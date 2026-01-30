const { Worker } = require("worker_threads");
const path = require("path");
const EventEmitter = require("events");
const crypto = require("crypto");

class WorkerPool extends EventEmitter {
  constructor(options = {}) {
    super();

    this.poolSize = options.poolSize || 4;
    this.workerScript = options.workerScript || path.resolve(__dirname, "pptxToHtmlWorker.js");
    this.maxQueueSize = options.maxQueueSize || 100;
    this.workerTimeout = options.workerTimeout || 10 * 60 * 1000;
    this.idleTimeout = options.idleTimeout || 5 * 60 * 1000;

    if (!path.isAbsolute(this.workerScript)) {
      throw new Error('workerScript must be an absolute path');
    }

    this.workers = [];
    this.availableWorkers = [];
    this.busyWorkers = new Set();
    this.taskQueue = [];
    this.isShuttingDown = false;
    this.minWorkers = Math.max(1, Math.floor(this.poolSize / 2));
    this.activeTimers = new Set();

    // Track active tasks for better debugging
    this.activeTasks = new Map(); // taskId -> { workerId, userId, sessionId, startTime }

    this.stats = {
      totalTasks: 0,
      completedTasks: 0,
      failedTasks: 0,
      activeWorkers: 0,
      queuedTasks: 0,
      workerCrashes: 0,
      concurrentTasksMax: 0
    };

    this.initializePool();
  }

  async initializePool() {
    const workerPromises = [];
    for (let i = 0; i < this.poolSize; i++) {
      workerPromises.push(this.createWorker(i));
    }

    try {
      const workers = await Promise.allSettled(workerPromises);
      workers.forEach((result, index) => {
        if (result.status === 'fulfilled') {
          this.workers.push(result.value);
          this.availableWorkers.push(result.value);
        } else {
          console.error(`Failed to create worker ${index}:`, result.reason);
        }
      });

      this.stats.activeWorkers = this.workers.length;
      this.emit('ready');
      console.log(`Worker pool initialized with ${this.workers.length} workers`);
    } catch (error) {
      console.error('Failed to initialize worker pool:', error);
      this.emit('error', error);
    }
  }

  async createWorker(workerId = null) {
    return new Promise((resolve, reject) => {
      const id = workerId !== null ? workerId : Date.now();
      try {
        const worker = new Worker(this.workerScript, {
          workerData: { workerId: id }
        });

        worker.workerId = id;
        worker.isIdle = true;
        worker.createdAt = Date.now();
        worker.lastUsed = Date.now();
        worker.tasksCompleted = 0;
        worker.currentTask = null;
        worker.currentTaskId = null; // Track current task ID
        worker.isTerminated = false;
        worker.idleTimer = null;

        worker.on('error', (error) => {
          console.error(`Worker ${worker.workerId} error:`, error);
          this.handleWorkerError(worker, error);
        });

        worker.on('exit', (code) => {
          this.handleWorkerExit(worker, code);
        });

        this.startIdleMonitoring(worker);
        resolve(worker);
      } catch (error) {
        console.error(`Failed to create worker ${id}:`, error);
        reject(error);
      }
    });
  }

  startIdleMonitoring(worker) {
    if (worker.idleTimer) {
      clearTimeout(worker.idleTimer);
      this.activeTimers.delete(worker.idleTimer);
    }

    const timer = setTimeout(() => {
      this.activeTimers.delete(timer);

      if (worker.isIdle &&
        !worker.isTerminated &&
        Date.now() - worker.lastUsed > this.idleTimeout &&
        this.availableWorkers.length > this.minWorkers) {

        this.terminateWorker(worker);
      } else if (!worker.isTerminated) {
        this.startIdleMonitoring(worker);
      }
    }, this.idleTimeout);

    worker.idleTimer = timer;
    this.activeTimers.add(timer);
  }

  async execute(taskData, options = {}) {
    if (this.isShuttingDown) {
      throw new Error('Worker pool is shutting down');
    }
    const timeout = options.timeout || this.workerTimeout;

    // Generate unique task ID if not provided
    const taskId = taskData.taskId || `task_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
    taskData.taskId = taskId;

    return new Promise((resolve, reject) => {
      const task = {
        id: taskId,
        data: taskData,
        resolve,
        reject,
        timeout,
        createdAt: Date.now(),
        isResolved: false,
        userId: taskData.userId || null,
        sessionId: taskData.sessionId || null
      };
      
      this.stats.totalTasks++;
      this.stats.queuedTasks++;

      if (this.taskQueue.length >= this.maxQueueSize) {
        this.stats.failedTasks++;
        this.stats.queuedTasks--;
        reject(new Error('Task queue is full'));
        return;
      }

      this.taskQueue.push(task);
      
      // Log queue status for concurrent operations
      if (this.taskQueue.length > 1 || this.busyWorkers.size > 0) {
        console.log(`[Pool] Queued task ${taskId} - Queue: ${this.taskQueue.length}, Busy: ${this.busyWorkers.size}, Available: ${this.availableWorkers.length}`);
      }
      
      this.processNextTask();
    });
  }

  async processNextTask() {
    if (this.taskQueue.length === 0 || this.availableWorkers.length === 0) {
      return;
    }
    
    const task = this.taskQueue.shift();
    const worker = this.availableWorkers.shift();

    this.stats.queuedTasks--;
    this.busyWorkers.add(worker);
    worker.isIdle = false;
    worker.currentTask = task;
    worker.currentTaskId = task.id;
    worker.lastUsed = Date.now();

    // Track active task
    this.activeTasks.set(task.id, {
      workerId: worker.workerId,
      userId: task.userId,
      sessionId: task.sessionId,
      startTime: Date.now()
    });

    // Update concurrent tasks max
    if (this.busyWorkers.size > this.stats.concurrentTasksMax) {
      this.stats.concurrentTasksMax = this.busyWorkers.size;
    }

    if (worker.idleTimer) {
      clearTimeout(worker.idleTimer);
      this.activeTimers.delete(worker.idleTimer);
      worker.idleTimer = null;
    }

    try {
      await this.executeTask(worker, task);
    } catch (error) {
      console.error(`[Pool] Task execution error for ${task.id}:`, error);
      this.handleTaskError(worker, task, error);
    }
  }

  async executeTask(worker, task) {
    return new Promise((resolve) => {
      let timeoutId = null;
      let messageHandler = null;
      let errorHandler = null;

      const cleanup = () => {
        if (timeoutId) {
          clearTimeout(timeoutId);
          this.activeTimers.delete(timeoutId);
          timeoutId = null;
        }
        if (messageHandler) worker.removeListener('message', messageHandler);
        if (errorHandler) worker.removeListener('error', errorHandler);
        
        // Remove from active tasks
        this.activeTasks.delete(task.id);
      };

      const resolveTask = (success, result, error = null) => {
        if (task.isResolved) return;
        task.isResolved = true;

        cleanup();
        this.releaseWorker(worker);

        if (success) {
          this.stats.completedTasks++;
          task.resolve(result);
        } else {
          this.stats.failedTasks++;
          const errorMessage = error || 'Task failed';
          console.error(`[Pool] Task ${task.id} failed: ${errorMessage}`);
          task.reject(new Error(errorMessage));
        }
        resolve();
      };

      timeoutId = setTimeout(() => {
        console.warn(`[Pool] Task ${task.id} timed out on worker ${worker.workerId}`);
        resolveTask(false, null, `Task timed out after ${task.timeout}ms`);
        this.handleWorkerTimeout(worker);
      }, task.timeout);
      this.activeTimers.add(timeoutId);

      messageHandler = (result) => {
        // Log successful completion
        if (result.success && result.taskId) {
          const taskInfo = this.activeTasks.get(result.taskId);
          if (taskInfo) {
            const duration = Date.now() - taskInfo.startTime;
            console.log(`[Pool] Task ${result.taskId} completed in ${duration}ms by worker ${worker.workerId}`);
          }
        }
        resolveTask(result.success, result, result.error);
      };
      
      errorHandler = (error) => {
        resolveTask(false, null, `Worker error: ${error.message}`);
        this.handleWorkerError(worker, error);
      };

      worker.on('message', messageHandler);
      worker.on('error', errorHandler);

      try {
        worker.postMessage(task.data);
      } catch (error) {
        resolveTask(false, null, `Failed to send task: ${error.message}`);
      }
    });
  }

  releaseWorker(worker) {
    if (worker.isTerminated) return;

    this.busyWorkers.delete(worker);
    worker.isIdle = true;
    worker.currentTask = null;
    worker.currentTaskId = null;
    worker.tasksCompleted++;
    worker.lastUsed = Date.now();

    if (!this.isShuttingDown && !worker.isTerminated) {
      this.availableWorkers.push(worker);
      this.startIdleMonitoring(worker);
      setImmediate(() => this.processNextTask());
    }
  }

  handleWorkerTimeout(worker) {
    console.warn(`[Pool] Worker ${worker.workerId} timed out (task: ${worker.currentTaskId}), terminating...`);
    
    if (worker.currentTaskId) {
      this.activeTasks.delete(worker.currentTaskId);
    }
    
    this.terminateWorker(worker);
    if (!this.isShuttingDown) {
      this.replaceWorker(worker);
    }
  }

  handleWorkerError(worker, error) {
    console.error(`[Pool] Worker ${worker.workerId} error (task: ${worker.currentTaskId}):`, error.message);
    this.stats.workerCrashes++;
    
    if (worker.currentTaskId) {
      this.activeTasks.delete(worker.currentTaskId);
    }
    
    this.removeWorkerFromCollections(worker);

    if (worker.currentTask && !worker.currentTask.isResolved) {
      worker.currentTask.isResolved = true;
      this.stats.failedTasks++;
      worker.currentTask.reject(new Error(`Worker crashed: ${error.message}`));
    }

    if (!this.isShuttingDown) {
      this.replaceWorker(worker);
    }
  }

  handleWorkerExit(worker, code) {
    if (code !== 0) {
      console.error(`[Pool] Worker ${worker.workerId} exited with code ${code} (task: ${worker.currentTaskId})`);
      this.stats.workerCrashes++;
    }
    
    if (worker.currentTaskId) {
      this.activeTasks.delete(worker.currentTaskId);
    }
    
    this.removeWorkerFromCollections(worker);

    if (worker.currentTask && !worker.currentTask.isResolved) {
      worker.currentTask.isResolved = true;
      this.stats.failedTasks++;
      worker.currentTask.reject(new Error(`Worker exited with code ${code}`));
    }

    if (!this.isShuttingDown && code !== 0) {
      this.replaceWorker(worker);
    }
  }

  removeWorkerFromCollections(worker) {
    worker.isTerminated = true;

    if (worker.idleTimer) {
      clearTimeout(worker.idleTimer);
      this.activeTimers.delete(worker.idleTimer);
      worker.idleTimer = null;
    }

    const workerIndex = this.workers.indexOf(worker);
    if (workerIndex > -1) this.workers.splice(workerIndex, 1);

    const availableIndex = this.availableWorkers.indexOf(worker);
    if (availableIndex > -1) this.availableWorkers.splice(availableIndex, 1);

    this.busyWorkers.delete(worker);
    this.stats.activeWorkers = this.workers.length;
  }

  async replaceWorker(failedWorker) {
    const maxRetries = 3;
    let retries = 0;

    console.log(`[Pool] Replacing failed worker ${failedWorker.workerId}`);

    while (retries < maxRetries && !this.isShuttingDown) {
      try {
        const newWorker = await this.createWorker();
        this.workers.push(newWorker);
        this.availableWorkers.push(newWorker);
        this.stats.activeWorkers = this.workers.length;
        setImmediate(() => this.processNextTask());
        this.emit('worker-created', newWorker.workerId);
        console.log(`[Pool] Successfully created replacement worker ${newWorker.workerId}`);
        return;
      } catch (error) {
        retries++;
        console.error(`[Pool] Failed to create replacement worker (attempt ${retries}):`, error);
        if (retries < maxRetries) {
          await new Promise(resolve => setTimeout(resolve, 1000 * retries));
        }
      }
    }

    if (this.workers.length < this.minWorkers) {
      console.error(`[Pool] CRITICAL: Pool has ${this.workers.length} workers, minimum is ${this.minWorkers}`);
      this.emit('critical-low-workers', this.workers.length);
    }
  }

  async terminateWorker(worker) {
    if (worker.isTerminated) return;

    console.log(`[Pool] Terminating worker ${worker.workerId}`);
    this.removeWorkerFromCollections(worker);

    try {
      await worker.terminate();
      this.emit('worker-terminated', worker.workerId, 'manual');
    } catch (error) {
      console.error(`[Pool] Error terminating worker ${worker.workerId}:`, error);
    }
  }

  getStats() {
    return {
      ...this.stats,
      activeWorkers: this.workers.length,
      availableWorkers: this.availableWorkers.length,
      busyWorkers: this.busyWorkers.size,
      queuedTasks: this.taskQueue.length,
      activeTasks: this.activeTasks.size,
      poolSize: this.poolSize,
      minWorkers: this.minWorkers,
      idleTimeoutSeconds: this.idleTimeout / 1000,
      activeTimers: this.activeTimers.size
    };
  }

  getActiveTasks() {
    const tasks = [];
    for (const [taskId, info] of this.activeTasks.entries()) {
      tasks.push({
        taskId,
        workerId: info.workerId,
        userId: info.userId,
        sessionId: info.sessionId,
        durationMs: Date.now() - info.startTime
      });
    }
    return tasks;
  }

  healthCheck() {
    const stats = this.getStats();
    const activeTasks = this.getActiveTasks();
    const isHealthy = stats.activeWorkers >= this.minWorkers && !this.isShuttingDown;
    
    return { 
      healthy: isHealthy, 
      stats, 
      activeTasks,
      timestamp: new Date().toISOString() 
    };
  }

  async shutdown() {
    if (this.isShuttingDown) return;
    this.isShuttingDown = true;

    console.log(`[Pool] Shutting down with ${this.taskQueue.length} queued tasks and ${this.busyWorkers.size} busy workers`);

    for (const timer of this.activeTimers) clearTimeout(timer);
    this.activeTimers.clear();

    while (this.taskQueue.length > 0) {
      const task = this.taskQueue.shift();
      if (!task.isResolved) {
        task.isResolved = true;
        task.reject(new Error('Worker pool is shutting down'));
      }
    }

    const shutdownTimeout = 30000;
    const start = Date.now();
    while (this.busyWorkers.size > 0 && Date.now() - start < shutdownTimeout) {
      await new Promise(r => setTimeout(r, 100));
    }

    const terminationPromises = [...this.workers].map(w =>
      this.terminateWorker(w).catch(() => { })
    );
    await Promise.allSettled(terminationPromises);

    this.activeTasks.clear();
    console.log(`[Pool] Shutdown complete`);
    this.emit('shutdown');
  }
}

module.exports = WorkerPool;