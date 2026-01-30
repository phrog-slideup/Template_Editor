require("dotenv").config();
const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
const cors = require("cors");

// Import WorkerPool
const WorkerPool = require("./app/workers/workerPool");

const app = express();

// Import the scheduler (it will auto-run daily at 11PM)

require("./app/scheduler/FileCleanupScheduler.js");

// ASYNC FUNCTION TO INITIALIZE SERVER WITH DUAL WORKER POOLS
async function initializeServer() {
  try {
    console.log('🔧 Initializing dual worker pools...');

    // Initialize PPTX to HTML Worker Pool   Creating PPTX-to-HTML worker pool.

    const pptxWorkerPool = new WorkerPool({
      poolSize: 4,
      workerScript: path.resolve(__dirname, "./app/workers/pptxToHtmlWorker.js"),
      maxQueueSize: 100,
      workerTimeout: 10 * 60 * 1000, // 10 minutes
      idleTimeout: 5 * 60 * 1000     // 5 minutes
    });

    // Initialize HTML to PPTX Worker Pool  Creating HTML-to-PPTX worker pool...

    const htmlWorkerPool = new WorkerPool({
      poolSize: 2, // Smaller pool since PPTX conversion is typically slower
      workerScript: path.resolve(__dirname, "./app/workers/htmlToPptxWorker.js"),
      maxQueueSize: 50,
      workerTimeout: 15 * 60 * 1000, // 15 minutes (PPTX generation can be slower)
      idleTimeout: 5 * 60 * 1000     // 5 minutes
    });

    // console.log("Worker Pools Configuration:", {
    //   pptxPool: {
    //     poolSize: pptxWorkerPool.poolSize,
    //     workerScript: pptxWorkerPool.workerScript,
    //     isShuttingDown: pptxWorkerPool.isShuttingDown
    //   },
    //   htmlPool: {
    //     poolSize: htmlWorkerPool.poolSize,
    //     workerScript: htmlWorkerPool.workerScript,
    //     isShuttingDown: htmlWorkerPool.isShuttingDown
    //   }
    // });

    // Add event listeners for PPTX worker pool   PPTX-to-HTML worker pool is ready!
    pptxWorkerPool.on('ready', () => {
      console.log('✅ PPTX-to-HTML worker pool is ready!');
      const stats = pptxWorkerPool.getStats();
      console.log('📊 PPTX worker pool stats:', JSON.stringify(stats, null, 2));
    });

    pptxWorkerPool.on('shutdown', () => {
      console.log('🛑 PPTX worker pool has shut down');
    });

    pptxWorkerPool.on('critical-low-workers', (count) => {
      console.error(`🚨 CRITICAL: PPTX pool has only ${count} workers remaining!`);
    });

    // Add event listeners for HTML worker pool
    htmlWorkerPool.on('ready', () => {
      console.log('✅ HTML-to-PPTX worker pool is ready!');
      const stats = htmlWorkerPool.getStats();
      console.log('📊 HTML worker pool stats:', JSON.stringify(stats, null, 2));
    });

    htmlWorkerPool.on('shutdown', () => {
      console.log('🛑 HTML worker pool has shut down');
    });

    htmlWorkerPool.on('critical-low-workers', (count) => {
      console.error(`🚨 CRITICAL: HTML pool has only ${count} workers remaining!`);
    });

    // WAIT FOR BOTH WORKER POOLS TO BE READY
    console.log('⏳ Waiting for both worker pools to initialize...');
    await Promise.all([
      new Promise((resolve, reject) => {
        let resolved = false;

        pptxWorkerPool.on('ready', () => {
          if (!resolved) {
            resolved = true;
            console.log('🎉 PPTX worker pool initialization complete!');
            resolve();
          }
        });

        setTimeout(() => {
          if (!resolved) {
            resolved = true;
            console.error('❌ PPTX worker pool initialization timeout (60s)');
            reject(new Error('PPTX worker pool initialization timeout'));
          }
        }, 60000);
      }),

      new Promise((resolve, reject) => {
        let resolved = false;

        htmlWorkerPool.on('ready', () => {
          if (!resolved) {
            resolved = true;
            console.log('🎉 HTML worker pool initialization complete!');
            resolve();
          }
        });

        setTimeout(() => {
          if (!resolved) {
            resolved = true;
            console.error('❌ HTML worker pool initialization timeout (60s)');
            reject(new Error('HTML worker pool initialization timeout'));
          }
        }, 60000);
      })
    ]);

    // Make worker pools available to controllers
    app.locals.pptxWorkerPool = pptxWorkerPool;  // For PPTX to HTML conversion
    app.locals.htmlWorkerPool = htmlWorkerPool;  // For HTML to PPTX conversion
    app.locals.workerPool = pptxWorkerPool;      // Keep backward compatibility

    console.log('🎊 Both worker pools initialized successfully!');

    // Setup graceful shutdown handling for both pools
    const gracefulShutdown = async (signal) => {
      console.log(`\n📨 Received ${signal}. Starting graceful shutdown...`);
      try {
        console.log('🛑 Shutting down both worker pools...');

        // Shutdown both pools with timeout
        const shutdownPromises = [
          pptxWorkerPool.shutdown(),
          htmlWorkerPool.shutdown()
        ].map(promise =>
          Promise.race([
            promise,
            new Promise((_, reject) =>
              setTimeout(() => reject(new Error('Worker pool shutdown timeout')), 15000)
            )
          ])
        );

        await Promise.allSettled(shutdownPromises);
        console.log('✅ All worker pools shutdown completed');
        process.exit(0);
      } catch (error) {
        console.error('❌ Error during shutdown (forcing exit):', error.message);
        process.exit(1);
      }
    };

    process.on('SIGINT', () => gracefulShutdown('SIGINT'));
    process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));

    // Fix CORS Configuration (YOUR EXISTING CODE - UNCHANGED)
    const corsOptions = {
      origin: ["https://deckster.slideuplift.com"], // Allow frontend URL
      methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
      allowedHeaders: ["Origin", "X-Requested-With", "Content-Type", "Accept", "Authorization"],
      credentials: true,
    };

    app.use(cors(corsOptions));
    app.options("*", cors(corsOptions)); // Handle preflight requests

    // Middleware for parsing JSON and URL-encoded data (YOUR EXISTING CODE - UNCHANGED)
    app.use(bodyParser.urlencoded({ limit: "100mb", extended: true }));
    app.use(bodyParser.json({ limit: "100mb", extended: true }));

    // Ensure CORS Headers on All Responses (YOUR EXISTING CODE - UNCHANGED)
    app.use((req, res, next) => {
      res.header("Access-Control-Allow-Origin", "https://deckster.slideuplift.com");
      res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, PUT, DELETE");
      res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept, Authorization");
      res.header("Access-Control-Allow-Credentials", "true");

      // Ensure OPTIONS request gets correct response
      if (req.method === "OPTIONS") {
        return res.status(204).end();
      }

      next();
    });

    // Serve static files (YOUR EXISTING CODE - UNCHANGED)
    app.use(express.static(path.join(__dirname, "app/views")));
    app.use("/uploads", express.static(path.join(__dirname, "app/uploads")));

    console.log(`Serving static files from: ${path.join(__dirname, "uploads")}`);

    // Serve the frontend page (YOUR EXISTING CODE - UNCHANGED)
    app.get("/", (req, res) => {
      res.sendFile(path.join(__dirname, "app/views/ppt-edit.html"));
    });

    // Add comprehensive health check endpoint
    app.get("/api/health", (req, res) => {
      try {
        const pptxHealthy = pptxWorkerPool && !pptxWorkerPool.isShuttingDown;
        const htmlHealthy = htmlWorkerPool && !htmlWorkerPool.isShuttingDown;

        if (!pptxHealthy || !htmlHealthy) {
          return res.status(503).json({
            healthy: false,
            services: {
              pptxToHtml: pptxHealthy ? 'healthy' : 'unavailable',
              htmlToPptx: htmlHealthy ? 'healthy' : 'unavailable'
            },
            timestamp: new Date().toISOString()
          });
        }

        const pptxHealth = pptxWorkerPool.healthCheck();
        const htmlHealth = htmlWorkerPool.healthCheck();

        res.json({
          healthy: pptxHealth.healthy && htmlHealth.healthy,
          services: {
            pptxToHtml: pptxHealth,
            htmlToPptx: htmlHealth
          },
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        res.status(500).json({
          healthy: false,
          error: 'Health check failed',
          details: error.message,
          timestamp: new Date().toISOString()
        });
      }
    });

    // Add comprehensive worker pool stats endpoint
    app.get("/api/worker-stats", (req, res) => {
      try {
        const pptxAvailable = pptxWorkerPool && !pptxWorkerPool.isShuttingDown;
        const htmlAvailable = htmlWorkerPool && !htmlWorkerPool.isShuttingDown;

        const stats = {
          timestamp: new Date().toISOString(),
          uptime: process.uptime(),
          services: {
            pptxToHtml: pptxAvailable ? pptxWorkerPool.getStats() : 'unavailable',
            htmlToPptx: htmlAvailable ? htmlWorkerPool.getStats() : 'unavailable'
          },
          summary: {
            totalWorkers: (pptxAvailable ? pptxWorkerPool.getStats().activeWorkers : 0) +
              (htmlAvailable ? htmlWorkerPool.getStats().activeWorkers : 0),
            totalTasks: (pptxAvailable ? pptxWorkerPool.getStats().totalTasks : 0) +
              (htmlAvailable ? htmlWorkerPool.getStats().totalTasks : 0),
            totalCompleted: (pptxAvailable ? pptxWorkerPool.getStats().completedTasks : 0) +
              (htmlAvailable ? htmlWorkerPool.getStats().completedTasks : 0),
            totalFailed: (pptxAvailable ? pptxWorkerPool.getStats().failedTasks : 0) +
              (htmlAvailable ? htmlWorkerPool.getStats().failedTasks : 0)
          }
        };

        res.json(stats);
      } catch (error) {
        res.status(500).json({
          error: 'Failed to get worker stats',
          details: error.message,
          timestamp: new Date().toISOString()
        });
      }
    });

    // Mount API routes (YOUR EXISTING CODE - UNCHANGED)
    const routes = require("./app/api/routes/slide.routes.js");
    app.use("/api/slides", routes);

    // Start the server (YOUR EXISTING CODE - UNCHANGED)
    const PORT = 3006;
    const server = app.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
      console.log(`📊 PPTX Worker pool: ${pptxWorkerPool.getStats().activeWorkers} workers`);
      console.log(`🎯 HTML Worker pool: ${htmlWorkerPool.getStats().activeWorkers} workers`);
      console.log(`🔗 Health check: http://localhost:${PORT}/api/health`);
      console.log(`📈 Worker stats: http://localhost:${PORT}/api/worker-stats`);
      console.log(`🌐 Frontend: http://localhost:${PORT}`);
      console.log(`📋 API Routes: http://localhost:${PORT}/api/slides/`);
    });

    // Handle server shutdown
    server.on('close', async () => {
      console.log('🏃 Server closing. Shutting down worker pools...');
      try {
        const shutdownPromises = [];

        if (pptxWorkerPool && !pptxWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              pptxWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('PPTX pool shutdown timeout')), 10000)
              )
            ])
          );
        }

        if (htmlWorkerPool && !htmlWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              htmlWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('HTML pool shutdown timeout')), 10000)
              )
            ])
          );
        }

        await Promise.allSettled(shutdownPromises);
        console.log('✅ Server shutdown complete');
      } catch (error) {
        console.error('❌ Error during server shutdown:', error.message);
      }
    });

    // Handle uncaught exceptions with dual pool shutdown
    process.on('uncaughtException', async (error) => {
      console.error('💥 Uncaught Exception:', error);
      try {
        const shutdownPromises = [];

        if (pptxWorkerPool && !pptxWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              pptxWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('Emergency PPTX shutdown timeout')), 3000)
              )
            ])
          );
        }

        if (htmlWorkerPool && !htmlWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              htmlWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('Emergency HTML shutdown timeout')), 3000)
              )
            ])
          );
        }

        await Promise.allSettled(shutdownPromises);
      } catch (shutdownError) {
        console.error('❌ Error during emergency shutdown:', shutdownError.message);
      }
      process.exit(1);
    });

    // Handle unhandled promise rejections with dual pool shutdown
    process.on('unhandledRejection', async (reason, promise) => {
      console.error('💥 Unhandled Rejection at:', promise, 'reason:', reason);
      try {
        const shutdownPromises = [];

        if (pptxWorkerPool && !pptxWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              pptxWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('Emergency PPTX shutdown timeout')), 3000)
              )
            ])
          );
        }

        if (htmlWorkerPool && !htmlWorkerPool.isShuttingDown) {
          shutdownPromises.push(
            Promise.race([
              htmlWorkerPool.shutdown(),
              new Promise((_, reject) =>
                setTimeout(() => reject(new Error('Emergency HTML shutdown timeout')), 3000)
              )
            ])
          );
        }

        await Promise.allSettled(shutdownPromises);
      } catch (shutdownError) {
        console.error('❌ Error during emergency shutdown:', shutdownError.message);
      }
      process.exit(1);
    });

    return { app, server, pptxWorkerPool, htmlWorkerPool };

  } catch (error) {
    console.error('❌ Failed to initialize server:', error);
    console.error('Stack trace:', error.stack);
    process.exit(1);
  }
}

// Start the server
console.log('🔄 Starting dual-pool server initialization...');
initializeServer().then(({ server, pptxWorkerPool, htmlWorkerPool }) => {
  console.log('✅ Dual-pool server initialization successful!');
  console.log(`🎊 Total active workers: ${pptxWorkerPool.getStats().activeWorkers + htmlWorkerPool.getStats().activeWorkers}`);
}).catch((error) => {
  console.error('❌ Server initialization failed:', error);
  process.exit(1);
});

module.exports = app;