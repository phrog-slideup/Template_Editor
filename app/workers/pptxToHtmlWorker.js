const { parentPort, threadId, workerData } = require("worker_threads");
const path = require("path");
const crypto = require("crypto");

const workerId = workerData?.workerId || threadId;

// Generate unique task ID
function generateUniqueTaskId() {
    return `${Date.now()}_${crypto.randomBytes(8).toString('hex')}_${process.pid}_${threadId}`;
}

// Robust path resolution
const getModulePath = (relativePath) => {
    const basePaths = [
        __dirname,
        path.join(__dirname, '..'),
        path.join(__dirname, '../..'),
        process.cwd(),
        path.join(process.cwd(), 'app')
    ];

    for (const basePath of basePaths) {
        try {
            const fullPath = path.resolve(basePath, relativePath);
            require.resolve(fullPath);
            return fullPath;
        } catch (e) {
            // continue
        }
    }
    return relativePath;
};

// === CRITICAL IMPORTS WITH FATAL EXIT ON FAILURE ===
let pptxToHtml, pptxParser, overrideGetImageFromPicture;

try {
    const pptxToHtmlPath = getModulePath('template-engine/pptxToHtml');
    pptxToHtml = require(pptxToHtmlPath);
} catch (error) {
    console.error(`FATAL: Failed to import pptxToHtml:`, error.message);
    parentPort.postMessage({
        success: false,
        error: `Import failed: pptxToHtml - ${error.message}`,
        fatal: true,
        workerId,
        threadId
    });
    process.exit(1);
}

try {
    const pptxParserPath = getModulePath('api/helper/pptParser.helper');
    pptxParser = require(pptxParserPath);
} catch (error) {
    console.error(`FATAL: Failed to import pptxParser:`, error.message);
    parentPort.postMessage({
        success: false,
        error: `Import failed: pptxParser - ${error.message}`,
        fatal: true,
        workerId,
        threadId
    });
    process.exit(1);
}

try {
    const imageHandlingPath = getModulePath('api/helper/pptImagesHandling');
    const imageHandling = require(imageHandlingPath);
    overrideGetImageFromPicture = imageHandling.overrideGetImageFromPicture;
} catch (error) {
    console.error(`FATAL: Failed to import pptImagesHandling:`, error.message);
    parentPort.postMessage({
        success: false,
        error: `Import failed: pptImagesHandling - ${error.message}`,
        fatal: true,
        workerId,
        threadId
    });
    process.exit(1);
}

// === MESSAGE HANDLER ===
parentPort.on("message", async (taskData) => {
    const taskId = generateUniqueTaskId();
    const startTime = Date.now();
    let extractor = null;
    let parser = null;

    try {
        const { fileBuffer, filePath, flag, userId, sessionId } = taskData;

        // Input validation
        if (!fileBuffer) throw new Error('fileBuffer is required');
        if (!filePath) throw new Error('filePath is required');
        if (!Buffer.isBuffer(fileBuffer) && !(fileBuffer instanceof Uint8Array)) {
            throw new Error('fileBuffer must be a Buffer or Uint8Array');
        }
        if (fileBuffer.length === 0) {
            throw new Error('fileBuffer is empty');
        }

        // ============================================================================
        // CRITICAL FIX FOR RACE CONDITION:
        // Use unique file name per task to prevent collisions when multiple workers
        // process files with the same original name simultaneously
        // ============================================================================
        const originalFileName = `${path.basename(filePath, '.pptx')}_${taskId}.pptx`;
        
        // console.log(`[${taskId}] Original file: ${filePath}`);
        // console.log(`[${taskId}] Task-isolated name: ${originalFileName}`);

        // Create parser with task isolation
        parser = new pptxParser(fileBuffer);
        const unzippedFiles = await parser.unzip();
        
        if (!unzippedFiles) {
            throw new Error('Failed to unzip PPTX file - invalid or corrupted file');
        }

        console.log(`[${taskId}] File extracted successfully`);

        // Create extractor with isolated filename to prevent race conditions
        extractor = new pptxToHtml(unzippedFiles, originalFileName);
        overrideGetImageFromPicture(extractor);

        const slides = await extractor.convertAllSlidesToHTML(flag);
        
        // Validate conversion result
        if (!slides) {
            throw new Error('Conversion returned null or undefined');
        }
        if (!Array.isArray(slides)) {
            throw new Error(`Conversion failed - expected array but got ${typeof slides}`);
        }
        if (slides.length === 0) {
            throw new Error('Conversion resulted in zero slides - file may be empty or invalid');
        }

        const processingTime = Date.now() - startTime;

        parentPort.postMessage({
            success: true,
            slides,
            slideCount: slides.length,
            processingTime,
            workerId,
            threadId,
            taskId,
            userId: userId || null,
            sessionId: sessionId || null,
            timestamp: new Date().toISOString(),
            processingDetails: {
                uniqueTaskId: taskId,
                originalFileName: filePath,
                isolatedFileName: originalFileName,
                bufferSize: fileBuffer.length,
                raceConditionPrevented: true,
                concurrentOperationSafe: true
            }
        });

    } catch (error) {
        const processingTime = Date.now() - startTime;
        console.error(`[${taskId}] Worker ${workerId} task error after ${processingTime}ms:`, error.message);
        console.error(`[${taskId}] Error stack:`, error.stack);

        parentPort.postMessage({
            success: false,
            error: error.message,
            errorStack: error.stack,
            processingTime,
            workerId,
            threadId,
            taskId,
            timestamp: new Date().toISOString()
        });
    } finally {
        try {
            // Cleanup extractor and parser objects
            if (extractor) {
                if (typeof extractor.cleanup === 'function') extractor.cleanup();
                if (typeof extractor.destroy === 'function') extractor.destroy();
                extractor = null;
            }
            if (parser) {
                if (typeof parser.cleanup === 'function') parser.cleanup();
                if (typeof parser.destroy === 'function') parser.destroy();
                parser = null;
            }
            if (global.gc) global.gc();
        } catch (cleanupError) {
            console.error(`[${taskId}] Worker ${workerId} cleanup error:`, cleanupError.message);
        }
    }
});

// === GLOBAL ERROR HANDLING ===
process.on('uncaughtException', (error) => {
    console.error(`Worker ${workerId} uncaught exception:`, error.message);
    console.error(`Stack:`, error.stack);

    try {
        parentPort.postMessage({
            success: false,
            error: `Uncaught exception: ${error.message}`,
            errorStack: error.stack,
            fatal: true,
            workerId,
            threadId,
            timestamp: new Date().toISOString()
        });
    } catch (e) {
        console.error(`Failed to send error message:`, e.message);
    }

    setTimeout(() => process.exit(1), 100);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error(`Worker ${workerId} unhandled rejection:`, reason);

    try {
        parentPort.postMessage({
            success: false,
            error: `Unhandled rejection: ${reason}`,
            fatal: true,
            workerId,
            threadId,
            timestamp: new Date().toISOString()
        });
    } catch (e) {
        console.error(`Failed to send error message:`, e.message);
    }

    setTimeout(() => process.exit(1), 100);
});

// Graceful shutdown
parentPort.on("close", () => { });
process.on('SIGTERM', () => process.exit(0));