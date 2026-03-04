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

    try {
        const { fileBuffer, filePath, flag, userId, sessionId } = taskData;

        // Input validation
        if (!fileBuffer) throw new Error('fileBuffer is required');
        if (!filePath) throw new Error('filePath is required');
        // if (flag === undefined || flag === null) throw new Error('flag parameter is required');
        if (!Buffer.isBuffer(fileBuffer) && !(fileBuffer instanceof Uint8Array)) {
            throw new Error('fileBuffer must be a Buffer or Uint8Array');
        }

        console.log(`[${taskId}] Starting PPTX to HTML conversion for user:${userId || 'unknown'}, session:${sessionId || 'unknown'}, file:${path.basename(filePath)}`);

        // Create parser with task isolation
        const parser = new pptxParser(fileBuffer);
        const unzippedFiles = await parser.unzip();
        if (!unzippedFiles) throw new Error('Failed to unzip PPTX file');

        // Use unique file name for this task to prevent collision
        const originalFileName = `${path.basename(filePath, '.pptx')}_${taskId}.pptx`;
        extractor = new pptxToHtml(unzippedFiles, originalFileName);
        overrideGetImageFromPicture(extractor);

        const slides = await extractor.convertAllSlidesToHTML(flag);
        if (!slides || !Array.isArray(slides)) {
            throw new Error('Conversion failed - slides is not a valid array');
        }

        const processingTime = Date.now() - startTime;

        console.log(`[${taskId}] Conversion successful in ${processingTime}ms, ${slides.length} slides`);

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
                isolatedFileName: originalFileName
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
            // Only cleanup extractor object, not files
            if (extractor) {
                if (typeof extractor.cleanup === 'function') extractor.cleanup();
                if (typeof extractor.destroy === 'function') extractor.destroy();
                extractor = null;
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
parentPort.on("close", () => {
    // Silent close
});

process.on('SIGTERM', () => {
    process.exit(0);
});