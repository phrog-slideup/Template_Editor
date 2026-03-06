const { parentPort, threadId, workerData } = require("worker_threads");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");

const workerId = workerData?.workerId || threadId;

// Generate unique task ID for tracking/logging
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

// === CRITICAL IMPORT WITH FATAL EXIT ===
let htmlToPptx;

try {
    const htmlToPptxPath = getModulePath('template-engine/htmlToPptx');
    htmlToPptx = require(htmlToPptxPath);
} catch (error) {
    console.error(`FATAL: Failed to import htmlToPptx:`, error.message);
    parentPort.postMessage({
        success: false,
        error: `Import failed: htmlToPptx - ${error.message}`,
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
    let tempFileCreated = false;
    let processingPaths = { tempFiles: [], directories: [] };

    try {
        const { html, outputFilePath, originalFolderName, userId, sessionId } = taskData;

        // Validation
        if (!html) throw new Error('html parameter is required');
        if (!outputFilePath) throw new Error('outputFilePath parameter is required');
        if (typeof html !== 'string') throw new Error('html must be a string');
        if (html.trim().length === 0) throw new Error('html cannot be empty');
        if (!html.includes('class="sli-slide"')) {
            throw new Error('Invalid HTML content - must contain slides with class="sli-slide"');
        }

        // CRITICAL: Keep originalFolderName unchanged - it's used to locate extracted PPTX files
        // For concurrent editing protection, ensure outputFilePath is unique per user/session
        const actualFolderName = originalFolderName || path.basename(outputFilePath, '.pptx');

        // console.log(`[${taskId}] Starting conversion - user:${userId || 'unknown'}, session:${sessionId || 'unknown'}, folder:${actualFolderName}`);

        // Ensure output directory
        const outputDir = path.dirname(outputFilePath);
        try {
            await fs.promises.access(outputDir, fs.constants.F_OK);
        } catch {
            await fs.promises.mkdir(outputDir, { recursive: true });
        }

        // Count slides
        const slideMatches = html.match(/class="sli-slide"/g);
        const slideCount = slideMatches ? slideMatches.length : 0;

        // Convert - Pass outputFilePath directly (original behavior)
        const conversionResultString = await htmlToPptx.convertHTMLToPPTX(
            html,
            outputFilePath,
            actualFolderName
        );
        tempFileCreated = true;

        let conversionResult;
        try {
            conversionResult = JSON.parse(conversionResultString);
        } catch (e) {
            throw new Error(`Failed to parse conversion result: ${e.message}`);
        }

        if (!conversionResult.success) {
            throw new Error(`Conversion failed: ${conversionResult.error}`);
        }

        const finalFilePath = conversionResult.fileName || outputFilePath;

        // Collect paths from result for reference
        if (conversionResult.debugInfo) {
            if (conversionResult.debugInfo.slideXmlsDir) processingPaths.directories.push(conversionResult.debugInfo.slideXmlsDir);
            if (conversionResult.debugInfo.extractedPptxDir) processingPaths.directories.push(conversionResult.debugInfo.extractedPptxDir);
        }
        if (conversionResult.conversionInfo) {
            if (conversionResult.conversionInfo.sourceFolder) processingPaths.directories.push(conversionResult.conversionInfo.sourceFolder);
            if (conversionResult.conversionInfo.originalZipPath) processingPaths.tempFiles.push(conversionResult.conversionInfo.originalZipPath);
            if (conversionResult.conversionInfo.tempSlideXmlsDir) processingPaths.directories.push(conversionResult.conversionInfo.tempSlideXmlsDir);
        }

        // Verify file
        let fileStats;
        try {
            fileStats = await fs.promises.stat(finalFilePath);
        } catch (e) {
            throw new Error(`Final PPTX file not created at ${finalFilePath}: ${e.message}`);
        }
        if (fileStats.size === 0) throw new Error('Generated PPTX file is empty');

        const processingTime = Date.now() - startTime;
        // console.log(`[${taskId}] Conversion successful in ${processingTime}ms`);

        parentPort.postMessage({
            success: true,
            outputFilePath: finalFilePath,
            originalOutputFilePath: outputFilePath,
            fileSize: fileStats.size,
            slideCount,
            processingTime,
            result: conversionResult,
            workerId,
            threadId,
            taskId,
            userId: userId || null,
            sessionId: sessionId || null,
            timestamp: new Date().toISOString(),
            processingDetails: {
                uniqueTaskId: taskId,
                directoryProcessingFixed: true,
                slideXmlsExtracted: !!conversionResult.debugInfo?.slideXmlsDir,
                pptxExtracted: !!conversionResult.debugInfo?.extractedPptxDir,
                structureModified: !!conversionResult.conversionInfo?.sourceFolder,
                zipConverted: !!conversionResult.conversionInfo?.originalZipPath,
                finalFileSize: fileStats.size,
                compressionRatio: conversionResult.conversionInfo?.sizeKB || 'unknown',
                originalFolderSimulated: actualFolderName
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
            timestamp: new Date().toISOString(),
            processingDetails: {
                uniqueTaskId: taskId,
                directoryProcessingAttempted: true,
                errorDuringProcessing: true
            }
        });
    } finally {
        try {
            if (global.gc) global.gc();
        } catch (e) {
            console.warn(`[${taskId}] Worker ${workerId} cleanup error:`, e.message);
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

process.on('unhandledRejection', (reason) => {
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