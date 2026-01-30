const path = require("path");
const fs = require("fs").promises; // Use promises for async operations
const fsSync = require("fs"); // For synchronous operations like existsSync
const htmlToPptx = require("../../template-engine/htmlToPptx");
const sharedCache = require('../shared/cache'); // Import shared cache

const imageSavePath = path.resolve(__dirname, "../../../Uploads");
if (!fsSync.existsSync(imageSavePath)) {
  fsSync.mkdirSync(imageSavePath, { recursive: true });
}

/**
 * Convert HTML to PPTX, send response, and clean up all related files
 */
async function convertToPPTX(req, res) {
  try {
    const updatedHTML = req.body.html;

    if (!updatedHTML) {
      return res.status(400).json({ error: "No HTML content provided." });
    }

    if (!updatedHTML.includes("class=\"sli-slide\"")) {
      return res.status(400).json({ error: "Invalid HTML content: No slides found." });
    }

    // Use the original filename from shared cache
    let outputFileName = "converted_presentation.pptx";
    let nameWithoutExt = '';
    if (sharedCache.extractionFolderName) {
      nameWithoutExt = path.parse(sharedCache.extractionFolderName).name;
      outputFileName = `${nameWithoutExt}_converted.pptx`;
    }

    const outputFilePath = path.resolve(__dirname, `../../files/${outputFileName}`);

    // Call the conversion function and get the result
    const conversionResultString = await htmlToPptx.convertHTMLToPPTX(updatedHTML, outputFilePath, nameWithoutExt);

    // Parse the JSON response to get the actual file path
    const conversionResult = JSON.parse(conversionResultString);

    if (!conversionResult.success) {
      console.error('Conversion failed:', conversionResult.error);
      return res.status(500).json({
        error: "Conversion failed",
        details: conversionResult.error
      });
    }

    // Use the actual file path returned by the conversion function
    const finalOutputPath = conversionResult.fileName;

    // Check if file exists before trying to download
    if (!fsSync.existsSync(finalOutputPath)) {
      console.error(`File not found at: ${finalOutputPath}`);
      return res.status(404).json({
        error: "Converted file not found",
        expectedPath: finalOutputPath
      });
    }

    // If conversionResult includes the source folder path, use it; otherwise, derive it
    let convertedExtractedFolder = conversionResult.sourceFolder
      ? conversionResult.sourceFolder
      : path.join(path.dirname(finalOutputPath), nameWithoutExt);

    // Send the actual converted file for download
    res.download(finalOutputPath, path.basename(finalOutputPath), async (err) => {
      if (err) {
        console.error("Error sending PPTX file:", err);
        res.status(500).json({ error: "Error generating the PPTX file." });
      } else {
        // Cleanup: Delete original PPTX, converted PPTX, and all extracted 

        // try {
        //   // Delete original PPTX file
        //   if (sharedCache.originalPptxPath && fsSync.existsSync(sharedCache.originalPptxPath)) {
        //     // await fs.unlink(sharedCache.originalPptxPath);
        //     // console.log(`Deleted original PPTX file: ${sharedCache.originalPptxPath}`);
        //   } else {
        //     console.log(`Original PPTX file not found at: ${sharedCache.originalPptxPath}`);
        //   }

        //   // Delete Converted PPTX file

        //   if (outputFilePath) {
        //     // await fs.unlink(outputFilePath);
        //     console.log(`Deleted outputFilePath PPTX file: ${outputFilePath}`);

        //   }

        //   // Delete converted PPTX file
        //   if (fsSync.existsSync(finalOutputPath)) {
        //     // await fs.unlink(finalOutputPath);
        //     console.log(`Deleted converted PPTX file: ${finalOutputPath}`);
        //   } else {
        //     console.log(`Converted PPTX file not found at: ${finalOutputPath}`);
        //   }

        //   // Delete original extracted files folder
        //   if (sharedCache.extractedFilesPath && fsSync.existsSync(sharedCache.extractedFilesPath)) {
        //     // await fs.rm(sharedCache.extractedFilesPath, { recursive: true, force: true });
        //     console.log(`Deleted original extracted files folder: ${sharedCache.extractedFilesPath}`);
        //   } else {
        //     console.log(`Original extracted folder not found at: ${sharedCache.extractedFilesPath}`);
        //   }

        //   // Delete converted PPTX extracted folder (if it exists)
        //   if (convertedExtractedFolder && fsSync.existsSync(convertedExtractedFolder)) {
        //     // await fs.rm(convertedExtractedFolder, { recursive: true, force: true });
        //     console.log(`Deleted converted PPTX extracted folder: ${convertedExtractedFolder}`);
        //   } else {
        //     console.log(`Converted PPTX extracted folder not found at: ${convertedExtractedFolder}`);
        //   }

        //   // // Clear shared cache

        //   // sharedCache.originalPptxPath = null;
        //   // sharedCache.extractedFilesPath = null;
        //   // sharedCache.originalFileName = null;
        //   // sharedCache.fileNameWithoutExt = null;
        //   // sharedCache.extractionFolderName = null;
        //   // sharedCache.slidesHTMLCache = null;
        //   // console.log('Cleared shared cache after cleanup');
        // } catch (cleanupError) {
        //   console.error('Error during file cleanup:', cleanupError.message, cleanupError.stack);
        // }


        console.log("PPTX file sent successfully.");
      }
    });

  } catch (error) {
    console.error("Error converting HTML to PPTX:", error);
    res.status(500).json({ error: "Error during conversion." });
  }
}

module.exports = {
  convertToPPTX,
};