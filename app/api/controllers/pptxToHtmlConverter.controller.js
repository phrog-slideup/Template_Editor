const path = require('path');
const fs = require('fs').promises; // Use promises for async/await
const fsSync = require('fs'); // Import synchronous fs for existsSync and mkdirSync
const pptxToHtml = require('../../template-engine/pptxToHtml');
const JSZip = require('jszip');
const { overrideGetImageFromPicture } = require('../helper/pptImagesHandling');
const axios = require('axios');
const sharedCache = require('../shared/cache'); // Import shared cache
const { v4: uuidv4 } = require('uuid'); // For generating unique IDs


// Define the files directory
const filesDir = path.resolve(__dirname, '../../files');

async function uploadFile(req, res) {
  const startTime = Date.now();

  try {
    const { filePath, flag, bucketName, key, generatePresignedUrl = false } = req.body;
    console.log("filePath ================ controller >>>>>>>>>> ", filePath);

    let fileBuffer;
    let originalFileName;
    let s3Details = null;
    let localFilePath = null;

    // Generate unique filename for the PPTX file
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const uniqueFileName = `${uuidv4()}_${timestamp}.pptx`;

    // Ensure the files directory exists
    try {
      if (!fsSync.existsSync(filesDir)) {
        fsSync.mkdirSync(filesDir, { recursive: true });
      }
    } catch (err) {
      console.error(`Error creating directory ${filesDir}:`, err.message);
      throw new Error(`Failed to create files directory: ${err.message}`);
    }

    // Download file based on input
    if (bucketName && key) {
      // console.log(`Downloading from S3: ${bucketName}/${key}`);
      fileBuffer = await downloadPPTXFromS3(bucketName, key);
      originalFileName = path.basename(key);
      s3Details = { bucketName, key, fileName: originalFileName, uniqueFileName };
    } else if (filePath && (filePath.includes("s3.amazonaws.com") || filePath.startsWith("s3://"))) {
      // console.log(`Downloading from S3 URL: ${filePath}`);
      fileBuffer = await downloadPPTXFromS3Url(filePath);
      const parsedS3 = parseS3Url(filePath);
      originalFileName = path.basename(parsedS3.key);
      s3Details = {
        bucketName: parsedS3.bucketName,
        key: parsedS3.key,
        fileName: originalFileName,
        uniqueFileName,
        originalUrl: filePath,
      };
    } else if (filePath && filePath.includes("kamaiipro.in")) {

      fileBuffer = await downloadFileBuffer(filePath);
      originalFileName = path.basename(filePath);
      s3Details = { fileName: originalFileName, uniqueFileName };
    } else if (filePath && filePath.startsWith("http")) {

      console.log(`Downloading from HTTP URL: ${filePath}`);

      fileBuffer = await downloadFileBuffer(filePath);
      originalFileName = path.basename(filePath);
      s3Details = { fileName: originalFileName, uniqueFileName };
    } else if (filePath) {
      s3Details = extractS3Details(filePath);
      fileBuffer = await downloadPPTXFromS3(s3Details.bucketName, s3Details.key);
      originalFileName = s3Details.fileName;
      s3Details.uniqueFileName = uniqueFileName;
    } else {
      return res.status(400).json({
        error:
          "Invalid file source. Provide one of: " +
          "1) bucketName and key parameters, " +
          "2) S3 URL (s3.amazonaws.com), " +
          "3) kamaiipro.in website URL, " +
          "4) S3 path format (bucket/folder/file.pptx), " +
          "5) Other HTTP URLs.",
      });
    }

    // Validate file buffer
    if (!fileBuffer || fileBuffer.length === 0) {
      return res.status(400).json({ error: "Failed to download file or file is empty." });
    }

    // Save the file to the files directory
    try {
      localFilePath = path.join(filesDir, uniqueFileName);
      await fs.writeFile(localFilePath, fileBuffer);

      // console.log(`File saved successfully to: ${localFilePath}`);
    } catch (err) {
      console.error(`Error saving file to ${localFilePath}:`, err.message);
      throw new Error(`Failed to save file to disk: ${err.message}`);
    }

    // ========== PPTX EXTRACTION LOGIC ==========
    // Generate extraction folder name using the unique file name (same as first controller)
    const fileNameWithoutExt = path.parse(uniqueFileName).name; // Get unique filename without extension
    const extractionFolderName = `${fileNameWithoutExt}_extracted`;
    const extractionDir = path.join(filesDir, extractionFolderName);

    // Extract PPTX files maintaining folder structure
    try {
      await saveExtractedFiles(fileBuffer, extractionDir);

    } catch (extractionError) {
      console.error("❌ PPTX extraction failed:", extractionError);
      // Continue processing even if extraction fails, but log the error
    }

    // ========== STORE IN SHARED CACHE (Same as first controller) ==========
    sharedCache.extractedFilesPath = extractionDir;
    sharedCache.originalPptxPath = localFilePath;
    sharedCache.originalFileName = uniqueFileName; // Store the unique filename
    sharedCache.fileNameWithoutExt = path.parse(uniqueFileName).name;
    sharedCache.extractionFolderName = extractionFolderName;

    // Store S3 file details separately (not in cache as per requirement)
    currentFileDetails = {
      ...s3Details,
      localFilePath,
      extractionDir,
      extractionFolderName,
      extractedFiles: true,
    };

    // Get worker pool
    const workerPool = req.app.locals.workerPool;
    if (!workerPool) {
      throw new Error("Worker pool not available");
    }

    // Process file with worker pool
    try {
      // console.log("Submitting task to worker pool...");
      const result = await workerPool.execute({ fileBuffer, filePath: uniqueFileName, flag }, { timeout: 5 * 60 * 1000 });

      if (result.success) {
        // Store slides in shared cache (same as first controller)
        sharedCache.slidesHTMLCache = result.slides;

        // console.log("Stored in cache - slidesHTMLCache length:", sharedCache.slidesHTMLCache?.length);

        const processingTime = Date.now() - startTime;
        // console.log(`Processing completed in ${processingTime}ms`);

        const response = {
          message: "File processed successfully.",
          slides: result.slides,
          fileName: uniqueFileName,
          originalFileName: uniqueFileName, // Return unique filename same as first controller
          fileSize: fileBuffer.length,
          processingTime,
          localFilePath,
          // Add extraction details to response (same as first controller)
          originalFile: localFilePath,
          extractedPath: extractionDir,
          extractionFolderName: extractionFolderName,
        };

        // Generate presigned URL if requested
        if (generatePresignedUrl && s3Details && s3Details.bucketName && s3Details.key) {
          try {
            const presignedUrl = await createPresignedPPTXDownloadUrl(s3Details.bucketName, s3Details.key, {
              expiresIn: 3600,
              downloadFilename: uniqueFileName,
              forceDownload: true,
            });
            response.presignedDownloadUrl = presignedUrl;
            response.presignedUrlExpires = new Date(Date.now() + 3600 * 1000).toISOString();
            // console.log(`✅ Presigned URL created for: ${uniqueFileName}`);
          } catch (presignedError) {
            console.error("Failed to create presigned URL:", presignedError);
            response.presignedUrlError = "Failed to create presigned download URL";
          }
        }

        if (s3Details) {
          response.s3Details = s3Details;
        }

        res.status(200).json(response);
      } else {
        console.error("Worker pool task failed:", result.error);
        return res.status(500).json({
          error: result.error || "Processing failed",
          processingTime: Date.now() - startTime,
        });
      }
    } catch (workerError) {
      console.error("Worker pool execution error:", workerError);
      return res.status(500).json({
        error: "Worker pool execution failed: " + workerError.message,
        processingTime: Date.now() - startTime,
      });
    }
  } catch (error) {
    console.error("Error in uploadFile:", error);
    if (error.message.includes("S3")) {
      res.status(500).json({
        error: `S3 Error: ${error.message}`,
        processingTime: Date.now() - startTime,
      });
    } else if (error.message.includes("Network")) {
      res.status(500).json({
        error: `Network Error: ${error.message}`,
        processingTime: Date.now() - startTime,
      });
    } else {
      res.status(500).json({
        error: `Error processing file: ${error.message}`,
        processingTime: Date.now() - startTime,
      });
    }
  }
}

// HELPER FUNCTION - Direct file download for HTTP URLs
async function downloadFileBuffer(filePath) {
  // console.log(`Downloading file: ${filePath}`);
  try {
    const response = await axios.get(filePath, {
      responseType: "arraybuffer",
      timeout: 30000,
      maxContentLength: 50 * 1024 * 1024
    });
    const fileBuffer = Buffer.from(response.data);
    // console.log(`File downloaded successfully: ${filePath}`);
    return fileBuffer;
  } catch (error) {
    console.error(`Error downloading file ${filePath}:`, error.message);
    throw new Error(`Failed to download file: ${error.message}`);
  }
}


// Function to save extracted files maintaining folder structure
async function saveExtractedFiles(pptxFileBuffer, extractionDir) {
  try {
    // Load the PPTX file as a ZIP archive
    const zip = await JSZip.loadAsync(pptxFileBuffer);

    // Ensure the extraction directory exists
    await fs.mkdir(extractionDir, { recursive: true });

    // Process each file in the ZIP archive
    for (const [filePath, zipEntry] of Object.entries(zip.files)) {
      const fullPath = path.join(extractionDir, filePath);
      const dirPath = path.dirname(fullPath);

      // Create directory if it doesn't exist
      await fs.mkdir(dirPath, { recursive: true });

      // Check if the file is an XML file based on extension
      const isXmlFile = filePath.endsWith('.xml') || filePath.endsWith('.rels');

      if (!zipEntry.dir) { // Skip directories
        if (isXmlFile) {
          // Extract XML content as a string
          const xmlContent = await zipEntry.async('string');
          await fs.writeFile(fullPath, xmlContent, 'utf8');
          // console.log(`Saved XML file: ${filePath}`);
        } else {
          // Extract binary content (e.g., images) as a buffer
          const binaryContent = await zipEntry.async('nodebuffer');
          await fs.writeFile(fullPath, binaryContent);
          // console.log(`Saved binary file: ${filePath}`);
        }
      }
    }

    // console.log(`All files extracted to: ${extractionDir}`);
  } catch (error) {
    console.error('Error saving extracted files:', error);
    throw error;
  }
}

async function downloadHTML(req, res) {
  console.log('Downloading HTML file...');

  try {
    // Use shared cache instead of local cache
    if (!sharedCache.slidesHTMLCache) {
      console.error('No slides available for download.');
      return res.status(400).json({ error: 'No slides available for download. Please upload and process a PPTX file first.' });
    }

    // Combine cached slides into a single HTML file
    const fullHTML = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Converted Slides</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
        </style>
      </head>
      <body>
        ${sharedCache.slidesHTMLCache.join('\n')}
      </body>
      </html>
    `;

    // Define the output path for the generated HTML file
    const htmlFilePath = path.resolve(__dirname, '../../public/converted-slides.html');

    // Write the HTML content to a file
    await fs.writeFile(htmlFilePath, fullHTML, 'utf8');
    console.log(`HTML file saved at: ${htmlFilePath}`);

    // Send the file for download
    res.download(htmlFilePath, 'converted-slides.html', (err) => {
      if (err) {
        console.error('Error sending file for download:', err);
        res.status(500).send('Error generating the HTML file.');
      }
    });
  } catch (error) {
    console.error('Error generating HTML file:', error);
    res.status(500).json({ error: 'Error generating HTML file' });
  }
}

async function getSlides(req, res) {
  console.log('Fetching slides...');
  try {
    // Use shared cache instead of local cache
    if (!sharedCache.slidesHTMLCache) {
      console.error('No slides available in cache.');
      return res.status(400).json({ error: 'No slides available. Please upload a PPTX file first.' });
    }

    console.log('Slides fetched successfully.');
    res.json(sharedCache.slidesHTMLCache);
  } catch (error) {
    console.error('Error fetching slides:', error);
    res.status(500).json({ error: 'Error fetching slides.' });
  }
}

async function saveSlides(req, res) {
  try {
    const updatedSlides = req.body.updatedSlides;

    // Check if updatedSlides is defined and an array
    if (!updatedSlides || !Array.isArray(updatedSlides)) {
      console.error('No slides provided or invalid format.');
      return res.status(400).json({ error: 'No slides provided or invalid format.' });
    }

    // Use shared cache instead of local cache
    if (!sharedCache.slidesHTMLCache || !Array.isArray(sharedCache.slidesHTMLCache)) {
      console.error('Slides cache is not initialized.');
      return res.status(400).json({ error: 'Slides cache is not initialized.' });
    }

    updatedSlides.forEach((slideHTML, index) => {
      // Check if slideHTML is not null or empty
      if (!slideHTML || typeof slideHTML !== 'string') {
        console.error(`Invalid slideHTML at index ${index}`);
        return;
      }

      // Match and process base64 images in slideHTML
      const matches = slideHTML.match(/<img[^>]+src="data:image\/[^;]+;base64,[^"]+"/g);
      if (matches) {
        matches.forEach((imgTag) => {
          try {
            const base64Data = imgTag.match(/src="data:image\/[^;]+;base64,([^"]+)"/)[1];
            const buffer = Buffer.from(base64Data, 'base64');

            // Save the image to disk
            const fileName = `slide_${index + 1}_image_${Date.now()}.png`;
            const filePath = path.join(imageSavePath, fileName);

            fsSync.writeFileSync(filePath, buffer);

            const newImgSrc = `/uploads/${fileName}`;
            console.log(`Updated <img> tag src: ${newImgSrc}`);

            // Replace the base64 image source with the new image path
            slideHTML = slideHTML.replace(imgTag, imgTag.replace(/src="[^"]+"/, `src="${newImgSrc}"`));
          } catch (err) {
            console.error('Error processing image tag:', imgTag, err);
          }
        });
      }

      // Update the shared cache instead of local cache
      if (index < sharedCache.slidesHTMLCache.length) {
        sharedCache.slidesHTMLCache[index] = slideHTML;
      } else {
        console.error(`No slide found in cache at index ${index}`);
      }
    });

    console.log('Slides saved successfully.');
    res.json({ status: 'Slides saved successfully.', slides: sharedCache.slidesHTMLCache });
  } catch (error) {
    console.error('Error saving slides:', error);
    res.status(500).json({ error: 'Error saving slides.' });
  }
}

module.exports = {
  uploadFile,
  downloadHTML,
  getSlides,
  saveSlides,
};