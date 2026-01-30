"use strict";

const path = require("path");
const fs = require("fs").promises;
const fsSync = require("fs");
const axios = require("axios");
const JSZip = require("jszip");
const { v4: uuidv4 } = require("uuid");
const { parseStringPromise } = require("xml2js");
const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
const sharedCache = require('../shared/cache'); // Import shared cache
const { downloadPPTXFromS3 } = require("../helper/s3Helper");

const filesDir = path.resolve(__dirname, "../../files");

// S3 client setup with validation
const s3Client = new S3Client({
    region: process.env.AWS_REGION || 'us-east-1',
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY
    }
});

// Enhanced S3 upload function
async function uploadToS3(bucketName, folderName, fileName, filePath) {
    try {
        // Validate AWS credentials
        if (!process.env.AWS_ACCESS_KEY_ID || !process.env.AWS_SECRET_ACCESS_KEY) {
            throw new Error('AWS credentials not configured');
        }

        // Read file and validate
        if (!fsSync.existsSync(filePath)) {
            throw new Error(`File not found: ${filePath}`);
        }

        const fileBuffer = fsSync.readFileSync(filePath);
        const fileSize = fileBuffer.length;

        if (fileSize === 0) {
            throw new Error('File is empty');
        }

        const params = {
            Bucket: bucketName,
            Key: `${folderName}/${fileName}`,
            Body: fileBuffer,
            ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            Metadata: {
                'upload-time': new Date().toISOString(),
                'file-size': fileSize.toString(),
                'processed': 'true',
                'source': 'html-conversion-with-structure-modification'
            }
        };

        const command = new PutObjectCommand(params);
        const result = await s3Client.send(command);

        const s3Url = `${bucketName}/${folderName}/${fileName}`;
        return s3Url;

    } catch (error) {
        console.error('❌ S3 Upload Error:', error.message);
        console.error('   📊 Error details:', {
            code: error.code,
            statusCode: error.$metadata?.httpStatusCode,
            region: process.env.AWS_REGION,
            bucket: bucketName
        });
        throw error;
    }
}

// HTTP file download
async function downloadFileBuffer(filePath) {
    const response = await axios.get(filePath, { responseType: "arraybuffer", timeout: 30000 });
    return Buffer.from(response.data);
}

function extractS3Details(s3Path) {
    const parts = s3Path.split("/");
    return { bucketName: parts[0], key: parts.slice(1).join("/"), fileName: parts[parts.length - 1] };
}

function getExtensionFromURL(url) {
    const lower = url.toLowerCase();
    const match = lower.match(/\.([a-z0-9]+)(\?|$)/);
    if (match) return match[1];
    if (lower.includes("png")) return "png";
    if (lower.includes("jpeg")) return "jpeg";
    if (lower.includes("jpg")) return "jpg";
    if (lower.includes("gif")) return "gif";
    if (lower.includes("webp")) return "webp";
    return "png";
}

function getMimeType(ext) {
    switch (ext) {
        case "png": return "image/png";
        case "jpg":
        case "jpeg": return "image/jpeg";
        case "gif": return "image/gif";
        case "webp": return "image/webp";
        default: return `image/${ext}`;
    }
}

function normalizeMediaPath(slidePath, target) {
    let cleaned = target.replace(/^\.\.\//, "");
    if (!cleaned.startsWith("ppt/")) cleaned = "ppt/" + cleaned;
    return cleaned;
}

function findMediaFileInZip(zip, rawTarget) {

    // Generate multiple path variations
    const pathVariations = [
        path.posix.normalize(`ppt/${rawTarget.replace("../", "")}`),
        path.posix.normalize(rawTarget.replace(/^\/+/, "")),
        rawTarget.startsWith("/ppt/") ? rawTarget.substring(1) : null,
        rawTarget.replace(/^\/+/, ""),
        !rawTarget.includes("ppt/") ? `ppt/${rawTarget.replace(/^\/+/, "")}` : null,
        `ppt/media/${path.basename(rawTarget)}`,
        normalizeMediaPath("", rawTarget)
    ].filter(Boolean);

    for (const pathVariant of pathVariations) {
        if (zip.file(pathVariant)) {
            return pathVariant;
        }
    }

    const availableFiles = Object.keys(zip.files).filter(k =>
        k.includes('media') || k.includes('image')
    );
    console.warn("File not found in ZIP. Tried paths:", pathVariations);
    console.warn("Available media files:", availableFiles);

    return null;
}

async function findPlaceholderImages(zip, placeholderName) {
    const slidePaths = Object.keys(zip.files).filter(
        p => p.startsWith("ppt/slides/slide") && p.endsWith(".xml")
    );

    for (const slidePath of slidePaths) {
        const slideXml = await zip.file(slidePath).async("string");
        const slideObj = await parseStringPromise(slideXml);

        const spTree = slideObj?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0];
        if (!spTree) continue;

        const shapes = [...(spTree["p:pic"] || []), ...(spTree["p:sp"] || [])];

        for (const pic of shapes) {
            const nvPic = pic["p:nvPicPr"] || pic["p:nvSpPr"];
            if (!nvPic || !nvPic[0] || !nvPic[0]["p:cNvPr"]) continue;

            const name = nvPic[0]["p:cNvPr"][0]["$"]?.name;
            if (name !== placeholderName) continue;

            let primaryEmbed = null;
            let hdEmbed = null;

            const blipFill =
                pic?.["p:blipFill"]?.[0] ||
                pic?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0];

            if (blipFill?.["a:blip"]?.[0]?.["$"]?.["r:embed"]) {
                primaryEmbed = blipFill["a:blip"][0]["$"]["r:embed"];
            }

            const imgLayer = blipFill?.["a:blip"]?.[0]?.["a:extLst"]?.[0]?.["a:ext"]?.[0]?.["a14:imgProps"]?.[0]?.["a14:imgLayer"]?.[0];
            if (imgLayer?.["$"]?.["r:embed"]) {
                hdEmbed = imgLayer["$"]["r:embed"];
            }

            return { slidePath, primaryEmbed, hdEmbed };
        }
    }

    return null;
}

async function replaceSingleImage(zip, placeholderName, imageBuffer, ext) {
    try {

        // FIND SHAPE + rId2 + rId3
        const shapeInfo = await findPlaceholderImages(zip, placeholderName);
        if (!shapeInfo) {
            console.warn(`⚠️ Placeholder "${placeholderName}" not found - skipping`);
            return { success: false, placeholder: placeholderName, error: "Placeholder not found" };
        }

        const { slidePath, primaryEmbed, hdEmbed } = shapeInfo;
        const relsPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
        let relsXml = await zip.file(relsPath).async("string");

        const mime = getMimeType(ext);

        // ===== PRIMARY IMAGE =====
        let primaryTarget = null;
        const primaryRelMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${primaryEmbed}"[^>]+/>`, "i"));
        if (primaryRelMatch) {
            const targetMatch = primaryRelMatch[0].match(/Target="([^"]+)"/i);
            if (targetMatch) {
                primaryTarget = targetMatch[1];
            }
        }

        if (!primaryTarget) {
            console.error(`❌ Primary Relationship missing for embed: ${primaryEmbed}`);
            return { success: false, placeholder: placeholderName, error: "Primary Relationship missing" };
        }

        // Find actual path in ZIP
        let primaryPath = findMediaFileInZip(zip, primaryTarget);
        if (!primaryPath) {
            console.error(`❌ Could not find primary image file for target: ${primaryTarget}`);
            return { success: false, placeholder: placeholderName, error: "Primary image file not found" };
        }

        const baseName = path.basename(primaryPath, path.extname(primaryPath));
        const newPrimaryName = `${baseName}.${ext}`;
        const newPrimaryTarget = primaryTarget.replace(path.basename(primaryTarget), newPrimaryName);
        const newPrimaryPath = primaryPath.replace(path.basename(primaryPath), newPrimaryName);

        // Remove old file if different
        if (primaryPath !== newPrimaryPath && zip.file(primaryPath)) {
            zip.remove(primaryPath);
        }

        // Add new image
        zip.file(newPrimaryPath, imageBuffer);

        // Update relationships XML
        if (primaryTarget !== newPrimaryTarget) {
            relsXml = relsXml.replace(primaryTarget, newPrimaryTarget);
        }

        // ===== HD IMAGE (optional) =====
        if (hdEmbed) {
            let hdTarget = null;
            const hdRelMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id="${hdEmbed}"[^>]+/>`, "i"));
            if (hdRelMatch) {
                const targetMatch = hdRelMatch[0].match(/Target="([^"]+)"/i);
                if (targetMatch) {
                    hdTarget = targetMatch[1];
                }
            }

            if (hdTarget) {
                let hdPath = findMediaFileInZip(zip, hdTarget);

                if (hdPath) {
                    const newHDName = `${baseName}_hd.${ext}`;
                    const newHDTarget = hdTarget.replace(path.basename(hdTarget), newHDName);
                    const newHDPath = hdPath.replace(path.basename(hdPath), newHDName);

                    if (hdPath !== newHDPath && zip.file(hdPath)) {
                        zip.remove(hdPath);
                    }

                    zip.file(newHDPath, imageBuffer);

                    if (hdTarget !== newHDTarget) {
                        relsXml = relsXml.replace(hdTarget, newHDTarget);
                    }
                } else {
                    console.warn(`⚠️ HD image file not found for target: ${hdTarget}`);
                }
            }
        }

        // Update relationships file
        zip.file(relsPath, relsXml);

        // UPDATE CONTENT TYPES
        const ctPath = "[Content_Types].xml";
        let ctXml = await zip.file(ctPath).async("string");
        if (!ctXml.includes(`Extension="${ext}"`)) {
            ctXml = ctXml.replace("</Types>", `  <Default Extension="${ext}" ContentType="${mime}"/>\n</Types>`);
            zip.file(ctPath, ctXml);
        }

        return { success: true, placeholder: placeholderName };

    } catch (error) {
        console.error(`❌ Error replacing image for ${placeholderName}:`, error.message);
        return { success: false, placeholder: placeholderName, error: error.message };
    }
}

async function replaceImage(req, res) {
    const startTime = Date.now();

    try {
        const filePath = req.body.filePath ?? req.query.filePath;
        const mapping = req.body.mapping ?? req.query.mapping;

        // Backward compatibility: support old single image format
        const imageUrl = req.body.imageUrl ?? req.query.imageUrl;
        const placeholderName = req.body.placeholderName ?? req.query.placeholderName;

        // Validate input
        if (!filePath) {
            return res.status(400).json({ error: "filePath is required" });
        }

        let imageMapping = {};

        // Check if using new format (mapping) or old format (single image)
        if (mapping && typeof mapping === 'object') {
            imageMapping = mapping;
        } else if (imageUrl && placeholderName) {
            // Old format - convert to new format
            imageMapping[placeholderName] = imageUrl;
        } else {
            return res.status(400).json({ 
                error: "Either 'mapping' object or both 'imageUrl' and 'placeholderName' are required" 
            });
        }

        if (Object.keys(imageMapping).length === 0) {
            return res.status(400).json({ error: "No images to replace" });
        }

        // Create files directory if needed
        if (!fsSync.existsSync(filesDir)) {
            fsSync.mkdirSync(filesDir, { recursive: true });
        }

        let fileBuffer;
        let originalFileName;

        // LOAD PPTX
        if (filePath?.startsWith("http://") || filePath?.startsWith("https://")) {
            fileBuffer = await downloadFileBuffer(filePath);
            originalFileName = path.basename(filePath);
        } else if (filePath && filePath.includes("/") && !filePath.includes(":")) {
            const s3 = extractS3Details(filePath);
            fileBuffer = await downloadPPTXFromS3(s3.bucketName, s3.key);
            originalFileName = s3.fileName;
        } else if (filePath?.includes(":")) {
            fileBuffer = await fs.readFile(filePath);
            originalFileName = path.basename(filePath);
        } else {
            return res.status(400).json({ error: "Invalid PPTX path" });
        }

        // Load ZIP once
        const zip = await JSZip.loadAsync(fileBuffer);

        // Process each image replacement
        const results = [];
        let successCount = 0;
        let failCount = 0;

        for (const [placeholder, imageUrl] of Object.entries(imageMapping)) {
            try {
                // LOAD IMAGE
                let imageBuffer;
                if (imageUrl.startsWith("http://") || imageUrl.startsWith("https://")) {
                    imageBuffer = Buffer.from((await axios.get(imageUrl, { responseType: "arraybuffer" })).data);
                } else if (imageUrl.includes(":")) {
                    imageBuffer = await fs.readFile(imageUrl);
                } else {
                    console.error(`❌ Invalid imageUrl for ${placeholder}: ${imageUrl}`);
                    results.push({ 
                        success: false, 
                        placeholder, 
                        imageUrl, 
                        error: "Invalid image URL/path" 
                    });
                    failCount++;
                    continue;
                }

                const ext = getExtensionFromURL(imageUrl);

                // Replace the image
                const result = await replaceSingleImage(zip, placeholder, imageBuffer, ext);
                results.push({ ...result, imageUrl });

                if (result.success) {
                    successCount++;
                } else {
                    failCount++;
                }

            } catch (error) {
                console.error(`❌ Error processing ${placeholder}:`, error.message);
                results.push({ 
                    success: false, 
                    placeholder, 
                    imageUrl, 
                    error: error.message 
                });
                failCount++;
            }
        }

        // EXPORT PPTX (save once after all replacements)
        const parsed = path.parse(originalFileName);
        const savedName = `${parsed.name}.pptx`;
        const savedPath = path.join(filesDir, savedName);

        const updatedBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
        await fs.writeFile(savedPath, updatedBuffer);

        // Update cache
        sharedCache.originalPptxPath = savedPath;
        sharedCache.originalFileName = parsed.name;
        sharedCache.fileNameWithoutExt = parsed.name;

        // // UPLOAD TO S3 (upload once)
        // const s3FileName = path.basename(savedPath);
        // const s3Url = await uploadToS3(
        //     'creatos-neo',
        //     'trans-pptx',
        //     s3FileName,
        //     savedPath
        // );

        const totalTime = Date.now() - startTime;

        return res.status(200).json({
            success: true,
            message: `Batch image replacement completed`,
            summary: {
                total: results.length,
                successful: successCount,
                failed: failCount
            },
            results: results,
            updatedPPTX: savedPath,
            // s3Url: s3Url,
            processingTime: totalTime
        });

    } catch (error) {
        console.error("❌ Error ->", error);
        return res.status(500).json({ 
            error: `Error replacing images: ${error.message}`,
            stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
}

module.exports = { replaceImage };