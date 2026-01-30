const { S3Client, GetObjectCommand } = require('@aws-sdk/client-s3');
const { getSignedUrl } = require('@aws-sdk/s3-request-presigner');

// Validate and get AWS region
function getValidAWSRegion() {
    const region = process.env.AWS_REGION;
        
    // List of valid AWS regions
    const validRegions = [
        'us-east-1', 'us-east-2', 'us-west-1', 'us-west-2',
        'eu-west-1', 'eu-west-2', 'eu-central-1', 'ap-south-1',
        'ap-southeast-1', 'ap-southeast-2', 'ap-northeast-1'
    ];
    
    // If no region specified or invalid, default to us-east-1
    if (!region || typeof region !== 'string' || region.trim() === '' || !validRegions.includes(region)) {
        console.warn(`⚠️ Invalid or missing AWS_REGION: "${region}". Using default: us-east-1`);
        return 'us-east-1';
    }
    
    // console.log(`✅ Using valid AWS region: ${region}`);
    return region;
}

// Configure AWS SDK v3 Client
const s3Client = new S3Client({
    region: getValidAWSRegion(),
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
    // Additional configuration for better error handling
    maxAttempts: 3,
    retryMode: 'adaptive',
});

// Convert stream to buffer utility
async function streamToBuffer(stream) {
    const chunks = [];
    return new Promise((resolve, reject) => {
        stream.on('data', (chunk) => chunks.push(chunk));
        stream.on('error', reject);
        stream.on('end', () => resolve(Buffer.concat(chunks)));
    });
}


function parseS3Url(s3Url) {
    try {
        let bucketName, key, region;

        if (s3Url.startsWith('s3://')) {
            // Handle s3:// protocol URLs
            const withoutProtocol = s3Url.replace('s3://', '');
            const parts = withoutProtocol.split('/');
            bucketName = parts[0];
            key = parts.slice(1).join('/');
        } else if (s3Url.includes('.s3.amazonaws.com') || s3Url.includes('s3.amazonaws.com')) {
            // Handle HTTPS S3 URLs
            const url = new URL(s3Url);
            
            if (url.hostname.includes('.s3.amazonaws.com')) {
                // Virtual-hosted-style URL: https://bucket.s3.amazonaws.com/path/file.pptx
                bucketName = url.hostname.split('.s3.amazonaws.com')[0];
            } else if (url.hostname.includes('s3.amazonaws.com')) {
                // Path-style URL: https://s3.amazonaws.com/bucket/path/file.pptx
                const pathParts = url.pathname.split('/').filter(part => part);
                bucketName = pathParts[0];
                key = pathParts.slice(1).join('/');
            }
            
            if (!key) {
                key = url.pathname.startsWith('/') ? url.pathname.slice(1) : url.pathname;
            }
            
            // Extract region if present
            const regionMatch = url.hostname.match(/s3\.([^.]+)\.amazonaws\.com/);
            if (regionMatch) {
                region = regionMatch[1];
            }
        } else {
            throw new Error('Invalid S3 URL format');
        }

        if (!bucketName || !key) {
            throw new Error('Could not extract bucket name and key from S3 URL');
        }

        console.log(`Parsed S3 URL - Bucket: ${bucketName}, Key: ${key}, Region: ${region || 'default'}`);
        
        return {
            bucketName,
            key,
            region: region || process.env.AWS_REGION || 'us-east-1'
        };
    } catch (error) {
        console.error('Error parsing S3 URL:', error);
        throw new Error(`Failed to parse S3 URL: ${s3Url}`);
    }
}


async function downloadPPTXFromS3(bucketName, key) {
    try {
        console.log(`🔍 Starting download from S3:`);
        console.log(`   📁 Bucket: ${bucketName}`);
        console.log(`   🔑 Key: ${key}`);
        console.log(`   🌍 Region: ${getValidAWSRegion()}`);
        
        // Validate inputs
        if (!bucketName || typeof bucketName !== 'string') {
            throw new Error(`Invalid bucket name: ${bucketName}`);
        }
        
        if (!key || typeof key !== 'string') {
            throw new Error(`Invalid key: ${key}`);
        }
        
        const command = new GetObjectCommand({
            Bucket: bucketName.trim(),
            Key: key.trim(),
        });

        console.log(`⚡ Sending request to S3...`);
        const response = await s3Client.send(command);
        
        if (!response.Body) {
            throw new Error('No data received from S3');
        }

        console.log(`📦 Received response, converting stream to buffer...`);
        const buffer = await streamToBuffer(response.Body);
        
        // Validate it's a PPTX file by checking magic bytes
        if (!isPPTXFile(buffer)) {
            throw new Error('Downloaded file is not a valid PPTX file');
        }

        console.log(`✅ Successfully downloaded PPTX file: ${buffer.length} bytes`);
        return buffer;
    } catch (error) {
        console.error(`❌ Error downloading PPTX from S3:`);
        console.error(`   📁 Bucket: ${bucketName}`);
        console.error(`   🔑 Key: ${key}`);
        console.error(`   🌍 Region: ${getValidAWSRegion()}`);
        console.error(`   ⚠️ Error: ${error.message}`);
        
        // Provide more specific error messages
        if (error.name === 'NoSuchKey') {
            throw new Error(`File not found in S3: ${bucketName}/${key}`);
        } else if (error.name === 'NoSuchBucket') {
            throw new Error(`Bucket not found: ${bucketName}`);
        } else if (error.name === 'AccessDenied') {
            throw new Error(`Access denied to S3 resource: ${bucketName}/${key}`);
        } else if (error.message.includes('region')) {
            throw new Error(`AWS Region error: ${error.message}. Please check your AWS_REGION environment variable.`);
        }
        
        throw new Error(`Failed to download PPTX file from S3: ${error.message}`);
    }
}


async function downloadPPTXFromS3Url(s3Url) {
    try {
        console.log(`Downloading PPTX from S3 URL: ${s3Url}`);
        
        const { bucketName, key } = parseS3Url(s3Url);
        return await downloadPPTXFromS3(bucketName, key);
    } catch (error) {
        console.error(`Error downloading PPTX from S3 URL (${s3Url}):`, error);
        throw new Error(`Failed to download PPTX file from S3 URL: ${error.message}`);
    }
}


async function createPresignedDownloadUrl(bucketName, key, expiresIn = 3600) {
    try {
        console.log(`Creating presigned download URL for: ${bucketName}/${key}`);
        
        const command = new GetObjectCommand({
            Bucket: bucketName,
            Key: key,
        });

        const signedUrl = await getSignedUrl(s3Client, command, {
            expiresIn, // URL expires in seconds
        });

        console.log(`✅ Presigned URL created successfully (expires in ${expiresIn} seconds)`);
        console.log(`🔗 Presigned Download URL: ${signedUrl}`);
        return signedUrl;
    } catch (error) {
        console.error(`Error creating presigned URL for ${bucketName}/${key}:`, error);
        throw new Error(`Failed to create presigned download URL: ${error.message}`);
    }
}


async function createPresignedPPTXDownloadUrl(bucketName, key, options = {}) {
    const {
        expiresIn = 3600,
        downloadFilename,
        forceDownload = true
    } = options;

    try {
        console.log(`Creating presigned PPTX download URL for: ${bucketName}/${key}`);
        
        const commandParams = {
            Bucket: bucketName,
            Key: key,
        };

        // Set proper PPTX content type
        commandParams.ResponseContentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

        // Force download with custom filename if specified
        if (forceDownload || downloadFilename) {
            const filename = downloadFilename || key.split('/').pop();
            commandParams.ResponseContentDisposition = `attachment; filename="${filename}"`;
        }

        const command = new GetObjectCommand(commandParams);
        const signedUrl = await getSignedUrl(s3Client, command, { expiresIn });

        console.log(`✅ Presigned PPTX download URL created successfully`);
        console.log(`🔗 PPTX Download URL: ${signedUrl}`);
        console.log(`📁 Bucket: ${bucketName}`);
        console.log(`🔑 Key: ${key}`);
        console.log(`⏰ Expires in: ${expiresIn} seconds`);
        if (downloadFilename) {
            console.log(`📄 Download as: ${downloadFilename}`);
        }
        
        return signedUrl;
    } catch (error) {
        console.error(`Error creating presigned PPTX URL for ${bucketName}/${key}:`, error);
        throw new Error(`Failed to create presigned PPTX download URL: ${error.message}`);
    }
}


async function createPresignedUrlFromS3Url(s3Url, options = {}) {
    try {
        console.log(`🌐 Creating presigned URL from S3 URL: ${s3Url}`);
        const { bucketName, key } = parseS3Url(s3Url);
        const signedUrl = await createPresignedPPTXDownloadUrl(bucketName, key, options);
        console.log(`✅ Presigned URL from S3 URL created successfully: ${signedUrl}`);
        return signedUrl;
    } catch (error) {
        console.error(`Error creating presigned URL from S3 URL (${s3Url}):`, error);
        throw new Error(`Failed to create presigned URL from S3 URL: ${error.message}`);
    }
}


function isPPTXFile(buffer) {
    if (!buffer || buffer.length < 4) {
        return false;
    }

    // Check for ZIP magic number (PPTX files are ZIP archives)
    const zipSignature = [0x50, 0x4B, 0x03, 0x04]; // "PK" + version
    const zipSignature2 = [0x50, 0x4B, 0x05, 0x06]; // "PK" + end of central directory
    const zipSignature3 = [0x50, 0x4B, 0x07, 0x08]; // "PK" + data descriptor

    const firstFourBytes = Array.from(buffer.slice(0, 4));
    
    return (
        JSON.stringify(firstFourBytes) === JSON.stringify(zipSignature) ||
        JSON.stringify(firstFourBytes) === JSON.stringify(zipSignature2) ||
        JSON.stringify(firstFourBytes) === JSON.stringify(zipSignature3)
    );
}


async function getFileMetadata(bucketName, key) {
    try {
        const command = new GetObjectCommand({
            Bucket: bucketName,
            Key: key,
        });

        const response = await s3Client.send(command);
        
        return {
            contentType: response.ContentType,
            contentLength: response.ContentLength,
            lastModified: response.LastModified,
            etag: response.ETag,
            metadata: response.Metadata,
        };
    } catch (error) {
        console.error(`Error getting file metadata for ${bucketName}/${key}:`, error);
        throw new Error(`Failed to get file metadata: ${error.message}`);
    }
}

module.exports = {
    // Core download functions (used by your controller)
    downloadPPTXFromS3,
    downloadPPTXFromS3Url,
    parseS3Url,
    
    // Presigned URL functions
    createPresignedDownloadUrl,
    createPresignedPPTXDownloadUrl,
    createPresignedUrlFromS3Url,
    
    // Utility functions
    streamToBuffer,
    isPPTXFile,
    getFileMetadata,
    getValidAWSRegion,
    
    // S3 client (if needed elsewhere)
    s3Client
};