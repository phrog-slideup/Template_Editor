const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
const fs = require('fs');

// Initialize S3 client
const s3Client = new S3Client({
    region: process.env.AWS_REGION || 'us-east-1',
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY
    }
});

/**
 * Upload file to S3 bucket
 */
async function uploadToS3(bucketName, folderName, fileName, fileContent, additionalParams = {}) {
    try {
        // If fileContent is a file path, read the file
        const body = typeof fileContent === 'string' && fs.existsSync(fileContent) 
            ? fs.readFileSync(fileContent) 
            : fileContent;

        const params = {
            Bucket: bucketName,
            Key: `${folderName}/${fileName}`,
            Body: body,
            ...additionalParams
        };

        const command = new PutObjectCommand(params);
        await s3Client.send(command);

        // Return the object URL
        return `https://${bucketName}/${folderName}/${fileName}`;

    } catch (error) {
        console.error('S3 Upload Error:', error.message);
        return false;
    }
}

module.exports = { uploadToS3 };