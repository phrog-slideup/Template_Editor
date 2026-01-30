const cron = require('node-cron');
const fs = require('fs');
const path = require('path');

// ============ CONFIGURATION ============
const filesDir = path.join(__dirname, '../files');  // Files directory path
const uploadsDir = path.join(__dirname, '../uploads');  // Uploads directory path
const hoursOld = 2;  // Delete files older than 2 hours
const schedule = '0 * * * *';  // Run every hour (change as needed)

function cleanupOldFiles() {
    console.log(`\n[${new Date().toLocaleString()}] 🧹 Starting cleanup...`);

    try {
        let totalDeleted = 0;

        // ===== CLEANUP FILES DIRECTORY =====
        if (!fs.existsSync(filesDir)) {
            console.log('Files directory not found');
        } else {
            const items = fs.readdirSync(filesDir);
            let deleted = 0;

            items.forEach(item => {
                const itemPath = path.join(filesDir, item);
                const stats = fs.statSync(itemPath);
                
                // Calculate age in hours
                const ageHours = (Date.now() - stats.mtime.getTime()) / (1000 * 60 * 60);

                // Delete if older than threshold
                if (ageHours > hoursOld) {
                    try {
                        if (stats.isDirectory()) {
                            // Delete directory
                            fs.rmSync(itemPath, { recursive: true, force: true });
                            deleted++;
                        } else if (item.endsWith('.pptx') || item.endsWith('.zip')) {
                            // Delete PPTX or ZIP file
                            fs.unlinkSync(itemPath);
                            deleted++;
                        }
                    } catch (err) {
                        console.log(`❌ Error deleting ${item}:`, err.message);
                    }
                }
            });

            totalDeleted += deleted;
        }

        // ===== CLEANUP UPLOADS DIRECTORY =====
        if (!fs.existsSync(uploadsDir)) {
            console.log('Uploads directory not found');
        } else {
            const items = fs.readdirSync(uploadsDir);
            let deleted = 0;

            items.forEach(item => {
                const itemPath = path.join(uploadsDir, item);
                const stats = fs.statSync(itemPath);
                
                // Calculate age in hours
                const ageHours = (Date.now() - stats.mtime.getTime()) / (1000 * 60 * 60);

                // Delete if older than threshold
                if (ageHours > hoursOld) {
                    try {
                        if (stats.isDirectory()) {
                            // Delete directory
                            fs.rmSync(itemPath, { recursive: true, force: true });
                            deleted++;
                        } else if (item.match(/\.(jpg|jpeg|png|gif|svg|webp|bmp|ico)$/i)) {
                            // Delete image files
                            fs.unlinkSync(itemPath);
                            deleted++;
                        }
                    } catch (err) {
                        console.log(`❌ Error deleting ${item}:`, err.message);
                    }
                }
            });

            totalDeleted += deleted;
        }

    } catch (error) {
        console.log('❌ Cleanup error:', error.message);
    }
}

// Run cleanup immediately on start
cleanupOldFiles();

// Schedule automatic cleanup
cron.schedule(schedule, cleanupOldFiles);

// Handle shutdown
process.on('SIGINT', () => {
    console.log('\n👋 Scheduler stopped');
    process.exit(0);
});