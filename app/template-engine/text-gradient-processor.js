const fsPromises = require('fs').promises;
const path = require('path');
const { postProcessSlideXMLForGradients } = require('./pptxgen-text-gradient-patch');

async function injectTextGradientsIntoSlideXML(slideXmlsDir) {
    try {
        console.log('🎨 Post-processing slides for text gradients...');
        
        const slidesDir = path.join(slideXmlsDir, 'slides');
        const slideFiles = await fsPromises.readdir(slidesDir);
        const slideXmlFiles = slideFiles.filter(file => file.match(/^slide\d+\.xml$/));
        
        let processedCount = 0;
        let totalGradients = 0;
        
        for (const slideFile of slideXmlFiles) {
            const slidePath = path.join(slidesDir, slideFile);
            const slideContent = await fsPromises.readFile(slidePath, 'utf8');
            
            const slideMatch = slideFile.match(/slide(\d+)\.xml/);
            if (!slideMatch) continue;
            
            const slideIndex = parseInt(slideMatch[1]) - 1;
            
            console.log(`\n📄 Processing ${slideFile}...`);
            
            // Post-process the slide XML
            const processedContent = await postProcessSlideXMLForGradients(slideContent, slideIndex);
            
            if (processedContent !== slideContent) {
                await fsPromises.writeFile(slidePath, processedContent, 'utf8');
                processedCount++;
                
                // Count how many gradients were added
                const gradientCount = (processedContent.match(/<a:gradFill/g) || []).length;
                totalGradients += gradientCount;
                
                console.log(`   ✅ Updated ${slideFile} with ${gradientCount} gradient(s)`);
            } else {
                console.log(`   ℹ️  No changes needed for ${slideFile}`);
            }
        }
        
        console.log(`\n   🎉 Gradient post-processing complete:`);
        console.log(`   📊 ${processedCount} slide(s) updated`);
        console.log(`   🎨 ${totalGradients} total gradient(s) injected`);
        
        return {
            success: true,
            slidesProcessed: processedCount,
            gradientsInjected: totalGradients
        };
        
    } catch (error) {
        console.error('   ❌ Error in gradient post-processing:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    injectTextGradientsIntoSlideXML
};