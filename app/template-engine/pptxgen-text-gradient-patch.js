/**
 * Simplified PptxGenJS Text Gradient Patch
 * Direct approach - track gradients by text content matching
 */

// Global registry to track text runs that need gradients
const textGradientRegistry = new Map();

/**
 * Register gradient by unique text content
 */
function registerTextGradient(textContent, gradientData) {
    // Use first 50 chars of text as key to avoid duplicates
    const key = textContent.substring(0, 50).trim();
    if (key) {
        textGradientRegistry.set(key, gradientData);
        console.log(`📝 Registered gradient for text: "${key.substring(0, 20)}..."`);
    }
}

/**
 * Clear all registered gradients
 */
function clearGradientMetadata() {
    textGradientRegistry.clear();
    console.log('🧹 Cleared gradient metadata');
}

/**
 * Generate OOXML gradient fill XML
 */
function generateGradientFillXML(gradientData) {
    if (!gradientData || !gradientData.stops || gradientData.stops.length < 2) {
        return null;
    }

    const { type, angle, stops, path, centerX, centerY } = gradientData;
    
    let xml = '<a:gradFill rotWithShape="1">';
    xml += '<a:gsLst>';
    
    stops.forEach(stop => {
        const position = Math.round(stop.position * 1000);
        const color = stop.color.replace('#', '').toUpperCase();
        
        xml += `<a:gs pos="${position}">`;
        xml += `<a:srgbClr val="${color}"/>`;
        xml += `</a:gs>`;
    });
    
    xml += '</a:gsLst>';
    
    if (type === 'linear') {
        const cssAngle = angle || 90;
        let ooxmlDegrees = (cssAngle-90) % 360;
        if (ooxmlDegrees < 0) ooxmlDegrees += 360;
        const ooxmlAngle = Math.round(ooxmlDegrees * 60000);
        
        xml += `<a:lin ang="${ooxmlAngle}" scaled="0"/>`;
        
    } else if (type === 'radial') {
        const pathType = path || 'circle';
        xml += `<a:path path="${pathType}">`;
        
        const cx = centerX !== undefined ? centerX : 50;
        const cy = centerY !== undefined ? centerY : 50;
        
        const l = Math.round(cx * 1000);
        const t = Math.round(cy * 1000);
        const r = Math.round((100 - cx) * 1000);
        const b = Math.round((100 - cy) * 1000);
        
        xml += `<a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/>`;
        xml += `</a:path>`;
    }
    
    xml += '</a:gradFill>';
    
    return xml;
}


async function postProcessSlideXMLForGradients(slideXmlContent, slideIndex) {
    try {
        if (textGradientRegistry.size === 0) {
            console.log(`   ℹ️  Slide ${slideIndex + 1}: No gradients to inject`);
            return slideXmlContent;
        }

        let modifiedXml = slideXmlContent;
        let gradientsInjected = 0;

        console.log(`   🔍 Slide ${slideIndex + 1}: Processing ${textGradientRegistry.size} gradient texts...`);

        // For each registered gradient text
        for (const [textKey, gradientData] of textGradientRegistry.entries()) {
            // Escape special regex characters in text
            const escapedText = textKey
                .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                .substring(0, 30); // Use first 30 chars for matching
            
            const pattern = new RegExp(
                '(<a:r>\\s*' +                          // Start of text run
                '<a:rPr[^>]*>\\s*' +                   // Run properties start
                '(?:(?!<a:solidFill>).)*?' +           // Any content before solidFill (non-greedy)
                ')' +                                   // GROUP 1: everything before solidFill
                '(<a:solidFill>.*?</a:solidFill>)' +  // GROUP 2: solidFill to replace
                '(' +                                   // GROUP 3: everything after solidFill
                '.*?' +                                 // Any other run properties
                '</a:rPr>\\s*' +                       // Close run properties
                '<a:t[^>]*>' +                         // Text element
                escapedText +                           // Match the text content
                '.*?' +                                 // Rest of text
                '</a:t>\\s*' +                         // Close text
                '</a:r>)',                             // Close run
                'gs'                                    // Global, singleline (. matches newline)
            );

            let matchCount = 0;
            
            // Replace all matches
            modifiedXml = modifiedXml.replace(pattern, (fullMatch, beforeFill, solidFill, afterFillAndRest) => {
                const gradientXml = generateGradientFillXML(gradientData);
                
                if (gradientXml) {
                    matchCount++;
                    gradientsInjected++;
                    console.log(`   ✅ Gradient #${gradientsInjected}: "${textKey.substring(0, 35)}..."`);
                    // Replace solidFill with gradFill, keep everything else
                    return beforeFill + gradientXml + afterFillAndRest;
                }
                
                return fullMatch;
            });

            if (matchCount > 0) {
                console.log(`   📊 Replaced ${matchCount} instance(s) of "${textKey.substring(0, 25)}..."`);
            } else {
                console.log(`   ⚠️  No XML match for: "${textKey.substring(0, 35)}..."`);
            }
        }

        if (gradientsInjected > 0) {
            console.log(`   🎉 Slide ${slideIndex + 1}: ${gradientsInjected} gradients successfully injected`);
        } else {
            console.log(`   ⚠️  Slide ${slideIndex + 1}: No gradients were injected - check XML structure`);
        }

        return modifiedXml;
        
    } catch (error) {
        console.error(`   ❌ Error post-processing slide ${slideIndex + 1}:`, error);
        return slideXmlContent; // Return original on error
    }
}
module.exports = {
    clearGradientMetadata,
    registerTextGradient,
    generateGradientFillXML,
    postProcessSlideXMLForGradients
};