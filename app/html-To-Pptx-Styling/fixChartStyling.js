/**
 * POST-PROCESSING FUNCTION TO FIX CHART XML
 * Add this function to your htmlToPptx.js file
 * Call it after fixChartXmlBugs() in the conversion pipeline
 */

const fsPromises = require('fs').promises;
const path = require('path');

/**
 * Fix chart styling issues in the generated XML
 * This directly modifies the chart XML to match the original styling
 * 
 * @param {string} slideXmlsDir - Directory containing extracted slide XMLs
 * @returns {Promise<Object>} Result with success status and fixes applied
 */
async function fixChartStyling(slideXmlsDir) {
    try {
        const chartsDir = path.join(slideXmlsDir, 'charts');

        // Check if charts directory exists
        const chartsExist = await checkDirectoryExists(chartsDir);
        if (!chartsExist) {
            console.log('   ℹ️  No charts directory found');
            return { success: true, chartsFixed: 0 };
        }

        const chartFiles = await fsPromises.readdir(chartsDir);
        const xmlFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

        if (xmlFiles.length === 0) {
            console.log('   ℹ️  No chart XML files found');
            return { success: true, chartsFixed: 0 };
        }

        let fixedCount = 0;

        for (const chartFile of xmlFiles) {
            try {
                const chartPath = path.join(chartsDir, chartFile);
                let chartXml = await fsPromises.readFile(chartPath, 'utf8');
                const originalXml = chartXml;
                let modified = false;

                console.log(`   🔧 Processing ${chartFile}...`);

                // ============================================
                // FIX #1: REMOVE ALL BORDERS FROM BARS (ANY COLOR)
                // ============================================
                // This handles FFFFFF, F9F9F9, or any other border color
                // Pattern matches any <a:ln> with <a:solidFill> inside <c:ser> sections
                
                // Pattern 1: Full border structure within series
                const borderPatternInSeries = /(<c:ser>[\s\S]*?<c:spPr>[\s\S]*?)<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="[A-Fa-f0-9]{6}" \/>\s*<\/a:solidFill>\s*<a:prstDash val="solid" \/>\s*<a:round \/>\s*<\/a:ln>/g;
                
                if (chartXml.match(borderPatternInSeries)) {
                    chartXml = chartXml.replace(
                        borderPatternInSeries,
                        '$1<a:ln cap="flat">\n                        <a:noFill />\n                        <a:round />\n                    </a:ln>'
                    );
                    console.log('      ✅ Removed borders from bars (full pattern)');
                    modified = true;
                }

                // Pattern 2: Simpler pattern for any border in series
                // This catches variations where formatting might differ
                const simpleBorderPattern = /(<c:ser>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="[A-Fa-f0-9]{6}" \/>\s*<\/a:solidFill>\s*)<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="[A-Fa-f0-9]{6}" \/>\s*<\/a:solidFill>\s*<a:prstDash val="solid" \/>\s*<a:round \/>\s*<\/a:ln>/g;
                
                if (chartXml.match(simpleBorderPattern)) {
                    chartXml = chartXml.replace(
                        simpleBorderPattern,
                        '$1<a:ln cap="flat">\n                        <a:noFill />\n                        <a:round />\n                    </a:ln>'
                    );
                    console.log('      ✅ Removed borders from bars (simple pattern)');
                    modified = true;
                }

                // Pattern 3: Direct replacement for any white/light borders (F9F9F9, FFFFFF, etc.)
                const lightBorderPattern = /<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="F[0-9A-Fa-f]{5}" \/>\s*<\/a:solidFill>\s*<a:prstDash val="solid" \/>\s*<a:round \/>\s*<\/a:ln>/g;
                
                if (chartXml.match(lightBorderPattern)) {
                    chartXml = chartXml.replace(
                        lightBorderPattern,
                        '<a:ln cap="flat">\n                        <a:noFill />\n                        <a:round />\n                    </a:ln>'
                    );
                    console.log('      ✅ Removed light-colored borders from bars');
                    modified = true;
                }

                // ============================================
                // FIX #2: FIX BAR SPACING (gapWidth)
                // ============================================
                // Change gapWidth from 150 to 219
                const gapWidthPattern = /<c:gapWidth val="150" \/>/g;
                if (chartXml.match(gapWidthPattern)) {
                    chartXml = chartXml.replace(gapWidthPattern, '<c:gapWidth val="219" />');
                    console.log('      ✅ Fixed bar gap width (150 → 219)');
                    modified = true;
                }

                // ============================================
                // FIX #3: FIX BAR OVERLAP
                // ============================================
                // Change overlap from 0 to -27
                const overlapPattern = /<c:overlap val="0" \/>/g;
                if (chartXml.match(overlapPattern)) {
                    chartXml = chartXml.replace(overlapPattern, '<c:overlap val="-27" />');
                    console.log('      ✅ Fixed bar overlap (0 → -27)');
                    modified = true;
                }

                // ============================================
                // FIX #4: FIX GRID LINE COLOR (ANY DARK COLOR)
                // ============================================
                // Change grid line color from EEEEEE, 888888, or any dark color to D9D9D9
                // This is in the majorGridlines section
                
                // Pattern 1: Fix 888888 (dark gray)
                const gridColor888Pattern = /<c:majorGridlines>\s*<c:spPr>\s*<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="888888" \/>/g;
                if (chartXml.match(gridColor888Pattern)) {
                    chartXml = chartXml.replace(
                        /<c:majorGridlines>([\s\S]*?)<a:srgbClr val="888888" \/>/g,
                        '<c:majorGridlines>$1<a:srgbClr val="D9D9D9" />'
                    );
                    console.log('      ✅ Fixed grid line color (888888 → D9D9D9)');
                    modified = true;
                }
                
                // Pattern 2: Fix EEEEEE (light gray)
                const gridColorEEEPattern = /<c:majorGridlines>\s*<c:spPr>\s*<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="EEEEEE" \/>/g;
                if (chartXml.match(gridColorEEEPattern)) {
                    chartXml = chartXml.replace(
                        /<c:majorGridlines>([\s\S]*?)<a:srgbClr val="EEEEEE" \/>/g,
                        '<c:majorGridlines>$1<a:srgbClr val="D9D9D9" />'
                    );
                    console.log('      ✅ Fixed grid line color (EEEEEE → D9D9D9)');
                    modified = true;
                }
                
                // Pattern 3: Generic fix for any grid color that's not D9D9D9
                const gridColorGenericPattern = /<c:majorGridlines>\s*<c:spPr>\s*<a:ln w="\d+" cap="flat">\s*<a:solidFill>\s*<a:srgbClr val="(?!D9D9D9)([A-Fa-f0-9]{6})" \/>/g;
                if (chartXml.match(gridColorGenericPattern)) {
                    chartXml = chartXml.replace(
                        gridColorGenericPattern,
                        '<c:majorGridlines>\n                    <c:spPr>\n                        <a:ln w="9525" cap="flat">\n                            <a:solidFill>\n                                <a:srgbClr val="D9D9D9" />'
                    );
                    console.log('      ✅ Fixed grid line color (other → D9D9D9)');
                    modified = true;
                }

                // Also change grid line width from 12700 to 9525 (0.75pt)
                const gridWidthPattern = /<c:majorGridlines>\s*<c:spPr>\s*<a:ln w="12700"/g;
                if (chartXml.match(gridWidthPattern)) {
                    chartXml = chartXml.replace(
                        /(<c:majorGridlines>\s*<c:spPr>\s*<a:ln w=")12700"/g,
                        '$19525"'
                    );
                    console.log('      ✅ Fixed grid line width (12700 → 9525)');
                    modified = true;
                }

                // ============================================
                // FIX #5: FIX AXIS LINE COLORS
                // ============================================
                // Change axis line colors from 888888 to D9D9D9
                const axisColorPattern888 = /<a:srgbClr val="888888" \/>/g;
                if (chartXml.match(axisColorPattern888)) {
                    chartXml = chartXml.replace(
                        axisColorPattern888,
                        '<a:srgbClr val="D9D9D9" />'
                    );
                    console.log('      ✅ Fixed axis line colors (888888 → D9D9D9)');
                    modified = true;
                }

                // Change axis line width from 12700 to 9525
                // Find axis lines (catAx and valAx) and fix their widths
                chartXml = chartXml.replace(
                    /(<c:(?:catAx|valAx)>[\s\S]*?<c:spPr>\s*<a:ln w=")12700"/g,
                    '$19525"'
                );

                // ============================================
                // FIX #6: REMOVE TICK MARKS
                // ============================================
                // Change majorTickMark from "out" to "none"
                const tickMarkOutPattern = /<c:majorTickMark val="out" \/>/g;
                if (chartXml.match(tickMarkOutPattern)) {
                    chartXml = chartXml.replace(
                        tickMarkOutPattern,
                        '<c:majorTickMark val="none" />'
                    );
                    console.log('      ✅ Removed tick marks (out → none)');
                    modified = true;
                }

                // ============================================
                // ADDITIONAL FIX: Ensure proper legend position
                // ============================================
                // Make sure legend is at bottom (b) not right (r)
                const legendPosPattern = /<c:legendPos val="r" \/>/g;
                if (chartXml.match(legendPosPattern)) {
                    chartXml = chartXml.replace(
                        legendPosPattern,
                        '<c:legendPos val="b" />'
                    );
                    console.log('      ✅ Fixed legend position (r → b)');
                    modified = true;
                }

                // ============================================
                // ADDITIONAL FIX: Change roundedCorners for Syncfusion compatibility
                // ============================================
                // Syncfusion has issues with rounded corners, change to square
                const roundedCornersPattern = /<c:roundedCorners val="1" \/>/g;
                if (chartXml.match(roundedCornersPattern)) {
                    chartXml = chartXml.replace(
                        roundedCornersPattern,
                        '<c:roundedCorners val="0" />'
                    );
                    console.log('      ✅ Fixed rounded corners (1 → 0 for Syncfusion)');
                    modified = true;
                }

                // ============================================
                // ADDITIONAL FIX: Add roundedCorners if missing
                // ============================================
                if (!chartXml.includes('<c:roundedCorners')) {
                    // Add after date1904
                    chartXml = chartXml.replace(
                        /(<c:date1904 val="0" \/>)/,
                        '$1\n    <c:roundedCorners val="0" />'
                    );
                    modified = true;
                }

                // Write fixed XML back if modified
                if (modified && chartXml !== originalXml) {
                    await fsPromises.writeFile(chartPath, chartXml, 'utf8');
                    console.log(`   ✅ Fixed chart styling in ${chartFile}`);
                    fixedCount++;
                } else {
                    console.log(`   ✓ No styling issues found in ${chartFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error fixing ${chartFile}: ${error.message}`);
            }
        }

        return {
            success: true,
            chartsFixed: fixedCount,
            totalCharts: xmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in fixChartStyling:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Helper function to check if directory exists
 */
async function checkDirectoryExists(dirPath) {
    try {
        const stat = await fsPromises.stat(dirPath);
        return stat.isDirectory();
    } catch {
        return false;
    }
}

// Export the function
module.exports = {
    fixChartStyling
};