/**
 * addChartToSlide.js
 * Converts HTML charts to PPTX charts using pptxgenjs
 * Node.js/JSDOM compatible version - CORRECTED
 */

function addChartToSlide(pptx, pptSlide, element, slideContext) {
    try {
        // Check if this is a chart container
        if (!isChartElement(element)) {
            return false;
        }

        // Extract chart data from HTML
        const chartData = extractChartDataFromHTML(element);
        if (!chartData || !chartData.isValid) {
            console.log("   ❌ Failed to extract valid chart data");
            return false;
        }

        // Extract positioning and styling
        const position = extractChartPosition(element, slideContext);
        const styling = extractChartStyling(element, chartData);

        // Convert to pptxgenjs format
        const pptxChartData = convertToPptxFormat(chartData);

        // Create chart options
        const chartOptions = createChartOptions(chartData, position, styling);

        // Add chart to slide
        pptSlide.addChart(getChartType(chartData.type), pptxChartData, chartOptions);

        console.log("   ✅ Chart added to slide successfully");
        return true;

    } catch (error) {
        console.error("   ❌ Error adding chart to slide:", error);
        return false;
    }
}

function isChartElement(element) {
    if (!element || !element.classList) return false;

    // Check for chart container classes
    const chartClasses = [
        'chart-container',
        'chart',
        'graph',
        'plot',
        'visualization'
    ];

    if (chartClasses.some(cls => element.classList.contains(cls))) {
        return true;
    }

    // Check for chart-specific elements inside
    const hasChartElements =
        element.querySelector('.bar') ||
        element.querySelector('.chart-area') ||
        element.querySelector('[data-value]') ||
        element.querySelector('.category-label') ||
        element.querySelector('canvas[data-chart]') ||
        element.querySelector('svg[data-chart]');

    return !!hasChartElements;
}

function extractChartDataFromHTML(chartContainer) {
    try {
        // Extract title
        const titleElement = chartContainer.querySelector('.chart-title, .title');
        const title = titleElement ? titleElement.textContent.trim() : 'Chart';

        // Detect chart type
        let chartType = detectChartType(chartContainer);

        // Extract data based on chart type
        let extractedData;
        switch (chartType) {
            case 'bar':
            case 'column':
                extractedData = extractBarChartData(chartContainer);
                break;
            case 'line':
                extractedData = extractLineChartData(chartContainer);
                break;
            case 'pie':
                extractedData = extractPieChartData(chartContainer);
                break;
            default:
                extractedData = extractBarChartData(chartContainer); // Fallback to bar chart
        }

        return {
            ...extractedData,
            title,
            type: chartType,
            isValid: extractedData && extractedData.categories && extractedData.series && extractedData.series.length > 0
        };

    } catch (error) {
        console.error("Error extracting chart data:", error);
        return { isValid: false };
    }
}

function detectChartType(chartContainer) {
    // Check for specific chart type indicators
    if (chartContainer.querySelector('.bar, [data-series]')) {
        return 'bar';
    }
    if (chartContainer.querySelector('.line, .path, polyline')) {
        return 'line';
    }
    if (chartContainer.querySelector('.pie, .slice, .arc')) {
        return 'pie';
    }
    if (chartContainer.querySelector('.scatter, .dot')) {
        return 'scatter';
    }

    // Check CSS classes for type hints
    const classList = chartContainer.className.toLowerCase();
    if (classList.includes('bar') || classList.includes('column')) return 'bar';
    if (classList.includes('line')) return 'line';
    if (classList.includes('pie') || classList.includes('donut')) return 'pie';
    if (classList.includes('scatter')) return 'scatter';

    // FIXED: Default to bar chart instead of null
    return 'bar';
}

function extractBarChartData(chartContainer) {
    const categories = [];
    const seriesMap = new Map();
    const seriesColors = new Map();
    const seriesNames = new Map();

    try {
        // Extract category labels
        const categoryElements = chartContainer.querySelectorAll('.category-label');

        if (categoryElements.length > 0) {
            for (let i = 0; i < categoryElements.length; i++) {
                const el = categoryElements[i];
                const text = el.textContent ? el.textContent.trim() : '';
                if (text) {
                    categories.push(text);
                }
            }
        }

        // Extract bars and their data
        const bars = chartContainer.querySelectorAll('.bar[data-value][data-series][data-category]');

        if (bars.length > 0) {
            for (let i = 0; i < bars.length; i++) {
                const bar = bars[i];

                const value = parseFloat(bar.getAttribute('data-value')) || 0;
                const seriesIndex = parseInt(bar.getAttribute('data-series')) || 0;
                const categoryIndex = parseInt(bar.getAttribute('data-category')) || 0;

                // Extract color from style
                const color = extractColorFromElement(bar);

                // Extract series name from title attribute
                const title = bar.getAttribute('title') || '';
                const seriesName = title.split(':')[0].trim() || `Series ${seriesIndex + 1}`;

                // Initialize series if not exists
                if (!seriesMap.has(seriesIndex)) {
                    seriesMap.set(seriesIndex, new Map());
                    seriesColors.set(seriesIndex, color);
                    seriesNames.set(seriesIndex, seriesName);
                }

                // Store value for this series and category
                seriesMap.get(seriesIndex).set(categoryIndex, value);
            }
        }

        // If no categories found from labels, generate them based on the data
        if (categories.length === 0) {
            const maxCategoryIndex = Math.max(...Array.from(seriesMap.values()).map(categoryMap =>
                Math.max(...categoryMap.keys())
            ));
            for (let i = 0; i <= maxCategoryIndex; i++) {
                categories.push(`Category ${i + 1}`);
            }
        }

        // Convert series map to array format
        const series = [];
        const sortedSeriesIndices = Array.from(seriesMap.keys()).sort((a, b) => a - b);

        sortedSeriesIndices.forEach(seriesIndex => {
            const categoryMap = seriesMap.get(seriesIndex);
            const values = [];

            // Fill values array in category order
            for (let catIndex = 0; catIndex < categories.length; catIndex++) {
                values.push(categoryMap.get(catIndex) || 0);
            }

            series.push({
                name: seriesNames.get(seriesIndex) || `Series ${seriesIndex + 1}`,
                values: values,
                color: normalizeColor(seriesColors.get(seriesIndex))
            });
        });

        // Calculate min/max values
        const allValues = series.flatMap(s => s.values).filter(v => !isNaN(v));
        const maxValue = allValues.length > 0 ? Math.max(...allValues) : 10;
        const minValue = allValues.length > 0 ? Math.min(...allValues, 0) : 0;

        return {
            categories,
            series,
            maxValue,
            minValue
        };

    } catch (error) {
        console.error("Error in extractBarChartData:", error);
        return {
            categories: ['Category 1'],
            series: [{ name: 'Series 1', values: [1], color: '#4472C4' }],
            maxValue: 1,
            minValue: 0
        };
    }
}

function extractColorFromElement(element) {
    try {
        // Check inline style attribute
        const styleAttr = element.getAttribute('style');
        if (styleAttr) {
            // Parse background-color
            const bgColorMatch = styleAttr.match(/background-color:\s*([^;]+)/i);
            if (bgColorMatch) {
                return bgColorMatch[1].trim();
            }
        }

        // Check computed style if available
        if (element.style && element.style.backgroundColor) {
            return element.style.backgroundColor;
        }

        return '#4472C4'; // Default color

    } catch (error) {
        console.error("Error extracting color:", error);
        return '#4472C4';
    }
}

function extractLineChartData(chartContainer) {
    // Fallback to bar chart extraction for now
    return extractBarChartData(chartContainer);
}

function extractPieChartData(chartContainer) {
    // Fallback to bar chart extraction for now
    return extractBarChartData(chartContainer);
}

function extractChartPosition(chartContainer, slideContext) {
    try {
        // First, try to get the actual chart area dimensions if available
        const chartArea = chartContainer.querySelector('.chart-area');
        let left = 0, top = 0, width = 400, height = 300;

        // Extract container position
        const containerStyle = chartContainer.getAttribute('style');
        if (containerStyle) {
            const leftMatch = containerStyle.match(/left:\s*([0-9]*\.?[0-9]+)px/i);
            const topMatch = containerStyle.match(/top:\s*([0-9]*\.?[0-9]+)px/i);

            if (leftMatch) left = parseFloat(leftMatch[1]) || 0;
            if (topMatch) top = parseFloat(topMatch[1]) || 0;
        }

        // Use the full container dimensions for better sizing
        if (containerStyle) {
            const widthMatch = containerStyle.match(/width:\s*([0-9]*\.?[0-9]+)px/i);
            const heightMatch = containerStyle.match(/height:\s*([0-9]*\.?[0-9]+)px/i);

            if (widthMatch) width = parseFloat(widthMatch[1]) || 400;
            if (heightMatch) height = parseFloat(heightMatch[1]) || 300;
        }

        // COMPLETELY REVISED APPROACH: Use proportional sizing instead of exact pixel conversion
        // Get slide dimensions - try to detect from HTML structure
        let slideWidthPx = 960;
        let slideHeightPx = 540;

        const slideElement = chartContainer.closest('.sli-slide');
        if (slideElement) {
            const slideStyle = slideElement.getAttribute('style');
            if (slideStyle) {
                const slideWidthMatch = slideStyle.match(/width:\s*([0-9]*\.?[0-9]+)px/i);
                const slideHeightMatch = slideStyle.match(/height:\s*([0-9]*\.?[0-9]+)px/i);
                if (slideWidthMatch) slideWidthPx = parseFloat(slideWidthMatch[1]);
                if (slideHeightMatch) slideHeightPx = parseFloat(slideHeightMatch[1]);
            }
        }

        // Calculate proportions of the chart relative to the slide
        const widthRatio = width / slideWidthPx;
        const heightRatio = height / slideHeightPx;
        const leftRatio = left / slideWidthPx;
        const topRatio = top / slideHeightPx;

        // FIXED: Use standard PowerPoint slide dimensions (10" x 7.5" for 4:3 or 13.33" x 7.5" for 16:9)
        // Most modern presentations use 16:9, so let's use that
        const pptSlideWidthIn = 13.33;
        const pptSlideHeightIn = 7.5;

        // Apply the same proportions to PowerPoint slide
        let pptX = leftRatio * pptSlideWidthIn;
        let pptY = topRatio * pptSlideHeightIn;
        let pptW = widthRatio * pptSlideWidthIn;
        let pptH = heightRatio * pptSlideHeightIn;

        // CRITICAL FIX: Apply size multiplier if chart is still too small
        // Based on your screenshots, we need to make it significantly larger
        const minVisibleWidth = 4.5;  // Minimum 4.5 inches to be clearly visible
        const minVisibleHeight = 3.0; // Minimum 3 inches to be clearly visible

        if (pptW < minVisibleWidth) {
            const scaleFactor = minVisibleWidth / pptW;
            pptW = minVisibleWidth;
            pptH *= scaleFactor; // Scale height proportionally
        }

        if (pptH < minVisibleHeight) {
            const scaleFactor = minVisibleHeight / pptH;
            pptH = minVisibleHeight;
            pptW *= scaleFactor; // Scale width proportionally  
        }

        // Ensure chart fits within slide boundaries after scaling
        if (pptX + pptW > pptSlideWidthIn) {
            if (pptW > pptSlideWidthIn - 1) {
                // Chart is wider than slide, scale down
                const scale = (pptSlideWidthIn - 1) / pptW;
                pptW = pptSlideWidthIn - 1;
                pptH *= scale;
                pptX = 0.5;
            } else {
                // Just reposition
                pptX = pptSlideWidthIn - pptW - 0.5;
            }
        }

        if (pptY + pptH > pptSlideHeightIn) {
            if (pptH > pptSlideHeightIn - 1) {
                // Chart is taller than slide, scale down
                const scale = (pptSlideHeightIn - 1) / pptH;
                pptH = pptSlideHeightIn - 1;
                pptW *= scale;
                pptY = 0.5;
            } else {
                // Just reposition
                pptY = pptSlideHeightIn - pptH - 0.5;
            }
        }

        // Final positioning with safety margins
        const finalPosition = {
            x: Math.max(0.25, pptX),
            y: Math.max(0.25, pptY),
            w: pptW,
            h: pptH
        };

        // Additional debugging info
        return finalPosition;

    } catch (error) {
        console.error("Error extracting position:", error);
        // Return a large, visible fallback
        return { x: 2, y: 2, w: 7, h: 4.5 };
    }
}

function extractChartStyling(chartContainer, chartData) {
    const styling = {
        colors: [],
        fontSize: 12,
        fontFamily: 'Arial',
        showLegend: true,
        showGridLines: true,
        showDataLabels: false
    };

    try {
        // Extract colors from series
        styling.colors = chartData.series.map(s => s.color || '#4472C4');

        // Extract font styling from title element
        const titleElement = chartContainer.querySelector('.chart-title');
        if (titleElement) {
            const styleAttr = titleElement.getAttribute('style');
            if (styleAttr) {
                const fontSizeMatch = styleAttr.match(/font-size:\s*([^;]+)/i);
                const fontFamilyMatch = styleAttr.match(/font-family:\s*([^;]+)/i);

                if (fontSizeMatch) styling.fontSize = parseInt(fontSizeMatch[1]) || 12;
                if (fontFamilyMatch) styling.fontFamily = fontFamilyMatch[1].trim().replace(/['"]/g, '') || 'Arial';
            }
        }

        // Check for grid lines - if axis labels exist, assume grid lines are wanted
        const axisLabels = chartContainer.querySelectorAll('.axis-label');
        styling.showGridLines = axisLabels.length > 0;

    } catch (error) {
        console.error("Error extracting styling:", error);
    }

    return styling;
}

// FIXED: Correct PPTX data format matching pptxgenjs expectations
// function convertToPptxFormat(chartData) {
//     try {
//         // Validate input data
//         if (!chartData || !chartData.categories || !chartData.series) {
//             console.error('Invalid chart data structure');
//             return getDefaultChartData();
//         }

//         if (chartData.categories.length === 0 || chartData.series.length === 0) {
//             console.error('Chart data is empty');
//             return getDefaultChartData();
//         }

//         // Validate that all series have the same number of values as categories
//         const expectedLength = chartData.categories.length;
//         const validSeries = chartData.series.filter(series => {
//             if (!series || !Array.isArray(series.values)) {
//                 console.warn(`Invalid series found:`, series);
//                 return false;
//             }
//             if (series.values.length !== expectedLength) {
//                 console.warn(`Series ${series.name} has ${series.values.length} values but expected ${expectedLength}`);
//             }
//             return true;
//         });

//         if (validSeries.length === 0) {
//             console.error('No valid series found');
//             return getDefaultChartData();
//         }

//         // CORRECTED: Use the exact format pptxgenjs expects
//         // Each series object needs: name, labels (array of arrays), values (array)
//         const pptxData = validSeries.map((series, idx) => {
//             const cleanValues = series.values.map(v => (typeof v === 'number' && !isNaN(v)) ? v : 0);
//             const base = { name: series.name || 'Unnamed Series', values: cleanValues };
//             if (idx === 0) {
//                 // Only first series carries labels
//                 base.labels = [chartData.categories.slice()];
//             }
//             return base;
//         });


//         // Validate final data structure
//         if (!validatePptxData(pptxData)) {
//             console.error('Generated invalid PPTX data, using fallback');
//             return getDefaultChartData();
//         }

//         return pptxData;

//     } catch (error) {
//         console.error('Error converting to PPTX format:', error);
//         return getDefaultChartData();
//     }
// }

function convertToPptxFormat(chartData) {
    try {
        if (!chartData || !Array.isArray(chartData.categories) || !Array.isArray(chartData.series)) {
            return getDefaultChartData();
        }
        const cats = chartData.categories.slice(); // flat array of strings
        if (cats.length === 0 || chartData.series.length === 0) return getDefaultChartData();

        const expected = cats.length;

        const pptxData = chartData.series
            .filter(s => s && Array.isArray(s.values))
            .map((s, idx) => {
                const cleanValues = s.values.map(v => (typeof v === 'number' && !isNaN(v)) ? v : 0);
                return {
                    name: s.name || `Series ${idx + 1}`,
                    labels: cats,          // <-- IMPORTANT: flat labels on EVERY series
                    values: cleanValues
                };
            });

        // Ensure all series lengths match
        if (!pptxData.length || !pptxData.every(s => s.values.length === expected)) {
            return getDefaultChartData();
        }
        return pptxData;
    } catch {
        return getDefaultChartData();
    }
}


// Helper function to get default/fallback chart data in correct format
function getDefaultChartData() {
    // Flat labels; not nested
    return [
        { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3'], values: [1, 2, 3] }
    ];
}


// Helper function to validate PPTX data structure
function validatePptxData(data) {
    if (!Array.isArray(data) || data.length === 0) return false;
    const lblCount = Array.isArray(data[0].labels) ? data[0].labels.length : 0;
    const valCount = Array.isArray(data[0].values) ? data[0].values.length : 0;
    if (!lblCount || !valCount) return false;

    for (let i = 0; i < data.length; i++) {
        const s = data[i];
        if (typeof s.name !== 'string') return false;
        if (!Array.isArray(s.labels) || s.labels.length !== lblCount) return false;
        if (!Array.isArray(s.values) || s.values.length !== valCount) return false;
        if (!s.values.every(v => typeof v === 'number' && !isNaN(v))) return false;
    }
    return true;
}

// FIXED: Simplified chart options to avoid compatibility issues
function createChartOptions(chartData, position, styling) {
    const options = {
        x: position.x, y: position.y, w: position.w, h: position.h,
        title: chartData.title || 'Chart', showTitle: true,
        showLegend: chartData.series && chartData.series.length > 1,
        barDir: 'col',
        barGrouping: 'clustered',
        barGapWidthPct: 150,
        barOverlapPct: -25,
        // leave gridlines/legend fonts as you had
    };

    // Colors (hex without '#')
    if (styling.colors?.length) {
        const hexes = styling.colors.map(c => normalizeColor(c, true)).filter(h => typeof h === 'string' && h.length === 6);
        if (hexes.length) options.chartColors = hexes;
    }

    return validateChartOptions(options);
}


// Helper function to validate chart options
function validateChartOptions(options) {
    const validated = { ...options };

    // Ensure required dimensions are valid numbers
    validated.x = (typeof options.x === 'number' && !isNaN(options.x)) ? Math.max(0, options.x) : 0.5;
    validated.y = (typeof options.y === 'number' && !isNaN(options.y)) ? Math.max(0, options.y) : 1;
    validated.w = (typeof options.w === 'number' && !isNaN(options.w)) ? Math.max(2, options.w) : 6;
    validated.h = (typeof options.h === 'number' && !isNaN(options.h)) ? Math.max(1.5, options.h) : 4;

    // Ensure title is a string
    validated.title = typeof options.title === 'string' ? options.title : 'Chart';

    // Ensure boolean values are actually boolean
    validated.showLegend = Boolean(options.showLegend);

    // Remove any undefined or null properties
    Object.keys(validated).forEach(key => {
        if (validated[key] === undefined || validated[key] === null) {
            delete validated[key];
        }
    });

    return validated;
}

function getChartType(chartType) {
    const typeMap = {
        'bar': 'bar',
        'column': 'bar', // pptxgenjs uses 'bar' for column charts
        'line': 'line',
        'pie': 'pie',
        'scatter': 'scatter',
        'area': 'area'
    };

    return typeMap[chartType] || 'bar';
}

// FIXED: Better color normalization with option for pptxgenjs format
function normalizeColor(color, forPptx = false) {
    if (!color) return forPptx ? '4472C4' : '#4472C4';

    // If already hex
    if (color.startsWith('#')) {
        const hex = color.length === 7 ? color.substring(1) : '4472C4';
        return forPptx ? hex.toUpperCase() : color;
    }

    // Convert rgb() to hex
    const rgbMatch = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (rgbMatch) {
        const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
        const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
        const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
        const hex = `${r}${g}${b}`.toUpperCase();
        return forPptx ? hex : `#${hex}`;
    }

    // Convert rgba() to hex (ignore alpha)
    const rgbaMatch = color.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*[\d.]+\)/);
    if (rgbaMatch) {
        const r = parseInt(rgbaMatch[1]).toString(16).padStart(2, '0');
        const g = parseInt(rgbaMatch[2]).toString(16).padStart(2, '0');
        const b = parseInt(rgbaMatch[3]).toString(16).padStart(2, '0');
        const hex = `${r}${g}${b}`.toUpperCase();
        return forPptx ? hex : `#${hex}`;
    }

    // Named colors to hex (basic set)
    const namedColors = {
        'red': forPptx ? 'FF0000' : '#FF0000',
        'green': forPptx ? '008000' : '#008000',
        'blue': forPptx ? '0000FF' : '#0000FF',
        'yellow': forPptx ? 'FFFF00' : '#FFFF00',
        'orange': forPptx ? 'FFA500' : '#FFA500',
        'purple': forPptx ? '800080' : '#800080',
        'pink': forPptx ? 'FFC0CB' : '#FFC0CB',
        'brown': forPptx ? 'A52A2A' : '#A52A2A',
        'black': forPptx ? '000000' : '#000000',
        'white': forPptx ? 'FFFFFF' : '#FFFFFF',
        'gray': forPptx ? '808080' : '#808080',
        'grey': forPptx ? '808080' : '#808080'
    };

    return namedColors[color.toLowerCase()] || (forPptx ? '4472C4' : '#4472C4');
}

async function processChartElement(pptx, pptSlide, chartElement, slideContext) {
    if (!isChartElement(chartElement)) {
        return false;
    }

    return await addChartToSlide(pptx, pptSlide, chartElement, slideContext);
}

// Export functions for use in the main conversion system
module.exports = {
    addChartToSlide,
    processChartElement,
    isChartElement
};