
function postProcessChartXML(chartXML, chartType, chartData) {
    if (chartType !== 'line') {
        return chartXML; // Only fix line charts
    }

    try {
        // Fix 1: Add c:grouping element if missing

        // Fix 1: Add c:grouping element if missing
        if (!chartXML.includes('<c:grouping')) {
            // Insert c:grouping right after <c:lineChart> opening tag
            chartXML = chartXML.replace(
                /(<c:lineChart[^>]*>)\s*/,
                '$1<c:grouping val="standard"/>'
            );
        }        // Fix 2: Remove third axis ID from c:lineChart if present
        // Line charts should only have 2 axis IDs
        const axIdMatches = chartXML.match(/<c:axId val="[^"]+"\s*\/>/g) || [];
        if (axIdMatches.length > 2) {
            // Find the c:lineChart section
            const lineChartMatch = chartXML.match(/<c:lineChart>([\s\S]*?)<\/c:lineChart>/);
            if (lineChartMatch) {
                let lineChartContent = lineChartMatch[1];
                const axIds = lineChartContent.match(/<c:axId val="[^"]+"\s*\/>/g) || [];

                if (axIds.length > 2) {
                    // Remove the third axis ID
                    const thirdAxisId = axIds[2];
                    lineChartContent = lineChartContent.replace(thirdAxisId, '');
                    chartXML = chartXML.replace(lineChartMatch[0], `<c:lineChart>${lineChartContent}</c:lineChart>`);
                }
            }
        }

        // Fix 3: Ensure only 2 axes are defined (catAx and valAx)
        // Remove any extra axis definitions
        const catAxCount = (chartXML.match(/<c:catAx>/g) || []).length;
        const valAxCount = (chartXML.match(/<c:valAx>/g) || []).length;

        if (valAxCount > 1) {
            // Remove duplicate valAx - keep only the first one
            const valAxMatches = chartXML.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g);
            if (valAxMatches && valAxMatches.length > 1) {
                // Remove the second valAx
                chartXML = chartXML.replace(valAxMatches[1], '');
            }
        }

        // Fix 4: Inject per-point marker customizations (c:dPt elements)
        if (chartData && chartData.series) {
            chartData.series.forEach((series, seriesIdx) => {
                if (series.markers && Array.isArray(series.markers)) {
                    // Find the c:ser element for this series
                    // AFTER:
                    const serPattern = new RegExp(`<c:ser>\\s*<c:idx val="${seriesIdx}"[\\s\\S]*?<\\/c:ser>`, 'g');

                    chartXML = chartXML.replace(serPattern, (serMatch) => {
                        // Build c:dPt elements for non-default markers
                        let dPtElements = '';

                        series.markers.forEach((marker, pointIdx) => {
                            if (marker.symbol !== 'circle' || marker.size !== 5) {
                                dPtElements += `<c:dPt><c:idx val="${pointIdx}"/><c:marker><c:symbol val="${marker.symbol}"/><c:size val="${marker.size}"/></c:marker></c:dPt>`;
                            }
                        });

                        if (dPtElements) {
                            // Insert c:dPt elements immediately after c:marker
                            serMatch = serMatch.replace(
                                /(<\/c:marker>)([\s\S]*?)(<c:cat>)/,
                                `$1${dPtElements}$2$3`
                            );
                        }

                        return serMatch;
                    });
                }
            });
        }

        // Add this new section to postProcessChartXML function, AFTER the existing Fix 4

        // Fix 5: Inject marker COLORS and SYMBOLS into series-level c:marker and c:dPt elements
        if (chartData && chartData.series) {
            chartData.series.forEach((series, seriesIdx) => {
                const markerColor = (series.markerColor || series.lineColor).replace('#', '').toUpperCase();
                const lineColor = (series.lineColor || '#4472C4').replace('#', '').toUpperCase();

                // Find this series in the XML
                const serPattern = new RegExp(
                    `<c:ser>\\s*<c:idx val="${seriesIdx}"[\\s\\S]*?<\\/c:ser>`,
                    'g'
                );

                chartXML = chartXML.replace(serPattern, (serMatch) => {
                    // Step 1: Inject marker color/symbol into SERIES-LEVEL c:marker
                    serMatch = serMatch.replace(
                        /(<c:marker>)([\s\S]*?)(<\/c:marker>)/,
                        (markerMatch, openTag, markerContent, closeTag) => {
                            // Check if spPr already exists in marker
                            if (!markerContent.includes('<c:spPr>')) {
                                // Add complete spPr with fill and border
                                const spPrXML = `<c:spPr><a:solidFill><a:srgbClr val="${markerColor}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln><a:effectLst/></c:spPr>`;
                                return `${openTag}${markerContent}${spPrXML}${closeTag}`;
                            }
                            return markerMatch;
                        }
                    );

                    // Step 2: Inject marker color into ALL c:dPt elements
                    const dPtPattern = /<c:dPt>(<c:idx val="\d+"\/><c:marker>.*?<\/c:marker>)<\/c:dPt>/g;

                    serMatch = serMatch.replace(dPtPattern, (dPtMatch, innerContent) => {
                        // Check if spPr already exists in this dPt
                        if (!dPtMatch.includes('<c:spPr>')) {
                            // Add spPr with marker fill color and line border
                            return `<c:dPt>${innerContent}<c:spPr><a:solidFill><a:srgbClr val="${markerColor}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>`;
                        }
                        return dPtMatch;
                    });

                    return serMatch;
                });
            });
        }


        return chartXML;
    } catch (error) {
        console.error('Error post-processing chart XML:', error);
        return chartXML; // Return original if processing fails
    }
}

/**
 * addChartToSlide.js
 * Converts HTML charts to PPTX charts using pptxgenjs
 * Node.js/JSDOM compatible version - CORRECTED
 */

function createChartOptions(chartData, position, styling) {
    const options = {
        x: position.x,
        y: position.y,
        w: position.w,
        h: position.h,
        showTitle: !!chartData.title,
        title: chartData.title || '',
        showLegend: true,
        legendPos: 'r',
        showValue: false,

        // ===== ADDED: Proper axis configuration for line charts =====
        catAxisOptions: {
            showTitle: false,
            showGridlines: chartData.type === 'line', // Show gridlines for line charts
            gridlineColor: '888888',
            gridlineSize: 1,
        },
        valAxisOptions: {
            showTitle: false,
            showGridlines: true,
            gridlineColor: '888888',
            gridlineSize: 1,
        }
        // ===== END ADDITION =====
    };

    // Apply series-specific styling
    if (styling.seriesOptions && chartData.series) {
        options.chartColors = chartData.series.map(s => s.lineColor || '#4472C4');
    }

    // ===== ADDED: Line chart specific options =====

    // ===== END ADDITION =====

    return options;
}

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

        // ===== CHANGED: Add special handling for line charts =====
        if (chartData.type === 'line') {
            // Ensure only 2 axes are configured
            if (!chartOptions.catAxisOptions) {
                chartOptions.catAxisOptions = {};
            }
            if (!chartOptions.valAxisOptions) {
                chartOptions.valAxisOptions = {};
            }

            // Remove any third axis configuration
            delete chartOptions.serAxisOptions;
            delete chartOptions.catAxisOptions2;
            delete chartOptions.valAxisOptions2;
        }
        // ===== END CHANGE =====

        // Add chart to slide

        // Collect marker data so postProcessPPTXForMarkers can inject c:dPt elements
        if (chartData.type === 'line' && chartData.series?.length) {
            if (!pptx._chartMarkerData) pptx._chartMarkerData = [];
            pptx._chartMarkerData.push({
                chartIndex: pptx._chartMarkerData.length,
                series: chartData.series.map(s => ({
                    markers: s.markers || [],
                    markerColor: s.markerColor,
                    lineColor: s.lineColor
                }))
            });
        }
        pptSlide.addChart(getChartType(chartData.type), pptxChartData, chartOptions);


       
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

    // Check data-chart-type attribute (set by LineChartHandler / BarChartHandler)
    if (element.getAttribute && element.getAttribute('data-chart-type')) {
        return true;
    }

    // Check for chart-specific elements inside
    const hasChartElements =
        element.querySelector('.bar') ||
        element.querySelector('.line-series') ||
        element.querySelector('.line-point') ||
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
            isValid: true
        };
    } catch (error) {
        console.error("Error extracting chart data:", error);
        return {
            isValid: false
        };
    }
}

function detectChartType(container) {
    // Check data-chart-type attribute first
    const dataType = container.getAttribute('data-chart-type');
    if (dataType) {
        return dataType;
    }

    // Check for line-specific elements
    if (container.querySelector('.line-series') ||
        container.querySelector('.line-point') ||
        container.classList.contains('line-chart')) {
        return 'line';
    }

    // Check for bar-specific elements
    if (container.querySelector('.bar') ||
        container.classList.contains('bar-chart') ||
        container.classList.contains('column-chart')) {
        return 'bar';
    }

    // Check for pie-specific elements
    if (container.querySelector('.pie-slice') ||
        container.classList.contains('pie-chart')) {
        return 'pie';
    }

    // Default to bar chart
    return 'bar';
}

function extractBarChartData(chartContainer) {
    try {
        const categories = [];
        const series = [];

        // Try to get categories from .category-label elements
        const categoryLabels = chartContainer.querySelectorAll('.category-label');
        categoryLabels.forEach(label => {
            categories.push(label.textContent.trim());
        });

        // Try to get series data from .bar elements
        const bars = chartContainer.querySelectorAll('.bar');
        const seriesMap = new Map();

        bars.forEach((bar, idx) => {
            const seriesName = bar.getAttribute('data-series-name') || `Series ${idx + 1}`;
            const value = parseFloat(bar.getAttribute('data-value') || '0');
            const color = bar.getAttribute('data-color') || '#4472C4';

            if (!seriesMap.has(seriesName)) {
                seriesMap.set(seriesName, {
                    name: seriesName,
                    values: [],
                    color: color
                });
            }

            seriesMap.get(seriesName).values.push(value);
        });

        seriesMap.forEach(s => series.push(s));

        // If no data found, try generic approach
        if (series.length === 0) {
            const dataElements = chartContainer.querySelectorAll('[data-value]');
            if (dataElements.length > 0) {
                const values = Array.from(dataElements).map(el =>
                    parseFloat(el.getAttribute('data-value') || '0')
                );
                series.push({
                    name: 'Series 1',
                    values: values,
                    color: '#4472C4'
                });
            }
        }

        // Ensure we have categories
        if (categories.length === 0 && series.length > 0) {
            const valueCount = series[0].values.length;
            for (let i = 0; i < valueCount; i++) {
                categories.push(`Category ${i + 1}`);
            }
        }

        const allValues = series.flatMap(s => s.values);
        return {
            categories,
            series,
            maxValue: allValues.length ? Math.max(...allValues) : 10,
            minValue: allValues.length ? Math.min(...allValues, 0) : 0
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

function extractLineChartData(chartContainer) {
    try {
        // ── Try to read the structured HTML emitted by LineChartHandler ──────
        const seriesNodes = chartContainer.querySelectorAll('.line-series');

        if (seriesNodes.length > 0) {
            const categories = [];
            const series = [];

            // ── Collect categories from the first series ──────────────────
            const firstSeriesPoints = seriesNodes[0].querySelectorAll('.line-point');
            firstSeriesPoints.forEach(pt => {
                const label = pt.getAttribute('data-category-label') || '';
                categories.push(label);
            });

            // ── Also collect from .category-labels span elements ──────────
            if (categories.length === 0) {
                const labelEls = chartContainer.querySelectorAll('.category-label');
                labelEls.forEach(el => categories.push(el.textContent.trim()));
            }

            // ── Parse each series ─────────────────────────────────────────
            seriesNodes.forEach((serEl, sIdx) => {
                const name = serEl.getAttribute('data-series-name') || `Series ${sIdx + 1}`;
                const lineColor = serEl.getAttribute('data-series-line-color') || '#4472C4';
                const markerColor = serEl.getAttribute('data-series-marker-color') || lineColor;

                const points = serEl.querySelectorAll('.line-point');
                const values = [];
                const markers = [];  // ADDED - Array to store marker data

                // Build values indexed by category position
                const byIdx = new Map();
                const markerByIdx = new Map();  // ADDED - Map to store markers by index
                points.forEach(pt => {
                    const catIdx = parseInt(pt.getAttribute('data-category-index') ?? '0', 10);
                    const val = parseFloat(pt.getAttribute('data-value') ?? '0');
                    byIdx.set(catIdx, isNaN(val) ? 0 : val);

                    // ADDED - Extract marker data
                    const markerSymbol = pt.getAttribute('data-marker-symbol') || 'circle';
                    const markerSize = parseInt(pt.getAttribute('data-marker-size') || '5', 10);
                    markerByIdx.set(catIdx, { symbol: markerSymbol, size: markerSize });
                });

                // Ensure we have one value per category
                const count = Math.max(categories.length, byIdx.size);
                for (let i = 0; i < count; i++) {
                    values.push(byIdx.get(i) ?? 0);
                    markers.push(markerByIdx.get(i) ?? { symbol: 'circle', size: 5 });  // ADDED
                }

                series.push({
                    name,
                    lineColor: normalizeColor(lineColor),
                    markerColor: normalizeColor(markerColor),
                    values,
                    markers
                });  // ADDED markers
            });

            // Fill categories if still empty
            if (categories.length === 0 && series.length > 0) {
                const count = series[0].values.length;
                for (let i = 0; i < count; i++) categories.push(`Category ${i + 1}`);
            }

            const allVals = series.flatMap(s => s.values).filter(v => !isNaN(v));
            return {
                categories,
                series,
                maxValue: allVals.length ? Math.max(...allVals) : 10,
                minValue: allVals.length ? Math.min(...allVals, 0) : 0,
            };
        }

        // ── Fallback: generic [data-value] elements (same as bar chart) ──
        return extractBarChartData(chartContainer);

    } catch (error) {
        console.error("Error in extractLineChartData:", error);
        return {
            categories: ['Category 1'],
            series: [{ name: 'Series 1', values: [1], color: '#4472C4' }],
            maxValue: 1,
            minValue: 0,
        };
    }
}

function extractPieChartData(chartContainer) {
    try {
        const categories = [];
        const values = [];
        const colors = [];

        // Try to get pie slices
        const slices = chartContainer.querySelectorAll('.pie-slice, [data-slice]');
        slices.forEach(slice => {
            const label = slice.getAttribute('data-label') ||
                slice.getAttribute('data-category') ||
                `Slice ${categories.length + 1}`;
            const value = parseFloat(slice.getAttribute('data-value') || '0');
            const color = slice.getAttribute('data-color') ||
                getComputedStyle(slice).fill || '#4472C4';

            categories.push(label);
            values.push(value);
            colors.push(normalizeColor(color));
        });

        // If no slices found, try generic approach
        if (categories.length === 0) {
            const dataElements = chartContainer.querySelectorAll('[data-value]');
            dataElements.forEach((el, idx) => {
                categories.push(`Category ${idx + 1}`);
                values.push(parseFloat(el.getAttribute('data-value') || '0'));
                colors.push('#4472C4');
            });
        }

        return {
            categories,
            series: [{
                name: 'Values',
                values: values,
                colors: colors
            }],
            maxValue: values.length ? Math.max(...values) : 100,
            minValue: 0
        };
    } catch (error) {
        console.error("Error in extractPieChartData:", error);
        return {
            categories: ['Category 1'],
            series: [{ name: 'Values', values: [100], colors: ['#4472C4'] }],
            maxValue: 100,
            minValue: 0
        };
    }
}

function extractChartPosition(element, slideContext) {
    const style = element.style;
    const left = parseFloat(style.left) || 0;
    const top = parseFloat(style.top) || 0;
    let width = parseFloat(style.width) || 400;
    let height = parseFloat(style.height) || 300;
    
    // Get chart type
    const chartType = element.getAttribute('data-chart-type');
    
    // Increase dimensions for line charts to prevent marker collision
    if (chartType === 'line') {
        width = width * 1.45;   // 15% increase in width
        height = height * 1.25;  // 25% increase in height
    }
    
    // Convert px to inches (assuming 96 DPI)
    return {
        x: left / 96,
        y: top / 96,
        w: width / 96,
        h: height / 96
    };
}

function extractChartStyling(element, chartData) {
    const styling = {
        colors: [],
        seriesOptions: {}
    };

    try {
        // Extract colors from series data
        if (chartData.series) {
            styling.colors = chartData.series.map(s => s.lineColor || '#4472C4');
            styling.markerColors = chartData.series.map(s => s.markerColor || s.lineColor || '#4472C4');
        }

        // Extract additional styling from element attributes
        const bgColor = element.getAttribute('data-bg-color');
        if (bgColor) {
            styling.backgroundColor = normalizeColor(bgColor);
        }

        const borderColor = element.getAttribute('data-border-color');
        if (borderColor) {
            styling.borderColor = normalizeColor(borderColor);
        }

    } catch (error) {
        console.error("Error extracting chart styling:", error);
    }

    return styling;
}

function convertToPptxFormat(chartData) {
    try {
        // ────────────────────────────────────────────────────────────────────────
        // CRITICAL FIX: Flat labels array (NOT nested under "labels" key)
        // ────────────────────────────────────────────────────────────────────────
        const cats = chartData.categories || [];
        if (cats.length === 0 || chartData.series.length === 0) return getDefaultChartData();

        const expected = cats.length;

        const pptxData = chartData.series
            .filter(s => s && Array.isArray(s.values))
            .map((s, idx) => {
                const cleanValues = s.values.map(v => (typeof v === 'number' && !isNaN(v)) ? v : 0);

                // ADDED - Build dataPointStyles for markers
                // ADDED - Build dataPointStyles for markers
                const dataPointStyles = [];
                if (s.markers && Array.isArray(s.markers)) {

                }

                const seriesData = {
                    name: s.name || `Series ${idx + 1}`,
                    labels: cats,          // <-- IMPORTANT: flat labels on EVERY series
                    values: cleanValues
                };

                // Add dataPointStyles only if we have custom markers
                if (dataPointStyles.length > 0) {
                    seriesData.dataPointStyles = dataPointStyles;
                }

                return seriesData;
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

// FIXED: Chart options that produce Syncfusion-compatible XML
function createChartOptions(chartData, position, styling) {
    const isLine = chartData.type === 'line';

    const options = {
        // Position & size
        x: position.x, y: position.y, w: position.w, h: position.h,

        // Title
        title: chartData.title || 'Chart',
        showTitle: !!chartData.title,

        // Legend — bottom position to match original XML (legendPos="b")
        showLegend: chartData.series && chartData.series.length > 1,
        legendPos: 'b',

        // Blank cells — use "gap" (original uses "gap", NOT "span")
        plotAreaBorder: { pt: 0 },

        // Axis tick marks & label position matching original XML
        catAxisLabelPos: 'nextTo',   // tickLblPos="nextTo"
        valAxisLabelPos: 'nextTo',
        catAxisMajorTickMark: 'none',   // majorTickMark="none"
        catAxisMinorTickMark: 'none',
        valAxisMajorTickMark: 'none',
        valAxisMinorTickMark: 'none',

        // Grid lines — only valAx gets majorGridlines (same as original)
        catGridLine: { style: 'none' },
        valGridLine: { style: 'solid', color: 'D9D9D9', pt: 0.75 },

        // Rounded corners OFF (original has roundedCorners val="0")
        chartArea: { roundedCorners: false },

        // Display blanks as gap (matches original dispBlanksAs val="gap")
        displayBlanksAs: 'gap',
    };

    if (isLine) {
        // ── Line chart specific ──────────────────────────────────────────
        // grouping="standard" — this is the key missing tag in converted XML
        options.lineGrouping = 'standard';
        options.lineDataSymbol = 'circle';
        options.lineDataSymbolSize = 5;
        options.lineSize = 2;        // ~28575 EMU in original
        options.smooth = false;

        // Only 2 axes for line chart (original has exactly 2 axId entries)
        options.catAxisLineShow = true;
        options.valAxisLineShow = false;   // valAx line is noFill in original

    } else {
        // ── Bar / Column chart specific ──────────────────────────────────
        options.barDir = 'col';
        options.barGrouping = 'clustered';
        options.barGapWidthPct = 150;
        options.barOverlapPct = -25;
    }

    // Colors (hex without '#') — one per series
    if (styling.colors?.length) {
        const hexes = styling.colors
            .map(c => normalizeColor(c, true))
            .filter(h => typeof h === 'string' && h.length === 6);
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

async function postProcessPPTXForMarkers(pptxPath, markerDataArray) {
    try {
        const JSZip = require('jszip');
        const fs = require('fs').promises;

        const pptxBuffer = await fs.readFile(pptxPath);
        const zip = await JSZip.loadAsync(pptxBuffer);

        for (const markerData of markerDataArray) {
            const chartPath = `ppt/charts/chart${markerData.chartIndex + 1}.xml`;
            const chartFile = zip.file(chartPath);

            if (!chartFile) continue;

            let chartXML = await chartFile.async('string');

            // Process each series
            markerData.series.forEach((series, seriesIdx) => {
                const markerColor = (series.markerColor || series.lineColor || '#4472C4').replace('#', '').toUpperCase();
                const lineColor = (series.lineColor || '#4472C4').replace('#', '').toUpperCase();
                
                // Find this series in the XML
                const serPattern = new RegExp(
                    `(<c:ser>\\s*<c:idx val="${seriesIdx}"\\s*\\/>[\\s\\S]*?)(<c:marker>)([\\s\\S]*?)(<\\/c:marker>)`,
                    'g'
                );
                
                chartXML = chartXML.replace(serPattern, (match, beforeMarker, openMarker, markerContent, closeMarker) => {
                    // Check if spPr exists in the marker
                    if (markerContent.includes('<c:spPr>')) {
                        // Replace existing spPr colors
                        markerContent = markerContent.replace(
                            /<a:srgbClr val="[A-F0-9]{6}"\s*\/>/g,
                            (colorMatch, offset) => {
                                // First occurrence is fill, second is border
                                if (markerContent.indexOf(colorMatch) === offset) {
                                    return `<a:srgbClr val="${markerColor}"/>`;
                                }
                                return `<a:srgbClr val="${lineColor}"/>`;
                            }
                        );
                    } else {
                        // Add spPr with colors
                        markerContent += `<c:spPr><a:solidFill><a:srgbClr val="${markerColor}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln><a:effectLst/></c:spPr>`;
                    }
                    
                    return beforeMarker + openMarker + markerContent + closeMarker;
                });

                // Also inject c:dPt elements for per-point markers if needed
                if (series.markers && series.markers.length > 0) {
                    const serPattern2 = new RegExp(`<c:ser>\\s*<c:idx val="${seriesIdx}"[\\s\\S]*?<\\/c:ser>`, 'g');
                    chartXML = chartXML.replace(serPattern2, (serMatch) => {
                        let dPtElements = '';

                        series.markers.forEach((marker, pointIdx) => {
                            if (marker.symbol !== 'circle' || marker.size !== 5) {
                                dPtElements += `<c:dPt><c:idx val="${pointIdx}"/><c:marker><c:symbol val="${marker.symbol}"/><c:size val="${marker.size}"/><c:spPr><a:solidFill><a:srgbClr val="${markerColor}"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:marker></c:dPt>`;
                            }
                        });

                        if (dPtElements) {
                            serMatch = serMatch.replace(
                                /(<\/c:marker>)([\s\S]*?)(<c:cat>)/,
                                `$1${dPtElements}$2$3`
                            );
                        }

                        return serMatch;
                    });
                }
            });

            zip.file(chartPath, chartXML);
        }

        const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
        await fs.writeFile(pptxPath, modifiedBuffer);

       

    } catch (error) {
        console.error('   ⚠️ Error post-processing chart markers:', error);
    }
}
// Export functions for use in the main conversion system
module.exports = {
    addChartToSlide,
    processChartElement,
    isChartElement,
    postProcessChartXML,
    postProcessPPTXForMarkers
};