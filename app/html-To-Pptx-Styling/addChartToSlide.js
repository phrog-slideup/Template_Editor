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
 * Node.js/JSDOM compatible version - FIXED FOR BOTH BAR AND LINE CHARTS
 */
function createChartOptions(chartData, position, styling) {
    const baseOptions = {
        x: position.x,
        y: position.y,
        w: position.w,
        h: position.h,
        showTitle: !!chartData.title,
        title: chartData.title || '',
        showLegend: true,
        legendPos: 'r',
    };

    if (chartData.type === 'line') {
        return {
            ...baseOptions,
            catAxisOptions: {
                showTitle: false,
                showGridlines: true,
                gridlineColor: '888888',
                gridlineSize: 1,
            },
            valAxisOptions: {
                showTitle: false,
                showGridlines: true,
                gridlineColor: '888888',
                gridlineSize: 1,
            }
        };
    } else if (chartData.type === 'bar' || chartData.type === 'column') {
    const isHorizontal = chartData.isHorizontal === true;
    const axisMax = Math.ceil(chartData.maxValue * 10) / 10 + 1;

    // For horizontal bar in pptxgenjs:
    //   catAxisTitle = the LEFT axis (Y) = where categories are = XML's catAx = "NR Axis Title"
    //   valAxisTitle = the BOTTOM axis (X) = where values are  = XML's valAx = ""
    // So pass catAxisTitle directly to catAxisTitle — NO swap needed
    return {
        ...baseOptions,
        barDir: isHorizontal ? 'bar' : 'col',
        barGrouping: 'clustered',
        barGapWidthPct: 150,
        barOverlapPct: -25,
        legendPos: 'b',
        valAxisMinVal: 0,
        valAxisMaxVal: axisMax,
        showValue: true,
        dataLabelPosition: 'outEnd',
        dataLabelFontSize: 11,
        dataLabelColor: '666666',
        dataLabelFormatCode: 'General',       // ✅ FIX: no trailing dots
        showTitle: !!chartData.title,         // ✅ FIX: ensure chart title shows
        title: chartData.title || '',
        // ✅ FIX: catAxisTitle in pptxgenjs = left/Y axis for horizontal = XML catAx
        showCatAxisTitle: !!chartData.catAxisTitle,
        catAxisTitle: chartData.catAxisTitle || '',
        showValAxisTitle: !!chartData.valAxisTitle,
        valAxisTitle: chartData.valAxisTitle || '',
    };
} else {
        return baseOptions;
    }
}

function addChartToSlide(pptx, pptSlide, element, slideContext) {
    try {
        // Check if this is a chart container
        if (!isChartElement(element)) {
            return false;
        }

        // Extract chart type
        // Extract chart type
        const chartType = element.getAttribute('data-chart-type') || 'bar';

        // Extract chart data based on type
        let chartData;
        if (chartType === 'line') {
            chartData = extractLineChartData(element);
        } else if (chartType === 'bar' || chartType === 'column') {
            chartData = extractBarChartData(element);
        } else if (chartType === 'pie' || chartType === 'doughnut') {
            chartData = extractPieChartData(element);
        } else {
            chartData = extractBarChartData(element); // Default fallback
        }

        // Add type to chartData for options generation
        chartData.type = chartType;

        // Detect horizontal bar chart from data-chart-direction attribute
        // (set this attribute in barChartHandler.js when isHorizontal=true)
        chartData.isHorizontal = element.getAttribute('data-chart-direction') === 'horizontal';

        // Extract position and styling
        const position = extractChartPosition(element, slideContext);
        const styling = extractChartStyling(element, chartData);

        // Convert data to pptxgenjs format
        const pptxData = convertToPptxFormat(chartData);

        // Create chart options (now with type-specific logic)
        const options = createChartOptions(chartData, position, styling);

        // Apply series-specific colors if available
        if (styling.colors && styling.colors.length > 0) {
            options.chartColors = styling.colors.map(c => normalizeColor(c, true));
        } else if (chartData.series && chartData.series.length > 0) {
            options.chartColors = chartData.series.map(s =>
                normalizeColor(s.lineColor || s.color || '#4472C4', true)
            );
        }

        // Validate options before adding chart
        const validatedOptions = validateChartOptions(options);

        // Add the chart to the slide
        const pptxChartType = getChartType(chartType);

        // ⭐ CRITICAL: Different handling for line vs bar charts
        if (chartType === 'line') {
            // Store the original chartData for post-processing
            pptSlide.addChart(pptxChartType, pptxData, validatedOptions);

            // ⭐ Attach metadata for post-processing
            if (!pptx._chartMarkerData) {
                pptx._chartMarkerData = [];
            }

            // Extract marker data from chartData
            const markerData = {
                chartIndex: (pptx._chartCounter || 0),
                series: chartData.series.map(s => ({
                    lineColor: s.lineColor || '#4472C4',
                    markerColor: s.markerColor || s.lineColor || '#4472C4',
                    markers: s.markers || []
                }))
            };

            pptx._chartMarkerData.push(markerData);
            pptx._chartCounter = (pptx._chartCounter || 0) + 1;

        } else {
            // Bar charts and others - standard pptxgenjs behavior
            pptSlide.addChart(pptxChartType, pptxData, validatedOptions);
        }

        return true;

    } catch (error) {
        console.error('Error adding chart to slide:', error);
        return false;
    }
}

function isChartElement(element) {
    try {
        if (!element || !element.classList) return false;
        return element.classList.contains('chart-container') ||
            element.classList.contains('bar-chart') ||
            element.classList.contains('line-chart') ||
            element.classList.contains('pie-chart');
    } catch (error) {
        return false;
    }
}

// ✅ FIXED: Extract line chart data from HTML with proper data-* attributes
function extractLineChartData(chartContainer) {
    const categories = [];
    const series = [];

    try {
        // Extract category labels
        const categoryElements = chartContainer.querySelectorAll('.category-label');
        categoryElements.forEach(el => {
            const text = el.textContent?.trim() || '';
            if (text) categories.push(text);
        });

        // Extract series
        const seriesElements = chartContainer.querySelectorAll('.line-series');
        seriesElements.forEach((seriesEl, idx) => {
            const name = seriesEl.getAttribute('data-series-name') || `Series ${idx + 1}`;
            const lineColor = seriesEl.getAttribute('data-series-line-color') || '#4472C4';
            const markerColor = seriesEl.getAttribute('data-series-marker-color') || lineColor;

            const values = [];
            const markers = [];

            // Extract points
            const pointElements = seriesEl.querySelectorAll('.line-point');
            pointElements.forEach(pointEl => {
                const value = parseFloat(pointEl.getAttribute('data-value')) || 0;
                const markerSymbol = pointEl.getAttribute('data-marker-symbol') || 'circle';
                const markerSize = parseInt(pointEl.getAttribute('data-marker-size')) || 5;

                values.push(value);
                markers.push({ symbol: markerSymbol, size: markerSize });
            });

            series.push({
                name,
                lineColor,
                markerColor,
                values,
                markers
            });
        });

        // Fallback if no data found
        if (categories.length === 0) {
            categories.push('Category 1', 'Category 2', 'Category 3');
        }
        if (series.length === 0) {
            series.push({
                name: 'Series 1',
                lineColor: '#4472C4',
                markerColor: '#4472C4',
                values: [1, 2, 3],
                markers: [
                    { symbol: 'circle', size: 5 },
                    { symbol: 'circle', size: 5 },
                    { symbol: 'circle', size: 5 }
                ]
            });
        }

        return { categories, series };

    } catch (error) {
        console.error("Error in extractLineChartData:", error);
        return {
            categories: ['Category 1'],
            series: [{
                name: 'Series 1',
                lineColor: '#4472C4',
                markerColor: '#4472C4',
                values: [1],
                markers: [{ symbol: 'circle', size: 5 }]
            }]
        };
    }
}

function extractBarChartData(chartContainer) {

     const titleEl = chartContainer.querySelector('.chart-title');
    const chartTitle = titleEl ? titleEl.textContent.trim() : '';
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

        // Horizontal bar charts render categories bottom-up in the DOM
        // (Cat4 appears first in DOM, Cat1 last). Reverse to restore natural order
        // so PPTX receives Cat1, Cat2, Cat3, Cat4.


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

        // Detect if this is a horizontal bar chart
        // In horizontal mode, HTML renders categories bottom-up (Cat4 at top DOM position)
        // but data-category indices still go 0,1,2,3 = Cat1,Cat2,Cat3,Cat4
        // PPT expects categories in natural order (Cat1 first), so no reversal needed for values.
        // However, series in HTML are rendered reversed (S3 on top, S1 on bottom),
        // so data-series indices 0,1,2 = S1,S2,S3 — this is already correct natural order.
        // No reversal needed: pptxgenjs receives S1,S2,S3 and renders them correctly.

        sortedSeriesIndices.forEach(seriesIndex => {
            const categoryMap = seriesMap.get(seriesIndex);
            const values = [];

            // Fill values in natural category order (Cat1=0, Cat2=1, ...)
            for (let catIndex = 0; catIndex < categories.length; catIndex++) {
                values.push(categoryMap.get(catIndex) !== undefined ? categoryMap.get(catIndex) : 0);
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

        const catAxisTitle = chartContainer.getAttribute('data-cat-axis-title') || '';
        const valAxisTitle = chartContainer.getAttribute('data-val-axis-title') || '';

        return {
            categories,
            series,
            maxValue,
            minValue,
            catAxisTitle,
            valAxisTitle,
            title: chartTitle
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

function extractPieChartData(chartContainer) {
    // Fallback to bar chart extraction for now
    return extractBarChartData(chartContainer);
}

function extractChartPosition(chartContainer, slideContext) {
    try {
        let left = 0, top = 0, width = 600, height = 400;

        const containerStyle = chartContainer.getAttribute('style');
        if (containerStyle) {
            const leftMatch = containerStyle.match(/left:\s*([0-9]*\.?[0-9]+)px/i);
            const topMatch = containerStyle.match(/top:\s*([0-9]*\.?[0-9]+)px/i);
            const widthMatch = containerStyle.match(/width:\s*([0-9]*\.?[0-9]+)px/i);
            const heightMatch = containerStyle.match(/height:\s*([0-9]*\.?[0-9]+)px/i);

            if (leftMatch) left = parseFloat(leftMatch[1]);
            if (topMatch) top = parseFloat(topMatch[1]);
            if (widthMatch) width = parseFloat(widthMatch[1]);
            if (heightMatch) height = parseFloat(heightMatch[1]);
        }

        // ✅ FIX: Try multiple selectors to find the actual slide container
        let slideWidthPx = 0;
        let slideHeightPx = 0;

        const slideSelectors = ['.sli-slide', '.slide', '.slide-container', '[class*="slide"]'];
        for (const selector of slideSelectors) {
            const el = chartContainer.closest(selector);
            if (el) {
                const s = el.getAttribute('style') || '';
                const sw = s.match(/width:\s*([0-9]*\.?[0-9]+)px/i);
                const sh = s.match(/height:\s*([0-9]*\.?[0-9]+)px/i);
                if (sw && sh) {
                    slideWidthPx = parseFloat(sw[1]);
                    slideHeightPx = parseFloat(sh[1]);
                    break;
                }
            }
        }

        // ✅ FIX: If no slide container found, infer from chart position + size
        // The chart at left:48, width:858 implies slide width ≈ 48+858+~54 = ~960
        if (!slideWidthPx || !slideHeightPx) {
            // Use slideContext if passed, otherwise infer from chart bounds
            if (slideContext && slideContext.slideWidth && slideContext.slideHeight) {
                slideWidthPx = slideContext.slideWidth;
                slideHeightPx = slideContext.slideHeight;
            } else {
                // Infer: chart occupies most of slide, so slide ≈ chart + margins
                // left:48 + width:858 + ~54 right margin = 960
                slideWidthPx = Math.round((left + width) / 0.9); // assume chart is ~90% of slide
                slideHeightPx = Math.round((top + height) / 0.9);
            }
        }

        const pptW_in = 13.33;
        const pptH_in = 7.5;

        const pptX = (left / slideWidthPx) * pptW_in;
        const pptY = (top / slideHeightPx) * pptH_in;
        const pptW = (width / slideWidthPx) * pptW_in;
        const pptH = (height / slideHeightPx) * pptH_in;

        return {
            x: Math.max(0, parseFloat(pptX.toFixed(3))),
            y: Math.max(0, parseFloat(pptY.toFixed(3))),
            w: parseFloat(pptW.toFixed(3)),
            h: parseFloat(pptH.toFixed(3))
        };

    } catch (error) {
        console.error("Error extracting position:", error);
        return { x: 0.5, y: 0.5, w: 12, h: 6.5 };
    }
}

function extractChartStyling(chartContainer, chartData) {
    const styling = {
        colors: [],
        fontSize: 12,
        fontFamily: 'Arial',
        showLegend: true,
        showGridLines: true,
        seriesOptions: null
    };

    try {
        // Extract colors from series data
        if (chartData && chartData.series) {
            styling.colors = chartData.series.map(s => s.color || s.lineColor || '#4472C4');
            styling.seriesOptions = chartData.series;
        }

        // Check for grid lines - if axis labels exist, assume grid lines are wanted
        const axisLabels = chartContainer.querySelectorAll('.axis-label');
        styling.showGridLines = axisLabels.length > 0;

    } catch (error) {
        console.error("Error extracting styling:", error);
    }

    return styling;
}

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