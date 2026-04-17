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

// ✅ FIXED: Chart-type-specific options
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
        showValue: false,
    };

    // ✅ Apply chart-type-specific options
    if (chartData.type === 'line') {
        // Line chart specific options
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
        // Bar/Column chart specific options
        return {
            ...baseOptions,
            barDir: 'col',
            barGrouping: 'clustered',
            barGapWidthPct: 150,
            barOverlapPct: -25,
        };
    } else if (chartData.type === 'doughnut') {
        return {
            ...baseOptions,
            legendPos: chartData.legendPos || 'b',
            holeSize: typeof chartData.holeSize === 'number' ? chartData.holeSize : 50,
            firstSliceAng: typeof chartData.firstSliceAng === 'number' ? chartData.firstSliceAng : 0,
            showLegend: true,
            showValue: Boolean(chartData.dataLabels?.showValue),
            showPercent: Boolean(chartData.dataLabels?.showPercent),
            showLeaderLines: Boolean(chartData.dataLabels?.showLeaderLines),
            dataLabelFormatCode: chartData.dataLabels?.formatCode || 'General',
            dataLabelColor: normalizeColor(chartData.dataLabels?.color || '#4D4D4D', true)
        };
    } else if (chartData.type === 'pie') {
        return {
            ...baseOptions,
            legendPos: chartData.legendPos || 'b',
            firstSliceAng: typeof chartData.firstSliceAng === 'number' ? chartData.firstSliceAng : 0,
            showLegend: true,
            showValue: Boolean(chartData.dataLabels?.showValue),
            showPercent: Boolean(chartData.dataLabels?.showPercent),
            showLeaderLines: Boolean(chartData.dataLabels?.showLeaderLines),
            dataLabelFormatCode: chartData.dataLabels?.formatCode || 'General',
            dataLabelColor: normalizeColor(chartData.dataLabels?.color || '#4D4D4D', true)
        };
    } else {
        // Default options for other chart types
        return baseOptions;
    }
}

async function addChartToSlide(pptx, pptSlide, element, slideContext) {
    try {
        // Check if this is a chart container
        if (!isChartElement(element)) {
            return false;
        }

        // Extract chart type
        const chartType = element.getAttribute('data-chart-type') || 'bar';

        if (shouldRenderPieChartAsSvg(element, chartType)) {
            return await addStyledPieChartImage(pptSlide, element, slideContext);
        }

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
        const chartIndex = (pptx._customChartCounter || 0);

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
                type: 'lineMarkers',
                chartIndex,
                series: chartData.series.map(s => ({
                    lineColor: s.lineColor || '#4472C4',
                    markerColor: s.markerColor || s.lineColor || '#4472C4',
                    markers: s.markers || []
                }))
            };

            pptx._chartMarkerData.push(markerData);

        } else {
            // Bar charts and others - standard pptxgenjs behavior
            pptSlide.addChart(pptxChartType, pptxData, validatedOptions);

            if ((chartType === 'pie' || chartType === 'doughnut') && chartData.dataLabels) {
                if (!pptx._chartMarkerData) {
                    pptx._chartMarkerData = [];
                }

                pptx._chartMarkerData.push({
                    type: 'pieDataLabels',
                    chartIndex,
                    chartType,
                    dataLabels: chartData.dataLabels
                });
            }
        }

        pptx._customChartCounter = chartIndex + 1;

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

function extractPieChartData(chartContainer) {
    try {
        const decodeJsonAttr = (attrName, fallback = []) => {
            const raw = chartContainer.getAttribute(attrName);
            if (!raw) return fallback;
            try {
                return JSON.parse(decodeURIComponent(raw));
            } catch {
                return fallback;
            }
        };

        const title = chartContainer.getAttribute('data-chart-title') || '';
        const categories = decodeJsonAttr('data-chart-labels', []);
        const values = decodeJsonAttr('data-chart-values', []).map(v => parseFloat(v) || 0);
        const colors = decodeJsonAttr('data-chart-colors', []);
        const dataLabelMeta = decodeJsonAttr('data-chart-data-labels', null);
        const holeSize = parseInt(chartContainer.getAttribute('data-chart-hole-size') || '50', 10);
        const firstSliceAng = parseInt(chartContainer.getAttribute('data-chart-first-slice-ang') || '0', 10);
        const legendPosRaw = (chartContainer.getAttribute('data-chart-legend-pos') || 'bottom').toLowerCase();
        const explicitShowValue = chartContainer.getAttribute('data-chart-show-value');
        const explicitShowPercent = chartContainer.getAttribute('data-chart-show-percent');
        const explicitShowLeaderLines = chartContainer.getAttribute('data-chart-show-leader-lines');
        const explicitFormatCode = chartContainer.getAttribute('data-chart-label-format');
        const explicitLabelColor = chartContainer.getAttribute('data-chart-label-color');
        const visibleTextNodes = Array.from(chartContainer.querySelectorAll('.doughnut-chart div, .pie-chart div'))
            .map(node => (node.textContent || '').trim())
            .filter(Boolean);
        const hasPercentLabelText = visibleTextNodes.some(text => /-?\d+(?:\.\d+)?%$/.test(text));
        const hasValueLabelText = visibleTextNodes.some(text => /^-?\d+(?:\.\d+)?$/.test(text));
        const showValue = explicitShowValue ? explicitShowValue === '1' : hasValueLabelText;
        const showPercent = explicitShowPercent ? explicitShowPercent === '1' : false;
        const showLeaderLines = explicitShowLeaderLines ? explicitShowLeaderLines === '1' : hasPercentLabelText;
        const labelFormatCode = explicitFormatCode || (hasPercentLabelText ? '0.00%' : 'General');
        const labelColor = explicitLabelColor || '#4D4D4D';
        const legendPosMap = {
            bottom: 'b',
            top: 't',
            left: 'l',
            right: 'r'
        };

        if (categories.length > 0 && values.length === categories.length) {
            return {
                title,
                categories,
                series: [
                    {
                        name: title || 'Sales',
                        values,
                        color: colors[0] || '#4472C4'
                    }
                ],
                colors,
                holeSize: Number.isFinite(holeSize) ? holeSize : 50,
                firstSliceAng: Number.isFinite(firstSliceAng) ? firstSliceAng : 0,
                legendPos: legendPosMap[legendPosRaw] || 'b',
                dataLabels: {
                    showValue,
                    showPercent,
                    showLeaderLines,
                    formatCode: labelFormatCode,
                    color: labelColor,
                    items: Array.isArray(dataLabelMeta?.items) ? dataLabelMeta.items : []
                }
            };
        }
    } catch (error) {
        console.error("Error in extractPieChartData:", error);
    }

    return {
        title: '',
        categories: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
        series: [{ name: 'Sales', values: [8.2, 3.2, 1.4, 1.2], color: '#4472C4' }],
        colors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000'],
        holeSize: 50,
        firstSliceAng: 0,
        legendPos: 'b',
        dataLabels: {
            showValue: false,
            showPercent: false,
            showLeaderLines: false,
            formatCode: 'General',
            color: '#4D4D4D',
            items: []
        }
    };
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
        seriesOptions: null
    };

    try {
        // Extract colors from series data
        if (chartData && chartData.series) {
            styling.colors = chartData.series.map(s => s.color || s.lineColor || '#4472C4');
            styling.seriesOptions = chartData.series;
        }

        if (chartData && Array.isArray(chartData.colors) && chartData.colors.length > 0) {
            styling.colors = chartData.colors;
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
        'doughnut': 'doughnut',
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

function parsePxValue(value, fallback = 0) {
    if (value == null) return fallback;
    const match = String(value).match(/-?\d+(?:\.\d+)?/);
    return match ? parseFloat(match[0]) : fallback;
}

function escapeSvgText(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function darkenHexColor(color, factor = 0.62) {
    const hex = normalizeColor(color, true);
    const r = Math.max(0, Math.min(255, Math.round(parseInt(hex.slice(0, 2), 16) * factor)));
    const g = Math.max(0, Math.min(255, Math.round(parseInt(hex.slice(2, 4), 16) * factor)));
    const b = Math.max(0, Math.min(255, Math.round(parseInt(hex.slice(4, 6), 16) * factor)));
    return [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

function polarToCartesian(radius, angleDeg) {
    const angleRad = (angleDeg - 90) * Math.PI / 180;
    return {
        x: radius * Math.cos(angleRad),
        y: radius * Math.sin(angleRad)
    };
}

function buildPieSlicePath(radius, startAngle, endAngle) {
    const start = polarToCartesian(radius, startAngle);
    const end = polarToCartesian(radius, endAngle);
    const largeArcFlag = endAngle - startAngle > 180 ? 1 : 0;
    return [
        `M 0 0`,
        `L ${start.x.toFixed(4)} ${start.y.toFixed(4)}`,
        `A ${radius.toFixed(4)} ${radius.toFixed(4)} 0 ${largeArcFlag} 1 ${end.x.toFixed(4)} ${end.y.toFixed(4)}`,
        'Z'
    ].join(' ');
}

function shouldRenderPieChartAsSvg(chartContainer, chartType) {
    return chartType === 'pie' || chartType === 'ofPie';
}

function parseTranslatePercent(styleText) {
    const transformMatch = styleText.match(/transform:\s*translate\(\s*([-0-9.]+)%\s*,\s*([-0-9.]+)%\s*\)/i);
    if (!transformMatch) {
        return { x: -50, y: -50 };
    }

    return {
        x: parseFloat(transformMatch[1]) || 0,
        y: parseFloat(transformMatch[2]) || 0
    };
}

function parseStyleNumber(styleText, propertyName, fallback) {
    const match = styleText.match(new RegExp(`${propertyName}:\\s*([^;]+)`, 'i'));
    return parsePxValue(match?.[1], fallback);
}

function parsePieLayerGeometry(styleText, fallbackWidth, fallbackHeight) {
    return {
        left: parseStyleNumber(styleText, 'left', 0),
        top: parseStyleNumber(styleText, 'top', 0),
        width: parseStyleNumber(styleText, 'width', fallbackWidth),
        height: parseStyleNumber(styleText, 'height', fallbackHeight),
        scaleX: parsePxValue(styleText.match(/scaleX\(([^)]+)\)/i)?.[1], 1),
        scaleY: parsePxValue(styleText.match(/scaleY\(([^)]+)\)/i)?.[1], 1),
        rotation: parsePxValue(styleText.match(/rotate\(([-0-9.]+)deg\)/i)?.[1], 0)
    };
}

function estimateLegendLayout(chartContainer, outerW) {
    const legendEl = chartContainer.querySelector('.chart-legend');
    if (!legendEl) {
        return { rows: [], height: 0, top: 0 };
    }

    const legendStyle = legendEl.getAttribute('style') || '';
    const gap = parseStyleNumber(legendStyle, 'gap', 14);
    const marginTop = parseStyleNumber(legendStyle, 'margin-top', 14);
    const availableWidth = Math.max(outerW - 20, 40);
    const entries = Array.from(legendEl.querySelectorAll(':scope > div')).map((entry, originalIndex) => {
        const boxStyle = entry.querySelector('div')?.getAttribute('style') || '';
        const text = entry.querySelector('span')?.textContent?.trim() || '';
        const textStyle = entry.querySelector('span')?.getAttribute('style') || '';
        const fontSize = parseStyleNumber(textStyle, 'font-size', 11);
        const itemHeight = Math.max(fontSize * 1.35, 8) + 4;
        const approxTextWidth = text.length * fontSize * 0.56;
        const itemWidth = 8 + 5 + approxTextWidth;
        return {
            entry,
            text,
            boxStyle,
            textStyle,
            fontSize,
            itemHeight,
            itemWidth,
            originalIndex
        };
    });

    const rows = [];
    let currentRow = [];
    let currentWidth = 0;
    for (const item of entries) {
        const nextWidth = currentRow.length === 0 ? item.itemWidth : currentWidth + gap + item.itemWidth;
        if (currentRow.length > 0 && nextWidth > availableWidth) {
            rows.push(currentRow);
            currentRow = [item];
            currentWidth = item.itemWidth;
        } else {
            currentRow.push(item);
            currentWidth = nextWidth;
        }
    }
    if (currentRow.length) {
        rows.push(currentRow);
    }

    const rowGap = Math.max(6, gap * 0.5);
    const rowsHeight = rows.reduce((sum, row, idx) => {
        const rowHeight = Math.max(...row.map(item => item.itemHeight), 16);
        return sum + rowHeight + (idx > 0 ? rowGap : 0);
    }, 0);

    return {
        rows,
        gap,
        marginTop,
        rowGap,
        height: rows.length ? rowsHeight + marginTop : 0
    };
}

let pieChartCaptureBrowserPromise = null;
let pieChartTempDirPromise = null;
const PIE_CHART_EXPORT_PADDING = {
    top: 4,
    right: 8,
    bottom: 14,
    left: 8
};

async function getPieChartCaptureBrowser() {
    if (!pieChartCaptureBrowserPromise) {
        const puppeteer = require('puppeteer');
        pieChartCaptureBrowserPromise = puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        }).catch((error) => {
            pieChartCaptureBrowserPromise = null;
            throw error;
        });
    }

    return pieChartCaptureBrowserPromise;
}

async function getPieChartTempDir() {
    if (!pieChartTempDirPromise) {
        const fs = require('fs').promises;
        const path = require('path');
        const tempDir = path.resolve(process.cwd(), 'tmp', 'pie-chart-export');
        pieChartTempDirPromise = fs.mkdir(tempDir, { recursive: true })
            .then(() => tempDir)
            .catch((error) => {
                pieChartTempDirPromise = null;
                throw error;
            });
    }

    return pieChartTempDirPromise;
}

async function renderPieChartHtmlToPngBuffer(chartContainer) {
    const containerStyle = chartContainer.getAttribute('style') || '';
    const width = Math.max(1, parseStyleNumber(containerStyle, 'width', 320));
    const height = Math.max(1, parseStyleNumber(containerStyle, 'height', 213.33));
    const paddedWidth = width + PIE_CHART_EXPORT_PADDING.left + PIE_CHART_EXPORT_PADDING.right;
    const paddedHeight = height + PIE_CHART_EXPORT_PADDING.top + PIE_CHART_EXPORT_PADDING.bottom;
    const browser = await getPieChartCaptureBrowser();
    const page = await browser.newPage();
    const captureNode = chartContainer.cloneNode(true);
    const captureStyle = captureNode.getAttribute('style') || '';
    const normalizedCaptureStyle = captureStyle
        .replace(/left:\s*[^;]+;?/i, 'left:0px;')
        .replace(/top:\s*[^;]+;?/i, 'top:0px;')
        .replace(/margin:\s*[^;]+;?/i, 'margin:0;')
        .replace(/position:\s*[^;]+;?/i, 'position:relative;');
    captureNode.setAttribute('style', normalizedCaptureStyle);

    try {
        await page.setViewport({
            width: Math.max(1, Math.ceil(paddedWidth)),
            height: Math.max(1, Math.ceil(paddedHeight)),
            deviceScaleFactor: 2
        });

        const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    html, body {
      margin: 0;
      padding: 0;
      background: transparent;
      width: ${paddedWidth}px;
      height: ${paddedHeight}px;
      overflow: visible;
    }
    body {
      font-family: Arial, sans-serif;
    }
    #capture-root {
      position: relative;
      width: ${paddedWidth}px;
      height: ${paddedHeight}px;
      padding: ${PIE_CHART_EXPORT_PADDING.top}px ${PIE_CHART_EXPORT_PADDING.right}px ${PIE_CHART_EXPORT_PADDING.bottom}px ${PIE_CHART_EXPORT_PADDING.left}px;
      box-sizing: border-box;
      overflow: visible;
      background: transparent;
    }
  </style>
</head>
<body>
  <div id="capture-root">${captureNode.outerHTML}</div>
</body>
</html>`;

        await page.setContent(html, { waitUntil: 'domcontentloaded' });
        const root = await page.$('#capture-root');
        if (!root) {
            throw new Error('Pie chart capture root not found');
        }

        const pngBuffer = await root.screenshot({
            omitBackground: true,
            type: 'png'
        });

        return pngBuffer;
    } finally {
        await page.close();
    }
}

async function writePieChartTempAsset(bufferOrText, extension) {
    const fs = require('fs').promises;
    const path = require('path');
    const tempDir = await getPieChartTempDir();
    const fileName = `pie_chart_${Date.now()}_${Math.random().toString(36).slice(2)}.${extension}`;
    const filePath = path.join(tempDir, fileName);
    await fs.writeFile(filePath, bufferOrText);
    return filePath;
}

function buildStyledPieChartSvg(chartContainer) {
    const containerStyle = chartContainer.getAttribute('style') || '';
    const outerW = Math.max(1, parsePxValue(containerStyle.match(/width:\s*([^;]+)/i)?.[1], 320));
    const outerH = Math.max(1, parsePxValue(containerStyle.match(/height:\s*([^;]+)/i)?.[1], 213.33));
    const exportOuterW = outerW + PIE_CHART_EXPORT_PADDING.left + PIE_CHART_EXPORT_PADDING.right;
    const exportOuterH = outerH + PIE_CHART_EXPORT_PADDING.top + PIE_CHART_EXPORT_PADDING.bottom;
    const wrapperStyle = chartContainer.querySelector('.doughnut-wrapper')?.getAttribute('style') || '';
    const wrapperPadTop = parseStyleNumber(wrapperStyle, 'padding', 10);
    const wrapperPadBottom = wrapperPadTop;

    const chartTitleEl = chartContainer.querySelector('.chart-title');
    const chartTitle = chartTitleEl?.textContent?.trim() || chartContainer.getAttribute('data-chart-title') || '';
    const chartTitleStyle = chartTitleEl?.getAttribute('style') || '';
    const chartTitleColor = normalizeColor(chartTitleStyle.match(/color:\s*([^;]+)/i)?.[1] || '#666666');
    const chartTitleSize = parsePxValue(chartTitleStyle.match(/font-size:\s*([^;]+)/i)?.[1], 12);
    const chartTitleLineHeight = parsePxValue(chartTitleStyle.match(/line-height:\s*([^;]+)/i)?.[1], chartTitleSize * 1.2);
    const chartTitleMarginBottom = parseStyleNumber(chartTitleStyle, 'margin-bottom', 8);

    const chartSurface = chartContainer.querySelector('.doughnut-chart');
    const surfaceStyle = chartSurface?.getAttribute('style') || '';
    const chartW = Math.max(1, parsePxValue(surfaceStyle.match(/width:\s*([^;]+)/i)?.[1], 120));
    const chartH = Math.max(1, parsePxValue(surfaceStyle.match(/height:\s*([^;]+)/i)?.[1], 120));

    const values = (() => {
        try {
            return JSON.parse(decodeURIComponent(chartContainer.getAttribute('data-chart-values') || '[]')).map(v => parseFloat(v) || 0);
        } catch {
            return [];
        }
    })();
    const labels = (() => {
        try {
            return JSON.parse(decodeURIComponent(chartContainer.getAttribute('data-chart-labels') || '[]'));
        } catch {
            return [];
        }
    })();
    const colors = (() => {
        try {
            return JSON.parse(decodeURIComponent(chartContainer.getAttribute('data-chart-colors') || '[]'));
        } catch {
            return [];
        }
    })();

    const total = values.reduce((sum, value) => sum + value, 0) || 1;
    const absoluteLayers = Array.from(chartSurface?.children || []).filter((node) => {
        const style = node.getAttribute?.('style') || '';
        const isHole = node.classList && node.classList.contains('doughnut-hole');
        const isLabel = /font-size:\s*[^;]+/i.test(style);
        return /position:\s*absolute/i.test(style) && !isHole && !isLabel;
    });
    const topLayerStyle = absoluteLayers.find((node) => {
        const style = node.getAttribute('style') || '';
        return /background:/i.test(style) && /scaleX\(/i.test(style) && !/translateY\(/i.test(style);
    })?.getAttribute('style') || '';
    const depthLayerStyle = absoluteLayers.find((node) => {
        const style = node.getAttribute('style') || '';
        return /background:/i.test(style) && /translateY\(/i.test(style);
    })?.getAttribute('style') || '';
    const has3DTransforms = Boolean(topLayerStyle || depthLayerStyle);
    const topGeometry = parsePieLayerGeometry(topLayerStyle, chartW, chartH);
    const depthGeometry = parsePieLayerGeometry(depthLayerStyle, topGeometry.width, topGeometry.height);
    const scaleX = topGeometry.scaleX;
    const scaleY = topGeometry.scaleY;
    const depth = depthLayerStyle
        ? parsePxValue(depthLayerStyle.match(/translateY\(([^)]+)\)/i)?.[1], 0)
        : Math.max(0, depthGeometry.top - topGeometry.top);
    const visualChartHeight = chartH * scaleY + depth;
    const titleBlockHeight = chartTitle ? chartTitleLineHeight + chartTitleMarginBottom : 0;
    const legendLayout = estimateLegendLayout(chartContainer, outerW);
    const availableChartBand = Math.max(0, outerH - wrapperPadTop - wrapperPadBottom - titleBlockHeight - legendLayout.height);
    const chartAreaTop = wrapperPadTop + titleBlockHeight + Math.max(0, (availableChartBand - visualChartHeight) / 2);
    const chartX = (outerW - chartW) / 2;
    const chartY = chartAreaTop;
    const centerX = chartX + chartW / 2;
    const centerY = chartY + chartH / 2;
    const radius = Math.min(chartW, chartH) / 2;

    let runningAngle = 0;
    const sliceSvgs = values.map((value, index) => {
        const sweep = (value / total) * 360;
        const startAngle = runningAngle;
        const endAngle = runningAngle + sweep;
        runningAngle = endAngle;
        const topRadius = Math.min(topGeometry.width || chartW, topGeometry.height || chartH) / 2;
        const pathD = buildPieSlicePath(topRadius, startAngle, endAngle);
        const topColor = normalizeColor(colors[index] || '#4472C4');
        const sideColor = `#${darkenHexColor(colors[index] || '#4472C4')}`;
        const topCenterX = chartX + topGeometry.left + (topGeometry.width / 2);
        const topCenterY = chartY + topGeometry.top + (topGeometry.height / 2);
        const sideCenterX = chartX + depthGeometry.left + (depthGeometry.width / 2);
        const sideCenterY = chartY + depthGeometry.top + (depthGeometry.height / 2);

        return {
            side: has3DTransforms
                ? `<path d="${pathD}" fill="${sideColor}" stroke="#FFFFFF" stroke-width="1" transform="translate(${sideCenterX.toFixed(3)} ${sideCenterY.toFixed(3)}) rotate(${depthGeometry.rotation || topGeometry.rotation || 0}) scale(${depthGeometry.scaleX || scaleX} ${depthGeometry.scaleY || scaleY})"/>`
                : '',
            top: `<path d="${pathD}" fill="${topColor}" stroke="#FFFFFF" stroke-width="1" transform="translate(${topCenterX.toFixed(3)} ${topCenterY.toFixed(3)}) rotate(${topGeometry.rotation || 0}) scale(${scaleX} ${scaleY})"/>`
        };
    });

    const lineSvgs = Array.from(chartSurface?.querySelectorAll('svg polyline, svg line, svg path') || [])
        .map((node) => {
            const tag = node.tagName.toLowerCase();
            const stroke = normalizeColor(node.getAttribute('stroke') || '#AAAAAA');
            const strokeWidth = parsePxValue(node.getAttribute('stroke-width') || '1', 1);
            const fill = node.getAttribute('fill') || 'none';

            if (tag === 'polyline') {
                const points = (node.getAttribute('points') || '').trim();
                if (!points) return '';
                return `<polyline points="${escapeSvgText(points)}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"/>`;
            }

            if (tag === 'line') {
                const x1 = parsePxValue(node.getAttribute('x1') || '0', 0);
                const y1 = parsePxValue(node.getAttribute('y1') || '0', 0);
                const x2 = parsePxValue(node.getAttribute('x2') || '0', 0);
                const y2 = parsePxValue(node.getAttribute('y2') || '0', 0);
                return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"/>`;
            }

            const d = node.getAttribute('d') || '';
            if (!d) return '';
            return `<path d="${escapeSvgText(d)}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}"/>`;
        })
        .filter(Boolean)
        .join('');

    const labelSvgs = Array.from(chartSurface?.querySelectorAll('div') || [])
        .filter(node => !(node.classList && node.classList.contains('doughnut-hole')))
        .map(node => {
            const text = (node.textContent || '').trim();
            const style = node.getAttribute('style') || '';
            if (!text || !/position:\s*absolute/i.test(style)) return '';

            const left = parsePxValue(style.match(/left:\s*([^;]+)/i)?.[1], 0);
            const top = parsePxValue(style.match(/top:\s*([^;]+)/i)?.[1], 0);
            const color = normalizeColor(style.match(/color:\s*([^;]+)/i)?.[1] || '#FFFFFF');
            const fontSize = parsePxValue(style.match(/font-size:\s*([^;]+)/i)?.[1], 11);
            const translate = parseTranslatePercent(style);
            const rawTextAlign = (style.match(/text-align:\s*([^;]+)/i)?.[1] || '').trim().toLowerCase();

            let textAnchor = 'middle';
            if (translate.x <= -90 || rawTextAlign === 'right') {
                textAnchor = 'end';
            } else if (translate.x >= -10 || rawTextAlign === 'left') {
                textAnchor = 'start';
            }

            return `<text x="${(chartX + left).toFixed(3)}" y="${(chartY + top).toFixed(3)}" fill="${color}" font-family="Arial, sans-serif" font-size="${fontSize}" text-anchor="${textAnchor}" dominant-baseline="middle">${escapeSvgText(text)}</text>`;
        })
        .filter(Boolean)
        .join('');

    const legendBaseTop = outerH - wrapperPadBottom - Math.max(0, legendLayout.height - legendLayout.marginTop);
    const legendSvgs = legendLayout.rows.map((row, rowIndex) => {
        const rowWidth = row.reduce((sum, item, idx) => sum + item.itemWidth + (idx > 0 ? legendLayout.gap : 0), 0);
        let cursorX = (outerW - rowWidth) / 2;
        const rowHeight = Math.max(...row.map(item => item.itemHeight), 16);
        const rowY = legendBaseTop + rowIndex * (rowHeight + legendLayout.rowGap);

        return row.map((item, index) => {
            const fill = normalizeColor(item.boxStyle.match(/background:\s*([^;]+)/i)?.[1] || colors[item.originalIndex] || '#4472C4');
            const textColor = normalizeColor(item.textStyle.match(/color:\s*([^;]+)/i)?.[1] || '#555555');
            const squareX = cursorX;
            const squareY = rowY + (rowHeight - 8) / 2;
            const textX = squareX + 12;
            const textY = rowY + Math.max(1, (rowHeight - item.fontSize) / 2);
            const svg = [
                `<rect x="${squareX.toFixed(3)}" y="${squareY.toFixed(3)}" width="8" height="8" fill="${fill}"/>`,
                `<text x="${textX.toFixed(3)}" y="${textY.toFixed(3)}" fill="${textColor}" font-family="Arial, sans-serif" font-size="${item.fontSize}" dominant-baseline="text-before-edge">${escapeSvgText(item.text || labels[item.originalIndex] || '')}</text>`
            ].join('');
            cursorX += item.itemWidth + legendLayout.gap;
            return svg;
        }).join('');
    }).join('');

    const titleSvg = chartTitle
        ? `<text x="${(outerW / 2).toFixed(3)}" y="${Math.max(12, chartTitleSize + 2).toFixed(3)}" fill="${chartTitleColor}" font-family="Arial, sans-serif" font-size="${chartTitleSize}" text-anchor="middle">${escapeSvgText(chartTitle)}</text>`
        : '';

    return [
        `<svg xmlns="http://www.w3.org/2000/svg" width="${exportOuterW}" height="${exportOuterH}" viewBox="0 0 ${exportOuterW} ${exportOuterH}">`,
        `<rect width="100%" height="100%" fill="transparent"/>`,
        `<g transform="translate(${PIE_CHART_EXPORT_PADDING.left} ${PIE_CHART_EXPORT_PADDING.top})">`,
        titleSvg,
        sliceSvgs.map(slice => slice.side).join(''),
        sliceSvgs.map(slice => slice.top).join(''),
        lineSvgs ? `<g transform="translate(${chartX.toFixed(3)} ${chartY.toFixed(3)})">${lineSvgs}</g>` : '',
        labelSvgs,
        legendSvgs,
        `</g>`,
        `</svg>`
    ].join('');
}

async function addStyledPieChartImage(pptSlide, chartContainer, slideContext) {
    const position = extractChartPosition(chartContainer, slideContext);
    let imagePath;

    try {
        const pngBuffer = await renderPieChartHtmlToPngBuffer(chartContainer);
        imagePath = await writePieChartTempAsset(pngBuffer, 'png');
    } catch (error) {
        console.error('Pie chart HTML capture failed, falling back to SVG rebuild:', error);
        const svgMarkup = buildStyledPieChartSvg(chartContainer);
        imagePath = await writePieChartTempAsset(svgMarkup, 'svg');
    }

    pptSlide.addImage({
        path: imagePath,
        x: position.x,
        y: position.y,
        w: position.w,
        h: position.h,
        objectName: chartContainer.getAttribute('data-name') || 'Styled Pie Chart'
    });

    return true;
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

            if (markerData.type === 'pieDataLabels') {
                const dataLabelsXml = buildPieLikeDataLabelsXML(markerData.dataLabels);
                if (/<c:dLbls>[\s\S]*?<\/c:dLbls>/.test(chartXML)) {
                    chartXML = chartXML.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, dataLabelsXml);
                } else if (markerData.chartType === 'doughnut') {
                    chartXML = chartXML.replace(/(<c:doughnutChart>[\s\S]*?)(<c:firstSliceAng\b[^>]*\/>)/, `$1${dataLabelsXml}$2`);
                } else {
                    chartXML = chartXML.replace(/(<c:pieChart>[\s\S]*?)(<c:firstSliceAng\b[^>]*\/>)/, `$1${dataLabelsXml}$2`);
                }
            } else {
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
            }

            zip.file(chartPath, chartXML);
        }

        const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
        await fs.writeFile(pptxPath, modifiedBuffer);



    } catch (error) {
        console.error('   ⚠️ Error post-processing chart markers:', error);
    }
}

function buildPieLikeDataLabelsXML(dataLabels) {
    const showValue = dataLabels?.showValue ? '1' : '0';
    const showPercent = dataLabels?.showPercent ? '1' : '0';
    const showLeaderLines = dataLabels?.showLeaderLines ? '1' : '0';
    const formatCode = escapeXmlAttr(dataLabels?.formatCode || 'General');
    const labelColor = (dataLabels?.color || '4D4D4D').replace('#', '').toUpperCase();
    const globalDLblPos = dataLabels?.dLblPos ? escapeXmlAttr(dataLabels.dLblPos) : '';
    const items = Array.isArray(dataLabels?.items) ? dataLabels.items : [];

    const perPoint = items.map((item) => {
        const idx = Number.isFinite(item?.idx) ? item.idx : 0;
        const x = Number.isFinite(item?.x) ? item.x : 0;
        const y = Number.isFinite(item?.y) ? item.y : 0;
        const hasManualLayout = Boolean(item?.hasManualLayout);
        const itemDLblPos = item?.dLblPos ? `<c:dLblPos val="${escapeXmlAttr(item.dLblPos)}"/>` : '';
        const itemLabelColor = item?.textColor ? String(item.textColor).replace('#', '').toUpperCase() : '';
        const itemTxPr = itemLabelColor
            ? `<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:srgbClr val="${itemLabelColor}"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`
            : '';
        const layoutXml = hasManualLayout
            ? `<c:layout><c:manualLayout><c:x val="${x}"/><c:y val="${y}"/></c:manualLayout></c:layout>`
            : '';
        return `<c:dLbl><c:idx val="${idx}"/><c:numFmt formatCode="${formatCode}" sourceLinked="0"/>${layoutXml}${itemTxPr}${itemDLblPos}<c:showLegendKey val="0"/><c:showVal val="${showValue}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${showPercent}"/><c:showBubbleSize val="0"/></c:dLbl>`;
    }).join('');

    return `<c:dLbls>${perPoint}<c:numFmt formatCode="${formatCode}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:srgbClr val="${labelColor}"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>${globalDLblPos ? `<c:dLblPos val="${globalDLblPos}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="${showValue}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${showPercent}"/><c:showBubbleSize val="0"/><c:showLeaderLines val="${showLeaderLines}"/></c:dLbls>`;
}

function escapeXmlAttr(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}
// Export functions for use in the main conversion system
module.exports = {
    addChartToSlide,
    processChartElement,
    isChartElement,
    postProcessChartXML,
    postProcessPPTXForMarkers
};
