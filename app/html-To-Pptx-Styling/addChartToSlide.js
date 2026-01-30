
function addChartToSlide(pptx, pptSlide, element, slideContext) {
    try {

        // Check if this is a chart container
        if (!isChartElement(element)) {
            return false;
        }

        // Extract chart data from HTML
        const chartData = extractChartDataFromHTML(element);
        if (!chartData || !chartData.isValid) {
            return false;
        }


        // Extract positioning
        const position = extractChartPosition(element, slideContext);

        // Extract styling WITH Y-axis config
        const styling = extractChartStyling(element, chartData);

        // Convert to pptxgenjs format
        const pptxChartData = convertToPptxFormat(chartData);

        const shapeName = element.getAttribute('data-name') || '';

        // Create chart options
        const chartOptions = createChartOptions(chartData, position, styling, shapeName, pptxChartData);

        // Clean up series names
        pptxChartData.forEach((s, i) => {
            if (typeof s.name !== 'string') {
                let rawName = s.name;
                if (rawName && typeof rawName === 'object') {
                    s.name = rawName.name || rawName.label || rawName.id || `Series ${i + 1}`;
                } else {
                    s.name = `Series ${i + 1}`;
                }
            }
        });

        // Add chart to slide
        pptSlide.addChart(getChartType(chartData.type), pptxChartData, chartOptions);

        return true;

    } catch (error) {
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
            case 'area':
                extractedData = extractAreaChartData(chartContainer);
                break;
            case 'pie':
                extractedData = extractPieChartData(chartContainer);
                break;
            case 'doughnut':
                extractedData = extractDoughnutChartData(chartContainer);
                break;
            default:
                extractedData = extractBarChartData(chartContainer);
        }


        return {
            ...extractedData,
            title,
            type: chartType,
            isValid: extractedData && extractedData.categories && extractedData.series && extractedData.series.length > 0
        };

    } catch (error) {
        return { isValid: false };
    }
}

function detectChartType(chartContainer) {

    // Check data attribute first
    const dataType = chartContainer.getAttribute('data-chart-type');
    if (dataType) {
        return dataType.toLowerCase();
    }

    // Check data-name attribute - THIS IS KEY FOR YOUR CHART
    const dataName = chartContainer.getAttribute('data-name');

    if (dataName) {
        const nameLower = dataName.toLowerCase();
        if (nameLower.includes('area')) {
            return 'area';
        }
        if (nameLower.includes('doughnut') || nameLower.includes('donut')) return 'doughnut';
        if (nameLower.includes('bar') || nameLower.includes('column')) return 'bar';
        if (nameLower.includes('line')) return 'line';
        if (nameLower.includes('pie')) return 'pie';
    }

    // Check for SVG structure
    const svg = chartContainer.querySelector('svg');
    if (svg) {
        const path = svg.querySelector('path[fill]:not([fill="white"]):not([fill="none"])');
        const circles = svg.querySelectorAll('circle');

        if (path && circles.length > 0) {
            const d = path.getAttribute('d');
            if (d && d.toUpperCase().includes('Z')) {
                return 'area';
            }
        }
    }

    // Fallback checks
    if (chartContainer.querySelector('.doughnut-chart, .doughnut-wrapper')) return 'doughnut';
    if (chartContainer.querySelector('.bar, [data-series]')) return 'bar';
    if (chartContainer.querySelector('.pie, .slice, .arc')) return 'pie';

    const classList = chartContainer.className.toLowerCase();
    if (classList.includes('area')) return 'area';
    if (classList.includes('doughnut') || classList.includes('donut')) return 'doughnut';
    if (classList.includes('bar') || classList.includes('column')) return 'bar';
    if (classList.includes('line')) return 'line';
    if (classList.includes('pie')) return 'pie';

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
                    seriesNames.set(
                        seriesIndex,
                        typeof seriesName === "string" ? seriesName : JSON.stringify(seriesName.name || seriesName || `Series ${seriesIndex + 1}`)
                    );

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

function extractDoughnutChartData(chartContainer) {
    try {

        const doughnutChart = chartContainer.querySelector('.doughnut-chart');
        if (!doughnutChart) {
            return getDefaultDoughnutData();
        }

        const styleAttr = doughnutChart.getAttribute('style');
        if (!styleAttr) {
            return getDefaultDoughnutData();
        }

        // Extract rotation
        const rotationMatch = styleAttr.match(/transform:\s*rotate\(([\d.-]+)deg\)/i);
        const rotation = rotationMatch ? parseFloat(rotationMatch[1]) : 0;

        // Parse conic-gradient
        const gradientMatch = styleAttr.match(/background:\s*conic-gradient\((.*?)\);/);
        if (!gradientMatch) {
            return getDefaultDoughnutData();
        }

        const gradientString = gradientMatch[1];
        const segments = parseConicGradient(gradientString);

        if (!segments || segments.length === 0) {
            return getDefaultDoughnutData();
        }

        // Filter out white separators
        const dataSegments = segments.filter(seg => {
            const color = seg.color.toUpperCase();
            return color !== '#FFFFFF' && color !== '#FFF' && color !== 'FFFFFF';
        });

        if (dataSegments.length === 0) {
            return getDefaultDoughnutData();
        }

        const categories = [];
        const values = [];
        const colors = [];

        dataSegments.forEach((segment, index) => {
            categories.push(`Segment ${index + 1}`);
            values.push(segment.value);
            colors.push(segment.color);
        });

        return {
            categories,
            series: [{
                name: 'Values',
                values: values,
                colors: colors
            }],
            maxValue: Math.max(...values),
            minValue: 0,
            isDoughnut: true,
            rotation: rotation
        };

    } catch (error) {
        return getDefaultDoughnutData();
    }
}


// 1. First, add the area chart extraction function
function extractAreaChartData(chartContainer) {
    try {

        // Get chart area for height calculation
        const chartArea = chartContainer.querySelector('.chart-area');
        if (!chartArea) {
            return getDefaultAreaChartData();
        }

        const chartAreaStyle = chartArea.getAttribute('style');
        const heightMatch = chartAreaStyle?.match(/height:\s*([\d.]+)px/);
        const chartHeight = heightMatch ? parseFloat(heightMatch[1]) : 201.5784251968504;

        // Extract Y-axis range
        const yAxisLabels = Array.from(chartContainer.querySelectorAll('[style*="text-align: right"]'))
            .map(el => {
                const text = el.textContent.trim().replace(/,/g, '');
                return parseFloat(text);
            })
            .filter(val => !isNaN(val))
            .sort((a, b) => a - b);

        const minY = yAxisLabels.length > 0 ? yAxisLabels[0] : 0;
        const maxY = yAxisLabels.length > 0 ? yAxisLabels[yAxisLabels.length - 1] : 90000;

        // Extract all circles
        const svg = chartContainer.querySelector('svg');
        if (!svg) {
            return getDefaultAreaChartData();
        }

        const circles = Array.from(svg.querySelectorAll('circle'));

        if (circles.length === 0) {
            return getDefaultAreaChartData();
        }

        // Get color
        let color = circles[0].getAttribute('fill') || '#963963';

        // Extract and convert all data points
        const dataPoints = circles.map((circle, idx) => {
            const cx = parseFloat(circle.getAttribute('cx'));
            const cy = parseFloat(circle.getAttribute('cy'));

            // Convert cy to actual value
            const value = maxY - ((cy / chartHeight) * (maxY - minY));

            return {
                cx: cx,
                cy: cy,
                value: Math.round(value)
            };
        }).sort((a, b) => a.cx - b.cx);

        // Extract visible X-axis labels
        const xAxisLabels = Array.from(chartContainer.querySelectorAll('[style*="bottom: -25px"]'))
            .map(el => ({
                text: el.textContent.trim(),
                left: parseFloat(el.style.left)
            }))
            .filter(label => label.text && !isNaN(label.left))
            .sort((a, b) => a.left - b.left);


        // Generate categories
        let categories = [];
        if (xAxisLabels.length >= 2 && /^\d{4}$/.test(xAxisLabels[0].text)) {
            const firstYear = parseInt(xAxisLabels[0].text);
            for (let i = 0; i < dataPoints.length; i++) {
                categories.push(String(firstYear + i));
            }
        } else {
            for (let i = 0; i < dataPoints.length; i++) {
                categories.push(`Point ${i + 1}`);
            }
        }

        const values = dataPoints.map(dp => dp.value);

        const result = {
            categories,
            series: [{
                name: 'Series 1',
                values: values,
                color: normalizeColor(color)
            }],
            maxValue: Math.max(...values),
            minValue: Math.min(...values, 0)
        };

        return result;

    } catch (error) {
        return getDefaultAreaChartData();
    }
}

function getDefaultAreaChartData() {
    return {
        categories: ['Point 1', 'Point 2', 'Point 3'],
        series: [{
            name: 'Series 1',
            values: [30, 50, 40],
            color: '#963963'
        }],
        maxValue: 50,
        minValue: 0
    };
}



/**
 * Parse conic-gradient string to extract color segments
 */
function parseConicGradient(gradientString) {
    try {
        let cleanString = gradientString.replace(/from\s+[\d.-]+deg,?\s*/, '').trim();
        const colorStopRegex = /(#[0-9A-Fa-f]{6}|#[0-9A-Fa-f]{3}|\w+)\s+([\d.]+)deg\s+([\d.]+)deg/g;

        const segments = [];
        let match;

        while ((match = colorStopRegex.exec(cleanString)) !== null) {
            const color = match[1];
            let startDeg = parseFloat(match[2]);
            let endDeg = parseFloat(match[3]);

            let angleDiff = endDeg - startDeg;
            if (angleDiff < 0) angleDiff += 360;
            if (angleDiff < 2) continue; // Skip separators

            const percentage = parseFloat(((angleDiff / 360) * 100).toFixed(2));
            const normalizedColor = color.startsWith('#') ? color.toUpperCase() : '#' + color.toUpperCase();

            segments.push({
                color: normalizedColor,
                startDeg: startDeg,
                endDeg: endDeg,
                angle: angleDiff,
                value: percentage
            });
        }

        return segments;

    } catch (error) {
        return [];
    }
}

function getDefaultDoughnutData() {
    return {
        categories: ['Segment 1', 'Segment 2', 'Segment 3'],
        series: [{
            name: 'Values',
            values: [33, 33, 34],
            colors: ['#4472C4', '#ED7D31', '#A5A5A5']
        }],
        maxValue: 34,
        minValue: 0,
        isDoughnut: true
    };
}


/**
 * Extract chart position using the SAME method as tables
 * Direct px to inches conversion at 72 DPI
 */
function extractChartPosition(chartContainer, slideContext) {
    try {

        const styleAttr = chartContainer.getAttribute('style');

        if (!styleAttr) {
            return getDefaultPosition();
        }

        // Extract pixel values from style
        const styleValues = extractStyleValues(styleAttr);

        // Convert to inches using 72 DPI (SAME as table conversion)
        const xPos = parseToInches(styleValues.left || '0px');
        const yPos = parseToInches(styleValues.top || '0px');
        const width = parseToInches(styleValues.width || '360px');
        const height = parseToInches(styleValues.height || '345px');


        // Ensure minimum dimensions
        const finalWidth = width > 0 ? width : 3;
        const finalHeight = height > 0 ? height : 3;

        const result = {
            x: parseFloat(xPos.toFixed(3)),
            y: parseFloat(yPos.toFixed(3)),
            w: parseFloat(finalWidth.toFixed(3)),
            h: parseFloat(finalHeight.toFixed(3))
        };

        return result;

    } catch (error) {
        return getDefaultPosition();
    }
}

/**
 * Helper function to extract positioning from style attribute string
 * EXACT SAME as table code
 */
function extractStyleValues(styleAttr) {
    if (!styleAttr) return {};

    const values = {};
    const leftMatch = styleAttr.match(/left:\s*([^;]+)/i);
    const topMatch = styleAttr.match(/top:\s*([^;]+)/i);
    const widthMatch = styleAttr.match(/width:\s*([^;]+)/i);
    const heightMatch = styleAttr.match(/height:\s*([^;]+)/i);

    if (leftMatch?.[1]) values.left = leftMatch[1].trim();
    if (topMatch?.[1]) values.top = topMatch[1].trim();
    if (widthMatch?.[1]) values.width = widthMatch[1].trim();
    if (heightMatch?.[1]) values.height = heightMatch[1].trim();

    return values;
}

/**
 * Helper function to parse dimension value and convert to inches
 * EXACT SAME as table code
 */
function parseToInches(value) {
    if (!value) return 0;
    const numValue = parseFloat(value);
    if (value.includes('px')) {
        return numValue / 72; // Convert px to inches (72 DPI)
    } else if (value.includes('pt')) {
        return numValue / 72; // Convert pt to inches
    } else if (value.includes('in')) {
        return numValue; // Already in inches
    } else {
        return numValue / 72; // Default to px conversion
    }
}

/**
 * Get default fallback position
 */
function getDefaultPosition() {
    return {
        x: 1.0,
        y: 1.5,
        w: 5.0,
        h: 4.0
    };
}



function extractChartStyling(chartContainer, chartData) {
    const styling = {
        colors: [],
        fontSize: 12,
        fontFamily: 'Arial',
        showLegend: true,
        showGridLines: true,
        showDataLabels: false,
        yAxisConfig: null
    };

    try {
        // Extract Y-axis configuration
        styling.yAxisConfig = extractYAxisConfig(chartContainer);

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

        // Check for grid lines
        const axisLabels = chartContainer.querySelectorAll('.axis-label');
        styling.showGridLines = axisLabels.length > 0;

    } catch (error) {
    }

    return styling;
}

function convertToPptxFormat(chartData) {
    try {

        // Validate input data
        if (!chartData || !chartData.categories || !chartData.series) {
            return getDefaultChartData();
        }

        if (chartData.categories.length === 0 || chartData.series.length === 0) {
            return getDefaultChartData();
        }

        // Handle doughnut/pie charts
        if (chartData.isDoughnut || chartData.type === 'doughnut' || chartData.type === 'pie') {
            return convertDoughnutToPptxFormat(chartData);
        }

        const validSeries = chartData.series.filter(series => {
            return series && Array.isArray(series.values);
        });

        if (validSeries.length === 0) {
            return getDefaultChartData();
        }

        const pptxData = validSeries.map((series, idx) => {
            // Ensure name is a string
            let rawName = series.name;
            let safeName;
            if (typeof rawName === 'string') {
                safeName = rawName;
            } else if (rawName && typeof rawName === 'object') {
                safeName = rawName.name || rawName.label || rawName.id || `Series ${idx + 1}`;
            } else if (rawName != null) {
                safeName = String(rawName);
            } else {
                safeName = `Series ${idx + 1}`;
            }

            // Clean values
            const cleanValues = (series.values || []).map(value =>
                (typeof value === 'number' && !isNaN(value)) ? value : 0
            );

            // 🔥 FIX: Wrap categories in array for pptxgenjs format
            const labels = [chartData.categories.slice()];

            return {
                name: safeName,
                labels,
                values: cleanValues
            };
        });

        if (!validatePptxData(pptxData)) {
            return getDefaultChartData();
        }
        return pptxData;

    } catch (error) {
        return getDefaultChartData();
    }
}


/**
 * Convert doughnut/pie chart data to pptxgenjs format
 * For doughnut/pie charts, pptxgenjs needs ONE series with all values
 */
function convertDoughnutToPptxFormat(chartData) {
    try {
        const series = chartData.series[0];

        if (!series || !series.values || series.values.length === 0) {
            return [{
                name: 'Data',
                labels: ['Segment 1'],
                values: [100]
            }];
        }

        const pptxData = [{
            name: 'Chart Data',
            labels: chartData.categories,
            values: series.values
        }];

        pptxData.colors = series.colors;

        return pptxData;

    } catch (error) {
        return [{
            name: 'Data',
            labels: ['Segment 1'],
            values: [100]
        }];
    }
}


// Helper function to get default/fallback chart data in correct format
function getDefaultChartData() {
    return [
        {
            name: 'Series 1',
            labels: [['Category 1', 'Category 2', 'Category 3']],
            values: [1, 2, 3]
        }
    ];
}

// Helper function to validate PPTX data structure
function validatePptxData(data) {
    if (!Array.isArray(data) || data.length === 0) {
        return false;
    }

    for (let i = 0; i < data.length; i++) {
        const series = data[i];

        // Check required properties
        if (!series.name || typeof series.name !== 'string') {
            return false;
        }

        if (!series.labels || !Array.isArray(series.labels) || series.labels.length === 0) {
            return false;
        }

        if (!Array.isArray(series.labels[0])) {
            return false;
        }

        if (!series.values || !Array.isArray(series.values)) {
            return false;
        }

        // Check that values are numbers
        for (let j = 0; j < series.values.length; j++) {
            if (typeof series.values[j] !== 'number' || isNaN(series.values[j])) {
                return false;
            }
        }

        // Check that all series have the same label and value counts
        if (i > 0) {
            if (series.labels[0].length !== data[0].labels[0].length) {
                return false;
            }

            if (series.values.length !== data[0].values.length) {
                return false;
            }
        }
    }

    return true;
}
function extractYAxisConfig(chartContainer) {
    try {
        const rawLabels = [];

        // axis-label elements
        chartContainer.querySelectorAll('.axis-label').forEach(el => {
            const txt = (el.textContent || '').trim();
            if (txt) rawLabels.push(txt);
        });

        // fallback: any right-aligned text (used in some charts)
        if (rawLabels.length === 0) {
            chartContainer.querySelectorAll('[style*="text-align: right"]').forEach(el => {
                const txt = (el.textContent || '').trim();
                if (txt) rawLabels.push(txt);
            });
        }

        if (rawLabels.length < 2) return null;

        let hasPercent = false;
        let hasCurrency = false;
        let hasComma = false;
        let hasDecimal = false;

        rawLabels.forEach(t => {
            if (t.includes('%')) hasPercent = true;
            if (/[₹$£€]/.test(t)) hasCurrency = true;
            if (t.includes(',')) hasComma = true;
            if (/\d+\.\d+/.test(t)) hasDecimal = true;
        });

        // numeric values
        const numeric = [];
        rawLabels.forEach(t0 => {
            let t = t0.replace(/,/g, '').replace(/[$₹£€]/g, '').replace('%', '');
            const n = parseFloat(t);
            if (!isNaN(n)) numeric.push(n);
        });

        if (numeric.length < 2) return null;

        // If labels are 0%, 13%, 25%... numeric is [0,13,25,...]
        // For axis we want 0–0.5, so scale by 100 when percent
        let scaled = numeric.slice();
        if (hasPercent) {
            scaled = numeric.map(v => v / 100);
        }

        scaled.sort((a, b) => a - b);
        const min = scaled[0];
        const max = scaled[scaled.length - 1];

        const steps = [];
        for (let i = 1; i < scaled.length; i++) {
            steps.push(scaled[i] - scaled[i - 1]);
        }
        const interval = steps.length ? steps[0] : (max - min) / 5;

        // decide format
        let numberFormat = 'General';
        if (hasPercent) numberFormat = '0%';
        else if (hasCurrency) numberFormat = '$#,##0.00';
        else if (hasDecimal) numberFormat = '0.00';
        else if (hasComma || Math.max(...numeric) >= 1000) numberFormat = '#,##0';

        return {
            min,
            max,
            interval,
            isPercent: hasPercent,
            isCurrency: hasCurrency,
            hasDecimals: hasDecimal,
            hasThousands: hasComma || Math.max(...numeric) >= 1000,
            numberFormat
        };
    } catch {
        return null;
    }
}



function createChartOptions(chartData, position, styling, shapeName, pptxChartData) {
    const options = {
        x: position.x,
        y: position.y,
        w: position.w,
        h: position.h,
        title: chartData.title || 'Chart',
        showTitle: false,
        showLegend: false,
        objectName: shapeName || 'chart 0'
    };

    if (chartData.isDoughnut || chartData.type === 'doughnut') {
        // Doughnut chart options
        options.holeSize = 75;

        if (chartData.rotation !== undefined && chartData.rotation !== 0) {
            options.firstSliceAng = Math.round(chartData.rotation);
        }

        if (pptxChartData && pptxChartData.colors) {
            const pptxColors = pptxChartData.colors.map(color => {
                let hex = color.replace('#', '').toUpperCase();
                if (hex.length === 3) {
                    hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                }
                return hex;
            });
            options.chartColors = pptxColors;
        }

        options.dataBorder = {
            pt: 2,
            color: 'FFFFFF'
        };

    } else if (chartData.type === 'area') {
        // Area chart specific options
        options.showLegend = false;

        // Show every 2nd label on x-axis
        options.catAxisLabelFrequency = 2;
        options.catAxisLabelRotate = 0;

        // Set axis label font size
        options.catAxisLabelFontSize = 11;
        options.valAxisLabelFontSize = 11;

        // Grid lines
        options.catGridLine = { style: 'none' };
        options.valGridLine = { color: 'DFDFDF', style: 'solid' };

        // 🔥 FIX: Apply Y-axis configuration
        if (styling.yAxisConfig) {
            const yConfig = styling.yAxisConfig;


            // Set number format
            // Set number format (new logic)
            if (yConfig.numberFormat) {
                options.valAxisNumFmt = yConfig.numberFormat;
            } else if (yConfig.hasThousands) {
                options.valAxisNumFmt = '#,##0';
            } else if (yConfig.hasDecimals) {
                options.valAxisNumFmt = '0.00';
            } else {
                options.valAxisNumFmt = 'General';
            }


            // Set major unit (interval)
            options.valAxisMajorUnit = yConfig.interval;

            // Set min and max to ensure correct scale
            options.valAxisMinVal = yConfig.min;
            options.valAxisMaxVal = yConfig.max;
        }

        // Set colors
        if (styling.colors && Array.isArray(styling.colors)) {
            const validColors = styling.colors
                .map(color => normalizeColor(color, true))
                .filter(color => color && color.length === 6);
            if (validColors.length > 0) {
                options.chartColors = validColors;
            }
        }

        options.showCatAxisTitle = false;
        options.showValAxisTitle = false;

    } else if (chartData.type === 'bar' || chartData.type === 'column') {
        // BAR/COLUMN CHART SPECIFIC OPTIONS

        options.showLegend = chartData.series && chartData.series.length > 1;

        // Bar spacing settings
        options.barGrouping = 'clustered';
        options.barGapWidthPct = 150;
        options.barGapDepthPct = 25;

        // Set colors
        if (styling.colors && Array.isArray(styling.colors)) {
            const validColors = styling.colors
                .map(color => normalizeColor(color, true))
                .filter(color => color && color.length === 6);
            if (validColors.length > 0) {
                options.chartColors = validColors;
            }
        }

        if (options.showLegend) {
            options.legendPos = 'b';
            options.legendFontSize = 10;
        }

        if (styling.showGridLines) {
            options.valGridLine = { color: 'EEEEEE', style: 'solid' };
            options.catGridLine = { style: 'none' };
        }

        if (styling.yAxisConfig) {
            const yConfig = styling.yAxisConfig;


            // Set number format (new logic)
            if (yConfig.numberFormat) {
                options.valAxisNumFmt = yConfig.numberFormat;
            } else if (yConfig.hasThousands) {
                options.valAxisNumFmt = '#,##0';
            } else if (yConfig.hasDecimals) {
                options.valAxisNumFmt = '0.00';
            } else {
                options.valAxisNumFmt = 'General';
            }


            options.valAxisMajorUnit = yConfig.interval;

            // For bar charts, only set min if it's exactly 0
            if (yConfig.min === 0) {
                options.valAxisMinVal = 0;
            }
            // Set max value
            if (yConfig.max !== undefined) {
                options.valAxisMaxVal = yConfig.max;
            }
        }

        options.catAxisLabelFontSize = 11;
        options.valAxisLabelFontSize = 10;
        options.catAxisLabelRotate = 0;

        options.dataBorder = {
            pt: 0.5,
            color: 'FFFFFF'
        };

    } else {
        // Other chart types (line, etc.)
        options.showLegend = chartData.series && chartData.series.length > 1;

        if (styling.colors && Array.isArray(styling.colors)) {
            const validColors = styling.colors
                .map(color => normalizeColor(color, true))
                .filter(color => color && color.length === 6);
            if (validColors.length > 0) {
                options.chartColors = validColors;
            }
        }

        if (options.showLegend) {
            options.legendPos = 'b';
        }

        if (styling.showGridLines) {
            options.valGridLine = { color: 'E5E5E5' };
        }

        if (styling.yAxisConfig) {
            const yConfig = styling.yAxisConfig;

            // Set number format (new logic)
            if (yConfig.numberFormat) {
                options.valAxisNumFmt = yConfig.numberFormat;
            } else if (yConfig.hasThousands) {
                options.valAxisNumFmt = '#,##0';
            } else if (yConfig.hasDecimals) {
                options.valAxisNumFmt = '0.00';
            } else {
                options.valAxisNumFmt = 'General';
            }


            options.valAxisMajorUnit = yConfig.interval;
        }
    }

    // ---- FINAL: apply numeric format for value axis (all non-doughnut charts) ----
    if (chartData && styling && styling.yAxisConfig) {
        const y = styling.yAxisConfig;

        // Skip doughnut / pie – they don't use axes
        const chartType = (chartData.type || '').toLowerCase();
        const isDoughnut = chartType === 'doughnut' || chartType === 'pie';

        if (!isDoughnut) {
            if (y.isPercent) {
                options.valAxisNumFmt = '0%';
                options.valAxisLabelFormatCode = '0%';
            } else if (y.isCurrency) {
                options.valAxisNumFmt = '$#,##0.00';
                options.valAxisLabelFormatCode = '$#,##0.00';
            } else if (y.hasThousands) {
                options.valAxisNumFmt = '#,##0';
                options.valAxisLabelFormatCode = '#,##0';
            } else if (y.hasDecimals) {
                options.valAxisNumFmt = '0.00';
                options.valAxisLabelFormatCode = '0.00';
            } else if (y.numberFormat) {
                options.valAxisNumFmt = y.numberFormat;
                options.valAxisLabelFormatCode = y.numberFormat;
            }

            // Min always exact
            options.valAxisMinVal = y.min;

            // ⭐ exact max (no bump)
            options.valAxisMaxVal = y.max;

            // ⭐ calculate tick spacing without causing rounding
            if (typeof y.interval === 'number' && y.interval > 0) {
                const totalRange = y.max - y.min;
                const tickCount = Math.round(totalRange / y.interval);
                if (tickCount > 0) {
                    options.valAxisMajorUnit = totalRange / tickCount;
                } else {
                    options.valAxisMajorUnit = y.interval;
                }
            }
        }
    }


    return options;
}

function getChartType(chartType) {
    const typeMap = {
        'bar': 'bar',
        'column': 'bar',
        'line': 'line',
        'area': 'area',
        'pie': 'pie',
        'doughnut': 'doughnut',
        'donut': 'doughnut',
        'scatter': 'scatter'
    };

    const mapped = typeMap[chartType] || 'bar';
    return mapped;
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