const colorHelper = require("../../api/helper/colorHelper.js");

class BarChartHandler {

    constructor(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML) {
        this.graphicsNode = graphicsNode;
        this.chartXML = chartXML;
        this.chartRelsXML = chartRelsXML;
        this.chartColorsXML = chartColorsXML;
        this.chartStyleXML = chartStyleXML;
        this.themeXML = themeXML;
    }

    // EMU -> px divisor
    getEMUDivisor() {
        return 12700;
    }

    async convertChartToHTML(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, zIndex = 0) {
        try {
            const sanitizeShapeName = (name) => {
                if (!name) return '';
                return name.replace(/['"<>&]/g, '').trim() || '';
            };

            // Extract unique identifiers from graphicsNode
            const cNvPr = graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0];
            const shapeName = cNvPr?.["$"]?.name || 'Chart';
            const shapeId = cNvPr?.["$"]?.id || Math.random().toString(36).substr(2, 9);

            // Extract relationship ID (r:id) from the chart reference
            const chartRef = graphicsNode?.["a:graphic"]?.[0]?.["a:graphicData"]?.[0]?.["c:chart"]?.[0];
            const relationshipId = chartRef?.["$"]?.["r:id"] || '';

            const position = this.getChartPosition(graphicsNode);
            const chartData = this.parseChartData(chartXML);

            // If no valid chart data found, try to extract from alternative sources or create fallback
            if (!chartData.series || chartData.series.length === 0) {
                const fallbackData = this.createFallbackChartData(chartXML, shapeName);
                const htmlOutput = this.generateChartHTML(fallbackData, position, shapeName, shapeId, relationshipId, zIndex);
                return htmlOutput;
            }

            const htmlOutput = this.generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, zIndex);
            return htmlOutput;
        } catch (error) {
            return this.generateErrorHTML(error.message);
        }
    }

    createFallbackChartData(chartXML, shapeName) {

        // Try to extract any available data or create sample data
        const chart = chartXML["c:chartSpace"]?.["c:chart"]?.[0];
        const title = this.getChartTitle(chart) || shapeName || "Chart";

        // Create sample data structure
        const series = [
            {
                index: 0,
                name: "Series 1",
                values: [4.3, 2.5, 3.5, 4.5],
                categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
                color: "#A5C249"
            },
            {
                index: 1,
                name: "Series 2",
                values: [2.4, 4.4, 1.8, 2.8],
                categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
                color: "#7CCA62"
            },
            {
                index: 2,
                name: "Series 3",
                values: [2, 2, 3, 5],
                categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
                color: "#10CF9B"
            }
        ];

        const allValues = series.flatMap(s => s.values);

        return {
            type: "bar",
            isHorizontal: false,
            title: title,
            series: series,
            categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
            axes: {
                yAxis: null,
                dataMin: 1.8,
                dataMax: 5,
                allValues: allValues
            },
        };
    }

    parseChartData(chartXML, axisPreferences = null) {
        try {
            const chart = chartXML["c:chartSpace"]?.["c:chart"]?.[0];
            if (!chart) {
                return null;
            }

            const plotArea = chart["c:plotArea"]?.[0];
            if (!plotArea) {
                return null;
            }

            const barChart = plotArea["c:barChart"]?.[0] ||
                plotArea["c:bar3DChart"]?.[0] ||
                plotArea["c:columnChart"]?.[0] ||
                plotArea["c:col3DChart"]?.[0];

            if (!barChart) {
                return null;
            }

            const barDir = barChart["c:barDir"]?.[0]?.["$"]?.["val"] || "col";
            const isHorizontal = barDir === "bar";

            const series = barChart["c:ser"] || [];

            if (series.length === 0) {
                return null;
            }

            const chartSeries = series.map((ser, index) => {
                const seriesData = {
                    index,
                    name: this.getSeriesName(ser, index),
                    values: this.getSeriesValues(ser),
                    categories: this.getSeriesCategories(ser),
                    color: this.getSeriesColor(ser, index),
                };
                return seriesData;
            });

            const title = this.getChartTitle(chart);

            // Get axis info and data extents
            let yAxis = this.getValueAxis(plotArea);

            // Apply axis preferences if provided (allows manual override)
            if (axisPreferences) {
                yAxis = {
                    ...yAxis,
                    ...axisPreferences
                };
            }

            const allValues = chartSeries.flatMap((s) => s.values).filter(v => v !== null && v !== undefined && !isNaN(v));
            const dataMin = allValues.length ? Math.min(...allValues) : 0;
            const dataMax = allValues.length ? Math.max(...allValues) : 5;

            const result = {
                type: "bar",
                isHorizontal,
                title,
                series: chartSeries,
                categories: chartSeries.length ? chartSeries[0].categories : [],
                axes: { yAxis, dataMin, dataMax, allValues }, // Include allValues for analysis
            };


            return result;
        } catch (error) {
            return null;
        }
    }

    // Helper function to safely extract text from XML value
    extractTextValue(value) {
        if (!value) return null;

        // If it's already a string, return it
        if (typeof value === 'string') return value;

        // If it's an object with _ property (xml2js text node with attributes)
        if (typeof value === 'object' && value._) return String(value._);

        // If it's an object with text property
        if (typeof value === 'object' && value.text) return String(value.text);

        // Try to convert to string
        if (value !== null && value !== undefined) {
            const str = String(value);
            // Don't return [object Object]
            if (str !== '[object Object]') return str;
        }

        return null;
    }

    getSeriesName(seriesXML, index) {
        try {
            const tx = seriesXML["c:tx"];

            if (tx && tx[0]) {
                // Try string reference first
                if (tx[0]["c:strRef"]) {
                    const strCache = tx[0]["c:strRef"][0]["c:strCache"];
                    if (strCache && strCache[0]["c:pt"] && strCache[0]["c:pt"][0]) {
                        const value = strCache[0]["c:pt"][0]["c:v"];
                        if (value && value[0]) {
                            const extracted = this.extractTextValue(value[0]);
                            if (extracted) return extracted;
                        }
                    }
                }

                // Try direct value
                if (tx[0]["c:v"] && tx[0]["c:v"][0]) {
                    const extracted = this.extractTextValue(tx[0]["c:v"][0]);
                    if (extracted) return extracted;
                }

                // Try rich text
                if (tx[0]["c:rich"]) {
                    const rich = tx[0]["c:rich"][0];

                    if (rich["a:p"]) {
                        for (const p of rich["a:p"]) {
                            // Try a:r -> a:t path
                            if (p["a:r"]) {
                                for (const r of p["a:r"]) {
                                    if (r["a:t"] && r["a:t"][0]) {
                                        const extracted = this.extractTextValue(r["a:t"][0]);
                                        if (extracted) return extracted;
                                    }
                                }
                            }

                            // Try direct a:t
                            if (p["a:t"] && p["a:t"][0]) {
                                const extracted = this.extractTextValue(p["a:t"][0]);
                                if (extracted) return extracted;
                            }
                        }
                    }
                }
            }

            // Fallback to index from series
            const idxVal = seriesXML["c:idx"]?.[0]?.["$"]?.["val"];
            if (idxVal !== undefined) {
                return `Series ${parseInt(idxVal) + 1}`;
            }
            return `Series ${index + 1}`;
        } catch (error) {
            return `Series ${index + 1}`;
        }
    }

    getSeriesValues(seriesXML) {
        try {
            const val = seriesXML["c:val"];
            if (val && val[0]) {
                if (val[0]["c:numRef"]) {
                    const numCache = val[0]["c:numRef"][0]["c:numCache"];
                    if (numCache && numCache[0]["c:pt"]) {
                        const values = numCache[0]["c:pt"].map((pt) => {
                            const v = pt["c:v"]?.[0];
                            return v !== undefined ? parseFloat(v) : 0;
                        }).filter(v => !isNaN(v));
                        return values;
                    }
                }
                // Try literal values
                if (val[0]["c:numLit"]) {
                    const numLit = val[0]["c:numLit"][0];
                    if (numLit["c:pt"]) {
                        return numLit["c:pt"].map((pt) => {
                            const v = pt["c:v"]?.[0];
                            return v !== undefined ? parseFloat(v) : 0;
                        }).filter(v => !isNaN(v));
                    }
                }
            }
            return [];
        } catch (error) {
            return [];
        }
    }

    getSeriesCategories(seriesXML) {
        try {
            const cat = seriesXML["c:cat"];
            if (cat && cat[0]) {
                // Try string reference
                if (cat[0]["c:strRef"]) {
                    const strCache = cat[0]["c:strRef"][0]["c:strCache"];
                    if (strCache && strCache[0]["c:pt"]) {
                        const categories = strCache[0]["c:pt"].map((pt) => pt["c:v"]?.[0] || "");
                        return categories;
                    }
                }
                // Try multi-level string reference
                if (cat[0]["c:multiLvlStrRef"]) {
                    const multiCache = cat[0]["c:multiLvlStrRef"][0]["c:multiLvlStrCache"];
                    if (multiCache && multiCache[0]["c:lvl"] && multiCache[0]["c:lvl"][0]["c:pt"]) {
                        return multiCache[0]["c:lvl"][0]["c:pt"].map((pt) => pt["c:v"]?.[0] || "");
                    }
                }
                // Try literal categories
                if (cat[0]["c:strLit"]) {
                    const strLit = cat[0]["c:strLit"][0];
                    if (strLit["c:pt"]) {
                        return strLit["c:pt"].map((pt) => pt["c:v"]?.[0] || "");
                    }
                }
            }
            return [];
        } catch (error) {
            return [];
        }
    }

    getValueAxis(plotArea) {
        const valAx = (plotArea["c:valAx"] || [])[0];
        if (!valAx) {
            return null;
        }

        const scaling = valAx["c:scaling"]?.[0];
        const numFmt = valAx["c:numFmt"]?.[0]?.["$"]?.formatCode || "General";
        const crosses = valAx["c:crosses"]?.[0]?.["$"]?.val || "autoZero";

        const toNum = (n) => {
            const v = Number(n);
            return Number.isFinite(v) ? v : undefined;
        };

        // Extract min value
        let minVal = toNum(scaling?.["c:min"]?.[0]?.["$"]?.val);

        // Extract max value
        let maxVal = toNum(scaling?.["c:max"]?.[0]?.["$"]?.val);

        // Extract majorUnit
        let majorUnitVal = toNum(valAx["c:majorUnit"]?.[0]?.["$"]?.val);

        // Check if orientation is reversed
        const orientation = scaling?.["c:orientation"]?.[0]?.["$"]?.val;

        // Check for auto settings
        const autoMin = valAx["c:auto"]?.[0]?.["$"]?.val === "1";
        const autoMax = valAx["c:autoMax"]?.[0]?.["$"]?.val === "1";

        // NEW: Check for explicitly stored tick labels (sometimes PowerPoint stores these)
        const tickLblSkip = valAx["c:tickLblSkip"]?.[0]?.["$"]?.val;
        const tickMarkSkip = valAx["c:tickMarkSkip"]?.[0]?.["$"]?.val;

        const axisData = {
            min: minVal,
            max: maxVal,
            majorUnit: majorUnitVal,
            minorUnit: toNum(valAx["c:minorUnit"]?.[0]?.["$"]?.val),
            numFmt,
            crosses,
            orientation,
            autoMin,
            autoMax,
            tickLblSkip,
            tickMarkSkip
        };

        return axisData;
    }

    niceStep(rawStep) {
        if (!Number.isFinite(rawStep) || rawStep <= 0) return 1;
        const pow = Math.pow(10, Math.floor(Math.log10(rawStep)));
        const frac = rawStep / pow;
        let niceFrac;
        if (frac <= 1) niceFrac = 1;
        else if (frac <= 2) niceFrac = 2;
        else if (frac <= 2.5) niceFrac = 2.5;
        else if (frac <= 5) niceFrac = 5;
        else niceFrac = 10;
        return niceFrac * pow;
    }

    computeValueTicks(yAxis, dataMin, dataMax, desired = 10, allDataValues = []) {

        // Use exact min/max from PPTX if available, otherwise calculate from data
        let min = yAxis?.min !== undefined ? yAxis.min :
            yAxis?.crosses === "autoZero" && dataMin > 0 ? 0 : dataMin;
        let max = yAxis?.max !== undefined ? yAxis.max : dataMax;

        if (min === max) {
            if (max === 0) max = 1;
            else min = 0;
        }

        // Use exact majorUnit from PPTX if available
        let step;
        const hasExplicitMajorUnit = yAxis?.majorUnit !== undefined && yAxis.majorUnit > 0;

        if (hasExplicitMajorUnit) {
            // Use the exact step from PPTX without any modification
            step = yAxis.majorUnit;
        } else {
            // No explicit value in PPTX - calculate using niceStep
            const rawStep = (max - min) / desired;
            step = this.niceStep(rawStep);


            // SMART DETECTION: Analyze data to decide between integer and decimal steps
            if (!Number.isInteger(step) && allDataValues.length > 0) {
                const detection = this.detectStepPreference(allDataValues, step, max - min);

                if (detection.preferInteger) {
                    step = detection.suggestedStep;
                } else {
                    // console.log("🎯 Smart detection: Keeping decimal step", step, `(${detection.reason})`);
                }
            }
        }

        // Use exact min/max if provided in PPTX
        const hasExactMinMax = yAxis?.min !== undefined && yAxis?.max !== undefined;

        const originalMin = min;
        const originalMax = max;

        if (!hasExactMinMax) {
            // Round min/max to align with step for clean axis
            min = Math.floor(min / step) * step;
            max = Math.ceil(max / step) * step;

        } else {
            // console.log("✅ Using exact min/max from PPTX:", { min, max });
        }

        const span = max - min || step;
        const decimals = this.decimalsFor(step);
        const ticks = [];

        // Generate ticks from min to max using the step
        for (let v = min; v <= max + step * 1e-6; v += step) {
            ticks.push(+v.toFixed(decimals));
        }

        return { ticks, min, max, step, numFmt: yAxis?.numFmt || "General", span };
    }

    detectStepPreference(allDataValues, calculatedStep, range) {
        // Analyze data values to determine if integer or decimal steps are more appropriate

        // Count how many values have significant decimal parts
        const decimalThreshold = 0.15; // Values with decimals > 0.15 are considered "truly decimal"
        const valuesWithDecimals = allDataValues.filter(v => {
            const decimalPart = Math.abs(v - Math.round(v));
            return decimalPart > decimalThreshold;
        });

        const percentDecimal = valuesWithDecimals.length / allDataValues.length;

        // Decision logic - MORE AGGRESSIVE for integer steps

        // For calculated step of 0.5 with small ranges (≤ 6)
        if (calculatedStep === 0.5 && range <= 6) {
            // Use integer steps unless MOST values (>75%) have significant decimals
            // This means we prefer integer steps for mixed data
            if (percentDecimal < 0.75) {
                return {
                    preferInteger: true,
                    suggestedStep: 1,
                    reason: `Only ${(percentDecimal * 100).toFixed(0)}% of values have significant decimals (threshold: 75%), range is ${range}`
                };
            } else {
                return {
                    preferInteger: false,
                    suggestedStep: calculatedStep,
                    reason: `${(percentDecimal * 100).toFixed(0)}% of values have significant decimals (above 75% threshold)`
                };
            }
        }

        // For other decimal steps with small ranges
        if (!Number.isInteger(calculatedStep) && range <= 10) {
            // If less than 50% of values have decimals, prefer integer steps
            if (percentDecimal < 0.5) {
                const intStep = Math.max(1, Math.round(calculatedStep));
                const numTicks = Math.ceil(range / intStep) + 1;

                // Only use integer step if it gives reasonable tick count
                if (numTicks >= 4 && numTicks <= 15) {
                    return {
                        preferInteger: true,
                        suggestedStep: intStep,
                        reason: `${(percentDecimal * 100).toFixed(0)}% of values have decimals (below 50%), would give ${numTicks} ticks`
                    };
                }
            }
        }

        // Default: keep the calculated step
        return {
            preferInteger: false,
            suggestedStep: calculatedStep,
            reason: "Data characteristics match calculated step"
        };
    }

    decimalsFor(num) {
        const s = String(num);
        return s.includes(".") ? s.split(".")[1].length : 0;
    }

    formatTick(val, numFmt, step) {
        // Check if this is a percentage format
        const isPercentage = numFmt && (numFmt.includes('%') || numFmt.toLowerCase().includes('percent'));

        if (isPercentage) {
            // For percentage format, multiply by 100 and add %
            const percentValue = val * 100;

            // Parse the format code to determine decimal places
            let decimals = 0;

            // Check the format string for decimal places
            // Examples: "#,##0%" = 0 decimals, "#,##0.0%" = 1 decimal, "0.00%" = 2 decimals
            const decimalMatch = numFmt.match(/0\.(0+)%/);
            if (decimalMatch) {
                // Format specifies decimals (e.g., "0.00%" means 2 decimals)
                decimals = decimalMatch[1].length;
            } else if (numFmt.includes('.')) {
                // Has decimal point but didn't match pattern, check manually
                const parts = numFmt.split('.');
                if (parts.length > 1) {
                    const afterDecimal = parts[1].replace(/[^0]/g, '');
                    decimals = afterDecimal.length;
                }
            } else {
                // No decimal point in format = 0 decimals (e.g., "#,##0%", "0%")
                decimals = 0;
            }

            // Format the percentage value
            let formatted;
            if (decimals === 0) {
                // For whole numbers, just round - don't use replace to avoid removing the "0" from "0"
                formatted = Math.round(percentValue).toString();
            } else {
                // For decimals, format and remove unnecessary trailing zeros after decimal point
                formatted = percentValue.toFixed(decimals).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
            }

            return formatted + '%';
        }

        // For non-percentage formats
        const fixed = this.decimalsFor(step);
        if (numFmt && /^0(\.0+)?$/.test(numFmt)) {
            const dec = (numFmt.split(".")[1] || "").length;
            return Number(val).toFixed(dec);
        }
        return fixed ? Number(val).toFixed(fixed).replace(/\.?0+$/, "") : String(val);
    }

    getSeriesColor(seriesXML, index) {
        try {
            const spPr = seriesXML["c:spPr"];
            if (spPr && spPr[0]) {
                const solidFill = spPr[0]["a:solidFill"];
                if (solidFill && solidFill[0]) {
                    const schemeClr = solidFill[0]["a:schemeClr"];
                    if (schemeClr && schemeClr[0]) {
                        const colorName = schemeClr[0]["$"]["val"];
                        return this.resolveThemeColor(colorName);
                    }
                    const srgbClr = solidFill[0]["a:srgbClr"];
                    if (srgbClr && srgbClr[0]) {
                        return "#" + srgbClr[0]["$"]["val"];
                    }
                }
            }

            const defaultColors = [
                "#A5C249", "#7CCA62", "#10CF9B", "#5B9BD5", "#70AD47",
                "#FFC000", "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4"
            ];
            return defaultColors[index % defaultColors.length];
        } catch {
            const fallback = ["#A5C249", "#7CCA62", "#10CF9B"];
            return fallback[index % fallback.length];
        }
    }

    resolveThemeColor(colorName) {
        const themeColors = {
            accent1: "#5B9BD5", accent2: "#70AD47", accent3: "#FFC000",
            accent4: "#10CF9B", accent5: "#7CCA62", accent6: "#A5C249",
        };
        return themeColors[colorName] || "#5B9BD5";
    }

    getChartTitle(chartXML) {
        try {
            const title = chartXML["c:title"];
            if (title && title[0] && title[0]["c:tx"] && title[0]["c:tx"][0]["c:rich"]) {
                const rich = title[0]["c:tx"][0]["c:rich"][0];
                if (rich["a:p"] && rich["a:p"][0] && rich["a:p"][0]["a:r"]) {
                    return rich["a:p"][0]["a:r"][0]["a:t"][0];
                }
            }
            return "";
        } catch {
            return "";
        }
    }

    getChartPosition(graphicsNode) {
        const emuDivisor = this.getEMUDivisor();
        const chartXfrm = graphicsNode?.["p:xfrm"]?.[0];
        if (chartXfrm) {
            const offset = chartXfrm?.["a:off"]?.[0]?.["$"];
            const extent = chartXfrm?.["a:ext"]?.[0]?.["$"];
            if (offset && extent) {
                const left = parseInt(offset.x || 0) / emuDivisor;
                const top = parseInt(offset.y || 0) / emuDivisor;
                const width = parseInt(extent.cx || 100) / emuDivisor;
                const height = parseInt(extent.cy || 100) / emuDivisor;
                return { left, top, width, height };
            }
        }
        return { left: 0, top: 0, width: 400, height: 300 };
    }

    generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, zIndex = 0) {
        // Create deterministic chartId using shapeId
        const chartId = "chart_" + shapeId;

        const { yAxis, dataMin, dataMax, allValues = [] } = chartData.axes || {};
        const tickInfo = this.computeValueTicks(yAxis, dataMin, dataMax, 10, allValues);

        let html = `<div class="chart-container" 
                 data-name="${shapeName}" 
                 data-shape-id="${shapeId}"
                 data-rel-id="${relationshipId}"
                 id="${chartId}" 
                 style="
                     position: absolute; 
                     top: ${position.top}px;
                     left: ${position.left}px;
                     width: ${position.width}px;
                     height: ${position.height}px;
                     padding: 20px;
                     border: 1px solid #ddd;
                     border-radius: 4px;
                     background: white;
                     font-family: Arial, sans-serif;
                     z-index: ${zIndex};  
                     cursor: pointer;
                     box-sizing: border-box;">`;

        if (chartData.title) {
            html += `<div class="chart-title" style="
                        text-align: center;
                        font-weight: bold;
                        font-size: 16px;
                        margin-bottom: 20px;
                        color: #333;">
                            ${chartData.title}</div>`;
        }

        const chartAreaHeight = position.height - 100;
        const chartAreaWidth = position.width - 100;

        html += `<div class="chart-area" style="
                        position: relative;
                        height: ${chartAreaHeight}px;
                        width: ${chartAreaWidth}px;
                        margin-left: 60px;
                        margin-top: 20px;">`;

        if (chartData.isHorizontal) {
            html += this.generateHorizontalBars(chartData, chartAreaWidth, chartAreaHeight, tickInfo);
        } else {
            html += this.generateVerticalBars(chartData, chartAreaWidth, chartAreaHeight, tickInfo);
        }

        html += this.generateAxes(chartData, chartAreaWidth, chartAreaHeight, tickInfo);
        html += "</div>";
        html += "</div>";

        return html;
    }

    generateVerticalBars(chartData, width, height, tickInfo) {
        let html = "";
        const categoryCount = chartData.categories.length || 1;
        const seriesCount = chartData.series.length || 1;
        const categoryWidth = width / categoryCount;
        const barWidth = Math.round((categoryWidth * 0.8) / seriesCount);
        const barSpacing = categoryWidth * 0.1;

        const min = tickInfo.min;
        const span = tickInfo.span || tickInfo.max - tickInfo.min || 1;
        const innerH = height - 40;

        const valueToY = (v) => height - 20 - ((v - min) / span) * innerH;
        const zeroY = valueToY(0);

        chartData.categories.forEach((category, categoryIndex) => {
            chartData.series.forEach((series, seriesIndex) => {
                const value = series.values[categoryIndex] ?? 0;
                const yVal = valueToY(value);
                const top = Math.round(Math.min(yVal, zeroY));
                const barH = Math.round(Math.max(1, Math.abs(yVal - zeroY)));
                const x = Math.round(categoryIndex * categoryWidth + barSpacing + seriesIndex * barWidth);

                const barId = `bar_${seriesIndex}_${categoryIndex}`;

                html += `<div class="bar" 
                id="${barId}"
                data-series="${seriesIndex}" 
                data-category="${categoryIndex}"
                data-value="${value}"
                style="
                    position: absolute;
                    left: ${x}px;
                    top: ${top}px;
                    width: ${barWidth - 2}px;
                    height: ${barH}px;
                    background-color: ${series.color};
                    border-radius: 2px;
                    transition: opacity 0.3s;
                    z-index: 2;
                    cursor: ns-resize;
                    user-select: none;
                " title="${series.name}: ${value}"></div>`;
            });

            html += `<div class="category-label" style="
                        position: absolute;
                        left: ${categoryIndex * categoryWidth + categoryWidth / 2 - 30}px;
                        top: ${height - 15}px;
                        width: 60px;
                        text-align: center;
                        font-size: 12px;
                        color: #666;
                        z-index: 4;
                    ">${category}</div>`;
        });

        return html;
    }

    generateHorizontalBars(chartData, width, height, tickInfo) {
        let html = "";
        const categoryCount = chartData.categories.length || 1;
        const seriesCount = chartData.series.length || 1;
        const categoryHeight = height / categoryCount;
        const barHeight = Math.round((categoryHeight * 0.8) / seriesCount);
        const barSpacing = categoryHeight * 0.1;

        const min = tickInfo.min;
        const span = tickInfo.span || tickInfo.max - tickInfo.min || 1;
        const innerW = width - 60;

        const valueToX = (v) => 40 + ((v - min) / span) * innerW;
        const zeroX = valueToX(0);

        chartData.categories.forEach((category, categoryIndex) => {
            chartData.series.forEach((series, seriesIndex) => {
                const value = series.values[categoryIndex] ?? 0;
                const xVal = valueToX(value);
                const left = Math.round(Math.min(xVal, zeroX));
                const bWidth = Math.round(Math.max(1, Math.abs(xVal - zeroX)));
                const y = Math.round(categoryIndex * categoryHeight + barSpacing + seriesIndex * barHeight);

                html += `<div class="bar" style="
                                position: absolute;
                                left: ${left}px;
                                top: ${y}px;
                                width: ${bWidth}px;
                                height: ${barHeight - 2}px;
                                background-color: ${series.color};
                                border-radius: 2px;
                                transition: opacity 0.3s;
                                z-index: 2;
                                " title="${series.name}: ${value}"></div>`;
            });

            html += `<div class="category-label" style="
                                position: absolute;
                                left: 5px;
                                top: ${categoryIndex * categoryHeight + categoryHeight / 2 - 8}px;
                                width: 35px;
                                text-align: right;
                                font-size: 12px;
                                color: #666;
                                z-index: 4;
                            ">${category}</div>`;
        });

        return html;
    }

    generateAxes(chartData, width, height, tickInfo) {
        let html = "";
        const { ticks, min, step, numFmt, span } = tickInfo;
        const innerH = height - 40;
        const innerW = width - 60;

        if (chartData.isHorizontal) {
            ticks.forEach((t) => {
                const x = 40 + ((t - min) / span) * innerW;
                html += `<div style="position:absolute; left:${x}px; top:0; width:1px; height:${height - 20}px; background:#eee; z-index:1; pointer-events:none;"></div>`;
                html += `<div class="axis-label" style="position:absolute; left:${x - 15}px; top:${height + 5}px; width:30px; text-align:center; font-size:10px; color:#666; z-index:4;">${this.formatTick(t, numFmt, step)}</div>`;
            });
        } else {
            ticks.forEach((t) => {
                const y = height - 20 - ((t - min) / span) * innerH;
                html += `<div style="position:absolute; left:0; top:${y}px; width:${width}px; height:1px; background:#eee; z-index:1; pointer-events:none;"></div>`;
                html += `<div class="axis-label" style="position:absolute; left:-35px; top:${y - 8}px; width:30px; text-align:right; font-size:10px; color:#666; z-index:4;">${this.formatTick(t, numFmt, step)}</div>`;
            });
        }

        return html;
    }

    generateErrorHTML(errorMessage) {
        return `<div class="chart-error" style="
                    width: 400px;
                    height: 300px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    border: 2px dashed #ccc;
                    color: #666;
                    font-family: Arial, sans-serif;">
            <div style="text-align: center;">
                <div style="font-size: 48px; margin-bottom: 10px;">📊</div>
                <div>Chart could not be rendered</div>
                <div style="font-size: 12px; margin-top: 10px; color: #999;">${errorMessage}</div>
            </div>
        </div>`;
    }
}

module.exports = BarChartHandler;