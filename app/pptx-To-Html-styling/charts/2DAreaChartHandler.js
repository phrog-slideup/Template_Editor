class TwoDAreaChartHandler {
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
    /**
     * Main conversion method - converts area chart XML to HTML
     */
    async convertChartToHTML(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, chartZIndex) {
        try {
            
            // Extract chart position and dimensions
           const position = this.getChartPosition(graphicsNode);
            
            // Extract chart data
            const chartData = this.extractChartData(chartXML);
          
            // Extract styling information
            chartData.style = this.extractChartStyle(chartXML, chartColorsXML, chartStyleXML, themeXML);
            
            // Extract shape metadata
            const shapeName = graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name || "Area Chart";
            const shapeId = graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.id || "0";
            
            // Get relationship ID if available
            const graphicData = graphicsNode?.["a:graphic"]?.[0]?.["a:graphicData"]?.[0];
            const relationshipId = graphicData?.["c:chart"]?.[0]?.["$"]?.["r:id"] || "";
            
            // Generate HTML
            const html = this.generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, chartZIndex);
            
            return html;
        } catch (error) {
            return this.generateErrorHTML(error.message);
        }
    }

    /**
     * Extract chart dimensions and position from graphic frame
     */
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

    /**
     * Extract chart data (series, categories, values)
     */
    extractChartData(chartXML) {
        const space = chartXML.chartSpace || chartXML["c:chartSpace"];
        const chart = space?.chart?.[0] || space?.["c:chart"]?.[0];
        const plotArea = chart?.plotArea?.[0] || chart?.["c:plotArea"]?.[0];
        const areaChart = plotArea?.areaChart?.[0] || plotArea?.["c:areaChart"]?.[0];

        if (!areaChart) {
            throw new Error("Area chart data not found in XML");
        }

        const series = areaChart.ser || areaChart["c:ser"] || [];
        
        if (series.length === 0) {
            throw new Error("No series found in area chart");
        }
        
        const datasets = series.map(ser => this.extractSeriesData(ser));
        
        // Extract categories from first series
        const categories = datasets.length > 0 ? datasets[0].categories : [];

        // Calculate data min/max for axis scaling
        // Exclude zeros if they're just placeholders
        let dataMin = Infinity;
        let dataMax = -Infinity;
        let hasNonZeroValues = false;

        datasets.forEach((dataset, dsIdx) => {
            
            dataset.values.forEach((val, valIdx) => {
                if (!isNaN(val) && isFinite(val)) {
                    // Track if we have any non-zero values
                    if (val !== 0) {
                        hasNonZeroValues = true;
                    }
                    
                    // Include zero in range only if it's the only value OR we have negative values
                    if (val !== 0 || dataMin < 0) {
                        if (val < dataMin) dataMin = val;
                        if (val > dataMax) dataMax = val;
                    }
                }
            });
        });

        // If we only found zeros, include them in range
        if (!hasNonZeroValues && dataMin === Infinity) {
            dataMin = 0;
            dataMax = 0;
        }

        // Validate data min/max - set defaults if invalid
        if (!isFinite(dataMin) || !isFinite(dataMax) || dataMin === Infinity || dataMax === -Infinity) {
            dataMin = 0;
            dataMax = 100;
        }

        // Ensure max > min
        if (dataMax <= dataMin) {
            dataMax = dataMin + 100;
        }

        // Extract title
        const title = this.extractTitle(chart);

        // Extract axes configuration
        const valAx = plotArea?.valAx?.[0] || plotArea?.["c:valAx"]?.[0];
        const catAx = plotArea?.catAx?.[0] || plotArea?.["c:catAx"]?.[0];

        return {
            datasets,
            categories,
            grouping: this.extractGrouping(areaChart),
            title: title.display ? title.text : null,
            dataMin,
            dataMax,
            axes: {
                yAxis: this.extractAxisConfig(valAx, "value"),
                xAxis: this.extractAxisConfig(catAx, "category")
            }
        };
    }

    /**
     * Extract individual series data
     */
    extractSeriesData(seriesNode) {
        // Series name
        const txNode = seriesNode.tx?.[0] || seriesNode["c:tx"]?.[0];
        const seriesName = this.extractSeriesName(txNode);

        // Categories
        const catNode = seriesNode.cat?.[0] || seriesNode["c:cat"]?.[0];
        const categories = this.extractCategories(catNode);

        // Values
        const valNode = seriesNode.val?.[0] || seriesNode["c:val"]?.[0];
        const values = this.extractValues(valNode);

        // Series color
        const color = this.extractSeriesColor(seriesNode);

        return {
            label: seriesName,
            categories,
            values,
            color
        };
    }

    /**
     * Extract series name from tx node
     */
    extractSeriesName(txNode) {
        if (!txNode) return "Series";

        const strRef = txNode.strRef?.[0] || txNode["c:strRef"]?.[0];
        const strCache = strRef?.strCache?.[0] || strRef?.["c:strCache"]?.[0];
        const pt = strCache?.pt?.[0] || strCache?.["c:pt"]?.[0];
        const v = pt?.v?.[0] || pt?.["c:v"]?.[0];

        return v || "Series";
    }

    /**
     * Extract categories from cat node
     */
    extractCategories(catNode) {
        if (!catNode) return [];

        const strRef = catNode.strRef?.[0] || catNode["c:strRef"]?.[0];
        const strCache = strRef?.strCache?.[0] || strRef?.["c:strCache"]?.[0];
        const points = strCache?.pt || strCache?.["c:pt"] || [];

        return points.map(pt => {
            const v = pt.v?.[0] || pt["c:v"]?.[0];
            return v || "";
        });
    }

    /**
     * Extract values from val node
     */
    extractValues(valNode) {
        if (!valNode) {
            return [];
        }

        const numRef = valNode.numRef?.[0] || valNode["c:numRef"]?.[0];
        if (!numRef) {
            return [];
        }

        const numCache = numRef.numCache?.[0] || numRef["c:numCache"]?.[0];
        if (!numCache) {
            return [];
        }

        const points = numCache.pt || numCache["c:pt"] || [];

        if (points.length === 0) {
            return [];
        }

        const values = [];
        
        for (let idx = 0; idx < points.length; idx++) {
            const pt = points[idx];
            const v = pt.v?.[0] || pt["c:v"]?.[0];
            
            if (v === undefined || v === null) {
                values.push(0);
                continue;
            }
            
            const parsed = parseFloat(v);
            
            if (idx < 5) { // Log first 5 values for debugging
               // console.log(`  📍 Value ${idx}: "${v}" → ${parsed} ${isNaN(parsed) ? '❌ INVALID' : '✅'}`);
            }
            
            values.push(isNaN(parsed) ? 0 : parsed);
        }

        const validValues = values.filter(v => v > 0);
        return values;
    }

    /**
     * Extract series color from spPr node
     */
    extractSeriesColor(seriesNode) {
        const spPr = seriesNode.spPr?.[0] || seriesNode["c:spPr"]?.[0];
        
        if (!spPr) return "#4472C4"; // Default blue

        const solidFill = spPr["a:solidFill"]?.[0];
        if (!solidFill) return "#4472C4";

        const srgbClr = solidFill["a:srgbClr"]?.[0];
        if (srgbClr?.["$"]?.val) {
            return `#${srgbClr["$"].val}`;
        }

        return "#4472C4";
    }

    /**
     * Extract grouping type (standard, stacked, percentStacked)
     */
    extractGrouping(areaChart) {
        const grouping = areaChart.grouping?.[0] || areaChart["c:grouping"]?.[0];
        const val = grouping?.["$"]?.val || "standard";
        return val;
    }

    /**
     * Extract chart styling information
     */
    extractChartStyle(chartXML, chartColorsXML, chartStyleXML, themeXML) {
        const space = chartXML.chartSpace || chartXML["c:chartSpace"];
        const chart = space?.chart?.[0] || space?.["c:chart"]?.[0];
        const plotArea = chart?.plotArea?.[0] || chart?.["c:plotArea"]?.[0];

        const style = {
            backgroundColor: this.extractBackgroundColor(plotArea),
            gridlines: this.extractGridlines(plotArea),
            axes: this.extractAxesStyle(plotArea),
            legend: this.extractLegendStyle(chart),
            title: this.extractTitle(chart)
        };

        return style;
    }

    /**
     * Extract background color
     */
    extractBackgroundColor(plotArea) {
        const spPr = plotArea?.spPr?.[0] || plotArea?.["c:spPr"]?.[0];
        
        if (!spPr) return "transparent";

        const solidFill = spPr["a:solidFill"]?.[0];
        if (!solidFill) return "transparent";

        const srgbClr = solidFill["a:srgbClr"]?.[0];
        if (srgbClr?.["$"]?.val) {
            return `#${srgbClr["$"].val}`;
        }

        return "transparent";
    }

    /**
     * Extract gridlines configuration
     */
    extractGridlines(plotArea) {
        const valAx = plotArea?.valAx?.[0] || plotArea?.["c:valAx"]?.[0];
        const majorGridlines = valAx?.majorGridlines?.[0] || valAx?.["c:majorGridlines"]?.[0];

        if (!majorGridlines) {
            return { display: false, color: "#E0E0E0" };
        }

        const spPr = majorGridlines.spPr?.[0] || majorGridlines["c:spPr"]?.[0];
        const ln = spPr?.["a:ln"]?.[0];
        const solidFill = ln?.["a:solidFill"]?.[0];
        const srgbClr = solidFill?.["a:srgbClr"]?.[0];

        const color = srgbClr?.["$"]?.val ? `#${srgbClr["$"].val}` : "#DFDFDF";

        return {
            display: true,
            color: color
        };
    }

    /**
     * Extract axes configuration
     */
    extractAxesStyle(plotArea) {
        const catAx = plotArea?.catAx?.[0] || plotArea?.["c:catAx"]?.[0];
        const valAx = plotArea?.valAx?.[0] || plotArea?.["c:valAx"]?.[0];

        const axes = {
            x: this.extractAxisConfig(catAx, "category"),
            y: this.extractAxisConfig(valAx, "value")
        };

        return axes;
    }

    /**
     * Extract individual axis configuration
     */
    extractAxisConfig(axisNode, type) {
        if (!axisNode) {
          return { display: true, color: "#666666" };
        }

        const scaling = axisNode.scaling?.[0] || axisNode["c:scaling"]?.[0];
        
        let max = undefined;
        let min = undefined;
        
        if (scaling) {
            const maxNode = scaling.max?.[0] || scaling["c:max"]?.[0];
            const minNode = scaling.min?.[0] || scaling["c:min"]?.[0];
            
            if (maxNode?.["$"]?.val) {
                max = parseFloat(maxNode["$"].val);
            }
            
            if (minNode?.["$"]?.val) {
                min = parseFloat(minNode["$"].val);
            }
        }

        const numFmt = axisNode.numFmt?.[0] || axisNode["c:numFmt"]?.[0];
        const formatCode = numFmt?.["$"]?.formatCode || "General";

        const spPr = axisNode.spPr?.[0] || axisNode["c:spPr"]?.[0];
        const ln = spPr?.["a:ln"]?.[0];
        const solidFill = ln?.["a:solidFill"]?.[0];
        const srgbClr = solidFill?.["a:srgbClr"]?.[0];
        const color = srgbClr?.["$"]?.val ? `#${srgbClr["$"].val}` : "#8F9298";

        // Check if axis is deleted
        const deleteNode = axisNode.delete?.[0] || axisNode["c:delete"]?.[0];
        const isDeleted = deleteNode?.["$"]?.val === "1";

        const majorUnitNode = axisNode.majorUnit?.[0] || axisNode["c:majorUnit"]?.[0];
        const majorUnit = majorUnitNode?.["$"]?.val ? parseFloat(majorUnitNode["$"].val) : undefined;
        
        if (majorUnit) {
           // console.log(`  📏 ${type} axis MAJOR UNIT from XML: ${majorUnit}`);
        }

        const minorUnitNode = axisNode.minorUnit?.[0] || axisNode["c:minorUnit"]?.[0];
        const minorUnit = minorUnitNode?.["$"]?.val ? parseFloat(minorUnitNode["$"].val) : undefined;

        const config = {
            display: !isDeleted,
            color: color,
            max: max !== undefined && isFinite(max) ? max : undefined,
            min: min !== undefined && isFinite(min) ? min : undefined,
            formatCode: formatCode,
            majorUnit: majorUnit !== undefined && isFinite(majorUnit) ? majorUnit : undefined,
            minorUnit: minorUnit !== undefined && isFinite(minorUnit) ? minorUnit : undefined
        };
       return config;
    }

    /**
     * Extract legend configuration
     */
    extractLegendStyle(chart) {
        const legend = chart?.legend?.[0] || chart?.["c:legend"]?.[0];

        if (!legend) {
            return { display: false };
        }

        const legendPos = legend.legendPos?.[0] || legend["c:legendPos"]?.[0];
        const position = legendPos?.["$"]?.val || "r"; // r = right

        const positionMap = {
            "r": "right",
            "l": "left",
            "t": "top",
            "b": "bottom"
        };

        return {
            display: true,
            position: positionMap[position] || "right"
        };
    }

    /**
     * Extract chart title
     */
    extractTitle(chart) {
        const autoTitleDeleted = chart?.autoTitleDeleted?.[0] || chart?.["c:autoTitleDeleted"]?.[0];
        const isDeleted = autoTitleDeleted?.["$"]?.val === "1";

        if (isDeleted) {
            return { display: false, text: "" };
        }

        const title = chart?.title?.[0] || chart?.["c:title"]?.[0];
        
        if (!title) {
            return { display: false, text: "" };
        }

        // Extract title text
        const tx = title.tx?.[0] || title["c:tx"]?.[0];
        const rich = tx?.rich?.[0] || tx["c:rich"]?.[0];
        const p = rich?.["a:p"]?.[0];
        const r = p?.["a:r"]?.[0];
        const t = r?.["a:t"]?.[0];

        return {
            display: true,
            text: t || ""
        };
    }

    /**
     * Compute value axis ticks
     */
    computeValueTicks(yAxis, dataMin, dataMax, maxTicks = 10) {
     
        // Validate input data
        if (!isFinite(dataMin) || !isFinite(dataMax)) {
            dataMin = 0;
            dataMax = 100;
        }

        let min = yAxis?.min !== undefined ? yAxis.min : dataMin;
        let max = yAxis?.max !== undefined ? yAxis.max : dataMax;

        // Validate min/max
        if (!isFinite(min)) {
           min = 0;
        }
        if (!isFinite(max)) {
           max = 100;
        }
        
        // Ensure max > min
        if (max <= min) {
            max = min + 100;
        }

        // Add padding if not using XML values
        if (yAxis?.min === undefined || yAxis?.max === undefined) {
            const range = max - min;
            const oldMin = min;
            const oldMax = max;
            
            if (yAxis?.min === undefined) {
                min = Math.floor(min - range * 0.05);
            }
            if (yAxis?.max === undefined) {
                max = Math.ceil(max + range * 0.05);
            }
            
        }

        // Ensure min is not negative if all data is positive
        if (dataMin >= 0 && min < 0) {
            min = 0;
        }

        // Calculate nice step
        const rawStep = (max - min) / maxTicks;
        const magnitude = Math.pow(10, Math.floor(Math.log10(rawStep)));
        const residual = rawStep / magnitude;
        
        let step;
        if (residual > 5) step = 10 * magnitude;
        else if (residual > 2) step = 5 * magnitude;
        else if (residual > 1) step = 2 * magnitude;
        else step = magnitude;

        // Use majorUnit if specified
        if (yAxis?.majorUnit && isFinite(yAxis.majorUnit) && yAxis.majorUnit > 0) {
           step = yAxis.majorUnit;
        }

        // Validate step
        if (!isFinite(step) || step <= 0) {
           step = (max - min) / 5; // Default to 5 ticks
        }

        // Generate ticks
        const ticks = [];
        let current = Math.floor(min / step) * step;
        let iterations = 0;
        const maxIterations = 100; // Prevent infinite loops
        
        while (current <= max && iterations < maxIterations) {
            ticks.push(current);
            current += step;
            iterations++;
        }

        // Ensure we have at least 2 ticks
        if (ticks.length < 2) {
            ticks.length = 0;
            ticks.push(min);
            ticks.push(max);
        }

        return { min, max, step, ticks };
    }

    /**
     * Generate main chart HTML structure
     */
    generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, zIndex = 0) {
        // Create deterministic chartId using shapeId
        const chartId = "chart_" + shapeId;

        const { yAxis, dataMin, dataMax } = chartData.axes || {};
        const tickInfo = this.computeValueTicks(yAxis, dataMin, dataMax, 10);

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
                     font-family: Arial, sans-serif;
                     z-index: ${zIndex};  
                     cursor: pointer;
                     box-sizing: border-box;">`;

        // Title
        if (chartData.title) {
            html += `<div class="chart-title" style="
                        text-align: center;
                        font-weight: bold;
                        font-size: 16px;
                        margin-bottom: 20px;
                        color: #333;">
                            ${chartData.title}</div>`;
        }

        const chartAreaHeight = position.height - 120;
        const chartAreaWidth = position.width - 120;

        html += `<div class="chart-area" style="
                        position: relative;
                        height: ${chartAreaHeight}px;
                        width: ${chartAreaWidth}px;
                        margin-left: 60px;
                        margin-top: 20px;
                        background: ${chartData.style?.backgroundColor || 'transparent'};">`;

        // Generate SVG area paths
        html += this.generateAreaPaths(chartData, chartAreaWidth, chartAreaHeight, tickInfo);

        // Generate axes
        html += this.generateAxes(chartData, chartAreaWidth, chartAreaHeight, tickInfo);

        html += "</div>";

        // Legend
        if (chartData.style?.legend?.display) {
            html += this.generateLegend(chartData);
        }

        html += "</div>";

        return html;
    }

    /**
     * Generate SVG area paths for all series
     */
    generateAreaPaths(chartData, width, height, tickInfo) {
        const categories = chartData.categories;
        const datasets = chartData.datasets;
        
        if (!categories || categories.length === 0 || !datasets || datasets.length === 0) {
            return '';
        }

        // Validate tickInfo
        if (!tickInfo || !isFinite(tickInfo.min) || !isFinite(tickInfo.max)) {
            return '';
        }

        const xStep = width / (categories.length - 1);
        const valueRange = tickInfo.max - tickInfo.min;

        // Validate value range
        if (valueRange <= 0 || !isFinite(valueRange)) {
            return '';
        }

        let svg = `<svg style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none;">`;

        // Draw gridlines
        if (chartData.style?.gridlines?.display) {
            svg += this.generateGridlines(width, height, tickInfo, chartData.style.gridlines.color);
        }

        // Draw each series
        for (let seriesIdx = 0; seriesIdx < datasets.length; seriesIdx++) {
            const series = datasets[seriesIdx];
            const values = series.values;

            if (!values || values.length === 0) {
                continue;
            }

            // Generate path for area
            let pathData = '';
            
            // Start from bottom left
            const firstX = 0;
            const firstValue = values[0];
            
            // Validate first value
            if (!isFinite(firstValue)) {
                continue;
            }
            
            const firstY = height - ((firstValue - tickInfo.min) / valueRange) * height;
            
            // Validate Y coordinate
            if (!isFinite(firstY)) {
                continue;
            }
            
            pathData += `M ${firstX} ${height} `;
            pathData += `L ${firstX} ${firstY} `;

            // Draw line through points
            for (let i = 1; i < values.length; i++) {
                const x = i * xStep;
                const value = values[i];
                
                // Skip invalid values
                if (!isFinite(value)) {
                    continue;
                }
                
                const y = height - ((value - tickInfo.min) / valueRange) * height;
                
                // Validate Y coordinate
                if (!isFinite(y)) {
                    continue;
                }
                
                pathData += `L ${x} ${y} `;
            }

            // Close path to bottom right
            const lastX = (values.length - 1) * xStep;
            pathData += `L ${lastX} ${height} `;
            pathData += 'Z';

            // Create area with fill and stroke fill="${fillColor}"
            const fillColor = this.hexToRgba(series.color, 0.3);
            svg += `<path d="${pathData}" 
                         fill="${series.color}" 
                         stroke="${series.color}" 
                         stroke-width="2" 
                         stroke-linejoin="round"/>`;

            // Draw data points
            for (let i = 0; i < values.length; i++) {
                const x = i * xStep;
                const value = values[i];
                
                // Skip invalid values
                if (!isFinite(value)) continue;
                
                const y = height - ((value - tickInfo.min) / valueRange) * height;
                
                // Skip invalid Y coordinates
                if (!isFinite(y)) continue;
                
                svg += `<circle cx="${x}" cy="${y}" r="1" fill="${series.color}" stroke="white" stroke-width="0"/>`;
            }
        }

        svg += `</svg>`;
        return svg;
    }

    /**
     * Generate gridlines
     */
    generateGridlines(width, height, tickInfo, color) {
        let gridHtml = '';
        const { ticks, min, max } = tickInfo;
        const valueRange = max - min;

        for (const tick of ticks) {
            if (tick === min) continue; // Skip bottom line
            
            const y = height - ((tick - min) / valueRange) * height;
            gridHtml += `<line x1="0" y1="${y}" x2="${width}" y2="${y}" 
                              stroke="${color || '#DFDFDF'}" 
                              stroke-width="1" 
                              stroke-dasharray="none"/>`;
        }

        return gridHtml;
    }

    /**
     * Generate axes (X and Y)
     */
    generateAxes(chartData, width, height, tickInfo) {
        let html = '';

        // Validate inputs
        if (!chartData || !tickInfo) {
            return html;
        }

        // X-axis (categories)
        const categories = chartData.categories;
        if (!categories || categories.length === 0) {
        } else {
            const xStep = width / (categories.length - 1);

            // X-axis line
            const xAxisColor = chartData.style?.axes?.x?.color || '#8F9298';
            html += `<div style="position: absolute; bottom: -1px; left: 0; width: 100%; height: 2px; background: ${xAxisColor};"></div>`;

            // X-axis labels
            for (let i = 0; i < categories.length; i++) {
                // Skip some labels if too many
                const skipFactor = categories.length > 15 ? 2 : 1;
                if (i % skipFactor !== 0 && i !== categories.length - 1) continue;

                const x = i * xStep;
                html += `<div style="
                            position: absolute;
                            bottom: -25px;
                            left: ${x}px;
                            transform: translateX(-50%);
                            font-size: 11px;
                            color: ${xAxisColor};
                            white-space: nowrap;">
                                ${categories[i]}</div>`;
            }
        }

        // Y-axis line
        const yAxisColor = chartData.style?.axes?.y?.color || '#8F9298';
        html += `<div style="position: absolute; left: -1px; top: 0; width: 2px; height: 100%; background: ${yAxisColor};"></div>`;

        // Y-axis labels and ticks
        const { ticks, min, max } = tickInfo;
        
        if (!ticks || ticks.length === 0) {
            return html;
        }

        const valueRange = max - min;
        
        if (!isFinite(valueRange) || valueRange <= 0) {
            return html;
        }

        for (const tick of ticks) {
            if (!isFinite(tick)) {
                continue;
            }
            
            const y = height - ((tick - min) / valueRange) * height;
            
            if (!isFinite(y)) {
                continue;
            }
            
            // Tick mark
            html += `<div style="
                        position: absolute;
                        left: -5px;
                        top: ${y}px;
                        width: 5px;
                        height: 1px;
                        background: ${yAxisColor};"></div>`;

            // Label
            html += `<div style="
                        position: absolute;
                        right: ${width + 10}px;
                        top: ${y}px;
                        transform: translateY(-50%);
                        font-size: 11px;
                        color: ${yAxisColor};
                        text-align: right;">
                            ${this.formatNumber(tick)}</div>`;
        }

        return html;
    }

    /**
     * Generate legend
     */
    generateLegend(chartData) {
        const datasets = chartData.datasets;
        if (!datasets || datasets.length === 0) return '';

        let html = `<div class="chart-legend" style="
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        margin-top: 15px;
                        flex-wrap: wrap;
                        gap: 15px;">`;

        for (const series of datasets) {
            html += `<div style="display: flex; align-items: center; gap: 5px;">
                        <div style="width: 16px; height: 16px; background: ${series.color}; border-radius: 2px;"></div>
                        <span style="font-size: 12px; color: #333;">${series.label}</span>
                     </div>`;
        }

        html += `</div>`;
        return html;
    }

    /**
     * Format number with commas
     */
    formatNumber(num) {
        return num.toLocaleString();
    }

    /**
     * Convert hex color to rgba
     */
    hexToRgba(hex, alpha = 1) {
        // Remove # if present
        hex = hex.replace('#', '');
        
        // Parse hex values
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    /**
     * Generate error HTML when conversion fails
     */
    generateErrorHTML(errorMessage) {
        return `
<div class="chart-error" style="position: absolute; padding: 20px; background: #ffebee; border: 1px solid #f44336; border-radius: 4px;">
    <p style="color: #d32f2f; margin: 0; font-family: Arial, sans-serif;">
        <strong>Error:</strong> Failed to convert area chart<br>
        <small>${errorMessage}</small>
    </p>
</div>
        `.trim();
    }
}

module.exports = TwoDAreaChartHandler;