const colorHelper = require("../../api/helper/colorHelper.js");

class BarChartHandler {

    constructor(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, masterXML = null) {
        this.graphicsNode = graphicsNode;
        this.chartXML = chartXML;
        this.chartRelsXML = chartRelsXML;
        this.chartColorsXML = chartColorsXML;
        this.chartStyleXML = chartStyleXML;
        this.themeXML = themeXML;
        this.masterXML = masterXML;
    }

    getEMUDivisor() {
        return 12700;
    }

    escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    async convertChartToHTML(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, zIndex = 0) {
        try {
            const cNvPr = graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0];
            const shapeName = cNvPr?.["$"]?.name || "Chart";
            const shapeId = cNvPr?.["$"]?.id || Math.random().toString(36).substr(2, 9);
            const chartRef = graphicsNode?.["a:graphic"]?.[0]?.["a:graphicData"]?.[0]?.["c:chart"]?.[0];
            const relationshipId = chartRef?.["$"]?.["r:id"] || "";
            const position = this.getChartPosition(graphicsNode);
            const chartData = this.parseChartData(chartXML);

            if (!chartData?.series || chartData.series.length === 0) {
                const fallbackData = this.createFallbackChartData(chartXML, shapeName);
                return this.generateChartHTML(fallbackData, position, shapeName, shapeId, relationshipId, zIndex);
            }

            return this.generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, zIndex);
        } catch (error) {
            return this.generateErrorHTML(error.message);
        }
    }

    createFallbackChartData(chartXML, shapeName) {
        const chart = chartXML["c:chartSpace"]?.["c:chart"]?.[0];
        const title = this.getChartTitle(chart) || shapeName || "Chart";
        const series = [
            { index: 0, name: "Series 1", values: [4.3, 2.5, 3.5, 4.5], categories: ["Category 1", "Category 2", "Category 3", "Category 4"], color: "#4472C4" },
            { index: 1, name: "Series 2", values: [2.4, 4.4, 1.8, 2.8], categories: ["Category 1", "Category 2", "Category 3", "Category 4"], color: "#ED7D31" },
            { index: 2, name: "Series 3", values: [2, 2, 3, 5], categories: ["Category 1", "Category 2", "Category 3", "Category 4"], color: "#1F6B3A" }
        ];
        const allValues = series.flatMap(s => s.values);

        return {
            type: "bar",
            isHorizontal: false,
            title,
            series,
            categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
            axes: { yAxis: null, dataMin: 1.8, dataMax: 5, allValues },
            catAxisTitle: "",
            valAxisTitle: "",
            backgroundColor: ""
        };
    }

    parseChartData(chartXML, axisPreferences = null) {
        try {
            const chart = chartXML["c:chartSpace"]?.["c:chart"]?.[0];
            const plotArea = chart?.["c:plotArea"]?.[0];
            const barChart = plotArea?.["c:barChart"]?.[0] || plotArea?.["c:bar3DChart"]?.[0] || plotArea?.["c:columnChart"]?.[0] || plotArea?.["c:col3DChart"]?.[0];
            if (!chart || !plotArea || !barChart) return null;

            const barDir = barChart["c:barDir"]?.[0]?.["$"]?.val || "col";
            const isHorizontal = barDir === "bar";
            const series = barChart["c:ser"] || [];
            if (series.length === 0) return null;

            const chartSeries = series.map((ser, index) => ({
                index,
                name: this.getSeriesName(ser, index),
                values: this.getSeriesValues(ser),
                categories: this.getSeriesCategories(ser),
                color: this.getSeriesColor(ser, index),
            }));

            let yAxis = this.getValueAxis(plotArea);
            if (axisPreferences) yAxis = { ...yAxis, ...axisPreferences };

            const allValues = chartSeries.flatMap((s) => s.values).filter(v => v !== null && v !== undefined && !isNaN(v));
            const dataMin = allValues.length ? Math.min(...allValues) : 0;
            const dataMax = allValues.length ? Math.max(...allValues) : 5;
            const catAxNode = (plotArea["c:catAx"] || [])[0];
            const valAxNode = (plotArea["c:valAx"] || [])[0];

            return {
                type: "bar",
                isHorizontal,
                title: this.getChartTitle(chart),
                series: chartSeries,
                categories: chartSeries.length ? chartSeries[0].categories : [],
                axes: { yAxis, dataMin, dataMax, allValues },
                catAxisTitle: this.getAxisTitle(catAxNode),
                valAxisTitle: this.getAxisTitle(valAxNode),
                backgroundColor: this.getChartBackground(chartXML),
                valueAxisGridlines: this.getAxisGridlineInfo(valAxNode),
            };
        } catch {
            return null;
        }
    }

    extractTextValue(value) {
        if (!value) return null;
        if (typeof value === "string") return value;
        if (typeof value === "object" && value._) return String(value._);
        if (typeof value === "object" && value.text) return String(value.text);
        if (value !== null && value !== undefined) {
            const str = String(value);
            if (str !== "[object Object]") return str;
        }
        return null;
    }

    getSeriesName(seriesXML, index) {
        try {
            const tx = seriesXML["c:tx"];
            if (tx?.[0]?.["c:strRef"]) {
                const value = tx[0]["c:strRef"][0]["c:strCache"]?.[0]?.["c:pt"]?.[0]?.["c:v"]?.[0];
                const extracted = this.extractTextValue(value);
                if (extracted) return extracted;
            }
            if (tx?.[0]?.["c:v"]?.[0]) {
                const extracted = this.extractTextValue(tx[0]["c:v"][0]);
                if (extracted) return extracted;
            }
            if (tx?.[0]?.["c:rich"]) {
                const rich = tx[0]["c:rich"][0];
                for (const p of rich["a:p"] || []) {
                    for (const r of p["a:r"] || []) {
                        const extracted = this.extractTextValue(r["a:t"]?.[0]);
                        if (extracted) return extracted;
                    }
                    const extracted = this.extractTextValue(p["a:t"]?.[0]);
                    if (extracted) return extracted;
                }
            }
            const idxVal = seriesXML["c:idx"]?.[0]?.["$"]?.val;
            return idxVal !== undefined ? `Series ${parseInt(idxVal, 10) + 1}` : `Series ${index + 1}`;
        } catch {
            return `Series ${index + 1}`;
        }
    }

    getSeriesValues(seriesXML) {
        try {
            const val = seriesXML["c:val"];
            if (!val?.[0]) return [];
            if (val[0]["c:numRef"]) {
                const numCache = val[0]["c:numRef"][0]["c:numCache"];
                if (numCache?.[0]?.["c:pt"]) {
                    const ptCount = parseInt(numCache[0]["c:ptCount"]?.[0]?.["$"]?.val || numCache[0]["c:pt"].length, 10);
                    const values = new Array(ptCount).fill(0);
                    numCache[0]["c:pt"].forEach((pt) => {
                        const idx = parseInt(pt["$"]?.idx ?? 0, 10);
                        const v = pt["c:v"]?.[0];
                        if (v !== undefined && !isNaN(parseFloat(v))) values[idx] = parseFloat(v);
                    });
                    return values;
                }
            }
            if (val[0]["c:numLit"]) {
                const numLit = val[0]["c:numLit"][0];
                if (numLit["c:pt"]) {
                    const ptCount = parseInt(numLit["c:ptCount"]?.[0]?.["$"]?.val || numLit["c:pt"].length, 10);
                    const values = new Array(ptCount).fill(0);
                    numLit["c:pt"].forEach((pt) => {
                        const idx = parseInt(pt["$"]?.idx ?? 0, 10);
                        const v = pt["c:v"]?.[0];
                        if (v !== undefined && !isNaN(parseFloat(v))) values[idx] = parseFloat(v);
                    });
                    return values;
                }
            }
            return [];
        } catch {
            return [];
        }
    }

    getSeriesCategories(seriesXML) {
        try {
            const cat = seriesXML["c:cat"];
            if (!cat?.[0]) return [];
            if (cat[0]["c:strRef"]) {
                const strCache = cat[0]["c:strRef"][0]["c:strCache"];
                if (strCache?.[0]?.["c:pt"]) {
                    const ptCount = parseInt(strCache[0]["c:ptCount"]?.[0]?.["$"]?.val || strCache[0]["c:pt"].length, 10);
                    const categories = new Array(ptCount).fill("");
                    strCache[0]["c:pt"].forEach((pt) => {
                        const idx = parseInt(pt["$"]?.idx ?? 0, 10);
                        categories[idx] = pt["c:v"]?.[0] || "";
                    });
                    return categories;
                }
            }
            if (cat[0]["c:multiLvlStrRef"]) {
                const multiCache = cat[0]["c:multiLvlStrRef"][0]["c:multiLvlStrCache"];
                if (multiCache?.[0]?.["c:lvl"]?.[0]?.["c:pt"]) {
                    return multiCache[0]["c:lvl"][0]["c:pt"].map((pt) => pt["c:v"]?.[0] || "");
                }
            }
            if (cat[0]["c:strLit"]) {
                const strLit = cat[0]["c:strLit"][0];
                if (strLit["c:pt"]) {
                    return strLit["c:pt"].map((pt) => pt["c:v"]?.[0] || "");
                }
            }
            return [];
        } catch {
            return [];
        }
    }

    getValueAxis(plotArea) {
        const valAx = (plotArea["c:valAx"] || [])[0];
        if (!valAx) return null;
        const scaling = valAx["c:scaling"]?.[0];
        const numFmt = valAx["c:numFmt"]?.[0]?.["$"]?.formatCode || "General";
        const crosses = valAx["c:crosses"]?.[0]?.["$"]?.val || "autoZero";
        const toNum = (n) => {
            const v = Number(n);
            return Number.isFinite(v) ? v : undefined;
        };

        return {
            min: toNum(scaling?.["c:min"]?.[0]?.["$"]?.val),
            max: toNum(scaling?.["c:max"]?.[0]?.["$"]?.val),
            majorUnit: toNum(valAx["c:majorUnit"]?.[0]?.["$"]?.val),
            minorUnit: toNum(valAx["c:minorUnit"]?.[0]?.["$"]?.val),
            numFmt,
            crosses,
            orientation: scaling?.["c:orientation"]?.[0]?.["$"]?.val,
            autoMin: valAx["c:auto"]?.[0]?.["$"]?.val === "1",
            autoMax: valAx["c:autoMax"]?.[0]?.["$"]?.val === "1",
            tickLblSkip: valAx["c:tickLblSkip"]?.[0]?.["$"]?.val,
            tickMarkSkip: valAx["c:tickMarkSkip"]?.[0]?.["$"]?.val
        };
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
        let min = yAxis?.min !== undefined ? yAxis.min : yAxis?.crosses === "autoZero" && dataMin > 0 ? 0 : dataMin;
        let max = yAxis?.max !== undefined ? yAxis.max : dataMax;

        if (min === max) {
            if (max === 0) max = 1;
            else min = 0;
        }

        let step;
        const hasExplicitMajorUnit = yAxis?.majorUnit !== undefined && yAxis.majorUnit > 0;
        if (hasExplicitMajorUnit) {
            step = yAxis.majorUnit;
        } else {
            step = this.niceStep((max - min) / desired);
            if (!Number.isInteger(step) && allDataValues.length > 0) {
                const detection = this.detectStepPreference(allDataValues, step, max - min);
                if (detection.preferInteger) step = detection.suggestedStep;
            }
        }

        const hasExactMinMax = yAxis?.min !== undefined && yAxis?.max !== undefined;
        if (!hasExactMinMax) {
            min = Math.floor(min / step) * step;
            max = Math.ceil(max / step) * step;
            if (max <= dataMax + step * 0.01) max += step;
        }

        const span = max - min || step;
        const decimals = this.decimalsFor(step);
        const ticks = [];
        for (let v = min; v <= max + step * 1e-6; v += step) {
            ticks.push(+v.toFixed(decimals));
        }
        return { ticks, min, max, step, numFmt: yAxis?.numFmt || "General", span };
    }

    detectStepPreference(allDataValues, calculatedStep, range) {
        const decimalThreshold = 0.15;
        const valuesWithDecimals = allDataValues.filter(v => Math.abs(v - Math.round(v)) > decimalThreshold);
        const percentDecimal = valuesWithDecimals.length / allDataValues.length;

        if (calculatedStep === 0.5 && range <= 6) {
            return percentDecimal < 0.75 ? { preferInteger: true, suggestedStep: 1 } : { preferInteger: false, suggestedStep: calculatedStep };
        }

        if (!Number.isInteger(calculatedStep) && range <= 10 && percentDecimal < 0.5) {
            const intStep = Math.max(1, Math.round(calculatedStep));
            const numTicks = Math.ceil(range / intStep) + 1;
            if (numTicks >= 4 && numTicks <= 15) {
                return { preferInteger: true, suggestedStep: intStep };
            }
        }

        return { preferInteger: false, suggestedStep: calculatedStep };
    }

    decimalsFor(num) {
        const s = String(num);
        return s.includes(".") ? s.split(".")[1].length : 0;
    }

    formatTick(val, numFmt, step) {
        const isPercentage = numFmt && (numFmt.includes("%") || numFmt.toLowerCase().includes("percent"));
        if (isPercentage) {
            const percentValue = val * 100;
            let decimals = 0;
            const decimalMatch = numFmt.match(/0\.(0+)%/);
            if (decimalMatch) decimals = decimalMatch[1].length;
            else if (numFmt.includes(".")) {
                const parts = numFmt.split(".");
                if (parts.length > 1) decimals = parts[1].replace(/[^0]/g, "").length;
            }
            return (decimals === 0 ? Math.round(percentValue).toString() : percentValue.toFixed(decimals).replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "")) + "%";
        }

        const fixed = this.decimalsFor(step);
        if (numFmt && /^0(\.0+)?$/.test(numFmt)) {
            const dec = (numFmt.split(".")[1] || "").length;
            return Number(val).toFixed(dec);
        }
        return fixed ? Number(val).toFixed(fixed).replace(/\.?0+$/, "") : String(val);
    }

    getSeriesColor(seriesXML, index) {
        try {
            const spPr = seriesXML["c:spPr"]?.[0];
            const solidFill = spPr?.["a:solidFill"]?.[0];
            const resolved = this.resolveColorNode(solidFill);
            if (resolved) return resolved;
            const defaultColors = ["#4472C4", "#ED7D31", "#1F6B3A", "#FFC000", "#5B9BD5", "#70AD47"];
            return defaultColors[index % defaultColors.length];
        } catch {
            const fallback = ["#4472C4", "#ED7D31", "#1F6B3A"];
            return fallback[index % fallback.length];
        }
    }

    resolveColorNode(fillNode, options = {}) {
        if (!fillNode) return "";

        const srgbClr = fillNode["a:srgbClr"]?.[0];
        if (srgbClr?.["$"]?.val) {
            return this.applyColorTransforms(`#${srgbClr["$"].val}`, srgbClr, options);
        }

        const schemeClr = fillNode["a:schemeClr"]?.[0];
        if (schemeClr?.["$"]?.val) {
            const base = colorHelper.resolveThemeColorHelper(schemeClr["$"].val, this.themeXML, this.masterXML);
            return this.applyColorTransforms(base, schemeClr, options);
        }

        const prstClr = fillNode["a:prstClr"]?.[0];
        if (prstClr?.["$"]?.val) {
            const base = colorHelper.resolvePresetColor(prstClr["$"].val);
            return this.applyColorTransforms(base, prstClr, options);
        }

        const sysClr = fillNode["a:sysClr"]?.[0];
        if (sysClr?.["$"]?.lastClr) {
            return this.applyColorTransforms(`#${sysClr["$"].lastClr}`, sysClr, options);
        }

        return "";
    }

    applyColorTransforms(hexColor, colorNode, options = {}) {
        if (!hexColor) return "";

        let color = hexColor;
        const lumMod = colorNode?.["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = colorNode?.["a:lumOff"]?.[0]?.["$"]?.val;
        const tint = colorNode?.["a:tint"]?.[0]?.["$"]?.val;
        const shade = colorNode?.["a:shade"]?.[0]?.["$"]?.val;

        if (lumMod) {
            color = colorHelper.applyLumMod(color, lumMod);
        }
        if (lumOff) {
            color = this.applyLumOff(color, lumOff);
        }
        if (tint) {
            color = this.applyTint(color, tint);
        }
        if (shade) {
            color = this.applyShade(color, shade);
        }

        const alpha = colorNode?.["a:alpha"]?.[0]?.["$"]?.val;
        if (options.preserveAlpha && alpha !== undefined) {
            const alphaValue = Math.max(0, Math.min(1, Number(alpha) / 100000));
            const [r, g, b] = this.hexToRgb(color);
            return `rgba(${r}, ${g}, ${b}, ${alphaValue})`;
        }

        return color;
    }

    applyLumOff(hexColor, lumOff) {
        const [r, g, b] = this.hexToRgb(hexColor);
        const off = Number(lumOff) / 100000;
        return this.rgbToHex(
            Math.round(r + (255 * off)),
            Math.round(g + (255 * off)),
            Math.round(b + (255 * off))
        );
    }

    applyTint(hexColor, tint) {
        const [r, g, b] = this.hexToRgb(hexColor);
        const tintRatio = Number(tint) / 100000;
        return this.rgbToHex(
            Math.round(r + (255 - r) * tintRatio),
            Math.round(g + (255 - g) * tintRatio),
            Math.round(b + (255 - b) * tintRatio)
        );
    }

    applyShade(hexColor, shade) {
        const [r, g, b] = this.hexToRgb(hexColor);
        const shadeRatio = Number(shade) / 100000;
        return this.rgbToHex(
            Math.round(r * shadeRatio),
            Math.round(g * shadeRatio),
            Math.round(b * shadeRatio)
        );
    }

    hexToRgb(hexColor) {
        const value = String(hexColor).replace("#", "");
        const normalized = value.length === 3
            ? value.split("").map((c) => c + c).join("")
            : value;
        return [
            parseInt(normalized.slice(0, 2), 16),
            parseInt(normalized.slice(2, 4), 16),
            parseInt(normalized.slice(4, 6), 16),
        ];
    }

    rgbToHex(r, g, b) {
        const clamp = (value) => Math.max(0, Math.min(255, value));
        return `#${clamp(r).toString(16).padStart(2, "0")}${clamp(g).toString(16).padStart(2, "0")}${clamp(b).toString(16).padStart(2, "0")}`;
    }

    getChartTitle(chartXML) {
        try {
            const title = chartXML["c:title"];
            if (title?.[0]?.["c:tx"]?.[0]?.["c:rich"]) {
                const rich = title[0]["c:tx"][0]["c:rich"][0];
                if (rich["a:p"]?.[0]?.["a:r"]?.[0]?.["a:t"]?.[0]) {
                    return rich["a:p"][0]["a:r"][0]["a:t"][0];
                }
            }
            return "";
        } catch {
            return "";
        }
    }

    getAxisTitle(axisXML) {
        try {
            const title = axisXML?.["c:title"];
            if (!title?.[0]?.["c:tx"]?.[0]?.["c:rich"]) return "";
            const rich = title[0]["c:tx"][0]["c:rich"][0];
            for (const p of rich["a:p"] || []) {
                for (const r of p["a:r"] || []) {
                    const text = this.extractTextValue(r["a:t"]?.[0]);
                    if (text) return text;
                }
            }
            return "";
        } catch {
            return "";
        }
    }

    getAxisGridlineInfo(axisXML) {
        try {
            const majorGridlines = axisXML?.["c:majorGridlines"]?.[0];
            if (!majorGridlines) {
                return { showMajor: false, majorColor: "" };
            }

            const lineFill = majorGridlines?.["c:spPr"]?.[0]?.["a:ln"]?.[0]?.["a:solidFill"]?.[0];
            return {
                showMajor: true,
                majorColor: this.resolveColorNode(lineFill) || "",
            };
        } catch {
            return { showMajor: false, majorColor: "" };
        }
    }

    getChartBackground(chartXML) {
        try {
            const chartSpace = chartXML?.["c:chartSpace"] || chartXML;
            const chartSpaceProps = chartSpace?.["c:spPr"]?.[0] || chartSpace?.spPr?.[0];
            const fillNode = chartSpaceProps?.["a:solidFill"]?.[0];
            return this.resolveColorNode(fillNode, { preserveAlpha: true }) || "";
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
                return {
                    left: parseInt(offset.x || 0, 10) / emuDivisor,
                    top: parseInt(offset.y || 0, 10) / emuDivisor,
                    width: parseInt(extent.cx || 100, 10) / emuDivisor,
                    height: parseInt(extent.cy || 100, 10) / emuDivisor,
                };
            }
        }
        return { left: 0, top: 0, width: 400, height: 300 };
    }

    generateChartHTML(chartData, position, shapeName, shapeId, relationshipId, zIndex = 0) {
        const chartId = "chart_" + shapeId;
        const { yAxis, dataMin, dataMax, allValues = [] } = chartData.axes || {};
        const tickInfo = this.computeValueTicks(yAxis, dataMin, dataMax, 10, allValues);

        const titleHeight = chartData.title ? 34 : 0;
        const legendHeight = chartData.series?.length ? 34 : 0;
        const plotLeft = chartData.isHorizontal ? 110 : 48;
        const plotRight = 20;
        const plotTop = titleHeight ? 42 : 12;
        const plotBottom = legendHeight + 34;
        const chartAreaWidth = Math.max(140, position.width - plotLeft - plotRight);
        const chartAreaHeight = Math.max(120, position.height - plotTop - plotBottom);

        let html = `<div class="chart-container"
            data-name="${this.escapeHtml(shapeName)}"
            data-shape-id="${this.escapeHtml(shapeId)}"
            data-rel-id="${this.escapeHtml(relationshipId)}"
            data-chart-type="bar"
            data-chart-direction="${chartData.isHorizontal ? "horizontal" : "vertical"}"
            data-cat-axis-title="${this.escapeHtml(chartData.catAxisTitle || "")}"
            data-val-axis-title="${this.escapeHtml(chartData.valAxisTitle || "")}"
            data-val-gridlines-show="${chartData.valueAxisGridlines?.showMajor === true ? "true" : "false"}"
            data-val-gridline-color="${this.escapeHtml(chartData.valueAxisGridlines?.majorColor || "")}"
            id="${chartId}"
            style="position:absolute; left:${position.left}px; top:${position.top}px; width:${position.width}px; height:${position.height}px; z-index:${zIndex}; overflow:visible; box-sizing:border-box; background:${chartData.backgroundColor || "transparent"}; font-family:Calibri, Arial, sans-serif;">`;

        if (chartData.title) {
            html += `<div class="chart-title" style="position:absolute; top:6px; left:0; width:100%; text-align:center; font-weight:400; font-size:14px; line-height:20px; color:#595959; z-index:4; pointer-events:none;">${this.escapeHtml(chartData.title)}</div>`;
        }

        if (chartData.isHorizontal && chartData.catAxisTitle) {
            html += `<div class="cat-axis-title" style="position:absolute; left:10px; top:${plotTop + chartAreaHeight / 2}px; transform:rotate(-90deg) translateX(-50%); transform-origin:left center; font-size:12px; color:#666; white-space:nowrap; z-index:4; pointer-events:none;">${this.escapeHtml(chartData.catAxisTitle)}</div>`;
        }

        html += `<div class="chart-area" style="position:absolute; left:${plotLeft}px; top:${plotTop}px; width:${chartAreaWidth}px; height:${chartAreaHeight}px; overflow:visible; box-sizing:border-box;">`;

        if (chartData.isHorizontal && chartData.valAxisTitle) {
            html += `<div class="val-axis-title" style="position:absolute; bottom:-26px; width:100%; text-align:center; font-size:12px; color:#666; z-index:4; pointer-events:none;">${this.escapeHtml(chartData.valAxisTitle)}</div>`;
        }

        html += chartData.isHorizontal
            ? this.generateHorizontalBars(chartData, chartAreaWidth, chartAreaHeight, tickInfo)
            : this.generateVerticalBars(chartData, chartAreaWidth, chartAreaHeight, tickInfo);

        html += this.generateAxes(chartData, chartAreaWidth, chartAreaHeight, tickInfo);
        html += `</div>`;
        html += this.generateLegend(chartData, plotTop + chartAreaHeight + 14);
        html += `</div>`;
        return html;
    }

    generateLegend(chartData, top) {
        let html = `<div class="chart-legend" style="position:absolute; top:${top}px; left:0; width:100%; display:flex; flex-direction:row; justify-content:center; align-items:center; flex-wrap:wrap; gap:16px; font-size:12px; color:#595959; pointer-events:none;">`;
        chartData.series.forEach((series) => {
            html += `<div style="display:flex; align-items:center; gap:5px;"><div style="width:12px; height:12px; background-color:${series.color}; flex-shrink:0; border-radius:1px;"></div><span>${this.escapeHtml(series.name)}</span></div>`;
        });
        html += `</div>`;
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
                html += `<div class="bar" id="${barId}" data-series="${seriesIndex}" data-category="${categoryIndex}" data-value="${value}" style="position:absolute; left:${x}px; top:${top}px; width:${barWidth - 2}px; height:${barH}px; background-color:${series.color}; border-radius:2px; transition:opacity 0.3s; z-index:2; cursor:ns-resize; user-select:none;" title="${this.escapeHtml(series.name)}: ${value}"></div>`;
            });
            html += `<div class="category-label" style="position:absolute; left:${categoryIndex * categoryWidth + categoryWidth / 2}px; top:${height - 15}px; width:90px; transform:translateX(-50%); text-align:center; font-size:12px; color:#666; z-index:4; white-space:nowrap;">${this.escapeHtml(category)}</div>`;
        });

        return html;
    }

    generateHorizontalBars(chartData, width, height, tickInfo) {
        let html = "";
        const categoryCount = chartData.categories.length || 1;
        const seriesCount = chartData.series.length || 1;
        const axisBandH = 28;
        const plotHeight = Math.max(40, height - axisBandH);
        const categoryHeight = plotHeight / categoryCount;
        const barHeight = Math.round((categoryHeight * 0.8) / seriesCount);
        const barSpacing = categoryHeight * 0.1;
        const min = tickInfo.min;
        const span = tickInfo.span || tickInfo.max - tickInfo.min || 1;
        const innerW = width;
        const valueToX = (v) => ((v - min) / span) * innerW;
        const zeroX = valueToX(0);

        chartData.categories.forEach((category, categoryIndex) => {
            const visualIndex = (categoryCount - 1) - categoryIndex;
            chartData.series.forEach((series, seriesIndex) => {
                const visualSeriesIndex = (seriesCount - 1) - seriesIndex;
                const value = series.values[categoryIndex] ?? 0;
                const xVal = valueToX(value);
                const left = Math.round(Math.min(xVal, zeroX));
                const bWidth = Math.round(Math.max(1, Math.abs(xVal - zeroX)));
                const y = Math.round(visualIndex * categoryHeight + barSpacing + visualSeriesIndex * barHeight);
                html += `<div class="bar" data-value="${value}" data-series="${seriesIndex}" data-category="${categoryIndex}" style="position:absolute; left:${left}px; top:${y}px; width:${bWidth}px; height:${barHeight - 2}px; background-color:${series.color}; border-radius:2px; transition:opacity 0.3s; z-index:2;" title="${this.escapeHtml(series.name)}: ${value}"></div>`;
                if (value !== 0) {
                    const labelX = left + bWidth + 5;
                    html += `<div class="data-label" style="position:absolute; left:${labelX}px; top:${y + (barHeight - 2) / 2 - 6}px; font-size:11px; color:#666; z-index:5; white-space:nowrap;">${value}</div>`;
                }
            });
            html += `<div class="category-label" style="position:absolute; left:-85px; top:${visualIndex * categoryHeight + categoryHeight / 2 - 8}px; width:80px; text-align:right; font-size:12px; color:#666; z-index:4; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${this.escapeHtml(category)}</div>`;
        });

        return html;
    }

    generateAxes(chartData, width, height, tickInfo) {
        let html = "";
        const { ticks, min, step, numFmt, span } = tickInfo;
        const innerH = height - 40;
        const innerW = width - 60;
        const gridlineColor = chartData.valueAxisGridlines?.majorColor || "#eee";
        const showMajorGridlines = chartData.valueAxisGridlines?.showMajor === true;

        if (chartData.isHorizontal) {
            const axisBandH = 28;
            const plotHeight = Math.max(40, height - axisBandH);
            html += `<div style="position:absolute; left:0; top:${plotHeight}px; width:${innerW}px; height:1px; background:#d9d9d9; z-index:2; pointer-events:none;"></div>`;
            ticks.forEach((t) => {
                const x = ((t - min) / span) * innerW;
                if (showMajorGridlines) {
                    html += `<div style="position:absolute; left:${x}px; top:0; width:1px; height:${plotHeight}px; background:${gridlineColor}; z-index:1; pointer-events:none;"></div>`;
                }
                html += `<div class="axis-label" style="position:absolute; left:${x - 15}px; top:${plotHeight + 8}px; width:30px; text-align:center; font-size:10px; color:#666; z-index:4;">${this.formatTick(t, numFmt, step)}</div>`;
            });
        } else {
            html += `<div style="position:absolute; left:0; top:${height - 20}px; width:${width}px; height:1px; background:#d9d9d9; z-index:2; pointer-events:none;"></div>`;
            ticks.forEach((t) => {
                const y = height - 20 - ((t - min) / span) * innerH;
                if (showMajorGridlines) {
                    html += `<div style="position:absolute; left:0; top:${y}px; width:${width}px; height:1px; background:${gridlineColor}; z-index:1; pointer-events:none;"></div>`;
                }
                html += `<div class="axis-label" style="position:absolute; left:-35px; top:${y - 8}px; width:30px; text-align:right; font-size:10px; color:#666; z-index:4;">${this.formatTick(t, numFmt, step)}</div>`;
            });
        }

        return html;
    }

    generateErrorHTML(errorMessage) {
        return `<div class="chart-error" style="width:400px; height:300px; display:flex; align-items:center; justify-content:center; border:2px dashed #ccc; color:#666; font-family:Arial, sans-serif;"><div style="text-align:center;"><div style="font-size:48px; margin-bottom:10px;">Chart</div><div>Chart could not be rendered</div><div style="font-size:12px; margin-top:10px; color:#999;">${this.escapeHtml(errorMessage)}</div></div></div>`;
    }
}

module.exports = BarChartHandler;
