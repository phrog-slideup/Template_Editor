// DoughnutChartHandler.js – FINAL VERSION with borders between slices
const colorHelper = require("../../api/helper/colorHelper.js");

class DoughnutChartHandler {
    constructor(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML) {
        this.graphicsNode = graphicsNode;
        this.chartXML = chartXML;
        this.chartRelsXML = chartRelsXML;
        this.chartColorsXML = chartColorsXML;
        this.chartStyleXML = chartStyleXML;
        this.themeXML = themeXML;
    }

    getEMUDivisor() {
        return 12700;
    }

    // -------------------------------
    // MAIN ENTRY POINT
    // -------------------------------
    async convertChartToHTML(
        graphicsNode,
        chartXML,
        chartRelsXML,
        chartColorsXML,
        chartStyleXML,
        themeXML,
        zIndex = 0
    ) {
        try {
            const shapeName =
                graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name ||
                "DoughnutChart";

            const position = this.getChartPosition(graphicsNode);
            const chartData = this.extractDoughnutData(chartXML);

            if (!chartData) {
               // return this.generateErrorHTML("No doughnut chart data found.");
            }

            return this.generateChartHTML(chartData, position, shapeName, zIndex);
        } catch (err) {
            console.error("DoughnutChartHandler convert error:", err);
           // return this.generateErrorHTML(err.message || "Unknown error");
        }
    }

    // -------------------------------
    // PARSE DOUGHNUT XML
    // -------------------------------
    extractDoughnutData(chartXML) {
        const cs = chartXML["c:chartSpace"];
        if (!cs) return null;

        const chart = cs["c:chart"]?.[0];
        if (!chart) return null;

        const plot = chart["c:plotArea"]?.[0];
        if (!plot) return null;

        const chartNode =
            plot["c:doughnutChart"]?.[0] ||
            plot["c:pieChart"]?.[0] ||
            plot["c:pie3DChart"]?.[0] ||
            plot["c:ofPieChart"]?.[0];
        if (!chartNode) return null;

        const chartType = plot["c:doughnutChart"]?.[0]
            ? "doughnut"
            : plot["c:pie3DChart"]?.[0]
                ? "pie3D"
                : plot["c:ofPieChart"]?.[0]
                    ? "ofPie"
                    : "pie";

        const firstSliceAngVal = chartNode["c:firstSliceAng"]?.[0]?.["$"]?.val;
        const firstSliceAng = firstSliceAngVal !== undefined
            ? parseFloat(firstSliceAngVal)
            : 0;

        const holeSizeVal = chartNode["c:holeSize"]?.[0]?.["$"]?.val;
        const holeSize = holeSizeVal !== undefined
            ? parseFloat(holeSizeVal) / 100
            : chartType === "doughnut" ? 0.6 : 0;

        const ser = chartNode["c:ser"]?.[0];
        if (!ser) return null;

        const values = this.getValues(ser);
        const labels = this.getCategories(ser);
        const seriesName = this.getSeriesName(ser);
        const title = this.getChartTitle(chart) || seriesName || "";
        const legend = this.getLegendConfig(chart);
        const dataLabels = this.getDataLabelConfig(chartNode, ser);
        const { colors, borderColor, borderWidth } = this.getColorsAndBorders(ser, values.length);
        const threeD = this.getThreeDPieStyle(chart, ser, chartType);
        const ofPie = this.getOfPieConfig(chartNode, chartType, values, colors, dataLabels);

        if (!values.length) return null;

        return {
            values,
            labels,
            colors,
            holeSize,
            firstSliceAng,
            borderColor,
            borderWidth,
            title,
            legend,
            dataLabels,
            chartType,
            threeD,
            ofPie
        };
    }

    getOfPieConfig(chartNode, chartType, values, colors, dataLabels) {
        if (chartType !== "ofPie") {
            return null;
        }

        const ofPieType = chartNode?.["c:ofPieType"]?.[0]?.["$"]?.val || "pie";
        const secondPieSize = parseFloat(chartNode?.["c:secondPieSize"]?.[0]?.["$"]?.val || 75);
        const syntheticColor = colors[values.length] || "#A02B93";
        const secondaryCount = values.length >= 4 ? 2 : 1;
        const splitIndex = Math.max(1, values.length - secondaryCount);
        const primaryValues = values.slice(0, splitIndex);
        const secondaryValues = values.slice(splitIndex);
        const aggregateValue = secondaryValues.reduce((sum, value) => sum + value, 0);
        const primaryColors = colors.slice(0, splitIndex).concat([syntheticColor]);
        const secondaryColors = colors.slice(splitIndex, splitIndex + secondaryCount);
        const aggregateLabelItem = (dataLabels?.items || []).find(item => item.idx >= values.length) || null;

        return {
            ofPieType,
            secondPieSize: Number.isFinite(secondPieSize) ? secondPieSize : 75,
            splitIndex,
            primaryValues: primaryValues.concat([aggregateValue]),
            secondaryValues,
            primaryColors,
            secondaryColors,
            aggregateValue,
            aggregateLabelItem
        };
    }

    getThreeDPieStyle(chart, ser, chartType) {
        if (chartType !== "pie3D") {
            return { rotX: 0, depthPercent: 0, contourWidth: 0 };
        }

        const view3D = chart?.["c:view3D"]?.[0];
        const rotX = parseFloat(view3D?.["c:rotX"]?.[0]?.["$"]?.val || 30);
        const depthPercent = parseFloat(view3D?.["c:depthPercent"]?.[0]?.["$"]?.val || 100);
        const contourWidth = parseFloat(
            ser?.["c:dPt"]?.[0]?.["c:spPr"]?.[0]?.["a:sp3d"]?.[0]?.["$"]?.contourW || 0
        );

        return { rotX, depthPercent, contourWidth };
    }

    getValues(ser) {
        try {
            const pts =
                ser["c:val"]?.[0]["c:numRef"]?.[0]["c:numCache"]?.[0]["c:pt"] || [];
            return pts.map(pt => parseFloat(pt["c:v"]?.[0] || 0));
        } catch {
            return [];
        }
    }

    getCategories(ser) {
        try {
            const pts =
                ser["c:cat"]?.[0]["c:strRef"]?.[0]["c:strCache"]?.[0]["c:pt"] || [];
            return pts.map(pt => pt["c:v"]?.[0] || "");
        } catch {
            return [];
        }
    }

    getSeriesName(ser) {
        try {
            return (
                ser?.["c:tx"]?.[0]?.["c:strRef"]?.[0]?.["c:strCache"]?.[0]?.["c:pt"]?.[0]?.["c:v"]?.[0] ||
                ser?.["c:tx"]?.[0]?.["c:v"]?.[0] ||
                ""
            );
        } catch {
            return "";
        }
    }

    getChartTitle(chart) {
        try {
            const autoTitleDeleted = chart?.["c:autoTitleDeleted"]?.[0]?.["$"]?.val === "1";
            if (autoTitleDeleted) return "";

            const titleNode = chart?.["c:title"]?.[0];
            if (!titleNode) return "";

            const richText =
                titleNode?.["c:tx"]?.[0]?.["c:rich"]?.[0]?.["a:p"] ||
                titleNode?.["c:tx"]?.[0]?.rich?.[0]?.["a:p"] ||
                [];

            const textRuns = [];
            richText.forEach((paragraph) => {
                const runs = paragraph?.["a:r"] || [];
                runs.forEach((run) => {
                    const text = run?.["a:t"]?.[0];
                    if (text) textRuns.push(text);
                });
            });

            return textRuns.join(" ").trim();
        } catch {
            return "";
        }
    }

    getLegendConfig(chart) {
        try {
            const legendNode = chart?.["c:legend"]?.[0];
            if (!legendNode) {
                return { display: false, position: "bottom" };
            }

            const overlayVal = legendNode?.["c:overlay"]?.[0]?.["$"]?.val;
            const legendPos = legendNode?.["c:legendPos"]?.[0]?.["$"]?.val || "r";
            const positionMap = {
                r: "right",
                l: "left",
                t: "top",
                b: "bottom"
            };

            return {
                display: true,
                overlay: overlayVal === "1",
                position: positionMap[legendPos] || "bottom"
            };
        } catch {
            return { display: false, position: "bottom" };
        }
    }

    getDataLabelConfig(doughnut, ser) {
        try {
            const dLbls =
                ser?.["c:dLbls"]?.[0] ||
                doughnut?.["c:dLbls"]?.[0];

            if (!dLbls) {
                return { display: false, items: [] };
            }

            const numFmt = dLbls?.["c:numFmt"]?.[0]?.["$"]?.formatCode || "";
            const showVal = dLbls?.["c:showVal"]?.[0]?.["$"]?.val === "1";
            const showPercent = dLbls?.["c:showPercent"]?.[0]?.["$"]?.val === "1";
            const showLeaderLines = dLbls?.["c:showLeaderLines"]?.[0]?.["$"]?.val === "1";
            const textColor = this.extractLabelTextColor(dLbls?.["c:txPr"]?.[0]);
            const dLblPos = dLbls?.["c:dLblPos"]?.[0]?.["$"]?.val || "";

            const items = (dLbls?.["c:dLbl"] || []).map((labelNode) => {
                const idx = parseInt(labelNode?.["c:idx"]?.[0]?.["$"]?.val, 10);
                const layout = labelNode?.["c:layout"]?.[0]?.["c:manualLayout"]?.[0];
                const x = parseFloat(layout?.["c:x"]?.[0]?.["$"]?.val || 0);
                const y = parseFloat(layout?.["c:y"]?.[0]?.["$"]?.val || 0);
                const itemTextColor = this.extractLabelTextColor(labelNode?.["c:txPr"]?.[0]);
                const itemDLblPos = labelNode?.["c:dLblPos"]?.[0]?.["$"]?.val || "";

                return {
                    idx,
                    x: Number.isFinite(x) ? x : 0,
                    y: Number.isFinite(y) ? y : 0,
                    textColor: itemTextColor || null,
                    dLblPos: itemDLblPos,
                    hasManualLayout: Boolean(layout),
                    isExplicit: true
                };
            }).filter(item => Number.isFinite(item.idx));

            return {
                display: showVal || showPercent,
                showVal,
                showPercent,
                showLeaderLines,
                numFmt,
                textColor: textColor || "#4D4D4D",
                dLblPos,
                items
            };
        } catch {
            return { display: false, items: [] };
        }
    }

    extractLabelTextColor(txPrNode) {
        try {
            const defRPr =
                txPrNode?.["a:p"]?.[0]?.["a:pPr"]?.[0]?.["a:defRPr"]?.[0] ||
                txPrNode?.["a:p"]?.[0]?.["a:endParaRPr"]?.[0];

            const solidFill = defRPr?.["a:solidFill"]?.[0];
            return this.resolveColorNode(solidFill);
        } catch {
            return null;
        }
    }

    // -------------------------------
    // DYNAMIC COLOR EXTRACTION
    // -------------------------------
    getColorsAndBorders(ser, count) {
        const colors = [];
        let borderColor = "#FFFFFF";
        let borderWidth = 0;

        try {

            // Extract series-level color
            let seriesColor = null;
            const serSpPr = ser["c:spPr"]?.[0];

            if (serSpPr) {
                const serFill = serSpPr["a:solidFill"]?.[0];
                if (serFill) {
                    seriesColor = this.resolveColorNode(serFill);
                }

                // Extract border
                const serLn = serSpPr["a:ln"]?.[0];
                if (serLn) {
                    const lnWidth = serLn["$"]?.w;
                    if (lnWidth) {
                        borderWidth = parseInt(lnWidth, 10) / 12700;
                    }

                    const lnFill = serLn["a:solidFill"]?.[0];
                    if (lnFill) {
                        borderColor = this.resolveColorNode(lnFill) || borderColor;
                    }
                }
            }

            // Initialize colors array
            for (let i = 0; i < count; i++) {
                colors[i] = null;
            }

            // Extract data point colors
            const dPtsRaw = ser["c:dPt"];
            if (dPtsRaw) {
                const dPts = Array.isArray(dPtsRaw) ? dPtsRaw : [dPtsRaw];

                dPts.forEach((dp, dpIndex) => {
                    try {
                        const idxVal = dp["c:idx"]?.[0]?.["$"]?.val;
                        const idx = idxVal !== undefined ? parseInt(idxVal, 10) : dpIndex;

                        const dpSpPr = dp["c:spPr"]?.[0];
                        if (dpSpPr) {
                            const dpFill = dpSpPr["a:solidFill"]?.[0];
                            if (dpFill) {
                                colors[idx] = this.resolveColorNode(dpFill);
                            }
                        }
                    } catch (e) {
                        console.warn(`Error at slice ${dpIndex}:`, e);
                    }
                });
            }

            // Fill missing colors
            const defaults = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
            for (let i = 0; i < count; i++) {
                if (!colors[i]) {
                    colors[i] = seriesColor || defaults[i % defaults.length];
                }
            }

        } catch (err) {
            console.error("Color extraction error:", err);
            const defaults = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
            for (let i = 0; i < count; i++) {
                colors[i] = defaults[i % defaults.length];
            }
        }

        return { colors, borderColor, borderWidth };
    }

    resolveColorNode(fillNode) {
        if (!fillNode) return null;

        const srgbClr = fillNode["a:srgbClr"]?.[0];
        if (srgbClr?.["$"]?.val) {
            return "#" + srgbClr["$"].val.toUpperCase();
        }

        const prstClr = fillNode["a:prstClr"]?.[0]?.["$"]?.val;
        if (prstClr) {
            return colorHelper.resolvePresetColor(prstClr);
        }

        const schemeClr = fillNode["a:schemeClr"]?.[0];
        if (schemeClr?.["$"]?.val) {
            const chartSchemeFallbacks = {
                bg1: "lt1",
                tx1: "dk1",
                bg2: "lt2",
                tx2: "dk2"
            };
            let resolved = colorHelper.resolveThemeColorHelper(
                schemeClr["$"].val,
                this.themeXML
            );

            if (!resolved && chartSchemeFallbacks[schemeClr["$"].val]) {
                resolved = colorHelper.resolveThemeColorHelper(
                    chartSchemeFallbacks[schemeClr["$"].val],
                    this.themeXML
                );
            }

            if (!resolved) return null;

            const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = schemeClr["a:lumOff"]?.[0]?.["$"]?.val;

            if (lumMod) {
                resolved = colorHelper.applyLumMod(
                    resolved,
                    parseInt(lumMod, 10),
                    lumOff ? parseInt(lumOff, 10) : 0
                );
            } else if (lumOff) {
                resolved = colorHelper.applyLumOff(resolved, lumOff);
            }

            return resolved;
        }

        return null;
    }


    // -------------------------------
    // POSITION
    // -------------------------------
    getChartPosition(graphicsNode) {
        const div = this.getEMUDivisor();
        const xfrm = graphicsNode?.["p:xfrm"]?.[0];

        const off = xfrm?.["a:off"]?.[0]?.["$"] || {};
        const ext = xfrm?.["a:ext"]?.[0]?.["$"] || {};

        return {
            left: parseInt(off.x || 0, 10) / div,
            top: parseInt(off.y || 0, 10) / div,
            width: parseInt(ext.cx || 2000000, 10) / div,
            height: parseInt(ext.cy || 1500000, 10) / div,
        };
    }

    // Calculate angular width of border based on pixel width and chart radius
    calculateBorderAngle(borderWidthPx, chartDiameter) {
        if (borderWidthPx <= 0) return 0;

        // Calculate the radius
        const radius = chartDiameter / 2;

        // Calculate the angular width of the border
        // Using the formula: angle = (arcWidth / radius) * (180 / PI)
        // For small angles, we can approximate: angle ≈ (borderWidth / radius) * (180 / PI)
        const angleRadians = borderWidthPx / radius;
        const angleDegrees = angleRadians * (180 / Math.PI);

        // Return the angle, with a minimum of 0.5 degrees for visibility
        return Math.max(angleDegrees, 0.5);
    }

    buildConicGradient(values, colors, borderColor, borderWidth, chartDiameter) {
        return this.buildConicGradientAtAngle(values, colors, borderColor, borderWidth, chartDiameter, 0);
    }

    getSliceSeparatorColor(borderWidth, borderColor) {
        if (!(borderWidth > 0)) {
            return borderColor;
        }

        // PowerPoint pie-family charts consistently render slice separators as light lines,
        // even when theme resolution yields a darker stroke in our HTML path.
        return "#FFFFFF";
    }

    buildConicGradientAtAngle(values, colors, borderColor, borderWidth, chartDiameter, startAngleDeg = 0) {
        const total = values.reduce((sum, v) => sum + v, 0) || 1;
        const separatorColor = this.getSliceSeparatorColor(borderWidth, borderColor);
        const borderAngle = borderWidth > 0
            ? this.calculateBorderAngle(borderWidth, chartDiameter)
            : 0;

        let currentAngle = 0;
        const stops = [];

        for (let i = 0; i < values.length; i++) {
            const value = values[i];
            const sliceDegrees = (value / total) * 360;
            const adjustedSliceDegrees = sliceDegrees - borderAngle;
            const endAngle = currentAngle + adjustedSliceDegrees;
            stops.push(`${colors[i]} ${currentAngle.toFixed(2)}deg ${endAngle.toFixed(2)}deg`);

            if (borderWidth > 0) {
                const borderStart = endAngle;
                const borderEnd = endAngle + borderAngle;
                stops.push(`${separatorColor} ${borderStart.toFixed(2)}deg ${borderEnd.toFixed(2)}deg`);
                currentAngle = borderEnd;
            } else {
                currentAngle = endAngle;
            }
        }

        return {
            total,
            gradientCSS: `conic-gradient(from ${startAngleDeg}deg, ${stops.join(", ")})`
        };
    }

    formatDataLabelValue(value, total, dataLabels) {
        if (dataLabels?.showPercent) {
            return `${((value / total) * 100).toFixed(2)}%`;
        }

        if (dataLabels?.numFmt && dataLabels.numFmt.includes('%')) {
            return `${(value * 100).toFixed(2)}%`;
        }

        if (dataLabels?.showVal) {
            const rounded = Math.round(value * 100) / 100;
            return Number.isInteger(rounded)
                ? String(rounded)
                : rounded.toFixed(2).replace(/\.?0+$/, "");
        }

        return "";
    }

    getRotationForSliceCenter(values, targetSliceIndex, targetCenterDegFromTop = 90) {
        if (!Array.isArray(values) || !values.length) {
            return 0;
        }

        const total = values.reduce((sum, value) => sum + value, 0) || 1;
        const previousDegrees = values
            .slice(0, targetSliceIndex)
            .reduce((sum, value) => sum + ((value / total) * 360), 0);
        const targetSliceDegrees = ((values[targetSliceIndex] || 0) / total) * 360;
        const baseCenterDeg = previousDegrees + (targetSliceDegrees / 2);

        return (targetCenterDegFromTop - baseCenterDeg + 360) % 360;
    }

    buildPieSeparatorSvg(values, centerX, centerY, diameter, startAngleDeg, stroke = "#FFFFFF", strokeWidth = 2) {
        if (!Array.isArray(values) || values.length < 2) {
            return "";
        }

        const total = values.reduce((sum, value) => sum + value, 0) || 1;
        const radius = diameter / 2;
        let cumulativeAngle = startAngleDeg;
        const lines = [];

        const buildLine = (angleDeg) => {
            const radians = (angleDeg - 90) * (Math.PI / 180);
            const x2 = centerX + Math.cos(radians) * radius;
            const y2 = centerY + Math.sin(radians) * radius;
            return `<line x1="${centerX.toFixed(2)}" y1="${centerY.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${y2.toFixed(2)}" stroke="${stroke}" stroke-width="${strokeWidth}" stroke-linecap="round"></line>`;
        };

        // Include the seam between last and first slice as well.
        lines.push(buildLine(startAngleDeg));

        for (let index = 0; index < values.length; index++) {
            const sliceDegrees = (values[index] / total) * 360;
            cumulativeAngle += sliceDegrees;
            if (index < values.length - 1) {
                lines.push(buildLine(cumulativeAngle));
            }
        }

        if (!lines.length) {
            return "";
        }

        return `<svg style="position:absolute; inset:0; width:100%; height:100%; overflow:visible; pointer-events:none; z-index:2;">${lines.join("")}</svg>`;
    }

    renderOfPieChartHTML(data, position, shapeName, zIndex, chartId) {
        const {
            labels,
            colors,
            borderColor,
            borderWidth,
            title,
            legend,
            dataLabels,
            ofPie
        } = data;
        const separatorColor = this.getSliceSeparatorColor(borderWidth, borderColor);
        const { left, top, width, height } = position;
        const hasTitle = Boolean(title);
        const showLegend = legend?.display !== false && labels.some(Boolean);
        const titleHeight = hasTitle ? 22 : 0;
        const legendHeight = showLegend ? 48 : 0;
        const verticalPadding = 12;
        const availableChartHeight = Math.max(92, height - titleHeight - legendHeight - verticalPadding);
        const secondarySizeFactor = Math.max(0.45, Math.min(0.84, (ofPie.secondPieSize || 75) / 100));
        const chartWidthFactor = ofPie.ofPieType === "bar" ? 0.40 : 0.43;
        const primaryDiameter = Math.max(112, Math.min(width * chartWidthFactor, availableChartHeight * 1.02));
        const secondaryDiameter = primaryDiameter * secondarySizeFactor;
        const connectorWidth = Math.max(24, width * (ofPie.ofPieType === "bar" ? 0.10 : 0.11));
        const secondaryWidth = ofPie.ofPieType === "bar" ? secondaryDiameter * 0.44 : secondaryDiameter;
        const chartBandWidth = primaryDiameter + connectorWidth + secondaryWidth;
        const chartBandHeight = Math.max(primaryDiameter, secondaryDiameter);
        const bandLeft = Math.max(0, (width - chartBandWidth) / 2);
        const chartAreaTop = hasTitle ? titleHeight + 4 : 8;
        const bandTop = chartAreaTop + Math.max(0, (availableChartHeight - chartBandHeight) / 2);
        const primaryCenter = {
            x: bandLeft + primaryDiameter / 2,
            y: bandTop + chartBandHeight / 2
        };
        const secondaryCenter = {
            x: bandLeft + primaryDiameter + connectorWidth + secondaryWidth / 2,
            y: bandTop + chartBandHeight / 2
        };

        const aggregateIndex = Math.max(0, ofPie.primaryValues.length - 1);
        const primaryRotationDeg = this.getRotationForSliceCenter(ofPie.primaryValues, aggregateIndex, 90);
        const secondaryRotationDeg = ofPie.ofPieType === "pie"
            ? this.getRotationForSliceCenter(ofPie.secondaryValues, 0, 225)
            : primaryRotationDeg;

        const primaryRender = this.buildConicGradientAtAngle(
            ofPie.primaryValues,
            ofPie.primaryColors,
            borderColor,
            borderWidth,
            primaryDiameter,
            primaryRotationDeg
        );
        const secondaryRender = ofPie.ofPieType === "pie"
            ? this.buildConicGradientAtAngle(
                ofPie.secondaryValues,
                ofPie.secondaryColors,
                borderColor,
                borderWidth,
                secondaryDiameter,
                secondaryRotationDeg
            )
            : null;
        const primarySeparatorSvg = this.buildPieSeparatorSvg(
            ofPie.primaryValues,
            primaryCenter.x,
            primaryCenter.y,
            primaryDiameter,
            primaryRotationDeg,
            separatorColor,
            2
        );
        const secondarySeparatorSvg = ofPie.ofPieType === "pie"
            ? this.buildPieSeparatorSvg(
                ofPie.secondaryValues,
                secondaryCenter.x,
                secondaryCenter.y,
                secondaryDiameter,
                secondaryRotationDeg,
                separatorColor,
                2
            )
            : "";

        let cumulativePrimaryAngle = primaryRotationDeg;
        const mainLabels = ofPie.primaryValues.map((value, index) => {
            const text = this.formatDataLabelValue(value, primaryRender.total, dataLabels);
            if (!text) return "";
            const sliceDegrees = (value / primaryRender.total) * 360;
            const midAngle = cumulativePrimaryAngle + sliceDegrees / 2;
            cumulativePrimaryAngle += sliceDegrees;
            const radians = (midAngle - 90) * (Math.PI / 180);
            const radiusFactor = index === 0 && ofPie.ofPieType === "bar"
                ? 0.50
                : index === ofPie.primaryValues.length - 1
                    ? 0.54
                    : 0.57;
            const labelX = primaryCenter.x + Math.cos(radians) * (primaryDiameter / 2) * radiusFactor;
            const labelY = primaryCenter.y + Math.sin(radians) * (primaryDiameter / 2) * radiusFactor;
            const labelColor = index === ofPie.primaryValues.length - 1 && ofPie.aggregateLabelItem?.textColor
                ? ofPie.aggregateLabelItem.textColor
                : dataLabels?.textColor || "#FFFFFF";

            return `<div style="position:absolute; left:${labelX.toFixed(2)}px; top:${labelY.toFixed(2)}px; transform:translate(-50%, -50%); font-size:11px; line-height:1; color:${labelColor}; white-space:nowrap; text-align:center; z-index:3;">${text}</div>`;
        }).join("");

        let cumulativeSecondaryAngle = secondaryRotationDeg;
        const secondaryPieTotal = ofPie.secondaryValues.reduce((sum, num) => sum + num, 0) || 1;
        const secondaryPieLabels = ofPie.ofPieType === "pie"
            ? ofPie.secondaryValues.map((value, index) => {
                const text = this.formatDataLabelValue(value, secondaryPieTotal, dataLabels);
                if (!text) return "";
                const sliceDegrees = (value / secondaryPieTotal) * 360;
                const midAngle = cumulativeSecondaryAngle + sliceDegrees / 2;
                cumulativeSecondaryAngle += sliceDegrees;
                const radians = (midAngle - 90) * (Math.PI / 180);
                const labelX = secondaryCenter.x + Math.cos(radians) * (secondaryDiameter / 2) * 0.56;
                const labelY = secondaryCenter.y + Math.sin(radians) * (secondaryDiameter / 2) * 0.56;
                return `<div style="position:absolute; left:${labelX.toFixed(2)}px; top:${labelY.toFixed(2)}px; transform:translate(-50%, -50%); font-size:11px; line-height:1; color:${dataLabels?.textColor || "#FFFFFF"}; white-space:nowrap; text-align:center; z-index:3;">${text}</div>`;
            }).join("")
            : "";

        const secondaryBarHeight = secondaryDiameter;
        const secondaryBarWidth = secondaryWidth;
        const secondaryBarTop = secondaryCenter.y - secondaryBarHeight / 2;
        const secondaryBarLeft = secondaryCenter.x - secondaryBarWidth / 2;
        const secondaryTotal = ofPie.secondaryValues.reduce((sum, num) => sum + num, 0) || 1;
        let runningBarTop = secondaryBarTop;
        const secondaryBarSegments = ofPie.ofPieType === "bar"
            ? ofPie.secondaryValues.map((value, index) => {
                const segHeight = secondaryBarHeight * (value / secondaryTotal);
                const html = `<div style="position:absolute; left:${secondaryBarLeft.toFixed(2)}px; top:${runningBarTop.toFixed(2)}px; width:${secondaryBarWidth.toFixed(2)}px; height:${segHeight.toFixed(2)}px; background:${ofPie.secondaryColors[index]}; border-top:${index > 0 ? `1px solid ${separatorColor}` : '0'}; box-sizing:border-box; z-index:2;"></div>`;
                runningBarTop += segHeight;
                return html;
            }).join("")
            : "";

        runningBarTop = secondaryBarTop;
        const secondaryBarLabels = ofPie.ofPieType === "bar"
            ? ofPie.secondaryValues.map((value) => {
                const segHeight = secondaryBarHeight * (value / secondaryTotal);
                const text = this.formatDataLabelValue(value, secondaryTotal, dataLabels);
                const labelY = runningBarTop + segHeight / 2;
                runningBarTop += segHeight;
                return `<div style="position:absolute; left:${(secondaryBarLeft + secondaryBarWidth + 10).toFixed(2)}px; top:${labelY.toFixed(2)}px; transform:translate(0, -50%); font-size:11px; line-height:1; color:${dataLabels?.textColor || "#404040"}; white-space:nowrap; text-align:left; z-index:3;">${text}</div>`;
            }).join("")
            : "";

        const connectorStartX = bandLeft + primaryDiameter * 0.86;
        const connectorTopY = bandTop + chartBandHeight * 0.28;
        const connectorBottomY = bandTop + chartBandHeight * 0.72;
        const connectorEndX = secondaryCenter.x - secondaryWidth / 2;
        const connectorSvg = ofPie.ofPieType === "pie"
            ? `<svg style="position:absolute; inset:0; width:100%; height:100%; overflow:visible; pointer-events:none; z-index:2;">
                    <line x1="${connectorStartX.toFixed(2)}" y1="${connectorTopY.toFixed(2)}" x2="${connectorEndX.toFixed(2)}" y2="${(secondaryCenter.y - secondaryDiameter * 0.26).toFixed(2)}" stroke="rgba(170,170,170,0.95)" stroke-width="1"></line>
                    <line x1="${connectorStartX.toFixed(2)}" y1="${connectorBottomY.toFixed(2)}" x2="${connectorEndX.toFixed(2)}" y2="${(secondaryCenter.y + secondaryDiameter * 0.26).toFixed(2)}" stroke="rgba(170,170,170,0.95)" stroke-width="1"></line>
               </svg>`
            : `<svg style="position:absolute; inset:0; width:100%; height:100%; overflow:visible; pointer-events:none; z-index:2;">
                    <line x1="${connectorStartX.toFixed(2)}" y1="${connectorTopY.toFixed(2)}" x2="${connectorEndX.toFixed(2)}" y2="${secondaryBarTop.toFixed(2)}" stroke="rgba(170,170,170,0.95)" stroke-width="1"></line>
                    <line x1="${connectorStartX.toFixed(2)}" y1="${connectorBottomY.toFixed(2)}" x2="${connectorEndX.toFixed(2)}" y2="${(secondaryBarTop + secondaryBarHeight).toFixed(2)}" stroke="rgba(170,170,170,0.95)" stroke-width="1"></line>
               </svg>`;

        const titleHTML = hasTitle
            ? `<div class="chart-title" style="width:100%; text-align:center; font-size:12px; line-height:1.2; color:#666666; margin-bottom:8px; flex:0 0 auto;">${title}</div>`
            : "";

        const legendHTML = showLegend
            ? this.generateLegend(labels, colors, legend?.position || "bottom")
            : "";

        const mainPieHtml = `<div class="doughnut-chart" style="position:relative; width:${width - 20}px; height:${availableChartHeight}px; max-width:${width - 20}px; max-height:${availableChartHeight}px; margin:auto; margin-bottom:12px; flex:0 0 auto; background:transparent;">
                <div style="position:absolute; left:${(primaryCenter.x - primaryDiameter / 2).toFixed(2)}px; top:${(primaryCenter.y - primaryDiameter / 2).toFixed(2)}px; width:${primaryDiameter.toFixed(2)}px; height:${primaryDiameter.toFixed(2)}px; border-radius:50%; background:${primaryRender.gradientCSS}; z-index:1;"></div>
                ${primarySeparatorSvg}
                ${ofPie.ofPieType === "pie"
                    ? `<div style="position:absolute; left:${(secondaryCenter.x - secondaryDiameter / 2).toFixed(2)}px; top:${(secondaryCenter.y - secondaryDiameter / 2).toFixed(2)}px; width:${secondaryDiameter.toFixed(2)}px; height:${secondaryDiameter.toFixed(2)}px; border-radius:50%; background:${secondaryRender.gradientCSS}; z-index:1;"></div>${secondarySeparatorSvg}`
                    : secondaryBarSegments}
                ${connectorSvg}
                ${mainLabels}
                ${secondaryPieLabels}
                ${secondaryBarLabels}
            </div>`;

        const encodedValues = encodeURIComponent(JSON.stringify(data.values));
        const encodedLabels = encodeURIComponent(JSON.stringify(labels));
        const encodedColors = encodeURIComponent(JSON.stringify(data.colors));
        const encodedDataLabels = encodeURIComponent(JSON.stringify(dataLabels || {}));

        return `
<div class="chart-container"
     id="${chartId}"
     data-name="${shapeName}"
     data-chart-type="ofPie"
     data-chart-title="${String(title || "").replace(/"/g, '&quot;')}"
     data-chart-values="${encodedValues}"
     data-chart-labels="${encodedLabels}"
     data-chart-colors="${encodedColors}"
     data-chart-hole-size="0"
     data-chart-first-slice-ang="0"
     data-chart-legend-pos="${legend?.position || "bottom"}"
     data-chart-data-labels="${encodedDataLabels}"
     style="
         position:absolute;
         top:${top}px;
         left:${left}px;
         width:${width}px;
         height:${height}px;
         margin:0;
         padding:0;
         border-radius:4px;
         background:transparent;
         font-family:Arial, sans-serif;
         overflow:hidden;
         z-index:${zIndex};
         box-sizing:border-box;
     ">
    <div class="doughnut-wrapper"
         style="
           position:relative;
           width:100%;
           height:100%;
           display:flex;
           flex-direction:column;
           align-items:center;
           justify-content:flex-start;
           padding:10px;
           box-sizing:border-box;
         ">
        ${titleHTML}
        ${mainPieHtml}
        ${legendHTML}
    </div>
</div>`;
    }

    // -------------------------------
    // HTML RENDERING WITH SLICE BORDERS
    // -------------------------------
    generateChartHTML(data, position, shapeName, zIndex) {
        const chartId = "chart_" + Math.random().toString(36).slice(2);

        if (data.chartType === "ofPie" && data.ofPie) {
            return this.renderOfPieChartHTML(data, position, shapeName, zIndex, chartId);
        }

        const {
            values,
            labels,
            colors,
            holeSize,
            firstSliceAng,
            borderColor,
            borderWidth,
            title,
            legend,
            dataLabels,
            chartType,
            threeD
        } = data;
        const { left, top, width, height } = position;

        const total = values.reduce((sum, v) => sum + v, 0);

        if (total === 0) {
          //  return this.generateErrorHTML("Chart has no data (total is zero)");
        }

        const hasTitle = Boolean(title);
        const showLegend = legend?.display !== false && labels.some(Boolean);
        const titleHeight = hasTitle ? 30 : 0;
        const legendHeight = showLegend ? 50 : 0;
        const verticalPadding = 20;

        const availableChartHeight = Math.max(
            80,
            height - titleHeight - legendHeight - verticalPadding
        );

        const rotX = Number.isFinite(threeD?.rotX) ? threeD.rotX : 30;
        const depthPercent = Number.isFinite(threeD?.depthPercent) ? threeD.depthPercent : 100;
        const contourWidth = Number.isFinite(threeD?.contourWidth) ? threeD.contourWidth : 0;
        const contourBoost = Math.min(0.05, contourWidth / 500000);
        const pieScaleX = chartType === "pie3D" ? Math.max(1.08, 1 + (rotX / 190) + contourBoost) : 1;
        const pieScaleY = chartType === "pie3D" ? Math.max(0.52, 1 - (rotX / 85) - contourBoost * 0.25) : 1;
        const isThreeDPie = chartType === "pie3D" && holeSize === 0;
        const depthFactor = isThreeDPie ? Math.max(0.045, (depthPercent / 1400) + contourBoost * 0.18) : 0;

        // Size the chart by its visual footprint, not its unscaled box.
        // This keeps flat pies and 3D pies dynamic while letting 3D pies
        // grow wider like the original PPT when their vertical scale is smaller.
        const maxVisualWidth = Math.max(60, width - 20);
        const maxVisualHeight = Math.max(60, availableChartHeight - 10);
        const maxDiameterByWidth = maxVisualWidth / Math.max(pieScaleX, 0.01);
        const maxDiameterByHeight = maxVisualHeight / Math.max(pieScaleY + depthFactor, 0.01);
        const chartDiameter = Math.max(60, Math.min(maxDiameterByWidth, maxDiameterByHeight));
        const pieDepthPx = isThreeDPie ? chartDiameter * depthFactor : 0;
        const chartSurfaceWidth = isThreeDPie ? chartDiameter * pieScaleX : chartDiameter;
        const chartSurfaceHeight = isThreeDPie ? (chartDiameter * pieScaleY) + pieDepthPx : chartDiameter;
        const layerOffsetX = isThreeDPie ? (chartSurfaceWidth - chartDiameter) / 2 : 0;
        const layerOffsetY = isThreeDPie ? -((chartDiameter - (chartDiameter * pieScaleY)) / 2) : 0;

        // Calculate border angle (gap between slices)
        const separatorColor = this.getSliceSeparatorColor(borderWidth, borderColor);
        const borderAngle = borderWidth > 0
            ? this.calculateBorderAngle(borderWidth, chartDiameter)
            : 0;

        // Build gradient with borders between slices
        let currentAngle = 0;
        let stops = [];

        for (let i = 0; i < values.length; i++) {
            const value = values[i];
            const sliceDegrees = (value / total) * 360;

            // Reduce slice size to account for border
            const adjustedSliceDegrees = sliceDegrees - borderAngle;
            const endAngle = currentAngle + adjustedSliceDegrees;

            // Add the slice color
            stops.push(`${colors[i]} ${currentAngle.toFixed(2)}deg ${endAngle.toFixed(2)}deg`);

            // Add border/gap (except after the last slice)
            if (borderWidth > 0) {
                const borderStart = endAngle;
                const borderEnd = endAngle + borderAngle;
                stops.push(`${separatorColor} ${borderStart.toFixed(2)}deg ${borderEnd.toFixed(2)}deg`);
                currentAngle = borderEnd;
            } else {
                currentAngle = endAngle;
            }
        }

        // Build gradient starting from 0°
        const gradientCSS = `conic-gradient(from 0deg, ${stops.join(", ")})`;

        // Rotate the entire element by firstSliceAng
        const rotationDeg = firstSliceAng;

        const holePct = holeSize * 100;
        const holeOffset = (100 - holePct) / 2;
        const encodedValues = encodeURIComponent(JSON.stringify(values));
        const encodedLabels = encodeURIComponent(JSON.stringify(labels));
        const encodedColors = encodeURIComponent(JSON.stringify(colors));
        const encodedDataLabels = encodeURIComponent(JSON.stringify(dataLabels || {}));
        const baseStops = stops.map((stop) => {
            const parts = stop.split(" ");
            if (!parts[0]?.startsWith("#")) return stop;
            parts[0] = this.darkenHex(parts[0], 0.38);
            return parts.join(" ");
        });
        const baseGradientCSS = `conic-gradient(from 0deg, ${baseStops.join(", ")})`;
        const dataLabelMarkup = this.generateDataLabels({
            values,
            labels,
            holeSize,
            firstSliceAng,
            chartDiameter,
            dataLabels,
            chartType,
            pieScaleX,
            pieScaleY,
            surfaceOffsetX: layerOffsetX,
            surfaceOffsetY: layerOffsetY
        });

        const titleHTML = hasTitle
            ? `<div class="chart-title"
                    style="
                        width:100%;
                        text-align:center;
                        font-size:12px;
                        line-height:1.2;
                        color:#666666;
                        margin-bottom:8px;
                        flex:0 0 auto;
                    ">${title}</div>`
            : "";

        const legendHTML = showLegend
            ? this.generateLegend(labels, colors, legend?.position || "bottom")
            : "";

        const pieDepthLayer = isThreeDPie
            ? `<div style="
                   position:absolute;
                   left:${layerOffsetX}px;
                   top:${(layerOffsetY + pieDepthPx).toFixed(2)}px;
                   width:${chartDiameter}px;
                   height:${chartDiameter}px;
                   border-radius:50%;
                   background:${baseGradientCSS};
                   transform:rotate(${rotationDeg}deg) scaleX(${pieScaleX}) scaleY(${pieScaleY});
                   transform-origin:center center;
                   z-index:1;
               "></div>`
            : "";

        const pieFaceLayer = isThreeDPie
            ? `<div style="
                   position:absolute;
                   left:${layerOffsetX}px;
                   top:${layerOffsetY.toFixed(2)}px;
                   width:${chartDiameter}px;
                   height:${chartDiameter}px;
                   border-radius:50%;
                   background:${gradientCSS};
                   transform:rotate(${rotationDeg}deg) scaleX(${pieScaleX}) scaleY(${pieScaleY});
                   transform-origin:center center;
                   z-index:2;
               "></div>`
            : "";

        const pieFaceOutlineLayer = isThreeDPie
            ? `<div style="
                   position:absolute;
                   left:${layerOffsetX}px;
                   top:${layerOffsetY.toFixed(2)}px;
                   width:${chartDiameter}px;
                   height:${chartDiameter}px;
                   border-radius:50%;
                   border:1px solid #FFFFFF;
                   box-sizing:border-box;
                   transform:rotate(${rotationDeg}deg) scaleX(${pieScaleX}) scaleY(${pieScaleY});
                   transform-origin:center center;
                   z-index:3;
                   pointer-events:none;
               "></div>`
            : "";

        const chartSurfaceStyle = isThreeDPie
            ? `
               position:relative;
               width:${chartSurfaceWidth}px;
               height:${chartSurfaceHeight}px;
               max-width:${chartSurfaceWidth}px;
               max-height:${chartSurfaceHeight}px;
               margin:auto;
               flex:0 0 auto;
               background:transparent;
             `
            : `
               position:relative;
               width:${chartDiameter}px;
               height:${chartDiameter}px;
               max-width:${chartDiameter}px;
               max-height:${chartDiameter}px;
               border-radius:50%;
               background:${gradientCSS};
               transform:rotate(${rotationDeg}deg) scaleX(${pieScaleX}) scaleY(${pieScaleY});
               transform-origin:center center;
               margin:auto;
               flex:0 0 auto;
             `;

        return `
<div class="chart-container"
     id="${chartId}"
     data-name="${shapeName}"
     data-chart-type="${chartType === "doughnut" ? "doughnut" : "pie"}"
     data-chart-title="${String(title || "").replace(/"/g, '&quot;')}"
     data-chart-values="${encodedValues}"
     data-chart-labels="${encodedLabels}"
     data-chart-colors="${encodedColors}"
     data-chart-hole-size="${Math.round(holeSize * 100)}"
     data-chart-first-slice-ang="${Math.round(firstSliceAng)}"
     data-chart-legend-pos="${legend?.position || "bottom"}"
     data-chart-data-labels="${encodedDataLabels}"
     data-chart-show-value="${dataLabels?.showVal ? "1" : "0"}"
     data-chart-show-percent="${dataLabels?.showPercent ? "1" : "0"}"
     data-chart-show-leader-lines="${dataLabels?.showLeaderLines ? "1" : "0"}"
     data-chart-label-format="${String(dataLabels?.numFmt || "").replace(/"/g, '&quot;')}"
     data-chart-label-color="${String(dataLabels?.textColor || "").replace(/"/g, '&quot;')}"
     style="
         position:absolute;
         top:${top}px;
         left:${left}px;
         width:${width}px;
         height:${height}px;
         margin:0;
         padding:0;
         border-radius:4px;
         background:transparent;
         font-family:Arial, sans-serif;
         overflow:hidden;
         z-index:${zIndex};
         box-sizing:border-box;
     ">

    <div class="doughnut-wrapper"
         style="
           position:relative;
           width:100%;
           height:100%;
           display:flex;
           flex-direction:column;
           align-items:center;
           justify-content:flex-start;
           padding:10px;
           box-sizing:border-box;
         ">

        ${titleHTML}

        <div class="doughnut-chart"
             style="${chartSurfaceStyle}">

            ${pieDepthLayer}
            ${pieFaceLayer}
            ${pieFaceOutlineLayer}

            <div class="doughnut-hole"
                 style="
                   position:absolute;
                   width:${holePct}%;
                   height:${holePct}%;
                   top:${holeOffset}%;
                   left:${holeOffset}%;
                   background:white;
                   border-radius:50%;
                   z-index:2;
                    display:${holeSize > 0 ? "block" : "none"};
                 ">
            </div>

            ${dataLabelMarkup}

        </div>

        ${legendHTML}

    </div>

</div>
        `;
    }

    generateLegend(labels, colors, position = "bottom") {
        const isBottom = position === "bottom";
        const containerStyle = isBottom
            ? `
                display:flex;
                justify-content:center;
                align-items:center;
                flex-wrap:wrap;
                gap:14px;
                width:100%;
                margin-top:22px;
                overflow:visible;
                flex:0 0 auto;
            `
            : `
                display:flex;
                flex-direction:column;
                align-items:flex-start;
                gap:8px;
                margin-left:12px;
                overflow:visible;
                flex:0 0 auto;
            `;

        const items = labels.map((label, index) => `
            <div style="display:flex; align-items:center; gap:5px; min-height:14px; overflow:visible;">
                <div style="
                    width:8px;
                    height:8px;
                    background:${colors[index]};
                    flex:0 0 auto;
                "></div>
                <span style="
                    font-size:11px;
                    line-height:1.15;
                    color:#555555;
                    white-space:nowrap;
                    display:inline-block;
                    overflow:visible;
                ">${label}</span>
            </div>
        `).join("");

        return `<div class="chart-legend" style="${containerStyle}">${items}</div>`;
    }

    generateDataLabels({ values, holeSize, firstSliceAng, chartDiameter, dataLabels, chartType, pieScaleX = 1, pieScaleY = 1, surfaceOffsetX = 0, surfaceOffsetY = 0 }) {
        if (!dataLabels?.display) {
            return "";
        }

        const total = values.reduce((sum, value) => sum + value, 0);
        if (!total) return "";

        const center = chartDiameter / 2;
        const radius = chartDiameter / 2;
        const labelColor = dataLabels.textColor || "#4D4D4D";
        const leaderColor = "rgba(170,170,170,0.95)";

        let cumulativeAngle = Number(firstSliceAng) || 0;
        const itemsHtml = [];
        const lines = [];

        const formatValue = (value) => {
            if (dataLabels.showPercent) {
                return `${((value / total) * 100).toFixed(2)}%`;
            }

            if (dataLabels.numFmt && dataLabels.numFmt.includes('%')) {
                return `${(value * 100).toFixed(2)}%`;
            }

            if (dataLabels.showVal) {
                const rounded = Math.round(value * 100) / 100;
                return Number.isInteger(rounded)
                    ? String(rounded)
                    : rounded.toFixed(2).replace(/\.?0+$/, "");
            }

            return "";
        };

        const explicitItems = Array.isArray(dataLabels.items) ? dataLabels.items : [];
        const explicitItemMap = new Map(explicitItems.map((item) => [item.idx, item]));
        const labelItems = values.map((_, idx) => {
            const existing = explicitItemMap.get(idx);
            if (existing) return existing;
            return { idx, x: 0, y: 0, auto: true, isExplicit: false };
        });

        labelItems.forEach((item) => {
            const value = values[item.idx];
            if (!Number.isFinite(value)) return;

            const text = formatValue(value);
            if (!text) return;

            const sliceDegrees = (value / total) * 360;
            const midAngle = cumulativeAngle + (sliceDegrees / 2);
            cumulativeAngle += sliceDegrees;

            const radians = (midAngle - 90) * (Math.PI / 180);
            const cos = Math.cos(radians);
            const sin = Math.sin(radians);
            const scaledCos = cos * pieScaleX;
            const scaledSin = sin * pieScaleY;

            const startX = surfaceOffsetX + center + scaledCos * (radius * 0.78);
            const startY = surfaceOffsetY + center + scaledSin * (radius * 0.78);
            const elbowX = surfaceOffsetX + center + scaledCos * (radius * 1.02);
            const elbowY = surfaceOffsetY + center + scaledSin * (radius * 1.02);
            const innerRadius = radius * holeSize;
            const dLblPos = item.dLblPos || dataLabels.dLblPos || "";
            const useInsidePlacement =
                chartType === "pie3D"
                    ? (dLblPos === "bestFit" || dLblPos === "ctr" || dLblPos === "inEnd")
                    : (
                        item.isExplicit &&
                        (dLblPos === "bestFit" || dLblPos === "ctr" || dLblPos === "inEnd")
                    );

            let labelX;
            let labelY;
            let textAlign;
            let anchorX;

            if (useInsidePlacement) {
                const insideRadius = holeSize > 0
                    ? innerRadius + ((radius - innerRadius) * 0.52)
                    : radius * 0.62;
                labelX = surfaceOffsetX + center + scaledCos * insideRadius;
                labelY = surfaceOffsetY + center + scaledSin * insideRadius;
                textAlign = "center";
                anchorX = labelX;
            } else {
                const baseLabelRadius = radius * 1.28;
                const nudgeX = (Number.isFinite(item.x) ? item.x : 0) * radius * 0.55;
                const nudgeY = (Number.isFinite(item.y) ? item.y : 0) * radius * 0.55;
                labelX = surfaceOffsetX + center + scaledCos * baseLabelRadius + nudgeX;
                labelY = surfaceOffsetY + center + scaledSin * baseLabelRadius + nudgeY;
                textAlign = cos < -0.18 ? "right" : cos > 0.18 ? "left" : "center";
                anchorX = textAlign === "left"
                    ? labelX - 4
                    : textAlign === "right"
                        ? labelX + 4
                        : labelX;
            }

            const leaderDx = Math.abs(anchorX - elbowX);
            const leaderDy = Math.abs(labelY - elbowY);
            const labelLayoutY = Number.isFinite(item.y) ? item.y : 0;
            const shouldDrawLeader =
                !useInsidePlacement &&
                dataLabels.showLeaderLines &&
                labelLayoutY > -0.03 &&
                (
                    leaderDx > radius * 0.16 ||
                    (labelY > center && leaderDy > radius * 0.08)
                );

            if (shouldDrawLeader) {
                lines.push(
                    `<polyline points="${startX.toFixed(2)},${startY.toFixed(2)} ${elbowX.toFixed(2)},${elbowY.toFixed(2)} ${anchorX.toFixed(2)},${labelY.toFixed(2)}" fill="none" stroke="${leaderColor}" stroke-width="1"/>`
                );
            }

            const transform = textAlign === "left"
                ? "translate(0, -50%)"
                : textAlign === "right"
                    ? "translate(-100%, -50%)"
                    : "translate(-50%, -50%)";

            itemsHtml.push(`
                <div style="
                    position:absolute;
                    left:${labelX.toFixed(2)}px;
                    top:${labelY.toFixed(2)}px;
                    transform:${transform};
                    font-size:11px;
                    line-height:1;
                    color:${item.textColor || labelColor};
                    white-space:nowrap;
                    text-align:${textAlign};
                    z-index:3;
                ">${text}</div>
            `);
        });

        const svg = lines.length
            ? `<svg style="position:absolute; inset:0; width:100%; height:100%; overflow:visible; pointer-events:none; z-index:2;">${lines.join("")}</svg>`
            : "";

        return `${svg}${itemsHtml.join("")}`;
    }

    // generateErrorHTML(msg) {
    //     return `
    //     <div class="chart-container chart-error"
    //          style="border:2px dashed #aaa; padding:20px; width:300px; height:200px;
    //                 display:flex;align-items:center;justify-content:center;font-family:Arial;">
    //         <div style="text-align:center;">
    //             <div style="font-size:32px;margin-bottom:8px;">📊</div>
    //             <div>Doughnut chart could not be rendered</div>
    //             <div style="font-size:12px;margin-top:8px;color:#999;">${msg}</div>
    //         </div>
    //     </div>`;
    // }

    darkenHex(hex, amount = 0.3) {
        try {
            const normalized = hex.replace('#', '');
            const r = parseInt(normalized.slice(0, 2), 16);
            const g = parseInt(normalized.slice(2, 4), 16);
            const b = parseInt(normalized.slice(4, 6), 16);
            const factor = Math.max(0, Math.min(1, 1 - amount));
            const toHex = (value) => Math.max(0, Math.min(255, Math.round(value * factor)))
                .toString(16)
                .padStart(2, '0')
                .toUpperCase();
            return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
        } catch {
            return hex;
        }
    }
}

module.exports = DoughnutChartHandler;
