// DoughnutChartHandler.js – FINAL VERSION with borders between slices

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

        const doughnut = plot["c:doughnutChart"]?.[0];
        if (!doughnut) return null;

        const firstSliceAngVal = doughnut["c:firstSliceAng"]?.[0]?.["$"]?.val;
        const firstSliceAng = firstSliceAngVal !== undefined
            ? parseFloat(firstSliceAngVal)
            : 0;

        const holeSizeVal = doughnut["c:holeSize"]?.[0]?.["$"]?.val;
        const holeSize = holeSizeVal !== undefined
            ? parseFloat(holeSizeVal) / 100
            : 0.6;

        const ser = doughnut["c:ser"]?.[0];
        if (!ser) return null;

        const values = this.getValues(ser);
        const labels = this.getCategories(ser);
        const { colors, borderColor, borderWidth } = this.getColorsAndBorders(ser, values.length);

        if (!values.length) return null;

        return { values, labels, colors, holeSize, firstSliceAng, borderColor, borderWidth };
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
                    const serRgb = serFill["a:srgbClr"]?.[0];
                    if (serRgb?.["$"]?.val) {
                        seriesColor = "#" + serRgb["$"].val.toUpperCase();
                    }
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
                        const lnRgb = lnFill["a:srgbClr"]?.[0];
                        if (lnRgb?.["$"]?.val) {
                            borderColor = "#" + lnRgb["$"].val.toUpperCase();
                        }
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
                                // RGB color
                                const dpRgb = dpFill["a:srgbClr"]?.[0];
                                if (dpRgb?.["$"]?.val) {
                                    colors[idx] = "#" + dpRgb["$"].val.toUpperCase();
                                }
                                // Theme color
                                else {
                                    const schemeClr = dpFill["a:schemeClr"]?.[0];
                                    if (schemeClr?.["$"]?.val) {
                                        const themeMap = {
                                            'accent1': '#4472C4', 'accent2': '#ED7D31',
                                            'accent3': '#A5A5A5', 'accent4': '#FFC000',
                                            'accent5': '#5B9BD5', 'accent6': '#70AD47'
                                        };
                                        const theme = schemeClr["$"].val;
                                        colors[idx] = themeMap[theme] || '#4472C4';
                                    }
                                }
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

    // -------------------------------
    // HTML RENDERING WITH SLICE BORDERS
    // -------------------------------
    generateChartHTML(data, position, shapeName, zIndex) {
        const chartId = "chart_" + Math.random().toString(36).slice(2);

        const { values, labels, colors, holeSize, firstSliceAng, borderColor, borderWidth } = data;
        const { left, top, width, height } = position;

        const total = values.reduce((sum, v) => sum + v, 0);

        if (total === 0) {
          //  return this.generateErrorHTML("Chart has no data (total is zero)");
        }

        // Calculate the chart diameter for border angle calculation
        const chartDiameter = Math.min(width, height) - 20;

        // Calculate border angle (gap between slices)
        const borderAngle = borderWidth > 0
            ? this.calculateBorderAngle(borderWidth, chartDiameter)
            : 0;

        // Build gradient with borders between slices
        let currentAngle = 0;
        let stops = [];

        // console.log("=== Doughnut Chart Rendering ===");
        // console.log("First Slice Angle:", firstSliceAng);
        // console.log("Values:", values);
        // console.log("Colors:", colors);
        // console.log("Total:", total);
        // console.log("Border Angle:", borderAngle.toFixed(2) + "°");

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
                stops.push(`${borderColor} ${borderStart.toFixed(2)}deg ${borderEnd.toFixed(2)}deg`);
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


        return `
<div class="chart-container"
     id="${chartId}"
     data-name="${shapeName}"
     data-chart-type="doughnut"
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
           align-items:center;
           justify-content:center;
           padding:10px;
           box-sizing:border-box;
         ">

        <div class="doughnut-chart"
             style="
               position:relative;
               width:100%;
               height:100%;
               max-width:${chartDiameter}px;
               max-height:${chartDiameter}px;
               border-radius:50%;
               background:${gradientCSS};
               transform:rotate(${rotationDeg}deg);
               margin:auto;
             ">

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
                 ">
            </div>

        </div>

    </div>

</div>
        `;
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
}

module.exports = DoughnutChartHandler;