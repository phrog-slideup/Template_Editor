const pptBackgroundColors = require("../../pptx-To-Html-styling/pptBackgroundColors.js");
const colorHelper = require("../../api/helper/colorHelper.js");

function getShapeFillColor(shapeNode, themeXML, masterXML = null) {

    let fillColor = "transparent";
    let fillOpacity = 1;
    let strokeColor = "transparent";
    let strokeOpacity = 1.0;
    let originalAlpha = '';

    try {
        const shapeFill = shapeNode?.["p:spPr"]?.[0];
        let originalThemeColor = '', originalLumMod = '', originalLumOff = '';
        // let txtPhType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type || '';

        // Handle Shape Fill (interior)
        const solidFill = shapeFill?.["a:solidFill"]?.[0];

        if (solidFill?.["a:srgbClr"]?.[0]?.["$"]?.val) {

            originalThemeColor = solidFill["a:srgbClr"][0]["$"].val;

            fillColor = `#${solidFill["a:srgbClr"][0]["$"].val}`;

            const lumMod = solidFill["a:srgbClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            if (lumMod) {
                originalLumMod = lumMod;
                const lumOff = solidFill["a:srgbClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                if (lumOff) {
                    originalLumOff = lumOff;
                    fillColor = pptBackgroundColors.applyLuminanceModifier(fillColor, lumMod, lumOff);
                } else {
                    fillColor = colorHelper.applyLumMod(fillColor, lumMod);
                }
            }

            // Extract fill opacity
            const fillAlphaNode = solidFill["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
            if (fillAlphaNode) {
                const alphaValue = parseInt(fillAlphaNode, 10);
                originalAlpha = alphaValue;

                if (!isNaN(alphaValue)) {
                    fillOpacity = alphaValue / 100000;
                }
            }

        } else if (solidFill?.["a:schemeClr"]?.[0]?.["$"]?.val) {
            let schmClr = solidFill["a:schemeClr"][0]["$"].val;

            originalThemeColor = schmClr;

            if (schmClr) {
                fillColor = colorHelper.resolveThemeColorHelper(schmClr, themeXML, masterXML);
            }

            const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;

            if (lumMod) {
                originalLumMod = lumMod;

                const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                if (lumOff) {
                    originalLumOff = lumOff;
                    fillColor = pptBackgroundColors.applyLuminanceModifier(fillColor, lumMod, lumOff);
                } else {
                    fillColor = colorHelper.applyLumMod(fillColor, lumMod);
                }
            }

            // Extract fill opacity for scheme color
            const fillAlphaNode = solidFill["a:schemeClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
            if (fillAlphaNode) {
                const alphaValue = parseInt(fillAlphaNode, 10);
                originalAlpha = alphaValue;
                if (!isNaN(alphaValue)) {
                    fillOpacity = alphaValue / 100000;
                }
            }
        }

        // Handle Line Stroke (border)
        const lineNode = shapeFill?.["a:ln"]?.[0];

        if (lineNode?.["a:solidFill"]?.[0]) {
            const strokeFill = lineNode["a:solidFill"][0];

            if (strokeFill?.["a:srgbClr"]?.[0]?.["$"]?.val) {
                strokeColor = `#${strokeFill["a:srgbClr"][0]["$"].val}`;

                const lumMod = strokeFill["a:srgbClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    originalLumMod = lumMod;
                    const lumOff = strokeFill["a:srgbClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        originalLumOff = lumOff;
                        strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                    } else {
                        strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                    }
                }

                const strokeAlphaNode = strokeFill["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                if (strokeAlphaNode) {
                    const alphaValue = parseInt(strokeAlphaNode, 10);
                    originalAlpha = alphaValue;

                    if (!isNaN(alphaValue)) {
                        strokeOpacity = alphaValue / 100000;
                    }
                }

            } else if (strokeFill?.["a:schemeClr"]?.[0]?.["$"]?.val) {
                let schmClr = strokeFill["a:schemeClr"][0]["$"].val;

                // NOTE: you had this.resolveMasterColor here - keeping your code
                if (masterXML && schmClr && this?.resolveMasterColor) {
                    const resolvedColor = this.resolveMasterColor(schmClr, masterXML);
                    if (resolvedColor) schmClr = resolvedColor;
                }

                if (schmClr) {
                    strokeColor = colorHelper.resolveThemeColorHelper(schmClr, themeXML, masterXML);
                }

                const lumMod = strokeFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    const lumOff = strokeFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                    } else {
                        strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                    }
                }

                const strokeAlphaNode = strokeFill["a:schemeClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                if (strokeAlphaNode) {
                    const alphaValue = parseInt(strokeAlphaNode, 10);
                    originalAlpha = alphaValue;

                    if (!isNaN(alphaValue)) {
                        strokeOpacity = alphaValue / 100000;
                    }
                }
            }
        }

        // ✅ NEW: Handle Gradient Stroke on <a:ln>
        // Your XML has: lineNode["a:gradFill"][0]...
        if (lineNode?.["a:gradFill"]?.[0]) {
            const lnGradFill = lineNode["a:gradFill"][0];
            const gradResult = getGradientFill(lnGradFill, themeXML, masterXML);

            // CSS border can't be true gradient; BUT you can still return it for SVG renderers
            // Keep strokeColor as gradient string so your SVG renderer can use it.
            if (gradResult?.fillColor) {
                strokeColor = gradResult.fillColor;
            }
            if (gradResult?.opacity !== undefined) {
                strokeOpacity = gradResult.opacity;
            }
        }

        // Handle Gradient Fill (shape interior)
        // ❌ Old bug: shapeFill["a:gradFill"] is an array
        const gradFill = shapeFill?.["a:gradFill"]?.[0];

        if (gradFill?.["a:gsLst"]) {
            const gradientResult = getGradientFill(gradFill, themeXML, masterXML);
            if (gradientResult) {
                fillColor = gradientResult.fillColor;
                if (gradientResult.opacity !== undefined) {
                    fillOpacity = gradientResult.opacity;
                }
            }
        }

        return {
            originalThemeColor,
            originalLumMod,
            originalLumOff,
            originalAlpha,
            fillColor,
            opacity: fillOpacity,
            fillOpacity: fillOpacity,
            strokeColor: strokeColor,
            strokeOpacity: strokeOpacity
        };

    } catch (error) {
        console.error("Error processing shape fill color:", error);
        return {
            fillColor: "transparent",
            opacity: 1.0,
            fillOpacity: 1.0,
            strokeColor: "transparent",
            strokeOpacity: 1.0
        };
    }
}

/**
 * ✅ UPDATED: getGradientFill now works for BOTH shape <a:gradFill> and line <a:ln><a:gradFill>
 * and does NOT rely on "this.*" (which was breaking your schemeClr gradient stops).
 */
function getGradientFill(gradFill, themeXML, masterXML) {
    if (!gradFill || !gradFill["a:gsLst"] || !gradFill["a:gsLst"][0]["a:gs"]) {
        console.error("Invalid gradient fill structure:", gradFill);
        return { fillColor: 'linear-gradient(90deg, transparent, transparent)', opacity: 1 };
    }

    const gsList = gradFill["a:gsLst"][0]["a:gs"];
    const hasLinear = gradFill["a:lin"];
    const hasPath = gradFill["a:path"];
    const pathType = hasPath ? gradFill["a:path"][0]["$"]?.path || "circle" : null;
    let gradientType = hasLinear ? "linear" : "radial";

    if (pathType === "circle") gradientType = "radial";
    else if (pathType === "rect") gradientType = "rectangular";
    else if (pathType === "shape") gradientType = "path";

    // --- Extract & convert PowerPoint angle to CSS ---
    let pptxAngle = parseInt(gradFill["a:lin"]?.[0]?.["$"]?.ang || 0, 10);
    pptxAngle = isNaN(pptxAngle) ? 0 : pptxAngle / 60000;

    // PPTX 0° (bottom->top) -> CSS 90° (top->bottom)
    let cssAngle = (pptxAngle + 90) % 360;

    // --- Handle flip (x/y/xy) ---
    const flip = gradFill["$"]?.flip || "none";
    if (flip === "x") cssAngle = (180 - cssAngle + 360) % 360;
    else if (flip === "y") cssAngle = (360 - cssAngle) % 360;
    else if (flip === "xy") cssAngle = (cssAngle + 180) % 360;

    let totalAlpha = 0;
    let alphaCount = 0;

    let stopsObj = gsList.map(stop => {
        let hex = "#000000";
        let alpha = 1;

        if (stop["a:srgbClr"]?.[0]?.["$"]?.val) {
            hex = `#${stop["a:srgbClr"][0]["$"].val}`;
            if (stop["a:srgbClr"][0]["a:alpha"]) {
                alpha = parseInt(stop["a:srgbClr"][0]["a:alpha"][0]["$"].val, 10) / 100000;
            }
        } else if (stop["a:schemeClr"]?.[0]?.["$"]?.val) {
            const schemeClr = stop["a:schemeClr"][0];
            const key = schemeClr["$"].val;

            // ✅ FIX: resolve theme scheme color properly using your helper
            hex = colorHelper.resolveThemeColorHelper(key, themeXML, masterXML) || "#000000";

            // ✅ FIX: apply luminance mods the same way as your solid fill code
            const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = schemeClr["a:lumOff"]?.[0]?.["$"]?.val;

            if (lumMod) {
                if (lumOff) hex = pptBackgroundColors.applyLuminanceModifier(hex, lumMod, lumOff);
                else hex = colorHelper.applyLumMod(hex, lumMod);
            }

            if (schemeClr["a:alpha"]) {
                alpha = parseInt(schemeClr["a:alpha"][0]["$"].val, 10) / 100000;
            }
        }

        const posVal = parseInt(stop["$"].pos, 10); // 0–100000
        const posPct = posVal / 1000;               // 0–100%

        totalAlpha += alpha;
        alphaCount++;

        return {
            posVal,
            posPct,
            rgbaStr: `rgba(${hexToRGB(hex)}, ${alpha}) ${posPct}%`
        };
    });

    const avgOpacity = alphaCount > 0 ? totalAlpha / alphaCount : 1;

    // --- Radial center (from fillToRect) ---
    let radialPosition = "center";
    if (hasPath && gradFill["a:path"][0]["a:fillToRect"]) {
        const rect = gradFill["a:path"][0]["a:fillToRect"][0]["$"] || {};
        const l = rect.l !== undefined ? parseInt(rect.l) : null;
        const t = rect.t !== undefined ? parseInt(rect.t) : null;
        const r = rect.r !== undefined ? parseInt(rect.r) : null;
        const b = rect.b !== undefined ? parseInt(rect.b) : null;

        if (l === 50000 && t === 50000 && r === 50000 && b === 50000) radialPosition = "center";
        else if (l === 100000 && t === 100000) radialPosition = "right bottom";
        else if (l === 100000 && b === 100000) radialPosition = "right top";
        else if (r === 100000 && t === 100000) radialPosition = "left bottom";
        else if (r === 100000 && b === 100000) radialPosition = "left top";
        else if (l === 100000) radialPosition = "right center";
        else if (r === 100000) radialPosition = "left center";
        else if (t === 100000) radialPosition = "center bottom";
        else if (b === 100000) radialPosition = "center top";
    }

    // --- Build final CSS stops string ---
    let gradientStops;
    if (gradientType === "radial") {
        stopsObj.sort((a, b) => a.posVal - b.posVal);
        gradientStops = stopsObj.map(s => s.rgbaStr);
    } else if (gradientType === "linear") {
        stopsObj.sort((a, b) => a.posVal - b.posVal);
        gradientStops = stopsObj.map(s => s.rgbaStr);
    } else {
        gradientStops = stopsObj.map(s => s.rgbaStr);
    }

    // --- Compose final CSS gradient ---
    let fillColor = "";
    switch (gradientType) {
        case "linear":
            fillColor = `linear-gradient(${cssAngle}deg, ${gradientStops.join(", ")})`;
            break;
        case "radial":
            fillColor = `radial-gradient(circle at ${radialPosition}, ${gradientStops.join(", ")})`;
            break;
        case "rectangular":
            fillColor = `radial-gradient(ellipse at ${radialPosition}, ${gradientStops.join(", ")})`;
            break;
        case "path":
            fillColor = `radial-gradient(circle closest-side at ${radialPosition}, ${gradientStops.join(", ")})`;
            break;
        default:
            fillColor = `linear-gradient(${cssAngle}deg, ${gradientStops.join(", ")})`;
            break;
    }

    return { fillColor, opacity: avgOpacity };
}

function hexToRGB(hex) {
    if (!hex) return "0, 0, 0";
    const clean = String(hex).replace("#", "");
    const r = parseInt(clean.substring(0, 2), 16) || 0;
    const g = parseInt(clean.substring(2, 4), 16) || 0;
    const b = parseInt(clean.substring(4, 6), 16) || 0;
    return `${r}, ${g}, ${b}`;
}

module.exports = { getShapeFillColor };
