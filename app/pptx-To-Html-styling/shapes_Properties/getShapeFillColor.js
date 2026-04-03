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
        const pattFill = shapeFill?.["a:pattFill"]?.[0];
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

        // Handle Pattern Fill (pattFill)
        let patternPrst = '';
        let patternFg = '';
        let patternBg = '';
        if (pattFill) {
            console.log("osadksa");
            const prst = pattFill["$"]?.prst || "pct5";

            let fgColor = "#000000";
            const fgClr = pattFill["a:fgClr"]?.[0];
            if (fgClr?.["a:schemeClr"]?.[0]) {
                fgColor = colorHelper.resolveThemeColorHelper(
                    fgClr["a:schemeClr"][0]["$"].val, themeXML, masterXML
                );
            } else if (fgClr?.["a:srgbClr"]?.[0]) {
                fgColor = `#${fgClr["a:srgbClr"][0]["$"].val}`;
            }

            let bgColor = "#ffffff";
            const bgClr = pattFill["a:bgClr"]?.[0];
            if (bgClr?.["a:schemeClr"]?.[0]) {
                bgColor = colorHelper.resolveThemeColorHelper(
                    bgClr["a:schemeClr"][0]["$"].val, themeXML, masterXML
                );
            } else if (bgClr?.["a:srgbClr"]?.[0]) {
                bgColor = `#${bgClr["a:srgbClr"][0]["$"].val}`;
            }

            fillColor = generatePatternFillCSS(prst, fgColor, bgColor);
            // ✅ Store in outer variables so return can access them
            patternPrst = prst;
            patternFg = fgColor;
            patternBg = bgColor;
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
            strokeOpacity: strokeOpacity,
            patternPrst,
            patternFg,
            patternBg
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

function generatePatternFillCSS(prst, fgColor, bgColor) {
    const patternMap = {
        'pct5': { type: 'dots', size: 1, spacing: 6 },
        'pct10': { type: 'dots', size: 1, spacing: 4 },
        'pct20': { type: 'dots', size: 2, spacing: 5 },
        'pct25': { type: 'dots', size: 2, spacing: 4 },
        'pct30': { type: 'dots', size: 2, spacing: 3 },
        'pct40': { type: 'dots', size: 3, spacing: 4 },
        'pct50': { type: 'checker', size: 4 },
        'pct60': { type: 'dots', size: 3, spacing: 2 },
        'pct75': { type: 'dots', size: 4, spacing: 2 },
        'pct80': { type: 'dots', size: 5, spacing: 2 },
        'pct90': { type: 'dots', size: 6, spacing: 2 },
        'horz': { type: 'hlines', gap: 4 },
        'vert': { type: 'vlines', gap: 4 },
        'ltHorz': { type: 'hlines', gap: 2 },
        'ltVert': { type: 'vlines', gap: 2 },
        'dkHorz': { type: 'hlines', gap: 8 },
        'dkVert': { type: 'vlines', gap: 8 },
        'narHorz': { type: 'hlines', gap: 1 },
        'narVert': { type: 'vlines', gap: 1 },
        'dashHorz': { type: 'dashH' },
        'dashVert': { type: 'dashV' },
        'dnDiag': { type: 'diagDown', gap: 6 },
        'upDiag': { type: 'diagUp', gap: 6 },
        'ltDnDiag': { type: 'diagDown', gap: 4 },
        'ltUpDiag': { type: 'diagUp', gap: 4 },
        'dkDnDiag': { type: 'diagDown', gap: 8 },
        'dkUpDiag': { type: 'diagUp', gap: 8 },
        'wdDnDiag': { type: 'diagDown', gap: 12 },
        'wdUpDiag': { type: 'diagUp', gap: 12 },
        'cross': { type: 'cross', gap: 8 },
        'ltGrid': { type: 'cross', gap: 4 },
        'dkGrid': { type: 'cross', gap: 12 },
        'diagCross': { type: 'diagCross', gap: 8 },
        'smGrid': { type: 'cross', gap: 2 },
        'lgGrid': { type: 'cross', gap: 16 },
        'diagBrick': { type: 'brick', gap: 8 },
        'horzBrick': { type: 'horzBrick', gap: 8 },
        'smCheck': { type: 'checker', size: 2 },
        'lgCheck': { type: 'checker', size: 8 },
        'smConfetti': { type: 'dots', size: 1, spacing: 3 },
        'lgConfetti': { type: 'dots', size: 2, spacing: 5 },
        'zigZag': { type: 'diagCross', gap: 6 },
        'wave': { type: 'hlines', gap: 4 },
        'divot': { type: 'diagCross', gap: 4 },
        'shingle': { type: 'diagDown', gap: 6 },
        'trellis': { type: 'cross', gap: 4 },
        'sphere': { type: 'checker', size: 6 },
        'dotGrid': { type: 'cross', gap: 4 },
        'dotDmnd': { type: 'diagCross', gap: 4 },
        'plaid': { type: 'cross', gap: 6 },
        'weave': { type: 'diagCross', gap: 6 },
        'openDmnd': { type: 'diagCross', gap: 8 },
        'solidDmnd': { type: 'checker', size: 4 },
        'dkTrellis': { type: 'cross', gap: 8 },
        'smDmnd': { type: 'diagCross', gap: 4 },
        'dkDmnd': { type: 'diagCross', gap: 8 },
    };

    const pattern = patternMap[prst] || { type: 'dots', size: 1, spacing: 6 };

    // ✅ FIX: Encode colors BEFORE embedding in SVG
    // This handles # in hex colors and any other special chars
    const fg = fgColor || 'black';
    const bg = bgColor || 'white';

    // ✅ FIX: Use encodeURIComponent-safe SVG builder
    // All SVG attributes use double quotes — these will be %22 encoded at the end
    const makeSVG = (width, height, content) =>
        `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">${content}</svg>`;

    let svgContent = '';

    switch (pattern.type) {
        case 'dots': {
            const s = pattern.size || 1;
            const sp = pattern.spacing || 6;
            const total = s + sp;
            svgContent = makeSVG(total, total,
                `<rect width="${total}" height="${total}" fill="${bg}"/>` +
                `<circle cx="${s / 2}" cy="${s / 2}" r="${s / 2}" fill="${fg}"/>`
            );
            break;
        }
        case 'hlines': {
            const gap = pattern.gap || 4;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="0" y1="0" x2="${gap}" y2="0" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'vlines': {
            const gap = pattern.gap || 4;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="0" y1="0" x2="0" y2="${gap}" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'diagDown': {
            const gap = pattern.gap || 6;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="0" y1="0" x2="${gap}" y2="${gap}" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="${-gap / 2}" y1="${gap / 2}" x2="${gap / 2}" y2="${gap * 1.5}" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="${gap / 2}" y1="${-gap / 2}" x2="${gap * 1.5}" y2="${gap / 2}" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'diagUp': {
            const gap = pattern.gap || 6;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="0" y1="${gap}" x2="${gap}" y2="0" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="${-gap / 2}" y1="${gap / 2}" x2="${gap / 2}" y2="${-gap / 2}" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="${gap / 2}" y1="${gap * 1.5}" x2="${gap * 1.5}" y2="${gap / 2}" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'cross': {
            const gap = pattern.gap || 8;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="${gap / 2}" y1="0" x2="${gap / 2}" y2="${gap}" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="0" y1="${gap / 2}" x2="${gap}" y2="${gap / 2}" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'diagCross': {
            const gap = pattern.gap || 8;
            svgContent = makeSVG(gap, gap,
                `<rect width="${gap}" height="${gap}" fill="${bg}"/>` +
                `<line x1="0" y1="0" x2="${gap}" y2="${gap}" stroke="${fg}" stroke-width="1"/>` +
                `<line x1="${gap}" y1="0" x2="0" y2="${gap}" stroke="${fg}" stroke-width="1"/>`
            );
            break;
        }
        case 'checker': {
            const s = pattern.size || 4;
            const d = s * 2;
            svgContent = makeSVG(d, d,
                `<rect width="${d}" height="${d}" fill="${bg}"/>` +
                `<rect x="0" y="0" width="${s}" height="${s}" fill="${fg}"/>` +
                `<rect x="${s}" y="${s}" width="${s}" height="${s}" fill="${fg}"/>`
            );
            break;
        }
        default: {
            // dots fallback
            const s = 1, sp = 6, total = s + sp;
            svgContent = makeSVG(total, total,
                `<rect width="${total}" height="${total}" fill="${bg}"/>` +
                `<circle cx="${s / 2}" cy="${s / 2}" r="${s / 2}" fill="${fg}"/>`
            );
            break;
        }
    }

    // ✅ FIX: Use encodeURIComponent for the entire SVG string
    // This is the most reliable encoding for data URIs used in CSS
    const encoded = encodeURIComponent(svgContent);
    return `url(data:image/svg+xml,${encoded})`;
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
