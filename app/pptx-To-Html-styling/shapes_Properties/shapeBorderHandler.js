const colorHelper = require("../../api/helper/colorHelper.js");

function getShapeBorderCSS(shapeNode, clrMap, themeXML, masterXML) {
    try {
        const spPr = shapeNode?.["p:spPr"]?.[0];
        const styleNode = shapeNode?.["p:style"]?.[0];

        if (!spPr && !styleNode) {
            return { border: "none" };
        }

        let ln = spPr?.["a:ln"]?.[0] || null;
        let lnRefColor = null;

        // ========== Resolve lnRef (line reference from style/theme) ==========
        if (styleNode?.["a:lnRef"]?.[0] && themeXML) {

            const lnRef = styleNode["a:lnRef"][0];
            const idx = parseInt(lnRef.$?.idx || "0", 10);

            const lnStyleLst = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:fmtScheme"]?.[0]?.["a:lnStyleLst"]?.[0]?.["a:ln"];

            if (lnStyleLst && lnStyleLst[idx]) {
                const themeLn = lnStyleLst[idx];

                const mergedLn = JSON.parse(JSON.stringify(themeLn));

                if (ln) {
                    if (ln.$) {
                        Object.assign(mergedLn.$, ln.$);
                    }
                    for (const key in ln) {
                        if (key !== '$') {
                            mergedLn[key] = JSON.parse(JSON.stringify(ln[key]));
                        }
                    }
                }

                ln = mergedLn;

                const schemeClr = lnRef["a:schemeClr"]?.[0];
                if (schemeClr) {
                    lnRefColor = schemeClr;
                }
            }
        }

        if (!ln) {
            return { border: "none" };
        }

        if (ln["a:noFill"]) {
            return { border: "none" };
        }

        if (ln.$?.w) {
            const wEmu = parseInt(ln.$.w, 10);
            if (wEmu === 0 || wEmu <= 3000) {   //3175
                return { border: "none" };
            }
        }
        const hasAnyFill = !!(
            ln["a:solidFill"] ||
            ln["a:gradFill"] ||
            ln["a:pattFill"]
        );


        const hasWidthOrDashOrCompound = !!(
            (ln.$?.w && parseInt(ln.$.w, 10) > 0) ||
            ln["a:prstDash"] ||
            ln.$?.cmpd ||
            ln["a:custDash"]
        );

        if (!hasAnyFill && !hasWidthOrDashOrCompound) {
            return { border: "none" };
        }

        let widthPt = ln.$?.w ? parseInt(ln.$.w, 10) / 12700 : 1;



        if (widthPt < 0.2) {
            return { border: "none" };
        }
        
        let style = "solid";

        const prstDash = ln["a:prstDash"]?.[0]?.$?.val;
        const dashMap = {
            "dot": "dotted",
            "sysDot": "dotted",
            "dash": "dashed",
            "lgDash": "dashed",
            "dashDot": "dashed",
            "lgDashDot": "dashed",
            "lgDashDotDot": "dashed",
            "sysDash": "dashed",
            "sysDashDot": "dashed",
            "sysDashDotDot": "dashed",
            "solid": "solid"
        };

        if (prstDash) {
            style = dashMap[prstDash] || "solid";
        }

        if (ln["a:custDash"]) {
            style = "dashed";
        }

        const compound = ln.$?.cmpd;
        if (compound === "dbl" || compound === "tri") {
            style = "double";
        } else if (compound === "thickThin" || compound === "thinThick") {
            widthPt *= 1.5;
        }

        let color = "#000000";

        if (ln["a:solidFill"]?.[0]) {
            color = resolveFillColor(ln["a:solidFill"][0], clrMap, themeXML, masterXML);
        } else if (ln["a:gradFill"]?.[0]) {
            const firstStop = ln["a:gradFill"][0]["a:gsLst"]?.[0]?.["a:gs"]?.[0];
            if (firstStop) {
                color = resolveFillColor(firstStop, clrMap, themeXML, masterXML);
            }
        } else if (ln["a:pattFill"]?.[0]) {
            const fgClr = ln["a:pattFill"][0]["a:fgClr"]?.[0];
            if (fgClr) {
                color = resolveFillColor(fgClr, clrMap, themeXML, masterXML);
            }
        }
        else if (lnRefColor) {
            color = resolveSchemeColor(lnRefColor, clrMap, themeXML, masterXML);
        }

        if (!color || color === "#00000000") {
            color = "#000000";
        }

        return {
            border: `${widthPt.toFixed(2)}px ${style} ${color}`
        };
    } catch (error) {
        console.error("Border parse error:", error);
        return { border: "none" };
    }
}

/**
 * Resolve fill color from various color nodes
 */
function resolveFillColor(colorNode, clrMap, themeXML, masterXML) {
    if (!colorNode) return "#000000";

    if (colorNode["a:srgbClr"]?.[0]) {
        const rgb = colorNode["a:srgbClr"][0].$?.val;
        if (rgb) {
            return applyColorModifiers("#" + rgb, colorNode["a:srgbClr"][0]);
        }
    }

    if (colorNode["a:schemeClr"]?.[0]) {
        return resolveSchemeColor(colorNode["a:schemeClr"][0], clrMap, themeXML, masterXML);
    }

    if (colorNode["a:sysClr"]?.[0]) {
        const lastClr = colorNode["a:sysClr"][0].$?.lastClr;
        if (lastClr) {
            return applyColorModifiers("#" + lastClr, colorNode["a:sysClr"][0]);
        }
    }

    if (colorNode["a:prstClr"]?.[0]) {
        const prstVal = colorNode["a:prstClr"][0].$?.val;
        if (prstVal) {
            const baseColor = colorHelper.resolvePresetColor(prstVal);
            return applyColorModifiers(baseColor, colorNode["a:prstClr"][0]);
        }
    }

    if (colorNode["a:hslClr"]?.[0]) {
        return "#808080"; // fallback
    }

    return "#000000";
}

/**
 * Resolve scheme color with clrMap and modifiers
 */
function resolveSchemeColor(schemeNode, clrMap, themeXML, masterXML) {
    if (!schemeNode || !schemeNode.$?.val) return "#000000";

    const colorKey = schemeNode.$.val;
    let baseColor = colorHelper.resolveThemeColorHelper(colorKey, themeXML, masterXML) || "#000000";

    return applyColorModifiers(baseColor, schemeNode);
}

/**
 * Apply color modifiers – FIXED lumMod & lumOff using HSL space
 */
function applyColorModifiers(hexColor, modifierNode) {
    if (!hexColor || !modifierNode) return hexColor || "#000000";

    if (!hexColor.startsWith("#")) {
        hexColor = "#" + hexColor.replace(/^#/, '');
    }

    // Parse RGB [0-255]
    let r = parseInt(hexColor.slice(1, 3), 16);
    let g = parseInt(hexColor.slice(3, 5), 16);
    let b = parseInt(hexColor.slice(5, 7), 16);

    const clamp = (val) => Math.max(0, Math.min(255, Math.round(val)));

    // ───────────────────────────────────────────────
    // Helper: RGB → HSL (H in degrees [0-360], S/L in [0-1])
    // ───────────────────────────────────────────────
    function rgbToHsl(r, g, b) {
        r /= 255; g /= 255; b /= 255;
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h, s, l = (max + min) / 2;

        if (max === min) {
            h = s = 0;
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }
        return { h: h * 360, s, l };
    }

    // ───────────────────────────────────────────────
    // Helper: HSL → RGB
    // ───────────────────────────────────────────────
    function hslToRgb(h, s, l) {
        h /= 360;
        let r, g, b;

        if (s === 0) {
            r = g = b = l;
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1 / 6) return p + (q - p) * 6 * t;
                if (t < 1 / 2) return q;
                if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            };

            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }

        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }

    // Apply lumMod & lumOff first (they operate in luminance space)
    let { h, s, l } = rgbToHsl(r, g, b);

    // lumMod: scale luminance (100000 = 100% = no change)
    if (modifierNode["a:lumMod"]?.[0]?.$?.val) {
        const lumMod = parseInt(modifierNode["a:lumMod"][0].$.val, 10) / 100000;
        l *= lumMod;
    }

    // lumOff: add offset to luminance (positive = lighter, negative = darker)
    if (modifierNode["a:lumOff"]?.[0]?.$?.val) {
        const lumOff = parseInt(modifierNode["a:lumOff"][0].$.val, 10) / 100000;
        l += lumOff;
    }

    // Clamp luminance [0-1]
    l = Math.max(0, Math.min(1, l));

    // Convert back to RGB
    const rgb = hslToRgb(h, s, l);
    r = rgb.r;
    g = rgb.g;
    b = rgb.b;

    // Then apply other modifiers (tint, shade, satMod, etc.)
    if (modifierNode["a:shade"]?.[0]?.$?.val) {
        const shade = parseInt(modifierNode["a:shade"][0].$.val, 10) / 100000;
        r *= shade;
        g *= shade;
        b *= shade;
    }

    if (modifierNode["a:tint"]?.[0]?.$?.val) {
        const tint = parseInt(modifierNode["a:tint"][0].$.val, 10) / 100000;
        r = r + (255 - r) * tint;
        g = g + (255 - g) * tint;
        b = b + (255 - b) * tint;
    }

    if (modifierNode["a:satMod"]?.[0]?.$?.val) {
        const satMod = parseInt(modifierNode["a:satMod"][0].$.val, 10) / 100000;
        const gray = 0.299 * r + 0.587 * g + 0.114 * b;
        r = gray + (r - gray) * satMod;
        g = gray + (g - gray) * satMod;
        b = gray + (b - gray) * satMod;
    }

    // satOff skipped for now (rare in borders, can be added similarly in HSL)

    r = clamp(r);
    g = clamp(g);
    b = clamp(b);

    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

module.exports = { getShapeBorderCSS };