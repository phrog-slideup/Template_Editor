/**
 * getShapeGlow.js
 * ---------------
 */

const colorHelper = require("../../api/helper/colorHelper.js");

// ─── Constants ────────────────────────────────────────────────────────────────

const EMU_PER_PX = 12700; // 96 dpi — same divisor used across shapeHandler.js

function getShapeGlowStyle(shapeNode, themeXML, masterXML, clrMap) {
    try {
        const glowNode = _extractGlowNode(shapeNode);
        if (!glowNode) return "";

        // ── 1. Radius ─────────────────────────────────────────────────────────
        const radEmu = parseInt(glowNode?.["$"]?.rad || "0", 10);
        if (isNaN(radEmu) || radEmu <= 0) return "";

        const radPx = radEmu / EMU_PER_PX; // e.g. 228600 / 12700 = 18 px

        // ── 2. Colour + alpha ─────────────────────────────────────────────────
        const { hex, alpha } = _resolveGlowColor(glowNode, themeXML, masterXML, clrMap);

        // ── 3. Build CSS ───────────────────────────────────────────────────────
        return _buildGlowCSS(hex, alpha, radPx);

    } catch (err) {
        console.error("[getShapeGlow] Error parsing glow effect:", err);
        return "";
    }
}

// ─── Private helpers ──────────────────────────────────────────────────────────

/**
 * Walks shapeNode → p:spPr → a:effectLst → a:glow
 * Returns the glow node or null.
 */
function _extractGlowNode(shapeNode) {
    return shapeNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:glow"]?.[0] ?? null;
}

function _resolveGlowColor(glowNode, themeXML, masterXML, clrMap) {
    // ── srgbClr ───────────────────────────────────────────────────────────────
    const srgbNode = glowNode?.["a:srgbClr"]?.[0];
    if (srgbNode) {
        const hex = `#${srgbNode["$"]?.val || "000000"}`;
        const alpha = _extractAlpha(srgbNode);
        return { hex, alpha };
    }

    // ── schemeClr ─────────────────────────────────────────────────────────────
    const schemeNode = glowNode?.["a:schemeClr"]?.[0];
    if (schemeNode) {
        const schemeVal = schemeNode["$"]?.val || "accent1";
        let hex = _resolveSchemeColor(schemeVal, themeXML, masterXML, clrMap);

        // Apply satMod if present — shifts HSL saturation
        const satMod = parseInt(schemeNode?.["a:satMod"]?.[0]?.["$"]?.val || "100000", 10);
        if (satMod !== 100000) {
            hex = _applySatMod(hex, satMod / 100000);
        }

        // Apply lumMod / lumOff if present
        const lumMod = schemeNode?.["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = schemeNode?.["a:lumOff"]?.[0]?.["$"]?.val;
        if (lumMod) {
            hex = _applyLuminance(hex, parseInt(lumMod, 10), parseInt(lumOff || "0", 10));
        }

        const alpha = _extractAlpha(schemeNode);
        return { hex, alpha };
    }

    // ── prstClr ───────────────────────────────────────────────────────────────
    const prstNode = glowNode?.["a:prstClr"]?.[0];
    if (prstNode) {
        const hex = _prstColorToHex(prstNode["$"]?.val || "black");
        const alpha = _extractAlpha(prstNode);
        return { hex, alpha };
    }

    // ── sysClr ────────────────────────────────────────────────────────────────
    const sysNode = glowNode?.["a:sysClr"]?.[0];
    if (sysNode) {
        const hex = `#${sysNode["$"]?.lastClr || "000000"}`;
        const alpha = _extractAlpha(sysNode);
        return { hex, alpha };
    }

    // Fallback
    return { hex: "#000000", alpha: 0.5 };
}

/**
 * Extracts <a:alpha val="73000"/> → 0.73
 * Returns 1.0 when not present.
 */
function _extractAlpha(colorNode) {
    const raw = colorNode?.["a:alpha"]?.[0]?.["$"]?.val;
    if (!raw) return 1.0;
    const parsed = parseInt(raw, 10);
    return isNaN(parsed) ? 1.0 : parsed / 100000;
}

/**
 * Resolves a DrawingML scheme colour name (e.g. "accent2") to a hex string
 * by walking the theme colour scheme, with clrMap remapping applied first.
 */
function _resolveSchemeColor(schemeVal, themeXML, masterXML, clrMap) {
    try {
        // Use colorHelper if available — it already handles clrMap + theme lookup
        if (colorHelper?.resolveThemeColorHelper) {
            const resolved = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, masterXML);
            if (resolved && resolved !== "#000000") return resolved;
        }

        // Manual fallback: walk theme XML directly
        const mappedKey = clrMap?.[schemeVal] || schemeVal;
        const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
        if (!clrScheme) return "#000000";

        const colorNode = clrScheme[`a:${mappedKey}`]?.[0];
        if (!colorNode) return "#000000";

        if (colorNode["a:srgbClr"]) return `#${colorNode["a:srgbClr"][0]["$"].val}`;
        if (colorNode["a:sysClr"]) return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;

        return "#000000";
    } catch {
        return "#000000";
    }
}

/**
 * Applies a saturation modifier to a hex colour.
 * satFactor: 0.0–2.0  (1.0 = no change, 1.75 = 175 % saturation)
 */
function _applySatMod(hex, satFactor) {
    try {
        const [r, g, b] = _hexToRgb(hex);
        const [h, s, l] = _rgbToHsl(r, g, b);
        const newS = Math.min(1, Math.max(0, s * satFactor));
        const [nr, ng, nb] = _hslToRgb(h, newS, l);
        return _rgbToHex(nr, ng, nb);
    } catch {
        return hex;
    }
}

/**
 * Applies lumMod (and optional lumOff) to a hex colour.
 * Both values are in the PPTX 0–100000 range.
 */
function _applyLuminance(hex, lumMod, lumOff) {
    try {
        const [r, g, b] = _hexToRgb(hex);
        const [h, s, l] = _rgbToHsl(r, g, b);
        // PPTX: newL = l * (lumMod/100000) + (lumOff/100000)
        const newL = Math.min(1, Math.max(0, l * (lumMod / 100000) + (lumOff / 100000)));
        const [nr, ng, nb] = _hslToRgb(h, s, newL);
        return _rgbToHex(nr, ng, nb);
    } catch {
        return hex;
    }
}

/**
 * Maps PPTX preset colour names to hex values.
 * Only the most common ones are listed; extend as needed.
 */
function _prstColorToHex(prstVal) {
    const map = {
        black: "#000000",
        white: "#FFFFFF",
        red: "#FF0000",
        green: "#00FF00",
        blue: "#0000FF",
        yellow: "#FFFF00",
        cyan: "#00FFFF",
        magenta: "#FF00FF",
        gray: "#808080",
        darkGray: "#A9A9A9",
        lightGray: "#D3D3D3",
        orange: "#FFA500",
        pink: "#FFC0CB",
        purple: "#800080",
        navy: "#000080",
        teal: "#008080",
        maroon: "#800000",
        olive: "#808000",
        silver: "#C0C0C0",
        gold: "#FFD700",
        coral: "#FF7F50",
        salmon: "#FA8072",
        violet: "#EE82EE",
        indigo: "#4B0082",
        brown: "#A52A2A",
        khaki: "#F0E68C",
        lavender: "#E6E6FA",
        lime: "#00FF00",
        aqua: "#00FFFF",
        fuchsia: "#FF00FF",
    };
    return map[prstVal] || "#000000";
}

/**
 * Builds the final CSS filter string.
 *
 */
function _buildGlowCSS(hex, alpha, radPx) {
    const [r, g, b] = _hexToRgb(hex);

    const REF_RAD = 18;
    const comp = Math.min(Math.sqrt(radPx / REF_RAD), 3.0);

    const blur1 = (radPx * 1.0).toFixed(2);   // immediate halo
    const blur2 = (radPx * 2.0).toFixed(2);   // fills spread distance
    const blur3 = (radPx * 3.5).toFixed(2);   // wide soft feather

    const a1 = Math.min(alpha * 0.65 * comp, 1.0).toFixed(4);
    const a2 = Math.min(alpha * 0.40 * comp, 1.0).toFixed(4);
    const a3 = Math.min(alpha * 0.18 * comp, 1.0).toFixed(4);

    const l1 = `drop-shadow(0 0 ${blur1}px rgba(${r},${g},${b},${a1}))`;
    const l2 = `drop-shadow(0 0 ${blur2}px rgba(${r},${g},${b},${a2}))`;
    const l3 = `drop-shadow(0 0 ${blur3}px rgba(${r},${g},${b},${a3}))`;

    return `filter: ${l1} ${l2} ${l3};`;
}

// ─── Colour math utilities ────────────────────────────────────────────────────

function _hexToRgb(hex) {
    const clean = hex.replace("#", "");
    return [
        parseInt(clean.slice(0, 2), 16),
        parseInt(clean.slice(2, 4), 16),
        parseInt(clean.slice(4, 6), 16),
    ];
}

function _rgbToHex(r, g, b) {
    return `#${[r, g, b].map(v => Math.round(v).toString(16).padStart(2, "0")).join("")}`;
}

/** RGB (0-255) → HSL (0-1 each) */
function _rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h;
    switch (max) {
        case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
        case g: h = ((b - r) / d + 2) / 6; break;
        default: h = ((r - g) / d + 4) / 6; break;
    }
    return [h, s, l];
}

/** HSL (0-1 each) → RGB (0-255) */
function _hslToRgb(h, s, l) {
    if (s === 0) {
        const v = Math.round(l * 255);
        return [v, v, v];
    }
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    const hue2rgb = (p, q, t) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    };
    return [
        Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
        Math.round(hue2rgb(p, q, h) * 255),
        Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
    ];
}

// ─── Exports ──────────────────────────────────────────────────────────────────

module.exports = { getShapeGlowStyle };