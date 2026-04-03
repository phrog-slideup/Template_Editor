function resolveThemeColorHelper(colorKey, themeXML, masterXML) {
    const clrMapNode = masterXML?.["p:sldMaster"]?.["p:clrMap"]?.[0];

    let mappedColorKey = colorKey;

    if (clrMapNode && clrMapNode['$']) {
        const dynamicColorMap = clrMapNode['$'];

        // If the color key exists in the map, use the mapped value
        if (dynamicColorMap[colorKey]) {
            mappedColorKey = dynamicColorMap[colorKey];
        }
    }

    // Get the color node using the mapped color key
    const colorNode = getThemeColor(themeXML, mappedColorKey);
    if (!colorNode) {
        console.warn("No colorNode found for the mapped color key:", mappedColorKey);
        return null;
    }

    return colorNode;
}


function getThemeColor(themeXML, schemeKey) {

    const colorNode = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0]?.[`a:${schemeKey}`]?.[0];

    if (!colorNode) return null; // FIX: explicit null instead of implicit undefined

    if (colorNode?.["a:sysClr"]) {
        return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;
    } else if (colorNode?.["a:srgbClr"]) {
        return `#${colorNode["a:srgbClr"][0]["$"].val}`;
    }

    return null; // FIX: explicit null fallback for unrecognized color node types
}

/**
 * Apply OOXML luminance modifiers to a hex color using the HSL colorspace.
 *
 * WHY HSL and not RGB:
 *   The OOXML spec defines lumMod/lumOff as operations on the HSL Luminance channel.
 *   Applying them directly to RGB channels produces wrong results for any non-gray
 *   color, and completely fails for black (#000000) because multiplying 0 by any
 *   lumMod always stays 0 — lumOff is then never applied, so black never brightens.
 *
 *   Example (tx1=dk1=#000000, lumMod=65000, lumOff=35000 → should be #595959):
 *     Old RGB approach:  0 * 0.65 = 0  →  #000000  ✗ (wrong — black stays black)
 *     New HSL approach:  L=0 → L*0.65+0.35=0.35 → #595959  ✓
 *
 * @param {string} hexColor   - Input color as #RRGGBB
 * @param {number} lumMod     - Luminance multiplier in OOXML units (100% = 100000)
 * @param {number} [lumOff=0] - Luminance additive offset in OOXML units (100% = 100000)
 * @returns {string} Modified color as #RRGGBB
 */
function applyLumMod(hexColor, lumMod, lumOff = 0) {
    if (!hexColor || !hexColor.startsWith('#') || hexColor.length < 7) return hexColor;

    let [r, g, b] = hexColor.substring(1).match(/.{2}/g).map(hex => parseInt(hex, 16));

    // Convert RGB → HSL
    const [h, s, l] = _rgbToHsl(r, g, b);

    // Apply OOXML luminance modifiers on the L channel
    let newL = l * (lumMod / 100000) + (lumOff / 100000);
    newL = Math.max(0, Math.min(1, newL)); // Clamp to [0, 1]

    // Convert HSL → RGB
    const [nr, ng, nb] = _hslToRgb(h, s, newL);

    return `#${Math.round(nr).toString(16).padStart(2, '0')}${Math.round(ng).toString(16).padStart(2, '0')}${Math.round(nb).toString(16).padStart(2, '0')}`;
}


// In colorHelper.js - add this if missing:
function applyLumOff(hexColor, lumOffVal) {
    // lumOff adds to luminance in HLS space
    // lumOffVal is in 1/1000ths of a percent (e.g. 35000 = 35%)
    const lumOffPct = parseInt(lumOffVal, 10) / 100000;

    // Convert hex to RGB
    const r = parseInt(hexColor.slice(1, 3), 16);
    const g = parseInt(hexColor.slice(3, 5), 16);
    const b = parseInt(hexColor.slice(5, 7), 16);

    // Add luminance offset (clamp to 0-255)
    const newR = Math.min(255, Math.round(r + 255 * lumOffPct));
    const newG = Math.min(255, Math.round(g + 255 * lumOffPct));
    const newB = Math.min(255, Math.round(b + 255 * lumOffPct));

    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}


/** @private RGB [0-255] → HSL [0-1] */
function _rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
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
    return [h, s, l];
}

/** @private HSL [0-1] → RGB [0-255] */
function _hslToRgb(h, s, l) {
    if (s === 0) {
        const v = Math.round(l * 255);
        return [v, v, v];
    }
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
    return [
        hue2rgb(p, q, h + 1 / 3) * 255,
        hue2rgb(p, q, h) * 255,
        hue2rgb(p, q, h - 1 / 3) * 255,
    ];
}

// Function to resolve preset colors like "white", "black", "red", etc.
function resolvePresetColor(prstClr) {
    const presetColors = {
        white: "#FFFFFF",
        black: "#000000",
        red: "#FF0000",
        blue: "#0000FF",
        green: "#008000",
        yellow: "#FFFF00",
        gray: "#808080",
        orange: "#FFA500",
        purple: "#800080",
        cyan: "#00FFFF",
        magenta: "#FF00FF",
        brown: "#A52A2A",
        pink: "#FFC0CB",
        lime: "#00FF00",
        navy: "#000080",
        maroon: "#800000",
        olive: "#808000",
        silver: "#C0C0C0",
        teal: "#008080",
    };

    return presetColors[prstClr.toLowerCase()] || "#000000"; // Default to black if not found
}

/**
 * Utility: Convert RGB or RGBA color to Hex
 */

function rgbToHex(color) {

    if (!color || typeof color !== "string") return "#000000"; // Default to black

    // If color is already in hex format, just return it
    if (color.startsWith('#') && (color.length === 7 || color.length === 9)) {
        return color;
    }

    if (color.includes("rgba") && color.includes("0, 0, 0, 0")) {
        return "#FFFFFF"; // Default to white for transparent
    }

    const result = color.match(/\d+/g);
    if (result && result.length >= 3) {
        const r = parseInt(result[0]).toString(16).padStart(2, "0");
        const g = parseInt(result[1]).toString(16).padStart(2, "0");
        const b = parseInt(result[2]).toString(16).padStart(2, "0");
        return `#${r}${g}${b}`;
    }

    return "#FFFFFF"; // Default to white for invalid format
}

function normalizeStyleValue(value, defaultValue = 0) {
    return parseFloat(value || `${defaultValue}`.replace("px", "")) || defaultValue;
}

module.exports = {
    applyLumMod,
    resolvePresetColor,
    rgbToHex,
    normalizeStyleValue,
    resolveThemeColorHelper,
    applyLumOff
}