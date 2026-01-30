
function resolveThemeColorHelper(colorKey, themeXML, masterXML) {
    const clrMapNode = masterXML?.["p:sldMaster"]?.["p:clrMap"]?.[0];

    // Log clrMapNode to check its structure
    // console.log("clrMapNode:", clrMapNode);

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
        console.log("No colorNode found for the mapped color key:", mappedColorKey);
        return "";
    }

    return colorNode;
}


function getThemeColor(themeXML, schemeKey) {

    const colorNode = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0]?.[`a:${schemeKey}`]?.[0];

    if (colorNode?.["a:sysClr"]) {
        return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;
    } else if (colorNode?.["a:srgbClr"]) {

        return `#${colorNode["a:srgbClr"][0]["$"].val}`;
    }
}

function applyLumMod(hexColor, lumPercent) {
    let [r, g, b] = hexColor.substring(1).match(/.{2}/g).map(hex => parseInt(hex, 16));
    r = Math.round(r * lumPercent / 100000);
    g = Math.round(g * lumPercent / 100000);
    b = Math.round(b * lumPercent / 100000);
    r = Math.min(255, r);
    g = Math.min(255, g);
    b = Math.min(255, b);
    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
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
}