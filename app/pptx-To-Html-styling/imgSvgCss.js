// Conversion functions
function emuToPx(emu) {
    return (emu * 0.0001052).toFixed(2); // Convert EMU to px and round to 2 decimal places
}

function pxToEmu(px) {
    return Math.round(px * 914400 / 96); // Convert px to EMU
}

// Conversion lookup tables
const emuToPxTable = {
    3175: 0.33,
    6350: 0.67,
    9525: 1.00,
    12700: 1.34,
    19050: 2.01,
    28575: 3.01,
    38100: 4.01,
    57150: 6.01,
    76200: 8.01,
};

const pxToEmuTable = {
    0.33: 3175,
    0.67: 6350,
    1.00: 9525,
    1.34: 12700,
    2.01: 19050,
    3.01: 28575,
    4.01: 38100,
    6.01: 57150,
    8.01: 76200,
};

// Helper function to fetch EMU-to-px conversion
function emuToPxWithLookup(emu) {
    return emuToPxTable[emu] || emuToPx(emu); // Use lookup or fallback
}


// Function to extract styles for images from PPTX XML
function returnImgSvgStyle(picNode) {

    let alphaValue = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0]?.["a:alphaModFix"]?.[0]?.["$"]?.amt;

    if (!alphaValue) {
        alphaValue = picNode?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0]?.["a:blip"]?.[0]?.["a:alphaModFix"]?.[0]?.["$"]?.amt || "100000" // getting 20000
    }

    const opacity = parseInt(alphaValue, 10) / 100000;

    const borderWidthEmu = picNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0]?.["$"]?.w || "0";
    const borderWidth = borderWidthEmu; // Assuming emuToPxWithLookup is a function you have defined elsewhere to convert EMU to pixels
    const borderColor = picNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0]?.["a:solidFill"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val || "#000000";

    const shadowColor = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val || "#000000";
    const shadowOffsetXEmu = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["$"]?.dx || "0";
    const shadowOffsetYEmu = picNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0]?.["$"]?.dy || "0";

    const shadowOffsetX = shadowOffsetXEmu; // Convert to px if necessary
    const shadowOffsetY = shadowOffsetYEmu; // Convert to px if necessary

    // Only add flip transformations if they are specified
    let flipH = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipH === "1" ? "scaleY(-1)" : "";
    let flipV = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipV === "1" ? "scaleX(-1)" : "";

    let rotation = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.rot;

    let rotationFlipH = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipH === "1";
    let rotationFlipV = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["$"]?.flipV === "1";

    if (rotation == "10800000" && rotationFlipH) {
        flipH = "scaleX(-1)";
    }

    if (rotation == "10800000" && rotationFlipV) {
        flipV = "scaleY(-1)";
    }
    if (!rotation && rotationFlipH) {
        flipH = "scaleX(-1)";
    }
    if (!rotation && rotationFlipV) {
        flipV = "scaleY(-1)";
    }

    return {
        opacity,
        borderWidth,
        borderColor,
        shadowColor,
        shadowOffsetX,
        shadowOffsetY,
        transform: [flipH, flipV].filter(t => t !== "").join(" "),
    };
}

// Export the updated function with the rest of the module
function getRotation(shapeNode) {
    const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"];
    return xfrm?.[0]?.["$"]?.rot ? parseInt(xfrm[0]["$"]?.rot, 10) / 60000 : 0;
}

module.exports = {
    returnImgSvgStyle,
    getRotation,
};
