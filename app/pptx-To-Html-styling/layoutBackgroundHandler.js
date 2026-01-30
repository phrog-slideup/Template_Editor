const path = require("path");
const fs = require("fs");
const colorHelper = require("../api/helper/colorHelper.js");
const config = require('../config.js');

// Define the directory to save images using config
const imageSavePath = path.resolve(__dirname, "../uploads");
if (!fs.existsSync(imageSavePath)) {
    fs.mkdirSync(imageSavePath, { recursive: true });
}

async function getLayoutBackgroundColor(layoutXML, themeXML, relationshipsXML, masterXML, extractor) {
    try {

        // console.log("slideLayoutBg ==========----------------111111------- >>>>>>>>>>>>>> ", layoutXML);

        if (!layoutXML) {
            console.log("No layout XML provided");
            return { backgroundCSS: "" };
        }

        // console.log("Checking layout background...");

        const layoutCsld = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0];
        const layoutBgPr = layoutCsld?.["p:bg"]?.[0]?.["p:bgPr"]?.[0];

        if (!layoutBgPr) {
            // console.log("No layout background properties found");
            return { backgroundCSS: "" };
        }

        // Check for layout solid fill
        const layoutSolidFill = layoutBgPr?.["a:solidFill"]?.[0];
        if (layoutSolidFill) {
            const solidFillResult = await processLayoutSolidFill(layoutSolidFill, themeXML, masterXML);
            if (solidFillResult) {
                return solidFillResult;
            }
        }

        // Check for layout gradient fill
        const layoutGradFill = layoutBgPr?.["a:gradFill"]?.[0];
        if (layoutGradFill) {
            const gradientResult = await processLayoutGradientFill(layoutGradFill, themeXML, masterXML);
            if (gradientResult) {
                return gradientResult;
            }
        }

        // Check for layout pattern fill
        const layoutPattFill = layoutBgPr?.["a:pattFill"]?.[0];
        if (layoutPattFill) {
            const patternResult = await processLayoutPatternFill(layoutPattFill, themeXML, masterXML);
            if (patternResult) {
                return patternResult;
            }
        }

        // Check for layout picture or texture fill
        const layoutPictureTextureFill = layoutBgPr?.["a:blipFill"]?.[0];

        // console.log("layoutPictureTextureFill ==========----##########-----111111------- >>>>>>>>>>>>>> ", layoutPictureTextureFill);

        if (layoutPictureTextureFill) {
            const pictureResult = await processLayoutPictureFill(layoutPictureTextureFill, relationshipsXML, extractor);
            if (pictureResult) {
                return pictureResult;
            }
        }

        // Check for background reference (inherits from master)
        const layoutBgRef = layoutCsld?.["p:bg"]?.[0]?.["p:bgRef"]?.[0];
        if (layoutBgRef) {
            const bgRefResult = await processLayoutBackgroundReference(layoutBgRef, masterXML, themeXML);
            if (bgRefResult) {
                return bgRefResult;
            }
        }

        console.log("No layout background found");
        return { backgroundCSS: "" };

    } catch (error) {
        console.error("Error resolving layout background color:", error);
        return { backgroundCSS: "" };
    }
}

async function processLayoutSolidFill(layoutSolidFill, themeXML, masterXML) {
    try {
        let transparency = 0;

        // Check for direct RGB color with alpha
        if (layoutSolidFill?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
            const alpha = layoutSolidFill["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
            if (alpha) {
                transparency = 100 - (parseInt(alpha, 10) / 1000);
                console.log(`Found layout solid fill alpha: ${alpha}, calculated transparency: ${transparency}%`);
            }

            const color = `#${layoutSolidFill["a:srgbClr"][0]["$"]?.val}`;
            console.log(`Layout solid fill with alpha: ${color}, transparency: ${transparency}%`);

            return {
                backgroundCSS: color,
                transparency: transparency,
                source: 'layout-solid-fill-alpha'
            };
        }
        // Check for direct RGB color without alpha
        else if (layoutSolidFill?.["a:srgbClr"]) {
            const color = `#${layoutSolidFill["a:srgbClr"][0]["$"]?.val}`;
            console.log(`Found layout solid fill: ${color}`);

            return {
                backgroundCSS: color,
                transparency: 0,
                source: 'layout-solid-fill'
            };
        }

        // Check for layout theme color reference
        const layoutSchemeClr = layoutSolidFill?.["a:schemeClr"]?.[0];
        if (layoutSchemeClr) {
            const schemeClrVal = layoutSchemeClr["$"]?.val;

            // Check for alpha in layout theme color
            if (layoutSchemeClr?.["a:alpha"]?.[0]) {
                const alpha = layoutSchemeClr["a:alpha"][0]["$"]?.val;
                if (alpha) {
                    transparency = 100 - (parseInt(alpha, 10) / 1000);
                    console.log(`Found layout scheme color alpha: ${alpha}, calculated transparency: ${transparency}%`);
                }
            }

            const lumMod = layoutSchemeClr?.["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = layoutSchemeClr?.["a:lumOff"]?.[0]?.["$"]?.val;

            const baseColor = colorHelper.resolveThemeColorHelper(schemeClrVal, themeXML, masterXML);

            console.log(`schemeClrVal =========>>>> ${schemeClrVal}, Found layout scheme color: ${baseColor}, transparency: ${transparency}%`);

            if (baseColor) {
                const finalColor = applyLuminanceModifier(baseColor, lumMod, lumOff);
                console.log(`Layout scheme color: ${finalColor}, transparency: ${transparency}%`);

                return {
                    backgroundCSS: finalColor,
                    transparency: transparency,
                    source: 'layout-scheme-color'
                };
            }
        }

        return null;
    } catch (error) {
        console.error("Error processing layout solid fill:", error);
        return null;
    }
}


async function processLayoutGradientFill(layoutGradFill, themeXML, masterXML) {
    try {
        const gradientCSS = getGradientFillColor(layoutGradFill, themeXML, masterXML);

        // Check for transparency in layout gradient stops
        let transparency = 0;
        const stops = layoutGradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];

        for (const stop of stops) {
            // Check for alpha in srgbClr
            if (stop?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
                const alpha = stop["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
                if (alpha) {
                    const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
                    transparency = Math.max(transparency, stopTransparency);
                }
            }

            // Check for alpha in schemeClr
            if (stop?.["a:schemeClr"]?.[0]?.["a:alpha"]?.[0]) {
                const alpha = stop["a:schemeClr"][0]["a:alpha"][0]["$"]?.val;
                if (alpha) {
                    const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
                    transparency = Math.max(transparency, stopTransparency);
                }
            }
        }

        if (gradientCSS) {
            console.log(`Found layout gradient fill: ${gradientCSS}`);
            return {
                backgroundCSS: gradientCSS,
                transparency: transparency,
                source: 'layout-gradient-fill'
            };
        }

        return null;
    } catch (error) {
        console.error("Error processing layout gradient fill:", error);
        return null;
    }
}

async function processLayoutPatternFill(layoutPattFill, themeXML, masterXML) {
    try {
        const patternCSS = getPatternFillCSS(layoutPattFill, themeXML, masterXML);

        if (patternCSS) {
            console.log(`Found layout pattern fill: ${patternCSS}`);
            return {
                backgroundCSS: patternCSS,
                source: 'layout-pattern-fill'
            };
        }

        return null;
    } catch (error) {
        console.error("Error processing layout pattern fill:", error);
        return null;
    }
}

async function processLayoutPictureFill(layoutPictureTextureFill, relationshipsXML, extractor) {
    try {
        // Extract fillRect values from stretch element if available
        const srcRect = layoutPictureTextureFill?.["a:srcRect"]?.[0];
        const stretchElem = layoutPictureTextureFill?.["a:stretch"]?.[0];
        const fillRectElem = stretchElem?.["a:fillRect"]?.[0];

        // Get inset values (negative values mean the image extends beyond the slide)
        const topInsetEMU = fillRectElem?.["$"]?.t || 0;
        const bottomInsetEMU = fillRectElem?.["$"]?.b || 0;
        const leftInsetEMU = fillRectElem?.["$"]?.l || 0;
        const rightInsetEMU = fillRectElem?.["$"]?.r || 0;

        // Convert EMUs to pixels (12700 EMUs = 1px)
        const topInset = parseInt(topInsetEMU, 10) / 12700;
        const bottomInset = parseInt(bottomInsetEMU, 10) / 12700;
        const leftInset = parseInt(leftInsetEMU, 10) / 12700;
        const rightInset = parseInt(rightInsetEMU, 10) / 12700;

        console.log(`Layout background insets: top=${topInset}px, bottom=${bottomInset}px, left=${leftInset}px, right=${rightInset}px`);

        // Get rotation info if available
        const rotWithShape = layoutPictureTextureFill?.["$"]?.rotWithShape === "1";

        // Check for alphaModFix in the blip node
        const blip = layoutPictureTextureFill?.["a:blip"]?.[0];
        let transparency = 0; // Default is 0% transparency (100% opacity)

        if (blip?.["a:alphaModFix"]?.[0]) {
            const alphaModFix = blip["a:alphaModFix"][0]["$"]?.amt;
            if (alphaModFix) {
                transparency = 100 - (parseInt(alphaModFix, 10) / 1000);
            }
        }

        const pictureTextureCSS = await getPictureTextureFillCSS(layoutPictureTextureFill, relationshipsXML, extractor);

        if (pictureTextureCSS) {
            console.log(`Found layout picture fill: ${pictureTextureCSS}`);
            return {
                backgroundCSS: pictureTextureCSS,
                transparency: transparency,
                insets: {
                    top: topInset,
                    bottom: bottomInset,
                    left: leftInset,
                    right: rightInset
                },
                rotWithShape: rotWithShape,
                source: 'layout-picture-fill'
            };
        }

        return null;
    } catch (error) {
        console.error("Error processing layout picture fill:", error);
        return null;
    }
}

async function processLayoutBackgroundReference(layoutBgRef, masterXML, themeXML) {
    try {
        const bgRefIndex = layoutBgRef?.["$"]?.idx;
        const schemeClr = layoutBgRef?.["a:schemeClr"]?.[0];

        if (schemeClr) {
            const schemeClrVal = schemeClr["$"]?.val;
            const baseColor = colorHelper.resolveThemeColorHelper(schemeClrVal, themeXML, masterXML);

            if (baseColor) {
                console.log(`Layout background reference to scheme color: ${baseColor}`);
                return {
                    backgroundCSS: baseColor,
                    transparency: 0,
                    source: 'layout-bg-reference'
                };
            }
        }

        // If no scheme color, try to get from master based on index
        if (bgRefIndex && masterXML) {
            // This would need implementation based on your master structure
            console.log(`Layout references master background index: ${bgRefIndex}`);
        }

        return null;
    } catch (error) {
        console.error("Error processing layout background reference:", error);
        return null;
    }
}

function hasLayoutBackground(layoutXML) {
    if (!layoutXML) return false;

    const layoutCsld = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0];
    const layoutBg = layoutCsld?.["p:bg"]?.[0];

    return !!(layoutBg?.["p:bgPr"] || layoutBg?.["p:bgRef"]);
}

function getLayoutBackgroundType(layoutXML) {
    if (!layoutXML) return 'none';

    const layoutCsld = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0];
    const layoutBgPr = layoutCsld?.["p:bg"]?.[0]?.["p:bgPr"]?.[0];

    if (!layoutBgPr) {
        const layoutBgRef = layoutCsld?.["p:bg"]?.[0]?.["p:bgRef"]?.[0];
        return layoutBgRef ? 'reference' : 'none';
    }

    if (layoutBgPr?.["a:solidFill"]) return 'solid';
    if (layoutBgPr?.["a:gradFill"]) return 'gradient';
    if (layoutBgPr?.["a:pattFill"]) return 'pattern';
    if (layoutBgPr?.["a:blipFill"]) return 'picture';

    return 'unknown';
}

// Helper functions (imported from main background handler)

function applyLuminanceModifier(hexColor, lumMod, lumOff) {
    if (!hexColor || !hexColor.startsWith('#') || hexColor.length !== 7) return hexColor;

    // Convert hex to RGB
    let r = parseInt(hexColor.substring(1, 3), 16);
    let g = parseInt(hexColor.substring(3, 5), 16);
    let b = parseInt(hexColor.substring(5, 7), 16);

    // Convert RGB to HSL
    const [h, s, l] = rgbToHsl(r, g, b);

    // Apply luminance modifications
    let newL = l;

    // According to OpenXML spec, lumMod is applied first as a multiplier
    if (lumMod) {
        newL = newL * (lumMod / 100000);
    }

    // Then lumOff is applied as an additive value
    if (lumOff) {
        newL = newL + (lumOff / 100000);
    }

    // Clamp luminance to valid range [0, 1]
    newL = Math.max(0, Math.min(1, newL));

    // Convert back to RGB
    const [newR, newG, newB] = hslToRgb(h, s, newL);

    // Convert back to hex
    return `#${Math.round(newR).toString(16).padStart(2, "0")}${Math.round(newG).toString(16).padStart(2, "0")}${Math.round(newB).toString(16).padStart(2, "0")}`;
}

// Helper function to convert RGB to HSL
function rgbToHsl(r, g, b) {
    r /= 255;
    g /= 255;
    b /= 255;

    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h, s, l = (max + min) / 2;

    if (max === min) {
        h = s = 0; // achromatic
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

// Helper function to convert HSL to RGB
function hslToRgb(h, s, l) {
    let r, g, b;

    if (s === 0) {
        r = g = b = l; // achromatic
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

    return [r * 255, g * 255, b * 255];
}

function getGradientFillColor(gradFill, themeXML, masterXML) {
    const stops = gradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];

    // Extract gradient angle or default to 90 degrees (vertical gradient)
    const angle = parseInt(gradFill?.["a:lin"]?.[0]?.["$"]?.ang, 10) || 5400000; // Default angle: 90 degrees
    const gradientAngle = angle / 60000; // Convert angle to degrees

    let gradientStops = [];

    stops.forEach((stop) => {
        const position = parseInt(stop["$"]?.pos, 10) / 1000; // Convert position to percentage
        const schemeClr = stop?.["a:schemeClr"]?.[0]?.["$"]?.val;
        const srgbClr = stop?.["a:srgbClr"]?.[0]?.["$"]?.val;

        let color = null;

        if (srgbClr) {
            color = `#${srgbClr}`; // Direct RGB color
        } else if (schemeClr) {
            let baseColor = colorHelper.resolveThemeColorHelper(schemeClr, themeXML, masterXML);

            // Apply luminance modifiers
            const lumMod = stop?.["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = stop?.["a:schemeClr"]?.[0]?.["a:lumOff"]?.[0]?.["$"]?.val;
            color = applyLuminanceModifier(baseColor, lumMod, lumOff);
        }

        if (color) {
            gradientStops.push(`${color} ${position}%`);
        }
    });

    // Build CSS gradient
    if (gradientStops.length > 0) {
        const stopsCSS = gradientStops.join(", ");
        return `linear-gradient(${gradientAngle}deg, ${stopsCSS})`;
    }

    return null; // Default if no gradient stops
}

async function getPictureTextureFillCSS(blipFill, relationshipsXML, extractor) {
    if (!relationshipsXML || !relationshipsXML.Relationships) {
        console.error("Invalid or missing relationships XML.");
        return null;
    }

    const embedId = blipFill?.["a:blip"]?.[0]?.["$"]?.["r:embed"];
    if (!embedId) return null;

    const relationship = relationshipsXML.Relationships.Relationship.find(
        (rel) => rel["$"]?.Id === embedId
    );

    if (!relationship) {
        console.error(`Relationship not found for embedId: ${embedId}`);
        return null;
    }

    try {
        // Get the relationship target path
        const targetPath = relationship["$"]?.Target;

        if (!targetPath) {
            console.error("Target path not found in relationship");
            return null;
        }

        // Normalize the path to handle relative references (like "../media/image1.png")
        const normalizedPath = targetPath.replace(/^\.\.\//, "ppt/");

        // Check if extractor is valid
        if (!extractor) {
            console.error("Extractor is undefined or null");
            return null;
        }

        // Get the image file from the extractor
        const imageFile = extractor.files[normalizedPath];

        if (!imageFile) {
            console.error(`Layout image file not found in extractor: ${normalizedPath}`);

            // Try alternative paths
            const alternativePath1 = `ppt/${targetPath}`;
            const alternativePath2 = targetPath.replace(/^\.\.\//, "");

            console.log("Trying alternative path 1:", alternativePath1);
            const altFile1 = extractor.files[alternativePath1];

            if (altFile1) {
                console.log("Found layout image at alternative path 1");
                const imageBuffer = await altFile1.async("nodebuffer");
                return await saveAndReturnImageUrl(imageBuffer, targetPath, 'layout');
            }

            console.log("Trying alternative path 2:", alternativePath2);
            const altFile2 = extractor.files[alternativePath2];

            if (altFile2) {
                console.log("Found layout image at alternative path 2");
                const imageBuffer = await altFile2.async("nodebuffer");
                return await saveAndReturnImageUrl(imageBuffer, targetPath, 'layout');
            }

            return null;
        }

        // If we found the image, process it
        const imageBuffer = await imageFile.async("nodebuffer");
        return await saveAndReturnImageUrl(imageBuffer, targetPath, 'layout');

    } catch (error) {
        console.error(`Error processing layout texture fill image: ${error.message}`);
        console.error(error.stack);
        return null;
    }
}

// Helper function to save layout image and return URL
async function saveAndReturnImageUrl(imageBuffer, originalPath, prefix = 'layout') {
    const fileExtension = path.extname(originalPath);
    const baseFileName = path.basename(originalPath, fileExtension);
    const uniqueFileName = `${prefix}_${baseFileName}_${Date.now()}${fileExtension}`;

    // Use the specific "../uploads" directory
    const uploadsDir = path.resolve(__dirname, "../uploads");

    // Ensure the uploads directory exists
    if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
    }

    const filePath = path.join(uploadsDir, uniqueFileName);

    // Save the image to disk
    fs.writeFileSync(filePath, imageBuffer);

    // Generate the URL for the saved image
    const imageUrl = `/uploads/${uniqueFileName}`;

    console.log(`Layout image saved to: ${filePath}`);
    console.log(`Layout image URL: ${imageUrl}`);

    return `url('${imageUrl}')`;
}

function getPatternFillCSS(pattFill, themeXML, masterXML) {
    const fgClrNode = pattFill?.["a:fgClr"]?.[0];
    const bgClrNode = pattFill?.["a:bgClr"]?.[0];

    // Resolve foreground color
    let fgColor = null;
    if (fgClrNode?.["a:srgbClr"]) {
        fgColor = `#${fgClrNode["a:srgbClr"][0]["$"]?.val}`;
    } else if (fgClrNode?.["a:schemeClr"]) {
        const schemeVal = fgClrNode["a:schemeClr"][0]["$"]?.val;
        fgColor = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, masterXML);
    }

    // Resolve background color
    let bgColor = null;
    if (bgClrNode?.["a:srgbClr"]) {
        bgColor = `#${bgClrNode["a:srgbClr"][0]["$"]?.val}`;
    } else if (bgClrNode?.["a:schemeClr"]) {
        const schemeVal = bgClrNode["a:schemeClr"][0]["$"]?.val;
        bgColor = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, masterXML);
    }

    // Map preset patterns to CSS equivalents
    const patternMap = {
        // Percentage Patterns
        pct5: "5%",
        pct10: "10%",
        pct20: "20%",
        pct25: "25%",
        pct30: "30%",
        pct40: "40%",
        pct50: "50%",
        pct60: "60%",
        pct70: "70%",
        pct80: "80%",
        pct90: "90%",

        // Line-Based Patterns
        horzStripe: "horizontal-stripes",
        thinHorzStripe: "thin-horizontal-stripes",
        vertStripe: "vertical-stripes",
        thinVertStripe: "thin-vertical-stripes",
        diagStripe: "diagonal-stripes",
        thinDiagStripe: "thin-diagonal-stripes",
        zigZag: "zigzag-lines",

        // Grid Patterns
        dkGrid: "dark-grid",
        ltGrid: "light-grid",
        squareGrid: "square-grid",
        hexGrid: "hexagonal-grid",
        dottedGrid: "dotted-grid",

        // Cross Patterns
        dkDiagCross: "dark-diagonal-cross",
        ltDiagCross: "light-diagonal-cross",
        solidCross: "solid-cross",

        // Special Patterns
        weave: "weave-pattern",
        wave: "wave-pattern",
    };

    const pattern = pattFill?.["$"]?.prst || "pct5"; // Default to 5% pattern
    const cssPattern = patternMap[pattern] || "none";

    // CSS generation for patterns
    if (fgColor && bgColor && cssPattern) {
        return `
      background-color: ${bgColor};
      background-image: repeating-linear-gradient(${cssPattern}, ${fgColor} 0%, ${fgColor} 50%, ${bgColor} 50%, ${bgColor} 100%);
    `;
    }

    return null;
}

module.exports = {
    getLayoutBackgroundColor,
    processLayoutSolidFill,
    processLayoutGradientFill,
    processLayoutPatternFill,
    processLayoutPictureFill,
    processLayoutBackgroundReference,
    hasLayoutBackground,
    getLayoutBackgroundType,
    applyLuminanceModifier
};