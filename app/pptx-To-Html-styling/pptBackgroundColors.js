
const path = require("path");
const fs = require("fs");
const colorHelper = require("../api/helper/colorHelper.js");
const config = require('../config.js');


// Define the directory to save images using config
const imageSavePath = path.resolve(__dirname, "../uploads");
if (!fs.existsSync(imageSavePath)) {
  fs.mkdirSync(imageSavePath, { recursive: true });
}

async function getBackgroundColor(slideXML, masterXML, themeXML, relationshipsXML, extractor, layoutXML) {
  try {
    const cSld = slideXML?.["p:sld"]?.["p:cSld"]?.[0];
    const bgPr = cSld?.["p:bg"]?.[0]?.["p:bgPr"]?.[0];

    // Check for solid fill with direct RGB color
    const solidFill = bgPr?.["a:solidFill"]?.[0];

    // Check for pattern fill
    const pattFill = bgPr?.["a:pattFill"]?.[0];
    if (pattFill) {
      const patternCSS = getPatternFillCSS(pattFill, themeXML);
      
      if (patternCSS) return { backgroundCSS: patternCSS };
    }

    // Check for picture or texture Fill
    const pictureTextureFill = bgPr?.["a:blipFill"]?.[0];

    if (pictureTextureFill) {
      // Extract fillRect values from stretch element if available
      const srcRect = pictureTextureFill?.["a:srcRect"]?.[0];
      const stretchElem = pictureTextureFill?.["a:stretch"]?.[0];
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

      // Get rotation info if available
      const rotWithShape = pictureTextureFill?.["$"]?.rotWithShape === "1";

      // Check for alphaModFix in the blip node
      const blip = pictureTextureFill?.["a:blip"]?.[0];
      let transparency = 0; // Default is 0% transparency (100% opacity)

      if (blip?.["a:alphaModFix"]?.[0]) {
        const alphaModFix = blip["a:alphaModFix"][0]["$"]?.amt;
        if (alphaModFix) {
          transparency = 100 - (parseInt(alphaModFix, 10) / 1000);
        }
      }

      // Get DPI setting if available
      const dpi = pictureTextureFill?.["$"]?.dpi || 0;

      const pictureTextureCSS = await getPictureTextureFillCSS(pictureTextureFill, relationshipsXML, extractor);

      if (pictureTextureCSS) {
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
          dpi: dpi
        };
      }
    }

    if (solidFill) {
      // Check for transparency in solid fill
      let transparency = 0;

      if (solidFill?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
        // Direct RGB color with alpha
        const alpha = solidFill["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
        if (alpha) {
          transparency = 100 - (parseInt(alpha, 10) / 1000);
        }

        return {
          backgroundCSS: `#${solidFill["a:srgbClr"][0]["$"]?.val}`,
          transparency: transparency
        };
      }
      else if (solidFill?.["a:srgbClr"]) {
        // Direct RGB color without alpha
        return {
          backgroundCSS: `#${solidFill["a:srgbClr"][0]["$"]?.val}`,
          transparency: 0
        };
      }

      // Check for theme color reference
      const schemeClr = solidFill?.["a:schemeClr"]?.[0];
      if (schemeClr) {
        const schemeClrVal = schemeClr["$"]?.val;

        // Check for alpha in theme color
        if (schemeClr?.["a:alpha"]?.[0]) {
          const alpha = schemeClr["a:alpha"][0]["$"]?.val;
          if (alpha) {
            transparency = 100 - (parseInt(alpha, 10) / 1000);
          }
        }

        const lumMod = schemeClr?.["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = schemeClr?.["a:lumOff"]?.[0]?.["$"]?.val;

        const baseColor = colorHelper.resolveThemeColorHelper(schemeClrVal, themeXML, masterXML);

        if (baseColor) {
          return {
            backgroundCSS: applyLuminanceModifier(baseColor, lumMod, lumOff),
            transparency: transparency
          };
        }
      }
    }

    // Check for gradient fill
    const gradFill = bgPr?.["a:gradFill"]?.[0];
    if (gradFill) {
      const gradientCSS = getGradientFillColor(gradFill, themeXML, masterXML);

      // Check for transparency in gradient stops
      let transparency = 0;
      const stops = gradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];
      for (const stop of stops) {
        // Check for alpha in srgbClr
        if (stop?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
          const alpha = stop["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
          if (alpha) {
            // Use the highest transparency value found
            const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
            transparency = Math.max(transparency, stopTransparency);
          }
        }

        // Check for alpha in schemeClr
        if (stop?.["a:schemeClr"]?.[0]?.["a:alpha"]?.[0]) {
          const alpha = stop["a:schemeClr"][0]["a:alpha"][0]["$"]?.val;
          if (alpha) {
            // Use the highest transparency value found
            const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
            transparency = Math.max(transparency, stopTransparency);
          }
        }
      }

      if (gradientCSS) {
        return {
          backgroundCSS: gradientCSS,
          transparency: transparency
        };
      }
    }

    // ===== NEW: Check layout background if no slide background found =====
    if (layoutXML) {

      const layoutCsld = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0];
      const layoutBgPr = layoutCsld?.["p:bg"]?.[0]?.["p:bgPr"]?.[0];

      if (layoutBgPr) {
        // Check for layout solid fill
        const layoutSolidFill = layoutBgPr?.["a:solidFill"]?.[0];

        if (layoutSolidFill) {
          // Check for transparency in layout solid fill
          let transparency = 0;

          if (layoutSolidFill?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
            // Direct RGB color with alpha
            const alpha = layoutSolidFill["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
            if (alpha) {
              transparency = 100 - (parseInt(alpha, 10) / 1000);
            }

            return {
              backgroundCSS: `#${layoutSolidFill["a:srgbClr"][0]["$"]?.val}`,
              transparency: transparency
            };
          }
          else if (layoutSolidFill?.["a:srgbClr"]) {
            // Direct RGB color without alpha
            return {
              backgroundCSS: `#${layoutSolidFill["a:srgbClr"][0]["$"]?.val}`,
              transparency: 0
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
              }
            }

            const lumMod = layoutSchemeClr?.["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = layoutSchemeClr?.["a:lumOff"]?.[0]?.["$"]?.val;

            const baseColor = colorHelper.resolveThemeColorHelper(schemeClrVal, themeXML, masterXML);

            if (baseColor) {
              return {
                backgroundCSS: applyLuminanceModifier(baseColor, lumMod, lumOff),
                transparency: transparency
              };
            }
          }
        }

        // Check for layout gradient fill
        const layoutGradFill = layoutBgPr?.["a:gradFill"]?.[0];
        if (layoutGradFill) {
          const gradientCSS = getGradientFillColor(layoutGradFill, themeXML, masterXML);

          // Check for transparency in layout gradient stops
          let transparency = 0;
          const stops = layoutGradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];
          for (const stop of stops) {
            // Check for alpha in srgbClr
            if (stop?.["a:srgbClr"]?.[0]?.["a:alpha"]?.[0]) {
              const alpha = stop["a:srgbClr"][0]["a:alpha"][0]["$"]?.val;
              if (alpha) {
                // Use the highest transparency value found
                const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
                transparency = Math.max(transparency, stopTransparency);
              }
            }

            // Check for alpha in schemeClr
            if (stop?.["a:schemeClr"]?.[0]?.["a:alpha"]?.[0]) {
              const alpha = stop["a:schemeClr"][0]["a:alpha"][0]["$"]?.val;
              if (alpha) {
                // Use the highest transparency value found
                const stopTransparency = 100 - (parseInt(alpha, 10) / 1000);
                transparency = Math.max(transparency, stopTransparency);
              }
            }
          }

          if (gradientCSS) {
            return {
              backgroundCSS: gradientCSS,
              transparency: transparency
            };
          }
        }

        // Check for layout pattern fill
        const layoutPattFill = layoutBgPr?.["a:pattFill"]?.[0];
        if (layoutPattFill) {
          const patternCSS = getPatternFillCSS(layoutPattFill, themeXML);
          if (patternCSS) {
            return { backgroundCSS: patternCSS };
          }
        }

        // Check for layout picture or texture Fill
        const layoutPictureTextureFill = layoutBgPr?.["a:blipFill"]?.[0];
        if (layoutPictureTextureFill) {
          const pictureTextureCSS = await getPictureTextureFillCSS(layoutPictureTextureFill, relationshipsXML, extractor);
          if (pictureTextureCSS) {
            return {
              backgroundCSS: pictureTextureCSS,
              transparency: 0
            };
          }
        }
      }
    }
    // ===== END NEW SECTION =====

    // Default fallback for no background
    return { backgroundCSS: "" };
  } catch (error) {
    console.error("Error resolving background color:", error);
    return { backgroundCSS: "" }; // Fallback in case of error
  }
}

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
      console.error(`Image file not found in extractor: ${normalizedPath}`);

      // Try alternative paths
      const alternativePath1 = `ppt/${targetPath}`;
      const alternativePath2 = targetPath.replace(/^\.\.\//, "");

      const altFile1 = extractor.files[alternativePath1];

      if (altFile1) {
        const imageBuffer = await altFile1.async("nodebuffer");
        return await saveAndReturnImageUrl(imageBuffer, targetPath);
      }

      const altFile2 = extractor.files[alternativePath2];

      if (altFile2) {
        const imageBuffer = await altFile2.async("nodebuffer");
        return await saveAndReturnImageUrl(imageBuffer, targetPath);
      }

      return null;
    }

    // If we found the image, process it
    const imageBuffer = await imageFile.async("nodebuffer");
    return await saveAndReturnImageUrl(imageBuffer, targetPath);

  } catch (error) {
    console.error(`Error processing texture fill image: ${error.message}`);
    console.error(error.stack); // Log the full stack trace for debugging
    return null;
  }
}

// Helper function to save image and return URL
async function saveAndReturnImageUrl(imageBuffer, originalPath) {
  const fileExtension = path.extname(originalPath);
  const baseFileName = path.basename(originalPath, fileExtension);
  const uniqueFileName = `${baseFileName}_${Date.now()}${fileExtension}`;

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

  console.log(`Image saved to: ${filePath}`);

  return `url('${imageUrl}')`;
}

function getPatternFillCSS(pattFill, themeXML) {
  const fgClrNode = pattFill?.["a:fgClr"]?.[0];
  const bgClrNode = pattFill?.["a:bgClr"]?.[0];

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

function resolveThemeColor(colorKey, themeXML) {
  const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) return null;

  const colorNode = clrScheme[`a:${colorKey}`]?.[0];
  if (!colorNode) return null;

  // Handle direct RGB color
  if (colorNode?.["a:srgbClr"]) {
    return `#${colorNode["a:srgbClr"][0]["$"]?.val}`;
  }

  // Handle system color
  if (colorNode?.["a:sysClr"]) {
    return `#${colorNode["a:sysClr"][0]["$"]?.lastClr || "FFFFFF"}`;
  }

  return null;
}


function getColorMappingFromMaster(masterXML) {
  const clrMapNode = masterXML?.["p:sldMaster"]?.[0]?.["p:clrMap"]?.[0]?.["$"];

  if (!clrMapNode) {
    return {
      bg1: "lt1",
      tx1: "dk1", // Default dark text
      bg2: "lt2",
      tx2: "dk2", // Secondary dark text
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hlink: "hlink",
      folHlink: "folHlink",
    };
  }

  return {
    bg1: clrMapNode.bg1 || "lt1",
    tx1: clrMapNode.tx1 || "dk1",
    bg2: clrMapNode.bg2 || "lt2",
    tx2: clrMapNode.tx2 || "dk2",
    accent1: clrMapNode.accent1 || "accent1",
    accent2: clrMapNode.accent2 || "accent2",
    accent3: clrMapNode.accent3 || "accent3",
    accent4: clrMapNode.accent4 || "accent4",
    accent5: clrMapNode.accent5 || "accent5",
    accent6: clrMapNode.accent6 || "accent6",
    hlink: clrMapNode.hlink || "hlink",
    folHlink: clrMapNode.folHlink || "folHlink",
  };
}


// function getShapeTransparency(shapeNode, themeXML) {
//   let fillColor = "#ffffff"; // Default black
//   let opacity = 1; // Fully visible

//   const shapeFill = shapeNode?.["p:spPr"]?.[0];

//   // 🔹 Check for Solid Fill
//   const solidFill = shapeFill?.["a:solidFill"]?.[0];

//   if (solidFill?.["a:srgbClr"]) {
//     fillColor = `#${solidFill["a:srgbClr"][0]["$"]?.val}`;
//   }
//   if (solidFill?.["a:schemeClr"]) {
//     fillColor = colorHelper.resolveThemeColorHelper(solidFill["a:schemeClr"][0]["$"]?.val, themeXML);

//   }
//   if (solidFill?.["a:alpha"]) {
//     opacity = parseInt(solidFill["a:alpha"][0]["$"].val, 10) / 100000;
//   }

//   // 🔹 Check for Gradient Fill (New Fix)
//   const gradFill = shapeFill?.["a:gradFill"]?.[0];
//   if (gradFill?.["a:gsLst"]) {
//     const stops = gradFill["a:gsLst"][0]["a:gs"];
//     let minAlpha = 100000; // Max value

//     stops.forEach(stop => {
//       const schemeClr = stop?.["a:schemeClr"]?.[0];
//       if (schemeClr?.["a:alpha"]) {
//         const alphaVal = parseInt(schemeClr["a:alpha"][0]["$"].val, 10);
//         minAlpha = Math.min(minAlpha, alphaVal);
//       }
//     });

//     if (minAlpha < 100000) {
//       opacity = minAlpha / 100000;
//     }
//   }

//   return { fillColor, opacity };
// }


module.exports = {
  getBackgroundColor,
  applyLuminanceModifier,
  getGradientFillColor,
  getPictureTextureFillCSS,
  resolveThemeColor,
  getColorMappingFromMaster,
  // getShapeTransparency
};
