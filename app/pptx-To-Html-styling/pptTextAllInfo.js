const colorHelper = require("../api/helper/colorHelper.js");

// Add a helper function to determine the correct divisor
function getEMUDivisor() {
  // return parseInt(flag) === 1 ? 9525 : 12700;
  return 12700;
}

function getPositionFromShape(shapeNode) {
  const divisor = getEMUDivisor();

  const nvPr = shapeNode?.["p:nvPicPr"]?.[0]?.["p:nvPr"]?.[0];
  const placeholder = nvPr?.["p:ph"]?.[0]?.["$"];

  // ✅ CORRECT srcRect PATH    ?.["a:srcRect"]?.[0]?.["$"]
  const srcRectNode = shapeNode?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"];
  // const srcRectNode = shapeNode?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0]?.["a:stretch"]?.[0];
  let srcRectL = 0,
    srcRectR = 0,
    srcRectT = 0,
    srcRectB = 0;

  if (srcRectNode) {
    srcRectL = srcRectNode.l ? parseInt(srcRectNode.l) : 0;
    srcRectR = srcRectNode.r ? parseInt(srcRectNode.r) : 0;
    srcRectT = srcRectNode.t ? parseInt(srcRectNode.t) : 0;
    srcRectB = srcRectNode.b ? parseInt(srcRectNode.b) : 0;
  }

  let phType = placeholder?.type || null;
  let phIdx = placeholder?.idx || null;

  const x = Math.round((shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:off"]?.[0]?.["$"]?.x || 0) / divisor);
  const y = Math.round((shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:off"]?.[0]?.["$"]?.y || 0) / divisor);
  const width = Math.round((shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:ext"]?.[0]?.["$"]?.cx || 100) / divisor);
  const height = Math.round((shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0]?.["a:ext"]?.[0]?.["$"]?.cy || 100) / divisor);

  const picZIndex = (shapeNode?.["p:nvPicPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.id || 0);

  return { x, y, width, height, picZIndex, phType, phIdx, srcRectL, srcRectR, srcRectT, srcRectB };
}

function getPositionFromPlaceholderLayout(picNode, layoutXML, flag = 0) {
  try {
    const divisor = getEMUDivisor();

    const nvPr = picNode?.["p:nvPicPr"]?.[0]?.["p:nvPr"]?.[0];
    const placeholder = nvPr?.["p:ph"]?.[0]?.["$"];

    // Extract srcRectL and srcRectR if available
    const srcRectL = picNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.l || 0;
    const srcRectR = picNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.r || 0;
    const srcRectT = picNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.t || 0;
    const srcRectB = picNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.b || 0;

    let phType = placeholder?.type || null;
    let phIdx = placeholder?.idx || null;

    if (!placeholder) {
      return null; // Not a placeholder
    }

    // Find matching placeholder in layout
    const layoutShapes = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];

    for (const layoutShape of layoutShapes) {
      const layoutNvPr = layoutShape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0];
      const layoutPh = layoutNvPr?.["p:ph"]?.[0]?.["$"];

      if (!layoutPh) continue;

      // Match by type and idx
      if (layoutPh.type === placeholder.type && layoutPh.idx === placeholder.idx) {
        console.log('Found matching layout placeholder!');

        // Extract position from layout
        const layoutXfrm = layoutShape?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];

        if (layoutXfrm) {
          const offset = layoutXfrm?.["a:off"]?.[0]?.["$"];
          const extent = layoutXfrm?.["a:ext"]?.[0]?.["$"];

          if (offset && extent) {
            const x = parseInt(offset.x || 0, 10) / divisor;
            const y = parseInt(offset.y || 0, 10) / divisor;
            const width = parseInt(extent.cx || 0, 10) / divisor;
            const height = parseInt(extent.cy || 0, 10) / divisor;
            const picZIndex = picNode?.["p:nvPicPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.id || 0;

            return { x, y, width, height, picZIndex, phType, phIdx, srcRectL, srcRectR, srcRectT, srcRectB };
          }
        }
      }
    }

    return null; // No matching placeholder found

  } catch (error) {
    console.error('Error extracting placeholder position:', error);
    return null;
  }
}

// ADD: Enhanced function that tries both methods
function getPositionFromShapeOrPlaceholder(shapeNode, layoutXML = null, flag = 0) {
  // Try regular extraction first

  const cropData = getCroppingData(shapeNode);

  const regularPosition = getPositionFromShape(shapeNode, flag);

  if (cropData) {
    regularPosition.phIdx;
    regularPosition.phType;
    regularPosition.srcRectL;
    regularPosition.srcRectR;
    regularPosition.srcRectT;
    regularPosition.srcRectB;
  }

  // If regular extraction gives valid dimensions, use it
  if (regularPosition && regularPosition.width > 1 && regularPosition.height > 1) {

    return regularPosition;
  }

  // If we have layoutXML, try placeholder extraction
  if (layoutXML) {
    const placeholderPosition = getPositionFromPlaceholderLayout(shapeNode, layoutXML, flag);

    if (placeholderPosition && placeholderPosition.width > 1) {
      return placeholderPosition;
    }
  }

  // Fallback
  return { x: 0, y: 0, width: 100, height: 100, picZIndex: 0 };
}

function getCroppingData(shapeNode) {

  const nvPr = shapeNode?.["p:nvPicPr"]?.[0]?.["p:nvPr"]?.[0];
  const placeholder = nvPr?.["p:ph"]?.[0]?.["$"];

  // Extract srcRectL and srcRectR if available
  const srcRectL = shapeNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.l || 0;
  const srcRectR = shapeNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.r || 0;
  const srcRectT = shapeNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.t || 0;
  const srcRectB = shapeNode?.["p:blipFill"]?.[0]?.["a:srcRect"]?.[0]?.["$"]?.b || 0;

  let phType = placeholder?.type || null;
  let phIdx = placeholder?.idx || null;

  return { srcRectL, srcRectR, srcRectT, srcRectB };
}

function isTextShape(shapeNode) {
  return !!shapeNode?.["p:txBody"];
}

// IMPROVED: Better font size extraction with proper fallback hierarchy
function getFontSize(shapeNode, masterXML) {

  // Try to get from first run (most specific)
  const firstRun = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:r"]?.[0];
  if (firstRun?.["a:rPr"]?.[0]?.["$"]?.sz) {
    const sizeInHundredths = parseInt(firstRun["a:rPr"][0]["$"].sz);
    const pointSize = sizeInHundredths / 100;
    return pointSize;
  }

  // Try paragraph default properties
  const pPr = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:pPr"]?.[0];
  if (pPr?.["a:defRPr"]?.[0]?.["$"]?.sz) {
    const sizeInHundredths = parseInt(pPr["a:defRPr"][0]["$"].sz);
    const pointSize = sizeInHundredths / 100;
    return pointSize;
  }

  // Try end paragraph properties
  const endParaRPr = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:endParaRPr"]?.[0];
  if (endParaRPr?.["$"]?.sz) {
    const sizeInHundredths = parseInt(endParaRPr["$"].sz);
    const pointSize = sizeInHundredths / 100;
    return pointSize;
  }

  // Try list style if it's a list
  const lstStyle = shapeNode?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
  if (lstStyle?.["a:defPPr"]?.[0]?.["a:defRPr"]?.[0]?.["$"]?.sz) {
    const sizeInHundredths = parseInt(lstStyle["a:defPPr"][0]["a:defRPr"][0]["$"].sz);
    const pointSize = sizeInHundredths / 100;
    return pointSize;
  }
  
  if (lstStyle?.["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["$"]?.sz) {
    // fallback fontSize for slideLayout.xml
    const layoutSz = parseInt(lstStyle?.["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["$"]?.sz);
    const pointSize = layoutSz / 100;
    return pointSize;
  }

  // ✅ NEW: Check if it's a text box (not a placeholder)
  if (masterXML) {
    // console.log("popsaodsapdsadsad============");
    const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
    const hasPlaceholder = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"];

    if (isTextBox || !hasPlaceholder) {
      // Text boxes use otherStyle, not bodyStyle
      const txStyles = masterXML?.["p:sldMaster"]?.["p:txStyles"]?.[0];
      const otherStyle = txStyles?.["p:otherStyle"]?.[0];
      const level = getTextLevel(shapeNode);
      const levelPropName = level === 0 ? "a:lvl1pPr" : `a:lvl${level + 1}pPr`;
      
      const levelProp = otherStyle?.[levelPropName]?.[0];
      if (levelProp?.["a:defRPr"]?.[0]?.["$"]?.sz) {
        const sizeInHundredths = parseInt(levelProp["a:defRPr"][0]["$"].sz);
        console.log(`✅ Text box font size from otherStyle: ${sizeInHundredths / 100}pt`);
        return sizeInHundredths / 100;
      }
    }
    
    // For placeholders, use existing logic
    const placeholderType = getPlaceholderType(shapeNode);
    const level = getTextLevel(shapeNode);
    const masterFontSize = getFontSizeFromMaster(masterXML, placeholderType, level);

    if (masterFontSize) {
      return masterFontSize;
    }
  }
  
  // Fallback to default
  return 16;
}

function getPlaceholderType(shapeNode) {
  const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
  return phType || "body";
}

// Helper: Get text level from paragraph properties
function getTextLevel(shapeNode) {
  const pPr = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:pPr"]?.[0];
  const lvl = pPr?.["$"]?.lvl;
  return lvl ? parseInt(lvl) : 0; // Default to level 0
}

// Helper: Extract font size from master XML based on placeholder type and level
function getFontSizeFromMaster(masterXML, placeholderType, level = 0) {
  try {
    const txStyles = masterXML?.["p:sldMaster"]?.["p:txStyles"]?.[0];

    if (!txStyles) return null;

    let styleSection;

    // Determine which style section to use based on placeholder type
    switch (placeholderType) {
      case "title":
      case "ctrTitle":
        styleSection = txStyles["p:titleStyle"]?.[0];
        break;

      case "body":
      case "subTitle":
      case "obj":
        styleSection = txStyles["p:bodyStyle"]?.[0];
        break;

      case "dt":
      case "ftr":
      case "sldNum":
      default:
        styleSection = txStyles["p:otherStyle"]?.[0];
        break;
    }

    if (!styleSection) return null;

    // Map level to property name (lvl0 = lvl1pPr, lvl1 = lvl2pPr, etc.)
    const levelPropName = level === 0 ? "a:lvl1pPr" : `a:lvl${level + 1}pPr`;

    // Try to get size from the specific level
    const levelProp = styleSection[levelPropName]?.[0];

    if (levelProp?.["a:defRPr"]?.[0]?.["$"]?.sz) {
      const sizeInHundredths = parseInt(levelProp["a:defRPr"][0]["$"].sz);
      return sizeInHundredths / 100;
    }

    // Fallback to level 1 if specific level not found
    if (level > 0) {
      const lvl1 = styleSection["a:lvl1pPr"]?.[0];
      if (lvl1?.["a:defRPr"]?.[0]?.["$"]?.sz) {
        const sizeInHundredths = parseInt(lvl1["a:defRPr"][0]["$"].sz);
        return sizeInHundredths / 100;
      }
    }

    return null;
  } catch (error) {
    console.warn("Error extracting font size from master:", error);
    return null;
  }
}


// NEW: Get font size for specific run (useful for mixed formatting)
function getRunFontSize(runNode) {
  if (runNode?.["a:rPr"]?.[0]?.["$"]?.sz) {
    const sizeInHundredths = parseInt(runNode["a:rPr"][0]["$"].sz);
    return sizeInHundredths / 100;
  }
  return null; // Return null if no specific size found
}

// NEW: Get all font sizes in a shape (for debugging)
function getAllFontSizesInShape(shapeNode) {
  const fontSizes = [];

  const paragraphs = shapeNode?.["p:txBody"]?.[0]?.["a:p"] || [];

  paragraphs.forEach((paragraph, pIndex) => {
    const runs = paragraph?.["a:r"] || [];

    runs.forEach((run, rIndex) => {
      const runFontSize = getRunFontSize(run);
      if (runFontSize) {
        fontSizes.push({
          paragraph: pIndex,
          run: rIndex,
          fontSize: runFontSize,
          text: run?.["a:t"]?.[0] || ""
        });
      }
    });

    // Also check paragraph-level properties
    const pPr = paragraph?.["a:pPr"]?.[0];
    if (pPr?.["a:defRPr"]?.[0]?.["$"]?.sz) {
      const sizeInHundredths = parseInt(pPr["a:defRPr"][0]["$"].sz);
      fontSizes.push({
        paragraph: pIndex,
        run: "paragraph-default",
        fontSize: sizeInHundredths / 100,
        text: "paragraph default"
      });
    }

    const endParaRPr = paragraph?.["a:endParaRPr"]?.[0];
    if (endParaRPr?.["$"]?.sz) {
      const sizeInHundredths = parseInt(endParaRPr["$"].sz);
      fontSizes.push({
        paragraph: pIndex,
        run: "end-paragraph",
        fontSize: sizeInHundredths / 100,
        text: "end paragraph"
      });
    }
  });

  return fontSizes;
}

function getFontColor(shapeNode, themeXML, clrMap, masterXML) {
  const run = shapeNode?.["p:txBody"]?.[0]?.["a:p"]?.[0]?.["a:r"]?.[0]?.["a:rPr"]?.[0];

  // Check for direct run properties first
  if (run && run["a:solidFill"]) {
    // Handle RGB color
    const colorCode = run?.["a:solidFill"]?.[0]?.["a:srgbClr"]?.[0]?.["$"]?.val;
    if (colorCode) return `#${colorCode}`;

    // Handle theme color (schemeClr)
    const schemeClr = run?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;
    if (schemeClr) {
      const lumMod = run?.["a:solidFill"]?.[0]?.["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
      let fillColor = resolveFontColor(schemeClr, themeXML, clrMap);
      if (lumMod) {
        fillColor = colorHelper.applyLumMod(fillColor, lumMod);
      }
      return fillColor;
    }

    // Handle preset colors (prstClr)
    const presetClr = run?.["a:solidFill"]?.[0]?.["a:prstClr"]?.[0]?.["$"]?.val;
    if (presetClr) {
      const prstLumMod = run?.["a:solidFill"]?.[0]?.["a:prstClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;
      const prstColor = colorHelper.resolvePresetColor(presetClr);
      if (prstLumMod) {
        return colorHelper.applyLumMod(prstColor, prstLumMod);
      } else {
        return prstColor;
      }
    }
  }

  // Check for color in text runs (a:rPr)
  const txBody = shapeNode?.["p:txBody"]?.[0];
  const paragraphs = txBody?.["a:p"] || [];

  for (const paragraph of paragraphs) {
    const runs = paragraph?.["a:r"] || [];

    for (const run of runs) {
      const rPr = run?.["a:rPr"]?.[0];

      if (rPr) {
        const solidFill = rPr["a:solidFill"]?.[0];

        if (solidFill) {
          // 1. Check for direct RGB color (a:srgbClr)
          const srgbClr = solidFill["a:srgbClr"]?.[0];
          if (srgbClr) {
            const rgbVal = srgbClr["$"]?.val;
            if (rgbVal) {
              let fontColor = rgbVal.startsWith('#') ? rgbVal : `#${rgbVal}`;

              // Check for lumMod
              const lumMod = srgbClr["a:lumMod"]?.[0]?.["$"]?.val;
              if (lumMod) {
                fontColor = colorHelper.applyLumMod(fontColor, lumMod);
              }

              return fontColor;
            }
          }

          // 2. Check for scheme color (a:schemeClr)
          const schemeClr = solidFill["a:schemeClr"]?.[0];
          if (schemeClr) {
            const schemeVal = schemeClr["$"]?.val;
            if (schemeVal) {
              let fontColor = resolveFontColor(schemeVal, themeXML, clrMap);

              // Check for lumMod
              const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
              if (lumMod) {
                fontColor = colorHelper.applyLumMod(fontColor, lumMod);
              }

              return fontColor;
            }
          }

          // 3. Check for preset color (a:prstClr)
          const prstClr = solidFill["a:prstClr"]?.[0];
          if (prstClr) {
            const prstVal = prstClr["$"]?.val;
            if (prstVal) {

              // Map preset colors to hex values
              const presetColorMap = {
                'white': '#FFFFFF',
                'black': '#000000',
                'red': '#FF0000',
                'green': '#00FF00',
                'blue': '#0000FF',
                'yellow': '#FFFF00',
                'cyan': '#00FFFF',
                'magenta': '#FF00FF',
                'gray': '#808080',
                'darkGray': '#404040',
                'lightGray': '#C0C0C0',
                // Add more preset colors as needed
              };

              let fontColor = presetColorMap[prstVal] || '#000000'; // Default to black if not found

              // Check for lumMod
              const lumMod = prstClr["a:lumMod"]?.[0]?.["$"]?.val;
              if (lumMod) {
                fontColor = colorHelper.applyLumMod(fontColor, lumMod);
              }

              return fontColor;
            }
          }

          // 4. Check for system color with lastClr (a:sysClr)
          const sysClr = solidFill["a:sysClr"]?.[0];
          if (sysClr) {
            const lastClr = sysClr["$"]?.lastClr;

            if (lastClr) {
              // lastClr is already in hex format (e.g., "000000")
              let fontColor = lastClr.startsWith('#') ? lastClr : `#${lastClr}`;

              // Check for lumMod
              const lumMod = sysClr["a:lumMod"]?.[0]?.["$"]?.val;
              if (lumMod) {
                fontColor = colorHelper.applyLumMod(fontColor, lumMod);
              }

              return fontColor;
            }
          }
        }
      }
    }
  }

  // If no direct color is found in run, check for font reference in style
  const fontRef = shapeNode?.["p:style"]?.[0]?.["a:fontRef"]?.[0]; // Access first item of array

  if (fontRef) {
    const schemeClr = fontRef["a:schemeClr"]?.[0]?.["$"]?.val;
    if (schemeClr) {
      const lumMod = fontRef["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;

      let fontColor = resolveFontColor(schemeClr, themeXML, clrMap);

      if (lumMod) {
        fontColor = colorHelper.applyLumMod(fontColor, lumMod);
      }
      return fontColor;
    }
  }

  // ✅ NEW: Check master slide default text color (title/body/other)
  const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type || "body";

  let masterColorKey = null;

  // masterXML?.["p:sldMaster"]?.["p:txStyles"]?.[0]

  if (masterXML?.["p:sldMaster"]?.["p:txStyles"]) {
    const txStyles = masterXML?.["p:sldMaster"]?.["p:txStyles"]?.[0];

    if (phType === "title" && txStyles["p:titleStyle"]) {
      masterColorKey = txStyles["p:titleStyle"][0]["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;

    } else if (phType === "body" && txStyles["p:bodyStyle"]) {
      masterColorKey = txStyles["p:bodyStyle"][0]["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;

    } else if (txStyles["p:otherStyle"]) {
      masterColorKey = txStyles["p:otherStyle"][0]["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;

    }
  }

  if (masterColorKey) {
    const masterColor = resolveFontColor(masterColorKey, themeXML, clrMap);
    if (masterColor) return masterColor;
  }

  // Fallback to default if no color found
  return "#000000"; // Default fallback
}

function resolveFontColor(colorKey, themeXML, clrMap) {
  if (!clrMap || typeof clrMap !== "object") {
    console.warn("clrMap is undefined or not an object.");
    return "#000000"; // Default to black
  }

  const mappedKey = clrMap[colorKey] || colorKey;
  let colorNode = '';
  // Ensure colorKey exists in the mapping
  if (!mappedKey) {
    // console.warn(`Color key '${colorKey}' not found in clrMap. Falling back to black.`);
    return "#000000"; // Default to black
  }

  colorNode = getFontColorFromThemeColor(themeXML, mappedKey);

  if (!colorNode) {
    console.warn(`Color key '${colorKey}' not found in theme. Defaulting to black.`);
    return "#000000"; // Default to black
  }

  return resolveColorFromNode(colorNode, themeXML);
}

function getFontColorFromThemeColor(themeXML, schemeKey) {
  const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) return null;

  const colorNode = clrScheme[`a:${schemeKey}`]?.[0];
  if (!colorNode) {
    console.warn(`Theme color '${schemeKey}' not found.`);
    return null;
  }

  return colorNode;
}

function resolveColorFromNode(colorNode, themeXML) {
  if (!colorNode) return "#000000"; // Default fallback

  // Handle direct RGB color
  if (colorNode?.["a:srgbClr"]) {
    return `#${colorNode["a:srgbClr"][0]["$"]?.val}`;
  }

  // Handle system color
  if (colorNode?.["a:sysClr"]) {
    return `#${colorNode["a:sysClr"][0]["$"]?.lastClr || "FFFFFF"}`;
  }

  // Handle theme color reference
  if (colorNode?.["a:schemeClr"]) {
    const schemeColorKey = colorNode["a:schemeClr"][0]["$"]?.val;
    if (!schemeColorKey) {
      console.warn("No scheme color found, defaulting to black.");
      return "#000000"; // Default to black
    }

    const colorMap = {
      bg1: "lt1", // Light 1
      bg2: "lt2", // Light 2
      tx1: "dk1", // Dark 1
      tx2: "dk2", // Dark 2
    };

    const mappedColorKey = colorMap[schemeColorKey] || schemeColorKey;

    // Resolve base color from the theme
    const baseColor = resolveThemeColor(mappedColorKey, themeXML);

    if (baseColor) {
      return baseColor;
    } else {
      console.warn(`Could not resolve base color for '${schemeColorKey}'. Defaulting to black.`);
      return "#000000"; // Default to black
    }
  }

  return "#000000"; // Default fallback 
}

function resolveThemeColor(colorKey, themeXML) {
  const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
  if (!clrScheme) {
    return "#000000"; // Default color
  }

  // Check if the requested color exists
  const colorNode = clrScheme[`a:${colorKey}`]?.[0];
  if (!colorNode) {
    console.warn(`Color key '${colorKey}' not found in the theme. Using default black.`);
    return "#000000"; // Default color
  }

  return resolveColorFromNode(colorNode, themeXML);
}

function getTextAlignment(shapeNode) {
  const bodyPr = shapeNode?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
  const anchor = bodyPr?.["$"]?.anchor;

  const alignmentMap = {
    ctr: "center",
    r: "right",
    t: "top",
    b: "flex-end",
  };

  return alignmentMap[anchor] || "left";
}

function getRotation(shapeNode) {
  const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"];
  return xfrm?.[0]?.["$"]?.rot ? parseInt(xfrm[0]["$"]?.rot, 10) / 60000 : 0;
}

module.exports = {
  getPositionFromShape,
  getPositionFromPlaceholderLayout,
  getPositionFromShapeOrPlaceholder,
  isTextShape,
  getFontSize,
  getRunFontSize,
  getAllFontSizesInShape,
  getFontColor,
  getTextAlignment,
  getRotation,
};