const fs = require("fs");
const path = require("path");
const { resolveColor } = require("../../pptx-To-Html-styling/pptBackgroundColors.js"); // Import theme functions
const config = require('../../config.js');

// Define the directory to save images using config
const imageSavePath = path.resolve(__dirname, "../../uploads");
if (!fs.existsSync(imageSavePath)) {
  fs.mkdirSync(imageSavePath, { recursive: true });
}

/**
 * Normalize path to prevent double ppt/ prefix
 */
function normalizeImagePath(target) {
  if (!target) return null;

  // Convert to string and handle backslashes
  let normalized = String(target).replace(/\\/g, '/');

  // Handle relative paths with ../
  if (normalized.startsWith('../')) {
    // Remove ../ prefix
    normalized = normalized.replace(/^\.\.\//, '');
  }

  // Remove leading slashes
  normalized = normalized.replace(/^\/+/, '');

  // Remove any double ppt/ prefix that might exist
  while (normalized.includes('ppt/ppt/')) {
    normalized = normalized.replace('ppt/ppt/', 'ppt/');
  }

  // Ensure it starts with ppt/ if it doesn't already
  if (!normalized.startsWith('ppt/')) {
    normalized = 'ppt/' + normalized;
  }

  return normalized;
}

async function getImageFromPicture(node, filePath, files, relationshipsXML, themeXML) {
  try {

    console.log(" @@@@@@22@@@@@@@ Extracting image from picture node...  @@@@@@@@@@@@@@");
    console.log("Theme accent colors:", themeXML?.["a:theme"]?.[0]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0]);

    const nodeName = node?.["p:nvPicPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
    console.log('\n=== PICTURE NODE DEBUG ===');
    console.log('Picture name:', nodeName);
    console.log('Full nvPicPr:', JSON.stringify(node?.["p:nvPicPr"], null, 2));
    const altText = node?.["p:nvPicPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.descr || '';
    console.log('Alt text (descr):====================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>', altText);

    const blipFillNode = node?.["p:blipFill"]?.[0] ||
      node?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0];
    const blip = blipFillNode?.["a:blip"]?.[0];

    const imageId = blip?.["$"]?.["r:embed"];
    if (!imageId) return null;

    // Find the relationship for this image
    const relationship = relationshipsXML?.["Relationships"]?.["Relationship"]?.find(
      (rel) => rel["$"].Id === imageId
    );

    if (!relationship) {
      console.warn(`No relationship found for image ID: ${imageId}`);
      return null;
    }

    // FIXED: Use normalizeImagePath to properly construct the path
    const targetPath = normalizeImagePath(relationship["$"].Target);

    const imageFile = files[targetPath];

    if (!imageFile) {
      console.warn("Image file not found in unzipped files:", targetPath);
      // Debug: Show available media files
      const mediaFiles = Object.keys(files).filter(k => k.includes('media'));
      console.warn("Available media files:", mediaFiles.slice(0, 10));
      return null;
    }

    // Detect if image is an ellipse (used for masking)
    const shapeType = node?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["$"]?.prst || "rect";
    const isEllipse = shapeType === "ellipse";

    // Extract border properties
    const lineNode = node?.["p:spPr"]?.[0]?.["a:ln"]?.[0];
    const borderWidth = lineNode?.["$"]?.w ? parseInt(lineNode["$"].w, 10) / 12700 : 0;

    let borderColor = "black"; // Default border color
    if (lineNode?.["a:solidFill"]?.[0]?.["a:srgbClr"]) {
      borderColor = `#${lineNode["a:solidFill"][0]["a:srgbClr"][0]["$"].val}`;
    } else if (lineNode?.["a:solidFill"]?.[0]?.["a:schemeClr"]) {
      const schemeClrKey = lineNode["a:solidFill"][0]["a:schemeClr"][0]["$"].val;
      borderColor = resolveBorderColor(schemeClrKey, themeXML); // Convert theme color
    }

    // ✅ ADD THIS: Extract border dash style
    let borderStyle = 'solid'; // Default to solid

    if (lineNode) {
      // Check for dash preset in PowerPoint XML
      const prstDash = lineNode["a:prstDash"]?.[0]?.["$"]?.val;

      if (prstDash) {
        // Map PowerPoint dash types to CSS border styles
        const dashStyleMap = {
          'solid': 'solid',
          'dot': 'dotted',
          'dash': 'dashed',
          'lgDash': 'dashed',
          'dashDot': 'dashed',
          'lgDashDot': 'dashed',
          'lgDashDotDot': 'dashed',
          'sysDot': 'dotted',
          'sysDash': 'dashed',
          'sysDashDot': 'dashed',
          'sysDashDotDot': 'dashed'
        };

        borderStyle = dashStyleMap[prstDash] || 'solid';
        console.log('✅ Border style extracted from PPT:', prstDash, '→', borderStyle);
      } else {
        console.log('No prstDash found, using solid border');
      }
    }


    // ✅ IMPROVED: Extract hyperlink from picture
    let hyperlink = null;

    const nvPicPr = node?.["p:nvPicPr"]?.[0];
    const cNvPr = nvPicPr?.["p:cNvPr"]?.[0];
    const hlinkClick = cNvPr?.["a:hlinkClick"]?.[0];

    console.log('=== HYPERLINK EXTRACTION DEBUG ===');
    console.log('cNvPr exists:', !!cNvPr);
    console.log('hlinkClick exists:', !!hlinkClick);

    if (hlinkClick) {
      // The hyperlink ID is stored in r:id attribute
      const hyperlinkId = hlinkClick?.["$"]?.["r:id"];
      console.log('Hyperlink ID from picture:', hyperlinkId);

      if (hyperlinkId && relationshipsXML) {
        // Find the relationship where Type ends with "hyperlink"
        const relationships = relationshipsXML?.["Relationships"]?.["Relationship"];

        if (relationships) {
          const hyperlinkRel = relationships.find(rel => {
            const relId = rel["$"].Id;
            const relType = rel["$"].Type;
            const isMatch = relId === hyperlinkId && relType.includes('hyperlink');

            console.log(`Checking relationship ${relId}: ${relType} - Match: ${isMatch}`);
            return isMatch;
          });

          if (hyperlinkRel) {
            hyperlink = hyperlinkRel["$"].Target;
            console.log('✅ Hyperlink extracted:', hyperlink);
          } else {
            console.warn('❌ No matching hyperlink relationship found for ID:', hyperlinkId);
          }
        }
      }
    } else {
      console.log('No hyperlink on this image');
    }
    // Extract Transparency (Opacity)
    let opacity = 1;
    const alphaNode = node?.["p:spPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:alpha"];

    if (alphaNode) {
      opacity = parseInt(alphaNode[0]["$"].val, 10) / 100000;
    }

    // Extract Shadow Properties
    const effectLst = node?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0];
    const outerShdw = effectLst?.["a:outerShdw"]?.[0];
    // ADD THESE DEBUG LINES:
    console.log("=== SHADOW DEBUG ===");
    console.log("effectLst exists:", !!effectLst);
    console.log("outerShdw full:", JSON.stringify(outerShdw, null, 2));

    // Extract Shadow AND Glow Properties
    let shadowData = null;

    if (effectLst) {
      // Extract GLOW effect
      const glow = effectLst?.["a:glow"]?.[0];
      let glowData = null;

      if (glow) {
        const glowRadius = glow?.["$"]?.rad ? parseInt(glow["$"].rad, 10) / 12700 : 0;
        let glowColor = "rgba(128, 0, 128, 0.6)"; // default purple

        if (glow["a:srgbClr"]) {
          const colorVal = glow["a:srgbClr"][0]["$"].val;
          const alpha = glow["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
          const glowAlpha = alpha ? parseInt(alpha, 10) / 100000 : 0.6;
          const r = parseInt(colorVal.substring(0, 2), 16);
          const g = parseInt(colorVal.substring(2, 4), 16);
          const b = parseInt(colorVal.substring(4, 6), 16);
          glowColor = `rgba(${r}, ${g}, ${b}, ${glowAlpha})`;
        } else if (glow["a:schemeClr"]) {
          const schemeClr = glow["a:schemeClr"][0]["$"].val;
          let hexColor = null;

          // Try theme first, then fallback
          if (themeXML) {
            try {
              hexColor = resolveBorderColor(schemeClr, themeXML);
            } catch (e) {
              console.warn("Could not resolve glow color from theme:", e);
            }
          }

          // Use fallback if theme failed
          if (!hexColor || hexColor === 'black' || hexColor === '#000000') {
            hexColor = getAccentColorFallback(schemeClr);
          }

          if (hexColor && hexColor.startsWith('#') && hexColor.length >= 7) {
            const r = parseInt(hexColor.substring(1, 3), 16);
            const g = parseInt(hexColor.substring(3, 5), 16);
            const b = parseInt(hexColor.substring(5, 7), 16);
            glowColor = `rgba(${r}, ${g}, ${b}, 0.6)`;
          }
        }

        glowData = {
          radius: glowRadius,
          color: glowColor
        };
      }

      // Extract OUTER SHADOW effect
      let outerShadowData = null;

      if (outerShdw) {
        const blurRad = outerShdw?.["$"]?.blurRad ? parseInt(outerShdw["$"].blurRad, 10) / 12700 : 0;
        const sx = outerShdw?.["$"]?.sx ? parseInt(outerShdw["$"].sx, 10) / 12700 : 0;
        const sy = outerShdw?.["$"]?.sy ? parseInt(outerShdw["$"].sy, 10) / 12700 : 0;

        let shadowColor = "rgba(0, 0, 0, 0.5)"; // default

        if (outerShdw["a:srgbClr"]) {
          const colorVal = outerShdw["a:srgbClr"][0]["$"].val;
          const alpha = outerShdw["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
          const shadowAlpha = alpha ? parseInt(alpha, 10) / 100000 : 1;
          const r = parseInt(colorVal.substring(0, 2), 16);
          const g = parseInt(colorVal.substring(2, 4), 16);
          const b = parseInt(colorVal.substring(4, 6), 16);
          shadowColor = `rgba(${r}, ${g}, ${b}, ${shadowAlpha})`;
        } else if (outerShdw["a:schemeClr"]) {
          const schemeClr = outerShdw["a:schemeClr"][0]["$"].val;
          let hexColor = null;

          // Try theme first, then fallback
          if (themeXML) {
            try {
              hexColor = resolveBorderColor(schemeClr, themeXML);
            } catch (e) {
              console.warn("Could not resolve shadow color from theme:", e);
            }
          }

          // Use fallback if theme failed
          if (!hexColor || hexColor === 'black' || hexColor === '#000000') {
            hexColor = getAccentColorFallback(schemeClr);
          }

          if (hexColor && hexColor.startsWith('#') && hexColor.length >= 7) {
            const r = parseInt(hexColor.substring(1, 3), 16);
            const g = parseInt(hexColor.substring(3, 5), 16);
            const b = parseInt(hexColor.substring(5, 7), 16);
            shadowColor = `rgba(${r}, ${g}, ${b}, 1)`;
          }
        }

        outerShadowData = {
          blur: blurRad,
          offsetX: sx,
          offsetY: sy,
          color: shadowColor
        };
        console.log("outerShadowData===============================>", outerShadowData);
      }

      // Combine glow and shadow
      if (glowData || outerShadowData) {
        shadowData = {
          glow: glowData,
          shadow: outerShadowData
        };
      }
    }
    // Extract Image Transformations (Flip, Rotation)
    const xfrm = node?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];

    let flipH = "";
    let flipV = "";
    let rotation = 0;

    if (xfrm?.["$"]?.flipH === "1") flipH = "scaleX(-1)";
    if (xfrm?.["$"]?.flipV === "1") flipV = "scaleY(-1)";
    if (xfrm?.["$"]?.rot) rotation = parseInt(xfrm["$"].rot, 10) / 60000;

    // NEW: Extract cropping information
    const cropData = extractImageCropping(blipFillNode);
    const croppingStyles = cropData ? generateCroppingStyles(cropData) : null;

    // Save Image to Disk
    const imageBuffer = await imageFile.async("nodebuffer");
    const imageKey = `image_${Date.now()}_${path.basename(targetPath)}`;
    const filePath = path.join(imageSavePath, imageKey);

    fs.writeFileSync(filePath, imageBuffer);

    return {
      src: `${config.uploadPath}/${imageKey}`,
      shape: isEllipse ? "ellipse" : "rect",
      border: { width: borderWidth, color: borderColor, style: borderStyle },
      opacity,
      transform: `${flipH} ${flipV} rotate(${rotation}deg)`,
      cropping: cropData,
      croppingStyles: croppingStyles,
      hasCropping: cropData !== null,
      shadow: shadowData,
      hasShadow: shadowData !== null,
      hyperlink: hyperlink,  // ✅ ADD THIS,
      altText: altText  // ✅ ADD THIS
    };
  } catch (error) {
    console.error("Error in getImageFromPicture:", error);
    return null;
  }
}

// Helper function to resolve border colors based on theme XML
// Helper function to resolve border colors based on theme XML or use fallback
function resolveBorderColor(colorKey, themeXML) {
  if (!colorKey) return "black";

  // Standard Office theme accent colors (fallback)
  const accentColors = {
    'accent1': '#4472C4',
    'accent2': '#ED7D31',
    'accent3': '#A5A5A5',
    'accent4': '#FFC000',
    'accent5': '#C55A9F', // Purple/Magenta
    'accent6': '#70AD47', // Green
    'lt1': '#FFFFFF',     // Light 1
    'lt2': '#E7E6E6',     // Light 2
    'dk1': '#000000',     // Dark 1
    'dk2': '#44546A',     // Dark 2
    'bg1': '#FFFFFF',     // Background 1
    'bg2': '#E7E6E6',     // Background 2
    'tx1': '#000000',     // Text 1
    'tx2': '#44546A'      // Text 2
  };

  // Color mapping
  const colorMap = {
    bg1: "lt1",
    bg2: "lt2",
    tx1: "dk1",
    tx2: "dk2",
  };

  const mappedColorKey = colorMap[colorKey] || colorKey;

  // Try to resolve from theme if available
  if (themeXML) {
    try {
      // Try using the imported resolveColor if it exists
      const { resolveColor } = require("../../pptx-To-Html-styling/pptBackgroundColors.js");
      if (resolveColor && typeof resolveColor === 'function') {
        const themeColor = resolveColor(mappedColorKey, themeXML);
        if (themeColor && themeColor !== 'black' && themeColor !== '#000000') {
          return themeColor;
        }
      }
    } catch (e) {
      // resolveColor not available, use fallback
      console.warn('resolveColor not available, using fallback colors');
    }
  }

  // Use fallback colors
  return accentColors[mappedColorKey] || accentColors[colorKey] || '#000000';
}
// Helper function to get accent colors when theme is unavailable
function getAccentColorFallback(accentName) {
  // Standard Office theme accent colors
  const accentColors = {
    'accent1': '#4472C4',
    'accent2': '#ED7D31',
    'accent3': '#A5A5A5',
    'accent4': '#FFC000',
    'accent5': '#C55A9F', // Purple/Magenta - matches your PowerPoint
    'accent6': '#70AD47'
  };

  return accentColors[accentName] || '#000000';
}

// Function that overrides the method in the pptxToHtml class
function overrideGetImageFromPicture(extractor) {
  extractor.getImageFromPicture = async function (node, slidePath, relationshipsXML) {
    try {
      // Call the standalone function with the necessary parameters
      return await getImageFromPicture(
        node,
        slidePath,
        this.files,
        relationshipsXML,
        this.themeXML
      );
    } catch (error) {
      console.error("Error in overridden getImageFromPicture:", error);
      return null;
    }
  };
}

/**
 * Extract cropping information from blipFill srcRect
 */
function extractImageCropping(blipFillNode) {
  if (!blipFillNode) return null;



  const srcRect = blipFillNode?.["a:srcRect"]?.[0]?.["$"];
  if (!srcRect) return null;

  // ✅ ADD THIS DEBUG
  console.log('===== RAW srcRect from XML =====');
  console.log('srcRect.l:', srcRect.l);
  console.log('srcRect.r:', srcRect.r);
  console.log('srcRect.t:', srcRect.t);
  console.log('srcRect.b:', srcRect.b);
  const cropping = {
    left: (srcRect.l / 100000) * 100,
    right: (srcRect.r / 100000) * 100,
    top: (srcRect.t / 100000) * 100,
    bottom: (srcRect.b / 100000) * 100,
    l: parseInt(srcRect.l) || 0,  // ✅ KEEP NEGATIVES
    r: parseInt(srcRect.r) || 0,
    t: parseInt(srcRect.t) || 0,
    b: parseInt(srcRect.b) || 0
  };

  console.log('Converted to percentages:', cropping);


  // Check if any cropping is applied
  if (cropping.left === 0 && cropping.top === 0 && cropping.right === 0 && cropping.bottom === 0) {
    return null;
  }

  return cropping;
}

/**
 * Generate CSS styles for cropping
 */
function generateCroppingStyles(cropData) {
  if (!cropData) return null;

  const { left, top, right, bottom } = cropData;
  const visibleWidth = Math.max(1, 100 - Math.max(0, left) - Math.max(0, right));
  const visibleHeight = Math.max(1, 100 - Math.max(0, top) - Math.max(0, bottom));

  let imageWidth, imageHeight, imageLeft, imageTop;

  if (left < 0 || right < 0) {
    const totalWidth = 100 + Math.abs(Math.min(0, left)) + Math.abs(Math.min(0, right));
    imageWidth = (totalWidth / visibleWidth) * 100;
    imageLeft = Math.min(0, left);
  } else {
    imageWidth = (100 / visibleWidth) * 100;
    imageLeft = -(Math.max(0, left) / visibleWidth) * 100;
  }

  if (top < 0 || bottom < 0) {
    const totalHeight = 100 + Math.abs(Math.min(0, top)) + Math.abs(Math.min(0, bottom));
    imageHeight = (totalHeight / visibleHeight) * 100;
    imageTop = Math.min(0, top);
  } else {
    imageHeight = (100 / visibleHeight) * 100;
    imageTop = -(Math.max(0, top) / visibleHeight) * 100;
  }

  return {
    containerStyles: `overflow: hidden; position: relative;`,
    imageStyles: `
            width: ${imageWidth.toFixed(2)}%;
            height: ${imageHeight.toFixed(2)}%;
            position: absolute;
            left: ${imageLeft.toFixed(2)}%;
            top: ${imageTop.toFixed(2)}%;
            object-fit: cover;
        `.replace(/\s+/g, ' ').trim()
  };
}

/**
 * Generate CSS box-shadow from shadow data
 */
function generateShadowCSS(shadowData) {
  if (!shadowData) return "";

  const { blur, distance, angle, color } = shadowData;

  // Convert angle and distance to x,y offsets
  const angleRad = (angle * Math.PI) / 180;
  const offsetX = Math.cos(angleRad) * distance;
  const offsetY = Math.sin(angleRad) * distance;

  return `${offsetX.toFixed(2)}px ${offsetY.toFixed(2)}px ${blur.toFixed(2)}px ${color}`;
}

// Export both functions
module.exports = {
  overrideGetImageFromPicture,
  getImageFromPicture,
  generateShadowCSS  // Add this
};