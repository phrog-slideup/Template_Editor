// MasterBackgroundHandler.js - FIXED with proper transparency handling

/**
 * Enhanced Master Background Handler with proper transparency to hex conversion
 */

async function getMasterBackground(masterXML, themeXML, relationshipsXML = null, layoutXML = null, pptxInstance = null) {
  try {
    console.log("-================================Master========Call======================================");

    if (!masterXML || !themeXML) {
      console.warn('Missing required parameters: masterXML or themeXML');
      return { type: 'solid', color: '#FFFFFF', css: '#FFFFFF' };
    }

    // Navigate to the master slide background
    const masterBgNode = masterXML?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:bg"]?.[0];

    if (!masterBgNode) {
      console.warn('No background node found in master slide');
      return await getDefaultMasterBackground(themeXML);
    }

    let result = null;

    // Check for background properties
    const bgPr = masterBgNode?.["p:bgPr"]?.[0];
    if (bgPr) {
      result = await processBgPr(bgPr, themeXML, relationshipsXML, layoutXML, pptxInstance);

      console.log("result ==========-----111------------- >>>>>>>>>>>> ", result);

    }

    // Check for background reference
    if (!result) {
      const bgRef = masterBgNode?.["p:bgRef"]?.[0];
      if (bgRef) {
        result = await processBgRef(bgRef, themeXML, layoutXML);
        console.log("result ==========-------222----------- >>>>>>>>>>>> ", result);

      }
    }

    // If no specific background found, get default from theme
    if (!result) {
      result = await getDefaultMasterBackground(themeXML);

      console.log("result ==========-------3333----------- >>>>>>>>>>>> ", result);

    }

    // FIXED: Convert RGBA to hex if needed
    if (result && result.color) {
      result.color = ensureHexColor(result.color);
      result.css = result.color;
    }

    return result;

  } catch (error) {
    console.error('Error processing master background:', error);
    return { type: 'solid', color: '#FFFFFF', css: '#FFFFFF' };
  }
}

/**
 * FIXED: Apply alpha (transparency) by blending with white background
 */
function applyAlpha(color, alphaValue) {
  try {
    const rgb = hexToRgb(color);
    if (!rgb) return color;

    const { r, g, b } = rgb;

    // Blend with white background based on alpha value
    // Formula: newColor = sourceColor * alpha + backgroundColor * (1 - alpha)
    const newR = Math.round(r * alphaValue + 255 * (1 - alphaValue));
    const newG = Math.round(g * alphaValue + 255 * (1 - alphaValue));
    const newB = Math.round(b * alphaValue + 255 * (1 - alphaValue));

    return rgbToHex(newR, newG, newB);

  } catch (error) {
    console.error('Error applying alpha:', error);
    return color;
  }
}

/**
 * NEW: Ensure color is in hex format, convert RGBA if needed
 */
function ensureHexColor(color) {
  if (!color) return '#FFFFFF';

  // If already hex, return as is
  if (color.startsWith('#')) return color;

  // Convert RGBA to hex
  if (color.startsWith('rgba') || color.startsWith('rgb')) {
    return rgbaToHex(color);
  }

  return color;
}

/**
 * NEW: Convert RGBA/RGB to hex with transparency handling
 */
function rgbaToHex(rgba) {
  const match = rgba.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
  if (!match) return rgba;

  let r = parseInt(match[1]);
  let g = parseInt(match[2]);
  let b = parseInt(match[3]);
  const alpha = match[4] ? parseFloat(match[4]) : 1;

  // If there's transparency, blend with white background
  if (alpha < 1) {
    r = Math.round(r * alpha + 255 * (1 - alpha));
    g = Math.round(g * alpha + 255 * (1 - alpha));
    b = Math.round(b * alpha + 255 * (1 - alpha));
  }

  return rgbToHex(r, g, b);
}

/**
 * Process solid fill backgrounds - UPDATED
 */
async function processSolidFill(solidFill, themeXML, layoutXML) {
  try {
    // Handle scheme color
    const schemeClr = solidFill?.["a:schemeClr"]?.[0];
    if (schemeClr) {
      const colorValue = schemeClr?.["$"]?.val;
      if (colorValue) {
        let baseColor = await resolveColor(colorValue, themeXML, layoutXML);
        if (baseColor) {
          baseColor = await applyColorModifiers(baseColor, schemeClr);
          return { type: 'solid', color: ensureHexColor(baseColor), css: ensureHexColor(baseColor) };
        }
      }
    }

    // Handle sRGB color
    const srgbClr = solidFill?.["a:srgbClr"]?.[0];
    if (srgbClr) {
      let color = `#${srgbClr?.["$"]?.val}`;
      color = await applyColorModifiers(color, srgbClr);
      return { type: 'solid', color: ensureHexColor(color), css: ensureHexColor(color) };
    }

    // Handle system color
    const sysClr = solidFill?.["a:sysClr"]?.[0];
    if (sysClr) {
      let color = `#${sysClr?.["$"]?.lastClr || 'FFFFFF'}`;
      color = await applyColorModifiers(color, sysClr);
      return { type: 'solid', color: ensureHexColor(color), css: ensureHexColor(color) };
    }

    return null;
  } catch (error) {
    console.error('Error processing solid fill:', error);
    return null;
  }
}

/**
 * Process background reference (bgRef) - UPDATED
 */
async function processBgRef(bgRef, themeXML, layoutXML) {
  try {
    const idx = bgRef?.["$"]?.idx;
    const schemeClr = bgRef?.["a:schemeClr"]?.[0];

    if (schemeClr) {
      const colorValue = schemeClr?.["$"]?.val;
      if (colorValue) {
        let baseColor = await resolveColor(colorValue, themeXML, layoutXML);
        if (baseColor) {
          baseColor = await applyColorModifiers(baseColor, schemeClr);
          return { type: 'solid', color: ensureHexColor(baseColor), css: ensureHexColor(baseColor) };
        }
      }
    }

    return null;
  } catch (error) {
    console.error('Error processing background reference:', error);
    return null;
  }
}

/**
 * Process gradient fill backgrounds - UPDATED
 */
async function processGradientFill(gradFill, themeXML, layoutXML) {
  try {
    const gsLst = gradFill?.["a:gsLst"]?.[0]?.["a:gs"];
    if (!gsLst || !Array.isArray(gsLst)) {
      return null;
    }

    const gradientStops = [];

    for (const gs of gsLst) {
      const pos = gs?.["$"]?.pos || 0;
      const position = Math.round(pos / 1000); // Convert from 1000ths to percentage

      let color = null;

      // Process scheme color
      const schemeClr = gs?.["a:schemeClr"]?.[0];
      if (schemeClr) {
        const colorValue = schemeClr?.["$"]?.val;
        color = await resolveColor(colorValue, themeXML, layoutXML);
        color = await applyColorModifiers(color, schemeClr);
      }

      // Process sRGB color
      const srgbClr = gs?.["a:srgbClr"]?.[0];
      if (srgbClr) {
        color = `#${srgbClr?.["$"]?.val}`;
        color = await applyColorModifiers(color, srgbClr);
      }

      if (color) {
        gradientStops.push({ color: ensureHexColor(color), position });
      }
    }

    if (gradientStops.length === 0) {
      return null;
    }

    // Create CSS gradient
    const gradientCSS = createGradientCSS(gradientStops, gradFill);

    return {
      type: 'gradient',
      stops: gradientStops,
      css: gradientCSS
    };
  } catch (error) {
    console.error('Error processing gradient fill:', error);
    return null;
  }
}

/**
 * Process background properties (bgPr)
 */
async function processBgPr(bgPr, themeXML, relationshipsXML, layoutXML, pptxInstance) {
  try {
    // Handle solid fill
    const solidFill = bgPr?.["a:solidFill"]?.[0];
    if (solidFill) {
      return await processSolidFill(solidFill, themeXML, layoutXML);
    }

    // Handle gradient fill
    const gradFill = bgPr?.["a:gradFill"]?.[0];
    if (gradFill) {
      return await processGradientFill(gradFill, themeXML, layoutXML);
    }

    // Handle pattern fill
    const pattFill = bgPr?.["a:pattFill"]?.[0];
    if (pattFill) {
      return await processPatternFill(pattFill, themeXML, layoutXML);
    }

    // Handle blip fill (images)
    const blipFill = bgPr?.["a:blipFill"]?.[0];
    if (blipFill && relationshipsXML && pptxInstance) {
      return await processBlipFill(blipFill, relationshipsXML, pptxInstance);
    }

    // Handle no fill
    const noFill = bgPr?.["a:noFill"];
    if (noFill) {
      return { type: 'none', color: 'transparent', css: 'transparent' };
    }

    return null;
  } catch (error) {
    console.error('Error processing background properties:', error);
    return null;
  }
}

/**
 * Process pattern fill backgrounds - UPDATED
 */
async function processPatternFill(pattFill, themeXML, layoutXML) {
  try {
    const prst = pattFill?.["$"]?.prst;

    // Get foreground color
    let fgColor = '#ffffff';
    const fgClr = pattFill?.["a:fgClr"]?.[0];
    if (fgClr) {
      const schemeClr = fgClr?.["a:schemeClr"]?.[0];
      if (schemeClr) {
        const colorValue = schemeClr?.["$"]?.val;
        fgColor = await resolveColor(colorValue, themeXML, layoutXML);
        fgColor = await applyColorModifiers(fgColor, schemeClr);
      }

      const srgbClr = fgClr?.["a:srgbClr"]?.[0];
      if (srgbClr) {
        fgColor = `#${srgbClr?.["$"]?.val}`;
        fgColor = await applyColorModifiers(fgColor, srgbClr);
      }
    }

    // Get background color
    let bgColor = '#FFFFFF';
    const bgClr = pattFill?.["a:bgClr"]?.[0];
    if (bgClr) {
      const schemeClr = bgClr?.["a:schemeClr"]?.[0];
      if (schemeClr) {
        const colorValue = schemeClr?.["$"]?.val;
        bgColor = await resolveColor(colorValue, themeXML, layoutXML);
        bgColor = await applyColorModifiers(bgColor, schemeClr);
      }

      const srgbClr = bgClr?.["a:srgbClr"]?.[0];
      if (srgbClr) {
        bgColor = `#${srgbClr?.["$"]?.val}`;
        bgColor = await applyColorModifiers(bgColor, srgbClr);
      }
    }

    return {
      type: 'pattern',
      pattern: prst,
      fgColor: ensureHexColor(fgColor),
      bgColor: ensureHexColor(bgColor),
      css: ensureHexColor(bgColor) // Fallback to background color
    };
  } catch (error) {
    console.error('Error processing pattern fill:', error);
    return null;
  }
}

/**
 * Process blip fill (image) backgrounds
 */
async function processBlipFill(blipFill, relationshipsXML, pptxInstance) {
  try {
    const blip = blipFill?.["a:blip"]?.[0];
    if (!blip) {
      return null;
    }

    // Get the relationship ID for the image
    const embedId = blip?.["$"]?.["r:embed"];
    const linkId = blip?.["$"]?.["r:link"];

    let imageId = embedId || linkId;
    if (!imageId) {
      console.warn('No image relationship ID found in blip fill');
      return null;
    }

    // Find the relationship in the relationships XML
    const relationships = relationshipsXML?.["Relationships"]?.["Relationship"];
    if (!relationships) {
      console.warn('No relationships found for image processing');
      return null;
    }

    const imageRelationship = relationships.find(rel => rel?.["$"]?.Id === imageId);
    if (!imageRelationship) {
      console.warn(`Image relationship not found for ID: ${imageId}`);
      return null;
    }

    const imageTarget = imageRelationship?.["$"]?.Target;
    if (!imageTarget) {
      console.warn('No target found for image relationship');
      return null;
    }

    // Construct the full image path
    let imagePath = imageTarget;
    if (!imagePath.startsWith('ppt/')) {
      imagePath = `ppt/media/${imagePath.replace('../media/', '')}`;
    }

    // Get image data using the pptx instance
    let imageDataUrl = null;
    if (pptxInstance && pptxInstance.files && pptxInstance.files[imagePath]) {
      try {
        const imageFile = pptxInstance.files[imagePath];
        const imageData = await imageFile.async("base64");

        // Determine MIME type based on file extension
        const fileExtension = imagePath.split('.').pop().toLowerCase();
        const mimeType = getMimeType(fileExtension);

        imageDataUrl = `data:${mimeType};base64,${imageData}`;
      } catch (imageError) {
        console.error('Error processing image data:', imageError);
        return null;
      }
    }

    // Process stretch or tile properties
    const stretch = blipFill?.["a:stretch"]?.[0];
    const tile = blipFill?.["a:tile"]?.[0];

    let cssProperties = {
      'background-image': `url(${imageDataUrl})`,
      'background-repeat': tile ? 'repeat' : 'no-repeat',
      'background-size': stretch ? 'cover' : 'auto',
      'background-position': 'center center'
    };

    const cssString = Object.entries(cssProperties)
      .map(([key, value]) => `${key}: ${value}`)
      .join('; ');

    return {
      type: 'image',
      imagePath: imagePath,
      imageDataUrl: imageDataUrl,
      css: cssString,
      properties: cssProperties
    };

  } catch (error) {
    console.error('Error processing blip fill:', error);
    return null;
  }
}

// [Include all the remaining utility functions - resolveColor, applyColorModifiers, etc.]
// These remain the same as in your original code

/**
 * Get MIME type based on file extension
 */
function getMimeType(extension) {
  const mimeTypes = {
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'webp': 'image/webp',
    'svg': 'image/svg+xml'
  };

  return mimeTypes[extension] || 'image/jpeg';
}

/**
 * Create CSS gradient from gradient stops
 */
function createGradientCSS(stops, gradFill) {
  // Sort stops by position
  stops.sort((a, b) => a.position - b.position);

  // Create gradient string
  const stopStrings = stops.map(stop => `${stop.color} ${stop.position}%`);

  // Check gradient direction/type
  const lin = gradFill?.["a:lin"]?.[0];
  if (lin) {
    const angle = lin?.["$"]?.ang || 0;
    const degrees = Math.round(angle / 60000); // Convert from 60000ths of a degree
    return `linear-gradient(${degrees}deg, ${stopStrings.join(', ')})`;
  }

  const radial = gradFill?.["a:radial"]?.[0];
  if (radial) {
    return `radial-gradient(circle, ${stopStrings.join(', ')})`;
  }

  // Default to linear gradient
  return `linear-gradient(45deg, ${stopStrings.join(', ')})`;
}

/**
 * Resolve color from theme
 */
async function resolveColor(colorName, themeXML, layoutXML = null) {
  try {
    // Check for layout color mapping override first
    if (layoutXML) {
      const overrideClrMapping = layoutXML?.["p:sldLayout"]?.["p:clrMapOvr"]?.[0]?.["a:overrideClrMapping"]?.[0]?.["$"];
      if (overrideClrMapping && overrideClrMapping[colorName]) {
        colorName = overrideClrMapping[colorName];
      }
    }

    const themeElements = themeXML?.["a:theme"]?.["a:themeElements"]?.[0];
    const clrScheme = themeElements?.["a:clrScheme"]?.[0];

    if (!clrScheme) {
      console.warn('No color scheme found in theme');
      return null;
    }

    // Map color names to theme elements
    const colorMap = {
      'bg1': clrScheme?.["a:lt1"]?.[0],
      'bg2': clrScheme?.["a:lt2"]?.[0],
      'tx1': clrScheme?.["a:dk1"]?.[0],
      'tx2': clrScheme?.["a:dk2"]?.[0],
      'lt1': clrScheme?.["a:lt1"]?.[0],
      'lt2': clrScheme?.["a:lt2"]?.[0],
      'dk1': clrScheme?.["a:dk1"]?.[0],
      'dk2': clrScheme?.["a:dk2"]?.[0],
      'accent1': clrScheme?.["a:accent1"]?.[0],
      'accent2': clrScheme?.["a:accent2"]?.[0],
      'accent3': clrScheme?.["a:accent3"]?.[0],
      'accent4': clrScheme?.["a:accent4"]?.[0],
      'accent5': clrScheme?.["a:accent5"]?.[0],
      'accent6': clrScheme?.["a:accent6"]?.[0],
      'hlink': clrScheme?.["a:hlink"]?.[0],
      'folHlink': clrScheme?.["a:folHlink"]?.[0]
    };

    const colorElement = colorMap[colorName];
    if (!colorElement) {
      console.warn(`Unknown scheme color: ${colorName}`);
      return null;
    }

    // Extract color value
    const srgbClr = colorElement?.["a:srgbClr"]?.[0];
    if (srgbClr) {
      return `#${srgbClr?.["$"]?.val}`;
    }

    const sysClr = colorElement?.["a:sysClr"]?.[0];
    if (sysClr) {
      return `#${sysClr?.["$"]?.lastClr || 'FFFFFF'}`;
    }

    return null;
  } catch (error) {
    console.error('Error resolving color:', error);
    return null;
  }
}

/**
 * Apply color modifiers like lumMod, lumOff, tint, shade, etc. - UPDATED
 */
async function applyColorModifiers(baseColor, colorNode) {
  try {
    if (!baseColor || !colorNode) {
      return baseColor;
    }

    let color = baseColor;

    // Handle luminance modifier (lumMod)
    const lumMod = colorNode?.["a:lumMod"]?.[0]?.["$"]?.val;
    if (lumMod) {
      const modValue = validateLuminanceValue(lumMod, 100000) / 100000;
      console.debug(`Applying lumMod: ${lumMod} -> ${modValue}`);
      color = adjustLuminance(color, modValue, 'multiply');
    }

    // Handle luminance offset (lumOff)
    const lumOff = colorNode?.["a:lumOff"]?.[0]?.["$"]?.val;
    if (lumOff) {
      const offValue = validateLuminanceValue(lumOff, 0) / 100000;
      console.debug(`Applying lumOff: ${lumOff} -> ${offValue}`);
      color = adjustLuminance(color, offValue, 'offset');
    }

    // Handle tint
    const tint = colorNode?.["a:tint"]?.[0]?.["$"]?.val;
    if (tint) {
      const tintValue = validateLuminanceValue(tint, 0) / 100000;
      color = applyTint(color, tintValue);
    }

    // Handle shade
    const shade = colorNode?.["a:shade"]?.[0]?.["$"]?.val;
    if (shade) {
      const shadeValue = validateLuminanceValue(shade, 0) / 100000;
      color = applyShade(color, shadeValue);
    }

    // Handle alpha (transparency) - FIXED to return hex
    const alpha = colorNode?.["a:alpha"]?.[0]?.["$"]?.val;
    if (alpha) {
      const alphaValue = validateLuminanceValue(alpha, 100000) / 100000;
      color = applyAlpha(color, alphaValue);
    }

    return color;
  } catch (error) {
    console.error('Error applying color modifiers:', error);
    return baseColor;
  }
}

/**
 * Validate luminance values
 */
function validateLuminanceValue(value, defaultValue = 100000) {
  if (!value) return defaultValue;

  const numValue = parseInt(value);

  // Luminance values should be between 0 and 200000 (0% to 200%)
  if (isNaN(numValue) || numValue < 0 || numValue > 200000) {
    console.warn(`Invalid luminance value: ${value}, using default: ${defaultValue}`);
    return defaultValue;
  }

  return numValue;
}

/**
 * Apply luminance modifier (for lumMod and lumOff)
 */
function adjustLuminance(color, value, operation) {
  try {
    const rgb = hexToRgb(color);
    if (!rgb) return color;

    let { r, g, b } = rgb;

    if (operation === 'multiply') {
      r = Math.round(Math.max(0, Math.min(255, r * value)));
      g = Math.round(Math.max(0, Math.min(255, g * value)));
      b = Math.round(Math.max(0, Math.min(255, b * value)));
    } else if (operation === 'offset') {
      const offset = Math.round(255 * value);
      r = Math.max(0, Math.min(255, r + offset));
      g = Math.max(0, Math.min(255, g + offset));
      b = Math.max(0, Math.min(255, b + offset));
    }

    return rgbToHex(r, g, b);
  } catch (error) {
    console.error('Error adjusting luminance:', error);
    return color;
  }
}

/**
 * Apply tint (mix with white)
 */
function applyTint(color, tintValue) {
  try {
    const rgb = hexToRgb(color);
    if (!rgb) return color;

    const { r, g, b } = rgb;

    const newR = Math.round(r + (255 - r) * tintValue);
    const newG = Math.round(g + (255 - g) * tintValue);
    const newB = Math.round(b + (255 - b) * tintValue);

    return rgbToHex(newR, newG, newB);
  } catch (error) {
    console.error('Error applying tint:', error);
    return color;
  }
}

/**
 * Apply shade (mix with black)
 */
function applyShade(color, shadeValue) {
  try {
    const rgb = hexToRgb(color);
    if (!rgb) return color;

    const { r, g, b } = rgb;

    const newR = Math.round(r * (1 - shadeValue));
    const newG = Math.round(g * (1 - shadeValue));
    const newB = Math.round(b * (1 - shadeValue));

    return rgbToHex(newR, newG, newB);
  } catch (error) {
    console.error('Error applying shade:', error);
    return color;
  }
}

/**
 * Get default master background from theme
 */
async function getDefaultMasterBackground(themeXML) {
  try {
    // Default to bg1 scheme color
    const defaultColor = await resolveColor('bg1', themeXML);
    return {
      type: 'solid',
      color: ensureHexColor(defaultColor || '#FFFFFF'),
      css: ensureHexColor(defaultColor || '#FFFFFF')
    };
  } catch (error) {
    console.error('Error getting default master background:', error);
    return { type: 'solid', color: '#FFFFFF', css: '#FFFFFF' };
  }
}

/**
 * Convert hex to RGB
 */
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    r: parseInt(result[1], 16),
    g: parseInt(result[2], 16),
    b: parseInt(result[3], 16)
  } : null;
}

/**
 * Convert RGB to hex
 */
function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

module.exports = {
  getMasterBackground
};