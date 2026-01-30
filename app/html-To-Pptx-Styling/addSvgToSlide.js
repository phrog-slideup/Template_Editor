// FIXED: addSvgToSlide.js - Properly handles stroke-width="0" to prevent unwanted borders
const PptxGenJS = require("pptxgenjs");

/**
 * Convert SVG path to PptxGenJS points format
 */
function convertSvgPathToPptxPoints(pathData, viewBoxWidth, viewBoxHeight, shapeWidth, shapeHeight) {
    if (!pathData || !viewBoxWidth || !viewBoxHeight) {
        console.log('   ❌ Invalid parameters for path conversion');
        return [];
    }

    const points = [];
    try {
        // Clean up path data
        const cleanPathData = pathData.replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
        const commands = cleanPathData.match(/[MLHVCSQTAZ][^MLHVCSQTAZ]*/gi);

        if (!commands || commands.length === 0) {
            console.log('   ❌ No valid path commands found');
            return [];
        }

        console.log(`   🔍 Found ${commands.length} path commands`);

        let currentX = 0, currentY = 0;
        let pathStartX = 0, pathStartY = 0;
        let hasMoveTo = false;

        commands.forEach((cmd, index) => {
            const type = cmd[0].toUpperCase();
            const isRelative = cmd[0] !== cmd[0].toUpperCase();
            const coords = cmd.slice(1).trim().split(/[\s,]+/).map(parseFloat).filter(n => !isNaN(n));

            switch (type) {
                case 'M': // Move to
                    if (coords.length >= 2) {
                        currentX = isRelative ? currentX + coords[0] : coords[0];
                        currentY = isRelative ? currentY + coords[1] : coords[1];
                        pathStartX = currentX;
                        pathStartY = currentY;

                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight,
                            moveTo: true
                        };
                        points.push(point);
                        hasMoveTo = true;
                    }
                    break;

                case 'L': // Line to
                    if (coords.length >= 2) {
                        currentX = isRelative ? currentX + coords[0] : coords[0];
                        currentY = isRelative ? currentY + coords[1] : coords[1];

                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight
                        };
                        points.push(point);
                    }
                    break;

                case 'H': // Horizontal line
                    if (coords.length >= 1) {
                        currentX = isRelative ? currentX + coords[0] : coords[0];
                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight
                        };
                        points.push(point);
                    }
                    break;

                case 'V': // Vertical line
                    if (coords.length >= 1) {
                        currentY = isRelative ? currentY + coords[0] : coords[0];
                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight
                        };
                        points.push(point);
                    }
                    break;

                case 'C': // Cubic Bezier curve
                    if (coords.length >= 6) {
                        const x1 = isRelative ? currentX + coords[0] : coords[0];
                        const y1 = isRelative ? currentY + coords[1] : coords[1];
                        const x2 = isRelative ? currentX + coords[2] : coords[2];
                        const y2 = isRelative ? currentY + coords[3] : coords[3];
                        currentX = isRelative ? currentX + coords[4] : coords[4];
                        currentY = isRelative ? currentY + coords[5] : coords[5];

                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight,
                            curve: {
                                type: 'cubic',
                                x1: (x1 / viewBoxWidth) * shapeWidth,
                                y1: (y1 / viewBoxHeight) * shapeHeight,
                                x2: (x2 / viewBoxWidth) * shapeWidth,
                                y2: (y2 / viewBoxHeight) * shapeHeight
                            }
                        };
                        points.push(point);
                    }
                    break;

                case 'Q': // Quadratic Bezier curve
                    if (coords.length >= 4) {
                        const x1 = isRelative ? currentX + coords[0] : coords[0];
                        const y1 = isRelative ? currentY + coords[1] : coords[1];
                        currentX = isRelative ? currentX + coords[2] : coords[2];
                        currentY = isRelative ? currentY + coords[3] : coords[3];

                        const point = {
                            x: (currentX / viewBoxWidth) * shapeWidth,
                            y: (currentY / viewBoxHeight) * shapeHeight,
                            curve: {
                                type: 'quadratic',
                                x1: (x1 / viewBoxWidth) * shapeWidth,
                                y1: (y1 / viewBoxHeight) * shapeHeight
                            }
                        };
                        points.push(point);
                    }
                    break;

                case 'Z': // Close path
                    if (hasMoveTo) {
                        points.push({ close: true });
                    }
                    currentX = pathStartX;
                    currentY = pathStartY;
                    break;
            }
        });

        console.log(`   ✅ Converted to ${points.length} points (moveTo: ${hasMoveTo})`);
        return points;

    } catch (error) {
        console.error('   ❌ Error converting SVG path to points:', error.message);
        return [];
    }
}

/**
 * FIXED: Create shape options with proper stroke-width="0" handling
 */
function createDynamicShapeOptions(element, slideContext, points, svgStyles) {
    const style = element.style;

    // Extract positioning and dimensions
    const left = parseFloat(style.left) || 0;
    const top = parseFloat(style.top) || 0;
    const width = parseFloat(style.width) || 100;
    const height = parseFloat(style.height) || 100;

    // Convert to inches (72 DPI)
    const x = left / 72;
    const y = top / 72;
    const w = width / 72;
    const h = height / 72;

    // Apply slide context scaling if available
    const finalX = slideContext?.scaleX ? x * slideContext.scaleX : x;
    const finalY = slideContext?.scaleY ? y * slideContext.scaleY : y;
    const finalW = slideContext?.scaleX ? w * slideContext.scaleX : w;
    const finalH = slideContext?.scaleY ? h * slideContext.scaleY : h;

    // Extract styling from SVG with proper transparent handling
    const fillColor = extractColor(svgStyles.fill);
    const stroke = extractColor(svgStyles.stroke);
    
    // FIXED: Properly handle stroke-width="0" - don't use || 1 fallback
    const strokeWidthRaw = svgStyles.strokeWidth;
    let strokeWidth = 0;
    if (strokeWidthRaw !== undefined && strokeWidthRaw !== null && strokeWidthRaw !== '') {
        strokeWidth = parseFloat(strokeWidthRaw);
        // If parseFloat fails, strokeWidth remains 0
        if (isNaN(strokeWidth)) {
            strokeWidth = 0;
        }
    }

    console.log(`   🎨 Stroke analysis: strokeWidthRaw="${strokeWidthRaw}", parsed=${strokeWidth}`);

    // Extract opacity
    const opacity = parseFloat(style.opacity || svgStyles.opacity || '1');
    const transparency = opacity < 1 ? Math.round((1 - opacity) * 100) : 0;

    // Create shape options for custom geometry
    const shapeOptions = {
        x: Math.round(finalX * 100) / 100,
        y: Math.round(finalY * 100) / 100,
        w: Math.round(finalW * 100) / 100,
        h: Math.round(finalH * 100) / 100,
        points: points
    };

    // Add fill if it's not transparent/none
    if (fillColor !== null) {
        shapeOptions.fill = fillColor;
        console.log(`   🎨 Added fill: ${fillColor}`);
    }

    // FIXED: Only add stroke if strokeWidth > 0 (not just truthy)
    if (stroke && stroke !== 'none' && strokeWidth > 0) {
        shapeOptions.line = {
            color: stroke,
            width: Math.min(strokeWidth, 10)
        };
        console.log(`   🎨 Added stroke: ${stroke}, width: ${strokeWidth}`);
    } else {
        console.log(`   🎨 No stroke: stroke=${stroke}, strokeWidth=${strokeWidth}`);
    }

    // Add transparency if needed
    if (transparency > 0 && transparency <= 100) {
        shapeOptions.transparency = transparency;
    }

    // Add rotation if present
    const transform = style.transform || '';
    const rotateMatch = transform.match(/rotate\((-?\d*\.?\d+)deg\)/);
    if (rotateMatch) {
        const rotation = parseFloat(rotateMatch[1]);
        if (rotation !== 0 && Math.abs(rotation) <= 360) {
            shapeOptions.rotate = Math.round(rotation);
        }
    }

    return shapeOptions;
}

/**
 * Extract color from CSS color value and return hex without #
 */
function extractColor(colorValue) {
    if (!colorValue || colorValue === 'transparent' || colorValue === 'none') {
        return null;
    }

    // Handle hex colors
    if (colorValue.startsWith('#')) {
        const hex = colorValue.substring(1).toUpperCase();
        return hex.length === 6 ? hex : (hex.length === 3 ?
            hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2] : null);
    }

    // Handle rgb/rgba colors
    const rgbMatch = colorValue.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgbMatch) {
        const r = Math.max(0, Math.min(255, parseInt(rgbMatch[1]))).toString(16).padStart(2, '0');
        const g = Math.max(0, Math.min(255, parseInt(rgbMatch[2]))).toString(16).padStart(2, '0');
        const b = Math.max(0, Math.min(255, parseInt(rgbMatch[3]))).toString(16).padStart(2, '0');
        return `${r}${g}${b}`.toUpperCase();
    }

    // Handle named colors
    const namedColors = {
        'white': 'FFFFFF', 'black': '000000', 'red': 'FF0000', 'green': '008000',
        'blue': '0000FF', 'yellow': 'FFFF00', 'cyan': '00FFFF', 'magenta': 'FF00FF',
        'silver': 'C0C0C0', 'gray': '808080', 'grey': '808080', 'maroon': '800000',
        'olive': '808000', 'purple': '800080', 'teal': '008080', 'navy': '000080'
    };

    return namedColors[colorValue.toLowerCase()] || null;
}

/**
 * Create fallback shape with proper transparent handling
 */
function createFallbackShape(shapeOptions, fallbackType = 'rect') {
    const fallback = {
        x: shapeOptions.x,
        y: shapeOptions.y,
        w: shapeOptions.w,
        h: shapeOptions.h
    };

    // Only add fill if it exists in the original options
    if (shapeOptions.fill) {
        fallback.fill = shapeOptions.fill;
    }

    if (shapeOptions.line) {
        fallback.line = shapeOptions.line;
    }

    if (shapeOptions.transparency) {
        fallback.transparency = shapeOptions.transparency;
    }

    if (shapeOptions.rotate) {
        fallback.rotate = shapeOptions.rotate;
    }

    return { fallback, type: fallbackType };
}

/**
 * FIXED: Validate shape options without adding unwanted strokes
 */
function validateShapeOptions(shapeOptions) {
    console.log('   🔍 Validating shape options...');

    // Check dimensions
    if (shapeOptions.w <= 0 || shapeOptions.h <= 0) {
        console.log(`   ❌ Invalid dimensions: ${shapeOptions.w}x${shapeOptions.h}`);
        return false;
    }

    // Check position bounds and fix if needed
    if (shapeOptions.x < 0 || shapeOptions.y < 0) {
        console.log(`   ⚠️ Negative position detected: (${shapeOptions.x}, ${shapeOptions.y}), fixing...`);
        shapeOptions.x = Math.max(0, shapeOptions.x);
        shapeOptions.y = Math.max(0, shapeOptions.y);
    }

    // Check points array
    if (!shapeOptions.points || shapeOptions.points.length === 0) {
        console.log('   ❌ No points array or empty points');
        return false;
    }

    // Ensure we have a valid moveTo command
    const hasValidMoveTo = shapeOptions.points.some(p => p.moveTo === true);
    if (!hasValidMoveTo) {
        console.log('   ❌ No valid moveTo command found in points');
        return false;
    }

    // FIXED: Check visibility but don't force add stroke if fill exists
    const hasFill = shapeOptions.fill !== undefined && shapeOptions.fill !== null;
    const hasStroke = shapeOptions.line && shapeOptions.line.width > 0;

    console.log(`   📊 Visibility: fill=${hasFill}, stroke=${hasStroke}`);

    // FIXED: Only add stroke if BOTH fill and stroke are missing (truly invisible shape)
    if (!hasFill && !hasStroke) {
        console.log('   ⚠️ Shape has no fill and no stroke - adding minimal stroke for visibility');
        shapeOptions.line = {
            color: "CCCCCC", // Use light gray instead of black for minimal visibility
            width: 0.5       // Use thinner line
        };
    } else if (hasFill && !hasStroke) {
        console.log('   ✅ Shape has fill but no stroke - this is valid, no stroke added');
    }

    console.log('   ✅ Shape options validation passed');
    return true;
}

/**
 * Main function to add SVG to slide
 */
function addSvgToSlide(pptSlide, svgElement, elementStyle, slideContext) {
    try {
        console.log('   🎯 Processing SVG element...');

        // Extract SVG properties
        const viewBox = svgElement.getAttribute('viewBox');
        const pathElement = svgElement.querySelector('path');

        if (!pathElement) {
            console.log('   ⚠️ No path element found in SVG');
            return false;
        }

        // Parse viewBox with better error handling
        let viewBoxWidth = 100, viewBoxHeight = 100;
        if (viewBox) {
            const viewBoxValues = viewBox.split(/\s+/).map(parseFloat);
            if (viewBoxValues.length >= 4 && !viewBoxValues.some(isNaN)) {
                viewBoxWidth = viewBoxValues[2];
                viewBoxHeight = viewBoxValues[3];
                console.log(`   📏 ViewBox: ${viewBoxWidth}x${viewBoxHeight}`);
            }
        }

        // Extract path data
        const pathData = pathElement.getAttribute('d');
        if (!pathData || pathData.trim() === '') {
            console.log('   ⚠️ No valid path data found');
            return false;
        }

        console.log(`   📝 Path data: ${pathData.substring(0, 100)}${pathData.length > 100 ? '...' : ''}`);

        // Get parent element for positioning
        const parentElement = svgElement.closest('.shape.custom-shape') ||
            svgElement.closest('.custom-shape') ||
            svgElement.closest('#custGeom') ||
            svgElement.parentElement;

        if (!parentElement) {
            console.log('   ⚠️ No valid parent element found');
            return false;
        }

        const style = parentElement.style;
        console.log(`   📍 Parent element: ${parentElement.className}, ID: ${parentElement.id}`);

        // Calculate shape dimensions in inches (72 DPI)
        const shapeWidthInches = (parseFloat(style.width) || 100) / 72;
        const shapeHeightInches = (parseFloat(style.height) || 100) / 72;

        console.log(`   📐 Shape dimensions: ${shapeWidthInches}" x ${shapeHeightInches}"`);

        // Convert SVG path to PowerPoint points
        const points = convertSvgPathToPptxPoints(
            pathData,
            viewBoxWidth,
            viewBoxHeight,
            shapeWidthInches,
            shapeHeightInches
        );

        if (points.length === 0) {
            console.log('   ❌ Failed to convert SVG path to points');
            return false;
        }

        console.log(`   ✅ Converted to ${points.length} points`);

        // FIXED: Extract SVG styling with explicit stroke-width handling
        const svgStyles = {
            fill: pathElement.getAttribute('fill') || svgElement.style.fill,
            stroke: pathElement.getAttribute('stroke') || svgElement.style.stroke,
            strokeWidth: pathElement.getAttribute('stroke-width') || svgElement.style.strokeWidth,
            opacity: pathElement.getAttribute('opacity') || svgElement.style.opacity || '1'
        };

        console.log(`   🎨 SVG styles extracted:`, {
            fill: svgStyles.fill,
            stroke: svgStyles.stroke,
            strokeWidth: svgStyles.strokeWidth,
            opacity: svgStyles.opacity
        });

        // Create shape options
        const shapeOptions = createDynamicShapeOptions(parentElement, slideContext, points, svgStyles);

        // Validate shape options
        if (!validateShapeOptions(shapeOptions)) {
            console.log('   ⚠️ Shape options validation failed, using fallback');
            const { fallback, type } = createFallbackShape(shapeOptions);
            pptSlide.addShape(type, fallback);
            return true;
        }

        // Add object name for debugging
        shapeOptions.objectName = `Custom SVG Shape (${parentElement.className || 'unnamed'})`;

        // Try to add the shape
        try {
            console.log(`   🔧 Attempting to add custGeom shape with options:`, {
                x: shapeOptions.x,
                y: shapeOptions.y,
                w: shapeOptions.w,
                h: shapeOptions.h,
                pointsCount: shapeOptions.points.length,
                fill: shapeOptions.fill,
                hasLine: !!shapeOptions.line,
                lineWidth: shapeOptions.line?.width
            });

            pptSlide.addShape('custGeom', shapeOptions);
            console.log('   ✅ Successfully added custom geometry shape');
            return true;

        } catch (custGeomError) {
            console.log(`   ❌ custGeom failed: ${custGeomError.message}`);

            // Fallback: Try polygon approach
            try {
                const polygonPoints = points.filter(p => !p.curve && !p.close && p.x !== undefined && p.y !== undefined);
                if (polygonPoints.length >= 3) {
                    const simplifiedOptions = { ...shapeOptions };
                    simplifiedOptions.points = polygonPoints;
                    simplifiedOptions.objectName = `Simplified SVG Polygon (${parentElement.className || 'unnamed'})`;

                    pptSlide.addShape('custGeom', simplifiedOptions);
                    console.log('   ✅ Added simplified polygon shape');
                    return true;
                }
            } catch (simplifiedError) {
                console.log(`   ❌ Simplified polygon failed: ${simplifiedError.message}`);
            }

            // Final fallback: Rectangle
            const { fallback, type } = createFallbackShape(shapeOptions, 'rect');
            fallback.objectName = `SVG Fallback Rectangle (${parentElement.className || 'unnamed'})`;
            pptSlide.addShape(type, fallback);
            console.log('   ⚠️ Used rectangle fallback');
            return true;
        }

    } catch (error) {
        console.error('   ❌ Critical error in addSvgToSlide:', error.message);
        console.error('   📍 Stack trace:', error.stack);

        // Emergency fallback
        try {
            const emergencyOptions = {
                x: 1, y: 1, w: 2, h: 1,
                fill: "FFCCCC", // Light red fill instead of border
                objectName: "Emergency SVG Fallback"
            };
            pptSlide.addShape('rect', emergencyOptions);
            console.log('   🚨 Added emergency fallback shape');
            return true;
        } catch (emergencyError) {
            console.error('   💥 Even emergency fallback failed:', emergencyError.message);
            return false;
        }
    }
}

/**
 * Add SVG connector to slide (for connector lines/arrows)
 */
function addSvgConnectorToSlide(pptSlide, svgElement, elementStyle, slideContext) {
    try {
        const pathElement = svgElement.querySelector('path') || svgElement.querySelector('line');
        if (!pathElement) {
            return false;
        }

        // Extract connector properties
        const parentElement = svgElement.closest('.sli-svg-connector') || svgElement.parentElement;
        const style = parentElement.style;

        // Extract positioning (72 DPI)
        const left = parseFloat(style.left) || 0;
        const top = parseFloat(style.top) || 0;
        const width = parseFloat(style.width) || 100;
        const height = parseFloat(style.height) || 100;

        // Convert to inches
        const x = (left / 72) * (slideContext?.scaleX || 1);
        const y = (top / 72) * (slideContext?.scaleY || 1);
        const w = (width / 72) * (slideContext?.scaleX || 1);
        const h = (height / 72) * (slideContext?.scaleY || 1);

        // Extract line styling
        const stroke = extractColor(pathElement.getAttribute('stroke') || '#000000');
        const strokeWidthAttr = pathElement.getAttribute('stroke-width');
        let strokeWidth = 1; // Default for connectors
        if (strokeWidthAttr !== null && strokeWidthAttr !== undefined) {
            const parsed = parseFloat(strokeWidthAttr);
            if (!isNaN(parsed)) {
                strokeWidth = Math.min(parsed, 10);
            }
        }

        // Skip connector if stroke width is 0
        if (strokeWidth <= 0) {
            console.log('   ⚠️ Connector has stroke-width=0, skipping');
            return false;
        }

        // Create line object
        const lineOptions = {
            x: Math.round(x * 100) / 100,
            y: Math.round(y * 100) / 100,
            w: Math.round(w * 100) / 100,
            h: Math.round(h * 100) / 100,
            line: {
                color: stroke || "000000",
                width: strokeWidth
            }
        };

        // Add line to slide
        pptSlide.addShape('line', lineOptions);
        return true;

    } catch (error) {
        console.error('   ❌ Error adding SVG connector to slide:', error.message);
        return false;
    }
}

/**
 * Enhanced SVG processor that handles multiple SVG types
 */
function processSvgElement(pptSlide, element, slideContext) {
    try {
        console.log('   🔍 >>>>>>>>>>>>>>>>>>>>>>>>>>--- >>  Processing SVG element with classes:', element.className);

        const svgElement = element.querySelector('svg');
        if (!svgElement) {
            console.log('   ❌ No SVG element found');
            return false;
        }

        // Log SVG properties for debugging
        const viewBox = svgElement.getAttribute('viewBox');
        const pathElements = svgElement.querySelectorAll('path');
        console.log(`   📊 SVG info: viewBox=${viewBox}, paths=${pathElements.length}`);

        // Determine SVG type based on parent classes
        const isConnector = element.classList.contains('sli-svg-connector');
        const isCustomShape = element.classList.contains('custom-shape') || element.id === 'custGeom';

        if (isConnector) {
            console.log('   🔗 Processing as connector');
            return addSvgConnectorToSlide(pptSlide, svgElement, element.style, slideContext);
        }
        else if (isCustomShape) {
            console.log('   🎨 Processing as custom shape');
            return addSvgToSlide(pptSlide, svgElement, element.style, slideContext);
        } else {
            console.log('   🎨 Processing as default custom shape');
            return addSvgToSlide(pptSlide, svgElement, element.style, slideContext);
        }

    } catch (error) {
        console.error('   ❌ Error processing SVG element:', error.message);
        return false;
    }
}

module.exports = {
    addSvgToSlide,
    addSvgConnectorToSlide,
    processSvgElement,
    convertSvgPathToPptxPoints,
    createDynamicShapeOptions
};