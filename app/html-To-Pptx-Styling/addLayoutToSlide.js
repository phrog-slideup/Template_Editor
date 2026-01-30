const path = require("path");
const fs = require("fs");

// New function to get slide dimensions
function getSlideDimensions(pptSlide) {
    try {
        // Default PowerPoint slide dimensions (in inches)
        let width = 10; // Default width
        let height = 7.5; // Default height

        // Try to get dimensions from slide object
        if (pptSlide && pptSlide.slideLayout) {
            const layout = pptSlide.slideLayout;
            if (layout.width) width = layout.width;
            if (layout.height) height = layout.height;
        }

        // Try alternative property names
        if (pptSlide && pptSlide.width) width = pptSlide.width;
        if (pptSlide && pptSlide.height) height = pptSlide.height;

        // Convert to pixels (72 DPI)
        const widthPx = width * 72;
        const heightPx = height * 72;

        return {
            width: width,
            height: height,
            widthPx: widthPx,
            heightPx: heightPx,
            widthInches: width,
            heightInches: height
        };
    } catch (error) {
        console.warn('Error getting slide dimensions:', error.message);
        // Return default dimensions
        return {
            width: 10,
            height: 7.5,
            widthPx: 960,
            heightPx: 720,
            widthInches: 10,
            heightInches: 7.5
        };
    }
}

function addLayoutToSlide(pptx, pptSlide, element) {
    const layoutStyles = element.style;

    // Get slide dimensions for proper scaling
    const slideDimensions = getSlideDimensions(pptSlide);

    // Extract layout properties with better parsing
    const width = parseFloat(layoutStyles.width) || 0;
    const height = parseFloat(layoutStyles.height) || 0;
    const left = parseFloat(layoutStyles.left) || 0;
    const top = parseFloat(layoutStyles.top) || 0;
    
    // Convert pixels to inches for PowerPoint
    const x = left / 72;
    const y = top / 72;
    const w = width / 72;
    const h = height / 72;

    // Extract styling properties
    const bgColor = layoutStyles.backgroundColor || "#FFFFFF";
    const borderColor = layoutStyles.borderColor || "#000000";
    const borderWidth = parseFloat(layoutStyles.borderWidth) || 0;
    const borderRadius = layoutStyles.borderRadius || "0px";
    const opacity = parseFloat(layoutStyles.opacity || "1");
    
    // Handle rotation
    const transform = layoutStyles.transform || "";
    const rotationMatch = transform.match(/rotate\(([-\d.]+)deg\)/);
    const rotation = rotationMatch ? parseFloat(rotationMatch[1]) : 0;

    // Convert opacity to transparency percentage
    const transparencyPercentage = Math.round((1 - opacity) * 100);

    // Parse border radius
    let borderRadiusValue = 0;
    if (borderRadius.includes("px")) {
        borderRadiusValue = parseFloat(borderRadius) / 72; // Convert px to inches
        const smallerDimension = Math.min(w, h);
        borderRadiusValue = smallerDimension > 0 ? borderRadiusValue / smallerDimension : 0;
        borderRadiusValue = Math.min(borderRadiusValue, 1); // Cap at 1
    } else if (borderRadius.includes("%")) {
        borderRadiusValue = parseFloat(borderRadius) / 100;
    }

    // Layout options for PowerPoint
    const layoutOptions = {
        x: x,
        y: y,
        w: w,
        h: h,
        fill: {
            color: rgbToHex(bgColor),
            transparency: transparencyPercentage
        },
        line: {
            color: rgbToHex(borderColor),
            width: borderWidth / 72 // Convert to inches
        }
    };

    // Add rotation if present
    if (rotation !== 0) {
        layoutOptions.rotate = rotation;
    }

    // Add border radius if present
    if (borderRadiusValue > 0) {
        layoutOptions.rectRadius = borderRadiusValue;
    }

    // Store layout data for XML generation
    const layoutData = {
        type: 'layout',
        id: element.id || `layout_${Date.now()}`,
        x: x,
        y: y,
        w: w,
        h: h,
        rotation: rotation,
        fill: {
            color: rgbToHex(bgColor),
            transparency: transparencyPercentage
        },
        line: {
            color: rgbToHex(borderColor),
            width: borderWidth / 72
        },
        borderRadius: borderRadiusValue,
        className: element.className || '',
        zIndex: parseInt(layoutStyles.zIndex) || 0
    };

    try {
        // Add the layout as a shape to the slide
        pptSlide.addShape(pptx.shapes.RECTANGLE, layoutOptions);

        // Store layout data on slide for XML generation
        if (!pptSlide.layouts) pptSlide.layouts = [];
        pptSlide.layouts.push(layoutData);

        console.log(`Layout added successfully: ${w}" x ${h}" at (${x}", ${y}")`);

    } catch (error) {
        console.error(`Failed to add layout to slide:`, error);
        
        // Store error data
        const errorData = {
            type: 'layout-error',
            id: element.id || `layout_error_${Date.now()}`,
            x: x,
            y: y,
            w: w,
            h: h,
            error: error.message
        };

        if (!pptSlide.layouts) pptSlide.layouts = [];
        pptSlide.layouts.push(errorData);
    }
}

// Function to generate layout XML for PPTX
function generateLayoutXml(layout, layoutId, slideDimensions) {
    if (layout.type === 'layout-error') {
        return generateLayoutErrorXml(layout, layoutId);
    }

    const x = Math.round((layout.x || 0) * 914400); // Convert inches to EMU
    const y = Math.round((layout.y || 0) * 914400);
    const w = Math.round((layout.w || 1) * 914400);
    const h = Math.round((layout.h || 1) * 914400);
    
    const rotation = layout.rotation ? Math.round(layout.rotation * 60000) : 0; // Convert degrees to 1/60000th
    
    // Handle fill
    const fillColor = layout.fill && layout.fill.color ? layout.fill.color.replace('#', '') : 'FFFFFF';
    const fillTransparency = layout.fill && layout.fill.transparency ? layout.fill.transparency : 0;
    
    let fillXml = '';
    if (fillTransparency >= 99) {
        fillXml = '<a:noFill/>';
    } else {
        fillXml = `
            <a:solidFill>
                <a:srgbClr val="${fillColor}">
                    ${fillTransparency > 0 ? `<a:alpha val="${100 - fillTransparency}000"/>` : ''}
                </a:srgbClr>
            </a:solidFill>`;
    }

    // Handle line/border
    const lineColor = layout.line && layout.line.color ? layout.line.color.replace('#', '') : '000000';
    const lineWidth = layout.line && layout.line.width ? Math.round(layout.line.width * 914400) : 0; // Convert to EMU
    
    let lineXml = '';
    if (lineWidth > 0) {
        lineXml = `
            <a:ln w="${lineWidth}">
                <a:solidFill>
                    <a:srgbClr val="${lineColor}"/>
                </a:solidFill>
            </a:ln>`;
    } else {
        lineXml = '<a:ln><a:noFill/></a:ln>';
    }

    const layoutXml = `
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="${layoutId}" name="Layout ${layoutId}"/>
                <p:cNvSpPr/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm${rotation !== 0 ? ` rot="${rotation}"` : ''}>
                    <a:off x="${x}" y="${y}"/>
                    <a:ext cx="${w}" cy="${h}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
                ${fillXml}
                ${lineXml}
            </p:spPr>
        </p:sp>`;

    return layoutXml;
}

// Function to generate layout error XML
function generateLayoutErrorXml(errorData, shapeId) {
    const x = Math.round((errorData.x || 0) * 914400);
    const y = Math.round((errorData.y || 0) * 914400);
    const w = Math.round((errorData.w || 1) * 914400);
    const h = Math.round((errorData.h || 1) * 914400);

    const errorXml = `
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="${shapeId}" name="Layout Error ${shapeId}"/>
                <p:cNvSpPr/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm>
                    <a:off x="${x}" y="${y}"/>
                    <a:ext cx="${w}" cy="${h}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
                <a:solidFill>
                    <a:srgbClr val="FFCCCC"/>
                </a:solidFill>
                <a:ln w="12700">
                    <a:solidFill>
                        <a:srgbClr val="FF0000"/>
                    </a:solidFill>
                </a:ln>
            </p:spPr>
            <p:txBody>
                <a:bodyPr wrap="square" rtlCol="0"/>
                <a:lstStyle/>
                <a:p>
                    <a:r>
                        <a:rPr lang="en-US" sz="1200">
                            <a:solidFill>
                                <a:srgbClr val="000000"/>
                            </a:solidFill>
                        </a:rPr>
                        <a:t>Layout Error: ${escapeXml(errorData.error?.substring(0, 50) || 'Unknown error')}</a:t>
                    </a:r>
                </a:p>
            </p:txBody>
        </p:sp>`;

    return errorXml;
}

// Function to generate all layouts XML for a slide
function generateLayoutsXml(layouts) {
    let layoutsXml = '';
    layouts.forEach((layout, index) => {
        const layoutId = index + 400; // Start layout IDs from 400
        layoutsXml += generateLayoutXml(layout, layoutId);
    });
    return layoutsXml;
}

// Function to create slide XML with layouts
function createSlideXmlWithLayouts(slideData, slideDimensions) {
    const { width, height } = slideDimensions;
    const slideId = slideData.slideId || 1;
    
    const widthEmu = Math.round(width * 914400);
    const heightEmu = Math.round(height * 914400);

    const layoutsXml = generateLayoutsXml(slideData.layouts || []);

    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="${widthEmu}" cy="${heightEmu}"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="${widthEmu}" cy="${heightEmu}"/>
                </a:xfrm>
            </p:grpSpPr>
            ${layoutsXml}
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`;

    return slideXml;
}

// Enhanced color conversion function
function rgbToHex(color) {
    if (!color) return '#FFFFFF'; // Default to white if no color provided

    // If already hex, return as is
    if (color.startsWith('#')) {
        return color;
    }

    // Handle rgba colors
    if (color.startsWith('rgba')) {
        const rgbaRegex = /rgba\((\d+),\s*(\d+),\s*(\d+),\s*([0-9.]+)\)/;
        const match = color.match(rgbaRegex);

        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);
            // We ignore alpha when converting to hex

            return '#' +
                r.toString(16).padStart(2, '0') +
                g.toString(16).padStart(2, '0') +
                b.toString(16).padStart(2, '0');
        }
    }

    // Handle rgb colors
    if (color.startsWith('rgb')) {
        const rgbRegex = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/;
        const match = color.match(rgbRegex);

        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);

            return '#' +
                r.toString(16).padStart(2, '0') +
                g.toString(16).padStart(2, '0') +
                b.toString(16).padStart(2, '0');
        }
    }

    // Handle named colors
    const namedColors = {
        'white': '#FFFFFF',
        'black': '#000000',
        'red': '#FF0000',
        'green': '#00FF00',
        'blue': '#0000FF',
        'yellow': '#FFFF00',
        'cyan': '#00FFFF',
        'magenta': '#FF00FF',
        'gray': '#808080',
        'grey': '#808080',
        'transparent': '#FFFFFF'
    };

    if (namedColors[color.toLowerCase()]) {
        return namedColors[color.toLowerCase()];
    }

    return '#FFFFFF'; // Default fallback to white
}

// Helper function to escape XML special characters
function escapeXml(text) {
    if (!text) return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Function to normalize style values
function normalizeStyleValue(value, defaultValue = 0) {
    if (!value) return defaultValue;
    
    // Remove units and parse as float
    const numericValue = parseFloat(value.toString().replace(/[^\d.-]/g, ''));
    return isNaN(numericValue) ? defaultValue : numericValue;
}

// Function to get layout statistics
function getLayoutStatistics(layouts) {
    if (!layouts || layouts.length === 0) {
        return {
            count: 0,
            totalArea: 0,
            errors: 0
        };
    }

    let totalArea = 0;
    let errors = 0;

    layouts.forEach(layout => {
        if (layout.type === 'layout-error') {
            errors++;
        } else {
            totalArea += (layout.w || 0) * (layout.h || 0);
        }
    });

    return {
        count: layouts.length,
        totalArea: totalArea,
        errors: errors,
        avgArea: layouts.length > 0 ? totalArea / (layouts.length - errors) : 0
    };
}

module.exports = {
    addLayoutToSlide,              // Original enhanced function
    getSlideDimensions,            // NEW: Slide dimension detection
    generateLayoutXml,             // NEW: Individual layout XML
    generateLayoutsXml,            // NEW: Multiple layouts XML
    generateLayoutErrorXml,        // NEW: Layout error XML
    createSlideXmlWithLayouts,     // NEW: Complete slide XML with layouts
    rgbToHex,                      // Enhanced color conversion
    escapeXml,                     // NEW: XML escaping utility
    normalizeStyleValue,           // NEW: Style value normalization
    getLayoutStatistics            // NEW: Layout statistics
};