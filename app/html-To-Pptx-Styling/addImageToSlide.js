const path = require("path");
const fs = require("fs");
const axios = require("axios");
const normalizeStyle = require("../api/helper/colorHelper");
const clrHelper = require("../api/helper/colorHelper.js");
async function addImageToSlide(pptx, pptSlide, imgElement, slideContext) {
    console.log('Adding image to slide:', imgElement);
    let src = imgElement.getAttribute("src");

    const style = imgElement.style;
    const parent = imgElement.closest(".image-container");
    const objName = parent.getAttribute("data-name");

    const altText = parent?.getAttribute("data-alt-text") || '';
    console.log('🏷️ Alt text from HTML:', altText);

    const parentStyle = parent ? parent.style : {};

    const slideDimensions = getSlideDimensions(pptSlide);

    const parentLeft = normalizeStyle.normalizeStyleValue(parentStyle.left, 0);
    const parentTop = normalizeStyle.normalizeStyleValue(parentStyle.top, 0);
    const x = parentLeft / 72;
    const y = parentTop / 72;

    let w = normalizeStyle.normalizeStyleValue(parentStyle.width || "100");
    let h = normalizeStyle.normalizeStyleValue(parentStyle.height || "100");

    w = w / 72;
    h = h / 72;

    if (w === 0 || h === 0) {
        console.warn(`Warning: Image (${src}) has zero width or height! Adjusting.`);
        if (w === 0) w = 100 / 72;
        if (h === 0) h = 100 / 72;
    }

    // --- TRANSFORM BLOCK ---
    const transformStr = (parentStyle.transform || style.transform || "").trim();

    let cssRotation = 0;
    let cssFlipH = false;
    let cssFlipV = false;

    const rotateMatch = transformStr.match(/rotate\(([-\d.]+)deg\)/i);
    if (rotateMatch) cssRotation = parseFloat(rotateMatch[1]);

    cssFlipH = /\bscaleY\(-1\)/.test(transformStr);
    cssFlipV = /\bscaleX\(-1\)/.test(transformStr);

    let flipBeforeRotate = false;
    if (rotateMatch) {
        const rotateIndex = transformStr.indexOf('rotate');
        const scaleXIndex = transformStr.indexOf('scaleX(-1)');
        const scaleYIndex = transformStr.indexOf('scaleY(-1)');

        if ((cssFlipH && scaleXIndex !== -1 && scaleXIndex < rotateIndex) ||
            (cssFlipV && scaleYIndex !== -1 && scaleYIndex < rotateIndex)) {
            flipBeforeRotate = true;
        }
    }

    let finalRotation = cssRotation;
    if (Math.abs(cssRotation) === 180) {
        if (cssFlipH && cssFlipV) {
            cssFlipH = false;
            cssFlipV = false;
        } else if (cssFlipH || cssFlipV) {
            const tmp = cssFlipH;
            cssFlipH = cssFlipV;
            cssFlipV = tmp;
        }
    }

    if (flipBeforeRotate && (Math.abs(cssRotation) === 90 || Math.abs(cssRotation) === 270)) {
        finalRotation = cssRotation;
    }

    const rotation = finalRotation;

    if (!rotation && cssFlipH === true && cssFlipV === false) {
        cssFlipH = false;
        cssFlipV = true;
    }
    else if (!rotation && cssFlipH === false && cssFlipV === true) {
        cssFlipH = true;
        cssFlipV = false;
    }

    // Extract shadow/glow from box-shadow CSS
    let pptxShadow = null;
    if (parent) {
        let boxShadowValue = parentStyle.boxShadow;

        if (!boxShadowValue || boxShadowValue === 'none') {
            const styleAttr = parent.getAttribute('style') || '';
            const match = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
            if (match && match[1]) {
                boxShadowValue = match[1].trim();
            }
        }

        console.log('=== SHADOW EXTRACTION ===');
        console.log('Box-shadow value:', boxShadowValue);

        if (boxShadowValue && boxShadowValue !== 'none') {
            pptxShadow = convertBoxShadowToPptxFormat(boxShadowValue);
            console.log('✅ Converted to pptxgenjs format:', pptxShadow);
        }
    }

    let flipH = cssFlipH;
    let flipV = cssFlipV;

    // Extract border properties
    const borderWidthPx = parseFloat(parentStyle.borderWidth || "0");
    const borderWidth = borderWidthPx / 72;
    const borderColor = clrHelper.rgbToHex(parentStyle.borderColor || "#000000");

    let borderStyle = parentStyle.borderStyle || 'solid';
    if (!borderStyle || borderStyle === '') {
        const styleAttr = parent?.getAttribute('style') || '';
        const borderStyleMatch = styleAttr.match(/border-style\s*:\s*([^;]+)/i);
        if (borderStyleMatch && borderStyleMatch[1]) {
            borderStyle = borderStyleMatch[1].trim();
        } else {
            const borderMatch = styleAttr.match(/border\s*:\s*[^;]*\s+(solid|dotted|dashed|double|groove|ridge|inset|outset)/i);
            if (borderMatch && borderMatch[1]) {
                borderStyle = borderMatch[1].trim();
            }
        }
    }

    const borderRadius = parentStyle.borderRadius || "0%";
    const isRounded = borderRadius === "50%";

    // Extract opacity
    let containerOpacity = 1;
    if (parent) {
        const parentStyleAttr = parent.getAttribute('style') || '';
        const containerOpacityMatch = parentStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
        if (containerOpacityMatch && containerOpacityMatch[1]) {
            containerOpacity = parseFloat(containerOpacityMatch[1]);
        }
    }

    let imgOpacity = 1;
    const imgStyleAttr = imgElement.getAttribute('style') || '';
    const imgOpacityMatch = imgStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
    if (imgOpacityMatch && imgOpacityMatch[1]) {
        imgOpacity = parseFloat(imgOpacityMatch[1]);
    }

    const finalOpacity = containerOpacity * imgOpacity;
    const transparencyPercentage = Math.round((1 - finalOpacity) * 100);

    // ========================================
    // ✅ STEP 1: Extract srcRect values from HTML
    // ========================================
    let srcRectL = '';
    let srcRectR = '';
    let srcRectT = '';
    let srcRectB = '';

    if (parent) {
        srcRectL = parent.getAttribute('srcrectl') || '';
        srcRectR = parent.getAttribute('srcrectr') || '';
        srcRectT = parent.getAttribute('srcrectt') || '';
        srcRectB = parent.getAttribute('srcrectb') || '';

        console.log('🔍 Extracted srcRect from parent:', { srcRectL, srcRectR, srcRectT, srcRectB });
    }
    // Fallback to img element
    if (!srcRectL && !srcRectR && !srcRectT && !srcRectB) {
        srcRectL = imgElement.getAttribute('srcrectl') || imgElement.getAttribute('srcRectL') || '';
        srcRectR = imgElement.getAttribute('srcrectr') || imgElement.getAttribute('srcRectR') || '';
        srcRectT = imgElement.getAttribute('srcrectt') || imgElement.getAttribute('srcRectT') || '';
        srcRectB = imgElement.getAttribute('srcrectb') || imgElement.getAttribute('srcRectB') || '';

        console.log('🔍 Fallback: Extracted srcRect from img:', { srcRectL, srcRectR, srcRectT, srcRectB });
    }

    // ========================================
    // ✅ STEP 2: Convert srcRect to pptxgenjs crop format
    // PowerPoint srcRect values are in 1/100,000ths of the image size
    // pptxgenjs expects decimal values (0.0 to 1.0)
    // ========================================
    // Parse srcRect values (can be negative!)
    const parsedL = parseInt(srcRectL) || 0;
    const parsedR = parseInt(srcRectR) || 0;
    const parsedT = parseInt(srcRectT) || 0;
    const parsedB = parseInt(srcRectB) || 0;

    const hasCrop = parsedL !== 0 || parsedR !== 0 || parsedT !== 0 || parsedB !== 0;

    console.log('📊 srcRect from HTML:', { parsedL, parsedR, parsedT, parsedB });



    // Extract hyperlink
    let hyperlink = null;
    if (parent) {
        hyperlink = parent.getAttribute('data-hyperlink');
        if (hyperlink) {
            console.log('✅ Hyperlink extracted from HTML:', hyperlink);
        }
    }

    try {
        let base64Data = null;

        // Image loading logic (unchanged)
        if (src.includes("uploads")) {
            try {
                const imageName = path.basename(src);
                let imagePath = null;

                const possiblePaths = [
                    src,
                    path.join("uploads", imageName),
                    path.join(__dirname, "..", "uploads", imageName),
                    path.join(__dirname, "..", src)
                ];

                for (const tryPath of possiblePaths) {
                    try {
                        if (fs.existsSync(tryPath)) {
                            imagePath = tryPath;
                            break;
                        }
                    } catch (e) { }
                }

                if (!imagePath) {
                    throw new Error(`Image file not found: ${src}`);
                }

                const imageDataBuffer = fs.readFileSync(imagePath);
                const mimeType = getMimeType(imagePath);
                base64Data = `data:${mimeType};base64,${imageDataBuffer.toString("base64")}`;
            } catch (localError) {
                console.error(`Error processing local image (${src}):`, localError);
                throw localError;
            }
        } else if (src.startsWith("http")) {
            try {
                const response = await axios.get(src, { responseType: "arraybuffer" });
                const mimeType = response.headers["content-type"] || "image/png";
                base64Data = `data:${mimeType};base64,${Buffer.from(response.data).toString("base64")}`;
            } catch (remoteError) {
                console.error(`Error fetching remote image (${src}):`, remoteError);
                throw remoteError;
            }
        } else if (src.startsWith("data:image/")) {
            base64Data = src;
        } else {
            throw new Error(`Unsupported image source format: ${src}`);
        }

        if (!base64Data) {
            throw new Error("Failed to create image data");
        }

        // ========================================
        // ✅ STEP 3: Build image options for pptxgenjs
        // ========================================
        const imageOptions = {
            data: base64Data,
            x: x,
            y: y,
            w: w,
            h: h,
            rotate: rotation,
            flipH: flipH,
            flipV: flipV,
            rounding: isRounded,
            transparency: transparencyPercentage,
            objectName: objName || '',
            altText: altText  // ✅ ADD THIS
        };

        // ========================================
        // ✅ CRITICAL: Add sizing with crop
        // pptxgenjs requires sizing.type = 'crop' with x, y, w, h in EMU units (0-100000)
        // ========================================
        if (hasCrop) {
            // Store the raw srcRect values for later XML manipulation
            // imageOptions.sizing = {
            //     type: 'crop',
            //     x: parsedL * 50,
            //     y: parsedT * 50,
            //     w: (100000 - parsedL - parsedR) * 50,
            //     h: (100000 - parsedT - parsedB) * 50
            // };

            // console.log('✅ Applied sizing:', imageOptions.sizing);
            // console.log('✅ Applied cropping to pptxgenjs:', imageOptions.sizing);
            imageOptions._srcRect = {
                l: parsedL,
                r: parsedR,
                t: parsedT,
                b: parsedB
            };
            console.log('🔧 Stored srcRect for manual XML injection:', imageOptions._srcRect);
        } else {
            // No crop - use cover to maintain aspect ratio
            imageOptions.sizing = { type: 'cover' };
        }

        // Add border
        if (borderWidth > 0) {
            imageOptions.line = {
                color: borderColor.replace('#', ''),
                width: borderWidth,
                type: borderStyle || 'solid'
            };
        }

        // Add shadow if exists
        if (pptxShadow) {
            imageOptions.shadow = pptxShadow;
            console.log('✅ Shadow added to pptxgenjs:', pptxShadow);
        }

        // Add hyperlink if exists
        if (hyperlink) {
            imageOptions.hyperlink = { url: hyperlink };
            console.log('✅ Hyperlink added to image:', hyperlink);
        }

        // Add image to slide
        pptSlide.addImage(imageOptions);
        // ✅ NOW ADD BORDER as a separate transparent rectangle shape
        // if (borderWidth > 0) {
        //     console.log('Adding border rectangle:', {
        //         color: borderColor.replace('#', ''),
        //         width: borderWidth * 72 // Convert inches back to points
        //     });

        //     pptSlide.addShape(pptx.shapes.RECTANGLE, {
        //         x: x,
        //         y: y,
        //         w: w,
        //         h: h,
        //         fill: { transparency: 100 }, // Completely transparent
        //         line: {
        //             color: borderColor.replace('#', ''),
        //             width: borderWidth * 72, // Points (1pt ≈ 1/72 inch)
        //             dashType: 'solid'
        //         },
        //         rotate: rotation // Match image rotation
        //     });
        // }
        // ✅ ADD BORDER with style support and fallback
        if (borderWidth > 0) {
            // Map CSS border styles to pptxgenjs dashType
            const borderStyleMap = {
                'solid': 'solid',
                'dotted': 'dot',
                'dashed': 'dash',
                'double': 'solid',      // Fallback: use solid
                'groove': 'solid',      // Fallback: use solid
                'ridge': 'solid',       // Fallback: use solid
                'inset': 'solid',       // Fallback: use solid
                'outset': 'solid',      // Fallback: use solid
                'none': 'solid',        // Fallback: use solid
                'hidden': 'solid'       // Fallback: use solid
            };

            // Get pptxgenjs dashType with fallback to 'solid'
            const pptxDashType = borderStyleMap[borderStyle.toLowerCase()] || 'solid';

            // Adjust width for dotted/dashed styles (they look thinner)
            let adjustedWidth = borderWidth * 72;
            if (pptxDashType === 'dot' || pptxDashType === 'dash') {
                adjustedWidth = Math.max(adjustedWidth * 1.2, 1); // Boost by 20%, minimum 1pt
            }

            console.log('Adding border rectangle:', {
                color: borderColor.replace('#', ''),
                width: adjustedWidth,
                dashType: pptxDashType,
                originalStyle: borderStyle
            });

            try {
                const parent = imgElement.closest(".image-container");
                const objName = parent.getAttribute("data-name");

                const parentStyle = parent ? parent.style : {};


                pptSlide.addShape(pptx.shapes.RECTANGLE, {
                    x: x,
                    y: y,
                    w: w,
                    h: h,
                    fill: { transparency: 100 }, // Completely transparent
                    line: {
                        color: borderColor.replace('#', ''),
                        width: adjustedWidth,
                        dashType: pptxDashType
                    },
                    rotate: rotation // Match image rotation
                });

                console.log('✅ Border added successfully');
            } catch (borderError) {
                // Fallback: try with just solid border if specific style fails
                console.warn('Border style failed, falling back to solid:', borderError);
                try {
                    pptSlide.addShape(pptx.shapes.RECTANGLE, {
                        x: x,
                        y: y,
                        w: w,
                        h: h,
                        fill: { transparency: 100 },
                        line: {
                            color: borderColor.replace('#', ''),
                            width: borderWidth * 72,
                            dashType: 'solid' // Fallback to solid
                        },
                        rotate: rotation
                    });
                    console.log('✅ Border added with solid fallback');
                } catch (fallbackError) {
                    console.error('❌ Border completely failed:', fallbackError);
                }
            }
        }

    } catch (error) {
        console.error(`❌ Failed to add image (${src}) to slide:`, error);
        pptSlide.addShape(pptx.shapes.RECTANGLE, {
            x: x,
            y: y,
            w: w,
            h: h,
            fill: { color: "#EEEEEE" },
            line: { color: "#FF0000", width: 1 / 72 },
            text: `Image Error: ${error.message?.substring(0, 50) || 'Unknown error'}`
        });
    }
}

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

// Add function to handle background images
async function addBackgroundImageToSlide(pptx, pptSlide, bgElement) {
    // Extract the background image URL from the style
    const style = bgElement.style || {};
    const backgroundImageStyle = style.backgroundImage || "";
    const urlMatch = backgroundImageStyle.match(/url\(['"]?([^'"]+)['"]?\)/);

    if (!urlMatch) {
        console.warn("No background image found in element");
        return;
    }

    const src = urlMatch[1];

    // Get slide dimensions for proper scaling
    const slideDimensions = getSlideDimensions(pptSlide);

    // Get the opacity value
    const opacityStr = style.opacity || "1";
    const opacity = parseFloat(opacityStr);

    // Convert to PowerPoint transparency percentage (0-100)
    const transparencyPercentage = Math.round((1 - opacity) * 100);

    // Get background positioning and sizing
    const backgroundSize = style.backgroundSize || "cover";
    const backgroundPosition = style.backgroundPosition || "0px 0px";
    const backgroundRepeat = style.backgroundRepeat || "no-repeat";

    // Store background image data for XML generation
    const bgImageData = {
        type: 'background',
        src: src,
        x: 0,
        y: 0,
        w: slideDimensions.widthInches,
        h: slideDimensions.heightInches,
        transparency: transparencyPercentage,
        backgroundSize: backgroundSize,
        backgroundPosition: backgroundPosition,
        backgroundRepeat: backgroundRepeat,
        zIndex: -1
    };

    // Process the background as a full slide image
    try {
        let base64Data = null;

        // Handle different types of image paths
        if (src.includes("uploads")) {
            try {
                // Try to find the image file
                const imageName = path.basename(src);
                let imagePath = null;

                // Look in several possible locations
                const possiblePaths = [
                    src,
                    path.join("uploads", imageName),
                    path.join(__dirname, "..", "uploads", imageName),
                    path.join(__dirname, "..", src)
                ];

                for (const tryPath of possiblePaths) {
                    try {
                        if (fs.existsSync(tryPath)) {
                            imagePath = tryPath;
                            break;
                        }
                    } catch (e) {
                        // Ignore errors and try next path
                    }
                }

                if (!imagePath) {
                    throw new Error(`Background image file not found: ${src}`);
                }

                const imageData = fs.readFileSync(imagePath);
                const mimeType = getMimeType(imagePath);
                base64Data = `data:${mimeType};base64,${imageData.toString("base64")}`;
            } catch (localError) {
                console.error(`Error processing local background image (${src}):`, localError);
                throw localError;
            }
        } else if (src.startsWith("http")) {
            // Handle remote images
            try {
                const response = await axios.get(src, { responseType: "arraybuffer" });
                const mimeType = response.headers["content-type"] || "image/png";
                base64Data = `data:${mimeType};base64,${Buffer.from(response.data).toString("base64")}`;
            } catch (remoteError) {
                console.error(`Error fetching remote background image (${src}):`, remoteError);
                throw remoteError;
            }
        } else if (src.startsWith("data:image/")) {
            // Already a data URL
            base64Data = src;
        } else {
            throw new Error(`Unsupported background image source format: ${src}`);
        }

        if (!base64Data) {
            throw new Error("Failed to create background image data");
        }

        // Store the base64 data
        bgImageData.base64Data = base64Data;

        // Add the background image to slide with transparency
        pptSlide.addImage({
            data: base64Data,
            x: 0,
            y: 0,
            w: '100%',
            h: '100%',
            sizing: backgroundSize === "100% 100%" ? 'stretch' : 'cover',
            transparency: transparencyPercentage,
            zIndex: -1
        });

        if (!pptSlide.backgroundImages) pptSlide.backgroundImages = [];
        pptSlide.backgroundImages.push(bgImageData);

    } catch (error) {
        console.error(`Failed to add background image (${src}) to slide:`, error);
    }
}


// Function to generate image XML for PPTX
function generateImageXml(image, imageId, slideDimensions) {
    if (image.type === 'error') {
        // Generate error placeholder XML
        return generateErrorPlaceholderXml(image, imageId);
    }
    const x = Math.round((image.x || 0) * 914400); // Convert inches to EMU
    const y = Math.round((image.y || 0) * 914400);
    const w = Math.round((image.w || 1) * 914400);
    const h = Math.round((image.h || 1) * 914400);
    const rotation = image.rotation ? Math.round(image.rotation * 60000) : 0; // Convert degrees to 1/60000th
    const flipH = image.flipH ? '1' : '0';
    // Handle transparency
    const transparencyXml = image.transparency > 0 ?
        `<a:alpha val="${100 - image.transparency}000"/>` : '';

    // Handle srcRect / fillRect cropping
    // srcRectL, srcRectR, srcRectT, srcRectB can be positive (crop) or negative (zoom/fill)
    let fillRectXml = '<a:fillRect/>';

    const hasL = image.srcRectL && image.srcRectL !== '' && image.srcRectL !== '0';
    const hasR = image.srcRectR && image.srcRectR !== '' && image.srcRectR !== '0';
    const hasT = image.srcRectT && image.srcRectT !== '' && image.srcRectT !== '0';
    const hasB = image.srcRectB && image.srcRectB !== '' && image.srcRectB !== '0';

    if (hasL || hasR || hasT || hasB) {
        const attrs = [];
        if (hasL) attrs.push(`l="${image.srcRectL}"`);
        if (hasR) attrs.push(`r="${image.srcRectR}"`);
        if (hasT) attrs.push(`t="${image.srcRectT}"`);
        if (hasB) attrs.push(`b="${image.srcRectB}"`);

        fillRectXml = `<a:fillRect ${attrs.join(' ')}/>`;
    }

    // Handle border
    let borderXml = '';
    if (image.borderWidth > 0) {
        const borderWidthEmu = Math.round(image.borderWidth * 914400);
        const borderColor = (image.borderColor || '#000000').replace('#', '');
        borderXml = `
            <a:ln w="${borderWidthEmu}">
                <a:solidFill>
                    <a:srgbClr val="${borderColor}"/>
                </a:solidFill>
            </a:ln>`;
        // Handle shadow and glow effects
        let effectsXml = '';
        if (image.shadowEffects && (image.shadowEffects.hasGlow || image.shadowEffects.hasShadow)) {
            effectsXml = '<a:effectLst>';

            // Add glow effect
            if (image.shadowEffects.hasGlow && image.shadowEffects.glow) {
                const glow = image.shadowEffects.glow;
                const glowRadius = Math.round(glow.blur * 12700);
                const glowAlpha = Math.round(glow.alpha * 100000);
                const colorHex = ((glow.r << 16) | (glow.g << 8) | glow.b).toString(16).padStart(6, '0').toUpperCase();

                effectsXml += `
            <a:glow rad="${glowRadius}">
                <a:srgbClr val="${colorHex}">
                    <a:alpha val="${glowAlpha}"/>
                </a:srgbClr>
            </a:glow>`;
            }

            // Add outer shadow effect
            if (image.shadowEffects.hasShadow && image.shadowEffects.shadow) {
                const shadow = image.shadowEffects.shadow;
                const blurRad = Math.round(shadow.blur * 12700);
                const sx = Math.round(shadow.offsetX * 12700);
                const sy = Math.round(shadow.offsetY * 12700);
                const shadowAlpha = Math.round(shadow.alpha * 100000);
                const colorHex = ((shadow.r << 16) | (shadow.g << 8) | shadow.b).toString(16).padStart(6, '0').toUpperCase();

                effectsXml += `
            <a:outerShdw blurRad="${blurRad}" sx="${sx}" sy="${sy}" algn="ctr" rotWithShape="0">
                <a:srgbClr val="${colorHex}">
                    <a:alpha val="${shadowAlpha}"/>
                </a:srgbClr>
            </a:outerShdw>`;
            }

            effectsXml += '</a:effectLst>';
        }
    } else {
        borderXml = '<a:ln><a:noFill/></a:ln>';
    }
    const imageXml = `
        <p:pic>
            <p:nvPicPr>
                <p:cNvPr id="${imageId}" name="Image ${imageId}"/>
                <p:cNvPicPr/>
                <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
                <a:blip r:embed="rId${imageId}">
                    ${transparencyXml}
                </a:blip>
                <a:stretch>
                    ${fillRectXml}
                </a:stretch>
            </p:blipFill>
            <p:spPr>
                <a:xfrm${rotation !== 0 ? ` rot="${rotation}"` : ''}${flipH === '1' ? ` flipH="${flipH}"` : ''}>
                    <a:off x="${x}" y="${y}"/>
                    <a:ext cx="${w}" cy="${h}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
                <a:noFill/>
                ${borderXml}
                ${effectsXml}
            </p:spPr>
        </p:pic>`;
    return imageXml;
}

// Function to generate error placeholder XML
function generateErrorPlaceholderXml(errorData, shapeId) {
    const x = Math.round((errorData.x || 0) * 914400);
    const y = Math.round((errorData.y || 0) * 914400);
    const w = Math.round((errorData.w || 1) * 914400);
    const h = Math.round((errorData.h || 1) * 914400);
    const errorXml = `
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="${shapeId}" name="Error ${shapeId}"/>
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
                    <a:srgbClr val="EEEEEE"/>
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
                        <a:t>Image Error: ${escapeXml(errorData.error?.substring(0, 50) || 'Unknown error')}</a:t>
                    </a:r>
                </a:p>
            </p:txBody>
        </p:sp>`;
    return errorXml;
}

// Function to generate all images XML for a slide
function generateImagesXml(images) {
    let imagesXml = '';
    images.forEach((image, index) => {
        const imageId = index + 200; // Start image IDs from 200
        imagesXml += generateImageXml(image, imageId);
    });
    return imagesXml;
}

// Function to generate background images XML
function generateBackgroundImagesXml(backgroundImages) {
    let bgImagesXml = '';
    backgroundImages.forEach((bgImage, index) => {
        const imageId = index + 300; // Start background image IDs from 300
        // Background images are treated as regular images but with full slide dimensions
        bgImagesXml += generateImageXml(bgImage, imageId);
    });
    return bgImagesXml;
}

// Function to create slide XML with images
function createSlideXmlWithImages(slideData, slideDimensions) {
    const { width, height } = slideDimensions;
    const slideId = slideData.slideId || 1;
    const widthEmu = Math.round(width * 914400);
    const heightEmu = Math.round(height * 914400);
    const backgroundImagesXml = generateBackgroundImagesXml(slideData.backgroundImages || []);
    const imagesXml = generateImagesXml(slideData.images || []);
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
            ${backgroundImagesXml}
            ${imagesXml}
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>`;
    return slideXml;
}

// Function to generate image relationships for PPTX
function generateImageRelationships(images, backgroundImages) {
    let relationships = '';
    let rIdCounter = 1;
    // Add background image relationships
    backgroundImages.forEach((bgImage, index) => {
        const imageId = index + 300;
        relationships += `<Relationship Id="rId${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${imageId}.${getImageExtension(bgImage.src)}"/>`;
        rIdCounter++;
    });
    // Add regular image relationships
    images.forEach((image, index) => {
        if (image.type !== 'error') {
            const imageId = index + 200;
            relationships += `<Relationship Id="rId${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${imageId}.${getImageExtension(image.src)}"/>`;
            rIdCounter++;
        }
    });
    return relationships;
}

// Helper function to get image file extension
function getImageExtension(src) {
    if (src.includes('data:image/')) {
        const match = src.match(/data:image\/([^;]+)/);
        return match ? match[1] : 'png';
    }
    const ext = path.extname(src).toLowerCase().replace('.', '');
    return ext || 'png';
}

// Helper function to escape XML special characters
function escapeXml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function getMimeType(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    switch (ext) {
        case ".png":
            return "image/png";
        case ".jpg":
        case ".jpeg":
            return "image/jpeg";
        case ".gif":
            return "image/gif";
        case ".svg":
            return "image/svg+xml";
        case ".webp":
            return "image/webp";
        default:
            return "image/jpeg"; // Default fallback
    }
}
// Helper function to parse box-shadow into glow and shadow effects
function parseBoxShadowEffects(boxShadowString) {
    const shadows = boxShadowString.split(/,(?![^(]*\))/);

    let glowShadows = [];
    let offsetShadow = null;

    shadows.forEach(shadow => {
        const parts = shadow.trim().match(/([-\d.]+)px\s+([-\d.]+)px\s+([-\d.]+)px\s+rgba?\(([^)]+)\)/);
        if (parts) {
            const offsetX = parseFloat(parts[1]);
            const offsetY = parseFloat(parts[2]);
            const blur = parseFloat(parts[3]);
            const colorParts = parts[4].split(',').map(s => s.trim());

            if (offsetX === 0 && offsetY === 0) {
                // It's a glow
                glowShadows.push({
                    blur,
                    r: parseInt(colorParts[0]),
                    g: parseInt(colorParts[1]),
                    b: parseInt(colorParts[2]),
                    alpha: colorParts[3] ? parseFloat(colorParts[3]) : 1
                });
            } else {
                // It's an outer shadow
                offsetShadow = {
                    offsetX,
                    offsetY,
                    blur,
                    r: parseInt(colorParts[0]),
                    g: parseInt(colorParts[1]),
                    b: parseInt(colorParts[2]),
                    alpha: colorParts[3] ? parseFloat(colorParts[3]) : 1
                };
            }
        }
    });

    return {
        hasGlow: glowShadows.length > 0,
        glow: glowShadows.length > 0 ? glowShadows[0] : null,
        hasShadow: offsetShadow !== null,
        shadow: offsetShadow
    };
}

// Complete box-shadow to PPTX converter - all functions in one
function convertBoxShadowToPptxFormat(boxShadowString) {
    // Helper: Split by commas outside parentheses
    function splitOutsideParens(str) {
        const out = [];
        let cur = "";
        let depth = 0;
        for (let i = 0; i < str.length; i++) {
            const ch = str[i];
            if (ch === "(") depth++;
            if (ch === ")") depth = Math.max(0, depth - 1);

            if (ch === "," && depth === 0) {
                out.push(cur.trim());
                cur = "";
            } else {
                cur += ch;
            }
        }
        if (cur.trim()) out.push(cur.trim());
        return out;
    }

    // Helper: Tokenize shadow part by whitespace
    function tokenizeShadowPart(part) {
        const tokens = [];
        let cur = "";
        let depth = 0;
        for (let i = 0; i < part.length; i++) {
            const ch = part[i];
            if (ch === "(") depth++;
            if (ch === ")") depth = Math.max(0, depth - 1);

            const isSpace = /\s/.test(ch);
            if (isSpace && depth === 0) {
                if (cur.trim()) tokens.push(cur.trim());
                cur = "";
            } else {
                cur += ch;
            }
        }
        if (cur.trim()) tokens.push(cur.trim());
        return tokens;
    }

    // Helper: Parse color token
    function parseColorToken(tok) {
        const m = tok.match(/^rgba?\((.+)\)$/i);
        if (m) {
            const parts = m[1].split(",").map(s => s.trim());
            const r = parseFloat(parts[0]);
            const g = parseFloat(parts[1]);
            const b = parseFloat(parts[2]);
            const a = parts[3] !== undefined ? parseFloat(parts[3]) : 1;
            if ([r, g, b, a].some(x => Number.isNaN(x))) return null;

            const hex = ((Math.round(r) << 16) | (Math.round(g) << 8) | Math.round(b))
                .toString(16)
                .padStart(6, "0")
                .toUpperCase();

            return { hex, alpha: a };
        }

        const h = tok.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
        if (h) {
            let v = h[1];
            if (v.length === 3) v = v.split("").map(c => c + c).join("");
            return { hex: v.toUpperCase(), alpha: 1 };
        }

        return null;
    }

    // Helper: Parse length value
    function parseLength(tok) {
        const m = tok.match(/^([+-]?\d*\.?\d+)(px)?$/i);
        if (!m) return null;
        return parseFloat(m[1]);
    }

    // Helper: Parse one shadow from tokens
    function parseOneShadow(part) {
        const tokens = tokenizeShadowPart(part);

        let isInset = false;
        let color = null;

        const remaining = [];
        for (const t of tokens) {
            if (t.toLowerCase() === "inset") {
                isInset = true;
                continue;
            }
            const c = parseColorToken(t);
            if (c) {
                color = c;
                continue;
            }
            remaining.push(t);
        }

        const nums = remaining.map(parseLength).filter(v => v !== null);
        if (nums.length < 2) return null;

        const offsetX = nums[0];
        const offsetY = nums[1];
        const blur = nums.length >= 3 ? nums[2] : 0;
        const spread = nums.length >= 4 ? nums[3] : 0;

        const alpha = color ? color.alpha : 1;
        const hex = color ? color.hex : "000000";

        const offset = Math.sqrt(offsetX * offsetX + offsetY * offsetY);
        let angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (angle < 0) angle += 360;

        return {
            type: isInset ? "inner" : "outer",
            offsetX, offsetY, blur, spread,
            opacity: alpha,
            colorHex: hex,
            offset,
            angle
        };
    }

    // Main conversion logic
    if (!boxShadowString || boxShadowString === "none") return null;

    const parts = splitOutsideParens(boxShadowString);
    const parsed = parts.map(parseOneShadow).filter(Boolean);
    if (!parsed.length) return null;

    // Filter out completely transparent shadows (alpha = 0)
    const visibleShadows = parsed.filter(s => s.opacity > 0.001);

    if (!visibleShadows.length) return null; // All shadows are transparent

    // Check if this is a glow effect (multiple shadows with low offset)
    const lowOffsetShadows = visibleShadows.filter(s => s.offset < 5);
    const isLikelyGlow = visibleShadows.length > 1 && lowOffsetShadows.length >= visibleShadows.length * 0.7;

    if (isLikelyGlow) {
        // Combine multiple shadows for glow effect
        const totalBlur = visibleShadows.reduce((sum, s) => sum + s.blur, 0);
        const maxBlur = Math.max(...visibleShadows.map(s => s.blur));
        const avgBlur = totalBlur / visibleShadows.length;

        // Find the most opaque shadow for color
        const dominantShadow = visibleShadows.reduce((best, curr) =>
            curr.opacity > best.opacity ? curr : best
        );

        // Combine opacities
        const combinedOpacity = Math.min(1, visibleShadows.reduce((sum, s) => sum + s.opacity * 0.5, 0));

        // Use max blur + average blur for intensity
        const effectiveBlur = (maxBlur + avgBlur) / 2;

        return {
            type: "outer",
            opacity: String(Math.min(0.9, combinedOpacity)),
            blur: String(Math.round(effectiveBlur * 2.5)), // Boost blur for PowerPoint
            color: dominantShadow.colorHex,
            offset: String(0),
            angle: String(0)
        };
    }

    // Regular shadow: pick the most visible one
    function score(s) {
        const a = Math.max(0, Math.min(1, s.opacity));
        const offsetBoost = s.offset > 0.01 ? 1 : 0;
        const size = Math.max(0, s.blur) + Math.max(0, s.spread || 0);
        return (a * 1000) + (offsetBoost * 200) + (size * 1);
    }

    visibleShadows.sort((a, b) => score(b) - score(a));
    const best = visibleShadows[0];

    return {
        type: best.type,
        opacity: String(best.opacity),
        blur: String(Math.round(best.blur + (best.spread || 0) * 0.3)),
        color: best.colorHex,
        offset: String(Math.round(best.offset)),
        angle: String(Math.round(best.angle))
    };
}

module.exports = {
    addImageToSlide,
    addBackgroundImageToSlide,
    getSlideDimensions,
    generateImageXml,
    generateImagesXml,
    generateBackgroundImagesXml,
    createSlideXmlWithImages,
    generateImageRelationships,
    getMimeType,
    getImageExtension
};