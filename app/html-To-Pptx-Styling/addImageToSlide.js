const path = require("path");
const fs = require("fs");
const axios = require("axios");
const normalizeStyle = require("../api/helper/colorHelper");
const clrHelper = require("../api/helper/colorHelper.js");
const imageSavePath = path.join(process.cwd(), "uploads");

function ensureImageSavePath() {
    if (!fs.existsSync(imageSavePath)) {
        fs.mkdirSync(imageSavePath, { recursive: true });
    }
}

function inferScene3dCameraFromTransform(transformStr = "") {
    const normalized = String(transformStr || "").replace(/\s+/g, " ").trim();
    if (!normalized) return "";

    if (normalized.includes("rotateX(58deg)") && normalized.includes("rotateZ(-45deg)")) {
        return "isometricTopUp";
    }

    if (normalized.includes("rotateX(18deg)")) {
        return "perspectiveRelaxedModerately";
    }

    return "";
}

async function addImageToSlide(pptx, pptSlide, imgElement, slideContext) {

    const isSvgElement = imgElement?.tagName?.toLowerCase() === "svg";
    let src = isSvgElement ? "" : imgElement.getAttribute("src");

    const style = imgElement.style;
    const parent = imgElement.closest(".image-container");
    const objName = parent.getAttribute("data-name");

    const altText = parent?.getAttribute("data-alt-text") || '';

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

        if (boxShadowValue && boxShadowValue !== 'none') {
            pptxShadow = convertBoxShadowToPptxFormat(boxShadowValue);
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
    // ✅ ADD THIS - detect grayscale CSS filter
    const isGrayscale = imgStyleAttr.includes('grayscale(100%)') || imgStyleAttr.includes('grayscale(1)');
    if (imgOpacityMatch && imgOpacityMatch[1]) {
        imgOpacity = parseFloat(imgOpacityMatch[1]);
    }

    const scene3dCamera = isSvgElement
        ? inferScene3dCameraFromTransform(style.transform || "")
        : "";

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

    }
    // Fallback to img element
    if (!srcRectL && !srcRectR && !srcRectT && !srcRectB) {
        srcRectL = imgElement.getAttribute('srcrectl') || imgElement.getAttribute('srcRectL') || '';
        srcRectR = imgElement.getAttribute('srcrectr') || imgElement.getAttribute('srcRectR') || '';
        srcRectT = imgElement.getAttribute('srcrectt') || imgElement.getAttribute('srcRectT') || '';
        srcRectB = imgElement.getAttribute('srcrectb') || imgElement.getAttribute('srcRectB') || '';

    }

    // Parse srcRect values (can be negative!)
    const parsedL = parseInt(srcRectL) || 0;
    const parsedR = parseInt(srcRectR) || 0;
    const parsedT = parseInt(srcRectT) || 0;
    const parsedB = parseInt(srcRectB) || 0;

    const hasCrop = parsedL !== 0 || parsedR !== 0 || parsedT !== 0 || parsedB !== 0;


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

        let imagePathForPpt = null;

        if (isSvgElement) {
            ensureImageSavePath();
            const innerSvgElement = imgElement.querySelector("svg");
            let svgMarkup = (innerSvgElement || imgElement).outerHTML;

            // Drop HTML-only accessibility/style wrappers that can produce
            // invalid nested SVG payloads for PowerPoint embedding.
            svgMarkup = svgMarkup
                .replace(/\srole="[^"]*"/gi, "")
                .replace(/\saria-label="[^"]*"/gi, "");

            const svgFileName = `html_export_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.svg`;
            imagePathForPpt = path.join(imageSavePath, svgFileName);
            fs.writeFileSync(imagePathForPpt, svgMarkup, "utf8");

        } else if (src.includes("uploads")) {
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

        if (!base64Data && !imagePathForPpt) {
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
            altText: altText,
            _isGrayscale: isGrayscale,
        };

        if (imagePathForPpt) {
            imageOptions.path = imagePathForPpt;
        } else {
            imageOptions.data = base64Data;
        }

        // ========================================
        // ✅ CRITICAL: Add sizing with crop
        // pptxgenjs requires sizing.type = 'crop' with x, y, w, h in EMU units (0-100000)
        // ========================================
        if (hasCrop) {
            imageOptions._srcRect = {
                l: parsedL,
                r: parsedR,
                t: parsedT,
                b: parsedB
            };
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
        }
        if (pptxShadow && pptxShadow.spread) {
            // Store sx/sy for XML post-processing
            // spread in px → approximate scale percentage (e.g., 9px spread ≈ 103%)
            const scalePct = 100 + Math.round(pptxShadow.spread / 3);
            imageOptions._shadowScale = {
                sx: scalePct * 1000,  // e.g., 103000
                sy: scalePct * 1000
            };
        }

        // Add hyperlink if exists
        if (hyperlink) {
            imageOptions.hyperlink = { url: hyperlink };
        }

        if (scene3dCamera && objName) {
            if (!Array.isArray(pptx._svgScene3dStore)) {
                pptx._svgScene3dStore = [];
            }

            pptx._svgScene3dStore.push({
                slideIndex: slideContext?.slideIndex ?? 0,
                objectName: objName,
                cameraPrst: scene3dCamera
            });
        }

        // Add image to slide
        pptSlide.addImage(imageOptions);

        if (isGrayscale) {
            if (!global.grayscaleImageStore) global.grayscaleImageStore = new Map();
            const storeKey = objName || `img_${x.toFixed(3)}_${y.toFixed(3)}`;
            global.grayscaleImageStore.set(storeKey, {
                objectName: objName,
                x, y, w, h
            });
            // console.log(`🔲 Queued grayscale image: "${objName}"`);
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

    // Spread in CSS creates thickness — in PPT this maps to higher blur + offset
    const spreadBoost = best.spread || 0;
    const effectiveBlur = best.blur + spreadBoost * 0.8;
    const effectiveOffset = best.offset + spreadBoost * 0.5;

    return {
        type: best.type,
        opacity: String(Math.min(1, best.opacity * 1.5)),
        blur: String(Math.round(effectiveBlur)),
        color: best.colorHex,
        offset: String(Math.round(effectiveOffset)),
        angle: String(Math.round(best.angle))
    };
}

module.exports = {
    addImageToSlide,
    getSlideDimensions,
    getMimeType,
};