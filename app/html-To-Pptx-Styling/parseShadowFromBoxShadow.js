function parseShadowFromBoxShadow(shapeElement) {
    try {
        // Read box-shadow from inline style attribute (most reliable in JSDOM/Node)
        const styleAttr = shapeElement.getAttribute('style') || '';

        // Extract box-shadow value from the style string
        // Handles multi-line styles with semicolons
        const shadowMatch = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
        if (!shadowMatch) return null;

        const boxShadowValue = shadowMatch[1].trim();
        if (!boxShadowValue || boxShadowValue === 'none') return null;

        const parsed = parseBoxShadowCSS(boxShadowValue);
        if (!parsed) return null;

        // Convert offsets → PPTX angle + distance
        const { angle, offset } = offsetToAngleAndDist(parsed.offsetX, parsed.offsetY);

        // Convert blur: CSS px → pt (at 72dpi they are equal)
        const blurPt = Math.round(parsed.blur * 10) / 10;

        // Convert spread → size
        const size = spreadToSize(parsed.spread, parsed.blur);

        // Clamp opacity
        const opacity = Math.max(0, Math.min(1, parsed.alpha));

        const shadowOptions = {
            type: 'outer',
            color: parsed.color,    // 6-char hex, no #
            opacity: opacity,
            blur: blurPt,
            offset: offset,
            angle: angle,
            size: size
        };

        // If all offsets are 0 and spread is very small, it's a centered glow
        // PPTX still handles it with type:'outer', angle:0, offset:0 — which is correct

        return shadowOptions;

    } catch (err) {
        console.error('parseShadowFromBoxShadow error:', err);
        return null;
    }
}

function spreadToSize(spread, blur) {
    if (!spread || spread <= 0) return 100;
    const base = blur > 0 ? blur : 10;
    const size = 100 + Math.round((spread / base) * 100);
    return Math.max(100, Math.min(size, 200));
}

function offsetToAngleAndDist(offsetX, offsetY) {
    const distPx = Math.sqrt(offsetX * offsetX + offsetY * offsetY);
    // 1 CSS px at 72 dpi = 1 pt
    const distPt = distPx;

    let angleDeg = 0;
    if (distPx > 0.01) {
        // atan2 gives angle where 0 = right, counterclockwise positive in math coords
        // but CSS Y increases downward, so atan2(y,x) gives clockwise angle from East
        angleDeg = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        // Normalize to 0–360
        if (angleDeg < 0) angleDeg += 360;
    }

    return {
        angle: Math.round(angleDeg),
        offset: Math.round(distPt * 10) / 10
    };
}

function parseBoxShadowCSS(cssValue) {
    if (!cssValue || cssValue === 'none' || cssValue.trim() === '') return null;

    // Take only the first shadow (ignore comma-separated multiples)
    // We need to be careful: rgba() contains commas, so we split at top-level commas only
    let firstShadow = cssValue.trim();
    {
        let depth = 0;
        let splitIdx = -1;
        for (let i = 0; i < firstShadow.length; i++) {
            if (firstShadow[i] === '(') depth++;
            else if (firstShadow[i] === ')') depth--;
            else if (firstShadow[i] === ',' && depth === 0) {
                splitIdx = i;
                break;
            }
        }
        if (splitIdx !== -1) firstShadow = firstShadow.slice(0, splitIdx).trim();
    }

    // Extract color token first (it can be at start or end)
    let colorRgba = null;
    let remaining = firstShadow;

    // Try to extract rgba/rgb color at START
    const rgbaAtStart = remaining.match(/^(rgba?\([^)]+\))\s*(.*)/i);
    if (rgbaAtStart) {
        colorRgba = parseRgba(rgbaAtStart[1]);
        remaining = rgbaAtStart[2];
    }

    // Try to extract rgba/rgb color at END
    if (!colorRgba) {
        const rgbaAtEnd = remaining.match(/(.*?)\s+(rgba?\([^)]+\))$/i);
        if (rgbaAtEnd) {
            colorRgba = parseRgba(rgbaAtEnd[2]);
            remaining = rgbaAtEnd[1];
        }
    }

    // Try hex color at start
    if (!colorRgba) {
        const hexAtStart = remaining.match(/^(#[0-9a-f]{3,8})\s+(.*)/i);
        if (hexAtStart) {
            colorRgba = parseHex(hexAtStart[1]);
            remaining = hexAtStart[2];
        }
    }

    // Try hex color at end
    if (!colorRgba) {
        const hexAtEnd = remaining.match(/(.*?)\s+(#[0-9a-f]{3,8})$/i);
        if (hexAtEnd) {
            colorRgba = parseHex(hexAtEnd[2]);
            remaining = hexAtEnd[1];
        }
    }

    // Default black if no color found
    if (!colorRgba) {
        colorRgba = { r: 0, g: 0, b: 0, a: 0.8 };
    }

    // Now parse the numeric length values: offsetX offsetY [blur] [spread]
    const numbers = remaining.trim().match(/([-\d.]+)px/g) || [];
    const vals = numbers.map(n => parseFloat(n));

    if (vals.length < 2) return null; // need at least offsetX + offsetY

    return {
        offsetX: vals[0] ?? 0,
        offsetY: vals[1] ?? 0,
        blur: vals[2] ?? 0,
        spread: vals[3] ?? 0,
        color: rgbToHex6(colorRgba.r, colorRgba.g, colorRgba.b),
        alpha: colorRgba.a ?? 1
    };
}

function rgbToHex6(r, g, b) {
    return [r, g, b]
        .map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
}

function parseHex(str) {
    const m = str.match(/^#([0-9a-f]{3,8})$/i);
    if (!m) return null;
    let hex = m[1];
    if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
    return {
        r: parseInt(hex.slice(0, 2), 16),
        g: parseInt(hex.slice(2, 4), 16),
        b: parseInt(hex.slice(4, 6), 16),
        a: hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1
    };
}

function parseRgba(str) {
    // rgba(r, g, b, a)
    const rgbaMatch = str.match(/rgba\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)/i);
    if (rgbaMatch) {
        return {
            r: parseInt(rgbaMatch[1], 10),
            g: parseInt(rgbaMatch[2], 10),
            b: parseInt(rgbaMatch[3], 10),
            a: parseFloat(rgbaMatch[4])
        };
    }
    // rgb(r, g, b)
    const rgbMatch = str.match(/rgb\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)/i);
    if (rgbMatch) {
        return {
            r: parseInt(rgbMatch[1], 10),
            g: parseInt(rgbMatch[2], 10),
            b: parseInt(rgbMatch[3], 10),
            a: 1
        };
    }
    return null;
}



module.exports = { parseShadowFromBoxShadow, parseBoxShadowCSS };

