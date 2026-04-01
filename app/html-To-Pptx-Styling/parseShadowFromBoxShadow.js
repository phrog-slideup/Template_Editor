// parseShadowFromBoxShadow.js
//

const ONEPT = 12700; // 1 pt = 12700 EMU (same constant pptxgenjs uses internally)


function buildShadowXml(shapeElement) {
    try {
        const styleAttr = shapeElement.getAttribute('style') || '';
        const shadowMatch = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
        if (!shadowMatch) return null;

        const boxShadowValue = shadowMatch[1].trim();
        if (!boxShadowValue || boxShadowValue.toLowerCase() === 'none') return null;

        const isInner = /\binset\b/i.test(boxShadowValue);
        const parsed = parseBoxShadowCSS(boxShadowValue);

        if (!parsed) return null;

        // ── Spread simulation ────────────────────────────────────────────────────
        // PowerPoint's outerShdw has no "spread" concept.  The old approach of
        // mapping spread → sx/sy scale produced sx=0/sy=0 for negative spread
        // (e.g. blur=4 spread=-4 → scale=0), which collapses the shadow to a
        // point and then re-expands it symmetrically via blur — exactly the
        // "top/bottom/right bleed" bug seen in the images.
        //
        // Correct approximation for NEGATIVE spread (the common "one-sided
        // shadow" pattern):
        //   • Reduce blurRad by |spread|  → tighten the feather so it doesn't
        //     bleed past the shape edges on the non-offset sides.
        //   • Increase dist by |spread|   → keep the visible shadow edge in
        //     roughly the same visual position.
        //
        // For POSITIVE spread we keep the old scale approach (sx/sy > 100 000)
        // because that faithfully enlarges the shadow footprint.
        // ────────────────────────────────────────────────────────────────────────

        let effectiveBlur = parsed.blur;
        let effectiveDist = Math.sqrt(
            parsed.offsetX * parsed.offsetX + parsed.offsetY * parsed.offsetY
        );

        if (!isInner && parsed.spread < -0.0001) {
            // Negative spread: shrink blur, push dist outward to compensate
            const shrink = Math.min(Math.abs(parsed.spread), effectiveBlur); // never go negative
            effectiveBlur = Math.max(0, effectiveBlur - shrink);
            effectiveDist = effectiveDist + shrink;
        }

        // Convert CSS px offsets back to OOXML dist (EMU) + dir (60000ths of a degree).
        // When we have adjusted the effective distance we rebuild distEmu from it
        // directly instead of re-deriving from raw offsets.
        const { dirEmu } = offsetToOoxmlDirDist(parsed.offsetX, parsed.offsetY, isInner);
        const distEmu = Math.round(effectiveDist * ONEPT);

        // CSS px = pt at 72 dpi, 1pt = ONEPT EMU
        const blurEmu = Math.round(effectiveBlur * ONEPT);

        let sxVal, syVal;
        if (isInner) {
            // innerShdw has no sx/sy attributes in OOXML — leave undefined
            sxVal = undefined;
            syVal = undefined;
        } else if (!isInner && parsed.spread > 0.0001) {
            // Positive spread only: enlarge the shadow via scale factor
            const referenceSize = parsed.blur > 0.0001 ? parsed.blur : 1;
            const avgScale = 1 + (parsed.spread / referenceSize);
            sxVal = Math.round(Math.max(0, avgScale) * 100000);
            syVal = sxVal; // PowerPoint always writes sx === sy (symmetric)
        } else {
            // spread=0 or negative (already handled via blur/dist above) → no scaling
            sxVal = 100000;
            syVal = 100000;
        }

        // Alpha: CSS rgba 0-1 float -> OOXML 0-100000 (no intermediate rounding)
        const alphaVal = Math.round(Math.max(0, Math.min(1, parsed.alpha)) * 100000);

        const color = parsed.color; // 6-char uppercase hex, no '#'
        const tag = isInner ? 'innerShdw' : 'outerShdw';

        // outerShdw requires sx/sy/kx/ky/algn attrs.
        // innerShdw must NOT have them (sxVal/syVal are undefined for inner).
        const outerAttrs = isInner
            ? ''
            : ` sx="${sxVal}" sy="${syVal}" kx="0" ky="0" algn="bl" rotWithShape="0"`;

        const alphaTag = alphaVal < 100000 ? `<a:alpha val="${alphaVal}"/>` : '';

        let xml = '<a:effectLst>';
        xml += `<a:${tag}${outerAttrs} blurRad="${blurEmu}" dist="${distEmu}" dir="${dirEmu}">`;
        xml += `<a:srgbClr val="${color}">${alphaTag}</a:srgbClr>`;
        xml += `</a:${tag}>`;  // FIX 4: correct closing tag, NOT hardcoded "outerShdw"
        xml += '</a:effectLst>';

        return xml;

    } catch (err) {
        console.error('buildShadowXml error:', err);
        return null;
    }
}


// ── parseShadowFromBoxShadow ──────────────────────────────────────────────────
//

function parseShadowFromBoxShadow(shapeElement) {
    try {
        const styleAttr = shapeElement.getAttribute('style') || '';

        const shadowMatch = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
        if (!shadowMatch) return null;

        const boxShadowValue = shadowMatch[1].trim();
        if (!boxShadowValue || boxShadowValue.toLowerCase() === 'none') return null;

        // FIX 1: detect "inset" before any other processing
        const isInner = /\binset\b/i.test(boxShadowValue);

        const parsed = parseBoxShadowCSS(boxShadowValue);
        if (!parsed) return null;

        // FIX 2: correct angle/distance conversion for inner vs outer
        const { angle, offset: rawOffset } = offsetToAngleAndDist(parsed.offsetX, parsed.offsetY, isInner);

        // ── Spread compensation (negative spread = "one-sided shadow") ──────────
        // PowerPoint has no spread concept.  A negative CSS spread (e.g. -4px)
        // combined with a matching blur (4px) creates a directional-only shadow
        // by clipping the feather on non-offset sides.  Without compensation the
        // shadow bleeds to top/bottom/right in PowerPoint.
        //
        // Fix: reduce blur by |spread| (tighten the feather) and increase the
        // displacement distance by the same amount (keep the visible edge in place).
        // ────────────────────────────────────────────────────────────────────────
        let effectiveBlurPt = parsed.blur;
        let effectiveOffsetPt = rawOffset;

        if (!isInner && parsed.spread < -0.0001) {
            const shrinkPx = Math.min(Math.abs(parsed.spread), parsed.blur);
            effectiveBlurPt = Math.max(0, parsed.blur - shrinkPx);
            effectiveOffsetPt = rawOffset + shrinkPx;
        }

        // blur: CSS px = pt at 72 dpi, pass directly to pptxgenjs
        const blurPt = Math.round(effectiveBlurPt * 10) / 10;
        const offset = Math.round(effectiveOffsetPt * 10) / 10;

        // FIX 3 / opacity: keep full 0-1 float, no lossy rounding
        const opacity = Math.max(0, Math.min(1, parsed.alpha));

        return {
            type: isInner ? 'inner' : 'outer',
            color: parsed.color,  // 6-char uppercase hex, no '#'
            opacity: opacity,       // 0-1 float (pptxgenjs multiplies by 100000 internally)
            blur: blurPt,        // points (pptxgenjs converts to EMU internally)
            offset: offset,        // points (displacement distance)
            angle: angle,         // degrees 0-360, CW from East
        };

    } catch (err) {
        console.error('parseShadowFromBoxShadow error:', err);
        return null;
    }
}


// ── offsetToOoxmlDirDist ──────────────────────────────────────────────────────

function offsetToOoxmlDirDist(offsetX, offsetY, isInner) {
    // Undo the negation that getShapeShadowStyle applies for inner shadows
    const ox = isInner ? -offsetX : offsetX;
    const oy = isInner ? -offsetY : offsetY;

    const distPx = Math.sqrt(ox * ox + oy * oy);
    const distEmu = Math.round(distPx * ONEPT);

    let dirEmu = 0;
    if (distPx > 0.01) {
        let atan2Deg = Math.atan2(oy, ox) * (180 / Math.PI);
        if (atan2Deg < 0) atan2Deg += 360;
        dirEmu = Math.round(atan2Deg * 60000);
    }

    return { dirEmu, distEmu };
}


// ── offsetToAngleAndDist ──────────────────────────────────────────────────────
//

function offsetToAngleAndDist(offsetX, offsetY, isInner) {
    const distPx = Math.sqrt(offsetX * offsetX + offsetY * offsetY);
    const distPt = Math.round(distPx * 10) / 10;

    let pptxAngle = 0;
    if (distPx > 0.01) {
        let atan2Deg = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (atan2Deg < 0) atan2Deg += 360;
        // Outer: CSS dir == OOXML dir -> atan2 directly
        // Inner: CSS offsets are pre-negated -> rotate +180 deg to recover OOXML dir
        pptxAngle = isInner ? (atan2Deg + 180) % 360 : atan2Deg;
    }

    return {
        angle: Math.round(pptxAngle),
        offset: distPt,
    };
}


// ── parseBoxShadowCSS ─────────────────────────────────────────────────────────
//

function parseBoxShadowCSS(cssValue) {
    if (!cssValue || cssValue === 'none' || cssValue.trim() === '') return null;

    // Take only the first shadow (multi-shadow: split at top-level commas)
    let firstShadow = cssValue.trim();
    {
        let depth = 0, splitIdx = -1;
        for (let i = 0; i < firstShadow.length; i++) {
            if (firstShadow[i] === '(') depth++;
            else if (firstShadow[i] === ')') depth--;
            else if (firstShadow[i] === ',' && depth === 0) { splitIdx = i; break; }
        }
        if (splitIdx !== -1) firstShadow = firstShadow.slice(0, splitIdx).trim();
    }

    // Strip "inset" keyword before numeric extraction
    let remaining = firstShadow.replace(/\binset\b/gi, '').trim();

    // Extract color token (rgba / rgb / hex — at start or end of the value)
    let colorRgba = null;

    const rgbaAtStart = remaining.match(/^(rgba?\([^)]+\))\s*(.*)/i);
    if (rgbaAtStart) { colorRgba = parseRgba(rgbaAtStart[1]); remaining = rgbaAtStart[2]; }

    if (!colorRgba) {
        const rgbaAtEnd = remaining.match(/(.*?)\s+(rgba?\([^)]+\))$/i);
        if (rgbaAtEnd) { colorRgba = parseRgba(rgbaAtEnd[2]); remaining = rgbaAtEnd[1]; }
    }

    if (!colorRgba) {
        const hexAtStart = remaining.match(/^(#[0-9a-f]{3,8})\s+(.*)/i);
        if (hexAtStart) { colorRgba = parseHex(hexAtStart[1]); remaining = hexAtStart[2]; }
    }

    if (!colorRgba) {
        const hexAtEnd = remaining.match(/(.*?)\s+(#[0-9a-f]{3,8})$/i);
        if (hexAtEnd) { colorRgba = parseHex(hexAtEnd[2]); remaining = hexAtEnd[1]; }
    }

    // Fallback: black at 80% (PowerPoint default shadow)
    if (!colorRgba) colorRgba = { r: 0, g: 0, b: 0, a: 0.8 };

    // Parse numeric lengths: offsetX offsetY [blur] [spread]
    const numbers = remaining.trim().match(/([-\d.]+)px/g) || [];
    const vals = numbers.map(n => parseFloat(n));

    if (vals.length < 2) return null;

    return {
        offsetX: vals[0] ?? 0,
        offsetY: vals[1] ?? 0,
        blur: vals[2] ?? 0,
        spread: vals[3] ?? 0,
        color: rgbToHex6(colorRgba.r, colorRgba.g, colorRgba.b),
        alpha: colorRgba.a ?? 1,
    };
}


// ── Color helpers ─────────────────────────────────────────────────────────────

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
        a: hex.length === 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1,
    };
}

function parseRgba(str) {
    const rgba = str.match(/rgba\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)/i);
    if (rgba) return { r: +rgba[1], g: +rgba[2], b: +rgba[3], a: +rgba[4] };

    const rgb = str.match(/rgb\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)/i);
    if (rgb) return { r: +rgb[1], g: +rgb[2], b: +rgb[3], a: 1 };

    return null;
}


module.exports = { buildShadowXml, parseShadowFromBoxShadow, parseBoxShadowCSS };