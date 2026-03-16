// ─── parseShadowFromBoxShadow.js ─────────────────────────────────────────────
//
// Converts a CSS `box-shadow` string (from an HTML shape element) into a
// PptxGenJS `shadow` options object that can be passed directly to
// addShapeToSlide / addTextBoxToSlide as `{ shadow: <result> }`.
//
// Handles all 18 PowerPoint shadow-panel presets:
//   Outer × 9: Bottom Right / Bottom / Bottom Left / Right / Center /
//              Left / Top Right / Top / Top Left
//   Inner × 9: same nine positions  (CSS "inset" keyword → type:"inner")
//
// ─── Three bugs fixed vs the previous version ────────────────────────────────
//
//   FIX 1 — "inset" detection  (type was always "outer")
//     CSS "box-shadow: inset …" → { type: 'inner' }
//     CSS "box-shadow: …"       → { type: 'outer' }
//
//   FIX 2 — angle convention differs between outer and inner shadows
//
//     For OUTER shadows:
//       OOXML outerShdw `dir` = direction the shadow is CAST from the shape.
//       CSS offsetX/offsetY point in the SAME direction as the cast shadow.
//       Therefore: pptxgenjs_angle = atan2(offsetY, offsetX)  ← no rotation
//
//       Example – "Bottom" outer shadow: offY=+6 → atan2=90° → dir=5400000=90° ✓
//       Example – "Right"  outer shadow: offX=+5 → atan2=0°  → dir=0=0°        ✓
//
//     For INNER shadows:
//       getShapeShadowStyle.js already NEGATES the CSS offsets so that the
//       inset shadow renders on the correct inner edge (CSS inset convention is
//       opposite to OOXML). When we reverse those negated offsets back to the
//       original OOXML direction we must add +180°:
//       pptxgenjs_angle = (atan2(offsetY, offsetX) + 180) % 360
//
//       Example – "Left" inner shadow: CSS inset offX=+4 → atan2=0° → +180=180°
//         → dir=10800000=180° (West) ✓
//
//   FIX 3 — "size" field removed
//     PptxGenJS createShadowElement() hardcodes sx/sy to 100000 and has no
//     `size` option; the computed value was silently discarded.
//
// ─── Integration (addShapeToSlide.js) ────────────────────────────────────────
//
//   const { parseShadowFromBoxShadow } = require('./parseShadowFromBoxShadow');
//
//   // Before building shapeOptions:
//   const shadowOptions = parseShadowFromBoxShadow(shapeElement);
//
//   // Inside shapeOptions object literal:
//   ...(shadowOptions ? { shadow: shadowOptions } : {})
//
// ─────────────────────────────────────────────────────────────────────────────


/**
 * Parse a shape element's inline `box-shadow` style and return a PptxGenJS
 * shadow options object, or `null` if no shadow is present.
 *
 * @param {Element} shapeElement  DOM/JSDOM element with an inline `style` attr
 * @returns {{ type, color, opacity, blur, offset, angle } | null}
 */
function parseShadowFromBoxShadow(shapeElement) {
    try {
        const styleAttr = shapeElement.getAttribute('style') || '';

        // Extract box-shadow value (everything between "box-shadow:" and ";")
        const shadowMatch = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
        if (!shadowMatch) return null;

        const boxShadowValue = shadowMatch[1].trim();
        if (!boxShadowValue || boxShadowValue.toLowerCase() === 'none') return null;

        // ── FIX 1: detect "inset" BEFORE handing off to the parser ───────────
        // The CSS "inset" keyword can appear anywhere in the value string.
        const isInner = /\binset\b/i.test(boxShadowValue);

        const parsed = parseBoxShadowCSS(boxShadowValue);
        if (!parsed) return null;

        // ── FIX 2: correct angle conversion ──────────────────────────────────
        // For outer shadows: atan2 directly (shadow cast direction = CSS offset direction).
        // For inner shadows: atan2 + 180° (CSS inset offsets are pre-negated by
        // getShapeShadowStyle, so we rotate back to the original OOXML dir).
        const { angle, offset } = offsetToAngleAndDist(parsed.offsetX, parsed.offsetY, isInner);

        // blur: CSS px = pt at 72 dpi → pass directly
        const blurPt = Math.round(parsed.blur * 10) / 10;

        // opacity clamped to [0, 1]
        const opacity = Math.max(0, Math.min(1, parsed.alpha));

        // ── FIX 3: no "size" field — pptxgenjs ignores it anyway ─────────────
        return {
            type: isInner ? 'inner' : 'outer',
            color: parsed.color,   // 6-char uppercase hex, no '#'
            opacity: opacity,        // 0–1 float
            blur: blurPt,         // points (= CSS px at 72 dpi)
            offset: offset,         // points (displacement distance)
            angle: angle,          // degrees 0–360, CW from East (light-source dir)
        };

    } catch (err) {
        console.error('parseShadowFromBoxShadow error:', err);
        return null;
    }
}


// ─── offsetToAngleAndDist ─────────────────────────────────────────────────────
//
// Converts CSS box-shadow offsetX / offsetY into:
//   angle  – PptxGenJS angle (OOXML dir in degrees, CW from East)
//   offset – displacement distance in points
//
// Angle convention differs between outer and inner shadows:
//
//   OUTER (isInner=false):
//     OOXML outerShdw `dir` = direction the shadow is CAST (same as CSS offset).
//     pptxAngle = atan2(offsetY, offsetX)   ← no rotation needed
//
//     CSS offsets              atan2  pptxAngle  OOXML dir     preset
//     offX=+3.5 offY=+3.5      45°     45°       2700000    Bottom Right  ✓
//     offX=0    offY=+6        90°     90°       5400000    Bottom        ✓
//     offX=-2.8 offY=+2.8     135°    135°       8100000    Bottom Left   ✓
//     offX=+5   offY=0          0°      0°            0    Right         ✓
//     offX=0    offY=0        (dist=0)  0°            0    Center        ✓
//     offX=-4   offY=0        180°    180°      10800000    Left          ✓
//     offX=+2.8 offY=-2.8     315°    315°      18900000    Top Right     ✓
//     offX=0    offY=-6       270°    270°      16200000    Top           ✓
//     offX=-4.2 offY=-4.2     225°    225°      13500000    Top Left      ✓
//
//   INNER (isInner=true):
//     getShapeShadowStyle.js pre-negates CSS inset offsets so the shadow renders
//     on the correct inner edge. Reversing that negation requires +180°:
//     pptxAngle = (atan2(offsetY, offsetX) + 180) % 360
//
//     CSS offsets (negated)    atan2  pptxAngle  OOXML dir     preset
//     offX=+2.8 offY=+2.8      45°    225°      13500000    Inner Top Left  ✓
//     offX=0    offY=+4        90°    270°      16200000    Inner Top       ✓
//     offX=-3.5 offY=+3.5     135°    315°      18900000    Inner Top Right ✓
//     offX=+4   offY=0          0°    180°      10800000    Inner Left      ✓
//     offX=0    offY=0        (dist=0) 0°            0     Inner Center    ✓
//     offX=-5   offY=0        180°      0°            0    Inner Right     ✓
//     offX=+4.2 offY=-4.2     315°    135°       8100000    Inner SW        ✓
//     offX=0    offY=-4       270°     90°       5400000    Inner Bottom    ✓
//     offX=-2.8 offY=-2.8     225°     45°       2700000    Inner SE        ✓
//
// @param {number}  offsetX  – CSS box-shadow offsetX in px
// @param {number}  offsetY  – CSS box-shadow offsetY in px
// @param {boolean} isInner  – true for inset (inner) shadows
//
function offsetToAngleAndDist(offsetX, offsetY, isInner) {
    const distPx = Math.sqrt(offsetX * offsetX + offsetY * offsetY);
    const distPt = Math.round(distPx * 10) / 10;

    let pptxAngle = 0;
    if (distPx > 0.01) {
        // Direction the shadow offset points (CW from East, Y-axis down)
        let atan2Deg = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (atan2Deg < 0) atan2Deg += 360;

        // Outer: CSS offset direction == OOXML dir → use atan2 directly.
        // Inner: CSS inset offsets are pre-negated vs OOXML dir → rotate +180°
        //        to recover the original OOXML shadow-cast direction.
        pptxAngle = isInner ? (atan2Deg + 180) % 360 : atan2Deg;
    }

    return {
        angle: Math.round(pptxAngle),
        offset: distPt,
    };
}


// ─── parseBoxShadowCSS ────────────────────────────────────────────────────────
//
// Parses a single CSS box-shadow string into its numeric components.
// The "inset" keyword is stripped before number extraction so it does not
// interfere with the numeric parsing (no parseFloat side-effects).
//
// Returns: { offsetX, offsetY, blur, spread, color, alpha }  or  null
//
function parseBoxShadowCSS(cssValue) {
    if (!cssValue || cssValue === 'none' || cssValue.trim() === '') return null;

    // ── Take only the first shadow (multi-shadow: split at top-level commas) ─
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

    // ── Strip "inset" keyword before numeric extraction ───────────────────────
    let remaining = firstShadow.replace(/\binset\b/gi, '').trim();

    // ── Extract color token (rgba / rgb / hex, at start or end) ──────────────
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

    // Fallback: black at 80% (PowerPoint's default shadow color/opacity)
    if (!colorRgba) colorRgba = { r: 0, g: 0, b: 0, a: 0.8 };

    // ── Parse numeric lengths: offsetX offsetY [blur] [spread] ───────────────
    const numbers = remaining.trim().match(/([-\d.]+)px/g) || [];
    const vals = numbers.map(n => parseFloat(n));

    if (vals.length < 2) return null; // need at least offsetX + offsetY

    return {
        offsetX: vals[0] ?? 0,
        offsetY: vals[1] ?? 0,
        blur: vals[2] ?? 0,
        spread: vals[3] ?? 0,
        color: rgbToHex6(colorRgba.r, colorRgba.g, colorRgba.b),
        alpha: colorRgba.a ?? 1,
    };
}


// ─── Color helpers ────────────────────────────────────────────────────────────

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


module.exports = { parseShadowFromBoxShadow, parseBoxShadowCSS };