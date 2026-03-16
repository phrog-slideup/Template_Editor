// ─── getShapeReflection.js ────────────────────────────────────────────────────
//
// Converts an OOXML <a:reflection> element (inside <a:effectLst>) into the
// HTML/CSS needed to render a PowerPoint-style shape reflection.
//
// ─── How PowerPoint reflections work ─────────────────────────────────────────
//
//   A reflection is a vertically-flipped, partially-transparent copy of the
//   shape placed at a configurable distance (gap) below the original shape.
//   The reflected copy fades from startAlpha → endAlpha across a gradient
//   that runs from stPos% to endPos% of the shape height, and the whole copy
//   can have a blur applied.
//
//   OOXML <a:reflection> attributes used here:
//     blurRad   – blur radius (EMU).  Defaults to 0.
//     dist      – gap between shape bottom and reflection top (EMU).
//     dir       – direction in 60000ths of a degree.
//                 5400000 = 90° = downward (the only common value).
//     stA       – start alpha at stPos   (0-100000, where 100000 = fully opaque).
//     stPos     – gradient start position (0-100000 = 0–100%). Defaults to 0.
//     endA      – end alpha at endPos    (0-100000). Defaults to 0 (transparent).
//     endPos    – gradient end position  (0-100000 = 0–100%). Defaults to 100000.
//     sx        – horizontal scale (100000 = 100%).  Negative = flip horizontal.
//     sy        – vertical scale   (-100000 = -100% = flip upside-down).
//     kx / ky   – skew angles in 60000ths of a degree (rarely used, supported).
//     algn      – alignment anchor for the reflection transform.
//     rotWithShape – "1" if the reflection should rotate together with the shape.
//
// ─── CSS implementation strategy ─────────────────────────────────────────────
//
//   Because the reflection sits OUTSIDE the shape's own bounding box, we
//   cannot implement it with a single div.  Instead:
//
//   1. getShapeReflectionData()   – parses the XML node, returns a plain
//                                   data object (or null if no reflection).
//
//   2. buildReflectionHTML()      – takes the data object + shape geometry
//                                   and returns an HTML string for the
//                                   reflection <div>.
//
//   3. wrapShapeWithReflection()  – wraps the original shape HTML + the
//                                   reflection HTML in a common container
//                                   <div> so both are positioned relative
//                                   to the same origin.
//
//   The caller (shapeHandler.js) should:
//     a) Call getShapeReflectionData(shapeNode) after extracting position.
//     b) If data is non-null, call wrapShapeWithReflection(shapeHTML, data, position).
//     c) Return the wrapped string instead of the plain shapeHTML string.
//
// ─── Integration example (shapeHandler.js) ───────────────────────────────────
//
//   const { getShapeReflectionData, wrapShapeWithReflection } =
//       require('./shapes_Properties/getShapeReflection.js');
//
//   // … (existing position / fill / shadow extraction) …
//
//   // Build the normal shape HTML as usual (existing code, no changes needed):
//   let shapeHTML = `<div class="shape" …> … </div>`;
//
//   // Apply reflection if present:
//   const reflectionData = getShapeReflectionData(shapeNode);
//   if (reflectionData) {
//       shapeHTML = wrapShapeWithReflection(shapeHTML, reflectionData, position, fillColor, borderRadius);
//   }
//
//   return shapeHTML;
//
// ─────────────────────────────────────────────────────────────────────────────

"use strict";

// ─── Constants ────────────────────────────────────────────────────────────────

const EMU_PER_PX = 12700;   // 1px = 12700 EMU at 72 DPI (matches getEMUDivisor())

// ─── getShapeReflectionData ───────────────────────────────────────────────────
//
// Parses the <a:reflection> node from a parsed xml2js shapeNode object.
//
// @param  {object} shapeNode  – xml2js-parsed shape node (p:sp)
// @returns {{ 
//   blurPx, distPx, dirDeg,
//   startAlpha, startPosPct, endAlpha, endPosPct,
//   scaleX, scaleY,
//   skewXDeg, skewYDeg,
//   algn, rotWithShape
// } | null}
//
function getShapeReflectionData(shapeNode) {
    try {
        const reflNode = shapeNode
            ?.["p:spPr"]?.[0]
            ?.["a:effectLst"]?.[0]
            ?.["a:reflection"]?.[0];

        if (!reflNode) return null;

        const a = reflNode["$"] || {};

        // ── Core geometry ─────────────────────────────────────────────────────
        const blurPx = a.blurRad ? parseInt(a.blurRad, 10) / EMU_PER_PX : 0;
        const distPx = a.dist ? parseInt(a.dist, 10) / EMU_PER_PX : 0;
        const dirDeg = a.dir ? parseInt(a.dir, 10) / 60000 : 90;  // default: downward

        // ── Alpha / fade gradient ─────────────────────────────────────────────
        // stA / endA are in units of 1/1000 percent (100000 = 100% = fully opaque)
        const startAlpha = a.stA ? parseInt(a.stA, 10) / 100000 : 1;
        const endAlpha = a.endA ? parseInt(a.endA, 10) / 100000 : 0;
        const startPosPct = a.stPos ? parseInt(a.stPos, 10) / 1000 : 0;    // 0–100
        const endPosPct = a.endPos ? parseInt(a.endPos, 10) / 1000 : 100;  // 0–100

        // ── Scale (sx/sy) – negative value means flip that axis ───────────────
        // 100000 = 100 % = no scaling.  -100000 = mirror.
        const scaleX = a.sx ? parseInt(a.sx, 10) / 100000 : 1;
        const scaleY = a.sy ? parseInt(a.sy, 10) / 100000 : 1;

        // ── Skew (kx/ky) ──────────────────────────────────────────────────────
        const skewXDeg = a.kx ? parseInt(a.kx, 10) / 60000 : 0;
        const skewYDeg = a.ky ? parseInt(a.ky, 10) / 60000 : 0;

        // ── Misc ──────────────────────────────────────────────────────────────
        const algn = a.algn || "bl";  // bottom-left is the PPT default
        const rotWithShape = a.rotWithShape === "1";

        return {
            blurPx: Math.round(blurPx * 10) / 10,
            distPx: Math.round(distPx * 10) / 10,
            dirDeg,
            startAlpha,
            startPosPct: Math.round(startPosPct * 100) / 100,
            endAlpha,
            endPosPct: Math.round(endPosPct * 100) / 100,
            scaleX,
            scaleY,
            skewXDeg,
            skewYDeg,
            algn,
            rotWithShape,
        };

    } catch (err) {
        console.error("getShapeReflectionData error:", err);
        return null;
    }
}


// ─── buildReflectionHTML ──────────────────────────────────────────────────────
//
// Builds the <div> that renders the PowerPoint-style reflection below the shape.
//
// ── Why a SINGLE div is correct ──────────────────────────────────────────────
//
// CSS applies visual effects on a single element in this strict pipeline order:
//
//   Paint background → transform → filter:blur → mask-image
//
// This means on ONE element:
//   1. Background is painted (solid fill)
//   2. scaleY(-1) flips it → bottom of shape now at top of reflection
//   3. filter:blur() blurs the flipped result
//   4. mask-image fades the blurred result top→bottom
//
// This is exactly the correct order to match PowerPoint's rendering.
//
// ── Why the previous two-div approach FAILED ─────────────────────────────────
//
// The two-div structure placed mask-image on the OUTER div and filter:blur on
// the INNER div. This triggers a Chromium browser bug:
//
//   "parent mask-image does NOT correctly apply to children that have
//    filter:blur() because filter creates a new GPU compositing layer.
//    The child's composited layer bypasses the parent's mask."
//
// Result: the blurred inner div renders at full opacity, ignoring the mask.
// The reflection appears either nearly invisible (masked incorrectly) or
// fully visible with no fade — neither matching PowerPoint.
//
// ── Why no overflow:hidden ────────────────────────────────────────────────────
//
// Without overflow:hidden, filter:blur() spreads pixels slightly OUTSIDE the
// div box (by blurPx pixels). This is intentional — it creates a soft glow at
// the gap boundary between the shape and its reflection, which is exactly what
// PowerPoint renders. overflow:hidden would clip this soft edge.
//
// @param {object} data           – output of getShapeReflectionData()
// @param {object} position       – { x, y, width, height, rotation }
// @param {string} fillColor      – CSS colour / gradient string (background)
// @param {string} borderRadius   – CSS border-radius e.g. "0px" or "50%"
// @param {string} extraShapeCSS  – additional shape CSS (border, etc.)
// @param {number} zIndex         – z-index of the parent shape
//
// @returns {string}  HTML string for the reflection <div>
//
function buildReflectionHTML(data, position, fillColor, borderRadius, extraShapeCSS, zIndex) {
    const {
        blurPx,
        distPx,
        dirDeg,
        startAlpha,
        startPosPct,
        endAlpha,
        endPosPct,
        scaleX,
        scaleY,
        skewXDeg,
        skewYDeg,
    } = data;

    const { width, height } = position;

    // ── 1. Position ───────────────────────────────────────────────────────────
    // dir=90° (OOXML 5400000 / 60000) = reflection appears directly below shape.
    // Place the reflection div 'distPx' below the shape's bottom edge.
    // Inside the wrapper, the shape sits at top:0, so reflection top = height + dist.

    const dirRad = dirDeg * (Math.PI / 180);
    const reflLeft = Math.round(Math.cos(dirRad) * distPx * 10) / 10;
    const reflTop = height + Math.round(Math.sin(dirRad) * distPx * 10) / 10;

    // ── 2. Transform ──────────────────────────────────────────────────────────
    //
    // scaleY(-1) mirrors the fill vertically so the shape's BOTTOM edge appears
    // at the TOP of the reflection (nearest the original = most opaque).
    //
    // CRITICAL — transform-origin must be "center" (50% 50%), NOT "top center":
    //
    // The reflection div sits at top:reflTop (= shapeHeight + gap) inside the wrapper.
    // scaleY(-1) flips content around the transform-origin point.
    //
    // With transform-origin: top center (WRONG):
    //   Flip axis = top edge of div (y=reflTop in wrapper).
    //   A point at local y=v maps to wrapper y = reflTop - v.
    //   The entire content renders UPWARD: from reflTop to reflTop-height.
    //   For reflTop=181, height=176 → content renders from y=5 to y=181 in wrapper.
    //   That overlaps the shape (y=0 to y=176) → reflection appears BEHIND the shape.
    //
    // With transform-origin: center (CORRECT):
    //   Flip axis = center of div (y = reflTop + height/2 in wrapper).
    //   A point at local y=v maps to wrapper y = reflTop + (height - v).
    //   v=0   (original top)    → wrapper y = reflTop + height  (bottom of refl area)
    //   v=h   (original bottom) → wrapper y = reflTop            (top of refl area = gap boundary)
    //   Content renders DOWNWARD from reflTop to reflTop+height → entirely below shape ✓
    //
    // NOTE: scaleY(-1) also flips the element's LOCAL coordinate system, which
    // affects how mask-image gradients are interpreted. See section 3 below.

    const transforms = [];
    if (Math.abs(scaleY - 1) > 0.001) transforms.push(`scaleY(${scaleY})`);
    if (Math.abs(scaleX - 1) > 0.001) transforms.push(`scaleX(${scaleX})`);
    if (Math.abs(skewXDeg) > 0.01) transforms.push(`skewX(${Math.round(skewXDeg * 100) / 100}deg)`);
    if (Math.abs(skewYDeg) > 0.01) transforms.push(`skewY(${Math.round(skewYDeg * 100) / 100}deg)`);
    const transformCSS = transforms.length ? transforms.join(' ') : 'none';

    // ── 3. Mask gradient ──────────────────────────────────────────────────────
    //
    // KEY INSIGHT — scaleY(-1) flips the element's LOCAL coordinate system:
    //
    //   CSS mask-image is applied in the element's LOCAL (pre-transform) space.
    //   After scaleY(-1):
    //     Local y=0%   (top of div)    → renders visually at the BOTTOM of the reflection
    //     Local y=100% (bottom of div) → renders visually at the TOP of the reflection
    //                                    (i.e. directly below the shape = nearest to it)
    //
    // Therefore the mask gradient must be written in LOCAL space so that after
    // the flip it produces the correct visual result:
    //
    //   Desired visual result (post-flip, top → bottom):
    //     visual-top    (near shape) = startAlpha  (most opaque)
    //     visual-bottom (far shape)  = endAlpha    (transparent)
    //
    //   In local space this must map to:
    //     local y=100% (= visual top)  = startAlpha
    //     local y=(100-endPosPct)%     = endAlpha
    //     local y=0%   (= visual btm)  = endAlpha  (or whatever the far-end value is)
    //
    // So we use "to bottom" in local space but with the stops REVERSED relative
    // to their visual intent:
    //
    //   linear-gradient(to bottom,
    //     endAlpha   at 0%,                  ← local top = visual bottom (far)
    //     endAlpha   at (100-endPosPct)%,    ← fade boundary (mirror of endPos)
    //     startAlpha at (100-startPosPct)%,  ← most opaque boundary (mirror of stPos)
    //     startAlpha at 100%                 ← local bottom = visual top (near shape)
    //   )
    //
    // After scaleY(-1) this reads visually as:
    //   visual top (near shape) → startAlpha (opaque) ✓
    //   visual bottom (far)     → endAlpha (transparent) ✓
    //
    // OOXML stA/endA: 0–100000 where 100000 = fully opaque.
    // CSS mask: rgba(0,0,0,A) where A=1 = fully visible, A=0 = hidden.
    // stA=87000 → 0.87, stA=50000 → 0.5  ✓

    // Detect whether scaleY will flip the element (sy is negative).
    // When flipped we must invert the gradient stops; when not flipped we use
    // the natural direction (legacy / future non-standard reflections).
    const isFlippedY = scaleY < 0;

    let maskGradient;
    if (isFlippedY) {
        // Build stops in LOCAL space (pre-flip). After scaleY(-1) these render
        // in the correct visual order (opaque near shape, transparent far).
        const mirrorStops = [];

        // local y=0% = visual bottom (far from shape) → endAlpha
        mirrorStops.push(`rgba(0,0,0,${endAlpha}) 0%`);

        // fade boundary: mirror of endPosPct
        const mirrorEndPct = Math.round((100 - endPosPct) * 100) / 100;
        mirrorStops.push(`rgba(0,0,0,${endAlpha}) ${mirrorEndPct}%`);

        // If stPos > 0 there is a flat-opaque band before the gradient starts
        if (startPosPct > 0) {
            const mirrorStartPct = Math.round((100 - startPosPct) * 100) / 100;
            mirrorStops.push(`rgba(0,0,0,${startAlpha}) ${mirrorStartPct}%`);
        }

        // local y=100% = visual top (nearest shape) → startAlpha
        mirrorStops.push(`rgba(0,0,0,${startAlpha}) 100%`);

        maskGradient = `linear-gradient(to bottom, ${mirrorStops.join(', ')})`;
    } else {
        // No Y-flip: gradient runs naturally from visual top (near shape) downward.
        const maskStops = [];
        if (startPosPct > 0) {
            maskStops.push(`rgba(0,0,0,${startAlpha}) 0%`);
        }
        maskStops.push(`rgba(0,0,0,${startAlpha}) ${startPosPct}%`);
        maskStops.push(`rgba(0,0,0,${endAlpha}) ${endPosPct}%`);
        if (endPosPct < 100 && endAlpha > 0) {
            maskStops.push(`transparent 100%`);
        }
        maskGradient = `linear-gradient(to bottom, ${maskStops.join(', ')})`;
    }

    // ── 4. Filter blur ────────────────────────────────────────────────────────
    // Applied on the SAME element as the mask so the mask correctly fades the
    // blurred result (avoids the Chromium compositing-layer / parent-mask bug).
    const filterCSS = blurPx > 0 ? `filter: blur(${blurPx}px);` : '';

    // ── 5. Assemble — single div with ALL properties ──────────────────────────
    // All CSS effects (transform → filter → mask) live on one element.
    // No overflow:hidden — lets blur spread naturally at the gap boundary.
    // No inner div — eliminates the Chromium mask-bypassing bug entirely.
    //
    // extraShapeCSS is stripped of any overflow:hidden / overflow:visible rules
    // to avoid accidentally clipping the blur spread.
    const safeExtraCSS = (extraShapeCSS || '')
        .replace(/overflow\s*:\s*\w+\s*;?/g, '')
        .trim();

    return `<div class="shape-reflection" aria-hidden="true" style="
        position: absolute;
        left: ${reflLeft}px;
        top: ${reflTop}px;
        width: ${width}px;
        height: ${height}px;
        background: ${fillColor};
        border-radius: ${borderRadius};
        ${safeExtraCSS}
        transform: ${transformCSS};
        transform-origin: center;
        ${filterCSS}
        -webkit-mask-image: ${maskGradient};
        mask-image: ${maskGradient};
        -webkit-mask-size: 100% 100%;
        mask-size: 100% 100%;
        pointer-events: none;
        z-index: ${Math.max(0, (zIndex || 1) - 1)};
        box-sizing: border-box;
    "></div>`;
}


// ─── wrapShapeWithReflection ──────────────────────────────────────────────────
//
// Wraps the original shape HTML + the reflection HTML inside a common
// absolutely-positioned container so both share the same coordinate origin.
//
// ── Double-positioning fix ────────────────────────────────────────────────────
// The incoming shapeHTML has position:absolute; left:Xpx; top:Ypx already set.
// The wrapper is also placed at left:X top:Y. Without correction the inner shape
// would render at X+X, Y+Y (doubled). The fix: regex-patch the inner shape's
// left/top to 0px so the wrapper owns the slide coordinate.
//
// ── pointer-events ────────────────────────────────────────────────────────────
// The wrapper must NOT have pointer-events:none — that would block all user
// interaction (click, drag, text edit) on the shape inside it.
// Only the reflection div gets pointer-events:none (set in buildReflectionHTML).
//
// @param {string} shapeHTML        – existing shape HTML (from convertShapeToHTML)
// @param {object} data             – output of getShapeReflectionData()
// @param {object} position         – { x, y, width, height, rotation }
// @param {string} fillColor        – CSS colour/gradient for the reflection clone
// @param {string} borderRadius     – CSS border-radius ("0px", "50%", etc.)
// @param {string} [extraShapeCSS]  – additional shape CSS (border, etc.)
// @param {number} [zIndex]         – z-index of the shape
//
// @returns {string}  HTML string – wrapper div containing shape + reflection
//
function wrapShapeWithReflection(shapeHTML, data, position, fillColor, borderRadius, extraShapeCSS, zIndex) {
    if (!data) return shapeHTML;

    const { width, height, x, y, rotation } = position;

    // Patch the inner shape's absolute position to (0,0) within the wrapper.
    // Only replace the FIRST occurrence so we don't affect nested child divs.
    const innerShapeHTML = shapeHTML
        .replace(/\bleft\s*:\s*[\d.]+px/, 'left: 0px')
        .replace(/\btop\s*:\s*[\d.]+px/, 'top: 0px');

    // Container height: shape height + gap + visible reflection portion.
    // The reflection is visually faded to transparent at endPosPct% of shape height.
    const visibleReflectionHeight = height * (data.endPosPct / 100);
    const containerHeight = height + data.distPx + visibleReflectionHeight;

    const reflectionHTML = buildReflectionHTML(
        data, position, fillColor, borderRadius, extraShapeCSS || '', zIndex || 1
    );

    return `<div class="shape-reflection-wrapper" data-reflection="true" style="
        position: absolute;
        left: ${x}px;
        top: ${y}px;
        width: ${width}px;
        height: ${containerHeight}px;
        transform: rotate(${rotation || 0}deg);
        transform-origin: top left;
        overflow: visible;
        z-index: ${zIndex || 1};
        box-sizing: border-box;
    ">
        ${innerShapeHTML}
        ${reflectionHTML}
    </div>`;
}


module.exports = {
    getShapeReflectionData,
    buildReflectionHTML,
    wrapShapeWithReflection,
};