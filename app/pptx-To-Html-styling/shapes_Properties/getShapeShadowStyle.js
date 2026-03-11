const colorHelper = require("../../api/helper/colorHelper.js");

// ─── EMU constants ────────────────────────────────────────────────────────────
const EMU_PER_PX = 12700; // 96 dpi: 914400 EMU/in ÷ 72 pt/in * (72/96) ≈ 12700 per CSS px

// ─── Preset color table (subset used by shadows) ─────────────────────────────
const PRESET_COLORS = {
    black: "#000000",
    white: "#ffffff",
    red: "#ff0000",
    green: "#008000",
    blue: "#0000ff",
    yellow: "#ffff00",
    cyan: "#00ffff",
    magenta: "#ff00ff",
    gray: "#808080",
    grey: "#808080",
    silver: "#c0c0c0",
    darkgray: "#a9a9a9",
    darkgrey: "#a9a9a9",
    navy: "#000080",
    teal: "#008080",
    maroon: "#800000",
    purple: "#800080",
    orange: "#ffa500",
    brown: "#a52a2a",
    pink: "#ffc0cb",
    gold: "#ffd700",
    transparent: "transparent",
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Convert a 6-char hex color string to "r, g, b" components.
 */
function hexToRgbComponents(hex) {
    if (!hex || hex === "transparent") return "0, 0, 0";
    const clean = hex.replace(/^#/, "");
    if (clean.length < 6) return "0, 0, 0";
    const r = parseInt(clean.substring(0, 2), 16) || 0;
    const g = parseInt(clean.substring(2, 4), 16) || 0;
    const b = parseInt(clean.substring(4, 6), 16) || 0;
    return `${r}, ${g}, ${b}`;
}

/**
 * Resolve any OOXML color node inside <a:outerShdw> to a hex string.
 * Supports: <a:prstClr>, <a:srgbClr>, <a:schemeClr>, <a:sysClr>
 */
function resolveShadowColor(shadowNode, themeXML, masterXML) {
    // 1. Preset color  (most common for shadows: val="black")
    const prstClrNode = shadowNode["a:prstClr"]?.[0];
    if (prstClrNode) {
        const val = prstClrNode?.["$"]?.val || "black";
        return PRESET_COLORS[val] ?? PRESET_COLORS["black"];
    }

    // 2. Direct RGB color
    const srgbClrNode = shadowNode["a:srgbClr"]?.[0];
    if (srgbClrNode) {
        const val = srgbClrNode?.["$"]?.val;
        if (val) return `#${val}`;
    }

    // 3. Scheme color (resolved against theme)
    const schemeClrNode = shadowNode["a:schemeClr"]?.[0];
    if (schemeClrNode) {
        const key = schemeClrNode?.["$"]?.val;
        if (key) {
            const resolved = colorHelper.resolveThemeColorHelper(key, themeXML, masterXML);
            if (resolved) return resolved;
        }
    }

    // 4. System color fallback
    const sysClrNode = shadowNode["a:sysClr"]?.[0];
    if (sysClrNode) {
        const lastClr = sysClrNode?.["$"]?.lastClr;
        if (lastClr) return `#${lastClr}`;
    }

    return "#000000"; // final fallback
}

/**
 * Extract alpha value from the color child node (whichever color type was present).
 * Returns a float 0–1.
 */
function resolveShadowAlpha(shadowNode) {
    // Check all possible color child nodes for <a:alpha>
    const colorTypes = ["a:prstClr", "a:srgbClr", "a:schemeClr", "a:sysClr"];

    for (const type of colorTypes) {
        const colorNode = shadowNode[type]?.[0];
        if (!colorNode) continue;

        const alphaVal = colorNode["a:alpha"]?.[0]?.["$"]?.val;
        if (alphaVal !== undefined && alphaVal !== null) {
            const parsed = parseInt(alphaVal, 10);
            if (!isNaN(parsed)) {
                return parsed / 100000; // OOXML: 100000 = 100% opaque
            }
        }
    }

    return 0.8; // default: 80% opacity (common for PPT shadows)
}

/**
 * Convert dir (60000ths-of-a-degree) + dist (EMU) to CSS offsetX/offsetY (px).
 *
 * OOXML angle convention:
 *   0°       = rightward  (East)
 *   5400000  = downward   (South)   [90°]
 *   10800000 = leftward   (West)    [180°]
 *   16200000 = upward     (North)   [270°]
 *
 * CSS box-shadow: positive X = right, positive Y = down.
 */
function dirDistToOffset(dirRaw, distRaw) {
    const dir = (dirRaw ?? 0) / 60000; // degrees, 0–359
    const dist = (distRaw ?? 0) / EMU_PER_PX; // px

    const rad = (dir * Math.PI) / 180;
    const offX = Math.round(dist * Math.cos(rad) * 10) / 10;
    const offY = Math.round(dist * Math.sin(rad) * 10) / 10;

    return { offX, offY };
}

/**
 * Convert blurRad (EMU) to CSS blur px.
 * PPT blurRad is the *radius* (same as CSS blur-radius).
 */
function blurRadToPx(blurRadRaw) {
    if (!blurRadRaw) return 0;
    return Math.round((parseInt(blurRadRaw, 10) / EMU_PER_PX) * 10) / 10;
}

/**
 * Approximate spread from sx/sy scale factors.
 * sx/sy = 100000 means 1× (no scale).  101000 → 1.01× (1% bigger).
 * We express the extra size as a spread px value relative to the shape,
 * but since CSS spread is absolute px we just use a small approximation:
 * spread ≈ (avgScale - 1) × blurPx   (feels right visually).
 */
function scaleToSpread(sxRaw, syRaw, blurPx) {
    const sx = sxRaw ? parseInt(sxRaw, 10) / 100000 : 1;
    const sy = syRaw ? parseInt(syRaw, 10) / 100000 : 1;
    const avg = (sx + sy) / 2;
    if (avg <= 1) return 0;
    return Math.round((avg - 1) * blurPx * 10) / 10;
}

/**
 * algn affects which corner of the shape the shadow originates from.
 * We translate this to a small additional CSS offset nudge so the shadow
 * visually matches the "alignment" (origin point) of the shadow.
 *
 * algn values (OOXML ST_RectAlignment):
 *   tl, t, tr, l, ctr, r, bl, b, br
 *
 * The offsets here are proportional adjustments in the same direction as
 * the alignment corner/edge relative to the shape center.  They are small
 * and only noticeable when dist=0 (glow-style shadows).
 */
function algnToNudge(algn) {
    // These nudges are deliberately 0 for most cases when dist is already set.
    // They matter most for "glow" (dist=0) shadows where algn shifts the glow origin.
    // Return [nudgeX, nudgeY] — expressed as fractions; caller multiplies by blurPx.
    const map = {
        tl: [-0.2, -0.2],
        t: [0.0, -0.2],
        tr: [0.2, -0.2],
        l: [-0.2, 0.0],
        ctr: [0.0, 0.0],
        r: [0.2, 0.0],
        bl: [-0.2, 0.2],
        b: [0.0, 0.2],
        br: [0.2, 0.2],
    };
    return map[algn] ?? [0, 0];
}


// ─── Main exported function ───────────────────────────────────────────────────

/**
 * getShapeShadowStyle(shapeNode, themeXML, masterXML, clrMap)
 *
 * Returns a CSS property string like:
 *   "box-shadow: 4px 6px 12px 0.5px rgba(0, 0, 0, 0.8);"
 * or "" when no shadow is defined.
 *
 * Can be used as a standalone function OR as a class method:
 *   // As a method inside ShapeHandler (rename getShadowStyle → getShapeShadowStyle):
 *   getShapeShadowStyle(shapeNode) {
 *       return getShapeShadowStyle(shapeNode, this.themeXML, this.masterXML, this.clrMap);
 *   }
 */
function getShapeShadowStyle(shapeNode, themeXML, masterXML, clrMap) {
    try {
        const effectLst = shapeNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0];
        if (!effectLst) return "";

        const outerShdw = effectLst?.["a:outerShdw"]?.[0];
        if (!outerShdw) return "";

        const attrs = outerShdw["$"] ?? {};

        // ── 1. Color + alpha ──────────────────────────────────────────────────
        const color = resolveShadowColor(outerShdw, themeXML, masterXML);
        const alpha = resolveShadowAlpha(outerShdw);

        // ── 2. Blur ───────────────────────────────────────────────────────────
        const blurPx = blurRadToPx(attrs.blurRad);

        // ── 3. Offset from dir + dist ─────────────────────────────────────────
        const distRaw = attrs.dist ? parseInt(attrs.dist, 10) : 0;
        const dirRaw = attrs.dir ? parseInt(attrs.dir, 10) : 0;
        let { offX, offY } = dirDistToOffset(dirRaw, distRaw);

        // ── 4. Alignment nudge (only significant when dist=0) ─────────────────
        if (distRaw === 0 && attrs.algn) {
            const [nx, ny] = algnToNudge(attrs.algn);
            offX += Math.round(nx * blurPx * 10) / 10;
            offY += Math.round(ny * blurPx * 10) / 10;
        }

        // ── 5. Spread from sx/sy ──────────────────────────────────────────────
        const spread = scaleToSpread(attrs.sx, attrs.sy, blurPx);

        // ── 6. Compose box-shadow ─────────────────────────────────────────────
        const rgb = hexToRgbComponents(color);

        // Round values for cleaner CSS output
        const cssOffX = Math.round(offX * 10) / 10;
        const cssOffY = Math.round(offY * 10) / 10;
        const cssBlur = Math.round(blurPx * 10) / 10;
        const cssSpread = Math.round(spread * 10) / 10;
        const cssAlpha = Math.round(alpha * 100) / 100;

        return `box-shadow: ${cssOffX}px ${cssOffY}px ${cssBlur}px ${cssSpread}px rgba(${rgb}, ${cssAlpha});`;

    } catch (err) {
        console.error("getShapeShadowStyle error:", err);
        return "";
    }
}

module.exports = { getShapeShadowStyle };