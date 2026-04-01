const colorHelper = require("../../api/helper/colorHelper.js");

// ─── EMU constants ────────────────────────────────────────────────────────────
const EMU_PER_PX = 12700; // 914400 EMU/in ÷ 72 pt/in = 12700 EMU per CSS px


// ─── Main exported function ───────────────────────────────────────────────────

function getShapeShadowStyle(shapeNode, themeXML, masterXML, clrMap) {
    try {

        const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
        const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
        const TEXT_PH_TYPES = ["ctrTitle", "title", "subTitle", "body", "obj"];
        const isTextPlaceholder = TEXT_PH_TYPES.includes(phType);
        const hasExplicitShapeFill =
            shapeNode?.["p:spPr"]?.[0]?.["a:solidFill"] ||
            shapeNode?.["p:spPr"]?.[0]?.["a:gradFill"] ||
            shapeNode?.["p:spPr"]?.[0]?.["a:pattFill"];

        if ((isTextBox || isTextPlaceholder) && !hasExplicitShapeFill) {
            return ""; // text-shadow is handled by textHandler.js on the spans
        }

        // ── Resolve effectLst (multi-level fallback) ──────────────────────────
        const effectLst =
            shapeNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]                                // Level A
            ?? shapeNode?.["a:effectLst"]?.[0]                                               // Level B
            ?? (shapeNode?.["a:outerShdw"] || shapeNode?.["a:innerShdw"] ? shapeNode : null);// Level C

        if (!effectLst) return "";

        // ── Detect outer vs inner shadow ──────────────────────────────────────
        const outerShdw = effectLst?.["a:outerShdw"]?.[0];
        const innerShdw = effectLst?.["a:innerShdw"]?.[0];
        const shdwNode = outerShdw ?? innerShdw;
        if (!shdwNode) return "";

        const isInner = !outerShdw && !!innerShdw;
        const attrs = shdwNode["$"] ?? {};

        // ── 1. Color + alpha ──────────────────────────────────────────────────
        const color = resolveShadowColor(shdwNode, themeXML, masterXML);
        const alpha = resolveShadowAlpha(shdwNode);

        // ── 2. Blur ───────────────────────────────────────────────────────────
        const blurPx = blurRadToPx(attrs.blurRad);

        // ── 3. Offset (dir + dist + algn) ─────────────────────────────────────
        // Use null for absent dir so dirDistToOffset can distinguish it from
        // an explicit 0° (East) direction.
        const distRaw = attrs.dist ? parseInt(attrs.dist, 10) : 0;
        const dirRaw = (attrs.dir != null && attrs.dir !== "")
            ? parseInt(attrs.dir, 10)
            : null;

        let { offX, offY } = dirDistToOffset(dirRaw, distRaw, attrs.algn);

        if (isInner) {
            offX = -offX;
            offY = -offY;
        }

        // ── 5. Spread (outerShdw sx/sy; innerShdw has no scale attrs) ─────────
        let spread = isInner ? 0 : scaleToSpread(attrs.sx, attrs.sy, blurPx);

        // ── 5b. Alignment-aware clip adjustment ───────────────────────────────
        // When algn != ctr and sx/sy < 1, PowerPoint anchors the shadow scaling
        // to a specific edge/corner so the shadow only bleeds on the intended
        // side(s).  Apply a negative-spread + offset compensation to match.
        if (!isInner) {
            const sxVal = attrs.sx ? parseInt(attrs.sx, 10) / 100000 : 1;
            const syVal = attrs.sy ? parseInt(attrs.sy, 10) / 100000 : 1;
            const dirDeg = (dirRaw != null) ? (dirRaw / 60000) : 0;
            const { spreadDelta, offXDelta, offYDelta } =
                getAlgnClipAdjustment(attrs.algn, dirDeg, blurPx, sxVal, syVal);
            spread += spreadDelta;
            offX += offXDelta;
            offY += offYDelta;
        }

        // ── 6. Compose final CSS ──────────────────────────────────────────────
        const rgb = hexToRgbComponents(color);
        const cssOffX = Math.round(offX * 1000) / 1000;
        const cssOffY = Math.round(offY * 1000) / 1000;
        const cssBlur = Math.round(blurPx * 10) / 10;
        const cssSpread = Math.round(spread * 1000) / 1000;
        // alpha at 3dp: preserves non-round OOXML values like 66667 → 0.667
        const cssAlpha = Math.round(alpha * 1000) / 1000;
        const inset = isInner ? "inset " : "";

        return `box-shadow: ${inset}${cssOffX}px ${cssOffY}px ${cssBlur}px ${cssSpread}px rgba(${rgb}, ${cssAlpha});`;

    } catch (err) {
        console.error("getShapeShadowStyle error:", err);
        return "";
    }
}

function resolveShadowColor(shadowNode, themeXML, masterXML) {
    // 1. Preset color (e.g. val="black")
    const prstClrNode = shadowNode["a:prstClr"]?.[0];
    if (prstClrNode) {
        const val = prstClrNode?.["$"]?.val || "black";
        const baseColor = PRESET_COLORS[val] ?? PRESET_COLORS["black"];
        const transforms = extractColorTransforms(prstClrNode);
        return applyColorTransforms(baseColor, transforms);
    }

    // 2. Direct sRGB hex
    const srgbClrNode = shadowNode["a:srgbClr"]?.[0];
    if (srgbClrNode) {
        const val = srgbClrNode?.["$"]?.val;
        if (val) {
            const baseColor = `#${val}`;
            const transforms = extractColorTransforms(srgbClrNode);
            return applyColorTransforms(baseColor, transforms);
        }
    }

    // 3. Scheme color resolved against theme
    const schemeClrNode = shadowNode["a:schemeClr"]?.[0];
    if (schemeClrNode) {
        const key = schemeClrNode?.["$"]?.val;
        if (key) {
            const resolved = colorHelper.resolveThemeColorHelper(key, themeXML, masterXML);
            if (resolved) {
                // Apply color transforms (e.g. lumMod=50000 darkens the theme color by 50%)
                const transforms = extractColorTransforms(schemeClrNode);
                return applyColorTransforms(resolved, transforms);
            }
        }
    }

    // 4. System color (use lastClr attribute as fallback hex)
    const sysClrNode = shadowNode["a:sysClr"]?.[0];
    if (sysClrNode) {
        const lastClr = sysClrNode?.["$"]?.lastClr;
        if (lastClr) {
            const baseColor = `#${lastClr}`;
            const transforms = extractColorTransforms(sysClrNode);
            return applyColorTransforms(baseColor, transforms);
        }
    }

    return "#000000";
}


function resolveShadowAlpha(shadowNode) {
    const colorTypes = ["a:prstClr", "a:srgbClr", "a:schemeClr", "a:sysClr"];

    for (const type of colorTypes) {
        const colorNode = shadowNode[type]?.[0];
        if (!colorNode) continue;

        const alphaVal = colorNode["a:alpha"]?.[0]?.["$"]?.val;
        if (alphaVal !== undefined && alphaVal !== null) {
            const parsed = parseInt(alphaVal, 10);
            if (!isNaN(parsed)) return parsed / 100000; // 100000 = fully opaque
        }

        // Color node present but no <a:alpha> child → fully opaque
        return 1.0;
    }

    return 1.0; // no colour node found → fully opaque fallback
}


/**
 * Convert blurRad (EMU) to CSS blur px.
 */

function blurRadToPx(blurRadRaw) {
    if (!blurRadRaw) return 0;
    return Math.round((parseInt(blurRadRaw, 10) / EMU_PER_PX) * 10) / 10;
}


/**
 * When algn is a non-center alignment AND the shadow is scaled DOWN (sx/sy < 1),
 * PowerPoint anchors the scaling to that edge/corner of the shape.  The shadow
 * ends up inset from the opposite sides, so it only "bleeds" out the intended
 * side(s).  CSS box-shadow has no such concept, so we approximate it with:
 *   • spread = −blurPx        → collapses the shadow inward, killing unwanted bleed
 *   • offset += blurPx in dir → compensates so the intended side stays visible
 *
 * Returns deltas to add onto spread, offX, offY.
 */
function getAlgnClipAdjustment(algn, dirDeg, blurPx, sx, sy) {
    // No adjustment needed for center alignment or no alignment
    if (!algn || algn === "ctr") return { spreadDelta: 0, offXDelta: 0, offYDelta: 0 };

    // Only clip when the shadow is actually scaled DOWN
    const avgScale = (sx + sy) / 2;
    if (avgScale >= 1) return { spreadDelta: 0, offXDelta: 0, offYDelta: 0 };

    // Negative spread clips the unwanted bleed on all sides
    const spreadDelta = -blurPx;

    // Re-introduce blurPx in the shadow direction so the correct side stays visible
    const rad = (dirDeg * Math.PI) / 180;
    const offXDelta = blurPx * Math.cos(rad);
    const offYDelta = blurPx * Math.sin(rad);

    return { spreadDelta, offXDelta, offYDelta };
}

function dirDistToOffset(dirRaw, distRaw, algn) {
    const dist = (distRaw ?? 0) / EMU_PER_PX;

    if (dist === 0) {
        // No displacement — ambient / glow shadow
        return { offX: 0, offY: 0 };
    }

    // Use the explicit dir, or fall back to the OOXML schema default of 0° (East)
    const dirDeg = (dirRaw != null) ? (dirRaw / 60000) : 0;
    const rad = (dirDeg * Math.PI) / 180;

    return {
        offX: Math.round(dist * Math.cos(rad) * 1000) / 1000,
        offY: Math.round(dist * Math.sin(rad) * 1000) / 1000,
    };
}

function scaleToSpread(sxRaw, syRaw, blurPx) {
    const sx = sxRaw ? parseInt(sxRaw, 10) / 100000 : 1;
    const sy = syRaw ? parseInt(syRaw, 10) / 100000 : 1;
    const avg = (sx + sy) / 2;
    if (avg <= 1) return 0;
    return Math.round((avg - 1) * blurPx * 1000) / 1000;
}

/**
 * Convert a 6-char hex string to "r, g, b" components.
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

function hexToHls(hex) {
    const clean = hex.replace(/^#/, "");
    const r = parseInt(clean.substring(0, 2), 16) / 255;
    const g = parseInt(clean.substring(2, 4), 16) / 255;
    const b = parseInt(clean.substring(4, 6), 16) / 255;

    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const l = (max + min) / 2;
    let h = 0, s = 0;

    if (max !== min) {
        const d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
            case g: h = ((b - r) / d + 2) / 6; break;
            case b: h = ((r - g) / d + 4) / 6; break;
        }
    }
    return [h, l, s];
}

/**
 * Convert [H, L, S] (0–1 each) back to a #RRGGBB hex string.
 */
function hlsToHex(h, l, s) {
    l = Math.max(0, Math.min(1, l));
    s = Math.max(0, Math.min(1, s));

    function hueToRgb(p, q, t) {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    }

    let r, g, b;
    if (s === 0) {
        r = g = b = l;
    } else {
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r = hueToRgb(p, q, h + 1 / 3);
        g = hueToRgb(p, q, h);
        b = hueToRgb(p, q, h - 1 / 3);
    }

    const toHex = (x) => Math.round(x * 255).toString(16).padStart(2, "0").toUpperCase();
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}


function applyColorTransforms(hexColor, transforms) {
    if (!transforms || transforms.length === 0) return hexColor;
    if (!hexColor || hexColor === "transparent") return hexColor;

    let [h, l, s] = hexToHls(hexColor);

    for (const { name, val } of transforms) {
        const factor = parseInt(val, 10) / 100000;
        switch (name) {
            case "lumMod": l = l * factor; break;
            case "lumOff": l = l + factor; break;
            case "satMod": s = s * factor; break;
            case "tint": l = l + (1.0 - l) * factor; break;
            case "shade": l = l * factor; break;
            // "alpha" is handled separately by resolveShadowAlpha() — skip here
        }
    }

    return hlsToHex(h, l, s);
}

function extractColorTransforms(colorNode) {
    const transformNames = ["lumMod", "lumOff", "satMod", "tint", "shade"];
    const result = [];
    for (const tname of transformNames) {
        const tnode = colorNode[`a:${tname}`]?.[0];
        if (tnode) {
            const val = tnode["$"]?.val;
            if (val != null) result.push({ name: tname, val });
        }
    }
    return result;
}


// ─── Custom shape shadow (filter: drop-shadow) ───────────────────────────────

function getCustomShapeShadowStyle(shapeNode, themeXML, masterXML, clrMap) {
    try {
        // ── Resolve effectLst (same multi-level fallback as getShapeShadowStyle) ─
        const effectLst =
            shapeNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]
            ?? shapeNode?.["a:effectLst"]?.[0]
            ?? (shapeNode?.["a:outerShdw"] || shapeNode?.["a:innerShdw"] ? shapeNode : null);

        if (!effectLst) return "";

        const outerShdw = effectLst?.["a:outerShdw"]?.[0];
        const innerShdw = effectLst?.["a:innerShdw"]?.[0];
        if (!outerShdw) {
            if (innerShdw) {
                console.warn(
                    "[getCustomShapeShadowStyle] Inner shadow on custGeom shape skipped — " +
                    "CSS drop-shadow() has no inset variant. Shape:",
                    shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name ?? "(unnamed)"
                );
            }
            return "";
        }

        const attrs = outerShdw["$"] ?? {};

        // ── 1. Color + alpha ──────────────────────────────────────────────────
        const color = resolveShadowColor(outerShdw, themeXML, masterXML);
        const alpha = resolveShadowAlpha(outerShdw);

        // ── 2. Blur radius ────────────────────────────────────────────────────
        let blurPx = blurRadToPx(attrs.blurRad);

        // ── 3. Spread approximation added into blur ───────────────────────────

        const sx = attrs.sx ? parseInt(attrs.sx, 10) / 100000 : 1;
        const sy = attrs.sy ? parseInt(attrs.sy, 10) / 100000 : 1;
        const avgScale = (sx + sy) / 2;
        if (avgScale > 1) {
            // Spread surplus: e.g. scale=1.02 → adds 2% of blurPx as extra blur
            blurPx = Math.round((blurPx + (avgScale - 1) * blurPx) * 10) / 10;
        }

        // ── 4. Offset (dir + dist) ────────────────────────────────────────────
        const distRaw = attrs.dist ? parseInt(attrs.dist, 10) : 0;
        const dirRaw = (attrs.dir != null && attrs.dir !== "")
            ? parseInt(attrs.dir, 10)
            : null;

        let { offX, offY } = dirDistToOffset(dirRaw, distRaw, attrs.algn);

        // ── 4b. Alignment-aware clip adjustment ───────────────────────────────
        // drop-shadow() has no spread parameter, so we bake the compensation
        // directly into blur and offset instead.
        const sxVal = attrs.sx ? parseInt(attrs.sx, 10) / 100000 : 1;
        const syVal = attrs.sy ? parseInt(attrs.sy, 10) / 100000 : 1;
        const dirDeg = (dirRaw != null) ? (dirRaw / 60000) : 0;
        const { spreadDelta, offXDelta, offYDelta } =
            getAlgnClipAdjustment(attrs.algn, dirDeg, blurPx, sxVal, syVal);
        // For drop-shadow we absorb the negative spread into a blur reduction
        blurPx = Math.max(0, Math.round((blurPx + spreadDelta) * 10) / 10);
        offX += offXDelta;
        offY += offYDelta;

        // ── 5. Compose CSS ────────────────────────────────────────────────────
        const rgb = hexToRgbComponents(color);
        const cssOffX = Math.round(offX * 10) / 10;
        const cssOffY = Math.round(offY * 10) / 10;
        const cssBlur = Math.round(blurPx * 10) / 10;
        // alpha at 3dp for same precision as getShapeShadowStyle
        const cssAlpha = Math.round(alpha * 1000) / 1000;

        return `filter: drop-shadow(${cssOffX}px ${cssOffY}px ${cssBlur}px rgba(${rgb}, ${cssAlpha}));`;

    } catch (err) {
        console.error("getCustomShapeShadowStyle error:", err);
        return "";
    }
}


module.exports = { getShapeShadowStyle, getCustomShapeShadowStyle };