/**
 * handleShapeGlow.js
 * ------------------
 * HTML → PPTX glow conversion for shapes.
 */

'use strict';

// ─── Constants ────────────────────────────────────────────────────────────────

const PX_TO_EMU = 12700; // 96 dpi
const LAYER1_BLUR_FACTOR = 1.0;  // blur1 = radPx * 1.0  (from getShapeGlow.js)
const LAYER2_BLUR_FACTOR = 2.0;  // blur2 = radPx * 2.0
const LAYER1_ALPHA_BASE = 0.65; // L1 alpha = origAlpha * 0.65 * comp
const LAYER2_ALPHA_BASE = 0.40; // L2 alpha = origAlpha * 0.40 * comp
const GLOW_REF_RAD = 18;   // reference radius used when tuning multipliers
const GLOW_COMP_CAP = 3.0;  // max compensation factor (matches getShapeGlow.js)

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Extracts glow parameters from a .shape-glow-wrapper element or its style string.
 */
function extractGlowFromWrapperHTML(styleOrHTML, wrapperEl) {
    if (!styleOrHTML) return null;

    // ── Priority 1: data attributes (lossless, set by shapeHandler.js) ────────
    if (wrapperEl && typeof wrapperEl.getAttribute === 'function') {
        const dataAlpha = wrapperEl.getAttribute('data-glow-alpha');
        const dataRad = wrapperEl.getAttribute('data-glow-rad');
        if (dataAlpha && dataRad) {
            const alphaVal = parseInt(dataAlpha, 10);
            const radEmu = parseInt(dataRad, 10);
            if (!isNaN(alphaVal) && !isNaN(radEmu) && radEmu > 0) {
                // Still parse colour from CSS filter (not stored in data attr)
                const filterMatch = styleOrHTML.match(/filter\s*:\s*([^;]+)/i);
                if (filterMatch) {
                    const layers = _extractDropShadowLayers(filterMatch[1].trim());
                    if (layers.length > 0) {
                        const colorHex = _rgbToHex(layers[0].r, layers[0].g, layers[0].b);
                        return { radEmu, alphaVal, colorHex };
                    }
                }
            }
        }
    }

    // ── Priority 2: CSS reverse-engineering (fallback) ────────────────────────
    const filterMatch = styleOrHTML.match(/filter\s*:\s*([^;]+)/i);
    if (!filterMatch) return null;
    return extractGlowFromFilterCSS(filterMatch[1].trim());
}

/**
 * Extracts glow parameters from a CSS filter value string using comp-aware
 * reverse-engineering.
 */
function extractGlowFromFilterCSS(filterValue) {
    if (!filterValue) return null;

    const layers = _extractDropShadowLayers(filterValue);
    if (layers.length === 0) return null;

    const l1 = layers[0];
    if (Math.abs(l1.offsetX) > 0.5 || Math.abs(l1.offsetY) > 0.5) return null;

    // ── Radius from blur1 ──────────────────────────────────────────────────────
    const radPx = l1.blur / LAYER1_BLUR_FACTOR;
    const radEmu = Math.round(radPx * PX_TO_EMU);

    // ── Compensation factor (must match getShapeGlow.js exactly) ──────────────
    const comp = Math.min(Math.sqrt(radPx / GLOW_REF_RAD), GLOW_COMP_CAP);

    // ── Recover original alpha ─────────────────────────────────────────────────
    let origAlpha;
    if (l1.alpha < 0.999) {
        // L1 was not clamped — direct inversion
        origAlpha = l1.alpha / (LAYER1_ALPHA_BASE * comp);
    } else {
        // L1 was clamped to 1.0 — use L2 which is not clamped
        const l2 = layers[1];
        if (l2) {
            origAlpha = l2.alpha / (LAYER2_ALPHA_BASE * comp);
        } else {
            // Only 1 layer (shouldn't happen) — best effort
            origAlpha = 1.0 / (LAYER1_ALPHA_BASE * comp);
        }
    }

    const alphaVal = Math.round(Math.min(1, Math.max(0, origAlpha)) * 100000);
    const colorHex = _rgbToHex(l1.r, l1.g, l1.b);

    return { radEmu, colorHex, alphaVal };
}

/**
 * Shared helper: extracts all drop-shadow layers from a filter value string.
 * Handles nested parens in rgba().
 */
function _extractDropShadowLayers(filterValue) {
    const re = /drop-shadow\(([^()]*(?:\([^()]*\)[^()]*)*)\)/gi;
    const layers = [];
    let m;
    while ((m = re.exec(filterValue)) !== null) {
        const parsed = _parseDropShadow(m[1].trim());
        if (parsed) layers.push(parsed);
    }
    return layers;
}

/**
 * Generates the OOXML <a:effectLst><a:glow> XML string.
 *
 * @param {{ radEmu: number, colorHex: string, alphaVal: number }} glowData
 * @returns {string}  e.g. '<a:effectLst><a:glow rad="228600">...</a:glow></a:effectLst>'
 */
function buildGlowXML(glowData) {
    if (!glowData) return '';
    const { radEmu, colorHex, alphaVal } = glowData;
    const hex = colorHex.replace('#', '').toUpperCase();

    return (
        `<a:effectLst>` +
        `<a:glow rad="${radEmu}">` +
        `<a:srgbClr val="${hex}">` +
        `<a:alpha val="${alphaVal}"/>` +
        `</a:srgbClr>` +
        `</a:glow>` +
        `</a:effectLst>`
    );
}

/**
 * Stores glow data on the slide object so the post-processor can find it.
 * Keyed by shape name (data-name attribute).
 */
function storeShapeGlow(pptxSlide, shapeName, glowData) {
    if (!pptxSlide._glowMap) pptxSlide._glowMap = {};
    pptxSlide._glowMap[shapeName] = glowData;
}

/**
 * Post-processing step: injects <a:effectLst><a:glow> into the raw slide XML
 * for every shape recorded in glowMap.
 */
function injectGlowIntoSlideXml(slideXml, glowMap) {
    if (!glowMap || !slideXml) return slideXml;

    let xml = slideXml;

    for (const [shapeName, glowData] of Object.entries(glowMap)) {
        if (!glowData) continue;

        const glowXml = buildGlowXML(glowData);
        const glowInner = buildGlowXML(glowData).replace('<a:effectLst>', '').replace('</a:effectLst>', '');
        // e.g.  <a:glow rad="228600"><a:srgbClr val="FF6A1C"><a:alpha val="73000"/></a:srgbClr></a:glow>

        // Escape shape name for use in regex
        const escapedName = shapeName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        // Match the entire <p:sp> block for this shape name
        // We look for the cNvPr name attribute then walk to </p:spPr>
        const spBlockRe = new RegExp(
            `(<p:sp>(?:(?!<p:sp>)[\\s\\S])*?<p:cNvPr[^>]+name="${escapedName}"[^>]*>[\\s\\S]*?<\\/p:spPr>)`,
            'g'
        );

        xml = xml.replace(spBlockRe, (spBlock) => {
            // Case 1: <a:effectLst> already exists — merge <a:glow> into it
            if (spBlock.includes('<a:effectLst>')) {
                // Insert glow inside existing effectLst (before closing tag)
                return spBlock.replace(
                    '<a:effectLst>',
                    `<a:effectLst>${glowInner}`
                );
            }

            // Case 2: no <a:effectLst> — inject the full block before </p:spPr>
            return spBlock.replace('</p:spPr>', `${glowXml}</p:spPr>`);
        });
    }

    return xml;
}

// ─── Private helpers ──────────────────────────────────────────────────────────

/**
 * Parses a single drop-shadow() argument string.
 * Handles both:
 */
function _parseDropShadow(arg) {
    // Extract rgba/rgb color
    const rgbaRe = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i;
    const colorMatch = arg.match(rgbaRe);
    if (!colorMatch) return null;

    const r = parseInt(colorMatch[1], 10);
    const g = parseInt(colorMatch[2], 10);
    const b = parseInt(colorMatch[3], 10);
    const alpha = colorMatch[4] !== undefined ? parseFloat(colorMatch[4]) : 1;

    // Remove color token from arg to parse the length values
    const withoutColor = arg.replace(rgbaRe, '').trim();

    // Extract all numeric length values — handles both "0px" and bare "0"
    // Format: offsetX offsetY blur [spread]
    const tokenRe = /(-?[\d.]+)(px)?/g;
    const values = [];
    let m;
    while ((m = tokenRe.exec(withoutColor)) !== null) {
        values.push(parseFloat(m[1]));
    }

    const offsetX = values[0] ?? 0;
    const offsetY = values[1] ?? 0;
    const blur = values[2] ?? values[0] ?? 0;

    return { offsetX, offsetY, blur, r, g, b, alpha };
}

function _rgbToHex(r, g, b) {
    return '#' + [r, g, b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

// ─── Exports ──────────────────────────────────────────────────────────────────

module.exports = {
    extractGlowFromFilterCSS,
    extractGlowFromWrapperHTML,
    buildGlowXML,
    storeShapeGlow,
    injectGlowIntoSlideXml,
};