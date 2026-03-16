/**
 * shape-reflection-processor.js
 *
 * Post-processing module that injects <a:reflection> OOXML into slide XML files
 * for shapes that have a .shape-reflection sibling in the source HTML.
 *
 * The HTML→PPTX pipeline cannot express reflections through pptxgenjs directly,
 * so we collect reflection metadata during the HTML parsing phase, store it in a
 * shared cache (keyed by shape name / data-name attribute), and then inject the
 * correct DrawingML <a:effectLst><a:reflection .../></a:effectLst> into the
 * already-generated slide XML files.
 *
 * OOXML attribute reference for <a:reflection>:
 *   blurRad  – blur radius in EMU  (1 px @ 72 dpi = 12700 EMU)
 *   stA      – start alpha, 0–100000  (100000 = fully opaque)
 *   endA     – end alpha,   0–100000  (0 = fully transparent, default)
 *   stPos    – start position, 0–100000 (default 0, omit when 0)
 *   endPos   – end position,   0–100000
 *   dist     – distance from shape bottom, in EMU
 *   dir      – direction, 5400000 = straight down (always for standard reflection)
 *   sy       – vertical scale, -100000 = flipped (always for standard reflection)
 *   algn     – alignment "bl" = bottom-left (always)
 *   rotWithShape – "0" (always)
 */

'use strict';

const fs = require('fs');
const fsP = require('fs').promises;
const path = require('path');

// ─── Shared in-memory store ────────────────────────────────────────────────
// Key:   shape data-name  (e.g. "Rectangle 3")
// Value: reflectionParams object (see parseReflectionFromHTML)
const _reflectionStore = new Map();

function clearReflectionStore() {
    _reflectionStore.clear();
}

function storeReflectionParams(shapeName, params) {
    _reflectionStore.set(shapeName, params);
}

function getReflectionStore() {
    return _reflectionStore;
}

// ─── HTML parsing helper ───────────────────────────────────────────────────

/**
 * Parse all shape-reflection-wrapper elements from a JSDOM slide element and
 * populate the reflection store.
 *
 * @param {Element} slideElement  – JSDOM element for a .sli-slide
 */
function collectReflectionsFromSlide(slideElement) {
    const wrappers = slideElement.querySelectorAll('.shape-reflection-wrapper[data-reflection="true"]');

    for (const wrapper of wrappers) {
        const shapeEl = wrapper.querySelector('.shape');
        const reflectionEl = wrapper.querySelector('.shape-reflection');

        if (!shapeEl || !reflectionEl) continue;

        const shapeName = shapeEl.getAttribute('data-name');
        if (!shapeName) continue;

        const params = parseReflectionElement(reflectionEl, shapeEl);
        if (params) {
            storeReflectionParams(shapeName, params);
        }
    }
}

/**
 * Extract reflection parameters from a .shape-reflection div.
 *
 * CSS properties we read:
 *   - filter: blur(Xpx)                      → blurPx
 *   - top  (absolute px from wrapper top)    → distPx = top - shapeHeight
 *   - mask-image / -webkit-mask-image        → gradient stops → startAlpha, endAlpha, endPosPct
 *
 * @param {Element} reflEl   – .shape-reflection element
 * @param {Element} shapeEl  – the real .shape element (for height)
 * @returns {object|null}
 */
function parseReflectionElement(reflEl, shapeEl) {
    try {
        const reflStyle = reflEl.getAttribute('style') || '';
        const styleObj = parseInlineStyle(reflStyle);

        // ── 1. blur radius ────────────────────────────────────────────────
        let blurPx = 0;
        const filterVal = styleObj['filter'] || '';
        const blurMatch = filterVal.match(/blur\(\s*([\d.]+)px\s*\)/i);
        if (blurMatch) blurPx = parseFloat(blurMatch[1]);

        // ── 2. dist (gap between shape bottom and reflection top) ─────────
        const reflTop = parseFloat(styleObj['top'] || '0');
        const shapeH = parseFloat(shapeEl.getAttribute('style') ? parseInlineStyle(shapeEl.getAttribute('style'))['height'] || '0' : '0');
        const distPx = Math.max(0, reflTop - shapeH);

        // ── 3. gradient stops from mask-image ─────────────────────────────
        const maskImage = styleObj['-webkit-mask-image'] || styleObj['mask-image'] || '';
        const gradStops = parseMaskGradientStops(maskImage);

        // gradStops is ordered: [transparentEnd%, opaqueStart%, opaqueAlpha]
        // The mirrored (corrected) gradient looks like:
        //   0%→ transparent, mirrorEndPct%→ transparent, 100%→ opaque(stA)
        // We need to reverse-engineer the OOXML params from this.
        //
        // OOXML endPos = (100 - mirrorEndPct) * 1000   [position where opacity fades to 0]
        // OOXML stA    = opaqueAlpha * 100000
        // OOXML endA   = 0 (transparent end, default)

        let startAlpha = 1.0;   // CSS alpha at the "near shape" end
        let endPosPct = 55;    // percent position where the gradient ends fading

        if (gradStops) {
            startAlpha = gradStops.opaqueAlpha;
            // The mask gradient stop at 100% holds the opaque side.
            // The mid-stop tells us where opacity starts (mirrors endPos).
            // mirrorEndPct = 100 - endPosPct  ↔  endPosPct = 100 - mirrorEndPct
            endPosPct = 100 - gradStops.midPosPct;
        }

        // ── 4. Convert to EMU / per-mille ─────────────────────────────────
        const PX_TO_EMU = 12700; // 914400 EMU/inch ÷ 72 px/inch

        return {
            blurRad: Math.round(blurPx * PX_TO_EMU),
            dist: Math.round(distPx * PX_TO_EMU),
            stA: Math.round(startAlpha * 100000),
            endA: 0,
            endPos: Math.round(endPosPct * 1000),
            dir: 5400000,   // straight down
            sy: -100000,   // flip vertical
            algn: 'bl',
            rotWithShape: 0
        };

    } catch (err) {
        console.warn(`[reflection-processor] Failed to parse reflection for shape:`, err.message);
        return null;
    }
}

/**
 * Parse a CSS mask linear-gradient string and return the key stop positions.
 *
 * Expected format (after the fix in getShapeReflection.js):
 *   linear-gradient(to bottom, rgba(0,0,0,0) 0%, rgba(0,0,0,0) 45%, rgba(0,0,0,0.87) 100%)
 *
 * Returns { midPosPct, opaqueAlpha } or null.
 */
function parseMaskGradientStops(maskImage) {
    if (!maskImage) return null;

    // Extract all "rgba(...) X%" tokens
    const stopRegex = /rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)\s*([\d.]+)%/g;
    const stops = [];
    let m;
    while ((m = stopRegex.exec(maskImage)) !== null) {
        stops.push({
            a: m[4] !== undefined ? parseFloat(m[4]) : 1,
            pos: parseFloat(m[5])
        });
    }

    if (stops.length < 2) return null;

    // Find the opaque stop (highest alpha) and the last transparent stop before it
    let opaqueStop = null;
    let midStop = null;

    for (const s of stops) {
        if (s.a > 0) {
            opaqueStop = s;
            break;
        }
        midStop = s;  // last transparent stop before the opaque one
    }

    // If only 2 stops: [transparent 0%, opaque 100%] – simple case
    if (!midStop) {
        // All stops before opaqueStop are at 0 alpha
        const transparentStops = stops.filter(s => s.a === 0);
        midStop = transparentStops[transparentStops.length - 1] || { pos: 0 };
    }

    if (!opaqueStop) return null;

    return {
        midPosPct: midStop.pos,
        opaqueAlpha: opaqueStop.a
    };
}

/**
 * Very minimal inline-style parser → plain key:value object.
 * Only needed for the properties we care about.
 */
function parseInlineStyle(styleStr) {
    const result = {};
    // Split on semicolons but be careful with nested parens (e.g. rgba(), blur())
    const parts = styleStr.split(';');
    for (const part of parts) {
        const colon = part.indexOf(':');
        if (colon < 0) continue;
        const key = part.slice(0, colon).trim().toLowerCase();
        const val = part.slice(colon + 1).trim();
        if (key) result[key] = val;
    }
    return result;
}

// ─── OOXML builder ─────────────────────────────────────────────────────────

/**
 * Build the <a:effectLst> XML snippet containing the <a:reflection> element.
 *
 * @param {object} p  – reflection params from parseReflectionElement
 * @returns {string}  – XML string
 */
function buildReflectionXML(p) {
    const attrs = [
        `blurRad="${p.blurRad}"`,
        `stA="${p.stA}"`,
        p.endA !== 0 ? `endA="${p.endA}"` : null,
        p.endPos ? `endPos="${p.endPos}"` : null,
        `dist="${p.dist}"`,
        `dir="${p.dir}"`,
        `sy="${p.sy}"`,
        `algn="${p.algn}"`,
        `rotWithShape="${p.rotWithShape}"`
    ].filter(Boolean).join(' ');

    return `<a:effectLst><a:reflection ${attrs}/></a:effectLst>`;
}

// ─── XML injection into slide files ───────────────────────────────────────

/**
 * Walk every slideN.xml in slideXmlsDir/slides/ and inject <a:reflection>
 * into any <p:sp> whose <p:cNvPr name="..."> matches a stored shape name.
 *
 * Strategy:
 *   For each matching <p:sp>:
 *     1. If <p:spPr> already contains <a:effectLst>, replace it.
 *     2. If <p:spPr> exists but has no <a:effectLst>, append before </p:spPr>.
 *     3. If <p:spPr> does not exist, this is a degenerate case — skip.
 *
 * @param {string} slideXmlsDir  – directory that contains slides/ subdir
 * @returns {Promise<{success:boolean, shapesInjected:number}>}
 */
async function injectReflectionsIntoSlideXML(slideXmlsDir) {
    const store = getReflectionStore();

    if (store.size === 0) {
        return { success: true, shapesInjected: 0 };
    }

    const slidesDir = path.join(slideXmlsDir, 'slides');

    let dirExists = false;
    try { dirExists = (await fsP.stat(slidesDir)).isDirectory(); } catch (_) { }
    if (!dirExists) return { success: true, shapesInjected: 0 };

    const slideFiles = (await fsP.readdir(slidesDir))
        .filter(f => /^slide\d+\.xml$/.test(f));

    let totalInjected = 0;

    for (const slideFile of slideFiles) {
        const slidePath = path.join(slidesDir, slideFile);
        let xml = await fsP.readFile(slidePath, 'utf8');
        const original = xml;
        let injectedInFile = 0;

        for (const [shapeName, params] of store.entries()) {
            const reflXml = buildReflectionXML(params);
            const newXml = injectReflectionForShape(xml, shapeName, reflXml);
            if (newXml !== xml) {
                totalInjected++;
                xml = newXml;
            }
        }

        if (xml !== original) {
            await fsP.writeFile(slidePath, xml, 'utf8');
        }
    }

    return { success: true, shapesInjected: totalInjected };
}

/**
 * Inject reflection XML into the <p:spPr> of the shape matching shapeName
 * within the given slide XML string.
 *
 * Matching: find <p:cNvPr ... name="SHAPE_NAME" ...> then locate the enclosing
 * <p:sp> ... </p:sp> block and edit its <p:spPr>.
 *
 * @param {string} slideXml   – full XML text of one slide
 * @param {string} shapeName  – value of data-name / objectName
 * @param {string} reflXml    – the <a:effectLst>...</a:effectLst> snippet
 * @returns {string}          – modified XML
 */
function injectReflectionForShape(slideXml, shapeName, reflXml) {
    // Escape special regex characters in shape name
    const escapedName = shapeName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    // Find <p:cNvPr with this name (handles both name="..." and id="..." orderings)
    const cNvPrRegex = new RegExp(
        `<p:cNvPr[^>]+name="${escapedName}"[^>]*/>`,
        'g'
    );

    let match;
    let resultXml = slideXml;
    let offset = 0; // track how much the string has grown/shrunk

    // Reset regex
    cNvPrRegex.lastIndex = 0;

    while ((match = cNvPrRegex.exec(slideXml)) !== null) {
        const cNvPrPos = match.index;

        // Find the opening <p:sp> that encloses this <p:cNvPr>
        // Walk backwards from cNvPrPos
        const spOpenRegex = /<p:sp[ >]/g;
        let lastSpStart = -1;
        let spMatch;
        const searchRegion = slideXml.slice(0, cNvPrPos);
        spOpenRegex.lastIndex = 0;
        while ((spMatch = spOpenRegex.exec(searchRegion)) !== null) {
            lastSpStart = spMatch.index;
        }

        if (lastSpStart < 0) continue;

        // Find closing </p:sp> after the cNvPr position
        const spCloseIdx = slideXml.indexOf('</p:sp>', cNvPrPos);
        if (spCloseIdx < 0) continue;

        const spBlock = slideXml.slice(lastSpStart, spCloseIdx + '</p:sp>'.length);

        // ── Now edit spPr inside spBlock ──────────────────────────────────

        let newSpBlock = spBlock;

        // Case A: <a:effectLst> already exists → replace it
        if (spBlock.includes('<a:effectLst')) {
            newSpBlock = spBlock.replace(
                /<a:effectLst>[\s\S]*?<\/a:effectLst>|<a:effectLst\s*\/>/g,
                reflXml
            );
        }
        // Case B: <p:spPr> exists, no effectLst → insert before </p:spPr>
        else if (spBlock.includes('<p:spPr')) {
            newSpBlock = spBlock.replace(
                /(<\/p:spPr>)/,
                reflXml + '$1'
            );
        }
        // Case C: no <p:spPr> at all → skip (degenerate)
        else {
            continue;
        }

        if (newSpBlock !== spBlock) {
            // Apply the replacement in resultXml (accounting for accumulated offset)
            const adjustedStart = lastSpStart + offset;
            const adjustedEnd = adjustedStart + spBlock.length;

            resultXml = resultXml.slice(0, adjustedStart)
                + newSpBlock
                + resultXml.slice(adjustedEnd);

            offset += (newSpBlock.length - spBlock.length);
        }
    }

    return resultXml;
}

// ─── Exports ───────────────────────────────────────────────────────────────

module.exports = {
    clearReflectionStore,
    storeReflectionParams,
    collectReflectionsFromSlide,
    injectReflectionsIntoSlideXML,
    // Exposed for testing
    parseReflectionElement,
    buildReflectionXML
};