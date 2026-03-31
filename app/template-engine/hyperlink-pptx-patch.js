/**
 * HYPERLINK PATCH (hyperlink-pptx-patch.js)
 */
function extractHyperlinkFromSpan(span) {
    let el = span.parentElement;

    // Walk up but stop at <p> or the text-box container — don't escape the paragraph
    while (el && el.tagName && el.tagName.toLowerCase() !== 'p' && el.tagName.toLowerCase() !== 'div') {
        if (el.tagName.toLowerCase() === 'a') {
            const action  = el.getAttribute('data-action')  || '';
            const rId     = el.getAttribute('data-rid')     || '';
            const tooltip = el.getAttribute('title')        || '';
            const href    = el.getAttribute('href')         || '';

            // action="ppaction://hlinksldjump" → internal slide jump
            // everything else → external URL (the rId resolves to a real URL in the PPTX rels)
            const isSlideJump = action.startsWith('ppaction://');

            if (isSlideJump) {
                return {
                    // PptxGenJS will write a generic hlinkClick; we patch action + rId after
                    url: tooltip || '#',    // use tooltip URL as the visible href fallback
                    tooltip: tooltip,
                    action: action,
                    rId: rId,
                    isSlideJump: true
                };
            }

            // External URL — href="#rId2" is a placeholder; real URL is in rels.
            // Pass the tooltip (which is the real URL) as the link target.
            if (tooltip) {
                return {
                    url: tooltip,
                    tooltip: tooltip,
                    action: action,
                    rId: rId,
                    isSlideJump: false
                };
            }

            // href is "#rId2" — we have no real URL, but we can still mark the run
            // so the PPTX patch can insert the correct <a:hlinkClick r:id="rId2">
            return {
                url: href,
                tooltip: tooltip,
                action: action,
                rId: rId,
                isSlideJump: false
            };
        }
        el = el.parentElement;
    }
    return null;
}


function extractSpanFormattingWithLineSpacing_TAIL_EXAMPLE(span, options) {
    // (existing shadow handling — already in your file)
    const shadowCss = span.getAttribute('data-shadow');
    if (shadowCss) {
        const pptxShadow = parseCssShadowToPptx(shadowCss);
        if (pptxShadow) {
            options.shadow = pptxShadow;
        }
    }

    // ── Hyperlink support ──────────────────────────────────────────────
    const hlinkData = extractHyperlinkFromSpan(span);
    if (hlinkData) {
        options.hyperlink = {
            url:     hlinkData.url     || '#',
            tooltip: hlinkData.tooltip || ''
        };

        // Store raw PPT metadata for the post-generation XML patch
        if (hlinkData.isSlideJump || hlinkData.action) {
            options._hlinkAction       = hlinkData.action;
            options._hlinkRId          = hlinkData.rId;
            options._hlinkIsSlideJump  = hlinkData.isSlideJump;
        }
    }
    // ── End hyperlink support ───────────────────────────────────────────

    return options;
}

/**
 * Post-processes a PptxGenJS-generated PPTX ArrayBuffer to fix slide-jump
 * hyperlinks that PptxGenJS cannot express through its public API.
 *
 * @param {ArrayBuffer} pptxBuffer  Output of pptxGenJS.write('arraybuffer')
 * @param {Array}       slideHlinks Array of { slideIndex, rId, action, tooltip } objects
 *                                  collected while building each slide
 * @returns {Promise<ArrayBuffer>}  Patched PPTX buffer
 */
async function patchSlideJumpHyperlinks(pptxBuffer, slideHlinks) {
    // Requires JSZip — already available in your project (PptxGenJS depends on it)
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(pptxBuffer);

    for (const hlink of slideHlinks) {
        const slideFile = `ppt/slides/slide${hlink.slideIndex + 1}.xml`;
        let xml = await zip.file(slideFile).async('string');


        if (hlink.tooltip) {
            // Find the relationship file for this slide
            const relsFile = `ppt/slides/_rels/slide${hlink.slideIndex + 1}.xml.rels`;
            let relsXml = await zip.file(relsFile).async('string');

            // Find the auto-generated rId that points to hlink.tooltip URL
            const rIdMatch = relsXml.match(
                new RegExp(`Id="(rId[^"]+)"[^/]*Target="${escapeRegex(hlink.tooltip)}"`)
            );

            if (rIdMatch) {
                const autoRId = rIdMatch[1];
                // Add action attribute to the hlinkClick element
                xml = xml.replace(
                    new RegExp(`(<a:hlinkClick[^>]*r:id="${autoRId}"[^/]*)(/>|>)`),
                    (match, opening, closing) => {
                        if (!opening.includes('action=')) {
                            opening += ` action="${hlink.action}"`;
                        }
                        if (hlink.tooltip && !opening.includes('tooltip=')) {
                            opening += ` tooltip="${hlink.tooltip}"`;
                        }
                        return opening + closing;
                    }
                );

                zip.file(slideFile, xml);
            }
        }
    }

    return zip.generateAsync({ type: 'arraybuffer' });
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


module.exports = {
    extractHyperlinkFromSpan,
    patchSlideJumpHyperlinks
};