/**
 * Converts an HTML table to a PowerPoint table with all CSS properties dynamically extracted
 * and properly handles <ul>/<ol> lists inside cells as real PPT bullets/numbering.
 */
function addTableToSlide(slide, tableElement, slideContext, containerElement = null) {
    // ----------------------------
    // Color utilities
    // ----------------------------
    function rgbToHex(rgb) {
        if (!rgb) return '000000';
        rgb = rgb.trim();

        if (rgb.startsWith('#')) {
            const hex = rgb.replace('#', '').toUpperCase();
            if (hex.length === 3) return hex.split('').map(c => c + c).join('');
            return hex;
        }

        const namedColors = {
            white: 'FFFFFF', black: '000000', red: 'FF0000', green: '008000',
            blue: '0000FF', yellow: 'FFFF00', purple: '800080', gray: '808080',
            grey: '808080', orange: 'FFA500', pink: 'FFC0CB', brown: 'A52A2A',
            cyan: '00FFFF', magenta: 'FF00FF', lime: '00FF00', navy: '000080',
            maroon: '800000', olive: '808000', teal: '008080', silver: 'C0C0C0',
            gold: 'FFD700', violet: 'EE82EE', indigo: '4B0082', crimson: 'DC143C'
        };
        const lower = rgb.toLowerCase();
        if (namedColors[lower]) return namedColors[lower];

        const m1 = rgb.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (m1) {
            const [r, g, b] = [m1[1], m1[2], m1[3]].map(n => Math.min(255, Math.max(0, parseInt(n))));
            const toHex = n => n.toString(16).toUpperCase().padStart(2, '0');
            return toHex(r) + toHex(g) + toHex(b);
        }

        const m2 = rgb.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([0-9.]+)\s*\)/i);
        if (m2) {
            const [r, g, b] = [m2[1], m2[2], m2[3]].map(n => Math.min(255, Math.max(0, parseInt(n))));
            const toHex = n => n.toString(16).toUpperCase().padStart(2, '0');
            return toHex(r) + toHex(g) + toHex(b);
        }

        return '000000';
    }

    function isColorDark(color) {
        if (!color) return false;
        let r, g, b;
        let c = color.trim();

        if (c.startsWith('#')) {
            c = c.replace('#', '');
            if (c.length === 3) c = c.split('').map(x => x + x).join('');
            r = parseInt(c.substr(0, 2), 16);
            g = parseInt(c.substr(2, 2), 16);
            b = parseInt(c.substr(4, 2), 16);
        } else {
            const m = c.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
            if (m) {
                r = parseInt(m[1]); g = parseInt(m[2]); b = parseInt(m[3]);
            } else {
                const low = c.toLowerCase();
                if (['white', 'yellow', 'cyan', 'lime', 'silver'].includes(low)) return false;
                if (['black', 'navy', 'maroon', 'green', 'purple', 'olive', 'teal'].includes(low)) return true;
                return false;
            }
        }
        const L = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        return L < 0.5;
    }

    // ----------------------------
    // NEW: Function to extract background color from element
    // ----------------------------
    function extractBackgroundColor(element) {
        if (!element) return null;

        // Check inline style first
        const inlineStyle = element.getAttribute('style') || '';
        const bgMatch = inlineStyle.match(/background-color:\s*([^;]+)/i);
        if (bgMatch?.[1]) {
            return bgMatch[1].trim();
        }

        // Check computed style if available (for CSS classes)
        if (typeof window !== 'undefined' && window.getComputedStyle) {
            const computedStyle = window.getComputedStyle(element);
            const bgColor = computedStyle.backgroundColor;
            if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
                return bgColor;
            }
        }

        return null;
    }

    // ----------------------------
    // NEW: Function to extract text color from element
    // ----------------------------
    function extractTextColor(element) {
        if (!element) return null;

        // Check inline style first
        const inlineStyle = element.getAttribute('style') || '';
        const colorMatch = inlineStyle.match(/(?:^|;|\s)color:\s*([^;]+)/i);
        if (colorMatch?.[1]) {
            return colorMatch[1].trim();
        }

        // Check computed style if available (for CSS classes)
        if (typeof window !== 'undefined' && window.getComputedStyle) {
            const computedStyle = window.getComputedStyle(element);
            const textColor = computedStyle.color;
            if (textColor && textColor !== 'rgba(0, 0, 0, 0)') {
                return textColor;
            }
        }

        return null;
    }

    // ----------------------------
    // Rich text + list helpers
    // ----------------------------
    function listStyleToBulletOpts(listEl) {
        const isOrdered = listEl.tagName.toLowerCase() === 'ol';
        const styleStr = (listEl.getAttribute('style') || '');
        const listStyle = styleStr.match(/list-style-type:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
        const startAt = parseInt(listEl.getAttribute('start') || '1', 10);

        if (isOrdered) {
            return { bullet: { type: 'number', startAt: isNaN(startAt) ? 1 : startAt } };
        } else {
            if (listStyle === 'square') return { bullet: { char: '■' } };
            if (listStyle === 'circle') return { bullet: { char: '○' } };
            if (listStyle === 'disc') return { bullet: { char: '•' } };
            return { bullet: true }; // default bullet
        }
    }

    // Flatten a DOM node into pptx "runs" (array of IText) preserving inline styles (b/i/u, color, size, font)
    function extractRunsFromNode(node, inherited = {}) {
        const runs = [];
        const pushRun = (text, opt = {}) => {
            if (!text) return;
            runs.push({ text, options: opt });
        };
        const merge = (a, b) => Object.assign({}, a, b);

        const styleFromEl = (el) => {
            const s = (el.getAttribute && el.getAttribute('style')) || '';
            const o = {};
            const ff = s.match(/font-family:\s*([^;]+)/i)?.[1]?.replace(/['"]/g, '').split(',')[0]?.trim();
            if (ff) o.fontFace = ff;

            const fs = s.match(/font-size:\s*([^;]+)/i)?.[1]?.trim();
            if (fs) {
                // Keep numeric value exactly as authored (e.g., 18px -> 18, 10.5px -> 10.5, 16pt -> 16)
                const num = parseFloat(fs);
                if (!isNaN(num)) o.fontSize = num;
            }

            const color = s.match(/(?:^|;|\s)color:\s*([^;]+)/i)?.[1];
            if (color) o.color = rgbToHex(color);

            const fw = s.match(/font-weight:\s*([^;]+)/i)?.[1]?.trim();
            if (fw && (fw === 'bold' || parseInt(fw) >= 700)) o.bold = true;

            const fst = s.match(/font-style:\s*([^;]+)/i)?.[1]?.trim();
            if (fst && fst === 'italic') o.italic = true;

            const td = s.match(/text-decoration:\s*([^;]+)/i)?.[1]?.toLowerCase();
            if (td?.includes('underline')) o.underline = true;

            return o;
        };

        const walk = (n, inh) => {
            if (!n) return;
            if (n.nodeType === 3) {
                const txt = n.nodeValue.replace(/\s+/g, ' ').replace(/^\s|\s$/g, '');
                if (txt) pushRun(txt, inh);
                return;
            }
            if (n.nodeType !== 1) return;

            const tag = n.tagName.toLowerCase();
            let next = merge(inh, styleFromEl(n));
            if (tag === 'b' || tag === 'strong') next = merge(next, { bold: true });
            if (tag === 'i' || tag === 'em') next = merge(next, { italic: true });
            if (tag === 'u') next = merge(next, { underline: true });

            for (const c of n.childNodes) walk(c, next);
        };

        walk(node, inherited);
        return runs.length ? runs : [{ text: '' }];
    }

    // Build content for a table cell:
    //  - If it contains a list: return array of IText entries (each li is a bullet/number paragraph)
    //  - Otherwise: return array of runs (IText[]) representing a single paragraph (no bullets)
    function buildCellContent(cell) {
        const firstList = cell.querySelector('ul,ol');

        // Helper: pick a dominant style from runs (first non-empty run wins)
        const dominantFromRuns = (runs) => {
            for (const r of runs) {
                if (r && r.text && r.text.trim()) {
                    return r.options || {};
                }
            }
            return {};
        };

        if (!firstList) {
            // Single paragraph: keep rich runs (pptxgen accepts IText[] in table cells)
            return extractRunsFromNode(cell);
        }

        // Lists → build bullet/numbered paragraphs with text:string for compatibility
        const paras = [];
        const emitList = (listEl, indentLevel = 0) => {
            const isOrdered = listEl.tagName.toLowerCase() === 'ol';
            const styleStr = (listEl.getAttribute('style') || '');
            const listType = styleStr.match(/list-style-type:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            const startAt = parseInt(listEl.getAttribute('start') || '1', 10);

            // base bullet options (no static chars unless HTML says so)
            const bulletBase = isOrdered
                ? { bullet: { type: 'number', startAt: isNaN(startAt) ? 1 : startAt } }
                : (listType === 'square' ? { bullet: { char: '■' } }
                    : listType === 'circle' ? { bullet: { char: '○' } }
                        : listType === 'disc' ? { bullet: { char: '•' } }
                            : { bullet: true });

            for (const li of listEl.children) {
                if (li.tagName?.toLowerCase() !== 'li') continue;

                const runs = extractRunsFromNode(li);
                const paragraphText = runs.map(r => r.text).join(''); // <- ensure string, avoids [object][object]
                const domStyle = dominantFromRuns(runs);

                // paragraph-level align (from inner <p>)
                const liAlign = (li.querySelector('p[style*="text-align"]')?.getAttribute('style') || '')
                    .match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();

                const paraOpts = Object.assign(
                    {},
                    bulletBase,
                    indentLevel ? { indentLevel: Math.min(5, indentLevel) } : {},
                    liAlign === 'center' ? { align: 'center' } :
                        liAlign === 'right' ? { align: 'right' } :
                            liAlign === 'left' ? { align: 'left' } : {},
                    // lift dominant inline styles to paragraph level for appearance
                    ['fontFace', 'fontSize', 'color', 'bold', 'italic', 'underline']
                        .reduce((acc, k) => (domStyle[k] !== undefined ? (acc[k] = domStyle[k], acc) : acc), {})
                );

                paras.push({
                    text: paragraphText || ' ',
                    options: paraOpts
                });

                const nested = li.querySelector(':scope > ul, :scope > ol');
                if (nested) emitList(nested, indentLevel + 1);
            }
        };

        emitList(firstList, 0);
        return paras.length ? paras : [{ text: (cell.textContent || '').trim() }];
    }

    // ----------------------------
    // Position & dimensions (values are in "px" based on your 72-DPI flow)
    // ----------------------------
    let xPos = 0, yPos = 0, tableWidth = 5, tableHeight = 2;

    // Helper function to extract positioning from style attribute string
    function extractStyleValues(styleAttr) {
        if (!styleAttr) return {};

        const values = {};
        const leftMatch = styleAttr.match(/left:\s*([^;]+)/i);
        const topMatch = styleAttr.match(/top:\s*([^;]+)/i);
        const widthMatch = styleAttr.match(/width:\s*([^;]+)/i);
        const heightMatch = styleAttr.match(/height:\s*([^;]+)/i);

        if (leftMatch?.[1]) values.left = leftMatch[1].trim();
        if (topMatch?.[1]) values.top = topMatch[1].trim();
        if (widthMatch?.[1]) values.width = widthMatch[1].trim();
        if (heightMatch?.[1]) values.height = heightMatch[1].trim();

        return values;
    }

    // Helper function to parse dimension value and convert to inches
    function parseToInches(value) {
        if (!value) return 0;
        const numValue = parseFloat(value);
        if (value.includes('px')) {
            return numValue / 72; // Convert px to inches (72 DPI)
        } else if (value.includes('pt')) {
            return numValue / 72; // Convert pt to inches
        } else if (value.includes('in')) {
            return numValue; // Already in inches
        } else {
            return numValue / 72; // Default to px conversion
        }
    }

    // Priority 1: Use container element positioning if provided
    if (containerElement) {
        const containerStyles = extractStyleValues(containerElement.getAttribute('style'));

        if (containerStyles.left) xPos = parseToInches(containerStyles.left);
        if (containerStyles.top) yPos = parseToInches(containerStyles.top);
        if (containerStyles.width) tableWidth = parseToInches(containerStyles.width);
        if (containerStyles.height) tableHeight = parseToInches(containerStyles.height);

    }

    // Priority 2: Fallback to table element positioning if container values not found
    if (xPos === 0 && yPos === 0) {
        // Try table.style properties first
        if (tableElement.style) {
            if (tableElement.style.left) xPos = parseToInches(tableElement.style.left);
            if (tableElement.style.top) yPos = parseToInches(tableElement.style.top);
            if (tableElement.style.width) tableWidth = parseToInches(tableElement.style.width);
            if (tableElement.style.height) tableHeight = parseToInches(tableElement.style.height);
        }

        // If still no positioning, try style attribute
        if (xPos === 0 && yPos === 0) {
            const tableStyles = extractStyleValues(tableElement.getAttribute('style'));

            if (tableStyles.left) xPos = parseToInches(tableStyles.left);
            if (tableStyles.top) yPos = parseToInches(tableStyles.top);
            if (tableStyles.width) tableWidth = parseToInches(tableStyles.width);
            if (tableStyles.height) tableHeight = parseToInches(tableStyles.height);
        }

    }

    // Ensure minimum dimensions
    if (tableWidth <= 0) tableWidth = 5;
    if (tableHeight <= 0) tableHeight = 2;

    // ----------------------------
    // Build PPTX table rows with height extraction and dynamic background colors
    // ----------------------------
    const rows = tableElement.rows;
    const pptxRows = [];

    // Extract row heights from HTML tr elements - STEP 1
    let rowHeights = [];
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowStyleAttr = row.getAttribute('style') || '';

        let rowHeight = 0;
        const rowHeightMatch = rowStyleAttr.match(/height:\s*([^;]+)/i);
        if (rowHeightMatch?.[1]) {
            const heightValue = rowHeightMatch[1].trim();
            rowHeight = parseToInches(heightValue);
        }

        if (rowHeight === 0) {
            rowHeight = 0.5; // Default
        }

        rowHeights.push(rowHeight);
    }

    // Process each row
    for (let i = 0; i < rows.length; i++) {
        const htmlRow = rows[i];
        const pptxCells = [];

        // NEW: Extract row-level background color
        const rowBgColor = extractBackgroundColor(htmlRow);
        const rowTextColor = extractTextColor(htmlRow);

        for (let j = 0; j < htmlRow.cells.length; j++) {
            const cell = htmlRow.cells[j];

            // Base cell options
            const cellOpts = { color: '000000' };

            // Inline style attribute
            const cellStyleAttr = cell.getAttribute('style') || '';

            // STEP 2: Extract cell height and update row height if larger
            const cellHeightMatch = cellStyleAttr.match(/height:\s*([^;]+)/i);
            if (cellHeightMatch?.[1]) {
                const heightValue = cellHeightMatch[1].trim();
                const cellHeight = parseToInches(heightValue);

                // Update row height if cell height is larger
                if (cellHeight > rowHeights[i]) {
                    rowHeights[i] = cellHeight;
                }
            }

            // NEW: Dynamic background color extraction (priority: cell > row)
            let bgColor = extractBackgroundColor(cell) || rowBgColor;
            if (bgColor) {
                const hex = rgbToHex(bgColor);
                cellOpts.fill = { color: hex };            // << must be an object for tables
            }

            // NEW: Dynamic text color extraction (priority: cell > row)
            let textColor = extractTextColor(cell) || rowTextColor;
            if (textColor) {
                cellOpts.color = rgbToHex(textColor);
            } else if (bgColor) {
                // Auto-determine text color based on background if no explicit text color
                cellOpts.color = isColorDark(bgColor) ? 'FFFFFF' : '000000';
            }

            // Check if it's a header cell for additional styling
            const isHeaderCell = cell.tagName.toLowerCase() === 'th';

            // Apply header-specific styling only if not already styled
            if (isHeaderCell) {
                // Only apply bold if not already set
                if (!cellStyleAttr.includes('font-weight')) {
                    cellOpts.bold = true;
                }

                // Only center-align if no alignment is specified
                if (!cellStyleAttr.includes('text-align')) {
                    const pAlign = (cell.querySelector('p[style*="text-align"]')?.getAttribute('style') || '')
                        .match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
                    if (!pAlign) {
                        cellOpts.align = 'center';
                    }
                }
            }

            // Continue with other style extractions (alignments, fonts, etc.)
            const ta = cellStyleAttr.match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            if (ta) {
                if (ta === 'center') cellOpts.align = 'center';
                else if (ta === 'right') cellOpts.align = 'right';
                else if (ta === 'left') cellOpts.align = 'left';
            } else if (!cellOpts.align) {
                const pAlign = (cell.querySelector('p[style*="text-align"]')?.getAttribute('style') || '')
                    .match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
                if (pAlign === 'center') cellOpts.align = 'center';
                else if (pAlign === 'right') cellOpts.align = 'right';
                else if (pAlign === 'left') cellOpts.align = 'left';
            }

            const va = cellStyleAttr.match(/vertical-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            if (va === 'top') cellOpts.valign = 'top';
            else if (va === 'bottom') cellOpts.valign = 'bottom';
            else if (va === 'middle') cellOpts.valign = 'middle';

            // Font size
            const fs = cellStyleAttr.match(/font-size:\s*([^;]+)/i)?.[1]?.trim();
            if (fs) {
                const num = parseFloat(fs);
                if (!isNaN(num)) cellOpts.fontSize = num;
            }

            // Font weight
            const fw = cellStyleAttr.match(/font-weight:\s*([^;]+)/i)?.[1]?.trim();
            if (fw && (fw === 'bold' || parseInt(fw) >= 700)) {
                cellOpts.bold = true;
            }

            // Font family
            const ff = cellStyleAttr.match(/font-family:\s*([^;]+)/i)?.[1]?.trim().replace(/['"]/g, '').split(',')[0];
            if (ff) cellOpts.fontFace = ff;

            // Border (shorthand)
            const borderMatch = cellStyleAttr.match(/border:\s*([^;]+)/i);
            if (borderMatch?.[1]) {
                const parts = borderMatch[1].trim().split(/\s+/);
                let bPt = 1, bColor = '000000';
                parts.forEach(p => {
                    const w = p.match(/^(\d+(?:\.\d+)?)(px|pt)?$/i);
                    if (w) {
                        let width = parseFloat(w[1]);
                        const unit = (w[2] || 'px').toLowerCase();
                        if (unit === 'px') width *= 0.75;
                        bPt = Math.max(0.25, width);
                    } else if (/^(#[0-9a-f]{3,6}|rgb|rgba|[a-z]+)$/i.test(p)) {
                        bColor = rgbToHex(p);
                    }
                });
                cellOpts.border = [
                    { pt: bPt, color: bColor },
                    { pt: bPt, color: bColor },
                    { pt: bPt, color: bColor },
                    { pt: bPt, color: bColor },
                ];
            }

            // Wrapping hints
            if (cellStyleAttr.includes('word-wrap: break-word') || cellStyleAttr.includes('overflow-wrap: break-word')) {
                cellOpts.breakLine = true;
            }

            // Build cell content (using existing buildCellContent function)
            const rich = buildCellContent(cell);

            pptxCells.push({
                text: rich,
                options: cellOpts
            });
        }

        if (pptxCells.length) pptxRows.push(pptxCells);
    }

    if (!pptxRows.length) {
        console.warn('No rows to add to PowerPoint table');
        return;
    }

    // Table options with extracted row heights
    const colCount = pptxRows[0].length;
    const tableOpts = {
        x: xPos,
        y: yPos,
        w: tableWidth,
        colW: new Array(colCount).fill(tableWidth / colCount),
        rowH: rowHeights, // Use extracted row/cell heights
    };

    slide.addTable(pptxRows, tableOpts);
}

module.exports = {
    addTableToSlide
};