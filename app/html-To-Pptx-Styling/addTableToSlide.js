/**
 * Converts an HTML table to a PowerPoint table with all CSS properties dynamically extracted
 * and properly handles <ul>/<ol> lists inside cells as real PPT bullets/numbering.
 */

function addTableToSlide(slide, tableElement, slideContext, containerElement = null) {
    // ----------------------------
    const verticalOverlays = [];
    const TABLE_FONT_BOOST = 1.0;
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

    function boostTableFontSize(value) {
        const num = parseFloat(value);
        if (!Number.isFinite(num)) return value;
        return Math.round(num * TABLE_FONT_BOOST * 100) / 100;
    }

    function roundMarginInches(value) {
        return Math.round(value * 1000) / 1000;
    }

    function expandCellMarginValue(margin) {
        const topBottomBoost = 0.75 / 72;
        const leftRightBoost = 1.5 / 72;
        const uniformBoost = 1.25 / 72;

        if (Array.isArray(margin)) {
            const [top = 0, right = 0, bottom = 0, left = 0] = margin;
            return [
                roundMarginInches(top + topBottomBoost),
                roundMarginInches(right + leftRightBoost),
                roundMarginInches(bottom + topBottomBoost),
                roundMarginInches(left + leftRightBoost)
            ];
        }

        const num = parseFloat(margin);
        if (!Number.isFinite(num)) return margin;
        return roundMarginInches(num + uniformBoost);
    }

    function extractBackgroundColor(element, inlineOnly = false) {
        if (!element) return null;

        const inlineStyle = element.getAttribute('style') || '';
        const bgMatch = inlineStyle.match(/background-color:\s*([^;]+)/i);
        if (bgMatch?.[1]) {
            return bgMatch[1].trim();
        }

        if (inlineOnly) {
            return null;
        }

        if (typeof window !== 'undefined' && window.getComputedStyle) {
            const computedStyle = window.getComputedStyle(element);
            const bgColor = computedStyle.backgroundColor;
            if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
                return bgColor;
            }
        }

        return null;
    }

    function extractTextColor(element) {
        if (!element) return null;

        const inlineStyle = element.getAttribute('style') || '';
        const colorMatch = inlineStyle.match(/(?:^|;|\s)color:\s*([^;]+)/i);
        if (colorMatch?.[1]) {
            return colorMatch[1].trim();
        }

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
    function extractRunsFromNode(node, inherited = {}) {
        const runs = [];
        const pushRun = (text, opt = {}) => {
            if (!text) return;
            runs.push({ text, options: { ...opt } });
        };
        const merge = (a, b) => Object.assign({}, a, b);

        const styleFromEl = (el) => {
            const s = (el.getAttribute && el.getAttribute('style')) || '';
            const o = {};
            const ff = s.match(/font-family:\s*([^;]+)/i)?.[1]?.replace(/['"]/g, '').split(',')[0]?.trim();
            if (ff) o.fontFace = ff;

            const fs = s.match(/font-size:\s*([^;]+)/i)?.[1]?.trim();
            if (fs) {
                const num = parseFloat(fs);
                if (!isNaN(num)) o.fontSize = boostTableFontSize(num);
            }

            const color = s.match(/(?:^|;|\s)color:\s*([^;]+)/i)?.[1];
            if (color) o.color = rgbToHex(color);

            const fw = s.match(/font-weight:\s*([^;]+)/i)?.[1]?.trim();
            if (fw && (fw === 'bold' || parseInt(fw) >= 700)) o.bold = true;

            const fst = s.match(/font-style:\s*([^;]+)/i)?.[1]?.trim();
            if (fst && fst === 'italic') o.italic = true;

            const td = s.match(/text-decoration:\s*([^;]+)/i)?.[1]?.toLowerCase();
            if (td?.includes('underline')) {
                o.underline = { style: 'sng' };
            }
            // ✅ ADD LINE HEIGHT EXTRACTION            

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
            if (tag === 'u') next = merge(next, { underline: { style: 'sng' } });

            for (const c of n.childNodes) walk(c, next);
        };

        walk(node, inherited);
        return runs.length ? runs : [{ text: '', options: {} }];
    }

    function buildCellContent(cell) {
        const firstList = cell.querySelector('ul,ol');

        const hasOnlyBreaks =
            cell &&
            !cell.textContent.trim() &&
            Array.from(cell.childNodes).every(node =>
                node.nodeType === 1 ? node.tagName.toLowerCase() === 'br' : !String(node.nodeValue || '').trim()
            );

        if (hasOnlyBreaks) {
            return [{ text: '', options: {} }];
        }

        // REPLACE WITH:
        if (!firstList) {
            const paragraphs = cell.querySelectorAll('p');
            if (paragraphs.length <= 1) {
                return extractRunsFromNode(cell);
            }
            // Multiple <p> tags → preserve each as separate line with formatting
            const result = [];
            paragraphs.forEach((p, idx) => {
                const runs = extractRunsFromNode(p);
                runs.forEach(r => result.push(r));
                if (idx < paragraphs.length - 1) {
                    result.push({ text: '\n', options: { breakLine: true } });
                }
            });
            return result.length ? result : [{ text: '', options: {} }];
        }

        // Lists → build bullet/numbered paragraphs
        const paras = [];
        const emitList = (listEl, indentLevel = 0) => {
            const isOrdered = listEl.tagName.toLowerCase() === 'ol';
            const styleStr = (listEl.getAttribute('style') || '');
            const listType = styleStr.match(/list-style-type:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            const startAt = parseInt(listEl.getAttribute('start') || '1', 10);

            const bulletBase = isOrdered
                ? { bullet: { type: 'number', startAt: isNaN(startAt) ? 1 : startAt } }
                : (listType === 'square' ? { bullet: { char: '■' } }
                    : listType === 'circle' ? { bullet: { char: '○' } }
                        : listType === 'disc' ? { bullet: { char: '•' } }
                            : { bullet: true });

            for (const li of listEl.children) {
                if (li.tagName?.toLowerCase() !== 'li') continue;

                const runs = [];
                for (const child of li.childNodes) {
                    if (child.nodeType === 1) {
                        const childTag = child.tagName.toLowerCase();
                        if (childTag === 'ul' || childTag === 'ol') continue;
                    }
                    extractRunsFromNode(child).forEach(run => runs.push(run));
                }

                const liAlign = (li.querySelector('p[style*="text-align"]')?.getAttribute('style') || '')
                    .match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();

                // ✅ Extract line spacing from <p> inside <li>
                // ✅ Extract line spacing from <p> inside <li>
                const liPara = li.querySelector('p');
                let lineSpacingMultiplier = null;
                if (liPara && liPara.style.lineHeight) {
                    const lh = parseFloat(liPara.style.lineHeight);
                    if (!isNaN(lh)) {
                        lineSpacingMultiplier = lh;
                    }
                }

                const paraOpts = Object.assign(
                    {},
                    bulletBase,
                    indentLevel ? { indentLevel: Math.min(5, indentLevel) } : {},
                    lineSpacingMultiplier ? { lineSpacing: Math.round(lineSpacingMultiplier * 12) } : {}, // Convert to points
                    liAlign === 'center' ? { align: 'center' } :
                        liAlign === 'right' ? { align: 'right' } :
                            liAlign === 'left' ? { align: 'left' } : {}
                );

                paras.push({
                    text: (runs.length > 0 ? runs : [{ text: ' ' }])
                        .map(r => r.text || '')
                        .join(''),
                    options: paraOpts
                });

                const nested = li.querySelector(':scope > ul, :scope > ol');
                if (nested) emitList(nested, indentLevel + 1);
            }
        };

        emitList(firstList, 0);
        return paras.length ? paras : [{ text: (cell.textContent || '').trim(), options: {} }];
    }

    function getPreferredCellTextAlign(cell, cellStyleAttr) {
        const cellAlign = cellStyleAttr.match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
        if (cellAlign) return cellAlign;

        const paragraphAlignments = Array.from(cell.querySelectorAll('p'))
            .map(p => (p.getAttribute('style') || '').match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase())
            .filter(Boolean);

        if (paragraphAlignments.length > 0) {
            return paragraphAlignments[0];
        }

        return null;
    }

    // ----------------------------
    // Position & dimensions
    // ----------------------------
    let xPos = 0, yPos = 0, tableWidth = 5, tableHeight = 2;

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

    function parseToInches(value) {
        if (!value) return 0;
        const normalizedValue = String(value).trim().toLowerCase();
        const numValue = parseFloat(normalizedValue);
        if (normalizedValue.includes('%')) {
            return 0;
        }
        if (normalizedValue.includes('px')) {
            // PPT -> HTML emits coordinates using a 72-DPI mapping (12700 EMU per px),
            // so table positions and sizes must come back through the same scale.
            return numValue / 72;
        } else if (value.includes('pt')) {
            return numValue / 72;
        } else if (value.includes('in')) {
            return numValue;
        } else {
            return numValue / 72;
        }
    }

    function parseAbsoluteLengthToInches(value) {
        if (!value) return null;
        const normalizedValue = String(value).trim().toLowerCase();
        if (!normalizedValue || normalizedValue.includes('%') || normalizedValue === 'auto') {
            return null;
        }

        const parsed = parseToInches(normalizedValue);
        return parsed > 0 ? parsed : null;
    }

    function parseCssFontSizeToPoints(value) {
        if (!value) return null;
        const normalizedValue = String(value).trim().toLowerCase();
        const numValue = parseFloat(normalizedValue);
        if (!Number.isFinite(numValue)) return null;
        if (normalizedValue.includes('pt')) return numValue;
        if (normalizedValue.includes('px')) return Math.round(numValue * 100) / 100;
        if (normalizedValue.includes('in')) return numValue * 72;
        return Math.round(numValue * 100) / 100;
    }

    if (containerElement) {
        const containerStyles = extractStyleValues(containerElement.getAttribute('style'));
        if (containerStyles.left) xPos = parseToInches(containerStyles.left);
        if (containerStyles.top) yPos = parseToInches(containerStyles.top);
        if (containerStyles.width) tableWidth = parseToInches(containerStyles.width);
        if (containerStyles.height) tableHeight = parseToInches(containerStyles.height);
    }

    if (xPos === 0 && yPos === 0) {
        if (tableElement.style) {
            if (tableElement.style.left) xPos = parseToInches(tableElement.style.left);
            if (tableElement.style.top) yPos = parseToInches(tableElement.style.top);
            if (tableElement.style.width) tableWidth = parseToInches(tableElement.style.width);
            if (tableElement.style.height) tableHeight = parseToInches(tableElement.style.height);
        }

        if (xPos === 0 && yPos === 0) {
            const tableStyles = extractStyleValues(tableElement.getAttribute('style'));
            if (tableStyles.left) xPos = parseToInches(tableStyles.left);
            if (tableStyles.top) yPos = parseToInches(tableStyles.top);
            if (tableStyles.width) tableWidth = parseToInches(tableStyles.width);
            if (tableStyles.height) tableHeight = parseToInches(tableStyles.height);
        }
    }

    if (tableWidth <= 0) tableWidth = 5;
    if (tableHeight <= 0) tableHeight = 2;

    // ----------------------------
    // Build PPTX table rows
    // ----------------------------
    const rows = tableElement.rows;
    const pptxRows = [];

    let maxColumns = 0;
    for (let i = 0; i < rows.length; i++) {
        const htmlRow = rows[i];
        let colCursor = 0;

        for (let j = 0; j < htmlRow.cells.length; j++) {
            const cell = htmlRow.cells[j];
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
            colCursor += colspan;
        }
        if (colCursor > maxColumns) maxColumns = colCursor;
    }


    let rowHeights = [];
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowStyleAttr = row.getAttribute('style') || '';
        let rowHeight = 0;
        const rowHeightMatch = rowStyleAttr.match(/height:\s*([^;]+)/i);
        if (rowHeightMatch?.[1]) {
            const heightValue = rowHeightMatch[1].trim();
            rowHeight = parseAbsoluteLengthToInches(heightValue) || 0;
        }
        if (rowHeight === 0) {
            rowHeight = 0.5;
        }
        rowHeights.push(rowHeight);
    }

    let colWArr = [];
    const colEls = tableElement.querySelectorAll('colgroup col');

    if (colEls.length > 0) {
        let totalPct = 0;
        const pctWidths = [];

        colEls.forEach(col => {
            const style = col.getAttribute('style') || '';
            const match = style.match(/width:\s*([\d.]+)%/i);
            const pct = match ? parseFloat(match[1]) : 0;
            pctWidths.push(pct);
            totalPct += pct;
        });

        if (totalPct > 0) {
            colWArr = pctWidths.map(pct => (tableWidth * pct) / totalPct);
        }
    }

    if (!colWArr.length) {
        colWArr = new Array(maxColumns).fill(tableWidth / maxColumns);
    }

    const colOffsets = [0];
    for (let idx = 0; idx < colWArr.length; idx++) {
        colOffsets[idx + 1] = colOffsets[idx] + colWArr[idx];
    }

    const rowOffsets = [0];
    for (let idx = 0; idx < rowHeights.length; idx++) {
        rowOffsets[idx + 1] = rowOffsets[idx] + rowHeights[idx];
    }

    const dashedOverlayRegions = (() => {
        if (!containerElement?.parentElement) return [];
        const candidates = Array.from(containerElement.parentElement.querySelectorAll('.shape#roundRect, .shape[id="roundRect"]'));
        return candidates
            .filter(el => /dashed/i.test(el.getAttribute('style') || ''))
            .map(el => {
                const styleAttr = el.getAttribute('style') || '';
                const getPx = (prop) => {
                    const match = styleAttr.match(new RegExp(`${prop}:\\s*([\\d.]+)px`, 'i'));
                    return match ? parseFloat(match[1]) : null;
                };
                const left = getPx('left');
                const top = getPx('top');
                const width = getPx('width');
                const height = getPx('height');
                if ([left, top, width, height].some(v => !Number.isFinite(v))) return null;
                return {
                    leftIn: left / 72,
                    topIn: top / 72,
                    rightIn: (left + width) / 72,
                    bottomIn: (top + height) / 72
                };
            })
            .filter(Boolean);
    })();

    for (let i = 0; i < rows.length; i++) {
        const htmlRow = rows[i];
        const pptxCells = [];
        let colCursor = 0;

        const rowBgColor = extractBackgroundColor(htmlRow, true);
        const rowTextColor = extractTextColor(htmlRow);

        for (let j = 0; j < htmlRow.cells.length; j++) {
            const cell = htmlRow.cells[j];
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
            const colIndex = colCursor;
            colCursor += colspan;

            // ✅ Get cellStyleAttr FIRST before using it
            const cellStyleAttr = cell.getAttribute('style') || '';

            // Base cell options
            const cellOpts = {
                color: '000000',
                border: [
                    { pt: 0, color: 'FFFFFF' },
                    { pt: 0, color: 'FFFFFF' },
                    { pt: 0, color: 'FFFFFF' },
                    { pt: 0, color: 'FFFFFF' }
                ]
            };

            // Text rotation/vertical text
            // Text rotation/vertical text - check BOTH writing-mode AND transform
            const writeDir = cellStyleAttr.match(/writing-mode:\s*([^;]+)/i)?.[1]?.trim();
            const transform = cellStyleAttr.match(/transform:\s*rotate\(([^)]+)\)/i)?.[1]?.trim();
            const isVerticalText =
                writeDir === 'vertical-rl' ||
                writeDir === 'tb-rl' ||
                transform === '90deg' ||
                transform === '270deg' ||
                transform === '-90deg';

            // PptxGenJS valid vert values: 'eaVert', 'horz', 'mongolianVert', 'vert', 'vert270', 'wordArtVert', 'wordArtVertRtl'
            if (writeDir && (writeDir === 'vertical-rl' || writeDir === 'tb-rl' || writeDir === 'vertical-lr')) {
                if (transform && (transform === '180deg' || transform === '180')) {
                    // vertical-rl + rotate(180deg) = bottom-to-top
                    cellOpts.vert = 'vert270';
                } else {
                    // Just vertical-rl = top-to-bottom  
                    cellOpts.vert = 'vert';
                }
            } else if (transform) {
                if (transform === '90deg' || transform === '90') {
                    cellOpts.vert = 'vert';
                } else if (transform === '270deg' || transform === '270' || transform === '-90deg') {
                    cellOpts.vert = 'vert270';
                }
            }

            // ✅ CRITICAL: Set vertical property at the TOP LEVEL, not nested
            if (cellOpts.vert) {
                cellOpts.vertical = cellOpts.vert; // Some versions use 'vertical' instead of 'vert'
            }


            if (colspan > 1) cellOpts.colspan = colspan;
            if (rowspan > 1) cellOpts.rowspan = rowspan;

            // Cell height
            const cellHeightMatch = cellStyleAttr.match(/height:\s*([^;]+)/i);
            if (cellHeightMatch?.[1]) {
                const heightValue = cellHeightMatch[1].trim();
                const cellHeight = parseAbsoluteLengthToInches(heightValue) || 0;
                if (cellHeight > rowHeights[i]) {
                    rowHeights[i] = cellHeight;
                }
            }

            // Background color
            // Header cell
            const isHeaderCell = cell.tagName.toLowerCase() === 'th';
            if (isHeaderCell && !cellStyleAttr.includes('font-weight')) {
                cellOpts.bold = true;
            }

            // Background color
            let bgColor = extractBackgroundColor(cell, true) || rowBgColor;
            if (bgColor) {
                const hex = rgbToHex(bgColor);
                cellOpts.fill = { color: hex };
            }

            // Text color
            let textColor = extractTextColor(cell) || rowTextColor;
            if (textColor) {
                cellOpts.color = rgbToHex(textColor);
            } else if (bgColor) {
                cellOpts.color = isColorDark(bgColor) ? 'FFFFFF' : '000000';
            } else if (isHeaderCell) {
                cellOpts.color = 'FFFFFF';
            }



            // Vertical alignment
            // Vertical alignment
            const va = cellStyleAttr.match(/vertical-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            if (va === 'top') {
                cellOpts.valign = 'top';
            } else if (va === 'bottom') {
                cellOpts.valign = 'bottom';
            } else if (va === 'middle' || va === 'center') {
                cellOpts.valign = 'middle';
            }

            // ✅ For vertical text, also ensure proper alignment
            if (cellOpts.vert) {
                // Vertical text defaults to top if not specified
                if (!cellOpts.valign) {
                    cellOpts.valign = 'top';
                }
            }

            // Cell padding
            const paddingTop = cellStyleAttr.match(/padding-top:\s*([^;]+)/i)?.[1];
            const paddingRight = cellStyleAttr.match(/padding-right:\s*([^;]+)/i)?.[1];
            const paddingBottom = cellStyleAttr.match(/padding-bottom:\s*([^;]+)/i)?.[1];
            const paddingLeft = cellStyleAttr.match(/padding-left:\s*([^;]+)/i)?.[1];
            const paddingAll = cellStyleAttr.match(/(?:^|;)\s*padding:\s*([^;]+)/i)?.[1];

            const parsePaddingToInches = (value) => {
                if (!value) return null;
                const num = parseFloat(value);
                if (isNaN(num)) return null;

                if (value.includes('px')) {
                    return num / 72;
                } else if (value.includes('pt')) {
                    return num / 72;
                } else if (value.includes('in')) {
                    return num;
                } else {
                    return num / 72;
                }
            };

            if (paddingAll) {
                const parts = paddingAll.trim().split(/\s+/);
                if (parts.length === 1) {
                    const marginInches = parsePaddingToInches(parts[0]);
                    if (marginInches !== null) cellOpts.margin = expandCellMarginValue(marginInches);
                } else if (parts.length === 2) {
                    const marginV = parsePaddingToInches(parts[0]);
                    const marginH = parsePaddingToInches(parts[1]);
                    if (marginV !== null && marginH !== null) {
                        cellOpts.margin = expandCellMarginValue([marginV, marginH, marginV, marginH]);
                    }
                } else if (parts.length === 4) {
                    const margins = parts.map(p => parsePaddingToInches(p));
                    if (margins.every(m => m !== null)) {
                        cellOpts.margin = expandCellMarginValue(margins);
                    }
                }
            } else if (paddingTop || paddingRight || paddingBottom || paddingLeft) {
                const ptT = parsePaddingToInches(paddingTop) || 0;
                const ptR = parsePaddingToInches(paddingRight) || 0;
                const ptB = parsePaddingToInches(paddingBottom) || 0;
                const ptL = parsePaddingToInches(paddingLeft) || 0;

                cellOpts.margin = expandCellMarginValue([ptT, ptR, ptB, ptL]);
            }

            // Font size
            const fs = cellStyleAttr.match(/font-size:\s*([^;]+)/i)?.[1]?.trim();
            if (fs) {
                const fontSizePt = parseCssFontSizeToPoints(fs);
                if (fontSizePt !== null) cellOpts.fontSize = boostTableFontSize(fontSizePt);
            }

            // Font weight
            const fw = cellStyleAttr.match(/font-weight:\s*([^;]+)/i)?.[1]?.trim();
            if (fw && (fw === 'bold' || parseInt(fw) >= 700)) {
                cellOpts.bold = true;
            }

            // Font family
            const ff = cellStyleAttr.match(/font-family:\s*([^;]+)/i)?.[1]?.trim().replace(/['"]/g, '').split(',')[0];
            if (ff) cellOpts.fontFace = ff;

            // Font style (italic)
            const fst = cellStyleAttr.match(/font-style:\s*([^;]+)/i)?.[1]?.trim();
            if (fst && fst === 'italic') {
                cellOpts.italic = true;
            }

            // Text decoration (underline)
            const td = cellStyleAttr.match(/text-decoration:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
            if (td && td.includes('underline')) {
                cellOpts.underline = true;
            }

            // Borders
            const parseBorderSide = (side) => {
                const match = cellStyleAttr.match(new RegExp(`border-${side}:\\s*([^;]+)`, 'i'));
                if (!match?.[1]) return null;

                const borderValue = match[1].trim();
                if (borderValue === 'none' || borderValue === '0') {
                    return { pt: 0, color: '000000' };
                }

                const parts = borderValue.split(/\s+/);
                let bPt = 1, bColor = 'FFFFFF';

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

                return { pt: bPt, color: bColor };
            };

            const topBorder = parseBorderSide('top');
            const rightBorder = parseBorderSide('right');
            const bottomBorder = parseBorderSide('bottom');
            const leftBorder = parseBorderSide('left');

            if (topBorder) cellOpts.border[0] = topBorder;
            if (rightBorder) cellOpts.border[1] = rightBorder;
            if (bottomBorder) cellOpts.border[2] = bottomBorder;
            if (leftBorder) cellOpts.border[3] = leftBorder;

            // Content
            let rich = buildCellContent(cell);

            // ✅ CRITICAL FIX: Apply cell alignment to text runs
            // ✅ Extract text alignment from cell style
            // ✅ Extract and apply text alignment
            const ta = getPreferredCellTextAlign(cell, cellStyleAttr);
            let paragraphAlign = null;

            if (ta === 'center') paragraphAlign = 'center';
            else if (ta === 'right') paragraphAlign = 'right';
            else if (ta === 'left') paragraphAlign = 'left';
            else if (ta === 'justify') paragraphAlign = 'justify';

            const hasDashedOverlay =
                (topBorder && topBorder.pt > 0 && topBorder.color !== 'FFFFFF' && String(cellStyleAttr).includes('dashed')) ||
                /dashed/i.test(cellStyleAttr);
            const cellLeft = xPos + colOffsets[colIndex];
            const cellTop = yPos + rowOffsets[i];
            const cellRight = xPos + colOffsets[colIndex + colspan];
            const cellBottom = yPos + rowOffsets[i + rowspan];
            const insideDashedRegion = dashedOverlayRegions.some(region =>
                cellLeft >= region.leftIn - 0.02 &&
                cellRight <= region.rightIn + 0.02 &&
                cellTop >= region.topIn - 0.02 &&
                cellBottom <= region.bottomIn + 0.02
            );

            if (insideDashedRegion) {
                paragraphAlign = 'center';
                cellOpts.valign = 'middle';
            }
            if (hasDashedOverlay) {
                const currentMargin = cellOpts.margin || [0, 0, 0, 0];
                const expanded = Array.isArray(currentMargin)
                    ? [
                        roundMarginInches(currentMargin[0] + (1.5 / 72)),
                        roundMarginInches(currentMargin[1] + (1.75 / 72)),
                        roundMarginInches(currentMargin[2] + (1.5 / 72)),
                        roundMarginInches(currentMargin[3] + (1.75 / 72))
                    ]
                    : roundMarginInches(parseFloat(currentMargin) + (1.5 / 72));
                cellOpts.margin = expanded;
            }
            if (insideDashedRegion) {
                const currentMargin = cellOpts.margin || [0, 0, 0, 0];
                cellOpts.margin = Array.isArray(currentMargin)
                    ? [
                        roundMarginInches(currentMargin[0] + (1.25 / 72)),
                        roundMarginInches(currentMargin[1] + (1.25 / 72)),
                        roundMarginInches(currentMargin[2] + (1.25 / 72)),
                        roundMarginInches(currentMargin[3] + (1.25 / 72))
                    ]
                    : roundMarginInches(parseFloat(currentMargin) + (1.25 / 72));
            }



            // Check if rich is already in paragraph structure (from lists: array of {text: runs[], options})
            const isParagraphArray = Array.isArray(rich) && rich.length > 0 &&
                typeof rich[0] === 'object' && rich[0] !== null &&
                typeof rich[0].text === 'string';

            if (isParagraphArray) {
                if (paragraphAlign) {
                    rich = rich.map(item => ({
                        ...item,
                        options: { ...(item.options || {}), align: (item.options && item.options.align) || paragraphAlign }
                    }));
                }
            } else if (paragraphAlign) {
                rich = [{ text: String(rich || ''), options: { align: paragraphAlign } }];
            }

            if (isVerticalText) {
                // Keep table cell empty
                pptxCells.push({
                    text: '',
                    options: cellOpts
                });

                // Store overlay info for later
                    verticalOverlays.push({
                    text: rich,
                    rowIndex: i,
                        colIndex: j,
                    colspan,
                    rowspan,
                    cellOpts
                });

            } else {
                pptxCells.push({
                    text: rich,
                    options: cellOpts
                });
            }

        }

        if (pptxCells.length) {
            pptxRows.push(pptxCells);
        }
    }

    if (!pptxRows.length) {
        console.warn('No rows to add to PowerPoint table');
        return;
    }

    const tableOpts = {
        x: xPos,
        y: yPos,
        w: tableWidth,
        colW: colWArr,
        rowH: rowHeights
    };

    slide.addTable(pptxRows, tableOpts);
    // ---- Overlay vertical text as rotated text boxes ----
    verticalOverlays.forEach(item => {
        const { text, rowIndex, colIndex, colspan, rowspan, cellOpts } = item;

        // Calculate X position
        let x = xPos;
        for (let c = 0; c < colIndex; c++) {
            x += colWArr[c];
        }

        // Calculate Y position
        let y = yPos;
        for (let r = 0; r < rowIndex; r++) {
            y += rowHeights[r];
        }

        // Cell dimensions
        let w = 0;
        for (let c = 0; c < colspan; c++) {
            w += colWArr[colIndex + c];
        }

        let h = 0;
        for (let r = 0; r < rowspan; r++) {
            h += rowHeights[rowIndex + r];
        }

        slide.addText(text, {
            x,
            y,
            w,
            h,
            vert: cellOpts.vert === 'vert270' ? 'vert270' : 'vert',
            align: 'center',
            valign: 'middle',
            fontFace: cellOpts.fontFace,
            fontSize: cellOpts.fontSize,
            color: cellOpts.color,
            bold: cellOpts.bold,
            italic: cellOpts.italic
        });
    });

}

module.exports = {
    addTableToSlide
};
