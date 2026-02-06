/**
 * Converts an HTML table to a PowerPoint table with all CSS properties dynamically extracted
 * and properly handles <ul>/<ol> lists inside cells as real PPT bullets/numbering.
 */

function addTableToSlide(slide, tableElement, slideContext, containerElement = null) {
    // ----------------------------
    const verticalOverlays = [];
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

    function extractBackgroundColor(element) {
        if (!element) return null;

        const inlineStyle = element.getAttribute('style') || '';
        const bgMatch = inlineStyle.match(/background-color:\s*([^;]+)/i);
        if (bgMatch?.[1]) {
            return bgMatch[1].trim();
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
            if (td?.includes('underline')) {
                o.underline = { style: 'sng' };
            }

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
        return runs.length ? runs : [{ text: '' }];
    }

    function buildCellContent(cell) {
        const firstList = cell.querySelector('ul,ol');

        if (!firstList) {
            // Single paragraph: return runs as-is
            return extractRunsFromNode(cell);
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

                const runs = extractRunsFromNode(li);
                
                const liAlign = (li.querySelector('p[style*="text-align"]')?.getAttribute('style') || '')
                    .match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();

                const paraOpts = Object.assign(
                    {},
                    bulletBase,
                    indentLevel ? { indentLevel: Math.min(5, indentLevel) } : {},
                    liAlign === 'center' ? { align: 'center' } :
                        liAlign === 'right' ? { align: 'right' } :
                            liAlign === 'left' ? { align: 'left' } : {}
                );

                paras.push({
                    text: runs.length > 0 ? runs : [{ text: ' ' }],
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
        const numValue = parseFloat(value);
        if (value.includes('px')) {
            return numValue / 72;
        } else if (value.includes('pt')) {
            return numValue / 72;
        } else if (value.includes('in')) {
            return numValue;
        } else {
            return numValue / 72;
        }
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

    let maxColumns = Infinity;
    for (let i = 0; i < rows.length; i++) {
        const htmlRow = rows[i];
        let rowColumnCount = 0;
        
        for (let j = 0; j < htmlRow.cells.length; j++) {
            const cell = htmlRow.cells[j];
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            rowColumnCount += colspan;
        }
        
        if (rowColumnCount < maxColumns) {
            maxColumns = rowColumnCount;
        }
    }

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
            rowHeight = 0.5;
        }
        rowHeights.push(rowHeight);
    }

    for (let i = 0; i < rows.length; i++) {
        const htmlRow = rows[i];
        const pptxCells = [];
        
        const rowBgColor = extractBackgroundColor(htmlRow);
        const rowTextColor = extractTextColor(htmlRow);

        let columnsUsed = 0;

        for (let j = 0; j < htmlRow.cells.length; j++) {
            const cell = htmlRow.cells[j];
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

            if (columnsUsed + colspan > maxColumns) {
                break;
            }

            // ✅ Get cellStyleAttr FIRST before using it
            const cellStyleAttr = cell.getAttribute('style') || '';

            // Base cell options
            const cellOpts = { 
                color: '000000',
                border: [
                    { pt: 1, color: 'FFFFFF' },
                    { pt: 1, color: 'FFFFFF' },
                    { pt: 1, color: 'FFFFFF' },
                    { pt: 1, color: 'FFFFFF' }
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

console.log('Cell vertical text detection:', { writeDir, transform, vert: cellOpts.vert });

            if (colspan > 1) cellOpts.colspan = colspan;
            if (rowspan > 1) cellOpts.rowspan = rowspan;
            columnsUsed += colspan;

            // Cell height
            const cellHeightMatch = cellStyleAttr.match(/height:\s*([^;]+)/i);
            if (cellHeightMatch?.[1]) {
                const heightValue = cellHeightMatch[1].trim();
                const cellHeight = parseToInches(heightValue);
                if (cellHeight > rowHeights[i]) {
                    rowHeights[i] = cellHeight;
                }
            }

            // Background color
            let bgColor = extractBackgroundColor(cell) || rowBgColor;
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
            }

            // Header cell
            const isHeaderCell = cell.tagName.toLowerCase() === 'th';
            if (isHeaderCell && !cellStyleAttr.includes('font-weight')) {
                cellOpts.bold = true;
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

            const parsePaddingToPoints = (value) => {
    if (!value) return null;
    const num = parseFloat(value);
    if (isNaN(num)) return null;
    
    // PptxGenJS expects points, converts internally: points * 12700 = EMUs
    
    if (value.includes('px')) {
        return num * 0.75; // px to points (96 DPI: 1pt = 1.333px)
    } else if (value.includes('pt')) {
        return num; // Already in points
    } else if (value.includes('in')) {
        return num * 72; // inches to points (72pt = 1in)
    } else {
        return num * 0.75; // Default: assume px
    }
};

            if (paddingAll) {
    const parts = paddingAll.trim().split(/\s+/);
    if (parts.length === 1) {
        const pt = parsePaddingToPoints(parts[0]);
        if (pt !== null) cellOpts.margin = pt;
    } else if (parts.length === 2) {
        const ptV = parsePaddingToPoints(parts[0]);
        const ptH = parsePaddingToPoints(parts[1]);
        if (ptV !== null && ptH !== null) {
            cellOpts.margin = [ptV, ptH, ptV, ptH];
        }
    } else if (parts.length === 4) {
        const margins = parts.map(p => parsePaddingToPoints(p));
        if (margins.every(m => m !== null)) {
            cellOpts.margin = margins;
        }
    }
} else if (paddingTop || paddingRight || paddingBottom || paddingLeft) {
    const ptT = parsePaddingToPoints(paddingTop) || 0;
    const ptR = parsePaddingToPoints(paddingRight) || 0;
    const ptB = parsePaddingToPoints(paddingBottom) || 0;
    const ptL = parsePaddingToPoints(paddingLeft) || 0;
    
    cellOpts.margin = [ptT, ptR, ptB, ptL];
}

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
const ta = cellStyleAttr.match(/text-align:\s*([^;]+)/i)?.[1]?.trim().toLowerCase();
let paragraphAlign = null;

if (ta === 'center') paragraphAlign = 'center';
else if (ta === 'right') paragraphAlign = 'right';
else if (ta === 'left') paragraphAlign = 'left';
else if (ta === 'justify') paragraphAlign = 'justify';

// Apply alignment to content structure
if (Array.isArray(rich) && rich.length > 0 && typeof rich[0] === 'object') {
    // Has paragraph structure (from lists or formatted content)
    rich = rich.map(item => {
        if (item && typeof item === 'object' && item.text !== undefined) {
            const existingAlign = item.options?.align;
            
            // Don't override if paragraph already has explicit alignment
            if (!existingAlign && paragraphAlign) {
                return {
                    text: item.text,
                    options: {
                        ...item.options,
                        align: paragraphAlign
                    }
                };
            }
        }
        return item;
    });
} else if (paragraphAlign) {
    // Plain runs array - wrap in paragraph with alignment
    rich = [{
        text: rich,
        options: { align: paragraphAlign }
    }];
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
    colIndex: columnsUsed - colspan,
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

    const colWArr = new Array(maxColumns).fill(tableWidth / maxColumns);

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