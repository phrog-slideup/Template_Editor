const IDENTITY_MATRIX = [1, 0, 0, 1, 0, 0];

function multiplyMatrices(a, b) {
    return [
        a[0] * b[0] + a[2] * b[1],
        a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3],
        a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4],
        a[1] * b[4] + a[3] * b[5] + a[5]
    ];
}

function applyMatrixToPoint(x, y, matrix) {
    return {
        x: matrix[0] * x + matrix[2] * y + matrix[4],
        y: matrix[1] * x + matrix[3] * y + matrix[5]
    };
}

function parseTransform(transformText) {
    if (!transformText) return IDENTITY_MATRIX;

    const transformRegex = /([a-zA-Z]+)\(([^)]*)\)/g;
    let result = IDENTITY_MATRIX;
    let match;

    while ((match = transformRegex.exec(transformText))) {
        const type = match[1].toLowerCase();
        const nums = match[2]
            .trim()
            .split(/[\s,]+/)
            .map(Number)
            .filter(num => Number.isFinite(num));

        let current = IDENTITY_MATRIX;

        if (type === 'matrix' && nums.length >= 6) {
            current = nums.slice(0, 6);
        } else if (type === 'translate') {
            current = [1, 0, 0, 1, nums[0] || 0, nums[1] || 0];
        } else if (type === 'scale') {
            const sx = nums[0] ?? 1;
            const sy = nums[1] ?? sx;
            current = [sx, 0, 0, sy, 0, 0];
        } else if (type === 'rotate' && nums.length >= 1) {
            const angle = (nums[0] * Math.PI) / 180;
            const cos = Math.cos(angle);
            const sin = Math.sin(angle);
            if (nums.length >= 3) {
                const [cx, cy] = [nums[1], nums[2]];
                current = multiplyMatrices(
                    multiplyMatrices([1, 0, 0, 1, cx, cy], [cos, sin, -sin, cos, 0, 0]),
                    [1, 0, 0, 1, -cx, -cy]
                );
            } else {
                current = [cos, sin, -sin, cos, 0, 0];
            }
        }

        result = multiplyMatrices(result, current);
    }

    return result;
}

function getNodeTransformMatrix(node, svgRoot) {
    let matrix = IDENTITY_MATRIX;
    let current = node;

    while (current && current !== svgRoot) {
        if (current.getAttribute) {
            matrix = multiplyMatrices(parseTransform(current.getAttribute('transform')), matrix);
        }
        current = current.parentElement;
    }

    return matrix;
}

function clonePoint(point) {
    const cloned = { ...point };
    if (point.curve) cloned.curve = { ...point.curve };
    return cloned;
}

function transformPoints(points, matrix) {
    return points.map(point => {
        if (point.close) return { close: true };

        const next = clonePoint(point);
        const end = applyMatrixToPoint(next.x, next.y, matrix);
        next.x = end.x;
        next.y = end.y;

        if (next.curve?.type === 'cubic') {
            const cp1 = applyMatrixToPoint(next.curve.x1, next.curve.y1, matrix);
            const cp2 = applyMatrixToPoint(next.curve.x2, next.curve.y2, matrix);
            next.curve.x1 = cp1.x;
            next.curve.y1 = cp1.y;
            next.curve.x2 = cp2.x;
            next.curve.y2 = cp2.y;
        } else if (next.curve?.type === 'quadratic') {
            const cp = applyMatrixToPoint(next.curve.x1, next.curve.y1, matrix);
            next.curve.x1 = cp.x;
            next.curve.y1 = cp.y;
        }

        return next;
    });
}

function computePointsBounds(points) {
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;

    const mark = (x, y) => {
        if (!Number.isFinite(x) || !Number.isFinite(y)) return;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
    };

    points.forEach(point => {
        if (point.close) return;
        mark(point.x, point.y);
        if (point.curve?.type === 'cubic') {
            mark(point.curve.x1, point.curve.y1);
            mark(point.curve.x2, point.curve.y2);
        } else if (point.curve?.type === 'quadratic') {
            mark(point.curve.x1, point.curve.y1);
        }
    });

    if (!Number.isFinite(minX) || !Number.isFinite(minY)) {
        return { minX: 0, minY: 0, maxX: 1, maxY: 1 };
    }

    if (maxX === minX) maxX = minX + 1;
    if (maxY === minY) maxY = minY + 1;

    return { minX, minY, maxX, maxY };
}

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function normalizePoints(points, bounds, shapeWidth, shapeHeight) {
    const width = Math.max(bounds.maxX - bounds.minX, 1);
    const height = Math.max(bounds.maxY - bounds.minY, 1);
    const normX = value => clamp(((value - bounds.minX) / width) * shapeWidth, 0, shapeWidth);
    const normY = value => clamp(((value - bounds.minY) / height) * shapeHeight, 0, shapeHeight);

    return points.map(point => {
        if (point.close) return { close: true };

        const next = clonePoint(point);
        next.x = normX(next.x);
        next.y = normY(next.y);

        if (next.curve?.type === 'cubic') {
            next.curve.x1 = normX(next.curve.x1);
            next.curve.y1 = normY(next.curve.y1);
            next.curve.x2 = normX(next.curve.x2);
            next.curve.y2 = normY(next.curve.y2);
        } else if (next.curve?.type === 'quadratic') {
            next.curve.x1 = normX(next.curve.x1);
            next.curve.y1 = normY(next.curve.y1);
        }

        return next;
    });
}

function parsePathDataToRawPoints(pathData) {
    if (!pathData) return [];

    const cleanData = pathData.replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
    const cmds = cleanData.match(/[MLHVCSQTAZmlhvcsqtaz][^MLHVCSQTAZmlhvcsqtaz]*/g);
    if (!cmds || cmds.length === 0) return [];

    const points = [];
    let curX = 0;
    let curY = 0;
    let startX = 0;
    let startY = 0;
    let lastCubicCPX = null;
    let lastCubicCPY = null;
    let lastQuadCPX = null;
    let lastQuadCPY = null;
    let hasMoveTo = false;

    cmds.forEach(cmd => {
        const type = cmd[0].toUpperCase();
        const rel = cmd[0] === cmd[0].toLowerCase() && type !== 'Z';
        const nums = cmd.slice(1).trim().split(/[\s,]+/).map(parseFloat).filter(n => !isNaN(n));
        const ax = value => rel ? curX + value : value;
        const ay = value => rel ? curY + value : value;

        switch (type) {
            case 'M':
                curX = ax(nums[0] || 0);
                curY = ay(nums[1] || 0);
                startX = curX;
                startY = curY;
                points.push({ x: curX, y: curY, moveTo: true });
                hasMoveTo = true;
                for (let i = 2; i + 1 < nums.length; i += 2) {
                    curX = ax(nums[i]);
                    curY = ay(nums[i + 1]);
                    points.push({ x: curX, y: curY });
                }
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
            case 'L':
                for (let i = 0; i + 1 < nums.length; i += 2) {
                    curX = ax(nums[i]);
                    curY = ay(nums[i + 1]);
                    points.push({ x: curX, y: curY });
                }
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
            case 'H':
                nums.forEach(value => {
                    curX = rel ? curX + value : value;
                    points.push({ x: curX, y: curY });
                });
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
            case 'V':
                nums.forEach(value => {
                    curY = rel ? curY + value : value;
                    points.push({ x: curX, y: curY });
                });
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
            case 'C':
                for (let i = 0; i + 5 < nums.length; i += 6) {
                    const x1 = ax(nums[i]);
                    const y1 = ay(nums[i + 1]);
                    const x2 = ax(nums[i + 2]);
                    const y2 = ay(nums[i + 3]);
                    curX = ax(nums[i + 4]);
                    curY = ay(nums[i + 5]);
                    lastCubicCPX = x2;
                    lastCubicCPY = y2;
                    points.push({ x: curX, y: curY, curve: { type: 'cubic', x1, y1, x2, y2 } });
                }
                lastQuadCPX = lastQuadCPY = null;
                break;
            case 'S':
                for (let i = 0; i + 3 < nums.length; i += 4) {
                    const x1 = lastCubicCPX !== null ? 2 * curX - lastCubicCPX : curX;
                    const y1 = lastCubicCPY !== null ? 2 * curY - lastCubicCPY : curY;
                    const x2 = ax(nums[i]);
                    const y2 = ay(nums[i + 1]);
                    curX = ax(nums[i + 2]);
                    curY = ay(nums[i + 3]);
                    lastCubicCPX = x2;
                    lastCubicCPY = y2;
                    points.push({ x: curX, y: curY, curve: { type: 'cubic', x1, y1, x2, y2 } });
                }
                lastQuadCPX = lastQuadCPY = null;
                break;
            case 'Q':
                for (let i = 0; i + 3 < nums.length; i += 4) {
                    const x1 = ax(nums[i]);
                    const y1 = ay(nums[i + 1]);
                    curX = ax(nums[i + 2]);
                    curY = ay(nums[i + 3]);
                    lastQuadCPX = x1;
                    lastQuadCPY = y1;
                    points.push({ x: curX, y: curY, curve: { type: 'quadratic', x1, y1 } });
                }
                lastCubicCPX = lastCubicCPY = null;
                break;
            case 'T':
                for (let i = 0; i + 1 < nums.length; i += 2) {
                    const x1 = lastQuadCPX !== null ? 2 * curX - lastQuadCPX : curX;
                    const y1 = lastQuadCPY !== null ? 2 * curY - lastQuadCPY : curY;
                    curX = ax(nums[i]);
                    curY = ay(nums[i + 1]);
                    lastQuadCPX = x1;
                    lastQuadCPY = y1;
                    points.push({ x: curX, y: curY, curve: { type: 'quadratic', x1, y1 } });
                }
                lastCubicCPX = lastCubicCPY = null;
                break;
            case 'A':
                if (nums.length >= 7) {
                    curX = ax(nums[5]);
                    curY = ay(nums[6]);
                    points.push({ x: curX, y: curY });
                }
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
            case 'Z':
                if (hasMoveTo) points.push({ close: true });
                curX = startX;
                curY = startY;
                lastCubicCPX = lastCubicCPY = lastQuadCPX = lastQuadCPY = null;
                break;
        }
    });

    return points;
}


function parseViewBox(svgElement) {
    const viewBoxAttr = svgElement?.getAttribute?.('viewBox');
    if (!viewBoxAttr) return null;

    const values = viewBoxAttr
        .trim()
        .split(/[\s,]+/)
        .map(Number)
        .filter(Number.isFinite);

    if (values.length < 4) return null;

    const [minX, minY, width, height] = values;
    if (!(width > 0) || !(height > 0)) return null;

    return {
        minX,
        minY,
        maxX: minX + width,
        maxY: minY + height
    };
}

function boundsFitInside(containerBounds, candidateBounds, tolerance = 0.01) {
    if (!containerBounds || !candidateBounds) return false;

    const width = candidateBounds.maxX - candidateBounds.minX;
    const height = candidateBounds.maxY - candidateBounds.minY;
    const padX = Math.max(width * tolerance, tolerance);
    const padY = Math.max(height * tolerance, tolerance);

    return (
        candidateBounds.minX >= containerBounds.minX - padX &&
        candidateBounds.minY >= containerBounds.minY - padY &&
        candidateBounds.maxX <= containerBounds.maxX + padX &&
        candidateBounds.maxY <= containerBounds.maxY + padY
    );
}

function resolveNormalizationBounds(svgElement, drawables) {
    const overallBounds = drawables.reduce((acc, drawable) => {
        const bounds = computePointsBounds(drawable.points);
        return {
            minX: Math.min(acc.minX, bounds.minX),
            minY: Math.min(acc.minY, bounds.minY),
            maxX: Math.max(acc.maxX, bounds.maxX),
            maxY: Math.max(acc.maxY, bounds.maxY)
        };
    }, { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity });

    const viewBoxBounds = parseViewBox(svgElement);
    if (viewBoxBounds && boundsFitInside(viewBoxBounds, overallBounds)) {
        return viewBoxBounds;
    }

    return overallBounds;
}

function resolveShapePath(node) {
    const tag = node.tagName.toLowerCase();
    if (tag === 'path') return (node.getAttribute('d') || '').trim();

    if (tag === 'polygon' || tag === 'polyline') {
        const rawPoints = (node.getAttribute('points') || '').trim();
        if (!rawPoints) return '';

        const nums = rawPoints.split(/[\s,]+/).map(Number).filter(Number.isFinite);
        if (nums.length < 4) return '';

        const coords = [];
        for (let i = 0; i + 1 < nums.length; i += 2) {
            coords.push(`${nums[i]} ${nums[i + 1]}`);
        }
        return `M ${coords[0]} ${coords.slice(1).map(point => `L ${point}`).join(' ')}${tag === 'polygon' ? ' Z' : ''}`;
    }

    if (tag === 'rect') {
        const x = parseFloat(node.getAttribute('x') || '0');
        const y = parseFloat(node.getAttribute('y') || '0');
        const width = parseFloat(node.getAttribute('width') || '0');
        const height = parseFloat(node.getAttribute('height') || '0');
        if (!(width > 0 && height > 0)) return '';
        return `M ${x} ${y} L ${x + width} ${y} L ${x + width} ${y + height} L ${x} ${y + height} Z`;
    }

    if (tag === 'circle' || tag === 'ellipse') {
        const cx = parseFloat(node.getAttribute('cx') || '0');
        const cy = parseFloat(node.getAttribute('cy') || '0');
        const rx = tag === 'circle'
            ? parseFloat(node.getAttribute('r') || '0')
            : parseFloat(node.getAttribute('rx') || '0');
        const ry = tag === 'circle'
            ? parseFloat(node.getAttribute('r') || '0')
            : parseFloat(node.getAttribute('ry') || '0');
        if (!(rx > 0 && ry > 0)) return '';

        const k = 0.5522847498307936;
        return [
            `M ${cx + rx} ${cy}`,
            `C ${cx + rx} ${cy + ry * k} ${cx + rx * k} ${cy + ry} ${cx} ${cy + ry}`,
            `C ${cx - rx * k} ${cy + ry} ${cx - rx} ${cy + ry * k} ${cx - rx} ${cy}`,
            `C ${cx - rx} ${cy - ry * k} ${cx - rx * k} ${cy - ry} ${cx} ${cy - ry}`,
            `C ${cx + rx * k} ${cy - ry} ${cx + rx} ${cy - ry * k} ${cx + rx} ${cy} Z`
        ].join(' ');
    }

    return '';
}

function extractColor(colorValue) {
    if (!colorValue || colorValue === 'transparent' || colorValue === 'none') return null;

    if (colorValue.startsWith('#')) {
        const hex = colorValue.substring(1).toUpperCase();
        if (hex.length === 6) return hex;
        if (hex.length === 3) return hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        return null;
    }

    const rgbMatch = colorValue.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgbMatch) {
        const r = Math.max(0, Math.min(255, parseInt(rgbMatch[1], 10))).toString(16).padStart(2, '0');
        const g = Math.max(0, Math.min(255, parseInt(rgbMatch[2], 10))).toString(16).padStart(2, '0');
        const b = Math.max(0, Math.min(255, parseInt(rgbMatch[3], 10))).toString(16).padStart(2, '0');
        return `${r}${g}${b}`.toUpperCase();
    }

    const namedColors = {
        white: 'FFFFFF',
        black: '000000',
        red: 'FF0000',
        green: '008000',
        blue: '0000FF',
        yellow: 'FFFF00',
        cyan: '00FFFF',
        magenta: 'FF00FF',
        silver: 'C0C0C0',
        gray: '808080',
        grey: '808080',
        maroon: '800000',
        olive: '808000',
        purple: '800080',
        teal: '008080',
        navy: '000080'
    };

    return namedColors[colorValue.toLowerCase()] || null;
}

function getSvgStyleValue(node, svgElement, attributeName, cssName) {
    if (!node && !svgElement) return null;

    const nodeAttr = node?.getAttribute?.(attributeName);
    if (nodeAttr != null && nodeAttr !== '') return nodeAttr;

    const nodeStyle = node?.style?.getPropertyValue?.(cssName);
    if (nodeStyle) return nodeStyle;

    const svgAttr = svgElement?.getAttribute?.(attributeName);
    if (svgAttr != null && svgAttr !== '') return svgAttr;

    const svgStyle = svgElement?.style?.getPropertyValue?.(cssName);
    if (svgStyle) return svgStyle;

    return null;
}

function mapStrokeDasharrayToPptDashType(strokeDasharray, strokeWidth) {
    if (!strokeDasharray) return null;

    const normalizedDash = String(strokeDasharray).trim().toLowerCase();
    if (!normalizedDash || normalizedDash === 'none') return null;

    const dashParts = normalizedDash
        .split(/[\s,]+/)
        .map(Number)
        .filter(value => Number.isFinite(value) && value > 0);

    if (dashParts.length === 0) return null;

    const [dashLen = 0, gapLen = dashLen, thirdLen = 0] = dashParts;
    const safeStrokeWidth = Number.isFinite(strokeWidth) && strokeWidth > 0 ? strokeWidth : 1;

    if (dashParts.length >= 4) {
        const isDotLike = thirdLen > 0 && thirdLen <= safeStrokeWidth * 1.4;
        return isDotLike ? 'dashDotDot' : 'dashDot';
    }

    if (dashLen <= safeStrokeWidth * 1.4) return 'sysDot';
    if (dashLen >= gapLen * 1.8) return 'lgDash';
    return 'dash';
}

// ─── SVG Gradient Helpers ────────────────────────────────────────────────────

/**
 * Returns true if `node` is nested inside a <defs> element.
 * Paths inside <defs> are clip-geometry/marker templates, NOT drawable shapes.
 */
function isInsideDefs(node) {
    let current = node.parentElement;
    while (current) {
        if (current.tagName && current.tagName.toLowerCase() === 'defs') return true;
        current = current.parentElement;
    }
    return false;
}

/**
 * Safely extract the background / background-image value from a DOM element's
 * raw style attribute.
 *
 */
function extractRawBackground(element) {
    if (!element) return '';
    const styleAttr = element.getAttribute('style') || '';
    // Match "background: ..." or "background-image: ..."
    const match = styleAttr.match(/background(?:-image)?\s*:\s*([^;]+)/i);
    return match ? match[1].trim() : '';
}

/**
 * JSDOM-safe raw box-shadow extractor.
 * JSDOM's CSSOM parser silently drops box-shadow from element.style —
 * must read the raw style attribute string via getAttribute().
 */
function extractRawBoxShadow(element) {
    if (!element) return '';
    const styleAttr = element.getAttribute('style') || '';
    const match = styleAttr.match(/box-shadow\s*:\s*([^;]+)/i);
    return match ? match[1].trim() : '';
}

function parseBoxShadowForPptxgenjs(boxShadow) {
    if (!boxShadow || boxShadow === 'none') return null;
    try {
        const isInset = /\binset\b/i.test(boxShadow);

        // ── Colour + alpha ───────────────────────────────────────────────────
        let colorHex = '000000';
        let opacity = 1;
        const rgbaMatch = boxShadow.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/);
        if (rgbaMatch) {
            const r = parseInt(rgbaMatch[1], 10).toString(16).padStart(2, '0');
            const g = parseInt(rgbaMatch[2], 10).toString(16).padStart(2, '0');
            const b = parseInt(rgbaMatch[3], 10).toString(16).padStart(2, '0');
            colorHex = `${r}${g}${b}`.toUpperCase();
            if (rgbaMatch[4] !== undefined) opacity = parseFloat(rgbaMatch[4]);
        } else {
            const hexMatch = boxShadow.match(/#([0-9a-fA-F]{3,8})/);
            if (hexMatch) {
                let hex = hexMatch[1];
                if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                colorHex = hex.toUpperCase().slice(0, 6);
            }
        }

        // ── Numeric values: offset-x, offset-y, blur-radius ─────────────────
        // Strip colour and 'inset' keyword, then read remaining px numbers
        const stripped = boxShadow
            .replace(/rgba?\([^)]+\)/gi, '')
            .replace(/#[0-9a-fA-F]{3,8}/gi, '')
            .replace(/\binset\b/gi, '')
            .trim();

        const nums = stripped.match(/-?[\d.]+px/g) || [];
        const offsetXPx = nums[0] ? parseFloat(nums[0]) : 0;
        const offsetYPx = nums[1] ? parseFloat(nums[1]) : 0;
        const blurPx = nums[2] ? parseFloat(nums[2]) : 0;

        const PX_TO_PT = 0.75;
        const blurPt = Math.round(Math.abs(blurPx) * PX_TO_PT);
        const distPt = Math.round(Math.sqrt(offsetXPx ** 2 + offsetYPx ** 2) * PX_TO_PT);
        const angleDeg = Math.round(((Math.atan2(offsetYPx, offsetXPx) * 180 / Math.PI) + 360) % 360);

        return {
            type: isInset ? 'inner' : 'outer',
            color: colorHex,
            blur: blurPt,
            offset: distPt,
            angle: angleDeg,
            opacity: Number.isFinite(opacity) ? Math.max(0, Math.min(1, opacity)) : 1
        };
    } catch (_) {
        return null;
    }
}

function parseCssGradient(cssGradStr) {
    if (!cssGradStr || !cssGradStr.includes('gradient')) return null;
    try {
        const gradientMatch = cssGradStr.match(/(linear|radial)-gradient\(([\s\S]+)\)/i);
        if (!gradientMatch) return null;

        const type = gradientMatch[1].toLowerCase();
        const content = gradientMatch[2].trim();

        // Split top-level comma-separated parts (respects nested parens)
        const parts = [];
        let depth = 0, cur = '';
        for (const ch of content) {
            if (ch === '(') depth++;
            else if (ch === ')') depth--;
            if (ch === ',' && depth === 0) { parts.push(cur.trim()); cur = ''; }
            else cur += ch;
        }
        if (cur.trim()) parts.push(cur.trim());

        // Parse colour stops (skip non-colour config parts)
        const stops = [];
        for (const part of parts) {
            // rgba(r,g,b,a?) [pos%]
            const rgbaMatch = part.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)\s*([\d.]+%)?/);
            if (rgbaMatch) {
                const r = parseInt(rgbaMatch[1], 10).toString(16).padStart(2, '0');
                const g = parseInt(rgbaMatch[2], 10).toString(16).padStart(2, '0');
                const b = parseInt(rgbaMatch[3], 10).toString(16).padStart(2, '0');
                const alpha = rgbaMatch[4] !== undefined ? parseFloat(rgbaMatch[4]) : 1;
                const clampedAlpha = Math.max(0, Math.min(1, Number.isFinite(alpha) ? alpha : 1));
                const pos = rgbaMatch[5] ? parseFloat(rgbaMatch[5]) / 100 : null;
                // alphaVal: OOXML <a:alpha val="..."/> uses 1/1000ths of a percent
                // (100000 = fully opaque, 0 = fully transparent) — pre-computed here
                // for per-stop transparency, consistent with shape/shadow opacity handling.
                stops.push({ color: `${r}${g}${b}`.toUpperCase(), alpha: clampedAlpha, alphaVal: Math.round(clampedAlpha * 100000), pos });
                continue;
            }
            // #hex [pos%]
            const hexMatch = part.match(/#([0-9a-fA-F]{3,8})\s*([\d.]+%)?/);
            if (hexMatch) {
                let hex = hexMatch[1];
                if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                const pos = hexMatch[2] ? parseFloat(hexMatch[2]) / 100 : null;
                stops.push({ color: hex.toUpperCase().slice(0, 6), alpha: 1, alphaVal: 100000, pos });
            }
        }

        if (stops.length === 0) return null;

        // Fill in implied positions (spread evenly for nulls)
        stops.forEach((stop, i) => {
            if (stop.pos === null) {
                if (i === 0) stop.pos = 0;
                else if (i === stops.length - 1) stop.pos = 1;
                else stop.pos = i / (stops.length - 1);
            }
        });

        // Radial: focal point (centre by default)
        let focusX = 50, focusY = 50, radialPath = 'circle';
        if (type === 'radial') {
            const focalMatch = content.match(/(?:circle|ellipse)\s+at\s+([\w\s%]+?)(?:,|$)/i);
            if (focalMatch) {
                const mapPos = v => {
                    if (!v) return 50;
                    const lc = v.toLowerCase().trim();
                    if (lc === 'left') return 0;
                    if (lc === 'right') return 100;
                    if (lc === 'top') return 0;
                    if (lc === 'bottom') return 100;
                    if (lc === 'center') return 50;
                    if (lc.endsWith('%')) return parseFloat(lc);
                    return 50;
                };
                const pos = focalMatch[1].trim().split(/\s+/);
                focusX = mapPos(pos[0]);
                focusY = pos.length >= 2 ? mapPos(pos[1]) : focusX;
            }
            if (/ellipse/i.test(content)) radialPath = 'rect';
        }

        // Linear: angle (degrees, CSS convention)
        let angleDeg = 180; // default: top→bottom
        if (type === 'linear') {
            const angleMatch = content.match(/^([-\d.]+)deg/);
            if (angleMatch) angleDeg = parseFloat(angleMatch[1]);
            else if (/to bottom/i.test(content)) angleDeg = 180;
            else if (/to top/i.test(content)) angleDeg = 0;
            else if (/to right/i.test(content)) angleDeg = 90;
            else if (/to left/i.test(content)) angleDeg = 270;
        }

        return { type, stops, focusX, focusY, radialPath, angleDeg };
    } catch (_) {
        return null;
    }
}

function collectNodeStyles(node, svgElement) {
    return {
        fill: getSvgStyleValue(node, svgElement, 'fill', 'fill'),
        fillOpacity: getSvgStyleValue(node, svgElement, 'fill-opacity', 'fill-opacity'),
        stroke: getSvgStyleValue(node, svgElement, 'stroke', 'stroke'),
        strokeWidth: getSvgStyleValue(node, svgElement, 'stroke-width', 'stroke-width') || '0',
        strokeDasharray: getSvgStyleValue(node, svgElement, 'stroke-dasharray', 'stroke-dasharray'),
        strokeLinecap: getSvgStyleValue(node, svgElement, 'stroke-linecap', 'stroke-linecap'),
        strokeLinejoin: getSvgStyleValue(node, svgElement, 'stroke-linejoin', 'stroke-linejoin'),
        opacity: getSvgStyleValue(node, svgElement, 'opacity', 'opacity') || '1'
    };
}

/**
 * parseDropShadowFilter(svgElement)
 *
 */
function parseDropShadowFilter(svgElement) {
    if (!svgElement) return null;
    try {
        const styleAttr = svgElement.getAttribute('style') || '';
        const filterMatch = styleAttr.match(/filter\s*:\s*([^;]+)/i);
        if (!filterMatch) return null;

        const filterVal = filterMatch[1].trim();
        // Use a paren-aware regex so rgba(r,g,b,a) nested inside drop-shadow(...)

        const dsMatch = filterVal.match(/drop-shadow\(((?:[^)(]|\([^)]*\))+)\)/i);
        if (!dsMatch) return null;

        const params = dsMatch[1].trim();

        // ── Colour + alpha ────────────────────────────────────────────────────
        let colorHex = '000000';
        let opacity = 1;

        const rgbaMatch = params.match(
            /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i
        );
        if (rgbaMatch) {
            colorHex = [rgbaMatch[1], rgbaMatch[2], rgbaMatch[3]]
                .map(v => parseInt(v, 10).toString(16).padStart(2, '0'))
                .join('').toUpperCase();
            if (rgbaMatch[4] !== undefined) opacity = parseFloat(rgbaMatch[4]);
        } else {
            const hexMatch = params.match(/#([0-9a-fA-F]{3,8})/);
            if (hexMatch) {
                let hex = hexMatch[1];
                if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                colorHex = hex.toUpperCase().slice(0, 6);
            }
        }

        // ── Numeric values: offset-x  offset-y  blur-radius ──────────────────
        const stripped = params
            .replace(/rgba?\([^)]+\)/gi, '')
            .replace(/#[0-9a-fA-F]{3,8}/gi, '')
            .trim();

        const nums = stripped.match(/-?[\d.]+(?:px)?/g) || [];
        const offsetXPx = nums[0] ? parseFloat(nums[0]) : 0;
        const offsetYPx = nums[1] ? parseFloat(nums[1]) : 0;
        const blurPx = nums[2] ? Math.abs(parseFloat(nums[2])) : 0;

        // ── Convert px → points (72 DPI: 1 px = 1 pt) ────────────────────────
        const blurPt = Math.round(blurPx);
        const distPt = Math.round(Math.sqrt(offsetXPx ** 2 + offsetYPx ** 2));
        const angleDeg = Math.round(
            ((Math.atan2(offsetYPx, offsetXPx) * 180 / Math.PI) + 360) % 360
        );

        return {
            type: 'outer',
            color: colorHex,
            blur: blurPt,
            offset: distPt,
            angle: angleDeg,
            opacity: Math.max(0, Math.min(1, Number.isFinite(opacity) ? opacity : 1))
        };
    } catch (_) {
        return null;
    }
}

/**
 * parseSvgGradientElement(gradEl, svgElement)
 *
 */
function parseSvgGradientElement(gradEl, svgElement) {
    if (!gradEl) return null;

    const tagName = (gradEl.tagName || '').toLowerCase().replace(/^[a-z]+:/, '');
    const isLinear = tagName === 'lineargradient';
    const isRadial = tagName === 'radialgradient';
    if (!isLinear && !isRadial) return null;

    // ── Resolve inherited stops via href / xlink:href ─────────────────────────
    let stopSource = gradEl;
    const visited = new Set([gradEl.getAttribute('id')]);
    while (stopSource.querySelectorAll('stop').length === 0) {
        const ref = stopSource.getAttribute('href') || stopSource.getAttribute('xlink:href') || '';
        if (!ref.startsWith('#')) break;
        const refId = ref.slice(1);
        if (visited.has(refId)) break;
        visited.add(refId);
        const refEl = svgElement.querySelector(`[id="${refId}"]`);
        if (!refEl) break;
        stopSource = refEl;
    }

    const stopEls = Array.from(stopSource.querySelectorAll('stop'));
    if (stopEls.length === 0) return null;

    const stops = stopEls.map(stop => {
        const offsetRaw = stop.getAttribute('offset') || '0';
        const pos = offsetRaw.endsWith('%') ? parseFloat(offsetRaw) / 100 : parseFloat(offsetRaw);

        let color = '000000';
        let alpha = 1;

        const styleAttr = stop.getAttribute('style') || '';

        const rgbaM = styleAttr.match(
            /stop-color\s*:\s*rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i
        );
        if (rgbaM) {
            color = [rgbaM[1], rgbaM[2], rgbaM[3]]
                .map(v => parseInt(v, 10).toString(16).padStart(2, '0'))
                .join('').toUpperCase();
            if (rgbaM[4] !== undefined) alpha = parseFloat(rgbaM[4]);
        } else {
            const hexM = styleAttr.match(/stop-color\s*:\s*#([0-9a-fA-F]{3,8})/i);
            if (hexM) {
                let hex = hexM[1];
                if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                color = hex.toUpperCase().slice(0, 6);
            } else {
                const scAttr = stop.getAttribute('stop-color') || '';
                const hexA = scAttr.match(/#([0-9a-fA-F]{3,8})/);
                if (hexA) {
                    let hex = hexA[1];
                    if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                    color = hex.toUpperCase().slice(0, 6);
                } else {
                    const rgbA = scAttr.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i);
                    if (rgbA) {
                        color = [rgbA[1], rgbA[2], rgbA[3]]
                            .map(v => parseInt(v, 10).toString(16).padStart(2, '0'))
                            .join('').toUpperCase();
                        if (rgbA[4] !== undefined) alpha = parseFloat(rgbA[4]);
                    }
                }
            }
        }

        const soM = styleAttr.match(/stop-opacity\s*:\s*([\d.]+)/i);
        if (soM) alpha = parseFloat(soM[1]);
        const soAttr = stop.getAttribute('stop-opacity');
        if (soAttr && !soM) alpha = parseFloat(soAttr);

        const clampedAlpha = Math.max(0, Math.min(1, alpha));
        return {
            color,
            alpha: clampedAlpha,
            // OOXML <a:alpha val="..."/> uses 1/1000ths of a percent
            // (100000 = fully opaque, 0 = fully transparent) — pre-computed
            // here so the post-processor applies per-stop transparency correctly,
            // consistent with how shape opacity and shadow opacity are handled.
            alphaVal: Math.round(clampedAlpha * 100000),
            pos: Math.max(0, Math.min(1, Number.isFinite(pos) ? pos : 0))
        };
    });

    let angleDeg = 90;
    let focusX = 50, focusY = 50;
    const radialPath = 'circle';

    if (isLinear) {
        const parseCoord = (raw, fallback) => {
            if (raw == null) return fallback;
            const n = parseFloat(raw);
            return raw.includes('%') ? n : (n <= 1 ? n * 100 : n);
        };
        const x1 = parseCoord(gradEl.getAttribute('x1'), 0);
        const y1 = parseCoord(gradEl.getAttribute('y1'), 50);
        const x2 = parseCoord(gradEl.getAttribute('x2'), 100);
        const y2 = parseCoord(gradEl.getAttribute('y2'), 50);
        const dx = x2 - x1;
        const dy = y2 - y1;
        angleDeg = Math.round(((Math.atan2(dx, -dy) * 180 / Math.PI) + 360) % 360);
    } else if (isRadial) {
        const parseCoord = (raw, fallback) => {
            if (raw == null) return fallback;
            const n = parseFloat(raw);
            return raw.includes('%') ? n : (n <= 1 ? n * 100 : n);
        };
        focusX = parseCoord(gradEl.getAttribute('cx'), 50);
        focusY = parseCoord(gradEl.getAttribute('cy'), 50);
    }

    return { type: isLinear ? 'linear' : 'radial', stops, angleDeg, focusX, focusY, radialPath };
}

function collectSvgDrawables(svgElement) {
    const nodes = Array.from(svgElement.querySelectorAll('path, polygon, polyline, rect, circle, ellipse'));

    // ── Filter out nodes inside <defs> ──────────────────────────────────────

    const visibleNodes = nodes.filter(node => !isInsideDefs(node));

    const drawables = visibleNodes
        .map(node => {
            const pathData = resolveShapePath(node);
            if (!pathData) return null;
            const rawPoints = parsePathDataToRawPoints(pathData);
            if (rawPoints.length === 0) return null;
            return {
                node,
                points: transformPoints(rawPoints, getNodeTransformMatrix(node, svgElement)),
                styles: collectNodeStyles(node, svgElement),
                gradient: null,
                strokeGradient: null
            };
        })
        .filter(Boolean);

    // ── Resolve url(#id) gradient fill references ─────────────────────────────

    // store in drawable.gradient / drawable.strokeGradient, and set a solid
    for (const drawable of drawables) {
        const fillAttr = drawable.styles.fill || '';
        const strokeAttr = drawable.styles.stroke || '';

        const resolveGradientReference = (gradientRef) => {
            if (!gradientRef.startsWith('url(#')) return null;

            const idMatch = gradientRef.match(/url\(#([^)]+)\)/);
            if (!idMatch) return null;
            const gradId = idMatch[1];

            let gradEl = null;
            try {
                gradEl = svgElement.querySelector(`#${CSS.escape(gradId)}`);
            } catch (_) {
                gradEl = svgElement.querySelector(`[id="${gradId}"]`);
            }
            if (!gradEl) return null;

            const parsed = parseSvgGradientElement(gradEl, svgElement);
            return parsed && parsed.stops.length > 0 ? parsed : null;
        };

        const fillGradient = resolveGradientReference(fillAttr);
        if (fillGradient) {
            drawable.gradient = fillGradient;
            drawable.styles.fill = `#${fillGradient.stops[0].color}`;
        }

        const strokeGradient = resolveGradientReference(strokeAttr);
        if (strokeGradient) {
            drawable.strokeGradient = strokeGradient;
            drawable.styles.stroke = `#${strokeGradient.stops[0].color}`;
        }
    }

    // ── Handle foreignObject + clipPath gradient pattern ─────────────────────

    if (drawables.length === 0) {
        const foreignObjects = Array.from(svgElement.querySelectorAll('foreignObject'));
        for (const fo of foreignObjects) {
            // Locate the associated clipPath
            const clipPathRef = fo.getAttribute('clip-path') || '';
            const clipIdMatch = clipPathRef.match(/url\(#([^)]+)\)/);
            if (!clipIdMatch) continue;

            const clipId = clipIdMatch[1];
            const clipPathEl =
                svgElement.querySelector(`clipPath[id="${clipId}"]`) ||
                svgElement.querySelector(`[id="${clipId}"]`);
            if (!clipPathEl) continue;

            const pathEl = clipPathEl.querySelector('path, polygon, polyline, rect, circle, ellipse');
            if (!pathEl) continue;

            const pathData = resolveShapePath(pathEl);
            if (!pathData) continue;

            const rawPoints = parsePathDataToRawPoints(pathData);
            if (rawPoints.length === 0) continue;

            // ── Extract gradient via raw attribute (JSDOM-safe) ──────────────
            let gradient = null;
            let shadow = null;
            const divEl = fo.querySelector('div');
            if (divEl) {
                const bg = extractRawBackground(divEl);
                if (bg.includes('gradient')) {
                    gradient = parseCssGradient(bg);
                }


                const rawShadow = extractRawBoxShadow(divEl);
                if (rawShadow) shadow = parseBoxShadowForPptxgenjs(rawShadow);
            }


            const fallbackHex = gradient?.stops?.[0]?.color || 'CCCCCC';

            drawables.push({
                node: pathEl,
                points: transformPoints(rawPoints, getNodeTransformMatrix(pathEl, svgElement)),
                styles: {
                    fill: `#${fallbackHex}`,   // extractColor-compatible hex
                    stroke: null,
                    strokeWidth: '0',
                    opacity: String(gradient?.stops?.[0]?.alpha ?? 1)
                },
                gradient,
                shadow
            });
        }
    }

    return drawables;
}

function createDynamicShapeOptions(element, slideContext, points, svgStyles) {
    const style = element.style;
    const left = parseFloat(style.left) || 0;
    const top = parseFloat(style.top) || 0;
    const width = parseFloat(style.width) || 100;
    const height = parseFloat(style.height) || 100;
    const scaleX = slideContext?.scaleX || 1;
    const scaleY = slideContext?.scaleY || 1;

    const shapeOptions = {
        x: Math.round(((left / 72) * scaleX) * 100) / 100,
        y: Math.round(((top / 72) * scaleY) * 100) / 100,
        w: Math.round((((width / 72) * scaleX)) * 100) / 100,
        h: Math.round((((height / 72) * scaleY)) * 100) / 100,
        points
    };

    // ── Opacity / Transparency ────────────────────────────────────────────────

    // background-gradient extraction).
    let opacityStr = style.opacity;
    if ((opacityStr === '' || opacityStr == null) && element.getAttribute) {
        const rawStyle = element.getAttribute('style') || '';
        const m = rawStyle.match(/\bopacity\s*:\s*([\d.]+)/i);
        if (m) opacityStr = m[1];
    }


    const svgNodeOpacity = (svgStyles.opacity !== '1' && svgStyles.opacity != null) ? svgStyles.opacity : null;
    const rawOpacity = parseFloat(opacityStr || svgNodeOpacity || svgStyles.fillOpacity || '1');
    const opacity = Number.isFinite(rawOpacity) ? Math.max(0, Math.min(1, rawOpacity)) : 1;
    const transparency = opacity < 1 ? Math.round((1 - opacity) * 100) : 0;

    // ── Fill ──────────────────────────────────────────────────────────────────

    const fillColor = extractColor(svgStyles.fill);
    if (fillColor) {
        shapeOptions.fill = (transparency > 0 && transparency <= 100)
            ? { color: fillColor, transparency }
            : fillColor;
    }

    // ── Stroke ───────────────────────────────────────────────────────────────
    const strokeColor = extractColor(svgStyles.stroke);
    const strokeWidth = parseFloat(svgStyles.strokeWidth || '0');
    if (strokeColor && Number.isFinite(strokeWidth) && strokeWidth > 0) {
        shapeOptions.line = { color: strokeColor, width: strokeWidth };

        const dashType = mapStrokeDasharrayToPptDashType(svgStyles.strokeDasharray, strokeWidth);
        if (dashType) shapeOptions.line.dashType = dashType;
    }

    if (transparency > 0 && transparency <= 100) shapeOptions.transparency = transparency;

    if (opacity < 1) shapeOptions._opacity = opacity;

    const transform = style.transform || '';
    const rotateMatch = transform.match(/rotate\((-?\d*\.?\d+)deg\)/);    
    if (rotateMatch) {
        const rotation = parseFloat(rotateMatch[1]);
        if (Number.isFinite(rotation) && rotation !== 0 && Math.abs(rotation) <= 360) {
            shapeOptions.rotate = Math.round(rotation);
        }
    }

    const hasFlipBoth = /scale\(\s*-1\s*,\s*-1\s*\)/.test(transform);
    const hasFlipH = /scaleX\(\s*-1\s*\)|scale\(\s*-1\s*,\s*1\s*\)/.test(transform);
    const hasFlipV = /scaleY\(\s*-1\s*\)|scale\(\s*1\s*,\s*-1\s*\)/.test(transform);
    if (hasFlipBoth) {
        shapeOptions.flipH = true;
        shapeOptions.flipV = true;
    } else {
        if (hasFlipH) shapeOptions.flipH = true;
        if (hasFlipV) shapeOptions.flipV = true;
    }

    // ── Shadow ────────────────────────────────────────────────────────────────

    const rawWrapperShadow = extractRawBoxShadow(element);
    if (rawWrapperShadow) {
        const parsedShadow = parseBoxShadowForPptxgenjs(rawWrapperShadow);
        if (parsedShadow) shapeOptions.shadow = parsedShadow;
    }

    // Detect flipH / flipV from CSS transform scaleX(-1) / scaleY(-1)
    const scaleXMatch = transform.match(/scaleX\((-?\d*\.?\d+)\)/);
    const scaleYMatch = transform.match(/scaleY\((-?\d*\.?\d+)\)/);
    if (scaleXMatch && parseFloat(scaleXMatch[1]) < 0) shapeOptions.flipH = true;
    if (scaleYMatch && parseFloat(scaleYMatch[1]) < 0) shapeOptions.flipV = true;

    // Detect flipH / flipV from SVG inner <g transform="translate(w,h) scale(-1,-1)">
    if (!shapeOptions.flipH || !shapeOptions.flipV) {
        const svgGroup = element.querySelector && element.querySelector('g[transform]');
        if (svgGroup) {
            const gTransform = svgGroup.getAttribute('transform') || '';
            const scaleMatch = gTransform.match(/scale\(\s*(-?\d*\.?\d+)\s*(?:,\s*(-?\d*\.?\d+)\s*)?\)/);
            if (scaleMatch) {
                const sx = parseFloat(scaleMatch[1]);
                const sy = scaleMatch[2] !== undefined ? parseFloat(scaleMatch[2]) : sx;
                if (sx < 0) shapeOptions.flipH = true;
                if (sy < 0) shapeOptions.flipV = true;
            }
        }
    }

    return shapeOptions;
}

function validateShapeOptions(shapeOptions) {
    if (!shapeOptions.w || shapeOptions.w <= 0 || !shapeOptions.h || shapeOptions.h <= 0) return false;
    if (!shapeOptions.points || shapeOptions.points.length === 0) return false;
    if (!shapeOptions.points.some(point => point.moveTo === true)) return false;

    shapeOptions.x = Number.isFinite(shapeOptions.x) ? shapeOptions.x : 0;
    shapeOptions.y = Number.isFinite(shapeOptions.y) ? shapeOptions.y : 0;

    const hasFill = shapeOptions.fill !== undefined && shapeOptions.fill !== null;
    const hasStroke = shapeOptions.line && shapeOptions.line.width > 0;
    if (!hasFill && !hasStroke) {
        shapeOptions.line = { color: 'CCCCCC', width: 0.5 };
    }

    return true;
}

function createFallbackShape(shapeOptions, fallbackType = 'rect') {
    const fallback = { x: shapeOptions.x, y: shapeOptions.y, w: shapeOptions.w, h: shapeOptions.h };
    if (shapeOptions.fill) fallback.fill = shapeOptions.fill;
    if (shapeOptions.line) fallback.line = shapeOptions.line;
    if (shapeOptions.transparency) fallback.transparency = shapeOptions.transparency;
    if (shapeOptions.rotate) fallback.rotate = shapeOptions.rotate;
    return { fallback, type: fallbackType };
}

function normalizeSvgObjectBaseName(rawName) {
    if (!rawName) return 'unnamed';

    let name = String(rawName).trim();
    if (!name) return 'unnamed';

    while (true) {
        const nestedMatch = name.match(/^Custom SVG Shape \((.*)\)(?: #\d+)?$/);
        if (!nestedMatch) break;
        name = nestedMatch[1].trim();
    }

    return name || 'unnamed';
}

function addSvgToSlide(pptSlide, svgElement, elementStyle, slideContext) {
    try {
        const parentElement =
            svgElement.closest('.shape.custom-shape') ||
            svgElement.closest('.custom-shape') ||
            svgElement.closest('#custGeom') ||
            svgElement.closest('.sli-svg-container') ||
            svgElement.parentElement;

        if (!parentElement) return false;

        const style = parentElement.style;
        const scaleX = slideContext?.scaleX || 1;
        const scaleY = slideContext?.scaleY || 1;
        const shapeWidthInches = Math.round((((parseFloat(style.width) || 100) / 72) * scaleX) * 100) / 100;
        const shapeHeightInches = Math.round((((parseFloat(style.height) || 100) / 72) * scaleY) * 100) / 100;
        if (shapeWidthInches <= 0 || shapeHeightInches <= 0) return false;

        const drawables = collectSvgDrawables(svgElement);
        if (drawables.length === 0) return false;

        const normalizationBounds = resolveNormalizationBounds(svgElement, drawables);

        let addedAny = false;

        drawables.forEach((drawable, index) => {
            const points = normalizePoints(drawable.points, normalizationBounds, shapeWidthInches, shapeHeightInches);
            const shapeOptions = createDynamicShapeOptions(parentElement, slideContext, points, drawable.styles);

            if (!validateShapeOptions(shapeOptions)) {
                const { fallback, type } = createFallbackShape(shapeOptions);
                pptSlide.addShape(type, fallback);
                addedAny = true;
                return;
            }

            const baseName = normalizeSvgObjectBaseName(parentElement.dataset?.name || parentElement.className || 'unnamed');
            shapeOptions.objectName = `Custom SVG Shape (${baseName}) #${index + 1}`;

            const rawFill = (drawable.styles.fill || '').trim().toLowerCase();
            const rawLinecap = (drawable.styles.strokeLinecap || '').trim().toLowerCase();
            const rawLinejoin = (drawable.styles.strokeLinejoin || '').trim().toLowerCase();
            const strokeWidth = parseFloat(drawable.styles.strokeWidth || '0');
            const dashType = shapeOptions.line?.dashType || mapStrokeDasharrayToPptDashType(drawable.styles.strokeDasharray, strokeWidth);
            const isOpenPath = !drawable.points.some(point => point.close === true);
            const hasStroke = !!shapeOptions.line?.color && Number.isFinite(shapeOptions.line?.width) && shapeOptions.line.width > 0;
            const hasExplicitNoFill = rawFill === 'none';

            if (hasStroke || hasExplicitNoFill || dashType || rawLinecap || rawLinejoin) {
                if (!global.svgLineStyleStore) global.svgLineStyleStore = new Map();
                global.svgLineStyleStore.set(shapeOptions.objectName, {
                    strokeColor: shapeOptions.line?.color || null,
                    strokeWidth: hasStroke ? shapeOptions.line.width : 0,
                    strokeGradient: drawable.strokeGradient || null,
                    dashType: dashType || null,
                    cap: rawLinecap || (hasStroke && isOpenPath ? 'butt' : null),
                    join: rawLinejoin || (hasStroke && isOpenPath ? 'miter' : null),
                    noFill: hasExplicitNoFill || (!shapeOptions.fill && hasStroke),
                    isOpenPath
                });
            }


            if (shapeOptions._opacity !== undefined && shapeOptions._opacity < 1) {
                if (!global.svgOpacityStore) global.svgOpacityStore = new Map();
                global.svgOpacityStore.set(shapeOptions.objectName, {
                    opacity: shapeOptions._opacity,
                    // OOXML <a:alpha val="..."/> uses 1/1000ths of a percent
                    // (100 000 = fully opaque, 0 = fully transparent).
                    alphaVal: Math.round(shapeOptions._opacity * 100000)
                });
            }

            delete shapeOptions._opacity;


            if (drawable.gradient) {
                if (!global.svgGradientFillStore) global.svgGradientFillStore = new Map();
                global.svgGradientFillStore.set(shapeOptions.objectName, drawable.gradient);
            }


            if (drawable.shadow) {
                shapeOptions.shadow = drawable.shadow;
            }

            if (!shapeOptions.shadow) {
                const svgDropShadow = parseDropShadowFilter(svgElement);
                if (svgDropShadow) shapeOptions.shadow = svgDropShadow;
            }

            const activeShadow = shapeOptions.shadow;
            if (activeShadow && typeof activeShadow.opacity === 'number' && activeShadow.opacity < 1) {
                if (!global.svgShadowStore) global.svgShadowStore = new Map();
                global.svgShadowStore.set(shapeOptions.objectName, {
                    type: activeShadow.type || 'outer',
                    color: activeShadow.color || '000000',
                    // Pre-calculate OOXML alpha so post-processor needs no math
                    alphaVal: Math.round(activeShadow.opacity * 100000)
                });
            }

            try {
                pptSlide.addShape('custGeom', shapeOptions);
                addedAny = true;
            } catch (custGeomError) {
                const polyPoints = points.filter(point => !point.curve && !point.close && point.x !== undefined);
                if (polyPoints.length >= 3) {
                    pptSlide.addShape('custGeom', { ...shapeOptions, points: polyPoints });
                    addedAny = true;
                    return;
                }

                const { fallback, type } = createFallbackShape(shapeOptions);
                pptSlide.addShape(type, fallback);
                addedAny = true;
            }
        });

        return addedAny;
    } catch (err) {
        console.error('   Critical error in addSvgToSlide:', err.message);
        try {
            pptSlide.addShape('rect', {
                x: 0,
                y: 0,
                w: 0.5,
                h: 0.5,
                fill: 'FFFFFF',
                objectName: 'SVG Emergency Fallback'
            });
        } catch (_) {
        }
        return false;
    }
}

function addSvgConnectorToSlide(pptSlide, svgElement, elementStyle, slideContext) {
    try {
        const pathElement = svgElement.querySelector('path') || svgElement.querySelector('line');
        if (!pathElement) return false;

        const parentElement = svgElement.closest('.sli-svg-connector') || svgElement.parentElement;
        const style = parentElement.style;
        const scaleX = slideContext?.scaleX || 1;
        const scaleY = slideContext?.scaleY || 1;

        const x = ((parseFloat(style.left) || 0) / 72) * scaleX;
        const y = ((parseFloat(style.top) || 0) / 72) * scaleY;
        const w = Math.max(((parseFloat(style.width) || 100) / 72) * scaleX, 0.01);
        const h = Math.max(((parseFloat(style.height) || 100) / 72) * scaleY, 0.01);

        const stroke = extractColor(pathElement.getAttribute('stroke') || '#000000');
        const strokeWidth = parseFloat(pathElement.getAttribute('stroke-width') || '1');
        if (!(strokeWidth > 0)) return false;
        const dashType = mapStrokeDasharrayToPptDashType(pathElement.getAttribute('stroke-dasharray'), strokeWidth);

        pptSlide.addShape('line', {
            x: Math.round(x * 100) / 100,
            y: Math.round(y * 100) / 100,
            w: Math.round(w * 100) / 100,
            h: Math.round(h * 100) / 100,
            line: {
                color: stroke || '000000',
                width: strokeWidth,
                ...(dashType ? { dashType } : {})
            }
        });

        return true;
    } catch (err) {
        console.error('   Error adding SVG connector:', err.message);
        return false;
    }
}

function processSvgElement(pptSlide, element, slideContext) {
    try {
        let svgElement = element.querySelector('svg');
        if (!svgElement) return false;

        const innerSvg = svgElement.querySelector('svg');
        if (innerSvg) svgElement = innerSvg;

        if (element.classList.contains('sli-svg-connector')) {
            return addSvgConnectorToSlide(pptSlide, svgElement, element.style, slideContext);
        }

        return addSvgToSlide(pptSlide, svgElement, element.style, slideContext);
    } catch (err) {
        console.error('   Error processing SVG element:', err.message);
        return false;
    }
}

module.exports = {
    addSvgToSlide,
    addSvgConnectorToSlide,
    processSvgElement,

    createDynamicShapeOptions,
    parsePathDataToRawPoints,
    collectSvgDrawables
};