const PptxGenJS = require("pptxgenjs");

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

function computePathBounds(pathData) {
    return computePointsBounds(parsePathDataToRawPoints(pathData));
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

function collectNodeStyles(node, svgElement) {
    return {
        fill: node.getAttribute('fill') || svgElement.style.fill || null,
        fillOpacity: node.getAttribute('fill-opacity') || null,
        stroke: node.getAttribute('stroke') || svgElement.style.stroke || null,
        strokeWidth: node.getAttribute('stroke-width') || svgElement.style.strokeWidth || '0',
        opacity: node.getAttribute('opacity') || svgElement.style.opacity || '1'
    };
}

function collectSvgDrawables(svgElement) {
    const nodes = Array.from(svgElement.querySelectorAll('path, polygon, polyline, rect, circle, ellipse'));
    return nodes
        .map(node => {
            const pathData = resolveShapePath(node);
            if (!pathData) return null;

            const rawPoints = parsePathDataToRawPoints(pathData);
            if (rawPoints.length === 0) return null;

            return {
                node,
                points: transformPoints(rawPoints, getNodeTransformMatrix(node, svgElement)),
                styles: collectNodeStyles(node, svgElement)
            };
        })
        .filter(Boolean);
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

    const fillColor = extractColor(svgStyles.fill);
    if (fillColor) shapeOptions.fill = fillColor;

    const strokeColor = extractColor(svgStyles.stroke);
    const strokeWidth = parseFloat(svgStyles.strokeWidth || '0');
    if (strokeColor && Number.isFinite(strokeWidth) && strokeWidth > 0) {
        shapeOptions.line = { color: strokeColor, width: Math.min(strokeWidth, 10) };
    }

    const opacity = parseFloat(style.opacity || svgStyles.opacity || svgStyles.fillOpacity || '1');
    const transparency = opacity < 1 ? Math.round((1 - opacity) * 100) : 0;
    if (transparency > 0 && transparency <= 100) shapeOptions.transparency = transparency;

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

    return shapeOptions;
}

function validateShapeOptions(shapeOptions) {
    if (!shapeOptions.w || shapeOptions.w <= 0 || !shapeOptions.h || shapeOptions.h <= 0) return false;
    if (!shapeOptions.points || shapeOptions.points.length === 0) return false;
    if (!shapeOptions.points.some(point => point.moveTo === true)) return false;

    shapeOptions.x = Math.max(0, shapeOptions.x || 0);
    shapeOptions.y = Math.max(0, shapeOptions.y || 0);

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

        pptSlide.addShape('line', {
            x: Math.round(x * 100) / 100,
            y: Math.round(y * 100) / 100,
            w: Math.round(w * 100) / 100,
            h: Math.round(h * 100) / 100,
            line: { color: stroke || '000000', width: Math.min(strokeWidth, 10) }
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
    convertSvgPathToPptxPoints: (pathData, viewBoxWidth, viewBoxHeight, shapeWidth, shapeHeight, boundsOverride) => {
        const rawPoints = parsePathDataToRawPoints(pathData);
        const bounds = boundsOverride || computePointsBounds(rawPoints);
        return normalizePoints(rawPoints, bounds, shapeWidth, shapeHeight);
    },
    createDynamicShapeOptions,
    computePathBounds,
    parsePathDataToRawPoints,
    collectSvgDrawables,
    PptxGenJS
};
