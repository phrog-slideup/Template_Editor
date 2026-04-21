const addTextBox = require("./addTextBoxToSlide.js");
const svgAddToSlide = require("./addSvgToSlide.js");
const { parseShadowFromBoxShadow } = require('./parseShadowFromBoxShadow.js');
const { extractGlowFromWrapperHTML } = require('./handleShapeGlow.js');


function extractSvgPresentationProps(shapeElement) {
    const svgEl = shapeElement.querySelector('svg');
    if (!svgEl) return null;

    const shapeNode = svgEl.querySelector('polygon, path, rect, ellipse, circle');
    if (!shapeNode) return null;

    const fill = shapeNode.getAttribute('fill');
    const stroke = shapeNode.getAttribute('stroke');
    const strokeWidthRaw = shapeNode.getAttribute('stroke-width');
    const opacityRaw = shapeNode.getAttribute('opacity');

    let strokeWidth = 0;
    if (strokeWidthRaw !== null && strokeWidthRaw !== undefined && strokeWidthRaw !== '') {
        strokeWidth = parseFloat(strokeWidthRaw);
        if (isNaN(strokeWidth)) strokeWidth = 0;
    }

    let opacity = 1;
    if (opacityRaw !== null && opacityRaw !== undefined && opacityRaw !== '') {
        const parsedOpacity = parseFloat(opacityRaw);
        if (!isNaN(parsedOpacity)) opacity = parsedOpacity;
    }

    return {
        fill,
        stroke,
        strokeWidth,
        opacity
    };
}


// ========== CORRECTED Connector ==========

function convertConnectorToPPTX(pptx, pptSlide, connectorData) {

    // Convert stroke color (remove # if present)
    const lineColor = connectorData.strokeColor.replace('#', '');

    // Convert stroke width from pixels to points (1px = 0.75pt)
    const lineWidth = connectorData.strokeWidth * 0.75;

    // Build line options
    const lineOptions = {
        color: lineColor,
        width: lineWidth
    };

    // Add dash type if not solid
    if (connectorData.dashType && connectorData.dashType !== 'solid') {
        lineOptions.dashType = connectorData.dashType;
    }

    // ✅ CHECK IF THIS IS A STRAIGHT CONNECTOR OR LINE
    const isStraightConnector = connectorData.shapeType === 'straightConnector1' ||
        connectorData.shapeType === 'line';

    // ========== SPECIAL HANDLING FOR STRAIGHT CONNECTORS ==========
    if (isStraightConnector && (!connectorData.segments || connectorData.segments.length === 0)) {

        // Use position data as fallback
        const x = connectorData.position.x / 72;
        const y = connectorData.position.y / 72;
        const w = Math.max(connectorData.position.width / 72, 0.01);
        const h = Math.max(connectorData.position.height / 72, 0.01);

        const simpleLineOptions = {
            x: x,
            y: y,
            w: w,
            h: h,
            line: lineOptions,
            objectName: connectorData.shapeName
        };

        // Apply arrows
        let startArrow = null;
        let endArrow = null;

        if (connectorData.lineEnds) {
            if (connectorData.lineEnds.headType && connectorData.lineEnds.headType !== 'none') {
                startArrow = mapArrowType(connectorData.lineEnds.headType);
            }
            if (connectorData.lineEnds.tailType && connectorData.lineEnds.tailType !== 'none') {
                endArrow = mapArrowType(connectorData.lineEnds.tailType);
            }
        }

        if (startArrow) simpleLineOptions.line.beginArrowType = startArrow;
        if (endArrow) simpleLineOptions.line.endArrowType = endArrow;

        if (connectorData.flipH) simpleLineOptions.flipH = true;
        if (connectorData.flipV) simpleLineOptions.flipV = true;

        pptSlide.addShape(pptx.shapes.LINE, simpleLineOptions);
        return;
    }

    // ✅ STRAIGHT CONNECTOR WITH SEGMENT DATA - THIS IS THE KEY FIX
    if (isStraightConnector && connectorData.segments && connectorData.segments.length > 0) {

        const segment = connectorData.segments[0];

        // ✅ CRITICAL: Extract actual line endpoints from segment
        // These are in ABSOLUTE pixel coordinates
        const x1Px = segment.x1;
        const y1Px = segment.y1;
        const x2Px = segment.x2;
        const y2Px = segment.y2;

        // Calculate the actual direction vector
        const deltaX = x2Px - x1Px;
        const deltaY = y2Px - y1Px;

        // Determine which direction the line is going
        const goingRight = deltaX >= 0;
        const goingDown = deltaY >= 0;
        const goingLeft = deltaX < 0;
        const goingUp = deltaY < 0;

        // ✅ PPTX LINE POSITIONING RULES:
        // - x, y = top-left corner of the line's bounding box
        // - w, h = width and height (always positive)
        // - flipH = true if line goes from right to left
        // - flipV = true if line goes from bottom to top

        // Calculate bounding box top-left corner
        const topLeftX = Math.min(x1Px, x2Px);
        const topLeftY = Math.min(y1Px, y2Px);

        // Calculate bounding box dimensions (always positive)
        const width = Math.abs(deltaX);
        const height = Math.abs(deltaY);

        // Convert to inches (72 DPI)
        const x = topLeftX / 72;
        const y = topLeftY / 72;
        const w = Math.max(width / 72, 0.01);
        const h = Math.max(height / 72, 0.01);

        // ✅ DETERMINE FLIP STATES BASED ON ACTUAL LINE DIRECTION
        let needsFlipH = false;
        let needsFlipV = false;

        // If line goes left (x2 < x1), we need horizontal flip
        if (goingLeft) {
            needsFlipH = true;
        }

        // If line goes up (y2 < y1), we need vertical flip
        if (goingUp) {
            needsFlipV = true;
        }

        // Build line options
        const lineShapeOptions = {
            x: x,
            y: y,
            w: w,
            h: h,
            line: lineOptions,
            objectName: connectorData.shapeName
        };

        // Apply flip states
        if (needsFlipH) {
            lineShapeOptions.flipH = true;
        }
        if (needsFlipV) {
            lineShapeOptions.flipV = true;
        }

        // ✅ APPLY ARROWS (Direct mapping for straight connectors)
        let startArrow = null;
        let endArrow = null;

        if (connectorData.lineEnds) {
            if (connectorData.lineEnds.headType && connectorData.lineEnds.headType !== 'none') {
                startArrow = mapArrowType(connectorData.lineEnds.headType);
            }
            if (connectorData.lineEnds.tailType && connectorData.lineEnds.tailType !== 'none') {
                endArrow = mapArrowType(connectorData.lineEnds.tailType);
            }
        }

        if (startArrow) {
            lineShapeOptions.line.beginArrowType = startArrow;
        }
        if (endArrow) {
            lineShapeOptions.line.endArrowType = endArrow;
        }

        pptSlide.addShape(pptx.shapes.LINE, lineShapeOptions);
        return;
    }

    // ========== ORIGINAL CODE FOR OTHER CONNECTOR TYPES ==========
    // All the existing code for bent and curved connectors remains unchanged

    let startArrow = null;
    let endArrow = null;

    if (connectorData.lineEnds) {
        if (connectorData.lineEnds.tailType && connectorData.lineEnds.tailType !== 'none') {
            startArrow = mapArrowType(connectorData.lineEnds.tailType);
        }
        if (connectorData.lineEnds.headType && connectorData.lineEnds.headType !== 'none') {
            endArrow = mapArrowType(connectorData.lineEnds.headType);
        }
    }

    const isCurvedConnector = connectorData.shapeType && connectorData.shapeType.includes('curved');

    if (isCurvedConnector) {

        const originalRotation = connectorData.rotation || 0;
        const flipH = connectorData.flipH || false;
        const flipV = connectorData.flipV || false;

        let effectiveRotation = originalRotation;

        if (flipH && !flipV) {
            effectiveRotation = -originalRotation;
        } else if (!flipH && flipV) {
            effectiveRotation = -originalRotation;
        }

        const curveSegments = generateCurveSegmentsForPPTX(connectorData);

        const arrowSegmentLength = 5;

        let startSegment = null;
        if (startArrow && curveSegments.length >= arrowSegmentLength) {
            const firstPoint = curveSegments[0];
            const fifthPoint = curveSegments[Math.min(arrowSegmentLength, curveSegments.length - 1)];
            startSegment = {
                x1: firstPoint.x1,
                y1: firstPoint.y1,
                x2: fifthPoint.x2,
                y2: fifthPoint.y2
            };
        }

        let endSegment = null;
        if (endArrow && curveSegments.length >= arrowSegmentLength) {
            const lastIndex = curveSegments.length - 1;
            const fifthFromLast = curveSegments[Math.max(0, lastIndex - arrowSegmentLength)];
            const lastPoint = curveSegments[lastIndex];
            endSegment = {
                x1: fifthFromLast.x1,
                y1: fifthFromLast.y1,
                x2: lastPoint.x2,
                y2: lastPoint.y2
            };
        }

        const transformedSegments = curveSegments.map(segment => {
            return transformSegmentForPPTX(
                segment,
                connectorData.position,
                effectiveRotation,
                flipH,
                flipV
            );
        });

        transformedSegments.forEach((segment, index) => {
            const segX1 = segment.x1 / 72;
            const segY1 = segment.y1 / 72;
            const segX2 = segment.x2 / 72;
            const segY2 = segment.y2 / 72;

            const segW = Math.abs(segX2 - segX1);
            const segH = Math.abs(segY2 - segY1);

            let finalW = Math.max(segW, 0.01);
            let finalH = Math.max(segH, 0.01);

            const goingLeft = segX2 < segX1;
            const goingUp = segY2 < segY1;

            const segmentLineOptions = {
                color: lineColor,
                width: lineWidth
            };

            if (connectorData.dashType && connectorData.dashType !== 'solid') {
                segmentLineOptions.dashType = connectorData.dashType;
            }

            const shapeOptions = {
                x: Math.min(segX1, segX2),
                y: Math.min(segY1, segY2),
                w: finalW,
                h: finalH,
                line: segmentLineOptions,
                objectName: `${connectorData.shapeName}_curve_seg${index + 1}`
            };

            if (goingLeft) shapeOptions.flipH = true;
            if (goingUp) shapeOptions.flipV = true;

            pptSlide.addShape(pptx.shapes.LINE, shapeOptions);
        });

        if (startArrow && startSegment) {
            const transformedStartSegment = transformSegmentForPPTX(
                startSegment,
                connectorData.position,
                effectiveRotation,
                flipH,
                flipV
            );
            addArrowSegment(pptx, pptSlide, transformedStartSegment, lineColor, lineWidth, startArrow, true, connectorData.shapeName);
        }

        if (endArrow && endSegment) {
            const transformedEndSegment = transformSegmentForPPTX(
                endSegment,
                connectorData.position,
                effectiveRotation,
                flipH,
                flipV
            );
            addArrowSegment(pptx, pptSlide, transformedEndSegment, lineColor, lineWidth, endArrow, false, connectorData.shapeName);
        }

        return;
    }

    // Handle bent connectors (unchanged)
    if (connectorData.segments && connectorData.segments.length > 0) {

        const originalRotation = connectorData.originalRotation !== undefined ? connectorData.originalRotation : connectorData.rotation;
        const flipH = connectorData.flipH || false;
        const flipV = connectorData.flipV || false;

        let effectiveRotation = originalRotation;

        if (flipH && !flipV) {
            effectiveRotation = -originalRotation;
        } else if (!flipH && flipV) {
            effectiveRotation = -originalRotation;
        }

        const transformedSegments = connectorData.segments.map(segment => {
            return transformSegmentForPPTX(
                segment,
                connectorData.position,
                effectiveRotation,
                flipH,
                flipV
            );
        });

        transformedSegments.forEach((segment, index) => {
            const isFirst = index === 0;
            const isLast = index === transformedSegments.length - 1;

            const segX1 = segment.x1 / 72;
            const segY1 = segment.y1 / 72;
            const segX2 = segment.x2 / 72;
            const segY2 = segment.y2 / 72;

            const segW = Math.abs(segX2 - segX1);
            const segH = Math.abs(segY2 - segY1);

            let finalW = Math.max(segW, 0.01);
            let finalH = Math.max(segH, 0.01);

            const goingLeft = segX2 < segX1;
            const goingUp = segY2 < segY1;

            const segmentLineOptions = {
                color: lineColor,
                width: lineWidth
            };

            if (connectorData.dashType && connectorData.dashType !== 'solid') {
                segmentLineOptions.dashType = connectorData.dashType;
            }

            if (isFirst && startArrow) {
                segmentLineOptions.beginArrowType = startArrow;
            }

            if (isLast && endArrow) {
                segmentLineOptions.endArrowType = endArrow;
            }

            const shapeOptions = {
                x: Math.min(segX1, segX2),
                y: Math.min(segY1, segY2),
                w: finalW,
                h: finalH,
                line: segmentLineOptions,
                objectName: `${connectorData.shapeName}_seg${index + 1}`
            };

            if (goingLeft) shapeOptions.flipH = true;
            if (goingUp) shapeOptions.flipV = true;

            pptSlide.addShape(pptx.shapes.LINE, shapeOptions);
        });

    } else {
        // Simple line without segments
        const x = connectorData.position.x / 72;
        const y = connectorData.position.y / 72;
        const w = Math.max(connectorData.position.width / 72, 0.01);
        const h = Math.max(connectorData.position.height / 72, 0.01);

        const simpleLineOptions = {
            x: x,
            y: y,
            w: w,
            h: h,
            line: lineOptions,
            objectName: connectorData.shapeName
        };

        if (startArrow) simpleLineOptions.line.beginArrowType = startArrow;
        if (endArrow) simpleLineOptions.line.endArrowType = endArrow;

        if (connectorData.flipH) simpleLineOptions.flipH = true;
        if (connectorData.flipV) simpleLineOptions.flipV = true;

        pptSlide.addShape(pptx.shapes.LINE, simpleLineOptions);
    }
}

// Helper functions (unchanged)
function transformSegmentForPPTX(segment, containerPosition, rotation, flipH, flipV) {
    const relX1 = segment.x1 - containerPosition.x;
    const relY1 = segment.y1 - containerPosition.y;
    const relX2 = segment.x2 - containerPosition.x;
    const relY2 = segment.y2 - containerPosition.y;

    let newX1 = relX1;
    let newY1 = relY1;
    let newX2 = relX2;
    let newY2 = relY2;

    if (rotation !== 0) {
        const radians = (rotation * Math.PI) / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);

        const centerX = containerPosition.width / 2;
        const centerY = containerPosition.height / 2;

        const dx1 = relX1 - centerX;
        const dy1 = relY1 - centerY;
        newX1 = centerX + (dx1 * cos - dy1 * sin);
        newY1 = centerY + (dx1 * sin + dy1 * cos);

        const dx2 = relX2 - centerX;
        const dy2 = relY2 - centerY;
        newX2 = centerX + (dx2 * cos - dy2 * sin);
        newY2 = centerY + (dx2 * sin + dy2 * cos);
    }

    if (flipH) {
        newX1 = containerPosition.width - newX1;
        newX2 = containerPosition.width - newX2;
    }

    if (flipV) {
        newY1 = containerPosition.height - newY1;
        newY2 = containerPosition.height - newY2;
    }

    return {
        x1: containerPosition.x + newX1,
        y1: containerPosition.y + newY1,
        x2: containerPosition.x + newX2,
        y2: containerPosition.y + newY2,
        type: segment.type
    };
}

function addArrowSegment(pptx, pptSlide, segment, lineColor, lineWidth, arrowType, isStart, shapeName) {
    const segX1 = segment.x1 / 72;
    const segY1 = segment.y1 / 72;
    const segX2 = segment.x2 / 72;
    const segY2 = segment.y2 / 72;

    const segW = Math.abs(segX2 - segX1);
    const segH = Math.abs(segY2 - segY1);
    const finalW = Math.max(segW, 0.01);
    const finalH = Math.max(segH, 0.01);

    const goingLeft = segX2 < segX1;
    const goingUp = segY2 < segY1;

    const arrowOptions = {
        x: Math.min(segX1, segX2),
        y: Math.min(segY1, segY2),
        w: finalW,
        h: finalH,
        line: {
            color: lineColor,
            width: lineWidth,
            [isStart ? 'beginArrowType' : 'endArrowType']: arrowType
        },
        objectName: `${shapeName}_${isStart ? 'start' : 'end'}_arrow`
    };

    if (goingLeft) arrowOptions.flipH = true;
    if (goingUp) arrowOptions.flipV = true;

    pptSlide.addShape(pptx.shapes.LINE, arrowOptions);
}

function mapArrowType(htmlArrowType) {
    const arrowMap = {
        'triangle': 'triangle',
        'arrow': 'arrow',
        'stealth': 'stealth',
        'diamond': 'diamond',
        'oval': 'oval',
        'circle': 'oval',
        'dot': 'oval',
        'none': null
    };

    return arrowMap[htmlArrowType] || null;
}

function generateCurveSegmentsForPPTX(connectorData) {
    const width = connectorData.position.width;
    const height = connectorData.position.height;
    const startX = connectorData.position.x;
    const startY = connectorData.position.y;
    const shapeType = connectorData.shapeType;

    const numSegments = 50;
    const points = [];

    switch (shapeType) {
        case "curvedConnector2":
            for (let i = 0; i <= numSegments; i++) {
                const t = i / numSegments;
                const controlX = width * 0.5;
                const controlY = height * 0.5;
                const x = startX + (1 - t) * (1 - t) * 0 + 2 * (1 - t) * t * controlX + t * t * width;
                const y = startY + (1 - t) * (1 - t) * 0 + 2 * (1 - t) * t * controlY + t * t * height;
                points.push({ x, y });
            }
            break;
        case "curvedConnector3":
            for (let i = 0; i <= numSegments; i++) {
                const t = i / numSegments;
                const easing = t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
                const x = startX + width * t;
                const y = startY + height * easing;
                points.push({ x, y });
            }
            break;
        case "curvedConnector4":
            for (let i = 0; i <= numSegments; i++) {
                const t = i / numSegments;
                const x = startX + width * t;
                const wave = Math.sin(t * Math.PI);
                const y = startY + height * t + height * 0.15 * wave;
                points.push({ x, y });
            }
            break;
        case "curvedConnector5":
            for (let i = 0; i <= numSegments; i++) {
                const t = i / numSegments;
                const x = startX + width * t;
                const wave = Math.sin(t * Math.PI * 1.5);
                const y = startY + height * t + height * 0.12 * wave;
                points.push({ x, y });
            }
            break;
        default:
            points.push({ x: startX, y: startY });
            points.push({ x: startX + width, y: startY + height });
            break;
    }

    const segments = [];
    for (let i = 0; i < points.length - 1; i++) {
        segments.push({
            type: 'curve',
            x1: points[i].x,
            y1: points[i].y,
            x2: points[i + 1].x,
            y2: points[i + 1].y
        });
    }

    return segments;
}
// ========== END OF connector Functionality ==========

function addShapeToSlide(pptx, pptSlide, shapeElement, slideContext) {
    const style = shapeElement.style;
    const shapeId = shapeElement.getAttribute("id");
    const textBox = shapeElement.querySelector('.sli-txt-box');
    const objName = shapeElement.getAttribute("data-name");

    // Extract theme color and luminance/alpha attributes
    const originalThemeColor = shapeElement.getAttribute("data-original-color");
    const originalLumMod = shapeElement.getAttribute("originallummod");
    const originalLumOff = shapeElement.getAttribute("originallumoff");
    const originalAlpha = shapeElement.getAttribute("originalalpha");
    const softEdgeRad = parseInt(shapeElement.getAttribute('data-soft-edge-rad') || '0', 10);

    // Helper function to create proper theme color object
    // NEW: Detect if originalThemeColor is srgbClr (hex) or schemeClr; add alpha when provided
    function createThemeColorObject(themeColor, lumMod, lumOff, alpha) {
        if (!themeColor || themeColor === 'null' || themeColor === 'undefined') {
            return null;
        }

        const raw = String(themeColor).trim();

        // Try explicit prefixes first: "srgbClr:#RRGGBB" / "schemeClr:accent1"
        const mSrgbPref = raw.match(/^(?:srgbclr|srgb|hex)\s*:\s*([#0-9a-fA-F]{3,8})$/i);
        const mSchemePref = raw.match(/^(?:schemeclr|scheme)\s*:\s*([A-Za-z0-9]+)$/i);

        // Hex helpers
        const stripHash = (s) => s.replace(/^#/, '');
        const isHex3 = /^[0-9a-fA-F]{3}$/;
        const isHex6 = /^[0-9a-fA-F]{6}$/;

        // Build transforms (lumMod/lumOff/alpha) onto color object
        const applyTransforms = (obj) => {
            if (lumMod && lumMod !== 'null' && lumMod !== 'undefined' && !isNaN(parseInt(lumMod))) {
                obj.lumMod = parseInt(lumMod);
            }
            if (lumOff && lumOff !== 'null' && lumOff !== 'undefined' && !isNaN(parseInt(lumOff))) {
                obj.lumOff = parseInt(lumOff);
            }
            // NEW: alpha support; accept 0–1, 0–100, or 0–100000 and clamp to 0–100000
            if (alpha && alpha !== 'null' && alpha !== 'undefined' && !isNaN(parseFloat(alpha))) {
                let aNum = parseFloat(alpha);
                if (aNum <= 1) aNum = Math.round(aNum * 100000);
                else if (aNum <= 100) aNum = Math.round(aNum * 1000);
                else aNum = Math.round(aNum);
                if (aNum < 0) aNum = 0;
                if (aNum > 100000) aNum = 100000;
                obj.alpha = aNum;
            }
            return obj;
        };

        // If explicitly marked as srgbClr:<hex>
        if (mSrgbPref) {
            let hx = stripHash(mSrgbPref[1]);
            if (isHex3.test(hx)) hx = hx.replace(/(.)/g, '$1$1');
            return applyTransforms({ srgbClr: hx.toUpperCase() });
        }

        // If explicitly marked as schemeClr:<token>
        if (mSchemePref) {
            return applyTransforms({ type: 'schemeClr', val: mSchemePref[1] });
        }

        // If it looks like a HEX, use srgbClr
        let hx = stripHash(raw);
        if (isHex3.test(hx) || isHex6.test(hx)) {
            if (isHex3.test(hx)) hx = hx.replace(/(.)/g, '$1$1');
            return applyTransforms({ srgbClr: hx.toUpperCase() });
        }

        // Otherwise treat it as a scheme token (normalize a few common aliases)
        const mapToOoxml = { text1: 'tx1', text2: 'tx2', background1: 'bg1', background2: 'bg2' };
        let schemeVal = (mapToOoxml[raw.toLowerCase()] || raw);
        return applyTransforms({ type: 'schemeClr', val: schemeVal });
    }

    const bgColor = getBackgroundColor(shapeElement);
    const hasVisibleBackground = bgColor && bgColor !== 'transparent' && bgColor !== 'none';

    const hasBorder = checkForValidBorder(style);

    const hasPatternFill = !!(
        shapeElement.getAttribute('data-pattern-prst') &&
        shapeElement.getAttribute('data-pattern-fg') &&
        shapeElement.getAttribute('data-pattern-bg')
    );

    // ✅ custGeom/SVG shapes must NEVER be early-returned here — their fill color lives on
    // the inner <path fill="..."> attribute, not on the wrapper div's CSS background-color.
    // getBackgroundColor() always returns 'transparent' for these, which would cause the
    // shape to be skipped entirely (only text would be written to the PPTX).
    const isCustGeomShape = (
        shapeElement.classList.contains('custom-shape') ||
        shapeElement.id === 'custGeom' ||
        shapeId === 'custGeom' ||
        shapeElement.classList.contains('sli-svg-connector') ||
        shapeElement.classList.contains('sli-svg-container')
    );

    // Skip ONLY if it's a plain text box with no background AND no border (never skip custGeom shapes)
    if (!isCustGeomShape && textBox && textBox.textContent.trim() && !hasVisibleBackground && !hasBorder && !hasPatternFill) {
        return;
    }

    // ✅ Freeform / SVG-based shapes (custGeom, custom-shape, svg connectors, inline graphics)
    // Route these to addSvgToSlide so we actually create a real custGeom in PPTX.
    // FIX Bug 4: sli-svg-container added — inline isometric/graphic SVGs were previously skipped.
    if (isCustGeomShape) {
        const ok = svgAddToSlide.processSvgElement(pptSlide, shapeElement, slideContext);
        if (ok) return;
        // If SVG parsing failed, continue with normal shape pipeline (or bail out below).
    }

    // ✅ If no ID, still try to process SVG-based shapes (most freeforms come here)
    if (shapeId === null) {
        const ok = svgAddToSlide.processSvgElement(pptSlide, shapeElement, slideContext);
        if (ok) return;
        return; // nothing we can do without a shape id
    }

    // Get slide dimensions for proper scaling
    // const slideDimensions = getSlideDimensions(pptSlide);

    // Check if this is a special shape (not a rectangle)
    const isSpecialShape = shapeId !== 'rect';

    // Get inline background style
    let bgStyle = getBackgroundColor(shapeElement);

    // FIXED: Parse all transform properties including rotation and flipping
    const transform = style.transform || '';
    const rotationMatch = transform.match(/rotate\(([-\d.]+)deg\)/);
    const scaleXMatch = transform.match(/scaleX\((-?\d*\.?\d+)\)/);
    const scaleYMatch = transform.match(/scaleY\((-?\d*\.?\d+)\)/);

    let dominantColor;
    const rotation = rotationMatch ? parseFloat(rotationMatch[1]) : 0;

    // FIXED: Extract flip information from scale transforms
    let flipHorizontal = false;
    let flipVertical = false;

    if (scaleXMatch && parseFloat(scaleXMatch[1]) < 0) {
        flipHorizontal = true;
    }
    if (scaleYMatch && parseFloat(scaleYMatch[1]) < 0) {
        flipVertical = true;
    }

    let gradient = parseGradient(bgStyle);

    if (gradient) {
        dominantColor = getDominantColorFromGradient(gradient);
    } else {
        dominantColor = rgbToHex(bgStyle);
    }

    // ===== Convert HTML/CSS border-radius → PptxGenJS rectRadius (inches) with EMU-aware rounding =====
    const EMU = 914400;                 // EMUs per inch (PowerPoint)
    const DPI = 72;                     // your pipeline uses 72 DPI

    // Read sizes (prefer decimals if available via inline styles)
    const widthPx = parseFloat(style.width || "0");
    const heightPx = parseFloat(style.height || "0");
    const smallerPx = Math.min(widthPx, heightPx);

    // Parse border-radius
    const borderRadiusRaw = String(style.borderRadius || "0").trim();
    let rPx = 0;

    if (borderRadiusRaw.endsWith("%")) {
        // % is per-axis in CSS; to map to roundRect (single adj), use the tighter (min) side
        const p = (parseFloat(borderRadiusRaw) || 0) / 100;
        rPx = p * smallerPx;
    } else {
        // px
        rPx = parseFloat(borderRadiusRaw) || 0;
    }

    // Clamp to half the smaller side (max roundness)
    if (smallerPx > 0) rPx = Math.min(rPx, smallerPx / 2);

    // --- Compute target adj the same way DrawingML defines it: adj = 100000 * r / minDim ---
    const desiredAdjInt = Math.round((100000 * rPx) / (smallerPx || 1));

    // --- Now invert PptxGenJS's XML formula to get rectRadius (inches) that rounds back to desiredAdjInt ---
    // PptxGenJS: adj_xml = Math.round( (rectRadiusIn * EMU * 100000) / Math.min(cx,cy) )
    const minInches = (smallerPx / DPI);
    const minCxCyEMU = Math.round(minInches * EMU);

    let borderRadiusValue = 0; // <-- rectRadius in inches for slideItemObj.options.rectRadius
    if (minCxCyEMU > 0) {
        borderRadiusValue = (desiredAdjInt * minCxCyEMU) / (100000 * EMU);
    }

    // Optional: final safety clamp (should already be safe)
    borderRadiusValue = Math.max(0, Math.min(borderRadiusValue, minInches / 2));


    // ── Glow wrapper detection ─────────────────────────────────────────────────
    function safeClosest(el, selector) {
        if (!el) return null;
        if (typeof el.closest === 'function') return el.closest(selector);
        // jsdom fallback: walk parentElement chain manually
        let cur = el.parentElement;
        while (cur) {
            if (typeof cur.matches === 'function' && cur.matches(selector)) return cur;
            if (cur.className && typeof cur.className === 'string' &&
                cur.className.includes('shape-glow-wrapper') && selector === '.shape-glow-wrapper') return cur;
            cur = cur.parentElement;
        }
        return null;
    }

    const glowWrapper = safeClosest(shapeElement, '.shape-glow-wrapper');

    // ── Position source ────────────────────────────────────────────────────────

    const positionSource = glowWrapper || shapeElement;
    const positionStyle = positionSource.getAttribute
        ? (() => {
            // Parse inline style attribute string into an object with left/top/width/height
            const raw = positionSource.getAttribute('style') || '';
            const get = (prop) => {
                const m = raw.match(new RegExp(prop + '\\s*:\\s*([0-9.-]+)px', 'i'));
                return m ? m[1] : null;
            };
            return {
                left: get('left'),
                top: get('top'),
                width: get('width'),
                height: get('height'),
            };
        })()
        : {};

    // Use wrapper values when available, fall back to shapeElement style
    const rawLeft = positionStyle.left !== null ? positionStyle.left : (style.left || '0').replace('px', '');
    const rawTop = positionStyle.top !== null ? positionStyle.top : (style.top || '0').replace('px', '');
    const rawWidth = positionStyle.width !== null ? positionStyle.width : (style.width || '0').replace('px', '');
    const rawHeight = positionStyle.height !== null ? positionStyle.height : (style.height || '0').replace('px', '');


    // ── Glow parsing ───────────────────────────────────────────────────────────
    // The filter: drop-shadow(...) is on the wrapper div — parse it back to
    // PPTX glow parameters (radEmu, colorHex, alphaVal) via handleShapeGlow.js
    let glowOptions = null;
    if (glowWrapper) {
        const wrapperStyleAttr = glowWrapper.getAttribute('style') || '';
        const glowData = extractGlowFromWrapperHTML(wrapperStyleAttr);
        if (glowData) {
            glowOptions = {
                size: glowData.radEmu / 12700,           // EMU → points
                color: glowData.colorHex.replace('#', ''), // hex without #
                opacity: glowData.alphaVal / 100000,         // 0–1
            };
        }
    }

    // ── Extract position and dimensions ───────────────────────────────────────
    const x = parseFloat(rawLeft) / 72;
    const y = parseFloat(rawTop) / 72;
    const w = parseFloat(rawWidth) / 72;
    const h = parseFloat(rawHeight) / 72;



    const shadowOptions = parseShadowFromBoxShadow(shapeElement);

    let shapeOptions = {
        x: x,
        y: y,
        w: w,
        h: h,
        rotate: rotation,
        objectName: objName || '',
        hidden: true,
        ...(shadowOptions ? { shadow: shadowOptions } : {}),
        ...(glowOptions ? { glow: glowOptions } : {}),
    };

    //     shapeOptions.shadow = {
    //     type: 'outer',        // 'outer', 'inner', or 'perspective'
    //     color: 'f8ff1d',      // Hex color (without #)
    //     blur: 28,              // Blur amount in points
    //     offset: 3,            // Distance/offset in points
    //     angle: 180,            // Angle in degrees
    //     opacity: 1          // 0-1 transparency
    //   }
    // FIXED: Add flip properties to shape options
    if (flipHorizontal) {
        shapeOptions.flipH = true;
    }
    if (flipVertical) {
        shapeOptions.flipV = true;
    }



    // if (style.opacity) {
    //     const opacityValue = parseFloat(style.opacity);
    //     const transparencyValue = Math.round((1 - opacityValue) * 100);

    //     // NEW: pass originalAlpha into the color object builder
    //     const themeColorObj = createThemeColorObject(originalThemeColor, originalLumMod, originalLumOff, originalAlpha);

    //     if (themeColorObj) {
    //         shapeOptions.fill = {
    //             color: themeColorObj
    //         };
    //     } else {
    //         shapeOptions.fill = {
    //             color: rgbToHex(dominantColor),
    //             transparency: transparencyValue
    //         };
    //     }
    // }

    // if (style.borderWidth && parseFloat(style.borderWidth) > 0) {

    //     const cssStyle = (style.borderStyle || 'solid').toLowerCase();

    //     const dashMap = {
    //         solid: 'solid',
    //         dotted: 'sysDot',
    //         dashed: 'dash'
    //     };

    //     const dashType = dashMap[cssStyle] || '';

    //     shapeOptions.line = {
    //         color: rgbToHex(style.borderColor || "#CD0C2C"),
    //         width: parseFloat(style.borderWidth) * (72 / 72),
    //         dashType: dashType
    //     };
    // }


    // note rakesh::::::: aboove is the previous code below is mine new
    if (style.opacity && hasVisibleBackground) {
        const opacityValue = parseFloat(style.opacity);
        const transparencyValue = Math.round((1 - opacityValue) * 100);

        // Only use theme color if background is not an explicit hex/rgb value
        const isExplicitColor = dominantColor && (dominantColor.startsWith('#') || dominantColor.startsWith('rgb'));

        if (!isExplicitColor && originalThemeColor && originalThemeColor !== 'null' && originalThemeColor !== 'undefined') {
            const themeColorObj = createThemeColorObject(originalThemeColor, originalLumMod, originalLumOff, originalAlpha);
            if (themeColorObj) {
                shapeOptions.fill = {
                    color: themeColorObj
                };
            } else if (dominantColor) {
                shapeOptions.fill = {
                    color: rgbToHex(dominantColor),
                    transparency: transparencyValue
                };
            }
        } else if (dominantColor) {
            shapeOptions.fill = {
                color: rgbToHex(dominantColor),
                transparency: transparencyValue
            };
        }
    }

    if (style.borderWidth && parseFloat(style.borderWidth) > 0) {

        const cssStyle = (style.borderStyle || 'solid').toLowerCase();

        const dashMap = {
            solid: 'solid',
            dotted: 'sysDot',
            dashed: 'dash'
        };

        const dashType = dashMap[cssStyle] || '';

        shapeOptions.line = {
            color: rgbToHex(style.borderColor || "#CD0C2C"),
            width: parseFloat(style.borderWidth) * (72 / 72),
            dashType: dashType
        };


        // rakesh note : before it is taking the border color as fallback here i add more first check for the original border color then fallbacktoblack

        const originalBorderColor = shapeElement.getAttribute("data-original-border-color");

        let borderColor = style.borderColor;

        // If no style.borderColor, try original color
        if (!borderColor || borderColor === 'transparent') {
            borderColor = originalBorderColor;
        }

        // Only use fallback if absolutely no color is found
        if (!borderColor || borderColor === 'transparent') {
            borderColor = "#000000"; // Use black as fallback instead of red
        }
        const borderDashStyle = getBorderDashFromStyle(style, shapeElement);
        shapeOptions.line = {
            color: rgbToHex(borderColor),
            width: parseFloat(style.borderWidth) * (72 / 72)
        };

        if (borderDashStyle && borderDashStyle !== 'solid') {
            shapeOptions.line.dashType = borderDashStyle;
        }

    }

    if (borderRadiusValue > 0) {
        if (shapeId === 'rect' ||
            shapeId === 'roundRect' ||
            shapeId === 'round1Rect' ||
            shapeId === 'round2DiagRect' ||
            shapeId === 'snip1Rect' ||
            shapeId === 'snip2DiagRect' ||
            shapeId === 'snip2SameRect' ||
            shapeId === 'snipRoundRect') {
            shapeOptions.rectRadius = borderRadiusValue;
        }
    }

    let calculatedTransparency = 0;
    if (style.opacity) {
        const opacityValue = parseFloat(style.opacity);
        calculatedTransparency = Math.round((1 - opacityValue) * 100);
    }

    // ✅ PATTERN FILL: Check for pattern data attributes first
    const patternPrst = shapeElement.getAttribute('data-pattern-prst');
    const patternFg = shapeElement.getAttribute('data-pattern-fg');
    const patternBg = shapeElement.getAttribute('data-pattern-bg');

    // Around line where patternFillStore.set() is called
    if (patternPrst && patternFg && patternBg) {
        if (!global.patternFillStore) global.patternFillStore = new Map();

        // Use objName if available, otherwise use position-based key as fallback
        const storeKey = objName || `shape_${x.toFixed(2)}_${y.toFixed(2)}`;

        global.patternFillStore.set(storeKey, {
            prst: patternPrst,
            fg: patternFg.replace('#', '').toUpperCase(),
            bg: patternBg.replace('#', '').toUpperCase(),
            shapeName: objName,  // keep original name for XML matching
            x: x,
            y: y,
            w: w,
            h: h
        });
        shapeOptions.fill = { color: patternFg.replace('#', '').toUpperCase() };
    }

    else if (gradient) {
        const gradientInfo = parseGradient(bgStyle);

        if (gradientInfo && gradientInfo.stops && gradientInfo.stops.length >= 2) {
            // Detect gradient type (linear or radial)
            const isRadial = /radial-gradient/i.test(bgStyle);

            const gradientType = isRadial ? 'radial' : 'linear';

            // Extract CSS angle (if linear)
            let cssAngle = 0;
            const angleMatch = bgStyle.match(/linear-gradient\(([-\d.]+)deg/);
            if (angleMatch) cssAngle = parseFloat(angleMatch[1]);

            // Convert CSS angle → PowerPoint compatible
            const pptxAngle = isRadial ? null : (cssAngle + 270) % 360;

            // ✅ NEW: Extract radial gradient focal point position
            let radialFocus = null;
            if (isRadial) {
                // Parse focal point from CSS: "radial-gradient(circle at right bottom, ...)"
                const focalMatch = bgStyle.match(/radial-gradient\s*\(\s*circle\s+at\s+([\w\s]+?)\s*,/i);
                if (focalMatch) {
                    const position = focalMatch[1].trim().toLowerCase();

                    // Map CSS positions to PowerPoint focal point percentages
                    // PowerPoint uses: path (shape type) and focus points (percentage from center)
                    const positionMap = {
                        'center': { path: 'circle', focusX: 50, focusY: 50 },
                        'center center': { path: 'circle', focusX: 50, focusY: 50 },

                        // Top positions
                        'top': { path: 'circle', focusX: 50, focusY: 0 },
                        'center top': { path: 'circle', focusX: 50, focusY: 0 },
                        'top center': { path: 'circle', focusX: 50, focusY: 0 },
                        'left top': { path: 'circle', focusX: 0, focusY: 0 },
                        'top left': { path: 'circle', focusX: 0, focusY: 0 },
                        'right top': { path: 'circle', focusX: 100, focusY: 0 },
                        'top right': { path: 'circle', focusX: 100, focusY: 0 },

                        // Bottom positions
                        'bottom': { path: 'circle', focusX: 50, focusY: 100 },
                        'center bottom': { path: 'circle', focusX: 50, focusY: 100 },
                        'bottom center': { path: 'circle', focusX: 50, focusY: 100 },
                        'left bottom': { path: 'circle', focusX: 0, focusY: 100 },
                        'bottom left': { path: 'circle', focusX: 0, focusY: 100 },
                        'right bottom': { path: 'circle', focusX: 100, focusY: 100 },
                        'bottom right': { path: 'circle', focusX: 100, focusY: 100 },

                        // Middle sides
                        'left': { path: 'circle', focusX: 0, focusY: 50 },
                        'left center': { path: 'circle', focusX: 0, focusY: 50 },
                        'center left': { path: 'circle', focusX: 0, focusY: 50 },
                        'right': { path: 'circle', focusX: 100, focusY: 50 },
                        'right center': { path: 'circle', focusX: 100, focusY: 50 },
                        'center right': { path: 'circle', focusX: 100, focusY: 50 }
                    };

                    radialFocus = positionMap[position] || { path: 'circle', focusX: 50, focusY: 50 };
                } else {
                    // Default to center if no position specified
                    radialFocus = { path: 'circle', focusX: 50, focusY: 50 };
                }
            }

            // Extract color stops, keep RGBA and positions
            const colorStops = gradientInfo.stops.map((stop) => {
                const rgbaMatch = stop.color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([0-9.]+))?\)/);
                let colorHex = '#000000';
                let alpha = 1;

                if (rgbaMatch) {
                    const [r, g, b] = [rgbaMatch[1], rgbaMatch[2], rgbaMatch[3]];
                    colorHex = `#${(+r).toString(16).padStart(2, '0')}${(+g)
                        .toString(16)
                        .padStart(2, '0')}${(+b).toString(16).padStart(2, '0')}`;
                    if (rgbaMatch[4]) alpha = parseFloat(rgbaMatch[4]);
                } else if (stop.color.startsWith('#')) {
                    colorHex = stop.color;
                }

                const position = parseFloat(stop.position) || 0;
                return { color: colorHex, position, alpha };
            });

            // ✅ FIXED: Calculate per-stop transparency from alpha values ONLY
            // DO NOT combine with global opacity - the alpha values already represent the final intended transparency
            const transparencies = colorStops.map((s) => {
                // Convert alpha to transparency: transparency = (1 - alpha) * 100
                return Math.round((1 - s.alpha) * 100);
            });

            // ✅ Pass per-stop transparency array to the gradient
            const gradientObj = {
                type: gradientType,
                colors: colorStops.map((s) => s.color),
                stops: colorStops.map((s) => s.position / 100),
                transparency: transparencies  // <-- FIXED: Added per-stop transparency array
            };

            // Add angle for linear gradients
            if (!isRadial) {
                gradientObj.angleDeg = pptxAngle;
            }

            // ✅ NEW: Add focal point for radial gradients
            if (isRadial && radialFocus) {
                gradientObj.path = radialFocus.path;
                // PowerPoint focal points: 0-100 where 50 is center
                // Convert our percentage to PowerPoint's expected format
                gradientObj.focusX = radialFocus.focusX;
                gradientObj.focusY = radialFocus.focusY;
            }

            shapeOptions.fill = {
                type: 'gradient',
                gradient: gradientObj
            };

        } else {
            // Fallback to solid fill with opacity
            let dominant = gradientInfo?.dominantColor || rgbToHex(bgStyle);
            shapeOptions.fill = {
                color: rgbToHex(dominant),
                transparency: calculatedTransparency
            };
        }
    } else if (style.backgroundColor && style.backgroundColor !== 'transparent') {
        // Normal solid fill (unchanged, but uses calculatedTransparency)
        const themeColorObj = createThemeColorObject(
            originalThemeColor,
            originalLumMod,
            originalLumOff,
            originalAlpha
        );
        if (themeColorObj) {
            shapeOptions.fill = { color: themeColorObj };
        } else if (style.backgroundColor) {
            shapeOptions.fill = {
                color: rgbToHex(style.backgroundColor),
                transparency: calculatedTransparency
            };
        }
    }



    // Store shape data for XML generation (Include flip information)
    const shapeData = {
        id: shapeId,
        x: x,
        y: y,
        w: w,
        h: h,
        rotation: rotation,
        flipH: flipHorizontal,
        flipV: flipVertical,
        fill: shapeOptions.fill,
        line: shapeOptions.line,
        borderRadius: borderRadiusValue,
        isSpecialShape: isSpecialShape,
        shapeType: getShapeTypeForXml(shapeId)
    };



    try {

        // ======= SVG FALLBACK FOR STANDARD SHAPES ======== hexagon shape not visible issue resolved.

        const svgProps = extractSvgPresentationProps(shapeElement);

        if (svgProps) {

            // ✅ FIX FILL (important)
            if (
                !shapeOptions.fill &&
                svgProps.fill &&
                svgProps.fill !== 'none' &&
                svgProps.fill !== 'transparent'
            ) {
                shapeOptions.fill = {
                    color: rgbToHex(svgProps.fill)
                };
            }

            // ✅ FIX STROKE (critical for your issue)
            if (
                !shapeOptions.line &&
                svgProps.stroke &&
                svgProps.stroke !== 'none' &&
                svgProps.stroke !== 'transparent' &&
                svgProps.strokeWidth > 0
            ) {
                shapeOptions.line = {
                    color: rgbToHex(svgProps.stroke),
                    width: svgProps.strokeWidth
                };
            }

            // ✅ FIX OPACITY
            if (
                shapeOptions.transparency === undefined &&
                svgProps.opacity !== undefined &&
                svgProps.opacity < 1
            ) {
                shapeOptions.transparency = Math.round((1 - svgProps.opacity) * 100);
            }
        }
        // All the existing switch cases remain the same...
        switch (shapeId) {
            // ========== CONNECTORS (These are elbow/bent/curved connectors) ==========
            case 'straightConnector1':
            case 'bentConnector2':
            case 'bentConnector3':
            case 'bentConnector4':
            case 'bentConnector5':
            case 'curvedConnector2':
            case 'curvedConnector3':
            case 'curvedConnector4':
            case 'curvedConnector5': {

                // Get connector name from element
                const connectorName = shapeElement.getAttribute('data-name') || objName;

                // ✅ PRIMARY METHOD: Read from data-connector-info HTML attribute
                const connectorInfoAttr = shapeElement.getAttribute('data-connector-info');
                let connectorData = null;

                if (connectorInfoAttr) {
                    try {
                        // Parse JSON from HTML attribute
                        const parsedData = JSON.parse(connectorInfoAttr);
                        const originalX = parseFloat(style.left || "0");
                        const originalY = parseFloat(style.top || "0");
                        const originalW = parseFloat(style.width || "0");
                        const originalH = parseFloat(style.height || "0");
                        // Reconstruct full connector data object
                        connectorData = {
                            shapeType: parsedData.shapeType || shapeId,
                            shapeName: connectorName,
                            position: {
                                x: originalX,  // ✅ Use original pixel values from HTML
                                y: originalY,
                                width: originalW,
                                height: originalH
                            },
                            strokeColor: parsedData.strokeColor,
                            strokeWidth: parsedData.strokeWidth,
                            dashType: parsedData.dashType || 'solid',
                            rotation: shapeOptions.rotate || 0,
                            flipH: shapeOptions.flipH || false,
                            flipV: shapeOptions.flipV || false,
                            lineEnds: parsedData.lineEnds || { headType: 'none', tailType: 'none' },
                            segments: parsedData.segments || []
                        };
                    } catch (error) {
                        console.error(`❌ Failed to parse data-connector-info:`, error);
                        connectorData = null;
                    }
                }

                // ⚠️ FALLBACK: If no data attribute, try global store
                if (!connectorData && global.connectorDataStore) {
                    connectorData = global.connectorDataStore.get(connectorName);
                }

                // ⚠️ LAST RESORT: Build from HTML element properties
                if (!connectorData) {

                    // Extract line properties from HTML/SVG
                    let lineProperties = getLineColorFromSvg(shapeElement);
                    let lineColor = lineProperties.color;
                    let lineWidth = lineProperties.width;

                    // Fallback to style properties if SVG properties not found
                    if (!lineColor) {
                        lineColor = style.borderColor || style.color || style.stroke || style.backgroundColor || '#000000';
                    }

                    if (!lineWidth) {
                        const cssStrokeWidth = style.strokeWidth || style.borderWidth;
                        if (cssStrokeWidth) {
                            lineWidth = parseFloat(cssStrokeWidth.replace('px', ''));
                        } else {
                            lineWidth = 2;
                        }
                    }

                    // Check for arrows in HTML
                    const arrowDivs = shapeElement.querySelectorAll('div[style*="border-left"][style*="border-top"]');
                    const hasArrow = arrowDivs.length > 0;

                    // Determine arrow position
                    let headType = 'none';
                    let tailType = 'none';

                    if (hasArrow) {
                        const arrowDiv = arrowDivs[0];
                        const arrowStyle = arrowDiv.style;
                        const arrowLeft = parseFloat(arrowStyle.left || '0');
                        const containerWidth = parseFloat(style.width || '100');

                        if (arrowLeft > containerWidth * 0.8) {
                            tailType = 'triangle';
                        } else {
                            headType = 'triangle';
                        }
                    }

                    // Create fallback connector data object
                    connectorData = {
                        shapeType: shapeId,
                        shapeName: connectorName,
                        position: {
                            x: shapeOptions.x * 72,
                            y: shapeOptions.y * 72,
                            width: shapeOptions.w * 72,
                            height: shapeOptions.h * 72
                        },
                        strokeColor: lineColor,
                        strokeWidth: lineWidth,
                        dashType: style.borderStyle === 'dashed' ? 'dash' :
                            style.borderStyle === 'dotted' ? 'dot' : 'solid',
                        rotation: shapeOptions.rotate || 0,
                        flipH: shapeOptions.flipH || false,
                        flipV: shapeOptions.flipV || false,
                        lineEnds: {
                            headType: headType,
                            tailType: tailType
                        },
                        segments: []
                    };

                }

                // Convert to PPTX using the connector data
                convertConnectorToPPTX(pptx, pptSlide, connectorData);

                break;
            }

            case 'actionButtonBackPrevious':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_BACK_OR_PREVIOUS, shapeOptions);
                break;
            case 'actionButtonBeginning':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_BEGINNING, shapeOptions);
                break;
            case 'actionButtonBlank':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_CUSTOM, shapeOptions);
                break;
            case 'actionButtonDocument':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_DOCUMENT, shapeOptions);
                break;
            case 'actionButtonEnd':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_END, shapeOptions);
                break;
            case 'actionButtonForwardNext':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_FORWARD_OR_NEXT, shapeOptions);
                break;
            case 'actionButtonHelp':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_HELP, shapeOptions);
                break;
            case 'actionButtonHome':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_HOME, shapeOptions);
                break;
            case 'actionButtonInformation':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_INFORMATION, shapeOptions);
                break;
            case 'actionButtonMovie':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_MOVIE, shapeOptions);
                break;
            case 'actionButtonReturn':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_RETURN, shapeOptions);
                break;
            case 'actionButtonSound':
                pptSlide.addShape(pptx.shapes.ACTION_BUTTON_SOUND, shapeOptions);
                break;

            // Basic Shapes
            case 'arc':
                pptSlide.addShape(pptx.shapes.ARC, shapeOptions);
                break;
            case 'wedgeRoundRectCallout':
                pptSlide.addShape(pptx.shapes.BALLOON, shapeOptions);
                break;
            case 'bentArrow':
                pptSlide.addShape(pptx.shapes.BENT_ARROW, shapeOptions);
                break;
            case 'bentUpArrow':
                pptSlide.addShape(pptx.shapes.BENT_UP_ARROW, shapeOptions);
                break;
            case 'bevel':
                pptSlide.addShape(pptx.shapes.BEVEL, shapeOptions);
                break;
            case 'blockArc':
                pptSlide.addShape(pptx.shapes.BLOCK_ARC, shapeOptions);
                break;
            case 'can': {
                const svgEl = shapeElement.querySelector('svg');

                // Fallback if svg missing
                if (!svgEl) {
                    pptSlide.addShape(pptx.shapes.CAN, {
                        ...shapeOptions,
                        hidden: false,
                        objectName: objName || shapeElement.getAttribute("data-name") || "can"
                    });
                    break;
                }

                const dataAdj = parseInt(shapeElement.getAttribute('data-adj') || '50000', 10);
                const wrapperOpacity = !isNaN(parseFloat(style.opacity)) ? parseFloat(style.opacity) : 1;

                // -----------------------------
                // helpers
                // -----------------------------
                const normalizeHex = (val) => {
                    if (!val) return null;
                    let c = String(val).trim();

                    if (c === 'none' || c === 'transparent') return null;

                    // rgb / rgba
                    if (c.startsWith('rgb')) {
                        return rgbToHex(c).replace('#', '').toUpperCase();
                    }

                    // #hex
                    if (c.startsWith('#')) {
                        c = c.replace('#', '').trim();
                        if (c.length === 3) {
                            c = c.split('').map(ch => ch + ch).join('');
                        }
                        if (/^[0-9A-Fa-f]{6}$/.test(c)) {
                            return c.toUpperCase();
                        }
                    }

                    // raw 6 hex
                    if (/^[0-9A-Fa-f]{6}$/.test(c)) {
                        return c.toUpperCase();
                    }

                    return null;
                };

                const parseOpacityValue = (val, fallback = 1) => {
                    const n = parseFloat(val);
                    return isNaN(n) ? fallback : n;
                };

                const getGradientById = (svgNode, id) => {
                    if (!id) return null;
                    const all = svgNode.querySelectorAll('linearGradient, radialGradient');
                    for (const node of all) {
                        if (node.getAttribute('id') === id) return node;
                    }
                    return null;
                };

                const parseGradientFillFromSvg = (svgNode, fillValue, wrapperOpacityVal) => {
                    if (!fillValue) return null;

                    const m = String(fillValue).match(/^url\(#(.+)\)$/);
                    if (!m) return null;

                    const gradId = m[1];
                    const gradEl = getGradientById(svgNode, gradId);
                    if (!gradEl) return null;

                    const tag = (gradEl.tagName || '').toLowerCase();
                    const stopEls = Array.from(gradEl.querySelectorAll('stop'));
                    if (!stopEls.length) return null;

                    const colors = [];
                    const stops = [];
                    const transparency = [];

                    for (const stopEl of stopEls) {
                        let stopColor =
                            stopEl.getAttribute('stop-color') ||
                            null;

                        if (!stopColor) {
                            const styleAttr = stopEl.getAttribute('style') || '';
                            const colorMatch = styleAttr.match(/stop-color\s*:\s*([^;]+)/i);
                            if (colorMatch) stopColor = colorMatch[1].trim();
                        }

                        const hex = normalizeHex(stopColor);
                        if (!hex) continue;

                        let offsetRaw = stopEl.getAttribute('offset') || '0%';
                        let offsetNum = 0;

                        if (offsetRaw.includes('%')) {
                            offsetNum = parseFloat(offsetRaw) / 100;
                        } else {
                            offsetNum = parseFloat(offsetRaw);
                            if (offsetNum > 1) offsetNum = offsetNum / 100;
                        }
                        if (isNaN(offsetNum)) offsetNum = 0;

                        let stopOpacity =
                            stopEl.getAttribute('stop-opacity');

                        if (stopOpacity == null || stopOpacity === '') {
                            const styleAttr = stopEl.getAttribute('style') || '';
                            const opacityMatch = styleAttr.match(/stop-opacity\s*:\s*([^;]+)/i);
                            if (opacityMatch) stopOpacity = opacityMatch[1].trim();
                        }

                        const alpha = parseOpacityValue(stopOpacity, 1) * wrapperOpacityVal;
                        const tr = Math.max(0, Math.min(100, Math.round((1 - alpha) * 100)));

                        colors.push(hex);
                        stops.push(Math.max(0, Math.min(1, offsetNum)));
                        transparency.push(tr);
                    }

                    if (!colors.length) return null;

                    if (tag === 'lineargradient') {
                        const x1 = parseFloat(gradEl.getAttribute('x1') || '0');
                        const y1 = parseFloat(gradEl.getAttribute('y1') || '0');
                        const x2 = parseFloat(gradEl.getAttribute('x2') || '0');
                        const y2 = parseFloat(gradEl.getAttribute('y2') || '100');

                        const dx = x2 - x1;
                        const dy = y2 - y1;

                        let angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
                        if (angleDeg < 0) angleDeg += 360;

                        return {
                            type: 'gradient',
                            gradient: {
                                type: 'linear',
                                angleDeg,
                                colors,
                                stops,
                                transparency
                            }
                        };
                    }

                    return {
                        type: 'gradient',
                        gradient: {
                            type: 'radial',
                            path: 'circle',
                            focusX: 50,
                            focusY: 50,
                            colors,
                            stops,
                            transparency
                        }
                    };
                };

                // -----------------------------
                // use actual generated SVG nodes
                // -----------------------------
                const ellipseEls = svgEl.querySelectorAll('ellipse');
                const pathEls = svgEl.querySelectorAll('path');

                const topEllipse = ellipseEls[0] || null;
                const bodyPath = pathEls[0] || null;
                const bottomPath = pathEls[1] || null;

                // Prefer body path for fill/stroke because that is your main can wall
                const primaryFill =
                    bodyPath?.getAttribute('fill') ||
                    topEllipse?.getAttribute('fill') ||
                    null;

                const primaryStroke =
                    bodyPath?.getAttribute('stroke') ||
                    topEllipse?.getAttribute('stroke') ||
                    null;

                const primaryStrokeWidth =
                    bodyPath?.getAttribute('stroke-width') ||
                    topEllipse?.getAttribute('stroke-width') ||
                    '0';

                const primaryStrokeOpacity =
                    bodyPath?.getAttribute('stroke-opacity') ||
                    topEllipse?.getAttribute('stroke-opacity') ||
                    '1';

                const gradientFill = parseGradientFillFromSvg(svgEl, primaryFill, wrapperOpacity);
                const solidFillHex = normalizeHex(primaryFill);
                const strokeHex = normalizeHex(primaryStroke);
                const strokeWidthNum = parseFloat(String(primaryStrokeWidth).replace('px', '')) || 0;
                const strokeOpacityNum = parseOpacityValue(primaryStrokeOpacity, 1) * wrapperOpacity;

                const canOptions = {
                    ...shapeOptions,
                    hidden: false,
                    objectName: objName || shapeElement.getAttribute("data-name") || "can"
                };

                // keep adj for later XML patching if needed
                canOptions.adjustPoint = dataAdj;

                // fill
                if (gradientFill) {
                    canOptions.fill = gradientFill;
                } else if (solidFillHex) {
                    canOptions.fill = {
                        color: solidFillHex,
                        transparency: Math.max(0, Math.min(100, Math.round((1 - wrapperOpacity) * 100)))
                    };
                } else {
                    // safe fallback
                    canOptions.fill = {
                        color: 'D9D9D9',
                        transparency: Math.max(0, Math.min(100, Math.round((1 - wrapperOpacity) * 100)))
                    };
                }

                // line
                if (!strokeHex || strokeWidthNum <= 0 || primaryStroke === 'none') {
                    canOptions.line = {
                        color: '000000',
                        transparency: 100,
                        width: 0
                    };
                } else {
                    canOptions.line = {
                        color: strokeHex,
                        width: strokeWidthNum,
                        transparency: Math.max(0, Math.min(100, Math.round((1 - strokeOpacityNum) * 100)))
                    };
                }

                pptSlide.addShape(pptx.shapes.CAN, canOptions);
                break;
            }
            case 'chartPlus':
                pptSlide.addShape(pptx.shapes.CHART_PLUS, shapeOptions);
                break;
            case 'chartStar':
                pptSlide.addShape(pptx.shapes.CHART_STAR, shapeOptions);
                break;
            case 'chartX':
                pptSlide.addShape(pptx.shapes.CHART_X, shapeOptions);
                break;
            case 'chevron': {
                const svgEl = shapeElement.querySelector('svg');
                const polygonEl = svgEl?.querySelector('polygon');

                // Fallback only if SVG/polygon is missing
                if (!svgEl || !polygonEl) {
                    pptSlide.addShape(pptx.shapes.CHEVRON, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                // Read actual SVG viewBox
                const viewBox = svgEl.getAttribute('viewBox') || '0 0 100 100';
                const vbParts = viewBox.trim().split(/\s+/).map(Number);
                const vbW = vbParts[2] || 100;
                const vbH = vbParts[3] || 100;

                // Read exact polygon points from HTML
                const rawPoints = (polygonEl.getAttribute('points') || '')
                    .trim()
                    .split(/\s+/)
                    .map(pt => pt.trim())
                    .filter(Boolean);

                if (rawPoints.length < 3) {
                    pptSlide.addShape(pptx.shapes.CHEVRON, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                // Convert polygon points -> SVG path
                let pathD = '';
                rawPoints.forEach((pt, index) => {
                    const [xStr, yStr] = pt.split(',');
                    const x = parseFloat(xStr);
                    const y = parseFloat(yStr);

                    if (Number.isNaN(x) || Number.isNaN(y)) return;

                    if (!pathD) {
                        pathD = `M ${x} ${y}`;
                    } else {
                        pathD += ` L ${x} ${y}`;
                    }
                });
                pathD += ' Z';

                // Convert SVG path -> pptx custGeom points
                const custGeomPoints = svgAddToSlide.convertSvgPathToPptxPoints(
                    pathD,
                    vbW,
                    vbH,
                    shapeOptions.w,
                    shapeOptions.h
                );

                if (!custGeomPoints || custGeomPoints.filter(p => !p.close).length < 3) {
                    pptSlide.addShape(pptx.shapes.CHEVRON, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                // Exact fill from SVG polygon
                const svgFill = polygonEl.getAttribute('fill');
                const fillColor = svgFill && svgFill !== 'none'
                    ? rgbToHex(svgFill).replace('#', '').toUpperCase()
                    : (
                        shapeOptions.fill?.color
                            ? String(shapeOptions.fill.color).replace('#', '').toUpperCase()
                            : '7CCA62'
                    );

                // Respect no-stroke from SVG
                const svgStroke = polygonEl.getAttribute('stroke');
                const svgStrokeWidth = parseFloat(polygonEl.getAttribute('stroke-width') || '0');

                let lineOptions = { color: fillColor, transparency: 100, width: 0 };

                if (svgStroke && svgStroke !== 'none' && svgStroke !== 'transparent' && svgStrokeWidth > 0) {
                    lineOptions = {
                        color: rgbToHex(svgStroke).replace('#', '').toUpperCase(),
                        width: svgStrokeWidth
                    };
                }

                pptSlide.addShape('custGeom', {
                    ...shapeOptions,
                    hidden: false,
                    objectName: objName || shapeElement.getAttribute('data-name') || 'chevron',
                    fill: {
                        color: fillColor,
                        transparency: calculatedTransparency || 0
                    },
                    line: lineOptions,
                    points: custGeomPoints
                });

                break;
            }
            case 'chord':
                pptSlide.addShape(pptx.shapes.CHORD, shapeOptions);
                break;
            case 'circularArrow':
                pptSlide.addShape(pptx.shapes.CIRCULAR_ARROW, shapeOptions);
                break;
            case 'cloud':
                pptSlide.addShape(pptx.shapes.CLOUD, shapeOptions);
                break;
            case 'cloudCallout':
                pptSlide.addShape(pptx.shapes.CLOUD_CALLOUT, shapeOptions);
                break;
            case 'corner':
                pptSlide.addShape(pptx.shapes.CORNER, shapeOptions);
                break;
            case 'cornerTabs':
                pptSlide.addShape(pptx.shapes.CORNER_TABS, shapeOptions);
                break;
            case 'plus':
                pptSlide.addShape(pptx.shapes.CROSS, shapeOptions);
                break;
            case 'cube': {
                const svgEl = shapeElement.querySelector('svg');

                // Fallback to preset if no SVG found
                if (!svgEl) {
                    pptSlide.addShape(pptx.shapes.CUBE, shapeOptions);
                    break;
                }

                // ── ViewBox dimensions ───────────────────────────────────────────────
                const viewBox = svgEl.getAttribute('viewBox') || '0 0 100 100';
                const vbParts = viewBox.trim().split(/\s+/).map(Number);
                const vbW = vbParts[2] || 100;
                const vbH = vbParts[3] || 100;

                // Shape position/size in inches (already computed above the switch)
                const baseX = shapeOptions.x;
                const baseY = shapeOptions.y;
                const baseW = shapeOptions.w;
                const baseH = shapeOptions.h;

                // ── Fill color from front face <rect> ────────────────────────────────
                const frontRect = svgEl.querySelector('rect');
                const fillAttr = frontRect?.getAttribute('fill') || '#808080';
                const fillHex = rgbToHex(fillAttr).replace('#', '').toUpperCase();

                // ── Stroke from front face <rect> ────────────────────────────────────
                const strokeAttr = frontRect?.getAttribute('stroke') || 'none';
                const strokeWidthPx = parseFloat(frontRect?.getAttribute('stroke-width') || '0');
                const strokeDashAttr = frontRect?.getAttribute('stroke-dasharray') || '';
                let lineOpts = null;
                if (strokeAttr !== 'none' && strokeWidthPx > 0) {
                    const dashMap = {
                        '8,4': 'dash', '8,4,2,4': 'dashDot', '2,4': 'dot',
                        '16,4': 'lgDash', '5,3': 'sysDash', '2,2': 'sysDot'
                    };
                    lineOpts = {
                        color: rgbToHex(strokeAttr).replace('#', '').toUpperCase(),
                        width: strokeWidthPx * 0.75,                           // px → pt
                        ...(strokeDashAttr ? { dashType: dashMap[strokeDashAttr.trim()] || 'solid' } : {})
                    };
                }

                // ── Shadow from <feDropShadow> ───────────────────────────────────────
                // Try both cases: browsers keep 'feDropShadow', jsdom lowercases it
                const feDropShadow = svgEl.querySelector('feDropShadow')
                    || svgEl.querySelector('fedropshadow');
                let shadowOpts = null;
                if (feDropShadow) {
                    const dx = parseFloat(feDropShadow.getAttribute('dx') || '0');
                    const dy = parseFloat(feDropShadow.getAttribute('dy') || '0');
                    const stdDev = parseFloat(
                        feDropShadow.getAttribute('stdDeviation') ||
                        feDropShadow.getAttribute('stddeviation') || '0'
                    );
                    const floodClr = feDropShadow.getAttribute('flood-color') || 'rgba(0,0,0,0.5)';
                    const rgba = floodClr.match(/rgba?\((\d+),(\d+),(\d+)(?:,([\d.]+))?\)/);
                    let sColor = '000000', sOpacity = 0.5;
                    if (rgba) {
                        sColor = [rgba[1], rgba[2], rgba[3]]
                            .map(v => parseInt(v).toString(16).padStart(2, '0'))
                            .join('').toUpperCase();
                        sOpacity = parseFloat(rgba[4] ?? '1');
                    }
                    const dist = Math.max(Math.round(Math.sqrt(dx * dx + dy * dy) * 0.75), 1);
                    const blur = Math.max(Math.round(stdDev * 2 * 0.75), 1);
                    let angle = Math.round(Math.atan2(dy, dx) * (180 / Math.PI));
                    if (angle < 0) angle += 360;
                    shadowOpts = { type: 'outer', color: sColor, blur, offset: dist, angle, opacity: sOpacity };
                }

                // ── Base options shared by all faces ─────────────────────────────────
                const baseOpts = {
                    x: baseX, y: baseY, w: baseW, h: baseH,
                    ...(lineOpts ? { line: lineOpts } : {}),
                    objectName: objName || 'cube'
                };

                // ── Front face (<rect>) drawn as custGeom ────────────────────────────
                // Rect → path string → convertSvgPathToPptxPoints (already in addSvgToSlide.js)
                if (frontRect) {
                    const rx = parseFloat(frontRect.getAttribute('x') || '0');
                    const ry = parseFloat(frontRect.getAttribute('y') || '0');
                    const rw = parseFloat(frontRect.getAttribute('width') || String(vbW));
                    const rh = parseFloat(frontRect.getAttribute('height') || String(vbH));

                    const rectPath = `M${rx} ${ry} L${rx + rw} ${ry} L${rx + rw} ${ry + rh} L${rx} ${ry + rh} Z`;
                    const frontPts = svgAddToSlide.convertSvgPathToPptxPoints(
                        rectPath, vbW, vbH, baseW, baseH
                    );

                    if (frontPts.length > 0) {
                        pptSlide.addShape('custGeom', {
                            ...baseOpts,
                            ...(shadowOpts ? { shadow: shadowOpts } : {}), // shadow only on front face
                            fill: { color: fillHex },
                            objectName: `${objName || 'cube'}_front`,
                            points: frontPts
                        });
                    }
                }

                // ── Top and right faces (<path> elements) ────────────────────────────
                // Read opacity directly from each path element (same values set in shapeHandler.js)
                // Convert SVG opacity → PPTX fill transparency (e.g. opacity 0.85 → 15% transparent)
                const pathEls = svgEl.querySelectorAll('path');
                pathEls.forEach((pathEl, i) => {
                    const d = pathEl.getAttribute('d') || '';
                    if (!d.trim()) return;

                    // Use the existing convertSvgPathToPptxPoints — handles M, L, C, Z correctly
                    const pts = svgAddToSlide.convertSvgPathToPptxPoints(
                        d, vbW, vbH, baseW, baseH
                    );
                    if (pts.filter(p => !p.close).length < 3) return;

                    // Read opacity from element (0.85 top, 0.7 right from shapeHandler.js)
                    const faceOpacity = parseFloat(pathEl.getAttribute('opacity') || '1');
                    const transparency = Math.round((1 - faceOpacity) * 100); // 0.85→15, 0.7→30

                    pptSlide.addShape('custGeom', {
                        ...baseOpts,
                        fill: { color: fillHex, transparency },
                        objectName: `${objName || 'cube'}_face${i + 1}`,
                        points: pts
                    });
                });

                break;
            }
            case 'curvedDownArrow':
                pptSlide.addShape(pptx.shapes.CURVED_DOWN_ARROW, shapeOptions);
                break;
            case 'ellipseRibbon':
                pptSlide.addShape(pptx.shapes.CURVED_DOWN_RIBBON, shapeOptions);
                break;
            case 'curvedLeftArrow':
                pptSlide.addShape(pptx.shapes.CURVED_LEFT_ARROW, shapeOptions);
                break;
            case 'curvedRightArrow':
                pptSlide.addShape(pptx.shapes.CURVED_RIGHT_ARROW, shapeOptions);
                break;
            case 'curvedUpArrow':
                pptSlide.addShape(pptx.shapes.CURVED_UP_ARROW, shapeOptions);
                break;
            case 'ellipseRibbon2':
                pptSlide.addShape(pptx.shapes.CURVED_UP_RIBBON, shapeOptions);
                break;
            case 'decagon':
                pptSlide.addShape(pptx.shapes.DECAGON, shapeOptions);
                break;
            case 'diagStripe':
                pptSlide.addShape(pptx.shapes.DIAGONAL_STRIPE, shapeOptions);
                break;
            case 'diamond':
                pptSlide.addShape(pptx.shapes.DIAMOND, shapeOptions);
                break;
            case 'dodecagon':
                pptSlide.addShape(pptx.shapes.DODECAGON, shapeOptions);
                break;
            case 'donut':
                pptSlide.addShape(pptx.shapes.DONUT, shapeOptions);
                break;
            case 'bracePair':
                pptSlide.addShape(pptx.shapes.DOUBLE_BRACE, shapeOptions);
                break;
            case 'bracketPair':
                pptSlide.addShape(pptx.shapes.DOUBLE_BRACKET, shapeOptions);
                break;
            case 'doubleWave':
                pptSlide.addShape(pptx.shapes.DOUBLE_WAVE, shapeOptions);
                break;
            case 'downArrow':
                pptSlide.addShape(pptx.shapes.DOWN_ARROW, shapeOptions);
                break;
            case 'downArrowCallout':
                pptSlide.addShape(pptx.shapes.DOWN_ARROW_CALLOUT, shapeOptions);
                break;
            case 'ribbon':
                pptSlide.addShape(pptx.shapes.DOWN_RIBBON, shapeOptions);
                break;
            case 'irregularSeal1':
                pptSlide.addShape(pptx.shapes.EXPLOSION1, shapeOptions);
                break;
            case 'irregularSeal2':
                pptSlide.addShape(pptx.shapes.EXPLOSION2, shapeOptions);
                break;

            // Flowchart Shapes
            case 'flowChartAlternateProcess':
                pptSlide.addShape(pptx.shapes.FLOWCHART_ALTERNATE_PROCESS, shapeOptions);
                break;
            case 'flowChartPunchedCard':
                pptSlide.addShape(pptx.shapes.FLOWCHART_CARD, shapeOptions);
                break;
            case 'flowChartCollate':
                pptSlide.addShape(pptx.shapes.FLOWCHART_COLLATE, shapeOptions);
                break;
            case 'flowChartConnector':
                pptSlide.addShape(pptx.shapes.FLOWCHART_CONNECTOR, shapeOptions);
                break;
            case 'flowChartInputOutput':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DATA, shapeOptions);
                break;
            case 'flowChartDecision':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DECISION, shapeOptions);
                break;
            case 'flowChartDelay':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DELAY, shapeOptions);
                break;
            case 'flowChartMagneticDrum':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DIRECT_ACCESS_STORAGE, shapeOptions);
                break;
            case 'flowChartDisplay':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DISPLAY, shapeOptions);
                break;
            case 'flowChartDocument': {
                const svgEl = shapeElement.querySelector('svg');
                const pathEl = svgEl?.querySelector('path');

                // Fallback if SVG/path missing
                if (!svgEl || !pathEl) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_DOCUMENT, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                const rawStyleAttr = shapeElement.getAttribute('style') || '';

                const transformMatch = rawStyleAttr.match(/transform\s*:\s*([^;]+)/i);
                const rawTransform = transformMatch ? transformMatch[1].trim() : '';

                const rotationMatchLocal = rawTransform.match(/rotate\(\s*([-\d.]+)deg\s*\)/i);
                const scaleXMatchLocal = rawTransform.match(/scaleX\(\s*(-?\d*\.?\d+)\s*\)/i);
                const scaleYMatchLocal = rawTransform.match(/scaleY\(\s*(-?\d*\.?\d+)\s*\)/i);

                let localRotate = 0;
                let localFlipH = false;
                let localFlipV = false;

                if (rotationMatchLocal && !isNaN(parseFloat(rotationMatchLocal[1]))) {
                    localRotate = parseFloat(rotationMatchLocal[1]);
                }

                if (scaleXMatchLocal && parseFloat(scaleXMatchLocal[1]) < 0) {
                    localFlipH = true;
                }

                if (scaleYMatchLocal && parseFloat(scaleYMatchLocal[1]) < 0) {
                    localFlipV = true;
                }

                localRotate = ((localRotate % 360) + 360) % 360;

                const flowDocShapeOptions = {
                    ...shapeOptions,
                    rotate: localRotate
                };

                if (localFlipH) flowDocShapeOptions.flipH = true;
                else delete flowDocShapeOptions.flipH;

                if (localFlipV) flowDocShapeOptions.flipV = true;
                else delete flowDocShapeOptions.flipV;

                const viewBox = svgEl.getAttribute('viewBox') || '0 0 100 100';
                const vbParts = viewBox.trim().split(/\s+/).map(Number);

                const vbW = (Number.isFinite(vbParts[2]) && vbParts[2] > 0) ? vbParts[2] : 100;
                const vbH = (Number.isFinite(vbParts[3]) && vbParts[3] > 0) ? vbParts[3] : 100;

                const pathD = (pathEl.getAttribute('d') || '').trim();

                if (!pathD) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_DOCUMENT, {
                        ...flowDocShapeOptions,
                        hidden: false
                    });
                    break;
                }

                const custGeomPoints = svgAddToSlide.convertSvgPathToPptxPoints(
                    pathD,
                    vbW,
                    vbH,
                    flowDocShapeOptions.w,
                    flowDocShapeOptions.h
                );

                if (!custGeomPoints || custGeomPoints.filter(p => !p.close).length < 3) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_DOCUMENT, {
                        ...flowDocShapeOptions,
                        hidden: false
                    });
                    break;
                }

                const svgFill = pathEl.getAttribute('fill');
                let fillOptions = {
                    color: 'E7E6E6'
                };

                if (svgFill && svgFill !== 'none' && svgFill !== 'transparent') {
                    fillOptions = {
                        color: rgbToHex(svgFill).replace('#', '').toUpperCase(),
                        transparency: calculatedTransparency || 0
                    };
                }

                // Preserve original theme color if available
                if (
                    originalThemeColor &&
                    originalThemeColor !== 'null' &&
                    originalThemeColor !== 'undefined'
                ) {
                    const themeColorObj = createThemeColorObject(
                        originalThemeColor,
                        originalLumMod,
                        originalLumOff,
                        originalAlpha
                    );

                    if (themeColorObj) {
                        fillOptions = { color: themeColorObj };
                    }
                }

                const svgStroke = pathEl.getAttribute('stroke');
                const svgStrokeWidth = parseFloat(pathEl.getAttribute('stroke-width') || '0');

                let lineOptions = { color: '000000', transparency: 100, width: 0 };

                if (
                    svgStroke &&
                    svgStroke !== 'none' &&
                    svgStroke !== 'transparent' &&
                    !isNaN(svgStrokeWidth) &&
                    svgStrokeWidth > 0
                ) {
                    lineOptions = {
                        color: rgbToHex(svgStroke).replace('#', '').toUpperCase(),
                        width: svgStrokeWidth
                    };
                } else if (flowDocShapeOptions.line) {
                    lineOptions = flowDocShapeOptions.line;
                }

                const pathOpacityRaw = pathEl.getAttribute('opacity');
                if (pathOpacityRaw !== null && pathOpacityRaw !== '') {
                    const pathOpacity = parseFloat(pathOpacityRaw);
                    if (!isNaN(pathOpacity) && pathOpacity >= 0 && pathOpacity < 1) {
                        fillOptions = {
                            ...fillOptions,
                            transparency: Math.round((1 - pathOpacity) * 100)
                        };
                    }
                }

                pptSlide.addShape('custGeom', {
                    ...flowDocShapeOptions,
                    hidden: false,
                    objectName: objName || shapeElement.getAttribute('data-name') || 'flowChartDocument',
                    fill: fillOptions,
                    line: lineOptions,
                    points: custGeomPoints
                });

                break;
            }
            case 'flowChartExtract':
                pptSlide.addShape(pptx.shapes.FLOWCHART_EXTRACT, shapeOptions);
                break;
            case 'flowChartInternalStorage':
                pptSlide.addShape(pptx.shapes.FLOWCHART_INTERNAL_STORAGE, shapeOptions);
                break;
            case 'flowChartMagneticDisk':
                pptSlide.addShape(pptx.shapes.FLOWCHART_MAGNETIC_DISK, shapeOptions);
                break;
            case 'flowChartManualInput':
                pptSlide.addShape(pptx.shapes.FLOWCHART_MANUAL_INPUT, shapeOptions);
                break;
            case 'flowChartManualOperation':
                pptSlide.addShape(pptx.shapes.FLOWCHART_MANUAL_OPERATION, shapeOptions);
                break;
            case 'flowChartMerge':
                pptSlide.addShape(pptx.shapes.FLOWCHART_MERGE, shapeOptions);
                break;
            case 'flowChartMultidocument':
                pptSlide.addShape(pptx.shapes.FLOWCHART_MULTIDOCUMENT, shapeOptions);
                break;
            case 'flowChartOfflineStorage':
                pptSlide.addShape(pptx.shapes.FLOWCHART_OFFLINE_STORAGE, shapeOptions);
                break;
            case 'flowChartOffpageConnector':
                pptSlide.addShape(pptx.shapes.FLOWCHART_OFFPAGE_CONNECTOR, shapeOptions);
                break;
            case 'flowChartOr':
                pptSlide.addShape(pptx.shapes.FLOWCHART_OR, shapeOptions);
                break;
            case 'flowChartPredefinedProcess':
                pptSlide.addShape(pptx.shapes.FLOWCHART_PREDEFINED_PROCESS, shapeOptions);
                break;
            case 'flowChartPreparation':
                pptSlide.addShape(pptx.shapes.FLOWCHART_PREPARATION, shapeOptions);
                break;
            case 'flowChartProcess':
                pptSlide.addShape(pptx.shapes.FLOWCHART_PROCESS, shapeOptions);
                break;
            case 'flowChartPunchedTape':
                pptSlide.addShape(pptx.shapes.FLOWCHART_PUNCHED_TAPE, shapeOptions);
                break;
            case 'flowChartMagneticTape':
                pptSlide.addShape(pptx.shapes.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE, shapeOptions);
                break;
            case 'flowChartSort':
                pptSlide.addShape(pptx.shapes.FLOWCHART_SORT, shapeOptions);
                break;
            case 'flowChartOnlineStorage': {
                const rawStyle = shapeElement.getAttribute('style') || '';
                const svgEl = shapeElement.querySelector('svg');
                const pathEl = svgEl?.querySelector('path');

                if (!svgEl || !pathEl) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_STORED_DATA, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                // ── helpers ─────────────────────────────────────────────────────
                const getPx = (prop) => {
                    const m = rawStyle.match(new RegExp(prop + '\\s*:\\s*([0-9.-]+)px', 'i'));
                    return m ? parseFloat(m[1]) : 0;
                };

                const toHex = (val) => {
                    if (!val) return null;
                    const s = String(val).trim();
                    if (!s || s === 'none' || s === 'transparent') return null;

                    if (/^#[0-9a-fA-F]{6}$/.test(s)) return s.replace('#', '').toUpperCase();
                    if (/^[0-9a-fA-F]{6}$/.test(s)) return s.toUpperCase();

                    const rgb = s.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
                    if (rgb) {
                        return [rgb[1], rgb[2], rgb[3]]
                            .map(v => Math.max(0, Math.min(255, parseInt(v, 10))).toString(16).padStart(2, '0'))
                            .join('')
                            .toUpperCase();
                    }

                    const rgba = s.match(/^rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([0-9.]+)\s*\)$/i);
                    if (rgba) {
                        return [rgba[1], rgba[2], rgba[3]]
                            .map(v => Math.max(0, Math.min(255, parseInt(v, 10))).toString(16).padStart(2, '0'))
                            .join('')
                            .toUpperCase();
                    }

                    return null;
                };

                const mapJoin = (joinVal) => {
                    const v = String(joinVal || '').toLowerCase();
                    if (v === 'miter') return 'miter';
                    if (v === 'bevel') return 'bevel';
                    return 'round';
                };

                const mapCap = (capVal) => {
                    const v = String(capVal || '').toLowerCase();
                    if (v === 'round') return 'round';
                    if (v === 'square') return 'square';
                    return 'flat';
                };

                const parseDashType = (pathEl) => {
                    const dashArray = pathEl.getAttribute('stroke-dasharray');
                    if (!dashArray) return undefined;

                    const arr = dashArray
                        .split(/[ ,]+/)
                        .map(v => parseFloat(v))
                        .filter(v => !isNaN(v));

                    if (!arr.length) return undefined;
                    if (arr.length === 2) {
                        if (arr[0] <= 2) return 'sysDot';
                        return 'dash';
                    }
                    if (arr.length === 4) return 'dashDot';
                    if (arr.length >= 6) return 'lgDashDotDot';
                    return undefined;
                };

                // ── exact position / size from HTML ───────────────────────────
                const fosLeft = getPx('left');
                const fosTop = getPx('top');
                const fosWidth = getPx('width');
                const fosHeight = getPx('height');

                const fosX = fosLeft / 72;
                const fosY = fosTop / 72;
                const fosW = fosWidth / 72;
                const fosH = fosHeight / 72;

                // ── exact transform from HTML style string ────────────────────
                const fosTransform = rawStyle.match(/transform\s*:\s*([^;]+)/i)?.[1]?.trim() || '';

                const fosRotMatch = fosTransform.match(/rotate\(\s*([-\d.]+)deg\s*\)/i);
                const fosRotation = fosRotMatch ? parseFloat(fosRotMatch[1]) : 0;

                const fosScaleXMatch = fosTransform.match(/scaleX\(\s*([-\d.]+)\s*\)/i);
                const fosScaleYMatch = fosTransform.match(/scaleY\(\s*([-\d.]+)\s*\)/i);
                const fosScaleMatch = fosTransform.match(/scale\(\s*([-\d.]+)(?:\s*,\s*([-\d.]+))?\s*\)/i);

                let fosFlipH = false;
                let fosFlipV = false;

                if (fosScaleXMatch && parseFloat(fosScaleXMatch[1]) < 0) fosFlipH = true;
                if (fosScaleYMatch && parseFloat(fosScaleYMatch[1]) < 0) fosFlipV = true;

                if (fosScaleMatch) {
                    const sx = parseFloat(fosScaleMatch[1]);
                    const sy = fosScaleMatch[2] !== undefined ? parseFloat(fosScaleMatch[2]) : sx;
                    if (!isNaN(sx) && sx < 0) fosFlipH = true;
                    if (!isNaN(sy) && sy < 0) fosFlipV = true;
                }

                // ── original SVG path, no pre-transform ───────────────────────
                const d = pathEl.getAttribute('d') || '';
                if (!d.trim()) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_STORED_DATA, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                const viewBox = (svgEl.getAttribute('viewBox') || `0 0 ${fosWidth || 100} ${fosHeight || 100}`)
                    .trim()
                    .split(/\s+/)
                    .map(Number);

                const vbX = viewBox[0] || 0;
                const vbY = viewBox[1] || 0;
                const vbW = viewBox[2] || fosWidth || 100;
                const vbH = viewBox[3] || fosHeight || 100;

                let fosPts = svgAddToSlide.convertSvgPathToPptxPoints(
                    d,
                    vbW,
                    vbH,
                    fosW,
                    fosH
                );

                if (!fosPts || fosPts.filter(p => !p.close).length < 3) {
                    pptSlide.addShape(pptx.shapes.FLOWCHART_STORED_DATA, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                // adjust if viewBox origin is not 0,0
                if (vbX !== 0 || vbY !== 0) {
                    const sx = fosW / vbW;
                    const sy = fosH / vbH;

                    fosPts = fosPts.map(pt => {
                        if (pt.close) return pt;
                        return {
                            ...pt,
                            x: pt.x - (vbX * sx),
                            y: pt.y - (vbY * sy)
                        };
                    });
                }

                // ── fill ──────────────────────────────────────────────────────
                const fosFillHex = toHex(pathEl.getAttribute('fill')) || 'FFFFFF';
                const fosFill = { color: fosFillHex };

                // ── line ──────────────────────────────────────────────────────
                const fosStroke = pathEl.getAttribute('stroke');
                const fosStrokeHex = toHex(fosStroke);
                const fosStrokeWidth = parseFloat(pathEl.getAttribute('stroke-width') || '0');
                const fosStrokeJoin = pathEl.getAttribute('stroke-linejoin') || 'round';
                const fosStrokeCap = pathEl.getAttribute('stroke-linecap') || 'butt';

                let fosLine = {
                    color: fosFillHex,
                    transparency: 100,
                    width: 0
                };

                if (fosStrokeHex && fosStroke !== 'none' && !isNaN(fosStrokeWidth) && fosStrokeWidth > 0) {
                    fosLine = {
                        color: fosStrokeHex,
                        width: fosStrokeWidth,
                        join: mapJoin(fosStrokeJoin),
                        cap: mapCap(fosStrokeCap)
                    };

                    const dashType = parseDashType(pathEl);
                    if (dashType) fosLine.dashType = dashType;
                }

                // ── final shape options: transform goes to XML xfrm ──────────
                const fosShapeOptions = {
                    x: fosX,
                    y: fosY,
                    w: fosW,
                    h: fosH,
                    rotate: fosRotation,
                    fill: fosFill,
                    line: fosLine,
                    objectName: objName || shapeElement.getAttribute('data-name') || 'flowChartOnlineStorage',
                    hidden: false,
                    points: fosPts
                };

                if (fosFlipH) fosShapeOptions.flipH = true;
                if (fosFlipV) fosShapeOptions.flipV = true;

                if (shapeOptions.shadow) fosShapeOptions.shadow = shapeOptions.shadow;
                if (shapeOptions.glow) fosShapeOptions.glow = shapeOptions.glow;

                pptSlide.addShape('custGeom', fosShapeOptions);
                break;
            }
            case 'flowChartSummingJunction':
                pptSlide.addShape(pptx.shapes.FLOWCHART_SUMMING_JUNCTION, shapeOptions);
                break;
            case 'flowChartTerminator':
                pptSlide.addShape(pptx.shapes.FLOWCHART_TERMINATOR, shapeOptions);
                break;

            // Geometric Shapes
            case 'folderCorner':
                pptSlide.addShape(pptx.shapes.FOLDED_CORNER, shapeOptions);
                break;
            case 'frame':
                pptSlide.addShape(pptx.shapes.FRAME, shapeOptions);
                break;
            case 'funnel':
                pptSlide.addShape(pptx.shapes.FUNNEL, shapeOptions);
                break;
            case 'gear6':
                pptSlide.addShape(pptx.shapes.GEAR_6, shapeOptions);
                break;
            case 'gear9':
                pptSlide.addShape(pptx.shapes.GEAR_9, shapeOptions);
                break;
            case 'halfFrame':
                pptSlide.addShape(pptx.shapes.HALF_FRAME, shapeOptions);
                break;
            case 'heart':
                pptSlide.addShape(pptx.shapes.HEART, shapeOptions);
                break;
            case 'heptagon':
                pptSlide.addShape(pptx.shapes.HEPTAGON, shapeOptions);
                break;
            case 'hexagon':
                pptSlide.addShape(pptx.shapes.HEXAGON, shapeOptions);
                break;
            case 'horizontalScroll':
                pptSlide.addShape(pptx.shapes.HORIZONTAL_SCROLL, shapeOptions);
                break;
            case 'triangle': {
                const triSvgEl = shapeElement.querySelector('svg');
                if (!triSvgEl) {
                    pptSlide.addShape(pptx.shapes.ISOSCELES_TRIANGLE, { ...shapeOptions, hidden: false });
                    break;
                }

                // ── 1. Position / size (px → inches, 72 DPI) ─────────────────────
                const triStyleStr = shapeElement.getAttribute('style') || '';
                const triGetPx = (prop) => {
                    const m = triStyleStr.match(new RegExp(prop + '\\s*:\\s*([0-9.-]+)px', 'i'));
                    return m ? parseFloat(m[1]) : null;
                };
                const triLeft = triGetPx('left') ?? parseFloat(style.left || '0');
                const triTop = triGetPx('top') ?? parseFloat(style.top || '0');
                const triWidth = triGetPx('width') ?? parseFloat(style.width || '0');
                const triHeight = triGetPx('height') ?? parseFloat(style.height || '0');
                const triX = triLeft / 72, triY = triTop / 72;
                const triW = triWidth / 72, triH = triHeight / 72;

                // ── 2. Rotation — parse rotate() from transform (supports fractions) ─
                const triTransform = triStyleStr.match(/transform\s*:\s*([^;]+)/i)?.[1]?.trim() || '';
                const triRotMatch = triTransform.match(/rotate\(\s*([-\d.]+)deg\s*\)/i);
                const triRotation = triRotMatch ? parseFloat(triRotMatch[1]) : 0;

                // ── 3. Flip — scaleX(-1) / scaleY(-1) in CSS transform ───────────
                const triScaleX = triTransform.match(/scaleX\(\s*([-\d.]+)\s*\)/i);
                const triScaleY = triTransform.match(/scaleY\(\s*([-\d.]+)\s*\)/i);
                const triFlipH = triScaleX ? parseFloat(triScaleX[1]) < 0 : false;
                const triFlipV = triScaleY ? parseFloat(triScaleY[1]) < 0 : false;

                // ── 4. Opacity → fill transparency ────────────────────────────────
                // Pre-switch block gates opacity on hasVisibleBackground (always
                // false for SVG triangles), so we read it directly from the style string.
                const triOpacityRaw = triStyleStr.match(/(?:^|;)\s*opacity\s*:\s*([\d.]+)/i)?.[1]
                    ?? style.opacity ?? '1';
                const triOpacity = Math.min(1, Math.max(0, parseFloat(triOpacityRaw) || 1));
                const triFillTrsp = Math.round((1 - triOpacity) * 100);

                // ── 5. SVG viewBox ────────────────────────────────────────────────
                const triVbParts = (triSvgEl.getAttribute('viewBox') || `0 0 ${triWidth} ${triHeight}`)
                    .trim().split(/\s+/).map(Number);
                const triVbW = triVbParts[2] || triWidth;
                const triVbH = triVbParts[3] || triHeight;

                // ── 6. Filter type detection ──────────────────────────────────────
                // Two patterns used in triangle SVGs:
                // A) <feDropShadow>  → outer drop shadow (Triangle 5, 11)
                // B) feGaussianBlur + feFlood + feComposite(operator=atop)
                //    → white or black overlay composited ATOP the shape
                //    → creates the 3D highlight / shadow effect (Triangle 7,8,14,15)
                //    → approximated by blending flood-color into fill at flood-opacity

                // Pattern A: feDropShadow
                const triFeDS = triSvgEl.querySelector('feDropShadow')
                    || triSvgEl.querySelector('fedropshadow');
                let triShadow = null;
                if (triFeDS) {
                    const sdx = parseFloat(triFeDS.getAttribute('dx') || '0');
                    const sdy = parseFloat(triFeDS.getAttribute('dy') || '0');
                    const stdDev = parseFloat(
                        triFeDS.getAttribute('stdDeviation') ||
                        triFeDS.getAttribute('stddeviation') || '0'
                    );
                    const floodClr = triFeDS.getAttribute('flood-color') || '#000000';
                    const floodOp = parseFloat(triFeDS.getAttribute('flood-opacity') || '1');

                    const rgbaM = floodClr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
                    let sColor = '000000', sOpacity = floodOp;
                    if (rgbaM) {
                        sColor = [rgbaM[1], rgbaM[2], rgbaM[3]]
                            .map(v => parseInt(v).toString(16).padStart(2, '0'))
                            .join('').toUpperCase();
                        if (rgbaM[4]) sOpacity = parseFloat(rgbaM[4]) * floodOp;
                    } else if (floodClr.startsWith('#')) {
                        sColor = floodClr.replace('#', '').toUpperCase();
                    }
                    const dist = Math.max(Math.round(Math.sqrt(sdx * sdx + sdy * sdy) * 0.75), 1);
                    const blur = Math.max(Math.round(stdDev * 2 * 0.75), 1);
                    let angle = Math.round(Math.atan2(sdy, sdx) * (180 / Math.PI));
                    if (angle < 0) angle += 360;
                    triShadow = { type: 'outer', color: sColor, blur, offset: dist, angle, opacity: sOpacity };
                }

                // Pattern B: feGaussianBlur composite highlight/shadow overlay
                // Detect last feComposite with operator="atop" (composites overlay atop source).
                // The feFlood before it defines the overlay color and opacity.
                let triOverlayR = -1, triOverlayG = -1, triOverlayB = -1, triOverlayOp = 0;
                const triFeComposites = triSvgEl.querySelectorAll('feComposite, fecomposite');
                let triAtopComposite = null;
                triFeComposites.forEach(fc => {
                    const op = fc.getAttribute('operator') || '';
                    if (op === 'atop') triAtopComposite = fc;
                });
                if (triAtopComposite && !triFeDS) {
                    // Walk backwards to find the preceding feFlood
                    const triFeFloods = triSvgEl.querySelectorAll('feFlood, feflood');
                    const lastFlood = triFeFloods[triFeFloods.length - 1];
                    if (lastFlood) {
                        const fc = lastFlood.getAttribute('flood-color') || lastFlood.getAttribute('floodColor') || '#000000';
                        triOverlayOp = parseFloat(lastFlood.getAttribute('flood-opacity') || lastFlood.getAttribute('floodOpacity') || '0');
                        const fcM = fc.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
                        if (fcM) {
                            triOverlayR = parseInt(fcM[1]);
                            triOverlayG = parseInt(fcM[2]);
                            triOverlayB = parseInt(fcM[3]);
                        } else if (fc.startsWith('#')) {
                            const h = fc.replace('#', '').padStart(6, '0');
                            triOverlayR = parseInt(h.slice(0, 2), 16);
                            triOverlayG = parseInt(h.slice(2, 4), 16);
                            triOverlayB = parseInt(h.slice(4, 6), 16);
                        } else if (/^white$/i.test(fc)) { triOverlayR = triOverlayG = triOverlayB = 255; }
                        else if (/^black$/i.test(fc)) { triOverlayR = triOverlayG = triOverlayB = 0; }
                    }
                }
                // Helper: blend base fill color with overlay (simulates feComposite atop)
                function applyOverlay(hexFill) {
                    if (triOverlayR < 0 || triOverlayOp <= 0) return hexFill;
                    const bR = parseInt(hexFill.slice(0, 2), 16);
                    const bG = parseInt(hexFill.slice(2, 4), 16);
                    const bB = parseInt(hexFill.slice(4, 6), 16);
                    const blend = (b, o) => Math.min(255, Math.max(0, Math.round(b * (1 - triOverlayOp) + o * triOverlayOp)));
                    return [blend(bR, triOverlayR), blend(bG, triOverlayG), blend(bB, triOverlayB)]
                        .map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
                }

                // ── 7. adj (apex position) ────────────────────────────────────────
                const triAdjRaw = shapeElement.getAttribute('data-adj');
                const triAdj = triAdjRaw ? parseInt(triAdjRaw, 10) : 50000;

                // ── 8. Shared base options ────────────────────────────────────────
                const triBase = {
                    x: triX, y: triY, w: triW, h: triH,
                    rotate: triRotation,
                    flipH: triFlipH,
                    flipV: triFlipV,
                    hidden: false
                };

                // ── 9. Process each polygon in the SVG ───────────────────────────
                // Some triangles contain two polygons:
                //   .triangle-outer       → filled body (use ISOSCELES_TRIANGLE preset)
                //   .triangle-inner-border → border-only inset (use custGeom, exact points)
                // Single-polygon triangles contain only .triangle-outer.
                const triAllPolygons = Array.from(triSvgEl.querySelectorAll('polygon'));
                const triName = objName || shapeElement.getAttribute('data-name') || 'triangle';

                triAllPolygons.forEach((triPoly, triPolyIdx) => {
                    const polyFill = triPoly.getAttribute('fill') || 'none';
                    const polyStroke = triPoly.getAttribute('stroke') || 'none';
                    const polyStrokeW = parseFloat(triPoly.getAttribute('stroke-width') || '0');
                    const isOuter = triPoly.classList.contains('triangle-outer') || triPolyIdx === 0;
                    const isBorder = triPoly.classList.contains('triangle-inner-border');
                    const polyName = triAllPolygons.length > 1
                        ? `${triName}_${isOuter ? 'outer' : 'border'}`
                        : triName;

                    // Build line options for this polygon
                    let polyLineOpts;
                    if (polyStroke !== 'none' && polyStroke !== 'transparent' && polyStrokeW > 0) {
                        polyLineOpts = {
                            color: rgbToHex(polyStroke).replace('#', '').toUpperCase(),
                            width: polyStrokeW * 0.75   // px → pt
                        };
                    } else {
                        // Suppress pptxgenjs default border
                        const noStrokeColor = polyFill !== 'none'
                            ? rgbToHex(polyFill).replace('#', '').toUpperCase()
                            : 'FFFFFF';
                        polyLineOpts = { color: noStrokeColor, transparency: 100, width: 0 };
                    }

                    if (isBorder || polyFill === 'none') {
                        // ── Border-only polygon → custGeom with exact points ──────
                        // This preserves the precise inner border inset regardless
                        // of rotation/flip (points are in viewBox local space).
                        const rawPtStr = triPoly.getAttribute('points') || '';
                        const rawPts = rawPtStr.trim().split(/\s+/).filter(Boolean);
                        if (rawPts.length < 3) return;
                        let borderPath = '';
                        rawPts.forEach((pt, i) => {
                            const [px, py] = pt.split(',').map(Number);
                            if (Number.isNaN(px) || Number.isNaN(py)) return;
                            borderPath += i === 0 ? `M ${px} ${py}` : ` L ${px} ${py}`;
                        });
                        borderPath += ' Z';
                        const borderPts = svgAddToSlide.convertSvgPathToPptxPoints(
                            borderPath, triVbW, triVbH, triW, triH
                        );
                        if (!borderPts || borderPts.filter(p => !p.close).length < 3) return;
                        pptSlide.addShape('custGeom', {
                            ...triBase,
                            objectName: polyName,
                            fill: { color: 'FFFFFF', transparency: 100 },
                            line: polyLineOpts,
                            points: borderPts
                        });

                    } else {
                        // ── Filled polygon → ISOSCELES_TRIANGLE preset ────────────
                        // Apply feGaussianBlur overlay blend to get correct face color.
                        let polyFillHex = rgbToHex(polyFill).replace('#', '').toUpperCase();
                        polyFillHex = applyOverlay(polyFillHex);

                        const presetOpts = {
                            ...triBase,
                            objectName: polyName,
                            fill: { color: polyFillHex, transparency: triFillTrsp },
                            line: polyLineOpts,
                            // Shadow and glow apply only to the primary (outer) polygon
                            ...(isOuter && triShadow ? { shadow: triShadow } : {}),
                            ...(isOuter && glowOptions ? { glow: glowOptions } : {})
                        };
                        // adj only on the primary polygon and when non-default
                        if (isOuter && triAdj !== 50000) presetOpts.adj = triAdj;

                        pptSlide.addShape(pptx.shapes.ISOSCELES_TRIANGLE, presetOpts);
                    }
                });

                break;
            }
            case 'leftArrow':
                pptSlide.addShape(pptx.shapes.LEFT_ARROW, shapeOptions);
                break;
            case 'leftArrowCallout':
                pptSlide.addShape(pptx.shapes.LEFT_ARROW_CALLOUT, shapeOptions);
                break;
            case 'leftBrace':
                pptSlide.addShape(pptx.shapes.LEFT_BRACE, shapeOptions);
                break;
            case 'leftBracket':
                pptSlide.addShape(pptx.shapes.LEFT_BRACKET, shapeOptions);
                break;
            case 'leftCircularArrow':
                pptSlide.addShape(pptx.shapes.LEFT_CIRCULAR_ARROW, shapeOptions);
                break;
            case 'leftRightArrow':
                pptSlide.addShape(pptx.shapes.LEFT_RIGHT_ARROW, shapeOptions);
                break;
            case 'leftRightArrowCallout':
                pptSlide.addShape(pptx.shapes.LEFT_RIGHT_ARROW_CALLOUT, shapeOptions);
                break;
            case 'leftRightCircularArrow':
                pptSlide.addShape(pptx.shapes.LEFT_RIGHT_CIRCULAR_ARROW, shapeOptions);
                break;
            case 'leftRightRibbon':
                pptSlide.addShape(pptx.shapes.LEFT_RIGHT_RIBBON, shapeOptions);
                break;
            case 'leftRightUpArrow':
                pptSlide.addShape(pptx.shapes.LEFT_RIGHT_UP_ARROW, shapeOptions);
                break;
            case 'leftUpArrow':
                pptSlide.addShape(pptx.shapes.LEFT_UP_ARROW, shapeOptions);
                break;
            case 'lightningBolt':
                pptSlide.addShape(pptx.shapes.LIGHTNING_BOLT, shapeOptions);
                break;

            // Line Callout Shapes
            case 'borderCallout1':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_1, shapeOptions);
                break;
            case 'accentCallout1':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_1_ACCENT_BAR, shapeOptions);
                break;
            case 'accentBorderCallout1':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR, shapeOptions);
                break;
            case 'callout1':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_1_NO_BORDER, shapeOptions);
                break;
            case 'borderCallout2':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_2, shapeOptions);
                break;
            case 'accentCallout2':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_2_ACCENT_BAR, shapeOptions);
                break;
            case 'accentBorderCallout2':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR, shapeOptions);
                break;
            case 'callout2':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_2_NO_BORDER, shapeOptions);
                break;
            case 'borderCallout3':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_3, shapeOptions);
                break;
            case 'accentCallout3':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_3_ACCENT_BAR, shapeOptions);
                break;
            case 'accentBorderCallout3':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR, shapeOptions);
                break;
            case 'callout3':
                pptSlide.addShape(pptx.shapes.LINE_CALLOUT_3_NO_BORDER, shapeOptions);
                break;

            // Line Shapes
            case 'line':
                // Check if this is actually a connector with stored data
                const lineConnectorInfo = shapeElement.getAttribute('data-connector-info');

                if (lineConnectorInfo) {
                    // ✅ This is a connector (straight arrow connector)
                    try {
                        const parsedData = JSON.parse(lineConnectorInfo);

                        const connectorData = {
                            shapeType: parsedData.shapeType || 'straightConnector1',
                            shapeName: shapeElement.getAttribute('data-name') || objName,
                            position: {
                                x: shapeOptions.x * 72,
                                y: shapeOptions.y * 72,
                                width: shapeOptions.w * 72,
                                height: shapeOptions.h * 72
                            },
                            strokeColor: parsedData.strokeColor,
                            strokeWidth: parsedData.strokeWidth,
                            dashType: parsedData.dashType || 'solid',
                            rotation: shapeOptions.rotate || 0,
                            flipH: parsedData.flipH || false,
                            flipV: parsedData.flipV || false,
                            lineEnds: parsedData.lineEnds || { headType: 'none', tailType: 'none' },
                            segments: parsedData.segments || []
                        };

                        convertConnectorToPPTX(pptx, pptSlide, connectorData);

                    } catch (error) {
                        console.error('Failed to parse connector data for line:', error);
                        // Fall through to regular line handling
                    }
                } else {
                    // ❌ Regular line (no connector data)
                    let lineProperties = getLineColorFromSvg(shapeElement);
                    let lineColor = lineProperties.color;
                    let lineWidth = lineProperties.width;

                    // Fallback to style.color if SVG color not found
                    if (!lineColor) {
                        lineColor = style.color || style.stroke || style.backgroundColor || '#000000';
                    }

                    // Fallback to CSS stroke-width or border-width if SVG width not found
                    if (!lineWidth) {
                        const cssStrokeWidth = style.strokeWidth || style.borderWidth;
                        if (cssStrokeWidth) {
                            lineWidth = parseFloat(cssStrokeWidth.replace('px', ''));
                        } else {
                            lineWidth = 1; // Default width
                        }
                    }

                    pptSlide.addShape(pptx.shapes.LINE, {
                        ...shapeOptions,
                        line: {
                            color: rgbToHex(lineColor),
                            width: lineWidth
                        }
                    });
                }
                break;
            // CORRECTED VERSION - Replace the existing 'line-lineheight' case (lines 730-846) with this:

            case 'line-lineheight':

                let lineProperties1 = getLineColorFromSvg(shapeElement);
                let lineColor2 = lineProperties1.color;
                let lineWidth2 = lineProperties1.width;

                // Fallback to style.color if SVG color not found
                if (!lineColor2) {
                    lineColor2 = style.color || style.stroke || style.backgroundColor || '#000000';
                }

                // Read stroke width from visual height (or stroke/border width)
                if (!lineWidth2) {
                    const lineStrokeWidth = style.height || style.strokeWidth || style.borderWidth;
                    if (lineStrokeWidth) {
                        lineWidth2 = parseFloat(lineStrokeWidth.toString().replace('px', ''));
                    } else {
                        lineWidth2 = 1; // Default width
                    }
                }

                // Extract line properties from HTML
                const lineStartX = parseFloat(style.left || "0");
                const lineStartY = parseFloat(style.top || "0");
                const lineLength = parseFloat(style.width || "0"); // visual length

                // Extract rotation angle
                const lineTransform = style.transform || '';
                const lineRotationMatch = lineTransform.match(/rotate\(([-\d.]+)deg\)/);
                const lineRotationDeg = lineRotationMatch ? parseFloat(lineRotationMatch[1]) : 0;

                // Compute end point from length + angle
                const rotationRad = (lineRotationDeg * Math.PI) / 180;
                const deltaX = lineLength * Math.cos(rotationRad);
                const deltaY = lineLength * Math.sin(rotationRad);

                // px → inches for PPTX (72 DPI)
                const lineX = lineStartX / 72;
                const lineY = lineStartY / 72;
                const lineW = Math.abs(deltaX) / 72;
                const lineH = Math.abs(deltaY) / 72;

                // Flip flags
                const lineFlipH = deltaX < 0;
                const lineFlipV = deltaY < 0;

                // Adjust x/y if flipped so the visual line starts at the correct end
                let adjustedX = lineX;
                let adjustedY = lineY;

                if (lineFlipH) {
                    adjustedX = lineX + deltaX / 72;
                }
                if (lineFlipV) {
                    adjustedY = lineY + deltaY / 72;
                }

                const lineOptions = {
                    x: adjustedX,
                    y: adjustedY,
                    w: lineW,
                    h: lineH,
                    flipH: lineFlipH,
                    flipV: lineFlipV,
                    line: {
                        color: rgbToHex(lineColor2),
                        width: lineWidth2
                    },
                    objectName: objName || ''
                };

                pptSlide.addShape(pptx.shapes.LINE, lineOptions);
                break;
            case 'lineInv':
                pptSlide.addShape(pptx.shapes.LINE_INVERSE, shapeOptions);
                break;

            // Math Shapes
            case 'mathDivide':
                pptSlide.addShape(pptx.shapes.MATH_DIVIDE, shapeOptions);
                break;
            case 'mathEqual':
                pptSlide.addShape(pptx.shapes.MATH_EQUAL, shapeOptions);
                break;
            case 'mathMinus':
                pptSlide.addShape(pptx.shapes.MATH_MINUS, shapeOptions);
                break;
            case 'mathMultiply':
                pptSlide.addShape(pptx.shapes.MATH_MULTIPLY, shapeOptions);
                break;
            case 'mathNotEqual':
                pptSlide.addShape(pptx.shapes.MATH_NOT_EQUAL, shapeOptions);
                break;
            case 'mathPlus':
                pptSlide.addShape(pptx.shapes.MATH_PLUS, shapeOptions);
                break;

            // Various Other Shapes
            case 'moon':
                pptSlide.addShape(pptx.shapes.MOON, shapeOptions);
                break;
            case 'nonIsoscelesTrapezoid':
                pptSlide.addShape(pptx.shapes.NON_ISOSCELES_TRAPEZOID, shapeOptions);
                break;
            case 'notchedRightArrow':
                pptSlide.addShape(pptx.shapes.NOTCHED_RIGHT_ARROW, shapeOptions);
                break;
            case 'noSmoking':
                pptSlide.addShape(pptx.shapes.NO_SYMBOL, shapeOptions);
                break;
            case 'octagon':
                pptSlide.addShape(pptx.shapes.OCTAGON, shapeOptions);
                break;
            case 'ellipse':
                pptSlide.addShape(pptx.shapes.OVAL, shapeOptions);
                break;
            case 'wedgeEllipseCallout':
                pptSlide.addShape(pptx.shapes.OVAL_CALLOUT, shapeOptions);
                break;
            case 'parallelogram':
                pptSlide.addShape(pptx.shapes.PARALLELOGRAM, shapeOptions);
                break;
            case 'homePlate':
                pptSlide.addShape(pptx.shapes.PENTAGON, shapeOptions);
                break;
            case 'pie': {
                let pieAdj1 = shapeElement.getAttribute('data-adj1');
                let pieAdj2 = shapeElement.getAttribute('data-adj2');

                pieAdj1 = pieAdj1 !== null && pieAdj1 !== '' ? parseInt(pieAdj1, 10) : null;
                pieAdj2 = pieAdj2 !== null && pieAdj2 !== '' ? parseInt(pieAdj2, 10) : null;

                // If adj values are missing, fallback to preset PIE rather than broken custGeom
                if (
                    pieAdj1 === null || Number.isNaN(pieAdj1) ||
                    pieAdj2 === null || Number.isNaN(pieAdj2)
                ) {
                    pptSlide.addShape(pptx.shapes.PIE, {
                        ...shapeOptions,
                        hidden: false
                    });
                    break;
                }

                const pieX = x;
                const pieY = y;
                const pieW = w;
                const pieH = h;

                // local shape dimensions in px (not inches)
                const localW = parseFloat(rawWidth);
                const localH = parseFloat(rawHeight);

                const cx = localW / 2;
                const cy = localH / 2;
                const rx = localW / 2;
                const ry = localH / 2;

                const normDeg = (deg) => {
                    let d = deg % 360;
                    if (d < 0) d += 360;
                    return d;
                };

                // OOXML angle units -> degrees
                const startDeg = pieAdj1 / 60000;
                const endDeg = pieAdj2 / 60000;

                let sweepDeg = normDeg(endDeg - startDeg);
                if (sweepDeg === 0) sweepDeg = 360;

                // sample arc into many line points
                const pointCount = Math.max(24, Math.ceil(sweepDeg / 6));

                const toLocalPoint = (deg) => {
                    const rad = (normDeg(deg) * Math.PI) / 180;
                    return {
                        x: cx + rx * Math.cos(rad),
                        y: cy + ry * Math.sin(rad)
                    };
                };

                const ptsPx = [];
                ptsPx.push({ x: cx, y: cy }); // center

                for (let i = 0; i <= pointCount; i++) {
                    const t = i / pointCount;
                    const deg = startDeg + sweepDeg * t;
                    ptsPx.push(toLocalPoint(deg));
                }

                // convert to custGeom points in ppt local units
                const custGeomPoints = ptsPx.map((pt) => ({
                    x: (pt.x / localW) * pieW,
                    y: (pt.y / localH) * pieH
                }));

                custGeomPoints.push({ close: true });

                // fill
                let pieFill = shapeOptions.fill || { color: '808080' };

                // line
                let pieLine = shapeOptions.line || { color: '000000', transparency: 100, width: 0 };

                // If SVG exists, prefer exact fill/stroke from it
                const svgEl = shapeElement.querySelector('svg');
                const pathEl = svgEl?.querySelector('path');

                if (pathEl) {
                    const svgFill = pathEl.getAttribute('fill');
                    if (svgFill && svgFill !== 'none' && svgFill !== 'transparent') {
                        pieFill = {
                            color: rgbToHex(svgFill).replace('#', '').toUpperCase(),
                            transparency: calculatedTransparency || 0
                        };
                    }

                    const svgStroke = pathEl.getAttribute('stroke');
                    const svgStrokeWidth = parseFloat(pathEl.getAttribute('stroke-width') || '0');

                    if (
                        svgStroke &&
                        svgStroke !== 'none' &&
                        svgStroke !== 'transparent' &&
                        !Number.isNaN(svgStrokeWidth) &&
                        svgStrokeWidth > 0
                    ) {
                        pieLine = {
                            color: rgbToHex(svgStroke).replace('#', '').toUpperCase(),
                            width: svgStrokeWidth
                        };
                    } else {
                        pieLine = { color: '000000', transparency: 100, width: 0 };
                    }

                    const fillOpacityRaw = pathEl.getAttribute('fill-opacity') || pathEl.getAttribute('opacity');
                    if (fillOpacityRaw !== null && fillOpacityRaw !== '') {
                        const fillOpacity = parseFloat(fillOpacityRaw);
                        if (!Number.isNaN(fillOpacity) && fillOpacity >= 0 && fillOpacity < 1) {
                            pieFill.transparency = Math.round((1 - fillOpacity) * 100);
                        }
                    }
                }

                // preserve original theme fill if available
                if (
                    originalThemeColor &&
                    originalThemeColor !== 'null' &&
                    originalThemeColor !== 'undefined'
                ) {
                    const themeColorObj = createThemeColorObject(
                        originalThemeColor,
                        originalLumMod,
                        originalLumOff,
                        originalAlpha
                    );
                    if (themeColorObj) {
                        pieFill = { color: themeColorObj };
                    }
                }

                const pieShapeOptions = {
                    x: pieX,
                    y: pieY,
                    w: pieW,
                    h: pieH,
                    rotate: rotation,
                    objectName: objName || shapeElement.getAttribute('data-name') || 'pie',
                    hidden: false,
                    fill: pieFill,
                    line: pieLine,
                    points: custGeomPoints,
                    ...(shapeOptions.shadow ? { shadow: shapeOptions.shadow } : {}),
                    ...(shapeOptions.glow ? { glow: shapeOptions.glow } : {})
                };

                if (shapeOptions.flipH) pieShapeOptions.flipH = true;
                if (shapeOptions.flipV) pieShapeOptions.flipV = true;

                pptSlide.addShape('custGeom', pieShapeOptions);
                break;
            }
            case 'pieWedge':
                pptSlide.addShape(pptx.shapes.PIE_WEDGE, shapeOptions);
                break;
            case 'plaque':
                pptSlide.addShape(pptx.shapes.PLAQUE, shapeOptions);
                break;
            case 'plaqueTabs':
                pptSlide.addShape(pptx.shapes.PLAQUE_TABS, shapeOptions);
                break;
            case 'quadArrow':
                pptSlide.addShape(pptx.shapes.QUAD_ARROW, shapeOptions);
                break;
            case 'quadArrowCallout':
                pptSlide.addShape(pptx.shapes.QUAD_ARROW_CALLOUT, shapeOptions);
                break;
            case 'rect':
                pptSlide.addShape(pptx.shapes.RECTANGLE, shapeOptions);
                break;
            case 'wedgeRectCallout':
                pptSlide.addShape(pptx.shapes.RECTANGULAR_CALLOUT, shapeOptions);
                break;
            case 'pentagon':
                pptSlide.addShape(pptx.shapes.REGULAR_PENTAGON, shapeOptions);
                break;
            case 'rightArrow': {
                const clipPathValue =
                    shapeElement.style.clipPath ||
                    shapeElement.style.webkitClipPath ||
                    shapeElement.getAttribute('data-clip-path') ||
                    '';

                const polygonMatch = clipPathValue.match(/polygon\((.*)\)/i);

                // Fallback to preset arrow if polygon is missing
                if (!polygonMatch || !polygonMatch[1]) {
                    pptSlide.addShape(pptx.shapes.RIGHT_ARROW, shapeOptions);
                    break;
                }

                // Example input:
                // polygon(0% 17.9925%, 92.6457% 17.9925%, 92.6457% 0%, 100% 50%, 92.6457% 100%, 92.6457% 82.0075%, 0% 82.0075%)
                const rawPoints = polygonMatch[1]
                    .split(',')
                    .map(p => p.trim())
                    .filter(Boolean);

                const parsedPoints = [];

                for (const pt of rawPoints) {
                    const m = pt.match(/(-?\d*\.?\d+)%\s+(-?\d*\.?\d+)%/);
                    if (!m) continue;

                    const xPct = parseFloat(m[1]);
                    const yPct = parseFloat(m[2]);

                    if (Number.isNaN(xPct) || Number.isNaN(yPct)) continue;

                    parsedPoints.push({ x: xPct, y: yPct });
                }

                // Need at least 3 points for a valid custom geometry
                if (parsedPoints.length < 3) {
                    pptSlide.addShape(pptx.shapes.RIGHT_ARROW, shapeOptions);
                    break;
                }

                // Build SVG path in a normalized 100x100 viewBox
                let pathD = `M ${parsedPoints[0].x} ${parsedPoints[0].y}`;
                for (let i = 1; i < parsedPoints.length; i++) {
                    pathD += ` L ${parsedPoints[i].x} ${parsedPoints[i].y}`;
                }
                pathD += ' Z';

                // Convert normalized polygon -> pptx custGeom points
                const custGeomPoints = svgAddToSlide.convertSvgPathToPptxPoints(
                    pathD,
                    100,              // viewBox width
                    100,              // viewBox height
                    shapeOptions.w,   // actual width in inches
                    shapeOptions.h    // actual height in inches
                );

                if (!custGeomPoints || custGeomPoints.filter(p => !p.close).length < 3) {
                    pptSlide.addShape(pptx.shapes.RIGHT_ARROW, shapeOptions);
                    break;
                }

                pptSlide.addShape('custGeom', {
                    ...shapeOptions,
                    objectName: objName || 'rightArrow',
                    points: custGeomPoints
                });

                break;
            }
            case 'rightArrowCallout':
                pptSlide.addShape(pptx.shapes.RIGHT_ARROW_CALLOUT, shapeOptions);
                break;
            case 'rightBrace':
                pptSlide.addShape(pptx.shapes.RIGHT_BRACE, shapeOptions);
                break;
            case 'rightBracket':
                pptSlide.addShape(pptx.shapes.RIGHT_BRACKET, shapeOptions);
                break;
            case 'rtTriangle':
                pptSlide.addShape(pptx.shapes.RIGHT_TRIANGLE, shapeOptions);
                break;
            case 'roundRect':
                pptSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, shapeOptions);
                break;
            case 'wedgeRoundRectCallout':
                pptSlide.addShape(pptx.shapes.ROUNDED_RECTANGULAR_CALLOUT, shapeOptions);
                break;
            case 'round1Rect':
                pptSlide.addShape(pptx.shapes.ROUND_1_RECTANGLE, shapeOptions);
                break;
            case 'round2DiagRect':
                pptSlide.addShape(pptx.shapes.ROUND_2_DIAG_RECTANGLE, shapeOptions);
                break;
            case 'round2SameRect': {
                const svgEl = shapeElement.querySelector('svg');
                const styleAttr = shapeElement.getAttribute('style') || '';

                let rTopPx = 0;     // top-left + top-right
                let rBottomPx = 0;  // bottom-right + bottom-left

                // 1. Try clip-path first
                const clipMatch = styleAttr.match(
                    /clip-path\s*:\s*inset\([^)]*round\s+([\d.]+)px\s+([\d.]+)px\s+([\d.]+)px\s+([\d.]+)px/i
                );
                if (clipMatch) {
                    rTopPx = parseFloat(clipMatch[1] || '0');
                    rBottomPx = parseFloat(clipMatch[3] || '0');
                } else {
                    // 2. Fallback to border-radius
                    const brMatch = styleAttr.match(
                        /border-radius\s*:\s*([\d.]+)px(?:\s+([\d.]+)px)?(?:\s+([\d.]+)px)?(?:\s+([\d.]+)px)?/i
                    );
                    if (brMatch) {
                        rTopPx = parseFloat(brMatch[1] || '0');
                        rBottomPx = parseFloat(brMatch[3] || brMatch[1] || '0');
                    }
                }

                const shapeWpx = Math.max(widthPx || 0, 1);
                const shapeHpx = Math.max(heightPx || 0, 1);
                const minDimPx = Math.min(shapeWpx, shapeHpx);

                rTopPx = Math.max(0, Math.min(rTopPx, minDimPx / 2));
                rBottomPx = Math.max(0, Math.min(rBottomPx, minDimPx / 2));

                // Convert px radii into the local SVG/viewBox space
                let vbW = shapeWpx;
                let vbH = shapeHpx;
                if (svgEl) {
                    const vb = (svgEl.getAttribute('viewBox') || '').trim().split(/\s+/).map(Number);
                    if (vb.length === 4 && Number.isFinite(vb[2]) && Number.isFinite(vb[3]) && vb[2] > 0 && vb[3] > 0) {
                        vbW = vb[2];
                        vbH = vb[3];
                    }
                }

                const sx = vbW / shapeWpx;
                const sy = vbH / shapeHpx;
                const rt = rTopPx * Math.min(sx, sy);
                const rb = rBottomPx * Math.min(sx, sy);

                // Cubic bezier approximation for quarter circle
                const K = 0.5522847498307936;

                const pathD = [
                    `M ${rt} 0`,
                    `L ${vbW - rt} 0`,
                    `C ${vbW - rt + rt * K} 0 ${vbW} ${rt - rt * K} ${vbW} ${rt}`,
                    `L ${vbW} ${vbH - rb}`,
                    `C ${vbW} ${vbH - rb + rb * K} ${vbW - rb + rb * K} ${vbH} ${vbW - rb} ${vbH}`,
                    `L ${rb} ${vbH}`,
                    `C ${rb - rb * K} ${vbH} 0 ${vbH - rb + rb * K} 0 ${vbH - rb}`,
                    `L 0 ${rt}`,
                    `C 0 ${rt - rt * K} ${rt - rt * K} 0 ${rt} 0`,
                    'Z'
                ].join(' ');

                const custGeomPoints = svgAddToSlide.convertSvgPathToPptxPoints(
                    pathD,
                    vbW,
                    vbH,
                    shapeOptions.w,
                    shapeOptions.h
                );

                if (!custGeomPoints || custGeomPoints.filter(p => !p.close).length < 3) {
                    // safe fallback
                    pptSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                        ...shapeOptions,
                        objectName: objName || 'round2SameRect'
                    });
                    break;
                }

                const customOpts = {
                    ...shapeOptions,
                    objectName: objName || 'round2SameRect',
                    points: custGeomPoints
                };

                // rectRadius must NOT be present on custGeom
                delete customOpts.rectRadius;

                pptSlide.addShape('custGeom', customOpts);
                break;
            }
            case 'smileyFace':
                pptSlide.addShape(pptx.shapes.SMILEY_FACE, shapeOptions);
                break;
            case 'snip1Rect':
                pptSlide.addShape(pptx.shapes.SNIP_1_RECTANGLE, shapeOptions);
                break;
            case 'snip2DiagRect':
                pptSlide.addShape(pptx.shapes.SNIP_2_DIAG_RECTANGLE, shapeOptions);
                break;
            case 'snip2SameRect':
                pptSlide.addShape(pptx.shapes.SNIP_2_SAME_RECTANGLE, shapeOptions);
                break;
            case 'snipRoundRect':
                pptSlide.addShape(pptx.shapes.SNIP_ROUND_RECTANGLE, shapeOptions);
                break;
            case 'squareTabs':
                pptSlide.addShape(pptx.shapes.SQUARE_TABS, shapeOptions);
                break;

            // Star Shapes
            case 'star10':
                pptSlide.addShape(pptx.shapes.STAR_10_POINT, shapeOptions);
                break;
            case 'star12':
                pptSlide.addShape(pptx.shapes.STAR_12_POINT, shapeOptions);
                break;
            case 'star16':
                pptSlide.addShape(pptx.shapes.STAR_16_POINT, shapeOptions);
                break;
            case 'star24':
                pptSlide.addShape(pptx.shapes.STAR_24_POINT, shapeOptions);
                break;
            case 'star32':
                pptSlide.addShape(pptx.shapes.STAR_32_POINT, shapeOptions);
                break;
            case 'star4':
                pptSlide.addShape(pptx.shapes.STAR_4_POINT, shapeOptions);
                break;
            case 'star5':
                pptSlide.addShape(pptx.shapes.STAR_5_POINT, shapeOptions);
                break;
            case 'star6':
                pptSlide.addShape(pptx.shapes.STAR_6_POINT, shapeOptions);
                break;
            case 'star7':
                pptSlide.addShape(pptx.shapes.STAR_7_POINT, shapeOptions);
                break;
            case 'star8':
                pptSlide.addShape(pptx.shapes.STAR_8_POINT, shapeOptions);
                break;
            case 'stripedRightArrow':
                pptSlide.addShape(pptx.shapes.STRIPED_RIGHT_ARROW, shapeOptions);
                break;
            case 'sun':
                pptSlide.addShape(pptx.shapes.SUN, shapeOptions);
                break;
            case 'swooshArrow':
                pptSlide.addShape(pptx.shapes.SWOOSH_ARROW, shapeOptions);
                break;
            case 'teardrop': {

                const tdRawStyle = shapeElement.getAttribute('style') || '';

                // ── Position (px → inches at 72 DPI) ──────────────────────────────
                const tdGet = (prop) => {
                    const m = tdRawStyle.match(new RegExp(prop + '\\s*:\\s*([0-9.-]+)px', 'i'));
                    return m ? parseFloat(m[1]) : 0;
                };
                const tdX = tdGet('left') / 72;
                const tdY = tdGet('top') / 72;
                const tdW = tdGet('width') / 72;
                const tdH = tdGet('height') / 72;

                // ── Rotation (degrees) ─────────────────────────────────────────────
                const tdRotM = tdRawStyle.match(/rotate\(([-\d.]+)deg\)/i);
                const tdRotation = tdRotM ? parseFloat(tdRotM[1]) : 0;

                // ── Flip (explicit data attributes written by shapeHandler.js) ─────
                const tdFlipH = shapeElement.getAttribute('data-flip-h') === '1';
                const tdFlipV = shapeElement.getAttribute('data-flip-v') === '1';

                // ── Fill color from SVG <path fill="..."> ─────────────────────────
                let tdFillColor = 'FFFFFF';
                const tdSvgPath = shapeElement.querySelector('svg path');
                if (tdSvgPath) {
                    const tdFillRaw = tdSvgPath.getAttribute('fill') || '';
                    if (tdFillRaw && tdFillRaw !== 'none' && tdFillRaw !== 'transparent') {
                        tdFillColor = rgbToHex(tdFillRaw).replace('#', '').toUpperCase();
                    }
                }

                // ── Stroke from SVG <path stroke="..."> ───────────────────────────
                let tdLineOptions = undefined;  // undefined = no line (noFill in PPTX)
                if (tdSvgPath) {
                    const tdStrokeRaw = tdSvgPath.getAttribute('stroke') || 'none';
                    const tdStrokeWidth = parseFloat(tdSvgPath.getAttribute('stroke-width') || '0');
                    if (tdStrokeRaw !== 'none' && tdStrokeRaw !== 'transparent' && tdStrokeWidth > 0) {
                        tdLineOptions = {
                            color: rgbToHex(tdStrokeRaw).replace('#', '').toUpperCase(),
                            width: tdStrokeWidth
                        };
                    }
                }

                // ── Opacity ────────────────────────────────────────────────────────
                let tdTransparency = undefined;
                const tdOpacityM = tdRawStyle.match(/(?:^|;)\s*opacity\s*:\s*([\d.]+)/i);
                if (tdOpacityM) {
                    const tdOp = parseFloat(tdOpacityM[1]);
                    if (!isNaN(tdOp) && tdOp < 1) {
                        tdTransparency = Math.round((1 - tdOp) * 100);
                    }
                }

                // ── adj (tip sharpness) ────────────────────────────────────────────
                const tdAdj = parseInt(shapeElement.getAttribute('data-adj') || '100000', 10);

                // ── Build final shape options ──────────────────────────────────────
                const tdShapeOptions = {
                    x: tdX,
                    y: tdY,
                    w: tdW,
                    h: tdH,
                    rotate: tdRotation,
                    fill: { color: tdFillColor },
                    objectName: objName || '',
                };

                if (tdFlipH) tdShapeOptions.flipH = true;
                if (tdFlipV) tdShapeOptions.flipV = true;
                if (tdLineOptions) tdShapeOptions.line = tdLineOptions;
                if (tdTransparency !== undefined) tdShapeOptions.transparency = tdTransparency;
                if (!isNaN(tdAdj) && tdAdj !== 100000) tdShapeOptions.adjVal1 = tdAdj;

                // Carry over shadow / glow if present
                if (shapeOptions.shadow) tdShapeOptions.shadow = shapeOptions.shadow;
                if (shapeOptions.glow) tdShapeOptions.glow = shapeOptions.glow;

                pptSlide.addShape(pptx.shapes.TEAR, tdShapeOptions);
                break;
            }
            case 'trapezoid':
                pptSlide.addShape(pptx.shapes.TRAPEZOID, shapeOptions);
                break;
            case 'upArrow':
                pptSlide.addShape(pptx.shapes.UP_ARROW, shapeOptions);
                break;
            case 'upArrowCallout':
                pptSlide.addShape(pptx.shapes.UP_ARROW_CALLOUT, shapeOptions);
                break;
            case 'upDownArrow':
                pptSlide.addShape(pptx.shapes.UP_DOWN_ARROW, shapeOptions);
                break;
            case 'upDownArrowCallout':
                pptSlide.addShape(pptx.shapes.UP_DOWN_ARROW_CALLOUT, shapeOptions);
                break;
            case 'ribbon2':
                pptSlide.addShape(pptx.shapes.UP_RIBBON, shapeOptions);
                break;
            case 'uturnArrow':
                pptSlide.addShape(pptx.shapes.U_TURN_ARROW, shapeOptions);
                break;
            case 'verticalScroll':
                pptSlide.addShape(pptx.shapes.VERTICAL_SCROLL, shapeOptions);
                break;
            case 'wave':
                pptSlide.addShape(pptx.shapes.WAVE, shapeOptions);
                break;
            // If none of these shape types match, use a default rectangle
            // default:
            //     pptSlide.addShape(pptx.shapes.RECTANGLE, shapeOptions);
            //     break;
        }

        // Store shape data on slide for XML generation
        if (!pptSlide.shapes) pptSlide.shapes = [];

        pptSlide.shapes.push(shapeData);
        // After pptSlide.shapes.push(shapeData)
        if (softEdgeRad > 0) {
            if (!global.softEdgeStore) global.softEdgeStore = new Map();
            const storeKey = objName || `shape_${x.toFixed(2)}_${y.toFixed(2)}`;
            global.softEdgeStore.set(storeKey, {
                radEMU: softEdgeRad,
                shapeName: objName,
                x, y, w, h
            });
            console.log(`   🌫️ Queued soft edge: shape="${objName}", rad=${softEdgeRad} EMU`);
        }
        // // Log successful shape addition with flip information
        // if (flipHorizontal || flipVertical) {
        // }

    } catch (error) {
        console.error(`Failed to add shape '${shapeId}' with error: ${error}`);
    }
}


// Helper function to properly detect if there's a valid border
function checkForValidBorder(style) {
    // Check if border is explicitly set to 'none'
    if (!style.border || style.border === 'none') {
        return false;
    }

    // Parse border width and color from style.border property
    let borderWidth = 0;
    let borderColor = '';

    // Try to extract width and color from style.border (e.g., "1.25px solid transparent" or "1.25px solid #000")
    if (style.border) {
        const borderMatch = style.border.match(/^(\d*\.?\d+)px\s+(\w+)\s+(.+)$/);
        if (borderMatch) {
            borderWidth = parseFloat(borderMatch[1]);
            borderColor = borderMatch[3].trim();
        }
    }

    // Fallback to individual style properties if border shorthand parsing fails
    if (borderWidth === 0 && style.borderWidth) {
        borderWidth = parseFloat(style.borderWidth);
    }

    if (!borderColor && style.borderColor) {
        borderColor = style.borderColor;
    }

    // Only consider it a valid border if:
    // 1. Width > 0
    // 2. Color is not transparent
    return borderWidth > 0 && borderColor !== 'transparent';
}


// NEW FUNCTION: Handle custom SVG shapes
function handleCustomSvgShape(pptx, pptSlide, shapeElement, svgElement, slideContext) {
    try {
        const style = shapeElement.style;

        // Extract position and dimensions
        const x = parseFloat(style.left || "0") / 72;
        const y = parseFloat(style.top || "0") / 72;
        const w = parseFloat(style.width || "0") / 72;
        const h = parseFloat(style.height || "0") / 72;

        // Extract rotation and flipping
        const transform = style.transform || '';
        const rotationMatch = transform.match(/rotate\(([-\d.]+)deg\)/);
        const rotation = rotationMatch ? parseFloat(rotationMatch[1]) : 0;

        // Check for scale transforms (flipping)
        const scaleXMatch = transform.match(/scaleX\((-?\d*\.?\d+)\)/);
        const scaleYMatch = transform.match(/scaleY\((-?\d*\.?\d+)\)/);

        let flipHorizontal = false;
        let flipVertical = false;

        if (scaleXMatch && parseFloat(scaleXMatch[1]) < 0) {
            flipHorizontal = true;
        }
        if (scaleYMatch && parseFloat(scaleYMatch[1]) < 0) {
            flipVertical = true;
        }

        // Extract color from SVG path
        const path = svgElement.querySelector('path');
        let fillColor = '#000000'; // default

        if (path) {
            const pathFill = path.getAttribute('fill');
            if (pathFill && pathFill !== 'none') {
                fillColor = pathFill;
            }
        }

        // Extract opacity
        const opacity = parseFloat(style.opacity || '1');
        const transparency = Math.round((1 - opacity) * 100);

        // Determine shape type based on SVG path and dimensions
        let shapeType = pptx.shapes.RECTANGLE; // default

        // For very thin rectangles (height < 10px), treat as line
        const heightPx = parseFloat(style.height || "0");
        if (heightPx <= 10) {
            shapeType = pptx.shapes.RECTANGLE; // Use thin rectangle instead of line for better control
        }

        // Create shape options
        const shapeOptions = {
            x: x,
            y: y,
            w: w,
            h: h,
            rotate: rotation,
            // fill: {
            //     color: rgbToHex(fillColor),
            //     transparency: transparency
            // }
        };

        // Add flipping if needed
        if (flipHorizontal) {
            shapeOptions.flipH = true;
        }
        if (flipVertical) {
            shapeOptions.flipV = true;
        }

        // Add border if stroke is defined
        if (path) {
            const stroke = path.getAttribute('stroke');
            const strokeWidth = path.getAttribute('stroke-width');

            if (stroke && stroke !== 'none' && stroke !== '') {
                shapeOptions.line = {
                    color: rgbToHex(stroke),
                    width: parseFloat(strokeWidth || '1')
                };
            }
        }

        // Add the shape to slide
        pptSlide.addShape(shapeType, shapeOptions);

        // Store shape data for XML generation
        const shapeData = {
            id: 'customSvg',
            x: x,
            y: y,
            w: w,
            h: h,
            rotation: rotation,
            flipH: flipHorizontal,
            flipV: flipVertical,
            fill: shapeOptions.fill,
            line: shapeOptions.line,
            isSpecialShape: false,
            shapeType: 'rect'
        };

        if (!pptSlide.shapes) pptSlide.shapes = [];
        pptSlide.shapes.push(shapeData);

        return true;

    } catch (error) {
        console.error('   ❌ Error processing custom SVG shape:', error);
        return false;
    }
}

// Rest of the existing functions remain unchanged...
function getLineColorFromSvg(shapeElement) {
    try {
        const svgElement = shapeElement.querySelector('svg');
        if (!svgElement) {
            return { color: null, width: null };
        }

        let strokeColor = null;
        let strokeWidth = null;

        // First check line elements with stroke attributes
        const lineElement = svgElement.querySelector('line[stroke]');
        if (lineElement) {
            const color = lineElement.getAttribute('stroke');
            if (color && color !== 'none' && color !== 'transparent') {
                strokeColor = color;
            }

            const width = lineElement.getAttribute('stroke-width');
            if (width) {
                strokeWidth = parseFloat(width.replace('px', ''));
            }
        }

        // If not found, check all SVG elements with stroke attributes
        if (!strokeColor || !strokeWidth) {
            const svgElementsWithStroke = svgElement.querySelectorAll('[stroke]');
            for (const element of svgElementsWithStroke) {
                if (!strokeColor) {
                    const color = element.getAttribute('stroke');
                    if (color && color !== 'none' && color !== 'transparent') {
                        strokeColor = color;
                    }
                }

                if (!strokeWidth) {
                    const width = element.getAttribute('stroke-width');
                    if (width) {
                        strokeWidth = parseFloat(width.replace('px', ''));
                    }
                }

                // Break if both are found
                if (strokeColor && strokeWidth) break;
            }
        }

        // Finally check SVG style attribute
        const svgStyle = svgElement.getAttribute('style');
        if (svgStyle) {
            if (!strokeColor) {
                const strokeMatch = svgStyle.match(/stroke:\s*([^;]+)/i);
                if (strokeMatch && strokeMatch[1]) {
                    strokeColor = strokeMatch[1].trim();
                }
            }

            if (!strokeWidth) {
                const widthMatch = svgStyle.match(/stroke-width:\s*([^;]+)/i);
                if (widthMatch && widthMatch[1]) {
                    strokeWidth = parseFloat(widthMatch[1].replace('px', ''));
                }
            }
        }

        return {
            color: strokeColor,
            width: strokeWidth
        };
    } catch (error) {
        console.error('Error extracting line properties from SVG:', error);
        return { color: null, width: null };
    }
}

function getBackgroundColor(shapeElement) {
    const styleAttr = shapeElement.getAttribute("style");

    if (styleAttr) {
        const patterns = [
            /background:\s*([^;\n\r]+)/i,
            /background-color:\s*([^;\n\r]+)/i,
            /background\s*:\s*([^;\n\r]+)/i
        ];

        for (const pattern of patterns) {
            const match = styleAttr.match(pattern);
            if (match && match[1]) {
                const extracted = match[1].trim();
                if (extracted !== 'transparent' && extracted !== 'none') {
                    return extracted;
                }
            }
        }
    }

    const computedBg = shapeElement.style.background;
    if (computedBg && computedBg !== 'transparent' && computedBg !== 'none') {
        return computedBg;
    }

    const computedBgColor = shapeElement.style.backgroundColor;
    if (computedBgColor && computedBgColor !== 'transparent' && computedBgColor !== 'none') {
        return computedBgColor;
    }

    if (typeof getComputedStyle !== 'undefined') {
        try {
            const computed = getComputedStyle(shapeElement);
            const computedBackground = computed.backgroundColor;
            if (computedBackground && computedBackground !== 'transparent' && computedBackground !== 'rgba(0, 0, 0, 0)') {
                return computedBackground;
            }
        } catch (e) {
            console.log("getComputedStyle not available or failed:", e.message);
        }
    }

    return 'transparent';
}

function getSlideDimensions(pptSlide) {
    if (addTextBox.getSlideDimensions) {

        return addTextBox.getSlideDimensions(pptSlide);
    }

    try {
        let width = 10;
        let height = 7.5;

        if (pptSlide && pptSlide.slideLayout) {
            const layout = pptSlide.slideLayout;
            if (layout.width) width = layout.width;
            if (layout.height) height = layout.height;
        }

        if (pptSlide && pptSlide.width) width = pptSlide.width;
        if (pptSlide && pptSlide.height) height = pptSlide.height;

        const widthPx = width * 72;
        const heightPx = height * 72;

        return {
            width: width,
            height: height,
            widthPx: widthPx,
            heightPx: heightPx,
            widthInches: width,
            heightInches: height
        };
    } catch (error) {
        console.warn('Error getting slide dimensions:', error.message);
        return {
            width: 10,
            height: 7.5,
            widthPx: 960,
            heightPx: 720,
            widthInches: 10,
            heightInches: 7.5
        };
    }
}

function getShapeTypeForXml(shapeId) {
    const shapeTypeMap = {
        'rect': 'rect',
        'ellipse': 'ellipse',
        'triangle': 'triangle',
        'diamond': 'diamond',
        'pentagon': 'pentagon',
        'hexagon': 'hexagon',
        'octagon': 'octagon',
        'star5': 'star5',
        'roundRect': 'roundRect',
        'rightArrow': 'rightArrow',
        'leftArrow': 'leftArrow',
        'upArrow': 'upArrow',
        'downArrow': 'downArrow',
        'homePlate': 'pentagon'  // Added mapping for homePlate
    };

    return shapeTypeMap[shapeId] || 'rect';
}

function parseGradient(backgroundCss) {
    if (!backgroundCss) return null;

    const gradientRegex = /(linear|radial)-gradient\s*\((.*)\)/i;
    const match = backgroundCss.match(gradientRegex);
    if (!match) return null;

    const type = match[1].toLowerCase();
    const gradientArgs = match[2];
    const colorStops = [];

    // Extract color stops safely
    let current = '';
    let depth = 0;
    for (let i = 0; i < gradientArgs.length; i++) {
        const ch = gradientArgs[i];
        if (ch === '(') depth++;
        if (ch === ')') depth--;
        if (ch === ',' && depth === 0) {
            colorStops.push(current.trim());
            current = '';
        } else {
            current += ch;
        }
    }
    if (current.trim()) colorStops.push(current.trim());

    const stops = [];
    colorStops.forEach((stop) => {
        const rgba = stop.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([0-9.]+))?\)/);
        const hex = stop.match(/#([0-9a-f]{3,8})/i);
        const pos = stop.match(/(\d+)%/);
        if (rgba) {
            stops.push({
                color: `rgba(${rgba[1]}, ${rgba[2]}, ${rgba[3]}, ${rgba[4] || 1})`,
                position: pos ? parseFloat(pos[1]) : 100
            });
        } else if (hex) {
            stops.push({
                color: `#${hex[1]}`,
                position: pos ? parseFloat(pos[1]) : 100
            });
        }
    });

    if (!stops.length) return null;
    return {
        type,
        stops,
        dominantColor: stops[stops.length - 1].color
    };
}


function rgbToHex(color) {
    if (!color) return '';

    if (color.startsWith('#')) return color;

    if (color.startsWith('rgba')) {
        const rgbaRegex = /rgba\((\d+),\s*(\d+),\s*(\d+),\s*([0-9.]+)\)/;
        const match = color.match(rgbaRegex);

        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);

            return '#' +
                r.toString(16).padStart(2, '0') +
                g.toString(16).padStart(2, '0') +
                b.toString(16).padStart(2, '0');
        }
    } else if (color.startsWith('rgb')) {
        const rgbRegex = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/;
        const match = color.match(rgbRegex);

        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);

            return '#' +
                r.toString(16).padStart(2, '0') +
                g.toString(16).padStart(2, '0') +
                b.toString(16).padStart(2, '0');
        }
    }

    return '#000000';
}

function getDominantColorFromGradient(gradient) {
    if (!gradient) return null;

    if (gradient.dominantColor) {
        return gradient.dominantColor;
    }

    if (gradient.stops && gradient.stops.length > 0) {
        return gradient.stops[0].color;
    }

    return null;
}



function getBorderDashFromStyle(style, shapeElement) {
    // First try data attribute (most reliable)
    const dataDash = shapeElement.getAttribute("data-border-dash");
    if (dataDash) {
        return mapDashStyleToPptx(dataDash);
    }

    // Try CSS border-style
    const borderStyle = style.borderStyle;
    if (borderStyle && borderStyle !== 'solid') {
        const cssToXmlMap = {
            'dashed': 'dash',
            'dotted': 'dot'
        };
        return cssToXmlMap[borderStyle] || 'solid';
    }

    return 'solid';
}

module.exports = {
    addShapeToSlide,
};