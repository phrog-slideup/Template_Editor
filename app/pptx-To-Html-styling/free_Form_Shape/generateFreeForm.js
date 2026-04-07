const { getCustomShapeShadowStyle, getCustomShapeShadowData } = require("../shapes_Properties/getShapeShadowStyle");

function generateCustomShapeSVG(shapeNode, custGeom, position, fillColor, stroke, flipOptions = {}) {

    const viewBoxWidth = position.width;
    const viewBoxHeight = position.height;

    // Shadow routing:
    //   Outer shadows  -> filter: drop-shadow() on <svg>  (clips to path outline)
    //   Inner shadows  -> box-shadow: inset on <foreignObject><div> for radial branches
    //                     (CSS drop-shadow has no inset, so inner shadows on solid/linear
    //                      branches are approximated as a zero-offset glow on <svg>)
    const shadowData = getCustomShapeShadowData(shapeNode, this.themeXML, this.masterXML, this.clrMap);

    // Outer shadow style for <svg> element (used by all branches)
    const svgFilterStyle = (!shadowData || shadowData.isInner)
        ? ""
        : ` style="${shadowData.svgFilter}"`;

    // Inner shadow approximation for solid/linear branches (zero-offset drop-shadow glow)
    const svgInnerApproxStyle = (shadowData?.isInner)
        ? (() => {
            const blurMatch = shadowData.boxShadow.match(/(\d+(?:\.\d+)?)px\s+rgba/);
            const rgbaMatch = shadowData.boxShadow.match(/rgba\([^)]+\)/);
            const blur = blurMatch ? blurMatch[1] : "0";
            const rgba = rgbaMatch ? rgbaMatch[0] : "rgba(0,0,0,1)";
            return ` style="filter: drop-shadow(0px 0px ${blur}px ${rgba});"`;
        })()
        : "";

    // ✅ FIX 1: Build SVG-level flip transform for the <path> element.
    const { flipH = false, flipV = false } = flipOptions;
    let pathTransform = "";
    // flip issue change start
    let pathTransformAttr = "";
    if (flipH && flipV) {
        pathTransform = `translate(${viewBoxWidth},${viewBoxHeight}) scale(-1,-1)`;
    } else if (flipH) {
        pathTransform = `translate(${viewBoxWidth},0) scale(-1,1)`;
    } else if (flipV) {
        pathTransform = `translate(0,${viewBoxHeight}) scale(1,-1)`;
    }

    pathTransformAttr = pathTransform ? `transform="${pathTransform}"` : "";
    // flip issue change end
    if (!custGeom || !custGeom["a:pathLst"] || !custGeom["a:pathLst"][0] || !custGeom["a:pathLst"][0]["a:path"]) {
        const _rectSvgStyle = svgFilterStyle || svgInnerApproxStyle;
        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${_rectSvgStyle}>
            <rect width="${viewBoxWidth}" height="${viewBoxHeight}" fill="${fillColor}" stroke="${stroke.color}" stroke-width="${stroke.width}" ${pathTransform}/>
        </svg>`;
    }

    let combinedPathData = "";
    const paths = custGeom["a:pathLst"][0]["a:path"];

    for (let pathIndex = 0; pathIndex < paths.length; pathIndex++) {
        const path = paths[pathIndex];

        // Get original dimensions for scaling
        const originalWidth = parseFloat(path["$"]?.w || viewBoxWidth);
        const originalHeight = parseFloat(path["$"]?.h || viewBoxHeight);

        // Process guide list (a:gdLst) if present
        const guides = {};
        if (custGeom["a:gdLst"] && custGeom["a:gdLst"][0] && custGeom["a:gdLst"][0]["a:gd"]) {
            const gdList = Array.isArray(custGeom["a:gdLst"][0]["a:gd"])
                ? custGeom["a:gdLst"][0]["a:gd"]
                : [custGeom["a:gdLst"][0]["a:gd"]];

            for (const gd of gdList) {
                if (gd && gd["$"]) {
                    const name = gd["$"].name;
                    // ✅ FIX 2: Skip connsite* guide names. PowerPoint repeats them many
                    // times (once per edit-history entry) and they are only used by
                    // <a:cxnLst> connection sites — never for path coordinate resolution.
                    // Including them pollutes the guides map and can corrupt scaleX/scaleY.
                    if (/^connsite/.test(name)) continue;
                    const fmla = gd["$"].fmla;
                    guides[name] = evaluateFormula(fmla, originalWidth, originalHeight, guides);
                }
            }
        }

        function scaleX(x) {
            if (typeof x === 'string' && guides[x] !== undefined) {
                x = guides[x];
            }
            return (parseFloat(x) / originalWidth) * viewBoxWidth;
        }

        function scaleY(y) {
            if (typeof y === 'string' && guides[y] !== undefined) {
                y = guides[y];
            }
            return (parseFloat(y) / originalHeight) * viewBoxHeight;
        }

        try {
            const pathData = parsePathCommandsDynamic(path, scaleX, scaleY);
            combinedPathData += pathData + " ";
        } catch (error) {
            console.error("Error processing custom shape:", error);
            combinedPathData += `M0,0 L${viewBoxWidth},0 L${viewBoxWidth},${viewBoxHeight} L0,${viewBoxHeight} Z `;
        }
    }

    combinedPathData = combinedPathData.trim();

    // ✅ FIX 3: Convert CSS linear-gradient angle to correct SVG linearGradient vector.
    // radial-gradient pass-through for non-linear fills.
    const isLinearGradient = /^linear-gradient\(/.test(fillColor);
    const isRadialGradient = /^radial-gradient\(/.test(fillColor);

    if (isLinearGradient) {
        // Extract angle and all stops from the CSS linear-gradient string
        const angleMatch = fillColor.match(/linear-gradient\((\d+(?:\.\d+)?)deg,\s*([\s\S]+)\)/);
        if (angleMatch) {
            const angleDeg = parseFloat(angleMatch[1]);
            const stopsStr = angleMatch[2];

            // Convert CSS angle → SVG gradient vector
            const rad = (angleDeg * Math.PI) / 180;
            const sinA = Math.sin(rad);
            const cosA = Math.cos(rad);
            const x1 = ((0.5 - 0.5 * sinA) * 100).toFixed(2);
            const y1 = ((0.5 + 0.5 * cosA) * 100).toFixed(2);
            const x2 = ((0.5 + 0.5 * sinA) * 100).toFixed(2);
            const y2 = ((0.5 - 0.5 * cosA) * 100).toFixed(2);

            // Parse individual stops: each is "rgba(...) XX%" or "#hex XX%"
            const stopRegex = /(rgba?\([^)]+\)|#[0-9a-fA-F]+)\s+(\d+(?:\.\d+)?%)/g;
            const svgStops = [];
            let m;
            while ((m = stopRegex.exec(stopsStr)) !== null) {
                svgStops.push({ color: m[1], offset: m[2] });
            }

            // Sort stops by offset ascending (SVG spec requirement)
            svgStops.sort((a, b) => parseFloat(a.offset) - parseFloat(b.offset));

            const gradientId = `gradient_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            const stopElements = svgStops
                .map(s => `<stop offset="${s.offset}" style="stop-color:${s.color}; stop-opacity:1"/>`)
                .join("\n               ");
            //flip issue change
            const _linSvgStyle = svgFilterStyle || svgInnerApproxStyle;
            return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${_linSvgStyle}>
           <defs>
               <linearGradient id="${gradientId}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%">
               ${stopElements}
               </linearGradient>
           </defs>
          <g ${pathTransformAttr}>
    <path d="${combinedPathData}" fill="url(#${gradientId})" stroke="${stroke.color || 'none'}" stroke-width="${stroke.width || 0}" />
</g>
       </svg>`;
        }
    }

    if (isRadialGradient) {
        // Radial gradients: use CSS directly via foreignObject trick is unreliable in SVG.
        // Best approach: pass through as a CSS background on a <foreignObject> rect,
        // clipped to the path shape. We use a clipPath for correctness.
        const clipId = `clip_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        //flip issue change
        // For inner shadows on radial-gradient shapes we apply box-shadow:inset on
        // the clipped <div> — the only element that can honour "inset" here.
        // For outer shadows we keep filter:drop-shadow on <svg> as usual.
        const _radDivExtra = (shadowData?.isInner)
            ? `box-shadow:${shadowData.boxShadow};`
            : "";
        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${svgFilterStyle}>
           <defs>
               <clipPath id="${clipId}">
                   <path d="${combinedPathData}" />
               </clipPath>
           </defs>
           <foreignObject width="${viewBoxWidth}" height="${viewBoxHeight}" clip-path="url(#${clipId})">
               <div xmlns="http://www.w3.org/1999/xhtml"
                    style="width:100%;height:100%;${_radDivExtra}background:${fillColor};"></div>
           </foreignObject>
       </svg>`;
    }

    // Solid fill fallback
    //flip issue change
    const _solidSvgStyle = svgFilterStyle || svgInnerApproxStyle;
    return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${_solidSvgStyle}>
       <g ${pathTransformAttr}>
    <path d="${combinedPathData}" fill="${fillColor}" stroke="${stroke.color || 'none'}" stroke-width="${stroke.width || 0}" />
</g>
   </svg>`;
}


function getShadowStyle(shapeNode) {
    return getShapeShadowStyle(shapeNode, this.themeXML, this.masterXML, this.clrMap);
}

// Evaluate DrawingML formulas used in guide definitions
function evaluateFormula(fmla, width, height, guides) {
    if (!fmla) return 0;

    // Parse formula format: operator param1 param2 param3
    const parts = fmla.trim().split(/\s+/);
    const operator = parts[0];

    // Helper to resolve a value (could be a number, 'w', 'h', or guide reference)
    const resolveValue = (val) => {
        if (val === 'w') return width;
        if (val === 'h') return height;
        if (val === 'l') return 0; // left
        if (val === 't') return 0; // top
        if (val === 'r') return width; // right
        if (val === 'b') return height; // bottom
        if (guides[val] !== undefined) return guides[val];
        return parseFloat(val) || 0;
    };

    switch (operator) {
        case '*/': // Multiply and divide: */ param1 param2 param3 = (param1 * param2) / param3
            if (parts.length >= 4) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                const p3 = resolveValue(parts[3]);
                return p3 !== 0 ? (p1 * p2) / p3 : 0;
            }
            break;
        case '*': // Multiply
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return p1 * p2;
            }
            break;
        case '/': // Divide
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return p2 !== 0 ? p1 / p2 : 0;
            }
            break;
        case '+': // Add
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return p1 + p2;
            }
            break;
        case '-': // Subtract
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return p1 - p2;
            }
            break;
        case 'val': // Direct value
            if (parts.length >= 2) {
                return resolveValue(parts[1]);
            }
            break;
        case '+-': // Add or subtract
            if (parts.length >= 4) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                const p3 = resolveValue(parts[3]);
                return p1 + p2 - p3;
            }
            break;
        case 'abs': // Absolute value
            if (parts.length >= 2) {
                return Math.abs(resolveValue(parts[1]));
            }
            break;
        case 'max': // Maximum
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return Math.max(p1, p2);
            }
            break;
        case 'min': // Minimum
            if (parts.length >= 3) {
                const p1 = resolveValue(parts[1]);
                const p2 = resolveValue(parts[2]);
                return Math.min(p1, p2);
            }
            break;
        case 'sqrt': // Square root
            if (parts.length >= 2) {
                return Math.sqrt(resolveValue(parts[1]));
            }
            break;
        case 'sin': // Sine (input in 60000ths of a degree)
            if (parts.length >= 3) {
                const angle = resolveValue(parts[1]);
                const mult = resolveValue(parts[2]);
                return Math.sin((angle / 60000) * Math.PI / 180) * mult;
            }
            break;
        case 'cos': // Cosine
            if (parts.length >= 3) {
                const angle = resolveValue(parts[1]);
                const mult = resolveValue(parts[2]);
                return Math.cos((angle / 60000) * Math.PI / 180) * mult;
            }
            break;
        case '?:': // Ternary: ?: condition if_true if_false
            if (parts.length >= 4) {
                const condition = resolveValue(parts[1]);
                const ifTrue = resolveValue(parts[2]);
                const ifFalse = resolveValue(parts[3]);
                return condition > 0 ? ifTrue : ifFalse;
            }
            break;
    }

    // If we can't parse it, try to return it as a number
    return parseFloat(fmla) || 0;
}

// New dynamic path parser that handles any path structure
function parsePathCommandsDynamic(path, scaleX, scaleY) {
    let pathData = "";

    // Helper function to extract point coordinates
    const getPoint = (item) => {
        if (item && item["a:pt"] && item["a:pt"][0] && item["a:pt"][0]["$"]) {
            return {
                x: scaleX(item["a:pt"][0]["$"].x),
                y: scaleY(item["a:pt"][0]["$"].y)
            };
        }
        return null;
    };

    // Helper function to extract cubic bezier points
    const getCubicBezierPoints = (item) => {
        if (item && item["a:pt"] && Array.isArray(item["a:pt"]) && item["a:pt"].length === 3) {
            return item["a:pt"].map(pt => ({
                x: scaleX(pt["$"].x),
                y: scaleY(pt["$"].y)
            }));
        }
        return null;
    };

    // Try to parse in the order commands appear in XML
    if (tryParseInXMLOrder) {
        const result = tryParseInXMLOrder(path, scaleX, scaleY);
        if (result) return result;
    }

    // Extract all command types
    const moveTos = path["a:moveTo"] ? (Array.isArray(path["a:moveTo"]) ? path["a:moveTo"] : [path["a:moveTo"]]) : [];
    const lineTos = path["a:lnTo"] ? (Array.isArray(path["a:lnTo"]) ? path["a:lnTo"] : [path["a:lnTo"]]) : [];
    const cubicBezTos = path["a:cubicBezTo"] ? (Array.isArray(path["a:cubicBezTo"]) ? path["a:cubicBezTo"] : [path["a:cubicBezTo"]]) : [];
    const quadBezTos = path["a:quadBezTo"] ? (Array.isArray(path["a:quadBezTo"]) ? path["a:quadBezTo"] : [path["a:quadBezTo"]]) : [];
    const arcTos = path["a:arcTo"] ? (Array.isArray(path["a:arcTo"]) ? path["a:arcTo"] : [path["a:arcTo"]]) : [];
    const closes = path["a:close"] ? (Array.isArray(path["a:close"]) ? path["a:close"] : [path["a:close"]]) : [];

    // Build path data sequentially
    let commandIndex = 0;
    let lineToIndex = 0;
    let cubicBezToIndex = 0;
    let quadBezToIndex = 0;
    let arcToIndex = 0;

    // Start with moveTo
    if (moveTos.length > 0) {
        const pt = getPoint(moveTos[0]);
        if (pt) {
            pathData += `M${pt.x},${pt.y} `;
        }
    }

    // ✅ FIX 4: Emit lnTo commands BEFORE cubicBezTo in the fallback path.
    //
    // xml2js (without explicitChildren) collapses all same-name child elements into
    // separate arrays, losing their interleaved XML document order. The most common
    // custGeom pattern is: moveTo → lnTo × N → cubicBezTo (closing arc) → close.
    // Emitting cubicBezTos first (the old order) produced: M→C→L→L→L→Z which is
    // completely wrong geometry. The correct fallback is lnTos-first → cubicBezTos.
    //
    // For shapes where a cubicBezTo appears BEFORE lnTos in the XML, you MUST
    // configure xml2js with { explicitChildren: true, preserveChildrenOrder: true }
    // so that tryParseInXMLOrder() handles it correctly above.

    // Process line segments first
    for (const lineTo of lineTos) {
        const pt = getPoint(lineTo);
        if (pt) {
            pathData += `L${pt.x},${pt.y} `;
        }
    }

    // Then cubic bezier curves (typically closing arcs at end of path)
    for (const cubicBez of cubicBezTos) {
        const points = getCubicBezierPoints(cubicBez);
        if (points && points.length === 3) {
            pathData += `C${points[0].x},${points[0].y} ${points[1].x},${points[1].y} ${points[2].x},${points[2].y} `;
        }
    }

    // Process quadratic bezier curves
    for (const quadBez of quadBezTos) {
        if (quadBez && quadBez["a:pt"] && Array.isArray(quadBez["a:pt"]) && quadBez["a:pt"].length === 2) {
            const pt1 = { x: scaleX(quadBez["a:pt"][0]["$"].x), y: scaleY(quadBez["a:pt"][0]["$"].y) };
            const pt2 = { x: scaleX(quadBez["a:pt"][1]["$"].x), y: scaleY(quadBez["a:pt"][1]["$"].y) };
            pathData += `Q${pt1.x},${pt1.y} ${pt2.x},${pt2.y} `;
        }
    }

    // Process arcs (simplified - may need more complex handling)
    for (const arc of arcTos) {
        if (arc && arc["$"]) {
            const wR = scaleX(arc["$"].wR || 0);
            const hR = scaleY(arc["$"].hR || 0);
            const stAng = parseFloat(arc["$"].stAng || 0) / 60000;
            const swAng = parseFloat(arc["$"].swAng || 0) / 60000;
            // This is a simplified arc handling - you may need to convert to SVG arc format
            pathData += `A${wR},${hR} 0 0 1 `;
        }
    }

    // Close path if needed
    if (closes.length > 0) {
        pathData += "Z ";
    }

    return pathData.trim();
}

// Optional: Enhanced XML order parser (if XML structure is accessible)
function tryParseInXMLOrder(path, scaleX, scaleY) {
    // This requires xml2js to be configured with:
    //   preserveChildrenOrder: true,
    //   explicitChildren: true,
    // so that child elements appear in `path.$$` in the same order as in XML.
    // If `path.$$` is missing, we return null and the caller will fall back.
    if (!path || !Array.isArray(path.$$) || path.$$.length === 0) return null;

    // Local helpers (mirror the ones in parsePathCommandsDynamic)
    const getPoint = (item) => {
        if (!item) return null;
        // Most xml2js configs: item["a:pt"][0]["$"]
        if (item["a:pt"] && item["a:pt"][0] && item["a:pt"][0]["$"]) {
            return {
                x: scaleX(item["a:pt"][0]["$"].x),
                y: scaleY(item["a:pt"][0]["$"].y)
            };
        }
        // Sometimes with explicitChildren, the pt can itself be a child in item.$$
        if (Array.isArray(item.$$)) {
            const ptChild = item.$$.find(ch => ch && (ch["a:pt"] || ch["#name"] === "a:pt"));
            if (ptChild && ptChild["$"]) {
                return { x: scaleX(ptChild["$"].x), y: scaleY(ptChild["$"].y) };
            }
            if (ptChild && ptChild["a:pt"] && ptChild["a:pt"][0] && ptChild["a:pt"][0]["$"]) {
                return { x: scaleX(ptChild["a:pt"][0]["$"].x), y: scaleY(ptChild["a:pt"][0]["$"].y) };
            }
        }
        return null;
    };

    const getCubicBezierPoints = (item) => {
        if (!item) return null;
        if (item["a:pt"] && Array.isArray(item["a:pt"]) && item["a:pt"].length === 3) {
            return item["a:pt"].map(pt => ({
                x: scaleX(pt["$"].x),
                y: scaleY(pt["$"].y)
            }));
        }
        // With explicitChildren, pts may be in item.$$
        if (Array.isArray(item.$$)) {
            const pts = item.$$.filter(ch => ch && (ch["#name"] === "a:pt" || ch["$"] && ch["$"].x !== undefined && ch["$"].y !== undefined));
            if (pts.length === 3) {
                return pts.map(pt => ({
                    x: scaleX(pt["$"].x),
                    y: scaleY(pt["$"].y)
                }));
            }
        }
        return null;
    };

    let d = "";
    for (const child of path.$$) {
        if (!child) continue;
        const name = child["#name"] || child.__name || child.name; // defensive
        switch (name) {
            case "a:moveTo": {
                const pt = getPoint(child);
                if (pt) d += `M${pt.x},${pt.y} `;
                break;
            }
            case "a:lnTo": {
                const pt = getPoint(child);
                if (pt) d += `L${pt.x},${pt.y} `;
                break;
            }
            case "a:cubicBezTo": {
                const pts = getCubicBezierPoints(child);
                if (pts && pts.length === 3) {
                    d += `C${pts[0].x},${pts[0].y} ${pts[1].x},${pts[1].y} ${pts[2].x},${pts[2].y} `;
                }
                break;
            }
            case "a:quadBezTo": {
                // 2 points: control + end
                if (child["a:pt"] && Array.isArray(child["a:pt"]) && child["a:pt"].length === 2) {
                    const p1 = { x: scaleX(child["a:pt"][0]["$"].x), y: scaleY(child["a:pt"][0]["$"].y) };
                    const p2 = { x: scaleX(child["a:pt"][1]["$"].x), y: scaleY(child["a:pt"][1]["$"].y) };
                    d += `Q${p1.x},${p1.y} ${p2.x},${p2.y} `;
                } else if (Array.isArray(child.$$)) {
                    const pts = child.$$.filter(ch => ch && ch["#name"] === "a:pt" && ch["$"]);
                    if (pts.length === 2) {
                        const p1 = { x: scaleX(pts[0]["$"].x), y: scaleY(pts[0]["$"].y) };
                        const p2 = { x: scaleX(pts[1]["$"].x), y: scaleY(pts[1]["$"].y) };
                        d += `Q${p1.x},${p1.y} ${p2.x},${p2.y} `;
                    }
                }
                break;
            }
            case "a:close": {
                d += "Z ";
                break;
            }
            // arcTo not used in your sample; keep no-op for now
            default:
                break;
        }
    }

    const out = d.trim();
    return out.length ? out : null;
}


// Handle single continuous path (most shapes)
function handleSinglePath(moveTos, lineTos, closes, getPoint) {
    let result = "";

    // Start with moveTo
    if (moveTos.length > 0) {
        const pt = getPoint(moveTos[0]);
        if (pt) {
            result += `M${pt.x},${pt.y} `;
        }
    }

    // Add all lineTo commands
    for (const lineTo of lineTos) {
        const pt = getPoint(lineTo);
        if (pt) {
            result += `L${pt.x},${pt.y} `;
        }
    }

    // Add close if present
    if (closes.length > 0) {
        result += "Z ";
    }

    return result;
}

// Handle multiple sub-paths
function handleMultiplePaths(moveTos, lineTos, closes, getPoint) {
    let result = "";

    // Try to intelligently distribute lineTo commands among moveTo commands
    const lineToPerMove = Math.floor(lineTos.length / moveTos.length);
    const remainder = lineTos.length % moveTos.length;

    let lineToIndex = 0;

    for (let i = 0; i < moveTos.length; i++) {
        // Add moveTo
        const pt = getPoint(moveTos[i]);
        if (pt) {
            result += `M${pt.x},${pt.y} `;
        }

        // Calculate how many lineTo commands for this sub-path
        let lineToCount = lineToPerMove;
        if (i < remainder) lineToCount += 1; // Distribute remainder

        // Add lineTo commands for this sub-path
        for (let j = 0; j < lineToCount && lineToIndex < lineTos.length; j++) {
            const lnPt = getPoint(lineTos[lineToIndex]);
            if (lnPt) {
                result += `L${lnPt.x},${lnPt.y} `;
            }
            lineToIndex++;
        }

        // Add close for this sub-path
        if (i < closes.length) {
            result += "Z ";
        }
    }

    return result;
}

// Fallback handler for unusual structures
function handleFallbackPath(moveTos, lineTos, closes, getPoint) {
    let result = "";

    // If no moveTo, start with a default moveTo
    if (moveTos.length === 0 && lineTos.length > 0) {
        const firstPt = getPoint(lineTos[0]);
        if (firstPt) {
            result += `M${firstPt.x},${firstPt.y} `;
            // Skip the first lineTo since we used it as moveTo
            for (let i = 1; i < lineTos.length; i++) {
                const pt = getPoint(lineTos[i]);
                if (pt) {
                    result += `L${pt.x},${pt.y} `;
                }
            }
        }
    } else {
        // Standard processing
        result = handleSinglePath(moveTos, lineTos, closes, getPoint);
    }

    return result;
}



module.exports = {
    generateCustomShapeSVG,
};