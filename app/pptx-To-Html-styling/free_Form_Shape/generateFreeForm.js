
function generateCustomShapeSVG(custGeom, position, fillColor, stroke) {

    const viewBoxWidth = position.width;
    const viewBoxHeight = position.height;

    if (!custGeom || !custGeom["a:pathLst"] || !custGeom["a:pathLst"][0] || !custGeom["a:pathLst"][0]["a:path"]) {
        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none">
            <rect width="${viewBoxWidth}" height="${viewBoxHeight}" fill="${fillColor}" stroke="${stroke.color}" stroke-width="${stroke.width}"/>
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
                    const fmla = gd["$"].fmla;
                    guides[name] = evaluateFormula(fmla, originalWidth, originalHeight, guides);
                }
            }
        }

        function scaleX(x) {
            // If x is a guide reference, resolve it first
            if (typeof x === 'string' && guides[x] !== undefined) {
                x = guides[x];
            }
            return (parseFloat(x) / originalWidth) * viewBoxWidth;
        }

        function scaleY(y) {
            // If y is a guide reference, resolve it first
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
            // Ultimate fallback
            combinedPathData += `M0,0 L${viewBoxWidth},0 L${viewBoxWidth},${viewBoxHeight} L0,${viewBoxHeight} Z `;
        }
    }

    combinedPathData = combinedPathData.trim();

    // Handle gradient fills
    const gradientRegex = /linear-gradient\((\d+)deg, (rgba?\([^\)]+\)) (\d+%)?, (rgba?\([^\)]+\)) (\d+%)?\)/;
    const match = fillColor.match(gradientRegex);

    if (match) {
        const angle = match[1];
        const color1 = match[2];
        const offset1 = match[3] || "100%";
        const color2 = match[4];
        const offset2 = match[5] || "0%";
        const gradientId = `gradient_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none">
           <defs>
               <linearGradient id="${gradientId}" x1="100%" y1="0%" x2="0%" y2="0%">
                   <stop offset="${offset2}" style="stop-color:${color2}; stop-opacity:1" />
                   <stop offset="${offset1}" style="stop-color:${color1}; stop-opacity:1" />
               </linearGradient>
           </defs>
           <path d="${combinedPathData}" fill="url(#${gradientId})" stroke="${stroke.color || '#000'}" stroke-width="${stroke.width || 0}"/>
       </svg>`;
    } else {
        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none">
           <path d="${combinedPathData}" fill="${fillColor}" stroke="${stroke.color || '#000'}" stroke-width="${stroke.width || 0}"/>
       </svg>`;
    }
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

    // Process cubic bezier curves
    for (const cubicBez of cubicBezTos) {
        const points = getCubicBezierPoints(cubicBez);
        if (points && points.length === 3) {
            pathData += `C${points[0].x},${points[0].y} ${points[1].x},${points[1].y} ${points[2].x},${points[2].y} `;
        }
    }

    // Process line segments
    for (const lineTo of lineTos) {
        const pt = getPoint(lineTo);
        if (pt) {
            pathData += `L${pt.x},${pt.y} `;
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