const { getCustomShapeShadowStyle, getCustomShapeShadowData } = require("../shapes_Properties/getShapeShadowStyle");

function normalizeSvgFill(fillColor) {
    if (!fillColor) return "none";
    if (fillColor === "transparent" || fillColor === "none") return "none";
    if (/^rgba\([^)]*,\s*0(?:\.0+)?\s*\)$/i.test(fillColor)) return "none";
    return fillColor;
}

function buildStrokeAttributes(stroke = {}) {
    let attrs = `stroke="${stroke.color || 'none'}" stroke-width="${stroke.width || 0}"`;
    if (stroke.join) attrs += ` stroke-linejoin="${stroke.join}"`;
    if (stroke.cap && stroke.cap !== "butt") attrs += ` stroke-linecap="${stroke.cap}"`;
    if (stroke.dashArray) attrs += ` stroke-dasharray="${stroke.dashArray}"`;
    return attrs;
}

function composeSvgStyle(...styleFragments) {
    const fragments = styleFragments
        .map(fragment => (fragment || "").trim())
        .filter(Boolean)
        .map(fragment => fragment.replace(/^style="/, "").replace(/"$/, "").trim().replace(/;$/, ""));

    const styleSet = new Set(["overflow:visible"]);
    fragments.forEach(fragment => {
        fragment.split(";").map(part => part.trim()).filter(Boolean).forEach(part => styleSet.add(part));
    });

    return ` style="${Array.from(styleSet).join(";")};"`;
}

function generateCustomShapeSVG(shapeNode, custGeom, position, fillColor, stroke, flipOptions = {}) {

    const viewBoxWidth = position.width;
    const viewBoxHeight = position.height;
    const svgFill = normalizeSvgFill(fillColor);
    const strokeAttrs = buildStrokeAttributes(stroke);

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
    const svgBaseStyle = composeSvgStyle(svgFilterStyle, svgInnerApproxStyle);

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
        return `<svg viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${svgBaseStyle}>
            <rect width="${viewBoxWidth}" height="${viewBoxHeight}" fill="${svgFill}" ${strokeAttrs} ${pathTransform}/>
        </svg>`;
    }

    let combinedPathData = "";
    const paths = custGeom["a:pathLst"][0]["a:path"];

    // ── Determine the viewBox coordinate space ──────────────────────────────
    // When <a:path> has explicit w/h attributes, those define the coordinate
    // space and we scale path coords to position.width/height as usual.
    // When w/h are ABSENT (common for PowerPoint freeforms), using
    // position.width/height (which is already in pixels, often sub-pixel) as
    // the fallback produces an identity scaleX/scaleY = 1, leaving path coords
    // raw inside a ~1×1 px viewBox.  The shape then overflows its wrapper div
    // via overflow:visible, giving wrong size and wrong hit-area.
    //
    // FIX: when w/h are absent, compute the natural bounding box of the path
    // commands themselves and use those dimensions as the viewBox coordinate
    // space.  The SVG viewBox is then set to those natural dimensions, and the
    // path coords map 1:1 — no scaling needed — while the SVG element still
    // renders at position.width × position.height pixels via width/height="100%".
    const firstPath = paths[0];
    const hasExplicitPathDims = firstPath?.["$"]?.w != null && firstPath?.["$"]?.h != null;

    let resolvedViewBoxWidth  = viewBoxWidth;
    let resolvedViewBoxHeight = viewBoxHeight;

    if (!hasExplicitPathDims) {
        // Scan all paths for the natural coordinate bounds
        let naturalBounds = null;
        for (const p of paths) {
            const b = computePathNaturalBounds(p);
            if (b) {
                naturalBounds = naturalBounds
                    ? { w: Math.max(naturalBounds.w, b.w), h: Math.max(naturalBounds.h, b.h) }
                    : b;
            }
        }
        if (naturalBounds) {
            // Use the natural path coordinate space as the viewBox dimensions.
            // The SVG will render at position.width × position.height via CSS
            // width/height="100%", so the shape fills the wrapper correctly.
            resolvedViewBoxWidth  = naturalBounds.w;
            resolvedViewBoxHeight = naturalBounds.h;
        }
    }

    for (let pathIndex = 0; pathIndex < paths.length; pathIndex++) {
        const path = paths[pathIndex];

        // Get original dimensions for scaling.
        // When w/h are absent on <a:path>, use the resolved natural bounds
        // (already computed above) so scaleX/scaleY produce correct proportions.
        // Do NOT fall back to viewBoxWidth/viewBoxHeight (pixel sizes) — that
        // causes an identity transform and leaves path coords raw.
        const originalWidth  = parseFloat(path["$"]?.w  || resolvedViewBoxWidth);
        const originalHeight = parseFloat(path["$"]?.h || resolvedViewBoxHeight);

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
            return (parseFloat(x) / originalWidth) * resolvedViewBoxWidth;
        }

        function scaleY(y) {
            if (typeof y === 'string' && guides[y] !== undefined) {
                y = guides[y];
            }
            return (parseFloat(y) / originalHeight) * resolvedViewBoxHeight;
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
    const isLinearGradient = /^linear-gradient\(/.test(svgFill);
    const isRadialGradient = /^radial-gradient\(/.test(svgFill);

    if (isLinearGradient) {
        // Extract angle and all stops from the CSS linear-gradient string
        const angleMatch = svgFill.match(/linear-gradient\((\d+(?:\.\d+)?)deg,\s*([\s\S]+)\)/);
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
                .map(s => `<stop offset="${s.offset}" style="stop-color:${s.color};"/>`)
                .join("\n               ");
            //flip issue change
            return `<svg viewBox="0 0 ${resolvedViewBoxWidth} ${resolvedViewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${svgBaseStyle}>
           <defs>
               <linearGradient id="${gradientId}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%">
               ${stopElements}
               </linearGradient>
           </defs>
          <g ${pathTransformAttr}>
    <path d="${combinedPathData}" fill="url(#${gradientId})" ${strokeAttrs} />
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
        return `<svg viewBox="0 0 ${resolvedViewBoxWidth} ${resolvedViewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${svgBaseStyle}>
           <defs>
               <clipPath id="${clipId}">
                   <path d="${combinedPathData}" />
               </clipPath>
           </defs>
           <foreignObject width="${resolvedViewBoxWidth}" height="${resolvedViewBoxHeight}" clip-path="url(#${clipId})">
               <div xmlns="http://www.w3.org/1999/xhtml"
                    style="width:100%;height:100%;${_radDivExtra}background:${svgFill};"></div>
           </foreignObject>
       </svg>`;
    }

    // Solid fill fallback
    //flip issue change
    return `<svg viewBox="0 0 ${resolvedViewBoxWidth} ${resolvedViewBoxHeight}" width="100%" height="100%" preserveAspectRatio="none"${svgBaseStyle}>
       <g ${pathTransformAttr}>
    <path d="${combinedPathData}" fill="${svgFill}" ${strokeAttrs} />
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

/**
 * computePathNaturalBounds(path)
 *
 * When an <a:path> element has no w/h attributes, the OOXML path coordinate
 * space cannot be inferred from position.width/height (which are already scaled
 * to pixels and are often sub-pixel for freeforms). Instead we scan every <a:pt>
 * inside the path commands to find the actual coordinate extents and use those
 * as the coordinate space dimensions.
 *
 * Returns { w, h } — the natural bounding box of the path command coordinates,
 * or null if no points are found.
 */
function computePathNaturalBounds(path) {
    let maxX = 0;
    let maxY = 0;
    let found = false;

    const scanPoints = (pts) => {
        if (!Array.isArray(pts)) return;
        for (const pt of pts) {
            const $ = pt?.["$"];
            if (!$) continue;
            const x = parseFloat($.x);
            const y = parseFloat($.y);
            if (Number.isFinite(x)) { maxX = Math.max(maxX, x); found = true; }
            if (Number.isFinite(y)) { maxY = Math.max(maxY, y); found = true; }
        }
    };

    const commandKeys = ["a:moveTo", "a:lnTo", "a:cubicBezTo", "a:quadBezTo"];
    for (const key of commandKeys) {
        const cmds = path[key];
        if (!cmds) continue;
        const list = Array.isArray(cmds) ? cmds : [cmds];
        for (const cmd of list) {
            if (cmd?.["a:pt"]) scanPoints(Array.isArray(cmd["a:pt"]) ? cmd["a:pt"] : [cmd["a:pt"]]);
        }
    }

    // Also scan explicitChildren order array ($$) if present
    if (Array.isArray(path.$$)) {
        for (const child of path.$$) {
            if (child?.["a:pt"]) scanPoints(Array.isArray(child["a:pt"]) ? child["a:pt"] : [child["a:pt"]]);
        }
    }

    if (!found || maxX === 0 && maxY === 0) return null;
    return { w: maxX || 1, h: maxY || 1 };
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

// In freeFormShape.js — add this export:
function generateCustomShapePathOnly(shapeNode, custGeom, position, options = {}) {
    try {
        const { skipFlip = false } = options;  // ✅ NEW option

        const pathLst = custGeom?.["a:pathLst"]?.[0]?.["a:path"] || [];
        if (!pathLst.length) return { pathD: "" };

        const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
        
        // ✅ Only read flip values if NOT skipping
        const flipH = !skipFlip && xfrm?.["$"]?.flipH === "1";
        const flipV = !skipFlip && xfrm?.["$"]?.flipV === "1";

        const W = position.width;
        const H = position.height;

        let allPathD = "";

        for (const pathNode of pathLst) {
            const pathW = parseFloat(pathNode?.["$"]?.w || 1);
            const pathH = parseFloat(pathNode?.["$"]?.h || 1);

            const scaleX = W / pathW;
            const scaleY = H / pathH;

            const scalePoint = (x, y) => {
                let sx = parseFloat(x) * scaleX;
                let sy = parseFloat(y) * scaleY;
                if (flipH) sx = W - sx;
                if (flipV) sy = H - sy;
                return { x: sx, y: sy };
            };

            const commands = pathNode?.["$$"] || [];
            let d = "";

            for (const cmd of commands) {
                const type = cmd["#name"];

                switch (type) {
                    case "a:moveTo": {
                        const pt = cmd?.["a:pt"]?.[0]?.["$"];
                        if (pt) {
                            const { x, y } = scalePoint(pt.x, pt.y);
                            d += `M ${x.toFixed(3)} ${y.toFixed(3)} `;
                        }
                        break;
                    }
                    case "a:lnTo": {
                        const pt = cmd?.["a:pt"]?.[0]?.["$"];
                        if (pt) {
                            const { x, y } = scalePoint(pt.x, pt.y);
                            d += `L ${x.toFixed(3)} ${y.toFixed(3)} `;
                        }
                        break;
                    }
                    case "a:cubicBezTo": {
                        const pts = cmd?.["a:pt"] || [];
                        if (pts.length === 3) {
                            const p1 = scalePoint(pts[0]["$"].x, pts[0]["$"].y);
                            const p2 = scalePoint(pts[1]["$"].x, pts[1]["$"].y);
                            const p3 = scalePoint(pts[2]["$"].x, pts[2]["$"].y);
                            d += `C ${p1.x.toFixed(3)} ${p1.y.toFixed(3)}, `;
                            d += `${p2.x.toFixed(3)} ${p2.y.toFixed(3)}, `;
                            d += `${p3.x.toFixed(3)} ${p3.y.toFixed(3)} `;
                        }
                        break;
                    }
                    case "a:quadBezTo": {
                        const pts = cmd?.["a:pt"] || [];
                        if (pts.length === 2) {
                            const p1 = scalePoint(pts[0]["$"].x, pts[0]["$"].y);
                            const p2 = scalePoint(pts[1]["$"].x, pts[1]["$"].y);
                            d += `Q ${p1.x.toFixed(3)} ${p1.y.toFixed(3)}, `;
                            d += `${p2.x.toFixed(3)} ${p2.y.toFixed(3)} `;
                        }
                        break;
                    }
                    case "a:arcTo": {
                        const $ = cmd?.["$"] || {};
                        const wR = parseFloat($.wR || 0) * scaleX;
                        const hR = parseFloat($.hR || 0) * scaleY;
                        const stAng = parseFloat($.stAng || 0) / 60000;
                        const swAng = parseFloat($.swAng || 0) / 60000;
                        const endAng = stAng + swAng;
                        const endX = (W / 2 + wR * Math.cos((endAng * Math.PI) / 180)).toFixed(3);
                        const endY = (H / 2 + hR * Math.sin((endAng * Math.PI) / 180)).toFixed(3);
                        const largeArc = Math.abs(swAng) > 180 ? 1 : 0;
                        const sweep = swAng > 0 ? 1 : 0;
                        d += `A ${wR.toFixed(3)} ${hR.toFixed(3)} 0 ${largeArc} ${sweep} ${endX} ${endY} `;
                        break;
                    }
                    case "a:close": {
                        d += "Z ";
                        break;
                    }
                    default:
                        break;
                }
            }

            allPathD += d.trim() + " ";
        }

        return { pathD: allPathD.trim() };

    } catch (err) {
        console.error("generateCustomShapePathOnly error:", err);
        return { pathD: "" };
    }
}



module.exports = {
    generateCustomShapeSVG,
    generateCustomShapePathOnly
};