function generateHexagonSVG({
    width,
    height,
    fillColor = "#cccccc",
    strokeColor = "#000",
    strokeWidth = 1,
    depth = 0,
    rotation = 0,
    opacity = 1,
    hexagonAdj = 50000,  // Default for manual override
    hexagonVF = 100000,  // Default for manual override
    extrusionColor = null,
    contourColor = null,
    contourWidth = 0,
    material = 'matte',
    lightDirection = 't',
    ignoreRotation = true,
    useDefaultAdj = true,  // Use PowerPoint's default adj value
    xmlData = null,  // XML data input parameter
}) {
    const w = width;
    const h = height;

    // Extract adj and vf from the provided XML data
    let effectiveAdj = hexagonAdj;
    let effectiveVF = hexagonVF;

    // Only extract from XML if `useDefaultAdj` is false
    if (xmlData && !useDefaultAdj) {
        const adjValue = extractAdjFromXML(xmlData);
        const vfValue = extractVFFromXML(xmlData);

        if (adjValue !== null) {
            effectiveAdj = adjValue;
        }
        if (vfValue !== null) {
            effectiveVF = vfValue;
        }

        // Extract rotation value from XML
        const rotationEMUs = extractRotationFromXML(xmlData);
        const rotationDegrees = convertEMUToDegrees(rotationEMUs);  // Convert to degrees
        if (rotationDegrees !== null) {
            rotation = rotationDegrees;  // Use the extracted rotation value
        }
    }

    // Apply or ignore rotation based on flag
    const effectiveRotation = ignoreRotation ? 0 : rotation;

    // Log which adj value is being used
    if (useDefaultAdj) {
        console.log(`ℹ️ Using PowerPoint default adj: 25000 (instead of extracted XML adj: ${effectiveAdj})`);
    } else {
        console.log(`ℹ️ Using extracted adj value: ${effectiveAdj}`);
    }

    // Calculate hexagon with effective adj value
    const front = calculateCorrectPowerPointHexagon(w, h, effectiveAdj, effectiveVF);

    // For 2D hexagon (no depth)
    if (depth <= 0) {
        const pointsString = front.map(p => p.join(",")).join(" ");

        // Only add rotation transform if not ignoring rotation
        const rotationTransform = effectiveRotation !== 0
            ? `<g transform="rotate(${effectiveRotation} ${w / 2} ${h / 2})">`
            : '';
        const rotationTransformClose = effectiveRotation !== 0 ? '</g>' : '';

        return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" style="overflow: visible;">
            ${rotationTransform}
                <polygon points="${pointsString}" 
                         fill="${fillColor}" 
                         stroke="${strokeColor}" 
                         stroke-width="${strokeWidth}"
                         opacity="${opacity}" />
            ${rotationTransformClose}
        </svg>`;
    }

    // 3D hexagon logic
    const perspective = calculatePowerPointPerspective(depth);
    const offsetX = perspective.offsetX;
    const offsetY = perspective.offsetY;

    const back = front.map(([x, y]) => [x + offsetX, y + offsetY]);
    const lighting = calculateAccurateLighting(fillColor, extrusionColor, lightDirection);

    const allPoints = [...front, ...back];
    const minX = Math.min(...allPoints.map(p => p[0]));
    const maxX = Math.max(...allPoints.map(p => p[0]));
    const minY = Math.min(...allPoints.map(p => p[1]));
    const maxY = Math.max(...allPoints.map(p => p[1]));

    const padding = 5;
    const viewBoxWidth = maxX - minX + padding * 2;
    const viewBoxHeight = maxY - minY + padding * 2;
    const viewBoxX = minX - padding;
    const viewBoxY = minY - padding;

    const adjustedFront = front.map(([x, y]) => [x - viewBoxX, y - viewBoxY]);
    const adjustedBack = back.map(([x, y]) => [x - viewBoxX, y - viewBoxY]);

    const faces = generateHexagon3DFaces(adjustedFront, adjustedBack, lighting, extrusionColor);
    const gradientDefs = generateAccurateGradients(lighting, fillColor);
    const contourEffect = contourWidth > 0 ? generatePowerPointContour(adjustedFront, contourColor, contourWidth) : '';

    const frontPointsString = adjustedFront.map(p => p.join(",")).join(" ");

    // Calculate rotation center accounting for viewBox offset
    const centerX = w / 2;
    const centerY = h / 2;
    const transformedCenterX = centerX - viewBoxX;
    const transformedCenterY = centerY - viewBoxY;

    // Only add rotation transform if not ignoring rotation
    const rotationTransform = effectiveRotation !== 0
        ? `<g transform="rotate(${effectiveRotation} ${transformedCenterX} ${transformedCenterY})">`
        : '';
    const rotationTransformClose = effectiveRotation !== 0 ? '</g>' : '';

    return `<svg width="100%" height="100%" viewBox="${viewBoxX} ${viewBoxY} ${viewBoxWidth} ${viewBoxHeight}" xmlns="http://www.w3.org/2000/svg" style="overflow: visible;">
        <defs>
            ${gradientDefs}
        </defs>
        
        ${rotationTransform}
            ${faces.map(face =>
        `<polygon points="${face.points}"
                          fill="${face.color}" 
                          stroke="${strokeColor}" 
                          stroke-width="${strokeWidth}"
                          opacity="0.95" />`
    ).join('\n')}
            
            ${contourEffect}
            
            <polygon points="${frontPointsString}" 
                     fill="url(#frontGradient)" 
                     stroke="${strokeColor}" 
                     stroke-width="${strokeWidth}"
                     opacity="${opacity}" />
        ${rotationTransformClose}
    </svg>`;
}

/**
 * Extracts the 'adj' value from the XML data
 */
function extractAdjFromXML(xmlData) {
    const adjMatch = xmlData.match(/<a:gd\s+name="adj"\s+fmla="val\s+(\d+)"/);
    return adjMatch ? parseInt(adjMatch[1]) : null;
}

/**
 * Extracts the 'vf' value from the XML data
 */
function extractVFFromXML(xmlData) {
    const vfMatch = xmlData.match(/<a:gd\s+name="vf"\s+fmla="val\s+(\d+)"/);
    return vfMatch ? parseInt(vfMatch[1]) : null;
}

/**
 * Extracts the 'rot' value (rotation) from the XML data
 */
function extractRotationFromXML(xmlData) {
    const rotationMatch = xmlData.match(/<a:xfrm\s+rot="(\d+)"/);
    return rotationMatch ? parseInt(rotationMatch[1]) : null;
}

/**
 * Converts the rotation value from EMU to degrees
 */
function convertEMUToDegrees(emu) {
    // PowerPoint uses 60000 EMUs per degree
    const degrees = emu / 60000;
    return degrees % 360; // Normalize to a 0-360 range
}

/**
 * Calculates standard OOXML hexagon points based on 'adj'.
 */
function calculateCorrectPowerPointHexagon(width, height, adj, vf) {
    const w = width;
    const h = height;
    const a = adj / 100000;

    return [
        [a * w, 0],
        [(1 - a) * w, 0],
        [w, h / 2],
        [(1 - a) * w, h],
        [a * w, h],
        [0, h / 2]
    ];
}

/**
 * Function to extract the 3D properties like depth, extrusion, etc.
 */
function extractHexagon3DProperties(shapeNode, options = {}) {
    const spPr = shapeNode?.["p:spPr"]?.[0];
    const sp3d = spPr?.["a:sp3d"]?.[0];
    const scene3d = spPr?.["a:scene3d"]?.[0];
    const xfrm = spPr?.["a:xfrm"]?.[0];
    const prstGeom = spPr?.["a:prstGeom"]?.[0];

    const properties = {
        depth: 0,
        material: 'matte',
        lightDirection: 't',
        rotation: 0,
        bevel: null,
        opacity: 1,
        extrusionColor: null,
        contourColor: null,
        contourWidth: 0,
        hexagonAdj: 25000,    // PowerPoint's default
        hexagonVF: 100000,
        ignoreRotation: options.ignoreRotation || false,
        useDefaultAdj: options.useDefaultAdj || false
    };

    // Extract rotation
    if (xfrm?.["$"]?.rot) {
        const rotValue = parseInt(xfrm["$"].rot);
        const rotationDegrees = (rotValue / 60000) % 360;
        properties.rotation = rotationDegrees;

        if (options.ignoreRotation) {
            console.log(`ℹ️  XML rotation: ${rotationDegrees.toFixed(2)}° (IGNORED)`);
        } else {
            console.log(`✅ XML rotation: ${rotationDegrees.toFixed(2)}° (APPLIED)`);
        }
    }

    // Extract hexagon adjustment values
    if (prstGeom?.["a:avLst"]?.[0]?.["a:gd"]) {
        const adjustments = prstGeom["a:avLst"][0]["a:gd"];

        adjustments.forEach(gd => {
            const name = gd["$"]?.name;
            const formula = gd["$"]?.fmla;

            if (name === "adj" && formula?.startsWith("val ")) {
                const xmlAdj = parseInt(formula.replace("val ", ""));
                properties.hexagonAdj = xmlAdj;

                if (options.useDefaultAdj) {
                    console.log(`ℹ️  XML adj: ${xmlAdj} (${(xmlAdj / 1000).toFixed(2)}%) - Using default 25000 (25%) instead`);
                } else {
                    console.log(`✅ Using XML adj: ${xmlAdj} (${(xmlAdj / 1000).toFixed(2)}%)`);
                }
            } else if (name === "vf" && formula?.startsWith("val ")) {
                properties.hexagonVF = parseInt(formula.replace("val ", ""));
            }
        });
    }

    // Extract 3D properties
    if (sp3d) {
        if (sp3d["$"]?.extrusionH) {
            properties.depth = parseInt(sp3d["$"].extrusionH) / 12700;
        }
        if (sp3d["$"]?.contourW) {
            properties.contourWidth = parseInt(sp3d["$"].contourW) / 12700;
        }
        const extrusionClr = sp3d["a:extrusionClr"]?.[0];
        if (extrusionClr?.["a:srgbClr"]?.[0]?.["$"]?.val) {
            properties.extrusionColor = `#${extrusionClr["a:srgbClr"][0]["$"].val}`;
        }
    }

    if (scene3d) {
        const lightRig = scene3d["a:lightRig"]?.[0];
        if (lightRig) {
            properties.lightDirection = lightRig["$"]?.dir || "t";
        }
    }

    return properties;
}

// Helper functions
function calculatePowerPointPerspective(depth) {
    const isoAngle = 30 * Math.PI / 180;
    return {
        offsetX: depth * Math.cos(isoAngle) * 0.4,
        offsetY: -depth * Math.sin(isoAngle) * 0.6
    };
}

function calculateAccurateLighting(baseColor, extrusionColor, lightDirection) {
    const lightFactors = {
        't': { front: 1.0, sides: 0.85 },
        'tr': { front: 1.0, sides: 0.80 },
        'br': { front: 1.0, sides: 0.75 },
        'b': { front: 1.0, sides: 0.80 },
        'bl': { front: 1.0, sides: 0.75 },
        'tl': { front: 1.0, sides: 0.85 }
    };

    const factors = lightFactors[lightDirection] || lightFactors.t;

    return {
        frontColor: baseColor,
        sideColor: extrusionColor || adjustColorBrightness(baseColor, factors.sides),
        backColor: adjustColorBrightness(baseColor, 0.3)
    };
}

function generateHexagon3DFaces(front, back, lighting, extrusionColor) {
    const faces = [];
    const visibleFaceIndices = [0, 1, 5];

    visibleFaceIndices.forEach(i => {
        const next = (i + 1) % 6;
        const f1 = front[i];
        const f2 = front[next];
        const b1 = back[i];
        const b2 = back[next];

        faces.push({
            points: `${f1.join(",")} ${f2.join(",")} ${b2.join(",")} ${b1.join(",")}`,
            color: extrusionColor || lighting.sideColor
        });
    });

    return faces;
}

function generateAccurateGradients(lighting, baseColor) {
    return `
        <linearGradient id="frontGradient" x1="30%" y1="30%" x2="70%" y2="70%">
            <stop offset="0%" style="stop-color:${adjustColorBrightness(baseColor, 1.15)};stop-opacity:1" />
            <stop offset="60%" style="stop-color:${baseColor};stop-opacity:1" />
            <stop offset="100%" style="stop-color:${adjustColorBrightness(baseColor, 0.85)};stop-opacity:1" />
        </linearGradient>`;
}

function generatePowerPointContour(front, contourColor, contourWidth) {
    if (!contourColor || contourWidth <= 0) return '';
    const frontPointsString = front.map(p => p.join(",")).join(" ");
    return `<polygon points="${frontPointsString}" 
                     fill="none" 
                     stroke="${contourColor}" 
                     stroke-width="${contourWidth}"
                     opacity="0.8" />`;
}

function adjustColorBrightness(hexColor, factor) {
    const hex = hexColor.replace('#', '');
    const r = Math.min(255, Math.max(0, parseInt(hex.substr(0, 2), 16) * factor));
    const g = Math.min(255, Math.max(0, parseInt(hex.substr(2, 2), 16) * factor));
    const b = Math.min(255, Math.max(0, parseInt(hex.substr(4, 2), 16) * factor));

    const toHex = (val) => Math.round(val).toString(16).padStart(2, '0');
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

module.exports = {
    generateHexagonSVG,
    extractHexagon3DProperties
};
