const colorHelper = require("../../api/helper/colorHelper.js");
const shapeFillColor = require("../shapes_Properties/getShapeFillColor.js");

// Add helper method to get the correct divisor
function getEMUDivisor() {
    return 12700;
}

function getDashArray(dashType, strokeWidth) {
    const baseSize = Math.max(2, strokeWidth);
    
    switch (dashType) {
        case "solid":
            return "";
        case "dash":
            return `${baseSize * 4}px ${baseSize * 3}px`;
        case "dot":
            return `${baseSize}px ${baseSize * 2}px`;
        case "dashDot":
            return `${baseSize * 4}px ${baseSize * 2}px ${baseSize}px ${baseSize * 2}px`;
        case "lgDash":
            return `${baseSize * 8}px ${baseSize * 3}px`;
        case "lgDashDot":
            return `${baseSize * 8}px ${baseSize * 2}px ${baseSize}px ${baseSize * 2}px`;
        case "lgDashDotDot":
            return `${baseSize * 8}px ${baseSize * 2}px ${baseSize}px ${baseSize * 2}px ${baseSize}px ${baseSize * 2}px`;
        case "sysDash":
            return `${baseSize * 3}px ${baseSize}px`;
        case "sysDot":
            return `${baseSize}px ${baseSize}px`;
        case "sysDashDot":
            return `${baseSize * 3}px ${baseSize}px ${baseSize}px ${baseSize}px`;
        case "sysDashDotDot":
            return `${baseSize * 3}px ${baseSize}px ${baseSize}px ${baseSize}px ${baseSize}px ${baseSize}px`;
        default:
            return "";
    }
}

/**
 * Generate HTML/CSS arrow marker for line endings
 */
function generateArrowMarker(endType, strokeColor, strokeWidth, x, y, rotation = 0) {
    if (!endType || endType === "none") return "";
    
    const markerSize = Math.max(10, strokeWidth * 4);
    
    switch (endType) {
        case "triangle":
        case "arrow":
        case "stealth":
            // CSS triangle using borders
            return `<div style="
                position: absolute;
                left: ${x}px;
                top: ${y}px;
                width: 0;
                height: 0;
                border-left: ${markerSize}px solid ${strokeColor};
                border-top: ${markerSize / 2}px solid transparent;
                border-bottom: ${markerSize / 2}px solid transparent;
                transform: translate(-50%, -50%) rotate(${rotation}deg);
                transform-origin: center;
            "></div>`;
        
        case "oval":
        case "circle":
        case "dot":
            const radius = markerSize / 2.5;
            return `<div style="
                position: absolute;
                left: ${x}px;
                top: ${y}px;
                width: ${radius * 2}px;
                height: ${radius * 2}px;
                border: ${Math.max(1, strokeWidth)}px solid ${strokeColor};
                border-radius: 50%;
                background: ${strokeColor};
                transform: translate(-50%, -50%);
            "></div>`;
        
        case "diamond":
            return `<div style="
                position: absolute;
                left: ${x}px;
                top: ${y}px;
                width: ${markerSize * 0.7}px;
                height: ${markerSize * 0.7}px;
                background: ${strokeColor};
                transform: translate(-50%, -50%) rotate(45deg);
            "></div>`;
        
        default:
            return "";
    }
}

/**
 * Extract line end properties from XML
 */
function extractLineEnds(lineNode) {
    const headEnd = lineNode?.["a:headEnd"]?.[0]?.["$"];
    const tailEnd = lineNode?.["a:tailEnd"]?.[0]?.["$"];
    
    return {
        headType: headEnd?.type || "none",
        headSize: headEnd?.w || "med",
        tailType: tailEnd?.type || "none",
        tailSize: tailEnd?.w || "med"
    };
}

/**
 * Calculate path endpoints and segments for different connector types
 */
function calculateConnectorPath(shapeType, width, height, adj1Pct, adj2Pct = 0.5) {
    const segments = [];
    let startPoint = { x: 0, y: 0 };
    let endPoint = { x: width, y: height };
    
    switch (shapeType) {
        case "bentConnector2":
            segments.push({ 
                type: 'horizontal', 
                x1: 0, y1: 0, 
                x2: width, y2: 0 
            });
            segments.push({ 
                type: 'vertical', 
                x1: width, y1: 0, 
                x2: width, y2: height 
            });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
            
        case "bentConnector3":
            const bendX = width * adj1Pct;
            segments.push({ 
                type: 'horizontal', 
                x1: 0, y1: 0, 
                x2: bendX, y2: 0 
            });
            segments.push({ 
                type: 'vertical', 
                x1: bendX, y1: 0, 
                x2: bendX, y2: height 
            });
            segments.push({ 
                type: 'horizontal', 
                x1: bendX, y1: height, 
                x2: width, y2: height 
            });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
            
        case "bentConnector4":
            const b4_bend1X = width * adj1Pct;
            const b4_midY = height * adj2Pct;
            
            // 5 segments creating 4 bends: H → V → H → V → H
            segments.push({ 
                type: 'horizontal', 
                x1: 0, y1: 0, 
                x2: b4_bend1X, y2: 0 
            });
            segments.push({ 
                type: 'vertical', 
                x1: b4_bend1X, y1: 0, 
                x2: b4_bend1X, y2: b4_midY 
            });
            segments.push({ 
                type: 'horizontal', 
                x1: b4_bend1X, y1: b4_midY, 
                x2: width, y2: b4_midY 
            });
            segments.push({ 
                type: 'vertical', 
                x1: width, y1: b4_midY, 
                x2: width, y2: height 
            });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
            
        case "bentConnector5":
            segments.push({ 
                type: 'horizontal', 
                x1: 0, y1: 0, 
                x2: width * 0.25, y2: 0 
            });
            segments.push({ 
                type: 'vertical', 
                x1: width * 0.25, y1: 0, 
                x2: width * 0.25, y2: height * 0.33 
            });
            segments.push({ 
                type: 'horizontal', 
                x1: width * 0.25, y1: height * 0.33, 
                x2: width * 0.5, y2: height * 0.33 
            });
            segments.push({ 
                type: 'vertical', 
                x1: width * 0.5, y1: height * 0.33, 
                x2: width * 0.5, y2: height * 0.67 
            });
            segments.push({ 
                type: 'horizontal', 
                x1: width * 0.5, y1: height * 0.67, 
                x2: width * 0.75, y2: height * 0.67 
            });
            segments.push({ 
                type: 'vertical', 
                x1: width * 0.75, y1: height * 0.67, 
                x2: width * 0.75, y2: height 
            });
            segments.push({ 
                type: 'horizontal', 
                x1: width * 0.75, y1: height, 
                x2: width, y2: height 
            });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
            
        default:
            segments.push({ 
                type: 'diagonal', 
                x1: 0, y1: 0, 
                x2: width, y2: height 
            });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
    }
    
    return { segments, startPoint, endPoint };
}

/**
 * Calculate arrow rotation angle based on segment direction
 */
function getArrowRotation(segment, isStart) {
    // Calculate the direction angle of the segment
    const dx = segment.x2 - segment.x1;
    const dy = segment.y2 - segment.y1;
    
    const segmentAngle = Math.atan2(dy, dx) * 180 / Math.PI;
    if (isStart) {
        return segmentAngle + 180;
    } else {
        return segmentAngle;
    }
}

async function convertConnectorToHTML(shapeNode, themeXML, clrMap, masterXML, layoutXML) {
    if (!shapeNode) {
        console.warn("convertConnectorToHTML: shapeNode is null or undefined");
        return "";
    }

    const shapeName =
        shapeNode?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name ||
        shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name ||
        "Connector";

    // z-index from node ordering
    let zIndex = 0;
    if (shapeName && this.nodes) {
        const matchedNode = this.nodes.find(node => node.name === shapeName);
        if (matchedNode) zIndex = matchedNode.id;
    }

    const position = getShapePosition(shapeNode);
    const shapeType = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["$"]?.prst || "unknown";
    const line = shapeNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0];

    console.log(`Connector ${shapeName}: type=${shapeType}, pos=`, position);

    // Extract line properties
    let strokeColor = "#000000";
    let strokeWidth = 2;
    let strokeDashArray = "";
    if (line) {
        console.log("pppppppl,mnjhjbnm,jlsadiusa98d79sad8sa8");
        // Extract stroke width
        if (line["$"]?.w) {
            strokeWidth = Math.max(1, parseInt(line["$"].w, 10) / getEMUDivisor());
        }

        // Extract stroke color
        const solidFill = line?.["a:solidFill"];
        if (solidFill && solidFill.length > 0) {
            console.log("line==============");
            if (solidFill[0]["a:srgbClr"]) {
                strokeColor = `#${solidFill[0]["a:srgbClr"][0]["$"].val}`;
            } else if (solidFill[0]["a:schemeClr"]) {
                strokeColor = resolveColor(solidFill[0]["a:schemeClr"][0]["$"].val, clrMap, themeXML);

                const lumMod = solidFill[0]["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                }
            }
        } else {
            console.log("line==============else");
            const lnRef = shapeNode?.["p:style"]?.[0]?.["a:lnRef"]?.[0];
            if (lnRef?.["a:schemeClr"]) {
                strokeColor = resolveColor(lnRef["a:schemeClr"][0]["$"].val, clrMap, themeXML);
                
                // Apply lumMod if present
                const lumMod = lnRef["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                }
            } else if (lnRef?.["a:srgbClr"]) {
                strokeColor = `#${lnRef["a:srgbClr"][0]["$"].val}`;
            }
        }
        console.log("strokeColor0====",strokeColor);
        const dashType = line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid";
        strokeDashArray = getDashArray(dashType, strokeWidth);
    }
    // If no color was found from line, check style reference (for connectors without line object or without solidFill)
    if (strokeColor === "#000000") {
        console.log("Checking style reference for color");
        const lnRef = shapeNode?.["p:style"]?.[0]?.["a:lnRef"]?.[0];
        if (lnRef?.["a:schemeClr"]) {
            strokeColor = resolveColor(lnRef["a:schemeClr"][0]["$"].val, clrMap, themeXML);
            
            // Apply lumMod if present
            const lumMod = lnRef["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            if (lumMod) {
                strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
            }
        } else if (lnRef?.["a:srgbClr"]) {
            strokeColor = `#${lnRef["a:srgbClr"][0]["$"].val}`;
        }
    }
    // Extract line end properties (arrows, dots, etc.)
    const lineEnds = line ? extractLineEnds(line) : { headType: "none", tailType: "none" };

    // Handle straight connectors (straightConnector1, line)

    if (shapeType === "straightConnector1" || shapeType === "line") {
        console.log("Processing straight connector");
        const shapeColors = shapeFillColor.getShapeFillColor(shapeNode, themeXML, masterXML);
        const lineOpacity = shapeColors.strokeOpacity || 1.0;
        const finalStrokeColor = shapeColors.strokeColor && shapeColors.strokeColor !== "transparent" && shapeColors.strokeColor !== "none" ? shapeColors.strokeColor : strokeColor;

        const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
        const flipH = xfrm?.["$"]?.flipH === "1";
        const flipV = xfrm?.["$"]?.flipV === "1";

        let deltaX = position.width;
        let deltaY = position.height;

        if (flipH) deltaX = -deltaX;
        if (flipV) deltaY = -deltaY;

        const angleRad = Math.atan2(deltaY, deltaX);
        const angleDeg = angleRad * (180 / Math.PI);
        const lineLength = Math.sqrt(deltaX * deltaX + deltaY * deltaY);

        let startX = position.x;
        let startY = position.y;

        if (flipH && !flipV) {
            startX = position.x + position.width;
            startY = position.y;
        } else if (!flipH && flipV) {
            startX = position.x;
            startY = position.y + position.height;
        } else if (flipH && flipV) {
            startX = position.x + position.width;
            startY = position.y + position.height;
        }
        
        const isLineReversed = (flipH && !flipV) || (!flipH && flipV);
        
        let arrowTypeAtPosition0;     // Arrow at start of rendered line
        let arrowTypeAtPositionEnd;   // Arrow at end of rendered line
        
        if (isLineReversed) {
            arrowTypeAtPosition0 = lineEnds.tailType;
            arrowTypeAtPositionEnd = lineEnds.headType;
        } else {
            arrowTypeAtPosition0 = lineEnds.headType;
            arrowTypeAtPositionEnd = lineEnds.tailType;
        }
        
        // Generate markers
        const startMarker = generateArrowMarker(arrowTypeAtPosition0, finalStrokeColor, strokeWidth, 0, strokeWidth / 2, 180);
        const endMarker = generateArrowMarker(arrowTypeAtPositionEnd, finalStrokeColor, strokeWidth, lineLength, strokeWidth / 2, 0);

        // Store connector data with original PowerPoint types
        const straightConnectorData = JSON.stringify({
            shapeType: shapeType,
            strokeColor: finalStrokeColor,
            strokeWidth: strokeWidth,
            dashType: line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid",
            lineEnds: {
                headType: lineEnds.headType || "none",
                tailType: lineEnds.tailType || "none"
            },
            segments: [{
                type: 'diagonal',
                x1: startX,
                y1: startY,
                x2: startX + deltaX,
                y2: startY + deltaY
            }],
            position: {
                x: startX,
                y: startY,
                width: position.width,
                height: position.height
            },
            rotation: position.rotation || 0,
            flipH: flipH,
            flipV: flipV
        });

        return `<div class="shape line-connector"
            id="${shapeType}"
            data-shape-type="${shapeName}"
            data-name="${shapeName}"
            data-connector-info='${straightConnectorData.replace(/'/g, "&#39;")}'
            style="
                position: absolute;
                left: ${startX}px;
                top: ${startY}px;
                width: ${Math.abs(lineLength)}px;
                height: ${strokeWidth}px;
                transform: rotate(${angleDeg}deg);
                transform-origin: left center;
                overflow: visible;
                z-index: ${zIndex};
                pointer-events: auto;
            ">
            <div style="
                position: relative;
                width: 100%;
                height: 100%;
                background: ${finalStrokeColor};
                opacity: ${lineOpacity};
                ${strokeDashArray ? `
                    background: linear-gradient(to right, ${finalStrokeColor} 50%, transparent 50%);
                    background-size: ${strokeDashArray.split(' ')[0]} ${strokeWidth}px;
                ` : ''}
            ">
                ${startMarker}
                ${endMarker}
            </div>
        </div>`;
    }

    // Check if this is a curved connector
    const isCurvedConnector = shapeType.includes("curvedConnector");
    console.log("isCurvedConnector========",isCurvedConnector);
    if (isCurvedConnector) {
        // Handle curved connectors using smooth HTML curves
        const width = Math.max(1, position.width);
        const height = Math.max(1, position.height);

        const avLst = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["a:avLst"]?.[0];
        let adj1 = 50000;
        
        if (avLst?.["a:gd"]) {
            const gdList = Array.isArray(avLst["a:gd"]) ? avLst["a:gd"] : [avLst["a:gd"]];
            for (const gd of gdList) {
                const name = gd.$?.name;
                const fmla = gd.$?.fmla || "";
                if (name === "adj1" && fmla.includes("val")) {
                    adj1 = parseInt(fmla.replace("val ", ""), 10) || 50000;
                }
            }
        }

        const adj1Pct = adj1 / 100000;
        const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
        let rotation = xfrm?.["$"]?.rot ? parseInt(xfrm["$"].rot, 10) / 60000 : 0;
        const flipH = xfrm?.["$"]?.flipH === "1";
        const flipV = xfrm?.["$"]?.flipV === "1";
        
        // Generate smooth curved HTML
        // const curveData = generateCurvedHTML(shapeType, width, height, strokeColor, strokeWidth, adj1Pct);
        const curveData = generateCurvedHTML(shapeType, width, height, strokeColor, strokeWidth, adj1Pct, flipH, flipV);
        // ✅ FIX: Adjust rotation based on flip combinations to match PowerPoint behavior
        // PowerPoint applies rotation in shape space, then flips in parent space
        // CSS applies transforms left-to-right, so we need to compensate
        if (flipH && !flipV) {
            rotation = -rotation; // Horizontal flip only: negate rotation
        } else if (!flipH && flipV) {
            rotation = -rotation; // Vertical flip only: negate rotation
        }
        // If both flipH and flipV are true, rotation stays the same (180° flip = no rotation change)
    
        // Generate markers at curve endpoints with proper angles
        const tailMarker = generateArrowMarker(lineEnds.tailType, strokeColor, strokeWidth, 
            curveData.startPoint.x, curveData.startPoint.y, curveData.startAngle);
        const headMarker = generateArrowMarker(lineEnds.headType, strokeColor, strokeWidth, 
            curveData.endPoint.x, curveData.endPoint.y, curveData.endAngle);

        let transform = '';
        if (rotation !== 0) {
            transform += `rotate(${rotation}deg) `;
        }
        if (flipH) {
            transform += 'scaleX(-1) ';
        }
        if (flipV) {
            transform += 'scaleY(-1) ';
        }

        // ✅ Curved connectors don't store segments (generated dynamically)
        const connectorDataJSON = JSON.stringify({
            shapeType: shapeType,
            strokeColor: strokeColor,
            strokeWidth: strokeWidth,
            dashType: line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid",
            lineEnds: {
                headType: lineEnds.headType || "none",
                tailType: lineEnds.tailType || "none"
            },
            position: {
                x: position.x,
                y: position.y,
                width: width,
                height: height
            },
            rotation: rotation,
            flipH: flipH,
            flipV: flipV
        });

        console.log(`Curved connector ${shapeName}: rendered with smooth CSS`);

        return `<div class="shape connector curved-connector"
            data-shape-type="${shapeType}"
            id="${shapeType}"
            data-name="${shapeName}"
            data-connector-info='${connectorDataJSON.replace(/'/g, "&#39;")}'
            style="
                position: absolute;
                left: ${position.x}px;
                top: ${position.y}px;
                width: ${width}px;
                height: ${height}px;
                ${transform ? `transform: ${transform.trim()};` : ''}
                overflow: visible;
                z-index: ${zIndex};
                pointer-events: auto;
            ">
            <div style="position: relative; width: 100%; height: 100%;">
                ${curveData.html}
                ${tailMarker}
                ${headMarker}
            </div>
        </div>`;
    }
    
    // Handle bent connectors using HTML/CSS
    const width = Math.max(1, position.width);
    const height = Math.max(1, position.height);

    // Get adjustment values if present
    const avLst = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["a:avLst"]?.[0];
    let adj1 = 50000;
    let adj2 = 50000;
    if (avLst?.["a:gd"]) {
        const gdList = Array.isArray(avLst["a:gd"]) ? avLst["a:gd"] : [avLst["a:gd"]];
        for (const gd of gdList) {
            const name = gd.$?.name;
            const fmla = gd.$?.fmla || "";
            if (name === "adj1" && fmla.includes("val")) {
                adj1 = parseInt(fmla.replace("val ", ""), 10) || 50000;
            }
            else if (name === "adj2" && fmla.includes("val")) {
                adj2 = parseInt(fmla.replace("val ", ""), 10) || 50000;    
            }
        }
    }

    const adj1Pct = adj1 / 100000;
    const adj2Pct = adj2 / 100000;
    // Calculate path and get endpoints
    // const pathInfo = calculateConnectorPath(shapeType, width, height, adj1Pct);
    const pathInfo = calculateConnectorPath(shapeType, width, height, adj1Pct, adj2Pct);
    const { segments, startPoint, endPoint } = pathInfo;

    // Generate HTML segments
    let segmentsHTML = "";
    const borderStyle = strokeDashArray ? `${strokeWidth}px dashed ${strokeColor}` : `${strokeWidth}px solid ${strokeColor}`;

    for (const seg of segments) {
        if (seg.type === 'horizontal') {
            const left = Math.min(seg.x1, seg.x2);
            const segWidth = Math.abs(seg.x2 - seg.x1);
            segmentsHTML += `<div style="position: absolute; left: ${left}px; top: ${seg.y1}px; width: ${segWidth}px; border-top: ${borderStyle};"></div>`;
        } else if (seg.type === 'vertical') {
            const top = Math.min(seg.y1, seg.y2);
            const segHeight = Math.abs(seg.y2 - seg.y1);
            segmentsHTML += `<div style="position: absolute; left: ${seg.x1}px; top: ${top}px; height: ${segHeight}px; border-left: ${borderStyle};"></div>`;
        } else if (seg.type === 'diagonal') {
            const lineLength = Math.sqrt(Math.pow(seg.x2 - seg.x1, 2) + Math.pow(seg.y2 - seg.y1, 2));
            const angle = Math.atan2(seg.y2 - seg.y1, seg.x2 - seg.x1) * 180 / Math.PI;
            segmentsHTML += `<div style="
                position: absolute;
                left: ${seg.x1}px;
                top: ${seg.y1}px;
                width: ${lineLength}px;
                height: ${strokeWidth}px;
                background: ${strokeColor};
                transform-origin: left top;
                transform: rotate(${angle}deg);
                ${strokeDashArray ? `
                    background: linear-gradient(to right, ${strokeColor} 50%, transparent 50%);
                    background-size: ${strokeDashArray.split(' ')[0]} ${strokeWidth}px;
                ` : ''}
            "></div>`;
        }
    }

    // Generate markers at actual path endpoints
    const firstSegment = segments[0];
    const lastSegment = segments[segments.length - 1];
    
    const tailRotation = getArrowRotation(firstSegment, true);
    const headRotation = getArrowRotation(lastSegment, false);
    
    const tailMarker = generateArrowMarker(lineEnds.tailType, strokeColor, strokeWidth, startPoint.x, startPoint.y, tailRotation);
    const headMarker = generateArrowMarker(lineEnds.headType, strokeColor, strokeWidth, endPoint.x, endPoint.y, headRotation);

    // Check for rotation and flips
    const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
    let rotation = xfrm?.["$"]?.rot ? parseInt(xfrm["$"].rot, 10) / 60000 : 0;
    const flipH = xfrm?.["$"]?.flipH === "1";
    const flipV = xfrm?.["$"]?.flipV === "1";
    const originalRotation = rotation;

    // ✅ Build transform: rotation first, then swapped flips (no rotation compensation needed)
    let transform = '';
    if (rotation !== 0) {
        transform += `rotate(${rotation}deg) `;
    }
    if (flipV) {
        transform += 'scaleX(-1) ';
    }
    if (flipH) {
        transform += 'scaleY(-1) ';
    }

    // ✅ FIXED: Store ABSOLUTE positions by adding parent position to each segment
    const connectorDataJSON = JSON.stringify({
        shapeType: shapeType,
        strokeColor: strokeColor,
        strokeWidth: strokeWidth,
        dashType: line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid",
        lineEnds: {
            headType: lineEnds.headType || "none",
            tailType: lineEnds.tailType || "none"
        },
        // ✅ Convert relative segment positions to absolute positions
        segments: segments && segments.length > 0 ? segments.map(seg => ({
            type: seg.type,
            x1: position.x + seg.x1,  // Add parent X to make absolute
            y1: position.y + seg.y1,  // Add parent Y to make absolute
            x2: position.x + seg.x2,  // Add parent X to make absolute
            y2: position.y + seg.y2   // Add parent Y to make absolute
        })) : [],
        position: {
            x: position.x,
            y: position.y,
            width: width,
            height: height
        },
        rotation: rotation,
        originalRotation: originalRotation,
        flipH: flipH,
        flipV: flipV
    });
    
    return `<div class="shape connector"
        data-shape-type="${shapeType}"
        id="${shapeType}"
        data-name="${shapeName}"
        data-connector-info='${connectorDataJSON.replace(/'/g, "&#39;")}'
        style="
            position: absolute;
            left: ${position.x}px;
            top: ${position.y}px;
            width: ${width}px;
            height: ${height}px;
            ${transform ? `transform: ${transform.trim()};` : ''}
            overflow: visible;
            z-index: ${zIndex};
            pointer-events: auto;
        ">
        <div style="position: relative; width: 100%; height: 100%;">
            ${segmentsHTML}
            ${tailMarker}
            ${headMarker}
        </div>
    </div>`;
}

function getShapePosition(shapeNode, masterXML = null) {
    let xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];

    if (masterXML && (!xfrm || !xfrm?.["a:off"]?.[0] || !xfrm?.["a:ext"]?.[0])) {
        const placeholderType = getPlaceholderType(shapeNode);
        const masterPosition = getPositionFromMaster(masterXML, shapeNode, placeholderType);

        if (masterPosition) {
            xfrm = {
                "$": { rot: masterPosition.rot },
                "a:off": [{ "$": { x: masterPosition.x, y: masterPosition.y } }],
                "a:ext": [{ "$": { cx: masterPosition.cx, cy: masterPosition.cy } }]
            };
        }
    }

    const flipH = xfrm?.["$"]?.flipH === "1" || xfrm?.["$"]?.flipH === true;
    const flipV = xfrm?.["$"]?.flipV === "1" || xfrm?.["$"]?.flipV === true;

    const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
    
    // Use the existing divisor from getEMUDivisor() - DON'T CHANGE IT
    const divisor = getEMUDivisor();
    
    const width = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cx || 100) / divisor);
    const height = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cy || 100) / divisor);

    return {
        x: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.x || 0) / divisor),
        y: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.y || 0) / divisor),
        width: width < 1 ? 1 : width,
        height: height < 1 ? 1 : height,
        rotation: xfrm?.["$"]?.rot ? parseInt(xfrm["$"].rot, 10) / 60000 : 0,
        flipH: flipH,
        flipV: flipV,
        isTextBox: isTextBox
    };
}
function resolveColor(colorKey, clrMap, themeXML) {
    const mappedKey = clrMap[colorKey] || colorKey;
    const colorNode = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0]?.[`a:${mappedKey}`]?.[0];

    if (colorNode?.["a:sysClr"]) {
        return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;
    }

    return colorNode?.["a:srgbClr"] ? `#${colorNode["a:srgbClr"][0]["$"].val}` : "#000000";
}

function getPlaceholderType(shapeNode) {
    const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
    return phType || null;
}

function getPositionFromMaster(masterXML, shapeNode, placeholderType) {
    try {
        const spTree = masterXML?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0];
        if (!spTree || !placeholderType) return null;

        const placeholderIdx = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.idx;
        const placeholderIdxNum = placeholderIdx ? parseInt(placeholderIdx) : null;
        const shapes = spTree["p:sp"] || [];

        for (const shape of shapes) {
            const masterPh = shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];
            const masterPhType = masterPh?.["$"]?.type;
            const masterPhIdx = masterPh?.["$"]?.idx;
            const masterPhIdxNum = masterPhIdx ? parseInt(masterPhIdx) : null;

            const typeMatches = masterPhType === placeholderType;
            const indexMatches = (placeholderIdxNum === null && masterPhIdxNum === null) ||
                (placeholderIdxNum === masterPhIdxNum);

            if (typeMatches && indexMatches) {
                const xfrm = shape?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
                if (xfrm) {
                    return {
                        x: xfrm?.["a:off"]?.[0]?.["$"]?.x,
                        y: xfrm?.["a:off"]?.[0]?.["$"]?.y,
                        cx: xfrm?.["a:ext"]?.[0]?.["$"]?.cx,
                        cy: xfrm?.["a:ext"]?.[0]?.["$"]?.cy,
                        rot: xfrm?.["$"]?.rot
                    };
                }
            }
        }
        return null;
    } catch (error) {
        console.warn("Error extracting position from master:", error);
        return null;
    }
}

function generateCurvedHTML(shapeType, width, height, strokeColor, strokeWidth, adj1Pct = 0.5, flipH = false, flipV = false) {
    let curveHTML = '';
    let startPoint = { x: 0, y: 0 };
    let endPoint = { x: width, y: height };
    let startAngle = 0;
    let endAngle = 0;
    
    // High-resolution curve rendering
    const numPoints = 150; // More points for smoother curves
    const points = [];
    
    // PowerPoint's curved connectors use specific Bezier curve formulas
    switch (shapeType) {
        case "curvedConnector2": {
            // Quadratic Bezier: control point at (width, 0)
            const controlX = width;
            const controlY = 0;
            
            for (let i = 0; i <= numPoints; i++) {
                const t = i / numPoints;
                // Quadratic Bezier formula: B(t) = (1-t)²P0 + 2(1-t)tP1 + t²P2
                const x = (1 - t) * (1 - t) * 0 + 2 * (1 - t) * t * controlX + t * t * width;
                const y = (1 - t) * (1 - t) * 0 + 2 * (1 - t) * t * controlY + t * t * height;
                points.push({ x, y });
            }
            
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
        }
            
        case "curvedConnector3": {
            // Cubic Bezier S-curve
            const cp1x = width * adj1Pct;
            const cp1y = 0;
            const cp2x = width * adj1Pct;
            const cp2y = height;
            
            for (let i = 0; i <= numPoints; i++) {
                const t = i / numPoints;
                // Cubic Bezier formula: B(t) = (1-t)³P0 + 3(1-t)²tP1 + 3(1-t)t²P2 + t³P3
                const x = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1x + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2x + 
                         Math.pow(t, 3) * width;
                const y = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1y + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2y + 
                         Math.pow(t, 3) * height;
                points.push({ x, y });
            }
            
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
        }
            
        case "curvedConnector4": {
            // Cubic Bezier with adjusted control points
            const cp1x = width * adj1Pct;
            const cp1y = 0;
            const cp2x = width * (1 - adj1Pct);
            const cp2y = height;
            
            for (let i = 0; i <= numPoints; i++) {
                const t = i / numPoints;
                const x = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1x + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2x + 
                         Math.pow(t, 3) * width;
                const y = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1y + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2y + 
                         Math.pow(t, 3) * height;
                points.push({ x, y });
            }
            
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
        }
            
        case "curvedConnector5": {
            // Complex multi-segment Bezier curve
            const halfPoints = Math.floor(numPoints / 2);
            
            // First segment: (0,0) to (width/2, height/2)
            const cp1x = width * 0.25;
            const cp1y = 0;
            const cp2x = width * 0.5;
            const cp2y = height * 0.25;
            const midX = width * 0.5;
            const midY = height * 0.5;
            
            for (let i = 0; i <= halfPoints; i++) {
                const t = i / halfPoints;
                const x = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1x + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2x + 
                         Math.pow(t, 3) * midX;
                const y = Math.pow(1 - t, 3) * 0 + 
                         3 * Math.pow(1 - t, 2) * t * cp1y + 
                         3 * (1 - t) * Math.pow(t, 2) * cp2y + 
                         Math.pow(t, 3) * midY;
                points.push({ x, y });
            }
            
            // Second segment: (width/2, height/2) to (width, height)
            const cp3x = width * 0.5;
            const cp3y = height * 0.75;
            const cp4x = width * 0.75;
            const cp4y = height;
            
            for (let i = 1; i <= halfPoints; i++) {
                const t = i / halfPoints;
                const x = Math.pow(1 - t, 3) * midX + 
                         3 * Math.pow(1 - t, 2) * t * cp3x + 
                         3 * (1 - t) * Math.pow(t, 2) * cp4x + 
                         Math.pow(t, 3) * width;
                const y = Math.pow(1 - t, 3) * midY + 
                         3 * Math.pow(1 - t, 2) * t * cp3y + 
                         3 * (1 - t) * Math.pow(t, 2) * cp4y + 
                         Math.pow(t, 3) * height;
                points.push({ x, y });
            }
            
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
        }
            
        default: {
            // Fallback to straight line
            points.push({ x: 0, y: 0 });
            points.push({ x: width, y: height });
            startPoint = { x: 0, y: 0 };
            endPoint = { x: width, y: height };
            break;
        }
    }
    
    // Render curve using overlapping line segments
    for (let i = 0; i < points.length - 1; i++) {
        const p1 = points[i];
        const p2 = points[i + 1];
        
        const dx = p2.x - p1.x;
        const dy = p2.y - p1.y;
        const length = Math.sqrt(dx * dx + dy * dy);
        const angle = Math.atan2(dy, dx) * 180 / Math.PI;
        
        // Create smooth segments with slight overlap
        curveHTML += `<div style="
            position: absolute;
            left: ${p1.x}px;
            top: ${p1.y - strokeWidth / 2}px;
            width: ${length + 0.5}px;
            height: ${strokeWidth}px;
            background: ${strokeColor};
            transform: rotate(${angle}deg);
            transform-origin: 0 center;
            border-radius: ${strokeWidth / 4}px;
        "></div>`;
    }
    
    // Calculate start and end tangent angles from first/last few points
    if (points.length > 5) {
        const startDx = points[5].x - points[0].x;
        const startDy = points[5].y - points[0].y;
        startAngle = Math.atan2(startDy, startDx) * 180 / Math.PI + 180;
        
        const endDx = points[points.length - 1].x - points[points.length - 6].x;
        const endDy = points[points.length - 1].y - points[points.length - 6].y;
        endAngle = Math.atan2(endDy, endDx) * 180 / Math.PI;
    }
    
    return {
        html: curveHTML,
        startPoint: startPoint,
        endPoint: endPoint,
        startAngle: startAngle,
        endAngle: endAngle
    };
}






module.exports = {
    convertConnectorToHTML
};