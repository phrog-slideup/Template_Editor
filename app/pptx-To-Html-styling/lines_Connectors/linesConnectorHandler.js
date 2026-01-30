const colorHelper = require("../../api/helper/colorHelper.js");
const shapeFillColor = require("../shapes_Properties/getShapeFillColor.js");


// Add helper method to get the correct divisor
function getEMUDivisor() {
    // return parseInt(this.flag) === 1 ? 9525 : 12700;
    return 12700;
}

async function convertConnectorToHTML(shapeNode, themeXML, clrMap, masterXML, layoutXML) {
    if (!shapeNode) return "";

    const shapeName =
        shapeNode?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name ||
        shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;

    // z-index from node ordering
    let zIndex = 0;
    if (shapeName && this.nodes) {
        const matchedNode = this.nodes.find(node => node.name === shapeName);
        if (matchedNode) zIndex = matchedNode.id;
    }

    const position = getShapePosition(shapeNode);
    const shapeType =
        shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["$"]?.prst || "unknown";
    const line = shapeNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0];

    // Extract line properties
    let strokeColor = "#000000";
    let strokeWidth = 1;
    let strokeDashArray = "";

    if (line) {
        if (line["$"]?.w) {
            strokeWidth = parseInt(line["$"].w, 10) / getEMUDivisor();
        }

        const solidFill = line?.["a:solidFill"];
        if (solidFill && solidFill.length > 0) {
            if (solidFill[0]["a:srgbClr"]) {
                strokeColor = `#${solidFill[0]["a:srgbClr"][0]["$"].val}`;
            } else if (solidFill[0]["a:schemeClr"]) {
                strokeColor = resolveColor(solidFill[0]["a:schemeClr"][0]["$"].val, clrMap, themeXML);

                const lumMod =
                    solidFill[0]["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                }
            }
        } else {
            const styleColor =
                shapeNode?.["p:style"]?.[0]?.["a:lnRef"]?.[0]?.["a:schemeClr"]?.[0]?.[
                    "$"
                ]?.val;
            if (styleColor) strokeColor = resolveColor(styleColor, clrMap, themeXML);
        }

        const dashType = line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid";
        if (dashType === "dash") strokeDashArray = "5,5";
        if (dashType === "dot") strokeDashArray = "2,5";
        if (dashType === "dashDot") strokeDashArray = "5,5,2,5";
    }

    if (shapeType === "straightConnector1" || shapeType === "line") {
        const shapeColors = shapeFillColor.getShapeFillColor(shapeNode, themeXML, masterXML);

        const lineOpacity = shapeColors.strokeOpacity || 1.0;
        const finalStrokeColor = shapeColors.strokeColor || strokeColor || "#000000";

        // Get flip flags from xfrm
        const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
        const flipH = xfrm?.["$"]?.flipH === "1";
        const flipV = xfrm?.["$"]?.flipV === "1";

        // vector with flip applied
        let deltaX = position.width;
        let deltaY = position.height;

        if (flipH) deltaX = -deltaX;
        if (flipV) deltaY = -deltaY;

        const angleRad = Math.atan2(deltaY, deltaX);
        const angleDeg = angleRad * (180 / Math.PI);
        const lineLength = Math.sqrt(deltaX * deltaX + deltaY * deltaY);

        // ⭐ FIXED POSITION LOGIC FOR FLIPPED LINES
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

        return `<div class="shape line-lineheight"
            id="line-lineheight"
            data-shape-type="line"
            data-name="${shapeName}"
            style="
                position: absolute;
                left: ${startX}px;
                top: ${startY}px;
                width: ${Math.abs(lineLength)}px;
                height: ${strokeWidth}px;
                background: ${finalStrokeColor};
                transform: rotate(${angleDeg}deg);
                transform-origin: left center;
                opacity: ${lineOpacity};
                border-radius: ${Math.ceil(strokeWidth / 2)}px;
                z-index: ${zIndex};
                cursor: pointer;
            ">
        </div>`;
    }

    let pathD = "";
    const width = position.width;
    const height = position.height;

    switch (shapeType) {
        case "bentConnector2":
        case "bentConnector3":
            pathD = `M 0 0 H ${width / 2} V ${height} H ${width}`;
            break;
        case "bentConnector4":
        case "bentConnector5":
            pathD = `M 5,5 H ${width / 2} V ${height - 5} H ${width - 5}`;
            break;
        case "curvedConnector2":
            pathD = `M 0 0 Q ${width / 2} ${height / 2}, ${width} ${height}`;
            break;
        case "curvedConnector3":
        case "curvedConnector4":
        case "curvedConnector5":
            pathD = `M 5,5 Q ${width / 2},${height / 2} ${width - 5}, ${height - 5
                }`;
            break;
        default:
            pathD = `M 5,5 L ${width - 5},${height - 5}`;
            break;
    }

    const svgContent = `
        <svg width="${width}px" height="${height}px"
             xmlns="http://www.w3.org/2000/svg"
             style="overflow: visible;">
            <path d="${pathD}"
                  stroke="${strokeColor}"
                  stroke-width="${strokeWidth}px"
                  fill="none"
                  ${strokeDashArray ? `stroke-dasharray="${strokeDashArray}"` : ""}
            />
        </svg>`;

    return `<div class="shape"
        id="${shapeType}"
        data-name="${shapeName}"
        style="
            position: absolute;
            left: ${position.x}px;
            top: ${position.y}px;
            width: ${width}px;
            height: ${height}px;
            transform: rotate(${position.rotation}deg);
            overflow: visible;
            z-index: ${zIndex};
        ">
        ${svgContent.trim()}
    </div>`;
}


function getShapePosition(shapeNode, masterXML = null) {

    let xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];

    // NEW: If xfrm is missing or incomplete, try to get from master XML
    if (masterXML && (!xfrm || !xfrm?.["a:off"]?.[0] || !xfrm?.["a:ext"]?.[0])) {
        const placeholderType = getPlaceholderType(shapeNode);
        const masterPosition = getPositionFromMaster(masterXML, shapeNode, placeholderType);

        if (masterPosition) {
            // Create xfrm-like structure from master data to match original code
            xfrm = {
                "$": { rot: masterPosition.rot },
                "a:off": [{ "$": { x: masterPosition.x, y: masterPosition.y } }],
                "a:ext": [{ "$": { cx: masterPosition.cx, cy: masterPosition.cy } }]
            };
        }
    }

    const flipH = xfrm?.["$"]?.flipH === "1" || xfrm?.["$"]?.flipH === true;
    const flipV = xfrm?.["$"]?.flipV === "1" || xfrm?.["$"]?.flipV === true;

    // Check if it's a text box
    const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
    const width = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cx || 100) / getEMUDivisor());
    const height = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cy || 100) / getEMUDivisor());

    return {
        x: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.x || 0) / getEMUDivisor()),
        y: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.y || 0) / getEMUDivisor()),
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

// Helper: Get placeholder type from shape
function getPlaceholderType(shapeNode) {
    const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
    return phType || null;
}

// Helper: Get position and size from master XML by matching placeholder in <p:spTree>
function getPositionFromMaster(masterXML, shapeNode, placeholderType) {
    try {
        // Navigate to spTree which contains all shapes in the master
        const spTree = masterXML?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0];

        if (!spTree || !placeholderType) return null;

        // Get placeholder idx from the shape node
        const placeholderIdx = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.idx;
        const placeholderIdxNum = placeholderIdx ? parseInt(placeholderIdx) : null;

        // Get all <p:sp> shapes from master
        const shapes = spTree["p:sp"] || [];

        // Find the matching placeholder shape by type and index
        for (const shape of shapes) {
            // Get placeholder info from master shape
            const masterPh = shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];
            const masterPhType = masterPh?.["$"]?.type;
            const masterPhIdx = masterPh?.["$"]?.idx;
            const masterPhIdxNum = masterPhIdx ? parseInt(masterPhIdx) : null;

            // Match by type and index
            // If placeholderIdx is null (no idx attribute), match by type only
            const typeMatches = masterPhType === placeholderType;
            const indexMatches = (placeholderIdxNum === null && masterPhIdxNum === null) ||
                (placeholderIdxNum === masterPhIdxNum);

            if (typeMatches && indexMatches) {
                // Found matching placeholder! Extract position from <p:spPr><a:xfrm>
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


module.exports = {
    convertConnectorToHTML
};