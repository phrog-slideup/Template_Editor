const fs = require("fs");
const path = require("path");
const pptBackgroundColors = require("../../pptx-To-Html-styling/pptBackgroundColors.js");
const colorHelper = require("../../api/helper/colorHelper.js");

// Module-level cache for table styles (FIXED: was using 'this' in standalone function)
let cachedTableStyles = null;

// Add helper method to get the correct divisor
function getEMUDivisor() {
    return 12700;
}

/**
 * Convert a table XML node to its corresponding HTML representation
 */
async function convertTableXMLToHTML(tableNode, themeXML, extractor, nodes, flag, masterXML = null, layoutXML = null, clrMap = null) {
    if (!tableNode || typeof tableNode !== "object") {
        console.error("Invalid tableNode:", tableNode);
        return "";
    }

    // Import pptTextAllInfo to extract font size properly (like shapeHandler does)
    const pptTextAllInfo = require("../../pptx-To-Html-styling/pptTextAllInfo.js");

    // Extract the base font size for the table using pptTextAllInfo (same as shapeHandler)
    const tableFontSize = pptTextAllInfo.getFontSize(tableNode, flag);

    const styleConfig = await extractDynamicTableStyle(tableNode, themeXML, extractor)


    const tableName = tableNode["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
    const zIndex = nodes.find(node => node.name === tableName)?.id || 5;

    const xfrm = tableNode["p:xfrm"]?.[0];
    const x = parseInt(xfrm?.["a:off"]?.[0]?.["$"]?.x || 0, 10) / getEMUDivisor();
    const y = parseInt(xfrm?.["a:off"]?.[0]?.["$"]?.y || 0, 10) / getEMUDivisor();
    const width = parseInt(xfrm?.["a:ext"]?.[0]?.["$"]?.cx || 0, 10) / getEMUDivisor();
    const height = parseInt(xfrm?.["a:ext"]?.[0]?.["$"]?.cy || 0, 10) / getEMUDivisor(); 

    let tableHTML = `<div class="table-container" 
                            style="position: absolute; 
                            left: ${x}px; 
                            top: ${y}px; 
                            width: ${width}px; 
                            height: ${height}px;
                            z-index: ${zIndex};">

                                <table class="pptx-table" 
                                    style="width: 100%; 
                                    height: 100%; 
                                    border-collapse: collapse; 
                                    table-layout: fixed; 
                                    font-size: ${tableFontSize}px;">`;

    const tableData = tableNode["a:graphic"]?.[0]?.["a:graphicData"]?.[0]?.["a:tbl"]?.[0];
    if (!tableData) return "</table></div>";

    const gridCols = tableData["a:tblGrid"]?.[0]?.["a:gridCol"] || [];
    if (gridCols.length > 0) {
        tableHTML += "<colgroup>";
        const totalWidth = gridCols.reduce((sum, col) => sum + parseInt(col["$"]?.w || 0, 10), 0);
        for (const col of gridCols) {
            const colWidth = parseInt(col["$"]?.w || 0, 10);
            const widthPercent = totalWidth > 0 ? (colWidth / totalWidth) * 100 : (100 / gridCols.length);
            tableHTML += `<col style="width: ${widthPercent.toFixed(2)}%;">`;
        }
        tableHTML += "</colgroup>";
    }

    tableHTML += "<tbody>";
    const rows = tableData["a:tr"] || [];
    const totalRows = rows.length;

    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
        const row = rows[rowIndex];
        tableHTML += `<tr style="height: ${parseInt(row["$"]?.h || 0, 10) / getEMUDivisor()}px;">`;
        const cells = row["a:tc"] || [];
        const totalCols = cells.length;

  
let colIndex = 0;

for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
    const cell = cells[cellIndex];

    const gridSpan = parseInt(cell["$"]?.gridSpan || 1, 10);
    const rowSpan  = parseInt(cell["$"]?.rowSpan || 1, 10);

    const cellTcPr = cell["a:tcPr"]?.[0];

    const finalStyles = await getTableElementStyle(
        styleConfig,
        'cell',
        cellTcPr,
        rowIndex,
        colIndex,
        totalRows,
        gridCols.length,
        themeXML
    );

    const isHeaderCell = styleConfig.hasHeaderRow && rowIndex === 0;

    const cellContent = await extractCellContent(
        cell,
        isHeaderCell,
        tableNode,
        themeXML,
        tableFontSize,
        clrMap,
        masterXML
    );

    // Use middle alignment for header cells, top alignment for data cells
// ✅ Extract vertical alignment and text rotation from tcPr
const anchor = cellTcPr?.["$"]?.anchor || "t";  // t=top, ctr=center, b=bottom
const anchorCtr = cellTcPr?.["$"]?.anchorCtr === "1";
const vert = cellTcPr?.["$"]?.vert;  // vert270, vert90, etc.

// Map PowerPoint anchor to CSS vertical-align
let verticalAlign = 'top';
if (anchor === "ctr") verticalAlign = 'middle';
else if (anchor === "b") verticalAlign = 'bottom';

// ✅ Extract custom margins from tcPr
const marL = parseInt(cellTcPr?.["$"]?.marL || "91440") / getEMUDivisor();  // Default: 91440 EMU = ~7.2px
const marR = parseInt(cellTcPr?.["$"]?.marR || "91440") / getEMUDivisor();
const marT = parseInt(cellTcPr?.["$"]?.marT || "45720") / getEMUDivisor();  // Default: 45720 EMU = ~3.6px
const marB = parseInt(cellTcPr?.["$"]?.marB || "45720") / getEMUDivisor();


let cellStyles = `padding: ${marT}px ${marR}px ${marB}px ${marL}px; vertical-align: ${verticalAlign}; word-wrap: break-word; overflow: hidden;`;
// ✅ Handle vertical text rotation


// ✅ ADD THIS: Force cell content to respect vertical alignment
if (verticalAlign === 'bottom') {
    cellStyles += ` height: 100%; display: table-cell;`;
} else if (verticalAlign === 'middle') {
    cellStyles += ` height: 100%; display: table-cell;`;
}
// ✅ Handle vertical text rotation with proper constraints
// ✅ Handle vertical text rotation - allow text to use full cell width
// ✅ Handle vertical text rotation - use row height from XML
if (vert === "vert270") {
    const rowHeight = parseInt(row["$"]?.h || 0, 10) / getEMUDivisor();
    cellStyles += ` writing-mode: vertical-rl; transform: rotate(180deg); white-space: normal; width: 100%; height: ${rowHeight}px; box-sizing: border-box;`;
} else if (vert === "vert" || vert === "vert90") {
    const rowHeight = parseInt(row["$"]?.h || 0, 10) / getEMUDivisor();
    cellStyles += ` writing-mode: vertical-rl; white-space: normal; width: 100%; height: ${rowHeight}px; box-sizing: border-box;`;
}

// ✅ FIX ISSUE 2: Handle horizontal centering (anchorCtr) and text justification for header row
if (isHeaderCell) {
    // Row 1, Column 1: Bottom-right alignment is handled by anchor="b" (already set to bottom)
    // The right alignment comes from text-align in finalStyles
    // Column 2: Should use justify
    if (colIndex === 1) {
        // Override text-align for second column to use justify
        finalStyles.textAlign = 'justify';
    }
} else if (anchorCtr && !vert) {
    cellStyles += ` text-align: center;`;
}

    if (finalStyles.backgroundColor) cellStyles += ` background-color: ${finalStyles.backgroundColor};`;
    if (finalStyles.color) cellStyles += ` color: ${finalStyles.color};`;
    if (finalStyles.fontWeight !== 'normal') cellStyles += ` font-weight: ${finalStyles.fontWeight};`;
    if (finalStyles.fontSize) cellStyles += ` font-size: ${finalStyles.fontSize};`;

    if (finalStyles.textAlign) {
        cellStyles += ` text-align: ${finalStyles.textAlign};`;
    }

    if (finalStyles.borderTop && finalStyles.borderTop !== 'none') cellStyles += ` border-top: ${finalStyles.borderTop};`;
    if (finalStyles.borderBottom && finalStyles.borderBottom !== 'none') cellStyles += ` border-bottom: ${finalStyles.borderBottom};`;
    if (finalStyles.borderLeft && finalStyles.borderLeft !== 'none') cellStyles += ` border-left: ${finalStyles.borderLeft};`;
    if (finalStyles.borderRight && finalStyles.borderRight !== 'none') cellStyles += ` border-right: ${finalStyles.borderRight};`;

    const spanAttrs = `
        ${gridSpan > 1 ? `colspan="${gridSpan}"` : ""}
        ${rowSpan > 1 ? `rowspan="${rowSpan}"` : ""}
    `.trim();

    const cellTag = isHeaderCell ? "th" : "td";

    tableHTML += `<${cellTag} style="${cellStyles}" ${spanAttrs}>${cellContent}</${cellTag}>`;

    // Update column index for next iteration
    colIndex += gridSpan;
}
        tableHTML += `</tr>`;
    }

    tableHTML += "</tbody></table></div>";
    return tableHTML;
}

async function extractDynamicTableStyle(tableNode, themeXML, extractor) {
    const tableData = tableNode["a:graphic"]?.[0]?.["a:graphicData"]?.[0]?.["a:tbl"]?.[0];
    if (!tableData) return getEmptyTableStyle();

    // Extract table properties
    const tblPr = tableData["a:tblPr"]?.[0];
    const styleId = tblPr?.["a:tableStyleId"]?.[0];

    // Try to get style from tableStyles.xml first
    if (styleId) {
        const styleDefinition = await getTableStyleById(styleId, extractor);

        if (styleDefinition) {
            // Use tableStyles.xml
            return {
                styleDefinition,
                hasHeaderRow: tblPr?.["$"]?.firstRow === "1",
                hasBandedRows: tblPr?.["$"]?.bandRow === "1",
                hasBandedCols: tblPr?.["$"]?.bandCol === "1",
                hasFirstCol: tblPr?.["$"]?.firstCol === "1",
                hasLastCol: tblPr?.["$"]?.lastCol === "1",
                hasLastRow: tblPr?.["$"]?.lastRow === "1",
                styleId,
                source: 'tableStyles'
            };
        }
    }

    // Fallback: Extract styling directly from slide XML
    const slideBasedStyle = extractSlideTableStyle(tableData, tblPr, themeXML);

    return {
        ...slideBasedStyle,
        source: 'slideXML'
    };
}

function getEmptyTableStyle() {
    return {
        styleDefinition: {},
        hasHeaderRow: false,
        hasBandedRows: false,
        hasBandedCols: false,
        hasFirstCol: false,
        hasLastCol: false,
        hasLastRow: false,
        styleId: "empty"
    };
}

async function getTableStyleById(styleId, extractor) {
    const tableStyles = await loadTableStyles(extractor);

    if (!tableStyles || !styleId) {
        console.warn("No table styles loaded or no styleId provided:", {
            tableStyles: !!tableStyles,
            styleId: styleId
        });
        return null;
    }

    const styles = tableStyles["a:tblStyle"] || [];

    const foundStyle = styles.find(style => style["$"]?.styleId === styleId);
    if (!foundStyle) {
        console.warn(`Table style not found for ID: ${styleId}`);
        console.warn(`Available styles:`, styles.map(s => ({
            id: s["$"]?.styleId,
            name: s["$"]?.styleName
        })));
    }

    return foundStyle;
}

function extractSlideTableStyle(tableData, tblPr, themeXML) {
    const rows = tableData["a:tr"] || [];
    const hasHeaderRow = tblPr?.["$"]?.firstRow === "1";
    const hasBandedRows = tblPr?.["$"]?.bandRow === "1";

    // Analyze actual cell styling patterns
    const styleAnalysis = analyzeSlideTableCells(rows, hasHeaderRow, themeXML);

    // Build a dynamic style definition based on actual cell properties
    const dynamicStyleDefinition = buildDynamicStyleDefinition(styleAnalysis);

    return {
        styleDefinition: dynamicStyleDefinition,
        hasHeaderRow: hasHeaderRow,
        hasBandedRows: hasBandedRows,
        hasBandedCols: tblPr?.["$"]?.bandCol === "1",
        hasFirstCol: tblPr?.["$"]?.firstCol === "1",
        hasLastCol: tblPr?.["$"]?.lastCol === "1",
        hasLastRow: tblPr?.["$"]?.lastRow === "1",
        styleId: "dynamic"
    };
}

function analyzeSlideTableCells(rows, hasHeaderRow, themeXML) {
    if (!rows || rows.length === 0) return null;

    const analysis = {
        headerStyle: null,
        dataStyle: null,
        bandingPattern: null
    };

    // Analyze header row (first row)
    if (hasHeaderRow && rows.length > 0) {
        const headerCells = rows[0]["a:tc"] || [];
        if (headerCells.length > 0) {
            analysis.headerStyle = extractCellStyleFromSlide(headerCells[0], themeXML);
        }
    }

    // Analyze data rows for patterns
    const dataStartIndex = hasHeaderRow ? 1 : 0;
    if (rows.length > dataStartIndex) {
        // Get style from first data row
        const firstDataCells = rows[dataStartIndex]["a:tc"] || [];
        if (firstDataCells.length > 0) {
            analysis.dataStyle = extractCellStyleFromSlide(firstDataCells[0], themeXML);
        }

        // Check for alternating row colors (banding)
        if (rows.length > dataStartIndex + 1) {
            const secondDataCells = rows[dataStartIndex + 1]["a:tc"] || [];
            if (secondDataCells.length > 0) {
                const secondRowStyle = extractCellStyleFromSlide(secondDataCells[0], themeXML);
                if (secondRowStyle?.backgroundColor &&
                    analysis.dataStyle?.backgroundColor &&
                    secondRowStyle.backgroundColor !== analysis.dataStyle.backgroundColor) {
                    analysis.bandingPattern = {
                        evenRow: analysis.dataStyle.backgroundColor,
                        oddRow: secondRowStyle.backgroundColor
                    };
                }
            }
        }
    }

    return analysis;
}

function buildDynamicStyleDefinition(analysis) {
    if (!analysis) return {};

    const styleDefinition = {};

    // Build wholeTbl style
    if (analysis.dataStyle) {
        const firstDataStyle = analysis.dataStyle;
        styleDefinition["a:wholeTbl"] = [{
            "a:tcTxStyle": firstDataStyle.color ? [{
                "a:color": [{
                    "a:srgbClr": [{ "$": { "val": firstDataStyle.color.replace('#', '') } }]
                }]
            }] : undefined,
            "a:tcStyle": [{
                ...(firstDataStyle.backgroundColor ? {
                    "a:fill": [{
                        "a:solidFill": [{
                            "a:srgbClr": [{ "$": { "val": firstDataStyle.backgroundColor.replace('#', '') } }]
                        }]
                    }]
                } : {}),
                ...(firstDataStyle.borderTop || firstDataStyle.borderBottom || firstDataStyle.borderLeft || firstDataStyle.borderRight ? {
                    "a:tcBdr": [{
                        ...(firstDataStyle.borderTop ? { "a:top": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderTop)] }] } : {}),
                        ...(firstDataStyle.borderBottom ? { "a:bottom": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderBottom)] }] } : {}),
                        ...(firstDataStyle.borderLeft ? { "a:left": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderLeft)] }] } : {}),
                        ...(firstDataStyle.borderRight ? { "a:right": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderRight)] }] } : {})
                    }]
                } : {})
            }]
        }];
    }

    // Build header row style
    if (analysis.headerStyle) {
        const headerStyle = analysis.headerStyle;
        styleDefinition["a:firstRow"] = [{
            "a:tcTxStyle": [{
                ...(headerStyle.fontWeight === 'bold' ? { "b": "1" } : {}),
                ...(headerStyle.color ? {
                    "a:color": [{
                        "a:srgbClr": [{ "$": { "val": headerStyle.color.replace('#', '') } }]
                    }]
                } : {})
            }],
            "a:tcStyle": [{
                ...(headerStyle.backgroundColor ? {
                    "a:fill": [{
                        "a:solidFill": [{
                            "a:srgbClr": [{ "$": { "val": headerStyle.backgroundColor.replace('#', '') } }]
                        }]
                    }]
                } : {}),
                ...(headerStyle.borderTop || headerStyle.borderBottom || headerStyle.borderLeft || headerStyle.borderRight ? {
                    "a:tcBdr": [{
                        ...(headerStyle.borderTop ? { "a:top": [{ "a:ln": [convertBorderToXMLFormat(headerStyle.borderTop)] }] } : {}),
                        ...(headerStyle.borderBottom ? { "a:bottom": [{ "a:ln": [convertBorderToXMLFormat(headerStyle.borderBottom)] }] } : {}),
                        ...(headerStyle.borderLeft ? { "a:left": [{ "a:ln": [convertBorderToXMLFormat(headerStyle.borderLeft)] }] } : {}),
                        ...(headerStyle.borderRight ? { "a:right": [{ "a:ln": [convertBorderToXMLFormat(headerStyle.borderRight)] }] } : {})
                    }]
                } : {})
            }]
        }];

        // If we also have first data row style, include it in the definition
        if (analysis.dataStyle) {
            const firstDataStyle = analysis.dataStyle;
            styleDefinition["a:firstRow"][0]["a:tcStyle"] = [{
                ...(firstDataStyle.backgroundColor ? {
                    "a:fill": [{
                        "a:solidFill": [{
                            "a:srgbClr": [{ "$": { "val": firstDataStyle.backgroundColor.replace('#', '') } }]
                        }]
                    }]
                } : {}),
                ...(firstDataStyle.borderTop || firstDataStyle.borderBottom || firstDataStyle.borderLeft || firstDataStyle.borderRight ? {
                    "a:tcBdr": [{
                        ...(firstDataStyle.borderTop ? { "a:top": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderTop)] }] } : {}),
                        ...(firstDataStyle.borderBottom ? { "a:bottom": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderBottom)] }] } : {}),
                        ...(firstDataStyle.borderLeft ? { "a:left": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderLeft)] }] } : {}),
                        ...(firstDataStyle.borderRight ? { "a:right": [{ "a:ln": [convertBorderToXMLFormat(firstDataStyle.borderRight)] }] } : {})
                    }]
                } : {})
            }];
        }
    }

    // ✅ FIX ISSUE 1: Build banding styles - band1H should NOT have a fill (use wholeTbl color)
   // In buildDynamicStyleDefinition function


    return styleDefinition;
}

function convertBorderToXMLFormat(borderString) {
    if (!borderString || borderString === "none") return { "a:noFill": [{}] };

    // Parse "1px solid #CCCCCC" format
    const parts = borderString.split(' ');
    const width = parseFloat(parts[0]) * getEMUDivisor(); // Convert px back to EMU
    const color = parts[2]?.replace('#', '') || '000000';

    return {
        "$": { "w": width.toString() },
        "a:solidFill": [{
            "a:srgbClr": [{ "$": { "val": color } }]
        }]
    };
}

function extractCellStyleFromSlide(cell, themeXML) {
    if (!cell) return null;

    const tcPr = cell["a:tcPr"]?.[0];
    if (!tcPr) return null;

    const style = {
        backgroundColor: null,
        color: null,
        fontWeight: 'normal',
        borderTop: null,
        borderBottom: null,
        borderLeft: null,
        borderRight: null
    };

// Extract background color
const solidFill = tcPr["a:solidFill"]?.[0];
if (solidFill?.["a:schemeClr"]) {
    const colorVal = solidFill["a:schemeClr"][0]["$"]?.val;
    if (colorVal) {
        style.backgroundColor = getThemeColorValue(colorVal, themeXML);

        // ✅ Apply tint/lumMod/lumOff modifiers
        const tint = solidFill["a:schemeClr"][0]["a:tint"]?.[0]?.["$"]?.val;
        const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
        
        if (tint) {
            style.backgroundColor = applyTintToColor(style.backgroundColor, parseInt(tint));
        } else if (lumMod && lumOff) {
            // Apply both luminance modifiers
            style.backgroundColor = pptBackgroundColors.applyLuminanceModifier ?
                pptBackgroundColors.applyLuminanceModifier(style.backgroundColor, lumMod, lumOff) : 
                style.backgroundColor;
        } else if (lumMod) {
            // Apply only lumMod
            style.backgroundColor = colorHelper.applyLumMod ?
                colorHelper.applyLumMod(style.backgroundColor, lumMod) : 
                style.backgroundColor;
        }
    }
    } else if (solidFill?.["a:srgbClr"]) {
        style.backgroundColor = `#${solidFill["a:srgbClr"][0]["$"].val}`;
    }

    // Extract text color from cell content
    const textRun = cell["a:txBody"]?.[0]?.["a:p"]?.[0]?.["a:r"]?.[0];
    if (textRun?.["a:rPr"]?.[0]?.["a:solidFill"]?.[0]) {
        const textFill = textRun["a:rPr"][0]["a:solidFill"][0];
        if (textFill["a:schemeClr"]) {
            const colorVal = textFill["a:schemeClr"][0]["$"]?.val;
            style.color = getThemeColorValue(colorVal, themeXML);
        } else if (textFill["a:srgbClr"]) {
            style.color = `#${textFill["a:srgbClr"][0]["$"].val}`;
        } else if (textFill["a:prstClr"]) {
            style.color = textFill["a:prstClr"][0]["$"]?.val === "white" ? "#FFFFFF" : "#000000";
        }
    }

    // FIXED: Extract borders from slide XML format (a:lnL, a:lnR, a:lnT, a:lnB)
    const borderMappings = {
        'a:lnL': 'borderLeft',    // Left border
        'a:lnR': 'borderRight',   // Right border  
        'a:lnT': 'borderTop',     // Top border
        'a:lnB': 'borderBottom'   // Bottom border
    };

    Object.entries(borderMappings).forEach(([xmlKey, styleKey]) => {
        const borderElement = tcPr[xmlKey]?.[0];
        if (borderElement) {
            style[styleKey] = parseSlideCellBorder(borderElement, themeXML);
        }
    });

    return style;
}

function parseSlideCellBorder(borderElement, themeXML) {
    if (!borderElement) return null;

    // Check for no fill
    if (borderElement["a:noFill"]) {
        return "none";
    }

    // Extract width (convert EMU to pixels)
    const width = borderElement["$"]?.w ?
        Math.max(1, parseInt(borderElement["$"].w) / getEMUDivisor()) : 1;

    let color = "#000000"; // Default

    // Extract color from solid fill
    const solidFill = borderElement["a:solidFill"]?.[0];
    if (solidFill) {
        if (solidFill["a:schemeClr"]) {
            const schemeVal = solidFill["a:schemeClr"][0]["$"]?.val;
            color = getThemeColorValue(schemeVal, themeXML) || color;

            // Apply luminance modifiers (lumMod and lumOff)
            const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;

            if (lumMod && lumOff) {
                color = pptBackgroundColors.applyLuminanceModifier ?
                    pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff) : color;
            } else if (lumMod) {
                color = colorHelper.applyLumMod ?
                    colorHelper.applyLumMod(color, lumMod) : color;
            }
        } else if (solidFill["a:srgbClr"]) {
            color = `#${solidFill["a:srgbClr"][0]["$"].val}`;
        } else if (solidFill["a:prstClr"]) {
            const presetVal = solidFill["a:prstClr"][0]["$"]?.val;
            color = presetVal === "white" ? "#FFFFFF" :
                presetVal === "black" ? "#000000" : color;
        }
    }

    // Extract line style
    const compound = borderElement["$"]?.cmpd || "sng";
    const lineStyle = compound === "dbl" ? "double" : "solid";

    // Handle dash patterns
    const dashType = borderElement["a:prstDash"]?.[0]?.["$"]?.val;
    const finalLineStyle = dashType && dashType !== "solid" ? "dashed" : lineStyle;

    return `${width}px ${finalLineStyle} ${color}`;
}

function resolveSystemColor(sysColorValue) {
    // These are Windows system colors - can be dynamic based on user's system theme
    const systemColors = {
        'windowText': '#000000',
        'window': '#FFFFFF',
        'captionText': '#000000',
        'activeCaption': '#0078D4',
        'inactiveCaption': '#CCCCCC',
        'menu': '#F0F0F0',
        'menuText': '#000000'
    };

    return systemColors[sysColorValue] || '#000000';
}

// FIXED: Using module-level cache instead of 'this'
async function loadTableStyles(extractor) {
    if (cachedTableStyles) {
        return cachedTableStyles; // Return cached styles
    }

    try {
        // Use the extractor to access the tableStyles.xml file
        const tableStylesPath = "ppt/tableStyles.xml";
        const tableStylesFile = extractor.files[tableStylesPath];

        if (!tableStylesFile) {
            console.warn("tableStyles.xml not found in the PPTX file");
            return null;
        }

        // Get the XML content and parse it
        const xmlContent = await tableStylesFile.async("text");
        const xml2js = require('xml2js');
        const parser = new xml2js.Parser({ explicitArray: true });
        const tableStylesXML = await parser.parseStringPromise(xmlContent);

        if (tableStylesXML?.["a:tblStyleLst"]) {
            cachedTableStyles = tableStylesXML["a:tblStyleLst"];
            return cachedTableStyles;
        } else {
            console.warn("No table style list found in tableStyles.xml");
            return null;
        }
    } catch (error) {
        console.warn("Could not load table styles, using defaults:", error.message);
    }

    return null;
}

async function getTableElementStyle(styleConfig, elementType, cellTcPr, rowIndex = 0, colIndex = 0, totalRows = 0, totalCols = 0, themeXML = null) {
    const { styleDefinition, hasHeaderRow, hasBandedRows, hasBandedCols, hasFirstCol, hasLastCol, hasLastRow } = styleConfig;
    const isHeaderRow = hasHeaderRow && rowIndex === 0;

    // ✅ FIX: If cell has direct tcPr styling, prioritize it completely
    if (cellTcPr) {
        const directStyles = parseTableElementStyle(null, cellTcPr, isHeaderRow, themeXML);
        
        // If direct styling has background color, use ONLY direct styling
        if (directStyles.backgroundColor) {
            // Add header alignment
            if (isHeaderRow && !directStyles.textAlign) {
                directStyles.textAlign = 'center';
            }
            return directStyles;
        }
    }

    // 1. Get the base style from wholeTbl
    const baseStyleElement = styleDefinition["a:wholeTbl"]?.[0];
    const finalStyles = parseTableElementStyle(baseStyleElement, null, isHeaderRow, themeXML);

    if (isHeaderRow && !finalStyles.textAlign) {
        finalStyles.textAlign = 'center';
    }

    // 2. Find the specific override style element
    let overrideStyleElement = null;
    const isLastRow = rowIndex === totalRows - 1;
    const isFirstCol = colIndex === 0;
    const isLastCol = colIndex === totalCols - 1;

    if (isHeaderRow) {
        overrideStyleElement = styleDefinition["a:firstRow"]?.[0];
    } else if (isLastRow && hasLastRow) {
        overrideStyleElement = styleDefinition["a:lastRow"]?.[0];
    } else if (isFirstCol && hasFirstCol) {
        overrideStyleElement = styleDefinition["a:firstCol"]?.[0];
    } else if (isLastCol && hasLastCol) {
        overrideStyleElement = styleDefinition["a:lastCol"]?.[0];
    } else if (hasBandedRows) {
    const dataRowIndex = hasHeaderRow ? rowIndex - 1 : rowIndex;
    // Per tableStyles.xml:
    // - band1H: tint="40000" (40% = LIGHTER)
    // - band2H: NO fill (inherits wholeTbl tint="20000" = DARKER)
    // Even data rows (0,2,4) should be LIGHTER → use band1H
    // Odd data rows (1,3,5) should be DARKER → don't override, use wholeTbl
    if (dataRowIndex % 2 === 0) {
        overrideStyleElement = styleDefinition["a:band1H"]?.[0];
    }
    } else if (hasBandedCols) {
        overrideStyleElement = (colIndex % 2 === 0) ? styleDefinition["a:band1V"]?.[0] : styleDefinition["a:band2V"]?.[0];
    }

    // 3. Merge override style
   // 3. Merge override style (only if it was set above)
if (overrideStyleElement) {
    const overrideStyles = parseTableElementStyle(overrideStyleElement, null, isHeaderRow, themeXML);
    
    Object.keys(overrideStyles).forEach(key => {
        if (overrideStyles[key] !== null) {
            finalStyles[key] = overrideStyles[key];
        }
    });
}


    // 4. Merge direct cell formatting WITHOUT background if already processed
    if (cellTcPr) {
        const directStyles = parseTableElementStyle(null, cellTcPr, isHeaderRow, themeXML);
        Object.keys(directStyles).forEach(key => {
            if (directStyles[key] !== null) {
                finalStyles[key] = directStyles[key];
            }
        });
    }

    return finalStyles;
}

async function extractCellContent(cell, isHeader = false, shapeNode = null, themeXML = null, tableFontSize = 18, clrMap = null, masterXML = null) {
    if (!cell || !cell["a:txBody"]) {
        return "&nbsp;";
    }

    let content = [];
    const paragraphs = cell["a:txBody"]?.[0]?.["a:p"] || [];

    let currentListTag = null;
    let currentListType = null;

    for (const paragraph of paragraphs) {
        if (!paragraph) continue;

        // Check if this is an empty paragraph
        const isEmptyPara = isEmptyParagraph(paragraph);

        if (isEmptyPara) {
            // Close any open list for empty paragraphs
            if (currentListTag) {
                content.push(`</${currentListTag}>`);
                currentListTag = null;
                currentListType = null;
            }
            content.push('<br>');
            continue;
        }

        // Handle the specific pattern where one paragraph contains multiple bullet items
        const bulletItems = [];

        // Get arrays of each element type
        const pPrElements = paragraph["a:pPr"] || [];
        const rElements = paragraph["a:r"] || [];
        const brElements = paragraph["a:br"] || [];

        // If we have multiple pPr elements, each should be a separate bullet item
        if (pPrElements.length > 1) {
            // Multiple bullet items in one paragraph
            for (let i = 0; i < pPrElements.length; i++) {
                const pPr = pPrElements[i];
                const bulletInfo = extractBulletInformation(pPr, shapeNode);

                if (bulletInfo.hasListMarker) {
                    // Each pPr gets its corresponding run
                    const correspondingRun = rElements[i] || null;

                    bulletItems.push({
                        pPr: pPr,
                        runs: correspondingRun ? [correspondingRun] : [],
                        lineBreaks: [],
                        bulletInfo: bulletInfo
                    });
                }
            }
        } else {
            // Single bullet item or non-bullet paragraph (original logic)
            const pPrNode = pPrElements[0] || null;
            const bulletInfo = extractBulletInformation(pPrNode, shapeNode);

            bulletItems.push({
                pPr: pPrNode,
                runs: rElements,
                lineBreaks: brElements,
                bulletInfo: bulletInfo
            });
        }

        // If no bullet items were found, handle as a simple paragraph (fallback)
        if (bulletItems.length === 0) {
            const runs = paragraph["a:r"] || [];
            const lineBreaks = paragraph["a:br"] || [];
            const pPrNode = paragraph["a:pPr"]?.[0];

            bulletItems.push({
                pPr: pPrNode,
                runs: runs,
                lineBreaks: lineBreaks,
                bulletInfo: extractBulletInformation(pPrNode, shapeNode)
            });
        }

        // Process each bullet item
        bulletItems.forEach((item, itemIndex) => {
            const pPrNode = item.pPr;
            const runs = item.runs;
            const lineBreaks = item.lineBreaks;
            const bulletInfo = item.bulletInfo;

            const runTexts = [];

            // Process line breaks and runs
            if (lineBreaks.length > 0) {
                if (runs.length > 0) {
                    const firstRun = runs[0];
                    const textElement = firstRun?.["a:t"]?.[0];
                    const textValue = typeof textElement === 'string' ? textElement : "";

                    if (textValue !== undefined && textValue !== null) {
                        const runRPrNode = firstRun?.["a:rPr"]?.[0];
                        // Pass tableFontSize as fallback like shapeHandler does
                        const spanText = createCellSpanFromRun(runRPrNode, textValue, isHeader, themeXML, pPrNode, tableFontSize);
                        runTexts.push(spanText);
                    }
                }

                for (let i = 0; i < lineBreaks.length; i++) {
                    runTexts.push('<br>');
                }

                for (let i = 1; i < runs.length; i++) {
                    const run = runs[i];
                    const textElement = run?.["a:t"]?.[0];
                    const textValue = typeof textElement === 'string' ? textElement : "";

                    if (textValue !== undefined && textValue !== null) {
                        const runRPrNode = run?.["a:rPr"]?.[0];
                        // Pass tableFontSize as fallback like shapeHandler does
                        const spanText = createCellSpanFromRun(runRPrNode, textValue, isHeader, themeXML, pPrNode, tableFontSize);
                        runTexts.push(spanText);
                    }
                }
            } else {
                runs.forEach((run, runIndex) => {
                    const textElement = run?.["a:t"]?.[0];
                    const textValue = typeof textElement === 'string' ? textElement : " ";

                    if (textValue !== undefined && textValue !== null) {
                        const runRPrNode = run?.["a:rPr"]?.[0];
                        // Pass tableFontSize as fallback like shapeHandler does
                        const spanText = createCellSpanFromRun(runRPrNode, textValue, isHeader, themeXML, pPrNode, tableFontSize);
                        runTexts.push(spanText);
                    }
                });
            }

            // Handle fields (dates, slide numbers, etc.)
            const fields = paragraph["a:fld"] || [];
            for (const field of fields) {
                const fieldText = field?.["a:t"]?.[0];
                if (fieldText && typeof fieldText === 'string' && fieldText.trim()) {
                    const fieldStyles = extractRunStyles(field["a:rPr"]?.[0], themeXML, pPrNode, tableFontSize);
                    const styledField = fieldStyles ?
                        `<span style="${fieldStyles}">${escapeHtml(fieldText)}</span>` :
                        escapeHtml(fieldText);
                    runTexts.push(styledField);
                }
            }

            // Check if we have actual text content
            const hasActualText = runs.some(run => {
                const textElement = run?.["a:t"]?.[0];
                const textValue = typeof textElement === 'string' ? textElement : "";
                return textValue.length > 0;
            }) || fields.length > 0;

            const hasContent = runTexts.length > 0 && hasActualText;

            // Handle list formatting using existing functions
            if (bulletInfo.hasListMarker) {
            const listKey = `${bulletInfo.listTag}-${bulletInfo.listStyle}-${bulletInfo.bulletChar || bulletInfo.numberingType}`;

            if (currentListType !== listKey) {
                if (currentListTag) content.push(`</${currentListTag}>`);

                const listStyles = generateCellListStyles(bulletInfo);
                content.push(`<${bulletInfo.listTag} style="${listStyles}">`);
                currentListTag = bulletInfo.listTag;
                currentListType = listKey;
            }

            if (hasContent) {
               content.push(`<li style="margin: 0;">${runTexts.join('')}</li>`);
            }
        } else {
            // Close any open list for non-list items
            if (currentListTag) {
                content.push(`</${currentListTag}>`);
                currentListTag = null;
                currentListType = null;
            }

            if (hasContent) {
                // ✅ REMOVED text-align - let cell style handle alignment
                // ✅ Remove margin and let cell's vertical-align handle positioning
content.push(`<p style="margin: 0; line-height: 1.2;">${runTexts.join('')}</p>`);
            } else if (runTexts.length > 0 || lineBreaks.length > 0) {
                content.push('<br>');
            }
        }
        });
    }

    // Close any remaining open list
    if (currentListTag) {
        content.push(`</${currentListTag}>`);
    }

    // If no content found, return non-breaking space
    return content.length > 0 ? content.join('') : "&nbsp;";
}

function parseTableElementStyle(styleElement, directTcPr = null, isHeader = false, themeXML = null) {
    const styles = {
        backgroundColor: null,
        color: null,
        fontWeight: 'normal',
        fontSize: null,
        textAlign: null,
        borderTop: null,
        borderBottom: null,
        borderLeft: null,
        borderRight: null
    };

    // 1. Parse base styles from the table style definition (tableStyles.xml)
    if (styleElement) {
        const tcTxStyle = styleElement["a:tcTxStyle"]?.[0];
        if (tcTxStyle) {
            // Extract text color
            const colorNode = tcTxStyle["a:color"]?.[0];
            if (colorNode?.["a:srgbClr"]) {
                styles.color = `#${colorNode["a:srgbClr"][0]["$"].val}`;
            } else if (colorNode?.["a:schemeClr"]) {
                const schemeVal = colorNode["a:schemeClr"][0]["$"].val;
                styles.color = getThemeColorValue(schemeVal, themeXML);
            }

            // Extract font size
            const fontSize = tcTxStyle["$"]?.sz;
if (fontSize) {
    const sizeInPt = (parseInt(fontSize) / 100);
    styles.fontSize = `${sizeInPt}pt`;
}

            // Extract font weight
            if (tcTxStyle["$"]?.b === "on" || tcTxStyle["b"] || tcTxStyle["$"]?.b === "1") {
                styles.fontWeight = 'bold';
            }
        }

        // Extract background color and borders from tcStyle
        const tcStyle = styleElement["a:tcStyle"]?.[0];
        if (tcStyle) {
            // Background color
            const fillNode = tcStyle["a:fill"]?.[0]?.["a:solidFill"]?.[0];
            if (fillNode?.["a:srgbClr"]) {
                styles.backgroundColor = `#${fillNode["a:srgbClr"][0]["$"].val}`;
            } else if (fillNode?.["a:schemeClr"]) {
    const schemeClrNode = fillNode["a:schemeClr"][0];
    let schemeVal = schemeClrNode["$"].val;
    
    
    
    // Get base color from theme
    styles.backgroundColor = getThemeColorValue(schemeVal, themeXML);

    // Apply color modifiers in order
    const tint = schemeClrNode["a:tint"]?.[0]?.["$"]?.val;
    const shade = schemeClrNode["a:shade"]?.[0]?.["$"]?.val;
    const lumMod = schemeClrNode["a:lumMod"]?.[0]?.["$"]?.val;
    const lumOff = schemeClrNode["a:lumOff"]?.[0]?.["$"]?.val;
    
    if (styles.backgroundColor) {
        if (tint) {
            styles.backgroundColor = applyTintToColor(styles.backgroundColor, parseInt(tint));
        } else if (shade) {
            styles.backgroundColor = applyTintToColor(styles.backgroundColor, -parseInt(shade));
        } else if (lumMod || lumOff) {
            if (pptBackgroundColors && pptBackgroundColors.applyLuminanceModifier) {
                styles.backgroundColor = pptBackgroundColors.applyLuminanceModifier(
                    styles.backgroundColor, 
                    lumMod || "100000", 
                    lumOff || "0"
                );
            }
        }
    }
}

            // Extract borders
            const borders = tcStyle["a:tcBdr"]?.[0];
            if (borders) {
                const borderMappings = {
                    'a:left': 'borderLeft',
                    'a:right': 'borderRight',
                    'a:top': 'borderTop',
                    'a:bottom': 'borderBottom'
                };

                Object.entries(borderMappings).forEach(([xmlKey, styleKey]) => {
                    const borderElement = borders[xmlKey]?.[0]?.["a:ln"]?.[0];
                    if (borderElement) {
                        styles[styleKey] = parseBorderStyle(borderElement, themeXML);
                    }
                });
            }
        }
    }

    // 2. Merge direct cell properties (`tcPr`) if provided
    if (directTcPr) {
        // Background color
        const solidFill = directTcPr["a:solidFill"]?.[0];
        if (solidFill?.["a:schemeClr"]) {
            const colorVal = solidFill["a:schemeClr"][0]["$"]?.val;
            if (colorVal) {
                styles.backgroundColor = getThemeColorValue(colorVal, themeXML);

                const tint = solidFill["a:schemeClr"][0]["a:tint"]?.[0]?.["$"]?.val;
                if (tint) {
                    styles.backgroundColor = applyTintToColor(styles.backgroundColor, parseInt(tint));
                }
            }
        } else if (solidFill?.["a:srgbClr"]) {
            styles.backgroundColor = `#${solidFill["a:srgbClr"][0]["$"].val}`;
        }

        // Extract borders from directTcPr (slide XML format)
        const borderMappings = {
            'a:lnL': 'borderLeft',
            'a:lnR': 'borderRight',
            'a:lnT': 'borderTop',
            'a:lnB': 'borderBottom'
        };

        Object.entries(borderMappings).forEach(([xmlKey, styleKey]) => {
            const borderElement = directTcPr[xmlKey]?.[0];
            if (borderElement) {
                const parsedBorder = parseSlideCellBorder(borderElement, themeXML);
                if (parsedBorder) {
                    styles[styleKey] = parsedBorder;
                }
            }
        });
    }

    return styles;
}

function parseBorderStyle(borderElement, themeXML) {
    if (!borderElement) return null;

    // Check for no fill
    if (borderElement["a:noFill"]) {
        return "none";
    }

    // Extract width
    const width = borderElement["$"]?.w ?
        Math.max(1, parseInt(borderElement["$"].w) / getEMUDivisor()) : 1;

    let color = "#000000"; // Default

    // Extract color from solid fill
    const solidFill = borderElement["a:solidFill"]?.[0];
    if (solidFill) {
        if (solidFill["a:schemeClr"]) {
            const schemeVal = solidFill["a:schemeClr"][0]["$"]?.val;
            color = getThemeColorValue(schemeVal, themeXML) || color;

            // Apply luminance modifiers
            const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;

            if (lumMod && lumOff) {
                color = pptBackgroundColors.applyLuminanceModifier ?
                    pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff) : color;
            } else if (lumMod) {
                color = colorHelper.applyLumMod ?
                    colorHelper.applyLumMod(color, lumMod) : color;
            }
        } else if (solidFill["a:srgbClr"]) {
            color = `#${solidFill["a:srgbClr"][0]["$"].val}`;
        }
    }

    // Extract line style
    const compound = borderElement["$"]?.cmpd || "sng";
    const lineStyle = compound === "dbl" ? "double" : "solid";

    // Handle dash patterns
    const dashType = borderElement["a:prstDash"]?.[0]?.["$"]?.val;
    const finalLineStyle = dashType && dashType !== "solid" ? "dashed" : lineStyle;

    return `${width}px ${finalLineStyle} ${color}`;
}

function getThemeColorValue(schemeVal, themeXML) {
    if (!themeXML) return null;

    // ✅ ADD THIS MAPPING FIRST
    const schemeMap = {
        'bg1': 'lt1',   // Background 1 = Light 1
        'bg2': 'lt2',   // Background 2 = Light 2
        'tx1': 'dk1',   // Text 1 = Dark 1
        'tx2': 'dk2'    // Text 2 = Dark 2
    };
     schemeVal = schemeMap[schemeVal] || schemeVal;

    // Try to resolve via colorHelper first
    if (colorHelper && colorHelper.resolveThemeColorHelper) {
        const resolved = colorHelper.resolveThemeColorHelper(schemeVal, themeXML);
        if (resolved) return resolved;
    }

    // Fallback: Direct theme lookup
    const colorScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
    if (!colorScheme) return null;

    const colorNode = colorScheme[`a:${schemeVal}`]?.[0];
    if (!colorNode) return null;

    // Check for system color
    if (colorNode["a:sysClr"]) {
        return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;
    }

    // Check for RGB color
    if (colorNode["a:srgbClr"]) {
        return `#${colorNode["a:srgbClr"][0]["$"].val}`;
    }

    return null;
}

function isEmptyParagraph(paragraph) {
    if (!paragraph) return true;

    const runs = paragraph["a:r"] || [];
    const fields = paragraph["a:fld"] || [];
    const endParaRPr = paragraph["a:endParaRPr"];

    // If there are no runs, no fields, and only endParaRPr, it's likely empty
    if (runs.length === 0 && fields.length === 0 && endParaRPr) {
        return true;
    }

    // Check if all runs have empty or whitespace-only text
    const hasText = runs.some(run => {
        const textElement = run?.["a:t"]?.[0];
        const textValue = typeof textElement === 'string' ? textElement.trim() : "";
        return textValue.length > 0;
    });

    return !hasText && fields.length === 0;
}

function createCellSpanFromRun(rPr, textValue, isHeader, themeXML, pPr = null, tableFontSize = 18) {
    const styles = extractRunStyles(rPr, themeXML, pPr, tableFontSize);
    const escapedText = escapeHtml(textValue);

    if (!styles) {
        return escapedText;
    }

    return `<span style="${styles}">${escapedText}</span>`;
}

function extractRunStyles(rPr, themeXML, pPr = null, tableFontSize = 18) {
    const styles = [];
        // ✅ ADD THIS: Check paragraph properties FIRST for inherited text color
    let textColorFound = false;
    if (pPr) {
        const defRPr = pPr["a:defRPr"]?.[0];
        if (defRPr) {
            const defTextColor = extractTextColor(defRPr, themeXML);
            if (defTextColor) {
                styles.push(`color: ${defTextColor}`);
                textColorFound = true;
            }
        }
    }

    let fontSize = rPr?.["$"]?.sz;

    // If no font size in run properties, check paragraph default run properties
    if (!fontSize && pPr) {
        const defRPr = pPr["a:defRPr"]?.[0];
        fontSize = defRPr?.["$"]?.sz;
    }
console.log('Extracted font size:', fontSize);
   if (fontSize) {
    const sizeInPt = (parseInt(fontSize) / 100) * 0.75;  // Reduce by 25%
    styles.push(`font-size: ${sizeInPt}px`);
} else {
    styles.push(`font-size: 10px`);  // Smaller default
}

    

    if (!rPr) return styles.length > 0 ? styles.join('; ') : null;

    // Bold
    if (rPr["$"]?.b === "1" || rPr["$"]?.b === true) {
        styles.push('font-weight: bold');
    }

    // Italic
    if (rPr["$"]?.i === "1" || rPr["$"]?.i === true) {
        styles.push('font-style: italic');
    }

    // Underline
    if (rPr["$"]?.u && rPr["$"]?.u !== "none") {
        styles.push('text-decoration: underline');
    }

    // ✅ Text color from run properties (only if not already found from paragraph)
if (!textColorFound) {
    const textColor = extractTextColor(rPr, themeXML);
    if (textColor) {
        styles.push(`color: ${textColor}`);
    }
}

    return styles.length > 0 ? styles.join('; ') : null;
}

function extractBulletInformation(pPr, shapeNode) {
    if (!pPr) {
        return {
            hasListMarker: false,
            listTag: null,
            listStyle: null,
            bulletChar: null,
            bulletFont: null,
            numberingType: null,
            indentLevel: 0,
            leftMargin: 0,
            textIndent: 0
        };
    }

    const buChar = pPr["a:buChar"]?.[0];
    const buAutoNum = pPr["a:buAutoNum"]?.[0];
    const buNone = pPr["a:buNone"];
    const bulletFont = buChar?.["$"]?.typeface;

    // Extract indentation info
    const lvl = parseInt(pPr["$"]?.lvl || "0");
    const indent = parseInt(pPr["$"]?.indent || "0") / getEMUDivisor();
    const marL = parseInt(pPr["$"]?.marL || "0") / getEMUDivisor();

    // If buNone is present, there is no bullet
    if (buNone) {
        return {
            hasListMarker: false,
            listTag: null,
            listStyle: null,
            bulletChar: null,
            bulletFont: null,
            numberingType: null,
            indentLevel: lvl,
            leftMargin: marL,
            textIndent: indent
        };
    }

    // Ordered list (numbered)
    if (buAutoNum) {
        const autoNumType = buAutoNum["$"]?.type || "arabicPeriod";
        return {
            hasListMarker: true,
            listTag: 'ol',
            listStyle: convertNumberingType(autoNumType),
            bulletChar: null,
            bulletFont: bulletFont,
            numberingType: autoNumType,
            indentLevel: lvl,
            leftMargin: marL,
            textIndent: indent
        };
    }

    // Unordered list (bulleted)
    if (buChar) {
        const bulletCharVal = buChar["$"]?.char || "•";
        return {
            hasListMarker: true,
            listTag: 'ul',
            listStyle: getBulletStyleFromChar(bulletCharVal, bulletFont),
            bulletChar: bulletCharVal,
            bulletFont: bulletFont,
            numberingType: null,
            indentLevel: lvl,
            leftMargin: marL,
            textIndent: indent
        };
    }

    // No explicit bullet info found, but check if shape has default bullet style
    return {
        hasListMarker: false,
        listTag: null,
        listStyle: null,
        bulletChar: null,
        bulletFont: null,
        numberingType: null,
        indentLevel: lvl,
        leftMargin: marL,
        textIndent: indent
    };
}

function generateCellListStyles(bulletInfo) {
    const styles = [];

    // Add list style based on bullet type
    if (bulletInfo.listTag === 'ul') {
        const ulStyle = getCellUnorderedListStyle(bulletInfo);
        if (ulStyle) {
            styles.push(ulStyle);
        }
    } else if (bulletInfo.listTag === 'ol') {
        styles.push(`list-style-type: ${bulletInfo.listStyle}`);
    }

    // Add indentation
    if (bulletInfo.leftMargin > 0) {
        styles.push(`margin-left: ${bulletInfo.leftMargin}px`);
    }

    // Add padding for nested lists
    if (bulletInfo.indentLevel > 0) {
        styles.push(`padding-left: ${bulletInfo.indentLevel * 20}px`);
    }

    // Remove default margins
    styles.push('margin-top: 2px');
    styles.push('margin-bottom: 2px');

    return styles.join('; ');
}

function getBulletStyleFromChar(bulletChar, bulletFont) {
    // Map common bullet characters to CSS list styles
    const bulletMap = {
        '•': 'disc',           // Filled circle (most common)
        '○': 'circle',         // Hollow circle
        '▪': 'square',         // Filled square
        '▫': 'square',         // Hollow square
        '◆': 'diamond',        // Filled diamond
        '◇': 'diamond',        // Hollow diamond
        '►': 'triangle',       // Right-pointing triangle
        '▶': 'triangle',       // Right-pointing triangle (filled)
        '✓': 'checkmark',      // Checkmark
        '✔': 'checkmark',      // Heavy checkmark
        '✗': 'x-mark',         // X mark
        '✘': 'x-mark',         // Heavy X mark
        '➢': 'arrow',          // Right arrow
        '➤': 'arrow',          // Right arrow (filled)
        '→': 'arrow',          // Right arrow
        '★': 'star',           // Filled star
        '☆': 'star',           // Hollow star
        '❤': 'heart',          // Heart
        '♥': 'heart',          // Heart (filled)
        '☺': 'smiley',         // Smiley face
        '☻': 'smiley',         // Smiley face (filled)
        '-': 'disc',           // Hyphen/dash
        '–': 'disc',           // En dash
        '—': 'disc'            // Em dash
    };

    // Special handling for Wingdings and other symbol fonts
    if (bulletFont) {
        const fontLower = bulletFont.toLowerCase();
        if (fontLower.includes('wingdings') || fontLower.includes('symbol')) {
            // Map Wingdings characters
            const wingdingsMap = {
                'ü': 'checkmark',
                'û': 'x-mark',
                'l': 'square',
                'n': 'square',
                'Ø': 'circle',
                '': 'disc'
            };
            return wingdingsMap[bulletChar] || 'disc';
        }
    }

    return bulletMap[bulletChar] || 'disc';
}

function getTextAlignmentFromParagraph(pPr, shapeNode) {
    if (!pPr) return 'left';

    const algn = pPr["$"]?.algn;
    if (algn) {
        return convertAlgnToCSS(algn);
    }

    return 'left';
}

function escapeHtml(text) {
    if (typeof text !== 'string') return '';

    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    };

    return text.replace(/[&<>"']/g, m => map[m]);
}

function convertNumberingType(autoNumType) {
    const numberingStyles = {
        // Arabic numerals
        'arabicPeriod': 'decimal',
        'arabicParenR': 'decimal',
        'arabicParenBoth': 'decimal',           // (1) (2) (3)
        'arabicPlain': 'decimal',               // 1 2 3

        // Lowercase alphabetical
        'alphaLcPeriod': 'lower-alpha',         // a. b. c.
        'alphaLcParenR': 'lower-alpha',         // a) b) c)
        'alphaLcParenBoth': 'lower-alpha',      // (a) (b) (c)
        'alphaLcPlain': 'lower-alpha',          // a b c

        // Uppercase alphabetical
        'alphaUcPeriod': 'upper-alpha',         // A. B. C.
        'alphaUcParenR': 'upper-alpha',         // A) B) C)
        'alphaUcParenBoth': 'upper-alpha',      // (A) (B) (C)
        'alphaUcPlain': 'upper-alpha',          // A B C

        // Lowercase Roman numerals
        'romanLcPeriod': 'lower-roman',         // i. ii. iii.
        'romanLcParenR': 'lower-roman',         // i) ii) iii)
        'romanLcParenBoth': 'lower-roman',      // (i) (ii) (iii)
        'romanLcPlain': 'lower-roman',          // i ii iii

        // Uppercase Roman numerals
        'romanUcPeriod': 'upper-roman',         // I. II. III.
        'romanUcParenR': 'upper-roman',         // I) II) III)
        'romanUcParenBoth': 'upper-roman',      // (I) (II) (III)
        'romanUcPlain': 'upper-roman',          // I II III

        // Special numbering formats
        'circleNumDbPlain': 'decimal-leading-zero',    // 01 02 03
        'circleNumWdBlackPlain': 'decimal-leading-zero', // 01 02 03
        'circleNumWdWhitePlain': 'decimal',            // ? ? ?
        'arabicAbjadDash': 'decimal',                  // Arabic Abjad
        'arabicAlphaDash': 'decimal',                  // Arabic Alpha
        'hindiAlphaPeriod': 'decimal',                 // Hindi Alpha
        'hindiNumPeriod': 'decimal',                   // Hindi Numbers
        'hindiVowels': 'decimal',                      // Hindi Vowels
        'thaiAlphaPeriod': 'decimal',                  // Thai Alpha
        'thaiAlphaParenR': 'decimal',                  // Thai Alpha
        'thaiNumPeriod': 'decimal',                    // Thai Numbers
        'hebrewAlphaDash': 'decimal',                  // Hebrew Alpha
        'hebrew2Minus': 'decimal'                      // Hebrew Numbers
    };

    return numberingStyles[autoNumType] || 'decimal';
}

function extractTextColor(rPr, themeXML = null) {
    const solidFill = rPr?.["a:solidFill"]?.[0];
    if (!solidFill) return null;

    // RGB color (direct)
    if (solidFill["a:srgbClr"]) {
        return `#${solidFill["a:srgbClr"][0]["$"].val}`;
    }
    
    // Scheme color (theme-based)
    // Scheme color (theme-based)
if (solidFill["a:schemeClr"]) {
    const schemeVal = solidFill["a:schemeClr"][0]["$"].val;
    
    // ✅ ADD THIS ENTIRE BLOCK:
    const schemeMap = {
        'bg1': 'lt1',
        'bg2': 'lt2',
        'tx1': 'dk1',
        'tx2': 'dk2'
    };
    
    const mappedScheme = schemeMap[schemeVal] || schemeVal;
    
    let color = getThemeColorValue(mappedScheme, themeXML);
    
    // ✅ Apply lumMod if present
    const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
    if (lumMod && color) {
        color = colorHelper.applyLumMod ?
            colorHelper.applyLumMod(color, lumMod) : color;
    }
    
    return color;
}
    
    // Preset color
    if (solidFill["a:prstClr"]) {
        const presetVal = solidFill["a:prstClr"][0]["$"]?.val;
        const presetColors = {
            'white': '#FFFFFF',
            'black': '#000000'
        };
        return presetColors[presetVal] || null;
    }
    
    // System color
    if (solidFill["a:sysClr"]) {
        const lastClr = solidFill["a:sysClr"][0]["$"]?.lastClr;
        if (lastClr) return `#${lastClr}`;
    }
    
    return null;
}

function getCellUnorderedListStyle(bulletInfo) {
    let detectedStyle = bulletInfo.listStyle; // Default to the existing listStyle

    // If we have bullet character info, re-detect the style using existing function
    if (bulletInfo.bulletChar) {
        detectedStyle = getBulletStyleFromChar(bulletInfo.bulletChar, bulletInfo.bulletFont);
    }

    // Map the detected style to CSS
    const standardStyles = {
        'disc': 'list-style-type: disc',
        'circle': 'list-style-type: circle',
        'square': 'list-style-type: square'
    };

    // Return standard CSS if available
    if (standardStyles[detectedStyle]) {
        return standardStyles[detectedStyle];
    }

    // For custom bullets that CSS doesn't support, use SVG images
    const customStyles = {
        'triangle': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpolygon points=\'0,0 8,4 0,8\' fill=\'black\'/%3E%3C/svg%3E")',
        'diamond': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpolygon points=\'4,0 8,4 4,8 0,4\' fill=\'black\'/%3E%3C/svg%3E")',
        'checkmark': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M1,4 L3,6 L7,2\' stroke=\'black\' stroke-width=\'1\' fill=\'none\'/%3E%3C/svg%3E")',
        'x-mark': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M2,2 L6,6 M6,2 L2,6\' stroke=\'black\' stroke-width=\'1\' fill=\'none\'/%3E%3C/svg%3E")',
        'arrow': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M1,4 L6,4 M4,2 L6,4 L4,6\' stroke=\'black\' stroke-width=\'1\' fill=\'none\'/%3E%3C/svg%3E")',
        'star': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M4,0 L5,3 L8,3 L6,5 L7,8 L4,6 L1,8 L2,5 L0,3 L3,3 Z\' fill=\'black\'/%3E%3C/svg%3E")',
        'heart': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M4,7 C1,4 0,2 2,2 C3,2 4,3 4,3 C4,3 5,2 6,2 C8,2 7,4 4,7 Z\' fill=\'black\'/%3E%3C/svg%3E")',
        'smiley': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Ccircle cx=\'4\' cy=\'4\' r=\'3\' fill=\'none\' stroke=\'black\'/%3E%3Ccircle cx=\'3\' cy=\'3\' r=\'0.5\' fill=\'black\'/%3E%3Ccircle cx=\'5\' cy=\'3\' r=\'0.5\' fill=\'black\'/%3E%3Cpath d=\'M2.5,5 Q4,6.5 5.5,5\' stroke=\'black\' fill=\'none\'/%3E%3C/svg%3E")'
    };

    // Return custom style if available
    if (customStyles[detectedStyle]) {
        return customStyles[detectedStyle];
    }

    // Final fallback to disc
    return 'list-style-type: disc';
}

function getLevelProperties(lstStyle, level) {
    const levelMap = {
        1: "a:lvl1pPr",
        2: "a:lvl2pPr",
        3: "a:lvl3pPr",
        4: "a:lvl4pPr",
        5: "a:lvl5pPr",
        6: "a:lvl6pPr",
        7: "a:lvl7pPr",
        8: "a:lvl8pPr",
        9: "a:lvl9pPr"
    };

    const levelKey = levelMap[level];
    return levelKey ? lstStyle[levelKey]?.[0] : null;
}

function convertAlgnToCSS(algn) {
    switch (algn) {
        case "l": return "left";
        case "ctr": return "center";
        case "r": return "right";
        case "just": return "justify";
        case "dist": return "justify";
        default: return "left";
    }
}


function findPlaceholderInLayout(placeholderType, layoutXML) {
    if (!layoutXML || !placeholderType) return null;
    const shapes = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
    for (const shape of shapes) {
        const phType = shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
        if (phType === placeholderType) {
            return shape;
        }
    }
    return null;
}

function applyTintToColor(hexColor, tintValue) {
    if (!tintValue || tintValue === 0) return hexColor;
    if (!hexColor || hexColor === 'null' || hexColor === 'undefined') return hexColor;

    // Ensure hex color format
    hexColor = hexColor.replace('#', '');
    if (hexColor.length !== 6) return `#${hexColor}`;

    // Convert hex to RGB
    const r = parseInt(hexColor.substr(0, 2), 16);
    const g = parseInt(hexColor.substr(2, 2), 16);
    const b = parseInt(hexColor.substr(4, 2), 16);

    let newR, newG, newB;

    if (tintValue > 0) {
        // Tint (lighten) - blend with white
        const tintFactor = tintValue / 100000;
        newR = Math.round(r + (255 - r) * tintFactor);
        newG = Math.round(g + (255 - g) * tintFactor);
        newB = Math.round(b + (255 - b) * tintFactor);
    } else {
        // Shade (darken) - multiply by factor
        const shadeFactor = 1 + (tintValue / 100000);
        newR = Math.round(r * shadeFactor);
        newG = Math.round(g * shadeFactor);
        newB = Math.round(b * shadeFactor);
    }

    // Clamp values to 0-255 range
    newR = Math.max(0, Math.min(255, newR));
    newG = Math.max(0, Math.min(255, newG));
    newB = Math.max(0, Math.min(255, newB));

    // Convert back to hex
    return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
}

// Helper function to reset cache (useful for testing or when processing multiple presentations)
function resetTableStylesCache() {
    cachedTableStyles = null;
}


module.exports = {
    convertTableXMLToHTML,
    resetTableStylesCache
};