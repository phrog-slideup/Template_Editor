const fs = require("fs");
const path = require("path");

const resolveThemeFont = require("../../api/helper/resolveThemeFont.js");
const pptTextAllInfo = require("../../pptx-To-Html-styling/pptTextAllInfo.js");
const pptBackgroundColors = require("../../pptx-To-Html-styling/pptBackgroundColors.js");
const colorHelper = require("../../api/helper/colorHelper.js");
// If using CommonJS (Node.js)
const { getFontFallbackStack, loadWebFont } = require('./fontFallbacks');

// OR if using ES6 modules
// import { getFontFallbackStack, loadWebFont } from './fontFallbacks.js';
function getEMUDivisor() {
    return 12700;
}

function getAllTextInformationFromShape(shapeNode, themeXML, clrMap, masterXML, layoutXML) {

    const paragraphs = shapeNode?.["p:txBody"]?.[0]?.["a:p"] || [];

    let htmlContent = [];
    let currentListTag = null;
    let currentListType = null;
    let lineHeight = '1.2';
    let spaceBefore = '0';
    let spaceAfter = '0';
    let textAlign = 'left';
    let justifyContent = "flex-start";
    let textbgColor = '';
    let getAlignItem = '';
    const textKind = resolveThemeFont.getTextKindFromShape(shapeNode); // 'title' | 'body' | 'other'
    // console.log("font type =====",textKind);
    let txtPhPath = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"];

    let txtPhValues = [];
    if (txtPhPath && txtPhPath.length > 0) {
        txtPhValues = txtPhPath.map(ph => {
            return {
                type: ph["$"]?.type || '',
                idx: ph["$"]?.idx || '',
                sz: ph["$"]?.sz || ''
            };
        });
    }

    // FIXED: Add flag parameter to getPositionFromShape
    const position = pptTextAllInfo.getPositionFromShape(shapeNode);
    const rotation = pptTextAllInfo.getRotation(shapeNode);
    const placeholderType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
    getAlignItem = getAlignItemsFromPptx(shapeNode, layoutXML, placeholderType);
    // Get font size with proper flag parameter
    const fontSize = pptTextAllInfo.getFontSize(shapeNode,masterXML);
    // DEBUG: Add comprehensive font size debugging
    const allFontSizes = pptTextAllInfo.getAllFontSizesInShape(shapeNode);
    const shapeName = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
    const fontColor = pptTextAllInfo.getFontColor(shapeNode, themeXML, clrMap);
    const outlineStyle = extractTextOutline(shapeNode, themeXML, clrMap);

    let marginTop = 0;
    let marginRight = 0;
    let marginBottom = 0;
    let marginLeft = 0;

    try {
        const bodyPr = shapeNode?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
        
        const emuToPixels = (emuValue) => {
            if (!emuValue) return 0;
            const emu = parseInt(emuValue, 10);
            return emu / getEMUDivisor(); // 12700
        };
    
        // First try to get margins from the shape itself
        if (bodyPr && bodyPr["$"]) {
            const attributes = bodyPr["$"];
            marginTop = Math.round(emuToPixels(attributes.tIns));
            marginRight = Math.round(emuToPixels(attributes.rIns));
            marginBottom = Math.round(emuToPixels(attributes.bIns));
            marginLeft = Math.round(emuToPixels(attributes.lIns));
        }
    
        // ✅ NEW: If margins are still 0, inherit from layout/master
        if (marginLeft === 0 && marginRight === 0 && marginTop === 0 && marginBottom === 0) {
            const placeholderType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
            
            // Try layout first
            if (placeholderType && layoutXML) {
                const layoutShape = findPlaceholderInLayout(placeholderType, layoutXML);
                if (layoutShape) {
                    const layoutBodyPr = layoutShape?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
                    if (layoutBodyPr && layoutBodyPr["$"]) {
                        const layoutAttrs = layoutBodyPr["$"];
                        marginTop = Math.round(emuToPixels(layoutAttrs.tIns)) || marginTop;
                        marginRight = Math.round(emuToPixels(layoutAttrs.rIns)) || marginRight;
                        marginBottom = Math.round(emuToPixels(layoutAttrs.bIns)) || marginBottom;
                        marginLeft = Math.round(emuToPixels(layoutAttrs.lIns)) || marginLeft;
                        
                        console.log(`📦 Layout margins: T=${marginTop}px, R=${marginRight}px, B=${marginBottom}px, L=${marginLeft}px`);
                    }
                }
            }
    
            // ✅ NEW: If still 0, try master
            if (marginLeft === 0 && marginRight === 0 && marginTop === 0 && marginBottom === 0) {
                if (masterXML && placeholderType) {
                    const masterShapes = masterXML?.["p:sldMaster"]?.[0]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
                    
                    const masterShape = masterShapes.find(shape => {
                        const phType = shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
                        return phType === placeholderType;
                    });
    
                    if (masterShape) {
                        const masterBodyPr = masterShape?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
                        if (masterBodyPr && masterBodyPr["$"]) {
                            const masterAttrs = masterBodyPr["$"];
                            marginTop = Math.round(emuToPixels(masterAttrs.tIns)) || marginTop;
                            marginRight = Math.round(emuToPixels(masterAttrs.rIns)) || marginRight;
                            marginBottom = Math.round(emuToPixels(masterAttrs.bIns)) || marginBottom;
                            marginLeft = Math.round(emuToPixels(masterAttrs.lIns)) || marginLeft;
                            
                            console.log(`📦 Master margins: T=${marginTop}px, R=${marginRight}px, B=${marginBottom}px, L=${marginLeft}px`);
                        }
                    }
                }
            }
        }
    } catch (error) {
        console.error("Error extracting margin:", error);
    }

    paragraphs.forEach((paragraph, paragraphIndex) => {
        if (!paragraph) return;

        const isEmptyParagrahs = isEmptyParagraph(paragraph);

        if (isEmptyParagrahs) {
            // Handle empty paragraph
            const pPrNode = paragraph["a:pPr"]?.[0];

            const endParaRPr = paragraph["a:endParaRPr"]?.[0];
            const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"];

            // Extract styling from paragraph properties or endParaRPr
            textAlign = getTextAlignmentFromParagraph(pPrNode, shapeNode, placeholderType, layoutXML);

            // Calculate line height from endParaRPr if available
            let emptyParagraphLineHeight = lineHeight;
            if (endParaRPr?.["$"]?.sz) {
                const endParaFontSize = parseInt(endParaRPr["$"].sz) / 100;
                emptyParagraphLineHeight = Math.round(endParaFontSize * 1.2) + 'px';
            } else {
                emptyParagraphLineHeight = Math.round(fontSize * 1.2) + 'px';
            }

            // Extract line spacing from paragraph properties
            const lineSpacingNode = pPrNode?.["a:lnSpc"]?.[0];
            if (lineSpacingNode) {
                const lineSpacingPts = lineSpacingNode?.["a:spcPts"]?.[0]?.["$"]?.val;
                const lineSpacingPct = lineSpacingNode?.["a:spcPct"]?.[0]?.["$"]?.val;

                if (lineSpacingPts) {
                    const pointsValue = parseInt(lineSpacingPts) / 100;
                    emptyParagraphLineHeight = Math.round(pointsValue) + 'px';
                } else if (lineSpacingPct) {
                    let effectiveFontSize = fontSize;
                    if (endParaRPr?.["$"]?.sz) {
                        effectiveFontSize = parseInt(endParaRPr["$"].sz) / 100;
                    }
                    const pctVal = parseInt(lineSpacingPct);
                    const lineHeightPoints = (pctVal / 100000) * effectiveFontSize;
                    const minLineHeight = effectiveFontSize * 1.1;
                    const finalLineHeight = Math.max(lineHeightPoints, minLineHeight);
                    emptyParagraphLineHeight = Math.round(finalLineHeight) + 'px';
                }
            }

            switch (textAlign) {
                case "center":
                    justifyContent = "center";
                    break;
                case "right":
                    justifyContent = "flex-end";
                    break;
                case "justify":
                    justifyContent = "space-between";
                    break;
                case "left":
                default:
                    justifyContent = "flex-start";
                    break;
            }

            // Close any open list if this is an empty paragraph
            if (currentListTag) {
                htmlContent.push(`</${currentListTag}>`);
                currentListTag = null;
                currentListType = null;
            }

            // Add empty paragraph with proper styling
            // htmlContent.push(`<p style="text-align: ${textAlign}; line-height: ${emptyParagraphLineHeight}; ${margin}; min-height: ${emptyParagraphLineHeight};">&nbsp;</p>`);
            // Add empty paragraph with proper styling
            htmlContent.push(`<p style="text-align: ${textAlign}; line-height: ${emptyParagraphLineHeight}; margin-left: ${marginLeft}px; margin-right: ${marginRight}px; margin-top: ${marginTop}px; margin-bottom: ${marginBottom}px; min-height: ${emptyParagraphLineHeight};">&nbsp;</p>`);
            return;
        }

        // Pattern: <a:pPr><a:r><a:pPr><a:r><a:pPr><a:r>
        const bulletItems = [];

        // Get arrays of each element type in order
        const pPrElements = paragraph["a:pPr"] || [];
        const rElements = paragraph["a:r"] || [];
        const brElements = paragraph["a:br"] || [];

        // If we have multiple pPr elements, each should be a separate bullet item
        if (pPrElements.length > 1) {
            // Multiple bullet items in one paragraph
            for (let i = 0; i < pPrElements.length; i++) {
                const pPr = pPrElements[i];
                const bulletInfo = extractBulletInformation(pPr);

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
            const bulletInfo = extractBulletInformation(pPrNode);

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
                bulletInfo: extractBulletInformation(pPrNode)
            });
        }

        // Process each bullet item
        bulletItems.forEach((item, itemIndex) => {
            const pPrNode = item.pPr;
            const runs = item.runs;
            const lineBreaks = item.lineBreaks;
            const bulletInfo = item.bulletInfo;

            textAlign = getTextAlignmentFromParagraph(pPrNode, shapeNode, placeholderType, layoutXML);

            switch (textAlign) {
                case "center":
                    justifyContent = "center";
                    break;
                case "right":
                    justifyContent = "flex-end";
                    break;
                case "justify":
                    justifyContent = "space-between";
                    break;
                case "left":
                default:
                    justifyContent = "flex-start";
                    break;
            }

            // ✅ CRITICAL FIX: Extract FIRST-LINE INDENT
            let textIndent = 0;
            const indentAttr = pPrNode?.["$"]?.indent;
            if (indentAttr) {
                // Convert EMUs to pixels
                textIndent = Math.round(parseInt(indentAttr) / getEMUDivisor());
            }

            spaceBefore = parseInt(pPrNode?.["a:spcBef"]?.[0]?.["a:spcPts"]?.[0]?.["$"]?.val || "0", 10) / 100;
            spaceAfter = parseInt(pPrNode?.["a:spcAft"]?.[0]?.["a:spcPts"]?.[0]?.["$"]?.val || "0", 10) / 100;
            const lineSpacingNode = pPrNode?.["a:lnSpc"]?.[0];

            if (lineSpacingNode) {
                const lineSpacingPts = lineSpacingNode?.["a:spcPts"]?.[0]?.["$"]?.val;
                const lineSpacingPct = lineSpacingNode?.["a:spcPct"]?.[0]?.["$"]?.val;

                if (lineSpacingPts) {
                    const pointsValue = parseInt(lineSpacingPts) / 100;
                    lineHeight = Math.round(pointsValue) + 'px';
                } else if (lineSpacingPct) {
                    let effectiveFontSizeInPoints = fontSize;

                    const run = runs[0];
                    const runRPr = run?.["a:rPr"]?.[0];
                    const runFontSize = runRPr?.["$"]?.sz;
                    if (runFontSize) {
                        effectiveFontSizeInPoints = parseInt(runFontSize) / 100;
                    }

                    const pctVal = parseInt(lineSpacingPct);
                    const lineHeightPoints = (pctVal / 100000) * effectiveFontSizeInPoints;

                    const minLineHeight = effectiveFontSizeInPoints * 1.1;
                    const finalLineHeight = Math.max(lineHeightPoints, minLineHeight);

                    lineHeight = Math.round(finalLineHeight) + 'px';
                } else {
                    const defaultLineHeightPx = Math.round(fontSize * 1.2);
                    lineHeight = defaultLineHeightPx + 'px';
                }
            } 
            else {
                const defaultLineHeightPx = Math.round(fontSize * 1.2);
                lineHeight = defaultLineHeightPx + 'px';
            }

            const runTexts = [];

            const fallbackTxtWeight = shapeNode?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0]?.["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["$"]?.b;

            if (paragraph.$$ && Array.isArray(paragraph.$$)) {

                paragraph.$$.forEach(child => {
                    const childName = child['#name'];

                    if (childName === 'a:r') {
                        // Process run
                        let textElement = child?.["a:t"]?.[0];
                        let textValue = typeof textElement === 'string' ? textElement : "";

                        const capValue = child?.["a:rPr"]?.[0]?.["$"]?.cap;
                        if (capValue === "all") {
                            textValue = textValue.toUpperCase();
                        } else if (capValue === "none") {
                            textValue = textValue;
                        }

                        if (textValue !== undefined && textValue !== null) {
                            const runRPrNode = child?.["a:rPr"]?.[0];
                            const spanText = createSpanFromRun(runRPrNode, fallbackTxtWeight, textValue, lineHeight, themeXML, fontSize, masterXML, textKind, shapeNode, clrMap);
                            runTexts.push(spanText);
                        }
                    } else if (childName === 'a:br') {
                        runTexts.push('<br>');
                    }

                });
            } else {

                runs.forEach((run, runIndex) => {
                    let textElement = run?.["a:t"]?.[0];
                    let textValue = typeof textElement === 'string' ? textElement : "";

                    const capValue = run?.["a:rPr"]?.[0]?.["$"]?.cap;
                    if (capValue === "all") {
                        textValue = textValue.toUpperCase();
                    } else if (capValue === "none") {
                        textValue = textValue;
                    }

                    if (textValue !== undefined && textValue !== null) {
                        const runRPrNode = run?.["a:rPr"]?.[0];
                        const spanText = createSpanFromRun(runRPrNode, fallbackTxtWeight, textValue, lineHeight, themeXML, fontSize, masterXML, textKind, shapeNode, clrMap);
                        runTexts.push(spanText);
                    }
                });
            }

            const hasActualText = runs.some(run => {
                const textElement = run?.["a:t"]?.[0];
                const textValue = typeof textElement === 'string' ? textElement : "";

                return textValue.length > 0;
            });

            const hasContent = runTexts.length > 0 && hasActualText;

            if (bulletInfo.hasListMarker) {
                const listKey = `${bulletInfo.listTag}-${bulletInfo.listStyle}-${bulletInfo.bulletChar || bulletInfo.numberingType}`;

                if (currentListType !== listKey) {
                    if (currentListTag) htmlContent.push(`</${currentListTag}>`);

                    const listStyles = generateListStyles(bulletInfo);
                    htmlContent.push(`<${bulletInfo.listTag} style="  margin-left: ${marginLeft}px; margin-right: ${marginRight}px; margin-top: ${marginTop + spaceBefore}px; margin-bottom: ${marginBottom + spaceAfter}px; line-height: ${lineHeight}; ${listStyles}">`);
                    currentListTag = bulletInfo.listTag;
                    currentListType = listKey;
                }

                if (hasContent) {
                    htmlContent.push(`<li>${runTexts.join('')}</li>`);
                }
            } else {
                if (currentListTag) {
                    htmlContent.push(`</${currentListTag}>`);
                    currentListTag = null;
                    currentListType = null;
                }

                if (hasContent) {
                    const isFirstParagraph = paragraphIndex === 0;
                    const isLastParagraph = paragraphIndex === paragraphs.length - 1;
                
                    // Extract first-line indent
                    let textIndent = 0;
                    const indentAttr = pPrNode?.["$"]?.indent;
                    if (indentAttr) {
                        textIndent = Math.round(parseInt(indentAttr) / getEMUDivisor());
                    }
                
                    // Extract left margin from paragraph (ADDITIONAL to text box margin)
                    let paragraphMarginLeft = 0;
                    let paragraphMarginTop = 0;
                    let paragraphMarginRight = 0;
                    let paragraphMarginBottom = 0;
                    
                    const marLAttr = pPrNode?.["$"]?.marL;
                    
                    if (marLAttr) {
                        // Direct margin from slide
                        paragraphMarginLeft = Math.round(parseInt(marLAttr) / getEMUDivisor());
                    } else {
                        // 🔧 FIX: Inherit from master/layout if missing (returns all 4 margins)
                        const inheritedMargins = getInheritedParagraphMargin(
                            shapeNode, 
                            layoutXML, 
                            masterXML, 
                            placeholderType,
                            pPrNode?.["$"]?.lvl ? parseInt(pPrNode["$"].lvl) : 0
                        );
                        
                        if (inheritedMargins !== null) {
                            paragraphMarginLeft = inheritedMargins.left || 0;
                            paragraphMarginTop = inheritedMargins.top || 0;
                            paragraphMarginRight = inheritedMargins.right || 0;
                            paragraphMarginBottom = inheritedMargins.bottom || 0;
                        }
                    }
                    
                    const hasMargin = paragraphMarginLeft > 0;
                    if (hasMargin && runTexts.length > 0) {
                        // Strip leading &nbsp; entities AND regular spaces from first run only
                        runTexts[0] = runTexts[0].replace(/^(&nbsp;)+/, '');
                    }

                    let effectiveMarginTop;
                    let effectiveMarginBottom;
                    let effectiveMarginLeft;
                    let effectiveMarginRight;
                
                    if (isFirstParagraph) {                        
                        // First paragraph: paragraph spacing + all paragraph-specific margins
                        effectiveMarginTop = paragraphMarginTop;
                        effectiveMarginLeft = paragraphMarginLeft;  // Don't add marginLeft (text box margin)
                        effectiveMarginRight = paragraphMarginRight;  // Add paragraph-specific right margin
                        effectiveMarginBottom = 0; // Will be set below if last paragraph
                    } else {                        
                        // Subsequent paragraphs: paragraph spacing + all paragraph-specific margins
                        effectiveMarginTop = spaceBefore + paragraphMarginTop;
                        effectiveMarginLeft = paragraphMarginLeft;
                        effectiveMarginRight = paragraphMarginRight;
                        effectiveMarginBottom = 0;
                    }
                
                    if (isLastParagraph) {
                        effectiveMarginBottom = spaceAfter + paragraphMarginBottom;
                    } else {
                        effectiveMarginBottom = spaceAfter;
                    }

                    const paragraphStyle = `text-align: ${textAlign}; line-height: ${lineHeight}; margin-left: ${effectiveMarginLeft}px; margin-right: ${effectiveMarginRight}px; margin-top: ${effectiveMarginTop}px; margin-bottom: ${effectiveMarginBottom}px;${textIndent !== 0 ? ` text-indent: ${textIndent}px;` : ''}`;
                
                    htmlContent.push(`<p style="${paragraphStyle}">${runTexts.join('')}</p>`);
                }
                else if (runTexts.length > 0 || lineBreaks.length > 0) {
                    htmlContent.push(`<p style="text-align: ${textAlign}; line-height: ${lineHeight}; margin-left: ${marginLeft}px; margin-right: ${marginRight}px; margin-top: ${marginTop + spaceBefore}px; margin-bottom: ${marginBottom + spaceAfter}px; min-height: ${lineHeight};"><br/></p>`);
                }
            }
        });

    });

    if (currentListTag) {
        htmlContent.push(`</${currentListTag}>`);
    }

    const text = htmlContent.join("\n");

    // Calculate estimated content height
    let estimatedHeight = marginTop + marginBottom;
    paragraphs.forEach((paragraph) => {
        if (!paragraph) return;
        const pPrNode = paragraph["a:pPr"]?.[0];

        let pLineHeight = parseInt(lineHeight) || Math.round(fontSize * 1.2);
        const lineSpacingNode = pPrNode?.["a:lnSpc"]?.[0];
        if (lineSpacingNode?.["a:spcPts"]?.[0]?.["$"]?.val) {
            pLineHeight = parseInt(lineSpacingNode["a:spcPts"][0]["$"].val) / 100;
        }

        const pSpaceBefore = parseInt(pPrNode?.["a:spcBef"]?.[0]?.["a:spcPts"]?.[0]?.["$"]?.val || "0") / 100;
        const pSpaceAfter = parseInt(pPrNode?.["a:spcAft"]?.[0]?.["a:spcPts"]?.[0]?.["$"]?.val || "0") / 100;

        // rakesh notes ::: height restricted making the height of the textbox cut off
        const runs = paragraph["a:r"] || [];
        const hasText = runs.some(run => {
            const textElement = run?.["a:t"]?.[0];
            return textElement && typeof textElement === 'string' && textElement.length > 0;
        });

        if (hasText) {
            // Estimate multiple lines based on text length
            let textLength = 0;
            runs.forEach(run => {
                const textElement = run?.["a:t"]?.[0];
                if (textElement && typeof textElement === 'string') {
                    textLength += textElement.length;
                }
            });

            // Estimate lines: ~60 chars per line for 863px width at 14px font
            const charsPerLine = Math.floor(position.width / (fontSize * 0.6));
            const estimatedLines = Math.max(1, Math.ceil(textLength / charsPerLine));
            estimatedHeight += (pLineHeight * estimatedLines) + pSpaceBefore + pSpaceAfter;
        } else {
            estimatedHeight += pLineHeight + pSpaceBefore + pSpaceAfter;
        }

    });

    return {
        position: position,
        text: text || " ",
        txtPhValues,
        textAlign: textAlign,
        // justifyContent: justifyContent,
        fontSize: fontSize,
        textbgColor: textbgColor || '',
        fontColor: fontColor,
        rotation: rotation,
        lineHeight: lineHeight,
        spaceBefore: spaceBefore,
        spaceAfter: spaceAfter,
        getAlignItem: getAlignItem,
        outlineStyle: outlineStyle,
        htmlContent: htmlContent.join("\n"),
        estimatedContentHeight: Math.max(estimatedHeight, position?.height || 0)  // NEW LINE
    };

}


function getAlignItemsFromPptx(shapeNode, layoutXML, placeholderType) {

    const anchor = shapeNode?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0]?.["$"]?.anchor;
    let txtalg = '';
    if (anchor) {
        txtalg = convertAnchorToFlexAlign(anchor);

        return txtalg;
    }

    if (placeholderType && layoutXML) {
        const layoutShape = findPlaceholderInLayout(placeholderType, layoutXML);

        if (layoutShape) {
            const layoutAnchor = layoutShape?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0]?.["$"]?.anchor;

            if (layoutAnchor) {
                let AnchorToFlexAlign = convertAnchorToFlexAlign(layoutAnchor);
                return AnchorToFlexAlign;
            }
        }
    }

    return "flex-start";
}

// NEW: Helper method to detect empty paragraphs
function isEmptyParagraph(paragraph) {
    // Check if paragraph has no runs or only empty runs
    const runs = paragraph["a:r"] || [];
    const lineBreaks = paragraph["a:br"] || [];

    // If there are line breaks, it's not empty
    if (lineBreaks.length > 0) {
        return false;
    }

    // If there are no runs at all, it's empty
    if (runs.length === 0) {
        return true;
    }

    // Check if all runs are empty (no text content)
    const hasTextContent = runs.some(run => {
        const textElement = run?.["a:t"]?.[0];
        const textValue = typeof textElement === 'string' ? textElement : "";
        return textValue.trim() !== '';
    });

    return !hasTextContent;
}


function getTextAlignmentFromParagraph(pPrNode, shapeNode, placeholderType, layoutXML) {
    // First, check if paragraph has explicit alignment
    if (pPrNode && pPrNode['$'] && pPrNode['$']['algn']) {
        const algn = pPrNode['$']['algn'];
        switch (algn) {
            case "l": return "left";
            case "ctr": return "center";
            case "r": return "right";
            case "just": return "justify";
            default: return "left";
        }
    }

    // Check list style in current shape
    const lstStyle = shapeNode?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
    if (lstStyle) {
        const paragraphLevel = pPrNode?.["$"]?.lvl ? parseInt(pPrNode["$"].lvl) + 1 : 1;
        let levelPPr = getLevelProperties(lstStyle, paragraphLevel);

        if (levelPPr && levelPPr["$"] && levelPPr["$"]["algn"]) {
            return convertAlgnToCSS(levelPPr["$"]["algn"]);
        }

        const defPPr = lstStyle["a:defPPr"]?.[0];
        if (defPPr && defPPr["$"] && defPPr["$"]["algn"]) {
            return convertAlgnToCSS(defPPr["$"]["algn"]);
        }
    }

    if (placeholderType && layoutXML) {
        const layoutShape = findPlaceholderInLayout(placeholderType, layoutXML);
        if (layoutShape) {
            const layoutLstStyle = layoutShape?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
            if (layoutLstStyle) {
                const paragraphLevel = pPrNode?.["$"]?.lvl ? parseInt(pPrNode["$"].lvl) + 1 : 1;
                let levelPPr = getLevelProperties(layoutLstStyle, paragraphLevel);

                if (levelPPr && levelPPr["$"] && levelPPr["$"]["algn"]) {
                    return convertAlgnToCSS(levelPPr["$"]["algn"]);
                }

                const defPPr = layoutLstStyle["a:defPPr"]?.[0];
                if (defPPr && defPPr["$"] && defPPr["$"]["algn"]) {
                    return convertAlgnToCSS(defPPr["$"]["algn"]);
                }
            }
        }
    }

    return "left";
}


function extractBulletInformation(pPrNode, shapeNode) {
    // --- 0️⃣ Default return object ---
    const defaultBulletInfo = {
        hasListMarker: false,
        listTag: null,
        listStyle: null,
        bulletChar: null,
        bulletFont: null,
        bulletSize: null,
        bulletColor: null,
        numberingType: null,
        marginLeft: null,
        indent: null,
        level: 0
    };

    if (!pPrNode && !shapeNode) return defaultBulletInfo;

    // --- 1️⃣ Explicit bullet disabling (a:buNone) ---
    const hasBuNone =
        pPrNode?.["a:buNone"] !== undefined ||
        (Array.isArray(pPrNode?.["a:buNone"]) && pPrNode["a:buNone"].length > 0);
    if (hasBuNone) return defaultBulletInfo;

    // --- 2️⃣ Direct bullet definitions inside paragraph ---
    const bulletChar = pPrNode?.["a:buChar"]?.[0]?.["$"]?.char || null;
    const autoNum = pPrNode?.["a:buAutoNum"]?.[0]?.["$"]?.type || null;
    const bulletFont = pPrNode?.["a:buFont"]?.[0]?.["$"]?.typeface || null;
    const bulletSizeRaw = pPrNode?.["a:buSzPct"]?.[0]?.["$"]?.val || null;
    const bulletSize = bulletSizeRaw ? parseInt(bulletSizeRaw, 10) / 1000 : null;
    const bulletColor = pPrNode?.["a:buClr"]?.[0] || null;
    const bulletScheme = pPrNode?.["a:buAutoNum"]?.[0]?.["$"]?.startAt || null;

    const marginLeft = pPrNode?.["$"]?.marL ? parseInt(pPrNode["$"].marL) : null;
    const indent = pPrNode?.["$"]?.indent ? parseInt(pPrNode["$"].indent) : null;
    const level = pPrNode?.["$"]?.lvl ? parseInt(pPrNode["$"].lvl) : 0;

    // --- 3️⃣ Explicit bullets ---
    if (bulletChar) {
        return {
            hasListMarker: true,
            listTag: "ul",
            listStyle: getBulletStyleFromChar(bulletChar, bulletFont),
            bulletChar,
            bulletFont,
            bulletSize,
            bulletColor,
            marginLeft,
            indent,
            level
        };
    }

    // --- 4️⃣ Auto numbering ---
    if (autoNum) {
        return {
            hasListMarker: true,
            listTag: "ol",
            listStyle: getNumberingStyleFromType(autoNum),
            numberingType: autoNum,
            bulletSize,
            bulletColor,
            startAt: bulletScheme,
            marginLeft,
            indent,
            level
        };
    }

    // --- 5️⃣ Inherited bullets from <a:lstStyle> (if no buNone and no explicit bullet) ---
    try {
        const lstStyle = shapeNode?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
        if (lstStyle) {
            const lvl1 = lstStyle["a:lvl1pPr"]?.[0];
            if (lvl1?.["a:buChar"]?.[0]?.["$"]?.char) {
                const inheritedChar = lvl1["a:buChar"][0]["$"].char;
                const inheritedFont = lvl1["a:buFont"]?.[0]?.["$"]?.typeface || "Arial";
                const inheritedColor = lvl1["a:buClr"]?.[0] || null;
                const inheritedSizeRaw = lvl1?.["a:buSzPct"]?.[0]?.["$"]?.val || null;
                const inheritedSize = inheritedSizeRaw ? parseInt(inheritedSizeRaw, 10) / 1000 : null;

                return {
                    hasListMarker: true,
                    listTag: "ul",
                    listStyle: getBulletStyleFromChar(inheritedChar, inheritedFont),
                    bulletChar: inheritedChar,
                    bulletFont: inheritedFont,
                    bulletColor: inheritedColor,
                    bulletSize: inheritedSize,
                    marginLeft,
                    indent,
                    level
                };
            }

            if (lvl1?.["a:buAutoNum"]?.[0]?.["$"]?.type) {
                const inheritedNumType = lvl1["a:buAutoNum"][0]["$"].type;
                const inheritedStart = lvl1["a:buAutoNum"][0]["$"].startAt || 1;
                const inheritedColor = lvl1["a:buClr"]?.[0] || null;
                const inheritedSizeRaw = lvl1?.["a:buSzPct"]?.[0]?.["$"]?.val || null;
                const inheritedSize = inheritedSizeRaw ? parseInt(inheritedSizeRaw, 10) / 1000 : null;

                return {
                    hasListMarker: true,
                    listTag: "ol",
                    listStyle: getNumberingStyleFromType(inheritedNumType),
                    numberingType: inheritedNumType,
                    bulletColor: inheritedColor,
                    bulletSize: inheritedSize,
                    startAt: inheritedStart,
                    marginLeft,
                    indent,
                    level
                };
            }
        }
    } catch (err) {
        console.warn("Bullet inheritance check failed:", err);
    }

    return defaultBulletInfo;
}


//MODIFIED: Update createSpanFromRun to accept outlineStyle parameter
function createSpanFromRun(runRPrNode, fallbackTxtWeight, textValue, lineHeight, themeXML, shapeFontSize, masterXML, textKind, shapeNode, clrMap) {
    if (!runRPrNode) {
        return `<span class="span-txt" style="white-space: pre-wrap; font-size: ${shapeFontSize}px; line-height: ${lineHeight};">${textValue}</span>`;
    }

    let originalTxtColor, originalLumOff, originalLumMod;

    const fontSize = runRPrNode?.["$"]?.sz ?
        parseInt(runRPrNode["$"].sz) / 100 :
        shapeFontSize;

    // Font inheritance logic (existing code)
    const runHasLatinFont = runRPrNode?.["a:latin"]?.[0]?.["$"]?.typeface;
    const runHasEaFont = runRPrNode?.["a:ea"]?.[0]?.["$"]?.typeface;
    const runHasCsFont = runRPrNode?.["a:cs"]?.[0]?.["$"]?.typeface;
    const runHasSymFont = runRPrNode?.["a:sym"]?.[0]?.["$"]?.typeface;

    const runHasAnyFont = runHasLatinFont || runHasEaFont || runHasCsFont || runHasSymFont;

    let latinFont, originalEA, originCS, originSYM, runTypeface;
    let inheritedDefRPr = null;

    const hasHyperlink = runRPrNode?.["a:hlinkClick"];
    let hyperlinkStart = '';
    let hyperlinkEnd = '';
    
    if (hasHyperlink) {
        const tooltip = hasHyperlink[0]?.["$"]?.tooltip || '';
        const rId = hasHyperlink[0]?.["$"]?.["r:id"] || '';
        
        // You'll need to resolve rId to actual URL from slide relationships
        hyperlinkStart = `<a href="#" title="${tooltip}" style="color: blue; text-decoration: underline;">`;
        hyperlinkEnd = '</a>';
    }

    if (runHasAnyFont) {
        latinFont = runHasLatinFont || "";
        originalEA = runHasEaFont || "";
        originCS = runHasCsFont || "";
        originSYM = runHasSymFont || "";
        runTypeface = latinFont;
    } else {
        const inheritedData = getInheritedPropertiesFromShape(shapeNode);

        if (inheritedData) {
            latinFont = inheritedData.fonts.latin || "";
            originalEA = inheritedData.fonts.ea || "";
            originCS = inheritedData.fonts.cs || "";
            originSYM = inheritedData.fonts.sym || "";
            runTypeface = latinFont;
            inheritedDefRPr = inheritedData.defRPr;
        } else {
            latinFont = "";
            originalEA = "";
            originCS = "";
            originSYM = "";
            runTypeface = null;
        }
    }

    const masterDefaultTypeface = resolveThemeFont.getDefaultTypefaceFromMaster(masterXML, textKind);

    let fontFamily;
    if (runTypeface) {
        fontFamily = resolveThemeFont.resolveThemeFont(runTypeface, themeXML);
    } else if (masterDefaultTypeface) {
        fontFamily = resolveThemeFont.resolveThemeFont(masterDefaultTypeface, themeXML);
    } else {
        fontFamily = 'Calibri';
    }

    if (fontFamily == "Univers Condensed Light (Body)") {
        fontFamily = "UniversCondensedLightBody";
    }

    if (fontFamily?.startsWith?.('+')) fontFamily = 'Calibri';

    let fontWeight = runRPrNode?.["$"]?.b;

    if (!fontWeight && fallbackTxtWeight == 1) {
        fontWeight = fallbackTxtWeight;
    }
    if (fontWeight == 1) {
        fontWeight = "bold";
    } else if (!fontWeight && fontFamily == "Helvetica Neue Ltd Std-Bd") {
        fontWeight = "bold";
    } else if (!fontWeight == 0 && fontFamily == "Helvetica Neue Ltd Std-Bd") {
        fontFamily = "Helvetica Neue Ltd Std";
        fontWeight = "normal";
    } else {
        fontWeight = "normal";
    }

    const fontStyle = runRPrNode?.["$"]?.i === "1" ? "italic" : "normal";

    // Capitalization handling (existing code)
    let textTransform = 'none';
    let fontVariant = 'normal';

    let capValue = runRPrNode?.["$"]?.cap;

    if (!capValue && inheritedDefRPr) {
        capValue = inheritedDefRPr?.["$"]?.cap;
    }

    if (capValue === "all") {
        textTransform = 'uppercase';
    } else if (capValue === "small") {
        fontVariant = 'small-caps';
    }

    // Extract text color OR gradient
    let color = "#000000";
    let opacity = 1;
    let textGradientCSS = '';
    let gradientDataAttrs = ''; // NEW: Store gradient data as attributes

    const solidFillNode = runRPrNode?.["a:solidFill"]?.[0];

    // Check for gradient fill FIRST (priority over solid fill)
    const textGradient = extractTextGradient(runRPrNode, themeXML);
    if (textGradient) {
        textGradientCSS = textGradient.css;
        color = '';
        if (textGradient.type === 'linear') {
            gradientDataAttrs = `data-has-gradient="true" data-gradient-type="linear" data-gradient-angle="${textGradient.degrees}" data-gradient-stops='${JSON.stringify(textGradient.stops)}'`;
        } else if (textGradient.type === 'radial') {
            gradientDataAttrs = `data-has-gradient="true" data-gradient-type="radial" data-gradient-path="${textGradient.pathType}" data-gradient-center-x="${textGradient.centerX}" data-gradient-center-y="${textGradient.centerY}" data-gradient-stops='${JSON.stringify(textGradient.stops)}'`;
        }
        // Set a fallback color (first gradient stop color) for visibility
        if (textGradient.stops && textGradient.stops.length > 0) {
            color = textGradient.stops[0].color;
        }
    }
    // If no gradient, check for solid fill
    else if (solidFillNode) {
        if (solidFillNode["a:srgbClr"]) {
            originalTxtColor = solidFillNode["a:srgbClr"][0]["$"].val;
            originalLumMod = solidFillNode["a:srgbClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            originalLumOff = solidFillNode["a:srgbClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;

            const srgbNode = solidFillNode["a:srgbClr"][0];
            color = `#${srgbNode["$"].val}`;

            const lumMod = srgbNode["a:lumMod"]?.[0]?.["$"]?.val;
            if (lumMod) {
                const lumOff = srgbNode["a:lumOff"]?.[0]?.["$"]?.val;
                if (lumOff) {
                    color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
                } else {
                    color = colorHelper.applyLumMod(color, lumMod);
                }
            }

            const alpha = srgbNode["a:alpha"]?.[0]?.["$"]?.val;
            if (alpha) {
                opacity = parseInt(alpha, 10) / 100000;
            }
        } else if (solidFillNode?.["a:schemeClr"]) {
            originalTxtColor = solidFillNode["a:schemeClr"][0]["$"].val;
            originalLumMod = solidFillNode["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            originalLumOff = solidFillNode["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;

            let schemeNode = solidFillNode["a:schemeClr"][0]["$"].val;

            if (masterXML && schemeNode) {
                const resolvedColor = resolveMasterColor(schemeNode, masterXML);
                if (resolvedColor) {
                    schemeNode = resolvedColor;
                }
            }

            if (schemeNode) {
                color = colorHelper.resolveThemeColorHelper(schemeNode, themeXML, masterXML);
            }

            const lumMod = solidFillNode["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
            if (lumMod) {
                const lumOff = solidFillNode["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                if (lumOff) {
                    color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
                } else {
                    color = colorHelper.applyLumMod(color, lumMod);
                }
            }
        } else if (solidFillNode["a:prstClr"]) {
            originalTxtColor = solidFillNode["a:prstClr"][0]["$"].val;
            color = solidFillNode["a:prstClr"]?.[0]?.["$"]?.val;
        }
    }


    // Fallback to master color (existing code)
    if (color == "#000000" || !color) {
        try {
            const masterFontColor = pptTextAllInfo.getFontColor(shapeNode, themeXML, clrMap, masterXML);

            const txStyles = masterXML?.["p:sldMaster"]?.["p:txStyles"]?.[0];
            if (txStyles) {
                let masterColorKey = txStyles["p:titleStyle"][0]["a:lvl1pPr"]?.[0]?.["a:defRPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:schemeClr"]?.[0]?.["$"]?.val;
                const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type

                if (masterColorKey, phType === "ctrTitle" || phType === "title") {
                    originalTxtColor = masterColorKey;
                }
            }

            if (masterFontColor && masterFontColor !== "#000000") {
                color = masterFontColor;
            }
        } catch (err) {
            console.warn("Error applying master text color fallback:", err);
        }
    }

    let finalColor;
    if (opacity < 1) {
        const rgbValues = hexToRGB(color);
        finalColor = `rgba(${rgbValues}, ${opacity})`;
    } else {
        finalColor = color;
    }

    // Baseline for superscript/subscript (existing code)
    const baseline = runRPrNode?.["$"]?.baseline;
    let positionStyle = '';
    let adjustedFontSize = fontSize;

    if (baseline) {
        const baselineValue = parseInt(baseline);
        // PowerPoint uses 30000 = 30% for typical superscript
        const percent = baselineValue / 100000; // Convert to decimal
        
        if (baselineValue > 0) {
            // Superscript: negative top offset, smaller font
            const topOffset = -(fontSize * 0.4); // More pronounced lift
            adjustedFontSize = fontSize * 0.58; // Smaller size (58% instead of 67%)
            positionStyle = `position: relative; top: ${topOffset}px; vertical-align: super; `;
        } else if (baselineValue < 0) {
            // Subscript: positive top offset, smaller font
            const topOffset = fontSize * 0.2;
            adjustedFontSize = fontSize * 0.58;
            positionStyle = `position: relative; top: ${topOffset}px; vertical-align: sub; `;
        }
    }

    // Text decoration (existing code)
    let textDecorations = [];
    if (runRPrNode?.['$']?.u === "sng" || runRPrNode?.['$']?.u === "1") {
        textDecorations.push("underline");
    }
    if (runRPrNode?.['$']?.strike === "1") {
        textDecorations.push("line-through");
    }

    // NEW: Apply outline style if present
    // const outlineCSS = outlineStyle && outlineStyle.css ? outlineStyle.css : '';

    const styles = [
        fontWeight !== 'normal' ? `font-weight: ${fontWeight}` : '',
        fontStyle !== 'normal' ? `font-style: ${fontStyle}` : '',
        `font-family: ${getFontFallbackStack(fontFamily)}`,
        // `font-family: ${fontFamily}`,
        `font-size: ${adjustedFontSize}px`,
        // `color: ${finalColor}`,
        finalColor ? `color: ${finalColor}` : '', // Only add color if not using gradient
        textGradientCSS, // Add gradient CSS
        textTransform !== 'none' ? `text-transform: ${textTransform}` : '',
        fontVariant !== 'normal' ? `font-variant: ${fontVariant}` : '',
        positionStyle.trim() ? positionStyle.trim() : '',
        textDecorations.length > 0 ? `text-decoration: ${textDecorations.join(" ")}` : '',
        //outlineCSS, // ✅ Add outline CSS
        'white-space: pre-wrap',
        `line-height: ${lineHeight}`
    ].filter(s => s).join('; ');

    // ✅ CRITICAL FIX: Preserve leading spaces (convert to &nbsp;)
    const preservedText = textValue.replace(/^ +/, match => '&nbsp;'.repeat(match.length));

    return `<span class="span-txt" ${gradientDataAttrs} originalEA="${originalEA}" originCS="${originCS}" originSYM="${originSYM}" latinFont="${latinFont}" originalTxtColor="${originalTxtColor}" originalLumMod="${originalLumMod}" originalLumOff="${originalLumOff}" alpha="${opacity}" cap="${capValue || ''}" style="${styles};">${preservedText}</span>`;
    
}

// NEW: Extract text gradient fill
function extractTextGradient(runRPrNode, themeXML) {
    try {
        if (!runRPrNode) return null;

        const gradFill = runRPrNode?.["a:gradFill"]?.[0];

        if (!gradFill) return null;

        // Check for LINEAR gradient
        const lin = gradFill?.["a:lin"]?.[0];

        // Check for RADIAL/PATH gradient
        const path = gradFill?.["a:path"]?.[0];

        // Extract gradient stops (common for both types)
        const gsLst = gradFill?.["a:gsLst"]?.[0];
        const gradientStops = gsLst?.["a:gs"];

        if (!gradientStops || gradientStops.length === 0) return null;

        let stops = [];

        // Process gradient stops
        gradientStops.forEach(stop => {
            const pos = stop?.["$"]?.pos;
            const position = pos ? (parseInt(pos) / 1000) : 0;

            let color = '#000000';

            // Check for scheme color
            const schemeClr = stop?.["a:schemeClr"]?.[0];
            if (schemeClr) {
                const schemeVal = schemeClr["$"]?.val;
                color = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null);

                // Apply color modifications
                const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
                const lumOff = schemeClr["a:lumOff"]?.[0]?.["$"]?.val;
                const tint = schemeClr["a:tint"]?.[0]?.["$"]?.val;
                const shade = schemeClr["a:shade"]?.[0]?.["$"]?.val;

                if (lumMod && lumOff) {
                    color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
                } else if (lumMod) {
                    color = colorHelper.applyLumMod(color, lumMod);
                }

                if (tint) {
                    color = applyShadeOrTint(color, tint, 'tint');
                }
                if (shade) {
                    color = applyShadeOrTint(color, shade, 'shade');
                }
            }

            // Check for direct RGB color
            const srgbClr = stop?.["a:srgbClr"]?.[0];
            if (srgbClr) {
                color = `#${srgbClr["$"].val}`;

                const lumMod = srgbClr["a:lumMod"]?.[0]?.["$"]?.val;
                const lumOff = srgbClr["a:lumOff"]?.[0]?.["$"]?.val;

                if (lumMod && lumOff) {
                    color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
                } else if (lumMod) {
                    color = colorHelper.applyLumMod(color, lumMod);
                }
            }

            stops.push({ position, color });
        });

        // Sort by position
        stops.sort((a, b) => a.position - b.position);

        // Build gradient string based on type
        const gradientString = stops.map(s => `${s.color} ${s.position}%`).join(', ');

        // ========================================
        // LINEAR GRADIENT
        // ========================================
        if (lin) {
            const angle = lin?.["$"]?.ang;
            const pptDegrees = angle ? (parseInt(angle) / 60000) : 90;
            let cssDegrees = (pptDegrees + 90) % 360;

            return {
                type: 'linear',
                degrees: cssDegrees,
                stops,
                gradientString,
                css: `background: linear-gradient(${cssDegrees}deg, ${gradientString}); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;`
            };
        }

        // ========================================
        // RADIAL/PATH GRADIENT
        // ========================================
        if (path) {
            const pathType = path?.["$"]?.path; // "circle" or "rect" or "shape"

            // Extract fill-to-rect (defines the focal point and size)
            const fillToRect = path?.["a:fillToRect"]?.[0]?.["$"];

            let centerX = 50; // Default center
            let centerY = 50;
            let radiusX = 50;
            let radiusY = 50;

            if (fillToRect) {
                // PowerPoint uses percentages * 1000
                // l = left, t = top, r = right, b = bottom
                const l = fillToRect.l ? parseInt(fillToRect.l) / 1000 : 0;
                const t = fillToRect.t ? parseInt(fillToRect.t) / 1000 : 0;
                const r = fillToRect.r ? parseInt(fillToRect.r) / 1000 : 0;
                const b = fillToRect.b ? parseInt(fillToRect.b) / 1000 : 0;

                // Calculate center point
                centerX = l + ((100 - l - r) / 2);
                centerY = t + ((100 - t - b) / 2);

                // Calculate radii (approximation)
                radiusX = (100 - l - r) / 2;
                radiusY = (100 - t - b) / 2;
            }

            let shape = 'circle';
            let size = 'farthest-corner';

            // Map PowerPoint path types to CSS
            if (pathType === 'circle') {
                shape = 'circle';
                size = 'farthest-corner';
            } else if (pathType === 'rect') {
                shape = 'ellipse';
                size = 'farthest-corner';
            } else if (pathType === 'shape') {
                // Shape follows the text/object shape
                shape = 'ellipse';
                size = 'closest-side';
            }

            // Build radial gradient CSS
            const radialCSS = `background: radial-gradient(${shape} ${size} at ${centerX}% ${centerY}%, ${gradientString}); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;`;

            return {
                type: 'radial',
                pathType,
                centerX,
                centerY,
                radiusX,
                radiusY,
                shape,
                size,
                stops,
                gradientString,
                css: radialCSS
            };
        }

        // Fallback: if neither lin nor path is found, default to linear
        return {
            type: 'linear',
            degrees: 90,
            stops,
            gradientString,
            css: `background: linear-gradient(90deg, ${gradientString}); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;`
        };

    } catch (error) {
        console.error("Error extracting text gradient:", error);
        return null;
    }
}



/** IMPROVED: Better list style generation with proper indentation */
function generateListStyles(bulletInfo) {
    let styles = '';

    // Base list styles
    styles += 'list-style-position: outside; ';
    styles += 'padding-left: 20px; '; // Standard indent for nested lists
    styles += 'margin-left: 0; ';

    // Apply list style type
    if (bulletInfo.listTag === 'ul') {
        // Unordered list styles
        if (bulletInfo.bulletChar) {
            // Custom bullet character
            styles += `list-style-type: '${bulletInfo.bulletChar}  '; `;
        } else {
            // Default bullet style
            styles += `list-style-type: ${bulletInfo.listStyle || 'disc'}; `;
        }
    } else if (bulletInfo.listTag === 'ol') {
        // Ordered list styles
        if (bulletInfo.numberingType) {
            styles += `list-style-type: ${bulletInfo.numberingType}; `;
        } else {
            styles += 'list-style-type: decimal; ';
        }

        // Add start attribute if specified
        if (bulletInfo.startAt && bulletInfo.startAt > 1) {
            // Note: This should be handled as an HTML attribute, not CSS
            // But we include it in the styles string for reference
        }
    }

    // Apply bullet color if specified
    if (bulletInfo.bulletColor) {
        styles += `color: ${bulletInfo.bulletColor}; `;
    }

    // Apply bullet size if specified
    if (bulletInfo.bulletSize) {
        styles += `font-size: ${bulletInfo.bulletSize}%; `;
    }

    return styles.trim();
}

function convertAnchorToFlexAlign(anchor) {
    switch (anchor) {
        case "ctr": return "center";
        case "b": return "flex-end";
        case "t": return "flex-start";
        case "just": return "space-between";
        case "dist": return "space-around";
        default: return "flex-start";
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

// IMPROVED: Better bullet style detection with proper character codes
function getBulletStyleFromChar(bulletChar, bulletFont) {
    // Normalize font name for comparison
    const normalizedFont = bulletFont ? bulletFont.toLowerCase() : '';

    // Handle Wingdings font characters
    if (normalizedFont.includes('wingdings')) {
        return getWingdingsStyle(bulletChar);
    }

    // Handle Symbol font characters
    if (normalizedFont.includes('symbol')) {
        return getSymbolFontStyle(bulletChar);
    }

    // FIXED: Use character codes and proper Unicode characters
    const standardStyles = {
        // Common bullet characters by Unicode/ASCII code
        '\u2022': 'disc',
        '\u25CF': 'disc',
        '\u25CB': 'circle',
        '\u25A0': 'square',
        '\u25A1': 'square',
        '\u25B6': 'triangle',
        '\u25C6': 'diamond',
        '\u25C7': 'diamond',
        '\u2713': 'checkmark',
        '\u2717': 'x-mark',
        '\u2192': 'arrow',
        '\u2190': 'arrow',
        '\u2605': 'star',
        '\u2606': 'star',
        '\u00B7': 'disc',
        '\u2043': 'disc',
        '\u2014': 'disc',
        '\u2013': 'disc',

        // ASCII fallbacks
        '*': 'disc',
        '-': 'disc',
        '+': 'disc',
        'o': 'circle',
        'O': 'circle',
        '>': 'arrow',
        '~': 'disc'
    };

    // Try exact match first
    if (standardStyles[bulletChar]) {
        return standardStyles[bulletChar];
    }

    // IMPROVED: Try to detect by character code if direct match fails
    const charCode = bulletChar.charCodeAt(0);

    // Unicode bullet range detection
    if (charCode >= 0x2022 && charCode <= 0x2043) {
        return 'disc'; // Most bullets in this range
    }
    if (charCode >= 0x25A0 && charCode <= 0x25FF) {
        return 'square'; // Geometric shapes range
    }
    if (charCode >= 0x2600 && charCode <= 0x26FF) {
        return 'star'; // Miscellaneous symbols range
    }

    return 'disc'; // Default fallback
}

// Enhanced numbering style detection with all PowerPoint types
function getNumberingStyleFromType(autoNumType) {
    const numberingStyles = {
        // Arabic numerals
        'arabicPeriod': 'decimal',              // 1. 2. 3.
        'arabicParenR': 'decimal',              // 1) 2) 3)
        'arabicParenBoth': 'decimal',           // (1) (2) (3)
        'arabic1Minus': 'decimal',              // 1- 2- 3-
        'arabic2Minus': 'decimal',              // -1- -2- -3-
        'arabicDbPeriod': 'decimal',            // 1.. 2.. 3..
        'arabicDbPlain': 'decimal',             // 1 2 3

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


function getInheritedPropertiesFromShape(shapeNode, level = 0) {
    if (!shapeNode) return null;

    try {
        // Navigate: shapeNode -> p:txBody -> a:lstStyle -> a:lvl1pPr (or lvl2pPr, etc.) -> a:defRPr
        const txBody = shapeNode?.["p:txBody"]?.[0];
        if (!txBody) return null;

        const lstStyle = txBody?.["a:lstStyle"]?.[0];
        if (!lstStyle) return null;

        // Try to get the level from the first paragraph if available
        const paragraphs = txBody?.["a:p"];
        if (paragraphs && paragraphs.length > 0) {
            const firstPPr = paragraphs[0]?.["a:pPr"]?.[0];
            if (firstPPr?.["$"]?.lvl !== undefined) {
                level = parseInt(firstPPr["$"].lvl);
            }
        }

        // Get the appropriate level properties (lvl1pPr, lvl2pPr, etc.)
        const levelKey = `a:lvl${level + 1}pPr`;
        const levelPr = lstStyle?.[levelKey]?.[0];

        if (!levelPr) {
            // Try level 0 as fallback
            const fallbackLevelPr = lstStyle?.["a:lvl1pPr"]?.[0];
            if (!fallbackLevelPr) return null;

            const fallbackDefRPr = fallbackLevelPr?.["a:defRPr"]?.[0];
            if (!fallbackDefRPr) return null;

            return {
                fonts: {
                    latin: fallbackDefRPr?.["a:latin"]?.[0]?.["$"]?.typeface || null,
                    ea: fallbackDefRPr?.["a:ea"]?.[0]?.["$"]?.typeface || null,
                    cs: fallbackDefRPr?.["a:cs"]?.[0]?.["$"]?.typeface || null,
                    sym: fallbackDefRPr?.["a:sym"]?.[0]?.["$"]?.typeface || null
                },
                defRPr: fallbackDefRPr // Return entire defRPr for other properties
            };
        }

        // Get defRPr (default run properties)
        const defRPr = levelPr?.["a:defRPr"]?.[0];
        if (!defRPr) return null;

        // Extract fonts from defRPr and return defRPr itself
        return {
            fonts: {
                latin: defRPr?.["a:latin"]?.[0]?.["$"]?.typeface || null,
                ea: defRPr?.["a:ea"]?.[0]?.["$"]?.typeface || null,
                cs: defRPr?.["a:cs"]?.[0]?.["$"]?.typeface || null,
                sym: defRPr?.["a:sym"]?.[0]?.["$"]?.typeface || null
            },
            defRPr: defRPr // Return entire defRPr for cap and other properties
        };
    } catch (error) {
        console.warn("Error getting inherited properties from shape:", error);
        return null;
    }
}

// NEW: Helper function to resolve master color mapping
function resolveMasterColor(schemeColor, masterXML) {
    try {
        // Look for the color map in master slide
        const colorMap = masterXML?.["p:sldMaster"]?.["p:clrMap"]?.[0]?.["$"];

        if (!colorMap) {
            return null;
        }

        // Check if the scheme color exists in the color map
        if (colorMap[schemeColor]) {
            return colorMap[schemeColor];
        }

        return null;
    } catch (error) {
        console.error("Error resolving master color:", error);
        return null;
    }
}

// IMPROVED: Better Wingdings character handling
function getWingdingsStyle(bulletChar) {
    const charCode = bulletChar.charCodeAt(0);

    const wingdingsMap = {
        // Common Wingdings characters by character code
        252: 'checkmark',
        251: 'x-mark',
        159: 'circle',
        110: 'square',
        108: 'diamond',
        224: 'arrow',
        232: 'arrow',
        171: 'star',
        118: 'arrow',
        94: 'arrow',
        169: 'heart',
        74: 'smiley',
        40: 'phone',
        42: 'envelope'
    };

    // Also check by actual character for common ones
    const wingdingsChars = {
        'ü': 'checkmark',
        'û': 'x-mark',
        'n': 'square',
        'l': 'diamond',
        'v': 'down-arrow',
        '^': 'up-arrow'
    };

    return wingdingsChars[bulletChar] || wingdingsMap[charCode] || 'disc';
}

// IMPROVED: Better Symbol font handling
function getSymbolFontStyle(bulletChar) {
    const charCode = bulletChar.charCodeAt(0);

    const symbolMap = {
        // Symbol font character codes
        183: 'disc',         // · Symbol bullet
        176: 'circle',       // ° Symbol degree (used as hollow bullet)
        254: 'square',       // þ Symbol square
        68: 'triangle',      // D Symbol triangle
        168: 'diamond',      // ¨ Symbol diamond
        72: 'star'           // H Symbol star
    };

    return symbolMap[charCode] || 'disc';
}


function hexToRGB(hex) {
    let r = parseInt(hex.slice(1, 3), 16);
    let g = parseInt(hex.slice(3, 5), 16);
    let b = parseInt(hex.slice(5, 7), 16);
    return `${r}, ${g}, ${b}`;
}




// UPDATED: Extract shape border with gradient support
function extractTextOutline(shapeNode, themeXML, clrMap) {
    try {
        let outlineInfo = {
            width: 0,
            color: '#000000',
            style: 'solid',
            css: '',
            type: 'shape',
            borderRadius: 0,
            hasClipPath: false  // NEW: Track if shape needs clip-path adjustment
        };

        // Extract border radius from geometry
        const spPr = shapeNode?.["p:spPr"]?.[0];
        const prstGeom = spPr?.["a:prstGeom"]?.[0];
        const geomType = prstGeom?.["$"]?.prst;

        if (geomType === 'roundRect' || geomType === 'round1Rect' || geomType === 'round2SameRect' || geomType === 'round2DiagRect') {
            const avLst = prstGeom?.["a:avLst"]?.[0];
            const adj = avLst?.["a:gd"]?.[0]?.["$"]?.fmla;

            if (adj) {
                const adjValue = parseInt(adj.replace('val ', ''));
                outlineInfo.borderRadius = Math.max(5, Math.round(adjValue / 3500));
            } else {
                outlineInfo.borderRadius = 10;
            }
            outlineInfo.hasClipPath = true; // Shape has rounded corners
        }

        // Method 1: Check for direct line in p:spPr -> a:ln
        const lnNode = spPr?.["a:ln"]?.[0];

        if (lnNode) {
            // Check if line is explicitly set to noFill
            const noFill = lnNode?.["a:noFill"];
            if (noFill) {
                // Shape explicitly has no border
                return null;
            }

            // Extract line width (in EMUs)
            const widthEMU = lnNode?.["$"]?.w;
            if (widthEMU) {
                outlineInfo.width = Math.round(parseInt(widthEMU) / getEMUDivisor());
            } else {
                outlineInfo.width = 1;
            }

            // Check for gradient fill first
            const gradFill = lnNode?.["a:gradFill"]?.[0];
            // console.log("gradFill=======",gradFill);

            if (gradFill) {
                // Extract gradient stops
                const gsLst = gradFill?.["a:gsLst"]?.[0];
                const gradientStops = gsLst?.["a:gs"];

                if (gradientStops && gradientStops.length > 0) {
                    // Get the last gradient stop (end color) - usually the accent color
                    const lastStop = gradientStops[gradientStops.length - 1];
                    const schemeClr = lastStop?.["a:schemeClr"]?.[0];

                    if (schemeClr) {
                        const schemeVal = schemeClr["$"]?.val;
                        console.log(`Scheme ${schemeVal} resolved to:`, colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null));
                        outlineInfo.color = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null);

                        // Apply any color modifications
                        const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
                        const lumOff = schemeClr["a:lumOff"]?.[0]?.["$"]?.val;
                        if (lumMod && lumOff) {
                            outlineInfo.color = pptBackgroundColors.applyLuminanceModifier(outlineInfo.color, lumMod, lumOff);
                        } else if (lumMod) {
                            outlineInfo.color = colorHelper.applyLumMod(outlineInfo.color, lumMod);
                        }
                    }
                }
            }
            // Check for solid fill
            else {
                const solidFill = lnNode?.["a:solidFill"]?.[0];
                if (solidFill) {
                    outlineInfo.color = extractColorFromNode(solidFill, themeXML);
                }
            }

            // Extract line style
            const prstDash = lnNode?.["a:prstDash"]?.[0]?.["$"]?.val;
            outlineInfo.style = getDashStyle(prstDash);

            if (outlineInfo.width > 0) {
                // NEW: Adjust clip-path to show border
                if (outlineInfo.hasClipPath) {
                    // Negative inset to prevent clip-path from cutting off the border
                    const inset = -outlineInfo.width;
                    console.log("outlineInfo.borderRadius", outlineInfo.borderRadius);
                    outlineInfo.css = `
                        border: ${outlineInfo.width}px ${outlineInfo.style} ${outlineInfo.color}; 
                        border-radius: ${outlineInfo.borderRadius}px;
                        clip-path: inset(${inset}px round ${outlineInfo.borderRadius}px) !important;
                    `.replace(/\s+/g, ' ').trim();
                } else {
                    outlineInfo.css = `border: ${outlineInfo.width}px ${outlineInfo.style} ${outlineInfo.color}; border-radius: ${outlineInfo.borderRadius}px;`;
                }
            }

            if (isTextBox(shapeNode)) {
                return null;
            }

            return outlineInfo;
        }

        // Method 2: Check for line reference in p:style -> a:lnRef
        const styleNode = shapeNode?.["p:style"]?.[0];
        const lnRef = styleNode?.["a:lnRef"]?.[0];

        if (lnRef) {
            const idx = lnRef?.["$"]?.idx;

            // If idx is "0", it means NO border
            if (idx === "0" || idx === 0) {
                return null;
            }

            const schemeClr = lnRef?.["a:schemeClr"]?.[0];
            if (schemeClr) {
                const schemeVal = schemeClr["$"]?.val;
                console.log(`Scheme ${schemeVal} resolved to:`, colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null));
                outlineInfo.color = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null);

                const shade = schemeClr["a:shade"]?.[0]?.["$"]?.val;
                if (shade) {
                    outlineInfo.color = applyShadeOrTint(outlineInfo.color, shade, 'shade');
                }

                const tint = schemeClr["a:tint"]?.[0]?.["$"]?.val;
                if (tint) {
                    outlineInfo.color = applyShadeOrTint(outlineInfo.color, tint, 'tint');
                }

                const lumMod = schemeClr["a:lumMod"]?.[0]?.["$"]?.val;
                const lumOff = schemeClr["a:lumOff"]?.[0]?.["$"]?.val;
                if (lumMod && lumOff) {
                    outlineInfo.color = pptBackgroundColors.applyLuminanceModifier(outlineInfo.color, lumMod, lumOff);
                } else if (lumMod) {
                    outlineInfo.color = colorHelper.applyLumMod(outlineInfo.color, lumMod);
                }
            }

            const lineWidths = {
                '0': 0,
                '1': 1,
                '2': 2,
                '3': 3
            };

            outlineInfo.width = lineWidths[idx] || 2;

            if (outlineInfo.width > 0) {
                // NEW: Adjust clip-path to show border
                if (outlineInfo.hasClipPath) {
                    // Negative inset to prevent clip-path from cutting off the border
                    const inset = -outlineInfo.width;
                    outlineInfo.css = `
                        border: ${outlineInfo.width}px ${outlineInfo.style} ${outlineInfo.color}; 
                        border-radius: ${outlineInfo.borderRadius}px;
                        clip-path: inset(${inset}px round ${outlineInfo.borderRadius}px) !important;
                    `.replace(/\s+/g, ' ').trim();
                } else {
                    outlineInfo.css = `border: ${outlineInfo.width}px ${outlineInfo.style} ${outlineInfo.color}; border-radius: ${outlineInfo.borderRadius}px;`;
                }
            }

            if (isTextBox(shapeNode)) {
                return null;
            }

            return outlineInfo;
        }

        return null;

    } catch (error) {
        console.error("Error extracting text outline:", error);
        return null;
    }
}

function isTextBox(shapeNode) {
    try {
        // Check if shape has nvSpPr (non-visual shape properties)
        const nvSpPr = shapeNode?.["p:nvSpPr"]?.[0];

        // Check for ph (placeholder) - text placeholders shouldn't have borders
        const nvPr = nvSpPr?.["p:nvPr"]?.[0];
        const ph = nvPr?.["p:ph"]?.[0];

        if (ph) {
            const phType = ph?.["$"]?.type;
            // Only filter out pure text placeholder types
            if (phType === 'body' || phType === 'title' || phType === 'subTitle' || phType === 'ctrTitle') {
                return true;
            }
        }

        // NEW LOGIC: Check if it's a text box BUT also check if it has an explicit border
        const cNvSpPr = nvSpPr?.["p:cNvSpPr"]?.[0];
        const txBox = cNvSpPr?.["$"]?.txBox;

        if (txBox === "1") {
            // Check if the shape has an explicit border defined
            const spPr = shapeNode?.["p:spPr"]?.[0];
            const lnNode = spPr?.["a:ln"]?.[0];

            // If it has a border defined with fill, DON'T filter it out
            if (lnNode) {
                const hasFill = lnNode?.["a:solidFill"] ||
                    lnNode?.["a:gradFill"] ||
                    lnNode?.["a:pattFill"];

                // If border has fill, keep the border (return false)
                if (hasFill) {
                    return false;
                }

                // If border is explicitly set to no fill, filter it out
                const noFill = lnNode?.["a:noFill"];
                if (noFill) {
                    return true;
                }
            }

            // If txBox="1" but no border info, filter it out
            return true;
        }

        return false;
    } catch (error) {
        console.error("Error checking if text box:", error);
        return false;
    }
}

function applyShadeOrTint(hexColor, value, type) {
    const hex = hexColor.replace('#', '');

    let r = parseInt(hex.substring(0, 2), 16);
    let g = parseInt(hex.substring(2, 4), 16);
    let b = parseInt(hex.substring(4, 6), 16);

    const factor = parseInt(value) / 100000;

    if (type === 'shade') {
        r = Math.round(r * factor);
        g = Math.round(g * factor);
        b = Math.round(b * factor);
    } else if (type === 'tint') {
        r = Math.round(r + (255 - r) * factor);
        g = Math.round(g + (255 - g) * factor);
        b = Math.round(b + (255 - b) * factor);
    }

    r = Math.max(0, Math.min(255, r));
    g = Math.max(0, Math.min(255, g));
    b = Math.max(0, Math.min(255, b));

    const toHex = (n) => {
        const hex = n.toString(16);
        return hex.length === 1 ? '0' + hex : hex;
    };

    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}


// UPDATED: Better color extraction with proper scheme color handling
function extractColorFromNode(solidFill, themeXML) {
    // Handle direct RGB color
    if (solidFill["a:srgbClr"]) {
        const srgbNode = solidFill["a:srgbClr"][0];
        let color = `#${srgbNode["$"].val}`;

        const lumMod = srgbNode["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = srgbNode["a:lumOff"]?.[0]?.["$"]?.val;
        if (lumMod && lumOff) {
            color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
        } else if (lumMod) {
            color = colorHelper.applyLumMod(color, lumMod);
        }

        return color;
    }
    // Handle scheme color (like accent1, accent2, etc.)
    else if (solidFill["a:schemeClr"]) {
        const schemeNode = solidFill["a:schemeClr"][0];
        const schemeVal = schemeNode["$"].val;
        let color = colorHelper.resolveThemeColorHelper(schemeVal, themeXML, null);

        // Apply shade if present
        const shade = schemeNode["a:shade"]?.[0]?.["$"]?.val;
        if (shade) {
            color = applyShadeOrTint(color, shade, 'shade');
        }

        // Apply tint if present
        const tint = schemeNode["a:tint"]?.[0]?.["$"]?.val;
        if (tint) {
            color = applyShadeOrTint(color, tint, 'tint');
        }

        // Apply luminance modifications
        const lumMod = schemeNode["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = schemeNode["a:lumOff"]?.[0]?.["$"]?.val;
        if (lumMod && lumOff) {
            color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
        } else if (lumMod) {
            color = colorHelper.applyLumMod(color, lumMod);
        }

        return color;
    }
    // Handle preset color
    else if (solidFill["a:prstClr"]) {
        const prstNode = solidFill["a:prstClr"][0];
        const prstVal = prstNode["$"].val;
        let color = colorHelper.resolvePresetColor(prstVal);

        // Apply luminance modifications if present
        const lumMod = prstNode["a:lumMod"]?.[0]?.["$"]?.val;
        const lumOff = prstNode["a:lumOff"]?.[0]?.["$"]?.val;
        if (lumMod && lumOff) {
            color = pptBackgroundColors.applyLuminanceModifier(color, lumMod, lumOff);
        } else if (lumMod) {
            color = colorHelper.applyLumMod(color, lumMod);
        }

        return color;
    }

    return '#000000'; // Default fallback
}

function getInheritedParagraphMargin(shapeNode, layoutXML, masterXML, placeholderType, level) {
    const divisor = 12700;
    
    // Initialize return object with all four margins
    let margins = {
        left: null,
        top: null,
        right: null,
        bottom: null
    };
    
    // 🔧 FIX 1: Check shape's own lstStyle FIRST (before layout/master)
    const lstStyle = shapeNode?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
    if (lstStyle) {
        const levelKey = `a:lvl${level + 1}pPr`;
        const levelPr = lstStyle[levelKey]?.[0];
        const marL = levelPr?.["$"]?.marL;
        if (marL) {
            console.log(`Found marL in shape lstStyle: ${marL}`);
            margins.left = Math.round(parseInt(marL) / divisor);
            return margins; // Return with left margin from lstStyle
        }
    }
    
    // 🔧 FIX 2: For ctrTitle, check BOTH titleStyle AND bodyStyle in master
    if (masterXML && (placeholderType === "ctrTitle" || placeholderType === "title")) {
        const txStyles = masterXML?.["p:sldMaster"]?.[0]?.["p:txStyles"]?.[0];
        
        if (txStyles) {
            // Try titleStyle first
            const titleStyle = txStyles["p:titleStyle"]?.[0];
            if (titleStyle) {
                const levelKey = `a:lvl${level + 1}pPr`;
                const levelPr = titleStyle?.[levelKey]?.[0];
                const marL = levelPr?.["$"]?.marL;
                
                if (marL) {
                    // console.log(`Found marL in master titleStyle: ${marL}`);
                    margins.left = Math.round(parseInt(marL) / divisor);
                    return margins;
                }
            }
            
            // 🔧 NEW: Also check bodyStyle for ctrTitle (some templates store it there)
            const bodyStyle = txStyles["p:bodyStyle"]?.[0];
            if (bodyStyle) {
                const levelKey = `a:lvl${level + 1}pPr`;
                const levelPr = bodyStyle?.[levelKey]?.[0];
                const marL = levelPr?.["$"]?.marL;
                
                if (marL) {
                    // console.log(`Found marL in master bodyStyle: ${marL}`);
                    margins.left = Math.round(parseInt(marL) / divisor);
                    return margins;
                }
            }
        }
    }
    
    // Try layout (for non-title placeholders)
    if (placeholderType && layoutXML) {
        const layoutShape = findPlaceholderInLayout(placeholderType, layoutXML);
        if (layoutShape) {
            const layoutLstStyle = layoutShape?.["p:txBody"]?.[0]?.["a:lstStyle"]?.[0];
            if (layoutLstStyle) {
                const levelKey = `a:lvl${level + 1}pPr`;
                const levelPr = layoutLstStyle[levelKey]?.[0];
                const marL = levelPr?.["$"]?.marL;
                if (marL) {
                    // console.log(`Found marL in layout: ${marL}`);
                    margins.left = Math.round(parseInt(marL) / divisor);
                    return margins;
                }
            }
        }
    }
    
    // 🔧 NEW FIX: Use bodyPr margins (lIns, tIns, rIns, bIns) as fallback
    const bodyPr = shapeNode?.["p:txBody"]?.[0]?.["a:bodyPr"]?.[0];
    if (bodyPr?.["$"]) {
        const attrs = bodyPr["$"];
        
        if (attrs.lIns) {
            margins.left = Math.round(parseInt(attrs.lIns) / divisor);
            // console.log(`Using lIns as fallback margin: ${attrs.lIns} -> ${margins.left}px`);
        }
        
        if (attrs.tIns) {
            margins.top = Math.round(parseInt(attrs.tIns) / divisor);
            // console.log(`Using tIns as fallback margin: ${attrs.tIns} -> ${margins.top}px`);
        }
        
        if (attrs.rIns) {
            margins.right = Math.round(parseInt(attrs.rIns) / divisor);
            // console.log(`Using rIns as fallback margin: ${attrs.rIns} -> ${margins.right}px`);
        }
        
        if (attrs.bIns) {
            margins.bottom = Math.round(parseInt(attrs.bIns) / divisor);
            // console.log(`Using bIns as fallback margin: ${attrs.bIns} -> ${margins.bottom}px`);
        }
        
        // Return if any margin was found
        if (margins.left !== null || margins.top !== null || margins.right !== null || margins.bottom !== null) {
            return margins;
        }
    }
    
    console.log(`No marL or bodyPr margins found for placeholder type: ${placeholderType}, level: ${level}`);
    return null;
}

function getDashStyle(prstDash) {
    const dashStyles = {
        'solid': 'solid',
        'dot': 'dotted',
        'dash': 'dashed',
        'lgDash': 'dashed',
        'dashDot': 'dashed',
        'lgDashDot': 'dashed',
        'lgDashDotDot': 'dashed',
        'sysDot': 'dotted',
        'sysDash': 'dashed',
        'sysDashDot': 'dashed',
        'sysDashDotDot': 'dashed'
    };
    return dashStyles[prstDash] || 'solid';
}


module.exports = {
    getAllTextInformationFromShape
};;