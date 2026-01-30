function addTextBoxToSlide(pptSlide, textBox, shapeTxtStyle, slideContext = null) {
    textBox = preprocessTextFormatting(textBox);
    const objName = textBox.getAttribute("data-name");

    if (!textBox) return;

    const textContent = textBox.textContent || "";
    if (!textContent.trim()) return;

    const style = textBox.style;
    const parentStyle = shapeTxtStyle.style;

    // Extract position and size from the parent shape (convert px to inches)
    let x = normalizeStyleValue(parentStyle.left) / 72;
    let y = normalizeStyleValue(parentStyle.top) / 72;
    let w = normalizeStyleValue(parentStyle.width) / 72;
    let h = normalizeStyleValue(parentStyle.height) / 72;

    // Apply scaling if slideContext is available
    if (slideContext && slideContext.scaleX && slideContext.scaleY) {
        x *= slideContext.scaleX;
        y *= slideContext.scaleY;
        w *= slideContext.scaleX;
        h *= slideContext.scaleY;
    }

    w = w * 0.995;

    const textAlign = style.textAlign || "left";

    // Extract rotation
    const parentTransform = parentStyle.transform || "";
    const textTransform = style.transform || "";
    let rotation = 0;
    const textRotationMatch = textTransform.match ? textTransform.match(/rotate\(([-\d.]+)deg\)/) : null;
    if (textRotationMatch) {
        rotation = parseFloat(textRotationMatch[1]);
    } else {
        const parentRotationMatch = parentTransform.match ? parentTransform.match(/rotate\(([-\d.]+)deg\)/) : null;
        if (parentRotationMatch) {
            rotation = parseFloat(parentRotationMatch[1]);
        }
    }

    // ========== MARGIN EXTRACTION AND CONVERSION FUNCTIONALITY ==========

    // const pixelsToEmus = (pixels) => {
    //     return Math.round(pixels * 12700); // 72 DPI: 1px = 12,700 EMUs
    // };

    const pixelsToEmus = (pixels) => {
        return pixels; // Convert pixels to inches (72 DPI)
    };


    const extractMarginValues = (elementStyle) => {
        const margins = { top: 0, right: 0, bottom: 0, left: 0 };

        if (!elementStyle) return margins;

        // Try to get individual margin properties first
        if (elementStyle.marginTop) margins.top = parseFloat(elementStyle.marginTop) || 0;
        if (elementStyle.marginRight) margins.right = parseFloat(elementStyle.marginRight) || 0;
        if (elementStyle.marginBottom) margins.bottom = parseFloat(elementStyle.marginBottom) || 0;
        if (elementStyle.marginLeft) margins.left = parseFloat(elementStyle.marginLeft) || 0;

        // Handle shorthand margin property
        if (elementStyle.margin) {

            const marginValue = elementStyle.margin.toString().trim().replace(/;+$/, ''); // Remove trailing semicolons
            const marginParts = marginValue.split(/\s+/).filter(part => part.length > 0);

            if (marginParts.length === 1) {
                // margin: 10px (all sides)
                const value = parseFloat(marginParts[0]) || 0;
                margins.top = margins.right = margins.bottom = margins.left = value;
            } else if (marginParts.length === 2) {

                // margin: 10px 20px (vertical horizontal)
                const vertical = parseFloat(marginParts[0]) || 0;
                const horizontal = parseFloat(marginParts[1]) || 0;
                margins.top = margins.bottom = vertical;
                margins.right = margins.left = horizontal;
            } else if (marginParts.length === 3) {

                // margin: 10px 20px 30px (top horizontal bottom)
                margins.top = parseFloat(marginParts[0]) || 0;
                margins.right = margins.left = parseFloat(marginParts[1]) || 0;
                margins.bottom = parseFloat(marginParts[2]) || 0;
            } else if (marginParts.length === 4) {

                // margin: 10px 20px 30px 40px (top right bottom left)
                margins.top = parseFloat(marginParts[0]) || 0;
                margins.right = parseFloat(marginParts[1]) || 0;
                margins.bottom = parseFloat(marginParts[2]) || 0;
                margins.left = parseFloat(marginParts[3]) || 0;
            }
        }

        return margins;
    };

    // FIXED: Initialize margin variables - will extract later from paragraphs
    let finalMargins = { top: 0, right: 0, bottom: 0, left: 0 };

    // FIXED: Extract global line spacing from the first span (where it actually exists in your HTML)
    const firstParagraph = textBox.querySelector('p');
    let globalLineSpacing = null;
    let globalFontSize = 14; // Default font size

    if (firstParagraph) {
        // Try to get line spacing from first span (where it actually exists in your HTML)
        const firstSpan = firstParagraph.querySelector('span');
        let lineHeightSource = null;

        if (firstSpan) {
            // Extract from span style attribute or computed style
            const spanStyle = firstSpan.getAttribute('style') || '';
            lineHeightSource = extractLineHeightFromCSS(spanStyle);

            // Also get font size from the same span
            if (firstSpan.style.fontSize) {
                globalFontSize = parseFloat(firstSpan.style.fontSize);
            }
        }

        // Fallback to paragraph level if not found in span
        if (!lineHeightSource) {
            const paragraphStyle = firstParagraph.style || {};
            lineHeightSource = extractLineHeightFromCSS(paragraphStyle.cssText || firstParagraph.getAttribute('style'));

            if (paragraphStyle.fontSize) {
                globalFontSize = parseFloat(paragraphStyle.fontSize);
            }
        }

        if (lineHeightSource) {
            globalLineSpacing = convertLineHeightToPowerPoint(lineHeightSource, globalFontSize);
        }
    }

    // Continue with normal paragraph processing...
    const paragraphs = textBox.querySelectorAll ? Array.from(textBox.querySelectorAll("p")) : [];

    // FIXED: Extract margins from paragraphs (where they actually exist in your HTML)
    if (paragraphs.length > 0) {
        // Use the first paragraph's margin as the text box margin
        const firstParagraph = paragraphs[0];
        if (firstParagraph && firstParagraph.style) {
            finalMargins = extractMarginValues(firstParagraph.style);
        }
    }

    // Fallback: Check textBox style and parent style if no paragraph margins found
    if (finalMargins.top === 0 && finalMargins.right === 0 && finalMargins.bottom === 0 && finalMargins.left === 0) {
        const textBoxMargins = extractMarginValues(style);
        const parentMargins = extractMarginValues(parentStyle);

        finalMargins = {
            top: textBoxMargins.top || parentMargins.top,
            right: textBoxMargins.right || parentMargins.right,
            bottom: textBoxMargins.bottom || parentMargins.bottom,
            left: textBoxMargins.left || parentMargins.left
        };
    }

    const marginInches = {
        top: pixelsToEmus(finalMargins.left),
        left: pixelsToEmus(finalMargins.top),
        right: pixelsToEmus(finalMargins.right),
        bottom: pixelsToEmus(finalMargins.bottom)
    };

    // Detect vertical alignment from flexbox properties
    const detectVerticalAlign = (textBoxStyle, parentShapeStyle) => {
        const justifyContent = textBoxStyle.justifyContent || textBoxStyle['justify-content'] || '';
        const alignItems = parentShapeStyle.alignItems || parentShapeStyle['align-items'] || '';
        const display = textBoxStyle.display || '';
        const flexDirection = textBoxStyle.flexDirection || textBoxStyle['flex-direction'] || '';

        // Priority 1: Check textBox's justify-content (if it's a flex column container)
        if (display === 'flex' && flexDirection === 'column') {

            if (justifyContent === 'center') return 'middle';
            if (justifyContent === 'flex-end' || justifyContent === 'end') return 'bottom';
            if (justifyContent === 'flex-start' || justifyContent === 'start') return 'center';
        }

        // Priority 2: Check parent's align-items (common pattern in your HTML)
        if (alignItems === 'center') return 'middle';
        if (alignItems === 'flex-end' || alignItems === 'end') return 'bottom';
        if (alignItems === 'flex-start' || alignItems === 'start') return 'top';

        // Priority 3: Default to top if no alignment detected
        return 'top';
    };

    // In the addTextBoxToSlide function, use the following logic to detect vertical alignment
    const verticalAlign = detectVerticalAlign(style, parentStyle);

    // Text box options with margin support
    const textBoxOptions = {
        x: x,
        y: y,
        w: w,
        h: h,
        transparent: 100,
        autofit: false,
        valign: verticalAlign,
        objectName: objName || '',
    };

    textBoxOptions.margin = [marginInches.top, marginInches.right, marginInches.bottom, marginInches.left];

    // FIXED: Add global line spacing with correct property name
    if (globalLineSpacing) {
        textBoxOptions.lineSpacing = globalLineSpacing;
    }

    // Add rotation if found
    if (rotation !== 0) {
        textBoxOptions.rotate = rotation;
    }

    // ENHANCED: Check for lists with better detection and debugging
    const hasLists = textBox.querySelector && (
        textBox.querySelector("ul") !== null ||
        textBox.querySelector("ol") !== null
    );

    // FIXED: Check if content is ONLY lists (no paragraphs) or mixed content
    const hasParagraphs = paragraphs.length > 0;
    const hasMixedContent = hasLists && hasParagraphs;

    // If there's mixed content (lists + paragraphs), we need to handle them in order
    if (hasMixedContent) {
        // DEBUG: Analyze list structure before processing
        debugListStructure(textBox);

        // Process mixed content in order
        handleMixedContent(pptSlide, textBox, textBoxOptions, globalLineSpacing, globalFontSize, textAlign);
        return;
    }

    // If ONLY lists (no paragraphs), use the original list handler
    if (hasLists && !hasParagraphs) {
        // DEBUG: Analyze list structure before processing
        debugListStructure(textBox);

        // FIXED: Pass global line spacing to list handler
        handleListContent(pptSlide, textBox, textBoxOptions, globalLineSpacing, globalFontSize);
        return;
    }


    const processedParagraphs = [];

    paragraphs.forEach((p, index) => {

        console.log(`\n=== Processing paragraph ${index} ===`);
        console.log("HTML:", p.innerHTML);
        console.log("Text:", p.textContent);
        const paragraphHTML = p.innerHTML.trim();

        const isBreakOnlyParagraph = paragraphHTML === '<br>' ||
            paragraphHTML === '<br/>' ||
            paragraphHTML === '<br />' ||
            paragraphHTML === '&nbsp;' ||
            paragraphHTML === '&#160;' ||
            (paragraphHTML.replace(/<br\s*\/?>/gi, '').trim() === '') ||
            (paragraphHTML.replace(/&nbsp;/gi, '').replace(/&#160;/gi, '').trim() === '');

        const hasMinHeight = p.style.minHeight && parseFloat(p.style.minHeight) > 0;
        const textContent = p.textContent || p.innerText || '';
        const isEmptyParagraph = isBreakOnlyParagraph ||
            (textContent.replace(/\u00A0/g, '').replace(/\s/g, '') === '' && hasMinHeight);

        if (isEmptyParagraph) {
            if (processedParagraphs.length > 0) {
                const lastParagraph = processedParagraphs[processedParagraphs.length - 1];
                if (lastParagraph && lastParagraph.length > 0) {
                    const lastRun = lastParagraph[lastParagraph.length - 1];
                    lastRun.text += '\n';
                }
            } else {
                if (hasMinHeight) {
                    const emptyParagraphFontSize = extractFontSizeFromParagraph(p) || globalFontSize;
                    const emptyParagraphAlign = p.style.textAlign || textAlign;

                    processedParagraphs.push([{
                        text: '\n',
                        options: {
                            fontFace: 'Arial',
                            fontSize: emptyParagraphFontSize,
                            color: '#000000',
                            align: emptyParagraphAlign
                        }
                    }]);
                }
            }
                } else {
            // Process non-empty paragraphs
            const paragraphText = processSpansInParagraphWithLineSpacing(p);

            if (paragraphText && paragraphText.length > 0) {
                // Instead of adding a separate "\n" run, mark the *previous*
                // paragraph's last run with breakLine = true.
                if (processedParagraphs.length > 0) {
                    const prevParagraphRuns = processedParagraphs[processedParagraphs.length - 1];

                    if (Array.isArray(prevParagraphRuns) && prevParagraphRuns.length > 0) {
                        const lastRun = prevParagraphRuns[prevParagraphRuns.length - 1];

                        // Clone options and remove lineSpacing for the break
                        const prevOpts = { ...(lastRun.options || {}) };
                        delete prevOpts.lineSpacing;

                        lastRun.options = {
                            ...prevOpts,
                            breakLine: true  // << key flag for new <a:p>
                        };
                    }
                }

                // Push current paragraph runs as-is (each span is its own <a:r>)
                processedParagraphs.push(paragraphText);
            }
        }

    });

    if (processedParagraphs.length === 0 && textContent.trim()) {
        const fallbackText = processFallbackText(textBox, style);

        if (fallbackText && fallbackText.length > 0) {
            processedParagraphs.push(fallbackText);
        }
    }

    if (processedParagraphs.length > 0) {
        try {

            textBoxOptions.name = "Sample Title";
            const flattenedText = processedParagraphs.flat();
            console.log("textBoxOptions===",textBoxOptions);
            console.log("flattenedText=======",flattenedText);
            pptSlide.addText(flattenedText, textBoxOptions);
        } catch (error) {
            console.error("Error adding text to slide:", error);
            const fallbackText = processedParagraphs.flat().map(t => ({
                ...t,
                options: { ...t.options, fontFace: "Arial" }
            }));
            try {

                pptSlide.addText(fallbackText, textBoxOptions);
            } catch (fallbackError) {
                console.error("Fallback error adding text to slide:", fallbackError);
            }
        }
    }
}

/**-------------- */

function processSpansInParagraphWithLineSpacing(paragraph) {
    const textRuns = [];
    const pStyle = paragraph.style || {};
    const defaultAlign = pStyle.textAlign || "left";

    // Rakesh Notes::: i have added this line
    const paragraphSpacing = extractParagraphSpacing(paragraph);
    console.log("Paragraph spacing:", paragraphSpacing); // DEBUG
    const spans = paragraph.querySelectorAll('span.span-txt, span');

    if (spans.length > 0) {
        spans.forEach((span, index) => {
            let spanText = span.textContent || span.innerText || "";

            const originalTxtColor = span.getAttribute('originaltxtcolor');
            const originalLumMod = span.getAttribute('originallummod');
            const originalLumOff = span.getAttribute('originallumoff');

            if (spanText.length > 0) {

                let nextSibling = span.nextSibling;
                let hasBrAfter = false;

                while (nextSibling) {
                    if (nextSibling.nodeType === 1 && nextSibling.tagName === 'BR') {
                        hasBrAfter = true;
                        break;
                    } else if (nextSibling.nodeType === 3) {
                        const textContent = nextSibling.textContent || "";
                        if (textContent.trim()) {
                            break;
                        }
                    } else if (nextSibling.nodeType === 1) {
                        break;
                    }
                    nextSibling = nextSibling.nextSibling;
                }

                if (hasBrAfter) {
                    spanText += '\n';
                }

                // FIXED: Extract line spacing from span
                const spanOptions = extractSpanFormattingWithLineSpacing(span, defaultAlign, originalTxtColor, originalLumMod, originalLumOff);
                
                // Rakesh Notes::: ADD PARAGRAPH SPACING HERE: this if condiotion are added for spacing
                if (paragraphSpacing.spaceBefore > 0) {
                    spanOptions.paraSpaceBefore = paragraphSpacing.spaceBefore;
                }
                if (paragraphSpacing.spaceAfter > 0) {
                    spanOptions.paraSpaceAfter = paragraphSpacing.spaceAfter;
                }

                textRuns.push({
                    text: spanText, // Keep original text including spaces
                    options: spanOptions
                });
            }
        });
    } else {
        // Handle paragraphs without spans
        let paragraphText = paragraph.textContent || paragraph.innerText || "";

        if (paragraph.innerHTML && paragraph.innerHTML.includes('<br')) {
            paragraphText = paragraph.innerHTML.replace(/<br\s*\/?>/gi, '\n');
            paragraphText = paragraphText.replace(/<[^>]*>/g, '');
        }

        // FIXED: Don't process paragraphs that only contain &nbsp; or whitespace
        const cleanText = paragraphText.replace(/\u00A0/g, '').trim();
        if (cleanText || paragraphText.includes('\n')) {
            const paragraphOptions = extractParagraphFormattingWithLineSpacing(paragraph, defaultAlign);
    // ADD PARAGRAPH SPACING HERE:
    if (paragraphSpacing.spaceBefore > 0) {
        paragraphOptions.paraSpaceBefore = paragraphSpacing.spaceBefore;
    }
    if (paragraphSpacing.spaceAfter > 0) {
        paragraphOptions.paraSpaceAfter = paragraphSpacing.spaceAfter;
    }
            textRuns.push({
                text: paragraphText,
                options: paragraphOptions
            });
        }
    }

    return textRuns;
}

// Rakesh notes::: for pragraphspaceing logic
// Extract paragraph spacing from HTML
function extractParagraphSpacing(pElement) {
    const style = pElement.style;
    
    // Extract margin-top and margin-bottom (these are in pixels)
    const marginTop = parseFloat(style.marginTop) || 0;
    const marginBottom = parseFloat(style.marginBottom) || 0;
    const spaceBeforePts = Math.round(marginTop); // Keep as points
    const spaceAfterPts = Math.round(marginBottom); // Keep as points
    
    return {
        spaceBefore: spaceBeforePts,
        spaceAfter: spaceAfterPts
    };
}
/** ------------- */


/** --------- */
// NEW: Helper function to extract font size from paragraph styles
function extractFontSizeFromParagraph(paragraph) {
    const style = paragraph.style || {};
    const styleAttr = paragraph.getAttribute('style') || '';

    // Try to get font size from inline styles
    if (style.fontSize) {
        return parseFloat(style.fontSize);
    }

    // Try to extract from style attribute string
    const fontSizeMatch = styleAttr.match(/font-size:\s*([0-9.]+)px/i);
    if (fontSizeMatch) {
        return parseFloat(fontSizeMatch[1]);
    }

    return null;
}
/**------------- */


function extractParagraphFormattingWithLineSpacing(paragraph, defaultAlign = "left") {
    const style = paragraph.style || {};
    const paragraphStyleAttr = paragraph.getAttribute('style') || '';

    let fontSizePx = parseFloat(style.fontSize || "14");
    const fontSizePt = fontSizePx;

    const fontColor = rgbToHex(style.color || "#000000");

    let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
    if (!fontFamily) fontFamily = "Arial";

    const fontWeight = style.fontWeight || "normal";
    const isBold = fontWeight === "bold" || parseInt(fontWeight) >= 700;

    const fontStyle = style.fontStyle || "";
    const isItalic = fontStyle.includes("italic");

    let isUnderlined = false;
    if (style.textDecoration && style.textDecoration.includes('underline')) {
        isUnderlined = true;
    }

    const options = {
        fontFace: fontFamily,
        fontSize: fontSizePt,
        color: fontColor,
        align: defaultAlign,
        bold: isBold,
        italic: isItalic,
        underline: isUnderlined
    };

    // Extract line spacing from paragraph
    const lineHeightFromParagraph = extractLineHeightFromCSS(paragraphStyleAttr);
    if (lineHeightFromParagraph) {
        const lineSpacing = convertLineHeightToPowerPoint(lineHeightFromParagraph, fontSizePx);
        if (lineSpacing) {
            options.lineSpacing = lineSpacing;
        }
    }

    return options;
}


function extractSpanFormattingWithLineSpacing(span, defaultAlign = "left", originalTxtColor, originalLumMod, originalLumOff) {
    const style = span.style || {};
    const spanStyleAttr = span.getAttribute('style') || '';

    // Detect super/subscript and get correct font size
    const superSubInfo = detectSuperSubscriptSimple(span);

    let fontSizePx = superSubInfo.baseFontSize; // Use base font size, not reduced
    const fontSizePt = fontSizePx;
    let finalColor;

    // Helper function to check if a value is a valid scheme color
    function isValidSchemeColor(colorVal) {
        if (!colorVal || colorVal === 'undefined') return false;

        const validSchemeColors = [
            'tx1', 'tx2', 'bg1', 'bg2',
            'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
            'hlink', 'folHlink', 'dk1', 'dk2', 'lt1', 'lt2', 'phclr',
            'text1', 'text2', 'background1', 'background2'
        ];

        return validSchemeColors.includes(colorVal.toLowerCase());
    }

    // Helper function to convert common color names to hex
    function colorNameToHex(colorName) {
        const colorMap = {
            'black': '#000000',
            'white': '#FFFFFF',
            'red': '#FF0000',
            'green': '#008000',
            'blue': '#0000FF',
            'yellow': '#FFFF00',
            'cyan': '#00FFFF',
            'magenta': '#FF00FF',
            'gray': '#808080',
            'grey': '#808080'
        };
        return colorMap[colorName.toLowerCase()] || null;
    }

    // Helper function to check if string is a valid hex color
    function isValidHex(colorStr) {
        if (!colorStr) return false;
        const hex = colorStr.replace('#', '');
        return /^[0-9A-Fa-f]{6}$/.test(hex);
    }

    if (originalTxtColor && originalTxtColor !== 'undefined') {

        // Check if it's a valid scheme color
        if (isValidSchemeColor(originalTxtColor)) {
            // Create theme color object for valid scheme colors
            finalColor = {
                type: 'schemeClr',
                val: originalTxtColor.toLowerCase()
            };

            // Add luminance modifications if they exist and are valid
            if (originalLumMod && !isNaN(parseInt(originalLumMod))) {
                finalColor.lumMod = parseInt(originalLumMod);
            }
            if (originalLumOff && !isNaN(parseInt(originalLumOff))) {
                finalColor.lumOff = parseInt(originalLumOff);
            }
        }
        // Check if it's a hex color (like "000000")
        else if (isValidHex(originalTxtColor)) {
            // For hex colors, we need to handle luminance modifications differently
            // If there are luminance modifications, we should treat this as a scheme color
            if ((originalLumMod && !isNaN(parseInt(originalLumMod))) ||
                (originalLumOff && !isNaN(parseInt(originalLumOff)))) {

                // Convert hex to scheme color (this is a workaround - you might need to map hex to appropriate scheme colors)
                // For black (000000), we can use 'tx1' or 'dk1' as a reasonable scheme color
                let schemeColorVal = 'tx1'; // Default to text1
                if (originalTxtColor.toLowerCase() === '000000') {
                    schemeColorVal = 'tx1'; // Black text
                } else if (originalTxtColor.toLowerCase() === 'ffffff') {
                    schemeColorVal = 'bg1'; // White background
                }

                finalColor = {
                    type: 'schemeClr',
                    val: schemeColorVal
                };

                if (originalLumMod && !isNaN(parseInt(originalLumMod))) {
                    finalColor.lumMod = parseInt(originalLumMod);
                }
                if (originalLumOff && !isNaN(parseInt(originalLumOff))) {
                    finalColor.lumOff = parseInt(originalLumOff);
                }
            } else {
                // Pure hex color without luminance modifications
                finalColor = '#' + originalTxtColor;
            }
        }
        // Check if it's a color name (like "black")
        else {
            const hexFromName = colorNameToHex(originalTxtColor);
            if (hexFromName) {
                // Handle luminance modifications for color names too
                if ((originalLumMod && !isNaN(parseInt(originalLumMod))) ||
                    (originalLumOff && !isNaN(parseInt(originalLumOff)))) {

                    console.log("Color name with luminance modifications, treating as scheme color");

                    let schemeColorVal = 'tx1'; // Default
                    if (originalTxtColor.toLowerCase() === 'black') {
                        schemeColorVal = 'tx1';
                    } else if (originalTxtColor.toLowerCase() === 'white') {
                        schemeColorVal = 'bg1';
                    }

                    finalColor = {
                        type: 'schemeClr',
                        val: schemeColorVal
                    };

                    if (originalLumMod && !isNaN(parseInt(originalLumMod))) {
                        finalColor.lumMod = parseInt(originalLumMod);
                    }
                    if (originalLumOff && !isNaN(parseInt(originalLumOff))) {
                        finalColor.lumOff = parseInt(originalLumOff);
                    }
                } else {
                    finalColor = hexFromName;
                }
            } else {
                console.warn(`Unrecognized color value "${originalTxtColor}", falling back to default`);
                finalColor = rgbToHex(style.color || "#000000");
            }
        }
    } else {
        console.log("No originalTxtColor found, using default color");
        finalColor = rgbToHex(style.color || "#000000");
    }

    let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
    if (!fontFamily) fontFamily = "";

    const fontWeight = style.fontWeight || "normal";
    const isBold = fontWeight === "bold" || parseInt(fontWeight) >= 700;

    const fontStyle = style.fontStyle || "";
    const isItalic = fontStyle.includes("italic");

    let isUnderlined = false;
    if (style.textDecoration && style.textDecoration.includes('underline')) {
        isUnderlined = true;
    }

    const options = {
        fontSize: fontSizePt,
        color: finalColor,
        align: defaultAlign,
        bold: isBold,
        italic: isItalic,
        underline: isUnderlined
    };

    // 🔹 NEW: pick up original EA / CS / SYM from HTML attributes
    const originalEa = span.getAttribute('originalea');
    const originalCs = span.getAttribute('origincs');
    const originalSym = span.getAttribute('originsym');
    const latinFont = span.getAttribute('latinFont');

    if (originalEa) {
        options.fontFaceEa = originalEa;      // custom metadata, not used by pptxgenjs
    }
    if (originalCs) {
        options.fontFaceCs = originalCs;      // custom metadata, not used by pptxgenjs
    }
    if (originalSym) {
        options.fontFaceSym = originalSym;    // custom metadata, not used by pptxgenjs
    }
    if (latinFont) {
        options.fontFace = latinFont;    // custom metadata, not used by pptxgenjs
    }

    // Add superscript/subscript properties with correct baseline values
    if (superSubInfo.isSuperscript) {
        options.baseline = 300; // PptxGenJS expects smaller values (converts to OOXML)
    } else if (superSubInfo.isSubscript) {
        options.baseline = -250; // PptxGenJS expects smaller values (converts to OOXML)
    } else {
        // Original positioning logic for non-super/sub cases
        const position = style.position || "";
        const top = style.top || "";
        if (position === "relative" && top) {
            const topValue = parseFloat(top);
            if (topValue < 0) {
                options.baseline = 300;
            } else if (topValue > 0) {
                options.baseline = -250;
            }
        }
    }

    // Extract line spacing from span
    const lineHeightFromSpan = extractLineHeightFromCSS(spanStyleAttr);
    if (lineHeightFromSpan) {
        const lineSpacing = convertLineHeightToPowerPoint(lineHeightFromSpan, fontSizePx);
        if (lineSpacing) {
            options.lineSpacing = lineSpacing;
        }
    }

    return options;
}




// FIXED: Improved line height extraction function
function extractLineHeightFromCSS(cssText) {
    if (!cssText) return null;

    // Look for line-height property in CSS text
    const lineHeightMatch = cssText.match(/line-height:\s*([^;]+)/i);
    if (!lineHeightMatch) {
        return null;
    }

    const lineHeightValue = lineHeightMatch[1].trim();

    // Handle different line-height formats
    if (lineHeightValue === 'normal') {
        return 1.2; // Default normal line height
    }

    // Handle unitless numbers (1.5, 2.0, etc.)
    if (/^\d*\.?\d+$/.test(lineHeightValue)) {
        const value = parseFloat(lineHeightValue);
        return value;
    }

    // Handle percentages (150%, 200%, etc.)
    if (lineHeightValue.endsWith('%')) {
        const value = parseFloat(lineHeightValue) / 100;
        return value;
    }

    // Handle pixel values (24px, 30px, etc.)
    if (lineHeightValue.endsWith('px')) {
        const pixelValue = parseFloat(lineHeightValue);
        return { type: 'absolute', value: pixelValue };
    }

    // Handle em values (1.5em, 2em, etc.)
    if (lineHeightValue.endsWith('em')) {
        const value = parseFloat(lineHeightValue);
        return value;
    }

    // Handle pt values (18pt, 24pt, etc.)
    if (lineHeightValue.endsWith('pt')) {
        const ptValue = parseFloat(lineHeightValue);
        return { type: 'absolute', value: ptValue };
    }

    return null;
}

function convertLineHeightToPowerPoint(lineHeight, fontSize) {
    if (!lineHeight || !fontSize) {
        return null;
    }

    // Convert fontSize from px to pt (assuming 72 DPI where 1px = 1pt)
    const fontSizePt = fontSize;

    // Handle absolute values (px or pt)
    if (typeof lineHeight === 'object' && lineHeight.type === 'absolute') {
        const lineHeightPt = lineHeight.value; // Assuming 72 DPI: 1px = 1pt

        // Calculate the multiplier: lineHeight / fontSize
        const multiplier = lineHeightPt / fontSizePt;

        // For PptxGenJS, we can return the actual line height in points
        // or return the multiplier - check PptxGenJS documentation
        return Math.round(lineHeightPt * 100) / 100; // Return line height in points
    }

    // Handle relative values (unitless numbers, em, percentages)
    if (typeof lineHeight === 'number') {
        // For relative values, calculate the actual line height in points
        const lineHeightPt = fontSizePt * lineHeight;
        return Math.round(lineHeightPt * 100) / 100;
    }

    return null;
}


function processSpansInParagraphWithoutLineSpacing(paragraph) {
    const textRuns = [];
    const pStyle = paragraph.style || {};
    const defaultAlign = pStyle.textAlign || "left";

    const spans = paragraph.querySelectorAll('span.portion, span');

    if (spans.length > 0) {
        spans.forEach((span, index) => {
            let spanText = span.textContent || span.innerText || "";

            // FIXED: Don't filter out spans - preserve ALL spans including whitespace-only ones
            if (spanText.length > 0) {  // Changed from spanText.trim() to spanText.length > 0
                // Check if there's a <br> tag immediately after this span
                let nextSibling = span.nextSibling;
                let hasBrAfter = false;

                // Look through the next siblings to find a <br> tag
                while (nextSibling) {
                    if (nextSibling.nodeType === 1 && nextSibling.tagName === 'BR') {
                        hasBrAfter = true;
                        break;
                    } else if (nextSibling.nodeType === 3) {
                        // Text node - check if it's just whitespace
                        const textContent = nextSibling.textContent || "";
                        if (textContent.trim()) {
                            // Non-whitespace text node, stop looking
                            break;
                        }
                    } else if (nextSibling.nodeType === 1) {
                        // Another element (not BR), stop looking
                        break;
                    }
                    nextSibling = nextSibling.nextSibling;
                }

                // If there's a <br> after this span, add a line break to the text
                if (hasBrAfter) {
                    spanText += '\n';
                }

                const spanOptions = extractSpanFormattingWithoutLineSpacing(span, defaultAlign);
                textRuns.push({
                    text: spanText, // Keep original text including spaces
                    options: spanOptions
                });
            }
        });
    } else {
        // Handle paragraphs without spans
        let paragraphText = paragraph.textContent || paragraph.innerText || "";

        // Handle <br> tags within paragraphs by converting them to line breaks
        if (paragraph.innerHTML && paragraph.innerHTML.includes('<br')) {
            paragraphText = paragraph.innerHTML.replace(/<br\s*\/?>/gi, '\n');
            // Remove any remaining HTML tags
            paragraphText = paragraphText.replace(/<[^>]*>/g, '');
        }

        if (paragraphText.trim() || paragraphText.includes('\n')) {
            const paragraphOptions = extractParagraphFormattingWithoutLineSpacing(paragraph, defaultAlign);
            textRuns.push({
                text: paragraphText,
                options: paragraphOptions
            });
        }
    }

    return textRuns;
}

function extractSpanFormattingWithoutLineSpacing(span, defaultAlign = "left") {
    const style = span.style || {};

    // Detect super/subscript and get correct font size
    const superSubInfo = detectSuperSubscriptSimple(span);

    let fontSizePx = superSubInfo.baseFontSize; // Use base font size, not reduced
    const fontSizePt = fontSizePx;

    const fontColor = rgbToHex(style.color || "#000000");

    let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
    if (!fontFamily) fontFamily = "Arial";

    const fontWeight = style.fontWeight || "normal";
    const isBold = fontWeight === "bold" || parseInt(fontWeight) >= 700;

    const fontStyle = style.fontStyle || "";
    const isItalic = fontStyle.includes("italic");

    let isUnderlined = false;
    if (style.textDecoration && style.textDecoration.includes('underline')) {
        isUnderlined = true;
    }

    const options = {
        fontFace: fontFamily,
        fontSize: fontSizePt,
        color: fontColor,
        align: defaultAlign,
        bold: isBold,
        italic: isItalic,
        underline: isUnderlined
    };

    // Add superscript/subscript properties
    if (superSubInfo.isSuperscript) {
        options.baseline = 300; // PptxGenJS expects smaller values
    } else if (superSubInfo.isSubscript) {
        options.baseline = -250; // PptxGenJS expects smaller values
    } else {
        // Original positioning logic for non-super/sub cases
        const position = style.position || "";
        const top = style.top || "";
        if (position === "relative" && top) {
            const topValue = parseFloat(top);
            if (topValue < 0) {
                options.baseline = 300;
            } else if (topValue > 0) {
                options.baseline = -250;
            }
        }
    }

    return options;
}

function extractParagraphFormattingWithoutLineSpacing(paragraph, defaultAlign = "left") {
    const style = paragraph.style || {};

    let fontSizePx = parseFloat(style.fontSize || "14");
    const fontSizePt = fontSizePx; // 1:1 conversion for 72 DPI

    const fontColor = rgbToHex(style.color || "#000000");

    let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
    if (!fontFamily) fontFamily = "Arial";

    const fontWeight = style.fontWeight || "normal";
    const isBold = fontWeight === "bold" || parseInt(fontWeight) >= 700;

    const fontStyle = style.fontStyle || "";
    const isItalic = fontStyle.includes("italic");

    let isUnderlined = false;
    if (style.textDecoration && style.textDecoration.includes('underline')) {
        isUnderlined = true;
    }

    return {
        fontFace: fontFamily,
        fontSize: fontSizePt,
        color: fontColor,
        align: defaultAlign,
        bold: isBold,
        italic: isItalic,
        underline: isUnderlined
    };
}


function detectSuperSubscriptSimple(span) {
    const style = span.style || {};
    const position = style.position || "";
    const top = style.top || "";
    const fontSize = parseFloat(style.fontSize || 24);

    // Get parent container to find base font size
    let baseFontSize = 24;
    try {
        const parent = span.closest('p, div');
        if (parent) {
            const siblings = parent.querySelectorAll('span');
            const fontSizes = Array.from(siblings).map(s => parseFloat(s.style.fontSize || 24));
            baseFontSize = Math.max(...fontSizes.filter(s => s > 0));
        }
    } catch (e) {
        // Fallback to default
    }

    // Check if this is super/subscript
    if (position === "relative" && top) {
        const topValue = parseFloat(top);
        const fontRatio = fontSize / baseFontSize;

        // If font is smaller AND positioned vertically
        if (fontRatio < 0.9 && Math.abs(topValue) > 2) {
            if (topValue < 0) {
                return { isSuperscript: true, baseFontSize };
            } else {
                return { isSubscript: true, baseFontSize };
            }
        }
    }

    return { isSuperscript: false, isSubscript: false, baseFontSize: fontSize };
}


function processFallbackText(textBox, style) {
    const textContent = textBox.textContent || textBox.innerText || "";
    if (!textContent.trim()) return [];

    const defaultOptions = {
        fontFace: decodeAndCleanFontFamily(style.fontFamily || "") || "Arial",
        fontSize: parseFloat(style.fontSize || "14"), // 1:1 conversion for 72 DPI
        color: rgbToHex(style.color || "#000000"),
        align: style.textAlign || "left",
        bold: false,
        italic: false,
        underline: false
    };

    return [{
        text: textContent.trim(),
        options: defaultOptions
    }];
}

function handleMixedContent(pptSlide, textBox, textBoxOptions, globalLineSpacing = null, globalFontSize = 14, defaultAlign = "left") {
    const allContent = [];

    // Get the sli-txt-box div which contains all the content
    const contentContainer = textBox.querySelector('.sli-txt-box') || textBox;

    // Process all direct children in order
    const children = Array.from(contentContainer.children);

    children.forEach((child, index) => {
        const tagName = child.tagName ? child.tagName.toLowerCase() : '';

        if (tagName === 'ul' || tagName === 'ol') {
            // Process list
            const listItems = [];
            processListRecursively(child, listItems, 0, globalLineSpacing, globalFontSize);

            // FIXED: Remove breakLine from the last item if there's more content after this list
            const hasContentAfter = index < children.length - 1;

            // Add list items to content
            listItems.forEach((item, itemIdx) => {
                const isLastItem = itemIdx === listItems.length - 1;

                // If this is the last item of this list AND there's content after, 
                // keep breakLine to ensure proper spacing
                allContent.push({
                    text: item.text,
                    options: item.options
                });
            });
        } else if (tagName === 'p') {
            // Process paragraph
            const paragraphHTML = child.innerHTML.trim();

            // Check if it's an empty paragraph
            const isBreakOnlyParagraph = paragraphHTML === '<br>' ||
                paragraphHTML === '<br/>' ||
                paragraphHTML === '<br />' ||
                paragraphHTML === '&nbsp;' ||
                paragraphHTML === '&#160;' ||
                (paragraphHTML.replace(/<br\s*\/?>/gi, '').trim() === '') ||
                (paragraphHTML.replace(/&nbsp;/gi, '').replace(/&#160;/gi, '').trim() === '');

            const hasMinHeight = child.style.minHeight && parseFloat(child.style.minHeight) > 0;
            const textContent = child.textContent || child.innerText || '';
            const isEmptyParagraph = isBreakOnlyParagraph ||
                (textContent.replace(/\u00A0/g, '').replace(/\s/g, '') === '' && hasMinHeight);

            if (isEmptyParagraph) {
                // FIXED: For empty paragraphs, add a line break to the last item if it exists
                // This maintains spacing between lists and paragraphs
                if (allContent.length > 0) {
                    const lastItem = allContent[allContent.length - 1];
                    // Only add newline if the last item doesn't already end with one
                    if (!lastItem.text.endsWith('\n')) {
                        lastItem.text += '\n';
                    }
                }
            } else {
                // Process non-empty paragraph
                const paragraphText = processSpansInParagraphWithLineSpacing(child);

                if (paragraphText && paragraphText.length > 0) {
                    // FIXED: Add paragraph break without line spacing
                    if (allContent.length > 0) {
                        const breakOptions = { ...paragraphText[0].options };
                        delete breakOptions.lineSpacing; // Remove line spacing for the break

                        allContent.push({
                            text: '\n',
                            options: breakOptions
                        });
                    }

                    paragraphText.forEach(textRun => {
                        allContent.push(textRun);
                    });
                }
            }
        }
    });

    // FIXED: Remove breakLine from the very last item to prevent extra space at the end
    if (allContent.length > 0) {
        const lastItem = allContent[allContent.length - 1];
        if (lastItem.options && lastItem.options.breakLine) {
            delete lastItem.options.breakLine;
        }
    }

    // Add all content to the slide
    if (allContent.length > 0) {
        try {
            const mixedTextBoxOptions = { ...textBoxOptions };
            delete mixedTextBoxOptions.lineSpacing;  // Let individual text runs handle line spacing

            pptSlide.addText(allContent, mixedTextBoxOptions);
        } catch (error) {
            console.error("Error adding mixed content to slide:", error);
            // Attempt fallback
            try {
                const fallbackContent = allContent.map(item => {
                    const safeOptions = { ...item.options };
                    delete safeOptions.lineSpacing;
                    return {
                        text: item.text.replace(/[\r\n\t\f\v]/g, ' ').trim(),
                        options: safeOptions
                    };
                });

                pptSlide.addText(fallbackContent, mixedTextBoxOptions);
            } catch (fallbackError) {
                console.error("Fallback failed:", fallbackError);
            }
        }
    }
}

function handleListContent(pptSlide, textBox, textBoxOptions, globalLineSpacing = null, globalFontSize = 14) {
    // Only get top-level lists (not nested ones)
    const topLevelLists = [];
    const allLists = textBox.querySelectorAll ? Array.from(textBox.querySelectorAll("ul, ol")) : [];

    // Filter to only include lists that are direct children of textBox or don't have a parent list
    allLists.forEach(list => {
        let hasParentList = false;
        let parent = list.parentElement;
        while (parent && parent !== textBox) {
            if (parent.tagName && (parent.tagName.toLowerCase() === 'ul' || parent.tagName.toLowerCase() === 'ol')) {
                hasParentList = true;
                break;
            }
            parent = parent.parentElement;
        }
        if (!hasParentList) {
            topLevelLists.push(list);
        }
    });

    if (topLevelLists.length === 0) return;

    const formattedListItems = [];

    topLevelLists.forEach((list) => {
        processListRecursively(list, formattedListItems, 0, globalLineSpacing, globalFontSize);
    });

    // Add all the processed list items to the slide in one call
    if (formattedListItems.length > 0) {
        try {
            const listTextBoxOptions = { ...textBoxOptions };
            delete listTextBoxOptions.lineSpacing;  // Remove global line spacing for lists

            pptSlide.addText(formattedListItems, listTextBoxOptions);
        } catch (error) {
            console.error("Error adding list to slide:", error);
            // Attempt fallback processing
            try {
                const fallbackItems = formattedListItems.map((item, itemIndex) => {
                    const cleanText = item.text.replace(/[\r\n\t\f\v]/g, ' ').trim();
                    const safeOptions = { ...item.options };
                    delete safeOptions.lineSpacing;

                    return {
                        text: cleanText,
                        options: {
                            ...safeOptions,
                            bullet: true,
                            breakLine: itemIndex < formattedListItems.length - 1
                        }
                    };
                });

                pptSlide.addText(fallbackItems, listTextBoxOptions);
            } catch (fallbackError) {
                console.error("All attempts failed:", fallbackError);
            }
        }
    }
}

// NEW: Recursive function to process lists with proper nesting
function processListRecursively(list, formattedListItems, indentLevel, globalLineSpacing, globalFontSize) {
    const isOrdered = list.tagName.toLowerCase() === 'ol';

    // Get only direct child list items (not nested ones)
    const directChildItems = Array.from(list.children).filter(child =>
        child.tagName && child.tagName.toLowerCase() === 'li'
    );

    // FIXED: Also get any direct child lists (malformed HTML where lists are direct children of lists)
    const directChildLists = Array.from(list.children).filter(child =>
        child.tagName && (child.tagName.toLowerCase() === 'ul' || child.tagName.toLowerCase() === 'ol')
    );

    directChildItems.forEach((item, itemIndex) => {
        // Extract only the direct text content (not from nested lists)
        const directTextContent = getDirectTextContent(item);

        if (!directTextContent || directTextContent.trim().length === 0) {
            // If no direct text, check nested lists anyway
            const nestedLists = Array.from(item.children).filter(child =>
                child.tagName && (child.tagName.toLowerCase() === 'ul' || child.tagName.toLowerCase() === 'ol')
            );

            nestedLists.forEach(nestedList => {
                processListRecursively(nestedList, formattedListItems, indentLevel + 1, globalLineSpacing, globalFontSize);
            });
            return;
        }

        // Get spans only from direct children (not nested lists)
        const directSpans = getDirectSpans(item);

        let combinedText = "";
        let itemFormatting = null;
        let hasValidContent = false;

        if (directSpans.length > 0) {
            directSpans.forEach((span) => {
                const spanText = getCleanTextContent(span);

                if (spanText && spanText.trim().length > 0) {
                    combinedText += spanText;

                    // Use formatting from the first span, including line spacing
                    if (!itemFormatting) {
                        itemFormatting = extractSpanFormattingFromListWithLineSpacing(span, globalLineSpacing, globalFontSize);
                    }

                    hasValidContent = true;
                }
            });
        } else {
            combinedText = directTextContent;
            itemFormatting = extractListItemFormattingWithLineSpacing(item, globalLineSpacing, globalFontSize);
            hasValidContent = !!combinedText.trim();
        }

        // Add bullet formatting
        if (hasValidContent && combinedText.trim()) {
            if (isOrdered) {
                itemFormatting.bullet = { type: "number" };
            } else {
                // Use standard bullet format
                itemFormatting.bullet = true;
            }

            // Add indentation level for nested items
            if (indentLevel > 0) {
                itemFormatting.indentLevel = indentLevel;
            }

            // Always add breakLine except for the very last item
            itemFormatting.breakLine = true;

            formattedListItems.push({
                text: combinedText.trim(),
                options: itemFormatting
            });
        }

        // Process nested lists recursively
        const nestedLists = Array.from(item.children).filter(child =>
            child.tagName && (child.tagName.toLowerCase() === 'ul' || child.tagName.toLowerCase() === 'ol')
        );

        nestedLists.forEach(nestedList => {
            processListRecursively(nestedList, formattedListItems, indentLevel + 1, globalLineSpacing, globalFontSize);
        });
    });

    // FIXED: Process any direct child lists (malformed HTML handling)
    // These will be treated at the same indent level as they are siblings to the list items
    directChildLists.forEach(childList => {
        processListRecursively(childList, formattedListItems, indentLevel, globalLineSpacing, globalFontSize);
    });
}

// NEW: Get only direct text content from an element (excluding nested lists)
function getDirectTextContent(element) {
    let text = "";

    for (let i = 0; i < element.childNodes.length; i++) {
        const node = element.childNodes[i];

        // Text node
        if (node.nodeType === 3) {
            text += node.textContent || "";
        }
        // Element node but NOT a list
        else if (node.nodeType === 1) {
            const tagName = node.tagName.toLowerCase();
            if (tagName !== 'ul' && tagName !== 'ol') {
                // For non-list elements, include their text content
                text += node.textContent || "";
            }
        }
    }

    return text.replace(/[\u00A0\u2000-\u200F\u2028\u2029\u202F\u205F\u3000\uFEFF]/g, '')
        .replace(/[\r\n\t\f\v]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

// NEW: Get only direct child spans (not from nested lists)
function getDirectSpans(element) {
    const directSpans = [];

    function collectSpansRecursive(node, isInNestedList) {
        for (let i = 0; i < node.childNodes.length; i++) {
            const child = node.childNodes[i];

            if (child.nodeType === 1) {
                const tagName = child.tagName.toLowerCase();

                // If we hit a nested list, mark it
                if (tagName === 'ul' || tagName === 'ol') {
                    continue; // Don't process nested lists
                }

                // If it's a span and we're not in a nested list, collect it
                if (tagName === 'span' && !isInNestedList) {
                    directSpans.push(child);
                }

                // Recurse into other elements (but not lists)
                if (tagName !== 'ul' && tagName !== 'ol') {
                    collectSpansRecursive(child, isInNestedList);
                }
            }
        }
    }

    collectSpansRecursive(element, false);
    return directSpans;
}

// NEW: Enhanced span formatting for lists with line spacing
function extractSpanFormattingFromListWithLineSpacing(span, globalLineSpacing = null, globalFontSize = 14) {
    const style = span.style || {};
    const spanStyleAttr = span.getAttribute('style') || '';

    // For lists, use globalFontSize as baseline
    const fontSize = parseFloat(style.fontSize || globalFontSize);
    const position = style.position || "";
    const top = style.top || "";

    let finalFontSize = fontSize;
    let isSuperscript = false;
    let isSubscript = false;

    // Simple detection for lists
    if (position === "relative" && top) {
        const topValue = parseFloat(top);
        const fontRatio = fontSize / globalFontSize;

        if (fontRatio < 0.9 && Math.abs(topValue) > 2) {
            finalFontSize = globalFontSize; // Use base font size
            if (topValue < 0) {
                isSuperscript = true;
            } else {
                isSubscript = true;
            }
        }
    }

    try {
        let fontSizePx = finalFontSize;
        if (fontSizePx < 6 || fontSizePx > 72) {
            fontSizePx = globalFontSize;
        }
        const fontSizePt = fontSizePx;

        let fontColor = "FFFFFF";
        try {
            const extractedColor = rgbToHex(style.color || "#FFFFFF");
            fontColor = extractedColor.replace('#', '');
        } catch (colorError) {
            console.warn("Color extraction failed, using white:", colorError);
        }

        let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
        if (!fontFamily || fontFamily.length === 0) {
            fontFamily = "Arial";
        }

        const fontWeight = style.fontWeight || "normal";
        const isBold = fontWeight === "bold" || fontWeight === "bolder" || parseInt(fontWeight) >= 700;

        const fontStyle = style.fontStyle || "";
        const isItalic = fontStyle.includes("italic") || fontStyle.includes("oblique");

        let isUnderlined = false;
        const textDecoration = style.textDecoration || "";
        if (textDecoration.includes('underline')) {
            isUnderlined = true;
        }

        const textAlign = style.textAlign || "left";
        const validAlignments = ["left", "center", "right", "justify"];
        const finalAlign = validAlignments.includes(textAlign) ? textAlign : "left";

        const options = {
            fontFace: fontFamily,
            fontSize: fontSizePt,
            color: fontColor,
            align: finalAlign,
            bold: isBold,
            italic: isItalic,
            underline: isUnderlined
        };

        // Add superscript/subscript
        if (isSuperscript) {
            options.baseline = 300; // PptxGenJS expects smaller values
        } else if (isSubscript) {
            options.baseline = -250; // PptxGenJS expects smaller values
        }

        // Extract line spacing from span, fallback to global
        const lineHeightFromSpan = extractLineHeightFromCSS(spanStyleAttr);
        if (lineHeightFromSpan) {
            const lineSpacing = convertLineHeightToPowerPoint(lineHeightFromSpan, fontSizePx);
            if (lineSpacing) {
                options.lineSpacing = lineSpacing;
            }
        } else if (globalLineSpacing) {
            options.lineSpacing = globalLineSpacing;
        }

        return options;
    } catch (error) {
        console.warn("Error extracting span formatting, using defaults:", error);
        return getDefaultFormattingOptions();
    }
}

// NEW: Enhanced list item formatting with line spacing
function extractListItemFormattingWithLineSpacing(listItem, globalLineSpacing = null, globalFontSize = 14) {
    const style = listItem.style || {};
    const itemStyleAttr = listItem.getAttribute('style') || '';

    try {
        let fontSizePx = parseFloat(style.fontSize) || globalFontSize;
        if (fontSizePx < 6 || fontSizePx > 72) {
            fontSizePx = globalFontSize;
        }
        const fontSizePt = fontSizePx;

        let fontColor = "FFFFFF";
        try {
            const colorSource = style.color || "#FFFFFF";
            const extractedColor = rgbToHex(colorSource);
            fontColor = extractedColor.replace('#', '');
        } catch (colorError) {
            console.warn("Color extraction failed, using white:", colorError);
        }

        let fontFamily = decodeAndCleanFontFamily(style.fontFamily || "");
        if (!fontFamily || fontFamily.length === 0) {
            fontFamily = "Arial";
        }

        const fontWeight = style.fontWeight || "normal";
        const isBold = fontWeight === "bold" || fontWeight === "bolder" || parseInt(fontWeight) >= 700;

        const fontStyle = style.fontStyle || "";
        const isItalic = fontStyle.includes("italic") || fontStyle.includes("oblique");

        const textDecoration = style.textDecoration || "";
        const isUnderlined = textDecoration.includes('underline');

        const textAlign = style.textAlign || "left";
        const validAlignments = ["left", "center", "right", "justify"];
        const finalAlign = validAlignments.includes(textAlign) ? textAlign : "left";

        const options = {
            fontFace: fontFamily,
            fontSize: fontSizePt,
            color: fontColor,
            align: finalAlign,
            bold: isBold,
            italic: isItalic,
            underline: isUnderlined
        };

        // Extract line spacing from list item, fallback to global
        const lineHeightFromItem = extractLineHeightFromCSS(itemStyleAttr);
        if (lineHeightFromItem) {
            const lineSpacing = convertLineHeightToPowerPoint(lineHeightFromItem, fontSizePx);
            if (lineSpacing) {
                options.lineSpacing = lineSpacing;
            }
        } else if (globalLineSpacing) {
            options.lineSpacing = globalLineSpacing;
        }

        return options;
    } catch (error) {
        console.warn("Error extracting list item formatting, using defaults:", error);
        return getDefaultFormattingOptions();
    }
}

function getIndentationLevel(listItem) {
    let level = 0;
    let currentElement = listItem.parentElement;

    // Traverse up the DOM to count nested list levels
    while (currentElement) {
        if (currentElement.tagName &&
            (currentElement.tagName.toLowerCase() === 'ul' ||
                currentElement.tagName.toLowerCase() === 'ol')) {
            level++;
        }
        currentElement = currentElement.parentElement;
    }

    // Adjust: first level should be 0
    return Math.max(0, level - 1);
}

function extractListItemFormatting(listItem) {
    const style = listItem.style || {};

    try {
        // Extract font size with validation
        let fontSizePx = parseFloat(style.fontSize) ||
            parseFloat(getComputedStyle(listItem)?.fontSize) ||
            14;
        if (fontSizePx < 6 || fontSizePx > 72) {
            fontSizePx = 14;
        }
        const fontSizePt = fontSizePx;

        // Extract color with multiple fallbacks
        let fontColor = "FFFFFF"; // Default white for dark slides
        try {
            const colorSource = style.color || getComputedStyle(listItem)?.color || "#FFFFFF";
            const extractedColor = rgbToHex(colorSource);
            fontColor = extractedColor.replace('#', '');
        } catch (colorError) {
            console.warn("Color extraction failed, using white:", colorError);
        }

        // Extract font family with validation
        let fontFamily = decodeAndCleanFontFamily(
            style.fontFamily ||
            getComputedStyle(listItem)?.fontFamily ||
            ""
        );
        if (!fontFamily || fontFamily.length === 0) {
            fontFamily = "Arial";
        }

        // Extract font weight
        const fontWeight = style.fontWeight ||
            getComputedStyle(listItem)?.fontWeight ||
            "normal";
        const isBold = fontWeight === "bold" ||
            fontWeight === "bolder" ||
            parseInt(fontWeight) >= 700;

        // Extract font style
        const fontStyle = style.fontStyle ||
            getComputedStyle(listItem)?.fontStyle ||
            "";
        const isItalic = fontStyle.includes("italic") || fontStyle.includes("oblique");

        // Extract text decoration
        const textDecoration = style.textDecoration ||
            getComputedStyle(listItem)?.textDecoration ||
            "";
        const isUnderlined = textDecoration.includes('underline');

        // Extract text alignment
        const textAlign = style.textAlign ||
            getComputedStyle(listItem)?.textAlign ||
            "left";
        const validAlignments = ["left", "center", "right", "justify"];
        const finalAlign = validAlignments.includes(textAlign) ? textAlign : "left";

        return {
            fontFace: fontFamily,
            fontSize: fontSizePt,
            color: fontColor,
            align: finalAlign,
            bold: isBold,
            italic: isItalic,
            underline: isUnderlined
        };
    } catch (error) {
        console.warn("Error extracting list item formatting, using defaults:", error);
        return getDefaultFormattingOptions();
    }
}

function normalizeStyleValue(value) {
    if (!value) return 0;
    return parseFloat(value.replace(/[^-\d.]/g, '')) || 0;
}

function decodeAndCleanFontFamily(fontFamily) {
    if (!fontFamily) return "";
    if (fontFamily == "UniversCondensedLightBody") {
        fontFamily = "Univers Condensed Light (Body)";
    }


    let decodedFont = fontFamily
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&amp;/g, '&');
    decodedFont = decodedFont.replace(/['"]/g, '');
    return decodedFont.trim();
}



// MODIFIED: Enhanced preprocessing to handle empty paragraphs
function preprocessTextFormatting(textBox) {
    if (!textBox) return textBox;

    if (typeof textBox !== 'object' || !textBox.querySelectorAll) {
        return textBox;
    }

    try {
        const paragraphs = textBox.querySelectorAll('p');

        paragraphs.forEach(p => {
            let html = p.innerHTML;

            // Handle underline formatting
            if (html.includes('<u style="text-decoration: none">') ||
                html.includes('<u style="text-decoration:none">')) {
                html = html.replace(/<u style="text-decoration:\s*none">(.*?)<\/u>/g, '$1');
                p.innerHTML = html;
            }

            const uTags = p.querySelectorAll('u');

            if (uTags.length > 0) {
                let style = p.getAttribute('style') || '';

                if (!style.includes('text-decoration')) {
                    style += (style ? '; ' : '') + 'text-decoration: underline;';
                    p.setAttribute('style', style);
                } else if (style.includes('text-decoration: none')) {
                    style = style.replace('text-decoration: none', 'text-decoration: underline');
                    p.setAttribute('style', style);
                } else if (!style.includes('text-decoration: underline')) {
                    style += (style ? '; ' : '') + 'text-decoration: underline;';
                    p.setAttribute('style', style);
                }

                html = p.innerHTML;
                html = html.replace(/<u(?:\s+[^>]*)?>(.*?)<\/u>/g, '$1');
                p.innerHTML = html;
            }

            // FIXED: Normalize &nbsp; entities in empty paragraphs
            const textContent = p.textContent || p.innerText || '';
            const hasMinHeight = p.style.minHeight && parseFloat(p.style.minHeight) > 0;

            // If paragraph only contains non-breaking spaces and has min-height, mark it as empty
            if (textContent.replace(/\u00A0/g, '').replace(/\s/g, '') === '' && hasMinHeight) {
                // Ensure the paragraph contains &nbsp; for consistent detection
                if (html.trim() !== '&nbsp;') {
                    p.innerHTML = '&nbsp;';
                }
            }
        });

    } catch (error) {
        console.warn('Error in preprocessTextFormatting, skipping: ', error.message);
    }

    return textBox;
}

function getCleanTextContent(element) {
    if (!element) return "";

    let text = element.textContent || element.innerText || "";

    // Remove ALL possible invisible/whitespace characters
    text = text
        .replace(/[\u00A0\u2000-\u200F\u2028\u2029\u202F\u205F\u3000\uFEFF]/g, '') // Remove unicode spaces
        .replace(/[\r\n\t\f\v]/g, ' ') // Replace all whitespace with regular spaces
        .replace(/\s+/g, ' ') // Collapse multiple spaces
        .replace(/&nbsp;/gi, '') // Remove HTML entities
        .replace(/&#160;/g, '') // Remove numeric HTML entities
        .trim();

    return text;
}

function isValidListItemContent(textContent, element) {

    // Only reject truly empty or whitespace-only content
    if (!textContent || textContent.trim().length === 0) {
        return false;
    }

    // Must have at least some meaningful characters (letters or numbers)
    if (!textContent.match(/[a-zA-Z0-9]/)) {
        return false;
    }

    // Must have at least 2 characters total
    if (textContent.trim().length < 2) {
        return false;
    }

    // Accept almost everything else - bullet points can be short phrases
    return true;
}

function isValidSpanContent(spanText) {
    // Use the same permissive validation for spans
    return isValidListItemContent(spanText, null);
}

function debugListStructure(textBox) {

    try {
        const lists = textBox.querySelectorAll ? textBox.querySelectorAll("ul, ol") : [];

        lists.forEach((list, listIndex) => {

            const listItems = list.querySelectorAll("li");

            listItems.forEach((item, itemIndex) => {
                try {
                    const rawText = item.textContent || item.innerText || "";
                    const cleanText = getCleanTextContent(item);
                    const isValid = isValidListItemContent(cleanText, item);

                    const spans = item.querySelectorAll('span');
                    if (spans.length > 0) {
                        spans.forEach((span, spanIndex) => {
                            try {
                                const spanRaw = span.textContent || span.innerText || "";
                                const spanClean = getCleanTextContent(span);
                                const spanValid = isValidSpanContent(spanClean);
                            } catch (spanError) {
                                console.error(`Error processing span ${spanIndex}:`, spanError.message);
                            }
                        });
                    }
                } catch (itemError) {
                    console.error(`Error processing list item ${itemIndex}:`, itemError.message);
                }
            });
        });
    } catch (error) {
        console.error("Error in debugListStructure:", error.message);
    }

}

function preFilterListItems(listElement) {
    const listItems = Array.from(listElement.querySelectorAll("li"));
    const validItems = [];

    listItems.forEach((item, index) => {
        const textContent = getCleanTextContent(item);
        const isValid = isValidListItemContent(textContent, item);

        if (isValid) {
            validItems.push(item);
        } else {
            // DON'T remove from DOM - just don't include in validItems
        }
    });

    return validItems;
}

function getDefaultFormattingOptions() {
    return {
        fontFace: "Arial",
        fontSize: 14,
        color: "FFFFFF",
        align: "left",
        bold: false,
        italic: false,
        underline: false
    };
}



function rgbToHex(color) {
    if (!color) return '#000000';

    color = color.toLowerCase().trim();

    if (color.startsWith('#')) {
        let hex = color.substring(1);
        if (hex.length === 3) {
            hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        }
        return '#' + hex.toUpperCase();
    }

    const namedColors = {
        'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'green': '008000',
        'blue': '0000FF', 'yellow': 'FFFF00', 'cyan': '00FFFF', 'magenta': 'FF00FF',
        'silver': 'C0C0C0', 'gray': '808080', 'grey': '808080', 'maroon': '800000',
        'olive': '808000', 'lime': '00FF00', 'aqua': '00FFFF', 'teal': '008080',
        'navy': '000080', 'fuchsia': 'FF00FF', 'purple': '800080', 'orange': 'FFA500'
    };

    if (namedColors[color]) {
        return '#' + namedColors[color];
    }

    if (color.startsWith('rgba')) {
        const match = color.match(/rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([0-9.]+)\s*\)/);
        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);
            const alpha = parseFloat(match[4]);

            if (alpha < 1) {
                const blendedR = Math.round(r * alpha + 255 * (1 - alpha));
                const blendedG = Math.round(g * alpha + 255 * (1 - alpha));
                const blendedB = Math.round(b * alpha + 255 * (1 - alpha));

                return '#' +
                    blendedR.toString(16).padStart(2, '0').toUpperCase() +
                    blendedG.toString(16).padStart(2, '0').toUpperCase() +
                    blendedB.toString(16).padStart(2, '0').toUpperCase();
            }

            return '#' +
                r.toString(16).padStart(2, '0').toUpperCase() +
                g.toString(16).padStart(2, '0').toUpperCase() +
                b.toString(16).padStart(2, '0').toUpperCase();
        }
    }

    else if (color.startsWith('rgb')) {
        const match = color.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);
            return '#' +
                r.toString(16).padStart(2, '0').toUpperCase() +
                g.toString(16).padStart(2, '0').toUpperCase() +
                b.toString(16).padStart(2, '0').toUpperCase();
        }
    }

    const numbers = color.match(/\d+/g);
    if (numbers && numbers.length >= 3) {
        const r = Math.min(255, parseInt(numbers[0]));
        const g = Math.min(255, parseInt(numbers[1]));
        const b = Math.min(255, parseInt(numbers[2]));
        return '#' +
            r.toString(16).padStart(2, '0').toUpperCase() +
            g.toString(16).padStart(2, '0').toUpperCase() +
            b.toString(16).padStart(2, '0').toUpperCase();
    }

    return '#000000';
}

module.exports = {
    addTextBoxToSlide
};