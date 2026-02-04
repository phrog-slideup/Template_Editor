const addTextBox = require("./addTextBoxToSlide.js");
const svgAddToSlide = require("./addSvgToSlide.js");

function addShapeToSlide(pptx, pptSlide, shapeElement, slideContext) {
    const style = shapeElement.style;
    const shapeId = shapeElement.getAttribute("id");
    const textBox = shapeElement.querySelector('.sli-txt-box');
    const objName = shapeElement.getAttribute("data-name");

    console.log(" >>>>>>>>>>>>>>> ---- >>>> ", shapeId);

    // Extract theme color and luminance/alpha attributes
    const originalThemeColor = shapeElement.getAttribute("data-original-color");
    const originalLumMod = shapeElement.getAttribute("originallummod");
    const originalLumOff = shapeElement.getAttribute("originallumoff");
    const originalAlpha = shapeElement.getAttribute("originalalpha");

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

    // Skip ONLY if it's a text box with no background AND no border
    if (textBox && textBox.textContent.trim() && !hasVisibleBackground && !hasBorder) {
        return;
    }


    // ✅ Freeform / SVG-based shapes (custGeom, custom-shape, svg connectors)
    // Route these to addSvgToSlide so we actually create a real custGeom in PPTX.
    if (
        shapeElement.classList.contains('custom-shape') ||
        shapeElement.id === 'custGeom' ||
        shapeId === 'custGeom' ||
        shapeElement.classList.contains('sli-svg-connector')
    ) {
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
    const slideDimensions = getSlideDimensions(pptSlide);

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

    // Extract position and dimensions with slide scaling
    const x = parseFloat(style.left || "0") / 72;
    const y = parseFloat(style.top || "0") / 72;
    const w = parseFloat(style.width || "0") / 72;
    const h = parseFloat(style.height || "0") / 72;

    let shapeOptions = {
        x: x,
        y: y,
        w: w,
        h: h,
        rotate: rotation,
        objectName: objName || '',
        hidden: true
    };

    console.log(" >>>>>>>>>>>>>>>>>>> -- Shape Options so far:", shapeOptions);

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

    if (style.opacity) {
        const opacityValue = parseFloat(style.opacity);
        const transparencyValue = Math.round((1 - opacityValue) * 100);

        // NEW: pass originalAlpha into the color object builder
        const themeColorObj = createThemeColorObject(originalThemeColor, originalLumMod, originalLumOff, originalAlpha);

        if (themeColorObj) {
            shapeOptions.fill = {
                color: themeColorObj
            };
        } else {
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
    }

    if (borderRadiusValue > 0) {
        if (shapeId === 'rect' ||
            shapeId === 'roundRect' ||
            shapeId === 'round1Rect' ||
            shapeId === 'round2DiagRect' ||
            shapeId === 'round2SameRect' ||
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

    if (gradient) {
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
        // All the existing switch cases remain the same...
        switch (shapeId) {
            // Action Buttons
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
            case 'can':
                pptSlide.addShape(pptx.shapes.CAN, shapeOptions);
                break;
            case 'chartPlus':
                pptSlide.addShape(pptx.shapes.CHART_PLUS, shapeOptions);
                break;
            case 'chartStar':
                pptSlide.addShape(pptx.shapes.CHART_STAR, shapeOptions);
                break;
            case 'chartX':
                pptSlide.addShape(pptx.shapes.CHART_X, shapeOptions);
                break;
            case 'chevron':
                pptSlide.addShape(pptx.shapes.CHEVRON, shapeOptions);
                break;
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
            case 'cube':
                pptSlide.addShape(pptx.shapes.CUBE, shapeOptions);
                break;
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
            case 'flowChartDocument':
                pptSlide.addShape(pptx.shapes.FLOWCHART_DOCUMENT, shapeOptions);
                break;
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
            case 'flowChartOnlineStorage':
                pptSlide.addShape(pptx.shapes.FLOWCHART_STORED_DATA, shapeOptions);
                break;
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
            case 'triangle':
                pptSlide.addShape(pptx.shapes.ISOSCELES_TRIANGLE, shapeOptions);
                break;
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

                // console.log("shapeOptions ====-----------!!!!!!!!!!!!!!!!------------- >>>>>>>>>>>>> ", {
                //     ...shapeOptions,
                //     line: {
                //         color: rgbToHex(lineColor),
                //         width: lineWidth
                //     }
                // });

                pptSlide.addShape(pptx.shapes.LINE, {
                    ...shapeOptions,
                    line: {
                        color: rgbToHex(lineColor),
                        width: lineWidth
                    }
                });
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
            case 'pie':
                pptSlide.addShape(pptx.shapes.PIE, shapeOptions);
                break;
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

                console.log(" >>>>>>>>> 222222 >>>>>>>>>> -- Shape Options so far:", shapeOptions);

                pptSlide.addShape(pptx.shapes.RECTANGLE, shapeOptions);
                break;
            case 'wedgeRectCallout':
                pptSlide.addShape(pptx.shapes.RECTANGULAR_CALLOUT, shapeOptions);
                break;
            case 'pentagon':
                pptSlide.addShape(pptx.shapes.REGULAR_PENTAGON, shapeOptions);
                break;
            case 'rightArrow':
                pptSlide.addShape(pptx.shapes.RIGHT_ARROW, shapeOptions);
                break;
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

                console.log(" >>>>>>>>  33333333 >>>>>>>>>>> -- Shape Options so far:", shapeOptions);

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
            case 'round2SameRect':
                pptSlide.addShape(pptx.shapes.ROUND_2_SAME_RECTANGLE, shapeOptions);
                break;
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
            case 'teardrop':
                pptSlide.addShape(pptx.shapes.TEAR, shapeOptions);
                break;
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

module.exports = {
    addShapeToSlide,
};