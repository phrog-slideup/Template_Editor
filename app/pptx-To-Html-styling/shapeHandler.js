const fs = require("fs");

const path = require("path");
const pptTextAllInfo = require("../pptx-To-Html-styling/pptTextAllInfo.js");
const TransparencyHandler = require("../pptx-To-Html-styling/transparencyHandler.js");
const colorHelper = require("../api/helper/colorHelper.js");
const resolveThemeFont = require("../api/helper/resolveThemeFont.js");
const pptBackgroundColors = require("../pptx-To-Html-styling/pptBackgroundColors.js");
const { overrideGetImageFromPicture } = require("../api/helper/pptImagesHandling.js");
const imgSvgStyle = require("../pptx-To-Html-styling/imgSvgCss.js");
const config = require('../config.js');
const { generateHexagonSVG, extractHexagon3DProperties } = require("./shapes/svgHexagonHandler.js");
const BarChartHandler = require("./charts/barChartHandler.js");
const DoughnutChartHandler = require("./charts/DoughnutChartHandler.js");
const TwoDAreaChartHandler = require("./charts/2DAreaChartHandler.js");
const textHandler = require("../pptx-To-Html-styling/text/textHandler.js");
const convertTableXMLToHTML = require("../pptx-To-Html-styling/tables/convertTableXMLToHTML.js");
const getShapeBorderCSS = require("./shapes_Properties/shapeBorderHandler.js");
const shapeFillColor = require("./shapes_Properties/getShapeFillColor.js");
const getLineConnectorHandler = require("./lines_Connectors/linesConnectorHandler.js");
const freeFormShape = require("./free_Form_Shape/generateFreeForm.js");

// Define the directory to save images using config
const imageSavePath = path.resolve(__dirname, "../uploads/");
 
if (!fs.existsSync(imageSavePath)) {
    fs.mkdirSync(imageSavePath, { recursive: true });
}

class ShapeHandler {
    constructor(themeXML, clrMap, nodes, extractor, slidePath, relationshipsXML, masterXML, layoutXML, flag) {
        this.themeXML = themeXML;
        this.clrMap = clrMap;
        this.nodes = nodes;
        this.extractor = extractor;
        this.transparencyHandler = new TransparencyHandler();
        this.slidePath = slidePath;
        this.relationshipsXML = relationshipsXML;
        this.masterXML = masterXML;
        this.layoutXML = layoutXML;
        this.flag = flag;
        this.tableStyles = null;
    }

    // Add helper method to get the correct divisor
    getEMUDivisor() {
        // return parseInt(this.flag) === 1 ? 9525 : 12700;
        return 12700;
    }

    isCenterTitlePlaceholder(shapeNode) {
        const phElement = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];
        if (!phElement) return false;
        return phElement?.["$"]?.type === "ctrTitle";
    }

    isPlaceholder(shapeNode) {
        const phElement = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];
        return Boolean(phElement);
    }

    async convertAllShapesToHTML(shapeNodes, lineShapeTag, graphicNodes, themeXML, masterXML, zIndexMap = null) {

        this.zIndexMap = zIndexMap;

        this.shapeNameOccurrences = {};
        const allHtmlElements = [];

        // Process shapes
        for (const shapeNode of shapeNodes) {

            if (this.isMasterPlaceholder(shapeNode, masterXML)) {
                continue;
            }
            const phElement = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];

            const isPlaceholder = Boolean(phElement);

            let hasType = Boolean(phElement?.["$"]?.type);
            let hasIdx = Boolean(phElement?.["$"]?.idx);

            let shapeHtml = await this.convertShapeToHTML(shapeNode, themeXML, masterXML);

            if (typeof shapeHtml !== 'string') {
                console.error('Non-string shape HTML:', typeof shapeHtml, shapeHtml);
            }
            allHtmlElements.push(shapeHtml);
        }

        // Process connectors
        if (lineShapeTag && lineShapeTag.length > 0) {
            console.log(`Processing ${lineShapeTag.length} connectors`); // DEBUG
            for (const cxnSpNode of lineShapeTag) {
                try {
                    const connectorHtml = await getLineConnectorHandler.convertConnectorToHTML(
                        cxnSpNode, this.themeXML, this.clrMap, this.masterXML, this.layoutXML);
                    
                    if (typeof connectorHtml === 'string' && connectorHtml.trim()) {
                        allHtmlElements.push(connectorHtml);
                        console.log('Connector added:', cxnSpNode?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name); // DEBUG
                    } else {
                        console.warn('Empty connector:', cxnSpNode?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name);
                    }
                } catch (error) {
                    console.error('Connector error:', error);
                }
            }
        } else {
            console.warn('No connectors found in lineShapeTag');
        }

        // Process tables And Charts.
        for (const graphicsNode of graphicNodes) {

            // Detect if this graphicFrame contains a chart
            const checkGraphics = graphicsNode?.["a:graphic"]?.[0]?.["a:graphicData"]?.[0]?.["c:chart"];

            if (checkGraphics) {

                const chartRId = checkGraphics?.[0]?.['$']?.['r:id'];
                const targetPath = this.getTargetFromRels(this.relationshipsXML, chartRId);

                // PARSED chart XML (object) - used for reading chart contents
                const chartXML = await this.extractor.parseXML(targetPath);

                const chartName = targetPath.match(/ppt\/charts\/(chart\d+)\.xml/)[1];
                const chartRelsPath = `ppt/charts/_rels/${chartName}.xml.rels`;

                const chartRelsXML = await this.extractor.parseXML(chartRelsPath);

                const chartsFiles = this.getChartStyleFiles(chartRelsXML);

                let chartColorsXML = "";
                let chartStyleXML = "";

                if (chartsFiles.colors) {
                    chartColorsXML = await this.extractor.parseXML(`ppt/charts/${chartsFiles.colors}`);
                } else if (chartsFiles.style) {
                    chartStyleXML = await this.extractor.parseXML(`ppt/charts/${chartsFiles.style}`);
                }

                const chartType = this.getChartTypeFromParsed(chartXML);

                const chartShapeName =
                    graphicsNode?.["p:nvGraphicFramePr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name || "Chart";

                const chartZIndex = this.getZIndexForShape(chartShapeName);

                console.log(`📊 Chart "${chartShapeName}" → zIndex ${chartZIndex}`);

                let chartHandler = null;

                if (chartType === "bar") {
                    chartHandler = new BarChartHandler(
                        graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, this.themeXML
                    );
                }
                else if (chartType === "doughnut") {
                    chartHandler = new DoughnutChartHandler(
                        graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, this.themeXML
                    );
                } else if (chartType === "area") {
                    chartHandler = new TwoDAreaChartHandler(
                        graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, this.themeXML
                    );
                }
                else {
                    continue;
                }
                const chartHtml = await chartHandler.convertChartToHTML(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, this.themeXML, chartZIndex);

                if (typeof chartHtml !== "string") {
                    console.error("❌ Chart handler did not return HTML string");
                }

                allHtmlElements.push(chartHtml);
            } else {
                const tableHtml = await convertTableXMLToHTML.convertTableXMLToHTML(graphicsNode, this.themeXML, this.extractor, this.nodes);

                allHtmlElements.push(tableHtml);
            }
        }

        return allHtmlElements.join("");
    }

    getChartTypeFromParsed(chartXML) {
        if (!chartXML) return "unknown";

        // Handle namespaces (c:, a:, etc.)
        const space =
            chartXML.chartSpace ||
            chartXML["c:chartSpace"] ||
            chartXML["chartSpace"] ||
            chartXML;

        const chart =
            space.chart?.[0] ||
            space["c:chart"]?.[0] ||
            space.chart ||
            space["c:chart"];

        if (!chart) return "unknown";

        const plotArea =
            chart.plotArea?.[0] ||
            chart["c:plotArea"]?.[0] ||
            chart.plotArea ||
            chart["c:plotArea"];

        if (!plotArea) return "unknown";

        // Chart type detection
        if (plotArea.doughnutChart || plotArea["c:doughnutChart"]) return "doughnut";
        if (plotArea.pieChart || plotArea["c:pieChart"]) return "pie";
        if (plotArea.barChart || plotArea["c:barChart"]) return "bar";
        if (plotArea.lineChart || plotArea["c:lineChart"]) return "line";
        if (plotArea.areaChart || plotArea["c:areaChart"]) return "area";
        if (plotArea.scatterChart || plotArea["c:scatterChart"]) return "scatter";

        return "unknown";
    }

    getTargetFromRels(relsContent, relationshipId) {
        // If using xml2js, the structure would be something like:
        const relationships = relsContent?.Relationships?.Relationship;

        if (!relationships) return null;

        // Find the relationship with matching Id
        const targetRelationship = relationships.find(rel =>
            rel.$?.Id === relationshipId
        );

        if (targetRelationship?.$.Target) {
            let target = targetRelationship.$.Target;

            // Convert relative path "../charts/chart1.xml" to "ppt/charts/chart1.xml"
            if (target.startsWith('../')) {
                target = target.replace('../', 'ppt/');
            }
            if (target.startsWith('/')) {
                target = target.replace('/', '');
            }

            return target;
        }

        return null;
    }

    getChartStyleFiles(chartRelsXML) {
        const relationships = chartRelsXML?.Relationships?.Relationship || [];

        const colorsRel = relationships.find(rel =>
            rel.$?.Type === "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
        );

        const styleRel = relationships.find(rel =>
            rel.$?.Type === "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
        );

        return {
            colors: colorsRel?.$?.Target || null,
            style: styleRel?.$?.Target || null
        };
    }

    async convertShapeToHTML(shapeNode, themeXML) {
        // Add this helper function at the top
        const sanitizeShapeName = (name) => {
            if (!name) return '';
            return name.replace(/['"<>&]/g, '').trim() || '';
        };

        const shapeName = sanitizeShapeName(shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name);
        const hidden = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.hidden === "1";

        const line = shapeNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0];

        const zIndex = this.getZIndexForShape(shapeName);
        let strokeWidth = "1px";
        let strokeColor = "#042433";

        const blip = shapeNode?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0]?.["a:blip"]?.[0];

        // Check if this shape is a placeholder
        const isPlaceholder = this.isPlaceholder(shapeNode);
        const placeholderClass = isPlaceholder ? ' placeholder-picture' : '';

        if (blip && blip["$"] && blip["$"]["r:embed"] || blip === undefined) {
            const imageId = blip?.["$"]?.["r:embed"];
            const imageInfo = await this.getImageFromPicture(shapeNode, this.slidePath, this.relationshipsXML);

            if (imageInfo) {
                const rotation = pptTextAllInfo.getRotation(shapeNode);
                const imgcss = imgSvgStyle.returnImgSvgStyle(shapeNode);
                const position = pptTextAllInfo.getPositionFromShape(shapeNode);

                // Extract cropping for shapes with images
                const blipFillNode = shapeNode?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0];
                const cropData = this.extractImageCroppingComprehensive(blipFillNode);
                const croppingStyles = cropData ? this.generateComprehensiveCroppingStyles(cropData) : null;

                // Apply border-radius for ellipse detection
                const shapeStyle = imageInfo.shape === "ellipse" ? "border-radius: 50%; overflow: hidden;" : "";
                const borderCSS = imageInfo.border.width > 0 ? `border: ${imageInfo.border.width}px solid ${imageInfo.border.color};` : "";
                const boxShadowCSS = (imgcss.shadowOffsetX || imgcss.shadowOffsetY) ? `box-shadow: ${imgcss.shadowOffsetX}px ${imgcss.shadowOffsetY}px ${imgcss.shadowColor};` : "";
                const combinedFilter = imgcss.blurAmount ? `filter: blur(${imgcss.blurAmount}px) contrast(${imgcss.contrastValue || 1});` : "filter:none;";

                // Apply cropping styles if available
                const containerCroppingStyles = croppingStyles ? croppingStyles.containerStyles : 'overflow: hidden;';
                const imageCroppingStyles = croppingStyles ? croppingStyles.imageStyles : 'width: 100%; height: 100%; object-fit: cover; position: absolute; left: 0; top: 0;';

                if (imageInfo) {
                    return `<div class="shape image-container${placeholderClass}"
        data-shape-type="image"
        data-name="${shapeName}"
        srcRectL="${imageInfo.cropping && imageInfo.cropping.leftRaw ? imageInfo.cropping.leftRaw : ''}"
        srcRectR="${imageInfo.cropping && imageInfo.cropping.rightRaw ? imageInfo.cropping.rightRaw : ''}"
        srcRectT="${imageInfo.cropping && imageInfo.cropping.topRaw ? imageInfo.cropping.topRaw : ''}"
        srcRectB="${imageInfo.cropping && imageInfo.cropping.bottomRaw ? imageInfo.cropping.bottomRaw : ''}"
        style="position:absolute;
        left:${position.x}px;
        top:${position.y}px;
        height:${position.height}px;
        width:${position.width}px;
        ${shapeStyle}
        ${boxShadowCSS}
        ${containerCroppingStyles}
        transform: ${imgcss.transform} rotate(${rotation}deg);
        z-index:${zIndex};">
                            
                            <img src="${imageInfo.src}" alt="Img"
                                style="${imageCroppingStyles}
                                transform: ${imgcss.flipH} ${imgcss.flipV} rotate(${rotation}deg);
                                opacity:${imgcss.opacity};
                                z-index:${zIndex};
                                object-fit: cover;
                                ${boxShadowCSS}
                                ${combinedFilter}" />
                        </div>`;
                }
                // Rest of your shape rendering code...
            }
        }

        const position = this.getShapePosition(shapeNode, this.masterXML);

        const stroke = this.extractStrokeProperties(shapeNode, themeXML);
        let shapeInfo = '';
        // Ensure get ShapeFillColor is correctly called once

        const fillProps = shapeFillColor.getShapeFillColor(shapeNode, themeXML, this.masterXML) ? shapeFillColor.getShapeFillColor(shapeNode, themeXML, this.masterXML) : { fillColor: "gray", opacity: 1 };
        const shadowStyle = this.getShadowStyle(shapeNode);

        const fillColor = fillProps?.fillColor || "gray";
        const originalThemeColor = fillProps?.originalThemeColor || null;
        const originalLumMod = fillProps?.originalLumMod || null;
        const originalLumOff = fillProps?.originalLumOff || null;
        const originalAlpha = fillProps?.originalAlpha || null;

        const opacity = `opacity : ${fillProps.opacity ? fillProps.opacity : 1}`;

        // const borderStyle = "solid";
        let cornerRadius = this.getCornerRadius ? this.getCornerRadius(shapeNode) : 0;
        const shapeType = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["$"]?.prst || "rect";

        const custGeom = shapeNode?.["p:spPr"]?.[0]?.["a:custGeom"]?.[0];

        const prstGeom = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["a:avLst"]?.[0]?.["a:gd"]?.[0]?.["$"].fmla;

        let borderRadius = shapeType === "ellipse" ? "50%" : `${cornerRadius}px`;
        let clipPath = "";
        let extraStyles = "";
        let maskPath = "";

        // Extracting border (stroke) properties
        const outline = shapeNode['p:spPr']?.[0]?.['a:ln']?.[0];
        // const shapeBorder = shapeNode['p:spPr']?.[0]?.['a:ln']?.[0]?.['a:solidFill'];
        const shapeBorder = await this.getShapeBorderStyle(outline);

        const shapeBorderStyle = getShapeBorderCSS.getShapeBorderCSS(shapeNode, this.clrMap, this.themeXML, this.masterXML);

        if (outline) {
            // Extract Stroke Width (Convert from EMUs to px)

            const sWidth = outline?.["$"]?.w;
            strokeWidth = `${parseInt(sWidth, 10) / this.getEMUDivisor()}px`;


            // Extract Stroke Color
            if (outline?.['a:solidFill']) {
                const solidFill = outline['a:solidFill'][0];

                if (solidFill['a:srgbClr']) {
                    strokeColor = `#${solidFill['a:srgbClr'][0].$.val}`; // RGB Hex
                } else if (solidFill['a:schemeClr']) {
                    strokeColor = colorHelper.resolveThemeColorHelper(solidFill['a:schemeClr'][0].$.val, themeXML, this.masterXML);
                }
            }
        }

        const txBoxBorder = this.extractStrokePropertiesForBorder(outline, themeXML, this.masterXML);

        let borderBoxStyle = "";
        if (txBoxBorder.width > 0) {
            borderBoxStyle = `border: ${txBoxBorder.width}px solid ${txBoxBorder.color};`;
        }

        // Check if the shape contains only text
        const isTextOnly = pptTextAllInfo.isTextShape(shapeNode);
        const textPlaceholderClass = (isTextOnly && isPlaceholder) ? ' placeholder-text' : '';

        const borderStyle = isTextOnly ? "none" : `solid ${strokeColor}`;

        function generateUniqueId(length = 10) {
            const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
            if (length > chars.length) {
                throw new Error(`Cannot generate an ID longer than ${chars.length} with unique characters.`);
            }
            const shuffled = chars.split('').sort(() => 0.5 - Math.random());
            return shuffled.slice(0, length).join('');
        }

        const uniqueId = generateUniqueId(10);

        // Extract text content from the shape
        let textContent = "";

        if (pptTextAllInfo.isTextShape(shapeNode)) {
            shapeInfo = textHandler.getAllTextInformationFromShape(shapeNode, this.themeXML, this.clrMap, this.masterXML, this.layoutXML);

            // Access the first element's properties (if you expect only one element)
            let txtPhType = shapeInfo.txtPhValues[0]?.type || '';
            let txtPhIdx = shapeInfo.txtPhValues[0]?.idx || '';
            let txtPhSz = shapeInfo.txtPhValues[0]?.sz || '';

            // Enhanced placeholder text cleaning
            const cleanedText = this.stripPlaceholderPrompts(shapeInfo.text || "");

            // For placeholders, be more aggressive about hiding master content
            if (isPlaceholder) {
                // Check if this is a master placeholder
                if (this.isMasterPlaceholder(shapeNode, this.masterXML)) {
                    return ""; // Hide master placeholders completely
                }

                // For slide placeholders, hide if no meaningful content
                if (!this.hasMeaningfulText(cleanedText)) {
                    return "";
                }
            }

            const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
            const transformString = this.getTransformString(position, isTextBox);
            // Rakesh Notes::: here from the below style i have remove the transformString variable as it is adding double rotation to bot the inner and outer (transform: ${transformString};)

            // NEW: Calculate if we need dynamic height
            const useHeightauto = shapeInfo.estimatedContentHeight && shapeInfo.estimatedContentHeight > position.height;
            const effectiveHeight = useHeightauto ? shapeInfo.estimatedContentHeight : position.height;

            if (this.hasMeaningfulText(cleanedText)) {
                textContent = `<div class="sli-txt-box${textPlaceholderClass}"
    contenteditable="true"
    spellcheck="false"
    data-editable="text"
    txtPhType="${txtPhType}"
    txtPhIdx="${txtPhIdx}"
    txtPhSz="${txtPhSz}"
    data-name="${shapeName}"
    id="${uniqueId}"
    style="
      color:${shapeInfo.fontColor};
      font-size:${shapeInfo.fontSize}px;
      display:flex;
      flex-direction: column;
      transform: ${transformString};
      ${opacity};
      justify-content: ${shapeInfo.justifyContent};
      text-align:${shapeInfo.textAlign};
      z-index:${zIndex};
      pointer-events:auto;
      user-select:text;
      outline:none;
    ">
    ${cleanedText}
</div>`;
            }
        }

        if (custGeom) {
            const svgMarkup = freeFormShape.generateCustomShapeSVG(custGeom, position, fillColor, stroke);
            return `<div id="custGeom" class="custom-shape" data-name="${shapeName}"
                        style="position:absolute; 
                        left:${position.x}px; 
                        top:${position.y}px; 
                        z-index:${zIndex};
                        width:${position.width}px; 
                        height:${position.height}px; 
                        transform: rotate(${position.rotation}deg) ${position.flipH ? 'scaleX(-1)' : ''} ${position.flipV ? 'scaleY(-1)' : ''};                    
                        ${opacity}; 
                        overflow:visible;">${textContent}${svgMarkup}
                    </div>`;
        }

        // Apply transparency and gradients if applicable
        const { gradientCSS } = this.transparencyHandler.getTransparency(shapeNode);
        let backgroundStyle = gradientCSS || fillColor;
        let caseName = '';

        let strokeDashArray = "";
        if (line) {
            const dashType = line?.["a:prstDash"]?.[0]?.["$"]?.val || "solid";
            if (dashType === "dash") {
                strokeDashArray = "5, 5";
            } else if (dashType === "dot") {
                strokeDashArray = "2, 5";
            } else if (dashType === "dashDot") {
                strokeDashArray = "5, 5, 2, 5";
            }
        }
        // Handling Different Shapes
        switch (shapeType) {
            // CORRECTED VERSION - Replace the 'line' case in convertShapeToHTML function
            case "line":
                caseName = "line";

                // Extract stroke width from XML
                let extractedStrokeWidth = 1;
                const lineElement = shapeNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0];

                if (lineElement && lineElement["$"]?.w) {
                    // EMU → px (72 DPI via this.getEMUDivisor)
                    extractedStrokeWidth = parseInt(lineElement["$"].w, 10) / this.getEMUDivisor();
                } else {
                    // Fallback to style reference
                    const styleRef = shapeNode?.["p:style"]?.[0];
                    const lnRef = styleRef?.["a:lnRef"]?.[0];
                    if (lnRef) {
                        const lnIdx = lnRef?.["$"]?.idx;
                        if (lnIdx) {
                            const lineWidthMap = { "0": 0.5, "1": 1, "2": 2, "3": 3, "4": 4.5, "5": 6 };
                            extractedStrokeWidth = lineWidthMap[lnIdx] || 1;
                        }
                    }
                }

                // Get line color / opacity
                const shapeColorsForLine = this.getShapeFillColor(shapeNode, themeXML, this.masterXML);
                const lineOpacity2 = shapeColorsForLine.strokeOpacity || 1.0;

                let finalStrokeColor2 = '#000000';
                if (shapeColorsForLine.strokeColor) {
                    finalStrokeColor2 = shapeColorsForLine.strokeColor;
                } else if (strokeColor && strokeColor !== '#042433') {
                    finalStrokeColor2 = strokeColor;
                } else if (shapeColorsForLine.fillColor) {
                    finalStrokeColor2 = shapeColorsForLine.fillColor;
                }

                // --- FIXED ANGLE AND LENGTH FOR p:sp LINES ---
                const xfrm2 = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
                const flipH2 = xfrm2?.["$"]?.flipH === "1";
                const flipV2 = xfrm2?.["$"]?.flipV === "1";

                let deltaX2 = position.width;
                let deltaY2 = position.height;

                if (flipH2) deltaX2 = -deltaX2;
                if (flipV2) deltaY2 = -deltaY2;

                const angleRad2 = Math.atan2(deltaY2, deltaX2);
                const angleDeg2 = angleRad2 * (180 / Math.PI);
                const lineLength2 = Math.sqrt(deltaX2 * deltaX2 + deltaY2 * deltaY2);

                // --- FIXED START POSITION (same logic as connectors) ---
                let startX2 = position.x;
                let startY2 = position.y;

                if (flipH2 && !flipV2) {
                    startX2 = position.x + position.width;
                    startY2 = position.y;
                } else if (!flipH2 && flipV2) {
                    startX2 = position.x;
                    startY2 = position.y + position.height;
                } else if (flipH2 && flipV2) {
                    startX2 = position.x + position.width;
                    startY2 = position.y + position.height;
                }

                return `<div class="shape line-lineheight" 
                    id="line-lineheight" 
                    data-shape-type="${caseName}" 
                    data-name="${shapeName}" 
                    style="position: absolute;
                            left: ${startX2}px;
                            top: ${startY2}px;
                            width: ${Math.abs(lineLength2)}px;
                            height: ${extractedStrokeWidth}px;
                            background: ${finalStrokeColor2};
                            transform: rotate(${angleDeg2}deg);
                            transform-origin: left center;
                            opacity: ${lineOpacity2};
                            border-radius: ${Math.ceil(extractedStrokeWidth / 2)}px;
                            z-index: ${zIndex};
                            cursor: pointer;"></div>`;
                break;

            case "rect":
                caseName = "rect";
                clipPath = "polygon(0% 0%, 100% 0%, 100% 100%, 0% 100%);" // Normal Rectangle
                break;
            case "ellipse":
                caseName = "ellipse";
                clipPath = "ellipse(50% 50% at 50% 50%)";
                break;
            case "triangle":
                caseName = "triangle";
                clipPath = "polygon(50% 0%, 100% 100%, 0% 100%)";
                break;
            case "hexagon":
                caseName = "hexagon";

                // Extract 3D properties from PPTX XML
                const hexagon3DProps = extractHexagon3DProperties(shapeNode, {
                    ignoreRotation: true,   // ← ADD THIS
                    useDefaultAdj: true     // ← ADD THIS
                });

                const hexagonSVG = generateHexagonSVG({
                    width: position.width,
                    height: position.height,
                    fillColor: fillColor,
                    strokeColor,
                    strokeWidth: parseFloat(strokeWidth.replace('px', '') || strokeWidth),
                    depth: hexagon3DProps.depth || 0,
                    rotation: hexagon3DProps.rotation || 0,
                    bevel: hexagon3DProps.bevel,
                    extrusionColor: hexagon3DProps.extrusionColor,
                    contourColor: hexagon3DProps.contourColor,
                    contourWidth: hexagon3DProps.contourWidth || 0,
                    hexagonAdj: hexagon3DProps.hexagonAdj,
                    hexagonVF: hexagon3DProps.hexagonVF,
                    material: 'matte',
                    lightDirection: hexagon3DProps.lightDirection || 't',
                    opacity: fillProps.opacity || 1
                });

                // Calculate transform string (same as other shapes)
                const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
                const transformString = this.getTransformString(position, isTextBox);

                return `<div class="shape" id="${caseName}" 
                            data-original-color="${originalThemeColor}" 
                            originalLumMod="${originalLumMod}" 
                            originalLumOff="${originalLumOff}" 
                            originalAlpha="${originalAlpha}"
                            data-name="${shapeName}" 
                            style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                display: flex;
                                transform: ${transformString};
                                box-sizing: border-box;
                                overflow: visible;
                                justify-content: ${shapeInfo.justifyContent || 'center'};
                                align-items: ${shapeInfo.getAlignItem || 'center'};
                                z-index: ${zIndex};">
                        ${hexagonSVG}
                        ${textContent}
                    </div>`;
                break;

            case "parallelogram":
                caseName = "parallelogram";
                clipPath = "polygon(15% 0%, 100% 0%, 85% 100%, 0% 100%)";
                break;
            case "pentagon":
                caseName = "pentagon";
                clipPath = "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)";
                break;
            case "trapezoid":
                caseName = "trapezoid";
                clipPath = "polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)";
                break;
            case "heptagon":
                caseName = "heptagon";
                clipPath = "polygon(50% 0%, 90% 20%, 100% 60%, 75% 100%, 25% 100%, 0% 60%, 10% 20%)";
                break;
            case "octagon":
                caseName = "octagon";
                clipPath = "polygon(30% 0%, 70% 0%, 100% 30%, 100% 70%, 70% 100%, 30% 100%, 0% 70%, 0% 30%)";
                break;
            case "nonagon":
                caseName = "nonagon";
                clipPath = "polygon(50% 0%, 83% 12%, 100% 43%, 94% 78%, 68% 100%, 32% 100%, 6% 78%, 0% 43%, 17% 12%)";
                break;
            case "decagon":
                caseName = "decagon";
                clipPath = "polygon(50% 0%, 80% 10%, 100% 35%, 100% 70%, 80% 90%, 50% 100%, 20% 90%, 0% 70%, 0% 35%, 20% 10%)";
                break;
            case "star":
                caseName = "star";
                clipPath = "polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)";
                break;
            case "rightArrow": {
                caseName = "rightArrow";

                // --- 1. Read adj1 (shaft height) & adj2 (head width) from a:gd ---
                let adj1 = 50000;  // PPT default if missing
                let adj2 = 50000;  // PPT default if missing

                const avLst =
                    shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["a:avLst"]?.[0];

                // xml2js can give "a:gd" or "gd" depending on config, handle both
                const gdList = (avLst && (avLst["a:gd"] || avLst["gd"])) || [];

                for (const gd of gdList) {
                    const name = gd.$?.name;
                    const fmla = gd.$?.fmla || "";
                    let val = parseInt(fmla.replace("val ", ""), 10);

                    const DEFAULT_ADJ1 = 50000;
                    const DEFAULT_ADJ2 = 50000;

                    if (name === "adj1") adj1 = val || DEFAULT_ADJ1;
                    if (name === "adj2") adj2 = val || DEFAULT_ADJ2;
                }

                // --- 2. Convert to percentages (PPT uses 0–100000) ---
                const shaftHeightPct = (adj1 / 100000) * 100;
                const headWidthPct = (adj2 / 100000) * 100;

                // --- 3. Compute polygon points in percentage space ---
                const shaftTopPct = Math.round((100 - shaftHeightPct) / 2);
                const shaftBottomPct = Math.round(shaftTopPct + shaftHeightPct);
                const headStartPct = Math.round(100 - headWidthPct);

                // --- 4. Dynamic clip-path that matches PPT rightArrow geometry ---
                clipPath = `polygon(
                        0% ${shaftTopPct}%,
                        ${headStartPct}% ${shaftTopPct}%,
                        ${headStartPct}% 0%,
                        100% 50%,
                        ${headStartPct}% 100%,
                        ${headStartPct}% ${shaftBottomPct}%,
                        0% ${shaftBottomPct}%
                    )`;

                break;
            }

            case "leftArrow":
                caseName = "leftArrow";
                clipPath = "polygon(40% 0%, 40% 20%, 100% 20%, 100% 80%, 40% 80%, 40% 100%, 0% 50%)";
                break;
            case "upArrow":
                caseName = "upArrow";
                clipPath = "polygon(40% 100%, 40% 30%, 20% 30%, 50% 0%, 80% 30%, 60% 30%, 60% 100%)";
                break;
            case "downArrow":
                caseName = "downArrow";
                clipPath = "polygon(40% 0%, 40% 70%, 20% 70%, 50% 100%, 80% 70%, 60% 70%, 60% 0%)";
                break;
            case "leftRightArrow":
                caseName = "leftRightArrow";
                clipPath = " polygon( 0% 50%,20% 30%,20% 40%,80% 40%, 80% 30%, 100% 50%,80% 70%,80% 60%,20% 60%, 20% 70%);"
                break;
            case "quadArrow":
                caseName = "quadArrow";
                clipPath = " polygon( 0% 50%, 25% 30%, 25% 40%, 40% 40%,40% 20%,30% 20%, 50% 0%, 70% 20%, 60% 20%, 60% 40%, 75% 40%, 75% 30%, 100% 50%, 75% 70%, 75% 60%, 60% 60%, 60% 80%, 70% 80%, 50% 100%, 30% 80%, 40% 80%, 40% 60%, 25% 60%, 25% 70% );"
                break;
            case "quadArrowCallout":
                caseName = "quadArrowCallout";
                clipPath = "polygon(0% 50%, 25% 30%, 25% 40%, 31.67% 40%, 31.67% 30%, 40% 30%, 40% 20%, 30% 20%, 50% 0%, 70% 20%, 60% 20%, 60% 30%, 67.5% 30%, 67.5% 40%, 75% 40%, 75% 30%, 100% 50%, 75% 70%, 75% 60%, 67.5% 60%, 67.5% 70%, 60% 70%, 60% 80%, 70% 80%, 50% 100%, 30% 80%, 40% 80%, 40% 70%, 32.5% 70%, 32.5% 60%, 25% 60%, 25% 70%)";
                break;
            case "leftRightArrowCallout":
                caseName = "leftRightArrowCallout";
                clipPath = " polygon(0% 50%, 25% 30%, 25% 40%, 31.67% 40%, 31.67% 15.25%, 67.5% 15.25%, 67.5% 40%, 75% 40%, 75% 30%, 100% 50%, 75% 70%, 75% 60%, 67.5% 60%, 67.5% 84.17%, 32.5% 84.17%, 32.5% 60%, 25% 60%, 25% 70%);"
                break;
            case "upDownArrow":
                caseName = "upDownArrow";
                clipPath = "polygon( 50% 0%,70% 20%, 60% 20%, 60% 45%, 80% 45%,50% 75%,20% 45%, 40% 45%, 40% 20%,  30% 20%,50% 0%,50% 75%,  80% 45%,  60% 45%, 60% 80%,  70% 80%, 50% 100%,30% 80%, 40% 80%,  40% 45%, 20% 45%,50% 75%)";
                break;
            case "chevron":
                caseName = "chevron";
                clipPath = "polygon(75% 0%, 100% 50%, 75% 100%, 0% 100%, 25% 50%, 0% 0%);"
                break;
            case "rightArrowCallout":
                caseName = "rightArrowCallout";
                clipPath = "polygon(82% 45.93%, 82% 37.51%, 100% 50%, 82% 62.49%, 82% 54.36%, 69.44% 54.36%, 69.44% 69.74%, 33.94% 69.74%, 33.94% 31.94%, 69.44% 31.94%, 69.44% 45.93%);"
                break;
            case "homePlate":
                caseName = "homePlate";
                const adjtValue = 95000;
                const arrowStart = (adjtValue / 100000) * 100; // converts to 95%
                clipPath = `polygon(0% 0%, ${arrowStart}% 0%, 100% 50%, ${arrowStart}% 100%, 0% 100%);`
                break;

            case "notchedRightArrow":
                caseName = "notchedRightArrow";
                clipPath = "polygon(28.76% 60.98%, 37.43% 69.65%, 28.76% 78.32%, 78.1% 78.05%, 78.98% 86.06%, 88.94% 68.99%, 78.98% 53.31%, 78.1% 60.98%);"
                break;
            case "leftRightUpArrow":
                caseName = "leftRightUpArrow";
                clipPath = "polygon(28.54% 53.31%, 18.58% 68.99%, 18.58% 70.03%, 28.76% 86.06%, 29.42% 86.06%, 29.42% 78.05%, 72.12% 77.70%, 78.10% 78.05%, 78.10% 85.71%, 78.98% 86.06%, 80.09% 84.32%, 80.09% 83.62%, 80.97% 82.23%, 81.42% 82.23%, 82.52% 80.49%, 82.52% 79.79%, 82.96% 79.79%, 83.19% 78.75%, 83.85% 78.40%, 84.07% 77.35%, 84.51% 77.35%, 84.51% 76.66%, 86.73% 73.17%, 87.17% 73.17%, 87.61% 71.78%, 88.05% 71.78%, 88.94% 70.38%, 88.94% 68.99%, 78.98% 53.31%, 78.10% 53.66%, 78.10% 60.98%, 59.29% 60.98%, 59.29% 39.02%, 63.94% 39.02%, 64.38% 37.98%, 54.20% 21.95%, 53.32% 21.95%, 43.14% 38.33%, 43.58% 39.02%, 48.23% 39.02%, 48.23% 60.98%, 29.42% 60.98%, 29.42% 53.66%, 28.54% 53.31%);"
                break;
            case "leftUpArrow":
                caseName = "leftUpArrow";
                clipPath = "polygon(37.58% 51.64%, 37.58% 52.24%, 25.26% 69.85%, 25.26% 70.75%, 37.78% 88.96%, 38.60% 88.96%, 38.60% 80.00%, 69.40% 80.00%, 69.82% 79.10%, 69.82% 34.93%, 75.36% 35.22%, 75.98% 34.63%, 75.98% 34.03%, 70.84% 26.57%, 70.43% 26.57%, 69.20% 24.78%, 69.20% 24.18%, 68.58% 23.88%, 63.86% 17.01%, 63.45% 15.82%, 62.63% 15.82%, 50.10% 34.03%, 50.31% 35.22%, 56.26% 34.93%, 56.26% 60.60%, 38.60% 60.60%, 38.60% 51.64%, 37.58% 51.64%);"
                break;
            case "bentUpArrow":
                caseName = "bentUpArrow";
                clipPath = "polygon(15.94% 65.33%, 15.70% 84.67%, 16.18% 85.33%, 44.69% 84.67%, 68.12% 85.33%, 68.60% 84.67%, 68.60% 29.33%, 75.12% 29.33%, 75.60% 29.00%, 75.60% 28.33%, 61.84% 9.33%, 60.87% 9.33%, 47.34% 28.00%, 47.58% 29.33%, 54.35% 29.33%, 54.11% 65.33%, 15.94% 65.33%)";
                break;
            case "diamond":
                caseName = "diamond";
                clipPath = "polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)";
                break;

            case "frame":
                // Get adjustment value for frame
                const avLst = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["a:avLst"]?.[0];
                const adj1 = avLst?.["gd"]?.[0]?.["$"]?.fmla || "val 4278"; // Default or extract value
                const adjValue = parseInt(adj1.replace("val ", ""), 10); // 4278

                // Dynamically get shape dimensions from XML
                // const xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
                const shapeWidth = parseInt(xfrm?.["a:ext"]?.[0]?.["$"]?.cx || xfrm?.["ext"]?.[0]?.["$"]?.cx || "12192000", 10);
                const shapeHeight = parseInt(xfrm?.["a:ext"]?.[0]?.["$"]?.cy || xfrm?.["ext"]?.[0]?.["$"]?.cy || "6858000", 10);

                const thicknessPercent = adjValue / 100000; // 4278/100000 = 0.04278 (4.278%)

                // Calculate the aspect ratio to determine correct horizontal thickness
                const aspectRatio = shapeWidth / shapeHeight;

                // The absolute EMU thickness is the same on all sides
                const thicknessEMUs = shapeHeight * thicknessPercent;

                const verticalThicknessPct = thicknessPercent * 100; // 4.278%
                const horizontalThicknessPct = (thicknessEMUs / shapeWidth) * 100; // 2.406% for aspect ratio 1.78

                backgroundStyle = fillColor;
                const gradFill = shapeNode?.["p:spPr"]?.[0]?.["a:gradFill"]?.[0];
                if (gradFill) {
                    const gsLst = gradFill?.["gsLst"]?.[0]?.["gs"] || [];
                    if (gsLst.length >= 2) {
                        const pos1 = gsLst[0]?.["$"]?.pos || "81000";
                        const pos2 = gsLst[1]?.["$"]?.pos || "0";
                        const color1 = gsLst[0]?.["srgbClr"]?.[0]?.["$"]?.val || "6F8A9A";
                        const color2 = gsLst[1]?.["srgbClr"]?.[0]?.["$"]?.val || "526876";

                        const pos1Percent = parseInt(pos1, 10) / 1000; // 81%
                        const pos2Percent = parseInt(pos2, 10) / 1000; // 0%

                        const angle = gradFill?.["lin"]?.[0]?.["$"]?.ang || "16200000";
                        const cssAngle = (parseInt(angle, 10) / 60000) % 360; // 270deg

                        backgroundStyle = `linear-gradient(${cssAngle}deg, #${color1} ${pos1Percent}%, #${color2} ${pos2Percent}%)`;
                    }
                }

                caseName = "frame";

                // Calculate the inner boundaries
                const leftEdge = horizontalThicknessPct;
                const rightEdge = 100 - horizontalThicknessPct;
                const topEdge = verticalThicknessPct;
                const bottomEdge = 100 - verticalThicknessPct;

                // Create polygon path for frame shape with visually consistent thickness
                clipPath = `polygon(
                        0% 0%,                      
                        100% 0%,                     
                        100% 100%,                 
                        0% 100%,                   
                        0% 0%,                       
                        ${leftEdge}% ${topEdge}%,    
                        ${leftEdge}% ${bottomEdge}%, 
                        ${rightEdge}% ${bottomEdge}%, 
                        ${rightEdge}% ${topEdge}%,  
                        ${leftEdge}% ${topEdge}%  
                    )`;
                break;

            case "halfFrame":
                caseName = "halfFrame";
                clipPath = "polygon(0% 0%, 80% 0%, 70% 20%, 40% 20%, 40% 80%, 0% 100%, 0% 80%);"
                break;


            case "diagStripe":
                caseName = "diagStripe";
                clipPath = "polygon(evenodd, 67% 0%, 100% 0%, 0% 100%, 0% 67%);"
                break;
            case "corner":
                caseName = "corner";
                clipPath = "polygon(10% 10%, 35.75% 10%, 35.75% 67.25%, 90% 67.25%, 90% 90%, 10% 90%);"
                break;
            case "plus":
                caseName = "plus";
                clipPath = "polygon(35% 0%,65% 0%,65% 35%,100% 35%,100% 60%,65% 60%,65% 100%,35% 100%,35% 60%,0% 60%,0% 35%,35% 35%)";
                break;
            case "minus":
                caseName = "minus";
                clipPath = "polygon(100% 35%,100% 60%,0% 60%,0% 35%);"
                break;
            case "close":
                caseName = "close";
                clipPath = "polygon(20% 0%, 0% 20%, 30% 50%, 0% 80%, 20% 100%, 50% 70%, 80% 100%, 100% 80%, 70% 50%, 100% 20%, 80% 0%, 50% 30%);"
                break;
            // case "roundRect":
            //     caseName = "roundRect";
            //     clipPath = "inset(0% round 2%);";
            //     break;

            case "roundRect":
                caseName = "roundRect";
                cornerRadius = this.getCornerRadius(shapeNode); // Get the radius in pixels
                const adjustedRadius = Math.min(cornerRadius, position.width / 2, position.height / 2); // Ensure radius doesn't exceed half the width/height
                clipPath = `inset(0% round ${adjustedRadius}px ${adjustedRadius}px ${adjustedRadius}px ${adjustedRadius}px)`; // Apply uniform rounding
                borderRadius = `${adjustedRadius}px`; // Sync border-radius with clip-path
                break;
            case "round2SameRect":
                caseName = "round2SameRect";
                clipPath = "inset(0% 0% 0% 0% round 20% 20% 0% 0%);"
                break;
            case "round1Rect":
                caseName = "round1Rect";
                clipPath = "inset(0% 0% 0% 0% round 0% 20% 0% 0%);"
                break;
            case "round2DiagRect":
                caseName = "round2DiagRect";
                clipPath = "inset(0% 0% 0% 0% round 25px 0 25px 0);"
                break;
            case "snip1Rect":
                caseName = "snip1Rect";
                clipPath = "polygon(0% 0%, 90% 0%, 100% 10%, 100% 100%, 0% 100%, 0% 0%);"
                break;
            case "snip2SameRect":
                caseName = "snip2SameRect";
                clipPath = "polygon(10% 0%, 90% 0%, 100% 10%, 100% 100%, 0% 100%, 0% 10%);"
                break;
            case "snip2DiagRect":
                caseName = "snip2DiagRect";
                clipPath = "polygon(0% 40.5%, 0% 0%, 88.75% 0%, 100% 11.25%, 100% 100%, 11.25% 100%, 0% 86%);"
                break;
            case "bentArrow":
                return `<div class="shape" id="bentArrow" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height}px;
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};
                                    ">
                                <svg width="${position.width}" height="${position.height}" viewBox="405 1685 573 540" xmlns="http://www.w3.org/2000/svg">
                            <path d="
                                M408.5 2221.5 
                                V1990.5 
                                C408.5 1862.92 511.922 1759.5 639.499 1759.5 
                                H840.5 
                                V1693.5 
                                L972.5 1825.5 
                                L840.5 1957.5 
                                V1891.5 
                                H639.499 
                                C584.823 1891.5 540.5 1935.82 540.5 1990.5 
                                V2221.5 
                                Z
                            " 
                            stroke="#042433" 
                            stroke-width="2" 
                            fill="${fillColor}" 
                            fill-rule="evenodd"/>
                        </svg>

                    </div>`;
                break;



            case "uturnArrow":
                return `<div class="shape" id="uturnArrow" data-name="${shapeName}" style="
                        position: absolute;
                        left: ${position.x}px;
                        top: ${position.y}px;
                        width: ${position.width}px;
                        height: ${position.height}px;
                        ${opacity};
                        transform: rotate(${position.rotation}deg);
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        z-index: ${zIndex};
                    ">
                    
                        <svg width="${position.width}" height="${position.height}" 
                            viewBox="1874 1296 300 798" 
                            xmlns="http://www.w3.org/2000/svg">
                                
                            <path d="M1877.5 2090.5 1877.5 1425.5C1877.5 1355.91 1933.91 1299.5 2003.5 1299.5L2003.5 1299.5C2073.09 1299.5 2129.5 1355.91 2129.5 1425.5L2129.5 1820.75 2165.5 1820.75 2093.5 1892.75 2021.5 1820.75 2057.5 1820.75 2057.5 1425.5C2057.5 1395.68 2033.32 1371.5 2003.5 1371.5L2003.5 1371.5C1973.68 1371.5 1949.5 1395.68 1949.5 1425.5L1949.5 2090.5Z" 
                                stroke="#042433" stroke-width="2" stroke-miterlimit="8" fill="${fillColor}" fill-rule="evenodd"/>
                                
                        </svg> 
                    </div>`;
                break;
            case "curvedRightArrow":
                return `<div class="shape" id="curvedRightArrow" data-name="${shapeName}" style="
                        position: absolute;
                        left: ${position.x}px;
                        top: ${position.y}px;
                        width: ${position.width}px;
                        height: ${position.height}px;
                        ${opacity};
                        transform: rotate(${position.rotation}deg);
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        z-index: ${zIndex};">
                    
                        <svg width="${position.width}" height="${position.height}" 
                            viewBox="2314 895 321 392" 
                            xmlns="http://www.w3.org/2000/svg">
                                
                            <!-- Main Curved Arrow Path -->
                            <path d="
                                M2317.5 1032  
                                C2317.5 1092.88 2413.75 1146.04 2551.5 1161.26  
                                L2551.5 1122.26  
                                L2629.5 1204.5  
                                L2551.5 1278.26  
                                L2551.5 1239.26  
                                C2413.75 1224.04 2317.5 1170.88 2317.5 1110  
                                Z
                            " 
                            fill="${fillColor}" fill-rule="evenodd"/>
                            
                            <!-- Upper Curve -->
                            <path d="
                                M2629.5 976.5  
                                C2492.29 976.5 2371.19 1014.85 2331.11 1071  
                                C2280.77 1000.49 2373.56 925.863 2538.35 904.324  
                                C2567.89 900.463 2598.61 898.5 2629.5 898.5  
                                Z
                            " 
                            fill="${fillColor}" fill-rule="evenodd"/>
                            
                            <!-- Outline -->
                            <path d="
                                M2317.5 1032  
                                C2317.5 1092.88 2413.75 1146.04 2551.5 1161.26  
                                L2551.5 1122.26  
                                L2629.5 1204.5  
                                L2551.5 1278.26  
                                L2551.5 1239.26  
                                C2413.75 1224.04 2317.5 1170.88 2317.5 1110  
                                L2317.5 1032  
                                C2317.5 958.27 2457.19 898.5 2629.5 898.5  
                                L2629.5 976.5  
                                C2492.29 976.5 2371.19 1014.85 2331.11 1071
                            " 
                            stroke="#042433" stroke-width="2" stroke-miterlimit="8" fill="none" fill-rule="evenodd"/>
                                
                        </svg> 
                    </div>`;
                break;


            case "donut":
                return `<div class="shape" id= "donut" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};
                                border-radius: 50%;
                                background: ${backgroundStyle && backgroundStyle !== 'transparent' ? backgroundStyle : 'gray'};">
                            <div style="
                                width: ${position.width * 0.6}px;
                                height: ${position.height * 0.6}px;
                                background: transparent;
                                border-radius: 50%;
                                background-color: white; /* Ensures hole is visible */
                            ">
                            </div>
                        </div>`;
                break;

            case "noSmoking":
                return `<div class="shape" id= "noSmoking" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};
                                border-radius: 50%;
                                background: ${fillColor};
                                border: 2px solid #000000; 
                                position: relative;">

                             <!-- Inner White Circle (Donut Hole) -->

                        <div style="
                            width: ${position.width * 0.6}px;
                            height: ${position.height * 0.6}px;
                            background: white;
                            border-radius: 50%;
                            border: 2px solid #000000;">
                        </div>

                            <!-- Diagonal Slash -->
                        <div style="
                            position: absolute;
                            width: ${position.width * 0.8}px;
                            height: ${position.height * 0.15}px;
                            background: ${fillColor};
                            transform: rotate(45deg);
                            border: 2px solid #000000; 
                            top: 50%;
                            left: 50%;
                            translate: -50% -50%;
                        ">
                        ${textContent}
                        </div>
                    </div>`;
                break;
            case "blockArc":
                return `<div class="shape" id= "blockArc" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height / 2}px; /* Half height for semi-circle */
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};
                            background: ${backgroundStyle && backgroundStyle !== 'transparent' ? backgroundStyle : 'gray'};
                            border-top-left-radius: ${position.width / 2}px;
                            border-top-right-radius: ${position.width / 2}px;
                            border-bottom-left-radius: 0;
                            border-bottom-right-radius: 0;
                            overflow: hidden;">

                                <div style="
                                    width: ${position.width * 0.6}px;
                                    height: ${position.width * 0.6}px;
                                    background: white;
                                    border-radius: 50%;
                                    position: absolute;
                                    bottom: -${position.width * 0.3}px; /* Adjusted to center the cutout */
                                ">
                                </div>
                        </div>`;
                break;
            case "heart":
                return `<div class="shape" id= "heart" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">

                                    <svg width="${position.width}" height="${position.height}" 
                                        viewBox="0 0 ${position.width} ${position.height}" 
                                        xmlns="http://www.w3.org/2000/svg">
                                        <path d="
                                            M${position.width * 0.5},${position.height * 0.9}
                                            C${position.width * 0.1},${position.height * 0.6} 
                                            ${position.width * 0.0},${position.height * 0.35}  
                                            ${position.width * 0.15},${position.height * 0.2} 
                                            C${position.width * 0.3},${position.height * 0.05} 
                                            ${position.width * 0.5},${position.height * 0.15} 
                                            ${position.width * 0.5},${position.height * 0.25} 
                                            C${position.width * 0.5},${position.height * 0.15} 
                                            ${position.width * 0.7},${position.height * 0.05} 
                                            ${position.width * 0.85},${position.height * 0.2} 
                                            C${position.width * 1.0},${position.height * 0.35} 
                                            ${position.width * 0.9},${position.height * 0.6} 
                                            ${position.width * 0.5},${position.height * 0.9}
                                        " 
                                        fill="${fillColor}" 
                                        stroke="#042433" stroke-width="${Math.max(1, position.width * 0.02)}" />
                                        ${textContent}
                                    </svg>
                            </div>`;
                break;

            case "foldedCorner":
                return `<div class="shape" id= "foldedCorner" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex}; ">

                                    <svg width="${position.width}" height="${position.height}" 
                                        viewBox="0 0 ${position.width} ${position.height}" 
                                        xmlns="http://www.w3.org/2000/svg">
                                        <path d="
                                            M0,0 
                                            L${position.width},0 
                                            L${position.width},${position.height - position.height * 0.15} 
                                            L${position.width - position.width * 0.1},${position.height} 
                                            L0,${position.height} 
                                            Z
                                            M${position.width},${position.height - position.height * 0.15} 
                                            Q${position.width - position.width * 0.1},${position.height - position.height * 0.1} 
                                            ${position.width - position.width * 0.1},${position.height}
                                            " 
                                        fill="${fillColor}" 
                                        stroke="#042433" stroke-width="${Math.max(1, position.width * 0.02)}" />
                                        ${textContent}
                                    </svg>
                            </div>`;
                break;
            case "cloud":
                return `<div class="shape" id= "cloud" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height}px;
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};">
                                <svg width="${position.width}" height="${position.height}" 
                                    viewBox="0 0 ${position.width} ${position.height}" 
                                    xmlns="http://www.w3.org/2000/svg">
                                    
                                    <!-- Fully Dynamic Cloud Shape -->
                                    <path d="
                                        M${position.width * 0.3},${position.height * 0.8}
                                        C${position.width * 0.1},${position.height * 0.9} 
                                        ${position.width * 0.05},${position.height * 0.65} 
                                        ${position.width * 0.1},${position.height * 0.5}
                                        C${position.width * 0.03},${position.height * 0.42} 
                                        ${position.width * 0.05},${position.height * 0.2} 
                                        ${position.width * 0.15},${position.height * 0.2}
                                        C${position.width * 0.15},${position.height * 0.05} 
                                        ${position.width * 0.25},${position.height * 0.02} 
                                        ${position.width * 0.35},${position.height * 0.15}
                                        C${position.width * 0.4},${position.height * 0} 
                                        ${position.width * 0.6},${position.height * 0} 
                                        ${position.width * 0.65},${position.height * 0.15}
                                        C${position.width * 0.75},${position.height * 0.05} 
                                        ${position.width * 0.9},${position.height * 0.15} 
                                        ${position.width * 0.87},${position.height * 0.35}
                                        C${position.width * 1},${position.height * 0.4} 
                                        ${position.width * 1},${position.height * 0.7} 
                                        ${position.width * 0.85},${position.height * 0.75}
                                        C${position.width * 0.8},${position.height * 0.9} 
                                        ${position.width * 0.6},${position.height * 0.9} 
                                        ${position.width * 0.5},${position.height * 0.8}
                                        C${position.width * 0.45},${position.height * 0.95} 
                                        ${position.width * 0.35},${position.height * 0.95} 
                                        ${position.width * 0.3},${position.height * 0.8}
                                        Z
                                    " 
                                    fill="${fillColor}" 
                                    stroke="#042433" stroke-width="${Math.max(2, position.width * 0.015)}" />
                                    ${textContent}
                                </svg>
                        </div>`;
                break;
            case "smileyFace":
                return `<div class="shape" id= "smileyFace" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height}px;
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};">
                                <svg width="${position.width}" height="${position.height}" 
                                    viewBox="0 0 100 100" 
                                    xmlns="http://www.w3.org/2000/svg">
                                    
                                    <!-- Face Circle -->
                                    <circle cx="50" cy="50" r="48" 
                                        fill="${fillColor}" 
                                        stroke="#042433" stroke-width="2" />
                                    
                                    <!-- Left Eye -->
                                    <circle cx="35" cy="35" r="4" fill="#4B2C3E" />
                                    
                                    <!-- Right Eye -->
                                    <circle cx="65" cy="35" r="4" fill="#4B2C3E" />
                                    
                                    <!-- Smile (Curved Line) -->
                                    <path d="
                                        M35,65 
                                        Q50,80 65,65
                                    " 
                                    stroke="#042433" stroke-width="2" fill="none" stroke-linecap="round"/>
                                    
                                </svg>
                        </div>`;
                break;


            case "snipRoundRect":
                const svgWidth = position.width;
                const svgHeight = position.height;
                cornerRadius = 20;
                const cutSize = svgWidth * 0.15;
                // const strokeColor = "#042433";
                // const strokeWidth = 3;

                return `<div class="shape" id= "snipRoundRect" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${svgWidth}px;
                                height: ${svgHeight}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">
                                
                                    <svg width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}" xmlns="http://www.w3.org/2000/svg">
                                        <path d="
                                            M ${cornerRadius} 0 
                                            H ${svgWidth - cutSize} 
                                            L ${svgWidth} ${cutSize} 
                                            V ${svgHeight} 
                                            H 0 
                                            V ${cornerRadius} 
                                            Q 0 0 ${cornerRadius} 0 
                                            Z" 
                                            stroke="${strokeColor}" 
                                            stroke-width="${strokeWidth}" 
                                            stroke-miterlimit="8" 
                                            fill="${fillColor}" 
                                            fill-rule="evenodd"/>
                                            ${textContent}
                                    </svg>
                            </div>`;
                break;
            case "chord":
                return `<div class="shape" id="chord" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height}px;
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};">
                            
                                <svg width="${position.width}" height="${position.height}" viewBox="0 0 ${position.width} ${position.height}" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M${position.width} ${0.92 * position.height} 
                                        C${0.7 * position.width} ${1.06 * position.height}, 
                                        ${0.3 * position.width} ${0.98 * position.height}, 
                                        ${0.1 * position.width} ${0.76 * position.height} 
                                        C${-0.08 * position.width} ${0.52 * position.height}, 
                                        ${0.012 * position.width} ${0.22 * position.height}, 
                                        ${0.34 * position.width} ${0.07 * position.height} 
                                        C${0.42 * position.width} ${0.02 * position.height}, 
                                        ${0.54 * position.width} ${0} , 
                                        ${0.66 * position.width} ${0} Z" 
                                        stroke="#042433" stroke-width="2" stroke-miterlimit="8" 
                                        fill="${fillColor}" fill-rule="evenodd"/>
                                        ${textContent}
                                </svg>
                        </div>`;
                break;
            case "teardrop":
                return `<div class="shape" id= "teardrop" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${position.width}px;
                                height: ${position.height}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">
                                
                                    <svg width="${position.width}" height="${position.height}" viewBox="0 0 ${position.width} ${position.height}" xmlns="http://www.w3.org/2000/svg">
                                        <path d="
                                            M${position.width * 0.5},0 
                                            L${position.width},0 
                                            L${position.width},${position.height * 0.5} 
                                            C${position.width},${position.height * 0.85}, 
                                            ${position.width * 0.85},${position.height}, 
                                            ${position.width * 0.5},${position.height} 
                                            C${position.width * 0.15},${position.height}, 
                                            0,${position.height * 0.85}, 
                                            0,${position.height * 0.5} 
                                            C0,${position.height * 0.15}, 
                                            ${position.width * 0.15},0, 
                                            ${position.width * 0.5},0 
                                            Z" 
                                            stroke="#042433" stroke-width="2" stroke-miterlimit="8" 
                                            fill="${fillColor}" fill-rule="evenodd"/>
                                            ${textContent}
                                    </svg>
                            </div>`;
                break;

            case "cube":
                const cubeWidth = position.width || 200;
                const cubeHeight = position.height || 200;

                const sideOffset = cubeWidth * 0.3;
                const topOffset = cubeHeight * 0.3;

                const frontX = sideOffset;
                const frontY = topOffset;

                // Adjust viewBox to ensure the entire cube fits inside
                const viewBoxWidth = cubeWidth + sideOffset * 2;
                const viewBoxHeight = cubeHeight + topOffset * 2;

                return `<div class="shape" id="cube" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${viewBoxWidth}px;
                                height: ${viewBoxHeight}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">
                                
                                    <svg width="100%" height="100%" viewBox="0 0 ${viewBoxWidth} ${viewBoxHeight}" xmlns="http://www.w3.org/2000/svg">
                                        <!-- Front face -->
                                        <rect x="${frontX}" y="${frontY}" width="${cubeWidth}" height="${cubeHeight}" fill="${fillColor}" stroke="#042433" stroke-width="2"/>

                                        <!-- Right face -->
                                        <path d="M${frontX + cubeWidth} ${frontY} 
                                                L${frontX + cubeWidth + sideOffset} ${frontY - topOffset} 
                                                L${frontX + cubeWidth + sideOffset} ${frontY + cubeHeight - topOffset} 
                                                L${frontX + cubeWidth} ${frontY + cubeHeight}Z" 
                                            fill="${fillColor}" stroke="#042433" stroke-width="1"/>

                                        <!-- Top face -->
                                        <path d="M${frontX} ${frontY} 
                                                L${frontX + sideOffset} ${frontY - topOffset} 
                                                L${frontX + cubeWidth + sideOffset} ${frontY - topOffset} 
                                                L${frontX + cubeWidth} ${frontY}Z" 
                                            fill="${fillColor}" stroke="#042433" stroke-width="1"/>

                                        <!-- Cube edges -->
                                        <path d="M${frontX} ${frontY} 
                                                L${frontX + sideOffset} ${frontY - topOffset} 
                                                L${frontX + cubeWidth + sideOffset} ${frontY - topOffset} 
                                                L${frontX + cubeWidth + sideOffset} ${frontY + cubeHeight - topOffset} 
                                                L${frontX + cubeWidth} ${frontY + cubeHeight} 
                                                L${frontX} ${frontY + cubeHeight} 
                                                L${frontX} ${frontY}Z"
                                            stroke="#042433" stroke-width="1" fill="none"/>
                                            ${textContent}
                                    </svg>
                            </div>`;
                break;

            case "can":
                const cylinderWidth = position.width || 200;
                const cylinderHeight = position.height || 400;
                const ellipseHeight = cylinderWidth * 0.3;

                const viewBoxWidth1 = cylinderWidth;
                const viewBoxHeight1 = cylinderHeight + ellipseHeight * 2;

                return `<div class="shape" id="can" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${viewBoxWidth1}px;
                                height: ${viewBoxHeight1}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">
                                
                                    <svg width="100%" height="100%" viewBox="0 0 ${viewBoxWidth1} ${viewBoxHeight1}" xmlns="http://www.w3.org/2000/svg">
                                        <!-- Top ellipse (Only border, no fill to prevent double stroke) -->
                                        <ellipse cx="${cylinderWidth / 2}" cy="${ellipseHeight}" rx="${cylinderWidth / 2}" ry="${ellipseHeight / 2}" 
                                                fill="none" stroke="#042433" stroke-width="2"/>

                                        <!-- Cylinder body (Fixed Path to Remove Straight Line at the Top) -->
                                        <path d="M0 ${ellipseHeight} 
                                                C0 ${ellipseHeight * 1.5}, ${cylinderWidth} ${ellipseHeight * 1.5}, ${cylinderWidth} ${ellipseHeight} 
                                                V${cylinderHeight} 
                                                C${cylinderWidth} ${cylinderHeight + ellipseHeight / 2}, 0 ${cylinderHeight + ellipseHeight / 2}, 0 ${cylinderHeight} 
                                                Z" 
                                            fill="none" stroke="#042433" stroke-width="2"/>

                                        <!-- Bottom ellipse (Properly aligned for realistic curvature) -->
                                        <ellipse cx="${cylinderWidth / 2}" cy="${cylinderHeight}" rx="${cylinderWidth / 2}" ry="${ellipseHeight / 2}" 
                                                fill="none" stroke="#042433" stroke-width="2"/>
                                                
                                        <text x="50%" y="${cylinderHeight / 2}" font-size="16px" text-anchor="middle" fill="black">${textContent}</text>
                                    </svg>
                            </div>`;
                break;

            case "pie":
                const pieWidth = position.width;
                const pieHeight = position.height;
                const pieRadius = Math.min(pieWidth, pieHeight) / 2;
                const pieStrokeColor = "#042433";
                const pieStrokeWidth = 2;
                const centerX = pieWidth / 2;
                const centerY = pieHeight / 2;

                return `<div class="shape" id= "pie" data-name="${shapeName}" style="
                                position: absolute;
                                left: ${position.x}px;
                                top: ${position.y}px;
                                width: ${pieWidth}px;
                                height: ${pieHeight}px;
                                ${opacity};
                                transform: rotate(${position.rotation}deg);
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                z-index: ${zIndex};">
                                
                                    <svg width="${pieWidth}" height="${pieHeight}" viewBox="0 0 ${pieWidth} ${pieHeight}" xmlns="http://www.w3.org/2000/svg">
                                        <path d="
                                            M ${centerX} ${centerY} 
                                            L ${pieWidth} ${centerY}
                                            A ${pieRadius} ${pieRadius} 0 0 1 ${centerX + pieRadius * Math.cos(Math.PI / 2)} ${centerY + pieRadius * Math.sin(Math.PI / 2)}
                                            A ${pieRadius} ${pieRadius} 0 0 1 ${centerX - pieRadius} ${centerY}
                                            A ${pieRadius} ${pieRadius} 0 0 1 ${centerX} ${centerY - pieRadius}
                                            Z"
                                            stroke="${pieStrokeColor}" 
                                            stroke-width="${pieStrokeWidth}" 
                                            stroke-linejoin="round"
                                            fill="${fillColor}" 
                                            fill-rule="evenodd"/>
                                            ${textContent}
                                    </svg>
                            </div>`;
                break;
            case "plaque":
                return `<div class="shape" id="plaque" data-name="${shapeName}" style="
                            position: absolute;
                            left: ${position.x}px;
                            top: ${position.y}px;
                            width: ${position.width}px;
                            height: ${position.height}px;
                            ${opacity};
                            transform: rotate(${position.rotation}deg);
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            z-index: ${zIndex};">
                            
                                <svg width="${position.width}" height="${position.height}" viewBox="0 0 ${position.width} ${position.height}" xmlns="http://www.w3.org/2000/svg">
                                    <path d="
                                        M3 ${position.height * 0.16} 
                                        C${position.width * 0.1} ${position.height * 0.16}, 
                                        ${position.width * 0.2} ${position.height * 0.07}, 
                                        ${position.width * 0.2} 3 
                                        L${position.width * 0.8} 3 
                                        C${position.width * 0.8} ${position.height * 0.07}, 
                                        ${position.width * 0.9} ${position.height * 0.16}, 
                                        ${position.width - 3} ${position.height * 0.16} 
                                        L${position.width - 3} ${position.height * 0.84} 
                                        C${position.width * 0.9} ${position.height * 0.84}, 
                                        ${position.width * 0.8} ${position.height * 0.93}, 
                                        ${position.width * 0.8} ${position.height - 3} 
                                        L${position.width * 0.2} ${position.height - 3} 
                                        C${position.width * 0.2} ${position.height * 0.93}, 
                                        ${position.width * 0.1} ${position.height * 0.84}, 
                                        3 ${position.height * 0.84} 
                                        Z" 
                                        stroke="#042433" stroke-width="2" fill="${fillColor}"/>
                                </svg>

                        </div>`;
                break;

            default:
                clipPath = "";
                maskPath = "";
                break;
        }

        const isTextBox = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvSpPr"]?.[0]?.["$"]?.txBox === "1";
        const transformString = this.getTransformString(position, isTextBox);

        const borderShape = shapeBorderStyle.border;

        return `<div class="shape" id="${caseName}" 
                data-name="${shapeName}" 
                data-original-color="${originalThemeColor}" 
                originalLumMod="${originalLumMod}" 
                originalLumOff="${originalLumOff}" 
                originalAlpha="${originalAlpha}"

                    style="
                    position: absolute;
                    left: ${position.x}px;
                    top: ${position.y}px;
                    width: ${position.width}px;
                    height: ${position.height}px;
                    background: ${fillColor};
                    ${opacity};
                    border-radius: ${borderRadius};
                    border: ${borderShape};
                    display: ${hidden ? "none" : "flex"};
                    transform:  ${transformString};                    
                    box-sizing: border-box;
                    overflow: hidden;
                    justify-content: ${shapeInfo.justifyContent};
                    align-items: ${shapeInfo.getAlignItem};
                    z-index: ${zIndex};
                    ${clipPath ? `clip-path: ${clipPath};` : ""}    
                    ${maskPath ? `mask: ${maskPath};` : ""}
                    
                ">
                ${textContent}
            </div>`;

    }

    async getImageFromPicture(node, slidePath, relationshipsXML, height) {
        try {

            const blipFillNode = node?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0] ||
                node?.["p:blipFill"]?.[0];
            const blip = blipFillNode?.["a:blip"]?.[0];

            const imageId = blip?.["$"]?.["r:embed"];
            if (!imageId) return null;

            const relationship = relationshipsXML?.["Relationships"]?.["Relationship"]?.find(
                (rel) => rel["$"].Id === imageId
            );

            if (!relationship) {
                console.warn(`No relationship found for image ID: ${imageId}`);
                return null;
            }

            const rawTarget = relationship["$"].Target;

            // Try multiple path variations to find the image
            const pathVariations = [
                // Variation 1: Handle paths starting with "../" (relative to ppt folder)
                path.posix.normalize(`ppt/${rawTarget.replace("../", "")}`),

                // Variation 2: Remove leading "/" and normalize
                path.posix.normalize(rawTarget.replace(/^\/+/, "")),

                // Variation 3: If it starts with "/ppt/", just remove the leading slash
                rawTarget.startsWith("/ppt/") ? rawTarget.substring(1) : null,

                // Variation 4: As-is without leading slashes
                rawTarget.replace(/^\/+/, ""),

                // Variation 5: Prepend "ppt/" if it doesn't start with "ppt/"
                !rawTarget.includes("ppt/") ? `ppt/${rawTarget.replace(/^\/+/, "")}` : null,

                // Variation 6: Try with just the filename in media folder
                `ppt/media/${path.basename(rawTarget)}`
            ].filter(Boolean); // Remove null values

            let imageFile = null;
            let targetPath = null;

            // Try each path variation
            for (const pathVariant of pathVariations) {

                if (this.extractor.files[pathVariant]) {
                    imageFile = this.extractor.files[pathVariant];
                    targetPath = pathVariant;
                    break;
                }
            }

            if (!imageFile) {
                console.warn("Image file not found in unzipped files. Tried paths:", pathVariations);
                console.warn("Available paths in extractor.files:", Object.keys(this.extractor.files).filter(k => k.includes('media') || k.includes('image')));
                return null;
            }

            // ?? Detect if image is an ellipse (used for masking)
            const shapeType = node?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0]?.["$"]?.prst || "rect";
            const isEllipse = shapeType === "ellipse";

            // ?? Extract border properties
            const lineNode = node?.["p:spPr"]?.[0]?.["a:ln"]?.[0];
            const borderWidth = lineNode?.["$"]?.w ? parseInt(lineNode["$"].w, 10) / this.getEMUDivisor() : 0;

            let borderColor = "black"; // Default border color
            if (lineNode?.["a:solidFill"]?.[0]?.["a:srgbClr"]) {
                borderColor = `#${lineNode["a:solidFill"][0]["a:srgbClr"][0]["$"].val}`;
            } else if (lineNode?.["a:solidFill"]?.[0]?.["a:schemeClr"]) {
                const schemeClrKey = lineNode["a:solidFill"][0]["a:schemeClr"][0]["$"].val;
                borderColor = colorHelper.resolveThemeColorHelper(schemeClrKey, this.extractor.themeXML, this.masterXML);
            }

            // ?? Extract Transparency (Opacity)
            // let opacity = 1;
            // const alphaNode = node?.["p:spPr"]?.[0]?.["a:solidFill"]?.[0]?.["a:alpha"];
            // if (alphaNode) {
            //     opacity = parseInt(alphaNode[0]["$"].val, 10) / 100000;
            // }

            // ?? Extract Image Transformations (Flip, Rotation)
            const xfrm = node?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
            let flipH = "";
            let flipV = "";
            let rotation = 0;

            if (xfrm?.["$"]?.flipH === "1") flipH = "scaleX(-1)";
            if (xfrm?.["$"]?.flipV === "1") flipV = "scaleY(-1)";
            if (xfrm?.["$"]?.rot) rotation = parseInt(xfrm["$"].rot, 10) / 60000;

            // ?? NEW: Comprehensive cropping extraction
            const cropData = this.extractImageCroppingComprehensive(blipFillNode);

            const position = this.getShapePosition(node);

            const croppingStyles = cropData ? this.generateComprehensiveCroppingStyles(cropData, position.height) : null;

            // ?? Save Image to Disk
            const imageBuffer = await imageFile.async("nodebuffer");
            const imageKey = `image_${Date.now()}_${path.basename(targetPath)}`;
            const filePath = path.join(imageSavePath, imageKey);

            fs.writeFileSync(filePath, imageBuffer);

            return {
                src: `${config.uploadPath}/${imageKey}`,
                shape: isEllipse ? "ellipse" : "rect",
                border: { width: borderWidth, color: borderColor },
                // opacity,
                transform: `${flipH} ${flipV} rotate(${rotation}deg)`,
                // ?? Enhanced cropping data
                cropping: cropData,
                croppingStyles: croppingStyles,
                hasCropping: cropData !== null
            };
        } catch (error) {
            console.error("Error in getImageFromPicture:", error);
            return null;
        }
    }

    generateComprehensiveCroppingStyles(cropData) {
        if (!cropData || !cropData.hasCropping) {
            return null;
        }

        const left = cropData.hasLeft ? cropData.left : 0;
        const top = cropData.hasTop ? cropData.top : 0;
        const right = cropData.hasRight ? cropData.right : 0;
        const bottom = cropData.hasBottom ? cropData.bottom : 0;

        let imageWidth = 100;
        let imageHeight = 100;
        let imageLeft = 0;
        let imageTop = 0;

        // Handle horizontal cropping/extension
        if (left !== 0 || right !== 0) {
            // Calculate the total width needed to show the required portion
            const leftExtension = Math.abs(Math.min(0, left));
            const rightExtension = Math.abs(Math.min(0, right));
            const leftCrop = Math.max(0, left);
            const rightCrop = Math.max(0, right);

            const visibleWidth = 100 - leftCrop - rightCrop;
            const totalRequiredWidth = visibleWidth + leftExtension + rightExtension;
            imageWidth = (totalRequiredWidth / visibleWidth) * 100;
            imageLeft = -(leftCrop + leftExtension) * (100 / visibleWidth);
        }

        // Handle vertical cropping/extension
        if (top !== 0 || bottom !== 0) {
            // Calculate the total height needed to show the required portion
            const topExtension = Math.abs(Math.min(0, top));
            const bottomExtension = Math.abs(Math.min(0, bottom));
            const topCrop = Math.max(0, top);
            const bottomCrop = Math.max(0, bottom);

            const visibleHeight = 100 - topCrop - bottomCrop;
            const totalRequiredHeight = visibleHeight + topExtension + bottomExtension;
            imageHeight = (totalRequiredHeight / visibleHeight) * 100;
            imageTop = -(topCrop + topExtension) * (100 / visibleHeight);
        }

        const containerStyles = `overflow: hidden;`;
        const imageStyles = `
        width: ${imageWidth.toFixed(2)}%;
        height: ${imageHeight.toFixed(2)}%;
        left: ${imageLeft.toFixed(2)}%;
        top: ${imageTop.toFixed(2)}%;
        object-fit: cover;
        position: absolute;
        `.replace(/\s+/g, ' ').trim();

        return {
            containerStyles,
            imageStyles,
            debug: {
                originalCrop: { left, top, right, bottom },
                imageTransform: {
                    width: imageWidth,
                    height: imageHeight,
                    left: imageLeft,
                    top: imageTop
                }
            }
        };
    }

    extractImageCroppingComprehensive(blipFillNode) {
        if (!blipFillNode) {
            return null;
        }

        // First check a:srcRect for standard cropping
        let srcRect = blipFillNode?.["a:srcRect"];
        let isFillRect = false;

        // If srcRect is empty or doesn't have attributes, check a:fillRect inside a:stretch
        if (!srcRect || !this.hasAttributes(srcRect)) {
            const stretch = blipFillNode?.["a:stretch"];
            if (stretch && stretch[0]?.["a:fillRect"]) {
                srcRect = stretch[0]["a:fillRect"][0];
                isFillRect = true;
            }
        } else if (Array.isArray(srcRect) && srcRect.length > 0) {
            // Handle array case
            srcRect = srcRect[0];
        }

        if (!srcRect || !this.hasAttributes(srcRect)) {
            return null;
        }

        const cropping = {
            hasLeft: false,
            hasTop: false,
            hasRight: false,
            hasBottom: false,
            left: 0,
            top: 0,
            right: 0,
            bottom: 0,
            leftRaw: null,
            topRaw: null,
            rightRaw: null,
            bottomRaw: null,
            isFillRect: isFillRect  // Flag to indicate if this is fillRect or srcRect
        };

        // Get attributes from srcRect/fillRect
        const attrs = srcRect.$ || srcRect;

        // Check each cropping value individually
        if (attrs.hasOwnProperty('l') && attrs.l !== undefined && attrs.l !== null) {
            cropping.hasLeft = true;
            cropping.leftRaw = parseInt(attrs.l, 10);
            cropping.left = (cropping.leftRaw / 100000) * 100;
        }

        if (attrs.hasOwnProperty('t') && attrs.t !== undefined && attrs.t !== null) {
            cropping.hasTop = true;
            cropping.topRaw = parseInt(attrs.t, 10);
            cropping.top = (cropping.topRaw / 100000) * 100;
        }

        if (attrs.hasOwnProperty('r') && attrs.r !== undefined && attrs.r !== null) {
            cropping.hasRight = true;
            cropping.rightRaw = parseInt(attrs.r, 10);
            cropping.right = (cropping.rightRaw / 100000) * 100;
        }

        if (attrs.hasOwnProperty('b') && attrs.b !== undefined && attrs.b !== null) {
            cropping.hasBottom = true;
            cropping.bottomRaw = parseInt(attrs.b, 10);
            cropping.bottom = (cropping.bottomRaw / 100000) * 100;
        }

        if (!cropping.hasLeft && !cropping.hasTop && !cropping.hasRight && !cropping.hasBottom) {
            return null;
        }

        cropping.hasCropping = true;
        cropping.croppingType = this.determineCroppingType(cropping);

        return cropping;
    }

    // Helper function to check if object has attributes
    hasAttributes(obj) {
        if (!obj) return false;

        // Check if it has $ property (xml2js format)
        if (obj.$ && typeof obj.$ === 'object') {
            const attrs = obj.$;
            return Object.keys(attrs).some(key =>
                ['l', 't', 'r', 'b'].includes(key) &&
                attrs[key] !== undefined &&
                attrs[key] !== null &&
                attrs[key] !== ''
            );
        }

        // Check direct properties
        return Object.keys(obj).some(key =>
            ['l', 't', 'r', 'b'].includes(key) &&
            obj[key] !== undefined &&
            obj[key] !== null &&
            obj[key] !== ''
        );
    }

    determineCroppingType(cropping) {
        const types = [];

        if (cropping.hasLeft) {
            types.push(cropping.left >= 0 ? "crop-left" : "extend-left");
        }
        if (cropping.hasTop) {
            types.push(cropping.top >= 0 ? "crop-top" : "extend-top");
        }
        if (cropping.hasRight) {
            types.push(cropping.right >= 0 ? "crop-right" : "extend-right");
        }
        if (cropping.hasBottom) {
            types.push(cropping.bottom >= 0 ? "crop-bottom" : "extend-bottom");
        }

        return types.join(", ");
    }


    extractStrokeProperties(shapeNode, themeXML) {
        const outline = shapeNode["p:spPr"]?.[0]?.["a:ln"]?.[0];

        let strokeWidth = 0; // Default stroke width
        let strokeColor = ""; // Default stroke color
        let strokeCap = "butt"; // Default line cap (CSS default)
        let strokeDashArray = ""; // Default dash pattern

        // Extracting the stroke width from the 'w' attribute and converting from EMUs to pixels
        const emuWidth = outline?.["$"]?.w;
        if (emuWidth) {
            strokeWidth = parseInt(emuWidth) / this.getEMUDivisor(); // Use existing method
        }

        // Extract line cap style
        const cap = outline?.["$"]?.cap;
        if (cap) {
            switch (cap) {
                case "rnd":
                    strokeCap = "round";
                    break;
                case "sq":
                    strokeCap = "square";
                    break;
                case "flat":
                default:
                    strokeCap = "butt";
                    break;
            }
        }

        // Extract dash pattern
        const dashType = outline?.["a:prstDash"]?.[0]?.["$"]?.val;
        if (dashType && dashType !== "solid") {
            switch (dashType) {
                case "dash":
                    strokeDashArray = "5, 5";
                    break;
                case "dot":
                    strokeDashArray = "2, 5";
                    break;
                case "dashDot":
                    strokeDashArray = "5, 5, 2, 5";
                    break;
                case "lgDash":
                    strokeDashArray = "10, 5";
                    break;
                case "lgDashDot":
                    strokeDashArray = "10, 5, 2, 5";
                    break;
                case "lgDashDotDot":
                    strokeDashArray = "10, 5, 2, 5, 2, 5";
                    break;
                case "sysDash":
                    strokeDashArray = "3, 3";
                    break;
                case "sysDot":
                    strokeDashArray = "1, 3";
                    break;
                case "sysDashDot":
                    strokeDashArray = "3, 3, 1, 3";
                    break;
                case "sysDashDotDot":
                    strokeDashArray = "3, 3, 1, 3, 1, 3";
                    break;
                default:
                    strokeDashArray = "";
                    break;
            }
        }

        // Extract stroke color from various color types
        const solidFill = outline?.["a:solidFill"]?.[0];

        if (solidFill) {
            // Handle direct RGB colors (srgbClr)
            if (solidFill["a:srgbClr"]) {
                const srgbNode = solidFill["a:srgbClr"][0];
                strokeColor = `#${srgbNode["$"].val}`;

                // Apply luminance modifiers if present (following existing pattern)
                const lumMod = srgbNode["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    const lumOff = srgbNode["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                    } else {
                        strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                    }
                }
            }
            // Handle scheme colors (schemeClr) - theme-based colors
            else if (solidFill["a:schemeClr"]) {
                const schemeNode = solidFill["a:schemeClr"][0];
                const schemeValue = schemeNode["$"]["val"];

                // Use existing colorHelper method (following pattern from getShapeFillColor)
                strokeColor = colorHelper.resolveThemeColorHelper(schemeValue, themeXML || this.themeXML, this.masterXML);

                // Apply luminance modifiers if present
                const lumMod = schemeNode["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    const lumOff = schemeNode["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                    } else {
                        strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                    }
                }
            }
            // Handle scRGB colors (scrgbClr) - extended RGB
            else if (solidFill["a:scrgbClr"]) {
                const scrgbNode = solidFill["a:scrgbClr"][0]["$"];
                const r = Math.round(parseFloat(scrgbNode.r || 0) * 255);
                const g = Math.round(parseFloat(scrgbNode.g || 0) * 255);
                const b = Math.round(parseFloat(scrgbNode.b || 0) * 255);
                strokeColor = `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
            }
            // Handle HSL colors (hslClr)
            else if (solidFill["a:hslClr"]) {
                const hslNode = solidFill["a:hslClr"][0]["$"];
                const h = parseInt(hslNode.hue || 0) / 60000; // Convert from 60000ths to degrees
                const s = parseInt(hslNode.sat || 0) / 100000; // Convert from 100000ths to percentage
                const l = parseInt(hslNode.lum || 0) / 100000; // Convert from 100000ths to percentage

                // Convert HSL to Hex using inline conversion
                const lVal = l / 100;
                const a = s * Math.min(lVal, 1 - lVal) / 100;
                const f = n => {
                    const k = (n + h / 30) % 12;
                    const color = lVal - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
                    return Math.round(255 * color).toString(16).padStart(2, '0');
                };
                strokeColor = `#${f(0)}${f(8)}${f(4)}`;
            }
            // Handle preset colors (prstClr)
            else if (solidFill["a:prstClr"]) {
                const presetValue = solidFill["a:prstClr"][0]["$"]["val"];
                // Use existing resolveColor method pattern if available in theme
                strokeColor = this.resolvePresetColorFromTheme(presetValue, themeXML || this.themeXML) || this.getPresetColorFallback(presetValue);
            }
            // Handle system colors (sysClr)
            else if (solidFill["a:sysClr"]) {
                const sysNode = solidFill["a:sysClr"][0]["$"];
                const lastClr = sysNode["lastClr"];
                if (lastClr) {
                    strokeColor = `#${lastClr}`;
                } else {
                    // Try to resolve from theme or use the class's existing resolveColor method
                    const sysValue = sysNode["val"];
                    strokeColor = this.resolveColor(sysValue) || "#000000";
                }
            }
        }

        return {
            width: strokeWidth,
            color: strokeColor,
            cap: strokeCap,
            dashArray: strokeDashArray
        };
    }

    resolvePresetColorFromTheme(presetValue, themeXML) {
        if (!themeXML) return null;

        try {
            // Try to find preset color in theme using existing pattern
            const themeRoot = themeXML["a:theme"]?.[0] || themeXML;
            const clrScheme = themeRoot["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];

            if (clrScheme) {
                // Look through theme colors for matching preset
                for (const [key, value] of Object.entries(clrScheme)) {
                    if (value[0] && value[0]["a:prstClr"] && value[0]["a:prstClr"][0]["$"]["val"] === presetValue) {
                        if (value[0]["a:srgbClr"]) {
                            return `#${value[0]["a:srgbClr"][0]["$"]["val"]}`;
                        }
                    }
                }
            }
        } catch (error) {
            console.warn('Error resolving preset color from theme:', error);
        }

        return null;
    }

    getPresetColorFallback(presetValue) {
        const commonPresets = {
            'black': '#000000',
            'white': '#FFFFFF',
            'red': '#FF0000',
            'green': '#008000',
            'blue': '#0000FF',
            'yellow': '#FFFF00',
            'cyan': '#00FFFF',
            'magenta': '#FF00FF'
        };

        return commonPresets[presetValue] || '#000000';
    }

    getZIndexForShape(shapeName) {

        if (this.zIndexMap) {
            if (!this.shapeNameOccurrences[shapeName]) {
                this.shapeNameOccurrences[shapeName] = 0;
            }

            const occurrence = this.shapeNameOccurrences[shapeName];
            const uniqueKey = occurrence === 0 ? shapeName : `${shapeName}__${occurrence}`;
            this.shapeNameOccurrences[shapeName]++;
            const zIndex = this.zIndexMap[uniqueKey] || 0;

            return zIndex;
        }

        const matchingNode = this.nodes.find(node => node.name === shapeName);
        return matchingNode ? matchingNode.id : 0;
    }

    findPlaceholderInLayout(placeholderType, layoutXML) {
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

    convertAnchorToFlexAlign(anchor) {
        switch (anchor) {
            case "ctr": return "center";
            case "b": return "flex-end";
            case "t": return "flex-start";
            case "just": return "space-between";
            case "dist": return "space-around";
            default: return "flex-start";
        }
    }


    // Master placeholder detection (unchanged)
    isMasterPlaceholder(shapeNode, masterXML) {
        if (!this.isPlaceholder(shapeNode)) {
            return false;
        }

        const shapeName = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;

        if (masterXML && shapeName) {
            const masterShapes = masterXML?.["p:sldMaster"]?.[0]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];

            const masterShape = masterShapes.find(shape => {
                const masterShapeName = shape?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
                return masterShapeName === shapeName;
            });

            if (masterShape) {
                return true;
            }
        }

        if (shapeName) {
            const masterPlaceholderNames = [
                'Title Placeholder 1',
                'Text Placeholder 2',
                'Date Placeholder 3',
                'Footer Placeholder 4',
                'Slide Number Placeholder 5'
            ];

            if (masterPlaceholderNames.includes(shapeName)) {
                return true;
            }
        }

        return false;
    }

    isOnlyPlaceholderText(textContent) {
        if (!textContent) return true; // Empty is considered placeholder

        const cleanText = textContent
            .replace(/<[^>]*>/g, '') // Remove HTML tags
            .replace(/&nbsp;/g, ' ') // Replace &nbsp; with space
            .replace(/\s+/g, ' ') // Normalize whitespace
            .toLowerCase()
            .trim();

        if (cleanText === '') return true; // Empty after cleanup

        // EXACT matches for common PowerPoint placeholder texts
        const exactPlaceholderTexts = [
            'click to add title',
            'click to add subtitle',
            'click to add text',
            'click to add content',
            'click to add notes',
            'click to add first slide title',
            'slide title',
            'title',
            'subtitle'
        ];

        // Check for exact matches only
        return exactPlaceholderTexts.includes(cleanText);
    }

    // Returns true if a line is just a default placeholder prompt
    isPromptLine(line) {
        if (!line) return false;
        const prompts = [
            "click to add text",
            "click to add title",
            "click to add subtitle",
            "click to add content",
            "click to add notes",
            "insert picture here",
            "click icon to add picture",
            "title",
            "subtitle"
        ];
        return prompts.includes(String(line).trim().toLowerCase());
    }

    // Remove placeholder-prompt lines but keep any user text
    stripPlaceholderPrompts(rawHtmlOrText) {
        if (!rawHtmlOrText) return "";

        // Master slide prompts to detect
        const masterPrompts = [
            "click to edit master title style",
            "click to edit master text styles",
            "click to edit master subtitle style",
            "second level",
            "third level",
            "fourth level",
            "fifth level"
        ];

        // Regular slide prompts  
        const slidePrompts = [
            "click to add text",
            "click to add title",
            "click to add subtitle",
            "click to add content",
            "click to add notes",
            "insert picture here",
            "click icon to add picture"
        ];

        // Extract ONLY the text content for checking (without modifying original HTML)
        const textOnly = String(rawHtmlOrText)
            .replace(/<[^>]*>/g, '') // Remove HTML tags for analysis only
            .replace(/&nbsp;/g, ' ')
            .replace(/\s+/g, ' ')
            .toLowerCase()
            .trim();

        // Check if content is ONLY placeholder prompts
        const isOnlyMasterPrompts = masterPrompts.some(prompt =>
            textOnly.includes(prompt.toLowerCase())
        );

        const isOnlySlidePrompts = slidePrompts.some(prompt =>
            textOnly === prompt.toLowerCase()
        );

        // If it's only placeholder prompts, return empty (to skip rendering)
        if (isOnlyMasterPrompts || isOnlySlidePrompts) {
            return "";
        }

        // Otherwise, return the ORIGINAL HTML completely unchanged
        return rawHtmlOrText;
    }

    hasMeaningfulText(text) {
        if (text == null || text === "") return false;

        // Extract text content for analysis only (don't modify original)
        const textContent = String(text)
            .replace(/<[^>]*>/g, "")
            .replace(/&nbsp;/g, "")
            .trim();

        return textContent.length > 0;
    }

    // IMPROVED: Calculate proper list indentation
    calculateListIndentation(bulletInfo) {
        // Convert EMUs to pixels if marginLeft is available
        if (bulletInfo.marginLeft) {
            // PowerPoint EMUs to pixels: 1 inch = 914400 EMUs, 1 inch = 72 pixels (at 72 DPI)
            const pixels = Math.round((bulletInfo.marginLeft / 914400) * 72);
            return Math.max(0, pixels);
        }

        // Calculate based on indent level
        if (bulletInfo.level > 0) {
            return bulletInfo.level * 20; // 20px per level
        }

        return 0;
    }

    // IMPROVED: Better unordered list styling
    getUnorderedListStyle(bulletInfo) {
        const standardStyles = {
            'disc': 'list-style-type: disc',
            'circle': 'list-style-type: circle',
            'square': 'list-style-type: square'
        };

        if (standardStyles[bulletInfo.listStyle]) {
            return standardStyles[bulletInfo.listStyle];
        }

        // For custom bullets that CSS doesn't support, we'll need custom styling
        const customStyles = {
            'triangle': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpolygon points=\'0,0 8,4 0,8\' fill=\'black\'/%3E%3C/svg%3E")',
            'diamond': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpolygon points=\'4,0 8,4 4,8 0,4\' fill=\'black\'/%3E%3C/svg%3E")',
            'checkmark': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M1,4 L3,6 L7,2\' stroke=\'black\' stroke-width=\'1\' fill=\'none\'/%3E%3C/svg%3E")',
            'arrow': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M1,4 L6,4 M4,2 L6,4 L4,6\' stroke=\'black\' stroke-width=\'1\' fill=\'none\'/%3E%3C/svg%3E")',
            'star': 'list-style-type: none; list-style-image: url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'8\' height=\'8\'%3E%3Cpath d=\'M4,0 L5,3 L8,3 L6,5 L7,8 L4,6 L1,8 L2,5 L0,3 L3,3 Z\' fill=\'black\'/%3E%3C/svg%3E")'
        };

        return customStyles[bulletInfo.listStyle] || 'list-style-type: disc';
    }


    determineListTypeFromChar(char) {
        if (['�', '?', '?', '?'].includes(char)) {
            return 'ul'; // Standard bullet characters
        }
        return 'ul'; // Default to unordered list for all other characters
    }

    determineListTypeFromAutoNum(autoNum) {
        switch (autoNum) {
            case 'arabicPeriod':
                return 'ol'; // '1.', '2.', '3.', ...
            case 'alphaLcPeriod':
                return 'ol type="a"'; // 'a.', 'b.', 'c.', ...
            case 'alphaUcPeriod':
                return 'ol type="A"'; // 'A.', 'B.', 'C.', ...
            case 'romanLcPeriod':
                return 'ol type="i"'; // 'i.', 'ii.', 'iii.', ...
            case 'romanUcPeriod':
                return 'ol type="I"'; // 'I.', 'II.', 'III.', ...
            default:
                return 'ul';
        }
    }

    getCornerRadius(shapeNode) {
        const geom = shapeNode?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0];

        if (!geom || geom?.["$"]?.prst !== "roundRect") return 0; // Return 0 for non-roundRect shapes

        const adj = geom?.["a:avLst"]?.[0]?.["a:gd"]?.[0]?.["$"]?.fmla;
        if (adj) {
            const rawRadius = parseFloat(adj.replace("val ", "")) / 1000; // Convert from PowerPoint units (1/1000th of EMU)
            const position = this.getShapePosition(shapeNode);

            const scaleFactor = Math.min(position.width, position.height) / 100; // Scale relative to smaller dimension
            return Math.round(Math.max(5, rawRadius * scaleFactor)); // Minimum 5px, scaled based on shape size
        }

        return 12; // Default rounded corner radius if no adjustment is found
    }

    getStrokeProperties(outline, themeXML, defaultWidth = 2, defaultColor = "#000000") {
        let strokeWidth = defaultWidth;
        let strokeColor = defaultColor;


        if (outline) {
            // Extract Stroke Width (Convert from EMUs to px)
            if (outline.$?.w) {
                strokeWidth = parseInt(outline.$.w, 10) / this.getEMUDivisor();
            }

            // Extract Stroke Color
            if (outline['a:solidFill']) {
                const solidFill = outline['a:solidFill'][0];

                if (solidFill['a:srgbClr']) {
                    strokeColor = `#${solidFill['a:srgbClr'][0].$.val}`;
                } else if (solidFill['a:schemeClr'] && typeof getThemeColor === "function") {
                    strokeColor = getThemeColor(themeXML, solidFill['a:schemeClr'][0].$.val) || defaultColor;
                }
            }
        }

        return { strokeWidth, strokeColor };
    }

    extractStrokePropertiesForBorder(outline, themeXML, masterXML = null) {
        let strokeWidth = 0;
        let strokeColor = "#000000";

        if (outline) {
            // Extract Stroke Width (Convert from EMUs to px)
            if (outline["$"]?.w) {
                strokeWidth = parseInt(outline["$"].w, 10) / this.getEMUDivisor();
            }

            // Extract Stroke Color
            if (outline["a:solidFill"]) {
                const solidFill = outline["a:solidFill"][0];
                if (solidFill["a:srgbClr"]) {
                    strokeColor = `#${solidFill["a:srgbClr"][0]["$"].val}`;
                    // Apply alpha if present
                    const alpha = solidFill["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                    if (alpha) {
                        const opacity = parseInt(alpha, 10) / 100000;
                        strokeColor = this.applyAlphaToColor(strokeColor, opacity);
                    }
                } else if (solidFill["a:schemeClr"]) {
                    const schemeClr = solidFill["a:schemeClr"][0]["$"].val;
                    strokeColor = colorHelper.resolveThemeColorHelper(schemeClr, themeXML, this.masterXML);
                    // Apply luminance modification
                    const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                    if (lumMod) {
                        const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                        if (lumOff) {
                            strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff); // Assuming this method exists
                        } else {
                            strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                        }
                    }
                    // Apply luminance offset if present


                    // Apply alpha if present
                    const alpha = solidFill["a:schemeClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                    if (alpha) {
                        const opacity = parseInt(alpha, 10) / 100000;
                        strokeColor = this.applyAlphaToColor(strokeColor, opacity);
                    }
                }
            }
        }

        return { width: strokeWidth, color: strokeColor };
    }

    // Helper method to apply alpha to a color
    applyAlphaToColor(color, alpha) {
        const rgb = this.hexToRGB(color).split(",").map(Number);
        return `rgba(${rgb[0]}, ${rgb[1]}, ${rgb[2]}, ${alpha})`;
    }

    getShapePosition(shapeNode, masterXML = null) {

        let xfrm = shapeNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];

        // NEW: If xfrm is missing or incomplete, try to get from master XML
        if (masterXML && (!xfrm || !xfrm?.["a:off"]?.[0] || !xfrm?.["a:ext"]?.[0])) {
            const placeholderType = this.getPlaceholderType(shapeNode);
            const masterPosition = this.getPositionFromMaster(masterXML, shapeNode, placeholderType);

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
        const width = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cx || 100) / this.getEMUDivisor());
        const height = Math.round((xfrm?.["a:ext"]?.[0]?.["$"]?.cy || 100) / this.getEMUDivisor());

        return {
            x: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.x || 0) / this.getEMUDivisor()),
            y: Math.round((xfrm?.["a:off"]?.[0]?.["$"]?.y || 0) / this.getEMUDivisor()),
            width: width < 1 ? 1 : width,
            height: height < 1 ? 1 : height,
            rotation: xfrm?.["$"]?.rot ? parseInt(xfrm["$"].rot, 10) / 60000 : 0,
            flipH: flipH,
            flipV: flipV,
            isTextBox: isTextBox
        };
    }

    // Helper: Get placeholder type from shape
    getPlaceholderType(shapeNode) {
        const phType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type;
        return phType || null;
    }

    // Helper: Get position and size from master XML by matching placeholder in <p:spTree>
    getPositionFromMaster(masterXML, shapeNode, placeholderType) {
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

    // When generating CSS transform, use this logic:
    // getTransformString(position, isTextBox = false) {
    //     let transforms = [];

    //     if (position.rotation !== 0) {
    //         transforms.push(`rotate(${position.rotation}deg)`);
    //     }

    //     if (position.flipH) {
    //         transforms.push('scaleX(-1)');
    //     }

    //     if (position.flipV) {
    //         transforms.push('scaleY(-1)');
    //     }

    //     return transforms.length > 0 ? transforms.join(' ') : 'none';
    // }

    // Rakesh Notes::: here above is the original function i have change it 
    getTransformString(position, isTextBox = false) {
        const transforms = [];

        if (!isTextBox) {
            if (position.flipH) transforms.push('scaleX(-1)');
            if (position.flipV) transforms.push('scaleY(-1)');
        }

        if (position.rotation && position.rotation !== 0) {
            transforms.push(`rotate(${position.rotation}deg)`);
        }

        return transforms.length ? transforms.join(' ') : 'none';
    }

    getShapeFillColor(shapeNode, themeXML, masterXML = null) {

        let fillColor = "transparent";
        let fillOpacity = 1;
        let strokeColor = "transparent";
        let strokeOpacity = 1.0;
        let originalAlpha = '';

        try {
            const shapeFill = shapeNode?.["p:spPr"]?.[0];
            let originalThemeColor = '', originalLumMod = '', originalLumOff = '';
            // let txtPhType = shapeNode?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"]?.type || '';

            // Handle Shape Fill (interior)
            const solidFill = shapeFill?.["a:solidFill"]?.[0];

            if (solidFill?.["a:srgbClr"]?.[0]?.["$"]?.val) {

                originalThemeColor = solidFill["a:srgbClr"][0]["$"].val;

                fillColor = `#${solidFill["a:srgbClr"][0]["$"].val}`;

                const lumMod = solidFill["a:srgbClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                if (lumMod) {
                    originalLumMod = lumMod;
                    const lumOff = solidFill["a:srgbClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        originalLumOff = lumOff;
                        fillColor = pptBackgroundColors.applyLuminanceModifier(fillColor, lumMod, lumOff);
                    } else {
                        fillColor = colorHelper.applyLumMod(fillColor, lumMod);
                    }
                }

                // Extract fill opacity
                const fillAlphaNode = solidFill["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                if (fillAlphaNode) {
                    const alphaValue = parseInt(fillAlphaNode, 10);
                    originalAlpha = alphaValue;

                    if (!isNaN(alphaValue)) {
                        fillOpacity = alphaValue / 100000;
                    }
                }

            } else if (solidFill?.["a:schemeClr"]?.[0]?.["$"]?.val) {
                let schmClr = solidFill["a:schemeClr"][0]["$"].val;

                originalThemeColor = schmClr;

                // NEW: Check if we need to resolve through master color mapping
                // if (masterXML && schmClr) {
                //     const resolvedColor = this.resolveMasterColor(schmClr, masterXML);
                //     if (resolvedColor) {
                //         schmClr = resolvedColor;
                //     }
                // }

                if (schmClr) {


                    fillColor = colorHelper.resolveThemeColorHelper(schmClr, themeXML, masterXML);
                }

                const lumMod = solidFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;

                if (lumMod) {
                    originalLumMod = lumMod;

                    const lumOff = solidFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                    if (lumOff) {
                        originalLumOff = lumOff;
                        fillColor = pptBackgroundColors.applyLuminanceModifier(fillColor, lumMod, lumOff);
                    } else {
                        fillColor = colorHelper.applyLumMod(fillColor, lumMod);
                    }
                }

                // Extract fill opacity for scheme color
                const fillAlphaNode = solidFill["a:schemeClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                if (fillAlphaNode) {
                    const alphaValue = parseInt(fillAlphaNode, 10);
                    originalAlpha = alphaValue;
                    if (!isNaN(alphaValue)) {
                        fillOpacity = alphaValue / 100000;
                    }
                }
            }

            // Handle Line Stroke (border) - NEW SECTION
            const lineNode = shapeFill?.["a:ln"]?.[0];
            if (lineNode?.["a:solidFill"]?.[0]) {
                const strokeFill = lineNode["a:solidFill"][0];

                if (strokeFill?.["a:srgbClr"]?.[0]?.["$"]?.val) {
                    strokeColor = `#${strokeFill["a:srgbClr"][0]["$"].val}`;
                    // originalThemeColor = strokeColor; // It fill the border to shape background.

                    // Handle stroke luminance modifiers
                    const lumMod = strokeFill["a:srgbClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                    if (lumMod) {
                        originalLumMod = lumMod;
                        const lumOff = strokeFill["a:srgbClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                        if (lumOff) {
                            originalLumOff = lumOff;
                            strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                        } else {
                            strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                        }
                    }

                    // Extract stroke opacity (usually no alpha = fully opaque)
                    const strokeAlphaNode = strokeFill["a:srgbClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                    if (strokeAlphaNode) {
                        const alphaValue = parseInt(strokeAlphaNode, 10);
                        originalAlpha = alphaValue;

                        if (!isNaN(alphaValue)) {
                            strokeOpacity = alphaValue / 100000;
                        }
                    }

                } else if (strokeFill?.["a:schemeClr"]?.[0]?.["$"]?.val) {
                    let schmClr = strokeFill["a:schemeClr"][0]["$"].val;
                    // originalThemeColor = schmClr; // Use the same original theme color for stroke

                    // NEW: Check if we need to resolve through master color mapping for stroke
                    if (masterXML && schmClr) {
                        const resolvedColor = this.resolveMasterColor(schmClr, masterXML);
                        if (resolvedColor) {
                            schmClr = resolvedColor;
                        }
                    }

                    if (schmClr) {
                        strokeColor = colorHelper.resolveThemeColorHelper(schmClr, themeXML, masterXML);
                    }

                    // Handle stroke luminance modifiers for scheme color
                    const lumMod = strokeFill["a:schemeClr"][0]["a:lumMod"]?.[0]?.["$"]?.val;
                    if (lumMod) {
                        // originalLumMod = lumMod;
                        const lumOff = strokeFill["a:schemeClr"][0]["a:lumOff"]?.[0]?.["$"]?.val;
                        if (lumOff) {
                            // originalLumOff = lumOff;
                            strokeColor = pptBackgroundColors.applyLuminanceModifier(strokeColor, lumMod, lumOff);
                        } else {
                            strokeColor = colorHelper.applyLumMod(strokeColor, lumMod);
                        }
                    }

                    // Extract stroke opacity for scheme color
                    const strokeAlphaNode = strokeFill["a:schemeClr"][0]["a:alpha"]?.[0]?.["$"]?.val;
                    if (strokeAlphaNode) {
                        const alphaValue = parseInt(strokeAlphaNode, 10);
                        originalAlpha = alphaValue;

                        if (!isNaN(alphaValue)) {
                            strokeOpacity = alphaValue / 100000;
                        }
                    }
                }
            }

            // Handle Gradient Fill
            const gradFill = shapeFill?.["a:gradFill"]?.[0];
            if (gradFill?.["a:gsLst"]) {
                const gradientResult = this.getGradientFill(gradFill);
                if (gradientResult) {
                    fillColor = gradientResult.fillColor;
                    if (gradientResult.opacity !== undefined) {
                        fillOpacity = gradientResult.opacity;
                    }
                }
            }

            // Return both fill and stroke properties
            return {
                originalThemeColor, // NEW: return original theme color for reference
                originalLumMod,     // NEW: return original luminance modifier
                originalLumOff,     // NEW: return original luminance offset
                // txtPhType,
                originalAlpha,      // NEW: return original alpha value
                fillColor,
                opacity: fillOpacity,           // For backward compatibility
                fillOpacity: fillOpacity,       // Explicit fill opacity
                strokeColor: strokeColor,       // NEW: stroke color
                strokeOpacity: strokeOpacity    // NEW: stroke opacity
            };

        } catch (error) {
            console.error("Error processing shape fill color:", error);
            return {
                fillColor: "transparent",
                opacity: 1.0,
                fillOpacity: 1.0,
                strokeColor: "transparent",
                strokeOpacity: 1.0
            };
        }
    }

    // NEW: Helper function to resolve master color mapping
    resolveMasterColor(schemeColor, masterXML) {
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
    // Extract solid fill processing to separate method
    processSolidFill(solidFill, themeXML) {
        let fillColor = "transparent";
        let opacity = 1.0;

        // Process sRGB Color
        const srgbClr = solidFill?.["a:srgbClr"]?.[0];
        if (srgbClr?.["$"]?.val) {
            fillColor = `#${srgbClr["$"].val}`;
            opacity = this.extractOpacity(srgbClr) || opacity;
            fillColor = this.applyLuminanceModifiers(fillColor, srgbClr);
            return { fillColor, opacity };
        }

        // Process Scheme Color
        const schemeClr = solidFill?.["a:schemeClr"]?.[0];
        if (schemeClr?.["$"]?.val) {
            const schmClr = schemeClr["$"].val;
            fillColor = colorHelper.resolveThemeColorHelper(schmClr, themeXML, this.masterXML);
            opacity = this.extractOpacity(schemeClr) || opacity;
            fillColor = this.applyLuminanceModifiers(fillColor, schemeClr);
            return { fillColor, opacity };
        }

        return null;
    }

    // Extract opacity calculation
    extractOpacity(colorNode) {
        const alphaNode = colorNode?.["a:alpha"]?.[0]?.["$"]?.val;
        if (alphaNode) {
            const alphaValue = parseInt(alphaNode, 10);
            // Convert from PowerPoint's alpha format (0-100000) to CSS opacity (0-1)
            return isNaN(alphaValue) ? 1.0 : alphaValue / 100000;
        }
        return null;
    }

    // Extract luminance modifier logic
    applyLuminanceModifiers(fillColor, colorNode) {
        const lumMod = colorNode?.["a:lumMod"]?.[0]?.["$"]?.val;
        if (!lumMod) return fillColor;

        const lumOff = colorNode?.["a:lumOff"]?.[0]?.["$"]?.val;

        if (lumOff) {
            return pptBackgroundColors.applyLuminanceModifier(fillColor, lumMod, lumOff);
        } else {
            return colorHelper.applyLumMod(fillColor, lumMod);
        }
    }

    getLineStyle(shapeNode) {
        const lineRef = shapeNode?.["p:style"]?.[0]?.["a:lnRef"]?.[0];
        if (!lineRef) return "";
        const color = this.resolveColor(lineRef["a:schemeClr"][0]["$"]?.val);
        const shade = lineRef["a:schemeClr"][0]?.["a:shade"]?.[0]?.["$"]?.val || 100000;
        return `border-color: ${color}; filter: brightness(${shade / 100000});`;
    }

    getGradientFill(gradFill) {
        if (!gradFill || !gradFill["a:gsLst"] || !gradFill["a:gsLst"][0]["a:gs"]) {
            console.error("Invalid gradient fill structure:", gradFill);
            return { fillColor: 'linear-gradient(90deg, transparent, transparent)', opacity: 1 };
        }

        const gsList = gradFill["a:gsLst"][0]["a:gs"];
        const hasLinear = gradFill["a:lin"];
        const hasPath = gradFill["a:path"];
        const pathType = hasPath ? gradFill["a:path"][0]["$"]?.path || "circle" : null;
        let gradientType = hasLinear ? "linear" : "radial";

        if (pathType === "circle") gradientType = "radial";
        else if (pathType === "rect") gradientType = "rectangular";
        else if (pathType === "shape") gradientType = "path";

        // --- Extract & convert PowerPoint angle to CSS ---
        let pptxAngle = parseInt(gradFill["a:lin"]?.[0]?.["$"]?.ang || 0, 10);
        pptxAngle = isNaN(pptxAngle) ? 0 : pptxAngle / 60000;

        // ✅ Correct mapping: PPTX 0° (bottom→top) → CSS 90° (top→bottom)
        let cssAngle = (pptxAngle + 90) % 360;

        // --- Handle flip (x/y/xy) ---
        const flip = gradFill["$"]?.flip || "none";
        if (flip === "x") cssAngle = (180 - cssAngle + 360) % 360;
        else if (flip === "y") cssAngle = (360 - cssAngle) % 360;
        else if (flip === "xy") cssAngle = (cssAngle + 180) % 360;

        // --- Parse gradient stops into objects ---
        let totalAlpha = 0;
        let alphaCount = 0;

        let stopsObj = gsList.map(stop => {
            let hex = "#000000";
            let alpha = 1;

            if (stop["a:srgbClr"]) {
                hex = `#${stop["a:srgbClr"][0]["$"].val}`;
                if (stop["a:srgbClr"][0]["a:alpha"]) {
                    alpha = parseInt(stop["a:srgbClr"][0]["a:alpha"][0]["$"].val, 10) / 100000;
                }
            } else if (stop["a:schemeClr"]) {
                const schemeClr = stop["a:schemeClr"][0];
                let base = this.resolveColor(schemeClr["$"].val);
                const lumMod = parseInt(schemeClr["a:lumMod"]?.[0]?.["$"]?.val || 100000, 10) / 100000;
                const lumOff = parseInt(schemeClr["a:lumOff"]?.[0]?.["$"]?.val || 0, 10) / 100000;
                hex = this.adjustLuminance(base, lumMod, lumOff);

                if (schemeClr["a:alpha"]) {
                    alpha = parseInt(schemeClr["a:alpha"][0]["$"].val, 10) / 100000;
                }
            }

            const posVal = parseInt(stop["$"].pos, 10); // 0–100000
            const posPct = posVal / 1000;               // 0–100%

            totalAlpha += alpha;
            alphaCount++;

            return {
                posVal,
                posPct,
                rgbaStr: `rgba(${this.hexToRGB(hex)}, ${alpha}) ${posPct}%`
            };
        });

        const avgOpacity = alphaCount > 0 ? totalAlpha / alphaCount : 1;

        // --- Radial center (from fillToRect) ---
        let radialPosition = "center";
        if (hasPath && gradFill["a:path"][0]["a:fillToRect"]) {
            const rect = gradFill["a:path"][0]["a:fillToRect"][0]["$"] || {};
            const l = rect.l !== undefined ? parseInt(rect.l) : null;
            const t = rect.t !== undefined ? parseInt(rect.t) : null;
            const r = rect.r !== undefined ? parseInt(rect.r) : null;
            const b = rect.b !== undefined ? parseInt(rect.b) : null;

            if (l === 50000 && t === 50000 && r === 50000 && b === 50000) radialPosition = "center";
            else if (l === 100000 && t === 100000) radialPosition = "right bottom";
            else if (l === 100000 && b === 100000) radialPosition = "right top";
            else if (r === 100000 && t === 100000) radialPosition = "left bottom";
            else if (r === 100000 && b === 100000) radialPosition = "left top";
            else if (l === 100000) radialPosition = "right center";
            else if (r === 100000) radialPosition = "left center";
            else if (t === 100000) radialPosition = "center bottom";
            else if (b === 100000) radialPosition = "center top";
        }

        // --- Build final CSS stops string ---
        let gradientStops;
        if (gradientType === "radial") {
            stopsObj.sort((a, b) => a.posVal - b.posVal); // center → edge
            gradientStops = stopsObj.map(s => s.rgbaStr);
        } else if (gradientType === "linear") {
            stopsObj.sort((a, b) => a.posVal - b.posVal); // left → right
            gradientStops = stopsObj.map(s => s.rgbaStr);
        } else {
            gradientStops = stopsObj.map(s => s.rgbaStr);
        }

        // --- Compose final CSS gradient ---
        let fillColor = "";
        switch (gradientType) {
            case "linear":
                fillColor = `linear-gradient(${cssAngle}deg, ${gradientStops.join(", ")})`;
                break;
            case "radial":
                fillColor = `radial-gradient(circle at ${radialPosition}, ${gradientStops.join(", ")})`;
                break;
            case "rectangular":
                fillColor = `radial-gradient(ellipse at ${radialPosition}, ${gradientStops.join(", ")})`;
                break;
            case "path":
                fillColor = `radial-gradient(circle closest-side at ${radialPosition}, ${gradientStops.join(", ")})`;
                break;
            default:
                fillColor = `linear-gradient(${cssAngle}deg, ${gradientStops.join(", ")})`;
                break;
        }

        return { fillColor, opacity: avgOpacity };
    }

    adjustLuminance(hexColor, lumMod, lumOff) {
        const rgb = this.hexToRGB(hexColor).split(",").map(v => parseInt(v.trim(), 10));
        const adjust = val => Math.min(255, Math.round((val / 255) * lumMod * 255 + lumOff * 255));
        return `#${[adjust(rgb[0]), adjust(rgb[1]), adjust(rgb[2])].map(v => v.toString(16).padStart(2, "0")).join("")}`;
    }

    getShadowStyle(shapeNode) {
        const shadowNode = shapeNode?.["p:spPr"]?.[0]?.["a:effectLst"]?.[0]?.["a:outerShdw"]?.[0];

        if (!shadowNode) return "";

        const dist = parseInt(shadowNode["$"].dist, 10) / this.getEMUDivisor(); // Convert EMU to pixels, distance will be used as border width
        const dir = parseInt(shadowNode["$"].dir, 10) / 60000; // Convert 60000ths of a degree to degrees

        // Extract color and alpha
        const colorNode = shadowNode["a:prstClr"]?.[0];
        let color = "#000000"; // Default black
        let alpha = 1; // Default fully opaque

        if (colorNode) {
            if (colorNode["a:srgbClr"]) {
                color = `#${colorNode["a:srgbClr"][0]["$"].val}`;
            } else if (colorNode["a:schemeClr"]) {
                // Handle scheme color mapping
                color = this.resolveColor(colorNode["a:schemeClr"][0]["$"].val);
            }

            const alphaVal = colorNode["a:alpha"]?.[0]?.["$"]?.val;
            if (alphaVal) {
                alpha = parseInt(alphaVal, 10) / 100000; // Convert to CSS alpha range
            }
        }

        // Convert color to RGB for border color
        const rgbColor = this.hexToRGB(color);

        // Determine border placement based on direction
        const borderPlacement = this.calculateBorderPlacement(dir);
        // return `border: ${borderPlacement}px solid rgba(${rgbColor}, ${alpha});`;

        return `border: 2px solid rgba(${rgbColor}, ${alpha});`;
    }

    calculateBorderPlacement(directionDegrees) {
        if (directionDegrees >= 45 && directionDegrees < 135) {
            return 'border-top';
        } else if (directionDegrees >= 135 && directionDegrees < 225) {
            return 'border-right';
        } else if (directionDegrees >= 225 && directionDegrees < 315) {
            return 'border-bottom';
        } else {
            return 'border-left';
        }
    }


    hexToRGB(hex) {
        let r = parseInt(hex.slice(1, 3), 16);
        let g = parseInt(hex.slice(3, 5), 16);
        let b = parseInt(hex.slice(5, 7), 16);
        return `${r}, ${g}, ${b}`;
    }

    resolveColor(colorKey) {
        const mappedKey = this.clrMap[colorKey] || colorKey;

        const colorNode = this.themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0]?.[`a:${mappedKey}`]?.[0];

        if (colorNode?.["a:sysClr"]) {
            return `#${colorNode["a:sysClr"][0]["$"].lastClr}`;
        }

        return colorNode?.["a:srgbClr"] ? `#${colorNode["a:srgbClr"][0]["$"].val}` : "#000000";
    }

    getShapeBorderStyle(outline) {
        // let line = shapeNode?.["p:spPr"]?.[0]?.["a:ln"]?.[0];

        if (!outline) return '';

        let width = outline["$"]?.w ? parseInt(outline["$"].w, 10) / this.getEMUDivisor() : 0;
        let color = "transparent";
        const solidFill = outline?.["a:solidFill"]?.[0];

        if (solidFill?.["a:srgbClr"]) {
            color = `#${solidFill["a:srgbClr"][0]["$"].val}`;
        } else if (solidFill?.["a:schemeClr"]) {
            color = this.resolveColor(solidFill["a:schemeClr"][0]["$"].val);
        }
        if (color === "transparent") {
            width = 0
        }
        return `border: ${width}px solid ${color};`;
    }
    // ✅ ADD METHOD 1
    extractTableCellTextColor(tcNode, themeXML, clrMap) {
        try {
            const textBody = tcNode?.["a:txBody"]?.[0];
            if (!textBody) return null;

            const paragraphs = textBody["a:p"] || [];
            
            for (const p of paragraphs) {
                const pPr = p["a:pPr"]?.[0];
                if (pPr) {
                    const defRPr = pPr["a:defRPr"]?.[0];
                    if (defRPr) {
                        const color = this.extractColorFromRunProperties(defRPr, themeXML, clrMap);
                        if (color) return color;
                    }
                }

                const runs = p["a:r"] || [];
                for (const run of runs) {
                    const rPr = run["a:rPr"]?.[0];
                    if (rPr) {
                        const color = this.extractColorFromRunProperties(rPr, themeXML, clrMap);
                        if (color) return color;
                    }
                }
            }

            return null;
        } catch (error) {
            console.error("Error extracting table cell text color:", error);
            return null;
        }
    }

    // ✅ ADD METHOD 2
    extractColorFromRunProperties(rPr, themeXML, clrMap) {
        try {
            const solidFill = rPr["a:solidFill"]?.[0];
            if (!solidFill) return null;

            const srgbClr = solidFill["a:srgbClr"]?.[0];
            if (srgbClr) {
                const val = srgbClr["$"]?.val;
                if (val) return `#${val}`;
            }

            const schemeClr = solidFill["a:schemeClr"]?.[0];
            if (schemeClr) {
                const val = schemeClr["$"]?.val;
                if (val) {
                    const resolved = this.resolveSchemeColor(val, themeXML, clrMap);
                    if (resolved) return resolved;
                }
            }

            const sysClr = solidFill["a:sysClr"]?.[0];
            if (sysClr) {
                const lastClr = sysClr["$"]?.lastClr;
                if (lastClr) return `#${lastClr}`;
            }

            return null;
        } catch (error) {
            return null;
        }
    }

    // ✅ ADD METHOD 3
    resolveSchemeColor(schemeClr, themeXML, clrMap) {
        try {
            if (!themeXML) return null;

            let mappedColor = schemeClr;
            if (clrMap && clrMap[schemeClr]) {
                mappedColor = clrMap[schemeClr];
            }

            const clrScheme = themeXML?.["a:theme"]?.["a:themeElements"]?.[0]?.["a:clrScheme"]?.[0];
            if (!clrScheme) return null;

            const colorNode = clrScheme[`a:${mappedColor}`]?.[0];
            if (!colorNode) return null;

            const srgbClr = colorNode["a:srgbClr"]?.[0]?.["$"]?.val;
            if (srgbClr) return `#${srgbClr}`;

            const sysClr = colorNode["a:sysClr"]?.[0]?.["$"]?.lastClr;
            if (sysClr) return `#${sysClr}`;

            return null;
        } catch (error) {
            return null;
        }
    }

    // ✅ ADD METHOD 4
    extractTableCellText(tcNode) {
        try {
            const textBody = tcNode?.["a:txBody"]?.[0];
            if (!textBody) return '';

            const paragraphs = textBody["a:p"] || [];
            const textParts = [];

            for (const p of paragraphs) {
                const runs = p["a:r"] || [];
                const paragraphText = runs.map(run => {
                    const text = run["a:t"]?.[0];
                    return typeof text === 'string' ? text : '';
                }).join('');

                if (paragraphText) {
                    textParts.push(paragraphText);
                }
            }

            return textParts.join('<br>');
        } catch (error) {
            return '';
        }
    }

}

module.exports = ShapeHandler;
