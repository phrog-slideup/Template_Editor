const fs = require("fs");
const path = require("path");
const config = require('../config.js');

const ShapeHandler = require("../pptx-To-Html-styling/shapeHandler.js");
const colorHelper = require("../api/helper/colorHelper.js");
const pptBackgroundColors = require("../pptx-To-Html-styling/pptBackgroundColors.js");
const pptTextAllInfo = require("../pptx-To-Html-styling/pptTextAllInfo.js");
const imgSvgStyle = require("../pptx-To-Html-styling/imgSvgCss.js");
const { getImageFromPicture } = require("../api/helper/pptImagesHandling.js");


// Define the directory to save images
const imageSavePath = path.resolve(__dirname, "../uploads/");

if (!fs.existsSync(imageSavePath)) {
    fs.mkdirSync(imageSavePath, { recursive: true });
}

class slideLayoutForm {
    constructor(themeXML, extractor, slidePath, slideRelXML, layoutPath, layoutRelationship, masterXML = null) {
        this.themeXML = themeXML;
        this.slidePath = slidePath;
        this.extractor = extractor;
        this.slideRelXML = slideRelXML;
        this.layoutPath = layoutPath;
        this.layoutRelationship = layoutRelationship;
        this.masterXML = masterXML; // Store master XML for color resolution
        this.nodes = []; // Initialize nodes for z-index tracking
    }

    // NEW: Helper function to add "master" class to all HTML elements
    addMasterClassToHTML(htmlContent) {
        if (!htmlContent || htmlContent.trim() === '') {
            return htmlContent;
        }

        // Replace class attributes to add "master" class
        return htmlContent.replace(/class="([^"]*)"/g, (match, classes) => {
            // Check if "master" class is already present
            if (classes.includes('master')) {
                return match; // Already has master class, return as is
            }
            // Add "master" class to existing classes
            return `class="${classes} master"`;
        });
    }

    async createSlideLayoutForm(slideLayoutXML, themeXML, extractor, slidePath, slideRelXML, layoutRelationship, masterXML = null) {

        if (!slideLayoutXML) return "";

        // Extract color mapping similar to main slide processing
        const clrMap = await pptBackgroundColors.getColorMappingFromMaster(slideLayoutXML);

        // Get nodes in correct order for z-index assignment
        this.nodes = await this.getNodesInCorrectOrder(slideLayoutXML, 0);

        // Extract all shape types from slide layout
        const shapeNodes = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
        const lineShapeNodes = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:cxnSp"] || [];
        const tableNodes = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:graphicFrame"] || [];

        let htmlContent = "";

        // Create a ShapeHandler instance similar to main slide processing
        const shapeHandler = new ShapeHandler(themeXML, clrMap, this.nodes, extractor, this.layoutPath, this.layoutRelationship, 0);

        // Use the same method as main slide to convert all shapes to HTML
        htmlContent += await shapeHandler.convertAllShapesToHTML(shapeNodes, lineShapeNodes, tableNodes, themeXML, masterXML);

        // Process any pictures/images that might be in the layout
        const picNodesImg = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:pic"] || [];

        for (const picNode of picNodesImg) {
            const position = pptTextAllInfo.getPositionFromShape(picNode, 0); // flag = 0 for layout
            const blipNode = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0];
            const svgContent = await this.getSVGContent(blipNode, this.layoutRelationship);
            const imageInfo = await this.getImageFromPicture(picNode, this.layoutPath, this.layoutRelationship);
            const nodeName = picNode["p:nvPicPr"][0]["p:cNvPr"][0].$.name;
            const matchingNode = this.nodes.find(node => node.name === nodeName);

            // If a matching node is found, use its id as the z-index
            const zIndex = matchingNode ? matchingNode.id : 0;

            const rotation = pptTextAllInfo.getRotation(picNode);
            const imgcss = imgSvgStyle.returnImgSvgStyle(picNode);

            if (!imageInfo && svgContent) {
                htmlContent += `<div class="sli-svg-container master" 
                                style="position:absolute; 
                                left:${position.x}px; 
                                top:${position.y}px; 
                                width:${position.width}px; 
                                height:${position.height}px;
                                transform : ${imgcss.flipH} ${imgcss.flipV} rotate(${rotation}deg);
                                overflow:hidden;
                                z-index:${zIndex};">
                                <svg style="width:100%; height:100%;" xmlns="http://www.w3.org/2000/svg">${svgContent}</svg>
                            </div>`;
            } else if (imageInfo) {
                const shapeStyle = imageInfo.shape === "ellipse" ? "border-radius: 50%; overflow: hidden;" : "";
                const borderCSS = imageInfo.border.width > 0 ? `border: ${imageInfo.border.width}px solid ${imageInfo.border.color};` : "";
                const boxShadowCSS = (imgcss.shadowOffsetX || imgcss.shadowOffsetY) ? `box-shadow: ${imgcss.shadowOffsetX}px ${imgcss.shadowOffsetY}px ${imgcss.shadowColor};` : "";
                const combinedFilter = imgcss.blurAmount ? `filter: blur(${imgcss.blurAmount}px) contrast(${imgcss.contrastValue || 1});` : "filter:none;";

                htmlContent += `
                <div class="image-container master"
                    style="position:absolute;
                    left:${position.x}px;
                    top:${position.y}px;
                    width:${position.width}px;
                    height:${position.height}px;
                    ${shapeStyle}
                    opacity:${imgcss.opacity};
                    transform: ${imgcss.transform} rotate(${rotation}deg);
                    overflow:hidden;
                    z-index:${zIndex};">
                    
                    <img src="${imageInfo.src}" alt="Img"
                        style="position:absolute;
                        width:${position.width}px;
                        height:${position.height}px;
                        opacity:${imgcss.opacity};
                        z-index:${zIndex};
                        ${borderCSS}
                        ${boxShadowCSS}
                        ${combinedFilter}" />
                </div>`;
            }
        }

        // NEW: Add "master" class to all HTML elements
        if (htmlContent.trim()) {
            const htmlWithMasterClasses = this.addMasterClassToHTML(htmlContent);

            return `<div class="sli-master" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none; z-index: -1;">
                        ${htmlWithMasterClasses}
                    </div>`;
        }

        return ""; // Return empty string if no content
    }

    // Get nodes in correct order for z-index assignment (similar to main slide processing)
    async getNodesInCorrectOrder(xmlFile, highestZIndex) {
        try {
            // Get the shape tree from the slide layout XML
            const spTree = xmlFile['p:sldLayout']['p:cSld'][0]['p:spTree'][0];
            const orderedNodes = [];
            let sequentialId = highestZIndex + 1;

            // Helper: traverse a list of nodes in order
            function traverseNodes(nodeList) {
                nodeList.forEach(node => {
                    const nodeType = node['#name'];
                    // Only process actual shape elements
                    if (nodeType === 'p:sp' || nodeType === 'p:cxnSp' ||
                        nodeType === 'p:pic' || nodeType === 'p:grpSp') {
                        // Create node info with sequential z-index ID
                        const nodeInfo = { type: nodeType, id: sequentialId++ };

                        // Extract shape name if available (from p:cNvPr within the non-visual props)
                        if (node.$$) {
                            const nvPrNode = node.$$.find(child =>
                                child['#name'].includes('nvGrpSpPr') ||
                                child['#name'].includes('nvSpPr') ||
                                child['#name'].includes('nvPicPr') ||
                                child['#name'].includes('nvCxnSpPr')
                            );
                            if (nvPrNode && nvPrNode.$$) {
                                const cNvPr = nvPrNode.$$.find(n => n['#name'] === 'p:cNvPr');
                                if (cNvPr && cNvPr.$ && cNvPr.$.name) {
                                    nodeInfo.name = cNvPr.$.name;
                                }
                            }
                        }

                        orderedNodes.push(nodeInfo);

                        // If this is a group shape, recursively process its children
                        if (nodeType === 'p:grpSp' && node.$$) {
                            traverseNodes(node.$$);  // Traverse all child nodes inside the group
                        }
                    }
                });
            }

            // Start traversal from the root shape tree's children (in order)
            if (spTree.$$) {
                traverseNodes(spTree.$$);
            }

            return orderedNodes;
        } catch (error) {
            console.error('Error in getNodesInCorrectOrder:', error);
            return [];
        }
    }

    // Get SVG content from <a:blip> node (similar to main slide processing)
    async getSVGContent(blipNode, relationshipsXML) {
        const svgEmbedId = blipNode?.["a:extLst"]?.[0]?.["a:ext"]?.[0]?.["asvg:svgBlip"]?.[0]?.["$"]?.["r:embed"];
        if (!svgEmbedId) {
            return null;
        }
        const relationship = relationshipsXML?.["Relationships"]?.["Relationship"]?.find((rel) => rel["$"]?.Id === svgEmbedId);

        if (!relationship) {
            return null;
        }
        const svgPath = path.posix.normalize(`ppt/media/${relationship["$"]?.Target}`);

        const svgFile = this.extractor.files[svgPath];
        if (!svgFile) {
            return null;
        }

        return await svgFile.async("string");
    }

    // Get image from picture (wrapper for the imported function)
    async getImageFromPicture(picNode, layoutPath, relationshipsXML) {
        try {
            return await getImageFromPicture(
                picNode,
                layoutPath,
                this.extractor.files,
                relationshipsXML,
                this.themeXML
            );
        } catch (error) {
            console.error('Error getting image from picture:', error);
            return null;
        }
    }

    // Legacy methods for backward compatibility (if still needed elsewhere)
    async extractShapes(slideLayoutXML, themeXML, slidePath, slideRelXML) {
        const shapeNodes = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
        const connectorNodes = slideLayoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:cxnSp"] || [];

        // Process both shapes and connectors
        const shapePromises = shapeNodes.map(shape => this.processShape(shape, themeXML, slidePath, slideRelXML));
        const connectorPromises = connectorNodes.map(connector => this.processConnector(connector, themeXML, slidePath, slideRelXML));

        const shapes = await Promise.all(shapePromises);
        const connectors = await Promise.all(connectorPromises);

        return [...shapes, ...connectors];
    }

    async processShape(shape, themeXML, slidePath, slideRelXML) {
        // Log the entire shape object for debugging purposes
        const zIndex = await this.getZIndexForShape(shape);

        const shapeHandler = new ShapeHandler();
        let htmlContent = "";

        const shapeId = shape?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.$.id || "Unknown";
        const shapeName = shape?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.$.name || "Shape";

        const blip = shape?.["p:spPr"]?.[0]?.["a:blipFill"]?.[0]?.["a:blip"]?.[0];

        if (blip && blip["$"] && blip["$"]["r:embed"] || blip !== undefined) {

            const imageId = blip?.["$"]?.["r:embed"];

            // Use the imported getImageFromPicture function
            const imageInfo = await getImageFromPicture(
                shape,
                this.layoutPath,
                this.extractor.files,
                this.layoutRelationship,
                this.themeXML
            );

            if (imageInfo) {
                const rotation = pptTextAllInfo.getRotation(shape);
                const imgcss = imgSvgStyle.returnImgSvgStyle(shape);
                const position = pptTextAllInfo.getPositionFromShape(shape);

                // Apply border-radius for ellipse detection
                const shapeStyle = imageInfo.shape === "ellipse" ? "border-radius: 50%; overflow: hidden;" : "";
                const borderCSS = imageInfo.border.width > 0 ? `border: ${imageInfo.border.width}px solid ${imageInfo.border.color};` : "";
                const boxShadowCSS = (imgcss.shadowOffsetX || imgcss.shadowOffsetY) ? `box-shadow: ${imgcss.shadowOffsetX}px ${imgcss.shadowOffsetY}px ${imgcss.shadowColor};` : "";
                const combinedFilter = imgcss.blurAmount ? `filter: blur(${imgcss.blurAmount}px) contrast(${imgcss.contrastValue || 1});` : "filter:none;";

                htmlContent += `<div class="image-container master"
                                    style="position:absolute;
                                    left:${position.x}px;
                                    top:${position.y}px;
                                    width:${position.width}px;
                                    height:${position.height}px;
                                    ${shapeStyle}
                                    ${borderCSS}
                                    ${boxShadowCSS}
                                    opacity:${imgcss.opacity};
                                    transform: ${imgcss.transform} 
                                    rotate(${rotation}deg);
                                    overflow:hidden;
                                    z-index:${zIndex};">
                                    
                                    <img src="${imageInfo.src}" alt="Img"
                                        style="position:absolute;
                                        width:100%;
                                        height:100%;
                                        transform: ${imgcss.flipH} ${imgcss.flipV} rotate(${rotation}deg);
                                        opacity:${imgcss.opacity};
                                        z-index:${zIndex};
                                        ${borderCSS}
                                        ${boxShadowCSS}
                                        ${combinedFilter}" />
                            </div>`;
            }
        }

        // Initialize color to default white
        let color = "#FFFFFF"; // Default color if not found

        // Handle color extraction from solid fill
        const solidFill = shape?.["p:spPr"]?.[0]?.["a:solidFill"]?.[0];

        if (solidFill) {
            if (solidFill["a:srgbClr"]) {
                // RGB Hex color
                color = `#${solidFill["a:srgbClr"][0]?.["$"]?.["val"]}`;
            } else if (solidFill["a:schemeClr"]) {
                // Scheme color (e.g., bg1, accent1, etc.)
                const scheme = solidFill["a:schemeClr"][0]?.["$"]?.["val"];

                const resolveColor = colorHelper.resolveThemeColorHelper(scheme, themeXML, this.masterXML);

                let lumMod = solidFill?.["a:schemeClr"]?.[0]?.["a:lumMod"]?.[0]?.["$"]?.val;

                if (lumMod) {
                    // Apply luminance modification if present
                    color = colorHelper.applyLumMod(resolveColor, lumMod);
                } else {
                    color = resolveColor;
                }
            }
        }

        // Handle gradient fill if present
        const gradFill = shape?.["p:spPr"]?.[0]?.["a:gradFill"]?.[0];

        if (gradFill) {
            color = this.extractGradientFill(gradFill); // Process gradient if available
        }

        // Extract position and size
        const position = this.extractPosition(shape); // Safely extract position
        const width = position?.width || 0; // Fallback to 0 if width is undefined
        const height = position?.height || 0; // Fallback to 0 if height is undefined

        // Extract custom geometry if present and call convertShapeToHTML
        const customGeometry = this.extractCustomGeometry(shape);

        if (customGeometry) {

            // Custom geometry detected, use convertShapeToHTML to convert it to HTML
            let shapeHtml = await shapeHandler.convertShapeToHTML(shape, themeXML, slidePath, this.masterXML);

            // Process the HTML to remove text content
            shapeHtml = await this.removeTextBoxContent(shapeHtml);

            // Add master class to the generated HTML
            shapeHtml = this.addMasterClassToHTML(shapeHtml);

            // Now add the cleaned HTML to htmlContent
            htmlContent += shapeHtml;
        }

        // Return processed shape details
        return {
            id: shapeId,
            name: shapeName,
            fillColor: color,
            position: position,
            width: width,
            height: height,
            customGeometry: customGeometry,
            htmlContent: htmlContent // Return the HTML content for the shape
        };
    }

    // New method to process connectors
    async processConnector(connector, themeXML, slidePath, slideRelXML) {
        const zIndex = await this.getZIndexForConnector(connector);

        const shapeHandler = new ShapeHandler();
        let htmlContent = "";

        const connectorId = connector?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.$.id || "Unknown";
        const connectorName = connector?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.$.name || "Connector";

        // Process connector using ShapeHandler
        htmlContent += await shapeHandler.convertConnectorToHTML(connector, themeXML, slidePath);

        // Add master class to connector HTML
        htmlContent = this.addMasterClassToHTML(htmlContent);

        // Extract position and size
        const position = this.extractConnectorPosition(connector);

        return {
            id: connectorId,
            name: connectorName,
            type: "connector",
            position: position,
            htmlContent: htmlContent
        };
    }

    // Function to remove text content from HTML
    removeTextBoxContent(html) {
        // Create a temporary DOM element to parse the HTML
        const tempElement = new (require('jsdom').JSDOM)(`<!DOCTYPE html><div id="temp">${html}</div>`).window.document.getElementById('temp');

        // Find all elements with class "sli-txt-box"
        const textBoxes = tempElement.querySelectorAll('.sli-txt-box');

        // For each text box, empty its content but keep the element itself
        textBoxes.forEach(textBox => {
            textBox.innerHTML = '';
        });

        // Return the modified HTML
        return tempElement.innerHTML;
    }

    async getZIndexForShape(shapeNode) {
        if (!this.nodes || this.nodes.length === 0) return 0;

        // Extract name from shapeNode
        let shapeName = null;
        const nvPrNode = shapeNode?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0];

        if (nvPrNode && nvPrNode.$ && nvPrNode.$.name) {
            shapeName = nvPrNode.$.name;
        }

        if (!shapeName) return 0; // If no name, return default z-index

        // Find matching node from ordered list
        const matchedNode = this.nodes.find(node => node.name === shapeName);
        return matchedNode ? matchedNode.id : 0;
    }

    async getZIndexForConnector(connectorNode) {
        if (!this.nodes || this.nodes.length === 0) return 0;

        // Extract name from connectorNode
        let connectorName = null;
        const nvPrNode = connectorNode?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0];

        if (nvPrNode && nvPrNode.$ && nvPrNode.$.name) {
            connectorName = nvPrNode.$.name;
        }

        if (!connectorName) return 0; // If no name, return default z-index

        // Find matching node from ordered list
        const matchedNode = this.nodes.find(node => node.name === connectorName);
        return matchedNode ? matchedNode.id : 0;
    }

    extractPosition(shape) {
        const xfrm = shape?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0] || {};
        const x = xfrm["a:off"]?.[0]?.$.x || 0;
        const y = xfrm["a:off"]?.[0]?.$.y || 0;
        const width = xfrm?.["a:ext"]?.[0]?.$.cx || 100;
        const height = xfrm?.["a:ext"]?.[0]?.$.cy || 100;

        const divisor = 12700;

        return {
            x: x / divisor,
            y: y / divisor,
            width: width / divisor,
            height: height / divisor
        };
    }

    extractConnectorPosition(connector) {
        const xfrm = connector?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0] || {};
        const x = xfrm["a:off"]?.[0]?.$.x || 0;
        const y = xfrm["a:off"]?.[0]?.$.y || 0;
        const width = xfrm?.["a:ext"]?.[0]?.$.cx || 100;
        const height = xfrm?.["a:ext"]?.[0]?.$.cy || 100;

        const divisor = 12700;

        return {
            x: x / divisor,
            y: y / divisor,
            width: width / divisor,
            height: height / divisor
        };
    }

    // Extract custom geometry (for shapes like freeforms)
    extractCustomGeometry(shape) {
        const custGeom = shape?.["p:spPr"]?.[0]?.["a:custGeom"]?.[0];
        const prstGeom = shape?.["p:spPr"]?.[0]?.["a:prstGeom"]?.[0];

        if (prstGeom) {
            return {
                gdData: custGeom?.["a:avLst"]?.[0]?.["a:gd"] || [],
            }
        }
        if (custGeom) {
            return {
                pathData: custGeom?.["a:pathLst"]?.[0]?.["a:path"] || [],
                gdData: custGeom?.["a:gdLst"]?.[0]?.["a:gd"] || []
            };
        }
        return {};
    }

    // Extract gradient fill data
    extractGradientFill(gradFill) {
        const gsList = gradFill?.["a:gsLst"]?.[0]?.["a:gs"] || [];
        const gradientColors = gsList.map(gs => `#${gs?.["a:srgbClr"]?.[0]?.$.val}`).join(", ");

        return `linear-gradient(${gradientColors})`;
    }
}

module.exports = {
    slideLayoutForm
};