const fs = require("fs");
const path = require("path");
const colorHelper = require("../api/helper/colorHelper.js");
const pptTextAllInfo = require("../pptx-To-Html-styling/pptTextAllInfo.js");
const ShapeHandler = require("../pptx-To-Html-styling/shapeHandler.js");

class masterHtml {

    constructor(themeXml, masterXml, masterRelsXml, slideSize) {
        this.themeXml = themeXml;
        this.masterXml = masterXml;
        this.masterRelsXml = masterRelsXml;
        this.slideSize = slideSize;
    }

    async generateMasterHTML(shapeTree, lineShapes, extractor) {
        try {
            const masterHtmlParts = [];

            // Use the passed shapeTree or extract from masterXml
            const masterShapes = shapeTree || this.masterXml["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
            const masterLineShapes = lineShapes || this.masterXml["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:cxnSp"] || [];

            // Check if we have any shapes to process
            if (masterShapes.length === 0 && masterLineShapes.length === 0) {
                return `<div class="sli-master"></div>`;
            }

            // Get color mapping for master slide
            const clrMap = this.getColorMappingFromMaster();

            // Create nodes array for z-index ordering (simplified for master)
            const nodes = this.createNodesArray(masterShapes, masterLineShapes);

            // Create ShapeHandler with correct parameters
            const shapeHandler = new ShapeHandler(
                this.themeXml,
                clrMap,
                nodes,
                extractor,
                this.masterRelsXml,
            );

            // Process regular shapes
            if (masterShapes.length > 0) {
                for (const shape of masterShapes) {
                    try {

                        const phElement = shape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0];

                        // Check if it's a placeholder (has p:ph element)
                        const isPlaceholder = Boolean(phElement);

                        // // Check placeholder attributes safely
                        // let hasType = Boolean(phElement?.["$"]?.type);
                        // let hasIdx = Boolean(phElement?.["$"]?.idx);

                        let shapeHtml = await shapeHandler.convertShapeToHTML(shape, this.themeXml, this.masterXml);

                        if (isPlaceholder) {
                            // IMPROVED: Better regex to handle existing styles
                            if (shapeHtml.includes('style="')) {
                                // Add to existing style
                                shapeHtml = shapeHtml.replace(
                                    /(<div[^>]*style="[^"]*)/g,
                                    '$1; visibility: hidden'
                                );
                            } else {
                                // Add new style attribute
                                shapeHtml = shapeHtml.replace(
                                    /<div([^>]*)>/,
                                    '<div$1 style="visibility: hidden;">'
                                );
                            }

                        }

                        if (shapeHtml && shapeHtml.trim() !== '') {
                            // Add master class to the shape HTML
                            const masterShapeHtml = this.addMasterClassToShapes(shapeHtml);
                            masterHtmlParts.push(masterShapeHtml);
                        }
                    } catch (error) {
                        console.error("Error processing master shape:", error);
                    }
                }
            }

            // Process connector shapes (lines)
            if (masterLineShapes.length > 0) {
                for (const lineShape of masterLineShapes) {
                    try {
                        const lineHtml = await shapeHandler.convertConnectorToHTML(lineShape, this.themeXml);
                        if (lineHtml && lineHtml.trim() !== '') {
                            // Add master class to the line shape HTML
                            const masterLineHtml = this.addMasterClassToShapes(lineHtml);
                            masterHtmlParts.push(masterLineHtml);
                        }
                    } catch (error) {
                        console.error("Error processing master line shape:", error);
                    }
                }
            }

            // Return the master content with proper styling
            return `<div class="sli-master main-master" 
                            style="position: absolute; 
                            top: 0; left: 0; 
                            width: ${this.slideSize.width}px; 
                            height: ${this.slideSize.height}px; 
                            z-index: 0;">
                                ${masterHtmlParts.join("")}
                            </div>`;

        } catch (error) {
            console.error("Error in generateMasterHTML:", error);
            return `<div class="sli-master"></div>`;
        }
    }

    /**
     * Adds "master" class to shape divs to identify them as master elements
     * @param {string} htmlString - The HTML string containing shape divs
     * @returns {string} - Modified HTML string with "master" class added
     */
    addMasterClassToShapes(htmlString) {
        try {
            // Regular expression to find and replace "custom-shape" with "custom-shape-master"
            // This will specifically target the "custom-shape" class and replace it
            const customShapeRegex = /class="([^"]*)\bcustom-shape\b([^"]*)"/g;

            // Replace "custom-shape" with "custom-shape-master"
            const modifiedHtml = htmlString.replace(customShapeRegex, (match, beforeCustomShape, afterCustomShape) => {
                // Check if "custom-shape-master" is not already present
                if (!match.includes('custom-shape-master')) {
                    return `class="${beforeCustomShape}custom-shape-master${afterCustomShape}"`;
                }
                return match; // Return unchanged if "custom-shape-master" already exists
            });

            return modifiedHtml;
        } catch (error) {
            console.error("Error modifying custom-shape class:", error);
            return htmlString; // Return original HTML if error occurs
        }
    }

    // Helper method to get color mapping from master
    getColorMappingFromMaster() {
        try {
            const clrMap = this.masterXml?.["p:sldMaster"]?.["p:clrMap"]?.[0]?.["$"];
            return clrMap || {
                bg1: "lt1",
                tx1: "dk1",
                bg2: "lt2",
                tx2: "dk2",
                accent1: "accent1",
                accent2: "accent2",
                accent3: "accent3",
                accent4: "accent4",
                accent5: "accent5",
                accent6: "accent6",
                hlink: "hlink",
                folHlink: "folHlink"
            };
        } catch (error) {
            console.error("Error getting color mapping:", error);
            return {
                bg1: "lt1",
                tx1: "dk1",
                bg2: "lt2",
                tx2: "dk2",
                accent1: "accent1",
                accent2: "accent2",
                accent3: "accent3",
                accent4: "accent4",
                accent5: "accent5",
                accent6: "accent6",
                hlink: "hlink",
                folHlink: "folHlink"
            };
        }
    }

    // Helper method to create nodes array for z-index ordering
    createNodesArray(shapes, lineShapes) {
        const nodes = [];
        let sequentialId = 1;

        // Add regular shapes
        if (shapes && Array.isArray(shapes)) {
            shapes.forEach(shape => {
                const shapeName = shape?.["p:nvSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
                if (shapeName) {
                    nodes.push({
                        type: 'p:sp',
                        id: sequentialId++,
                        name: shapeName
                    });
                }
            });
        }

        // Add line shapes
        if (lineShapes && Array.isArray(lineShapes)) {
            lineShapes.forEach(lineShape => {
                const shapeName = lineShape?.["p:nvCxnSpPr"]?.[0]?.["p:cNvPr"]?.[0]?.["$"]?.name;
                if (shapeName) {
                    nodes.push({
                        type: 'p:cxnSp',
                        id: sequentialId++,
                        name: shapeName
                    });
                }
            });
        }

        return nodes;
    }

    // Helper method to check if master slide has background
    hasMasterBackground() {
        const bgProps = this.masterXml?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:bg"];
        return bgProps && bgProps.length > 0;
    }

    // Helper method to get master background if exists
    getMasterBackground() {
        try {
            const bgProps = this.masterXml?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:bg"]?.[0];
            if (bgProps) {
                return this.processMasterBackground(bgProps);
            }
            return null;
        } catch (error) {
            console.error("Error getting master background:", error);
            return null;
        }
    }

    // Helper method to process master background
    processMasterBackground(bgProps) {
        return null;
    }
}

module.exports = masterHtml;