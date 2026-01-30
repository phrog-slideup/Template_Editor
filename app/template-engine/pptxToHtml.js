const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const xml2js = require("xml2js");
const path = require("path");
const { posix } = require("path");
const fs = require("fs").promises;

const pptBackgroundColors = require("../pptx-To-Html-styling/pptBackgroundColors.js");
const pptTextAllInfo = require("../pptx-To-Html-styling/pptTextAllInfo.js");
const imgSvgStyle = require("../pptx-To-Html-styling/imgSvgCss.js");
const ShapeHandler = require("../pptx-To-Html-styling/shapeHandler.js");

const masterBackgroundHandler = require("../pptx-To-Html-styling/MasterBackgroundHandler.js");
const masterHtml = require("../pptx-To-Html-styling/masterHtml.js");
const { overrideGetImageFromPicture } = require("../api/helper/pptImagesHandling.js");
const { slideLayoutForm: SlideLayoutForm } = require('../pptx-To-Html-styling/slideLayoutForm.js');
const pptLayoutBackgroundColors = require("../pptx-To-Html-styling/layoutBackgroundHandler.js");

const PptxUngrouper = require("../api/helper/pptxUngrouper.js");

require("dotenv").config();

const sharedCache = require('../api/shared/cache.js');

/** Normalize ANY path to forward-slash + collapse ../ and remove leading slashes */
function normalizePath(p) {
  return posix.normalize(String(p).replace(/\\/g, '/')).replace(/^\/+/, '');
}

class pptxToHtml {
  constructor(unzippedFiles, extractor) {
    this.files = {};
    for (const [k, v] of Object.entries(unzippedFiles)) {
      this.files[normalizePath(k)] = v;
    }
    this.themeColors = null;
    this.cache = {};
    this.extractor = extractor;
  }

  // Clear the cache
  clearCache() {
    this.cache = {};
  }

  async parseXML(filePath, returnRaw = false) {

    if (!returnRaw && this.cache[filePath]) {
      return this.cache[filePath];
    }

    if (this.cache[filePath]) {
      return this.cache[filePath];
    }

    const file = this.files[filePath];
    if (!file) {
      console.error(`File not found: ${filePath}`);
      return null;
    }

    try {
      const xmlContent = await file.async("string");

      // Only ungroup slide XML files, not .rels files or other XML files
      if (typeof filePath === 'string' &&
        (filePath.match(/ppt\/slides\/slide\d+\.xml$/) ||
          filePath.match(/ppt\/slideMasters\/slideMaster\d+\.xml$/)) &&
        !filePath.includes('_rels')) {

        // Try to ungroup elements, but handle failures gracefully
        let processedXmlContent;
        try {
          processedXmlContent = await PptxUngrouper.ungroupElements(xmlContent);
        } catch (ungroupError) {
          console.warn(`Ungrouping failed for ${filePath}, using original content:`, ungroupError.message);
          processedXmlContent = xmlContent; // Use original content if ungrouping fails
        }

        // ⭐ ADD THIS: Return raw XML if requested
        if (returnRaw) {
          return processedXmlContent;
        }

        // Parse the processed XML (either ungrouped or original)
        const parser = new xml2js.Parser({
          preserveChildrenOrder: true,
          explicitChildren: true,
          explicitArray: true
        });

        const result = await parser.parseStringPromise(processedXmlContent);
        this.cache[filePath] = result; // Cache the parsed XML
        return result;
      } else {

        // ⭐ ADD THIS: Return raw XML if requested
        if (returnRaw) {
          return xmlContent;
        }
        // For non-slide files (including .rels files), proceed as normal
        const parser = new xml2js.Parser({
          preserveChildrenOrder: true,
          explicitChildren: true,
          explicitArray: true
        });

        const result = await parser.parseStringPromise(xmlContent);
        this.cache[filePath] = result; // Cache the parsed XML
        return result;
      }
    } catch (error) {
      console.error(`Error parsing XML file at ${filePath}:`, error);
      throw new Error(`Failed to parse XML: ${filePath}`, error);
    }
  }

  /**Convert all slides in the presentation to HTML*/
  async convertAllSlidesToHTML(flag) {
    const slideHTMLs = [];

    // Step 1: Get rIds from presentation.xml
    const sldPathFromPresentation = "ppt/presentation.xml";
    const sldXML = await this.parseXML(sldPathFromPresentation);

    const sldIdListContainer = sldXML?.["p:presentation"]?.["p:sldIdLst"];
    const sldIds = Array.isArray(sldIdListContainer)
      ? sldIdListContainer[0]?.["p:sldId"]
      : sldIdListContainer?.["p:sldId"];

    const rIds = Array.isArray(sldIds)
      ? sldIds.map(slide => slide?.["$"]?.["r:id"]).filter(Boolean)
      : [];

    // Step 2: Parse presentation.xml.rels and map rId -> Target
    const presentationRelPath = "ppt/_rels/presentation.xml.rels";
    const presentationRelXML = await this.parseXML(presentationRelPath);

    const rels = presentationRelXML?.["Relationships"]?.["Relationship"] ?? [];
    const rIdToTargetMap = {};

    for (const rel of rels) {
      const attrs = rel?.["$"];
      if (
        attrs &&
        attrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
      ) {
        rIdToTargetMap[attrs.Id] = attrs.Target; // rId5 → slides/slide1.xml
      }
    }

    // Step 3: Build full slide paths (with "ppt/" prefix)
    const slidePaths = rIds
      .map(rid => rIdToTargetMap[rid])
      .filter(Boolean)
      .map(target => this.fixTargetPath(target));

    for (const slidePath of slidePaths) {
      const slideHTML = await this.convertSlideToHTML(slidePath, flag);
      if (slideHTML) slideHTMLs.push(slideHTML);
    }
    return slideHTMLs;
  }

  fixTargetPath(target) {
    if (!target) return null;

    // Convert to string and handle backslashes
    target = String(target).replace(/\\/g, '/');

    // Handle relative paths with ../
    if (target.startsWith('../')) {
      // Remove ../ and build proper path
      target = target.replace(/^\.\.\//, '');

      // If it doesn't start with ppt/, add it
      if (!target.startsWith('ppt/')) {
        target = 'ppt/' + target;
      }
    }

    // Remove leading slashes
    target = target.replace(/^\/+/, '');

    // Remove any double ppt/ prefix
    while (target.includes('ppt/ppt/')) {
      target = target.replace('ppt/ppt/', 'ppt/');
    }

    // Ensure it starts with ppt/
    if (!target.startsWith('ppt/')) {
      target = 'ppt/' + target;
    }

    return normalizePath(target);
  }

  async convertSlideToHTML(slidePath, flag) {
    try {

      sharedCache.slideX = slidePath;

      this.shapeNameOccurrences = {};
      let slideXML = await this.parseXML(slidePath);

      if (!this.isSlideEmpty(slideXML)) {
        const slideName = slidePath.match(/ppt\/slides\/(slide\d+)\.xml/)[1];
        const slideRelsPaths = `ppt/slides/_rels/${slideName}.xml.rels`;

        const slideRelXML = await this.parseXML(slideRelsPaths);

        // DEBUG: Check if slideRelXML was loaded
        if (!slideRelXML) {
          console.error(`Failed to load slide relationships for ${slidePath}`);
          return null;
        }

        // Get slide layout path first (simpler relationship)
        const sldLayoutPath = await this.getSlideLayoutFromRels(slideRelXML);

        if (!sldLayoutPath) {
          console.error(`No slide layout found for ${slidePath}`);
          return null;
        }

        let layoutPath = this.fixTargetPath(sldLayoutPath);

        if (!layoutPath) {
          console.error(`Invalid layout path for ${slidePath}. Original: ${sldLayoutPath}`);
          return null;
        }

        let layoutXML = await this.parseXML(layoutPath);

        if (!layoutXML) {
          console.error(`Failed to parse layout XML at ${layoutPath}`);
          return null;
        }

        // Get slide master from layout relationships
        const layoutFileName = layoutPath.split("/").pop().replace(".xml", "");
        let layoutPathrel = `ppt/slideLayouts/_rels/${layoutFileName}.xml.rels`;

        let layoutRelationship = await this.parseXML(layoutPathrel);

        if (!layoutRelationship) {
          console.error(`Failed to parse layout relationship file at ${layoutPathrel}`);
          return null;
        }

        // Get slide master path from layout relationships
        const slideMasterPath = await this.getSlideMasterFromLayoutRels(layoutRelationship);

        if (!slideMasterPath) {
          console.error(`No slide master found in layout relationships for ${layoutPath}`);
          return null;
        }

        const masterPath = this.fixTargetPath(slideMasterPath);

        if (!masterPath) {
          console.error(`Invalid master path. Original: ${slideMasterPath}`);
          return null;
        }

        const masterXML = await this.parseXML(masterPath);

        if (!masterXML) {
          console.error(`Failed to parse master XML at ${masterPath}`);
          return null;
        }

        const MasterName = masterPath.match(/ppt\/slideMasters\/(slideMaster\d+)\.xml/)[1];

        const MasterRelFile = `ppt/slideMasters/_rels/${MasterName}.xml.rels`;
        const MasterRelFileXml = await this.parseXML(MasterRelFile);

        const themePath = await this.getThemeFilePathFromMasterRels(MasterRelFileXml);
        const themeXML = await this.parseXML(themePath);

        this.themeXML = themeXML; // New Change

        // NEW: prefill slide placeholders with layout position + text defaults
        this.enrichSlideFromLayout(slideXML, layoutXML);

        const presentationPath = "ppt/presentation.xml";
        const presentationXML = await this.parseXML(presentationPath);
        const slideSize = await this.getSlideSize(presentationXML);

        let relPath;
        if (slidePath.endsWith(".rels")) {
          relPath = slidePath; // It's already a .rels file
        } else {
          const relDir = path.posix.dirname(slidePath);
          const relFile = `${path.posix.basename(slidePath)}.rels`;
          relPath = path.posix.join(relDir, "_rels", relFile);
        }

        const relationshipsXML = await this.parseXML(relPath);

        // Get color mapping from the master slide
        const clrMap = await pptBackgroundColors.getColorMappingFromMaster(masterXML);

        let slideBgResult = await pptBackgroundColors.getBackgroundColor(slideXML, masterXML, themeXML, relationshipsXML, this, layoutXML);
        let slideBg = slideBgResult.backgroundCSS;

        // getLayoutBackgroundColor
        let slideLayoutBg
        if (slideBg === "#000000" || !slideBg) {
          slideLayoutBg = await pptLayoutBackgroundColors.getLayoutBackgroundColor(layoutXML, themeXML, layoutRelationship, masterXML, this);
        }

        if (slideBg === "#000000" || slideBg == "") {
          slideBg = "";
        }

        let bgInsets = slideBgResult.insets;
        let transparency = slideBgResult.transparency || 0;

        let htmlContent = '';

        if (parseInt(flag) === 1) {
          htmlContent = `<div class="sli-slide"
            data-original-width="${slideSize.width}"
            data-original-height="${slideSize.height}"
            data-slide-xml="${slidePath}"
            style="position:relative;
           scale: 0.9;
           width:${slideSize.width}px;
           height:${slideSize.height}px;">`;
        } else {
          htmlContent = `<div class="sli-slide"
          data-original-width="${slideSize.width}"
          data-original-height="${slideSize.height}"
          data-slide-xml="${slidePath}"
          style="position:relative;
           scale: 0.9;
           width:${slideSize.width}px;
           height:${slideSize.height}px;">`;
        }

        if (slideBg) {
          // Apply opacity to the background
          const opacityValue = transparency > 0 ? (100 - transparency) / 100 : 1;
          const opacityStyle = `opacity: ${opacityValue.toFixed(2)};`;

          if (slideBg.startsWith('url(')) {
            // For background images - use consistent approach for all flags
            htmlContent += `<div class="sli-background" 
                  style="position:absolute; 
                  top:0; 
                  left:0; 
                  width:${slideSize.width}px; 
                  height:${slideSize.height}px;
                  background-image: ${slideBg}; 
                  background-size: cover;
                  background-position: center center;
                  background-repeat: no-repeat;
                  ${opacityStyle}
                  z-index:1;"></div>`;
          } else {
            // For solid colors or gradients, apply styling based on flag
            if (parseInt(flag) === 1) {
              // Flag = 1: Use fixed dimensions
              htmlContent += `<div class="sli-background" 
                      style="position:absolute; 
                      top:0; 
                      left:0; 
                      width:${slideSize.width}px; 
                      height:${slideSize.height}px;
                      background: ${slideBg}; 
                      ${opacityStyle}
                      z-index:1;"></div>`;
            } else {
              // Flag = 0: Use aspect ratio
              htmlContent += `<div class="sli-background" 
                      style="position:absolute; 
                      top:0; 
                      left:0; 
                      width:${slideSize.width}px; 
                      height:${slideSize.height}px;
                      margin: auto;
                      aspect-ratio:${slideSize.aspectRatio};
                      background: ${slideBg}; 
                      ${opacityStyle}
                      z-index:1;"></div>`;
            }
          }
        } else if (slideLayoutBg) {
          // For solid colors or gradients, apply styling based on flag
          if (parseInt(flag) === 1) {
            // Flag = 1: Use fixed dimensions
            htmlContent += `<div class="layout-background" 
                      style="position:absolute; 
                      top:0; 
                      left:0; 
                      width:${slideSize.width}px; 
                      height:${slideSize.height}px;
                      background: ${slideLayoutBg.backgroundCSS}; 
                      z-index:1;"></div>`;
          } else {
            // Flag = 0: Use aspect ratio
            htmlContent += `<div class="layout-background" 
                      style="position:absolute; 
                      top:0; 
                      left:0; 
                      width:${slideSize.width}px; 
                      height:${slideSize.height}px;
                      margin: auto;
                      aspect-ratio:${slideSize.aspectRatio};
                      background: ${slideLayoutBg.backgroundCSS}; 
                      z-index:1;"></div>`;
          }
        } else {
          // Handled background color from Master.
          const masterBgColor = await masterBackgroundHandler.getMasterBackground(masterXML, themeXML, MasterRelFileXml, layoutXML, this);

          let backgroundStyle = '';

          if (masterBgColor && masterBgColor.type) {
            switch (masterBgColor.type) {
              case 'solid':
                // For solid colors, use the color property
                backgroundStyle = `background-color: ${masterBgColor.color || '#FFFFFF'};`;
                break;

              case 'gradient':
                // For gradients, use the css property
                backgroundStyle = `background: ${masterBgColor.css};`;
                break;

              case 'image':
                // For images, use the complete CSS string which includes all image properties
                backgroundStyle = masterBgColor.css;
                break;

              case 'pattern':
                // For patterns, use the css property (fallback to bgColor)
                backgroundStyle = `background: ${masterBgColor.css};`;
                break;

              case 'none':
                // For transparent backgrounds
                backgroundStyle = 'background: transparent;';
                break;

              default:
                // Fallback to white background
                backgroundStyle = 'background-color: #FFFFFF;';
                console.warn('Unknown master background type:', masterBgColor.type);
            }
          } else {
            // Fallback if no result or invalid result
            backgroundStyle = 'background-color: #FFFFFF;';
            console.warn('No valid master background result, using white fallback');
          }

          if (parseInt(flag) === 1) {
            // Flag = 1: Use fixed dimensions
            htmlContent += `<div class="sli-master-background" 
                  style="position:absolute; 
                  top:0; 
                  left:0;
                  width:${slideSize.width}px; 
                  height:${slideSize.height}px;
                  ${backgroundStyle}; 
                  z-index:1;"></div>`;
          } else {
            // Flag = 0: Use aspect ratio
            htmlContent += `<div class="sli-master-background" 
                  style="position:absolute; 
                  top:0; 
                  left:0;
                  width:${slideSize.width}px; 
                  height:${slideSize.height}px;
                  margin: auto;
                  aspect-ratio:${slideSize.aspectRatio}; 
                  ${backgroundStyle}; 
                  z-index:1;"></div>`;
          }
        }

        // Add a content container to hold all slide content
        htmlContent += `<div class="sli-content" style="position:relative; z-index:2;">`;

        // Get Slide Layout background data
        const slideLayoutForm = new SlideLayoutForm(themeXML, this, slidePath, slideRelXML, layoutPath, layoutRelationship, masterXML);

        let htmlContentData = await slideLayoutForm.createSlideLayoutForm(layoutXML, themeXML, this, slidePath, slideRelXML, layoutRelationship, masterXML);

        // if (!htmlContentData === 'undefined') {
        htmlContent += htmlContentData;
        // }

        const regex = /z-index:\s*(\d+)/g;
        let zIndexes = [];
        let match;

        while ((match = regex.exec(htmlContentData)) !== null) {
          // Extract the z-index value and push to the array
          zIndexes.push(parseInt(match[1], 10));
        }

        let highestZIndex = 0;
        if (zIndexes && zIndexes.length > 0) {
          // Find the highest z-index value
          highestZIndex = Math.max(...zIndexes);
        }
        const overrideClrMapping = this.extractOverrideClrMapping(layoutXML);
        const nodes = await this.getNodesInCorrectOrder(slideXML, highestZIndex);

        const zIndexMap = {};
        const nameCounts = {};

        nodes.forEach((node) => {
          if (node.name) {
            // Create a unique key using name + occurrence count
            if (!nameCounts[node.name]) {
              nameCounts[node.name] = 0;
            }

            const occurrence = nameCounts[node.name];
            const uniqueKey = occurrence === 0 ? node.name : `${node.name}__${occurrence}`;
            zIndexMap[uniqueKey] = node.id;
            nameCounts[node.name]++;
          }
        });

        const shapeTree = masterXML["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
        const MasterLineShape = masterXML?.["p:sldMaster"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:cxnSp"] || [];

        // After loading themeXML, masterXML, MasterRelFileXml, slideSize
        const masterHtmlGen = new masterHtml(themeXML, masterXML, MasterRelFileXml, slideSize);

        // FIXED: Pass the required parameters to generateMasterHTML
        const MasterHtmlContent = await masterHtmlGen.generateMasterHTML(shapeTree, MasterLineShape, this);

        // // Append to full slide HTML
        htmlContent += MasterHtmlContent;

        // ----------------------------------------------------------------------------------------------

        // Process text shapes
        const shapeNodes = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
        // Extract line connector
        const lineShapeTag = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:cxnSp"] || [];
        // Extract tables from the slide XML
        const tableNodes = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:graphicFrame"] || [];
        // Process pictures and SVGs
        const picNodesImg = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:pic"] || [];

        // Create a ShapeHandler instance
        const shapeHandler = new ShapeHandler(themeXML, clrMap, nodes, this, slidePath, relationshipsXML, masterXML, layoutXML, flag);

        htmlContent += await shapeHandler.convertAllShapesToHTML(shapeNodes, lineShapeTag, tableNodes, themeXML, layoutXML, zIndexMap);

        // ==================== FIXED IMAGE PROCESSING LOOP ====================
        for (const picNode of picNodesImg) {
          const nodeName = picNode["p:nvPicPr"][0]["p:cNvPr"][0].$.name;

          if (!this.shapeNameOccurrences) {
            this.shapeNameOccurrences = {};
          }
          if (!this.shapeNameOccurrences[nodeName]) {
            this.shapeNameOccurrences[nodeName] = 0;
          }
          const occurrence = this.shapeNameOccurrences[nodeName];
          const uniqueKey = occurrence === 0 ? nodeName : `${nodeName}__${occurrence}`;
          this.shapeNameOccurrences[nodeName]++;
          const zIndex = zIndexMap[uniqueKey] || 0;

          // FIRST: Extract raw dimensions directly from XML to detect lines
          const xfrm = picNode?.["p:spPr"]?.[0]?.["a:xfrm"]?.[0];
          const rawWidth = xfrm?.["a:ext"]?.[0]?.["$"]?.cx;
          const rawHeight = xfrm?.["a:ext"]?.[0]?.["$"]?.cy;

          // Convert EMU to pixels (1 EMU = 1/12700 pixels at 72 DPI)
          const actualWidth = rawWidth ? parseFloat(rawWidth) / 12700 : null;
          const actualHeight = rawHeight ? parseFloat(rawHeight) / 12700 : null;

          // Detect line shapes BEFORE calling getPositionFromShapeOrPlaceholder
          const isHorizontalLine = actualHeight && actualWidth && actualHeight < 2 && actualWidth > 10;
          const isVerticalLine = actualHeight && actualWidth && actualWidth < 2 && actualHeight > 10;
          const isLine = isHorizontalLine || isVerticalLine;

          // Get position (but we'll override dimensions for lines)
          const position = pptTextAllInfo.getPositionFromShapeOrPlaceholder(picNode, layoutXML, flag);

          // CRITICAL FIX: For line shapes, use actual XML dimensions instead of processed ones
          if (isLine && actualWidth && actualHeight) {
            // Extract position from xfrm
            const rawX = xfrm?.["a:off"]?.[0]?.["$"]?.x;
            const rawY = xfrm?.["a:off"]?.[0]?.["$"]?.y;

            if (rawX && rawY) {
              position.x = parseFloat(rawX) / 12700;
              position.y = parseFloat(rawY) / 12700;
            }

            // Use ACTUAL dimensions from XML, not processed ones
            position.width = actualWidth;
            position.height = actualHeight;
          }

          // Check if this picture is a placeholder
          const isPicPlaceholder = Boolean(picNode?.["p:nvPicPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]);
          const picPlaceholderClass = isPicPlaceholder ? ' placeholder-picture' : '';

          // Set rendering dimensions (enforce minimum 1px for lines)
          let renderWidth = position.width;
          let renderHeight = position.height;

          if (isHorizontalLine) {
            renderHeight = Math.max(position.height, 1);
            renderWidth = position.width; // Keep actual width
          } else if (isVerticalLine) {
            renderWidth = Math.max(position.width, 1);
            renderHeight = position.height; // Keep actual height
          }

          // Skip only if BOTH dimensions are invalid (not line shapes)
          if (!isLine && (position.width <= 1 || position.height <= 1)) {
            console.warn(`Skipping ${nodeName} - invalid dimensions (not a line)`);
            continue;
          }

          const blipNode = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0];
          const svgContent = await this.getSVGContent(blipNode, relationshipsXML);

          console.log("     theme File **** -- > ", themeXML);

          const imageInfo = await this.getImageFromPicture(picNode, slidePath, relationshipsXML, themeXML);

          console.log(" @@@@@@##---##@@@@@@@ Extracting image from picture node... -->", imageInfo);

          const matchingNode = nodes.find(node => node.name === nodeName);
          // Fetch imgProps, blurEffect, and blurAmount
          const imgProps = picNode?.["p:blipFill"]?.[0]?.["a:blip"]?.[0]?.["a:extLst"]?.[0]?.["a:ext"]?.find(
            (ext) => ext?.["a14:imgProps"]?.[0]
          )?.["a14:imgProps"];

          let blurEffect = null;
          let blurAmount = null;
          let cssFilters = []; // Collect all filters in an array

          if (imgProps && Array.isArray(imgProps)) {
            const imgLayers = imgProps
              .map((prop) => prop?.["a14:imgLayer"])
              .flat()
              .filter((layer) => layer); // Remove undefined or null layers

            for (const imgLayer of imgLayers) {
              const imgEffects = imgLayer?.["a14:imgEffect"] || [];
              for (const imgEffect of imgEffects) {
                // Check for artisticBlur effect and add to CSS filters
                if (imgEffect?.["a14:artisticBlur"]) {
                  const blurNode = imgEffect["a14:artisticBlur"];
                  const blurValue = blurNode?.["$"]?.intensity || "5px"; // Use dynamic value or fallback
                  cssFilters.push(`blur(${blurValue})`);
                  blurEffect = "artisticBlur";
                }

                // Check for sharpenSoften effect and add to CSS filters
                if (imgEffect?.["a14:sharpenSoften"]) {
                  blurAmount = imgEffect["a14:sharpenSoften"]?.[0]?.["$"]?.amount;

                  // Map blurAmount to contrast range [0.5, 2]
                  const contrastValue = Math.max(0.5, Math.min(2, 1 + parseInt(blurAmount, 10) / 200000));
                  cssFilters.push(`contrast(${contrastValue})`);
                }
              }
            }
          }

          // Combine all filter effects into a single CSS filter string
          const combinedFilter = cssFilters.length > 0 ? `filter: ${cssFilters.join(" ")};` : `filter:blur(0px) contrast(1);`;

          const rotation = pptTextAllInfo.getRotation(picNode);
          const imgcss = imgSvgStyle.returnImgSvgStyle(picNode);

          if (!imageInfo && svgContent) {
            htmlContent += `<div class="sli-svg-container${picPlaceholderClass}" data-name="${nodeName}"
                              style="position:absolute; 
                              left:${position.x}px; 
                              top:${position.y}px; 
                              width:${renderWidth}px; 
                              height:${renderHeight}px;
                              transform: ${imgcss.flipH} ${imgcss.flipV} rotate(${rotation}deg);
                              overflow:hidden;
                              z-index:${zIndex};">
                              <svg style="width:100%; height:100%;" xmlns="http://www.w3.org/2000/svg">${svgContent}</svg>
                          </div>`;
          } else if (imageInfo) {
            // Apply border-radius for ellipse detection
            const shapeStyle = imageInfo.shape === "ellipse" ? "border-radius: 50%; overflow: hidden;" : "";
            // const borderCSS = imageInfo.border.width > 0 ? `border: ${imageInfo.border.width}px solid ${imageInfo.border.color};` : "";
            const borderStyle = imageInfo.border.style || "solid"; // Fallback to solid
            const borderCSS =
              imageInfo.border.width > 0
                ? `border: ${imageInfo.border.width}px ${borderStyle} ${imageInfo.border.color};`
                : "";
            // ✅ ADD THIS LINE: Prepare hyperlink attribute
            const hyperlinkAttr = imageInfo.hyperlink ? `data-hyperlink="${imageInfo.hyperlink}"` : "";

            // Generate shadow CSS from imageInfo
            let boxShadowCSS = "";
            if (imageInfo.hasShadow && imageInfo.shadow) {
              const shadows = [];

              // Add glow effect (rendered as multiple box-shadows with 0 offset)
              if (imageInfo.shadow.glow) {
                const { radius, color } = imageInfo.shadow.glow;

                // PowerPoint glow uses multiple layers for a soft, diffused effect
                // We need more layers and bigger spread
                shadows.push(`0 0 ${(radius * 0.5).toFixed(2)}px ${color}`);
                shadows.push(`0 0 ${radius.toFixed(2)}px ${color}`);
                shadows.push(`0 0 ${(radius * 1.5).toFixed(2)}px ${color}`);
                shadows.push(`0 0 ${(radius * 2).toFixed(2)}px ${color}`);
                shadows.push(`0 0 ${(radius * 2.5).toFixed(2)}px ${color}`);
              }

              // Add outer shadow effect
              if (imageInfo.shadow.shadow) {
                const { blur, offsetX, offsetY, color } = imageInfo.shadow.shadow;
                shadows.push(`${offsetX.toFixed(2)}px ${offsetY.toFixed(2)}px ${blur.toFixed(2)}px ${color}`);
              }

              if (shadows.length > 0) {
                boxShadowCSS = `box-shadow: ${shadows.join(", ")};`;
                console.log("Generated Shadow CSS:", boxShadowCSS);
              }
            }
            const finalFilter = imgcss.blurAmount ? `filter: blur(${imgcss.blurAmount}px) contrast(${imgcss.contrastValue || 1});` : combinedFilter;

            // For lines: use object-fit fill and ensure no distortion
            const objectFit = isLine ? "fill" : "cover";

            // Only create anchor wrapper if hyperlink exists
            let hyperlinkOpen = "";
            let hyperlinkClose = "";

            if (imageInfo.hyperlink) {
              hyperlinkOpen = `<a href="${imageInfo.hyperlink}" style="display:block; width:100%; height:100%; position:absolute; top:0; left:0;">`;
              hyperlinkClose = "</a>";
              console.log("✅ Wrapping image in anchor tag:", imageInfo.hyperlink);
            }

            // ✅ SIMPLE ONCLICK APPROACH - No anchor tag needed
            const onclickAttr = imageInfo.hyperlink ? `onclick="window.location.href='${imageInfo.hyperlink}'"` : "";
            const cursorStyle = imageInfo.hyperlink ? "cursor: pointer;" : "";

            // ✅ Simplified - show image first, then add cropping
            let imgStyle = `position:absolute; width:100%; height:100%; object-fit: ${objectFit}; opacity:${imgcss.opacity}; pointer-events: none; ${finalFilter}`;

            console.log("Image:", imageInfo.src);
            console.log("Has cropping?", imageInfo.hasCropping);
            console.log("Cropping data:", imageInfo.cropping);

            if (imageInfo.hasCropping && imageInfo.cropping) {
              const crop = imageInfo.cropping;

              const cropL = crop.left || 0;
              const cropR = crop.right || 0;
              const cropT = crop.top || 0;
              const cropB = crop.bottom || 0;

              const visibleW = 100 - cropL - cropR;
              const visibleH = 100 - cropT - cropB;

              if (visibleW > 0 && visibleH > 0) {
                const scaleX = 100 / visibleW;
                const scaleY = 100 / visibleH;

                const offsetX = -(cropL / visibleW) * 100;
                const offsetY = -(cropT / visibleH) * 100;

                imgStyle = `
                      position:absolute;
                      width:${(scaleX * 100).toFixed(3)}%;
                      height:${(scaleY * 100).toFixed(3)}%;
                      left:${offsetX.toFixed(3)}%;
                      top:${offsetY.toFixed(3)}%;
                      opacity:${imgcss.opacity};
                      pointer-events:none;
                      ${finalFilter}`;
              }
            }

            htmlContent += `
                      <div class="image-container${picPlaceholderClass}${isLine ? " line-image" : ""}" 
                          srcrectl="${imageInfo.cropping?.l || 0}" 
                          srcrectr="${imageInfo.cropping?.r || 0}" 
                          srcrectt="${imageInfo.cropping?.t || 0}" 
                          srcrectb="${imageInfo.cropping?.b || 0}"
                          data-name="${nodeName}"
                          data-alt-text="${imageInfo.altText || ""}"
                          ${imageInfo.hyperlink ? `data-hyperlink="${imageInfo.hyperlink}"` : ""}
                          ${onclickAttr}
                          phType="${position.phType || ""}" 
                          phIdx="${position.phIdx || ""}" 
                          data-is-line="${isLine}"
                          data-actual-width="${actualWidth}"
                          data-actual-height="${actualHeight}"
                          style="position:absolute;
                              left:${position.x}px;
                              top:${position.y}px;
                              width:${renderWidth}px;
                              height:${renderHeight}px;
                              ${shapeStyle}
                              ${borderCSS}
                              ${boxShadowCSS}
                              transform: ${imgcss.transform} rotate(${rotation}deg);
                              overflow:hidden;
                              ${cursorStyle}
                              z-index:${zIndex};">
                          
                          <img src="${imageInfo.src}" alt="${nodeName}"
                              style="${imgStyle}" />
                        </div>`;
          } else {
            // No image info and no SVG content
          }

        }
        // ==================== END FIXED IMAGE PROCESSING LOOP ====================

        htmlContent += "</div>"; // Close sli-content div

        let htmlcontentString = await this.cleanGeneratedHTML(htmlContent);

        return htmlcontentString;
      }
    } catch (error) {
      console.error('Error occurred:', error);

      // Create a standardized error response object
      const errorResponse = {
        success: false,
        error: {
          message: error.message || 'An unexpected error occurred',
          code: error.code || 'UNKNOWN_ERROR',
          status: error.status || 500
        }
      };

      // Return the error response instead of trying to use res.status
      return errorResponse;
    }
  }

  /**
   * FIXED: New method to get slide layout directly from slide relationships
   */
  async getSlideLayoutFromRels(slideRelXML) {
    if (!slideRelXML || !slideRelXML.Relationships) {
      console.error("Invalid slideRelXML structure");
      return null;
    }

    const relationships = slideRelXML.Relationships.Relationship || [];

    for (const rel of relationships) {
      const attrs = rel.$ || {};

      // Look for slide layout relationship
      if (attrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
        // console.log("Found slide layout relationship:", attrs.Target);
        return attrs.Target; // Returns something like "../slideLayouts/slideLayout2.xml"
      }
    }

    console.error("No slide layout relationship found in relationships");
    return null;
  }

  /**
   * FIXED: New method to get slide master from layout relationships
   */
  async getSlideMasterFromLayoutRels(layoutRelXML) {
    if (!layoutRelXML || !layoutRelXML.Relationships) {
      console.error("Invalid layoutRelXML structure");
      return null;
    }

    const relationships = layoutRelXML.Relationships.Relationship || [];

    for (const rel of relationships) {
      const attrs = rel.$ || {};

      // Look for slide master relationship
      if (attrs.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
        // console.log("Found slide master relationship:", attrs.Target);
        return attrs.Target; // Returns something like "../slideMasters/slideMaster1.xml"
      }
    }

    console.error("No slide master relationship found in layout relationships");
    return null;
  }

  /**
   * FIXED: Get theme file path from master relationships
   */
  async getThemeFilePathFromMasterRels(masterRelXML) {
    try {
      if (!masterRelXML || !masterRelXML.Relationships || !masterRelXML.Relationships.Relationship) {
        console.warn('Invalid master relation XML structure');
        return "ppt/theme/theme1.xml"; // Default fallback
      }

      const relationships = masterRelXML.Relationships.Relationship;

      // Find the theme relationship
      for (const relationship of relationships) {
        const type = relationship.$.Type;

        if (type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {

          const themeTarget = relationship.$.Target;

          const fullThemePath = this.fixTargetPath(themeTarget);
          return fullThemePath;
        }
      }

      // If no theme relationship found, return default
      console.warn('No theme relationship found, using default theme1.xml');
      return "ppt/theme/theme1.xml";

    } catch (error) {
      console.error('Error extracting theme file path:', error);
      return "ppt/theme/theme1.xml"; // Default fallback
    }
  }

  /**
   * Check if slide is empty
   */
  isSlideEmpty(slideXML) {
    if (!slideXML) return true;

    const spTree = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0];
    if (!spTree) return true;

    // Check if there are any shapes, pictures, or other content
    const hasShapes = (spTree["p:sp"] && spTree["p:sp"].length > 0);
    const hasPictures = (spTree["p:pic"] && spTree["p:pic"].length > 0);
    const hasGraphics = (spTree["p:graphicFrame"] && spTree["p:graphicFrame"].length > 0);
    const hasConnectors = (spTree["p:cxnSp"] && spTree["p:cxnSp"].length > 0);

    return !(hasShapes || hasPictures || hasGraphics || hasConnectors);
  }

  async getSlideSize(presentationXML) {
    const sldSz = presentationXML?.["p:presentation"]?.["p:sldSz"]?.[0]?.["$"];

    if (!sldSz) {

      return {
        width: 960,
        height: 540,
        aspectRatio: "16/9"
      };
    }

    const width = parseInt(sldSz.cx) / 12700;
    const height = parseInt(sldSz.cy) / 12700;
    const aspectRatio = `${width}/${height}`;

    return { width, height, aspectRatio };
  }

  async cleanGeneratedHTML(htmlContent) {

    htmlContent = htmlContent.replace(/\s+/g, ' ');
    htmlContent = htmlContent.replace(/\s+>/g, '>');
    htmlContent = htmlContent.replace(/>\s+/g, '>');

    if (!htmlContent.endsWith('</div>')) {
      htmlContent += '</div>';
    }

    return htmlContent;
  }

  enrichSlideFromLayout(slideXML, layoutXML) {
    try {

      const layoutShapes = layoutXML?.["p:sldLayout"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];
      const slideShapes = slideXML?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.["p:sp"] || [];

      const clone = (obj) => JSON.parse(JSON.stringify(obj));

      for (const sShape of slideShapes) {
        const ph = sShape?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"];
        if (!ph) continue;

        const phType = ph.type || "body";
        const phIdx = ph.idx;

        // Find matching layout placeholder
        const lShape = layoutShapes.find(ls => {
          const lph = ls?.["p:nvSpPr"]?.[0]?.["p:nvPr"]?.[0]?.["p:ph"]?.[0]?.["$"];
          if (!lph) return false;
          const lType = lph.type || "body";
          const lIdx = lph.idx;
          return lType === phType && lIdx === phIdx;
        });

        if (!lShape) continue;

        // 1) Transform (xfrm) – copy if slide has none
        if (!sShape["p:spPr"]?.[0]?.["a:xfrm"] && lShape["p:spPr"]?.[0]?.["a:xfrm"]) {
          if (!sShape["p:spPr"]) sShape["p:spPr"] = [{}];
          sShape["p:spPr"][0]["a:xfrm"] = clone(lShape["p:spPr"][0]["a:xfrm"]);
        }

        // 2) Text body defaults
        const sTxBody = sShape["p:txBody"]?.[0];
        const lTxBody = lShape["p:txBody"]?.[0];
        if (sTxBody && lTxBody) {
          // bodyPr
          if (!sTxBody["a:bodyPr"] && lTxBody["a:bodyPr"]) {
            sTxBody["a:bodyPr"] = clone(lTxBody["a:bodyPr"]);
          }

          // lstStyle (might have level-based defaults)
          if (!sTxBody["a:lstStyle"] && lTxBody["a:lstStyle"]) {
            sTxBody["a:lstStyle"] = clone(lTxBody["a:lstStyle"]);
          }
        }

        // 3) Paragraph/run defaults from layout
        const lvl1 = lTxBody?.["a:lstStyle"]?.[0]?.["a:lvl1pPr"]?.[0];
        const lvl1DefRPr = lvl1?.["a:defRPr"]?.[0];
        const lvl1DefRPr$ = lvl1DefRPr?.["$"] || {};

        const paragraphs = sTxBody?.["a:p"] || [];
        for (const p of paragraphs) {
          if (!p['a:pPr']) p['a:pPr'] = [{}];
          const pPr = p['a:pPr'][0];

          // paragraph alignment from lvl1
          if (lvl1?.['$']?.algn) {
            if (!pPr['$']) pPr['$'] = {};
            if (pPr['$'].algn == null) pPr['$'].algn = lvl1['$'].algn;
          }

          // default run props on paragraph if missing
          if (!pPr['a:defRPr'] && lvl1DefRPr) pPr['a:defRPr'] = [clone(lvl1DefRPr)];

          // Apply defaults to runs if they lack them
          const runs = p['a:r'] || [];
          for (const r of runs) {
            if (!r['a:rPr']) r['a:rPr'] = [{}];
            const rPr = r['a:rPr'][0];
            if (!rPr['$']) rPr['$'] = {};

            if (rPr['$'].sz == null && lvl1DefRPr$.sz != null) rPr['$'].sz = lvl1DefRPr$.sz;
            if (rPr['$'].b == null && lvl1DefRPr$.b != null) rPr['$'].b = lvl1DefRPr$.b;
            if (rPr['$'].i == null && lvl1DefRPr$.i != null) rPr['$'].i = lvl1DefRPr$.i;

            // inherit color if run has none
            if (!rPr['a:solidFill'] && lvl1DefRPr?.['a:solidFill']) {
              rPr['a:solidFill'] = clone(lvl1DefRPr['a:solidFill']);
            }
          }
        }

        // 4) p:style (scheme refs) – copy only if slide lacks it
        if (!sShape['p:style'] && lShape['p:style']) {
          sShape['p:style'] = clone(lShape['p:style']);
        }
      }
    } catch (e) {
      console.error('enrichSlideFromLayout error:', e);
    }
  }

  async getNodesInCorrectOrder(xmlFile, highestZIndex) {
    try {
      // Get the shape tree from the slide XML
      const spTree = xmlFile['p:sld']['p:cSld'][0]['p:spTree'][0];
      const orderedNodes = [];
      let sequentialId = highestZIndex + 1;

      // Helper: traverse a list of nodes in order
      function traverseNodes(nodeList) {
        nodeList.forEach(node => {
          const nodeType = node['#name'];

          // ✅ FIXED: Added p:graphicFrame to the list
          if (nodeType === 'p:sp' ||
            nodeType === 'p:cxnSp' ||
            nodeType === 'p:pic' ||
            nodeType === 'p:grpSp' ||
            nodeType === 'p:graphicFrame') {  // ✅ Added this line

            // Create node info with sequential z-index ID
            const nodeInfo = { type: nodeType, id: sequentialId++ };

            if (node.$$) {
              const nvPrNode = node.$$.find(child =>
                child['#name'].includes('nvGrpSpPr') ||
                child['#name'].includes('nvSpPr') ||
                child['#name'].includes('nvPicPr') ||
                child['#name'].includes('nvCxnSpPr') ||
                child['#name'].includes('nvGraphicFramePr')
              );

              if (nvPrNode && nvPrNode.$$) {
                const cNvPr = nvPrNode.$$.find(n => n['#name'] === 'p:cNvPr');
                if (cNvPr && cNvPr.$ && cNvPr.$.name) {
                  nodeInfo.name = cNvPr.$.name;
                }
              }
            }

            orderedNodes.push(nodeInfo);

            if (nodeType === 'p:grpSp' && node.$$) {
              traverseNodes(node.$$);
            }
          }
        });
      }

      if (spTree.$$) {
        traverseNodes(spTree.$$);
      }

      return orderedNodes;
    } catch (error) {
      console.error('Error:', error);
      throw error;
    }
  }

  extractOverrideClrMapping(layoutXML) {

    const clrMapOvr = layoutXML?.['p:sldLayout']?.['p:clrMapOvr']?.[0]?.['a:overrideClrMapping'];
    if (clrMapOvr) {
      return clrMapOvr[0].$;
    }
    return null;
  }

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

    const svgFile = this.files[svgPath];
    if (!svgFile) {
      return null;
    }

    return await svgFile.async("string");
  }
}

module.exports = pptxToHtml;