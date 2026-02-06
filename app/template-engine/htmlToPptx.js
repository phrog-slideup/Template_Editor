const PptxGenJS = require("pptxgenjs");
const jsdom = require("jsdom");
const { JSDOM } = jsdom;
const path = require("path");
const fs = require("fs");
const fsPromises = require("fs").promises; 
const puppeteer = require("puppeteer");
const axios = require("axios");
const JSZip = require('jszip');

const addShapeToSlide = require("../html-To-Pptx-Styling/addShapeToSlide");
const addTextBox = require("../html-To-Pptx-Styling/addTextBoxToSlide");
const addTable = require("../html-To-Pptx-Styling/addTableToSlide");
const addSlideLayout = require("../html-To-Pptx-Styling/addLayoutToSlide");
const addImage = require("../html-To-Pptx-Styling/addImageToSlide");
const gradientColor = require("../api/helper/parseGradient.js");
const clrHelper = require("../api/helper/colorHelper.js");
const addChartToSlide = require("../html-To-Pptx-Styling/addChartToSlide");
const sharedCache = require('../api/shared/cache.js');

async function convertHTMLToPPTX(htmlString, outputFilePath, originalFolderName) {
    try {
        console.log('🚀 Starting HTML to PPTX conversion with FIXED master support...', originalFolderName);

        const dom = new JSDOM(htmlString);
        const document = dom.window.document;
        const slides = document.querySelectorAll(".sli-slide:not(.sli-slide .sli-slide)");

        const hasPlaceholders = Array.from(slides).some(slide =>
            slide.querySelector('.placeholder-picture, .placeholder-text'));

        console.log(`🎯 Detected ${hasPlaceholders ? 'placeholders' : 'no placeholders'} in presentation`);

        if (slides.length === 0) {
            throw new Error("No slides found in the provided HTML.");
        }

        // Extract slide dimensions
        const firstSlide = slides[0];
        const extractedDimensions = extractSlideDimensions(firstSlide);
        const slideDimensions = validateSlideDimensions(extractedDimensions);

        // Initialize PowerPoint without masters (we'll inject them later)
        const pptx = new PptxGenJS();

        // console.log(pptx.)
        if (pptx._chartCounter !== undefined) {
            pptx._chartCounter = 0;
        }

        if (slideDimensions.isValid) {
            try {
                pptx.defineLayout({
                    name: 'CUSTOM_SLIDE',
                    width: slideDimensions.width,
                    height: slideDimensions.height
                });
                pptx.layout = 'CUSTOM_SLIDE';
            } catch (layoutError) {
                console.error('   ⚠️ Error setting custom layout, using default');
            }
        }

        for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
            const slideElement = slides[slideIndex];
            const currentSlideDimensions = extractSlideDimensions(slideElement);
            const currentValidatedDimensions = validateSlideDimensions(currentSlideDimensions);

            const slideContext = {
                dimensions: currentValidatedDimensions,
                slideIndex,
                totalSlides: slides.length,
                scaleX: 1,
                scaleY: 1
            };

        }

        // STEP 2: Process slides normally (masters will be injected later)
        console.log('📄 Step 2: Processing slides...');

        for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
            try {
                const slideElement = slides[slideIndex];
                const currentSlideDimensions = extractSlideDimensions(slideElement);
                const currentValidatedDimensions = validateSlideDimensions(currentSlideDimensions);

                const slideContext = {
                    dimensions: currentValidatedDimensions,
                    slideIndex,
                    totalSlides: slides.length,
                    scaleX: 1,
                    scaleY: 1
                };

                // Create slide
                const pptSlide = pptx.addSlide();

                // Set background
                try {
                    await setBackground(pptx, slideElement, pptSlide, slideContext);
                } catch (bgError) {
                    console.error(`   ❌ Error setting background: ${bgError.message}`);
                }

                // Process regular slide content (excluding masters)
                await processSlideContent(pptx, pptSlide, slideElement, slideContext);

            } catch (slideError) {
                console.error(`❌ Error processing slide ${slideIndex + 1}:`, slideError);
            }
        }

        // STEP 3: Write initial PPTX file
        console.log('💾 Step 3: Writing initial PPTX file...');

        await pptx.writeFile({ fileName: outputFilePath });

        // STEP 5: Final cleanup
        await fixPptxFile(outputFilePath);

        // STEP 6: Extract and process slideX.xml files with unique directory
        console.log('📜 Step 6: Extracting and processing slideX.xml files...');

        // FIXED: Create unique directory path BEFORE calling extractAndProcessSlideXMLs
        const uniqueDirName = `slide_xmls_${originalFolderName}_${Date.now()}`;
        const slideXmlsDir = path.join(path.dirname(outputFilePath), uniqueDirName);

        const firstSlideEl = document.querySelector(".sli-slide");
        const extractedSlideX = firstSlideEl?.getAttribute("data-slide-xml") || null;
        sharedCache.slideX = extractedSlideX; // ⭐ Update global cache

        console.log("📌 Slide XML detected from HTML:", sharedCache.slideX);

        // FIXED: Pass the unique directory path to extractAndProcessSlideXMLs
        console.log("=====================22=======================");
        console.log("   sharedCache.slideX :- ", sharedCache.slideX);
        console.log("=====================22=======================");
        const extractionResult = await extractAndProcessSlideXMLs(outputFilePath, slideXmlsDir, sharedCache.slideX);

        // Verify extraction was successful
        if (!extractionResult.success) {
            throw new Error('Failed to extract slide XML files');
        }

        // // NEW STEP 6.5: Fix chart XML bugs before replacing files
        // console.log('🔧 Step 6.5: Fixing chart XML bugs...');
        // const chartFixResult = await fixChartXmlBugs(slideXmlsDir);

        // // NEW STEP 6.5A: Fix chart styling (white borders, spacing, colors, grid lines)
        // console.log('🎨 Step 6.5A: Fixing chart styling...');
        // const chartStylingResult = await fixChartStyling(slideXmlsDir);
        // if (chartStylingResult.success && chartStylingResult.chartsFixed > 0) {
        //     console.log(`   ✅ Fixed styling in ${chartStylingResult.chartsFixed} charts`);
        // } else if (chartStylingResult.success) {
        //     console.log('   ℹ️  No chart styling issues found');
        // } else {
        //     console.log('   ⚠️  Chart styling fix had issues:', chartStylingResult.error);
        // }

        // // NEW STEP 6.6: Comprehensive chart XML fix for Syncfusion
        // console.log('🔧 Step 6.6: Comprehensive chart XML fix...');
        // const comprehensiveChartResult = await comprehensiveChartXmlFix(slideXmlsDir);

        // NEW STEP 6.7: Fix chart relationships
        console.log('🔗 Step 6.7: Fixing chart relationships...');
        const chartRelsResult = await fixChartRelationships(slideXmlsDir);

        // NEW STEP 6.8: Validate embedded Excel files
        console.log('📊 Step 6.8: Validating embedded Excel...');
        const excelValidation = await validateEmbeddedExcel(slideXmlsDir);

        // NEW STEP 6.9: Clean slide XML for Syncfusion compatibility
        console.log('🧹 Step 6.9: Cleaning slides for Syncfusion compatibility...');
        const slideCleanResult = await cleanSlideXmlForSyncfusion(slideXmlsDir);

        // NEW STEP 6.10: Fix table cell alignment in merged cells
console.log('🔧 Step 6.10: Fixing table cell alignment...');
const alignmentFixResult = await fixTableCellAlignment(slideXmlsDir);

        const fileFolderName = originalFolderName; // e.g., 'Agenda'
        const filesDir = path.resolve(__dirname, '../files');

        // STEP 7: Replace slide1.xml in the PPTX with the converted slide1.xml

        console.log('🔄 Step 7: Replacing slide1.xml in PPTX...', fileFolderName);
        console.log('🔄 Step 7.1: req path...', fileFolderName);

        // Now slideXmlsDir contains the actual directory where files were extracted
        // Replace both slideX.xml and slideX.xml.rels files
        await replaceSlideXMLInPPTX(fileFolderName, slideXmlsDir);

        // Replace Images 
        await replaceSlideImages(fileFolderName, slideXmlsDir);

        // NEW STEP 7.5: Normalize chart references
        console.log('🔄 Step 7.5: Normalizing chart references...');
        await normalizeChartReferences(fileFolderName, slideXmlsDir);


        const sourceFilePath = path.join(filesDir, fileFolderName);

        // Fixed: Specify complete file path instead of just directory
        const zipFileOutput = path.join(filesDir, `${fileFolderName}.zip`);

        // STEP 8: Convert zip to final PPTX file
        console.log('🔄 Step 8: Converting zip to final PPTX...');
        const conversionResult = await convertZipToPptxFile(sourceFilePath, zipFileOutput, fileFolderName);

        console.log("🎉 PPTX file created successfully with proper master support:", conversionResult);

        return JSON.stringify({
            success: true,
            fileName: conversionResult.finalPptxPath, // Use the final converted PPTX path
            finalFileName: conversionResult.fileName, // Just the filename
            message: "PPTX file created successfully with FIXED master support",
        });

    } catch (error) {
        console.error("❌ Error in convertHTMLToPPTX:", error);
        return JSON.stringify({
            success: false,
            error: error.message || "Unknown error in HTML to PPTX conversion",
            details: error.stack
        });
    }
}

function extractSlideDimensions(slideElement) {
    const style = slideElement.getAttribute('style') || '';
    let width = null, height = null;
    let originalWidth = null, originalHeight = null;

    const dataOriginalWidth = slideElement.getAttribute('data-original-width');
    const dataOriginalHeight = slideElement.getAttribute('data-original-height');

    if (dataOriginalWidth && dataOriginalHeight) {
        const originalWidthPx = parseFloat(dataOriginalWidth);
        const originalHeightPx = parseFloat(dataOriginalHeight);
        originalWidth = `${originalWidthPx}px`;
        originalHeight = `${originalHeightPx}px`;
        width = originalWidthPx / 72;
        height = originalHeightPx / 72;
    } else {
        const widthMatch = style.match(/width:\s*([0-9.]+)(px|cm|in|pt|mm)/i);
        const heightMatch = style.match(/height:\s*([0-9.]+)(px|cm|in|pt|mm)/i);

        if (widthMatch && heightMatch) {
            const widthValue = parseFloat(widthMatch[1]);
            const widthUnit = widthMatch[2].toLowerCase();
            const heightValue = parseFloat(heightMatch[1]);
            const heightUnit = heightMatch[2].toLowerCase();

            originalWidth = `${widthValue}${widthUnit}`;
            originalHeight = `${heightValue}${heightUnit}`;

            width = convertToInches(widthValue, widthUnit);
            height = convertToInches(heightValue, heightUnit);
        }
    }

    return {
        width,
        height,
        originalWidth,
        originalHeight,
        aspectRatio: width && height ? width / height : null
    };
}

function validateSlideDimensions(dimensions) {
    const { width, height } = dimensions;
    const MIN_SIZE = 1.0;
    const MAX_SIZE = 56.0;

    let validatedWidth = width;
    let validatedHeight = height;

    if (!width || width < MIN_SIZE || width > MAX_SIZE) {
        validatedWidth = 10.0;
    }

    if (!height || height < MIN_SIZE || height > MAX_SIZE) {
        validatedHeight = 7.5;
    }

    return {
        ...dimensions,
        width: validatedWidth,
        height: validatedHeight,
        aspectRatio: validatedWidth / validatedHeight,
        isValid: !!(width && height),
        isCustomSize: validatedWidth !== 10.0 || validatedHeight !== 7.5
    };
}

function convertToInches(value, unit) {
    switch (unit.toLowerCase()) {
        case 'cm': return value / 2.54;
        case 'mm': return value / 25.4;
        case 'px': return value / 72;
        case 'pt': return value / 72;
        case 'in': return value;
        default: return value / 72;
    }
}

async function setBackground(pptx, slideElement, pptSlide, slideContext = null) {
    const backgroundElement = slideElement.querySelector('.sli-background');

    if (!backgroundElement) {
        pptSlide.background = { fill: "" };
        return;
    }

    const backgroundStyle = backgroundElement.style;
    let backgroundColor = null;
    let backgroundOpacity = 1;

    if (backgroundStyle.opacity) {
        backgroundOpacity = parseFloat(backgroundStyle.opacity);
    }

    const background = backgroundStyle.background || backgroundStyle.backgroundColor;

    if (background) {
        if (background.includes('gradient')) {
            const gradientColor = parseGradientBackground(background);
            if (gradientColor) {
                pptSlide.background = gradientColor;
                return;
            }
        }

        if (background.includes('url(')) {
            const imageUrl = extractBackgroundImageUrl(background);
            if (imageUrl) {
                const resolvedImagePath = resolveImagePath(imageUrl);
                if (resolvedImagePath) {
                    pptSlide.background = {
                        path: resolvedImagePath,
                        sizing: 'cover'
                    };
                    return;
                }
            }
        }

        backgroundColor = extractColor(background);
    }

    if (backgroundColor) {
        const backgroundConfig = { fill: backgroundColor };
        if (backgroundOpacity < 1) {
            backgroundConfig.transparency = Math.round((1 - backgroundOpacity) * 100);
        }
        pptSlide.background = backgroundConfig;
    } else {
        pptSlide.background = { fill: "FFFFFF" };
    }
}

function parseGradientBackground(gradientString) {
    try {
        const colorMatches = gradientString.match(/#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}|rgba?\([^)]+\)|[a-zA-Z]+/g);

        if (colorMatches && colorMatches.length >= 2) {
            const color1 = extractColor(colorMatches[0]);
            const color2 = extractColor(colorMatches[1]);

            if (color1 && color2) {
                return {
                    fill: {
                        type: 'gradient',
                        colors: [
                            { color: color1, position: 0 },
                            { color: color2, position: 100 }
                        ],
                        angle: 90
                    }
                };
            }
        }
        return null;
    } catch (error) {
        console.error('   ❌ Error parsing gradient:', error);
        return null;
    }
}

function extractBackgroundImageUrl(backgroundString) {
    const urlMatch = backgroundString.match(/url\(['"]?([^'"]+)['"]?\)/);
    return urlMatch ? urlMatch[1] : null;
}

function extractColor(colorValue) {
    if (!colorValue || colorValue === 'transparent') return null;

    if (colorValue.startsWith('#')) {
        return colorValue.substring(1);
    }

    const rgbMatch = colorValue.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgbMatch) {
        const r = parseInt(rgbMatch[1]).toString(16).padStart(2, '0');
        const g = parseInt(rgbMatch[2]).toString(16).padStart(2, '0');
        const b = parseInt(rgbMatch[3]).toString(16).padStart(2, '0');
        return `${r}${g}${b}`;
    }

    const namedColors = {
        'white': 'FFFFFF', 'black': '000000', 'red': 'FF0000', 'green': '008000',
        'blue': '0000FF', 'yellow': 'FFFF00', 'cyan': '00FFFF', 'magenta': 'FF00FF',
        'silver': 'C0C0C0', 'gray': '808080', 'grey': '808080', 'maroon': '800000',
        'olive': '808000', 'purple': '800080', 'teal': '008080', 'navy': '000080'
    };

    return namedColors[colorValue.toLowerCase()] || null;
}

async function processSlideContent(pptx, pptSlide, slideElement, slideContext) {
    const processedElements = new Set();
    const contentContainer = slideElement.querySelector('.sli-content') || slideElement;

    let allElements = [];
    const slideChildren = Array.from(slideElement.children);

    if (contentContainer !== slideElement && contentContainer.parentElement === slideElement) {
        allElements = slideChildren
            .filter(child =>
                child !== contentContainer &&
                !child.classList.contains('sli-background') &&
                !isMasterElement(child)
            )
            .map(el => ({
                element: el,
                zIndex: parseFloat(el.style.zIndex) || 0
            }));

        const contentChildren = Array.from(contentContainer.children);
        allElements = allElements.concat(
            contentChildren
                .filter(el => !isMasterElement(el))
                .map(el => ({
                    element: el,
                    zIndex: parseFloat(el.style.zIndex) || 0
                }))
        );
    } else {
        allElements = slideChildren
            .filter(child =>
                !child.classList.contains('sli-background') &&
                !isMasterElement(child)
            )
            .map(el => ({
                element: el,
                zIndex: parseFloat(el.style.zIndex) || 0
            }));
    }

    allElements.sort((a, b) => a.zIndex - b.zIndex);

    for (const { element } of allElements) {
        try {
            if (processedElements.has(element) || isMasterElement(element)) {
                continue;
            }

            // Check for placeholder elements first
            if (isPlaceholderElement(element)) {
                await processPlaceholderElement(pptx, pptSlide, element, slideContext);
                processedElements.add(element);
                continue;
            }

            // NEW: Add chart processing before other element types
            if (addChartToSlide.isChartElement(element)) {
                console.log("Processing chart element:", element.className);
                // const chartProcessed = await addChartToSlide.processChartElement(pptx, pptSlide, element, slideContext);
                const chartProcessed = await addChartToSlide.addChartToSlide(pptx, pptSlide, element, slideContext);
                if (chartProcessed) {
                    console.log("   ✅ Chart processed successfully");
                    processedElements.add(element);
                    continue;
                } else {
                    console.log("   ⚠️ Chart processing failed, trying other processors");
                    // Don't continue here - let it fall through to other processors
                }
            }

            if (element.classList.contains("shape") || element.classList.contains("custom-shape") || element.id === "custGeom" || element.classList.contains("sli-svg-connector")) {
                await processShapeElement(pptx, pptSlide, element, slideContext);
                processedElements.add(element);
            }
            else if (element.classList.contains("sli-layout")) {
                await processLayoutElement(pptx, pptSlide, element, slideContext, processedElements);
            }
            else if (element.classList.contains("image-container")) {
                const imgElement = element.querySelector("img");

                if (imgElement && !isMasterElement(imgElement)) {

                    if (element.classList.contains('placeholder-picture')) {

                        await processPlaceholderElement(pptx, pptSlide, element, slideContext);
                    } else {
                        await addImage.addImageToSlide(pptx, pptSlide, imgElement, slideContext);
                    }
                    processedElements.add(element);
                }
            }
            else if (element.tagName === "IMG") {
                if (element.closest('.placeholder-picture')) {

                    await processPlaceholderElement(pptx, pptSlide, element, slideContext);
                } else {
                    await addImage.addImageToSlide(pptx, pptSlide, element, slideContext);
                }
                processedElements.add(element);
            }
            else if (element.tagName === "DIV" && element.classList.contains('table-container') && !processedElements.has(element)) {
                const table = element.querySelector('table');
                if (table) {
                    console.log('Processing table-container with table element'); // Debug log

                    addTable.addTableToSlide(pptSlide, table, slideContext, element);

                    processedElements.add(element);
                    processedElements.add(table);
                }
            }
            // Keep the original table processing for backward compatibility
            else if (element.tagName === "TABLE" && !processedElements.has(element)) {
                // Only process if not already processed as part of table-container
                const container = element.closest('.table-container');
                if (!container || processedElements.has(container)) {
                    console.log('Processing direct table element'); // Debug log
                    addTable.addTableToSlide(pptSlide, element, slideContext);
                    processedElements.add(element);
                }
            }
        } catch (elementError) {
            console.error(`   ❌ Error processing element:`, elementError);
        }
    }
}

function isMasterElement(element) {
    if (!element || !element.classList) return false;

    // Check for various master class combinations
    const masterClasses = [
        'master',
        'sli-master',
        'custom-shape-master',
        'shape-master',
        'image-master',
        'text-master'
    ];

    // Check if element has any master class
    for (const masterClass of masterClasses) {
        if (element.classList.contains(masterClass)) {
            return true;
        }
    }

    // Check if element is inside a master container
    if (element.closest('.master') ||
        element.closest('.sli-master') ||
        element.closest('.custom-shape-master')) {
        console.log(`   🚫 Filtering out element inside master container`);
        return true;
    }

    // Check for specific master IDs or attributes
    if (element.id && element.id.includes('master')) {
        console.log(`   🚫 Filtering out element with master ID: ${element.id}`);
        return true;
    }

    return false;
}

function isPlaceholderElement(element) {
    return element.classList.contains('placeholder-picture') ||
        element.classList.contains('placeholder-text') ||
        (element.querySelector && (
            element.querySelector('.placeholder-picture') ||
            element.querySelector('.placeholder-text')
        ));
}

async function processPlaceholderElement(pptx, pptSlide, element, slideContext) {
    try {
        // Handle placeholder images
        if (element.classList.contains('placeholder-picture') || element.classList.contains('image-container')) {

            const placeholderTypeDetected = determinePlaceholderType(element, 'image');
            const position = extractElementPosition(element, slideContext);

            // Try to find an actual image inside the container
            const imgElement = element.querySelector('img');

            if (imgElement && imgElement.src && !imgElement.src.startsWith('data:')) {
                const resolvedImagePath = resolveImagePath(imgElement.src);
                if (resolvedImagePath) {
                    // Read attributes (case-insensitive in HTML)
                    const rawPhType = element.getAttribute('phtype') || element.getAttribute('phType') || '';
                    const rawPhIdx = element.getAttribute('phIdx') || element.getAttribute('phidx') || element.getAttribute('data-ph-idx') || '';

                    const srcRectL = element.getAttribute('srcRectL') || element.getAttribute('srcrectl') || '';
                    const srcRectR = element.getAttribute('srcRectR') || element.getAttribute('srcrectr') || '';
                    const srcRectT = element.getAttribute('srcRectT') || element.getAttribute('srcRectt') || '';
                    const srcRectB = element.getAttribute('srcrectB') || element.getAttribute('srcrectb') || '';

                    // Normalize
                    const phType = (rawPhType || 'pic').toString().trim().toLowerCase();
                    const phIdx = rawPhIdx !== '' ? parseInt(rawPhIdx, 10) : 14;

                    // Get transform property for flip (scaleX(-1) for horizontal flip)
                    const transformStyle = element.style.transform || '';
                    const isFlipped = transformStyle.includes('scaleX(-1)');

                    // FIX: Add vertical flip detection

                    const isFlippedVertical = transformStyle.includes('scaleY(-1)');

                    // ✅ NEW: Compute transparency from container + image opacity
                    let containerOpacity = 1;
                    const parentStyleAttr = element.getAttribute('style') || '';
                    const mCont = parentStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
                    if (mCont && mCont[1]) {
                        containerOpacity = Math.min(Math.max(parseFloat(mCont[1]), 0), 1);
                    }
                    let imgOpacity = 1;
                    const imgStyleAttr = imgElement.getAttribute('style') || '';
                    const mImg = imgStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
                    if (mImg && mImg[1]) {
                        imgOpacity = Math.min(Math.max(parseFloat(mImg[1]), 0), 1);
                    }
                    const finalOpacity = containerOpacity * imgOpacity;
                    const transparencyPercentage = Math.min(Math.max(Math.round((1 - finalOpacity) * 100), 0), 100);

                    // Add image as placeholder with potential flip
                    pptSlide.addImage({
                        path: resolvedImagePath,
                        x: position.x,
                        y: position.y,
                        w: position.w,
                        h: position.h,
                        objectName: element.getAttribute('data-name') || `Placeholder Image (${placeholderTypeDetected})`,
                        flipH: isFlipped,
                        flipV: isFlippedVertical,
                        transparency: transparencyPercentage,
                        _placeholderType: phType,
                        _placeholderIdx: phIdx,
                        placeholderType: phType,
                        placeholderIdx: phIdx,
                        srcRect: { l: srcRectL, r: srcRectR, t: srcRectT, b: srcRectB }
                    });
                    return;
                }
            }
            // ✅ NEW: Support base64 data URLs without changing existing logic
            else if (imgElement && imgElement.src && imgElement.src.startsWith('data:')) {
                // Read attributes (case-insensitive in HTML)
                const rawPhType = element.getAttribute('phtype') || element.getAttribute('phType') || '';
                const rawPhIdx = element.getAttribute('phIdx') || element.getAttribute('phidx') || element.getAttribute('data-ph-idx') || '';

                const srcRectL = element.getAttribute('srcRectL') || element.getAttribute('srcrectl') || '';
                const srcRectR = element.getAttribute('srcRectR') || element.getAttribute('srcrectr') || '';

                // Normalize
                const phType = (rawPhType || 'pic').toString().trim().toLowerCase();
                const phIdx = rawPhIdx !== '' ? parseInt(rawPhIdx, 10) : 14;

                // Flip (reuse your existing container-based detection)
                const transformStyle = element.style.transform || '';
                const isFlipped = transformStyle.includes('scaleX(-1)');
                const isFlippedVertical = transformStyle.includes('scaleY(-1)');

                // ✅ NEW: Compute transparency from container + image opacity
                let containerOpacity = 1;
                const parentStyleAttr = element.getAttribute('style') || '';
                const mCont = parentStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
                if (mCont && mCont[1]) {
                    containerOpacity = Math.min(Math.max(parseFloat(mCont[1]), 0), 1);
                }
                let imgOpacity = 1;
                const imgStyleAttr = imgElement.getAttribute('style') || '';
                const mImg = imgStyleAttr.match(/opacity\s*:\s*([0-9.]+)/i);
                if (mImg && mImg[1]) {
                    imgOpacity = Math.min(Math.max(parseFloat(mImg[1]), 0), 1);
                }
                const finalOpacity = containerOpacity * imgOpacity;
                const transparencyPercentage = Math.min(Math.max(Math.round((1 - finalOpacity) * 100), 0), 100);

                // Use data URL directly; pptxgen accepts data URLs or pure base64 payloads
                const dataUrl = imgElement.src;

                pptSlide.addImage({
                    data: dataUrl,
                    x: position.x,
                    y: position.y,
                    w: position.w,
                    h: position.h,
                    objectName: element.getAttribute('data-name') || `Placeholder Image (${placeholderTypeDetected})`,
                    flipH: isFlipped,
                    flipV: isFlippedVertical,
                    transparency: transparencyPercentage,
                    _placeholderType: phType,
                    _placeholderIdx: phIdx,
                    placeholderType: phType,
                    placeholderIdx: phIdx,
                    srcRect: { l: srcRectL, r: srcRectR }
                });
                return;
            }

            // No actual image: create a visible placeholder box + prompt
            pptSlide.addShape(pptx.ShapeType.rect, {
                x: position.x,
                y: position.y,
                w: position.w,
                h: position.h,
                fill: { color: 'F5F5F5' },
                line: { color: 'CCCCCC', width: 1, dashType: 'dash' },
                objectName: element.getAttribute('data-name') || `Image Placeholder (${placeholderTypeDetected})`,
                _placeholderType: 'pic',
                _placeholderIdx: 14,
                placeholderType: 'pic',
                placeholderIdx: 14,
            });

            pptSlide.addText('Click to add image', {
                x: position.x,
                y: position.y + (position.h / 2) - 0.2,
                w: position.w,
                h: 0.4,
                fontSize: 14,
                color: '999999',
                align: 'center'
            });
            return;
        }

        // Handle placeholder text
        if (element.classList.contains('placeholder-text') ||
            element.querySelector('.sli-txt-box.placeholder-text')) {

            const placeholderType = determinePlaceholderType(element, 'text');
            const position = extractElementPosition(element, slideContext);

            const textElement = element.classList.contains('sli-txt-box')
                ? element
                : element.querySelector('.sli-txt-box');

            // ✅ FIXED: Get txtph attributes from the textElement, not the outer element
            const txtPhType = textElement?.getAttribute('txtphtype') || textElement?.getAttribute('txtPhType') || '';
            const txtPhIdx = textElement?.getAttribute('txtphidx') || textElement?.getAttribute('txtPhIdx') || '';
            const txtPhSz = textElement?.getAttribute('txtphsz') || textElement?.getAttribute('txtPhSz') || '';

            let placeholderText = getPlaceholderText(placeholderType);
            let textOptions = extractTextStyle(textElement);

            if (textElement && textElement.textContent.trim()) {
                placeholderText = textElement.textContent.trim();
            }

            // Extract color and luminosity adjustments from the span
            const spanElement = textElement.querySelector('span');
            let textColor = spanElement?.getAttribute('originaltxtcolor') || ''; // Get color
            const lumMod = spanElement?.getAttribute('originallummod');
            const lumOff = spanElement?.getAttribute('originallumoff');

            textOptions.color = textColor;

            // Convert txtPhType for placeholder usage (if needed)
            const normalizedPhType = txtPhType || 'body';
            const normalizedPhIdx = txtPhIdx.trim() !== '' ? parseInt(txtPhIdx.trim(), 10) : undefined;

            // Handle text positioning and alignment
            const justifyContent = element.style.justifyContent || 'flex-start'; // Default to flex-start
            const alignItems = element.style.alignItems || 'flex-start'; // Default to flex-start
            const textAlign = textElement.style.textAlign || 'left'; // Default to left

            // Setting horizontal alignment based on text-align and justify-content
            let align = 'left';
            if (justifyContent === 'flex-start') {
                align = 'left'; // Left align the text
            } else if (justifyContent === 'center') {
                align = 'center'; // Center align the text
            } else if (justifyContent === 'flex-end') {
                align = 'right'; // Right align the text
            }

            // Vertical alignment
            let valign = 'top'; // Default alignment
            if (alignItems === 'flex-start') {
                valign = 'top'; // Top align
            } else if (alignItems === 'center') {
                valign = 'middle'; // Center align
            } else if (alignItems === 'flex-end') {
                valign = 'bottom'; // Bottom align
            }

            // For textAlign (CSS), map it to PowerPoint alignment
            if (textAlign === 'left') {
                align = 'left';
            } else if (textAlign === 'center') {
                align = 'center';
            } else if (textAlign === 'right') {
                align = 'right';
            }

            pptSlide.addText(
                [{ text: placeholderText, options: { bold: false } }],
                {

                    x: position.x,
                    y: position.y,
                    w: position.w,
                    h: position.h,
                    fontSize: textOptions.fontSize,
                    color: textOptions.color,
                    fontFace: textOptions.fontFace || (isTitle ? '+mj-lt' : undefined),
                    align: align,
                    valign: valign,
                    objectName: element.getAttribute('data-name') || `Text Placeholder (${placeholderType})`,
                    _placeholderType: txtPhType,
                    _placeholderIdx: txtPhIdx,
                    _placeholderSz: txtPhSz,
                    placeholderType: txtPhType,
                    placeholderIdx: txtPhIdx,
                    placeholderSz: txtPhSz,
                    _isPlaceholderFormattingOverride: true
                });
            return;
        }

        // Handle shapes with placeholder text
        if (element.classList.contains('shape') &&
            element.querySelector('.placeholder-text')) {

            // First add the shape
            await addShapeToSlide.addShapeToSlide(pptx, pptSlide, element, slideContext);

            // Then add placeholder text on top
            const textElement = element.querySelector('.sli-txt-box.placeholder-text');
            if (textElement) {
                const placeholderType = determinePlaceholderType(textElement, 'text');
                const position = extractElementPosition(element, slideContext);
                const textOptions = extractTextStyle(textElement);

                let placeholderText = getPlaceholderText(placeholderType);
                if (textElement.textContent.trim()) {
                    placeholderText = textElement.textContent.trim();
                }

                // Extract color and luminosity adjustments from the span
                const spanElement = textElement.querySelector('span');
                let textColor = spanElement?.getAttribute('originaltxtcolor') || ''; // Get color
                const lumMod = spanElement?.getAttribute('originallummod');
                const lumOff = spanElement?.getAttribute('originallumoff');

                // Only set the text color if lumMod and lumOff are provided
                // if (lumMod && lumOff) {
                textOptions.color = textColor; // Apply the color if luminosity modifications are provided
                // }\

                pptSlide.addText(placeholderText, {
                    x: position.x,
                    y: position.y,
                    w: position.w,
                    h: position.h,
                    fontSize: textOptions.fontSize,
                    color: textOptions.color,
                    fontFace: textOptions.fontFace,
                    align: textOptions.align,
                    valign: textOptions.valign,
                    objectName: element.getAttribute('data-name') || `Shape Text Placeholder (${placeholderType})`,
                    _placeholderType: 'body',
                    _placeholderIdx: undefined,
                    placeholderType: 'body'
                });
            }
            return;
        }

    } catch (error) {
        console.error(`   ❌ Error processing placeholder element:`, error);
        // Fallback to regular processing for shapes if needed
        if (element.classList.contains('shape')) {
            await addShapeToSlide.addShapeToSlide(pptx, pptSlide, element, slideContext);
        }
    }
}

function determinePlaceholderType(element, defaultType = 'body') {
    // Check for data attributes
    const placeholderTypeAttr = element.getAttribute('data-placeholder-type');
    if (placeholderTypeAttr) {
        return placeholderTypeAttr;
    }

    // Check element position to guess type
    const style = element.style;
    const position = {
        y: parseFloat(style.top) || 0,
        width: parseFloat(style.width) || 0,
        height: parseFloat(style.height) || 0
    };

    if (defaultType === 'text') {
        // Title placeholders are usually at the top and wider
        if (position.y < 100 && position.width > 400) {
            return 'title';
        }
        // Subtitle placeholders are below title
        if (position.y > 50 && position.y < 200 && position.width > 300) {
            return 'subTitle';
        }
        return 'body';
    }

    if (defaultType === 'image') {
        return 'pic';
    }

    return 'body';
}

async function fixChartXmlBugs(slideXmlsDir) {
    try {
        const chartsDir = path.join(slideXmlsDir, 'charts');

        // Check if charts directory exists
        const chartsExist = await checkDirectoryExists(chartsDir);
        if (!chartsExist) {
            return { success: true, chartsFixed: 0 };
        }

        const chartFiles = await fsPromises.readdir(chartsDir);
        const xmlFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

        if (xmlFiles.length === 0) {
            console.log('   ℹ️  No chart XML files found');
            return { success: true, chartsFixed: 0 };
        }

        let fixedCount = 0;

        for (const chartFile of xmlFiles) {
            try {
                const chartPath = path.join(chartsDir, chartFile);
                let chartXml = await fsPromises.readFile(chartPath, 'utf8');
                const originalXml = chartXml;
                let modified = false;

                // Fix 1: Remove extra third axId ONLY from barChart section
                // Extract the barChart section
                const barChartMatch = chartXml.match(/(<c:barChart>)([\s\S]*?)(<\/c:barChart>)/);

                if (barChartMatch) {
                    const barChartContent = barChartMatch[2];
                    const axIdMatches = barChartContent.match(/<c:axId val="\d+"[^>]*\/>/g);

                    if (axIdMatches && axIdMatches.length > 2) {

                        // Keep only first 2 axId elements in barChart section
                        let axIdCount = 0;
                        const fixedBarChartContent = barChartContent.replace(/<c:axId val="\d+"[^>]*\/>/g, (match) => {
                            axIdCount++;
                            return axIdCount <= 2 ? match : '';
                        });

                        chartXml = chartXml.replace(barChartMatch[0], barChartMatch[1] + fixedBarChartContent + barChartMatch[3]);
                        modified = true;
                    }
                }

                // Fix 2: Ensure axis definitions have their axId elements
                // Check catAx
                const catAxMatch = chartXml.match(/(<c:catAx>)([\s\S]*?)(<\/c:catAx>)/);
                if (catAxMatch) {
                    const catAxContent = catAxMatch[2];
                    // Check if axId is missing at the start
                    if (!catAxContent.trim().startsWith('<c:axId')) {
                        // Extract the axId value from barChart references
                        const barChartAxIds = chartXml.match(/<c:barChart>[\s\S]*?<c:axId val="(\d+)"[^>]*\/>/g);
                        if (barChartAxIds && barChartAxIds.length > 0) {
                            // Get the first axId value
                            const firstAxIdMatch = barChartAxIds[0].match(/val="(\d+)"/);
                            if (firstAxIdMatch) {
                                const axIdValue = firstAxIdMatch[1];
                                const fixedCatAxContent = `\n    <c:axId val="${axIdValue}" />` + catAxContent;
                                chartXml = chartXml.replace(catAxMatch[0], catAxMatch[1] + fixedCatAxContent + catAxMatch[3]);
                                modified = true;
                            }
                        }
                    }
                }

                // Check valAx
                const valAxMatch = chartXml.match(/(<c:valAx>)([\s\S]*?)(<\/c:valAx>)/);
                if (valAxMatch) {
                    const valAxContent = valAxMatch[2];
                    // Check if axId is missing at the start
                    if (!valAxContent.trim().startsWith('<c:axId')) {
                        // Extract the second axId value from barChart references
                        const barChartAxIds = chartXml.match(/<c:barChart>[\s\S]*?<c:axId val="(\d+)"[^>]*\/>/g);
                        if (barChartAxIds && barChartAxIds.length > 1) {
                            // Get the second axId value
                            const secondAxIdMatch = barChartAxIds[1].match(/val="(\d+)"/);
                            if (secondAxIdMatch) {
                                const axIdValue = secondAxIdMatch[1];
                                const fixedValAxContent = `\n    <c:axId val="${axIdValue}" />` + valAxContent;
                                chartXml = chartXml.replace(valAxMatch[0], valAxMatch[1] + fixedValAxContent + valAxMatch[3]);
                                modified = true;
                            }
                        }
                    }
                }

                // Fix 3: Replace multiLvlStrRef with strRef for simple categories
                if (chartXml.includes('<c:multiLvlStrRef>')) {

                    chartXml = chartXml.replace(
                        /<c:multiLvlStrRef>([\s\S]*?)<c:multiLvlStrCache>([\s\S]*?)<c:lvl>([\s\S]*?)<\/c:lvl>([\s\S]*?)<\/c:multiLvlStrCache>([\s\S]*?)<\/c:multiLvlStrRef>/g,
                        (match, before, cacheStart, lvlContent, cacheEnd, after) => {
                            return `<c:strRef>${before}<c:strCache>${cacheStart}${lvlContent}${cacheEnd}</c:strCache>${after}</c:strRef>`;
                        }
                    );
                    modified = true;
                }

                // Write fixed XML back if modified
                if (modified && chartXml !== originalXml) {
                    await fsPromises.writeFile(chartPath, chartXml, 'utf8');
                    console.log(`   ✅ Fixed chart XML bugs in ${chartFile}`);
                    fixedCount++;
                } else {
                    console.log(`   ✓ No issues found in ${chartFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error fixing ${chartFile}: ${error.message}`);
            }
        }

        return {
            success: true,
            chartsFixed: fixedCount,
            totalCharts: xmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in fixChartXmlBugs:', error);
        return { success: false, error: error.message };
    }
}

async function cleanChartXmlForSyncfusion(slideXmlsDir) {
    try {
        console.log('🧹 Cleaning chart XML for Syncfusion compatibility...');

        const chartsDir = path.join(slideXmlsDir, 'charts');

        // Check if charts directory exists
        if (!await checkDirectoryExists(chartsDir)) {
            console.log('   ℹ️  No charts directory found, skipping Syncfusion cleanup');
            return { success: true, chartsFixed: 0 };
        }

        const chartFiles = await fsPromises.readdir(chartsDir);
        const chartXmlFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

        if (chartXmlFiles.length === 0) {
            console.log('   ℹ️  No chart XML files found');
            return { success: true, chartsFixed: 0 };
        }

        let fixedCount = 0;

        for (const chartFile of chartXmlFiles) {
            try {
                const chartPath = path.join(chartsDir, chartFile);
                let chartContent = await fsPromises.readFile(chartPath, 'utf8');
                const originalContent = chartContent;

                // Fix 1: Remove whitespace-only content from defRPr tags
                // This specific issue causes Syncfusion PNG export to fail
                chartContent = chartContent.replace(
                    /<a:defRPr([^>]*)>\s+<\/a:defRPr>/g,
                    '<a:defRPr$1/>'
                );

                // Fix 2: Ensure all self-closing tags are properly formatted
                // Clean up other common empty tags
                chartContent = chartContent.replace(
                    /<(a:bodyPr|a:lstStyle|a:effectLst)>\s*<\/\1>/g,
                    '<$1/>'
                );

                // Check if any changes were made
                if (chartContent !== originalContent) {
                    await fsPromises.writeFile(chartPath, chartContent, 'utf8');
                    console.log(`   ✅ Fixed Syncfusion compatibility issues in ${chartFile}`);
                    fixedCount++;
                } else {
                    console.log(`   ✓ No Syncfusion issues found in ${chartFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error cleaning ${chartFile}: ${error.message}`);
            }
        }

        console.log(`   🎉 Syncfusion cleanup complete - ${fixedCount} file(s) fixed`);

        return {
            success: true,
            chartsFixed: fixedCount,
            totalCharts: chartXmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in cleanChartXmlForSyncfusion:', error);
        return { success: false, error: error.message };
    }
}

async function cleanSlideXmlForSyncfusion(slideXmlsDir) {
    try {
        console.log('🧹 Cleaning slide XML for Syncfusion compatibility...');

        const slidesDir = path.join(slideXmlsDir, 'slides');

        // Check if slides directory exists
        if (!await checkDirectoryExists(slidesDir)) {
            console.log('   ℹ️  No slides directory found, skipping slide cleanup');
            return { success: true, slidesFixed: 0 };
        }

        const slideFiles = await fsPromises.readdir(slidesDir);
        const slideXmlFiles = slideFiles.filter(file => file.match(/^slide\d+\.xml$/));

        if (slideXmlFiles.length === 0) {
            console.log('   ℹ️  No slide XML files found');
            return { success: true, slidesFixed: 0 };
        }

        let fixedCount = 0;

        for (const slideFile of slideXmlFiles) {
            try {
                const slidePath = path.join(slidesDir, slideFile);
                let slideContent = await fsPromises.readFile(slidePath, 'utf8');
                const originalContent = slideContent;

                // Fix 1: Convert empty nvPr tags to self-closing
                slideContent = slideContent.replace(
                    /<p:nvPr>\s*<\/p:nvPr>/g,
                    '<p:nvPr/>'
                );

                // Fix 2: Convert empty line (ln) tags to self-closing
                // These are VERY common and cause Syncfusion failures
                slideContent = slideContent.replace(
                    /<a:ln>\s*<\/a:ln>/g,
                    '<a:ln/>'
                );

                // Fix 3: Convert empty avLst (adjust value list) tags to self-closing
                slideContent = slideContent.replace(
                    /<a:avLst>\s*<\/a:avLst>/g,
                    '<a:avLst/>'
                );

                // Fix 4: Convert empty bodyPr tags to self-closing (even with attributes)
                slideContent = slideContent.replace(
                    /<a:bodyPr([^>]*)>\s*<\/a:bodyPr>/g,
                    '<a:bodyPr$1/>'
                );

                // Fix 5: Convert empty lstStyle tags to self-closing
                slideContent = slideContent.replace(
                    /<a:lstStyle>\s*<\/a:lstStyle>/g,
                    '<a:lstStyle/>'
                );

                // Fix 6: Convert other common empty presentation tags to self-closing
                slideContent = slideContent.replace(
                    /<(p:spPr|p:txBody|a:effectLst)>\s*<\/\1>/g,
                    '<$1/>'
                );

                // Fix 7: Remove any invalid control characters
                slideContent = slideContent.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

                // Check if any changes were made
                if (slideContent !== originalContent) {
                    await fsPromises.writeFile(slidePath, slideContent, 'utf8');
                    console.log(`   ✅ Fixed Syncfusion compatibility issues in ${slideFile}`);
                    fixedCount++;
                } else {
                    console.log(`   ✓ No Syncfusion issues found in ${slideFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error cleaning ${slideFile}: ${error.message}`);
            }
        }

        console.log(`   🎉 Slide XML cleanup complete - ${fixedCount} file(s) fixed`);

        return {
            success: true,
            slidesFixed: fixedCount,
            totalSlides: slideXmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in cleanSlideXmlForSyncfusion:', error);
        return { success: false, error: error.message };
    }
}

async function comprehensiveChartXmlFix(slideXmlsDir) {
    try {
        console.log('🔧 Applying comprehensive chart XML fixes for Syncfusion...');

        const chartsDir = path.join(slideXmlsDir, 'charts');

        if (!await checkDirectoryExists(chartsDir)) {
            console.log('   ℹ️  No charts directory found');
            return { success: true, chartsFixed: 0 };
        }

        const chartFiles = await fsPromises.readdir(chartsDir);
        const chartXmlFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

        if (chartXmlFiles.length === 0) {
            console.log('   ℹ️  No chart XML files found');
            return { success: true, chartsFixed: 0 };
        }

        let fixedCount = 0;

        for (const chartFile of chartXmlFiles) {
            try {
                const chartPath = path.join(chartsDir, chartFile);
                let chartContent = await fsPromises.readFile(chartPath, 'utf8');
                const originalContent = chartContent;

                console.log(`   🔍 Processing ${chartFile}...`);

                // FIX 1: Remove whitespace before closing tags
                chartContent = chartContent.replace(/\s+(<\/[ac]:[^>]+>)/g, '$1');

                // FIX 2: bodyPr tags
                chartContent = chartContent.replace(
                    /<a:bodyPr([^>]*)>\s*<\/a:bodyPr>/g,
                    '<a:bodyPr$1/>'
                );

                // FIX 3: lstStyle tags
                chartContent = chartContent.replace(
                    /<a:lstStyle>\s*<\/a:lstStyle>/g,
                    '<a:lstStyle/>'
                );

                // FIX 4: effectLst tags
                chartContent = chartContent.replace(
                    /<a:effectLst>\s*<\/a:effectLst>/g,
                    '<a:effectLst/>'
                );

                // FIX 5: defRPr tags
                chartContent = chartContent.replace(
                    /<a:defRPr([^>]*)>\s*<\/a:defRPr>/g,
                    '<a:defRPr$1/>'
                );

                // FIX 6: ln (line) tags
                chartContent = chartContent.replace(
                    /<a:ln([^>]*)>\s*<\/a:ln>/g,
                    '<a:ln$1/>'
                );

                // FIX 7: Fix spacing in spPr
                chartContent = chartContent.replace(/<c:spPr>\s+/g, '<c:spPr>');
                chartContent = chartContent.replace(/\s+<\/c:spPr>/g, '</c:spPr>');

                // FIX 8: Fix txPr spacing
                chartContent = chartContent.replace(/<c:txPr>\s+/g, '<c:txPr>');
                chartContent = chartContent.replace(/\s+<\/c:txPr>/g, '</c:txPr>');

                // FIX 9: Remove excessive whitespace
                chartContent = chartContent.replace(/>\s{2,}</g, '><');

                // FIX 10: majorGridlines spacing
                chartContent = chartContent.replace(/<c:majorGridlines>\s+/g, '<c:majorGridlines>');
                chartContent = chartContent.replace(/\s+<\/c:majorGridlines>/g, '</c:majorGridlines>');

                // FIX 11: Remove invalid control characters
                chartContent = chartContent.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

                // FIX 12: noFill tags
                chartContent = chartContent.replace(
                    /<a:noFill>\s*<\/a:noFill>/g,
                    '<a:noFill/>'
                );

                // FIX 13: layout tags
                chartContent = chartContent.replace(
                    /<c:layout>\s*<\/c:layout>/g,
                    '<c:layout/>'
                );

                // FIX 14: endParaRPr spacing
                chartContent = chartContent.replace(/\s+<a:endParaRPr/g, '<a:endParaRPr');

                if (chartContent !== originalContent) {
                    await fsPromises.writeFile(chartPath, chartContent, 'utf8');
                    console.log(`   ✅ Fixed ${chartFile}`);
                    fixedCount++;
                } else {
                    console.log(`   ✓ No fixes needed for ${chartFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error processing ${chartFile}:`, error.message);
            }
        }

        console.log(`   🎉 Chart XML fix complete - ${fixedCount} file(s) fixed`);

        return {
            success: true,
            chartsFixed: fixedCount,
            totalCharts: chartXmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in comprehensiveChartXmlFix:', error);
        return { success: false, error: error.message };
    }
}

async function fixChartRelationships(slideXmlsDir) {
    try {
        console.log('🔗 Fixing chart relationships...');

        const chartsRelsDir = path.join(slideXmlsDir, 'charts', '_rels');

        if (!await checkDirectoryExists(chartsRelsDir)) {
            console.log('   ℹ️  No chart relationships directory');
            return { success: true };
        }

        const relsFiles = await fsPromises.readdir(chartsRelsDir);
        const chartRelsFiles = relsFiles.filter(file => file.match(/^chart\d+\.xml\.rels$/));

        for (const relsFile of chartRelsFiles) {
            try {
                const relsPath = path.join(chartsRelsDir, relsFile);
                let relsContent = await fsPromises.readFile(relsPath, 'utf8');
                const originalContent = relsContent;

                relsContent = relsContent.replace(
                    /<Relationship[^>]*Type="[^"]*\/chartStyle"[^>]*\/>/g,
                    ''
                );
                relsContent = relsContent.replace(
                    /<Relationship[^>]*Type="[^"]*\/chartColorStyle"[^>]*\/>/g,
                    ''
                );
                relsContent = relsContent.replace(/\s{2,}/g, '');

                if (relsContent !== originalContent) {
                    await fsPromises.writeFile(relsPath, relsContent, 'utf8');
                    console.log(`   ✅ Fixed relationships in ${relsFile}`);
                }

            } catch (error) {
                console.error(`   ❌ Error fixing ${relsFile}:`, error.message);
            }
        }

        console.log('   ✅ Chart relationships check complete');
        return { success: true };

    } catch (error) {
        console.error('   ❌ Error in fixChartRelationships:', error);
        return { success: false, error: error.message };
    }
}

async function validateEmbeddedExcel(slideXmlsDir) {
    try {
        console.log('📊 Validating embedded Excel files...');

        const embeddingsDir = path.join(slideXmlsDir, 'embeddings');

        if (!await checkDirectoryExists(embeddingsDir)) {
            console.log('   ℹ️  No embeddings directory');
            return { success: true };
        }

        const embeddingFiles = await fsPromises.readdir(embeddingsDir);
        const excelFiles = embeddingFiles.filter(file =>
            file.match(/\.xlsx$/i) || file.match(/Microsoft_Excel/i)
        );

        for (const excelFile of excelFiles) {
            try {
                const excelPath = path.join(embeddingsDir, excelFile);
                const stats = await fsPromises.stat(excelPath);

                console.log(`   📄 ${excelFile}: ${stats.size} bytes`);

                if (stats.size === 0) {
                    console.error(`   ❌ Excel file is empty!`);
                    return { success: false, error: 'Empty Excel file' };
                }

                if (stats.size < 100) {
                    console.warn(`   ⚠️  Excel file seems too small`);
                }

                console.log(`   ✅ Excel file validated`);

            } catch (error) {
                console.error(`   ❌ Error validating ${excelFile}:`, error.message);
                return { success: false, error: error.message };
            }
        }

        return { success: true };

    } catch (error) {
        console.error('   ❌ Error in validateEmbeddedExcel:', error);
        return { success: false, error: error.message };
    }
}

function extractElementPosition(element, slideContext) {
    const style = element.style;

    const left = parseFloat(style.left) || 0;
    const top = parseFloat(style.top) || 0;
    const width = parseFloat(style.width) || 100;
    const height = parseFloat(style.height) || 50;

    // Convert pixels to inches (72 DPI)
    return {
        x: left / 72,
        y: top / 72,
        w: width / 72,
        h: height / 72
    };
}

function extractTextStyle(textElement) {
    if (!textElement) {
        return {
            fontSize: 18,
            color: '000000',
            fontFace: undefined, // let theme handle fallback
            align: 'left',
            valign: 'top'
        };
    }

    const style = textElement.style || {};
    const spanFF = dominantSpanFontFamily(textElement);          // prefer inner .span-txt
    const boxFF = firstFontFamily(style.fontFamily);
    const fontFace = spanFF || boxFF || undefined;

    return {
        fontSize: parseFloat(style.fontSize) || 18,
        color: extractColor(style.color) || '000000',
        fontFace,
        align: getTextAlign(style.textAlign),
        valign: getVerticalAlign(style.verticalAlign || style.alignItems)
    };
}

function getPlaceholderText(placeholderType) {
    const placeholderTexts = {
        'title': 'Click to add title',
        'subTitle': 'Click to add subtitle',
        'body': 'Click to add text',
        'content': 'Click to add content'
    };
    return placeholderTexts[placeholderType] || 'Click to add text';
}

function getTextAlign(textAlign) {
    switch (textAlign) {
        case 'center': return 'center';
        case 'right': return 'right';
        case 'justify': return 'justify';
        default: return 'left';
    }
}

function getVerticalAlign(valign) {
    switch (valign) {
        case 'top':
        case 'flex-start': return 'top';
        case 'middle':
        case 'center': return 'middle';
        case 'bottom':
        case 'flex-end': return 'bottom';
        default: return 'top';
    }
}

function resolveImagePath(imageSrc) {
    try {
        if (imageSrc.startsWith('data:') || imageSrc.startsWith('http://') || imageSrc.startsWith('https://')) {
            return imageSrc;
        }

        let searchPaths = [];
        let filename = '';

        if (imageSrc.startsWith('/')) {
            filename = path.basename(imageSrc);
            const relativePath = imageSrc.substring(1);

            for (const basePath of [IMAGE_CONFIG.baseImagePath, ...IMAGE_CONFIG.alternativeBasePaths]) {
                for (const subdir of IMAGE_CONFIG.imageSubdirs) {
                    searchPaths.push(path.join(basePath, subdir, relativePath));
                    searchPaths.push(path.join(basePath, subdir, filename));
                }
            }
        } else {
            filename = path.basename(imageSrc);
            for (const basePath of [process.cwd(), IMAGE_CONFIG.baseImagePath, ...IMAGE_CONFIG.alternativeBasePaths]) {
                for (const subdir of IMAGE_CONFIG.imageSubdirs) {
                    searchPaths.push(path.join(basePath, subdir, imageSrc));
                    searchPaths.push(path.join(basePath, subdir, filename));
                }
            }
        }

        searchPaths = [...new Set(searchPaths)];

        for (const searchPath of searchPaths) {
            if (fs.existsSync(searchPath)) {
                return searchPath;
            }
        }

        if (IMAGE_CONFIG.enableRecursiveSearch && filename) {
            for (const basePath of [IMAGE_CONFIG.baseImagePath, ...IMAGE_CONFIG.alternativeBasePaths]) {
                const found = findFileRecursively(basePath, filename, IMAGE_CONFIG.maxSearchDepth);
                if (found) return found;
            }
        }

        return null;
    } catch (error) {
        console.error(`   ❌ Error resolving image path ${imageSrc}:`, error);
        return null;
    }
}

const IMAGE_CONFIG = {
    baseImagePath: '/var/www/html/deckbuilder/Template-Editor',
    alternativeBasePaths: [
        '/var/www/html/deckbuilder',
        '/var/www/html',
        process.cwd(),
    ],
    imageSubdirs: ['', 'public', 'uploads', 'assets', 'static', 'images', 'media'],
    enableRecursiveSearch: true,
    maxSearchDepth: 3
};

function findFileRecursively(dir, filename, maxDepth = 3, currentDepth = 0) {
    if (currentDepth > maxDepth) return null;

    try {
        if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) {
            return null;
        }

        const files = fs.readdirSync(dir);
        const directPath = path.join(dir, filename);
        if (fs.existsSync(directPath)) {
            return directPath;
        }

        for (const file of files) {
            const fullPath = path.join(dir, file);
            try {
                if (fs.statSync(fullPath).isDirectory()) {
                    const found = findFileRecursively(fullPath, filename, maxDepth, currentDepth + 1);
                    if (found) return found;
                }
            } catch (e) {
                continue;
            }
        }
    } catch (error) {
        return null;
    }
    return null;
}

function firstFontFamily(ff) {
    if (!ff) return undefined;
    return String(ff).split(',')[0].trim().replace(/^['"]|['"]$/g, '');
}

function dominantSpanFontFamily(textEl) {
    if (!textEl) return undefined;
    const spans = textEl.querySelectorAll('.span-txt');
    if (!spans.length) return undefined;

    const weights = new Map();
    for (const sp of spans) {
        const ff = firstFontFamily(sp.style.fontFamily);
        if (!ff) continue;
        const w = (sp.textContent || '').length || 1;
        weights.set(ff, (weights.get(ff) || 0) + w);
    }
    let best, max = 0;
    for (const [ff, w] of weights) if (w > max) { best = ff; max = w; }
    return best;
}

async function processLayoutElement(pptx, pptSlide, element, slideContext, processedElements) {
    // Filter out master elements from nested shapes and images
    const nestedShapes = element.querySelectorAll(".shape");
    const nestedImages = element.querySelectorAll(".image-container");

    const filteredShapes = Array.from(nestedShapes).filter(shape => !isMasterElement(shape));
    const filteredImages = Array.from(nestedImages).filter(img => !isMasterElement(img));

    if (filteredShapes.length > 0 || filteredImages.length > 0) {
        for (const shapeElement of filteredShapes) {
            if (!processedElements.has(shapeElement)) {
                try {
                    await processShapeElement(pptx, pptSlide, shapeElement, slideContext);
                    processedElements.add(shapeElement);
                } catch (nestedShapeError) {
                    console.error(`   ❌ Error processing nested shape:`, nestedShapeError);
                }
            }
        }

        for (const imgContainer of filteredImages) {
            if (!processedElements.has(imgContainer)) {
                try {
                    const imgElement = imgContainer.querySelector("img");
                    if (imgElement && !isMasterElement(imgElement)) {
                        await addImage.addImageToSlide(pptx, pptSlide, imgElement, slideContext);
                        processedElements.add(imgContainer);
                    }
                } catch (nestedImageError) {
                    console.error(`   ❌ Error processing nested image:`, nestedImageError);
                }
            }
        }
        processedElements.add(element);
    } else {
        processedElements.add(element);
    }
}

async function processShapeElement(pptx, pptSlide, shapeElement, slideContext = null) {
    // Filter out master elements
    if (isMasterElement(shapeElement)) {
        return;
    }

    // Check if this is a placeholder element
    if (isPlaceholderElement(shapeElement)) {
        await processPlaceholderElement(pptx, pptSlide, shapeElement, slideContext);
        return;
    }

    const elementStyle = shapeElement?.getAttribute('style');
    if (elementStyle && elementStyle.includes('visibility: hidden')) {
        return;
    }

    // Special check for custom geometry shapes with master classes
    if (shapeElement.id === 'custGeom' || shapeElement.classList.contains('custom-shape')) {
        if (shapeElement.classList.contains('custom-shape-master') ||
            shapeElement.classList.contains('master')) {
            return;
        }
    }

    addShapeToSlide.addShapeToSlide(pptx, pptSlide, shapeElement, slideContext);

    // Handle text boxes
    const txBox = shapeElement.querySelector(".sli-txt-box");
    if (txBox && !isMasterElement(txBox) && !isMasterElement(txBox.parentElement)) {
        if (txBox.classList.contains('placeholder-text')) {
            const placeholderType = determinePlaceholderType(txBox, 'text');
            const position = extractElementPosition(shapeElement, slideContext);
            const textOptions = extractTextStyle(txBox);

            let placeholderText = getPlaceholderText(placeholderType);
            if (txBox.textContent.trim()) {
                placeholderText = txBox.textContent.trim();
            }

            pptSlide.addText(placeholderText, {
                x: position.x,
                y: position.y,
                w: position.w,
                h: position.h,
                fontSize: textOptions.fontSize,
                color: textOptions.color,
                fontFace: textOptions.fontFace,
                align: textOptions.align,
                valign: textOptions.valign,
                objectName: `Text Box Placeholder (${placeholderType})`
            });
        } else {
            addTextBox.addTextBoxToSlide(pptSlide, txBox, shapeElement, slideContext);
        }
    }
}

async function fixPptxFile(filePath) {
    try {
        const pptxBuffer = fs.readFileSync(filePath);
        const zip = new JSZip();
        const pptxZip = await zip.loadAsync(pptxBuffer);

        if (pptxZip.files['ppt/presentation.xml']) {
            let presentationXml = await pptxZip.files['ppt/presentation.xml'].async('text');

            presentationXml = presentationXml.replace(
                /<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/gi,
                ''
            );

            presentationXml = presentationXml.replace(
                /<p:notesMasterId[^>]*\/>/gi,
                ''
            );

            pptxZip.file('ppt/presentation.xml', presentationXml);
        }

        if (pptxZip.files['ppt/_rels/presentation.xml.rels']) {
            let relsXml = await pptxZip.files['ppt/_rels/presentation.xml.rels'].async('text');

            relsXml = relsXml.replace(
                /<Relationship[^>]*Target="[^"]*notesMaster[^"]*"[^>]*\/>/gi,
                ''
            );

            pptxZip.file('ppt/_rels/presentation.xml.rels', relsXml);
        }

        const modifiedPptxBuffer = await pptxZip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });

        fs.writeFileSync(filePath, modifiedPptxBuffer);
        return true;
    } catch (error) {
        console.error('❌ Failed to fix PPTX file:', error);
        return false;
    }
}

async function extractAndProcessSlideXMLs(pptxFilePath, customOutputDir = null, slideX) {
    try {
        // Read the PPTX file
        const zipData = await fsPromises.readFile(pptxFilePath);
        const zip = await JSZip.loadAsync(zipData);


        // Extract the original slide number from slideX (e.g., "5" from "ppt/slides/slide5.xml")
        let originalSlideNumber = null;
        if (slideX) {
            const slideMatch = slideX.match(/slide(\d+)\.xml/);
            if (slideMatch) {
                originalSlideNumber = slideMatch[1];
            }
        }

        console.log("originalSlideNumber ------- >> ", originalSlideNumber);

        // Get www-data user and group IDs
        let wwwDataUid, wwwDataGid;
        try {
            const { execSync } = require('child_process');
            wwwDataUid = parseInt(execSync('id -u www-data', { encoding: 'utf8' }).trim());
            wwwDataGid = parseInt(execSync('id -g www-data', { encoding: 'utf8' }).trim());
        } catch (error) {
            wwwDataUid = 33;
            wwwDataGid = 33;
        }

        // FIXED: Use custom directory path if provided, otherwise use default
        const outputDir = customOutputDir || path.join(path.dirname(pptxFilePath), 'slide_xmls');

        // Create main output directory with www-data ownership
        await fsPromises.mkdir(outputDir, { recursive: true, mode: 0o755 });
        try {
            await fsPromises.chown(outputDir, wwwDataUid, wwwDataGid);
        } catch (chownError) {
            console.warn(`⚠️ Could not set ownership for output directory: ${chownError.message}`);
        }

        // Create subdirectories for organization with www-data ownership
        const slidesDir = path.join(outputDir, 'slides');
        const relsDir = path.join(outputDir, '_rels');
        const mediaDir = path.join(outputDir, 'media');
        const chartDir = path.join(outputDir, 'charts');
        const chartRelsDir = path.join(outputDir, 'charts', '_rels');
        const embeddingsDir = path.join(outputDir, 'embeddings');

        // Create slides directory
        await fsPromises.mkdir(slidesDir, { recursive: true, mode: 0o755 });
        try {
            await fsPromises.chown(slidesDir, wwwDataUid, wwwDataGid);
        } catch (chownError) {
            console.warn(`⚠️ Could not set ownership for slides directory: ${chownError.message}`);
        }

        // Create rels directory
        await fsPromises.mkdir(relsDir, { recursive: true, mode: 0o755 });
        try {
            await fsPromises.chown(relsDir, wwwDataUid, wwwDataGid);
        } catch (chownError) {
            console.warn(`⚠️ Could not set ownership for _rels directory: ${chownError.message}`);
        }

        // Create media directory
        await fsPromises.mkdir(mediaDir, { recursive: true, mode: 0o755 });
        try {
            await fsPromises.chown(mediaDir, wwwDataUid, wwwDataGid);
        } catch (chownError) {
            console.warn(`⚠️ Could not set ownership for media directory: ${chownError.message}`);
        }

        // Get slideX.xml files
        const slideXMLs = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/slides\/slide\d+\.xml$/)
        );

        // Get slideX.xml.rels files
        const slideRelsXMLs = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/)
        );

        // Get media files
        const mediaFiles = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/media\/.+$/)
        );

        // ADDED: Get chart files
        const chartXMLs = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/charts\/chart\d+\.xml$/)
        );

        // ADDED: Get chart relationship files
        const chartRelsXMLs = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/charts\/_rels\/chart\d+\.xml\.rels$/)
        );

        // ADDED: Get embedded Excel files (chart data)
        const embeddingFiles = Object.keys(zip.files).filter(fileName =>
            fileName.match(/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/)
        );

        // Check if we have any charts to extract
        const hasCharts = chartXMLs.length > 0 || chartRelsXMLs.length > 0 || embeddingFiles.length > 0;

        if (slideXMLs.length === 0 && slideRelsXMLs.length === 0 && mediaFiles.length === 0 && !hasCharts) {
            return;
        }

        // Process slideX.xml files
        for (const fileName of slideXMLs) {
            const xmlContent = await zip.file(fileName).async('string');

            // Extract the current slide number from the converted PPTX (e.g., "1" from "slide1.xml")
            const currentSlideMatch = fileName.match(/slide(\d+)\.xml/);
            let slideName = path.basename(fileName);

            // If we have an original slide number and this is the first slide in the converted PPTX,
            // rename it to match the original slide number
            if (originalSlideNumber && currentSlideMatch) {
                const currentSlideNum = currentSlideMatch[1];
                // Map the converted slide to the original slide number
                slideName = `slide${originalSlideNumber}.xml`;
            }

            const outputPath = path.join(slidesDir, slideName);

            // Write file with appropriate permissions
            await fsPromises.writeFile(outputPath, xmlContent, { mode: 0o644 });

            // Set file ownership to www-data
            try {
                await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
            } catch (chownError) {
                console.warn(`   ⚠️ Could not set ownership for ${slideName}: ${chownError.message}`);
            }
        }

        // Process slideX.xml.rels files
        for (const fileName of slideRelsXMLs) {
            const xmlContent = await zip.file(fileName).async('string');

            // Extract the current slide number from the converted PPTX (e.g., "1" from "slide1.xml.rels")
            const currentSlideMatch = fileName.match(/slide(\d+)\.xml\.rels/);
            let relsName = path.basename(fileName);

            // If we have an original slide number and this is the first slide in the converted PPTX,
            // rename it to match the original slide number
            if (originalSlideNumber && currentSlideMatch) {
                const currentSlideNum = currentSlideMatch[1];
                // Map the converted slide rels to the original slide number
                relsName = `slide${originalSlideNumber}.xml.rels`;
            }

            const outputPath = path.join(relsDir, relsName);

            // Write file with appropriate permissions
            await fsPromises.writeFile(outputPath, xmlContent, { mode: 0o644 });

            // Set file ownership to www-data
            try {
                await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
            } catch (chownError) {
                console.warn(`   ⚠️ Could not set ownership for ${relsName}: ${chownError.message}`);
            }
        }

        // Process media files
        for (const fileName of mediaFiles) {
            const mediaContent = await zip.file(fileName).async('nodebuffer'); // Use nodebuffer for binary files
            const mediaName = path.basename(fileName);
            const outputPath = path.join(mediaDir, mediaName);

            // Write file with appropriate permissions
            await fsPromises.writeFile(outputPath, mediaContent, { mode: 0o644 });

            // Set file ownership to www-data
            try {
                await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
            } catch (chownError) {
                console.warn(`   ⚠️ Could not set ownership for ${mediaName}: ${chownError.message}`);
            }
        }

        // ADDED: Process chart files only if charts exist
        let chartsExtracted = 0;
        let chartRelsExtracted = 0;
        let embeddingsExtracted = 0;

        if (hasCharts) {

            // Create chart directories only if we have charts
            if (chartXMLs.length > 0 || chartRelsXMLs.length > 0) {
                await fsPromises.mkdir(chartDir, { recursive: true, mode: 0o755 });
                try {
                    await fsPromises.chown(chartDir, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`⚠️ Could not set ownership for charts directory: ${chownError.message}`);
                }
            }

            if (chartRelsXMLs.length > 0) {
                await fsPromises.mkdir(chartRelsDir, { recursive: true, mode: 0o755 });
                try {
                    await fsPromises.chown(chartRelsDir, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`⚠️ Could not set ownership for chart _rels directory: ${chownError.message}`);
                }
            }

            if (embeddingFiles.length > 0) {
                await fsPromises.mkdir(embeddingsDir, { recursive: true, mode: 0o755 });
                try {
                    await fsPromises.chown(embeddingsDir, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`⚠️ Could not set ownership for embeddings directory: ${chownError.message}`);
                }
            }

            // Process chart XML files
            for (const fileName of chartXMLs) {
                const xmlContent = await zip.file(fileName).async('string');
                const chartName = path.basename(fileName);
                const outputPath = path.join(chartDir, chartName);

                // Write file with appropriate permissions
                await fsPromises.writeFile(outputPath, xmlContent, { mode: 0o644 });

                // Set file ownership to www-data
                try {
                    await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`   ⚠️ Could not set ownership for ${chartName}: ${chownError.message}`);
                }

                chartsExtracted++;
            }

            // Process chart relationship files
            for (const fileName of chartRelsXMLs) {
                const xmlContent = await zip.file(fileName).async('string');
                const chartRelsName = path.basename(fileName);
                const outputPath = path.join(chartRelsDir, chartRelsName);

                // Write file with appropriate permissions
                await fsPromises.writeFile(outputPath, xmlContent, { mode: 0o644 });

                // Set file ownership to www-data
                try {
                    await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`   ⚠️ Could not set ownership for ${chartRelsName}: ${chownError.message}`);
                }

                chartRelsExtracted++;
            }

            // Process embedded Excel files (chart data)
            for (const fileName of embeddingFiles) {
                const embeddingContent = await zip.file(fileName).async('nodebuffer'); // Use nodebuffer for binary files
                const embeddingName = path.basename(fileName);
                const outputPath = path.join(embeddingsDir, embeddingName);

                // Write file with appropriate permissions
                await fsPromises.writeFile(outputPath, embeddingContent, { mode: 0o644 });

                // Set file ownership to www-data
                try {
                    await fsPromises.chown(outputPath, wwwDataUid, wwwDataGid);
                } catch (chownError) {
                    console.warn(`   ⚠️ Could not set ownership for ${embeddingName}: ${chownError.message}`);
                }

                embeddingsExtracted++;
            }
        }

        // Return the output directory for verification
        return {
            success: true,
            outputDir: outputDir,
            slidesExtracted: slideXMLs.length,
            relsExtracted: slideRelsXMLs.length,
            mediaExtracted: mediaFiles.length,
            chartsExtracted: chartsExtracted,
            chartRelsExtracted: chartRelsExtracted,
            embeddingsExtracted: embeddingsExtracted,
            hasCharts: hasCharts,
            ownershipSet: true,
            originalSlideNumber: originalSlideNumber, // Include the mapped slide number
            slideMappingApplied: !!originalSlideNumber // Boolean flag indicating if mapping was applied
        };

    } catch (error) {
        console.error('   ❌ Error extracting files:', error.message);
        throw error;
    }
}

async function replaceSlideXMLInPPTX(fileFolderName, extractedSlidesDir) {
    try {

        // Validate inputs
        if (!fileFolderName || !extractedSlidesDir) {
            throw new Error('Invalid folder name or extracted slides directory path.');
        }

        const filesDir = path.resolve(__dirname, '../files');
        const targetDir = path.join(filesDir, fileFolderName);

        // Check if target directory exists
        try {
            await fsPromises.access(targetDir);
        } catch {
            throw new Error(`Target directory not found: ${targetDir}`);
        }

        // Paths for extracted files
        const extractedSlidesPath = path.join(extractedSlidesDir, 'slides');
        const extractedRelsPath = path.join(extractedSlidesDir, '_rels');
        const extractedChartsPath = path.join(extractedSlidesDir, 'charts');
        const extractedChartRelsPath = path.join(extractedSlidesDir, 'charts', '_rels');
        const extractedEmbeddingsPath = path.join(extractedSlidesDir, 'embeddings');

        // Check if extracted directories exist
        const slidesExist = await checkDirectoryExists(extractedSlidesPath);
        const relsExist = await checkDirectoryExists(extractedRelsPath);
        const chartsExist = await checkDirectoryExists(extractedChartsPath);
        const chartRelsExist = await checkDirectoryExists(extractedChartRelsPath);
        const embeddingsExist = await checkDirectoryExists(extractedEmbeddingsPath);

        if (!slidesExist && !relsExist && !chartsExist && !chartRelsExist && !embeddingsExist) {
            throw new Error('No extracted files found to replace');
        }

        let replacedCount = 0;

        // Replace slideX.xml files
        if (slidesExist) {

            const slideFiles = await fsPromises.readdir(extractedSlidesPath);

            const xmlFiles = slideFiles.filter(file => file.match(/^slide\d+\.xml$/));

            console.log(`📄 Found ${xmlFiles.length} slide XML files to replace`);

            for (const xmlFile of xmlFiles) {
                try {
                    // Read converted XML content
                    const convertedXMLPath = path.join(extractedSlidesPath, xmlFile);
                    const convertedXMLContent = await fsPromises.readFile(convertedXMLPath, 'utf8');

                    // Target path in PPTX structure
                    const targetXMLPath = path.join(targetDir, 'ppt', 'slides', xmlFile);

                    // Check if target file exists
                    try {
                        await fsPromises.access(targetXMLPath);
                        await fsPromises.writeFile(targetXMLPath, convertedXMLContent, 'utf8');
                        replacedCount++;
                    } catch {
                        console.log(`   ⚠️ Target file not found, skipping: ${xmlFile}`);
                    }
                } catch (error) {
                    console.error(`   ❌ Error replacing ${xmlFile}: ${error.message}`);
                }
            }
        }

        // Replace slideX.xml.rels files
        if (relsExist) {
            const relsFiles = await fsPromises.readdir(extractedRelsPath);
            const relsXmlFiles = relsFiles.filter(file => file.match(/^slide\d+\.xml\.rels$/));

            for (const relsFile of relsXmlFiles) {
                try {
                    // Read converted rels content
                    const convertedRelsPath = path.join(extractedRelsPath, relsFile);
                    const convertedRelsContent = await fsPromises.readFile(convertedRelsPath, 'utf8');

                    // Target path in PPTX structure
                    const targetRelsPath = path.join(targetDir, 'ppt', 'slides', '_rels', relsFile);

                    // Ensure _rels directory exists
                    const relsDir = path.dirname(targetRelsPath);
                    await fsPromises.mkdir(relsDir, { recursive: true });

                    // Read original rels file to preserve slideLayout relationship
                    let finalRelsContent = convertedRelsContent;

                    try {
                        await fsPromises.access(targetRelsPath);
                        const originalRelsContent = await fsPromises.readFile(targetRelsPath, 'utf8');

                        // Merge rels files - preserve slideLayout from original
                        finalRelsContent = await mergeRelsFiles(originalRelsContent, convertedRelsContent, relsFile);

                    } catch {
                        console.log(`   ℹ️ Creating new rels file: ${relsFile}`);
                    }

                    await fsPromises.writeFile(targetRelsPath, finalRelsContent, 'utf8');
                    replacedCount++;
                } catch (error) {
                    console.error(`   ❌ Error replacing ${relsFile}: ${error.message}`);
                }
            }
        }

        // Replace chartX.xml files
        if (chartsExist) {
            const chartFiles = await fsPromises.readdir(extractedChartsPath);
            const chartXmlFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

            for (const chartFile of chartXmlFiles) {
                try {
                    // Read converted chart XML content
                    const convertedChartPath = path.join(extractedChartsPath, chartFile);
                    const convertedChartContent = await fsPromises.readFile(convertedChartPath, 'utf8');

                    // Target path in PPTX structure
                    const targetChartPath = path.join(targetDir, 'ppt', 'charts', chartFile);

                    // Ensure charts directory exists
                    const chartsDir = path.dirname(targetChartPath);
                    await fsPromises.mkdir(chartsDir, { recursive: true });

                    // Check if target file exists
                    try {
                        await fsPromises.access(targetChartPath);
                        await fsPromises.writeFile(targetChartPath, convertedChartContent, 'utf8');

                        replacedCount++;
                    } catch {
                        console.log(`   ℹ️ Creating new chart file: ${chartFile}`);
                        await fsPromises.writeFile(targetChartPath, convertedChartContent, 'utf8');
                        replacedCount++;
                    }
                } catch (error) {
                    console.error(`   ❌ Error replacing ${chartFile}: ${error.message}`);
                }
            }
        }

        // Replace chartX.xml.rels files
        if (chartRelsExist) {
            const chartRelsFiles = await fsPromises.readdir(extractedChartRelsPath);
            const chartRelsXmlFiles = chartRelsFiles.filter(file => file.match(/^chart\d+\.xml\.rels$/));

            for (const chartRelsFile of chartRelsXmlFiles) {
                try {
                    // Read converted chart rels content
                    const convertedChartRelsPath = path.join(extractedChartRelsPath, chartRelsFile);
                    const convertedChartRelsContent = await fsPromises.readFile(convertedChartRelsPath, 'utf8');

                    // Target path in PPTX structure
                    const targetChartRelsPath = path.join(targetDir, 'ppt', 'charts', '_rels', chartRelsFile);

                    // Ensure chart _rels directory exists
                    const chartRelsDir = path.dirname(targetChartRelsPath);
                    await fsPromises.mkdir(chartRelsDir, { recursive: true });

                    // Check if target file exists and merge if needed
                    let finalChartRelsContent = convertedChartRelsContent;

                    try {
                        await fsPromises.access(targetChartRelsPath);
                        const originalChartRelsContent = await fsPromises.readFile(targetChartRelsPath, 'utf8');

                        // You might want to merge chart rels files similar to slide rels
                        // For now, we'll replace completely, but you can add merge logic if needed
                        finalChartRelsContent = convertedChartRelsContent;
                    } catch {
                        console.log(`   ℹ️ Creating new chart rels file: ${chartRelsFile}`);
                    }

                    // 🆕 Remove style1.xml and colors1.xml relationships from chart rels
                    finalChartRelsContent = finalChartRelsContent.replace(
                        /<Relationship[^>]*Target="style1\.xml"[^>]*\/>\s*/g,
                        ''
                    );
                    finalChartRelsContent = finalChartRelsContent.replace(
                        /<Relationship[^>]*Target="colors1\.xml"[^>]*\/>\s*/g,
                        ''
                    );
                    console.log(`   🧹 Cleaned chart rels: removed style1.xml and colors1.xml relationships`);

                    await fsPromises.writeFile(targetChartRelsPath, finalChartRelsContent, 'utf8');

                    replacedCount++;
                } catch (error) {
                    console.error(`   ❌ Error replacing ${chartRelsFile}: ${error.message}`);
                }
            }
        }

        // Replace embedded Excel files
        if (embeddingsExist) {
            // Source (extracted) files
            const embeddingFiles = await fsPromises.readdir(extractedEmbeddingsPath);
            const excelFiles = embeddingFiles.filter(file =>
                /^Microsoft_Excel_Worksheet\d+\.xlsx$/i.test(file)
            );

            // Target embeddings dir inside the PPTX folder
            const targetEmbeddingsDir = path.join(targetDir, 'ppt', 'embeddings');
            await fsPromises.mkdir(targetEmbeddingsDir, { recursive: true });

            // 1) Remove any existing Excel files in ppt/embeddings
            try {
                const existingTargetEmbeds = await fsPromises.readdir(targetEmbeddingsDir);
                const existingExcel = existingTargetEmbeds.filter(name => /\.xlsx$/i.test(name));
                if (existingExcel.length) {

                    for (const oldXlsx of existingExcel) {
                        const oldPath = path.join(targetEmbeddingsDir, oldXlsx);
                        try {
                            await fsPromises.unlink(oldPath);
                        } catch (delErr) {
                            console.warn(`   ⚠️ Could not remove ${oldXlsx}: ${delErr.message}`);
                        }
                    }
                } else {
                    console.log("   ℹ️ No existing Excel embeddings to remove");
                }
            } catch (scanErr) {
                console.warn(`   ⚠️ Could not scan target embeddings dir: ${scanErr.message}`);
            }

            // 2) Create (write) fresh Excel embeddings
            for (const excelFile of excelFiles) {
                try {
                    const srcPath = path.join(extractedEmbeddingsPath, excelFile);
                    const buf = await fsPromises.readFile(srcPath);

                    const targetExcelPath = path.join(targetEmbeddingsDir, excelFile);
                    await fsPromises.writeFile(targetExcelPath, buf);

                    replacedCount++;
                } catch (error) {
                    console.error(`   ❌ Error writing ${excelFile}: ${error.message}`);
                }
            }
        }


        console.log(`   🎉 Successfully replaced ${replacedCount} files total`);

        return {
            success: true,
            message: `Successfully replaced ${replacedCount} files (slides, rels, charts, chart rels, and embeddings)`,
            replacedCount,
            details: {
                slidesProcessed: slidesExist,
                relsProcessed: relsExist,
                chartsProcessed: chartsExist,
                chartRelsProcessed: chartRelsExist,
                embeddingsProcessed: embeddingsExist
            }
        };

    } catch (error) {
        console.error(`   ❌ Error in replaceSlideXMLInPPTX: ${error.message}`);
        throw error;
    }
}

async function replaceSlideImages(fileFolderName, extractedSlidesDir) {
    try {
        // Validate inputs
        if (!fileFolderName || !extractedSlidesDir) {
            throw new Error('Invalid folder name or extracted slides directory path.');
        }

        const filesDir = path.resolve(__dirname, '../files');
        const targetDir = path.join(filesDir, fileFolderName);

        // Check if target directory exists
        try {
            await fsPromises.access(targetDir);
        } catch {
            throw new Error(`Target directory not found: ${targetDir}`);
        }

        // Path for extracted media files
        const extractedMediaPath = path.join(extractedSlidesDir, 'media');

        // Check if extracted media directory exists
        const mediaExist = await checkDirectoryExists(extractedMediaPath);

        if (!mediaExist) {
            throw new Error('No extracted media files found to replace');
        }

        let replacedOrAddedCount = 0;

        // Replace or add media files
        if (mediaExist) {
            const mediaFiles = await fsPromises.readdir(extractedMediaPath);
            const validMediaFiles = mediaFiles.filter(file =>
                file.match(/\.(png|jpg|jpeg|gif|bmp|mp4|wmv|avi)$/i)
            );

            for (const mediaFile of validMediaFiles) {
                try {
                    // Read converted media content
                    const convertedMediaPath = path.join(extractedMediaPath, mediaFile);
                    const convertedMediaContent = await fsPromises.readFile(convertedMediaPath);

                    // Target path in PPTX structure
                    const targetMediaPath = path.join(targetDir, 'ppt', 'media', mediaFile);

                    // Ensure media directory exists
                    const mediaDir = path.dirname(targetMediaPath);
                    await fsPromises.mkdir(mediaDir, { recursive: true });

                    // Check if target file exists
                    try {
                        await fsPromises.access(targetMediaPath);
                        await fsPromises.writeFile(targetMediaPath, convertedMediaContent);
                    } catch {
                        // If file doesn't exist, add it
                        await fsPromises.writeFile(targetMediaPath, convertedMediaContent);
                        console.log(`   ✅ Added: ${mediaFile}`);
                    }
                    replacedOrAddedCount++;
                } catch (error) {
                    console.error(`   ❌ Error processing ${mediaFile}: ${error.message}`);
                }
            }
        }

        console.log(`   🎉 Successfully processed ${replacedOrAddedCount} media files total`);

        return {
            success: true,
            message: `Successfully processed ${replacedOrAddedCount} media files`,
            replacedOrAddedCount
        };

    } catch (error) {
        console.error(`   ❌ Error in replaceSlideImages: ${error.message}`);
        throw error;
    }
}

async function mergeRelsFiles(originalRelsContent, convertedRelsContent, fileName) {
    try {
        // Extract slideLayout file name from original file
        const slideLayoutMatch = originalRelsContent.match(
            /<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slideLayout"[^>]+Target="([^"]+)"[^>]*>/i
        );

        // Remove notesSlide relationship from converted content
        let cleanedConvertedContent = convertedRelsContent.replace(
            /<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*>/ig,
            ''
        );

        if (!slideLayoutMatch) {
            return cleanedConvertedContent;
        }

        const originalSlideLayoutFileName = slideLayoutMatch[1];
        // Check if converted file has slideLayout relationship
        const convertedSlideLayoutMatch = cleanedConvertedContent.match(
            /<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slideLayout"[^>]+Target="([^"]+)"[^>]*>/i
        );

        let mergedContent = cleanedConvertedContent;

        if (convertedSlideLayoutMatch) {
            // Replace only the slideLayout file name in the converted file
            mergedContent = cleanedConvertedContent.replace(
                /(<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slideLayout"[^>]+Target=")[^"]+("[^>]*>)/i,
                `$1${originalSlideLayoutFileName}$2`
            );
        } else {
            // If converted file doesn't have slideLayout, add it with original file name
            const insertPosition = mergedContent.lastIndexOf('</Relationships>');
            if (insertPosition !== -1) {
                const newRelationship = `    <Relationship Id="rIdSlideLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${originalSlideLayoutFileName}"/>`;
                mergedContent =
                    mergedContent.slice(0, insertPosition) +
                    newRelationship + '\n' +
                    mergedContent.slice(insertPosition);
            } else {
                mergedContent = cleanedConvertedContent;
            }
        }

        return mergedContent;

    } catch (error) {
        console.error(`   ❌ Error merging rels files: ${error.message}`);
        return cleanedConvertedContent;
    }
}

async function normalizeChartReferences(fileFolderName, slideXmlsDir) {
    try {
        const filesDir = path.resolve(__dirname, '../files');
        const targetDir = path.join(filesDir, fileFolderName);

        // 1. Normalize slide XML chart references
        const slidesPath = path.join(slideXmlsDir, 'slides');
        const slideFiles = await fsPromises.readdir(slidesPath);

        for (const slideFile of slideFiles) {
            if (slideFile.match(/^slide\d+\.xml$/)) {
                const slidePath = path.join(slidesPath, slideFile);
                let slideContent = await fsPromises.readFile(slidePath, 'utf8');

                // Normalize chart references to chart1.xml
                slideContent = slideContent.replace(
                    /r:id="rId\d+" .*?\/ppt\/charts\/chart\d+\.xml/g,
                    'r:id="rId3" ../charts/chart1.xml'
                );

                await fsPromises.writeFile(slidePath, slideContent, 'utf8');
            }
        }

        // 2. Normalize rels file chart references
        const relsPath = path.join(slideXmlsDir, '_rels');
        const relsFiles = await fsPromises.readdir(relsPath);

        for (const relsFile of relsFiles) {
            if (relsFile.match(/^slide\d+\.xml\.rels$/)) {
                const relsFilePath = path.join(relsPath, relsFile);
                let relsContent = await fsPromises.readFile(relsFilePath, 'utf8');

                // Normalize chart relationship references
                relsContent = relsContent.replace(
                    /Target="\.\.\/charts\/chart\d+\.xml"/g,
                    'Target="../charts/chart1.xml"'
                );

                // Normalize relationship IDs for charts
                relsContent = relsContent.replace(
                    /<Relationship Id="rId\d+" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/chart" Target="\.\.\/charts\/chart1\.xml"/g,
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"'
                );

                await fsPromises.writeFile(relsFilePath, relsContent, 'utf8');
            }
        }

        // 3. Rename chart files in the target directory to chart1.xml
        const chartsDir = path.join(targetDir, 'ppt', 'charts');
        if (await checkDirectoryExists(chartsDir)) {
            const chartFiles = await fsPromises.readdir(chartsDir);
            const xmlChartFiles = chartFiles.filter(file => file.match(/^chart\d+\.xml$/));

            if (xmlChartFiles.length > 0) {
                // Find the latest chart file (highest number)
                const latestChart = xmlChartFiles.sort((a, b) => {
                    const aNum = parseInt(a.match(/\d+/)[0]);
                    const bNum = parseInt(b.match(/\d+/)[0]);
                    return bNum - aNum;
                })[0];

                const latestChartPath = path.join(chartsDir, latestChart);
                const targetChartPath = path.join(chartsDir, 'chart1.xml');

                // If it's not already chart1.xml, rename it
                if (latestChart !== 'chart1.xml') {
                    // Remove existing chart1.xml if it exists
                    try {
                        await fsPromises.unlink(targetChartPath);
                    } catch (error) {
                        // File doesn't exist, that's fine
                    }

                    // Rename latest chart to chart1.xml
                    await fsPromises.rename(latestChartPath, targetChartPath);

                    // Remove other chart files
                    for (const chartFile of xmlChartFiles) {
                        if (chartFile !== latestChart) {
                            try {
                                await fsPromises.unlink(path.join(chartsDir, chartFile));
                                console.log(`   Removed duplicate ${chartFile}`);
                            } catch (error) {
                                console.log(`   Could not remove ${chartFile}`);
                            }
                        }
                    }
                }
            }

            // 🆕 Remove style1.xml and colors1.xml files
            console.log('🗑️  Removing chart style and color files...');
            const filesToRemove = ['style1.xml', 'colors1.xml'];
            for (const fileToRemove of filesToRemove) {
                try {
                    const filePath = path.join(chartsDir, fileToRemove);
                    if (await checkFileExists(filePath)) {
                        await fsPromises.unlink(filePath);
                        console.log(`   ✅ Removed ${fileToRemove}`);
                    }
                } catch (error) {
                    console.log(`   ⚠️  Could not remove ${fileToRemove}: ${error.message}`);
                }
            }
        }

        // 4. Update [Content_Types].xml to remove style1.xml and colors1.xml Override entries
        console.log('📝 Updating [Content_Types].xml...');
        const contentTypesPath = path.join(targetDir, '[Content_Types].xml');

        if (await checkFileExists(contentTypesPath)) {
            try {
                let contentTypesXml = await fsPromises.readFile(contentTypesPath, 'utf8');

                // Remove style1.xml Override entry
                contentTypesXml = contentTypesXml.replace(
                    /<Override\s+PartName="\/ppt\/charts\/style1\.xml"[^>]*\/>\s*/g,
                    ''
                );

                // Remove colors1.xml Override entry
                contentTypesXml = contentTypesXml.replace(
                    /<Override\s+PartName="\/ppt\/charts\/colors1\.xml"[^>]*\/>\s*/g,
                    ''
                );

                await fsPromises.writeFile(contentTypesPath, contentTypesXml, 'utf8');
                console.log('   ✅ Updated [Content_Types].xml - removed style1.xml and colors1.xml entries');
            } catch (error) {
                console.log(`   ⚠️  Error updating [Content_Types].xml: ${error.message}`);
            }
        }

        return { success: true };

    } catch (error) {
        console.error('Error normalizing chart references:', error);
        return { success: false, error: error.message };
    }
}

// Helper function to check if file exists
async function checkFileExists(filePath) {
    try {
        await fsPromises.access(filePath);
        return true;
    } catch {
        return false;
    }
}

async function checkDirectoryExists(dirPath) {
    try {
        const stat = await fsPromises.stat(dirPath);
        return stat.isDirectory();
    } catch {
        return false;
    }
}

async function convertZipToPptxFile(sourceFolder, zipFileOutput, customName = null) {
    const zip = new JSZip();

    // Recursive function to add files and folders to zip
    function addToZip(currentPath, zipFolder = zip) {
        const items = fs.readdirSync(currentPath);

        items.forEach(item => {
            const itemPath = path.join(currentPath, item);
            const stat = fs.statSync(itemPath);

            if (stat.isFile()) {
                // Add file to zip
                const relativePath = path.relative(sourceFolder, itemPath);
                const fileContent = fs.readFileSync(itemPath);
                zipFolder.file(relativePath, fileContent);
            } else if (stat.isDirectory()) {
                // Create folder in zip and recursively add its contents
                const relativePath = path.relative(sourceFolder, itemPath);

                // Recursively process the directory
                addToZip(itemPath, zip);
            }
        });
    }

    try {
        // Verify source folder exists
        if (!fs.existsSync(sourceFolder)) {
            throw new Error(`Source folder does not exist: ${sourceFolder}`);
        }

        // Start the recursive process
        addToZip(sourceFolder);

        // Generate the zip buffer
        console.log(`🔄 Generating archive...`);
        const zipBuffer = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: {
                level: 6  // Better compression
            }
        });

        // Save the zip file to the output path
        fs.writeFileSync(zipFileOutput, zipBuffer);
        console.log(`🎉 ZIP file created successfully at: ${zipFileOutput}`);

        // 🛡️ PROTECTED RENAME FROM .zip TO .pptx
        const directory = path.dirname(zipFileOutput);
        const baseNameFromZip = path.basename(zipFileOutput, '.zip');
        let tempPptxName = `${baseNameFromZip}.pptx`;

        let tempPptxPath = path.join(directory, tempPptxName);

        // Check if PPTX already exists and create a safe temporary name
        if (fs.existsSync(tempPptxPath)) {
            console.log(`⚠️ File ${tempPptxName} already exists! Creating with temporary name...`);

            // Generate a unique temporary name
            const timestamp = Date.now();
            tempPptxName = `${baseNameFromZip}_temp_${timestamp}.pptx`;
            tempPptxPath = path.join(directory, tempPptxName);
        }

        // Safely rename to PPTX (with temporary name if needed)
        fs.renameSync(zipFileOutput, tempPptxPath);
        console.log(`🔄 Safely renamed to: ${tempPptxName}`);

        // Generate final filename with timestamp and optional custom name
        const finalBaseName = customName || baseNameFromZip;

        // Create timestamp for uniqueness
        const now = new Date();
        const timestamp = now.toISOString()
            .replace(/T/, '_')
            .replace(/\..+/, '')
            .replace(/:/g, '-');

        const dateStr = now.toISOString().slice(0, 10); // YYYY-MM-DD format

        // Generate final filename
        const finalFileName = path.join(directory, `${finalBaseName}_converted.pptx`);

        // Handle file conflicts by adding counter if needed
        let actualFileName = finalFileName;
        let counter = 1;
        while (fs.existsSync(actualFileName)) {
            const nameWithoutExt = finalFileName.replace('.pptx', '');
            actualFileName = `${nameWithoutExt}_v${counter}.pptx`;
            counter++;
        }

        // Final rename from temp to actual filename
        fs.renameSync(tempPptxPath, actualFileName);
        console.log(`🏷️ Final renamed file: ${path.basename(actualFileName)}`);

        // 🔍 Check if any existing files were preserved
        const originalPptx = path.join(directory, `${baseNameFromZip}.pptx`);

        // if (fs.existsSync(originalPptx) && originalPptx !== actualFileName) {
        //     console.log(`✅ Original file preserved: ${baseNameFromZip}.pptx`);
        // }

        return {
            success: true,
            originalZipPath: zipFileOutput,
            finalPptxPath: actualFileName,
            fileName: path.basename(actualFileName),
            size: zipBuffer.length,
            sizeKB: (zipBuffer.length / 1024).toFixed(2),
            originalFilePreserved: fs.existsSync(originalPptx) && originalPptx !== actualFileName
        };

    } catch (error) {
        console.error(`❌ Error creating zip/pptx file: ${error.message}`);
        throw error;
    }
}
async function fixTableCellAlignment(slideXmlsDir) {
    try {
        console.log('🔧 Fixing table cell alignment in merged cells...');

        const slidesDir = path.join(slideXmlsDir, 'slides');
        
        if (!await checkDirectoryExists(slidesDir)) {
            console.log('   ℹ️  No slides directory found');
            return { success: true, slidesFixed: 0 };
        }

        const slideFiles = await fsPromises.readdir(slidesDir);
        const slideXmlFiles = slideFiles.filter(file => file.match(/^slide\d+\.xml$/));
        
        let fixedCount = 0;

        for (const slideFile of slideXmlFiles) {
            try {
                const slidePath = path.join(slidesDir, slideFile);
                let slideContent = await fsPromises.readFile(slidePath, 'utf8');
                const originalContent = slideContent;

                // Find all table rows
                const rowRegex = /<a:tr[^>]*>([\s\S]*?)<\/a:tr>/g;
                let rowMatch;
                const rowsToFix = [];

                while ((rowMatch = rowRegex.exec(slideContent)) !== null) {
                    const rowContent = rowMatch[1];
                    const fullRow = rowMatch[0];
                    
                    // Check if row has merged cells (hMerge)
                    if (rowContent.includes('hMerge="1"')) {
                        // Find first cell with alignment
                        const firstCellMatch = rowContent.match(/<a:tc[^>]*>([\s\S]*?)<\/a:tc>/);
                        if (firstCellMatch) {
                            const firstCellContent = firstCellMatch[1];
                            
                            // Extract alignment from first cell
                            const alignMatch = firstCellContent.match(/<a:pPr\s+algn="([^"]+)"/);
                            if (alignMatch) {
                                const alignment = alignMatch[1];
                                console.log(`   📌 Found alignment="${alignment}" in merged row`);
                                
                                // Fix all hMerge cells in this row to have alignment
                                let fixedRow = fullRow.replace(
                                    /<a:tc\s+hMerge="1">([\s\S]*?)<\/a:tc>/g,
                                    (match, cellContent) => {
                                        // Check if cell already has <a:pPr>
                                        if (cellContent.includes('<a:pPr')) {
                                            // Already has pPr, just add alignment if missing
                                            if (!cellContent.includes('algn=')) {
                                                return match.replace(
                                                    /<a:pPr/,
                                                    `<a:pPr algn="${alignment}"`
                                                );
                                            }
                                            return match;
                                        } else {
                                            // No pPr, add it inside <a:p>
                                            return match.replace(
                                                /<a:p>/,
                                                `<a:p><a:pPr algn="${alignment}"/>`
                                            );
                                        }
                                    }
                                );
                                
                                rowsToFix.push({ original: fullRow, fixed: fixedRow });
                            }
                        }
                    }
                }

                // Apply all row fixes
                for (const { original, fixed } of rowsToFix) {
                    slideContent = slideContent.replace(original, fixed);
                }

                if (slideContent !== originalContent) {
                    await fsPromises.writeFile(slidePath, slideContent, 'utf8');
                    console.log(`   ✅ Fixed merged cell alignment in ${slideFile}`);
                    fixedCount++;
                }

            } catch (error) {
                console.error(`   ❌ Error fixing ${slideFile}: ${error.message}`);
            }
        }

        console.log(`   🎉 Table alignment fix complete - ${fixedCount} slide(s) fixed`);

        return {
            success: true,
            slidesFixed: fixedCount,
            totalSlides: slideXmlFiles.length
        };

    } catch (error) {
        console.error('   ❌ Error in fixTableCellAlignment:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    convertHTMLToPPTX
};