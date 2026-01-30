"use strict";

const express = require("express");
const router = express.Router();
const multer = require("multer");
const pptxToHtmlConverter = require("../controllers/pptxToHtmlConverter.controller");
const htmlToPptxConverter = require("../controllers/htmlToPptxConverter.controller");
const imageReplace = require("../controllers/imageReplace.controller");

// Configure multer for in-memory file uploads
const upload = multer({ storage: multer.memoryStorage() });


// Define routes
router.post("/upload", upload.single("pptxFile"), pptxToHtmlConverter.uploadFile);
router.get("/getSlides", pptxToHtmlConverter.getSlides);
router.post("/saveSlides", pptxToHtmlConverter.saveSlides);
router.get("/downloadHTML", pptxToHtmlConverter.downloadHTML);
router.post("/convertToPPTX", htmlToPptxConverter.convertToPPTX);

router.post("/replaceImage", imageReplace.replaceImage);

module.exports = router;
