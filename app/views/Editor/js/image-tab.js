/**
 * IMAGE FORMAT TAB FUNCTIONALITY
 * FIXED: All formatting applies to .image-container (NOT img)
 */

class ImageFormatTab {
    constructor(editor) {
        this.editor = editor;
        this.selectedImage = null;
        this.cropMode = false;
        this.aspectRatioLocked = true;
        this.originalAspectRatio = 1;
        this.originalImageSrc = null;
        this.cropOverlay = null;
        this.cropArea = null;

        this.init();
    }

    init() {
        this.setupToolbar();
        this.setupEventListeners();
    }

    /* -------------------------------
       SAFE HISTORY
    -------------------------------- */

    saveState() {
        if (this.editor && typeof this.editor.saveState === 'function') {
            this.editor.saveState();
        }
    }

    ensurePreStateSaved() {
        if (this.editor && Array.isArray(this.editor.history) && this.editor.history.length === 0) {
            this.editor.saveState();
        }
    }

    /* -------------------------------
       HELPERS
    -------------------------------- */

    getContainer() {
        return this.selectedImage;
    }

    getImg() {
        return this.selectedImage?.querySelector('img') || null;
    }

    /* -------------------------------
       SELECTION
    -------------------------------- */

    selectImage(imageElement) {
        document.querySelectorAll('.sli-txt-box, .shape, .custom-shape, .image-container')
            .forEach(el => el.classList.remove('selected'));

        this.selectedImage = imageElement;
        imageElement.classList.add('selected');

        const img = this.getImg();
        if (!img) return;

        if (!imageElement.dataset.originalSrc) {
            imageElement.dataset.originalSrc = img.src;
        }
        this.originalImageSrc = imageElement.dataset.originalSrc;

        const w = parseInt(imageElement.style.width) || imageElement.offsetWidth;
        const h = parseInt(imageElement.style.height) || imageElement.offsetHeight;

        this.originalAspectRatio = w / h;

        document.getElementById('imageWidth').value = Math.round(w);
        document.getElementById('imageHeight').value = Math.round(h);

        this.syncToolbarFromContainer();
    }

    /* -------------------------------
       TOOLBAR
    -------------------------------- */

    setupToolbar() {
        const imageToolbar = document.querySelector('.imageformat-toolbar');
        if (!imageToolbar) return;

        imageToolbar.innerHTML = `
            <!-- Crop & Change -->
            <div class="toolbar-group">
                <button class="tool-btn large" id="cropImage" title="Crop Image">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M6 2v14a2 2 0 002 2h14M18 6v14a2 2 0 01-2 2H2"/>
                    </svg>
                    <span>Crop</span>
                </button>
                <button class="tool-btn large" id="changeImage" title="Change Image">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <circle cx="8.5" cy="8.5" r="1.5"/>
                        <path d="M21 15l-5-5L5 21"/>
                    </svg>
                </button>

            </div>

            <!-- Picture Styles / Frames -->
            <div class="toolbar-group" style="display: none;">
                <select class="tool-select" id="imageFrameStyle" style="width: 130px;">
                    <option value="none">No Style</option>
                    <option value="simple-frame">Simple Frame</option>
                    <option value="rounded-rectangle">Rounded Rectangle</option>
                    <option value="circle">Circle</option>
                    <option value="soft-edge">Soft Edge Rectangle</option>
                    <option value="metal-frame">Metal Frame</option>
                    <option value="reflected-bevel">Reflected Bevel</option>
                    <option value="beveled">Beveled Matte</option>
                    <option value="compound-frame">Compound Frame</option>
                    <option value="center-shadow">Center Shadow</option>
                    <option value="drop-shadow">Drop Shadow</option>
                </select>
            </div>

            <!-- Corrections -->
            <div class="toolbar-group" style="display: none;">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Brightness</label>
            <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageBrightness" min="0" max="200" value="100" title="Brightness">
                <span class="slider-value" id="brightnessValue">100%</span>
            </div>
            </div>
            </div>

            <div class="toolbar-group" style="display: none;">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Contrast</label>
                <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageContrast" min="0" max="200" value="100" title="Contrast">
                <span class="slider-value" id="contrastValue">100%</span>
                </div>
            </div>
            </div>

            <div class="toolbar-group" style="display: none;">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Saturation</label>
                <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageSaturation" min="0" max="200" value="100" title="Saturation">
                <span class="slider-value" id="saturationValue">100%</span>
            </div>
            </div>
            </div>

            <!-- Color Adjustments -->
            <div class="toolbar-group" style="display: none;">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Hue Rotation</label>
                <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageHueRotate" min="0" max="360" value="0" title="Hue Rotation">
                <span class="slider-value" id="hueRotateValue">0°</span>
            </div>
            </div>
            </div>

            <div class="toolbar-group" style="display: none;">
                <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Filters</label>
                <div class="img-toolbar-slider" style="display: flex;">
                <button class="tool-btn" id="grayscaleImage" title="Grayscale">
                    <span style="filter: grayscale(1);">⬛</span>
                </button>
                <button class="tool-btn" id="sepiaImage" title="Sepia">
                    <span style="color: #704214;">📜</span>
                </button>
                <button class="tool-btn" id="invertImage" title="Invert Colors">
                    <span>🔄</span>
                </button>
            </div>
            </div>
            </div>

            <!-- Effects -->
            <div class="toolbar-group" style="display: none;">
             <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Effects</label>
                <div class="img-toolbar-slider" style="display: flex;">
                <button class="tool-btn" id="addShadow" title="Add Shadow">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                        <rect x="6" y="6" width="12" height="12" opacity="0.3" transform="translate(2, 2)"/>
                        <rect x="6" y="6" width="12" height="12"/>
                    </svg>
                </button>
                <button class="tool-btn" id="addGlow" title="Glow">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                        <circle cx="12" cy="12" r="8" opacity="0.3"/>
                        <circle cx="12" cy="12" r="4"/>
                    </svg>
                </button>
                <button class="tool-btn" id="softEdges" title="Soft Edges">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="6" y="6" width="12" height="12" rx="2" style="filter: blur(1px);"/>
                    </svg>
                </button>
                <button class="tool-btn" id="bevelEffect" title="3D Bevel">
                    <span>◪</span>
                </button>
            </div>
            </div>
            </div>

            <!-- Transparency & Blur -->
            <div class="toolbar-group">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Transparency</label>
                <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageTransparency" min="0" max="100" value="0" title="Transparency">
                <span class="slider-value" id="transparencyValue">0%</span>
            </div>
            </div>
            </div>

            <div class="toolbar-group" style="display: none;">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Blur</label>
                <div class="img-toolbar-slider">
                <input type="range" class="tool-slider" id="imageBlur" min="0" max="20" value="0" step="0.5" title="Blur">
                <span class="slider-value" id="blurValue">0px</span>
            </div>
            </div>
            </div>

            <!-- Arrange -->
            <div class="toolbar-group">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Arrange</label>
                 <div class="img-toolbar-subcontent" style="display: flex;">
                <button class="tool-btn" id="imgBringToFront" title="Bring to Front">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                        <rect x="8" y="4" width="12" height="12" opacity="0.3"/>
                        <rect x="4" y="8" width="12" height="12"/>
                    </svg>
                </button>
                <button class="tool-btn" id="imgSendToBack" title="Send to Back">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                        <rect x="4" y="8" width="12" height="12" opacity="0.3"/>
                        <rect x="8" y="4" width="12" height="12"/>
                    </svg>
                </button>
            </div>
            </div>
            </div>

            <!-- Rotate & Flip -->
            <div class="toolbar-group">
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Rotate</label>
                <div class="img-toolbar-subcontent" style="display: flex;">
                <button class="tool-btn" id="imgRotateLeft90" title="Rotate Left 90°">↶</button>
                <button class="tool-btn" id="imgRotateRight90" title="Rotate Right 90°">↷</button>
                <button class="tool-btn" id="imgFlipVertical" title="Flip Vertical">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M8 3H5a2 2 0 00-2 2v3m18 0V5a2 2 0 00-2-2h-3m0 18h3a2 2 0 002-2v-3M3 16v3a2 2 0 002 2h3"/>
                        <line x1="3" y1="12" x2="21" y2="12"/>
                    </svg>
                </button>
                <button class="tool-btn" id="imgFlipHorizontal" title="Flip Horizontal">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M8 3H5a2 2 0 00-2 2v3m18 0V5a2 2 0 00-2-2h-3m0 18h3a2 2 0 002-2v-3M3 16v3a2 2 0 002 2h3"/>
                        <line x1="12" y1="3" x2="12" y2="21"/>
                    </svg>
                </button>
                </div>
                </div>
            </div>

            <!-- Link & Alt Text -->
          <!--  <div class="toolbar-group">
            
            <div class="img-toolbar-subcontent">
                <label class="toolbar-label">Properties</label>
                <div class="img-toolbar-subcontent" style="display: flex;">
                <button class="tool-btn" id="addImageLink" title="Add Hyperlink">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71"/>
                        <path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71"/>
                    </svg>
                </button>
                <button class="tool-btn" id="editAltText" title="Alt Text">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/>
                        <text x="8" y="14" font-size="10" fill="currentColor">Alt</text>
                    </svg>
                </button>
            </div>
            </div>
            </div> -->


            <!-- Size -->
            <div class="toolbar-group">
                <label class="toolbar-label">Size</label>
                <input type="number" class="tool-input" id="imageWidth" value="300" min="10" max="1000" placeholder="Width" style="width: 80px;">
                <span style="color: #b0b0b0;">×</span>
                <input type="number" class="tool-input" id="imageHeight" value="200" min="10" max="1000" placeholder="Height" style="width: 80px;">
            </div>
        `;
    }

    setupEventListeners() {
        const on = (id, fn, ev = 'click') => {
            const el = document.getElementById(id);
            if (el) el.addEventListener(ev, fn);
        };
        on('cropImage', () => this.enableCropMode());
        on('changeImage', () => this.changeImage());
        on('imageFrameStyle', e => this.applyFrameStyle(e.target.value), 'change');

        on('imageBrightness', e => this.setFilter('brightness', e.target.value), 'input');
        on('imageContrast', e => this.setFilter('contrast', e.target.value), 'input');
        on('imageSaturation', e => this.setFilter('saturate', e.target.value), 'input');
        on('imageHueRotate', e => this.setFilter('hue-rotate', e.target.value), 'input');

        on('grayscaleImage', () => this.toggleSimpleFilter('grayscale'));
        on('sepiaImage', () => this.toggleSimpleFilter('sepia'));
        on('invertImage', () => this.toggleSimpleFilter('invert'));

        on('addShadow', () => this.toggleShadow());
        on('addGlow', () => this.toggleGlow());
        on('softEdges', () => this.toggleSoftEdges());
        on('bevelEffect', () => this.toggleBevel());

        on('imageTransparency', e => this.setTransparency(e.target.value), 'input');
        on('imageBlur', e => this.setBlur(e.target.value), 'input');

        on('imgRotateLeft90', () => this.rotateLeft());
        on('imgRotateRight90', () => this.rotateRight());
        on('imgFlipHorizontal', () => this.flipHorizontal());
        on('imgFlipVertical', () => this.flipVertical());

        // Arrange buttons — both image-tab and shape-tab buttons use the same shared utility
        on('imgBringToFront', () => this.bringToFront());
        on('imgSendToBack', () => this.sendToBack());
        on('bringToFront', () => this.bringToFront());
        on('sendToBack', () => this.sendToBack());

        on('imageWidth', e => this.updateImageSize('width', e.target.value), 'input');
        on('imageHeight', e => this.updateImageSize('height', e.target.value), 'input');

        // Clamp imageWidth / imageHeight to [min, max] in real-time
        ['imageWidth', 'imageHeight'].forEach(id => {
            const el = document.getElementById(id);
            if (!el) return;
            const MIN = parseInt(el.min) || 10;
            const MAX = parseInt(el.max) || 1000;
            el.addEventListener('input', () => {
                const val = parseFloat(el.value);
                if (!isNaN(val) && val > MAX) el.value = MAX;
                if (!isNaN(val) && val < MIN && el.value.length >= String(MIN).length) el.value = MIN;
            });
            el.addEventListener('blur', () => {
                const val = parseFloat(el.value);
                if (isNaN(val) || val < MIN) el.value = MIN;
                if (val > MAX) el.value = MAX;
            });
        });

        document.getElementById('canvas').addEventListener('click', e => {
            const c = e.target.closest('.image-container');
            if (c) this.selectImage(c);
        });

        ['undoBtn', 'redoBtn', 'undoBtnTop', 'redoBtnTop'].forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                btn.addEventListener('click', () => {
                    if (this.cropMode) {
                        this.disableCropMode();
                    }
                }, true); // capture phase
            }
        });
    }


    // ============================================
    // CROP & CHANGE
    // ============================================

    enableCropMode() {
        if (!this.selectedImage) {
            alert('Please select an image first');
            return;
        }

        if (this.cropMode) {
            // Already in crop mode, disable it
            this.disableCropMode();
            return;
        }

        this.cropMode = true;

        // Create crop overlay
        this.createCropOverlay();
    }

    createCropOverlay() {
        const img = this.selectedImage?.querySelector('img');
        if (!img) return;

        const canvas = document.getElementById('canvas');
        const canvasRect = canvas.getBoundingClientRect();
        const rect = this.selectedImage.getBoundingClientRect();

        // FULL CANVAS OVERLAY (blocks interactions with everything else)
        this.cropOverlay = document.createElement('div');
        this.cropOverlay.className = 'crop-overlay active';
        this.cropOverlay.style.position = 'absolute';
        this.cropOverlay.style.left = '0px';
        this.cropOverlay.style.top = '0px';
        this.cropOverlay.style.width = canvas.offsetWidth + 'px';
        this.cropOverlay.style.height = canvas.offsetHeight + 'px';
        this.cropOverlay.style.zIndex = '10000';
        this.cropOverlay.style.pointerEvents = 'auto';
        this.cropOverlay.style.background = 'rgba(0,0,0,0.02)'; // almost transparent, but blocks hover/click

        // Frame positioned exactly over the image container
        const frame = document.createElement('div');
        frame.className = 'crop-frame';
        frame.style.position = 'absolute';
        frame.style.left = (rect.left - canvasRect.left) + 'px';
        frame.style.top = (rect.top - canvasRect.top) + 'px';
        frame.style.width = rect.width + 'px';
        frame.style.height = rect.height + 'px';
        frame.style.pointerEvents = 'auto';

        // Create crop area (initially 80%)
        this.cropArea = document.createElement('div');
        this.cropArea.className = 'crop-area';
        this.cropArea.style.position = 'absolute';
        this.cropArea.style.left = '10%';
        this.cropArea.style.top = '10%';
        this.cropArea.style.width = '80%';
        this.cropArea.style.height = '80%';
        this.cropArea.style.border = '2px dashed #ff6c36';
        this.cropArea.style.background = 'rgba(255, 108, 54, 0.1)';
        this.cropArea.style.cursor = 'move';
        this.cropArea.style.pointerEvents = 'auto';

        // Add crop handles
        const positions = ['nw', 'ne', 'sw', 'se', 'n', 's', 'e', 'w'];
        positions.forEach(pos => {
            const handle = document.createElement('div');
            handle.className = `crop-handle ${pos}`;
            handle.dataset.position = pos;
            this.cropArea.appendChild(handle);
        });

        // Buttons
        const buttonContainer = document.createElement('div');
        buttonContainer.style.position = 'absolute';
        buttonContainer.style.bottom = '-50px';
        buttonContainer.style.left = '50%';
        buttonContainer.style.transform = 'translateX(-50%)';
        buttonContainer.style.display = 'flex';
        buttonContainer.style.gap = '10px';
        buttonContainer.style.pointerEvents = 'auto';

        const applyCropBtn = document.createElement('button');
        applyCropBtn.className = 'tool-btn large';
        applyCropBtn.innerHTML = '<span>✓ Apply Crop</span>';
        applyCropBtn.style.background = '#4CAF50';
        applyCropBtn.style.color = 'white';
        applyCropBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); this.applyCrop(); };

        const cancelCropBtn = document.createElement('button');
        cancelCropBtn.className = 'tool-btn large';
        cancelCropBtn.innerHTML = '<span>✗ Cancel</span>';
        cancelCropBtn.style.background = '#f44336';
        cancelCropBtn.style.color = 'white';
        cancelCropBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); this.disableCropMode(); };

        buttonContainer.appendChild(applyCropBtn);
        buttonContainer.appendChild(cancelCropBtn);
        this.cropArea.appendChild(buttonContainer);

        frame.appendChild(this.cropArea);
        this.cropOverlay.appendChild(frame);
        canvas.appendChild(this.cropOverlay);

        // Prevent clicks anywhere else (belt + suspenders)
        this.cropOverlay.addEventListener('mousedown', (e) => {
            // allow crop-area interactions
            if (e.target.closest('.crop-area')) return;
            e.preventDefault();
            e.stopPropagation();
        }, true);

        this.cropOverlay.addEventListener('click', (e) => {
            if (e.target.closest('.crop-area')) return;
            e.preventDefault();
            e.stopPropagation();
        }, true);

        // Make crop area draggable/resizable
        this.makeCropAreaDraggable();
    }


    makeCropAreaDraggable() {
        let isDragging = false;
        let isResizing = false;
        let resizeHandle = null;
        let startX, startY, startLeft, startTop, startWidth, startHeight;

        // Dragging crop area
        this.cropArea.addEventListener('mousedown', (e) => {
            if (e.target.classList.contains('crop-handle')) {
                // Start resizing
                isResizing = true;
                resizeHandle = e.target.dataset.position;
                startX = e.clientX;
                startY = e.clientY;
                startLeft = parseFloat(this.cropArea.style.left);
                startTop = parseFloat(this.cropArea.style.top);
                startWidth = parseFloat(this.cropArea.style.width);
                startHeight = parseFloat(this.cropArea.style.height);
                e.preventDefault();
                return;
            }

            // Start dragging
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            startLeft = parseFloat(this.cropArea.style.left);
            startTop = parseFloat(this.cropArea.style.top);
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (isResizing) {
                const deltaX = ((e.clientX - startX) / this.cropOverlay.offsetWidth) * 100;
                const deltaY = ((e.clientY - startY) / this.cropOverlay.offsetHeight) * 100;

                let newLeft = startLeft;
                let newTop = startTop;
                let newWidth = startWidth;
                let newHeight = startHeight;

                // Handle different resize directions
                switch (resizeHandle) {
                    case 'nw': // North-West
                        newLeft = startLeft + deltaX;
                        newTop = startTop + deltaY;
                        newWidth = startWidth - deltaX;
                        newHeight = startHeight - deltaY;
                        break;
                    case 'ne': // North-East
                        newTop = startTop + deltaY;
                        newWidth = startWidth + deltaX;
                        newHeight = startHeight - deltaY;
                        break;
                    case 'sw': // South-West
                        newLeft = startLeft + deltaX;
                        newWidth = startWidth - deltaX;
                        newHeight = startHeight + deltaY;
                        break;
                    case 'se': // South-East
                        newWidth = startWidth + deltaX;
                        newHeight = startHeight + deltaY;
                        break;
                    case 'n': // North
                        newTop = startTop + deltaY;
                        newHeight = startHeight - deltaY;
                        break;
                    case 's': // South
                        newHeight = startHeight + deltaY;
                        break;
                    case 'w': // West
                        newLeft = startLeft + deltaX;
                        newWidth = startWidth - deltaX;
                        break;
                    case 'e': // East
                        newWidth = startWidth + deltaX;
                        break;
                }

                // Keep within bounds (minimum 10%, maximum 100%)
                newWidth = Math.max(10, Math.min(newWidth, 100 - newLeft));
                newHeight = Math.max(10, Math.min(newHeight, 100 - newTop));
                newLeft = Math.max(0, Math.min(newLeft, 100 - newWidth));
                newTop = Math.max(0, Math.min(newTop, 100 - newHeight));

                this.cropArea.style.left = newLeft + '%';
                this.cropArea.style.top = newTop + '%';
                this.cropArea.style.width = newWidth + '%';
                this.cropArea.style.height = newHeight + '%';

            } else if (isDragging) {
                const deltaX = ((e.clientX - startX) / this.cropOverlay.offsetWidth) * 100;
                const deltaY = ((e.clientY - startY) / this.cropOverlay.offsetHeight) * 100;

                let newLeft = startLeft + deltaX;
                let newTop = startTop + deltaY;

                // Keep within bounds
                const width = parseFloat(this.cropArea.style.width);
                const height = parseFloat(this.cropArea.style.height);

                newLeft = Math.max(0, Math.min(newLeft, 100 - width));
                newTop = Math.max(0, Math.min(newTop, 100 - height));

                this.cropArea.style.left = newLeft + '%';
                this.cropArea.style.top = newTop + '%';
            }
        });

        document.addEventListener('mouseup', () => {
            isDragging = false;
            isResizing = false;
            resizeHandle = null;
        });
    }

    applyCrop() {
        if (!this.selectedImage || !this.cropArea || !this.cropOverlay) return;

        const img = this.selectedImage.querySelector('img');
        if (!img) return;

        // Save original container geometry ONCE
        if (!this.selectedImage.dataset.origGeomSaved) {
            this.selectedImage.dataset.origGeomSaved = '1';
            this.selectedImage.dataset.origLeft = this.selectedImage.style.left || '';
            this.selectedImage.dataset.origTop = this.selectedImage.style.top || '';
            this.selectedImage.dataset.origWidth = this.selectedImage.style.width || '';
            this.selectedImage.dataset.origHeight = this.selectedImage.style.height || '';
        }

        // ✅ 1) Capture crop box values BEFORE removing overlay
        const cropLeft = parseFloat(this.cropArea.style.left);
        const cropTop = parseFloat(this.cropArea.style.top);
        const cropWidth = parseFloat(this.cropArea.style.width);
        const cropHeight = parseFloat(this.cropArea.style.height);

        // Container size (unchanged)
        const containerWidth = this.selectedImage.offsetWidth;
        const containerHeight = this.selectedImage.offsetHeight;

        // ✅ 2) Remove crop overlay BEFORE saving history state
        // This prevents undo restoring overlay/buttons.
        this.disableCropMode();

        // ✅ 3) Ensure undo has baseline, then save "before crop" WITHOUT overlay
        this.ensurePreStateSaved();
        this.editor.saveState();

        // Convert % crop rect to px
        const cropLeftPx = (cropLeft / 100) * containerWidth;
        const cropTopPx = (cropTop / 100) * containerHeight;
        const cropWidthPx = (cropWidth / 100) * containerWidth;
        const cropHeightPx = (cropHeight / 100) * containerHeight;

        // Create canvas for crop output
        const c = document.createElement('canvas');
        const ctx = c.getContext('2d');

        c.width = Math.max(1, Math.round(cropWidthPx));
        c.height = Math.max(1, Math.round(cropHeightPx));

        const imgW = img.naturalWidth || img.width;
        const imgH = img.naturalHeight || img.height;

        const scaleX = imgW / containerWidth;
        const scaleY = imgH / containerHeight;

        const sx = cropLeftPx * scaleX;
        const sy = cropTopPx * scaleY;
        const sw = cropWidthPx * scaleX;
        const sh = cropHeightPx * scaleY;

        ctx.drawImage(img, sx, sy, sw, sh, 0, 0, c.width, c.height);

        // Apply cropped bitmap (container unchanged)
        img.src = c.toDataURL('image/png');

        img.style.width = '100%';
        img.style.height = '100%';
        img.style.objectFit = 'fill';
        img.style.objectPosition = 'center';
        img.style.clipPath = '';
        img.style.webkitClipPath = '';
        img.style.transform = '';
        img.style.filter = '';

        this.selectedImage.dataset.isCropped = '1';

        // ✅ 4) Save "after crop" (clean DOM, no overlay)
        this.editor.saveState();
    }


    disableCropMode() {
        this.cropMode = false;

        if (this.cropOverlay) {
            this.cropOverlay.remove();
            this.cropOverlay = null;
            this.cropArea = null;
        }

    }



    changeImage() {
        if (!this.selectedImage) {
            alert('Please select an image first');
            return;
        }

        const img = this.selectedImage.querySelector('img');
        if (!img) return;

        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';

        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (ev) => {
                // IMPORTANT: do NOT overwrite originalSrc
                img.src = ev.target.result;
                this.saveState();
            };
            reader.readAsDataURL(file);
        };

        input.click();
    }



    updateCurrentFilters() {
        if (!this.selectedImage) return;

        // Parse current filter values
        const c = this.selectedImage;
        const filterStyle = c.style.filter || '';

        // Transparency (opacity)
        const opacity = parseFloat(c.style.opacity || 1);
        const transparencyPercent = Math.round((1 - opacity) * 100);
        document.getElementById('imageTransparency').value = transparencyPercent;
        document.getElementById('transparencyValue').textContent = transparencyPercent + '%';

        // Blur
        const blurMatch = filterStyle.match(/blur\((\d+\.?\d*)px\)/);
        const blurValue = blurMatch ? parseFloat(blurMatch[1]) : 0;
        document.getElementById('imageBlur').value = blurValue;
        document.getElementById('blurValue').textContent = blurValue + 'px';

        // Brightness
        const brightnessMatch = filterStyle.match(/brightness\((\d+\.?\d*)%?\)/);
        const brightnessValue = brightnessMatch ? parseFloat(brightnessMatch[1]) : 100;
        document.getElementById('imageBrightness').value = brightnessValue;
        document.getElementById('brightnessValue').textContent = brightnessValue + '%';

        // Contrast
        const contrastMatch = filterStyle.match(/contrast\((\d+\.?\d*)%?\)/);
        const contrastValue = contrastMatch ? parseFloat(contrastMatch[1]) : 100;
        document.getElementById('imageContrast').value = contrastValue;
        document.getElementById('contrastValue').textContent = contrastValue + '%';

        // Saturation
        const saturationMatch = filterStyle.match(/saturate\((\d+\.?\d*)%?\)/);
        const saturationValue = saturationMatch ? parseFloat(saturationMatch[1]) : 100;
        document.getElementById('imageSaturation').value = saturationValue;
        document.getElementById('saturationValue').textContent = saturationValue + '%';

        // Hue Rotate
        const hueRotateMatch = filterStyle.match(/hue-rotate\((\d+\.?\d*)deg\)/);
        const hueRotateValue = hueRotateMatch ? parseFloat(hueRotateMatch[1]) : 0;
        document.getElementById('imageHueRotate').value = hueRotateValue;
        document.getElementById('hueRotateValue').textContent = hueRotateValue + '°';
    }

    /* -------------------------------
       FRAME STYLES (CONTAINER)
    -------------------------------- */

    applyFrameStyle(style) {
        const c = this.getContainer();
        if (!c) return;

        c.style.border = 'none';
        c.style.borderRadius = '0';
        c.style.boxShadow = 'none';

        switch (style) {
            case 'simple-frame':
                c.style.border = '3px solid #666';
                break;
            case 'rounded-rectangle':
                c.style.border = '2px solid #888';
                c.style.borderRadius = '15px';
                break;
            case 'circle':
                c.style.border = '2px solid #888';
                c.style.borderRadius = '50%';
                break;
            case 'soft-edge':
                c.style.boxShadow = '0 0 15px rgba(0,0,0,0.3)';
                break;
            case 'metal-frame':
                c.style.border = '4px solid #aaa';
                c.style.boxShadow = 'inset 0 0 10px rgba(0,0,0,.4)';
                break;
            case 'drop-shadow':
                c.style.boxShadow = '8px 8px 15px rgba(0,0,0,.4)';
                break;
        }

        this.saveState();
    }

    /* -------------------------------
       FILTER PIPELINE (CONTAINER)
    -------------------------------- */

    updateSliderLabels() {
        const setText = (id, txt) => {
            const el = document.getElementById(id);
            if (el) el.textContent = txt;
        };

        const b = document.getElementById('imageBrightness')?.value ?? 100;
        const c = document.getElementById('imageContrast')?.value ?? 100;
        const s = document.getElementById('imageSaturation')?.value ?? 100;
        const h = document.getElementById('imageHueRotate')?.value ?? 0;
        const t = document.getElementById('imageTransparency')?.value ?? 0;
        const blur = document.getElementById('imageBlur')?.value ?? 0;

        setText('brightnessValue', `${b}%`);
        setText('contrastValue', `${c}%`);
        setText('saturationValue', `${s}%`);
        setText('hueRotateValue', `${h}°`);
        setText('transparencyValue', `${t}%`);
        setText('blurValue', `${blur}px`);
    }


    setFilter(type, value) {
        const c = this.getContainer();
        if (!c) return;

        // FIX: dataset keys cannot contain hyphens
        const key = (type === 'hue-rotate') ? 'hueRotate' : type;

        c.dataset[key] = value;
        this.rebuildFilter();
        this.updateSliderLabels();
        this.saveState();
    }


    toggleSimpleFilter(name) {
        const c = this.getContainer();
        if (!c) return;

        c.dataset[name] = c.dataset[name] ? '' : '100';
        this.rebuildFilter();
        this.saveState();
    }

    setBlur(v) {
        const c = this.getContainer();
        if (!c) return;
        c.dataset.blur = v;
        this.rebuildFilter();
        document.getElementById('blurValue').textContent = v + 'px';
        this.saveState();
    }

    rebuildFilter() {
        const c = this.getContainer();
        if (!c) return;

        const f = [];
        if (c.dataset.brightness) f.push(`brightness(${c.dataset.brightness}%)`);
        if (c.dataset.contrast) f.push(`contrast(${c.dataset.contrast}%)`);
        if (c.dataset.saturate) f.push(`saturate(${c.dataset.saturate}%)`);
        if (c.dataset.hueRotate) f.push(`hue-rotate(${c.dataset.hueRotate}deg)`);
        if (c.dataset.grayscale) f.push('grayscale(100%)');
        if (c.dataset.sepia) f.push('sepia(100%)');
        if (c.dataset.invert) f.push('invert(100%)');
        if (c.dataset.blur && +c.dataset.blur > 0) f.push(`blur(${c.dataset.blur}px)`);

        c.style.filter = f.join(' ');
    }

    /* -------------------------------
       EFFECTS
    -------------------------------- */

    toggleShadow() {
        const c = this.getContainer();
        if (!c) return;
        c.style.boxShadow = c.style.boxShadow ? '' : '0 10px 25px rgba(0,0,0,.35)';
        this.saveState();
    }

    toggleGlow() {
        const c = this.getContainer();
        if (!c) return;
        c.style.boxShadow = c.style.boxShadow ? '' : '0 0 18px rgba(255,255,255,.6)';
        this.saveState();
    }

    toggleSoftEdges() {
        const c = this.getContainer();
        if (!c) return;
        c.style.borderRadius = c.style.borderRadius ? '0' : '16px';
        this.saveState();
    }

    toggleBevel() {
        const c = this.getContainer();
        if (!c) return;
        c.style.boxShadow = c.style.boxShadow
            ? ''
            : 'inset 2px 2px 4px rgba(255,255,255,.5), inset -2px -2px 4px rgba(0,0,0,.4)';
        this.saveState();
    }

    /* -------------------------------
       TRANSPARENCY
    -------------------------------- */

    setTransparency(v) {
        const c = this.getImg();
        if (!c) return;
        c.style.opacity = 1 - v / 100;
        document.getElementById('transparencyValue').textContent = v + '%';
        this.saveState();
    }

    /* -------------------------------
       TRANSFORMS - UPDATED WITH SINGLE-FLIP LOGIC
    -------------------------------- */

    // Normalize rotation to 0-360 range
    normalizeRotation(angle) {
        return ((angle % 360) + 360) % 360;
    }

    // Get current transform state from container
    getTransformState() {
        const c = this.getContainer();
        if (!c) return { rotation: 0, scaleX: 1, scaleY: 1 };

        return {
            rotation: parseInt(c.dataset.rotation || 0),
            scaleX: parseFloat(c.dataset.scaleX || 1),
            scaleY: parseFloat(c.dataset.scaleY || 1)
        };
    }

    // Rotate Right (clockwise 90 degrees)
    rotateRight() {
        const c = this.getContainer();
        if (!c) return;

        const state = this.getTransformState();
        state.rotation = this.normalizeRotation(state.rotation + 90);

        c.dataset.rotation = state.rotation;
        this.applyTransform();
        this.saveState();
    }

    // Rotate Left (counter-clockwise 90 degrees)
    rotateLeft() {
        const c = this.getContainer();
        if (!c) return;

        const state = this.getTransformState();
        state.rotation = this.normalizeRotation(state.rotation - 90);

        c.dataset.rotation = state.rotation;
        this.applyTransform();
        this.saveState();
    }

    // Flip Horizontal - Only ONE flip can be active at a time
    flipHorizontal() {
        const c = this.getContainer();
        if (!c) return;

        const state = this.getTransformState();

        // Toggle horizontal flip
        if (state.scaleX === -1) {
            // Already flipped horizontally, remove flip
            state.scaleX = 1;
        } else {
            // Flip horizontally and reset vertical flip
            state.scaleX = -1;
            state.scaleY = 1;  // Reset vertical flip
        }

        c.dataset.scaleX = state.scaleX;
        c.dataset.scaleY = state.scaleY;
        this.applyTransform();
        this.saveState();
    }

    // Flip Vertical - Only ONE flip can be active at a time
    flipVertical() {
        const c = this.getContainer();
        if (!c) return;

        const state = this.getTransformState();

        // Toggle vertical flip
        if (state.scaleY === -1) {
            // Already flipped vertically, remove flip
            state.scaleY = 1;
        } else {
            // Flip vertically and reset horizontal flip
            state.scaleY = -1;
            state.scaleX = 1;  // Reset horizontal flip
        }

        c.dataset.scaleX = state.scaleX;
        c.dataset.scaleY = state.scaleY;
        this.applyTransform();
        this.saveState();
    }

    // Apply the transform based on current state
    applyTransform() {
        const c = this.getContainer();
        if (!c) return;

        const state = this.getTransformState();
        let transform = '';

        // Add rotation if not 0
        if (state.rotation !== 0) {
            transform += `rotate(${state.rotation}deg)`;
        }

        // Add scaleX if flipped horizontally
        if (state.scaleX === -1) {
            transform += (transform ? ' ' : '') + 'scaleX(-1)';
        }

        // Add scaleY if flipped vertically
        if (state.scaleY === -1) {
            transform += (transform ? ' ' : '') + 'scaleY(-1)';
        }

        c.style.transform = transform || 'none';
    }

    /* -------------------------------
       ARRANGE (Z-INDEX)
    -------------------------------- */

    bringToFront() {
        ArrangeUtils.bringToFront(() => this.saveState());
    }

    sendToBack() {
        ArrangeUtils.sendToBack(() => this.saveState());
    }


    /* -------------------------------
       SIZE
    -------------------------------- */

    updateImageSize(type, value) {
        const c = this.getContainer();
        if (!c) return;

        let v = parseFloat(value);
        if (isNaN(v) || v <= 0) return;
        v = Math.min(1000, Math.max(10, v)); // enforce min:10 max:1000

        if (type === 'width') {
            c.style.width = v + 'px';
            if (this.aspectRatioLocked) {
                c.style.height = (v / this.originalAspectRatio) + 'px';
            }
        } else {
            c.style.height = v + 'px';
            if (this.aspectRatioLocked) {
                c.style.width = (v * this.originalAspectRatio) + 'px';
            }
        }

        this.saveState();
    }

    /* -------------------------------
       UI SYNC
    -------------------------------- */

    syncToolbarFromContainer() {
        const c = this.getContainer();
        if (!c) return;

        document.getElementById('imageTransparency').value =
            Math.round((1 - (parseFloat(c.style.opacity) || 1)) * 100);
    }
}

/* -------------------------------
   INIT
-------------------------------- */

document.addEventListener('DOMContentLoaded', () => {
    if (window.editor) {
        window.imageFormatTab = new ImageFormatTab(window.editor);
    }
});