class PresentationEditor {
    constructor() {
        this.currentSlide = 1;
        this.totalSlides = 3;
        this.zoomLevel = 64;
        this.history = [];
        this.historyIndex = -1;
        this.maxHistory = 50;
        this.lastAssignedZIndex = 0; // Track last assigned z-index for incremental values
        this._textMeasurerEl = null;
        this._textAutoGrowRaf = null;

        // 3-state text-box click machine (mirrors toolbar.js approach)
        this._txtClickBox     = null;   // the .shape wrapper currently tracked
        this._isShapeSelected = false;  // State 1: shape selected / drag mode
        this._isTextSelected  = false;  // State 2: all text selected
        this._isEditing       = false;  // State 3: direct contenteditable edit

        this.init();
    }

    init() {
        this.setupCoreEventListeners();
        this.setupSlideDeleteButtons();
        this.updateUI();
        this.updateUndoRedoButtons();

        // ✅ Set initial slide visibility
        this.selectSlide(1);

        // ✅ Try to save initial state (will defer if canvas is empty)
        this.saveInitialState();
        
        // ✅ Also try again after a delay for dynamically loaded slides
        setTimeout(() => this.saveInitialState(), 500);
        
        // Draw/Image toolbars (shape formatting, etc.)
        this.setupShapeFormatToolbar();
        this.normalizeExistingShapes();
    }

    setupSlideDeleteButtons() {
        const allSlides = document.querySelectorAll('.slide-thumb');
        allSlides.forEach((slide) => {
            // Add delete button if it doesn't exist
            if (!slide.querySelector('.slide-delete-btn')) {
                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'slide-delete-btn';
                deleteBtn.innerHTML = '×';
                deleteBtn.title = 'Delete slide';

                deleteBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const slideNumber = parseInt(e.currentTarget.closest('.slide-thumb').dataset.slide);
                    this.deleteSlide(slideNumber);
                });

                slide.appendChild(deleteBtn);
            }
        });
    }

    /**
     * Generate slide thumbnails from loaded slides
     * Called after slides are injected into canvas
     */
   generateSlidePreviews() {
    const canvas = document.querySelector('.canvas');
    const slidesContainer = document.querySelector('.slides-container');
    
    if (!canvas || !slidesContainer) {
        console.warn('Canvas or slides container not found');
        return;
    }

    // Get all sli-slide elements from canvas
    const slides = canvas.querySelectorAll('.sli-slide');
    
    if (slides.length === 0) {
        console.warn('No slides found in canvas');
        return;
    }

    // Clear existing thumbnails
    slidesContainer.innerHTML = '';

    // Update total slides count
    this.totalSlides = slides.length;

    // Generate thumbnail for each slide
    slides.forEach((slide, index) => {
        const slideNumber = index + 1;
        
        // Create thumbnail wrapper
        const thumbWrapper = document.createElement('div');
        thumbWrapper.className = 'slide-thumb';
        thumbWrapper.dataset.slide = slideNumber;
        if (slideNumber === this.currentSlide) {
            thumbWrapper.classList.add('active');
        }

        // Add slide number badge
        const numberBadge = document.createElement('div');
        numberBadge.className = 'slide-number';
        numberBadge.textContent = slideNumber;
        thumbWrapper.appendChild(numberBadge);

        // Create preview container
        const previewDiv = document.createElement('div');
        previewDiv.className = 'sli-preview';

        // Clone the slide content for preview
        const slideClone = slide.cloneNode(true);
        
        // ✅ CRITICAL: Disable all interactions on preview elements
        slideClone.style.pointerEvents = 'none';
        
        // Remove all event listeners and interactive elements from clone
        slideClone.querySelectorAll('.resize-handle, .delete-btn, .rotate-handle').forEach(el => el.remove());
        slideClone.querySelectorAll('.shape, .custom-shape, .sli-txt-box, .image-container').forEach(el => {
            el.style.pointerEvents = 'none';
            el.classList.remove('selected', 'hover');
            el.removeAttribute('contenteditable');
        });
        
        // Calculate scale to fit preview container
        const previewWidth = previewDiv.offsetWidth || 250;
        const scale = previewWidth / 1040;
        
        slideClone.style.transform = `scale(${scale})`;
        slideClone.style.transformOrigin = 'top left';
        slideClone.style.width = '960px';
        slideClone.style.height = '540px';
        slideClone.style.position = 'absolute';
        slideClone.style.top = '0';
        slideClone.style.left = '0';

        // Wrapper for scaled content with proper dimensions
        const scaleWrapper = document.createElement('div');
        scaleWrapper.style.width = `${960 * scale}px`;
        scaleWrapper.style.height = `${540 * scale}px`;
        scaleWrapper.style.position = 'relative';
        scaleWrapper.style.overflow = 'hidden';
        scaleWrapper.style.margin = '0 auto';
        scaleWrapper.appendChild(slideClone);

        previewDiv.appendChild(scaleWrapper);
        thumbWrapper.appendChild(previewDiv);

        // Add delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'slide-delete-btn';
        deleteBtn.innerHTML = '×';
        deleteBtn.title = 'Delete slide';
        deleteBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.deleteSlide(slideNumber);
        });
        thumbWrapper.appendChild(deleteBtn);

        // Add click handler to switch slides
        thumbWrapper.addEventListener('click', () => {
            this.selectSlide(slideNumber);
        });

        slidesContainer.appendChild(thumbWrapper);
    });

    // Update UI
    this.updateUI();
    
    console.log(`Generated ${slides.length} slide previews`);
}

    /**
     * Refresh a specific slide preview
     */
   refreshSlidePreview(slideNumber) {
    const canvas = document.querySelector('.canvas');
    const slideThumb = document.querySelector(`.slide-thumb[data-slide="${slideNumber}"]`);
    
    if (!canvas || !slideThumb) return;

    const slides = canvas.querySelectorAll('.sli-slide');
    const slide = slides[slideNumber - 1];
    
    if (!slide) return;

    // Find and update the preview
    const previewDiv = slideThumb.querySelector('.sli-preview');
    if (!previewDiv) return;

    // Clear existing preview content
    previewDiv.innerHTML = '';

    // Clone the slide content for preview
    const slideClone = slide.cloneNode(true);
    
    // ✅ CRITICAL: Disable all interactions on preview elements
    slideClone.style.pointerEvents = 'none';
    
    // Remove all event listeners and interactive elements from clone
    slideClone.querySelectorAll('.resize-handle, .delete-btn, .rotate-handle').forEach(el => el.remove());
    slideClone.querySelectorAll('.shape, .custom-shape, .sli-txt-box, .image-container').forEach(el => {
        el.style.pointerEvents = 'none';
        el.classList.remove('selected', 'hover');
        el.removeAttribute('contenteditable');
    });
    
    // Calculate scale to fit preview container
    const previewWidth = previewDiv.offsetWidth || 250;
    const scale = previewWidth / 860;
    
    slideClone.style.transform = `scale(${scale})`;
    slideClone.style.transformOrigin = 'top left';
    slideClone.style.width = '960px';
    slideClone.style.height = '540px';
    slideClone.style.position = 'absolute';
    slideClone.style.top = '0';
    slideClone.style.left = '0';

    // Wrapper for scaled content with proper dimensions
    const scaleWrapper = document.createElement('div');
    scaleWrapper.style.width = `${960 * scale}px`;
    scaleWrapper.style.height = `${540 * scale}px`;
    scaleWrapper.style.position = 'relative';
    scaleWrapper.style.overflow = 'hidden';
    scaleWrapper.style.margin = '0 auto';
    scaleWrapper.appendChild(slideClone);

    previewDiv.appendChild(scaleWrapper);
}

    resetShapeDefaults(shape) {
        shape.style.opacity = '1';
        shape.style.filter = '';
        shape.style.boxShadow = '';
        shape.style.transform = 'none';
        shape.style.borderRadius = shape.style.borderRadius || '0';
    }

    /**
     * Get the highest z-index from all elements in the current slide's sli-content
     * Returns highest + 1 for new elements, and tracks last assigned value for incremental assignment
     */
    getHighestZIndex() {
        const canvas = document.getElementById('canvas');
        if (!canvas) return 1;

        // Find currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) return 1;

        // Check elements within sli-content of current slide
        const sliContent = currentSlide.querySelector('.sli-content');
        if (!sliContent) {
            // Fallback: check directly in slide
            const elements = currentSlide.querySelectorAll('.shape, .custom-shape, .shape-group, .image-container, .sli-txt-box, .insertable-element, .chart-element, .table-element');
            let highest = 0;
            
            elements.forEach(el => {
                const zIndex = parseInt(el.style.zIndex || 0);
                if (zIndex > highest) {
                    highest = zIndex;
                }
            });
            
            // Use the maximum of: highest in DOM, or last assigned value
            // This ensures each call returns an incremental value even in quick succession
            const nextZ = Math.max(highest, this.lastAssignedZIndex) + 1;
            this.lastAssignedZIndex = nextZ;
            return nextZ;
        }

        // Check elements within sli-content
        const elements = sliContent.querySelectorAll('.shape, .custom-shape, .shape-group, .image-container, .sli-txt-box, .insertable-element, .chart-element, .table-element');
        let highest = 0;
        
        elements.forEach(el => {
            const zIndex = parseInt(el.style.zIndex || 0);
            if (zIndex > highest) {
                highest = zIndex;
            }
        });
        
        // Use the maximum of: highest in DOM, or last assigned value
        // This ensures each call returns an incremental value even in quick succession
        const nextZ = Math.max(highest, this.lastAssignedZIndex) + 1;
        this.lastAssignedZIndex = nextZ;
        return nextZ;
    }

    setupCoreEventListeners() {
        // Menu Tab Switching
        document.querySelectorAll('.menu-item').forEach(item => {
            item.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        // Slide Thumbnails
        document.querySelectorAll('.slide-thumb').forEach(thumb => {
            thumb.addEventListener('click', (e) => this.selectSlide(e.currentTarget.dataset.slide));
        });

        // Keyboard Shortcuts
        document.addEventListener('keydown', (e) => this.handleKeyboardShortcuts(e));

        // Content Changes for History
        const canvas = document.getElementById('canvas');
        // Lazy-init interactive controls ONLY on hover
// ── Single canvas-level mousemove tracker ────────────────────────────────────
// This is the ONLY place that adds/removes .hover. Using mousemove (not
// mouseover/mouseenter) means we always know the CURRENT topmost element via
// document.elementFromPoint — no sibling bleed, no stacking-order ambiguity.
this._hoverEl = null; // currently hovered interactive element (class property for global access)

canvas.addEventListener('mousemove', (e) => {
    const topNode = document.elementFromPoint(e.clientX, e.clientY);
    if (!topNode) return;

    let topEl = topNode.closest('.image-container, .shape, .custom-shape, .shape-group');

    // If elementFromPoint landed on a child handle of _hoverEl, keep hovering it
    if (!topEl && this._hoverEl && this._hoverEl.contains(topNode)) {
        topEl = this._hoverEl;
    }

    // Same element — just ensure lazy-init ran
    if (topEl === this._hoverEl) {
        if (this._hoverEl && this._hoverEl.dataset.interactiveInit !== '1') {
            const type = this._hoverEl.dataset.interactiveType || 'shape';
            const opts = (type === 'text')
                ? { resize:true, move:true, rotate:true, delete:true, minWidth:80, minHeight:30 }
                : { resize:true, move:true, rotate:true, delete:true, minWidth:20, minHeight:20 };
            this.makeElementInteractive(this._hoverEl, opts);
            this._hoverEl.dataset.interactiveInit = '1';
        }
        return;
    }

    // ── Leave previous element ────────────────────────────────────────────────
    if (this._hoverEl && !this._hoverEl.classList.contains('selected')) {
        this._hoverEl.classList.remove('hover');
    }

    // ── Enter new element ─────────────────────────────────────────────────────
    this._hoverEl = topEl || null;

    if (this._hoverEl) {
        if (this._hoverEl.dataset.interactiveInit !== '1') {
            const type = this._hoverEl.dataset.interactiveType || 'shape';
            const opts = (type === 'text')
                ? { resize:true, move:true, rotate:true, delete:true, minWidth:80, minHeight:30 }
                : { resize:true, move:true, rotate:true, delete:true, minWidth:20, minHeight:20 };
            this.makeElementInteractive(this._hoverEl, opts);
            this._hoverEl.dataset.interactiveInit = '1';
        }
        if (!this._hoverEl.classList.contains('selected')) {
            this._hoverEl.classList.add('hover');
        }
    }
});

// When mouse leaves canvas entirely, clear hover
canvas.addEventListener('mouseleave', () => {
    if (this._hoverEl && !this._hoverEl.classList.contains('selected')) {
        this._hoverEl.classList.remove('hover');
    }
    this._hoverEl = null;
});

        // NOTE: canvas 'input' does NOT call saveState — formatting methods save
        // after their own changes. Saving on every keystroke would create too many
        // history entries and break sequential undo.
        // canvas.addEventListener('input', ...) intentionally removed.


        // Track selection changes
        document.addEventListener('selectionchange', () => this.updateFormatButtons());

        // Undo / Redo (Home toolbar + top menu buttons)
        const undoBtn = document.getElementById('undoBtn');
        const redoBtn = document.getElementById('redoBtn');
        const undoBtnTop = document.getElementById('undoBtnTop');
        const redoBtnTop = document.getElementById('redoBtnTop');
        
        // ✅ Respect disabled state in event handlers
        if (undoBtn) undoBtn.addEventListener('click', (e) => {
            if (!e.currentTarget.disabled) this.undo();
        });
        if (redoBtn) redoBtn.addEventListener('click', (e) => {
            if (!e.currentTarget.disabled) this.redo();
        });
        if (undoBtnTop) undoBtnTop.addEventListener('click', (e) => {
            if (!e.currentTarget.disabled) this.undo();
        });
        if (redoBtnTop) redoBtnTop.addEventListener('click', (e) => {
            if (!e.currentTarget.disabled) this.redo();
        });

        // ─────────────────────────────────────────────────────────────────────
        // UNIFIED CLICK HANDLER — single listener, 3-state machine (Canva-style)
        //   State 0 → Click 1 : select shape (drag/move mode)
        //   State 1 → Click 2 : select all text in textbox
        //   State 2 → Click 3+: enter direct edit (contenteditable = true)
        //
        // Uses instance-level state (_txtClickBox, _isShapeSelected, _isTextSelected,
        // _isEditing) to avoid dataset race conditions between two listeners.
        // ─────────────────────────────────────────────────────────────────────

        // Deselect when clicking empty canvas / slide background
        canvas.addEventListener('mousedown', (e) => {
            const clickedElement = e.target;
            const isEmptySpace = clickedElement === canvas ||
                                 clickedElement.classList.contains('sli-slide') ||
                                 clickedElement.classList.contains('sli-content');
            const isInteractiveElement = clickedElement.closest(
                '.shape, .image-container, .sli-txt-box, .insertable-element, .resize-handle, .delete-btn, .rotate-handle, .move-handle'
            );
            if (isEmptySpace || (!isInteractiveElement && clickedElement.closest('.sli-slide'))) {
                this._exitTextEditState();
                this.deselectAllElements(true);
            }
        });

        // Single unified click listener
        canvas.addEventListener('click', (e) => {
            // Ignore toolbar / handle clicks
            if (e.target.classList.contains('resize-handle') ||
                e.target.classList.contains('delete-btn') ||
                e.target.classList.contains('rotate-handle')) return;

            if (e.target.closest('#textToolPanel, #imageToolPanel, button, select')) return;

            // ── IMAGE / NON-TEXT SHAPE ─────────────────────────────────────
            const clickedShape = e.target.closest('.image-container, .shape, .custom-shape, .shape-group');
            if (!clickedShape) return;

            const txtBox = clickedShape.querySelector('.sli-txt-box');

            // Non-text shapes (images, pure draw shapes without sli-txt-box)
            if (!txtBox) {
                if (this._txtClickBox) this._exitTextEditState();
                e.stopPropagation();
                this.selectElement(clickedShape, e.ctrlKey || e.metaKey);
                return;
            }

            // ── TEXT SHAPE ─────────────────────────────────────────────────

            // Clicking a DIFFERENT shape → commit any active edit, go to State 1
            if (clickedShape !== this._txtClickBox) {
                this._exitTextEditState();
                this.deselectAllElements(false);

                // STATE 1 — shape selected, drag enabled
                // Always fully deselect everything first so no other element
                // keeps .selected or .hover from a previous interaction.
                this.deselectAllElements(false);

                this._txtClickBox     = clickedShape;
                this._isShapeSelected = true;
                this._isTextSelected  = false;
                this._isEditing       = false;

                clickedShape.classList.add('selected', 'active-text-box');
                clickedShape.classList.remove('editing-mode', 'text-selected');
                clickedShape.style.cursor = 'move';
                txtBox.setAttribute('contenteditable', 'false');
                txtBox.style.pointerEvents = 'none';
                txtBox.classList.remove('text-all-selected');

                // Init interactive handles (resize, delete, rotate, drag)
                if (clickedShape.dataset.interactiveInit !== '1') {
                    this.makeElementInteractive(clickedShape, {
                        resize: true, move: true, rotate: true, delete: true,
                        minWidth: 80, minHeight: 30
                    });
                    clickedShape.dataset.interactiveInit = '1';
                }

                // Expose active txtBox so home-tab formatting works from State 1
                this._activeTxtBox = txtBox;

                this.setContextualTabFromElement(clickedShape);
                e.stopPropagation();
                return;
            }

            // Clicking the SAME shape — advance state ──────────────────────

            // Already fully in edit mode — let native browser cursor placement work
            if (this._isEditing && this._activeTextEditor && this._activeTextEditor.textBox === txtBox) {
                return;
            }

            // STATE 1 → 2 : select all text
            if (this._isShapeSelected && !this._isTextSelected && !this._isEditing) {
                const sel = window.getSelection();
                if (sel && sel.toString().length > 0) return; // user is drag-selecting

                this._isTextSelected  = true;
                this._isShapeSelected = true;

                clickedShape.classList.add('text-selected');
                clickedShape.classList.remove('editing-mode');
                clickedShape.style.cursor = 'default';

                txtBox.setAttribute('contenteditable', 'true');
                txtBox.style.pointerEvents = 'auto';
                txtBox.style.cursor        = 'default';

                setTimeout(() => {
                    txtBox.focus();
                    try {
                        const range = document.createRange();
                        range.selectNodeContents(txtBox);
                        const s = window.getSelection();
                        s.removeAllRanges();
                        s.addRange(range);
                    } catch (_) {}
                    txtBox.classList.add('text-all-selected');
                }, 10);

                e.stopPropagation();
                e.preventDefault();
                return;
            }

            // STATE 2 → 3 : enter direct edit
            if (this._isTextSelected && !this._isEditing) {
                this._isEditing       = true;
                this._isTextSelected  = false;
                this._isShapeSelected = false;

                txtBox.classList.remove('text-all-selected');
                clickedShape.classList.add('editing-mode');
                clickedShape.classList.remove('text-selected');
                clickedShape.style.cursor = 'text';

                this.startDirectEdit(txtBox, e);

                e.stopPropagation();
                e.preventDefault();
                return;
            }
        }, false);

        // Setup text box interactions for existing text boxes
        this.setupTextBoxInteractions();
    }

    insertShape(type) {
        const canvas = document.getElementById('canvas');
        if (!canvas) return;

        // Find currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            console.error('No active slide found');
            return;
        }

        // Get sli-content container
        const sliContent = currentSlide.querySelector('.sli-content') || currentSlide;

        const shape = document.createElement('div');
        shape.className = 'shape';
        shape.dataset.shapeType = type;
        shape.dataset.originalColor = 'accent1';
        shape.setAttribute('originallummod', 'null');
        shape.setAttribute('originallumoff', 'null');
        shape.setAttribute('originalalpha', 'null');
        shape.dataset.interactiveType = 'shape';

        shape.style.position = 'absolute';
        shape.style.left = '100px';
        shape.style.top = '100px';
        shape.style.width = '200px';
        shape.style.height = '120px';
        shape.style.boxSizing = 'border-box';
        shape.style.cursor = 'move';
        shape.style.display = 'flex';
        shape.style.alignItems = 'center';
        shape.style.overflow = 'visible';
        shape.style.opacity = '1';
        shape.style.transform = 'none';
        shape.style.border = '0px solid transparent';
        
        // Auto-assign z-index to be on top of existing elements
        shape.style.zIndex = this.getHighestZIndex();

        // Generate unique ID for text box
        const textBoxId = Math.random().toString(36).substring(2, 12);

        // Create text box structure for shape
        const textBox = document.createElement('div');
        textBox.className = 'sli-txt-box';
        textBox.id = textBoxId;
        textBox.setAttribute('txtphtype', '');
        textBox.setAttribute('txtphidx', '');
        textBox.setAttribute('txtphsz', '');
        textBox.dataset.name = 'TextBox';
        textBox.setAttribute('contenteditable', 'false');
        textBox.setAttribute('spellcheck', 'false');
        textBox.style.cssText = `
            color: #000000;
            position: absolute;
            font-size: 18px;
            display: flex;
            transform: none;
            opacity: 1;
            text-align: center;
            width: 100%;
            height: 100%;
            top: 0;
            left: 0;
            padding: 8px;
            box-sizing: border-box;
            overflow-wrap: break-word;
            word-break: break-word;
            white-space: pre-wrap;
            line-height: 1.4;
            align-items: center;
            justify-content: center;
            pointer-events: none;
            z-index: 2;
            cursor: text;
        `;

        // Create paragraph and span structure
        const p = document.createElement('p');
        p.style.cssText = `
            text-align: center;
            line-height: 20px;
            margin: 0;
            padding: 0;
            width: 100%;
        `;

        const span = document.createElement('span');
        span.className = 'span-txt';
        span.setAttribute('originalea', '+mn-ea');
        span.setAttribute('origincs', 'Arial');
        span.setAttribute('originsym', '');
        span.setAttribute('latinfont', 'Arial');
        span.setAttribute('originaltxtcolor', '000000');
        span.setAttribute('originallummod', 'null');
        span.setAttribute('originallumoff', 'null');
        span.setAttribute('alpha', '1');
        span.setAttribute('cap', 'none');
        span.style.cssText = `
            font-family: Arial, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif;
            font-size: 18px;
            color: #000000;
            white-space: pre-wrap;
            line-height: 20px;
        `;
        span.textContent = 'Click to edit text';

        p.appendChild(span);
        textBox.appendChild(p);

        // Add text box to shape first
        shape.appendChild(textBox);

        // Shape geometry using clip-path (except hexagon which uses SVG)
        let svgHtml = '';
        switch (type) {
            case 'rectangle':
                shape.id = 'rect';
                shape.dataset.name = 'Rectangle';
                shape.style.background = '#4A90E2';
                shape.style.borderRadius = '0';
                shape.style.setProperty('clip-path', 'polygon(0% 0%, 100% 0%, 100% 100%, 0% 100%)', 'important');
                shape.style.setProperty('-webkit-clip-path', 'polygon(0% 0%, 100% 0%, 100% 100%, 0% 100%)', 'important');
                break;

            case 'circle':
                shape.id = 'ellipse';
                shape.dataset.name = 'Circle';
                shape.style.background = '#4A90E2';
                shape.style.borderRadius = '50%';
                shape.style.setProperty('clip-path', 'ellipse(50% 50% at 50% 50%)', 'important');
                shape.style.setProperty('-webkit-clip-path', 'ellipse(50% 50% at 50% 50%)', 'important');
                break;

            case 'oval':
                shape.id = 'ellipse';
                shape.dataset.name = 'Oval';
                shape.style.background = '#4A90E2';
                shape.style.borderRadius = '50%';
                shape.style.setProperty('clip-path', 'ellipse(50% 50% at 50% 50%)', 'important');
                shape.style.setProperty('-webkit-clip-path', 'ellipse(50% 50% at 50% 50%)', 'important');
                break;

            case 'triangle':
                shape.id = 'triangle';
                shape.dataset.name = 'Triangle';
                shape.style.background = '#4A90E2';
                shape.style.borderRadius = '0';
                shape.style.setProperty('clip-path', 'polygon(50% 0%, 100% 100%, 0% 100%)', 'important');
                shape.style.setProperty('-webkit-clip-path', 'polygon(50% 0%, 100% 100%, 0% 100%)', 'important');
                break;

            case 'pentagon':
                shape.id = 'pentagon';
                shape.dataset.name = 'Pentagon';
                shape.style.background = '#4A90E2';
                shape.style.borderRadius = '0';
                shape.style.setProperty('clip-path', 'polygon(50% 0%, 100% 38%, 81% 100%, 19% 100%, 0% 38%)', 'important');
                shape.style.setProperty('-webkit-clip-path', 'polygon(50% 0%, 100% 38%, 81% 100%, 19% 100%, 0% 38%)', 'important');
                break;

            case 'hexagon':
                shape.id = 'hexagon';
                shape.dataset.name = 'Hexagon';
                shape.style.background = 'transparent';
                shape.style.borderRadius = '0';
                svgHtml = `
            <svg viewBox="0 0 100 100" preserveAspectRatio="none" style="width:100%;height:100%;display:block;position:absolute;top:0;left:0;pointer-events:none;z-index:1;">
                <polygon points="25,0 75,0 100,50 75,100 25,100 0,50"
                    fill="#4A90E2" stroke="#2E5C8A" stroke-width="0"/>
            </svg>`;
                break;
        }

        // Add SVG for hexagon
        if (svgHtml) {
            shape.insertAdjacentHTML('beforeend', svgHtml);
        }


        // Append to sli-content instead of canvas
        sliContent.appendChild(shape);

        this.makeElementInteractive(shape, {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 30,
            minHeight: 30
        });

        this.resetShapeDefaults(shape);
        this.selectElement(shape);

        this.saveState();

        const effectSelect = document.getElementById('shapeEffect');
        if (effectSelect) effectSelect.value = 'none';

    }


    // ============================================
    // CONTEXTUAL TAB (PowerPoint-like)
    // ============================================
    getContextualTabForElement(element) {
        if (!element) return 'home';

        // Images always → image format tab
        if (element.classList.contains('image-container')) return 'imageformat';
        if (element.closest('.image-container')) return 'imageformat';

        // Shape with text-box → home tab (like PowerPoint/Canva)
        // Pure draw shape (no text) → draw/shape-format tab
        if ((element.classList.contains('shape') || element.classList.contains('custom-shape')) || element.classList.contains('shape-group')) {
            return element.querySelector('.sli-txt-box') ? 'home' : 'draw';
        }
        if (element.classList.contains('sli-txt-box')) return 'home';

        // Fallbacks
        const shapeAncestor = element.closest('.shape, .custom-shape, .shape-group');
        if (shapeAncestor) {
            return shapeAncestor.querySelector('.sli-txt-box') ? 'home' : 'draw';
        }
        if (element.closest('.sli-txt-box')) return 'home';

        return 'home';
    }

    setContextualTabFromElement(element) {
        const tab = this.getContextualTabForElement(element);
        this.activateContextualTab(tab);
    }

    // ============================================
    // TAB SWITCHING
    // ============================================
    switchTab(tabName) {
        // Update menu items
        document.querySelectorAll('.menu-item').forEach(item => {
            item.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // Show corresponding toolbar
        document.querySelectorAll('.toolbar').forEach(toolbar => {
            toolbar.classList.remove('active');
        });

        const toolbar = document.querySelector(`.${tabName}-toolbar`);
        if (toolbar) {
            toolbar.classList.add('active');
        }
    }

    /**
     * Contextual (PowerPoint-like) activation: keeps all tabs visible,
     * but automatically activates the right one based on selection.
     */
    activateContextualTab(tabName) {
        const menuItem = document.querySelector(`[data-tab="${tabName}"]`);
        if (!menuItem) return;

        // Avoid redundant work
        if (menuItem.classList.contains('active')) return;
        this.switchTab(tabName);
    }

    getTabForSelectedElement(element) {
        if (!element) return 'home';

        // Image selection
        if (element.classList.contains('image-container') || element.closest('.image-container')) {
            return 'imageformat';
        }

        // Shape selection (single shapes + groups)
        if ((element.classList.contains('shape') || element.classList.contains('custom-shape')) || element.classList.contains('shape-group') ||
            element.closest('.shape, .custom-shape') || element.closest('.shape-group')) {
            return 'draw';
        }

        // Text selection
        if (element.classList.contains('sli-txt-box') || element.closest('.sli-txt-box')) {
            return 'home';
        }

        return 'home';
    }

    // ============================================
    // SLIDE MANAGEMENT
    // ============================================
    /**
     * Normalize z-indexes for all elements in a slide
     * Ensures all elements have proper z-index values and syncs tracker
     */
    normalizeZIndexes(slide) {
        if (!slide) return;

        const sliContent = slide.querySelector('.sli-content') || slide;
        const elements = sliContent.querySelectorAll('.shape, .custom-shape, .shape-group, .image-container, .sli-txt-box, .insertable-element, .chart-element, .table-element');
        
        let highest = 0;
        elements.forEach(el => {
            // If element doesn't have z-index, assign one
            if (!el.style.zIndex || el.style.zIndex === 'auto' || el.style.zIndex === '0') {
                el.style.zIndex = '1';
            }
            
            const zIndex = parseInt(el.style.zIndex || 0);
            if (zIndex > highest) {
                highest = zIndex;
            }
        });
        
        // Sync the tracker with the actual highest in the slide
        this.lastAssignedZIndex = highest;
    }

    selectSlide(slideNumber) {
        this.currentSlide = parseInt(slideNumber);

        // Update active state in sidebar
        document.querySelectorAll('.slide-thumb').forEach(thumb => {
            thumb.classList.remove('active');
        });
        const activeThumb = document.querySelector(`[data-slide="${slideNumber}"]`);
        if (activeThumb) {
            activeThumb.classList.add('active');
        }

        // Show/hide slides in canvas
        const canvas = document.querySelector('.canvas');
        if (canvas) {
            const slides = canvas.querySelectorAll('.sli-slide');
            slides.forEach((slide, index) => {
                if (index === slideNumber - 1) {
                    slide.style.display = 'block';
                    // Normalize z-indexes when slide becomes visible
                    // This also syncs lastAssignedZIndex with the highest in the slide
                    this.normalizeZIndexes(slide);
                } else {
                    slide.style.display = 'none';
                }
            });
        }

        this.updateUI();
        // this.showNotification(`Switched to slide ${slideNumber}`);
    }

    addSlide() {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Create a new blank slide
        const newSlide = document.createElement('div');
        newSlide.className = 'sli-slide';
        newSlide.style.width = '960px';
        newSlide.style.height = '540px';
        newSlide.style.position = 'relative';
        newSlide.style.background = 'white';
        
        // Add to canvas
        canvas.appendChild(newSlide);
        
        this.totalSlides++;

        // ✅ Regenerate all previews
        if (this.generateSlidePreviews) {
            this.generateSlidePreviews();
        }

        // Switch to the new slide
        this.selectSlide(this.totalSlides);

        this.updateUI();
        this.saveState();
        this.showNotification('New slide added');
    }

   deleteSlide(slideNumber) {
    // Don't allow deleting if only one slide left
    if (this.totalSlides <= 1) {
        this.showNotification('Cannot delete the last slide');
        return;
    }

    // Remove the slide from canvas
    const canvas = document.querySelector('.canvas');
    const slides = canvas.querySelectorAll('.sli-slide');
    const slideToDelete = slides[slideNumber - 1];
    
    if (slideToDelete) {
        slideToDelete.remove();
        this.totalSlides--;

        // If deleted slide was active, select another one
        if (this.currentSlide === parseInt(slideNumber)) {
            this.currentSlide = Math.min(this.currentSlide, this.totalSlides);
        } else if (this.currentSlide > parseInt(slideNumber)) {
            this.currentSlide--;
        }

        // ✅ Regenerate all previews
        if (this.generateSlidePreviews) {
            this.generateSlidePreviews();
        }

        // Save state and update UI
        this.saveState();
        this.updateUI();
        this.showNotification(`Slide ${slideNumber} deleted`);
    }
}


    // ============================================
    // HISTORY (UNDO/REDO)
    // ============================================

    resetHistory(initialHtml) {
        this.history = [];
        this.historyIndex = -1;

        if (initialHtml && initialHtml.trim()) {
            this.history.push(initialHtml);
            this.historyIndex = 0;
        }

        this.updateUndoRedoButtons();
    }

resetOnModalClose() {
    // Clear undo/redo history
    this.history = [];
    this.historyIndex = -1;

    // Deselect any selected element
    if (this.selectedElement) {
        this.selectedElement = null;
    }

    // Hide any active toolbars/panels
    const toolbars = document.querySelectorAll('.text-toolbar, .image-toolbar, .shape-toolbar, .element-toolbar');
    toolbars.forEach(tb => { tb.style.display = 'none'; });

    // Reset undo/redo button states
    this.updateUndoRedoButtons();
}

  saveInitialState() {
    const canvas = document.getElementById('canvas');
    if (!canvas) return;
    
    const state = canvas.innerHTML;
    
    // Don't save if canvas is empty or only has whitespace
    if (!state || !state.trim()) {
        return;
    }

    // Initialize history with the first valid state if history is empty
    if (this.history.length === 0) {
        this.history.push(state);
        this.historyIndex = 0;
    } else {
        // console.log('✓ Initial state already saved');
    }
    this.updateUndoRedoButtons();  // Ensure the buttons are updated immediately
}


    // ============================================
    // UNDO / REDO SYSTEM - Simple and Working
    // ============================================
    /**
     * Serialize canvas content to a clean, UI-noise-free string.
     * Strips ALL transient UI state so that selection, hover, editing, and
     * drag-cursor changes never create spurious history entries.
     */
    _getCleanState() {
        const canvas = document.getElementById('canvas');
        if (!canvas) return '';
        const clone = canvas.cloneNode(true);

        // ── Remove handle/overlay DOM elements ───────────────────────────────
        clone.querySelectorAll(
            '.resize-handle, .delete-btn, .rotate-handle, .move-handle, .crop-overlay'
        ).forEach(el => el.remove());

        // ── Strip transient CSS classes ──────────────────────────────────────
        clone.querySelectorAll('*').forEach(el => {
            el.classList.remove(
                'selected', 'hover', 'active-text-box', 'text-selected',
                'editing-mode', 'text-all-selected', 'text-shape',
                'text-all-selected', 'shape-selected'
            );
        });

        // ── Remove data attributes that are purely runtime state ─────────────
        clone.querySelectorAll('[data-interactive-init]').forEach(el =>
            el.removeAttribute('data-interactive-init')
        );
        clone.querySelectorAll('[data-editing]').forEach(el =>
            el.removeAttribute('data-editing')
        );

        // ── Strip transient inline styles on every element ───────────────────
        // These are set during selection / drag / edit and must not be saved.
        const TRANSIENT_STYLES = [
            'cursor', 'pointer-events', 'user-select',
            'min-height', 'flex-shrink', 'align-self'
        ];
        clone.querySelectorAll('*').forEach(el => {
            TRANSIENT_STYLES.forEach(s => el.style.removeProperty(s));
        });
        // Restore image-container overflow to hidden (was set to visible for handles)
        clone.querySelectorAll('.image-container').forEach(el => {
            el.style.overflow = 'hidden';
        });

        // ── Normalize textboxes to resting state ──────────────────────────────
        clone.querySelectorAll('.sli-txt-box').forEach(tb => {
            tb.setAttribute('contenteditable', 'false');
            tb.style.height = '100%';
        });

        return clone.innerHTML;
    }

    saveState() {
        // Re-entrancy guard: restoreState internally manipulates DOM;
        // any saveState triggered during restore must be ignored.
        if (this._restoringState) return;

        const state = this._getCleanState();
        if (!state || !state.trim()) return;

        // Dedup: identical content → no new history entry.
        // Because _getCleanState strips ALL transient UI (selected classes,
        // cursor styles, contenteditable, handles), clicking/selecting a shape
        // produces the exact same string as before → dedup fires → no entry added.
        if (this.history[this.historyIndex] === state) return;

        // Branch: discard any redo states
        this.history = this.history.slice(0, this.historyIndex + 1);
        this.history.push(state);
        this.historyIndex = this.history.length - 1;

        if (this.history.length > this.maxHistory) {
            this.history.shift();
            this.historyIndex--;
        }

        this.updateUndoRedoButtons();

        if (this.refreshSlidePreview) {
            this.refreshSlidePreview(this.currentSlide);
        }
    }

    saveInitialState() {
        if (this.history.length > 0) return; // only on first load
        const state = this._getCleanState();
        if (!state || !state.trim()) return;
        this.history.push(state);
        this.historyIndex = 0;
        this.updateUndoRedoButtons();
    }

    undo() {
        if (this.historyIndex <= 0) return;
        this.historyIndex--;
        this.restoreState();
        this.showNotification('Undo');
    }

    redo() {
        if (this.historyIndex >= this.history.length - 1) return;

        this.historyIndex++;
        this.restoreState();
        this.showNotification('Redo');
    }

    restoreState() {
        const canvas = document.getElementById('canvas');
        if (!canvas || !this.history[this.historyIndex]) return;

        // ── 1. Cleanly exit any active edit session ───────────────────────
        // _commitDirectEdit would call saveState which we don't want here,
        // so just cancel quietly.
        if (this._activeTextEditor) {
            try { this._cancelDirectEdit(); } catch (_) {}
        }

        // ── 2. Reset ALL state-machine flags ─────────────────────────────
        // These point to DOM nodes about to be replaced — must be nulled first.
        this._activeTextEditor  = null;
        this._activeTxtBox      = null;
        this._txtClickBox       = null;
        this._isShapeSelected   = false;
        this._isTextSelected    = false;
        this._isEditing         = false;
        this.activeTextBox      = null;
        // Reset home-tab's saved selection range — it points to destroyed nodes
        if (window._homeTab) {
            window._homeTab._savedSelectionRange = null;
        }

        // ── 3. Write clean HTML (guard against saveState re-entry) ─────────
        this._restoringState = true;
        canvas.innerHTML = this.history[this.historyIndex];

        // ── 4. Paranoia strip — remove any leftover UI noise ─────────────
        canvas.querySelectorAll(
            '.resize-handle, .delete-btn, .rotate-handle, .move-handle, .crop-overlay'
        ).forEach(el => el.remove());
        canvas.querySelectorAll('*').forEach(el => {
            el.classList.remove(
                'selected', 'hover', 'active-text-box', 'text-selected',
                'editing-mode', 'text-all-selected', 'text-shape', 'shape-selected'
            );
        });
        // Also reset wrapper-level state attributes
        canvas.querySelectorAll('.shape, .custom-shape').forEach(el => {
            el.style.removeProperty('cursor');
            el.style.removeProperty('user-select');
        });
        canvas.querySelectorAll('[data-interactive-init]').forEach(el =>
            el.removeAttribute('data-interactive-init')
        );
        canvas.querySelectorAll('.sli-txt-box').forEach(tb => {
            tb.setAttribute('contenteditable', 'false');
            tb.style.pointerEvents = 'none';
        });

        // ── 5. Re-normalize shapes ────────────────────────────────────────
        this.normalizeExistingShapes();
        this.setupTextBoxInteractions();

        // ── 6. Eagerly make ALL elements interactive ──────────────────────
        // Lazy mouseover-init won't work until the user hovers; force it now
        // so handles, drag, rotate, delete all work immediately after restore.
        canvas.querySelectorAll('.shape, .custom-shape, .image-container, .shape-group').forEach(el => {
            const type = el.dataset.interactiveType || 'shape';
            const opts = (type === 'text')
                ? { resize:true, move:true, rotate:true, delete:true, minWidth:80, minHeight:30 }
                : { resize:true, move:true, rotate:true, delete:true, minWidth:20, minHeight:20 };
            this.makeElementInteractive(el, opts);
            el.dataset.interactiveInit = '1';
        });

        // ── 7. Update UI ──────────────────────────────────────────────────
        // ── 8. Release re-entrancy guard ─────────────────────────────────
        this._restoringState = false;

        this.updateUndoRedoButtons();

        if (this.refreshSlidePreview) {
            this.refreshSlidePreview(this.currentSlide);
        }
    }

    updateUndoRedoButtons() {
        const canUndo = this.historyIndex > 0;
        const canRedo = this.historyIndex < this.history.length - 1;
        [
            document.getElementById('undoBtn'),
            document.getElementById('undoBtnTop')
        ].forEach(btn => {
            if (!btn) return;
            btn.disabled = !canUndo;
            btn.style.opacity = canUndo ? '1' : '0.4';
            btn.style.cursor = canUndo ? 'pointer' : 'not-allowed';
            btn.style.pointerEvents = canUndo ? 'auto' : 'none';
        });
        [
            document.getElementById('redoBtn'),
            document.getElementById('redoBtnTop')
        ].forEach(btn => {
            if (!btn) return;
            btn.disabled = !canRedo;
            btn.style.opacity = canRedo ? '1' : '0.4';
            btn.style.cursor = canRedo ? 'pointer' : 'not-allowed';
            btn.style.pointerEvents = canRedo ? 'auto' : 'none';
        });
    }


    // ============================================
    // TEXT BOX MANAGEMENT
    // ============================================
    setupTextBoxInteractions() {
        // Make ALL existing slide elements interactive based on your structure.
        // Important: for text boxes, attach handles to the wrapper .shape (because many shapes have overflow:hidden)
        // while keeping actual editing inside .sli-txt-box and .span-txt.
        const textBoxes = document.querySelectorAll('.sli-txt-box');
        textBoxes.forEach(tb => {
            tb.setAttribute('spellcheck', 'false');
            tb.setAttribute('contenteditable', 'false');

            // Remove any old handles (from older DOM)
            tb.querySelectorAll('.resize-handle, .delete-btn, .move-handle, .rotate-handle').forEach(h => h.remove());

            const wrapper = tb.closest('.shape, .custom-shape') || tb;
            // Mark wrapper as a text-shape for easier detection
            wrapper.classList.add('text-shape');
            wrapper.dataset.textboxId = tb.id || '';

            // Clear old handles on wrapper too
            wrapper.querySelectorAll(':scope > .resize-handle, :scope > .delete-btn, :scope > .move-handle, :scope > .rotate-handle')
                .forEach(h => h.remove());

            // Make wrapper interactive so handles are not clipped
            wrapper.dataset.interactiveType = 'text';

            // ── Guard: only attach mousedown ONCE per textbox instance ──────
            // After undo/redo, canvas.innerHTML is replaced so ALL nodes are
            // brand-new — _tbMousedownBound will never be set on them.
            // The check prevents double-attachment if setupTextBoxInteractions
            // is called twice on the same live DOM (e.g. on initial load).
            if (!tb._tbMousedownBound) {
                const tbMousedown = (e) => {
                    if (e.target.classList.contains('resize-handle') ||
                        e.target.classList.contains('delete-btn') ||
                        e.target.classList.contains('move-handle') ||
                        e.target.classList.contains('rotate-handle')) return;

                    // State 3 — already editing: let browser place cursor, stop drag
                    if (this._activeTextEditor && this._activeTextEditor.textBox === tb) {
                        e.stopPropagation();
                        return;
                    }

                    // States 1 & 2 — stop propagation so drag doesn't swallow click.
                    // The unified canvas click handler owns all state transitions.
                    if (this._txtClickBox === tb.closest('.shape, .custom-shape')) {
                        e.stopPropagation();
                    }
                };
                tb.addEventListener('mousedown', tbMousedown);
                tb._tbMousedownBound = true;
            }

            // dblclick handled by 3-click state machine
        });

        // Shapes — only reset handles on elements NOT yet interactively init'd.
        // Elements with interactiveInit='1' already have working handles; removing
        // them here would break them until next hover triggers re-init.
        document.querySelectorAll('.shape, .custom-shape').forEach(shape => {
            if (shape.classList.contains('text-shape')) return;
            if (shape.dataset.interactiveInit === '1') return; // already init'd, leave alone
            shape.querySelectorAll(':scope > .resize-handle, :scope > .delete-btn, :scope > .move-handle, :scope > .rotate-handle')
                .forEach(h => h.remove());
            shape.dataset.interactiveType = 'shape';
        });

        // Images — same guard
        document.querySelectorAll('.image-container').forEach(img => {
            if (img.dataset.interactiveInit === '1') return; // already init'd, leave alone
            img.querySelectorAll(':scope > .resize-handle, :scope > .delete-btn, :scope > .move-handle, :scope > .rotate-handle')
                .forEach(h => h.remove());
            img.dataset.interactiveType = 'image';
        });
    }
    makeTextBoxInteractive(textBox) {
        const wrapper = textBox.closest('.shape, .custom-shape') || textBox;
        wrapper.classList.add('text-shape');
        wrapper.dataset.textboxId = textBox.id || '';

        // Ensure text box floats on top
        if (!wrapper.style.zIndex || parseInt(wrapper.style.zIndex) < 1000) {
            wrapper.style.zIndex = '1000';
        }
        wrapper.style.overflow = 'visible';

        // Use universal method from core with text box specific settings (on wrapper)
        return this.makeElementInteractive(wrapper, {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 80,
            minHeight: 30
        });
    }


    // Convenience methods for text boxes
    // Convenience methods for text boxes
    selectTextBox(textBox, isMultiSelect = false) {
        const wrapper = textBox.closest('.shape, .custom-shape') || textBox;

        if (!isMultiSelect) {
            this.deselectAllElements(false);
        } else if (wrapper.classList.contains('selected')) {
            wrapper.classList.remove('selected');
            textBox.classList.remove('selected');
            return;
        }

        wrapper.classList.add('selected');
        textBox.classList.add('selected');
        this.activeTextBox = textBox;

        // Keep a DOM-readable pointer for toolbar logic
        if (textBox && textBox.id) document.body.dataset.activeTextBoxId = textBox.id;

        // IMPORTANT: Do NOT set contenteditable here.
        // Editing is handled via an overlay textarea opened on double-click.
        this.setContextualTabFromElement(textBox);
    }


    
    // ============================================
    // TEXT AUTO-GROW HELPERS (stable measurement)
    // ============================================

    // ─────────────────────────────
    // TEXT EDITING: OVERLAY TEXTAREA (CANVA-STYLE)
    // ─────────────────────────────
    
    // ─────────────────────────────
    // TEXT EDITING: RICH OVERLAY (CONTENTEDITABLE) — supports selection + formatting
    // ─────────────────────────────

    // ─────────────────────────────────────────────────────────────────────────
    // DIRECT EDIT — edits sli-txt-box in-place (contenteditable), no overlay div.
    // DOM stays: shape > sli-txt-box > p > span.span-txt
    // ─────────────────────────────────────────────────────────────────────────

    // ─────────────────────────────────────────────────────────────────────────
    // Helper: commit/exit any active edit and reset 3-state instance vars
    // ─────────────────────────────────────────────────────────────────────────
    _exitTextEditState() {
        if (this._activeTextEditor) {
            this._commitDirectEdit();
        }
        if (this._txtClickBox) {
            const box = this._txtClickBox;
            box.classList.remove('editing-mode', 'text-selected', 'active-text-box');
            box.style.cursor   = '';
            box.style.overflow = box.dataset._preEditWrapperOverflow || 'visible';
            delete box.dataset._preEditWrapperOverflow;
            const tb = box.querySelector('.sli-txt-box');
            if (tb) {
                tb.setAttribute('contenteditable', 'false');
                tb.style.pointerEvents = '';
                tb.style.cursor     = '';
                tb.style.flexShrink = '';
                tb.style.alignSelf  = '';
                tb.classList.remove('text-all-selected');
                tb.style.height    = '100%';
                tb.style.minHeight = '';
                tb.style.overflow  = '';
            }
        }
        this._txtClickBox     = null;
        this._activeTxtBox    = null;
        this._isShapeSelected = false;
        this._isTextSelected  = false;
        this._isEditing       = false;
    }

    startDirectEdit(textBox, clickEvent, skipFocusRaf = false) {
        if (this._activeTextEditor && this._activeTextEditor.textBox !== textBox) {
            this._commitDirectEdit();
        }
        if (this._activeTextEditor && this._activeTextEditor.textBox === textBox) return;

        const wrapper = textBox.closest('.shape, .custom-shape') || textBox;
        wrapper.dataset._editing = '1';

        // ── LOCK WRAPPER WIDTH, allow height to grow only with text wrap ──────────
        const lockedW = parseFloat(wrapper.style.width) || wrapper.getBoundingClientRect().width || wrapper.offsetWidth;
        const lockedH = parseFloat(wrapper.style.height) || wrapper.getBoundingClientRect().height || wrapper.offsetHeight;
        wrapper.style.width  = lockedW + 'px';
        wrapper.style.height = lockedH + 'px';   // start at original height
        wrapper.dataset._editBaseHeight = String(lockedH);
        // Keep overflow:visible on wrapper so text/handles aren't clipped
        wrapper.dataset._preEditWrapperOverflow = wrapper.style.overflow || '';

        // Enable contenteditable directly on sli-txt-box — no extra div
        textBox.setAttribute('contenteditable', 'true');
        textBox.setAttribute('spellcheck', 'false');
        textBox.style.pointerEvents = 'auto';
        textBox.style.cursor        = 'text';
        textBox.style.caretColor    = 'black';
        // Remove fixed height constraints — let the textbox size to its content.
        // The wrapper height is controlled explicitly by the input handler below.
        textBox.style.height        = 'auto';
        textBox.style.minHeight     = '0';
        textBox.style.overflow      = 'visible';
        textBox.style.flexShrink    = '0';
        textBox.style.alignSelf     = 'flex-start';

        // Measurer kept for commit-time height calculation but NOT used to resize during input
        const cs = window.getComputedStyle(textBox);
        const measurer = document.createElement('div');
        Object.assign(measurer.style, {
            position: 'fixed', top: '-99999px', left: '-99999px',
            width: lockedW + 'px', height: 'auto',
            padding: cs.padding, fontFamily: cs.fontFamily, fontSize: cs.fontSize,
            fontWeight: cs.fontWeight, fontStyle: cs.fontStyle,
            lineHeight: cs.lineHeight, letterSpacing: cs.letterSpacing,
            whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowWrap: 'break-word',
            boxSizing: 'border-box', overflow: 'hidden',
            visibility: 'hidden', pointerEvents: 'none',
        });
        document.body.appendChild(measurer);

        // autoSize: grow wrapper height when text wraps, never shrink below original.
        // Since wrapper is position:absolute, growing it never shifts other elements.
        const autoSize = () => {
            // Use a hidden off-screen clone to measure true content height
            const clone = textBox.cloneNode(true);
            clone.style.cssText = [
                'position:fixed', 'top:-99999px', 'left:-99999px',
                'visibility:hidden', 'pointer-events:none',
                `width:${lockedW}px`, 'height:auto', 'min-height:0',
                'overflow:visible', 'white-space:pre-wrap',
                'word-break:break-word', 'overflow-wrap:break-word',
                'box-sizing:border-box', 'display:block'
            ].join(';');
            document.body.appendChild(clone);
            const needed = Math.ceil(clone.scrollHeight || clone.offsetHeight || 0);
            clone.remove();
            // Allow grow AND shrink — always clamp to at least lockedH
            const target = Math.max(lockedH, needed);
            wrapper.style.height = target + 'px';
            textBox.style.height = target + 'px';
        };

        // Save selection on toolbar mousedown (fires before blur, so range is still alive)
        let _savedRange = null;
        const onToolbarMousedown = (evt) => {
            const onToolbar = evt.target.closest(
                '.toolbar, .home-toolbar, [data-tab], #textToolPanel, button, select, input, label, .tool-btn'
            );
            if (!onToolbar) return;
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0) {
                try { _savedRange = sel.getRangeAt(0).cloneRange(); } catch (_) {}
            }
            window._editorSavedRange = _savedRange;
            window._editorOverlayEl  = textBox;
        };
        document.addEventListener('mousedown', onToolbarMousedown, true);

        const onKeyDown = (e) => {
            if (e.key === 'Escape') { e.preventDefault(); this._cancelDirectEdit(); return; }
            if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); this._commitDirectEdit(); return; }
        };

        // ── STRUCTURE GUARD ──────────────────────────────────────────────────────
        // Snapshot span + paragraph attributes NOW (before any editing mutates the DOM).
        // These are used to rebuild a clean structure if the browser tries to inject
        // <font> tags when the user does Ctrl+A → type/paste/backspace.

        const _snapFirstSpan = () => textBox.querySelector('.span-txt');
        const _snapFirstP    = () => textBox.querySelector('p');

        // Capture the "template" span attrs at startDirectEdit time so even if the
        // span is destroyed mid-edit we can still restore its styles.
        const _snapshotSpanAttrs = (() => {
            const src = _snapFirstSpan();
            if (!src) return { className: 'span-txt', attrs: [] };
            return {
                className : src.className,
                attrs     : Array.from(src.attributes)
                                 .filter(a => a.name !== 'class')
                                 .map(a => ({ name: a.name, value: a.value }))
            };
        })();

        const _snapshotPStyle = (() => {
            const p = _snapFirstP();
            return p ? (p.getAttribute('style') || '') : '';
        })();

        // Build a fresh span carrying the snapshotted attributes + styles
        const _buildCleanSpan = (text = '') => {
            const span = document.createElement('span');
            span.className = _snapshotSpanAttrs.className;
            _snapshotSpanAttrs.attrs.forEach(({ name, value }) => span.setAttribute(name, value));
            if (text) span.textContent = text;
            return span;
        };

        const _buildCleanP = (spanEl) => {
            const p = document.createElement('p');
            if (_snapshotPStyle) p.setAttribute('style', _snapshotPStyle);
            p.appendChild(spanEl);
            return p;
        };

        // Robust "all text selected" check — works for Ctrl+A, triple-click, and
        // programmatic selectNodeContents by comparing the selection's plain-text
        // length against the textbox's plain-text length.
        const _isAllSelected = () => {
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return false;

            // Quick check: selection string vs textBox text content
            const selText = sel.toString();
            const tbText  = textBox.textContent;

            // If lengths match it's almost certainly all-selected
            if (selText.length > 0 && selText.length >= tbText.length) return true;

            // Fallback: try range boundary comparison
            try {
                const range   = sel.getRangeAt(0);
                const tbRange = document.createRange();
                tbRange.selectNodeContents(textBox);
                const cmpStart = range.compareBoundaryPoints(Range.START_TO_START, tbRange);
                const cmpEnd   = range.compareBoundaryPoints(Range.END_TO_END,     tbRange);
                return cmpStart <= 0 && cmpEnd >= 0;
            } catch (_) {
                return false;
            }
        };

        // Intercept text insertion when ALL content is selected — prevents <font> injection
        const onBeforeInput = (e) => {
            const type = e.inputType;
            if (!['insertText', 'insertFromPaste', 'insertFromDrop'].includes(type)) return;

            const sel = window.getSelection();
            const anchorNode = sel && sel.anchorNode;

            // Case 1: All text selected → replace-all scenario
            const allSelected = _isAllSelected();

            // Case 2: Caret is NOT inside a .span-txt (e.g. after Backspace cleared the box
            //         and the caret ended up directly in the <p> or in the textBox root)
            const caretInSpan = anchorNode && (
                anchorNode.nodeType === 3
                    ? anchorNode.parentElement && anchorNode.parentElement.closest('.span-txt')
                    : anchorNode.closest && anchorNode.closest('.span-txt')
            );
            const caretLost = anchorNode && textBox.contains(anchorNode) && !caretInSpan;

            if (!allSelected && !caretLost) return; // normal mid-text edit, let browser handle

            e.preventDefault();

            const insertedText = e.data ||
                (e.dataTransfer && e.dataTransfer.getData('text/plain')) || '';

            if (allSelected) {
                // Replace-all: build fresh structure with new text
                const newSpan = _buildCleanSpan(insertedText);
                const newP    = _buildCleanP(newSpan);
                textBox.innerHTML = '';
                textBox.appendChild(newP);

                requestAnimationFrame(() => {
                    try {
                        const r = document.createRange();
                        const textNode = newSpan.firstChild;
                        if (textNode) {
                            r.setStart(textNode, textNode.length);
                        } else {
                            r.setStart(newSpan, 0);
                        }
                        r.collapse(true);
                        const s = window.getSelection();
                        s.removeAllRanges();
                        s.addRange(r);
                    } catch (_) {}
                    autoSize();
                });
            } else {
                // Caret lost (not inside span) — find the nearest span or the empty one we created,
                // insert text there, and redirect the caret properly.
                let targetSpan = textBox.querySelector('.span-txt');
                if (!targetSpan) {
                    targetSpan = _buildCleanSpan('');
                    const newP = _buildCleanP(targetSpan);
                    textBox.innerHTML = '';
                    textBox.appendChild(newP);
                }
                // Append the new character to span's text content
                const existingText = targetSpan.textContent || '';
                targetSpan.textContent = existingText + insertedText;

                requestAnimationFrame(() => {
                    try {
                        const r = document.createRange();
                        const textNode = targetSpan.firstChild;
                        if (textNode) {
                            r.setStart(textNode, textNode.length);
                        } else {
                            r.setStart(targetSpan, 0);
                        }
                        r.collapse(true);
                        const s = window.getSelection();
                        s.removeAllRanges();
                        s.addRange(r);
                    } catch (_) {}
                    autoSize();
                });
            }
        };

        // Shared: clear all content but keep p > span.span-txt structure intact
        const clearAllPreserveStructure = () => {
            const newSpan = _buildCleanSpan('');
            const newP    = _buildCleanP(newSpan);
            textBox.innerHTML = '';
            textBox.appendChild(newP);
            requestAnimationFrame(() => {
                try {
                    const r = document.createRange();
                    r.setStart(newSpan, 0);
                    r.collapse(true);
                    const s = window.getSelection();
                    s.removeAllRanges();
                    s.addRange(r);
                } catch (_) {}
                autoSize();
            });
        };

        // Guard Backspace/Delete when everything is selected — clear content but keep structure
        const onKeyDownStructure = (e) => {
            if ((e.key === 'Backspace' || e.key === 'Delete') && _isAllSelected()) {
                e.preventDefault();
                clearAllPreserveStructure();
            }
        };
        // ── END STRUCTURE GUARD ──────────────────────────────────────────────────

        const onBlur = (e) => {
            const related = e.relatedTarget;

            // Is focus going to a toolbar element?
            const isToolbar = related && related.closest(
                '.toolbar, .home-toolbar, [data-tab], #textToolPanel, button, select, input, label, .tool-btn'
            );

            if (isToolbar) {
                const isSelectEl = related.tagName === 'SELECT' || !!related.closest('select');
                if (isSelectEl) {
                    // <select> dropdown: save range but do NOT refocus.
                    // Calling textBox.focus() here would close the native dropdown.
                    // changeFontFamily/changeFontSize will restore focus after change fires.
                    const sel = window.getSelection();
                    if (sel && sel.rangeCount > 0) {
                        try { _savedRange = sel.getRangeAt(0).cloneRange(); } catch (_) {}
                    }
                } else {
                    // Button/input: refocus and restore selection after tick
                    setTimeout(() => {
                        textBox.focus();
                        if (_savedRange) {
                            try { const s = window.getSelection(); s.removeAllRanges(); s.addRange(_savedRange); } catch (_) {}
                        }
                    }, 0);
                }
                return; // never commit when toolbar element took focus
            }

            // relatedTarget is null when the OS-level select dropdown opens (2nd blur).
            // If home-tab saved a selection range during mousedown pointing inside this
            // textbox, a select interaction is in progress — do not commit.
            const htRange = window._homeTab && window._homeTab._savedSelectionRange;
            if (htRange) {
                try {
                    if (textBox.contains(htRange.commonAncestorContainer)) return;
                } catch (_) {}
            }

            this._commitDirectEdit();
        };

        const stopProp  = (e) => e.stopPropagation();
        const stopClick = (e) => e.stopPropagation();
        textBox.addEventListener('mousedown',   stopProp);
        textBox.addEventListener('click',       stopClick);
        textBox.addEventListener('input',       autoSize);
        textBox.addEventListener('keydown',     onKeyDown);
        textBox.addEventListener('keydown',     onKeyDownStructure);
        textBox.addEventListener('beforeinput', onBeforeInput);
        textBox.addEventListener('blur',        onBlur);

        this._activeTextEditor = { textBox, wrapper, measurer, onToolbarMousedown, autoSize, onKeyDown, onKeyDownStructure, onBeforeInput, onBlur, stopProp, stopClick, clearAllPreserveStructure };

        window._editorOverlayEl  = textBox;
        window._editorSavedRange = null;

        requestAnimationFrame(() => {
            autoSize();
            if (!skipFocusRaf) {
                textBox.focus();
            }
            try {
                let range = null;
                // Place cursor at exact click position if event available
                if (clickEvent) {
                    if (document.caretRangeFromPoint) {
                        range = document.caretRangeFromPoint(clickEvent.clientX, clickEvent.clientY);
                    } else if (document.caretPositionFromPoint) {
                        const pos = document.caretPositionFromPoint(clickEvent.clientX, clickEvent.clientY);
                        if (pos) {
                            range = document.createRange();
                            range.setStart(pos.offsetNode, pos.offset);
                        }
                    }
                }
                // Fallback: place cursor at end
                if (!range) {
                    range = document.createRange();
                    range.selectNodeContents(textBox);
                    range.collapse(false);
                }
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            } catch (_) {}
        });
    }

    _commitDirectEdit() {
        const s = this._activeTextEditor;
        if (!s) return;
        const { textBox, wrapper, measurer, onToolbarMousedown, autoSize, onKeyDown, onKeyDownStructure, onBeforeInput, onBlur, stopProp, stopClick } = s;

        if (measurer && measurer.parentNode) measurer.remove();
        document.removeEventListener('mousedown', onToolbarMousedown, true);
        textBox.removeEventListener('mousedown',   stopProp);
        textBox.removeEventListener('click',       stopClick);
        textBox.removeEventListener('input',       autoSize);
        textBox.removeEventListener('keydown',     onKeyDown);
        textBox.removeEventListener('keydown',     onKeyDownStructure);
        textBox.removeEventListener('beforeinput', onBeforeInput);
        textBox.removeEventListener('blur',        onBlur);

        window._editorSavedRange = null;
        window._editorOverlayEl  = null;

        this._normalizeTbInPlace(textBox);

        textBox.setAttribute('contenteditable', 'false');
        textBox.style.pointerEvents = 'none';
        textBox.style.cursor        = '';
        textBox.style.overflow      = '';
        textBox.style.height        = '100%';
        textBox.style.minHeight     = '';
        textBox.style.flexShrink    = '';
        textBox.style.alignSelf     = '';

        // Restore wrapper overflow
        wrapper.style.overflow   = wrapper.dataset._preEditWrapperOverflow || 'visible';
        delete wrapper.dataset._preEditWrapperOverflow;
        wrapper.dataset._editing = '0';
        this._activeTextEditor   = null;

        wrapper.classList.remove('editing-mode', 'text-selected');
        wrapper.classList.add('active-text-box', 'selected');
        wrapper.style.cursor     = 'move';
        wrapper.style.userSelect = 'none';
        wrapper.removeAttribute('contenteditable');

        this._isEditing       = false;
        this._isTextSelected  = false;
        this._isShapeSelected = true;

        // After committing, go back to State 1 (shape selected / drag mode)
        this._txtClickBox = wrapper;
        // NOTE: do NOT call saveState() here. Formatting methods already save AFTER
        // applying their change. Saving here would add a duplicate entry every time
        // the user clicks away from a textbox, breaking sequential undo.
    }

    _cancelDirectEdit() {
        const s = this._activeTextEditor;
        if (!s) return;
        const { textBox, wrapper, measurer, onToolbarMousedown, autoSize, onKeyDown, onKeyDownStructure, onBeforeInput, onBlur, stopProp, stopClick } = s;

        if (measurer && measurer.parentNode) measurer.remove();
        document.removeEventListener('mousedown', onToolbarMousedown, true);
        textBox.removeEventListener('mousedown',   stopProp);
        textBox.removeEventListener('click',       stopClick);
        textBox.removeEventListener('input',       autoSize);
        textBox.removeEventListener('keydown',     onKeyDown);
        textBox.removeEventListener('keydown',     onKeyDownStructure);
        textBox.removeEventListener('beforeinput', onBeforeInput);
        textBox.removeEventListener('blur',        onBlur);

        window._editorSavedRange = null;
        window._editorOverlayEl  = null;

        textBox.setAttribute('contenteditable', 'false');
        textBox.style.pointerEvents = 'none';
        textBox.style.cursor        = '';
        textBox.style.overflow      = '';
        textBox.style.height        = '100%';
        textBox.style.minHeight     = '';
        textBox.style.flexShrink    = '';
        textBox.style.alignSelf     = '';

        // Restore wrapper overflow and height
        wrapper.style.overflow   = wrapper.dataset._preEditWrapperOverflow || 'visible';
        delete wrapper.dataset._preEditWrapperOverflow;
        wrapper.dataset._editing = '0';
        this._activeTextEditor   = null;

        wrapper.classList.remove('editing-mode', 'text-selected');
        wrapper.classList.add('active-text-box', 'selected');
        wrapper.style.cursor     = 'move';
        wrapper.style.userSelect = 'none';
        wrapper.removeAttribute('contenteditable');

        // Restore original locked height on cancel
        const baseH = parseFloat(wrapper.dataset._editBaseHeight || '0') || 0;
        if (baseH) { wrapper.style.height = baseH + 'px'; }

        // Reset to State 1 (shape selected)
        this._isEditing       = false;
        this._isTextSelected  = false;
        this._isShapeSelected = true;
        this._txtClickBox     = wrapper;
    }

    _normalizeTbInPlace(tb) {
        const baseSpan      = tb.querySelector('.span-txt');
        const baseSpanStyle = baseSpan ? baseSpan.getAttribute('style') : '';
        const baseAttrs     = baseSpan ? Array.from(baseSpan.attributes) : [];

        const tmp = document.createElement('div');
        tmp.innerHTML = tb.innerHTML;

        // ── Strip <font> tags (browser inserts them when contenteditable replaces
        //    all content) — unwrap them, keeping their text content.
        tmp.querySelectorAll('font').forEach(font => {
            const frag = document.createDocumentFragment();
            while (font.firstChild) frag.appendChild(font.firstChild);
            font.replaceWith(frag);
        });

        tmp.querySelectorAll('div').forEach(d => {
            const p = document.createElement('p');
            p.innerHTML = d.innerHTML;
            d.replaceWith(p);
        });

        if (!tmp.querySelector('p')) {
            const p = document.createElement('p');
            p.innerHTML = tmp.innerHTML;
            tmp.innerHTML = '';
            tmp.appendChild(p);
        }

        tmp.querySelectorAll('p').forEach(p => {
            p.style.margin = '0';
            if (!p.textContent.trim()) p.innerHTML = '&nbsp;';
            if (!p.querySelector('.span-txt')) {
                const span = document.createElement('span');
                span.className = 'span-txt';
                if (baseSpanStyle) span.setAttribute('style', baseSpanStyle);
                baseAttrs.forEach(a => {
                    if (['id','class','style'].includes(a.name)) return;
                    span.setAttribute(a.name, a.value);
                });
                span.innerHTML = p.innerHTML || '&nbsp;';
                p.innerHTML = '';
                p.appendChild(span);
            } else {
                p.querySelectorAll('.span-txt').forEach(s => {
                    if (baseSpanStyle && !s.getAttribute('style')) s.setAttribute('style', baseSpanStyle);
                    s.style.whiteSpace = 'pre-wrap';
                });
            }
        });

        tb.innerHTML = tmp.innerHTML;
    }

    // Shims so any remaining code using old names still works
    startTextEditOverlay(tb)   { return this.startDirectEdit(tb); }
    _commitTextEditOverlay()   { return this._commitDirectEdit(); }
    _cancelTextEditOverlay()   { return this._cancelDirectEdit(); }

    _getTextMeasurer() {
        if (this._textMeasurerEl) return this._textMeasurerEl;

        const el = document.createElement('div');
        el.setAttribute('aria-hidden', 'true');
        el.style.position = 'fixed';
        el.style.left = '-100000px';
        el.style.top = '0';
        el.style.visibility = 'hidden';
        el.style.pointerEvents = 'none';
        el.style.whiteSpace = 'normal';
        el.style.wordBreak = 'break-word';
        el.style.overflowWrap = 'break-word';
        el.style.padding = '0';
        el.style.margin = '0';
        el.style.boxSizing = 'border-box';
        el.style.height = 'auto';
        el.style.maxHeight = 'none';
        el.style.border = '0';
        document.body.appendChild(el);

        this._textMeasurerEl = el;
        return el;
    }

    _measureTextBoxContentHeight(textBox, widthPx) {
        const meas = this._getTextMeasurer();

        // Sync width
        meas.style.width = `${Math.max(1, Math.floor(widthPx || 1))}px`;

        // Copy key typography styles from textbox (spans inside will override as needed)
        const cs = window.getComputedStyle(textBox);
        meas.style.fontFamily = cs.fontFamily;
        meas.style.fontSize = cs.fontSize;
        meas.style.fontWeight = cs.fontWeight;
        meas.style.fontStyle = cs.fontStyle;
        meas.style.letterSpacing = cs.letterSpacing;
        meas.style.lineHeight = cs.lineHeight;
        meas.style.textAlign = cs.textAlign;

        // Important: use the same HTML so inline spans/paragraphs match
        meas.innerHTML = textBox.innerHTML || '';

        // Force <p> margins to 0 in measurer to match your template conventions
        meas.querySelectorAll('p').forEach(p => {
            p.style.margin = '0';
        });

        // scrollHeight is stable because measurer has no constrained height
        const h = Math.ceil(meas.scrollHeight || 0);

        // Fallback to at least 1 line
        const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) || 16;
        return Math.max(Math.ceil(lh), h);
    }

    _applyTextAutoGrow(textBox) {
        if (!textBox || !document.body.contains(textBox)) return;

        const wrapper = textBox.closest('.shape, .custom-shape') || textBox.parentElement;
        if (!wrapper) return;

        const base = parseFloat(wrapper.dataset._editBaseHeight || '0') || (wrapper.getBoundingClientRect().height || wrapper.offsetHeight || 0);

        // Use wrapper content box width (more accurate than offsetWidth when borders)
        const width = wrapper.clientWidth || parseFloat(wrapper.style.width) || textBox.clientWidth || 1;

        const needed = this._measureTextBoxContentHeight(textBox, width);

        // Only grow; never shrink while typing (Canva-like)
        const target = Math.max(base, needed);

        const current = wrapper.getBoundingClientRect().height || wrapper.offsetHeight || 0;

        // Threshold prevents 1px jitter / drift due to rounding
        if (target > current + 2) {
            wrapper.style.height = `${target}px`;
        }
    }

    _scheduleTextAutoGrow(textBox) {
        if (this._textAutoGrowRaf) cancelAnimationFrame(this._textAutoGrowRaf);
        this._textAutoGrowRaf = requestAnimationFrame(() => {
            this._textAutoGrowRaf = null;
            this._applyTextAutoGrow(textBox);
        });
    }

deselectAllTextBoxes() {
        this.deselectAllElements();
    }

    deleteTextBox(textBox) {
        this.deleteElement(textBox);
    }

    // ============================================
    // TEXT FORMATTING HELPERS
    // ============================================
    updateFormatButtons() {
        // Update button states based on current selection
        const selection = window.getSelection();
        if (!selection.rangeCount) return;

        const parentElement = selection.anchorNode.parentElement;

        // Update bold button
        const isBold = document.queryCommandState('bold') ||
            window.getComputedStyle(parentElement).fontWeight === 'bold' ||
            window.getComputedStyle(parentElement).fontWeight >= 700;
        const boldBtn = document.getElementById('boldBtn');
        if (boldBtn) boldBtn.classList.toggle('active', isBold);

        // Update italic button
        const isItalic = document.queryCommandState('italic') ||
            window.getComputedStyle(parentElement).fontStyle === 'italic';
        const italicBtn = document.getElementById('italicBtn');
        if (italicBtn) italicBtn.classList.toggle('active', isItalic);

        // Update underline button
        const isUnderline = document.queryCommandState('underline');
        const underlineBtn = document.getElementById('underlineBtn');
        if (underlineBtn) underlineBtn.classList.toggle('active', isUnderline);

        // Update undo/redo button states
    }

    updateUndoRedoButtons() {
    const undoBtn = document.getElementById('undoBtn');
    const redoBtn = document.getElementById('redoBtn');
    const undoBtnTop = document.getElementById('undoBtnTop');
    const redoBtnTop = document.getElementById('redoBtnTop');

    // Undo button should be enabled after the first change
    const shouldDisableUndo = this.historyIndex <= 0;  // Enable undo after the first change
    const shouldDisableRedo = this.historyIndex >= this.history.length - 1;

    if (undoBtn) {
        undoBtn.disabled = shouldDisableUndo;
        undoBtn.style.opacity = shouldDisableUndo ? '0.4' : '1';
        undoBtn.style.cursor = shouldDisableUndo ? 'not-allowed' : 'pointer';
        undoBtn.style.pointerEvents = shouldDisableUndo ? 'none' : 'auto';
    }

    if (redoBtn) {
        redoBtn.disabled = shouldDisableRedo;
        redoBtn.style.opacity = shouldDisableRedo ? '0.4' : '1';
        redoBtn.style.cursor = shouldDisableRedo ? 'not-allowed' : 'pointer';
        redoBtn.style.pointerEvents = shouldDisableRedo ? 'none' : 'auto';
    }

    // Mirror the same state on top buttons
    if (undoBtnTop) {
        undoBtnTop.disabled = shouldDisableUndo;
        undoBtnTop.style.opacity = shouldDisableUndo ? '0.4' : '1';
        undoBtnTop.style.cursor = shouldDisableUndo ? 'not-allowed' : 'pointer';
        undoBtnTop.style.pointerEvents = shouldDisableUndo ? 'none' : 'auto';
    }
    if (redoBtnTop) {
        redoBtnTop.disabled = shouldDisableRedo;
        redoBtnTop.style.opacity = shouldDisableRedo ? '0.4' : '1';
        redoBtnTop.style.cursor = shouldDisableRedo ? 'not-allowed' : 'pointer';
        redoBtnTop.style.pointerEvents = shouldDisableRedo ? 'none' : 'auto';
    }
}


    // ============================================
    // KEYBOARD SHORTCUTS
    // ============================================
    handleKeyboardShortcuts(e) {

        // ── STATE 2 → 3 AUTO-ADVANCE ─────────────────────────────────────────────
        // If the user is in State 2 (all text selected) and presses Backspace,
        // Delete, or a printable key, advance to State 3 immediately so structure
        // guards are active — without triggering a blur/refocus cycle.
        if (this._isTextSelected && !this._isEditing && this._txtClickBox) {
            const isPrintable = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
            const isClear     = e.key === 'Backspace' || e.key === 'Delete';

            if (isPrintable || isClear) {
                const txtBox = this._txtClickBox.querySelector('.sli-txt-box');
                if (txtBox) {
                    e.preventDefault(); // stop browser acting on key before guards are ready

                    // Transition state flags (same as State 2→3 click)
                    this._isEditing       = true;
                    this._isTextSelected  = false;
                    this._isShapeSelected = false;
                    txtBox.classList.remove('text-all-selected');
                    this._txtClickBox.classList.add('editing-mode');
                    this._txtClickBox.classList.remove('text-selected');
                    this._txtClickBox.style.cursor = 'text';

                    // skipFocusRaf=true: textbox already has focus from State 2,
                    // skipping the rAF focus call prevents a blur→commit cycle.
                    this.startDirectEdit(txtBox, null, true);

                    // Now apply the key action directly (guards are now registered)
                    const s = this._activeTextEditor;
                    if (!s) return;

                    if (isClear) {
                        s.clearAllPreserveStructure();
                    } else {
                        // Printable: replace all selected content with the typed char
                        const span = txtBox.querySelector('.span-txt');
                        if (span) {
                            span.textContent = e.key;
                            try {
                                const r  = document.createRange();
                                const tn = span.firstChild;
                                if (tn) { r.setStart(tn, tn.length); }
                                else    { r.setStart(span, 0); }
                                r.collapse(true);
                                window.getSelection().removeAllRanges();
                                window.getSelection().addRange(r);
                            } catch (_) {}
                        }
                    }
                }
                return;
            }
        }
        // ── END STATE 2 → 3 AUTO-ADVANCE ─────────────────────────────────────────

        // Ctrl/Cmd + B = Bold
        if ((e.ctrlKey || e.metaKey) && e.key === 'b') {
            e.preventDefault();
            document.execCommand('bold');
            this.updateFormatButtons();
        }

        // Ctrl/Cmd + I = Italic
        if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
            e.preventDefault();
            document.execCommand('italic');
            this.updateFormatButtons();
        }

        // Ctrl/Cmd + U = Underline
        if ((e.ctrlKey || e.metaKey) && e.key === 'u') {
            e.preventDefault();
            document.execCommand('underline');
            this.updateFormatButtons();
        }

        // Ctrl/Cmd + Z = Undo
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
            e.preventDefault();
            this.undo();
        }

        // Ctrl/Cmd + Shift + Z = Redo
        if ((e.ctrlKey || e.metaKey) && e.key === 'z' && e.shiftKey) {
            e.preventDefault();
            this.redo();
        }

        // Ctrl/Cmd + Y = Redo (alternate)
        if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
            e.preventDefault();
            this.redo();
        }

        // Delete key = Delete selected text box
        if (e.key === 'Delete' || e.key === 'Backspace') {
            const selectedTextBox = document.querySelector('.sli-txt-box.selected');
            if (selectedTextBox) {
                const selection = window.getSelection();
                const isEditing = selectedTextBox.contains(selection.anchorNode) &&
                    selection.toString().length > 0;

                if (!isEditing && document.activeElement !== selectedTextBox) {
                    e.preventDefault();
                    this.deleteTextBox(selectedTextBox);
                }
            }
        }

        // Escape key = Deselect all
        if (e.key === 'Escape') {
            this.deselectAllTextBoxes();
        }
    }



    // ============================================
    // WINDOW CONTROLS
    // ============================================
    minimize() {
        this.showNotification('Minimize clicked');
    }

    maximize() {
        if (document.fullscreenElement) {
            document.exitFullscreen();
        } else {
            document.documentElement.requestFullscreen();
        }
    }

    close() {
        if (confirm('Are you sure you want to close? Any unsaved changes will be lost.')) {
            window.close();
        }
    }

    // ============================================
    // UI UPDATES
    // ============================================
    updateUI() {
        document.getElementById('currentSlide').textContent = this.currentSlide;
        document.getElementById('totalSlides').textContent = this.totalSlides;
    }

    // ============================================
    // NOTIFICATIONS
    // ============================================
    showNotification(message) {
        const notification = document.createElement('div');
        notification.className = 'notification';
        notification.textContent = message;
        notification.style.cssText = `
            position: fixed;
            bottom: 40px;
            left: 50%;
            transform: translateX(-50%);
            background: #3a3a3a;
            color: #e0e0e0;
            padding: 12px 24px;
            border-radius: 8px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
            z-index: 10000;
            animation: slideIn 0.3s ease-out;
            font-size: 14px;
            font-weight: 500;
        `;

        document.body.appendChild(notification);

        setTimeout(() => {
            notification.style.animation = 'fadeOut 0.3s ease-out';
            setTimeout(() => notification.remove(), 300);
        }, 2000);
    }

    // ============================================
    // UNIVERSAL ELEMENT INTERACTIVE (For ALL Elements)
    // ============================================
    makeElementInteractive(element, options = {}) {

        const defaults = {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 50,
            minHeight: 50
        };

        const settings = { ...defaults, ...options };

        // ── Ensure handles are not clipped by parent overflow ────────────────
        // image-container is created with overflow:hidden (so image doesn't spill),
        // but handles must protrude outside the element boundary → overflow:visible.
        // We store the original value so _getCleanState can restore it.
        if (!element._origOverflow) {
            element._origOverflow = element.style.overflow || '';
        }
        element.style.overflow = 'visible';

        // ── Remove stale handles first ────────────────────────────────────────
        element.querySelectorAll('.resize-handle, .delete-btn, .rotate-handle').forEach(h => h.remove());

        // ── Detach previous named listeners to prevent stacking ───────────────
        // After undo/redo the element is a NEW DOM node (innerHTML replaced), so
        // _meiHoverIn etc. are undefined — these removeEventListener calls are
        // no-ops on new nodes but correctly clean up on re-calls on live nodes.
        if (element._meiHoverIn)    element.removeEventListener('mouseenter', element._meiHoverIn);
        if (element._meiHoverOut)   element.removeEventListener('mouseleave', element._meiHoverOut);
        if (element._meiClick)      element.removeEventListener('click',      element._meiClick);

        // ── Define and store named handlers ──────────────────────────────────

        // Hover is now managed exclusively by the canvas mousemove tracker
        // (set up once in setupCanvasHoverTracker). Per-element mouseenter/
        // mouseleave are NOT used — they cause sibling bleed when elements overlap.
        // We only need the click handler here.
        element._meiClick = (e) => {
            if (e.target.classList.contains('resize-handle') ||
                e.target.classList.contains('delete-btn') ||
                e.target.classList.contains('rotate-handle')) return;
            // Text shapes are handled by the unified canvas click handler
            if (element.querySelector('.sli-txt-box')) return;
            this.selectElement(element, e.ctrlKey || e.metaKey);
        };

        if (element._meiClick_prev) element.removeEventListener('click', element._meiClick_prev);
        element._meiClick_prev = element._meiClick;
        element.addEventListener('click', element._meiClick);

        // ── Add handles ───────────────────────────────────────────────────────
        if (settings.resize) this.addResizeHandles(element, settings.minWidth, settings.minHeight);
        if (settings.delete) this.addDeleteButton(element);
        if (settings.move)   this.setupPowerPointDrag(element);
        if (settings.rotate) this.addRotateHandle(element);

        return element;
    }
    selectElement(element, isMultiSelect = false) {
        if (!isMultiSelect) {
            this.deselectAllElements(false);
        } else {
            if (element.classList.contains('selected')) {
                element.classList.remove('selected');
                return;
            }
        }
        if (element.dataset.interactiveInit !== '1') {
    const type = element.dataset.interactiveType || 'shape';
    const opts = (type === 'text')
        ? { resize:true, move:true, rotate:true, delete:true, minWidth:80, minHeight:30 }
        : { resize:true, move:true, rotate:true, delete:true, minWidth:20, minHeight:20 };

    this.makeElementInteractive(element, opts);
    element.dataset.interactiveInit = '1';
}

        element.classList.add('selected');
        element.style.overflow = 'visible';
        
        // Preserve clip-path for shapes (CSS tries to remove it on .selected)
        if ((element.classList.contains('shape') || element.classList.contains('custom-shape')) && element.dataset.shapeType) {
            this.reapplyClipPath(element);
        }
        
        this.setContextualTabFromElement(element);
        // Ensure Shape Format tab reacts
        if ((element.classList.contains('shape') || element.classList.contains('custom-shape'))) {
            this.setContextualTabFromElement(element);
        }
    }

    // Reapply clip-path to shapes when selected (CSS removes it)
    reapplyClipPath(element) {
        const shapeType = element.dataset.shapeType;
        if (!shapeType) return;

        const clipPaths = {
            'rectangle': 'polygon(0% 0%, 100% 0%, 100% 100%, 0% 100%)',
            'circle': 'ellipse(50% 50% at 50% 50%)',
            'oval': 'ellipse(50% 50% at 50% 50%)',
            'triangle': 'polygon(50% 0%, 100% 100%, 0% 100%)',
            'pentagon': 'polygon(50% 0%, 100% 38%, 81% 100%, 19% 100%, 0% 38%)'
        };

        const clipPath = clipPaths[shapeType];
        if (clipPath) {
            element.style.setProperty('clip-path', clipPath, 'important');
            element.style.setProperty('-webkit-clip-path', clipPath, 'important');
        }
    }

    deselectAllElements(switchToHome = true) {
        // ── Remove .selected AND .hover from every interactive element ────────
        // This is the single canonical cleanup — must strip both classes so
        // no element keeps handles visible after deselection.
        document.querySelectorAll(
            '.sli-txt-box, .insertable-element, .image-container, .shape, .custom-shape, .shape-group'
        ).forEach(el => {
            el.classList.remove('selected', 'hover', 'active-text-box', 'text-selected', 'editing-mode', 'text-all-selected');
            el.style.outline = '';

            if (el.classList.contains('sli-txt-box')) {
                el.setAttribute('contenteditable', 'false');
                el.style.position  = el.dataset.preEditPosition  || '';
                el.style.height    = el.dataset.preEditHeight    || '100%';
                el.style.minHeight = el.dataset.preEditMinHeight || '';
                el.style.top       = el.dataset.preEditTop       || '';
                el.style.left      = el.dataset.preEditLeft      || '';
                el.style.width     = el.dataset.preEditWidth     || '100%';
                el.style.zIndex    = el.dataset.preEditZIndex    || '';
                const _c = el.parentElement;
                if (_c && _c !== document.body) {
                    _c.style.display  = _c.dataset.preEditDisplay  || '';
                    // image-container needs overflow:hidden; shapes need overflow:visible
                    if (_c.classList.contains('image-container')) {
                        _c.style.overflow = 'visible'; // keep visible for handles
                    } else {
                        _c.style.overflow = _c.dataset.preEditOverflow || 'visible';
                    }
                }
            }
        });

        // Reset 3-click state machine on text shapes
        document.querySelectorAll('.shape[data-click-state]').forEach(s => {
            s.dataset.clickState = '0';
            s.querySelector('.sli-txt-box')?.classList.remove('text-all-selected');
        });

        // Reset the canvas-level hover tracker so the next mousemove starts clean.
        // Without this, the tracker would compare topEl === this._hoverEl (stale),
        // think "same element, nothing to do", and skip removing .hover.
        this._hoverEl = null;

        this.activeTextBox    = null;
        this._txtClickBox     = null;
        this._isShapeSelected = false;
        this._isTextSelected  = false;
        this._isEditing       = false;

        delete document.body.dataset.activeTextBoxId;
        if (switchToHome) {
            this.switchTab('home');
        }
    }
    // ============================================
    // SHAPE FORMATTING (toolbar -> selected shapes)
    // ============================================
    getSelectedDrawElements() {
        // Shapes + groups (for future multi-select support)
        return Array.from(document.querySelectorAll('.shape.selected, .custom-shape.selected, .shape-group.selected, .image-container.selected'));
    }

    applyToSelectedShapes(applyFn) {
        const selected = this.getSelectedDrawElements().filter(el => el.classList.contains('shape') || el.classList.contains('shape-group'));
        selected.forEach(applyFn);
        return selected.length;
    }

    setupShapeFormatToolbar() {
        // Shape format controls (IDs must exist in presentation-editor.html)
        const fill = document.getElementById('shapeFillColor');
        const fillBar = document.getElementById('shapeFillColorBar');
        const outline = document.getElementById('shapeOutlineColor');
        const outlineBar = document.getElementById('shapeOutlineColorBar');
        const outlineWidth = document.getElementById('shapeOutlineWidth');
        const opacity = document.getElementById('shapeOpacity');
        const corner = document.getElementById('shapeCornerRadius');
        const widthInput = document.getElementById('shapeWidth');
        const heightInput = document.getElementById('shapeHeight');
        const align = document.getElementById('alignShapes');

        // If your HTML doesn't have these controls, just exit quietly
        if (!fill && !outline && !outlineWidth && !opacity && !corner && !widthInput && !heightInput && !align) return;

        const updateBar = (barEl, color) => {
            if (barEl) barEl.style.background = color || 'transparent';
        };

        // DISABLED: These are now handled by ShapeFormatController in shape-format.js
        // which properly handles SVG shapes
        /*
        if (fill) {
            fill.addEventListener('input', (e) => {
                const color = e.target.value;
                updateBar(fillBar, color);
                this.applyToSelectedShapes((el) => {
                    el.style.background = color;
                });
            });
            updateBar(fillBar, fill.value);
        }

        if (outline) {
            outline.addEventListener('input', (e) => {
                const color = e.target.value;
                updateBar(outlineBar, color);
                this.applyToSelectedShapes((el) => {
                    el.style.borderStyle = 'solid';
                    el.style.borderColor = color;
                    if (!el.style.borderWidth) el.style.borderWidth = '1px';
                });
            });
            updateBar(outlineBar, outline.value);
        }

        if (outlineWidth) {
            outlineWidth.addEventListener('input', (e) => {
                const px = parseFloat(e.target.value);
                if (Number.isNaN(px)) return;
                this.applyToSelectedShapes((el) => {
                    el.style.borderStyle = 'solid';
                    el.style.borderWidth = `${px}px`;
                    if (!el.style.borderColor) el.style.borderColor = '#000000';
                });
            });
        }
        */

        if (opacity) {
            opacity.addEventListener('input', (e) => {
                const val = parseFloat(e.target.value);
                if (Number.isNaN(val)) return;
                const o = Math.max(0, Math.min(100, val)) / 100;
                this.applyToSelectedShapes((el) => {
                    el.style.opacity = String(o);
                });
            });
        }

        if (corner) {
            corner.addEventListener('input', (e) => {
                const px = parseFloat(e.target.value);
                if (Number.isNaN(px)) return;
                this.applyToSelectedShapes((el) => {
                    el.style.borderRadius = `${px}px`;
                });
            });
        }

        const applySize = () => {
            const w = widthInput ? parseFloat(widthInput.value) : NaN;
            const h = heightInput ? parseFloat(heightInput.value) : NaN;
            this.applyToSelectedShapes((el) => {
                if (!Number.isNaN(w)) el.style.width = `${Math.max(1, w)}px`;
                if (!Number.isNaN(h)) el.style.height = `${Math.max(1, h)}px`;
            });
        };

        if (widthInput) widthInput.addEventListener('change', applySize);
        if (heightInput) heightInput.addEventListener('change', applySize);

        if (align) {
            align.addEventListener('change', (e) => {
                const mode = e.target.value;
                this.alignSelectedShapes(mode);
            });
        }

        const bind = (id, fn) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('click', fn);
        };

        bind('insertRectangle', () => this.insertShape('rectangle'));
        bind('insertCircle', () => this.insertShape('circle'));
        bind('insertOval', () => this.insertShape('oval'));
        bind('insertTriangle', () => this.insertShape('triangle'));
        bind('insertPentagon', () => this.insertShape('pentagon'));
        bind('insertHexagon', () => this.insertShape('hexagon'));

    }

    
    alignSelectedShapes(mode) {
        // Minimal align feature for multi-select: align to the first selected shape
        const selected = Array.from(document.querySelectorAll('.shape.selected'));
        if (selected.length < 2) return;

        const base = selected[0];
        const baseLeft = base.offsetLeft;
        const baseTop = base.offsetTop;
        const baseRight = baseLeft + base.offsetWidth;
        const baseBottom = baseTop + base.offsetHeight;
        const baseCenterX = baseLeft + base.offsetWidth / 2;
        const baseCenterY = baseTop + base.offsetHeight / 2;

        selected.slice(1).forEach(el => {
            switch (mode) {
                case 'align-left':
                    el.style.left = `${baseLeft}px`;
                    break;
                case 'align-center':
                    el.style.left = `${baseCenterX - el.offsetWidth / 2}px`;
                    break;
                case 'align-right':
                    el.style.left = `${baseRight - el.offsetWidth}px`;
                    break;
                case 'align-top':
                    el.style.top = `${baseTop}px`;
                    break;
                case 'align-middle':
                    el.style.top = `${baseCenterY - el.offsetHeight / 2}px`;
                    break;
                case 'align-bottom':
                    el.style.top = `${baseBottom - el.offsetHeight}px`;
                    break;
                default:
                    break;
            }
        });
    }




    // ============================================
    // UNIVERSAL RESIZE HANDLES
    // ============================================
    addResizeHandles(element, minWidth = 50, minHeight = 50) {

        // Check if handles already exist
        if (element.querySelector('.resize-handle')) {
            return;
        }

        const handles = [
            'top-left', 'top-center', 'top-right',
            'left-center', 'right-center',
            'bottom-left', 'bottom-center', 'bottom-right'
        ];

        handles.forEach(position => {
            const handle = document.createElement('div');
            handle.className = `resize-handle ${position}`;
            handle.addEventListener('mousedown', (e) => this.startResize(e, element, position, minWidth, minHeight));
            // Stop hover events from leaking out of handle into/from parent
            handle.addEventListener('mouseenter', (e) => e.stopPropagation());
            handle.addEventListener('mouseleave', (e) => e.stopPropagation());
            element.appendChild(handle);
        });

    }

    startResize(e, element, position, minWidth, minHeight) {
        e.preventDefault();
        e.stopPropagation();

        const startX = e.clientX;
        const startY = e.clientY;
        const startWidth = element.offsetWidth;
        const startHeight = element.offsetHeight;
        const startLeft = element.offsetLeft;
        const startTop = element.offsetTop;

        // Check if this is a group
        const isGroup = element.classList.contains('shape-group') || element.dataset.isGroup === 'true';

        // If it's a group, store initial positions and sizes of all inner shapes
        let innerShapesData = [];
        if (isGroup) {
            const innerShapes = element.querySelectorAll('.shape, .custom-shape');
            innerShapes.forEach(shape => {
                innerShapesData.push({
                    element: shape,
                    startLeft: shape.offsetLeft,
                    startTop: shape.offsetTop,
                    startWidth: shape.offsetWidth,
                    startHeight: shape.offsetHeight
                });
            });
        }

        // ── CANVA-STYLE TEXT SCALING ──────────────────────────────────────────
        // Snapshot every span's font-size and line-height at drag-start so we
        // can scale them proportionally as the box is resized.
        const txtBox = element.querySelector(':scope > .sli-txt-box');
        let spanSnapshots = [];   // [{ el, fontSize, lineHeight, hasLhPx }]
        let pSnapshots    = [];   // [{ el, lineHeight, hasLhPx }]

        if (txtBox) {
            txtBox.querySelectorAll('.span-txt').forEach(span => {
                const cs = window.getComputedStyle(span);
                const fs = parseFloat(span.style.fontSize) || parseFloat(cs.fontSize) || 16;
                const lhStr = span.style.lineHeight || cs.lineHeight || 'normal';
                const lhPx  = parseFloat(lhStr);
                spanSnapshots.push({
                    el: span,
                    fontSize: fs,
                    lineHeight: isNaN(lhPx) ? null : lhPx,
                    lineHeightIsMultiplier: !span.style.lineHeight || isNaN(parseFloat(span.style.lineHeight)),
                });
            });
            txtBox.querySelectorAll('p').forEach(p => {
                const cs = window.getComputedStyle(p);
                const lhStr = p.style.lineHeight || cs.lineHeight || 'normal';
                const lhPx  = parseFloat(lhStr);
                pSnapshots.push({
                    el: p,
                    lineHeight: isNaN(lhPx) ? null : lhPx,
                });
            });
            // Ensure sli-txt-box itself stays 100% of wrapper during drag
            txtBox.style.width  = '100%';
            txtBox.style.height = '100%';
            txtBox.style.overflow = 'visible';
        }
        // ─────────────────────────────────────────────────────────────────────

        const onMouseMove = (e) => {
            const deltaX = e.clientX - startX;
            const deltaY = e.clientY - startY;

            let newWidth = startWidth;
            let newHeight = startHeight;
            let newLeft = startLeft;
            let newTop = startTop;

            // Calculate new dimensions based on handle position
            if (position.includes('right')) {
                newWidth = Math.max(minWidth, startWidth + deltaX);
            }
            if (position.includes('left')) {
                newWidth = Math.max(minWidth, startWidth - deltaX);
                if (newWidth > minWidth) {
                    newLeft = startLeft + deltaX;
                }
            }
            if (position.includes('bottom')) {
                newHeight = Math.max(minHeight, startHeight + deltaY);
            }
            if (position.includes('top')) {
                newHeight = Math.max(minHeight, startHeight - deltaY);
                if (newHeight > minHeight) {
                    newTop = startTop + deltaY;
                }
            }

            // Apply new dimensions to the wrapper shape
            element.style.width  = newWidth  + 'px';
            element.style.height = newHeight + 'px';
            element.style.left   = newLeft   + 'px';
            element.style.top    = newTop    + 'px';

            // ── CANVA-STYLE: scale all span font-sizes proportionally ──────────
            if (spanSnapshots.length > 0) {
                // Use the smaller of the two ratios so text always fits both dimensions
                const scaleW = newWidth  / startWidth;
                const scaleH = newHeight / startHeight;

                // If only width changed (left/right handle) use width scale;
                // if only height changed (top/bottom handle) use height scale;
                // if both changed (corner handle) use the smaller to ensure fit.
                let scale;
                const widthChanged  = position.includes('left')  || position.includes('right');
                const heightChanged = position.includes('top')   || position.includes('bottom');
                if (widthChanged && heightChanged) {
                    scale = Math.min(scaleW, scaleH);
                } else if (widthChanged) {
                    scale = scaleW;
                } else {
                    scale = scaleH;
                }

                spanSnapshots.forEach(({ el, fontSize, lineHeight }) => {
                    const newFs = Math.max(6, Math.round(fontSize * scale * 10) / 10);
                    el.style.fontSize = newFs + 'px';
                    if (lineHeight !== null) {
                        const newLh = Math.max(6, Math.round(lineHeight * scale * 10) / 10);
                        el.style.lineHeight = newLh + 'px';
                    }
                });

                pSnapshots.forEach(({ el, lineHeight }) => {
                    if (lineHeight !== null) {
                        const newLh = Math.max(6, Math.round(lineHeight * scale * 10) / 10);
                        el.style.lineHeight = newLh + 'px';
                    }
                });
            }
            // ─────────────────────────────────────────────────────────────────

            // If it's a group, scale inner shapes proportionally
            if (isGroup && innerShapesData.length > 0) {
                const scaleX = newWidth / startWidth;
                const scaleY = newHeight / startHeight;

                innerShapesData.forEach(shapeData => {
                    const shape = shapeData.element;

                    // Scale position
                    const newShapeLeft = shapeData.startLeft * scaleX;
                    const newShapeTop  = shapeData.startTop  * scaleY;

                    // Scale size
                    const newShapeWidth  = shapeData.startWidth  * scaleX;
                    const newShapeHeight = shapeData.startHeight * scaleY;

                    // Apply new position and size
                    shape.style.left   = newShapeLeft   + 'px';
                    shape.style.top    = newShapeTop    + 'px';
                    shape.style.width  = newShapeWidth  + 'px';
                    shape.style.height = newShapeHeight + 'px';
                });
            }
        };

        const onMouseUp = () => {
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);

            // Reset sli-txt-box to 100% — wrapper now has correct explicit px dimensions
            if (txtBox) {
                txtBox.style.width    = '100%';
                txtBox.style.height   = '100%';
                txtBox.style.overflow = '';
            }

            this.saveState();

            if (isGroup) {
                // console.log('Group resize complete');
            }
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    }

    // ============================================
    // UNIVERSAL DELETE BUTTON
    // ============================================
    addDeleteButton(element) {
        // Check if delete button already exists
        if (element.querySelector('.delete-btn')) return;

        const deleteBtn = document.createElement('div');
        deleteBtn.className = 'delete-btn';
        deleteBtn.innerHTML = '×';
        deleteBtn.title = 'Delete';
        deleteBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.deleteElement(element);
        });
        // Prevent the delete button's own mouseenter/mouseleave from
        // bubbling up and confusing the parent's hover state.
        // Without this, entering the delete button fires mouseleave on the
        // parent → .hover removed → button disappears before user can click.
        deleteBtn.addEventListener('mouseenter', (e) => e.stopPropagation());
        deleteBtn.addEventListener('mouseleave', (e) => e.stopPropagation());
        element.appendChild(deleteBtn);
    }

   deleteElement(element) {
    element.remove();
    this.saveState();
    this.showNotification('Element deleted');
}


    // ============================================
    // POWERPOINT-STYLE DRAG FUNCTIONALITY
    // ============================================
    
    /**
     * Check if mouse is near the border of an element
     * @param {MouseEvent} e - Mouse event
     * @param {HTMLElement} element - Element to check
     * @param {number} threshold - Distance from edge in pixels (default: 10)
     * @returns {boolean} - True if mouse is on/near border
     */
    isNearBorder(e, element, threshold = 10) {
        const rect = element.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        
        // Check if within threshold pixels of any edge
        const nearLeft = x <= threshold;
        const nearRight = x >= rect.width - threshold;
        const nearTop = y <= threshold;
        const nearBottom = y >= rect.height - threshold;
        
        return nearLeft || nearRight || nearTop || nearBottom;
    }

    /**
     * Check if element has a text box child
     * @param {HTMLElement} element - Element to check
     * @returns {boolean} - True if element contains a text box
     */
    hasTextBox(element) {
        return element.querySelector('.sli-txt-box') !== null;
    }

    /**
     * Setup PowerPoint-style drag functionality
     * - Shapes with text-boxes: drag only on border
     * - Shapes without text-boxes: drag anywhere
     * - Image-containers: drag anywhere
     */
    setupPowerPointDrag(element) {
        const isImageContainer = element.classList.contains('image-container');
        const isShape = (element.classList.contains('shape') || element.classList.contains('custom-shape')) || element.classList.contains('shape-group');
        const hasTextContent = this.hasTextBox(element);

        // Cleanup old listeners if they exist
        if (element._dragCleanup) {
            element._dragCleanup();
        }

        // Store drag state on the element itself to avoid closure issues
        element._dragState = {
            isDragging: false,
            startX: 0,
            startY: 0,
            startLeft: 0,
            startTop: 0
        };

        // Update cursor based on element type and mouse position
        const updateCursor = (e) => {
            // Don't change cursor while dragging
            if (element._dragState.isDragging) return;

            // Skip if hovering over control handles — let the handle's own CSS cursor: *-resize win
            if (e.target.classList.contains('resize-handle') ||
                e.target.classList.contains('delete-btn') ||
                e.target.classList.contains('rotate-handle')) {
                // Don't touch element.style.cursor; handle CSS cursor takes precedence
                return;
            }

            // Image containers: always show move cursor
            if (isImageContainer) {
                element.style.cursor = 'move';
                return;
            }

            // Shapes without text: always show move cursor
            if (isShape && !hasTextContent) {
                element.style.cursor = 'move';
                return;
            }

            // Shapes with text: show move cursor only on border
            if (isShape && hasTextContent) {
                if (this.isNearBorder(e, element, 10)) {
                    element.style.cursor = 'move';
                } else {
                    element.style.cursor = 'default';
                }
            }
        };

        // Mousemove handler for cursor updates
        const onMouseMove = (e) => {
            updateCursor(e);
        };

        // Mousedown handler to start drag
        const onMouseDown = (e) => {
            // If this textbox is in overlay-edit mode, do not start drag/resize
            if (element && element.dataset && element.dataset._editing === '1') return;
            if (e && e.target && e.target.classList && false /* no overlay — direct edit in sli-txt-box */) return;
            // Skip if clicking on control handles
            if (e.target.classList.contains('resize-handle') ||
                e.target.classList.contains('delete-btn') ||
                e.target.classList.contains('rotate-handle')) {
                return;
            }

            // Skip if clicking on image inside image-container
            if (isImageContainer && e.target.tagName === 'IMG') {
                // Continue to allow drag
            }

                        // If the element contains text, allow drag from anywhere when NOT editing.
            // During rich editing, the overlay blocks propagation so we won't reach here.
            const hasTextContent = element.classList.contains('text-shape') || element.querySelector('.sli-txt-box');
            if (hasTextContent) {
                // If someone accidentally made the textbox itself editable, don't drag.
                const tb = element.querySelector('.sli-txt-box');
                if (tb && tb.isContentEditable) return;
            }
            // For images & non-text shapes, default to draggable behavior
            const canStartDrag = true;

            if (!canStartDrag) return;

            // Start dragging
            e.preventDefault();
            e.stopPropagation();

            element._dragState.isDragging = true;
            element._dragState.startX = e.clientX;
            element._dragState.startY = e.clientY;
            element._dragState.startLeft = element.offsetLeft;
            element._dragState.startTop = element.offsetTop;

            element.classList.add('dragable');
            document.body.style.cursor = 'move';
            element.style.cursor = 'move';

            // Attach global move and up handlers
            document.addEventListener('mousemove', onGlobalMouseMove);
            document.addEventListener('mouseup', onGlobalMouseUp);
        };

        // Global mousemove handler for actual dragging
        const onGlobalMouseMove = (e) => {
            if (!element._dragState.isDragging) return;

            const canvas = document.getElementById('canvas');
            const deltaX = e.clientX - element._dragState.startX;
            const deltaY = e.clientY - element._dragState.startY;

            let newLeft = element._dragState.startLeft + deltaX;
            let newTop = element._dragState.startTop + deltaY;

            // Keep within canvas bounds
            newLeft = Math.max(0, Math.min(newLeft, canvas.offsetWidth - element.offsetWidth));
            newTop = Math.max(0, Math.min(newTop, canvas.offsetHeight - element.offsetHeight));

            element.style.left = newLeft + 'px';
            element.style.top = newTop + 'px';
        };

        // Global mouseup handler to end drag
        const onGlobalMouseUp = () => {
            if (!element._dragState.isDragging) return;

            element._dragState.isDragging = false;
            element.classList.remove('dragable');
            document.body.style.cursor = '';
            element.style.cursor = '';
            
            // Remove global listeners
            document.removeEventListener('mousemove', onGlobalMouseMove);
            document.removeEventListener('mouseup', onGlobalMouseUp);
            
            this.saveState();
        };

        // Attach local listeners
        element.addEventListener('mousemove', onMouseMove);
        element.addEventListener('mousedown', onMouseDown);

        // Set initial cursor
        const fakeEvent = { 
            clientX: element.getBoundingClientRect().left + element.offsetWidth / 2,
            clientY: element.getBoundingClientRect().top + element.offsetHeight / 2,
            target: element
        };
        updateCursor(fakeEvent);

        // Store cleanup function
        element._dragCleanup = () => {
            element.removeEventListener('mousemove', onMouseMove);
            element.removeEventListener('mousedown', onMouseDown);
            document.removeEventListener('mousemove', onGlobalMouseMove);
            document.removeEventListener('mouseup', onGlobalMouseUp);
            delete element._dragState;
        };
    }


    normalizeExistingShapes() {
        document.querySelectorAll('.shape, .custom-shape').forEach(shape => {
            // Infer shapeType if missing
            if (!shape.dataset.shapeType) {
                if (shape.id && shape.id.toLowerCase().includes('round')) {
                    shape.dataset.shapeType = 'rectangle';
                } else if (shape.style.borderRadius === '50%') {
                    shape.dataset.shapeType = 'circle';
                } else {
                    shape.dataset.shapeType = 'rectangle';
                }
            }
        });
    }


    // ============================================
    // UNIVERSAL ROTATE HANDLE
    // ============================================
    addRotateHandle(element) {
        // Check if rotate handle already exists
        if (element.querySelector('.rotate-handle')) return;

        const rotateHandle = document.createElement('div');
        rotateHandle.className = 'rotate-handle';
        rotateHandle.innerHTML = '↻';
        rotateHandle.title = 'Rotate';

        rotateHandle.addEventListener('mousedown', (e) => this.startRotate(e, element));
        rotateHandle.addEventListener('mouseenter', (e) => e.stopPropagation());
        rotateHandle.addEventListener('mouseleave', (e) => e.stopPropagation());
        element.appendChild(rotateHandle);
    }

    startRotate(e, element) {
        e.preventDefault();
        e.stopPropagation();

        const rect = element.getBoundingClientRect();
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        // Get current rotation or default to 0
        const currentTransform = window.getComputedStyle(element).transform;
        let currentRotation = 0;
        if (currentTransform && currentTransform !== 'none') {
            const values = currentTransform.split('(')[1].split(')')[0].split(',');
            const a = parseFloat(values[0]);
            const b = parseFloat(values[1]);
            currentRotation = Math.atan2(b, a) * (180 / Math.PI);
        }

        // Calculate initial angle from center to mouse position
        const startDeltaX = e.clientX - centerX;
        const startDeltaY = e.clientY - centerY;
        const startAngle = Math.atan2(startDeltaY, startDeltaX) * (180 / Math.PI);

        const onMouseMove = (moveEvent) => {
            const deltaX = moveEvent.clientX - centerX;
            const deltaY = moveEvent.clientY - centerY;
            const currentAngle = Math.atan2(deltaY, deltaX) * (180 / Math.PI);
            
            // Calculate the rotation delta from start position
            let angleDiff = currentAngle - startAngle;
            
            // Apply the delta to the current rotation
            let newRotation = currentRotation + angleDiff;

            // Optional: Snap to 15-degree increments when Shift is held (like PowerPoint)
            if (moveEvent.shiftKey) {
                newRotation = Math.round(newRotation / 15) * 15;
            } else {
                // Snap to whole degrees for smoother feel
                newRotation = Math.round(newRotation);
            }

            element.style.transform = `rotate(${newRotation}deg)`;
        };

        const onMouseUp = () => {
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);
            this.saveState();
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
    }


// ============================================
// CANVAS SELECTION HELPERS (ADD TO PresentationEditor CLASS)
// ============================================
// Add these methods to the PresentationEditor class in editor-core.js
// These ensure ALL operations work ONLY on canvas elements, never on preview elements

/**
 * Get the currently visible slide in canvas
 */
getVisibleCanvasSlide() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return null;
    }

    const slides = canvas.querySelectorAll('.sli-slide');
    for (let slide of slides) {
        // Check if slide is visible (not hidden)
        if (slide.style.display !== 'none') {
            console.log('✅ Found visible canvas slide');
            return slide;
        }
    }

    // If no visible slide found, return first slide
    if (slides.length > 0) {
        console.log('✅ Returning first canvas slide');
        return slides[0];
    }

    console.error('❌ No slides found in canvas');
    return null;
}

/**
 * Ensure element is from canvas, not preview
 * If element is from preview, find corresponding canvas element
 */
ensureCanvasElement(element) {
    if (!element) return null;

    // CRITICAL: Check if element is in a preview
    const inPreview = element.closest('.sli-preview') !== null;

    if (!inPreview) {
        // Already in canvas, verify it's actually in canvas
        const canvas = document.querySelector('.canvas');
        if (canvas && canvas.contains(element)) {
            console.log('✅ Element is in canvas');
            return element;
        }
    }

    console.warn('⚠️ Element is in preview, finding canvas equivalent');

    // Element is in preview or not in canvas, find canvas equivalent
    const elementId = element.id || element.getAttribute('data-textbox-id') || 
                     element.getAttribute('data-name');

    if (!elementId) {
        console.warn('⚠️ Preview element has no ID, cannot find canvas equivalent');
        return null; // Return null instead of preview element
    }

    // Search in canvas
    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) return null;

    // Try different selectors
    let canvasElement = null;
    
    // Try by ID
    if (element.id) {
        canvasElement = visibleSlide.querySelector(`#${element.id}`);
    }
    
    // Try by data-textbox-id
    if (!canvasElement && element.getAttribute('data-textbox-id')) {
        const tbId = element.getAttribute('data-textbox-id');
        canvasElement = visibleSlide.querySelector(`[data-textbox-id="${tbId}"]`);
    }
    
    // Try by data-name
    if (!canvasElement && element.getAttribute('data-name')) {
        const name = element.getAttribute('data-name');
        canvasElement = visibleSlide.querySelector(`[data-name="${name}"]`);
    }

    if (canvasElement) {
        console.log('✅ Found canvas element for preview element');
        return canvasElement;
    }

    console.error('❌ Could not find canvas equivalent for preview element');
    return null;
}

/**
 * Get selected shape from canvas only
 */
getCanvasSelectedShape() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return null;
    }

    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) return null;

    const selected = visibleSlide.querySelector('.shape.selected, .shape-group.selected');
    
    if (!selected) {
        console.warn('⚠️ No shape selected in canvas');
        return null;
    }

    // Verify it's actually in canvas, not preview
    if (canvas.contains(selected) && !selected.closest('.sli-preview')) {
        console.log('✅ Found selected shape in canvas');
        return selected;
    }

    console.warn('⚠️ Selected shape is in preview, not canvas');
    return this.ensureCanvasElement(selected);
}

/**
 * Get all selected shapes from canvas only
 */
getCanvasSelectedShapes() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return [];
    }

    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) return [];

    const selected = Array.from(visibleSlide.querySelectorAll('.shape.selected'));
    
    // Filter out any preview elements and ensure canvas elements
    const canvasShapes = selected
        .filter(el => canvas.contains(el) && !el.closest('.sli-preview'))
        .map(el => this.ensureCanvasElement(el))
        .filter(Boolean);

    console.log(`✅ Found ${canvasShapes.length} selected shapes in canvas`);
    return canvasShapes;
}

/**
 * Get selected image from canvas only
 */
getCanvasSelectedImage() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return null;
    }

    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) return null;

    const selected = visibleSlide.querySelector('.image-container.selected');
    
    if (!selected) {
        console.warn('⚠️ No image selected in canvas');
        return null;
    }

    // Verify it's actually in canvas, not preview
    if (canvas.contains(selected) && !selected.closest('.sli-preview')) {
        console.log('✅ Found selected image in canvas');
        return selected;
    }

    console.warn('⚠️ Selected image is in preview, not canvas');
    return this.ensureCanvasElement(selected);
}

/**
 * Get active textbox from canvas only - STRICT VERSION
 */
getCanvasActiveTextBox() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return null;
    }

    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) {
        console.warn('⚠️ No visible canvas slide');
        return null;
    }

    // First try: selected textbox in canvas
    let textbox = visibleSlide.querySelector('.sli-txt-box.selected');
    if (textbox) {
        // STRICT CHECK: Verify it's in canvas, not preview
        if (!canvas.contains(textbox) || textbox.closest('.sli-preview')) {
            console.warn('⚠️ Selected textbox is in preview, not canvas');
            textbox = null;
        } else {
            console.log('✅ Found selected textbox in canvas:', textbox.id || textbox.getAttribute('data-textbox-id'));
            return textbox;
        }
    }

    // Second try: use text selection
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
        const node = selection.anchorNode;
        const el = node?.nodeType === 3 ? node.parentElement : node;
        textbox = el?.closest('.sli-txt-box');

        // STRICT CHECK: Verify it's in canvas visible slide, NOT in preview
        if (textbox && visibleSlide.contains(textbox) && canvas.contains(textbox) && !textbox.closest('.sli-preview')) {
            console.log('✅ Found textbox from selection in canvas:', textbox.id || textbox.getAttribute('data-textbox-id'));
            return textbox;
        } else if (textbox) {
            console.warn('⚠️ Textbox from selection is in preview, not canvas');
        }
    }

    console.warn('⚠️ No active textbox found in canvas');
    return null;
}

/**
 * Sync changes from canvas element to all preview instances
 */
syncCanvasToPreview(canvasElement) {
    if (!canvasElement) {
        console.warn('⚠️ No canvas element to sync');
        return;
    }

    // Verify element is from canvas
    const canvas = document.querySelector('.canvas');
    if (!canvas || !canvas.contains(canvasElement) || canvasElement.closest('.sli-preview')) {
        console.error('❌ Element is not from canvas, cannot sync');
        return;
    }

    console.log('🔄 Syncing canvas element to preview...');

    const elementId = canvasElement.id || canvasElement.getAttribute('data-textbox-id') || 
                     canvasElement.getAttribute('data-name');

    if (!elementId) {
        console.warn('⚠️ Element has no ID, will refresh entire preview');
        if (this.currentSlide && this.refreshSlidePreview) {
            setTimeout(() => this.refreshSlidePreview(this.currentSlide), 50);
        }
        return;
    }

    // Find all preview instances
    const previews = document.querySelectorAll('.sli-preview');
    let syncedCount = 0;
    
    previews.forEach(preview => {
        let previewElement = null;

        // Try different selectors
        if (canvasElement.id) {
            previewElement = preview.querySelector(`#${canvasElement.id}`);
        }
        if (!previewElement && canvasElement.getAttribute('data-textbox-id')) {
            const tbId = canvasElement.getAttribute('data-textbox-id');
            previewElement = preview.querySelector(`[data-textbox-id="${tbId}"]`);
        }
        if (!previewElement && canvasElement.getAttribute('data-name')) {
            const name = canvasElement.getAttribute('data-name');
            previewElement = preview.querySelector(`[data-name="${name}"]`);
        }

        if (previewElement) {
            // Copy styles
            const stylesToCopy = [
                'backgroundColor', 'background', 'color', 'fontSize', 'fontFamily', 
                'fontWeight', 'fontStyle', 'textDecoration', 'textAlign',
                'borderColor', 'borderWidth', 'borderStyle', 'borderRadius',
                'boxShadow', 'filter', 'transform', 'opacity', 'width', 'height'
            ];

            stylesToCopy.forEach(prop => {
                if (canvasElement.style[prop]) {
                    previewElement.style[prop] = canvasElement.style[prop];
                }
            });

            // For textboxes, also copy innerHTML
            if (canvasElement.classList.contains('sli-txt-box') && 
                previewElement.classList.contains('sli-txt-box')) {
                previewElement.innerHTML = canvasElement.innerHTML;
            }
            
            syncedCount++;
        }
    });

    console.log(`✅ Synced to ${syncedCount} preview instance(s)`);

    // Refresh preview for complete sync
    if (this.currentSlide && this.refreshSlidePreview) {
        setTimeout(() => {
            this.refreshSlidePreview(this.currentSlide);
            console.log('✅ Preview refreshed');
        }, 50);
    }
}

/**
 * Clear all selections in canvas
 */
clearCanvasSelection() {
    const canvas = document.querySelector('.canvas');
    if (!canvas) {
        console.error('❌ Canvas not found');
        return;
    }

    const visibleSlide = this.getVisibleCanvasSlide();
    if (!visibleSlide) return;

    visibleSlide.querySelectorAll('.selected').forEach(el => {
        // Only clear selection if element is actually in canvas
        if (canvas.contains(el) && !el.closest('.sli-preview')) {
            el.classList.remove('selected');
        }
    });
    
    console.log('✅ Canvas selection cleared');
}

    
}

// ============================================
// INITIALIZE EDITOR
// ============================================
window.editor = null;

document.addEventListener('DOMContentLoaded', () => {
    window.editor = new PresentationEditor();
});

// ============================================
// ANIMATIONS
// ============================================
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(-50%) translateY(20px); opacity: 0; }
        to   { transform: translateX(-50%) translateY(0);    opacity: 1; }
    }
    @keyframes fadeOut {
        from { opacity: 1; }
        to   { opacity: 0; }
    }

    /* ── Handle visibility ────────────────────────────────────────────────
       Handles are always in the DOM (appended once, never removed on hover-out).
       They are hidden by default and shown only when parent is hovered or selected.
       pointer-events:none when hidden so they never steal mouse from siblings.
    ──────────────────────────────────────────────────────────────────────── */
    .resize-handle,
    .delete-btn,
    .rotate-handle {
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.1s;
    }

    /* Show handles when parent is hovered or selected */
    .hover > .resize-handle,
    .hover > .delete-btn,
    .hover > .rotate-handle,
    .selected > .resize-handle,
    .selected > .delete-btn,
    .selected > .rotate-handle {
        opacity: 1;
        pointer-events: auto;
    }

    /* The border/outline for selection — show only on selected */
    .shape.selected,
    .image-container.selected,
    .shape-group.selected {
        outline: 2px dashed #4A90E2;
        outline-offset: 2px;
    }

    /* Do NOT show selection border on hover alone */
    .shape.hover:not(.selected),
    .image-container.hover:not(.selected),
    .shape-group.hover:not(.selected) {
        outline: 1px dashed rgba(74,144,226,0.5);
        outline-offset: 2px;
    }
`;
document.head.appendChild(style);