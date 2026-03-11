/**
 * SHAPE FORMAT TAB FUNCTIONALITY
 * PowerPoint-like behavior
 */

/* =========================
   SHARED ARRANGE UTILITY
   Used by both ShapeFormatController (shapes/textboxes)
   and ImageFormatTab (images) so that bringToFront /
   imgBringToFront and sendToBack / imgSendToBack all
   operate on whatever element is currently selected,
   regardless of its type.
========================== */
window.ArrangeUtils = {

    /**
     * Returns the currently selected element of any type:
     * shape, shape-group, image-container, or sli-txt-box.
     */
    getSelectedElement() {
        return document.querySelector(
            '.shape.selected, .custom-shape.selected, .shape-group.selected, .image-container.selected, .sli-txt-box.selected'
        );
    },

    /**
     * Finds the active slide's sli-content wrapper.
     */
    getSliContent() {
        const canvas = document.getElementById('canvas');
        if (!canvas) return null;
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(s => s.style.display !== 'none');
        if (!currentSlide) return null;
        return currentSlide.querySelector('.sli-content') || currentSlide;
    },

    /**
     * Returns all interactive elements inside the active slide.
     */
    getAllElements(sliContent) {
        return Array.from(sliContent.querySelectorAll(
            '.shape, .custom-shape, .shape-group, .image-container, .sli-txt-box, .insertable-element'
        ));
    },

    /** Bring the selected element above every other element. */
    bringToFront(saveStateFn) {
        const target = this.getSelectedElement();
        if (!target) return;

        const sliContent = this.getSliContent();
        if (!sliContent) return;

        let maxZ = 0;
        this.getAllElements(sliContent).forEach(el => {
            if (el !== target) maxZ = Math.max(maxZ, parseInt(el.style.zIndex || 0));
        });

        target.style.zIndex = String(maxZ + 1);
        if (typeof saveStateFn === 'function') saveStateFn();
    },

    /** Send the selected element below every other element. */
    sendToBack(saveStateFn) {
        const target = this.getSelectedElement();
        if (!target) return;

        const sliContent = this.getSliContent();
        if (!sliContent) return;

        let minZ = 999999;
        this.getAllElements(sliContent).forEach(el => {
            if (el !== target) minZ = Math.min(minZ, parseInt(el.style.zIndex || 0));
        });

        const newZ = (minZ === 999999) ? 1 : Math.max(1, minZ - 1);
        target.style.zIndex = String(newZ);
        if (typeof saveStateFn === 'function') saveStateFn();
    }
};

class ShapeFormatController {
    constructor(editor) {
        this.editor = editor;
        this.shapeGroup = null;
        this.init();
    }

    init() {
        this.setupEventListeners();
    }

    /* =========================
       SELECTION HELPERS
    ========================== */

    getSelectedShape() {
        return document.querySelector('.shape.selected, .custom-shape.selected, .shape-group.selected');
    }

    getSelectedShapes() {
        return Array.from(document.querySelectorAll('.shape.selected, .custom-shape.selected'));
    }

    clearSelection() {
        document
            .querySelectorAll('.shape.selected, .custom-shape.selected, .shape-group.selected')
            .forEach(el => el.classList.remove('selected'));
    }

    setSelection(elements) {
        this.clearSelection();
        elements.forEach(el => el.classList.add('selected'));
    }

    /* =========================
       EVENT WIRING
    ========================== */

    setupEventListeners() {
        const bind = (id, fn) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('click', fn);
        };

        const bindChange = (id, fn) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('change', e => fn(e.target.value));
        };

        const bindInput = (id, fn) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('input', e => fn(e.target.value));
        };

        // Use 'input' for color pickers for immediate updates
        bindInput('shapeFillColor', v => this.changeShapeFill(v));
        bindInput('shapeOutlineColor', v => this.changeShapeOutline(v));
        bindInput('outlineWidth', v => this.changeOutlineWidth(v));
        bindChange('shapeEffect', v => this.applyShapeEffect(v));

        bind('bringToFront', () => ArrangeUtils.bringToFront(() => this.editor.saveState()));
        bind('imgBringToFront', () => ArrangeUtils.bringToFront(() => this.editor.saveState()));
        bind('sendToBack', () => ArrangeUtils.sendToBack(() => this.editor.saveState()));
        bind('imgSendToBack', () => ArrangeUtils.sendToBack(() => this.editor.saveState()));

        bind('rotateLeft90', () => this.rotateShape(-90));
        bind('rotateRight90', () => this.rotateShape(90));
        bind('flipHorizontal', () => this.flipShape('horizontal'));
        bind('flipVertical', () => this.flipShape('vertical'));

        bind('groupShapes', () => this.groupShapes());
        bind('ungroupShapes', () => this.ungroupShapes());

        bindChange('alignShapes', v => v && this.alignShapes(v));

        const on = (id, fn, ev = 'click') => {
            const el = document.getElementById(id);
            if (el) el.addEventListener(ev, fn);
        };
        
        // Capture the shape when user focuses the input, apply to IT on change/blur
        on('shapeWidth', e => this.updateShapeSize('width', e.target.value), 'input');
        on('shapeHeight', e => this.updateShapeSize('height', e.target.value), 'input');

        // Clamp value to min/max in real-time as user types
        ['shapeWidth', 'shapeHeight'].forEach(id => {
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
            const c = e.target.closest('.shape, .custom-shape');
            if (c) this.getSelectedShape(c);
        });
    }


    /* -------------------------------
       SIZE
    -------------------------------- */

    updateShapeSize(type, value) {
        const c = this.getSelectedShape();
        if (!c) return;

        let v = parseFloat(value);
        if (isNaN(v) || v <= 0) return;
        v = Math.min(1000, Math.max(10, v)); // enforce min:10 max:1000

        if (type === 'width') {
            c.style.width = v + 'px';
            
        } else {
            c.style.height = v + 'px';
           
        }

       this.editor.saveState();
    }

    /* =========================
       FILL / OUTLINE
    ========================== */

    changeShapeFill(color) {
        const shape = this.getSelectedShape();
        if (!shape) return;

        const svgElement = shape.querySelector('svg path, svg polygon');

        if (svgElement) {
            // SVG-based shape (hexagon) — colour the SVG fill
            svgElement.setAttribute('fill', color);
            shape.style.background = 'transparent';
        } else if (shape.style.getPropertyValue('--shape-clip')) {
            // CSS custom property shape — update --shape-bg, ::before reads it
            shape.style.setProperty('--shape-bg', color);
            shape.style.background = 'transparent';
        } else {
            // Legacy fallback
            const inner = shape.querySelector('.shape-inner') || shape;
            inner.style.backgroundColor = color;
        }

        this.editor.saveState();
    }

    changeShapeOutline(color) {
        const shape = this.getSelectedShape();
        if (!shape) return;

        const inner = shape.querySelector('.shape-inner') || shape;
        const svgElement = shape.querySelector('svg path, svg polygon');

        if (svgElement) svgElement.setAttribute('stroke', color);
        else inner.style.borderColor = color;

        this.editor.saveState();
    }

    changeOutlineWidth(width) {
        const shape = this.getSelectedShape();
        if (!shape) return;

        const inner = shape.querySelector('.shape-inner') || shape;
        const svgElement = shape.querySelector('svg path, svg polygon');

        if (svgElement) svgElement.setAttribute('stroke-width', width);
        else inner.style.borderWidth = `${width}px`;

        this.editor.saveState();
    }

    /* =========================
       EFFECTS
    ========================== */

    applyShapeEffect(effect) {
    const shape = this.getSelectedShape();
    if (!shape) return;

    // Reset previous effects
    shape.style.boxShadow = '';
    shape.style.filter = '';
    shape.style.webkitBoxReflect = '';
    shape.classList.remove('effect-3d-format');
    shape.dataset.effect = effect || 'none';

    // Clean any previous 3D rotate additions from transform
    let t = shape.style.transform || '';
    if (t === 'none') t = '';
    t = t
        .replace(/perspective\([^)]+\)/g, '')
        .replace(/rotateX\([^)]+\)/g, '')
        .replace(/rotateY\([^)]+\)/g, '')
        .trim();
    shape.style.transform = t;

    // Effects
    if (effect === 'shadow') {
        // alpha MUST be 0..1
        shape.style.boxShadow = '0 6px 14px rgba(0,0,0,0.3)';
    }
    else if (effect === 'glow') {
        shape.style.boxShadow = '0 0 18px rgba(74,144,226,0.9)';
    }
    else if (effect === 'soft-edges') {
        // Subtle blur + drop shadow to mimic "soft edges"
        // (avoid heavy blur because it will look like out-of-focus)
        shape.style.filter = 'blur(0.6px) drop-shadow(0 3px 10px rgba(0,0,0,0.18))';
    }
    else if (effect === 'reflection') {
        // Works in Chrome/Edge (Chromium): simplest working reflection
        // (PowerPoint-like "reflection below")
        shape.style.webkitBoxReflect =
            'below 4px linear-gradient(transparent, rgba(0,0,0,0.22))';
    }
    else if (effect === '3d-format') {
        // Fake 3D bevel using inset shadows (works on div + svg shapes)
        shape.style.boxShadow =
            'inset 2px 2px 6px rgba(255,255,255,0.35), inset -3px -3px 8px rgba(0,0,0,0.25), 0 6px 14px rgba(0,0,0,0.18)';
    }
    else if (effect === '3d-rotation') {
        // Add a 3D tilt without breaking existing rotate()/scale() too much
        shape.style.transform = `${t} perspective(900px) rotateX(18deg) rotateY(-18deg)`;
        shape.style.boxShadow = '0 10px 22px rgba(0,0,0,0.22)';
    }

    this.editor.saveState();
}


    /* =========================
       ARRANGE
    ========================== */

    bringToFront() {
        ArrangeUtils.bringToFront(() => this.editor.saveState());
    }

    sendToBack() {
        ArrangeUtils.sendToBack(() => this.editor.saveState());
    }

    /* =========================
       ROTATE / FLIP
    ========================== */

    rotateShape(deg) {
        const s = this.getSelectedShape();
        if (!s) return;

        let t = s.style.transform || '';
        if (t === 'none') t = '';

        const m = t.match(/rotate\(([-\d.]+)deg\)/);
        const cur = m ? parseFloat(m[1]) : 0;

        s.style.transform =
            t.replace(/rotate\([^)]+\)/, '').trim() +
            ` rotate(${cur + deg}deg)`;

        this.editor.saveState();
    }

    flipShape(dir) {
        const s = this.getSelectedShape();
        if (!s) return;

        let t = s.style.transform || '';
        if (t === 'none') t = '';

        const prop = dir === 'horizontal' ? 'scaleX' : 'scaleY';
        const has = new RegExp(`${prop}\\(-1\\)`).test(t);

        t = t.replace(new RegExp(`${prop}\\([^)]+\\)`), '').trim();
        s.style.transform = `${t} ${prop}(${has ? 1 : -1})`;

        this.editor.saveState();
    }

    /* =========================
       GROUP / UNGROUP
    ========================== */

    groupShapes() {
        const shapes = this.getSelectedShapes();
        if (shapes.length < 2) return;

        const canvas = document.getElementById('canvas');
        const group = document.createElement('div');
        group.className = 'shape-group';

        let minL = Infinity, minT = Infinity, maxR = 0, maxB = 0;

        shapes.forEach(s => {
            minL = Math.min(minL, s.offsetLeft);
            minT = Math.min(minT, s.offsetTop);
            maxR = Math.max(maxR, s.offsetLeft + s.offsetWidth);
            maxB = Math.max(maxB, s.offsetTop + s.offsetHeight);
        });

        group.style.position = 'absolute';
        group.style.left = `${minL}px`;
        group.style.top = `${minT}px`;
        group.style.width = `${maxR - minL}px`;
        group.style.height = `${maxB - minT}px`;

        shapes.forEach(s => {
            s.style.left = `${s.offsetLeft - minL}px`;
            s.style.top = `${s.offsetTop - minT}px`;
            s.classList.remove('selected');
            group.appendChild(s);
        });

        canvas.appendChild(group);
        this.editor.makeElementInteractive(group, { move: true, resize: true, rotate: true, delete: true });

        this.setSelection([group]);
        this.shapeGroup = group;
        this.editor.saveState();
    }

    ungroupShapes() {
        const group = document.querySelector('.shape-group.selected');
        if (!group) return;

        const canvas = document.getElementById('canvas');
        const gl = group.offsetLeft;
        const gt = group.offsetTop;

        Array.from(group.children).forEach(s => {
            s.style.left = `${gl + s.offsetLeft}px`;
            s.style.top = `${gt + s.offsetTop}px`;
            canvas.appendChild(s);
            this.editor.makeElementInteractive(s, { move: true, resize: true, rotate: true, delete: true });
        });

        group.remove();
        this.clearSelection();
        this.editor.saveState();
    }

    /* =========================
       ALIGN
    ========================== */

    alignShapes(type) {
        const shapes = this.getSelectedShapes();
        if (!shapes.length) return;

        const c = document.getElementById('sli-slide');
        const cw = c.offsetWidth;
        const ch = c.offsetHeight;

        shapes.forEach(s => {
            if (type === 'left') s.style.left = '0px';
            if (type === 'center') s.style.left = `${(cw - s.offsetWidth) / 2}px`;
            if (type === 'right') s.style.left = `${cw - s.offsetWidth}px`;
            if (type === 'top') s.style.top = '0px';
            if (type === 'middle') s.style.top = `${(ch - s.offsetHeight) / 2}px`;
            if (type === 'bottom') s.style.top = `${ch - s.offsetHeight}px`;
        });

        this.editor.saveState();
    }

    /* =========================
       RESIZE
    ========================== */

    resizeShape(w, h) {
        const s = this.getSelectedShape();
        if (!s) return;

        if (w !== null && w !== undefined) {
            const val = parseFloat(w);
            if (!isNaN(val) && val > 0) {
                s.style.setProperty('width', `${val}px`, 'important');
                s.setAttribute('data-width', val);
            }
        }
        if (h !== null && h !== undefined) {
            const val = parseFloat(h);
            if (!isNaN(val) && val > 0) {
                s.style.setProperty('height', `${val}px`, 'important');
                s.setAttribute('data-height', val);
            }
        }

        // Force reflow so the browser commits the new size before saveState reads it
        void s.offsetWidth;

        // Defer saveState slightly so any blur/focus side-effects settle first
        setTimeout(() => {
            // Re-apply in case something reset it during the event cycle
            if (w !== null && w !== undefined) {
                const val = parseFloat(w);
                if (!isNaN(val) && val > 0) s.style.setProperty('width', `${val}px`, 'important');
            }
            if (h !== null && h !== undefined) {
                const val = parseFloat(h);
                if (!isNaN(val) && val > 0) s.style.setProperty('height', `${val}px`, 'important');
            }
            this.editor.saveState();
        }, 50);
    }

    /**
     * Populate the width/height inputs with the selected shape's current dimensions.
     * Call this whenever a shape becomes selected.
     */
    updateSizeInputs() {
        const s = this.getSelectedShape();
        const wInput = document.getElementById('shapeWidth');
        const hInput = document.getElementById('shapeHeight');
        if (!wInput || !hInput) return;

        if (s) {
            // Prefer data attributes set by resizeShape; fall back to rendered size
            wInput.value = s.dataset.width ? Math.round(s.dataset.width) : Math.round(s.offsetWidth);
            hInput.value = s.dataset.height ? Math.round(s.dataset.height) : Math.round(s.offsetHeight);
        } else {
            wInput.value = '';
            hInput.value = '';
        }
    }
}

/* =========================
   INIT
========================== */

document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        if (window.editor) {
            new ShapeFormatController(window.editor);
        }
    }, 100);
});