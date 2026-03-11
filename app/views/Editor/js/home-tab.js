/**
 * HOME TAB FUNCTIONALITY
 * Text formatting, alignment, colors, lists, and zoom controls
 */

class HomeTabController {
    constructor(editor) {
        this.editor = editor;
        this.init();
    }

    init() {
        this.setupEventListeners();
    }


    applyToTextHierarchy(tb, applyFn, spanMetaFn) {
        if (!tb) return;

        const shape = tb.closest('.shape, .custom-shape') || tb;

        // 1️⃣ shape
        applyFn(shape, 'shape');

        // 2️⃣ sli-txt-box
        applyFn(tb, 'textbox');

        // 3️⃣ paragraphs
        tb.querySelectorAll('p').forEach(p => {
            applyFn(p, 'p');
        });

        // 4️⃣ spans + metadata
        tb.querySelectorAll('.span-txt').forEach(span => {
            applyFn(span, 'span');
            if (spanMetaFn) spanMetaFn(span);
        });
    }

    /**
     * Returns true if user has a partial text selection inside the given textbox.
     * "Partial" means: some text is selected, and it's NOT the entire textbox content.
     */
    hasPartialSelection(tb) {
        // Check live selection first
        const sel = window.getSelection();
        let range = null;
        if (sel && !sel.isCollapsed && sel.rangeCount > 0) {
            range = sel.getRangeAt(0);
            // Must be inside the textbox
            if (!tb.contains(range.commonAncestorContainer)) range = null;
        }
        // Fall back to range saved on toolbar mousedown (before focus shifted)
        if (!range && this._savedSelectionRange) {
            const saved = this._savedSelectionRange;
            if (tb.contains(saved.commonAncestorContainer)) {
                range = saved;
                // Restore it so execCommand and extractContents work
                try { sel.removeAllRanges(); sel.addRange(range); } catch (_) {}
            }
        }
        if (!range) return false;
        const selectedText = (range.toString() || '').trim();
        if (!selectedText) return false;
        // If entire textbox content is selected → treat as "apply to all"
        const totalText = tb.textContent.trim();
        return selectedText !== totalText;
    }

    /**
     * Apply a style to only the currently selected range by wrapping it in a
     * <span class="span-txt"> with the given style property set.
     * Uses execCommand for simple commands, or manual span wrapping for complex styles.
     */
    applyToSelection(command, value) {
        // Map command to CSS property + value
        const sel = window.getSelection();
        if (!sel || sel.isCollapsed || sel.rangeCount === 0) return;

        const cssProp = {
            bold:          'fontWeight',
            italic:        'fontStyle',
            underline:     'textDecoration',
            strikeThrough: 'textDecoration',
            foreColor:     'color',
            fontSize:      'fontSize',
            fontFamily:    'fontFamily',
        };

        // Compute toggle value for bold/italic/underline/strike
        const range = sel.getRangeAt(0);
        const anchorNode = range.commonAncestorContainer;
        const anchorEl = anchorNode.nodeType === 3 ? anchorNode.parentElement : anchorNode;
        const cs = anchorEl ? window.getComputedStyle(anchorEl) : null;

        let prop, val;
        if (command === 'bold') {
            prop = 'fontWeight';
            val = cs && (parseInt(cs.fontWeight) >= 600 || cs.fontWeight === 'bold') ? 'normal' : 'bold';
        } else if (command === 'italic') {
            prop = 'fontStyle';
            val = cs && cs.fontStyle === 'italic' ? 'normal' : 'italic';
        } else if (command === 'underline') {
            prop = 'textDecoration';
            val = cs && cs.textDecorationLine.includes('underline') ? 'none' : 'underline';
        } else if (command === 'strikeThrough') {
            prop = 'textDecoration';
            val = cs && cs.textDecorationLine.includes('line-through') ? 'none' : 'line-through';
        } else if (command === 'foreColor') {
            prop = 'color'; val = value;
        } else if (command === 'fontSize') {
            prop = 'fontSize'; val = value + 'px';
        } else if (command === 'fontFamily') {
            prop = 'fontFamily'; val = value;
        } else {
            prop = command; val = value;
        }

        this._wrapSelectionWithStyle(prop, val);

        // After applying color to a selection inside a <li>, sync the <li>'s own
        // color so the bullet marker matches. The bullet ::marker inherits from <li>,
        // not from child <span> elements.
        if (command === 'foreColor' && value) {
            const range2 = sel.getRangeAt ? sel.getRangeAt(0) : null;
            if (range2) {
                const node = range2.commonAncestorContainer;
                const li = (node.nodeType === 3 ? node.parentElement : node).closest('li');
                if (li) {
                    // Use the color of the first span in this li as the bullet color
                    const firstSpan = li.querySelector('.span-txt');
                    li.style.color = firstSpan ? (firstSpan.style.color || value) : value;
                }
            }
        }
    }

    _wrapSelectionWithStyle(prop, value) {
        const sel = window.getSelection();
        if (!sel || sel.isCollapsed || sel.rangeCount === 0) return;
        const range = sel.getRangeAt(0);

        // ── MERGE PATH ────────────────────────────────────────────────────────────
        // If the selection is already exactly one .span-txt, mutate it in-place.
        const targetSpan = this._getExactSpan(range);
        if (targetSpan) {
            targetSpan.style[prop] = value;
            this._setSpanMeta(targetSpan, prop, value);
            const nr = document.createRange();
            nr.selectNodeContents(targetSpan);
            sel.removeAllRanges();
            sel.addRange(nr);
            this._savedSelectionRange = nr.cloneRange();
            return;
        }

        // ── SPLIT-PARENT PATH ─────────────────────────────────────────────────────
        // The selection is a substring of text inside a parent .span-txt.
        // We must SPLIT the parent span at the selection boundaries — never nest.
        // Result: [before-span] [new-span with styles] [after-span]
        const parentSpan = this._getParentSpan(range);
        if (parentSpan) {
            this._splitAndWrap(parentSpan, range, prop, value, sel);
            return;
        }

        // ── FALLBACK WRAP PATH ────────────────────────────────────────────────────
        // Selection crosses multiple elements or has no span parent — generic wrap.
        const anchor = range.commonAncestorContainer;
        const ctxEl  = anchor.nodeType === 3 ? anchor.parentElement : anchor;
        const ctxSpan = ctxEl && ctxEl.closest ? ctxEl.closest('.span-txt') : null;
        const inheritedCss   = ctxSpan ? (ctxSpan.style.cssText || '') : '';
        const inheritedAttrs = {};
        if (ctxSpan) {
            ['originalea','origincs','latinfont','originaltxtcolor','originsym',
             'alpha','cap','originallummod','originallumoff'].forEach(a => {
                if (ctxSpan.hasAttribute(a)) inheritedAttrs[a] = ctxSpan.getAttribute(a);
            });
        }
        const fragment = range.extractContents();
        const wrapper  = document.createElement('span');
        wrapper.className = 'span-txt';
        wrapper.style.cssText = inheritedCss;
        wrapper.style[prop] = value;
        for (const [k, v] of Object.entries(inheritedAttrs)) wrapper.setAttribute(k, v);
        this._setSpanMeta(wrapper, prop, value);
        wrapper.appendChild(fragment);
        range.insertNode(wrapper);
        const nr = document.createRange();
        nr.selectNodeContents(wrapper);
        sel.removeAllRanges();
        sel.addRange(nr);
        this._savedSelectionRange = nr.cloneRange();
    }

    /**
     * Split parentSpan at the selection boundaries, producing up to 3 sibling spans:
     *   [text before selection] [selected text with new style] [text after selection]
     * All three inherit parentSpan's styles. The middle one gets the new property too.
     */
    _splitAndWrap(parentSpan, range, prop, value, sel) {
        const fullText   = parentSpan.textContent;
        const selText    = range.toString();
        const beforeText = this._textBefore(parentSpan, range);
        const afterText  = fullText.slice(beforeText.length + selText.length);

        // Inherit styles + attributes from parent
        const baseCss   = parentSpan.style.cssText || '';
        const baseAttrs = {};
        ['originalea','origincs','latinfont','originaltxtcolor','originsym',
         'alpha','cap','originallummod','originallumoff'].forEach(a => {
            if (parentSpan.hasAttribute(a)) baseAttrs[a] = parentSpan.getAttribute(a);
        });

        const makeSpan = (text, extraProp, extraVal) => {
            const s = document.createElement('span');
            s.className = 'span-txt';
            s.style.cssText = baseCss;
            for (const [k, v] of Object.entries(baseAttrs)) s.setAttribute(k, v);
            if (extraProp) {
                s.style[extraProp] = extraVal;
                this._setSpanMeta(s, extraProp, extraVal);
            }
            s.textContent = text;
            return s;
        };

        const parent = parentSpan.parentNode;

        // Middle span — inherits base styles + new property
        const midSpan = makeSpan(selText, prop, value);

        if (beforeText) parent.insertBefore(makeSpan(beforeText, null, null), parentSpan);
        parent.insertBefore(midSpan, parentSpan);
        if (afterText)  parent.insertBefore(makeSpan(afterText, null, null),  parentSpan);
        parent.removeChild(parentSpan);

        // Select the middle span so chained formats hit merge path
        const nr = document.createRange();
        nr.selectNodeContents(midSpan);
        sel.removeAllRanges();
        sel.addRange(nr);
        this._savedSelectionRange = nr.cloneRange();
    }

    /** Get the text content of parentSpan that comes BEFORE the range start. */
    _textBefore(parentSpan, range) {
        const beforeRange = document.createRange();
        beforeRange.setStart(parentSpan, 0);
        beforeRange.setEnd(range.startContainer, range.startOffset);
        return beforeRange.toString();
    }

    /**
     * Returns the nearest .span-txt ancestor if the selection is a substring
     * entirely within one span (both endpoints in the same span, not the whole text).
     */
    _getParentSpan(range) {
        const startNode = range.startContainer;
        const endNode   = range.endContainer;
        const startEl   = startNode.nodeType === 3 ? startNode.parentElement : startNode;
        const endEl     = endNode.nodeType   === 3 ? endNode.parentElement   : endNode;
        const startSpan = startEl && startEl.closest ? startEl.closest('.span-txt') : null;
        const endSpan   = endEl   && endEl.closest   ? endEl.closest('.span-txt')   : null;
        // Both ends in same span and text doesn't cover the whole span
        if (startSpan && startSpan === endSpan && startSpan.textContent !== range.toString()) {
            return startSpan;
        }
        return null;
    }

    /**
     * If the Range exactly covers one .span-txt (entire text content), return it.
     */
    _getExactSpan(range) {
        const selText = range.toString();
        if (!selText) return null;
        const startNode = range.startContainer;
        const endNode   = range.endContainer;
        const startEl   = startNode.nodeType === 3 ? startNode.parentElement : startNode;
        const endEl     = endNode.nodeType   === 3 ? endNode.parentElement   : endNode;
        const startSpan = startEl && startEl.closest ? startEl.closest('.span-txt') : null;
        const endSpan   = endEl   && endEl.closest   ? endEl.closest('.span-txt')   : null;
        if (startSpan && startSpan === endSpan && startSpan.textContent === selText) {
            return startSpan;
        }
        const ancestor = range.commonAncestorContainer;
        if (ancestor.nodeType === 1 && ancestor.classList &&
            ancestor.classList.contains('span-txt') && ancestor.textContent === selText) {
            return ancestor;
        }
        return null;
    }

    _setSpanMeta(span, prop, value) {
        if (prop === 'color')        span.setAttribute('originaltxtcolor', value);
        if (prop === 'fontFamily') { span.setAttribute('origincs', value); span.setAttribute('latinfont', value); }
        if (prop === 'fontSize')     span.setAttribute('fontsize', parseFloat(value));
        if (prop === 'fontWeight')   span.setAttribute('fontweight', value);
        if (prop === 'fontStyle')    span.setAttribute('fontstyle', value);
        if (prop === 'textDecoration') {
            span.setAttribute('underline', value === 'underline' ? 'true' : 'false');
            span.setAttribute('strikethrough', value === 'line-through' ? 'true' : 'false');
        }
    }

    getActiveTextBox() {
        // State 3: direct edit active — return the editing textBox
        if (this.editor._activeTextEditor) {
            return this.editor._activeTextEditor.textBox;
        }
        // State 1 & 2: shape selected but not editing yet — return _activeTxtBox
        // This allows formatting tools to work from the very first click (Canva-style)
        if (this.editor._activeTxtBox) {
            return this.editor._activeTxtBox;
        }
        // Fallbacks
        return document.querySelector('.sli-txt-box.selected') ||
            window.getSelection()?.anchorNode?.parentElement?.closest('.sli-txt-box');
    }

    // Returns the active overlay editor element (contenteditable div) if open
    getOverlayEditor() {
        return this.editor._activeTextEditor?.ed || null;
    }


    setupEventListeners() {
        // Prevent toolbar button clicks from blurring the overlay editor
        // by intercepting mousedown on all toolbar buttons (mousedown fires before blur)
        document.addEventListener('mousedown', (e) => {
            const isToolbarClick =
                e.target.closest('.toolbar') ||
                e.target.closest('[class*="toolbar"]') ||
                e.target.id === 'boldBtn' || e.target.id === 'italicBtn' ||
                e.target.id === 'underlineBtn' || e.target.id === 'strikeBtn' ||
                e.target.id === 'alignLeftBtn' || e.target.id === 'alignCenterBtn' ||
                e.target.id === 'alignRightBtn' || e.target.id === 'alignJustifyBtn' ||
                e.target.id === 'bulletBtn' || e.target.id === 'numberBtn' ||
                e.target.id === 'textColor' || e.target.id === 'highlightColor' ||
                e.target.closest('.font-select') || e.target.closest('.font-size');
            if (!isToolbarClick) return;

            // Always save the selection range first (synchronous, before any blur fires)
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
                try { this._savedSelectionRange = sel.getRangeAt(0).cloneRange(); } catch (_) {}
            } else {
                this._savedSelectionRange = null;
            }

            // <select> dropdowns: NEVER preventDefault — it prevents the dropdown opening.
            // We rely on _savedSelectionRange to apply formatting after the change event.
            const isSelect = e.target.tagName === 'SELECT' || e.target.closest('select');
            if (isSelect) return;

            // For all other toolbar elements (buttons, color pickers, etc.):
            // preventDefault whenever text is selected (State 2 OR State 3) so blur
            // does not fire and the selection stays alive for hasPartialSelection.
            if (this._savedSelectionRange || this.editor._activeTextEditor) {
                e.preventDefault();
            }
        }, true);


        // Undo/Redo Buttons
        const undoBtn = document.getElementById('undoBtn');
        const redoBtn = document.getElementById('redoBtn');
        if (undoBtn) undoBtn.addEventListener('click', () => this.editor.undo());
        if (redoBtn) redoBtn.addEventListener('click', () => this.editor.redo());

        // Add Slide Button
        const addSlideBtn = document.getElementById('addSlideBtn');
        if (addSlideBtn) {
            addSlideBtn.addEventListener('click', () => this.editor.addSlide());
        }

        // Text Formatting Buttons
        const boldBtn = document.getElementById('boldBtn');
        const italicBtn = document.getElementById('italicBtn');
        const underlineBtn = document.getElementById('underlineBtn');
        const strikeBtn = document.getElementById('strikeBtn');

        if (boldBtn) boldBtn.addEventListener('click', () => this.toggleBold());
        if (italicBtn) italicBtn.addEventListener('click', () => this.toggleItalic());
        if (underlineBtn) underlineBtn.addEventListener('click', () => this.toggleUnderline());
        if (strikeBtn) strikeBtn.addEventListener('click', () => this.toggleStrikethrough());


        // Font Controls
        const fontFamily = document.getElementById('fontFamily');
        const fontSize = document.getElementById('fontSize');

        if (fontFamily) {
            fontFamily.addEventListener('change', (e) => this.changeFontFamily(e.target.value));
        }
        if (fontSize) {
            fontSize.addEventListener('change', (e) => {
                this.changeFontSize(e.target.value);
                this.editor.showNotification(`Font size changed to ${e.target.value}px`);
            });
        }

        // Color Controls
        const textColor = document.getElementById('textColor');
        const highlightColor = document.getElementById('highlightColor');

        if (textColor) {
            textColor.addEventListener('input', (e) => this.changeTextColor(e.target.value));
            textColor.addEventListener('change', (e) => this.changeTextColor(e.target.value));
        }
        if (highlightColor) {
            highlightColor.addEventListener('input', (e) => this.changeHighlightColor(e.target.value));
            highlightColor.addEventListener('change', (e) => this.changeHighlightColor(e.target.value));
        }

        // Alignment Buttons
        const alignLeftBtn = document.getElementById('alignLeftBtn');
        const alignCenterBtn = document.getElementById('alignCenterBtn');
        const alignRightBtn = document.getElementById('alignRightBtn');
        const alignJustifyBtn = document.getElementById('alignJustifyBtn');

        if (alignLeftBtn) alignLeftBtn.addEventListener('click', () => this.setTextAlign('left'));
        if (alignCenterBtn) alignCenterBtn.addEventListener('click', () => this.setTextAlign('center'));
        if (alignRightBtn) alignRightBtn.addEventListener('click', () => this.setTextAlign('right'));
        if (alignJustifyBtn) alignJustifyBtn.addEventListener('click', () => this.setTextAlign('justify'));


        // List Buttons
        const bulletBtn = document.getElementById('bulletBtn');
        const numberBtn = document.getElementById('numberBtn');

        if (bulletBtn) bulletBtn.addEventListener('click', () => this.toggleList('bullet'));
        if (numberBtn) numberBtn.addEventListener('click', () => this.toggleList('number'));


        // Zoom Controls
        const zoomInBtn = document.getElementById('zoomInBtn');
        const zoomOutBtn = document.getElementById('zoomOutBtn');

        if (zoomInBtn) zoomInBtn.addEventListener('click', () => this.zoomIn());
        if (zoomOutBtn) zoomOutBtn.addEventListener('click', () => this.zoomOut());

        // Keyboard shortcuts for undo/redo and list Enter handling
        document.addEventListener('keydown', (e) => {
            // Ctrl/Cmd + Z = Undo (without Shift)
            if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z') && !e.shiftKey) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
                if (this.editor && typeof this.editor.undo === 'function') {
                    this.editor.undo();
                }
                return false;
            }

            // Ctrl/Cmd + Shift + Z = Redo (Mac/Windows)
            if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z') && e.shiftKey) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
                if (this.editor && typeof this.editor.redo === 'function') {
                    this.editor.redo();
                }
                return false;
            }

            // Ctrl/Cmd + Y = Redo (alternate for Windows)
            if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y')) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
                if (this.editor && typeof this.editor.redo === 'function') {
                    this.editor.redo();
                }
                return false;
            }

            // Handle Enter and Backspace keys in lists
            if (e.key !== 'Enter' && e.key !== 'Backspace') return;

            const tb = this.getActiveTextBox();
            if (!tb) return;

            const sel = window.getSelection();
            const node = sel?.anchorNode;
            const el = node?.nodeType === 3 ? node.parentElement : node;
            const li = el?.closest?.('li');
            const list = li?.parentElement;

            if (!li || !list || !['UL', 'OL'].includes(list.tagName)) return;

            // ── BACKSPACE handler ──────────────────────────────────────────────
            if (e.key === 'Backspace') {
                if (!sel.rangeCount) return;
                const range = sel.getRangeAt(0);
                if (!range.collapsed) return; // let browser handle selection deletion

                // Determine if cursor is effectively at the start of this <li>
                // "Effectively at start" means: all content before cursor is empty or ZWS only
                const getTextBeforeCursor = () => {
                    const r = document.createRange();
                    r.setStart(li, 0);
                    r.setEnd(range.startContainer, range.startOffset);
                    return r.toString().replace(/\u200B/g, ''); // strip ZWS
                };

                const textBefore = getTextBeforeCursor();
                if (textBefore.length > 0) return; // cursor is mid-text, let browser handle

                e.preventDefault();
                this.editor.saveState();

                const allLis = Array.from(list.querySelectorAll('li'));
                const liIndex = allLis.indexOf(li);

                if (liIndex > 0) {
                    // Merge into previous <li>
                    const prevLi = allLis[liIndex - 1];

                    // Get caret position at end of prevLi before merge
                    const prevLastChild = prevLi.lastChild;
                    const prevTextNode = prevLastChild?.nodeType === 3
                        ? prevLastChild
                        : prevLastChild?.lastChild?.nodeType === 3
                            ? prevLastChild.lastChild
                            : null;

                    // Move all children from current li to prevLi
                    // But skip pure-ZWS spans if prevLi already has content
                    const prevHasRealContent = prevLi.textContent.replace(/\u200B/g, '').trim().length > 0;
                    Array.from(li.childNodes).forEach(child => {
                        const isZWSOnly = child.nodeType === 3
                            ? child.textContent.replace(/\u200B/g, '') === ''
                            : child.textContent.replace(/\u200B/g, '') === '';
                        if (prevHasRealContent && isZWSOnly) return; // skip ZWS placeholders
                        prevLi.appendChild(child);
                    });
                    li.remove();

                    // Place caret at join point
                    requestAnimationFrame(() => {
                        const target = prevTextNode || (() => {
                            const spans = prevLi.querySelectorAll('.span-txt');
                            const lastSpan = spans[spans.length - 1];
                            return lastSpan?.firstChild || lastSpan || prevLi;
                        })();
                        const newRange = document.createRange();
                        if (target.nodeType === 3) {
                            const offset = prevTextNode
                                ? prevTextNode.textContent.replace(/\u200B$/, '').length
                                : target.textContent.length;
                            newRange.setStart(target, Math.min(offset, target.textContent.length));
                        } else {
                            newRange.setStart(target, target.childNodes?.length || 0);
                        }
                        newRange.collapse(true);
                        sel.removeAllRanges();
                        sel.addRange(newRange);
                    });
                } else {
                    // First <li> — remove the list entirely, convert to plain paragraph
                    // Collect ALL li content to restore as paragraphs
                    const ps = [];
                    allLis.forEach(item => {
                        const p = document.createElement('p');
                        p.style.margin = '0';
                        while (item.firstChild) p.appendChild(item.firstChild);
                        ps.push(p);
                    });

                    list.replaceWith(...ps);
                    this._teardownListMutationObserver(tb);
                    this._removeEmptyParagraphs(tb);

                    // Ensure textbox has at least one usable paragraph
                    if (!tb.querySelector('p')) {
                        const cs = window.getComputedStyle(tb);
                        const fallbackP = document.createElement('p');
                        fallbackP.style.margin = '0';
                        const fallbackSpan = document.createElement('span');
                        fallbackSpan.className = 'span-txt';
                        fallbackSpan.style.cssText = `font-family:${cs.fontFamily};font-size:${cs.fontSize};color:${cs.color};line-height:${cs.lineHeight||'1.4'};`;
                        fallbackSpan.setAttribute('originalea', '+mn-ea');
                        fallbackSpan.setAttribute('origincs', cs.fontFamily);
                        fallbackSpan.setAttribute('latinfont', cs.fontFamily);
                        fallbackSpan.setAttribute('originaltxtcolor', cs.color);
                        fallbackSpan.textContent = '\u200B';
                        fallbackP.appendChild(fallbackSpan);
                        tb.appendChild(fallbackP);
                    }

                    // Place caret in the first paragraph
                    requestAnimationFrame(() => {
                        const firstP = tb.querySelector('p');
                        if (!firstP) return;
                        const textNode = firstP.querySelector('.span-txt')?.firstChild || firstP.firstChild;
                        const newRange = document.createRange();
                        if (textNode?.nodeType === 3) {
                            newRange.setStart(textNode, 0);
                        } else {
                            newRange.setStart(firstP, 0);
                        }
                        newRange.collapse(true);
                        sel.removeAllRanges();
                        sel.addRange(newRange);
                    });
                }

                this.autoGrowAndFlow(tb);
                return;
            }

            e.preventDefault();
            this.editor.saveState();

            const range = sel.getRangeAt(0);

            // ── Find which span the cursor is in ─────────────────────────────────
            const cursorNode   = range.startContainer; // text node or element
            const cursorOffset = range.startOffset;
            const cursorSpan   = cursorNode.nodeType === 3
                ? cursorNode.parentElement.closest('.span-txt')
                : cursorNode.closest && cursorNode.closest('.span-txt');

            // Get all direct child spans of the current <li>
            const allSpans = Array.from(li.querySelectorAll(':scope > .span-txt'));

            // ── New <li> inherits alignment/line-height from current <li> ────────
            const newLi = document.createElement('li');
            if (li.style.textAlign)  newLi.style.textAlign  = li.style.textAlign;
            if (li.style.lineHeight) newLi.style.lineHeight = li.style.lineHeight;

            // Helper: clone a span with same styles/attrs but given text content
            const cloneSpanWithText = (srcSpan, text) => {
                const s = document.createElement('span');
                s.className = 'span-txt';
                s.style.cssText = srcSpan.style.cssText || '';
                ['originalea','origincs','latinfont','originaltxtcolor','originsym',
                 'alpha','cap','originallummod','originallumoff'].forEach(a => {
                    if (srcSpan.hasAttribute(a)) s.setAttribute(a, srcSpan.getAttribute(a));
                });
                s.textContent = text || '';
                return s;
            };

            // ── Split logic ───────────────────────────────────────────────────────
            let caretSpan; // the span in newLi where caret will be placed

            if (!cursorSpan || !li.contains(cursorSpan)) {
                // Cursor not in a span — append empty span to new li
                const firstSpan = allSpans[0];
                const emptySpan = firstSpan
                    ? cloneSpanWithText(firstSpan, '​')
                    : (() => { const s = document.createElement('span'); s.className='span-txt'; s.textContent='​'; return s; })();
                newLi.appendChild(emptySpan);
                caretSpan = emptySpan;
            } else {
                const spanIdx = allSpans.indexOf(cursorSpan);

                // Split the current span at the cursor offset
                const textNode  = cursorNode.nodeType === 3 ? cursorNode : cursorSpan.firstChild;
                const textOff   = cursorNode.nodeType === 3 ? cursorOffset : 0;
                const fullText  = textNode ? textNode.textContent : '';
                const textBefore = fullText.slice(0, textOff);
                const textAfter  = fullText.slice(textOff);

                // Update current span to only contain text before cursor
                if (textNode) textNode.textContent = textBefore;

                // Build new li: [after-part of split span] + [all spans after cursorSpan]
                if (textAfter) {
                    const afterSpan = cloneSpanWithText(cursorSpan, textAfter);
                    newLi.appendChild(afterSpan);
                    caretSpan = afterSpan;
                }

                // Move all spans that came AFTER cursorSpan in the original li
                for (let i = spanIdx + 1; i < allSpans.length; i++) {
                    newLi.appendChild(allSpans[i]); // moves DOM node, removes from li
                }

                // If newLi is still empty, add a placeholder span
                if (!newLi.hasChildNodes()) {
                    const ph = cloneSpanWithText(cursorSpan, '​');
                    newLi.appendChild(ph);
                    caretSpan = ph;
                }
                if (!caretSpan) caretSpan = newLi.firstChild;

                // Clean up: if current li's split span is now empty, give it a ZWS
                if (!textBefore && cursorSpan.textContent === '') {
                    cursorSpan.textContent = '​';
                }
            }

            // Insert new li after current li
            if (li.nextSibling) list.insertBefore(newLi, li.nextSibling);
            else list.appendChild(newLi);

            // Place caret at start of first text node in caretSpan
            const caretTarget = caretSpan.firstChild || caretSpan;
            const newRange = document.createRange();
            if (caretTarget.nodeType === 3) {
                // Skip leading zero-width space if present
                const off = caretTarget.textContent.startsWith('​') ? 1 : 0;
                newRange.setStart(caretTarget, off);
            } else {
                newRange.setStart(caretTarget, 0);
            }
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);

            this.autoGrowAndFlow(tb);
            this._removeEmptyParagraphs(tb);
        }, true); // Use capture phase to handle event before other listeners


        document.addEventListener('input', (e) => {
            const tb = e.target?.closest?.('.sli-txt-box');
            if (!tb) return;

            // Remove any empty <p> tags the browser may have inserted around the list
            this._removeEmptyParagraphs(tb);

            // If we are in list context, keep DOM clean
            if (this.isInListContext(tb)) {
                this.normalizeListStructure(tb);
            }

            // Auto expand and flow shapes down
            this.autoGrowAndFlow(tb);
        });


    }


    /** Remove empty <p> tags that browser inserts around lists in contenteditable. */
    _removeEmptyParagraphs(tb) {
        if (!tb) return;
        // Remove all empty <p> tags — browser inserts these around lists
        // when contenteditable is active. Keep only <p> tags with real content.
        tb.querySelectorAll('p').forEach(p => {
            const hasContent = p.textContent.trim() || p.querySelector('img, br, span');
            if (!hasContent) p.remove();
        });
    }

    _setupListMutationObserver(tb) {
        // Watch for browser-inserted empty <p> tags around lists and remove immediately
        if (tb._listObserver) return; // already watching
        const observer = new MutationObserver(() => {
            tb.querySelectorAll('p').forEach(p => {
                const hasContent = p.textContent.trim() || p.querySelector('img, br, span');
                if (!hasContent) p.remove();
            });
        });
        observer.observe(tb, { childList: true, subtree: false });
        tb._listObserver = observer;
    }

    _teardownListMutationObserver(tb) {
        if (tb && tb._listObserver) {
            tb._listObserver.disconnect();
            delete tb._listObserver;
        }
    }

    autoGrowAndFlow(tb) {
        if (!tb) return;

        const shape = tb.closest('.shape, .custom-shape');
        if (!shape) return;

        // Shape is position:absolute — growing it NEVER displaces other shapes.
        // Do NOT push siblings down: each shape is independently positioned.
        // Only grow the shape itself based on content, never move other elements.
        if (shape.dataset._editing !== '1') return;

        // Trigger the editor's autoSize if available (startDirectEdit sets it up)
        const editor = this.editor;
        if (editor && editor._activeTextEditor) {
            const { autoSize } = editor._activeTextEditor;
            if (typeof autoSize === 'function') {
                requestAnimationFrame(() => autoSize());
            }
        }
    }



    // ============================================
    // TEXT FORMATTING
    // ============================================

    changeFontFamily(fontName) {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('fontName', false, fontName); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        // Restore saved selection before applying
        const saved = this._savedSelectionRange;
        if (saved) {
            try {
                tb.focus();
                const s = window.getSelection();
                s.removeAllRanges();
                s.addRange(saved);
            } catch (_) {}
        }
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('fontFamily', fontName);
        } else {
            this.applyToTextHierarchy(tb,
                (el) => { el.style.fontFamily = fontName; },
                (span) => {
                    span.setAttribute('origincs', fontName);
                    span.setAttribute('latinfont', fontName);
                    if (!span.getAttribute('originalea')) span.setAttribute('originalea', '+mn-ea');
                    span.setAttribute('originsym', '');
                }
            );
        }
        this.editor.saveState();
    }



    changeFontSize(px) {
        const overlay = this.getOverlayEditor();
        if (overlay) {
            overlay.focus();
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
                document.execCommand('fontSize', false, '7');
                overlay.querySelectorAll('font[size="7"]').forEach(f => {
                    f.removeAttribute('size');
                    f.style.fontSize = `${px}px`;
                });
            }
            return;
        }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        // Restore saved selection before applying
        const saved = this._savedSelectionRange;
        if (saved) {
            try {
                tb.focus();
                const s = window.getSelection();
                s.removeAllRanges();
                s.addRange(saved);
            } catch (_) {}
        }
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('fontSize', px);
        } else {
            this.applyToTextHierarchy(tb,
                (el) => { el.style.fontSize = `${px}px`; },
                (span) => { span.setAttribute('fontsize', px); }
            );
        }
        this.editor.saveState();
    }


    cleanupNestedSpans(element) {
        const spans = element.querySelectorAll('span');
        spans.forEach(span => {
            if (span.style.fontSize && span.parentElement.style.fontSize === span.style.fontSize) {
                while (span.firstChild) {
                    span.parentElement.insertBefore(span.firstChild, span);
                }
                span.remove();
            }
        });
    }

    toggleBold() {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('bold'); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('bold');
        } else {
            const isBold = window.getComputedStyle(tb).fontWeight >= 600;
            this.applyToTextHierarchy(tb,
                (el) => { el.style.fontWeight = isBold ? 'normal' : 'bold'; },
                (span) => { span.setAttribute('fontweight', isBold ? 'normal' : 'bold'); }
            );
        }
        this.editor.saveState();
    }


    toggleItalic() {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('italic'); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('italic');
        } else {
            const isItalic = window.getComputedStyle(tb).fontStyle === 'italic';
            this.applyToTextHierarchy(tb,
                (el) => { el.style.fontStyle = isItalic ? 'normal' : 'italic'; },
                (span) => { span.setAttribute('fontstyle', isItalic ? 'normal' : 'italic'); }
            );
        }
        this.editor.saveState();
    }

    toggleUnderline() {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('underline'); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('underline');
        } else {
            const isUnderline = window.getComputedStyle(tb).textDecorationLine.includes('underline');
            this.applyToTextHierarchy(tb,
                (el) => { el.style.textDecoration = isUnderline ? 'none' : 'underline'; },
                (span) => { span.setAttribute('underline', isUnderline ? 'false' : 'true'); }
            );
        }
        this.editor.saveState();
    }

    toggleStrikethrough() {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('strikeThrough'); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;
        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('strikeThrough');
        } else {
            const isStrike = window.getComputedStyle(tb).textDecorationLine.includes('line-through');
            this.applyToTextHierarchy(tb,
                (el) => { el.style.textDecoration = isStrike ? 'none' : 'line-through'; },
                (span) => { span.setAttribute('strikethrough', isStrike ? 'false' : 'true'); }
            );
        }
        this.editor.saveState();
    }




    changeTextColor(color) {
        const overlay = this.getOverlayEditor();
        if (overlay) { overlay.focus(); document.execCommand('foreColor', false, color); return; }
        const tb = this.getActiveTextBox();
        if (!tb) return;

        if (this.hasPartialSelection(tb)) {
            this.applyToSelection('foreColor', color);
            // Also update the <li> bullet color for the li containing the selection
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0) {
                const anchor = sel.getRangeAt(0).commonAncestorContainer;
                const li = (anchor.nodeType === 3 ? anchor.parentElement : anchor).closest('li');
                if (li) li.style.color = color;
            }
            this.editor.saveState();
            return;
        }

        // Apply color to all spans
        this.applyToTextHierarchy(tb,
            (el, level) => { if (level === 'span') el.style.color = color; },
            (span) => {
                span.setAttribute('originaltxtcolor', color);
                span.removeAttribute('originallummod');
                span.removeAttribute('originallumoff');
            }
        );

        // Bullet color is driven by the <li> element's own color property.
        // Spans inside <li> do NOT affect the ::marker — only the <li> itself does.
        // Set color on every <li> so bullets always match the text color.
        tb.querySelectorAll('li').forEach(li => {
            li.style.color = color;
        });

        this.editor.saveState();
    }


   changeHighlightColor(color) {
        const tb = this.getActiveTextBox();
        if (!tb) return;

        const shape = tb.closest('.shape, .custom-shape');
        if (!shape) return;

        this.editor.saveState();

        // Apply background color only to the shape (parent div), not to sli-txt-box, p, or span
        shape.style.backgroundColor = color;
    }


    updateColorBar(barId, color) {
        const colorBar = document.getElementById(barId);
        if (colorBar) {
            colorBar.style.backgroundColor = color;
        }
    }

    // ============================================
    // ALIGNMENT
    // ============================================

    mapAlignToFlex(align) {
        switch (align) {
            case 'left': return 'flex-start';
            case 'center': return 'center';
            case 'right': return 'flex-end';
            case 'justify': return 'flex-start'; // keep start; justify is for text spacing in <p>
            default: return 'flex-start';
        }
    }


    setTextAlign(align) {
        const overlay = this.getOverlayEditor();
        if (overlay) {
            overlay.focus();
            document.execCommand('justifyLeft', false, null); // reset
            const cmdMap = { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' };
            document.execCommand(cmdMap[align] || 'justifyLeft');
            return;
        }
        const tb = this.getActiveTextBox();
        if (!tb) return;

        const shape = tb.closest('.shape, .custom-shape');
        const flexJustify = this.mapAlignToFlex(align);

        // ✅ 1) Update FLEX alignment for containers
        if (shape) {
            shape.style.justifyContent = flexJustify;
            shape.style.textAlign = align;
        }

        tb.style.justifyContent = flexJustify;
        tb.style.textAlign = align;

        // ✅ 2) Update paragraph alignment
        tb.querySelectorAll('p').forEach(p => {
            p.style.textAlign = align;
        });

        // ✅ 3) Apply text-align to spans
        tb.querySelectorAll('.span-txt').forEach(span => {
            span.style.textAlign = align;
        });

        // ✅ 4) For "justify", make sure paragraph width can justify properly
        if (align === 'justify') {
            tb.querySelectorAll('p').forEach(p => {
                p.style.textAlignLast = 'left';
            });
        } else {
            tb.querySelectorAll('p').forEach(p => {
                p.style.textAlignLast = '';
            });
        }

        this.editor.saveState();
    }



    // ============================================
    // LISTS
    // ============================================

   toggleList(type) {
    const tb = this.getActiveTextBox();
    if (!tb) return;

    const existingList = tb.querySelector('ul,ol');
    const targetTag = type === 'number' ? 'OL' : 'UL';

    // ==================================================
    // CASE 1: LIST EXISTS
    // ==================================================
    if (existingList) {

        // 🔁 TOGGLE-OFF (same button clicked again)
        if (existingList.tagName === targetTag) {
            
            this.editor.saveState();
            const items = Array.from(existingList.querySelectorAll('li'));

            items.forEach(li => {
                const p = document.createElement('p');
                p.style.margin = '0';

                // Copy alignment/line-height from li or list
                if (li.style.textAlign)  p.style.textAlign  = li.style.textAlign;
                if (li.style.lineHeight) p.style.lineHeight = li.style.lineHeight;

                // Move ALL child nodes (all span-txt elements) back into paragraph
                while (li.firstChild) {
                    p.appendChild(li.firstChild);
                }

                tb.insertBefore(p, existingList);
            });

            existingList.remove();
            this._teardownListMutationObserver(tb);
            this._removeEmptyParagraphs(tb);
            return;
        }

        // 🔄 CONVERT UL ↔ OL
        const newList = document.createElement(targetTag.toLowerCase());
        newList.style.margin = existingList.style.margin || '0px';
        newList.style.paddingLeft = existingList.style.paddingLeft || '25px';
        
        this.editor.saveState();
        Array.from(existingList.children).forEach(li => {
            if (li.tagName === 'LI') {
                newList.appendChild(li);
            }
        });

        existingList.replaceWith(newList);
        return;
    }

    // ==================================================
    // CASE 2: NO LIST → CREATE LIST FROM PARAGRAPHS
    // ==================================================

    /** Helper: create a default styled span for this textbox */
    const makeDefaultSpan = (text = '\u200B') => {
        const cs = window.getComputedStyle(tb);
        const s = document.createElement('span');
        s.className = 'span-txt';
        s.style.cssText = `font-family:${cs.fontFamily};font-size:${cs.fontSize};color:${cs.color};line-height:${cs.lineHeight||'1.4'};white-space:pre-wrap;`;
        s.setAttribute('originalea', '+mn-ea');
        s.setAttribute('origincs', cs.fontFamily);
        s.setAttribute('latinfont', cs.fontFamily);
        s.setAttribute('originaltxtcolor', cs.color);
        s.textContent = text;
        return s;
    };

    const list = document.createElement(type === 'number' ? 'ol' : 'ul');
    list.style.margin = '0px';
    list.style.paddingLeft = '25px';

    this.editor.saveState();

    let paragraphs = Array.from(tb.querySelectorAll('p'));

    // If textbox is empty (no paragraphs, or all paragraphs are ZWS/whitespace only)
    const hasRealContent = paragraphs.some(p => p.textContent.replace(/\u200B/g, '').trim().length > 0);
    if (!paragraphs.length || !hasRealContent) {
        // Remove whatever empty paragraphs exist
        paragraphs.forEach(p => p.remove());

        // Create a single empty list item
        const li = document.createElement('li');
        li.appendChild(makeDefaultSpan('\u200B'));
        list.appendChild(li);
        tb.appendChild(list);
        this._removeEmptyParagraphs(tb);
        this._setupListMutationObserver(tb);

        // Place caret inside the new li
        requestAnimationFrame(() => {
            const textNode = li.querySelector('.span-txt')?.firstChild;
            if (!textNode) return;
            const sel = window.getSelection();
            const r = document.createRange();
            r.setStart(textNode, Math.min(1, textNode.textContent.length));
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
        });
        return;
    }

    paragraphs.forEach(p => {
        // Skip truly empty paragraphs (no spans, no real content)
        const hasSpans = p.querySelector('.span-txt');
        const hasText = p.textContent.replace(/\u200B/g, '').trim().length > 0;
        if (!hasText && !hasSpans) {
            p.remove();
            return;
        }

        const li = document.createElement('li');

        // Move ALL child nodes (spans with their formatting) directly into <li>
        // This preserves every span-txt with its existing inline styles and attributes.
        while (p.firstChild) {
            li.appendChild(p.firstChild);
        }

        // Copy paragraph-level styles onto the li for alignment/line-height
        if (p.style.textAlign)  li.style.textAlign  = p.style.textAlign;
        if (p.style.lineHeight) li.style.lineHeight = p.style.lineHeight;

        // Set li color from first span so bullet marker inherits the text color.
        // The ::marker pseudo-element inherits color from <li>, not from child spans.
        const firstSpan = li.querySelector('.span-txt');
        if (firstSpan && firstSpan.style.color) {
            li.style.color = firstSpan.style.color;
        }

        list.appendChild(li);
        p.remove();
    });

    tb.appendChild(list);
    this._removeEmptyParagraphs(tb);
    this._setupListMutationObserver(tb);
}


    isInListContext(tb) {
        if (!tb) return false;
        const sel = window.getSelection();
        const node = sel?.anchorNode;
        const el = node?.nodeType === 3 ? node.parentElement : node;
        return !!el?.closest?.('ul,ol');
    }

    normalizeListStructure(tb) {
        if (!tb) return;
        const list = tb.querySelector('ul,ol');
        if (!list) return;

        list.querySelectorAll('li').forEach(li => {
            // Ensure one span-txt exists
            let span = li.querySelector(':scope > .span-txt');
            if (!span) {
                span = document.createElement('span');
                span.className = 'span-txt';
                // inherit current textbox styling
                const cs = window.getComputedStyle(tb);
                span.style.fontFamily = cs.fontFamily;
                span.style.fontSize = cs.fontSize;
                span.style.color = cs.color;
                span.style.lineHeight = cs.lineHeight || '1.2';
                span.setAttribute('originalea', '+mn-ea');
                span.setAttribute('origincs', cs.fontFamily);
                span.setAttribute('latinfont', cs.fontFamily);
                span.setAttribute('originaltxtcolor', cs.color);
                li.insertBefore(span, li.firstChild);
            }

            // Move any stray text nodes / non-span nodes into span
            const nodes = Array.from(li.childNodes);
            nodes.forEach(n => {
                if (n === span) return;
                // move text nodes and inline nodes into span
                if (n.nodeType === 3) {
                    const txt = n.textContent || '';
                    if (txt.trim().length) span.appendChild(document.createTextNode(txt));
                    n.remove();
                } else if (n.nodeType === 1) {
                    // if it's not the span-txt itself, move its text into span
                    if (!n.classList.contains('span-txt')) {
                        const txt = n.innerText || '';
                        if (txt.trim().length) span.appendChild(document.createTextNode(txt));
                        n.remove();
                    }
                }
            });

            // If span becomes empty, keep a zero-width space to hold caret safely
            if (!span.textContent || !span.textContent.length) {
                span.textContent = '\u200B';
            }
        });
    }


    // ============================================
    // ZOOM CONTROLS
    // ============================================
    zoomIn() {
        if (this.editor.zoomLevel < 124) {
            this.editor.zoomLevel += 10;
            this.applyZoom();
            this.editor.saveState();
        }
    }

    zoomOut() {
        if (this.editor.zoomLevel > 20) {
            this.editor.zoomLevel -= 10;
            this.applyZoom();
            this.editor.saveState();
        }
    }

    fitToScreen() {
        this.editor.zoomLevel = 64;
        this.applyZoom();
        this.editor.saveState();
    }

    applyZoom() {
        // Only apply zoom to slides in the main canvas, not in the preview sidebar
        const canvasSlides = document.querySelectorAll('.canvas .sli-slide');
        canvasSlides.forEach(slide => {
            slide.style.transform = `scale(${this.editor.zoomLevel / 100})`;
        });
        
        const zoomLevel = document.getElementById('zoomLevel');
        if (zoomLevel) {
            zoomLevel.textContent = this.editor.zoomLevel + '%';
        }
    }
}

// ============================================
// INITIALIZE HOME TAB CONTROLLER
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    // Wait for editor to be initialized
    setTimeout(() => {
        if (window.editor) {
            const homeTab = new HomeTabController(window.editor);
            window._homeTab = homeTab;
        }
    }, 100);
});