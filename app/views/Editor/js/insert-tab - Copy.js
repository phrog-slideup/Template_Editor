/**
 * INSERT TAB FUNCTIONALITY
 * Insert text boxes, images, charts, tables, videos, audio, links, and comments
 */

class InsertTabController {
    constructor(editor) {
        this.editor = editor;
        this.init();
    }

    init() {
        this.setupEventListeners();
    }

    setupEventListeners() {
        // Text Box
        const insertTextBoxBtn = document.getElementById('insertTextBoxBtn');
        if (insertTextBoxBtn) {
            insertTextBoxBtn.addEventListener('click', () => this.insertTextBox());
        }

        // Image
        const insertImageBtn = document.getElementById('insertImageBtn');
        if (insertImageBtn) {
            insertImageBtn.addEventListener('click', () => this.insertImage());
        }

        // Chart
        const insertChartBtn = document.getElementById('insertChartBtn');
        if (insertChartBtn) {
            insertChartBtn.addEventListener('click', () => this.insertChart());
        }

        // Table
        const insertTableBtn = document.getElementById('insertTableBtn');
        if (insertTableBtn) {
            insertTableBtn.addEventListener('click', () => this.insertTable());
        }

        // Video
        const insertVideoBtn = document.getElementById('insertVideoBtn');
        if (insertVideoBtn) {
            insertVideoBtn.addEventListener('click', () => this.insertVideo());
        }

        // Audio
        const insertAudioBtn = document.getElementById('insertAudioBtn');
        if (insertAudioBtn) {
            insertAudioBtn.addEventListener('click', () => this.insertAudio());
        }

        // Link
        const insertLinkBtn = document.getElementById('insertLinkBtn');
        if (insertLinkBtn) {
            insertLinkBtn.addEventListener('click', () => this.insertLink());
        }

        // Comment
        const insertCommentBtn = document.getElementById('insertCommentBtn');
        if (insertCommentBtn) {
            insertCommentBtn.addEventListener('click', () => this.insertComment());
        }
    }

    // ============================================
    // INSERT TEXT BOX
    // ============================================
  insertTextBox() {
    // Find the currently visible slide
    const canvas = document.querySelector('.canvas');
    if (!canvas) return;
    
    const slides = canvas.querySelectorAll('.sli-slide');
    const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
    
    if (!currentSlide) {
      this.editor.showNotification('No active slide found', true);
      return;
    }

    // Generate unique ID for the text box
    const textBoxId = this.generateUniqueId();

    // Create a wrapper for the text box (resizable shape container)
    const wrapper = document.createElement('div');
    wrapper.className = 'shape text-shape';
    wrapper.id = 'rect';
    wrapper.dataset.name = 'TextBox';
    wrapper.dataset.originalColor = 'null';
    wrapper.setAttribute('originallummod', 'null');
    wrapper.setAttribute('originallumoff', 'null');
    wrapper.setAttribute('originalalpha', 'null');
    wrapper.style.cssText = `
        position: absolute;
        left: 120px;
        top: 120px;
        width: 320px;
        height: auto;
        min-height: 50px;
        background: transparent;
        opacity: 1;
        border-radius: 0px;
        border: none;
        display: block;
        transform: none;
        box-sizing: border-box;
        overflow: visible;
        z-index: 1000;
        cursor: move;
        pointer-events: auto;
    `;

    // Text box itself (contenteditable for typing)
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
        font-size: 18px;
        display: block;
        transform: none;
        opacity: 1;
        text-align: left;
        width: 100%;
        min-height: 50px;
        height: auto;
        padding: 8px;
        box-sizing: border-box;
        overflow-wrap: break-word;
        word-wrap: break-word;
        word-break: break-word;
        white-space: pre-wrap;
        line-height: 1.4;
    `;

    // Create paragraph element
    const p = document.createElement('p');
    p.style.cssText = `
        text-align: left;
        line-height: 20px;
        margin: 0;
        padding: 0;
    `;

    // Create span element with text
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

    // Build the structure: wrapper > textBox > p > span
    p.appendChild(span);
    textBox.appendChild(p);
    wrapper.appendChild(textBox);
    currentSlide.appendChild(wrapper);

    // Auto-resize functionality - wrapper grows with content
    const autoResize = () => {
        // Get the actual content height
        const contentHeight = textBox.scrollHeight;
        wrapper.style.height = Math.max(50, contentHeight + 16) + 'px'; // +16 for padding
    };

    // Add input event listener for auto-resize
    textBox.addEventListener('input', autoResize);
    textBox.addEventListener('paste', () => setTimeout(autoResize, 10));

    // Make the text box interactive (for resizing and dragging)
    this.editor.makeTextBoxInteractive(textBox);
    this.editor.selectTextBox(textBox, false);

    // Initial resize
    setTimeout(autoResize, 10);

    this.editor.saveState();
    this.editor.showNotification('Text box added');
}

  // Generate unique ID for text boxes
  generateUniqueId() {
    return Math.random().toString(36).substring(2, 12);
  }


    // ============================================
    // INSERT IMAGE
    // ============================================
    insertImage() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (file && file.type.startsWith('image/')) {
                const reader = new FileReader();
                reader.onload = (event) => {
                    this.createImageElement(event.target.result);
                };
                reader.readAsDataURL(file);
            }
        };
        
        input.click();
    }
    createImageElement(imageSrc) {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const slideContent = currentSlide.querySelector('.sli-content') || currentSlide;

        const imageContainer = document.createElement('div');
        imageContainer.className = 'image-container';
        imageContainer.setAttribute('srcrectl', '');
        imageContainer.setAttribute('srcrectr', '');
        imageContainer.setAttribute('srcrectt', '');
        imageContainer.setAttribute('srcrectb', '');
        imageContainer.dataset.name = 'InsertedImage';
        imageContainer.setAttribute('phtype', '');
        imageContainer.setAttribute('phidx', '');
        imageContainer.dataset.isLine = 'false';
        imageContainer.dataset.actualWidth = '300';
        imageContainer.dataset.actualHeight = '200';
        imageContainer.style.cssText = `
            position:absolute;
            left:120px;
            top:120px;
            width:300px;
            height:200px;
            box-shadow: 0px 0px #000000;
            transform: rotate(0deg);
            overflow:hidden;
            z-index: 20;
        `;

        const img = document.createElement('img');
        img.src = imageSrc;
        img.alt = 'InsertedImage';
        img.style.cssText = `
            position:absolute;
            width:100%;
            height:100%;
            object-fit: cover;
            opacity:1;
            filter:blur(0px) contrast(1);
            pointer-events:none;
        `;

        imageContainer.appendChild(img);
        slideContent.appendChild(imageContainer);

        // Make interactive and select
        this.editor.setupTextBoxInteractions();
        this.editor.makeElementInteractive(imageContainer, { minWidth: 20, minHeight: 20 });
        this.editor.selectElement(imageContainer, false);

        this.editor.saveState();
        this.editor.showNotification('Image added');
    }


    // ============================================
    // INSERT CHART
    // ============================================
    insertChart() {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const chartContainer = document.createElement('div');
        chartContainer.className = 'chart-element insertable-element';
        chartContainer.style.cssText = `
            position: absolute;
            top: 100px;
            left: 100px;
            width: 400px;
            height: 300px;
            background: white;
            border: 2px solid #4a90e2;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: move;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        `;

        // Create simple bar chart visualization
        const chartContent = document.createElement('div');
        chartContent.style.cssText = `
            width: 90%;
            height: 90%;
            display: flex;
            align-items: flex-end;
            justify-content: space-around;
            padding: 20px;
            pointer-events: none;
        `;

        // Create sample bars
        const barHeights = [60, 80, 50, 90, 70];
        const barColors = ['#4a90e2', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6'];
        
        barHeights.forEach((height, index) => {
            const bar = document.createElement('div');
            bar.style.cssText = `
                width: 50px;
                height: ${height}%;
                background: ${barColors[index]};
                border-radius: 4px 4px 0 0;
                transition: height 0.3s ease;
            `;
            chartContent.appendChild(bar);
        });

        chartContainer.appendChild(chartContent);
        currentSlide.appendChild(chartContainer);

        // Use universal interactive method from core
        this.editor.makeElementInteractive(chartContainer);

        this.editor.saveState();
        this.editor.showNotification('Chart added');
    }

    // ============================================
    // INSERT TABLE
    // ============================================
    insertTable() {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'table-element insertable-element';
        tableContainer.style.cssText = `
            position: absolute;
            top: 100px;
            left: 100px;
            cursor: move;
            background: white;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        `;

        const table = document.createElement('table');
        table.style.cssText = `
            border-collapse: collapse;
            width: 400px;
            pointer-events: auto;
        `;

        // Create table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        for (let i = 0; i < 3; i++) {
            const th = document.createElement('th');
            th.contentEditable = 'true';
            th.textContent = `Header ${i + 1}`;
            th.style.cssText = `
                border: 1px solid #ddd;
                padding: 12px;
                background: #4a90e2;
                color: white;
                font-weight: 600;
                text-align: left;
            `;
            headerRow.appendChild(th);
        }
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Create table body
        const tbody = document.createElement('tbody');
        for (let i = 0; i < 3; i++) {
            const row = document.createElement('tr');
            for (let j = 0; j < 3; j++) {
                const td = document.createElement('td');
                td.contentEditable = 'true';
                td.textContent = `Cell ${i + 1},${j + 1}`;
                td.style.cssText = `
                    border: 1px solid #ddd;
                    padding: 12px;
                    background: ${i % 2 === 0 ? '#f9f9f9' : 'white'};
                `;
                row.appendChild(td);
            }
            tbody.appendChild(row);
        }
        table.appendChild(tbody);

        tableContainer.appendChild(table);
        currentSlide.appendChild(tableContainer);

        // Use universal interactive method from core
        this.editor.makeElementInteractive(tableContainer);

        this.editor.saveState();
        this.editor.showNotification('Table added');
    }

    // ============================================
    // INSERT VIDEO
    // ============================================
    insertVideo() {
        const url = prompt('Enter video URL (YouTube, Vimeo, or direct video link):');
        if (!url) return;

        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const videoContainer = document.createElement('div');
        videoContainer.className = 'video-element';
        videoContainer.style.cssText = `
            position: absolute;
            top: 100px;
            left: 100px;
            width: 480px;
            height: 270px;
            cursor: move;
            border: 2px solid #4a90e2;
            border-radius: 8px;
            overflow: hidden;
            background: #000;
        `;

        // Check if it's a YouTube or Vimeo URL
        let embedUrl = url;
        if (url.includes('youtube.com') || url.includes('youtu.be')) {
            const videoId = url.split('v=')[1] || url.split('/').pop();
            embedUrl = `https://www.youtube.com/embed/${videoId}`;
        } else if (url.includes('vimeo.com')) {
            const videoId = url.split('/').pop();
            embedUrl = `https://player.vimeo.com/video/${videoId}`;
        }

        const iframe = document.createElement('iframe');
        iframe.src = embedUrl;
        iframe.style.cssText = `
            width: 100%;
            height: 100%;
            border: none;
        `;
        iframe.allowFullscreen = true;

        videoContainer.appendChild(iframe);
        currentSlide.appendChild(videoContainer);

        // Make it fully interactive with handles (resize, move, rotate, delete)
        this.editor.makeElementInteractive(videoContainer, {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 200,
            minHeight: 150
        });

        this.editor.saveState();
        this.editor.showNotification('Video added - fully interactive');
    }

    // ============================================
    // INSERT AUDIO
    // ============================================
    insertAudio() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'audio/*';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (file && file.type.startsWith('audio/')) {
                const reader = new FileReader();
                reader.onload = (event) => {
                    this.createAudioElement(event.target.result, file.name);
                };
                reader.readAsDataURL(file);
            }
        };
        
        input.click();
    }

    createAudioElement(audioSrc, fileName) {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const audioContainer = document.createElement('div');
        audioContainer.className = 'audio-element insertable-element';
        audioContainer.dataset.type = 'audio';
        audioContainer.style.cssText = `
            position: absolute;
            top: 100px;
            left: 100px;
            width: 400px;
            padding: 20px;
            background: white;
            border: 2px solid #4a90e2;
            border-radius: 8px;
            cursor: move;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        `;

        const title = document.createElement('div');
        title.textContent = fileName || 'Audio File';
        title.style.cssText = `
            margin-bottom: 10px;
            font-weight: 600;
            color: #333;
        `;

        const audio = document.createElement('audio');
        audio.src = audioSrc;
        audio.controls = true;
        audio.style.cssText = `
            width: 100%;
        `;

        audioContainer.appendChild(title);
        audioContainer.appendChild(audio);
        currentSlide.appendChild(audioContainer);

        // Make it fully interactive with handles (resize, move, rotate, delete)
        this.editor.makeElementInteractive(audioContainer, {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 200,
            minHeight: 100
        });

        this.editor.saveState();
        this.editor.showNotification('Audio added - fully interactive');
    }

    // ============================================
    // INSERT LINK
    // ============================================
    insertLink() {
        // Check if an element is selected (shape, image, textbox, etc)
        const selectedElement = this.getSelectedElement();
        
        if (selectedElement) {
            // Link for shapes, images, and other elements
            this.addLinkToElement(selectedElement);
        } else {
            // Link for selected text (traditional behavior)
            const selection = window.getSelection();
            if (!selection.rangeCount || selection.toString().trim() === '') {
                alert('Please select some text or click on a shape/image first');
                return;
            }
            this.addLinkToText(selection);
        }
    }
    
    getSelectedElement() {
        // Check for selected shape
        const canvas = document.getElementById('canvas');
        if (!canvas) return null;
        
        // Look for selected shapes, images, or other elements
        const selected = canvas.querySelector('.shape.selected, .image-container.selected, .insertable-element.selected, .sli-txt-box.selected, .chart-element.selected, .table-element.selected, .video-element.selected, .audio-element.selected, .comment-element.selected');
        
        return selected;
    }
    
    addLinkToElement(element) {
        const url = prompt('Enter URL:', 'https://');
        if (!url) return;
        
        // Store link data on element
        element.dataset.link = url;
        element.dataset.hasLink = 'true';
        
        // If it's a text element, add underline styling
        const isTextElement = element.classList.contains('sli-txt-box') || 
                             element.classList.contains('text-box') ||
                             element.querySelector('[contenteditable]');
        
        if (isTextElement) {
            // Add underline to text
            element.style.textDecoration = 'underline';
            element.style.color = element.style.color || '#4a90e2';
            
            // Find the editable content and apply underline
            const editableContent = element.querySelector('[contenteditable]') || element;
            if (editableContent.style) {
                editableContent.style.textDecoration = 'underline';
                editableContent.style.color = editableContent.style.color || '#4a90e2';
            }
            
        }
        
        // Add visual indicator (small link icon)
        this.addLinkIndicator(element);
        
        // Add click handler for Ctrl+Click
        const clickHandler = (e) => {
            if (e.ctrlKey || e.metaKey) {
                e.preventDefault();
                e.stopPropagation();
                window.open(element.dataset.link, '_blank');
                this.editor.showNotification('Opening: ' + element.dataset.link);
            }
        };
        
        // Remove old click handler if exists
        if (element._linkClickHandler) {
            element.removeEventListener('click', element._linkClickHandler);
        }
        
        // Store handler reference for future removal
        element._linkClickHandler = clickHandler;
        element.addEventListener('click', clickHandler);
        
        // Update tooltip
        const existingTitle = element.title || '';
        element.title = existingTitle ? existingTitle + ' | Link: ' + url + ' (Ctrl+Click)' : url + ' (Ctrl+Click to follow)';
        
        this.editor.saveState();
        this.editor.showNotification('Link added to element - Ctrl+Click to follow');
    }
    
    addLinkIndicator(element) {
        // Remove existing indicator if any
        const existingIndicator = element.querySelector('.link-indicator');
        if (existingIndicator) {
            existingIndicator.remove();
        }
        
        // Create link indicator icon
        const indicator = document.createElement('div');
        indicator.className = 'link-indicator';
        indicator.innerHTML = '🔗';
        indicator.style.cssText = `
            position: absolute;
            top: 5px;
            right: 5px;
            width: 24px;
            height: 24px;
            background: rgba(74, 144, 226, 0.9);
            color: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            pointer-events: none;
            z-index: 1000;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        `;
        
        element.appendChild(indicator);
    }
    
    addLinkToText(selection) {
        const url = prompt('Enter URL:', 'https://');
        if (!url) return;

        const range = selection.getRangeAt(0);
        const link = document.createElement('a');
        link.href = url;
        link.target = '_blank';
        link.title = url + ' (Ctrl+Click to follow)';
        link.style.cssText = `
            color: #4a90e2;
            text-decoration: underline;
            cursor: pointer;
        `;
        
        // Add hover effect
        link.addEventListener('mouseenter', () => {
            link.style.textDecoration = 'underline';
            link.style.color = '#2171d9';
        });
        
        link.addEventListener('mouseleave', () => {
            link.style.textDecoration = 'underline';
            link.style.color = '#4a90e2';
        });
        
        // Add Ctrl+Click to follow link
        link.addEventListener('click', (e) => {
            if (e.ctrlKey || e.metaKey) {
                window.open(link.href, '_blank');
            } else {
                e.preventDefault();
            }
        });

        try {
            const fragment = range.extractContents();
            link.appendChild(fragment);
            range.insertNode(link);
            
            range.selectNodeContents(link);
            selection.removeAllRanges();
            selection.addRange(range);
        } catch (e) {
            // console.error('Link insertion error:', e);
        }

        this.editor.saveState();
        this.editor.showNotification('Link added - Ctrl+Click to follow');
    }

    // ============================================
    // INSERT COMMENT
    // ============================================
    insertComment() {
        const canvas = document.querySelector('.canvas');
        if (!canvas) return;

        // Find the currently visible slide
        const slides = canvas.querySelectorAll('.sli-slide');
        const currentSlide = Array.from(slides).find(slide => slide.style.display !== 'none');
        
        if (!currentSlide) {
            this.editor.showNotification('No active slide found', true);
            return;
        }

        const commentText = prompt('Enter your comment:');
        if (!commentText) return;

        const commentContainer = document.createElement('div');
        commentContainer.className = 'comment-element insertable-element';
        commentContainer.dataset.type = 'comment';
        commentContainer.style.cssText = `
            position: absolute;
            top: 50px;
            right: 20px;
            width: 250px;
            padding: 15px;
            background: #fff3cd;
            border: 2px solid #ffc107;
            border-radius: 8px;
            cursor: move;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        `;

        const commentHeader = document.createElement('div');
        commentHeader.style.cssText = `
            display: flex;
            align-items: center;
            margin-bottom: 10px;
            font-weight: 600;
            color: #856404;
        `;
        commentHeader.innerHTML = `
            <span style="margin-right: 8px;">💬</span>
            <span>Comment</span>
        `;

        const commentBody = document.createElement('div');
        commentBody.style.cssText = `
            font-size: 14px;
            color: #333;
            line-height: 1.5;
        `;
        commentBody.textContent = commentText;

        const commentFooter = document.createElement('div');
        commentFooter.style.cssText = `
            margin-top: 10px;
            font-size: 12px;
            color: #666;
            text-align: right;
        `;
        commentFooter.textContent = new Date().toLocaleString();

        commentContainer.appendChild(commentHeader);
        commentContainer.appendChild(commentBody);
        commentContainer.appendChild(commentFooter);
        currentSlide.appendChild(commentContainer);

        // Make it fully interactive with handles (resize, move, rotate, delete)
        this.editor.makeElementInteractive(commentContainer, {
            resize: true,
            move: true,
            rotate: true,
            delete: true,
            minWidth: 150,
            minHeight: 100
        });

        this.editor.saveState();
        this.editor.showNotification('Comment added - fully interactive');
    }
}

// ============================================
// INITIALIZE INSERT TAB CONTROLLER
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    // Wait for editor to be initialized
    setTimeout(() => {
        if (window.editor) {
            const insertTab = new InsertTabController(window.editor);
        }
    }, 100);
});