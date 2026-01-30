// fontFallbacks.js - Font fallback configuration and utilities


// this is sever side code so dynamic loading is not possible from here loadWebFont should be implement in frontend
const WEB_FONTS_CONFIG = {
    // Google Fonts (free)
    'Poppins': 'https://fonts.googleapis.com/css2?family=Poppins:wght@100;200;300;400;500;600;700;800;900&display=swap',
    'Roboto': 'https://fonts.googleapis.com/css2?family=Roboto:wght@100;300;400;500;700;900&display=swap',
    'Open Sans': 'https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;500;600;700;800&display=swap',
    'Montserrat': 'https://fonts.googleapis.com/css2?family=Montserrat:wght@100;200;300;400;500;600;700;800;900&display=swap',
    'Lato': 'https://fonts.googleapis.com/css2?family=Lato:wght@100;300;400;700;900&display=swap',
    'Inter': 'https://fonts.googleapis.com/css2?family=Inter:wght@100;200;300;400;500;600;700;800;900&display=swap',
    'Nunito': 'https://fonts.googleapis.com/css2?family=Nunito:wght@200;300;400;500;600;700;800;900&display=swap',
    'Raleway': 'https://fonts.googleapis.com/css2?family=Raleway:wght@100;200;300;400;500;600;700;800;900&display=swap',
    'Source Sans Pro': 'https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@200;300;400;600;700;900&display=swap',
    'PT Sans': 'https://fonts.googleapis.com/css2?family=PT+Sans:wght@400;700&display=swap',
    'Merriweather': 'https://fonts.googleapis.com/css2?family=Merriweather:wght@300;400;700;900&display=swap',
    'Playfair Display': 'https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700;800;900&display=swap',
    'Oswald': 'https://fonts.googleapis.com/css2?family=Oswald:wght@200;300;400;500;600;700&display=swap',
    'Quicksand': 'https://fonts.googleapis.com/css2?family=Quicksand:wght@300;400;500;600;700&display=swap',
};

// Track loaded fonts to avoid duplicate loading
const loadedFonts = new Set();

/**
 * Dynamically load web fonts
 * @param {string} fontFamily - The font family name
 */
function loadWebFont(fontFamily) {
    if (typeof document === 'undefined') return; // Skip on server-side
    
    const cleanFontName = fontFamily.replace(/['"]/g, '').trim();
    
    if (WEB_FONTS_CONFIG[cleanFontName] && !loadedFonts.has(cleanFontName)) {
        const linkId = `font-${cleanFontName.replace(/\s+/g, '-')}`;
        
        // Check if already loaded in DOM
        if (!document.getElementById(linkId)) {
            const link = document.createElement('link');
            link.id = linkId;
            link.rel = 'stylesheet';
            link.href = WEB_FONTS_CONFIG[cleanFontName];
            document.head.appendChild(link);
            loadedFonts.add(cleanFontName);
            console.log(`✓ Loaded web font: ${cleanFontName}`);
        }
    }
}

/**
 * Get comprehensive font fallback stack
 * @param {string} fontFamily - The font family name
 * @returns {string} - CSS font-family string with fallbacks
 */
function getFontFallbackStack(fontFamily) {
    if (!fontFamily) return 'Arial, sans-serif';
    
    const cleanFont = fontFamily.replace(/['"]/g, '').trim();
    // console.log("typeof document",typeof document);
    // Try to load web font if available
    if (typeof document !== 'undefined') {
        loadWebFont(cleanFont);
    }
    
    const fontFallbacks = {
        // Microsoft Office fonts
        'Calibri': 'Calibri, Carlito, Helvetica Neue, Helvetica, Arial, sans-serif',
        'Aptos': 'Poppins, sans-serif',
        'Cambria': 'Cambria, Georgia, Times New Roman, serif',
        'Corbel': 'Corbel, Lucida Grande, Lucida Sans Unicode, sans-serif',
        'Candara': 'Candara, Optima, Segoe UI, sans-serif',
        'Constantia': 'Constantia, Georgia, Times New Roman, serif',
        'Consolas': 'Consolas, Courier New, Monaco, monospace',
        
        // Google Fonts & Modern Web Fonts
        'Poppins': 'Poppins, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif',
        'Roboto': 'Roboto, -apple-system, BlinkMacSystemFont, Segoe UI, Arial, sans-serif',
        'Open Sans': 'Open Sans, Helvetica Neue, Helvetica, Arial, sans-serif',
        'Montserrat': 'Montserrat, Helvetica Neue, Helvetica, Arial, sans-serif',
        'Lato': 'Lato, Helvetica Neue, Helvetica, Arial, sans-serif',
        'Inter': 'Inter, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif',
        'Nunito': 'Nunito, Segoe UI, Verdana, Arial, sans-serif',
        'Raleway': 'Raleway, Helvetica Neue, Helvetica, sans-serif',
        'Source Sans Pro': 'Source Sans Pro, Segoe UI, Arial, sans-serif',
        'PT Sans': 'PT Sans, Arial, sans-serif',
        'Merriweather': 'Merriweather, Georgia, Times New Roman, serif',
        'Playfair Display': 'Playfair Display, Georgia, serif',
        'Oswald': 'Oswald, Impact, Arial Narrow, sans-serif',
        'Quicksand': 'Quicksand, Verdana, Segoe UI, sans-serif',
        
        // Classic fonts
        'Arial': 'Arial, Helvetica, Nimbus Sans L, sans-serif',
        'Arial Black': 'Arial Black, Arial Bold, Gadget, sans-serif',
        'Arial Narrow': 'Arial Narrow, Arial, sans-serif',
        'Times New Roman': 'Times New Roman, Times, Georgia, serif',
        'Helvetica': 'Helvetica, Helvetica Neue, Arial, sans-serif',
        'Helvetica Neue': 'Helvetica Neue, Helvetica, Arial, sans-serif',
        'Helvetica Neue Ltd Std': 'Helvetica Neue, Helvetica, Arial, sans-serif',
        'Helvetica Neue Ltd Std-Bd': 'Helvetica Neue, Helvetica, Arial, sans-serif',
        'Verdana': 'Verdana, Geneva, DejaVu Sans, sans-serif',
        'Georgia': 'Georgia, Times New Roman, Times, serif',
        'Courier New': 'Courier New, Courier, Lucida Sans Typewriter, monospace',
        'Tahoma': 'Tahoma, Verdana, Geneva, sans-serif',
        'Trebuchet MS': 'Trebuchet MS, Lucida Grande, sans-serif',
        'Impact': 'Impact, Arial Black, Helvetica Inserat, sans-serif',
        'Comic Sans MS': 'Comic Sans MS, Comic Sans, Chalkboard SE, cursive',
        'Palatino': 'Palatino, Palatino Linotype, Book Antiqua, Georgia, serif',
        'Garamond': 'Garamond, Times New Roman, serif',
        'Bookman': 'Bookman, Bookman Old Style, Georgia, serif',
        'Century Gothic': 'Century Gothic, Apple Gothic, sans-serif',
        'Lucida Sans': 'Lucida Sans, Lucida Grande, sans-serif',
        'Franklin Gothic': 'Franklin Gothic Medium, Franklin Gothic, Arial, sans-serif',
        
        // Custom/Corporate fonts
        'UniversCondensedLightBody': 'Univers Condensed, Arial Narrow, Arial, sans-serif',
        'Univers Condensed Light (Body)': 'Univers Condensed, Arial Narrow, Arial, sans-serif',
        'Segoe UI': 'Segoe UI, -apple-system, BlinkMacSystemFont, Arial, sans-serif',
        'San Francisco': 'San Francisco, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif',
    };
    
    // Return specific fallback if exists
    if (fontFallbacks[cleanFont]) {
        return fontFallbacks[cleanFont];
    }
    
    // Smart fallback based on font characteristics
    const lowerFont = cleanFont.toLowerCase();
    
    // Condensed/Narrow fonts
    if (lowerFont.includes('condensed') || lowerFont.includes('narrow')) {
        return `${cleanFont}, Arial Narrow, Helvetica Condensed, Arial, sans-serif`;
    }
    
    // Light/Thin fonts
    if (lowerFont.includes('light') || lowerFont.includes('thin')) {
        return `${cleanFont}, Helvetica Neue Light, Segoe UI Light, Arial, sans-serif`;
    }
    
    // Bold/Heavy fonts
    if (lowerFont.includes('bold') || lowerFont.includes('heavy') || lowerFont.includes('black')) {
        return `${cleanFont}, Arial Black, Helvetica Bold, sans-serif`;
    }
    
    // Display/Decorative fonts
    if (lowerFont.includes('display')) {
        return `${cleanFont}, Impact, Arial Black, sans-serif`;
    }
    
    // Serif fonts
    if (lowerFont.includes('serif') || lowerFont.includes('times') || 
        lowerFont.includes('garamond') || lowerFont.includes('baskerville') ||
        lowerFont.includes('palatino') || lowerFont.includes('bookman')) {
        return `${cleanFont}, Georgia, Times New Roman, serif`;
    }
    
    // Monospace fonts
    if (lowerFont.includes('mono') || lowerFont.includes('code') || 
        lowerFont.includes('courier') || lowerFont.includes('consolas') ||
        lowerFont.includes('terminal')) {
        return `${cleanFont}, Courier New, Courier, Monaco, monospace`;
    }
    
    // Script/Handwriting fonts
    if (lowerFont.includes('script') || lowerFont.includes('handwriting') || 
        lowerFont.includes('brush') || lowerFont.includes('cursive')) {
        return `${cleanFont}, cursive`;
    }
    
    // Rounded fonts
    if (lowerFont.includes('rounded') || lowerFont.includes('round')) {
        return `${cleanFont}, Verdana, Segoe UI, sans-serif`;
    }
    
    // Modern sans-serif with system font fallback (default)
    return `${cleanFont}, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica Neue, Arial, sans-serif`;
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { getFontFallbackStack, loadWebFont, WEB_FONTS_CONFIG };
}