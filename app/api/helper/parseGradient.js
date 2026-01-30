function parseGradient(backgroundCss) {
    // First, extract the linear-gradient part of the string if present
    const gradientRegex = /linear-gradient\((.*)\)/;
    const match = backgroundCss.match(gradientRegex);

    if (!match) return null;

    // Extract the gradient arguments
    const gradientArgs = match[1];
    
    // Initialize result object
    const result = {
        angle: 0,
        colors: []
    };
    
    // Parse the color stops correctly, handling the nested parentheses in rgba values
    const parts = [];
    let currentPart = '';
    let parenthesesCount = 0;
    
    for (let i = 0; i < gradientArgs.length; i++) {
        const char = gradientArgs[i];
        
        if (char === '(') {
            parenthesesCount++;
            currentPart += char;
        } else if (char === ')') {
            parenthesesCount--;
            currentPart += char;
        } else if (char === ',' && parenthesesCount === 0) {
            parts.push(currentPart.trim());
            currentPart = '';
        } else {
            currentPart += char;
        }
    }
    
    if (currentPart.trim()) {
        parts.push(currentPart.trim());
    }
    
    // Process parts - check if first part is an angle
    let startIndex = 0;
    const firstPart = parts[0];
    
    // Check for angle (in degrees)
    if (firstPart.includes('deg')) {
        const angleMatch = firstPart.match(/([0-9.-]+)deg/);
        if (angleMatch) {
            result.angle = parseFloat(angleMatch[1]);
            startIndex = 1;
        }
    }
    // Check for directional syntax (to top, to bottom, etc.)
    else if (firstPart.startsWith('to ')) {
        switch (firstPart) {
            case 'to top': result.angle = 0; break;
            case 'to right': result.angle = 90; break;
            case 'to bottom': result.angle = 180; break;
            case 'to left': result.angle = 270; break;
            case 'to top right': case 'to right top': result.angle = 45; break;
            case 'to bottom right': case 'to right bottom': result.angle = 135; break;
            case 'to bottom left': case 'to left bottom': result.angle = 225; break;
            case 'to top left': case 'to left top': result.angle = 315; break;
            default: result.angle = 180; break;
        }
        startIndex = 1;
    }
    
    // Extract color stops
    for (let i = startIndex; i < parts.length; i++) {
        const part = parts[i].trim();
        let color, position = null;
        
        // Extract color and position
        if (part.includes(' ')) {
            // Color with position (e.g., "#00B0F0 0%")
            const spaceIndex = part.lastIndexOf(' ');
            color = part.substring(0, spaceIndex).trim();
            const posStr = part.substring(spaceIndex).trim();
            if (posStr.endsWith('%')) {
                position = parseFloat(posStr);
            }
        } else {
            // Just color
            color = part;
        }
        
        // Add to colors array
        result.colors.push({
            color: color,
            position: position !== null ? position : (i === startIndex ? 0 : 100)
        });
    }
    
    // Ensure we have at least two colors
    if (result.colors.length === 1) {
        // Duplicate the single color with different positions
        result.colors.push({
            color: result.colors[0].color,
            position: 100
        });
        result.colors[0].position = 0;
    }
    
    return result;
}

function convertColorToHex(color) {
    if (color.startsWith('#')) return color;
    
    if (color.startsWith('rgba')) {
        const rgbaRegex = /rgba\((\d+),\s*(\d+),\s*(\d+),\s*([0-9.]+)\)/;
        const match = color.match(rgbaRegex);
        
        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);
            // We ignore alpha when converting to hex
            
            return '#' + 
                r.toString(16).padStart(2, '0') + 
                g.toString(16).padStart(2, '0') + 
                b.toString(16).padStart(2, '0');
        }
    } else if (color.startsWith('rgb')) {
        const rgbRegex = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/;
        const match = color.match(rgbRegex);
        
        if (match) {
            const r = parseInt(match[1]);
            const g = parseInt(match[2]);
            const b = parseInt(match[3]);
            
            return '#' + 
                r.toString(16).padStart(2, '0') + 
                g.toString(16).padStart(2, '0') + 
                b.toString(16).padStart(2, '0');
        }
    }
    
    return '#000000'; // Default fallback
}

function getDominantColorFromGradient(gradient) {
    // If the gradient stops are defined, return the color of the first stop
    if (gradient && gradient.stops && gradient.stops.length > 0) {
        return gradient.stops[0].color;  // Dominant color is typically the first stop
    }
    return null; // If no valid gradient is provided
}

module.exports = {
    parseGradient,
    convertColorToHex,
    getDominantColorFromGradient,
}