class TransparencyHandler {
    constructor() {}

    getTransparency(shapeNode) {
        let opacity = 1; // Default fully visible
        let gradientCSS = null;

        const shapeFill = shapeNode?.["p:spPr"]?.[0];

        // ✅ Check for Solid Fill Transparency
        const solidFill = shapeFill?.["a:solidFill"]?.[0];
        if (solidFill?.["a:alpha"]) {
            opacity = parseInt(solidFill["a:alpha"][0]["$"].val, 10) / 100000;
        }

        // ✅ Check for Gradient Fill Transparency
        const gradFill = shapeFill?.["a:gradFill"]?.[0];
        if (gradFill?.["a:gsLst"]) {
            const stops = gradFill["a:gsLst"][0]["a:gs"];
            let minAlpha = 100000; // Maximum transparency value

            stops.forEach(stop => {
                const schemeClr = stop?.["a:schemeClr"]?.[0];
                if (schemeClr?.["a:alpha"]) {
                    const alphaVal = parseInt(schemeClr["a:alpha"][0]["$"].val, 10);
                    minAlpha = Math.min(minAlpha, alphaVal);
                }
            });

            if (minAlpha < 100000) {
                opacity = minAlpha / 100000;
                gradientCSS = this.getGradientCSS(gradFill);
            }
        }

        return { opacity, gradientCSS };
    }

    getGradientCSS(gradFill) {
        const stops = gradFill["a:gsLst"][0]["a:gs"] || [];
        let gradientStops = [];

        stops.forEach(stop => {
            const position = parseInt(stop["$"]?.pos, 10) / 1000;
            const schemeClr = stop?.["a:schemeClr"]?.[0]?.["$"]?.val;
            const srgbClr = stop?.["a:srgbClr"]?.[0]?.["$"]?.val;
            const color = srgbClr ? `#${srgbClr}` : schemeClr ? `var(--${schemeClr})` : "transparent";
            gradientStops.push(`${color} ${position}%`);
        });

        return `linear-gradient(180deg, ${gradientStops.join(", ")})`;
    }
}

module.exports = TransparencyHandler;
