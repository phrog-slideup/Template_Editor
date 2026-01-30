function resolveThemeFont(typeface, themeXML, fallbackTypeface) {
    const pick = (v, d) => (v && String(v).trim() ? v : d);

    let t = pick(typeface, fallbackTypeface);

    if (!t) return 'Calibri';                 // ultimate fallback
    if (!t.startsWith('+')) return t;         // concrete name already

    // Resolve +mj / +mn to theme
    const themeFontScheme = themeXML?.['a:theme']?.['a:themeElements']?.[0]?.['a:fontScheme']?.[0];
    if (!themeFontScheme) return 'Calibri'; // Default if theme not found

    const majorFont = themeFontScheme?.['a:majorFont']?.[0];
    const minorFont = themeFontScheme?.['a:minorFont']?.[0];

    // Handle different theme font references
    switch (t) {
        case '+mj-lt':  // Major font, Latin
        case '+mj':     // Major font
        case '+head':   // Heading font
            return majorFont?.['a:latin']?.[0]?.['$']?.typeface || 'Calibri Light';

        case '+mn-lt':  // Minor font, Latin
        case '+mn':     // Minor font
        case '+body':   // Body font
            return minorFont?.['a:latin']?.[0]?.['$']?.typeface || 'Calibri';

        // Handle East Asian, Complex Script, and other script types if needed
        case '+mj-ea':  // Major font, East Asian
            return majorFont?.['a:ea']?.[0]?.['$']?.typeface || 'Calibri Light';
        case '+mn-ea':  // Minor font, East Asian
            return minorFont?.['a:ea']?.[0]?.['$']?.typeface || 'Calibri';

        case '+mj-cs':  // Major font, Complex Script
            return majorFont?.['a:cs']?.[0]?.['$']?.typeface || 'Calibri Light';
        case '+mn-cs':  // Minor font, Complex Script
            return minorFont?.['a:cs']?.[0]?.['$']?.typeface || 'Calibri';

        default:
            // For unknown theme references, default to Calibri
            return 'Calibri';
    }
}

// Returns "+mj-lt" / "+mn-lt" (or a concrete font) from the master defaults
function getDefaultTypefaceFromMaster(masterXML, textKind = 'body') {
    try {
        const txStyles = masterXML?.['p:sldMaster']?.['p:txStyles']?.[0];
        if (!txStyles) return null;

        const pickLatin = (node) =>
            node?.[0]?.['a:lvl1pPr']?.[0]?.['a:defRPr']?.[0]?.['a:latin']?.[0]?.['$']?.typeface || null;

        if (textKind === 'title') return pickLatin(txStyles['p:titleStyle']);
        if (textKind === 'body') return pickLatin(txStyles['p:bodyStyle']);
        return pickLatin(txStyles['p:otherStyle']);
    } catch {
        return null;
    }
}

// Utility: infer 'title' | 'body' | 'other' from shape placeholder
function getTextKindFromShape(shapeNode) {
    const phType = shapeNode?.['p:nvSpPr']?.[0]?.['p:nvPr']?.[0]?.['p:ph']?.[0]?.['$']?.type;
    if (phType === 'title' || phType === 'ctrTitle' || phType === 'subTitle') return 'title';
    if (phType === 'body' || phType === 'obj' || phType === 'tbl') return 'body';
    return 'other';
}


module.exports = {
    resolveThemeFont,
    getDefaultTypefaceFromMaster,
    getTextKindFromShape
};
