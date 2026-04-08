/**
 * LineChartHandler.js (FINAL COMPLETE FIX)
 */

const colorHelper = require("../../api/helper/colorHelper.js");

const DEFAULT_COLORS = [
    "#4472C4", "#ED7D31", "#70AD47", "#FF0000",
    "#FFC000", "#9DC3E6", "#A9D18E", "#FF7F7F",
];

// Layout constants (px)
const PAD_LEFT = 55;
const PAD_RIGHT = 20;
const PAD_BOTTOM = 45;  // x-axis labels
const LEGEND_H = 30;  // legend row height at bottom

class LineChartHandler {
    constructor(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, masterXML = null) {
        this.graphicsNode = graphicsNode;
        this.chartXML = chartXML;
        this.chartRelsXML = chartRelsXML;
        this.chartColorsXML = chartColorsXML;
        this.chartStyleXML = chartStyleXML;
        this.themeXML = themeXML;
        this.masterXML = masterXML;
    }

    // ─── public entry point ──────────────────────────────────────────────────
    async convertChartToHTML(graphicsNode, chartXML, chartRelsXML, chartColorsXML, chartStyleXML, themeXML, zIndex = 0) {
        try {
            const position = this._getPosition(graphicsNode);
            const chartSpace = this._getChartSpace(chartXML);
            const plotArea = this._getPlotArea(chartSpace);

            if (!plotArea) {
                console.error("LineChartHandler: no plotArea found");
                return this._fallbackHTML(position, zIndex);
            }

            const lineChart = plotArea["c:lineChart"]?.[0] || plotArea["lineChart"]?.[0];
            if (!lineChart) {
                console.error("LineChartHandler: no lineChart element in plotArea");
                return this._fallbackHTML(position, zIndex);
            }

            const seriesNodes = lineChart["c:ser"] || lineChart["ser"] || [];
            const seriesData = seriesNodes.map((ser, idx) =>
                this._parseSeries(ser, idx, themeXML)
            );

            if (seriesData.length === 0) {
                return this._fallbackHTML(position, zIndex);
            }

            const categories = this._extractCategories(seriesNodes[0]);
            const title = this._extractTitle(chartSpace);
            const titleStyle = this._extractTitleStyle(chartSpace);
            const chartBackground = this._getChartBackground(chartSpace);
            const gridlineConfig = this._getGridlineConfig(plotArea);

            // Normalise lengths
            const catCount = Math.max(categories.length, ...seriesData.map(s => s.values.length));
            seriesData.forEach(s => {
                while (s.values.length < catCount) s.values.push(0);
                // Extend markers array to match
                while (s.markers.length < catCount) {
                    s.markers.push({ symbol: 'circle', size: 5 });
                }
            });
            while (categories.length < catCount) {
                categories.push(`Category ${categories.length + 1}`);
            }

            // ── Visible SVG chart ─────────────────────────────────────────
            const svgHTML = this._buildSVG(
                seriesData, categories, title, titleStyle, gridlineConfig,
                position.width, position.height
            );

            // ── Hidden data divs (PPTX round-trip) ───────────────────────
            const seriesDataHTML = seriesData.map((series, sIdx) => {
                const pointsHTML = series.values.map((val, pIdx) =>
                    `<div class="line-point"
                          data-category-index="${pIdx}"
                          data-category-label="${this._esc(String(categories[pIdx] ?? ''))}"
                          data-value="${isNaN(val) ? 0 : val}"
                          data-marker-symbol="${series.markers[pIdx]?.symbol || 'circle'}"
                          data-marker-size="${series.markers[pIdx]?.size || 5}"></div>`
                ).join("");
                return `<div class="line-series"
                              data-series-index="${sIdx}"
                              data-series-name="${this._esc(series.name)}"
                              data-series-line-color="${series.lineColor}"
                              data-series-marker-color="${series.markerColor}"
                              style="display:none;">${pointsHTML}</div>`;
            }).join("");

            const catLabelsHTML = categories.map((cat, i) =>
                `<span class="category-label" data-index="${i}"
                       style="display:none;">${this._esc(String(cat))}</span>`
            ).join("");

            return `<div class="chart-container line-chart"
                         data-chart-type="line"
                         style="position:absolute;
                                left:${position.x}px;
                                top:${position.y}px;
                                width:${position.width}px;
                                height:${position.height}px;
                                z-index:${zIndex};
                                overflow:hidden;
                                box-sizing:border-box;
                                background:${chartBackground || "#ffffff"};">
                        ${svgHTML}
                        ${seriesDataHTML}
                        <div class="category-labels" style="display:none;">${catLabelsHTML}</div>
                    </div>`;

        } catch (err) {
            console.error("LineChartHandler.convertChartToHTML error:", err);
            return "";
        }
    }

    // ─── SVG renderer ────────────────────────────────────────────────────────
    _buildSVG(seriesData, categories, title, titleStyle, gridlineConfig, totalW, totalH) {
        const titleH = title ? 36 : 10;
        const legendH = seriesData.length > 0 ? LEGEND_H : 0;

        const plotX = PAD_LEFT;
        const plotY = titleH;
        const plotW = totalW - PAD_LEFT - PAD_RIGHT;
        const plotH = totalH - titleH - PAD_BOTTOM - legendH;

        // ── Value range ───────────────────────────────────────────────────
        const allVals = seriesData.flatMap(s => s.values).filter(v => !isNaN(v));
        const rawMax = allVals.length ? Math.max(...allVals) : 6;
        const rawMin = 0; // Always start from 0 like PowerPoint
        const { axisMax, axisMin, step } = this._niceScale(rawMin, rawMax);
        const valueRange = axisMax - axisMin || 1;

        // ── Coordinate helpers with spacing from axes ────────────────────
        // Add small spacing so lines don't touch the Y-axis
        const categoryStep = categories.length > 0 ? plotW / categories.length : plotW;
        const xPos = i => plotX + (categories.length > 1
            ? (categoryStep * 0.5) + (i * categoryStep)
            : plotW / 2);
        const yPos = val => plotY + plotH - ((val - axisMin) / valueRange) * plotH;

        // ── Chart title ───────────────────────────────────────────────────
        let titleSVG = "";
        if (title) {
            const tx = (plotX + plotW / 2).toFixed(1);
            titleSVG = `<text x="${tx}" y="${(titleH - 8).toFixed(1)}"
                              text-anchor="middle"
                              font-size="${titleStyle.fontSize}"
                              font-family="${this._esc(titleStyle.fontFamily)}"
                              font-weight="${titleStyle.fontWeight}"
                              fill="${titleStyle.color}">${this._esc(title)}</text>`;
        }

        // ── Plot area background (white) ──────────────────────────────────
        const plotBg = `<rect x="${plotX}" y="${plotY}"
                              width="${plotW}" height="${plotH}"
                              fill="transparent" stroke="none"/>`;

        // ── Horizontal grid lines (for Y-axis values) ────────────────────
        // ── Horizontal grid lines (for Y-axis values) ────────────────────
        let gridLinesH = "";
        const numSteps = Math.round((axisMax - axisMin) / step);

        if (gridlineConfig.valAxis.showMajor) {
            for (let i = 0; i <= numSteps; i++) {
                const v = axisMin + i * step;
                const y = yPos(v);
                gridLinesH += `<line x1="${plotX}" y1="${y.toFixed(1)}"
                     x2="${(plotX + plotW).toFixed(1)}" y2="${y.toFixed(1)}"
                     stroke="${gridlineConfig.valAxis.majorColor}" stroke-width="0.75" opacity="1"/>`;
            }
        }

        if (gridlineConfig.valAxis.showMinor) {
            const minorSteps = 5;
            for (let i = 0; i < numSteps; i++) {
                for (let j = 1; j < minorSteps; j++) {
                    const v = axisMin + i * step + (j * step / minorSteps);
                    if (v < axisMax) {
                        const y = yPos(v);
                        gridLinesH += `<line x1="${plotX}" y1="${y.toFixed(1)}"
                             x2="${(plotX + plotW).toFixed(1)}" y2="${y.toFixed(1)}"
                             stroke="${gridlineConfig.valAxis.minorColor}" stroke-width="0.5" opacity="1"/>`;
                    }
                }
            }
        }

        // ── Vertical grid lines (for X-axis categories) ──────────────────
        let gridLinesV = "";

        if (gridlineConfig.catAxis.showMajor) {
            categories.forEach((cat, i) => {
                const x = xPos(i);
                gridLinesV += `<line x1="${x.toFixed(1)}" y1="${plotY.toFixed(1)}"
                             x2="${x.toFixed(1)}" y2="${(plotY + plotH).toFixed(1)}"
                             stroke="${gridlineConfig.catAxis.majorColor}" stroke-width="0.75" opacity="1"/>`;
            });
        }

        if (gridlineConfig.catAxis.showMinor && categories.length > 1) {
            for (let i = 1; i < categories.length; i++) {
                const x = plotX + (i * categoryStep);
                gridLinesV += `<line x1="${x.toFixed(1)}" y1="${plotY.toFixed(1)}"
                             x2="${x.toFixed(1)}" y2="${(plotY + plotH).toFixed(1)}"
                             stroke="${gridlineConfig.catAxis.minorColor}" stroke-width="0.5" opacity="1"/>`;
            }
        }

        // ── Y-axis labels ─────────────────────────────────────────────────
        let yLabels = "";
        for (let i = 0; i <= numSteps; i++) {
            const v = axisMin + i * step;
            const y = yPos(v);
            const label = Number.isInteger(v) ? v : parseFloat(v.toFixed(2));
            yLabels += `<text x="${(plotX - 6).toFixed(1)}" y="${(y + 4).toFixed(1)}"
                              text-anchor="end" font-size="11" fill="#595959">${label}</text>`;
        }

        // ── X-axis labels ─────────────────────────────────────────────────
        let xLabels = "";
        categories.forEach((cat, i) => {
            const x = xPos(i);
            xLabels += `<text x="${x.toFixed(1)}" y="${(plotY + plotH + 18).toFixed(1)}"
                              text-anchor="middle" font-size="11" fill="#595959"
                              >${this._esc(String(cat))}</text>`;
        });

        // ── Axes (drawn AFTER gridlines, BEFORE data) ────────────────────
        const axes = `
            <line x1="${plotX}" y1="${plotY.toFixed(1)}"
                  x2="${plotX}" y2="${(plotY + plotH).toFixed(1)}"
                  stroke="#BFBFBF" stroke-width="1"/>
            <line x1="${plotX}" y1="${(plotY + plotH).toFixed(1)}"
                  x2="${(plotX + plotW).toFixed(1)}" y2="${(plotY + plotH).toFixed(1)}"
                  stroke="#BFBFBF" stroke-width="1"/>`;

        // ── Series lines + markers ────────────────────────────────────────
        let seriesSVG = "";
        seriesData.forEach(series => {
            if (series.values.length === 0) return;

            const points = series.values
                .map((v, i) => `${xPos(i).toFixed(1)},${yPos(v).toFixed(1)}`)
                .join(" ");

            seriesSVG += `<polyline points="${points}"
                                    fill="none"
                                    stroke="${series.lineColor}"
                                    stroke-width="2.5"
                                    stroke-linejoin="round"
                                    stroke-linecap="round"/>`;

            // Render markers with per-point shapes
            series.values.forEach((v, i) => {
                const cx = xPos(i);
                const cy = yPos(v);
                const marker = series.markers[i] || { symbol: 'circle', size: 5 };
                const markerSVG = this._renderMarker(
                    marker.symbol,
                    cx,
                    cy,
                    marker.size,
                    series.markerColor  // ← This should already be correct
                );
                seriesSVG += markerSVG;
            });
        });

        // ── Legend ────────────────────────────────────────────────────────
        let legendSVG = "";
        if (seriesData.length > 0) {
            const legendY = plotY + plotH + PAD_BOTTOM - 2;
            const swatchW = 22;
            const textPad = 4;
            const itemGap = 14;
            const itemWidths = seriesData.map(s => swatchW + textPad + s.name.length * 6.5);
            const totalLegW = itemWidths.reduce((a, b) => a + b, 0) + itemGap * (seriesData.length - 1);
            let lx = plotX + (plotW - totalLegW) / 2;

            seriesData.forEach((series, i) => {
                const midY = legendY - 5;
                legendSVG += `
                    <line x1="${lx.toFixed(1)}" y1="${midY.toFixed(1)}"
                          x2="${(lx + swatchW).toFixed(1)}" y2="${midY.toFixed(1)}"
                          stroke="${series.lineColor}" stroke-width="2"/>
                    <circle cx="${(lx + swatchW / 2).toFixed(1)}" cy="${midY.toFixed(1)}"
        r="3.5" fill="${series.markerColor}" stroke="#fff" stroke-width="1"/>
                    <text x="${(lx + swatchW + textPad).toFixed(1)}" y="${legendY.toFixed(1)}"
                          font-size="11" fill="#595959">${this._esc(series.name)}</text>`;
                lx += itemWidths[i] + itemGap;
            });
        }

        // ── CRITICAL: SVG assembly order matters! ─────────────────────────
        // 1. Title
        // 2. Plot background (white)
        // 3. Gridlines (both H and V)
        // 4. Axes
        // 5. Data lines & markers
        // 6. Labels (X and Y)
        // 7. Legend
        return `<svg xmlns="http://www.w3.org/2000/svg"
                     width="${totalW}" height="${totalH}"
                     style="display:block; font-family:Calibri,Arial,sans-serif;">
                    ${titleSVG}
                    ${plotBg}
                    ${gridLinesH}
                    ${gridLinesV}
                    ${axes}
                    ${seriesSVG}
                    ${xLabels}
                    ${yLabels}
                    ${legendSVG}
                </svg>`;
    }

    // ── Render individual marker shape ──────────────────────────────────────
    _renderMarker(symbol, cx, cy, size, color) {
        // Size scaling factor (PPT size 5 ≈ radius 4px)
        const baseRadius = 4;
        const scaleFactor = size / 5;
        const r = baseRadius * scaleFactor;

        switch (symbol.toLowerCase()) {
            case 'circle':
                return `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}"
                                r="${r.toFixed(1)}" fill="${color}"
                                stroke="#ffffff" stroke-width="1.5"/>`;

            case 'square':
                const halfSide = r * 0.85;
                return `<rect x="${(cx - halfSide).toFixed(1)}" y="${(cy - halfSide).toFixed(1)}"
                              width="${(halfSide * 2).toFixed(1)}" height="${(halfSide * 2).toFixed(1)}"
                              fill="${color}" stroke="#ffffff" stroke-width="1.5"/>`;

            case 'diamond':
                const d = r * 1.2;
                return `<path d="M ${cx.toFixed(1)} ${(cy - d).toFixed(1)}
                                 L ${(cx + d).toFixed(1)} ${cy.toFixed(1)}
                                 L ${cx.toFixed(1)} ${(cy + d).toFixed(1)}
                                 L ${(cx - d).toFixed(1)} ${cy.toFixed(1)} Z"
                              fill="${color}" stroke="#ffffff" stroke-width="1.5"/>`;

            case 'triangle':
                const h = r * 1.3;
                const w = h * 0.866;
                return `<path d="M ${cx.toFixed(1)} ${(cy - h).toFixed(1)}
                                 L ${(cx + w).toFixed(1)} ${(cy + h * 0.5).toFixed(1)}
                                 L ${(cx - w).toFixed(1)} ${(cy + h * 0.5).toFixed(1)} Z"
                              fill="${color}" stroke="#ffffff" stroke-width="1.5"/>`;

            default:
                return `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}"
                                r="${r.toFixed(1)}" fill="${color}"
                                stroke="#ffffff" stroke-width="1.5"/>`;
        }
    }

    // ── Nice axis scale ───────────────────────────────────────────────────
    _niceScale(rawMin, rawMax) {
        if (rawMax === rawMin) { rawMax = rawMin + 1; }
        const range = rawMax - rawMin;
        const roughStep = range / 5;
        const mag = Math.pow(10, Math.floor(Math.log10(roughStep || 1)));
        const niceSteps = [1, 2, 2.5, 5, 10];
        const step = mag * (niceSteps.find(s => s * mag >= roughStep) || 10);
        const axisMin = Math.floor(rawMin / step) * step;
        const axisMax = Math.ceil(rawMax / step) * step;
        return { axisMin, axisMax, step };
    }

    // ─── XML parsers ─────────────────────────────────────────────────────────

    _getPosition(graphicsNode) {
        const frameXfrm = graphicsNode?.["p:xfrm"]?.[0];
        const off = frameXfrm?.["a:off"]?.[0]?.["$"] || {};
        const ext = frameXfrm?.["a:ext"]?.[0]?.["$"] || {};
        const EMU = 12700;
        return {
            x: Math.round(parseInt(off.x || 0, 10) / EMU),
            y: Math.round(parseInt(off.y || 0, 10) / EMU),
            width: Math.round(parseInt(ext.cx || 3200000, 10) / EMU),
            height: Math.round(parseInt(ext.cy || 2400000, 10) / EMU),
        };
    }

    _getChartSpace(chartXML) {
        return chartXML?.chartSpace?.[0] ||
            chartXML?.["c:chartSpace"]?.[0] ||
            chartXML?.chartSpace ||
            chartXML?.["c:chartSpace"] ||
            chartXML;
    }

    _getPlotArea(chartSpace) {
        const chart = chartSpace?.["c:chart"]?.[0] || chartSpace?.chart?.[0] ||
            chartSpace?.["c:chart"] || chartSpace?.chart;
        return chart?.["c:plotArea"]?.[0] || chart?.plotArea?.[0] ||
            chart?.["c:plotArea"] || chart?.plotArea;
    }

    _parseSeries(ser, idx, themeXML) {
        // Name
        const txNode = ser?.["c:tx"]?.[0] || ser?.tx?.[0];
        let name = `Series ${idx + 1}`;
        const strRef = txNode?.["c:strRef"]?.[0] || txNode?.strRef?.[0];
        if (strRef) {
            const cache = strRef?.["c:strCache"]?.[0] || strRef?.strCache?.[0];
            const pt = cache?.["c:pt"]?.[0] || cache?.pt?.[0];
            const v = pt?.["c:v"]?.[0] || pt?.v?.[0];
            if (typeof v === "string") name = v;
        }

        // Color – line stroke first, then marker spPr
        // Extract LINE color
        const spPr = ser?.["c:spPr"]?.[0] || ser?.spPr?.[0];
        const ln = spPr?.["a:ln"]?.[0] || spPr?.ln?.[0];
        const lnFill = ln?.["a:solidFill"]?.[0] || ln?.solidFill?.[0];
        let lineColor = DEFAULT_COLORS[idx % DEFAULT_COLORS.length];
        if (lnFill) {
            lineColor = this._resolveColor(lnFill, themeXML) || lineColor;
        }

        // Extract MARKER color (separate from line color)
        const marker = ser?.["c:marker"]?.[0] || ser?.marker?.[0];
        const mSpPr = marker?.["c:spPr"]?.[0] || marker?.spPr?.[0];
        const mFill = mSpPr?.["a:solidFill"]?.[0] || mSpPr?.solidFill?.[0];
        let markerColor = lineColor; // fallback to line color if no marker color
        if (mFill) {
            // Pass the entire mFill node to resolve color with modifiers
            markerColor = this._resolveColor(mFill, themeXML, mFill) || markerColor;
        }

        // Default series-level marker
        const seriesMarkerNode = ser?.["c:marker"]?.[0] || ser?.marker?.[0];
        const defaultMarker = this._parseMarkerNode(seriesMarkerNode);

        // Values
        const yNode = ser?.["c:val"]?.[0] || ser?.val?.[0];
        const numRef = yNode?.["c:numRef"]?.[0] || yNode?.numRef?.[0];
        const cache2 = numRef?.["c:numCache"]?.[0] || numRef?.numCache?.[0];
        const pts = cache2?.["c:pt"] || cache2?.pt || [];
        const ptCount = parseInt(
            cache2?.["c:ptCount"]?.[0]?.["$"]?.val ??
            cache2?.ptCount?.[0]?.["$"]?.val ??
            pts.length, 10
        );

        const values = new Array(ptCount).fill(0);
        pts.forEach(pt => {
            const i = parseInt(pt?.["$"]?.idx ?? "0", 10);
            const v = pt?.["c:v"]?.[0] ?? pt?.v?.[0];
            values[i] = v != null ? parseFloat(v) : 0;
        });

        // Per-point marker overrides (c:dPt)
        const markers = new Array(ptCount).fill(null).map(() => ({ ...defaultMarker }));
        const dPtNodes = ser?.["c:dPt"] || ser?.dPt || [];
        dPtNodes.forEach(dPt => {
            const ptIdx = parseInt(dPt?.["c:idx"]?.[0]?.["$"]?.val ?? dPt?.idx?.[0]?.["$"]?.val ?? "0", 10);
            const dPtMarkerNode = dPt?.["c:marker"]?.[0] || dPt?.marker?.[0];
            if (dPtMarkerNode && ptIdx < markers.length) {
                markers[ptIdx] = this._parseMarkerNode(dPtMarkerNode);
            }
        });

        return { name, lineColor, markerColor, values, markers };
    }

    _parseMarkerNode(markerNode) {
        if (!markerNode) {
            return { symbol: 'circle', size: 5 };
        }

        const symbolNode = markerNode?.["c:symbol"]?.[0] || markerNode?.symbol?.[0];
        const sizeNode = markerNode?.["c:size"]?.[0] || markerNode?.size?.[0];

        const symbol = symbolNode?.["$"]?.val || 'circle';
        const size = parseInt(sizeNode?.["$"]?.val ?? "5", 10);

        return { symbol, size };
    }

    _extractCategories(firstSer) {
        if (!firstSer) return [];
        const catNode = firstSer?.["c:cat"]?.[0] || firstSer?.cat?.[0];
        if (!catNode) return [];

        const strRef = catNode?.["c:strRef"]?.[0] || catNode?.strRef?.[0];
        const numRef = catNode?.["c:numRef"]?.[0] || catNode?.numRef?.[0];
        const ref = strRef || numRef;
        if (!ref) return [];

        const strCache = ref?.["c:strCache"]?.[0] || ref?.strCache?.[0];
        const numCache = ref?.["c:numCache"]?.[0] || ref?.numCache?.[0];
        const cache = strCache || numCache;
        if (!cache) return [];

        const pts = cache?.["c:pt"] || cache?.pt || [];
        const count = parseInt(
            cache?.["c:ptCount"]?.[0]?.["$"]?.val ?? pts.length, 10
        );
        const labels = new Array(count).fill("");
        pts.forEach(pt => {
            const i = parseInt(pt?.["$"]?.idx ?? "0", 10);
            const v = pt?.["c:v"]?.[0] ?? pt?.v?.[0];
            if (v != null) labels[i] = String(v);
        });
        return labels;
    }

    _extractTitle(chartSpace) {
        try {
            const chart = chartSpace?.["c:chart"]?.[0] || chartSpace?.chart?.[0] ||
                chartSpace?.["c:chart"] || chartSpace?.chart;
            const titleNode = chart?.["c:title"]?.[0] || chart?.title?.[0];
            if (!titleNode) return "";
            const tx = titleNode?.["c:tx"]?.[0] || titleNode?.tx?.[0];
            const rich = tx?.["c:rich"]?.[0] || tx?.rich?.[0];
            if (!rich) return "";
            const paras = rich?.["a:p"] || rich?.p || [];
            return paras.flatMap(p => {
                const runs = p?.["a:r"] || p?.r || [];
                return runs.map(r => r?.["a:t"]?.[0] ?? r?.t?.[0] ?? "");
            }).join("");
        } catch { return ""; }
    }

    _extractTitleStyle(chartSpace) {
        return { fontSize: 14, fontFamily: "Calibri", fontWeight: "normal", color: "#595959" };
    }

    _getChartBackground(chartSpace) {
        try {
            const chartSpaceProps = chartSpace?.["c:spPr"]?.[0] || chartSpace?.spPr?.[0];
            const fillNode = chartSpaceProps?.["a:solidFill"]?.[0] || chartSpaceProps?.solidFill?.[0];
            return this._resolveColor(fillNode, this.themeXML, fillNode) || "";
        } catch {
            return "";
        }
    }

    _getGridlineConfig(plotArea) {
        const catAx = plotArea?.["c:catAx"]?.[0] || plotArea?.catAx?.[0];
        const valAx = plotArea?.["c:valAx"]?.[0] || plotArea?.valAx?.[0];
        return {
            catAxis: this._extractAxisGridlines(catAx),
            valAxis: this._extractAxisGridlines(valAx),
        };
    }

    _extractAxisGridlines(axisNode) {
        try {
            const majorNode = axisNode?.["c:majorGridlines"]?.[0] || axisNode?.majorGridlines?.[0];
            const minorNode = axisNode?.["c:minorGridlines"]?.[0] || axisNode?.minorGridlines?.[0];
            return {
                showMajor: Boolean(majorNode),
                showMinor: Boolean(minorNode),
                majorColor: this._extractGridlineColor(majorNode) || "#D9D9D9",
                minorColor: this._extractGridlineColor(minorNode) || "#F2F2F2",
            };
        } catch {
            return {
                showMajor: false,
                showMinor: false,
                majorColor: "#D9D9D9",
                minorColor: "#F2F2F2",
            };
        }
    }

    _extractGridlineColor(gridlineNode) {
        const solidFill =
            gridlineNode?.["c:spPr"]?.[0]?.["a:ln"]?.[0]?.["a:solidFill"]?.[0] ||
            gridlineNode?.spPr?.[0]?.ln?.[0]?.solidFill?.[0];
        return this._resolveColor(solidFill, this.themeXML, solidFill) || "";
    }

    _resolveColor(solidFill, themeXML, fullNode = null) {
        try {
            const srgb = solidFill?.["a:srgbClr"]?.[0]?.["$"]?.val;
            if (srgb) return `#${srgb}`;

            const schemeClr = solidFill?.["a:schemeClr"]?.[0];
            const scheme = schemeClr?.["$"]?.val;
            if (scheme && themeXML) {
                const lumMod = schemeClr?.["a:lumMod"]?.[0]?.["$"]?.val;
                const lumOff = schemeClr?.["a:lumOff"]?.[0]?.["$"]?.val;
                const alpha = schemeClr?.["a:alpha"]?.[0]?.["$"]?.val;

                const modifiers = {};
                if (lumMod) modifiers.lumMod = parseInt(lumMod, 10);
                if (lumOff) modifiers.lumOff = parseInt(lumOff, 10);
                let color = colorHelper.resolveThemeColorHelper(scheme, themeXML, this.masterXML) || null;
                if (!color) return null;
                if (modifiers.lumMod) {
                    color = colorHelper.applyLumMod(color, modifiers.lumMod);
                }
                if (modifiers.lumOff) {
                    color = this._applyLumOff(color, modifiers.lumOff);
                }
                if (alpha !== undefined) {
                    const [r, g, b] = this._hexToRgb(color);
                    return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, Number(alpha) / 100000))})`;
                }
                return color;
            }
            return null;
        } catch { return null; }
    }

    _applyLumOff(hexColor, lumOff) {
        const [r, g, b] = this._hexToRgb(hexColor);
        const off = Number(lumOff) / 100000;
        return this._rgbToHex(
            Math.round(r + (255 * off)),
            Math.round(g + (255 * off)),
            Math.round(b + (255 * off))
        );
    }

    _hexToRgb(hexColor) {
        const value = String(hexColor).replace("#", "");
        const normalized = value.length === 3
            ? value.split("").map((c) => c + c).join("")
            : value;
        return [
            parseInt(normalized.slice(0, 2), 16),
            parseInt(normalized.slice(2, 4), 16),
            parseInt(normalized.slice(4, 6), 16),
        ];
    }

    _rgbToHex(r, g, b) {
        const clamp = (value) => Math.max(0, Math.min(255, value));
        return `#${clamp(r).toString(16).padStart(2, "0")}${clamp(g).toString(16).padStart(2, "0")}${clamp(b).toString(16).padStart(2, "0")}`;
    }

    _esc(str) {
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    _fallbackHTML(position, zIndex) {
        return `<div class="chart-container line-chart"
                     data-chart-type="line"
                     style="position:absolute;
                            left:${position?.x ?? 0}px;
                            top:${position?.y ?? 0}px;
                            width:${position?.width ?? 400}px;
                            height:${position?.height ?? 300}px;
                            z-index:${zIndex};
                            background:#ffffff;">
                </div>`;
    }
}

module.exports = LineChartHandler;
