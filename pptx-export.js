(function () {
    function stripHtml(html) {
        var d = document.createElement('div');
        d.innerHTML = html;
        return d.textContent || d.innerText || '';
    }

    function sanitizeForPpt(text) {
        if (!text) return '';
        // Keep all XML-valid code points, including supplementary-plane characters
        // (emoji live here), and drop only truly invalid XML chars.
        var out = '';
        for (var ch of text) {
            var cp = ch.codePointAt(0);
            var isValidXml =
                cp === 0x9 ||
                cp === 0xA ||
                cp === 0xD ||
                (cp >= 0x20 && cp <= 0xD7FF) ||
                (cp >= 0xE000 && cp <= 0xFFFD) ||
                (cp >= 0x10000 && cp <= 0x10FFFF);
            if (isValidXml) out += ch;
        }
        return out;
    }

    function htmlToPptRuns(html) {
        var root = document.createElement('div');
        root.innerHTML = html || '';
        var runs = [];

        function endsWithBreak() {
            if (!runs.length) return false;
            var last = runs[runs.length - 1];
            return last.text === '' && last.options && last.options.breakLine;
        }

        function pushRun(text, fmt) {
            var clean = sanitizeForPpt(text);
            if (!clean) return;
            runs.push({
                text: clean,
                options: {
                    bold: !!fmt.bold,
                    italic: !!fmt.italic,
                    underline: !!fmt.underline
                }
            });
        }

        function pushBreak() {
            // Dedicated break runs are the most reliable way to preserve
            // author-entered hard line breaks in generated PPTX text.
            if (!runs.length) return;
            runs.push({ text: '', options: { breakLine: true } });
        }

        function walk(node, fmt) {
            if (node.nodeType === Node.TEXT_NODE) {
                var txt = (node.nodeValue || '').replace(/\r\n?/g, '\n');
                var parts = txt.split('\n');
                for (var pi = 0; pi < parts.length; pi++) {
                    pushRun(parts[pi], fmt);
                    if (pi < parts.length - 1) pushBreak();
                }
                return;
            }
            if (node.nodeType !== Node.ELEMENT_NODE) return;

            var tag = node.tagName.toLowerCase();
            var nextFmt = {
                bold: fmt.bold,
                italic: fmt.italic,
                underline: fmt.underline
            };

            // If a new block starts after inline/text content, force a new line.
            // This preserves editor intent for structures like "text + <div>...".
            if ((tag === 'div' || tag === 'p' || tag === 'ul' || tag === 'ol') && runs.length && !endsWithBreak()) {
                pushBreak();
            }

            if (tag === 'strong' || tag === 'b') nextFmt.bold = true;
            if (tag === 'em' || tag === 'i') nextFmt.italic = true;
            if (tag === 'u') nextFmt.underline = true;

            if (tag === 'br') {
                pushBreak();
                return;
            }

            if (tag === 'li') {
                var parentTag = node.parentElement && node.parentElement.tagName
                    ? node.parentElement.tagName.toLowerCase()
                    : '';
                var marker = '\u2022 ';
                if (parentTag === 'ol') {
                    var idx = Array.prototype.indexOf.call(node.parentElement.children, node);
                    marker = String(idx + 1) + '. ';
                }
                pushRun(marker, nextFmt);
            }

            for (var i = 0; i < node.childNodes.length; i++) {
                walk(node.childNodes[i], nextFmt);
            }

            if (tag === 'div' || tag === 'p' || tag === 'li') {
                pushBreak();
            }
        }

        var baseFmt = { bold: false, italic: false, underline: false };
        for (var i = 0; i < root.childNodes.length; i++) {
            walk(root.childNodes[i], baseFmt);
        }

        // Avoid an extra blank line at the very end introduced by block tags.
        while (runs.length) {
            var last = runs[runs.length - 1];
            if (last.text === '' && last.options && last.options.breakLine) {
                runs.pop();
                continue;
            }
            break;
        }

        return runs;
    }

    function riskColor(risk) {
        var r = (risk || '').toLowerCase();
        if (r.includes('on track')) return '16A34A';
        if (r.includes('at risk')) return 'D97706';
        if (r.includes('off track')) return 'DC2626';
        return '888888';
    }

    function riskLabel(risk) {
        var r = (risk || '').toLowerCase();
        if (r.includes('on track')) return '\u2713 On Track';
        if (r.includes('at risk')) return '\u26a0 At Risk';
        if (r.includes('off track')) return '\u2717 Off Track';
        return '\u2014';
    }

    function pptxSlideHeader(slide, title) {
        slide.background = { color: 'FFFFFF' };
        slide.addText(title, {
            x: 0.5,
            y: 0.25,
            w: 12.6,
            h: 0.65,
            fontSize: 24,
            bold: true,
            color: '16213E',
            fontFace: 'Segoe UI',
            valign: 'top'
        });
    }

    function getState() {
        var state = window.__arPptxState;
        if (!state) throw new Error('PPTX state bridge is missing');
        return {
            hasMidpoint: !!state.getHasMidpoint(),
            lastSlides: state.getLastSlides() || [],
            PL_STAGES: state.getPipelineStages() || [],
            plData: state.loadPipeline()
        };
    }

    function normalizePipelineStageValue(value) {
        var v = (value || '').trim();
        var lower = v.toLowerCase();
        if (lower === 'complete' || lower === 'completed') return 'complete';
        return v;
    }

    function downloadPptx() {
        var state = getState();
        var hasMidpoint = state.hasMidpoint;
        var lastSlides = state.lastSlides;
        var PL_STAGES = state.PL_STAGES;
        var plData = state.plData;

        var pres = new PptxGenJS();
        pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
        var PPT_X = 0.2;
        var PPT_W = 13;
        var TABLE_TOP_Y = 1.5;
        var TABLE_ROW_H = 0.5;
        var WORKITEM_LINK_BASE = 'https://microsoft.visualstudio.com/Edge/_workitems/edit/';

        // Slide 1: Top of Mind
        var s1 = pres.addSlide();
        pptxSlideHeader(s1, 'Top of Mind');
        var tomEl = document.getElementById('tom');
        var tomHtml = tomEl ? tomEl.innerHTML : '';
        var tomRuns = htmlToPptRuns(tomHtml);
        if (tomRuns.length) {
            s1.addText(tomRuns, {
                x: PPT_X,
                y: TABLE_TOP_Y,
                w: PPT_W,
                h: 6.1,
                fontSize: 14,
                color: '1F2937',
                fontFace: 'Segoe UI',
                valign: 'top',
                wrap: true
            });
        }

        // Slide 2 (+ optional Slide 3): Pipeline — paginated
        var plTitleEl = document.querySelector('.pl-title');
        var plTitle = (plTitleEl && plTitleEl.textContent) ? plTitleEl.textContent : 'Pipeline';
        var plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

        var plChartX  = 0.20;
        var plChartY  = 1.35;
        var plChartW  = 12.93;
        var plChartH  = 5.85;
        var plTopicW  = 3.40;
        var plStageW  = plChartW - plTopicW; // shortened to give topic more room
        var plStageColW = plStageW / 6;

        // Cap at 12 rows; split >7 as 6 + remainder
        var allPlRows = (plData && plData.rows) ? plData.rows.slice(0, 12) : [];
        var plPageGroups;
        if (allPlRows.length > 7) {
            plPageGroups = [allPlRows.slice(0, 6), allPlRows.slice(6)];
        } else {
            plPageGroups = [allPlRows];
        }
        var plTotalPages = plPageGroups.length;

        for (var plPg = 0; plPg < plTotalPages; plPg++) {
            var pageRows = plPageGroups[plPg];
            var spl = pres.addSlide();
            var pageTitle = plTotalPages > 1 ? plTitle + ' (' + (plPg + 1) + '/' + plTotalPages + ')' : plTitle;
            pptxSlideHeader(spl, pageTitle);

            if (pageRows.length === 0) {
                spl.addText('No pipeline data \u2014 use the \u270e Edit button in the web view to add rows.', {
                    x: PPT_X, y: TABLE_TOP_Y, w: PPT_W, h: 0.6,
                    fontSize: 13, color: '888888', fontFace: 'Segoe UI', italic: true
                });
                continue;
            }

            // Legend (top-right, two stacked columns)
            var legX = plChartX + plChartW - 4.9;
            var legY = 0.25;
            var legColGap = 2.45;
            var legRowH = 0.22;
            var legSymW = 0.24;
            var legTxtW = 2.25;

            function addPlLegendItem(slide, col, lrow, symbol, text, opts) {
                var lx = legX + col * legColGap;
                var ly = legY + lrow * legRowH;
                if (opts.symbolType === 'dotFilled' || opts.symbolType === 'dotOpen') {
                    var d = opts.symbolDiameter || 0.18;
                    slide.addShape(pres.ShapeType.ellipse, {
                        x: lx + (legSymW - d) / 2, y: ly + 0.03, w: d, h: d,
                        line: { color: opts.symbolColor, pt: opts.symbolType === 'dotOpen' ? 2 : 1.2 },
                        fill: { color: opts.symbolType === 'dotFilled' ? opts.symbolColor : 'FFFFFF' }
                    });
                } else {
                    slide.addText(symbol, {
                        x: lx, y: ly, w: legSymW, h: 0.2,
                        align: 'center', valign: 'mid',
                        fontSize: opts.symbolSize, bold: !!opts.symbolBold,
                        color: opts.symbolColor, fontFace: opts.symbolFont || 'Segoe UI'
                    });
                }
                slide.addText(text, {
                    x: lx + legSymW + 0.04, y: ly, w: legTxtW, h: 0.2,
                    fontSize: 9, color: '4B5563', fontFace: 'Segoe UI'
                });
            }

            addPlLegendItem(spl, 0, 0, '\u25BC', 'Roughly where we are',        { symbolSize: 11, symbolBold: true, symbolColor: '2A8E2A' });
            addPlLegendItem(spl, 0, 1, '\u26A0\uFE0F', 'Risk Identified',         { symbolSize: 12, symbolColor: 'D97706', symbolFont: 'Segoe UI Emoji' });
            addPlLegendItem(spl, 1, 0, '', 'Completed',                           { symbolType: 'dotFilled', symbolColor: '2A8E2A', symbolDiameter: 0.18 });
            addPlLegendItem(spl, 1, 1, '', 'Committed for the current cycle',     { symbolType: 'dotOpen',   symbolColor: '2A8E2A', symbolDiameter: 0.18 });
            addPlLegendItem(spl, 1, 2, '', 'Planned for a future cycle',          { symbolType: 'dotOpen',   symbolColor: '9CA3AF', symbolDiameter: 0.18 });

            // Column headers
            var plHeaderY = plChartY + 0.72;
            for (var ph = 0; ph < 6; ph++) {
                spl.addText(plHeaders[ph], {
                    x: plChartX + plTopicW + ph * plStageColW,
                    y: plHeaderY, w: plStageColW, h: 0.36,
                    align: 'center', valign: 'mid', bold: true,
                    fontSize: 9, color: '666666', fontFace: 'Segoe UI'
                });
            }

            // Row geometry — larger nodes since we cap at 6 per page
            var plRowStartY = plHeaderY + 0.44;
            var plRowGap    = 0.15;
            var plN         = pageRows.length;
            var plAvail     = (plChartY + plChartH) - plRowStartY - (plN - 1) * plRowGap;
            var plRowH      = Math.min(0.95, plAvail / Math.max(plN, 1));
            var plNodeD     = Math.min(0.65, plRowH * 1.05);

            for (var r = 0; r < pageRows.length; r++) {
                var row = pageRows[r];
                var rowY    = plRowStartY + r * (plRowH + plRowGap);
                var centerY = rowY + plRowH * 0.5;

                // Topic pill
                spl.addShape(pres.ShapeType.roundRect, {
                    x: plChartX, y: rowY + plRowH * 0.10,
                    w: plTopicW - 0.18, h: plRowH * 1.05,
                    rectRadius: 0.05,
                    line: { color: 'DDDDDD', pt: 1 }, fill: { color: 'DDDDDD' }
                });
                spl.addText(row.topic || '', {
                    x: plChartX + 0.06, y: rowY + plRowH * 0.14,
                    w: plTopicW - 0.30, h: plRowH * 0.88,
                    align: 'center', valign: 'mid', bold: true,
                    fontSize: 10, color: '333333', fontFace: 'Segoe UI'
                });

                // Connector line
                spl.addShape(pres.ShapeType.line, {
                    x: plChartX + plTopicW - 0.18,
                    y: centerY,
                    w: (plChartX + plTopicW + plStageColW * 5.5 + plNodeD * 0.5) - (plChartX + plTopicW - 0.18),
                    h: 0,
                    line: { color: '999999', pt: 5 }
                });

                // Stage nodes
                for (var sc = 0; sc < 6; sc++) {
                    var cx = plChartX + plTopicW + plStageColW * (sc + 0.5);
                    var v  = normalizePipelineStageValue((row.stages && row.stages[sc]) || '');
                    var isHere = row.hereCol === sc;
                    var isWarn = row.warnCol === sc;
                    var isDone = v === 'complete';

                    if (isHere) {
                        spl.addText('\u25BC', {
                            x: cx - 0.11, y: centerY - plNodeD * 0.70 - 0.09,
                            w: 0.22, h: 0.22,
                            align: 'center', color: '2A8E2A', bold: true,
                            fontSize: 17, fontFace: 'Segoe UI'
                        });
                    }

                    if (isWarn) {
                        spl.addText('\u26A0\uFE0F', {
                            x: cx - plNodeD * 0.5 - 0.28, y: centerY - 0.14,
                            w: 0.26, h: 0.26,
                            align: 'center', valign: 'mid',
                            fontSize: 19, fontFace: 'Segoe UI Emoji'
                        });
                    }

                    if (isDone) {
                        spl.addShape(pres.ShapeType.ellipse, {
                            x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                            w: plNodeD, h: plNodeD,
                            line: { color: '2A8E2A', pt: 1.5 }, fill: { color: '2A8E2A' }
                        });
                        continue;
                    }

                    if (v) {
                        var committed = row.hereCol >= 0 && sc <= row.hereCol;
                        spl.addShape(pres.ShapeType.ellipse, {
                            x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                            w: plNodeD, h: plNodeD,
                            line: { color: committed ? '2A8E2A' : 'AAAAAA', pt: 2.5 },
                            fill: { color: 'FFFFFF' }
                        });
                        spl.addText(v, {
                            x: cx - plNodeD * 0.5 + 0.02, y: centerY - plNodeD * 0.5 + 0.02,
                            w: plNodeD - 0.04, h: plNodeD - 0.04,
                            align: 'center', valign: 'mid', fit: 'shrink',
                            fontSize: 10, bold: true,
                            color: committed ? '333333' : '666666', fontFace: 'Segoe UI'
                        });
                    }
                }
            }
        }

        // Slides 3+: Feature group slides
        if (lastSlides && lastSlides.length > 0) {
            for (var si = 0; si < lastSlides.length; si++) {
                var sd = lastSlides[si];
                var fs = pres.addSlide();
                pptxSlideHeader(fs, sd.title);

                if (sd.items.length === 0) {
                    fs.addText('No child items', {
                        x: PPT_X,
                        y: TABLE_TOP_Y,
                        w: PPT_W,
                        h: 0.5,
                        fontSize: 12,
                        color: '888888',
                        italic: true,
                        fontFace: 'Segoe UI'
                    });
                    continue;
                }

                var cols;
                var colW;
                if (hasMidpoint) {
                    cols = ['ID', 'Title', 'Midpoint Risk', 'Midpoint Details', 'Final Risk', 'Final Details'];
                    colW = [1, 3.5, 1, 3.23, 1, 3.22];
                } else {
                    cols = ['ID', 'Title', 'Risk', 'Details'];
                    colW = [1, 4, 1, 6];
                }

                var tblHdr = cols.map(function (c) {
                    return { text: c, options: { bold: true, color: 'FFFFFF', fill: { color: '16213E' }, fontSize: 10, align: 'center', valign: 'middle' } };
                });

                var tblRows = sd.items.map(function (it, idx) {
                    var bg = (idx % 2 === 0) ? 'F8FAFC' : 'FFFFFF';
                    var idCell = {
                        text: String(it.id),
                        options: {
                            fontSize: 10,
                            color: '0078d4',
                            fill: { color: bg },
                            align: 'center',
                            valign: 'middle'
                        }
                    };
                    if (hasMidpoint) {
                        return [
                            idCell,
                            { text: it.title, options: { fontSize: 10, color: '1F2937', fill: { color: bg } } },
                            { text: riskLabel(it.midpointRisk), options: { fontSize: 10, bold: true, color: riskColor(it.midpointRisk), fill: { color: bg }, align: 'center', valign: 'middle' } },
                            { text: it.midpointComment || '\u2014', options: { fontSize: 10, color: '4B5563', fill: { color: bg } } },
                            { text: riskLabel(it.risk), options: { fontSize: 10, bold: true, color: riskColor(it.risk), fill: { color: bg }, align: 'center', valign: 'middle' } },
                            { text: it.riskComment || '\u2014', options: { fontSize: 10, color: '4B5563', fill: { color: bg } } }
                        ];
                    }
                    return [
                        idCell,
                        { text: it.title, options: { fontSize: 10, color: '1F2937', fill: { color: bg } } },
                        { text: riskLabel(it.risk), options: { fontSize: 10, bold: true, color: riskColor(it.risk), fill: { color: bg } } },
                        { text: it.riskComment || '\u2014', options: { fontSize: 10, color: '4B5563', fill: { color: bg } } }
                    ];
                });

                var pageSizes;
                if (sd.items.length <= 8) {
                    pageSizes = [sd.items.length];
                } else if (sd.items.length <= 13) {
                    pageSizes = [5, sd.items.length - 5];
                } else {
                    pageSizes = [5, 5, sd.items.length - 10];
                }

                var start = 0;
                for (var p = 0; p < pageSizes.length; p++) {
                    var size = pageSizes[p];
                    var pageItems = sd.items.slice(start, start + size);
                    var pageRows = tblRows.slice(start, start + size);
                    var pageSlide = (p === 0) ? fs : pres.addSlide();

                    if (p > 0) {
                        pptxSlideHeader(pageSlide, sd.title + ' (cont.)');
                    }

                    pageSlide.addTable([tblHdr].concat(pageRows), {
                        x: PPT_X,
                        y: TABLE_TOP_Y,
                        w: PPT_W,
                        border: { type: 'solid', pt: 0.5, color: 'D1D5DB' },
                        rowH: TABLE_ROW_H,
                        colW: colW
                    });

                    // Keep ID text styling plain while preserving click-through links.
                    for (var r = 0; r < pageItems.length; r++) {
                        pageSlide.addShape(pres.ShapeType.rect, {
                            x: PPT_X,
                            y: TABLE_TOP_Y + TABLE_ROW_H * (r + 1),
                            w: colW[0],
                            h: TABLE_ROW_H,
                            line: { color: 'FFFFFF', transparency: 100, pt: 0 },
                            fill: { color: 'FFFFFF', transparency: 100 },
                            hyperlink: { url: WORKITEM_LINK_BASE + pageItems[r].id }
                        });
                    }

                    start += size;
                }
            }
        }

        pres.writeFile({ fileName: 'anything-report.pptx' });
    }

    var btn = document.getElementById('pptxBtn');
    if (btn) btn.addEventListener('click', downloadPptx);
})();
