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

        // Slide 2: Pipeline
        var plTitleEl = document.querySelector('.pl-title');
        var plTitle = (plTitleEl && plTitleEl.textContent) ? plTitleEl.textContent : 'Pipeline';
        var s2 = pres.addSlide();
        pptxSlideHeader(s2, plTitle);
        if (plData && plData.rows && plData.rows.length > 0) {
            var plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

            var chartX = 0.55;
            var chartY = 1.35;
            var chartW = 12.2;
            var chartH = 5.85;
            var topicW = 2.05;
            var stageW = chartW - topicW;
            var stageColW = stageW / 6;

            // Legend at top-right in two stacked columns.
            var legendX = chartX + chartW - 4.9;
            var legendY = 0.25;
            var legendColGap = 2.45;
            var legendRowH = 0.22;
            var legendSymbolW = 0.24;
            var legendTextW = 2.25;

            function addLegendItem(col, row, symbol, text, opts) {
                var lx = legendX + col * legendColGap;
                var ly = legendY + row * legendRowH;
                if (opts.symbolType === 'dotFilled' || opts.symbolType === 'dotOpen') {
                    var d = opts.symbolDiameter || 0.14;
                    s2.addShape(pres.ShapeType.ellipse, {
                        x: lx + (legendSymbolW - d) / 2,
                        y: ly + 0.03,
                        w: d,
                        h: d,
                        line: {
                            color: opts.symbolColor,
                            pt: opts.symbolType === 'dotOpen' ? 1.8 : 1.2
                        },
                        fill: {
                            color: opts.symbolType === 'dotFilled' ? opts.symbolColor : 'FFFFFF'
                        }
                    });
                } else {
                    s2.addText(symbol, {
                        x: lx,
                        y: ly,
                        w: legendSymbolW,
                        h: 0.2,
                        align: 'center',
                        valign: 'mid',
                        fontSize: opts.symbolSize,
                        bold: !!opts.symbolBold,
                        color: opts.symbolColor,
                        fontFace: opts.symbolFont || 'Segoe UI'
                    });
                }
                s2.addText(text, {
                    x: lx + legendSymbolW + 0.04,
                    y: ly,
                    w: legendTextW,
                    h: 0.2,
                    fontSize: 9,
                    color: '4B5563',
                    fontFace: 'Segoe UI'
                });
            }

            addLegendItem(0, 0, '\u25BC', 'Roughly where we are', { symbolSize: 11, symbolBold: true, symbolColor: '2A8E2A' });
            addLegendItem(0, 1, '\u26A0\uFE0F', 'Risk Identified', { symbolSize: 12, symbolColor: 'D97706', symbolFont: 'Segoe UI Emoji' });
            addLegendItem(1, 0, '', 'Completed', { symbolType: 'dotFilled', symbolColor: '2A8E2A', symbolDiameter: 0.16 });
            addLegendItem(1, 1, '', 'Committed for the current cycle', { symbolType: 'dotOpen', symbolColor: '2A8E2A', symbolDiameter: 0.16 });
            addLegendItem(1, 2, '', 'Planned for a future cycle', { symbolType: 'dotOpen', symbolColor: '9CA3AF', symbolDiameter: 0.16 });

            var headerY = chartY + 0.72;
            for (var ph = 0; ph < 6; ph++) {
                s2.addText(plHeaders[ph], {
                    x: chartX + topicW + ph * stageColW,
                    y: headerY,
                    w: stageColW,
                    h: 0.36,
                    align: 'center',
                    valign: 'mid',
                    bold: true,
                    fontSize: 9,
                    color: '666666',
                    fontFace: 'Segoe UI'
                });
            }

            var rows = plData.rows || [];
            var rowStartY = headerY + 0.44;
            var rowGap = 0.18;
            var availForRows = chartH - (rowStartY - chartY) - (rows.length - 1) * rowGap;
            // ~50px at 96dpi is about 0.52in; cap row height there to avoid oversized rows.
            var rowH = Math.min(0.58, availForRows / Math.max(rows.length, 1));
            var nodeD = Math.min(0.54, rowH * 0.92);

            for (var r = 0; r < rows.length; r++) {
                var row = rows[r];
                var rowY = rowStartY + r * (rowH + rowGap);
                var centerY = rowY + rowH * 0.5;

                s2.addShape(pres.ShapeType.roundRect, {
                    x: chartX,
                    y: rowY + rowH * 0.10,
                    w: topicW - 0.18,
                    h: rowH * 0.80,
                    rectRadius: 0.05,
                    line: { color: 'DDDDDD', pt: 1 },
                    fill: { color: 'DDDDDD' }
                });
                s2.addText(row.topic || '', {
                    x: chartX + 0.06,
                    y: rowY + rowH * 0.18,
                    w: topicW - 0.30,
                    h: rowH * 0.64,
                    align: 'center',
                    valign: 'mid',
                    bold: true,
                    fontSize: 10,
                    color: '333333',
                    fontFace: 'Segoe UI'
                });

                s2.addShape(pres.ShapeType.line, {
                    x: chartX + topicW - 0.18,
                    y: centerY,
                    w: (chartX + topicW + stageColW * 5.5 + nodeD * 0.5) - (chartX + topicW - 0.18),
                    h: 0,
                    line: { color: '999999', pt: 4.5 }
                });

                for (var sc = 0; sc < 6; sc++) {
                    var cx = chartX + topicW + stageColW * (sc + 0.5);
                    var v = normalizePipelineStageValue((row.stages && row.stages[sc]) || '');
                    var isHere = row.hereCol === sc;
                    var isWarn = row.warnCol === sc;
                    var isDone = v === 'complete';

                    if (isHere) {
                        s2.addText('\u25BC', {
                            x: cx - 0.08,
                            y: centerY - nodeD * 0.70,
                            w: 0.16,
                            h: 0.18,
                            align: 'center',
                            color: '2A8E2A',
                            bold: true,
                            fontSize: 15,
                            fontFace: 'Segoe UI'
                        });
                    }

                    if (isWarn) {
                        s2.addText('\u26A0\uFE0F', {
                            x: cx - nodeD * 0.5 - 0.24,
                            y: centerY - 0.12,
                            w: 0.22,
                            h: 0.22,
                            align: 'center',
                            valign: 'mid',
                            fontSize: 17,
                            fontFace: 'Segoe UI Emoji'
                        });
                    }

                    if (isDone) {
                        s2.addShape(pres.ShapeType.ellipse, {
                            x: cx - nodeD * 0.5,
                            y: centerY - nodeD * 0.5,
                            w: nodeD,
                            h: nodeD,
                            line: { color: '2A8E2A', pt: 1.5 },
                            fill: { color: '2A8E2A' }
                        });
                        continue;
                    }

                    if (v) {
                        var committed = row.hereCol >= 0 && sc <= row.hereCol;
                        s2.addShape(pres.ShapeType.ellipse, {
                            x: cx - nodeD * 0.5,
                            y: centerY - nodeD * 0.5,
                            w: nodeD,
                            h: nodeD,
                            line: { color: committed ? '2A8E2A' : 'AAAAAA', pt: 2 },
                            fill: { color: 'FFFFFF' }
                        });
                        s2.addText(v, {
                            x: cx - nodeD * 0.5 + 0.01,
                            y: centerY - nodeD * 0.5 + 0.01,
                            w: nodeD - 0.02,
                            h: nodeD - 0.02,
                            align: 'center',
                            valign: 'mid',
                            fit: 'shrink',
                            fontSize: 7,
                            bold: true,
                            color: committed ? '333333' : '666666',
                            fontFace: 'Segoe UI'
                        });
                    }
                }
            }
        } else {
            s2.addText('No pipeline data \u2014 use the \u270e Edit button in the web view to add rows.', {
                x: PPT_X,
                y: TABLE_TOP_Y,
                w: PPT_W,
                h: 0.6,
                fontSize: 13,
                color: '888888',
                fontFace: 'Segoe UI',
                italic: true
            });
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
