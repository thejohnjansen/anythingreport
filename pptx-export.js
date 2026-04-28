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

    function logTopOfMindRuns() {
        var tomEl = document.getElementById('tom');
        var tomHtml = tomEl ? tomEl.innerHTML : '';
        var tomText = tomEl ? tomEl.innerText : '';
        var tomRuns = htmlToPptRuns(tomHtml);

        var summarizedRuns = tomRuns.map(function (r, idx) {
            return {
                i: idx,
                text: r.text,
                breakLine: !!(r.options && r.options.breakLine),
                bold: !!(r.options && r.options.bold),
                italic: !!(r.options && r.options.italic),
                underline: !!(r.options && r.options.underline)
            };
        });

        console.group('Anything Report PPTX Debug: Top of Mind');
        console.log('innerHTML:', tomHtml);
        console.log('innerText:', tomText);
        console.log('runCount:', tomRuns.length);
        console.table(summarizedRuns);
        console.log('rawRuns:', tomRuns);
        console.groupEnd();
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
        var DATA_ROWS_PER_SLIDE = 12;
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
            var plHdrCells = [{ text: 'Topic', options: { bold: true, color: 'FFFFFF', fill: { color: '16213E' }, fontSize: 10, align: 'center', valign: 'middle' } }];
            for (var ps = 0; ps < PL_STAGES.length; ps++) {
                plHdrCells.push({ text: PL_STAGES[ps], options: { bold: true, color: 'FFFFFF', fill: { color: '16213E' }, fontSize: 10, align: 'center', valign: 'middle' } });
            }
            plHdrCells.push({ text: 'Here', options: { bold: true, color: 'FFFFFF', fill: { color: '16213E' }, fontSize: 10, align: 'center', valign: 'middle' } });

            var plDataRows = plData.rows.map(function (row) {
                var cells = [{ text: row.topic || '', options: { fontSize: 10, color: '222222', fill: { color: 'F8FAFC' } } }];
                for (var sc = 0; sc < 6; sc++) {
                    var isHere = (row.hereCol === sc);
                    cells.push({
                        text: row.stages[sc] || '',
                        options: {
                            fontSize: 10,
                            color: '222222',
                            bold: isHere,
                            fill: { color: isHere ? 'DBEAFE' : 'F8FAFC' }
                        }
                    });
                }
                cells.push({ text: row.hereCol >= 0 ? PL_STAGES[row.hereCol] : '\u2014', options: { fontSize: 10, color: '555555', fill: { color: 'F8FAFC' } } });
                return cells;
            });

            s2.addTable([plHdrCells].concat(plDataRows), {
                x: PPT_X,
                y: TABLE_TOP_Y,
                w: PPT_W,
                border: { type: 'solid', pt: 1, color: 'D1D5DB' },
                rowH: 0.5,
                colW: [2.5, 1.4, 1.4, 1.4, 1.4, 1.4, 1.4, 2.0]
            });
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
                            { text: riskLabel(it.midpointRisk), options: { fontSize: 10, bold: true, color: riskColor(it.midpointRisk), fill: { color: bg } } },
                            { text: it.midpointComment || '\u2014', options: { fontSize: 10, color: '4B5563', fill: { color: bg } } },
                            { text: riskLabel(it.risk), options: { fontSize: 10, bold: true, color: riskColor(it.risk), fill: { color: bg } } },
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

                for (var start = 0; start < sd.items.length; start += DATA_ROWS_PER_SLIDE) {
                    var pageItems = sd.items.slice(start, start + DATA_ROWS_PER_SLIDE);
                    var pageRows = tblRows.slice(start, start + DATA_ROWS_PER_SLIDE);
                    var pageSlide = (start === 0) ? fs : pres.addSlide();

                    if (start > 0) {
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
                }
            }
        }

        pres.writeFile({ fileName: 'anything-report.pptx' });
    }

    var btn = document.getElementById('pptxBtn');
    if (btn) btn.addEventListener('click', downloadPptx);

    var debugBtn = document.getElementById('pptxDebugBtn');
    if (debugBtn) debugBtn.addEventListener('click', logTopOfMindRuns);

    // Temporary manual hook for debugging in DevTools.
    window.__arDebugPptxTopOfMind = logTopOfMindRuns;
})();
