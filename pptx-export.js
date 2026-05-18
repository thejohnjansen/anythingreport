(function () {
    var DEFAULT_THEME = {
        titleFont: 'Aptos Display',
        titleFontSize: 28,
        titleX: 2.0,
        titleY: 0.18,
        titleW: 10.6,
        titleH: 0.62,
        bodyFont: 'Aptos',
        emojiFont: 'Segoe UI Emoji',
        bgColor: 'FFFFFF',
        titleColor: '0E2841',
        bodyColor: '000000',
        mutedTextColor: '467886',
        headerFillColor: '156082',
        rowAltColor: 'F8FAFC',
        borderColor: 'D1D5DB',
        linkColor: '467886',
        pipelineLineColor: '999999',
        pipelineCompleteColor: '4EA72E',
        pipelineCommittedFillColor: 'E9F7EF',
        pipelineFutureColor: '9CA3AF',
        topicPillFillColor: 'DDDDDD',
        topicPillBorderColor: '666666',
        topicPillRadius: 0.12,
        topicPillTextColor: '0E2841',
        backgroundImageUrl: ''
    };

    function getTheme() {
        var userTheme = window.__arPptxTheme || {};
        return Object.assign({}, DEFAULT_THEME, userTheme);
    }

    function toDataUrl(url) {
        if (!url) return Promise.resolve('');
        return fetch(url)
            .then(function (res) {
                if (!res.ok) throw new Error('Failed to load background image: ' + url);
                return res.blob();
            })
            .then(function (blob) {
                return new Promise(function (resolve, reject) {
                    var reader = new FileReader();
                    reader.onloadend = function () { resolve(reader.result || ''); };
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
            })
            .catch(function () {
                return '';
            });
    }

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

        function pushRun(text, fmt, paraOpts) {
            var clean = sanitizeForPpt(text);
            if (!clean) return;
            var options = {
                bold: !!fmt.bold,
                italic: !!fmt.italic,
                underline: !!fmt.underline
            };
            if (paraOpts) {
                if (paraOpts.bullet) {
                    options.bullet = paraOpts.bullet;
                }
                if (typeof paraOpts.indentLevel === 'number') {
                    options.indentLevel = paraOpts.indentLevel;
                }
            }
            runs.push({
                text: clean,
                options: options
            });
        }

        function pushBreak() {
            // Dedicated break runs are the most reliable way to preserve
            // author-entered hard line breaks in generated PPTX text.
            if (!runs.length) return;
            runs.push({ text: '', options: { breakLine: true } });
        }

        function listDepth(liNode) {
            var depth = 0;
            var p = liNode && liNode.parentElement;
            while (p) {
                var pt = p.tagName ? p.tagName.toLowerCase() : '';
                if (pt === 'ul' || pt === 'ol') depth++;
                p = p.parentElement;
            }
            return depth;
        }

        function consumePendingParagraphOptions(state, textPart) {
            if (!state || !state.pendingParaOptions) return null;
            if (!/\S/.test(textPart || '')) return null;
            var opts = state.pendingParaOptions;
            state.pendingParaOptions = null;
            return opts;
        }

        function walk(node, fmt, state) {
            if (node.nodeType === Node.TEXT_NODE) {
                var txt = (node.nodeValue || '').replace(/\r\n?/g, '\n');
                var parts = txt.split('\n');
                for (var pi = 0; pi < parts.length; pi++) {
                    var paraOpts = consumePendingParagraphOptions(state, parts[pi]);
                    pushRun(parts[pi], fmt, paraOpts);
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
                var level = Math.max(0, listDepth(node) - 1);
                var liState = {
                    pendingParaOptions: parentTag === 'ol'
                        ? { bullet: { type: 'number' }, indentLevel: level }
                        : { bullet: true, indentLevel: level }
                };

                if (runs.length && !endsWithBreak()) {
                    pushBreak();
                }

                for (var lii = 0; lii < node.childNodes.length; lii++) {
                    walk(node.childNodes[lii], nextFmt, liState);
                }

                // Ensure an empty list item still emits a paragraph at the right list level.
                if (liState.pendingParaOptions) {
                    pushRun(' ', nextFmt, liState.pendingParaOptions);
                }

                pushBreak();
                return;
            }

            for (var i = 0; i < node.childNodes.length; i++) {
                walk(node.childNodes[i], nextFmt, state);
            }

            if (tag === 'div' || tag === 'p') {
                pushBreak();
            }
        }

        var baseFmt = { bold: false, italic: false, underline: false };
        for (var i = 0; i < root.childNodes.length; i++) {
            walk(root.childNodes[i], baseFmt, { pendingParaOptions: null });
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
        if (r.includes('at risk')) return 'B8860B';
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

    function riskFillColor(risk) {
        var r = (risk || '').toLowerCase();
        if (r.includes('on track')) return 'E9F7EF';
        if (r.includes('at risk')) return 'FFF6DB';
        if (r.includes('off track')) return 'FDECEC';
        return 'F5F5F5';
    }

    function tableNoBorder() {
        return { type: 'solid', pt: 0, color: 'FFFFFF' };
    }

    function tableHeaderBorder(theme) {
        return [
            { type: 'solid', pt: 1.2, color: theme.titleColor },
            tableNoBorder(),
            { type: 'solid', pt: 1.2, color: theme.titleColor },
            tableNoBorder()
        ];
    }

    function tableRowBottomBorder(theme) {
        return [
            tableNoBorder(),
            tableNoBorder(),
            { type: 'solid', pt: 0.8, color: theme.borderColor },
            tableNoBorder()
        ];
    }

    function pptxSlideHeader(slide, title, theme, bgImageData) {
        var headerTitle = String(title || '').replace(/^<\/>\s*/, '');
        slide.background = { color: theme.bgColor };
        if (bgImageData) {
            slide.addImage({
                data: bgImageData,
                x: 0,
                y: 0,
                w: 13.333,
                h: 7.5
            });
        }
        slide.addText(headerTitle, {
            x: theme.titleX,
            y: theme.titleY,
            w: theme.titleW,
            h: theme.titleH,
            fontSize: theme.titleFontSize,
            bold: true,
            color: theme.titleColor,
            fontFace: theme.titleFont,
            valign: 'mid'
        });
    }

    function getState() {
        var state = window.__arPptxState;
        if (!state) throw new Error('PPTX state bridge is missing');
        return {
            hasMidpoint: !!state.getHasMidpoint(),
            lastSlides: state.getLastSlides() || [],
            PL_STAGES: state.getPipelineStages() || [],
            plData: state.loadPipeline(),
            teamPipelines: state.loadTeamPipelines ? state.loadTeamPipelines() : {},
            baseSlideTitle: state.getBaseSlideTitle ? state.getBaseSlideTitle() : 'Layout',
            baseTeamName: state.getBaseTeamName ? state.getBaseTeamName() : '',
            deckFileName: state.getDeckFileName ? state.getDeckFileName() : 'anything-report',
            cycleNumber: state.getCycleNumber ? state.getCycleNumber() : ''
        };
    }

    function sanitizeFileName(name) {
        var value = String(name || '').trim();
        if (!value) return 'anything-report';
        return value.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim() || 'anything-report';
    }

    function normalizePipelineStageValue(value) {
        var v = (value || '').trim();
        var lower = v.toLowerCase();
        if (lower === 'complete' || lower === 'completed') return 'complete';
        return v;
    }

    function parsePipelineStageValue(value) {
        var raw = (value || '').trim();
        if (!raw) return { complete: false, committed: false, label: '' };

        var lower = raw.toLowerCase();
        if (lower === 'complete' || lower === 'completed') {
            return { complete: true, committed: false, label: '' };
        }
        if (lower.indexOf('complete|') === 0 || lower.indexOf('completed|') === 0) {
            var pipeIndex = raw.indexOf('|');
            var label = pipeIndex >= 0 ? raw.slice(pipeIndex + 1).trim() : '';
            return { complete: true, committed: false, label: label };
        }
        if (lower === 'committed' || lower === 'commit') {
            return { complete: false, committed: true, label: '' };
        }
        if (lower.indexOf('committed|') === 0 || lower.indexOf('commit|') === 0) {
            var pipeIndexCommitted = raw.indexOf('|');
            var committedLabel = pipeIndexCommitted >= 0 ? raw.slice(pipeIndexCommitted + 1).trim() : '';
            return { complete: false, committed: true, label: committedLabel };
        }
        return { complete: false, committed: false, label: normalizePipelineStageValue(raw) };
    }

    async function downloadPptx() {
        var state = getState();
        var theme = getTheme();
        var bgImageData = await toDataUrl(theme.backgroundImageUrl);
        var titleImageData = await toDataUrl('pptx_inspect/InitialSlideBackground.png');
        var hasMidpoint = state.hasMidpoint;
        var lastSlides = state.lastSlides;
        var PL_STAGES = state.PL_STAGES;
        var plData = state.plData;
        var teamPipelines = state.teamPipelines || {};
        var baseSlideTitle = state.baseSlideTitle || 'Layout';
        var baseTeamName = (state.baseTeamName || '').trim();
        var deckFileName = sanitizeFileName(state.deckFileName);
        // Derive cycle from deckFileName (e.g. "26C2 - Rendering Check In" → "26C2")
        // as a reliable fallback when cycleNumber hasn't been populated yet.
        var cycleNumber = (state.cycleNumber || '').trim();
        if (!cycleNumber) {
            var firstWord = deckFileName.split(' ')[0];
            if (/^\d+[A-Za-z]\d+$/.test(firstWord)) cycleNumber = firstWord;
        }
        var monthYear = new Date().toLocaleString('en-US', { month: 'long', year: 'numeric' });
        var topOfMindTitle = baseTeamName ? ('Top of Mind - ' + baseTeamName) : 'Top of Mind';
        var pipelineBaseTitle = baseTeamName ? ('Pipeline - ' + baseTeamName) : 'Pipeline';

        var pres = new PptxGenJS();
        pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
        var PPT_X = 0.2;
        var PPT_W = 13;
        var TABLE_TOP_Y = 1.5;
        var FEATURE_TABLE_TOP_Y = TABLE_TOP_Y - 0.5;
        var TABLE_ROW_H = 0.5;
        var WORKITEM_LINK_BASE = 'https://microsoft.visualstudio.com/Edge/_workitems/edit/';

        // Slide 0: Title slide
        var titleSlide = pres.addSlide();
        if (titleImageData) {
            titleSlide.addImage({ data: titleImageData, x: 0, y: 0, w: 13.333, h: 7.5 });
        } else {
            titleSlide.background = { color: '0E2841' };
        }
        if (baseTeamName) {
            titleSlide.addText(baseTeamName, {
                x: 0.4, y: 5.0, w: 9, h: 0.82,
                fontSize: 42, bold: true, color: 'FFFFFF',
                fontFace: 'Segoe UI Semibold', valign: 'mid'
            });
        }
        titleSlide.addText('End of Cycle Review', {
            x: 0.4, y: 5.8, w: 9, h: 0.65,
            fontSize: 32, bold: true, color: 'FFFFFF',
            fontFace: 'Segoe UI Semibold', valign: 'mid'
        });
        var subtitleParts = ['Web Platform'];
        if (cycleNumber) subtitleParts.push(cycleNumber);
        if (monthYear) subtitleParts.push(monthYear);
        titleSlide.addText(subtitleParts.join(' - '), {
            x: 0.4, y: 6.42, w: 10, h: 0.45,
            fontSize: 16, color: 'FFFFFF',
            fontFace: 'Segoe UI', valign: 'mid'
        });

        // Slide 1: Top of Mind
        var s1 = pres.addSlide();
        pptxSlideHeader(s1, topOfMindTitle, theme, bgImageData);
        var tomEl = document.getElementById('tom');
        var tomHtml = tomEl ? tomEl.innerHTML : '';
        var tomRuns = htmlToPptRuns(tomHtml);
        if (tomRuns.length) {
            s1.addText(tomRuns, {
                x: PPT_X,
                y: TABLE_TOP_Y,
                w: PPT_W,
                h: 6.1,
                fontSize: 20,
                color: theme.bodyColor,
                fontFace: theme.bodyFont,
                valign: 'top',
                wrap: true
            });
        }

        // Slide 2 (+ optional Slide 3): Pipeline — paginated
        var plTitle = pipelineBaseTitle;
        var plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

        var plChartX  = 0.20;
        var plChartW  = 12.93;
        var plChartBottomY = 7.05;
        var plTopicW  = 3.40;
        var plStageW  = plChartW - plTopicW; // shortened to give topic more room
        var plStageColW = plStageW / 6;
        var hasFlatTeamSlides = !!(lastSlides && lastSlides.length && lastSlides.some(function (sd) {
            return String((sd && sd.id) || '').indexOf('flat-') === 0;
        }));

        function extractTeamNameFromSlideTitle(title) {
            var t = String(title || '').trim();
            var sep = t.indexOf(' - ');
            if (sep >= 0) return t.slice(sep + 3).trim() || t;
            return t;
        }

        function addPipelinePlaceholderSlide(teamName) {
            var slide = pres.addSlide();
            pptxSlideHeader(slide, 'Pipeline - ' + teamName, theme, bgImageData);

            var plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

            // Legend (top-right, two stacked columns), mirrored from populated pipeline slides.
            var legX = plChartX + plChartW - 4.9;
            var legY = 0.25;
            var legColGap = 2.45;
            var legRowH = 0.22;
            var legSymW = 0.24;
            var legTxtW = 2.25;

            function addPlLegendItem(slideRef, col, lrow, symbol, text, opts) {
                var lx = legX + col * legColGap;
                var ly = legY + lrow * legRowH;
                if (opts.symbolType === 'dotFilled' || opts.symbolType === 'dotOpen') {
                    var d = opts.symbolDiameter || 0.18;
                    slideRef.addShape(pres.ShapeType.ellipse, {
                        x: lx + (legSymW - d) / 2, y: ly + 0.03, w: d, h: d,
                        line: { color: opts.symbolColor, pt: opts.symbolType === 'dotOpen' ? 2 : 1.2 },
                        fill: { color: opts.symbolType === 'dotFilled' ? opts.symbolColor : 'FFFFFF' }
                    });
                } else if (opts.symbolType === 'dotFilledBorder') {
                    var db = opts.symbolDiameter || 0.18;
                    slideRef.addShape(pres.ShapeType.ellipse, {
                        x: lx + (legSymW - db) / 2, y: ly + 0.03, w: db, h: db,
                        line: { color: opts.symbolColor, pt: 2 },
                        fill: { color: opts.fillColor || 'FFFFFF' }
                    });
                } else {
                    slideRef.addText(symbol, {
                        x: lx, y: ly, w: legSymW, h: 0.2,
                        align: 'center', valign: 'mid',
                        fontSize: opts.symbolSize, bold: !!opts.symbolBold,
                        color: opts.symbolColor, fontFace: opts.symbolFont || 'Segoe UI'
                    });
                }
                slideRef.addText(text, {
                    x: lx + legSymW + 0.04, y: ly, w: legTxtW, h: 0.2,
                    fontSize: 9, color: theme.mutedTextColor, fontFace: theme.bodyFont
                });
            }

            addPlLegendItem(slide, 0, 0, '\u25BC', 'Roughly where we are',        { symbolSize: 11, symbolBold: true, symbolColor: theme.pipelineCompleteColor });
            addPlLegendItem(slide, 0, 1, '\u26A0\uFE0F', 'Risk Identified',         { symbolSize: 12, symbolColor: 'D97706', symbolFont: theme.emojiFont });
            addPlLegendItem(slide, 1, 0, '', 'Completed',                           { symbolType: 'dotFilled', symbolColor: theme.pipelineCompleteColor, symbolDiameter: 0.18 });
            addPlLegendItem(slide, 1, 1, '', 'Committed for the current cycle',     { symbolType: 'dotFilledBorder', symbolColor: theme.pipelineCompleteColor, fillColor: theme.pipelineCommittedFillColor, symbolDiameter: 0.18 });
            addPlLegendItem(slide, 1, 2, '', 'Planned for a future cycle',          { symbolType: 'dotOpen',   symbolColor: theme.pipelineFutureColor, symbolDiameter: 0.18 });

            var plLegendBottomY = legY + legRowH * 3 + 0.02;
            var headerY = plLegendBottomY + 0.14;

            for (var ph = 0; ph < 6; ph++) {
                slide.addText(plHeaders[ph], {
                    x: plChartX + plTopicW + ph * plStageColW,
                    y: headerY,
                    w: plStageColW,
                    h: 0.36,
                    align: 'center',
                    valign: 'mid',
                    bold: true,
                    fontSize: 9,
                    color: '666666',
                    fontFace: theme.bodyFont
                });
            }

            var rowCount = 4;
            var rowStartY = headerY + 0.58;
            var rowGap = 0.15;
            var plFixedRows = 6;
            var plAvail = plChartBottomY - rowStartY - (plFixedRows - 1) * rowGap;
            var rowH = Math.min(0.95, plAvail / plFixedRows);
            var nodeD = Math.min(0.65, rowH * 1.05);
            var plReferenceRows = 5;
            var plReferenceAvail = plChartBottomY - rowStartY - (plReferenceRows - 1) * rowGap;
            var plReferenceRowH = Math.min(0.95, plReferenceAvail / plReferenceRows);
            var plTopicHeightTrim = 20 / 96;
            var plTopicBoxH = Math.max(0.3, plReferenceRowH * 1.05 - plTopicHeightTrim);

            for (var r = 0; r < rowCount; r++) {
                var rowY = rowStartY + r * (rowH + rowGap);
                var centerY = rowY + rowH * 0.5;

                slide.addShape(pres.ShapeType.roundRect, {
                    x: plChartX,
                    y: centerY - plTopicBoxH * 0.5,
                    w: plTopicW - 0.18,
                    h: plTopicBoxH,
                    rectRadius: theme.topicPillRadius,
                    line: { color: 'CFCFCF', pt: 1.2 },
                    fill: { color: 'E5E7EB' }
                });

                slide.addShape(pres.ShapeType.line, {
                    x: plChartX + plTopicW - 0.18,
                    y: centerY,
                    w: (plChartX + plTopicW + plStageColW * 5.5 + nodeD * 0.5) - (plChartX + plTopicW - 0.18),
                    h: 0,
                    line: { color: theme.pipelineLineColor, pt: 4 }
                });

                for (var sc = 0; sc < 6; sc++) {
                    var cx = plChartX + plTopicW + plStageColW * (sc + 0.5);
                    slide.addShape(pres.ShapeType.ellipse, {
                        x: cx - nodeD * 0.5,
                        y: centerY - nodeD * 0.5,
                        w: nodeD,
                        h: nodeD,
                        line: { color: theme.pipelineFutureColor, pt: 2.5 },
                        fill: { color: 'FFFFFF' }
                    });
                }
            }
        }

        function addPipelineDataSlides(teamName, pipelineData) {
            var pipelineRows = (pipelineData && pipelineData.rows) ? pipelineData.rows.slice(0, 12) : [];
            if (!pipelineRows.length) return false;

            var pageGroups;
            if (pipelineRows.length > 7) {
                pageGroups = [pipelineRows.slice(0, 6), pipelineRows.slice(6)];
            } else {
                pageGroups = [pipelineRows];
            }

            var title = 'Pipeline - ' + teamName;
            var totalPages = pageGroups.length;
            var plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

            for (var plPg = 0; plPg < totalPages; plPg++) {
                var pageRows = pageGroups[plPg];
                var spl = pres.addSlide();
                var pageTitle = totalPages > 1 ? title + ' (' + (plPg + 1) + '/' + totalPages + ')' : title;
                pptxSlideHeader(spl, pageTitle, theme, bgImageData);

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
                    } else if (opts.symbolType === 'dotFilledBorder') {
                        var db = opts.symbolDiameter || 0.18;
                        slide.addShape(pres.ShapeType.ellipse, {
                            x: lx + (legSymW - db) / 2, y: ly + 0.03, w: db, h: db,
                            line: { color: opts.symbolColor, pt: 2 },
                            fill: { color: opts.fillColor || 'FFFFFF' }
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
                        fontSize: 9, color: theme.mutedTextColor, fontFace: theme.bodyFont
                    });
                }

                addPlLegendItem(spl, 0, 0, '\u25BC', 'Roughly where we are',        { symbolSize: 11, symbolBold: true, symbolColor: theme.pipelineCompleteColor });
                addPlLegendItem(spl, 0, 1, '\u26A0\uFE0F', 'Risk Identified',         { symbolSize: 12, symbolColor: 'D97706', symbolFont: theme.emojiFont });
                addPlLegendItem(spl, 1, 0, '', 'Completed',                           { symbolType: 'dotFilled', symbolColor: theme.pipelineCompleteColor, symbolDiameter: 0.18 });
                addPlLegendItem(spl, 1, 1, '', 'Committed for the current cycle',     { symbolType: 'dotFilledBorder', symbolColor: theme.pipelineCompleteColor, fillColor: theme.pipelineCommittedFillColor, symbolDiameter: 0.18 });
                addPlLegendItem(spl, 1, 2, '', 'Planned for a future cycle',          { symbolType: 'dotOpen',   symbolColor: theme.pipelineFutureColor, symbolDiameter: 0.18 });

                var plLegendBottomY = legY + legRowH * 3 + 0.02;
                var plHeaderY = plLegendBottomY + 0.14;
                for (var ph = 0; ph < 6; ph++) {
                    spl.addText(plHeaders[ph], {
                        x: plChartX + plTopicW + ph * plStageColW,
                        y: plHeaderY, w: plStageColW, h: 0.36,
                        align: 'center', valign: 'mid', bold: true,
                        fontSize: 9, color: '666666', fontFace: theme.bodyFont
                    });
                }

                var plRowStartY = plHeaderY + 0.58;
                var plRowGap    = 0.15;
                var plFixedRows = 6;
                var plAvail     = plChartBottomY - plRowStartY - (plFixedRows - 1) * plRowGap;
                var plRowH      = Math.min(0.95, plAvail / plFixedRows);
                var plNodeD     = Math.min(0.65, plRowH * 1.05);
                var plReferenceRows = 5;
                var plReferenceAvail = plChartBottomY - plRowStartY - (plReferenceRows - 1) * plRowGap;
                var plReferenceRowH = Math.min(0.95, plReferenceAvail / plReferenceRows);
                var plTopicHeightTrim = 20 / 96;
                var plTopicBoxH = Math.max(0.3, plReferenceRowH * 1.05 - plTopicHeightTrim);
                var plTopicTextH = Math.max(0.42, plReferenceRowH * 0.88 - plTopicHeightTrim);

                for (var r = 0; r < pageRows.length; r++) {
                    var row = pageRows[r];
                    var rowY    = plRowStartY + r * (plRowH + plRowGap);
                    var centerY = rowY + plRowH * 0.5;

                    spl.addShape(pres.ShapeType.roundRect, {
                        x: plChartX, y: centerY - plTopicBoxH * 0.5,
                        w: plTopicW - 0.18, h: plTopicBoxH,
                        rectRadius: theme.topicPillRadius,
                        line: { color: theme.topicPillBorderColor, pt: 1.2 }, fill: { color: theme.topicPillFillColor }
                    });
                    spl.addText(row.topic || '', {
                        x: plChartX + 0.06, y: centerY - plTopicTextH * 0.5,
                        w: plTopicW - 0.30, h: plTopicTextH,
                        align: 'center', valign: 'mid', bold: true,
                        fontSize: 18, color: theme.topicPillTextColor, fontFace: theme.bodyFont
                    });

                    spl.addShape(pres.ShapeType.line, {
                        x: plChartX + plTopicW - 0.18,
                        y: centerY,
                        w: (plChartX + plTopicW + plStageColW * 5.5 + plNodeD * 0.5) - (plChartX + plTopicW - 0.18),
                        h: 0,
                        line: { color: theme.pipelineLineColor, pt: 5 }
                    });

                    for (var sc = 0; sc < 6; sc++) {
                        var cx = plChartX + plTopicW + plStageColW * (sc + 0.5);
                        var parsedStage = parsePipelineStageValue((row.stages && row.stages[sc]) || '');
                        var v  = parsedStage.label;
                        var isHere = row.hereCol === sc;
                        var isWarn = row.warnCol === sc;
                        var isDone = parsedStage.complete;
                        var isCommitted = !!parsedStage.committed;

                        if (isHere) {
                            spl.addShape(pres.ShapeType.downArrow, {
                                x: cx - plNodeD * 0.5 - 0.25, y: centerY - 0.27,
                                w: 0.12, h: 0.24,
                                line: { color: theme.pipelineCompleteColor, pt: 1.3 },
                                fill: { color: theme.pipelineCompleteColor }
                            });
                        }

                        if (isWarn) {
                            spl.addText('\u26A0\uFE0F', {
                                x: cx - plNodeD * 0.5 - 0.61, y: centerY - 0.24,
                                w: 0.32, h: 0.32,
                                align: 'center', valign: 'mid',
                                fontSize: 24, fontFace: theme.emojiFont
                            });
                        }

                        if (isDone) {
                            spl.addShape(pres.ShapeType.ellipse, {
                                x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                                w: plNodeD, h: plNodeD,
                                line: { color: theme.pipelineCompleteColor, pt: 1.5 }, fill: { color: theme.pipelineCompleteColor }
                            });
                            if (v) {
                                spl.addText(v, {
                                    x: cx - plNodeD * 0.5 + 0.02, y: centerY - plNodeD * 0.5 + 0.02,
                                    w: plNodeD - 0.04, h: plNodeD - 0.04,
                                    align: 'center', valign: 'mid', fit: 'shrink',
                                    fontSize: 12, bold: true,
                                    color: '000000', fontFace: theme.bodyFont
                                });
                            }
                            continue;
                        }

                        if (v || isCommitted) {
                            var committed = isCommitted;
                            spl.addShape(pres.ShapeType.ellipse, {
                                x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                                w: plNodeD, h: plNodeD,
                                line: { color: committed ? theme.pipelineCompleteColor : theme.pipelineFutureColor, pt: 2.5 },
                                fill: { color: committed ? theme.pipelineCommittedFillColor : 'FFFFFF' }
                            });
                            if (v) {
                                spl.addText(v, {
                                    x: cx - plNodeD * 0.5 + 0.02, y: centerY - plNodeD * 0.5 + 0.02,
                                    w: plNodeD - 0.04, h: plNodeD - 0.04,
                                    align: 'center', valign: 'mid', fit: 'shrink',
                                    fontSize: 12, bold: true,
                                    color: committed ? theme.topicPillTextColor : '666666', fontFace: theme.bodyFont
                                });
                            }
                        }
                    }
                }
            }

            return true;
        }

        // Cap at 12 rows; split >7 as 6 + remainder
        var allPlRows = (plData && plData.rows) ? plData.rows.slice(0, 12) : [];
        var plPageGroups;
        if (allPlRows.length > 7) {
            plPageGroups = [allPlRows.slice(0, 6), allPlRows.slice(6)];
        } else {
            plPageGroups = [allPlRows];
        }
        var plTotalPages = plPageGroups.length;

        if (!hasFlatTeamSlides) {
        for (var plPg = 0; plPg < plTotalPages; plPg++) {
            var pageRows = plPageGroups[plPg];
            var spl = pres.addSlide();
            var pageTitle = plTotalPages > 1 ? plTitle + ' (' + (plPg + 1) + '/' + plTotalPages + ')' : plTitle;
            pptxSlideHeader(spl, pageTitle, theme, bgImageData);

            if (pageRows.length === 0) {
                spl.addText('No pipeline data \u2014 use the \u270e Edit button in the web view to add rows.', {
                    x: PPT_X, y: TABLE_TOP_Y, w: PPT_W, h: 0.6,
                    fontSize: 13, color: '888888', fontFace: theme.bodyFont, italic: true
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
                } else if (opts.symbolType === 'dotFilledBorder') {
                    var db = opts.symbolDiameter || 0.18;
                    slide.addShape(pres.ShapeType.ellipse, {
                        x: lx + (legSymW - db) / 2, y: ly + 0.03, w: db, h: db,
                        line: { color: opts.symbolColor, pt: 2 },
                        fill: { color: opts.fillColor || 'FFFFFF' }
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
                    fontSize: 9, color: theme.mutedTextColor, fontFace: theme.bodyFont
                });
            }

            addPlLegendItem(spl, 0, 0, '\u25BC', 'Roughly where we are',        { symbolSize: 11, symbolBold: true, symbolColor: theme.pipelineCompleteColor });
            addPlLegendItem(spl, 0, 1, '\u26A0\uFE0F', 'Risk Identified',         { symbolSize: 12, symbolColor: 'D97706', symbolFont: theme.emojiFont });
            addPlLegendItem(spl, 1, 0, '', 'Completed',                           { symbolType: 'dotFilled', symbolColor: theme.pipelineCompleteColor, symbolDiameter: 0.18 });
            addPlLegendItem(spl, 1, 1, '', 'Committed for the current cycle',     { symbolType: 'dotFilledBorder', symbolColor: theme.pipelineCompleteColor, fillColor: theme.pipelineCommittedFillColor, symbolDiameter: 0.18 });
            addPlLegendItem(spl, 1, 2, '', 'Planned for a future cycle',          { symbolType: 'dotOpen',   symbolColor: theme.pipelineFutureColor, symbolDiameter: 0.18 });

            // Anchor the pipeline directly below the legend instead of centering it vertically.
            var plLegendBottomY = legY + legRowH * 3 + 0.02;
            var plHeaderY = plLegendBottomY + 0.14;
            for (var ph = 0; ph < 6; ph++) {
                spl.addText(plHeaders[ph], {
                    x: plChartX + plTopicW + ph * plStageColW,
                    y: plHeaderY, w: plStageColW, h: 0.36,
                    align: 'center', valign: 'mid', bold: true,
                    fontSize: 9, color: '666666', fontFace: theme.bodyFont
                });
            }

            // Row geometry — larger nodes since we cap at 6 per page
            var plRowStartY = plHeaderY + 0.58;
            var plRowGap    = 0.15;
            var plFixedRows = 6;
            var plAvail     = plChartBottomY - plRowStartY - (plFixedRows - 1) * plRowGap;
            var plRowH      = Math.min(0.95, plAvail / plFixedRows);
            var plNodeD     = Math.min(0.65, plRowH * 1.05);
            var plReferenceRows = 5;
            var plReferenceAvail = plChartBottomY - plRowStartY - (plReferenceRows - 1) * plRowGap;
            var plReferenceRowH = Math.min(0.95, plReferenceAvail / plReferenceRows);
            var plTopicHeightTrim = 20 / 96;
            var plTopicBoxH = Math.max(0.3, plReferenceRowH * 1.05 - plTopicHeightTrim);
            var plTopicTextH = Math.max(0.42, plReferenceRowH * 0.88 - plTopicHeightTrim);

            for (var r = 0; r < pageRows.length; r++) {
                var row = pageRows[r];
                var rowY    = plRowStartY + r * (plRowH + plRowGap);
                var centerY = rowY + plRowH * 0.5;

                // Keep topic pills at the five-row reference size so they stay readable.
                spl.addShape(pres.ShapeType.roundRect, {
                    x: plChartX, y: centerY - plTopicBoxH * 0.5,
                    w: plTopicW - 0.18, h: plTopicBoxH,
                    rectRadius: theme.topicPillRadius,
                    line: { color: theme.topicPillBorderColor, pt: 1.2 }, fill: { color: theme.topicPillFillColor }
                });
                spl.addText(row.topic || '', {
                    x: plChartX + 0.06, y: centerY - plTopicTextH * 0.5,
                    w: plTopicW - 0.30, h: plTopicTextH,
                    align: 'center', valign: 'mid', bold: true,
                    fontSize: 18, color: theme.topicPillTextColor, fontFace: theme.bodyFont
                });

                // Connector line
                spl.addShape(pres.ShapeType.line, {
                    x: plChartX + plTopicW - 0.18,
                    y: centerY,
                    w: (plChartX + plTopicW + plStageColW * 5.5 + plNodeD * 0.5) - (plChartX + plTopicW - 0.18),
                    h: 0,
                    line: { color: theme.pipelineLineColor, pt: 5 }
                });

                // Stage nodes
                for (var sc = 0; sc < 6; sc++) {
                    var cx = plChartX + plTopicW + plStageColW * (sc + 0.5);
                    var parsedStage = parsePipelineStageValue((row.stages && row.stages[sc]) || '');
                    var v  = parsedStage.label;
                    var isHere = row.hereCol === sc;
                    var isWarn = row.warnCol === sc;
                    var isDone = parsedStage.complete;
                    var isCommitted = !!parsedStage.committed;

                    if (isHere) {
                        spl.addShape(pres.ShapeType.downArrow, {
                            x: cx - plNodeD * 0.5 - 0.25, y: centerY - 0.27,
                            w: 0.12, h: 0.24,
                            line: { color: theme.pipelineCompleteColor, pt: 1.3 },
                            fill: { color: theme.pipelineCompleteColor }
                        });
                    }

                    if (isWarn) {
                        spl.addText('\u26A0\uFE0F', {
                            x: cx - plNodeD * 0.5 - 0.61, y: centerY - 0.24,
                            w: 0.32, h: 0.32,
                            align: 'center', valign: 'mid',
                            fontSize: 24, fontFace: theme.emojiFont
                        });
                    }

                    if (isDone) {
                        spl.addShape(pres.ShapeType.ellipse, {
                            x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                            w: plNodeD, h: plNodeD,
                            line: { color: theme.pipelineCompleteColor, pt: 1.5 }, fill: { color: theme.pipelineCompleteColor }
                        });
                        if (v) {
                            spl.addText(v, {
                                x: cx - plNodeD * 0.5 + 0.02, y: centerY - plNodeD * 0.5 + 0.02,
                                w: plNodeD - 0.04, h: plNodeD - 0.04,
                                align: 'center', valign: 'mid', fit: 'shrink',
                                fontSize: 12, bold: true,
                                color: '000000', fontFace: theme.bodyFont
                            });
                        }
                        continue;
                    }

                    if (v || isCommitted) {
                        var committed = isCommitted;
                        spl.addShape(pres.ShapeType.ellipse, {
                            x: cx - plNodeD * 0.5, y: centerY - plNodeD * 0.5,
                            w: plNodeD, h: plNodeD,
                            line: { color: committed ? theme.pipelineCompleteColor : theme.pipelineFutureColor, pt: 2.5 },
                            fill: { color: committed ? theme.pipelineCommittedFillColor : 'FFFFFF' }
                        });
                        if (v) {
                            spl.addText(v, {
                                x: cx - plNodeD * 0.5 + 0.02, y: centerY - plNodeD * 0.5 + 0.02,
                                w: plNodeD - 0.04, h: plNodeD - 0.04,
                                align: 'center', valign: 'mid', fit: 'shrink',
                                fontSize: 12, bold: true,
                                color: committed ? theme.topicPillTextColor : '666666', fontFace: theme.bodyFont
                            });
                        }
                    }
                }
            }
        }
        }

        // Slides 3+: Feature group slides
        if (lastSlides && lastSlides.length > 0) {
            for (var si = 0; si < lastSlides.length; si++) {
                var sd = lastSlides[si];
                var fs = pres.addSlide();
                pptxSlideHeader(fs, sd.title, theme, bgImageData);

                if (sd.items.length === 0) {
                    fs.addText('No child items', {
                        x: PPT_X,
                        y: FEATURE_TABLE_TOP_Y,
                        w: PPT_W,
                        h: 0.5,
                        fontSize: 12,
                        color: '888888',
                        italic: true,
                        fontFace: theme.bodyFont
                    });
                    continue;
                }

                var cols;
                var colW;
                if (hasMidpoint) {
                    cols = ['ID', 'Title', 'Midpoint Risk', 'Midpoint Details', 'Final Risk', 'Final Details'];
                    colW = [1, 3.2, 1.25, 2.95, 1.25, 3.3];
                } else {
                    cols = ['ID', 'Title', 'Risk', 'Details'];
                    colW = [1, 3.7, 1.35, 5.95];
                }

                var hdrBorder = tableHeaderBorder(theme);
                var rowBorder = tableRowBottomBorder(theme);

                var tblHdr = cols.map(function (c) {
                    return { text: c, options: { bold: true, color: '000000', fill: { color: 'FFFFFF' }, fontSize: 12, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: hdrBorder } };
                });

                var tblRows = sd.items.map(function (it) {
                    var finalBg = riskFillColor(it.risk);
                    var midpointBg = riskFillColor(it.midpointRisk);
                    var idCell = {
                        text: String(it.id),
                        options: {
                            fontSize: 12,
                            color: '000000',
                            fill: { color: 'FFFFFF' },
                            align: 'center',
                            valign: 'middle',
                            fontFace: theme.bodyFont,
                            border: rowBorder,
                            hyperlink: { url: WORKITEM_LINK_BASE + it.id }
                        }
                    };
                    if (hasMidpoint) {
                        return [
                            idCell,
                            { text: it.title, options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: rowBorder } },
                            { text: riskLabel(it.midpointRisk), options: { fontSize: 12, bold: true, color: riskColor(it.midpointRisk), fill: { color: midpointBg }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: rowBorder } },
                            { text: it.midpointComment || '\u2014', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: rowBorder } },
                            { text: riskLabel(it.risk), options: { fontSize: 12, bold: true, color: riskColor(it.risk), fill: { color: finalBg }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: rowBorder } },
                            { text: it.riskComment || '\u2014', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: rowBorder } }
                        ];
                    }
                    return [
                        idCell,
                        { text: it.title, options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: rowBorder } },
                        { text: riskLabel(it.risk), options: { fontSize: 12, bold: true, color: riskColor(it.risk), fill: { color: finalBg }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: rowBorder } },
                        { text: it.riskComment || '\u2014', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: rowBorder } }
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
                        pptxSlideHeader(pageSlide, sd.title + ' (cont.)', theme, bgImageData);
                    }

                    pageSlide.addTable([tblHdr].concat(pageRows), {
                        x: PPT_X,
                        y: FEATURE_TABLE_TOP_Y,
                        w: PPT_W,
                        border: tableNoBorder(),
                        rowH: TABLE_ROW_H,
                        colW: colW
                    });

                    // Keep ID text styling plain while preserving click-through links.
                    for (var r = 0; r < pageItems.length; r++) {
                        pageSlide.addShape(pres.ShapeType.rect, {
                            x: PPT_X,
                            y: FEATURE_TABLE_TOP_Y + TABLE_ROW_H * (r + 1),
                            w: colW[0],
                            h: TABLE_ROW_H,
                            line: { color: 'FFFFFF', transparency: 100, pt: 0 },
                            fill: { color: 'FFFFFF', transparency: 100 },
                            hyperlink: { url: WORKITEM_LINK_BASE + pageItems[r].id }
                        });
                    }

                    start += size;
                }

                var teamName = extractTeamNameFromSlideTitle(sd.title);
                var teamPipeline = teamPipelines[teamName] || null;
                if (!addPipelineDataSlides(teamName, teamPipeline)) {
                    addPipelinePlaceholderSlide(teamName);
                }
            }
        }

        if (window.electronBridge && typeof window.electronBridge.savePptx === 'function') {
            var arrayBuffer = await pres.write('arraybuffer');
            await window.electronBridge.savePptx(arrayBuffer, deckFileName + '.pptx');
        } else {
            pres.writeFile({ fileName: deckFileName + '.pptx' });
        }
    }

    var btn = document.getElementById('pptxBtn');
    if (btn) btn.addEventListener('click', downloadPptx);
})();
