import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { JSDOM } from 'jsdom';
import PptxGenJS from 'pptxgenjs';
import type { PreviewSlide } from './queryPreviewService';

const V_ALIGN_MID = 'mid' as any;

export interface PipelineRow {
  topic: string;
  stages: string[];
  hereCol?: number;
  warnCol?: number;
}

export interface PipelineData {
  title?: string;
  queryUrl?: string;
  rows: PipelineRow[];
}

export interface ReportPptTheme {
  titleFont: string;
  titleFontSize: number;
  titleX: number;
  titleY: number;
  titleW: number;
  titleH: number;
  bodyFont: string;
  emojiFont: string;
  bgColor: string;
  titleColor: string;
  bodyColor: string;
  mutedTextColor: string;
  headerFillColor: string;
  rowAltColor: string;
  borderColor: string;
  linkColor: string;
  pipelineLineColor: string;
  pipelineCompleteColor: string;
  pipelineCommittedFillColor: string;
  pipelineFutureColor: string;
  topicPillFillColor: string;
  topicPillBorderColor: string;
  topicPillRadius: number;
  topicPillTextColor: string;
  backgroundImageUrl: string;
  titleImageUrl: string;
}

export interface GenerateReportPptInput {
  hasMidpoint: boolean;
  slides: PreviewSlide[];
  topOfMindHtml?: string;
  pipeline?: PipelineData;
  teamPipelines?: Record<string, PipelineData>;
  baseSlideTitle?: string;
  baseTeamName?: string;
  deckFileName?: string;
  linkBase?: string;
  cycleNumber?: string;
  theme?: Partial<ReportPptTheme>;
}

const DEFAULT_THEME: ReportPptTheme = {
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
  backgroundImageUrl: '../pptx_inspect/slideTemplate.png',
  titleImageUrl: '../pptx_inspect/InitialSlideBackground.png'
};

function getTheme(theme?: Partial<ReportPptTheme>): ReportPptTheme {
  return { ...DEFAULT_THEME, ...(theme ?? {}) };
}

function sanitizeForPpt(text: string): string {
  if (!text) {
    return '';
  }

  let output = '';
  for (const character of text) {
    const codePoint = character.codePointAt(0) ?? 0;
    const isValidXml =
      codePoint === 0x9 ||
      codePoint === 0xa ||
      codePoint === 0xd ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff);
    if (isValidXml) {
      output += character;
    }
  }
  return output;
}

function htmlToPptRuns(html: string): Array<{ text: string; options?: Record<string, unknown> }> {
  const dom = new JSDOM('<!doctype html><html><body></body></html>');
  const root = dom.window.document.createElement('div');
  root.innerHTML = html || '';
  const runs: Array<{ text: string; options?: Record<string, unknown> }> = [];

  function endsWithBreak(): boolean {
    if (!runs.length) {
      return false;
    }
    const last = runs[runs.length - 1];
    return last.text === '' && !!last.options?.breakLine;
  }

  function pushRun(text: string, format: { bold: boolean; italic: boolean; underline: boolean }, paraOpts: Record<string, unknown> | null): void {
    const clean = sanitizeForPpt(text);
    if (!clean) {
      return;
    }

    const options: Record<string, unknown> = {
      bold: format.bold,
      italic: format.italic,
      underline: format.underline
    };
    if (paraOpts) {
      if (paraOpts.bullet) {
        options.bullet = paraOpts.bullet;
      }
      if (typeof paraOpts.indentLevel === 'number') {
        options.indentLevel = paraOpts.indentLevel;
      }
    }

    runs.push({ text: clean, options });
  }

  function pushBreak(): void {
    if (!runs.length) {
      return;
    }
    runs.push({ text: '', options: { breakLine: true } });
  }

  function listDepth(node: Element | null): number {
    let depth = 0;
    let parent = node?.parentElement ?? null;
    while (parent) {
      const tagName = parent.tagName ? parent.tagName.toLowerCase() : '';
      if (tagName === 'ul' || tagName === 'ol') {
        depth += 1;
      }
      parent = parent.parentElement;
    }
    return depth;
  }

  function consumePendingParagraphOptions(state: { pendingParaOptions: Record<string, unknown> | null }, textPart: string): Record<string, unknown> | null {
    if (!state.pendingParaOptions || !/\S/.test(textPart || '')) {
      return null;
    }
    const options = state.pendingParaOptions;
    state.pendingParaOptions = null;
    return options;
  }

  function walk(
    node: ChildNode,
    format: { bold: boolean; italic: boolean; underline: boolean },
    state: { pendingParaOptions: Record<string, unknown> | null }
  ): void {
    if (node.nodeType === dom.window.Node.TEXT_NODE) {
      const text = (node.nodeValue || '').replace(/\r\n?/g, '\n');
      const parts = text.split('\n');
      for (let index = 0; index < parts.length; index += 1) {
        const paraOpts = consumePendingParagraphOptions(state, parts[index]);
        pushRun(parts[index], format, paraOpts);
        if (index < parts.length - 1) {
          pushBreak();
        }
      }
      return;
    }

    if (node.nodeType !== dom.window.Node.ELEMENT_NODE) {
      return;
    }

    const element = node as Element;
    const tag = element.tagName.toLowerCase();
    const nextFormat = { ...format };

    if (['div', 'p', 'ul', 'ol'].includes(tag) && runs.length && !endsWithBreak()) {
      pushBreak();
    }
    if (tag === 'strong' || tag === 'b') {
      nextFormat.bold = true;
    }
    if (tag === 'em' || tag === 'i') {
      nextFormat.italic = true;
    }
    if (tag === 'u') {
      nextFormat.underline = true;
    }

    if (tag === 'br') {
      pushBreak();
      return;
    }

    if (tag === 'li') {
      const parentTag = element.parentElement?.tagName?.toLowerCase() || '';
      const level = Math.max(0, listDepth(element) - 1);
      const liState = {
        pendingParaOptions:
          parentTag === 'ol'
            ? ({ bullet: { type: 'number' }, indentLevel: level } as Record<string, unknown>)
            : ({ bullet: true, indentLevel: level } as Record<string, unknown>)
      };

      if (runs.length && !endsWithBreak()) {
        pushBreak();
      }

      for (const child of Array.from(element.childNodes)) {
        walk(child, nextFormat, liState);
      }
      if (liState.pendingParaOptions) {
        pushRun(' ', nextFormat, liState.pendingParaOptions);
      }
      pushBreak();
      return;
    }

    for (const child of Array.from(element.childNodes)) {
      walk(child, nextFormat, state);
    }
    if (tag === 'div' || tag === 'p') {
      pushBreak();
    }
  }

  for (const child of Array.from(root.childNodes) as ChildNode[]) {
    walk(child, { bold: false, italic: false, underline: false }, { pendingParaOptions: null });
  }

  while (runs.length) {
    const last = runs[runs.length - 1];
    if (last.text === '' && last.options?.breakLine) {
      runs.pop();
      continue;
    }
    break;
  }

  return runs;
}

function riskColor(risk: string): string {
  const value = (risk || '').toLowerCase();
  if (value.includes('on track')) {
    return '16A34A';
  }
  if (value.includes('at risk')) {
    return 'B8860B';
  }
  if (value.includes('off track')) {
    return 'DC2626';
  }
  return '888888';
}

function riskLabel(risk: string): string {
  const value = (risk || '').toLowerCase();
  if (value.includes('on track')) {
    return '✓ On Track';
  }
  if (value.includes('at risk')) {
    return '⚠ At Risk';
  }
  if (value.includes('off track')) {
    return '✗ Off Track';
  }
  return '—';
}

function riskFillColor(risk: string): string {
  const value = (risk || '').toLowerCase();
  if (value.includes('on track')) {
    return 'E9F7EF';
  }
  if (value.includes('at risk')) {
    return 'FFF6DB';
  }
  if (value.includes('off track')) {
    return 'FDECEC';
  }
  return 'F5F5F5';
}

function tableNoBorder() {
  return { type: 'solid' as const, pt: 0, color: 'FFFFFF' };
}

function tableHeaderBorder(theme: ReportPptTheme) {
  return [
    { type: 'solid', pt: 1.2, color: theme.titleColor },
    tableNoBorder(),
    { type: 'solid', pt: 1.2, color: theme.titleColor },
    tableNoBorder()
  ];
}

function tableRowBottomBorder(theme: ReportPptTheme) {
  return [
    tableNoBorder(),
    tableNoBorder(),
    { type: 'solid', pt: 0.8, color: theme.borderColor },
    tableNoBorder()
  ];
}

function sanitizeFileName(name: string): string {
  const value = String(name || '').trim();
  if (!value) {
    return 'anything-report';
  }
  return value.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim() || 'anything-report';
}

function normalizePipelineStageValue(value: string): string {
  const normalized = (value || '').trim();
  return ['complete', 'completed'].includes(normalized.toLowerCase()) ? 'complete' : normalized;
}

function parsePipelineStageValue(value: string): { complete: boolean; committed: boolean; label: string } {
  const raw = (value || '').trim();
  if (!raw) {
    return { complete: false, committed: false, label: '' };
  }

  const lower = raw.toLowerCase();
  if (lower === 'complete' || lower === 'completed') {
    return { complete: true, committed: false, label: '' };
  }
  if (lower.startsWith('complete|') || lower.startsWith('completed|')) {
    const pipeIndex = raw.indexOf('|');
    return { complete: true, committed: false, label: pipeIndex >= 0 ? raw.slice(pipeIndex + 1).trim() : '' };
  }
  if (lower === 'committed' || lower === 'commit') {
    return { complete: false, committed: true, label: '' };
  }
  if (lower.startsWith('committed|') || lower.startsWith('commit|')) {
    const pipeIndex = raw.indexOf('|');
    return { complete: false, committed: true, label: pipeIndex >= 0 ? raw.slice(pipeIndex + 1).trim() : '' };
  }
  return { complete: false, committed: false, label: normalizePipelineStageValue(raw) };
}

function extractTeamNameFromSlideTitle(title: string): string {
  const value = String(title || '').trim();
  const separator = value.indexOf(' - ');
  return separator >= 0 ? value.slice(separator + 3).trim() || value : value;
}

async function toDataUrl(urlOrPath: string): Promise<string> {
  if (!urlOrPath) {
    return '';
  }

  if (/^https?:\/\//i.test(urlOrPath)) {
    const response = await fetch(urlOrPath);
    if (!response.ok) {
      return '';
    }
    const arrayBuffer = await response.arrayBuffer();
    const extension = path.extname(new URL(urlOrPath).pathname).toLowerCase();
    const mimeType = extension === '.jpg' || extension === '.jpeg' ? 'image/jpeg' : 'image/png';
    return `data:${mimeType};base64,${Buffer.from(arrayBuffer).toString('base64')}`;
  }

  const candidates = [
    path.resolve(process.cwd(), urlOrPath),
    path.resolve(process.cwd(), '..', urlOrPath)
  ];

  for (const candidate of candidates) {
    try {
      const contents = await readFile(candidate);
      const extension = path.extname(candidate).toLowerCase();
      const mimeType = extension === '.jpg' || extension === '.jpeg' ? 'image/jpeg' : 'image/png';
      return `data:${mimeType};base64,${contents.toString('base64')}`;
    } catch {
      continue;
    }
  }

  return '';
}

function addSlideHeader(slide: PptxGenJS.Slide, title: string, theme: ReportPptTheme, bgImageData: string): void {
  slide.background = { color: theme.bgColor };
  if (bgImageData) {
    slide.addImage({ data: bgImageData, x: 0, y: 0, w: 13.333, h: 7.5 });
  }
  slide.addText(String(title || '').replace(/^<\/>\s*/, ''), {
    x: theme.titleX,
    y: theme.titleY,
    w: theme.titleW,
    h: theme.titleH,
    fontSize: theme.titleFontSize,
    bold: true,
    color: theme.titleColor,
    fontFace: theme.titleFont,
    valign: V_ALIGN_MID
  });
}

function addPipelineLegend(
  slide: PptxGenJS.Slide,
  pres: PptxGenJS,
  theme: ReportPptTheme,
  chartX: number,
  chartW: number
): number {
  const legendX = chartX + chartW - 4.9;
  const legendY = 0.25;
  const legendColGap = 2.45;
  const legendRowH = 0.22;
  const legendSymbolW = 0.24;
  const legendTextW = 2.25;

  function addLegendItem(col: number, row: number, symbol: string, text: string, options: Record<string, unknown>): void {
    const x = legendX + col * legendColGap;
    const y = legendY + row * legendRowH;
    if (options.symbolType === 'dotFilled' || options.symbolType === 'dotOpen') {
      const diameter = Number(options.symbolDiameter || 0.18);
      slide.addShape(pres.ShapeType.ellipse, {
        x: x + (legendSymbolW - diameter) / 2,
        y: y + 0.03,
        w: diameter,
        h: diameter,
        line: { color: String(options.symbolColor), pt: options.symbolType === 'dotOpen' ? 2 : 1.2 },
        fill: { color: options.symbolType === 'dotFilled' ? String(options.symbolColor) : 'FFFFFF' }
      });
    } else if (options.symbolType === 'dotFilledBorder') {
      const diameter = Number(options.symbolDiameter || 0.18);
      slide.addShape(pres.ShapeType.ellipse, {
        x: x + (legendSymbolW - diameter) / 2,
        y: y + 0.03,
        w: diameter,
        h: diameter,
        line: { color: String(options.symbolColor), pt: 2 },
        fill: { color: String(options.fillColor || 'FFFFFF') }
      });
    } else {
      slide.addText(symbol, {
        x,
        y,
        w: legendSymbolW,
        h: 0.2,
        align: 'center',
        valign: V_ALIGN_MID,
        fontSize: Number(options.symbolSize),
        bold: Boolean(options.symbolBold),
        color: String(options.symbolColor),
        fontFace: String(options.symbolFont || 'Segoe UI')
      });
    }

    slide.addText(text, {
      x: x + legendSymbolW + 0.04,
      y,
      w: legendTextW,
      h: 0.2,
      fontSize: 9,
      color: theme.mutedTextColor,
      fontFace: theme.bodyFont
    });
  }

  addLegendItem(0, 0, '▼', 'Roughly where we are', { symbolSize: 11, symbolBold: true, symbolColor: theme.pipelineCompleteColor });
  addLegendItem(0, 1, '⚠️', 'Risk Identified', { symbolSize: 12, symbolColor: 'D97706', symbolFont: theme.emojiFont });
  addLegendItem(1, 0, '', 'Completed', { symbolType: 'dotFilled', symbolColor: theme.pipelineCompleteColor, symbolDiameter: 0.18 });
  addLegendItem(1, 1, '', 'Committed for the current cycle', { symbolType: 'dotFilledBorder', symbolColor: theme.pipelineCompleteColor, fillColor: theme.pipelineCommittedFillColor, symbolDiameter: 0.18 });
  addLegendItem(1, 2, '', 'Planned for a future cycle', { symbolType: 'dotOpen', symbolColor: theme.pipelineFutureColor, symbolDiameter: 0.18 });
  return legendY + legendRowH * 3 + 0.02;
}

export async function generateReportPpt(input: GenerateReportPptInput): Promise<Buffer> {
  const theme = getTheme(input.theme);
  const bgImageData = await toDataUrl(theme.backgroundImageUrl);
  const titleImageData = await toDataUrl(theme.titleImageUrl);
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = sanitizeFileName(input.deckFileName || 'anything-report');
  pres.author = 'anything-report';

  const slides = input.slides || [];
  const pipeline = input.pipeline || { rows: [] };
  const teamPipelines = input.teamPipelines || {};
  const baseTeamName = (input.baseTeamName || '').trim();
  const deckFileName = sanitizeFileName(input.deckFileName || 'anything-report');
  let cycleNumber = (input.cycleNumber || '').trim();
  if (!cycleNumber) {
    const firstWord = deckFileName.split(' ')[0];
    if (/^\d+[A-Za-z]\d+$/.test(firstWord)) cycleNumber = firstWord;
  }
  const monthYear = new Date().toLocaleString('en-US', { month: 'long', year: 'numeric' });
  const topOfMindTitle = baseTeamName ? `Top of Mind - ${baseTeamName}` : 'Top of Mind';
  const pipelineBaseTitle = baseTeamName ? `Pipeline - ${baseTeamName}` : 'Pipeline';
  const linkBase = input.linkBase || 'https://microsoft.visualstudio.com/Edge/_workitems/edit/';

  const PPT_X = 0.2;
  const PPT_W = 13;
  const TABLE_TOP_Y = 1.5;
  const FEATURE_TABLE_TOP_Y = TABLE_TOP_Y - 0.5;
  const TABLE_ROW_H = 0.5;
  const plChartX = 0.2;
  const plChartW = 12.93;
  const plChartBottomY = 7.05;
  const plTopicW = 3.4;
  const plStageW = plChartW - plTopicW;
  const plStageColW = plStageW / 6;
  const plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

  // Title slide (first slide)
  const titleSlide = pres.addSlide();
  if (titleImageData) {
    titleSlide.addImage({ data: titleImageData, x: 0, y: 0, w: 13.333, h: 7.5 });
  } else {
    titleSlide.background = { color: theme.bgColor };
  }
  if (baseTeamName) {
    titleSlide.addText(baseTeamName, {
      x: 0.4, y: 5.0, w: 9, h: 0.82,
      fontSize: 42, bold: true, color: 'FFFFFF',
      fontFace: 'Segoe UI Semibold', valign: V_ALIGN_MID
    });
  }
  titleSlide.addText('End of Cycle Review', {
    x: 0.4, y: 5.8, w: 9, h: 0.65,
    fontSize: 32, bold: true, color: 'FFFFFF',
    fontFace: 'Segoe UI Semibold', valign: V_ALIGN_MID
  });
  const subtitleParts = ['Web Platform'];
  if (cycleNumber) subtitleParts.push(cycleNumber);
  if (monthYear) subtitleParts.push(monthYear);
  titleSlide.addText(subtitleParts.join(' - '), {
    x: 0.4, y: 6.42, w: 10, h: 0.45,
    fontSize: 16, color: 'FFFFFF',
    fontFace: 'Segoe UI', valign: V_ALIGN_MID
  });

  const PPT_X = 0.2;
  const PPT_W = 13;
  const TABLE_TOP_Y = 1.5;
  const FEATURE_TABLE_TOP_Y = TABLE_TOP_Y - 0.5;
  const TABLE_ROW_H = 0.5;
  const plChartX = 0.2;
  const plChartW = 12.93;
  const plChartBottomY = 7.05;
  const plTopicW = 3.4;
  const plStageW = plChartW - plTopicW;
  const plStageColW = plStageW / 6;
  const plHeaders = ['INVESTIGATE', 'EXPLAINER /\nDESIGN DOC', 'IMPLEMENTATION', 'DEV TRIAL', 'ORIGIN TRIAL /\nCFR', 'SHIP'];

  const topOfMindSlide = pres.addSlide();
  addSlideHeader(topOfMindSlide, topOfMindTitle, theme, bgImageData);
  const topOfMindRuns = htmlToPptRuns(input.topOfMindHtml || '');
  if (topOfMindRuns.length) {
    topOfMindSlide.addText(topOfMindRuns as any, {
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

  const allPipelineRows = (pipeline.rows || []).slice(0, 12);
  const pipelinePageGroups = allPipelineRows.length > 7 ? [allPipelineRows.slice(0, 6), allPipelineRows.slice(6)] : [allPipelineRows];
  const hasFlatTeamSlides = slides.some((slide) => String(slide.id || '').startsWith('flat-'));

  function renderPipelineSlides(title: string, pageGroups: PipelineRow[][], presentation: PptxGenJS): void {
    for (let pageIndex = 0; pageIndex < pageGroups.length; pageIndex += 1) {
      const pageRows = pageGroups[pageIndex];
      const slide = presentation.addSlide();
      addSlideHeader(slide, pageGroups.length > 1 ? `${title} (${pageIndex + 1}/${pageGroups.length})` : title, theme, bgImageData);

      if (pageRows.length === 0) {
        slide.addText('No pipeline data — use the Edit button in the web view to add rows.', {
          x: PPT_X,
          y: TABLE_TOP_Y,
          w: PPT_W,
          h: 0.6,
          fontSize: 13,
          color: '888888',
          fontFace: theme.bodyFont,
          italic: true
        });
        continue;
      }

      const legendBottomY = addPipelineLegend(slide, presentation, theme, plChartX, plChartW);
      const headerY = legendBottomY + 0.14;
      for (let headerIndex = 0; headerIndex < 6; headerIndex += 1) {
        slide.addText(plHeaders[headerIndex], {
          x: plChartX + plTopicW + headerIndex * plStageColW,
          y: headerY,
          w: plStageColW,
          h: 0.36,
          align: 'center',
          valign: V_ALIGN_MID,
          bold: true,
          fontSize: 9,
          color: '666666',
          fontFace: theme.bodyFont
        });
      }

      const rowStartY = headerY + 0.58;
      const rowGap = 0.15;
      const fixedRows = 6;
      const available = plChartBottomY - rowStartY - (fixedRows - 1) * rowGap;
      const rowH = Math.min(0.95, available / fixedRows);
      const nodeD = Math.min(0.65, rowH * 1.05);
      const referenceRows = 5;
      const referenceAvailable = plChartBottomY - rowStartY - (referenceRows - 1) * rowGap;
      const referenceRowH = Math.min(0.95, referenceAvailable / referenceRows);
      const topicHeightTrim = 20 / 96;
      const topicBoxH = Math.max(0.3, referenceRowH * 1.05 - topicHeightTrim);
      const topicTextH = Math.max(0.42, referenceRowH * 0.88 - topicHeightTrim);

      for (let rowIndex = 0; rowIndex < pageRows.length; rowIndex += 1) {
        const row = pageRows[rowIndex];
        const y = rowStartY + rowIndex * (rowH + rowGap);
        const centerY = y + rowH * 0.5;

        slide.addShape(presentation.ShapeType.roundRect, {
          x: plChartX,
          y: centerY - topicBoxH * 0.5,
          w: plTopicW - 0.18,
          h: topicBoxH,
          rectRadius: theme.topicPillRadius,
          line: { color: theme.topicPillBorderColor, pt: 1.2 },
          fill: { color: theme.topicPillFillColor }
        });
        slide.addText(row.topic || '', {
          x: plChartX + 0.06,
          y: centerY - topicTextH * 0.5,
          w: plTopicW - 0.3,
          h: topicTextH,
          align: 'center',
          valign: V_ALIGN_MID,
          bold: true,
          fontSize: 18,
          color: theme.topicPillTextColor,
          fontFace: theme.bodyFont
        });
        slide.addShape(presentation.ShapeType.line, {
          x: plChartX + plTopicW - 0.18,
          y: centerY,
          w: plChartX + plTopicW + plStageColW * 5.5 + nodeD * 0.5 - (plChartX + plTopicW - 0.18),
          h: 0,
          line: { color: theme.pipelineLineColor, pt: 5 }
        });

        for (let stageIndex = 0; stageIndex < 6; stageIndex += 1) {
          const centerX = plChartX + plTopicW + plStageColW * (stageIndex + 0.5);
          const parsedStage = parsePipelineStageValue(row.stages?.[stageIndex] || '');
          const label = parsedStage.label;
          const isHere = row.hereCol === stageIndex;
          const isWarn = row.warnCol === stageIndex;
          const isDone = parsedStage.complete;
          const isCommitted = parsedStage.committed;

          if (isHere) {
            slide.addShape(presentation.ShapeType.downArrow, {
              x: centerX - nodeD * 0.5 - 0.25,
              y: centerY - 0.27,
              w: 0.12,
              h: 0.24,
              line: { color: theme.pipelineCompleteColor, pt: 1.3 },
              fill: { color: theme.pipelineCompleteColor }
            });
          }
          if (isWarn) {
            slide.addText('⚠️', {
              x: centerX - nodeD * 0.5 - 0.61,
              y: centerY - 0.24,
              w: 0.32,
              h: 0.32,
              align: 'center',
              valign: V_ALIGN_MID,
              fontSize: 24,
              fontFace: theme.emojiFont
            });
          }
          if (isDone) {
            slide.addShape(presentation.ShapeType.ellipse, {
              x: centerX - nodeD * 0.5,
              y: centerY - nodeD * 0.5,
              w: nodeD,
              h: nodeD,
              line: { color: theme.pipelineCompleteColor, pt: 1.5 },
              fill: { color: theme.pipelineCompleteColor }
            });
            if (label) {
              slide.addText(label, {
                x: centerX - nodeD * 0.5 + 0.02,
                y: centerY - nodeD * 0.5 + 0.02,
                w: nodeD - 0.04,
                h: nodeD - 0.04,
                align: 'center',
                valign: V_ALIGN_MID,
                fit: 'shrink',
                fontSize: 12,
                bold: true,
                color: '000000',
                fontFace: theme.bodyFont
              });
            }
            continue;
          }

          if (label || isCommitted) {
            slide.addShape(presentation.ShapeType.ellipse, {
              x: centerX - nodeD * 0.5,
              y: centerY - nodeD * 0.5,
              w: nodeD,
              h: nodeD,
              line: { color: isCommitted ? theme.pipelineCompleteColor : theme.pipelineFutureColor, pt: 2.5 },
              fill: { color: isCommitted ? theme.pipelineCommittedFillColor : 'FFFFFF' }
            });
            if (label) {
              slide.addText(label, {
                x: centerX - nodeD * 0.5 + 0.02,
                y: centerY - nodeD * 0.5 + 0.02,
                w: nodeD - 0.04,
                h: nodeD - 0.04,
                align: 'center',
                valign: V_ALIGN_MID,
                fit: 'shrink',
                fontSize: 12,
                bold: true,
                color: isCommitted ? theme.topicPillTextColor : '666666',
                fontFace: theme.bodyFont
              });
            }
          }
        }
      }
    }
  }

  if (!hasFlatTeamSlides) {
    renderPipelineSlides(pipelineBaseTitle, pipelinePageGroups, pres);
  }

  for (const slideData of slides) {
    const featureSlide = pres.addSlide();
    addSlideHeader(featureSlide, slideData.title, theme, bgImageData);

    if (slideData.items.length === 0) {
      featureSlide.addText('No child items', {
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

    const columns = input.hasMidpoint
      ? ['ID', 'Title', 'Midpoint Risk', 'Midpoint Details', 'Final Risk', 'Final Details']
      : ['ID', 'Title', 'Risk', 'Details'];
    const columnWidths = input.hasMidpoint ? [1, 3.2, 1.25, 2.95, 1.25, 3.3] : [1, 3.7, 1.35, 5.95];
    const headerRow = columns.map((column) => ({
      text: column,
      options: {
        bold: true,
        color: '000000',
        fill: { color: 'FFFFFF' },
        fontSize: 12,
        align: 'center',
        valign: 'middle',
        fontFace: theme.bodyFont,
        border: tableHeaderBorder(theme)
      }
    }));

    const bodyRows = slideData.items.map((item) => {
      const idCell = {
        text: String(item.id),
        options: {
          fontSize: 12,
          color: '000000',
          fill: { color: 'FFFFFF' },
          align: 'center',
          valign: 'middle',
          fontFace: theme.bodyFont,
          border: tableRowBottomBorder(theme),
          hyperlink: { url: `${linkBase}${item.id}` }
        }
      };

      if (input.hasMidpoint) {
        return [
          idCell,
          { text: item.title, options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
          { text: riskLabel(item.midpointRisk || ''), options: { fontSize: 12, bold: true, color: riskColor(item.midpointRisk || ''), fill: { color: riskFillColor(item.midpointRisk || '') }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
          { text: item.midpointComment || '—', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
          { text: riskLabel(item.risk), options: { fontSize: 12, bold: true, color: riskColor(item.risk), fill: { color: riskFillColor(item.risk) }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
          { text: item.riskComment || '—', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } }
        ];
      }

      return [
        idCell,
        { text: item.title, options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
        { text: riskLabel(item.risk), options: { fontSize: 12, bold: true, color: riskColor(item.risk), fill: { color: riskFillColor(item.risk) }, align: 'center', valign: 'middle', fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } },
        { text: item.riskComment || '—', options: { fontSize: 12, color: '000000', fill: { color: 'FFFFFF' }, fontFace: theme.bodyFont, border: tableRowBottomBorder(theme) } }
      ];
    });

    const pageSizes = slideData.items.length <= 8 ? [slideData.items.length] : slideData.items.length <= 13 ? [5, slideData.items.length - 5] : [5, 5, slideData.items.length - 10];
    let start = 0;
    for (let pageIndex = 0; pageIndex < pageSizes.length; pageIndex += 1) {
      const size = pageSizes[pageIndex];
      const pageItems = slideData.items.slice(start, start + size);
      const pageRows = bodyRows.slice(start, start + size);
      const pageSlide = pageIndex === 0 ? featureSlide : pres.addSlide();

      if (pageIndex > 0) {
        addSlideHeader(pageSlide, `${slideData.title} (cont.)`, theme, bgImageData);
      }

      pageSlide.addTable([headerRow, ...pageRows] as any, {
        x: PPT_X,
        y: FEATURE_TABLE_TOP_Y,
        w: PPT_W,
        border: tableNoBorder(),
        rowH: TABLE_ROW_H,
        colW: columnWidths
      });

      for (let itemIndex = 0; itemIndex < pageItems.length; itemIndex += 1) {
        pageSlide.addShape(pres.ShapeType.rect, {
          x: PPT_X,
          y: FEATURE_TABLE_TOP_Y + TABLE_ROW_H * (itemIndex + 1),
          w: columnWidths[0],
          h: TABLE_ROW_H,
          line: { color: 'FFFFFF', transparency: 100, pt: 0 },
          fill: { color: 'FFFFFF', transparency: 100 },
          hyperlink: { url: `${linkBase}${pageItems[itemIndex].id}` }
        });
      }
      start += size;
    }

    const teamName = extractTeamNameFromSlideTitle(slideData.title);
    const teamPipeline = teamPipelines[teamName];
    const teamRows = (teamPipeline?.rows || []).slice(0, 12);
    const teamPages = teamRows.length > 7 ? [teamRows.slice(0, 6), teamRows.slice(6)] : [teamRows];
    renderPipelineSlides(`Pipeline - ${teamName}`, teamPages, pres);
  }

  return (await pres.write({ outputType: 'nodebuffer' })) as Buffer;
}
