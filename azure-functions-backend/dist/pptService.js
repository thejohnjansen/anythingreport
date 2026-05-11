"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generatePpt = generatePpt;
const pptxgenjs_1 = __importDefault(require("pptxgenjs"));
async function generatePpt(stories) {
    const pptx = new pptxgenjs_1.default();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'anything-report';
    pptx.subject = 'Work item storytelling';
    pptx.title = 'Stories Export';
    for (const story of stories) {
        const slide = pptx.addSlide();
        const header = `#${story.workItemId} - ${story.title}`;
        const subHeader = `Area: ${story.area} | Cycle: ${story.cycle}`;
        slide.addText(header, {
            x: 0.5,
            y: 0.3,
            w: 12.2,
            h: 0.6,
            fontSize: 22,
            bold: true,
            color: '1F2937'
        });
        slide.addText(subHeader, {
            x: 0.5,
            y: 0.95,
            w: 12.2,
            h: 0.3,
            fontSize: 12,
            color: '4B5563'
        });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.5,
            y: 1.4,
            w: 12.2,
            h: 3.2,
            fill: { color: 'F3F4F6' },
            line: { color: 'E5E7EB', pt: 1 }
        });
        slide.addText(story.narrative || '(No narrative provided)', {
            x: 0.75,
            y: 1.65,
            w: 11.7,
            h: 2.7,
            fontSize: 16,
            color: '111827',
            valign: 'top'
        });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.5,
            y: 4.9,
            w: 12.2,
            h: 2.1,
            fill: { color: 'FFF7ED' },
            line: { color: 'FDBA74', pt: 1 }
        });
        slide.addText('PM Comments', {
            x: 0.75,
            y: 5.1,
            w: 11.7,
            h: 0.35,
            fontSize: 13,
            bold: true,
            color: '9A3412'
        });
        slide.addText(story.pmComments || '(No comments provided)', {
            x: 0.75,
            y: 5.45,
            w: 11.7,
            h: 1.3,
            fontSize: 14,
            color: '7C2D12',
            valign: 'top'
        });
    }
    return (await pptx.write({ outputType: 'nodebuffer' }));
}
