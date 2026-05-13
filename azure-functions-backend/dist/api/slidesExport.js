"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const reportPptService_1 = require("../reportPptService");
const requestAuth_1 = require("../requestAuth");
function errorResponse(message, status = 500) {
    return {
        status,
        headers: { 'content-type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ error: message }, null, 2)
    };
}
functions_1.app.http('exportSlidesPpt', {
    methods: ['POST'],
    route: 'slides/export',
    authLevel: 'anonymous',
    handler: async (request) => {
        const accessError = (0, requestAuth_1.requireMicrosoftUser)(request);
        if (accessError) {
            return accessError;
        }
        const body = (await request.json());
        if (!Array.isArray(body.slides)) {
            return errorResponse('slides array is required.', 400);
        }
        try {
            const buffer = await (0, reportPptService_1.generateReportPpt)({
                hasMidpoint: !!body.hasMidpoint,
                slides: body.slides,
                topOfMindHtml: body.topOfMindHtml,
                pipeline: body.pipeline,
                teamPipelines: body.teamPipelines,
                baseSlideTitle: body.baseSlideTitle,
                baseTeamName: body.baseTeamName,
                deckFileName: body.deckFileName,
                linkBase: body.linkBase,
                theme: body.theme
            });
            return {
                status: 200,
                headers: {
                    'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    'content-disposition': `attachment; filename="${body.deckFileName || 'anything-report'}.pptx"`
                },
                body: new Uint8Array(buffer)
            };
        }
        catch (error) {
            return errorResponse(error instanceof Error ? error.message : 'Unknown error');
        }
    }
});
