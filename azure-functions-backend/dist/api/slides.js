"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const queryPreviewService_1 = require("../queryPreviewService");
function json(body, status = 200) {
    return {
        status,
        headers: { 'content-type': 'application/json; charset=utf-8' },
        body: JSON.stringify(body, null, 2)
    };
}
functions_1.app.http('getSlidesPreview', {
    methods: ['POST'],
    route: 'slides',
    authLevel: 'anonymous',
    handler: async (request) => {
        const body = (await request.json());
        if (!body.queryUrl) {
            return json({ error: 'queryUrl is required.' }, 400);
        }
        try {
            const preview = await (0, queryPreviewService_1.getSlidesPreview)(body);
            return json(preview);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : 'Unknown error';
            return json({ error: message }, 500);
        }
    }
});
