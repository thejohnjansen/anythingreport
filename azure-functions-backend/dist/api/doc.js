"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const queryAuxService_1 = require("../queryAuxService");
const requestAuth_1 = require("../requestAuth");
function json(body, status = 200) {
    return {
        status,
        headers: { 'content-type': 'application/json; charset=utf-8' },
        body: JSON.stringify(body, null, 2)
    };
}
functions_1.app.http('getDocSections', {
    methods: ['POST'],
    route: 'doc',
    authLevel: 'anonymous',
    handler: async (request) => {
        const accessError = (0, requestAuth_1.requireMicrosoftUser)(request);
        if (accessError) {
            return accessError;
        }
        const body = (await request.json());
        if (!body.queryUrl) {
            return json({ error: 'queryUrl is required.' }, 400);
        }
        try {
            const result = await (0, queryAuxService_1.getDocSections)(body.queryUrl, (0, requestAuth_1.getIncomingAdoToken)(request));
            return json(result);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : 'Unknown error';
            return json({ error: message }, 500);
        }
    }
});
