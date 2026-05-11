"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const workItemService_1 = require("../workItemService");
function json(body, status = 200) {
    return {
        status,
        headers: { 'content-type': 'application/json; charset=utf-8' },
        body: JSON.stringify(body, null, 2)
    };
}
functions_1.app.http('getWorkItems', {
    methods: ['GET'],
    route: 'workitems',
    authLevel: 'anonymous',
    handler: async (request) => {
        const iterationPath = request.query.get('iterationPath') ?? undefined;
        const areaPath = request.query.get('areaPath') ?? undefined;
        if (!iterationPath && !areaPath) {
            return json({ error: 'Provide iterationPath or areaPath query parameter.' }, 400);
        }
        try {
            const items = await (0, workItemService_1.fetchWorkItems)({ iterationPath, areaPath });
            return json({ items });
        }
        catch (error) {
            const message = error instanceof Error ? error.message : 'Unknown error';
            return json({ error: message }, 500);
        }
    }
});
