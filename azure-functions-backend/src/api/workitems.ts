import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { fetchWorkItems } from '../workItemService';
import { getIncomingAdoToken, requireMicrosoftUser } from '../requestAuth';

function json(body: unknown, status = 200): HttpResponseInit {
  return {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body, null, 2)
  };
}

app.http('getWorkItems', {
  methods: ['GET'],
  route: 'workitems',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const accessError = requireMicrosoftUser(request);
    if (accessError) {
      return accessError;
    }

    const iterationPath = request.query.get('iterationPath') ?? undefined;
    const areaPath = request.query.get('areaPath') ?? undefined;

    if (!iterationPath && !areaPath) {
      return json({ error: 'Provide iterationPath or areaPath query parameter.' }, 400);
    }

    try {
      const items = await fetchWorkItems({ iterationPath, areaPath }, getIncomingAdoToken(request));
      return json({ items });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return json({ error: message }, 500);
    }
  }
});
