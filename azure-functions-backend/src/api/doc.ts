import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { getDocSections } from '../queryAuxService';

function json(body: unknown, status = 200): HttpResponseInit {
  return {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body, null, 2)
  };
}

app.http('getDocSections', {
  methods: ['POST'],
  route: 'doc',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const body = (await request.json()) as { queryUrl?: string };
    if (!body.queryUrl) {
      return json({ error: 'queryUrl is required.' }, 400);
    }

    try {
      const result = await getDocSections(body.queryUrl);
      return json(result);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return json({ error: message }, 500);
    }
  }
});
