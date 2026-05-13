import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { getSlidesPreview } from '../queryPreviewService';
import { getIncomingAdoToken, requireMicrosoftUser } from '../requestAuth';

function json(body: unknown, status = 200): HttpResponseInit {
  return {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body, null, 2)
  };
}

app.http('getSlidesPreview', {
  methods: ['POST'],
  route: 'slides',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const accessError = requireMicrosoftUser(request);
    if (accessError) {
      return accessError;
    }

    const body = (await request.json()) as {
      queryUrl?: string;
      midpointDate?: string;
      flatView?: boolean;
    };

    if (!body.queryUrl) {
      return json({ error: 'queryUrl is required.' }, 400);
    }

    try {
      const preview = await getSlidesPreview(
        body as { queryUrl: string; midpointDate?: string; flatView?: boolean },
        getIncomingAdoToken(request)
      );
      return json(preview);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return json({ error: message }, 500);
    }
  }
});
