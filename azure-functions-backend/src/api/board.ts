import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { BoardRepository } from '../boardRepository';
import type { BoardUpsertRequest } from '../types';

function json(body: unknown, init: Partial<HttpResponseInit> = {}): HttpResponseInit {
  return {
    status: init.status ?? 200,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body, null, 2)
  };
}

app.http('getBoard', {
  methods: ['GET'],
  route: 'board',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const boardType = request.query.get('type');
    const teamSlug = request.query.get('teamSlug') ?? 'default';

    if (!boardType || (boardType !== 'topofmind' && boardType !== 'pipeline')) {
      return json({ error: 'Query param "type" must be "topofmind" or "pipeline".' }, { status: 400 });
    }

    try {
      const repo = new BoardRepository();
      const document = await repo.getBoardDocument(boardType, teamSlug);
      return json({ document });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return json({ error: message }, { status: 500 });
    }
  }
});

app.http('putBoard', {
  methods: ['PUT'],
  route: 'board',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    try {
      const body = (await request.json()) as BoardUpsertRequest;

      if (!body.boardType || (body.boardType !== 'topofmind' && body.boardType !== 'pipeline')) {
        return json({ error: 'boardType must be "topofmind" or "pipeline".' }, { status: 400 });
      }
      if (typeof body.content !== 'string') {
        return json({ error: 'content is required.' }, { status: 400 });
      }
      if (!body.lastEditedBy) {
        return json({ error: 'lastEditedBy is required.' }, { status: 400 });
      }

      const repo = new BoardRepository();
      const document = await repo.upsertBoardDocument(body);
      return json({ document });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      if (message.startsWith('Version conflict')) {
        const currentDocument = (error as any).currentDocument ?? null;
        return json({ error: message, currentDocument }, { status: 409 });
      }
      return json({ error: message }, { status: 500 });
    }
  }
});
