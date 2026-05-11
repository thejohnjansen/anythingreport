import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { uploadPptToGraph } from '../graphUploadService';
import { generatePpt } from '../pptService';
import { StoryRepository } from '../storyRepository';
import type { StoryUpdateRequest, StoryUpsertRequest } from '../types';

const repository = new StoryRepository();

function json(body: unknown, init: Partial<HttpResponseInit> = {}): HttpResponseInit {
  return {
    status: init.status ?? 200,
    headers: { 'content-type': 'application/json; charset=utf-8', ...(init.headers ?? {}) },
    body: JSON.stringify(body, null, 2)
  };
}

function getQueryValue(request: HttpRequest, name: string): string | undefined {
  return request.query.get(name) ?? undefined;
}

function readJsonBody<T>(request: HttpRequest): Promise<T> {
  return request.json() as Promise<T>;
}

function parseBearerToken(value: string | undefined): string | null {
  if (!value) {
    return null;
  }
  const match = value.match(/^Bearer\s+(.+)$/i);
  return match?.[1]?.trim() || null;
}

app.http('getStories', {
  methods: ['GET'],
  route: 'stories',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const cycle = getQueryValue(request, 'cycle');
    if (!cycle) {
      return json({ error: 'Query string parameter cycle is required.' }, { status: 400 });
    }

    context.log(`Fetching stories for cycle ${cycle}`);
    const stories = await repository.getStoriesByCycle(cycle);
    return json({ items: stories });
  }
});

app.http('postStory', {
  methods: ['POST'],
  route: 'stories',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const body = await readJsonBody<StoryUpsertRequest>(request);
    const story = await repository.upsertStory(body);
    return json(story, { status: 201 });
  }
});

app.http('putStory', {
  methods: ['PUT'],
  route: 'stories/{id}',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const id = request.params.id;
    if (!id) {
      return json({ error: 'Story id is required.' }, { status: 400 });
    }

    const body = await readJsonBody<StoryUpsertRequest & StoryUpdateRequest>(request);
    if (typeof body.expectedVersion !== 'number') {
      return json({ error: 'expectedVersion is required for concurrency-safe update.' }, { status: 400 });
    }

    try {
      const story = await repository.updateStoryWithConcurrencyCheck(id, body.expectedVersion, body);
      return json(story);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      if (message.startsWith('Version conflict')) {
        return json({ error: message }, { status: 409 });
      }
      return json({ error: message }, { status: 404 });
    }
  }
});

app.http('exportStoriesPpt', {
  methods: ['POST'],
  route: 'stories/export',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const cycle = getQueryValue(request, 'cycle');
    if (!cycle) {
      return json({ error: 'Query string parameter cycle is required.' }, { status: 400 });
    }

    const authHeader = request.headers.get('authorization') ?? undefined;
    const accessToken = parseBearerToken(authHeader) ?? process.env.GRAPH_ACCESS_TOKEN ?? null;
    if (!accessToken) {
      return json({ error: 'Graph access token missing. Provide Authorization: Bearer <token> or set GRAPH_ACCESS_TOKEN.' }, { status: 401 });
    }

    const folderPath = getQueryValue(request, 'folderPath') ?? process.env.GRAPH_UPLOAD_FOLDER ?? 'AnythingReport/Exports';
    const fileName = getQueryValue(request, 'fileName') ?? `stories-${cycle}-${Date.now()}.pptx`;
    const mode = getQueryValue(request, 'uploadMode') ?? process.env.GRAPH_UPLOAD_MODE ?? 'delegated';
    const driveId = getQueryValue(request, 'driveId') ?? process.env.GRAPH_DRIVE_ID;
    const userId = getQueryValue(request, 'userId') ?? process.env.GRAPH_USER_ID;

    try {
      context.log(`Exporting stories to PPT for cycle ${cycle}`);
      const stories = await repository.getStoriesByCycle(cycle);
      if (stories.length === 0) {
        return json({ error: `No stories found for cycle ${cycle}.` }, { status: 404 });
      }

      const pptBuffer = await generatePpt(stories);

      let fileUrl: string;
      if (mode === 'application') {
        if (!driveId) {
          return json({ error: 'driveId is required for uploadMode=application.' }, { status: 400 });
        }

        fileUrl = await uploadPptToGraph({
          accessToken,
          fileName,
          fileBuffer: pptBuffer,
          target: {
            mode: 'application',
            driveId,
            folderPath
          }
        });
      } else {
        fileUrl = await uploadPptToGraph({
          accessToken,
          fileName,
          fileBuffer: pptBuffer,
          target: {
            mode: 'delegated',
            folderPath,
            ...(userId ? { userId } : {})
          }
        });
      }

      return json({
        cycle,
        storiesExported: stories.length,
        fileName,
        fileUrl
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return json({ error: message }, { status: 500 });
    }
  }
});
