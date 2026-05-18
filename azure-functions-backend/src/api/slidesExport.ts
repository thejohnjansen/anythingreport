import { app, HttpRequest, HttpResponseInit } from '@azure/functions';
import { generateReportPpt, type GenerateReportPptInput } from '../reportPptService';
import { requireMicrosoftUser } from '../requestAuth';

function errorResponse(message: string, status = 500): HttpResponseInit {
  return {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ error: message }, null, 2)
  };
}

app.http('exportSlidesPpt', {
  methods: ['POST'],
  route: 'slides/export',
  authLevel: 'anonymous',
  handler: async (request: HttpRequest): Promise<HttpResponseInit> => {
    const accessError = requireMicrosoftUser(request);
    if (accessError) {
      return accessError;
    }

    const body = (await request.json()) as Partial<GenerateReportPptInput>;
    if (!Array.isArray(body.slides)) {
      return errorResponse('slides array is required.', 400);
    }

    try {
      const buffer = await generateReportPpt({
        hasMidpoint: !!body.hasMidpoint,
        slides: body.slides,
        topOfMindHtml: body.topOfMindHtml,
        pipeline: body.pipeline,
        teamPipelines: body.teamPipelines,
        baseSlideTitle: body.baseSlideTitle,
        baseTeamName: body.baseTeamName,
        deckFileName: body.deckFileName,
        linkBase: body.linkBase,
        cycleNumber: body.cycleNumber,
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
    } catch (error) {
      return errorResponse(error instanceof Error ? error.message : 'Unknown error');
    }
  }
});
