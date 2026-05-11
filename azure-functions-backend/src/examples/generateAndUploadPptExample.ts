import { generatePpt } from '../pptService';
import { uploadPptToGraph } from '../graphUploadService';

async function runExample(): Promise<void> {
  const accessToken = process.env.GRAPH_ACCESS_TOKEN;
  if (!accessToken) {
    throw new Error('Set GRAPH_ACCESS_TOKEN before running this example.');
  }

  const stories = [
    {
      workItemId: 2001,
      title: 'Improve onboarding workflow',
      cycle: 'FY26\\Q4',
      area: 'Edge\\Experience',
      narrative: 'Introduced a guided flow and reduced drop-off during setup by 22%.',
      pmComments: 'Track conversion for another sprint before broad rollout.'
    },
    {
      workItemId: 2002,
      title: 'Stabilize release health checks',
      cycle: 'FY26\\Q4',
      area: 'Edge\\Platform',
      narrative: 'Consolidated health checks and removed duplicate alerts across environments.',
      pmComments: 'Document escalation playbook updates.'
    }
  ];

  const pptBuffer = await generatePpt(stories);

  const fileUrl = await uploadPptToGraph({
    accessToken,
    fileName: `stories-${Date.now()}.pptx`,
    fileBuffer: pptBuffer,
    target: {
      mode: 'delegated',
      folderPath: 'AnythingReport/Exports'
    }
  });

  console.log('Uploaded PPT URL:', fileUrl);
}

void runExample();
