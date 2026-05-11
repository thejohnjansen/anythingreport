import { writeFile } from 'node:fs/promises';
import { generatePpt } from '../pptService';

async function runExample(): Promise<void> {
  const stories = [
    {
      workItemId: 1001,
      title: 'Reduce deployment friction',
      cycle: 'FY26\Q4',
      area: 'Edge\\Platform',
      narrative: 'This cycle we removed two manual gates and reduced deployment time by 35%.',
      pmComments: 'Need one more sprint to harden rollback workflow.'
    },
    {
      workItemId: 1002,
      title: 'Unify reporting views',
      cycle: 'FY26\\Q4',
      area: 'Edge\\Insights',
      narrative: 'Narrative templates are now shared and rendered consistently across teams.',
      pmComments: 'Pilot feedback is positive; publish guidance doc next week.'
    }
  ];

  const pptBuffer = await generatePpt(stories);
  await writeFile('stories-export.pptx', pptBuffer);
}

void runExample();
