import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { CosmosClient } from '@azure/cosmos';
import type { StoryUpsertRequest } from '../types';

async function loadLocalSettings(): Promise<void> {
  const localSettingsPath = path.resolve(process.cwd(), 'local.settings.json');
  const raw = await readFile(localSettingsPath, 'utf8');
  const parsed = JSON.parse(raw) as { Values?: Record<string, string> };

  for (const [key, value] of Object.entries(parsed.Values ?? {})) {
    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable ${name}.`);
  }
  return value;
}

function getSampleStories(): StoryUpsertRequest[] {
  return [
    {
      workItemId: 4101,
      cycle: 'FY26-Q4',
      title: 'Improve onboarding workflow',
      area: 'Edge\\Experience',
      pmComments: 'Pilot is healthy; watch activation rate before broad release.',
      narrative: 'The team simplified first-run setup and reduced user drop-off by removing two unnecessary approval steps.',
      lastEditedBy: 'local.seed'
    },
    {
      workItemId: 4102,
      cycle: 'FY26-Q4',
      title: 'Stabilize release health checks',
      area: 'Edge\\Platform',
      pmComments: 'Escalation playbook still needs final review.',
      narrative: 'Health checks were consolidated into one release gate, cutting duplicate alerts and improving operator confidence during rollout.',
      lastEditedBy: 'local.seed'
    },
    {
      workItemId: 4103,
      cycle: 'FY26-Q4',
      title: 'Unify reporting views',
      area: 'Edge\\Insights',
      pmComments: 'Messaging is strong; add one slide on adoption impact.',
      narrative: 'Reporting summaries now use a shared template so leadership sees consistent narrative structure across product groups.',
      lastEditedBy: 'local.seed'
    },
    {
      workItemId: 4104,
      cycle: 'FY26-Q4',
      title: 'Reduce deployment friction',
      area: 'Edge\\Operations',
      pmComments: 'Rollback experience improved, but still needs one more validation run.',
      narrative: 'Automation removed manual environment handoffs and shortened average deployment lead time for the release train.',
      lastEditedBy: 'local.seed'
    },
    {
      workItemId: 4105,
      cycle: 'FY26-Q4',
      title: 'Increase service observability',
      area: 'Edge\\Platform',
      pmComments: 'Keep refining error-budget summary for executive readout.',
      narrative: 'The team added targeted telemetry and dashboard coverage so incidents can be triaged faster with clearer customer impact signals.',
      lastEditedBy: 'local.seed'
    }
  ];
}

async function run(): Promise<void> {
  await loadLocalSettings();

  const connectionString = requireEnv('COSMOS_CONNECTION_STRING');
  const databaseName = requireEnv('COSMOS_DATABASE');
  const containerName = requireEnv('COSMOS_CONTAINER');

  const client = new CosmosClient(connectionString);
  const { database } = await client.databases.createIfNotExists({ id: databaseName });
  const { container } = await database.containers.createIfNotExists({
    id: containerName,
    partitionKey: { paths: ['/cycle'] }
  });

  const now = new Date().toISOString();
  for (const story of getSampleStories()) {
    const document = {
      id: `${story.cycle}:${story.workItemId}`,
      workItemId: story.workItemId,
      cycle: story.cycle,
      title: story.title,
      area: story.area,
      pmComments: story.pmComments ?? '',
      narrative: story.narrative,
      lastEditedBy: story.lastEditedBy,
      lastEditedAt: now,
      version: 1
    };

    await container.items.upsert(document);
    console.log(`Seeded ${document.id}`);
  }
}

void run();