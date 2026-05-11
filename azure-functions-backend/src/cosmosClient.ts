import { CosmosClient, Container, Database } from '@azure/cosmos';
import type { CosmosSettings } from './types';

let client: CosmosClient | null = null;
let database: Database | null = null;
let container: Container | null = null;

function getSettingsFromEnv(): CosmosSettings {
  const endpoint = process.env.COSMOS_ENDPOINT;
  const key = process.env.COSMOS_KEY;
  const databaseName = process.env.COSMOS_DATABASE;
  const containerName = process.env.COSMOS_CONTAINER;

  if (!endpoint || !key || !databaseName || !containerName) {
    throw new Error('Missing Cosmos configuration. Set COSMOS_ENDPOINT, COSMOS_KEY, COSMOS_DATABASE, and COSMOS_CONTAINER.');
  }

  return { endpoint, key, databaseName, containerName };
}

export function getCosmosContainer(): Container {
  if (container) {
    return container;
  }

  const settings = getSettingsFromEnv();
  client = new CosmosClient({ endpoint: settings.endpoint, key: settings.key });
  database = client.database(settings.databaseName);
  container = database.container(settings.containerName);
  return container;
}

export function buildStoryId(workItemId: number, cycle: string): string {
  return `${cycle}:${workItemId}`;
}