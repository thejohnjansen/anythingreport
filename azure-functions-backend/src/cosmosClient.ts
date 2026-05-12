import { CosmosClient, Container, Database } from '@azure/cosmos';
import type { CosmosSettings } from './types';

let client: CosmosClient | null = null;
let database: Database | null = null;
let container: Container | null = null;

function parseConnectionString(connectionString: string): Pick<CosmosSettings, 'endpoint' | 'key'> {
  const pairs = connectionString
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const separatorIndex = part.indexOf('=');
      if (separatorIndex < 0) {
        return null;
      }
      const name = part.slice(0, separatorIndex).trim().toLowerCase();
      const value = part.slice(separatorIndex + 1).trim();
      return [name, value] as const;
    })
    .filter((entry): entry is readonly [string, string] => entry !== null);

  const values = new Map<string, string>(pairs);
  const endpoint = values.get('accountendpoint');
  const key = values.get('accountkey');

  if (!endpoint || !key) {
    throw new Error(
      'Invalid COSMOS_CONNECTION_STRING. Expected AccountEndpoint and AccountKey entries.'
    );
  }

  return { endpoint, key };
}

function getSettingsFromEnv(): CosmosSettings {
  const connectionString = process.env.COSMOS_CONNECTION_STRING;
  const databaseName = process.env.COSMOS_DATABASE;
  const containerName = process.env.COSMOS_CONTAINER;

  if (!connectionString || !databaseName || !containerName) {
    throw new Error(
      'Missing Cosmos configuration. Set COSMOS_CONNECTION_STRING, COSMOS_DATABASE, and COSMOS_CONTAINER.'
    );
  }

  const { endpoint, key } = parseConnectionString(connectionString);

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