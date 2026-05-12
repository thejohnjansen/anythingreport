import { adoFetchJson, getAdoAuthHeader } from './adoClient';
import type { NormalizedWorkItem } from './types';

type WorkItemQueryFilter = {
  iterationPath?: string;
  areaPath?: string;
};

type AdoConfig = {
  organizationUrl: string;
  project: string;
};

type WiqlResponse = {
  workItems?: Array<{ id: number }>;
};

type WorkItemsBatchResponse = {
  value: Array<{
    id: number;
    fields?: {
      'System.Title'?: string;
      'System.AreaPath'?: string;
    };
  }>;
};

function getAdoConfigFromEnv(): AdoConfig {
  const organizationUrl = process.env.ADO_ORGANIZATION_URL;
  const project = process.env.ADO_PROJECT;

  if (!organizationUrl || !project) {
    throw new Error('Missing ADO configuration. Set ADO_ORGANIZATION_URL and ADO_PROJECT.');
  }

  return {
    organizationUrl: organizationUrl.replace(/\/+$/, ''),
    project
  };
}

function escapeWiqlString(value: string): string {
  return value.replace(/'/g, "''");
}

function buildWhereClause(filter: WorkItemQueryFilter): string {
  const clauses: string[] = ["[System.TeamProject] = @project"];

  if (filter.iterationPath) {
    clauses.push(`[System.IterationPath] UNDER '${escapeWiqlString(filter.iterationPath)}'`);
  }
  if (filter.areaPath) {
    clauses.push(`[System.AreaPath] UNDER '${escapeWiqlString(filter.areaPath)}'`);
  }

  return clauses.join(' AND ');
}

export async function fetchWorkItems(filter: WorkItemQueryFilter): Promise<NormalizedWorkItem[]> {
  if (!filter.iterationPath && !filter.areaPath) {
    throw new Error('Either iterationPath or areaPath must be provided.');
  }

  const config = getAdoConfigFromEnv();
  const authHeader = getAdoAuthHeader();
  const whereClause = buildWhereClause(filter);

  const wiqlBody = {
    query: `
      SELECT [System.Id]
      FROM WorkItems
      WHERE ${whereClause}
      ORDER BY [System.ChangedDate] DESC
    `
  };

  const wiqlUrl = `${config.organizationUrl}/${config.project}/_apis/wit/wiql?api-version=7.1`;
  const wiqlResponse = await adoFetchJson<WiqlResponse>(wiqlUrl, {
    method: 'POST',
    headers: {
      Authorization: authHeader,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(wiqlBody)
  });

  const ids = (wiqlResponse.workItems ?? []).map((item) => item.id).filter((id) => Number.isFinite(id));
  if (ids.length === 0) {
    return [];
  }

  const fields = ['System.Id', 'System.Title', 'System.AreaPath'].join(',');
  const itemsUrl = `${config.organizationUrl}/${config.project}/_apis/wit/workitems?ids=${ids.join(',')}&fields=${encodeURIComponent(fields)}&api-version=7.1`;
  const detailsResponse = await adoFetchJson<WorkItemsBatchResponse>(itemsUrl, {
    method: 'GET',
    headers: { Authorization: authHeader }
  });

  return detailsResponse.value.map((item) => ({
    workItemId: item.id,
    title: item.fields?.['System.Title'] ?? '',
    area: item.fields?.['System.AreaPath'] ?? ''
  }));
}
