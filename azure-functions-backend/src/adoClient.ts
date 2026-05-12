import { execSync } from 'node:child_process';

const ADO_RESOURCE = '499b84ac-1321-427f-aa17-267ca6975798';

let tokenCache: { token: string | null; expiry: number } = { token: null, expiry: 0 };

export interface ParsedAdoQueryUrl {
  baseUrl: string;
  project: string;
  queryId: string;
}

export interface WorkItemFields {
  'System.Id'?: number;
  'System.Title'?: string;
  'System.WorkItemType'?: string;
  'System.State'?: string;
  'System.AssignedTo'?: { displayName?: string };
  'System.IterationPath'?: string;
  'System.AreaPath'?: string;
  'OSG.RiskAssessment'?: string;
  'OSG.RiskAssessmentComment'?: string;
  'Microsoft.VSTS.Scheduling.OriginalEstimate'?: number;
  'OSG.OverallComments'?: string;
  'System.ChangedDate'?: string;
}

export interface AzureDevOpsWorkItem {
  id: number;
  fields: WorkItemFields;
}

export interface WiqlTreeRelation {
  source?: { id: number };
  target?: { id: number };
}

export interface WiqlTreeResult {
  workItemRelations?: WiqlTreeRelation[];
}

function getPatFromEnv(): string | null {
  return process.env.ADO_PAT || null;
}

function getCachedAzureCliToken(): string | null {
  if (tokenCache.token && Date.now() < tokenCache.expiry) {
    return tokenCache.token;
  }

  try {
    const json = execSync(
      `az account get-access-token --resource ${ADO_RESOURCE}`,
      { encoding: 'utf8', timeout: 30000 }
    );
    const parsed = JSON.parse(json) as { accessToken?: string; expiresOn?: string };
    if (!parsed.accessToken || !parsed.expiresOn) {
      return null;
    }

    tokenCache = {
      token: parsed.accessToken,
      expiry: new Date(parsed.expiresOn).getTime() - 60000
    };
    return tokenCache.token;
  } catch {
    return null;
  }
}

export function getAdoAuthHeader(): string {
  const azureCliToken = getCachedAzureCliToken();
  if (azureCliToken) {
    return `Bearer ${azureCliToken}`;
  }

  const pat = getPatFromEnv();
  if (pat) {
    return `Basic ${Buffer.from(`:${pat}`).toString('base64')}`;
  }

  throw new Error('Unable to authenticate to Azure DevOps. Sign in with az login or set ADO_PAT.');
}

export function parseQueryUrl(raw: string): ParsedAdoQueryUrl {
  const value = raw.trim().replace(/\/+$/, '');

  if (/^[a-f0-9-]{36}$/i.test(value)) {
    return { baseUrl: 'https://microsoft.visualstudio.com', project: 'Edge', queryId: value };
  }

  const visualStudioMatch = value.match(
    /https?:\/\/([^/]+\.visualstudio\.com)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i
  );
  if (visualStudioMatch) {
    return {
      baseUrl: `https://${visualStudioMatch[1]}`,
      project: visualStudioMatch[2],
      queryId: visualStudioMatch[3]
    };
  }

  const devAzureMatch = value.match(
    /https?:\/\/dev\.azure\.com\/([^/]+)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i
  );
  if (devAzureMatch) {
    return {
      baseUrl: `https://dev.azure.com/${devAzureMatch[1]}`,
      project: devAzureMatch[2],
      queryId: devAzureMatch[3]
    };
  }

  throw new Error('Could not parse ADO query URL. Expected an ADO query link or a query GUID.');
}

export async function adoFetchJson<T>(url: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers);
  if (!headers.has('Authorization')) {
    headers.set('Authorization', getAdoAuthHeader());
  }

  const response = await fetch(url, { ...init, headers });
  const bodyText = await response.text();
  const contentType = String(response.headers.get('content-type') || '').toLowerCase();

  if (!response.ok) {
    const snippet = bodyText.replace(/\s+/g, ' ').trim().slice(0, 300);
    throw new Error(`ADO ${response.status}: ${response.statusText}${snippet ? ` - ${snippet}` : ''}`);
  }

  if (!contentType.includes('application/json')) {
    const snippet = bodyText.replace(/\s+/g, ' ').trim().slice(0, 300);
    throw new Error(`ADO non-JSON response (${contentType || 'unknown content-type'})${snippet ? ` - ${snippet}` : ''}`);
  }

  return JSON.parse(bodyText) as T;
}

export async function runTreeQuery(baseUrl: string, project: string, queryId: string): Promise<WiqlTreeResult> {
  return adoFetchJson<WiqlTreeResult>(
    `${baseUrl}/${project}/_apis/wit/wiql/${queryId}?api-version=7.0`
  );
}

export async function fetchWorkItemsByIds(
  baseUrl: string,
  project: string,
  ids: number[]
): Promise<Record<number, AzureDevOpsWorkItem>> {
  if (!ids.length) {
    return {};
  }

  const fields = [
    'System.Id',
    'System.Title',
    'System.WorkItemType',
    'System.State',
    'System.AssignedTo',
    'System.IterationPath',
    'System.AreaPath',
    'OSG.RiskAssessment',
    'OSG.RiskAssessmentComment',
    'Microsoft.VSTS.Scheduling.OriginalEstimate',
    'OSG.OverallComments'
  ].join(',');

  const items: Record<number, AzureDevOpsWorkItem> = {};
  for (let index = 0; index < ids.length; index += 200) {
    const batch = ids.slice(index, index + 200);
    const data = await adoFetchJson<{ value?: AzureDevOpsWorkItem[] }>(
      `${baseUrl}/${project}/_apis/wit/workitems?ids=${batch.join(',')}&fields=${fields}&api-version=7.0`
    );
    for (const workItem of data.value ?? []) {
      items[workItem.id] = workItem;
    }
  }

  return items;
}
