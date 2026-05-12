"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getAdoAuthHeader = getAdoAuthHeader;
exports.parseQueryUrl = parseQueryUrl;
exports.adoFetchJson = adoFetchJson;
exports.runTreeQuery = runTreeQuery;
exports.fetchWorkItemsByIds = fetchWorkItemsByIds;
const node_child_process_1 = require("node:child_process");
const ADO_RESOURCE = '499b84ac-1321-427f-aa17-267ca6975798';
let tokenCache = { token: null, expiry: 0 };
function getPatFromEnv() {
    return process.env.ADO_PAT || null;
}
function getCachedAzureCliToken() {
    if (tokenCache.token && Date.now() < tokenCache.expiry) {
        return tokenCache.token;
    }
    try {
        const json = (0, node_child_process_1.execSync)(`az account get-access-token --resource ${ADO_RESOURCE}`, { encoding: 'utf8', timeout: 30000 });
        const parsed = JSON.parse(json);
        if (!parsed.accessToken || !parsed.expiresOn) {
            return null;
        }
        tokenCache = {
            token: parsed.accessToken,
            expiry: new Date(parsed.expiresOn).getTime() - 60000
        };
        return tokenCache.token;
    }
    catch {
        return null;
    }
}
function getAdoAuthHeader() {
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
function parseQueryUrl(raw) {
    const value = raw.trim().replace(/\/+$/, '');
    if (/^[a-f0-9-]{36}$/i.test(value)) {
        return { baseUrl: 'https://microsoft.visualstudio.com', project: 'Edge', queryId: value };
    }
    const visualStudioMatch = value.match(/https?:\/\/([^/]+\.visualstudio\.com)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i);
    if (visualStudioMatch) {
        return {
            baseUrl: `https://${visualStudioMatch[1]}`,
            project: visualStudioMatch[2],
            queryId: visualStudioMatch[3]
        };
    }
    const devAzureMatch = value.match(/https?:\/\/dev\.azure\.com\/([^/]+)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i);
    if (devAzureMatch) {
        return {
            baseUrl: `https://dev.azure.com/${devAzureMatch[1]}`,
            project: devAzureMatch[2],
            queryId: devAzureMatch[3]
        };
    }
    throw new Error('Could not parse ADO query URL. Expected an ADO query link or a query GUID.');
}
async function adoFetchJson(url, init) {
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
    return JSON.parse(bodyText);
}
async function runTreeQuery(baseUrl, project, queryId) {
    return adoFetchJson(`${baseUrl}/${project}/_apis/wit/wiql/${queryId}?api-version=7.0`);
}
async function fetchWorkItemsByIds(baseUrl, project, ids) {
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
    const items = {};
    for (let index = 0; index < ids.length; index += 200) {
        const batch = ids.slice(index, index + 200);
        const data = await adoFetchJson(`${baseUrl}/${project}/_apis/wit/workitems?ids=${batch.join(',')}&fields=${fields}&api-version=7.0`);
        for (const workItem of data.value ?? []) {
            items[workItem.id] = workItem;
        }
    }
    return items;
}
