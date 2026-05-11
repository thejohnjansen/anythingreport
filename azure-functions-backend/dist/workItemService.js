"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.fetchWorkItems = fetchWorkItems;
function getAdoConfigFromEnv() {
    const organizationUrl = process.env.ADO_ORGANIZATION_URL;
    const project = process.env.ADO_PROJECT;
    const pat = process.env.ADO_PAT;
    if (!organizationUrl || !project || !pat) {
        throw new Error('Missing ADO configuration. Set ADO_ORGANIZATION_URL, ADO_PROJECT, and ADO_PAT.');
    }
    return {
        organizationUrl: organizationUrl.replace(/\/+$/, ''),
        project,
        pat
    };
}
function getAuthHeaderValue(pat) {
    // Azure DevOps PAT auth uses basic auth with empty username and PAT as password.
    return `Basic ${Buffer.from(`:${pat}`).toString('base64')}`;
}
function escapeWiqlString(value) {
    return value.replace(/'/g, "''");
}
function buildWhereClause(filter) {
    const clauses = ["[System.TeamProject] = @project"];
    if (filter.iterationPath) {
        clauses.push(`[System.IterationPath] UNDER '${escapeWiqlString(filter.iterationPath)}'`);
    }
    if (filter.areaPath) {
        clauses.push(`[System.AreaPath] UNDER '${escapeWiqlString(filter.areaPath)}'`);
    }
    return clauses.join(' AND ');
}
async function adoFetchJson(url, init) {
    const response = await fetch(url, init);
    const bodyText = await response.text();
    if (!response.ok) {
        const snippet = bodyText.replace(/\s+/g, ' ').trim().slice(0, 300);
        throw new Error(`ADO ${response.status}: ${response.statusText}${snippet ? ` - ${snippet}` : ''}`);
    }
    return JSON.parse(bodyText);
}
async function fetchWorkItems(filter) {
    if (!filter.iterationPath && !filter.areaPath) {
        throw new Error('Either iterationPath or areaPath must be provided.');
    }
    const config = getAdoConfigFromEnv();
    const authHeader = getAuthHeaderValue(config.pat);
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
    const wiqlResponse = await adoFetchJson(wiqlUrl, {
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
    const detailsResponse = await adoFetchJson(itemsUrl, {
        method: 'GET',
        headers: { Authorization: authHeader }
    });
    return detailsResponse.value.map((item) => ({
        workItemId: item.id,
        title: item.fields?.['System.Title'] ?? '',
        area: item.fields?.['System.AreaPath'] ?? ''
    }));
}
