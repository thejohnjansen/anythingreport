/**
 * Anything Report — server
 *
 * Runs an ADO tree query, fetches work-item details, and returns
 * structured slide data.  Auth comes from `az account get-access-token`.
 */

const express = require('express');
const { execSync } = require('child_process');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3456;

app.use(express.json());
app.use(express.static(__dirname));

/* ── Azure CLI token (cached) ─────────────────────────────────── */

let tokenCache = { token: null, expiry: 0 };

function getToken() {
    if (tokenCache.token && Date.now() < tokenCache.expiry) return tokenCache.token;
    const json = execSync(
        'az account get-access-token --resource 499b84ac-1321-427f-aa17-267ca6975798',
        { encoding: 'utf8', timeout: 30_000 }
    );
    const { accessToken, expiresOn } = JSON.parse(json);
    tokenCache = { token: accessToken, expiry: new Date(expiresOn).getTime() - 60_000 };
    return accessToken;
}

/* ── URL parsing ──────────────────────────────────────────────── */

function parseQueryUrl(raw) {
    const s = raw.trim().replace(/\/+$/, '');

    // Plain GUID
    if (/^[a-f0-9-]{36}$/i.test(s)) {
        return { baseUrl: 'https://microsoft.visualstudio.com', project: 'Edge', queryId: s };
    }

    // visualstudio.com
    const vs = s.match(
        /https?:\/\/([^/]+\.visualstudio\.com)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i
    );
    if (vs) return { baseUrl: `https://${vs[1]}`, project: vs[2], queryId: vs[3] };

    // dev.azure.com
    const dev = s.match(
        /https?:\/\/dev\.azure\.com\/([^/]+)\/([^/]+)\/_queries\/query(?:-edit)?\/([a-f0-9-]{36})/i
    );
    if (dev) return { baseUrl: `https://dev.azure.com/${dev[1]}`, project: dev[2], queryId: dev[3] };

    throw new Error('Could not parse ADO query URL. Expected an ADO query link or a query GUID.');
}

/* ── ADO helpers ──────────────────────────────────────────────── */

async function adoFetch(url, token) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`ADO ${res.status}: ${res.statusText} — ${url}`);
    return res.json();
}

async function runQuery(baseUrl, project, queryId, token) {
    return adoFetch(
        `${baseUrl}/${project}/_apis/wit/wiql/${queryId}?api-version=7.0`,
        token
    );
}

async function fetchWorkItems(baseUrl, project, ids, token) {
    if (!ids.length) return {};
    const fields = [
        'System.Id', 'System.Title', 'System.WorkItemType',
        'System.State', 'System.AssignedTo',
        'OSG.RiskAssessment', 'OSG.RiskAssessmentComment',
        'Microsoft.VSTS.Scheduling.OriginalEstimate'
    ].join(',');

    const map = {};
    for (let i = 0; i < ids.length; i += 200) {
        const batch = ids.slice(i, i + 200);
        const data = await adoFetch(
            `${baseUrl}/${project}/_apis/wit/workitems?ids=${batch.join(',')}&fields=${fields}&api-version=7.0`,
            token
        );
        for (const wi of data.value) map[wi.id] = wi;
    }
    return map;
}

/* ── Revision history lookup ───────────────────────────────────── */

/**
 * Fetch all revisions for a work item and return the field values
 * that were active on `asOfDate` (the last revision on or before that date).
 * Only the risk-related fields are extracted.
 */
async function getFieldsAsOf(baseUrl, project, workItemId, asOfDate, token) {
    const url = `${baseUrl}/${project}/_apis/wit/workitems/${workItemId}/revisions?api-version=7.0`;
    const data = await adoFetch(url, token);

    const cutoff = new Date(asOfDate);
    cutoff.setUTCHours(23, 59, 59, 999);          // include the whole day

    let best = null;
    for (const rev of data.value || []) {
        const revDate = new Date(rev.fields['System.ChangedDate']);
        if (revDate <= cutoff) best = rev;
    }

    if (!best) best = data.value?.[0];             // fallback: earliest revision

    return {
        risk:        best?.fields?.['OSG.RiskAssessment'] || '',
        riskComment: best?.fields?.['OSG.RiskAssessmentComment'] || ''
    };
}

/**
 * For every leaf work-item ID, look up historical risk fields.
 * Returns a Map<id, { risk, riskComment }>.
 */
async function batchHistoricalFields(baseUrl, project, ids, asOfDate, token) {
    const map = {};
    // Run in parallel, up to 10 at a time to be polite to ADO
    const chunks = [];
    for (let i = 0; i < ids.length; i += 10) chunks.push(ids.slice(i, i + 10));

    for (const chunk of chunks) {
        const results = await Promise.all(
            chunk.map(id => getFieldsAsOf(baseUrl, project, id, asOfDate, token))
        );
        chunk.forEach((id, i) => { map[id] = results[i]; });
    }
    return map;
}

/* ── Build slide hierarchy ────────────────────────────────────── */

function buildSlides(queryResult, workItemMap, midpointMap) {
    const roots = [];
    const children = {};

    for (const rel of queryResult.workItemRelations || []) {
        if (!rel.source) {
            roots.push(rel.target.id);
        } else {
            (children[rel.source.id] ||= []).push(rel.target.id);
        }
    }

    const slides = [];
    for (const rootId of roots) {
        for (const l2Id of children[rootId] || []) {
            const wi = workItemMap[l2Id];
            const rows = (children[l2Id] || [])
                .map(id => workItemMap[id])
                .filter(Boolean);

            slides.push({
                title: wi?.fields?.['System.Title'] || `Work Item ${l2Id}`,
                id: l2Id,
                items: rows.map(w => {
                    const item = {
                        id:          w.id,
                        title:       w.fields['System.Title'] || '',
                        risk:        w.fields['OSG.RiskAssessment'] || '',
                        riskComment: w.fields['OSG.RiskAssessmentComment'] || '',
                        state:       w.fields['System.State'] || '',
                        assignedTo:  w.fields['System.AssignedTo']?.displayName || ''
                    };
                    if (midpointMap) {
                        const snap = midpointMap[w.id] || {};
                        item.midpointRisk    = snap.risk || '';
                        item.midpointComment = snap.riskComment || '';
                    }
                    return item;
                })
            });
        }
    }
    return slides;
}

/* ── API route ────────────────────────────────────────────────── */

app.post('/api/slides', async (req, res) => {
    try {
        const { queryUrl, midpointDate } = req.body;
        if (!queryUrl) return res.status(400).json({ error: 'queryUrl is required' });

        const { baseUrl, project, queryId } = parseQueryUrl(queryUrl);
        const token = getToken();

        // 1. Run the tree query
        const queryResult = await runQuery(baseUrl, project, queryId, token);

        // 2. Collect every referenced work-item ID & identify leaf IDs
        const allIds = new Set();
        const leafIds = new Set();
        const parentIds = new Set();
        for (const rel of queryResult.workItemRelations || []) {
            if (rel.target) allIds.add(rel.target.id);
            if (rel.source) parentIds.add(rel.source.id);
        }
        for (const id of allIds) {
            if (!parentIds.has(id)) leafIds.add(id);
        }

        // 3. Batch-fetch current work-item details
        const workItemMap = await fetchWorkItems(baseUrl, project, [...allIds], token);

        // 4. If midpoint date provided, fetch historical snapshots for leaf items
        let midpointMap = null;
        if (midpointDate) {
            midpointMap = await batchHistoricalFields(
                baseUrl, project, [...leafIds], midpointDate, token
            );
        }

        // 5. Build slides
        const slides = buildSlides(queryResult, workItemMap, midpointMap);

        // 6. Provide a link prefix so the UI can link to work items
        const linkBase = `${baseUrl}/${project}/_workitems/edit/`;

        res.json({ slides, linkBase, hasMidpoint: !!midpointDate });
    } catch (err) {
        console.error('❌', err);
        res.status(500).json({ error: err.message });
    }
});

/* ── Build doc hierarchy ──────────────────────────────────────── */

function buildDocSections(queryResult, workItemMap) {
    const roots = [];
    const children = {};

    for (const rel of queryResult.workItemRelations || []) {
        if (!rel.source) {
            roots.push(rel.target.id);
        } else {
            (children[rel.source.id] ||= []).push(rel.target.id);
        }
    }

    const sections = [];
    for (const rootId of roots) {
        const rootWi = workItemMap[rootId];
        const groups = [];

        for (const l2Id of children[rootId] || []) {
            const wi = workItemMap[l2Id];
            const items = (children[l2Id] || [])
                .map(id => workItemMap[id])
                .filter(Boolean)
                .map(w => ({
                    id:               w.id,
                    title:            w.fields['System.Title'] || '',
                    originalEstimate: w.fields['Microsoft.VSTS.Scheduling.OriginalEstimate'] ?? '',
                    dri:              w.fields['System.AssignedTo']?.displayName || '',
                    state:            w.fields['System.State'] || ''
                }));

            groups.push({
                title: wi?.fields?.['System.Title'] || `Work Item ${l2Id}`,
                id: l2Id,
                items
            });
        }

        sections.push({
            objectiveTitle: rootWi?.fields?.['System.Title'] || `Objective ${rootId}`,
            objectiveId: rootId,
            groups
        });
    }
    return sections;
}

/* ── Doc API route ────────────────────────────────────────────── */

app.post('/api/doc', async (req, res) => {
    try {
        const { queryUrl } = req.body;
        if (!queryUrl) return res.status(400).json({ error: 'queryUrl is required' });

        const { baseUrl, project, queryId } = parseQueryUrl(queryUrl);
        const token = getToken();

        const queryResult = await runQuery(baseUrl, project, queryId, token);

        const allIds = new Set();
        for (const rel of queryResult.workItemRelations || []) {
            if (rel.target) allIds.add(rel.target.id);
        }

        const workItemMap = await fetchWorkItems(baseUrl, project, [...allIds], token);
        const sections = buildDocSections(queryResult, workItemMap);
        const linkBase = `${baseUrl}/${project}/_workitems/edit/`;

        res.json({ sections, linkBase });
    } catch (err) {
        console.error('❌', err);
        res.status(500).json({ error: err.message });
    }
});

/* ── Start ────────────────────────────────────────────────────── */

app.listen(PORT, () => {
    console.log(`\n  Anything Report → http://localhost:${PORT}\n`);
});
