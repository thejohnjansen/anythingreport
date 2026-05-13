"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getDocSections = getDocSections;
exports.getRetroRows = getRetroRows;
const adoClient_1 = require("./adoClient");
function buildTree(queryResult) {
    const roots = [];
    const children = {};
    for (const relation of queryResult.workItemRelations || []) {
        if (!relation.target) {
            continue;
        }
        if (!relation.source) {
            roots.push(relation.target.id);
        }
        else {
            children[relation.source.id] ??= [];
            children[relation.source.id].push(relation.target.id);
        }
    }
    return { roots, children };
}
function parseAreaLevel4(areaPath) {
    const parts = String(areaPath || '').split('\\').filter(Boolean);
    return parts[3] || '';
}
function riskAssessmentToRag(risk) {
    const value = String(risk || '').toLowerCase();
    if (value.includes('on track')) {
        return 'green';
    }
    if (value.includes('at risk')) {
        return 'yellow';
    }
    if (value.includes('off track')) {
        return 'red';
    }
    return '';
}
function buildDocSections(queryResult, workItemMap) {
    const { roots, children } = buildTree(queryResult);
    return roots.map((rootId) => {
        const rootWi = workItemMap[rootId];
        const groups = (children[rootId] || []).map((l2Id) => {
            const wi = workItemMap[l2Id];
            const items = (children[l2Id] || [])
                .map((id) => workItemMap[id])
                .filter(Boolean)
                .map((workItem) => ({
                id: workItem.id,
                title: workItem.fields['System.Title'] || '',
                originalEstimate: workItem.fields['Microsoft.VSTS.Scheduling.OriginalEstimate'] ?? '',
                dri: workItem.fields['System.AssignedTo']?.displayName || '',
                state: workItem.fields['System.State'] || ''
            }));
            return {
                title: wi?.fields['System.Title'] || `Work Item ${l2Id}`,
                id: l2Id,
                dri: wi?.fields['System.AssignedTo']?.displayName || '',
                originalEstimate: wi?.fields['Microsoft.VSTS.Scheduling.OriginalEstimate'] ?? '',
                comments: wi?.fields['OSG.OverallComments'] || '',
                items
            };
        });
        return {
            objectiveTitle: rootWi?.fields['System.Title'] || `Objective ${rootId}`,
            objectiveId: rootId,
            groups
        };
    });
}
function buildRetroRows(queryResult, workItemMap) {
    const { roots, children } = buildTree(queryResult);
    const rows = [];
    for (const rootId of roots) {
        for (const l2Id of children[rootId] || []) {
            const wi = workItemMap[l2Id];
            const l3Items = (children[l2Id] || []).map((id) => workItemMap[id]).filter(Boolean);
            const l2Risk = wi?.fields['OSG.RiskAssessment'] || '';
            let rag = riskAssessmentToRag(l2Risk);
            if (!rag) {
                const childRags = l3Items.map((workItem) => riskAssessmentToRag(workItem.fields['OSG.RiskAssessment'] || ''));
                if (childRags.includes('red')) {
                    rag = 'red';
                }
                else if (childRags.includes('yellow')) {
                    rag = 'yellow';
                }
                else if (childRags.includes('green')) {
                    rag = 'green';
                }
            }
            rows.push({
                id: l2Id,
                rag,
                featureArea: parseAreaLevel4(wi?.fields['System.AreaPath'] || ''),
                mission: wi?.fields['System.Title'] || `Work Item ${l2Id}`,
                milestones: l3Items.map((workItem) => ({
                    id: workItem.id,
                    title: workItem.fields['System.Title'] || '',
                    risk: workItem.fields['OSG.RiskAssessment'] || ''
                })),
                justifications: l3Items.map((workItem) => {
                    const comment = workItem.fields['OSG.RiskAssessmentComment'] || '';
                    return comment.length > 200 ? `${comment.slice(0, 200)}…` : comment;
                })
            });
        }
    }
    return rows;
}
async function getQueryWorkItemMap(queryUrl, incomingBearerToken) {
    const { baseUrl, project, queryId } = (0, adoClient_1.parseQueryUrl)(queryUrl);
    const queryResult = await (0, adoClient_1.runTreeQuery)(baseUrl, project, queryId, incomingBearerToken);
    const allIds = new Set();
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.target) {
            allIds.add(relation.target.id);
        }
    }
    const workItemMap = await (0, adoClient_1.fetchWorkItemsByIds)(baseUrl, project, [...allIds], incomingBearerToken);
    return { baseUrl, project, queryResult, workItemMap };
}
async function getDocSections(queryUrl, incomingBearerToken) {
    const { baseUrl, project, queryResult, workItemMap } = await getQueryWorkItemMap(queryUrl, incomingBearerToken);
    return {
        sections: buildDocSections(queryResult, workItemMap),
        linkBase: `${baseUrl}/${project}/_workitems/edit/`
    };
}
async function getRetroRows(queryUrl, incomingBearerToken) {
    const { baseUrl, project, queryResult, workItemMap } = await getQueryWorkItemMap(queryUrl, incomingBearerToken);
    return {
        rows: buildRetroRows(queryResult, workItemMap),
        linkBase: `${baseUrl}/${project}/_workitems/edit/`
    };
}
