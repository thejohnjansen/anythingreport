"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getSlidesPreview = getSlidesPreview;
const adoClient_1 = require("./adoClient");
function parseIterationLevel2Token(iterationPath) {
    const parts = String(iterationPath || '').split('\\').filter(Boolean);
    const level2 = parts[1] || '';
    if (!level2) {
        return '';
    }
    const dashParts = level2.split('-');
    return dashParts.length > 1 ? dashParts.slice(1).join('-') : level2;
}
function parseAreaLevel4(areaPath) {
    const parts = String(areaPath || '').split('\\').filter(Boolean);
    return parts[3] || '';
}
function parseAreaLevel3(areaPath) {
    const parts = String(areaPath || '').split('\\').filter(Boolean);
    return parts[2] || '';
}
function findFirstLeafEpic(queryResult, workItemMap) {
    const children = {};
    const orderedTargets = [];
    const seenTargets = new Set();
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.source && relation.target) {
            children[relation.source.id] ??= [];
            children[relation.source.id].push(relation.target.id);
        }
        if (relation.target && !seenTargets.has(relation.target.id)) {
            orderedTargets.push(relation.target.id);
            seenTargets.add(relation.target.id);
        }
    }
    for (const id of orderedTargets) {
        const workItem = workItemMap[id];
        const workItemType = String(workItem?.fields?.['System.WorkItemType'] || '').toLowerCase();
        if (!workItem || !workItemType.includes('epic')) {
            continue;
        }
        if ((children[id] || []).length > 0) {
            continue;
        }
        return workItem;
    }
    return null;
}
function findLeafEpics(queryResult, workItemMap) {
    const children = {};
    const orderedTargets = [];
    const seenTargets = new Set();
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.source && relation.target) {
            children[relation.source.id] ??= [];
            children[relation.source.id].push(relation.target.id);
        }
        if (relation.target && !seenTargets.has(relation.target.id)) {
            orderedTargets.push(relation.target.id);
            seenTargets.add(relation.target.id);
        }
    }
    return orderedTargets
        .map((id) => workItemMap[id])
        .filter((workItem) => {
        const workItemType = String(workItem?.fields?.['System.WorkItemType'] || '').toLowerCase();
        return !!workItem && workItemType.includes('epic') && (children[workItem.id] || []).length === 0;
    });
}
function buildTitleContext(queryResult, workItemMap) {
    const leafEpic = findFirstLeafEpic(queryResult, workItemMap);
    if (!leafEpic) {
        return {
            baseTeamName: '',
            flatSlideTitle: 'Layout',
            flatSlidePrefix: '',
            deckFileName: 'anything-report'
        };
    }
    const iterationLevel2NoHyphen = String(leafEpic.fields['System.IterationPath'] || '')
        .split('\\')
        .filter(Boolean)[1]?.replace(/-/g, '') || '';
    const iterationToken = parseIterationLevel2Token(leafEpic.fields['System.IterationPath']);
    const areaLevel4 = parseAreaLevel4(leafEpic.fields['System.AreaPath']);
    const areaLevel3 = parseAreaLevel3(leafEpic.fields['System.AreaPath']);
    const uniqueTeams = new Set(findLeafEpics(queryResult, workItemMap)
        .map((workItem) => parseAreaLevel4(workItem.fields['System.AreaPath']))
        .filter(Boolean));
    let flatSlideTitle = 'Layout';
    if (iterationToken && areaLevel4) {
        flatSlideTitle = `${iterationToken} - ${areaLevel4}`;
    }
    else if (iterationToken) {
        flatSlideTitle = iterationToken;
    }
    else if (areaLevel4) {
        flatSlideTitle = areaLevel4;
    }
    let deckFileName = 'anything-report';
    if (uniqueTeams.size > 1 && iterationLevel2NoHyphen && areaLevel3) {
        deckFileName = `${iterationLevel2NoHyphen} ${areaLevel3}`;
    }
    else if (iterationLevel2NoHyphen && areaLevel4) {
        deckFileName = `${iterationLevel2NoHyphen} - ${areaLevel4} Check In`;
    }
    return {
        baseTeamName: uniqueTeams.size > 1 && areaLevel3 ? areaLevel3 : (areaLevel4 || ''),
        flatSlideTitle,
        flatSlidePrefix: iterationToken || '',
        deckFileName
    };
}
async function getFieldsAsOf(baseUrl, project, workItemId, asOfDate, incomingBearerToken) {
    const data = await (0, adoClient_1.adoFetchJson)(`${baseUrl}/${project}/_apis/wit/workitems/${workItemId}/revisions?api-version=7.0`, undefined, incomingBearerToken);
    const cutoff = new Date(asOfDate);
    cutoff.setUTCHours(23, 59, 59, 999);
    let best = null;
    for (const revision of data.value || []) {
        const revisionDate = new Date(revision.fields['System.ChangedDate'] || '');
        if (!Number.isNaN(revisionDate.getTime()) && revisionDate <= cutoff) {
            best = revision;
        }
    }
    if (!best) {
        best = data.value?.[0] || null;
    }
    return {
        risk: best?.fields['OSG.RiskAssessment'] || '',
        riskComment: best?.fields['OSG.RiskAssessmentComment'] || ''
    };
}
async function batchHistoricalFields(baseUrl, project, ids, asOfDate, incomingBearerToken) {
    const map = {};
    const chunks = [];
    for (let index = 0; index < ids.length; index += 10) {
        chunks.push(ids.slice(index, index + 10));
    }
    for (const chunk of chunks) {
        const results = await Promise.all(chunk.map((id) => getFieldsAsOf(baseUrl, project, id, asOfDate, incomingBearerToken)));
        chunk.forEach((id, resultIndex) => {
            map[id] = results[resultIndex];
        });
    }
    return map;
}
function buildSlides(queryResult, workItemMap, midpointMap) {
    const children = {};
    const targetIds = new Set();
    const rootIds = new Set();
    function typeContains(workItem, needle) {
        const workItemType = String(workItem?.fields?.['System.WorkItemType'] || '').toLowerCase();
        return workItemType.includes(needle);
    }
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.target) {
            targetIds.add(relation.target.id);
        }
        if (!relation.source && relation.target) {
            rootIds.add(relation.target.id);
        }
        if (relation.source && relation.target) {
            children[relation.source.id] ??= [];
            children[relation.source.id].push(relation.target.id);
        }
    }
    const slides = [];
    const additionalItems = [];
    const seenAdditional = new Set();
    const seenSlideIds = new Set();
    const objectiveIds = [];
    for (const id of targetIds) {
        const workItem = workItemMap[id];
        if (typeContains(workItem, 'objective')) {
            objectiveIds.push(id);
        }
    }
    if (objectiveIds.length === 0) {
        objectiveIds.push(...rootIds);
    }
    const candidateEpicIds = new Set();
    for (const objectiveId of objectiveIds) {
        for (const childId of children[objectiveId] || []) {
            const workItem = workItemMap[childId];
            if (workItem && typeContains(workItem, 'epic')) {
                candidateEpicIds.add(childId);
            }
        }
    }
    for (const rootId of rootIds) {
        const workItem = workItemMap[rootId];
        if (workItem && typeContains(workItem, 'epic')) {
            candidateEpicIds.add(rootId);
        }
    }
    for (const epicId of candidateEpicIds) {
        const epic = workItemMap[epicId];
        const rows = (children[epicId] || []).map((id) => workItemMap[id]).filter(Boolean);
        if (rows.length === 0) {
            if (seenAdditional.has(epicId)) {
                continue;
            }
            const item = {
                id: epicId,
                title: epic?.fields['System.Title'] || '',
                risk: epic?.fields['OSG.RiskAssessment'] || '',
                riskComment: epic?.fields['OSG.RiskAssessmentComment'] || '',
                state: epic?.fields['System.State'] || '',
                assignedTo: epic?.fields['System.AssignedTo']?.displayName || ''
            };
            if (midpointMap) {
                item.midpointRisk = midpointMap[epicId]?.risk || '';
                item.midpointComment = midpointMap[epicId]?.riskComment || '';
            }
            additionalItems.push(item);
            seenAdditional.add(epicId);
            continue;
        }
        if (seenSlideIds.has(epicId)) {
            continue;
        }
        slides.push({
            title: epic?.fields['System.Title'] || `Work Item ${epicId}`,
            id: epicId,
            items: rows.map((row) => {
                const item = {
                    id: row.id,
                    title: row.fields['System.Title'] || '',
                    risk: row.fields['OSG.RiskAssessment'] || '',
                    riskComment: row.fields['OSG.RiskAssessmentComment'] || '',
                    state: row.fields['System.State'] || '',
                    assignedTo: row.fields['System.AssignedTo']?.displayName || ''
                };
                if (midpointMap) {
                    item.midpointRisk = midpointMap[row.id]?.risk || '';
                    item.midpointComment = midpointMap[row.id]?.riskComment || '';
                }
                return item;
            })
        });
        seenSlideIds.add(epicId);
    }
    if (additionalItems.length) {
        slides.push({ title: 'Additional Work', id: 'additional-work', items: additionalItems });
    }
    return slides;
}
function buildFlatSlides(queryResult, workItemMap, midpointMap, flatSlideTitle, flatSlidePrefix) {
    const children = {};
    const targetIds = new Set();
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.target) {
            targetIds.add(relation.target.id);
        }
        if (relation.source && relation.target) {
            children[relation.source.id] ??= [];
            children[relation.source.id].push(relation.target.id);
        }
    }
    const itemsByArea = {};
    const orderedAreas = [];
    function pushItem(areaName, item) {
        if (!itemsByArea[areaName]) {
            itemsByArea[areaName] = [];
            orderedAreas.push(areaName);
        }
        itemsByArea[areaName].push(item);
    }
    for (const id of targetIds) {
        const workItem = workItemMap[id];
        if (!workItem) {
            continue;
        }
        const workItemType = String(workItem.fields['System.WorkItemType'] || '').toLowerCase();
        if (!workItemType.includes('epic') || (children[id] || []).length > 0) {
            continue;
        }
        const item = {
            id,
            title: workItem.fields['System.Title'] || '',
            risk: workItem.fields['OSG.RiskAssessment'] || '',
            riskComment: workItem.fields['OSG.RiskAssessmentComment'] || '',
            state: workItem.fields['System.State'] || '',
            assignedTo: workItem.fields['System.AssignedTo']?.displayName || ''
        };
        if (midpointMap) {
            item.midpointRisk = midpointMap[id]?.risk || '';
            item.midpointComment = midpointMap[id]?.riskComment || '';
        }
        const areaName = parseAreaLevel4(workItem.fields['System.AreaPath']) || 'Uncategorized';
        pushItem(areaName, item);
    }
    const areaNames = orderedAreas.length ? orderedAreas : [flatSlideTitle || 'Layout'];
    const titlePrefix = String(flatSlidePrefix || '').trim();
    return areaNames
        .map((areaName, index) => ({
        title: titlePrefix ? `${titlePrefix} - ${areaName}` : areaName,
        id: `flat-${String(areaName).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || index}`,
        items: itemsByArea[areaName] || []
    }))
        .filter((slide) => slide.items.length > 0);
}
async function getSlidesPreview(input, incomingBearerToken) {
    const { baseUrl, project, queryId } = (0, adoClient_1.parseQueryUrl)(input.queryUrl);
    const queryResult = await (0, adoClient_1.runTreeQuery)(baseUrl, project, queryId, incomingBearerToken);
    const allIds = new Set();
    const leafIds = new Set();
    const parentIds = new Set();
    for (const relation of queryResult.workItemRelations || []) {
        if (relation.target) {
            allIds.add(relation.target.id);
        }
        if (relation.source) {
            parentIds.add(relation.source.id);
        }
    }
    for (const id of allIds) {
        if (!parentIds.has(id)) {
            leafIds.add(id);
        }
    }
    const workItemMap = await (0, adoClient_1.fetchWorkItemsByIds)(baseUrl, project, [...allIds], incomingBearerToken);
    const midpointMap = input.midpointDate
        ? await batchHistoricalFields(baseUrl, project, [...leafIds], input.midpointDate, incomingBearerToken)
        : null;
    const titleContext = buildTitleContext(queryResult, workItemMap);
    const slides = input.flatView
        ? buildFlatSlides(queryResult, workItemMap, midpointMap, titleContext.flatSlideTitle, titleContext.flatSlidePrefix)
        : buildSlides(queryResult, workItemMap, midpointMap);
    return {
        slides,
        linkBase: `${baseUrl}/${project}/_workitems/edit/`,
        hasMidpoint: !!input.midpointDate,
        baseSlideTitle: titleContext.flatSlideTitle,
        baseTeamName: titleContext.baseTeamName,
        deckFileName: titleContext.deckFileName
    };
}
