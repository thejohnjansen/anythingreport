import {
  fetchWorkItemsByIds,
  parseQueryUrl,
  runTreeQuery,
  type AzureDevOpsWorkItem,
  type WiqlTreeResult
} from './adoClient';

export interface DocItem {
  id: number;
  title: string;
  originalEstimate: number | string;
  dri: string;
  state: string;
}

export interface DocGroup {
  title: string;
  id: number;
  dri: string;
  originalEstimate: number | string;
  comments: string;
  items: DocItem[];
}

export interface DocSection {
  objectiveTitle: string;
  objectiveId: number;
  groups: DocGroup[];
}

export interface RetroMilestone {
  id: number;
  title: string;
  risk: string;
}

export interface RetroRow {
  id: number;
  rag: string;
  featureArea: string;
  mission: string;
  milestones: RetroMilestone[];
  justifications: string[];
}

function buildTree(queryResult: WiqlTreeResult): { roots: number[]; children: Record<number, number[]> } {
  const roots: number[] = [];
  const children: Record<number, number[]> = {};

  for (const relation of queryResult.workItemRelations || []) {
    if (!relation.target) {
      continue;
    }

    if (!relation.source) {
      roots.push(relation.target.id);
    } else {
      children[relation.source.id] ??= [];
      children[relation.source.id].push(relation.target.id);
    }
  }

  return { roots, children };
}

function parseAreaLevel4(areaPath: string | undefined): string {
  const parts = String(areaPath || '').split('\\').filter(Boolean);
  return parts[3] || '';
}

function riskAssessmentToRag(risk: string | undefined): string {
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

function buildDocSections(queryResult: WiqlTreeResult, workItemMap: Record<number, AzureDevOpsWorkItem>): DocSection[] {
  const { roots, children } = buildTree(queryResult);

  return roots.map((rootId) => {
    const rootWi = workItemMap[rootId];
    const groups: DocGroup[] = (children[rootId] || []).map((l2Id) => {
      const wi = workItemMap[l2Id];
      const items: DocItem[] = (children[l2Id] || [])
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

function buildRetroRows(queryResult: WiqlTreeResult, workItemMap: Record<number, AzureDevOpsWorkItem>): RetroRow[] {
  const { roots, children } = buildTree(queryResult);
  const rows: RetroRow[] = [];

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
        } else if (childRags.includes('yellow')) {
          rag = 'yellow';
        } else if (childRags.includes('green')) {
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

async function getQueryWorkItemMap(queryUrl: string, incomingBearerToken?: string): Promise<{
  baseUrl: string;
  project: string;
  queryResult: WiqlTreeResult;
  workItemMap: Record<number, AzureDevOpsWorkItem>;
}> {
  const { baseUrl, project, queryId } = parseQueryUrl(queryUrl);
  const queryResult = await runTreeQuery(baseUrl, project, queryId, incomingBearerToken);

  const allIds = new Set<number>();
  for (const relation of queryResult.workItemRelations || []) {
    if (relation.target) {
      allIds.add(relation.target.id);
    }
  }

  const workItemMap = await fetchWorkItemsByIds(baseUrl, project, [...allIds], incomingBearerToken);
  return { baseUrl, project, queryResult, workItemMap };
}

export async function getDocSections(
  queryUrl: string,
  incomingBearerToken?: string
): Promise<{ sections: DocSection[]; linkBase: string }> {
  const { baseUrl, project, queryResult, workItemMap } = await getQueryWorkItemMap(queryUrl, incomingBearerToken);
  return {
    sections: buildDocSections(queryResult, workItemMap),
    linkBase: `${baseUrl}/${project}/_workitems/edit/`
  };
}

export async function getRetroRows(
  queryUrl: string,
  incomingBearerToken?: string
): Promise<{ rows: RetroRow[]; linkBase: string }> {
  const { baseUrl, project, queryResult, workItemMap } = await getQueryWorkItemMap(queryUrl, incomingBearerToken);
  return {
    rows: buildRetroRows(queryResult, workItemMap),
    linkBase: `${baseUrl}/${project}/_workitems/edit/`
  };
}
