export interface StoryDocument {
  id: string;
  workItemId: number;
  cycle: string;
  title: string;
  area: string;
  pmComments: string;
  narrative: string;
  lastEditedBy: string;
  lastEditedAt: string;
  version: number;
  _etag?: string;
}

export interface StoryUpsertRequest {
  workItemId: number;
  cycle: string;
  title: string;
  area: string;
  pmComments?: string;
  narrative: string;
  lastEditedBy: string;
  expectedVersion?: number;
}

export interface StoryUpdateRequest {
  title: string;
  area: string;
  pmComments?: string;
  narrative: string;
  lastEditedBy: string;
}

export interface NormalizedWorkItem {
  workItemId: number;
  title: string;
  area: string;
}

export interface CosmosSettings {
  endpoint: string;
  key: string;
  databaseName: string;
  containerName: string;
}

export interface BoardDocument {
  id: string;           // "topofmind:{teamSlug}" | "pipeline:{teamSlug}"
  boardType: 'topofmind' | 'pipeline';
  teamSlug: string;
  cycle: string;        // always "__board__" — serves as partition-key value
  content: string;      // raw HTML for topofmind; JSON string for pipeline
  lastEditedBy: string;
  lastEditedAt: string;
  version: number;
  _etag?: string;
}

export interface BoardUpsertRequest {
  boardType: 'topofmind' | 'pipeline';
  teamSlug: string;
  content: string;
  lastEditedBy: string;
  expectedVersion?: number;
}