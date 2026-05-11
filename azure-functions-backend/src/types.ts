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