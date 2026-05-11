import type { Container } from '@azure/cosmos';
import { buildStoryId, getCosmosContainer } from './cosmosClient';
import type { StoryDocument, StoryUpdateRequest, StoryUpsertRequest } from './types';

export class StoryRepository {
  constructor(private readonly container: Container = getCosmosContainer()) {}

  async getStoriesByCycle(cycle: string): Promise<StoryDocument[]> {
    const query = {
      query: 'SELECT * FROM c WHERE c.cycle = @cycle ORDER BY c.lastEditedAt DESC',
      parameters: [{ name: '@cycle', value: cycle }]
    };

    const { resources } = await this.container.items.query<StoryDocument>(query).fetchAll();
    return resources;
  }

  async upsertStory(story: StoryUpsertRequest): Promise<StoryDocument> {
    const id = buildStoryId(story.workItemId, story.cycle);
    const existing = await this.readStoryById(id);

    const now = new Date().toISOString();
    const document: StoryDocument = {
      id,
      workItemId: story.workItemId,
      cycle: story.cycle,
      title: story.title,
      area: story.area,
      pmComments: story.pmComments ?? '',
      narrative: story.narrative,
      lastEditedBy: story.lastEditedBy,
      lastEditedAt: now,
      version: (existing?.version ?? 0) + 1
    };

    const response = await this.container.items.upsert<StoryDocument>(document);
    if (!response.resource) {
      throw new Error('Cosmos upsert returned no resource.');
    }
    return response.resource;
  }

  async updateStoryWithConcurrencyCheck(
    id: string,
    version: number,
    update: StoryUpdateRequest
  ): Promise<StoryDocument> {
    const existing = await this.readStoryById(id);
    if (!existing) {
      throw new Error(`Story ${id} not found`);
    }

    if (version !== existing.version) {
      throw new Error(`Version conflict for story ${id}. Expected ${version}, found ${existing.version}.`);
    }
    if (!existing._etag) {
      throw new Error(`Missing ETag for story ${id}; cannot perform concurrency-safe update.`);
    }

    const updated: StoryDocument = {
      ...existing,
      title: update.title,
      area: update.area,
      pmComments: update.pmComments ?? existing.pmComments,
      narrative: update.narrative,
      lastEditedBy: update.lastEditedBy,
      lastEditedAt: new Date().toISOString(),
      version: existing.version + 1
    };

    const response = await this.container.item(id, existing.cycle).replace(updated, {
      accessCondition: {
        type: 'IfMatch',
        condition: existing._etag
      }
    });
    if (!response.resource) {
      throw new Error('Cosmos replace returned no resource.');
    }
    return response.resource as unknown as StoryDocument;
  }

  private async readStoryById(id: string): Promise<StoryDocument | null> {
    const query = {
      query: 'SELECT TOP 1 * FROM c WHERE c.id = @id',
      parameters: [{ name: '@id', value: id }]
    };

    const { resources } = await this.container.items.query<StoryDocument>(query).fetchAll();
    return resources[0] ?? null;
  }

  async listByCycle(cycle: string): Promise<StoryDocument[]> {
    return this.getStoriesByCycle(cycle);
  }

  async createOrUpdate(story: StoryUpsertRequest): Promise<StoryDocument> {
    return this.upsertStory(story);
  }
}
