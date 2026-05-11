"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.StoryRepository = void 0;
const cosmosClient_1 = require("./cosmosClient");
class StoryRepository {
    container;
    constructor(container = (0, cosmosClient_1.getCosmosContainer)()) {
        this.container = container;
    }
    async getStoriesByCycle(cycle) {
        const query = {
            query: 'SELECT * FROM c WHERE c.cycle = @cycle ORDER BY c.lastEditedAt DESC',
            parameters: [{ name: '@cycle', value: cycle }]
        };
        const { resources } = await this.container.items.query(query).fetchAll();
        return resources;
    }
    async upsertStory(story) {
        const id = (0, cosmosClient_1.buildStoryId)(story.workItemId, story.cycle);
        const existing = await this.readStoryById(id);
        const now = new Date().toISOString();
        const document = {
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
        const response = await this.container.items.upsert(document);
        if (!response.resource) {
            throw new Error('Cosmos upsert returned no resource.');
        }
        return response.resource;
    }
    async updateStoryWithConcurrencyCheck(id, version, update) {
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
        const updated = {
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
        return response.resource;
    }
    async readStoryById(id) {
        const query = {
            query: 'SELECT TOP 1 * FROM c WHERE c.id = @id',
            parameters: [{ name: '@id', value: id }]
        };
        const { resources } = await this.container.items.query(query).fetchAll();
        return resources[0] ?? null;
    }
    async listByCycle(cycle) {
        return this.getStoriesByCycle(cycle);
    }
    async createOrUpdate(story) {
        return this.upsertStory(story);
    }
}
exports.StoryRepository = StoryRepository;
