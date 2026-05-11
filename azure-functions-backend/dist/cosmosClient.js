"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getCosmosContainer = getCosmosContainer;
exports.buildStoryId = buildStoryId;
const cosmos_1 = require("@azure/cosmos");
let client = null;
let database = null;
let container = null;
function getSettingsFromEnv() {
    const endpoint = process.env.COSMOS_ENDPOINT;
    const key = process.env.COSMOS_KEY;
    const databaseName = process.env.COSMOS_DATABASE;
    const containerName = process.env.COSMOS_CONTAINER;
    if (!endpoint || !key || !databaseName || !containerName) {
        throw new Error('Missing Cosmos configuration. Set COSMOS_ENDPOINT, COSMOS_KEY, COSMOS_DATABASE, and COSMOS_CONTAINER.');
    }
    return { endpoint, key, databaseName, containerName };
}
function getCosmosContainer() {
    if (container) {
        return container;
    }
    const settings = getSettingsFromEnv();
    client = new cosmos_1.CosmosClient({ endpoint: settings.endpoint, key: settings.key });
    database = client.database(settings.databaseName);
    container = database.container(settings.containerName);
    return container;
}
function buildStoryId(workItemId, cycle) {
    return `${cycle}:${workItemId}`;
}
