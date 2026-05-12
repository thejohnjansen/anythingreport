"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getCosmosContainer = getCosmosContainer;
exports.buildStoryId = buildStoryId;
const cosmos_1 = require("@azure/cosmos");
let client = null;
let database = null;
let container = null;
function parseConnectionString(connectionString) {
    const pairs = connectionString
        .split(';')
        .map((part) => part.trim())
        .filter(Boolean)
        .map((part) => {
        const separatorIndex = part.indexOf('=');
        if (separatorIndex < 0) {
            return null;
        }
        const name = part.slice(0, separatorIndex).trim().toLowerCase();
        const value = part.slice(separatorIndex + 1).trim();
        return [name, value];
    })
        .filter((entry) => entry !== null);
    const values = new Map(pairs);
    const endpoint = values.get('accountendpoint');
    const key = values.get('accountkey');
    if (!endpoint || !key) {
        throw new Error('Invalid COSMOS_CONNECTION_STRING. Expected AccountEndpoint and AccountKey entries.');
    }
    return { endpoint, key };
}
function getSettingsFromEnv() {
    const connectionString = process.env.COSMOS_CONNECTION_STRING;
    const databaseName = process.env.COSMOS_DATABASE;
    const containerName = process.env.COSMOS_CONTAINER;
    if (!connectionString || !databaseName || !containerName) {
        throw new Error('Missing Cosmos configuration. Set COSMOS_CONNECTION_STRING, COSMOS_DATABASE, and COSMOS_CONTAINER.');
    }
    const { endpoint, key } = parseConnectionString(connectionString);
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
