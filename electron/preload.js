'use strict';

const { contextBridge } = require('electron');

// Expose a minimal surface to the renderer — no Node.js internals.
contextBridge.exposeInMainWorld('electronBridge', {
    platform: process.platform
});
