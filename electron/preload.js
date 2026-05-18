'use strict';

const { contextBridge, ipcRenderer } = require('electron');

// Expose a minimal surface to the renderer — no Node.js internals.
contextBridge.exposeInMainWorld('electronBridge', {
    platform: process.platform,
    savePptx: (buffer, defaultFileName) =>
        ipcRenderer.invoke('save-pptx', { buffer: Array.from(new Uint8Array(buffer)), defaultFileName })
});
