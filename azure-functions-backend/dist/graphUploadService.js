"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.uploadPptToGraph = uploadPptToGraph;
function normalizeFolderPath(folderPath) {
    return folderPath
        .replace(/\\/g, '/')
        .replace(/^\/+/, '')
        .replace(/\/+$/, '');
}
function encodePathSegment(value) {
    return encodeURIComponent(value).replace(/%2F/g, '/');
}
function getUploadUrl(target, fileName) {
    const normalizedFolder = normalizeFolderPath(target.folderPath);
    const pathWithFile = normalizedFolder ? `${normalizedFolder}/${fileName}` : fileName;
    const encodedPath = encodePathSegment(pathWithFile);
    if (target.mode === 'application') {
        return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(target.driveId)}/root:/${encodedPath}:/content`;
    }
    if (target.userId) {
        return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(target.userId)}/drive/root:/${encodedPath}:/content`;
    }
    return `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`;
}
async function uploadPptToGraph(input) {
    const uploadUrl = getUploadUrl(input.target, input.fileName);
    const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
            Authorization: `Bearer ${input.accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        },
        body: new Uint8Array(input.fileBuffer)
    });
    const bodyText = await response.text();
    if (!response.ok) {
        const snippet = bodyText.replace(/\s+/g, ' ').trim().slice(0, 500);
        throw new Error(`Graph upload failed (${response.status} ${response.statusText})${snippet ? ` - ${snippet}` : ''}`);
    }
    const payload = JSON.parse(bodyText);
    if (!payload.webUrl) {
        throw new Error('Graph upload succeeded but response did not include webUrl.');
    }
    return payload.webUrl;
}
