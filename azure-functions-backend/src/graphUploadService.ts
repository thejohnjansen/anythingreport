export type GraphUploadTarget =
  | {
      mode: 'delegated';
      folderPath: string;
      userId?: string;
    }
  | {
      mode: 'application';
      driveId: string;
      folderPath: string;
    };

export interface GraphUploadInput {
  accessToken: string;
  fileName: string;
  fileBuffer: Buffer;
  target: GraphUploadTarget;
}

interface GraphUploadResponse {
  webUrl?: string;
}

function normalizeFolderPath(folderPath: string): string {
  return folderPath
    .replace(/\\/g, '/')
    .replace(/^\/+/, '')
    .replace(/\/+$/, '');
}

function encodePathSegment(value: string): string {
  return encodeURIComponent(value).replace(/%2F/g, '/');
}

function getUploadUrl(target: GraphUploadTarget, fileName: string): string {
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

export async function uploadPptToGraph(input: GraphUploadInput): Promise<string> {
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

  const payload = JSON.parse(bodyText) as GraphUploadResponse;
  if (!payload.webUrl) {
    throw new Error('Graph upload succeeded but response did not include webUrl.');
  }

  return payload.webUrl;
}
