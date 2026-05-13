import type { HttpRequest, HttpResponseInit } from '@azure/functions';

function json(body: unknown, status: number): HttpResponseInit {
  return {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body, null, 2)
  };
}

function parseBearer(value: string | null): string | null {
  if (!value) {
    return null;
  }
  const match = value.match(/^Bearer\s+(.+)$/i);
  return match?.[1]?.trim() || null;
}

type ClientPrincipalClaim = { typ?: string; val?: string };
type ClientPrincipal = { userDetails?: string; claims?: ClientPrincipalClaim[] };

function readClientPrincipal(request: HttpRequest): ClientPrincipal | null {
  const encoded = request.headers.get('x-ms-client-principal');
  if (!encoded) {
    return null;
  }

  try {
    const decoded = Buffer.from(encoded, 'base64').toString('utf8');
    return JSON.parse(decoded) as ClientPrincipal;
  } catch {
    return null;
  }
}

function claimValue(principal: ClientPrincipal | null, claimType: string): string {
  if (!principal?.claims?.length) {
    return '';
  }
  const claim = principal.claims.find((c) => String(c.typ || '').toLowerCase() === claimType.toLowerCase());
  return String(claim?.val || '').trim();
}

function extractUserIdentifier(principal: ClientPrincipal | null): string {
  return (
    String(principal?.userDetails || '').trim() ||
    claimValue(principal, 'preferred_username') ||
    claimValue(principal, 'upn') ||
    claimValue(principal, 'email')
  );
}

export function getIncomingAdoToken(request: HttpRequest): string | undefined {
  const fromAuthHeader = parseBearer(request.headers.get('authorization'));
  if (fromAuthHeader) {
    return fromAuthHeader;
  }

  const fromEasyAuth = request.headers.get('x-ms-token-aad-access-token');
  if (fromEasyAuth && fromEasyAuth.trim()) {
    return fromEasyAuth.trim();
  }

  return undefined;
}

export function requireMicrosoftUser(request: HttpRequest): HttpResponseInit | null {
  // Local development can run without Static Web Apps identity headers.
  const principal = readClientPrincipal(request);
  if (!principal) {
    return null;
  }

  const user = extractUserIdentifier(principal).toLowerCase();
  if (!user) {
    return json({ error: 'Unable to determine signed-in user identity.' }, 401);
  }

  if (!user.endsWith('@microsoft.com')) {
    return json({ error: 'Access denied. Only @microsoft.com accounts are allowed.' }, 403);
  }

  return null;
}