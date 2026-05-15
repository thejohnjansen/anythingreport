'use strict';

/**
 * MSAL authentication for Anything Report (Electron).
 *
 * Prerequisites — add to .env:
 *   AZURE_CLIENT_ID=<your app registration client ID>
 *   AZURE_TENANT_ID=<your tenant ID, or "organizations" for any Microsoft org account>
 *
 * App registration settings (Azure Portal → Entra → App registrations):
 *   - Platform: Mobile and desktop application
 *   - Redirect URI: http://localhost  (MSAL picks the port dynamically)
 *   - API permission: Azure DevOps → user_impersonation (delegated)
 *   - Public client: Yes  (no client secret needed)
 */

const { PublicClientApplication, InteractionRequiredAuthError } = require('@azure/msal-node');
const { app, shell } = require('electron');
const path = require('path');
const fs   = require('fs');

// ADO resource scope
const SCOPES = ['https://app.vssps.visualstudio.com/user_impersonation'];

// Persist the MSAL token cache to disk so users stay logged in across launches
const CACHE_FILE = path.join(app.getPath('userData'), 'msal-cache.json');

const cachePlugin = {
    beforeCacheAccess: async (ctx) => {
        try {
            const raw = fs.existsSync(CACHE_FILE) ? fs.readFileSync(CACHE_FILE, 'utf8') : '';
            ctx.tokenCache.deserialize(raw);
        } catch { /* ignore — start with empty cache */ }
    },
    afterCacheAccess: async (ctx) => {
        if (ctx.cacheHasChanged) {
            try {
                fs.mkdirSync(path.dirname(CACHE_FILE), { recursive: true });
                fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize(), 'utf8');
            } catch { /* non-fatal — token will just be re-acquired next launch */ }
        }
    }
};

let _pca = null;

function getClient() {
    if (_pca) return _pca;

    const clientId = process.env.AZURE_CLIENT_ID;
    const tenantId = process.env.AZURE_TENANT_ID || 'organizations';

    if (!clientId) {
        throw new Error(
            'AZURE_CLIENT_ID is not set.\n' +
            'Add it to your .env file (see electron/msalAuth.js for setup instructions).'
        );
    }

    _pca = new PublicClientApplication({
        auth: {
            clientId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        },
        cache: { cachePlugin }
    });

    return _pca;
}

/**
 * Returns a valid ADO access token.
 * Tries the silent (cached / refresh-token) path first.
 * Falls back to an interactive browser sign-in if needed.
 */
async function getAccessToken() {
    const client = getClient();
    const accounts = await client.getAllAccounts();

    if (accounts.length > 0) {
        try {
            const result = await client.acquireTokenSilent({
                scopes:  SCOPES,
                account: accounts[0]
            });
            return result.accessToken;
        } catch (e) {
            // Only fall through to interactive if the error is "interaction required"
            if (!(e instanceof InteractionRequiredAuthError)) throw e;
        }
    }

    // Interactive: MSAL starts a loopback server, calls openBrowser with the auth URL,
    // then waits for the redirect to capture the authorization code automatically.
    const result = await client.acquireTokenInteractive({
        scopes: SCOPES,
        openBrowser: async (url) => { await shell.openExternal(url); },
        successTemplate: `<!DOCTYPE html>
<html>
<body style="margin:0;display:flex;align-items:center;justify-content:center;
             height:100vh;font-family:'Segoe UI',sans-serif;
             background:#0f172a;color:#e2e8f0;text-align:center">
  <div>
    <div style="font-size:3rem;margin-bottom:.75rem">&#x2713;</div>
    <h2 style="margin:0 0 .5rem">Signed in successfully</h2>
    <p style="color:#94a3b8">You can close this tab and return to Anything Report.</p>
  </div>
</body>
</html>`,
        errorTemplate: `<!DOCTYPE html>
<html>
<body style="margin:0;display:flex;align-items:center;justify-content:center;
             height:100vh;font-family:'Segoe UI',sans-serif;
             background:#0f172a;color:#e2e8f0;text-align:center">
  <div>
    <div style="font-size:3rem;margin-bottom:.75rem">&#x2717;</div>
    <h2 style="margin:0 0 .5rem">Sign-in failed</h2>
    <p style="color:#94a3b8">{errorMessage}</p>
  </div>
</body>
</html>`
    });

    return result.accessToken;
}

/**
 * Returns true if there is at least one cached account.
 * A cached account does NOT guarantee the token is still valid —
 * call getAccessToken() for that (it will refresh silently if needed).
 */
async function hasCachedAccount() {
    try {
        const accounts = await getClient().getAllAccounts();
        return accounts.length > 0;
    } catch {
        return false;
    }
}

module.exports = { getAccessToken, hasCachedAccount };
