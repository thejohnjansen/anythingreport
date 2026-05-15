'use strict';

/**
 * Token server — a localhost-only HTTP server that the Express child process
 * calls to obtain a fresh access token without depending on az CLI.
 *
 * Usage (in main process):
 *   const { startTokenServer } = require('./tokenServer');
 *   const { port, close } = await startTokenServer(getAccessToken);
 *   // Pass port to the server process as MSAL_TOKEN_PORT env var.
 *
 * Usage (in server.js):
 *   const token = await fetch(`http://127.0.0.1:${process.env.MSAL_TOKEN_PORT}/token`)
 *                          .then(r => r.json()).then(d => d.token);
 */

const http = require('http');

/**
 * @param {() => Promise<string>} getToken  Async function that returns an access token.
 * @returns {Promise<{ port: number, close: () => void }>}
 */
function startTokenServer(getToken) {
    // Cache the token in-process so we only call getToken() (which may trigger
    // an interactive browser flow) once. Expire 5 minutes before the actual
    // token expiry to give some headroom for clock skew.
    const CACHE_MS = 55 * 60 * 1000; // 55 minutes
    let cache = { token: null, expiry: 0 };

    async function getCachedToken() {
        if (cache.token && Date.now() < cache.expiry) return cache.token;
        const token = await getToken();
        cache = { token, expiry: Date.now() + CACHE_MS };
        return token;
    }

    return new Promise((resolve, reject) => {
        const server = http.createServer(async (req, res) => {
            // Only serve the /token route
            if (req.method !== 'GET' || req.url !== '/token') {
                res.writeHead(404);
                res.end();
                return;
            }

            // Reject requests that are not from localhost
            const remote = req.socket.remoteAddress;
            const isLocal = remote === '127.0.0.1'
                         || remote === '::1'
                         || remote === '::ffff:127.0.0.1';
            if (!isLocal) {
                res.writeHead(403);
                res.end();
                return;
            }

            try {
                const token = await getCachedToken();
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ token }));
            } catch (err) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: err.message }));
            }
        });

        // Bind to 127.0.0.1 (not 0.0.0.0) so it is never reachable from outside
        server.listen(0, '127.0.0.1', () => {
            const { port } = server.address();
            resolve({
                port,
                close: () => server.close()
            });
        });

        server.on('error', reject);
    });
}

module.exports = { startTokenServer };
