'use strict';

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const { app, BrowserWindow, shell, Menu } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const http = require('http');
const { getAccessToken, hasCachedAccount, signOut } = require('./msalAuth');
const { startTokenServer } = require('./tokenServer');

const SERVER_PORT = process.env.PORT || 3456;
let serverProcess = null;
let mainWindow = null;

/* ── Start Express server as child process ── */
function startServer(msalTokenPort) {
    const serverPath = path.join(__dirname, '..', 'server.js');
    const env = { ...process.env };
    if (msalTokenPort) env.MSAL_TOKEN_PORT = String(msalTokenPort);
    serverProcess = spawn(process.execPath, [serverPath], {
        cwd: path.join(__dirname, '..'),
        env,
        stdio: ['ignore', 'pipe', 'pipe']
    });

    serverProcess.stdout.on('data', (data) => {
        process.stdout.write('[server] ' + data.toString());
    });
    serverProcess.stderr.on('data', (data) => {
        process.stderr.write('[server:err] ' + data.toString());
    });
    serverProcess.on('exit', (code) => {
        console.log(`[server] process exited (code ${code})`);
    });
}

/* ── Poll until Express is accepting connections ── */
function waitForServer(maxAttempts = 40) {
    return new Promise((resolve, reject) => {
        let attempts = 0;

        function attempt() {
            const req = http.get(`http://localhost:${SERVER_PORT}/api/ping`, (res) => {
                res.resume();
                resolve();
            });
            req.setTimeout(800, () => req.destroy());
            req.on('error', () => {
                attempts++;
                if (attempts >= maxAttempts) {
                    reject(new Error('Express server did not become ready in time.'));
                } else {
                    setTimeout(attempt, 500);
                }
            });
        }

        attempt();
    });
}

/* ── Splash / status window (single instance, content updated in-place) ── */
function createSplashWindow() {
    const win = new BrowserWindow({
        width: 520,
        height: 260,
        frame: false,
        resizable: false,
        center: true,
        webPreferences: { nodeIntegration: false, contextIsolation: false }
    });

    const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: .75rem;
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: #0f172a;
    color: #e2e8f0;
    padding: 2rem;
    text-align: center;
  }
  #app-name { font-size: 1.1rem; font-weight: 600; color: #94a3b8; letter-spacing: .05em; }
  #icon     { font-size: 2.2rem; line-height: 1; }
  #message  { font-size: 1rem; font-weight: 500; }
  #detail   { font-size: .8rem; color: #94a3b8; max-width: 400px; line-height: 1.5; }
  a         { color: #60a5fa; }
  code      { font-family: 'Cascadia Code', Consolas, monospace; background: #1e293b;
              padding: .1em .35em; border-radius: 4px; }
</style>
</head>
<body>
  <div id="app-name">ANYTHING REPORT</div>
  <div id="icon">&#x23F3;</div>
  <div id="message">Starting&hellip;</div>
  <div id="detail"></div>
  <script>
    function update(icon, message, detail) {
      document.getElementById('icon').innerHTML    = icon;
      document.getElementById('message').innerHTML = message;
      document.getElementById('detail').innerHTML  = detail || '';
    }
  </script>
</body>
</html>`;

    win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    return win;
}

/* ── Update splash content ── */
function updateSplash(win, icon, message, detail = '') {
    return win.webContents.executeJavaScript(
        `update(${JSON.stringify(icon)}, ${JSON.stringify(message)}, ${JSON.stringify(detail)})`
    );
}

/* ── Create the main window ── */
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1440,
        height: 900,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            sandbox: false
        }
    });

    mainWindow.loadURL(`http://localhost:${SERVER_PORT}`);

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        if (url.startsWith('http://') || url.startsWith('https://')) {
            shell.openExternal(url);
        }
        return { action: 'deny' };
    });

    mainWindow.webContents.on('will-navigate', (event, url) => {
        const localBase = `http://localhost:${SERVER_PORT}`;
        if (!url.startsWith(localBase)) {
            event.preventDefault();
            shell.openExternal(url);
        }
    });

    mainWindow.on('closed', () => { mainWindow = null; });
}

/* ── App lifecycle ── */
app.whenReady().then(async () => {
    // Show splash immediately so users see something right away
    const splash = createSplashWindow();
    await new Promise((resolve) => splash.webContents.once('did-finish-load', resolve));

    const useMsal = !!process.env.AZURE_CLIENT_ID;

    // Start the MSAL token server before the Express server so the port is
    // available as an env var when the Express process starts.
    let tokenServerHandle;
    if (useMsal) {
        try {
            await updateSplash(splash, '&#x1F510;', 'Initialising authentication&hellip;');
            tokenServerHandle = await startTokenServer(getAccessToken);
        } catch (err) {
            await updateSplash(splash, '&#x274C;', 'Authentication setup failed.',
                err.message.replace(/</g, '&lt;'));
            setTimeout(() => app.quit(), 8000);
            return;
        }
    }

    await updateSplash(splash, '&#x23F3;', 'Starting server&hellip;');
    startServer(tokenServerHandle ? tokenServerHandle.port : null);

    try {
        await waitForServer();
    } catch (err) {
        await updateSplash(splash, '&#x274C;', 'Could not start the server.',
            'Please try restarting the app. If the problem persists, reinstall.');
        setTimeout(() => app.quit(), 8000);
        return;
    }

    if (useMsal) {
        // If there is no cached account, we know interactive sign-in is needed.
        // Otherwise getAccessToken() will silently refresh in the background.
        const cached = await hasCachedAccount();
        if (!cached) {
            await updateSplash(splash, '&#x1F511;', 'Sign in to your Microsoft account',
                'A browser window has opened &mdash; complete sign-in there, then return here.');
        } else {
            await updateSplash(splash, '&#x1F510;', 'Verifying sign-in&hellip;');
        }

        try {
            // Warm up the token now — call the token server so its in-memory
            // cache is pre-populated. This triggers interactive sign-in if needed,
            // or silently refreshes an expired token from the cache.
            // All subsequent calls from server.js will hit the warm cache instantly.
            const warmupRes = await fetch(`http://127.0.0.1:${tokenServerHandle.port}/token`);
            const warmupData = await warmupRes.json();
            if (!warmupRes.ok) throw new Error(warmupData.error || 'Token warmup failed');
        } catch (err) {
            await updateSplash(splash, '&#x274C;', 'Sign-in failed.',
                err.message.replace(/</g, '&lt;'));
            setTimeout(() => app.quit(), 8000);
            return;
        }
    }

    splash.close();
    createWindow();

    if (useMsal) {
        Menu.setApplicationMenu(Menu.buildFromTemplate([
            {
                label: 'Account',
                submenu: [
                    {
                        label: 'Sign Out',
                        click: async () => {
                            await signOut();
                            app.relaunch();
                            app.quit();
                        }
                    }
                ]
            }
        ]));
    }
});

app.on('window-all-closed', () => {
    if (serverProcess) {
        serverProcess.kill();
        serverProcess = null;
    }
    app.quit();
});

app.on('activate', () => {
    if (mainWindow === null) createWindow();
});
