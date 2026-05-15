'use strict';

const { app, BrowserWindow, shell } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const http = require('http');

const SERVER_PORT = process.env.PORT || 3456;
let serverProcess = null;
let mainWindow = null;

/* ── Start Express server as child process ── */
function startServer() {
    const serverPath = path.join(__dirname, '..', 'server.js');
    serverProcess = spawn(process.execPath, [serverPath], {
        cwd: path.join(__dirname, '..'),
        env: { ...process.env },
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

/* ── Tiny frameless status window ── */
function createStatusWindow(message) {
    const win = new BrowserWindow({
        width: 460,
        height: 140,
        frame: false,
        resizable: false,
        center: true,
        alwaysOnTop: true,
        webPreferences: { nodeIntegration: false, contextIsolation: true }
    });

    const html = `<!DOCTYPE html><html><body style="margin:0;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;font-family:'Segoe UI',system-ui,sans-serif;background:#0f172a;color:#e2e8f0;font-size:14px;gap:.6rem;"><div style="font-size:1.5rem;">&#128274;</div><div>${message}</div></body></html>`;
    win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    return win;
}

/* ── Check Azure CLI auth via the server endpoint ── */
async function checkAuth() {
    try {
        const res = await fetch(`http://localhost:${SERVER_PORT}/api/auth/status`);
        const data = await res.json();
        return data.authenticated === true;
    } catch {
        return false;
    }
}

/* ── Run az login (opens the default browser) ── */
function runAzLogin() {
    return new Promise((resolve, reject) => {
        const proc = spawn('az', ['login'], {
            shell: true,
            stdio: 'pipe',       // suppress console noise in the packaged app
            windowsHide: true    // no extra CMD window on Windows
        });
        proc.on('exit', (code) => {
            if (code === 0) resolve();
            else reject(new Error(`az login exited with code ${code}`));
        });
        proc.on('error', (err) => reject(err));
    });
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

    // Open external URLs (e.g. ADO links) in the default browser
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
    startServer();

    try {
        await waitForServer();
    } catch (err) {
        console.error('Fatal: could not start server —', err.message);
        app.quit();
        return;
    }

    const authenticated = await checkAuth();

    if (!authenticated) {
        const statusWin = createStatusWindow('Opening Azure sign-in in your browser&hellip;');

        try {
            await runAzLogin();
        } catch (err) {
            statusWin.close();
            const errWin = createStatusWindow('Sign-in failed. Run <code>az login</code> and restart.');
            setTimeout(() => app.quit(), 6000);
            return;
        }

        // Verify login succeeded
        const nowAuthenticated = await checkAuth();
        statusWin.close();

        if (!nowAuthenticated) {
            const errWin = createStatusWindow('Sign-in did not complete. Run <code>az login</code> and restart.');
            setTimeout(() => app.quit(), 6000);
            return;
        }
    }

    createWindow();
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
