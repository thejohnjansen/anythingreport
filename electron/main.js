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
            const req = http.get(`http://localhost:${SERVER_PORT}/api/config`, (res) => {
                res.resume(); // drain the response
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

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

/* ── App lifecycle ── */
app.whenReady().then(async () => {
    startServer();
    try {
        await waitForServer();
        createWindow();
    } catch (err) {
        console.error('Fatal: could not start server —', err.message);
        app.quit();
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
    if (mainWindow === null) {
        createWindow();
    }
});
