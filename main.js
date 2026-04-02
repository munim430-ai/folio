'use strict';

const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const os = require('os');

let mainWindow = null;
let serverPort = null;

// Start Express server
function startServer() {
  return new Promise((resolve, reject) => {
    const server = require('./server.js');
    server.start((port) => {
      serverPort = port;
      resolve(port);
    });
  });
}

function createWindow(port) {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 780,
    minWidth: 900,
    minHeight: 600,
    frame: false,
    show: false,
    backgroundColor: '#080808',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.loadURL(`http://127.0.0.1:${port}/auth`);

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// IPC handlers
ipcMain.on('window-close', () => {
  if (mainWindow) mainWindow.close();
});

ipcMain.on('window-minimize', () => {
  if (mainWindow) mainWindow.minimize();
});

ipcMain.on('window-maximize', () => {
  if (mainWindow) {
    if (mainWindow.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow.maximize();
    }
  }
});

ipcMain.handle('open-outputs', async () => {
  const outputDir = path.join(os.homedir(), 'Documents', 'Mokwon University');
  const { mkdirSync } = require('fs');
  try { mkdirSync(outputDir, { recursive: true }); } catch (_) {}
  shell.openPath(outputDir);
  return { ok: true };
});

ipcMain.handle('open-templates', async () => {
  let templatesDir;
  if (app.isPackaged) {
    templatesDir = path.join(process.resourcesPath, 'templates');
  } else {
    templatesDir = path.join(__dirname, 'templates');
  }
  shell.openPath(templatesDir);
  return { ok: true };
});

app.whenReady().then(async () => {
  try {
    const port = await startServer();
    createWindow(port);
  } catch (err) {
    console.error('Failed to start server:', err);
    app.quit();
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0 && serverPort) {
      createWindow(serverPort);
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
