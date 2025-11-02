const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs/promises');

const GRAPH_FILTER = [{ name: 'NodeFlow Graph', extensions: ['json'] }];
const SCRIPT_FILTER = [{ name: 'PowerShell Script', extensions: ['ps1', 'txt'] }];

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 960,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
};

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('graph:save', async (_event, payload = {}) => {
  const { content = '', suggestedName = 'nodeflow-graph.json' } = payload;
  const { canceled, filePath } = await dialog.showSaveDialog({
    defaultPath: suggestedName,
    filters: GRAPH_FILTER,
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  await fs.writeFile(filePath, content, 'utf-8');
  return { canceled: false, fileName: path.basename(filePath), fullPath: filePath };
});

ipcMain.handle('graph:load', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    filters: GRAPH_FILTER,
    properties: ['openFile'],
  });

  if (canceled || !filePaths?.length) {
    return { canceled: true };
  }

  const [filePath] = filePaths;
  const content = await fs.readFile(filePath, 'utf-8');
  return { canceled: false, fileName: path.basename(filePath), content };
});

ipcMain.handle('script:save', async (_event, payload = {}) => {
  const { content = '', suggestedName = 'flow.ps1' } = payload;
  const { canceled, filePath } = await dialog.showSaveDialog({
    defaultPath: suggestedName,
    filters: SCRIPT_FILTER,
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  await fs.writeFile(filePath, content, 'utf-8');
  return { canceled: false, fileName: path.basename(filePath), fullPath: filePath };
});
