const { app, BrowserWindow, dialog, ipcMain, protocol } = require('electron');
const path = require('path');
const fs = require('fs/promises');

const GRAPH_FILTER = [{ name: 'NodeFlow Graph', extensions: ['json'] }];
const SCRIPT_FILTER = [{ name: 'PowerShell Script', extensions: ['ps1', 'txt'] }];
const APP_PROTOCOL = 'app';

protocol.registerSchemesAsPrivileged?.([
  {
    scheme: APP_PROTOCOL,
    privileges: {
      secure: true,
      standard: true,
      supportFetchAPI: true,
      corsEnabled: true,
      stream: true,
    },
  },
]);

const registerLocalProtocol = () => {
  protocol.registerFileProtocol(APP_PROTOCOL, (request, callback) => {
    const url = request.url.replace(`${APP_PROTOCOL}://`, '');
    const relativePath = url.length ? url : 'index.html';
    const filePath = path.normalize(path.join(__dirname, relativePath));
    callback({ path: filePath });
  });
};

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

  mainWindow.loadURL(`${APP_PROTOCOL}://index.html`);
};

app.whenReady().then(() => {
  registerLocalProtocol();
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

ipcMain.handle('dialog:selectFile', async (_event, payload = {}) => {
  const defaultPath =
    typeof payload?.defaultPath === 'string' && payload.defaultPath.trim()
      ? payload.defaultPath
      : app.getPath('desktop');
  const mode = payload?.mode === 'directory' ? 'directory' : 'file';
  const properties = [];
  if (mode === 'directory') {
    properties.push('openDirectory');
  } else {
    properties.push('openFile');
  }
  const options = {
    defaultPath,
    properties,
  };
  if (Array.isArray(payload?.filters) && payload.filters.length) {
    options.filters = payload.filters;
  }
  const { canceled, filePaths } = await dialog.showOpenDialog(options);
  if (canceled || !filePaths?.length) {
    return { canceled: true };
  }
  const [filePath] = filePaths;
  return { canceled: false, filePath, fileName: path.basename(filePath) };
});
