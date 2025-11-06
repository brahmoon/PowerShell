const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  saveGraph: (payload) => ipcRenderer.invoke('graph:save', payload),
  loadGraph: () => ipcRenderer.invoke('graph:load'),
  saveScript: (payload) => ipcRenderer.invoke('script:save', payload),
  selectLocalFile: (payload) => ipcRenderer.invoke('dialog:selectFile', payload),
});
