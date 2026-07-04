const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  platform: process.platform,
  ping: () => ipcRenderer.invoke('ping'),
  extractPdfText: (fileName, buffer) => ipcRenderer.invoke('extract-pdf-text', { fileName, buffer }),
  parseXlsx: (fileName, buffer) => ipcRenderer.invoke('parse-xlsx', { fileName, buffer }),
  saveFile: (defaultName, buffer, mimeType) => ipcRenderer.invoke('save-file', { defaultName, buffer, mimeType }),
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  onImportFiles: (callback) => ipcRenderer.on('import-files', (event, filePaths) => callback(filePaths)),
  removeImportListener: () => ipcRenderer.removeAllListeners('import-files'),
});
