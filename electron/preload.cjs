const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('crcApp', {
  platform: process.platform,
  version: '1.0.0',
  runtime: 'electron',

  // Legacy invoke path retained as fallback.
  selectMboxFile: () => ipcRenderer.invoke('crc:select-mbox-file'),
  prepareResultsPackage: (payload) => ipcRenderer.invoke('crc:prepare-results-package', payload),
  openResultsFolder: (folderPath) => ipcRenderer.invoke('crc:open-results-folder', folderPath),
  revealPath: (targetPath) => ipcRenderer.invoke('crc:reveal-path', targetPath),
  copyFullExportFile: (payload) => ipcRenderer.invoke('crc:copy-full-export-file', payload),

  // Event-driven long-running import path.
  startMboxImport: () => ipcRenderer.send('crc:start-mbox-import'),

  onMboxImportProgress: (callback) => {
    if (typeof callback !== 'function') {
      return () => {};
    }

    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on('crc:mbox-import-progress', listener);

    return () => {
      ipcRenderer.removeListener('crc:mbox-import-progress', listener);
    };
  },

  onMboxImportComplete: (callback) => {
    if (typeof callback !== 'function') {
      return () => {};
    }

    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on('crc:mbox-import-complete', listener);

    return () => {
      ipcRenderer.removeListener('crc:mbox-import-complete', listener);
    };
  }
});
