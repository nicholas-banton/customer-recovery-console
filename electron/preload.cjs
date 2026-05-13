const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('crcApp', {
  platform: process.platform,
  version: '0.5.0',
  runtime: 'electron',

  // Legacy invoke path retained as fallback.
  selectMboxFile: () => ipcRenderer.invoke('crc:select-mbox-file'),

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
