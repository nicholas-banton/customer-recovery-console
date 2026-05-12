const { contextBridge } = require('electron');

contextBridge.exposeInMainWorld('crcApp', {
  platform: process.platform,
  version: '0.1.0',
  runtime: 'electron'
});
