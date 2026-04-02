'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('folioWindow', {
  close: () => ipcRenderer.send('window-close'),
  minimize: () => ipcRenderer.send('window-minimize'),
  maximize: () => ipcRenderer.send('window-maximize'),
  openOutputs: () => ipcRenderer.invoke('open-outputs'),
  onUpdateStatus: (cb) => ipcRenderer.on('update-status', (_, payload) => cb(payload))
});
