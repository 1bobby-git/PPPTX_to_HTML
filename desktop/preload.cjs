'use strict';

const {contextBridge, ipcRenderer} = require('electron');
const invoke = name => (...args) => ipcRenderer.invoke('pptx-desktop:' + name, ...args);
contextBridge.exposeInMainWorld('pptxDesktop', Object.freeze({
  getInfo: invoke('info'),
  chooseFile: invoke('choose'),
  convert: invoke('convert'),
  cancel: invoke('cancel'),
  preview: invoke('preview'),
  saveHtml: invoke('save'),
  onProgress(callback) {
    if (typeof callback !== 'function') throw new TypeError('A callback is required.');
    const listener = (_event, message) => callback(message);
    ipcRenderer.on('pptx-desktop:progress', listener);
    return () => ipcRenderer.removeListener('pptx-desktop:progress', listener);
  }
}));
