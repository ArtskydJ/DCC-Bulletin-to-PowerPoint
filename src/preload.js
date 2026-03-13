'use strict';

const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('dcc', {
  // Settings
  getSettings: () => ipcRenderer.invoke('get-settings'),
  saveSettings: (s) => ipcRenderer.invoke('save-settings', s),

  // File operations
  pickFile: (opts) => ipcRenderer.invoke('pick-file', opts),
  pickFolder: (opts) => ipcRenderer.invoke('pick-folder', opts),
  openPath: (p) => ipcRenderer.invoke('open-path', p),

  // Bulletin parsing
  getRecommendedBulletins: () => ipcRenderer.invoke('get-recommended-bulletins'),
  parseBulletin: (filePath) => ipcRenderer.invoke('parse-bulletin', filePath),

  // Scripture preview
  fetchScripturePreview: (ref) => ipcRenderer.invoke('fetch-scripture-preview', ref),

  // Song resolution
  findSong: (title) => ipcRenderer.invoke('find-song', title),
  setSongMapping: (title, filePath) => ipcRenderer.invoke('set-song-mapping', title, filePath),
  saveSongMappings: (mappings) => ipcRenderer.invoke('save-song-mappings', mappings),

  // PPT generation
  generatePpt: (items, outputPath) => ipcRenderer.invoke('generate-ppt', items, outputPath),

  // .ppt → .pptx batch conversion
  convertAllPpt: () => ipcRenderer.invoke('convert-all-ppt'),

  // Progress events
  onProgress: (callback) => ipcRenderer.on('progress', (_, data) => callback(data)),
  removeProgressListeners: () => ipcRenderer.removeAllListeners('progress'),

  // File path from drag-and-drop File object
  getPathForFile: (file) => webUtils.getPathForFile(file),

  // Menu events from main process
  onMenuOpen:       (cb) => ipcRenderer.on('menu-open',       () => cb()),
  onMenuSettings:   (cb) => ipcRenderer.on('menu-settings',   () => cb()),
  onMenuConvertAll: (cb) => ipcRenderer.on('menu-convert-all',() => cb()),
});
