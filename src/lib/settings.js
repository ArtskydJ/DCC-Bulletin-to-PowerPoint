'use strict';

const path = require('path');
const os = require('os');

let store = null;

function getStore() {
  if (!store) {
    const Store = require('electron-store');
    store = new Store({
      name: 'config',
      defaults: {
        songLibraryPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Songs'),
        outputPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Completed'),
        bulletinPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'WorshipPrograms'),
        referencePptxPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Completed', '2026-03-08.pptx'),
        songMappings: {},
        recentFiles: []
      }
    });
  }
  return store;
}

function get(key) {
  return getStore().get(key);
}

function set(key, value) {
  getStore().set(key, value);
}

function getAll() {
  const s = getStore();
  return {
    songLibraryPath: s.get('songLibraryPath'),
    outputPath: s.get('outputPath'),
    bulletinPath: s.get('bulletinPath'),
    referencePptxPath: s.get('referencePptxPath'),
    songMappings: s.get('songMappings'),
    recentFiles: s.get('recentFiles')
  };
}

function saveSettings(settings) {
  const s = getStore();
  if (settings.songLibraryPath !== undefined) s.set('songLibraryPath', settings.songLibraryPath);
  if (settings.outputPath !== undefined) s.set('outputPath', settings.outputPath);
  if (settings.bulletinPath !== undefined) s.set('bulletinPath', settings.bulletinPath);
  if (settings.referencePptxPath !== undefined) s.set('referencePptxPath', settings.referencePptxPath);
}

function getSongMapping(songName) {
  const mappings = getStore().get('songMappings') || {};
  const lower = songName.toLowerCase();
  // Case-insensitive lookup so stored keys can keep their original casing
  const key = Object.keys(mappings).find(k => k.toLowerCase() === lower);
  return key ? mappings[key] : null;
}

function setSongMapping(songName, filePath) {
  const s = getStore();
  const mappings = s.get('songMappings') || {};
  const lower = songName.toLowerCase();
  // Remove any existing entry with a different casing before writing
  const existing = Object.keys(mappings).find(k => k.toLowerCase() === lower);
  if (existing) delete mappings[existing];
  mappings[songName] = filePath;
  s.set('songMappings', mappings);
}

function addRecentFile(filePath) {
  const s = getStore();
  const recent = s.get('recentFiles') || [];
  const norm = p => p.replace(/\\/g, '/').toLowerCase();
  const filtered = recent.filter(f => norm(f) !== norm(filePath));
  filtered.unshift(filePath);
  s.set('recentFiles', filtered.slice(0, 10));
}

module.exports = { get, set, getAll, saveSettings, getSongMapping, setSongMapping, addRecentFile };
