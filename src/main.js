'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

let mainWindow = null;

// Persist window bounds across sessions using a simple JSON file
const windowStatePath = path.join(app.getPath('userData'), 'window-state.json');
function loadWindowState() {
  try { return JSON.parse(fs.readFileSync(windowStatePath, 'utf8')); } catch { return {}; }
}
function saveWindowState(win) {
  if (win.isMaximized() || win.isMinimized()) return;
  fs.writeFileSync(windowStatePath, JSON.stringify(win.getBounds()));
}

function createWindow() {
  const saved = loadWindowState();
  mainWindow = new BrowserWindow({
    width:  saved.width  || 960,
    height: saved.height || 700,
    x: saved.x,
    y: saved.y,
    minWidth: 700,
    minHeight: 500,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    title: 'DCC PPT Generator',
    icon: path.join(__dirname, '..', 'DCC.ico')
  });

  mainWindow.on('resize', () => saveWindowState(mainWindow));
  mainWindow.on('move',   () => saveWindowState(mainWindow));

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));

  // Build menu
  const menu = Menu.buildFromTemplate([
    {
      label: 'File',
      submenu: [
        { label: 'Open Bulletin...', accelerator: 'CmdOrCtrl+O', click: () => mainWindow.webContents.send('menu-open') },
        { type: 'separator' },
        { label: 'Settings', accelerator: 'CmdOrCtrl+,', click: () => mainWindow.webContents.send('menu-settings') },
        { type: 'separator' },
        { label: 'Quit', accelerator: 'CmdOrCtrl+Q', click: () => app.quit() }
      ]
    },
    {
      label: 'Advanced',
      submenu: [
        { label: 'Toggle DevTools', click: () => mainWindow.webContents.toggleDevTools() }
      ]
    }
  ]);
  Menu.setApplicationMenu(menu);
}

app.whenReady().then(() => {
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ── IPC Handlers ────────────────────────────────────────────────────────────

ipcMain.handle('get-settings', () => {
  const settings = require('./lib/settings');
  return settings.getAll();
});

ipcMain.handle('save-settings', (_, newSettings) => {
  const settings = require('./lib/settings');
  settings.saveSettings(newSettings);
  return { ok: true };
});

ipcMain.handle('pick-file', async (_, opts = {}) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: opts.title || 'Select File',
    defaultPath: opts.defaultPath ? path.normalize(opts.defaultPath) : undefined,
    filters: opts.filters || [{ name: 'All Files', extensions: ['*'] }],
    properties: ['openFile']
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('pick-folder', async (_, opts = {}) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: opts.title || 'Select Folder',
    defaultPath: opts.defaultPath ? path.normalize(opts.defaultPath) : undefined,
    properties: ['openDirectory']
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('open-path', (_, filePath) => {
  shell.openPath(filePath);
});

ipcMain.handle('get-recommended-bulletins', () => {
  const settings = require('./lib/settings');
  const bulletinDir = settings.get('bulletinPath');
  if (!bulletinDir || !fs.existsSync(bulletinDir)) return [];
  const today = new Date().toISOString().slice(0, 10);
  return fs.readdirSync(bulletinDir)
    .filter(f => /^\d{4}-\d{2}-\d{2}/.test(f) && /\.docx$/i.test(f) && f.slice(0, 10) >= today)
    .sort()
    .map(f => path.join(bulletinDir, f));
});

ipcMain.handle('parse-bulletin', async (_, filePath) => {
  const { parseBulletin } = require('./lib/bulletin-parser');
  const settings = require('./lib/settings');
  const { findSong, getPptxPath } = require('./lib/song-library');

  settings.addRecentFile(filePath);
  const libraryPath = settings.get('songLibraryPath');

  const items = await parseBulletin(filePath);

  // For each song, attempt fuzzy match
  for (const item of items) {
    if (item.type !== 'song') continue;

    // Check cached mapping first
    const cached = settings.getSongMapping(item.title);
    if (cached && fs.existsSync(cached)) {
      item.resolvedPath = cached;
      item.resolved = true;
      continue;
    }

    const match = findSong(item.title, libraryPath);
    if (match) {
      item.resolvedPath = match.path;
      item.matchName = match.name;
      item.resolved = true;
      item.matchScore = match.score;
      // Persist so it appears in Settings → Song Mappings for review/correction
      settings.setSongMapping(item.title, match.path);
    } else {
      item.resolved = false;
    }
  }

  return items;
});

ipcMain.handle('fetch-scripture-preview', async (_, ref) => {
  const { fetchScripture } = require('./lib/bible-fetcher');
  try {
    const slides = await fetchScripture(ref);
    return (slides[0]?.lines[0] || '').replace(/^\d+\s*/, '').trim();
  } catch {
    return null;
  }
});

ipcMain.handle('find-song', async (_, title) => {
  const settings = require('./lib/settings');
  const { findSong } = require('./lib/song-library');
  const libraryPath = settings.get('songLibraryPath');
  return findSong(title, libraryPath);
});

ipcMain.handle('set-song-mapping', (_, title, filePath) => {
  const settings = require('./lib/settings');
  settings.setSongMapping(title, filePath);
  return { ok: true };
});

ipcMain.handle('save-song-mappings', (_, mappings) => {
  const settings = require('./lib/settings');
  settings.set('songMappings', mappings);
  return { ok: true };
});

ipcMain.handle('generate-ppt', async (event, items, outputPath) => {
  const settings = require('./lib/settings');
  const { buildPptx, splitReadingLines } = require('./lib/ppt-builder');
  const { fetchScripture } = require('./lib/bible-fetcher');
  const { getPptxPath } = require('./lib/song-library');
  const referencePptxPath = settings.get('referencePptxPath');
  if (!referencePptxPath || !fs.existsSync(referencePptxPath)) {
    throw new Error('Reference PPTX not found. Please set it in Settings.');
  }

  const send = (msg) => {
    if (mainWindow) mainWindow.webContents.send('progress', msg);
  };

  // Build slide descriptors, fetching scripture as needed
  const slideDescriptors = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    send({ step: i + 1, total: items.length, message: `Processing: ${item.title || item.ref || item.type}` });

    if (item.type === 'reading') {
      slideDescriptors.push({ ...item, slideGroups: splitReadingLines(item.lines) });
      continue;
    }

    if (item.type === 'scripture') {
      send({ step: i + 1, total: items.length, message: `Fetching scripture: ${item.ref}` });
      try {
        const slides = await fetchScripture(item.ref);
        slideDescriptors.push({ ...item, slides });
      } catch (e) {
        slideDescriptors.push({
          ...item,
          slides: [{ title: item.ref, lines: [`Error fetching scripture: ${e.message}`] }]
        });
      }
      continue;
    }

    if (item.type === 'song') {
      let pptxPath = item.resolvedPath || null;

      // Convert .ppt → .pptx if needed
      if (pptxPath && /\.ppt$/i.test(pptxPath)) {
        send({ step: i + 1, total: items.length, message: `Converting .ppt: ${path.basename(pptxPath)}` });
        try {
          pptxPath = getPptxPath(pptxPath);
        } catch (e) {
          pptxPath = null;
        }
      }

      slideDescriptors.push({ ...item, pptxPath });
      continue;
    }
  }

  send({ step: items.length, total: items.length, message: 'Building PPTX...' });

  const { buffer, warnings } = await buildPptx(referencePptxPath, slideDescriptors, (step, total, desc) => {
    send({ step, total, message: `Adding slide: ${desc.title || desc.ref || desc.type}` });
  });

  // Write output — on EBUSY (file open in PowerPoint), prompt user to close it then retry once
  try {
    fs.writeFileSync(outputPath, buffer);
  } catch (err) {
    if (err.code !== 'EBUSY') throw err;
    await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      title: 'File In Use',
      message: 'The output file is open in another program (e.g. PowerPoint).\n\nClose it, then click OK.',
      buttons: ['OK'],
      defaultId: 0,
    });
    fs.writeFileSync(outputPath, buffer); // if still busy, let the error propagate normally
  }
  send({ done: true, outputPath, warnings });

  return { ok: true, outputPath, warnings };
});

ipcMain.handle('convert-all-ppt', async (event) => {
  const settings = require('./lib/settings');
  const { convertAllPpt } = require('./lib/song-library');
  const libraryPath = settings.get('songLibraryPath');

  const result = convertAllPpt(libraryPath, (converted, total, name) => {
    if (mainWindow) mainWindow.webContents.send('progress', { step: converted, total, message: `Converting: ${name}` });
  });

  return result;
});

