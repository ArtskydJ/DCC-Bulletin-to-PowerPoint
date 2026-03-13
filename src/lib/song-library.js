'use strict';

const fs = require('fs');
const path = require('path');
const { execSync, spawnSync } = require('child_process');

let Fuse = null;
let fuseIndex = null;
let songFiles = [];
let lastLibraryPath = null;

function loadFuse() {
  if (!Fuse) Fuse = require('fuse.js');
}

// Build file list and fuse index from song library
function buildIndex(libraryPath) {
  if (lastLibraryPath === libraryPath && fuseIndex) return;
  loadFuse();

  if (!fs.existsSync(libraryPath)) {
    songFiles = [];
    fuseIndex = null;
    return;
  }

  const files = fs.readdirSync(libraryPath);
  songFiles = files
    .filter(f => /\.(ppt|pptx)$/i.test(f))
    .map(f => ({
      name: f,
      base: path.basename(f, path.extname(f)),
      path: path.join(libraryPath, f)
    }));

  // Deduplicate: if both .ppt and .pptx exist for same base name, prefer .pptx
  const seen = new Map();
  const deduped = [];
  for (const f of songFiles) {
    const key = f.base.toLowerCase();
    if (seen.has(key)) {
      const prev = seen.get(key);
      if (/\.pptx$/i.test(f.name)) {
        deduped[prev.idx] = f;
        seen.set(key, { idx: prev.idx });
      }
    } else {
      seen.set(key, { idx: deduped.length });
      deduped.push(f);
    }
  }
  songFiles = deduped;

  fuseIndex = new Fuse(songFiles, {
    keys: ['base'],
    threshold: 0.6,
    includeScore: true,
    ignoreLocation: true,
    minMatchCharLength: 3
  });

  lastLibraryPath = libraryPath;
}

// Normalize a song name for searching
function normalizeSongName(name) {
  return name
    .replace(/\(pages?\s*[\d,\s–-]+\)/gi, '')  // remove page refs
    .replace(/\([^)]*\)$/, '')                  // remove trailing parenthetical
    .replace(/\.\d[\d–-]*(?=\s|$)/g, '')        // remove verse ranges like ".1-4" from psalm names
    .replace(/[,;:!?]/g, ' ')                   // remove punctuation that disrupts matching
    .replace(/\s+/g, ' ')
    .trim();
}

// Search for a song by name. Returns { path, score, name } or null.
function findSong(songName, libraryPath) {
  buildIndex(libraryPath);
  if (!fuseIndex) return null;

  // Try multiple search forms and take the best result
  const normalized = normalizeSongName(songName);
  const words = normalized.split(' ');
  const candidates = [
    normalized,
    // Try progressively shorter versions (catches extra subtitle words like "Exalted High")
    words.slice(0, 5).join(' '),
    words.slice(0, 4).join(' '),
    words.slice(0, 3).join(' ')
  ].filter((s, i, a) => s && s.length >= 6 && a.indexOf(s) === i); // deduplicate, min 6 chars

  let bestResult = null;
  for (const clean of candidates) {
    const results = fuseIndex.search(clean);
    if (results.length === 0) continue;
    const r = results[0];
    if (!bestResult || r.score < bestResult.score) {
      bestResult = r;
    }
  }

  if (!bestResult || bestResult.score > 0.6) return null;

  // Hard reject if the query contains a number immediately after a word (e.g. "Psalm 89")
  // but the matched filename has a different such number (e.g. "Psalm 19").
  // Prevents "Psalm 89B" from matching "Psalm 19b", etc.
  const queryNumMatch  = normalized.match(/\b([A-Za-z]+)\s*(\d+)/i);
  const resultNumMatch = bestResult.item.base.match(/\b([A-Za-z]+)\s*(\d+)/i);
  if (queryNumMatch && resultNumMatch) {
    if (queryNumMatch[2] !== resultNumMatch[2]) return null;
  }

  return {
    path: bestResult.item.path,
    name: bestResult.item.base,
    score: bestResult.score
  };
}

// Get the pptx path for a song file, converting .ppt → .pptx if needed
function getPptxPath(songPath) {
  if (/\.pptx$/i.test(songPath)) return songPath;

  // .ppt file: check if converted version already exists
  const pptxPath = songPath.replace(/\.ppt$/i, '.pptx');
  if (fs.existsSync(pptxPath)) return pptxPath;

  // Convert using PowerShell COM automation
  convertPptToPptx(songPath, pptxPath);
  return pptxPath;
}

function convertPptToPptx(pptPath, pptxPath) {
  // Use PowerShell COM to convert
  // Note: setting Visible=msoFalse can fail if PowerPoint is already open;
  // we omit it so the window may briefly appear during conversion.
  const script = `
$ErrorActionPreference = 'Stop'
$pptApp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pptApp.Presentations.Open("${pptPath.replace(/\//g, '\\')}", $true, $false, $false)
  $pres.SaveAs("${pptxPath.replace(/\//g, '\\')}", 24)
  $pres.Close()
} finally {
  $pptApp.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
}
`;
  const result = spawnSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', script], {
    timeout: 60000,
    encoding: 'utf8'
  });
  if (result.status !== 0) {
    throw new Error(`Failed to convert ${path.basename(pptPath)}: ${result.stderr || result.stdout}`);
  }
}

// Convert all .ppt files in the library at once (batch conversion)
function convertAllPpt(libraryPath, onProgress) {
  if (!fs.existsSync(libraryPath)) return { converted: 0, errors: [] };

  const files = fs.readdirSync(libraryPath);
  const pptFiles = files.filter(f => /\.ppt$/i.test(f));
  const toConvert = pptFiles.filter(f => {
    const pptxPath = path.join(libraryPath, f.replace(/\.ppt$/i, '.pptx'));
    return !fs.existsSync(pptxPath);
  });

  let converted = 0;
  const errors = [];

  for (const f of toConvert) {
    const pptPath = path.join(libraryPath, f);
    const pptxPath = path.join(libraryPath, f.replace(/\.ppt$/i, '.pptx'));
    try {
      convertPptToPptx(pptPath, pptxPath);
      converted++;
      if (onProgress) onProgress(converted, toConvert.length, f);
    } catch (e) {
      errors.push({ file: f, error: e.message });
    }
  }

  return { converted, total: toConvert.length, errors };
}

// List all songs in library (as name + path)
function listSongs(libraryPath) {
  buildIndex(libraryPath);
  return songFiles.map(f => ({ name: f.base, path: f.path }));
}

module.exports = { findSong, getPptxPath, convertAllPpt, listSongs, buildIndex };
