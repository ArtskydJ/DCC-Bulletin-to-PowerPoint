#!/usr/bin/env node
/**
 * DCC Bulletin → PowerPoint CLI
 * Usage: node dcc-bulletin-to-pptx-cli.js <input.docx> [output.pptx]
 *
 * Uses the default config paths. Assumes all parsing is correct (no interactive
 * song resolution dialogs — uses cached mappings and fuzzy matching only).
 */
'use strict';

const path = require('path');
const fs = require('fs');
const os = require('os');

// ── Args ─────────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
const openAfter = args.includes('--open');
const filteredArgs = args.filter(a => a !== '--open');
const [inputDocx, outputPptxArg] = filteredArgs;

if (!inputDocx) {
  console.error('Usage: node src/cli.js <input.docx> [output.pptx] [--open]');
  process.exit(1);
}

if (!fs.existsSync(inputDocx)) {
  console.error('Input file not found:', inputDocx);
  process.exit(1);
}

// ── Config / defaults ─────────────────────────────────────────────────────
// Read electron-store config if it exists, otherwise use defaults
const configPath = path.join(os.homedir(), 'AppData', 'Roaming', 'dcc-pptgen', 'config.json');
let config = {};
if (fs.existsSync(configPath)) {
  try { config = JSON.parse(fs.readFileSync(configPath, 'utf8')); } catch (e) {}
}

const DEFAULTS = {
  songLibraryPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Songs'),
  outputPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Completed'),
  referencePptxPath: path.join(os.homedir(), 'Dropbox', 'WorshipMaterials', 'SlideShows', 'Powerpoint Completed', '2026-03-08.pptx'),
  songMappings: {}
};
const songLibraryPath = config.songLibraryPath || DEFAULTS.songLibraryPath;
const songMappings = config.songMappings || DEFAULTS.songMappings;
const referencePptxPath = config.referencePptxPath || DEFAULTS.referencePptxPath;

if (!fs.existsSync(referencePptxPath)) {
  console.error('Reference PPTX not found:', referencePptxPath);
  console.error('Set referencePptxPath in config.json or update the DEFAULTS in this script.');
  process.exit(1);
}

// Determine output path
const dateMatch = path.basename(inputDocx).match(/(\d{4}-\d{2}-\d{2})/);
const dateStr = dateMatch ? dateMatch[1] : new Date().toISOString().slice(0, 10);
const outputPptx = outputPptxArg || path.join(
  config.outputPath || DEFAULTS.outputPath,
  dateStr + '.pptx'
);

// ── Run ───────────────────────────────────────────────────────────────────
async function main() {
  console.log('Input:', inputDocx);
  console.log('Output:', outputPptx);
  console.log('Song library:', songLibraryPath);

  // 1. Parse bulletin
  process.stdout.write('Parsing bulletin... ');
  const { parseBulletin } = require('./lib/bulletin-parser');
  const items = await parseBulletin(inputDocx);
  console.log(`found ${items.length} items`);

  // 2. Resolve songs
  const { findSong, getPptxPath } = require('./lib/song-library');
  const warnings = [];
  let songCount = 0;

  for (const item of items) {
    if (item.type !== 'song') continue;
    songCount++;

    // Check manual mapping
    const mapped = songMappings[item.title.toLowerCase()];
    if (mapped && fs.existsSync(mapped)) {
      item.resolvedPath = mapped;
      item.resolved = true;
      continue;
    }

    // Fuzzy match
    const match = findSong(item.title, songLibraryPath);
    if (match) {
      item.resolvedPath = match.path;
      item.resolved = true;
      console.log(`  Song matched: "${item.title}" → ${path.basename(match.path)} (score: ${match.score.toFixed(3)})`);
    } else {
      item.resolved = false;
      warnings.push({ type: 'missing-song', title: item.title });
      console.warn(`  ⚠ Song not found: "${item.title}" — inserting placeholder`);
    }
  }

  // 3. Fetch scripture
  const { fetchScripture } = require('./lib/bible-fetcher');
  const slideDescriptors = [];

  for (const item of items) {
    if (item.type === 'skip') continue;

    if (item.type === 'reading') {
      const { splitReadingLines } = require('./lib/ppt-builder');
      slideDescriptors.push({ ...item, slideGroups: splitReadingLines(item.lines) });
      continue;
    }

    if (item.type === 'scripture') {
      process.stdout.write(`  Fetching scripture: ${item.ref}... `);
      try {
        const slides = await fetchScripture(item.ref);
        console.log(`${slides.length} slide(s)`);
        slideDescriptors.push({ ...item, slides });
      } catch (e) {
        console.error(`FAILED: ${e.message}`);
        warnings.push({ type: 'scripture-error', ref: item.ref, error: e.message });
        slideDescriptors.push({
          ...item,
          slides: [{ title: item.ref, lines: [`[Error fetching: ${e.message}]`] }]
        });
      }
      continue;
    }

    if (item.type === 'song') {
      let pptxPath = item.resolvedPath || null;
      if (pptxPath && /\.ppt$/i.test(pptxPath)) {
        process.stdout.write(`  Converting .ppt: ${path.basename(pptxPath)}... `);
        try {
          pptxPath = getPptxPath(pptxPath);
          console.log('done');
        } catch (e) {
          console.error(`FAILED: ${e.message}`);
          warnings.push({ type: 'ppt-convert-error', title: item.title, error: e.message });
          pptxPath = null;
        }
      }
      slideDescriptors.push({ ...item, pptxPath });
      continue;
    }
  }

  // ── Text dump to stdout ───────────────────────────────────────────────
  {
    const lines = [];
    const dateLabel = path.basename(inputDocx).match(/(\d{4}-\d{2}-\d{2})/)?.[1] ?? path.basename(inputDocx);
    lines.push(`=== Bulletin: ${dateLabel} ===\n`);

    for (const item of slideDescriptors) {
      if (item.type === 'song') {
        const status = item.resolvedPath ? `→ ${path.basename(item.resolvedPath)}` : '→ [NOT FOUND]';
        lines.push(`[SONG] ${item.title}`);
        lines.push(`  ${status}`);

      } else if (item.type === 'reading') {
        lines.push(`[READING] ${item.title}`);
        item.slideGroups.forEach((slideLines, si) => {
          if (si > 0) lines.push('  ---');
          for (const l of slideLines) {
            const txt = l.role === 'All' ? `**${l.text}**` : l.text;
            lines.push(`  ${l.role}: ${txt}`);
          }
        });

      } else if (item.type === 'scripture') {
        lines.push(`[SCRIPTURE] ${item.ref}`);
        (item.slides || []).forEach((slide, si) => {
          if (si > 0) lines.push('  ---');
          lines.push(`  ${slide.title}`);
          for (const t of slide.lines) lines.push(`    ${t}`);
        });
      }
      lines.push('');
    }

    console.log(lines.join('\n'));
  }

  // 4. Build PPTX
  process.stdout.write('Building PPTX... ');
  const { buildPptx } = require('./lib/ppt-builder');
  const { buffer, warnings: buildWarnings } = await buildPptx(
    referencePptxPath,
    slideDescriptors,
    (step, total, desc) => {
      // Compact progress: print dots
      process.stdout.write('.');
    }
  );
  console.log(' done');

  // 5. Save
  const outDir = path.dirname(outputPptx);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(outputPptx, buffer);
  console.log('\n✅ Saved to:', outputPptx);

  if (openAfter) {
    const { spawn } = require('child_process');
    spawn('cmd', ['/c', 'start', '', outputPptx], { detached: true, stdio: 'ignore' }).unref();
  }

  // 6. Print warnings
  // buildWarnings already includes missing-song entries; merge without duplicating
  const allWarnings = [...warnings, ...buildWarnings.filter(w => w.type !== 'missing-song')];
  if (allWarnings.length > 0) {
    console.log('\n⚠ Warnings:');
    allWarnings.forEach(w => {
      if (w.type === 'missing-song') console.warn(`  • Missing song: "${w.title}"`);
      else if (w.type === 'song-error') console.warn(`  • Song error (${w.title}): ${w.error}`);
      else if (w.type === 'scripture-error') console.warn(`  • Scripture error (${w.ref}): ${w.error}`);
      else if (w.type === 'ppt-convert-error') console.warn(`  • .ppt conversion failed (${w.title}): ${w.error}`);
      else console.warn(`  • ${JSON.stringify(w)}`);
    });
  }
}

main().catch(e => {
  console.error('Fatal error:', e);
  process.exit(1);
});
