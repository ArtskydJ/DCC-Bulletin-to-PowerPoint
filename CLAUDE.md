# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Run the Electron app
npm start

# CLI (faster for testing — no Electron needed)
# output.pptx is optional; if omitted, output path is derived from settings + bulletin date.
# Always prints a text summary to stdout (readings show **bold** All lines, plain Leader lines).
node src/cli.js "path/to/bulletin.docx" ["path/to/output.pptx"] [--open]

# Build Windows installer (outputs to dist/)
npm run build:win

# Verify a generated PPTX opens correctly (requires PowerPoint installed)
powershell -File test-pptx.ps1 "path/to/file.pptx"
```

There are no automated tests. Manual testing uses the CLI against `C:/Users/Joseph/dcc/Bulletins docx/` files, verified by opening in PowerPoint.

## Architecture

This is an Electron app with a vanilla HTML/JS renderer and a Node.js main process. The core pipeline runs entirely in the main process and is also exposed via CLI.

### Directory structure

```
src/main.js, src/preload.js, src/cli.js  — entry points
src/lib/        — pipeline modules
src/renderer/   — HTML/CSS/JS frontend
```

### Data flow

```
.docx bulletin
    → bulletin-parser.js  (mammoth → item list)
    → song-library.js     (fuse.js fuzzy match .pptx files by name)
    → bible-fetcher.js    (local NKJV JSON lookup for scripture items)
    → ppt-builder.js      (ZIP manipulation → output .pptx buffer)
```

Each item in the parsed list has `type`: `'reading'` | `'song'` | `'scripture'` | `'skip'`.

### bible-fetcher.js

Reads from `src/lib/assets/NEW KING JAMES VERSION.json` — no network requests. Handles full chapters, single verses, and ranges; normalizes dot notation (`16.6` → `16:6`) and common book name aliases (`Psalm` → `Psalms`, etc.).
### ppt-builder.js — critical PPTX XML facts

The builder works by copying a **reference PPTX** (an existing completed slideshow) and replacing its slide content. It does NOT use pptxgenjs; it manipulates ZIP entries and XML directly via `adm-zip`.

Key invariants that must be preserved:
- Slide list entries in `ppt/presentation.xml` use `<p:sldId>` (NOT `<p:sld>`) — wrong tag = 0 slides shown
- `ppt/_rels/presentation.xml.rels` rId numbers: parse with `Id="rId(\d+)"` capturing only digits; `parseInt("rId3")` = NaN
- New slide IDs start at 257 (slide1 from reference is reset to 256); do NOT scan all `id=` attrs in presentation.xml to find max — slide master IDs like `2147483648` are present and would overflow
- `[Content_Types].xml` is rebuilt from scratch by scanning actual ZIP entries
- Song slides copied from library `.pptx` files get their rels sanitized: `notesSlides` refs stripped, missing `slideLayout` refs fall back to `slideLayout1`
- `adm-zip`: never call `updateFile()` on an entry added via `addFile()` — it silently does nothing

### Electron IPC boundary

All `window.dcc.*` calls in the renderer go through `preload.js` → `ipcMain.handle()` in `main.js`. The renderer has no direct Node access (`contextIsolation: true`, `nodeIntegration: false`). Drag-and-drop file paths require `webUtils.getPathForFile(file)` (not `file.path`).

### Settings

Persisted via `electron-store` at `%APPDATA%/dcc-pptgen/config.json`. Key settings:
- `referencePptxPath` — a completed .pptx whose title slide (slide1) and master/layouts are copied into every output file
- `songLibraryPath` — directory of 800+ .ppt/.pptx song files
- `songMappings` — manual overrides: lowercase song title → absolute file path
- `recentFiles` — last 10 opened bulletin paths

The CLI reads this same config file directly (bypassing electron-store).

### Song matching

`song-library.js` builds a `fuse.js` index of filenames (lazy, cached). It tries the full normalized title, then progressively shorter prefixes (5, 4, 3 words). Threshold 0.6. `.ppt` files are auto-converted to `.pptx` on first use via PowerShell COM (`SaveAs` format 24). Both `.ppt` and `.pptx` for the same base name are deduplicated in favor of `.pptx`.

### Bulletin parsing

`bulletin-parser.js` extracts raw text via `mammoth` then processes line-by-line. Songs are identified by `(page N)` patterns. Readings are identified by known heading strings (see `READING_HEADINGS`). Catechism, scripture refs, and unknown headings followed by `(Leader)`/`(All)` lines are also handled. `∆` prefix indicates "please stand" but appears on both songs and other items — song detection uses the page-number heuristic, not the `∆`.

### Local reference paths

- Song library: `C:/Users/Joseph/dcc/Slideshows/DCC_PPT_songs/ppt/`
- Completed slideshows: `C:/Users/Joseph/dcc/Slideshows/PowerPoint Completed/`
- Reference PPTX: `C:/Users/Joseph/dcc/Slideshows/PowerPoint Completed/2026-03-08.pptx`
- Bulletins: `C:/Users/Joseph/dcc/Bulletins docx/`
