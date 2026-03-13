'use strict';

// ── State ────────────────────────────────────────────────────────────────
let currentItems = [];   // parsed bulletin items
let currentBulletinPath = '';
let outputPath = '';
let settings = {};

// ── View management ──────────────────────────────────────────────────────
function showView(id) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

// ── Helpers ──────────────────────────────────────────────────────────────
function basename(p) {
  return p.replace(/\\/g, '/').split('/').pop();
}
function basenameNoExt(p) {
  const b = basename(p);
  return b.replace(/\.[^.]+$/, '');
}
function dateSuggest(bulletinPath) {
  const m = basename(bulletinPath).match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : new Date().toISOString().slice(0, 10);
}
function buildOutputPath(bulletinPath) {
  const date = dateSuggest(bulletinPath);
  const outDir = settings.outputPath || '';
  return outDir ? (outDir.replace(/[\\/]+$/, '') + '/' + date + '.pptx') : date + '.pptx';
}

// ── Init ─────────────────────────────────────────────────────────────────
async function init() {
  settings = await window.dcc.getSettings();
  loadRecommendedFiles();
  loadRecentFiles();

  // Auto-detect first run: if settings paths don't exist, show settings
  const { songLibraryPath, outputPath: op } = settings;
  const pathsMissing = !songLibraryPath;
  if (pathsMissing) showView('settings-view');
  else showView('main-view');

  setupDropZone();
  setupSettings();
  setupReview();
  setupResult();
setupMenuListeners();
  lucide.createIcons();
}

// ── Recommended files ─────────────────────────────────────────────────────
async function loadRecommendedFiles() {
  const section = document.getElementById('recommended-section');
  const list = document.getElementById('recommended-list');
  list.innerHTML = '';
  try {
    const files = await window.dcc.getRecommendedBulletins();
    if (!files || files.length === 0) { section.style.display = 'none'; return; }
    section.style.display = 'block';
    files.forEach(f => {
      const li = document.createElement('li');
      li.textContent = basename(f);
      li.title = f;
      li.addEventListener('click', () => openBulletin(f));
      list.appendChild(li);
    });
  } catch { section.style.display = 'none'; }
}

// ── Recent files ─────────────────────────────────────────────────────────
function loadRecentFiles() {
  const recent = settings.recentFiles || [];
  const section = document.getElementById('recent-section');
  const list = document.getElementById('recent-list');
  list.innerHTML = '';
  if (recent.length === 0) { section.style.display = 'none'; return; }
  section.style.display = 'block';
  recent.forEach(f => {
    const li = document.createElement('li');
    li.textContent = basename(f);
    li.title = f;
    li.addEventListener('click', () => openBulletin(f));
    list.appendChild(li);
  });
}

// ── Drop zone ─────────────────────────────────────────────────────────────
function setupDropZone() {
  const zone = document.getElementById('drop-zone');

  zone.addEventListener('click', async () => {
    const f = await window.dcc.pickFile({
      title: 'Open Bulletin',
      filters: [{ name: 'Word Document', extensions: ['docx'] }],
      defaultPath: settings.bulletinPath
    });
    if (f) openBulletin(f);
  });

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.docx'));
    if (files.length > 0) openBulletin(window.dcc.getPathForFile(files[0]));
  });

  document.getElementById('settings-btn').addEventListener('click', () => {
    populateSettings();
    showView('settings-view');
  });
}

// ── Open bulletin → parse → review ──────────────────────────────────────
async function openBulletin(filePath) {
  currentBulletinPath = filePath;
  document.getElementById('drop-zone').innerHTML = '<div class="progress-spinner" style="margin-top:8px"></div><h2>Parsing bulletin…</h2><p>' + basename(filePath) + '</p>';

  try {
    const items = await window.dcc.parseBulletin(filePath);
    currentItems = items;
    settings = await window.dcc.getSettings();
    loadRecentFiles();
    populateReview(filePath, items);
    showView('review-view');
  } catch (e) {
    alert('Error parsing bulletin:\n' + e.message);
  }

  // Restore drop zone
  document.getElementById('drop-zone').innerHTML = '<div class="drop-icon"><i data-lucide="file-text"></i></div><h2>Drop Bulletin Here</h2><p>Drop a .docx bulletin file, or click to browse</p>';
  lucide.createIcons();
}

// ── Review ────────────────────────────────────────────────────────────────
function populateReview(filePath, items) {
  document.getElementById('review-title').textContent = 'Parsed Bulletin — ' + dateSuggest(filePath);
  document.getElementById('review-subtitle').textContent = `${items.length} items found.`;

  const tbody = document.getElementById('review-body');
  tbody.innerHTML = '';

  items.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.dataset.idx = idx;

    // Type badge
    const tdType = document.createElement('td');
    tdType.className = 'type-cell';
    const badge = document.createElement('span');
    badge.className = `badge badge-${item.type}`;
    badge.textContent = item.type.charAt(0).toUpperCase() + item.type.slice(1);
    tdType.appendChild(badge);
    tr.appendChild(tdType);

    // Title / content
    const tdTitle = document.createElement('td');
    tdTitle.className = 'title-cell';
    if (item.type === 'scripture') {
      const refInput = document.createElement('input');
      refInput.type = 'text';
      refInput.className = 'ref-edit';
      refInput.value = item.ref || '';
      refInput.title = 'Click to edit scripture reference';
      refInput.addEventListener('change', () => {
        item.ref = refInput.value.trim();
        updateInfoCell(tdInfo, item);
      });
      tdTitle.appendChild(refInput);
    } else {
      tdTitle.textContent = item.title || item.ref || '(unknown)';
    }
    tr.appendChild(tdTitle);

    // Info/details
    const tdInfo = document.createElement('td');
    tdInfo.className = 'info-cell';
    updateInfoCell(tdInfo, item);
    tr.appendChild(tdInfo);

    tbody.appendChild(tr);
  });
}

function updateInfoCell(td, item) {
  if (item.type === 'song') {
    if (item.resolved && item.resolvedPath) {
      td.innerHTML = `<span class="badge badge-found"><i data-lucide="check"></i> matched</span> ${basenameNoExt(item.resolvedPath)}`;
      lucide.createIcons({ nodes: [td] });
    } else {
      td.innerHTML = '';
      const badge = document.createElement('span');
      badge.className = 'badge badge-missing';
      badge.innerHTML = '<i data-lucide="x"></i> not found';
      lucide.createIcons({ nodes: [badge] });
      const btn = document.createElement('button');
      btn.className = 'btn btn-secondary btn-sm';
      btn.textContent = 'Browse';
      btn.style.marginLeft = '8px';
      btn.addEventListener('click', async () => {
        const f = await window.dcc.pickFile({
          title: 'Locate Song File',
          filters: [{ name: 'PowerPoint', extensions: ['pptx', 'ppt'] }],
          defaultPath: settings.songLibraryPath
        });
        if (f) {
          item.resolvedPath = f;
          item.resolved = true;
          await window.dcc.setSongMapping(item.title, f);
          updateInfoCell(td, item);
        }
      });
      td.appendChild(badge);
      td.appendChild(btn);
    }
  } else if (item.type === 'reading') {
    const lines = item.lines || [];
    td.textContent = lines.map(l => `${l.role}: ${l.text}`).join(' / ') || `${lines.length} line(s)`;
  } else if (item.type === 'scripture') {
    td.innerHTML = '';
    window.dcc.fetchScripturePreview(item.ref).then(preview => {
      if (preview) {
        td.innerHTML = `<span class="badge badge-found"><i data-lucide="check"></i> valid</span> ${preview}`;
      } else {
        td.innerHTML = `<span class="badge badge-missing"><i data-lucide="x"></i> invalid</span>`;
      }
      lucide.createIcons({ nodes: [td] });
    }).catch(() => {
      td.innerHTML = `<span class="badge badge-missing"><i data-lucide="x"></i> invalid</span>`;
      lucide.createIcons({ nodes: [td] });
    });
  } else if (item.type === 'skip') {
    td.textContent = 'Will be omitted';
  }
}


function setupReview() {
  document.getElementById('review-cancel-btn').addEventListener('click', () => showView('main-view'));
  document.getElementById('review-generate-btn').addEventListener('click', () => startGeneration());
}

// ── Generation ─────────────────────────────────────────────────────────────
async function startGeneration() {
  // Filter out skipped items
  const toProcess = currentItems.filter(item => item.type !== 'skip');

  // Determine output path
  settings = await window.dcc.getSettings();
  outputPath = buildOutputPath(currentBulletinPath);

  // Show progress view
  showView('progress-view');
  document.getElementById('progress-bar').style.width = '0%';
  document.getElementById('progress-msg').textContent = 'Starting…';

  window.dcc.onProgress(({ step, total, message, done, outputPath: op, warnings }) => {
    if (done) {
      window.dcc.removeProgressListeners();
      showResult(op || outputPath, warnings || []);
      return;
    }
    const pct = total > 0 ? Math.round((step / total) * 100) : 0;
    document.getElementById('progress-bar').style.width = pct + '%';
    document.getElementById('progress-msg').textContent = message || '';
  });

  try {
    await window.dcc.generatePpt(toProcess, outputPath);
  } catch (e) {
    window.dcc.removeProgressListeners();
    alert('Error generating slideshow:\n' + e.message);
    showView('review-view');
  }
}


// ── Result ─────────────────────────────────────────────────────────────────
function showResult(outPath, warnings) {
  outputPath = outPath;
  document.getElementById('result-path').textContent = outPath;

  const warnBox = document.getElementById('warnings-box');
  const warnList = document.getElementById('warnings-list');
  warnList.innerHTML = '';
  if (warnings && warnings.length > 0) {
    warnBox.style.display = 'block';
    warnings.forEach(w => {
      const li = document.createElement('li');
      if (w.type === 'missing-song') li.textContent = `Missing song: "${w.title}"`;
      else if (w.type === 'song-error') li.textContent = `Song error (${w.title}): ${w.error}`;
      else li.textContent = JSON.stringify(w);
      warnList.appendChild(li);
    });
  } else {
    warnBox.style.display = 'none';
  }

  showView('result-view');
}

function setupResult() {
  document.getElementById('result-open-btn').addEventListener('click', () => {
    window.dcc.openPath(outputPath);
  });
  document.getElementById('result-bulletin-btn').addEventListener('click', () => {
    window.dcc.openPath(currentBulletinPath);
  });
  document.getElementById('result-again-btn').addEventListener('click', () => {
    currentItems = [];
    currentBulletinPath = '';
    showView('main-view');
  });
}

// ── Settings ──────────────────────────────────────────────────────────────
function populateSettings() {
  document.getElementById('setting-library').value = settings.songLibraryPath || '';
  document.getElementById('setting-output').value = settings.outputPath || '';
  document.getElementById('setting-bulletin').value = settings.bulletinPath || '';
  document.getElementById('setting-reference').value = settings.referencePptxPath || '';
  populateMappings(settings.songMappings || {});
}

function populateMappings(mappings) {
  const tbody = document.getElementById('mappings-body');
  tbody.innerHTML = '';
  const sorted = Object.entries(mappings).sort((a, b) => a[0].toLowerCase().localeCompare(b[0].toLowerCase()));
  for (const [title, filePath] of sorted) {
    tbody.appendChild(makeMappingRow(title, filePath));
  }
}

function makeMappingRow(title = '', filePath = '') {
  const tr = document.createElement('tr');

  // Title cell
  const tdTitle = document.createElement('td');
  const titleInput = document.createElement('input');
  titleInput.type = 'text';
  titleInput.className = 'mapping-title';
  titleInput.value = title;
  titleInput.placeholder = 'song title';
  tdTitle.appendChild(titleInput);
  tr.appendChild(tdTitle);

  // File cell
  const tdFile = document.createElement('td');
  const fileRow = document.createElement('div');
  fileRow.className = 'path-row';
  const fileInput = document.createElement('input');
  fileInput.type = 'text';
  fileInput.className = 'mapping-file';
  fileInput.readOnly = true;
  fileInput.placeholder = 'no file selected';
  fileInput.title = filePath;
  fileInput.value = filePath ? basename(filePath) : '';
  fileInput.dataset.fullpath = filePath;
  const browseBtn = document.createElement('button');
  browseBtn.className = 'btn btn-secondary btn-sm';
  browseBtn.textContent = 'Browse';
  browseBtn.addEventListener('click', async () => {
    const f = await window.dcc.pickFile({
      title: 'Select Song File',
      filters: [{ name: 'PowerPoint', extensions: ['pptx', 'ppt'] }],
      defaultPath: fileInput.dataset.fullpath || settings.songLibraryPath
    });
    if (f) {
      fileInput.value = basename(f);
      fileInput.title = f;
      fileInput.dataset.fullpath = f;
    }
  });
  fileRow.appendChild(fileInput);
  fileRow.appendChild(browseBtn);
  tdFile.appendChild(fileRow);
  tr.appendChild(tdFile);

  // Delete cell
  const tdDel = document.createElement('td');
  const delBtn = document.createElement('button');
  delBtn.className = 'btn btn-danger btn-sm';
  delBtn.textContent = 'Delete';
  delBtn.addEventListener('click', () => tr.remove());
  tdDel.appendChild(delBtn);
  tr.appendChild(tdDel);

  return tr;
}

function collectMappings() {
  const rows = document.querySelectorAll('#mappings-body tr');
  const mappings = {};
  for (const tr of rows) {
    const title = tr.cells[0].querySelector('input').value.trim();
    const fullpath = tr.cells[1].querySelector('input').dataset.fullpath || '';
    if (title && fullpath) mappings[title] = fullpath;
  }
  return mappings;
}

function setupSettings() {
  document.getElementById('browse-library').addEventListener('click', async () => {
    const f = await window.dcc.pickFolder({ title: 'Select Song Library Folder', defaultPath: document.getElementById('setting-library').value });
    if (f) document.getElementById('setting-library').value = f;
  });
  document.getElementById('browse-output').addEventListener('click', async () => {
    const f = await window.dcc.pickFolder({ title: 'Select Output Folder', defaultPath: document.getElementById('setting-output').value });
    if (f) document.getElementById('setting-output').value = f;
  });
  document.getElementById('browse-bulletin').addEventListener('click', async () => {
    const f = await window.dcc.pickFolder({ title: 'Select Bulletin Folder', defaultPath: document.getElementById('setting-bulletin').value });
    if (f) document.getElementById('setting-bulletin').value = f;
  });
  document.getElementById('browse-reference').addEventListener('click', async () => {
    const current = document.getElementById('setting-reference').value;
    const f = await window.dcc.pickFile({
      title: 'Select Reference PPTX',
      filters: [{ name: 'PowerPoint', extensions: ['pptx'] }],
      defaultPath: current || settings.outputPath
    });
    if (f) document.getElementById('setting-reference').value = f;
  });

  document.getElementById('add-mapping-btn').addEventListener('click', () => {
    const row = makeMappingRow();
    document.getElementById('mappings-body').appendChild(row);
    row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    row.querySelector('.mapping-title').focus();
  });

  document.getElementById('settings-save-btn').addEventListener('click', async () => {
    const newSettings = {
      songLibraryPath: document.getElementById('setting-library').value,
      outputPath: document.getElementById('setting-output').value,
      bulletinPath: document.getElementById('setting-bulletin').value,
      referencePptxPath: document.getElementById('setting-reference').value
    };
    await window.dcc.saveSettings(newSettings);
    await window.dcc.saveSongMappings(collectMappings());
    settings = await window.dcc.getSettings();
    showView('main-view');
  });

  document.getElementById('settings-cancel-btn').addEventListener('click', () => {
    showView(currentItems.length > 0 ? 'review-view' : 'main-view');
  });

  /* maintenance section hidden — convert-all-btn removed from HTML
  document.getElementById('convert-all-btn').addEventListener('click', async () => {
    document.getElementById('convert-all-btn').disabled = true;
    document.getElementById('convert-all-btn').textContent = 'Converting…';
    showView('progress-view');

    window.dcc.onProgress(({ step, total, message }) => {
      const pct = total > 0 ? Math.round((step / total) * 100) : 0;
      document.getElementById('progress-bar').style.width = pct + '%';
      document.getElementById('progress-msg').textContent = message || '';
    });

    try {
      const result = await window.dcc.convertAllPpt();
      window.dcc.removeProgressListeners();
      alert(`Converted ${result.converted} of ${result.total} files.\n` +
        (result.errors.length > 0 ? `\nErrors:\n${result.errors.map(e => e.file + ': ' + e.error).join('\n')}` : ''));
    } catch (e) {
      window.dcc.removeProgressListeners();
      alert('Error: ' + e.message);
    }

    document.getElementById('convert-all-btn').disabled = false;
    document.getElementById('convert-all-btn').textContent = 'Convert All .ppt Files';
    showView('settings-view');
  });
  */
}

// ── Menu listeners ─────────────────────────────────────────────────────────
function setupMenuListeners() {
  window.dcc.onMenuOpen(() => {
    document.getElementById('drop-zone').click();
  });
  window.dcc.onMenuSettings(() => {
    populateSettings();
    showView('settings-view');
  });
  window.dcc.onMenuConvertAll(() => {
    document.getElementById('convert-all-btn').click();
  });
}

// ── Bootstrap ────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', init);
