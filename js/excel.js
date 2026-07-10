/* ==========================================================================
   Hamleys Global Count Dashboard — excel.js (Admin Generator)
   Drag & drop Excel upload → SheetJS parse → validate → map → gc-data.json
   ========================================================================== */

'use strict';

const ADMIN = {
  columns: [],
  rows: [],
  storeMapping: [],
  generatedJSON: null,
  sourceFileName: '',
  mapColumns: [],
  mapRows: [],
  generatedMapJSON: null,
};

const REQUIRED_COLUMN_HINTS = [
  { key: 'storeCode', names: ['store code'], required: true },
  { key: 'storeName', names: ['store name'], required: true },
  { key: 'date', names: ['date of global count', 'date'], required: true },
  { key: 'gc', names: ["no. of sku's counted", 'no of skus counted', 'sku', 'global count'], required: true },
  { key: 'rm', names: ['rm'], required: false },
  { key: 'rom', names: ['rom'], required: false },
  { key: 'sd', names: ['sd'], required: false },
];

function $(sel) { return document.querySelector(sel); }
function $all(sel) { return Array.from(document.querySelectorAll(sel)); }

/* ---------------------------------------------------------------------- *
 * Theme (shared behaviour with main dashboard)
 * ---------------------------------------------------------------------- */
function initTheme() {
  const saved = localStorage.getItem('hamleys-gc-theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  $('#themeToggle').textContent = saved === 'dark' ? '☀️' : '🌙';
  $('#themeToggle').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('hamleys-gc-theme', next);
    $('#themeToggle').textContent = next === 'dark' ? '☀️' : '🌙';
  });
}

/* ---------------------------------------------------------------------- *
 * Column resolution helpers
 * ---------------------------------------------------------------------- */
function findColIndex(columns, candidates) {
  const lower = columns.map(c => (c || '').toString().toLowerCase().trim());
  for (const cand of candidates) {
    const idx = lower.findIndex(c => c === cand.toLowerCase());
    if (idx !== -1) return idx;
  }
  for (const cand of candidates) {
    const idx = lower.findIndex(c => c.includes(cand.toLowerCase()));
    if (idx !== -1) return idx;
  }
  return -1;
}

function excelSerialToDate(v) {
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
  return null;
}

/**
 * Format a Date parsed from Excel as "YYYY-MM-DD", safe against day-shift
 * bugs. SheetJS's cellDates:true values are meant to be plain calendar dates,
 * but formula-driven workbooks sometimes store a tiny fractional time
 * component alongside the serial (e.g. 18:29:50 instead of 00:00:00), which
 * can round to the wrong calendar day once rendered in a timezone ahead of
 * UTC (Hamleys India runs on IST, UTC+5:30). Anchoring at UTC-noon before
 * reading the calendar fields absorbs that drift either direction.
 */
function toLocalISODate(d) {
  const shifted = new Date(d.getTime() + 12 * 60 * 60 * 1000);
  const y = shifted.getUTCFullYear();
  const m = String(shifted.getUTCMonth() + 1).padStart(2, '0');
  const day = String(shifted.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/* ---------------------------------------------------------------------- *
 * Step UI helpers
 * ---------------------------------------------------------------------- */
function setStep(step, state, msg) {
  const el = document.querySelector(`.step-item[data-step="${step}"]`);
  if (!el) return;
  el.classList.remove('done', 'active', 'error');
  el.classList.add(state);
  el.querySelector('.step-msg').textContent = msg || '';
}

function showAlert(container, type, msg) {
  const el = document.createElement('div');
  el.className = `alert ${type}`;
  el.innerHTML = `<span>${type === 'success' ? '✅' : type === 'error' ? '⚠️' : 'ℹ️'}</span><span>${msg}</span>`;
  $(container).innerHTML = '';
  $(container).appendChild(el);
}

function setProgress(pct) {
  $('#progressTrack').classList.remove('hidden');
  $('#progressFill').style.width = pct + '%';
}

/* ---------------------------------------------------------------------- *
 * Generic Excel reader (dynamic — no hardcoded columns)
 * ---------------------------------------------------------------------- */
function readWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheetName];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
        // First row = headers, dynamic — no assumptions about count/order
        const columns = (aoa[0] || []).map(h => (h || '').toString().trim());
        const rows = aoa.slice(1)
          .filter(r => r.some(cell => cell !== '' && cell !== null && cell !== undefined))
          .map(r => columns.map((_, i) => {
            let v = r[i];
            if (v instanceof Date) return toLocalISODate(v);
            return v === undefined ? '' : v;
          }));
        resolve({ columns, rows, sheetName: firstSheetName });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

/* ---------------------------------------------------------------------- *
 * Validation
 * ---------------------------------------------------------------------- */
function validateColumns(columns) {
  const results = REQUIRED_COLUMN_HINTS.map(hint => {
    const idx = findColIndex(columns, hint.names);
    return { ...hint, idx, found: idx !== -1 };
  });
  const missingRequired = results.filter(r => r.required && !r.found);
  return { results, missingRequired, ok: missingRequired.length === 0 };
}

/* ---------------------------------------------------------------------- *
 * Mapping merge — fills RM / ROM / SD / Store Name using store-mapping.json
 * ---------------------------------------------------------------------- */
function mergeWithMapping(columns, rows, mapping) {
  const mapByCode = {};
  mapping.forEach(m => mapByCode[m.code] = m);

  let outColumns = [...columns];
  const idxStoreCode = findColIndex(outColumns, ['store code']);
  let idxRM = findColIndex(outColumns, ['rm']);
  let idxROM = findColIndex(outColumns, ['rom']);
  let idxSD = findColIndex(outColumns, ['sd']);
  let idxName = findColIndex(outColumns, ['store name']);

  // Append missing mapping columns dynamically if the sheet doesn't contain them at all
  if (idxRM === -1) { outColumns.push('RM'); idxRM = outColumns.length - 1; }
  if (idxROM === -1) { outColumns.push('ROM'); idxROM = outColumns.length - 1; }
  if (idxSD === -1) { outColumns.push('SD'); idxSD = outColumns.length - 1; }

  let filledCount = 0, unmatchedStores = new Set();

  const outRows = rows.map(row => {
    const r = [...row];
    while (r.length < outColumns.length) r.push('');
    const code = idxStoreCode !== -1 ? String(r[idxStoreCode] || '').trim() : '';
    const m = mapByCode[code];
    if (!m && code) unmatchedStores.add(code);

    if (m) {
      if (!r[idxRM] || String(r[idxRM]).trim() === '') { r[idxRM] = m.rm; filledCount++; }
      if (!r[idxROM] || String(r[idxROM]).trim() === '') { r[idxROM] = m.rom; filledCount++; }
      if (!r[idxSD] || String(r[idxSD]).trim() === '') { r[idxSD] = m.sd; filledCount++; }
      if (idxName !== -1 && (!r[idxName] || String(r[idxName]).trim() === '')) r[idxName] = m.name;
    }
    return r;
  });

  return { columns: outColumns, rows: outRows, filledCount, unmatchedStores: Array.from(unmatchedStores) };
}

/* ---------------------------------------------------------------------- *
 * Stats computation for preview cards
 * ---------------------------------------------------------------------- */
function computeStats(columns, rows, mapping) {
  const idxStoreCode = findColIndex(columns, ['store code']);
  const idxRM = findColIndex(columns, ['rm']);
  const idxROM = findColIndex(columns, ['rom']);
  const idxSD = findColIndex(columns, ['sd']);
  const idxDate = findColIndex(columns, ['date of global count', 'date']);
  const idxGC = findColIndex(columns, ["no. of sku's counted", 'no of skus counted', 'sku', 'global count']);

  const stores = new Set(), rms = new Set(), roms = new Set(), sds = new Set(), dates = new Set();
  let totalGC = 0;
  rows.forEach(r => {
    if (idxStoreCode !== -1 && r[idxStoreCode]) stores.add(String(r[idxStoreCode]).trim());
    if (idxRM !== -1 && r[idxRM]) rms.add(r[idxRM]);
    if (idxROM !== -1 && r[idxROM]) roms.add(r[idxROM]);
    if (idxSD !== -1 && r[idxSD]) sds.add(r[idxSD]);
    if (idxDate !== -1 && r[idxDate]) dates.add(String(r[idxDate]));
    if (idxGC !== -1) totalGC += Number(r[idxGC]) || 0;
  });

  return {
    totalRows: rows.length,
    totalColumns: columns.length,
    uniqueStores: stores.size,
    storesMasterCount: mapping.length,
    completion: mapping.length ? ((stores.size / mapping.length) * 100).toFixed(1) : '—',
    rmCount: rms.size, romCount: roms.size, sdCount: sds.size,
    dateRange: dates.size, totalGC,
  };
}

/* ---------------------------------------------------------------------- *
 * Main GC upload pipeline
 * ---------------------------------------------------------------------- */
async function processGCFile(file) {
  ADMIN.sourceFileName = file.name;
  $('#stepList').classList.remove('hidden');
  $('#alertsArea').innerHTML = '';
  $('#mappingStats').classList.add('hidden');
  $('#resultActions').classList.add('hidden');
  $('#previewWrap').classList.add('hidden');
  setProgress(5);

  try {
    // STEP 1: Read
    setStep('read', 'active', 'Reading…');
    const { columns, rows } = await readWorkbookFile(file);
    setStep('read', 'done', `${rows.length} rows found`);
    setProgress(25);

    // STEP 2: Parse (already parsed as part of read; simulate distinct stage)
    setStep('parse', 'active', 'Converting…');
    await wait(150);
    setStep('parse', 'done', `${columns.length} columns detected`);
    setProgress(45);

    // STEP 3: Validate
    setStep('validate', 'active', 'Checking…');
    const validation = validateColumns(columns);
    if (!validation.ok) {
      setStep('validate', 'error', 'Missing required columns');
      showAlert('#alertsArea', 'error', `Missing required column(s): <strong>${validation.missingRequired.map(m => m.names[0]).join(', ')}</strong>. Please check the uploaded file and try again.`);
      setProgress(0);
      return;
    }
    setStep('validate', 'done', 'All required columns present');
    setProgress(60);

    // STEP 4: Mapping merge
    setStep('map', 'active', 'Loading store mapping…');
    let mapping = ADMIN.storeMapping;
    if (!mapping || !mapping.length) {
      mapping = await fetchExistingMapping();
      ADMIN.storeMapping = mapping;
    }
    const merged = mergeWithMapping(columns, rows, mapping);
    ADMIN.columns = merged.columns;
    ADMIN.rows = merged.rows;
    setStep('map', 'done', merged.unmatchedStores.length ? `${merged.unmatchedStores.length} store(s) not in mapping` : 'All stores matched');
    setProgress(80);
    if (merged.unmatchedStores.length) {
      showAlert('#alertsArea', 'info', `${merged.unmatchedStores.length} store code(s) were not found in the mapping reference and could not be auto-filled: <strong>${merged.unmatchedStores.slice(0, 15).join(', ')}${merged.unmatchedStores.length > 15 ? '…' : ''}</strong>. Their RM/ROM/SD (if blank) will show as "Unassigned" on the dashboard.`);
    }

    // STEP 5: Build JSON
    setStep('build', 'active', 'Finalizing…');
    const stats = computeStats(ADMIN.columns, ADMIN.rows, mapping);
    const today = new Date();
    ADMIN.generatedJSON = {
      meta: {
        lastUpdated: today.toISOString().slice(0, 10),
        generatedAt: today.toISOString(),
        sourceFile: file.name,
        totalStoresMaster: mapping.length,
      },
      columns: ADMIN.columns,
      rows: ADMIN.rows,
      storeMaster: mapping,
    };
    setStep('build', 'done', 'gc-data.json ready');
    setProgress(100);

    renderStats(stats);
    renderPreview(ADMIN.columns, ADMIN.rows);
    $('#resultActions').classList.remove('hidden');
    if (!validation.missingRequired.length) {
      showAlert('#alertsArea', 'success', `Excel converted successfully — ${stats.totalRows} rows across ${stats.totalColumns} columns, ${stats.uniqueStores}/${stats.storesMasterCount} stores counted. Click "Download Updated gc-data.json" and replace the file at <code>data/gc-data.json</code> in your GitHub repo (or use Auto-Publish below).`);
    }
  } catch (err) {
    console.error(err);
    showAlert('#alertsArea', 'error', `Something went wrong: ${err.message}`);
    setProgress(0);
  }
}

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

async function fetchExistingMapping() {
  try {
    const res = await fetch('data/store-mapping.json?t=' + Date.now());
    if (!res.ok) return [];
    return await res.json();
  } catch { return []; }
}

function renderStats(stats) {
  const items = [
    { label: 'Rows Parsed', value: stats.totalRows },
    { label: 'Columns', value: stats.totalColumns },
    { label: 'Stores Uploaded', value: stats.uniqueStores },
    { label: 'Total Stores', value: stats.storesMasterCount },
    { label: 'Completion %', value: stats.completion + '%' },
    { label: 'RM Count', value: stats.rmCount },
    { label: 'ROM Count', value: stats.romCount },
    { label: 'SD Count', value: stats.sdCount },
    { label: 'Dates Covered', value: stats.dateRange },
    { label: 'Total Global Count', value: stats.totalGC },
  ];
  $('#mappingStats').classList.remove('hidden');
  $('#mappingStats').innerHTML = items.map(i => `<div class="mapping-stat"><div class="num">${i.value}</div><div class="lbl">${i.label}</div></div>`).join('');
}

function renderPreview(columns, rows) {
  $('#previewWrap').classList.remove('hidden');
  $('#previewHead').innerHTML = columns.map(c => `<th>${c}</th>`).join('');
  $('#previewBody').innerHTML = rows.slice(0, 15).map(r => `<tr>${r.map(c => `<td>${c ?? ''}</td>`).join('')}</tr>`).join('');
}

/* ---------------------------------------------------------------------- *
 * Store Mapping upload pipeline (Section 3)
 * ---------------------------------------------------------------------- */
async function processMapFile(file) {
  $('#mapAlertsArea').innerHTML = '';
  $('#mapResultActions').classList.add('hidden');
  try {
    const { columns, rows } = await readWorkbookFile(file);
    const idxCode = findColIndex(columns, ['store code']);
    const idxName = findColIndex(columns, ['name', 'store name']);
    const idxRM = findColIndex(columns, ['rm']);
    const idxSD = findColIndex(columns, ['sd/rom', 'sd']);
    const idxROM = findColIndex(columns, ['rom']);

    if (idxCode === -1 || idxRM === -1) {
      showAlert('#mapAlertsArea', 'error', 'Could not find "Store Code" and "RM" columns in the uploaded mapping file.');
      return;
    }

    const mapping = rows
      .filter(r => r[idxCode])
      .map(r => ({
        code: String(r[idxCode]).trim(),
        name: (r[idxName] ?? '').toString().trim(),
        rm: (r[idxRM] ?? '').toString().trim(),
        sd: (r[idxSD] ?? '').toString().trim(),
        rom: (r[idxROM] ?? '').toString().trim(),
      }));

    ADMIN.storeMapping = mapping;
    ADMIN.generatedMapJSON = mapping;
    showAlert('#mapAlertsArea', 'success', `Mapping parsed successfully — ${mapping.length} stores loaded. Download and replace <code>data/store-mapping.json</code>, then re-generate today's gc-data.json above so it picks up the latest mapping.`);
    $('#mapResultActions').classList.remove('hidden');
  } catch (err) {
    showAlert('#mapAlertsArea', 'error', `Failed to parse mapping file: ${err.message}`);
  }
}

/* ---------------------------------------------------------------------- *
 * Downloads
 * ---------------------------------------------------------------------- */
function downloadJSONFile(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 1)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------------------------------------------------------------- *
 * GitHub Auto-Publish (optional) — Contents API
 * ---------------------------------------------------------------------- */
function loadGithubSettings() {
  const s = JSON.parse(localStorage.getItem('hamleys-gh-settings') || '{}');
  $('#ghOwner').value = s.owner || '';
  $('#ghRepo').value = s.repo || '';
  $('#ghBranch').value = s.branch || '';
  $('#ghToken').value = s.token || '';
}

function saveGithubSettings() {
  const s = {
    owner: $('#ghOwner').value.trim(),
    repo: $('#ghRepo').value.trim(),
    branch: $('#ghBranch').value.trim() || 'main',
    token: $('#ghToken').value.trim(),
  };
  localStorage.setItem('hamleys-gh-settings', JSON.stringify(s));
  showAlert('#alertsArea', 'success', 'GitHub settings saved locally in this browser.');
}

async function publishToGithub() {
  const s = JSON.parse(localStorage.getItem('hamleys-gh-settings') || '{}');
  if (!s.owner || !s.repo || !s.token) {
    showAlert('#alertsArea', 'error', 'Please fill in and save your GitHub owner, repo and token first (Section 2).');
    return;
  }
  if (!ADMIN.generatedJSON) {
    showAlert('#alertsArea', 'error', 'Generate the JSON first by uploading today\'s Excel file.');
    return;
  }
  const path = 'data/gc-data.json';
  const branch = s.branch || 'main';
  const apiBase = `https://api.github.com/repos/${s.owner}/${s.repo}/contents/${path}`;
  try {
    showAlert('#alertsArea', 'info', 'Publishing to GitHub…');
    let sha;
    const getRes = await fetch(`${apiBase}?ref=${branch}`, { headers: { Authorization: `token ${s.token}` } });
    if (getRes.ok) { const j = await getRes.json(); sha = j.sha; }

    const content = b64EncodeUnicode(JSON.stringify(ADMIN.generatedJSON, null, 1));
    const putRes = await fetch(apiBase, {
      method: 'PUT',
      headers: { Authorization: `token ${s.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: `Update Global Count data — ${new Date().toISOString().slice(0, 10)}`,
        content, branch, ...(sha ? { sha } : {}),
      })
    });
    if (!putRes.ok) {
      const err = await putRes.json().catch(() => ({}));
      throw new Error(err.message || `GitHub API error (${putRes.status})`);
    }
    showAlert('#alertsArea', 'success', '✅ Published! data/gc-data.json has been committed to your repository. The live dashboard will pick it up automatically on next load.');
  } catch (err) {
    showAlert('#alertsArea', 'error', `GitHub publish failed: ${err.message}`);
  }
}

function b64EncodeUnicode(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

/* ---------------------------------------------------------------------- *
 * Drag & Drop wiring
 * ---------------------------------------------------------------------- */
function wireDropzone(zoneSel, inputSel, browseBtnSel, handler) {
  const zone = $(zoneSel), input = $(inputSel), browseBtn = $(browseBtnSel);
  zone.addEventListener('click', (e) => { if (e.target !== browseBtn) input.click(); });
  if (browseBtn) browseBtn.addEventListener('click', (e) => { e.stopPropagation(); input.click(); });
  input.addEventListener('change', () => { if (input.files[0]) handler(input.files[0]); });
  ['dragenter', 'dragover'].forEach(evt => zone.addEventListener(evt, (e) => { e.preventDefault(); zone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(evt => zone.addEventListener(evt, (e) => { e.preventDefault(); zone.classList.remove('dragover'); }));
  zone.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) handler(f); });
}

/* ---------------------------------------------------------------------- *
 * Init
 * ---------------------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  loadGithubSettings();

  wireDropzone('#gcDropzone', '#gcFileInput', '#browseBtn', processGCFile);
  wireDropzone('#mapDropzone', '#mapFileInput', '#mapBrowseBtn', processMapFile);

  $('#downloadJsonBtn').addEventListener('click', () => downloadJSONFile(ADMIN.generatedJSON, 'gc-data.json'));
  $('#downloadMapJsonBtn').addEventListener('click', () => downloadJSONFile(ADMIN.generatedMapJSON, 'store-mapping.json'));
  $('#previewJsonBtn').addEventListener('click', () => {
    const win = window.open('', '_blank');
    win.document.write(`<pre style="white-space:pre-wrap;font-family:monospace;padding:16px;">${JSON.stringify(ADMIN.generatedJSON, null, 2).replace(/</g, '&lt;')}</pre>`);
  });
  $('#ghSaveBtn').addEventListener('click', saveGithubSettings);
  $('#publishGithubBtn').addEventListener('click', publishToGithub);

  // Preload existing mapping reference so GC upload works standalone
  fetchExistingMapping().then(m => { ADMIN.storeMapping = m; });
});
