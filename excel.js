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

/**
 * Convert a raw Excel date serial (e.g. from a leftover numeric cell) into a
 * Date built via the LOCAL constructor from its literal Y/M/D/H/M/S fields —
 * see toLocalISODate() below for why this matters.
 */
function excelSerialToDate(v) {
  if (typeof v !== 'number') return null;
  const u = new Date(Math.round((v - 25569) * 86400 * 1000)); // literal epoch-day math, UTC fields = literal values
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
}

/**
 * Format a Date parsed from Excel as "YYYY-MM-DD".
 *
 * IMPORTANT: SheetJS's cellDates:true builds these Date objects using the
 * LOCAL Date constructor from the literal Y/M/D/H/M/S it decodes out of the
 * Excel serial (i.e. `new Date(y, m, d, h, mi, s)`), NOT via UTC. That means
 * `.getUTCHours()` etc. return the literal time shifted by whatever the
 * runtime's OS/browser timezone happens to be — e.g. under IST (UTC+5:30) an
 * entry typed as "18:25:47" comes back as "12:55:47" via the UTC getters.
 * A previous version of this function tried to compensate with a blanket
 * "+12 hours, then read UTC fields" shift, but that only masked the symptom
 * for near-midnight entries and actively corrupted every afternoon/evening
 * entry (the vast majority of real submissions here) by rolling them to the
 * next calendar day.
 *
 * The correct, timezone-independent fix: use the LOCAL getters
 * (getFullYear/getMonth/getDate), which — because of how the Date was
 * constructed above — always recover the exact literal values as typed in
 * Excel, regardless of what timezone the browser/OS is set to.
 */
function toLocalISODate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
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
/**
 * Some workbooks (especially ones generated/edited by external tools,
 * macros, or repeated exports) store a stale `!ref` dimension on the sheet
 * that undercounts the actually-populated cells. SheetJS's sheet_to_json()
 * trusts that declared range, so a stale `!ref` silently truncates rows —
 * this is the most common reason an upload "only takes" some fixed number
 * of rows even though the file clearly has more data below.
 *
 * This scans every real cell address on the sheet and returns the true
 * min/max row & column actually used, so we can widen `!ref` before
 * converting and guarantee no populated row/column is ever dropped.
 */
function computeActualUsedRange(ws) {
  let minR = null, minC = null, maxR = null, maxC = null;
  for (const addr in ws) {
    if (addr[0] === '!') continue; // skip meta keys like !ref, !merges, !cols
    const cell = XLSX.utils.decode_cell(addr);
    if (minR === null || cell.r < minR) minR = cell.r;
    if (maxR === null || cell.r > maxR) maxR = cell.r;
    if (minC === null || cell.c < minC) minC = cell.c;
    if (maxC === null || cell.c > maxC) maxC = cell.c;
  }
  if (minR === null) return null;
  return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
}

function readWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheetName];

        // Widen the sheet's declared range to the true used range so no row
        // is ever silently dropped because of a stale/undersized `!ref`.
        const declared = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
        const actual = computeActualUsedRange(ws);
        if (actual) {
          const widened = {
            s: {
              r: declared ? Math.min(declared.s.r, actual.s.r) : actual.s.r,
              c: declared ? Math.min(declared.s.c, actual.s.c) : actual.s.c,
            },
            e: {
              r: declared ? Math.max(declared.e.r, actual.e.r) : actual.e.r,
              c: declared ? Math.max(declared.e.c, actual.e.c) : actual.e.c,
            },
          };
          ws['!ref'] = XLSX.utils.encode_range(widened);
        }

        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', range: ws['!ref'] });
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
 * Mapping merge — RM / ROM / SD / Store Name always come from the
 * canonical store-mapping.json when the store code is found there.
 *
 * Why override instead of just filling blanks: the daily GC form lets
 * whoever fills it type RM/ROM/SD freely, and those values can drift from
 * the official mapping (typos, a store temporarily counted by someone
 * else, stale copy-paste from a previous day, etc). If a store's SD in the
 * upload doesn't match its SD in store-mapping.json, the dashboard's
 * "Total Stores" (computed from the mapping) and "GC Stores" (computed
 * from the upload's own SD column) stop tallying — you can end up with
 * completion percentages over 100% for one SD and under for another.
 * Forcing the canonical mapping's RM/ROM/SD onto every matched row
 * guarantees both numbers are always computed from the same source, so
 * they can never disagree.
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

  let overriddenCount = 0, unmatchedStores = new Set();

  const outRows = rows.map(row => {
    const r = [...row];
    while (r.length < outColumns.length) r.push('');
    const code = idxStoreCode !== -1 ? String(r[idxStoreCode] || '').trim() : '';
    const m = mapByCode[code];
    if (!m && code) unmatchedStores.add(code);

    if (m) {
      if (String(r[idxRM] || '').trim() !== m.rm) overriddenCount++;
      if (String(r[idxROM] || '').trim() !== m.rom) overriddenCount++;
      if (String(r[idxSD] || '').trim() !== m.sd) overriddenCount++;
      r[idxRM] = m.rm;
      r[idxROM] = m.rom;
      r[idxSD] = m.sd;
      if (idxName !== -1 && (!r[idxName] || String(r[idxName]).trim() === '')) r[idxName] = m.name;
    }
    return r;
  });

  return { columns: outColumns, rows: outRows, filledCount: overriddenCount, unmatchedStores: Array.from(unmatchedStores) };
}

/* ---------------------------------------------------------------------- *
 * Merge with already-published data — lets a daily upload contain only
 * the NEW day's rows instead of the full history every time.
 * ---------------------------------------------------------------------- */
async function fetchExistingGCData() {
  try {
    const res = await fetch('data/gc-data.json?t=' + Date.now());
    if (!res.ok) return null;
    return await res.json();
  } catch { return null; }
}

function rowsToObjects(columns, rows) {
  return rows.map(r => {
    const obj = {};
    columns.forEach((c, i) => { obj[c] = r[i] ?? ''; });
    return obj;
  });
}

function objectsToRows(columns, objects) {
  return objects.map(o => columns.map(c => (o[c] === undefined ? '' : o[c])));
}

/** Normalize a date-ish cell value (ISO string, free-text, or leftover Excel
 *  serial number) to "YYYY-MM-DD" purely for matching purposes, so the merge
 *  key is stable even if existing data stored dates slightly differently. */
function normalizeDateKey(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') {
    const d = excelSerialToDate(v);
    return d ? toLocalISODate(d) : String(v);
  }
  const s = String(v).trim();
  const iso = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return iso ? iso[1] : s;
}

/**
 * Merge newly-uploaded rows into the existing published dataset. Rows are
 * matched by "Store Code + Date" — if that combination already exists in
 * the published data, the new upload's row REPLACES it (so re-uploading a
 * day corrects it instead of duplicating); otherwise it's appended as a new
 * row. Everything already published for other dates/stores is left as-is.
 */
function mergeWithExistingData(existing, newColumns, newRows) {
  if (!existing || !Array.isArray(existing.columns) || !Array.isArray(existing.rows) || !existing.rows.length) {
    return { columns: newColumns, rows: newRows, addedCount: newRows.length, updatedCount: 0, totalCount: newRows.length, merged: false };
  }

  // Union of columns — existing layout first, then any genuinely new ones
  const combinedColumns = [...existing.columns];
  newColumns.forEach(c => { if (!combinedColumns.includes(c)) combinedColumns.push(c); });

  const codeIdx = findColIndex(combinedColumns, ['store code']);
  const dateIdx = findColIndex(combinedColumns, ['date of global count', 'date']);
  const codeCol = codeIdx !== -1 ? combinedColumns[codeIdx] : null;
  const dateCol = dateIdx !== -1 ? combinedColumns[dateIdx] : null;

  function keyOf(obj) {
    const code = codeCol ? String(obj[codeCol] || '').trim().toLowerCase() : '';
    const date = dateCol ? normalizeDateKey(obj[dateCol]) : '';
    return `${code}__${date}`;
  }

  const existingObjs = rowsToObjects(existing.columns, existing.rows);
  const newObjs = rowsToObjects(newColumns, newRows);

  const byKey = new Map();
  existingObjs.forEach(o => byKey.set(keyOf(o), o));

  let addedCount = 0, updatedCount = 0;
  newObjs.forEach(o => {
    const k = keyOf(o);
    if (byKey.has(k)) updatedCount++; else addedCount++;
    byKey.set(k, o); // new upload wins on conflict
  });

  const mergedRows = objectsToRows(combinedColumns, Array.from(byKey.values()));
  return { columns: combinedColumns, rows: mergedRows, addedCount, updatedCount, totalCount: mergedRows.length, merged: true };
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
    completion: mapping.length ? Math.round((stores.size / mapping.length) * 100) : '—',
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

    // STEP 4: Mapping merge (fills RM/ROM/SD for this upload's own rows)
    setStep('map', 'active', 'Loading store mapping…');
    let mapping = ADMIN.storeMapping;
    if (!mapping || !mapping.length) {
      mapping = await fetchExistingMapping();
      ADMIN.storeMapping = mapping;
    }
    const merged = mergeWithMapping(columns, rows, mapping);
    setStep('map', 'done', merged.unmatchedStores.length ? `${merged.unmatchedStores.length} store(s) not in mapping` : 'All stores matched');
    setProgress(70);
    if (merged.unmatchedStores.length) {
      showAlert('#alertsArea', 'info', `${merged.unmatchedStores.length} store code(s) were not found in the mapping reference and could not be auto-filled: <strong>${merged.unmatchedStores.slice(0, 15).join(', ')}${merged.unmatchedStores.length > 15 ? '…' : ''}</strong>. Their RM/ROM/SD (if blank) will show as "Unassigned" on the dashboard.`);
    }

    // STEP 5: Merge with already-published data (so a daily upload can
    // contain just the new day's rows — everything already published stays)
    const shouldMerge = $('#mergeWithExistingChk') ? $('#mergeWithExistingChk').checked : true;
    let mergeResult;
    if (shouldMerge) {
      setStep('merge', 'active', 'Fetching published gc-data.json…');
      const existing = await fetchExistingGCData();
      mergeResult = mergeWithExistingData(existing, merged.columns, merged.rows);
      if (mergeResult.merged) {
        setStep('merge', 'done', `+${mergeResult.addedCount} new, ${mergeResult.updatedCount} updated — ${mergeResult.totalCount} total rows`);
        if (mergeResult.updatedCount) {
          showAlert('#alertsArea', 'info', `${mergeResult.updatedCount} row(s) matched an already-published Store Code + Date and were replaced with this upload's values (corrections). ${mergeResult.addedCount} brand-new row(s) were appended. Nothing else already published was touched.`);
        }
      } else {
        setStep('merge', 'done', 'No published data.json found — this upload becomes the full dataset');
      }
    } else {
      setStep('merge', 'done', 'Merge skipped — replacing with this file only');
      mergeResult = { columns: merged.columns, rows: merged.rows, addedCount: merged.rows.length, updatedCount: 0, totalCount: merged.rows.length, merged: false };
    }
    ADMIN.columns = mergeResult.columns;
    ADMIN.rows = mergeResult.rows;
    setProgress(85);

    // STEP 6: Build JSON
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
      const mergeNote = shouldMerge && mergeResult.merged
        ? ` (${mergeResult.addedCount} new row(s) added, ${mergeResult.updatedCount} corrected, merged with what was already published)`
        : '';
      showAlert('#alertsArea', 'success', `Excel converted successfully — ${stats.totalRows} total rows across ${stats.totalColumns} columns${mergeNote}, ${stats.uniqueStores}/${stats.storesMasterCount} stores counted. Click "Download Updated gc-data.json" and replace the file at <code>data/gc-data.json</code> in your GitHub repo (or use Auto-Publish below).`);
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
  // Fetches the file's current SHA with cache-busting, so we never PUT against a stale version.
  async function fetchFreshSha() {
    const getRes = await fetch(`${apiBase}?ref=${branch}&_=${Date.now()}`, {
      headers: { Authorization: `token ${s.token}`, 'Cache-Control': 'no-cache' },
      cache: 'no-store',
    });
    if (getRes.ok) { const j = await getRes.json(); return j.sha; }
    return undefined;
  }

  const content = b64EncodeUnicode(JSON.stringify(ADMIN.generatedJSON, null, 1));

  async function attemptPublish(sha) {
    return fetch(apiBase, {
      method: 'PUT',
      headers: { Authorization: `token ${s.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: `Update Global Count data — ${new Date().toISOString().slice(0, 10)}`,
        content, branch, ...(sha ? { sha } : {}),
      })
    });
  }

  try {
    showAlert('#alertsArea', 'info', 'Publishing to GitHub…');
    let sha = await fetchFreshSha();
    let putRes = await attemptPublish(sha);

    // Classic race condition: SHA went stale between our GET and PUT. Refetch once and retry automatically.
    if (!putRes.ok) {
      const errBody = await putRes.json().catch(() => ({}));
      const isShaMismatch = putRes.status === 409 || /does not match|sha/i.test(errBody.message || '');
      if (isShaMismatch) {
        sha = await fetchFreshSha();
        putRes = await attemptPublish(sha);
      } else {
        throw new Error(errBody.message || `GitHub API error (${putRes.status})`);
      }
    }

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
