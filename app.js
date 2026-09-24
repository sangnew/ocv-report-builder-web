'use strict';

const C = window.OcvConvert;
const COLUMNS = C.COLUMNS;
const NUMERIC_KEYS = new Set(['docv', 'frozenIr', 'layer', 'docvV', 'x', 'y', 'longSide', 'shortSide', 'height']);

const state = {
  source: null,          // workbook loaded on this page: { fileName, cells, lots, sheetCount, masterName }
  selectedLots: new Set(),
  prev: null,            // previous report on this page: { fileName, byId }
  report: null,          // shared report: { sourceFile, lots, minDropV, builtAt, sheetName, rows: [...], excluded: [...] }
  entries: {},           // shared typed values: { CELLID: { key: value } }
  reportLoaded: false,
};

const $ = (id) => document.getElementById(id);
const hasOwn = (o, k) => Object.prototype.hasOwnProperty.call(o, k);

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
let toastTimer;
function showToast(msg, isError) {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}
function busy(text) {
  $('busyText').textContent = text || 'Working…';
  $('busy').hidden = !text;
}
function setStatus(el, html, isError) {
  el.hidden = false;
  el.className = 'status' + (isError ? ' error' : '');
  el.innerHTML = html;
}
function setLiveBadge(mode, text) {
  $('liveBadge').className = 'live-badge ' + mode;
  $('liveBadgeText').textContent = text;
}
function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---------------------------------------------------------------------------
// Worker (workbook parsing)
// ---------------------------------------------------------------------------
const worker = new Worker('worker.js?v=4');
let msgSeq = 0;
const pending = new Map();
worker.onmessage = (e) => {
  const p = pending.get(e.data.id);
  if (!p) return;
  pending.delete(e.data.id);
  e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error));
};
worker.onerror = (e) => {
  console.error(e);
  pending.forEach((p) => p.reject(new Error('The file reader failed to start. Check your internet connection and reload.')));
  pending.clear();
};
function callWorker(msg, transfer) {
  return new Promise((resolve, reject) => {
    const id = ++msgSeq;
    pending.set(id, { resolve, reject });
    worker.postMessage(Object.assign({ id }, msg), transfer || []);
  });
}

// ---------------------------------------------------------------------------
// Shared storage: Firestore (team) when configured, otherwise this browser
//   ocv_report/current            the report table (auto-filled values)
//   ocv_report_entries/{CELLID}   values typed by the team, { values: { key: value } }
// ---------------------------------------------------------------------------
const FB_CFG = window.BTT_FIREBASE_CONFIG;
const CLOUD = !!(FB_CFG && FB_CFG.apiKey && FB_CFG.projectId && window.firebase);
const BATCH_LIMIT = 450;

const CloudStore = {
  db: null,
  unsubs: [],
  init() {
    if (!firebase.apps.length) firebase.initializeApp(FB_CFG);
    this.db = firebase.firestore();
  },
  reportRef() { return this.db.collection('ocv_report').doc('current'); },
  entriesCol() { return this.db.collection('ocv_report_entries'); },
  subscribe(onReport, onEntries, onError) {
    this.unsubscribe();
    this.unsubs.push(this.reportRef().onSnapshot((d) => onReport(d.exists ? d.data() : null), onError));
    this.unsubs.push(this.entriesCol().onSnapshot((snap) => {
      const map = {};
      snap.docs.forEach((d) => { map[d.id] = (d.data() && d.data().values) || {}; });
      onEntries(map);
    }, onError));
  },
  unsubscribe() { this.unsubs.forEach((u) => u()); this.unsubs = []; },
  saveReport(report) { return this.reportRef().set(report); },
  updateReport(patch) { return this.reportRef().update(patch); },
  setExcluded(id, excluded) {
    const op = excluded ? firebase.firestore.FieldValue.arrayUnion(id) : firebase.firestore.FieldValue.arrayRemove(id);
    return this.reportRef().update({ excluded: op });
  },
  setEntry(id, key, value) {
    const ref = this.entriesCol().doc(id);
    if (value === null) return ref.set({ values: { [key]: firebase.firestore.FieldValue.delete() }, updatedAt: new Date().toISOString() }, { merge: true });
    return ref.set({ values: { [key]: value }, updatedAt: new Date().toISOString() }, { merge: true });
  },
  async deleteAll() {
    const snap = await this.entriesCol().get();
    for (let i = 0; i < snap.docs.length; i += BATCH_LIMIT) {
      const batch = this.db.batch();
      snap.docs.slice(i, i + BATCH_LIMIT).forEach((d) => batch.delete(d.ref));
      await batch.commit();
    }
    await this.reportRef().delete();
  },
};

const LOCAL_REPORT_KEY = 'ocvrb.report.v2';
const LOCAL_ENTRIES_KEY = 'ocvrb.entries.v2';
const LocalStore = {
  onReport: null,
  onEntries: null,
  read(key, fallback) { try { return JSON.parse(localStorage.getItem(key)) || fallback; } catch (e) { return fallback; } },
  write(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) {} },
  init() {},
  subscribe(onReport, onEntries) {
    this.onReport = onReport; this.onEntries = onEntries;
    onReport(this.read(LOCAL_REPORT_KEY, null));
    onEntries(this.read(LOCAL_ENTRIES_KEY, {}));
  },
  unsubscribe() {},
  async saveReport(report) { this.write(LOCAL_REPORT_KEY, report); this.onReport(report); },
  async updateReport(patch) { const r = Object.assign(this.read(LOCAL_REPORT_KEY, {}), patch); this.write(LOCAL_REPORT_KEY, r); this.onReport(r); },
  async setExcluded(id, excluded) {
    const r = this.read(LOCAL_REPORT_KEY, {});
    const set = new Set(r.excluded || []);
    excluded ? set.add(id) : set.delete(id);
    r.excluded = [...set];
    this.write(LOCAL_REPORT_KEY, r);
    this.onReport(r);
  },
  async setEntry(id, key, value) {
    const e = this.read(LOCAL_ENTRIES_KEY, {});
    e[id] = e[id] || {};
    if (value === null) delete e[id][key]; else e[id][key] = value;
    if (!Object.keys(e[id]).length) delete e[id];
    this.write(LOCAL_ENTRIES_KEY, e);
    this.onEntries(e);
  },
  async deleteAll() {
    try { localStorage.removeItem(LOCAL_REPORT_KEY); localStorage.removeItem(LOCAL_ENTRIES_KEY); } catch (e) {}
    this.onReport(null); this.onEntries({});
  },
};
const store = CLOUD ? CloudStore : LocalStore;

// ---------------------------------------------------------------------------
// Step 1: source workbook
// ---------------------------------------------------------------------------
async function loadSource(file) {
  if (!file) return;
  busy(`Reading "${file.name}"…`);
  try {
    const buf = await file.arrayBuffer();
    const res = await callWorker({ type: 'source', buf }, [buf]);
    if (!res.cells.length) throw new Error('No cell rows were found in the Master sheet.');
    state.source = { fileName: file.name, cells: res.cells, lots: res.lots, sheetCount: res.sheetCount, masterName: res.masterName };
    state.selectedLots.clear();
    const warn = res.missingHeaders.length
      ? `<br/>⚠ Columns not found in the Master sheet (they will be left for manual input): ${escapeHtml(res.missingHeaders.join(', '))}` : '';
    setStatus($('sourceStatus'), `Loaded <b>${escapeHtml(file.name)}</b> — ${res.cells.length} cells in "${escapeHtml(res.masterName)}", ${res.sheetCount} sheets.${warn}`, false);
    renderLots();
    $('step2').hidden = false;
  } catch (err) {
    console.error(err);
    setStatus($('sourceStatus'), `Could not read this file: ${escapeHtml(err.message)}`, true);
  } finally {
    busy(null);
  }
}

// ---------------------------------------------------------------------------
// Step 2: LOT selection and building
// ---------------------------------------------------------------------------
function selectedCells() {
  if (!state.source) return [];
  const onlyWithSheet = $('onlyWithSheet').checked;
  return state.source.cells.filter((c) => state.selectedLots.has(c.lot || '(no LOT)') && (!onlyWithSheet || c.trackingSheet));
}
function renderLots() {
  const grid = $('lotGrid');
  grid.innerHTML = '';
  const onlyWithSheet = $('onlyWithSheet').checked;
  state.source.lots.forEach((l) => {
    const count = onlyWithSheet ? l.withSheet : l.total;
    const el = document.createElement('label');
    el.className = 'lot' + (state.selectedLots.has(l.lot) ? ' on' : '') + (count ? '' : ' empty');
    el.innerHTML = `<input type="checkbox" ${state.selectedLots.has(l.lot) ? 'checked' : ''} /> <b>${escapeHtml(l.lot)}</b><small>${count} cell${count === 1 ? '' : 's'}</small>`;
    el.querySelector('input').onchange = (e) => {
      if (e.target.checked) state.selectedLots.add(l.lot); else state.selectedLots.delete(l.lot);
      el.classList.toggle('on', e.target.checked);
      updateSelectionHint();
    };
    grid.appendChild(el);
  });
  updateSelectionHint();
}
function updateSelectionHint() {
  const n = selectedCells().length;
  $('buildBtn').disabled = n === 0;
  $('selectionHint').textContent = state.selectedLots.size ? `${n} cell${n === 1 ? '' : 's'} selected` : 'Select one or more LOTs';
}

async function loadPrevReport(file) {
  if (!file) return;
  busy(`Reading "${file.name}"…`);
  try {
    const buf = await file.arrayBuffer();
    const res = await callWorker({ type: 'report', buf }, [buf]);
    if (!res.report) throw new Error('No "Cell ID" header row was found.');
    state.prev = { fileName: file.name, byId: res.report };
    $('prevStatus').textContent = `✓ ${file.name}: ${Object.keys(res.report).length} cells (sheet "${res.sheetName}") — used when you build the report`;
  } catch (err) {
    console.error(err);
    state.prev = null;
    $('prevStatus').textContent = `Could not read this report: ${err.message}`;
  } finally {
    busy(null);
  }
}

function defaultSheetName(lots) {
  const d = new Date();
  const stamp = String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
  const range = lots.length === 1 ? C.padLot(lots[0]) : `${C.padLot(lots[0])}~${C.padLot(lots[lots.length - 1])}`;
  return `${stamp} ${range}`.slice(0, 31);
}

async function buildReport() {
  const cells = selectedCells();
  if (!cells.length) return;
  if (state.report && state.report.rows && state.report.rows.length && CLOUD &&
      !confirm('Replace the shared report for the whole team? Values the team typed are kept (they are saved by Cell ID).')) return;
  const minDropMv = Number($('minDropInput').value);
  const minDropV = isFinite(minDropMv) && minDropMv >= 0 ? minDropMv / 1000 : C.DEFAULT_MIN_DROP_V;
  busy(`Analyzing ${cells.filter((c) => c.trackingSheet).length} OCV tracking sheets…`);
  try {
    const res = await callWorker({ type: 'tracking', sheets: cells.filter((c) => c.trackingSheet).map((c) => c.trackingSheet) });
    const prevById = state.prev ? state.prev.byId : {};
    const rows = cells.map((cell) => {
      const built = C.buildRow(cell, cell.trackingSheet ? res.analysis[cell.trackingSheet] : null, { minDropV });
      const prev = prevById[cell.cellId.toUpperCase()] || {};
      const auto = {}, src = {}, note = {};
      COLUMNS.forEach(({ key }) => {
        let f = built[key];
        if ((!f.value || C.MANUAL_ONLY.includes(key)) && prev[key]) f = { value: prev[key], source: 'prev', note: 'From ' + state.prev.fileName };
        auto[key] = f.value;
        src[key] = f.source;
        if (f.note) note[key] = f.note;
      });
      return { id: cell.cellId.toUpperCase(), auto, src, note };
    });
    const lots = [...new Set(cells.map((c) => c.lot))].sort((a, b) => C.lotSortKey(a).localeCompare(C.lotSortKey(b)));
    const report = {
      sourceFile: state.source.fileName,
      lots,
      minDropV,
      prevFile: state.prev ? state.prev.fileName : '',
      builtAt: new Date().toISOString(),
      sheetName: defaultSheetName(lots),
      rows,
      excluded: [],
    };
    busy('Saving the shared report…');
    await store.saveReport(report);
    $('step3').hidden = false;
    $('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
    showToast(CLOUD ? 'Report built and shared with the team.' : 'Report built.');
  } catch (err) {
    console.error(err);
    showToast('Could not build the report: ' + err.message, true);
  } finally {
    busy(null);
  }
}

// ---------------------------------------------------------------------------
// Step 3: table (auto values from the shared report + values typed by the team)
// ---------------------------------------------------------------------------
function field(row, key) {
  const typed = state.entries[row.id];
  if (typed && hasOwn(typed, key)) return { value: typed[key], source: typed[key] ? 'manual' : 'empty', note: 'Entered by the team' };
  return { value: row.auto[key] || '', source: row.src[key] || 'empty', note: row.note && row.note[key] };
}
function rowValues(row) {
  const v = {};
  COLUMNS.forEach(({ key }) => { v[key] = field(row, key).value; });
  return v;
}
function isIncluded(row) { return !(state.report.excluded || []).includes(row.id); }
function cellClass(f, required, key) {
  if (!f.value) return required.has(key) ? 'missing' : '';
  if (f.source === 'calc') return 'calc';
  if (f.source === 'manual' || f.source === 'prev') return 'entered';
  return '';
}
function missingCount(row) {
  const v = rowValues(row);
  return C.requiredKeys(v).filter((k) => !v[k]).length;
}
function sourceTitle(f) {
  const base = { master: 'From Master E & L', sheet: 'From the OCV tracking sheet', calc: 'Please verify', prev: 'From the previous report', manual: 'Entered by the team' }[f.source] || '';
  return [base, f.note && f.note !== base ? f.note : ''].filter(Boolean).join(' — ');
}

function renderTable() {
  const report = state.report;
  if (!report || !report.rows || !report.rows.length) { $('step3').hidden = true; return; }
  $('step3').hidden = false;
  $('reportMeta').textContent = `Shared report built from "${report.sourceFile}" · LOT ${report.lots.join(', ')} · Drop ≥ ${(report.minDropV * 1000).toFixed(1)} mV` +
    (report.prevFile ? ` · previous report "${report.prevFile}"` : '') + ` · ${new Date(report.builtAt).toLocaleString()}`;
  if (document.activeElement !== $('sheetNameInput')) $('sheetNameInput').value = report.sheetName || '';

  const active = document.activeElement;
  let focus = null;
  if (active && active.classList && active.classList.contains('cell')) {
    const tr = active.closest('tr');
    focus = { id: tr.dataset.id, key: active.parentElement.dataset.key, value: active.value, start: active.selectionStart, end: active.selectionEnd };
  }
  $('reportHead').innerHTML = '<th title="Include in download">✓</th><th>No.</th>' + COLUMNS.map((c) => `<th>${escapeHtml(c.label)}</th>`).join('');
  const body = $('reportBody');
  body.innerHTML = '';
  const onlyMissing = $('onlyMissing').checked;
  let no = 0;
  const frag = document.createDocumentFragment();
  report.rows.forEach((row) => {
    const included = isIncluded(row);
    if (included) no++;
    if (onlyMissing && (!included || missingCount(row) === 0)) return;
    frag.appendChild(createRow(row, included ? no : null, included));
  });
  body.appendChild(frag);
  if (focus) {
    const el = body.querySelector(`tr[data-id="${CSS.escape(focus.id)}"] td[data-key="${focus.key}"] input`);
    if (el) {
      el.value = focus.value;
      el.focus({ preventScroll: true });
      try { el.setSelectionRange(focus.start, focus.end); } catch (e) {}
    }
  }
  renderSummary();
}
function createRow(row, no, included) {
  const tr = document.createElement('tr');
  tr.dataset.id = row.id;
  if (!included) tr.className = 'excluded';
  const inc = document.createElement('td');
  inc.className = 'inc';
  inc.innerHTML = `<input type="checkbox" ${included ? 'checked' : ''} title="Include in download" />`;
  inc.querySelector('input').onchange = (e) => {
    store.setExcluded(row.id, !e.target.checked).catch((err) => { console.error(err); showToast('Could not save. Check your connection.', true); });
  };
  tr.appendChild(inc);
  const num = document.createElement('td');
  num.className = 'num';
  num.textContent = no == null ? '' : no;
  tr.appendChild(num);
  COLUMNS.forEach(({ key }) => {
    const td = document.createElement('td');
    td.dataset.key = key;
    const input = document.createElement('input');
    input.className = 'cell';
    input.addEventListener('change', () => onCellEdit(row, key, input.value));
    td.appendChild(input);
    tr.appendChild(td);
  });
  updateRowCells(row, tr);
  return tr;
}
// Refresh values and highlights of one rendered row (skips the input being edited)
function updateRowCells(row, tr) {
  const required = new Set(C.requiredKeys(rowValues(row)));
  tr.querySelectorAll('td[data-key]').forEach((td) => {
    const key = td.dataset.key;
    const f = field(row, key);
    const input = td.firstChild;
    if (document.activeElement !== input && input.value !== f.value) input.value = f.value;
    input.title = sourceTitle(f);
    td.className = [cellClass(f, required, key), key === 'cellId' ? 'w-id' : '', ['sem', 'location', 'frozenIr', 'voltageDrop'].includes(key) ? 'w-wide' : ''].filter(Boolean).join(' ');
  });
}
function refreshRenderedRows() {
  if (!state.report || !state.report.rows) return;
  const byId = new Map(state.report.rows.map((r) => [r.id, r]));
  $('reportBody').querySelectorAll('tr[data-id]').forEach((tr) => {
    const row = byId.get(tr.dataset.id);
    if (row) updateRowCells(row, tr);
  });
  renderSummary();
}
function onCellEdit(row, key, rawValue) {
  const value = rawValue.trim();
  const auto = row.auto[key] || '';
  // Typing the workbook value back removes the override; anything else (including blank) is saved
  const stored = value === auto && auto !== '' ? null : value;
  state.entries[row.id] = Object.assign({}, state.entries[row.id]);
  if (stored === null) delete state.entries[row.id][key]; else state.entries[row.id][key] = stored;
  refreshRenderedRows();
  store.setEntry(row.id, key, stored).catch((err) => { console.error(err); showToast('Could not save this value. Check your connection.', true); });
}
function renderSummary() {
  const included = state.report.rows.filter(isIncluded);
  let missingFields = 0, rowsMissing = 0, verify = 0;
  included.forEach((r) => {
    const m = missingCount(r);
    missingFields += m;
    if (m) rowsMissing++;
    COLUMNS.forEach(({ key }) => { const f = field(r, key); if (f.source === 'calc' && f.value) verify++; });
  });
  $('summaryText').innerHTML = `<b>${included.length}</b> row${included.length === 1 ? '' : 's'} · ` +
    (missingFields ? `<b>${rowsMissing}</b> row${rowsMissing === 1 ? '' : 's'} need input (<b>${missingFields}</b> highlighted cell${missingFields === 1 ? '' : 's'})` : 'nothing left to fill in ✓') +
    (verify ? ` · ${verify} value${verify === 1 ? '' : 's'} to verify` : '');
}

// ---------------------------------------------------------------------------
// Export (same layout as the report template: headers on row 2, data from row 3, No. in column B)
// ---------------------------------------------------------------------------
const FILL = {
  missing: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
  calc: { patternType: 'solid', fgColor: { rgb: 'DDEBF7' } },
  gray: { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } },
};
const THIN = { style: 'thin', color: { rgb: 'BFBFBF' } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };

function sanitizeSheetName(name) {
  const clean = String(name || '').replace(/[\[\]:*?\/\\]/g, '-').trim().slice(0, 31);
  return clean || 'Report';
}
function exportXlsx() {
  const rows = state.report.rows.filter(isIncluded);
  if (!rows.length) { showToast('No rows are selected for the download.', true); return; }
  const sheetName = sanitizeSheetName($('sheetNameInput').value);
  const ws = {};
  const put = (r, c, v, s) => {
    const cell = v === '' || v == null ? { t: 's', v: '' } : (typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) });
    if (s) cell.s = s;
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  };
  put(1, 1, '', { fill: FILL.gray, border: BORDER });
  COLUMNS.forEach((col, i) => put(1, i + 2, col.label, { font: { bold: true }, border: BORDER, alignment: { wrapText: true, vertical: 'center' } }));
  rows.forEach((row, ri) => {
    const r = ri + 2;
    const required = new Set(C.requiredKeys(rowValues(row)));
    put(r, 1, ri + 1, { border: BORDER });
    COLUMNS.forEach((col, i) => {
      const f = field(row, col.key);
      let v = f.value;
      if (v && NUMERIC_KEYS.has(col.key) && C.isNum(v)) v = Number(v);
      const style = { border: BORDER };
      if (!f.value && required.has(col.key)) style.fill = FILL.missing;
      else if (f.value && f.source === 'calc') style.fill = FILL.calc;
      put(r, i + 2, v, style);
    });
  });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 1, c: 1 }, e: { r: rows.length + 1, c: COLUMNS.length + 1 } });
  ws['!cols'] = [{ wch: 2 }, { wch: 5 }].concat(COLUMNS.map((c) => ({ wch: Math.max(8, Math.min(26, c.label.length + 2)) })));
  ws['!cols'][3] = { wch: 13 }; // Cell ID

  const report = state.report;
  const legend = XLSX.utils.aoa_to_sheet([
    ['Legend'],
    ['Yellow', 'Needs manual input (not found in the tracking workbook)'],
    ['Light blue', 'Voltage drop from the OCV tracking sheet disagrees with Master E & L — please verify'],
    ['Voltage drop', `Per layer dOCV = biggest fall between the tracking dates (C, D, E). Drop = a layer > ${C.SIGMA_LIMIT}σ above the others and ≥ ${(report.minDropV * 1000).toFixed(1)} mV; otherwise NTF.`],
    ['Frozen IR P/F', 'NG when Frozen IR < 35 MΩ, OK otherwise (when Master E & L has no result)'],
    [],
    ['Source file', report.sourceFile],
    ['Previous report', report.prevFile || '(none)'],
    ['Built', new Date(report.builtAt).toLocaleString()],
    ['Downloaded', new Date().toLocaleString()],
  ]);
  legend.A1.s = { font: { bold: true } };
  legend.A2.s = { fill: FILL.missing };
  legend.A3.s = { fill: FILL.calc };
  legend['!cols'] = [{ wch: 16 }, { wch: 110 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.utils.book_append_sheet(wb, legend, 'Legend');
  XLSX.writeFile(wb, `${sheetName}.xlsx`);
  const remaining = rows.reduce((n, r) => n + missingCount(r), 0);
  showToast(remaining ? `Downloaded. ${remaining} highlighted cell(s) still need input.` : 'Downloaded.');
}

// ---------------------------------------------------------------------------
// Reset
// ---------------------------------------------------------------------------
function resetScope() { return document.querySelector('input[name="resetScope"]:checked').value; }
function updateResetUI() {
  const all = resetScope() === 'all';
  $('resetAllConfirm').hidden = !all;
  $('resetConfirmBtn').disabled = all && $('resetConfirmInput').value.trim() !== 'RESET';
}
function unloadPage() {
  state.source = null;
  state.prev = null;
  state.selectedLots.clear();
  $('sourceInput').value = '';
  $('prevInput').value = '';
  $('prevStatus').textContent = '';
  $('sourceStatus').hidden = true;
  $('lotGrid').innerHTML = '';
  $('step2').hidden = true;
}
async function doReset() {
  const all = resetScope() === 'all';
  if (all && $('resetConfirmInput').value.trim() !== 'RESET') return;
  $('resetConfirmBtn').disabled = true;
  try {
    if (all) { busy('Deleting the shared report…'); await store.deleteAll(); }
    unloadPage();
    $('resetOverlay').classList.remove('open');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast(all ? 'The shared report and all typed values were deleted.' : 'Workbook unloaded from this page.');
  } catch (err) {
    console.error(err);
    showToast('Reset failed: ' + err.message, true);
  } finally {
    busy(null);
    updateResetUI();
  }
}

// ---------------------------------------------------------------------------
// Sign-in and live data
// ---------------------------------------------------------------------------
function onReport(report) {
  const rebuilt = !state.report || !report || state.report.builtAt !== report.builtAt;
  state.report = report;
  state.reportLoaded = true;
  if (!report) { $('step3').hidden = true; return; }
  if (rebuilt) { renderTable(); return; }
  // Only the row selection / sheet name changed
  if (document.activeElement !== $('sheetNameInput')) $('sheetNameInput').value = report.sheetName || '';
  renderTable();
}
function onEntries(map) {
  state.entries = map;
  if (state.report) refreshRenderedRows();
}
function onStoreError(err) {
  console.error(err);
  if (err && err.code === 'permission-denied') { setLiveBadge('error', 'No access'); showToast('Access denied. Sign in with the team password again.', true); }
  else { setLiveBadge('error', 'Sync error'); showToast('Lost connection to the shared database. Refresh the page.', true); }
}
function startSync() {
  store.subscribe((r) => { onReport(r); setLiveBadge('live', CLOUD ? 'Synced · shared' : 'Saved in this browser'); }, onEntries, onStoreError);
}

function initAuth() {
  if (!CLOUD) {
    $('step1Note').textContent = 'Shared database not configured — saved in this browser only';
    startSync();
    return;
  }
  try { store.init(); } catch (err) { console.error(err); setLiveBadge('error', 'Connection error'); return; }
  firebase.auth().onAuthStateChanged((user) => {
    if (user) {
      $('loginOverlay').classList.remove('open');
      $('loginPassword').value = '';
      $('signOutBtn').hidden = false;
      setLiveBadge('connecting', 'Connecting');
      startSync();
    } else {
      store.unsubscribe();
      state.report = null;
      state.entries = {};
      $('step3').hidden = true;
      $('signOutBtn').hidden = true;
      setLiveBadge('connecting', 'Sign-in required');
      $('loginError').textContent = '';
      $('loginOverlay').classList.add('open');
      setTimeout(() => $('loginPassword').focus(), 0);
    }
  });
  $('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const pw = $('loginPassword').value;
    if (!pw) { $('loginError').textContent = 'Enter the team password.'; return; }
    $('loginBtn').disabled = true;
    $('loginError').textContent = '';
    try {
      await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      await firebase.auth().signInWithEmailAndPassword(window.BTT_TEAM_EMAIL, pw);
    } catch (err) {
      console.error(err);
      const code = err && err.code;
      $('loginError').textContent = ['auth/wrong-password', 'auth/invalid-credential', 'auth/invalid-login-credentials'].includes(code) ? 'Incorrect password.'
        : code === 'auth/too-many-requests' ? 'Too many attempts. Wait a few minutes and try again.'
          : code === 'auth/network-request-failed' ? 'Network error. Check your connection.'
            : `Sign-in failed (${code || 'unknown error'}).`;
    } finally {
      $('loginBtn').disabled = false;
    }
  });
  $('signOutBtn').onclick = () => firebase.auth().signOut();
}

// ---------------------------------------------------------------------------
// Wiring
// ---------------------------------------------------------------------------
function init() {
  const drop = $('sourceDrop');
  $('sourceInput').addEventListener('change', (e) => loadSource(e.target.files[0]));
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('drag'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) loadSource(f); });

  $('onlyWithSheet').addEventListener('change', renderLots);
  $('lotsNoneBtn').onclick = () => { state.selectedLots.clear(); renderLots(); };
  $('prevInput').addEventListener('change', (e) => loadPrevReport(e.target.files[0]));
  $('buildBtn').onclick = buildReport;
  $('onlyMissing').addEventListener('change', renderTable);
  $('downloadBtn').onclick = exportXlsx;
  $('sheetNameInput').addEventListener('change', (e) => {
    store.updateReport({ sheetName: sanitizeSheetName(e.target.value) }).catch((err) => console.error(err));
  });

  $('resetBtn').onclick = () => {
    document.querySelector('input[name="resetScope"][value="page"]').checked = true;
    $('resetAllLabel').textContent = CLOUD ? 'Delete the shared report and every value the team typed (for everyone)' : 'Delete the saved report and every value you typed';
    $('resetConfirmInput').value = '';
    updateResetUI();
    $('resetOverlay').classList.add('open');
  };
  document.querySelectorAll('input[name="resetScope"]').forEach((r) => r.addEventListener('change', updateResetUI));
  $('resetConfirmInput').addEventListener('input', updateResetUI);
  $('resetCancelBtn').onclick = () => $('resetOverlay').classList.remove('open');
  $('resetOverlay').addEventListener('click', (e) => { if (e.target === $('resetOverlay')) $('resetOverlay').classList.remove('open'); });
  $('resetConfirmBtn').onclick = doReset;

  initAuth();
}
init();
