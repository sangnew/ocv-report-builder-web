'use strict';

const C = window.OcvConvert;
const COLUMNS = C.COLUMNS;
const NUMERIC_KEYS = new Set(['docv', 'frozenIr', 'layer', 'docvV', 'x', 'y', 'longSide', 'shortSide', 'height']);
const MANUAL_STORE_KEY = 'ocvrb.manualEntries.v1';

const state = {
  source: null,        // { fileName, cells, lots, sheetCount, masterName }
  selectedLots: new Set(),
  prev: null,          // { fileName, byId }
  rows: [],            // [{ cell, include, fields: { key: { value, source, auto } } }]
};

const $ = (id) => document.getElementById(id);

// ---------------------------------------------------------------------------
// Worker (workbook parsing)
// ---------------------------------------------------------------------------
const worker = new Worker('worker.js?v=3');
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
  el.className = 'status' + (el.classList.contains('inline') ? ' inline' : '') + (isError ? ' error' : '');
  el.innerHTML = html;
}
function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Values typed by hand, remembered per Cell ID in this browser
function loadManual() {
  try { return JSON.parse(localStorage.getItem(MANUAL_STORE_KEY) || '{}'); } catch (e) { return {}; }
}
function saveManual(store) {
  try { localStorage.setItem(MANUAL_STORE_KEY, JSON.stringify(store)); } catch (e) {}
}

// ---------------------------------------------------------------------------
// Step 1: source workbook
// ---------------------------------------------------------------------------
async function loadSource(file) {
  if (!file) return;
  busy(`Reading "${file.name}"… large workbooks take a few seconds.`);
  try {
    const buf = await file.arrayBuffer();
    const res = await callWorker({ type: 'source', buf }, [buf]);
    if (!res.cells.length) throw new Error('No cell rows were found in the Master sheet.');
    state.source = { fileName: file.name, cells: res.cells, lots: res.lots, sheetCount: res.sheetCount, masterName: res.masterName };
    state.selectedLots.clear();
    state.rows = [];
    $('step3').hidden = true;
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
// Step 2: LOT selection
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
    $('prevStatus').textContent = `✓ ${file.name}: ${Object.keys(res.report).length} cells (sheet "${res.sheetName}")`;
    if (state.rows.length) await buildReport();
  } catch (err) {
    console.error(err);
    state.prev = null;
    $('prevStatus').textContent = `Could not read this report: ${err.message}`;
  } finally {
    busy(null);
  }
}

// ---------------------------------------------------------------------------
// Build the report rows
// ---------------------------------------------------------------------------
async function buildReport() {
  const cells = selectedCells();
  if (!cells.length) return;
  busy('Building the report…');
  try {
    // Tracking sheets are only a fallback when the Master row has no voltage-drop data
    const needTracking = cells.filter((c) => c.trackingSheet && !c.anodeSheet && !c.voltageDrop && String(c.ntf).toUpperCase() !== 'NTF');
    const res = await callWorker({ type: 'tracking', sheets: needTracking.map((c) => c.trackingSheet) });
    const manual = loadManual();
    const prevById = state.prev ? state.prev.byId : {};
    state.rows = cells.map((cell) => {
      const built = C.buildRow(cell, cell.trackingSheet ? res.tracking[cell.trackingSheet] : null);
      const idKey = cell.cellId.toUpperCase();
      const prev = prevById[idKey] || {};
      const saved = manual[idKey] || {};
      const fields = {};
      COLUMNS.forEach(({ key }) => {
        const auto = built[key];
        let f = { value: auto.value, source: auto.source, auto: auto.value };
        if ((!f.value || C.MANUAL_ONLY.includes(key)) && prev[key]) f = { value: prev[key], source: 'prev', auto: auto.value };
        if (Object.prototype.hasOwnProperty.call(saved, key)) f = { value: saved[key], source: saved[key] ? 'manual' : 'empty', auto: auto.value };
        fields[key] = f;
      });
      return { cell, include: true, fields };
    });
    const lots = [...new Set(cells.map((c) => c.lot))].sort((a, b) => C.lotSortKey(a).localeCompare(C.lotSortKey(b)));
    $('sheetNameInput').value = defaultSheetName(lots);
    renderTable();
    $('step3').hidden = false;
    $('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    console.error(err);
    showToast('Could not build the report: ' + err.message, true);
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

// ---------------------------------------------------------------------------
// Step 3: table
// ---------------------------------------------------------------------------
function rowValues(row) {
  const v = {};
  COLUMNS.forEach(({ key }) => { v[key] = row.fields[key].value; });
  return v;
}
function cellClass(row, key, required) {
  const f = row.fields[key];
  if (!f.value) return required.has(key) ? 'missing' : '';
  if (f.source === 'calc') return 'calc';
  if (f.source === 'manual' || f.source === 'prev') return 'entered';
  return '';
}
function missingCount(row) {
  const v = rowValues(row);
  return C.requiredKeys(v).filter((k) => !v[k]).length;
}
function renderTable() {
  const head = $('reportHead');
  head.innerHTML = '<th title="Include in download">✓</th><th>No.</th>' + COLUMNS.map((c) => `<th>${escapeHtml(c.label)}</th>`).join('');
  const body = $('reportBody');
  body.innerHTML = '';
  const onlyMissing = $('onlyMissing').checked;
  let no = 0;
  const frag = document.createDocumentFragment();
  state.rows.forEach((row, idx) => {
    if (row.include) no++;
    if (onlyMissing && (!row.include || missingCount(row) === 0)) return;
    frag.appendChild(createRow(row, idx, row.include ? no : null));
  });
  body.appendChild(frag);
  renderSummary();
}
function createRow(row, idx, no) {
  const tr = document.createElement('tr');
  tr.dataset.idx = idx;
  if (!row.include) tr.className = 'excluded';
  const required = new Set(C.requiredKeys(rowValues(row)));
  const inc = document.createElement('td');
  inc.className = 'inc';
  inc.innerHTML = `<input type="checkbox" ${row.include ? 'checked' : ''} title="Include in download" />`;
  inc.querySelector('input').onchange = (e) => { row.include = e.target.checked; renderTable(); };
  tr.appendChild(inc);
  const num = document.createElement('td');
  num.className = 'num';
  num.textContent = no == null ? '' : no;
  tr.appendChild(num);
  COLUMNS.forEach(({ key }) => {
    const td = document.createElement('td');
    td.dataset.key = key;
    td.className = cellClass(row, key, required) + (key === 'cellId' ? ' w-id' : '') + (['sem', 'location', 'frozenIr', 'voltageDrop'].includes(key) ? ' w-wide' : '');
    const input = document.createElement('input');
    input.className = 'cell';
    input.value = row.fields[key].value;
    const f = row.fields[key];
    input.title = f.source === 'calc' ? 'From the tracking sheet — please verify'
      : f.source === 'prev' ? 'Copied from the previous report'
        : f.source === 'manual' ? 'Entered by you' : '';
    input.addEventListener('change', () => onCellEdit(row, key, input.value, tr));
    td.appendChild(input);
    tr.appendChild(td);
  });
  return tr;
}
function onCellEdit(row, key, rawValue, tr) {
  const value = rawValue.trim();
  const f = row.fields[key];
  const manual = loadManual();
  const idKey = row.cell.cellId.toUpperCase();
  manual[idKey] = manual[idKey] || {};
  if (value === f.auto && f.auto !== '') {
    // Back to the workbook value: no need to remember it
    delete manual[idKey][key];
    row.fields[key] = { value, source: 'master', auto: f.auto };
  } else {
    manual[idKey][key] = value;
    row.fields[key] = { value, source: value ? 'manual' : 'empty', auto: f.auto };
  }
  if (!Object.keys(manual[idKey]).length) delete manual[idKey];
  saveManual(manual);
  refreshRowClasses(row, tr);
  renderSummary();
}
// Re-evaluate highlights in place (a typed "Drop" or "Coating Inside" changes which cells are required)
function refreshRowClasses(row, tr) {
  const required = new Set(C.requiredKeys(rowValues(row)));
  tr.querySelectorAll('td[data-key]').forEach((td) => {
    const key = td.dataset.key;
    const base = cellClass(row, key, required);
    td.classList.remove('missing', 'calc', 'entered');
    if (base) td.classList.add(base);
  });
}
function renderSummary() {
  const included = state.rows.filter((r) => r.include);
  let missingFields = 0, rowsMissing = 0, calcFields = 0;
  included.forEach((r) => {
    const m = missingCount(r);
    missingFields += m;
    if (m) rowsMissing++;
    COLUMNS.forEach(({ key }) => { if (r.fields[key].source === 'calc' && r.fields[key].value) calcFields++; });
  });
  $('summaryText').innerHTML = `<b>${included.length}</b> row${included.length === 1 ? '' : 's'} · ` +
    (missingFields ? `<b>${rowsMissing}</b> row${rowsMissing === 1 ? '' : 's'} need input (<b>${missingFields}</b> highlighted cell${missingFields === 1 ? '' : 's'})` : 'nothing left to fill in ✓') +
    (calcFields ? ` · ${calcFields} calculated value${calcFields === 1 ? '' : 's'} to verify` : '');
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
  const rows = state.rows.filter((r) => r.include);
  if (!rows.length) { showToast('No rows are selected for the download.', true); return; }
  const sheetName = sanitizeSheetName($('sheetNameInput').value);
  const ws = {};
  const put = (r, c, v, s) => {
    const cell = v === '' || v == null ? { t: 's', v: '' } : (typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) });
    if (s) cell.s = s;
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  };
  // Header row (row 2): column B blank, then the template labels
  put(1, 1, '', { fill: FILL.gray, border: BORDER });
  COLUMNS.forEach((col, i) => {
    put(1, i + 2, col.label, { font: { bold: true }, border: BORDER, alignment: { wrapText: true, vertical: 'center' }, fill: col.optional ? FILL.gray : undefined });
  });
  rows.forEach((row, ri) => {
    const r = ri + 2;
    const required = new Set(C.requiredKeys(rowValues(row)));
    put(r, 1, ri + 1, { border: BORDER });
    COLUMNS.forEach((col, i) => {
      const f = row.fields[col.key];
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

  const legend = XLSX.utils.aoa_to_sheet([
    ['Legend'],
    ['Yellow', 'Needs manual input (not found in the tracking workbook)'],
    ['Light blue', "Taken from the cell's own OCV tracking sheet because the Master row was empty — please verify"],
    ['Note', 'Frozen IR Pass/Fail = NG when Frozen IR < 35 MΩ, OK otherwise (when the Master sheet has no result)'],
    [],
    ['Source file', state.source ? state.source.fileName : ''],
    ['Previous report', state.prev ? state.prev.fileName : '(none)'],
    ['Generated', new Date().toLocaleString()],
  ]);
  legend.A1.s = { font: { bold: true } };
  legend.A2.s = { fill: FILL.missing };
  legend.A3.s = { fill: FILL.calc };
  legend['!cols'] = [{ wch: 16 }, { wch: 80 }];

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
function doReset() {
  if ($('resetManual').checked) { try { localStorage.removeItem(MANUAL_STORE_KEY); } catch (e) {} }
  state.source = null;
  state.prev = null;
  state.rows = [];
  state.selectedLots.clear();
  $('sourceInput').value = '';
  $('prevInput').value = '';
  $('prevStatus').textContent = '';
  $('sourceStatus').hidden = true;
  $('lotGrid').innerHTML = '';
  $('reportBody').innerHTML = '';
  $('step2').hidden = true;
  $('step3').hidden = true;
  $('resetOverlay').classList.remove('open');
  window.scrollTo({ top: 0, behavior: 'smooth' });
  showToast($('resetManual').checked ? 'Reset complete. Saved manual entries were cleared.' : 'Reset complete.');
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

  $('resetBtn').onclick = () => { $('resetManual').checked = false; $('resetOverlay').classList.add('open'); };
  $('resetCancelBtn').onclick = () => $('resetOverlay').classList.remove('open');
  $('resetOverlay').addEventListener('click', (e) => { if (e.target === $('resetOverlay')) $('resetOverlay').classList.remove('open'); });
  $('resetConfirmBtn').onclick = doReset;
}
init();
