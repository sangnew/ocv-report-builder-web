// Conversion rules: "Mass Production E&L Grade OCV Tracking Sheet" -> Frozen IR / Spot report rows.
// Pure functions shared by the browser (window.OcvConvert) and Node tests (module.exports).
(function (root) {
  'use strict';

  // Output columns, in the exact order and wording of the report template (test1.xlsx).
  var COLUMNS = [
    { key: 'lot', label: 'LOT' },
    { key: 'cellId', label: 'Cell ID' },
    { key: 'grade', label: 'Grade' },
    { key: 'docv', label: 'dOCV' },
    { key: 'frozenIr', label: 'Frozen IR Result (35MOhm)' },
    { key: 'frozenPf', label: 'Frozen IR Pass/Fail' },
    { key: 'voltageDrop', label: 'voltage drop / no drop' },
    { key: 'layer', label: 'Voltage Dropped Layer' },
    { key: 'docvV', label: 'dOCV (V)' },
    { key: 'spot', label: 'Spot Found' },
    { key: 'topBack', label: 'Top/ Back' },
    { key: 'x', label: 'x' },
    { key: 'y', label: 'y' },
    { key: 'shape', label: 'Shape' },
    { key: 'sem', label: 'SEM/EDS Analysis' },
    { key: 'location', label: 'Location' },
    { key: 'longSide', label: 'Long side' },
    { key: 'shortSide', label: 'Short side' },
    { key: 'height', label: 'Height' },
    { key: 'azs', label: 'AZS', optional: true },
    { key: 'dnc', label: 'DNC', optional: true },
  ];
  // Columns that never exist in the source workbook: always typed by hand.
  var MANUAL_ONLY = ['shape', 'location', 'longSide', 'shortSide', 'height', 'azs', 'dnc'];
  var FROZEN_IR_LIMIT_MOHM = 35;

  function norm(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s\r\n]+/g, ''); }
  function text(v) { return String(v == null ? '' : v).trim(); }
  function isNum(s) { return s !== '' && !isNaN(Number(s)); }

  // Header names in the "Master E & L" sheet (matched loosely so moved columns still work).
  var MASTER_HEADERS = {
    lot: ['lotid'],
    cellId: ['cellid'],
    grade: ['grade'],
    docv: ['docv(mv)', 'docv'],
    ntf: ['ntf'],
    frozenIr: ['frozenir(mohm)'],
    frozenResult: ['frozenirresult'],
    anodeSheet: ['anodesheet'],
    voltageDrop: ['voltagedrop(v)'],
    topBack: ['anodetop/back'],
    x: ['x(mm)'],
    y: ['y(mm)'],
    burnMark: ['burnmark/pinhole/none'],
    eds: ['edsimpurityresults'],
  };

  // Find the header row (the one containing "Cell ID") and map each field to a column index.
  function detectMasterColumns(rows) {
    for (var r = 0; r < Math.min(rows.length, 30); r++) {
      var normalized = (rows[r] || []).map(norm);
      if (normalized.indexOf('cellid') === -1) continue;
      var map = {};
      Object.keys(MASTER_HEADERS).forEach(function (field) {
        var candidates = MASTER_HEADERS[field];
        for (var i = 0; i < candidates.length; i++) {
          var idx = normalized.indexOf(candidates[i]);
          if (idx !== -1) { map[field] = idx; return; }
        }
      });
      return { headerRow: r, cols: map };
    }
    return null;
  }

  function lotSortKey(lot) {
    var m = /^([A-Za-z]+)0*(\d+)/.exec(lot);
    return m ? m[1].toUpperCase() + String(m[2]).padStart(4, '0') : lot.toUpperCase();
  }
  // "FH9" -> "FH09" (used for the report's sheet name, e.g. "260921 FH09~FH14")
  function padLot(lot) {
    var m = /^([A-Za-z]+)0*(\d+)$/.exec(lot);
    return m ? m[1] + String(m[2]).padStart(2, '0') : lot;
  }

  // Parse master rows into cell records (in sheet order).
  function parseMaster(rows, sheetNames) {
    var det = detectMasterColumns(rows);
    if (!det) throw new Error('Could not find the "Cell ID" header row in the Master E & L sheet.');
    var c = det.cols;
    var missingHeaders = Object.keys(MASTER_HEADERS).filter(function (f) { return c[f] === undefined; });
    var sheetSet = {};
    (sheetNames || []).forEach(function (n) { sheetSet[text(n).toUpperCase()] = n; });
    var cells = [];
    var seen = {};
    for (var r = det.headerRow + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var id = text(row[c.cellId]);
      if (!id) continue;
      var key = id.toUpperCase();
      var rec = { row: r + 1, trackingSheet: sheetSet[key] || null, duplicate: !!seen[key] };
      Object.keys(c).forEach(function (f) { rec[f] = text(row[c[f]]); });
      seen[key] = true;
      cells.push(rec);
    }
    return { cells: cells, missingHeaders: missingHeaders };
  }

  function lotSummary(cells) {
    var lots = {};
    cells.forEach(function (c) {
      var lot = c.lot || '(no LOT)';
      lots[lot] = lots[lot] || { lot: lot, total: 0, withSheet: 0 };
      lots[lot].total++;
      if (c.trackingSheet) lots[lot].withSheet++;
    });
    return Object.keys(lots).map(function (k) { return lots[k]; })
      .sort(function (a, b) { return lotSortKey(a.lot).localeCompare(lotSortKey(b.lot)); });
  }

  // The OCV tracking sheet of one cell: R6 = dropped layer number (or "NTF"), S6 = max dOCV.
  function readTrackingSheet(ws) {
    if (!ws) return null;
    function get(addr) {
      var rowsArr = ws['!data'] || (Array.isArray(ws) ? ws : null); // dense mode (newer / older SheetJS)
      if (rowsArr) {
        var m = /^([A-Z]+)(\d+)$/.exec(addr);
        var col = 0;
        for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
        var rowArr = rowsArr[Number(m[2]) - 1];
        var cell = rowArr && rowArr[col - 1];
        return cell ? cell.v : undefined;
      }
      return ws[addr] ? ws[addr].v : undefined;
    }
    return { layer: get('R6'), maxDocv: get('S6') };
  }

  function fmtDocv(v) {
    var n = Number(v);
    if (v === undefined || v === null || v === '' || isNaN(n)) return '';
    return String(Math.round(n * 10000) / 10000);
  }
  function titleTopBack(v) {
    var s = text(v).toLowerCase();
    if (s === 'top') return 'Top';
    if (s === 'back') return 'Back';
    return '';
  }

  // Which columns must be filled for this row (depends on voltage drop / location values).
  function requiredKeys(values) {
    var req = ['lot', 'cellId', 'grade', 'docv', 'frozenIr', 'frozenPf', 'voltageDrop'];
    if (text(values.voltageDrop).toLowerCase() === 'drop') {
      req.push('layer', 'docvV', 'spot', 'topBack', 'x', 'y', 'shape', 'sem', 'location');
      if (/inside/i.test(text(values.location))) req.push('longSide', 'shortSide', 'height');
    }
    return req;
  }

  // Build one report row. Each field gets { value, source }:
  //   source: 'master' (copied from Master E & L), 'calc' (derived / from tracking sheet — verify),
  //           'empty' (nothing found)
  function buildRow(cell, tracking) {
    var f = {};
    function set(key, value, source) { f[key] = { value: text(value), source: text(value) ? source : 'empty' }; }

    set('lot', cell.lot, 'master');
    set('cellId', cell.cellId, 'master');
    set('grade', cell.grade, 'master');
    set('docv', cell.docv, 'master');
    set('frozenIr', cell.frozenIr, 'master');

    var pf = text(cell.frozenResult).toUpperCase();
    if (pf === 'NG' || pf === 'OK') set('frozenPf', pf, 'master');
    else if (isNum(text(cell.frozenIr))) set('frozenPf', Number(cell.frozenIr) < FROZEN_IR_LIMIT_MOHM ? 'NG' : 'OK', 'master'); // fixed 35 MΩ rule from the column header
    else set('frozenPf', '', 'empty');

    var hasDropData = !!(text(cell.anodeSheet) || text(cell.voltageDrop));
    var trackLayer = tracking ? text(tracking.layer) : '';
    if (text(cell.ntf).toUpperCase() === 'NTF') set('voltageDrop', 'NTF', 'master');
    else if (hasDropData) set('voltageDrop', 'Drop', 'master');
    else if (trackLayer.toUpperCase() === 'NTF') set('voltageDrop', 'NTF', 'calc');
    else if (isNum(trackLayer)) set('voltageDrop', 'Drop', 'calc');
    else set('voltageDrop', '', 'empty');

    if (text(cell.anodeSheet)) set('layer', cell.anodeSheet, 'master');
    else if (isNum(trackLayer) && f.voltageDrop.value === 'Drop') set('layer', trackLayer, 'calc');
    else set('layer', '', 'empty');

    if (text(cell.voltageDrop)) set('docvV', cell.voltageDrop, 'master');
    else if (tracking && f.voltageDrop.value === 'Drop') set('docvV', fmtDocv(tracking.maxDocv), 'calc');
    else set('docvV', '', 'empty');

    var burn = text(cell.burnMark);
    set('spot', burn && burn.toLowerCase() !== 'none' ? 'Spot Found' : '', 'master');
    set('topBack', titleTopBack(cell.topBack), 'master');
    set('x', cell.x, 'master');
    set('y', cell.y, 'master');

    var eds = text(cell.eds);
    var ntfCol = text(cell.ntf);
    if (eds && eds.toUpperCase() !== 'N/A') set('sem', eds, 'master');
    else if (ntfCol && ['NTF', 'OVER.F'].indexOf(ntfCol.toUpperCase()) === -1) set('sem', ntfCol, 'master');
    else set('sem', '', 'empty');

    MANUAL_ONLY.forEach(function (k) { set(k, '', 'empty'); });
    return f;
  }

  // Report header row as it appears in the template.
  function headerRow() { return COLUMNS.map(function (c) { return c.label; }); }

  // Map a previous report's header labels to column keys.
  function detectReportColumns(rows) {
    var byLabel = {};
    COLUMNS.forEach(function (c) { byLabel[norm(c.label)] = c.key; });
    for (var r = 0; r < Math.min(rows.length, 30); r++) {
      var normalized = (rows[r] || []).map(norm);
      if (normalized.indexOf('cellid') === -1) continue;
      var map = {};
      normalized.forEach(function (h, i) { if (byLabel[h]) map[byLabel[h]] = i; });
      return { headerRow: r, cols: map };
    }
    return null;
  }
  function parsePreviousReport(rows) {
    var det = detectReportColumns(rows);
    if (!det) return null;
    var out = {};
    for (var r = det.headerRow + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var id = text(row[det.cols.cellId]).toUpperCase();
      if (!id) continue;
      var rec = {};
      Object.keys(det.cols).forEach(function (k) { rec[k] = text(row[det.cols[k]]); });
      out[id] = rec;
    }
    return out;
  }

  var api = {
    COLUMNS: COLUMNS,
    MANUAL_ONLY: MANUAL_ONLY,
    detectMasterColumns: detectMasterColumns,
    parseMaster: parseMaster,
    lotSummary: lotSummary,
    lotSortKey: lotSortKey,
    padLot: padLot,
    readTrackingSheet: readTrackingSheet,
    buildRow: buildRow,
    requiredKeys: requiredKeys,
    headerRow: headerRow,
    parsePreviousReport: parsePreviousReport,
    isNum: isNum,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.OcvConvert = api;
})(typeof self !== 'undefined' ? self : this);
