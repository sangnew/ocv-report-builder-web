// Conversion rules: "Mass Production E&L Grade OCV Tracking Sheet" -> Frozen IR / Spot report rows.
// Pure functions shared by the browser (window.OcvConvert) and Node tests (module.exports).
(function (root) {
  'use strict';

  // Output columns, in the order and wording of the report template (test1.xlsx).
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
  ];
  // Columns that never exist in the source workbook: always typed by hand.
  var MANUAL_ONLY = ['shape', 'location', 'longSide', 'shortSide', 'height'];
  var FROZEN_IR_LIMIT_MOHM = 35;
  // Same outlier rule as the tracking sheet's R6 formula: a layer is the dropped layer when its
  // dOCV is more than 2.6 standard deviations above the average of the inner layers.
  var SIGMA_LIMIT = 2.6;
  var DEFAULT_MIN_DROP_V = 0.0015;

  function norm(s) { return String(s == null ? '' : s).toLowerCase().replace(/[\s\r\n]+/g, ''); }
  function text(v) { return String(v == null ? '' : v).trim(); }
  function isNum(s) { return s !== '' && s !== null && s !== undefined && !isNaN(Number(s)); }
  function num(v) { return isNum(v) ? Number(v) : null; }

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
    return m ? m[1].toUpperCase() + String(m[2]).padStart(4, '0') : String(lot).toUpperCase();
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
    for (var r = det.headerRow + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var id = text(row[c.cellId]);
      if (!id) continue;
      var rec = { row: r + 1, trackingSheet: sheetSet[id.toUpperCase()] || null };
      Object.keys(c).forEach(function (f) { rec[f] = text(row[c[f]]); });
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

  function sheetGetter(ws) {
    var rowsArr = ws['!data'] || (Array.isArray(ws) ? ws : null); // dense mode (newer / older SheetJS)
    return function (col, row) {
      if (rowsArr) {
        var ci = 0;
        for (var i = 0; i < col.length; i++) ci = ci * 26 + (col.charCodeAt(i) - 64);
        var r = rowsArr[row - 1];
        var cell = r && r[ci - 1];
        return cell ? cell.v : undefined;
      }
      var c = ws[col + row];
      return c ? c.v : undefined;
    };
  }

  // Analyze one cell's OCV tracking sheet from the raw readings.
  // Layout: row 5 = tracking dates in C/D/E, rows 6.. = one anode layer per row (layer number in B,
  // OCV readings per date in C/D/E). Per layer: dOCV = max(C-D, C-E). The dropped layer is the one
  // whose dOCV stands out (> 2.6σ over the inner layers, like the sheet's R6 formula).
  function analyzeTrackingSheet(ws) {
    if (!ws) return null;
    var g = sheetGetter(ws);
    var layers = [];
    for (var r = 6; r <= 200; r++) {
      var b = text(g('B', r));
      if (!b) break;
      var c = num(g('C', r)), d = num(g('D', r)), e = num(g('E', r));
      if (c === null || (d === null && e === null)) { layers.push({ layer: b, docv: null }); continue; }
      var d1 = d === null ? 0 : c - d;
      var d2 = e === null ? 0 : c - e;
      layers.push({ layer: b, docv: Math.max(d1, d2) });
    }
    var days = ['C', 'D', 'E'].filter(function (col) { return text(g(col, 5)); }).length;
    // Statistics over the inner layers (the sheet uses rows 7..41 of 6..42: first and last layer excluded)
    var inner = layers.slice(1, -1).filter(function (l) { return l.docv !== null; });
    if (inner.length < 3) return { ok: false, days: days, reason: 'Not enough OCV readings in the tracking sheet' };
    var mean = inner.reduce(function (s, l) { return s + l.docv; }, 0) / inner.length;
    var sd = Math.sqrt(inner.reduce(function (s, l) { return s + (l.docv - mean) * (l.docv - mean); }, 0) / (inner.length - 1));
    var best = inner[0];
    inner.forEach(function (l) { if (l.docv > best.docv) best = l; });
    var sigma = sd > 0 ? (best.docv - mean) / sd : 0;
    return { ok: true, days: days, layer: best.layer, docv: best.docv, sigma: sigma };
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

  // Build one report row. Each field gets { value, source, note }:
  //   source: 'master'  copied from Master E & L
  //           'sheet'   from the cell's OCV tracking sheet analysis
  //           'calc'    tracking sheet analysis disagrees with Master E & L — verify
  //           'empty'   nothing found
  function buildRow(cell, analysis, opts) {
    var minDrop = opts && isNum(opts.minDropV) ? Number(opts.minDropV) : DEFAULT_MIN_DROP_V;
    var f = {};
    function set(key, value, source, note) {
      f[key] = { value: text(value), source: text(value) ? source : 'empty' };
      if (note) f[key].note = note;
    }

    set('lot', cell.lot, 'master');
    set('cellId', cell.cellId, 'master');
    set('grade', cell.grade, 'master');
    set('docv', cell.docv, 'master');
    set('frozenIr', cell.frozenIr, 'master');

    var pf = text(cell.frozenResult).toUpperCase();
    if (pf === 'NG' || pf === 'OK') set('frozenPf', pf, 'master');
    else if (isNum(text(cell.frozenIr))) set('frozenPf', Number(cell.frozenIr) < FROZEN_IR_LIMIT_MOHM ? 'NG' : 'OK', 'master'); // fixed 35 MΩ rule from the column header
    else set('frozenPf', '', 'empty');

    // What the Master sheet says about the voltage drop (if anything)
    var masterVd = text(cell.ntf).toUpperCase() === 'NTF' ? 'NTF'
      : (text(cell.anodeSheet) || text(cell.voltageDrop)) ? 'Drop' : '';
    var masterLayer = text(cell.anodeSheet);

    if (analysis && analysis.ok) {
      var isDrop = analysis.sigma > SIGMA_LIMIT && analysis.docv >= minDrop;
      var vd = isDrop ? 'Drop' : 'NTF';
      var detail = 'Tracking sheet: max dOCV ' + (analysis.docv * 1000).toFixed(2) + ' mV at layer ' + analysis.layer +
        ' (' + analysis.sigma.toFixed(1) + 'σ)';
      var conflict = masterVd && (masterVd !== vd || (isDrop && masterLayer && masterLayer !== String(analysis.layer)));
      var masterNote = masterVd ? ' · Master E & L: ' + masterVd + (masterLayer ? ' layer ' + masterLayer : '') : '';
      var src = conflict ? 'calc' : 'sheet';
      set('voltageDrop', vd, src, detail + masterNote);
      if (isDrop) {
        set('layer', analysis.layer, src, detail + masterNote);
        // Prefer the Master's rounded value when it describes the same layer
        if (text(cell.voltageDrop) && masterLayer === String(analysis.layer)) set('docvV', cell.voltageDrop, 'master');
        else set('docvV', fmtDocv(analysis.docv), src, detail + masterNote);
      } else {
        set('layer', '', 'empty');
        set('docvV', '', 'empty');
      }
    } else {
      // No usable tracking sheet: fall back to the Master sheet
      set('voltageDrop', masterVd, 'master', analysis && analysis.reason);
      set('layer', masterVd === 'Drop' ? masterLayer : '', 'master');
      set('docvV', masterVd === 'Drop' ? cell.voltageDrop : '', 'master');
    }

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
    SIGMA_LIMIT: SIGMA_LIMIT,
    DEFAULT_MIN_DROP_V: DEFAULT_MIN_DROP_V,
    detectMasterColumns: detectMasterColumns,
    parseMaster: parseMaster,
    lotSummary: lotSummary,
    lotSortKey: lotSortKey,
    padLot: padLot,
    analyzeTrackingSheet: analyzeTrackingSheet,
    buildRow: buildRow,
    requiredKeys: requiredKeys,
    parsePreviousReport: parsePreviousReport,
    isNum: isNum,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.OcvConvert = api;
})(typeof self !== 'undefined' ? self : this);
