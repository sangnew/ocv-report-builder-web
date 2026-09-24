// Parses workbooks off the main thread so the page stays responsive with large (30MB+) files.
importScripts(
  'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
  'https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js',
  'xlsx-subset.js?v=1',
  'convert.js?v=4'
);

var source = null; // { handle (XlsxSubset workbook or null), data, names }

function findMasterSheet(names) {
  var norm = function (s) { return String(s).toLowerCase().replace(/[^a-z]/g, ''); };
  for (var i = 0; i < names.length; i++) if (norm(names[i]).indexOf('master') === 0) return names[i];
  return names[0];
}

// Open with the fast partial reader; fall back to a full SheetJS parse for unusual files (e.g. .xls)
function readSheets(names, opts) {
  var input = source.handle ? source.handle.open(names) : source.data;
  return XLSX.read(input, Object.assign({ type: 'array', sheets: names }, opts));
}

self.onmessage = function (e) {
  var msg = e.data;
  try {
    if (msg.type === 'source') {
      var data = new Uint8Array(msg.buf);
      var handle = null, names;
      try {
        handle = XlsxSubset.openWorkbook(fflate, data);
        names = handle.names;
      } catch (zipErr) {
        names = XLSX.read(data, { type: 'array', bookSheets: true }).SheetNames;
      }
      source = { handle: handle, data: data, names: names };
      var masterName = findMasterSheet(names);
      var wb = readSheets([masterName], { dense: true });
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[masterName], { header: 1, raw: false, defval: '' });
      var parsed = OcvConvert.parseMaster(rows, names);
      self.postMessage({
        id: msg.id, ok: true, masterName: masterName, sheetCount: names.length,
        cells: parsed.cells, missingHeaders: parsed.missingHeaders, lots: OcvConvert.lotSummary(parsed.cells),
      });
    } else if (msg.type === 'tracking') {
      var result = {};
      if (msg.sheets.length) {
        var twb = readSheets(msg.sheets, { dense: true, sheetRows: 60 });
        msg.sheets.forEach(function (name) { result[name] = OcvConvert.analyzeTrackingSheet(twb.Sheets[name]); });
      }
      self.postMessage({ id: msg.id, ok: true, analysis: result });
    } else if (msg.type === 'report') {
      var rwb = XLSX.read(new Uint8Array(msg.buf), { type: 'array' });
      var found = null, sheetName = null;
      for (var i = 0; i < rwb.SheetNames.length && !found; i++) {
        var rrows = XLSX.utils.sheet_to_json(rwb.Sheets[rwb.SheetNames[i]], { header: 1, raw: false, defval: '' });
        found = OcvConvert.parsePreviousReport(rrows);
        if (found) sheetName = rwb.SheetNames[i];
      }
      self.postMessage({ id: msg.id, ok: true, report: found, sheetName: sheetName });
    }
  } catch (err) {
    self.postMessage({ id: msg.id, ok: false, error: String(err && err.message || err) });
  }
};
