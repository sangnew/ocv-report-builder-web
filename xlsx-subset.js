// Fast partial reading of very large .xlsx files.
// SheetJS inflates every part of the zip up front, which takes 20s+ in the browser for a 36MB
// workbook full of images. Here fflate decompresses only the parts we need, and the result is
// repacked into a small uncompressed zip that SheetJS parses quickly.
(function (root) {
  'use strict';
  var CORE = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/sharedStrings.xml', 'xl/styles.xml'];

  function decodeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  function attr(tag, name) {
    var m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag);
    return m ? decodeXml(m[1]) : null;
  }

  // Returns { names: [...sheet names in order], open(sheetNames) -> Uint8Array (small xlsx) }
  function openWorkbook(fflate, data) {
    var core = fflate.unzipSync(data, { filter: function (f) { return CORE.indexOf(f.name) !== -1; } });
    var td = new TextDecoder();
    var wbXml = td.decode(core['xl/workbook.xml']);
    var relsXml = td.decode(core['xl/_rels/workbook.xml.rels']);
    var rels = {};
    (relsXml.match(/<Relationship\b[^>]*>/g) || []).forEach(function (tag) {
      var target = attr(tag, 'Target') || '';
      target = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      rels[attr(tag, 'Id')] = target;
    });
    var sheets = [];
    (wbXml.match(/<sheet\b[^>]*>/g) || []).forEach(function (tag) {
      var rid = (/\s[\w]+:id="([^"]*)"/.exec(tag) || [])[1];
      sheets.push({ name: attr(tag, 'name'), path: rels[rid] });
    });
    var byName = {};
    sheets.forEach(function (s) { byName[s.name] = s.path; });

    return {
      names: sheets.map(function (s) { return s.name; }),
      open: function (wanted) {
        var paths = {};
        wanted.forEach(function (n) { if (byName[n]) paths[byName[n]] = true; });
        var parts = fflate.unzipSync(data, { filter: function (f) { return !!paths[f.name]; } });
        var files = {};
        Object.keys(core).forEach(function (k) { files[k] = [core[k], { level: 0 }]; });
        Object.keys(parts).forEach(function (k) { files[k] = [parts[k], { level: 0 }]; });
        return fflate.zipSync(files, { level: 0 });
      },
    };
  }

  var api = { openWorkbook: openWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxSubset = api;
})(typeof self !== 'undefined' ? self : this);
