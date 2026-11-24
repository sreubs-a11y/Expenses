/* In-memory stand-ins for SpreadsheetApp, DriveApp etc. Works in node (vm) and the browser. */
function makeGoogleMock(rawBills, rawCar, opts) {
  opts = opts || {};
  function revive(v) { return v && v.$d ? new Date(+v.$d.slice(0, 4), +v.$d.slice(5, 7) - 1, +v.$d.slice(8, 10)) : v; }
  function load(raw) { var d = {}; for (var k in raw) d[k] = raw[k].map(function (r) { return r.map(revive); }); return d; }
  var books = { BILLS: { data: load(rawBills), rich: {} }, CARID: { data: load(rawCar), rich: {} } };
  function evalCell(book, name, r, c) {
    var v = (book.data[name][r - 1] || [])[c - 1];
    if (typeof v !== 'string' || v[0] !== '=') return v === undefined ? '' : v;
    var g = function (col, row) { var x = evalCell(book, name, row, col.charCodeAt(0) - 64); return typeof x === 'number' ? x : 0; }, m;
    if ((m = v.match(/^=IF\(AND\(N\(B(\d+)\)>0,N\(D\1\)>0\),D\1\/B\1,""\)$/))) { var b = g('B', +m[1]), d = g('D', +m[1]); return b > 0 && d > 0 ? d / b : ''; }
    if ((m = v.match(/^=E(\d+)-E(\d+)$/))) return g('E', +m[1]) - g('E', +m[2]);
    if ((m = v.match(/^=IF\(F(\d+)>0,SUM\(B(\d+):B(\d+)\)\/F\1\*100,""\)$/))) {
      var f = g('F', +m[1]); if (!(f > 0)) return ''; var s = 0; for (var i = +m[2]; i <= +m[3]; i++) s += g('B', i); return s / f * 100;
    }
    return '';
  }
  function width(book, name) { return Math.max.apply(null, [0].concat(book.data[name].map(function (r) { return r.length; }))); }
  function ensureRow(book, name, r) { var w = width(book, name); while (book.data[name].length < r) book.data[name].push(new Array(w).fill('')); var row = book.data[name][r - 1]; while (row.length < w) row.push(''); }
  function richOf(book, name, r, c) {
    var key = name + '!' + r + ',' + c, rt = book.rich[key];
    var text = String(evalCell(book, name, r, c) || '');
    var runs = rt ? rt.links.map(function (l) { return { getText: function () { return rt.text.slice(l.s, l.e); }, getLinkUrl: function () { return l.url; } }; })
                  : [{ getText: function () { return text; }, getLinkUrl: function () { return null; } }];
    return { getRuns: function () { return runs; } };
  }
  function sheet(book, name) {
    var D = book.data;
    return {
      getName: function () { return name; },
      getDataRange: function () { return { getValues: function () { var w = width(book, name); return D[name].map(function (row, i) { var o = []; for (var j = 1; j <= w; j++) o.push(evalCell(book, name, i + 1, j)); return o; }); } }; },
      getLastRow: function () { var n = D[name].length; while (n > 0 && D[name][n - 1].every(function (v) { return v === '' || v === null; })) n--; return n; },
      getLastColumn: function () { return width(book, name); },
      insertColumnAfter: function (c) { D[name].forEach(function (r) { r.splice(c, 0, ''); }); },
      appendRow: function (r) { var w = width(book, name); D[name].push(r.concat(new Array(Math.max(0, w - r.length)).fill(''))); },
      deleteRow: function (r) { D[name].splice(r - 1, 1); },
      setFrozenRows: function () {},
      getRange: function (r, c, nr, nc) {
        nr = nr || 1; nc = nc || 1;
        var api = {
          getValues: function () { var o = []; for (var i = 0; i < nr; i++) { var row = []; for (var j = 0; j < nc; j++) row.push(evalCell(book, name, r + i, c + j)); o.push(row); } return o; },
          getValue: function () { return evalCell(book, name, r, c); },
          setValues: function (v) { for (var i = 0; i < nr; i++) { ensureRow(book, name, r + i); for (var j = 0; j < nc; j++) D[name][r - 1 + i][c - 1 + j] = v[i][j]; } return api; },
          setValue: function (v) { ensureRow(book, name, r); D[name][r - 1][c - 1] = v; return api; },
          setFormula: function (f) { return api.setValue(f); },
          setFormulas: function (v) { return api.setValues(v); },
          getRichTextValues: function () { var o = []; for (var i = 0; i < nr; i++) o.push([richOf(book, name, r + i, c)]); return o; },
          setRichTextValue: function (rt) { ensureRow(book, name, r); book.rich[name + '!' + r + ',' + c] = rt; D[name][r - 1][c - 1] = rt.text; return api; },
          setNumberFormats: function () { return api; }, setFontFamily: function () { return api; }, setFontSize: function () { return api; }
        };
        return api;
      }
    };
  }
  function ss(id) { var b = books[id]; return { getSheetByName: function (n) { return b.data[n] ? sheet(b, n) : null; },
    insertSheet: function (n) { if (b.data[n]) throw new Error('A sheet with the name "' + n + '" already exists.'); b.data[n] = []; return sheet(b, n); }, getUrl: function () { return '#sheet-' + id; }, getSpreadsheetTimeZone: function () { return 'Australia/Sydney'; } }; }
  var files = [], props = {}, cache = {}, cacheLog = [];
  function folder(n) {
    var subs = {};
    return { getId: function () { return 'FOLDER-' + n; }, getName: function () { return n; },
      getFoldersByName: function (t) { return { hasNext: function () { return !!subs[t]; }, next: function () { return subs[t]; } }; },
      createFolder: function (t) { return (subs[t] = folder(t)); },
      createFile: function (b) { files.push(b); var url = opts.blobUrl ? opts.blobUrl(b) : 'https://drive.example/' + files.length + '/' + b.name; return { getUrl: function () { return url; } }; } };
  }
  var roots = {};
  return {
    books: books, files: files,
    SpreadsheetApp: {
      getActiveSpreadsheet: function () { return ss('BILLS'); },
      openById: function (id) { if (!books[id]) throw new Error('not found'); return ss(id); },
      newRichTextValue: function () { var t = '', links = []; var b = { setText: function (x) { t = x; return b; }, setLinkUrl: function (s, e, url) { links.push({ s: s, e: e, url: url }); return b; }, build: function () { return { text: t, links: links }; } }; return b; },
      getUi: function () { return { alert: function () {}, createMenu: function () { var m = { addItem: function () { return m; }, addToUi: function () {} }; return m; } }; }
    },
    HtmlService: {},
    ContentService: { MimeType: { JSON: 'application/json' }, createTextOutput: function (t) { var o = { content: t, mime: 'text/plain', setMimeType: function (m) { o.mime = m; return o; }, getContent: function () { return o.content; } }; return o; } },
    ScriptApp: { getService: function () { return { getUrl: function () { return 'https://script.google.com/macros/s/TEST/exec'; } }; } },
    props: props,
    CacheService: { getScriptCache: function () { return { get: function (k) { cacheLog.push('get ' + k); return cache[k] || null; }, put: function (k, v) { cacheLog.push('put'); cache[k] = v; }, remove: function (k) { cacheLog.push('remove'); delete cache[k]; } }; } },
    cacheLog: cacheLog,
    LockService: { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } },
    PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } }; } },
    Utilities: {
      base64Decode: function (s) { if (typeof Buffer !== 'undefined') return Buffer.from(s, 'base64'); var b = atob(s), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; },
      newBlob: function (bytes, mime, name) { return { bytes: bytes, mime: mime, name: name }; },
      sleep: function () {},
      getUuid: function () { var h = '0123456789abcdef', o = ''; for (var i = 0; i < 32; i++) o += h[Math.floor(Math.random() * 16)]; return o.slice(0, 8) + '-' + o.slice(8, 12) + '-' + o.slice(12, 16) + '-' + o.slice(16, 20) + '-' + o.slice(20); }
    },
    DriveApp: {
      getFolderById: function (id) { for (var k in roots) if (roots[k].getId() === id) return roots[k]; throw new Error('no folder'); },
      getFoldersByName: function (n) { return { hasNext: function () { return !!roots[n]; }, next: function () { return roots[n]; } }; },
      createFolder: function (n) { return (roots[n] = folder(n)); }
    }
  };
}
if (typeof module !== 'undefined') module.exports = makeGoogleMock;
