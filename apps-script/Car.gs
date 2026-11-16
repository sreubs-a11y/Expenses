/**
 * Car expenses - ported from the standalone Car Expenses app.
 * Lives in the bills project and opens the "Vehicle Expenses" sheet by ID
 * (Settings tab of the bills sheet: "Car Sheet ID").
 * Tabs used in the car sheet: Fuel, Tolls, Other, Allowances, Summary.
 * Everything here is prefixed car/CAR_ so nothing clashes with the bills code.
 */

var CAR_CONFIG = {
  ROOT_FOLDER_NAME: 'Vehicle Expenses Receipts',
  TZ: 'Australia/Sydney'
};

var CAR_TABS = {
  fuel:  { name: 'Fuel',  cols: 9, receiptCol: 9, folder: 'Fuel' },
  tolls: { name: 'Tolls', cols: 5, receiptCol: 5, folder: 'Tolls' },
  other: { name: 'Other', cols: 5, receiptCol: 5, folder: 'Other' }
};

/* The Dashboard menu (incl. Recalculate car fuel economy) is in Api.gs. */

/* ------------------------------------------------------------------ reads */

/** True when a car sheet ID is set in Settings. */
function carConfigured_() { return !!carSheetId_(); }

function carSheetId_() {
  var s = readSettings_()._rows['Car Sheet ID'];
  var v = s ? String(s.value || '').trim() : '';
  var m = v.match(/\/d\/([a-zA-Z0-9_-]{20,})/);   // accept a full URL too
  return m ? m[1] : v;
}

function carGetData() {
  if (!carConfigured_()) return { configured: false };
  var ss = carSs_();
  return {
    configured: true,
    fuel:  carReadTab_(ss, 'fuel',  function (r) {
      return { date: carD_(r[0]), litres: carN_(r[1]), price: carN_(r[2]), total: carN_(r[3]),
               odo: carN_(r[4]), km: carN_(r[5]), l100: carN_(r[6]), notes: carS_(r[7]) }; }),
    tolls: carReadTab_(ss, 'tolls', function (r) {
      return { from: carD_(r[0]), to: carD_(r[1]), total: carN_(r[2]), comment: carS_(r[3]) }; }),
    other: carReadTab_(ss, 'other', function (r) {
      return { date: carD_(r[0]), item: carS_(r[1]), total: carN_(r[2]), comment: carS_(r[3]) }; }),
    allowances: carReadAllowances_(ss),
    linked: carLinkedItems_(),
    sheetUrl: ss.getUrl()
  };
}

/** Car "Other" items that feed the bills forecast, from the Recurring tab's Car Item column: {pattern: recurring name} */
function carLinkedItems_() {
  var out = {};
  readTable_(SHEETS.recurring).rows.forEach(function (r) {
    var link = String(r['Car Item'] || '').trim();
    if (link && String(r.Active || 'Yes').toLowerCase().indexOf('n') !== 0) out[link] = String(r.Name).trim();
  });
  return out;
}

/**
 * Car "Other" entries as plain objects for the forecast. Empty if no car sheet is set or it can't be opened.
 * Opening a second spreadsheet is the slowest part of a load, so the entries are kept in Google's cache for
 * an hour. Saving or deleting through the app clears it; so does the dashboard's refresh button.
 */
var CAR_CACHE_KEY_ = 'carOther:v1';
function carOtherForForecast_() {
  if (!carConfigured_()) return [];
  var cache = CacheService.getScriptCache(), hit = cache.get(CAR_CACHE_KEY_ + ':' + carSheetId_());
  if (hit) {
    try { return JSON.parse(hit).map(function (o) { o.date = fromIso_(o.date); return o; }); } catch (e) { /* fall through */ }
  }
  try {
    var list = carReadTab_(carSs_(), 'other', function (r) {
      return { date: toDate_(r[0]), item: carS_(r[1]), total: carN_(r[2]) };
    }).map(function (o) { return { date: o.date, item: o.item, total: o.total, receipt: (o.receipts[0] || {}).url || '' }; })
      .filter(function (o) { return o.date && o.total; });
    var json = JSON.stringify(list.map(function (o) { return { date: iso_(o.date), item: o.item, total: o.total, receipt: o.receipt }; }));
    if (json.length < 90000) cache.put(CAR_CACHE_KEY_ + ':' + carSheetId_(), json, 3600);
    return list;
  } catch (e) {
    return [];
  }
}
function carForgetCache_() { try { if (carConfigured_()) CacheService.getScriptCache().remove(CAR_CACHE_KEY_ + ':' + carSheetId_()); } catch (e) {} }

function carReadTab_(ss, key, map) {
  var t = CAR_TABS[key];
  var sh = ss.getSheetByName(t.name);
  if (!sh) throw new Error('Car sheet has no "' + t.name + '" tab');
  var last = carLastDataRow_(sh);
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, t.cols).getValues();
  var rich = sh.getRange(2, t.receiptCol, last - 1, 1).getRichTextValues();
  var out = [];
  vals.forEach(function (r, i) {
    if (r[0] === '' || r[0] === null) return;
    var o = map(r);
    o.row = i + 2;
    o.receipts = carLinksOf_(rich[i][0]);
    out.push(o);
  });
  return out;
}

function carReadAllowances_(ss) {
  var sh = ss.getSheetByName('Allowances');
  var res = { vehicle: 0, fuel: 0, tolls: 0, pct: 1 };
  if (!sh) return res;
  var inAllow = false;
  sh.getDataRange().getValues().forEach(function (r) {
    var k = String(r[0]).trim().toLowerCase();
    if (k === '%') res.pct = Number(r[1]) || 1;
    if (k === 'allowances') { inAllow = true; return; }
    if (!inAllow) return;
    if (k === 'vehicle') res.vehicle = Number(r[1]) || 0;
    if (k === 'fuel card') res.fuel = Number(r[1]) || 0;
    if (k === 'tolls') res.tolls = Number(r[1]) || 0;
  });
  return res;
}

/* ------------------------------------------------------------------ writes */

function carAddFuel(e) {
  return carWithLock_(function () {
    var sh = carSs_().getSheetByName(CAR_TABS.fuel.name);
    var r = carLastDataRow_(sh) + 1;
    var date = carParseDate_(e.date);
    var litres = carNum_(e.litres, 'Litres');
    var total = carOptNum_(e.total);
    var odo = carOptNum_(e.odometer);
    sh.getRange(r, 1, 1, 5).setValues([[date, litres, '', total === null ? '' : total, odo === null ? '' : odo]]);
    sh.getRange(r, 3).setFormula('=IF(AND(N(B' + r + ')>0,N(D' + r + ')>0),D' + r + '/B' + r + ',"")');
    sh.getRange(r, 8).setValue(e.notes || '');
    carFormatRow_(sh, r, ['dd/mm/yyyy', '0.00', '$0.000', '$#,##0.00', '#,##0', '#,##0', '0.00']);
    carSetReceipts_(sh, r, CAR_TABS.fuel, e.files, carFmt_(date) + '_Fuel' + (odo ? '_' + odo + 'km' : ''));
    carRebuildEconomy();
    return carGetData();
  });
}

function carAddToll(e) {
  return carWithLock_(function () {
    var sh = carSs_().getSheetByName(CAR_TABS.tolls.name);
    var r = carLastDataRow_(sh) + 1;
    var from = carParseDate_(e.from), to = carParseDate_(e.to);
    sh.getRange(r, 1, 1, 4).setValues([[from, to, carNum_(e.total, 'Total'), e.comment || '']]);
    carFormatRow_(sh, r, ['dd/mm/yyyy', 'dd/mm/yyyy', '$#,##0.00']);
    carSetReceipts_(sh, r, CAR_TABS.tolls, e.files, carFmt_(from) + '_to_' + carFmt_(to) + '_Tolls');
    return carGetData();
  });
}

function carAddOther(e) {
  return carWithLock_(function () {
    var sh = carSs_().getSheetByName(CAR_TABS.other.name);
    var r = carLastDataRow_(sh) + 1;
    var date = carParseDate_(e.date);
    var item = String(e.item || '').trim();
    if (!item) throw new Error('Item is required');
    sh.getRange(r, 1, 1, 4).setValues([[date, item, carNum_(e.total, 'Total'), e.comment || '']]);
    carFormatRow_(sh, r, ['dd/mm/yyyy', '@', '$#,##0.00']);
    carSetReceipts_(sh, r, CAR_TABS.other, e.files, carFmt_(date) + '_' + item.replace(/[^\w]+/g, '-'));
    carForgetCache_();
    return carGetData();
  });
}

/** Deletes a row after checking it still holds the entry the UI showed. Receipt files stay in Drive. */
function carDeleteEntry(type, row, date) {
  return carWithLock_(function () {
    var t = CAR_TABS[type];
    if (!t) throw new Error('Unknown type');
    var sh = carSs_().getSheetByName(t.name);
    var actual = carD_(sh.getRange(row, 1).getValue());
    if (row < 2 || actual !== date) throw new Error('Sheet has changed - refresh and try again.');
    sh.deleteRow(row);
    if (type === 'other') carForgetCache_();
    if (type === 'fuel') carRebuildEconomy();
    return carGetData();
  });
}

/**
 * Rewrites KM Travelled / L/100km for every fuel row.
 * Each row with an odometer reading measures back to the previous reading and
 * sums every fill in between, so fills without an odometer roll into the next one.
 * Blank if any fill in that span is missing litres.
 */
function carRebuildEconomy() {
  var sh = carSs_().getSheetByName(CAR_TABS.fuel.name);
  var last = carLastDataRow_(sh);
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, 5).getValues();
  var out = [], prev = null;
  vals.forEach(function (v, i) {
    var r = i + 2, f = '', g = '';
    if (v[4] !== '' && v[4] !== null) {
      if (prev !== null) {
        var ok = true;
        for (var k = prev + 1; k <= r; k++) if (vals[k - 2][1] === '' || vals[k - 2][1] === null) ok = false;
        if (ok) {
          f = '=E' + r + '-E' + prev;
          g = '=IF(F' + r + '>0,SUM(B' + (prev + 1) + ':B' + r + ')/F' + r + '*100,"")';
        }
      }
      prev = r;
    }
    out.push([f, g]);
  });
  sh.getRange(2, 6, out.length, 2).setFormulas(out);
}

/* ------------------------------------------------------------------ receipts */

function carSetReceipts_(sh, r, tab, files, baseName) {
  if (!files || !files.length) return;
  var folder = carSubFolder_(tab.folder);
  var saved = files.map(function (f) {
    var ext = (f.name.match(/\.[a-z0-9]+$/i) || [carMimeExt_(f.type)])[0];
    var label = f.label || 'Receipt';
    var name = (baseName + '_' + label + ext).replace(/\s+/g, '-');
    var blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', name);
    return { label: label, url: folder.createFile(blob).getUrl() };
  });
  var b = SpreadsheetApp.newRichTextValue().setText(saved.map(function (s) { return s.label; }).join(', '));
  var pos = 0;
  saved.forEach(function (s) { b.setLinkUrl(pos, pos + s.label.length, s.url); pos += s.label.length + 2; });
  sh.getRange(r, tab.receiptCol).setRichTextValue(b.build());
}

/** Reuses the existing "Vehicle Expenses Receipts" folder, so receipts from the old app stay together. */
function carSubFolder_(name) {
  var props = PropertiesService.getScriptProperties();
  var root = null, id = props.getProperty('CAR_ROOT_FOLDER_ID');
  if (id) { try { root = DriveApp.getFolderById(id); } catch (err) { root = null; } }
  if (!root) {
    var found = DriveApp.getFoldersByName(CAR_CONFIG.ROOT_FOLDER_NAME);
    root = found.hasNext() ? found.next() : DriveApp.createFolder(CAR_CONFIG.ROOT_FOLDER_NAME);
    props.setProperty('CAR_ROOT_FOLDER_ID', root.getId());
  }
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function carLinksOf_(rt) {
  if (!rt) return [];
  var out = [];
  rt.getRuns().forEach(function (run) {
    var url = run.getLinkUrl();
    if (url) out.push({ label: run.getText().replace(/^[,\s]+|[,\s]+$/g, '') || 'Receipt', url: url });
  });
  return out;
}

function carMimeExt_(t) {
  return ({ 'image/jpeg': '.jpg', 'image/png': '.png', 'application/pdf': '.pdf', 'image/webp': '.webp', 'image/heic': '.heic' })[t] || '';
}

/* ------------------------------------------------------------------ helpers */

var CAR_SS_ = null;
function carSs_() {
  if (CAR_SS_) return CAR_SS_;
  var id = carSheetId_();
  if (!id) throw new Error('Set "Car Sheet ID" on the Settings tab first.');
  try { CAR_SS_ = SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Could not open the car sheet - check "Car Sheet ID" on the Settings tab.'); }
  return CAR_SS_;
}

function carWithLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function carLastDataRow_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return 1;
  var col = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) if (col[i][0] !== '' && col[i][0] !== null) return i + 2;
  return 1;
}

function carFormatRow_(sh, r, fmts) {
  sh.getRange(r, 1, 1, fmts.length).setNumberFormats([fmts]);
  sh.getRange(r, 1, 1, sh.getLastColumn()).setFontFamily('Arial').setFontSize(10);
}

function carParseDate_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Invalid date');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function carNum_(v, label) {
  var n = Number(v);
  if (v === '' || v === null || v === undefined || !isFinite(n)) throw new Error(label + ' is required');
  return n;
}
function carOptNum_(v) { return v === '' || v === null || v === undefined || !isFinite(Number(v)) ? null : Number(v); }
function carFmt_(d) { return iso_(d); }
function carD_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v) ? iso_(v) : ''; }
function carN_(v) { return typeof v === 'number' && isFinite(v) ? v : null; }
function carS_(v) { return v === null || v === undefined ? '' : String(v); }
