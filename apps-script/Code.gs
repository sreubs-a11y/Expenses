/**
 * Bills dashboard - server side.
 * Bound to the "Bills Dashboard" Google Sheet (Extensions > Apps Script).
 * The app itself is hosted on GitHub Pages and talks to this script through Api.gs.
 */

var SHEETS = { bills: 'Bills', recurring: 'Recurring', txns: 'Transactions', rec: 'Reconcile', settings: 'Settings' };
var BILL_TYPE_ORDER = ['Electricity', 'Gas', 'Water', 'Sewer'];

/**
 * Bank accounts with their own forecast and reconciliation. Each has its own Recurring / Transactions /
 * Reconcile tabs; Settings are shared, with per-account rows named "<prefix><Setting>" (e.g. "Holiday Opening Date").
 * Only the bills account has utility bills and car items.
 */
var ACCOUNTS_ = {
  bills: { id: 'bills', label: 'Bills', recurring: 'Recurring', txns: 'Transactions', rec: 'Reconcile', prefix: '', bills: true, car: true },
  holiday: { id: 'holiday', label: 'Holiday', recurring: 'Holiday Recurring', txns: 'Holiday Transactions', rec: 'Holiday Reconcile', prefix: 'Holiday ', bills: false, car: false }
};
var ACCOUNT_ = ACCOUNTS_.bills;
/** Points every helper at one account's tabs for the rest of this server call. */
function useAccount_(id) {
  var a = ACCOUNTS_[id || 'bills'];
  if (!a || !Object.prototype.hasOwnProperty.call(ACCOUNTS_, id || 'bills')) throw new Error('Unknown account: ' + id);
  ACCOUNT_ = a;
  SHEETS.recurring = a.recurring; SHEETS.txns = a.txns; SHEETS.rec = a.rec;
  forgetTables_();
  ensureSchema_();
  return a;
}

/* ------------------------------------------------------------------ web app */

/* doGet / doPost (the app's API) live in Api.gs. */

/**
 * Everything the dashboard needs, in one call.
 * force = true (the refresh button) also drops the cached car entries so edits made straight in the car sheet show up.
 */
function getDashboardData(msg, force) {
  if (force) carForgetCache_();
  ensureSchema_();
  var model = buildModel_();
  return JSON.parse(JSON.stringify({
    msg: msg || '',
    account: ACCOUNT_.id,
    accountLabel: ACCOUNT_.label,
    hasBills: ACCOUNT_.bills,
    accounts: Object.keys(ACCOUNTS_).map(function (k) { return { id: k, label: ACCOUNTS_[k].label }; }),
    car: carConfigured_(),
    rules: model.rules,
    today: iso_(today_()),
    settings: model.settings,
    sheetUrl: ss_().getUrl(),
    bills: model.bills,
    forecast: model.forecast,
    attention: model.attention,
    recentTxns: model.recentTxns,
    addresses: model.addresses
  }));
}

/* ------------------------------------------------------------------ sheet io */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

// Each tab is read at most once per server call; any write through the helpers below clears it.
var TABLES_ = {};
function forgetTables_(name) { if (name) delete TABLES_[name]; else TABLES_ = {}; }

function readTable_(name) {
  if (TABLES_[name]) return TABLES_[name];
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing tab: ' + name);
  var values = sh.getDataRange().getValues();
  var headers = values.shift().map(function (h) { return String(h).trim(); });
  var rows = [];
  values.forEach(function (r, i) {
    if (r.every(function (v) { return v === '' || v === null; })) return;
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) { if (h) o[h] = r[j]; });
    rows.push(o);
  });
  return (TABLES_[name] = { sheet: sh, headers: headers, rows: rows });
}

function appendRows_(name, objs) {
  if (!objs.length) return;
  var t = readTable_(name);
  var data = objs.map(function (o) { return t.headers.map(function (h) { return o[h] === undefined ? '' : o[h]; }); });
  t.sheet.getRange(t.sheet.getLastRow() + 1, 1, data.length, t.headers.length).setValues(data);
  forgetTables_(name);
}

function setCell_(name, row, header, value) {
  var t = readTable_(name);
  var col = t.headers.indexOf(header) + 1;
  if (col < 1) throw new Error('Missing column ' + header + ' in ' + name);
  t.sheet.getRange(row, col).setValue(value);
  forgetTables_(name);
}

function readSettings_() {
  var t = readTable_(SHEETS.settings), s = {};
  t.rows.forEach(function (r) { s[String(r.Setting).trim()] = { value: r.Value, row: r._row }; });
  // Another account reads "<prefix>Setting" first. Its own balance/dates never fall back to the bills account's.
  var P = ACCOUNT_.prefix;
  function blank(v) { return v === '' || v === undefined || v === null; }
  function raw(k, own) {
    if (P && s[P + k] && !blank(s[P + k].value)) return s[P + k].value;
    if (P && own) return '';
    return s[k] ? s[k].value : '';
  }
  function num(k, def, own) { var v = raw(k, own); return (blank(v) || isNaN(Number(v))) ? def : Number(v); }
  return {
    account: ACCOUNT_.id,
    openingBalance: num('Opening Balance', 0, true),
    openingBalanceSet: !blank(raw('Opening Balance', true)),
    openingDate: toDate_(raw('Opening Date', true)) || today_(),
    months: num('Forecast Months', P ? 12 : 9, true),
    historyMonths: num('History Months', P ? 12 : 3, true),
    lowBalance: num('Low Balance Alert', 0, true),
    windowDays: num('Match Window Days', 5),
    tolerancePct: num('Amount Tolerance %', 10),
    estimateYears: num('Estimate Years', 3),
    amountOnly: String(s['Amount-only Matching'] ? s['Amount-only Matching'].value : 'Yes').toLowerCase().indexOf('n') !== 0,
    amountDays: num('Amount Match Days', 2),
    cashbackText: P ? String(raw('Cashback Text', true) || '') : s['Cashback Text'] ? String(s['Cashback Text'].value || '') : 'CASHBACK',
    folderId: s['Bills Folder ID'] ? String(s['Bills Folder ID'].value || '') : '',
    _rows: s
  };
}

/* ------------------------------------------------------------------ dates */

function today_() { var d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function toDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'string' && v) { var d = parseDate_(v); return d; }
  return null;
}
function iso_(d) {
  if (!d) return '';
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function fromIso_(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function addMonths_(d, n) {
  var y = d.getFullYear(), m = d.getMonth() + n, day = d.getDate();
  var last = new Date(y, m + 1, 0).getDate();
  return new Date(y, m, Math.min(day, last));
}
function daysBetween_(a, b) { return Math.round((b - a) / 86400000); }

var MONTHS_ = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
/** Parses the date formats Australian banks export: 15/11/2026, 15/11/26, 2026-11-15, 15 Nov 2026, 15 Nov 26, 15-Nov-2026. */
function parseDate_(s) {
  s = String(s).trim();
  var m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/))) return new Date(fullYear_(+m[3]), +m[2] - 1, +m[1]);
  if ((m = s.match(/^(\d{1,2})[\s-]([A-Za-z]{3})[A-Za-z]*[\s-](\d{2,4})$/)) && MONTHS_[m[2].toLowerCase()] !== undefined)
    return new Date(fullYear_(+m[3]), MONTHS_[m[2].toLowerCase()], +m[1]);
  return null;
}
function fullYear_(y) { return y < 100 ? 2000 + y : y; }

/* ------------------------------------------------------------------ model */

var FREQ_ = {
  'weekly': { days: 7 }, 'fortnightly': { days: 14 }, 'monthly': { months: 1 }, 'two-monthly': { months: 2 },
  'bi-monthly': { months: 2 }, 'quarterly': { months: 3 }, 'six-monthly': { months: 6 }, 'half-yearly': { months: 6 },
  'annual': { months: 12 }, 'yearly': { months: 12 }, 'once': { once: true }
};
function freqOf_(s) { return FREQ_[String(s || '').trim().toLowerCase()] || null; }
function periodDays_(f) { return f.days || (f.months ? Math.round(f.months * 30.44) : 30); }

function buildModel_() {
  var settings = readSettings_();
  var billRows = ACCOUNT_.bills ? readTable_(SHEETS.bills).rows : [], recT = readTable_(SHEETS.recurring);
  var txT = readTable_(SHEETS.txns), rcT = readTable_(SHEETS.rec);
  return computeModel_(settings, billRows, recT.rows, txT.rows, rcT.rows, today_(), ACCOUNT_.car ? carOtherForForecast_() : []);
}

/** Pure: everything below works on plain row objects so it can be tested off-Google. */
function computeModel_(settings, billRows, recRows, txRows, rcRows, today, carOther) {
  var bills = billRows.map(normBill_).filter(function (b) { return b.type; });
  var allTx = txRows.map(function (r) {
    return { id: String(r['Txn ID']), date: toDate_(r.Date), desc: String(r.Description || ''), amount: Number(r.Amount) || 0,
             balance: (r.Balance === '' || r.Balance === null || r.Balance === undefined) ? null : Number(r.Balance), order: r._row,
             imported: toDate_(r.Imported) };
  }).filter(function (t) { return t.date; });
  // Transactions before the Opening Date are kept only for the balance history on the graph; they're never reconciled.
  var txns = allTx.filter(function (t) { return t.date >= settings.openingDate; });
  var dayEnd = endOfDayBalances_(allTx);
  var lastImport = allTx.reduce(function (m, t) { return t.imported && (!m || t.imported > m) ? t.imported : m; }, null);

  var status = {};      // occurrence key -> {status, txnId}
  var txnUsed = {};     // txn id -> {status, occ}
  rcRows.forEach(function (r) {
    var occ = String(r.Occurrence || ''), tid = String(r['Txn ID'] || ''), st = String(r.Status || '').toLowerCase();
    if (occ && !/cashback/i.test(String(r.Note || ''))) status[occ] = { status: st, txnId: tid };
    if (tid) txnUsed[tid] = { status: st, occ: occ, note: String(r.Note || '') };
  });

  // Where the projection starts: last known bank balance, else opening balance + imported movements.
  var start = { date: settings.openingDate, balance: settings.openingBalance, source: 'opening' };
  if (txns.length) {
    var sorted = txns.slice().sort(function (a, b) { return a.date - b.date || a.order - b.order; });
    var last = sorted[sorted.length - 1];
    var lastDay = iso_(last.date);
    if (dayEnd.map[lastDay] !== undefined) {
      start = { date: last.date, balance: dayEnd.map[lastDay], source: 'bank' };
    } else {
      start = { date: last.date, balance: settings.openingBalance + sorted.reduce(function (s, t) { return s + t.amount; }, 0), source: 'opening+imports' };
    }
  } else if (allTx.length && !settings.openingBalanceSet && dayEnd.days.length) {
    // Nothing since the Opening Date and no Opening Balance typed in: start from the latest bank balance before it.
    var lastK = dayEnd.days[dayEnd.days.length - 1];
    start = { date: fromIso_(lastK), balance: dayEnd.map[lastK], source: 'bank' };
  }

  var horizon = addMonths_(today > start.date ? today : start.date, settings.months);
  var est = makeEstimator_(bills, settings.estimateYears, today);
  // Occurrences go back as far as the bank history, so older transactions can be shown against the rules too.
  var occFrom = allTx.reduce(function (m, t) { return t.date < m ? t.date : m; }, settings.openingDate);
  var occs = generateOccurrences_(recRows, occFrom, horizon, bills, est, carOther || []);

  occs.forEach(function (o) {
    var s = status[o.key];
    if (s && s.status === 'matched') { o.state = 'paid'; o.txnId = s.txnId; }
    else if (s && s.status === 'skipped') o.state = 'skipped';
    else if (o.bill && o.bill.paid && o.bill.paid >= settings.openingDate) o.state = 'paid';
    else if (o.bill && o.bill.paid && o.bill.paid < settings.openingDate) o.state = 'before-start';
    else if (o.date < settings.openingDate) o.state = 'before-start';
    else o.state = (start.source !== 'opening' && o.date <= start.date) || o.date < today ? 'overdue' : 'upcoming';
  });

  // Projection: start balance + everything not yet paid/skipped that falls after the start (overdue items land on day 1).
  var open = occs.filter(function (o) { return o.state === 'upcoming' || o.state === 'overdue'; });
  var day0 = start.date;
  var byDay = {};
  // An opening balance is taken as at the start of its day; a bank balance as at the end of its day,
  // so anything still outstanding lands on the following day.
  var first = start.source === 'opening' ? day0 : addDays_(day0, 1);
  open.forEach(function (o) {
    var d = o.date < first ? first : o.date;
    o.projDate = d;
    var k = iso_(d); (byDay[k] = byDay[k] || []).push(o);
  });
  var series = [], bal = start.balance, low = { balance: Infinity, date: null };
  for (var d = day0; d <= horizon; d = addDays_(d, 1)) {
    var k = iso_(d);
    (byDay[k] || []).forEach(function (o) { bal += o.amount; });
    bal = Math.round(bal * 100) / 100;
    series.push([k, bal]);
    if (d >= today && bal < low.balance) low = { balance: bal, date: k };
  }
  open.sort(function (a, b) { return a.projDate - b.projDate || a.amount - b.amount; });
  var run = start.balance;
  var upcoming = open.map(function (o) {
    run = Math.round((run + o.amount) * 100) / 100;
    return occOut_(o, run);
  });

  var matchedTxn = {};
  occs.forEach(function (o) { if (o.txnId) matchedTxn[o.txnId] = o; });
  var unmatchedTxns = txns.filter(function (t) { return !txnUsed[t.id]; })
    .sort(function (a, b) { return b.date - a.date; })
    .map(function (t) {
      var cands = occs.filter(function (o) {
        return (o.state === 'overdue' || o.state === 'upcoming') && (o.amount < 0) === (t.amount < 0) && Math.abs(daysBetween_(o.date, t.date)) <= 45;
      }).sort(function (a, b) { return Math.abs(a.date - t.date) - Math.abs(b.date - t.date); }).slice(0, 12)
        .map(function (o) { return { key: o.key, label: o.name + ' ' + iso_(o.date) + ' ' + o.amount.toFixed(2) }; });
      return { id: t.id, date: iso_(t.date), desc: t.desc, amount: t.amount, candidates: cands, suggest: suggestPhrase_(t.desc) };
    });

  var addresses = {};
  bills.forEach(function (b) { addresses[b.address] = Math.max(addresses[b.address] || 0, b.issued ? b.issued.getTime() : 0); });

  // History (before the Opening Date) is matched to the rules here, in memory only - never written to Reconcile.
  var histUsed = historyMatches_(occs, allTx.filter(function (t) { return t.date < settings.openingDate && !txnUsed[t.id]; }), settings);
  var used = {}; Object.keys(histUsed).forEach(function (k) { used[k] = histUsed[k]; }); Object.keys(txnUsed).forEach(function (k) { used[k] = txnUsed[k]; });
  var past = pastActuals_(occs, allTx, used, settings.openingDate);
  var rules = recRows.filter(function (r) { return String(r.Name || '').trim(); }).map(function (r) {
    var name = String(r.Name).trim(), n = 0;
    occs.forEach(function (o) { if (o.name === name && o.state === 'paid' && o.txnId) n++; });
    return ruleOut_(r, name, n, upcoming, past, carOther || []);
  });

  return {
    settings: { openingBalance: settings.openingBalance, openingDate: iso_(settings.openingDate), months: settings.months,
                lowBalance: settings.lowBalance, windowDays: settings.windowDays, tolerancePct: settings.tolerancePct,
                amountOnly: settings.amountOnly, amountDays: settings.amountDays },
    bills: bills.map(billOut_),
    addresses: Object.keys(addresses).sort(function (a, b) { return addresses[b] - addresses[a]; }),
    forecast: { start: { date: iso_(start.date), balance: start.balance, source: start.source },
                low: low.date ? low : null, series: series, upcoming: upcoming, horizon: iso_(horizon),
                history: balanceHistory_(dayEnd, addMonths_(today, -settings.historyMonths), start.date),
                lastImport: lastImport ? iso_(lastImport) : '', past: past },
    attention: { overdue: upcoming.filter(function (u) { return u.state === 'overdue'; }), unmatchedTxns: unmatchedTxns },
    recentTxns: txns.sort(function (a, b) { return b.date - a.date || b.order - a.order; }).slice(0, 40).map(function (t) {
      var o = matchedTxn[t.id], u = txnUsed[t.id];
      return { id: t.id, date: iso_(t.date), desc: t.desc, amount: t.amount, balance: t.balance,
               matched: o ? o.name : (u && u.status === 'ignored' ? '(one-off)' : u && u.status === 'matched' && u.occ ? u.occ.split('|')[0] : ''),
               how: u ? matchMethod_(u) : '' };
    }),
    rules: rules,
    _occs: occs, _txns: txns, _status: status, _txnUsed: txnUsed
  };
}

/**
 * Pure: what actually happened since the Opening Date, newest first. One row per bank transaction
 * (matched to an item, marked one-off, or not yet matched), plus bills marked paid without a transaction and
 * skipped items. Cashback credits are folded into the payment they were paired with.
 */
/**
 * Pure: pairs bank history from before the Opening Date with the rules' occurrences, using the same passes as
 * live matching (text + window + amount, exact amount, cashback) but a wider amount tolerance, since amounts
 * drift over time (e.g. a transfer that was $650 before it became $750). Returns {txnId: {status, occ, note}}.
 */
function historyMatches_(occs, histTx, settings) {
  if (!histTx.length) return {};
  var pool = occs.filter(function (o) { return o.state === 'before-start'; })
    .map(function (o) { var c = {}; for (var k in o) c[k] = o[k]; c.state = 'history'; return c; });
  var s2 = {}; for (var k in settings) s2[k] = settings[k];
  s2.tolerancePct = Math.max(settings.tolerancePct, 40);
  var out = {};
  proposeMatches_({ _occs: pool, _txns: histTx, _txnUsed: {} }, s2).forEach(function (p) {
    out[p.txnId] = { status: 'matched', occ: p.key, note: 'history: ' + p.how };
  });
  return out;
}

function pastActuals_(occs, txns, txnUsed, openingDate) {
  var byKey = {}; occs.forEach(function (o) { byKey[o.key] = o; });
  var cb = {}, out = [], seenOcc = {};
  txns.forEach(function (t) {
    var u = txnUsed[t.id];
    if (u && u.status === 'matched' && u.occ && /cashback/i.test(u.note)) cb[u.occ] = Math.round(((cb[u.occ] || 0) + t.amount) * 100) / 100;
  });
  txns.forEach(function (t) {
    var u = txnUsed[t.id], hist = openingDate && t.date < openingDate;
    var row = { date: iso_(t.date), amount: t.amount, balance: t.balance, desc: t.desc, txnId: t.id, order: t.order };
    if (hist) row.history = true;
    if (!u) { row.state = hist ? 'unlisted' : 'unmatched'; row.name = t.desc; }
    else if (u.status === 'ignored') { row.state = 'one-off'; row.name = t.desc; }
    else if (u.status === 'matched' && u.occ && !/cashback/i.test(u.note)) {
      var o = byKey[u.occ];
      row.state = 'paid'; row.key = u.occ; row.how = /^history/.test(u.note) ? 'history' : matchMethod_(u);
      row.name = o ? o.name : u.occ.split('|')[0];
      row.due = o ? iso_(o.date) : u.occ.split('|')[1] || '';
      row.expected = o ? o.amount : null;
      row.source = o ? o.source : '';
      row.cashback = cb[u.occ] || 0;
      seenOcc[u.occ] = true;
    } else return;
    out.push(row);
  });
  occs.forEach(function (o) {
    if (seenOcc[o.key]) return;
    if (o.state === 'paid' && o.bill && !o.txnId)
      out.push({ date: iso_(o.bill.paid || o.date), amount: -o.bill.total, balance: null, state: 'paid', how: 'bill', key: o.key,
                 name: o.name, due: iso_(o.date), expected: o.amount, source: o.source, order: 0 });
    else if (o.state === 'skipped')
      out.push({ date: iso_(o.date), amount: null, balance: null, state: 'skipped', key: o.key, name: o.name, due: iso_(o.date),
                 expected: o.amount, source: o.source, order: 0 });
  });
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : (b.order || 0) - (a.order || 0); });
  out.forEach(function (r) { delete r.order; });
  return out;
}

/**
 * Pure: one Recurring row as the app's Rules tab shows it - how it's forecast, what's next, and how its last
 * actual payments compared with what was expected (amount and days early/late).
 */
function ruleOut_(r, name, matchedCount, upcoming, past, carOther) {
  var f = freqOf_(r.Frequency), billTypes = String(r['Bill Type'] || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  var carLink = String(r['Car Item'] || '').trim();
  var hasCar = !!carLink && carOther.some(function (e) { return carItemMatches_(e.item, carLink); });
  var kind = !f ? 'invalid' : !toDate_(r['Next Date']) ? 'invalid' : f.once ? 'once' : hasCar ? 'car' : billTypes.length ? 'estimate' : 'fixed';
  var next = upcoming.filter(function (u) { return u.name === name; })[0] || null;
  var recent = past.filter(function (p) { return p.name === name && p.state === 'paid' && p.amount !== null; }).slice(0, 6).map(function (p) {
    return { due: p.due, date: p.date, expected: p.expected, actual: p.amount, cashback: p.cashback || 0,
             days: p.due ? daysBetween_(fromIso_(p.due), fromIso_(p.date)) : null };
  });
  var withExp = recent.filter(function (x) { return x.expected !== null && x.expected !== undefined; });
  // Typical (median) difference, so one unusual payment doesn't make a rule look out of date.
  var avg = function (a, g) {
    if (!a.length) return null;
    var v = a.map(g).sort(function (x, y) { return x - y; }), m = Math.floor(v.length / 2);
    return Math.round((v.length % 2 ? v[m] : (v[m - 1] + v[m]) / 2) * 100) / 100;
  };
  var end = toDate_(r['End Date']), nd = toDate_(r['Next Date']);
  return {
    name: name, matchText: String(r['Match Text'] || ''), learned: splitPhrases_(r['Learned Text']),
    active: String(r.Active || 'Yes').toLowerCase().indexOf('n') !== 0, matched: matchedCount,
    amount: Number(r.Amount) || 0, frequency: String(r.Frequency || ''), nextDate: iso_(nd), endDate: iso_(end),
    skipMonths: String(r['Skip Months'] || ''), billTypes: billTypes, carItem: carLink, notes: String(r.Notes || ''), kind: kind,
    next: next ? { date: next.date, amount: next.amount, source: next.source, state: next.state } : null,
    recent: recent,
    avgDiff: avg(withExp, function (x) { return Math.abs(x.actual) - Math.abs(x.expected); }),
    avgDays: avg(recent.filter(function (x) { return x.days !== null; }), function (x) { return x.days; })
  };
}

function occOut_(o, run) {
  return { key: o.key, name: o.name, date: iso_(o.projDate || o.date), due: iso_(o.date), amount: o.amount,
           source: o.source, state: o.state, balance: run, billRow: o.bill ? o.bill.row : null,
           pdf: o.bill ? o.bill.pdf : (o.receipt || ''), note: o.note || '' };
}

function normBill_(r) {
  var from = toDate_(r['Period From']), to = toDate_(r['Period To']);
  var days = Number(r.Days) || (from && to ? daysBetween_(from, to) + 1 : 0);
  return {
    row: r._row, type: String(r['Bill Type'] || '').trim(), biller: String(r.Biller || ''), invoice: String(r['Invoice No.'] || ''),
    issued: toDate_(r['Date Issued']), due: toDate_(r['Date Due']), paid: toDate_(r['Date Paid']),
    from: from, to: to, days: days, qty: Number(r.Quantity) || 0, unit: String(r.Unit || ''),
    subtotal: Number(r.Subtotal) || 0, discount: Number(r.Discount) || 0, total: Number(r.Total) || 0,
    address: String(r['Supply Address'] || ''), pdf: String(r['Bill PDF'] || ''), notes: String(r.Notes || '')
  };
}
function billOut_(b) {
  return { row: b.row, type: b.type, biller: b.biller, invoice: b.invoice, issued: iso_(b.issued), due: iso_(b.due), paid: iso_(b.paid),
           from: iso_(b.from), to: iso_(b.to), days: b.days, qty: b.qty, unit: b.unit, total: b.total, discount: b.discount,
           address: b.address, pdf: b.pdf, notes: b.notes,
           perDay: b.days ? b.total / b.days : null, usePerDay: b.days ? b.qty / b.days : null, unitRate: b.qty ? b.total / b.qty : null };
}

/**
 * Seasonal estimator: average cost per day for each calendar month, from the last N years of bills
 * of that type at the address of the latest bill of that type.
 */
function makeEstimator_(bills, years, today) {
  var cache = {};
  return function (type, periodEnd, days) {
    var key = type;
    if (!cache[key]) {
      var ofType = bills.filter(function (b) { return b.type === type && b.from && b.to && b.days > 0; });
      if (!ofType.length) { cache[key] = null; }
      else {
        var latest = ofType.reduce(function (a, b) { return (b.to > a.to) ? b : a; });
        var cutoff = addMonths_(latest.to, -12 * years);
        var cost = [], dcount = [];
        for (var m = 0; m < 12; m++) { cost[m] = 0; dcount[m] = 0; }
        ofType.filter(function (b) { return b.address === latest.address && b.to > cutoff; }).forEach(function (b) {
          var perDay = b.total / b.days;
          for (var d = b.from; d <= b.to; d = addDays_(d, 1)) { cost[d.getMonth()] += perDay; dcount[d.getMonth()]++; }
        });
        var overall = cost.reduce(function (a, b) { return a + b; }, 0) / Math.max(1, dcount.reduce(function (a, b) { return a + b; }, 0));
        var shape = cost.map(function (c, m) { return dcount[m] ? c / dcount[m] : overall; });
        // Seasonal shape from N years, scaled to the last 12 months' actual level (prices drift up).
        var yearAgo = addDays_(latest.to, -365), actual = 0, expected = 0;
        ofType.filter(function (b) { return b.address === latest.address && b.to > yearAgo; }).forEach(function (b) {
          var perDay = b.total / b.days;
          for (var d2 = b.from; d2 <= b.to; d2 = addDays_(d2, 1)) if (d2 > yearAgo) { actual += perDay; expected += shape[d2.getMonth()]; }
        });
        var scale = expected > 0 ? actual / expected : 1;
        cache[key] = shape.map(function (v) { return v * scale; });
      }
    }
    var perMonth = cache[key];
    if (!perMonth) return null;
    var total = 0;
    for (var i = 0; i < days; i++) total += perMonth[addDays_(periodEnd, -i).getMonth()];
    return Math.round(total * 100) / 100;
  };
}

/**
 * Pure: end-of-day bank balance for each day with transactions. Within a day the file order can't be trusted
 * (banks list newest first), so the day's closing balance is the one no other transaction that day leads on from.
 */
function endOfDayBalances_(txs) {
  var byDay = {};
  txs.forEach(function (t) { if (t.balance !== null && !isNaN(t.balance)) (byDay[iso_(t.date)] = byDay[iso_(t.date)] || []).push(t); });
  var map = {}, open = {};
  Object.keys(byDay).forEach(function (k) {
    var list = byDay[k], before = {};
    list.forEach(function (t) { before[(t.balance - t.amount).toFixed(2)] = true; });
    var closers = list.filter(function (t) { return !before[t.balance.toFixed(2)]; });
    var closer = closers.length === 1 ? closers[0] : list.slice().sort(function (a, b) { return b.order - a.order; })[0];
    map[k] = Math.round(closer.balance * 100) / 100;
    var openers = list.filter(function (t) { return !list.some(function (u) { return u !== t && Math.abs(u.balance - (t.balance - t.amount)) < 0.005; }); });
    var first = openers.length === 1 ? openers[0] : list[0];
    open[k] = Math.round((first.balance - first.amount) * 100) / 100;
  });
  return { map: map, open: open, days: Object.keys(map).sort() };
}

/** Pure: daily closing balance from `from` to `to`, carrying the last known balance across quiet days. */
function balanceHistory_(dayEnd, from, to) {
  if (!dayEnd.days.length || from > to) return [];
  var out = [], fk = iso_(from), prev = null, i = 0, days = dayEnd.days;
  while (i < days.length && days[i] < fk) { prev = dayEnd.map[days[i]]; i++; }
  if (prev === null && i < days.length) prev = dayEnd.open[days[i]];
  if (prev === null) return [];
  for (var d = from; d <= to; d = addDays_(d, 1)) {
    var k = iso_(d);
    if (dayEnd.map[k] !== undefined) prev = dayEnd.map[k];
    out.push([k, prev]);
  }
  return out;
}

/** Pure: does a car "Other" item match a Car Item pattern? Exact, case-insensitive; | separates alternatives. */
function carItemMatches_(item, pattern) {
  var it = String(item || '').trim().toLowerCase();
  return String(pattern || '').split('|').some(function (p) { return p.trim().toLowerCase() === it && it !== ''; });
}

/**
 * Expands each Recurring row into dated occurrences between `from` and `to`, then swaps in real bills.
 * Rows with a Car Item use the car sheet instead: each real entry becomes an occurrence with its real
 * amount and date, and the schedule continues from the latest entry at the latest amount.
 */
function generateOccurrences_(recRows, from, to, bills, est, carOther) {
  var out = [];
  var BILL_LAG = 18; // a bill is due roughly this many days after its period ends
  recRows.forEach(function (r) {
    var name = String(r.Name || '').trim();
    if (!name || String(r.Active || 'Yes').toLowerCase().indexOf('n') === 0) return;
    var f = freqOf_(r.Frequency), anchor = toDate_(r['Next Date']);
    if (!f || !anchor) return;
    var end = toDate_(r['End Date']);
    var skip = String(r['Skip Months'] || '').split(/[,\s]+/).filter(String).map(Number);
    var amount = Number(r.Amount) || 0;
    var billTypes = String(r['Bill Type'] || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    var dates = [];
    var lo = addDays_(from, -40);
    var match = [String(r['Match Text'] || name), String(r['Learned Text'] || '')].join('|');

    var carLink = String(r['Car Item'] || '').trim();
    var ents = carLink && !f.once ? (carOther || []).filter(function (e) { return carItemMatches_(e.item, carLink); })
      .sort(function (a, b) { return a.date - b.date; }) : [];
    if (ents.length) {
      var latest = ents[ents.length - 1], carOut = [];
      ents.forEach(function (e) {
        if (e.date < lo || e.date > to) return;
        carOut.push({ key: name + '|' + iso_(e.date), name: name, date: e.date, sched: e.date, amount: -e.total, source: 'car',
                      match: match, billTypes: [], freq: f, receipt: e.receipt, note: 'Car: ' + e.item });
      });
      // Next one is a period after the latest car entry - or the Recurring tab's Next Date if that's later
      // (e.g. a service that's overdue by the schedule but booked for a set day).
      var sched1 = f.days ? addDays_(latest.date, f.days) : addMonths_(latest.date, f.months);
      var useNext = anchor > sched1, from0 = useNext ? anchor : latest.date, j0 = useNext ? 0 : 1;
      for (var j = j0; j < 400; j++) {
        var nd = f.days ? addDays_(from0, j * f.days) : addMonths_(from0, j * f.months);
        if (nd > to || (end && nd > end)) break;
        if (nd < lo || skip.indexOf(nd.getMonth() + 1) >= 0) continue;
        carOut.push({ key: name + '|' + iso_(nd), name: name, date: nd, sched: nd, amount: -latest.total, source: 'recurring',
                      match: match, billTypes: [], freq: f, note: useNext ? 'Next date from the Recurring tab' : 'Next after car entry ' + iso_(latest.date) });
      }
      out = out.concat(carOut);
      return;
    }
    if (f.once) dates = [anchor];
    else {
      for (var k = -400; k < 800; k++) {
        var d = f.days ? addDays_(anchor, k * f.days) : addMonths_(anchor, k * f.months);
        if (d < lo) continue;
        if (d > to) break;
        dates.push(d);
      }
    }
    var mine = [];
    dates.forEach(function (d) {
      if (end && d > end) return;
      if (skip.indexOf(d.getMonth() + 1) >= 0) return;
      var o = { key: name + '|' + iso_(d), name: name, date: d, sched: d, amount: amount, source: 'recurring',
                match: match, billTypes: billTypes, freq: f };
      if (billTypes.length) {
        var e = 0, ok = true;
        billTypes.forEach(function (t) { var v = est(t, addDays_(d, -BILL_LAG), periodDays_(f)); if (v === null) ok = false; else e += v; });
        if (ok) { o.amount = -e; o.source = 'estimate'; }
      }
      mine.push(o);
    });
    // Real bills replace the nearest occurrence (within half a period of the due date).
    if (billTypes.length && !f.once) {
      var half = periodDays_(f) / 2;
      bills.filter(function (b) { return billTypes.indexOf(b.type) >= 0 && b.due && b.due >= addDays_(from, -60); })
        .sort(function (a, b) { return a.due - b.due; })
        .forEach(function (b) {
          var best = null, bd = Infinity;
          mine.forEach(function (o) {
            if (o.bill) return;
            var dd = Math.abs(daysBetween_(o.sched, b.due));
            if (dd <= half && dd < bd) { bd = dd; best = o; }
          });
          if (best) {
            best.bill = b; best.amount = -b.total; best.date = b.due; best.source = 'bill';
            best.note = b.biller + ' ' + iso_(b.from) + ' to ' + iso_(b.to);
          }
        });
    }
    out = out.concat(mine);
  });
  return out.filter(function (o) { return o.date >= addDays_(from, -40) && o.date <= to; });
}

/* ------------------------------------------------------------------ reconciliation */

/** Upper-case, punctuation to spaces, single-spaced - so "1st-Energy*DD" and "1ST ENERGY DD" compare equal. */
function norm_(s) { return ' ' + String(s || '').toUpperCase().replace(/[^A-Z0-9&]+/g, ' ').trim() + ' '; }
function splitPhrases_(s) { return String(s || '').split('|').map(function (p) { return p.trim(); }).filter(String); }

function matchText_(desc, pattern) {
  var d = norm_(desc);
  return splitPhrases_(pattern).some(function (p) { var n = norm_(p).trim(); return n && d.indexOf(n) >= 0; });
}

// Words banks add to descriptions that say nothing about who was paid.
var NOISE_ = ('DD DIRECT DEBIT CREDIT BPAY VISA MASTERCARD EFTPOS PURCHASE CARD PAYMENT PYMT PMT POS AUS AU AUST AUSTRALIA NSW VIC ' +
  'QLD SA WA TAS ACT NT PTY LTD LIMITED INTERNET BANKING NETBANK ONLINE OSKO NPP REF REFERENCE RECEIPT TFR VALUE DATE AUD ' +
  'CBA SQ THE AND FROM TO FOR OF ON AT BY WITH NO TRANSACTION DEPOSIT WITHDRAWAL DEBITS CREDITS PAY TAP CONTACTLESS APPLE GOOGLE').split(' ');

/** Pure: picks the words in a bank description that name the payee, e.g. "Direct Debit 123456 AGL SALES PTY LTD" -> "AGL SALES". */
function suggestPhrase_(desc) {
  var toks = norm_(desc).trim().split(' ');
  function keep(t) {
    if (NOISE_.indexOf(t) >= 0) return false;
    var letters = (t.match(/[A-Z]/g) || []).length, digits = (t.match(/[0-9]/g) || []).length;
    if (letters < 2) return false;              // 123456, 4521751, X1
    if (digits > letters) return false;         // reference codes like A12345
    if (/^X+\d*$/.test(t)) return false;        // masked account numbers
    return true;
  }
  var out = [];
  for (var i = 0; i < toks.length; i++) {
    if (keep(toks[i])) { out.push(toks[i]); if (out.length === 3) break; }
    else if (out.length) break;                 // stop at the first noise word after the payee starts
  }
  var p = out.join(' ');
  return p.length >= 3 ? p : '';
}

function matchMethod_(u) {
  if (u.status === 'ignored') return 'one-off';
  var n = String(u.note || '').toLowerCase();
  if (n.indexOf('cashback') >= 0) return 'cashback';
  if (n.indexOf('amount') >= 0) return 'amount';
  if (n.indexOf('manual') >= 0) return 'manual';
  return 'text';
}

/** Pure: proposes matches for unmatched transactions. Returns [{txnId, key, occ, txn, how}] */
/**
 * How far (in days) a payment is from an occurrence's allowed window; 0 = inside it.
 * A real bill can be paid any time from its issue date until a week or so after it's due.
 * An estimated utility bill can be paid up to three weeks before its estimated due date, or a week or so after.
 */
function windowGap_(o, t, settings) {
  var w = settings.windowDays, lo, hi;
  if (o.bill) { lo = o.bill.issued ? addDays_(o.bill.issued, -1) : addDays_(o.date, -21); hi = addDays_(o.date, w + 7); }
  else if (o.billTypes && o.billTypes.length) { lo = addDays_(o.date, -21); hi = addDays_(o.date, w + 7); }
  else { lo = addDays_(o.date, -w); hi = addDays_(o.date, w); }
  if (t.date < lo) return daysBetween_(t.date, lo);
  if (t.date > hi) return daysBetween_(hi, t.date);
  return 0;
}

function proposeMatches_(model, settings) {
  var taken = {};
  model._occs.forEach(function (o) { if (o.state === 'paid' || o.state === 'skipped') taken[o.key] = true; });
  var out = [], left = [];
  var free = model._txns.filter(function (t) { return !model._txnUsed[t.id]; }).sort(function (a, b) { return a.date - b.date; });

  // Pass 1: description text + amount within tolerance + date within window.
  free.forEach(function (t) {
    var best = null, score = Infinity;
    model._occs.forEach(function (o) {
      if (taken[o.key] || o.state === 'before-start') return;
      if ((o.amount < 0) !== (t.amount < 0)) return;
      if (!matchText_(t.desc, o.match)) return;
      if (windowGap_(o, t, settings) > 0) return;
      var dd = Math.abs(daysBetween_(o.date, t.date)) / 10;   // inside the window, nearer is still better
      var tol = Math.abs(o.amount) * settings.tolerancePct / 100 * (o.source === 'estimate' ? 3 : 1);
      var diff = Math.abs(Math.abs(t.amount) - Math.abs(o.amount));
      if (diff > Math.max(tol, 2)) return;
      var s = dd + diff / Math.max(1, Math.abs(o.amount)) * 10;
      if (s < score) { score = s; best = o; }
    });
    if (best) { taken[best.key] = true; out.push({ txnId: t.id, key: best.key, occ: best, txn: t, how: 'text' }); }
    else left.push(t);
  });

  // Pass 2: exact amount, close date, and only one possible pairing either way. Never for estimates.
  if (settings.amountOnly) {
    var pairs = [];
    left.forEach(function (t) {
      model._occs.forEach(function (o) {
        if (taken[o.key] || o.state === 'before-start' || o.source === 'estimate') return;
        if (Math.abs(t.amount - o.amount) >= 0.005) return;
        if (Math.abs(daysBetween_(o.date, t.date)) > settings.amountDays) return;
        pairs.push({ t: t, o: o });
      });
    });
    var perT = {}, perO = {};
    pairs.forEach(function (p) { perT[p.t.id] = (perT[p.t.id] || 0) + 1; perO[p.o.key] = (perO[p.o.key] || 0) + 1; });
    pairs.forEach(function (p) {
      if (perT[p.t.id] === 1 && perO[p.o.key] === 1) { taken[p.o.key] = true; out.push({ txnId: p.t.id, key: p.o.key, occ: p.o, txn: p.t, how: 'amount' }); }
    });
  }

  // Pass 3: small cashback credits paired with a utility payment made the same day (or the day either side).
  var cbText = settings.cashbackText;
  if (cbText) {
    var byKey = {}; model._occs.forEach(function (o) { byKey[o.key] = o; });
    var pays = [], hasCb = {};
    model._txns.forEach(function (t) {
      var u = model._txnUsed[t.id];
      if (!u || u.status !== 'matched' || !u.occ) return;
      if (/cashback/i.test(u.note)) { hasCb[u.occ] = true; return; }
      var o = byKey[u.occ]; if (o && o.billTypes && o.billTypes.length) pays.push({ t: t, o: o });
    });
    out.forEach(function (p) { if (p.occ.billTypes && p.occ.billTypes.length) pays.push({ t: p.txn, o: p.occ }); });
    var paired = {}; out.forEach(function (p) { paired[p.txnId] = true; });
    free.forEach(function (t) {
      if (paired[t.id] || t.amount <= 0 || t.amount > 20 || !matchText_(t.desc, cbText)) return;
      var best = null;
      pays.forEach(function (p) {
        if (hasCb[p.o.key] || Math.abs(daysBetween_(p.t.date, t.date)) > 1) return;
        if (!best || Math.abs(p.t.date - t.date) < Math.abs(best.t.date - t.date)) best = p;
      });
      if (best) { hasCb[best.o.key] = true; paired[t.id] = true; out.push({ txnId: t.id, key: best.o.key, occ: best.o, txn: t, how: 'cashback' }); }
    });
  }
  return out;
}

function autoMatch() {
  var settings = readSettings_();
  var model = buildModel_();
  var props = proposeMatches_(model, settings);
  var now = new Date();
  appendRows_(SHEETS.rec, props.map(function (p) {
    return { Occurrence: p.key, 'Txn ID': p.txnId, Status: 'matched', Note: 'auto: ' + p.how, Updated: now };
  }));
  props.forEach(function (p) { if (p.how !== 'cashback') markBillPaid_(p.occ, p.txn.date); });
  return props.length;
}

/** Re-runs matching over everything unmatched - use after changing rules. */
function runMatching() {
  var n = autoMatch();
  return getDashboardData(n ? n + ' transaction' + (n === 1 ? '' : 's') + ' matched' : 'No new matches');
}

function markBillPaid_(occ, date) {
  if (occ && occ.bill && !occ.bill.paid) setCell_(SHEETS.bills, occ.bill.row, 'Date Paid', date);
}

/**
 * Manual match. If `phrase` is given and the item's rules don't already recognise this description,
 * the phrase is added to the item's Learned Text so next time it matches by itself.
 */
function matchTxn(txnId, occKey, phrase) {
  var model = buildModel_();
  var occ = model._occs.filter(function (o) { return o.key === occKey; })[0];
  var txn = model._txns.filter(function (t) { return t.id === txnId; })[0];
  if (!occ || !txn) throw new Error('Could not find that item - refresh and try again.');
  var learned = '';
  phrase = String(phrase || '').trim().toUpperCase();
  if (phrase) {
    if (!matchText_(txn.desc, phrase)) throw new Error('"' + phrase + '" isn\'t in the description "' + txn.desc + '" - shorten it to words that are.');
    if (!matchText_(txn.desc, occ.match) && addLearned_(occ.name, phrase)) learned = phrase;
  }
  appendRows_(SHEETS.rec, [{ Occurrence: occKey, 'Txn ID': txnId, Status: 'matched',
                             Note: 'manual' + (learned ? '; learned: ' + learned : ''), Updated: new Date() }]);
  markBillPaid_(occ, txn.date);
  var more = learned ? autoMatch() : 0;
  return getDashboardData('Matched to ' + occ.name + (learned ? '. Learned "' + learned + '"' : '') +
                          (more ? ' - and ' + more + ' more matched with it' : ''));
}

function ignoreTxn(txnId, note) {
  appendRows_(SHEETS.rec, [{ Occurrence: '', 'Txn ID': txnId, Status: 'ignored', Note: note || '', Updated: new Date() }]);
  return getDashboardData('Marked as a one-off');
}

function skipOccurrence(occKey, note) {
  appendRows_(SHEETS.rec, [{ Occurrence: occKey, 'Txn ID': '', Status: 'skipped', Note: note || '', Updated: new Date() }]);
  return getDashboardData('Skipped');
}

/**
 * Removes the Reconcile row(s) for an occurrence or transaction, putting it back in play.
 * Undoing a manual match also forgets the phrase that match taught.
 */
function undoReconcile(id) {
  var t = readTable_(SHEETS.rec), forgot = [];
  var rows = t.rows.filter(function (r) { return String(r.Occurrence) === id || String(r['Txn ID']) === id; });
  rows.forEach(function (r) {
    var m = String(r.Note || '').match(/learned: (.+)$/);
    if (m) { var name = String(r.Occurrence).split('|')[0]; if (removeLearned_(name, m[1])) forgot.push(m[1]); }
  });
  rows.map(function (r) { return r._row; }).sort(function (a, b) { return b - a; })
    .forEach(function (row) { t.sheet.deleteRow(row); });
  forgetTables_(SHEETS.rec);
  return getDashboardData('Undone' + (forgot.length ? ' and forgot "' + forgot.join('", "') + '"' : ''));
}

/* ------------------------------------------------------------------ matching rules */

function recurringRow_(name) {
  var t = readTable_(SHEETS.recurring);
  var r = t.rows.filter(function (x) { return String(x.Name).trim() === name; })[0];
  if (!r) throw new Error('No Recurring item called "' + name + '"');
  return r;
}
function addLearned_(name, phrase) {
  var r = recurringRow_(name), list = splitPhrases_(r['Learned Text']);
  phrase = String(phrase).trim().toUpperCase();
  if (!phrase || list.some(function (p) { return norm_(p) === norm_(phrase); })) return false;
  list.push(phrase);
  setCell_(SHEETS.recurring, r._row, 'Learned Text', list.join('|'));
  return true;
}
function removeLearned_(name, phrase) {
  var r; try { r = recurringRow_(name); } catch (e) { return false; }
  var list = splitPhrases_(r['Learned Text']), keep = list.filter(function (p) { return norm_(p) !== norm_(phrase); });
  if (keep.length === list.length) return false;
  setCell_(SHEETS.recurring, r._row, 'Learned Text', keep.join('|'));
  return true;
}

function addLearnedPhrase(name, phrase) {
  phrase = String(phrase || '').trim();
  if (norm_(phrase).trim().length < 3) throw new Error('Use at least 3 letters');
  var added = addLearned_(name, phrase);
  var n = added ? autoMatch() : 0;
  return getDashboardData(added ? 'Added "' + phrase.toUpperCase() + '" to ' + name + (n ? ' - ' + n + ' matched' : '') : 'Already there');
}
function removeLearnedPhrase(name, phrase) {
  removeLearned_(name, phrase);
  return getDashboardData('Removed "' + phrase + '" from ' + name + '. Past matches are kept.');
}
/**
 * Adds a forecast item to the current account's Recurring tab - a one-off planned payment (Frequency "Once")
 * or a repeating one. form: {name, amount (− out / + in), frequency, date, end, matchText, notes}
 */
function addForecastItem(form) {
  form = form || {};
  var name = String(form.name || '').trim(), amount = Number(form.amount), freq = String(form.frequency || 'Once').trim();
  if (!name) throw new Error('Give it a name.');
  if (!amount || isNaN(amount)) throw new Error('Enter an amount.');
  if (!freqOf_(freq)) throw new Error('Unknown frequency: ' + freq);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(form.date || ''))) throw new Error('Choose a date.');
  var f = freqOf_(freq);
  if (!f.once && readTable_(SHEETS.recurring).rows.some(function (r) { return String(r.Name).trim() === name; }))
    throw new Error('There is already an item called "' + name + '" - use another name, or change that one in the sheet.');
  appendRows_(SHEETS.recurring, [{
    Name: name, Amount: Math.round(amount * 100) / 100, Frequency: freq, 'Next Date': fromIso_(form.date),
    'End Date': form.end ? fromIso_(form.end) : '', 'Match Text': String(form.matchText || '').trim(), Active: 'Yes',
    Notes: String(form.notes || '').trim()
  }]);
  var n = autoMatch();
  return getDashboardData('Added ' + name + (n ? ' - ' + n + ' matched' : ''));
}

function setMatchText(name, text) {
  var r = recurringRow_(name);
  setCell_(SHEETS.recurring, r._row, 'Match Text', String(text || '').trim());
  var n = autoMatch();
  return getDashboardData('Saved ' + name + (n ? ' - ' + n + ' matched' : ''));
}

/**
 * Adds columns/settings introduced after the workbook was first set up, so old sheets keep working.
 * Runs its checks once per schema version, then remembers it's done.
 */
var SCHEMA_VERSION_ = '6';
var REC_HEADERS_ = ['Name', 'Amount', 'Frequency', 'Next Date', 'End Date', 'Skip Months', 'Match Text', 'Learned Text', 'Bill Type', 'Car Item', 'Active', 'Notes'];
var TXN_HEADERS_ = ['Txn ID', 'Date', 'Description', 'Amount', 'Balance', 'Imported'];
var RECON_HEADERS_ = ['Occurrence', 'Txn ID', 'Status', 'Note', 'Updated'];
function ensureTab_(name, headers) {
  var sh = ss_().getSheetByName(name);
  if (sh) return sh;
  sh = ss_().insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}
function ensureSchema_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION_) return;
  // Two requests arriving together must not both upgrade (duplicate rows, "sheet already exists").
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    forgetTables_();
    if (props.getProperty('SCHEMA_VERSION') !== SCHEMA_VERSION_) upgradeSchema_(props);
  } finally { lock.releaseLock(); }
}
function upgradeSchema_(props) {
  // v6: tabs for the other accounts (e.g. Holiday), each with its own Recurring / Transactions / Reconcile.
  Object.keys(ACCOUNTS_).forEach(function (k) {
    var a = ACCOUNTS_[k];
    if (a.id === 'bills') return;
    ensureTab_(a.recurring, REC_HEADERS_); ensureTab_(a.txns, TXN_HEADERS_); ensureTab_(a.rec, RECON_HEADERS_);
  });
  var rec = ss_().getSheetByName(ACCOUNTS_.bills.recurring);
  var head = rec.getRange(1, 1, 1, rec.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.indexOf('Learned Text') < 0) {
    var after = head.indexOf('Match Text') + 1 || head.length;
    rec.insertColumnAfter(after);
    rec.getRange(1, after + 1).setValue('Learned Text');
    head.splice(after, 0, 'Learned Text');
  }
  if (head.indexOf('Car Item') < 0) {
    var after2 = head.indexOf('Bill Type') + 1 || head.length;
    rec.insertColumnAfter(after2);
    rec.getRange(1, after2 + 1).setValue('Car Item');
  }
  var st = ss_().getSheetByName(SHEETS.settings);
  var keys = st.getDataRange().getValues().map(function (r) { return String(r[0]).trim(); });
  [['Amount-only Matching', 'Yes', 'Yes = when the description doesn\'t match, still match a transaction whose amount is exactly right and date is close, if it can only be one item. Never used for estimates.'],
   ['Amount Match Days', 2, 'How many days either side of the due date an amount-only match can be.'],
   ['Car Sheet ID', '', 'ID (or full link) of your Vehicle Expenses Google Sheet. Leave blank to hide the Car section.'],
   ['History Months', 3, 'How many months of actual bank balance to show before today on the graph.'],
   ['Cashback Text', 'CASHBACK', 'Small credits (under $20) with this in the description are paired with a utility payment made the same day. Blank = off.'],
   ['Holiday Opening Balance', '', 'Holiday account. Leave blank to start from the latest imported bank balance.'],
   ['Holiday Opening Date', today_(), 'Holiday account. Transactions from this date are reconciled; earlier ones only draw the balance history.'],
   ['Holiday Forecast Months', 12, 'Holiday account. How far ahead to forecast.'],
   ['Holiday History Months', 12, 'Holiday account. Months of actual balance to show (and keep on import) before today.'],
   ['Holiday Low Balance Alert', 0, 'Holiday account. Warn when the forecast drops below this.']]
    .forEach(function (row) { if (keys.indexOf(row[0]) < 0) st.appendRow(row); });
  forgetTables_();
  props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION_);
}

/* ------------------------------------------------------------------ CSV import */

function parseCsv_(text) {
  var rows = [], row = [], field = '', q = false;
  text = String(text).replace(/^﻿/, '');
  for (var i = 0; i < text.length; i++) {
    var c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      if (row.some(function (v) { return v.trim() !== ''; })) rows.push(row);
      row = [];
    } else field += c;
  }
  row.push(field);
  if (row.some(function (v) { return v.trim() !== ''; })) rows.push(row);
  return rows;
}

function parseAmount_(s) {
  s = String(s || '').trim();
  if (!s) return null;
  var neg = /^\(.*\)$/.test(s) || /-/.test(s) || /\bDR\b/i.test(s);
  var n = parseFloat(s.replace(/[^0-9.]/g, ''));
  if (isNaN(n)) return null;
  return neg ? -n : n;
}

/** Pure: turns a bank CSV into [{date, desc, amount, balance}], working out the columns itself. */
function parseBankCsv_(text) {
  var rows = parseCsv_(text);
  if (!rows.length) throw new Error('The file is empty.');
  var head = rows[0].map(function (h) { return h.trim().toLowerCase(); });
  var hasHeader = !parseDate_(rows[0][0]) && head.some(function (h) { return /date/.test(h); });
  var col = { date: -1, desc: -1, amount: -1, debit: -1, credit: -1, balance: -1 };
  var data = hasHeader ? rows.slice(1) : rows;
  if (hasHeader) {
    head.forEach(function (h, i) {
      if (col.date < 0 && /date/.test(h) && !/process|value/.test(h)) col.date = i;
      else if (/balance/.test(h)) col.balance = i;
      else if (/debit|withdraw|money out|paid out/.test(h) && !/credit/.test(h)) col.debit = i;
      else if (/credit|deposit|money in|paid in/.test(h)) col.credit = i;
      else if (/amount/.test(h) && col.amount < 0) col.amount = i;
    });
    var prefs = [/transaction details|narrative|description|details|particulars|memo/, /merchant|payee|reference/];
    prefs.some(function (re) { var i = head.findIndex(function (h) { return re.test(h); }); if (i >= 0) { col.desc = i; return true; } return false; });
    if (col.date < 0) col.date = head.findIndex(function (h) { return /date/.test(h); });
  } else {
    // Headerless (CBA: date, amount, description, balance; ANZ: date, amount, description)
    var sample = data.slice(0, 20), ncol = data[0].length;
    var isNum = [], isDate = [], textLen = [];
    for (var c = 0; c < ncol; c++) {
      isDate[c] = sample.every(function (r) { return !r[c] || parseDate_(r[c]); });
      isNum[c] = !isDate[c] && sample.every(function (r) { return !String(r[c] || '').trim() || /^[\s$+\-(]*[\d,]+(\.\d+)?\)?\s*(CR|DR)?$/i.test(String(r[c]).trim()); });
      textLen[c] = sample.reduce(function (s, r) { return s + String(r[c] || '').length; }, 0);
    }
    col.date = isDate.indexOf(true);
    var nums = []; for (c = 0; c < ncol; c++) if (isNum[c]) nums.push(c);
    col.amount = nums[0]; if (nums.length > 1) col.balance = nums[1];
    var best = -1; for (c = 0; c < ncol; c++) if (!isNum[c] && !isDate[c] && (best < 0 || textLen[c] > textLen[best])) best = c;
    col.desc = best;
  }
  if (col.date < 0 || (col.amount < 0 && col.debit < 0 && col.credit < 0))
    throw new Error('Could not work out the date and amount columns. First row was: ' + rows[0].join(', '));
  var out = [];
  data.forEach(function (r) {
    var d = parseDate_(r[col.date]);
    if (!d) return;
    var amt;
    if (col.amount >= 0 && String(r[col.amount] || '').trim()) amt = parseAmount_(r[col.amount]);
    else {
      var cr = col.credit >= 0 ? parseAmount_(r[col.credit]) : null, db = col.debit >= 0 ? parseAmount_(r[col.debit]) : null;
      amt = (cr ? Math.abs(cr) : 0) - (db ? Math.abs(db) : 0);
    }
    if (amt === null || isNaN(amt)) return;
    out.push({ date: d, desc: col.desc >= 0 ? String(r[col.desc] || '').trim().replace(/\s+/g, ' ') : '',
               amount: Math.round(amt * 100) / 100, balance: col.balance >= 0 ? parseAmount_(r[col.balance]) : null });
  });
  return { rows: out, columns: col, header: hasHeader };
}

/**
 * Pure: puts rows in date order and, within each day, oldest-first by following the balance chain
 * (a row's opening balance = balance - amount = the previous row's balance). Rows without balances, or a chain
 * that can't be resolved, fall back to reversing their incoming order (bank exports list newest first) -
 * or, with fallback 'keep', to leaving it as it is (rows already stored oldest-first).
 * Works on any row shape via the three accessors.
 */
function chainOrder_(rows, getDate, getAmt, getBal, fallback) {
  var byDay = {}, keys = [];
  rows.forEach(function (r, i) {
    var k = iso_(toDate_(getDate(r)));
    if (!byDay[k]) { byDay[k] = []; keys.push(k); }
    byDay[k].push({ r: r, i: i });
  });
  keys.sort();
  var out = [];
  keys.forEach(function (k) {
    var list = byDay[k];
    if (list.length === 1) { out.push(list[0].r); return; }
    var ok = list.every(function (x) { var b = getBal(x.r); return b !== null && b !== undefined && b !== '' && !isNaN(b); });
    var ordered = null;
    if (ok) {
      var open = function (x) { return (Number(getBal(x.r)) - Number(getAmt(x.r))).toFixed(2); };
      var bal = function (x) { return Number(getBal(x.r)).toFixed(2); };
      var starts = list.filter(function (x) { return !list.some(function (y) { return y !== x && bal(y) === open(x); }); });
      if (starts.length === 1) {
        ordered = [starts[0]];
        var used = {}; used[starts[0].i] = true;
        while (ordered.length < list.length) {
          var cur = ordered[ordered.length - 1];
          var next = list.filter(function (y) { return !used[y.i] && open(y) === bal(cur); });
          if (next.length !== 1) { ordered = null; break; }
          used[next[0].i] = true; ordered.push(next[0]);
        }
      }
    }
    if (!ordered) ordered = fallback === 'keep' ? list : list.slice().reverse();
    ordered.forEach(function (x) { out.push(x.r); });
  });
  return out;
}

/**
 * One-off repair (Dashboard menu): rewrites each account's Transactions tab in date order, oldest-first within
 * each day. Same rows, same IDs - Reconcile refers to Txn IDs, not row numbers, so it stays valid.
 */
function repairTransactionOrder() {
  var keep = ACCOUNT_.id, msgs = [];
  try {
    Object.keys(ACCOUNTS_).forEach(function (k) {
      if (!ss_().getSheetByName(ACCOUNTS_[k].txns)) return;
      useAccount_(k);
      var t = readTable_(SHEETS.txns);
      if (!t.rows.length) return;
      var sorted = chainOrder_(t.rows, function (r) { return r.Date; }, function (r) { return r.Amount; },
                               function (r) { return r.Balance; }, 'keep');
      var moved = sorted.filter(function (o, i) { return o !== t.rows[i]; }).length;
      var data = sorted.map(function (o) { return t.headers.map(function (h) { return o[h] === undefined ? '' : o[h]; }); });
      var last = t.sheet.getLastRow();
      for (var pad = data.length + 2; pad <= last; pad++) data.push(t.headers.map(function () { return ''; }));  // blank rows go to the end
      t.sheet.getRange(2, 1, data.length, t.headers.length).setValues(data);
      forgetTables_(SHEETS.txns);
      msgs.push(ACCOUNTS_[k].label + ': ' + t.rows.length + ' transactions, ' + moved + ' moved');
    });
  } finally { useAccount_(keep); }
  try { SpreadsheetApp.getUi().alert('Transaction order repaired\n\n' + msgs.join('\n')); } catch (e) {}
  return msgs;
}

function txnKey_(t) { return iso_(t.date) + '|' + t.amount.toFixed(2) + '|' + t.desc.toUpperCase() + '|' + (t.balance === null ? '' : t.balance.toFixed(2)); }
function hash_(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; return ('0000000' + h.toString(16)).slice(-8); }
function blankCell_(v) { return v === '' || v === null || v === undefined; }

/** Called from the dashboard with the CSV text. Skips anything already imported and anything before the Opening Date. */
function importCsv(text) {
  ensureSchema_();
  var settings = readSettings_();
  var parsed = parseBankCsv_(text);
  var existing = readTable_(SHEETS.txns).rows;
  var seen = {}, ids = {};
  existing.forEach(function (r) {
    // Same normalising as parseBankCsv_ (trimmed, single spaces), so rows typed or pasted into the sheet by hand
    // with the bank's double spaces are still recognised as already imported.
    var k = txnKey_({ date: toDate_(r.Date), desc: String(r.Description || '').trim().replace(/\s+/g, ' '), amount: Number(r.Amount),
                      balance: blankCell_(r.Balance) ? null : Number(r.Balance) });
    seen[k] = (seen[k] || 0) + 1; ids[String(r['Txn ID'])] = true;
  });
  var inFile = {}, add = [], now = new Date(), early = 0;
  // Everything is kept. Rows before the Opening Date are history: shown as past actuals and against the rules,
  // and drawn on the graph, but never reconciled.
  parsed.rows.sort(function (a, b) { return a.date - b.date; }).forEach(function (t) {
    var k = txnKey_(t);
    inFile[k] = (inFile[k] || 0) + 1;
    if (inFile[k] <= (seen[k] || 0)) return;
    var id = 'T' + hash_(k + '#' + inFile[k]);
    while (ids[id]) id = 'T' + hash_(id + 'x');
    ids[id] = true;
    if (t.date < settings.openingDate) early++;
    add.push({ 'Txn ID': id, Date: t.date, Description: t.desc, Amount: t.amount, Balance: t.balance === null ? '' : t.balance, Imported: now });
  });
  // IDs are already fixed above (file order); this only changes the order rows are written in: oldest-first within
  // a day, so the last row of a day is its closing balance.
  add = chainOrder_(add, function (r) { return r.Date; }, function (r) { return r.Amount; }, function (r) { return r.Balance; });
  appendRows_(SHEETS.txns, add);
  var matched = add.length ? autoMatch() : 0;
  return { added: add.length, duplicates: parsed.rows.length - add.length, beforeStart: early, matched: matched,
           data: getDashboardData() };
}

/* ------------------------------------------------------------------ bills + Drive */

function billsFolder_(type) {
  var settings = readSettings_();
  var folder = null;
  if (settings.folderId) { try { folder = DriveApp.getFolderById(settings.folderId); } catch (e) { folder = null; } }
  if (!folder) {
    folder = DriveApp.createFolder('Bills');
    var row = settings._rows['Bills Folder ID'];
    if (row) { ss_().getSheetByName(SHEETS.settings).getRange(row.row, 2).setValue(folder.getId()); forgetTables_(SHEETS.settings); }
  }
  if (!type) return folder;
  var sub = folder.getFoldersByName(type);
  return sub.hasNext() ? sub.next() : folder.createFolder(type);
}

function saveFile_(file, type, name) {
  var bytes = Utilities.base64Decode(file.base64);
  var ext = (file.name.match(/\.[A-Za-z0-9]+$/) || ['.pdf'])[0];
  var blob = Utilities.newBlob(bytes, file.mimeType || 'application/pdf', name.replace(/[\/\\:*?"<>|]/g, '-') + ext);
  return billsFolder_(type).createFile(blob).getUrl();
}

/** form: {type, biller, invoice, issued, due, paid, from, to, qty, unit, subtotal, discount, total, address, notes, file?} */
function addBill(form) {
  if (!form.type || !form.total || !form.due) throw new Error('Bill type, total and due date are required.');
  var from = form.from ? fromIso_(form.from) : '', to = form.to ? fromIso_(form.to) : '';
  var url = '';
  if (form.file && form.file.base64) url = saveFile_(form.file, form.type, (form.issued || form.due) + ' ' + form.biller + ' ' + form.type + (form.invoice ? ' ' + form.invoice : ''));
  appendRows_(SHEETS.bills, [{
    'Bill Type': form.type, Biller: form.biller, 'Invoice No.': form.invoice ? "'" + form.invoice : '',
    'Date Issued': form.issued ? fromIso_(form.issued) : '', 'Date Due': fromIso_(form.due), 'Date Paid': form.paid ? fromIso_(form.paid) : '',
    'Period From': from, 'Period To': to, Days: from && to ? daysBetween_(from, to) + 1 : '',
    Quantity: form.qty === '' ? '' : Number(form.qty), Unit: form.unit, Subtotal: form.subtotal === '' ? '' : Number(form.subtotal),
    Discount: Number(form.discount) || 0, Total: Number(form.total), 'Supply Address': form.address, 'Bill PDF': url, Notes: form.notes || ''
  }]);
  return getDashboardData();
}

function attachPdf(row, file) {
  var b = readTable_(SHEETS.bills).rows.filter(function (r) { return r._row === row; })[0];
  if (!b) throw new Error('Bill not found - refresh and try again.');
  var nb = normBill_(b);
  var url = saveFile_(file, nb.type, iso_(nb.issued || nb.due) + ' ' + nb.biller + ' ' + nb.type + (nb.invoice ? ' ' + nb.invoice : ''));
  setCell_(SHEETS.bills, row, 'Bill PDF', url);
  return getDashboardData();
}
