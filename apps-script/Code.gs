/**
 * Bills dashboard - server side.
 * Bound to the "Bills Dashboard" Google Sheet (Extensions > Apps Script).
 * The app itself is hosted on GitHub Pages and talks to this script through Api.gs.
 */

var SHEETS = { bills: 'Bills', recurring: 'Recurring', txns: 'Transactions', rec: 'Reconcile', settings: 'Settings' };
var BILL_TYPE_ORDER = ['Electricity', 'Gas', 'Water', 'Sewer'];

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
  function num(k, def) { var v = s[k] && s[k].value; return (v === '' || v === undefined || v === null || isNaN(Number(v))) ? def : Number(v); }
  return {
    openingBalance: num('Opening Balance', 0),
    openingDate: s['Opening Date'] && toDate_(s['Opening Date'].value) || today_(),
    months: num('Forecast Months', 6),
    lowBalance: num('Low Balance Alert', 0),
    windowDays: num('Match Window Days', 5),
    tolerancePct: num('Amount Tolerance %', 10),
    estimateYears: num('Estimate Years', 3),
    amountOnly: String(s['Amount-only Matching'] ? s['Amount-only Matching'].value : 'Yes').toLowerCase().indexOf('n') !== 0,
    amountDays: num('Amount Match Days', 2),
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
  var billsT = readTable_(SHEETS.bills), recT = readTable_(SHEETS.recurring);
  var txT = readTable_(SHEETS.txns), rcT = readTable_(SHEETS.rec);
  return computeModel_(settings, billsT.rows, recT.rows, txT.rows, rcT.rows, today_(), carOtherForForecast_());
}

/** Pure: everything below works on plain row objects so it can be tested off-Google. */
function computeModel_(settings, billRows, recRows, txRows, rcRows, today, carOther) {
  var bills = billRows.map(normBill_).filter(function (b) { return b.type; });
  var txns = txRows.map(function (r) {
    return { id: String(r['Txn ID']), date: toDate_(r.Date), desc: String(r.Description || ''), amount: Number(r.Amount) || 0,
             balance: (r.Balance === '' || r.Balance === null || r.Balance === undefined) ? null : Number(r.Balance), order: r._row };
  }).filter(function (t) { return t.date && t.date >= settings.openingDate; });

  var status = {};      // occurrence key -> {status, txnId}
  var txnUsed = {};     // txn id -> {status, occ}
  rcRows.forEach(function (r) {
    var occ = String(r.Occurrence || ''), tid = String(r['Txn ID'] || ''), st = String(r.Status || '').toLowerCase();
    if (occ) status[occ] = { status: st, txnId: tid };
    if (tid) txnUsed[tid] = { status: st, occ: occ, note: String(r.Note || '') };
  });

  // Where the projection starts: last known bank balance, else opening balance + imported movements.
  var start = { date: settings.openingDate, balance: settings.openingBalance, source: 'opening' };
  if (txns.length) {
    var sorted = txns.slice().sort(function (a, b) { return a.date - b.date || a.order - b.order; });
    var last = sorted[sorted.length - 1];
    var withBal = sorted.filter(function (t) { return t.balance !== null && !isNaN(t.balance); });
    if (withBal.length && withBal[withBal.length - 1].date >= last.date) {
      var lb = withBal[withBal.length - 1];
      start = { date: lb.date, balance: lb.balance, source: 'bank' };
    } else {
      start = { date: last.date, balance: settings.openingBalance + sorted.reduce(function (s, t) { return s + t.amount; }, 0), source: 'opening+imports' };
    }
  }

  var horizon = addMonths_(today > start.date ? today : start.date, settings.months);
  var est = makeEstimator_(bills, settings.estimateYears, today);
  var occs = generateOccurrences_(recRows, settings.openingDate, horizon, bills, est, carOther || []);

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

  return {
    settings: { openingBalance: settings.openingBalance, openingDate: iso_(settings.openingDate), months: settings.months,
                lowBalance: settings.lowBalance, windowDays: settings.windowDays, tolerancePct: settings.tolerancePct,
                amountOnly: settings.amountOnly, amountDays: settings.amountDays },
    bills: bills.map(billOut_),
    addresses: Object.keys(addresses).sort(function (a, b) { return addresses[b] - addresses[a]; }),
    forecast: { start: { date: iso_(start.date), balance: start.balance, source: start.source },
                low: low.date ? low : null, series: series, upcoming: upcoming, horizon: iso_(horizon) },
    attention: { overdue: upcoming.filter(function (u) { return u.state === 'overdue'; }), unmatchedTxns: unmatchedTxns },
    recentTxns: txns.sort(function (a, b) { return b.date - a.date || b.order - a.order; }).slice(0, 40).map(function (t) {
      var o = matchedTxn[t.id], u = txnUsed[t.id];
      return { id: t.id, date: iso_(t.date), desc: t.desc, amount: t.amount, balance: t.balance,
               matched: o ? o.name : (u && u.status === 'ignored' ? '(one-off)' : ''), how: u ? matchMethod_(u) : '' };
    }),
    rules: recRows.filter(function (r) { return String(r.Name || '').trim(); }).map(function (r) {
      var name = String(r.Name).trim(), n = 0;
      occs.forEach(function (o) { if (o.name === name && o.state === 'paid' && o.txnId) n++; });
      return { name: name, matchText: String(r['Match Text'] || ''), learned: splitPhrases_(r['Learned Text']),
               active: String(r.Active || 'Yes').toLowerCase().indexOf('n') !== 0, matched: n };
    }),
    _occs: occs, _txns: txns, _status: status, _txnUsed: txnUsed
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
      for (var j = 1; j < 400; j++) {
        var nd = f.days ? addDays_(latest.date, j * f.days) : addMonths_(latest.date, j * f.months);
        if (nd > to || (end && nd > end)) break;
        if (nd < lo || skip.indexOf(nd.getMonth() + 1) >= 0) continue;
        carOut.push({ key: name + '|' + iso_(nd), name: name, date: nd, sched: nd, amount: -latest.total, source: 'recurring',
                      match: match, billTypes: [], freq: f, note: 'Next after car entry ' + iso_(latest.date) });
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
  if (n.indexOf('amount') >= 0) return 'amount';
  if (n.indexOf('manual') >= 0) return 'manual';
  return 'text';
}

/** Pure: proposes matches for unmatched transactions. Returns [{txnId, key, occ, txn, how}] */
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
      var dd = Math.abs(daysBetween_(o.date, t.date));
      if (dd > settings.windowDays) return;
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
  return out;
}

function autoMatch() {
  var settings = readSettings_();
  var model = buildModel_();
  var props = proposeMatches_(model, settings);
  var now = new Date();
  appendRows_(SHEETS.rec, props.map(function (p) {
    return { Occurrence: p.key, 'Txn ID': p.txnId, Status: 'matched', Note: p.how === 'amount' ? 'auto: amount' : 'auto: text', Updated: now };
  }));
  props.forEach(function (p) { markBillPaid_(p.occ, p.txn.date); });
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
var SCHEMA_VERSION_ = '3';
function ensureSchema_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION_) return;
  var rec = ss_().getSheetByName(SHEETS.recurring);
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
   ['Car Sheet ID', '', 'ID (or full link) of your Vehicle Expenses Google Sheet. Leave blank to hide the Car section.']]
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

function txnKey_(t) { return iso_(t.date) + '|' + t.amount.toFixed(2) + '|' + t.desc.toUpperCase() + '|' + (t.balance === null ? '' : t.balance.toFixed(2)); }
function hash_(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; return ('0000000' + h.toString(16)).slice(-8); }

/** Called from the dashboard with the CSV text. Skips anything already imported and anything before the Opening Date. */
function importCsv(text) {
  var settings = readSettings_();
  var parsed = parseBankCsv_(text);
  var existing = readTable_(SHEETS.txns).rows;
  var seen = {}, ids = {};
  existing.forEach(function (r) {
    var k = txnKey_({ date: toDate_(r.Date), desc: String(r.Description || ''), amount: Number(r.Amount),
                      balance: r.Balance === '' ? null : Number(r.Balance) });
    seen[k] = (seen[k] || 0) + 1; ids[String(r['Txn ID'])] = true;
  });
  var inFile = {}, add = [], now = new Date(), early = 0;
  parsed.rows.sort(function (a, b) { return a.date - b.date; }).forEach(function (t) {
    if (t.date < settings.openingDate) { early++; return; }
    var k = txnKey_(t);
    inFile[k] = (inFile[k] || 0) + 1;
    if (inFile[k] <= (seen[k] || 0)) return;
    var id = 'T' + hash_(k + '#' + inFile[k]);
    while (ids[id]) id = 'T' + hash_(id + 'x');
    ids[id] = true;
    add.push({ 'Txn ID': id, Date: t.date, Description: t.desc, Amount: t.amount, Balance: t.balance === null ? '' : t.balance, Imported: now });
  });
  appendRows_(SHEETS.txns, add);
  var matched = add.length ? autoMatch() : 0;
  return { added: add.length, duplicates: parsed.rows.length - add.length - early, beforeStart: early, matched: matched,
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
