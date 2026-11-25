/** Holiday account, Upcoming past actuals, Rules tab data, importer dedupe and the transaction-order repair. */
const test = require('node:test'), assert = require('node:assert');
const fs = require('fs'), path = require('path');
const loadGs = require('../load-gs');
const fx = f => fs.readFileSync(path.join(__dirname, '..', 'fixtures', f), 'utf8');
const KEY = 'k'.repeat(48);
const post = (ctx, b) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ key: KEY }, b)) } }).getContent());
const D = s => new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

test('hand-typed rows with the bank\'s double spaces are not imported twice', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  const T = g.books.BILLS.data.Transactions;
  T.push(['Tmanual01', D('2026-11-23'), 'PHONECO  PrePaid - Visa  Purchase ', -12, 1761.71, '']);
  const r = ctx.importCsv(fx('bank-2026-11-24.csv'));
  assert.strictEqual(T.slice(1).filter(x => /PHONECO/.test(x[2])).length, 1);
  assert.strictEqual(r.added, 5);
});

test('Repair transaction order puts newest-first days right and keeps every ID', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  const T = g.books.BILLS.data.Transactions;
  const ids = T.slice(1).map(r => r[0]).sort();
  const day = T.slice(1).filter(r => iso(r[1]) === '2026-11-24');
  const others = T.slice(1).filter(r => iso(r[1]) !== '2026-11-24');
  T.splice(1, T.length, ...others, ...day.reverse());       // how the old importer stored them
  ctx.repairTransactionOrder();
  assert.deepStrictEqual(T.slice(1).filter(r => iso(r[1]) === '2026-11-24').map(r => r[4]), [1572.95, 1574.84, 1375.92, 1377.91]);
  assert.deepStrictEqual(T.slice(1).map(r => r[0]).sort(), ids);
});

test('holiday account: own tabs, starts from its bank balance, bills untouched', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24', appKey: KEY });
  const r = post(ctx, { account: 'holiday', fn: 'importCsv', args: [fx('holiday-bank.csv')] });
  assert.ok(r.ok, r.error);
  const B = g.books.BILLS.data;
  assert.strictEqual(B['Holiday Transactions'].length - 1, 7);
  assert.strictEqual(B.Transactions.length, 1, 'bills transactions untouched');
  assert.strictEqual(r.data.data.account, 'holiday');
  assert.strictEqual(r.data.data.hasBills, false);
  assert.deepStrictEqual(r.data.data.bills, []);
  // Opening Date defaults to the upgrade day (today), so everything imported is history; the forecast starts from it.
  assert.deepStrictEqual(r.data.data.forecast.start, { date: '2026-11-18', balance: 2000, source: 'bank' });
  const bills = post(ctx, { fn: 'getDashboardData', args: [] });
  assert.strictEqual(bills.data.account, 'bills');
  assert.strictEqual(post(ctx, { account: 'nope', fn: 'getDashboardData' }).ok, false);
  assert.strictEqual(post(ctx, { account: 'constructor', fn: 'getDashboardData' }).ok, false);
});

test('holiday account: weekly deposit and a planned withdrawal forecast and reconcile', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24', appKey: KEY });
  post(ctx, { account: 'holiday', fn: 'getDashboardData' });
  const S = g.books.BILLS.data.Settings;
  S.find(r => r[0] === 'Holiday Opening Date')[1] = D('2026-11-01');
  post(ctx, { account: 'holiday', fn: 'importCsv', args: [fx('holiday-bank.csv')] });
  let r = post(ctx, { account: 'holiday', fn: 'addForecastItem', args: [{ name: 'Weekly deposit', amount: 250, frequency: 'Weekly', date: '2026-11-25', matchText: 'FROM EVERYDAY' }] });
  assert.ok(r.ok, r.error);
  r = post(ctx, { account: 'holiday', fn: 'addForecastItem', args: [{ name: 'Flights', amount: -1200, frequency: 'Once', date: '2026-12-10' }] });
  const d = r.data, up = d.forecast.upcoming;
  assert.deepStrictEqual(d.forecast.start, { date: '2026-11-18', balance: 2000, source: 'bank' });
  assert.strictEqual(up[0].name, 'Weekly deposit'); assert.strictEqual(up[0].date, '2026-11-25'); assert.strictEqual(up[0].balance, 2250);
  const fl = up.find(u => u.name === 'Flights');
  assert.strictEqual(fl.date, '2026-12-10'); assert.strictEqual(fl.balance, 2000 + 250 * 3 - 1200);
  // the three November deposits matched themselves (anchored weekly from 25 Nov, going back); the October ones
  // are before the Opening Date, so they're paired as history (in memory, not in Reconcile)
  const paid = d.forecast.past.filter(p => p.state === 'paid');
  assert.deepStrictEqual(paid.map(p => p.date), ['2026-11-18', '2026-11-11', '2026-11-04', '2026-10-28', '2026-10-21']);
  assert.deepStrictEqual(paid.map(p => !!p.history), [false, false, false, true, true]);
  assert.strictEqual(g.books.BILLS.data['Holiday Reconcile'].length - 1, 3, 'history pairs are not written');
  const wd = d.forecast.past.find(p => p.amount === -500);
  assert.deepStrictEqual([wd.state, wd.history], ['unlisted', true]);
  assert.ok(paid.every(p => p.expected === 250 && p.amount === 250 && p.name === 'Weekly deposit'));
  assert.strictEqual(paid[0].balance, 2000);
  assert.strictEqual(g.books.BILLS.data.Recurring.length, 5, 'bills Recurring untouched');
  const rule = d.rules.find(x => x.name === 'Weekly deposit');
  assert.strictEqual(rule.kind, 'fixed'); assert.strictEqual(rule.next.date, '2026-11-25');
  assert.strictEqual(rule.recent.length, 5); assert.strictEqual(rule.avgDiff, 0); assert.strictEqual(rule.avgDays, 0);
  assert.strictEqual(d.rules.find(x => x.name === 'Flights').kind, 'once');
});

test('past actuals: matched with expected vs actual, one-offs, unmatched and skipped', () => {
  const { ctx } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  let d = ctx.getDashboardData();
  const filler = d.forecast.past.find(p => /Opening filler/.test(p.name));
  assert.deepStrictEqual([filler.state, filler.history], ['unlisted', true], 'history before the Opening Date is shown, not reconciled');
  assert.ok(!d.attention.unmatchedTxns.some(t => /Opening filler/.test(t.desc)));
  const ph = d.forecast.past.find(p => p.name === 'Phone');
  assert.deepStrictEqual([ph.state, ph.date, ph.amount, ph.expected], ['paid', '2026-11-23', -12, -12]);
  const el = d.forecast.past.find(p => p.name === 'Electricity');
  assert.strictEqual(el.amount, -188.76); assert.strictEqual(el.cashback, 1.89);
  assert.ok(!d.forecast.past.some(p => /Cashback/.test(p.name)), 'cashback folded into its payment');
  const um = d.attention.unmatchedTxns[0];
  if (um) { d = ctx.ignoreTxn(um.id, 'one-off'); assert.ok(d.forecast.past.some(p => p.state === 'one-off' && p.txnId === um.id)); }
  const next = d.forecast.upcoming.find(u => u.name === 'Transfer');
  d = ctx.skipOccurrence(next.key, '');
  assert.ok(d.forecast.past.some(p => p.state === 'skipped' && p.key === next.key));
  assert.strictEqual(d.rules.find(r => r.name === 'Electricity').kind, 'estimate');
  assert.strictEqual(d.rules.find(r => r.name === 'Phone').kind, 'fixed');
});

test('history before the Opening Date shows against its rule (Kieser case) and all of a long CSV is kept', () => {
  const { ctx } = loadGs({ today: '2026-11-24' });
  const rows = ['Date,Description,Credit,Debit,Balance'];
  let bal = 5000;
  const lines = [];
  // fortnightly gym payments from February, a fixed rule anchored on 4 Dec; Opening Date in the fixture is 23 Nov
  for (let d = new Date(2026, 1, 12); d < new Date(2026, 10, 23); d.setDate(d.getDate() + 14)) {
    bal = Math.round((bal - 122.07) * 100) / 100;
    lines.unshift(`${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()},EZI*GYMCO Northcote - Visa Purchase,,-122.07,${bal.toFixed(2)}`);
  }
  const csv = rows.concat(lines).join(String.fromCharCode(10));
  const r = ctx.importCsv(csv);
  assert.strictEqual(r.added, lines.length, 'nothing older than History Months is dropped');
  assert.strictEqual(r.beforeStart, lines.length);
  const g2 = ctx.getDashboardData();
  // add a rule after the fact - history pairs up with it straight away
  const { ctx: c2, g } = loadGs({ today: '2026-11-24' });
  g.books.BILLS.data.Recurring.push(['Gym', -122.07, 'Fortnightly', new Date(2026, 11, 3), '', '', 'GYMCO', '', '', '', 'Yes', '']);
  c2.importCsv(csv);
  const d = c2.getDashboardData();
  const gym = d.rules.find(x => x.name === 'Gym');
  assert.strictEqual(gym.recent.length, 6);
  assert.ok(gym.recent.every(x => x.actual === -122.07 && x.expected === -122.07));
  assert.strictEqual(d.forecast.past.filter(p => p.name === 'Gym').length, lines.length);
  assert.strictEqual(d.attention.unmatchedTxns.length, g2.attention.unmatchedTxns.length);
});
