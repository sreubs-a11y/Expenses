/**
 * Importer same-day ordering (bug found in the 24 Nov 2026 reconciliation).
 * The bank CSV lists newest first. importCsv() sorts by date only, so rows on the same day land in the
 * Transactions tab newest-first: the LAST row of a day is the EARLIEST transaction, not the closing balance.
 * Expected after the fix: same-day rows are written oldest-first (following the balance chain), IDs unchanged.
 */
const test = require('node:test'), assert = require('node:assert');
const fs = require('fs'), path = require('path');
const loadGs = require('../load-gs');
const fx = f => fs.readFileSync(path.join(__dirname, '..', 'fixtures', f), 'utf8');
const iso = d => d.toISOString ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` : d;
const txRows = g => g.books.BILLS.data.Transactions.slice(1).filter(r => r[0]).map(r => ({ id: r[0], date: iso(r[1]), amount: r[3], balance: r[4] }));

test('same-day rows are written oldest-first, so the last row of a day is its closing balance', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  const day = txRows(g).filter(r => r.date === '2026-11-24').map(r => r.balance);
  assert.deepStrictEqual(day, [1572.95, 1574.84, 1375.92, 1377.91]);
});

test('each row follows on from the one before it (balance chain holds within every day)', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  const rows = txRows(g);
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].date !== rows[i - 1].date) continue;
    assert.ok(Math.abs(rows[i - 1].balance + rows[i].amount - rows[i].balance) < 0.005, `row ${i} doesn't follow row ${i - 1}`);
  }
});

test('a later export the same day appends the new row after the earlier ones', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  const r = ctx.importCsv(fx('bank-2026-11-24-later.csv'));
  assert.strictEqual(r.added, 1);
  const day = txRows(g).filter(r => r.date === '2026-11-24').map(r => r.balance);
  assert.deepStrictEqual(day, [1572.95, 1574.84, 1375.92, 1377.91, 1135.18]);
});

test('Txn IDs are unchanged by the fix (re-importing an old export adds nothing)', () => {
  const { ctx, g } = loadGs({ today: '2026-11-24' });
  ctx.importCsv(fx('bank-2026-11-24.csv'));
  const ids = txRows(g).map(r => r.id).sort();
  // IDs the current importer produces for this file (hash of key + '#1'); hand-written sheet rows use the same scheme.
  assert.deepStrictEqual(ids, JSON.parse(fx('expected-ids.json')).sort());
  assert.strictEqual(ctx.importCsv(fx('bank-2026-11-24.csv')).added, 0);
});

test('forecast starts from the bank closing balance on a busy day', () => {
  const { ctx } = loadGs({ today: '2026-11-24' });
  const r = ctx.importCsv(fx('bank-2026-11-24.csv'));
  assert.deepStrictEqual(r.data.forecast.start, { date: '2026-11-24', balance: 1377.91, source: 'bank' });
});

test('endOfDayBalances_ is right even when the chain is ambiguous (falls back to stored order, oldest-first)', () => {
  const { ctx } = loadGs({ today: '2026-11-25' });
  const D = s => new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
  // -50 then +50: both rows "follow on" from each other's opening balance, so the chain can't decide.
  const tx = [{ id: 'a', date: D('2026-11-25'), amount: -50, balance: 1085.18, order: 2 },
              { id: 'b', date: D('2026-11-25'), amount: 50, balance: 1135.18, order: 3 }];
  assert.strictEqual(ctx.endOfDayBalances_(tx).map['2026-11-25'], 1135.18);
});
