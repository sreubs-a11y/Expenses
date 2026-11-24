/** API guard rails: key check, allowlist, error shape. Should pass before and after the importer fix. */
const test = require('node:test'), assert = require('node:assert');
const loadGs = require('../load-gs');
const KEY = 'x'.repeat(48);
const post = (ctx, b) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(b) } }).getContent());

test('rejects a missing or wrong key', () => {
  const { ctx } = loadGs({ today: '2026-11-24', appKey: KEY });
  assert.strictEqual(post(ctx, { fn: 'getDashboardData' }).code, 'AUTH');
  assert.strictEqual(post(ctx, { key: 'y'.repeat(48), fn: 'getDashboardData' }).code, 'AUTH');
});
test('only allowlisted functions can be called', () => {
  const { ctx } = loadGs({ today: '2026-11-24', appKey: KEY });
  assert.strictEqual(post(ctx, { key: KEY, fn: 'autoMatch' }).code, 'BAD_REQUEST');
  assert.strictEqual(post(ctx, { key: KEY, fn: 'constructor' }).code, 'BAD_REQUEST');
});
test('getDashboardData works through the API', () => {
  const { ctx } = loadGs({ today: '2026-11-24', appKey: KEY });
  const r = post(ctx, { key: KEY, fn: 'getDashboardData', args: [] });
  assert.ok(r.ok); assert.ok(Array.isArray(r.data.forecast.series));
});
