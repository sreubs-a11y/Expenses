/**
 * Loads apps-script/*.gs into a Node vm with in-memory stand-ins for SpreadsheetApp, DriveApp, etc. (mockcore.js),
 * so server code can be tested without Google. Mirrors Apps Script: one shared global scope, and every top-level
 * call starts with fresh per-call caches (TABLES_ is reset before each public function).
 *
 *   const { ctx, g } = loadGs({ today: '2026-11-24', bills, car });
 *   ctx.importCsv(text); g.books.BILLS.data.Transactions  // rows as written
 *
 * bills / car: { TabName: [[header...], [row...], ...] }  Dates as {$d:'yyyy-mm-dd'}.
 */
const fs = require('fs'), vm = require('vm'), path = require('path');
const makeGoogleMock = require('./mockcore');
const GS_DIR = path.join(__dirname, '..', 'apps-script');

module.exports = function loadGs({ today, bills, car, appKey } = {}) {
  const g = makeGoogleMock(bills || require('./fixtures/bills.json'), car || require('./fixtures/car.json'));
  if (appKey) g.props.APP_KEY = appKey;
  g.props.SCHEMA_VERSION = g.props.SCHEMA_VERSION || '5';
  const R = Date;
  class FakeDate extends R { constructor(...a) { if (a.length === 0) super(today + 'T10:00:00'); else super(...a); } static now() { return +new R(today + 'T10:00:00'); } }
  const ctx = Object.assign({ Date: today ? FakeDate : Date, Math, JSON, isNaN, parseFloat, Number, String, Error, Infinity, console, Object, Array, RegExp, isFinite, Buffer }, g);
  vm.createContext(ctx);
  const before = new Set(Object.keys(ctx));
  for (const f of ['Code.gs', 'Car.gs', 'Api.gs']) vm.runInContext(fs.readFileSync(path.join(GS_DIR, f), 'utf8'), ctx, { filename: f });
  for (const k of Object.keys(ctx)) {
    if (!before.has(k) && typeof ctx[k] === 'function' && !k.endsWith('_') && /^[a-z]/.test(k) && k !== 'onOpen') {
      const fn = ctx[k]; ctx[k] = function () { ctx.forgetTables_(); return fn.apply(this, arguments); };
    }
  }
  return { ctx, g };
};
