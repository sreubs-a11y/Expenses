/**
 * Local preview: serves the app from the repo root and answers its API calls with the real apps-script code
 * running on the synthetic fixtures (in memory, lost on restart). No Google account needed.
 *
 *   cd dev && node serve.js        then open http://localhost:8787
 *
 * The page gets a window.MOCK that POSTs {fn, args, account} to /api, so the Connect screen is skipped.
 */
const http = require('http'), fs = require('fs'), path = require('path');
const loadGs = require('./load-gs');
const ROOT = path.join(__dirname, '..'), PORT = +process.env.PORT || 8787, KEY = 'd'.repeat(48);
const fx = f => fs.readFileSync(path.join(__dirname, 'fixtures', f), 'utf8');
// BILLS_JSON / BILLS_CSV: optional paths to your own (private, never committed) sheet shape and bank export.
const bills = process.env.BILLS_JSON ? JSON.parse(fs.readFileSync(process.env.BILLS_JSON, 'utf8')) : undefined;
const { ctx } = loadGs({ today: process.env.TODAY || '2026-11-24', appKey: KEY, bills });
const call = (fn, args, account) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ key: KEY, fn, args: args || [], account }) } }).getContent());

// Seed: the bills fixture CSV, and a holiday account with a weekly deposit and one planned withdrawal.
call('importCsv', [process.env.BILLS_CSV ? fs.readFileSync(process.env.BILLS_CSV, 'utf8') : fx('bank-2026-11-24.csv')], 'bills');
call('getDashboardData', [], 'holiday');
call('importCsv', [fx('holiday-bank.csv')], 'holiday');
call('addForecastItem', [{ name: 'Weekly deposit', amount: 250, frequency: 'Weekly', date: '2026-11-25', matchText: 'FROM EVERYDAY' }], 'holiday');
call('addForecastItem', [{ name: 'Flights', amount: -1200, frequency: 'Once', date: '2026-12-10' }], 'holiday');

const MOCK = '<script>window.MOCK = function (fn, args, account) { return fetch("/api", { method: "POST", body: JSON.stringify({ fn: fn, args: args, account: account }) })' +
  '.then(function (r) { return r.json(); }).then(function (j) { if (!j.ok) throw new Error(j.error); return j.data; }); };</script>\n';
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.json': 'application/json', '.webmanifest': 'application/manifest+json', '.png': 'image/png', '.css': 'text/css' };

http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api') {
    let body = ''; req.on('data', c => body += c);
    req.on('end', () => { const b = JSON.parse(body || '{}'); res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(call(b.fn, b.args, b.account))); });
    return;
  }
  let p = decodeURIComponent(req.url.split('?')[0]); if (p === '/') p = '/index.html';
  const file = path.join(ROOT, p);
  if (!file.startsWith(ROOT) || /[\\/](dev|apps-script|\.git)([\\/]|$)/.test(path.relative(ROOT, file).replace(/^/, '/')) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.statusCode = 404; return res.end('Not found'); }
  let data = fs.readFileSync(file);
  if (p === '/index.html') data = String(data).replace('<script>\n', MOCK + '<script>\n').replace('<script>\r\n', MOCK + '<script>\r\n');
  res.setHeader('Content-Type', TYPES[path.extname(file)] || 'application/octet-stream');
  res.end(data);
}).listen(PORT, () => console.log('Preview on http://localhost:' + PORT));
