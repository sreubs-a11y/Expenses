# Bills & Car app

A personal, installable web app (PWA) for a household bills forecast, bank reconciliation, utility usage and car
expenses. Single user. The repo is public: it must never contain account data, keys, sheet IDs or URLs.

## How it fits together

```
Phone / browser                     Google (owner's account)
┌──────────────────────┐  POST     ┌───────────────────────────────┐     ┌────────────────────────┐
│ index.html (PWA)     │ ───────▶  │ Apps Script web app (Api.gs)  │ ──▶ │ "Bills Dashboard" sheet │
│ sw.js  (offline)     │  {key,fn, │ doPost → allowlisted fn       │     │ (script is bound here)  │
│ localStorage: conn,  │   args}   │ Code.gs  bills/forecast/match │ ──▶ │ "Vehicle Expenses" sheet│
│ last data (instant   │ ◀──────── │ Car.gs   car sheet + receipts │     │ (opened by ID)          │
│ open)                │   JSON    └───────────────────────────────┘     └────────────────────────┘
└──────────────────────┘                        │ Drive: receipts folders
```

- **GitHub Pages** serves the repo root: `index.html`, `sw.js`, `manifest.webmanifest`, `icons/`, `vendor/` (pdf.js).
- **`apps-script/`** is the server code. It is NOT run by GitHub; it's pasted into (or pushed with clasp to) the
  Apps Script project bound to the Bills Dashboard sheet. Deployed as a Web app: Execute as *Me*, access *Anyone*.
  The app key (Script Properties `APP_KEY`) is what keeps it private.
- The script URL and app key are entered once per device on the app's Connect screen and kept in localStorage.

## Files

| File | What it is |
|---|---|
| `index.html` | The whole client: CSS, markup, JS. Account sections Bills / Holiday (Overview, Upcoming with filters + past actuals, Utilities, Usage, Reconcile, Rules; Utilities/Usage are bills-only) and a Car section (Dashboard, Add, History) inside one IIFE. Hand-built SVG charts, no chart library. |
| `sw.js` | Service worker. Caches the app shell for instant/offline open; never caches API calls. **Bump `VERSION` on every change to any app file**, otherwise phones keep the old copy. Open apps reload themselves onto a new version. |
| `manifest.webmanifest` | Install metadata. Theme/background `#0d0d0d` so Android's nav bar matches the dark UI. |
| `vendor/pdf.min.mjs`, `vendor/pdf.worker.min.mjs` | pdf.js (Apache-2.0), used to read Linkt toll statements in Car → Add → Toll. |
| `apps-script/Code.gs` | Bills: sheet IO, forecast engine, estimator, matching/reconciliation, CSV import, bills + Drive uploads, schema upgrades. |
| `apps-script/Car.gs` | Car: reads/writes the car sheet (Fuel, Tolls, Other, Allowances), receipts, feeds car items into the forecast. |
| `apps-script/Api.gs` | `doGet` (one-line message), `doPost` JSON API with key check + function allowlist, Dashboard menu (show/rotate key). |
| `apps-script/appsscript.json` | Manifest: V8, `Australia/Melbourne`, web app `USER_DEPLOYING` / `ANYONE_ANONYMOUS`. |
| `dev/` | Node test harness (no Google needed) with synthetic fixtures. `cd dev && npm test`. `node dev/serve.js` = local UI preview on :8787 running the real server code on the fixtures. |

## Server conventions (Apps Script)

- All `.gs` files share **one global scope**. Function and variable names must be unique across files.
  Car code is prefixed `car`/`CAR_`. Private helpers end in `_` (Apps Script won't expose them to the client).
- Code is ES5-style (`var`, `function`) even though the runtime is V8 — keep it that way for consistency.
- Keep logic **pure** where possible (`computeModel_`, `generateOccurrences_`, `proposeMatches_`, `parseBankCsv_`,
  `endOfDayBalances_`, …) so it's testable in `dev/`. Sheet IO goes through `readTable_` / `appendRows_` / `setCell_`.
- `readTable_` memoises per call (`TABLES_`). Any write must go through the helpers or call `forgetTables_(name)`.
- Dates: compare as local-midnight `Date`s via `toDate_`, `iso_`, `addDays_`, `addMonths_`. Never `toISOString()`.
- New columns/settings are added by `ensureSchema_()`; bump `SCHEMA_VERSION_` when you add one so live sheets upgrade.
- Every new client-callable server function must be added to `API_FUNCTIONS_` in `Api.gs`.
- **Accounts.** `ACCOUNTS_` in `Code.gs` (bills, holiday). Each API request carries `account`; `doPost` calls
  `useAccount_()`, which repoints `SHEETS.recurring/txns/rec` at that account's tabs for the call. Settings are shared,
  with per-account rows prefixed (`Holiday Opening Date`, …); an account's own balance/date/alert settings never fall
  back to the bills ones. Only the bills account reads Bills and car items. Car API calls always go as `bills`.
- Date detection uses `Object.prototype.toString.call(v) === '[object Date]'` (not `instanceof`) so the vm harness works.

## Data model (Bills Dashboard sheet)

| Tab | Columns |
|---|---|
| Bills | Bill Type, Biller, Invoice No., Date Issued, Date Due, Date Paid, Period From, Period To, Days, Quantity, Unit, Subtotal, Discount, Total, Supply Address, Bill PDF, Notes |
| Recurring | Name, Amount (− out / + in), Frequency (Weekly…Annual, Once), Next Date, End Date, Skip Months, Match Text, Learned Text, Bill Type, Car Item, Active, Notes |
| Transactions | Txn ID, Date, Description, Amount, Balance, Imported |
| Reconcile | Occurrence (`Name\|yyyy-mm-dd` scheduled date), Txn ID, Status (matched / skipped / ignored), Note, Updated |
| Settings | Setting, Value, Notes (Opening Balance/Date, Forecast Months, History Months, alert, match windows, Car Sheet ID, Cashback Text, Holiday … overrides) |
| Holiday Recurring / Holiday Transactions / Holiday Reconcile | Same columns as Recurring / Transactions / Reconcile, for the holiday account (created by schema v6) |

Car sheet tabs are read **by column position**: Fuel (Date, Litres, Price, Total, Odometer, KM, L/100km, Notes,
Receipts), Tolls (From, To, Total, Comment, Receipts), Other (Date, Item, Total, Comment, Receipts), Allowances
(by row labels). Receipt cells are rich-text links.

**Txn ID** = `"T"` + djb2 hash (h=5381; h=h*33+c, unsigned 32-bit, 8 hex) of
`yyyy-mm-dd|amount(2dp)|DESCRIPTION UPPERCASED, spaces collapsed|balance(2dp)#n`. Re-imports dedupe on the key
(date|amount|desc|balance), so IDs must never change for existing data.

## Bank CSV

Always the bank's export: header `Date,Description,Credit,Debit,Balance`, dd/mm/yyyy, **newest first**, debits
negative, multiple spaces inside descriptions. Import keeps `History Months` before the Opening Date (graph only);
only rows on/after the Opening Date are reconciled. Rows are written oldest-first within a day (`chainOrder_` follows the
balance chain), so the last row of a day is its closing balance; **Dashboard → Repair transaction order** fixes older data.

## Testing

```
cd dev && npm test            # node:test, no dependencies
```
`dev/load-gs.js` loads `apps-script/*.gs` into a vm with in-memory stand-ins for SpreadsheetApp, DriveApp,
PropertiesService, CacheService, ContentService, LockService (`dev/mockcore.js`). Fixtures in `dev/fixtures/` are
synthetic. **Never add real bank, bill or car data to the repo.** For UI checks, serve the repo root
(`python -m http.server`) and point the Connect screen at a local server that calls `doPost` (see `dev/`), or test
against the real deployment.

## Releasing

1. Server: update the Apps Script project (paste, or `clasp push`), then **Deploy → Manage deployments → Edit →
   New version** (same deployment ⇒ same URL and key). Not a new deployment.
2. App: commit to `main`, bump `VERSION` in `sw.js`, wait for Pages to publish, open the app once.
