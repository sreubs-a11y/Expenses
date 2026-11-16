/**
 * The app's API. The app on GitHub Pages sends POST requests here:
 *   body (text): {"key": "<your app key>", "fn": "getDashboardData", "args": []}
 *   reply (JSON): {"ok": true, "data": ...}  or  {"ok": false, "error": "...", "code": "..."}
 *
 * Deploy as a Web app with "Execute as: Me" and "Who has access: Anyone". That makes the URL reachable
 * from the app, so the key is what keeps it private: without the right key nothing is read or changed.
 * Make or change the key from the sheet's Dashboard menu.
 */

/** Only these functions can be called through the API. */
var API_FUNCTIONS_ = {
  getDashboardData: getDashboardData, importCsv: importCsv, matchTxn: matchTxn, ignoreTxn: ignoreTxn,
  skipOccurrence: skipOccurrence, undoReconcile: undoReconcile, runMatching: runMatching,
  addLearnedPhrase: addLearnedPhrase, removeLearnedPhrase: removeLearnedPhrase, setMatchText: setMatchText,
  addBill: addBill, attachPdf: attachPdf,
  carGetData: carGetData, carAddFuel: carAddFuel, carAddToll: carAddToll, carAddOther: carAddOther, carDeleteEntry: carDeleteEntry
};

/** Visiting the URL in a browser shows nothing useful - the app is on GitHub Pages. */
function doGet() {
  return ContentService.createTextOutput('This is the data service for your bills app. Open the app instead.');
}

function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return apiReply_({ ok: false, error: 'Bad request', code: 'BAD_REQUEST' }); }

  if (!apiKeyOk_(req.key)) {
    Utilities.sleep(1500);   // slows down anyone guessing
    return apiReply_({ ok: false, error: 'Wrong or missing app key.', code: 'AUTH' });
  }
  var fn = API_FUNCTIONS_[req.fn];
  if (!fn || !Object.prototype.hasOwnProperty.call(API_FUNCTIONS_, req.fn))
    return apiReply_({ ok: false, error: 'Unknown action: ' + req.fn, code: 'BAD_REQUEST' });
  try {
    var data = fn.apply(null, Array.isArray(req.args) ? req.args : []);
    return apiReply_({ ok: true, data: data === undefined ? null : data });
  } catch (err) {
    return apiReply_({ ok: false, error: (err && err.message) || String(err), code: 'ERROR' });
  }
}

function apiReply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Compares in constant time so the reply time gives nothing away. */
function apiKeyOk_(given) {
  var real = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!real || typeof given !== 'string' || given.length !== real.length) return false;
  var diff = 0;
  for (var i = 0; i < real.length; i++) diff |= real.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}

/* ------------------------------------------------------------------ key management (from the sheet) */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Show app key', 'showAppKey')
    .addItem('Make a new app key (signs out every device)', 'newAppKey')
    .addSeparator()
    .addItem('Recalculate car fuel economy', 'carRebuildEconomy')
    .addToUi();
}

function makeKey_() {
  var raw = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
  return raw.replace(/-/g, '').slice(0, 48);
}

function showAppKey() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('APP_KEY');
  if (!key) { key = makeKey_(); props.setProperty('APP_KEY', key); }
  showKeyDialog_(key, 'Your app key');
}

function newAppKey() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Make a new app key?', 'Every phone and computer using the app will need the new key. Use this if a device is lost.', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var key = makeKey_();
  PropertiesService.getScriptProperties().setProperty('APP_KEY', key);
  showKeyDialog_(key, 'New app key');
}

function showKeyDialog_(key, title) {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px system-ui,sans-serif;line-height:1.45">' +
    '<p>Paste these into the app\'s <b>Connect</b> screen on each device. Treat the key like a password.</p>' +
    '<label style="font-size:12px;color:#555">Script URL</label>' +
    '<input readonly value="' + url.replace(/"/g, '&quot;') + '" onclick="this.select()" style="width:100%;padding:6px;margin:2px 0 10px;font:12px monospace">' +
    (url ? '' : '<p style="font-size:12px;color:#a00">No web app deployment found yet - deploy first (Deploy > New deployment > Web app), then use the /exec URL it shows.</p>') +
    '<label style="font-size:12px;color:#555">App key</label>' +
    '<input readonly value="' + key + '" onclick="this.select()" style="width:100%;padding:6px;margin:2px 0;font:12px monospace">' +
    '</div>').setWidth(460).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}
