/**
 * Kinara Health: lead store for the landing page and the admin dashboard.
 *
 * SETUP (about 3 minutes)
 * 1. Create a new Google Sheet (any name, e.g. "Kinara Leads").
 * 2. Extensions > Apps Script. Delete what is there and paste this whole file. Save.
 * 3. In the function dropdown pick `setup`, click Run, and allow access.
 *    Open View > Logs (or Execution log): copy the ADMIN KEY it prints.
 * 4. Deploy > New deployment > type "Web app".
 *    Execute as: Me. Who has access: Anyone. Deploy, then copy the Web app URL (ends in /exec).
 * 5. Send the Web app URL to Claude (or paste it into config.js). Keep the ADMIN KEY private:
 *    you type it once into the admin dashboard to unlock the leads.
 *
 * After any change to this file: Deploy > Manage deployments > Edit > Version: New version > Deploy.
 */

var SHEET_NAME = 'Leads';
var COLS = ['id', 'received_at', 'status', 'name', 'phone', 'email', 'age', 'service',
  'preferred_day', 'visit_type', 'reason', 'utm_source', 'utm_medium', 'utm_campaign',
  'utm_content', 'utm_term', 'gclid', 'fbclid', 'landing_page', 'referrer', 'variant',
  'notes', 'updated_at'];
var STATUSES = ['New', 'Contacted', 'Booked', 'Closed'];

function setup() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('ADMIN_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('ADMIN_KEY', key);
  }
  sheet_();
  Logger.log('ADMIN KEY: ' + key);
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Stops spreadsheet formula injection and trims long input. */
function clean_(v, max) {
  var s = (v === undefined || v === null) ? '' : String(v).slice(0, max || 400);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function isAdmin_(key) {
  var k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!k && key === k;
}

function doGet() {
  return json_({ ok: true, service: 'kinara-leads' });
}

function doPost(e) {
  var data = {};
  try { data = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  if (data.action) return admin_(data);

  /* New lead from the landing page */
  if (data.website) return json_({ ok: true }); /* honeypot filled: bot */
  if (!data.name || !data.phone) return json_({ ok: false, error: 'missing_fields' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    var id = 'KH-' + (1000 + sh.getLastRow());
    var now = new Date().toISOString();
    var row = COLS.map(function (c) {
      if (c === 'id') return id;
      if (c === 'received_at') return now;
      if (c === 'status') return 'New';
      if (c === 'notes' || c === 'updated_at') return '';
      return clean_(data[c]);
    });
    sh.appendRow(row);
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

function admin_(data) {
  if (!isAdmin_(data.key)) return json_({ ok: false, error: 'unauthorized' });
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var head = values.shift();

  if (data.action === 'list') {
    var leads = values.map(function (r) {
      var o = {};
      head.forEach(function (h, i) { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
      return o;
    });
    return json_({ ok: true, leads: leads });
  }

  if (data.action === 'update') {
    var ids = [].concat(data.ids || data.id || []);
    var idCol = head.indexOf('id'), stCol = head.indexOf('status'), nCol = head.indexOf('notes'), uCol = head.indexOf('updated_at');
    var count = 0;
    values.forEach(function (r, i) {
      if (ids.indexOf(r[idCol]) === -1) return;
      if (data.status && STATUSES.indexOf(data.status) !== -1) sh.getRange(i + 2, stCol + 1).setValue(data.status);
      if (typeof data.notes === 'string') sh.getRange(i + 2, nCol + 1).setValue(clean_(data.notes, 1000));
      sh.getRange(i + 2, uCol + 1).setValue(new Date().toISOString());
      count++;
    });
    return json_({ ok: true, updated: count });
  }

  return json_({ ok: false, error: 'unknown_action' });
}
