// Code.gs — deploy this as a Web App (Deploy > New deployment > Web app)
// Execute as: Me
// Who has access: Anyone with the link

const SHEET_NAME = 'Sheet1'; // reviews tab
const RECORDS_SHEET_NAME = 'Records'; // master record list tab

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getRecordsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORDS_SHEET_NAME);
}

function doGet(e) {
  const type = (e.parameter.type || 'reviews');

  if (type === 'records') {
    const sheet = getRecordsSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = row[idx]; });
      records.push(obj);
    }

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const reviews = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    if (!id) continue;
    reviews[id] = {
      action: row[1],
      comment: row[2],
      timestamp: row[3],
      name: row[4],
      conf: row[5],
      score: row[6],
      reviewer: row[7],
      previous_timestamp: row[8] || '',
      previous_reviewer: row[9] || ''
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify(reviews))
    .setMimeType(ContentService.MimeType.JSON);
}

function rocketReachCheckStatus(apiKey, lookupId) {
  const url = 'https://api.rocketreach.co/api/v2/person/checkStatus?ids=' + encodeURIComponent(lookupId);
  const options = {
    method: 'get',
    headers: { 'Api-Key': apiKey },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    return { error: true, status: response.getResponseCode(), message: response.getContentText() };
  }
  const results = JSON.parse(response.getContentText());
  return Array.isArray(results) ? results[0] : results;
}

function rocketReachLookup(name, company) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ROCKETREACH_API_KEY');

  const url = 'https://api.rocketreach.co/api/v2/person/lookup'
    + '?name=' + encodeURIComponent(name)
    + '&current_employer=' + encodeURIComponent(company || '');

  const options = {
    method: 'get',
    headers: { 'Api-Key': apiKey },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status !== 200) {
    return { error: true, status: status, message: body };
  }

  let result = JSON.parse(body);
  const lookupId = result.id;
  let currentStatus = result.status;

  let attempts = 0;
  while (
    (currentStatus === 'searching' || currentStatus === 'progress' || currentStatus === 'waiting')
    && attempts < 4
  ) {
    Utilities.sleep(2000);
    const checked = rocketReachCheckStatus(apiKey, lookupId);
    if (checked && !checked.error) {
      result = checked;
      currentStatus = checked.status;
    }
    attempts++;
  }

  return attachEmailValidation(result);
}

// Reject Postmark / official / automated inboxes so outreach uses a real person email.
var AUTOMATED_EMAIL_LOCAL_PARTS = {
  'noreply': true, 'no-reply': true, 'no_reply': true, 'donotreply': true,
  'do-not-reply': true, 'do_not_reply': true, 'mailer-daemon': true,
  'postmaster': true, 'bounce': true, 'bounces': true, 'notifications': true,
  'notification': true, 'notify': true, 'automated': true, 'auto': true,
  'newsletter': true, 'marketing': true, 'webmaster': true, 'official': true,
  'info': true, 'admin': true, 'support': true, 'sales': true, 'hello': true,
  'contact': true, 'office': true, 'hr': true, 'careers': true, 'press': true,
  'media': true, 'billing': true, 'invoices': true, 'privacy': true,
  'legal': true, 'abuse': true, 'spam': true, 'unsubscribe': true,
  'updates': true, 'alert': true, 'alerts': true, 'system': true, 'mailer': true
};

function isBlockedEmailDomain_(domain) {
  var d = String(domain || '').toLowerCase();
  if (!d) return false;
  if (d === 'postmarkapp.com' || d === 'inbound.postmarkapp.com' || d === 'pm.mtasv.net' || d === 'mtasv.net') {
    return true;
  }
  return d.slice(-15) === '.postmarkapp.com' || d.slice(-9) === '.mtasv.net';
}

function validateOutreachEmail(address, meta) {
  var email = String(address || '').trim();
  var info = meta || {};
  if (!email) {
    return { email: '', status: 'MISSING', reason: 'missing', smtp_valid: '', type: '', grade: '' };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { email: email, status: 'REJECTED', reason: 'invalid_format', smtp_valid: info.smtp_valid || '', type: info.type || '', grade: info.grade || '' };
  }
  var at = email.lastIndexOf('@');
  var local = email.slice(0, at).toLowerCase();
  var domain = email.slice(at + 1).toLowerCase();
  var smtp = String(info.smtp_valid || '').toLowerCase();
  var type = String(info.type || '').toLowerCase();
  var grade = String(info.grade || '');

  if (isBlockedEmailDomain_(domain)) {
    return { email: email, status: 'REJECTED', reason: 'postmark_or_provider', smtp_valid: smtp, type: type, grade: grade };
  }
  if (AUTOMATED_EMAIL_LOCAL_PARTS[local] || local.indexOf('noreply') !== -1 || local.indexOf('no-reply') !== -1) {
    return { email: email, status: 'REJECTED', reason: 'automated_or_official', smtp_valid: smtp, type: type, grade: grade };
  }
  if (type === 'disposable') {
    return { email: email, status: 'REJECTED', reason: 'disposable', smtp_valid: smtp, type: type, grade: grade };
  }
  if (type === 'role-based' || type === 'role_based') {
    return { email: email, status: 'REJECTED', reason: 'role_based', smtp_valid: smtp, type: type, grade: grade };
  }
  if (smtp === 'invalid') {
    return { email: email, status: 'REJECTED', reason: 'smtp_invalid', smtp_valid: smtp, type: type, grade: grade };
  }
  if (smtp === 'accept-all' || smtp === 'accept_all' || smtp === 'unknown') {
    return { email: email, status: 'REVIEW', reason: smtp === 'unknown' ? 'smtp_unknown' : 'smtp_accept_all', smtp_valid: smtp, type: type, grade: grade };
  }
  return { email: email, status: 'USABLE', reason: 'ok', smtp_valid: smtp, type: type, grade: grade };
}

function collectLookupEmails_(result) {
  var seen = {};
  var items = [];
  function add(email, meta) {
    if (!email) return;
    var addr = String(email).trim();
    var key = addr.toLowerCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    items.push({
      email: addr,
      smtp_valid: meta && meta.smtp_valid ? String(meta.smtp_valid) : '',
      type: meta && meta.type ? String(meta.type) : '',
      grade: meta && meta.grade ? String(meta.grade) : ''
    });
  }
  add(result.recommended_email, { type: 'recommended' });
  add(result.recommended_professional_email, { type: 'professional' });
  add(result.recommended_personal_email, { type: 'personal' });
  add(result.current_work_email, { type: 'professional' });
  add(result.email, null);
  var list = result.emails;
  if (Object.prototype.toString.call(list) === '[object Array]') {
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      if (item && typeof item === 'object') {
        add(item.email || item.address, item);
      } else {
        add(item, null);
      }
    }
  }
  return items;
}

function rankEmailStatus_(status) {
  if (status === 'USABLE') return 0;
  if (status === 'REVIEW') return 1;
  if (status === 'REJECTED') return 2;
  return 3;
}

function pickValidatedEmail(result) {
  var candidates = collectLookupEmails_(result).map(function (item) {
    return validateOutreachEmail(item.email, item);
  });
  candidates.sort(function (a, b) {
    var rank = rankEmailStatus_(a.status) - rankEmailStatus_(b.status);
    if (rank !== 0) return rank;
    var aPersonal = a.type === 'personal' ? 0 : 1;
    var bPersonal = b.type === 'personal' ? 0 : 1;
    return aPersonal - bPersonal;
  });
  var best = null;
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].status === 'USABLE' || candidates[i].status === 'REVIEW') {
      best = candidates[i];
      break;
    }
  }
  return {
    best: best,
    candidates: candidates,
    reason: best ? best.reason : (candidates.length ? 'automated_or_official' : 'missing')
  };
}

function attachEmailValidation(result) {
  if (!result || result.error) return result;
  result.email_validation = pickValidatedEmail(result);
  return result;
}

var TRACKED_RECORD_FIELDS = {
  rr_title: 'rr_title',
  rr_company: 'rr_company',
  rr_location: 'rr_location',
  rr_linkedin_url: 'rr_linkedin_url',
  rr_email: 'rr_email',
  rr_email_status: 'rr_email_status',
  review_queue: 'review_queue',
  action: 'action',
  last_updated: 'last_updated',
  last_updated_by: 'last_updated_by',
  updated_at: 'updated_at',
  updated_by: 'updated_by'
};

function ensureColumn(sheet, headers, name) {
  var idx = headers.indexOf(name);
  if (idx !== -1) return idx;
  idx = headers.length;
  sheet.getRange(1, idx + 1).setValue(name);
  headers.push(name);
  return idx;
}

function appendChangeLog(entries) {
  if (!entries || !entries.length) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName('Change Log');
  if (!log) {
    log = ss.insertSheet('Change Log');
    log.appendRow(['timestamp', 'updated_by', 'rr_index', 'name', 'field', 'previous_value', 'new_value']);
  }
  entries.forEach(function(entry) {
    log.appendRow([
      entry.timestamp || '',
      entry.updated_by || '',
      entry.id || '',
      entry.name || '',
      entry.field || '',
      entry.previous == null ? '' : String(entry.previous),
      entry.next == null ? '' : String(entry.next)
    ]);
  });
}

function updateRecord(payload) {
  const sheet = getRecordsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(header) { return String(header || ''); });

  const idColIndex = headers.indexOf('rr_index');
  if (idColIndex === -1) {
    return { error: true, message: 'rr_index column not found in Records sheet' };
  }

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(payload.id)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { error: true, message: 'Record with id ' + payload.id + ' not found' };
  }

  const changes = [];
  Object.keys(TRACKED_RECORD_FIELDS).forEach(function(key) {
    if (payload[key] === undefined) return;
    const colName = TRACKED_RECORD_FIELDS[key];
    const colIndex = ensureColumn(sheet, headers, colName);
    const previous = data[rowIndex - 1][colIndex];
    const next = payload[key];
    if (String(previous == null ? '' : previous) === String(next == null ? '' : next)) return;
    sheet.getRange(rowIndex, colIndex + 1).setValue(next);
    changes.push({
      timestamp: payload.last_updated || new Date().toISOString(),
      updated_by: payload.updated_by || payload.reviewer || '',
      id: payload.id,
      name: payload.name || '',
      field: colName,
      previous: previous,
      next: next
    });
  });
  appendChangeLog(changes);

  return { status: 'ok', updated_row: rowIndex, changes: changes.length };
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);

  if (payload.type === 'rocketreach_lookup') {
    const result = rocketReachLookup(payload.name, payload.company);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (payload.type === 'update_record') {
    const result = updateRecord(payload);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id)) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowValues = [
    payload.id,
    payload.action,
    payload.comment,
    payload.timestamp,
    payload.name,
    payload.conf,
    payload.score,
    payload.reviewer || '',
    payload.previous_timestamp || '',
    payload.previous_reviewer || ''
  ];

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  }

  var queue = payload.action === 'quick' ? 'QUICK_REVIEW'
    : payload.action === 'deep' ? 'DEEP_REVIEW'
    : payload.action === 'ignore' ? 'IGNORE'
    : '';
  if (queue) {
    updateRecord({
      id: payload.id,
      name: payload.name || '',
      action: payload.action,
      review_queue: queue,
      last_updated: payload.previous_timestamp || payload.timestamp || new Date().toISOString(),
      last_updated_by: payload.previous_reviewer || payload.reviewer || '',
      updated_at: payload.timestamp || new Date().toISOString(),
      updated_by: payload.reviewer || '',
      reviewer: payload.reviewer || ''
    });
    if (payload.comment) {
      appendChangeLog([{
        timestamp: payload.timestamp || new Date().toISOString(),
        updated_by: payload.reviewer || '',
        id: payload.id,
        name: payload.name || '',
        field: 'comment',
        previous: '',
        next: payload.comment
      }]);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testLookup() {
  const result = rocketReachLookup('Mark Fidler', 'Berkshire Grey');
  Logger.log(JSON.stringify(result));
}
