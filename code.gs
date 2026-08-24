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
      reviewer: row[7]
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

  return result;
}

function updateRecord(payload) {
  const sheet = getRecordsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

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

  const fieldMap = {
    rr_title: 'rr_title',
    rr_company: 'rr_company',
    rr_location: 'rr_location',
    rr_linkedin_url: 'rr_linkedin_url'
  };

  Object.keys(fieldMap).forEach(key => {
    if (payload[key] === undefined) return;
    const colName = fieldMap[key];
    const colIndex = headers.indexOf(colName);
    if (colIndex !== -1) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(payload[key]);
    }
  });

  return { status: 'ok', updated_row: rowIndex };
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
    payload.reviewer || ''
  ];

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testLookup() {
  const result = rocketReachLookup('Mark Fidler', 'Berkshire Grey');
  Logger.log(JSON.stringify(result));
}
