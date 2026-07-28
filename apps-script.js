var SHEET_ID = '1ZtuCoxGv3jnOx5_5ID9zVSasXsCjXvNi8lp67QJbClQ';

function doGet(e) {
  var action = e.parameter.action || 'getAll';
  if (action === 'add')    return addRow(e);
  if (action === 'update') return updateRow(e);
  if (action === 'delete') return deleteRow(e);
  return getAllRows();
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action || '';
  if (action === 'batchAdd')   return batchAddRows(body.rows || []);
  if (action === 'clearSheet') return clearSheetData();
  return jsonOk({ error: 'Unknown action' });
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
}

function getAllRows() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return jsonOk([]);
  var startRow = 0;
  if (String(data[0][0]).toLowerCase().includes('farsi') || String(data[0][0]).toLowerCase().includes('general')) {
    startRow = 1;
  }
  var rows = [];
  for (var i = startRow; i < data.length; i++) {
    var row = data[i];
    var gf = String(row[0] || '').trim();
    var tf = String(row[1] || '').trim();
    var ge = String(row[2] || '').trim();
    var te = String(row[3] || '').trim();
    var sc = String(row[4] || '').trim();
    if (!gf && !tf && !ge && !te) continue;
    rows.push({ generalFarsi: gf, technicalFarsi: tf, generalEnglish: ge, technicalEnglish: te, shortcut: sc, _rowIndex: i + 1 });
  }
  return jsonOk(rows);
}

function addRow(e) {
  var sheet = getSheet();
  sheet.appendRow([
    (e.parameter.generalFarsi     || '').trim(),
    (e.parameter.technicalFarsi   || '').trim(),
    (e.parameter.generalEnglish   || '').trim(),
    (e.parameter.technicalEnglish || '').trim(),
    (e.parameter.shortcut         || '').trim()
  ]);
  return jsonOk({ success: true });
}

function batchAddRows(rows) {
  var sheet = getSheet();
  var toAppend = rows.map(function(r) {
    return [
      (r.generalFarsi     || '').trim(),
      (r.technicalFarsi   || '').trim(),
      (r.generalEnglish   || '').trim(),
      (r.technicalEnglish || '').trim(),
      (r.shortcut         || '').trim()
    ];
  });
  if (toAppend.length === 0) return jsonOk({ added: 0 });
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, toAppend.length, 5).setValues(toAppend);
  return jsonOk({ added: toAppend.length });
}

function updateRow(e) {
  var sheet = getSheet();
  var rowIndex = parseInt(e.parameter.rowIndex || '0');
  if (rowIndex < 1) return jsonOk({ success: false, error: 'Missing rowIndex' });
  sheet.getRange(rowIndex, 1, 1, 5).setValues([[
    (e.parameter.generalFarsi     || '').trim(),
    (e.parameter.technicalFarsi   || '').trim(),
    (e.parameter.generalEnglish   || '').trim(),
    (e.parameter.technicalEnglish || '').trim(),
    (e.parameter.shortcut         || '').trim()
  ]]);
  return jsonOk({ success: true });
}

function deleteRow(e) {
  var sheet = getSheet();
  var rowIndex = parseInt(e.parameter.rowIndex || '0');
  if (rowIndex < 1) return jsonOk({ success: false, error: 'Missing rowIndex' });
  sheet.deleteRow(rowIndex);
  return jsonOk({ success: true });
}

function clearSheetData() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  return jsonOk({ cleared: lastRow - 1 });
}

function dedupSheet() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var header = null;
  var startRow = 0;
  if (String(data[0][0]).toLowerCase().includes('farsi') || String(data[0][0]).toLowerCase().includes('general')) {
    header = data[0];
    startRow = 1;
  }
  var seen = {};
  var unique = [];
  for (var i = startRow; i < data.length; i++) {
    var row = data[i];
    var key = String(row[0]||'').trim().slice(0,80) + '|' + String(row[2]||'').trim().slice(0,80) + '|' + String(row[4]||'').trim();
    if (!key.replace(/\|/g,'').trim()) continue;
    if (!seen[key]) { seen[key] = true; unique.push(row); }
  }
  sheet.clearContents();
  var writeStart = 1;
  if (header) { sheet.getRange(1, 1, 1, 5).setValues([header]); writeStart = 2; }
  if (unique.length > 0) sheet.getRange(writeStart, 1, unique.length, 5).setValues(unique);
  return jsonOk({ before: data.length - startRow, after: unique.length });
}

function populateShortcuts() {
  var sheet = getSheet();
  var shortcuts = [
    'واریز-ناموفق',
    'اسکرین‌شات',
    'حساب-دمو',
    'اسلیپیج',
    'تایید',
    'vpn',
    'forfx',
    'سرور-فعال',
    'اختلال'
  ];
  var data = sheet.getDataRange().getValues();
  var startRow = (String(data[0][0]).toLowerCase().includes('farsi') || String(data[0][0]).toLowerCase().includes('general')) ? 1 : 0;
  for (var i = 0; i < shortcuts.length; i++) {
    var sheetRow = startRow + i + 1;
    if (sheetRow <= data.length) {
      sheet.getRange(sheetRow, 5).setValue(shortcuts[i]);
    }
  }
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
