/**
 * ARTIME — Apps Script backend (with PIN login)
 * Users tab in this sheet: Name | PIN | Bank (optional)
 * Only people listed there can log in and make entries.
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Login check from the app
    if (data.action === 'verify') {
      var v = checkUser(data.name, data.pin);
      if (v.ok) return jsonResponse({ result: 'success', name: v.name, bank: v.bank });
      return jsonResponse({ result: 'error', code: 'auth', message: v.message });
    }

    // Every entry must come from a valid user
    var auth = checkUser(data.enteredBy, data.pin);
    if (!auth.ok) {
      return jsonResponse({ result: 'error', code: 'auth', message: auth.message });
    }
    var who = auth.name; // name exactly as written in the Users tab

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var dateStr = Utilities.formatDate(now, tz, 'dd-MM-yyyy');
    var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

    if (data.formType === 'cash') {
      var sheet = getOrCreateSheet(ss, 'Cash Book',
        ['Date', 'Time', 'Direction', 'Mode', 'Bank Name', 'Amount', 'Purpose', 'Entered By']);
      sheet.appendRow([
        dateStr, timeStr,
        data.direction, data.mode, data.bankName || '',
        data.amount, data.purpose, who
      ]);
    } else if (data.formType === 'stock') {
      var sheet2 = getOrCreateSheet(ss, 'Stock Register',
        ['Date', 'Time', 'Movement', 'Item / Description', 'Size', 'Design', 'Type', 'Stand',
         'Quantity', 'Rate', 'Party Name', 'Entered By']);
      (data.items || []).forEach(function (it) {
        sheet2.appendRow([
          dateStr, timeStr, data.movement,
          it.description || it.itemName || '',
          it.size || '', it.design || '', it.type || '', it.stand || '',
          it.quantity || '', it.rate || '',
          data.party || '', who
        ]);
      });
    } else {
      return jsonResponse({ result: 'error', message: 'Unknown formType' });
    }
    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.toString() });
  }
}

function checkUser(name, pin) {
  name = String(name || '').trim().toLowerCase();
  pin = String(pin || '').trim();
  if (!name || !pin) return { ok: false, message: 'Enter name and PIN' };

  // lock for 10 minutes after 5 wrong tries
  var cache = CacheService.getScriptCache();
  var key = 'fail_' + name.replace(/\s+/g, '_');
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) return { ok: false, message: 'Too many wrong attempts. Try again in 10 minutes.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Name', 'PIN', 'Bank (optional)']);
    sheet.getRange('B:B').setNumberFormat('@'); // keep PINs as text so 0123 stays 0123
    sheet.setFrozenRows(1);
  }
  var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === name && String(rows[i][1]).trim() === pin) {
      cache.remove(key);
      return { ok: true, name: String(rows[i][0]).trim(), bank: String(rows[i][2] || '').trim() };
    }
  }
  cache.put(key, String(fails + 1), 600);
  return { ok: false, message: 'Wrong name or PIN' };
}

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (action === 'parties') {
    return jsonResponse({ result: 'ok', values: getDistinctColumn(ss, 'Stock Register', 'Party Name') });
  }
  if (action === 'banks') {
    return jsonResponse({ result: 'ok', values: getDistinctColumn(ss, 'Cash Book', 'Bank Name') });
  }
  return jsonResponse({ result: 'ok', message: 'Running.' });
}

function getDistinctColumn(ss, sheetName, headerName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(headerName);
  if (col === -1) return [];
  var values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues().flat();
  var seen = {}, out = [];
  values.forEach(function (v) {
    v = String(v).trim();
    if (v && !seen[v]) { seen[v] = true; out.push(v); }
  });
  return out;
}

function getOrCreateSheet(ss, name, headerRow) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headerRow);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
