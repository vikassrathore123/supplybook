/**
 * COMPANY DATA ENTRY — Apps Script backend
 * -----------------------------------------
 * Receives entries from the HTML form and appends rows into two tabs
 * of THIS spreadsheet:
 *   - "Cash Book"       (money received / paid)
 *   - "Stock Register"  (sale items out / purchase items in — one row per item)
 *
 * Every row is timestamped automatically by the server (date + time of
 * the moment the entry was saved) — nobody has to type a date.
 *
 * SETUP:
 * 1. Open your Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code, paste this whole file in.
 * 4. Deploy -> New deployment -> type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL and paste it into APPS_SCRIPT_URL in index.html.
 * 6. Whenever you edit this code, make a NEW deployment version
 *    (Deploy -> Manage deployments -> Edit -> New version) so the
 *    live URL picks up the change.
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();

    if (data.formType === 'cash') {
      var sheet = getOrCreateSheet(ss, 'Cash Book',
        ['Date', 'Time', 'Direction', 'Mode', 'Bank Name', 'Amount', 'Purpose', 'Entered By']);
      sheet.appendRow([
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy'),
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
        data.direction,     // Received / Paid
        data.mode,          // Cash / Bank
        data.bankName || '',
        data.amount,
        data.purpose,
        data.enteredBy
      ]);
    } else if (data.formType === 'stock') {
      var sheet2 = getOrCreateSheet(ss, 'Stock Register',
        ['Date', 'Time', 'Movement', 'Item / Description', 'Size', 'Design', 'Type', 'Stand',
         'Quantity', 'Rate', 'Party Name', 'Entered By']);
      var items = data.items || [];
      var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MM-yyyy');
      var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      items.forEach(function (it) {
        sheet2.appendRow([
          dateStr,
          timeStr,
          data.movement,               // Sale / Purchase
          it.description || it.itemName || '',
          it.size || '',
          it.design || '',
          it.type || '',
          it.stand || '',
          it.quantity || '',
          it.rate || '',
          data.party || '',
          data.enteredBy
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

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'parties') {
    return jsonResponse({ result: 'ok', values: getDistinctColumn(ss, 'Stock Register', 'Party Name') });
  }
  if (action === 'banks') {
    return jsonResponse({ result: 'ok', values: getDistinctColumn(ss, 'Cash Book', 'Bank Name') });
  }
  return jsonResponse({ result: 'ok', message: 'Company data entry API is running.' });
}

function getDistinctColumn(ss, sheetName, headerName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(headerName);
  if (col === -1) return [];
  var values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues().flat();
  var seen = {};
  var out = [];
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
