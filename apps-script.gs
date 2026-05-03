// Paste this entire file into Google Apps Script (Extensions > Apps Script)
// Then: Deploy > New deployment > Web app
//   Execute as: Me
//   Who has access: Anyone
// Copy the deployment URL and paste it into SHEETS_URL in index.html

const SHEET_NAME = 'habits';
const HEADERS = [
  'date','habit1_done','habit2_done','habit3_done','perfect_day',
  'deep_work','recharge_j','recharge_f','recharge_e','recharge_m','recharge_count',
  'mit_text','reflection','tags','friend_count','energy_rating','synced_at'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    // Build the row values in header order
    const row = HEADERS.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');

    // Upsert: find existing row for this date and update it, otherwise append
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      const rowIdx = dates.indexOf(data.date);
      if (rowIdx >= 0) {
        sheet.getRange(rowIdx + 2, 1, 1, HEADERS.length).setValues([row]);
        return jsonResponse({ status: 'ok', action: 'updated' });
      }
    }

    sheet.appendRow(row);
    return jsonResponse({ status: 'ok', action: 'inserted' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// Health check — open the deployment URL in your browser to verify it works
function doGet() {
  return jsonResponse({ status: 'alive', sheet: SHEET_NAME });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
