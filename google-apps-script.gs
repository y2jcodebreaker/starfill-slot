/**
 * STARFILL Lucky Spin — Google Sheets logger
 *
 * Receives spin results from the slot machine and appends one row per spin
 * to the active Google Sheet. See SETUP.md for step-by-step instructions.
 */

function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Add header row once, on an empty sheet
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Phone', 'Reel 1', 'Reel 2', 'Reel 3', 'Won', 'Prize'
      ]);
    }

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name      || '',
      "'" + (data.phone || ''),  // leading quote keeps phone as text (no 9.98E+9)
      data.reel1     || '',
      data.reel2     || '',
      data.reel3     || '',
      data.won ? 'WON' : 'miss',
      data.won ? (data.reel1 || '') : ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you open the /exec URL in a browser to confirm it's deployed
function doGet() {
  return ContentService.createTextOutput('STARFILL logger is running.');
}
