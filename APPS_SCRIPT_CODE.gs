// Intern Quiz Telemetry — Apps Script Web App
// Deploy as: Web App, Execute as: Me, Who has access: Anyone
// After deploy, copy the Web App URL and send to Claude to wire into quiz HTML.

const SHEET_NAME = 'raw';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // Append one row per event. Schema must match the header row in 'raw' tab.
    sheet.appendRow([
      new Date().toISOString(),
      data.intern_name || '',
      data.day || '',
      data.question_id || '',
      data.section || '',
      data.type || '',
      data.answer || '',
      data.correct === true ? 'TRUE' : data.correct === false ? 'FALSE' : '',
      data.difficulty_rating || '',
      data.comment || '',
      data.session_id || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: GET endpoint for quick health check in browser
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'Intern quiz telemetry endpoint live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
