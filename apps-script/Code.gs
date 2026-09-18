/**
 * Spirelight referral intake — Apps Script backend.
 *
 * Setup:
 * 1. Open the "Spirelight Referral Tracker" Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code, paste this file's contents in, save.
 * 4. Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Authorize when prompted (it's your own script, click through the
 *    "unsafe" warning — that's Google's standard prompt for any
 *    unpublished script, not an actual problem).
 * 6. Copy the resulting Web app URL and paste it into SCRIPT_URL at the
 *    top of index.html.
 *
 * Every submission lands as a new row on a "Referidos" tab in this
 * spreadsheet (created automatically on first submission).
 */

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Referidos') || ss.insertSheet('Referidos');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Fecha',
      'Nombre del referente',
      'País del referente',
      'WhatsApp del referente',
      'Nombre del referido',
      'País del referido',
      'WhatsApp del referido',
      'Comentario'
    ]);
  }

  var p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.referrerName || '',
    p.referrerCountry || '',
    p.referrerWhatsapp || '',
    p.referredName || '',
    p.referredCountry || '',
    p.referredWhatsapp || '',
    p.comments || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
