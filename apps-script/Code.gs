/**
 * Spirelight intake — Apps Script backend.
 *
 * Handles TWO things from one Web app deployment:
 *  1. The existing "referrer refers someone" form (referral-form/index.html)
 *     — unchanged behavior, still writes to the "Referidos" tab.
 *  2. NEW: general lead intake (website form and/or Facebook Instant Form
 *     via Make.com) — writes to "Leads Sitio", checks country against the
 *     "Config" tab, dedupes by email, and (via a time trigger) emails
 *     anyone whose country later goes active.
 *
 * Setup:
 * 1. Open the "Spirelight Referral Tracker" Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this file's contents in (replacing what's there), save.
 * 4. Run `setupSheetsOnce` once from the editor (▶ button, pick that
 *    function first) — creates the "Config" and "Leads Sitio" tabs.
 *    Authorize when prompted (it's your own script — click through
 *    Google's "unsafe" warning, that's standard for any unpublished script).
 * 5. Run `installHourlyTrigger` once — schedules the waiting-list email
 *    sweep to run automatically every hour. (Safe to run twice; delete
 *    duplicate triggers from the editor's clock icon on the left if you
 *    ever do.)
 * 6. Deploy > Manage deployments > edit (pencil) the existing deployment
 *    > Version: New version > Deploy. The Web app URL stays the SAME as
 *    what's already in referral-form/index.html — no need to change it
 *    there. Use that same URL for the new confirmation page's JSONP call
 *    and for any new lead-intake form.
 * 7. Check the "Config" tab after setup — it's seeded with the current
 *    active/waiting status per country. Flip a country's second column
 *    to "activo" whenever it goes live; that's the only manual step from
 *    here on. Everyone already waiting for it gets emailed within the
 *    hour, automatically.
 */

var SIGNUP_LINK = 'https://voice.spirelight.ai/login?ref=QU2R4Y55';
var WHATSAPP_GROUP_LINK = 'https://chat.whatsapp.com/LnMEOkmKOc3COzB5Y0vgqG';
var CONFIG_SHEET_NAME = 'Config';
var LEADS_SHEET_NAME = 'Leads Sitio';

// ---------------------------------------------------------------------
// One-time setup
// ---------------------------------------------------------------------

function setupSheetsOnce() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.appendRow(['País', 'Estado']);
    var countries = [
      ['Panamá', 'activo'],
      ['Puerto Rico', 'activo'],
      ['Argentina', 'activo'],
      ['República Dominicana', 'activo'],
      ['Cuba', 'esperando'],
      ['Honduras', 'esperando'],
      ['Paraguay', 'esperando'],
      ['Nicaragua', 'esperando'],
      ['El Salvador', 'esperando'],
      ['Costa Rica', 'esperando'],
      ['Uruguay', 'esperando'],
      ['Ecuador', 'esperando'],
      ['Perú', 'esperando'],
      ['Venezuela', 'esperando'],
      ['Chile', 'esperando'],
      ['Guatemala', 'esperando']
    ];
    countries.forEach(function (row) { configSheet.appendRow(row); });
  }

  var leadsSheet = ss.getSheetByName(LEADS_SHEET_NAME);
  if (!leadsSheet) {
    leadsSheet = ss.insertSheet(LEADS_SHEET_NAME);
    leadsSheet.appendRow(['Fecha', 'Fuente', 'Nombre', 'Email', 'WhatsApp', 'País', 'Estado', 'CorreoEnviado']);
  }
}

function installHourlyTrigger() {
  ScriptApp.newTrigger('sendWaitingListEmails')
    .timeBased()
    .everyHours(1)
    .create();
}

// ---------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------

function getActiveCountries_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var active = {};
  if (!configSheet) return active;
  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var country = String(data[i][0] || '').trim();
    var status = String(data[i][1] || '').trim().toLowerCase();
    if (country && status === 'activo') active[country] = true;
  }
  return active;
}

function isCountryActive_(country) {
  return !!getActiveCountries_()[String(country || '').trim()];
}

// ---------------------------------------------------------------------
// doGet — JSONP country-check for the /gracias/ confirmation page.
// Usage: <script src=".../exec?country=Panamá&callback=xyz"></script>
// ---------------------------------------------------------------------

function doGet(e) {
  var country = (e.parameter && e.parameter.country) || '';
  var callback = (e.parameter && e.parameter.callback) || 'callback';
  var active = isCountryActive_(country);
  var payload = { active: active, link: active ? SIGNUP_LINK : null, whatsapp: WHATSAPP_GROUP_LINK };
  var body = callback + '(' + JSON.stringify(payload) + ');';
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ---------------------------------------------------------------------
// doPost — branches on formType. Defaults to 'referral' so the existing
// referral-form/index.html (which never sends formType) keeps working
// exactly as before, untouched.
// ---------------------------------------------------------------------

function doPost(e) {
  var p = e.parameter;
  var formType = p.formType || 'referral';

  if (formType === 'lead') {
    return handleLeadSubmission_(p);
  }

  return handleReferralSubmission_(p);
}

function handleReferralSubmission_(p) {
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

function handleLeadSubmission_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName(LEADS_SHEET_NAME) || ss.insertSheet(LEADS_SHEET_NAME);

  if (leadsSheet.getLastRow() === 0) {
    leadsSheet.appendRow(['Fecha', 'Fuente', 'Nombre', 'Email', 'WhatsApp', 'País', 'Estado', 'CorreoEnviado']);
  }

  var name = p.name || '';
  var email = String(p.email || '').trim().toLowerCase();
  var whatsapp = p.whatsapp || '';
  var country = p.country || '';
  var source = p.source || 'Website';

  // Dedup by email.
  var data = leadsSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (email && String(data[i][3] || '').trim().toLowerCase() === email) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'duplicate' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  var status = isCountryActive_(country) ? 'activo' : 'esperando';
  leadsSheet.appendRow([new Date(), source, name, email, whatsapp, country, status, false]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// Hourly sweep — emails anyone in "Leads Sitio" whose country has gone
// active since they submitted, and marks them as sent.
// ---------------------------------------------------------------------

function sendWaitingListEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName(LEADS_SHEET_NAME);
  if (!leadsSheet) return;

  var data = leadsSheet.getDataRange().getValues();
  var activeCountries = getActiveCountries_();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = row[2];
    var email = row[3];
    var country = String(row[5] || '').trim();
    var status = row[6];
    var sent = row[7];

    if (status === 'esperando' && sent !== true && activeCountries[country]) {
      sendActivationEmail_(name, email);
      leadsSheet.getRange(i + 1, 8).setValue(true);   // CorreoEnviado
      leadsSheet.getRange(i + 1, 7).setValue('activo'); // Estado
    }
  }
}

function sendActivationEmail_(name, email) {
  if (!email) return;
  var greeting = name ? ('¡Hola ' + name + '!') : '¡Hola!';
  var subject = '🎙️ ¡Tu país ya está activo en Spirelight!';
  var body = greeting + '\n\n'
    + 'Buenas noticias: tu país ya está activo en el programa de grabación de voz de Spirelight.\n\n'
    + 'Aplica aquí para comenzar: ' + SIGNUP_LINK + '\n\n'
    + 'Si aún no lo has hecho, únete a nuestro grupo de WhatsApp para el video explicativo y las preguntas frecuentes: '
    + WHATSAPP_GROUP_LINK + '\n\n'
    + '¡Nos vemos ahí!';
  GmailApp.sendEmail(email, subject, body);
}
