/**
 * Spirelight intake — Apps Script backend.
 *
 * This is the SAME project that already powers the existing referrer-
 * refers-someone form (referral-form/index.html) — you may have renamed
 * the project itself, but it's still the one bound to the "Spirelight
 * Referral Tracker" Sheet. This file now handles TWO things from that
 * one Web app deployment:
 *  1. The existing referral form — unchanged behavior, still writes to
 *     the "Referidos" tab.
 *  2. NEW: general lead intake (website form and/or Facebook Instant
 *     Form via Make.com) — writes to "Leads Sitio", checks country
 *     against the "Config" tab, dedupes by email, and (via a time
 *     trigger) emails anyone whose country later goes active.
 *
 * Setup:
 * 1. In this Apps Script project (the one already bound to the Sheet),
 *    paste this file's contents in (replacing what's there), save.
 * 2. Run `setupSheetsOnce` once from the editor (▶ button — pick that
 *    function from the dropdown next to it first). Authorize when
 *    prompted: Review permissions > pick your account > Advanced > Go to
 *    (project name) (unsafe) > Allow. That's Google's standard warning
 *    for any script you haven't published to the store — it's your own
 *    code running on your own Sheet, not an actual problem.
 * 3. Run `installHourlyTrigger` once — schedules the waiting-list email
 *    sweep to run automatically every hour.
 * 4. Deploy > Manage deployments > edit (pencil icon) the existing
 *    deployment > Version: New version > Deploy. The Web app URL stays
 *    the SAME as what's already in referral-form/index.html — nothing
 *    to change there. gracias/index.html already uses that same URL too.
 * 5. Check the "Config" tab in the Sheet after setup — it's seeded with
 *    the current active/waiting status per country. Flip a country's
 *    second column to "activo" whenever it goes live; that's the only
 *    manual step from here on. Everyone already waiting for it gets
 *    emailed within the hour, automatically.
 */

// Separate sheets on purpose: the referral-form (Referidos tab) keeps
// writing to the ORIGINAL sheet, unaffected by anything below. Only
// the lead-intake side (Config, Leads Sitio, doGet, the email sweep)
// points at the new sheet -- these must never be merged into one
// constant again, or referral submissions would silently split.
var REFERRAL_SHEET_ID = '1IN1iv6X-isl2grAIG3f_LXHk1KrgUleqGXWmd3fdAdI'; // "Spirelight Referral Tracker" -- Referidos tab only
var SHEET_ID = '13z3HtJpO7TPl67JVUPyrRxqrLhsMqcsSYz3iccqbRM4'; // new leads sheet, after the original ad/form was deleted
var SIGNUP_LINK = 'https://voice.spirelight.ai/login?ref=QU2R4Y55';
var WHATSAPP_GROUP_LINK = 'https://chat.whatsapp.com/LnMEOkmKOc3COzB5Y0vgqG';
var CONFIG_SHEET_NAME = 'Config';
var LEADS_SHEET_NAME = 'Leads Sitio';

// Opens the Sheet explicitly by ID rather than relying on
// getActiveSpreadsheet() — works the same whether this project is
// container-bound or not, so it's one less thing to get wrong.
function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// ---------------------------------------------------------------------
// One-time setup
// ---------------------------------------------------------------------

function setupSheetsOnce() {
  var ss = getSheet_();

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
  var configSheet = getSheet_().getSheetByName(CONFIG_SHEET_NAME);
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
  var ss = SpreadsheetApp.openById(REFERRAL_SHEET_ID);
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
  var ss = getSheet_();
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

// Meta's native Google Sheets lead delivery writes its OWN columns
// (id, created_time, ad_id, ..., the question text as a raw header,
// email, full_name, phone_number, lead_status) and does NOT respect
// whatever header row already existed on the sheet -- it just writes
// its own data regardless of what the original columns meant. So this
// reads everything by HEADER NAME (never a fixed column index), and
// only adds "Estado"/"CorreoEnviado" as new columns it fully owns.
var META_COUNTRY_HEADER = '¿de_qué_país_eres?';
var META_EMAIL_HEADER = 'email';
var META_NAME_HEADER = 'full_name';
var STATUS_HEADER = 'Estado';
var EMAIL_SENT_HEADER = 'CorreoEnviado';

function getHeaderIndexMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h || '').trim();
    if (key) map[key] = i + 1; // 1-based column index
  });
  return map;
}

function ensureColumn_(sheet, map, headerName) {
  if (map[headerName]) return map[headerName];
  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(headerName);
  map[headerName] = col;
  return col;
}

// Meta stores multiple-choice answers lowercased ("argentina"), while
// the Config tab uses proper capitalization ("Argentina") -- an exact
// string match would silently fail for every single lead.
function matchCountryCaseInsensitive_(raw, activeCountriesMap) {
  var rawLower = String(raw || '').trim().toLowerCase();
  for (var country in activeCountriesMap) {
    if (country.toLowerCase() === rawLower) return true;
  }
  return false;
}

function sendWaitingListEmails() {
  var leadsSheet = getSheet_().getSheetByName(LEADS_SHEET_NAME);
  if (!leadsSheet) return;

  var lastRow = leadsSheet.getLastRow();
  if (lastRow < 2) return;

  var map = getHeaderIndexMap_(leadsSheet);
  var countryCol = map[META_COUNTRY_HEADER];
  var emailCol = map[META_EMAIL_HEADER];
  if (!countryCol || !emailCol) return; // Meta hasn't delivered any leads yet

  var nameCol = map[META_NAME_HEADER] || null;
  var estadoCol = ensureColumn_(leadsSheet, map, STATUS_HEADER);
  var sentCol = ensureColumn_(leadsSheet, map, EMAIL_SENT_HEADER);

  var numRows = lastRow - 1;
  var countryValues = leadsSheet.getRange(2, countryCol, numRows, 1).getValues();
  var emailValues = leadsSheet.getRange(2, emailCol, numRows, 1).getValues();
  var nameValues = nameCol ? leadsSheet.getRange(2, nameCol, numRows, 1).getValues() : null;
  var estadoValues = leadsSheet.getRange(2, estadoCol, numRows, 1).getValues();
  var sentValues = leadsSheet.getRange(2, sentCol, numRows, 1).getValues();

  var activeCountries = getActiveCountries_();

  for (var i = 0; i < numRows; i++) {
    var rowNum = i + 2;
    var isActive = matchCountryCaseInsensitive_(countryValues[i][0], activeCountries);
    var newEstado = isActive ? 'activo' : 'esperando';

    if (String(estadoValues[i][0] || '').trim() !== newEstado) {
      leadsSheet.getRange(rowNum, estadoCol).setValue(newEstado);
    }

    var alreadySent = sentValues[i][0] === true;
    if (isActive && !alreadySent) {
      var email = String(emailValues[i][0] || '').trim();
      var name = nameValues ? String(nameValues[i][0] || '').trim() : '';
      if (email) {
        sendActivationEmail_(name, email);
        leadsSheet.getRange(rowNum, sentCol).setValue(true);
      }
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
