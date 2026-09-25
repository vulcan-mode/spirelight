/**
 * Spirelight intake — Apps Script backend.
 *
 * Bound to the "13z3Ht..." Sheet (SHEET_ID below), which now holds
 * everything in one place: Config, Leads Sitio, and Referidos. This
 * one Web app deployment handles:
 *  1. The referral form (referral-form/index.html) — writes to the
 *     "Referidos" tab (in THIS sheet, as of 2026-09-25 -- the old
 *     standalone "Spirelight Referral Tracker" sheet is retired/archived,
 *     nothing writes there anymore).
 *  2. General lead intake (website form and/or Facebook Instant Form
 *     via Make.com) — writes to "Leads Sitio", checks country against
 *     the "Config" tab, dedupes by email, and (via a time trigger)
 *     emails anyone whose country later goes active.
 *  3. syncReferralBonuses — hourly, fills in the referral bonus amount
 *     the moment a referral's "Resultado" is manually set to "Aprobado".
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
 * 3. Run `installHourlyTrigger` and `installReferralSyncTrigger` once
 *    each — schedules the waiting-list email sweep and the referral
 *    bonus sync to run automatically every hour.
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

// 2026-09-25: the old standalone "Spirelight Referral Tracker" sheet
// (REFERRAL_SHEET_ID) is retired -- kept only as a historical archive,
// nothing writes to it anymore. Referral tracking now lives as its own
// "Referidos" tab inside the SAME sheet as everything else (SHEET_ID),
// per the user's explicit request to have all data in one place.
var REFERRAL_SHEET_ID_ARCHIVED = '1IN1iv6X-isl2grAIG3f_LXHk1KrgUleqGXWmd3fdAdI'; // historical only, do not write
var SHEET_ID = '13z3HtJpO7TPl67JVUPyrRxqrLhsMqcsSYz3iccqbRM4'; // new leads sheet, after the original ad/form was deleted
var SIGNUP_LINK = 'https://voice.spirelight.ai/login?ref=QU2R4Y55';
var WHATSAPP_GROUP_LINK = 'https://chat.whatsapp.com/LnMEOkmKOc3COzB5Y0vgqG';
var CONFIG_SHEET_NAME = 'Config';
var LEADS_SHEET_NAME = 'Leads Sitio';
var REFERRALS_SHEET_NAME = 'Referidos';

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

  // Deliberately created BLANK, no header row at all. Meta's native
  // Sheets connector writes its own header + columns the moment it's
  // configured, positionally, and does not respect/merge with any
  // pre-existing header row -- pre-seeding one (as this used to do)
  // is exactly what caused the whole column-misalignment mess
  // tonight. "Estado"/"CorreoEnviado" get added later by
  // sendWaitingListEmails via ensureColumn_, once Meta's real headers
  // already exist to append alongside.
  var leadsSheet = ss.getSheetByName(LEADS_SHEET_NAME);
  if (!leadsSheet) {
    ss.insertSheet(LEADS_SHEET_NAME);
  }

  // Unlike Leads Sitio, this tab is entirely our own -- nothing external
  // writes to it, so it's fine (good, even) to seed a real header row
  // plus dropdown data validation up front.
  var referralsSheet = ss.getSheetByName(REFERRALS_SHEET_NAME);
  if (!referralsSheet) {
    referralsSheet = ss.insertSheet(REFERRALS_SHEET_NAME);
    referralsSheet.appendRow([
      'Fecha',
      'NombreReferente',
      'PaisReferente',
      'WhatsAppReferente',
      'NombreReferido',
      'PaisReferido',
      'WhatsAppReferido',
      'Comentario',
      'Resultado',
      'Bono',
      'FechaProgramada',
      'EstadoPago'
    ]);
    referralsSheet.setFrozenRows(1);

    var resultadoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente', 'Aprobado', 'Rechazado'], true)
      .setAllowInvalid(false)
      .build();
    referralsSheet.getRange(2, 9, referralsSheet.getMaxRows() - 1, 1).setDataValidation(resultadoRule);

    var estadoPagoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente', 'Programado', 'Pagado'], true)
      .setAllowInvalid(false)
      .build();
    referralsSheet.getRange(2, 12, referralsSheet.getMaxRows() - 1, 1).setDataValidation(estadoPagoRule);
  }
}

// One-off: wipes out the existing "Leads Sitio" tab (which still has
// the old pre-seeded header + misaligned test rows from tonight) and
// recreates it fully blank, ready for Meta to write its own structure
// into cleanly from scratch.
function resetLeadsSheet() {
  var ss = getSheet_();
  var existing = ss.getSheetByName(LEADS_SHEET_NAME);
  if (existing) {
    ss.deleteSheet(existing);
  }
  ss.insertSheet(LEADS_SHEET_NAME);
}

function installHourlyTrigger() {
  ScriptApp.newTrigger('sendWaitingListEmails')
    .timeBased()
    .everyHours(1)
    .create();
}

function installReferralSyncTrigger() {
  ScriptApp.newTrigger('syncReferralBonuses')
    .timeBased()
    .everyHours(1)
    .create();
}

// ---------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------

// Meta stores multi-word / accented multiple-choice answers as
// lowercased, underscore-slugged, ACCENT-STRIPPED text -- e.g.
// "República Dominicana" comes back as "republica_dominicana", and
// "Panamá" comes back as "panama" (accent dropped entirely, not just
// re-cased). Confirmed live 2026-09-25: this silently broke 3 of 4
// active countries (everything except single-word, unaccented
// Argentina) -- matchCountryCaseInsensitive_'s plain .toLowerCase()
// never caught it. Mirrors the same fix in spirelight-worker/src/index.js.
function normalizeCountryKey_(raw) {
  return String(raw || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/_/g, ' ')
    .trim()
    .toLowerCase();
}

function getActiveCountries_() {
  var configSheet = getSheet_().getSheetByName(CONFIG_SHEET_NAME);
  var active = {};
  if (!configSheet) return active;
  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var country = String(data[i][0] || '').trim();
    var status = String(data[i][1] || '').trim().toLowerCase();
    if (country && status === 'activo') active[normalizeCountryKey_(country)] = true;
  }
  return active;
}

function isCountryActive_(country) {
  return !!getActiveCountries_()[normalizeCountryKey_(country)];
}

// ---------------------------------------------------------------------
// Phone lookup — the real anti-abuse gate. Rather than letting a
// visitor freely self-select any country on /gracias/, they type the
// phone number they already gave Meta's form, we look it up in Leads
// Sitio, and pull back the country (and name) THEY submitted. Someone
// who never actually filled out the real form has no matching row and
// gets nothing -- they can't just claim to be from an active country.
// ---------------------------------------------------------------------

var META_PHONE_HEADER = 'phone_number';

// Meta stores this as e.g. "p:+18352273236"; strip everything but
// digits so formatting differences (spaces, dashes, the "p:" prefix,
// a leading "+") never cause a false negative.
function normalizePhone_(raw) {
  return String(raw || '').replace(/\D/g, '');
}

function findLeadByPhone_(rawPhone) {
  var target = normalizePhone_(rawPhone);
  if (!target) return { found: false };

  var sheet = getSheet_().getSheetByName(LEADS_SHEET_NAME);
  if (!sheet) return { found: false };

  var map = getHeaderIndexMap_(sheet);
  var phoneCol = map[META_PHONE_HEADER];
  var countryCol = map[META_COUNTRY_HEADER];
  var nameCol = map[META_NAME_HEADER];
  if (!phoneCol || !countryCol) return { found: false };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false };

  var numRows = lastRow - 1;
  var phoneValues = sheet.getRange(2, phoneCol, numRows, 1).getValues();
  var countryValues = sheet.getRange(2, countryCol, numRows, 1).getValues();
  var nameValues = nameCol ? sheet.getRange(2, nameCol, numRows, 1).getValues() : null;

  for (var i = 0; i < numRows; i++) {
    // Compare on a suffix match too (last 10 digits) so a stored
    // number with/without a country code still matches what someone
    // types without one.
    var stored = normalizePhone_(phoneValues[i][0]);
    if (stored && (stored === target || stored.slice(-10) === target.slice(-10))) {
      return {
        found: true,
        country: String(countryValues[i][0] || '').trim(),
        name: nameValues ? String(nameValues[i][0] || '').trim() : ''
      };
    }
  }
  return { found: false };
}

// ---------------------------------------------------------------------
// doGet — JSONP check for the /gracias/ confirmation page.
// Country mode:  ...?country=Panamá&callback=xyz
// Phone mode:    ...?phone=+18352273236&callback=xyz
//   (looks up the actual submitted record instead of trusting a
//   freely self-selected country)
// ---------------------------------------------------------------------

function doGet(e) {
  var callback = (e.parameter && e.parameter.callback) || 'callback';
  var phone = (e.parameter && e.parameter.phone) || '';
  var payload;

  if (phone) {
    var lead = findLeadByPhone_(phone);
    if (!lead.found) {
      payload = { found: false };
    } else {
      var active = isCountryActive_(lead.country);
      payload = {
        found: true,
        name: lead.name,
        country: lead.country,
        active: active,
        link: active ? SIGNUP_LINK : null,
        whatsapp: WHATSAPP_GROUP_LINK
      };
    }
  } else {
    var country = (e.parameter && e.parameter.country) || '';
    var isActive = isCountryActive_(country);
    payload = { found: true, active: isActive, link: isActive ? SIGNUP_LINK : null, whatsapp: WHATSAPP_GROUP_LINK };
  }

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
  var ss = getSheet_();
  var sheet = ss.getSheetByName(REFERRALS_SHEET_NAME);
  if (!sheet) {
    setupSheetsOnce(); // creates it with the right headers + dropdowns
    sheet = ss.getSheetByName(REFERRALS_SHEET_NAME);
  }

  sheet.appendRow([
    new Date(),
    p.referrerName || '',
    p.referrerCountry || '',
    p.referrerWhatsapp || '',
    p.referredName || '',
    p.referredCountry || '',
    p.referredWhatsapp || '',
    p.comments || '',
    'Pendiente', // Resultado
    '',          // Bono -- filled in automatically once Resultado -> Aprobado
    '',          // FechaProgramada -- manual for now
    'Pendiente'  // EstadoPago
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2026-09-25: rewritten to be header-name-based, matching every other
// writer of this sheet. The old version wrote its own fixed column
// order (Fecha/Fuente/Nombre/Email/WhatsApp/País/Estado/CorreoEnviado)
// which no longer matches Meta's actual layout -- using it as-is would
// have silently misaligned data exactly like the bug fixed earlier
// tonight. Backs /registro/index.html, the website intake form for
// people who never went through the Facebook ad.
var ID_QUESTION_HEADER = '¿tienes_una_identificación_oficial_válida_de_ese_país_(o_pasaporte_estadounidense_si_eres_de_puerto_rico)?';
var FUENTE_HEADER = 'Fuente';

function handleLeadSubmission_(p) {
  var leadsSheet = getSheet_().getSheetByName(LEADS_SHEET_NAME) || getSheet_().insertSheet(LEADS_SHEET_NAME);

  var name = String(p.name || '').trim();
  var phone = String(p.whatsapp || '').trim();
  var country = String(p.country || '').trim();
  var email = String(p.email || '').trim().toLowerCase();
  var idAnswer = String(p.idAnswer || '').trim();

  if (!phone) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: 'missing phone' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Dedup by phone -- the identifier used everywhere else in this
  // system (not email: plenty of people submitting this way won't
  // give one). Also the actual enforcement of "if they already did
  // the Facebook thing, don't make them do it again" -- the FRONT END
  // already checks this before ever showing the form, this is the
  // server-side backstop.
  var existing = findLeadByPhone_(phone);
  if (existing.found) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'duplicate' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var map = getHeaderIndexMap_(leadsSheet);
  if (Object.keys(map).length === 0) {
    // Sheet has never received a single lead from anywhere yet --
    // bootstrap with the SAME header names Meta itself uses, so a
    // real Meta lead arriving later lines up instead of colliding.
    leadsSheet.appendRow([META_COUNTRY_HEADER, META_EMAIL_HEADER, META_NAME_HEADER, META_PHONE_HEADER]);
    map = getHeaderIndexMap_(leadsSheet);
  }

  var countryCol = ensureColumn_(leadsSheet, map, META_COUNTRY_HEADER);
  var nameCol = ensureColumn_(leadsSheet, map, META_NAME_HEADER);
  var phoneCol = ensureColumn_(leadsSheet, map, META_PHONE_HEADER);
  var emailCol = ensureColumn_(leadsSheet, map, META_EMAIL_HEADER);
  var idQuestionCol = ensureColumn_(leadsSheet, map, ID_QUESTION_HEADER);
  var fuenteCol = ensureColumn_(leadsSheet, map, FUENTE_HEADER);
  var estadoCol = ensureColumn_(leadsSheet, map, STATUS_HEADER);
  var sentCol = ensureColumn_(leadsSheet, map, EMAIL_SENT_HEADER);
  ensureColumn_(leadsSheet, map, UNLOCK_HEADER);

  var rowNum = leadsSheet.getLastRow() + 1;
  leadsSheet.getRange(rowNum, countryCol).setValue(country);
  leadsSheet.getRange(rowNum, nameCol).setValue(name);
  leadsSheet.getRange(rowNum, phoneCol).setValue(phone);
  if (email) leadsSheet.getRange(rowNum, emailCol).setValue(email);
  leadsSheet.getRange(rowNum, idQuestionCol).setValue(idAnswer);
  leadsSheet.getRange(rowNum, fuenteCol).setValue('Sitio web');
  // Set immediately (not just left for the hourly sweep) so a fresh
  // submission redirecting straight to /gracias/ sees the right state
  // right away instead of a stale "esperando" for up to an hour.
  leadsSheet.getRange(rowNum, estadoCol).setValue(isCountryActive_(country) ? 'activo' : 'esperando');
  leadsSheet.getRange(rowNum, sentCol).setValue(false);

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
// Set by the Worker (not by anything here) the moment someone actually
// clicks through to the real Spirelight signup link on /gracias/ --
// this is what lets a returning visit, even on a different device,
// skip the 3-step checklist entirely instead of repeating it. Ensured
// here (not lazily by the Worker) so the column always exists before
// any click needs to write to it.
var UNLOCK_HEADER = 'Desbloqueado';

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

// activeCountriesMap keys are already normalizeCountryKey_'d (see
// getActiveCountries_) -- normalize the raw side the same way.
function matchCountryCaseInsensitive_(raw, activeCountriesMap) {
  return !!activeCountriesMap[normalizeCountryKey_(raw)];
}

// Paused 2026-09-25 at the user's request: rethinking email content and
// the wider site before any more activation emails go out. The hourly
// trigger keeps running and keeps "Estado" up to date, it just skips the
// actual send. Flip back to false when ready to resume.
var EMAILS_PAUSED = true;

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
  ensureColumn_(leadsSheet, map, UNLOCK_HEADER); // just guarantees it exists; the Worker writes to it, this never does

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
    if (isActive && !alreadySent && !EMAILS_PAUSED) {
      var email = String(emailValues[i][0] || '').trim();
      var name = nameValues ? String(nameValues[i][0] || '').trim() : '';
      if (email) {
        sendActivationEmail_(name, email);
        leadsSheet.getRange(rowNum, sentCol).setValue(true);
      }
    }
  }
}

// One-off, 2026-09-25: apology + real signup path for the 28 people
// confirmed to have been incorrectly told their (actually active)
// country wasn't ready yet, due to the country-matching bug fixed
// today (see normalizeCountryKey_). Deliberately NOT part of the
// regular EMAILS_PAUSED-gated flow above -- this is a separate,
// one-time send, run manually once from the editor. Reuses the same
// "CorreoEnviado" flag, so the regular sweep won't double-email these
// people once EMAILS_PAUSED is turned back off.
var BUG_FIX_AFFECTED_PHONES_ = [
  '18094916856', '8295024380', '18292460820', '18498486944', '18494983218', '18292128512',
  '18096984394', '18293744670', '18098765699', '18493983392', '18493903513', '18299349529',
  '18293633479', '18094793495', '18298185625', '18092058182', '18498813991', '18297580403',
  '18099387585', '8098044174', '18092035182', '18293127125', '18492088611', '18092309958',
  '18498599793', '50760961503', '50765111986', '50767379844'
];

function sendBugFixApologyEmails() {
  var leadsSheet = getSheet_().getSheetByName(LEADS_SHEET_NAME);
  if (!leadsSheet) return;

  var map = getHeaderIndexMap_(leadsSheet);
  var phoneCol = map[META_PHONE_HEADER];
  var emailCol = map[META_EMAIL_HEADER];
  var nameCol = map[META_NAME_HEADER];
  var sentCol = ensureColumn_(leadsSheet, map, EMAIL_SENT_HEADER);
  if (!phoneCol || !emailCol) return;

  var lastRow = leadsSheet.getLastRow();
  if (lastRow < 2) return;
  var numRows = lastRow - 1;

  var phoneValues = leadsSheet.getRange(2, phoneCol, numRows, 1).getValues();
  var emailValues = leadsSheet.getRange(2, emailCol, numRows, 1).getValues();
  var nameValues = nameCol ? leadsSheet.getRange(2, nameCol, numRows, 1).getValues() : null;
  var sentValues = leadsSheet.getRange(2, sentCol, numRows, 1).getValues();

  var targetSet = {};
  BUG_FIX_AFFECTED_PHONES_.forEach(function (p) { targetSet[p] = true; });

  var sentCount = 0, skippedNoEmail = 0;
  for (var i = 0; i < numRows; i++) {
    var normalized = normalizePhone_(phoneValues[i][0]);
    if (!targetSet[normalized]) continue;
    if (sentValues[i][0] === true) continue;

    var email = String(emailValues[i][0] || '').trim();
    if (!email) { skippedNoEmail++; continue; }

    var name = nameValues ? String(nameValues[i][0] || '').trim() : '';
    sendBugFixApologyEmail_(name, email);
    leadsSheet.getRange(i + 2, sentCol).setValue(true);
    sentCount++;
  }
  Logger.log('Sent apology emails to ' + sentCount + ' people. ' + skippedNoEmail + ' had no email on file.');
}

function sendBugFixApologyEmail_(name, email) {
  if (!email) return;
  var greeting = name ? ('¡Hola ' + name + '!') : '¡Hola!';
  // No emoji in the subject -- see sendActivationEmail_'s note below on
  // the encoding bug that mangled it there.
  var subject = 'Error de mi parte -- tu país SÍ está activo';
  var body = greeting + '\n\n'
    + 'Quiero ser directo contigo: cometí un error.\n\n'
    + 'Un problema técnico en mi sistema estaba marcando incorrectamente tu país como "todavía no activo", cuando en realidad ya está activo ahora mismo. '
    + 'Esto fue un error mío, y quiero que lo sepas de mi parte -- yo soy el único responsable de este proyecto de referidos y de este sitio, y esta vez me equivoqué.\n\n'
    + 'La buena noticia: ya está corregido, y puedes aplicar a Spirelight hoy mismo.\n\n'
    + 'Esto es dinero sobre la mesa ahora mismo -- no hay razón para esperar ni un día más.\n\n'
    + 'Entra aquí con el mismo número de WhatsApp que usaste antes: https://vulcan-mode.github.io/spirelight/gracias/\n\n'
    + 'Ahí vas a ver el video explicativo, cómo funciona todo, y el enlace directo para aplicar a Spirelight.\n\n'
    + 'De nuevo, lamento mucho el error. Gracias por tu paciencia, y espero verte grabando pronto 🎙️\n\n'
    + '-- Domingo';
  GmailApp.sendEmail(email, subject, body);
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

// ---------------------------------------------------------------------
// Referral bonus sync -- hourly, silent (no email, unlike the sweep
// above). The pass/fail call itself ("Resultado") is always manual --
// Spirelight gives no feed for who actually completed work and got
// paid, so that has to stay a human judgment call in the sheet. This
// only automates the math *after* that call is made: the moment
// Resultado is set to "Aprobado", it fills in the bonus amount so it
// isn't computed by hand every time.
// FechaProgramada (when the referrer can expect payment) is NOT
// touched here yet -- that needs Spirelight's actual biweekly payday
// anchor date, which hasn't been provided. Leave it manual until then.
// ---------------------------------------------------------------------

var REFERRED_COUNTRY_COL = 6;   // F: PaisReferido
var RESULTADO_COL = 9;          // I: Resultado
var BONO_COL = 10;              // J: Bono
var BONO_DEFAULT = 15;
var BONO_PUERTO_RICO = 20;

function syncReferralBonuses() {
  var sheet = getSheet_().getSheetByName(REFERRALS_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var numRows = lastRow - 1;
  var countryValues = sheet.getRange(2, REFERRED_COUNTRY_COL, numRows, 1).getValues();
  var resultadoValues = sheet.getRange(2, RESULTADO_COL, numRows, 1).getValues();
  var bonoValues = sheet.getRange(2, BONO_COL, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var resultado = String(resultadoValues[i][0] || '').trim();
    var bonoAlreadySet = bonoValues[i][0] !== '' && bonoValues[i][0] !== null;

    if (resultado === 'Aprobado' && !bonoAlreadySet) {
      var country = String(countryValues[i][0] || '').trim().toLowerCase();
      var bono = (country === 'puerto rico') ? BONO_PUERTO_RICO : BONO_DEFAULT;
      sheet.getRange(2 + i, BONO_COL).setValue(bono);
    }
  }
}
