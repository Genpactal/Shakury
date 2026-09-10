/**
 * Нұрдаулет & Манура — RSVP backend for Google Sheets.
 * 1) Create a Google Sheet.
 * 2) Extensions -> Apps Script.
 * 3) Paste this file and run setup() once.
 * 4) Deploy -> New deployment -> Web app.
 *    Execute as: Me. Who has access: Anyone.
 * 5) Copy the /exec URL into GOOGLE_SCRIPT_URL in index.html.
 */

const RSVP_SHEET = 'RSVP';
const SUMMARY_SHEET = 'Summary';
const DECLINED = 'Өкінішке орай, келе алмаймын';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script from the Google Sheet, then run setup().');

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  let rsvp = ss.getSheetByName(RSVP_SHEET);
  if (!rsvp) rsvp = ss.insertSheet(RSVP_SHEET);

  if (rsvp.getLastRow() === 0) {
    rsvp.getRange('A1:E1').setValues([[
      'Уақыты', 'Қонақтың аты', 'Жауабы', 'Қонақ саны', 'Пікір'
    ]]);
  }
  rsvp.setFrozenRows(1);
  rsvp.getRange('A1:E1').setFontWeight('bold').setBackground('#f1e7d7');
  rsvp.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm');
  rsvp.setColumnWidth(1, 150);
  rsvp.setColumnWidth(2, 220);
  rsvp.setColumnWidth(3, 250);
  rsvp.setColumnWidth(4, 110);
  rsvp.setColumnWidth(5, 280);

  let summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) summary = ss.insertSheet(SUMMARY_SHEET);
  summary.clear();
  summary.getRange('A1:B1').setValues([['Көрсеткіш', 'Саны']]);
  summary.getRange('A2:A6').setValues([
    ['Барлық жауаптар'],
    ['Келетін жауаптар'],
    ['Келетін қонақтар саны'],
    ['Келе алмайтындар'],
    ['Соңғы жауап уақыты']
  ]);
  summary.getRange('B2').setFormula('=COUNTA(RSVP!A2:A)');
  summary.getRange('B3').setFormula('=COUNTIF(RSVP!C2:C,"<>Өкінішке орай, келе алмаймын")');
  summary.getRange('B4').setFormula('=SUM(RSVP!D2:D)');
  summary.getRange('B5').setFormula('=COUNTIF(RSVP!C2:C,"Өкінішке орай, келе алмаймын")');
  summary.getRange('B6').setFormula('=IFERROR(MAX(RSVP!A2:A),"")').setNumberFormat('dd.mm.yyyy hh:mm');
  summary.getRange('A1:B1').setFontWeight('bold').setBackground('#f1e7d7');
  summary.setFrozenRows(1);
  summary.setColumnWidth(1, 230);
  summary.setColumnWidth(2, 160);

  SpreadsheetApp.flush();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!ssId) throw new Error('Run setup() once before deploying the web app.');

    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(RSVP_SHEET);
    if (!sheet) throw new Error('RSVP sheet not found. Run setup().');

    const p = e && e.parameter ? e.parameter : {};
    const name = String(p.name || '').trim().slice(0, 120);
    const attendance = String(p.attendance || '').trim().slice(0, 120);
    const comment = String(p.comment || '').trim().slice(0, 500);

    if (!name || !attendance) throw new Error('Name and attendance are required.');

    let guestCount = Number.parseInt(p.guestCount, 10);
    if (attendance === DECLINED) guestCount = 0;
    if (!Number.isFinite(guestCount)) guestCount = attendance === DECLINED ? 0 : 1;
    guestCount = Math.max(0, Math.min(20, guestCount));

    sheet.appendRow([new Date(), name, attendance, guestCount, comment]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService
    .createTextOutput('Nurdaulet & Manura RSVP endpoint is working.')
    .setMimeType(ContentService.MimeType.TEXT);
}
