/**
 * Seriyet waiting list -> Google Sheet
 *
 * 1. Create a Google Sheet, then Extensions > Apps Script, paste this file.
 * 2. Deploy > New deployment > type "Web app"
 *      Execute as: Me      Who has access: Anyone
 * 3. Copy the web-app URL (ends with /exec) into ENDPOINT in index.html.
 *
 * Each sign-up adds a row: date | WhatsApp number | source (tiktok, insta...) | already there?
 */
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(["Date", "WhatsApp", "Source", "Doublon"]);

  var p = e.parameter || {};
  var phone = String(p.phone || "").replace(/[^\d+]/g, "");
  if (!/^\+216[2459]\d{7}$/.test(phone)) return ContentService.createTextOutput("bad");

  var existing = sheet.getRange(1, 2, sheet.getLastRow(), 1).getValues().flat();
  var dup = existing.indexOf(phone) !== -1;
  sheet.appendRow([new Date(), "'" + phone,  // leading ' keeps the + (stored as text, not a number)
String(p.source || "").slice(0, 200), dup ? "oui" : ""]);
  return ContentService.createTextOutput("ok");
}
