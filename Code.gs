/** Google Apps Script: bind this script to a Google Sheet, then deploy as Web app. */
const SHEET_NAME = 'Подтверждения';
function doPost(e) {
  const p = (e && e.parameter) || {};
  const nonce = String(p.nonce || '').slice(0, 120);
  let result = {type:'tek30-rsvp', nonce, ok:false};
  try {
    const name = String(p.guest || '').trim().slice(0, 150);
    const attendance = String(p.attendance || '');
    const people = Number(p.people);
    if (!name || !['Катышам','Катыша албайм'].includes(attendance) || !Number.isInteger(people) || people < 1 || people > 4) throw Error('Invalid form');
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) throw Error('Bind the script to a Google Sheet');
    let sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = book.insertSheet(SHEET_NAME);
      sheet.appendRow(['Дата и время','ФИО гостя','Участие','Количество гостей']);
      sheet.setFrozenRows(1);
    }
    // Prevent user input from being interpreted as spreadsheet formulas.
    const safe = v => /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
    sheet.appendRow([new Date(), safe(name), attendance, attendance === 'Катышам' ? people : 0]);
    result.ok = true;
  } catch (err) {
    result.message = 'Маалыматты сактоо мүмкүн болгон жок.';
  }
  const json = JSON.stringify(result).replace(/</g,'\\u003c');
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>parent.postMessage('+json+',"*")<\/script>');
}
