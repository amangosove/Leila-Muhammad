/**
 * Google Sheets RSVP receiver for the Muhammad & Leila invitation.
 * Bind this script to the target Google Sheet (Extensions > Apps Script).
 */
const SHEET_NAME = 'RSVP';
// Optional: paste Telegram bot token and chat ID to receive every RSVP.
const TELEGRAM_BOT_TOKEN = '';
const TELEGRAM_CHAT_ID = '';

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Время получения', 'Имя гостя', 'Ответ', 'Взрослые', 'Дети', 'Всего гостей', 'Комментарий', 'Язык', 'Мероприятие']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    const attendance = ({yes:'Придут', no:'Не придут', maybe:'Пока не знают'})[data.attendance] || String(data.attendance || '');
    const submitted = data.submittedAt ? new Date(data.submittedAt) : new Date();
    sheet.appendRow([submitted, clean_(data.guestName), attendance, Number(data.adults || 0), Number(data.children || 0), Number(data.totalGuests || 0), clean_(data.note), clean_(data.language), clean_(data.event)]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    if (TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) sendTelegram_(data, attendance);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function clean_(value) {
  // Prevent spreadsheet formula injection from guest-provided text.
  const s = String(value == null ? '' : value).slice(0, 1000);
  return /^[=+@\-]/.test(s) ? "'" + s : s;
}

function sendTelegram_(data, attendance) {
  const text = [
    '🔔 Новый ответ на приглашение',
    'Имя: ' + clean_(data.guestName),
    'Ответ: ' + attendance,
    'Взрослые: ' + Number(data.adults || 0),
    'Дети: ' + Number(data.children || 0),
    'Всего гостей: ' + Number(data.totalGuests || 0),
    data.note ? 'Комментарий: ' + clean_(data.note) : '',
    'Язык формы: ' + clean_(data.language)
  ].filter(Boolean).join('\n');
  const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
  UrlFetchApp.fetch(url, {method:'post', contentType:'application/json', payload:JSON.stringify({chat_id:TELEGRAM_CHAT_ID, text:text}), muteHttpExceptions:true});
}

// Run once from the Apps Script editor to confirm the sheet is accessible.
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open Apps Script from inside the target Google Sheet.');
  Logger.log('Connected to: ' + ss.getName());
}
