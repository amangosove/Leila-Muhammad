/**
 * RSVP receiver for the Muhammad & Leila invitation.
 * Bound to the Google Sheet with ID below.
 */
const SPREADSHEET_ID = '1Mo1LZz8VTGCnqC4Aio6eFX8gsPZg3yp6m1VhoDmKkjA';
const SHEET_NAME = 'RSVP';

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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

    // Email notification to the Google account that owns/deploys this script.
    // If mail delivery fails, keep the RSVP row and log the error.
    try {
      const ownerEmail = Session.getEffectiveUser().getEmail();
      if (ownerEmail) {
        const subject = 'Новый ответ на приглашение — Мухаммад и Лейла';
        const body = [
          'Получен новый ответ на приглашение.',
          '',
          'Имя: ' + clean_(data.guestName),
          'Ответ: ' + attendance,
          'Взрослые: ' + Number(data.adults || 0),
          'Дети: ' + Number(data.children || 0),
          'Всего гостей: ' + Number(data.totalGuests || 0),
          data.note ? 'Комментарий: ' + clean_(data.note) : '',
          'Язык формы: ' + clean_(data.language),
          'Мероприятие: ' + clean_(data.event),
          '',
          'Таблица ответов: https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit'
        ].filter(Boolean).join('\\n');
        MailApp.sendEmail(ownerEmail, subject, body);
      } else {
        console.warn('Could not resolve the deploying Google account email.');
      }
    } catch (mailErr) {
      console.error('RSVP was saved, but email notification failed: ' + mailErr);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function clean_(value) {
  const s = String(value == null ? '' : value).slice(0, 1000);
  return /^[=+@\\-]/.test(s) ? "'" + s : s;
}

function testSetup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Connected to: ' + ss.getName());
}
