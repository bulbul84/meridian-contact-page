/**
 * Meridian Studio — contact form endpoint
 * Paste this into Google Apps Script (Extensions > Apps Script) of a Google Sheet,
 * then Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

const SHEET_NAME = 'Submissions';
const NOTIFY_EMAIL = 'hello@meridian.studio'; // set to '' to skip email notification

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.subject || '',
      data.budget || '',
      data.reference || '',
      data.message || ''
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New contact message: ' + (data.subject || 'general'),
        replyTo: data.email || NOTIFY_EMAIL,
        body:
          'Name: ' + (data.name || '') + '\n' +
          'Email: ' + (data.email || '') + '\n' +
          'Subject: ' + (data.subject || '') + '\n' +
          'Budget: ' + (data.budget || '') + '\n' +
          'Reference: ' + (data.reference || '') + '\n\n' +
          (data.message || '')
      });
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, message: 'Contact endpoint is alive.' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Subject', 'Budget', 'Reference', 'Message']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
