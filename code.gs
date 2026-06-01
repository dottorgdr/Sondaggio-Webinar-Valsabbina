const SHEET_ID = '1vChb-vtCS1eSaAzOYpuA3f9htOd4nt0AC307TkN6NXg';
const SHEET_NAME = 'Sondaggio Valsabbina';

const HEADERS = [
  'Timestamp',
  'UserAgent',
  'Email',
  'Risposta'
];

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Web App attiva' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    appendSurvey(
      {
        q1: data.q1 || '',
        email: data.email || ''
      },
      data.ua || '',
      data.tsISO || new Date().toISOString()
    );

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: err.message || String(err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function appendSurvey(answers, ua, tsISO) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }

  const firstCell = sh.getRange(1, 1).getValue();

  if (!firstCell) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  sh.appendRow([
    tsISO || new Date().toISOString(),
    (ua || '').substring(0, 512),
    answers.email || '',
    answers.q1 || ''
  ]);

  SpreadsheetApp.flush();

  return { ok: true };
}
