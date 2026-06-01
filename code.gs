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
    .createTextOutput('Web App attiva')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = e && e.parameter ? e.parameter : {};

    const q1 = data.q1 || '';
    const email = data.email || '';
    const ua = data.ua || '';
    const tsISO = data.tsISO || new Date().toISOString();

    const result = appendSurvey(
      {
        q1: q1,
        email: email
      },
      ua,
      tsISO
    );

    return ContentService
      .createTextOutput('OK - scritto riga ' + result.lastRow)
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService
      .createTextOutput('ERRORE: ' + (err.message || String(err)))
      .setMimeType(ContentService.MimeType.TEXT);
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

  return {
    ok: true,
    lastRow: sh.getLastRow()
  };
}

function testAppend() {
  appendSurvey(
    {
      q1: 'TEST',
      email: 'test@example.com'
    },
    'manual-test',
    new Date().toISOString()
  );
}
