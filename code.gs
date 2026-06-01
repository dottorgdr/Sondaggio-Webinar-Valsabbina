const SHEET_ID = '1vChb-vtCS1eSaAzOYpuA3f9htOd4nt0AC307TkN6NXg';
const SHEET_NAME = 'Sondaggio Valsabbina';

const HEADERS = [
  'Timestamp',
  'UserAgent',
  'Email',
  'Risposta'
];

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Sondaggio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
