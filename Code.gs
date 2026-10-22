/**
 * Sporto Lyga 2027 – duomenų bazė Google Sheets lentelėje.
 * Įklijuokite į: Google Sheets → Plėtiniai (Extensions) → Apps Script.
 * Diegimas: Deploy → New deployment → Web app,
 *   Execute as: Me, Who has access: Anyone.
 */

const SHEET_NAME = 'Veiklos';
const ATHLETES = ['Edvinas', 'Tomas', 'Vygantas', 'Justas', 'Paulius', 'Eduardas'];
const SPORTS = ['run', 'bike', 'swim', 'xc', 'row', 'gym', 'padel', 'racket', 'team', 'class', 'climb', 'skate', 'ski', 'yoga'];
const KM_SPORTS = ['run', 'bike', 'swim', 'xc', 'row'];
// Testavimas: datos nuo 2026-10-22 leidžiamos; programėlė jų neskaičiuoja nuo 2027-01-01.
const HEADERS = ['id', 'athlete', 'sport', 'amount', 'date', 'note', 'createdAt'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@'); // data kaip tekstas
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function doGet() {
  const values = sheet_().getDataRange().getValues();
  values.shift();
  const entries = values.filter(r => r[0]).map(r => ({
    id: String(r[0]),
    athlete: String(r[1]),
    sport: String(r[2]),
    amount: Number(r[3]),
    date: dateStr_(r[4]),
    note: String(r[5] || ''),
    createdAt: r[6] instanceof Date ? r[6].toISOString() : String(r[6] || '')
  }));
  return json_({ ok: true, entries: entries });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();

    if (body.action === 'add') {
      const amount = Number(body.amount);
      const date = String(body.date || '');
      const maxAmount = KM_SPORTS.indexOf(body.sport) >= 0 ? 300 : 720;
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      if (SPORTS.indexOf(body.sport) < 0) return json_({ ok: false, error: 'sport' });
      if (!(amount > 0) || amount > maxAmount) return json_({ ok: false, error: 'amount' });
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < '2026-10-22' || date > '2027-03-01') return json_({ ok: false, error: 'date' });
      const id = Utilities.getUuid();
      sh.appendRow([id, body.athlete, body.sport, amount, date, String(body.note || '').slice(0, 80), new Date().toISOString()]);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'delete') {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(body.id)) { sh.deleteRow(i + 1); break; }
      }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
