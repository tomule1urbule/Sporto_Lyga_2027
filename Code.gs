/**
 * Sporto Lyga 2027 – duomenų bazė Google Sheets lentelėje.
 * Įklijuokite į: Google Sheets → Plėtiniai (Extensions) → Apps Script.
 * Diegimas: Deploy → New deployment → Web app,
 *   Execute as: Me, Who has access: Anyone.
 */

const SHEET_NAME = 'Veiklos';
const ATHLETES = ['Edvinas', 'Tomas', 'Vygantas', 'Justas', 'Paulius', 'Eduardas'];
// Veiklų sąrašas tikrinamas programėlėje, todėl keičiant veiklas šio kodo keisti nereikia.
const HEADERS = ['id', 'athlete', 'sport', 'amount', 'date', 'note', 'createdAt', 'reactions'];
const REACTIONS = ['💪', '🤯', '🖕', '👍'];

// Reakcijos saugomos H stulpelyje: "Vardas:emoji,Vardas:emoji"
function parseReactions_(v) {
  const out = {};
  String(v || '').split(',').filter(Boolean).forEach(p => { const i = p.indexOf(':'); if (i > 0) out[p.slice(0, i)] = p.slice(i + 1); });
  return out;
}
// Žaidėjų spalvos saugomos skripto nustatymuose (p0–p9)
const COLORS_KEY = 'colors_v2';
function colors_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(COLORS_KEY) || '{}'); } catch (e) { return {}; } }
const CHAT_SHEET = 'Pokalbiai';
const CHAT_HEADERS = ['id', 'athlete', 'text', 'createdAt', 'likes'];

function chatSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CHAT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CHAT_SHEET);
    sh.appendRow(CHAT_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('C:D').setNumberFormat('@');
  }
  return sh;
}

function deleteById_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) {
    if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 1); break; }
  }
}

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
  // Lentelės (ne skripto) laiko juosta, kad data nepasislinktų per dieną
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).replace(/^'/, '').slice(0, 10);
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
    createdAt: r[6] instanceof Date ? r[6].toISOString() : String(r[6] || ''),
    reactions: parseReactions_(r[7])
  }));
  const cv = chatSheet_().getDataRange().getValues();
  cv.shift();
  const messages = cv.filter(r => r[0]).slice(-200).map(r => ({
    likes: String(r[4] || '').split(',').filter(Boolean),
    id: String(r[0]),
    athlete: String(r[1]),
    text: String(r[2] || '').replace(/^'/, ''),
    createdAt: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || '').replace(/^'/, '')
  }));
  return json_({ ok: true, likesV: 1, reactV: 1, colorsV: 1, colors: colors_(), entries: entries, messages: messages });
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
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      if (!/^[a-z]{2,12}$/.test(String(body.sport))) return json_({ ok: false, error: 'sport' });
      if (!(amount > 0) || amount > 720) return json_({ ok: false, error: 'amount' });
      // Ateities datų neleidžiama (+1 d. atsarga dėl laiko juostų)
      const tomorrow = Utilities.formatDate(new Date(Date.now() + 864e5), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > tomorrow) return json_({ ok: false, error: 'date' });
      const id = Utilities.getUuid();
      sh.appendRow([id, body.athlete, body.sport, amount, date, String(body.note || '').slice(0, 80), new Date().toISOString()]);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'chat') {
      const text = String(body.text || '').trim().slice(0, 500);
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      if (!text) return json_({ ok: false, error: 'text' });
      const id = Utilities.getUuid();
      // C ir D stulpeliai suformatuoti kaip tekstas, todėl Sheets jų nepaverčia formule ar data
      chatSheet_().appendRow([id, body.athlete, text, new Date().toISOString()]);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'like') {
      // Patiktukai saugomi E stulpelyje kaip vardų sąrašas, atskirtas kableliais
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      const csh = chatSheet_();
      const ids = csh.getRange(1, 1, csh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(body.id)) {
          const cell = csh.getRange(i + 1, 5);
          let names = String(cell.getValue() || '').split(',').filter(Boolean);
          names = names.filter(n => n !== body.athlete);
          if (body.on) names.push(body.athlete);
          cell.setValue(names.join(','));
          return json_({ ok: true, likes: names });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    if (body.action === 'react') {
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      const emoji = String(body.emoji || '');
      if (emoji && REACTIONS.indexOf(emoji) < 0) return json_({ ok: false, error: 'emoji' });
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(body.id)) {
          const cell = sh.getRange(i + 1, 8);
          const r = parseReactions_(cell.getValue());
          delete r[body.athlete];
          if (emoji) r[body.athlete] = emoji;
          cell.setValue(Object.keys(r).map(n => n + ':' + r[n]).join(','));
          return json_({ ok: true, reactions: r });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    if (body.action === 'color') {
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      const color = String(body.color || '');
      if (!/^p\d$/.test(color)) return json_({ ok: false, error: 'color' });
      const all = colors_();
      for (const n in all) { if (n !== body.athlete && all[n] === color) return json_({ ok: false, error: 'taken' }); }
      all[body.athlete] = color;
      PropertiesService.getScriptProperties().setProperty(COLORS_KEY, JSON.stringify(all));
      return json_({ ok: true, colors: all });
    }

    if (body.action === 'deleteChat') {
      deleteById_(chatSheet_(), body.id);
      return json_({ ok: true });
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
