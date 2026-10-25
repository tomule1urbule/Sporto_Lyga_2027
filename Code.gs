/**
 * Sporto Lyga 2027 – duomenų bazė Google Sheets lentelėje.
 * Įklijuokite į: Google Sheets → Plėtiniai (Extensions) → Apps Script.
 * Diegimas: Deploy → New deployment → Web app,
 *   Execute as: Me, Who has access: Anyone.
 */

const SHEET_NAME = 'Veiklos';
const ATHLETES = ['Edvinas', 'Tomas', 'Vygantas', 'Justas', 'Paulius', 'Eduardas', 'Dainius'];
// Veiklų sąrašas tikrinamas programėlėje, todėl keičiant veiklas šio kodo keisti nereikia.
const HEADERS = ['id', 'athlete', 'sport', 'amount', 'date', 'note', 'createdAt', 'reactions', 'minutes', 'image'];
const REACTIONS = ['💪', '🤯', '🖕', '👍'];
// Mažiausia užskaitoma veikla (km arba min.)
const MINS = { run: 3, walk: 3, bike: 10, swim: 0.3, ibike: 20, gym: 20, padel: 20, basket: 20, hiit: 20, martial: 20, xc: 20, ski: 20, hunt: 45 };

// Reakcijos saugomos H stulpelyje: "Vardas:emoji,Vardas:emoji"
function parseReactions_(v) {
  const out = {};
  String(v || '').split(',').filter(Boolean).forEach(p => { const i = p.indexOf(':'); if (i > 0) out[p.slice(0, i)] = p.slice(i + 1); });
  return out;
}
// Žaidėjų spalvos saugomos skripto nustatymuose (p0–p17)
const COLORS_KEY = 'colors_v2';
function colors_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(COLORS_KEY) || '{}'); } catch (e) { return {}; } }
const CHAT_SHEET = 'Pokalbiai';
const CHAT_HEADERS = ['id', 'athlete', 'text', 'createdAt', 'likes', 'image', 'replyTo'];
// Nuotraukos saugomos atskirame lape dalimis (langelyje telpa iki 50 000 simbolių)
const IMG_SHEET = 'Nuotraukos';
const IMG_CHUNK = 45000;
function imgSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IMG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IMG_SHEET);
    sh.appendRow(['imageId', 'part', 'data']);
    sh.setFrozenRows(1);
    sh.getRange('A:C').setNumberFormat('@');
  }
  // Nereikalingi stulpeliai pašalinami – mažiau langelių, greitesnės kopijos
  if (sh.getMaxColumns() > 3) sh.deleteColumns(4, sh.getMaxColumns() - 3);
  return sh;
}
// Nuotrauka išsaugoma dalimis; grąžinamas jos ID
function storeImage_(image) {
  const imageId = Utilities.getUuid();
  const rows = [];
  for (let i = 0, p = 0; i < image.length; i += IMG_CHUNK, p++) rows.push([imageId, String(p), image.slice(i, i + IMG_CHUNK)]);
  const ish = imgSheet_();
  ish.getRange(ish.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  return imageId;
}
function validImage_(image, max) {
  return /^data:image\/(jpeg|png|webp);base64,/.test(image) && image.length <= max;
}
function readImage_(id) {
  // Pirma skaitomas tik trumpas imageId stulpelis, tada – tik šios nuotraukos eilutės
  const sh = imgSheet_();
  const last = sh.getLastRow();
  if (last < 2) return '';
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  let first = -1, end = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { if (first < 0) first = i; end = i; }
  }
  if (first < 0) return '';
  const rows = sh.getRange(first + 2, 1, end - first + 1, 3).getValues().filter(r => String(r[0]) === String(id));
  rows.sort((a, b) => Number(a[1]) - Number(b[1]));
  return rows.map(r => String(r[2])).join('');
}

// ===== Naktinė atsarginė kopija =====
// Kiekvieną naktį visi lapai nukopijuojami į atskirą lentelę „Sporto Lyga atsarginės kopijos“ (saugoma 14 dienų).
// Nuotraukų lapas kopijuojamas tik sekmadieniais (jis didžiausias).
// Paleidimas kasdien 03–04 val. nustatytas per Apps Script → Triggers.
const BACKUP_DAYS = 14;
function backupNightly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const props = PropertiesService.getScriptProperties();
  let bk = null;
  const bid = props.getProperty('backup_id');
  if (bid) { try { bk = SpreadsheetApp.openById(bid); } catch (e) { bk = null; } }
  if (!bk) { bk = SpreadsheetApp.create('Sporto Lyga atsarginės kopijos'); props.setProperty('backup_id', bk.getId()); }
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm');
  const photos = Utilities.formatDate(new Date(), tz, 'u') === '7' || !bk.getSheets().some(s => s.getName().slice(-IMG_SHEET.length - 1) === ' ' + IMG_SHEET);
  ss.getSheets().forEach(sh => { if (sh.getName() !== IMG_SHEET || photos) sh.copyTo(bk).setName(stamp + ' ' + sh.getName()); });
  // Seni (ir tušti pradiniai) lapai ištrinami
  const cutoff = Utilities.formatDate(new Date(Date.now() - BACKUP_DAYS * 864e5), tz, 'yyyy-MM-dd');
  bk.getSheets().forEach(s => {
    const d = s.getName().slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d < cutoff) bk.deleteSheet(s);
  });
  return bk.getUrl();
}
function deleteImage_(id) {
  if (!id) return;
  const sh = imgSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) { if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2); }
}

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

function doGet(e) {
  if (e && e.parameter && e.parameter.img) {
    const data = readImage_(e.parameter.img);
    return json_({ ok: !!data, data: data });
  }
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
    reactions: parseReactions_(r[7]),
    minutes: Number(r[8]) || 0,
    image: String(r[9] || '')
  }));
  const cv = chatSheet_().getDataRange().getValues();
  cv.shift();
  const messages = cv.filter(r => r[0]).slice(-200).map(r => ({
    likes: String(r[4] || '').split(',').filter(Boolean),
    image: String(r[5] || ''),
    replyTo: String(r[6] || ''),
    id: String(r[0]),
    athlete: String(r[1]),
    text: String(r[2] || '').replace(/^'/, ''),
    createdAt: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || '').replace(/^'/, '')
  }));
  return json_({ ok: true, likesV: 1, reactV: 1, colorsV: 1, imgV: 1, actImgV: 1, editV: 1, replyV: 1, colors: colors_(), entries: entries, messages: messages });
}

// Bendras veiklos duomenų tikrinimas (pridedant ir redaguojant)
function checkEntry_(body) {
  const amount = Number(body.amount);
  const date = String(body.date || '');
  if (ATHLETES.indexOf(body.athlete) < 0) return { error: 'athlete' };
  if (!/^[a-z]{2,12}$/.test(String(body.sport))) return { error: 'sport' };
  if (!(amount > 0) || amount > 720) return { error: 'amount' };
  if (MINS[body.sport] != null && amount < MINS[body.sport] - 1e-9) return { error: 'too_short' };
  // Ateities datų neleidžiama (+1 d. atsarga dėl laiko juostų)
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 864e5), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > tomorrow) return { error: 'date' };
  const minutes = Number(body.minutes) || '';
  if (minutes !== '' && (minutes < 1 || minutes > 720)) return { error: 'minutes' };
  return { amount: amount, date: date, minutes: minutes, note: String(body.note || '').slice(0, 80) };
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();

    if (body.action === 'add') {
      const c = checkEntry_(body);
      if (c.error) return json_({ ok: false, error: c.error });
      const amount = c.amount, date = c.date, minutes = c.minutes;
      const id = Utilities.getUuid();
      const image = String(body.image || '');
      if (image && !validImage_(image, 400000)) return json_({ ok: false, error: 'image' });
      const imageId = image ? storeImage_(image) : '';
      // I stulpelis – trukmė minutėmis (nuotolio veikloms), J – veiklos nuotraukos ID
      sh.appendRow([id, body.athlete, body.sport, amount, date, c.note, new Date().toISOString(), '', minutes, imageId]);
      return json_({ ok: true, id: id, image: imageId });
    }

    if (body.action === 'edit') {
      // Redaguoti galima tik savo veiklą (ne bonusą)
      const c = checkEntry_(body);
      if (c.error) return json_({ ok: false, error: c.error });
      if (body.sport === 'bonus') return json_({ ok: false, error: 'sport' });
      const ids = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) !== String(body.id)) continue;
        if (String(ids[i][1]) !== body.athlete) return json_({ ok: false, error: 'forbidden' });
        const row = i + 1;
        const oldImg = String(sh.getRange(row, 10).getValue() || '');
        if (String(sh.getRange(row, 3).getValue()) === 'bonus') return json_({ ok: false, error: 'sport' });
        let imageId = oldImg;
        const image = String(body.image || '');
        if (image) {
          if (!validImage_(image, 400000)) return json_({ ok: false, error: 'image' });
          imageId = storeImage_(image);
          deleteImage_(oldImg);
        } else if (body.removeImage) {
          deleteImage_(oldImg);
          imageId = '';
        }
        sh.getRange(row, 3, 1, 4).setValues([[body.sport, c.amount, c.date, c.note]]);
        sh.getRange(row, 9, 1, 2).setValues([[c.minutes, imageId]]);
        return json_({ ok: true, id: String(body.id), image: imageId });
      }
      return json_({ ok: false, error: 'not_found' });
    }

    if (body.action === 'chat') {
      const text = String(body.text || '').trim().slice(0, 500);
      if (ATHLETES.indexOf(body.athlete) < 0) return json_({ ok: false, error: 'athlete' });
      const image = String(body.image || '');
      if (image && !validImage_(image, 1500000)) return json_({ ok: false, error: 'image' });
      if (!text && !image) return json_({ ok: false, error: 'text' });
      const id = Utilities.getUuid();
      const imageId = image ? storeImage_(image) : '';
      const replyTo = /^[\w-]{1,64}$/.test(String(body.replyTo || '')) ? String(body.replyTo) : '';
      // C ir D stulpeliai suformatuoti kaip tekstas, todėl Sheets jų nepaverčia formule ar data; G – į kurią žinutę atsakoma
      chatSheet_().appendRow([id, body.athlete, text, new Date().toISOString(), '', imageId, replyTo]);
      return json_({ ok: true, id: id, image: imageId });
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
      if (!/^p\d{1,2}$/.test(color) || Number(color.slice(1)) > 17) return json_({ ok: false, error: 'color' });
      const all = colors_();
      for (const n in all) { if (n !== body.athlete && all[n] === color) return json_({ ok: false, error: 'taken' }); }
      all[body.athlete] = color;
      PropertiesService.getScriptProperties().setProperty(COLORS_KEY, JSON.stringify(all));
      return json_({ ok: true, colors: all });
    }

    if (body.action === 'deleteChat') {
      const csh = chatSheet_();
      const last = csh.getLastRow();
      if (last >= 2) {
        const rows = csh.getRange(2, 1, last - 1, 6).getValues();
        for (const r of rows) { if (String(r[0]) === String(body.id)) { deleteImage_(r[5]); break; } }
      }
      deleteById_(csh, body.id);
      return json_({ ok: true });
    }

    if (body.action === 'delete') {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(body.id)) {
          const img = sh.getMaxColumns() >= 10 ? String(sh.getRange(i + 1, 10).getValue() || '') : '';
          sh.deleteRow(i + 1);
          deleteImage_(img);
          break;
        }
      }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
