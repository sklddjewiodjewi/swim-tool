// Общая история заплывов — одна на все устройства. Хранится в Google Таблице, лист history:
// колонки A дата (ГГГГ-ММ-ДД), B запись JSON, C когда сохранено. Одна строка на дату — дубликатов не бывает.
// Без ключа: читать и писать может любой, у кого есть адрес веб-приложения (по решению пользователя).
// Установка: README.md → «Общая история (Google Таблица)».
const SHEET = 'history';
const SEED_URL = 'https://sklddjewiodjewi.github.io/swim-tool/data/history.json'; // первое заполнение пустой таблицы

function doGet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (sheet_().getLastRow() < 1) seed_();
    return out_({ ok: true, list: list_() });
  } finally {
    lock.releaseLock();
  }
}

// Тело (text/plain, чтобы без preflight): { ops: [ {date, rec} | {date, del: true} ] }
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad json' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const s = sheet_(), now = new Date().toISOString();
    (body.ops || []).forEach(op => {
      const d = String(op.date || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
      const row = rowOf_(s, d);
      if (op.del) { if (row) s.deleteRow(row); return; }
      const vals = [["'" + d, JSON.stringify(op.rec), "'" + now]];
      if (row) s.getRange(row, 1, 1, 3).setValues(vals);
      else s.getRange(s.getLastRow() + 1, 1, 1, 3).setValues(vals);
    });
    return out_({ ok: true, list: list_() });
  } finally {
    lock.releaseLock();
  }
}

function list_() {
  const s = sheet_(), n = s.getLastRow(), by = {};
  if (n > 0) s.getRange(1, 1, n, 2).getValues().forEach(r => {
    if (!r[0]) return;
    try { by[String(r[0])] = JSON.parse(r[1]); } catch (err) {}
  });
  return Object.values(by).sort((a, b) => b.date.localeCompare(a.date));
}

function rowOf_(s, d) {
  const n = s.getLastRow();
  if (n < 1) return 0;
  const col = s.getRange(1, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(col[i][0]) === d) return i + 1;
  return 0;
}

function seed_() {
  const list = JSON.parse(UrlFetchApp.fetch(SEED_URL + '?t=' + Date.now()).getContentText());
  const now = new Date().toISOString(), by = {};
  list.forEach(r => { if (r && r.date) by[r.date] = r; });
  const rows = Object.values(by).map(r => ["'" + r.date, JSON.stringify(r), "'" + now]);
  if (rows.length) sheet_().getRange(1, 1, rows.length, 3).setValues(rows);
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
