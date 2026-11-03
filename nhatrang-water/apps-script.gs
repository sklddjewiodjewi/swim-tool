// Синхронизация лога «Прозрачность воды — пляж Нячанга», схема 2 (снимки + наблюдения)
// Вставить в Расширения → Apps Script у Google Таблицы. KEY — своё слово, его же вводишь на странице.
const KEY = 'ПРИДУМАЙ-СВОЁ-СЛОВО';
const SHEETS = { snaps: 'snaps', obs: 'obs_v2' };   // колонки: A ключ, B JSON, C когда сохранено, D версия записи

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out_({ error: 'bad key' });
  const since = e.parameter.since || '';
  const res = { now: new Date().toISOString() };
  for (const [name, sh] of Object.entries(SHEETS)) {
    const s = sheet_(sh), n = s.getLastRow();
    res[name] = [];
    if (n < 1) continue;
    s.getRange(1, 1, n, 3).getValues().forEach(r => {
      if (!r[0] || String(r[2]) <= since) return;
      try { res[name].push(JSON.parse(r[1])); } catch (err) {}
    });
  }
  return out_(res);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad json' }); }
  if (body.key !== KEY) return out_({ error: 'bad key' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const now = new Date().toISOString();
    for (const [name, sh] of Object.entries(SHEETS)) {
      const items = body[name] || [];
      if (!items.length) continue;
      const s = sheet_(sh), n = s.getLastRow(), idx = {};
      if (n > 0) s.getRange(1, 1, n, 4).getValues().forEach((r, i) => { if (r[0]) idx[String(r[0])] = { row: i + 1, ver: String(r[3]) }; });
      const append = [];
      items.forEach(it => {
        const id = name === 'snaps' ? it.key : it.id;
        if (!id) return;
        const ver = it.v || it.updated || '';
        const row = ["'" + id, JSON.stringify(it), "'" + now, "'" + ver];
        const cur = idx[id];
        if (!cur) { append.push(row); idx[id] = { row: -1, ver }; }
        else if (cur.row > 0 && ver >= cur.ver) { s.getRange(cur.row, 1, 1, 4).setValues([row]); cur.ver = ver; }
      });
      if (append.length) s.getRange(s.getLastRow() + 1, 1, append.length, 4).setValues(append);
    }
    return out_({ ok: true, now });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
