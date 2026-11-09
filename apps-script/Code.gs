// 對聯闖關・成績後端（Google Apps Script）
// 使用方式：
// 1. 在 Google 試算表：擴充功能 → Apps Script，把這個檔案全部貼上
// 2. 把第 8 行的密碼改成你自己的（不能是 change-me）
// 3. 部署 → 新增部署作業 → 類型「網頁應用程式」，執行身分「我」，存取權「所有人」
// 4. 把網址（結尾是 /exec）交給 Claude 填進 config.js

const PASSWORD = 'change-me';

const DATA_SHEET = 'data';         // 給程式讀：每人每項一列
const LOG_SHEET = '作答紀錄';       // 給老師看：每次作答一列
const MAX_ATTEMPTS = 20;
const SUS_SECONDS = 30;
const ITEMS = {
  L1: '第一關 字字相對', L2: '第二關 分辨上下聯', L3: '第三關 入聲字陷阱', L4: '第四關 找出下聯',
  L5: '第五關 對聯小知識', L6: '第六關 對聯貼哪裡', L7: '第七關 動手貼春聯', L8: '第八關 抽年獸對戰',
  R: '學習單'
};
const DATA_HEAD = ['key', 'cls', 'seat', 'name', 'item', 'best', 'bestScore', 'total', 'updated', 'attempts'];
const LOG_HEAD = ['完成時間', '班級', '座號', '姓名', '項目', '答對', '題數', '星星', '用時(秒)', '疑似亂答', '作答內容'];

function doGet() {
  return out({ ok: true, app: '對聯闖關成績後端', time: new Date().toISOString() });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'bad_json' }); }
  try {
    if (b.action === 'list') return out(list(b));
    if (b.action === 'save') return out(save(b));
    if (b.action === 'clearClass') return out(clearClass(b));
    return out({ ok: false, error: 'unknown_action' });
  } catch (x) {
    return out({ ok: false, error: String(x) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function pwOk(p) {
  return PASSWORD !== 'change-me' && typeof p === 'string' && p.length > 0 && p === PASSWORD;
}

function getSheet(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
  }
  return sh;
}

function readData() {
  const sh = getSheet(DATA_SHEET, DATA_HEAD);
  const v = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < v.length; r++) {
    const o = {};
    DATA_HEAD.forEach((h, i) => { o[h] = v[r][i]; });
    o._row = r + 1;
    rows.push(o);
  }
  return rows;
}

function parseJson(t, d) { try { return JSON.parse(t); } catch (e) { return d; } }
function str(v, n) { return String(v == null ? '' : v).slice(0, n); }

function list(b) {
  const teacher = pwOk(b.password);
  const all = readData();
  const classes = Array.from(new Set(all.map(r => String(r.cls)))).filter(c => c);
  const rows = all.filter(r => !b.cls || String(r.cls) === String(b.cls)).map(r => {
    const o = { cls: String(r.cls), seat: String(r.seat), name: String(r.name), item: String(r.item),
      best: Number(r.best) || 0, bestScore: Number(r.bestScore) || 0, total: Number(r.total) || 0 };
    if (teacher) { o.updated = String(r.updated); o.attempts = parseJson(r.attempts, []); }
    return o;
  }).filter(o => teacher || o.item !== 'R');
  return { ok: true, teacher: teacher, classes: classes, rows: rows };
}

function cleanAttempt(a) {
  a = a || {};
  const c = {
    at: str(a.at, 40) || new Date().toISOString(),
    dur: Math.max(0, Math.round(Number(a.dur) || 0)),
    score: Math.max(0, Number(a.score) || 0),
    total: Math.max(0, Number(a.total) || 0),
    stars: Math.max(0, Math.min(24, Number(a.stars) || 0)),
    ans: (Array.isArray(a.ans) ? a.ans : []).slice(0, 40).map(x => ({
      q: str(x && x.q, 60), p: (x && typeof x.p === 'number') ? x.p : null, pick: str(x && x.pick, 40), ok: !!(x && x.ok) }))
  };
  if (a.reflect) {
    c.reflect = { say: str(a.reflect.say, 60),
      skills: (Array.isArray(a.reflect.skills) ? a.reflect.skills : []).filter(i => typeof i === 'number').slice(0, 10) };
  }
  return c;
}

function isSus(item, a) {
  if (item === 'R') return false;
  if (a.dur < SUS_SECONDS) return true;
  const ps = a.ans.map(x => x.p);
  return ps.length >= 3 && ps.every(p => p !== null && p === ps[0]) && a.score < a.total;
}

function save(b) {
  const s = b.student || {};
  const cls = str(s.cls, 10).trim(), seat = str(s.seat, 5).trim(), name = str(s.name, 20).trim();
  const item = str(b.item, 4);
  if (!cls || !seat || !name) return { ok: false, error: 'missing_student' };
  if (!ITEMS[item]) return { ok: false, error: 'bad_item' };
  const att = cleanAttempt(b.attempt);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet(DATA_SHEET, DATA_HEAD);
    const key = cls + '|' + seat;
    const row = readData().find(r => String(r.key) === key && String(r.item) === item);
    let attempts = row ? parseJson(row.attempts, []) : [];
    const isNew = !attempts.some(a => a.at === att.at);
    if (isNew) attempts.push(att);
    attempts.sort((x, y) => (x.at < y.at ? -1 : 1));
    attempts = attempts.slice(-MAX_ATTEMPTS);
    const best = Math.max(row ? Number(row.best) || 0 : 0, item === 'R' ? 0 : att.stars);
    const bestScore = Math.max(row ? Number(row.bestScore) || 0 : 0, att.score);
    const total = att.total || (row ? Number(row.total) || 0 : 0);
    const updated = attempts[attempts.length - 1].at;
    const values = [key, cls, seat, name, item, best, bestScore, total, updated, JSON.stringify(attempts)];
    if (row) sh.getRange(row._row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);
    if (isNew) {
      const log = getSheet(LOG_SHEET, LOG_HEAD);
      const when = Utilities.formatDate(new Date(att.at), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      const text = item === 'R'
        ? (att.reflect ? att.reflect.say : '')
        : att.ans.map((x, i) => (i + 1) + '.' + (x.ok ? '○' : '×') + x.pick).join(' ');
      log.appendRow([when, cls, seat, name, ITEMS[item], att.score, att.total, att.stars, att.dur, isSus(item, att) ? '是' : '', text]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function clearClass(b) {
  if (!pwOk(b.password)) return { ok: false, error: 'wrong_password' };
  const cls = String(b.cls || '');
  if (!cls) return { ok: false, error: 'missing_class' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let removed = 0;
    [[DATA_SHEET, DATA_HEAD, 2], [LOG_SHEET, LOG_HEAD, 2]].forEach(([n, h, col]) => {
      const sh = getSheet(n, h);
      const v = sh.getDataRange().getValues();
      for (let r = v.length - 1; r >= 1; r--) {
        if (String(v[r][col - 1]) === cls) { sh.deleteRow(r + 1); removed++; }
      }
    });
    return { ok: true, removed: removed };
  } finally {
    lock.releaseLock();
  }
}
