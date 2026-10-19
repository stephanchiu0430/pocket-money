/**
 * 零用錢帳本：爸爸的 Google 試算表（正本）
 *
 * 安裝：
 * 1. 新建一個 Google 試算表 → 擴充功能 → Apps Script，把這整個檔案貼上、存檔。
 * 2. 上方選 setup → 執行（第一次會要求授權）。
 * 3. 部署 → 新增部署作業 → 類型「網頁應用程式」
 *    執行身分：我；誰可以存取：所有人 → 部署 → 複製結尾是 /exec 的網址。
 * 4. 在孩子手機的 App 第一次設定時貼上這個網址。
 *
 * 孩子沒有這個試算表的權限。這支程式只會「新增」紀錄，沒有修改或刪除的功能。
 */

const KIDS = ['Chantel', 'Luke', 'Ian'];
const T = { SUM: '總覽', LOG: '紀錄', PARENT: '家長入帳', SET: '設定' };
const LOG_HEAD = ['收到時間', '孩子', '時間', '類型', '金額', '分類', '說明', '確認', '警示',
  'id', 'seq', '裝置', 'ref', 'expected', 'flags', 'prev', 'hash'];
const C = {}; LOG_HEAD.forEach((h, i) => C[h] = i);
const TYPE_ZH = { out: '支出', in: '收入', void: '撤銷', count: '盤點' };
const TYPE_CODE = { '支出': 'out', '收入': 'in', '撤銷': 'void', '盤點': 'count' };
const TZ = 'Asia/Taipei';

/* ---------- 第一次安裝 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  const sum = getOrMake_(T.SUM);
  const log = getOrMake_(T.LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(LOG_HEAD); log.setFrozenRows(1);
    if (log.getMaxRows() < 5000) log.insertRowsAfter(log.getMaxRows(), 5000 - log.getMaxRows());
    ['分類', '說明', 'id', '裝置', 'ref', 'flags', 'prev', 'hash'].forEach(h =>
      log.getRange(2, C[h] + 1, log.getMaxRows() - 1, 1).setNumberFormat('@'));
    log.getRange(2, C['確認'] + 1, log.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['—', '待確認', '已確認', '不算'], true).build());
    log.getRange(2, C['時間'] + 1, log.getMaxRows() - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    log.getRange(2, C['收到時間'] + 1, log.getMaxRows() - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    log.hideColumns(C['id'] + 1, LOG_HEAD.length - C['id']);
  }

  const par = getOrMake_(T.PARENT);
  if (par.getLastRow() === 0) {
    par.appendRow(['日期', '孩子', '金額', '項目', '說明']); par.setFrozenRows(1);
    par.getRange('A2:A').setNumberFormat('yyyy/MM/dd');
    par.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KIDS, true).build());
    par.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['零用錢', '獎勵', '扣款', '其他'], true).build());
  }

  const set = getOrMake_(T.SET);
  if (set.getLastRow() === 0) {
    set.appendRow(['孩子', '每週零用錢', '發放日（1＝週一 … 7＝週日）', '自動入帳']);
    KIDS.forEach(k => set.appendRow([k, 0, 6, false]));
    set.getRange(2, 4, KIDS.length, 1).insertCheckboxes();
  }

  ss.setActiveSheet(sum); ss.moveActiveSheet(1);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyAllowance')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyAllowance').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  updateSummary();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('零用錢')
    .addItem('更新總覽', 'updateSummary')
    .addItem('今天的零用錢補發', 'dailyAllowance')
    .addToUi();
}

/* ---------- App 呼叫的入口 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (KIDS.indexOf(req.child) < 0) return out_({ ok: false, error: '不認得這個孩子' });
    if (req.action === 'push') pushEntries_(req.child, req.entries || []);
    const entries = computeFor_(req.child).entries;
    updateSummary();
    return out_({ ok: true, entries: entries });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
function doGet() { return out_({ ok: true, hello: '零用錢帳本' }); }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 收紀錄：只新增，並檢查有沒有被動過 ---------- */
function pushEntries_(child, entries) {
  const sh = sheet_(T.LOG);
  const props = PropertiesService.getScriptProperties();
  const existing = new Set(colValues_(sh, C['id']));
  const cur = computeFor_(child);
  let bal = cur.bal;
  const rows = [];

  entries.slice().sort((a, b) => a.seq - b.seq).forEach(x => {
    if (!x || !x.id || existing.has(x.id)) return;
    const alerts = [];
    if (x.child !== child) alerts.push('孩子不符');
    if (hashOf_(x) !== x.hash) alerts.push('內容被改過');

    const key = 'dev:' + x.device;
    const last = JSON.parse(props.getProperty(key) || 'null');
    if (!last) {
      if (Number(x.seq) !== 1 || x.prev !== 'GENESIS') alerts.push('前面缺紀錄');
      else if (otherDevices_(child, x.device).length) alerts.push('帳本重新開始（App 被清除或換手機）');
      addDevice_(child, x.device);
    } else {
      if (Number(x.seq) !== last.seq + 1 || x.prev !== last.hash) alerts.push('前後不連續');
      if (x.ts < last.ts) alerts.push('時間倒退');
    }
    if (new Date(x.ts).getTime() > Date.now() + 10 * 60 * 1000) alerts.push('時間在未來');

    const amt = Number(x.amount) || 0;
    const flags = String(x.flags || '');
    if (x.type === 'out') { if (amt > bal) alerts.push('超過餘額'); bal -= amt; }
    if (x.type === 'void') {
      const o = cur.byId[x.ref];
      if (o && o.type === 'out') bal += o.amount;
      if (o && o.type === 'in' && o.status === 'ok') bal -= o.amount;
    }
    if (flags.indexOf('big:notasked') >= 0) alerts.push('大額，沒先商量');
    else if (flags.indexOf('big:asked') >= 0) alerts.push('大額，說有商量過');
    if (x.type === 'count') {
      const d = amt - Number(x.expected);
      if (d !== 0) alerts.push('盤點差 ' + (d > 0 ? '+' : '') + d);
    }

    props.setProperty(key, JSON.stringify({ seq: Number(x.seq), hash: x.hash, ts: x.ts }));
    rows.push([new Date(), child, new Date(x.ts), TYPE_ZH[x.type] || x.type,
      x.type === 'void' ? '' : amt, x.cat || '', x.note || '',
      x.type === 'in' ? '待確認' : '—', alerts.join('；'),
      x.id, Number(x.seq), x.device, x.ref || '',
      (x.expected === '' || x.expected == null) ? '' : Number(x.expected),
      flags, x.prev, x.hash]);
    existing.add(x.id);
  });

  if (!rows.length) return;
  const start = sh.getLastRow() + 1;
  const need = start + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 500);
  sh.getRange(start, 1, rows.length, LOG_HEAD.length).setValues(rows);
}

// 必須跟 App 的 canon() 完全一樣
function hashOf_(x) {
  const s = [x.id, x.seq, x.ts, x.child, x.device, x.type, x.amount, x.cat, x.note, x.ref, x.expected, x.flags, x.prev]
    .map(v => v == null ? '' : String(v)).join('|');
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return b.map(v => ((v + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/* ---------- 計算（跟 App 用同一套規則） ---------- */
function readKid_(child) {
  const sh = sheet_(T.LOG), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, LOG_HEAD.length).getValues()
    .filter(r => r[C['孩子']] === child && r[C['id']])
    .map(r => ({
      id: String(r[C['id']]), seq: Number(r[C['seq']]), ts: iso_(r[C['時間']]),
      type: TYPE_CODE[r[C['類型']]] || String(r[C['類型']]),
      amount: Number(r[C['金額']]) || 0, cat: String(r[C['分類']]), note: String(r[C['說明']]),
      ref: String(r[C['ref']]), expected: r[C['expected']] === '' ? '' : Number(r[C['expected']]),
      flags: String(r[C['flags']]), device: String(r[C['裝置']]),
      status: statusOf_(r[C['確認']]), alert: String(r[C['警示']]), source: 'kid'
    }));
}
function statusOf_(v) {
  v = String(v);
  return v === '已確認' || v === '—' ? 'ok' : v === '不算' ? 'rejected' : 'pending';
}
function readParent_(child) {
  const sh = sheet_(T.PARENT), n = sh.getLastRow();
  if (n < 2) return [];
  const list = [];
  sh.getRange(2, 1, n - 1, 5).getValues().forEach((r, i) => {
    if (r[1] !== child || r[0] === '' || r[2] === '') return;
    const item = String(r[3] || '零用錢');
    let amt = Number(r[2]) || 0;
    if (item === '扣款') amt = -Math.abs(amt);
    list.push({
      id: 'P-' + (i + 2), seq: 0, ts: iso_(r[0]), type: amt < 0 ? 'adj' : 'in',
      amount: amt, cat: item, note: String(r[4] || ''), source: 'parent', status: 'ok'
    });
  });
  return list;
}
function computeFor_(child) {
  const entries = readKid_(child).concat(readParent_(child))
    .sort((a, b) => a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : (a.seq || 0) - (b.seq || 0));
  const voided = new Set(entries.filter(e => e.type === 'void').map(e => e.ref));
  const byId = {};
  let bal = 0, pendIn = 0;
  entries.forEach(e => {
    byId[e.id] = e;
    if (e.type === 'void' || voided.has(e.id)) return;
    if (e.type === 'out') bal -= e.amount;
    else if (e.type === 'adj') bal += e.amount;
    else if (e.type === 'in') {
      if (e.status === 'ok') bal += e.amount;
      else if (e.status === 'pending') pendIn += e.amount;
    }
  });
  return { entries: entries, bal: bal, pendIn: pendIn, byId: byId };
}

/* ---------- 總覽 ---------- */
function updateSummary() {
  const sh = sheet_(T.SUM);
  const head = ['孩子', '餘額', '待確認收入', '最近盤點', '盤點差額', '有警示的筆數', '最後一筆'];
  const rows = KIDS.map(k => {
    const c = computeFor_(k);
    const counts = c.entries.filter(e => e.type === 'count');
    const lc = counts[counts.length - 1];
    const kidRows = c.entries.filter(e => e.source === 'kid');
    const last = kidRows[kidRows.length - 1];
    return [k, c.bal, c.pendIn, lc ? new Date(lc.ts) : '', lc ? lc.amount - Number(lc.expected) : '',
      c.entries.filter(e => e.alert).length, last ? new Date(last.ts) : ''];
  });
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(2, 4, rows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm');
}

/* ---------- 每週零用錢（每天早上 7 點檢查） ---------- */
function dailyAllowance() {
  const set = sheet_(T.SET).getDataRange().getValues().slice(1);
  const now = new Date();
  const wd = Number(Utilities.formatDate(now, TZ, 'u'));
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const ps = sheet_(T.PARENT);
  const prows = ps.getLastRow() > 1 ? ps.getRange(2, 1, ps.getLastRow() - 1, 5).getValues() : [];
  set.forEach(r => {
    const kid = r[0], amt = Number(r[1]), day = Number(r[2]), auto = r[3] === true;
    if (!auto || !amt || day !== wd) return;
    const done = prows.some(p => p[1] === kid && p[3] === '零用錢' && p[0] instanceof Date &&
      Utilities.formatDate(p[0], TZ, 'yyyy-MM-dd') === today);
    if (!done) ps.appendRow([now, kid, amt, '零用錢', '每週零用錢（自動）']);
  });
}

/* ---------- 小工具 ---------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁「' + name + '」，請先執行 setup()');
  return sh;
}
function getOrMake_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function colValues_(sh, idx) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, idx + 1, n - 1, 1).getValues().map(r => String(r[0]));
}
function iso_(v) {
  if (v instanceof Date) return v.toISOString();
  const d = new Date(v);
  return isNaN(d) ? new Date(0).toISOString() : d.toISOString();
}
function otherDevices_(child, device) {
  const list = JSON.parse(PropertiesService.getScriptProperties().getProperty('kid:' + child) || '[]');
  return list.filter(d => d !== device);
}
function addDevice_(child, device) {
  const p = PropertiesService.getScriptProperties();
  const list = JSON.parse(p.getProperty('kid:' + child) || '[]');
  if (list.indexOf(device) < 0) { list.push(device); p.setProperty('kid:' + child, JSON.stringify(list)); }
}
