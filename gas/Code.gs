/**
 * 줍줍 API — Google Apps Script bound to the "줍줍 DB" spreadsheet.
 * All requests are POST with a text/plain JSON body: { action, token, ...payload }.
 * Run setup() once from the editor, then deploy as a web app (실행: 나 / 액세스: 모든 사용자).
 */

const SCHEMA = {
  Items: ['id', 'createdAt', 'updatedAt', 'status', 'title', 'category', 'source', 'url', 'place', 'cond',
    'buyPrice', 'askPrice', 'retail', 'listPrice', 'soldPrice', 'boughtAt', 'listedAt', 'soldAt', 'imageId', 'memo'],
  Comps: ['id', 'itemId', 'price', 'source', 'date', 'cond', 'createdAt'],
  Settings: ['key', 'value'],
  Log: ['at', 'itemId', 'from', 'to']
};
const NUM_COLS = ['buyPrice', 'askPrice', 'retail', 'listPrice', 'soldPrice', 'price'];
const DATE_COLS = ['createdAt', 'updatedAt', 'boughtAt', 'listedAt', 'soldAt', 'date', 'at'];
const STATUSES = ['new', 'watch', 'pass', 'bought', 'selling', 'sold'];
const STAMP = { bought: 'boughtAt', selling: 'listedAt', sold: 'soldAt' };
const LOCKED = ['id', 'createdAt', 'updatedAt'];
const MAX_IMAGE_CHARS = 6 * 1024 * 1024;

const props_ = () => PropertiesService.getScriptProperties();

/* ---------- one-time setup (run from the editor) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트의 [확장 프로그램 > Apps Script]에서 열어 실행하세요.');
  Object.keys(SCHEMA).forEach(name => {
    const cols = SCHEMA[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#E2F7F0');
    sh.setFrozenRows(1);
  });
  const blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  const p = props_();
  p.setProperty('SHEET_ID', ss.getId());
  if (!p.getProperty('API_TOKEN')) p.setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('준비 완료. API 토큰: ' + p.getProperty('API_TOKEN'));
}
function showToken() { Logger.log('API 토큰: ' + props_().getProperty('API_TOKEN')); }
function rotateToken() {
  props_().setProperty('API_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  Logger.log('새 API 토큰: ' + props_().getProperty('API_TOKEN') + '  (앱 설정에서 다시 연결하세요)');
}

/* ---------- HTTP ---------- */
function doGet() { return json_({ ok: true, app: 'jupjup', version: 1 }); }

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const token = props_().getProperty('API_TOKEN');
    if (!token || req.token !== token) return json_({ ok: false, error: '토큰이 맞지 않아요' });
    const fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: '알 수 없는 요청이에요: ' + req.action });
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try { return json_({ ok: true, data: fn(req) }); } finally { lock.releaseLock(); }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

const ACTIONS = {
  bootstrap() {
    const settings = {};
    rows_('Settings').forEach(r => { try { settings[r.key] = JSON.parse(r.value); } catch (e) { settings[r.key] = r.value; } });
    return { items: rows_('Items'), comps: rows_('Comps'), settings: settings };
  },

  createItem(req) {
    const now = new Date().toISOString();
    const o = clean_('Items', req.item || {});
    if (!o.title) throw new Error('상품명이 없어요');
    o.id = newId_('i'); o.createdAt = now; o.updatedAt = now;
    o.status = STATUSES.indexOf(o.status) >= 0 ? o.status : 'new';
    if (STAMP[o.status] && !o[STAMP[o.status]]) o[STAMP[o.status]] = now;
    sheet_('Items').appendRow(toRow_('Items', o));
    return normalize_('Items', o);
  },

  updateItem(req) {
    const at = find_('Items', req.id);
    const cur = toObj_('Items', at.sheet.getRange(at.row, 1, 1, SCHEMA.Items.length).getValues()[0]);
    const patch = clean_('Items', req.patch || {});
    const now = new Date().toISOString();
    if (patch.status && patch.status !== cur.status) {
      if (STATUSES.indexOf(patch.status) < 0) throw new Error('알 수 없는 상태예요: ' + patch.status);
      const k = STAMP[patch.status];
      if (k && !cur[k] && !patch[k]) patch[k] = now;
      sheet_('Log').appendRow([new Date(), cur.id, cur.status, patch.status]);
    }
    const next = Object.assign({}, cur, patch, { updatedAt: now });
    at.sheet.getRange(at.row, 1, 1, SCHEMA.Items.length).setValues([toRow_('Items', next)]);
    return normalize_('Items', next);
  },

  deleteItem(req) {
    const at = find_('Items', req.id);
    at.sheet.deleteRow(at.row);
    const cs = sheet_('Comps'), last = cs.getLastRow();
    if (last >= 2) {
      const ids = cs.getRange(2, 2, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(req.id)) cs.deleteRow(i + 2);
    }
    return { id: req.id };
  },

  addComp(req) {
    const c = clean_('Comps', req.comp || {});
    if (!c.itemId) throw new Error('매물 id가 없어요');
    find_('Items', c.itemId);
    if (!(Number(c.price) > 0)) throw new Error('팔린 가격이 없어요');
    c.id = newId_('c'); c.createdAt = new Date().toISOString(); c.date = c.date || c.createdAt;
    sheet_('Comps').appendRow(toRow_('Comps', c));
    return normalize_('Comps', c);
  },

  deleteComp(req) {
    const at = find_('Comps', req.id);
    at.sheet.deleteRow(at.row);
    return { id: req.id };
  },

  saveSettings(req) {
    const sh = sheet_('Settings'), s = req.settings || {};
    Object.keys(s).forEach(key => {
      const value = JSON.stringify(s[key]);
      const hit = sh.getRange('A:A').createTextFinder(key).matchEntireCell(true).findNext();
      if (hit && hit.getRow() > 1) sh.getRange(hit.getRow(), 2).setValue(value);
      else sh.appendRow([key, value]);
    });
    return { saved: Object.keys(s) };
  },

  uploadImage(req) {
    const m = /^data:(image\/[\w.+-]+);base64,(.+)$/.exec(req.dataUrl || '');
    if (!m) throw new Error('이미지 형식이 아니에요');
    if (m[2].length > MAX_IMAGE_CHARS) throw new Error('이미지가 너무 커요');
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'jupjup-' + Date.now() + '.jpg');
    const file = imageFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { imageId: file.getId() };
  }
};

/* ---------- helpers ---------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() {
  const id = props_().getProperty('SHEET_ID');
  if (!id) throw new Error('setup()을 먼저 실행하세요');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(name + ' 시트가 없어요. setup()을 다시 실행하세요');
  return sh;
}
function newId_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function rows_(name) {
  const sh = sheet_(name), cols = SCHEMA[name], last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => toObj_(name, r));
}
function find_(name, id) {
  if (!id) throw new Error('id가 없어요');
  const sh = sheet_(name);
  const hit = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!hit || hit.getRow() < 2) throw new Error('찾을 수 없어요: ' + id);
  return { sheet: sh, row: hit.getRow() };
}
/** Keep only known, writable columns and coerce types. */
function clean_(name, src) {
  const o = {};
  SCHEMA[name].forEach(c => {
    if (LOCKED.indexOf(c) >= 0 || !(c in src)) return;
    let v = src[c];
    if (v === undefined) return;
    if (v === null || v === '') { o[c] = null; return; }
    if (NUM_COLS.indexOf(c) >= 0) v = Math.round(Number(v)) || 0;
    else if (DATE_COLS.indexOf(c) >= 0) { const d = new Date(v); if (isNaN(d)) throw new Error(c + ' 날짜 형식이 잘못됐어요'); v = d.toISOString(); }
    else v = String(v).slice(0, 2000);
    o[c] = v;
  });
  return o;
}
function toRow_(name, o) {
  return SCHEMA[name].map(c => {
    const v = o[c];
    if (v === null || v === undefined || v === '') return '';
    if (DATE_COLS.indexOf(c) >= 0) return new Date(v);
    if (NUM_COLS.indexOf(c) >= 0) return Number(v) || 0;
    const s = String(v);
    return /^[=+\-@]/.test(s) ? "'" + s : s; // never let user text become a formula
  });
}
function toObj_(name, row) {
  const o = {};
  SCHEMA[name].forEach((c, i) => {
    let v = row[i];
    if (v instanceof Date) v = v.toISOString();
    else if (v === '') v = null;
    else if (NUM_COLS.indexOf(c) >= 0) v = Number(v);
    else if (typeof v !== 'string') v = String(v);
    o[c] = v;
  });
  return o;
}
function normalize_(name, o) {
  const out = {};
  SCHEMA[name].forEach(c => { out[c] = o[c] === undefined || o[c] === '' ? null : o[c]; });
  return out;
}
function imageFolder_() {
  const p = props_(), id = p.getProperty('IMAGE_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ } }
  const f = DriveApp.createFolder('줍줍 매물 사진');
  p.setProperty('IMAGE_FOLDER_ID', f.getId());
  return f;
}
