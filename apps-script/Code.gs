/**
 * Home Base — sync backend (Google Apps Script web app).
 * Stores your lists in one JSON file in your Google Drive (homebase-house.json),
 * shared by every phone that has the household key.
 *
 * 1. Put your own secret key below (any long random text).
 * 2. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * 3. Copy the /exec URL into the app (⚙︎ > Sync).
 */
const HOUSE_KEY = 'PUT-HOUSEHOLD-KEY-HERE';   // everyone in the house

function doGet(e)  { return handle_(e && e.parameter || {}); }
function doPost(e) { let p = {}; try { p = JSON.parse(e.postData.contents); } catch (err) {} return handle_(p); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function allowed_(doc, key) {
  if (!key) return false;
  return doc === 'house' && key === HOUSE_KEY;
}

function file_(doc) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('file_' + doc);
  if (id) { try { return DriveApp.getFileById(id); } catch (e) {} }
  const f = DriveApp.createFile('homebase-' + doc + '.json', '', MimeType.PLAIN_TEXT);
  props.setProperty('file_' + doc, f.getId());
  return f;
}

function read_(doc, v) {
  if (!v) return null;
  const txt = file_(doc).getBlob().getDataAsString();
  return txt ? JSON.parse(txt) : null;
}

function handle_(p) {
  const doc = p.doc;
  if (!allowed_(doc, p.key)) return out_({ ok: false, error: 'wrong key' });
  const props = PropertiesService.getScriptProperties();
  const ver = () => Number(props.getProperty('ver_' + doc) || 0);

  if (p.op === 'ver') return out_({ ok: true, version: ver() });
  if (p.op === 'bills') { try { return out_({ ok: true, data: bbData_() }); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'billAction') { try { return out_(bbAction_(p.action, JSON.parse(p.payload || '{}'))); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'get') { const v = ver(); return out_({ ok: true, version: v, data: read_(doc, v) }); }
  if (p.op === 'put') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const v = ver();
      if (Number(p.base) !== v) return out_({ ok: false, conflict: true, version: v, data: read_(doc, v) });
      file_(doc).setContent(JSON.stringify(p.data));
      props.setProperty('ver_' + doc, String(v + 1));
      return out_({ ok: true, version: v + 1 });
    } finally { lock.releaseLock(); }
  }
  return out_({ ok: false, error: 'unknown op' });
}


/* ---------- Daily gas bottle check (ntfy) ----------
 * Run setupDailyCheck() once from the editor (select it in the toolbar, press Run, allow access).
 * Every morning around 8am it checks the Gas bottle tracker and sends one ntfy alert when the
 * bottle reaches its average life. It uses the same ntfy topic as BillBell (Home Base picks it
 * up automatically once the Bills tab is connected). To use a different topic, put it below.
 */
const NTFY_TOPIC = '';
const APP_URL = 'https://bmpwestside.github.io/home-base/';

function checkGasBottle() {
  const props = PropertiesService.getScriptProperties();
  const data = read_('house', Number(props.getProperty('ver_house') || 0));
  if (!data) return;
  const topic = NTFY_TOPIC || data.ntfy;
  if (!topic) { Logger.log('No ntfy topic yet — connect the Bills tab in Home Base, or set NTFY_TOPIC.'); return; }
  try { checkDates_(data, topic); } catch (e) { Logger.log('Dates: ' + e.message); }
  const t = (data.trackers || []).find(x => /gas/i.test(x.name || '') || x.icon === '🔥');
  if (!t) return;
  const ch = (t.changes || []).slice().sort();
  if (ch.length < 2) return;                                   // need one full bottle to know the average
  const day = s => { const p = s.split('-').map(Number); return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5); };
  const today = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd');
  const gaps = ch.slice(1).map((d, i) => day(d) - day(ch[i]));
  const avg = Math.round(gaps.reduce((a, b) => a + b, 0) / gaps.length);
  const last = ch[ch.length - 1], since = day(today) - day(last);
  if (since < avg) return;
  if (props.getProperty('gas_alerted') === last) return;      // only once per bottle
  UrlFetchApp.fetch('https://ntfy.sh/' + topic, {
    method: 'post', muteHttpExceptions: true,
    payload: since + ' days since the last change. Your bottles usually last about ' + avg + ' days.',
    headers: { Title: 'Gas bottle due - Home Base', Tags: 'fire', Click: APP_URL }
  });
  props.setProperty('gas_alerted', last);
}

/** Sends a test alert so you can check ntfy is working. */
function testGasAlert() {
  const props = PropertiesService.getScriptProperties();
  const data = read_('house', Number(props.getProperty('ver_house') || 0)) || {};
  const topic = NTFY_TOPIC || data.ntfy;
  if (!topic) { Logger.log('No ntfy topic yet.'); return; }
  UrlFetchApp.fetch('https://ntfy.sh/' + topic, { method: 'post', payload: 'Gas bottle alerts are working.', headers: { Title: 'Home Base test', Tags: 'white_check_mark', Click: APP_URL } });
  Logger.log('Test sent to ' + topic);
}

function setupDailyCheck() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'checkGasBottle') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('checkGasBottle').timeBased().everyDays(1).atHour(8).inTimezone('Australia/Sydney').create();
  checkGasBottle();
}


/* ---------- Bills (reads the BillBell Data sheet in this same Google account) ---------- */
const BILLBELL_SHEET = 'BillBell Data';
const BILL_EDITABLE = ['provider', 'category', 'amount', 'dueDate', 'issueDate', 'status', 'paidDate', 'notes', 'invoiceNo', 'autopay', 'reason', 'flags'];

function bbSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('bb_sheet');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const it = DriveApp.getFilesByName(BILLBELL_SHEET);
  if (!it.hasNext()) throw new Error('Could not find the "' + BILLBELL_SHEET + '" sheet in this Google account');
  const f = it.next();
  props.setProperty('bb_sheet', f.getId());
  return SpreadsheetApp.openById(f.getId());
}
function bbTable_(name) {
  const sh = bbSheet_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues(), h = v[0];
  return v.slice(1).map(r => { const o = {}; h.forEach((k, i) => { const x = r[i]; o[k] = x instanceof Date ? Utilities.formatDate(x, 'Australia/Sydney', 'yyyy-MM-dd') : (x === null || x === undefined ? '' : x); }); return o; });
}
function bbUpdate_(name, id, patch) {
  const sh = bbSheet_().getSheetByName(name), v = sh.getDataRange().getValues(), h = v[0];
  for (let r = 1; r < v.length; r++) {
    if (String(v[r][0]) === String(id)) {
      Object.keys(patch).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(r + 1, c + 1).setNumberFormat('@').setValue(patch[k] == null ? '' : String(patch[k])); });
      return true;
    }
  }
  return false;
}
function bbAppend_(name, obj) {
  const sh = bbSheet_().getSheetByName(name), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = h.map(k => obj[k] == null ? '' : String(obj[k]));
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function bbProviders_() { return bbTable_('Providers').filter(p => String(p.active) !== 'false'); }
function bbData_() {
  const cutoff = Utilities.formatDate(new Date(Date.now() - 800 * 864e5), 'Australia/Sydney', 'yyyy-MM-dd');
  const bills = bbTable_('Bills').filter(b => b.status === 'due' || b.status === 'review' || (b.dueDate || b.paidDate || '') >= cutoff);
  return { bills: bills, providers: bbProviders_(), settings: { ntfyTopic: NTFY_TOPIC }, generated: new Date().toISOString() };
}
function bbAction_(action, body) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (action === 'update') {
      const patch = {}; BILL_EDITABLE.forEach(k => { if (body.patch && k in body.patch) patch[k] = body.patch[k]; });
      if (patch.provider && !patch.category) { const pr = bbProviders_().find(p => p.id === patch.provider); if (pr) patch.category = pr.category; }
      bbUpdate_('Bills', body.id, patch);
    } else if (action === 'add') {
      const b = body.bill || {}, pv = bbProviders_().find(p => p.id === b.provider);
      bbAppend_('Bills', Object.assign({}, b, {
        id: 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        category: pv ? pv.category : (b.category || 'Other'), status: b.status || 'due', source: 'manual', createdAt: new Date().toISOString()
      }));
    } else if (action === 'provider') {
      const row = body.provider || {};
      if (!bbUpdate_('Providers', row.id, row)) bbAppend_('Providers', Object.assign({ active: 'true', cycle: 30 }, row));
    } else if (action !== 'data') {
      return { ok: false, error: 'Use BillBell for ' + action };
    }
  } finally { lock.releaseLock(); }
  return { ok: true, data: bbData_() };
}


/* ---------- Important dates reminders (ntfy, 7 days before, the day before and on the day) ---------- */
function checkDates_(data, topic) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd');
  const day = s => { const p = s.split('-').map(Number); return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5); };
  (data.dates || []).forEach(dt => {
    if (!dt.date) return;
    const p = dt.date.split('-').map(Number), ty = Number(today.slice(0, 4));
    let iso = dt.yearly === false ? dt.date : ty + '-' + String(p[1]).padStart(2, '0') + '-' + String(p[2]).padStart(2, '0');
    if (dt.yearly !== false && iso < today) iso = (ty + 1) + iso.slice(4);
    const left = day(iso) - day(today);
    if ([7, 1, 0].indexOf(left) < 0) return;
    const key = 'date_' + dt.id + '_' + iso + '_' + left;
    if (props.getProperty(key)) return;
    const years = dt.yearly !== false && p[0] > 1900 ? Number(iso.slice(0, 4)) - p[0] : 0;
    const extra = years ? (dt.icon === '🎂' ? ' (turns ' + years + ')' : ' (' + years + ' years)') : '';
    UrlFetchApp.fetch('https://ntfy.sh/' + topic, {
      method: 'post', muteHttpExceptions: true,
      payload: (left === 0 ? 'Today' : left === 1 ? 'Tomorrow' : 'In a week') + ': ' + dt.name + extra,
      headers: { Title: (left === 0 ? 'Today' : 'Coming up') + ' - Home Base', Tags: 'tada', Click: APP_URL }
    });
    props.setProperty(key, '1');
  });
}
