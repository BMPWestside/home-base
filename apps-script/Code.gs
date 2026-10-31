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
