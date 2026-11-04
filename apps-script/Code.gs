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

function doGet(e)  { const p = e && e.parameter || {}; return (p.op || p.doc) ? handle_(p) : setupPage_(); }

/* The household key: the one written above, or (in a fresh copy) one made automatically the first time. */
function houseKey_() {
  if (HOUSE_KEY && HOUSE_KEY.indexOf('PUT-') !== 0) return HOUSE_KEY;
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('house_key');
  if (!k) { k = 'house-' + Utilities.getUuid().replace(/-/g, '').slice(0, 24); props.setProperty('house_key', k); }
  return k;
}
/* Opening the script's web-app address in Safari shows this page with your personal Home Base link. */
function setupPage_() {
  const url = ScriptApp.getService().getUrl();
  try {   // switch on the every-minute reminder check the first time this page is opened
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'checkReminders'))
      ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(1).create();
  } catch (e) {}
  const cfg = Utilities.base64EncodeWebSafe(JSON.stringify({ u: url, k: houseKey_(), w: 'S' }), Utilities.Charset.UTF_8).replace(/=+$/, '');
  const link = APP_URL + '?c=' + encodeURIComponent(cfg);
  const html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><title>Home Base</title>'
    + '<body style="font-family:-apple-system,Helvetica,sans-serif;background:#0f2b2b;color:#fff;padding:28px;line-height:1.5">'
    + '<h1 style="color:#5fe0cf;margin:0 0 6px">Home Base is ready 🎉</h1>'
    + '<p>This is your own private copy. Your lists are saved in your Google Drive.</p>'
    + '<ol><li>Tap the button below (on your iPhone, in Safari).</li><li>Tap <b>Share</b> → <b>Add to Home Screen</b> → <b>Add</b>.</li><li>Open Home Base from your Home Screen, go to <b>Reminders → ⋯</b> and tap <b>Turn on notifications</b>.</li></ol>'
    + '<p><a href="' + link + '" target="_top" style="display:inline-block;background:#5fe0cf;color:#0f2b2b;font-weight:800;padding:14px 22px;border-radius:14px;text-decoration:none">Open my Home Base</a></p>'
    + '<p style="opacity:.7;font-size:14px">Keep this page to yourself: the button opens your lists.</p></body>';
  return HtmlService.createHtmlOutput(html).setTitle('Home Base').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function doPost(e) { let p = {}; try { p = JSON.parse(e.postData.contents); } catch (err) {} return handle_(p); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function allowed_(doc, key) {
  if (!key) return false;
  return doc === 'house' && key === houseKey_();
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
  if (p.op === 'remAct' || p.op === 'remAdd') return ContentService.createTextOutput(remRemote_(p));
  if (!allowed_(doc, p.key)) return out_({ ok: false, error: 'wrong key' });
  const props = PropertiesService.getScriptProperties();
  const ver = () => Number(props.getProperty('ver_' + doc) || 0);

  if (p.op === 'ver') return out_({ ok: true, version: ver() });
  if (p.op === 'remSetup') { try { return out_(remSetup_(p)); } catch (e) { return out_({ ok: false, error: 'setup: ' + e.message }); } }
  if (p.op === 'remToken') { try { return out_(remToken_(p)); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'pushKey') { try { return out_({ ok: true, key: wpVapid_().pub }); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'pushSub') { try { return out_(wpSubscribe_(p)); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'remStatus') { try { return out_(remStatus_(p)); } catch (e) { return out_({ ok: false, error: e.message }); } }
  if (p.op === 'remSync') { try { return out_(remSync_(p)); } catch (e) { return out_({ ok: false, error: e.message }); } }
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
  const gasMsg = since + ' days since the last change. Your bottles usually last about ' + avg + ' days.';
  if (topic) UrlFetchApp.fetch('https://ntfy.sh/' + topic, {
    method: 'post', muteHttpExceptions: true, payload: gasMsg,
    headers: ntfyAuth_({ Title: 'Gas bottle due - Home Base', Tags: 'fire', Click: APP_URL })
  });
  hbBroadcast_('🔥 Gas bottle due', gasMsg);
  props.setProperty('gas_alerted', last);
}

/** Sends a test alert so you can check ntfy is working. */
function testGasAlert() {
  const props = PropertiesService.getScriptProperties();
  const data = read_('house', Number(props.getProperty('ver_house') || 0)) || {};
  const topic = NTFY_TOPIC || data.ntfy;
  if (!topic) { Logger.log('No ntfy topic yet.'); return; }
  UrlFetchApp.fetch('https://ntfy.sh/' + topic, { method: 'post', payload: 'Gas bottle alerts are working.', headers: ntfyAuth_({ Title: 'Home Base test', Tags: 'white_check_mark', Click: APP_URL }) });
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
    if (left !== 3) return;                                      // one reminder, 3 days before
    const key = 'date_' + dt.id + '_' + iso + '_' + left;
    if (props.getProperty(key)) return;
    const years = dt.yearly !== false && p[0] > 1900 ? Number(iso.slice(0, 4)) - p[0] : 0;
    const kind = dt.type || (dt.icon === '🎂' ? 'birthday' : dt.icon === '💍' ? 'anniversary' : 'other');
    const extra = years ? (kind === 'birthday' ? ' (turns ' + years + ')' : kind === 'anniversary' ? ' (' + years + ' years)' : ' (' + years + ' years)') : '';
    if (topic) UrlFetchApp.fetch('https://ntfy.sh/' + topic, {
      method: 'post', muteHttpExceptions: true,
      payload: 'In 3 days: ' + dt.name + extra,
      headers: ntfyAuth_({ Title: 'Coming up - Home Base', Tags: 'tada', Click: APP_URL })
    });
    hbBroadcast_((dt.icon || '🎉') + ' Coming up', 'In 3 days: ' + dt.name + extra);
    props.setProperty(key, '1');
  });
}


/* ---------- Reminders (private to each person, pushed through ntfy) ----------
 * Run setupReminders() once from the editor. It checks every minute and sends due reminders
 * to that person's own ntfy topic, with Done / 1 hour / Tomorrow buttons on the alert.
 */
const REM_TZ = 'Australia/Sydney';
const remNow_ = () => Utilities.formatDate(new Date(), REM_TZ, "yyyy-MM-dd'T'HH:mm");
function remRead_(who) { const t = file_('rem_' + who).getBlob().getDataAsString(); return t ? JSON.parse(t) : null; }
function remWrite_(who, d) { file_('rem_' + who).setContent(JSON.stringify(d)); }
const remWho_ = w => (w === 'D' ? 'D' : 'S');

function remSync_(p) {
  const who = remWho_(p.who), incoming = typeof p.data === 'string' ? JSON.parse(p.data || '{}') : (p.data || {});
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const merged = RL.merge(remRead_(who) || {}, incoming);
    remWrite_(who, merged);
    return { ok: true, data: merged };
  } finally { lock.releaseLock(); }
}

/* Done / Snooze from the alert buttons, and quick add from Siri. Checked against the person's secret key. */
function remRemote_(p) {
  const who = remWho_(p.who);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = remRead_(who);
    if (!d || !d.key || p.k !== d.key) return 'Not allowed';
    const now = remNow_(), stamp = Date.now();
    if (p.op === 'remAdd') {
      const r = RL.parse(p.text || '', now);
      if (!r.title) return 'Say what to remind you about';
      const due = r.due || RL.addMin(now.slice(0, 10) + 'T09:00', now.slice(11) >= '09:00' ? 1440 : 0);
      d.items = d.items || [];
      d.items.push({ id: 'r' + stamp.toString(36) + Math.random().toString(36).slice(2, 5), title: r.title, due: due, anchor: due, rep: r.rep, early: 0, nag: false, hi: false, u: stamp, created: now });
      remWrite_(who, d);
      return 'Added: ' + r.title + ' - ' + remWhen_(due) + (RL.repeats({ rep: r.rep }) ? ' (' + RL.repText({ rep: r.rep, due: due }) + ')' : '');
    }
    if (p.a === 'rtdone') {
      const r = (d.routines || []).find(x => x.id === p.id); if (!r) return 'Routine not found';
      d.rdone = d.rdone || {}; d.rdone[r.id] = { date: now.slice(0, 10), steps: (r.steps || []).map(s => s.id) }; d.su = stamp;
      remWrite_(who, d); return 'All done: ' + r.name;
    }
    const it = (d.items || []).find(x => x.id === p.id);
    if (!it || it.done) return 'Already done';
    if (p.a === 'done') { const h = RL.complete(it, now, stamp); d.hist = [h].concat(d.hist || []).slice(0, 60); remWrite_(who, d); return 'Done: ' + it.title; }
    if (p.a === 'snooze' || p.a === 'tmr') { RL.snooze(it, now, p.a === 'tmr' ? 'tmr' : 60, stamp); remWrite_(who, d); return 'Snoozed: ' + it.title + ' until ' + remWhen_(it.snooze); }
    return 'Unknown action';
  } finally { lock.releaseLock(); }
}

function remWhen_(w) {
  const now = remNow_(), day = w.slice(0, 10), h = +w.slice(11, 13), mi = w.slice(14, 16);
  const t = ((h % 12) || 12) + (mi === '00' ? '' : ':' + mi) + (h < 12 ? 'am' : 'pm');
  if (day === now.slice(0, 10)) return 'today ' + t;
  if (day === RL.addMin(now.slice(0, 10) + 'T00:00', 1440).slice(0, 10)) return 'tomorrow ' + t;
  const D = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return D[RL.dow(w)] + ' ' + (+w.slice(8, 10)) + ' ' + M[+w.slice(5, 7) - 1] + ' ' + t;
}

function checkReminders() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('rem_lastrun', remNow_());
  const today = remNow_().slice(0, 10);
  if (remNow_().slice(11) >= '08:00' && props.getProperty('daily_done') !== today) {    // gas bottle + important dates, once a day after 8am
    props.setProperty('daily_done', today);
    try { checkGasBottle(); } catch (e) { Logger.log('Daily: ' + e.message); }
  }
  try { checkReminders__(props); props.deleteProperty('rem_lasterr'); }
  catch (e) { props.setProperty('rem_lasterr', remNow_() + ' ' + e.message); throw e; }
}
function checkReminders__(props) {
  const url = ScriptApp.getService().getUrl();
  ['S', 'D'].forEach(who => {
    let d; try { d = remRead_(who); } catch (e) { return; }
    if (!d || !d.topic || (!(d.items || []).length && !(d.routines || []).length)) return;
    const now = remNow_(), hm = now.slice(11);
    const q = d.quiet || {};
    if (q.on && q.from && q.to && (q.from < q.to ? (hm >= q.from && hm < q.to) : (hm >= q.from || hm < q.to))) return;   // quiet hours: hold until they end
    const sent = JSON.parse(props.getProperty('remsent_' + who) || '{}'), keep = {};
    d.items.forEach(it => {
      if (it.done) return;
      const st = sent[it.id] || {}, when = it.snooze || it.due;
      if (it.early > 0 && !it.snooze) {
        const ea = RL.addMin(it.due, -it.early);
        if (now >= ea && now < it.due && st.e !== it.due && remOk_(st, remPush_(d, it, 'early', url, who))) st.e = it.due;
      }
      if (now >= when && st.m !== when) { if (remOk_(st, remPush_(d, it, 'due', url, who))) { st.m = when; st.n = RL.addMin(now, 15); st.c = 0; } }
      else if (it.nag && st.m === when && st.n && now >= st.n && (st.c || 0) < 12) { remPush_(d, it, 'nag', url, who); st.n = RL.addMin(now, 15); st.c = (st.c || 0) + 1; }
      keep[it.id] = st;
    });
    const today = now.slice(0, 10);
    (d.routines || []).forEach(r => {
      if (!r.alert || !r.time || !(r.steps || []).length) return;
      if (r.days && r.days.length && r.days.indexOf(RL.dow(today + 'T00:00')) < 0) return;
      const done = (d.rdone && d.rdone[r.id] && d.rdone[r.id].date === today) ? d.rdone[r.id].steps || [] : [];
      const left = r.steps.filter(s => done.indexOf(s.id) < 0);
      const k = 'rt_' + r.id, prev = sent[k] || {};
      if (left.length && now >= today + 'T' + r.time && prev.d !== today && remOk_(prev, remPushRoutine_(d, r, left, url, who))) prev.d = today;
      keep[k] = prev;
    });
    props.setProperty('remsent_' + who, JSON.stringify(keep));
  });
}

function remPush_(d, it, kind, url, who) {
  const act = a => url + '?op=remAct&who=' + who + '&id=' + encodeURIComponent(it.id) + '&a=' + a + '&k=' + encodeURIComponent(d.key);
  const mins = it.early >= 1440 ? (it.early / 1440) + ' day' + (it.early > 1440 ? 's' : '') : it.early >= 60 ? (it.early / 60) + ' hour' + (it.early > 60 ? 's' : '') : it.early + ' min';
  const title = d.hide ? 'Home Base' : kind === 'early' ? 'Coming up in ' + mins : kind === 'nag' ? 'Still to do' : 'Reminder';
  const message = d.hide ? 'You have a reminder' : it.title + (kind === 'early' ? ' - ' + remWhen_(it.due) : '');
  const body = {
    topic: d.topic, title: title, message: message,
    priority: it.hi ? 5 : kind === 'nag' ? 4 : 3, tags: [it.hi ? 'rotating_light' : 'alarm_clock'],
    actions: kind === 'early' ? [] : [
      { action: 'http', label: 'Done', url: act('done'), method: 'GET', clear: true },
      { action: 'http', label: '1 hour', url: act('snooze'), method: 'GET', clear: true },
      { action: 'http', label: 'Tomorrow', url: act('tmr'), method: 'GET', clear: true }
    ]
  };
  body.badge = remBadge_(d); body.tag = it.id;
  body.ptitle = d.hide ? 'Reminder' : '⏰ ' + it.title;
  body.pbody = d.hide ? 'You have a reminder' : kind === 'early' ? 'Coming up · ' + remWhen_(it.due) : kind === 'nag' ? 'Still to do · tap to open' : 'Due ' + remWhen_(it.snooze || it.due);
  return remSend_(who, body, it.title, d);
}
/* Number on the Home Base icon: overdue reminders plus bills due within 3 days. */
function remBadge_(d) {
  const now = remNow_(); let n = (d.items || []).filter(x => !x.done && (x.snooze || x.due) <= now).length;
  try { n += bbUrgent_(); } catch (e) {}
  return n;
}
function bbUrgent_() {
  if (typeof BILLS_ENABLED !== 'undefined' && !BILLS_ENABLED) return 0;
  const c = CacheService.getScriptCache(), hit = c.get('bb_urgent'); if (hit !== null) return Number(hit);
  let n = 0;
  try { const lim = Utilities.formatDate(new Date(Date.now() + 3 * 864e5), REM_TZ, 'yyyy-MM-dd');
    n = bbData_().bills.filter(b => b.status === 'due' && String(b.autopay) !== 'true' && b.dueDate && b.dueDate <= lim).length; } catch (e) {}
  c.put('bb_urgent', String(n), 900); return n;
}
/* A failed alert is tried again next minute; after 30 failed tries it is skipped so it can't loop forever. */
function remOk_(st, code) {
  if (code === 200) { st.f = 0; return true; }
  st.f = (st.f || 0) + 1;
  return st.f >= 30;
}
function remSend_(who, body, label, d) {
  const via = (d && d.via) || 'ntfy', hasPush = wpSubs_(who).length > 0;
  let push = '', ntfy = '';
  if (hasPush && via !== 'ntfy') push = wpSend_(who, { title: body.ptitle || body.title, body: body.pbody || body.message, tag: body.tag, badge: body.badge, urgent: body.priority >= 4 });
  if (body.topic && (via !== 'push' || !hasPush)) ntfy = remNtfy_(body);
  const code = push === 200 || ntfy === 200 ? 200 : (push || ntfy || 'nothing to send to');
  const props = PropertiesService.getScriptProperties(), k = 'remlog_' + who;
  const log = JSON.parse(props.getProperty(k) || '[]'); log.unshift({ at: remNow_(), t: label, code: code, push: push, ntfy: ntfy });
  props.setProperty(k, JSON.stringify(log.slice(0, 6)));
  return code;
}
function remNtfy_(body) {
  let code = 0;
  // Publish to https://ntfy.sh/<topic> with headers (Google can't reach ntfy's JSON address).
  const enc = t => /^[\x20-\x7e]*$/.test(t) ? t : '=?UTF-8?B?' + Utilities.base64Encode(t, Utilities.Charset.UTF_8) + '?=';
  const headers = ntfyAuth_({ Title: enc(body.title || 'Home Base'), Priority: String(body.priority || 3), Tags: (body.tags || []).join(',') });
  if (body.actions && body.actions.length) headers.Actions = JSON.stringify(body.actions);
  const opts = { method: 'post', contentType: 'text/plain; charset=utf-8', payload: body.message || '', headers: headers, muteHttpExceptions: true };
  for (let i = 0; i < 3 && code !== 200; i++) {
    if (i) Utilities.sleep(2000 * i);
    try { code = UrlFetchApp.fetch('https://ntfy.sh/' + encodeURIComponent(body.topic), opts).getResponseCode(); }
    catch (e) { code = 'error: ' + e.message; }
  }
  return code;
}
/* Gas bottle and important-date alerts go to everyone in the house who turned notifications on. */
function hbBroadcast_(title, message) {
  ['S', 'D'].forEach(w => { try { wpSend_(w, { title: title, body: message }); } catch (e) {} });
}
function remSetup_(p) {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'checkReminders') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(1).create();
  checkReminders();
  return remStatus_(p);
}
/* Your own free ntfy.sh account token: ntfy limits anonymous senders by address, and Google's are shared. */
function ntfyAuth_(headers) {
  const t = PropertiesService.getScriptProperties().getProperty('ntfy_token');
  if (t) headers.Authorization = 'Bearer ' + t;
  return headers;
}
function remToken_(p) {
  const props = PropertiesService.getScriptProperties(), t = String(p.token || '').trim();
  if (t) props.setProperty('ntfy_token', t); else props.deleteProperty('ntfy_token');
  return remStatus_(p);
}
function remStatus_(p) {
  const who = remWho_(p.who), props = PropertiesService.getScriptProperties(), d = remRead_(who) || {};
  const open = (d.items || []).filter(x => !x.done).sort((a, b) => ((a.snooze || a.due) < (b.snooze || b.due) ? -1 : 1));
  let res = { ok: true, version: 3, now: remNow_(), triggers: ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkReminders').length,
    hasToken: !!props.getProperty('ntfy_token'), pushSubs: wpSubs_(who).length, via: d.via || 'ntfy', engine: wpSelfTest_(), lastRun: props.getProperty('rem_lastrun') || '', lastErr: props.getProperty('rem_lasterr') || '', topic: d.topic || '', open: open.length,
    next: open[0] ? { t: open[0].title, at: open[0].snooze || open[0].due } : null, log: JSON.parse(props.getProperty('remlog_' + who) || '[]') };
  if (p.test) { res.test = remSend_(who, { topic: d.topic, title: 'Home Base', message: 'Test from your Google script ✅', tags: ['white_check_mark'], badge: remBadge_(d) }, 'Test', d); res.testLog = JSON.parse(props.getProperty('remlog_' + who) || '[]')[0]; }
  return res;
}

function remPushRoutine_(d, r, left, url, who) {
  const all = url + '?op=remAct&who=' + who + '&id=' + encodeURIComponent(r.id) + '&a=rtdone&k=' + encodeURIComponent(d.key);
  const body = {
    topic: d.topic, title: d.hide ? 'Home Base' : (r.icon || '') + ' ' + r.name,
    message: d.hide ? 'Time for your routine' : left.length + ' of ' + r.steps.length + ' to do: ' + left.slice(0, 4).map(s => s.text).join(', ') + (left.length > 4 ? '…' : ''),
    priority: 3, tags: ['sunrise'],
    actions: [{ action: 'http', label: 'All done', url: all, method: 'GET', clear: true }]
  };
  body.badge = remBadge_(d); body.tag = 'rt_' + r.id;
  return remSend_(who, body, r.name, d);
}

function wpSelfTest_() {
  try { const v = WP.makeVapid(); return v.pub.length > 80 ? 'ok' : 'bad'; } catch (e) { return 'error: ' + e.message; }
}

function setupReminders() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'checkReminders') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(1).create();
  checkReminders();
  Logger.log('Reminders are on. They are checked every minute.');
}


/* ---------- Home Base notifications (Web Push, the same standard SMail uses) ----------
 * Google Apps Script has no built-in tools for Apple's push encryption, so this file carries its own:
 * P-256 elliptic curves (signing and key exchange), SHA-256/HMAC and AES-128-GCM, all in plain JavaScript.
 */
var WP = (function(){
  /* ---- bytes ---- */
  const u8 = a => Uint8Array.from(a, x => x & 255);
  const cat = (...a) => { const n = a.reduce((s, x) => s + x.length, 0), o = new Uint8Array(n); let i = 0; a.forEach(x => { o.set(x, i); i += x.length; }); return o; };
  const utf8 = s => { const o = []; s = unescape(encodeURIComponent(s)); for (let i = 0; i < s.length; i++) o.push(s.charCodeAt(i)); return u8(o); };
  const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
  function b64u(b){ let s = "", i = 0; for (; i + 2 < b.length; i += 3){ const n = (b[i] << 16) | (b[i+1] << 8) | b[i+2]; s += B64[n >> 18] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63] + B64[n & 63]; }
    if (b.length - i === 1){ const n = b[i] << 16; s += B64[n >> 18] + B64[(n >> 12) & 63]; } else if (b.length - i === 2){ const n = (b[i] << 16) | (b[i+1] << 8); s += B64[n >> 18] + B64[(n >> 12) & 63] + B64[(n >> 6) & 63]; } return s; }
  function unb64u(s){ s = String(s).replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_"); const o = []; let buf = 0, bits = 0;
    for (const c of s){ const v = B64.indexOf(c); if (v < 0) continue; buf = (buf << 6) | v; bits += 6; if (bits >= 8){ bits -= 8; o.push((buf >> bits) & 255); } } return u8(o); }
  /* ---- 256-bit numbers as 16 little-endian 16-bit limbs (no BigInt: Google's script engine doesn't support it) ---- */
  const L = 16;
  const zero = () => new Array(L).fill(0);
  function limbs(b){ const x = zero(); for (let i = 0; i < L; i++) x[i] = b[31 - 2*i] | (b[30 - 2*i] << 8); return x; }
  function bytes(x){ const o = new Uint8Array(32); for (let i = 0; i < L; i++){ o[31 - 2*i] = x[i] & 255; o[30 - 2*i] = x[i] >>> 8; } return o; }
  function hexL(h){ const b = new Uint8Array(32); for (let i = 0; i < 32; i++) b[i] = parseInt(h.substr(2*i, 2), 16); return limbs(b); }
  const isZero = x => x.every(v => v === 0);
  function ge(a, b){ for (let i = L - 1; i >= 0; i--){ if (a[i] !== b[i]) return a[i] > b[i]; } return true; }
  function subRaw(a, b){ const o = zero(); let br = 0; for (let i = 0; i < L; i++){ let v = a[i] - b[i] - br; br = v < 0 ? 1 : 0; o[i] = v & 0xffff; } return o; }
  function addRaw(a, b){ const o = zero(); let c = 0; for (let i = 0; i < L; i++){ const v = a[i] + b[i] + c; o[i] = v & 0xffff; c = v >>> 16; } return [o, c]; }
  function ctx(hex){
    const m = hexL(hex), m0 = m[0];
    let inv = 1; for (let k = 0; k < 5; k++) inv = (inv * ((2 - ((m0 * inv) % 65536) + 65536) % 65536)) % 65536;
    const mp = (65536 - inv) % 65536;
    const C = { m, mp };
    C.add = (a, b) => { const [o, c] = addRaw(a, b); return (c || ge(o, m)) ? subRaw(o, m) : o; };
    C.sub = (a, b) => ge(a, b) ? subRaw(a, b) : addRaw(subRaw(a, b), m)[0];
    C.red = a => ge(a, m) ? subRaw(a, m) : a;
    C.mm = (a, b) => {            // Montgomery product a*b/2^256 mod m
      const t = new Array(L + 2).fill(0);
      for (let i = 0; i < L; i++){
        let c = 0; const bi = b[i];
        for (let j = 0; j < L; j++){ const v = t[j] + a[j] * bi + c; t[j] = v % 65536; c = Math.floor(v / 65536); }
        let v = t[L] + c; t[L] = v % 65536; t[L+1] = Math.floor(v / 65536);
        const u = (t[0] * mp) % 65536;
        v = t[0] + u * m[0]; c = Math.floor(v / 65536);
        for (let j = 1; j < L; j++){ v = t[j] + u * m[j] + c; t[j-1] = v % 65536; c = Math.floor(v / 65536); }
        v = t[L] + c; t[L-1] = v % 65536; c = Math.floor(v / 65536);
        t[L] = t[L+1] + c; t[L+1] = 0;
      }
      const o = t.slice(0, L); return (t[L] || ge(o, m)) ? subRaw(o, m) : o;
    };
    let r2 = zero(); r2[0] = 1; for (let k = 0; k < 512; k++) r2 = C.add(r2, r2);
    const one = zero(); one[0] = 1;
    C.to = a => C.mm(a, r2); C.from = a => C.mm(a, one); C.oneM = C.to(one);
    C.mul = (a, b) => C.mm(C.mm(a, b), r2);                  // plain a*b mod m
    C.pow = (aM, e) => { let r = C.oneM; for (let i = L*16 - 1; i >= 0; i--){ r = C.mm(r, r); if ((e[i >> 4] >>> (i & 15)) & 1) r = C.mm(r, aM); } return r; };
    C.inv = a => C.from(C.pow(C.to(a), subRaw(m, (() => { const t = zero(); t[0] = 2; return t; })())));
    return C;
  }

  /* ---- SHA-256 / HMAC ---- */
  const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  function sha256(msg){
    const l = msg.length, nb = ((l + 9 + 63) >> 6) << 6, m = new Uint8Array(nb); m.set(msg); m[l] = 0x80;
    const bl = l * 8; for (let i = 0; i < 8; i++) m[nb - 1 - i] = Math.floor(bl / Math.pow(2, 8 * i)) & 255;
    const H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19], W = new Array(64);
    const r = (x, n) => (x >>> n) | (x << (32 - n));
    for (let o = 0; o < nb; o += 64){
      for (let t = 0; t < 16; t++) W[t] = (m[o+4*t] << 24) | (m[o+4*t+1] << 16) | (m[o+4*t+2] << 8) | m[o+4*t+3];
      for (let t = 16; t < 64; t++){ const s0 = r(W[t-15],7) ^ r(W[t-15],18) ^ (W[t-15] >>> 3), s1 = r(W[t-2],17) ^ r(W[t-2],19) ^ (W[t-2] >>> 10); W[t] = (W[t-16] + s0 + W[t-7] + s1) | 0; }
      let [a,b,c,d,e,f,g,h] = H;
      for (let t = 0; t < 64; t++){
        const S1 = r(e,6) ^ r(e,11) ^ r(e,25), ch = (e & f) ^ (~e & g), t1 = (h + S1 + ch + K[t] + W[t]) | 0;
        const S0 = r(a,2) ^ r(a,13) ^ r(a,22), mj = (a & b) ^ (a & c) ^ (b & c), t2 = (S0 + mj) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0]+a)|0; H[1] = (H[1]+b)|0; H[2] = (H[2]+c)|0; H[3] = (H[3]+d)|0; H[4] = (H[4]+e)|0; H[5] = (H[5]+f)|0; H[6] = (H[6]+g)|0; H[7] = (H[7]+h)|0;
    }
    const out = new Uint8Array(32); H.forEach((v, i) => { out[4*i] = v >>> 24; out[4*i+1] = (v >>> 16) & 255; out[4*i+2] = (v >>> 8) & 255; out[4*i+3] = v & 255; }); return out;
  }
  function hmac(key, msg){
    if (key.length > 64) key = sha256(key);
    const k = new Uint8Array(64); k.set(key);
    return sha256(cat(k.map(x => x ^ 0x5c), sha256(cat(k.map(x => x ^ 0x36), msg))));
  }

  /* ---- randomness (Google's UUIDs come from a secure generator) ---- */
  let ctr = 0;
  function rand(n){
    const parts = []; let out = new Uint8Array(0);
    while (out.length < n){
      const seed = (typeof Utilities !== "undefined" ? Utilities.getUuid() + Utilities.getUuid() : WP_TEST_RANDOM()) + "|" + Date.now() + "|" + (++ctr) + "|" + Math.random();
      out = cat(out, sha256(utf8(seed)));
    }
    return out.slice(0, n);
  }

  /* ---- P-256 ---- */
  const FP = ctx("ffffffff00000001000000000000000000000000ffffffffffffffffffffffff");
  const FN = ctx("ffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551");
  const Bm = FP.to(hexL("5ac635d8aa3a93e7b3ebbd55769886bc651d06b0cc53b0f63bce3c3e27d2604b"));
  const G = [hexL("6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296"), hexL("4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5")];
  const halfN = hexL("7fffffff800000007fffffffffffffffde737d56d38bcf4279dce5617e3192a8");
  const fm = FP.mm, fa = FP.add, fs = FP.sub;
  // Jacobian points [X, Y, Z] in Montgomery form; Z = 0 is the point at infinity
  function dbl(p){ const [X, Y, Z] = p; if (isZero(Z) || isZero(Y)) return [zero(), FP.oneM, zero()];
    const YY = fm(Y, Y), XYY = fm(X, YY), S = fa(fa(XYY, XYY), fa(XYY, XYY)), ZZ = fm(Z, Z);
    const t = fm(fs(X, ZZ), fa(X, ZZ)), M = fa(fa(t, t), t);
    const X3 = fs(fm(M, M), fa(S, S)), YYYY = fm(YY, YY), e2 = fa(YYYY, YYYY), e4 = fa(e2, e2), e8 = fa(e4, e4);
    const YZ = fm(Y, Z); return [X3, fs(fm(M, fs(S, X3)), e8), fa(YZ, YZ)]; }
  function add(p, q){ if (isZero(p[2])) return q; if (isZero(q[2])) return p;
    const Z1Z1 = fm(p[2], p[2]), Z2Z2 = fm(q[2], q[2]), U1 = fm(p[0], Z2Z2), U2 = fm(q[0], Z1Z1);
    const S1 = fm(fm(p[1], q[2]), Z2Z2), S2 = fm(fm(q[1], p[2]), Z1Z1);
    const H = fs(U2, U1), R = fs(S2, S1);
    if (isZero(H)) return isZero(R) ? dbl(p) : [zero(), FP.oneM, zero()];
    const HH = fm(H, H), HHH = fm(H, HH), V = fm(U1, HH);
    const X3 = fs(fs(fm(R, R), HHH), fa(V, V)); return [X3, fs(fm(R, fs(V, X3)), fm(S1, HHH)), fm(fm(p[2], q[2]), H)]; }
  function mul(k, pt){
    let R = [zero(), FP.oneM, zero()]; const Q = [FP.to(pt[0]), FP.to(pt[1]), FP.oneM];
    for (let i = L*16 - 1; i >= 0; i--){ R = dbl(R); if ((k[i >> 4] >>> (i & 15)) & 1) R = add(R, Q); }
    if (isZero(R[2])) throw new Error("point at infinity");
    const zi = FP.pow(R[2], subRaw(FP.m, (() => { const t = zero(); t[0] = 2; return t; })())), zi2 = fm(zi, zi);
    return [FP.from(fm(R[0], zi2)), FP.from(fm(fm(R[1], zi2), zi))]; }
  const pub = pt => cat([4], bytes(pt[0]), bytes(pt[1]));
  function unpub(b){ if (b.length !== 65 || b[0] !== 4) throw new Error("bad public key");
    const x = limbs(b.subarray(1, 33)), y = limbs(b.subarray(33)), xm = FP.to(x), ym = FP.to(y);
    const rhs = fa(fs(fm(fm(xm, xm), xm), fa(fa(xm, xm), xm)), Bm);
    if (!isZero(fs(fm(ym, ym), rhs))) throw new Error("public key not on curve"); return [x, y]; }
  function scalar(){ for (;;){ const d = limbs(rand(32)); if (!isZero(d) && !ge(d, FN.m)) return d; } }
  function newKey(){ const d = scalar(); return { d, pub: pub(mul(d, G)) }; }

  /* ECDSA P-256 / SHA-256 with deterministic nonces (RFC 6979), so the private key never depends on luck */
  function sign(d, msg){
    const z = FN.red(limbs(sha256(msg))), x = bytes(d), h1 = bytes(z);
    let V = new Uint8Array(32).fill(1), Kk = new Uint8Array(32);
    Kk = hmac(Kk, cat(V, [0], x, h1)); V = hmac(Kk, V); Kk = hmac(Kk, cat(V, [1], x, h1)); V = hmac(Kk, V);
    for (;;){
      V = hmac(Kk, V); const k = limbs(V);
      if (!isZero(k) && !ge(k, FN.m)){
        const r = FN.red(mul(k, G)[0]);
        if (!isZero(r)){ let s = FN.mul(FN.inv(k), FN.add(z, FN.mul(r, d)));
          if (!isZero(s)){ if (!ge(halfN, s)) s = subRaw(FN.m, s); return cat(bytes(r), bytes(s)); } }
      }
      Kk = hmac(Kk, cat(V, [0])); V = hmac(Kk, V);
    }
  }

  /* ---- AES-128-GCM ---- */
  const SB = (() => { const s = new Uint8Array(256); let p = 1, q = 1;
    do { p = p ^ ((p << 1) & 255) ^ (p & 0x80 ? 0x1b : 0); q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 255; if (q & 0x80) q ^= 0x09;
      const x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4)); s[p] = (x ^ 0x63) & 255; } while (p !== 1);
    s[0] = 0x63; return s; })();
  const xt = b => ((b << 1) ^ (b & 0x80 ? 0x1b : 0)) & 255;
  function expand(key){ const w = new Uint8Array(176); w.set(key); let rc = 1;
    for (let i = 16; i < 176; i += 4){ let t = w.slice(i - 4, i);
      if (i % 16 === 0){ t = Uint8Array.of(SB[t[1]] ^ rc, SB[t[2]], SB[t[3]], SB[t[0]]); rc = xt(rc); }
      for (let j = 0; j < 4; j++) w[i + j] = w[i - 16 + j] ^ t[j]; } return w; }
  function encBlock(w, inp){ const s = inp.slice(); for (let i = 0; i < 16; i++) s[i] ^= w[i];
    for (let r = 1; r <= 10; r++){
      for (let i = 0; i < 16; i++) s[i] = SB[s[i]];
      const t = s.slice(); for (let c = 0; c < 4; c++) for (let rr = 0; rr < 4; rr++) s[c*4 + rr] = t[((c + rr) % 4)*4 + rr];
      if (r < 10) for (let c = 0; c < 4; c++){ const a0 = s[c*4], a1 = s[c*4+1], a2 = s[c*4+2], a3 = s[c*4+3], all = a0 ^ a1 ^ a2 ^ a3;
        s[c*4] ^= all ^ xt(a0 ^ a1); s[c*4+1] ^= all ^ xt(a1 ^ a2); s[c*4+2] ^= all ^ xt(a2 ^ a3); s[c*4+3] ^= all ^ xt(a3 ^ a0); }
      for (let i = 0; i < 16; i++) s[i] ^= w[r*16 + i];
    } return s; }
  const w4 = b => [0, 1, 2, 3].map(i => ((b[4*i] << 24) | (b[4*i+1] << 16) | (b[4*i+2] << 8) | b[4*i+3]) >>> 0);
  const b4 = w => { const o = new Uint8Array(16); w.forEach((v, i) => { o[4*i] = v >>> 24; o[4*i+1] = (v >>> 16) & 255; o[4*i+2] = (v >>> 8) & 255; o[4*i+3] = v & 255; }); return o; };
  function gmul(X, Y){ let Z = [0, 0, 0, 0], V = Y.slice();
    for (let i = 0; i < 128; i++){
      if ((X[i >> 5] >>> (31 - (i & 31))) & 1) Z = Z.map((z, k) => (z ^ V[k]) >>> 0);
      const lsb = V[3] & 1;
      V = [V[0] >>> 1, ((V[1] >>> 1) | ((V[0] & 1) << 31)) >>> 0, ((V[2] >>> 1) | ((V[1] & 1) << 31)) >>> 0, ((V[3] >>> 1) | ((V[2] & 1) << 31)) >>> 0];
      if (lsb) V[0] = (V[0] ^ 0xe1000000) >>> 0;
    } return Z; }
  function gcm(key, iv, pt){
    const w = expand(key), Hk = w4(encBlock(w, new Uint8Array(16)));
    const J0 = cat(iv, [0, 0, 0, 1]), ct = new Uint8Array(pt.length), cb = J0.slice();
    for (let i = 0; i < pt.length; i += 16){
      for (let k = 15; k >= 12; k--){ cb[k] = (cb[k] + 1) & 255; if (cb[k]) break; }
      const ks = encBlock(w, cb); for (let j = 0; j < 16 && i + j < pt.length; j++) ct[i + j] = pt[i + j] ^ ks[j]; }
    let y = [0, 0, 0, 0];
    for (let i = 0; i < ct.length; i += 16){ const blk = new Uint8Array(16); blk.set(ct.subarray(i, i + 16)); y = gmul(w4(blk).map((v, k) => (v ^ y[k]) >>> 0), Hk); }
    const bits = ct.length * 8; y = gmul([y[0], y[1], (y[2] ^ Math.floor(bits / 4294967296)) >>> 0, (y[3] ^ (bits >>> 0)) >>> 0], Hk);
    const ek = w4(encBlock(w, J0));
    return cat(ct, b4(y.map((v, k) => (v ^ ek[k]) >>> 0)));
  }

  /* ---- Web Push message (RFC 8291 aes128gcm + RFC 8292 VAPID) ---- */
  function encrypt(sub, payload){
    const ua = unb64u(sub.keys.p256dh), auth = unb64u(sub.keys.auth), uaPt = unpub(ua);
    const eph = newKey(), shared = bytes(mul(eph.d, uaPt)[0]), salt = rand(16);
    const prkKey = hmac(auth, shared), ikm = hmac(prkKey, cat(utf8("WebPush: info\0"), ua, eph.pub, [1]));
    const prk = hmac(salt, ikm);
    const cek = hmac(prk, cat(utf8("Content-Encoding: aes128gcm\0"), [1])).slice(0, 16);
    const nonce = hmac(prk, cat(utf8("Content-Encoding: nonce\0"), [1])).slice(0, 12);
    const body = gcm(cek, nonce, cat(utf8(payload), [2]));
    return cat(salt, [0, 0, 16, 0], [65], eph.pub, body);
  }
  function jwt(vapid, aud, subject){
    const enc = o => b64u(utf8(JSON.stringify(o)));
    const head = enc({ typ:"JWT", alg:"ES256" }) + "." + enc({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: subject });
    return head + "." + b64u(sign(limbs(unb64u(vapid.d)), utf8(head)));
  }
  function makeVapid(){ const k = newKey(); return { d: b64u(bytes(k.d)), pub: b64u(k.pub) }; }
  return { encrypt, jwt, makeVapid, b64u, unb64u, sha256, hmac, sign, gcm, mul, G, limbs, bytes, utf8 };
})();

/* ---------- sending ---------- */
function wpVapid_() {
  const props = PropertiesService.getScriptProperties();
  let v = props.getProperty('vapid');
  if (!v) { v = JSON.stringify(WP.makeVapid()); props.setProperty('vapid', v); }
  return JSON.parse(v);
}
function wpSubs_(who) { return JSON.parse(PropertiesService.getScriptProperties().getProperty('pushsubs_' + who) || '[]'); }
function wpSaveSubs_(who, subs) { PropertiesService.getScriptProperties().setProperty('pushsubs_' + who, JSON.stringify(subs.slice(-6))); }
function wpSubscribe_(p) {
  const who = remWho_(p.who), sub = typeof p.sub === 'string' ? JSON.parse(p.sub) : p.sub;
  if (!sub || !/^https:\/\//.test(sub.endpoint || '') || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) return { ok: false, error: 'bad subscription' };
  const subs = wpSubs_(who).filter(s => s.endpoint !== sub.endpoint);
  if (p.off) { wpSaveSubs_(who, subs); return { ok: true, subs: subs.length }; }
  subs.push({ endpoint: sub.endpoint, keys: { p256dh: sub.keys.p256dh, auth: sub.keys.auth }, at: remNow_() });
  wpSaveSubs_(who, subs);
  return { ok: true, subs: subs.length };
}
/* Sends one notification to every phone of this person that turned Home Base notifications on. Returns 200 if any phone got it. */
function wpSend_(who, msg) {
  const subs = wpSubs_(who); if (!subs.length) return 0;
  const vapid = wpVapid_(), payload = JSON.stringify(msg); let best = 0, keep = [];
  subs.forEach(s => {
    let code = 0;
    try {
      const aud = s.endpoint.match(/^https:\/\/[^/]+/)[0];
      const res = UrlFetchApp.fetch(s.endpoint, {
        method: 'post', muteHttpExceptions: true, contentType: 'application/octet-stream',
        payload: Array.from(WP.encrypt(s, payload), b => (b > 127 ? b - 256 : b)),
        headers: { TTL: '86400', Urgency: msg.urgent ? 'high' : 'normal', 'Content-Encoding': 'aes128gcm', Authorization: 'vapid t=' + WP.jwt(vapid, aud, APP_URL) + ', k=' + vapid.pub }
      });
      code = res.getResponseCode();
    } catch (e) { code = 'error: ' + e.message; }
    if (code !== 404 && code !== 410) keep.push(s);            // that phone removed Home Base or turned notifications off
    if (code === 201 || code === 200) best = 200; else if (!best) best = code;
  });
  if (keep.length !== subs.length) wpSaveSubs_(who, keep);
  return best;
}

/* ================= Reminders core (shared by the app and the Google script) =================
 * Times are Sydney wall-clock strings "YYYY-MM-DDTHH:MM", so daylight saving never shifts them
 * and plain string comparison works for "is it due yet".
 */
var RL = (function(){
  const pad = n => String(n).padStart(2, "0");
  const MONS = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  const DAYS = ["sun","mon","tue","wed","thu","fri","sat"];
  const toMs = w => Date.UTC(+w.slice(0,4), +w.slice(5,7) - 1, +w.slice(8,10), +w.slice(11,13) || 0, +w.slice(14,16) || 0);
  const fromMs = ms => { const d = new Date(ms); return d.getUTCFullYear() + "-" + pad(d.getUTCMonth()+1) + "-" + pad(d.getUTCDate()) + "T" + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes()); };
  const addMin = (w, m) => fromMs(toMs(w) + m * 60000);
  const dow = w => new Date(toMs(w)).getUTCDay();
  const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();          // m is 1-12
  const minsBetween = (a, b) => Math.round((toMs(b) - toMs(a)) / 60000);
  function wallNow(d){ d = d || new Date(); return d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate()) + "T" + pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function addMonths(w, k, day){
    let y = +w.slice(0,4), m = +w.slice(5,7) + k;
    while (m > 12){ m -= 12; y++; } while (m < 1){ m += 12; y--; }
    return y + "-" + pad(m) + "-" + pad(Math.min(day, daysIn(y, m))) + w.slice(10);
  }
  function step(r, w, anchorDay){
    switch (r.f){
      case "daily": return addMin(w, 1440);
      case "ndays": return addMin(w, 1440 * Math.max(1, +r.n || 1));
      case "weekdays": { let x = addMin(w, 1440); while (dow(x) === 0 || dow(x) === 6) x = addMin(x, 1440); return x; }
      case "weekly": {
        const days = (r.days && r.days.length) ? r.days : [dow(w)];
        let x = w; for (let i = 0; i < 8; i++){ x = addMin(x, 1440); if (days.indexOf(dow(x)) >= 0) return x; }
        return addMin(w, 7 * 1440);
      }
      case "monthly": return addMonths(w, Math.max(1, +r.n || 1), anchorDay);
      case "yearly": return addMonths(w, 12, anchorDay);
    }
    return null;
  }
  const repeats = it => !!(it.rep && it.rep.f && it.rep.f !== "none");
  const at = it => it.snooze || it.due;
  /* Next time a repeating reminder is due after `now` (null when the series has ended). */
  function nextDue(it, now){
    if (!repeats(it)) return null;
    const r = it.rep, anchor = +(it.anchor || it.due).slice(8,10);
    let x = it.due, guard = 0;
    do { x = step(r, x, anchor); } while (x && x <= now && ++guard < 5000);
    if (!x) return null;
    if (r.end && x.slice(0,10) > r.end) return null;
    if (r.count){ const left = (r.left == null ? r.count : r.left) - 1; if (left <= 0) return null; }
    return x;
  }
  /* Tick a reminder off. Repeating ones move to their next time; others are marked done. */
  function complete(it, now, stamp){
    const nx = nextDue(it, now), title = it.title;
    if (nx){ if (it.rep.count) it.rep.left = (it.rep.left == null ? it.rep.count : it.rep.left) - 1; it.due = nx; }
    else { it.done = true; it.doneAt = now; }
    it.snooze = null; it.u = stamp || Date.now();
    return { t: title, at: now, id: it.id + "@" + now, ms: stamp || Date.now() };
  }
  function snooze(it, now, mins, stamp){ it.snooze = mins === "tmr" ? addMin(now.slice(0,10) + "T09:00", 1440) : addMin(now, mins); it.u = stamp || Date.now(); }

  /* Merge two copies (phone + server): newest edit wins per reminder; deletes are kept as tombstones. */
  function merge(a, b){
    a = a || {}; b = b || {};
    const out = {}, del = Object.assign({}, a.del || {});
    Object.keys(b.del || {}).forEach(k => { del[k] = Math.max(del[k] || 0, b.del[k]); });
    const byId = {};
    (a.items || []).concat(b.items || []).forEach(it => { const o = byId[it.id]; if (!o || (it.u || 0) > (o.u || 0)) byId[it.id] = it; });
    out.items = Object.keys(byId).map(k => byId[k]).filter(it => !(del[it.id] && del[it.id] >= (it.u || 0)));
    const cut = Date.now() - 90 * 864e5; Object.keys(del).forEach(k => { if (del[k] < cut) delete del[k]; });
    out.del = del;
    const sa = a.su || 0, sb = b.su || 0, s = sb > sa ? b : a;
    ["topic","key","quiet","hide","lists","routines","rdone","via","su"].forEach(k => { if (s[k] !== undefined) out[k] = s[k]; else if ((s === a ? b : a)[k] !== undefined) out[k] = (s === a ? b : a)[k]; });
    const h = {}; (a.hist || []).concat(b.hist || []).forEach(x => { h[x.id] = x; });
    const hc = Math.max(a.hc || 0, b.hc || 0);
    out.hist = Object.keys(h).map(k => h[k]).filter(x => (x.ms || 0) > hc).sort((x, y) => x.at < y.at ? 1 : -1).slice(0, 60);
    out.hc = hc;
    return out;
  }

  /* ---------- typing a reminder in plain words: "Pay rego Fri 5pm", "Bins every Tue 7pm" ---------- */
  function parse(text, now){
    let s = " " + String(text || "").replace(/\s+/g, " ").trim() + " ";
    const low = () => s.toLowerCase();
    const take = re => { const m = low().match(re); if (!m) return null; s = s.slice(0, m.index) + " " + s.slice(m.index + m[0].length); return m; };
    take(/^ (remind me (to|about)?|reminder:?) /);
    let date = null, time = null, rep = null, rel = null, nth = false;
    const dayIdx = w => DAYS.indexOf(w.slice(0,3));
    const DAYRE = "(sun|mon|tue|tues|wed|thu|thur|thurs|fri|sat)(day|nesday|sday|rsday|urday)?";
    let m;
    // repeats
    if ((m = take(/ every ?day | daily /)) ) rep = { f:"daily" };
    else if ((m = take(/ every weekday | weekdays /))) rep = { f:"weekdays" };
    else if ((m = take(/ every (\d+) days? /))) rep = { f:"ndays", n:+m[1] };
    else if ((m = take(/ every (\d+) weeks? /))) rep = { f:"ndays", n:+m[1] * 7 };
    else if ((m = take(/ every (\d+) months? /))) rep = { f:"monthly", n:+m[1] };
    else if ((m = take(new RegExp(" every (" + DAYRE + "(?:(?:,| and|,? &) ?" + DAYRE + ")*) ")))){
      rep = { f:"weekly", days: (m[1].match(new RegExp(DAYRE, "g")) || []).map(dayIdx).filter(x => x >= 0) };
    }
    else if ((m = take(/ every week | weekly /))) rep = { f:"weekly" };
    else if ((m = take(/ every month | monthly /))) rep = { f:"monthly", n:1 };
    else if ((m = take(/ every year | yearly | annually /))) rep = { f:"yearly" };
    // relative
    if ((m = take(/ in (\d+|an?|half an?) (min|mins|minutes?|hours?|hrs?|days?|weeks?) /))){
      const n = /^half/.test(m[1]) ? 0.5 : /^an?$/.test(m[1]) ? 1 : +m[1], u = m[2][0];
      rel = Math.round(n * (u === "m" ? 1 : u === "h" ? 60 : u === "d" ? 1440 : 10080));
    }
    // time
    if ((m = take(/ (?:at |@ ?|by )?(\d{1,2})(?:[:.](\d{2}))? ?(am|pm|a\.m\.|p\.m\.) /))){
      let h = +m[1] % 12; if (m[3][0] === "p") h += 12; time = pad(h) + ":" + (m[2] || "00");
    } else if ((m = take(/ (?:at |@ ?)?([01]?\d|2[0-3])[:.](\d{2}) /))) time = pad(+m[1]) + ":" + m[2];
    else if ((m = take(/ (?:at )?(noon|midday) /))) time = "12:00";
    else if ((m = take(/ at (\d{1,2}) /))){ let h = +m[1]; if (h >= 1 && h <= 7) h += 12; time = pad(h) + ":00"; }
    if ((m = take(/ (this |in the )?(morning|afternoon|evening|tonight) /))){
      const w = m[2]; if (!time) time = w === "morning" ? "09:00" : w === "afternoon" ? "15:00" : w === "evening" ? "18:00" : "20:00";
      if (w === "tonight") date = now.slice(0,10);
    }
    const today = now.slice(0,10);
    // dates
    if ((m = take(/ (today|tonight) /))) date = today;
    else if ((m = take(/ (tomorrow|tmrw|tmr) /))) date = addMin(today + "T00:00", 1440).slice(0,10);
    else if ((m = take(/ (?:on )?(\d{1,2})(?:st|nd|rd|th)?(?: of)? (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*(?: (\d{4}))? /))) date = mkDate(+m[1], MONS.indexOf(m[2]) + 1, m[3]);
    else if ((m = take(/ (?:on )?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]* (\d{1,2})(?:st|nd|rd|th)?(?:,? (\d{4}))? /))) date = mkDate(+m[2], MONS.indexOf(m[1]) + 1, m[3]);
    else if ((m = take(/ (?:on )?(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))? /))) date = mkDate(+m[1], +m[2], m[3]);
    else if ((m = take(new RegExp(" (?:on |this |next )?" + DAYRE + " ")))){
      const want = dayIdx(m[1]); let x = today + "T00:00";
      for (let i = 0; i < 8; i++){ x = addMin(x, 1440); if (dow(x) === want) break; }
      if (dow(today + "T00:00") === want && !/next/.test(m[0])) x = today + "T00:00";
      date = x.slice(0,10);
    }
    else if ((m = take(/ (?:on )?the (\d{1,2})(?:st|nd|rd|th) /))){
      const d = +m[1]; let y = +today.slice(0,4), mo = +today.slice(5,7);
      let cand = y + "-" + pad(mo) + "-" + pad(Math.min(d, daysIn(y, mo)));
      if (cand < today){ mo++; if (mo > 12){ mo = 1; y++; } cand = y + "-" + pad(mo) + "-" + pad(Math.min(d, daysIn(y, mo))); }
      date = cand; nth = true;
    }
    function mkDate(d, mo, y){
      if (!(mo >= 1 && mo <= 12 && d >= 1 && d <= 31)) return null;
      let yy = y ? (+y < 100 ? 2000 + +y : +y) : +today.slice(0,4);
      let c = yy + "-" + pad(mo) + "-" + pad(Math.min(d, daysIn(yy, mo)));
      if (!y && c < today){ yy++; c = yy + "-" + pad(mo) + "-" + pad(Math.min(d, daysIn(yy, mo))); }
      return c;
    }
    let due = null;
    if (rel != null) due = addMin(now, rel);
    else if (rep && rep.f === "weekly" && rep.days && rep.days.length && !date){
      const t = time || "09:00"; let x = today + "T" + t;
      for (let i = 0; i < 8; i++){ if (rep.days.indexOf(dow(x)) >= 0 && x > now) break; x = addMin(x, 1440); }
      due = x;
    }
    else if (date || time){
      due = (date || today) + "T" + (time || "09:00");
      if (!date && due <= now) due = addMin(due, 1440);
      if (rep && !date && rep.f === "weekdays") while (dow(due) === 0 || dow(due) === 6) due = addMin(due, 1440);
    }
    else if (rep){ due = today + "T09:00"; if (due <= now) due = addMin(due, 1440); }
    if (nth && due && due <= now) due = addMonths(due, 1, +due.slice(8,10));
    let title = s.replace(/\s+/g, " ").trim().replace(/^(to|about|on|at|by)\s+/i, "").replace(/\s+(on|at|by|in|every|the)$/i, "").trim();
    if (title) title = title[0].toUpperCase() + title.slice(1);
    return { title, due, rep: rep || { f:"none" } };
  }

  const ord = n => n + (n % 100 >= 11 && n % 100 <= 13 ? "th" : ["th","st","nd","rd"][n % 10] || "th");
  function repText(it){
    const r = it.rep || {}, D = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
    switch (r.f){
      case "daily": return "Every day";
      case "weekdays": return "Weekdays";
      case "ndays": return r.n % 7 === 0 && r.n > 7 ? `Every ${r.n/7} weeks` : r.n === 7 ? "Every week" : `Every ${r.n} days`;
      case "weekly": return (r.days && r.days.length ? "Every " + r.days.slice().sort((a,b) => ((a+6)%7) - ((b+6)%7)).map(d => D[d]).join(", ") : "Every " + D[dow(it.due)]);
      case "monthly": return (+r.n > 1 ? `Every ${r.n} months` : "Monthly") + " on the " + ord(+(it.anchor || it.due).slice(8,10));
      case "yearly": return "Every year";
    }
    return "";
  }
  return { pad, toMs, addMin, dow, wallNow, minsBetween, nextDue, complete, snooze, merge, parse, repText, repeats, at, ord };
})();
