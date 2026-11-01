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
  if (p.op === 'remAct' || p.op === 'remAdd') return ContentService.createTextOutput(remRemote_(p));
  if (!allowed_(doc, p.key)) return out_({ ok: false, error: 'wrong key' });
  const props = PropertiesService.getScriptProperties();
  const ver = () => Number(props.getProperty('ver_' + doc) || 0);

  if (p.op === 'ver') return out_({ ok: true, version: ver() });
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
    if (left !== 3) return;                                      // one reminder, 3 days before
    const key = 'date_' + dt.id + '_' + iso + '_' + left;
    if (props.getProperty(key)) return;
    const years = dt.yearly !== false && p[0] > 1900 ? Number(iso.slice(0, 4)) - p[0] : 0;
    const kind = dt.type || (dt.icon === '🎂' ? 'birthday' : dt.icon === '💍' ? 'anniversary' : 'other');
    const extra = years ? (kind === 'birthday' ? ' (turns ' + years + ')' : kind === 'anniversary' ? ' (' + years + ' years)' : ' (' + years + ' years)') : '';
    UrlFetchApp.fetch('https://ntfy.sh/' + topic, {
      method: 'post', muteHttpExceptions: true,
      payload: 'In 3 days: ' + dt.name + extra,
      headers: { Title: 'Coming up - Home Base', Tags: 'tada', Click: APP_URL }
    });
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
        if (now >= ea && now < it.due && st.e !== it.due) { remPush_(d, it, 'early', url, who); st.e = it.due; }
      }
      if (now >= when && st.m !== when) { remPush_(d, it, 'due', url, who); st.m = when; st.n = RL.addMin(now, 15); st.c = 0; }
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
      if (left.length && now >= today + 'T' + r.time && prev.d !== today) { remPushRoutine_(d, r, left, url, who); prev.d = today; }
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
  UrlFetchApp.fetch('https://ntfy.sh/', { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
}

function remPushRoutine_(d, r, left, url, who) {
  const all = url + '?op=remAct&who=' + who + '&id=' + encodeURIComponent(r.id) + '&a=rtdone&k=' + encodeURIComponent(d.key);
  const body = {
    topic: d.topic, title: d.hide ? 'Home Base' : (r.icon || '') + ' ' + r.name,
    message: d.hide ? 'Time for your routine' : left.length + ' of ' + r.steps.length + ' to do: ' + left.slice(0, 4).map(s => s.text).join(', ') + (left.length > 4 ? '…' : ''),
    priority: 3, tags: ['sunrise'],
    actions: [{ action: 'http', label: 'All done', url: all, method: 'GET', clear: true }]
  };
  UrlFetchApp.fetch('https://ntfy.sh/', { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
}

function setupReminders() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'checkReminders') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(1).create();
  checkReminders();
  Logger.log('Reminders are on. They are checked every minute.');
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
    ["topic","key","quiet","hide","lists","routines","rdone","su"].forEach(k => { if (s[k] !== undefined) out[k] = s[k]; else if ((s === a ? b : a)[k] !== undefined) out[k] = (s === a ? b : a)[k]; });
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
