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
  if (p.op === 'notify') { try { ['S', 'D'].forEach(w => { if (w !== p.who) wpSend_(w, { title: String(p.title || 'Home Base').slice(0, 80), body: String(p.message || '').slice(0, 200) }); }); return out_({ ok: true }); } catch (e) { return out_({ ok: false, error: e.message }); } }
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


/* ---------- Daily checks: gas bottle, important dates and bills ----------
 * Run once a day after 8am by the every-minute check, and sent as Home Base notifications.
 */
const APP_URL = 'https://bmpwestside.github.io/home-base/';

function checkGasBottle() {
  const props = PropertiesService.getScriptProperties();
  const data = read_('house', Number(props.getProperty('ver_house') || 0));
  if (!data) return;
  try { checkDates_(data); } catch (e) { Logger.log('Dates: ' + e.message); }
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
  hbBroadcast_('🔥 Gas bottle due', gasMsg);
  props.setProperty('gas_alerted', last);
}

/* Bill alerts: 2 days before and on the due day. */
function checkBills_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM-dd');
  const day = s => { const p = s.split('-').map(Number); return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5); };
  let data; try { data = bbData_(); } catch (e) { return; }
  const names = {}; (data.providers || []).forEach(p => { names[p.id] = p.name; });
  data.bills.forEach(b => {
    if (b.status !== 'due' || String(b.autopay) === 'true' || !b.dueDate) return;
    const left = day(b.dueDate) - day(today);
    if (left !== 2 && left !== 0) return;
    const key = 'bill_' + b.id + '_' + left; if (props.getProperty(key)) return;
    hbBroadcast_('💸 ' + (names[b.provider] || 'Bill') + ' due ' + (left ? 'in 2 days' : 'today'), '$' + Number(b.amount || 0).toFixed(2) + ' due ' + b.dueDate.split('-').reverse().join('/'));
    props.setProperty(key, '1');
  });
}

/** Sends a test notification to every phone in the house that turned notifications on. */
function testGasAlert() { hbBroadcast_('Home Base', 'Notifications are working ✅'); }

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
  if (it.hasNext()) { const f = it.next(); props.setProperty('bb_sheet', f.getId()); return SpreadsheetApp.openById(f.getId()); }
  // No BillBell yet (a new household): make the bills sheet in this Google Drive, ready for bills added in the app.
  if (typeof BS !== 'undefined') { const s2 = BS.createSheet(); props.setProperty('bb_sheet', s2.getId()); return s2; }
  const ss = SpreadsheetApp.create(BILLBELL_SHEET);
  const bills = ss.getSheets()[0]; bills.setName('Bills');
  bills.getRange(1, 1, 1, 21).setValues([['id', 'provider', 'category', 'amount', 'dueDate', 'issueDate', 'period', 'status', 'paidDate', 'autopay', 'billerCode', 'bpayRef', 'accountNo', 'invoiceNo', 'notes', 'reason', 'flags', 'fileUrl', 'emailUrl', 'source', 'createdAt']]);
  ss.insertSheet('Providers').getRange(1, 1, 1, 10).setValues([['id', 'name', 'category', 'colour', 'short', 'cycle', 'active', 'payUrl', 'domain', 'logo']]);
  props.setProperty('bb_sheet', ss.getId());
  return ss;
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
  return { bills: bills, providers: bbProviders_(), settings: {}, generated: new Date().toISOString() };
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


/* ---------- Important dates: a notification 3 days before ---------- */
function checkDates_(data) {
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
    hbBroadcast_((dt.icon || '🎉') + ' Coming up', 'In 3 days: ' + dt.name + extra);
    props.setProperty(key, '1');
  });
}


/* ---------- Reminders (private to each person, sent as Home Base notifications) ----------
 * Run setupReminders() once from the editor. It checks every minute and sends due reminders
 * to that person's phones as Home Base notifications.
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
      d.rhist = d.rhist || {}; d.rhist[r.id] = (d.rhist[r.id] || []).filter(x => x !== now.slice(0, 10)).concat([now.slice(0, 10)]).sort().slice(-120);
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
    try { checkBills_(); } catch (e) { Logger.log('Bills: ' + e.message); }
  }
  try { checkReminders__(props); props.deleteProperty('rem_lasterr'); }
  catch (e) { props.setProperty('rem_lasterr', remNow_() + ' ' + e.message); throw e; }
  if (typeof BS !== 'undefined' && Date.now() - Number(props.getProperty('bs_last') || 0) > 14 * 60000) {   // read new bill emails every 15 minutes
    props.setProperty('bs_last', String(Date.now()));
    try { BS.scan(); props.deleteProperty('bs_err'); } catch (e) { props.setProperty('bs_err', remNow_() + ' ' + e.message); }
  }
}
function checkReminders__(props) {
  const url = ScriptApp.getService().getUrl();
  ['S', 'D'].forEach(who => {
    let d; try { d = remRead_(who); } catch (e) { return; }
    if (!d || (!(d.items || []).length && !(d.routines || []).length && !(d.summary && d.summary.on))) return;
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
    const sm = d.summary;
    if (sm && sm.on && hm >= (sm.time || '07:00') && props.getProperty('remsum_' + who) !== today) {
      props.setProperty('remsum_' + who, today);
      const msg = remSummary_(d, now);
      if (msg) remSend_(who, { title: '☀️ Today', message: msg, tag: 'summary', badge: remBadge_(d) }, 'Summary', d);
    }
  });
}

function remPush_(d, it, kind, url, who) {
  const title = d.hide ? 'Reminder' : '⏰ ' + it.title;
  const message = d.hide ? 'You have a reminder' : kind === 'early' ? 'Coming up · ' + remWhen_(it.due) : kind === 'nag' ? 'Still to do · tap to open' : 'Due ' + remWhen_(it.snooze || it.due);
  return remSend_(who, { title: title, message: message, tag: it.id, badge: remBadge_(d), priority: it.hi ? 5 : kind === 'nag' ? 4 : 3 }, it.title, d);
}
/* Number on the Home Base icon: overdue reminders plus bills due within 3 days. */
function remBadge_(d) {
  const now = remNow_(); let n = (d.items || []).filter(x => !x.done && (x.snooze || x.due) <= now).length;
  n += (d.routines || []).filter(r => remRoutineLate_(d, r, now)).length;
  try { n += bbUrgent_(); } catch (e) {}
  return n;
}
function remRoutineOn_(r, day) { return !r.days || !r.days.length || r.days.indexOf(RL.dow(day + 'T00:00')) >= 0; }
function remRoutineLate_(d, r, now) {
  const day = now.slice(0, 10);
  if (!r.time || !(r.steps || []).length || !remRoutineOn_(r, day) || now < day + 'T' + r.time) return false;
  const done = d.rdone && d.rdone[r.id] && d.rdone[r.id].date === day ? d.rdone[r.id].steps || [] : [];
  return r.steps.some(s => done.indexOf(s.id) < 0);
}
/* The morning summary: today's reminders, routines and bills due soon, in one notification. */
function remSummary_(d, now) {
  const day = now.slice(0, 10), parts = [];
  const due = (d.items || []).filter(x => !x.done && (x.snooze || x.due).slice(0, 10) <= day);
  if (due.length) parts.push(due.length + ' reminder' + (due.length > 1 ? 's' : '') + ': ' + due.slice(0, 3).map(x => x.title).join(', ') + (due.length > 3 ? '…' : ''));
  const rts = (d.routines || []).filter(r => (r.steps || []).length && remRoutineOn_(r, day));
  if (rts.length) parts.push('Routines: ' + rts.map(r => r.name).join(', '));
  try {
    if (typeof BILLS_ENABLED === 'undefined' || BILLS_ENABLED) {
      const lim = RL.addMin(day + 'T00:00', 2 * 1440).slice(0, 10), data = bbData_(), names = {};
      (data.providers || []).forEach(p => { names[p.id] = p.name; });
      const bills = data.bills.filter(b => b.status === 'due' && b.dueDate && b.dueDate <= lim && String(b.autopay) !== 'true');
      if (bills.length) parts.push('Bills: ' + bills.slice(0, 3).map(b => (names[b.provider] || 'Bill') + ' $' + Number(b.amount || 0).toFixed(0) + (b.dueDate < day ? ' (overdue)' : b.dueDate === day ? ' (today)' : '')).join(', '));
    }
  } catch (e) {}
  return parts.join(' · ');
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
  const code = wpSubs_(who).length ? wpSend_(who, { title: body.title, body: body.message, tag: body.tag, badge: body.badge, urgent: body.priority >= 4 }) : 'notifications are not turned on';
  const props = PropertiesService.getScriptProperties(), k = 'remlog_' + who;
  const log = JSON.parse(props.getProperty(k) || '[]'); log.unshift({ at: remNow_(), t: label, code: code });
  props.setProperty(k, JSON.stringify(log.slice(0, 6)));
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
function remStatus_(p) {
  const who = remWho_(p.who), props = PropertiesService.getScriptProperties(), d = remRead_(who) || {};
  const open = (d.items || []).filter(x => !x.done).sort((a, b) => ((a.snooze || a.due) < (b.snooze || b.due) ? -1 : 1));
  let res = { ok: true, version: 3, now: remNow_(), triggers: ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkReminders').length,
    pushSubs: wpSubs_(who).length, engine: wpSelfTest_(), lastRun: props.getProperty('rem_lastrun') || '', lastErr: props.getProperty('rem_lasterr') || '', open: open.length,
    next: open[0] ? { t: open[0].title, at: open[0].snooze || open[0].due } : null, log: JSON.parse(props.getProperty('remlog_' + who) || '[]') };
  if (p.test) { res.test = remSend_(who, { title: 'Home Base', message: 'Test from your Google script ✅', badge: remBadge_(d) }, 'Test', d); res.testLog = JSON.parse(props.getProperty('remlog_' + who) || '[]')[0]; }
  return res;
}

function remPushRoutine_(d, r, left, url, who) {
  const title = d.hide ? 'Home Base' : (r.icon || '') + ' ' + r.name;
  const message = d.hide ? 'Time for your routine' : left.length + ' of ' + r.steps.length + ' to do: ' + left.slice(0, 4).map(s => s.text).join(', ') + (left.length > 4 ? '…' : '');
  return remSend_(who, { title: title, message: message, tag: 'rt_' + r.id, badge: remBadge_(d) }, r.name, d);
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
    ["topic","key","quiet","hide","lists","routines","rdone","rhist","via","summary","su"].forEach(k => { if (s[k] !== undefined) out[k] = s[k]; else if ((s === a ? b : a)[k] !== undefined) out[k] = (s === a ? b : a)[k]; });
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

/* ---------- Bill reader: finds bills in this Gmail and adds them to the Bills tab ----------
 * Based on BillBell's scanner. Every 15 minutes it reads new bill emails (and their PDF bills, if
 * Services → Drive API is turned on), and saves each bill to the "BillBell Data" sheet in this Drive.
 * It looks at emails from known providers, emails whose subject says bill / invoice / statement,
 * and anything you put under the Gmail label "Bills".
 */
var BS = (function () {
  var CONFIG = { LABEL: 'Bills', SHEET_NAME: 'BillBell Data', FOLDER_NAME: 'Home Base Bills', PRICE_RISE: 0.20,
    SCAN_DAYS: 30, BACKFILL_DAYS: 120, MAX_MS: 50000, TZ: 'Australia/Sydney' };
  var BILL_COLS = ['id', 'provider', 'category', 'amount', 'dueDate', 'issueDate', 'period', 'invoiceNo',
    'accountNo', 'billerCode', 'bpayRef', 'status', 'paidDate', 'autopay', 'flags', 'reason',
    'fileUrl', 'emailUrl', 'messageId', 'source', 'notes', 'createdAt'];
  var PROVIDER_COLS = ['id', 'name', 'category', 'match', 'colour', 'cycle', 'payUrl', 'short', 'domain', 'active', 'logo'];
  var PAYMENT_COLS = ['date', 'provider', 'amount', 'invoiceNo', 'billId', 'messageId'];
  var LOG_COLS = ['messageId', 'date', 'subject', 'result'];
  var _cache = {};

/**
 * BillBell — bill text parser (pure JavaScript, no Google services).
 * Reads the text of an email or PDF and pulls out the bill details.
 */

var MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
var AMT = '\\$?\\s*((?:\\d{1,3}(?:,\\d{3})+|\\d+)\\.\\d{2})';
var DATE = '(\\d{1,2}[\\/.\\-]\\d{1,2}[\\/.\\-]\\d{2,4}|\\d{1,2}(?:st|nd|rd|th)?[\\s\\-]+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\\.?[\\s,\\-]+\\d{2,4})';

/** Default providers. Order matters: more specific patterns first. */
var DEFAULT_PROVIDERS = [
  // id, name, category, match (regex), colour, cycle days, pay URL, short, website
  ['telstra', 'Telstra', 'Telecommunications', 'telstra', '#0099F8', 30, 'https://www.telstra.com.au/my-account', 'T', 'telstra.com.au'],
  ['optus', 'Optus', 'Internet', 'optus', '#139C9F', 30, 'https://www.optus.com.au/paymybill', 'O', 'optus.com.au'],
  ['vodafone', 'Vodafone', 'Telecommunications', 'vodafone', '#E60000', 30, 'https://www.vodafone.com.au/my-vodafone', 'V', 'vodafone.com.au'],
  ['tpg', 'TPG', 'Internet', '\\btpg\\b', '#6A2C91', 30, 'https://www.tpg.com.au/myaccount', 'TP', 'tpg.com.au'],
  ['aussie-broadband', 'Aussie Broadband', 'Internet', 'aussie ?broadband|aussiebb', '#00A76F', 30, 'https://my.aussiebroadband.com.au', 'AB', 'aussiebroadband.com.au'],
  ['amaysim', 'Amaysim', 'Telecommunications', 'amaysim', '#F26A21', 28, 'https://www.amaysim.com.au/my-account', 'AM', 'amaysim.com.au'],
  ['agl', 'AGL', 'Electricity', '\\bagl\\b|agl\\.com\\.au', '#1E88E5', 30, 'https://myaccount.agl.com.au', 'AG', 'agl.com.au'],
  ['origin', 'Origin Energy', 'Electricity', 'origin energy|originenergy', '#F15A22', 30, 'https://www.originenergy.com.au/for-home/my-account', 'OE', 'originenergy.com.au'],
  ['energyaustralia', 'EnergyAustralia', 'Electricity', 'energy ?australia', '#00A651', 30, 'https://www.energyaustralia.com.au/home/my-account', 'EA', 'energyaustralia.com.au'],
  ['sydney-water', 'Sydney Water', 'Water', 'sydney ?water', '#0077C8', 91, 'https://www.sydneywater.com.au', 'SW', 'sydneywater.com.au']
];

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/** "26 Oct 2026", "30/11/2026", "06-OCT-26" → "2026-10-26" (Australian day-first). */
function parseDateStr(s) {
  if (!s) return '';
  s = String(s).trim().toLowerCase();
  var m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  var d, mo, y;
  if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else {
    m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[\s\-]+([a-z]{3})[a-z]*\.?[\s,\-]+(\d{2,4})$/);
    if (!m) return '';
    d = +m[1]; mo = MONTHS[m[2]]; y = +m[3];
  }
  if (y < 100) y += 2000;
  if (!mo || mo > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return '';
  return y + '-' + pad2_(mo) + '-' + pad2_(d);
}

function firstMatch_(text, patterns) {
  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m && m[1]) return m[1];
  }
  return '';
}

function toNum_(s) { return s ? parseFloat(String(s).replace(/,/g, '')) : 0; }

/** Finds the provider whose pattern matches the text. providers: array of {id, match,...}. */
function matchProvider(text, providers) {
  var t = String(text || '');
  for (var i = 0; i < providers.length; i++) {
    var p = providers[i];
    if (!p.match) continue;
    try { if (new RegExp(p.match, 'i').test(t)) return p; } catch (e) {}
  }
  return null;
}

/* ---------- New providers ---------- */

var PERSONAL_DOMAINS = /(^|\.)(gmail|googlemail|icloud|me|mac|outlook|hotmail|live|yahoo|bigpond)\.(com|com\.au|net|net\.au)$/i;
var CATEGORY_WORDS = [
  ['Electricity', /electricity|kwh|energy usage|power bill|solar feed/i],
  ['Gas', /\bgas\b|\blpg\b|cylinder|megajoules|\bmj\b/i],
  ['Water', /\bwater\b|sewerage|wastewater/i],
  ['Internet', /\bnbn\b|internet|broadband|wi-?fi|fibre/i],
  ['Telecommunications', /mobile|\bsim\b|prepaid|postpaid|phone plan|data plan/i],
  ['Council Rates', /council|\brates\b|instalment notice/i],
  ['Insurance', /insurance|policy|premium|insurer/i],
  ['Rego', /registration|\brego\b|\bctp\b|transport for nsw|service nsw|e-?toll|linkt/i],
  ['Strata', /strata|levy|owners corporation|body corporate/i],
  ['Subscriptions', /subscription|membership|renewal/i]
];
var NEW_COLOURS = ['#5B6BD6', '#139C9F', '#C0563A', '#2E7D32', '#8E44AD', '#D9822B', '#0077C8', '#B03A5B', '#3A7D7B', '#6D4C41'];

/** "AGL <hello@mail.agl.com.au>" → { name: 'AGL', email, domain: 'agl.com.au' } */
function senderInfo(from) {
  var s = String(from || '');
  var email = (s.match(/<([^>]+)>/) || s.match(/([^\s<>"]+@[^\s<>"]+)/) || [])[1] || '';
  var name = s.replace(/<[^>]*>/, '').replace(/["']/g, '').trim();
  var host = (email.split('@')[1] || '').toLowerCase();
  var parts = host.split('.');
  var keep = /\.(com|net|org|gov|edu|asn|id)\.au$|\.co\.(uk|nz)$/.test(host) ? 3 : 2;
  var domain = parts.slice(-keep).join('.');
  return { name: name, email: email.toLowerCase(), domain: domain };
}

function guessCategory(text) {
  var t = String(text || '').slice(0, 6000);
  for (var i = 0; i < CATEGORY_WORDS.length; i++) if (CATEGORY_WORDS[i][1].test(t)) return CATEGORY_WORDS[i][0];
  return 'Other';
}

/** Builds a provider row for a sender BillBell hasn't seen. Returns null for personal email addresses. */
function makeProvider(from, text) {
  var si = senderInfo(from);
  if (!si.domain || PERSONAL_DOMAINS.test(si.domain)) return null;
  var brand = si.domain.split('.')[0];
  var name = si.name && !/@|no.?reply|do.?not.?reply|notifications?|billing|accounts?$/i.test(si.name)
    ? si.name.replace(/\s+(billing|accounts|notifications?|team|customer service)$/i, '')
    : brand.charAt(0).toUpperCase() + brand.slice(1);
  var id = brand.toLowerCase().replace(/[^a-z0-9]+/g, '-');
  var hash = 0; for (var i = 0; i < id.length; i++) hash = (hash * 31 + id.charCodeAt(i)) >>> 0;
  var words = name.split(/\s+/).filter(Boolean);
  return {
    id: id, name: name, category: guessCategory(name + ' ' + si.domain + ' ' + text),
    match: [brand, si.domain.replace(/\./g, '\\.'), name.toLowerCase().replace(/[^a-z0-9 ]/g, '').trim()].filter(function (x, i, arr) { return x && arr.indexOf(x) === i; }).join('|'),
    colour: NEW_COLOURS[hash % NEW_COLOURS.length], cycle: 30,
    payUrl: 'https://www.' + si.domain, short: (words.length > 1 ? words[0][0] + words[1][0] : name.slice(0, 2)).toUpperCase(),
    domain: si.domain, active: 'true'
  };
}

/** Pulls the original sender/subject/date out of a forwarded email body. */
function parseForwarded(body) {
  var out = { from: '', subject: '', date: '' };
  var t = String(body || '');
  var block = t.match(/(?:forwarded message|begin forwarded message|original message)[\s\S]{0,800}/i);
  var src = block ? block[0] : t.slice(0, 1500);
  var f = src.match(/^\s*from:\s*(.+)$/im); if (f) out.from = f[1].trim();
  var s = src.match(/^\s*subject:\s*(.+)$/im); if (s) out.subject = s[1].trim();
  var d = src.match(/^\s*date:\s*(.+)$/im); if (d) out.date = d[1].trim();
  return out;
}

/**
 * Reads one document (email body or PDF text).
 * Returns { kind: 'bill' | 'paid-bill' | 'receipt' | 'unknown', ...fields }
 */
function parseBillText(text, emailDateIso) {
  var t = String(text || '').replace(/ /g, ' ').replace(/[ \t]+/g, ' ');
  var r = {
    kind: 'unknown', amount: 0, dueDate: '', issueDate: '', period: '', invoiceNo: '',
    accountNo: '', billerCode: '', bpayRef: '', autopay: false, dueNow: false,
    credits: [], receiptInvoices: [], paidDate: '', reasons: []
  };

  // ---- Payment receipt (e.g. AGL "PAYMENT RECEIPT") ----
  if (/payment receipt|receipt for your payment|payment confirmation/i.test(t) &&
      !/(?:total )?amount due[^0-9$]{0,40}\$?\s*[1-9]/i.test(t)) {
    r.kind = 'receipt';
    r.amount = toNum_(firstMatch_(t, [
      new RegExp('payment total[^0-9$]{0,30}' + AMT, 'i'),
      new RegExp('amount paid[^0-9$]{0,30}' + AMT, 'i'),
      new RegExp('payment of[^0-9$]{0,10}' + AMT, 'i'),
      new RegExp('\\btotal\\b[^0-9$]{0,20}' + AMT, 'i')
    ]));
    r.paidDate = parseDateStr(firstMatch_(t, [
      new RegExp('(?:payment receipt|payment date|date paid|received on)[^0-9a-z]{0,30}' + DATE, 'i'),
      new RegExp(DATE, 'i')
    ])) || emailDateIso || '';
    var re = /invoice\s*(?:no\.?|number|#)?\s*[:#]?\s*(\d{5,})/gi, mm;
    while ((mm = re.exec(t))) r.receiptInvoices.push(mm[1]);
    return r;
  }

  // ---- Amount ----
  var owing = toNum_(firstMatch_(t, [new RegExp('total owing[^0-9$]{0,20}' + AMT, 'i')]));
  var amt = firstMatch_(t, [
    new RegExp('total amount due[^0-9$]{0,40}' + AMT, 'i'),
    new RegExp('amount due[^0-9$]{0,40}' + AMT, 'i'),
    new RegExp('total due[^0-9$]{0,40}' + AMT, 'i'),
    new RegExp('(?:amount payable|balance due|amount to pay|please pay)[^0-9$]{0,40}' + AMT, 'i'),
    new RegExp('total (?:incl(?:\\.|uding)?|inc\\.?)\\s*gst[^0-9$]{0,20}' + AMT, 'i'),
    new RegExp('\\btotal\\b(?!\\s*(?:owing|excl|ex\\b|payments|for this|gst))[^0-9$]{0,20}' + AMT, 'i')
  ]);
  r.amount = toNum_(amt) || owing;
  if (!r.amount && /tax invoice|invoice|renewal|subscription|your receipt|billing/i.test(t)) {
    // Price shown without a label (e.g. Apple): use the largest dollar amount on the page
    var all = t.match(/\$\s*(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}/g) || [];
    var vals = all.map(function (x) { return toNum_(x.replace(/[$\s]/g, '')); }).filter(function (v) { return v > 0; });
    if (vals.length) { r.amount = Math.max.apply(null, vals); r.guessedAmount = true; }   // the total is the biggest figure
  }

  // ---- Dates ----
  r.dueDate = parseDateStr(firstMatch_(t, [
    new RegExp('(?:this bill due date|due date|payment due|instalment due|due by|due on|pay by|please pay by)[^0-9a-z]{0,30}' + DATE, 'i'),
    new RegExp('(?:debit date|will be debited on)[^0-9a-z]{0,30}' + DATE, 'i'),
    new RegExp('(?:next billing date|next payment date|renewal date|billing date|charge date|will be (?:charged|billed|renewed) on)[^0-9a-z]{0,30}' + DATE, 'i'),
    new RegExp('\\bdue\\b[^0-9a-z]{0,15}' + DATE, 'i')
  ]));
  r.issueDate = parseDateStr(firstMatch_(t, [
    new RegExp('(?:issue date|invoice date|date of issue|date issued|issued)[^0-9a-z]{0,30}' + DATE, 'i'),
    new RegExp('(?:tax invoice|customer copy)[^0-9]{0,60}' + DATE, 'i')
  ]));
  if (!r.dueDate && /due now|c\.?o\.?d\.?\b|payable immediately|due immediately/i.test(t)) {
    r.dueNow = true;
    r.dueDate = r.issueDate || emailDateIso || '';
  }

  // ---- Reference numbers ----
  r.invoiceNo = firstMatch_(t, [
    /invoice\s*(?:no\.?|number|#)\s*[:#]?\s*((?=[A-Z0-9\-]*\d)[A-Z0-9\-]{5,})/i,
    /notice\s*(?:no\.?|number)\s*:?\s*(\d{5,})/i,
    /tax invoice\s+((?=[A-Z0-9]*\d)[A-Z0-9]{5,})/i
  ]);
  r.accountNo = firstMatch_(t, [
    /(?:account\s*(?:number|no\.?|code)|customer\s*(?:number|no\.?)|amaysim no\.?)\s*[:.]?\s*([A-Z]?\d[\d ]{2,18}\d)/i
  ]).replace(/\s+/g, ' ').trim();
  r.billerCode = firstMatch_(t, [/biller\s*code\s*:?\s*(\d{3,7})/i]);
  if (r.billerCode) {
    var idx = t.search(/biller\s*code/i);
    var after = t.slice(idx, idx + 220);
    r.bpayRef = firstMatch_(after, [
      /(?:ref(?:erence)?\s*(?:number|no\.?)?|cust(?:omer)?\s*ref(?:erence)?|crn)\s*:?\s*(\d[\d ]{3,24}\d)/i
    ]).replace(/\s+/g, '');
  }
  if (r.billerCode && !r.bpayRef) {
    r.bpayRef = firstMatch_(t, [
      /(?:bpay\s*)?(?:ref(?:erence)?\s*(?:number|no\.?)|customer\s*ref(?:erence)?|crn)\s*:?\s*(\d[\d ]{4,24}\d)/i
    ]).replace(/\s+/g, '');
  }
  r.period = firstMatch_(t, [
    /(?:bill(?:ing)? period|invoice period|supply period|period)\s*:?\s*(\d{1,2}\s*[a-z]{3}[a-z]*\s*\d{0,4}\s*(?:to|\-|–)\s*\d{1,2}\s*[a-z]{3}[a-z]*\s*\d{2,4})/i
  ]);

  r.autopay = /(?:will be|is) (?:automatically )?(?:debited|deducted)|automatically debited|direct debit date|payment method\s*:?\s*(?:automatic|direct debit)|next billing date|auto-?renew|will be (?:automatically )?(?:charged|billed|renewed)/i.test(t);

  // Credits on this bill = payments for earlier bills (e.g. "$84.29 CR")
  var cre = /\$?((?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2})\s*cr\b/gi, cm;
  while ((cm = cre.exec(t))) r.credits.push(toNum_(cm[1]));

  // ---- Classify ----
  var paidInvoice = /no further action required|thanks? (?:you )?for your payment|total owing\s*\$?\s*0\.00|paid in full/i.test(t);
  if (r.amount > 0) {
    r.kind = paidInvoice ? 'paid-bill' : 'bill';
    if (paidInvoice) {
      r.paidDate = parseDateStr(firstMatch_(t, [new RegExp('(?:payment received on|paid on)[^0-9a-z]{0,20}' + DATE, 'i')])) || r.issueDate || emailDateIso || '';
    }
  }
  if (!r.amount) r.reasons.push('No amount found');
  if (r.guessedAmount) r.reasons.push('Amount taken from the price on the invoice — check it');
  if (r.kind === 'bill' && !r.dueDate) r.reasons.push('No due date found');
  if (r.dueNow) r.reasons.push('Bill says "Due now" — due date set to the invoice date');
  return r;
}

/** Bills that look like the same bill (same provider + invoice, or same amount + due date). */
function isDuplicateBill(newBill, bills) {
  for (var i = 0; i < bills.length; i++) {
    var b = bills[i];
    if (b.provider !== newBill.provider || b.status === 'ignored') continue;
    if (newBill.invoiceNo && b.invoiceNo && String(b.invoiceNo) === String(newBill.invoiceNo)) return b;
    if (Math.abs(Number(b.amount) - Number(newBill.amount)) < 0.005 && b.dueDate && b.dueDate === newBill.dueDate) return b;
  }
  return null;
}

/** % rise against the average of the last 3 bills for the provider, or 0. */
function priceRise(newBill, bills, threshold) {
  var prev = bills.filter(function (b) {
    return b.provider === newBill.provider && b.status !== 'ignored' && b.id !== newBill.id &&
      Number(b.amount) > 0 && (b.dueDate || '') < (newBill.dueDate || '9999');
  }).sort(function (a, b) { return a.dueDate < b.dueDate ? 1 : -1; }).slice(0, 3);
  if (prev.length < 2) return 0;
  var avg = prev.reduce(function (s, b) { return s + Number(b.amount); }, 0) / prev.length;
  var rise = (Number(newBill.amount) - avg) / avg;
  return rise >= (threshold || 0.2) ? Math.round(rise * 100) : 0;
}

/** Providers whose regular bill is late. Returns [{provider, daysLate, last, every}]. */
function missingBills(bills, providers, todayIso) {
  var out = [];
  var today = new Date(todayIso + 'T00:00:00');
  providers.forEach(function (p) {
    if (String(p.active) === 'false' || !(Number(p.cycle) > 0)) return;
    var dates = bills.filter(function (b) { return b.provider === p.id && b.status !== 'ignored'; })
      .map(function (b) { return b.issueDate || b.dueDate; }).filter(Boolean).sort();
    if (!dates.length) return;
    var gaps = [];
    for (var i = 1; i < dates.length; i++) {
      var g = (new Date(dates[i]) - new Date(dates[i - 1])) / 864e5;
      if (g > 5) gaps.push(g);
    }
    var every = Number(p.cycle);
    if (gaps.length >= 2) { gaps.sort(function (a, b) { return a - b; }); every = Math.round(gaps[Math.floor(gaps.length / 2)]); }
    var last = new Date(dates[dates.length - 1] + 'T00:00:00');
    var late = Math.floor((today - last) / 864e5) - every;
    if (late > 5) out.push({ provider: p.id, daysLate: late, last: dates[dates.length - 1], every: every });
  });
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = { parseBillText: parseBillText, parseDateStr: parseDateStr, matchProvider: matchProvider,
    parseForwarded: parseForwarded, isDuplicateBill: isDuplicateBill, priceRise: priceRise,
    missingBills: missingBills, DEFAULT_PROVIDERS: DEFAULT_PROVIDERS,
    senderInfo: senderInfo, guessCategory: guessCategory, makeProvider: makeProvider };
}

function scanBills_(days, quiet) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { added: 0 };
  var started = Date.now(), added = [];
  try {
    tidyReview_();
    var done = {};
    readTable_('Log').forEach(function (r) { done[r.messageId] = true; });
    var start = 0, threads;
    do {
      // Forwarded emails keep their original "To", so also look at where Gmail delivered them
      threads = GmailApp.search(searchQuery_() + ' newer_than:' + days + 'd', start, 100);
      
      for (var i = 0; i < threads.length; i++) {
        var msgs = threads[i].getMessages();
        for (var j = 0; j < msgs.length; j++) {
          var m = msgs[j];
          if (done[m.getId()]) continue;
          if (Date.now() - started > CONFIG.MAX_MS) {
            if (days >= CONFIG.BACKFILL_DAYS) PropertiesService.getScriptProperties().setProperty('BACKFILL_LEFT', '1');
            Logger.log('Read ' + added.length + ' bills so far. The rest will be read automatically over the next few runs.');
            return { added: added.length, more: true };
          }
          var res;
          try { res = processMessage_(m); } catch (e) { res = { result: 'error: ' + e.message, bills: [] }; }
          appendRow_('Log', { messageId: m.getId(), date: iso_(m.getDate()), subject: m.getSubject(), result: res.result });
          done[m.getId()] = true;
          added = added.concat(res.bills || []);
        }
      }
      start += 100;
    } while (threads.length === 100);
    if (days >= CONFIG.BACKFILL_DAYS) PropertiesService.getScriptProperties().deleteProperty('BACKFILL_LEFT');
  } finally { lock.releaseLock(); }

  if (!quiet) {
    added.forEach(function (b) {
      var p = providerById_(b.provider);
      var name = p ? p.name : 'Unknown provider';
      if (b.status === 'review') notify_('Check a bill: ' + name, money_(b.amount) + ' — ' + (b.reason || 'needs a quick check'), 'mag');
      else if (b.status === 'due') notify_('New bill: ' + name, money_(b.amount) + ' due ' + niceDate_(b.dueDate) + (b.flags ? ' · ' + b.flags : ''), 'receipt');
    });
  }
  return { added: added.length };
}

function processMessage_(msg) {
  var body = msg.getPlainBody() || stripHtml_(msg.getBody());
  var fwd = parseForwarded(body);
  var emailDate = iso_(msg.getDate());
  var fwdDate = parseDateStr((fwd.date.match(/\d{1,2}\s+[a-z]{3}[a-z]*\s+\d{4}/i) || [''])[0]);
  if (fwdDate) emailDate = fwdDate;
  var header = [msg.getFrom(), msg.getSubject(), fwd.from, fwd.subject].join('\n');
  var providers = getProviders_();
  var emailUrl = 'https://mail.google.com/mail/u/0/#all/' + msg.getId();
  var base = { messageId: msg.getId(), emailUrl: emailUrl, source: 'email', sender: fwd.from || msg.getFrom() };

  // 1) Each PDF / photo attachment is read on its own
  var docs = [];
  msg.getAttachments({ includeInlineImages: false }).forEach(function (a) {
    var type = a.getContentType() || '';
    var isPdf = /pdf/i.test(type) || /\.pdf$/i.test(a.getName());
    var isImg = /^image\/(jpe?g|png|heic|heif)/i.test(type) && a.getSize() > 30000;
    if (!isPdf && !isImg) return;
    var file = getFolder_().createFile(a.copyBlob()).setName(emailDate + ' ' + a.getName());
    var text = '';
    try { text = ocrText_(a.copyBlob(), a.getName()); } catch (e) { text = ''; }
    docs.push({ text: text, fileUrl: file.getUrl() });
  });

  var bills = [], results = [];
  docs.forEach(function (d) {
    var r = handleDoc_(header + '\n' + d.text, emailDate, providers, merge_(base, { fileUrl: d.fileUrl }));
    if (r.bill) bills.push(r.bill);
    results.push(r.result);
  });

  // 2) If no attachment held the bill, read the email text itself
  var anyHit = results.some(function (r) { return !/^skip/.test(r); });
  if (!anyHit) {
    var r = handleDoc_(header + '\n' + body, emailDate, providers, merge_(base, { fileUrl: docs[0] ? docs[0].fileUrl : '' }));
    if (r.bill) bills.push(r.bill);
    results.push(r.result);
    // Known provider but nothing readable → send to Review so nothing is missed
    if (/^skip/.test(r.result)) {
      var p = matchProvider(header + '\n' + body, providers);
      if (p && hasBillNear_(p.id, emailDate)) {
        results.push('skip: reminder or copy of a bill already saved');
      } else if (p && /bill|invoice|statement|notice|amount|due/i.test(header + body.slice(0, 2000))) {
        var b = saveBill_(merge_(base, {
          provider: p.id, category: p.category, amount: 0, dueDate: '', status: 'review',
          reason: 'Couldn\'t read the amount — open the email and enter it', fileUrl: docs[0] ? docs[0].fileUrl : ''
        }));
        bills.push(b); results.push('review: unreadable');
      }
    }
  }
  return { result: results.join(' | ') || 'skip: empty', bills: bills };
}

function handleDoc_(text, emailDate, providers, base) {
  var info = parseBillText(text, emailDate);
  var p = matchProvider(text, providers);
  if (info.kind === 'unknown') return { result: 'skip: no amount' };
  var newProvider = false;
  if (!p && info.kind !== 'receipt') {
    p = addNewProvider_(base.sender, text);
    newProvider = !!p;
  }
  base = merge_(base, {}); delete base.sender;

  if (info.kind === 'receipt') {
    var matched = applyPayment_(p ? p.id : '', info.amount, info.receiptInvoices, info.paidDate, base.messageId);
    return { result: 'receipt ' + money_(info.amount) + (matched ? ' → paid ' + matched : ' (no bill yet)') };
  }

  var bill = merge_(base, {
    provider: p ? p.id : 'unknown', category: p ? p.category : 'Other',
    amount: info.amount, dueDate: info.dueDate, issueDate: info.issueDate, period: info.period,
    invoiceNo: info.invoiceNo, accountNo: info.accountNo, billerCode: info.billerCode,
    bpayRef: info.bpayRef, autopay: info.autopay ? 'true' : '', flags: '', reason: ''
  });

  var all = readTable_('Bills');
  var dup = isDuplicateBill(bill, all);
  if (dup) return { result: 'duplicate of ' + dup.id };

  // Status
  var reasons = info.reasons.slice();
  if (!p) reasons.unshift('Unknown provider');
  if (newProvider) reasons.unshift('New provider — check the name and category');
  if (info.kind === 'paid-bill') { bill.status = 'paid'; bill.paidDate = info.paidDate; }
  else if (reasons.length) bill.status = 'review';
  else bill.status = 'due';
  bill.reason = reasons.join('; ');

  // Already paid? (a receipt for this invoice arrived earlier)
  if (bill.status === 'due' && bill.invoiceNo) {
    var pay = readTable_('Payments').filter(function (x) { return String(x.invoiceNo) === String(bill.invoiceNo) && !x.billId; })[0];
    if (pay) { bill.status = 'paid'; bill.paidDate = pay.date; }
  }

  var rise = priceRise(bill, all, CONFIG.PRICE_RISE);
  if (rise) bill.flags = 'Up ' + rise + '% on usual';

  var saved = saveBill_(bill);
  if (saved.status !== 'review') clearUnreadable_(saved.provider, bill.issueDate || bill.dueDate || emailDate);

  // Credits on this bill mean earlier bills were paid
  if (p && info.credits.length) {
    all.filter(function (b) { return b.provider === p.id && (b.status === 'due' || b.status === 'review') && b.id !== saved.id; })
      .forEach(function (b) {
        if (info.credits.some(function (c) { return Math.abs(c - Number(b.amount)) < 0.005; }))
          updateRow_('Bills', b.id, { status: 'paid', paidDate: bill.issueDate || emailDate });
      });
  }
  return { result: saved.status + ' ' + (p ? p.id : 'unknown') + ' ' + money_(bill.amount), bill: saved };
}

function tidyReview_() {
  var bills = readTable_('Bills');
  bills.filter(isUnreadable_).forEach(function (r) {
    var when = r.messageId ? emailDateOf_(r.messageId) : String(r.createdAt || '').slice(0, 10);
    var real = bills.some(function (b) {
      return b.provider === r.provider && b.id !== r.id && Number(b.amount) > 0 && b.status !== 'ignored' &&
        (nearDate_(b.issueDate, when) || nearDate_(b.dueDate, when) || (b.emailDate && nearDate_(b.emailDate, when)));
    });
    if (real) updateRow_('Bills', r.id, { status: 'ignored', reason: 'Reminder for a bill already saved' });
  });
}

function emailDateOf_(id) {
  try { return iso_(GmailApp.getMessageById(id).getDate()); } catch (e) { return ''; }
}

function applyPayment_(providerId, amount, invoices, date, messageId) {
  var bills = readTable_('Bills').filter(function (b) { return b.status === 'due' || b.status === 'review'; });
  var hit = bills.filter(function (b) { return b.invoiceNo && invoices.indexOf(String(b.invoiceNo)) >= 0; })[0] ||
    bills.filter(function (b) { return b.provider === providerId && Math.abs(Number(b.amount) - amount) < 0.005; })
      .sort(function (a, b) { return a.dueDate < b.dueDate ? -1 : 1; })[0];
  if (hit) updateRow_('Bills', hit.id, { status: 'paid', paidDate: date });
  appendRow_('Payments', { date: date, provider: providerId, amount: amount, invoiceNo: invoices[0] || '', billId: hit ? hit.id : '', messageId: messageId });
  return hit ? hit.id : '';
}

function saveBill_(bill) {
  bill.id = bill.id || ('b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
  bill.createdAt = bill.createdAt || new Date().toISOString();
  appendRow_('Bills', bill);
  return bill;
}

function ocrText_(blob, name) {
  if (typeof Drive === 'undefined') return '';      // PDF/photo reading needs Services → Drive API
  var file = Drive.Files.create
    ? Drive.Files.create({ name: 'tmp-' + name, mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'en', fields: 'id' })
    : Drive.Files.insert({ title: 'tmp-' + name, mimeType: blob.getContentType() }, blob, { ocr: true, ocrLanguage: 'en', convert: true });
  try { return DocumentApp.openById(file.id).getBody().getText(); }
  finally { try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e) {} }
}

function addNewProvider_(sender, text) {
  var np = makeProvider(sender, text);
  if (!np) return null;
  var existing = readTable_('Providers').filter(function (x) { return x.id === np.id || x.domain === np.domain; })[0];
  if (existing) return existing;
  try { np.logo = fetchLogo_(np.domain) || ''; } catch (e) { np.logo = ''; }
  appendRow_('Providers', np);
  delete _cache['Providers'];
  Logger.log('New provider added: ' + np.name + ' (' + np.category + ')');
  return np;
}

function refreshLogos() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshLogos') ScriptApp.deleteTrigger(t); });
  delete _cache['Providers'];
  var done = {};
  readTable_('Providers').forEach(function (p) {
    if (!p.domain) return;
    if (p.logo) return;          // keep logos already saved, including ones you uploaded in the app
    var uri = done[p.domain] || fetchLogo_(p.domain);
    if (uri) { done[p.domain] = uri; updateRow_('Providers', p.id, { logo: uri }); }
    Logger.log(p.name + ': ' + (uri ? 'logo saved' : 'no logo found, initials will show'));
  });
}

function fetchLogo_(domain) {
  var hosts = ['https://www.' + domain.replace(/^www\./, ''), 'https://' + domain.replace(/^www\./, '')];
  var cands = [];
  // 1) Icons the website itself declares (largest first)
  for (var h = 0; h < hosts.length && !cands.length; h++) {
    try {
      var res = UrlFetchApp.fetch(hosts[h], { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X)' } });
      if (res.getResponseCode() >= 400) continue;
      var html = res.getContentText().slice(0, 300000), re = /<link\b[^>]*>/gi, m, found = [];
      while ((m = re.exec(html))) {
        var tag = m[0], rel = (tag.match(/rel=["']?([^"'>]+)/i) || [])[1] || '', href = (tag.match(/href=["']?([^"' >]+)/i) || [])[1];
        if (!href || !/icon/i.test(rel)) continue;
        var size = parseInt((tag.match(/sizes=["']?(\d+)/i) || [])[1] || (/apple-touch/i.test(rel) ? '180' : '32'), 10);
        found.push({ url: absUrl_(href.replace(/&amp;/g, '&'), hosts[h]), size: size });
      }
      found.sort(function (a, b) { return b.size - a.size; });
      cands = found.map(function (f) { return f.url; });
      cands.push(hosts[h] + '/apple-touch-icon.png');
    } catch (e) {}
  }
  // 2) Fallback: Google's icon service
  cands.push('https://www.google.com/s2/favicons?domain=' + domain + '&sz=128');

  for (var i = 0; i < cands.length; i++) {
    try {
      var r = UrlFetchApp.fetch(cands[i], { muteHttpExceptions: true, followRedirects: true });
      if (r.getResponseCode() !== 200) continue;
      var blob = r.getBlob(), type = (blob.getContentType() || '').split(';')[0], bytes = blob.getBytes();
      if (!/^image\//.test(type) && /\.svg(\?|$)/i.test(cands[i])) type = 'image/svg+xml';
      if (!/^image\//.test(type) || bytes.length < 600 || bytes.length > 36000) continue;
      return 'data:' + type + ';base64,' + Utilities.base64Encode(bytes);
    } catch (e) {}
  }
  return '';
}

function absUrl_(href, base) {
  if (/^https?:\/\//i.test(href)) return href;
  if (/^\/\//.test(href)) return 'https:' + href;
  if (/^\//.test(href)) return base.replace(/\/$/, '') + href;
  return base.replace(/\/$/, '') + '/' + href;
}

function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function ensureSheet_(ss, name, cols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1B2A5E').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  // Plain text everywhere, so long reference numbers and leading zeros are kept exactly
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  return sh;
}

function sheet_(name) { return getSpreadsheet_().getSheetByName(name); }

function readTable_(name) {
  if (_cache[name]) return _cache[name];
  var sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return (_cache[name] = []);
  var v = sh.getDataRange().getValues(), h = v[0];
  _cache[name] = v.slice(1).map(function (row) {
    var o = {};
    h.forEach(function (k, i) {
      var x = row[i];
      o[k] = x instanceof Date ? Utilities.formatDate(x, CONFIG.TZ, 'yyyy-MM-dd') : (x === null || x === undefined ? '' : x);
    });
    return o;
  });
  return _cache[name];
}

function appendRow_(name, obj) {
  var sh = sheet_(name), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = h.map(function (k) { return obj[k] === undefined || obj[k] === null ? '' : obj[k]; });
  // Written as plain text so the sheet never reformats dates or long numbers
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row.map(String)]);
  if (_cache[name]) _cache[name].push(obj);
}

function updateRow_(name, id, patch) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), h = v[0];
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][0]) === String(id)) {
      Object.keys(patch).forEach(function (k) {
        var c = h.indexOf(k);
        if (c >= 0) sh.getRange(r + 1, c + 1).setNumberFormat('@').setValue(patch[k] === null || patch[k] === undefined ? '' : String(patch[k]));
      });
      break;
    }
  }
  delete _cache[name];
}

function getProviders_() {
  return readTable_('Providers').filter(function (p) { return String(p.active) !== 'false'; });
}

function providerById_(id) { return getProviders_().filter(function (p) { return p.id === id; })[0]; }

var NEAR_DAYS = 25;
function nearDate_(a, b) { return a && b && Math.abs(daysBetween_(a, b)) <= NEAR_DAYS; }
function isUnreadable_(b) { return b.status === 'review' && !(Number(b.amount) > 0) && /^Couldn't read/.test(b.reason || ''); }
function hasBillNear_(providerId, iso) {
  return readTable_('Bills').some(function (b) {
    return b.provider === providerId && Number(b.amount) > 0 && b.status !== 'ignored' &&
      (nearDate_(b.issueDate, iso) || nearDate_(b.dueDate, iso) || nearDate_(String(b.createdAt || '').slice(0, 10), iso));
  });
}
function clearUnreadable_(providerId, iso) {
  readTable_('Bills').filter(function (b) {
    return b.provider === providerId && isUnreadable_(b) && nearDate_(String(b.createdAt || '').slice(0, 10) || iso, iso);
  }).forEach(function (b) { updateRow_('Bills', b.id, { status: 'ignored', reason: 'Reminder for a bill already saved' }); });
}

function merge_(a, b) { var o = {}; [a, b].forEach(function (x) { for (var k in x) o[k] = x[k]; }); return o; }
function iso_(d) { return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd'); }
function todayIso_() { return iso_(new Date()); }
function daysBetween_(a, b) { return Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 864e5); }
function money_(n) { return '$' + Number(n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function niceDate_(s) { return s ? Utilities.formatDate(new Date(s + 'T00:00:00'), CONFIG.TZ, 'EEE d MMM') : '—'; }
function stripHtml_(h) { return String(h || '').replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<br\s*\/?>|<\/p>|<\/div>|<\/tr>/gi, '\n').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#36;|&dollar;/g, '$'); }

  function getSpreadsheet_() {
    var props = PropertiesService.getScriptProperties(), id = props.getProperty('bb_sheet');
    if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
    var it = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
    var ss = it.hasNext() ? SpreadsheetApp.openById(it.next().getId()) : SpreadsheetApp.create(CONFIG.SHEET_NAME);
    props.setProperty('bb_sheet', ss.getId());
    return ss;
  }
  /* Makes the bills sheet with the right columns and the common Australian providers. */
  function createSheet() {
    var ss = getSpreadsheet_();
    ensureSheet_(ss, 'Bills', BILL_COLS); ensureSheet_(ss, 'Payments', PAYMENT_COLS); ensureSheet_(ss, 'Log', LOG_COLS);
    var p = ensureSheet_(ss, 'Providers', PROVIDER_COLS);
    if (p.getLastRow() < 2) { var rows = DEFAULT_PROVIDERS.map(function (x) { return x.concat(['true', '']); }); p.getRange(2, 1, rows.length, PROVIDER_COLS.length).setValues(rows); }
    var def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
    return ss;
  }
  function searchQuery_() {
    var doms = {}; getProviders_().forEach(function (p) { if (p.domain) doms[String(p.domain).replace(/^www\./, '')] = 1; });
    var from = Object.keys(doms).slice(0, 40).join(' OR ');
    return '(label:' + CONFIG.LABEL + (from ? ' OR from:(' + from + ')' : '') + ' OR (subject:(bill OR invoice OR statement OR "account summary" OR "amount due") -category:promotions -category:social))';
  }
  function notify_(title, message) { try { hbBroadcast_(title, message); } catch (e) {} }
  function scan() {
    createSheet();
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('bs_started')) { props.setProperty('bs_started', '1'); props.setProperty('BACKFILL_LEFT', '1'); try { refreshLogos(); } catch (e) {} }
    var backfill = !!props.getProperty('BACKFILL_LEFT');
    return scanBills_(backfill ? CONFIG.BACKFILL_DAYS : CONFIG.SCAN_DAYS, backfill);
  }
  return { scan: scan, createSheet: createSheet, refreshLogos: refreshLogos };
})();

/* Run this from the editor to read your bills straight away (it also runs by itself every 15 minutes). */
function scanBillsNow() { const r = BS.scan(); Logger.log('New bills found: ' + r.added + (r.more ? ' (more to read on the next runs)' : '')); }
