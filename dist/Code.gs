/**
 * Scheduling App Prototype – session booking app for a solo practitioner (Google Apps Script, bound to a Google Sheet)
 *
 * Sheets:   Settings (edit values here), Patients, Bookings, Log
 * Calendar: a dedicated bookings calendar (created by setup) + any calendars listed in busyCalendarIds
 * Pages:    <app url>               patient booking
 *           <app url>?page=admin    therapist admin (PIN)
 *           <app url>?m=<token>     patient manage booking (cancel / reschedule)
 *           <app url>?pay=<token>   UPI payment page
 *           <app url>?page=privacy  privacy notice
 *
 * First run: Extensions ▸ Apps Script ▸ run setup()  (or Sheet menu ▸ Booking app ▸ Set up / repair)
 */

const TZ = 'Asia/Kolkata';
const TZ_OFFSET_MIN = 330; // IST has no daylight saving
const SH = { SETTINGS: 'Settings', DASH: 'Dashboard', PATIENTS: 'Patients', BOOKINGS: 'Bookings', LOG: 'Log' };

const P_COLS = ['id', 'firstName', 'lastName', 'city', 'state', 'country', 'phone', 'email',
  'privacyConsentAt', 'createdAt', 'updatedAt'];
const B_COLS = ['id', 'when', 'status', 'firstName', 'lastName', 'phone', 'email', 'city', 'type',
  'payStatus', 'payMethod', 'paidAt', 'start', 'end', 'patientId', 'token', 'eventId', 'meetLink',
  'fee', 'rem120', 'rem30', 'cancelReason', 'changeLog', 'createdAt', 'updatedAt'];

const STATUS = { CONFIRMED: 'Confirmed', CANCELLED: 'Cancelled', COMPLETED: 'Completed', NOSHOW: 'No-show' };
const PAY = { NONE: '', DUE: 'Due', REPORTED: 'Patient says paid', PAID: 'Paid' };

const DEFAULT_SETTINGS = [
  ['practiceName', 'My Practice', 'Shown on the booking page and in emails'],
  ['therapistName', 'Your therapist', 'Used in messages to patients'],
  ['therapistEmail', '', 'Where new-booking, cancellation and reminder emails for the therapist go'],
  ['therapistPhone', '', 'With country code, e.g. +91 98xxxxxxxx. Shown to patients for same-day changes'],
  ['clinicAddress', '', 'Shown for in-person sessions'],
  ['fee', '2000', 'Session fee in rupees'],
  ['inPersonCities', '', 'Cities that default to In-person (comma separated, e.g. pune, mumbai). Others default to Online'],
  ['workDays', 'Mon, Tue, Wed, Thu, Fri, Sat', 'Days patients can book'],
  ['dayStart', '10:00', 'First possible start time (24h)'],
  ['dayEnd', '18:00', 'Sessions must start early enough to finish by this time (24h)'],
  ['sessionMinutes', '60', 'Length booked on the calendar'],
  ['bufferMinutes', '120', 'Kept free before and after each session, since sessions can run long'],
  ['slotStepMinutes', '60', 'Gap between offered start times'],
  ['maxPerDay', '2', 'Most sessions bookable on one day'],
  ['minNoticeHours', '24', 'Earliest a patient can book ahead'],
  ['horizonDays', '21', 'How far ahead patients can book'],
  ['cancelCutoffHours', '24', 'Patients can cancel/reschedule online until this many hours before'],
  ['upiId', '', 'UPI ID for payments, e.g. name@okhdfcbank'],
  ['upiName', 'My Practice', 'Payee name shown in UPI apps'],
  ['slotNotice', 'A one hour slot will be booked but actual time may vary based on conversation with the therapist.', 'Message above the time picker'],
  ['policyText', 'Cancellations are ok as long as they\'re 24hrs in advance. If you need to cancel on the day of the appointment, please get in touch with the therapist. The same applies to rescheduling.', 'Added to confirmation emails'],
  ['privacyText', 'DRAFT – replace with lawyer-reviewed text. We collect your name, city, state, country, phone number and email only to book and manage your sessions, send you reminders and request payment. Only the therapist can see them. We do not share them with anyone else. You can ask us to correct or delete your details at any time by writing to the therapist.', 'Privacy notice patients agree to'],
  ['calendarId', '', 'Set by setup: calendar where bookings are written'],
  ['busyCalendarIds', '', 'Extra calendars whose events block slots (comma separated). Use "primary" for the account\'s main calendar'],
  ['appUrl', '', 'Optional: the web app /exec URL, if links in emails should use a specific deployment'],
  ['adminPin', '', 'Set by setup: PIN for the therapist page'],
];

/* ============================== Setup & menu ============================== */

function ownerOnly_() {
  // Web visitors can call any public function via google.script.run; only the owner (editor/menu) may run these.
  const active = Session.getActiveUser().getEmail();
  if (!active || active !== Session.getEffectiveUser().getEmail()) throw new Error('Not allowed.');
}

function setup() {
  ownerOnly_();
  let book = null;
  try { book = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* standalone script */ }
  const savedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!book && savedId) book = SpreadsheetApp.openById(savedId);
  if (!book) {
    book = SpreadsheetApp.create('Bookings');
    Logger.log('Created the bookings Sheet: ' + book.getUrl());
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', book.getId());

  // Settings
  let st = book.getSheetByName(SH.SETTINGS);
  if (!st) st = book.insertSheet(SH.SETTINGS, 0);
  const existing = {};
  st.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) existing[r[0]] = r[1]; });
  if (st.getLastRow() === 0) st.appendRow(['Setting', 'Value', 'What it does']);
  DEFAULT_SETTINGS.forEach(r => { if (!(r[0] in existing)) st.appendRow(r); });
  st.getRange('B:B').setNumberFormat('@');
  st.setFrozenRows(1); st.setColumnWidth(1, 170); st.setColumnWidth(2, 320); st.setColumnWidth(3, 420);
  st.getRange(1, 1, 1, 3).setFontWeight('bold');

  const setVal = (key, val) => {
    const vals = st.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) if (vals[i][0] === key) { st.getRange(i + 1, 2).setValue(val); return; }
  };
  const cur = {};
  st.getDataRange().getDisplayValues().slice(1).forEach(r => { cur[r[0]] = r[1]; });

  if (!cur.therapistEmail) setVal('therapistEmail', Session.getEffectiveUser().getEmail());
  if (!cur.adminPin) setVal('adminPin', String(Math.floor(100000 + Math.random() * 900000)));
  if (!cur.calendarId) {
    const cal = Calendar.Calendars.insert({ summary: cur.practiceName + ' – Bookings', timeZone: TZ });
    setVal('calendarId', cal.id);
  }

  // Data sheets
  [[SH.PATIENTS, P_COLS], [SH.BOOKINGS, B_COLS], [SH.LOG, ['time', 'what', 'detail']]].forEach(([name, cols]) => {
    let sh = book.getSheetByName(name);
    if (!sh) sh = book.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(cols);
      sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  });
  const s1 = book.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && book.getSheets().length > 1) book.deleteSheet(s1);

  // Dashboard tab + trigger for its month picker
  if (!book.getSheetByName(SH.DASH)) book.insertSheet(SH.DASH, 1);
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'onDashboardEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onDashboardEdit').forSpreadsheet(book).onEdit().create();

  // Reminder trigger every 5 minutes (one only)
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'reminderTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('reminderTick').timeBased().everyMinutes(5).create();

  try { refreshDashboard_(book); } catch (e) { log_('dashboard refresh failed', String(e)); }
  log_('setup', 'ok');
  try { SpreadsheetApp.getUi().alert('Booking app is set up. Next: Deploy ▸ New deployment ▸ Web app.'); } catch (e) { /* run from editor */ }
}

function showLinks() {
  ownerOnly_();
  const c = cfg_();
  const url = appUrl_(c);
  const msg = 'Admin PIN: ' + c.adminPin + '\n\nBooking page: ' + (url || '(deploy first)') +
    '\nAdmin page: ' + (url ? url + '?page=admin' : '(deploy first)');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Booking app')
    .addItem('Set up / repair', 'setup')
    .addItem('Show admin PIN and links', 'showLinks')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addToUi();
}

/* ============================== Web entry ============================== */

function doGet(e) {
  try {
    return renderPage_(e);
  } catch (err) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;max-width:480px;margin:40px auto;padding:20px">' +
      '<h2>Booking page isn\'t ready yet</h2><p>' + esc_(err.message) + '</p></div>')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function renderPage_(e) {
  const p = (e && e.parameter) || {};
  let file = 'Book';
  const t = {};
  if (p.page === 'admin') file = 'Admin';
  else if (p.page === 'privacy') file = 'Privacy';
  else if (p.m) { file = 'Manage'; t.token = String(p.m); }
  else if (p.pay) { file = 'Pay'; t.token = String(p.pay); }
  const tpl = typeof PAGES !== 'undefined' ? HtmlService.createTemplate(PAGES[file]) : HtmlService.createTemplateFromFile(file);
  tpl.boot = safeJson_(Object.assign({ cfg: publicCfg_() }, t));
  const c = cfg_();
  return tpl.evaluate()
    .setTitle('Book a session · ' + c.practiceName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  if (typeof PAGES !== 'undefined') return PAGES[name];
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function safeJson_(o) {
  return JSON.stringify(o).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
}

/* ============================== Settings & sheets ============================== */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Not set up yet: choose "setup" in the function menu at the top of the editor and click Run.');
  return book;
}

function cfg_() {
  const st = ss_().getSheetByName(SH.SETTINGS);
  if (!st) throw new Error('Not set up yet: choose "setup" in the function menu at the top of the editor and click Run.');
  const rows = st.getDataRange().getDisplayValues();
  const c = {};
  DEFAULT_SETTINGS.forEach(r => { c[r[0]] = r[1]; });
  rows.slice(1).forEach(r => { if (r[0]) c[String(r[0]).trim()] = String(r[1]).trim(); });
  ['fee', 'sessionMinutes', 'bufferMinutes', 'slotStepMinutes', 'maxPerDay', 'minNoticeHours',
    'horizonDays', 'cancelCutoffHours'].forEach(k => { c[k] = Number(c[k]) || 0; });
  c.fee = c.fee || 2000;
  c.sessionMinutes = c.sessionMinutes || 60;
  c.slotStepMinutes = c.slotStepMinutes || 60;
  c.horizonDays = c.horizonDays || 21;
  if (!c.therapistEmail) c.therapistEmail = Session.getEffectiveUser().getEmail();
  return c;
}

function publicCfg_() {
  const c = cfg_();
  return {
    practiceName: c.practiceName, therapistName: c.therapistName, fee: c.fee,
    inPersonCities: c.inPersonCities.toLowerCase().split(',').map(s => s.trim()).filter(String),
    clinicAddress: c.clinicAddress, slotNotice: c.slotNotice, policyText: c.policyText,
    therapistPhone: c.therapistPhone, cancelCutoffHours: c.cancelCutoffHours,
    sessionMinutes: c.sessionMinutes, minNoticeHours: c.minNoticeHours, appUrl: appUrl_(c),
  };
}

function appUrl_(c) {
  return (c && c.appUrl) || ScriptApp.getService().getUrl();
}

function sheet_(name) { return ss_().getSheetByName(name); }

function readTable_(name, cols) {
  const sh = sheet_(name);
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    const o = { _row: i + 1 };
    cols.forEach((k, j) => { o[k] = vals[i][j] === undefined ? '' : String(vals[i][j]); });
    out.push(o);
  }
  return out;
}

/** Stop Sheets from treating values like "+91 98…" or "=…" as formulas. */
function cell_(v) {
  const s = v === undefined || v === null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function appendRow_(name, cols, obj) {
  sheet_(name).appendRow(cols.map(k => cell_(obj[k])));
}

function updateRow_(name, cols, obj) {
  obj.updatedAt = new Date().toISOString();
  sheet_(name).getRange(obj._row, 1, 1, cols.length)
    .setValues([cols.map(k => cell_(obj[k]))]);
}

function log_(what, detail) {
  try { sheet_(SH.LOG).appendRow([new Date().toISOString(), what, detail || '']); } catch (e) { /* ignore */ }
}

/* ============================== Helpers ============================== */

function fmtWhen_(iso) {
  return Utilities.formatDate(new Date(iso), TZ, "EEE d MMM yyyy, h:mm a") + ' IST';
}
function fmtTime_(iso) { return Utilities.formatDate(new Date(iso), TZ, 'h:mm a'); }
function istDayKey_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd'); }

function normPhone_(s) { return String(s || '').replace(/\D/g, ''); }
function normEmail_(s) { return String(s || '').trim().toLowerCase(); }
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, ch =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}
function token_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function shortId_() { return 'B' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase(); }

/** Pure: available start times. Times in ms (UTC). busy: [{s,e}], perDay: {'yyyy-mm-dd': count}. */
function computeSlots_(c, busy, perDay, nowMs) {
  const DAY = 86400000, off = TZ_OFFSET_MIN * 60000;
  const names = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  const work = String(c.workDays).toLowerCase().split(',').map(s => s.trim().slice(0, 3));
  const hm = s => { const p = String(s).split(':'); return (Number(p[0]) || 0) * 60 + (Number(p[1]) || 0); };
  const startMin = hm(c.dayStart), endMin = hm(c.dayEnd);
  const session = c.sessionMinutes * 60000, buf = c.bufferMinutes * 60000, step = c.slotStepMinutes * 60000;
  const earliest = nowMs + c.minNoticeHours * 3600000;
  const todayMid = Math.floor((nowMs + off) / DAY) * DAY - off;
  const days = [];
  for (let i = 0; i <= c.horizonDays; i++) {
    const mid = todayMid + i * DAY;
    const local = new Date(mid + off);
    const key = local.toISOString().slice(0, 10);
    const wd = names[local.getUTCDay()];
    const day = { date: key, weekday: wd, slots: [], full: false };
    if (work.indexOf(wd) === -1) { day.closed = true; days.push(day); continue; }
    if (c.maxPerDay > 0 && (perDay[key] || 0) >= c.maxPerDay) { day.full = true; days.push(day); continue; }
    for (let s = mid + startMin * 60000; s + session <= mid + endMin * 60000; s += step) {
      if (s < earliest) continue;
      const bs = s - buf, be = s + session + buf;
      const clash = busy.some(b => b.s < be && b.e > bs);
      if (!clash) day.slots.push(new Date(s).toISOString());
    }
    days.push(day);
  }
  return days;
}

function busyIntervals_(c, fromMs, toMs, excludeEventId) {
  const ids = [c.calendarId].concat(String(c.busyCalendarIds || '').split(','))
    .map(s => s.trim()).filter(String);
  const out = [];
  if (!ids.length) return out;
  // Free/busy respects "free" (transparent) events and private details.
  const fb = Calendar.Freebusy.query({
    timeMin: new Date(fromMs).toISOString(), timeMax: new Date(toMs).toISOString(),
    timeZone: TZ, items: ids.map(id => ({ id: id })),
  });
  let exclude = null;
  if (excludeEventId) {
    try {
      const ev = Calendar.Events.get(c.calendarId, excludeEventId);
      exclude = { s: new Date(ev.start.dateTime).getTime(), e: new Date(ev.end.dateTime).getTime() };
    } catch (e) { /* ignore */ }
  }
  Object.keys(fb.calendars || {}).forEach(id => {
    (fb.calendars[id].busy || []).forEach(b => {
      const s = new Date(b.start).getTime(), e = new Date(b.end).getTime();
      if (exclude && id === c.calendarId && s === exclude.s && e === exclude.e) return;
      out.push({ s: s, e: e });
    });
  });
  return out;
}

function perDayCounts_(excludeId) {
  const counts = {};
  readTable_(SH.BOOKINGS, B_COLS).forEach(b => {
    if (b.status !== STATUS.CONFIRMED || b.id === excludeId) return;
    const k = istDayKey_(new Date(b.start).getTime());
    counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function availability_(c, excludeBooking) {
  const now = Date.now();
  const to = now + (c.horizonDays + 2) * 86400000;
  const busy = busyIntervals_(c, now - 86400000, to, excludeBooking && excludeBooking.eventId);
  const counts = perDayCounts_(excludeBooking && excludeBooking.id);
  return computeSlots_(c, busy, counts, now);
}

/* ============================== Public API: booking ============================== */

function getSlots() {
  return availability_(cfg_(), null);
}

function getPrivacyText() {
  return cfg_().privacyText;
}

/** Returning patient: send a 6-digit code to the email on file. Response never reveals whether they exist. */
function startLogin(identifier) {
  const id = String(identifier || '').trim();
  if (!id) throw new Error('Enter your email or phone number.');
  const cache = CacheService.getScriptCache();
  const rlKey = 'rl:' + id.toLowerCase();
  const n = Number(cache.get(rlKey) || 0);
  if (n >= 5) throw new Error('Too many attempts. Please try again in 15 minutes.');
  cache.put(rlKey, String(n + 1), 900);
  const p = findPatient_(id);
  if (p) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put('otp:' + p.id, JSON.stringify({ code: code, tries: 0 }), 600);
    const c = cfg_();
    MailApp.sendEmail({
      to: p.email, name: c.practiceName, replyTo: c.therapistEmail,
      subject: 'Your code for booking with ' + c.practiceName + ': ' + code,
      htmlBody: emailShell_(c, 'Your code', '<p>Hi ' + esc_(p.firstName) + ',</p><p>Your code is</p>' +
        '<p style="font-size:28px;font-weight:700;letter-spacing:4px">' + code + '</p>' +
        '<p style="color:#666">It expires in 10 minutes. If you didn\'t ask for this, you can ignore this email.</p>'),
    });
  }
  return { ok: true };
}

function verifyLogin(identifier, code) {
  const p = findPatient_(identifier);
  const cache = CacheService.getScriptCache();
  if (!p) throw new Error('That code didn\'t match. Please check and try again.');
  const raw = cache.get('otp:' + p.id);
  if (!raw) throw new Error('This code has expired. Please ask for a new one.');
  const o = JSON.parse(raw);
  if (o.tries >= 5) throw new Error('Too many tries. Please ask for a new code.');
  if (String(code).trim() !== o.code) {
    o.tries++; cache.put('otp:' + p.id, JSON.stringify(o), 600);
    throw new Error('That code didn\'t match. Please check and try again.');
  }
  cache.remove('otp:' + p.id);
  const session = token_();
  cache.put('sess:' + session, p.id, 3600);
  return {
    session: session,
    patient: { firstName: p.firstName, lastName: p.lastName, city: p.city, state: p.state,
      country: p.country, phone: p.phone, email: p.email },
  };
}

function findPatient_(identifier) {
  const id = String(identifier || '').trim();
  const em = normEmail_(id);
  const ph = normPhone_(id);
  const all = readTable_(SH.PATIENTS, P_COLS);
  if (em.indexOf('@') > -1) return all.find(p => normEmail_(p.email) === em) || null;
  if (ph.length >= 8) {
    const last10 = ph.slice(-10);
    return all.find(p => { const q = normPhone_(p.phone); return q === ph || q.slice(-10) === last10; }) || null;
  }
  return null;
}

/**
 * payload: { session?, patient:{firstName,lastName,city,state,country,phone,email}, privacy:true,
 *            type:'Online'|'In-person', feeAck:true, start: ISO }
 */
function book(payload) {
  const c = cfg_();
  const d = payload || {};
  const pt = d.patient || {};
  ['firstName', 'lastName', 'city', 'state', 'country', 'phone', 'email'].forEach(k => {
    if (!String(pt[k] || '').trim()) throw new Error('Please fill in all required details.');
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(pt.email).trim())) throw new Error('Please enter a valid email.');
  if (normPhone_(pt.phone).length < 8) throw new Error('Please enter a valid phone number.');
  if (!d.privacy) throw new Error('Please agree to the privacy notice.');
  if (!d.feeAck) throw new Error('Please acknowledge the session fee.');
  if (d.type !== 'Online' && d.type !== 'In-person') throw new Error('Please choose online or in-person.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const startMs = new Date(d.start).getTime();
    const ok = availability_(c, null).some(day => day.slots.some(s => new Date(s).getTime() === startMs));
    if (!ok) return { ok: false, code: 'SLOT_TAKEN', message: 'Sorry, that time was just taken. Please pick another.' };

    // Patient
    const now = new Date().toISOString();
    let patient = null;
    const cache = CacheService.getScriptCache();
    const sessPid = d.session ? cache.get('sess:' + d.session) : null;
    const all = readTable_(SH.PATIENTS, P_COLS);
    if (sessPid) {
      patient = all.find(p => p.id === sessPid);
      if (patient) {
        ['firstName', 'lastName', 'city', 'state', 'country', 'phone', 'email'].forEach(k => { patient[k] = String(pt[k]).trim(); });
        updateRow_(SH.PATIENTS, P_COLS, patient);
      }
    }
    if (!patient) {
      const em = normEmail_(pt.email), ph = normPhone_(pt.phone).slice(-10);
      const exists = all.some(p => normEmail_(p.email) === em || normPhone_(p.phone).slice(-10) === ph);
      if (exists) return { ok: false, code: 'EXISTS', message: 'Looks like you\'ve booked with us before. Please use "I\'ve booked before" so we can find your details.' };
      patient = {
        id: 'P' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase(),
        firstName: String(pt.firstName).trim(), lastName: String(pt.lastName).trim(), city: String(pt.city).trim(),
        state: String(pt.state).trim(), country: String(pt.country).trim(), phone: String(pt.phone).trim(),
        email: String(pt.email).trim(), privacyConsentAt: now, createdAt: now, updatedAt: now,
      };
      appendRow_(SH.PATIENTS, P_COLS, patient);
    }

    // Booking + calendar event
    const b = {
      id: shortId_(), token: token_(), patientId: patient.id, firstName: patient.firstName, lastName: patient.lastName,
      phone: patient.phone, email: patient.email, city: patient.city, type: d.type,
      start: new Date(startMs).toISOString(), end: new Date(startMs + c.sessionMinutes * 60000).toISOString(),
      status: STATUS.CONFIRMED, payStatus: PAY.NONE, payMethod: '', paidAt: '', fee: c.fee,
      rem120: '', rem30: '', cancelReason: '', changeLog: '', createdAt: now, updatedAt: now,
    };
    b.when = fmtWhen_(b.start);
    const ev = createEvent_(c, b);
    b.eventId = ev.eventId; b.meetLink = ev.meetLink;
    appendRow_(SH.BOOKINGS, B_COLS, b);
    log_('booked', b.id);

    sendConfirmation_(c, b, 'booked');
    notifyTherapist_(c, b, 'New booking');
    safeRefresh_();
    return { ok: true, booking: publicBooking_(c, b) };
  } finally {
    lock.releaseLock();
  }
}

function publicBooking_(c, b) {
  return {
    id: b.id, firstName: b.firstName, lastName: b.lastName, email: b.email, phone: b.phone, type: b.type,
    start: b.start, end: b.end, status: b.status, meetLink: b.type === 'Online' ? b.meetLink : '',
    fee: Number(b.fee) || c.fee, clinicAddress: b.type === 'In-person' ? c.clinicAddress : '',
    manageUrl: appUrl_(c) + '?m=' + b.token,
  };
}

/* ============================== Calendar ============================== */

function createEvent_(c, b) {
  const res = {
    summary: 'Session · ' + b.firstName + ' ' + b.lastName + ' (' + b.type + ')',
    description: 'Booking ' + b.id + '\nPhone: ' + b.phone + '\nEmail: ' + b.email + '\nCity: ' + b.city +
      '\nFee: ₹' + b.fee + '\n\nAdmin: ' + appUrl_(c) + '?page=admin',
    start: { dateTime: b.start, timeZone: TZ },
    end: { dateTime: b.end, timeZone: TZ },
    visibility: 'private',
    transparency: 'opaque',
    reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 120 }, { method: 'popup', minutes: 30 }] },
  };
  if (b.type === 'In-person' && c.clinicAddress) res.location = c.clinicAddress;
  if (b.type === 'Online') {
    res.conferenceData = { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
  }
  let ev = Calendar.Events.insert(res, c.calendarId, { conferenceDataVersion: 1, sendUpdates: 'none' });
  let link = ev.hangoutLink || '';
  for (let i = 0; i < 3 && b.type === 'Online' && !link; i++) {
    Utilities.sleep(1500);
    ev = Calendar.Events.get(c.calendarId, ev.id);
    link = ev.hangoutLink || '';
  }
  if (b.type === 'Online' && link) {
    Calendar.Events.patch({ location: link }, c.calendarId, ev.id, { sendUpdates: 'none' });
  }
  return { eventId: ev.id, meetLink: link };
}

function moveEvent_(c, b) {
  if (!b.eventId) return;
  Calendar.Events.patch({ start: { dateTime: b.start, timeZone: TZ }, end: { dateTime: b.end, timeZone: TZ } },
    c.calendarId, b.eventId, { sendUpdates: 'none' });
}

function removeEvent_(c, b) {
  if (!b.eventId) return;
  try { Calendar.Events.remove(c.calendarId, b.eventId, { sendUpdates: 'none' }); } catch (e) { log_('remove event failed', b.id + ' ' + e); }
}

/* ============================== Emails ============================== */

function emailShell_(c, title, inner) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:auto;color:#1f2a27;line-height:1.5">' +
    '<div style="padding:18px 20px;background:#3f6b5c;color:#fff;border-radius:12px 12px 0 0">' +
    '<div style="font-size:13px;opacity:.85">' + esc_(c.practiceName) + '</div>' +
    '<div style="font-size:20px;font-weight:700">' + esc_(title) + '</div></div>' +
    '<div style="padding:20px;border:1px solid #e3ddd3;border-top:none;border-radius:0 0 12px 12px">' + inner + '</div></div>';
}

function btn_(href, label) {
  return '<a href="' + esc_(href) + '" style="display:inline-block;background:#3f6b5c;color:#fff;text-decoration:none;' +
    'padding:12px 18px;border-radius:10px;font-weight:700;margin:6px 6px 6px 0">' + esc_(label) + '</a>';
}

function sessionBlock_(c, b) {
  let where = '';
  if (b.type === 'Online') {
    where = b.meetLink ? 'Online · Google Meet<br>' + btn_(b.meetLink, 'Join Google Meet') +
      '<br><span style="color:#666;font-size:13px">Open the link a few minutes early. ' + esc_(c.therapistName) + ' will let you in.</span>'
      : 'Online · the meeting link will follow by email.';
  } else {
    where = 'In-person' + (c.clinicAddress ? '<br>' + esc_(c.clinicAddress) : '');
  }
  return '<table style="width:100%;border-collapse:collapse;margin:8px 0 14px">' +
    '<tr><td style="color:#666;padding:4px 0;width:90px;vertical-align:top">When</td><td style="padding:4px 0"><b>' + esc_(fmtWhen_(b.start)) + '</b></td></tr>' +
    '<tr><td style="color:#666;padding:4px 0;vertical-align:top">Where</td><td style="padding:4px 0">' + where + '</td></tr>' +
    '<tr><td style="color:#666;padding:4px 0">Fee</td><td style="padding:4px 0">₹' + Number(b.fee || c.fee).toLocaleString('en-IN') + ' · paid after the session</td></tr>' +
    '</table>';
}

function icsBlob_(c, b, cancel) {
  const f = iso => iso.replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  const loc = b.type === 'Online' ? (b.meetLink || 'Online') : (c.clinicAddress || 'In-person');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//' + c.practiceName + '//Booking//EN',
    'METHOD:' + (cancel ? 'CANCEL' : 'PUBLISH'), 'BEGIN:VEVENT', 'UID:' + b.id + '@booking',
    'DTSTAMP:' + f(new Date().toISOString()), 'DTSTART:' + f(b.start), 'DTEND:' + f(b.end),
    'SUMMARY:Session with ' + c.therapistName, 'LOCATION:' + loc.replace(/[,;]/g, ' '),
    'DESCRIPTION:Manage your booking: ' + appUrl_(c) + '?m=' + b.token,
    'STATUS:' + (cancel ? 'CANCELLED' : 'CONFIRMED'), 'END:VEVENT', 'END:VCALENDAR'];
  return Utilities.newBlob(lines.join('\r\n'), 'text/calendar', 'session.ics');
}

function sendConfirmation_(c, b, kind) {
  const titles = { booked: 'Your session is booked', rescheduled: 'Your session has moved', cancelled: 'Your session is cancelled' };
  const manage = appUrl_(c) + '?m=' + b.token;
  let inner = '<p>Hi ' + esc_(b.firstName) + ',</p>';
  if (kind === 'cancelled') {
    inner += '<p>Your session on <b>' + esc_(fmtWhen_(b.start)) + '</b> has been cancelled.</p>' +
      btn_(appUrl_(c), 'Book another time');
  } else {
    inner += '<p>' + (kind === 'rescheduled' ? 'Your session has been moved. Here are the new details:' :
      'Thank you for booking a session with ' + esc_(c.therapistName) + '.') + '</p>' + sessionBlock_(c, b) +
      '<p style="padding:12px;background:#fbf3e3;border-radius:8px;font-size:14px">' + esc_(c.slotNotice) + '</p>' +
      '<p style="font-size:14px">' + esc_(c.policyText) + '</p>' + btn_(manage, 'Reschedule or cancel');
    if (c.therapistPhone) inner += '<p style="font-size:13px;color:#666">On the day, reach ' + esc_(c.therapistName) + ' at ' + esc_(c.therapistPhone) + '.</p>';
  }
  MailApp.sendEmail({
    to: b.email, name: c.practiceName, replyTo: c.therapistEmail,
    subject: titles[kind] + ' · ' + fmtWhen_(b.start),
    htmlBody: emailShell_(c, titles[kind], inner),
    attachments: [icsBlob_(c, b, kind === 'cancelled')],
  });
}

function notifyTherapist_(c, b, what, extra) {
  const inner = '<p><b>' + esc_(what) + '</b></p>' +
    '<p>' + esc_(b.firstName + ' ' + b.lastName) + '<br>' + esc_(b.type) + ' · ' + esc_(fmtWhen_(b.start)) + '<br>' +
    'Phone: ' + esc_(b.phone) + '<br>Email: ' + esc_(b.email) + '</p>' +
    (extra ? '<p>' + extra + '</p>' : '') + btn_(appUrl_(c) + '?page=admin', 'Open bookings');
  MailApp.sendEmail({
    to: c.therapistEmail, name: 'Booking app', subject: what + ': ' + b.firstName + ' ' + b.lastName + ' · ' + fmtWhen_(b.start),
    htmlBody: emailShell_(c, what, inner),
  });
}

/* ============================== Reminders (time trigger) ============================== */

function reminderTick() {
  const c = cfg_();
  const now = Date.now();
  readTable_(SH.BOOKINGS, B_COLS).forEach(b => {
    if (b.status !== STATUS.CONFIRMED) return;
    const mins = (new Date(b.start).getTime() - now) / 60000;
    if (mins <= 0) return;
    let which = null;
    if (!b.rem30 && mins <= 32) which = 'rem30';
    else if (!b.rem120 && mins <= 122 && mins > 32) which = 'rem120';
    if (!which) return;
    const label = which === 'rem30' ? 'in 30 minutes' : 'in 2 hours';
    try {
      const inner = '<p>Hi ' + esc_(b.firstName) + ', a quick reminder that your session with ' +
        esc_(c.therapistName) + ' starts <b>' + label + '</b>.</p>' + sessionBlock_(c, b);
      MailApp.sendEmail({ to: b.email, name: c.practiceName, replyTo: c.therapistEmail,
        subject: 'Reminder: your session starts ' + label + ' (' + fmtTime_(b.start) + ' IST)',
        htmlBody: emailShell_(c, 'Your session starts ' + label, inner) });
      notifyTherapist_(c, b, 'Session ' + label, b.type === 'Online' && b.meetLink ? btn_(b.meetLink, 'Join Google Meet') : '');
      b[which] = new Date().toISOString();
      if (which === 'rem30' && !b.rem120) b.rem120 = 'skipped';
      updateRow_(SH.BOOKINGS, B_COLS, b);
    } catch (e) { log_('reminder failed', b.id + ' ' + e); }
  });
}

/* ============================== Manage booking (patient) ============================== */

function bookingByToken_(token) {
  const t = String(token || '');
  if (t.length < 20) return null;
  return readTable_(SH.BOOKINGS, B_COLS).find(b => b.token === t) || null;
}

function getBooking(token) {
  const c = cfg_();
  const b = bookingByToken_(token);
  if (!b) throw new Error('We couldn\'t find this booking.');
  const hoursLeft = (new Date(b.start).getTime() - Date.now()) / 3600000;
  return Object.assign(publicBooking_(c, b), {
    canChange: b.status === STATUS.CONFIRMED && hoursLeft >= c.cancelCutoffHours,
    isPast: hoursLeft <= 0,
  });
}

function getRescheduleSlots(token) {
  const c = cfg_();
  const b = bookingByToken_(token);
  if (!b) throw new Error('We couldn\'t find this booking.');
  return availability_(c, b);
}

function cancelBooking(token, reason) {
  const c = cfg_();
  const r = String(reason || '').trim();
  if (!r) throw new Error('Please tell us the reason.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = bookingByToken_(token);
    if (!b || b.status !== STATUS.CONFIRMED) throw new Error('This booking can\'t be changed.');
    if ((new Date(b.start).getTime() - Date.now()) / 3600000 < c.cancelCutoffHours) {
      throw new Error('It\'s less than ' + c.cancelCutoffHours + ' hours to your session. Please contact ' + c.therapistName + ' directly.');
    }
    b.status = STATUS.CANCELLED; b.cancelReason = r;
    b.changeLog = (b.changeLog ? b.changeLog + '\n' : '') + new Date().toISOString() + ' cancelled by patient: ' + r;
    removeEvent_(c, b);
    updateRow_(SH.BOOKINGS, B_COLS, b);
    log_('cancelled', b.id);
    sendConfirmation_(c, b, 'cancelled');
    notifyTherapist_(c, b, 'Cancellation request', 'Reason: ' + esc_(r));
    safeRefresh_();
    return getBooking(token);
  } finally { lock.releaseLock(); }
}

function rescheduleBooking(token, newStart, reason) {
  const c = cfg_();
  const r = String(reason || '').trim();
  if (!r) throw new Error('Please tell us the reason.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = bookingByToken_(token);
    if (!b || b.status !== STATUS.CONFIRMED) throw new Error('This booking can\'t be changed.');
    if ((new Date(b.start).getTime() - Date.now()) / 3600000 < c.cancelCutoffHours) {
      throw new Error('It\'s less than ' + c.cancelCutoffHours + ' hours to your session. Please contact ' + c.therapistName + ' directly.');
    }
    const ms = new Date(newStart).getTime();
    const ok = availability_(c, b).some(d => d.slots.some(s => new Date(s).getTime() === ms));
    if (!ok) return { ok: false, message: 'Sorry, that time was just taken. Please pick another.' };
    const old = b.start;
    b.start = new Date(ms).toISOString();
    b.end = new Date(ms + c.sessionMinutes * 60000).toISOString();
    b.when = fmtWhen_(b.start);
    b.rem120 = ''; b.rem30 = '';
    b.changeLog = (b.changeLog ? b.changeLog + '\n' : '') + new Date().toISOString() + ' moved from ' + fmtWhen_(old) + ': ' + r;
    moveEvent_(c, b);
    updateRow_(SH.BOOKINGS, B_COLS, b);
    log_('rescheduled', b.id);
    sendConfirmation_(c, b, 'rescheduled');
    notifyTherapist_(c, b, 'Reschedule request', 'Was: ' + esc_(fmtWhen_(old)) + '<br>Reason: ' + esc_(r));
    safeRefresh_();
    return { ok: true, booking: getBooking(token) };
  } finally { lock.releaseLock(); }
}

/* ============================== Payment page ============================== */

function getPayment(token) {
  const c = cfg_();
  const b = bookingByToken_(token);
  if (!b) throw new Error('We couldn\'t find this booking.');
  return {
    firstName: b.firstName, when: fmtWhen_(b.start), amount: Number(b.fee) || c.fee, payStatus: b.payStatus,
    upiId: c.upiId, upiName: c.upiName, note: 'Session ' + b.id, therapistName: c.therapistName,
  };
}

function reportPaid(token) {
  const c = cfg_();
  const b = bookingByToken_(token);
  if (!b) throw new Error('We couldn\'t find this booking.');
  if (b.payStatus === PAY.DUE) {
    b.payStatus = PAY.REPORTED;
    updateRow_(SH.BOOKINGS, B_COLS, b);
    notifyTherapist_(c, b, 'Patient says they paid', 'Please check your UPI app, then mark it Paid.');
  }
  return getPayment(token);
}

function sendPaymentRequest_(c, b, isReminder) {
  const payUrl = appUrl_(c) + '?pay=' + b.token;
  const inner = '<p>Hi ' + esc_(b.firstName) + ',</p><p>' + (isReminder ? 'A gentle reminder about the payment for' : 'Thank you for') +
    ' your session with ' + esc_(c.therapistName) + ' on ' + esc_(fmtWhen_(b.start)) + '.</p>' +
    '<p style="font-size:22px;font-weight:700;margin:8px 0">₹' + Number(b.fee || c.fee).toLocaleString('en-IN') + '</p>' +
    btn_(payUrl, 'Pay with UPI') +
    (c.upiId ? '<p style="font-size:14px;color:#444">Or pay to UPI ID <b>' + esc_(c.upiId) + '</b> from any UPI app.</p>' : '') +
    '<p style="font-size:13px;color:#666">Questions? Just reply to this email.</p>';
  MailApp.sendEmail({ to: b.email, name: c.practiceName, replyTo: c.therapistEmail,
    subject: (isReminder ? 'Reminder: ' : '') + 'Payment for your session · ₹' + (b.fee || c.fee),
    htmlBody: emailShell_(c, isReminder ? 'Payment reminder' : 'Thank you for your session', inner) });
}

/* ============================== Insights (app + Sheet dashboard) ============================== */

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
function monthKeyIST_(iso) { return iso ? istDayKey_(new Date(iso).getTime()).slice(0, 7) : ''; }
function monthLabel_(key) { const p = key.split('-'); return MONTH_NAMES[Number(p[1]) - 1] + ' ' + p[0]; }

/**
 * Per IST month:
 *  sessions     = bookings in the month that weren't cancelled (completed + no-show + still to mark/upcoming)
 *  patients     = different patients with a session that month; new = their first-ever session is that month; returning = the rest
 *  totalPatients= everyone with at least one session up to the end of that month
 *  earned       = payments marked Paid during the month (by payment date), split UPI / cash
 *  due          = unpaid amounts for that month's sessions
 */
function computeStats_(c) {
  const rows = readTable_(SH.BOOKINGS, B_COLS);
  const curKey = istDayKey_(Date.now()).slice(0, 7);
  const first = {};
  rows.forEach(b => {
    if (b.status === STATUS.CANCELLED || !b.start) return;
    const m = monthKeyIST_(b.start);
    if (!first[b.patientId] || m < first[b.patientId]) first[b.patientId] = m;
  });
  const keys = [curKey];
  rows.forEach(b => { if (b.start) keys.push(monthKeyIST_(b.start)); if (b.paidAt) keys.push(monthKeyIST_(b.paidAt)); });
  keys.sort();
  const months = [];
  let y = Number(keys[0].slice(0, 4)), m = Number(keys[0].slice(5, 7));
  const last = keys[keys.length - 1];
  for (let guard = 0; guard < 240; guard++) {
    const k = y + '-' + String(m).padStart(2, '0');
    months.push(k);
    if (k >= last) break;
    m++; if (m > 12) { m = 1; y++; }
  }
  const by = {};
  months.forEach(k => {
    by[k] = { month: k, label: monthLabel_(k), sessions: 0, completed: 0, toMark: 0, noShows: 0, cancelled: 0,
      patients: 0, newPatients: 0, returning: 0, totalPatients: 0, earned: 0, upi: 0, cash: 0, due: 0 };
  });
  const seen = {};
  rows.forEach(b => {
    const fee = Number(b.fee) || c.fee;
    const k = monthKeyIST_(b.start), s = by[k];
    if (s) {
      if (b.status === STATUS.CANCELLED) s.cancelled++;
      else {
        s.sessions++;
        if (b.status === STATUS.COMPLETED) s.completed++;
        else if (b.status === STATUS.NOSHOW) s.noShows++;
        else s.toMark++;
        (seen[k] = seen[k] || {})[b.patientId] = true;
      }
      if (b.payStatus === PAY.DUE || b.payStatus === PAY.REPORTED) s.due += fee;
    }
    if (b.payStatus === PAY.PAID && b.paidAt) {
      const p = by[monthKeyIST_(b.paidAt)];
      if (p) { p.earned += fee; if (b.payMethod === 'Cash') p.cash += fee; else p.upi += fee; }
    }
  });
  months.forEach(k => {
    const s = by[k], ids = Object.keys(seen[k] || {});
    s.patients = ids.length;
    s.newPatients = ids.filter(id => first[id] === k).length;
    s.returning = s.patients - s.newPatients;
    s.totalPatients = Object.keys(first).filter(id => first[id] <= k).length;
  });
  return { current: curKey, months: months.map(k => by[k]).reverse() }; // newest first
}

function adminStats(pin) {
  return computeStats_(checkPin_(pin));
}

/** Rebuilds the Dashboard tab in the Sheet. Keeps the month picked in B3. */
function refreshDashboard_(book) {
  const c = cfg_();
  const st = computeStats_(c);
  const ss = book || ss_();
  let sh = ss.getSheetByName(SH.DASH);
  if (!sh) sh = ss.insertSheet(SH.DASH, 1);
  // Keep the month she picked, but go back to the current month once a new month starts.
  const props = PropertiesService.getScriptProperties();
  let picked = String(sh.getRange('B3').getDisplayValue() || '');
  if (props.getProperty('DASH_MONTH_SEEN') !== st.current) picked = '';
  props.setProperty('DASH_MONTH_SEEN', st.current);
  const m = st.months.find(x => x.label === picked) || st.months.find(x => x.month === st.current) || st.months[0];

  sh.clear();
  sh.getRange('B3').clearDataValidations();
  const money = '₹#,##0';
  sh.getRange('A1').setValue('Dashboard').setFontSize(18).setFontWeight('bold');
  sh.getRange('A2').setValue('Pick a month in B3. Updates by itself whenever a booking, payment or status changes. Last updated ' +
    Utilities.formatDate(new Date(), TZ, 'd MMM yyyy, h:mm a')).setFontColor('#5f6b67');
  sh.getRange('A3').setValue('Month').setFontWeight('bold');
  sh.getRange('B3').setValue(m.label).setFontWeight('bold').setBackground('#e6efeb')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(st.months.map(x => x.label), true).setAllowInvalid(false).build());

  const kpi = [
    ['Total patients (up to this month)', m.totalPatients],
    ['New patients', m.newPatients],
    ['Returning patients', m.returning],
    ['Earned this month', m.earned],
  ];
  sh.getRange(5, 1, kpi.length, 2).setValues(kpi);
  sh.getRange(5, 2, kpi.length, 1).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(8, 2).setNumberFormat(money);

  const more = [
    ['Sessions (not cancelled)', m.sessions],
    ['Completed', m.completed],
    ['No-shows', m.noShows],
    ['Not yet marked / upcoming', m.toMark],
    ['Cancelled', m.cancelled],
    ['Earned by UPI', m.upi],
    ['Earned in cash', m.cash],
    ['Payments due', m.due],
  ];
  sh.getRange(10, 1, more.length, 2).setValues(more);
  sh.getRange(10, 2, more.length, 1).setHorizontalAlignment('left');
  sh.getRange(15, 2, 3, 1).setNumberFormat(money);

  const head = ['Month', 'Sessions', 'Completed', 'No-shows', 'Cancelled', 'Patients', 'New', 'Returning', 'Total patients', 'Earned', 'UPI', 'Cash', 'Due'];
  sh.getRange(19, 1).setValue('Month by month').setFontSize(13).setFontWeight('bold');
  const table = [head].concat(st.months.map(x => [x.label, x.sessions, x.completed, x.noShows, x.cancelled, x.patients,
    x.newPatients, x.returning, x.totalPatients, x.earned, x.upi, x.cash, x.due]));
  sh.getRange(20, 1, table.length, head.length).setValues(table);
  sh.getRange(20, 1, 1, head.length).setFontWeight('bold').setBackground('#e6efeb');
  if (table.length > 1) sh.getRange(21, 10, table.length - 1, 4).setNumberFormat(money);
  sh.setColumnWidth(1, 250);
  sh.setFrozenRows(0);
  return m;
}

function refreshDashboard() { refreshDashboard_(); }

function safeRefresh_() {
  try { refreshDashboard_(); } catch (e) { log_('dashboard refresh failed', String(e)); }
}

/** Installable on-edit trigger (set up by setup): month changed in Dashboard!B3. */
function onDashboardEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === SH.DASH && e.range.getA1Notation() === 'B3') refreshDashboard_(e.source);
  } catch (err) { log_('dashboard edit failed', String(err)); }
}

/* ============================== Admin (therapist) ============================== */

function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfail') || 0);
  if (fails >= 10) throw new Error('Too many wrong PINs. Try again in 15 minutes.');
  const c = cfg_();
  if (!c.adminPin || String(pin || '').trim() !== String(c.adminPin)) {
    cache.put('pinfail', String(fails + 1), 900);
    throw new Error('Wrong PIN.');
  }
  return c;
}

function adminLogin(pin) { checkPin_(pin); return { ok: true }; }

function adminData(pin) {
  const c = checkPin_(pin);
  const now = Date.now();
  const todayKey = istDayKey_(now);
  const monthAgo = now - 45 * 86400000;
  const rows = readTable_(SH.BOOKINGS, B_COLS).map(b => ({
    id: b.id, name: b.firstName + ' ' + b.lastName, phone: b.phone, email: b.email, type: b.type, city: b.city,
    start: b.start, end: b.end, status: b.status, payStatus: b.payStatus, payMethod: b.payMethod, paidAt: b.paidAt,
    meetLink: b.meetLink, cancelReason: b.cancelReason, changeLog: b.changeLog, fee: Number(b.fee) || c.fee,
    _t: new Date(b.start).getTime(), _day: istDayKey_(new Date(b.start).getTime()),
  }));
  rows.sort((a, b) => a._t - b._t);
  return {
    therapistName: c.therapistName, upiId: c.upiId, upiName: c.upiName,
    today: rows.filter(r => r._day === todayKey && r.status !== STATUS.CANCELLED),
    upcoming: rows.filter(r => r._day > todayKey && r.status === STATUS.CONFIRMED),
    toClose: rows.filter(r => r._day < todayKey && r.status === STATUS.CONFIRMED),
    payments: rows.filter(r => r.payStatus === PAY.DUE || r.payStatus === PAY.REPORTED),
    cancelled: rows.filter(r => r.status === STATUS.CANCELLED && r._t > monthAgo).reverse(),
    history: rows.filter(r => (r.status === STATUS.COMPLETED || r.status === STATUS.NOSHOW) && r._t > monthAgo).reverse(),
    paidThisMonth: rows.filter(r => r.payStatus === PAY.PAID && r.paidAt && r.paidAt.slice(0, 7) === new Date().toISOString().slice(0, 7))
      .reduce((s, r) => s + r.fee, 0),
  };
}

/** action: completed_upi | completed_cash | paid_upi | paid_cash | resend_upi | noshow | cancel */
function adminAction(pin, id, action, note) {
  const c = checkPin_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = readTable_(SH.BOOKINGS, B_COLS).find(x => x.id === id);
    if (!b) throw new Error('Booking not found.');
    const started = new Date(b.start).getTime() <= Date.now();
    const stamp = new Date().toISOString();
    switch (action) {
      case 'completed_upi':
      case 'completed_upi_qr':
      case 'completed_cash':
        if (!started) throw new Error('You can mark this once the session has started.');
        if (b.status !== STATUS.CONFIRMED && b.status !== STATUS.NOSHOW) throw new Error('This session is already ' + b.status + '.');
        b.status = STATUS.COMPLETED;
        if (action === 'completed_cash') { b.payStatus = PAY.PAID; b.payMethod = 'Cash'; b.paidAt = stamp; }
        else {
          b.payStatus = PAY.DUE; b.payMethod = 'UPI';
          if (action === 'completed_upi') sendPaymentRequest_(c, b, false);
        }
        break;
      case 'paid_upi':
      case 'paid_cash':
        b.payStatus = PAY.PAID; b.payMethod = action === 'paid_cash' ? 'Cash' : 'UPI'; b.paidAt = stamp;
        break;
      case 'resend_upi':
        if (b.payStatus === PAY.PAID) throw new Error('Already paid.');
        sendPaymentRequest_(c, b, /payment request sent|completed_upi(?!_qr)/.test(b.changeLog || ''));
        b.changeLog = (b.changeLog ? b.changeLog + '\n' : '') + stamp + ' payment reminder sent';
        break;
      case 'noshow':
        if (!started) throw new Error('You can mark a no-show once the session time has passed.');
        b.status = STATUS.NOSHOW;
        break;
      case 'cancel':
        if (b.status !== STATUS.CONFIRMED) throw new Error('Only upcoming sessions can be cancelled.');
        b.status = STATUS.CANCELLED; b.cancelReason = 'By therapist' + (note ? ': ' + note : '');
        removeEvent_(c, b);
        sendConfirmation_(c, b, 'cancelled');
        break;
      default: throw new Error('Unknown action.');
    }
    b.changeLog = (b.changeLog ? b.changeLog + '\n' : '') + stamp + ' ' + action + (note ? ': ' + note : '');
    updateRow_(SH.BOOKINGS, B_COLS, b);
    log_('admin ' + action, b.id);
    safeRefresh_();
    return adminData(pin);
  } finally { lock.releaseLock(); }
}

/** Block time on the bookings calendar. date 'yyyy-mm-dd', from/to 'HH:mm' (both empty = whole day). */
function adminBlock(pin, date, from, to, label) {
  const c = checkPin_(pin);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error('Pick a date.');
  const res = { summary: 'Blocked' + (label ? ' · ' + label : ''), transparency: 'opaque', visibility: 'private' };
  if (!from && !to) {
    const next = new Date(new Date(date + 'T00:00:00Z').getTime() + 86400000).toISOString().slice(0, 10);
    res.start = { date: date }; res.end = { date: next };
  } else {
    if (!from || !to || to <= from) throw new Error('Pick a start and end time.');
    res.start = { dateTime: date + 'T' + from + ':00+05:30', timeZone: TZ };
    res.end = { dateTime: date + 'T' + to + ':00+05:30', timeZone: TZ };
  }
  const ev = Calendar.Events.insert(res, c.calendarId);
  log_('blocked', date + ' ' + (from || 'all day') + '-' + (to || ''));
  return { ok: true, id: ev.id };
}

function adminBlocks(pin) {
  const c = checkPin_(pin);
  const now = new Date();
  const items = Calendar.Events.list(c.calendarId, {
    timeMin: now.toISOString(), timeMax: new Date(now.getTime() + 60 * 86400000).toISOString(),
    singleEvents: true, orderBy: 'startTime', q: 'Blocked', maxResults: 50,
  }).items || [];
  return items.filter(e => /^Blocked/.test(e.summary || '')).map(e => ({
    id: e.id, summary: e.summary,
    start: e.start.dateTime || e.start.date, end: e.end.dateTime || e.end.date, allDay: !e.start.dateTime,
  }));
}

function adminUnblock(pin, eventId) {
  const c = checkPin_(pin);
  const e = Calendar.Events.get(c.calendarId, eventId);
  if (!/^Blocked/.test(e.summary || '')) throw new Error('That is not a blocked time.');
  Calendar.Events.remove(c.calendarId, eventId);
  return adminBlocks(pin);
}



/* ============================== Pages (bundled) ============================== */
const PAGES = {
  Book: "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <?!= include('Styles') ?>\n</head>\n<body>\n<div class=\"wrap\">\n  <div style=\"margin:6px 0 14px\">\n    <div class=\"hint\" style=\"margin:0\" id=\"brand\"></div>\n    <h1 id=\"title\">Book a session</h1>\n  </div>\n  <div class=\"progress hidden\" id=\"progress\"><i></i><i></i><i></i></div>\n\n  <!-- START -->\n  <section id=\"sStart\">\n    <p class=\"sub\">Choose a time that suits you. You'll get a confirmation by email, and reminders before your session.</p>\n    <button class=\"btn primary\" id=\"goNew\">Book my first session</button>\n    <button class=\"btn ghost\" id=\"goBack\">I've booked before</button>\n  </section>\n\n  <!-- RETURNING LOGIN -->\n  <section id=\"sLogin\" class=\"hidden\">\n    <div class=\"card\">\n      <h2>Welcome back</h2>\n      <p class=\"sub\">Enter the email or phone number you used last time. We'll email you a 6-digit code.</p>\n      <div id=\"loginStep1\">\n        <label class=\"f\" for=\"ident\">Email or phone number</label>\n        <input id=\"ident\" autocomplete=\"email\" placeholder=\"you@example.com or 98765 43210\">\n        <div class=\"err\" id=\"identErr\"></div>\n      </div>\n      <div id=\"loginStep2\" class=\"hidden\">\n        <p class=\"hint\" id=\"codeSent\"></p>\n        <label class=\"f\" for=\"code\">6-digit code</label>\n        <input id=\"code\" inputmode=\"numeric\" autocomplete=\"one-time-code\" maxlength=\"6\" placeholder=\"\u2022\u2022\u2022\u2022\u2022\u2022\" style=\"letter-spacing:6px;font-size:22px;text-align:center\">\n        <div class=\"err\" id=\"codeErr\"></div>\n        <p class=\"hint\">No email? Check spam, or <button class=\"link\" id=\"resend\">send a new code</button>.</p>\n      </div>\n    </div>\n  </section>\n\n  <!-- 1. DETAILS -->\n  <section id=\"s1\" class=\"hidden\">\n    <div class=\"card hidden\" id=\"welcomeCard\">\n      <h2 id=\"welcomeName\"></h2>\n      <p class=\"sub\" style=\"margin:0\" id=\"welcomeLine\"></p>\n      <button class=\"link\" id=\"editDetails\" style=\"margin-top:8px\">Edit my details</button>\n    </div>\n    <div class=\"card\" id=\"detailsCard\">\n      <h2>Your details</h2>\n      <p class=\"sub\">So we can confirm your booking and send reminders.</p>\n      <div class=\"row\">\n        <div><label class=\"f\" for=\"first\">First name <span class=\"req\">*</span></label><input id=\"first\" autocomplete=\"given-name\"><div class=\"err\"></div></div>\n        <div><label class=\"f\" for=\"last\">Last name <span class=\"req\">*</span></label><input id=\"last\" autocomplete=\"family-name\"><div class=\"err\"></div></div>\n      </div>\n      <label class=\"f\" for=\"country\">Country <span class=\"req\">*</span></label>\n      <select id=\"country\" autocomplete=\"country-name\"></select><div class=\"err\"></div>\n      <label class=\"f\" for=\"stateSel\">State <span class=\"req\">*</span></label>\n      <select id=\"stateSel\"></select><input id=\"stateTxt\" class=\"hidden\" autocomplete=\"address-level1\"><div class=\"err\"></div>\n      <label class=\"f\" for=\"city\">City <span class=\"req\">*</span></label>\n      <input id=\"city\" list=\"cities\" autocomplete=\"address-level2\"><datalist id=\"cities\"></datalist><div class=\"err\"></div>\n      <label class=\"f\" for=\"phone\">Mobile number <span class=\"req\">*</span></label>\n      <div class=\"row\"><span class=\"cc\" id=\"cc\">+91</span><input id=\"phone\" type=\"tel\" inputmode=\"numeric\" autocomplete=\"tel-national\" placeholder=\"98765 43210\" style=\"flex:1\"></div>\n      <div class=\"err\"></div>\n      <label class=\"f\" for=\"email\">Email <span class=\"req\">*</span></label>\n      <input id=\"email\" type=\"email\" inputmode=\"email\" autocomplete=\"email\" placeholder=\"you@example.com\"><div class=\"err\"></div>\n      <label class=\"check\" id=\"privacyWrap\"><input type=\"checkbox\" id=\"privacy\">\n        <span>I agree to the <a href=\"#\" id=\"privacyLink\" target=\"_blank\">privacy notice</a> and to my details being used to book and manage my sessions. <span class=\"req\">*</span></span></label>\n      <div class=\"err\" id=\"privacyErr\"></div>\n    </div>\n  </section>\n\n  <!-- 2. SESSION TYPE + FEE -->\n  <section id=\"s2\" class=\"hidden\">\n    <h2>How would you like to meet?</h2>\n    <p class=\"sub\" id=\"typeHint\"></p>\n    <div class=\"choice\" data-type=\"In-person\"><div class=\"ic\">\ud83d\udccd</div><div><b>In person</b><small id=\"clinicLine\">At the clinic</small></div></div>\n    <div class=\"choice\" data-type=\"Online\"><div class=\"ic\">\ud83d\udcbb</div><div><b>Online</b><small>Google Meet video call. The link comes by email</small></div></div>\n    <div class=\"err\" id=\"typeErr\"></div>\n    <div class=\"fee\" style=\"margin-top:18px\"><div><div class=\"hint\" style=\"margin:0\">Session fee</div><div class=\"amt\" id=\"feeAmt\"></div></div><span class=\"pill\">Pay after the session</span></div>\n    <label class=\"check\"><input type=\"checkbox\" id=\"ack\"><span>I understand the session fee is <b id=\"feeAmt2\"></b>. <span class=\"req\">*</span></span></label>\n    <div class=\"err\" id=\"ackErr\"></div>\n  </section>\n\n  <!-- 3. TIME -->\n  <section id=\"s3\" class=\"hidden\">\n    <h2>Pick a time</h2>\n    <div class=\"notice\" id=\"slotNotice\"></div>\n    <div id=\"picker\"><div class=\"spinner\"></div></div>\n    <div class=\"err\" id=\"slotErr\"></div>\n  </section>\n\n  <!-- 4. DONE -->\n  <section id=\"s4\" class=\"hidden\">\n    <div class=\"center\"><div class=\"big-tick\">\u2713</div><h2>You're booked</h2><p class=\"sub\" id=\"doneSub\"></p></div>\n    <div class=\"card\"><dl class=\"kv\" id=\"doneKv\"></dl></div>\n    <div class=\"card\" style=\"font-size:14.5px\" id=\"donePolicy\"></div>\n    <a class=\"btn ghost\" id=\"manageLink\" href=\"#\">Reschedule or cancel</a>\n  </section>\n</div>\n\n<div class=\"bar hidden\" id=\"bar\"><div class=\"inner\">\n  <button class=\"btn ghost\" id=\"backBtn\">Back</button>\n  <button class=\"btn primary\" id=\"nextBtn\">Continue</button>\n</div></div>\n\n<script>const BOOT = <?!= boot ?>;</script>\n<?!= include('Common') ?>\n<script>\n  const COUNTRIES = [['India', '+91'], ['United States', '+1'], ['United Kingdom', '+44'], ['United Arab Emirates', '+971'], ['Singapore', '+65'], ['Canada', '+1'], ['Australia', '+61'], ['Other', '+']];\n  const IN_STATES = ['Andaman and Nicobar Islands', 'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chandigarh', 'Chhattisgarh', 'Dadra and Nagar Haveli and Daman and Diu', 'Delhi', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir', 'Jharkhand', 'Karnataka', 'Kerala', 'Ladakh', 'Lakshadweep', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha', 'Puducherry', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal'];\n  const CITIES = ['Ahmedabad', 'Bengaluru', 'Chandigarh', 'Chennai', 'Delhi', 'Faridabad', 'Ghaziabad', 'Greater Noida', 'Gurugram', 'Hyderabad', 'Jaipur', 'Kolkata', 'Lucknow', 'Mumbai', 'New Delhi', 'Noida', 'Pune'];\n\n  const S = { step: 'start', session: null, patient: null, type: '', pick: null, returning: false };\n\n  // ---------- init ----------\n  $('brand').textContent = C.practiceName;\n  $('privacyLink').href = C.appUrl + '?page=privacy';\n  $('feeAmt').textContent = rupees(C.fee); $('feeAmt2').textContent = rupees(C.fee);\n  $('slotNotice').textContent = C.slotNotice;\n  if (C.clinicAddress) $('clinicLine').textContent = C.clinicAddress;\n  COUNTRIES.forEach(([n]) => $('country').add(new Option(n, n)));\n  $('stateSel').add(new Option('Select state', ''));\n  IN_STATES.forEach(s => $('stateSel').add(new Option(s, s)));\n  CITIES.forEach(c => { const o = document.createElement('option'); o.value = c; $('cities').appendChild(o); });\n  $('country').onchange = () => {\n    const india = $('country').value === 'India';\n    $('stateSel').classList.toggle('hidden', !india); $('stateTxt').classList.toggle('hidden', india);\n    $('cc').textContent = COUNTRIES.find(c => c[0] === $('country').value)[1];\n  };\n\n  document.querySelectorAll('#s1 input, #s1 select').forEach(el => {\n    const clear = () => {\n      el.classList.remove('invalid');\n      let n = el.nextElementSibling; while (n && !(n.classList && n.classList.contains('err'))) n = n.nextElementSibling;\n      if (!n && el.id === 'phone') n = el.parentElement.nextElementSibling;\n      if (n && n.classList.contains('err')) n.textContent = '';\n    };\n    el.addEventListener('input', clear); el.addEventListener('change', clear);\n  });\n  $('privacy').addEventListener('change', () => { if ($('privacy').checked) $('privacyErr').textContent = ''; });\n  $('ack').addEventListener('change', () => { if ($('ack').checked) $('ackErr').textContent = ''; });\n\n  // ---------- navigation ----------\n  const ORDER = ['s1', 's2', 's3'];\n  function show(id) {\n    ['sStart', 'sLogin', 's1', 's2', 's3', 's4'].forEach(s => $(s).classList.toggle('hidden', s !== id));\n    S.step = id;\n    const idx = ORDER.indexOf(id);\n    $('progress').classList.toggle('hidden', idx < 0);\n    document.querySelectorAll('#progress i').forEach((el, i) => el.classList.toggle('on', i <= idx));\n    $('bar').classList.toggle('hidden', id === 'sStart' || id === 's4');\n    $('nextBtn').disabled = false;\n    $('nextBtn').textContent = { sLogin: $('loginStep2').classList.contains('hidden') ? 'Send code' : 'Continue', s1: 'Continue', s2: 'Choose a time', s3: 'Book this time' }[id] || 'Continue';\n    $('title').textContent = id === 's4' ? 'All set' : 'Book a session';\n    if (id === 's3') $('nextBtn').disabled = !S.pick;\n    window.scrollTo(0, 0);\n  }\n  $('goNew').onclick = () => { S.returning = false; $('welcomeCard').classList.add('hidden'); $('detailsCard').classList.remove('hidden'); show('s1'); };\n  $('goBack').onclick = () => { S.returning = true; show('sLogin'); setTimeout(() => $('ident').focus(), 50); };\n  $('backBtn').onclick = () => {\n    const back = { sLogin: 'sStart', s1: S.returning ? 'sStart' : 'sStart', s2: 's1', s3: 's2' }[S.step];\n    if (S.step === 'sLogin' && !$('loginStep2').classList.contains('hidden')) { $('loginStep2').classList.add('hidden'); $('loginStep1').classList.remove('hidden'); show('sLogin'); return; }\n    show(back || 'sStart');\n  };\n  $('nextBtn').onclick = () => ({ sLogin: loginNext, s1: detailsNext, s2: typeNext, s3: bookNow }[S.step] || (() => {}))();\n\n  // ---------- returning login ----------\n  async function sendCode() {\n    const v = $('ident').value.trim();\n    $('identErr').textContent = '';\n    if (!v || (v.indexOf('@') < 0 && digits(v).length < 8)) { $('identErr').textContent = 'Enter your email or mobile number.'; return false; }\n    busy($('nextBtn'), true, 'Sending\u2026');\n    try {\n      await call('startLogin', v);\n      $('codeSent').textContent = 'If we have your details, we\\'ve emailed a code to the address you booked with.';\n      $('loginStep1').classList.add('hidden'); $('loginStep2').classList.remove('hidden');\n      busy($('nextBtn'), false); $('nextBtn').textContent = 'Continue';\n      setTimeout(() => $('code').focus(), 50);\n      return true;\n    } catch (e) { busy($('nextBtn'), false); $('identErr').textContent = e.message; return false; }\n  }\n  async function loginNext() {\n    if ($('loginStep2').classList.contains('hidden')) return sendCode();\n    const code = digits($('code').value);\n    $('codeErr').textContent = '';\n    if (code.length !== 6) { $('codeErr').textContent = 'Enter the 6-digit code from your email.'; return; }\n    busy($('nextBtn'), true, 'Checking\u2026');\n    try {\n      const r = await call('verifyLogin', $('ident').value.trim(), code);\n      busy($('nextBtn'), false);\n      S.session = r.session; fillDetails(r.patient);\n      $('welcomeName').textContent = 'Welcome back, ' + r.patient.firstName;\n      $('welcomeLine').textContent = r.patient.email + ' \u00b7 ' + r.patient.phone;\n      $('welcomeCard').classList.remove('hidden'); $('detailsCard').classList.add('hidden');\n      $('privacy').checked = true;\n      show('s1');\n    } catch (e) { busy($('nextBtn'), false); $('codeErr').textContent = e.message; }\n  }\n  $('resend').onclick = () => { $('loginStep2').classList.add('hidden'); $('loginStep1').classList.remove('hidden'); sendCode(); };\n  $('editDetails').onclick = () => { $('detailsCard').classList.remove('hidden'); $('privacyWrap').classList.add('hidden'); };\n\n  function fillDetails(p) {\n    $('first').value = p.firstName; $('last').value = p.lastName; $('city').value = p.city; $('email').value = p.email;\n    const ctry = COUNTRIES.some(c => c[0] === p.country) ? p.country : 'Other';\n    $('country').value = ctry; $('country').onchange();\n    if (ctry === 'India') $('stateSel').value = p.state; else $('stateTxt').value = p.state;\n    const ph = String(p.phone || ''); const m = ph.match(/^(\\+\\d{1,3})\\s*(.*)$/);\n    $('phone').value = m ? m[2] : ph;\n  }\n\n  // ---------- step 1 ----------\n  function setErr(el, msg) {\n    el.classList.toggle('invalid', !!msg);\n    let n = el.nextElementSibling; while (n && !n.classList.contains('err')) n = n.nextElementSibling;\n    if (!n && el.parentElement) n = el.parentElement.parentElement.querySelector(':scope > .err') || el.parentElement.nextElementSibling;\n    if (n && n.classList && n.classList.contains('err')) n.textContent = msg || '';\n    return !msg;\n  }\n  function detailsNext() {\n    const v = id => $(id).value.trim();\n    const india = $('country').value === 'India';\n    const stEl = india ? $('stateSel') : $('stateTxt');\n    const ph = digits(v('phone'));\n    let ok = true;\n    ok = setErr($('first'), v('first') ? '' : 'Please enter your first name.') && ok;\n    ok = setErr($('last'), v('last') ? '' : 'Please enter your last name.') && ok;\n    ok = setErr(stEl, stEl.value.trim() ? '' : 'Please enter your state.') && ok;\n    ok = setErr($('city'), v('city') ? '' : 'Please enter your city.') && ok;\n    const phOk = india ? /^[6-9]\\d{9}$/.test(ph.slice(-10)) && (ph.length === 10 || (ph.length === 12 && ph.startsWith('91'))) : ph.length >= 6;\n    $('phone').classList.toggle('invalid', !phOk);\n    $('phone').parentElement.nextElementSibling.textContent = phOk ? '' : (india ? 'Enter a 10-digit mobile number.' : 'Enter a valid phone number.');\n    ok = phOk && ok;\n    ok = setErr($('email'), /^[^\\s@]+@[^\\s@]+\\.[^\\s@]{2,}$/.test(v('email')) ? '' : 'Enter a valid email.') && ok;\n    $('privacyErr').textContent = $('privacy').checked ? '' : 'Please agree to continue.'; ok = $('privacy').checked && ok;\n    if (!ok) {\n      if (S.returning) { $('detailsCard').classList.remove('hidden'); }\n      const bad = document.querySelector('#s1 .invalid'); if (bad) bad.scrollIntoView({ block: 'center' });\n      return;\n    }\n    S.patient = { firstName: v('first'), lastName: v('last'), country: $('country').value, state: stEl.value.trim(), city: v('city'),\n      phone: $('cc').textContent + ' ' + (india ? ph.slice(-10) : ph), email: v('email') };\n    const city = S.patient.city.toLowerCase().replace(/\\s+/g, ' ');\n    if (!S.type) S.type = C.inPersonCities.indexOf(city) > -1 ? 'In-person' : 'Online';\n    $('typeHint').textContent = 'We\\'ve suggested ' + (S.type === 'In-person' ? 'in person' : 'online') + ' based on your city. You can change it.';\n    paintType(); show('s2');\n  }\n\n  // ---------- step 2 ----------\n  function paintType() { document.querySelectorAll('.choice').forEach(el => el.classList.toggle('sel', el.dataset.type === S.type)); }\n  document.querySelectorAll('.choice').forEach(el => el.onclick = () => { S.type = el.dataset.type; paintType(); $('typeErr').textContent = ''; });\n  let slotsLoaded = false;\n  function typeNext() {\n    let ok = true;\n    if (!S.type) { $('typeErr').textContent = 'Please choose one.'; ok = false; }\n    $('ackErr').textContent = $('ack').checked ? '' : 'Please tick to continue.'; ok = $('ack').checked && ok;\n    if (!ok) return;\n    show('s3');\n    if (!slotsLoaded) loadSlots();\n  }\n\n  // ---------- step 3 ----------\n  async function loadSlots() {\n    $('picker').innerHTML = '<div class=\"spinner\"></div>'; S.pick = null; $('nextBtn').disabled = true;\n    try {\n      const days = await call('getSlots');\n      slotsLoaded = true;\n      SlotPicker($('picker'), days, s => { S.pick = s; $('nextBtn').disabled = !s; $('slotErr').textContent = ''; });\n    } catch (e) { $('picker').innerHTML = '<div class=\"err\">' + esc(e.message) + '</div>'; }\n  }\n  async function bookNow() {\n    if (!S.pick) return;\n    busy($('nextBtn'), true, 'Booking\u2026'); $('backBtn').disabled = true;\n    try {\n      const r = await call('book', { session: S.session, patient: S.patient, privacy: $('privacy').checked, type: S.type, feeAck: $('ack').checked, start: S.pick });\n      busy($('nextBtn'), false); $('backBtn').disabled = false;\n      if (!r.ok) {\n        if (r.code === 'EXISTS') { await sheet('<h2>Welcome back</h2><p>' + esc(r.message) + '</p>', [{ label: 'Find my details', cls: 'primary', value: 1 }]); S.returning = true; $('ident').value = S.patient.email; show('sLogin'); return; }\n        $('slotErr').textContent = r.message; slotsLoaded = false; loadSlots(); return;\n      }\n      done(r.booking);\n    } catch (e) { busy($('nextBtn'), false); $('backBtn').disabled = false; $('slotErr').textContent = e.message; }\n  }\n\n  // ---------- done ----------\n  function done(b) {\n    $('doneSub').textContent = 'We\\'ve emailed the details to ' + b.email + '.';\n    $('doneKv').innerHTML = '<dt>When</dt><dd>' + esc(fmtLong(b.start)) + '</dd>' +\n      '<dt>How</dt><dd>' + (b.type === 'Online' ? 'Online \u00b7 Google Meet' + (b.meetLink ? '<br><a href=\"' + esc(b.meetLink) + '\" target=\"_blank\">' + esc(b.meetLink.replace('https://', '')) + '</a>' : '') : 'In person' + (b.clinicAddress ? '<br><span style=\"font-weight:400\">' + esc(b.clinicAddress) + '</span>' : '')) + '</dd>' +\n      '<dt>Fee</dt><dd>' + rupees(b.fee) + ' \u00b7 after the session</dd>';\n    $('donePolicy').textContent = C.policyText;\n    $('manageLink').href = b.manageUrl;\n    show('s4');\n  }\n\n  show('sStart');\n</script>\n</body>\n</html>\n",
  Manage: "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <?!= include('Styles') ?>\n</head>\n<body>\n<div class=\"wrap\">\n  <div style=\"margin:6px 0 14px\"><div class=\"hint\" style=\"margin:0\" id=\"brand\"></div><h1>Your booking</h1></div>\n  <div id=\"main\"><div class=\"spinner\"></div></div>\n</div>\n<div class=\"bar hidden\" id=\"bar\"><div class=\"inner\">\n  <button class=\"btn ghost\" id=\"backBtn\">Back</button>\n  <button class=\"btn primary\" id=\"saveBtn\" disabled>Save new time</button>\n</div></div>\n\n<script>const BOOT = <?!= boot ?>;</script>\n<?!= include('Common') ?>\n<script>\n  $('brand').textContent = C.practiceName;\n  const TOKEN = BOOT.token;\n  const REASONS = ['Schedule conflict', 'Not feeling well', 'Travelling', 'Work came up', 'Family reasons', 'Other'];\n  let B = null, pick = null, pendingReason = '';\n\n  function statusPill(b) {\n    if (b.status === 'Cancelled') return '<span class=\"pill red\">Cancelled</span>';\n    if (b.status === 'Completed') return '<span class=\"pill grey\">Completed</span>';\n    if (b.status === 'No-show') return '<span class=\"pill grey\">Missed</span>';\n    return '<span class=\"pill\">Confirmed</span>';\n  }\n\n  function contactHtml() {\n    if (!C.therapistPhone) return '<p>Please contact ' + esc(C.therapistName) + ' directly.</p>';\n    return '<a class=\"btn primary\" href=\"' + waLink(C.therapistPhone) + '\" target=\"_blank\">Message on WhatsApp</a>' +\n      '<a class=\"btn ghost\" href=\"tel:' + esc(digits(C.therapistPhone)) + '\">Call ' + esc(C.therapistPhone) + '</a>';\n  }\n\n  function render() {\n    $('bar').classList.add('hidden');\n    const b = B;\n    let h = '<div class=\"card\"><div style=\"display:flex;justify-content:space-between;align-items:center;margin-bottom:10px\"><b>' + esc(b.firstName + ' ' + b.lastName) + '</b>' + statusPill(b) + '</div>' +\n      '<dl class=\"kv\"><dt>When</dt><dd>' + (b.status === 'Cancelled' ? '<s>' + esc(fmtLong(b.start)) + '</s>' : esc(fmtLong(b.start))) + '</dd>' +\n      '<dt>How</dt><dd>' + (b.type === 'Online' ? 'Online \u00b7 Google Meet' : 'In person') + '</dd>' +\n      (b.type === 'Online' && b.meetLink && b.status === 'Confirmed' ? '<dt>Link</dt><dd><a href=\"' + esc(b.meetLink) + '\" target=\"_blank\">Join Google Meet</a></dd>' : '') +\n      (b.clinicAddress ? '<dt>Where</dt><dd>' + esc(b.clinicAddress) + '</dd>' : '') +\n      '<dt>Fee</dt><dd>' + rupees(b.fee) + '</dd></dl></div>';\n    if (b.status === 'Confirmed' && !b.isPast) {\n      if (b.canChange) {\n        h += '<button class=\"btn ghost\" id=\"resBtn\">Reschedule</button><button class=\"btn ghost\" id=\"canBtn\" style=\"color:var(--red)\">Cancel session</button>' +\n          '<p class=\"hint center\" style=\"margin-top:14px\">' + esc(C.policyText) + '</p>';\n      } else {\n        h += '<div class=\"card\"><p style=\"margin-top:0\">Your session is less than ' + C.cancelCutoffHours + ' hours away, so it can\\'t be changed online. Please get in touch with ' + esc(C.therapistName) + ':</p>' + contactHtml() + '</div>';\n      }\n    } else if (b.status === 'Cancelled') {\n      h += '<a class=\"btn primary\" href=\"' + esc(C.appUrl) + '\">Book another time</a>';\n    }\n    $('main').innerHTML = h;\n    if ($('resBtn')) $('resBtn').onclick = () => askReason('reschedule');\n    if ($('canBtn')) $('canBtn').onclick = () => askReason('cancel');\n  }\n\n  async function askReason(kind) {\n    const html = '<h2>' + (kind === 'cancel' ? 'Cancel this session?' : 'Move this session?') + '</h2>' +\n      '<p class=\"sub\">' + esc(fmtLong(B.start)) + '</p>' +\n      '<label class=\"f\">Reason <span class=\"req\">*</span></label><select id=\"rSel\"><option value=\"\">Choose a reason</option>' + REASONS.map(r => '<option>' + r + '</option>').join('') + '</select>' +\n      '<label class=\"f\">Anything else? <span class=\"hint\" id=\"rOpt\">(optional)</span></label><textarea id=\"rTxt\" rows=\"3\"></textarea><div class=\"err\" id=\"rErr\"></div>';\n    const reason = await sheet(html, [\n      { label: kind === 'cancel' ? 'Cancel session' : 'Choose a new time', cls: kind === 'cancel' ? 'danger' : 'primary',\n        validate: bg => {\n          const s = bg.querySelector('#rSel').value, t = bg.querySelector('#rTxt').value.trim();\n          const msg = !s ? 'Please choose a reason.' : (s === 'Other' && !t ? 'Please tell us a little more.' : '');\n          bg.querySelector('#rErr').textContent = msg; return !msg;\n        },\n        collect: bg => { const s = bg.querySelector('#rSel').value, t = bg.querySelector('#rTxt').value.trim(); return s + (t ? ' \u2014 ' + t : ''); } },\n      { label: 'Keep my session', value: null },\n    ]);\n    if (!reason) return;\n    if (kind === 'cancel') {\n      $('main').innerHTML = '<div class=\"spinner\"></div>';\n      try { B = await call('cancelBooking', TOKEN, reason); render(); toast('Cancelled. We\\'ve emailed you a confirmation.'); }\n      catch (e) { render(); toast(e.message); }\n    } else { pendingReason = reason; pickNewTime(); }\n  }\n\n  async function pickNewTime() {\n    pick = null;\n    $('main').innerHTML = '<h2>Pick a new time</h2><p class=\"sub\">Your current time stays booked until you save a new one.</p><div class=\"notice\">' + esc(C.slotNotice) + '</div><div id=\"picker\"><div class=\"spinner\"></div></div><div class=\"err\" id=\"pErr\"></div>';\n    $('bar').classList.remove('hidden'); $('saveBtn').disabled = true;\n    try {\n      const days = await call('getRescheduleSlots', TOKEN);\n      SlotPicker($('picker'), days, s => { pick = s; $('saveBtn').disabled = !s; });\n    } catch (e) { $('picker').innerHTML = '<div class=\"err\">' + esc(e.message) + '</div>'; }\n  }\n  $('backBtn').onclick = render;\n  $('saveBtn').onclick = async () => {\n    if (!pick) return;\n    busy($('saveBtn'), true, 'Saving\u2026');\n    try {\n      const r = await call('rescheduleBooking', TOKEN, pick, pendingReason);\n      busy($('saveBtn'), false);\n      if (!r.ok) { $('pErr').textContent = r.message; pickNewTime(); return; }\n      B = r.booking; render(); toast('Moved. We\\'ve emailed you the new details.');\n    } catch (e) { busy($('saveBtn'), false); $('pErr').textContent = e.message; }\n  };\n\n  call('getBooking', TOKEN).then(b => { B = b; render(); })\n    .catch(e => { $('main').innerHTML = '<div class=\"card\">' + esc(e.message) + '</div>'; });\n</script>\n</body>\n</html>\n",
  Admin: "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <meta name=\"theme-color\" content=\"#3f6b5c\">\n  <?!= include('Styles') ?>\n  <script src=\"https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js\"></script>\n  <style>\n    .tiles{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}\n    .tile{background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px}\n    .tile b{display:block;font-family:Fraunces,serif;font-size:30px;line-height:1.15}\n    .tile span{font-size:13px;color:var(--muted)}\n    .tile.hero{background:var(--sage);border-color:var(--sage);color:#fff} .tile.hero span{color:rgba(255,255,255,.85)}\n    .mini{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:10px}\n    .mini div{background:#fff;border:1px solid var(--line);border-radius:12px;padding:10px}\n    .mini b{display:block;font-size:18px} .mini span{font-size:12px;color:var(--muted)}\n    .qrwrap{display:flex;justify-content:center;margin:6px 0 10px}\n    .qrwrap>div{padding:12px;background:#fff;border:1px solid var(--line);border-radius:14px}\n    .top{display:flex;justify-content:space-between;align-items:flex-start;margin:6px 0 14px}\n    .icon-btn{background:#fff;border:1px solid var(--line);border-radius:12px;width:44px;height:44px;font-size:20px;cursor:pointer}\n    .stat{display:flex;gap:10px;margin-bottom:14px}\n    .stat div{flex:1;background:#fff;border:1px solid var(--line);border-radius:14px;padding:10px 12px}\n    .stat b{font-family:Fraunces,serif;font-size:22px;display:block}\n    .stat span{font-size:12.5px;color:var(--muted)}\n    .sec{font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:var(--muted);margin:18px 2px 8px}\n    .bk{background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px;margin-bottom:10px}\n    .bk.attn{border-color:var(--amber);box-shadow:0 0 0 2px var(--amber-l)}\n    .bk .t{font-family:Fraunces,serif;font-size:19px;font-weight:600}\n    .bk .n{font-weight:700;font-size:16.5px;margin-top:2px}\n    .bk .m{color:var(--muted);font-size:14px}\n    .bk .acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}\n    .bk .acts .btn{flex:1 1 auto}\n    .bk .quick{display:flex;gap:8px;margin-top:10px}\n    .bk .quick a{flex:1;text-align:center;padding:9px;border-radius:10px;background:#f3efe8;color:var(--ink);text-decoration:none;font-weight:600;font-size:14px}\n    .empty{text-align:center;color:var(--muted);padding:28px 10px;background:#fff;border:1px dashed var(--line);border-radius:16px}\n    nav.tabs{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid var(--line);display:flex;padding:6px 4px calc(6px + env(safe-area-inset-bottom));z-index:6}\n    nav.tabs button{flex:1;background:none;border:none;font:inherit;font-size:11.5px;font-weight:600;color:var(--muted);padding:6px 2px;cursor:pointer;position:relative}\n    nav.tabs button .i{display:block;font-size:20px;line-height:1.2;margin-bottom:1px}\n    nav.tabs button.on{color:var(--sage-d)}\n    nav.tabs button .badge{position:absolute;top:2px;left:calc(50% + 6px);background:var(--red);color:#fff;border-radius:99px;font-size:10.5px;padding:0 6px;line-height:17px}\n    .pinbox{max-width:320px;margin:12vh auto 0;text-align:center}\n    .pinbox input{text-align:center;font-size:28px;letter-spacing:10px}\n    .log{white-space:pre-line;font-size:12.5px;color:var(--muted);margin-top:8px;border-top:1px dashed var(--line);padding-top:8px}\n  </style>\n</head>\n<body>\n<!-- PIN -->\n<div class=\"wrap\" id=\"lock\">\n  <div class=\"pinbox\">\n    <div class=\"hint\" id=\"brand\"></div>\n    <h1 style=\"margin-bottom:6px\">Bookings</h1>\n    <p class=\"sub\">Enter your PIN</p>\n    <input id=\"pin\" type=\"password\" inputmode=\"numeric\" autocomplete=\"current-password\" maxlength=\"8\">\n    <label class=\"check\" style=\"text-align:left\"><input type=\"checkbox\" id=\"remember\" checked><span>Keep me signed in on this phone</span></label>\n    <div class=\"err\" id=\"pinErr\"></div>\n    <button class=\"btn primary\" id=\"unlock\" style=\"margin-top:14px\">Open</button>\n  </div>\n</div>\n\n<!-- APP -->\n<div class=\"wrap hidden\" id=\"app\" style=\"padding-bottom:100px\">\n  <div class=\"top\">\n    <div><div class=\"hint\" style=\"margin:0\" id=\"hello\"></div><h1 id=\"tabTitle\">Today</h1></div>\n    <button class=\"icon-btn\" id=\"refresh\" aria-label=\"Refresh\">\u21bb</button>\n  </div>\n  <div id=\"view\"><div class=\"spinner\"></div></div>\n  <p class=\"center\" style=\"margin-top:28px\"><button class=\"link\" id=\"signout\">Sign out on this phone</button></p>\n</div>\n\n<nav class=\"tabs hidden\" id=\"tabs\">\n  <button data-t=\"today\" class=\"on\"><span class=\"i\">\u2600\ufe0e</span>Today</button>\n  <button data-t=\"upcoming\"><span class=\"i\">\ud83d\uddd3</span>Upcoming</button>\n  <button data-t=\"payments\"><span class=\"i\">\u20b9</span>Payments</button>\n  <button data-t=\"stats\"><span class=\"i\">\ud83d\udcca</span>Insights</button>\n  <button data-t=\"cancelled\"><span class=\"i\">\u2715</span>Cancelled</button>\n  <button data-t=\"block\"><span class=\"i\">\u26d4\ufe0e</span>Block</button>\n</nav>\n\n<script>const BOOT = <?!= boot ?>;</script>\n<?!= include('Common') ?>\n<script>\n  $('brand').textContent = C.practiceName;\n  let PIN = null, D = null, TAB = 'today', STATS = null, MONTH = null;\n  const store = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },\n    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { } }, del(k) { try { localStorage.removeItem(k); } catch (e) { } } };\n\n  // ---------- sign in ----------\n  async function unlock(pin, silent) {\n    busy($('unlock'), true, 'Opening\u2026');\n    try {\n      await call('adminLogin', pin);\n      PIN = pin; if ($('remember').checked) store.set('bk_pin', pin);\n      $('lock').classList.add('hidden'); $('app').classList.remove('hidden'); $('tabs').classList.remove('hidden');\n      load();\n    } catch (e) { if (!silent) $('pinErr').textContent = e.message; store.del('bk_pin'); }\n    busy($('unlock'), false);\n  }\n  $('unlock').onclick = () => unlock($('pin').value.trim());\n  $('pin').onkeydown = e => { if (e.key === 'Enter') $('unlock').click(); };\n  $('signout').onclick = () => { store.del('bk_pin'); location.reload(); };\n  const saved = store.get('bk_pin'); if (saved) unlock(saved, true);\n\n  // ---------- data ----------\n  async function load() {\n    $('view').innerHTML = '<div class=\"spinner\"></div>';\n    STATS = null;\n    try { D = await call('adminData', PIN); paint(); } catch (e) { $('view').innerHTML = '<div class=\"empty\">' + esc(e.message) + '</div>'; }\n  }\n  $('refresh').onclick = load;\n  document.querySelectorAll('#tabs button').forEach(b => b.onclick = () => {\n    TAB = b.dataset.t; document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('on', x === b)); paint(); window.scrollTo(0, 0);\n  });\n\n  function greet() { const h = +fmt(new Date().toISOString(), { hour: 'numeric', hour12: false }); return (h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening') + ', ' + (D ? D.therapistName : ''); }\n\n  function badges() {\n    const set = (t, n) => { const b = document.querySelector('#tabs [data-t=\"' + t + '\"]'); let s = b.querySelector('.badge'); if (!n) { if (s) s.remove(); return; } if (!s) { s = document.createElement('span'); s.className = 'badge'; b.appendChild(s); } s.textContent = n; };\n    set('today', D.toClose.length); set('payments', D.payments.length);\n  }\n\n  function pills(r) {\n    const st = { 'Confirmed': '', 'Completed': ' grey', 'No-show': ' red', 'Cancelled': ' red' }[r.status] || '';\n    let h = '<span class=\"pill' + st + '\">' + esc(r.status === 'Confirmed' ? (r.type === 'Online' ? 'Online' : 'In person') : r.status) + '</span> ';\n    if (r.status !== 'Confirmed') h += '<span class=\"pill grey\">' + esc(r.type === 'Online' ? 'Online' : 'In person') + '</span> ';\n    if (r.payStatus === 'Due') h += '<span class=\"pill amber\">Payment due</span>';\n    if (r.payStatus === 'Patient says paid') h += '<span class=\"pill amber\">Says paid, check UPI</span>';\n    if (r.payStatus === 'Paid') h += '<span class=\"pill\">Paid \u00b7 ' + esc(r.payMethod) + '</span>';\n    return h;\n  }\n\n  function card(r, mode) {\n    const started = new Date(r.start).getTime() <= Date.now();\n    const soon = !started && new Date(r.start).getTime() - Date.now() < 3 * 3600000;\n    let acts = '';\n    if (mode === 'payments') {\n      acts = '<button class=\"btn primary small\" data-a=\"paid_upi\">Received on UPI</button>' +\n        '<button class=\"btn ghost small\" data-a=\"qr\">Show QR</button>' +\n        (r.type === 'In-person' ? '<button class=\"btn ghost small\" data-a=\"paid_cash\">Paid in cash</button>' : '') +\n        '<button class=\"btn ghost small\" data-a=\"resend_upi\">Email pay link</button>';\n    } else if (r.status === 'Confirmed' && started) {\n      acts = '<button class=\"btn primary small\" data-a=\"done\">Session done</button><button class=\"btn ghost small\" data-a=\"noshow\">Didn\\'t come</button>';\n    } else if (r.status === 'Confirmed') {\n      if (r.type === 'Online' && r.meetLink && soon) acts += '<a class=\"btn primary small\" href=\"' + esc(r.meetLink) + '\" target=\"_blank\">Join Meet</a>';\n      acts += '<button class=\"btn ghost small\" data-a=\"cancel\">Cancel</button>';\n    }\n    const when = mode === 'upcoming' || mode === 'payments' || mode === 'cancelled' || mode === 'toClose' ? fmtShort(r.start) : fmtTime(r.start);\n    return '<div class=\"bk' + (mode === 'toClose' ? ' attn' : '') + '\" data-id=\"' + esc(r.id) + '\">' +\n      '<div style=\"display:flex;justify-content:space-between;gap:8px\"><div class=\"t\">' + esc(when) + '</div><div>' + pills(r) + '</div></div>' +\n      '<div class=\"n\">' + esc(r.name) + '</div><div class=\"m\">' + esc(r.city || '') + (r.payStatus === 'Due' || mode === 'payments' ? ' \u00b7 ' + rupees(r.fee) : '') + '</div>' +\n      (r.cancelReason ? '<div class=\"m\" style=\"margin-top:6px\"><b>Reason:</b> ' + esc(r.cancelReason) + '</div>' : '') +\n      '<div class=\"quick\"><a href=\"tel:' + esc(digits(r.phone)) + '\">Call</a><a href=\"' + esc(waLink(r.phone)) + '\" target=\"_blank\">WhatsApp</a><a href=\"mailto:' + esc(r.email) + '\">Email</a></div>' +\n      (acts ? '<div class=\"acts\">' + acts + '</div>' : '') +\n      (r.changeLog && mode === 'cancelled' ? '<div class=\"log\">' + esc(r.changeLog.split('\\n').map(l => l.replace(/^\\S+\\s/, '')).join('\\n')) + '</div>' : '') +\n      '</div>';\n  }\n\n  function list(rows, mode, empty) { return rows.length ? rows.map(r => card(r, mode)).join('') : '<div class=\"empty\">' + empty + '</div>'; }\n\n  function paint() {\n    if (!D) return;\n    $('hello').textContent = greet(); badges();\n    const titles = { today: 'Today', upcoming: 'Upcoming', payments: 'Payments', stats: 'Insights', cancelled: 'Cancelled', block: 'Block time' };\n    $('tabTitle').textContent = titles[TAB];\n    let h = '';\n    if (TAB === 'today') {\n      h += '<div class=\"stat\"><div><b>' + D.today.length + '</b><span>today</span></div><div><b>' + D.upcoming.length + '</b><span>upcoming</span></div><div><b>' + rupees(D.paidThisMonth) + '</b><span>received this month</span></div></div>';\n      if (D.toClose.length) h += '<div class=\"sec\">Needs updating</div>' + list(D.toClose, 'toClose', '');\n      h += '<div class=\"sec\">' + fmt(new Date().toISOString(), { weekday: 'long', day: 'numeric', month: 'long' }) + '</div>' + list(D.today, 'today', 'No sessions today. Enjoy the day \u273f');\n    } else if (TAB === 'upcoming') {\n      h += list(D.upcoming, 'upcoming', 'No upcoming sessions yet.');\n    } else if (TAB === 'payments') {\n      h += '<p class=\"sub\">Sessions where a UPI request went out but you haven\\'t marked the payment received yet.</p>' + list(D.payments, 'payments', 'All payments are in \ud83c\udf89');\n    } else if (TAB === 'cancelled') {\n      h += '<p class=\"sub\">Last 45 days. Patients\\' reasons are shown on each.</p>' + list(D.cancelled, 'cancelled', 'No cancellations.');\n    } else if (TAB === 'block') {\n      h += blockForm();\n    } else if (TAB === 'stats') {\n      h += '<div id=\"stats\"><div class=\"spinner\"></div></div>';\n    }\n    $('view').innerHTML = h;\n    $('view').querySelectorAll('[data-a]').forEach(b => b.onclick = () => act(b.closest('.bk').dataset.id, b.dataset.a));\n    if (TAB === 'block') wireBlock();\n    if (TAB === 'stats') showStats();\n  }\n\n  function find(id) { return [].concat(D.today, D.upcoming, D.toClose, D.payments, D.cancelled, D.history).find(r => r.id === id); }\n\n  async function act(id, a) {\n    const r = find(id);\n    let action = a, note = '';\n    if (a === 'qr') { showQr(r, false); return; }\n    if (a === 'done') {\n      const first = esc(r.name.split(' ')[0]);\n      const btns = r.type === 'In-person'\n        ? [{ label: 'Done. Show UPI QR to scan', cls: 'primary', value: 'completed_upi_qr' }, { label: 'Done. Paid in cash', cls: 'ghost', value: 'completed_cash' }]\n        : [{ label: 'Done. Send UPI request', cls: 'primary', value: 'completed_upi' }];\n      btns.push({ label: 'Not yet', value: null });\n      action = await sheet('<h2>' + esc(r.name) + '</h2><p class=\"sub\">' + esc(fmtShort(r.start)) + ' \u00b7 ' + rupees(r.fee) + '</p>' +\n        (r.type === 'In-person'\n          ? '<p>Show ' + first + ' a UPI QR code on your phone to scan, or record a cash payment.</p>'\n          : '<p>We\\'ll email ' + first + ' a UPI payment link, and show you the QR code too in case you want to share it.</p>'), btns);\n    } else if (a === 'noshow') {\n      action = await sheet('<h2>Mark as no-show?</h2><p class=\"sub\">' + esc(r.name) + ' \u00b7 ' + esc(fmtShort(r.start)) + '</p>', [{ label: 'Yes, they didn\\'t come', cls: 'danger', value: 'noshow' }, { label: 'Back', value: null }]);\n    } else if (a === 'cancel') {\n      const res = await sheet('<h2>Cancel this session?</h2><p class=\"sub\">' + esc(r.name) + ' \u00b7 ' + esc(fmtShort(r.start)) + '</p><p>We\\'ll email ' + esc(r.name.split(' ')[0]) + ' that it\\'s cancelled.</p>' +\n        '<label class=\"f\">Note for your records (optional)</label><input id=\"cNote\">',\n        [{ label: 'Cancel session', cls: 'danger', collect: bg => ({ v: 'cancel', n: bg.querySelector('#cNote').value.trim() }) }, { label: 'Keep it', value: null }]);\n      action = res && res.v; note = res && res.n;\n    } else if (a === 'resend_upi') {\n      action = await sheet('<h2>Send a payment reminder?</h2><p class=\"sub\">' + esc(r.name) + ' \u00b7 ' + rupees(r.fee) + '</p>', [{ label: 'Send reminder email', cls: 'primary', value: 'resend_upi' }, { label: 'Back', value: null }]);\n    }\n    if (!action) return;\n    try {\n      toast('Saving\u2026');\n      D = await call('adminAction', PIN, id, action, note || '');\n      hideToast();\n      STATS = null;\n      paint();\n      if (action === 'completed_upi' || action === 'completed_upi_qr') { showQr(r, action === 'completed_upi'); return; }\n      toast({ completed_upi: 'Done. Payment request sent.', completed_cash: 'Done. Cash payment recorded.', paid_upi: 'Marked as paid.', paid_cash: 'Marked as paid in cash.', resend_upi: 'Reminder sent.', noshow: 'Marked as no-show.', cancel: 'Cancelled and patient emailed.' }[action] || 'Saved');\n    } catch (e) { toast(e.message); }\n  }\n\n  // ---------- UPI QR on her screen ----------\n  async function showQr(r, emailed) {\n    if (!D.upiId) { await sheet('<h2>UPI ID missing</h2><p>Add your UPI ID in the <b>Settings</b> tab of the Google Sheet (row \"upiId\"), then try again.</p>', [{ label: 'OK', value: null }]); return; }\n    const upi = 'upi://pay?pa=' + encodeURIComponent(D.upiId).replace(/%40/g, '@') + '&pn=' + encodeURIComponent(D.upiName || D.therapistName) +\n      '&am=' + Number(r.fee).toFixed(2) + '&cu=INR&tn=' + encodeURIComponent('Session ' + r.id);\n    const p = sheet('<h2 class=\"center\">Scan to pay ' + rupees(r.fee) + '</h2><p class=\"sub center\">' + esc(r.name) + ' \u00b7 ' + esc(fmtShort(r.start)) + '</p>' +\n      '<div class=\"qrwrap\"><div id=\"qrBox\"></div></div><p class=\"center\" style=\"margin:0\">UPI ID: <b>' + esc(D.upiId) + '</b></p>' +\n      '<p class=\"hint center\">Any UPI app can scan this: GPay, PhonePe, Paytm or a bank app.' + (emailed ? ' We\\'ve also emailed the payment link.' : '') + '</p>',\n      [{ label: 'Payment received', cls: 'primary', value: 'paid_upi' }, { label: 'They\\'ll pay later', value: null }]);\n    try { new QRCode(document.getElementById('qrBox'), { text: upi, width: 230, height: 230, correctLevel: QRCode.CorrectLevel.M }); }\n    catch (e) { document.getElementById('qrBox').textContent = 'QR could not load. Share the UPI ID instead.'; }\n    const v = await p;\n    if (v === 'paid_upi') {\n      try { D = await call('adminAction', PIN, r.id, 'paid_upi', ''); STATS = null; paint(); toast('Marked as paid.'); } catch (e) { toast(e.message); }\n    } else { paint(); toast('Saved. It\\'s in Payments until it\\'s paid.'); }\n  }\n\n  // ---------- insights ----------\n  async function showStats() {\n    try {\n      if (!STATS) STATS = await call('adminStats', PIN);\n      if (!MONTH || !STATS.months.some(m => m.month === MONTH)) MONTH = STATS.current;\n      const m = STATS.months.find(x => x.month === MONTH);\n      const opts = STATS.months.map(x => '<option value=\"' + x.month + '\"' + (x.month === MONTH ? ' selected' : '') + '>' + esc(x.label) + (x.month === STATS.current ? ' (this month)' : '') + '</option>').join('');\n      $('stats').innerHTML = '<select id=\"mSel\" style=\"margin-bottom:12px\">' + opts + '</select>' +\n        '<div class=\"tiles\">' +\n        '<div class=\"tile\"><b>' + m.totalPatients + '</b><span>Total patients</span></div>' +\n        '<div class=\"tile hero\"><b>' + rupees(m.earned) + '</b><span>Earned' + (m.month === STATS.current ? ' so far' : '') + '</span></div>' +\n        '<div class=\"tile\"><b>' + m.newPatients + '</b><span>New patients</span></div>' +\n        '<div class=\"tile\"><b>' + m.returning + '</b><span>Returning patients</span></div></div>' +\n        '<div class=\"sec\">Sessions</div><div class=\"mini\">' +\n        '<div><b>' + m.completed + '</b><span>Completed</span></div><div><b>' + m.noShows + '</b><span>No-shows</span></div><div><b>' + m.cancelled + '</b><span>Cancelled</span></div>' +\n        (m.toMark ? '<div><b>' + m.toMark + '</b><span>' + (m.month >= STATS.current ? 'Upcoming / to mark' : 'Not marked yet') + '</span></div>' : '') +\n        '</div><div class=\"sec\">Money</div><div class=\"mini\">' +\n        '<div><b>' + rupees(m.upi) + '</b><span>By UPI</span></div><div><b>' + rupees(m.cash) + '</b><span>In cash</span></div><div><b>' + rupees(m.due) + '</b><span>Still due</span></div></div>' +\n        '<p class=\"hint\">Total patients: everyone with at least one session up to the end of this month. New: first-ever session this month. ' +\n        'Returning: had a session before. Earned: payments marked received this month.</p>';\n      $('mSel').onchange = () => { MONTH = $('mSel').value; showStats(); };\n    } catch (e) { $('stats').innerHTML = '<div class=\"empty\">' + esc(e.message) + '</div>'; }\n  }\n\n  // ---------- block time ----------\n  function blockForm() {\n    const today = fmt(new Date().toISOString(), { year: 'numeric', month: '2-digit', day: '2-digit' }).split('/').reverse().join('-');\n    return '<div class=\"card\"><p class=\"sub\" style=\"margin-top:0\">Blocked time can\\'t be booked. You can also just add an event in your Google Calendar app.</p>' +\n      '<label class=\"f\" for=\"bDate\">Date</label><input type=\"date\" id=\"bDate\" value=\"' + today + '\" min=\"' + today + '\">' +\n      '<label class=\"check\"><input type=\"checkbox\" id=\"bAll\" checked><span>Whole day</span></label>' +\n      '<div class=\"row hidden\" id=\"bTimes\"><div><label class=\"f\" for=\"bFrom\">From</label><input type=\"time\" id=\"bFrom\" value=\"10:00\"></div><div><label class=\"f\" for=\"bTo\">To</label><input type=\"time\" id=\"bTo\" value=\"13:00\"></div></div>' +\n      '<label class=\"f\" for=\"bLabel\">Note (only you see this)</label><input id=\"bLabel\" placeholder=\"e.g. Travel\">' +\n      '<div class=\"err\" id=\"bErr\"></div><button class=\"btn primary\" id=\"bSave\" style=\"margin-top:16px\">Block this time</button></div>' +\n      '<div class=\"sec\">Blocked in the next 60 days</div><div id=\"blocks\"><div class=\"spinner\"></div></div>';\n  }\n  async function wireBlock() {\n    $('bAll').onchange = () => $('bTimes').classList.toggle('hidden', $('bAll').checked);\n    $('bSave').onclick = async () => {\n      $('bErr').textContent = '';\n      busy($('bSave'), true, 'Blocking\u2026');\n      try {\n        await call('adminBlock', PIN, $('bDate').value, $('bAll').checked ? '' : $('bFrom').value, $('bAll').checked ? '' : $('bTo').value, $('bLabel').value.trim());\n        toast('Blocked'); $('bLabel').value = ''; showBlocks();\n      } catch (e) { $('bErr').textContent = e.message; }\n      busy($('bSave'), false);\n    };\n    showBlocks();\n  }\n  async function showBlocks(listData) {\n    try {\n      const items = listData || await call('adminBlocks', PIN);\n      $('blocks').innerHTML = items.length ? items.map(b => '<div class=\"bk\" style=\"display:flex;justify-content:space-between;align-items:center\"><div><div class=\"n\" style=\"margin:0\">' +\n        esc(b.allDay ? fmt(b.start + 'T12:00:00+05:30', { weekday: 'short', day: 'numeric', month: 'short' }) + ' \u00b7 all day' : fmtShort(b.start) + ' \u2013 ' + fmtTime(b.end)) +\n        '</div><div class=\"m\">' + esc(b.summary.replace(/^Blocked( \u00b7 )?/, '') || 'Blocked') + '</div></div><button class=\"btn ghost small\" data-ev=\"' + esc(b.id) + '\">Remove</button></div>').join('')\n        : '<div class=\"empty\">Nothing blocked.</div>';\n      $('blocks').querySelectorAll('[data-ev]').forEach(btn => btn.onclick = async () => {\n        busy(btn, true, '\u2026');\n        try { showBlocks(await call('adminUnblock', PIN, btn.dataset.ev)); toast('Removed'); } catch (e) { toast(e.message); busy(btn, false); }\n      });\n    } catch (e) { $('blocks').innerHTML = '<div class=\"empty\">' + esc(e.message) + '</div>'; }\n  }\n</script>\n</body>\n</html>\n",
  Pay: "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <?!= include('Styles') ?>\n  <script src=\"https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js\"></script>\n</head>\n<body>\n<div class=\"wrap\">\n  <div style=\"margin:6px 0 14px\"><div class=\"hint\" style=\"margin:0\" id=\"brand\"></div><h1>Session payment</h1></div>\n  <div id=\"main\"><div class=\"spinner\"></div></div>\n</div>\n<script>const BOOT = <?!= boot ?>;</script>\n<?!= include('Common') ?>\n<script>\n  $('brand').textContent = C.practiceName;\n  function render(p) {\n    if (p.payStatus === 'Paid') { $('main').innerHTML = '<div class=\"card center\"><div class=\"big-tick\">\u2713</div><h2>Paid. Thank you!</h2><p class=\"sub\">' + esc(p.when) + '</p></div>'; return; }\n    if (!p.upiId) { $('main').innerHTML = '<div class=\"card\">Payment details aren\\'t set up yet. Please contact ' + esc(p.therapistName) + '.</div>'; return; }\n    const upi = 'upi://pay?pa=' + encodeURIComponent(p.upiId).replace(/%40/g, '@') + '&pn=' + encodeURIComponent(p.upiName) + '&am=' + Number(p.amount).toFixed(2) +\n      '&cu=INR&tn=' + encodeURIComponent(p.note);\n    const reported = p.payStatus === 'Patient says paid';\n    $('main').innerHTML =\n      '<div class=\"card center\"><p class=\"sub\" style=\"margin:0\">Hi ' + esc(p.firstName) + ', for your session on</p><b>' + esc(p.when) + '</b>' +\n      '<div style=\"font-family:Fraunces,serif;font-size:38px;font-weight:600;margin:10px 0\">' + rupees(p.amount) + '</div>' +\n      '<a class=\"btn primary\" href=\"' + upi + '\" target=\"_blank\">Pay with a UPI app</a>' +\n      '<p class=\"hint\">Opens GPay, PhonePe, Paytm or your bank app on this phone.</p></div>' +\n      '<div class=\"card center\"><p style=\"margin-top:0\"><b>Paying from another phone?</b> Scan this code.</p><div id=\"qr\" style=\"display:inline-block;padding:10px;background:#fff\"></div>' +\n      '<p style=\"margin-bottom:0\">UPI ID: <b>' + esc(p.upiId) + '</b> <button class=\"link\" id=\"copy\">Copy</button></p></div>' +\n      (reported ? '<div class=\"card center\">Thanks! We\\'ve told ' + esc(p.therapistName) + ' you\\'ve paid.</div>'\n        : '<button class=\"btn ghost\" id=\"paid\">I\\'ve paid</button><p class=\"hint center\">Lets ' + esc(p.therapistName) + ' know to look for your payment.</p>');\n    try { new QRCode($('qr'), { text: upi, width: 200, height: 200, correctLevel: QRCode.CorrectLevel.M }); } catch (e) { $('qr').textContent = ''; }\n    $('copy').onclick = () => { try { navigator.clipboard.writeText(p.upiId).then(() => toast('UPI ID copied')); } catch (e) { toast(p.upiId); } };\n    if ($('paid')) $('paid').onclick = async () => { busy($('paid'), true, 'Sending\u2026'); try { render(await call('reportPaid', BOOT.token)); } catch (e) { busy($('paid'), false); toast(e.message); } };\n  }\n  call('getPayment', BOOT.token).then(render).catch(e => { $('main').innerHTML = '<div class=\"card\">' + esc(e.message) + '</div>'; });\n</script>\n</body>\n</html>\n",
  Privacy: "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta charset=\"utf-8\">\n  <?!= include('Styles') ?>\n</head>\n<body>\n<div class=\"wrap\">\n  <div style=\"margin:6px 0 14px\"><div class=\"hint\" style=\"margin:0\" id=\"brand\"></div><h1>Privacy notice</h1></div>\n  <div class=\"card\" id=\"txt\" style=\"white-space:pre-line\"></div>\n  <a class=\"btn ghost\" id=\"back\" href=\"#\">Back to booking</a>\n</div>\n<script>const BOOT = <?!= boot ?>;</script>\n<?!= include('Common') ?>\n<script>\n  $('brand').textContent = C.practiceName;\n  $('back').href = C.appUrl;\n  call('getPrivacyText').then(t => { $('txt').textContent = t; }).catch(e => { $('txt').textContent = e.message; });\n</script>\n</body>\n</html>\n",
  Styles: "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link href=\"https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=DM+Sans:wght@400;500;600;700&display=swap\" rel=\"stylesheet\">\n<style>\n  :root{--bg:#f6f3ee;--card:#fff;--ink:#1f2a27;--muted:#5f6b67;--line:#e3ddd3;--sage:#3f6b5c;--sage-d:#2f5246;--sage-l:#e6efeb;\n    --amber:#b7791f;--amber-l:#fbf3e3;--red:#b23b3b;--red-l:#fbeaea;--r:14px}\n  *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}\n  html,body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 \"DM Sans\",system-ui,-apple-system,sans-serif}\n  .wrap{max-width:560px;margin:0 auto;padding:16px 16px 120px}\n  h1,h2{font-family:Fraunces,Georgia,serif;font-weight:600;letter-spacing:-.01em;margin:0}\n  h1{font-size:26px;line-height:1.2} h2{font-size:21px;margin-bottom:4px}\n  .sub{color:var(--muted);margin:2px 0 16px;font-size:15px}\n  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);padding:18px;margin-bottom:14px}\n  .hidden{display:none!important}\n  label.f{display:block;font-weight:600;font-size:14px;margin:14px 0 6px}\n  .req{color:var(--red)}\n  input,select,textarea{width:100%;font:inherit;font-size:16px;color:inherit;background:#fff;border:1px solid #cfc8bc;border-radius:12px;padding:13px 14px;outline:none;-webkit-appearance:none;appearance:none}\n  select{background-image:linear-gradient(45deg,transparent 50%,#5f6b67 50%),linear-gradient(135deg,#5f6b67 50%,transparent 50%);background-position:calc(100% - 20px) 55%,calc(100% - 15px) 55%;background-size:5px 5px;background-repeat:no-repeat;padding-right:36px}\n  input:focus,select:focus,textarea:focus{border-color:var(--sage);box-shadow:0 0 0 3px var(--sage-l)}\n  .invalid{border-color:var(--red)!important}\n  .err{color:var(--red);font-size:14px;margin-top:6px}\n  .hint{color:var(--muted);font-size:13.5px;margin-top:6px}\n  .row{display:flex;gap:10px} .row>*{flex:1}\n  .cc{flex:0 0 76px;display:flex;align-items:center;justify-content:center;background:#f3efe8;border:1px solid #cfc8bc;border-radius:12px;font-weight:600}\n  .check{display:flex;gap:12px;align-items:flex-start;padding:14px;border:1px solid var(--line);border-radius:12px;background:#fbfaf7;margin-top:12px;font-size:15px}\n  .check input{width:22px;height:22px;flex:0 0 auto;margin-top:1px;accent-color:var(--sage);-webkit-appearance:auto;appearance:auto;padding:0}\n  .btn{display:flex;align-items:center;justify-content:center;gap:8px;width:100%;font:inherit;font-weight:700;font-size:16px;border-radius:14px;padding:15px 18px;border:1px solid transparent;cursor:pointer;text-decoration:none;min-height:52px}\n  .btn.primary{background:var(--sage);color:#fff} .btn.primary:active{background:var(--sage-d)}\n  .btn.primary:disabled{background:#a9bdb5}\n  .btn.ghost{background:#fff;border-color:#cfc8bc;color:var(--ink)}\n  .btn.danger{background:var(--red);color:#fff}\n  .btn.small{min-height:42px;padding:9px 14px;font-size:14.5px;border-radius:11px;width:auto}\n  .btn+.btn{margin-top:10px}\n  .link{background:none;border:none;color:var(--sage-d);text-decoration:underline;font:inherit;padding:0;cursor:pointer}\n  .bar{position:fixed;left:0;right:0;bottom:0;background:rgba(246,243,238,.96);backdrop-filter:blur(6px);border-top:1px solid var(--line);padding:12px 16px calc(12px + env(safe-area-inset-bottom));z-index:5}\n  .bar .inner{max-width:560px;margin:0 auto;display:flex;gap:10px}\n  .bar .inner .btn{margin:0}\n  .bar .inner .btn.ghost{flex:0 0 96px}\n  .progress{display:flex;gap:6px;margin:4px 0 18px}\n  .progress i{flex:1;height:5px;border-radius:9px;background:var(--line)} .progress i.on{background:var(--sage)}\n  .choice{display:flex;gap:12px;align-items:center;padding:16px;border:2px solid var(--line);border-radius:14px;background:#fff;cursor:pointer;margin-bottom:10px}\n  .choice.sel{border-color:var(--sage);background:var(--sage-l)}\n  .choice .ic{font-size:24px;width:34px;text-align:center}\n  .choice b{display:block} .choice small{color:var(--muted)}\n  .fee{display:flex;justify-content:space-between;align-items:center;padding:16px;border-radius:14px;background:var(--sage-l)}\n  .fee .amt{font-family:Fraunces,serif;font-size:28px;font-weight:600}\n  .pill{display:inline-block;font-size:12.5px;padding:3px 10px;border-radius:99px;background:var(--sage-l);color:var(--sage-d);font-weight:700}\n  .pill.amber{background:var(--amber-l);color:var(--amber)} .pill.red{background:var(--red-l);color:var(--red)} .pill.grey{background:#eee;color:#555}\n  .notice{border-left:4px solid var(--amber);background:var(--amber-l);padding:12px 14px;border-radius:10px;font-size:14.5px;margin-bottom:14px}\n  .days{display:flex;gap:8px;overflow-x:auto;padding:2px 2px 8px;margin:0 -2px 8px;scroll-snap-type:x proximity}\n  .days::-webkit-scrollbar{display:none}\n  .day{flex:0 0 68px;text-align:center;border:2px solid var(--line);border-radius:14px;padding:8px 2px;background:#fff;cursor:pointer;scroll-snap-align:start}\n  .day .dw{font-size:12px;color:var(--muted)} .day .dn{font-size:21px;font-family:Fraunces,serif;font-weight:600;line-height:1.2} .day .dm{font-size:11px;color:var(--muted)}\n  .day .dc{font-size:11px;color:var(--sage);font-weight:600}\n  .day.sel{border-color:var(--sage);background:var(--sage-l)}\n  .day:disabled{opacity:.4;cursor:default}\n  .slots{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n  .slot{border:2px solid #cfc8bc;background:#fff;border-radius:12px;padding:12px 6px;text-align:center;font:inherit;font-weight:600;cursor:pointer;min-height:54px}\n  .slot.sel{background:var(--sage);border-color:var(--sage);color:#fff}\n  .slot small{display:block;font-weight:400;font-size:12px;opacity:.8}\n  .kv{display:grid;grid-template-columns:96px 1fr;gap:8px 12px;font-size:15px}\n  .kv dt{color:var(--muted)} .kv dd{margin:0;font-weight:600}\n  .center{text-align:center}\n  .big-tick{width:64px;height:64px;border-radius:50%;background:var(--sage);color:#fff;display:grid;place-items:center;font-size:32px;margin:8px auto 12px}\n  .spinner{width:28px;height:28px;border:3px solid var(--sage-l);border-top-color:var(--sage);border-radius:50%;animation:sp 0.8s linear infinite;margin:24px auto}\n  @keyframes sp{to{transform:rotate(360deg)}}\n  .toast{position:fixed;left:50%;bottom:96px;transform:translateX(-50%);background:#1f2a27;color:#fff;padding:12px 16px;border-radius:12px;font-size:14.5px;z-index:20;max-width:90%;box-shadow:0 6px 20px rgba(0,0,0,.2)}\n  .sheet-bg{position:fixed;inset:0;background:rgba(20,25,23,.45);z-index:15;display:flex;align-items:flex-end;justify-content:center}\n  .sheet{background:#fff;width:100%;max-width:560px;border-radius:20px 20px 0 0;padding:20px 18px calc(20px + env(safe-area-inset-bottom));max-height:88vh;overflow:auto}\n  .sheet .grab{width:44px;height:5px;border-radius:9px;background:#ddd;margin:-6px auto 14px}\n  a{color:var(--sage-d)}\n</style>\n",
  Common: "<script>\n  const C = BOOT.cfg;\n  const $ = id => document.getElementById(id);\n  const IST = 'Asia/Kolkata';\n  const LOCAL_TZ = Intl.DateTimeFormat().resolvedOptions().timeZone;\n  const SHOW_LOCAL = LOCAL_TZ !== 'Asia/Kolkata' && LOCAL_TZ !== 'Asia/Calcutta';\n\n  function call(fn) {\n    const args = Array.prototype.slice.call(arguments, 1);\n    return new Promise((resolve, reject) => {\n      google.script.run\n        .withSuccessHandler(resolve)\n        .withFailureHandler(e => reject(new Error(String((e && e.message) || e).replace(/^(Exception|Error):\\s*/, ''))))[fn].apply(null, args);\n    });\n  }\n  function fmt(iso, opts, tz) { return new Intl.DateTimeFormat('en-IN', Object.assign({ timeZone: tz || IST }, opts)).format(new Date(iso)); }\n  function fmtTime(iso, tz) { return fmt(iso, { hour: 'numeric', minute: '2-digit' }, tz); }\n  function fmtLong(iso) { return fmt(iso, { weekday: 'long', day: 'numeric', month: 'long' }) + ', ' + fmtTime(iso) + ' IST'; }\n  function fmtShort(iso) { return fmt(iso, { weekday: 'short', day: 'numeric', month: 'short' }) + ', ' + fmtTime(iso); }\n  function esc(s) { return String(s == null ? '' : s).replace(/[&<>\"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c])); }\n  function rupees(n) { return '\u20b9' + Number(n).toLocaleString('en-IN'); }\n  function digits(s) { return String(s || '').replace(/\\D/g, ''); }\n  function waLink(phone) { const d = digits(phone); return d ? 'https://wa.me/' + (d.length === 10 ? '91' + d : d) : ''; }\n\n  let toastTimer;\n  function toast(msg) {\n    let t = document.querySelector('.toast');\n    if (!t) { t = document.createElement('div'); t.className = 'toast'; document.body.appendChild(t); }\n    t.textContent = msg; t.classList.remove('hidden');\n    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add('hidden'), 3500);\n  }\n  function hideToast() { const t = document.querySelector('.toast'); if (t) t.classList.add('hidden'); }\n  function busy(el, on, label) {\n    if (!el) return;\n    if (on) { el.dataset.label = el.textContent; el.textContent = label || 'Please wait\u2026'; el.disabled = true; }\n    else { el.textContent = el.dataset.label || el.textContent; el.disabled = false; }\n  }\n\n  /** Bottom sheet. Returns a promise resolved with the clicked button's value (or null). */\n  function sheet(html, buttons) {\n    return new Promise(resolve => {\n      const bg = document.createElement('div'); bg.className = 'sheet-bg';\n      bg.innerHTML = '<div class=\"sheet\"><div class=\"grab\"></div>' + html + '<div class=\"sheet-btns\" style=\"margin-top:16px\"></div></div>';\n      const box = bg.querySelector('.sheet-btns');\n      (buttons || []).forEach(b => {\n        const el = document.createElement('button');\n        el.className = 'btn ' + (b.cls || 'ghost'); el.textContent = b.label;\n        el.onclick = () => {\n          if (b.validate && !b.validate(bg)) return;\n          const v = b.collect ? b.collect(bg) : b.value; document.body.removeChild(bg); resolve(v);\n        };\n        box.appendChild(el);\n      });\n      bg.onclick = e => { if (e.target === bg) { document.body.removeChild(bg); resolve(null); } };\n      document.body.appendChild(bg);\n    });\n  }\n\n  /** Slot picker. days: [{date, weekday, slots[], full, closed}] */\n  function SlotPicker(host, days, onPick) {\n    const open = days.filter(d => !d.closed);\n    let sel = Math.max(0, open.findIndex(d => d.slots.length));\n    let picked = null;\n    function render() {\n      if (!open.some(d => d.slots.length)) {\n        host.innerHTML = '<div class=\"card center\">No open times in the next few weeks. Please contact ' + esc(C.therapistName) +\n          (C.therapistPhone ? ' at <a href=\"tel:' + esc(C.therapistPhone) + '\">' + esc(C.therapistPhone) + '</a>' : '') + '.</div>';\n        return;\n      }\n      const daysHtml = open.map((d, i) => {\n        const dt = d.date + 'T12:00:00+05:30';\n        const n = d.slots.length;\n        return '<button type=\"button\" class=\"day' + (i === sel ? ' sel' : '') + '\" data-i=\"' + i + '\"' + (n ? '' : ' disabled') + '>' +\n          '<div class=\"dw\">' + fmt(dt, { weekday: 'short' }) + '</div><div class=\"dn\">' + fmt(dt, { day: 'numeric' }) + '</div>' +\n          '<div class=\"dm\">' + fmt(dt, { month: 'short' }) + '</div><div class=\"dc\">' + (n ? n + ' open' : (d.full ? 'full' : '\u2014')) + '</div></button>';\n      }).join('');\n      const day = open[sel];\n      const slotsHtml = day.slots.map(s => '<button type=\"button\" class=\"slot' + (picked === s ? ' sel' : '') + '\" data-s=\"' + s + '\">' +\n        fmtTime(s) + (SHOW_LOCAL ? '<small>' + fmt(s, { weekday: 'short', hour: 'numeric', minute: '2-digit' }, LOCAL_TZ) + ' your time</small>' : '') + '</button>').join('');\n      host.innerHTML = '<div class=\"days\">' + daysHtml + '</div>' +\n        '<div class=\"hint\" style=\"margin:-2px 0 12px\">' + fmt(day.date + 'T12:00:00+05:30', { weekday: 'long', day: 'numeric', month: 'long' }) +\n        ' \u00b7 times in IST' + (SHOW_LOCAL ? ' (your local time below each)' : '') + '</div>' +\n        '<div class=\"slots\">' + slotsHtml + '</div>';\n      host.querySelectorAll('.day').forEach(b => b.onclick = () => { sel = +b.dataset.i; picked = null; onPick(null); render(); });\n      host.querySelectorAll('.slot').forEach(b => b.onclick = () => { picked = b.dataset.s; onPick(picked); render(); });\n      const selDay = host.querySelector('.day.sel'); if (selDay && selDay.scrollIntoView) selDay.scrollIntoView({ inline: 'nearest', block: 'nearest' });\n    }\n    render();\n  }\n</script>\n"
};
