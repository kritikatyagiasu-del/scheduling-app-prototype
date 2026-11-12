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

