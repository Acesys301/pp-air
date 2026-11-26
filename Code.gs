/**
 * Pp air — staff check-in / check-out with QR + face check (Google Apps Script backend).
 * Data lives in this Google Sheet: tabs "Staff", "Logs", "Requests", "Leaves", "Settings".
 * Face photos and check-in selfies are saved (private) in the Google Drive folder "Pp air Photos".
 *
 * First time: run setup() once, then Deploy > New deployment > Web app
 *   Execute as: Me   |   Who has access: Anyone
 */

const APP = 'Pp air';
const QR_PREFIX = 'PPAIR:';
const TZ = 'Asia/Vientiane';
const MAX_TRIES = 5;            // wrong PIN (or 5 face mismatches) -> 15-minute lock
const REQUEST_DAYS = 7;         // missed-scan requests allowed for the last 7 days
const FACE_LEN = 128;           // face code length (face-api descriptor)

// Shifts: IN / OUT times. Change them here if Pp air works other hours.
const SHIFTS = {
  M: { name: 'Morning',   inT: '08:00', outT: '16:00', outNextDay: false },
  E: { name: 'Afternoon', inT: '16:00', outT: '00:00', outNextDay: true },
  N: { name: 'Night',     inT: '00:00', outT: '08:00', outNextDay: false }
};

const STAFF_COLS = ['StaffID', 'Name', 'Position', 'PIN', 'DeviceID (auto — clear to reset phone)', 'Active (YES/NO)', 'Role (STAFF/SUPERVISOR/ADMIN)', 'Face code (auto — clear both face cells to re-register)', 'Face photo (auto)'];
const LOG_COLS = ['Timestamp', 'Date', 'Time', 'StaffID', 'Name', 'Type', 'PhotoFileID', 'DeviceID', 'Status', 'Note', 'CheckedBy', 'Source', 'Face distance'];
const LEAVE_COLS = ['LeaveID', 'Created', 'StaffID', 'Name', 'Type', 'From', 'To', 'Days', 'Reason', 'Status', 'DecidedBy', 'DecidedAt', 'MedCert (Drive file id)', 'Dates'];
const REQ_COLS = ['ReqID', 'Created', 'StaffID', 'Name', 'Date', 'Shift', 'Missing', 'Reason', 'Status', 'DecidedBy', 'DecidedAt'];
const SICK_CERT_DAYS = 3;
const LEAVE_TYPES = { AL: 'Annual leave', SL: 'Sick leave', UL: 'Unpaid leave', PH: 'Public holiday', OFF: 'Other / OFF swap' };
const LEAVE_AHEAD_DAYS = 90, LEAVE_BACK_DAYS = 7, LEAVE_MAX_DAYS = 31;

// ---------------------------------------------------------------- setup
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  const st = ensureSheet('Staff', STAFF_COLS, ['A:A', 'D:E', 'H:I']);
  if (st.getLastRow() === 1) st.appendRow(['S001', 'Test Staff', 'Technician', '1234', '', 'YES', 'ADMIN']);
  ensureSheet('Logs', LOG_COLS, ['B:D']);
  ensureSheet('Requests', REQ_COLS, ['A:A', 'C:C', 'E:E']);
  ensureSheet('Leaves', LEAVE_COLS, ['A:A', 'C:C', 'E:G', 'N:N']);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('PHOTO_FOLDER')) props.setProperty('PHOTO_FOLDER', DriveApp.createFolder(APP + ' Photos').getId());

  const se = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (se.getLastRow() === 0) {
    se.getRange('B:B').setNumberFormat('@');
    se.appendRow(['Setting', 'Value', 'Note']);
    se.getRange('1:1').setFontWeight('bold');
    se.setColumnWidths(1, 3, 220);
  }
  const defaults = [
    ['Admin PIN', String(Math.floor(100000 + Math.random() * 900000)), 'PIN for the admin screen. You can change it.'],
    ['Kiosk key', Utilities.getUuid().replace(/-/g, '').slice(0, 10), 'Key for the entrance QR screen. Change it if it leaks.'],
    ['Shop latitude', '', 'Workplace latitude from Google Maps (e.g. 17.9667). Empty = GPS check off.'],
    ['Shop longitude', '', 'Workplace longitude from Google Maps (e.g. 102.6000).'],
    ['Radius (m)', '150', 'How far from the workplace check-in is allowed'],
    ['Face limit', '0.50', 'Face match strictness: 0.40 strict … 0.55 easy. Lower = harder to fool, more retries.'],
    ['Save selfies (YES/NO)', 'YES', 'YES = keep a photo of every check-in in Drive for supervisors.']
  ];
  const have = se.getDataRange().getDisplayValues().map(r => r[0]);
  defaults.forEach(d => { if (have.indexOf(d[0]) < 0) se.appendRow(d); });

  ss.getSheets().filter(s => /^(Sheet1|Trang tính1|ຊີດ1)$/.test(s.getName()) && s.getLastRow() === 0).forEach(s => ss.deleteSheet(s));
  Logger.log('Setup done. Open the Settings tab to see your Admin PIN and Kiosk key.');
}

function ensureSheet(name, cols, textRanges) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  (textRanges || []).forEach(r => sh.getRange(r).setNumberFormat('@'));
  sh.getRange(1, 1, 1, cols.length).setValues([cols])
    .setFontWeight('bold').setBackground('#B11F2B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

// ---------------------------------------------------------------- web entry
function doGet() {
  return ContentService.createTextOutput(APP + ' API is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.action];
    if (!fn) throw new Error('Unknown action');
    out = fn(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const API = {};

// ---------------------------------------------------------------- staff
API.login = req => withLock(() => {
  const s = auth(req);
  return { ok: true, id: s.id, name: s.name, position: s.position, role: s.role, hasFace: !!s.face, history: history(s.id), requests: myRequests(s.id), leaves: myLeaves(s.id), teamLeaves: teamLeaves() };
});

/** One-time face registration: face code (128 numbers) + a clear photo. Only HR can reset it (clear the two cells). */
API.setFace = req => withLock(() => {
  const s = auth(req);
  if (s.face) throw new Error('Face already registered. Ask HR to reset it.');
  const desc = cleanFace(req.face);
  const photo = String(req.photo || '');
  if (photo.length < 2000) throw new Error('Face photo missing. Please try again.');
  const now = new Date();
  const file = photoFolder().createFile(Utilities.newBlob(Utilities.base64Decode(photo.split(',').pop()), 'image/jpeg', 'FACE_' + s.id + '_' + fmt(now, 'yyyyMMdd_HHmmss') + '.jpg'));
  sheet('Staff').getRange(s.row, 8, 1, 2).setValues([[JSON.stringify(desc.map(x => Math.round(x * 10000) / 10000)), file.getId()]]);
  return { ok: true };
});

API.punch = req => withLock(() => {
  const s = auth(req);
  const type = req.type === 'OUT' ? 'OUT' : 'IN';
  if (!s.face) throw new Error('Register your face first.');
  if (!validQr(req.qr)) throw new Error('QR code not valid. Scan the code at the entrance again.');
  const gps = gpsCheck(req);
  const cache = CacheService.getScriptCache();
  const usedKey = 'used:' + s.id + ':' + type;
  if (cache.get(usedKey)) throw new Error('Already saved. Wait 1 minute before scanning again.');

  // face check (done on the server, the phone never gets the saved face code)
  guard('face:' + s.id);
  const dist = faceDistance(JSON.parse(s.face), cleanFace(req.face));
  const limit = parseFloat(settingOr('Face limit', '0.5')) || 0.5;
  if (!(dist <= limit)) {
    fail('face:' + s.id);
    throw new Error('Face does not match. Look straight at the camera in good light and try again.');
  }
  clearFails('face:' + s.id);

  const now = new Date();
  let photoId = '';
  const photo = String(req.photo || '');
  if (photo.length >= 2000 && settingOr('Save selfies (YES/NO)', 'YES').toUpperCase() !== 'NO') {
    const blob = Utilities.newBlob(Utilities.base64Decode(photo.split(',').pop()), 'image/jpeg', s.id + '_' + fmt(now, 'yyyyMMdd_HHmmss') + '_' + type + '.jpg');
    photoId = photoFolder().createFile(blob).getId();
  }
  const score = Math.round(dist * 100) / 100;
  sheet('Logs').appendRow([now, fmt(now, 'yyyy-MM-dd'), fmt(now, 'HH:mm:ss'), s.id, s.name, type, photoId, String(req.device || ''), '', gps.note + ' · face ' + score, '', 'QR+FACE', score]);
  cache.put(usedKey, '1', 60);
  return { ok: true, type: type, time: fmt(now, 'HH:mm'), date: fmt(now, 'dd MMM yyyy'), match: Math.max(0, Math.round((1 - dist) * 100)), history: history(s.id) };
});

function cleanFace(f) {
  if (!Array.isArray(f) || f.length !== FACE_LEN) throw new Error('Face not found. Please try again.');
  const v = f.map(Number);
  if (v.some(x => !isFinite(x) || Math.abs(x) > 1)) throw new Error('Face not found. Please try again.');
  return v;
}
function faceDistance(a, b) {
  let s = 0;
  for (let i = 0; i < FACE_LEN; i++) s += (a[i] - b[i]) * (a[i] - b[i]);
  return Math.sqrt(s);
}

/** Missed-scan request. Staff ask for themselves (PENDING).
 *  A supervisor or admin can add one for any staff member (forId) — approved at once. */
API.request = req => withLock(() => {
  let who, approver = null;
  if (req.forId) {
    approver = actor(req);
    who = staffById(req.forId);
    if (!who) throw new Error('Staff ID not found.');
    if (!approver.admin && who.id === approver.id) throw new Error('You cannot approve your own request.');
  } else {
    who = auth(req);
  }
  const shift = SHIFTS[req.shift] ? req.shift : null;
  if (!shift) throw new Error('Choose a shift.');
  const missing = ['IN', 'OUT', 'BOTH'].indexOf(req.missing) >= 0 ? req.missing : null;
  if (!missing) throw new Error('Choose what you forgot to scan.');
  const date = String(req.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Choose the date.');
  const todayStr = fmt(new Date(), 'yyyy-MM-dd');
  const oldest = fmt(new Date(Date.now() - REQUEST_DAYS * 86400000), 'yyyy-MM-dd');
  if (date > todayStr) throw new Error('The date cannot be in the future.');
  if (date < oldest) throw new Error('Too old. Requests are only for the last 7 days — ask HR.');
  const reason = String(req.reason || '').trim().slice(0, 300);
  if (!approver && reason.length < 3) throw new Error('Write the reason.');

  const rs = sheet('Requests');
  const dup = rs.getDataRange().getDisplayValues().slice(1)
    .some(r => r[2] === who.id && r[4] === date && r[5] === shift && r[8] === 'PENDING');
  if (dup) throw new Error('You already sent a request for this date and shift.');

  const id = 'R' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  rs.appendRow([id, new Date(), who.id, who.name, date, shift, missing, reason, 'PENDING', '', '']);
  if (approver) decide(rs.getLastRow(), 'APPROVED', approver);
  return { ok: true, id: id, status: approver ? 'APPROVED' : 'PENDING', requests: req.forId ? undefined : myRequests(who.id) };
});

// ---------------------------------------------------------------- leave requests
API.leave = req => withLock(() => {
  const s = auth(req);
  const type = String(req.type || '');
  if (!LEAVE_TYPES[type]) throw new Error('Choose the leave type.');
  let dates = Array.isArray(req.dates) ? req.dates.map(String) : expandDates(String(req.from || ''), String(req.to || req.from || ''));
  dates = dates.filter((d, i) => dates.indexOf(d) === i).sort();
  if (!dates.length || dates.some(d => !/^\d{4}-\d{2}-\d{2}$/.test(d))) throw new Error('Choose the dates.');
  if (dates[0] < fmt(new Date(Date.now() - LEAVE_BACK_DAYS * 86400000), 'yyyy-MM-dd')) throw new Error('Start date is too old. Ask HR.');
  if (dates[dates.length - 1] > fmt(new Date(Date.now() + LEAVE_AHEAD_DAYS * 86400000), 'yyyy-MM-dd')) throw new Error('Too far ahead. Leave can be requested up to 90 days ahead.');
  if (dates.length > LEAVE_MAX_DAYS) throw new Error('Too many days in one request (max 31).');
  const from = dates[0], to = dates[dates.length - 1], days = dates.length;
  const reason = String(req.reason || '').trim().slice(0, 300);
  if (type !== 'PH' && reason.length < 3) throw new Error('Write the reason.');
  const photo = String(req.cert || '');
  if (type === 'SL' && days >= SICK_CERT_DAYS && photo.length < 2000) throw new Error('Sick leave of 3 days or more needs a medical certificate. Attach a photo of it.');
  const ls = leaveSheet();
  const clash = ls.getDataRange().getDisplayValues().slice(1)
    .some(r => r[2] === s.id && (r[9] === 'PENDING' || r[9] === 'APPROVED') && leaveDatesOf(r).some(d => dates.indexOf(d) >= 0));
  if (clash) throw new Error('You already have a leave request for these dates.');
  const id = 'L' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  let certId = '';
  if (photo.length >= 2000) {
    const blob = Utilities.newBlob(Utilities.base64Decode(photo.split(',').pop()), 'image/jpeg', 'MEDCERT_' + s.id + '_' + from + '_' + id + '.jpg');
    certId = photoFolder().createFile(blob).getId();
  }
  ls.appendRow([id, new Date(), s.id, s.name, type, from, to, days, reason, 'PENDING', '', '', certId, dates.join(',')]);
  return { ok: true, id: id, leaves: myLeaves(s.id), teamLeaves: teamLeaves() };
});

API.leaveCancel = req => withLock(() => {
  const s = auth(req);
  const ls = leaveSheet();
  const v = ls.getDataRange().getDisplayValues();
  const i = v.findIndex((r, k) => k > 0 && r[0] === String(req.leaveId));
  if (i < 1 || v[i][2] !== s.id) throw new Error('Request not found.');
  if (v[i][9] !== 'PENDING') throw new Error('Already decided - you cannot cancel it.');
  ls.getRange(i + 1, 10).setValue('CANCELLED');
  return { ok: true, leaves: myLeaves(s.id), teamLeaves: teamLeaves() };
});

API.leaveDecide = req => withLock(() => {
  const a = actor(req);
  if (!a.admin) throw new Error('Admin only.');
  const ls = leaveSheet();
  const v = ls.getDataRange().getDisplayValues();
  const i = v.findIndex((r, k) => k > 0 && r[0] === String(req.leaveId));
  if (i < 1) throw new Error('Request not found.');
  if (v[i][9] !== 'PENDING') throw new Error('Already decided by ' + (v[i][10] || 'someone') + '.');
  ls.getRange(i + 1, 10, 1, 3).setValues([[req.decision === 'APPROVED' ? 'APPROVED' : 'REJECTED', a.name, fmt(new Date(), 'yyyy-MM-dd HH:mm')]]);
  return { ok: true };
});

API.exportLeaves = req => {
  if (!actor(req).admin) throw new Error('Admin only.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(req.from || '') || !/^\d{4}-\d{2}-\d{2}$/.test(req.to || '')) throw new Error('Pick both dates.');
  return { ok: true, rows: allLeaves().filter(r => r.dates.some(d => d >= req.from && d <= req.to)) };
};

function leaveSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Leaves');
  if (!sh || sh.getLastColumn() < LEAVE_COLS.length) return ensureSheet('Leaves', LEAVE_COLS, ['A:A', 'C:C', 'E:G', 'N:N']);
  return sh;
}
function expandDates(from, to) {
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(from) || !re.test(to) || to < from || dayDiff(from, to) > 400) return [];
  const out = [];
  for (let t = Date.parse(from + 'T00:00:00Z'); t <= Date.parse(to + 'T00:00:00Z'); t += 86400000) out.push(new Date(t).toISOString().slice(0, 10));
  return out;
}
function leaveDatesOf(r) {
  const own = String(r[13] || '').split(',').map(x => x.trim()).filter(Boolean);
  return own.length ? own : expandDates(String(r[5]), String(r[6]));
}
function leaveObj(r, row) {
  return { row: row, id: r[0], created: r[1], staffId: r[2], name: r[3], type: r[4], from: r[5], to: r[6], days: Number(r[7]) || 0, reason: r[8], status: r[9], by: r[10], decidedAt: r[11], cert: r[12] || '', dates: leaveDatesOf(r) };
}
function allLeaves() {
  const ls = leaveSheet();
  if (ls.getLastRow() < 2) return [];
  return ls.getDataRange().getDisplayValues().slice(1).map((r, i) => leaveObj(r, i + 2)).filter(r => r.id);
}
function teamLeaves() {
  const cut = fmt(new Date(Date.now() - 7 * 86400000), 'yyyy-MM-dd');
  return allLeaves().filter(l => l.status === 'PENDING' || l.status === 'APPROVED')
    .map(l => ({ name: l.name, type: l.type, status: l.status, dates: l.dates.filter(d => d >= cut) }))
    .filter(l => l.dates.length);
}
function myLeaves(id) { return allLeaves().filter(r => r.staffId === id).slice(-20).reverse(); }
function dayDiff(a, b) { return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000); }

// ---------------------------------------------------------------- entrance QR screen
API.qr = req => {
  if (!req.kiosk || String(req.kiosk) !== setting('Kiosk key')) throw new Error('Wrong kiosk key.');
  return { ok: true, token: QR_PREFIX + staticToken(), now: fmt(new Date(), 'HH:mm:ss') };
};

// ---------------------------------------------------------------- supervisor + admin
API.day = req => {
  const a = actor(req);
  const date = /^\d{4}-\d{2}-\d{2}$/.test(req.date || '') ? req.date : fmt(new Date(), 'yyyy-MM-dd');
  const st = sheet('Staff').getDataRange().getDisplayValues().slice(1)
    .filter(r => r[0] && String(r[5]).toUpperCase() !== 'NO')
    .map(r => ({ id: r[0], name: r[1], position: r[2], role: roleOf(r), device: r[4] ? 'locked' : 'not set', face: r[7] ? 'yes' : '', ref: r[8] || '' }));
  const refOf = {}; st.forEach(x => { refOf[x.id] = x.ref; });
  const reqs = sheet('Requests').getDataRange().getDisplayValues().slice(1)
    .map((r, i) => reqObj(r, i + 2))
    .filter(r => r.status === 'PENDING' || r.decidedAt.slice(0, 10) === date)
    .reverse();
  const rows = logRows(date, date); rows.forEach(r => { r.ref = refOf[r.id] || ''; });
  const out = { ok: true, date: date, rows: rows, staff: st, requests: reqs, me: a.name, admin: !!a.admin };
  if (a.admin) {
    out.kioskKey = setting('Kiosk key');
    const cut = fmt(new Date(Date.now() - 30 * 86400000), 'yyyy-MM-dd');
    out.leaves = allLeaves().filter(l => l.status === 'PENDING' || l.to >= cut).reverse();
  }
  return out;
};

API.photo = req => {
  actor(req);
  const f = DriveApp.getFileById(String(req.fileId));
  if (f.getParents().next().getId() !== photoFolder().getId()) throw new Error('Not a ' + APP + ' photo.');
  return { ok: true, data: 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
};

API.check = req => withLock(() => {
  const a = actor(req);
  const row = parseInt(req.row, 10);
  const lg = sheet('Logs');
  if (!(row > 1 && row <= lg.getLastRow())) throw new Error('Row not found.');
  const status = req.status === 'FLAG' ? 'FLAG' : (req.status === 'OK' ? 'OK' : '');
  if (!a.admin && lg.getRange(row, 4).getDisplayValue() === a.id) throw new Error('You cannot verify your own scan.');
  const note = lg.getRange(row, 10).getDisplayValue().split(' ‖ ')[0];
  lg.getRange(row, 9, 1, 3).setValues([[status, status === 'FLAG' ? note + ' ‖ ' + String(req.note || 'Check photo').slice(0, 200) : note, status ? a.name : '']]);
  return { ok: true };
});

API.decide = req => withLock(() => {
  const a = actor(req);
  const rs = sheet('Requests');
  const v = rs.getDataRange().getDisplayValues();
  const i = v.findIndex((r, k) => k > 0 && r[0] === String(req.reqId));
  if (i < 1) throw new Error('Request not found.');
  if (v[i][8] !== 'PENDING') throw new Error('Already decided by ' + (v[i][9] || 'someone') + '.');
  if (!a.admin && v[i][2] === a.id) throw new Error('You cannot approve your own request.');
  decide(i + 1, req.decision === 'APPROVED' ? 'APPROVED' : 'REJECTED', a);
  return { ok: true };
});

API.export = req => {
  if (!actor(req).admin) throw new Error('Admin only.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(req.from || '') || !/^\d{4}-\d{2}-\d{2}$/.test(req.to || '')) throw new Error('Pick both dates.');
  return { ok: true, rows: logRows(req.from, req.to) };
};

// ---------------------------------------------------------------- helpers
function decide(rowNum, decision, a) {
  const rs = sheet('Requests');
  const r = rs.getRange(rowNum, 1, 1, REQ_COLS.length).getDisplayValues()[0];
  rs.getRange(rowNum, 9, 1, 3).setValues([[decision, a.name, fmt(new Date(), 'yyyy-MM-dd HH:mm')]]);
  if (decision !== 'APPROVED') return;
  const sh = SHIFTS[r[5]];
  const types = r[6] === 'BOTH' ? ['IN', 'OUT'] : [r[6]];
  const lg = sheet('Logs');
  types.forEach(t => {
    let day = r[4];
    if (t === 'OUT' && sh.outNextDay) day = fmt(new Date(new Date(day + 'T12:00:00+07:00').getTime() + 86400000), 'yyyy-MM-dd');
    const hhmm = t === 'IN' ? sh.inT : sh.outT;
    const ts = new Date(day + 'T' + hhmm + ':00+07:00');
    lg.appendRow([ts, day, hhmm + ':00', r[2], r[3], t, '', '', 'OK',
      'Missed scan · ' + sh.name + ' shift · ' + r[0] + (r[7] ? ' · ' + r[7] : ''), a.name, 'MANUAL', '']);
  });
}

function reqObj(r, row) {
  return { row: row, id: r[0], created: r[1], staffId: r[2], name: r[3], date: r[4], shift: r[5], missing: r[6], reason: r[7], status: r[8], by: r[9], decidedAt: r[10] };
}

function myRequests(id) {
  const rs = sheet('Requests');
  if (rs.getLastRow() < 2) return [];
  return rs.getDataRange().getDisplayValues().slice(1).map((r, i) => reqObj(r, i + 2))
    .filter(r => r.staffId === id).slice(-15).reverse();
}

function roleOf(r) {
  const v = String(r[6] || '').trim().toUpperCase();
  return v === 'ADMIN' || v === 'SUPERVISOR' ? v : 'STAFF';
}

function staffById(id) {
  id = String(id || '').trim().toUpperCase();
  const v = sheet('Staff').getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toUpperCase() === id && String(v[i][5]).toUpperCase() !== 'NO') {
      return { id: String(v[i][0]).trim(), name: v[i][1], position: v[i][2], role: roleOf(v[i]) };
    }
  }
  return null;
}

function auth(req) {
  const id = String(req.id || '').trim().toUpperCase();
  if (!id) throw new Error('Enter your staff ID.');
  guard('staff:' + id);
  const sh = sheet('Staff');
  const v = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toUpperCase() !== id) continue;
    if (String(v[i][3]).trim() === '' || String(v[i][3]).trim() !== String(req.pin || '').trim()) break;
    if (String(v[i][5]).trim().toUpperCase() === 'NO') throw new Error('This staff ID is not active. Ask HR.');
    const device = String(req.device || '');
    if (device.length < 10) throw new Error('Phone not recognised. Reload the app.');
    const saved = String(v[i][4]).trim();
    if (!saved) sh.getRange(i + 1, 5).setValue(device);
    else if (saved !== device) throw new Error('Your account is locked to another phone. Ask HR to reset it.');
    clearFails('staff:' + id);
    return { row: i + 1, id: String(v[i][0]).trim(), name: v[i][1], position: v[i][2], role: roleOf(v[i]), face: String(v[i][7] || '').trim() };
  }
  fail('staff:' + id);
  throw new Error('Wrong staff ID or PIN.');
}

function actor(req) {
  if (req.admin) { checkAdmin(req.admin); return { admin: true, id: 'ADMIN', name: 'Admin' }; }
  const s = auth(req);
  if (s.role === 'ADMIN') { s.admin = true; return s; }
  if (s.role !== 'SUPERVISOR') throw new Error('Supervisors only.');
  return s;
}

function checkAdmin(pin) {
  guard('admin');
  if (!pin || String(pin) !== setting('Admin PIN')) { fail('admin'); throw new Error('Wrong admin PIN.'); }
  clearFails('admin');
}

function guard(key) {
  const n = Number(CacheService.getScriptCache().get('fail:' + key) || 0);
  if (n >= MAX_TRIES) throw new Error('Too many wrong tries. Wait 15 minutes.');
}
function fail(key) {
  const c = CacheService.getScriptCache();
  c.put('fail:' + key, String(Number(c.get('fail:' + key) || 0) + 1), 900);
}
function clearFails(key) { CacheService.getScriptCache().remove('fail:' + key); }

/** One fixed QR code. The token is signed with the secret, so it cannot be guessed. */
function staticToken() {
  const sig = Utilities.computeHmacSha256Signature('static-qr', secret());
  return 'S.' + Utilities.base64EncodeWebSafe(sig).slice(0, 12);
}
function validQr(t) {
  return String(t || '') === QR_PREFIX + staticToken();
}

function history(id) {
  const lg = sheet('Logs');
  const n = lg.getLastRow();
  if (n < 2) return [];
  const start = Math.max(2, n - 5000);
  const ym = fmt(new Date(), 'yyyy-MM');
  return lg.getRange(start, 1, n - start + 1, 12).getDisplayValues()
    .filter(r => r[3] === id && r[1].indexOf(ym) === 0)
    .map(r => ({ date: r[1], time: r[2].slice(0, 5), type: r[5], source: r[11] || 'QR', status: r[8] }))
    .sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time));
}

function logRows(from, to) {
  const lg = sheet('Logs');
  const n = lg.getLastRow();
  if (n < 2) return [];
  const w = Math.max(12, Math.min(13, lg.getLastColumn()));
  return lg.getRange(2, 1, n - 1, w).getDisplayValues()
    .map((r, i) => ({ row: i + 2, date: r[1], time: r[2], id: r[3], name: r[4], type: r[5], photo: r[6], status: r[8], note: r[9], by: r[10], source: r[11] || 'QR', face: r[12] || '' }))
    .filter(r => r.date >= from && r.date <= to)
    .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" missing. Run setup() first.');
  return s;
}
function setting(key) {
  const v = sheet('Settings').getDataRange().getDisplayValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) return String(v[i][1]).trim();
  throw new Error('Setting "' + key + '" missing. Run setup() first.');
}
function settingOr(key, def) { try { return setting(key); } catch (e) { return def; } }

/** GPS: phone must be within the radius of the workplace (only if the location is set in Settings). */
function gpsCheck(req) {
  const lat0 = parseFloat(settingOr('Shop latitude', '')), lng0 = parseFloat(settingOr('Shop longitude', ''));
  if (isNaN(lat0) || isNaN(lng0)) return { note: 'GPS check off' };
  const lat = parseFloat(req.lat), lng = parseFloat(req.lng), acc = Math.min(parseFloat(req.acc) || 0, 100);
  if (isNaN(lat) || isNaN(lng)) throw new Error('Location is off. Turn on location (GPS) and allow it for this app.');
  const R = 6371000, toR = x => x * Math.PI / 180;
  const dLat = toR(lat - lat0), dLng = toR(lng - lng0);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toR(lat0)) * Math.cos(toR(lat)) * Math.sin(dLng / 2) ** 2;
  const dist = Math.round(2 * R * Math.asin(Math.sqrt(a)));
  const radius = parseFloat(settingOr('Radius (m)', '150')) || 150;
  if (dist > radius + acc) throw new Error('You are not at the workplace (' + dist + ' m away).');
  return { note: 'GPS ' + dist + ' m' };
}
function secret() { return PropertiesService.getScriptProperties().getProperty('SECRET'); }
function photoFolder() { return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER')); }
function fmt(d, p) { return Utilities.formatDate(d, TZ, p); }
