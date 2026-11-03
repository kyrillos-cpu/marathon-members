/****************************************************************
 * MARATHON MEMBERS — Backend (Google Apps Script + Google Sheet)
 * ---------------------------------------------------------------
 * 1) Open your Google Sheet > Extensions > Apps Script
 * 2) Paste this whole file into Code.gs
 * 3) Change ADMIN_CODE below
 * 4) Run setup() once (accept permissions)
 * 5) Deploy > New deployment > Web app
 *      Execute as: Me    |    Who has access: Anyone
 * 6) Copy the /exec URL into config.js on GitHub
 ****************************************************************/

const ADMIN_CODE = 'CHANGE-ME';   // <-- the leaders' secret code

const SHEETS = {
  Members:    ['ID','Name','Phone','Teams','Role','Church','JoinDate','Status','Notes','CreatedAt'],
  Teams:      ['Name','Color','Lead','CreatedAt'],
  Sessions:   ['ID','Type','Title','Date','StartTime','MissionID','Notes','CreatedAt'],
  Attendance: ['SessionID','MemberID','Time','Status','Rating','Method','Note'],
  Missions:   ['ID','Title','Type','Eparchy','Date','EndDate','Location','SlotsNeeded','Status','Source','BookingKey','Notes','CreatedAt'],
  Signups:    ['MissionID','MemberID','Team','Status','SignedAt','Note'],
  Bookings:   ['Key','Eparchy','Type','Date','EndDate','Contact','Phone','Status','MissionID','SyncedAt'],
  Settings:   ['Key','Value']
};

const DEFAULT_TEAMS = [
  ['Training','#2563eb'], ['Media','#db2777'], ['Eparchy Coordination','#059669'],
  ['Transportation','#d97706'], ['Gifts & Cash Box','#7c3aed'], ['Sunday Meeting','#0891b2']
];

const DEFAULT_SETTINGS = {
  SERVICE_NAME: 'Marathon Members',
  LATE_AFTER_MIN: '15',
  REQUIRED_TYPES: 'Sunday Meeting,Training',
  W_ATTENDANCE: '50', W_COMMITMENT: '20', W_PARTICIPATION: '30',
  TYPE_WEIGHT_SUNDAY: '1', TYPE_WEIGHT_TRAINING: '2', TYPE_WEIGHT_CARNIVAL: '2',
  BOOKING_SHEET_ID: '1y5M1BQLA5QMegs22yBkJB1n1WIA5_KQo949K4HqQK1U', BOOKING_TAB: 'طلبات الحجز',
  BOOKING_CONFIRMED_VALUES: 'تم التأكيد,تمت الخدمة,confirmed,agreement finalized,finalized,service completed,completed,done,تم الاتفاق,مؤكد,تم التأكيد,تم التنفيذ',
  BOOKING_COL_KEY: '', BOOKING_COL_EPARCHY: '', BOOKING_COL_TYPE: '', BOOKING_COL_DATE: '',
  BOOKING_COL_ENDDATE: '', BOOKING_COL_STATUS: '', BOOKING_COL_CONTACT: '', BOOKING_COL_PHONE: '',
  LAST_BOOKING_SYNC: ''
};

/* ======================= HTTP ======================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.payload) return respond_(JSON.parse(p.payload));
  return respond_({ action: 'ping' });
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return respond_(body);
}
function respond_(req) {
  let out;
  try { out = { ok: true, data: route_(req.action, req.data || {}, req.code) }; }
  catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, d, code) {
  const PUBLIC = {
    ping: () => ({ pong: true, time: now_() }),
    publicMissions: publicMissions_, lookupMember: lookupMember_,
    signup: signup_, cancelSignup: cancelSignup_, memberCard: memberCard_
  };
  const ADMIN = {
    login: () => ({ ok: true }), bootstrap: bootstrap_,
    saveMember: saveMember_, importMembers: importMembers_, deleteMember: deleteMember_,
    saveTeam: saveTeam_, deleteTeam: deleteTeam_,
    saveSession: saveSession_, deleteSession: deleteSession_,
    checkin: checkin_, checkinBatch: checkinBatch_, setAttendance: setAttendance_,
    saveMission: saveMission_, deleteMission: deleteMission_, setSignup: setSignup_,
    syncBookings: syncBookings, saveSettings: saveSettings_
  };
  if (PUBLIC[action]) return withLock_(() => PUBLIC[action](d));
  if (ADMIN[action]) {
    if (String(code || '') !== String(ADMIN_CODE)) throw new Error('Wrong access code');
    return withLock_(() => ADMIN[action](d));
  }
  throw new Error('Unknown action: ' + action);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================= SETUP ======================= */
function setup() {
  Object.keys(SHEETS).forEach(n => sh_(n));
  if (readTable_('Teams').length === 0) {
    DEFAULT_TEAMS.forEach(t => writeObj_('Teams', { Name: t[0], Color: t[1], Lead: '', CreatedAt: now_() }));
  }
  const st = readTable_('Settings').map(r => r.Key);
  Object.keys(DEFAULT_SETTINGS).forEach(k => {
    if (st.indexOf(k) < 0) writeObj_('Settings', { Key: k, Value: DEFAULT_SETTINGS[k] });
  });
  const def = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('ورقة1');
  if (def && def.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(def);
  Logger.log('Setup done. Now deploy as Web App.');
}

/** Optional: run once to auto-sync bookings every hour */
function enableAutoSync() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'autoSync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoSync').timeBased().everyHours(1).create();
  Logger.log('Auto sync enabled (hourly).');
}
function autoSync() { try { withLock_(() => syncBookings()); } catch (e) { Logger.log(e); } }

/* ======================= SHEET HELPERS ======================= */
let _ss = null; const _sh = {};
function ss_() { return _ss || (_ss = SpreadsheetApp.getActiveSpreadsheet()); }
function sh_(name) {
  if (_sh[name]) return _sh[name];
  let s = ss_().getSheetByName(name);
  const want = SHEETS[name];
  if (!s) {
    s = ss_().insertSheet(name);
    s.getRange('A:Z').setNumberFormat('@');
    s.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold').setBackground('#eef2ff');
    s.setFrozenRows(1);
  } else {
    const lastCol = Math.max(1, s.getLastColumn());
    const have = s.getRange(1, 1, 1, lastCol).getDisplayValues()[0].filter(String);
    const missing = want.filter(h => have.indexOf(h) < 0);
    if (missing.length) s.getRange(1, have.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return (_sh[name] = s);
}
function headers_(name) {
  const s = sh_(name);
  return s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0];
}
function rows_(name) {
  const s = sh_(name);
  const v = s.getDataRange().getDisplayValues();
  const h = v[0]; const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i].some(c => c !== '')) continue;
    const o = {}; h.forEach((k, j) => { if (k) o[k] = v[i][j]; });
    out.push({ row: i + 1, obj: o });
  }
  return out;
}
function readTable_(name) { return rows_(name).map(r => r.obj); }
function writeObj_(name, obj, rowNum) {
  const s = sh_(name); const h = headers_(name);
  const row = h.map(k => (obj[k] == null ? '' : String(obj[k])));
  if (rowNum) s.getRange(rowNum, 1, 1, h.length).setValues([row]);
  else {
    const r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, h.length).setNumberFormat('@').setValues([row]);
  }
  return obj;
}
function upsert_(name, pred, patch) {
  const found = rows_(name).filter(r => pred(r.obj))[0];
  if (found) { const o = Object.assign({}, found.obj, patch); writeObj_(name, o, found.row); return o; }
  return writeObj_(name, patch);
}
function deleteWhere_(name, pred) {
  const rows = rows_(name).filter(r => pred(r.obj)).map(r => r.row).sort((a, b) => b - a);
  rows.forEach(r => sh_(name).deleteRow(r));
  return rows.length;
}
function nextId_(name, prefix, pad) {
  let max = 0;
  readTable_(name).forEach(r => {
    const m = String(r.ID || '').match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + String(max + 1).padStart(pad, '0');
}
function tz_() { return Session.getScriptTimeZone() || 'Africa/Cairo'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function phoneKey_(p) {
  // converts Arabic-Indic digits, keeps last 10 digits (Egyptian mobile)
  const d = String(p || '').replace(/[٠-٩]/g, c => '٠١٢٣٤٥٦٧٨٩'.indexOf(c)).replace(/[^\d]/g, '');
  return d.slice(-10);
}
function settings_() {
  const o = Object.assign({}, DEFAULT_SETTINGS);
  readTable_('Settings').forEach(r => { if (r.Key) o[r.Key] = r.Value; });
  return o;
}
function setSetting_(k, v) { upsert_('Settings', r => r.Key === k, { Key: k, Value: v }); }
function findMember_(key) {
  key = String(key || '').trim();
  if (!key) return null;
  const members = readTable_('Members');
  const idm = key.toUpperCase().match(/MM-\d+/);
  if (idm) return members.filter(m => m.ID.toUpperCase() === idm[0])[0] || null;
  const pk = phoneKey_(key);
  if (pk.length < 9) return null;
  return members.filter(m => phoneKey_(m.Phone) === pk)[0] || null;
}

/* ======================= ADMIN: DATA ======================= */
function bootstrap_() {
  const st = settings_();
  return {
    members: readTable_('Members'), teams: readTable_('Teams'),
    sessions: readTable_('Sessions'), attendance: readTable_('Attendance'),
    missions: readTable_('Missions'), signups: readTable_('Signups'),
    bookings: readTable_('Bookings'), settings: st, serverTime: now_()
  };
}

function saveMember_(d) {
  if (!d.Name || !String(d.Name).trim()) throw new Error('Name is required');
  const members = readTable_('Members');
  if (d.Phone) {
    const pk = phoneKey_(d.Phone);
    const dup = members.filter(m => m.ID !== d.ID && pk && phoneKey_(m.Phone) === pk)[0];
    if (dup) throw new Error('Phone already used by ' + dup.Name);
  }
  const patch = {
    Name: String(d.Name).trim(), Phone: d.Phone || '', Teams: d.Teams || '', Role: d.Role || 'Servant',
    Church: d.Church || '', JoinDate: d.JoinDate || today_(), Status: d.Status || 'Active', Notes: d.Notes || ''
  };
  if (d.ID && members.some(m => m.ID === d.ID)) return upsert_('Members', r => r.ID === d.ID, patch);
  patch.ID = nextId_('Members', 'MM-', 3); patch.CreatedAt = now_();
  return writeObj_('Members', patch);
}

function importMembers_(d) {
  const list = d.rows || [];
  let added = 0, skipped = 0;
  const seen = {};
  readTable_('Members').forEach(m => { const k = phoneKey_(m.Phone); if (k) seen[k] = 1; seen['n:' + m.Name.trim().toLowerCase()] = 1; });
  list.forEach(r => {
    const name = String(r.Name || '').trim(); if (!name) return;
    const pk = phoneKey_(r.Phone);
    if ((pk && seen[pk]) || seen['n:' + name.toLowerCase()]) { skipped++; return; }
    saveMember_({ Name: name, Phone: r.Phone || '', Teams: r.Teams || '', Role: r.Role || 'Servant', Church: r.Church || '' });
    if (pk) seen[pk] = 1; seen['n:' + name.toLowerCase()] = 1; added++;
  });
  return { added, skipped };
}

function deleteMember_(d) {
  deleteWhere_('Members', r => r.ID === d.ID);
  deleteWhere_('Attendance', r => r.MemberID === d.ID);
  deleteWhere_('Signups', r => r.MemberID === d.ID);
  return { deleted: d.ID };
}

function saveTeam_(d) {
  const name = String(d.Name || '').trim();
  if (!name) throw new Error('Team name is required');
  const old = d.OldName || name;
  const teams = readTable_('Teams');
  if (old !== name && teams.some(t => t.Name === name)) throw new Error('Team already exists');
  const res = upsert_('Teams', r => r.Name === old, {
    Name: name, Color: d.Color || '#64748b', Lead: d.Lead || '',
    CreatedAt: (teams.filter(t => t.Name === old)[0] || {}).CreatedAt || now_()
  });
  if (old !== name) renameTeamInMembers_(old, name);
  return res;
}
function renameTeamInMembers_(oldName, newName) {
  rows_('Members').forEach(r => {
    const list = splitList_(r.obj.Teams);
    if (list.indexOf(oldName) < 0) return;
    const nl = list.map(t => (t === oldName ? newName : t)).filter(Boolean);
    r.obj.Teams = Array.from(new Set(nl)).join(', ');
    writeObj_('Members', r.obj, r.row);
  });
}
function deleteTeam_(d) {
  deleteWhere_('Teams', r => r.Name === d.Name);
  renameTeamInMembers_(d.Name, '');
  return { deleted: d.Name };
}
function splitList_(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }

function saveSession_(d) {
  if (!d.Type || !d.Date) throw new Error('Type and date are required');
  const patch = { Type: d.Type, Title: d.Title || '', Date: d.Date, StartTime: d.StartTime || '', MissionID: d.MissionID || '', Notes: d.Notes || '' };
  if (d.ID && readTable_('Sessions').some(s => s.ID === d.ID)) return upsert_('Sessions', r => r.ID === d.ID, patch);
  patch.ID = nextId_('Sessions', 'S-', 4); patch.CreatedAt = now_();
  return writeObj_('Sessions', patch);
}
function deleteSession_(d) {
  deleteWhere_('Sessions', r => r.ID === d.ID);
  deleteWhere_('Attendance', r => r.SessionID === d.ID);
  return { deleted: d.ID };
}

function lateStatus_(session, time, lateMin) {
  if (!session.StartTime || !time) return 'Present';
  if (String(time).slice(0, 10) !== session.Date) return 'Present';
  const hm = s => { const m = String(s).match(/(\d{1,2}):(\d{2})/); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };
  const st = hm(session.StartTime), t = hm(String(time).slice(11));
  if (st == null || t == null) return 'Present';
  return t > st + Number(lateMin || 15) ? 'Late' : 'Present';
}

function checkin_(d) {
  const session = readTable_('Sessions').filter(s => s.ID === d.SessionID)[0];
  if (!session) throw new Error('Session not found');
  const member = findMember_(d.MemberID);
  if (!member) throw new Error('Member not found: ' + d.MemberID);
  const existing = readTable_('Attendance').filter(a => a.SessionID === session.ID && a.MemberID === member.ID)[0];
  if (existing) return { duplicate: true, member: { ID: member.ID, Name: member.Name }, record: existing };
  const time = d.Time || now_();
  const rec = {
    SessionID: session.ID, MemberID: member.ID, Time: time,
    Status: d.Status || lateStatus_(session, time, settings_().LATE_AFTER_MIN),
    Rating: d.Rating || '', Method: d.Method || 'QR', Note: d.Note || ''
  };
  writeObj_('Attendance', rec);
  // If session belongs to a mission, mark signup as attended
  if (session.MissionID) {
    const sp = r => r.MissionID === session.MissionID && r.MemberID === member.ID;
    const hasSignup = readTable_('Signups').some(sp);
    upsert_('Signups', sp, hasSignup ? { Status: 'Attended' } :
      { MissionID: session.MissionID, MemberID: member.ID, Team: (splitList_(member.Teams)[0] || ''), Status: 'Attended', SignedAt: time, Note: 'Checked in' });
  }
  return { duplicate: false, member: { ID: member.ID, Name: member.Name }, record: rec };
}
function checkinBatch_(d) {
  return (d.items || []).map(it => { try { return Object.assign({ ok: true }, checkin_(it)); } catch (e) { return { ok: false, error: String(e.message), item: it }; } });
}

function setAttendance_(d) {
  const pred = r => r.SessionID === d.SessionID && r.MemberID === d.MemberID;
  if (d.remove) { deleteWhere_('Attendance', pred); return { removed: true }; }
  const cur = readTable_('Attendance').filter(pred)[0];
  const patch = { SessionID: d.SessionID, MemberID: d.MemberID };
  ['Status', 'Rating', 'Note', 'Method', 'Time'].forEach(k => { if (d[k] !== undefined) patch[k] = d[k]; });
  if (!cur) { patch.Time = patch.Time || now_(); patch.Method = patch.Method || 'Manual'; patch.Status = patch.Status || 'Present'; }
  return upsert_('Attendance', pred, patch);
}

function saveMission_(d) {
  if (!d.Title || !d.Date) throw new Error('Title and date are required');
  const patch = {
    Title: d.Title, Type: d.Type || 'Carnival', Eparchy: d.Eparchy || '', Date: d.Date, EndDate: d.EndDate || '',
    Location: d.Location || '', SlotsNeeded: d.SlotsNeeded || '', Status: d.Status || 'Open', Notes: d.Notes || ''
  };
  if (d.ID && readTable_('Missions').some(m => m.ID === d.ID)) return upsert_('Missions', r => r.ID === d.ID, patch);
  patch.ID = nextId_('Missions', 'MS-', 3); patch.Source = 'Manual'; patch.BookingKey = ''; patch.CreatedAt = now_();
  return writeObj_('Missions', patch);
}
function deleteMission_(d) {
  deleteWhere_('Missions', r => r.ID === d.ID);
  deleteWhere_('Signups', r => r.MissionID === d.ID);
  rows_('Sessions').forEach(r => { if (r.obj.MissionID === d.ID) { r.obj.MissionID = ''; writeObj_('Sessions', r.obj, r.row); } });
  rows_('Bookings').forEach(r => { if (r.obj.MissionID === d.ID) { r.obj.MissionID = 'DELETED'; writeObj_('Bookings', r.obj, r.row); } });
  return { deleted: d.ID };
}
function setSignup_(d) {
  const pred = r => r.MissionID === d.MissionID && r.MemberID === d.MemberID;
  if (d.Status === 'Removed') { deleteWhere_('Signups', pred); return { removed: true }; }
  const cur = readTable_('Signups').filter(pred)[0];
  const patch = { MissionID: d.MissionID, MemberID: d.MemberID, Status: d.Status || 'Signed' };
  if (d.Team !== undefined) patch.Team = d.Team;
  if (!cur) { patch.SignedAt = now_(); patch.Note = d.Note || 'Added by leader'; }
  return upsert_('Signups', pred, patch);
}
function saveSettings_(d) {
  const s = d.settings || {};
  Object.keys(s).forEach(k => { if (k !== 'LAST_BOOKING_SYNC') setSetting_(k, s[k]); });
  return settings_();
}

/* ======================= PUBLIC ======================= */
function missionCounts_() {
  const c = {};
  readTable_('Signups').forEach(s => { if (s.Status !== 'Cancelled') c[s.MissionID] = (c[s.MissionID] || 0) + 1; });
  return c;
}
function publicMissions_() {
  const t = today_(); const c = missionCounts_();
  const missions = readTable_('Missions')
    .filter(m => m.Status === 'Open' && ((m.EndDate || m.Date) >= t))
    .sort((a, b) => a.Date < b.Date ? -1 : 1)
    .map(m => ({ ID: m.ID, Title: m.Title, Type: m.Type, Eparchy: m.Eparchy, Date: m.Date, EndDate: m.EndDate,
      Location: m.Location, SlotsNeeded: m.SlotsNeeded, Notes: m.Notes, Signed: c[m.ID] || 0 }));
  return { missions, teams: readTable_('Teams').map(x => x.Name), serviceName: settings_().SERVICE_NAME };
}
function lookupMember_(d) {
  const m = findMember_(d.key);
  if (!m) throw new Error('We could not find you. Check your phone number or member ID.');
  if (m.Status === 'Inactive') throw new Error('Your profile is inactive. Contact the service leader.');
  const mine = readTable_('Signups').filter(s => s.MemberID === m.ID).map(s => ({ MissionID: s.MissionID, Status: s.Status, Team: s.Team }));
  return { ID: m.ID, Name: m.Name, Teams: m.Teams, signups: mine };
}
function signup_(d) {
  const m = findMember_(d.key);
  if (!m || m.Status === 'Inactive') throw new Error('Member not found');
  const mission = readTable_('Missions').filter(x => x.ID === d.MissionID)[0];
  if (!mission || mission.Status !== 'Open') throw new Error('This mission is not open for sign-up');
  const ex = readTable_('Signups').filter(s => s.MissionID === mission.ID && s.MemberID === m.ID)[0];
  if (ex && ex.Status !== 'Cancelled') return { already: true, Status: ex.Status };
  const signed = missionCounts_()[mission.ID] || 0;
  const slots = Number(mission.SlotsNeeded || 0);
  const status = slots && signed >= slots ? 'Waitlist' : 'Signed';
  upsert_('Signups', r => r.MissionID === mission.ID && r.MemberID === m.ID, {
    MissionID: mission.ID, MemberID: m.ID, Team: d.Team || splitList_(m.Teams)[0] || '', Status: status, SignedAt: now_(), Note: d.Note || ''
  });
  return { Status: status };
}
function cancelSignup_(d) {
  const m = findMember_(d.key);
  if (!m) throw new Error('Member not found');
  const mission = readTable_('Missions').filter(x => x.ID === d.MissionID)[0];
  if (mission && mission.Date < today_()) throw new Error('This mission already started');
  deleteWhere_('Signups', r => r.MissionID === d.MissionID && r.MemberID === m.ID && r.Status !== 'Attended');
  return { cancelled: true };
}
function memberCard_(d) {
  const m = findMember_(d.ID);
  if (!m) throw new Error('Member not found');
  return { ID: m.ID, Name: m.Name, Teams: m.Teams, Role: m.Role, serviceName: settings_().SERVICE_NAME };
}

/* ======================= BOOKINGS SYNC ======================= */
/** Reads the other Google Sheet (eparchy bookings) and brings CONFIRMED dates here as Missions. */
function syncBookings() {
  const st = settings_();
  if (!st.BOOKING_SHEET_ID) throw new Error('Add the bookings sheet link in Settings first');
  const idm = String(st.BOOKING_SHEET_ID).match(/[-\w]{25,}/);
  const book = SpreadsheetApp.openById(idm ? idm[0] : st.BOOKING_SHEET_ID);
  const sheet = st.BOOKING_TAB ? book.getSheetByName(st.BOOKING_TAB) : book.getSheets()[0];
  if (!sheet) throw new Error('Tab not found in bookings sheet: ' + st.BOOKING_TAB);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { found: 0, created: 0, updated: 0 };

  // header row = first row (of first 5) with 3+ filled cells
  let hr = 0;
  for (let i = 0; i < Math.min(5, values.length); i++) { if (values[i].filter(c => String(c).trim()).length >= 3) { hr = i; break; } }
  const headers = values[hr].map(h => String(h).trim());
  const low = headers.map(h => h.toLowerCase());
  const used = {};
  const pick = (override, words, exact) => {
    if (override) { const i = low.indexOf(String(override).trim().toLowerCase()); if (i >= 0) { used[i] = 1; return i; } }
    for (const w of words) {
      for (let i = 0; i < low.length; i++) {
        if (used[i] || !low[i]) continue;
        if (/timestamp|طابع زمني|created|submitted/.test(low[i])) continue;
        if (exact ? low[i] === w : low[i].indexOf(w) >= 0) { used[i] = 1; return i; }
      }
    }
    return -1;
  };
  // Marathon booking form: separate training / carnival date + place columns
  const findCol = (must, avoid) => {
    for (let i = 0; i < low.length; i++) {
      if (used[i] || !low[i]) continue;
      if (must.every(w => w.some(x => low[i].indexOf(x) >= 0)) && !avoid.some(x => low[i].indexOf(x) >= 0)) { used[i] = 1; return i; }
    }
    return -1;
  };
  const ALT = ['بديل', 'alt', 'عدد', 'count'];
  const split = {
    tDate:  findCol([['تدريب', 'training'], ['ميعاد', 'تاريخ', 'date']], ALT),
    cDate:  findCol([['كرنفال', 'carnival'], ['ميعاد', 'تاريخ', 'date']], ALT),
    tPlace: findCol([['تدريب', 'training'], ['مكان', 'place', 'location']], []),
    cPlace: findCol([['كرنفال', 'carnival'], ['مكان', 'place', 'location']], []),
    tCount: findCol([['خدام', 'servants'], ['عدد', 'count']], []),
    cCount: findCol([['مخدوم', 'kids', 'أطفال'], ['عدد', 'count']], [])
  };
  const isSplit = split.tDate >= 0 || split.cDate >= 0;
  const col = {
    status:  pick(st.BOOKING_COL_STATUS,  ['status', 'الحالة', 'حالة']),
    key:     pick(st.BOOKING_COL_KEY,     ['id', 'booking id', 'request id', '#', 'كود', 'رقم الحجز', 'رقم الطلب'], true),
    enddate: isSplit ? -1 : pick(st.BOOKING_COL_ENDDATE, ['end date', 'to date', 'until', 'إلى', 'الى', 'حتى', 'نهاية']),
    date:    isSplit ? -1 : pick(st.BOOKING_COL_DATE, ['confirmed date', 'final date', 'agreed date', 'التاريخ المؤكد', 'الميعاد المؤكد', 'start date', 'date', 'تاريخ', 'ميعاد', 'الميعاد']),
    eparchy: pick(st.BOOKING_COL_EPARCHY, ['eparchy', 'diocese', 'إيبارشية', 'ايبارشية', 'إبارشية', 'ابارشية', 'church', 'كنيسة']),
    type:    pick(st.BOOKING_COL_TYPE,    ['activity type', 'mission type', 'type', 'نوع', 'النشاط']),
    contact: pick(st.BOOKING_COL_CONTACT, ['responsible', 'contact', 'servant name', 'مسئول', 'مسؤول', 'المسؤول', 'المسئول', 'الخادم', 'name', 'اسم']),
    phone:   pick(st.BOOKING_COL_PHONE,   ['phone', 'mobile', 'موبايل', 'تليفون', 'هاتف', 'whatsapp', 'واتس'])
  };
  // stable row key: ID column, else the submission time column
  let stampCol = -1;
  if (col.key < 0) low.forEach((h, i) => { if (stampCol < 0 && /timestamp|طابع زمني|وقت التسجيل|submitted|created/.test(h)) stampCol = i; });
  if (!isSplit && col.enddate < 0) col.enddate = pick('', ['to', 'end', 'إلى', 'الى'], true);
  if (col.status < 0) throw new Error('Could not find a Status column in the bookings sheet. Set "Status column" in Settings. Headers: ' + headers.join(' | '));
  if (!isSplit && col.date < 0) throw new Error('Could not find a Date column in the bookings sheet. Set "Date column" in Settings. Headers: ' + headers.join(' | '));
  const okVals = String(st.BOOKING_CONFIRMED_VALUES || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const isConfirmed = v => {
    v = String(v || '').trim().toLowerCase(); if (!v) return false;
    if (/\bnot\b|غير|لم |لا /.test(v)) return false;
    return okVals.some(o => v === o || v.indexOf(o) >= 0);
  };
  const fmtDate = v => {
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
    const s = String(v || '').trim();
    let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/);
    if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0'); }
    return s;
  };
  const mapType = v => {
    const s = String(v || '').toLowerCase();
    if (/تدريب|train/.test(s)) return 'Training';
    if (/كرنفال|carn/.test(s)) return 'Carnival';
    if (/متابع|follow/.test(s)) return 'Follow-up';
    return s ? String(v) : 'Carnival';
  };
  const get = (r, i) => (i >= 0 ? r[i] : '');

  const missions = rows_('Missions');
  const bookingsRows = rows_('Bookings');
  const seen = {};
  let found = 0, created = 0, updated = 0;
  const stamp = now_();

  for (let i = hr + 1; i < values.length; i++) {
    const r = values[i];
    if (!r.some(c => String(c).trim())) continue;
    if (!isConfirmed(get(r, col.status))) continue;
    const eparchy = String(get(r, col.eparchy) || '').trim();
    const rawType = String(get(r, col.type) || '');
    const stampV = stampCol >= 0 ? r[stampCol] : '';
    const rawKey = String(get(r, col.key) || '').trim() ||
      (stampV ? (Object.prototype.toString.call(stampV) === '[object Date]' ? String(stampV.getTime()) : String(stampV).trim()) : '');
    const items = [];
    if (isSplit) {
      const wantT = /تدريب|train/i.test(rawType) || (!rawType && split.tDate >= 0);
      const wantC = /كرنفال|carn/i.test(rawType) || (!rawType && split.cDate >= 0);
      if (wantT && split.tDate >= 0 && String(get(r, split.tDate)).trim())
        items.push({ type: 'Training', suffix: '|T', date: fmtDate(get(r, split.tDate)), place: String(get(r, split.tPlace) || ''), count: String(get(r, split.tCount) || '') });
      if (wantC && split.cDate >= 0 && String(get(r, split.cDate)).trim())
        items.push({ type: 'Carnival', suffix: '|C', date: fmtDate(get(r, split.cDate)), place: String(get(r, split.cPlace) || ''), count: String(get(r, split.cCount) || '') });
    } else {
      items.push({ type: mapType(rawType), suffix: '', date: fmtDate(get(r, col.date)), endDate: col.enddate >= 0 ? fmtDate(get(r, col.enddate)) : '', place: '', count: '' });
    }
    items.forEach(it => {
      found++;
      const key = (rawKey ? 'B-' + rawKey : 'R-' + [eparchy, it.type, it.date].join('|')) + it.suffix;
      const endDate = it.endDate || '';
      const location = it.place ? (it.place + (eparchy ? ' – ' + eparchy : '')) : eparchy;
      seen[key] = 1;
      let mRow = missions.filter(m => m.obj.BookingKey === key)[0];
      let missionId;
      if (!mRow) {
        const prev = bookingsRows.filter(b => b.obj.Key === key)[0];
        if (prev && prev.obj.MissionID === 'DELETED') return; // leader deleted it on purpose
        const m = {
          ID: nextId_('Missions', 'MS-', 3), Title: it.type + ' – ' + (eparchy || 'Eparchy'), Type: it.type, Eparchy: eparchy,
          Date: it.date, EndDate: endDate, Location: location, SlotsNeeded: '', Status: 'Open', Source: 'Booking',
          BookingKey: key, Notes: 'Confirmed booking' + (it.count ? ' · expected: ' + it.count : ''), CreatedAt: stamp
        };
        writeObj_('Missions', m);
        missions.push({ row: sh_('Missions').getLastRow(), obj: m });
        missionId = m.ID; created++;
      } else {
        missionId = mRow.obj.ID;
        if (mRow.obj.Date !== it.date || mRow.obj.EndDate !== endDate || mRow.obj.Eparchy !== eparchy) {
          Object.assign(mRow.obj, { Date: it.date, EndDate: endDate, Eparchy: eparchy });
          writeObj_('Missions', mRow.obj, mRow.row); updated++;
        }
      }
      upsert_('Bookings', b => b.Key === key, {
        Key: key, Eparchy: eparchy, Type: it.type, Date: it.date, EndDate: endDate,
        Contact: String(get(r, col.contact) || ''), Phone: String(get(r, col.phone) || ''),
        Status: String(get(r, col.status) || ''), MissionID: missionId, SyncedAt: stamp
      });
    });
  }
  // bookings that are no longer confirmed
  rows_('Bookings').forEach(b => {
    if (!seen[b.obj.Key] && b.obj.Status !== 'No longer confirmed' && b.obj.MissionID !== 'DELETED') {
      b.obj.Status = 'No longer confirmed'; b.obj.SyncedAt = stamp; writeObj_('Bookings', b.obj, b.row);
    }
  });
  setSetting_('LAST_BOOKING_SYNC', stamp);
  const detected = {}; Object.keys(col).forEach(k => detected[k] = col[k] >= 0 ? headers[col[k]] : '—');
  if (isSplit) ['tDate', 'cDate', 'tPlace', 'cPlace'].forEach(k => detected[k] = split[k] >= 0 ? headers[split[k]] : '—');
  delete detected.date; delete detected.enddate;
  return { found, created, updated, columns: detected, time: stamp };
}
