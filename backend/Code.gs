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

const ADMIN_CODE = 'CHANGE-ME';   // leaders' access code

const SHEETS = {
  Members:    ['ID','Name','Phone','Teams','Role','Church','JoinDate','Status','Notes','CreatedAt','Email','Photo'],
  Teams:      ['Name','Color','Lead','CreatedAt'],
  Sessions:   ['ID','Type','Title','Date','StartTime','MissionID','Notes','CreatedAt', 'Recap'],
  Attendance: ['SessionID','MemberID','Time','Status','Rating','Method','Note'],
  Missions:   ['ID','Title','Type','Eparchy','Date','EndDate','Location','SlotsNeeded','Status','Source','BookingKey','Notes','CreatedAt','EventID','CoordName','CoordPhone'],
  Signups:    ['MissionID','MemberID','Team','Status','SignedAt','Note'],
  Cancellations: ['MissionID','MemberID','Name','Mission','MissionDate','PrevStatus','SignedAt','CancelledAt','Seen'],
  Bookings:   ['Key','Eparchy','Type','Date','EndDate','Contact','Phone','Status','MissionID','SyncedAt','Coordinators'],
  Library:    ['ID','Category','Title','Link','Note','AddedAt'],
  Announcements: ['ID','Text','Until','CreatedAt','By','Status','Images'],
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
  LAST_BOOKING_SYNC: '',
  LIBRARY_CARNIVAL_TAB: 'ملفات الكرنفال',
  BOOKING_COORD_TAB: 'منسقين الإيبارشيات',
  CALENDAR_NAME: 'Marathon Missions',
  APPROVER_TEAMS: 'Training', ANNOUNCER_TEAMS: 'Media',
  REMINDER_HOUR: '18'
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
  try { ensureAbsentTrigger_(); } catch (e) { Logger.log(e); }
  try { out = { ok: true, data: route_(req.action, req.data || {}, req.code) }; }
  catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, d, code) {
  const PUBLIC = {
    ping: () => ({ pong: true, time: now_() }),
    publicMissions: publicMissions_, lookupMember: lookupMember_,
    signup: signup_, cancelSignup: cancelSignup_, memberCard: memberCard_,
    myTeam: myTeam_, setMyEmail: setMyEmail_, uploadPhoto: d => { const m = findMember_(d.key); if (!m || m.Status === 'Inactive') throw new Error('Member not found'); return savePhoto_(m, d.data); }, approvals: approvals_, decide: decide_,
    teamAnnouncements: d => { checkAnnouncer_(d); return readTable_('Announcements'); },
    teamSaveAnnouncement: d => { const m = checkAnnouncer_(d); return saveAnnouncement_(Object.assign({}, d, { By: m.Name })); },
    teamDeleteAnnouncement: d => { checkAnnouncer_(d); return deleteAnnouncement_(d.ID); },
    joinData: d => ({ me: lookupMember_(d), pub: publicMissions_() }),
    library: d => { const m = findMember_(d.key); if (!m || m.Status === 'Inactive') throw new Error('Member not found'); return getLibrary_(); }
  };
  const ADMIN = {
    login: () => ({ ok: true }), bootstrap: bootstrap_,
    saveMember: saveMember_, importMembers: importMembers_, deleteMember: deleteMember_,
    saveTeam: saveTeam_, deleteTeam: deleteTeam_,
    saveSession: saveSession_, deleteSession: deleteSession_,
    checkin: checkin_, checkinBatch: checkinBatch_, setAttendance: setAttendance_,
    saveMission: saveMission_, deleteMission: deleteMission_, setSignup: setSignup_,
    syncBookings: syncBookings, saveSettings: saveSettings_,
    getLibrary: getLibrary_, saveLibraryItem: saveLibraryItem_, deleteLibraryItem: deleteLibraryItem_,
    setMissionCoord: d => { upsert_('Missions', r => r.ID === d.ID, { CoordName: d.name || '', CoordPhone: d.phone || '' }); calMission_(d.ID); return readTable_('Missions').filter(x => x.ID === d.ID)[0]; },
    setMemberPhoto: d => { const m = readTable_('Members').filter(x => x.ID === d.ID)[0]; if (!m) throw new Error('Member not found'); return savePhoto_(m, d.data); },
    saveAnnouncement: saveAnnouncement_,
    seenCancellations: () => { const sh = sh_('Cancellations'); const h = headers_('Cancellations'); const col = h.indexOf('Seen') + 1; const n = sh.getLastRow() - 1; if (n > 0 && col > 0) sh.getRange(2, col, n, 1).setValues(Array(n).fill(['1'])); return readTable_('Cancellations'); },
    reviewAnnouncement: d => { upsert_('Announcements', r => r.ID === d.ID, { Status: d.Status === 'Approved' ? 'Approved' : 'Rejected' }); return readTable_('Announcements'); }, deleteAnnouncement: d => deleteAnnouncement_(d.ID)
  };
  // read-only actions don't wait for the write lock (faster loading)
  const READ_ONLY = ['teamAnnouncements', 'approvals', 'joinData', 'ping', 'publicMissions', 'lookupMember', 'memberCard', 'myTeam', 'library', 'login', 'bootstrap', 'getLibrary'];
  const run = (fn) => READ_ONLY.indexOf(action) >= 0 ? fn(d) : withLock_(() => fn(d));
  if (PUBLIC[action]) return run(PUBLIC[action]);
  if (ADMIN[action]) {
    if (String(code || '') !== String(ADMIN_CODE)) throw new Error('Wrong access code');
    return run(ADMIN[action]);
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
    missions: readTable_('Missions'), signups: readTable_('Signups'), cancellations: readTable_('Cancellations'),
    bookings: readTable_('Bookings'), announcements: readTable_('Announcements'), settings: st, serverTime: now_()
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
    Church: d.Church || '', JoinDate: d.JoinDate || today_(), Status: d.Status || 'Active', Notes: d.Notes || '',
    Email: String(d.Email || '').trim().toLowerCase()
  };
  const before = members.filter(m => m.ID === d.ID)[0];
  let res;
  if (before) res = upsert_('Members', r => r.ID === d.ID, patch);
  else { patch.ID = nextId_('Members', 'MM-', 3); patch.CreatedAt = now_(); res = writeObj_('Members', patch); }
  if (patch.Email && (!before || before.Email !== patch.Email) && patch.Status !== 'Inactive') calAddGuest_(patch.Email);
  return res;
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
    saveMember_({ Name: name, Phone: r.Phone || '', Teams: r.Teams || '', Role: r.Role || 'Servant', Church: r.Church || '', Notes: r.Notes || '' });
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
  if (d.Recap !== undefined) patch.Recap = String(d.Recap || '');
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
  if (d.ID && readTable_('Missions').some(m => m.ID === d.ID)) { const res = upsert_('Missions', r => r.ID === d.ID, patch); calMission_(d.ID); return res; }
  patch.ID = nextId_('Missions', 'MS-', 3); patch.Source = 'Manual'; patch.BookingKey = ''; patch.CreatedAt = now_();
  writeObj_('Missions', patch); calMission_(patch.ID);
  return readTable_('Missions').filter(x => x.ID === patch.ID)[0] || patch;
}
function deleteMission_(d) {
  readTable_('Missions').filter(r => r.ID === d.ID).forEach(r => calDelete_(r.EventID));
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
  readTable_('Signups').forEach(s => { if (['Cancelled', 'Declined'].indexOf(s.Status) < 0) c[s.MissionID] = (c[s.MissionID] || 0) + 1; });
  return c;
}
function publicMissions_() {
  const t = today_(); const c = missionCounts_();
  const missions = readTable_('Missions')
    .filter(m => m.Status === 'Open' && ((m.EndDate || m.Date) >= t))
    .sort((a, b) => a.Date < b.Date ? -1 : 1)
    .map(m => ({ ID: m.ID, Title: m.Title, Type: m.Type, Eparchy: m.Eparchy, Date: m.Date, EndDate: m.EndDate,
      Location: m.Location, SlotsNeeded: m.SlotsNeeded, Notes: m.Notes, Signed: c[m.ID] || 0 }));
  return { missions, teams: readTable_('Teams').map(x => x.Name), serviceName: settings_().SERVICE_NAME, announcements: activeAnnouncements_() };
}
function lookupMember_(d) {
  const m = findMember_(d.key);
  if (!m) throw new Error('We could not find you. Check your phone number or member ID.');
  if (m.Status === 'Inactive') throw new Error('Your profile is inactive. Contact the service leader.');
  const mine = readTable_('Signups').filter(s => s.MemberID === m.ID).map(s => ({ MissionID: s.MissionID, Status: s.Status, Team: s.Team }));
  return { ID: m.ID, Name: m.Name, Teams: m.Teams, Photo: m.Photo || '', HasEmail: !!m.Email, IsApprover: isApprover_(m), IsAnnouncer: isAnnouncer_(m), signups: mine };
}
function signup_(d) {
  const m = findMember_(d.key);
  if (!m || m.Status === 'Inactive') throw new Error('Member not found');
  const mission = readTable_('Missions').filter(x => x.ID === d.MissionID)[0];
  if (!mission || mission.Status !== 'Open') throw new Error('This mission is not open for sign-up');
  const ex = readTable_('Signups').filter(s => s.MissionID === mission.ID && s.MemberID === m.ID)[0];
  if (ex && ex.Status !== 'Cancelled') return { already: true, Status: ex.Status };
  const status = 'Pending'; // waits for approval by the service leader or the training team lead
  upsert_('Signups', r => r.MissionID === mission.ID && r.MemberID === m.ID, {
    MissionID: mission.ID, MemberID: m.ID, Team: d.Team || splitList_(m.Teams)[0] || '', Status: status, SignedAt: now_(), Note: d.Note || ''
  });
  try {
    const st = sundayStats_(m.ID, m.JoinDate);
    if (st.flag) {
      const to = Session.getEffectiveUser().getEmail();
      if (to) MailApp.sendEmail({ to: to, subject: '⚠️ ' + m.Name + ' requested ' + mission.Title + ' — missed Sunday meetings',
        body: m.Name + ' (' + (m.Phone || m.ID) + ') requested to join:\n' + mission.Title + ' · ' + mission.Date +
          '\n\nSunday meeting: missed ' + st.missed + ' of the last ' + st.of + (st.streak >= 2 ? ' (the last ' + st.streak + ' in a row)' : '') +
          '.\n\nOpen the app to approve or decline.' });
    }
  } catch (e) { Logger.log(e); }
  return { Status: status };
}
function cancelSignup_(d) {
  const m = findMember_(d.key);
  if (!m) throw new Error('Member not found');
  const mission = readTable_('Missions').filter(x => x.ID === d.MissionID)[0];
  if (mission && mission.Date < today_()) throw new Error('This mission already started');
  const sg = readTable_('Signups').filter(r => r.MissionID === d.MissionID && r.MemberID === m.ID && r.Status !== 'Attended')[0];
  deleteWhere_('Signups', r => r.MissionID === d.MissionID && r.MemberID === m.ID && r.Status !== 'Attended');
  if (sg) {
    // keep a record so the leader knows who cancelled
    writeObj_('Cancellations', { MissionID: d.MissionID, MemberID: m.ID, Name: m.Name, Mission: mission ? mission.Title : d.MissionID,
      MissionDate: mission ? mission.Date : '', PrevStatus: sg.Status, SignedAt: sg.SignedAt || '', CancelledAt: now_(), Seen: '' });
    try {
      const to = Session.getEffectiveUser().getEmail();
      if (to) MailApp.sendEmail({ to: to, subject: '❌ ' + m.Name + ' cancelled: ' + (mission ? mission.Title : d.MissionID),
        body: m.Name + ' (' + (m.Phone || m.ID) + ') cancelled their ' + (sg.Status === 'Pending' ? 'request' : 'sign-up') + ' for:\n\n' +
          (mission ? mission.Title + '\n' + mission.Date + (mission.Location ? ' · ' + mission.Location : '') : d.MissionID) +
          '\n\nStatus before cancelling: ' + sg.Status + '\nCancelled at: ' + now_() });
    } catch (e) { Logger.log(e); }
  }
  return { cancelled: true };
}
function memberCard_(d) {
  const m = findMember_(d.ID);
  if (!m) throw new Error('Member not found');
  return { ID: m.ID, Name: m.Name, Teams: m.Teams, Role: m.Role, Photo: m.Photo || '', serviceName: settings_().SERVICE_NAME };
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
  // eparchy coordinators registered by the admins (tab «منسقين الإيبارشيات») — their phone wins over whoever filled the booking
  const normEp = x => String(x || '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/^(ايبارشيه|ابارشيه|الايبارشيه|الابارشيه)\s*/, '').replace(/\s+/g, '').toLowerCase();
  const coords = {};
  try {
    const cs = book.getSheetByName(st.BOOKING_COORD_TAB || 'منسقين الإيبارشيات');
    if (cs && cs.getLastRow() > 1) {
      const cv = cs.getDataRange().getDisplayValues(); const ch = cv[0].map(h => String(h).trim());
      const ci = k => ch.findIndex(h => k.some(w => h.indexOf(w) >= 0));
      const iN = ci(['الاسم', 'name']), iE = ci(['الإيبارشية', 'الايبارشية', 'eparchy']), iP = ci(['الموبايل', 'mobile', 'phone']), iM = ci(['الإيميل', 'الايميل', 'email']);
      for (let i = 1; i < cv.length; i++) {
        const ep = normEp(cv[i][iE]); if (!ep) continue;
        (coords[ep] = coords[ep] || []).push({ name: String(cv[i][iN] || '').trim(), phone: String(cv[i][iP] || '').replace(/\s+/g, ''), email: iM >= 0 ? String(cv[i][iM] || '').trim().toLowerCase() : '' });
      }
    }
  } catch (e) {}
  const digits = x => String(x || '').replace(/[٠-٩]/g, c => '٠١٢٣٤٥٦٧٨٩'.indexOf(c)).replace(/\D/g, '').slice(-10);
  /** Which coordinator belongs to this booking?
   *  1) the coordinator who made the booking (same phone / email)  2) otherwise the first one registered for the eparchy. */
  const pickCoord = (list, r) => {
    if (!list || !list.length) return null;
    const ph = [digits(get(r, col.phone)), digits(get(r, col.whats))].filter(x => x.length >= 9);
    const em = [String(get(r, col.email) || '').trim().toLowerCase(), String(get(r, col.by) || '').trim().toLowerCase()].filter(Boolean);
    return list.filter(c => ph.indexOf(digits(c.phone)) >= 0 || (c.email && em.indexOf(c.email) >= 0))[0] || list[0];
  };

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
  col.by = pick('', ['سجّله', 'سجله', 'registered by', 'submitted by']);
  col.whats = pick('', ['واتساب', 'واتس', 'whatsapp']);
  col.email = pick('', ['الإيميل', 'الايميل', 'email', 'e-mail']);
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
      calMission_(missionId);
      } else {
        missionId = mRow.obj.ID;
        if (mRow.obj.Date !== it.date || mRow.obj.EndDate !== endDate || mRow.obj.Eparchy !== eparchy) {
          Object.assign(mRow.obj, { Date: it.date, EndDate: endDate, Eparchy: eparchy });
          writeObj_('Missions', mRow.obj, mRow.row); updated++;
        calMission_(missionId);
        }
      }
      upsert_('Bookings', b => b.Key === key, {
        Key: key, Eparchy: eparchy, Type: it.type, Date: it.date, EndDate: endDate,
        Contact: (pickCoord(coords[normEp(eparchy)], r) || {}).name || String(get(r, col.contact) || ''),
        Phone: (pickCoord(coords[normEp(eparchy)], r) || {}).phone || String(get(r, col.phone) || ''),
        Coordinators: JSON.stringify((coords[normEp(eparchy)] || []).map(c => ({ name: c.name, phone: c.phone }))),
        Status: String(get(r, col.status) || ''), MissionID: missionId, SyncedAt: stamp
      });
      calMission_(missionId); // keeps calendar text (coordinator) and guest list up to date
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

/* ======================= LIBRARY ======================= */
/** Carnival files live in the eparchy booking sheet (tab «ملفات الكرنفال») so they show in both apps.
 *  Other files live in this sheet's Library tab. */
function carnivalTab_(create) {
  const st = settings_();
  if (!st.BOOKING_SHEET_ID) return null;
  const idm = String(st.BOOKING_SHEET_ID).match(/[-\w]{25,}/);
  const book = SpreadsheetApp.openById(idm ? idm[0] : st.BOOKING_SHEET_ID);
  const name = st.LIBRARY_CARNIVAL_TAB || 'ملفات الكرنفال';
  let sh = book.getSheetByName(name);
  if (!sh && create) { sh = book.insertSheet(name); sh.getRange(1, 1, 1, 6).setValues([['المرحلة','الموضوع','الجزء','العنوان','اللينك','ملاحظة']]).setFontWeight('bold'); }
  return sh;
}
function getLibrary_() {
  let carnival = [];
  try {
    const sh = carnivalTab_(false);
    if (sh && sh.getLastRow() > 1) {
      const v = sh.getRange(1, 1, sh.getLastRow(), 6).getDisplayValues();
      for (let i = 1; i < v.length; i++) {
        const r = v[i]; if (!r[4]) continue;
        carnival.push({ row: i + 1, stage: r[0].trim(), topic: r[1].trim(), part: r[2].trim(), title: r[3], url: r[4].trim(), note: r[5] });
      }
    }
  } catch (e) { carnival = []; }
  const general = readTable_('Library').filter(x => x.Link).map(x => ({ id: x.ID, category: x.Category || 'General', title: x.Title, url: x.Link, note: x.Note, added: x.AddedAt }));
  return { carnival, general };
}
function saveLibraryItem_(d) {
  if (!d.url || !/^https?:\/\//i.test(String(d.url).trim())) throw new Error('Paste a full link starting with https://');
  if (d.section === 'carnival') {
    const sh = carnivalTab_(true);
    if (!sh) throw new Error('Bookings sheet is not set in Settings');
    sh.appendRow([d.stage || '', d.topic || '', d.part || '', d.title || '', String(d.url).trim(), d.note || '']);
  } else {
    writeObj_('Library', { ID: nextId_('Library', 'L-', 3), Category: d.category || 'General', Title: d.title || '', Link: String(d.url).trim(), Note: d.note || '', AddedAt: now_() });
  }
  return getLibrary_();
}
function deleteLibraryItem_(d) {
  if (d.section === 'carnival') {
    const sh = carnivalTab_(false);
    const row = Number(d.row);
    if (sh && row > 1 && row <= sh.getLastRow()) {
      const url = sh.getRange(row, 5).getDisplayValue().trim();
      if (url === String(d.url || '').trim()) sh.deleteRow(row); else throw new Error('The file list changed — refresh and try again');
    }
  } else {
    deleteWhere_('Library', r => r.ID === d.id);
  }
  return getLibrary_();
}

/** Servant view: names and mobile numbers of the people in the same team(s) */
function myTeam_(d) {
  const m = findMember_(d.key);
  if (!m || m.Status === 'Inactive') throw new Error('Member not found');
  const mine = splitList_(m.Teams);
  const members = readTable_('Members').filter(x => x.Status !== 'Inactive');
  const teams = readTable_('Teams');
  return mine.map(t => {
    const tm = teams.filter(x => x.Name === t)[0] || {};
    const lead = members.filter(x => x.ID === tm.Lead)[0];
    return {
      team: t, color: tm.Color || '#64748b', lead: lead ? lead.Name : '',
      members: members.filter(x => splitList_(x.Teams).indexOf(t) >= 0)
        .map(x => ({ ID: x.ID, Name: x.Name, Phone: x.Phone, Photo: x.Photo || '', me: x.ID === m.ID, lead: x.ID === tm.Lead }))
        .sort((a, b) => (b.lead - a.lead) || a.Name.localeCompare(b.Name))
    };
  });
}

/* ======================= CALENDAR + REMINDERS ======================= */
/** Every mission is ONE event on the "Marathon Missions" calendar (owner = this script's account).
 *  All active servants who have an email are guests → it appears in everybody's calendar,
 *  whether they signed up or not. Guests can't see each other's emails. */
function calendar_() {
  const name = settings_().CALENDAR_NAME || 'Marathon Missions';
  return CalendarApp.getCalendarsByName(name)[0] || CalendarApp.createCalendar(name, { color: CalendarApp.Color.ORANGE });
}
function dateOf_(s) { const p = String(s || '').slice(0, 10).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function missionTexts_(m) {
  const b = readTable_('Bookings').filter(x => x.MissionID === m.ID)[0] || {};
  const cn = m.CoordName || m.CoordPhone ? m.CoordName : b.Contact, cp = m.CoordName || m.CoordPhone ? m.CoordPhone : b.Phone;
  const contact = cn || cp ? 'Local coordinator: ' + (cn || '') + (cp ? ' – ' + cp : '') : '';
  const notes = m.Notes && !/^Confirmed booking/.test(m.Notes) ? m.Notes : '';
  return { title: 'Marathon – ' + m.Title, location: m.Location || m.Eparchy || '', contact: contact, notes: notes,
    description: [m.Type + ' mission' + (m.Eparchy ? ' – ' + m.Eparchy : ''), notes, contact, 'Sign up: https://kyrillos-cpu.github.io/marathon-members/?join=1'].filter(String).join('\n') };
}
function calDelete_(id) {
  if (!id) return;
  try { const ev = calendar_().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) { Logger.log(e); }
}
function allEmails_() {
  return readTable_('Members').filter(m => m.Status !== 'Inactive' && m.Email).map(m => String(m.Email).trim().toLowerCase())
    .filter((e, i, a) => e && a.indexOf(e) === i);
}
/** Create / update the calendar event of a mission. */
function calMission_(missionId) {
  try {
    const row = rows_('Missions').filter(r => r.obj.ID === missionId)[0];
    if (!row) return;
    const m = row.obj;
    if (m.Status === 'Cancelled') { if (m.EventID) { calDelete_(m.EventID); m.EventID = ''; writeObj_('Missions', m, row.row); } return; }
    if ((m.EndDate || m.Date) < today_()) return;
    const t = missionTexts_(m);
    const start = dateOf_(m.Date), end = dateOf_(m.EndDate || m.Date); end.setDate(end.getDate() + 1);
    let ev = null;
    if (m.EventID) { try { ev = calendar_().getEventById(m.EventID); } catch (e) {} }
    const emails = allEmails_();
    if (!ev) {
      ev = calendar_().createAllDayEvent(t.title, start, end, { description: t.description, location: t.location, guests: emails.join(','), sendInvites: true });
      try { ev.setGuestsCanSeeGuests(false); ev.setGuestsCanInviteOthers(false); ev.addPopupReminder(24 * 60); } catch (e) {}
      m.EventID = ev.getId(); writeObj_('Missions', m, row.row);
    } else {
      ev.setTitle(t.title); ev.setLocation(t.location); ev.setDescription(t.description); ev.setAllDayDates(start, end);
      const have = ev.getGuestList().map(g => String(g.getEmail()).toLowerCase());
      emails.filter(e => have.indexOf(e) < 0).forEach(e => ev.addGuest(e));
    }
  } catch (e) { Logger.log('calMission_ ' + e); }
}
/** New email → add it to every upcoming mission. */
function calAddGuest_(email) {
  try {
    email = String(email).trim().toLowerCase();
    readTable_('Missions').filter(m => m.Status !== 'Cancelled' && (m.EndDate || m.Date) >= today_()).forEach(m => {
      if (!m.EventID) { calMission_(m.ID); return; }
      const ev = calendar_().getEventById(m.EventID);
      if (ev && !ev.getGuestList().some(g => String(g.getEmail()).toLowerCase() === email)) ev.addGuest(email);
    });
  } catch (e) { Logger.log('calAddGuest_ ' + e); }
}
/** Servant adds his email from the sign-up page. */
function setMyEmail_(d) {
  const m = findMember_(d.key);
  if (!m || m.Status === 'Inactive') throw new Error('Member not found');
  const email = String(d.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Please write a correct email');
  upsert_('Members', r => r.ID === m.ID, { Email: email });
  calAddGuest_(email);
  return { ok: true };
}
/** Daily (evening): email every servant who is going tomorrow. */
function dailyReminder() {
  const tz = tz_();
  const t = new Date(); t.setDate(t.getDate() + 1);
  const tomorrow = Utilities.formatDate(t, tz, 'yyyy-MM-dd');
  const members = readTable_('Members');
  const signups = readTable_('Signups');
  readTable_('Missions').filter(m => m.Date === tomorrow && m.Status !== 'Cancelled').forEach(m => {
    const x = missionTexts_(m);
    signups.filter(s => s.MissionID === m.ID && (s.Status === 'Signed' || s.Status === 'Confirmed')).forEach(s => {
      const mem = members.filter(k => k.ID === s.MemberID)[0];
      if (!mem || !mem.Email) return;
      try {
        MailApp.sendEmail({
          to: mem.Email, name: 'Marathon Members', subject: '⏰ Tomorrow: ' + m.Title,
          htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.6">Hi ' + mem.Name.split(' ')[0] + ' 👋<br><br>' +
            'Reminder: you are serving <b>tomorrow</b> in <b>' + m.Title + '</b>.<br>📅 ' + m.Date + (m.Location ? '<br>📍 ' + m.Location : '') +
            (x.notes ? '<br>📝 ' + x.notes : '') + (x.contact ? '<br>📞 ' + x.contact : '') + '<br><br>God bless your service 🙏<br>Marathon</div>'
        });
      } catch (e) { Logger.log(e); }
    });
  });
}
/* ======================= ABSENT LIST (every Monday morning) ======================= */
const ABSENT_TAB = 'Absent Sunday';
/** Rebuilds the "Absent Sunday" tab from the latest Sunday Meeting attendance. */
function absentReport() {
  const t = today_();
  const sessions = readTable_('Sessions').filter(x => x.Type === 'Sunday Meeting' && x.Date && x.Date <= t)
    .sort((a, b) => a.Date < b.Date ? 1 : -1);
  let sh = ss_().getSheetByName(ABSENT_TAB);
  if (!sh) sh = ss_().insertSheet(ABSENT_TAB);
  sh.clear();
  if (!sessions.length) { sh.getRange(1, 1).setValue('No Sunday meeting yet'); return; }
  const last = sessions[0];
  const att = readTable_('Attendance');
  const bySession = {};
  att.forEach(a => { (bySession[a.SessionID] = bySession[a.SessionID] || {})[a.MemberID] = a.Status || 'Present'; });
  const here = bySession[last.ID] || {};
  const came = st => st === 'Present' || st === 'Late';
  const members = readTable_('Members').filter(m => m.Status !== 'Inactive' && m.Name && (m.JoinDate || '') <= last.Date);
  const rows = members.filter(m => !came(here[m.ID])).map(m => {
    let streak = 0;
    for (const s of sessions) { if ((m.JoinDate || '') > s.Date) break; if (came((bySession[s.ID] || {})[m.ID])) break; streak++; }
    const lastSeen = sessions.filter(s => came((bySession[s.ID] || {})[m.ID]))[0];
    return [m.Name, m.Phone || '', splitList_(m.Teams).join(', '), here[m.ID] === 'Excused' ? 'Excused' : 'Absent',
      streak, lastSeen ? lastSeen.Date : '—'];
  }).sort((a, b) => (a[3] === b[3] ? b[4] - a[4] : (a[3] === 'Absent' ? -1 : 1)));
  const presentN = members.length - rows.length;
  sh.getRange(1, 1).setValue('Sunday meeting ' + last.Date + ' — ' + rows.length + ' did not attend (out of ' + members.length + ' active, ' + presentN + ' attended)')
    .setFontWeight('bold').setFontSize(12);
  sh.getRange(2, 1).setValue('Updated automatically every Monday morning · last update ' + now_()).setFontColor('#64748b');
  const head = ['#', 'Name', 'Phone', 'Teams', 'Status', 'Sundays missed in a row', 'Last attended'];
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#eef2ff');
  if (rows.length) {
    sh.getRange(5, 3, rows.length, 1).setNumberFormat('@');
    sh.getRange(5, 1, rows.length, head.length).setValues(rows.map((r, i) => [i + 1].concat(r)));
    rows.forEach((r, i) => { if (r[3] === 'Absent' && r[4] >= 3) sh.getRange(5 + i, 1, 1, head.length).setBackground('#fee2e2'); });
  }
  sh.setFrozenRows(4);
  sh.autoResizeColumns(1, head.length);
}
/** Makes sure the Monday-morning trigger exists (runs once, automatically). */
function ensureAbsentTrigger_() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('ABSENT_TRIGGER') === '1') return;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'absentReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('absentReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(tz_()).create();
  pr.setProperty('ABSENT_TRIGGER', '1');
  try { absentReport(); } catch (e) { Logger.log(e); }
}

/** Run ONCE from the editor: hourly bookings sync + evening reminder emails + creates the calendar. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (['autoSync', 'dailyReminder'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoSync').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyReminder').timeBased().atHour(Number(settings_().REMINDER_HOUR || 18)).everyDays(1).inTimezone(tz_()).create();
  calendar_();
  sh_('Members'); sh_('Signups');
  // put every upcoming mission on everybody's calendar
  readTable_('Missions').filter(m => m.Status !== 'Cancelled' && (m.EndDate || m.Date) >= today_()).forEach(m => calMission_(m.ID));
  Logger.log('Done: hourly sync + daily reminder at ' + (settings_().REMINDER_HOUR || 18) + ':00 + calendar ready.');
}

/* ======================= ANNOUNCEMENTS ======================= */
function activeAnnouncements_() {
  const t = today_();
  return readTable_('Announcements').filter(a => (a.Text || a.Images) && (!a.Until || a.Until >= t) && ['Pending', 'Rejected'].indexOf(a.Status) < 0)
    .sort((a, b) => a.CreatedAt < b.CreatedAt ? 1 : -1).map(a => ({ ID: a.ID, Text: a.Text, CreatedAt: a.CreatedAt, Images: a.Images || '' }));
}
function saveAnnouncement_(d) {
  const text = String(d.Text || '').trim();
  const cur = d.ID ? readTable_('Announcements').filter(a => a.ID === d.ID)[0] : null;
  const keep = cur ? splitList_(d.keepImages !== undefined ? d.keepImages : cur.Images) : [];
  if (!text && !(d.images || []).length && !keep.length) throw new Error('Write the announcement first');
  const imgs = (d.images || []).slice(0, 4).map((b, i) => saveImage_(b, 'announcement_' + Date.now() + '_' + i + '.jpg'));
  if (cur) {
    splitList_(cur.Images).filter(id => keep.indexOf(id) < 0).forEach(trashFile_);
    upsert_('Announcements', r => r.ID === d.ID, { Text: text, Until: d.Until || '', Images: keep.concat(imgs).join(',') });
  } else {
    writeObj_('Announcements', { ID: nextId_('Announcements', 'A-', 3), Text: text, Until: d.Until || '', CreatedAt: now_(), By: d.By || 'Leader', Status: d.By ? 'Pending' : 'Approved', Images: imgs.join(',') });
  }
  return readTable_('Announcements');
}
function deleteAnnouncement_(id) {
  readTable_('Announcements').filter(r => r.ID === id).forEach(r => splitList_(r.Images).forEach(trashFile_));
  deleteWhere_('Announcements', r => r.ID === id);
  return readTable_('Announcements');
}

/* ======================= APPROVALS ======================= */
/** Sign-ups start as "Pending". The service leader (admin app) or a lead of an approver team
 *  (Settings → APPROVER_TEAMS, default Training) confirms or declines them — choosing the lead gives the permission. */
function isApprover_(m) {
  if (!m) return false;
  const teams = splitList_(settings_().APPROVER_TEAMS || 'Training');
  return readTable_('Teams').some(t => teams.indexOf(t.Name) >= 0 && t.Lead === m.ID);
}
function checkApprover_(d) {
  const m = findMember_(d.key);
  if (!m || m.Status === 'Inactive' || !isApprover_(m)) throw new Error('Only the team lead can approve');
  return m;
}
/** Sunday-meeting attendance of one member over the last 4 meetings. */
function sundayStats_(memberId, joinDate, sessions, att) {
  const t = today_();
  sessions = sessions || readTable_('Sessions'); att = att || readTable_('Attendance');
  const suns = sessions.filter(x => x.Type === 'Sunday Meeting' && x.Date && x.Date <= t && (!joinDate || String(joinDate).slice(0, 10) <= x.Date))
    .sort((a, b) => a.Date < b.Date ? 1 : -1).slice(0, 4);
  const came = {}; att.forEach(a => { if (a.MemberID === memberId && (a.Status === 'Present' || a.Status === 'Late' || !a.Status)) came[a.SessionID] = 1; });
  let streak = 0; for (const x of suns) { if (came[x.ID]) break; streak++; }
  const missed = suns.filter(x => !came[x.ID]).length;
  return { of: suns.length, missed: missed, streak: streak, flag: suns.length >= 2 && (streak >= 2 || missed >= 3) };
}
function approvals_(d) {
  checkApprover_(d);
  const t = today_();
  const members = readTable_('Members'), signups = readTable_('Signups');
  const sessions = readTable_('Sessions'), att = readTable_('Attendance');
  return readTable_('Missions').filter(m => m.Status !== 'Cancelled' && (m.EndDate || m.Date) >= t)
    .sort((a, b) => a.Date < b.Date ? -1 : 1)
    .map(m => ({ ID: m.ID, Title: m.Title, Type: m.Type, Date: m.Date, EndDate: m.EndDate, Location: m.Location, SlotsNeeded: m.SlotsNeeded,
      people: signups.filter(s => s.MissionID === m.ID && ['Cancelled'].indexOf(s.Status) < 0).map(s => {
        const x = members.filter(k => k.ID === s.MemberID)[0] || {};
        return { MemberID: s.MemberID, Name: x.Name || s.MemberID, Phone: x.Phone || '', Teams: x.Teams || '', Status: s.Status, SignedAt: s.SignedAt, Sun: s.Status === 'Pending' ? sundayStats_(s.MemberID, x.JoinDate, sessions, att) : null };
      }) }));
}
function decide_(d) {
  const lead = checkApprover_(d);
  if (['Confirmed', 'Declined', 'Pending'].indexOf(d.Status) < 0) throw new Error('Bad status');
  const pred = r => r.MissionID === d.MissionID && r.MemberID === d.MemberID;
  if (!readTable_('Signups').some(pred)) throw new Error('Sign-up not found');
  upsert_('Signups', pred, { Status: d.Status, Note: d.Status + ' by ' + lead.Name });
  return approvals_(d);
}

/** Members of the announcer teams (Settings → ANNOUNCER_TEAMS, default Media) can post announcements from the sign-up page. */
function isAnnouncer_(m) {
  if (!m || m.Status === 'Inactive') return false;
  const teams = splitList_(settings_().ANNOUNCER_TEAMS || 'Media');
  return splitList_(m.Teams).some(t => teams.indexOf(t) >= 0);
}
function checkAnnouncer_(d) {
  const m = findMember_(d.key);
  if (!isAnnouncer_(m)) throw new Error('Only the media team can post announcements');
  return m;
}

/* ======================= PHOTOS ======================= */
/** Profile photos are saved in the Drive folder "Marathon Members Photos" (owner's Drive), viewable by link. */
function photosFolder_() {
  const it = DriveApp.getFoldersByName('Marathon Members Photos');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Marathon Members Photos');
}
function saveImage_(b64, name) {
  b64 = String(b64 || '').replace(/^data:image\/\w+;base64,/, '');
  if (!b64) throw new Error('No photo');
  if (b64.length > 3500000) throw new Error('Photo is too big');
  const file = photosFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}
function trashFile_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
function savePhoto_(m, b64) {
  b64 = String(b64 || '').replace(/^data:image\/\w+;base64,/, '');
  if (!b64) throw new Error('No photo');
  if (b64.length > 2500000) throw new Error('Photo is too big');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', m.ID + '_' + Date.now() + '.jpg');
  const file = photosFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (m.Photo) { try { DriveApp.getFileById(m.Photo).setTrashed(true); } catch (e) {} }
  upsert_('Members', r => r.ID === m.ID, { Photo: file.getId() });
  return { Photo: file.getId() };
}
/** Run once from the editor after pasting this version (gives the script permission to save photos). */
function authorizePhotos() { photosFolder_(); Logger.log('Photos folder ready'); }
