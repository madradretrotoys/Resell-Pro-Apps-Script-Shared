/** ========= Time Entry (Timesheet) ========= **/

function getTimeSheetSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('TIME_ENTRY_SHEET');
  return SpreadsheetApp.openById(id);
}

function ensureSheets_() {
  const ss = getTimeSheetSpreadsheet_();
  const names = ss.getSheets().map(s => s.getName());
  if (!names.includes('Time Entries')) {
    ss.insertSheet('Time Entries').appendRow([
      'Date','User','LoginID','Clock In','Lunch Out','Lunch In','Clock Out',
      'Total Hours','Notes','Status','Edited By','Edited At','Period Start','Period End','RowID'
    ]);
  }
  if (!names.includes('Audit')) {
    ss.insertSheet('Audit').appendRow([
      'AuditID','RowID','When','Who','Action','Field','Old','New','Client','Notes'
    ]);
  }
}

/** Session → user (same store as POS) */
function getUserFromToken_(token){
  const p = loadSession_(token); // provided by your shared code.gs
  if (!p) throw new Error('Session expired');
  return { name: p.user, login: p.login, role: p.role };
}

/** Formatting helpers (America/Denver) */
function formatDate_(d) { return Utilities.formatDate(d, 'America/Denver', 'yyyy-MM-dd'); }  // yyyy-MM-dd
function formatTime_(d) { return Utilities.formatDate(d, 'America/Denver', 'hh:mm a'); }     // HH:MM AM/PM

/** Normalize Date column to yyyy-MM-dd key (prevents Date vs string mismatches) */
function _dateKey_(cell) {
  if (cell instanceof Date) return formatDate_(cell);
  if (!cell) return '';
  return formatDate_(new Date(cell));
}

/** Compute bi-weekly period (Sun..Sat x 2) using PAY_PERIOD_ANCHOR */
function computePeriod_(d) {
  const anchorStr = PropertiesService.getScriptProperties().getProperty('PAY_PERIOD_ANCHOR');
  const anchor = new Date(anchorStr + 'T00:00:00-07:00'); // America/Denver
  const days = Math.floor((d - anchor) / 86400000);
  const periodIndex = Math.floor(days / 14);
  const start = new Date(anchor.getTime() + periodIndex * 14 * 86400000);
  const end   = new Date(start.getTime() + 13 * 86400000);
  return { start: formatDate_(start), end: formatDate_(end) };
}

/** Hours math (safe if any stamp is blank) */
function calcHours_(clockIn, lunchOut, lunchIn, clockOut) {
  try {
    if (!clockIn || !clockOut) return '';
    const inT = new Date(clockIn);
    const outT = new Date(clockOut);
    let ms = outT - inT;
    if (lunchOut && lunchIn) ms -= (new Date(lunchIn) - new Date(lunchOut));
    return (ms / 3600000).toFixed(2);
  } catch (_) { return ''; }
}

/**
 * Parse "HH:MM AM/PM" at local wall time on dateKey (yyyy-MM-dd)
 * using the PROJECT TIME ZONE (America/Denver).
 */
function parseLocalTimeOnDate_(dateKey, hhmmAP) {
  const t = String(hhmmAP || '').trim();
  if (!t) return null;
  const m = t.match(/^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/);
  if (!m) throw new Error('Invalid time: ' + t + ' (use HH:MM AM/PM)');

  let hh = Number(m[1]), mm = Number(m[2]);
  const isPM = /pm$/i.test(m[3]);
  if (hh < 1 || hh > 12 || mm < 0 || mm > 59) throw new Error('Invalid time: ' + t);
  if (hh === 12) hh = 0;
  if (isPM) hh += 12;

  const parts = String(dateKey).split('-'); // yyyy-MM-dd
  const y = Number(parts[0]), mo = Number(parts[1]) - 1, d = Number(parts[2]);

  // IMPORTANT: constructs Date in project TZ (America/Denver), so 10:56 PM stays 10:56 PM.
  return new Date(y, mo, d, hh, mm, 0, 0);
}

/** Map row -> UI object */
function mapRow_(r){
  if (!r) return null;
  return {
    date: _dateKey_(r[0]),
    user: String(r[1] || ''),
    login: String(r[2] || ''),
    clockIn:  r[3] ? formatTime_(new Date(r[3])) : '',
    lunchOut: r[4] ? formatTime_(new Date(r[4])) : '',
    lunchIn:  r[5] ? formatTime_(new Date(r[5])) : '',
    clockOut: r[6] ? formatTime_(new Date(r[6])) : '',
    totalHours: r[7] || '',
    status: String(r[9] || 'Open')
  };
}

/** -------- APIs (all expect {token}) -------- */

function apiTime_GetToday(obj) {
  ensureSheets_();
  const user = getUserFromToken_(obj && obj.token);

  const sh = getTimeSheetSpreadsheet_().getSheetByName('Time Entries');
  const todayKey = formatDate_(new Date());
  const data = sh.getDataRange().getValues();

  let idx = -1;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (_dateKey_(r[0]) === todayKey && String(r[2]).trim() === String(user.login).trim()) { idx = i; break; }
  }

  let row;
  if (idx >= 0) {
    row = data[idx];
  } else {
    const period = computePeriod_(new Date());
    const rowId  = Utilities.getUuid();
    const now = new Date();
    const midnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // real Date cell
    row = [ midnight, user.name, user.login, '', '', '', '', '', '', 'Open', user.name, new Date(), period.start, period.end, rowId ];
    sh.appendRow(row);
  }
  return mapRow_(row);
}

function apiTime_Punch(obj) {
  ensureSheets_();
  const user = getUserFromToken_(obj && obj.token);

  const sh = getTimeSheetSpreadsheet_().getSheetByName('Time Entries');
  const todayKey = formatDate_(new Date());
  const data = sh.getDataRange().getValues();

  let idx = -1;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (_dateKey_(r[0]) === todayKey && String(r[2]).trim() === String(user.login).trim()) { idx = i; break; }
  }
  if (idx < 0) throw new Error('No row for today');

  const row = data[idx];

  let fieldIndex = -1;
  if (obj.type === 'CLOCK_IN')  fieldIndex = 3;
  if (obj.type === 'LUNCH_OUT') fieldIndex = 4;
  if (obj.type === 'LUNCH_IN')  fieldIndex = 5;
  if (obj.type === 'CLOCK_OUT') fieldIndex = 6;
  if (fieldIndex === -1) throw new Error('Invalid type');

  row[fieldIndex] = new Date();
  row[7]  = calcHours_(row[3], row[4], row[5], row[6]);
  row[9]  = (row[3] && row[6]) ? 'Complete' : 'Open';
  row[10] = user.name;
  row[11] = new Date();

  sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  return mapRow_(row);
}

function apiTime_ListMy(obj) {
  ensureSheets_();
  const user = getUserFromToken_(obj && obj.token);

  const sh = getTimeSheetSpreadsheet_().getSheetByName('Time Entries');
  const period = computePeriod_(new Date());

  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[2]).trim() !== String(user.login).trim()) continue;
    const k = _dateKey_(r[0]);
    if (k >= period.start && k <= period.end) out.push(mapRow_(r));
  }
  return out;
}

/**
 * Edit an existing row (employee: same-day & self; manager/admin: any date/user).
 * payload: { token, dateKey:'yyyy-MM-dd', login, fields:{clockIn, lunchOut, lunchIn, clockOut}, note? }
 */
function apiTime_EditRow(payload){
  ensureSheets_();
  const caller = getUserFromToken_(payload && payload.token);
  const targetLogin = String(payload.login || '').trim();
  const dateKey = String(payload.dateKey || '').trim();
  if (!targetLogin || !dateKey) throw new Error('Login and date are required.');

  const todayKey = formatDate_(new Date());
  const isMgr = (caller.role === 'Admin' || caller.role === 'Manager');
  if (!isMgr) {
    if (caller.login !== targetLogin) throw new Error('You can only edit your own time.');
    if (dateKey !== todayKey) throw new Error('Employees may edit only the current day.');
  }

  const sh = getTimeSheetSpreadsheet_().getSheetByName('Time Entries');
  const data = sh.getDataRange().getValues();
  let idx = -1;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (_dateKey_(r[0]) === dateKey && String(r[2]).trim() === targetLogin) { idx = i; break; }
  }
  if (idx < 0) throw new Error('Row not found.');

  const row = data[idx];
  const before = { cIn: row[3], lOut: row[4], lIn: row[5], cOut: row[6] };

  const f = (payload.fields || {});
  function setField(i, text){
    if (text == null) return;
    const s = String(text).trim();
    row[i] = s ? parseLocalTimeOnDate_(dateKey, s) : '';
  }
  setField(3, f.clockIn);
  setField(4, f.lunchOut);
  setField(5, f.lunchIn);
  setField(6, f.clockOut);

  row[7]  = calcHours_(row[3], row[4], row[5], row[6]);
  row[9]  = (row[3] && row[6]) ? 'Complete' : 'Open';
  row[10] = caller.name;
  row[11] = new Date();

  sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);

  // Audit trail (non-blocking)
  try {
    const audit = getTimeSheetSpreadsheet_().getSheetByName('Audit');
    const rowId = data[idx][14] || Utilities.getUuid();
    function addAudit(fieldLabel, oldV, newV){
      if (String(oldV||'') === String(newV||'')) return;
      audit.appendRow([
        Utilities.getUuid(),
        rowId,
        new Date(),
        caller.name + ' (' + caller.login + ')',
        isMgr ? 'MANAGER_EDIT' : 'EDIT',
        fieldLabel,
        oldV ? String(oldV) : '',
        newV ? String(newV) : '',
        'Resell Pro – Timesheet',
        String(payload.note || '')
      ]);
    }
    addAudit('Clock In', before.cIn, row[3]);
    addAudit('Lunch Out', before.lOut, row[4]);
    addAudit('Lunch In', before.lIn, row[5]);
    addAudit('Clock Out', before.cOut, row[6]);
  } catch (_) {}

  return mapRow_(row);
}
