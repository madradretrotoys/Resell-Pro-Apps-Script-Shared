function apiCD_DebugLoad(drawer) {
  try {
    var res = apiCD_LoadToday(drawer);
    // Force a simple, serializable envelope so the client always sees something
    return { ok: true, typeof: typeof res, payload: res };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}


/************ CASH DRAWER (server APIs) ************/

// Sheet + timezone (you already have CASH_TZ = 'America/Denver')
//const CD_SHEET  = 'Cash Drawer Log';
const CD_TZ     = Session.getScriptTimeZone() || 'America/Denver';
const CD_SHEET  = 'Cash Drawer Log';

// Header order we will create/expect in the sheet
const CD_HEADERS = [
  'Timestamp','CountID','Date','Period','Drawer','User',
  'Pennies','Nickels','Dimes','Quarters','HalfDollars',
  'Ones','Twos','Fives','Tens','Twenties','Fifties','Hundreds',
  'CoinTotal','BillTotal','GrandTotal','Second','Notes'
];

// Ensure sheet exists and is formatted; do NOT keep two versions of this function.
function cdEnsureSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(CD_SHEET);
  if (!sh) sh = ss.insertSheet(CD_SHEET);

  // If row 1 doesn't match our headers, rewrite just the header row.
  const firstRow = sh.getRange(1,1,1,CD_HEADERS.length).getValues()[0] || [];
  let mismatch = false;
  for (let i = 0; i < CD_HEADERS.length; i++) {
    if ((firstRow[i] || '') !== CD_HEADERS[i]) { mismatch = true; break; }
  }
  if (mismatch) {
    sh.getRange(1,1,1,CD_HEADERS.length).setValues([CD_HEADERS]);
    sh.setFrozenRows(1);
  }

  // Friendly formats (we don’t rely on them for logic)
  sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm'); // Timestamp
  sh.getRange('C:C').setNumberFormat('@');             // Date as plain text (yyyy-mm-dd)
  sh.getRange('D:E').setNumberFormat('@');             // Period/Drawer text
  sh.getRange('S:U').setNumberFormat('$#,##0.00');     // CoinTotal, BillTotal, GrandTotal
  return sh;
}

function cdNormDate_(v) {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; // already yyyy-MM-dd
  const d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

// Normalize “OPENING/CLOSING” to OPEN/CLOSE
function cdNormPeriod_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (s === 'OPENING') return 'OPEN';
  if (s === 'CLOSING') return 'CLOSE';
  return s;
}


// Date helpers (local to store timezone)
function cdTodayKey_() {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  const now = new Date();
  const y = Utilities.formatDate(now, tz, 'yyyy');
  const m = Utilities.formatDate(now, tz, 'MM');
  const d = Utilities.formatDate(now, tz, 'dd');
  return `${y}-${m}-${d}`; // e.g., 2025-08-13
}
function cdNow_(fmt) {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  return Utilities.formatDate(new Date(), tz, fmt || "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// Build row object from sheet row array
function cdRowToObj_(rowArr) {
  if (!rowArr || rowArr.length < CD_HEADERS.length) return null;
  const obj = {};
  CD_HEADERS.forEach((h, i) => obj[h] = rowArr[i]);
  // Map to UI field names
  return {
    timestamp: obj.Timestamp,
    countId: obj.CountID,
    date: obj.Date,
    period: obj.Period,
    drawer: obj.Drawer,
    user: obj.User,
    c001: Number(obj.Pennies)||0,
    c005: Number(obj.Nickels)||0,
    c010: Number(obj.Dimes)||0,
    c025: Number(obj.Quarters)||0,
    c050: Number(obj.HalfDollars)||0,
    b1: Number(obj.Ones)||0,
    b2: Number(obj.Twos)||0,
    b5: Number(obj.Fives)||0,
    b10: Number(obj.Tens)||0,
    b20: Number(obj.Twenties)||0,
    b50: Number(obj.Fifties)||0,
    b100: Number(obj.Hundreds)||0,
    coinTotal: Number(obj.CoinTotal)||0,
    billTotal: Number(obj.BillTotal)||0,
    grandTotal: Number(obj.GrandTotal)||0,
    second: String(obj.Second||''),
    notes: String(obj.Notes||'')
  };
}

// Find latest rows for OPEN and CLOSE for given date + drawer (robust to sheet formats)
function cdFindForDate_(dateKey, drawerNum) {
  const sh = cdEnsureSheet_();
  const last = sh.getLastRow();
  if (last <= 1) return { open: null, close: null, rowOpen: -1, rowClose: -1 };

  const values = sh.getRange(2, 1, last - 1, CD_HEADERS.length).getValues();
  let foundOpen = null, foundClose = null, rOpen = -1, rClose = -1;

  const wantDate   = cdNormDate_(dateKey);
  const wantDrawer = String(drawerNum || '1').trim();

  for (let i = values.length - 1; i >= 0; i--) { // newest → oldest
    const row     = values[i];
    const vDate   = cdNormDate_(row[2]);             // column C
    const vPeriod = cdNormPeriod_(row[3]);           // column D
    const vDrawer = String(row[4] || '').trim();     // column E

    if (vDate === wantDate && vDrawer === wantDrawer) {
      if (vPeriod === 'OPEN'  && !foundOpen ) { foundOpen  = cdRowToObj_(row);  rOpen  = i + 2; }
      if (vPeriod === 'CLOSE' && !foundClose) { foundClose = cdRowToObj_(row);  rClose = i + 2; }
      if (foundOpen && foundClose) break;
    }
  }
  return { open: foundOpen, close: foundClose, rowOpen: rOpen, rowClose: rClose };
}



// Public: health check (used by "Check connection")
function apiCD_Ping() {
  return { ok: true, now: cdNow_(), tz: (CASH_TZ || Session.getScriptTimeZone() || 'America/Denver') };
}

// Public: load today’s data for a drawer
function apiCD_LoadToday(drawer) {
  const logs = [];
  try {
    const dateKey = cdTodayKey_();
    logs.push(`dateKey=${dateKey}`);

    const drv = String(drawer||'1').trim();
    logs.push(`drawer=${drv}`);

    const res = cdFindForDate_(dateKey, drv);
    logs.push(`rowOpen=${res.rowOpen}, rowClose=${res.rowClose}`);

    // Make JSON-safe (Dates → ISO strings)
    ['open','close'].forEach(k => {
      const r = res[k];
      if (r && r.timestamp instanceof Date) r.timestamp = r.timestamp.toISOString();
    });

    return { open: res.open, close: res.close, logs };
  } catch (err) {
    logs.push('ERR: ' + (err && err.message ? err.message : String(err)));
    return { open: null, close: null, error: (err && err.message) ? err.message : String(err), logs };
  }
}

// Public: save (upsert) today’s count for a drawer & period
function apiCD_Save(payload) {
  const logs = [];
  try {
    const dateKey = cdTodayKey_();
    const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
    const now = new Date();
    const stamp = Utilities.formatDate(now, tz, 'M/d/yyyy H:mm');

    const p = payload || {};
    const drawer = String(p.drawer||'1').trim();
    const period = String(p.period||'OPEN').toUpperCase();
    logs.push(`save drawer=${drawer} period=${period} dateKey=${dateKey}`);

    const sh = cdEnsureSheet_();
    const found = cdFindForDate_(dateKey, drawer);
    const isOpen = (period === 'OPEN');
    const targetRow = isOpen ? found.rowOpen : found.rowClose;
    // If a row already exists for today+drawer+period, only allow managers to change it
    const perm = cdGetUserPerm_(p.user, p.loginId);
    const canLoad = !!(perm && perm.notify === 'Y'); // managers = notifications 'Y'
    if (targetRow > 0 && !canLoad) {
      logs.push('blocked: non-manager attempted to modify an existing count');
      return { ok:false, error:'Edits are restricted to managers. Please contact a manager.', logs };
    }
    const coinTotal =
    (Number(p.c001)||0)*0.01 +
    (Number(p.c005)||0)*0.05 +
    (Number(p.c010)||0)*0.10 +
    (Number(p.c025)||0)*0.25 +
    (Number(p.c050)||0)*0.50;

    const billTotal =
      (Number(p.b1)||0)*1 +
      (Number(p.b2)||0)*2 +
      (Number(p.b5)||0)*5 +
      (Number(p.b10)||0)*10 +
      (Number(p.b20)||0)*20 +
      (Number(p.b50)||0)*50 +
      (Number(p.b100)||0)*100;

    const grandTotal = coinTotal + billTotal;
    const countId = `${dateKey}#${drawer}#${period}`;

    const rowVals = [
      now,                     // Timestamp (actual Date object; sheet formats it)
      countId,                 // CountID
      dateKey,                 // Date
      period,                  // Period
      drawer,                  // Drawer
      p.user||'',              // User
      Number(p.c001)||0,       // Pennies
      Number(p.c005)||0,       // Nickels
      Number(p.c010)||0,       // Dimes
      Number(p.c025)||0,       // Quarters
      Number(p.c050)||0,       // HalfDollars
      Number(p.b1)||0,         // Ones
      Number(p.b2)||0,         // Twos
      Number(p.b5)||0,         // Fives
      Number(p.b10)||0,        // Tens
      Number(p.b20)||0,        // Twenties
      Number(p.b50)||0,        // Fifties
      Number(p.b100)||0,       // Hundreds
      coinTotal,               // CoinTotal
      billTotal,               // BillTotal
      grandTotal,              // GrandTotal
      p.second||'',            // Second
      p.notes||''              // Notes
    ];

    if (targetRow > 0) {
      sh.getRange(targetRow, 1, 1, CD_HEADERS.length).setValues([rowVals]);
      logs.push(`updated row ${targetRow}`);
    } else {
      sh.appendRow(rowVals);
      logs.push('appended new row');
    }

    return {
      ok: true,
      countId,
      date: dateKey,
      period,
      drawer,
      grandTotal,
      logs
    };
  } catch (err) {
    logs.push('ERR: ' + (err && err.message ? err.message : String(err)));
    return { ok: false, error: (err && err.message) ? err.message : String(err), logs };
  }
}

// Optional: existing debug wrapper can remain
function apiCD_DebugLoad(drawer) {
  try {
    var res = apiCD_LoadToday(drawer);
    return { ok: true, typeof: typeof res, payload: res };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/************ END CASH DRAWER (server APIs) ************/

/************ CASH DRAWER NOTIFICATIONS ************/

// Which drawers to check (expand if you ever add more drawers)
const CD_DRAWERS_TO_CHECK = ['1'];

// Try several common tab names for the permissions sheet
function cdGetUserNotifyRecipients_() {
  const ss = book_();
  const candidateTabs = ['User Permissions', 'Users', 'User'];
  let sh = null;
  for (const name of candidateTabs) { sh = ss.getSheetByName(name); if (sh) break; }
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idxUser  = headers.indexOf('User');
  const idxEmail = headers.indexOf('Notification Email');
  const idxFlag  = headers.indexOf('Cash Drawer Notifications');
  if (idxEmail === -1 || idxFlag === -1) return [];

  const out = [];
  for (const r of values) {
    const yn = String(r[idxFlag] || '').trim().toUpperCase();
    const email = String(r[idxEmail] || '').trim();
    const user  = (idxUser >= 0) ? String(r[idxUser] || '').trim() : '';
    if (yn === 'Y' && email) out.push({ user, email });
  }
  return out;
}

// Check today's OPEN/CLOSE presence for a drawer
function cdHasCount_(period, drawer) {
  const dateKey = cdTodayKey_();
  const res = cdFindForDate_(dateKey, String(drawer || '1'));
  return (String(period).toUpperCase() === 'OPEN') ? !!res.open : !!res.close;
}

// Build the drawer page link if we can (fallback is empty string)
function cdExecUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// Send an email only if the requested count is missing
function cdNotifyIfMissing_(period) {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  const dateKey = cdTodayKey_();
  const periodLabel = (String(period).toUpperCase() === 'OPEN') ? 'Opening' : 'Closing';
  const dueTime = (periodLabel === 'Opening') ? '12:00 PM' : '6:30 PM';
  const execUrl = cdExecUrl_();

  const recipients = cdGetUserNotifyRecipients_();
  if (!recipients.length) return { sent: 0, reason: 'no-recipients' };

  let sent = 0;
  for (const drv of CD_DRAWERS_TO_CHECK) {
    const hasIt = cdHasCount_(period, drv);
    if (hasIt) continue; // present -> no email

    const subj = `[Mad Rad] Missing ${periodLabel} Cash Drawer Count — ${dateKey} (Drawer ${drv})`;
    const body =
`Heads up:

No ${periodLabel.toLowerCase()} cash drawer count was found for today (${dateKey}) on Drawer ${drv} by ${dueTime} Mountain Time.

Please record it here:
${execUrl ? execUrl + '?page=drawer' : 'Open: Mad Rad Tools → Drawer'}

— Mad Rad Tools`;
    for (const r of recipients) {
      try { MailApp.sendEmail(r.email, subj, body); sent++; } catch (_) {}
    }
  }
  return { sent, dateKey, period: String(period).toUpperCase() };
}

// Public API for the UI to poll closing reminder status
function apiCD_CheckClosingReminder(drawer) {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  const now = new Date();
  const hh = Number(Utilities.formatDate(now, tz, 'H'));
  const mm = Number(Utilities.formatDate(now, tz, 'm'));

  const afterReminder = (hh > 18) || (hh === 18 && mm >= 0); // after 6:00 PM MT
  const res = cdFindForDate_(cdTodayKey_(), String(drawer || '1'));
  const missingClose = !res.close;

  return {
    afterReminder,
    missingClose,
    now: Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"),
  };
}

// CRON entrypoints (call these from time-based triggers)
function cdCronCheckOpening() { cdNotifyIfMissing_('OPEN'); }
function cdCronCheckClosing() { cdNotifyIfMissing_('CLOSE'); }

// One-time installer for the daily triggers (runs in MT if your project timezone is America/Denver)
function cdInstallDrawerTriggers() {
  // Clean old ones
  for (const t of ScriptApp.getProjectTriggers()) {
    if (['cdCronCheckOpening','cdCronCheckClosing'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  }
  // Opening: ~12:00 PM
  ScriptApp.newTrigger('cdCronCheckOpening').timeBased().atHour(12).nearMinute(0).everyDays(1).create();
  // Closing: ~6:30 PM
  ScriptApp.newTrigger('cdCronCheckClosing').timeBased().atHour(18).nearMinute(30).everyDays(1).create();
  return 'Triggers installed';
}
/************ END CASH DRAWER NOTIFICATIONS ************/

/************ CASH DRAWER ACCESS CONTROL ************/
// Look up the logged-in user in the "User Permissions" sheet
function cdGetUserPerm_(userName, loginId) {
  const ss = book_();
  const candidates = ['User Permissions','Users','User']; // try common tab names
  let sh = null;
  for (const name of candidates) { sh = ss.getSheetByName(name); if (sh) break; }
  if (!sh) return null;

  const rows = sh.getDataRange().getValues();
  const hdr  = rows.shift();
  const iUser  = hdr.indexOf('User');
  const iLogin = hdr.indexOf('Login ID');
  const iEmail = hdr.indexOf('Notification Email');
  const iFlag  = hdr.indexOf('Cash Drawer Notifications');
  if (iFlag === -1) return null;

  const wantUser  = String(userName||'').trim().toLowerCase();
  const wantLogin = String(loginId||'').trim().toLowerCase();

  for (const r of rows) {
    const u  = iUser  >= 0 ? String(r[iUser] || '').trim().toLowerCase()  : '';
    const li = iLogin >= 0 ? String(r[iLogin]|| '').trim().toLowerCase()  : '';
    if ((wantUser && u === wantUser) || (wantLogin && li === wantLogin)) {
      return {
        user:  iUser  >= 0 ? r[iUser]  : '',
        login: iLogin >= 0 ? r[iLogin] : '',
        email: iEmail >= 0 ? r[iEmail] : '',
        notify: String(r[iFlag] || '').trim().toUpperCase() // 'Y' or 'N'
      };
    }
  }
  return null;
}

// UI helper: return whether this user may "Load Today"
function apiCD_CanLoad(userName, loginId) {
  const perm = cdGetUserPerm_(userName, loginId);
  return { canLoad: !!(perm && perm.notify === 'Y'), notify: perm ? perm.notify : '' };
}
/************ END CASH DRAWER ACCESS CONTROL ************/

/* ====== CASH PAYOUTS (matches your headers) ====== */
const PAYOUT_SHEET = 'Cash Payouts';

function payoutSheet_() {
  // Uses your existing helper that opens the bound spreadsheet and gets a tab by name
  const sh = sheet_(PAYOUT_SHEET);
  ensurePayoutHeadersIfEmpty_(sh);
  ensureHelperColumns_(sh); // adds Payout ID + Deleted if missing
  return sh;
}

// Only set headers if the sheet is empty (won’t override your existing row 1)
function ensurePayoutHeadersIfEmpty_(sh) {
  if (sh.getLastRow() < 1 || sh.getRange(1,1,1,1).getValue() === '') {
    const desired = [
      'TimeStamp','Item Purchased','Short Description',
      'Drawer 1 Amount','Drawer 2 Amount','Safe Amount',
      'Grand Total','User','Payout ID','Deleted'
    ];
    sh.getRange(1,1,1,desired.length).setValues([desired]);
    sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm');
    sh.getRange('D:G').setNumberFormat('$#,##0.00');
  }
}

function ensureHelperColumns_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  if (!headers.includes('payout id')) {
    sh.getRange(1, lastCol+1).setValue('Payout ID');
  }
  const lastCol2 = sh.getLastColumn();
  const headers2 = sh.getRange(1,1,1,lastCol2).getValues()[0].map(h => String(h).trim().toLowerCase());
  if (!headers2.includes('deleted')) {
    sh.getRange(1, lastCol2+1).setValue('Deleted');
  }
  sh.getRange('D:G').setNumberFormat('$#,##0.00');
}

function toNum_(x){ const n = parseFloat(String(x).replace(/[^\d.]/g,'')); return isNaN(n)?0:Math.max(0,n); }
function makePayoutId_(){ return 'PO' + Date.now(); }

// ===== Buy Ticket integration: bump "Payout Amount" on Buy Tickets =====
// Safely adds `delta` to the "Payout Amount" for the ticket referenced by
// the Item Purchased (expects a BUY-XXXXXXXX id, case-insensitive).
function btPayout_bumpPayoutAmount_(rawItemPurchased, delta) {
  try {
    // 1) Extract the TicketID (prefer BUY-XXXXXXXX; else treat whole string)
    const raw = (rawItemPurchased == null ? '' : String(rawItemPurchased)).trim();
    const m = raw.match(/\bBUY-[A-Z0-9]{8}\b/i);
    const ticketId = m ? m[0].toUpperCase() : (typeof _norm_ === 'function' ? _norm_(raw) : raw);
    if (!ticketId) return false;

    // 2) Get Buy Tickets sheet and header
    const tsh = bt_ticketsSheet_(); // defined in buy_ticket.gs
    const last = tsh.getLastRow(); if (last <= 1) return false;
    const header = tsh.getRange(1, 1, 1, tsh.getLastColumn()).getValues()[0];

    // Column indices by header name (case-insensitive)
    const norm = s => String(s || '').trim().toLowerCase();
    const iTicket  = header.findIndex(h => norm(h) === 'ticketid');
    const iPayout  = header.findIndex(h => norm(h) === 'cash drawer payout amount');
    const iUpdated = header.findIndex(h => norm(h) === 'updatedat');

    if (iTicket < 0) return false;

    // 3) Find the ticket row
    const ids = tsh.getRange(2, iTicket + 1, last - 1, 1).getValues().flat().map(v => String(v || '').trim());
    const idx = ids.indexOf(ticketId);
    if (idx < 0) return false;
    const row = 2 + idx;

    // Default to column N (14) if header not present (you said column N)
    const colPayout = (iPayout >= 0 ? (iPayout + 1) : 14);

    // 4) Compute new value and write
    const cur = tsh.getRange(row, colPayout).getValue();
    const curNum = (function toNum(x) {
      const n = Number(x); 
      if (Number.isFinite(n)) return n;
      const m = String(x||'').replace(/[^\d.-]/g,''); 
      const p = parseFloat(m);
      return Number.isFinite(p) ? p : 0;
    })(cur);

    const d = Number(delta);
    if (!Number.isFinite(d) || d === 0) return true; // nothing to change

    const next = +(curNum + d).toFixed(2);
    tsh.getRange(row, colPayout).setValue(next);

    // Touch UpdatedAt if present
    if (iUpdated >= 0) {
      const stamp = (typeof _nowIso_ === 'function') ? _nowIso_() : new Date();
      tsh.getRange(row, iUpdated + 1).setValue(stamp);
    }
    return true;
  } catch (e) {
    // Keep payout flow resilient
    Logger.log('[btPayout_bumpPayoutAmount_] ' + (e && e.message ? e.message : e));
    return false;
  }
}


function apiPayout_Save(payload){
  const sh = payoutSheet_();
  const item = String(payload && payload.item || '').trim();
  if (!item) throw new Error('Item Purchased is required');

  const store   = toNum_(payload && payload.store);
  const nichols = toNum_(payload && payload.nichols);
  const safe    = toNum_(payload && payload.safe);
  const total   = +(store + nichols + safe).toFixed(2);
  if (total <= 0) throw new Error('Enter at least one amount > $0');

  const user = String(payload && payload.user || '').trim();
  const id = makePayoutId_();
  const ts = new Date();

  // A..J = TimeStamp, Item, Desc, D1, D2, Safe, Total, User, ID, Deleted
  sh.appendRow([ts, item, (payload.desc||''), store, nichols, safe, total, user, id, '']);
  const lr = sh.getLastRow();
  sh.getRange(lr, 1).setNumberFormat('m/d/yyyy h:mm');
  sh.getRange(lr, 4, 1, 4).setNumberFormat('$#,##0.00');

  btPayout_bumpPayoutAmount_(item, total);
  
  const displayTime = Utilities.formatDate(ts, CASH_TZ, 'M/d/yyyy h:mm a');
  return { id, item, desc: payload.desc||'', store, nichols, safe, total, user, displayTime };
}

function apiPayout_LoadLast(){
  const sh = payoutSheet_();
  const last = sh.getLastRow();
  if (last <= 1) return null;
  const vals = sh.getRange(2,1,last-1,10).getValues(); // A:J
  for (let i = vals.length - 1; i >= 0; i--){
    const r = vals[i];
    const deleted = String(r[10-1] || '').toUpperCase() === 'Y'; // J
    if (deleted) continue;
    const ts = r[1-1], item=r[2-1], desc=r[3-1],
          store=r[4-1]||0, nichols=r[5-1]||0, safe=r[6-1]||0,
          total=r[7-1]||0, user=r[8-1]||'', id=r[9-1]||'';
    const displayTime = Utilities.formatDate(new Date(ts), CASH_TZ, 'M/d/yyyy h:mm a');
    return { id, item, desc, store, nichols, safe, total, user, displayTime };
  }
  return null;
}

function apiPayout_Update(payload){
  if (!payload || !payload.id) throw new Error('Missing payout ID');
  const sh = payoutSheet_();
  const last = sh.getLastRow(); if (last <= 1) throw new Error('No payouts found');

  const ids = sh.getRange(2, 9, last-1, 1).getValues().flat(); // Col I = Payout ID
  const idx = ids.indexOf(payload.id);
  if (idx < 0) throw new Error('Payout not found: ' + payload.id);

  const row = 2 + idx;
  const prevTotal = Number(sh.getRange(row, 7).getValue()) || 0; // Col G = Grand Total

  const item = String(payload.item||'').trim();
  if (!item) throw new Error('Item Purchased is required');
  const store   = toNum_(payload.store);
  const nichols = toNum_(payload.nichols);
  const safe    = toNum_(payload.safe);
  const total   = +(store + nichols + safe).toFixed(2);
  if (total <= 0) throw new Error('Enter at least one amount > $0');
  const user = String(payload.user||'').trim();

  // Update B.H (keeps original timestamp and ID; clears Deleted)
  sh.getRange(row, 2, 1, 7).setValues([[item, (payload.desc||''), store, nichols, safe, total, user]]);
  sh.getRange(row, 10).setValue(''); // Deleted
  sh.getRange(row, 4, 1, 4).setNumberFormat('$#,##0.00');

  // NEW: bump by the delta so edits don’t double-count
  btPayout_bumpPayoutAmount_(item, total - prevTotal);

  const ts = sh.getRange(row,1).getValue();
  const displayTime = Utilities.formatDate(new Date(ts), CASH_TZ, 'M/d/yyyy h:mm a');
  return { id: payload.id, item, desc: payload.desc||'', store, nichols, safe, total, user, displayTime };

}

function apiPayout_Delete(id){
  if (!id) throw new Error('Missing payout ID');
  const sh = payoutSheet_();
  const last = sh.getLastRow(); if (last <= 1) return true;
  const ids = sh.getRange(2, 9, last-1, 1).getValues().flat(); // I
  const idx = ids.indexOf(id);
  if (idx < 0) return true;
  const row = 2 + idx;
  // Grab values before we mark as deleted
  const itemPurchased = String(sh.getRange(row, 2).getValue() || '').trim(); // Col B = Item Purchased
  const prevTotal     = Number(sh.getRange(row, 7).getValue()) || 0;         // Col G = Grand Total
  sh.getRange(row, 10).setValue('Y'); // J = Deleted
  // Subtract the deleted payout from Buy Tickets → Payout Amount
  if (prevTotal !== 0) btPayout_bumpPayoutAmount_(itemPurchased, -prevTotal);
  return true;
}

function apiDebug_SheetIdentity() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return { id: SPREADSHEET_ID, name: ss.getName(), url: ss.getUrl() };
}
/* ====== end CASH PAYOUTS  ====== */