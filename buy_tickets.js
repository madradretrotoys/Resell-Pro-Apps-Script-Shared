/************************************************************
 * Buy Ticket API (Cloud + Browser Logging, Safe Serialization)
 ************************************************************/

const BT_TICKETS_SHEET = 'Buy Tickets';
const BT_ITEMS_SHEET   = 'Buy Ticket Items';

const BT_TICKETS_HEADERS = [
  'TicketID','EstimateID','Status',
  'SellerName','SellerPhone','SellerEmail',
  'AskingSell','MeetInfo','Notes',
  'CreatedAt','UpdatedAt','Total','ClosedAt'
];

const BT_ITEMS_HEADERS = [
  'TicketID','ItemID','CreatedAt','UpdatedAt',
  'ShortTitle','LongDesc','PlannedSell','OfferAmount',
  'OfferSource','RemovedAt',
  // NEW
  'Qty','Total_Planned_Sell','Total_Offer'
];

function _norm_(v){ return String(v == null ? '' : v).trim(); }
function _nowIso_(){ return Utilities.formatDate(new Date(), CASH_TZ || Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }
function _makeTicketId_(){ return 'BUY-' + Utilities.getUuid().replace(/-/g,'').slice(-8).toUpperCase(); }

function _bt_num_(v){ const n = Number(v); return Number.isFinite(n) ? n : 0; }
function _bt_validatePlannedVsOffer_(planned, offer){
  const p = _bt_num_(planned), o = _bt_num_(offer);
  if (o > p) throw new Error('Offer cannot exceed the planned sell amount.');
}

// Deep sanitize any object/array so google.script.run can return it safely.
// - Date -> ISO string
// - NaN/undefined -> null
// - Only JSON primitives/arrays/objects
function _serializeForClient_(x){
  const t = Object.prototype.toString.call(x);
  if (x === null) return null;
  if (t === '[object Date]') return Utilities.formatDate(x, CASH_TZ || Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  if (t === '[object Number]') return Number.isFinite(x) ? x : null;
  if (t === '[object String]' || t === '[object Boolean]') return x;
  if (t === '[object Array]') return x.map(_serializeForClient_);
  if (t === '[object Object]'){
    const o = {};
    for (const k in x){
      if (!Object.prototype.hasOwnProperty.call(x,k)) continue;
      const v = x[k];
      if (v === undefined) { o[k] = null; continue; }
      o[k] = _serializeForClient_(v);
    }
    return o;
  }
  // fallback: stringify unknowns
  return x != null ? String(x) : null;
}

function _ensureSheet_(name, requiredHeaders){
  const logs = [];
  logs.push(`_ensureSheet_: start, name="${name}"`);

  const ss = book_();
  logs.push(`_ensureSheet_: opened by book_(), id=${ss.getId()}`);

  let sh = ss.getSheetByName(name);
  if (!sh) {
    logs.push(`_ensureSheet_: creating sheet "${name}"`);
    sh = ss.insertSheet(name);
    sh.appendRow(requiredHeaders);
    Logger.log(logs.join('\n'));
    return sh;
  }

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  logs.push(`_ensureSheet_: lastRow=${lastRow}, lastCol=${lastCol}`);

  if (lastRow === 0 || lastCol === 0) {
    logs.push(`_ensureSheet_: sheet blank, writing headers`);
    sh.clear();
    sh.appendRow(requiredHeaders);
    Logger.log(logs.join('\n'));
    return sh;
  }

  const row1 = sh.getRange(1,1,1,lastCol).getValues()[0];
  const have = new Set(row1.map(h => _norm_(h)).filter(Boolean));
  const missing = requiredHeaders.filter(h => !have.has(_norm_(h)));

  if (missing.length) {
    logs.push(`_ensureSheet_: missing headers = [${missing.join(', ')}], inserting`);
    sh.insertColumnsAfter(lastCol, missing.length);
    sh.getRange(1, row1.length + 1, 1, missing.length).setValues([missing]);
  }

  Logger.log(logs.join('\n'));
  return sh;
}

function bt_ticketsSheet_(){ return _ensureSheet_(BT_TICKETS_SHEET, BT_TICKETS_HEADERS); }
function bt_itemsSheet_(){ return _ensureSheet_(BT_ITEMS_SHEET, BT_ITEMS_HEADERS); }

function _readAllAsObjects_(sh){
  const logs = [];
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  logs.push(`_readAllAsObjects_: sheet=${sh.getName()}, lastRow=${lastRow}, lastCol=${lastCol}`);
  if (lastRow < 1 || lastCol < 1) { Logger.log(logs.join('\n')); return { header:[], rows:[], objects:[], _logs:logs }; }

  const header = sh.getRange(1,1,1,lastCol).getValues()[0];
  logs.push(`[readAll] header=${header.join('|')}`);
  if (lastRow === 1) { Logger.log(logs.join('\n')); return { header, rows:[], objects:[], _logs:logs }; }

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  const objects = rows.map(r => { const o={}; header.forEach((h,i)=> o[String(h)] = r[i]); return o; });
  logs.push(`[readAll] objects=${objects.length}`);
  Logger.log(logs.join('\n'));
  return { header, rows, objects, _logs:logs };
}

function _dbgSampleCol_(sh, headerName){
  try{
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { count:0, list:[] };
    const header = sh.getRange(1,1,1,lastCol).getValues()[0];
    const idx = header.findIndex(h => _norm_(h) === _norm_(headerName));
    if (idx === -1) return { count:lastRow-1, list:[], headerFound:false, headerRow:_norm_(header.join('|')) };
    const n = Math.min(10, lastRow-1);
    const vals = sh.getRange(2, idx+1, n, 1).getValues().map(r => _serializeForClient_(r[0]));
    return { count:lastRow-1, list:vals, headerFound:true, colIndex:idx+1 };
  }catch(e){
    return { error:String(e && e.message ? e.message : e) };
  }
}

function _filterBy_(objects, key, val){
  const needle = _norm_(val);
  return objects.filter(o => _norm_(o[key]) === needle);
}

function _bt_readEstimate_(estId){
  const sh = est_sheet_();
  const row = est_findRowById_(estId);
  if (!row) return null;

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const vals   = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  const map = {}; header.forEach((h,i)=> map[String(h)] = vals[i]);

  const first  = _norm_(map.First);
  const last   = _norm_(map.Last);
  const phone  = _norm_(map.Phone);
  const email  = _norm_(map.Email);
  const askRaw = map.WouldLike;
  const asking = (askRaw===''||askRaw==null) ? '' : Number(String(askRaw).replace(/[^\d.]/g,''));

  return {
    sellerName: [first,last].filter(Boolean).join(' ').trim(),
    phone, email, asking
  };
}

// ------------------------------------------------------------------
// Create or get ticket
// ------------------------------------------------------------------
function apiBT_createOrGetForEstimate(estId, sellerOpt){
  const logs = [];
  try {
    const id = _norm_(estId);
    logs.push(`→ apiBT_createOrGetForEstimate start, estId=${id}`);
    if (!id) throw new Error('Missing EstimateID');

    const tsh  = bt_ticketsSheet_();
    logs.push(`[A] ticketsSheet name="${tsh.getName()}" lr=${tsh.getLastRow()} lc=${tsh.getLastColumn()}`);
    const tTab = _readAllAsObjects_(tsh);
    logs.push(...tTab._logs);

    const existing = _filterBy_(tTab.objects, 'EstimateID', id);
    logs.push(`[A] existing by EstimateID count=${existing.length}`);
    if (existing.length) {
      const out = { ok:true, created:false, ticketId: existing[0].TicketID, estId:id, _logs:logs };
      Logger.log(logs.join('\n'));
      return _serializeForClient_(out);
    }

    const fromEstimate = _bt_readEstimate_(id) || {};
    logs.push(`[A] fromEstimate keys=${Object.keys(fromEstimate||{}).join(',')}`);
    const sellerName = _norm_(sellerOpt && sellerOpt.name) || fromEstimate.sellerName || '';
    const sellerPhone = _norm_(sellerOpt && sellerOpt.phone) || fromEstimate.phone || '';
    const sellerEmail = _norm_(sellerOpt && sellerOpt.email) || fromEstimate.email || '';
    const askingSell  = (fromEstimate.asking===''||fromEstimate.asking==null) ? '' : Number(fromEstimate.asking);

    const now = _nowIso_();
    const ticketId = _makeTicketId_();

    const header = tsh.getRange(1,1,1,tsh.getLastColumn()).getValues()[0];
    logs.push(`[A] header cols=${header.length} :: ${header.join('|')}`);
    // AFTER (adds 'Asking Sell' alias so either header gets the value)
    const map = {
      TicketID: ticketId, EstimateID: id, Status: 'Open',
      SellerName: sellerName, SellerPhone: sellerPhone, SellerEmail: sellerEmail,
      AskingSell: askingSell, 'Asking Sell': askingSell,  // <-- alias
      MeetInfo: '', Notes: '',
      CreatedAt: now, UpdatedAt: now, Total: 0, ClosedAt: ''
    };
    const row = header.map(h => (h in map) ? map[h] : '');
    tsh.appendRow(row);
    logs.push(`[A] created new ticket ${ticketId}`);

    const out = { ok:true, created:true, ticketId, estId:id, _logs:logs };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? String(err.stack) : '';
    logs.push('ERROR: ' + msg);
    const out = { ok:false, error: msg, stack, _logs: logs, where:'apiBT_createOrGetForEstimate' };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);
  }
}

// ------------------------------------------------------------------
// Get ticket + items
// ------------------------------------------------------------------
function apiBT_getTicket(ticketId){
  const logs = [];
  try {
    const id = _norm_(ticketId);
    logs.push(`→ apiBT_getTicket start, ticketId=${id}`);
    if (!id) throw new Error('Missing TicketID');

    const tsh  = bt_ticketsSheet_();
    logs.push(`[B] ticketsSheet name="${tsh.getName()}" lr=${tsh.getLastRow()} lc=${tsh.getLastColumn()}`);

    const headerRangeExists = (tsh.getLastColumn() > 0);
    if (!headerRangeExists) throw new Error('[B] Buy Tickets sheet has 0 columns');

    const header = tsh.getRange(1,1,1,tsh.getLastColumn()).getValues()[0];
    logs.push(`[B] header cols=${header.length} :: ${header.join('|')}`);

    const lastRow = tsh.getLastRow();
    if (lastRow < 1) throw new Error('[B] Buy Tickets sheet is empty (lastRow<1)');
    if (lastRow === 1) {
      const sample = _dbgSampleCol_(tsh, 'TicketID');
      logs.push(`[B] header-only sheet, TicketID sample=${JSON.stringify(sample)}`);
      Logger.log(logs.join('\n'));
      throw new Error(`Ticket not found: ${id}`);
    }

    const sample = _dbgSampleCol_(tsh, 'TicketID');
    logs.push(`[B] TicketID sample=${JSON.stringify(sample)}`);

    const tTab = _readAllAsObjects_(tsh);
    logs.push(...tTab._logs);
    logs.push(`[B] objects count=${tTab.objects.length}`);

    let tMatches = _filterBy_(tTab.objects, 'TicketID', id);
    logs.push(`[B] exact matches=${tMatches.length}`);

    if (!tMatches.length) {
      const idStr = _norm_(String(id));
      tMatches = tTab.objects.filter(o => _norm_(String(o['TicketID'])) === idStr);
      logs.push(`[B] fallback stringify/trim matches=${tMatches.length}`);
    }

    if (!tMatches.length) {
      const firstTen = tTab.objects.slice(0,10).map(o => _serializeForClient_(o['TicketID']));
      logs.push(`[B] first10 TicketIDs=${JSON.stringify(firstTen)}`);
      Logger.log(logs.join('\n'));
      throw new Error(`Ticket not found: ${id}`);
    }

    const ticket = tMatches[0];
    logs.push('[B] ticket keys=' + Object.keys(ticket).join(','));
    // Normalize AskingSell so the UI can always render it
    (function normalizeAskingSell(o){
      // Prefer canonical key; otherwise look for common variants
      const v = o.AskingSell ?? o['Asking Sell'] ?? o.askingSell;
      // Optional safety: if nothing found but MeetInfo is a number, treat that as a fallback
      const meetAsNum = (o.MeetInfo == null) ? null : Number(String(o.MeetInfo).replace(/[^\d.]/g,''));
      o.AskingSell = (v != null && v !== '') ? v
                  : (Number.isFinite(meetAsNum) ? meetAsNum : '');
    })(ticket);


    const ish  = bt_itemsSheet_();
    logs.push(`[B] itemsSheet name="${ish.getName()}" lr=${ish.getLastRow()} lc=${ish.getLastColumn()}`);
    const iTab = _readAllAsObjects_(ish);
    logs.push(...iTab._logs);
    const items = _filterBy_(iTab.objects, 'TicketID', id);
    logs.push(`[B] items matched=${items.length}`);

    const out = { ok:true, ticket, items, _logs:logs };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? String(err.stack) : '';
    logs.push('ERROR: ' + msg);
    const out = { ok:false, error: msg, stack, _logs:logs, where:'apiBT_getTicket' };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);
  }
}

/**
 * Increment the "Payout Amount" for a ticket by `delta` and return the fresh ticket.
 * - Looks for a header named "Payout Amount" (case-insensitive). If not found, falls back to column N (14).
 * - Touches UpdatedAt.
 * - Returns: { ok, ticket, _logs }
 */
function apiBT_addPayoutAmount(ticketId, delta){
  return apiBT_bumpPayComponent(ticketId, 'cashDrawer', delta);
}

// ------------------------------------------------------------------
// Add item to a Buy Ticket and return fresh ticket+items
// ------------------------------------------------------------------
function apiBT_addItem(ticketId, item) {
  const logs = [];
  try {
    const id = _norm_(ticketId);
    logs.push(`→ apiBT_addItem start, ticketId=${id}`);

    if (!id) throw new Error('Missing TicketID');

    const tsh = bt_ticketsSheet_();
    const ish = bt_itemsSheet_();
    logs.push(`[I] ticketsSheet="${tsh.getName()}", itemsSheet="${ish.getName()}"`);

    // verify ticket exists
    const tTab = _readAllAsObjects_(tsh);
    const tMatch = _filterBy_(tTab.objects, 'TicketID', id);
    if (!tMatch.length) throw new Error(`Ticket not found: ${id}`);

    // sanitize incoming payload
    const now = _nowIso_();
    const itemId = 'ITM-' + Utilities.getUuid().replace(/-/g,'').slice(-8).toUpperCase();
    const shortTitle  = _norm_(item && item.shortTitle) || '(untitled)';
    const longDesc    = _norm_(item && item.longDesc)   || '';
    const plannedSell = Number(item && item.plannedSell);
    const offerAmount = Number(item && item.offerAmount);
    const offerSource = _norm_(item && item.offerSource) || '';
    const qtyIn       = Number(item && item.qty);
    const qty         = (Number.isFinite(qtyIn) && qtyIn > 0) ? Math.floor(qtyIn) : 1;

    _bt_validatePlannedVsOffer_(plannedSell, offerAmount);

    const iHeader = ish.getRange(1,1,1,ish.getLastColumn()).getValues()[0];

    // compute per-row totals
    const ps = Number.isFinite(plannedSell) ? plannedSell : 0;
    const of = Number.isFinite(offerAmount) ? offerAmount : 0;
    const totalPlanned = +(ps * qty).toFixed(2);
    const totalOffer   = +(of * qty).toFixed(2);

    // map for appendRow in exact header order
    const iMap = {
      TicketID: id,
      ItemID: itemId,
      CreatedAt: now,
      UpdatedAt: now,
      ShortTitle: shortTitle,
      LongDesc: longDesc,
      PlannedSell: ps,
      OfferAmount: of,
      OfferSource: offerSource,
      RemovedAt: '',
      // NEW
      Qty: qty,
      Total_Planned_Sell: totalPlanned,
      Total_Offer: totalOffer
    };
    const iRow = iHeader.map(h => (h in iMap) ? iMap[h] : '');
    ish.appendRow(iRow);
    logs.push(`[I] appended item ${itemId} (qty=${qty}, totals: planned=${totalPlanned}, offer=${totalOffer})`);

    // recalc ticket total (prefer Total_Offer; fallback to OfferAmount*Qty; fallback to OfferAmount)
    const iTab = _readAllAsObjects_(ish);
    const itemsForTicket = _filterBy_(iTab.objects, 'TicketID', id).filter(o => !o.RemovedAt);
    const newTotal = itemsForTicket.reduce((sum, o) => {
      const tOffer = Number(o.Total_Offer);
      if (Number.isFinite(tOffer)) return sum + tOffer;
      const q = Number(o.Qty); const of2 = Number(o.OfferAmount);
      if (Number.isFinite(of2) && Number.isFinite(q)) return sum + (of2 * (q > 0 ? q : 1));
      return sum + (Number.isFinite(of2) ? of2 : 0);
    }, 0);

    // write UpdatedAt and Total back to ticket row
    const tHeader = tTab.header;
    const tRowIdx = tTab.objects.findIndex(o => _norm_(o.TicketID) === id);
    if (tRowIdx >= 0) {
      const sheetRow = 2 + tRowIdx;
      const colUpdatedAt = tHeader.indexOf('UpdatedAt') + 1;
      const colTotal     = tHeader.indexOf('Total') + 1;
      if (colUpdatedAt > 0) tsh.getRange(sheetRow, colUpdatedAt).setValue(now);
      if (colTotal > 0)     tsh.getRange(sheetRow, colTotal).setValue(+newTotal.toFixed(2));
      logs.push(`[I] updated ticket totals (Total= ${newTotal.toFixed(2)})`);
    }

    const out = apiBT_getTicket(id);
    if (out && out._logs) out._logs.push(...logs);
    return out;
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? err.stack : '';
    const out = { ok:false, error: msg, stack, _logs: logs, where:'apiBT_addItem' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

/**
 * Set ticket status and return fresh ticket + items.
 * Allowed statuses: 'Open' | 'Paid' | 'Declined' | 'picked_up'
 */
function apiBT_setTicketStatus(ticketId, status){
  const logs = [];
  try {
    const id = _norm_(ticketId);
    const s  = String(status || '').trim();
    logs.push(`→ apiBT_setTicketStatus start, ticketId=${id}, status=${s}`);

    if (!id) throw new Error('Missing TicketID');
    const allowed = ['Open','Paid','Declined','picked_up'];
    if (allowed.indexOf(s) < 0) throw new Error('Invalid status: ' + s);

    const tsh = bt_ticketsSheet_();
    const ish = bt_itemsSheet_();
    logs.push(`[S] ticketsSheet="${tsh.getName()}", itemsSheet="${ish.getName()}"`);

    // Read tickets
    const tTab = _readAllAsObjects_(tsh);
    logs.push(...(tTab._logs || []));
    const tHeader  = tTab.header || [];
    const tObjects = tTab.objects || [];

    // Find row for this TicketID
    const rowIdx = tObjects.findIndex(o => _norm_(o.TicketID) === id);
    if (rowIdx < 0) throw new Error(`Ticket not found: ${id}`);
    const sheetRow = 2 + rowIdx; // header is row 1

    // Column indices
    const colStatus    = tHeader.indexOf('Status')    + 1;
    const colUpdatedAt = tHeader.indexOf('UpdatedAt') + 1;
    const colClosedAt  = tHeader.indexOf('ClosedAt')  + 1;

    const now = _nowIso_();

    if (colStatus > 0)    tsh.getRange(sheetRow, colStatus).setValue(s);
    if (colUpdatedAt > 0) tsh.getRange(sheetRow, colUpdatedAt).setValue(now);

    // ClosedAt rules
    if (colClosedAt > 0){
      if (s === 'picked_up')       tsh.getRange(sheetRow, colClosedAt).setValue(now);
      else if (s === 'Open')      tsh.getRange(sheetRow, colClosedAt).setValue('');
      // for 'Paid' or 'Declined', leave ClosedAt as-is
    }

    // Read fresh ticket + items to return
    const tTab2   = _readAllAsObjects_(tsh);
    const ticketA = _filterBy_(tTab2.objects, 'TicketID', id);
    if (!ticketA.length) throw new Error(`Ticket not found after update: ${id}`);
    const ticket  = ticketA[0];

    const iTab    = _readAllAsObjects_(ish);
    const items   = _filterBy_(iTab.objects, 'TicketID', id);

    logs.push(`[S] status set OK: ${s}; returning fresh ticket+items`);
    const out = { ok:true, ticket, items, _logs:logs };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);

  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? String(err.stack) : '';
    logs.push('ERROR: ' + msg);
    const out = { ok:false, error: msg, stack, _logs:logs, where:'apiBT_setTicketStatus' };
    Logger.log(logs.join('\n'));
    return _serializeForClient_(out);
  }
}


/**
 * Set Trade/Other to an absolute value, then recompute and persist "Total Pay Amount".
 * kind: 'trade' | 'other'
 */
function apiBT_setPayComponent(ticketId, kind, value){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Buy Tickets');
  if (!sh) return { ok:false, message:'Sheet not found' };

  const data = sh.getDataRange().getValues();
  const headers = data[0];

  const idxTicket = headers.indexOf('TicketID');
  const idxTrade  = headers.indexOf('Trade Amount');
  const idxOther  = headers.indexOf('Other Pay (zelle, cashapp, etc.)');
  const idxCash   = headers.indexOf('Cash Drawer Payout Amount');

  // Ensure Total Pay Amount column exists (this was the missing piece)
  let idxTotal  = headers.indexOf('Total Pay Amount');
  if (idxTotal < 0) {
    idxTotal = headers.length;
    sh.getRange(1, idxTotal + 1).setValue('Total Pay Amount');
  }

  if (idxTicket < 0) return { ok:false, message:'No TicketID column' };
  const row = data.findIndex(r => String(r[idxTicket]) === String(ticketId));
  if (row < 1) return { ok:false, message:'Ticket not found' };

  const r = row + 1; // 1-based
  const v = Number(value);
  const val = isFinite(v) && v >= 0 ? v : 0;

  if (String(kind) === 'trade' && idxTrade >= 0) sh.getRange(r, idxTrade+1).setValue(val);
  if (String(kind) === 'other' && idxOther >= 0) sh.getRange(r, idxOther+1).setValue(val);

  // Recompute total = cash + trade + other
  const toNum = x => {
    const n = Number(x); if (Number.isFinite(n)) return n;
    const s = String(x == null ? '' : x).replace(/[^\d.-]/g,'');
    const p = parseFloat(s); return Number.isFinite(p) ? p : 0;
  };
  const cash  = idxCash  >= 0 ? toNum(sh.getRange(r, idxCash +1).getValue())  : 0;
  const trade = idxTrade >= 0 ? toNum(sh.getRange(r, idxTrade+1).getValue()) : 0;
  const other = idxOther >= 0 ? toNum(sh.getRange(r, idxOther+1).getValue()) : 0;
  const total = +(cash + trade + other).toFixed(2);

  sh.getRange(r, idxTotal + 1).setValue(total);

  // Return fresh ticket row
  const rowVals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const ticket = {};
  hdrs.forEach((h, i) => ticket[h] = rowVals[i]);
  return { ok:true, ticket };
}

/**
 * Accumulate a payment component and recompute "Total Pay Amount".
 * kind: 'cashDrawer' | 'trade' | 'other'
 * delta: number to add (>=0 typically)
 * Returns { ok, ticket, _logs }
 */
function apiBT_bumpPayComponent(ticketId, kind, delta){
  const logs = [];
  try {
    const id = _norm_(ticketId);
    if (!id) throw new Error('Missing TicketID');
    const d = Number(delta);
    if (!Number.isFinite(d)) throw new Error('Invalid delta (not a number)');

    // Preferred headers per component; keep compatibility for cash drawer
    const MAP = {
      cashDrawer: ['Cash Drawer Payout Amount', 'Payout Amount'],
      trade:      ['Trade Amount'],
      other:      ['Other Pay (zelle, cashapp, etc.)']
    };
    const TOTAL_HEADER = 'Total Pay Amount';

    const tsh = bt_ticketsSheet_();
    const lastCol = tsh.getLastColumn();
    const header  = tsh.getRange(1,1,1,lastCol).getValues()[0];

    const norm = s => String(s||'').trim().toLowerCase();
    const findOrAppendHeader = (names) => {
      for (const n of names){
        const i = header.findIndex(h => norm(h) === norm(n));
        if (i >= 0) return i + 1;
      }
      // append preferred (first) if none found
      const newCol = tsh.getLastColumn() + 1;
      tsh.getRange(1, newCol).setValue(names[0]);
      return newCol;
    };

    // Locate ticket row
    const tTab   = _readAllAsObjects_(tsh);
    const objs   = tTab.objects || [];
    const rowIdx = objs.findIndex(o => _norm_(o.TicketID) === id);
    if (rowIdx < 0) throw new Error(`Ticket not found: ${id}`);
    const row = 2 + rowIdx;

    // Ensure component columns exist
    const colCash  = findOrAppendHeader(MAP.cashDrawer);
    const colTrade = findOrAppendHeader(MAP.trade);
    const colOther = findOrAppendHeader(MAP.other);
    // Ensure total column exists
    let colTotal;
    {
      const i = header.findIndex(h => norm(h) === norm(TOTAL_HEADER));
      if (i >= 0) colTotal = i + 1;
      else { colTotal = tsh.getLastColumn() + 1; tsh.getRange(1, colTotal).setValue(TOTAL_HEADER); }
    }

    // Route to target column
    const targetCol =
      (kind === 'cashDrawer') ? colCash :
      (kind === 'trade')      ? colTrade :
      (kind === 'other')      ? colOther : null;
    if (!targetCol) throw new Error('Invalid kind: ' + kind);

    const toNum = (x) => {
      const n = Number(x); if (Number.isFinite(n)) return n;
      const s = String(x==null?'':x).replace(/[^\d.-]/g,'');
      const p = parseFloat(s); return Number.isFinite(p) ? p : 0;
    };

    // Bump the component
    const cur = toNum(tsh.getRange(row, targetCol).getValue());
    const next = +(cur + d).toFixed(2);
    tsh.getRange(row, targetCol).setValue(next);

    // Recompute total = cash + trade + other
    const cash  = toNum(tsh.getRange(row, colCash).getValue());
    const trade = toNum(tsh.getRange(row, colTrade).getValue());
    const other = toNum(tsh.getRange(row, colOther).getValue());
    const total = +(cash + trade + other).toFixed(2);
    tsh.getRange(row, colTotal).setValue(total);

    // Touch UpdatedAt if present
    const iUpd = header.findIndex(h => norm(h) === 'updatedat');
    if (iUpd >= 0) tsh.getRange(row, iUpd + 1).setValue(_nowIso_());

    // Return fresh ticket so UI can repaint
    const tTab2  = _readAllAsObjects_(tsh);
    const ticket = (tTab2.objects || []).find(o => _norm_(o.TicketID) === id);
    return _serializeForClient_({ ok:true, ticket, _logs: logs });
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? String(err.stack) : '';
    return _serializeForClient_({ ok:false, error: msg, stack, _logs: [] });
  }
}


// ------------------------------------------------------------------
// Compatibility alias (older UI code sometimes calls apiBT_get)
// ------------------------------------------------------------------
function apiBT_get(ticketId) { return apiBT_getTicket(ticketId); }

// --------------------------------------------------------------
// Update a single Buy Ticket item, recalc totals, return fresh data
// patch = { shortTitle?, plannedSell?, offerAmount? }
// --------------------------------------------------------------
function apiBT_updateItem(ticketId, itemId, patch) {
  const logs = [];
  try {
    const tId = _norm_(ticketId);
    const iId = _norm_(itemId);
    if (!tId) throw new Error('Missing TicketID');
    if (!iId) throw new Error('Missing ItemID');

    logs.push(`→ apiBT_updateItem start t=${tId} i=${iId}`);

    const tsh = bt_ticketsSheet_();
    const ish = bt_itemsSheet_();

    // Verify ticket exists
    const tTab = _readAllAsObjects_(tsh);
    const tRowIdx = tTab.objects.findIndex(o => _norm_(o.TicketID) === tId);
    if (tRowIdx < 0) throw new Error('Ticket not found: ' + tId);

    // Find item row
    const iTab = _readAllAsObjects_(ish);
    const iRowIdx = iTab.objects.findIndex(o => _norm_(o.ItemID) === iId && _norm_(o.TicketID) === tId);
    if (iRowIdx < 0) throw new Error('Item not found for this ticket');

    const iHeader = iTab.header;
    const sheetRow = 2 + iRowIdx; // data starts on row 2
    const now = _nowIso_();

    // columns (existing)
    const colShort   = iHeader.indexOf('ShortTitle')  + 1;
    const colPlanned = iHeader.indexOf('PlannedSell') + 1;
    const colOffer   = iHeader.indexOf('OfferAmount') + 1;
    const colUpdated = iHeader.indexOf('UpdatedAt')   + 1;
    // NEW columns
    const colQty     = iHeader.indexOf('Qty')                + 1;
    const colTPS     = iHeader.indexOf('Total_Planned_Sell') + 1;
    const colTOF     = iHeader.indexOf('Total_Offer')        + 1;

    // current row values
    const cur = iTab.objects[iRowIdx];

    // apply patch (only provided fields)
    if (patch) {
      if (patch.shortTitle != null && colShort > 0)   ish.getRange(sheetRow, colShort).setValue(String(patch.shortTitle || '').trim());
      if (patch.plannedSell != null && colPlanned > 0) ish.getRange(sheetRow, colPlanned).setValue(Number(patch.plannedSell) || 0);
      if (patch.offerAmount != null && colOffer > 0)   ish.getRange(sheetRow, colOffer).setValue(Number(patch.offerAmount) || 0);
      if (patch.qty != null && colQty > 0)    ish.getRange(sheetRow, colQty).setValue( Math.max(1, Math.floor(Number(patch.qty)||1)) );
    }
    if (colUpdated > 0) ish.getRange(sheetRow, colUpdated).setValue(now);

    // resolve effective values after patch (no duplicate declarations)
    const effPlanned = (patch && patch.plannedSell != null) ? Number(patch.plannedSell) : Number(cur.PlannedSell);
    const effOffer   = (patch && patch.offerAmount != null) ? Number(patch.offerAmount) : Number(cur.OfferAmount);

    // prefer patched qty; otherwise use the value currently in the sheet (or 1)
    const qtyCell = (colQty > 0) ? Number(ish.getRange(sheetRow, colQty).getValue()) : Number(cur.Qty);
    const effQty  = (patch && patch.qty != null)
      ? Math.max(1, Math.floor(Number(patch.qty) || 1))
      : (Number.isFinite(qtyCell) && qtyCell > 0 ? Math.floor(qtyCell) : 1);


    _bt_validatePlannedVsOffer_(effPlanned, effOffer);

    // write row totals if headers exist
    if (colTPS > 0) ish.getRange(sheetRow, colTPS).setValue(+( (Number.isFinite(effPlanned)?effPlanned:0) * effQty ).toFixed(2));
    if (colTOF > 0) ish.getRange(sheetRow, colTOF).setValue(+( (Number.isFinite(effOffer)?effOffer:0)   * effQty ).toFixed(2));

    // Recalc ticket Total from remaining (non-removed) items using totals first
    const iTab2 = _readAllAsObjects_(ish);
    const itemsForTicket = iTab2.objects.filter(o => _norm_(o.TicketID) === tId && !_norm_(o.RemovedAt));
    const offerTotal = itemsForTicket.reduce((s, o) => {
      const tOffer = Number(o.Total_Offer);
      if (Number.isFinite(tOffer)) return s + tOffer;
      const q = Number(o.Qty); const of2 = Number(o.OfferAmount);
      if (Number.isFinite(of2) && Number.isFinite(q)) return s + (of2 * (q > 0 ? q : 1));
      return s + (Number.isFinite(of2) ? of2 : 0);
    }, 0);

    const tHeader = tTab.header;
    const tSheetRow = 2 + tRowIdx;
    const colTUpd  = tHeader.indexOf('UpdatedAt') + 1;
    const colTTot  = tHeader.indexOf('Total')     + 1;
    if (colTUpd > 0) tsh.getRange(tSheetRow, colTUpd).setValue(now);
    if (colTTot > 0) tsh.getRange(tSheetRow, colTTot).setValue(+offerTotal.toFixed(2));
    logs.push(`← apiBT_updateItem ok; offerTotal=${offerTotal.toFixed(2)}`);

    const out = apiBT_getTicket(tId);
    if (out && out._logs) out._logs.push(...logs);
    return out;
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? err.stack : '';
    const out = { ok:false, error: msg, stack, _logs: logs, where:'apiBT_updateItem' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

// --------------------------------------------------------------
// Batch update items for a ticket, recalc per-row totals and ticket total
// patches: [{ itemId, shortTitle?, plannedSell?, offerAmount?, qty? }, ...]
// Returns: { ok, items, _logs }
function apiBT_updateItemsBulk(ticketId, patches){
  const logs = [];
  try {
    const tId = _norm_(ticketId);
    if (!tId) throw new Error('Missing TicketID');
    if (!Array.isArray(patches) || !patches.length) {
      return _serializeForClient_({ ok:true, items: [], _logs: logs });
    }

    logs.push(`→ apiBT_updateItemsBulk start t=${tId} count=${patches.length}`);

    const tsh = bt_ticketsSheet_();
    const ish = bt_itemsSheet_();
    const now = _nowIso_();

    // Verify ticket exists
    const tTab = _readAllAsObjects_(tsh);
    const tRowIdx = tTab.objects.findIndex(o => _norm_(o.TicketID) === tId);
    if (tRowIdx < 0) throw new Error('Ticket not found: ' + tId);
    const tHeader   = tTab.header;
    const tSheetRow = 2 + tRowIdx;

    // Read all items once
    const iTab    = _readAllAsObjects_(ish);
    const iHeader = iTab.header;
    const objects = iTab.objects;

    // Column indices (0 → not found)
    const idx = {
      Short   : iHeader.indexOf('ShortTitle') + 1,
      Planned : iHeader.indexOf('PlannedSell') + 1,
      Offer   : iHeader.indexOf('OfferAmount') + 1,
      Qty     : iHeader.indexOf('Qty') + 1,
      TPS     : iHeader.indexOf('Total_Planned_Sell') + 1,
      TOF     : iHeader.indexOf('Total_Offer') + 1,
      Updated : iHeader.indexOf('UpdatedAt') + 1,
      Removed : iHeader.indexOf('RemovedAt') + 1
    };

    // Map ItemID → sheet row (for this ticket only)
    const rowMap = new Map();
    objects.forEach((o, i) => {
      if (_norm_(o.TicketID) === tId) rowMap.set(_norm_(o.ItemID), 2 + i);
    });

    // Apply patches
    for (const p of patches) {
      const iId = _norm_(p && p.itemId);
      const r = rowMap.get(iId);
      if (!r) { logs.push(`[bulk] skip unknown itemId ${iId}`); continue; }

      // Validate planned vs offer if both provided
      const curPlanned = (idx.Planned>0) ? Number(ish.getRange(r, idx.Planned).getValue()) : 0;
      const curOffer   = (idx.Offer  >0) ? Number(ish.getRange(r, idx.Offer  ).getValue()) : 0;
      const planEff = (p.plannedSell != null) ? Number(p.plannedSell) : curPlanned;
      const offerEff = (p.offerAmount != null) ? Number(p.offerAmount) : curOffer;
      _bt_validatePlannedVsOffer_(planEff, offerEff);

      if (p.shortTitle  != null && idx.Short   > 0) ish.getRange(r, idx.Short  ).setValue(String(p.shortTitle || '').trim());
      if (p.plannedSell != null && idx.Planned > 0) ish.getRange(r, idx.Planned).setValue(Number(p.plannedSell) || 0);
      if (p.offerAmount != null && idx.Offer   > 0) ish.getRange(r, idx.Offer  ).setValue(Number(p.offerAmount) || 0);
      if (p.qty         != null && idx.Qty     > 0) ish.getRange(r, idx.Qty    ).setValue(Math.max(1, Math.floor(Number(p.qty)||1)));
      if (idx.Updated > 0) ish.getRange(r, idx.Updated).setValue(now);

      // Compute + write per-row totals (prefer values we just wrote)
      const ps  = (p.plannedSell != null) ? Number(p.plannedSell) : (idx.Planned>0 ? Number(ish.getRange(r, idx.Planned).getValue()) : 0);
      const of  = (p.offerAmount != null) ? Number(p.offerAmount) : (idx.Offer  >0 ? Number(ish.getRange(r, idx.Offer  ).getValue()) : 0);
      const qty = (p.qty         != null) ? Math.max(1, Math.floor(Number(p.qty)||1))
                                          : (idx.Qty    >0 ? Math.max(1, Math.floor(Number(ish.getRange(r, idx.Qty).getValue())||1)) : 1);
      if (idx.TPS > 0) ish.getRange(r, idx.TPS).setValue(+((Number.isFinite(ps)?ps:0) * qty).toFixed(2));
      if (idx.TOF > 0) ish.getRange(r, idx.TOF).setValue(+((Number.isFinite(of)?of:0) * qty).toFixed(2));
    }

    // Re-read items for this ticket and compute ticket Total (offers)
    const iTab2 = _readAllAsObjects_(ish);
    const itemsForTicket = iTab2.objects.filter(o => _norm_(o.TicketID) === tId && !_norm_(o.RemovedAt));
    const offerTotal = itemsForTicket.reduce((s, o) => {
      const tOffer = Number(o.Total_Offer);
      if (Number.isFinite(tOffer)) return s + tOffer;
      const q = Number(o.Qty); const of2 = Number(o.OfferAmount);
      if (Number.isFinite(of2) && Number.isFinite(q)) return s + (of2 * (q > 0 ? q : 1));
      return s + (Number.isFinite(of2) ? of2 : 0);
    }, 0);

    // Write ticket UpdatedAt + Total
    const colTUpd = tHeader.indexOf('UpdatedAt') + 1;
    const colTTot = tHeader.indexOf('Total')     + 1;
    if (colTUpd > 0) tsh.getRange(tSheetRow, colTUpd).setValue(now);
    if (colTTot > 0) tsh.getRange(tSheetRow, colTTot).setValue(+offerTotal.toFixed(2));

    logs.push(`[bulk] updated ${patches.length} rows; ticket Total=${offerTotal.toFixed(2)}`);

    return _serializeForClient_({ ok:true, items: itemsForTicket, _logs: logs });
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? String(err.stack) : '';
    const out = { ok:false, error: msg, stack, _logs: logs, where:'apiBT_updateItemsBulk' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

function apiBT_renderOfferHtml(ticketId, mode){
  const logs = [];
  try {
    const id = _norm_(ticketId);
    if (!id) throw new Error('Missing TicketID');

    // fetch fresh data
    const out = apiBT_getTicket(id);
    if (!out || out.ok !== true) throw new Error(out && out.error || 'Failed to load ticket');

    const t = out.ticket || {};
    const items = (out.items || []).filter(x => !x.RemovedAt);
    const offerTotal = items.reduce((s,o)=> s + (Number(o.OfferAmount)||0), 0);

    const fmt = n => '$' + (Number(n)||0).toFixed(2);

    // logo: data-url for print; cid for email
    const isEmail = String(mode||'').toLowerCase() === 'email';
    const logoSrc = isEmail ? 'cid:logo' : _logoDataUrl_();

    const html =
`<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Offer for ${t.TicketID||''}</title>
  <style>
    body { font: 14px/1.4 -apple-system, Segoe UI, Roboto, sans-serif; color:#111; margin:24px; }
    .hdr { display:flex; align-items:center; gap:16px; }
    .hdr img { height: 64px; }
    h1 { margin:8px 0 0; font-size:20px; }
    .meta { color:#555; margin-top:2px; }
    table { width:100%; border-collapse:collapse; margin-top:18px; }
    th, td { text-align:left; padding:10px; border-bottom:1px solid #eee; }
    th { font-weight:600; color:#444; background:#fafafa; }
    tfoot td { font-weight:700; }
    .right { text-align:right; }
    .muted { color:#777; }
    .note { margin-top:18px; color:#444; }
    @media print { .no-print { display:none; } }
  </style>
</head>
<body>
  <div class="hdr">
    ${logoSrc ? `<img src="${logoSrc}" alt="Mad Rad Retro Toys">` : ''}
    <div>
      <h1>Offer Summary</h1>
      <div class="meta">
        Ticket: <b>${t.TicketID||''}</b> • Estimate: ${t.EstimateID||''}<br>
        Customer: ${t.SellerName||''} • ${t.SellerPhone||''} • ${t.SellerEmail||''}<br>
        Generated: ${_nowIso_()}
      </div>
    </div>
  </div>

  <table>
    <thead><tr><th>Item</th><th class="right">Offer</th></tr></thead>
    <tbody>
      ${items.map(it => `
        <tr>
          <td>${(it.ShortTitle||'').toString().replace(/[<>]/g,'')}</td>
          <td class="right">${fmt(it.OfferAmount)}</td>
        </tr>
      `).join('')}
    </tbody>
    <tfoot>
      <tr><td class="right">Total Offer</td><td class="right">${fmt(offerTotal)}</td></tr>
    </tfoot>
  </table>

  <div class="note">Thank you for considering Mad Rad Retro Toys. Prices are valid for a limited time and subject to inspection at payout.</div>

  ${isEmail ? '' : `<div class="no-print" style="margin-top:16px">
    <button onclick="window.print()">Print</button>
  </div>`}
</body>
</html>`;

    return _serializeForClient_({ ok:true, html, _logs:logs, offerTotal, ticket:t, items });
  } catch (err){
    const out = { ok:false, error: String(err && err.message ? err.message : err), _logs: logs, where: 'apiBT_renderOfferHtml' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

// --------------------------------------------------------------
// Remove a single Buy Ticket item (soft delete via RemovedAt)
// Recalc totals and return fresh data
// --------------------------------------------------------------
function apiBT_removeItem(ticketId, itemId) {
  const logs = [];
  try {
    const tId = _norm_(ticketId);
    const iId = _norm_(itemId);
    if (!tId) throw new Error('Missing TicketID');
    if (!iId) throw new Error('Missing ItemID');

    logs.push(`→ apiBT_removeItem start t=${tId} i=${iId}`);

    const tsh = bt_ticketsSheet_();
    const ish = bt_itemsSheet_();

    // Verify ticket exists
    const tTab = _readAllAsObjects_(tsh);
    const tRowIdx = tTab.objects.findIndex(o => _norm_(o.TicketID) === tId);
    if (tRowIdx < 0) throw new Error('Ticket not found: ' + tId);

    // Find item row
    const iTab = _readAllAsObjects_(ish);
    const iRowIdx = iTab.objects.findIndex(o => _norm_(o.ItemID) === iId && _norm_(o.TicketID) === tId);
    if (iRowIdx < 0) throw new Error('Item not found for this ticket');

    const iHeader = iTab.header;
    const sheetRow = 2 + iRowIdx; // data starts on row 2
    const now = _nowIso_();

    const colRemoved = iHeader.indexOf('RemovedAt') + 1;
    const colUpdated = iHeader.indexOf('UpdatedAt') + 1;
    if (colRemoved > 0) ish.getRange(sheetRow, colRemoved).setValue(now);
    if (colUpdated > 0) ish.getRange(sheetRow, colUpdated).setValue(now);
    logs.push('[R] item marked removed');

    // Recalc totals from remaining (non-removed) items — use totals first
    const iTab2 = _readAllAsObjects_(ish);
    const itemsForTicket = iTab2.objects.filter(o => _norm_(o.TicketID) === tId && !_norm_(o.RemovedAt));
    const offerTotal = itemsForTicket.reduce((s, o) => {
      const tOffer = Number(o.Total_Offer);
      if (Number.isFinite(tOffer)) return s + tOffer;
      const q = Number(o.Qty); const of2 = Number(o.OfferAmount);
      if (Number.isFinite(of2) && Number.isFinite(q)) return s + (of2 * (q > 0 ? q : 1));
      return s + (Number.isFinite(of2) ? of2 : 0);
    }, 0);

    const tHeader = tTab.header;
    const tSheetRow = 2 + tRowIdx;
    const colTUpd  = tHeader.indexOf('UpdatedAt') + 1;
    const colTTot  = tHeader.indexOf('Total')     + 1;
    if (colTUpd > 0) tsh.getRange(tSheetRow, colTUpd).setValue(now);
    if (colTTot > 0) tsh.getRange(tSheetRow, colTTot).setValue(+offerTotal.toFixed(2));
    
    logs.push(`[R] totals updated to ${offerTotal.toFixed(2)}`);

    // Return only the updated row so the client doesn't pay for a full refresh here.
    // We already fetched iTab2 above, so use it to locate the updated row.
    const updated = iTab2.objects.find(o => _norm_(o.ItemID) === iId);
    const out = { ok: true, items: updated ? [updated] : [], _logs: logs };
    return _serializeForClient_(out);

  } catch (err){
    const msg = (err && err.message) ? err.message : String(err);
    const stack = (err && err.stack) ? err.stack : '';
    const out = { ok:false, error: msg, stack, _logs: logs, where:'apiBT_removeItem' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

function apiBT_sendOffer(ticketId, via, dest, carrier){
  const logs = [];
  try {
    const id = _norm_(ticketId);
    const mode = 'email'; // render email-friendly (cid logo)
    if (!id) throw new Error('Missing TicketID');
    const r = apiBT_renderOfferHtml(id, mode);
    if (!r || r.ok !== true) throw new Error(r && r.error || 'Failed to render HTML');

    // Resolve destination
    let to = String(dest||'').trim();
    const v = String(via||'').toLowerCase();

    if (v === 'sms') {
      const addr = _carrierEmail_(to, carrier);
      if (!addr) throw new Error('Unknown carrier or invalid phone for SMS gateway');
      to = addr;
    }
    if (!to) throw new Error('Missing destination address/phone');

    // Inline logo (cid:logo)
    const logo = _logoBlob_();
    const subject = `Offer for ${r.ticket && r.ticket.TicketID ? r.ticket.TicketID : 'your items'}`;
    const html = r.html;

    const opts = {
      name: 'Mad Rad Retro Toys',
      htmlBody: html
    };
    if (logo) opts.inlineImages = { 'logo': logo };

    MailApp.sendEmail(to, subject, 'HTML offer attached', opts);

    return _serializeForClient_({ ok:true, to, via:v, _logs:logs });
  } catch (err){
    const out = { ok:false, error: String(err && err.message ? err.message : err), _logs: logs, where:'apiBT_sendOffer' };
    Logger.log(JSON.stringify(out));
    return _serializeForClient_(out);
  }
}

function _prop_(k){ try { return PropertiesService.getScriptProperties().getProperty(k) || ''; } catch(e){ return ''; } }
function _logoBlob_(){
  const id = _prop_('MADRAD_LOGO_FILE_ID');
  if (!id) return null;
  try { return DriveApp.getFileById(id).getBlob().setName('logo.png'); } catch(e){ return null; }
}
function _logoDataUrl_(){ // for print (embed)
  const b = _logoBlob_(); if (!b) return '';
  const bytes = Utilities.base64Encode(b.getBytes());
  const ct = b.getContentType() || 'image/png';
  return `data:${ct};base64,${bytes}`;
}
function _carrierEmail_(phone, carrier){
  const p = String(phone||'').replace(/[^\d]/g,''); if (!p) return '';
  const c = String(carrier||'').toLowerCase();
  // Common US gateways (can add more later)
  const map = {
    'att':'txt.att.net',
    'at&t':'txt.att.net',
    'verizon':'vtext.com',
    'tmobile':'tmomail.net',
    't-mobile':'tmomail.net',
    'sprint':'messaging.sprintpcs.com',
    'cricket':'sms.mycricket.com',
    'uscellular':'email.uscc.net',
    'us cellular':'email.uscc.net',
    'googlefi':'msg.fi.google.com',
    'google fi':'msg.fi.google.com',
  };
  const host = map[c] || '';
  return host ? `${p}@${host}` : '';
}

/**
 * Recompute "Total Pay Amount" = Cash Drawer + Trade + Other for a ticket,
 * write it to the sheet, and return a fresh ticket.
 */
function apiBT_recalcPayTotal(ticketId){
  const id = _norm_(ticketId);
  if (!id) return _serializeForClient_({ ok:false, error:'Missing TicketID' });

  const tsh   = bt_ticketsSheet_();
  const lastC = tsh.getLastColumn();
  const hdr   = tsh.getRange(1,1,1,lastC).getValues()[0];

  const norm  = s => String(s||'').trim().toLowerCase();
  const toNum = x => {
    const n = Number(x); if (Number.isFinite(n)) return n;
    const s = String(x==null?'':x).replace(/[^\d.-]/g,'');
    const p = parseFloat(s); return Number.isFinite(p) ? p : 0;
  };

  // Locate row for this TicketID
  const tTab   = _readAllAsObjects_(tsh);
  const rowIdx = (tTab.objects || []).findIndex(o => _norm_(o.TicketID) === id);
  if (rowIdx < 0) return _serializeForClient_({ ok:false, error:'Ticket not found: ' + id });
  const row = 2 + rowIdx;

  // Ensure columns exist (same behavior as apiBT_bumpPayComponent)
  const findOrAppend = (names) => {
    for (const n of names){
      const i = hdr.findIndex(h => norm(h) === norm(n));
      if (i >= 0) return i + 1;
    }
    const newCol = tsh.getLastColumn() + 1;
    tsh.getRange(1, newCol).setValue(names[0]);
    return newCol;
  };

  const colCash  = findOrAppend(['Cash Drawer Payout Amount','Payout Amount']);
  const colTrade = findOrAppend(['Trade Amount']);
  const colOther = findOrAppend(['Other Pay (zelle, cashapp, etc.)']);

  // Total column
  let colTotal = hdr.findIndex(h => norm(h) === norm('Total Pay Amount')) + 1;
  if (colTotal <= 0) {
    colTotal = tsh.getLastColumn() + 1;
    tsh.getRange(1, colTotal).setValue('Total Pay Amount');
  }

  // Recompute = cash + trade + other
  const cash  = toNum(tsh.getRange(row, colCash).getValue());
  const trade = toNum(tsh.getRange(row, colTrade).getValue());
  const other = toNum(tsh.getRange(row, colOther).getValue());
  const total = +(cash + trade + other).toFixed(2);
  tsh.getRange(row, colTotal).setValue(total);

  // Touch UpdatedAt if present
  const iUpd = hdr.findIndex(h => norm(h) === 'updatedat');
  if (iUpd >= 0) tsh.getRange(row, iUpd + 1).setValue(_nowIso_());

  // Return fresh ticket
  const tTab2  = _readAllAsObjects_(tsh);
  const ticket = (tTab2.objects || []).find(o => _norm_(o.TicketID) === id);
  return _serializeForClient_({ ok:true, ticket });
}
