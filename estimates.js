/** Estimate Requests — endpoints and sheet helpers (Phase D1–D3, hardened) */

const EST_SHEET = 'Estimate Requests';

function est_sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(EST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EST_SHEET);
    sh.appendRow([
      'ReqID','Status','CreatedAt','UpdatedAt',
      'First','Last','Phone','Email','PreferredContact',
      'WouldLike','ItemsDesc','WaiverAgreed','SigName',
      'Notes','DeclinedAt','DuePickupBy','AcceptedAt','PaidAt','ClosedAt'
    ]);
  }
  return sh;
}

function est_makeId_() {
  const d = new Date();
  const pad = n => String(n).padStart(2,'0');
  const seq = Math.floor(Math.random()*9000+1000);
  return `EST-${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${seq}`;
}

function est_nowIso_() {
  return Utilities.formatDate(new Date(), CASH_TZ || Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function est_findRowById_(reqId) {
  const sh = est_sheet_();
  const ids = sh.getRange(2,1,Math.max(0, sh.getLastRow()-1),1).getValues().flat();
  const idx = ids.indexOf(String(reqId));
  return (idx === -1) ? null : (2 + idx);
}

/** D1: Customer submit endpoint (returns a rich object) */
function apiEst_CreateRequest(form) {
  try {
    const sh = est_sheet_();

    const first = String((form && (form.first || form.First))||'').trim();
    const last  = String((form && (form.last  || form.Last ))||'').trim();
    const phone = String((form && (form.phone || form.Phone))||'').trim();
    const email = String((form && (form.email || form.Email))||'').trim();
    const contact = String((form && (form.contact|| form.PreferredContact))||'').trim() || 'phone';

    // support either wouldLike / wish keys from older client
    const wouldLikeRaw = (form && (form.wouldLike ?? form.wish));
    const wouldLike = (wouldLikeRaw===''||wouldLikeRaw==null)
      ? ''
      : Number(String(wouldLikeRaw).replace(/[^\d.]/g,''));

    const items = String((form && (form.items || form.ItemsDesc))||'').trim();
    const agree = !!(form && (form.agree || form.WaiverAgreed));
    const sigName = String((form && (form.sigName || form.sig || form.SigName))||'').trim();

    if (!first || !last)      throw new Error('Please enter first and last name.');
    if (!phone && !email)     throw new Error('Please enter a phone or an email.');
    if (!items)               throw new Error('Please describe the items you’re dropping off.');
    if (!agree || !sigName)   throw new Error('You must agree to the waiver and sign.');

    const reqId = est_makeId_();
    const now = est_nowIso_();

    sh.appendRow([
      reqId,'Open',now,now,
      first,last,phone,email,contact,
      (wouldLike===''? '' : Number.isFinite(wouldLike)? wouldLike : ''),
      items, 'Y', sigName,
      '', '', '', '', '', ''
    ]);

    // Return a rich, serializable object (what the receipt UI expects)
    const res = {
      ok: true,
      reqId: reqId,
      id: reqId,
      first, last, phone, email, contact,
      wouldLike: (wouldLike===''? '' : wouldLike),
      wish: (wouldLike===''? '' : wouldLike),   // legacy alias
      items, sigName,
      sig: sigName,
      status: 'Open',
      createdAt: now,
      updatedAt: now
    };
    Logger.log('apiEst_CreateRequest -> %s', JSON.stringify(res));
    return JSON.parse(JSON.stringify(res));
  } catch (err) {
    const out = { ok:false, error: String(err && err.message ? err.message : err) };
    Logger.log('apiEst_CreateRequest ERROR: %s', out.error);
    return out;
  }
}

/** D2: List/search for the internal queue (hardened, always returns JSON) */
function apiEst_List(opts) {
  try {
    const qRaw = String((opts && opts.q) || '');
    const q = qRaw.toLowerCase().trim();

    // Accept pipe list, default to open-ish statuses
    const st = String((opts && opts.status) || 'Open|In Progress|Offered|Declined-WaitingPickup');
    const statuses = st.split('|').map(s => s.trim().toLowerCase()).filter(Boolean);

    const sh = est_sheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok:true, rows:[] };

    const values = sh.getRange(2,1,last-1, sh.getLastColumn()).getValues();
    const rows = values.map(r => ({
      reqId:        r[0],
      status:       r[1],
      createdAt:    r[2],
      updatedAt:    r[3],
      first:        r[4],
      last:         r[5],
      phone:        r[6],
      email:        r[7],
      contact:      r[8],
      wouldLike:    r[9],
      items:        r[10],
      waiverAgreed: r[11],
      sigName:      r[12],
      notes:        r[13]
    }));

    const filt = rows.filter(x => {
      const stx = String(x.status||'').toLowerCase().trim();           // <= trim fixes "Open " cases
      if (statuses.length && !statuses.includes(stx)) return false;
      if (!q) return true;
      const hay = [
        x.reqId, x.first, x.last, x.phone, x.email,
        (x.wouldLike==null?'':String(x.wouldLike)), x.items
      ].join(' ').toLowerCase();
      return hay.includes(q);
    });

    // newest first
    filt.sort((a,b)=> String(b.createdAt||'').localeCompare(String(a.createdAt||'')));

    const out = { ok:true, rows: filt };
    Logger.log('apiEst_List -> count=%s (filter=%s | %s)', filt.length, qRaw, st);
    return JSON.parse(JSON.stringify(out));
  } catch (err) {
    const out = { ok:false, rows:[], error: String(err && err.message ? err.message : err) };
    Logger.log('apiEst_List ERROR: %s', out.error);
    return out;
  }
}

/** D3: Update status (returns {ok:true}) */
function apiEst_UpdateStatus(reqId, nextStatus) {
  try {
    const row = est_findRowById_(reqId);
    if (!row) throw new Error('Request not found');

    const sh = est_sheet_();
    const statusCol = 2, updatedCol = 4, declinedAtCol = 16, duePickupCol = 17;
    const now = est_nowIso_();
    const next = String(nextStatus||'').trim();

    sh.getRange(row, statusCol).setValue(next);
    sh.getRange(row, updatedCol).setValue(now);

    if (next === 'Declined-WaitingPickup') {
      const declinedDate = new Date();
      const due = new Date(declinedDate.getTime() + 30*24*60*60*1000);
      const fmt = d => Utilities.formatDate(d, CASH_TZ || Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sh.getRange(row, declinedAtCol).setValue(fmt(declinedDate));
      sh.getRange(row, duePickupCol).setValue(fmt(due));
    }
    Logger.log('apiEst_UpdateStatus -> %s -> %s', reqId, next);
    return { ok:true };
  } catch (err) {
    const out = { ok:false, error: String(err && err.message ? err.message : err) };
    Logger.log('apiEst_UpdateStatus ERROR: %s', out.error);
    return out;
  }
}

/** Convenience: fetch a single request by id */
function apiEst_GetById(reqId) {
  try {
    const row = est_findRowById_(reqId);
    if (!row) return { ok:false, error:'not_found' };
    const sh = est_sheet_();
    const v = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
    const out = {
      ok:true,
      req: {
        reqId: v[0], status:v[1], createdAt:v[2], updatedAt:v[3],
        first:v[4], last:v[5], phone:v[6], email:v[7], contact:v[8],
        wouldLike:v[9], items:v[10], waiverAgreed:v[11], sigName:v[12], notes:v[13]
      }
    };
    Logger.log('apiEst_GetById -> %s', reqId);
    return JSON.parse(JSON.stringify(out));
  } catch (err) {
    const out = { ok:false, error: String(err && err.message ? err.message : err) };
    Logger.log('apiEst_GetById ERROR: %s', out.error);
    return out;
  }
}

/** Debug snapshot used by the UI when zero rows come back */
function apiEst_DebugSnapshot(opts){
  try {
    const base = apiEst_List(opts) || { ok:false, rows:[] };
    const sh = est_sheet_();
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    const headers = (lastCol>0) ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
    const sampleCount = Math.max(0, Math.min(5, lastRow-1));
    const sample = (sampleCount>0) ? sh.getRange(2,1,sampleCount,lastCol).getValues() : [];
    base.debug = {
      sheetName: sh.getName(),
      lastRow: lastRow,
      lastCol: lastCol,
      headers: headers,
      sampleFirst5Rows: sample,
      filterParams: {
        q: String((opts && opts.q) || ''),
        status: String((opts && opts.status) || '')
      }
    };
    Logger.log('apiEst_DebugSnapshot -> lastRow=%s lastCol=%s', lastRow, lastCol);
    return JSON.parse(JSON.stringify(base));
  } catch (err) {
    const out = { ok:false, rows:[], debug:null, error: String(err && err.message ? err.message : err) };
    Logger.log('apiEst_DebugSnapshot ERROR: %s', out.error);
    return out;
  }
}
