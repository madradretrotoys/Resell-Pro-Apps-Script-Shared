/***** Daily Sales Summary Email (isolated) *****/
const DE_USERS_SHEET = (typeof USERS_SHEET !== 'undefined') ? USERS_SHEET : 'User Permissions';
const DE_TZ = (typeof CASH_TZ !== 'undefined') ? CASH_TZ : (Session.getScriptTimeZone() || 'America/Denver');

// Public entry (use this in the trigger)
function sendDailySalesSummary(dateOpt) {
  return de_sendDailySalesSummary_(dateOpt);
}

// ---- Implementation (private helpers below) ----
function de_sendDailySalesSummary_(dateOpt) {
  const now = new Date();
  let target =
    (dateOpt instanceof Date) ? dateOpt :
    (typeof dateOpt === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateOpt)) ? new Date(dateOpt + 'T12:00:00') :
    now;

  const dayStr   = Utilities.formatDate(target, DE_TZ, 'yyyy-MM-dd');
  const dayStart = de_startOfDay_(target, DE_TZ);
  const dayEnd   = de_endOfDay_(target, DE_TZ);

  const recipients = de_getRecipients_();
  if (!recipients.length) return { ok:false, msg:'No recipients with Daily Sales Summary Email = Y' };

  const sales   = de_readSales_(dayStart, dayEnd);          // from "Sales Log"
  const cashTot = de_sumCashSales_(sales);

  const dr      = de_readDrawerOpenClose_(dayStr);          // from "Cash Drawer Log"
  const payouts = de_readPayouts_(dayStart, dayEnd);        // from "Cash Payouts"
  const cashRefundsTotal = de_readCashRefunds_(dayStart, dayEnd);

  const openBy  = dr.opening || {};
  const closeBy = dr.closing || {};
  const payBy   = payouts.byDrawer || {};
  const safeTot = payouts.safeTotal || 0;

  const drawers = [...new Set([].concat(Object.keys(openBy), Object.keys(closeBy))).filter(Boolean)].sort();
  const impliedBy = {};
  drawers.forEach(d => {
    const open = +openBy[d]  || 0;
    const close= +closeBy[d] || 0;
    const pay  = +payBy[d]   || 0;
    impliedBy[d] = +(close - open + pay).toFixed(2);
  });

  const sumOpen   = de_sumVals_(openBy);
  const sumClose  = de_sumVals_(closeBy);
  const sumPayout = de_sumVals_(payBy);
  const totalVariance = +(sumClose - sumOpen + sumPayout + cashRefundsTotal - cashTot).toFixed(2);

  const html = de_renderEmail_({
    tz: DE_TZ, dayStr,
    sales, cashSalesTotal: cashTot,
    openingByDrawer: openBy, closingByDrawer: closeBy,
    payoutByDrawer: payBy, impliedByDrawer: impliedBy,
    safeTotal: safeTot, totalVariance,
    cashRefundsTotal                                   // <-- pass to template
  });

  const subject = `Daily Cash Summary — ${dayStr}`;
  recipients.forEach(to => MailApp.sendEmail({ to, subject, htmlBody: html }));
  return { ok:true, sent:recipients.length, recipients, cashTot, totalVariance };
}

// Sum CASH refunds for the day by reading refund rows from "Sales Log".
// Looks for envelope.type === 'refund' and sums the CASH portion of envelope.payments (negative numbers).
function de_readCashRefunds_(start, end){
  const sh = sheet_('Sales Log');
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2,1,last-1, Math.max(12, sh.getLastColumn())).getValues();

  let cashRefunds = 0;
  for (let i = 0; i < vals.length; i++){
    const r = vals[i];
    const ts = r[0];
    if (!(ts instanceof Date) || ts < start || ts > end) continue;

    // Column K (index 10) stores the JSON envelope
    let env = null;
    try { env = r[10] ? JSON.parse(r[10]) : null; } catch(_) {}

    if (!env || String(env.type || '').toLowerCase() !== 'refund') continue;

    // Prefer explicit payments array in the envelope
    if (Array.isArray(env.payments) && env.payments.length){
      env.payments.forEach(p => {
        if (String(p.method || '').toLowerCase() === 'cash'){
          cashRefunds += Math.abs(Number(p.amount) || 0); // amounts are stored negative in refunds; take absolute
        }
      });
    } else {
      // Fallback: if Payment Method (col I) mentions Cash, use absolute of Total (col H)
      const paymentMethod = String(r[8] || '').toLowerCase();
      if (paymentMethod.includes('cash')){
        cashRefunds += Math.abs(Number(r[7]) || 0);
      }
    }
  }
  return +cashRefunds.toFixed(2);
}


/* ----------------- data collectors ----------------- */

function de_getRecipients_(){
  const sh = sheet_(DE_USERS_SHEET);
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return [];
  const hdr = vals.shift();
  const loginIdx = hdr.findIndex(h => String(h).trim().toLowerCase() === 'login');
  const dailyIdx = hdr.findIndex(h => String(h).trim().toLowerCase() === 'daily sales summary email');
  if (loginIdx < 0 || dailyIdx < 0) return [];
  const out = [];
  vals.forEach(r => {
    const login = String(r[loginIdx]||'').trim();
    const daily = String(r[dailyIdx]||'').trim().toUpperCase();
    if (login && daily === 'Y') out.push(login);
  });
  return out;
}

function de_readSales_(start, end){
  const sh = sheet_('Sales Log');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last-1, Math.max(12, sh.getLastColumn())).getValues();
  // Columns assumed: A Timestamp, B SaleID, I Payment, H Total, L Clerk (adjust if your schema differs)
  return rows
    .filter(r => r[0] && r[0] >= start && r[0] <= end)
    .map(r => ({
      ts: r[0],
      saleId: r[1],
      total: +r[7] || 0,
      payment: String(r[8]||''),
      clerk: String(r[11]||'')
    }));
}
function de_sumCashSales_(sales){
  return +sales.filter(s => String(s.payment||'').toLowerCase() === 'cash')
               .reduce((a,s)=>a+(+s.total||0),0).toFixed(2);
}

function de_readDrawerOpenClose_(dayStr){
  const sh = sheet_('Cash Drawer Log');
  const last = sh.getLastRow();
  const opening = {}, closing = {};
  if (last >= 2) {
    const vals = sh.getRange(2,1,last-1, Math.max(22, sh.getLastColumn())).getValues();
    vals.forEach(r => {
      const ts = r[0];                                   // Timestamp
      const period = String(r[3]||r[4-1]||'').toLowerCase(); // Period (D) or fallback
      const drawer = String(r[4]||r[5-1]||'1').trim();   // Drawer (E)
      const grand  = +r[19] || 0;                        // Grand Total (T = 20th col, index 19)
      const dKey = (ts instanceof Date)
        ? Utilities.formatDate(ts, DE_TZ, 'yyyy-MM-dd')
        : String(r[2]||'').slice(0,10);
      if (dKey !== dayStr) return;
      if (period === 'open' || period === 'opening') {
        if (!(drawer in opening)) opening[drawer] = grand;       // keep first open
      } else if (period === 'close' || period === 'closing') {
        closing[drawer] = grand;                                  // keep latest close
      }
    });
  }
  return { opening, closing };
}

function de_readPayouts_(start, end){
  const sh = sheet_('Cash Payouts');
  const last = sh.getLastRow();
  const out = { byDrawer:{}, safeTotal:0 };
  if (last >= 2) {
    const vals = sh.getRange(2,1,last-1, Math.max(10, sh.getLastColumn())).getValues();
    vals.forEach(r => {
      const ts = r[0];
      const deleted = String(r[9]||'').toUpperCase() === 'Y'; // J
      if (deleted) return;
      if (!(ts instanceof Date) || ts < start || ts > end) return;
      const d1 = +r[3]||0, d2 = +r[4]||0, safe = +r[5]||0;     // D,E,F
      out.byDrawer['1'] = (out.byDrawer['1']||0) + d1;
      out.byDrawer['2'] = (out.byDrawer['2']||0) + d2;
      out.safeTotal += safe;
    });
  }
  out.byDrawer['1'] = +((out.byDrawer['1']||0).toFixed(2));
  out.byDrawer['2'] = +((out.byDrawer['2']||0).toFixed(2));
  out.safeTotal     = +(out.safeTotal.toFixed(2));
  return out;
}

/* ----------------- template render ----------------- */
function de_renderEmail_(ctx){
  const t = HtmlService.createTemplateFromFile('email_daily_summary'); // HTML file
  Object.keys(ctx).forEach(k => t[k] = ctx[k]);
  return t.evaluate().getContent();
}

/* ----------------- utils ----------------- */
function de_startOfDay_(d, tz){
  const y = Utilities.formatDate(d, tz, 'yyyy');
  const m = Utilities.formatDate(d, tz, 'MM');
  const dd= Utilities.formatDate(d, tz, 'dd');
  return new Date(`${y}-${m}-${dd}T00:00:00`);
}
function de_endOfDay_(d, tz){
  const y = Utilities.formatDate(d, tz, 'yyyy');
  const m = Utilities.formatDate(d, tz, 'MM');
  const dd= Utilities.formatDate(d, tz, 'dd');
  return new Date(`${y}-${m}-${dd}T23:59:59`);
}
function de_sumVals_(obj){ return Object.values(obj||{}).reduce((a,b)=>a+(+b||0),0); }
