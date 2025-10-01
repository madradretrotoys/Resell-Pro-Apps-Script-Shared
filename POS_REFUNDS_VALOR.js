/***** ============================================================
 * Begin POS CODE
 * ============================================================ *****/

// Managers/Admins enter passcode to approve an over-limit discount
function normalizeMaxDiscount_(val) {
  if (val == null) return 0;

  // If the sheet kept it as a number (e.g., 0.05 for 5%)
  if (typeof val === 'number') {
    return val <= 1 ? val * 100 : val; // 0.05 -> 5
  }

  // If it's text
  const s = String(val).trim();
  if (!s) return 0;
  if (s.toLowerCase() === 'full') return 1000; // effectively unlimited

  // Grab the number portion (handles "5", "5%", " 5 % ", "0.05")
  const m = s.match(/(\d+(?:\.\d+)?)/);
  if (!m) return 0;
  let n = Number(m[1]);

  // If there was an explicit % sign, use as-is; otherwise, if <=1, it’s likely a percent value from the sheet (0.05)
  if (!s.includes('%') && n <= 1) n = n * 100;
  return n;
}



/* ==========================================================
   ================   POS HELPERS (Phase 1)   ================
   ========================================================== */

/** Write client traces into the execution log (called from the browser). */
function apiClientLog(entry) {
  try {
    var scope = (entry && entry.scope) || 'Client';
    var trace = Array.isArray(entry && entry.trace) ? entry.trace : [JSON.stringify(entry)];
    Logger.log('[%s] %s', scope, trace.join(' | '));
  } catch (e) {
    // swallow — logging must never throw
  }
  return true;
}

function ensureSalesLogHeaders_() {
  const sh = sheet_('Sales Log');
  const desired = [
    'Timestamp','Sale ID','Raw Subtotal','Line Discounts','Subtotal Discount',
    'Subtotal','Tax','Total','Payment Method','Fees','Items JSON','Clerk'
  ];
  const lastCol = Math.max(1, sh.getLastColumn());
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const matches =
    existing.length >= desired.length &&
    desired.every((h, i) => (existing[i] || '') === h);
  if (matches) return;
  sh.getRange(1, 1, 1, desired.length).setValues([desired]);
  sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm');
  sh.getRange('C:G').setNumberFormat('$#,##0.00');
  sh.getRange('H:H').setNumberFormat('$#,##0.00');
}
function getSalesTaxRate_() {
  try {
    const sh = sheet_('Sales Tax');
    const n = Math.max(0, sh.getLastRow() - 1);
    if (n <= 0) return 0.085;
    const val = sh.getRange(2, 1).getValue();
    const pct = parseFloat(val);
    if (isNaN(pct)) return 0.085;
    return pct / 100.0;
  } catch (_) {
    return 0.085;
  }
}
function getSalesTaxRate() { return getSalesTaxRate_(); }

/**
 * Batch-set Vendoo_Listing_Status for a list of SKUs.
 * - Non-blocking: intended to be called AFTER the sale row is written.
 * - Only writes the Vendoo_Listing_Status column; does NOT touch other columns/validations.
 */
function setVendooStatusForSkus_(skus, statusText) {
  if (!Array.isArray(skus) || !skus.length) return;
  const status = String(statusText || '').trim() || 'Pending Removal';

  const sh = sheet_('SKU Tracker');
  const last = sh.getLastRow();
  if (last <= 1) return;

  // Read header to find columns we need
  const header = sh.getRange(1,1,1,Math.max(20, sh.getLastColumn())).getValues()[0].map(String);
  const cSku  = header.indexOf('SKU') + 1;
  const cVend = header.indexOf('Vendoo_Listing_Status') + 1;
  if (!cSku || !cVend) return;

  // Build an index for SKU -> row
  const skuVals = sh.getRange(2, cSku, last - 1, 1).getValues().flat().map(v => String(v||'').trim());
  const idxBySku = new Map();
  for (let i = 0; i < skuVals.length; i++) {
    const s = skuVals[i]; if (s) idxBySku.set(s, 2 + i);
  }

  // Normalize unique SKUs
  const uniq = Array.from(new Set(skus.map(s => String(s||'').trim()).filter(Boolean)));

  // Write status to each found row
  uniq.forEach(sku => {
    const row = idxBySku.get(sku);
    if (row) sh.getRange(row, cVend).setValue(status);
  });
}

/** Ensure the Vendoo queue sheet exists with stable columns. */
function ensureVendooTaskSheet_() {
  // Auto-create the sheet if missing, then ensure the header row
  const ss = book_();
  let sh = ss.getSheetByName('Vendoo Tasks');
  if (!sh) sh = ss.insertSheet('Vendoo Tasks');

  const desired = [
    'Task ID','Status','Lease Until','Created At',
    'Sale ID','SKU','Vendoo URL','Price Cents','Qty',
    'Attempts','Last Error'
  ];

  const lastCol = Math.max(1, sh.getLastColumn());
  const firstRow = sh.getRange(1, 1, 1, Math.max(lastCol, desired.length)).getValues()[0].map(String);
  const ok = firstRow.length >= desired.length && desired.every((h,i)=> (firstRow[i]||'')===h);
  if (!ok) {
    sh.getRange(1, 1, 1, desired.length).setValues([desired]);
    // (optional niceties)
    sh.getRange('C:C').setNumberFormat('m/d/yyyy h:mm'); // Lease Until
    sh.getRange('D:D').setNumberFormat('m/d/yyyy h:mm'); // Created At
  }
  return sh;
}

/**
 * Helper: find 1-based column index by exact header name in row 1.
 * Returns 0 if not found.
 */
function _findCol_(sh, name){
  const want = String(name || '').trim();
  if (!want) return 0;
  const last = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, last).getValues()[0];
  for (let c = 1; c <= last; c++) {
    if (String(header[c-1] || '').trim() === want) return c;
  }
  return 0;
}

/** ===== Tiny utils used by Vendoo queue ===== */
function _now_(){ return new Date(); }
/** Stable unique id for a task */
function _id_(){ return Utilities.getUuid(); }

// MRAD-START vendoo-final-unit-price-helper
/**
 * Compute the final per-unit price **in cents** for a POS item,
 * after applying line-level discounts (and optionally a proportional
 * share of any ticket-level/subtotal discount when we enable that).
 *
 * Expected item shape (robust to missing fields):
 *   it.price      → unit price (number, dollars)
 *   it.qty        → quantity (number)
 *   it.discMode   → 'percent' | 'amount' | '' (string)
 *   it.discVal    → discount value (number; % if mode=percent, $ if amount)
 *
 * Returns: integer cents, clamped to >= 0
 */
function _vendoo_computeFinalUnitCents_(it) {
  const unit = Math.max(0, Number(it && it.price) || 0);
  const qty  = Math.max(1, Number(it && it.qty)   || 1);

  // ---- Line-level discount amount (for the WHOLE line) ----
  const rawMode = (it && (it.discMode ?? it.discountType ?? it.disc_type ?? it.discount_mode)) ?? '';
  const rawVal  = (it && (it.discVal  ?? it.discountValue ?? it.disc_value ?? it.discount_amount)) ?? 0;

  // Normalize mode + value
  let mode = String(rawMode).trim().toLowerCase();                  // 'percent' | 'amount' | ''
  let valNum = 0;

  if (typeof rawVal === 'number') {
    valNum = rawVal;
  } else {
    // handle strings like "5", "5%", " 5 % ", "0.05"
    const s = String(rawVal).trim();
    const m = s.match(/(\d+(?:\.\d+)?)/);
    if (m) {
      valNum = Number(m[1]);
      if (!s.includes('%') && valNum <= 1 && mode === 'percent') {
        // user sent 0.05 with mode=percent => 5%
        valNum = valNum * 100;
      }
    }
  }
  const val = Number(valNum) || 0;

  // Line base (whole line, dollars)
  const lineBase = unit * qty;

  let lineDisc = 0; // dollars
  if (mode === 'percent' && isFinite(val) && val > 0) {
    lineDisc = (lineBase * Math.min(val, 100)) / 100;
  } else if (mode === 'amount' && isFinite(val) && val > 0) {
    lineDisc = val;
  }
  // Clamp to [0, lineBase]
  if (lineDisc < 0) lineDisc = 0;
  if (lineDisc > lineBase) lineDisc = lineBase;

  // Net after line discount (dollars)
  const lineNet = Math.max(0, lineBase - lineDisc);

  // ---- (Future) subtotal discount allocation hook ----
  // When ticket-level/subtotal discount is re-enabled, allocate it
  // pro-rata by each line's share of the post-line-discount total and
  // subtract the per-line allocation here before dividing by qty.

  // Per-unit (dollars) → cents
  const unitNet = qty > 0 ? (lineNet / qty) : 0;
  const cents   = Math.round(unitNet * 100);

  return Math.max(0, cents);
}
// MRAD-END vendoo-final-unit-price-helper


/**
 * Turn the POS sale payload into Vendoo queue tasks (one per SKU line).
 * Each task needs: sku, vendoo_url, price_cents, qty
 * Now uses **final per-unit price after line-item discounts**.
 */
function _vendoo_buildTasksFromSale_(p) {
  const items = Array.isArray(p && p.items) ? p.items : [];
  const out = [];

  try { Logger.log('[VendooQ] BuildTasks: items=%s', items.length); } catch (_){}

  items.forEach(function(it, idx){
    const sku = String((it && it.sku) || '').trim();
    if (!sku) return; // skip MISC etc.

    const qty = Math.max(1, Number((it && it.qty) || 1));

    // NEW: compute the discounted per-unit cents
    const unitFinalCents = _vendoo_computeFinalUnitCents_(it);

    const url = String((it && it.vendooUrl) || '').trim(); // POS passes this through

    out.push({
      sku,
      vendoo_url: url,
      price_cents: unitFinalCents,  // << use discounted per-unit cents
      qty
    });

    // verbose trace (your preferred style)
    try {
      Logger.log(
        '[VendooQ] Line %s sku=%s qty=%s price=$%s discMode=%s discVal=%s → finalUnitCents=%s',
        idx, sku, qty,
        (Number(it && it.price) || 0).toFixed(2),
        String((it && it.discMode) || ''),
        String((it && it.discVal) || ''),
        unitFinalCents
      );
    } catch(_){}
  });

  try { Logger.log('[VendooQ] BuildTasks: built=%s', out.length); } catch(_){}

  return out;
}

/**
 * Enqueue a batch of tasks. Optionally tag all as blocked (e.g., missing_url),
 * which lets us surface them later for manual action without worker claiming them.
 */
function apiVendoo_EnqueueTasks(payload) {
  const sh = ensureVendooTaskSheet_();
  const list = Array.isArray(payload?.tasks) ? payload.tasks : [];
  const saleId = String(payload?.sale_id || '').trim();
  try { Logger.log(`[VendooQ] Enqueue start saleId=${saleId} count=${list.length}`); } catch(_){}

  if (!list.length) {
    try { Logger.log('[VendooQ] Enqueue: no tasks'); } catch(_){}
    return { ok: false, error: 'No tasks' };
  }

  const blockedReason = String(payload?.blocked || '').trim(); // e.g., 'missing_url'
  const isBlocked = !!blockedReason;
  if (isBlocked) { try { Logger.log(`[VendooQ] Enqueue: marking as blocked (${blockedReason})`); } catch(_){}} 

  const c = {
    id: _findCol_(sh, 'Task ID'), st: _findCol_(sh, 'Status'), lease: _findCol_(sh, 'Lease Until'),
    created: _findCol_(sh, 'Created At'), sale: _findCol_(sh, 'Sale ID'), sku: _findCol_(sh, 'SKU'),
    url: _findCol_(sh, 'Vendoo URL'), price: _findCol_(sh, 'Price Cents'), qty: _findCol_(sh, 'Qty'),
    att: _findCol_(sh, 'Attempts'), err: _findCol_(sh, 'Last Error')
  };

  const rows = list.map(t => ([
    _id_(),
    isBlocked ? 'blocked' : 'pending',
    '',
    _now_(),
    saleId,
    String(t.sku || '').trim(),
    String(t.vendoo_url || '').trim(),
    Math.max(0, Number(t.price_cents) || 0),
    Math.max(1, Number(t.qty) || 1),
    0,
    isBlocked ? blockedReason : ''
  ]));

  const before = sh.getLastRow() || 1;
  sh.insertRowsAfter(before, rows.length);
  const startRow = sh.getLastRow() - rows.length + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  try { Logger.log(`[VendooQ] Enqueue: wrote ${rows.length} row(s) @ ${startRow}..${startRow + rows.length - 1}`); } catch(_){}

  return { ok: true, queued: rows.length, blocked: isBlocked ? rows.length : 0 };
}

/**
 * Atomically lease the next pending task with a valid URL.
 * Skips blocked rows and anything with too many attempts.
 */
function apiVendoo_ClaimNextTask() {
  const lock = LockService.getScriptLock(); lock.tryLock(5000);
  try {
    const sh = ensureVendooTaskSheet_();
    const last = sh.getLastRow();
    try { Logger.log(`[VendooQ] Claim: scanning rows=${last-1}`); } catch(_){}
    if (last <= 1) return { ok: true, task: null };

    const c = {
      id: _findCol_(sh, 'Task ID'), st: _findCol_(sh, 'Status'), lease: _findCol_(sh, 'Lease Until'),
      created: _findCol_(sh, 'Created At'), sale: _findCol_(sh, 'Sale ID'), sku: _findCol_(sh, 'SKU'),
      url: _findCol_(sh, 'Vendoo URL'), price: _findCol_(sh, 'Price Cents'), qty: _findCol_(sh, 'Qty'),
      att: _findCol_(sh, 'Attempts'), err: _findCol_(sh, 'Last Error')
    };

    const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const now = _now_(), leaseUntil = new Date(now.getTime() + 2 * 60 * 1000);

    for (let i = 0; i < vals.length; i++) {
      const row = 2 + i;
      const st = String(vals[i][c.st - 1] || '').trim().toLowerCase();
      const leaseStr = vals[i][c.lease - 1];
      const attempts = Number(vals[i][c.att - 1]) || 0;
      const url = String(vals[i][c.url - 1] || '').trim();

      // Only 'pending' rows; require URL; limit attempts
      if (st !== 'pending') continue;
      if (!url) continue;                 // don't lease blank-URL (those are 'blocked' instead)
      if (attempts >= 6) continue;

      const leased = (leaseStr && (new Date(leaseStr))) || null;
      if (leased && leased > now) continue; // still leased by someone else

      // Lease it
      sh.getRange(row, c.st).setValue('in_progress');
      sh.getRange(row, c.lease).setValue(leaseUntil);

      // Return the minimal task object Tampermonkey expects
      const task = {
        row,
        task_id: String(vals[i][c.id - 1] || ''),
        sale_id: String(vals[i][c.sale - 1] || ''),
        sku:     String(vals[i][c.sku - 1] || ''),
        vendoo_url: url,
        price_cents: Number(vals[i][c.price - 1]) || 0,
        qty: Number(vals[i][c.qty - 1]) || 1
      };
      try { Logger.log(`[VendooQ] Claim: leased row=${row} sku=${task.sku} task_id=${task.task_id}`); } catch(_){}
      return { ok: true, task };
    }
    return { ok: true, task: null };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Mark a task complete and (optionally) store the Vendoo item id / url.
 * Phase 4 will flip SKU Tracker status to "Delisted"; for now, just complete.
 */
function apiVendoo_CompleteTask(payload) {
  const sh = ensureVendooTaskSheet_();
  const row = Number(payload?.row||0); if (!row) return { ok:false, error:'Missing row' };
  try { Logger.log(`[VendooQ] Complete: row=${row} sku=${payload?.sku||''} task_id=${payload?.task_id||''}`); } catch(_){}

  const cSt = _findCol_(sh,'Status'), cLease=_findCol_(sh,'Lease Until'), cErr=_findCol_(sh,'Last Error');
  sh.getRange(row, cSt).setValue('done');
  sh.getRange(row, cLease).setValue('');
  sh.getRange(row, cErr).setValue('');

  apiVendoo_SetDelisted({
    sku: payload?.sku,
    sale_id: payload?.sale_id,
    vendoo_item_id: payload?.vendoo_item_id || '',
    vendoo_item_url: payload?.vendoo_item_url || ''
  });

  return { ok:true };
}

/**
 * Mark a task failed and increment attempts with an error message.
 */
function apiVendoo_FailTask(payload) {
  const sh = ensureVendooTaskSheet_();
  const row = Number(payload?.row) || 0;
  const id  = String(payload?.task_id || '').trim();
  const msg = String(payload?.error || 'unknown');

  if (!row || !id) return { ok:false, error:'row/task_id required' };

  const c = {
    id: _findCol_(sh, 'Task ID'), st: _findCol_(sh, 'Status'), lease: _findCol_(sh, 'Lease Until'),
    att: _findCol_(sh, 'Attempts'), err: _findCol_(sh, 'Last Error')
  };

  const curId = String(sh.getRange(row, c.id).getValue() || '').trim();
  if (curId !== id) return { ok:false, error:'mismatch' };

  const attempts = Number(sh.getRange(row, c.att).getValue()) || 0;
  sh.getRange(row, c.att).setValue(attempts + 1);
  sh.getRange(row, c.st).setValue('pending'); // return to pool (lease window will avoid instant re-claim)
  sh.getRange(row, c.lease).setValue('');     // drop lease
  sh.getRange(row, c.err).setValue(msg);
  return { ok:true };
}



function searchInventory(query) {
  const sh = sheet_(SHEET_NAME); // "SKU Tracker"
  const last = sh.getLastRow();
  if (last <= HEADER_ROW) return [];

  const q = String(query || '').trim();
  if (!q) return [];

  // Read A..I (9 cols) to include Status (col I)
  const vals = sh.getRange(2, 1, last - 1, 9).getValues();

  // Vendoo columns are optional; detect by header name
  const header  = sh.getRange(1, 1, 1, Math.max(9, sh.getLastColumn())).getValues()[0].map(String);
  const colVId  = header.indexOf('Vendoo_Item_Number') + 1;
  const colVUrl = header.indexOf('Vendoo_ITEM_URL') + 1;

  const vendooIds  = colVId  > 0 ? sh.getRange(2, colVId,  last - 1, 1).getValues().flat()  : [];
  const vendooUrls = colVUrl > 0 ? sh.getRange(2, colVUrl, last - 1, 1).getValues().flat() : [];

  const qLower = q.toLowerCase();
  const isExactSku = q.includes('-');
  const out = [];

  // IMPORTANT: indexed loop so vendooIds[i] / vendooUrls[i] line up with vals[i]
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    // A:SKU, B:Category, C:Name, D:Price, E:Qty, F:Store, G:Case/Bin/Shelf, H:Online, I:Status
    const sku          = row[0];
    const name         = row[2];
    const price        = row[3];
    const qty          = row[4];
    const store        = row[5];
    const caseBinShelf = row[6];
    const status       = row[8];

    const skuStr  = String(sku  || '').trim();
    const nameStr = String(name || '').trim();

    // inventory gating
    const qtyNum    = Number(qty) || 0;
    const statusStr = String(status || '').trim().toLowerCase();
    const isBlocked =
      statusStr === 'sold' ||
      statusStr === 'sold out' ||
      statusStr === 'archived' ||
      statusStr === 'inactive';
    if (qtyNum <= 0 || isBlocked) continue;

    // Vendoo metadata for this row (if columns exist)
    const vendooId  = (colVId  > 0 ? String(vendooIds[i]  || '') : '');
    const vendooUrl = (colVUrl > 0 ? String(vendooUrls[i] || '') : '');

    const push = () => out.push({
      sku: skuStr,
      name: nameStr,
      price: Number(price) || 0,
      qty: qtyNum,
      storeLoc: String(store || ''),
      caseBinShelf: String(caseBinShelf || ''),
      vendooId,
      vendooUrl
    });

    if (isExactSku) {
      if (skuStr.toLowerCase() === qLower) push();
    } else {
      if (skuStr.toLowerCase().includes(qLower) || nameStr.toLowerCase().includes(qLower)) push();
    }

    if (out.length >= 25) break;
  }

  return out;
}

function saveSale(payload) {
  if (!payload || !Array.isArray(payload.items) || payload.items.length === 0) {
    throw new Error('No items to save.');
  }
  ensureSalesLogHeaders_();

  const acceptedPayments = [
    'Cash','Visa','Mastercard','American Express','Discover','Venmo','Zelle','Cashapp','Paypal'
  ];

  const taxRate = getSalesTaxRate_();

  // --- totals ---
  let rawSubtotal = 0, lineDiscountTotal = 0, netAfterLineDiscounts = 0;
  payload.items.forEach(it => {
    const qty = Math.max(1, Number(it.qty) || 1);
    const price = Math.max(0, Number(it.price) || 0);
    const base = qty * price;
    rawSubtotal += base;

    let disc = 0;
    if (it.discountType === 'percent') disc = base * (Number(it.discountValue) || 0) / 100;
    else if (it.discountType === 'amount')  disc = Number(it.discountValue) || 0;
    disc = Math.min(Math.max(disc, 0), base);

    lineDiscountTotal += disc;
    netAfterLineDiscounts += (base - disc);
  });

  let subtotalDiscount = 0;
  if (payload.subtotalDiscount && payload.subtotalDiscount.type) {
    const d = payload.subtotalDiscount;
    if (d.type === 'percent') subtotalDiscount = netAfterLineDiscounts * (Number(d.value)||0) / 100;
    if (d.type === 'amount')  subtotalDiscount = Number(d.value)||0;
    subtotalDiscount = Math.min(Math.max(subtotalDiscount, 0), netAfterLineDiscounts);
  }

  const subtotal = +(netAfterLineDiscounts - subtotalDiscount).toFixed(2);
  const tax      = +(subtotal * taxRate).toFixed(2);
  const total    = +(subtotal + tax).toFixed(2);

  const saleId = 'S' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

  // --- single or split payments ---
  let paymentMethodForColumn = '';
  let payments = null;
  let changeDue = 0;

  if (Array.isArray(payload.payments) && payload.payments.length) {
    payments = payload.payments.map(p => ({
      method: String(p.method || '').trim(),
      amount: Number(p.amount) || 0
    }));
    const bad = payments.find(p => !acceptedPayments.includes(p.method) || p.amount < 0);
    if (bad) throw new Error('Invalid split payment.');

    const paid = +(payments.reduce((s,p)=>s+p.amount,0)).toFixed(2);
    if (paid + 1e-9 < total) throw new Error(`Underpaid: paid ${paid.toFixed(2)} < total ${total.toFixed(2)}.`);

    // Overpay allowed only if Cash present (change due)
    if (paid > total + 1e-9) {
      const hasCash = payments.some(p => p.method === 'Cash');
      if (!hasCash) throw new Error('Overpayment requires a Cash tender (for change).');
      changeDue = +(paid - total).toFixed(2);
      for (let i = 0; i < payments.length; i++) {
        if (payments[i].method === 'Cash') {
          payments[i].amount = +(payments[i].amount - changeDue).toFixed(2);
          break;
        }
      }
    }
    paymentMethodForColumn = (payments.length === 1) ? payments[0].method : 'Split';
  } else {
    if (!acceptedPayments.includes(payload.paymentMethod)) {
      throw new Error('Invalid payment method: ' + payload.paymentMethod);
    }
    paymentMethodForColumn = payload.paymentMethod;
  }

  // Envelope for Sales Log (keeps your structure)
  const envelope = {
    items: payload.items,
    discounts: {
      rawSubtotal: +rawSubtotal.toFixed(2),
      lineDiscountTotal: +lineDiscountTotal.toFixed(2),
      subtotalDiscount: +subtotalDiscount.toFixed(2)
    },
    paymentMethod: paymentMethodForColumn,
    payments: payments || null,
    changeDue: changeDue,
    fees: Number(payload.fees || 0),
    clerk: payload.clerk || '',
    valor: payload.valor || null  // carry any valor metadata you passed from POS
  };

  const row = [
    new Date(), saleId,
    +rawSubtotal.toFixed(2),
    +lineDiscountTotal.toFixed(2),
    +subtotalDiscount.toFixed(2),
    +subtotal.toFixed(2),
    tax, total,
    paymentMethodForColumn,
    Number(payload.fees || 0),
    JSON.stringify(envelope),
    payload.clerk || ''
  ];

  const log = sheet_('Sales Log');
  log.appendRow(row);

  /* ==========================================================
     Decrement inventory in "SKU Tracker" — UPDATE ONLY E & I
     ========================================================== */
  try {
    // Build { SKU -> deltaQty } (negative)
    const deltas = {};
    for (const it of (payload.items || [])) {
      const sku = String(it.sku || '').trim();
      if (!sku) continue; // ignore MISC
      const q = Math.max(1, Number(it.qty) || 1);
      deltas[sku] = (deltas[sku] || 0) - q;
    }

    const affectedSkus = Object.keys(deltas);
    if (affectedSkus.length) {
      const lock = LockService.getDocumentLock();
      lock.waitLock(30000);

      try {
        const sh = sheet_('SKU Tracker');
        const last = sh.getLastRow();
        if (last > 1) {
          // Read once (A:I) so we can see current qty & status
          const rng  = sh.getRange(2, 1, last - 1, 9);
          const vals = rng.getValues();
          const COL = { SKU:0, CATEGORY:1, NAME:2, PRICE:3, QTY:4, STORE:5, CASE:6, ONLINE:7, STATUS:8 };

          // Index SKUs
          const idxBySku = new Map();
          for (let i = 0; i < vals.length; i++) {
            const s = String(vals[i][COL.SKU] || '').trim();
            if (s) idxBySku.set(s, i);
          }

          // Apply deltas, but **write back only E (Qty) and I (Status)**
          for (const sku of affectedSkus) {
            const i = idxBySku.get(sku);
            if (i == null) continue;
            const cur = Math.max(0, Number(vals[i][COL.QTY]) || 0);
            let next = cur + (Number(deltas[sku]) || 0); // negative
            if (next < 0) next = 0;

            if (next !== cur) {
              vals[i][COL.QTY] = next;
              if (next === 0) vals[i][COL.STATUS] = 'Sold';

              // Write only columns 5..9 for this row (Qty..Status),
              // which avoids touching Category (col B) and its validation.
              sh.getRange(2 + i, 5, 1, 5).setValues([[
                vals[i][COL.QTY],   // E
                vals[i][COL.STORE], // F
                vals[i][COL.CASE],  // G
                vals[i][COL.ONLINE],// H
                vals[i][COL.STATUS] // I
              ]]);
            }
          }
        }
      } finally {
        try { lock.releaseLock(); } catch (_) {}
      }
    }
  } catch (invErr) {
    // Don’t block the sale if inventory write fails; surface in logs
    Logger.log('Inventory decrement failed: ' + (invErr && invErr.message ? invErr.message : invErr));
  }

  /* ==========================================================
    Vendoo: flag sold SKUs as "Pending Removal" (non-blocking)
    ========================================================== */
  try {
    const soldSkus = (payload.items || [])
      .map(it => String(it.sku || '').trim())
      .filter(Boolean);
    if (soldSkus.length) {
      setVendooStatusForSkus_(soldSkus, 'Pending Removal');
    }
  } catch (vendErr) {
    // Do not block the sale; we’ll surface this later in UI in Phase 4
    Logger.log('Vendoo pending-flag failed: ' + (vendErr && vendErr.message ? vendErr.message : vendErr));
  }


  return { saleId, subtotal, tax, total, taxRate };
}

/** Wrapper for POS: forwards to saveSale(...), then enqueues Vendoo tasks (non-blocking). */
function apiSale_Save(payloadJson) {
  const logs = [];
  try {
    const p = (typeof payloadJson === 'string') ? JSON.parse(payloadJson || '{}') : payloadJson;
    logs.push('[apiSale_Save] start');

    // 1) Write the sale row (single source of truth)
    const res = saveSale(p); // { saleId, subtotal, tax, total, taxRate }
    logs.push(`[apiSale_Save] wrote sale row saleId=${res.saleId}`);

    // 2) Link Valor session if present (unchanged)
    try {
      if (p && p.valor && p.valor.invoice) {
        const invoice = String(p.valor.invoice).slice(0, 24);
        const attempt = String(p.valor.attempt || 'A1');
        upsertValorSession_({
          invoice: invoice,
          attempt: attempt,
          status: 'pending',
          saleId: res.saleId,
          posPayload: p
        });
        logs.push(`[apiSale_Save] linked Valor invoice=${invoice} attempt=${attempt}`);
      } else {
        logs.push('[apiSale_Save] no Valor payload present');
      }
    } catch (e) {
      logs.push('[apiSale_Save] upsertValorSession_ failed: ' + e);
    }

    // 3) Vendoo manual mode (no auto-enqueue)
    // We only set SKU Tracker → "Pending Removal" (already done above) and
    // append a manual delist entry so POS can show it in the "Vendoo Delist Queue".
    try {
      const added = apiVendooManual_AddFromSale_({
        sale_id: res.saleId,
        sale_ts: _now_(),
        items: Array.isArray(p.items) ? p.items : []
      });
      L(`[apiSale_Save] manual delist entries added: ${added}`);
    } catch (e) {
      L('[apiSale_Save] manual delist add failed: ' + (e && e.message ? e.message : e));
    }

    return { ...res, logs };
  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e), logs };
  }
}

function getSalesLogRows_() {
  const sh = sheet_('Sales Log');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(2, 1, lastRow - 1, 12).getValues(); // A..L
  return rng.map(r => ({
    ts: r[0], saleId: r[1],
    rawSubtotal: Number(r[2])||0,
    lineDisc: Number(r[3])||0,
    subDisc: Number(r[4])||0,
    subtotal: Number(r[5])||0,
    tax: Number(r[6])||0,
    total: Number(r[7])||0,
    payment: String(r[8]||''),
    fees: Number(r[9])||0,
    itemsJson: r[10],
    clerk: String(r[11]||'')
  }));
}

function buildDailySummary_(rows) {
  const dayMap = new Map();
  rows.forEach(x => {
    if (!(x.ts instanceof Date)) return;
    const d = new Date(x.ts.getFullYear(), x.ts.getMonth(), x.ts.getDate());
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!dayMap.has(key)) {
      dayMap.set(key, { date:d, raw:0, line:0, sub:0, subAfter:0, tax:0, total:0, fees:0, count:0 });
    }
    const a = dayMap.get(key);
    a.raw += x.rawSubtotal; a.line += x.lineDisc; a.sub += x.subDisc;
    a.subAfter += x.subtotal; a.tax += x.tax; a.total += x.total; a.fees += x.fees; a.count += 1;
  });
  const arr = [...dayMap.values()].sort((a,b)=>a.date-b.date);
  return arr.map(a => ([
    a.date, +a.raw.toFixed(2), +a.line.toFixed(2), +a.sub.toFixed(2),
    +a.subAfter.toFixed(2), +a.tax.toFixed(2), +a.total.toFixed(2),
    +a.fees.toFixed(2), a.count
  ]));
}

function buildPaymentSummary_(rows) {
  const pm = new Map();
  rows.forEach(x => {
    const key = x.payment || '(blank)';
    if (!pm.has(key)) pm.set(key, { raw:0, line:0, sub:0, subAfter:0, tax:0, total:0, fees:0, count:0 });
    const a = pm.get(key);
    a.raw += x.rawSubtotal; a.line += x.lineDisc; a.sub += x.subDisc;
    a.subAfter += x.subtotal; a.tax += x.tax; a.total += x.total; a.fees += x.fees; a.count += 1;
  });
  const arr = [...pm.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  return arr.map(([name,a]) => ([
    name, +a.raw.toFixed(2), +a.line.toFixed(2), +a.sub.toFixed(2),
    +a.subAfter.toFixed(2), +a.tax.toFixed(2), +a.total.toFixed(2),
    +a.fees.toFixed(2), a.count
  ]));
}

function refreshSalesSummary() {
  const ss = book_();
  const title = 'Sales Summary';
  const sh = ss.getSheetByName(title) || ss.insertSheet(title);
  sh.clear();
  sh.setFrozenRows(0);

  const moneyFmt = '$#,##0.00';
  const dateFmt  = 'm/d/yyyy';
  const rows = getSalesLogRows_();

  sh.getRange(1,1).setValue('Daily Summary');
  const dailyHeader = ['Date','Raw Subtotal','Line Discounts','Subtotal Discount','Subtotal','Tax','Total','Fees','#Sales'];
  sh.getRange(2,1,1,dailyHeader.length).setValues([dailyHeader]).setFontWeight('bold');

  const dailyData = buildDailySummary_(rows);
  if (dailyData.length) {
    sh.getRange(3,1,dailyData.length,dailyHeader.length).setValues(dailyData);
    sh.getRange(3,1,dailyData.length,1).setNumberFormat(dateFmt);
    sh.getRange(3,2,dailyData.length,6).setNumberFormat(moneyFmt);
    sh.getRange(3,8,dailyData.length,1).setNumberFormat(moneyFmt);
  } else {
    sh.getRange(3,1).setValue('No sales yet.');
  }

  const startRow = 3 + Math.max(1, dailyData.length) + 1;
  sh.getRange(startRow,1).setValue('Payment Method Summary');
  const payHeader = ['Payment Method','Raw Subtotal','Line Discounts','Subtotal Discount','Subtotal','Tax','Total','Fees','#Sales'];
  sh.getRange(startRow+1,1,1,payHeader.length).setValues([payHeader]).setFontWeight('bold');

  const payData = buildPaymentSummary_(rows);
  if (payData.length) {
    sh.getRange(startRow+2,1,payData.length,payHeader.length).setValues(payData);
    sh.getRange(startRow+2,2,payData.length,6).setNumberFormat(moneyFmt);
    sh.getRange(startRow+2,8,payData.length,1).setNumberFormat(moneyFmt);
  } else {
    sh.getRange(startRow+2,1).setValue('No sales yet.');
  }

  sh.autoResizeColumns(1, 9);
  sh.setFrozenRows(2);
}


/*** === POS: Sales list APIs (Today + Custom Range) === ***/

// Utilities to compute start/end of a day in store timezone
function _startOfDay_(d, tz) {
  const z = tz || CASH_TZ;
  const y = +Utilities.formatDate(d, z, 'yyyy');
  const m = +Utilities.formatDate(d, z, 'M') - 1;  // 0-based month
  const day = +Utilities.formatDate(d, z, 'd');
  return new Date(y, m, day, 0, 0, 0, 0);
}
function _endOfDay_(d, tz) {
  const z = tz || CASH_TZ;
  const y = +Utilities.formatDate(d, z, 'yyyy');
  const m = +Utilities.formatDate(d, z, 'M') - 1;
  const day = +Utilities.formatDate(d, z, 'd');
  return new Date(y, m, day, 23, 59, 59, 999);
}

// Internal: normalize [start, end] from ISO strings ("yyyy-mm-dd") or Date objects
function _normalizeRange_(startOpt, endOpt, tz) {
  tz = tz || CASH_TZ;
  let start, end;

  if (startOpt && endOpt) {
    // both yyyy-mm-dd
    start = _startOfDay_(new Date(startOpt + 'T12:00:00'), tz);
    end   = _endOfDay_(  new Date(endOpt   + 'T12:00:00'), tz);
  } else {
    const now = new Date();
    start = _startOfDay_(now, tz);
    end   = _endOfDay_(now, tz);
  }
  if (end < start) { const t = start; start = end; end = t; }
  return { start, end, tz };
}

// Parse "SYYYYMMDDHHmmss" into a Date (local time). Returns null if not parseable.
function _rp_saleIdToDate_(saleId) {
  const m = /^S(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})$/.exec(String(saleId||'').trim());
  if (!m) return null;
  const y=+m[1], mo=+m[2]-1, d=+m[3], hh=+m[4], mm=+m[5], ss=+m[6];
  return new Date(y, mo, d, hh, mm, ss, 0);
}

// Build compact row for POS (derive time from ts or saleId; render split payments like Refund)
function _rp_buildPosRow_(r, tz) {
  const zone = String(tz || CASH_TZ || Session.getScriptTimeZone() || 'America/Denver');

  // Prefer real timestamp; fallback to Sale ID parsing
  const ts = (r.ts instanceof Date) ? r.ts : _rp_saleIdToDate_(r.saleId);
  const display = ts ? Utilities.formatDate(ts, zone, 'M/d h:mm a') : '';

  // Match Refunds' payment rendering (split payments supported)
  let payText = r.payment || '';
  try {
    const env = (typeof parseJSONLoose_ === 'function') ? parseJSONLoose_(r.itemsJson) : null;
    if (env && Array.isArray(env.payments) && env.payments.length) {
      payText = env.payments.map(p => `${String(p.method||p.type||'')} $${(+p.amount||0).toFixed(2)}`).join(' + ');
    } else if (env && env.paymentMethod && !payText) {
      payText = env.paymentMethod;
    }
  } catch(_) {}

  // IMPORTANT: do not return Date objects (JSON-safe only)
  return {
    displayTime: display,
    saleId: r.saleId || '',
    payment: payText || '',
    total: Number(r.total)||0,
    clerk: r.clerk || ''
  };
}

/** Return sales for TODAY (store timezone). */
function apiSales_ListToday() {
  const logs = [];
  try {
    const tz = Session.getScriptTimeZone();
    const dayKey = 'S' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
    const all = getSalesLogRows_();
    logs.push(`rows=${all.length}`, `tz=${tz}`, `dayKey=${dayKey}`,
              `sampleIds=${all.slice(Math.max(0, all.length - 5)).map(r => r.saleId).join(',')}`);

    const sales = all
      .filter(r => String(r.saleId || '').startsWith(dayKey))
      .map(r => _rp_buildPosRow_(r, tz))
      .sort((a, b) => String(b.saleId).localeCompare(String(a.saleId))); // newest first

    return { ok: true, sales, info: { dayKey, tz, spreadsheetId: (typeof SPREADSHEET_ID !== 'undefined' ? SPREADSHEET_ID : '') }, logs };
  } catch (e) {
    logs.push('error=' + e.message);
    return { ok: false, error: e.message, logs };
  }
}

/** Sales list (Custom Range) — inclusive date range; returns logs */
function apiSales_ListRange(fromISO, toISO) {
  const logs = [];
  try {
    // inclusive day range
    const from = new Date(fromISO + 'T00:00:00');
    const to   = new Date(toISO   + 'T23:59:59');
    if (!(from instanceof Date) || isNaN(+from) || !(to instanceof Date) || isNaN(+to)) {
      return { ok: false, error: 'Invalid date(s).', logs };
    }
    const tz = Session.getScriptTimeZone();
    const all = getSalesLogRows_();
    logs.push(`rows=${all.length}`, `from=${from.toISOString()}`, `to=${to.toISOString()}`);

    const sales = all
      .filter(r => {
        const t = (r.ts instanceof Date) ? r.ts : new Date(r.ts);
        return t && !isNaN(+t) && t >= from && t <= to;
      })
      .map(r => _rp_buildPosRow_(r, tz))
      .sort((a, b) => String(b.saleId).localeCompare(String(a.saleId))); // newest first

    return { ok: true, sales, info: { from: fromISO, to: toISO, spreadsheetId: (typeof SPREADSHEET_ID !== 'undefined' ? SPREADSHEET_ID : '') }, logs };
  } catch (e) {
    logs.push('error=' + e.message);
    return { ok: false, error: e.message, logs };
  }
}

/**
 * Ensure "Vendoo Manual Queue" sheet exists and header is correct.
 * Columns: Added At | Sale ID | Sale TS | SKU | Qty | Price Cents | Vendoo URL | Status | Note
 */
function ensureVendooManualSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName('Vendoo Manual Queue');
  if (!sh) sh = ss.insertSheet('Vendoo Manual Queue');

  // Base header (existing 9 cols)
  const baseHeader = [
    'Added At','Sale ID','Sale TS','SKU','Qty','Price Cents','Vendoo URL','Status','Note'
  ];

  // Read current header
  const lastCol = Math.max(1, sh.getLastColumn());
  const cur = sh.getRange(1, 1, 1, Math.max(lastCol, baseHeader.length)).getValues()[0].map(String);

  // If first-time or header mismatched, reset to base header
  const okBase = baseHeader.every((h, i) => (cur[i] || '') === h);
  if (!okBase) {
    sh.getRange(1, 1, 1, baseHeader.length).setValues([baseHeader]);
    sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm'); // Added At
    sh.getRange('C:C').setNumberFormat('m/d/yyyy h:mm'); // Sale TS
  }

  // Ensure appended "Item Title" column exists (append by name; do not shift earlier columns)
  const headerAfter = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), baseHeader.length)).getValues()[0].map(String);
  if (!headerAfter.includes('Item Title')) {
    const col = (sh.getLastColumn() || baseHeader.length) + 1;
    sh.getRange(1, col).setValue('Item Title');
  }
  return sh;
}

/**
 * Lookup Online Location and Vendoo URL for a SKU from "SKU Tracker".
 * Returns { onlineLoc, vendooUrl }
 */
function _skuVendooMeta_(sku) {
  const sh = sheet_('SKU Tracker');
  const last = sh.getLastRow();
  if (last < 2) return { onlineLoc: '', vendooUrl: '', title: '' };

  const hc = {};
  (function mapHeaders(){
    const row = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
    row.forEach((h, i) => hc[String(h).trim()] = i+1);
  })();

  const cSKU = hc['SKU'] || 0;

  // Accept common variants for these headers
  const cLoc = hc['Online Location'] || hc['Online'] || hc['OnlineLocation'] || 0;
  const cUrl = hc['Vendoo_ITEM_URL'] || hc['Vendoo Item URL'] || hc['Vendoo URL'] || 0;

  // Prefer a title/name-like column if present; fall back to classic column C:Name
  const cTitle =
    hc['Title'] || hc['Short Description'] || hc['Name'] || 0;

  if (!cSKU) return { onlineLoc: '', vendooUrl: '', title: '' };

  // Build an index for SKU -> row index (2-based row number)
  const skuVals = sh.getRange(2, cSKU, last - 1, 1).getValues().flat();
  const idxBySku = new Map();
  for (let i = 0; i < skuVals.length; i++) {
    const s = String(skuVals[i] || '').trim();
    if (s) idxBySku.set(s, 2 + i);
  }

  const row = idxBySku.get(String(sku || '').trim());
  if (!row) return { onlineLoc: '', vendooUrl: '', title: '' };

  const onlineLoc = cLoc ? String(sh.getRange(row, cLoc).getValue() || '') : '';
  const vendooUrl = cUrl ? String(sh.getRange(row, cUrl).getValue() || '') : '';
  const title     = cTitle ? String(sh.getRange(row, cTitle).getValue() || '') : '';

  return { onlineLoc, vendooUrl, title };
}

/**
 * Append manual delist entries for items in a sale.
 * Expects payload: { sale_id, sale_ts, items:[{sku, qty, price, vendooUrl?}, ...] }
 * Returns number of rows written.
 */
function apiVendooManual_AddFromSale_(payload){
  try { Logger.log('[VendooManual] AddFromSale start ' + JSON.stringify(payload).slice(0,300)); } catch(_){}
  const sh = ensureVendooManualSheet_();  // guarantees header in row 1

  const saleId = String(payload && payload.sale_id || '').trim();
  const saleTs = payload && payload.sale_ts ? new Date(payload.sale_ts) : _now_();
  const items  = Array.isArray(payload && payload.items) ? payload.items : [];
  if (!items.length) return 0;

  const rows = [];
  for (const it of items) {
    const sku = String(it.sku || '').trim();
    if (!sku) { try{ Logger.log('[VendooManual] skip: no sku in item ' + JSON.stringify(it)); }catch(_){ }
                continue; }

    const qty         = Math.max(1, Number(it.qty)||1);
    const priceCents  = _vendoo_computeFinalUnitCents_(it);
    const meta        = _skuVendooMeta_(sku);
    const vendooUrl = String((it.vendooUrl || meta.vendooUrl) || '').trim();
    const locRaw    = String(meta.onlineLoc || '').trim();
    const loc       = locRaw.toLowerCase();

    // Consider several “Vendoo-ish” signals from the location, even if URL is blank
    const isVendooLoc = ['vendoo', 'vendoo only', 'vendoo-only', 'vendoo_only', 'both']
      .some(k => loc.includes(k));

    // Any Vendoo signal qualifies: location OR a Vendoo URL
    const looksVendoo = isVendooLoc || (/vendoo\.co/i.test(vendooUrl));

    if (looksVendoo) {
      // Append with Item Title column at the end (header ensures it exists)
      const itemTitle =
        String((it.name || it.title || meta.title) || '').trim();
      rows.push([ _now_(), saleId, saleTs, sku, qty, priceCents, vendooUrl, 'pending', '', itemTitle ]);
      try { Logger.log('[VendooManual] queued sku=' + sku + ' loc="' + locRaw + '" url=' + (vendooUrl||'(none)')); } catch(_){}
    } else {
      try { Logger.log('[VendooManual] skip sku=' + sku + ' loc="' + locRaw + '" (no vendoo signal)'); } catch(_){}
    }
  }

  if (!rows.length) { try{ Logger.log('[VendooManual] no rows to write'); }catch(_){ }
                      return 0; }

  // *** append after the header — never write to row 1 ***
  const start = Math.max(2, sh.getLastRow() + 1);   // first data row is 2
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);

  try { Logger.log('[VendooManual] wrote ' + rows.length + ' row(s) @ ' + start + '..' + (start + rows.length - 1)); } catch(_){}
  return rows.length;
}

/**
 * List pending manual delist entries (default: last 72h).
 * Returns: [{ addedAt, saleId, saleTs, sku, qty, price, vendooUrl, ageHours }]
 */
function apiVendoo_ListPending(args){
  const sh = ensureVendooManualSheet_();
  const maxHours = Math.max(1, Number(args && args.maxHours || 72));
  const now = _now_();
  const out = [];

  const hc = {};
  (function mapHeaders(){
    const row = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
    row.forEach((h, i) => hc[String(h).trim()] = i+1);
  })();
  const last = sh.getLastRow();
  if (last <= 1) return out;

  const vals = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const st = new Date(vals[i][(hc['Added At']||1)-1] || now);
    const saleTs = new Date(vals[i][(hc['Sale TS']||1)-1] || st);
    const status = String(vals[i][(hc['Status']||1)-1] || '').toLowerCase();
    if (status !== 'pending') continue;

    const ageH = Math.floor((now - st) / (1000*60*60));
    if (ageH > maxHours) continue;

    out.push({
      addedAt: st,
      saleId: String(vals[i][(hc['Sale ID']||1)-1] || ''),
      saleTs: saleTs,
      sku: String(vals[i][(hc['SKU']||1)-1] || ''),
      qty: Number(vals[i][(hc['Qty']||1)-1] || 1),
      price: Number(vals[i][(hc['Price Cents']||1)-1] || 0) / 100,
      vendooUrl: String(vals[i][(hc['Vendoo URL']||1)-1] || ''),
      title: String(vals[i][(hc['Item Title']||0)-1] || ''), // safe if column missing
      ageHours: ageH
    });
  }
  try { Logger.log(`[VendooManual] list → ${out.length} row(s)`); } catch(_){ }
  return JSON.parse(JSON.stringify(out));
}

/**
 * Helper: mark a SKU as Delisted in "SKU Tracker" and optionally store Vendoo id/url.
 * Payload: { sku, sale_id?, vendoo_item_id?, vendoo_item_url? }
 */
function apiVendoo_SetDelisted(payload) {
  var sku = String(payload && payload.sku || '').trim();
  if (!sku) return { ok:false, error:'Missing sku' };

  var vendooId  = String(payload && payload.vendoo_item_id  || '').trim();
  var vendooUrl = String(payload && payload.vendoo_item_url || '').trim();

  try {
    var sh = sheet_('SKU Tracker');
    var last = sh.getLastRow();
    if (last < 2) return { ok:false, error:'No SKU rows' };

    // Header map
    var header = sh.getRange(1, 1, 1, Math.max(20, sh.getLastColumn())).getValues()[0].map(String);
    var cSKU   = header.indexOf('SKU') + 1;
    var cStat  = header.indexOf('Vendoo_Listing_Status') + 1;
    var cId    = header.indexOf('Vendoo_Item_Number') + 1;   // if present
    var cUrl   = header.indexOf('Vendoo_ITEM_URL') + 1;      // if present
    if (!cSKU) return { ok:false, error:'Missing SKU column' };

    // Scan rows to find the SKU
    var vals = sh.getRange(2, 1, last - 1, Math.max(20, sh.getLastColumn())).getValues();
    var rowIndex = -1;
    for (var i = 0; i < vals.length; i++) {
      var s = String(vals[i][cSKU - 1] || '').trim();
      if (s === sku) { rowIndex = 2 + i; break; }
    }
    if (rowIndex < 0) return { ok:false, error:'SKU not found: ' + sku };

    // Write status
    if (cStat) sh.getRange(rowIndex, cStat).setValue('Delisted');

    // Backfill id/url if columns exist and we have values
    if (cId && vendooId) {
      var curId = String(sh.getRange(rowIndex, cId).getValue() || '').trim();
      if (!curId) sh.getRange(rowIndex, cId).setValue(vendooId);
    }
    if (cUrl && vendooUrl) {
      var curUrl = String(sh.getRange(rowIndex, cUrl).getValue() || '').trim();
      if (!curUrl) sh.getRange(rowIndex, cUrl).setValue(vendooUrl);
    }

    try {
      Logger.log('[Vendoo] SetDelisted sku=%s sale=%s id=%s url=%s row=%s',
        sku, String(payload && payload.sale_id || ''), vendooId, vendooUrl, rowIndex);
    } catch (_){}

    return { ok:true, row: rowIndex };
  } catch (e) {
    return { ok:false, error: String(e && e.message || e) };
  }
}

/**
 * Confirm delist from POS:
 * - Mark SKU Tracker Vendoo_Listing_Status = 'delisted' (existing helper)
 * - Mark manual row as done (or remove)
 */
function apiVendoo_ManualConfirmDelist(payload){
  const sku = String(payload && payload.sku || '').trim();
  const saleId = String(payload && payload.sale_id || '').trim();
  const vendooUrl = String(payload && payload.vendoo_item_url || payload && payload.vendooUrl || '').trim();

  // Flip SKU Tracker via the existing helper
  apiVendoo_SetDelisted({
    sku: sku,
    sale_id: saleId,
    vendoo_item_id: '',
    vendoo_item_url: vendooUrl
  });

  // Mark the first matching manual row as done
  const sh = ensureVendooManualSheet_();
  const last = sh.getLastRow();
  if (last > 1) {
    const hc = {};
    (function mapHeaders(){
      const row = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
      row.forEach((h, i) => hc[String(h).trim()] = i+1);
    })();
    const vals = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
    for (let i=0;i<vals.length;i++){
      const r = i+2;
      const sale = String(vals[i][(hc['Sale ID']||1)-1] || '');
      const s   = String(vals[i][(hc['SKU']||1)-1] || '');
      const stCol = hc['Status'] || 8;
      if (sale === saleId && s === sku && String(vals[i][stCol-1]||'').toLowerCase() === 'pending') {
        sh.getRange(r, stCol).setValue('done');
        break;
      }
    }
  }
  try { Logger.log(`[VendooManual] confirm → sku=${sku} sale=${saleId}`); } catch(_){}
  return { ok:true };
}

/***** ============================================================
 * End POS 
 * ============================================================ *****/

/***** ============================================================
 * Begin Refund / Void APIs — normalized, serializer-safe response
 * ============================================================ *****/

/** Safe JSON parse (returns null on failure/empty). */
function parseJSONLoose_(v) {
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

/**
 * apiSale_Get_v2(saleId)
 * Looks up the sale by exact match in column B ("Sales Log"),
 * returns primitives + normalized items[] and payments[].
 * (Unique name to avoid collisions with older helpers.)
 */
function apiSale_Get_v2(saleId) {
  const sh = sheet_('Sales Log');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('No sales found.');

  const id = String(saleId || '').trim();

  // Exact match in column B
  const cell = sh.getRange('B:B').createTextFinder(id).matchEntireCell(true).findNext();
  if (!cell) throw new Error('Sale not found: ' + id);

  const row = cell.getRow();
  const maxCols = Math.max(12, sh.getLastColumn());
  const r = sh.getRange(row, 1, 1, maxCols).getValues()[0]; // A..L+

  const ts       = r[0];
  const rawSub   = Number(r[2]) || 0;
  const lineDisc = Number(r[3]) || 0;
  const subDisc  = Number(r[4]) || 0;
  const subtotal = Number(r[5]) || 0;
  const tax      = Number(r[6]) || 0;
  const total    = Number(r[7]) || 0;
  const payCol   = r[8] || '';
  const fees     = Number(r[9]) || 0; // reserved
  const env      = parseJSONLoose_(r[10]);
  const clerk    = r[11] || '';
  const taxRate  = subtotal > 0 ? +(tax / subtotal).toFixed(4) : 0;

  // ---- Normalize items
  let items = [];
  if (Array.isArray(env)) { items = env; }
  else if (env && Array.isArray(env.items)) { items = env.items; }
  items = (items || []).map(it => ({
    sku:   String(it.sku || ''),
    name:  String(it.name || ''),
    price: Number(it.price) || 0,
    qty:   Number(it.qty)   || 1
  }));

  // ---- Normalize payments
  let payments = [];
  if (env && Array.isArray(env.payments) && env.payments.length) {
    payments = env.payments.map(p => ({
      method: String(p.method || p.type || p.name || ''),
      amount: Number(p.amount) || 0
    })).filter(p => p.method || p.amount);
  } else {
    const col = String(payCol || '').trim();
    if (col && col.toLowerCase() !== 'split') {
      payments = [{ method: col, amount: total }];
    } else {
      payments = []; // old “Split” rows without breakdown
    }
  }

  const sale = {
    saleId: String(r[1]),
    displayTime: (ts instanceof Date)
      ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'M/d h:mm a') : '',
    paymentMethod: payCol || '',
    clerk: clerk || '',
    subtotal, tax, total, taxRate,
    items, payments
  };

  Logger.log('[apiSale_Get_v2] returning sale=%s',
    JSON.stringify({ saleId: sale.saleId, itemsLen: sale.items.length, payments: sale.payments }));

  return sale;
}

/**
 * apiSale_Refund(payload)
 * Appends a negative entry to "Sales Log" with a refund for the selected lines.
 *
 * Accepts lines with either { qty } or { refundQty }, and also a linesJson fallback.
 * Returns both refund* fields and {subtotal,tax,total} for UI compatibility.
 */
function apiSale_Refund(payload) {
  const pay = payload || {};

  // Prefer array; fallback to linesJson string if needed
  let lines = Array.isArray(pay.lines) ? pay.lines : [];
  if ((!lines || !lines.length) && typeof pay.linesJson === 'string') {
    try { const parsed = JSON.parse(pay.linesJson); if (Array.isArray(parsed)) lines = parsed; } catch(e) {}
  }

  if (!pay.origSaleId) throw new Error('Missing original sale id.');

  // Normalize & sum
  const norm = [];
  let sub = 0;
  (lines || []).forEach(l => {
    const qty = Math.max(0, Math.floor(Number(l.qty ?? l.refundQty) || 0));
    if (!qty) return;
    const price = Number(l.price) || 0;
    norm.push({ sku: String(l.sku||''), name: String(l.name||''), price, qty });
    sub += price * qty;
  });
  if (norm.length === 0) throw new Error('Nothing to refund.');

  // Tax rate: use payload.taxRate, else a helper if you have one
  const taxRate = (typeof pay.taxRate === 'number' && pay.taxRate >= 0)
    ? pay.taxRate
    : (typeof getSalesTaxRate_ === 'function' ? getSalesTaxRate_() : 0);

  const tax   = +(sub * taxRate).toFixed(2);
  const total = +(sub + tax).toFixed(2);

  // Payment text for column I
  let paymentText = 'Refund';
  if (Array.isArray(pay.payments) && pay.payments.length === 1) {
    paymentText = 'Refund: ' + String(pay.payments[0].method || '').trim();
  } else if (Array.isArray(pay.payments) && pay.payments.length > 1) {
    paymentText = 'Refund: Split';
  }

  // Envelope marking this as a refund
  const env = {
    type: 'refund',
    origSaleId: String(pay.origSaleId || ''),
    taxRate: taxRate,
    items: norm, // [{sku,name,price,qty}]
    payments: Array.isArray(pay.payments)
      ? pay.payments.map(p => ({ method: String(p.method||''), amount: Number(p.amount)||0 }))
      : []
  };

  // Write negative row
  const sh = sheet_('Sales Log');
  if (typeof ensureSalesLogHeaders_ === 'function') ensureSalesLogHeaders_();

  const tz = Session.getScriptTimeZone();
  const ts = new Date();
  const refundId = 'R' + Utilities.formatDate(ts, tz, 'yyyyMMddHHmmss');

  sh.appendRow([
    ts,                         // A Timestamp
    refundId,                   // B Sale ID (refund id)
    0,                          // C Raw Subtotal (not used for refunds)
    0,                          // D Line Discounts
    0,                          // E Subtotal Discount
    -sub,                       // F Subtotal (negative)
    -tax,                       // G Tax (negative)
    -total,                     // H Total (negative)
    paymentText,                // I Payment Method
    0,                          // J Fees
    JSON.stringify(env),        // K Envelope JSON
    String(pay.clerk || '')     // L Clerk
  ]);

  Logger.log('[apiSale_Refund] wrote refund row id=%s total=$%s for sale=%s',
    refundId, total.toFixed(2), String(pay.origSaleId || ''));

  // return both naming styles for UI compatibility
  return {
    refundId,
    refundSubtotal: +sub.toFixed(2),
    refundTax:      +tax.toFixed(2),
    refundTotal:    +total.toFixed(2),
    subtotal:       +sub.toFixed(2),
    tax:            +tax.toFixed(2),
    total:          +total.toFixed(2)
  };
}
/***** ============================================================
 * end Refund code
 * ============================================================ *****/



/***** Begin Valor Code – publish + status fallback + compatibility aliases *****/

function valorProps_() {
  const p = PropertiesService.getScriptProperties();
  const env = (p.getProperty('VALOR_ENV') || 'uat').toLowerCase();

  const defaultBase =
    env === 'prod'
      ? 'https://securelink.valorpaytech.com'
      : 'https://securelink-staging.valorpaytech.com';

  return {
    VALOR_ENV: env,
    VALOR_BASE_URL: p.getProperty('VALOR_BASE_URL') || defaultBase,
    VALOR_PUBLISH_URL: p.getProperty('VALOR_PUBLISH_URL') || '',
    VALOR_CHANNEL_ID: p.getProperty('VALOR_CHANNEL_ID') || '',
    VALOR_APP_ID:     p.getProperty('VALOR_APP_ID') || '',
    VALOR_APP_KEY:    p.getProperty('VALOR_APP_KEY') || '',
    VALOR_EPI:        p.getProperty('VALOR_EPI') || '',
    VALOR_WEBHOOK_SECRET: p.getProperty('VALOR_WEBHOOK_SECRET') || ''
  };
}

function valorUuid_(){ return Utilities.getUuid(); }
function valorCentsToDollars_(val){ const n=Number(val); return Number.isFinite(n)?(n/100).toFixed(2):''; }
function valorMask_(s, keepEnd){
  const v = String(s||''); const k = Math.max(0, keepEnd||4);
  if (!v) return '';
  if (v.length <= k) return '*'.repeat(Math.max(0, v.length-1))+v.slice(-1);
  return '*'.repeat(v.length - k) + v.slice(-k);
}

// Auto-create the log tab if missing; append masked request/response rows
function valorLogPublish_(row){
  const ss = book_();
  let sh = ss.getSheetByName('Valor Publish Log');
  if (!sh) sh = ss.insertSheet('Valor Publish Log');

  // Ensure header exists; keep the original 7 columns in the same order,
  // and add H = invoicenumber without shifting earlier columns.
  if (sh.getLastRow() < 1) {
    sh.appendRow(['When','Phase','ReqTxnId','URL','HTTP','Ack/Msg','Payload','invoicenumber']);
  } else {
    // Make sure column H has a header "invoicenumber"
    const lastCol = sh.getLastColumn();
    if (lastCol < 8 || String(sh.getRange(1, 8).getValue()).trim() === '') {
      sh.getRange(1, 8).setValue('invoicenumber');
    }
  }

  // Extract invoicenumber from the logged payload (robust to string/object/nested)
  let inv = '';
  try {
    const p = (typeof row.payload === 'string')
      ? JSON.parse(row.payload || '{}')
      : (row.payload || {});
    inv = String(
      (p && (p.invoicenumber || p.INVOICENUMBER || p.invoice_number)) ||
      (p && p.payload && (p.payload.invoicenumber || p.payload.INVOICENUMBER || p.payload.invoice_number)) ||
      ''
    ).trim();
  } catch (_){ /* ignore parse errors; leave inv = '' */ }

  sh.appendRow([
    new Date(),
    row.phase || '',
    row.reqTxnId || '',
    row.url || '',
    String(row.http || ''),
    (typeof row.ack === 'string' ? row.ack : JSON.stringify(row.ack || row.msg || {})).slice(0, 2000),
    (typeof row.payload === 'string' ? row.payload : JSON.stringify(row.payload || {})).slice(0, 2000),
    inv
  ]);
}

// --- endpoints ---
function valorGetPublishUrl_() {
  const { VALOR_BASE_URL, VALOR_PUBLISH_URL } = valorProps_();
  let u = (VALOR_PUBLISH_URL || VALOR_BASE_URL || '').trim();
  if (!/^https?:\/\//i.test(u)) throw new Error('Valor publish URL is not a valid https URL.');
  const bareHost = /^https?:\/\/[^\/?#]+\/?$/.test(u);
  if (bareHost) {
    // Wrapper Publish path must be ?status
    throw new Error('VALOR_PUBLISH_URL must include the full publish path, e.g. https://securelink.valorpaytech.com:4430/?status');
  }
  u = u.replace(/(\?[^#]*?)\/+$/, '$1');
  return u;
}

// Derive the Status endpoint from your Publish URL and FORCE the `=` suffix.
function valorGetStatusUrl_() {
  const pub = valorGetPublishUrl_(); // e.g. https://securelink.valorpaytech.com:4430/?status
  if (!pub || !/\?status\b/i.test(pub)) {
    throw new Error('VALOR_PUBLISH_URL must include "?status" (e.g. https://...:4430/?status).');
  }
  // Replace everything from ?status to the end with ?txn_status=
  const u = pub.replace(/\?status\b.*/i, '?txn_status=');
  // Sanity: require the exact suffix
  if (!/\?txn_status=$/i.test(u)) {
    throw new Error('Failed to derive vc_status URL ending in "?txn_status=". Check VALOR_PUBLISH_URL.');
  }
  return u;
}

/**
 * Start a card-present charge on the Valor terminal via Wrapper Publish API.
 * Uses txn_type: "vc_publish" with nested "payload" per Publish API.
 *
 * Returns: { ok, status, accepted, reqTxnId, ack }
 */
function apiStartValorCheckout(amountCents, invoiceNumber, lineItems) {
  const {
    VALOR_ENV, VALOR_CHANNEL_ID, VALOR_APP_ID, VALOR_APP_KEY, VALOR_EPI
  } = valorProps_();

  if (!VALOR_CHANNEL_ID || !VALOR_APP_ID || !VALOR_APP_KEY || !VALOR_EPI) {
    throw new Error('Missing Valor credentials. Set VALOR_CHANNEL_ID, VALOR_APP_ID, VALOR_APP_KEY, VALOR_EPI in Script Properties.');
  }

  const url = valorGetPublishUrl_();   // must include ?status
  const reqTxnId = valorUuid_();

  const body = {
    appid:       String(VALOR_APP_ID),
    appkey:      String(VALOR_APP_KEY),
    epi:         String(VALOR_EPI),
    txn_type:    'vc_publish',
    channel_id:  String(VALOR_CHANNEL_ID),
    version:     '1',
    // Fadil's guidance: INVOICENUMBER with no underscore at TOP LEVEL
    INVOICENUMBER: String(invoiceNumber || '').slice(0, 24),
    payload: {
      TRAN_MODE:  '1',                                         // card-present
      TRAN_CODE:  '1',                                         // sale
      AMOUNT:     String(Math.max(0, Number(amountCents) || 0)),// cents
      REQ_TXN_ID: reqTxnId,                                     // still generated, but no longer used for runtime correlation
      INVOICENUMBER:  String(invoiceNumber || '').slice(0, 24),
    }
  };

  if (Array.isArray(lineItems) && lineItems.length) body.lineItems = lineItems;

  const masked = (function(){
    const m = JSON.parse(JSON.stringify(body));
    const mask = (s,k) => valorMask_(s, k);
    m.appid = mask(m.appid, 4);
    m.appkey = mask(m.appkey, 4);
    m.epi = mask(m.epi, 4);
    return m;
  })();

  valorLogPublish_({ phase:'request', reqTxnId, url, http:'→', payload: masked });
  Logger.log('[valor publish →] url=%s env=%s reqTxnId=%s amountCents=%s invoice=%s',
           url, VALOR_ENV, reqTxnId, body.payload.AMOUNT, body.INVOICENUMBER);

  let status = 0, bodyText = '', ack = null, accepted = false;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      muteHttpExceptions: true,
      contentType: 'application/json',
      headers: { 'accept': 'application/json' },
      payload: JSON.stringify(body)
    });
    status = res.getResponseCode();
    bodyText = res.getContentText() || '';
    try { ack = JSON.parse(bodyText); } catch(_) { ack = { raw: bodyText.slice(0,400) }; }

    const asString = JSON.stringify(ack||{}).toLowerCase();
    const looksError = /error|invalid|denied|not\s*found|unauthor/i.test(asString);
    const okHttp = (status >= 200 && status < 300);
    accepted = okHttp && !looksError;

    // VC07: treat as accepted; we'll wait for webhook
    if (!accepted) {
      const vc07 = /vc07/i.test(asString) || /transaction\s*timeout/i.test(asString);
      if (vc07) accepted = true;
    }

    valorLogPublish_({ phase:'response', reqTxnId, url, http: String(status), ack: accepted ? ack : { status, ack } });
    Logger.log('[valor publish ←] status=%s accepted=%s', status, accepted);
  } catch (err) {
    valorLogPublish_({ phase:'network-error', reqTxnId, url, http:'EXC', ack: String(err && err.message || err), payload: masked });
    throw new Error('Network/HTTP error reaching Valor: ' + (err && err.message ? err.message : String(err)));
  }

  if (!accepted) {
    throw new Error('Valor did not accept publish (HTTP ' + status + '). See "Valor Publish Log" for details.');
  }

  // Return the invoice so the front-end can poll on it
  return { ok:true, status, accepted:true, reqTxnId, invoice: String(invoiceNumber || ''), ack };
}
/**
 * Find the publish REQ_TXN_ID (UUID) that matches a webhook.
 * We match primarily by AMOUNT (cents), then EPI, then closeness in time.
 * Scans the last ~120 rows of "Valor Publish Log".
 *
 * @param {number} amountCents - cents from webhook (e.g., 11 for $0.11)
 * @param {string} epiFromHook - webhook epi_id (optional but preferred)
 * @param {Date}   hookWhen    - webhook timestamp (Date) for recency scoring
 * @param {string} invoiceFromHook - if your webhook ever carries an invoice id (often empty)
 * @return {string} reqTxnId or ''
 */
function valorFindReqIdFromPublish_(amountCents, epiFromHook, hookWhen, invoiceFromHook) {
  const sh = sheet_('Valor Publish Log');
  if (!sh) return '';

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return '';

  const startRow = Math.max(2, lastRow - 120);
  const wid = Math.min(12, sh.getLastColumn());
  const vals = sh.getRange(startRow, 1, lastRow - startRow + 1, wid).getValues();

  const hookTs = (hookWhen && hookWhen.getTime) ? +hookWhen : Date.now();
  let best = {score: -1, id: ''};

  for (let i = vals.length - 1; i >= 0; i--) {
    const row = vals[i];
    const phase = String(row[1] || '');          // "request" / "response"
    const reqIdFromCol = String(row[2] || '');   // column C usually stores REQ_TXN_ID
    if (phase !== 'request') continue;

    // Time
    const when = row[0] instanceof Date ? +row[0] : Date.parse(String(row[0] || ''));
    const diffMin = Math.abs(hookTs - (when || hookTs)) / 60000;

    // Parse the publish payload (column G usually)
    const raw = String(row[6] || '');
    let pub = {};
    try { pub = JSON.parse(raw || '{}'); } catch (_) { pub = {}; }

    const pay = (pub && (pub.payload || pub.PAYLOAD)) || {};
    const amtPub = Number(pay.AMOUNT ?? pub.AMOUNT ?? NaN);  // in cents
    const epiPub = String(pub.epi || pub.EPI || '');
    const reqId  = String(pay.REQ_TXN_ID || pub.REQ_TXN_ID || reqIdFromCol || '');
    // AFTER (covers both  cases)
    const invPub = String(pub.invoicenumber || pub.INVOICENUMBER || '');

    // Must match amount in cents
    if (!(Number.isFinite(amtPub) && amtPub === amountCents)) continue;

    // Score: amount match (mandatory) + EPI match + invoice match + recency
    let score = 100; // amount already matched
    if (epiFromHook && epiPub && epiPub === epiFromHook) score += 25;
    if (invoiceFromHook && invPub && invPub === invoiceFromHook) score += 25;
    // prefer closer in time (within ~10 min)
    if (diffMin <= 10) score += (10 - diffMin);  // 0..10

    if (score > best.score && reqId) best = {score, id: reqId};
  }

  return best.id;
}


/**
 * Heuristic correlation when webhook lacks REQ_TXN_ID:
 * - Match by EPI (if present) and amount (cents).
 * - Prefer the most recent publish within ±10 minutes of the webhook's timestamp.
 * - Scans the last ~50 publish rows from bottom up.
 * Returns the REQ_TXN_ID string or ''.
 */
function valorGuessReqIdFromPublish_(amountCents, epiFromHook, hookWhenStr, rawBody) {
  try {
    const ss = book_();
    const sh = ss.getSheetByName('Valor Publish Log');
    if (!sh) return '';

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return '';

    const maxBack = Math.max(2, lastRow - 60); // scan up to 60 recent rows
    const range = sh.getRange(maxBack, 1, lastRow - maxBack + 1, Math.min(12, sh.getLastColumn())); // first 12 cols is plenty
    const vals = range.getValues();

    // Parse time from webhook (UTC-ish strings in payload) with fallbacks
    let hookTs = Date.now();
    if (hookWhenStr && hookWhenStr.trim()) {
      const parsed = new Date(hookWhenStr.replace('T', ' '));
      if (!isNaN(+parsed)) hookTs = +parsed;
    }

    // Helpers to extract fields from a row's text blob
    const rxReq  = /"REQ_TXN_ID"\s*:\s*"([^"]+)"/i;
    const rxAmt  = /"AMOUNT"\s*:\s*"?(\d{1,10})"?/i;
    const rxEpi  = /"epi"\s*:\s*"(\d{6,})"/i;

    let bestId = '';
    let bestScore = -1;

    for (let i = vals.length - 1; i >= 0; i--) {
      const row = vals[i];
      const whenCell = row[0]; // column A assumed timestamp in your sheet
      const whenTs = (whenCell && whenCell.getTime) ? +whenCell : Date.now();

      // Concatenate row cells to search within (covers URL, request JSON, etc.)
      const rowText = row.map(v => (v == null ? '' : String(v))).join(' ');

      const mReq = rowText.match(rxReq);
      if (!mReq) continue; // must have a REQ_TXN_ID in the publish row
      const reqId = mReq[1];

      const mAmt = rowText.match(rxAmt);
      const amt  = mAmt ? parseInt(mAmt[1], 10) : NaN;
      if (!amt || amt !== amountCents) continue; // amount must match in cents

      // EPI: prefer matching; if not present in row, don't filter it out
      let epiOk = true;
      if (epiFromHook) {
        const mEpi = rowText.match(rxEpi);
        if (mEpi && mEpi[1] && String(mEpi[1]) !== String(epiFromHook)) epiOk = false;
      }
      if (!epiOk) continue;

      // Score by recency vs webhook time (closer is better). Accept within 10 minutes.
      const diffMs = Math.abs((hookTs || whenTs) - whenTs);
      const within = diffMs <= 10 * 60 * 1000;
      if (!within) continue;

      const score = 1 / (diffMs + 1);
      if (score > bestScore) {
        bestScore = score;
        bestId = reqId;
      }
    }

    return bestId || '';
  } catch (e) {
    return '';
  }
}

/** UI helper — minimize float errors by converting dollars→cents here. */
function apiStartValorCheckoutUI(totalDollars, invoice, items) {
  const cents = Math.round(Number(totalDollars) * 100);
  return apiStartValorCheckout(cents, invoice, items);
}

function valorWebhookHandler_(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    let obj = {}; try { obj = JSON.parse(raw || '{}'); } catch (_) {}

    // Normalize both shapes: flat & {"event":"transactions","data":{...}}
    let d = obj;
    if (obj && obj.event === 'transactions' && obj.data && typeof obj.data === 'object') d = obj.data;

    // Amounts (cents)
    const amtC  = Number(d.amount || 0) || 0;
    const totC  = Number(d.total_amount || d.amount || 0) || amtC;
    const epi   = String(d.epi_id || obj.epi || '').trim();
    const txnId = String(d.txn_id || obj.txn_id || '').trim();

    // Invoice can be in multiple places; handle lowercase and UPPERCASE
  const invRaw =
    (d && (d.invoicenumber ?? d.INVOICENUMBER)) ||
    (d && d.reference_descriptive_data && (d.reference_descriptive_data.invoicenumber ?? d.reference_descriptive_data.INVOICENUMBER)) ||
    (d && (d.invoice_no ?? obj.invoice_no)) ||
    '';

  const inv = String(invRaw || '').trim();

    // State/approval
    let state = String(d.STAT || d.status || d.STATE || '').toUpperCase();
    if (!state) {
      const rc = String(d.response_code || '').trim();
      const hr = String(d.host_response || '').trim().toUpperCase();
      if (rc === '00' || hr === 'APPROVAL') state = 'APPROVED';
    }

    // Compose a "when"
    let hookWhen = new Date();
    const ds = (d.response_date || d.request_date || '').trim();
    const ts = (d.response_time || d.request_time || '').trim();
    if (ds || ts) {
      const parsed = new Date((ds + ' ' + ts).trim().replace('T', ' '));
      if (!isNaN(+parsed)) hookWhen = parsed;
    }

    // Sheet setup: A..J
    const ss = book_();
    let sh = ss.getSheetByName('Valor Webhook Log');
    if (!sh) sh = ss.insertSheet('Valor Webhook Log');
    if (sh.getLastRow() < 1) {
      sh.appendRow(['Timestamp','ReqTxnId','State','Amount','TotalWithFees','Note','Raw (full)','txn_id','epi_id','invoicenumber']);
      sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm');
      sh.getRange('D:E').setNumberFormat('$#,##0.00');
    }

    const amount = +(amtC / 100).toFixed(2);
    const total  = +(totC / 100).toFixed(2);
    const note   = String(d.display_message || d.approval_code || obj.NOTE || obj.message || obj.ERROR_MSG || '');

    // Append immediately (B stays blank; we’ve moved off REQ_TXN_ID)
    sh.appendRow([ new Date(), '', state || '', amount, total, note, raw, txnId || '', epi || '', inv || '' ]);
    const row = sh.getLastRow();

    // Cache by INVOICENUMBER so the POS can complete on next poll
    if (inv) {
      const approved = (state === 'APPROVED' || state === 'AUTHCAPTURE' || state === '0');
      const out = {
        ok: true, found: true, approved,
        state, amount, total, fee: Math.max(0, +(total - amount).toFixed(2)),
        txn_id: txnId || '', epi_id: epi || '', invoicenumber: inv
      };
      const cache = CacheService.getScriptCache();
      cache.put('valor:inv:' + inv, JSON.stringify(out), 600);
    }

    // (Optional) legacy: try to fill ReqTxnId for audit if you want
    // (kept out here to avoid noise in your runtime path)

    return ContentService.createTextOutput(JSON.stringify({ ok:true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function apiValorCheckInvoice(invoicenumber) {
  const inv = String(invoicenumber || '').trim();
  if (!inv) return { ok:false, error:'Missing invoicenumber', found:false, approved:false, state:'', amount:0, total:0, fee:0 };

  const cache = CacheService.getScriptCache();
  const K = 'valor:inv:' + inv;

  // 1) Cache hit?
  const hit = cache.get(K);
  if (hit) return JSON.parse(hit);

  // 2) Scan recent webhook rows for a matching invoice (col J)
  try {
    const sh = sheet_('Valor Webhook Log');
    if (sh) {
      const last = sh.getLastRow();
      if (last > 1) {
        const take = Math.min(100, last - 1);
        const vals = sh.getRange(last - take + 1, 1, take, Math.min(10, sh.getLastColumn())).getValues();
        for (let i = vals.length - 1; i >= 0; i--) {
          const row = vals[i];
          if (String(row[9] || '') === inv) { // col J = invoicenumber
            const amount = Number(row[3] || 0);
            const total  = Number(row[4] || amount || 0);
            const state  = String(row[2] || '').toUpperCase();
            const approved = (state === 'APPROVED' || state === 'AUTHCAPTURE' || state === '0');
            const out = {
              ok: true, found: true, approved,
              state, amount, total, fee: Math.max(0, +(total - amount).toFixed(2)),
              invoicenumber: inv,
              txn_id: String(row[7] || ''), // H
              epi_id: String(row[8] || '')  // I
            };
            cache.put(K, JSON.stringify(out), 600);
            return out;
          }
        }
      }
    }
  } catch (_) {}

  // 3) Not found yet → tell UI to keep waiting (webhook is the source of truth)
  return { ok:true, found:false, approved:false, state:'', amount:0, total:0, fee:0, waiting:true };
}

// Server-side Status API call that mirrors Valor's docs.
function valorFetchStatus_(reqTxnId) {
  const p = valorProps_();
  const url = valorGetStatusUrl_(); // ...:4430/?txn_status=
  const body = {
    appid: String(p.VALOR_APP_ID),
    appkey: String(p.VALOR_APP_KEY),
    epi: String(p.VALOR_EPI),
    txn_type: 'vc_status',
    req_txn_id: String(reqTxnId || ''),

    // Some environments require these on vc_status:
    channel_id: String(p.VALOR_CHANNEL_ID || ''),
    version: '1'
  };

  let http = 0, text = '', json = null;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      muteHttpExceptions: true,
      contentType: 'application/json',
      headers: { accept: 'application/json' },
      payload: JSON.stringify(body)
    });
    http = res.getResponseCode();
    text = res.getContentText() || '';
    try { json = JSON.parse(text); } catch (_) { json = { raw: text }; }
  } catch (e) {
    return {
      ok: false, http: 0, found: false, approved: false, state: '',
      amount: 0, total: 0, fee: 0,
      ack: { error: String(e && e.message || e) }
    };
  }

  // Normalize ack
  const a = (json && json.ack) ? json.ack : json;
  const sRaw = String((a && (a.STATE || a.status || a.STATUS || a.STAT)) || '').toUpperCase();

  const amtC = Number((a && (a.AMOUNT || a.amount)) || 0) || 0;
  const totC = Number((a && (a.TOTAL_AMOUNT || a.total_amount)) || amtC) || amtC;

  const approved = (sRaw === 'APPROVED' || sRaw === 'AUTHCAPTURE' || sRaw === '0');
  const declined = /DECLIN/i.test(sRaw) || sRaw === '2';
  const found = approved || declined;

  return {
    ok: http >= 200 && http < 300,
    http,
    found,
    approved,
    state: sRaw,
    amount: +(amtC / 100).toFixed(2),
    total: +(totC / 100).toFixed(2),
    fee: Math.max(0, +((totC - amtC) / 100).toFixed(2)),
    ack: json
  };
}

/** Public: allow POS to call status directly if needed. */
function apiValorStatus(reqTxnId) {
  if (!reqTxnId) return { ok:false, error:'Missing reqTxnId' };
  try {
    return valorFetchStatus_(reqTxnId);
  } catch(e) {
    return { ok:false, error:e.message };
  }
}

function apiValorCheck(id) {
  const reqId = String(id || '').trim();
  const cache = CacheService.getScriptCache();
  const K     = 'valor:' + reqId;

  // 1) Cache first (webhook handler primes this)
  const hit = cache.get(K);
  if (hit) return JSON.parse(hit);

  // 2) Scan the last ~200 webhook rows for a direct match in column B
  try {
    const sh = sheet_('Valor Webhook Log');
    if (sh) {
      const last = sh.getLastRow();
      if (last > 1) {
        const take = Math.min(200, last - 1);
        const vals = sh.getRange(last - take + 1, 1, take, Math.min(9, sh.getLastColumn())).getValues();
        for (let i = vals.length - 1; i >= 0; i--) {
          const row = vals[i];
          if (String(row[1] || '') === reqId) {
            const amount = Number(row[3] || 0);
            const total  = Number(row[4] || amount || 0);
            const state  = String(row[2] || '').toUpperCase();
            const approved = (state === 'APPROVED' || state === 'AUTHCAPTURE' || state === '0');
            const out = {
              ok: true, found: true, approved,
              state, amount, total, fee: Math.max(0, +(total - amount).toFixed(2))
            };
            cache.put(K, JSON.stringify(out), 600);
            return out;
          }
        }
      }
    }
  } catch (_) {}

  // 3) Throttled Status fallback — avoid spamming VC09
  const throttleKey = 'valor:throttle:' + reqId;
  if (cache.get(throttleKey)) {
    // Still waiting; tell the UI to keep polling (webhook usually lands within ~60s)
    return { ok: true, found: false, approved: false, state: '', amount: 0, total: 0, fee: 0, waiting: true };
  }
  cache.put(throttleKey, '1', 6); // at most one status call every 6s per reqId

  const st = valorFetchStatus_(reqId); // your existing vc_status caller
  if (st && st.ok && st.found) {
    cache.put(K, JSON.stringify(st), 600);
  }
  return st;
}

/** Tiny diagnostic so you can sanity-check env/URLs from the browser console. */
function apiValorDiag() {
  const p = valorProps_();
  return {
    env: p.VALOR_ENV,
    baseUrl: p.VALOR_BASE_URL,
    publishUrl: p.VALOR_PUBLISH_URL || '(using VALOR_BASE_URL)',
    statusUrl: (function(){ try { return valorGetStatusUrl_(); } catch(e){ return '(derive failed)'; } })(),
    epiTail: p.VALOR_EPI ? p.VALOR_EPI.slice(-4) : '',
    haveIds: !!(p.VALOR_CHANNEL_ID && p.VALOR_APP_ID && p.VALOR_APP_KEY)
  };
}


function apiValorWebhookRaw(reqTxnId) {
  try {
    const id = String(reqTxnId || '').trim();
    const ss = book_();
    const sh = ss.getSheetByName('Valor Webhook Log');
    if (!sh) return { ok:false, error:'No "Valor Webhook Log" sheet' };
    const vals = sh.getDataRange().getValues();
    for (let r = vals.length - 1; r >= 1; r--) {
      if (String(vals[r][1]) === id) {
        const when  = vals[r][0];
        const state = vals[r][2];
        const amount= vals[r][3];
        const total = vals[r][4];
        const rawTrunc = vals[r][6] || '';
        // Column H (index 7) holds full raw if present; fall back to truncated
        const rawFull = (sh.getRange(r+1, 8).getValue() || rawTrunc || '');
        return { ok:true, row:r+1, when, state, amount, total, rawFull };
      }
    }
    return { ok:false, error:'Not found' };
  } catch (e) {
    return { ok:false, error:String(e && e.message || e) };
  }
}

/** ============================== 
 * Valor Sessions (invoice/attempt tracker)
 * ============================== */
function valorSessionsSheet_() {
  var ss = book_();
  var name = 'Valor Sessions';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var headers = ['invoice','attempt','reqTxnId','amountCents','status','startedAt','lastSeenAt','saleId','webhookJson'];
  if (String(sh.getRange(1,1).getValue()||'') !== 'invoice') {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}
function valorSessionKey_(invoice, attempt) {
  return String(invoice||'').trim() + '|' + String(attempt||'').trim();
}
function valorSessionsIndex_() {
  var sh = valorSessionsSheet_();
  var last = sh.getLastRow();
  var map = new Map();
  if (last >= 2) {
    var vals = sh.getRange(2,1,last-1,9).getValues();
    for (var i=0;i<vals.length;i++){
      var row = vals[i];
      var k = valorSessionKey_(row[0], row[1]);
      map.set(k, { rowIndex: i+2, row: row });
    }
  }
  return { sh: sh, map: map };
}
function upsertValorSession_(o) {
  var idx = valorSessionsIndex_();
  var k = valorSessionKey_(String(o.invoice||'').slice(0,24), String(o.attempt||'A1').trim());
  var now = new Date();

  var prev = null;
  if (idx.map.has(k)) {
    var hit = idx.map.get(k).row;
    prev = {
      invoice: hit[0], attempt: hit[1], reqTxnId: hit[2], amountCents: hit[3],
      status: hit[4], startedAt: hit[5], lastSeenAt: hit[6], saleId: hit[7],
      webhookJson: hit[8] ? JSON.parse(hit[8]) : null
    };
  }

  var merged = {
    invoice:  String(o.invoice|| (prev && prev.invoice) || '').slice(0,24),
    attempt:  String(o.attempt|| (prev && prev.attempt) || 'A1').trim(),
    reqTxnId: String((o.reqTxnId != null ? o.reqTxnId : (prev && prev.reqTxnId)) || ''),
    amountCents: Number((o.amountCents != null ? o.amountCents : (prev && prev.amountCents)) || 0),
    status: (o.status != null ? o.status : (prev && prev.status) || 'pending'),
    startedAt: (prev && prev.startedAt) || o.startedAt || now,
    lastSeenAt: now,
    saleId: (o.saleId ? String(o.saleId) : (prev && prev.saleId) || ''),
    webhookJson: (function() {
      var base = (prev && prev.webhookJson) || {};
      if (o.posPayload) base.posPayload = o.posPayload;     // STASH CART SNAPSHOT
      if (o.webhookJson) base.lastWebhook = o.webhookJson;
      return Object.keys(base).length ? base : null;
    })()
  };

  var rowVals = [
    merged.invoice, merged.attempt, merged.reqTxnId, merged.amountCents,
    merged.status, merged.startedAt, merged.lastSeenAt, merged.saleId,
    merged.webhookJson ? JSON.stringify(merged.webhookJson) : ''
  ];

  if (idx.map.has(k)) {
    var rowIndex = idx.map.get(k).rowIndex;
    idx.sh.getRange(rowIndex, 1, 1, 9).setValues([rowVals]);
  } else {
    idx.sh.appendRow(rowVals);
  }
  return merged;
}
function getValorSession_(invoice, attempt) {
  var idx = valorSessionsIndex_();
  var k = valorSessionKey_(String(invoice||'').slice(0,24), String(attempt||'A1').trim());
  if (!idx.map.has(k)) return null;
  var r = idx.map.get(k).row;
  return {
    invoice: r[0], attempt: r[1], reqTxnId: r[2], amountCents: r[3],
    status: r[4], startedAt: r[5], lastSeenAt: r[6], saleId: r[7],
    webhookJson: r[8] ? JSON.parse(r[8]) : null
  };
}

/** 
 * Public API: finalize NOW based on clerk seeing "Approved" on terminal.
 * Idempotent: if saleId already exists for (invoice,attempt), returns that.
 */
function apiPos_FinalizeNow(rawJson) {
  var p = {}; try { p = JSON.parse(String(rawJson||'')||'{}'); } catch(_){}
  var invoice = String(p.invoice||'').slice(0,24);
  var attempt = String(p.attempt||'A1').trim();

  var sess = getValorSession_(invoice, attempt);
  if (sess && sess.saleId) {
    return { ok:true, saleId: sess.saleId, alreadyFinalized:true };
  }

  var salePayload = p.posPayload || {};
  var status = (sess && String(sess.status).toLowerCase() === 'approved') ? 'approved' : 'pending';
  salePayload.valor = {
    invoice: invoice, attempt: attempt,
    status: status, finalizedBy: 'manual',
    brand: p.brand || '', last4: p.last4 || '', auth: p.auth || ''
  };

  var result = saveSale(salePayload); // SINGLE SOURCE OF TRUTH

  try {
    apiVendooManual_AddFromSale_({
      sale_id: result.saleId,
      sale_ts: new Date(),
      items: Array.isArray(salePayload.items) ? salePayload.items : []
    });
  } catch (e) {
    Logger.log('[apiPos_FinalizeNow] manual delist add failed: ' + (e && e.message ? e.message : e));
  }

  upsertValorSession_({
    invoice: invoice, attempt: attempt,
    reqTxnId: p.reqTxnId||'', amountCents: p.amountCents||0,
    status: status, startedAt: p.startedAt||new Date(),
    saleId: result.saleId, posPayload: p.posPayload
  });


  
  return {
    ok: true, saleId: result.saleId, pending: true,
    subtotal: result.subtotal, tax: result.tax, total: result.total
  };
}


// === Split-leg helper: finalize a single terminal leg WITHOUT saving the POS sale ===
function apiPos_FinalizeLeg(rawJson) {
  var p = {}; try { p = JSON.parse(String(rawJson||'')||'{}'); } catch(_){}
  var invoice = String(p.invoice||'').slice(0,24);
  var attempt = String(p.attempt||'A1').trim();

  if (!invoice) throw new Error('Missing invoice');

  var sess = getValorSession_(invoice, attempt);
  if (!sess) {
    // Defensive: poll once, in case the webhook wrote very recently
    try {
      // pollOnce will NO-OP if you don't have one; keep this light if not available
      // (ok to omit if your codebase doesn't include a single-poll util)
      // pollValorOnce_(invoice, attempt);
      sess = getValorSession_(invoice, attempt);
    } catch(_){}
  }

  var status = (sess && String(sess.status||'').toLowerCase()) || '';
  var approved = (status === 'approved' || status === 'ok');
  if (!approved) {
    return { ok:false, declined:true, status: status || 'pending' };
  }

  // Return useful leg details to the client for audit
  return {
    ok: true,
    status: 'approved',
    brand:  String(sess.brand || ''),
    last4:  String(sess.last4 || ''),
    auth:   String(sess.auth || ''),
    txnId:  String(sess.txnId || sess.reqTxnId || ''),
    epiId:  String(sess.epiId || ''),
    amountCents: Number(sess.amountCents||0)
  };
}


/**
 * Webhook handler (called when Cloudflare Worker forwards with ?source=valor).
 * Reconciles manual or auto finalization. Does not create duplicate rows.
 */
function valorWebhookHandlerV2_(rawBody) {
  try {
    var body = {}; try { body = JSON.parse(rawBody || '{}'); } catch (_){}

    var invoice = '';
    try {
      invoice = String(
        (body.invoicenumber || body.INVOICENUMBER || body.invoice || body.invoice_number) ||
        (body.data && (body.data.invoicenumber || body.data.INVOICENUMBER || body.data.invoice || body.data.invoice_number)) ||
        (body.data && body.data.reference_descriptive_data &&
          (body.data.reference_descriptive_data.invoicenumber || body.data.reference_descriptive_data.INVOICENUMBER)) ||
        ''
      ).trim();
    } catch(_){}
    invoice = String(invoice||'').slice(0,24);
    var attempt = String(body.attempt || 'A1').trim();

    var d = (body && body.data && typeof body.data === 'object') ? body.data : body;
    var amtC  = Number(d.amount || d.AMOUNT || 0) || 0;
    var totC  = Number(d.total_amount || d.TOTAL_AMOUNT || d.amount || d.AMOUNT || amtC) || 0;
    var txnId = String(d.txn_id || body.txn_id || '').trim();
    var epi   = String(d.epi_id || body.epi || '').trim();
    var state = String(d.STAT || d.status || d.STATE || '').toUpperCase();
    if (!state) {
      var rc = String(d.response_code || '').trim();
      var hr = String(d.host_response || '').trim().toUpperCase();
      if (rc === '00' || hr === 'APPROVAL') state = 'APPROVED';
    }
    var approved = (state === 'APPROVED' || state === 'AUTHCAPTURE' || state === '0');
    var declined = (state === 'DECLINED' || state === '2');

    if (!invoice) {
      return ContentService.createTextOutput(JSON.stringify({ ok:true, ignored:true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // minimal webhook log (kept same style)
    var ss = book_();
    var sh = ss.getSheetByName('Valor Webhook Log') || ss.insertSheet('Valor Webhook Log');
    if (sh.getLastRow() < 1) {
      sh.appendRow(['Timestamp','ReqTxnId','State','Amount','TotalWithFees','Note','Raw (full)','txn_id','epi_id','invoicenumber']);
      sh.getRange('A:A').setNumberFormat('m/d/yyyy h:mm');
      sh.getRange('D:E').setNumberFormat('$#,##0.00');
    }
    sh.appendRow([new Date(), txnId, (approved ? 'APPROVED' : (declined ? 'DECLINED' : String(state||'').toUpperCase())),
                  +(amtC/100).toFixed(2), +(totC/100).toFixed(2), (d.display_message||''),
                  JSON.stringify(body), txnId, epi, invoice]);

    var sess = getValorSession_(invoice, attempt);
    var existingSaleId = (sess && sess.saleId) ? String(sess.saleId) : '';

    // Always keep session fresh
    upsertValorSession_({
      invoice: invoice, attempt: attempt,
      reqTxnId: '', amountCents: amtC,
      status: approved ? 'approved' : (declined ? 'declined' : 'pending'),
      webhookJson: body,
      saleId: existingSaleId
    });

    // If manual/normal already wrote the sale, PATCH ONLY
    if (existingSaleId) {
      patchSaleValorEnvelope_(existingSaleId, {
        status: approved ? 'approved' : (declined ? 'declined' : 'pending'),
        pending: false,
        amount: +(amtC/100).toFixed(2),
        total:  +(totC/100).toFixed(2),
        fee: Math.max(0, +((totC - amtC)/100).toFixed(2)),
        txn_id: txnId || '', epi_id: epi || '',
        finalizedBy: 'manual'
      });
      return ContentService.createTextOutput(JSON.stringify({ ok:true, reconciled:true, saleId: existingSaleId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // No existing row yet — DO NOT auto-finalize anymore. Only record status and wait for clerk to finalize.
    if (approved) {
      upsertValorSession_({
        invoice: invoice,
        attempt: attempt,
        status: 'approved',
        // keep the latest webhook JSON merged elsewhere in this handler
      });
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true, recorded: true, approved: true, waitingForManualFinalize: true })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok:true, recorded:true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('valorWebhookHandlerV2_ error: ' + err);
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/** Patch a row’s Items JSON (envelope) to update valor.* fields (by saleId). */
function patchSaleValorEnvelope_(saleId, patch) {
  var sh = sheet_('Sales Log');
  var last = sh.getLastRow(); if (last < 2) return false;
  var rng = sh.getRange(2,1,last-1,12).getValues(); // A..L
  for (var i=0;i<rng.length;i++){
    var row = rng[i];
    if (String(row[1]||'') === String(saleId)) {
      var env = {};
      try { env = JSON.parse(row[10] || '{}'); } catch(_){}
      env.valor = env.valor || {};
      Object.keys(patch||{}).forEach(function(k){ env.valor[k] = patch[k]; });
      sh.getRange(2+i, 11).setValue(JSON.stringify(env)); // Items JSON (col K)
      return true;
    }
  }
  return false;
}

/**
 * Reconcile Sales Log rows still marked valor.pending:true by matching their
 * valor.invoice against "Valor Webhook Log". Patches envelope and logs results.
 * Run manually from the Apps Script editor.
 */
function reconcileValorPending() {
  var logs = [];
  var sh = sheet_('Sales Log');
  if (!sh) throw new Error('Missing "Sales Log" sheet.');
  var last = sh.getLastRow();
  if (last < 2) return { ok:true, fixed:0, unresolved:0, note:'No rows' };

  // Read A..L to get SaleId + Envelope (K)
  var vals = sh.getRange(2,1,last-1,12).getValues();
  var fixed = 0, unresolved = 0;

  // Normalize invoices to 24 chars so webhook vs. sales keys match
  var normInv = function(s){ return String(s||'').trim().slice(0,24); };

  // Build index of recent webhook rows by invoice (col J)
  var wh = sheet_('Valor Webhook Log');
  var hookIdx = new Map();
  if (wh && wh.getLastRow() > 1) {
    var take = Math.min(2000, wh.getLastRow() - 1);
    var wv = wh.getRange(wh.getLastRow() - take + 1, 1, take, Math.min(10, wh.getLastColumn())).getValues();
    for (var i = 0; i < wv.length; i++) {
      var row = wv[i];
      var inv = normInv(row[9]);  // col J = invoicenumber
      if (!inv) continue;
      var state = String(row[2]||'').toUpperCase(); // APPROVED/DECLINED/0/2
      var amount = Number(row[3]||0);
      var total  = Number(row[4]||amount||0);
      var approved = (state === 'APPROVED' || state === 'AUTHCAPTURE' || state === '0');
      hookIdx.set(inv, {
        state: state, approved: approved,
        amount: amount, total: total,
        fee: Math.max(0, +(total - amount).toFixed(2)),
        txn_id: String(row[7]||''), // H
        epi_id: String(row[8]||'')  // I
      });
    }
  }

  for (var r = 0; r < vals.length; r++) {
    var saleId = String(vals[r][1]||'').trim();
    var envRaw = vals[r][10];
    var env = {};
    try { env = envRaw ? JSON.parse(envRaw) : {}; } catch (_) { env = {}; }

    var v = env.valor || null;
    if (!v || !v.invoice) continue; // no valor info
    if (v.pending !== true && !/pending/i.test(String(v.status||''))) continue; // not pending

    var inv = normInv(v.invoice);
    var hit = hookIdx.get(inv);

    if (hit) {
      env.valor = env.valor || {};
      env.valor.pending = false;
      env.valor.status  = hit.approved ? 'approved' : 'declined';
      env.valor.amount  = hit.amount;
      env.valor.total   = hit.total;
      env.valor.fee     = hit.fee;
      env.valor.txnId   = hit.txn_id || env.valor.txnId || '';
      env.valor.epiId   = hit.epi_id || env.valor.epiId || '';

      sh.getRange(2 + r, 11).setValue(JSON.stringify(env)); // col K
      fixed++;
      logs.push('Fixed ' + saleId + ' via invoice ' + inv + ' => ' + env.valor.status +
                ' amt=$' + hit.amount.toFixed(2));
    } else {
      unresolved++;
      logs.push('Unresolved ' + saleId + ' (invoice ' + inv + '): no webhook match.');
    }
  }

  Logger.log('[reconcileValorPending] fixed=%s unresolved=%s', fixed, unresolved);
  logs.slice(0,50).forEach(function(x){ Logger.log(' - %s', x); });
  return { ok:true, fixed:fixed, unresolved:unresolved, notes:logs.slice(0,200) };
}

/***** end Valor code *****/