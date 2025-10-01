function apiSales_ListVendooDelist(dayStrOpt) {
  const tz = CASH_TZ || Session.getScriptTimeZone() || 'America/Denver';
  const day = dayStrOpt ? new Date(dayStrOpt + 'T12:00:00') : new Date();
  const start = new Date(Utilities.formatDate(day, tz, 'yyyy-MM-dd') + 'T00:00:00');
  const end   = new Date(Utilities.formatDate(day, tz, 'yyyy-MM-dd') + 'T23:59:59');

  const sh = sheet_('Sales Log');
  const last = sh.getLastRow();
  if (last < 2) return { ok:true, items: [] };

  const rows = sh.getRange(2, 1, last-1, Math.max(12, sh.getLastColumn())).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= start && r[0] <= end);

  // envelope in col K (index 10) may contain normalized items
  const soldMap = new Map(); // sku -> qty
  rows.forEach(r => {
    const envRaw = r[10];
    let env = null;
    try { env = envRaw ? JSON.parse(envRaw) : null; } catch(_) {}
    const items = (env && Array.isArray(env.items)) ? env.items : (Array.isArray(env) ? env : []);
    items.forEach(it => {
      const sku = String(it.sku || '').trim();
      const q = Math.max(1, Number(it.qty)||1);
      if (sku) soldMap.set(sku, (soldMap.get(sku)||0) + q);
    });
  });

  if (!soldMap.size) return { ok:true, items: [] };

  // Join to SKU Tracker to get name and Online Location
  const inv = sheet_('SKU Tracker');
  const invLast = inv.getLastRow();
  const invVals = inv.getRange(2, 1, Math.max(0, invLast-1), 9).getValues();
  const skuToRow = new Map();
  invVals.forEach((r, i) => { const s=String(r[0]||'').trim(); if (s) skuToRow.set(s, i); });

  const out = [];
  soldMap.forEach((qty, sku) => {
    const idx = skuToRow.get(sku);
    const row = idx != null ? invVals[idx] : null;
    const name = row ? String(row[2]||'') : '';
    const onlineLoc = row ? String(row[7]||'') : '';
    if (onlineLoc && /^(vendoo|both)$/i.test(onlineLoc)) {
      out.push({ sku, name, qty, onlineLoc });
    }
  });

  // Sort for convenience
  out.sort((a,b) => a.sku.localeCompare(b.sku));
  return { ok:true, items: out, day: Utilities.formatDate(day, tz, 'yyyy-MM-dd') };
}


/**
 * Vendoo helpers (placeholder for future bulk upload flows)
 * - Compose the Vendoo SKU from SKU + store + bin/shelf
 * - Build a consistent payload object (mirrors the client side)
 */
function vendooComposeSku_(sku, storeLoc, caseBinShelf) {
  const parts = [String(sku || '').trim(), String(storeLoc || '').trim(), String(caseBinShelf || '').trim()]
    .filter(Boolean);
  return parts.join(' | ');
}

/**
 * Build a Vendoo-like payload from an Intake row object.
 * Extend later for bulk runs.
 */
function vendooBuildPayload_(row) {
  return {
    title: row.name || '',
    description: row.longDesc || row.name || '',
    brand: row.brand || '',
    condition: row.condition || '',
    price: Number(row.price || 0),
    quantity: Number(row.qty || 1),
    sku: vendooComposeSku_(row.sku, row.storeLoc, row.caseBinShelf),
    zip: '80033',
    category: String(row.vendooCategoryPath || '').trim() ||
          'Toys & Hobbies > Action Figures & Accessories > Action Figures'
  };
}
