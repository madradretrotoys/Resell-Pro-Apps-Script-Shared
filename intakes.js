function getCategoryNames()   { return getCategoryNames_(); }
function getStoreLocations()  { return getStoreLocations_(); }

function getOnlineLocations() { return getOnlineLocations_(); }
function getConditions() { return getConditions_(); }
function getBrands()     { return getBrands_(); }

function getprimaryColors()     { return getprimaryColors_(); }

function getVendooDisplayList() { return getVendooDisplayList_(); }
function getVendooPathMap()     { return getVendooPathMap_(); }

// Shipping options + metadata (from "Shipping" sheet)
function getShippingBoxOptions() { return getShippingBoxOptions_(); }
function getShippingBoxMeta()    { return getShippingBoxMeta_(); }

function getVendooDisplayList_() {
  const ss = book_();
  const sh = ss.getSheetByName('Vendoo Categories'); // tab name in your workbook
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last <= 1) return [];
  const header = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
  const colIdx = header.findIndex(h => h.trim().toLowerCase() === 'resell pro display');
  if (colIdx < 0) return [];
  const vals = sh.getRange(2, colIdx+1, last-1, 1).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  // De-dupe while preserving order
  const seen = new Set(); const out = [];
  for (const v of vals) if (!seen.has(v)) { seen.add(v); out.push(v); }
  return out;
}

function getVendooPathMap_() {
  const ss = book_();
  const sh = ss.getSheetByName('Vendoo Categories');
  const out = {};
  if (!sh) return out;

  const last = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (last <= 1) return out;

  const header = sh.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||'').trim());
  const nameToIndex = (name) => header.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idxDisplay = nameToIndex('Resell Pro Display');
  const idxParent  = nameToIndex('Vendoo Parent');
  // Leaves may be fewer than 7 in some rows; find all present
  const leafIdxes = [];
  for (let i=1; i<=7; i++) {
    const k = nameToIndex(`Vendoo Leaf ${i}`);
    if (k >= 0) leafIdxes.push(k);
  }

  if (idxDisplay < 0 || idxParent < 0) return out;

  const rows = sh.getRange(2,1,last-1,lastCol).getValues();
  for (const row of rows) {
    const display = String(row[idxDisplay] || '').trim();
    if (!display) continue;
    const parts = [ String(row[idxParent]||'').trim() ];
    for (const li of leafIdxes) {
      const v = String(row[li]||'').trim();
      if (v) parts.push(v);
    }
    const path = parts.filter(Boolean).join(' > ');
    if (path) out[display] = path;
  }
  return out;
}


/* ====== LOOKUP HELPERS (dropdown sources) ====== */

function getConditions_() {
  try {
    const sh = sheet_('Conditions');
    const n  = Math.max(0, sh.getLastRow() - 1);
    if (n === 0) return [];
    return sh.getRange(2, 1, n, 1).getValues().flat().filter(Boolean);
  } catch (err) {
    const msg = `getConditions_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

function getShippingBoxOptions_() {
  try {
    const sh = sheet_('Shipping');                // tab name: Shipping
    const last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last <= 1 || lastCol < 1) return [];
    const head = sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
    const idxBox = head.findIndex(h => h.toLowerCase() === 'shipping_box_options');
    if (idxBox < 0) return [];
    const rows = sh.getRange(2,1,last-1,lastCol).getValues();
    return rows.map(r => String(r[idxBox] || '').trim()).filter(Boolean);
  } catch (err) {
    const msg = `getShippingBoxOptions_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

function getShippingBoxMeta_() {
  try {
    const sh = sheet_('Shipping');                // tab name: Shipping
    const last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last <= 1 || lastCol < 1) return {};
    const head = sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
    function idx(name){ return head.findIndex(h => h.toLowerCase() === name.toLowerCase()); }
    const iBox = idx('Shipping_Box_Options');
    const iLb  = idx('Weight_LB');
    const iOz  = idx('Weight_OZ');
    const iLen = idx('Length');
    const iWid = idx('Width');
    const iHei = idx('Height');
    if (iBox < 0) return {};

    const rows = sh.getRange(2,1,last-1,lastCol).getValues();
    const map = {};
    for (const r of rows) {
      const key = String(r[iBox] || '').trim();
      if (!key) continue;
      const v = (x) => {
        const n = parseFloat(String(x).replace(/[^\d.]/g,'')); 
        return Number.isFinite(n) ? n : '';
      };
      map[key] = {
        lb:  iLb  >= 0 ? v(r[iLb])  : '',
        oz:  iOz  >= 0 ? v(r[iOz])  : '',
        len: iLen >= 0 ? v(r[iLen]) : '',
        wid: iWid >= 0 ? v(r[iWid]) : '',
        hei: iHei >= 0 ? v(r[iHei]) : ''
      };
    }
    return map;
  } catch (err) {
    const msg = `getShippingBoxMeta_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

function getBrands_() {
  try {
    const sh = sheet_('Brands');
    const n  = Math.max(0, sh.getLastRow() - 1);
    if (n === 0) return [];
    return sh.getRange(2, 1, n, 1).getValues().flat().filter(Boolean);
  } catch (err) {
    const msg = `getBrands_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

function getprimaryColors_() {
  try {
    const sh = sheet_('Vendoo Primary Colors');
    const n  = Math.max(0, sh.getLastRow() - 1);
    if (n === 0) return [];
    return sh.getRange(2, 1, n, 1).getValues().flat().filter(Boolean);
  } catch (err) {
    const msg = `getprimaryColors_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

function getCategoryNames_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CODE_SHEET);
    if (!sheet) throw new Error(`Sheet '${CODE_SHEET}' not found`);
    const last = sheet.getLastRow();
    if (last <= 1) return [];
    const values = sheet.getRange(2, 1, last - 1, 1).getValues();
    return values.flat().filter(String);
  } catch (err) {
    return { error: err.toString() };
  }
}
function getStoreLocations_() {
  try {
    const sh = sheet_('Store Locations');
    const n = Math.max(0, sh.getLastRow() - 1);
    if (n === 0) return [];
    return sh.getRange(2, 1, n, 1).getValues().flat().filter(Boolean);
  } catch (err) {
    const msg = `getStoreLocations_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}
function getOnlineLocations_() {
  try {
    const sh = sheet_('Online Locations');
    const n = Math.max(0, sh.getLastRow() - 1);
    if (n === 0) return [];
    return sh.getRange(2, 1, n, 1).getValues().flat().filter(Boolean);
  } catch (err) {
    const msg = `getOnlineLocations_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

// Ensure a header exists on the first row and return its 1-based column index.
// If missing, append a new column at the end with that header name.
function ensureHeaderColumn_(sheet, headerName) {
  const lastCol = sheet.getLastColumn();
  const header = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  let idx = header.findIndex(h => String(h).trim() === String(headerName).trim());
  if (idx >= 0) return idx + 1;
  // append a new header column
  sheet.insertColumnAfter(Math.max(1, lastCol) || 1);
  const newCol = (lastCol || 1) + 1;
  sheet.getRange(1, newCol, 1, 1).setValue(headerName);
  return newCol;
}

// Read-only: given an array of Buy Ticket ItemIDs, return a map: { "<ItemID>": { sku, duplicate } }
function apiSku_FindByBuyTicketItemIds(ids) {
  try {
    const out = {};
    if (!Array.isArray(ids) || !ids.length) return out;

    const sh = sheet_(SHEET_NAME);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return out;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const colSku = header.indexOf('SKU') + 1;
    const colLink = header.indexOf('Buy_Ticket_ITEM_ID') + 1;

    if (colSku <= 0 || colLink <= 0) return out;

    const linkVals = sh.getRange(2, colLink, lastRow - 1, 1).getValues().flat().map(v => String(v || '').trim());
    const skuVals  = sh.getRange(2, colSku,  lastRow - 1, 1).getValues().flat().map(v => String(v || '').trim());

    const wanted = new Set(ids.map(x => String(x || '').trim()).filter(Boolean));

    // Build map: itemId -> all SKUs found
    const hits = new Map(); // id -> array of {sku}
    for (let i = 0; i < linkVals.length; i++) {
      const id = linkVals[i];
      if (!id || !wanted.has(id)) continue;
      const sku = skuVals[i] || '';
      if (!sku) continue;
      if (!hits.has(id)) hits.set(id, []);
      hits.get(id).push({ sku });
    }

    hits.forEach((list, id) => {
      const uniqueSkus = Array.from(new Set(list.map(x => x.sku)));
      out[id] = { sku: uniqueSkus[0] || '', duplicate: uniqueSkus.length > 1 };
    });

    return out;
  } catch (err) {
    Logger.log('apiSku_FindByBuyTicketItemIds error: ' + err.message);
    return {};
  }
}


/* ====== DIAGNOSTIC ====== */
function debugDiag_() {
  try {
    const ss = book_();
    const cats = getCategoryNames_();
    const stores = getStoreLocations_();
    const online = getOnlineLocations_();
    const skuSh = sheet_(SHEET_NAME);
    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      categoryTab: CODE_SHEET,
      categoryCount: Array.isArray(cats) ? cats.length : -1,
      storeTab: 'Store Locations',
      storeCount: stores.length,
      onlineTab: 'Online Locations',
      onlineCount: online.length,
      skuTrackerRows: skuSh.getLastRow()
    };
  } catch (err) {
    const msg = `debugDiag_ error: ${err.message}`;
    Logger.log(msg);
    return { error: msg };
  }
}

/* ====== CATEGORY CODE MAP & SKU ====== */
function getCategoryCodeMap_() {
  const sh = sheet_(CODE_SHEET);
  const n = Math.max(0, sh.getLastRow() - 1);
  const map = {};
  if (n === 0) return map;
  const values = sh.getRange(2, 1, n, 2).getValues(); // [name, code]
  values.forEach(([name, code]) => {
    if (name && code) map[String(name).trim().toLowerCase()] = String(code).trim().toUpperCase();
  });
  return map;
}
function generateNextSku_(sheet, category, catMapOpt) {
  try {
    const catMap = catMapOpt || getCategoryCodeMap_();
    const code = catMap[String(category).trim().toLowerCase()];
    if (!code) return '';

    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return `${code}-` + String(1).padStart(SKU_PAD, '0');

    const existing = sheet.getRange(HEADER_ROW + 1, SKU_COL, lastRow - HEADER_ROW, 1).getValues().flat();
    let maxNum = 0, prefix = code + '-';
    for (const v of existing) {
      const s = String(v || '');
      if (s.startsWith(prefix)) {
        const n = parseInt(s.slice(prefix.length), 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
    return `${code}-${String(maxNum + 1).padStart(SKU_PAD, '0')}`;
  } catch (err) {
    const msg = `generateNextSku_ error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

/* ====== SHEET UX (grid editing helpers) ====== */
function setCategoryDropdown_() {
  try {
    const sku = sheet_(SHEET_NAME);
    const cats = sheet_(CODE_SHEET);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(cats.getRange('A2:A1000'), true)
      .setAllowInvalid(false)
      .setHelpText('Choose a category from Category Codes.')
      .build();
    sku.getRange('B2:B5000').setDataValidation(rule);
  } catch (err) {
    Logger.log(`setCategoryDropdown_ error: ${err.message}`);
  }
}
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    const row = e.range.getRow(), col = e.range.getColumn();
    if (row <= HEADER_ROW) return;

    if (col === CATEGORY_COL) {
      const category = sh.getRange(row, CATEGORY_COL).getValue();
      const skuCell = sh.getRange(row, SKU_COL);
      const current = String(skuCell.getValue() || '').trim();
      const codeMap = getCategoryCodeMap_();
      const code = codeMap[String(category).trim().toLowerCase()];
      if (!code) return;
      const currentPrefix = current.includes('-') ? current.split('-')[0] : '';
      if (!current || currentPrefix !== code) {
        const next = generateNextSku_(sh, category, codeMap);
        if (next) skuCell.setNumberFormat('@').setValue(next);
      }
    }
  } catch (err) {
    Logger.log(`onEdit error: ${err.message}`);
  }
}
function assignSkusForAll_() {
  try {
    const sh = sheet_(SHEET_NAME);
    const last = sh.getLastRow();
    if (last <= HEADER_ROW) return;
    const map = getCategoryCodeMap_();
    const data = sh.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, 2).getValues(); // [SKU, Category]
    for (let i = 0; i < data.length; i++) {
      const [sku, category] = data[i];
      if (!sku && category) {
        const next = generateNextSku_(sh, category, map);
        if (next) sh.getRange(HEADER_ROW + 1 + i, SKU_COL).setNumberFormat('@').setValue(next);
      }
    }
  } catch (err) {
    Logger.log(`assignSkusForAll_ error: ${err.message}`);
  }
}

/* ====== GOOGLE FORM HANDLER (optional) ====== */
const FORM_FIELD_MAP = {
  category: 'Category',
  name: 'Item Name / Description',
  price: 'Price',
  qty: 'Qty',
  storeLoc: 'Store Location',
  caseBinShelf: 'Case#/BIN#/Shelf#',
  onlineLoc: 'Online Location',
  status: 'Status'
};
function onFormSubmit(e) {
  try {
    const nv = (e && e.namedValues) ? e.namedValues : {};
    const keys = Object.keys(nv).map(k => k.trim());
    const getByPrefix = (prefix) => {
      const p = prefix.toLowerCase();
      const key = keys.find(k => k.toLowerCase().startsWith(p));
      return key ? (nv[key] || [''])[0] : '';
    };

    const category     = getByPrefix('category');
    const name         = getByPrefix('item name');
    const rawPrice     = getByPrefix('price');
    const qtyRaw       = getByPrefix('qty');
    const storeLoc     = getByPrefix('store location');
    const caseBinShelf = getByPrefix('case#/bin#/shelf#');
    const onlineLoc    = getByPrefix('online location');
    const status       = getByPrefix('status');

    let priceNum = parseFloat(String(rawPrice).replace(/[^\d.]/g, '')); if (isNaN(priceNum)) priceNum = 0;
    let qty      = parseInt(String(qtyRaw).replace(/[^\d]/g, ''), 10);   if (isNaN(qty))      qty = 1;

    const sh = sheet_(SHEET_NAME);
    const sku = generateNextSku_(sh, category); if (!sku) return;

    sh.appendRow([sku, category, name, priceNum, qty, storeLoc, caseBinShelf, onlineLoc, status]);
    const lr = sh.getLastRow();
    sh.getRange(lr, 1).setNumberFormat('@').setValue(sku);
    sh.getRange(lr, PRICE_COL).setNumberFormat('$#,##0.00');

    if (e && e.range) {
      const resp = e.range.getSheet();
      const row = e.range.getRow();
      const header = resp.getRange(1,1,1,resp.getLastColumn()).getValues()[0];
      let skuCol = header.indexOf('SKU') + 1;
      if (skuCol === 0) { skuCol = header.length + 1; resp.getRange(1, skuCol).setValue('SKU'); }
      resp.getRange(row, skuCol).setNumberFormat('@').setValue(sku);

      const respPriceCol = header.findIndex(h => String(h).toLowerCase().startsWith('price')) + 1;
      if (respPriceCol > 0) {
        resp.getRange(row, respPriceCol).setValue(priceNum);
        resp.getRange(2, respPriceCol, resp.getMaxRows()-1).setNumberFormat('$#,##0.00');
      }
    }
  } catch (err) {
    Logger.log(`onFormSubmit error: ${err.message}`);
    throw err;
  }
}

/* ====== INTAKE ENDPOINTS (web app) ====== */
function addItemAndReturnLabel(payload, submissionId, bulkFlag) {
  try {
    if (submissionId) {
      const cache = CacheService.getScriptCache();
      const seen = cache.get(submissionId);
      if (seen) return JSON.parse(seen);
    }
    const { category, name, price, qty, storeLoc, caseBinShelf, onlineLoc, status, btItemId, costOfGoods, condition, brand, primaryColor} = payload;

    const sh = sheet_(SHEET_NAME);
    let priceNum = parseFloat(String(price).replace(/[^\d.]/g, '')); if (isNaN(priceNum)) priceNum = 0;
    let qtyNum   = parseInt(String(qty).replace(/[^\d]/g, ''), 10);   if (isNaN(qtyNum))   qtyNum = 1;

    const sku = generateNextSku_(sh, category); if (!sku) throw new Error('Could not generate SKU');
    sh.appendRow([sku, category, name, priceNum, qtyNum, storeLoc, caseBinShelf, onlineLoc, status]);

    const lr = sh.getLastRow();
    sh.getRange(lr, 1).setNumberFormat('@').setValue(sku);
    sh.getRange(lr, 4).setNumberFormat('$#,##0.00');

    // Ensure and write Condition (required in UI, but write safely here)
    (function () {
      const col = ensureHeaderColumn_(sh, 'Condition');
      sh.getRange(lr, col).setNumberFormat('@').setValue(String(condition || ''));
    })();

    // Ensure and write Brand (required in UI, but write safely here)
    (function () {
      const col = ensureHeaderColumn_(sh, 'Brand');
      sh.getRange(lr, col).setNumberFormat('@').setValue(String(brand || ''));
    })();

    // Ensure and write primary color (required in UI, but write safely here)
    (function () {
      const col = ensureHeaderColumn_(sh, 'PrimaryColor');
      // NOTE: use the destructured `primaryColor` (lowercase p)
      sh.getRange(lr, col).setNumberFormat('@').setValue(String(primaryColor || ''));
    })();

    // NEW: ensure and write Cost_of_Goods on the appended row
    (function () {
      // Parse to number; leave blank if not provided
      let cogNum = parseFloat(String(costOfGoods || '').replace(/[^\d.]/g, ''));
      if (isNaN(cogNum)) return; // do nothing if empty/invalid

      const cogCol = ensureHeaderColumn_(sh, 'Cost_of_Goods'); // creates header if missing
      sh.getRange(lr, cogCol).setValue(cogNum).setNumberFormat('$#,##0.00');
    })();

    // NEW: mark bulk-intake rows as Pending
    (function () {
      // Accept either explicit boolean arg OR payload.bulkAdd truthy
      var isBulk = (bulkFlag === true) || (payload && (payload.bulkAdd === true || String(payload.bulkAdd || '').toLowerCase() === 'true'));
      if (!isBulk) return;

      var col = ensureHeaderColumn_(sh, 'BULK_ADD_STATUS');
      sh.getRange(lr, col).setNumberFormat('@').setValue('Pending');
    })();

    // NEW: ensure and write Buy_Ticket_ITEM_ID on the appended row
    if (btItemId) {
      const col = ensureHeaderColumn_(sh, 'Buy_Ticket_ITEM_ID');
      sh.getRange(lr, col).setNumberFormat('@').setValue(String(btItemId));
    }


    (function () {
      const display = String((payload && payload.vendooCategoryDisplay) || '').trim();
      if (!display) return; // nothing selected

      const map = getVendooPathMap_();
      const pathStr = map[display] || '';

      const colDisp = ensureHeaderColumn_(sh, 'Vendoo_Category_Display');
      sh.getRange(lr, colDisp).setNumberFormat('@').setValue(display);

      const colPath = ensureHeaderColumn_(sh, 'Vendoo_Category_Path');
      sh.getRange(lr, colPath).setNumberFormat('@').setValue(pathStr);
    })();

    // NEW: write Long_Description = Title + blank + Long Desc + blank + "<SKU Store Case>"
    (function () {
      var longDesc = String((payload && payload.longDescription) || '').trim();
      var tail = [sku, storeLoc, caseBinShelf].filter(Boolean).join(' ');
      var title = String(name || '');
      // Client already deduped the default clause; we store verbatim here.
      var finalText = [title, longDesc, tail].join('\n\n').trim();
      var colLD = ensureHeaderColumn_(sh, 'Long_Description');
      sh.getRange(lr, colLD).setNumberFormat('@').setValue(finalText);
    })();

    // NEW: Shipping Box + Weight/Dimensions
    (function () {
      function setText(colName, val) {
        var col = ensureHeaderColumn_(sh, colName);
        sh.getRange(lr, col).setNumberFormat('@').setValue(String(val || ''));
      }
      function setNum(colName, val, fmt) {
        var col = ensureHeaderColumn_(sh, colName);
        if (val === '' || val == null) { sh.getRange(lr, col).setValue(''); return; }
        var n = parseFloat(String(val).replace(/[^\d.]/g,''));
        if (!Number.isFinite(n)) { sh.getRange(lr, col).setValue(''); return; }
        sh.getRange(lr, col).setValue(n).setNumberFormat(fmt || '0.##');
      }

      var shipBox = (payload && payload.shippingBox) || '';
      setText('Shipping_Box', shipBox);

      setNum('Weight_LB', (payload && payload.weightLb), '0');      // integer ok
      setNum('Weight_OZ', (payload && payload.weightOz), '0');      // integer ok
      setNum('Length',    (payload && payload.length),   '0.0');
      setNum('Width',     (payload && payload.width),    '0.0');
      setNum('Height',    (payload && payload.height),   '0.0');
    })();
    
    
    // Ensure all pending sheet writes are committed before the client immediately reads.
    SpreadsheetApp.flush();

    const labelLine = [sku, storeLoc, caseBinShelf].filter(Boolean).join(' ');
    const result = {
      sku, name,
      priceDisplay: `$${priceNum.toFixed(2)}`,
      qty: qtyNum, storeLoc, caseBinShelf, onlineLoc, status, labelLine
    };

    if (submissionId) {
      const cache = CacheService.getScriptCache();
      cache.put(submissionId, JSON.stringify(result), 30);
    }
    return result;
  } catch (err) {
    const msg = `addItemAndReturnLabel error: ${err.message}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}

/** Persist Vendoo ID + URL into SKU Tracker for a given base SKU. */
function apiSku_SetVendooInfo(sku, vendooId, vendooUrl) {
  try {
    const baseSku = String(sku || '').trim();
    if (!baseSku) throw new Error('Missing SKU');

    const sh = sheet_(SHEET_NAME); // "SKU Tracker"
    const last = sh.getLastRow();
    if (last <= HEADER_ROW) throw new Error('Empty SKU Tracker');

    // Ensure/locate columns
    const colVendooId  = ensureHeaderColumn_(sh, 'Vendoo_Item_Number');
    const colVendooUrl = ensureHeaderColumn_(sh, 'Vendoo_ITEM_URL');

    // Find the exact SKU in column A
    const cell = sh.getRange('A:A').createTextFinder(baseSku).matchEntireCell(true).findNext();
    if (!cell) throw new Error('SKU not found: ' + baseSku);
    const row = cell.getRow();

    // Write values (Vendoo ID as text)
    sh.getRange(row, colVendooId ).setNumberFormat('@').setValue(String(vendooId || ''));
    sh.getRange(row, colVendooUrl).setValue(String(vendooUrl || ''));

    return { ok: true, row, vendooId, vendooUrl };
  } catch (err) {
    Logger.log('apiSku_SetVendooInfo error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

/**
 * Update BULK_ADD_STATUS and audit columns.
 * status: "In Progress" | "Completed" | "Error"
 * msg: optional detail (stored in BULK_VENDOO_RESULT)
 */
function apiBulk_SetStatus(sku, status, msg) {
  try {
    var baseSku = String(sku || '').trim();
    if (!baseSku) throw new Error('Missing SKU');

    var sh = sheet_(SHEET_NAME);
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last <= HEADER_ROW) throw new Error('Empty SKU Tracker');

    var header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    var H = function(n){ return header.indexOf(n) + 1; };

    var colStatus = ensureHeaderColumn_(sh, 'BULK_ADD_STATUS');
    var colRes    = ensureHeaderColumn_(sh, 'BULK_VENDOO_RESULT');
    var colAt     = ensureHeaderColumn_(sh, 'BULK_VENDOO_AT');

    // find row by exact SKU match
    var cell = sh.getRange('A:A').createTextFinder(baseSku).matchEntireCell(true).findNext();
    if (!cell) throw new Error('SKU not found: ' + baseSku);
    var r = cell.getRow();

    sh.getRange(r, colStatus).setNumberFormat('@').setValue(String(status || ''));
    if (status === 'Completed') {
      sh.getRange(r, colRes).setNumberFormat('@').setValue('Completed');
      sh.getRange(r, colAt).setValue(new Date());
    } else if (status === 'Error') {
      sh.getRange(r, colRes).setNumberFormat('@').setValue('Error: ' + String(msg || ''));
      sh.getRange(r, colAt).setValue(new Date());
    }
    return { ok: true, row: r };
  } catch (err) {
    Logger.log('apiBulk_SetStatus error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

/**
 * Build a new Google Sheet (in the given folder) with Line 1/Line 2 for ALL Pending items.
 * Also writes XEASY_SHEET_STATUS/AT/URL into each included row.
 */
// === Option B: rebuild an XLSX and overwrite the existing .xlsx file (same ID/link) ===
function apiBulk_CreateXeasySheet() {
  function log(m){ try{ Logger.log('[Xeasy->XLSXOverwriteAB] ' + m); }catch(_){} }

  try {
    var TARGET_XLSX_FILE_ID = '1Hpo2TRUEDwEv1A0ASg73NG5Bs9SiVtx3'; // your current link's ID
    var HEADERS = ['Line 1','Line 2'];

    // Build lines from Pending rows
    var sh   = sheet_(SHEET_NAME);
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last <= HEADER_ROW) return { ok:true, count:0, url:'' };

    var header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    var H = function(n){ return header.indexOf(n) + 1; };
    var colStatus = ensureHeaderColumn_(sh, 'BULK_ADD_STATUS');
    var colSku    = H('SKU');
    var colPrice  = H('Price') || 4;
    var colStore  = H('Store Location') || 6;
    var colBin    = H('Case#/BIN#/Shelf#') || 7;
    var colOnline = H('Online Location') || 8;

    var colSheetStatus = ensureHeaderColumn_(sh, 'XEASY_SHEET_STATUS');
    var colSheetAt     = ensureHeaderColumn_(sh, 'XEASY_SHEET_AT');
    var colSheetUrl    = ensureHeaderColumn_(sh, 'XEASY_SHEET_URL');

    var values = sh.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, lastCol).getValues();
    var lines = [], rowsIncluded = [];
    function fmtPrice(p){ var n = parseFloat(String(p).replace(/[^\d.]/g,'')); return Number.isFinite(n) ? ('$' + n.toFixed(2)) : ''; }

    for (var i=0; i<values.length; i++){
      var row    = values[i];
      var status = String(row[colStatus-1] || '').trim();
      if (status !== 'Pending') continue;

      var sku    = String(row[colSku-1]   || '');
      var price  = fmtPrice(row[colPrice-1]);
      var store  = String(row[colStore-1] || '');
      var bin    = String(row[colBin-1]   || '');
      var online = String(row[colOnline-1]|| '').toLowerCase();

      var needsVEN = (online.indexOf('vendoo') >= 0) || (online.indexOf('both') >= 0);
      var line1 = needsVEN ? ((price ? (price + ' ') : '') + 'VEN') : (price || '');
      var line2 = [sku, store, bin].filter(Boolean).join(' ');
      lines.push([line1, line2]);
      rowsIncluded.push(HEADER_ROW + 1 + i);
    }
    if (!lines.length) return { ok:true, count:0, url:'' };

    // Create a temp Google Sheet → fill A/B → export as XLSX
    var base = 'MRAD_Xeasy_Bulk_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Denver', 'yyyyMMdd_HHmm');
    var tmp  = SpreadsheetApp.create(base);
    var ws   = tmp.getActiveSheet();
    ws.clear();
    ws.getRange(1,1,1,2).setValues([HEADERS]);
    ws.getRange(2,1,lines.length,2).setValues(lines);
    SpreadsheetApp.flush();

    var xlsxBlob = DriveApp.getFileById(tmp.getId()).getBlob().getAs(MimeType.MICROSOFT_EXCEL);
    xlsxBlob.setName(base + '.xlsx');

    // Overwrite your existing XLSX file in place (same ID/link)
    Drive.Files.update({}, TARGET_XLSX_FILE_ID, xlsxBlob);

    // Clean up temp
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch(_){}

    var url = 'https://docs.google.com/spreadsheets/d/' + TARGET_XLSX_FILE_ID + '/edit';
    var now = new Date();
    rowsIncluded.forEach(function(r){
      sh.getRange(r, colSheetStatus).setNumberFormat('@').setValue('Created');
      sh.getRange(r, colSheetAt).setValue(now);
      sh.getRange(r, colSheetUrl).setNumberFormat('@').setValue(url);
    });
    SpreadsheetApp.flush();

    return { ok:true, count:lines.length, url:url, id:TARGET_XLSX_FILE_ID };

  } catch (err) {
    log('ERROR: ' + (err && err.message ? err.message : err));
    return { ok:false, error:(err && err.message ? err.message : String(err)) };
  }
}

// --- TEMP DIAG: creates a tiny CSV in your target folder to force Drive auth
function apiDiag_CreateCsvTest() {
  const FOLDER_ID = '1Xw7wvWBNq5ROC8w_7GjIYOJ8HvccCT4u';
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const name   = 'MRAD_Xeasy_TEST_' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone() || 'America/Denver', 'yyyyMMdd_HHmmss'
    ) + '.csv';
    const blob   = Utilities.newBlob('hello,world\n1,2\n', 'text/csv', name);
    const file   = folder.createFile(blob);
    return { ok:true, url:file.getUrl(), id:file.getId() };
  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * List rows from SKU Tracker where BULK_ADD_STATUS == "Pending".
 * Returns minimal fields needed to render a list (SKU, Name, Price, Qty, Locs, Vendoo hints).
 */
function apiIntake_ListBulkPending(limit) {
  try {
    var N = Math.max(1, Number(limit) || 100);
    var sh = sheet_(SHEET_NAME);
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow <= HEADER_ROW) return { ok: true, items: [] };

    var header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    var H = function(name){ return header.indexOf(name) + 1; };

    var colStatus = ensureHeaderColumn_(sh, 'BULK_ADD_STATUS'); // ensures column exists
    var colSku = H('SKU'), colName = H('Item Name / Description') || H('name') || 3;

    // (kept but not used in the list UI anymore; safe for other flows)
    var colPrice = H('Price') || 4, colQty = H('Qty') || 5;
    var colStore = H('Store Location') || 6, colBin = H('Case#/BIN#/Shelf#') || 7;
    var colVendooDisp = H('Vendoo_Category_Display');

    // NEW: fields to show in Bulk list
    var colVendooUrl = ensureHeaderColumn_(sh, 'Vendoo_ITEM_URL');
    var colSheetUrl  = ensureHeaderColumn_(sh, 'XEASY_SHEET_URL');
    var colBulkRes   = ensureHeaderColumn_(sh, 'BULK_VENDOO_RESULT');

    // Read all (could optimize with filters later)
    var values = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
    var out = [];
    for (var i=0; i<values.length; i++){
      var row = values[i];
      var status = String(row[colStatus-1] || '').trim();
      // Show anything that still needs attention; skip Completed and blank cells
      if (!status || status === 'Completed') continue;
      out.push({
        row: HEADER_ROW + 1 + i,
        sku: String(row[colSku-1] || ''),
        name: String(row[colName-1] || ''),
        status: status,
        vendooUrl: String(row[colVendooUrl-1] || ''),
        xeasySheetUrl: String(row[colSheetUrl-1] || ''),
        bulkVendooResult: String(row[colBulkRes-1] || '')
      });
      if (out.length >= N) break;
    }
    return { ok: true, items: out };
  } catch (err) {
    Logger.log('apiIntake_ListBulkPending error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

/**
 * Full dataset for bulk runner. Includes Long_Description and Shipping fields.
 * Returns newest-first order hint via "row".
 */
function apiBulk_ListPendingFull(limit) {
  try {
    var N = Math.max(1, Number(limit) || 500);
    var sh = sheet_(SHEET_NAME);
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow <= HEADER_ROW) return { ok: true, items: [] };

    var header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    var H = function(name){ return header.indexOf(name) + 1; };

    var colStatus = ensureHeaderColumn_(sh, 'BULK_ADD_STATUS');
    var colSku    = H('SKU');
    var colName   = H('Item Name / Description') || H('name') || 3;
    var colPrice  = H('Price') || 4, colQty = H('Qty') || 5;
    var colStore  = H('Store Location') || 6, colBin = H('Case#/BIN#/Shelf#') || 7;
    var colOnline = H('Online Location') || 8;

    var colCond   = ensureHeaderColumn_(sh, 'Condition');
    var colBrand  = ensureHeaderColumn_(sh, 'Brand');
    var colPrim   = ensureHeaderColumn_(sh, 'PrimaryColor');
    var colVendooDisp = ensureHeaderColumn_(sh, 'Vendoo_Category_Display');
    var colLong   = ensureHeaderColumn_(sh, 'Long_Description');
    var colCogs   = ensureHeaderColumn_(sh, 'Cost_of_Goods');

    var colShipBox = ensureHeaderColumn_(sh, 'Shipping_Box');
    var colLb = ensureHeaderColumn_(sh, 'Weight_LB');
    var colOz = ensureHeaderColumn_(sh, 'Weight_OZ');
    var colLen = ensureHeaderColumn_(sh, 'Length');
    var colWid = ensureHeaderColumn_(sh, 'Width');
    var colHei = ensureHeaderColumn_(sh, 'Height');

    var values = sh.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
    var out = [];
    for (var i=0; i<values.length; i++){
      var row = values[i];
      var status = String(row[colStatus-1] || '').trim();
      // Only process valid queued statuses; blanks and anything else are excluded
      if (!(status === 'Pending' || status === 'In Progress')) continue;

      out.push({
        row: HEADER_ROW + 1 + i,
        sku: String(row[colSku-1] || ''),
        name: String(row[colName-1] || ''),
        price: row[colPrice-1],
        qty:   row[colQty-1],
        storeLoc: String(row[colStore-1] || ''),
        caseBinShelf: String(row[colBin-1] || ''),
        onlineLoc: String(row[colOnline-1] || ''),
        condition: String(row[colCond-1] || ''),
        brand:     String(row[colBrand-1] || ''),
        primaryColor: String(row[colPrim-1] || ''),
        vendooCategoryDisplay: String(row[colVendooDisp-1] || ''),
        longDescriptionFull: String(row[colLong-1] || ''), // full value, per your directive
        costOfGoods: row[colCogs-1],
        // shipping
        shippingBox: String(row[colShipBox-1] || ''),
        weightLb: row[colLb-1],
        weightOz: row[colOz-1],
        length:  row[colLen-1],
        width:   row[colWid-1],
        height:  row[colHei-1]
      });

      if (out.length >= N) break;
    }
    // newest-first by sheet row number (higher row => newer)
    out.sort(function(a,b){ return (b.row|0) - (a.row|0); });
    return { ok: true, items: out };
  } catch (err) {
    Logger.log('apiBulk_ListPendingFull error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// --- Duplicate source loader by SKU ---
// Returns the item's fields but leaves Title blank on the client and strips certain lines from Long_Description.
function apiIntake_GetItemForDuplicate(sku) {
  try {
    sku = String(sku || '').trim();
    if (!sku) return { ok: false, error: 'Missing SKU' };

    var sh = sheet_(SHEET_NAME);
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow <= HEADER_ROW) return { ok: false, error: 'Empty sheet' };

    var header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    var H = function(name){ return header.indexOf(name) + 1; };

    var colSku    = H('SKU');
    var colName   = H('Item Name / Description') || H('name') || 3;
    var colPrice  = H('Price') || 4;
    var colQty    = H('Qty') || 5;
    var colStore  = H('Store Location') || 6;
    var colBin    = H('Case#/BIN#/Shelf#') || 7;
    var colOnline = H('Online Location') || 8;

    var colCond   = ensureHeaderColumn_(sh, 'Condition');
    var colBrand  = ensureHeaderColumn_(sh, 'Brand');
    var colPrim   = ensureHeaderColumn_(sh, 'PrimaryColor');
    var colVendooDisp = ensureHeaderColumn_(sh, 'Vendoo_Category_Display');
    var colLong   = ensureHeaderColumn_(sh, 'Long_Description');
    var colCogs   = ensureHeaderColumn_(sh, 'Cost_of_Goods');

    var colShipBox = ensureHeaderColumn_(sh, 'Shipping_Box');
    var colLb = ensureHeaderColumn_(sh, 'Weight_LB');
    var colOz = ensureHeaderColumn_(sh, 'Weight_OZ');
    var colLen = ensureHeaderColumn_(sh, 'Length');
    var colWid = ensureHeaderColumn_(sh, 'Width');
    var colHei = ensureHeaderColumn_(sh, 'Height');

    // Find the row for this SKU
    var skuColVals = sh.getRange(HEADER_ROW + 1, colSku, lastRow - HEADER_ROW, 1).getValues().flat().map(String);
    var idx = skuColVals.indexOf(sku);
    if (idx < 0) return { ok: false, error: 'SKU not found: ' + sku };
    var row = HEADER_ROW + 1 + idx;

    var get = (c)=> sh.getRange(row, c).getValue();

    var name        = String(get(colName) || '');
    var storeLoc    = String(get(colStore) || '');
    var caseBin     = String(get(colBin) || '');
    var longFull    = String(get(colLong) || '');
    var DEF         = "The photos are part of the description. Be sure to look them over for condition and details. This is sold as is and it's ready for a new home."; // mirror client

    // Strip: Title line, default blurb, and "SKU StoreLoc Case#" line (exact matches)
    function stripForDuplicate(txt){
      var lines = String(txt||'').split(/\r?\n/).map(function(s){ return String(s||'').trim(); });
      var skuLine = [sku, storeLoc, caseBin].filter(Boolean).join(' ');
      var out = [];
      for (var i=0; i<lines.length; i++){
        var line = lines[i];
        if (!line) continue;
        if (name && line === name) continue;      // remove title line
        if (DEF && line === DEF) continue;        // remove default blurb
        if (skuLine && line === skuLine) continue;// remove SKU/Loc/Case line
        out.push(line);
      }
      return out.join('\n\n');
    }

    var item = {
      // identity & context
      sku: sku,
      name: name,

      // basics
      category: String(get(H('Category')) || ''),
      price: get(colPrice),
      qty: get(colQty),
      storeLoc: storeLoc,
      caseBinShelf: caseBin,
      onlineLoc: String(get(colOnline) || ''),
      costOfGoods: get(colCogs),

      // vendoo
      vendooCategoryDisplay: String(get(colVendooDisp) || ''),
      condition: String(get(colCond) || ''),
      brand: String(get(colBrand) || ''),
      primaryColor: String(get(colPrim) || ''),

      // long description (stripped on client too, but send it here for convenience)
      longDescriptionFull: stripForDuplicate(longFull),

      // shipping
      shippingBox: String(get(colShipBox) || ''),
      weightLb: get(colLb),
      weightOz: get(colOz),
      length: get(colLen),
      width: get(colWid),
      height: get(colHei)
    };

    return { ok: true, item: item };
  } catch (err) {
    Logger.log('apiIntake_GetItemForDuplicate error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

function updateItemBySku(update) {
  const sh = sheet_(SHEET_NAME);
  const data = sh.getRange(2, 1, Math.max(0, sh.getLastRow()-1), 1).getValues().flat();
  const idx = data.indexOf(String(update.sku));
  if (idx === -1) throw new Error('SKU not found: ' + update.sku);

  const row = 2 + idx;
  const priceNum = (function(x){ const n = parseFloat(String(x).replace(/[^\d.]/g,'')); return isNaN(n)?0:n; })(update.price);
  const qtyNum   = (function(x){ const n = parseInt(String(x).replace(/[^\d]/g,''),10); return isNaN(n)?1:n; })(update.qty);

  const values = [
    update.sku, update.category, update.name, priceNum, qtyNum,
    update.storeLoc, update.caseBinShelf, update.onlineLoc, update.status
  ];
  sh.getRange(row, 1, 1, values.length).setValues([values]);
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 4).setNumberFormat('$#,##0.00');

  // Keep Condition and Brand in sync on edits (safe header ensure)
  (function(){
    const condCol = ensureHeaderColumn_(sh, 'Condition');
    sh.getRange(row, condCol).setNumberFormat('@').setValue(String(update.condition || ''));
  })();
  (function(){
    const brandCol = ensureHeaderColumn_(sh, 'Brand');
    sh.getRange(row, brandCol).setNumberFormat('@').setValue(String(update.brand || ''));
  })();

  (function(){
    const primaryColorCol = ensureHeaderColumn_(sh, 'PrimaryColor');
    sh.getRange(row, primaryColorCol).setNumberFormat('@').setValue(String(update.primaryColor || ''));
  })();

  const priceDisplay = `$${priceNum.toFixed(2)}`;
  const labelLine = [update.sku, update.storeLoc, update.caseBinShelf].filter(Boolean).join(' ');
  return { ok: true, name: update.name, priceDisplay, labelLine };
}
function deleteItemBySku(sku) {
  const sh = sheet_(SHEET_NAME);
  const last = sh.getLastRow();
  if (last <= 1) throw new Error('No rows to delete.');
  const skus = sh.getRange(2, 1, last - 1, 1).getValues().flat();
  const idx = skus.indexOf(String(sku));
  if (idx === -1) throw new Error('SKU not found: ' + sku);
  sh.deleteRow(2 + idx);
  return { ok: true };
}

/** Return the most recent N drafts for Intake (default: 20). */
function apiIntake_GetRecentDrafts(limit) {
  var _logs = [];
  function push(msg){
    var line = '[GetRecentDrafts] ' + msg;
    try { console.log(line); } catch(_) {}
    try { Logger.log(line); } catch(_) {}
    _logs.push(line);
  }

  try {
    var N = Math.max(1, Number(limit) || 20);
    push('start N=' + N);

    // Prefer the bound spreadsheet if available; fall back to your helper if you use one.
    var ss = SpreadsheetApp.getActive();
    if (!ss) { push('No active spreadsheet'); throw new Error('No active spreadsheet'); }

    var sh = ss.getSheetByName('Intake Draft');
    if (!sh) { push('Sheet "Intake Draft" not found'); return { ok: true, drafts: [], _logs }; }

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    push('lastRow=' + lastRow + ', lastCol=' + lastCol);

    if (lastRow < 2) { push('No data rows'); return { ok: true, drafts: [], _logs }; }

    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    push('header=' + JSON.stringify(header));

    var rowsToRead = Math.min(N, lastRow - 1);
    var startRow   = Math.max(2, lastRow - rowsToRead + 1);
    push('startRow=' + startRow + ', rowsToRead=' + rowsToRead);

    var values = sh.getRange(startRow, 1, rowsToRead, lastCol).getValues();
    var H = function(n) { return header.indexOf(n); };

    var out = values.map(function(r) {
      var ts = H('Timestamp') >= 0 ? r[H('Timestamp')] : '';
      var tsMs = (ts && ts.getTime) ? ts.getTime() : (typeof ts === 'number' ? ts : null);

      return {
        timestampMs:  tsMs, // normalized for front-end
        shortDesc:    H('ShortDesc') >= 0 ? String(r[H('ShortDesc')] || '') : '',
        longDesc:     H('LongDesc')  >= 0 ? String(r[H('LongDesc')]  || '') : '',
        recPrice:     H('RecommendedPrice') >= 0 ? r[H('RecommendedPrice')] : null,
        altPrice:     H('AltPrice')  >= 0 ? r[H('AltPrice')]  : null,
        imageFileId:  H('ImageFileId') >= 0 ? String(r[H('ImageFileId')] || '') : ''
      };
    });

    out.reverse(); // newest first
    push('returning drafts=' + out.length);
    return { ok: true, drafts: out, _logs };
  } catch (e) {
    push('ERROR: ' + (e && e.message ? e.message : String(e)));
    return { ok: false, error: e && e.message ? e.message : String(e), _logs };
  }
}


