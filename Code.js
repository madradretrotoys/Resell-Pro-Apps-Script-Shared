/************************************************************
 * Resell Pro — Unified Code.gs (Base functions shared)
 * Drop-in replacement
 ************************************************************/

/* ====== CONFIG ====== */
const SPREADSHEET_ID = '1YEi4fZejWvTloqlydhavTO3Cllz6ikVoItSqom1GS_k';

const SHEET_NAME   = 'SKU Tracker';
const CODE_SHEET   = 'Category Codes';     // A: Category Name, B: Category Code
const SKU_COL      = 1;                    // Column A
const CATEGORY_COL = 2;                    // Column B
const HEADER_ROW   = 1;
const SKU_PAD      = 4;                    // 0001, 0002, etc.

const FORM_URL = 'https://docs.google.com/forms/d/e/1FAIpQLSeJU5e1NxSiWfAjKqGt5Zo3oXn6kDAWsP_KOFXQuGHnHZ-FtQ/viewform?usp=sharing&ouid=115942254913738298628';
const PRICE_COL = 4; // [SKU, Category, Name, Price, Qty, Store, Case/Bin/Shelf, Online, Status]
const USERS_SHEET = 'User Permissions';
const U_COL = { USER:1, LOGIN:2, PASS:3, ROLE:4, EDIT:5, VIEW:6, POS:7, INTAKE:8, MAXDISC:9, DISCPASS:10 };
// Use store timezone explicitly
const CASH_TZ = 'America/Denver';



/* ====== SIMPLE EXPORTS ====== */
function ping() { return 'pong ' + new Date(); }




/* ====== CORE OPENERS ====== */
function book_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found by ID.');
    return ss;
  } catch (err) {
    const msg = `book_ error: ${err && err.message ? err.message : err}`;
    Logger.log(msg);
    throw new Error(msg);
  }
}
function sheet_(name) {
  const ss = book_();
  const sh = ss.getSheetByName(name);
  if (!sh) {
    const msg = `Missing sheet/tab: "${name}"`;
    Logger.log(msg);
    throw new Error(msg);
  }
  return sh;
}




/* ====== MENU & STARTUP ====== */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Mad Rad Tools')
      .addItem('Open Item Form (link)', 'showFormLink_')
      .addItem('Assign SKUs (fill blanks)', 'assignSkusForAll_')
      .addItem('Refresh Sales Summary', 'refreshSalesSummary')
      .addToUi();

    const sh = sheet_(SHEET_NAME);
    sh.getRange('A2:A').setNumberFormat('@');           // SKU as text
    sh.getRange('D2:D').setNumberFormat('$#,##0.00');   // Price currency

    setCategoryDropdown_();
  } catch (err) {
    Logger.log(`onOpen error: ${err.message}`);
  }
}
function showFormLink_() {
  SpreadsheetApp.getUi().alert('Item Entry Form:\n' + FORM_URL);
}

/* ====== WEB APP ROUTING (TEMPLATED) ====== */
function apiGetBase() {
  // Returns the canonical /exec or /dev URL for this deployment.
  return ScriptApp.getService().getUrl();
}

/** Returns { ok, url, depId, isDev, version, description, label, source }  */
function apiGetVersionTag() {
  const url = ScriptApp.getService().getUrl() || '';
  const m = url.match(/\/s\/([a-zA-Z0-9_-]+)\/(exec|dev)/);
  const depId = m ? m[1] : '';
  const isDev = m ? (m[2] === 'dev') : false;

  // Single source of truth shown on the badge
  const props = PropertiesService.getScriptProperties();
  let propVer = (props.getProperty('APP_VERSION') || '').trim();

  // Default response uses the stored property (works for dev & prod)
  let version = propVer ? (Number(propVer) || propVer) : null;
  let description = '';
  let label = isDev ? 'DEV' : (propVer ? 'v' + String(propVer) : 'PROD');
  let source = 'properties';

  // In production, discover the real deployment version and sync APP_VERSION if needed
  if (!isDev && depId) {
    try {
      const infos = ScriptApp.getDeploymentInfo() || [];
      const hit =
        infos.find(d => d.getDeploymentId() === depId) ||
        infos.find(d => depId.startsWith(d.getDeploymentId())); // loose prefix match

      if (hit && typeof hit.getVersionNumber === 'function') {
        const realVer = hit.getVersionNumber();   // integer like 41, 42, ...
        description = hit.getDescription() || '';
        if (typeof realVer === 'number' && realVer > 0) {
          if (String(realVer) !== propVer) {
            // AUTO-SYNC the property so the badge stays correct thereafter
            props.setProperty('APP_VERSION', String(realVer));
            propVer = String(realVer);
          }
          version = realVer;
          label = 'v' + String(realVer);
          source = 'deployment';
        }
      }
    } catch (e) {
      // ignore; keep using the property
      source = 'properties-fallback';
    }
  }

  return { ok: true, url, depId, isDev, version, description, label, source };
}

/* Optional convenience helpers — lets you set APP_VERSION without opening Project Settings */
function setAppVersion(version) {
  PropertiesService.getScriptProperties().setProperty('APP_VERSION', String(version));
  return 'APP_VERSION set to ' + String(version);
}
function setAppVersionPrompt() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Set Production Version', 'Enter version number to display (e.g., 40):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    return setAppVersion(res.getResponseText().trim());
  }
  return 'Cancelled';
}

  

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Hex helper used by older challenge handler builds.
// Safe to keep even if your current code doesn't call it.
function toHex_(bytes) {
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function exportSource_(p) {
  var provided = (p && p.key) || '';
  var expect = PropertiesService.getScriptProperties().getProperty('SOURCE_EXPORT_KEY') || '';
  if (!expect || provided !== expect) {
    return ContentService.createTextOutput('forbidden').setMimeType(ContentService.MimeType.TEXT);
  }

  // Advanced Service “Script” (Apps Script API v1) must be added under Services.
  var scriptId = ScriptApp.getScriptId();
  var content  = Script.Projects.getContent(scriptId);   // { files:[...] }

  var files = (content && content.files) ? content.files.map(function (f) {
    return { name: f.name, type: f.type, source: f.source || '' };
  }) : [];

  var payload = { projectId: scriptId, ts: new Date().toISOString(), files: files };
  return ContentService.createTextOutput(JSON.stringify(payload))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) { 
  var p = (e && e.parameter) || {};

  if (String(p.export || p.action).toLowerCase() === 'source') return exportSource_(p);

  

  // --- 0) Lightweight JSON health response for external validators (Valor)
  // Use any of these query strings: ?ping=1  OR  ?source=valor  OR  ?hook=valor
  // Example to give Valor:  https://.../exec?source=valor
  if (p.ping == '1' || String(p.source || '').toLowerCase() === 'valor' || String(p.hook || '').toLowerCase() === 'valor') {
    try {
      var out = {
        ok: true,
        service: 'Resell Pro',
        hook: 'valor',
        ping: true,
        ts: new Date().toISOString()
      };
      return ContentService
        .createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      var fail = { ok:false, error:String(err) };
      return ContentService
        .createTextOutput(JSON.stringify(fail))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- 1) eBay Marketplace Account Deletion verification (GET) ---
  if (p.challenge_code) {
    var props    = PropertiesService.getScriptProperties();
    var token    = (props.getProperty('EBAY_DELETE_VERIFY_TOKEN') || '').trim();
    var endpoint = (props.getProperty('EBAY_DELETE_ENDPOINT') || '').trim();
    if (!endpoint) endpoint = ScriptApp.getService().getUrl();

    var challenge = String(p.challenge_code);

    // Compute lowercase hex SHA-256 of challenge + token + endpoint
    var bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      challenge + token + endpoint,
      Utilities.Charset.UTF_8
    );
    var hex = bytes.map(function (b) {
      var v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? '0' + v : v;
    }).join('');

    // TEMP LOGS (show in Apps Script → Executions → click the doGet row)
    try {
      console.log(JSON.stringify({
        kind: 'ebay-verify',
        challenge_len: challenge.length,
        endpoint_len: endpoint.length,
        token_len: token.length,
        hash_prefix: hex.slice(0, 8)
      }));
    } catch (_) {}

    var out2 = { challengeResponse: hex };
    if (p.debug == '1') {
      out2._debug = { endpoint: endpoint, endpointLen: endpoint.length, tokenLen: token.length };
    }

    return ContentService.createTextOutput(JSON.stringify(out2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // --- 2) Your normal router (unchanged) ---
  var page = p.page || 'login'; // default

  var view = (page === 'pos')        ? 'pos'
           : (page === 'drawer')     ? 'drawer'
           : (page === 'payout')     ? 'payout'
           : (page === 'refund')     ? 'refund'
           : (page === 'research2')  ? 'research2'
           : (page === 'research')   ? 'research'
           : (page === 'estimate_drop')   ? 'estimate_drop'     // NEW: customer form
           : (page === 'estimate')  ? 'estimate'    // NEW: internal queue
           : (page === 'bticket')  ? 'bticket'    // NEW: internal queue
           : (page === 'timesheet')   ? 'timesheet'
           : (page === 'login')      ? 'login'
           : 'intake';

  var t = HtmlService.createTemplateFromFile(view);
  t.execUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
           .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const EBAY_WORKER_KEY = props.getProperty('EBAY_WORKER_KEY') || '';
    const p = (e && e.parameter) || {};

    // ---- Guard: if caller claims to be our Worker (x_worker_key present) but key is wrong, ignore it
    if (p.x_worker_key && p.x_worker_key !== EBAY_WORKER_KEY) {
      // No logging, no processing (keeps your executions clean)
      return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
    }

    // Raw body once (used in routes below)
    const raw = (e && e.postData && e.postData.contents) || '';

    // ---- Route: Valor webhook (explicit query parameter)
    if (String(p.source || '') === 'valor') {
      return valorWebhookHandlerV2_(raw); // NEW: route to the renamed V2 handler that expects raw JSON
    }

    // ---- Route: Valor webhook (by payload having an invoicenumber)
    try {
      const body = raw ? JSON.parse(raw) : {};
      const inv =
        (body && (body.invoicenumber || body.INVOICENUMBER)) ||
        (body && body.data && (body.data.invoicenumber || body.data.INVOICENUMBER)) ||
        (body && body.reference_descriptive_data &&
          (body.reference_descriptive_data.invoicenumber || body.reference_descriptive_data.INVOICENUMBER)) ||
        '';
      if (String(inv || '').trim()) {
        return valorWebhookHandlerV2_(raw); // invoice-only reconciliation path
      }
    } catch (_err) {
      // ignore JSON parse errors; fall through to other routes
    }

    // ---- Route: eBay marketplace-account-deletion (must come via Cloudflare Worker)
    try {
      if (p.x_worker_key && EBAY_WORKER_KEY && p.x_worker_key === EBAY_WORKER_KEY) {
        // This is our Worker forwarding the eBay notification (already ACKed at the edge).
        // Keep it super light – tiny log only.
        const raw = (e.postData && e.postData.contents) || '';
        console.log('eBay delete (via Worker)', raw.slice(0, 256));
        return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
      }
    } catch (err) {
      console.warn('doPost eBay worker check failed:', err);
    }

    // ---- Unknown POSTs -> do nothing (quiet 200). Prevents bot noise from crowding logs.
    return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Validate a manager/admin passcode for approving over-limit discounts.
 * Called by the client: google.script.run.apiValidateDiscountPasscode(code)
 * Returns true (accepted) or false (rejected). Never throws to the client.
 */
function apiValidateDiscountPasscode(code) {
  try {
    // Normalize the input the same way we store it (trim + de-dash/space + lowercase)
    const input = String(code || '').trim();
    if (!input) return false;
    const normIn = input.replace(/[\s-]/g, '').toLowerCase();

    // Load allowed users from the "User Permissions" sheet
    const users = readUsers_(); // already maps { role, discPass, ... }

    // Define who can approve. Adjust list if you use a different role naming.
    const APPROVER_ROLES = /^(manager|admin|owner|supervisor)$/i;

    // Look for a matching passcode on an approver.
    const match = users.find(u => {
      const pass = String(u.discPass || '').trim();
      if (!pass) return false;
      const normPass = pass.replace(/[\s-]/g, '').toLowerCase();

      // Require role OR give a special case for “Full” discount users if you prefer:
      // const hasUnlimited = normalizeMaxDiscount_(u.maxDiscount) >= 1000;
      const canApprove = APPROVER_ROLES.test(String(u.role || ''));

      return canApprove && normPass === normIn;
    });

    // (Optional) audit log line – won’t break the UI if logging fails
    if (match) {
      try { Logger.log('Discount override approved by: %s (%s)', match.user, match.role); } catch (_){}
    }

    return !!match;
  } catch (_err) {
    // Never block the UI with an exception; just deny on error
    return false;
  }
}

function setAppVersion(v){ PropertiesService.getScriptProperties().setProperty('APP_VERSION', String(v)); }
function setAppVersionPrompt(){ var ui=SpreadsheetApp.getUi(); var r=ui.prompt('Set Production Version','Version (e.g., 42):',ui.ButtonSet.OK_CANCEL); if(r.getSelectedButton()===ui.Button.OK){ return setAppVersion(r.getResponseText().trim()); } return 'Cancelled'; }

/* ==========================================================
   ================   Login Permissions HELPERS   ================
   ========================================================== */

function readUsers_() {
  const cache = CacheService.getScriptCache();
  const key = 'mrad_users_v1';
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const sh = sheet_(USERS_SHEET); // your helper with underscore
  const vals = sh.getDataRange().getValues();
  vals.shift(); // header
  const rows = vals.filter(r => String(r[U_COL.LOGIN-1]).trim());
  const users = rows.map(r => ({
    user:        String(r[U_COL.USER-1]).trim(),
    login:       String(r[U_COL.LOGIN-1]).trim(),
    pass:        String(r[U_COL.PASS-1]).trim(),
    role:        String(r[U_COL.ROLE-1]).trim() || 'Clerk',
    canEdit:     String(r[U_COL.EDIT-1]).trim().toUpperCase() === 'Y',
    canView:     String(r[U_COL.VIEW-1]).trim().toUpperCase() === 'Y',
    canPOS:      String(r[U_COL.POS-1]).trim().toUpperCase() === 'Y',
    canIntake:   String(r[U_COL.INTAKE-1]).trim().toUpperCase() === 'Y',
    maxDiscount: String(r[U_COL.MAXDISC-1]).trim() || '0%',
    discPass:    String(r[U_COL.DISCPASS-1]).trim()
  }));

  cache.put(key, JSON.stringify(users), 600); // cache for 2 minutes
  return users;
}

// rudimentary token store (8 hours)
function makeToken_() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t=''; for (let i=0;i<40;i++) t += chars.charAt(Math.floor(Math.random()*chars.length));
  return t;
}
function saveSession_(token, profile) {
  const cache = CacheService.getScriptCache();
  cache.put('sess_'+token, JSON.stringify(profile), 8*60*60);
}
function loadSession_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('sess_'+token);
  return raw ? JSON.parse(raw) : null;
}
function killSession_(token) {
  const cache = CacheService.getScriptCache();
  cache.remove('sess_'+token);
}

// ===== Public API called from HTML =====
// Warm the container + prime the users cache so first login is fast
function apiWarmLogin() {
  try { readUsers_(); } catch (e) {}
  return 'ok';
}

function apiLogin(login, pass) {
  const users = readUsers_();
  const u = users.find(x => x.login === String(login).trim());
  if (!u || u.pass !== String(pass)) throw new Error('Invalid credentials');
  const profile = {
    user: u.user,
    login: u.login,
    role: u.role,
    canEdit: u.canEdit,
    canView: u.canView,
    canPOS: u.canPOS,
    canIntake: u.canIntake,
    maxDiscountPct: normalizeMaxDiscount_(u.maxDiscount),
    approverPasscodePresent: !!u.discPass
  };
  const token = makeToken_();
  saveSession_(token, profile);
  return { token, profile };
}

function apiVerify(token) {
  const p = loadSession_(token);
  if (!p) throw new Error('Session expired');
  return p;
}

function apiLogout(token) {
  killSession_(token);
  return true;
}









