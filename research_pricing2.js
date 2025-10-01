/** research_pricing2.gs â€” Research 2 API (image-first + eBay active/sold) */

const EBAY_FINDING_ENDPOINT = 'https://svcs.ebay.com/services/search/FindingService/v1';
const EBAY_APP_ID = PropertiesService.getScriptProperties().getProperty('EBAY_APP_ID'); // set once in Script Properties
const MAX_FINDING_RETRIES = 3;
const BACKOFF_START_MS = 600;   // first retry ~0.6s
const BACKOFF_FACTOR   = 2.2;   // 0.6s â†’ ~1.3s â†’ ~2.9s
const BACKOFF_JITTER   = 0.35;  // Â±35% jitter

/* ------------------- props ------------------- */
function R2_props_() {
  const p = PropertiesService.getScriptProperties();
  return {
    GCV_API_KEY:    p.getProperty('GCV_API_KEY'),
    GEMINI_API_KEY: p.getProperty('GEMINI_API_KEY'),
    GEMINI_MODEL:   p.getProperty('GEMINI_MODEL') || 'gemini-1.5-flash',
    EBAY_APP_ID:    p.getProperty('EBAY_APP_ID') || p.getProperty('EBAY_APPID') || '',
    EBAY_GLOBAL_ID: p.getProperty('EBAY_GLOBAL_ID') || 'EBAY-US',
    WORKER_EBAY_SEARCH_URL:p.getProperty('WORKER_EBAY_SEARCH_URL') || '',
    WORKER_KEY:     p.getProperty('WORKER_KEY') || ''
  };
}

function props_() {
  const p = PropertiesService.getScriptProperties();
  return {
    GCV_API_KEY: p.getProperty('GCV_API_KEY'),
    GEMINI_API_KEY: p.getProperty('GEMINI_API_KEY'),
    GEMINI_MODEL: p.getProperty('GEMINI_MODEL') || 'gemini-1.5-flash',
    IMG_FOLDER_ID: p.getProperty('INVENTORY_IMG_FOLDER_ID')
  };
}



/* ------------------- local debug helpers ------------------- */
function R2_dbgPush_(dbg, msg){
  try {
    const stamp = Utilities.formatDate(new Date(), 'Etc/UTC', 'HH:mm:ss.SSS');
    const line = '['+stamp+'] ' + msg;
    console.log(line);
    dbg.push(line);
  } catch (_) {}
}
function R2_peek_(s, n){ try{ s=String(s||''); return s.length>n? (s.slice(0,n)+'â€¦') : s; }catch(_){ return ''; } }

function R2_isEbayBotWallHtml_(html){
  const t = String(html||'').toLowerCase();
  // classic CAPTCHA + eBay's JS-shell "x-page-config" gate and common phrases
  return /verify you'?re not a robot|to continue, please verify|enter the characters you see|captcha|x-page-config/i.test(t);
}

function sleepMs(ms){ Utilities.sleep(ms); }
function jitter(ms){
  const j = BACKOFF_JITTER;
  return Math.max(0, ms * (1 + (Math.random() * 2 * j - j)));
}

function isBotWallHtml(html){
  if (!html) return false;
  const s = String(html);
  // very stable signatures we saw in your sample
  return (
    /radware_stormcaster/i.test(s) ||
    /SSJSConnectorObj|ssConf\("c1"|__uzdbm_/i.test(s) ||
    /There seems to be a problem serving the request at this time/i.test(s)
  );
}

// saves a short sample of the blocked page for debugging (you already do this â€” keep your version if you prefer)
function saveBotWallSample(html, tag){
  try{
    const fileName = `ebay-active-botwall-${tag}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss'Z'")}.txt`;
    DriveApp.createFile(fileName, html || '', MimeType.PLAIN_TEXT);
  }catch(e){ /* non-fatal */ }
}

/* ------------------------------------------------------------------ */
/* IMAGE â†’ EVIDENCE (Cloud Vision WebDetection + OCR)                  */
/* ------------------------------------------------------------------ */

function apiR2_Analyze(imageBase64, mimeType, meta, avoidSignatures) {
  const out = { ok:true, evidence:{ ocr:'', pages:[], images:[], thumbs:[], titles:[] }, candidates:[] };

  try {
    const gcv = R2_gcvWeb_(imageBase64||'');
    out.evidence.ocr    = gcv.ocr || '';
    out.evidence.pages  = (gcv.pages  || []).slice(0, 24);
    out.evidence.images = (gcv.images || []).slice(0, 24);
    out.evidence.titles = R2_fetchTitles_(out.evidence.pages, 16);
    out.evidence.thumbs = R2_makeThumbs_(out.evidence.images, out.evidence.pages, 18);
  } catch (e) {
    return { ok:false, error: 'Cloud Vision error: '+ (e && e.message ? e.message : e) };
  }

  const cand = R2_geminiCandidatesFromEvidence_(
    imageBase64||'',
    mimeType||'image/jpeg',
    meta||{},
    out.evidence.ocr,
    out.evidence.pages,
    out.evidence.titles,
    12
  );
  if (!cand.ok) return { ok:false, error:cand.error || 'candidate error' };

  const avoid = new Set((avoidSignatures||[]).map(s=>String(s).toLowerCase()));
  const mapped = [];
  for (let i=0;i<cand.candidates.length;i++){
    const c = cand.candidates[i];
    const sig = R2_candSig_(c);
    if (avoid.size && avoid.has(sig.toLowerCase())) continue;

    const src   = R2_bestSourceForHint_(c.srcHint || '', out.evidence.pages, out.evidence.titles);
    const thumb = out.evidence.thumbs[i % Math.max(1, out.evidence.thumbs.length)] || (out.evidence.images[i] || '');

    mapped.push({
      name: c.name||'',
      brand: c.brand||'',
      line: c.line||'',
      variant: c.variant||'',
      year: c.year||'',
      confidence: typeof c.confidence==='number' ? c.confidence : null,
      query: R2_buildQueryFromMeta_(meta||{}, c),
      links: R2_ebayLinks_(R2_buildQueryFromMeta_(meta||{}, c)),
      thumb: thumb,
      src: src,
      sig: sig
    });
  }
  out.candidates = mapped;
  return out;
}

/* ------------------------------------------------------------------ */
/* IMAGE â†’ TEMP PUBLIC URL (Drive one-time staging for Google Lens)   */
/* ------------------------------------------------------------------ */
function apiR2_StageImage(imageBase64, mimeType) {
  const out = { ok:false, url:'', fileId:'', error:'' };
  try {
    if (!imageBase64) { out.error = 'No image'; return out; }

    // Decode base64
    const parts = String(imageBase64).split(',');
    const b64 = parts.length > 1 ? parts[1] : parts[0];
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType || 'image/jpeg', 'resellpro.jpg');

    // Ensure a temp folder
    const folderName = 'ResellPro Temp Images';
    const root = DriveApp.getRootFolder();
    let folder = null;

    // Try to find existing folder (cheap linear scan; this is fine for a single folder)
    const it = DriveApp.getFoldersByName(folderName);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);

    // Create file, set anyone-with-link viewable
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Build direct fetch URL
    const fileId = file.getId();
    // Using uc?export=download yields a direct binary fetchable by Googleâ€™s fetchers.
    const directUrl = 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId);

    out.ok = true;
    out.fileId = fileId;
    out.url = directUrl;
    return out;
  } catch (e) {
    out.error = (e && e.message) ? e.message : String(e);
    return out;
  }
}

function R2_gcvWeb_(imageBase64){
  const key = R2_props_().GCV_API_KEY;
  if (!key) throw new Error('Missing GCV_API_KEY');
  const url = 'https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(key);
  const payload = {
    requests: [{
      image: { content: imageBase64 },
      features: [
        { type: 'WEB_DETECTION',  maxResults: 80 },
        { type: 'TEXT_DETECTION', maxResults: 1 }
      ],
      imageContext: { webDetectionParams: { includeGeoResults: false } }
    }]
  };
  const res  = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload: JSON.stringify(payload), muteHttpExceptions:true });
  const code = res.getResponseCode();
  const body = res.getContentText() || '';
  if (code !== 200) throw new Error('GCV HTTP ' + code + ' â€” ' + R2_peek_(body, 300));

  const data = JSON.parse(body);
  const a  = ((data.responses||[])[0]||{});
  const wd = a.webDetection || {};
  const ocr = (a.textAnnotations && a.textAnnotations[0] && a.textAnnotations[0].description) || '';

  const pageUrls = (wd.pagesWithMatchingImages || []).map(x => x && x.url).filter(Boolean);

  const imageUrls = []
    .concat((wd.fullMatchingImages     || []).map(x => x && x.url))
    .concat((wd.partialMatchingImages  || []).map(x => x && x.url))
    .concat((wd.visuallySimilarImages  || []).map(x => x && x.url))
    .filter(Boolean);

  const pageOut = []; const pageSeen = Object.create(null);
  pageUrls.forEach(u => { try{ const cu = new URL(u); const k = cu.hostname.replace(/^www\./,'') + cu.pathname; if (!pageSeen[k]) { pageSeen[k]=1; pageOut.push(u); } } catch(_){ } });

  const imgOut = []; const imgSeen = Object.create(null);
  imageUrls.forEach(u => { if (!imgSeen[u]) { imgSeen[u]=1; imgOut.push(u); } });

  return { pages: pageOut, images: imgOut, ocr: ocr };
}

// --- NEW: worker-first active
function R2_ebayActive_(q, dbg) {
  const cleaned = R2_stripNeg_(q || '');
  const w = R2_workerFetch_('active', cleaned, 60, dbg);
  if (w.ok && (w.items||[]).length) return { ok:true, items:w.items, source:'API' };
  // fall back to legacy calls only if worker empty
  return R2_ebayRssActive_(cleaned, dbg);
}

// --- NEW: worker-first sold
function R2_ebaySold_(q, dbg) {
  const cleaned = R2_stripNeg_(q || '');
  const w = R2_workerFetch_('sold', cleaned, 60, dbg);
  if (w.ok && (w.items||[]).length) return { ok:true, items:w.items, source:'API' };
  // as a last resort, RSS (avoids eBay bot-wall and 10001)
  return R2_ebayRssSold_(cleaned, dbg);
}

function R2_fetchTitles_(urls, limit){
  const out = [];
  const list = (urls||[]).slice(0, Math.min(limit||16, (urls||[]).length));
  if (!list.length) return out;
  const reqs = list.map(u => ({ url:u, muteHttpExceptions:true, followRedirects:true }));
  let resps = [];
  try { resps = UrlFetchApp.fetchAll(reqs); } catch(_) { resps = []; }

  for (let i=0;i<list.length;i++){
    try{
      const r = resps[i];
      if (!r) { out.push(''); continue; }
      if (r.getResponseCode() >= 400) { out.push(''); continue; }
      const html = r.getContentText();
      const m =
        html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)/i) ||
        html.match(/<meta[^>]+name=["']twitter:title["'][^>]+content=["']([^"']+)/i) ||
        html.match(/<title[^>]*>([^<]+)<\/title>/i);
      out.push(m ? (m[1] || '').trim() : '');
    } catch(_) { out.push(''); }
  }
  return out;
}

function R2_findTitleForUrl_(pageUrl){
  try{
    const res = UrlFetchApp.fetch(pageUrl, {muteHttpExceptions:true, followRedirects:true});
    if (res.getResponseCode() >= 400) return '';
    const html = res.getContentText();
    const m =
      html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)/i) ||
      html.match(/<meta[^>]+name=["']twitter:title["'][^>]+content=["']([^"']+)/i) ||
      html.match(/<title[^>]*>([^<]+)<\/title>/i);
    return m ? (m[1] || '').trim() : '';
  } catch(_) { return ''; }
}

/* ------------------------------------------------------------------ */
/* CANDIDATES (Gemini with evidence titles + OCR)                      */
/* ------------------------------------------------------------------ */

function R2_geminiCandidatesFromEvidence_(imageBase64, mimeType, meta, ocr, pages, titles, topN){
  const { GEMINI_MODEL } = R2_props_();
  const N = topN || 12;

  const prompt = [
    'You are an expert toy identifier for vintage/modern collectibles.',
    'Given: a product photo, OCR text from the image, and titles of web pages containing matching/visually-similar images.',
    'Return up to '+N+' plausible candidates (brand, line/series, product/character, variant/color, year).',
    'Use the OCR and page titles to increase recall. Prefer exact character/variant names found in titles.',
    'Confidence is from 0.0â€“1.0. If uncertain, still include plausible options with lower confidence.',
    'Also return a "srcHint" as a short snippet (either a page title or host) that best supports that candidate.',
    'STRICT JSON ONLY:\n{"candidates":[{"name":"","brand":"","line":"","variant":"","year":"","confidence":0.0,"srcHint":""}]}'
  ].join('\n');

  const evidenceBlock = [
    'OCR:\n'+(ocr||'(none)'),
    '\nTITLES:\n'+(titles||[]).slice(0,30).map((t,i)=>((i+1)+'. '+(t||''))).join('\n')
  ].join('\n');

  const parts = [{ text: prompt }, { text: 'FIELDS:\n' + JSON.stringify(meta||{}, null, 0) }, { text: evidenceBlock }];
  if (imageBase64) parts.push({ inline_data:{ mime_type: mimeType||'image/jpeg', data:imageBase64 } });

  let r = geminiCall_(GEMINI_MODEL, parts, { temperature:0.15, maxOutputTokens:900 });
  if (!r.ok) r = geminiCall_('gemini-1.5-flash-8b', parts, { temperature:0.15, maxOutputTokens:900 });
  if (!r.ok) return { ok:false, error:r.error||'candidate error' };

  let obj=null; try { obj=JSON.parse(r.text); } catch(_){ const m=r.text.match(/\{[\s\S]*\}/); if(m){ try{ obj=JSON.parse(m[0]); }catch(_){}}}
  const list = (obj && obj.candidates) || [];

  const seen = new Set();
  const dedup = [];
  list.forEach(c=>{
    const key = R2_candSig_(c).toLowerCase();
    if (!seen.has(key)) { seen.add(key); dedup.push(c); }
  });

  return { ok:true, candidates: dedup.slice(0, N) };
}

function R2_bestSourceForHint_(hint, pages, titles){
  if (!hint) return pages && pages[0] ? pages[0] : '';
  for (let i=0;i<titles.length;i++){ const t=titles[i]||''; if (t && t.toLowerCase().indexOf(String(hint).toLowerCase())>=0) return pages[i]||''; }
  try{
    const want = String(hint).toLowerCase().replace(/^www\./,'');
    for (let i=0;i<pages.length;i++){ const h=new URL(pages[i]).hostname.replace(/^www\./,'').toLowerCase(); if (h.indexOf(want)>=0) return pages[i]; }
  } catch(_){}
  return pages && pages[0] ? pages[0] : '';
}

/* ------------------------------------------------------------------ */
/* THUMBNAILS                                                          */
/* ------------------------------------------------------------------ */

function R2_makeThumbs_(imageUrls, pageUrls, maxCount){
  const want = Math.max(6, maxCount || 18);
  const out = [];
  (imageUrls || []).some(u => { if (out.length >= want) return true; out.push(u); return false; });

  if (out.length < want && pageUrls && pageUrls.length){
    const need = want - out.length;
    const pages = pageUrls.slice(0, Math.min(need, 10));
    const reqs = pages.map(u => ({ url:u, muteHttpExceptions:true, followRedirects:true }));
    let resps = [];
    try { resps = UrlFetchApp.fetchAll(reqs); } catch(_) { resps=[]; }

    for (let i=0;i<pages.length && out.length<want;i++){
      try{
        const r = resps[i]; if (!r || r.getResponseCode() >= 400) continue;
        const html = r.getContentText();
        let img =
          (html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)/i) || [])[1] ||
          (html.match(/<meta[^>]+name=["']twitter:image(:src)?["'][^>]+content=["']([^"']+)/i) || [])[2] ||
          (html.match(/<link[^>]+rel=["']image_src["'][^>]+href=["']([^"']+)/i) || [])[1] || '';
        if (img) {
          if (img.startsWith('//')) img = 'https:' + img;
          if (img.startsWith('/')) {
            try { const uo = new URL(pages[i]); img = uo.origin + img; } catch(_){}
          }
          out.push(img);
        }
      } catch(_) {}
    }
  }
  return out.slice(0, want);
}

function fetchJsonWithRetry_(url, payloadObj, tries, retryableCodes) {
  tries = tries || 4;
  retryableCodes = retryableCodes || {429:1, 500:1, 502:1, 503:1, 504:1};
  let waitMs = 600; // 0.6s -> 1.2s -> 2.4s -> 4.8s
  for (let i = 0; i < tries; i++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payloadObj),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (!retryableCodes[code]) return { code, text: res.getContentText() };
    Utilities.sleep(waitMs);
    waitMs *= 2;
  }
  const last = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payloadObj),
    muteHttpExceptions: true,
  });
  return { code: last.getResponseCode(), text: last.getContentText() };
}


function geminiCall_(model, parts, config) {
  const key = props_().GEMINI_API_KEY;
  if (!key) return { ok:false, error:'Missing GEMINI_API_KEY' };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              encodeURIComponent(model) + ':generateContent?key=' +
              encodeURIComponent(key);
  const payload = {
    contents: [{ role:'user', parts: parts }],
    generationConfig: Object.assign({ temperature:0.15, maxOutputTokens:600 }, config || {})
  };

  const r = fetchJsonWithRetry_(url, payload, 4, {429:1,500:1,502:1,503:1,504:1});
  if (r.code !== 200) return { ok:false, error:'Gemini '+r.code, raw:r.text };

  let data;
  try { data = JSON.parse(r.text); } catch(e){ return { ok:false, error:'Gemini parse error' }; }
  const text = ((data.candidates && data.candidates[0] && data.candidates[0].content &&
                 data.candidates[0].content.parts) || [])
                .map(p => p.text || '').join('').trim();
  return { ok:true, text };
}


/* ------------------------------------------------------------------ */
/* QUERY / LINKS / SIGS                                                */
/* ------------------------------------------------------------------ */

function R2_buildQueryFromMeta_(m, c){
  const parts=[]; const add = x => { if (x) parts.push(String(x)); };
  if (c) { add(c.brand); add(c.line); add(c.name); add(c.variant); add(c.year); }
  else   { add(m.brand); add(m.series); add(m.name||m.product); add(m.year); }
  return parts.join(' ').trim();
}
function R2_ebayLinks_(q){
  const clean = R2_stripNeg_(q || '');
  const enc   = encodeURIComponent(clean);
  return {
    active: 'https://www.ebay.com/sch/i.html?_nkw=' + enc + '&LH_BIN=1',
    sold:   'https://www.ebay.com/sch/i.html?_nkw=' + enc + '&LH_Sold=1&LH_Complete=1',
    preview:'https://www.google.com/search?tbm=isch&q=' + encodeURIComponent(clean)
  };
}
function R2_candSig_(c){ return [c.brand,c.line,c.name,c.variant,c.year].filter(Boolean).join('|'); }

/* ------------------------------------------------------------------ */
/* EBAY (low-level fetch, caching, helpers)                            */
/* ------------------------------------------------------------------ */

function R2_ebayErr_(res){
  var code = res.getResponseCode();
  var body = '';
  try { body = String(res.getContentText() || ''); } catch(_) {}
  if (body && body.length > 400) body = body.slice(0, 400) + 'â€¦';
  return 'eBay HTTP ' + code + (body ? (' â€” ' + body) : '');
}
function R2_isEbayRateLimit_(res){
  try {
    var code = res.getResponseCode();
    var body = String(res.getContentText() || '');
    return (code >= 500) && /RateLimit/i.test(body);
  } catch(_) { return false; }
}

function R2_cacheGet_(key){
  try { const v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch(_) { return null; }
}
function R2_cachePut_(key, obj, seconds){
  try { CacheService.getScriptCache().put(key, JSON.stringify(obj), Math.max(0, seconds||30)); } catch(_) {}
}

/**
 * Low-level fetch with:
 *  - Global script lock (serialize concurrent calls)
 *  - Server-side throttle between calls (minGapMs)
 *  - 5 retries w/ exponential backoff + jitter on 500 RateLimiter
 *  - Sends eBay SOA headers in addition to query params
 */
function R2_ebayFetch_(params, appId, globalId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (_){}

  try {
    const minGapMs = 1600;
    const cache = CacheService.getScriptCache();
    const lastTs = Number(cache.get('ebay:lastTs') || '0');
    const now = Date.now();
    const delta = now - lastTs;
    if (lastTs && delta < minGapMs) Utilities.sleep(minGapMs - delta);
    cache.put('ebay:lastTs', String(Date.now()), 10);

    const base = 'https://svcs.ebay.com/services/search/FindingService/v1?' + toQS_(params);
    const opts = {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'X-EBAY-SOA-SECURITY-APPNAME': appId,
        'X-EBAY-SOA-GLOBAL-ID': globalId,
        'X-EBAY-SOA-OPERATION-NAME': params['OPERATION-NAME']
      }
    };

    for (var attempt = 0; attempt < 5; attempt++) {
      const res = UrlFetchApp.fetch(base, opts);
      const code = res.getResponseCode();
      if (code === 200) return { ok:true, res };

      if (R2_isEbayRateLimit_(res) && attempt < 4) {
        const sleepMs = (1500 * Math.pow(2, attempt)) + Math.floor(Math.random()*500);
        Utilities.sleep(sleepMs);
        continue;
      }
      return { ok:false, res };
    }
    return { ok:false, res:null };
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/* ---------- HTML fallback (parse public search page) --------------- */

function R2_parseMoney_(txt){
  if (!txt) return null;
  const m = String(txt).replace(/,/g,'').match(/\$?\s*([\d]+(?:\.\d{1,2})?)/);
  return m ? Number(m[1]) : null;
}

function R2_htmlFetch_(url){
  const opts = {
    muteHttpExceptions:true,
    followRedirects:true,
    headers:{
      'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36',
      'Accept-Language':'en-US,en;q=0.9',
      'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Referer':'https://www.ebay.com/'
    }
  };
  return UrlFetchApp.fetch(url, opts);
}

function R2_parseEbayList_(html, assumeSold){
  const out = [];
  if (!html) return out;

  // Accept either <li ... class="s-item"> or <div ... class="s-item">
  const re = /<(?:li|div)[^>]+class="[^"]*\bs-item\b[^"]*"[\s\S]*?<\/(?:li|div)>/gi;
  let m;
  while ((m = re.exec(html)) && out.length < 60) {
    const block = m[0];

    // URL
    let url = '';
    const mLink = block.match(/<a[^>]+class="[^"]*\bs-item__link\b[^"]*"[^>]+href="([^"]+)"/i);
    if (mLink) url = mLink[1];

    // Title (some rows use <h3>, others <div>)
    let title = '';
    const mTitle =
      block.match(/<h3[^>]*class="[^"]*\bs-item__title\b[^"]*"[^>]*>([\s\S]*?)<\/h3>/i) ||
      block.match(/<div[^>]*class="[^"]*\bs-item__title\b[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
    if (mTitle) title = mTitle[1].replace(/<[^>]+>/g,'').replace(/\s+/g,' ').trim();

    // Image
    let img = '';
    const mImg = block.match(/<img[^>]+class="[^"]*\bs-item__image-img\b[^"]*"[^>]+(?:src|data-src)="([^"]+)"/i);
    if (mImg) img = mImg[1];

    // Price
    let priceTxt = '';
    const mPrice = block.match(/<span[^>]*class="[^"]*\bs-item__price\b[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
    if (mPrice) priceTxt = mPrice[1].replace(/<[^>]+>/g,'');
    const price = R2_parseMoney_(priceTxt);

    // Shipping
    let shipTxt = '';
    const mShip = block.match(/<span[^>]*class="[^"]*\b(s-item__shipping|s-item__logisticsCost)\b[^"]*"[^>]*>([\s\S]*?)<\/span>/i);
    if (mShip) shipTxt = mShip[2].replace(/<[^>]+>/g,'');
    let ship = 0;
    if (/free/i.test(shipTxt)) ship = 0;
    else {
      const shipNum = R2_parseMoney_(shipTxt);
      ship = (shipNum==null) ? 0 : shipNum;
    }

    if (!title || !url) continue;
    const total = (price!=null) ? +(price + (ship||0)).toFixed(2) : null;

    out.push({
      id:null, title:title, price:price, shipping:ship, total:total,
      url:url, image:img, sold: !!assumeSold
    });
  }
  return out;
}

function R2_ebayHtmlActive_(q){
  const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_BIN=1&_ipg=60';
  const res = R2_htmlFetch_(url);
  if (res.getResponseCode() >= 400) return { ok:false, error:'eBay HTML fetch failed ('+res.getResponseCode()+')' };
  const html = res.getContentText() || '';
  try { Logger.log('HTML(active) len=' + html.length); } catch(_){}

  // Bot wall â†’ dump and bail
  if (R2_isEbayBotWallHtml_(html)) {
    const link = R2_dumpToDrive_('ebay-active-botwall-'+q, html);
    return { ok:false, error:'eBay HTML bot-check page encountered'+(link?(' â€” sample: '+link):'') };
  }

  const items = R2_parseEbayList_(html, false);
  if (!items.length) {
    // Parsing produced zero â€” save sample so we can adjust the parser
    const link = R2_dumpToDrive_('ebay-active-zeroitems-'+q, html);
    return { ok:true, items:[], debug:[ 'Active HTML returned 0 items â€” sample saved: '+link ] };
  }
  return { ok:true, items: items };
}

function R2_ebayHtmlSold_(q){
  const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_Sold=1&LH_Complete=1&_ipg=60';
  const res = R2_htmlFetch_(url);
  if (res.getResponseCode() >= 400) return { ok:false, error:'eBay HTML fetch failed ('+res.getResponseCode()+')' };
  const html = res.getContentText() || '';
  try { Logger.log('HTML(sold) len=' + html.length); } catch(_){}

  // Bot wall â†’ dump and bail
  if (R2_isEbayBotWallHtml_(html)) {
    const link = R2_dumpToDrive_('ebay-sold-botwall-'+q, html);
    return { ok:false, error:'eBay HTML bot-check page encountered'+(link?(' â€” sample: '+link):'') };
  }

  const items = R2_parseEbayList_(html, true);
  if (!items.length) {
    const link = R2_dumpToDrive_('ebay-sold-zeroitems-'+q, html);
    return { ok:true, items:[], stats:{min:null,med:null,max:null}, debug:[ 'Sold HTML returned 0 items â€” sample saved: '+link ] };
  }
  // If you want stats here, keep your existing stats logic elsewhere
  return { ok:true, items: items };
}

// --- Finding API with backoff (Active) ---
function R2_ebayFindingActive_(q, dbg){
  const { EBAY_APP_ID, EBAY_GLOBAL_ID } = R2_props_();
  const app = EBAY_APP_ID || '';
  const gid = EBAY_GLOBAL_ID || 'EBAY-US';
  if (!app) return { ok:false, error:'Missing EBAY_APP_ID', items:[] };

  const baseParams = {
    'OPERATION-NAME':'findItemsByKeywords',
    'SERVICE-VERSION':'1.13.0',
    'SECURITY-APPNAME': app,
    'GLOBAL-ID': gid,
    'RESPONSE-DATA-FORMAT':'JSON',
    'keywords': q,
    'paginationInput.entriesPerPage':'60',
    'itemFilter(0).name':'ListingType',
    'itemFilter(0).value(0)':'FixedPrice',
    'itemFilter(0).value(1)':'StoreInventory',
    'itemFilter(0).value(2)':'AuctionWithBIN',
    'sortOrder':'BestMatch',
    'outputSelector(0)':'SellerInfo',
    'outputSelector(1)':'PictureURLLarge',
    'outputSelector(2)':'StoreInfo'
  };
  R2_dbgPush_(dbg, 'Active: API params => '+ JSON.stringify(baseParams));

  let delay = BACKOFF_START_MS || 600;
  let lastErr = '';
  for (let attempt = 1; attempt <= (MAX_FINDING_RETRIES||3); attempt++){
    const url = EBAY_FINDING_ENDPOINT + '?' + Object.keys(baseParams)
      .map(k => encodeURIComponent(k)+'='+encodeURIComponent(baseParams[k]))
      .join('&');

    let res=null, code=0, body='';
    try{
      res  = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
      code = res.getResponseCode();
      body = res.getContentText() || '';
    }catch(e){
      lastErr = 'fetch error: ' + (e && e.message ? e.message : String(e));
      R2_dbgPush_(dbg, 'Active: API fetch threw â€” '+ lastErr);
      code = 599;
    }

    if (code===200) {
      try{
        const data = JSON.parse(body);
        const searchResult = (((data || {}).findItemsByKeywordsResponse || [])[0] || {}).searchResult || [];
        const itemsRaw = (searchResult[0] && searchResult[0].item) || [];
        const items = itemsRaw.map(it => {
          const priceRaw = (((((it.sellingStatus||[])[0]||{}).currentPrice||[])[0]||{}).__value__);
          const shipRaw  = (((((it.shippingInfo ||[])[0]||{}).shippingServiceCost||[])[0]||{}).__value__);
          const price    = (priceRaw === undefined || priceRaw === null || priceRaw === '') ? null : Number(priceRaw);
          const ship     = (shipRaw  === undefined || shipRaw  === null || shipRaw  === '') ? 0    : Number(shipRaw);
          const img      = ((it.pictureURLLarge||[])[0]) || ((it.galleryURL||[])[0]) || '';
          const url      = (it.viewItemURL ||[])[0] || '';
          const title    = (it.title        ||[])[0] || '';
          const total    = (price!=null) ? +(price + (ship||0)).toFixed(2) : null;
          return { id:null, title, price, shipping:ship, total, url, image:img, sold:false };
        }).filter(x => x && x.title && x.url);

        return { ok:true, items, source:'API' };
      }catch(e){
        lastErr = 'json parse error: '+(e && e.message ? e.message : String(e));
        R2_dbgPush_(dbg, 'Active: API 200 but parse failed â€” '+ lastErr);
        break; // donâ€™t retry parse errors
      }
    }

    // 500 + RateLimiter "10001" â‡’ backoff and retry
    if (code===500 && /10001/.test(body)) {
      R2_dbgPush_(dbg, 'Active: API fail â€” eBay HTTP 500 (10001 RateLimiter); attempt '+attempt);
      if (attempt < (MAX_FINDING_RETRIES||3)) {
        sleepMs(jitter(delay));
        delay = Math.floor(delay * (BACKOFF_FACTOR||2.2));
        continue;
      }
      lastErr = 'eBay HTTP 500 â€” RateLimiter (10001)';
      break;
    }

    // Other errors â€” stop (no retry loop for nonâ€‘10001)
    lastErr = 'API HTTP '+code+' â€” '+ R2_peek_(body, 180);
    R2_dbgPush_(dbg, 'Active: API fail â€” ' + lastErr);
    break;
  }

  return { ok:false, error:lastErr||'API error', items:[] };
}


//SOLD
// --- Finding API with backoff (Active) ---
function R2_ebayFindingSold_(q, dbg){
  const { EBAY_APP_ID, EBAY_GLOBAL_ID } = R2_props_();
  const app = EBAY_APP_ID || '';
  const gid = EBAY_GLOBAL_ID || 'EBAY-US';
  if (!app) return { ok:false, error:'Missing EBAY_APP_ID', items:[] };

  const baseParams = {
    'OPERATION-NAME':'findCompletedItems',
    'SERVICE-VERSION':'1.13.0',
    'SECURITY-APPNAME': app,
    'GLOBAL-ID': gid,
    'RESPONSE-DATA-FORMAT':'JSON',
    'keywords': q,
    'paginationInput.entriesPerPage':'60',
    'itemFilter(0).name':'Condition',           // keep defaults; adjust if you like
    'itemFilter(0).value(0)':'Used',
    'sortOrder':'EndTimeSoonest',
    'outputSelector(0)':'SellerInfo',
    'outputSelector(1)':'PictureURLLarge',
    'outputSelector(2)':'StoreInfo'
  };
  R2_dbgPush_(dbg, 'Sold: API params => '+ JSON.stringify(baseParams));

  let delay = BACKOFF_START_MS || 600;
  let lastErr = '';
  for (let attempt = 1; attempt <= (MAX_FINDING_RETRIES||3); attempt++){
    const url = EBAY_FINDING_ENDPOINT + '?' + Object.keys(baseParams)
      .map(k => encodeURIComponent(k)+'='+encodeURIComponent(baseParams[k]))
      .join('&');

    let res=null, code=0, body='';
    try{
      res  = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
      code = res.getResponseCode();
      body = res.getContentText() || '';
    }catch(e){
      lastErr = 'fetch error: ' + (e && e.message ? e.message : String(e));
      R2_dbgPush_(dbg, 'Sold: API fetch threw â€” '+ lastErr);
      code = 599;
    }

    if (code===200) {
      try{
        const data = JSON.parse(body);
        const searchResult = (((data || {}).findCompletedItemsResponse || [])[0] || {}).searchResult || [];
        const itemsRaw = (searchResult[0] && searchResult[0].item) || [];
        const items = itemsRaw.map(it => {
          const price = Number(((((it.sellingStatus||[])[0]||{}).currentPrice||[])[0]||{}).__value__) || null;
          const ship  = Number(((((it.shippingInfo ||[])[0]||{}).shippingServiceCost||[])[0]||{}).__value__) || 0;
          const img   = ((it.pictureURLLarge||[])[0]) || ((it.galleryURL||[])[0]) || '';
          const url   = (it.viewItemURL ||[])[0] || '';
          const title = (it.title        ||[])[0] || '';
          const state = String((((it.sellingStatus||[])[0]||{}).sellingState||[])[0]||'').toLowerCase();
          const sold  = (state === 'endedwithsales');
          const total = (price!=null) ? +(price + (ship||0)).toFixed(2) : null;
          return { id:null, title, price, shipping:ship, total, url, image:img, sold:true };
        }).filter(x => x && x.title && x.url);

        return { ok:true, items, source:'API' };
      }catch(e){
        lastErr = 'json parse error: '+(e && e.message ? e.message : String(e));
        R2_dbgPush_(dbg, 'Active: API 200 but parse failed â€” '+ lastErr);
        break; // donâ€™t retry parse errors
      }
    }

    // 500 + RateLimiter "10001" â‡’ backoff and retry
    if (code===500 && /10001/.test(body)) {
      R2_dbgPush_(dbg, 'Active: API fail â€” eBay HTTP 500 (10001 RateLimiter); attempt '+attempt);
      if (attempt < (MAX_FINDING_RETRIES||3)) {
        sleepMs(jitter(delay));
        delay = Math.floor(delay * (BACKOFF_FACTOR||2.2));
        continue;
      }
      lastErr = 'eBay HTTP 500 â€” RateLimiter (10001)';
      break;
    }

    // Other errors â€” stop (no retry loop for nonâ€‘10001)
    lastErr = 'API HTTP '+code+' â€” '+ R2_peek_(body, 180);
    R2_dbgPush_(dbg, 'Active: API fail â€” ' + lastErr);
    break;
  }

  return { ok:false, error:lastErr||'API error', items:[] };
}
//end sold



// --- RSS fallback (works without bot-wall & without API key) ---
function R2_ebayRssActive_(q, dbg){
  const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_BIN=1&_ipg=60&_rss=1';
  return R2_fetchEbayRss_(url, false, dbg);
}
function R2_ebayRssSold_(q, dbg){
  const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_Sold=1&LH_Complete=1&_ipg=60&_rss=1';
  return R2_fetchEbayRss_(url, true, dbg);
}

function R2_fetchEbayRss_(url, assumeSold, dbg){
  try{
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true});
    const code = res.getResponseCode(); const xml = res.getContentText() || '';
    if (code >= 400) return { ok:false, error:'eBay RSS fetch failed ('+code+')', items:[] };
    const items = R2_parseEbayRss_(xml, !!assumeSold);
    return { ok:true, items, source:'RSS' };
  }catch(e){
    return { ok:false, error:(e && e.message ? e.message : String(e)), items:[] };
  }
}

function R2_parseEbayRss_(xml, assumeSold){
  const out = [];
  const itemRe = /<item>([\s\S]*?)<\/item>/gi; let m;
  while((m=itemRe.exec(xml))){
    const block = m[1];

    // title (often "Title â€” $Price" or similar)
    let title = ''; let price = null;
    const t = (block.match(/<title>([\s\S]*?)<\/title>/i)||[])[1] || '';
    const cleanTitle = t.replace(/<!\[CDATA\[|\]\]>/g, '').trim();
    // simple price pull from title if present
    const pm = cleanTitle.match(/\$\s?([\d,]+(?:\.\d{1,2})?)/);
    if (pm) price = Number(pm[1].replace(/,/g,''));
    title = cleanTitle;

    // link
    const link = ((block.match(/<link>([\s\S]*?)<\/link>/i)||[])[1]||'').trim();

    if (!title || !link) continue;
    out.push({
      id:null, title, price: (price!=null?price:null), shipping: null,
      total: price, url: link, image:'', sold: !!assumeSold
    });
  }
  return out;
}

/* ---------- Public API (uses JSON API, falls back to HTML) --------- */

function R2_stripNeg_(q){
  return String(q||'')
    .replace(/\s-\s*"(?:[^"]*)"/g, '')   // -"multi word"
    .replace(/\s-\s*'(?:[^']*)'/g, '')   // -'multi word'
    .replace(/\s-\s*\S+/g, '')           // -word
    .replace(/\s+/g, ' ')
    .trim();
}

function R2_workerFetch_(kind, query, ipg, dbg) {
  const { WORKER_EBAY_SEARCH_URL, WORKER_KEY } = R2_props_();
  if (!WORKER_EBAY_SEARCH_URL || !WORKER_KEY) return { ok:false, error:'worker not configured', items:[] };

  const url = WORKER_EBAY_SEARCH_URL.replace(/\/+$/,'') + '/ebay/' + (kind === 'sold' ? 'sold' : 'active') +
              '?q=' + encodeURIComponent(query || '') + '&ipg=' + encodeURIComponent(ipg || 60);

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + WORKER_KEY },
    muteHttpExceptions: true,
    followRedirects: true
  });

  const code = res.getResponseCode();
  const txt  = res.getContentText() || '';
  if (dbg) R2_dbgPush_(dbg, 'Worker '+kind+' HTTP '+code+' Â· peek=' + R2_peek_(txt, 160));

  if (code !== 200) return { ok:false, error:'worker http '+code, items:[] };

  let data = null; try { data = JSON.parse(txt); } catch (_){}
  if (!data || data.ok !== true) return { ok:false, error: (data && data.error) || 'worker bad json', items:[] };

  // normalize to your internal item schema
  const items = (data.items || []).map(it => ({
    title: it.title || '',
    url: it.url || '',
    image: it.image || '',
    price: typeof it.price === 'number' ? it.price : null,
    shipping: typeof it.shipping === 'number' ? it.shipping : null,
    total: typeof it.total === 'number'
            ? it.total
            : (typeof it.price === 'number' && typeof it.shipping === 'number'
                ? (it.price + it.shipping)
                : null),
    sold: !!it.sold
  }));

  // Pull through the worker's reported total (true total matches on eBay)
  const totalRaw = data && data.stats ? data.stats.total : null;
  const totalNum = (typeof totalRaw === 'number')
      ? totalRaw
      : (totalRaw != null && String(totalRaw).trim() !== '' ? Number(totalRaw) : null);
  const total = Number.isFinite(totalNum) ? totalNum : null;

  return { ok:true, source: data.source || 'HTML', note: data.note || '', total, items };
}


// --- helper to normalize query input from client (string OR {raw, cleaned, appPresent})
function R2_qStr_(q) {
  if (q == null) return '';
  if (typeof q === 'string') return q;
  if (typeof q === 'object') return String(q.cleaned || q.raw || '');
  return String(q);
}

function R2_qCoerce_(q){
  // Accepts string or object and returns {raw, cleaned, appPresent}
  if (typeof q === 'string' || q == null){
    const raw = String(q||'');
    return { raw, cleaned: R2_stripNeg_(raw), appPresent: !!R2_props_().EBAY_APP_ID };
  }
  const out = {
    raw: String(q.raw||''),
    cleaned: q.cleaned ? String(q.cleaned) : R2_stripNeg_(String(q.raw||'')),
    appPresent: (q.appPresent!=null ? !!q.appPresent : !!R2_props_().EBAY_APP_ID)
  };
  return out;
}

/** Active listings: 3â€‘tier flow (API â†’ HTML â†’ RSS) with consistent fields for UI */
function apiR2_EbayActive(q) {
  const dbg = [];
  const started = Date.now();
  try {
    // Keep your original coercion + logging
    q = R2_qCoerce_(q);
    R2_dbgPush_(dbg, `Active: qRaw="${q.raw}", cleaned="${q.cleaned}", appPresent=${!!q.appPresent}`);

    const qStr = (q.cleaned || q.raw || '').trim();
    if (!qStr) return { ok:false, items:[], source:'', error:'missing query', debug:dbg, ms: Date.now()-started };

    // 1) Cloudflare Worker first (HTML/RSS at the edge)
    const w = R2_workerFetch_('active', qStr, 60, dbg);
    if (w && w.ok && (w.items || []).length) {
      R2_dbgPush_(dbg, 'Active: Workerâ†’' + (w.source || 'HTML') + ' Â· ' + w.items.length + ' items');
      return { ok:true, items:w.items, total:(w.total||null), source:(w.source || 'HTML'), note:(w.note || ''), debug:dbg, ms: Date.now()-started };
    }
    if (w && w.ok && (!w.items || !w.items.length)) {
      R2_dbgPush_(dbg, 'Active: Worker returned 0 â†’ trying API/HTML/RSS fallback');
    } else if (w && !w.ok) {
      R2_dbgPush_(dbg, 'Active: Worker error â†’ ' + (w.error || 'unknown') + ' â†’ trying API/HTML/RSS fallback');
    }

    // 2) Your existing triage (unchanged)
    // 2a) API with backoff
    const a = R2_ebayFindingActive_(q.cleaned, dbg);
    if (a && a.ok && (a.items||[]).length) {
      return { ok:true, items:a.items, source:'API', note:'', debug:dbg, ms: Date.now()-started };
    }
    if (a && a.ok && (!a.items || a.items.length === 0)) {
      R2_dbgPush_(dbg, 'Active: API returned 0 â†’ trying HTML fallback');
    }

    // 2b) HTML fallback
    const h = R2_ebayHtmlActive_(q.cleaned);
    R2_dbgPush_(dbg, `Active: HTML fallback ok=${!!h.ok}, items=${h.items ? h.items.length : 0}, err=${h.error||''}`);
    if (h && h.ok && (h.items||[]).length) {
      return { ok:true, items:h.items, source:'HTML', note:h.error||'', debug:dbg, ms: Date.now()-started };
    }
    R2_dbgPush_(dbg, h && h.error ? 'Active: HTML had error â€” trying RSS fallbackâ€¦'
                                   : 'Active: HTML had 0 items â€” trying RSS fallbackâ€¦');

    // 2c) RSS fallback
    const r = R2_ebayRssActive_(q.cleaned, dbg);
    if (r && r.ok && (r.items||[]).length) {
      return { ok:true, items:r.items, source:'RSS', note:r.note||'', debug:dbg, ms: Date.now()-started };
    }

    // Nothing worked
    const err = (h && h.error) ? h.error : (a && a.error) ? a.error : (r && r.error) ? r.error : 'No matches';
    return { ok:false, items:[], source:'', error:String(err||'No matches'), debug:dbg, ms: Date.now()-started };

  } catch (e) {
    R2_dbgPush_(dbg, 'Active: exception: ' + (e && e.stack ? e.stack : e));
    return { ok:false, items:[], source:'', error:String(e), debug:dbg, ms: Date.now()-started };
  }
}

/** Sold (90d): 3â€‘tier flow (API â†’ HTML â†’ RSS) with consistent fields for UI */
function apiR2_EbaySold(qObj) {
  const dbg = [];
  const t0 = Date.now();
  try {
    const q = (qObj && (qObj.cleaned || qObj.raw || '')).trim();
    R2_dbgPush_(dbg, 'Sold: q="' + q + '"');
    if (!q) return { ok:false, error:'missing query', items:[], debug:dbg, ms: Date.now()-t0 };

    // 1) Cloudflare Worker first (HTML/RSS at the edge)
    const w = R2_workerFetch_('sold', q, 60, dbg);
    if (w.ok && w.items && w.items.length) {
      R2_dbgPush_(dbg, 'Sold: Workerâ†’' + (w.source||'HTML') + ' Â· ' + w.items.length + ' items');
      return { ok:true, source: w.source || 'HTML', note: w.note || '', items: w.items, debug: dbg, ms: Date.now()-t0 };
    }

    // 2) Finding API (findCompletedItems) with backoff
    const api1 = sold_viaFindingApi_(q, dbg);
    if (api1.ok && api1.items.length) {
      R2_dbgPush_(dbg, 'Sold: APIâ†’ok Â· ' + api1.items.length + ' items');
      return { ok:true, source:'API', note: api1.note || '', items: api1.items, debug: dbg, ms: Date.now()-t0 };
    }

    // 3) HTML fallback
    const html1 = sold_viaHtml_(q, dbg);
    if (html1.ok && html1.items.length) {
      R2_dbgPush_(dbg, 'Sold: HTMLâ†’ok Â· ' + html1.items.length + ' items');
      return { ok:true, source:'HTML', note: html1.note || '', items: html1.items, debug: dbg, ms: Date.now()-t0 };
    }

    // 4) RSS fallback
    const rss1 = sold_viaRss_(q, dbg);
    if (rss1.ok && rss1.items.length) {
      R2_dbgPush_(dbg, 'Sold: RSSâ†’ok Â· ' + rss1.items.length + ' items');
      return { ok:true, source:'RSS', note: rss1.note || '', items: rss1.items, debug: dbg, ms: Date.now()-t0 };
    }

    const why = w.ok ? 'worker-empty' : (api1.error || html1.error || rss1.error || 'no-results');
    R2_dbgPush_(dbg, 'Sold: no results Â· ' + why);
    return { ok:true, source:'NONE', note: why, items:[], debug:dbg, ms: Date.now()-t0 };

  } catch (e) {
    R2_dbgPush_(dbg, 'Sold: error ' + (e && e.message ? e.message : e));
    return { ok:false, error:String(e), items:[], debug:dbg, ms: Date.now()-t0 };
  }

  // ---------- inline helpers (Sold) ----------

  function sold_viaFindingApi_(q, dbg) {
    const { EBAY_APP_ID, EBAY_GLOBAL_ID } = R2_props_();
    const app = EBAY_APP_ID || '';
    const gid = EBAY_GLOBAL_ID || 'EBAY-US';
    if (!app) return { ok:false, error:'missing EBAY_APP_ID', items:[] };

    const base = 'https://svcs.ebay.com/services/search/FindingService/v1';
    const params = {
      'OPERATION-NAME':'findCompletedItems',
      'SERVICE-VERSION':'1.13.0',
      'SECURITY-APPNAME': app,
      'GLOBAL-ID': gid,
      'RESPONSE-DATA-FORMAT':'JSON',
      'keywords': q,
      'paginationInput.entriesPerPage':'60',
      'itemFilter(0).name':'Condition',
      'itemFilter(0).value(0)':'Used',
      'sortOrder':'EndTimeSoonest',
      'outputSelector(0)':'SellerInfo',
      'outputSelector(1)':'PictureURLLarge',
      'outputSelector(2)':'StoreInfo'
    };

    const url = base + '?' + Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
    const attempts = 3;
    for (let i=1; i<=attempts; i++) {
      try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
        const code = res.getResponseCode();
        const body = res.getContentText() || '';
        R2_dbgPush_(dbg, `Sold: API HTTP ${code} (attempt ${i})`);
        if (code !== 200) throw new Error('http '+code);

        const data = JSON.parse(body);
        const sr = ((((data||{}).findCompletedItemsResponse||[])[0]||{}).searchResult||[])[0] || {};
        const itemsRaw = sr.item || [];

        const items = itemsRaw.map(it => {
          const priceRaw = (((((it.sellingStatus||[])[0]||{}).currentPrice||[])[0]||{}).__value__);
          const shipRaw  = (((((it.shippingInfo ||[])[0]||{}).shippingServiceCost||[])[0]||{}).__value__);
          const price    = (priceRaw === undefined || priceRaw === null || priceRaw === '') ? null : Number(priceRaw);
          const ship     = (shipRaw  === undefined || shipRaw  === null || shipRaw  === '') ? 0    : Number(shipRaw);
          const img      = ((it.pictureURLLarge||[])[0]) || ((it.galleryURL||[])[0]) || '';
          const url      = (it.viewItemURL ||[])[0] || '';
          const title    = (it.title        ||[])[0] || '';
          const state    = String((((it.sellingStatus||[])[0]||{}).sellingState||[])[0] || '');
          const sold     = /endedwithsales|sold/i.test(state);
          const total    = (price != null) ? Number((price + (ship||0)).toFixed(2)) : null;
          return { title, url, image: img, price, shipping: ship, total, sold };
        });

        return { ok:true, items, note:'API' };
      } catch (err) {
        const msg = (err && err.message) ? String(err.message) : String(err);
        R2_dbgPush_(dbg, `Sold: API fail â€“ ${msg} (attempt ${i})`);
        Utilities.sleep(400 * i);
      }
    }
    return { ok:false, error:'api-rate-limit-or-error', items:[] };
  }

  function sold_viaHtml_(q, dbg) {
    const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_Sold=1&LH_Complete=1&_ipg=60';
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
      const code = res.getResponseCode();
      const html = res.getContentText() || '';
      R2_dbgPush_(dbg, 'Sold: HTML HTTP ' + code);
      if (code !== 200) return { ok:false, error:'html http '+code, items:[] };
      if (R2_isEbayBotWallHtml_(html)) return { ok:false, error:'html bot-wall', items:[] };
      const items = R2_parseEbayList_(html, /*sold=*/true);
      return { ok:true, items, note:'HTML' };
    } catch (e) {
      return { ok:false, error:String(e), items:[] };
    }
  }

  function sold_viaRss_(q, dbg) {
    const url = 'https://www.ebay.com/sch/i.html?_nkw=' + encodeURIComponent(q) + '&LH_Sold=1&LH_Complete=1&_ipg=60&_rss=1';
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions:true, followRedirects:true });
      const code = res.getResponseCode();
      const rss = res.getContentText() || '';
      R2_dbgPush_(dbg, 'Sold: RSS HTTP ' + code);
      if (code !== 200) return { ok:false, error:'rss http '+code, items:[] };
      const items = R2_parseEbayRss_(rss, /*sold=*/true);   // <-- use your R2_ helper
      return { ok:true, items, note:'RSS' };
    } catch (e) {
      return { ok:false, error:String(e), items:[] };
    }
  }
}

function R2_dumpToDrive_(name, content){
  try{
    const folderName = 'ResellPro-Debug';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const ts = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd-HHmmss'Z'");
    const file = folder.createFile(name.replace(/[^\w.-]+/g,'_')+'-'+ts+'.txt', String(content||''), 'text/plain');
    return file.getUrl();
  }catch(e){ return ''; }
}

function R2_dbgHttp_(dbg, label, res, body, extraMsg){
  try{
    const code = res ? res.getResponseCode() : '(no response)';
    const url  = res && res.getHeaders ? (res.getHeaders()['X-Resolved-Url'] || '') : '';
    const peek = (String(body||'').slice(0,200) || '').replace(/\s+/g,' ').trim();
    R2_dbgPush_(dbg, `${label}: code=${code}, url=${R2_peek_(url,120)}, peek="${R2_peek_(peek,140)}"${extraMsg?(' â€” '+extraMsg):''}`);
  }catch(_){}
}

function callEbayFindingWithBackoff(paramsObj){
  const params = Object.assign({
    'OPERATION-NAME': 'findItemsByKeywords',
    'SERVICE-VERSION': '1.13.0',
    'SECURITY-APPNAME': EBAY_APP_ID,
    'GLOBAL-ID': 'EBAY-US',
    'RESPONSE-DATA-FORMAT': 'JSON',
    'paginationInput.entriesPerPage': '60',
    'itemFilter(0).name': 'ListingType',
    'itemFilter(0).value(0)': 'FixedPrice',
    'itemFilter(1).name': 'HideDuplicateItems',
    'itemFilter(1).value(0)': 'true',
    'outputSelector(0)': 'PictureURLLarge',
    'outputSelector(1)': 'StoreInfo',
  }, paramsObj || {});

  let delay = BACKOFF_START_MS;
  let lastErr = null;

  for (let attempt = 0; attempt < MAX_FINDING_RETRIES; attempt++){
    try{
      const url = EBAY_FINDING_ENDPOINT + '?' + Object.keys(params).map(k =>
        encodeURIComponent(k)+'='+encodeURIComponent(params[k])
      ).join('&');

      const resp = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true});
      const code = resp.getResponseCode();
      const body = resp.getContentText();

      if (code !== 200){
        lastErr = new Error(`Finding HTTP ${code}`);
      } else {
        const j = JSON.parse(body);
        const ack = j && j.findItemsByKeywordsResponse && j.findItemsByKeywordsResponse[0]?.ack?.[0];
        const err  = j && j.findItemsByKeywordsResponse && j.findItemsByKeywordsResponse[0]?.errorMessage?.[0]?.error?.[0];
        const errId = err?.errorId?.[0];

        // success
        if (ack === 'Success' || !err){
          return { ok:true, json:j };
        }

        // rate limited â€” retry with backoff
        if (String(errId) === '10001'){
          lastErr = new Error('eBay 10001 RateLimiter');
          if (attempt < MAX_FINDING_RETRIES-1){
            sleepMs(jitter(delay)); delay = Math.round(delay * BACKOFF_FACTOR);
            continue;
          }
        }else{
          // other API error â€” bail to fallback
          lastErr = new Error(`eBay Finding error ${errId||'?'}: ${err?.message?.[0]||''}`);
        }
      }
    }catch(e){
      lastErr = e;
    }
    // if we got here without continuing, weâ€™ll fall through to fallback
  }
  return { ok:false, error:lastErr };
}

function fetchEbayRss(kind, query){
  // kind: 'active' | 'sold'
  const base = 'https://www.ebay.com/sch/i.html';
  const qs = {
    _nkw: query,
    _ipg: '60',
    _rss: '1'
  };
  if (kind === 'active'){ qs.LH_BIN = '1'; }          // Buy It Now
  if (kind === 'sold')  { qs.LH_Sold = '1'; qs.LH_Complete = '1'; }

  const url = base + '?' + Object.keys(qs).map(k=>k+'='+encodeURIComponent(qs[k])).join('&');

  const resp = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true});
  if (resp.getResponseCode() !== 200) return { ok:false, items:[], error:'RSS HTTP '+resp.getResponseCode() };

  const xml = XmlService.parse(resp.getContentText());
  const ch  = xml.getRootElement().getChild('channel');
  if (!ch) return { ok:false, items:[] };

  const items = [];
  (ch.getChildren('item')||[]).forEach(it=>{
    const title = it.getChildText('title') || '';
    const link  = it.getChildText('link') || '';
    // crude price pull from title (RSS doesnâ€™t expose price cleanly)
    const m = title.match(/\$([\d,]+(?:\.\d{2})?)/);
    const total = m ? Number(m[1].replace(/,/g,'')) : null;
    items.push({ title, url: link, total, sold: (kind==='sold') });
  });

  return { ok:true, items };
}

/* ------------------- normalizers & helpers (unchanged) ------------- */

function R2_price_(obj){ if(!obj) return null; const x = (obj[0]||{}); const v = Number(x.__value__); return isFinite(v)?v:null; }
function R2_text_(a){ return Array.isArray(a) && a.length ? String(a[0]) : ''; }

function R2_normActive_(it){
  const price = R2_price_(((it.sellingStatus||[])[0]||{}).currentPrice);
  let ship = null; try { ship = R2_price_(((it.shippingInfo||[])[0]||{}).shippingServiceCost); } catch(_) {}
  ship = (ship==null) ? 0 : ship;
  const url = R2_text_(it.viewItemURL);
  const img = R2_text_(it.galleryPlusPictureURL) || R2_text_(it.galleryURL);
  return { id:R2_text_(it.itemId), title:R2_text_(it.title), price:price, shipping:ship, total:(price!=null ? +(price + (ship||0)).toFixed(2) : null), url:url, image:img };
}
function R2_normCompleted_(it){
  const price = R2_price_(((it.sellingStatus||[])[0]||{}).currentPrice);
  let ship = null; try { ship = R2_price_(((it.shippingInfo||[])[0]||{}).shippingServiceCost); } catch(_) {}
  ship = (ship==null) ? 0 : ship;
  const url = R2_text_(it.viewItemURL);
  const img = R2_text_(it.galleryPlusPictureURL) || R2_text_(it.galleryURL);
  const state = R2_text_(((it.sellingStatus||[])[0]||{}).sellingState);
  const sold = String(state||'').toLowerCase() === 'endedwithsales';
  return { id:R2_text_(it.itemId), title:R2_text_(it.title), price:price, shipping:ship, total:(price!=null ? +(price + (ship||0)).toFixed(2) : null), url:url, image:img, sold:sold };
}

function toQS_(obj){
  const parts=[]; Object.keys(obj).forEach(k=>{ if (obj[k]==null || obj[k]==='') return; parts.push(encodeURIComponent(k)+'='+encodeURIComponent(String(obj[k]))); });
  return parts.join('&');
}

/* ------------------------------------------------------------------ */
/* PRICING / INTAKE passthrough (unchanged)                            */
/* ------------------------------------------------------------------ */

function apiR2_ComputePricing(input){ return computePricing(input); }

function computePricing(input){
  const n=x=> (x==null||x==='')?null:Number(x);
  const aC=n(input.activeCount), aL=n(input.activeLow), aM=n(input.activeMid), aH=n(input.activeHigh);
  const sC=n(input.soldCount),   sL=n(input.soldLow),  sM=n(input.soldMid),  sH=n(input.soldHigh);
  if(!aC || aC<0) return { ok:false, error:'Enter Total Active count.' };
  const pct=Math.round(((sC||0)/aC)*100);
  let band='poor', color='warn';
  if(pct>=80){band='excellent'; color='ok';}
  else if(pct>=70){band='great'; color='ok';}
  else if(pct>=60){band='good'; color='ok';}
  else if(pct>=50){band='fair';}
  const avg=arr=>{const xs=(arr||[]).filter(v=>typeof v==='number'&&isFinite(v));return xs.length?Math.round((xs.reduce((a,b)=>a+b,0)/xs.length)*100)/100:null;};
  let rec=null;
  switch(band){
    case 'poor':       rec=avg([sL, aM]); break;
    case 'fair':       rec=avg([sM, aM]); break;
    case 'good':       rec=avg([sM, aH]); break;
    case 'great':      rec=Math.round(avg([sH, aH])*0.95*100)/100; break;
    case 'excellent': { const base=avg([sH,aH]); const up10=base!=null?base*1.10:null; const cap=(typeof aH==='number'&&isFinite(aH))?(aH*0.90):null; rec = (up10!=null && cap!=null) ? Math.round(Math.min(up10,cap)*100)/100 : (up10!=null?Math.round(up10*100)/100:null); break; }
  }
  return { ok:true, sellThroughPct:pct, band, bandColor:color, recommended: rec };
}


function apiR2_SendToIntake(d){
  const r = saveIntakeDraft(d) || { ok:true };
  r.link = getIntakeUrl_();
  return r;
}
function getIntakeUrl_(){
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Settings') || ss.getSheetByName('Config');
    if (!sh) return '';
    const rows = sh.getDataRange().getValues();
    for (let i=0;i<rows.length;i++){
      const k = String(rows[i][0]||'').trim().toLowerCase();
      if (k==='intake_url' || k==='intake web app url' || k==='intake link'){
        return String(rows[i][1]||'').trim();
      }
    }
  } catch(_) {}
  return '';
}
function _n(x){ if(x==null||x==='') return ''; const n=Number(x); return Number.isFinite(n)?n:''; }

function saveIntakeDraft(d){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('Intake Draft')||ss.insertSheet('Intake Draft');
  if(sh.getLastRow()===0){
    sh.appendRow(['Timestamp','ShortDesc','LongDesc','ActiveCount','ActiveLow','ActiveMid','ActiveHigh','SoldCount','SoldLow','SoldMid','SoldHigh','SellThroughPct','RecommendedPrice','AltPrice','ImageFileId']);
  }
  let fileId=''; const { IMG_FOLDER_ID }=props_();
  if(d.imageB64 && IMG_FOLDER_ID){ try{ const blob=Utilities.newBlob(Utilities.base64Decode(d.imageB64),'image/jpeg','intake-'+Date.now()+'.jpg'); const folder=DriveApp.getFolderById(IMG_FOLDER_ID); fileId=folder.createFile(blob).getId(); }catch(_){ } }
  sh.appendRow([
    new Date(), d.shortDesc||'', d.longDesc||'',
    _n(d.activeCount), _n(d.activeLow), _n(d.activeMid), _n(d.activeHigh),
    _n(d.soldCount), _n(d.soldLow), _n(d.soldMid), _n(d.soldHigh),
    (d.sellThrough||'').replace('%',''), _n(d.recPrice), _n(d.altPrice), fileId
  ]);
  return { ok:true };
}

function apiIntake_GetLatestDraft(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('Intake Draft');
  if(!sh || sh.getLastRow()<2) return { ok:false, reason:'empty' };
  const header=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], last=sh.getRange(sh.getLastRow(),1,1,sh.getLastColumn()).getValues()[0];
  const H=n=>header.indexOf(n);
  return {
    ok:true,
    shortDesc:String(last[H('ShortDesc')]||''),
    longDesc:String(last[H('LongDesc')]||''),
    recPrice:last[H('RecommendedPrice')]||'',
    altPrice:last[H('AltPrice')]||'',
    imageFileId: String(last[H('ImageFileId')]||'')
  };
}

