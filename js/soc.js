/* ============================================================================
 * RCT OpsDesk — SOC Risky Sign-in Investigation (js/soc.js)
 * Ported from soc-dashboard v12.3 (standalone) — all analysis logic unchanged.
 * Uses the dashboard's vendored libs/xlsx.full.min.js and libs/chart.umd.min.js
 * instead of inlined copies. Fully offline.
 * Init: QBR.socInit() (idempotent) — called when the section is first shown.
 * ========================================================================== */
(function () {
"use strict";
const QBR = (window.QBR = window.QBR || {});

// ============================================================
// State
// ============================================================
const state = {
  riskyData: null,
  riskyFilename: null,
  allData: null,
  allFilename: null,
  authData: null,
  authFilename: null,
  authMap: null,        // Map<normalizedUPN, authRecord>
  noniData: null,       // v12.3 Source 04 — non-interactive sign-ins (column-trimmed)
  noniFilename: null,
  noniMeta: null,       // v12.3 {totalWithUser, filtered} from ingest
  noniAuthMap: null,    // v12.3 Source 05 — Map<Request ID, {satisfiedByToken, method, detail}>
  noniAuthFilename: null,
  report: null,
};

// ============================================================
// Utility
// ============================================================
function $(id) { return document.getElementById(id); }

function showToast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._tm);
  t._tm = setTimeout(() => t.classList.remove('show'), 4500);
}

function fmtDate(iso) {
  if (!iso) return '—';
  let d;
  if (iso instanceof Date) d = iso;
  else d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso);
  return d.toISOString().replace('T', ' ').replace(/\.\d+Z$/, 'Z');
}

// v9: observation-window label helpers — replace the old hardcoded "7-day".
// The window is computed in analyze() from the real data span and stored on agg.
function windowLabel(agg) {
  if (!agg || !agg.windowHas) return 'available observation window';
  return `${agg.windowDays}-day observation window`;
}
function windowRange(agg) {
  if (!agg || !agg.windowHas) return '';
  const s = fmtDate(agg.windowStart).slice(0, 10);
  const e = fmtDate(agg.windowEnd).slice(0, 10);
  return ` (${s} → ${e})`;
}

// ============================================================
// v10 risk scoring (heuristic, analyst-tunable) — bound to risky-context.
// Every term traces to a present data field; terms whose source column is
// absent contribute 0 and set partial=true so a low score is never misread
// as "safe". Clean-IP success contributes 0 (invariant 5).
// ============================================================
const RISK_TIERS = [
  { min: 90, tier: 'Critical', cls: 'crit',   priority: 'P1' },
  { min: 70, tier: 'High',     cls: 'high',   priority: 'P2' },
  { min: 40, tier: 'Medium',   cls: 'medium', priority: 'P3' },
  { min: 0,  tier: 'Low',      cls: 'low',    priority: 'P4' },
];
function tierFromScore(score) {
  return RISK_TIERS.find(t => score >= t.min) || RISK_TIERS[RISK_TIERS.length - 1];
}
function computeRiskScore(inv, opts) {
  const { authAvailable, caAvailable } = opts;
  const contribs = [];
  let partial = false;

  // +50 — success from an EXACT risky IP, post-detection (invariant 5).
  const riskyIpSuccess = inv.postRiskySuccess.some(s => s.ip && inv.riskyIpSet.has(s.ip));
  if (riskyIpSuccess) contribs.push({ label: 'Success from risky IP (post-detection)', pts: 50 });
  else if (inv.bgCompromise) contribs.push({ label: 'Success from risky IP via background token (post-detection)', pts: 50 }); // v12.3

  // +20 password-only / +10 weak MFA — require the auth-methods source.
  if (authAvailable) {
    if (inv.authPosture.tier === 'none')      contribs.push({ label: 'Password-only posture', pts: 20 });
    else if (inv.authPosture.tier === 'weak') contribs.push({ label: 'Weak MFA (SMS/Voice/Email)', pts: 10 });
  } else { partial = true; }

  // +10 — three or more risky detections.
  if ((inv.riskyEvents?.length || 0) >= 3) contribs.push({ label: '≥3 risky detections', pts: 10 });

  // +20 — Conditional Access failure within risky context (requires CA column).
  if (caAvailable) {
    const caFail = inv.riskyContextSignins.some(s => s.ca && /fail/i.test(String(s.ca)));
    if (caFail) contribs.push({ label: 'Conditional Access failure (risky context)', pts: 20 });
  } else { partial = true; }

  // +10 — repeat offender: risky detections spanning ≥2 distinct UTC days.
  const days = new Set((inv.riskyEvents || [])
    .map(r => r.date && Math.floor(r.date.getTime() / 86400000))
    .filter(v => v != null));
  if (days.size >= 2) contribs.push({ label: 'Repeat offender (≥2 days)', pts: 10 });

  const raw = contribs.reduce((s, c) => s + c.pts, 0);
  const score = Math.max(0, Math.min(100, raw));
  const t = tierFromScore(score);
  return { score, tier: t.tier, cls: t.cls, priority: t.priority, contribs, partial };
}

// ============================================================
// v12 — MITRE ATT&CK heuristic mapping (INFERRED, not detected).
// Conservative: a technique is emitted only when supporting rows exist.
// ============================================================
const MITRE_SEV = { high: 'crit', medium: 'high', low: 'medium' };
function computeMitre(inv) {
  const out = [];
  const fails = inv.riskyContextSignins.filter(s => /failure/i.test(String(s.status || ''))).length;
  const interrupts = inv.riskyContextSignins.filter(s => /interrupt/i.test(String(s.status || ''))).length;
  const validAcct = inv.postRiskySuccess.some(s => s.ip && inv.riskyIpSet.has(s.ip));
  if (validAcct) out.push({ id: 'T1078', name: 'Valid Accounts', tactic: 'Initial Access / Persistence',
    sev: 'high', evidence: inv.postRiskySuccess.length, note: 'Successful authentication from a risky IP after detection' });
  if (inv.bgCompromise) out.push({ id: 'T1550', name: 'Use Alternate Authentication Material', tactic: 'Defense Evasion / Lateral Movement',
    sev: 'high', evidence: inv.bgPostRiskySuccess.filter(s => s.ip && inv.riskyIpSet.has(s.ip)).length,
    note: 'Background token sign-in from a risky IP after detection (possible token replay)' });   // v12.3
  if (fails >= 3) out.push({ id: 'T1110', name: 'Brute Force', tactic: 'Credential Access',
    sev: 'medium', evidence: fails, note: fails + ' failed authentications in risky context' });
  if (interrupts >= 3) out.push({ id: 'T1621', name: 'MFA Request Generation', tactic: 'Credential Access',
    sev: 'medium', evidence: interrupts, note: interrupts + ' MFA challenges/interrupts in risky context' });
  return out;
}

// v12 — offline country resolution (bundled ISO-3166 alpha-2 names, no geocoding).
const COUNTRY_NAMES = {
  AD:'Andorra',AE:'United Arab Emirates',AF:'Afghanistan',AG:'Antigua & Barbuda',AI:'Anguilla',AL:'Albania',AM:'Armenia',AO:'Angola',AR:'Argentina',AT:'Austria',AU:'Australia',AW:'Aruba',AZ:'Azerbaijan',
  BA:'Bosnia & Herzegovina',BB:'Barbados',BD:'Bangladesh',BE:'Belgium',BF:'Burkina Faso',BG:'Bulgaria',BH:'Bahrain',BI:'Burundi',BJ:'Benin',BM:'Bermuda',BN:'Brunei',BO:'Bolivia',BR:'Brazil',BS:'Bahamas',BT:'Bhutan',BW:'Botswana',BY:'Belarus',BZ:'Belize',
  CA:'Canada',CD:'DR Congo',CF:'Central African Rep.',CG:'Congo',CH:'Switzerland',CI:"Côte d'Ivoire",CL:'Chile',CM:'Cameroon',CN:'China',CO:'Colombia',CR:'Costa Rica',CU:'Cuba',CV:'Cabo Verde',CY:'Cyprus',CZ:'Czechia',
  DE:'Germany',DJ:'Djibouti',DK:'Denmark',DM:'Dominica',DO:'Dominican Rep.',DZ:'Algeria',
  EC:'Ecuador',EE:'Estonia',EG:'Egypt',ER:'Eritrea',ES:'Spain',ET:'Ethiopia',
  FI:'Finland',FJ:'Fiji',FM:'Micronesia',FO:'Faroe Islands',FR:'France',
  GA:'Gabon',GB:'United Kingdom',GD:'Grenada',GE:'Georgia',GH:'Ghana',GI:'Gibraltar',GL:'Greenland',GM:'Gambia',GN:'Guinea',GQ:'Eq. Guinea',GR:'Greece',GT:'Guatemala',GU:'Guam',GW:'Guinea-Bissau',GY:'Guyana',
  HK:'Hong Kong',HN:'Honduras',HR:'Croatia',HT:'Haiti',HU:'Hungary',
  ID:'Indonesia',IE:'Ireland',IL:'Israel',IM:'Isle of Man',IN:'India',IQ:'Iraq',IR:'Iran',IS:'Iceland',IT:'Italy',
  JE:'Jersey',JM:'Jamaica',JO:'Jordan',JP:'Japan',
  KE:'Kenya',KG:'Kyrgyzstan',KH:'Cambodia',KI:'Kiribati',KM:'Comoros',KN:'St Kitts & Nevis',KP:'North Korea',KR:'South Korea',KW:'Kuwait',KY:'Cayman Islands',KZ:'Kazakhstan',
  LA:'Laos',LB:'Lebanon',LC:'St Lucia',LI:'Liechtenstein',LK:'Sri Lanka',LR:'Liberia',LS:'Lesotho',LT:'Lithuania',LU:'Luxembourg',LV:'Latvia',LY:'Libya',
  MA:'Morocco',MC:'Monaco',MD:'Moldova',ME:'Montenegro',MG:'Madagascar',MH:'Marshall Islands',MK:'North Macedonia',ML:'Mali',MM:'Myanmar',MN:'Mongolia',MO:'Macau',MR:'Mauritania',MT:'Malta',MU:'Mauritius',MV:'Maldives',MW:'Malawi',MX:'Mexico',MY:'Malaysia',MZ:'Mozambique',
  NA:'Namibia',NC:'New Caledonia',NE:'Niger',NG:'Nigeria',NI:'Nicaragua',NL:'Netherlands',NO:'Norway',NP:'Nepal',NZ:'New Zealand',
  OM:'Oman',
  PA:'Panama',PE:'Peru',PF:'French Polynesia',PG:'Papua New Guinea',PH:'Philippines',PK:'Pakistan',PL:'Poland',PR:'Puerto Rico',PS:'Palestine',PT:'Portugal',PW:'Palau',PY:'Paraguay',
  QA:'Qatar',
  RE:'Réunion',RO:'Romania',RS:'Serbia',RU:'Russia',RW:'Rwanda',
  SA:'Saudi Arabia',SB:'Solomon Islands',SC:'Seychelles',SD:'Sudan',SE:'Sweden',SG:'Singapore',SI:'Slovenia',SK:'Slovakia',SL:'Sierra Leone',SM:'San Marino',SN:'Senegal',SO:'Somalia',SR:'Suriname',SS:'South Sudan',ST:'São Tomé & Príncipe',SV:'El Salvador',SY:'Syria',SZ:'Eswatini',
  TC:'Turks & Caicos',TD:'Chad',TG:'Togo',TH:'Thailand',TJ:'Tajikistan',TL:'Timor-Leste',TM:'Turkmenistan',TN:'Tunisia',TO:'Tonga',TR:'Türkiye',TT:'Trinidad & Tobago',TV:'Tuvalu',TW:'Taiwan',TZ:'Tanzania',
  UA:'Ukraine',UG:'Uganda',US:'United States',UY:'Uruguay',UZ:'Uzbekistan',
  VA:'Vatican City',VC:'St Vincent',VE:'Venezuela',VG:'British Virgin Is.',VI:'US Virgin Is.',VN:'Vietnam',VU:'Vanuatu',
  WS:'Samoa',XK:'Kosovo',YE:'Yemen',YT:'Mayotte',ZA:'South Africa',ZM:'Zambia',ZW:'Zimbabwe',UK:'United Kingdom',
};
function parseCountry(loc) {
  if (!loc) return { code: null, name: 'Unmapped' };
  const parts = String(loc).split(',').map(s => s.trim()).filter(Boolean);
  if (!parts.length) return { code: null, name: 'Unmapped' };
  const last = parts[parts.length - 1];
  const up = last.toUpperCase();
  if (up.length === 2 && COUNTRY_NAMES[up]) return { code: up, name: COUNTRY_NAMES[up] };
  return { code: null, name: last }; // full country name given, or unrecognised token
}

function normalizeKey(k) {
  return String(k || '').toLowerCase().trim().replace(/\s+/g, ' ');
}

function getField(row, ...candidates) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const target = normalizeKey(cand);
    const hit = keys.find(k => normalizeKey(k) === target);
    if (hit !== undefined && row[hit] !== undefined && row[hit] !== null && row[hit] !== '')
      return row[hit];
  }
  return null;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function shortIp(ip) {
  if (!ip) return '—';
  const s = String(ip);
  if (s.length > 36) return s.slice(0, 34) + '…';
  return s;
}

// ============================================================
// File ingest
// ============================================================
function bindDrop(dropEl, inputEl, slot) {
  dropEl.addEventListener('click', () => inputEl.click());
  dropEl.addEventListener('dragover', e => { e.preventDefault(); dropEl.classList.add('dragging'); });
  dropEl.addEventListener('dragleave', () => dropEl.classList.remove('dragging'));
  dropEl.addEventListener('drop', e => {
    e.preventDefault();
    dropEl.classList.remove('dragging');
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f, slot);
  });
  inputEl.addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) handleFile(f, slot);
  });
}

// ---- Off-main-thread workbook parsing --------------------------------------
// v8: XLSX.read + sheet_to_json are heavy on large sign-in logs (10k–50k+ rows)
// and previously ran on the UI thread, freezing the page with no feedback.
// We now parse inside a Web Worker so the main thread stays responsive, and we
// paint a "reading…" state the instant a file is selected. If the worker can't
// be created (older browser / blocked), we fall back to a yielded main-thread
// parse so ingest still works — just without the anti-freeze guarantee.
// v12: the worker loads xlsx from the INLINED library via a same-origin blob URL
// (built once from the #xlsxlib script) so parsing works with zero network — even
// air-gapped. Falls back to the CDN only if the inlined script isn't found.
// Dashboard merge (v1.32.0): NO CDN fallback — the app must stay offline/no-CDN. When the
// vendored source can't be fetched (file://), parsing uses the main-thread fallback below.
const XLSX_CDN = null;
let _xlsxLibUrl = null;
function xlsxLibUrl() {
  if (_xlsxLibUrl) return _xlsxLibUrl;
  try {
    if (window.__XLSX_SRC__ && window.__XLSX_SRC__.length > 1000) {
      _xlsxLibUrl = URL.createObjectURL(new Blob([window.__XLSX_SRC__], { type: 'application/javascript' }));
      return _xlsxLibUrl;
    }
  } catch (e) {}
  return null; // no inline source -> caller falls back to main-thread parsing
}
function workerSrc(libUrl) {
  // v12.3 FIX: from file:// the page origin is "null", and a worker cannot
  // importScripts() a second blob:null URL — so since v12 the worker silently
  // failed and every ingest fell back to the (freezing) main thread. The inlined
  // xlsx source is now prepended to the worker's OWN script, so no import is needed.
  const inlineLib = (window.__XLSX_SRC__ && window.__XLSX_SRC__.length > 1000) ? window.__XLSX_SRC__ : null;
  const head = inlineLib ? inlineLib + "\n;\n" : "importScripts(" + JSON.stringify(libUrl) + ");\n";
  return head + [
    "self.onmessage = function(ev){",
    "  try {",
    "    var data = new Uint8Array(ev.data.buffer);",
    "    var wb = XLSX.read(data, { type:'array', cellDates:true });",
    "    var ws = wb.Sheets[wb.SheetNames[0]];",
    "    var json = XLSX.utils.sheet_to_json(ws, { defval:null, raw:false });",
    "    var meta = null;",
    "    var keep = ev.data.keepCols, users = ev.data.keepUsers;",
    "    if (keep && keep.length) {",
    "      var set = {}; for (var k = 0; k < keep.length; k++) set[keep[k]] = 1;",
    "      var uset = null; if (users) { uset = {}; for (var q = 0; q < users.length; q++) uset[users[q]] = 1; }",
    "      var out = [], withUser = 0;",
    "      for (var r = 0; r < json.length; r++) {",
    "        var row = json[r], o = {}, any = false, vu = null, vn = null, vd = null;",
    "        for (var key in row) {",
    "          var nk = String(key).toLowerCase().trim().replace(/\\s+/g, ' ');",
    "          var val = row[key];",
    "          if (nk === 'user' && val) vu = val; else if (nk === 'name' && val) vn = val; else if (nk === 'user display name' && val) vd = val;",
    "          if (set[nk]) { o[key] = val; if (val !== null && val !== '') any = true; }",
    "        }",
    "        var u = vu || vn || vd;",
    "        if (u) withUser++;",
    "        if (uset && (!u || !uset[String(u).toLowerCase().trim()])) continue;",
    "        if (any) out.push(o);",
    "      }",
    "      json = out; meta = { totalWithUser: withUser, filtered: !!uset };",
    "    }",
    "    self.postMessage({ ok:true, json:json, meta:meta });",
    "  } catch (e) {",
    "    self.postMessage({ ok:false, error:(e && e.message) || String(e) });",
    "  }",
    "};"
  ].join("\n");
}

function parseWorkbookInWorker(buffer, opts) {
  return new Promise((resolve, reject) => {
    let worker;
    try {
      const lib = xlsxLibUrl();
      if (!lib) { reject(new Error('xlsx source unavailable for the worker (offline/file://) - parsing on main thread')); return; }
      const url = URL.createObjectURL(new Blob([workerSrc(lib)], { type: 'application/javascript' }));
      worker = new Worker(url);
      URL.revokeObjectURL(url);
    } catch (e) { reject(e); return; }
    worker.onmessage = ev => {
      worker.terminate();
      if (ev.data && ev.data.ok) {
        const j = ev.data.json;
        if (ev.data.meta) Object.defineProperty(j, '__meta', { value: ev.data.meta, enumerable: false });
        resolve(j);
      }
      else reject(new Error((ev.data && ev.data.error) || 'worker parse failed'));
    };
    worker.onerror = err => {
      worker.terminate();
      reject(new Error((err && err.message) || 'worker error'));
    };
    // NOTE: do NOT transfer the buffer — a copy is cheap and keeps the original
    // intact so the main-thread fallback below can still use it if the worker dies.
    worker.postMessage({ buffer, keepCols: (opts && opts.keepCols) || null, keepUsers: (opts && opts.keepUsers) || null });
  });
}

// v12.3 — Source 04 column whitelist (normalised names). Non-interactive exports
// carry ~56 columns and can hit 100k rows; keeping only what the report reads
// cuts memory sharply. Applied ONLY to the non-interactive slot.
const BG_KEEP_COLS = [
  'date (utc)','date','datetime','request id','user','name','user display name',
  'username','user principal name','application','resource','ip address','ipaddress',
  'location','status','sign-in error code','failure reason','client app','device id',
  'browser','operating system','compliant','managed','join type',
  'multifactor authentication result','multifactor authentication auth method',
  'multifactor authentication auth detail','authentication requirement',
  'incoming token type','autonomous system number','conditional access',
];
function trimColumns(json, keep, keepUsers) {
  const set = new Set(keep);
  const uset = keepUsers ? new Set(keepUsers) : null;
  const out = []; let withUser = 0;
  for (const row of json) {
    const o = {}; let any = false;
    for (const key in row) {
      if (set.has(normalizeKey(key))) { o[key] = row[key]; if (row[key] !== null && row[key] !== '') any = true; }
    }
    const u = getField(row, 'User', 'Name', 'User display name');
    if (u) withUser++;
    if (uset && (!u || !uset.has(String(u).toLowerCase().trim()))) continue;
    if (any) out.push(o);
  }
  Object.defineProperty(out, '__meta', { value: { totalWithUser: withUser, filtered: !!uset }, enumerable: false });
  return out;
}
// Display-name keys of "At risk" users in Source 01 (same rule as analyze()).
function atRiskUserKeys() {
  if (!state.riskyData) return null;
  const keys = new Set();
  for (const r of state.riskyData) {
    const rs = getField(r, 'Risk state', 'risk state', 'RiskState');
    if (!rs || !String(rs).toLowerCase().includes('at risk')) continue;
    const u = getField(r, 'User', 'Name', 'User display name');
    if (u) keys.add(String(u).toLowerCase().trim());
  }
  return [...keys];
}

function parseWorkbookMainThread(buffer, opts) {
  return new Promise((resolve, reject) => {
    // Yield one paint so the "reading…" spinner renders before we block.
    setTimeout(() => {
      try {
        const data = new Uint8Array(buffer);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
        resolve(opts && opts.keepCols ? trimColumns(json, opts.keepCols, opts.keepUsers) : json);
      } catch (e) { reject(e); }
    }, 30);
  });
}

function showReadingState(slot, filename) {
  const pill = $('pill-' + slot);
  pill.classList.remove('hidden', 'done');
  pill.classList.add('reading');
  pill.querySelector('.filename').textContent = filename;
  pill.querySelector('.rows').innerHTML = '<span class="spin"></span>reading…';
  $('card-' + slot).classList.remove('loaded');
}

function failIngest(slot, filename, msg) {
  const pill = $('pill-' + slot);
  pill.classList.add('hidden');
  pill.classList.remove('reading', 'done');
  $('card-' + slot).classList.remove('loaded');
  showToast('Failed to parse ' + filename + ': ' + msg);
}

function applyParsed(slot, filename, json) {
  if (slot === 'noni' && json && !json.length && json.__meta && json.__meta.filtered) {
    failIngest(slot, filename, `none of ${json.__meta.totalWithUser.toLocaleString()} rows belong to the risky users in Source 01`); return;
  }
  if (!json || !json.length) { failIngest(slot, filename, 'Sheet is empty'); return; }
  if (slot === 'risky' && state.noniMeta && state.noniMeta.filtered) {
    // Source 04 was pre-filtered for the PREVIOUS risky list — it would now be stale.
    state.noniData = null; state.noniFilename = null; state.noniMeta = null;
    $('pill-noni').classList.add('hidden'); $('card-noni').classList.remove('loaded'); $('file-noni').value = '';
    showToast('Source 04 was filtered for the previous risky list — please re-load the non-interactive file.');
  }
  if (slot === 'risky') {
    state.riskyData = json;
    state.riskyFilename = filename;
  } else if (slot === 'all') {
    state.allData = json;
    state.allFilename = filename;
  } else if (slot === 'auth') {
    state.authData = json;
    state.authFilename = filename;
    state.authMap = buildAuthMap(json);
  } else if (slot === 'noni') {                       // v12.3
    state.noniData = json;
    state.noniFilename = filename;
    state.noniMeta = json.__meta || null;
  } else if (slot === 'noniad') {                     // v12.3
    state.noniAuthMap = buildBgAuthMap(json);
    state.noniAuthFilename = filename;
    if (!state.noniAuthMap.size) { state.noniAuthMap = null; failIngest(slot, filename, 'no Request ID rows found'); return; }
  }
  const pill = $('pill-' + slot);
  pill.classList.remove('hidden', 'reading');
  pill.classList.add('done');
  pill.querySelector('.filename').textContent = filename;
  pill.querySelector('.rows').textContent =
    (slot === 'auth')
      ? (json.length + ' users · ' + state.authMap.size + ' indexed')
      : (slot === 'noniad')
      ? (state.noniAuthMap.size.toLocaleString() + ' request IDs indexed')
      : (slot === 'noni' && state.noniMeta && state.noniMeta.filtered)
      ? (json.length.toLocaleString() + ' risky-user rows kept (of ' + state.noniMeta.totalWithUser.toLocaleString() + ')')
      : (slot === 'noni')
      ? (json.length.toLocaleString() + ' rows · load Source 01 first to pre-filter')
      : (json.length.toLocaleString() + ' rows');
  $('card-' + slot).classList.add('loaded');
  updateAnalyzeReady();
}

function handleFile(file, slot) {
  const reader = new FileReader();
  reader.onload = e => {
    const buffer = e.target.result;
    showReadingState(slot, file.name);            // immediate feedback (v8)
    const opts = slot === 'noni' ? { keepCols: BG_KEEP_COLS, keepUsers: atRiskUserKeys() } : undefined;  // v12.3
    parseWorkbookInWorker(buffer, opts)
      .then(json => applyParsed(slot, file.name, json))
      .catch(() =>                                 // worker unavailable → fallback
        parseWorkbookMainThread(buffer, opts)
          .then(json => applyParsed(slot, file.name, json))
          .catch(err => failIngest(slot, file.name, err.message))
      );
  };
  reader.onerror = () => { failIngest(slot, file.name, 'could not read file'); };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// Auth method enrichment helpers
// ============================================================

// ============================================================
// v12.3 — Background (non-interactive) layer helpers
// ============================================================
// Source 05: Map<Request ID, {satisfiedByToken, method, detail}>. A request can
// have several auth-detail rows; it counts as satisfied-by-token if ANY row says so.
function buildBgAuthMap(rows) {
  const m = new Map();
  for (const r of rows) {
    const id = getField(r, 'Request ID');
    if (!id) continue;
    const key = String(id).trim();
    const method = String(getField(r, 'Authentication method') || '');
    const detail = String(getField(r, 'Result detail') || '');
    const sat = /previously satisfied/i.test(method) || /satisfied by claim in the token/i.test(detail);
    const prev = m.get(key);
    if (!prev) m.set(key, { satisfiedByToken: sat, method, detail });
    else if (sat && !prev.satisfiedByToken) m.set(key, { satisfiedByToken: true, method, detail });
  }
  return m;
}

// Map one Source 04 row to the same sign-in shape the engine already uses,
// plus background-only forensic fields. signInType tags it as non-interactive.
function mapBgRow(r) {
  const requestId = getField(r, 'Request ID');
  const ad = (state.noniAuthMap && requestId) ? state.noniAuthMap.get(String(requestId).trim()) : null;
  return {
    date:       parseDate(getField(r, 'Date (UTC)', 'Date')),
    user:       getField(r, 'User'),
    username:   getField(r, 'Username', 'User principal name'),
    status:     getField(r, 'Status'),
    failure:    getField(r, 'Failure reason'),
    authMethod: getField(r, 'Authentication requirement', 'Authentication method', 'Multifactor authentication auth method'),
    mfaMethod:  getField(r, 'Multifactor authentication auth method'),
    mfaResult:  getField(r, 'Multifactor authentication result'),
    ip:         getField(r, 'IP address'),
    location:   getField(r, 'Location'),
    app:        getField(r, 'Application'),
    resource:   getField(r, 'Resource'),
    ca:         getField(r, 'Conditional Access'),
    browser:    getField(r, 'Browser'),
    os:         getField(r, 'Operating System'),
    errCode:    getField(r, 'Sign-in error code'),
    clientApp:  getField(r, 'Client app'),
    signInType: 'nonInteractive',
    requestId:  requestId || null,
    tokenType:  getField(r, 'Incoming token type'),
    asn:        getField(r, 'Autonomous system number'),
    deviceId:   getField(r, 'Device ID'),
    compliant:  getField(r, 'Compliant'),
    managed:    getField(r, 'Managed'),
    joinType:   getField(r, 'Join Type'),
    mfaByToken: ad ? ad.satisfiedByToken : null,   // null = Source 05 not loaded / no match
  };
}
const isBg = s => s && s.signInType === 'nonInteractive';

// Build Map<lowercaseUPN, authRecord> from the auth-methods CSV
function buildAuthMap(rows) {
  const m = new Map();
  for (const r of rows) {
    const upn = getField(r, 'UPN', 'UserPrincipalName', 'upn');
    if (!upn) continue;
    const key = String(upn).toLowerCase().trim();
    // Truthy parser: handles "True"/"TRUE"/true/1/"1"/"yes"
    const yes = v => /^(true|1|yes|y)$/i.test(String(v ?? '').trim());
    m.set(key, {
      upn: String(upn).trim(),
      password:     yes(getField(r, 'Password')),
      email:        yes(getField(r, 'Email')),
      authenticator:yes(getField(r, 'AuthenticatorApp', 'Authenticator', 'MicrosoftAuthenticator')),
      sms:          yes(getField(r, 'SMS')),
      voice:        yes(getField(r, 'VoiceCall', 'Voice')),
      fido2:        yes(getField(r, 'FIDO2')),
      windowsHello: yes(getField(r, 'WindowsHello', 'WindowsHelloForBusiness')),
      tap:          yes(getField(r, 'TemporaryAccessPass', 'TAP')),
      methodCount:  Number(getField(r, 'MethodCount')) || 0,
      defaultMethod:getField(r, 'DefaultSignInMethod', 'Default Sign-In Method') || 'Unknown',
      status:       getField(r, 'Status') || '—',
    });
  }
  return m;
}

// Classify auth posture into phishing-resistance tier
function classifyPosture(auth) {
  if (!auth) return { tier: 'unknown', label: 'Not in registry' };
  if (auth.fido2 || auth.windowsHello) return { tier: 'resistant', label: 'Phishing-resistant' };
  if (auth.authenticator)              return { tier: 'strong',    label: 'Strong MFA' };
  if (auth.sms || auth.voice || auth.email) return { tier: 'weak', label: 'Weak MFA' };
  if (auth.password)                   return { tier: 'none',      label: 'Password only' };
  return { tier: 'unknown', label: 'No methods registered' };
}

// Return list of registered method chips for a user
function authChips(auth) {
  if (!auth) return [];
  const out = [];
  if (auth.fido2)         out.push({ label: 'FIDO2',         cls: 'strong' });
  if (auth.windowsHello)  out.push({ label: 'Windows Hello', cls: 'strong' });
  if (auth.authenticator) out.push({ label: 'Authenticator', cls: 'medium' });
  if (auth.sms)           out.push({ label: 'SMS',           cls: 'weak' });
  if (auth.voice)         out.push({ label: 'Voice',         cls: 'weak' });
  if (auth.email)         out.push({ label: 'Email OTP',     cls: 'weak' });
  if (auth.tap)           out.push({ label: 'TAP',           cls: 'medium' });
  if (auth.password)      out.push({ label: 'Password',      cls: 'none' });
  return out;
}

// Lookup an auth record by username/UPN (or by display name fallback)
function lookupAuth(username, displayName) {
  if (!state.authMap) return null;
  if (username) {
    const hit = state.authMap.get(String(username).toLowerCase().trim());
    if (hit) return hit;
  }
  if (displayName) {
    // fallback: scan map for a record whose UPN local-part starts with the slugged display name
    const slug = String(displayName).toLowerCase().replace(/\s+/g, '.');
    for (const [k, v] of state.authMap) {
      if (k.startsWith(slug + '@') || k.startsWith(slug + '.')) return v;
    }
  }
  return null;
}

function updateAnalyzeReady() {
  const btn = $('analyzeBtn');
  const status = $('analyzeStatus');
  const dot = $('liveDot');
  const live = $('liveStatus');
  if (state.riskyData && state.allData) {
    btn.disabled = false;
    const authNote = state.authMap ? ` · Auth registry: ${state.authMap.size} users` : ' · Auth registry: not provided (optional)';
    const bgNote = state.noniData                                                       // v12.3
      ? ` · Background: ${state.noniData.length.toLocaleString()} rows` + (state.noniAuthMap ? ` + auth details (${state.noniAuthMap.size.toLocaleString()} IDs)` : '')
      : '';
    status.innerHTML = '<span class="accent">▸</span> Both sources loaded. Risky: ' +
      state.riskyData.length + ' rows · Interactive: ' + state.allData.length + ' rows' + authNote + bgNote + '. Ready to analyze.';
    dot.classList.remove('idle');
    live.textContent = 'SOURCES LOADED';
  } else if (state.riskyData || state.allData) {
    btn.disabled = true;
    const have = state.riskyData ? 'Risky logs' : 'Interactive logs';
    const need = state.riskyData ? 'Interactive logs' : 'Risky logs';
    status.innerHTML = '<span class="accent">▸</span> ' + have + ' received. Waiting for ' + need + '…';
    live.textContent = 'PARTIAL';
  } else {
    btn.disabled = true;
    status.innerHTML = '<span class="accent">▸</span> Awaiting both log sources to begin correlation analysis';
    dot.classList.add('idle');
    live.textContent = 'AWAITING INGEST';
  }
}

// ============================================================
// Investigation engine
// ============================================================
function runInvestigation() {
  try {
    const report = analyze(state.riskyData, state.allData, state.noniData);
    state.report = report;
    renderReport(report);
    $('report').classList.add('active');
    $('liveStatus').textContent = 'ANALYSIS COMPLETE';
    $('liveDot').classList.remove('idle');
    setTimeout(() => $('sec-summary').scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
  } catch (err) {
    console.error(err);
    showToast('Analysis error: ' + err.message);
  }
}

function analyze(risky, all, noni) {
  // v9: observed window — real min/max span across interactive + risky dates.
  // Drives every "N-day observation window" label so it matches the data
  // (7, 31, or any range) instead of a hardcoded 7. No effect on correlation.
  const DAY_MS = 86400000;
  let winMin = Infinity, winMax = -Infinity;
  const scanDate = r => {
    const d = parseDate(getField(r, 'Date (UTC)', 'Date', 'DateTime'));
    if (d) { const t = d.getTime(); if (t < winMin) winMin = t; if (t > winMax) winMax = t; }
  };
  for (const r of all) scanDate(r);
  for (const r of risky) scanDate(r);
  const winHas = isFinite(winMin) && isFinite(winMax) && winMax >= winMin;
  // Inclusive UTC calendar-day count: Jul25 00:00 → Jul31 23:59 = 7 days.
  const winDays = winHas
    ? (Math.floor(winMax / DAY_MS) - Math.floor(winMin / DAY_MS) + 1)
    : 0;
  const winPhrase = winHas ? `${winDays}-day` : 'available';

  const atRiskRows = risky.filter(r => {
    const rs = getField(r, 'Risk state', 'risk state', 'RiskState');
    return rs && String(rs).toLowerCase().includes('at risk');
  });

  const userMap = new Map();
  for (const r of atRiskRows) {
    const user = getField(r, 'User', 'Name', 'User display name');
    if (!user) continue;
    const key = String(user).toLowerCase().trim();
    if (!userMap.has(key)) {
      userMap.set(key, {
        name: String(user).trim(),
        riskyEvents: [],
      });
    }
    userMap.get(key).riskyEvents.push({
      date: parseDate(getField(r, 'Date (UTC)', 'Date', 'DateTime')),
      ip: getField(r, 'IP address', 'IpAddress'),
      location: getField(r, 'Location'),
      riskState: getField(r, 'Risk state'),
    });
  }

  // v12.3 — bucket Source 04 rows by risky user, ONCE. Rows for users who are
  // not on the risky list are dropped here, so large exports stay cheap.
  let bgCtx = null;
  if (noni && noni.length) {
    const byUser = new Map(); let withUser = 0, kept = 0;
    for (const r of noni) {
      const user = getField(r, 'User', 'Name', 'User display name');
      if (!user) continue;
      withUser++;
      const k = String(user).toLowerCase().trim();
      if (!userMap.has(k)) continue;
      if (!byUser.has(k)) byUser.set(k, []);
      byUser.get(k).push(r); kept++;
    }
    const total = (state.noniMeta && state.noniMeta.totalWithUser != null) ? state.noniMeta.totalWithUser : withUser;
    bgCtx = { byUser, total, kept };
  }

  const investigations = [];
  for (const [key, u] of userMap) {
    const signins = all.filter(r => {
      const user = getField(r, 'User', 'Name', 'User display name');
      if (user && String(user).toLowerCase().trim() === key) return true;
      return false;
    }).map(r => ({
      date:    parseDate(getField(r, 'Date (UTC)', 'Date')),
      user:    getField(r, 'User'),
      username:getField(r, 'Username', 'User principal name'),
      status:  getField(r, 'Status'),
      failure: getField(r, 'Failure reason'),
      authMethod: getField(r, 'Authentication requirement', 'Authentication method', 'Multifactor authentication auth method'),
      mfaMethod: getField(r, 'Multifactor authentication auth method'),
      mfaResult: getField(r, 'Multifactor authentication result'),
      ip:      getField(r, 'IP address'),
      location:getField(r, 'Location'),
      app:     getField(r, 'Application'),
      ca:      getField(r, 'Conditional Access'),
      browser: getField(r, 'Browser'),
      os:      getField(r, 'Operating System'),
      errCode: getField(r, 'Sign-in error code'),
    }));

    signins.sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
    const username = signins.find(s => s.username)?.username || null;
    const timeline = [...signins.map(s => ({ ...s, kind: 'signin' }))];
    for (const r of u.riskyEvents) {
      timeline.push({
        date: r.date,
        kind: 'risky',
        ip: r.ip,
        location: r.location,
        riskState: r.riskState,
      });
    }
    timeline.sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));

    // ---- Risky-context filtering ------------------------------------------
    // Build the sets of risky-flagged IPs and Locations first, then narrow
    // interactive sign-ins to only those whose IP OR Location matches. All
    // investigation columns (First / Latest Sign-in, Auth Method, Failure
    // Reason, counts) must be sourced from this narrowed subset so the report
    // reflects sign-ins tied to a risky event — not the user's normal traffic.
    const locSet = new Set(signins.map(s => s.location).filter(Boolean));
    const ipSet  = new Set(signins.map(s => s.ip).filter(Boolean));
    const riskyLocSet = new Set(u.riskyEvents.map(r => r.location).filter(Boolean));
    const riskyIpSet = new Set(u.riskyEvents.map(r => r.ip).filter(Boolean));

    const riskyContextSignins = signins.filter(s =>
      (s.ip && riskyIpSet.has(s.ip)) ||
      (s.location && riskyLocSet.has(s.location))
    );

    const first = riskyContextSignins[0] || null;
    const last  = riskyContextSignins[riskyContextSignins.length - 1] || null;

    // Counts are scoped to risky-context sign-ins only
    const counts = { success: 0, failure: 0, interrupt: 0 };
    for (const s of riskyContextSignins) {
      const st = String(s.status || '').toLowerCase();
      if (st.includes('success')) counts.success++;
      else if (st.includes('interrupt')) counts.interrupt++;
      else if (st.includes('failure')) counts.failure++;
    }

    // Post-risk buckets: risky-context AND at/after earliest risky event
    const earliestRiskTs = Math.min(...u.riskyEvents.map(r => r.date?.getTime() || Infinity));
    const postRisky = riskyContextSignins.filter(s => s.date && s.date.getTime() >= earliestRiskTs);
    const postRiskySuccess   = postRisky.filter(s => String(s.status || '').toLowerCase().includes('success'));
    const postRiskyInterrupt = postRisky.filter(s => String(s.status || '').toLowerCase().includes('interrupt'));
    const postRiskyFailure   = postRisky.filter(s => String(s.status || '').toLowerCase().includes('failure'));

    let assessment = 'INDETERMINATE';
    let assessmentClass = 'review';
    let assessmentReason = '';

    if (signins.length === 0) {
      assessment = 'NO INTERACTIVE LOGS';
      assessmentClass = 'review';
      assessmentReason = 'No matching records found in the interactive sign-in log for this user. Cannot determine outcome from supplied data.';
    } else if (riskyContextSignins.length === 0) {
      assessment = 'NO RISKY-CONTEXT ACTIVITY';
      assessmentClass = 'review';
      assessmentReason = 'Interactive sign-ins exist for this user but NONE match a risky-flagged IP or Location. The risky detection was not corroborated by any interactive log entry — verify the detection source and check whether interactive logs cover the same time window.';
    } else if (postRiskySuccess.length > 0 && riskyIpSet.size > 0 &&
               postRiskySuccess.some(s => riskyIpSet.has(s.ip))) {
      assessment = 'POSSIBLE COMPROMISE';
      assessmentClass = 'compromise';
      assessmentReason = 'Successful authentication observed from an IP address flagged as risky. Account warrants immediate response.';
    } else if (postRiskyInterrupt.length > 0 && postRiskySuccess.length === 0) {
      assessment = 'CONTROLS EFFECTIVE';
      assessmentClass = 'mitigated';
      assessmentReason = 'Sign-in attempts from risky IPs/locations were interrupted by security controls (MFA / Conditional Access). No completed authentication recorded from a risky context.';
    } else if (postRiskyFailure.length > 0 && postRiskySuccess.length === 0) {
      assessment = 'AUTH FAILED';
      assessmentClass = 'mitigated';
      assessmentReason = 'Authentication attempts from risky IPs/locations failed (likely invalid credentials). No successful sign-in from a risky context was recorded.';
    } else if (postRiskySuccess.length > 0) {
      assessment = 'REVIEW REQUIRED';
      assessmentClass = 'review';
      assessmentReason = 'Successful sign-ins occurred from a risky location after detection, but not from an exact risky IP match. Verify with the user.';
    } else {
      assessment = 'NO POST-RISK ACTIVITY';
      assessmentClass = 'ok';
      assessmentReason = `No risky-context sign-in activity recorded after the risky detection within the ${winPhrase} observation window.`;
    }

    // ---- v12.3 background (non-interactive) layer ----------------------------
    // Same risky-context rule as interactive (invariant 1/2). Interactive fields
    // (first/last/counts) are NOT touched; background only escalates.
    const bgSignins = bgCtx ? (bgCtx.byUser.get(key) || []).map(mapBgRow)
      .sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0)) : [];
    const bgRiskyContext = bgSignins.filter(s =>
      (s.ip && riskyIpSet.has(s.ip)) || (s.location && riskyLocSet.has(s.location)));
    const bgPostRisky = bgRiskyContext.filter(s => s.date && s.date.getTime() >= earliestRiskTs);
    const bgPostRiskySuccess = bgPostRisky.filter(s => String(s.status || '').toLowerCase().includes('success'));
    const bgPostRiskyFailure = bgPostRisky.filter(s => String(s.status || '').toLowerCase().includes('failure'));
    const bgCompromise = bgPostRiskySuccess.some(s => s.ip && riskyIpSet.has(s.ip));
    let bgEscalated = false;

    if (bgCtx) {
      const bgSucc = bgPostRiskySuccess.length, bgFail = bgPostRiskyFailure.length;
      if (bgCompromise) {
        if (assessmentClass === 'compromise') {
          assessmentReason += ' Background (non-interactive) token sign-ins from the same risky IP were also observed after detection.';
        } else {
          assessment = 'POSSIBLE COMPROMISE';
          assessmentClass = 'compromise';
          assessmentReason = 'Successful background (non-interactive) token sign-in from an IP address flagged as risky, after the detection — consistent with token replay. This path needs no password or MFA prompt, which is why it does not appear in interactive logs. Account warrants immediate response.';
          bgEscalated = true;
        }
      } else if (assessmentClass !== 'compromise') {
        const noEvidence = assessment === 'NO INTERACTIVE LOGS' || assessment === 'NO RISKY-CONTEXT ACTIVITY';
        if (bgSucc > 0 || (noEvidence && bgRiskyContext.length > 0)) {
          assessment = 'REVIEW REQUIRED';
          assessmentClass = 'review';
          assessmentReason = `Background (non-interactive) sign-ins matched a risky IP or location (${bgRiskyContext.length} total; after detection: ${bgSucc} successful, ${bgFail} failed), but no success from an exact risky IP. Verify whether the user's session or device is expected at that location.`;
          bgEscalated = true;
        } else if (bgSignins.length > 0) {
          assessmentReason += ` ${bgSignins.length} background (non-interactive) sign-in${bgSignins.length === 1 ? '' : 's'} reviewed; none from a risky IP or location${bgRiskyContext.length ? ' after detection' : ''}.`;
        }
      }
      if (bgRiskyContext.length) {
        for (const s of bgRiskyContext) timeline.push({ ...s, kind: 'signin' });
        timeline.sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
      }
    }
    const effUsername = username || bgSignins.find(s => s.username)?.username || null;

    investigations.push({
      name: u.name,
      username: effUsername,
      bgSignins, bgRiskyContext, bgPostRiskySuccess, bgPostRiskyFailure, bgCompromise, bgEscalated,
      riskyEvents: u.riskyEvents,
      signins,                 // full population (kept for timeline / anomaly cells)
      riskyContextSignins,     // narrowed subset used for investigation columns
      timeline,                // full timeline (unchanged — forensic context)
      first,
      last,
      counts,
      postRiskySuccess,
      postRiskyInterrupt,
      postRiskyFailure,
      locSet, ipSet, riskyLocSet, riskyIpSet,
      assessment,
      assessmentClass,
      assessmentReason,
      auth: lookupAuth(effUsername, u.name),
      authPosture: classifyPosture(lookupAuth(effUsername, u.name)),
    });
  }

  const sevRank = { compromise: 0, review: 1, mitigated: 2, ok: 3 };
  investigations.sort((a, b) =>
    (sevRank[a.assessmentClass] ?? 9) - (sevRank[b.assessmentClass] ?? 9)
    || a.name.localeCompare(b.name)
  );

  // ---- v10 risk scoring (bind to risky-context — invariants 1 & 5) ----------
  // Data-availability flags gate the terms that depend on optional columns so a
  // missing source lowers the term to 0 and marks the score "partial" rather
  // than silently understating risk.
  const authAvailable = !!(state.authMap && state.authMap.size);
  const caAvailable = all.some(r => {
    const v = getField(r, 'Conditional Access');
    return v != null && String(v).trim() !== '';
  });
  for (const inv of investigations) {
    inv.risk = computeRiskScore(inv, { authAvailable, caAvailable });
    inv.mitre = computeMitre(inv);
  }

  // Aggregate unique risky IPs / locations across all users
  const uAllIps = new Set(), uAllLocs = new Set();
  for (const inv of investigations) {
    inv.riskyIpSet.forEach(v => uAllIps.add(v));
    inv.riskyLocSet.forEach(v => uAllLocs.add(v));
  }

  const agg = {
    totalRiskyEvents: atRiskRows.length,
    totalRiskyUsers: userMap.size,
    totalInteractiveSignins: all.length,
    compromised: investigations.filter(i => i.assessmentClass === 'compromise').length,
    mitigated:   investigations.filter(i => i.assessmentClass === 'mitigated').length,
    review:      investigations.filter(i => i.assessmentClass === 'review').length,
    ok:          investigations.filter(i => i.assessmentClass === 'ok').length,
    anyPostRiskSuccess: investigations.some(i => i.postRiskySuccess.length > 0),
    // v9 observed-window fields
    windowHas:   winHas,
    windowDays:  winDays,
    windowStart: winHas ? new Date(winMin) : null,
    windowEnd:   winHas ? new Date(winMax) : null,
    // v10 analytics fields
    authAvailable,
    caAvailable,
    uniqueRiskyIps:   uAllIps.size,
    uniqueRiskyLocs:  uAllLocs.size,
    weakMfaUsers:     authAvailable ? investigations.filter(i => i.authPosture.tier === 'weak').length : null,
    passwordOnlyUsers:authAvailable ? investigations.filter(i => i.authPosture.tier === 'none').length : null,
    highRiskAccounts: investigations.filter(i => i.risk.tier === 'High' || i.risk.tier === 'Critical').length,
    // v12.3 background layer
    bgLoaded:        !!bgCtx,
    bgAuthLoaded:    !!state.noniAuthMap,
    bgTotal:         bgCtx ? bgCtx.total : 0,
    bgKept:          bgCtx ? bgCtx.kept : 0,
    bgRiskyContext:  investigations.reduce((n, i) => n + i.bgRiskyContext.length, 0),
    bgCompromised:   investigations.filter(i => i.bgCompromise).length,
    bgEscalations:   investigations.filter(i => i.bgEscalated).length,
    bgMfaByToken:    investigations.reduce((n, i) => n + i.bgSignins.filter(s => s.mfaByToken === true).length, 0),
    bgMfaChecked:    investigations.reduce((n, i) => n + i.bgSignins.filter(s => s.mfaByToken !== null).length, 0),
  };

  // ---- v11 trend + Conditional Access aggregations ------------------------
  // Detections by day/hour from the risky detections; risky-context sign-in
  // outcomes by day (bound to risky-context per invariant 1). Empty buckets stay
  // 0 — never interpolated.
  const detByDay = new Map();
  const detByHour = new Array(24).fill(0);
  for (const r of atRiskRows) {
    const d = parseDate(getField(r, 'Date (UTC)', 'Date', 'DateTime'));
    if (!d) continue;
    const day = d.toISOString().slice(0, 10);
    detByDay.set(day, (detByDay.get(day) || 0) + 1);
    detByHour[d.getUTCHours()]++;
  }
  const outByDay = new Map();
  for (const inv of investigations) {
    for (const s of inv.riskyContextSignins) {
      if (!s.date) continue;
      const day = s.date.toISOString().slice(0, 10);
      const o = outByDay.get(day) || { success: 0, failure: 0, interrupt: 0 };
      const st = String(s.status || '').toLowerCase();
      if (st.includes('success')) o.success++;
      else if (st.includes('interrupt')) o.interrupt++;
      else if (st.includes('failure')) o.failure++;
      outByDay.set(day, o);
    }
  }

  // Conditional Access analytics across all interactive sign-ins (degrades if
  // the column is absent). MFA-required rate from the auth-requirement column.
  let caS = 0, caF = 0, caNA = 0, caOther = 0;
  for (const r of all) {
    const v = getField(r, 'Conditional Access');
    if (v == null || String(v).trim() === '') continue;
    const t = String(v).toLowerCase();
    if (t.includes('success')) caS++;
    else if (t.includes('fail')) caF++;
    else if (t.includes('notapplied') || t.includes('not applied') || t.includes('notenabled')) caNA++;
    else caOther++;
  }
  let mfaReq = 0, arTotal = 0, arAvail = false;
  for (const r of all) {
    const v = getField(r, 'Authentication requirement', 'Authentication method', 'Multifactor authentication auth method');
    if (v == null || String(v).trim() === '') continue;
    arAvail = true; arTotal++;
    if (/multi.?factor|mfa/i.test(String(v))) mfaReq++;
  }
  agg.ca = {
    available: caAvailable,
    success: caS, failure: caF, notApplied: caNA, other: caOther,
    total: caS + caF + caNA + caOther,
    mfaAvailable: arAvail, mfaRequired: mfaReq, mfaTotal: arTotal,
  };

  return { investigations, agg, trends: { detByDay, detByHour, outByDay } };
}

// ============================================================
// Render
// ============================================================
function renderReport(report) {
  resetCharts();
  renderKpis(report);
  renderIncidentQueue(report);
  renderTrends(report);
  renderCaAnalytics(report);
  renderMitre(report);
  renderGeo(report);
  renderIoc(report);
  renderSummary(report);
  renderTable(report);
  renderIpRegistry(report);
  renderUserCards(report);
  renderAuthPostureSection(report);
  renderConclusion(report);
  renderRecs(report);
  renderBgExtras(report);   // v12.3 — no-op unless Source 04 loaded
}

// v12.3 — background-layer additions to existing sections. Pure appends; does
// nothing when Source 04 is absent, so the v12.2 report is unchanged.
function renderBgExtras(report) {
  const a = report.agg;
  if (!a.bgLoaded) return;
  // KPI tile
  const grid = $('kpiGrid');
  if (grid) grid.insertAdjacentHTML('beforeend', `
    <div class="kpi-card kpi-${a.bgCompromised > 0 ? 'crit' : 'ok'}">
      <div class="kpi-value">${a.bgCompromised}</div>
      <div class="kpi-label">Token-from-Risky-IP</div>
      <div class="kpi-desc">Background sign-in accepted from an exact risky IP (possible replay) · ${a.bgEscalations} user${a.bgEscalations === 1 ? '' : 's'} escalated by background logs</div>
    </div>`);
  // Conditional Access / MFA note — background rows are excluded from those rates
  const ca = $('caBody');
  if (ca) ca.insertAdjacentHTML('beforeend', `
    <div class="bg-note">Background (non-interactive) sign-ins are <strong>excluded</strong> from the CA and MFA-required rates above: they reuse MFA completed earlier, so they always read "single-factor / not applied".
    ${a.bgAuthLoaded ? `Auth details confirm <strong>${a.bgMfaByToken}</strong> of ${a.bgMfaChecked} matched background rows for risky users were <em>MFA satisfied by earlier token</em>.` : 'Load Source 05 (auth details) to confirm the earlier MFA per request.'}</div>`);
  // Executive summary sentence
  const sum = $('summaryNarrative');
  if (sum) sum.insertAdjacentHTML('beforeend', `
    <p style="margin-top:14px">Background (non-interactive) sign-ins were also reviewed: <strong class="hl-cyan">${a.bgKept.toLocaleString()}</strong> rows belonged to risky users (of ${a.bgTotal.toLocaleString()} ingested), of which <strong class="hl-cyan">${a.bgRiskyContext}</strong> matched a risky IP or location.
    ${a.bgCompromised > 0
      ? `<span class="hl-crimson">${a.bgCompromised} user${a.bgCompromised === 1 ? '' : 's'} had a background token accepted from an exact risky IP after detection</span> — consistent with token replay, which interactive logs alone would not show.`
      : 'No background token was accepted from an exact risky IP after detection.'}</p>`);
}

// ============================================================
// v11 — Chart engine (Tier B). Every panel renders a Chart.js canvas when the
// library loaded, and its underlying data TABLE when it did not (offline). No
// panel is ever left blank; empty windows show an explicit empty state.
// ============================================================
const CX = {
  green:   '#2d6b4a', crimson: '#a81f2d', amber: '#a8590f', cyan: '#0e6b8e',
  cyanFill:'rgba(14,107,142,0.12)', gray: '#9c988c',
  grid:    'rgba(26,24,20,0.07)',   tick: '#7a786f',
  font:    { family: "'JetBrains Mono', monospace", size: 11 },
};
let _charts = [];
function resetCharts() {
  _charts.forEach(c => { try { c.destroy(); } catch (e) {} });
  _charts = [];
}
function chartAvailable() { return typeof Chart !== 'undefined'; }

function chartOpts(type) {
  const base = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { labels: { color: CX.tick, font: CX.font, boxWidth: 12 } } },
  };
  if (type === 'doughnut' || type === 'pie') return base;
  base.scales = {
    x: { ticks: { color: CX.tick, font: CX.font, maxRotation: 0, autoSkip: true }, grid: { color: CX.grid } },
    y: { beginAtZero: true, ticks: { color: CX.tick, font: CX.font, precision: 0 }, grid: { color: CX.grid } },
  };
  return base;
}

// Render a chart into `mount`, or its data table if Chart.js is unavailable.
function chartOrTable(mount, spec) {
  const { type, labels, datasets, empty } = spec;
  if (!labels || !labels.length) {
    mount.style.height = 'auto';
    mount.innerHTML = `<div class="cx-empty">${escapeHtml(empty || 'No data in this window.')}</div>`;
    return;
  }
  if (!chartAvailable()) {
    mount.style.height = 'auto';
    const cols = ['', ...datasets.map(d => d.label || 'Value')];
    const rows = labels.map((lb, i) => [lb, ...datasets.map(d => d.data[i] ?? 0)]);
    mount.innerHTML =
      `<div class="cx-note">Chart library unavailable offline — data table shown instead.</div>`
      + `<div class="cx-tablewrap"><table class="cx-table"><thead><tr>`
      + cols.map(c => `<th>${escapeHtml(String(c))}</th>`).join('')
      + `</tr></thead><tbody>`
      + rows.map(r => `<tr>` + r.map((c, ci) =>
          `<td${ci === 0 ? ' class="cx-lbl"' : ''}>${escapeHtml(String(c))}</td>`).join('') + `</tr>`).join('')
      + `</tbody></table></div>`;
    return;
  }
  mount.style.height = '270px';
  mount.innerHTML = '';
  const cv = document.createElement('canvas');
  mount.appendChild(cv);
  const opts = chartOpts(type);
  if (spec.horizontal) opts.indexAxis = 'y';
  _charts.push(new Chart(cv, { type, data: { labels, datasets }, options: opts }));
}

// v11 — Threat trend charts
function renderTrends(report) {
  const t = report.trends;
  const days = [...t.detByDay.keys()].sort();
  chartOrTable($('cxDetDay'), {
    type: 'line', empty: 'No risky detections in this window.',
    labels: days,
    datasets: [{ label: 'Detections', data: days.map(d => t.detByDay.get(d)),
      borderColor: CX.cyan, backgroundColor: CX.cyanFill, tension: 0.25, fill: true, pointRadius: 2 }],
  });

  const hours = [...Array(24).keys()].map(h => String(h).padStart(2, '0'));
  chartOrTable($('cxDetHour'), {
    type: 'bar', empty: 'No risky detections in this window.',
    labels: hours,
    datasets: [{ label: 'Detections', data: t.detByHour, backgroundColor: CX.cyan }],
  });

  const odays = [...t.outByDay.keys()].sort();
  chartOrTable($('cxOutDay'), {
    type: 'line', empty: 'No risky-context sign-ins in this window.',
    labels: odays,
    datasets: [
      { label: 'Success',   data: odays.map(d => t.outByDay.get(d).success),   borderColor: CX.green,   backgroundColor: CX.green,   tension: 0.2, pointRadius: 2 },
      { label: 'Interrupt', data: odays.map(d => t.outByDay.get(d).interrupt), borderColor: CX.amber,   backgroundColor: CX.amber,   tension: 0.2, pointRadius: 2 },
      { label: 'Failure',   data: odays.map(d => t.outByDay.get(d).failure),   borderColor: CX.crimson, backgroundColor: CX.crimson, tension: 0.2, pointRadius: 2 },
    ],
  });

  renderPostureChart(report);
}

// v11 — Auth posture doughnut (tenant-wide; needs the auth-methods source)
function renderPostureChart(report) {
  const mount = $('cxPosture');
  if (!report.agg.authAvailable || !state.authMap) {
    mount.style.height = 'auto';
    mount.innerHTML = `<div class="cx-empty">Requires auth-methods source (Source 03).</div>`;
    return;
  }
  const tiers = { resistant: 0, strong: 0, weak: 0, none: 0, unknown: 0 };
  for (const [, a] of state.authMap) tiers[classifyPosture(a).tier]++;
  const labels = ['Phishing-resistant', 'Strong MFA', 'Weak MFA', 'Password-only', 'Unknown'];
  const data   = [tiers.resistant, tiers.strong, tiers.weak, tiers.none, tiers.unknown];
  chartOrTable(mount, {
    type: 'doughnut', empty: 'No registered users.',
    labels,
    datasets: [{ label: 'Users', data,
      backgroundColor: [CX.green, CX.cyan, CX.amber, CX.crimson, CX.gray] }],
  });
}

// v11 — Conditional Access analytics
function renderCaAnalytics(report) {
  const ca = report.agg.ca;
  const body = $('caBody');
  if (!ca.available) {
    body.innerHTML = `<div class="cx-empty">Conditional Access — Not present in source data.</div>`;
    return;
  }
  const applied = ca.success + ca.failure;
  const rate = (n, d) => d ? Math.round(100 * n / d) : 0;
  const kpis = [
    { label: 'CA Success Rate',    value: rate(ca.success, applied) + '%', desc: 'success / applied',        cls: 'ok' },
    { label: 'CA Block Rate',      value: rate(ca.failure, applied) + '%', desc: 'failure / applied',        cls: ca.failure > 0 ? 'crit' : 'ok' },
    { label: 'CA Enforcement Rate',value: rate(applied, ca.total) + '%',   desc: 'applied / all sign-ins',   cls: 'cyan' },
  ];
  if (ca.mfaAvailable) {
    kpis.push({ label: 'MFA-Required Rate', value: rate(ca.mfaRequired, ca.mfaTotal) + '%', desc: 'multi-factor required', cls: 'cyan' });
  }
  body.innerHTML =
    `<div class="kpi-grid">`
    + kpis.map(c => `
      <div class="kpi-card kpi-${c.cls}">
        <div class="kpi-value">${c.value}</div>
        <div class="kpi-label">${escapeHtml(c.label)}</div>
        <div class="kpi-desc">${escapeHtml(c.desc)}</div>
      </div>`).join('')
    + `</div>
    <div class="cx-grid">
      <div class="cx-card"><div class="cx-title">CA Outcome Share</div><div class="cx-mount" id="cxCaDough"></div></div>
      <div class="cx-card"><div class="cx-title">CA Outcome Counts</div><div class="cx-mount" id="cxCaBar"></div></div>
    </div>`;

  const labels = ['Success', 'Failure', 'Not applied'];
  const data   = [ca.success, ca.failure, ca.notApplied];
  const colors = [CX.green, CX.crimson, CX.gray];
  if (ca.other > 0) { labels.push('Other'); data.push(ca.other); colors.push(CX.cyan); }

  chartOrTable($('cxCaDough'), { type: 'doughnut', labels, datasets: [{ label: 'Sign-ins', data, backgroundColor: colors }] });
  chartOrTable($('cxCaBar'),   { type: 'bar',      labels, datasets: [{ label: 'Sign-ins', data, backgroundColor: colors }] });
}

// v12 — MITRE ATT&CK matrix cards (heuristic)
function renderMitre(report) {
  const map = new Map();
  for (const inv of report.investigations) {
    for (const t of inv.mitre) {
      const e = map.get(t.id) || { id: t.id, name: t.name, tactic: t.tactic, sev: t.sev, note: t.note, users: new Set(), evidence: 0 };
      e.users.add(inv.name); e.evidence += t.evidence; map.set(t.id, e);
    }
  }
  const rank = { high: 0, medium: 1, low: 2 };
  const techs = [...map.values()].sort((a, b) => (rank[a.sev] ?? 9) - (rank[b.sev] ?? 9));
  const body = $('mitreBody');
  if (!techs.length) { body.innerHTML = `<div class="cx-empty">No techniques inferred from the current data.</div>`; return; }
  body.innerHTML = `<div class="mitre-grid">` + techs.map(t => `
    <div class="mitre-card mitre-${MITRE_SEV[t.sev] || 'medium'}">
      <div class="mitre-id">${escapeHtml(t.id)}</div>
      <div class="mitre-name">${escapeHtml(t.name)}</div>
      <div class="mitre-tactic">${escapeHtml(t.tactic)}</div>
      <div class="mitre-stats"><span>${t.users.size} user${t.users.size === 1 ? '' : 's'}</span><span>${t.evidence} evidence row${t.evidence === 1 ? '' : 's'}</span></div>
      <div class="mitre-desc">${escapeHtml(t.note)}</div>
    </div>`).join('') + `</div>`;
}

// v12 — Attack geography (offline, country-level)
function renderGeo(report) {
  const body = $('geoBody');
  const map = new Map();
  const add = (name, code) => { let e = map.get(name); if (!e) { e = { name, code, det: 0, users: new Set(), success: 0, failure: 0, interrupt: 0 }; map.set(name, e); } return e; };
  for (const inv of report.investigations) {
    for (const ev of inv.riskyEvents) { const c = parseCountry(ev.location); const e = add(c.name, c.code); e.det++; e.users.add(inv.name); }
    for (const s of inv.riskyContextSignins) {
      const c = parseCountry(s.location); const e = add(c.name, c.code);
      const st = String(s.status || '').toLowerCase();
      if (st.includes('success')) e.success++; else if (st.includes('interrupt')) e.interrupt++; else if (st.includes('failure')) e.failure++;
    }
  }
  const rows = [...map.values()].sort((a, b) => b.det - a.det || b.users.size - a.users.size);
  if (!rows.length) { body.innerHTML = `<div class="cx-empty">No location data to map.</div>`; return; }
  const top = rows[0];
  const kpi = (label, value, desc, cls, big) =>
    `<div class="kpi-card kpi-${cls}"><div class="kpi-value" style="font-size:${big ? '26px' : '18px'}">${escapeHtml(String(value))}</div><div class="kpi-label">${escapeHtml(label)}</div><div class="kpi-desc">${escapeHtml(desc)}</div></div>`;
  body.innerHTML =
    `<div class="kpi-grid">`
    + kpi('Countries Observed', rows.filter(r => r.name !== 'Unmapped').length, 'distinct source countries', 'cyan', true)
    + kpi('Top Source Country', top.name, top.det + ' detection' + (top.det === 1 ? '' : 's'), top.name === 'Unmapped' ? 'muted' : 'crit', false)
    + `</div>`
    + `<div class="cx-grid"><div class="cx-card cx-wide"><div class="cx-title">Detections by Country (Top 15)</div><div class="cx-mount" id="cxGeo"></div></div></div>`
    + `<div class="cx-tablewrap"><table class="cx-table"><thead><tr><th>Country</th><th>Detections</th><th>Users</th><th>Success</th><th>Interrupt</th><th>Failure</th></tr></thead><tbody>`
    + rows.map(r => `<tr><td class="cx-lbl">${escapeHtml(r.name)}</td><td>${r.det}</td><td>${r.users.size}</td><td>${r.success}</td><td>${r.interrupt}</td><td>${r.failure}</td></tr>`).join('')
    + `</tbody></table></div>`;
  const top15 = rows.slice(0, 15);
  chartOrTable($('cxGeo'), { type: 'bar', horizontal: true, labels: top15.map(r => r.name),
    datasets: [{ label: 'Detections', data: top15.map(r => r.det), backgroundColor: CX.crimson }] });
}

// v12 — IOC prevalence & link analysis (searchable). Prevalence, not reputation.
let IOC_STATE = { ipRows: [], linkRows: [] };
function renderIoc(report) {
  const body = $('iocBody');
  const ipMap = new Map();
  for (const inv of report.investigations) {
    for (const ev of inv.riskyEvents) {
      if (!ev.ip) continue;
      let e = ipMap.get(ev.ip);
      if (!e) { e = { ip: ev.ip, users: new Set(), det: 0, latest: null, loc: ev.location || '' }; ipMap.set(ev.ip, e); }
      e.users.add(inv.name); e.det++;
      if (ev.date && (!e.latest || ev.date > e.latest)) e.latest = ev.date;
    }
  }
  const ipRows = [...ipMap.values()].sort((a, b) => b.det - a.det || b.users.size - a.users.size);
  const linkRows = report.investigations.map(inv => ({
    name: inv.name, username: inv.username || '',
    ips: [...inv.riskyIpSet], locs: [...inv.riskyLocSet], events: inv.riskyEvents.length,
  })).sort((a, b) => b.events - a.events);
  IOC_STATE = { ipRows, linkRows };
  if (!ipRows.length && !linkRows.length) { body.innerHTML = `<div class="cx-empty">No indicators to explore.</div>`; return; }
  body.innerHTML = `
    <div class="ioc-search"><input type="text" id="iocSearch" placeholder="Filter by IP, user, or location…" oninput="filterIoc()"></div>
    <div class="ioc-cols">
      <div class="ioc-panel"><div class="cx-title">IP Prevalence</div><div class="cx-tablewrap"><table class="cx-table" id="iocIpTable"><thead><tr><th>IP Address</th><th>Users</th><th>Detections</th><th>Latest (UTC)</th></tr></thead><tbody></tbody></table></div></div>
      <div class="ioc-panel"><div class="cx-title">User &rarr; IP / Location Link</div><div class="cx-tablewrap"><table class="cx-table" id="iocLinkTable"><thead><tr><th>User</th><th>IPs</th><th>Locations</th><th>Events</th></tr></thead><tbody></tbody></table></div></div>
    </div>`;
  renderIocTables('');
}
function filterIoc() { renderIocTables((($('iocSearch') || {}).value || '').toLowerCase().trim()); }
function renderIocTables(q) {
  const CAP = 200;
  const ipRows = IOC_STATE.ipRows.filter(r => !q || r.ip.toLowerCase().includes(q)
    || [...r.users].some(u => u.toLowerCase().includes(q)) || (r.loc || '').toLowerCase().includes(q));
  const linkRows = IOC_STATE.linkRows.filter(r => !q || r.name.toLowerCase().includes(q)
    || (r.username || '').toLowerCase().includes(q)
    || r.ips.some(x => String(x).toLowerCase().includes(q)) || r.locs.some(x => String(x).toLowerCase().includes(q)));
  const ipT = $('iocIpTable').querySelector('tbody');
  ipT.innerHTML = ipRows.slice(0, CAP).map(r =>
    `<tr><td class="cx-lbl">${escapeHtml(r.ip)}</td><td>${r.users.size}</td><td>${r.det}</td><td>${escapeHtml(r.latest ? fmtDate(r.latest) : '—')}</td></tr>`).join('')
    || `<tr><td colspan="4" class="cx-empty">No matching IPs.</td></tr>`;
  const lkT = $('iocLinkTable').querySelector('tbody');
  lkT.innerHTML = linkRows.slice(0, CAP).map(r =>
    `<tr><td class="cx-lbl">${escapeHtml(r.name)}</td><td>${escapeHtml(r.ips.join(' | ') || '—')}</td><td>${escapeHtml(r.locs.join(' | ') || '—')}</td><td>${r.events}</td></tr>`).join('')
    || `<tr><td colspan="4" class="cx-empty">No matching users.</td></tr>`;
}

// v10 — Executive KPI cards
function renderKpis(report) {
  const { agg } = report;
  const authNote = 'requires auth-methods source';
  const cards = [
    { label: 'Total Risky Users',      value: agg.totalRiskyUsers,   desc: 'Distinct "At risk" accounts',          cls: 'cyan' },
    { label: 'Possible Compromise',    value: agg.compromised,       desc: 'Success from an exact risky IP',        cls: agg.compromised > 0 ? 'crit' : 'ok' },
    { label: 'Mitigated',              value: agg.mitigated,         desc: 'Controls effective / auth failed',      cls: 'ok' },
    { label: 'Review Required',        value: agg.review,            desc: 'Needs analyst confirmation',            cls: agg.review > 0 ? 'warn' : 'muted' },
    { label: 'No Post-Risk Activity',  value: agg.ok,                desc: 'No risky-context sign-ins after event', cls: 'muted' },
    { label: 'High-Risk Accounts',     value: agg.highRiskAccounts,  desc: 'Risk tier High or Critical',            cls: agg.highRiskAccounts > 0 ? 'crit' : 'ok' },
    { label: 'Weak-MFA Users',         value: agg.weakMfaUsers,      desc: agg.authAvailable ? 'SMS / Voice / Email only' : authNote, cls: 'warn' },
    { label: 'Password-Only Users',    value: agg.passwordOnlyUsers, desc: agg.authAvailable ? 'No MFA method registered' : authNote, cls: 'crit' },
    { label: 'Unique Risky IPs',       value: agg.uniqueRiskyIps,    desc: 'Across all risky detections',           cls: 'cyan' },
    { label: 'Unique Risky Locations', value: agg.uniqueRiskyLocs,   desc: 'Across all risky detections',           cls: 'cyan' },
  ];
  $('kpiGrid').innerHTML = cards.map(c => {
    const missing = c.value === null || c.value === undefined;
    const val = missing ? '—' : c.value.toLocaleString();
    const cls = missing ? 'muted' : c.cls;
    return `
      <div class="kpi-card kpi-${cls}">
        <div class="kpi-value">${val}</div>
        <div class="kpi-label">${escapeHtml(c.label)}</div>
        <div class="kpi-desc">${escapeHtml(c.desc)}</div>
      </div>`;
  }).join('');
}

// v10 — SOC Incident Queue (sorted by risk score, highest first)
function renderIncidentQueue(report) {
  const { investigations, agg } = report;
  const statusLabel = {
    compromise: 'Possible compromise',
    mitigated:  'Mitigated',
    review:     'Review required',
    ok:         'No post-risk activity',
  };
  const rows = [...investigations].sort((a, b) =>
    b.risk.score - a.risk.score || a.name.localeCompare(b.name)
  );
  const CAP = 100;
  const shown = rows.slice(0, CAP);
  $('iqSub').textContent =
    `${rows.length} account${rows.length === 1 ? '' : 's'} · sorted by risk score`
    + (rows.length > CAP ? ` · showing top ${CAP}` : '');

  if (!shown.length) {
    $('iqTbody').innerHTML = `<tr><td colspan="7" class="iq-empty">No risky accounts to queue.</td></tr>`;
    return;
  }
  $('iqTbody').innerHTML = shown.map(inv => {
    const r = inv.risk;
    const latest = inv.last && inv.last.date ? fmtDate(inv.last.date) : '—';
    const mfa = agg.authAvailable ? escapeHtml(inv.authPosture.label) : '—';
    const partial = r.partial
      ? ` <span class="iq-partial" title="Score is partial — auth and/or Conditional Access data not loaded, so some terms could not be evaluated.">partial</span>`
      : '';
    return `
      <tr>
        <td><span class="pri pri-${r.cls}">${r.priority}</span></td>
        <td><div class="iq-user">${escapeHtml(inv.name)}</div><div class="iq-upn">${escapeHtml(inv.username || '—')}</div></td>
        <td class="iq-score">${r.score}${partial}</td>
        <td><span class="tier tier-${r.cls}">${r.tier}</span></td>
        <td>${escapeHtml(statusLabel[inv.assessmentClass] || inv.assessment || '—')}</td>
        <td>${mfa}</td>
        <td class="iq-latest">${escapeHtml(latest)}</td>
      </tr>`;
  }).join('');
}

function renderSummary(report) {
  const { agg, investigations } = report;
  const compromiseTxt = agg.compromised > 0
    ? `<span class="hl-crimson">${agg.compromised} potential compromise event${agg.compromised === 1 ? '' : 's'}</span> requiring immediate response`
    : `<span class="hl-green">no confirmed compromise</span> events`;

  const successTxt = agg.anyPostRiskSuccess
    ? `Some risky-flagged users authenticated successfully after the detection window opened.`
    : `No risky-flagged user completed a successful authentication after the risky event was logged.`;

  $('summaryNarrative').innerHTML = `
    <p>This investigation correlates <strong class="hl-cyan">${agg.totalRiskyEvents}</strong> risky sign-in event${agg.totalRiskyEvents === 1 ? '' : 's'} against
    <strong class="hl-cyan">${agg.totalInteractiveSignins.toLocaleString()}</strong> interactive authentication records spanning a ${windowLabel(agg)}${windowRange(agg)}.
    A unique population of <strong>${agg.totalRiskyUsers}</strong> user${agg.totalRiskyUsers === 1 ? '' : 's'} was placed in an "At risk" state during the period.</p>

    <p style="margin-top:14px">Cross-referencing the risky telemetry with completed authentication outcomes reveals
    ${compromiseTxt},
    <span class="hl-cyan">${agg.mitigated} case${agg.mitigated === 1 ? '' : 's'} where controls operated effectively</span>
    (MFA challenge, Conditional Access interrupt, or credential rejection), and
    <span class="hl-amber">${agg.review} case${agg.review === 1 ? '' : 's'}</span> requiring human review.</p>

    <p style="margin-top:14px">${successTxt}
    A <strong>"risky sign-in" classification alone is not evidence of compromise</strong> — only a downstream <em>Success</em> status from
    the flagged source constitutes a confirmed unauthorized authentication.</p>
  `;

  const overallVerdict = agg.compromised > 0
    ? { tag: 'CRITICAL', cls: 'crit', val: 'COMPROMISE', unit: 'INDICATED' }
    : agg.review > 0
    ? { tag: 'CAUTION', cls: 'warn', val: 'REVIEW', unit: 'REQUIRED' }
    : { tag: 'NOMINAL', cls: 'ok', val: 'CONTROLS', unit: 'HELD' };

  $('verdictStack').innerHTML = `
    <div class="verdict-card ${overallVerdict.cls}">
      <div class="v-bar"></div>
      <div class="v-tag">${overallVerdict.tag}</div>
      <div class="v-label">Overall posture</div>
      <div class="v-value">${overallVerdict.val}<span class="unit">${overallVerdict.unit}</span></div>
    </div>
    <div class="verdict-card info">
      <div class="v-bar"></div>
      <div class="v-tag">SCOPE</div>
      <div class="v-label">Users at risk</div>
      <div class="v-value">${agg.totalRiskyUsers}<span class="unit">of ${agg.totalRiskyEvents} events</span></div>
    </div>
    <div class="verdict-card ${agg.compromised > 0 ? 'crit' : 'ok'}">
      <div class="v-bar"></div>
      <div class="v-tag">${agg.compromised > 0 ? 'BREACH' : 'CLEAN'}</div>
      <div class="v-label">Confirmed compromise</div>
      <div class="v-value">${agg.compromised}<span class="unit">account${agg.compromised === 1 ? '' : 's'}</span></div>
    </div>
    <div class="verdict-card info">
      <div class="v-bar"></div>
      <div class="v-tag">MITIGATED</div>
      <div class="v-label">Controls held</div>
      <div class="v-value">${agg.mitigated}<span class="unit">account${agg.mitigated === 1 ? '' : 's'}</span></div>
    </div>
  `;
}

function statusBadge(status) {
  if (!status) return '<span class="badge none">—</span>';
  const s = String(status).toLowerCase();
  if (s.includes('success')) return '<span class="badge success">Success</span>';
  if (s.includes('interrupt')) return '<span class="badge interrupt">Interrupted</span>';
  if (s.includes('failure')) return '<span class="badge failure">Failure</span>';
  return `<span class="badge none">${escapeHtml(status)}</span>`;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// v12.3 — add/remove the "Background Activity" header so the table (and the
// DOM-driven 02_Investigation sheet) is unchanged when Source 04 is absent.
function syncBgHeader(on) {
  const headRow = document.querySelector('.inv thead tr');
  if (!headRow) return;
  let th = document.getElementById('thBgActivity');
  if (on && !th) {
    th = document.createElement('th');
    th.id = 'thBgActivity';
    th.textContent = 'Background Activity (risky-context)';
    headRow.insertBefore(th, headRow.lastElementChild);   // just before Final Assessment
  } else if (!on && th) {
    th.remove();
  }
}
function bgCellText(i) {
  if (!i.bgSignins.length) return 'No background logs for user';
  if (!i.bgRiskyContext.length) return `0 of ${i.bgSignins.length} matched`;
  const last = i.bgRiskyContext[i.bgRiskyContext.length - 1];
  const succ = i.bgRiskyContext.filter(s => /success/i.test(s.status || '')).length;
  const fail = i.bgRiskyContext.filter(s => /failure/i.test(s.status || '')).length;
  return `${i.bgRiskyContext.length} of ${i.bgSignins.length} matched · ${succ} success / ${fail} failure · latest: ${last.status || '—'}`
    + (i.bgCompromise ? ' · TOKEN FROM RISKY IP' : '');
}

function renderTable(report) {
  const tbody = $('invTbody');
  tbody.innerHTML = '';
  const bgOn = !!report.agg.bgLoaded;
  syncBgHeader(bgOn);
  for (const i of report.investigations) {
    const first = i.first;
    const last = i.last;
    const auth = first?.authMethod || first?.mfaMethod || 'Single-factor authentication';
    const failure = last?.failure || first?.failure || '—';
    const posture = i.authPosture || { tier: 'unknown', label: 'Not in registry' };
    const postureHtml = state.authMap
      ? `<span class="posture-mini tier-${posture.tier}" title="${escapeHtml(i.auth?.defaultMethod || posture.label)}">${escapeHtml(posture.label)}</span>`
      : '<span style="color:var(--text-muted);font-family:var(--font-mono);font-size:11px">—</span>';

    // Unique IPs / Locations tied to risky-context results (Failure / Interrupted / Success).
    // Using ' | ' as separator because Location values contain embedded commas
    // (e.g. "Montreal, Quebec, CA"), so a plain ', ' join would be ambiguous.
    const uniqIps  = [...new Set(i.riskyContextSignins.map(s => s.ip).filter(Boolean))];
    const uniqLocs = [...new Set(i.riskyContextSignins.map(s => s.location).filter(Boolean))];
    const ipCell  = uniqIps.length  ? uniqIps.join(' | ')  : '—';
    const locCell = uniqLocs.length ? uniqLocs.join(' | ') : '—';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="name-cell"><div class="full">${escapeHtml(i.name)}</div></td>
      <td>${escapeHtml(i.username || '—')}</td>
      <td><span class="badge risky">YES</span></td>
      <td>${statusBadge(first?.status)}</td>
      <td>${escapeHtml(auth)}</td>
      <td>${postureHtml}</td>
      <td>${statusBadge(last?.status)}</td>
      <td style="max-width:180px;white-space:normal;line-height:1.45;font-family:var(--font-mono);font-size:11.5px;">${escapeHtml(ipCell)}</td>
      <td style="max-width:220px;white-space:normal;line-height:1.45;font-family:var(--font-mono);font-size:11.5px;">${escapeHtml(locCell)}</td>
      <td style="max-width:240px;white-space:normal;line-height:1.45;">${escapeHtml(failure === '—' || !failure ? '—' : failure.length > 80 ? failure.slice(0, 78) + '…' : failure)}</td>${bgOn ? `
      <td style="max-width:240px;white-space:normal;line-height:1.45;font-family:var(--font-mono);font-size:11.5px;${i.bgCompromise ? 'color:var(--crimson);font-weight:700;' : ''}">${escapeHtml(bgCellText(i))}</td>` : ''}
      <td><span class="assessment ${i.assessmentClass}">${i.assessment}</span></td>
    `;
    tbody.appendChild(tr);
  }
  if (!report.investigations.length) {
    tbody.innerHTML = `<tr><td colspan="${bgOn ? 12 : 11}" class="empty-state">No "At risk" users found in the risky sign-ins log.</td></tr>`;
  }
}

// ============================================================
// Risky IP / Location registry — aggregate view across all affected users
// ============================================================
function renderIpRegistry(report) {
  const body = $('ipRegBody');
  if (!body) return;

  // Aggregate every risky event across all investigations.
  // Key = IP + '|' + Location so distinct location strings per IP are preserved.
  const map = new Map();
  for (const inv of report.investigations) {
    for (const ev of inv.riskyEvents) {
      const ip = ev.ip || '—';
      const loc = ev.location || '—';
      const key = ip + '|' + loc;
      if (!map.has(key)) {
        map.set(key, { ip, location: loc, hits: 0, users: new Set(), latest: null });
      }
      const e = map.get(key);
      e.hits += 1;
      e.users.add(inv.name);
      const ts = ev.date?.getTime() || 0;
      if (!e.latest || ts > e.latest.getTime()) e.latest = ev.date;
    }
  }

  const rows = [...map.values()].sort((a, b) => (b.latest?.getTime() || 0) - (a.latest?.getTime() || 0));

  if (!rows.length) {
    body.innerHTML = `<div class="empty-state" style="padding:20px;background:var(--bg-panel);border:1px solid var(--line);font-family:var(--font-mono);font-size:12.5px;color:var(--text-muted)">No risky detections found.</div>`;
    return;
  }

  const uniqueIps = new Set(rows.map(r => r.ip)).size;
  const uniqueLocs = new Set(rows.map(r => r.location)).size;
  const uniqueUsers = new Set([].concat(...rows.map(r => [...r.users]))).size;

  const rowsHtml = rows.map(r => `
    <tr>
      <td class="ip">${escapeHtml(r.ip)}</td>
      <td>${escapeHtml(r.location)}</td>
      <td class="hits">${r.hits}</td>
      <td class="users">${escapeHtml([...r.users].join(', '))}</td>
      <td style="color:var(--text-muted);white-space:nowrap">${escapeHtml(fmtDate(r.latest))}</td>
    </tr>
  `).join('');

  body.innerHTML = `
    <div style="font-family:var(--font-mono);font-size:12px;color:var(--text-muted);margin-bottom:12px;letter-spacing:0.05em">
      <strong style="color:var(--crimson);font-weight:700">${uniqueIps}</strong> unique IP${uniqueIps === 1 ? '' : 's'} ·
      <strong style="color:var(--crimson);font-weight:700">${uniqueLocs}</strong> unique location${uniqueLocs === 1 ? '' : 's'} ·
      <strong style="color:var(--crimson);font-weight:700">${uniqueUsers}</strong> affected user${uniqueUsers === 1 ? '' : 's'} ·
      <strong style="color:var(--crimson);font-weight:700">${rows.reduce((s, r) => s + r.hits, 0)}</strong> total risky detection${rows.reduce((s, r) => s + r.hits, 0) === 1 ? '' : 's'}
    </div>
    <div class="ip-registry">
      <table class="ipreg">
        <thead>
          <tr>
            <th style="width:16%">IP Address</th>
            <th>Location</th>
            <th style="width:9%;text-align:right">Hits</th>
            <th style="width:28%">Affected User(s)</th>
            <th style="width:18%">Latest Detection (UTC)</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;
}

function renderUserCards(report) {
  const wrap = $('userCards');
  wrap.innerHTML = '';
  report.investigations.forEach((inv, idx) => {
    const card = document.createElement('div');
    card.className = 'user-card';
    const idxStr = String(idx + 1).padStart(2, '0');

    const firstStatus = inv.first?.status || 'no risky-context sign-in';
    const firstAuth = inv.first?.authMethod || 'unspecified';
    const firstLoc  = inv.first?.location || 'unknown location';
    const firstIp   = inv.first?.ip || 'unknown IP';
    const lastStatus = inv.last?.status || 'n/a';
    const riskyLocs  = [...inv.riskyLocSet].join(', ') || '—';
    const riskyIps   = [...inv.riskyIpSet].join(', ') || '—';
    const totalSignins = inv.signins.length;
    const rcCount = inv.riskyContextSignins.length;
    const successRate = rcCount ? Math.round(100 * inv.counts.success / rcCount) : 0;

    const narrative = rcCount === 0 ? `
      The risky detection for <strong>${escapeHtml(inv.name)}</strong> originated from
      <strong>${escapeHtml(riskyLocs)}</strong> (IP <code>${escapeHtml(riskyIps)}</code>).
      Across the ${windowLabel(report.agg)}, <strong>${totalSignins}</strong> interactive sign-in${totalSignins === 1 ? '' : 's'}
      ${totalSignins === 1 ? 'was' : 'were'} recorded for this user, but <strong>none</strong> originated from a risky-flagged IP or Location.
      ${escapeHtml(inv.assessmentReason)}
    ` : `
      The risky detection for <strong>${escapeHtml(inv.name)}</strong> originated from
      <strong>${escapeHtml(riskyLocs)}</strong> (IP <code>${escapeHtml(riskyIps)}</code>).
      Of <strong>${totalSignins}</strong> total interactive sign-in${totalSignins === 1 ? '' : 's'} in the ${windowLabel(report.agg)},
      <strong>${rcCount}</strong> matched a risky IP or Location — ${inv.counts.success} successful, ${inv.counts.interrupt} interrupted, ${inv.counts.failure} failed
      (${successRate}% completion rate within risky context). The earliest risky-context attempt resolved with status
      <strong>${escapeHtml(firstStatus)}</strong> using <em>${escapeHtml(firstAuth)}</em> from
      <strong>${escapeHtml(firstLoc)}</strong>. The most recent risky-context activity ended with status
      <strong>${escapeHtml(lastStatus)}</strong>. ${escapeHtml(inv.assessmentReason)}
    `;

    // Show ALL events, no truncation — full forensic timeline
    const events = inv.timeline;
    const moreCount = 0;
    const timelineHtml = events.map(ev => {
      if (ev.kind === 'risky') {
        return `<div class="tl-event risky">
          <span class="ts">${escapeHtml(fmtDate(ev.date))}</span>
          <span class="status-text">⚠ RISKY DETECTION</span>
          <div class="meta">
            <span class="lbl">ip</span> ${escapeHtml(shortIp(ev.ip))}
            <span class="sep">·</span>
            <span class="lbl">loc</span> ${escapeHtml(ev.location || '—')}
            <span class="sep">·</span>
            <span class="lbl">state</span> ${escapeHtml(ev.riskState || '—')}
          </div>
        </div>`;
      }
      const stRaw = String(ev.status || '').toLowerCase();
      const stCls = stRaw.includes('success') ? 'success' : stRaw.includes('interrupt') ? 'interrupt' : stRaw.includes('failure') ? 'failure' : '';
      const stLabel = ev.status || '—';
      const fail = ev.failure && ev.failure !== 'Other.' ? `<span class="sep">·</span><span class="lbl">reason</span> ${escapeHtml(ev.failure.length > 70 ? ev.failure.slice(0,68)+'…' : ev.failure)}` : '';
      const bgTag = isBg(ev) ? ' <span class="bg-tag">BACKGROUND</span>' : '';   // v12.3
      const bgMeta = isBg(ev)
        ? (ev.tokenType ? `<span class="sep">·</span><span class="lbl">token</span> ${escapeHtml(ev.tokenType)}` : '')
          + (ev.mfaByToken === true ? '<span class="sep">·</span><span class="lbl">mfa</span> satisfied by earlier token' : '')
        : '';
      return `<div class="tl-event ${stCls}">
        <span class="ts">${escapeHtml(fmtDate(ev.date))}</span>
        <span class="status-text">${escapeHtml(stLabel)}${bgTag}</span>
        <div class="meta">
          <span class="lbl">ip</span> ${escapeHtml(shortIp(ev.ip))}
          <span class="sep">·</span>
          <span class="lbl">loc</span> ${escapeHtml(ev.location || '—')}
          <span class="sep">·</span>
          <span class="lbl">auth</span> ${escapeHtml(ev.authMethod || '—')}
          ${ev.app ? '<span class="sep">·</span><span class="lbl">app</span> ' + escapeHtml(ev.app) : ''}${bgMeta}
          ${fail}
        </div>
      </div>`;
    }).join('');

    // ---- Risky Detection Log (all rows from risky sign-ins CSV for this user) ----
    const detEventsSorted = [...inv.riskyEvents].sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0));
    const detectionRowsHtml = detEventsSorted.map(ev => `
      <tr>
        <td class="ts">${escapeHtml(fmtDate(ev.date))}</td>
        <td class="ip">${escapeHtml(ev.ip || '—')}</td>
        <td class="loc">${escapeHtml(ev.location || '—')}</td>
        <td><span class="risk-state">${escapeHtml(ev.riskState || 'At risk')}</span></td>
      </tr>
    `).join('');
    const detectionTable = `
      <div class="detlog-wrap">
        <table class="detlog">
          <thead>
            <tr>
              <th style="width:22%">Date (UTC)</th>
              <th style="width:20%">IP Address</th>
              <th>Location</th>
              <th style="width:12%">Risk State</th>
            </tr>
          </thead>
          <tbody>${detectionRowsHtml}</tbody>
        </table>
      </div>
    `;

    // ---- Risky-Context Sign-ins (interactive logs whose IP or Location matched a risky detection) ----
    const rcSortedDesc = [...inv.riskyContextSignins].sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0));
    const rcRowsHtml = rcSortedDesc.map(s => {
      const st = String(s.status || '').toLowerCase();
      const stCls = st.includes('success') ? 'status-succ' : st.includes('interrupt') ? 'status-intr' : st.includes('failure') ? 'status-fail' : '';
      const ipMatch = inv.riskyIpSet.has(s.ip);
      const locMatch = inv.riskyLocSet.has(s.location);
      const matchTag = ipMatch && locMatch ? 'IP+LOC' : ipMatch ? 'IP' : locMatch ? 'LOC' : '—';
      const reason = s.failure && s.failure !== 'Other.' ? s.failure : '';
      return `
        <tr class="row-risky-context">
          <td class="ts">${escapeHtml(fmtDate(s.date))}</td>
          <td class="${stCls}">${escapeHtml(s.status || '—')}</td>
          <td class="ip">${escapeHtml(s.ip || '—')}</td>
          <td class="loc">${escapeHtml(s.location || '—')}</td>
          <td>${escapeHtml(s.authMethod || '—')}</td>
          <td><span class="risk-state">${matchTag}</span></td>
          <td class="reason">${escapeHtml(reason || '—')}</td>
        </tr>
      `;
    }).join('');
    const rcTable = rcSortedDesc.length ? `
      <div class="detlog-wrap">
        <table class="detlog">
          <thead>
            <tr>
              <th style="width:16%">Date (UTC)</th>
              <th style="width:12%">Status</th>
              <th style="width:14%">IP Address</th>
              <th style="width:22%">Location</th>
              <th>Auth Method</th>
              <th style="width:9%">Match</th>
              <th>Failure Reason</th>
            </tr>
          </thead>
          <tbody>${rcRowsHtml}</tbody>
        </table>
      </div>
    ` : `<div class="empty-state" style="padding:14px;background:var(--bg-panel-2);border:1px solid var(--line);font-family:var(--font-mono);font-size:12px;color:var(--text-muted)">No interactive sign-ins matched a risky-flagged IP or Location.</div>`;

    // ---- v12.3 background (non-interactive) section ----
    let bgSectionHtml = '';
    if (report.agg.bgLoaded) {
      const bgRows = [...inv.bgRiskyContext].sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0));
      const bgRowsHtml = bgRows.map(s => {
        const st = String(s.status || '').toLowerCase();
        const cls = st.includes('success') ? 'status-succ' : st.includes('interrupt') ? 'status-intr' : st.includes('failure') ? 'status-fail' : '';
        const ipM = inv.riskyIpSet.has(s.ip), locM = inv.riskyLocSet.has(s.location);
        return `<tr class="row-risky-context">
          <td class="ts">${escapeHtml(fmtDate(s.date))}</td>
          <td class="${cls}">${escapeHtml(s.status || '—')}</td>
          <td class="ip">${escapeHtml(s.ip || '—')}</td>
          <td class="loc">${escapeHtml(s.location || '—')}</td>
          <td>${escapeHtml(s.app || '—')}</td>
          <td>${escapeHtml(s.tokenType || '—')}</td>
          <td><span class="risk-state">${ipM && locM ? 'IP+LOC' : ipM ? 'IP' : 'LOC'}</span></td>
          <td>${s.mfaByToken === true ? 'Yes' : s.mfaByToken === false ? 'No' : '—'}</td>
        </tr>`;
      }).join('');
      const bgSummary = inv.bgSignins.length
        ? `${inv.bgSignins.length} background (non-interactive) sign-in${inv.bgSignins.length === 1 ? '' : 's'} for this user; <strong>${inv.bgRiskyContext.length}</strong> matched a risky IP or Location.`
          + (inv.bgCompromise ? ' <strong style="color:var(--crimson)">A background token was accepted from an exact risky IP after the detection (possible token replay).</strong>' : '')
        : 'No background (non-interactive) sign-ins were found for this user in Source 04.';
      bgSectionHtml = `
        <div class="user-section-label">Background (non-interactive) sign-ins · risky-context only</div>
        <div class="narrative" style="margin-bottom:10px">${bgSummary}</div>
        ${bgRows.length ? `<div class="detlog-wrap"><table class="detlog"><thead><tr>
          <th style="width:16%">Date (UTC)</th><th style="width:10%">Status</th><th style="width:14%">IP Address</th>
          <th>Location</th><th>Application</th><th>Token Type</th><th style="width:8%">Match</th><th style="width:10%">MFA via token</th>
        </tr></thead><tbody>${bgRowsHtml}</tbody></table></div>` : ''}`;
    }

    const anoCells = [
      { lbl: 'Risky IP(s)', val: riskyIps },
      { lbl: 'Risky Location(s)', val: riskyLocs },
      { lbl: 'Risky-context matched', val: `${rcCount} of ${totalSignins} sign-in${totalSignins === 1 ? '' : 's'}` },
      { lbl: 'Distinct IPs seen', val: inv.ipSet.size },
      { lbl: 'Distinct Locations', val: inv.locSet.size },
      { lbl: 'MFA enforcement (risky-context)', val: inv.riskyContextSignins.some(s => /interrupt/i.test(s.status || '')) ? 'Triggered' : 'Not triggered' },
      { lbl: 'Conditional Access (risky-context)', val: (() => {
          const vals = new Set(inv.riskyContextSignins.map(s => s.ca).filter(Boolean));
          return [...vals].join(', ') || '—';
        })()
      },
    ];

    if (report.agg.bgLoaded) {                        // v12.3
      const tt = [...new Set(inv.bgRiskyContext.map(s => s.tokenType).filter(t => t && t !== 'None'))];
      anoCells.push(
        { lbl: 'Background risky-context', val: `${inv.bgRiskyContext.length} of ${inv.bgSignins.length}` },
        { lbl: 'Background token types (risky-context)', val: tt.join(', ') || '—' },
      );
    }

    // v12 — inferred MITRE technique tags for this user (heuristic)
    const mitreTagsHtml = inv.mitre.length ? `
      <div class="user-section-label">Inferred MITRE ATT&amp;CK techniques <span class="mitre-inline-note">· heuristic, analyst-confirm</span></div>
      <div class="mitre-tags">${inv.mitre.map(t =>
        `<span class="mitre-tag mitre-${MITRE_SEV[t.sev] || 'medium'}" title="${escapeHtml(t.note)}">${escapeHtml(t.id)} · ${escapeHtml(t.name)}</span>`).join('')}</div>` : '';

    // v12 — device / browser drilldown (degrades if columns absent)
    const devSrc = inv.riskyContextSignins.length ? inv.riskyContextSignins : inv.signins;
    const browsers = [...new Set(devSrc.map(s => s.browser).filter(Boolean))];
    const oses     = [...new Set(devSrc.map(s => s.os).filter(Boolean))];
    const deviceHtml = (browsers.length || oses.length) ? `
      <div class="user-section-label">Device &amp; browser (risky-context)</div>
      <div class="ano-grid">
        <div class="ano-item"><div class="lbl">Browser(s)</div><div class="val">${escapeHtml(browsers.join(', ') || '—')}</div></div>
        <div class="ano-item"><div class="lbl">Operating system(s)</div><div class="val">${escapeHtml(oses.join(', ') || '—')}</div></div>
      </div>` : `
      <div class="user-section-label">Device &amp; browser</div>
      <div class="empty-state" style="padding:12px;background:var(--bg-panel-2);border:1px solid var(--line);font-family:var(--font-mono);font-size:12px;color:var(--text-muted)">Not present in source data.</div>`;

    card.innerHTML = `
      <div class="user-head">
        <div class="user-head-left">
          <div class="user-index">${idxStr}</div>
          <div class="user-id">
            <div class="nm">${escapeHtml(inv.name)}</div>
            <div class="em">${escapeHtml(inv.username || 'username not located in interactive logs')}</div>
          </div>
        </div>
        <div class="user-head-right">
          <span class="assessment ${inv.assessmentClass}">${inv.assessment}</span>
          <svg class="caret" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"/></svg>
        </div>
      </div>
      <div class="user-body">
        <div class="user-body-inner">
          ${state.authMap ? renderUserAuthPostureBlock(inv) : ''}
          <div class="user-section-label">Narrative</div>
          <div class="narrative">${narrative}</div>

          <div class="user-section-label">Risky detection log · ${detEventsSorted.length} event${detEventsSorted.length === 1 ? '' : 's'} from admin-center risky sign-ins export</div>
          ${detectionTable}

          <div class="user-section-label">Risky-context sign-ins · ${rcSortedDesc.length} interactive log${rcSortedDesc.length === 1 ? '' : 's'} matched a risky IP or Location</div>
          ${rcTable}${bgSectionHtml}

          <div class="user-section-label">Chronological timeline · all ${events.length} event${events.length === 1 ? '' : 's'} (risky detections + interactive sign-ins)</div>
          <div class="timeline">${timelineHtml || '<div class="empty-state">No events</div>'}</div>

          <div class="user-section-label">Anomaly &amp; control indicators</div>
          <div class="ano-grid">
            ${anoCells.map(c => `
              <div class="ano-item">
                <div class="lbl">${escapeHtml(c.lbl)}</div>
                <div class="val">${escapeHtml(String(c.val))}</div>
              </div>
            `).join('')}
          </div>
          ${mitreTagsHtml}
          ${deviceHtml}
        </div>
      </div>
    `;

    card.querySelector('.user-head').addEventListener('click', () => {
      card.classList.toggle('open');
    });
    if (idx === 0) card.classList.add('open');
    wrap.appendChild(card);
  });
}

function renderUserAuthPostureBlock(inv) {
  const auth = inv.auth;
  const posture = inv.authPosture || { tier: 'unknown', label: 'Not in registry' };
  if (!auth) {
    return `
      <div class="user-section-label">Registered authentication methods</div>
      <div class="auth-posture-row">
        <div class="auth-default">
          <span class="lbl">Default sign-in method</span>
          <span style="color:var(--text-muted);font-weight:500">Not in auth-method export</span>
        </div>
        <div class="auth-chips">
          <span style="color:var(--text-muted);font-family:var(--font-mono);font-size:12px">
            Re-run Bulk-ExtractAuthMethods.ps1 with this UPN to populate.
          </span>
        </div>
        <span class="posture-tier tier-unknown">${escapeHtml(posture.label)}</span>
      </div>`;
  }
  const chips = authChips(auth);
  const chipsHtml = chips.map(c =>
    `<span class="auth-chip ${c.cls}">${escapeHtml(c.label)}</span>`
  ).join('') || '<span style="color:var(--text-muted);font-family:var(--font-mono);font-size:12px">No methods registered</span>';

  return `
    <div class="user-section-label">Registered authentication methods</div>
    <div class="auth-posture-row">
      <div class="auth-default">
        <span class="lbl">Default sign-in method</span>
        ${escapeHtml(auth.defaultMethod)} <span style="font-weight:400;color:var(--text-muted);font-size:11px">· ${auth.methodCount} method${auth.methodCount === 1 ? '' : 's'} registered</span>
      </div>
      <div class="auth-chips">${chipsHtml}</div>
      <span class="posture-tier tier-${posture.tier}">${escapeHtml(posture.label)}</span>
    </div>`;
}

function renderAuthPostureSection(report) {
  const sec = $('sec-authposture');
  if (!state.authMap) {
    sec.style.display = 'none';
    // Sections: 01 summary, 02 table, 03 ip-registry, 04 per-user → 05 conclusion, 06 recs
    $('conclusionNum').textContent = '// 05';
    $('recsNum').textContent       = '// 06';
    return;
  }
  sec.style.display = 'block';
  // Sections: 01..04, 05 auth-posture → 06 conclusion, 07 recs
  $('conclusionNum').textContent = '// 06';
  $('recsNum').textContent       = '// 07';

  // Tenant-wide distribution across the entire uploaded auth-method registry
  const tiers = { resistant: 0, strong: 0, weak: 0, none: 0, unknown: 0 };
  for (const [, a] of state.authMap) {
    const t = classifyPosture(a).tier;
    tiers[t] = (tiers[t] || 0) + 1;
  }
  const total = state.authMap.size;
  const pct = n => total ? Math.round(100 * n / total) : 0;

  // Among risky-flagged users specifically — how many sit on weak factors?
  const riskyTiers = { resistant: 0, strong: 0, weak: 0, none: 0, unknown: 0 };
  for (const inv of report.investigations) {
    riskyTiers[inv.authPosture.tier]++;
  }
  const riskyTotal = report.investigations.length;
  const riskyPct = n => riskyTotal ? Math.round(100 * n / riskyTotal) : 0;

  // High-risk-posture users: compromised/review with weak/none auth posture
  const highRiskPosture = report.investigations.filter(i =>
    (i.assessmentClass === 'compromise' || i.assessmentClass === 'review') &&
    (i.authPosture.tier === 'weak' || i.authPosture.tier === 'none' || i.authPosture.tier === 'unknown')
  );

  const seg = (cls, n) => n > 0
    ? `<div class="seg-${cls}" style="flex:${n} 0 0">${pct(n) >= 6 ? n : ''}</div>`
    : '';
  const rseg = (cls, n) => n > 0
    ? `<div class="seg-${cls}" style="flex:${n} 0 0">${riskyPct(n) >= 8 ? n : ''}</div>`
    : '';

  $('authPostureBody').innerHTML = `
    <div class="auth-dist-grid">
      <div class="auth-dist-card tier-resistant">
        <div class="dist-lbl">Phishing-resistant</div>
        <div class="dist-val">${tiers.resistant}<span class="unit">/${total} · ${pct(tiers.resistant)}%</span></div>
      </div>
      <div class="auth-dist-card tier-strong">
        <div class="dist-lbl">Strong MFA</div>
        <div class="dist-val">${tiers.strong}<span class="unit">/${total} · ${pct(tiers.strong)}%</span></div>
      </div>
      <div class="auth-dist-card tier-weak">
        <div class="dist-lbl">Weak MFA (SMS/Voice/Email)</div>
        <div class="dist-val">${tiers.weak}<span class="unit">/${total} · ${pct(tiers.weak)}%</span></div>
      </div>
      <div class="auth-dist-card tier-none">
        <div class="dist-lbl">Password only / none</div>
        <div class="dist-val">${tiers.none}<span class="unit">/${total} · ${pct(tiers.none)}%</span></div>
      </div>
    </div>

    <div class="user-section-label" style="margin-top:18px">Tenant-wide auth posture distribution (${total} user${total === 1 ? '' : 's'})</div>
    <div class="auth-dist-bar">
      ${seg('resistant', tiers.resistant)}
      ${seg('strong',    tiers.strong)}
      ${seg('weak',      tiers.weak)}
      ${seg('none',      tiers.none)}
      ${seg('unknown',   tiers.unknown)}
    </div>
    <div class="auth-dist-legend">
      <span class="leg-item"><span class="leg-dot" style="background:var(--green)"></span>Phishing-resistant (${pct(tiers.resistant)}%)</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--cyan)"></span>Strong (${pct(tiers.strong)}%)</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--amber)"></span>Weak (${pct(tiers.weak)}%)</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--crimson)"></span>None / Password-only (${pct(tiers.none)}%)</span>
    </div>

    <div class="user-section-label" style="margin-top:24px">Auth posture among ${riskyTotal} risky-flagged user${riskyTotal === 1 ? '' : 's'}</div>
    <div class="auth-dist-bar">
      ${rseg('resistant', riskyTiers.resistant)}
      ${rseg('strong',    riskyTiers.strong)}
      ${rseg('weak',      riskyTiers.weak)}
      ${rseg('none',      riskyTiers.none)}
      ${rseg('unknown',   riskyTiers.unknown)}
    </div>
    <div class="auth-dist-legend">
      <span class="leg-item"><span class="leg-dot" style="background:var(--green)"></span>Phishing-resistant: ${riskyTiers.resistant}</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--cyan)"></span>Strong: ${riskyTiers.strong}</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--amber)"></span>Weak: ${riskyTiers.weak}</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--crimson)"></span>None: ${riskyTiers.none}</span>
      <span class="leg-item"><span class="leg-dot" style="background:var(--text-muted)"></span>Not in registry: ${riskyTiers.unknown}</span>
    </div>

    ${highRiskPosture.length > 0 ? `
      <div class="user-section-label" style="margin-top:24px;color:var(--crimson)">⚠ High-priority remediation candidates</div>
      <div class="narrative" style="border-left-color:var(--crimson)">
        <strong>${highRiskPosture.length} user${highRiskPosture.length === 1 ? '' : 's'}</strong>
        sit at the intersection of a risky-sign-in event and a weak authentication posture:
        ${highRiskPosture.map(u =>
          `<strong>${escapeHtml(u.name)}</strong> (${escapeHtml(u.authPosture.label)})`
        ).join(', ')}.
        These accounts should be prioritized for MFA strengthening (migrate to Microsoft Authenticator or FIDO2)
        ahead of the broader tenant rollout.
      </div>
    ` : ''}
  `;
}

function renderConclusion(report) {
  const { agg, investigations } = report;
  const compromisedUsers = investigations.filter(i => i.assessmentClass === 'compromise');
  const mitigatedUsers = investigations.filter(i => i.assessmentClass === 'mitigated');
  const reviewUsers = investigations.filter(i => i.assessmentClass === 'review');

  let body = '';

  if (compromisedUsers.length > 0) {
    body += `<p><strong>Evidence of compromise:</strong> ${compromisedUsers.length} account${compromisedUsers.length === 1 ? '' : 's'}
      (${compromisedUsers.map(u => escapeHtml(u.name)).join(', ')})
      recorded successful authentication from an IP address that was flagged as risky. These accounts should be treated as
      potentially compromised pending forensic verification.</p>`;
  } else {
    body += `<p><strong>Evidence of compromise:</strong> No conclusive evidence of account compromise was identified.
      No user with an "At risk" sign-in completed a successful authentication from the flagged IP address within the
      observation window. The risky-sign-in classification, in isolation, does not constitute compromise.</p>`;
  }

  if (mitigatedUsers.length > 0) {
    body += `<p><strong>Effectiveness of controls:</strong> Security controls operated as intended in
      ${mitigatedUsers.length} case${mitigatedUsers.length === 1 ? '' : 's'}.
      Authentication attempts from risky contexts were either interrupted by MFA / Conditional Access policy or rejected
      due to invalid credentials, preventing successful sign-in.</p>`;
  }

  if (reviewUsers.length > 0) {
    body += `<p><strong>Residual risk:</strong> ${reviewUsers.length} account${reviewUsers.length === 1 ? '' : 's'}
      (${reviewUsers.map(u => escapeHtml(u.name)).join(', ')})
      registered successful sign-ins after the risky event, though from different IP addresses than the flagged source.
      This pattern is consistent with legitimate user activity following an opportunistic risky-IP attempt, but should be
      validated directly with the affected user${reviewUsers.length === 1 ? '' : 's'} before being closed out.</p>`;
  } else if (compromisedUsers.length === 0) {
    body += `<p><strong>Residual risk:</strong> Residual risk is assessed as <strong>low</strong>. Continued
      monitoring is recommended for the affected user population over the next 14 days to confirm no delayed-stage
      activity emerges from the risky source IPs.</p>`;
  }

  body += `<p><strong>Overall posture:</strong> Of ${agg.totalRiskyUsers} risky-flagged users,
    ${agg.compromised} require incident response, ${agg.mitigated} were successfully blocked by controls,
    ${agg.review} require manual review, and ${agg.ok} showed no follow-on activity. The tenant's preventive
    control stack ${agg.compromised === 0 ? 'performed effectively against the risky sign-in attempts observed in this window' : 'requires immediate reinforcement given the observed compromise indicator'}.</p>`;

  $('conclusionBody').innerHTML = body;
}

function renderRecs(report) {
  const { agg, investigations } = report;
  const compromisedUsers = investigations.filter(i => i.assessmentClass === 'compromise');
  const reviewUsers = investigations.filter(i => i.assessmentClass === 'review');
  const recs = [];

  if (compromisedUsers.length > 0) {
    recs.push({
      title: 'Immediate password reset and session revocation',
      desc: `Force a password reset and revoke active refresh tokens for the ${compromisedUsers.length} potentially compromised account${compromisedUsers.length === 1 ? '' : 's'}: ${compromisedUsers.map(u => u.name).join(', ')}. Treat each as an active incident pending forensic review of mailbox rules, OAuth grants, and registered MFA devices.`
    });
  }

  if (reviewUsers.length > 0) {
    recs.push({
      title: 'User attestation for review-flagged accounts',
      desc: `Contact ${reviewUsers.map(u => u.name).join(', ')} directly to confirm recent sign-in activity. Validate the user's home location and device against the timeline before closing each case.`
    });
  }

  // ===== Auth-method-aware recommendations (only when auth registry uploaded) =====
  if (state.authMap) {
    const noMfaRisky = investigations.filter(i =>
      i.auth && !i.auth.authenticator && !i.auth.fido2 && !i.auth.windowsHello && !i.auth.sms && !i.auth.voice && !i.auth.email
    );
    if (noMfaRisky.length > 0) {
      recs.push({
        title: 'CRITICAL — Risky users with no MFA registered',
        desc: `${noMfaRisky.length} risky-flagged user${noMfaRisky.length === 1 ? '' : 's'} (${noMfaRisky.map(u => u.name).join(', ')}) ${noMfaRisky.length === 1 ? 'has' : 'have'} NO multifactor authentication method registered — only a password. These accounts are protected solely by password strength and are the highest residual-risk population in the tenant. Enforce immediate Authenticator app or FIDO2 enrollment via a Conditional Access policy that blocks access until MFA is registered.`
      });
    }

    const smsOnlyRisky = investigations.filter(i =>
      i.auth && (i.auth.sms || i.auth.voice) && !i.auth.authenticator && !i.auth.fido2 && !i.auth.windowsHello
    );
    if (smsOnlyRisky.length > 0) {
      recs.push({
        title: 'Migrate risky users from SMS/Voice to phishing-resistant MFA',
        desc: `${smsOnlyRisky.length} risky-flagged user${smsOnlyRisky.length === 1 ? '' : 's'} ${smsOnlyRisky.length === 1 ? 'is' : 'are'} protected only by SMS or voice-call MFA — methods widely known to be bypassable via SIM swap, SS7 interception, and adversary-in-the-middle phishing. Prioritize migration to Microsoft Authenticator (number matching) or FIDO2 security keys for: ${smsOnlyRisky.map(u => u.name).join(', ')}.`
      });
    }

    // Tenant-wide weak-MFA population
    let weakCount = 0, noneCount = 0;
    for (const [, a] of state.authMap) {
      const t = classifyPosture(a).tier;
      if (t === 'weak') weakCount++;
      if (t === 'none') noneCount++;
    }
    if (weakCount + noneCount > 0) {
      recs.push({
        title: 'Tenant-wide MFA modernization campaign',
        desc: `Across the uploaded authentication registry, ${noneCount} user${noneCount === 1 ? ' is' : 's are'} on password-only and ${weakCount} ${weakCount === 1 ? 'is' : 'are'} on weak MFA factors (SMS/Voice/Email OTP). Plan a phased migration to phishing-resistant methods (Authenticator with number matching, then FIDO2) with quarterly compliance milestones. Track via the auth-method extract on a recurring schedule.`
      });
    }
  }

  recs.push({
    title: 'Enforce MFA on all interactive sign-ins',
    desc: 'Ensure a Conditional Access policy requires MFA for every interactive authentication, with no Single-Factor authentication legacy exceptions. The investigation surfaced multiple interactive sign-ins that resolved with a "Single-factor authentication" requirement — these represent the largest residual exposure surface in the current control posture.'
  });

  recs.push({
    title: 'Conditional Access — risky sign-in policy',
    desc: 'Configure a Conditional Access policy that blocks access (or requires MFA + compliant device) when the sign-in risk level is Medium or High, using Identity Protection signals. This shifts the response from detection to prevention.'
  });

  recs.push({
    title: 'Geo-velocity and impossible-travel alerting',
    desc: 'The risky telemetry shows authentication attempts originating from geographically dispersed IPs (US, EU, APAC). Enable atypical-travel and unfamiliar-sign-in alerts in Microsoft Entra Identity Protection and route them to the SOC queue with a 4-hour response SLA.'
  });

  recs.push({
    title: 'Continuous monitoring — 14-day watchlist',
    desc: `Add the ${agg.totalRiskyUsers} risky-flagged users to a 14-day watchlist with elevated logging. Re-run this correlation analysis daily for the duration of the watchlist period to surface any delayed compromise indicators.`
  });

  recs.push({
    title: 'Phishing-resistant MFA migration',
    desc: 'Where SMS or voice MFA is still in use, migrate users to phishing-resistant methods (FIDO2 security keys, Windows Hello, or certificate-based authentication). SMS-based MFA was bypassed in 70%+ of cloud-account takeovers in recent industry reporting.'
  });

  recs.push({
    title: 'User awareness — risky sign-in context',
    desc: 'Issue a targeted communication to affected users explaining that their account was the subject of a risky-sign-in attempt, what signals triggered it, and what to do if they see unfamiliar MFA prompts. Empowered users are the final layer of the prevention stack.'
  });

  const list = $('recList');
  list.innerHTML = '';
  recs.forEach((r, i) => {
    const li = document.createElement('li');
    li.className = 'rec-item';
    li.innerHTML = `
      <div class="rec-num">// ${String(i + 1).padStart(2, '0')}</div>
      <div class="rec-title">${escapeHtml(r.title)}</div>
      <div class="rec-desc">${escapeHtml(r.desc)}</div>
    `;
    list.appendChild(li);
  });
}

// ========== EXPORT FULL REPORT TO EXCEL ==========
// ============================================================
// v12.2 — Per-User Timeline export (SEPARATE button; does not touch the engine
// or the full-report export). One dedicated worksheet per risky user, built
// from that user's existing inv.timeline (risky detections + all interactive
// sign-ins, chronological), plus a 00_Index sheet for navigation.
// ============================================================
function safeSheetName(name, used) {
  // Excel: max 31 chars, no  [ ] : * ? / \  , must be unique, not blank.
  let base = String(name || 'user').replace(/[\[\]\:\*\?\/\\]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!base) base = 'user';
  base = base.slice(0, 31);
  let candidate = base, n = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = '~' + n++;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function exportPerUserTimelines() {
  const reportDiv = $('report');
  if (!reportDiv.classList.contains('active') || !state.report || !state.report.investigations.length) {
    showToast('⛔ No investigation data to export. Please run the analysis first.');
    return;
  }
  try {
    const { investigations } = state.report;
    const workbook = XLSX.utils.book_new();
    const used = new Set();
    // Pre-assign a safe sheet name per user so the index can reference it.
    const assigned = investigations.map(inv => ({ inv, sheet: safeSheetName(inv.name, used) }));
    const bgOn = !!state.report.agg.bgLoaded;   // v12.3 — extra columns only when Source 04 loaded

    // ---- 00_Index (navigation) ----
    const indexRows = assigned.map(({ inv, sheet }) => ({
      'User': inv.name,
      'Username': inv.username || '',
      'Worksheet': sheet,
      'Risk Score': inv.risk ? inv.risk.score : '',
      'Risk Tier': inv.risk ? inv.risk.tier : '',
      'Final Assessment': inv.assessment || '',
      'Risky Detections': inv.riskyEvents.length,
      'Timeline Events': inv.timeline.length,
      ...(bgOn ? { 'Background Risky-Context': inv.bgRiskyContext.length, 'Background Total': inv.bgSignins.length } : {}),
    }));
    const indexSheet = XLSX.utils.json_to_sheet(indexRows);
    indexSheet['!cols'] = [{wch:24},{wch:30},{wch:26},{wch:11},{wch:12},{wch:24},{wch:16},{wch:16}];
    if (bgOn) indexSheet['!cols'].push({wch:24},{wch:18});
    XLSX.utils.book_append_sheet(workbook, indexSheet, '00_Index');

    // ---- One sheet per user: full chronological timeline ----
    for (const { inv, sheet } of assigned) {
      const events = [...inv.timeline].sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
      const rows = events.map(ev => {
        const isRisky = ev.kind === 'risky';
        const ipMatch  = inv.riskyIpSet.has(ev.ip);
        const locMatch = inv.riskyLocSet.has(ev.location);
        return {
          'Date (UTC)':        fmtDate(ev.date),
          'Event Type':        isRisky ? 'Risky detection' : (isBg(ev) ? 'Sign-in (background)' : 'Sign-in'),
          'Status':            isRisky ? (ev.riskState || 'At risk') : (ev.status || ''),
          'IP Address':        ev.ip || '',
          'Location':          ev.location || '',
          'Application':       isRisky ? '' : (ev.app || ''),
          'Auth Method':       isRisky ? '' : (ev.authMethod || ''),
          'Conditional Access':isRisky ? '' : (ev.ca || ''),
          'Browser':           isRisky ? '' : (ev.browser || ''),
          'Operating System':  isRisky ? '' : (ev.os || ''),
          'Risky Match':       (ipMatch && locMatch) ? 'IP+LOC' : ipMatch ? 'IP' : locMatch ? 'LOC' : '',
          'Sign-in Error Code':isRisky ? '' : (ev.errCode || ''),
          'Failure Reason':    isRisky ? '' : (ev.failure || ''),
          ...(bgOn ? {
            'Token Type':            isBg(ev) ? (ev.tokenType || '') : '',
            'ASN':                   isBg(ev) ? (ev.asn || '') : '',
            'Device Compliant':      isBg(ev) ? (ev.compliant || '') : '',
            'MFA via Earlier Token': isBg(ev) ? (ev.mfaByToken === true ? 'Yes' : ev.mfaByToken === false ? 'No' : '') : '',
          } : {}),
        };
      });
      const ws = rows.length
        ? XLSX.utils.json_to_sheet(rows)
        : XLSX.utils.aoa_to_sheet([['No timeline events recorded for this user.']]);
      ws['!cols'] = [
        {wch:22},{wch:16},{wch:14},{wch:18},{wch:34},{wch:22},{wch:26},
        {wch:18},{wch:16},{wch:16},{wch:11},{wch:18},{wch:40}
      ];
      if (bgOn) ws['!cols'].push({wch:22},{wch:10},{wch:16},{wch:20});
      XLSX.utils.book_append_sheet(workbook, ws, sheet);
    }

    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    XLSX.writeFile(workbook, `${socFilePrefix()}SOC_PerUser_Timelines_${timestamp}.xlsx`);
    showToast(`✓ Exported ${assigned.length} per-user timeline sheet${assigned.length === 1 ? '' : 's'} (+ index).`);
  } catch (err) {
    console.error(err);
    showToast('⛔ Per-user timeline export failed: ' + err.message);
  }
}

function exportInvestigationTable() {
  const reportDiv = $('report');
  if (!reportDiv.classList.contains('active') || !state.report || !state.report.investigations.length) {
    showToast('⛔ No investigation data to export. Please run the analysis first.');
    return;
  }

  try {
    const { investigations, agg } = state.report;
    const workbook = XLSX.utils.book_new();

    // -----------------------------------------------------------------------
    // Sheet 01 — Summary
    // -----------------------------------------------------------------------
    // Aggregate risky-detection numbers across every affected user
    const allRiskyEvents = investigations.flatMap(i => i.riskyEvents);
    const uniqueRiskyIps  = new Set(allRiskyEvents.map(e => e.ip).filter(Boolean));
    const uniqueRiskyLocs = new Set(allRiskyEvents.map(e => e.location).filter(Boolean));

    const summaryRows = [
      ['SOC Risky Sign-in Investigation — Executive Summary'],
      [],
      ['Report generated (UTC)', new Date().toISOString().replace('T', ' ').slice(0, 19) + 'Z'],
      ['Risky sign-ins source',  state.riskyFilename || '—'],
      ['Interactive logs source', state.allFilename || '—'],
      ['Auth-methods source',    state.authFilename || 'not loaded'],
      ['Observation window',     agg.windowHas
        ? `${agg.windowDays}-day (${fmtDate(agg.windowStart).slice(0,10)} → ${fmtDate(agg.windowEnd).slice(0,10)})`
        : '—'],
      [],
      ['— Aggregate counts —'],
      ['Total risky detections',       agg.totalRiskyEvents],
      ['Distinct affected users',      agg.totalRiskyUsers],
      ['Unique risky IPs',             uniqueRiskyIps.size],
      ['Unique risky locations',       uniqueRiskyLocs.size],
      ['Interactive sign-ins ingested', agg.totalInteractiveSignins],
      [],
      ['— Assessment breakdown —'],
      ['Possible compromise',      agg.compromised],
      ['Controls effective / auth failed (mitigated)', agg.mitigated],
      ['Review required',          agg.review],
      ['No post-risk activity',    agg.ok],
    ];
    if (agg.bgLoaded) {                               // v12.3
      summaryRows.push(
        [],
        ['— Background (non-interactive) layer —'],
        ['Non-interactive source',                 state.noniFilename || '—'],
        ['Non-interactive auth-details source',    state.noniAuthFilename || 'not loaded'],
        ['Background rows ingested (with user)',   agg.bgTotal],
        ['Background rows for risky users',        agg.bgKept],
        ['Background risky-context matches',       agg.bgRiskyContext],
        ['Users escalated by background activity', agg.bgEscalations],
        ['Token-from-risky-IP (possible replay)',  agg.bgCompromised],
        ['Background rows: MFA satisfied by token', agg.bgAuthLoaded ? `${agg.bgMfaByToken} of ${agg.bgMfaChecked} matched to auth details` : 'auth details not loaded'],
      );
    }
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    summarySheet['!cols'] = [{wch:42}, {wch:36}];
    // Bold-ish visual hint via merged header
    summarySheet['!merges'] = [{s:{r:0,c:0}, e:{r:0,c:1}}];
    XLSX.utils.book_append_sheet(workbook, summarySheet, '01_Summary');

    // -----------------------------------------------------------------------
    // Sheet 02 — Investigation (mirrors the top on-screen table)
    // -----------------------------------------------------------------------
    const invTableEl = document.querySelector('.inv');
    if (invTableEl) {
      const invSheet = XLSX.utils.table_to_sheet(invTableEl.cloneNode(true), { raw: true });
      invSheet['!cols'] = [
        {wch:24}, {wch:32}, {wch:14}, {wch:18}, {wch:26}, {wch:20}, {wch:18},
        {wch:38}, {wch:44}, {wch:44}, {wch:22}
      ];
      if (agg.bgLoaded) invSheet['!cols'].splice(10, 0, { wch: 46 });   // v12.3 Background Activity
      XLSX.utils.book_append_sheet(workbook, invSheet, '02_Investigation');
    }

    // -----------------------------------------------------------------------
    // Sheet 03 — IP · Location Registry (aggregate)
    // -----------------------------------------------------------------------
    const ipRegMap = new Map();
    for (const inv of investigations) {
      for (const ev of inv.riskyEvents) {
        const key = (ev.ip || '—') + '|' + (ev.location || '—');
        if (!ipRegMap.has(key)) {
          ipRegMap.set(key, { ip: ev.ip || '—', location: ev.location || '—', hits: 0, users: new Set(), latest: null });
        }
        const e = ipRegMap.get(key);
        e.hits += 1;
        e.users.add(inv.name);
        const ts = ev.date?.getTime() || 0;
        if (!e.latest || ts > e.latest.getTime()) e.latest = ev.date;
      }
    }
    const ipRegRows = [...ipRegMap.values()]
      .sort((a, b) => (b.latest?.getTime() || 0) - (a.latest?.getTime() || 0))
      .map(r => ({
        'IP Address':               r.ip,
        'Location':                 r.location,
        'Hits':                     r.hits,
        'Affected User(s)':         [...r.users].join(', '),
        'Latest Detection (UTC)':   fmtDate(r.latest),
      }));
    const ipRegSheet = XLSX.utils.json_to_sheet(ipRegRows);
    ipRegSheet['!cols'] = [{wch:18},{wch:44},{wch:8},{wch:38},{wch:22}];
    XLSX.utils.book_append_sheet(workbook, ipRegSheet, '03_IP_Location_Registry');

    // -----------------------------------------------------------------------
    // Sheet 04 — Risky Detection Log (every row from the risky sign-ins CSV, per user)
    // -----------------------------------------------------------------------
    const detectionRows = [];
    for (const inv of investigations) {
      const sorted = [...inv.riskyEvents].sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
      for (const ev of sorted) {
        detectionRows.push({
          'Date (UTC)':  fmtDate(ev.date),
          'User':        inv.name,
          'Username':    inv.username || '',
          'IP Address':  ev.ip || '',
          'Location':    ev.location || '',
          'Risk State':  ev.riskState || 'At risk',
        });
      }
    }
    const detectionSheet = XLSX.utils.json_to_sheet(detectionRows);
    detectionSheet['!cols'] = [{wch:22},{wch:24},{wch:32},{wch:18},{wch:44},{wch:14}];
    XLSX.utils.book_append_sheet(workbook, detectionSheet, '04_Risky_Detection_Log');

    // -----------------------------------------------------------------------
    // Sheet 05 — Risky-Context Sign-ins (interactive logs matching a risky IP/Location)
    // -----------------------------------------------------------------------
    const rcRows = [];
    for (const inv of investigations) {
      const sorted = [...inv.riskyContextSignins].sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
      for (const s of sorted) {
        const ipMatch  = inv.riskyIpSet.has(s.ip);
        const locMatch = inv.riskyLocSet.has(s.location);
        rcRows.push({
          'Date (UTC)':      fmtDate(s.date),
          'User':            inv.name,
          'Username':        inv.username || '',
          'Status':          s.status || '',
          'IP Address':      s.ip || '',
          'Location':        s.location || '',
          'Match Type':      ipMatch && locMatch ? 'IP+LOC' : ipMatch ? 'IP' : locMatch ? 'LOC' : '',
          'Auth Method':     s.authMethod || '',
          'MFA Method':      s.mfaMethod || '',
          'MFA Result':      s.mfaResult || '',
          'Application':     s.app || '',
          'Conditional Access': s.ca || '',
          'Browser':         s.browser || '',
          'Operating System': s.os || '',
          'Sign-in Error Code': s.errCode || '',
          'Failure Reason':  s.failure || '',
        });
      }
    }
    const rcSheet = XLSX.utils.json_to_sheet(rcRows);
    rcSheet['!cols'] = [
      {wch:22},{wch:24},{wch:32},{wch:14},{wch:18},{wch:36},{wch:12},
      {wch:28},{wch:22},{wch:18},{wch:22},{wch:22},{wch:16},{wch:18},{wch:18},{wch:42}
    ];
    XLSX.utils.book_append_sheet(workbook, rcSheet, '05_RiskyContext_Signins');

    // Sheet 05b — Background (non-interactive) risky-context sign-ins (v12.3, only when loaded)
    if (agg.bgLoaded) {
      const bgRows = [];
      for (const inv of investigations) {
        for (const s of inv.bgRiskyContext) {
          const ipMatch = inv.riskyIpSet.has(s.ip), locMatch = inv.riskyLocSet.has(s.location);
          bgRows.push({
            'Date (UTC)':            fmtDate(s.date),
            'User':                  inv.name,
            'Username':              inv.username || '',
            'Status':                s.status || '',
            'IP Address':            s.ip || '',
            'Location':              s.location || '',
            'Match Type':            ipMatch && locMatch ? 'IP+LOC' : ipMatch ? 'IP' : locMatch ? 'LOC' : '',
            'Application':           s.app || '',
            'Resource':              s.resource || '',
            'Token Type':            s.tokenType || '',
            'Client App':            s.clientApp || '',
            'ASN':                   s.asn || '',
            'Device Compliant':      s.compliant || '',
            'Managed':               s.managed || '',
            'MFA via Earlier Token': s.mfaByToken === true ? 'Yes' : s.mfaByToken === false ? 'No' : '',
            'Sign-in Error Code':    s.errCode || '',
            'Failure Reason':        s.failure || '',
          });
        }
      }
      const bgSheet = bgRows.length ? XLSX.utils.json_to_sheet(bgRows)
        : XLSX.utils.aoa_to_sheet([['No background sign-ins matched a risky IP or location.']]);
      bgSheet['!cols'] = [{wch:22},{wch:24},{wch:30},{wch:12},{wch:18},{wch:34},{wch:10},{wch:26},
        {wch:30},{wch:22},{wch:24},{wch:10},{wch:14},{wch:10},{wch:18},{wch:16},{wch:40}];
      XLSX.utils.book_append_sheet(workbook, bgSheet, '05b_Background_Signins');
    }

    // -----------------------------------------------------------------------
    // Sheet 06 — Full Timeline (all events: risky detections + all interactive sign-ins)
    // -----------------------------------------------------------------------
    const timelineRows = [];
    for (const inv of investigations) {
      const sorted = [...inv.timeline].sort((a, b) => (a.date?.getTime() || 0) - (b.date?.getTime() || 0));
      for (const ev of sorted) {
        const isRisky   = ev.kind === 'risky';
        const ipMatch   = !isRisky && inv.riskyIpSet.has(ev.ip);
        const locMatch  = !isRisky && inv.riskyLocSet.has(ev.location);
        timelineRows.push({
          'Date (UTC)': fmtDate(ev.date),
          'User':       inv.name,
          'Event Type': isRisky ? 'RISKY DETECTION' : (isBg(ev) ? 'Background sign-in (non-interactive)' : 'Interactive sign-in'),
          'Status':     isRisky ? (ev.riskState || 'At risk') : (ev.status || ''),
          'IP Address': ev.ip || '',
          'Location':   ev.location || '',
          'Risky-Context Match': isRisky ? '—' : (ipMatch && locMatch ? 'IP+LOC' : ipMatch ? 'IP' : locMatch ? 'LOC' : 'no'),
          'Auth Method': isRisky ? '' : (ev.authMethod || ''),
          'Application': isRisky ? '' : (ev.app || ''),
          'Failure Reason': isRisky ? '' : (ev.failure || ''),
        });
      }
    }
    const timelineSheet = XLSX.utils.json_to_sheet(timelineRows);
    timelineSheet['!cols'] = [
      {wch:22},{wch:24},{wch:20},{wch:16},{wch:18},{wch:36},{wch:16},{wch:26},{wch:22},{wch:42}
    ];
    XLSX.utils.book_append_sheet(workbook, timelineSheet, '06_Full_Timeline');

    // -----------------------------------------------------------------------
    // Sheet 07 — Auth Posture (only when auth-methods CSV was uploaded)
    // -----------------------------------------------------------------------
    if (state.authMap && investigations.length) {
      const authRows = investigations.map(i => ({
        Name: i.name,
        UPN: i.username || '',
        FinalAssessment: i.assessment,
        AuthPosture: i.authPosture?.label || '',
        DefaultSignInMethod: i.auth?.defaultMethod || 'Not in registry',
        MethodCount: i.auth?.methodCount ?? '',
        FIDO2: i.auth?.fido2 ?? '',
        WindowsHello: i.auth?.windowsHello ?? '',
        Authenticator: i.auth?.authenticator ?? '',
        SMS: i.auth?.sms ?? '',
        VoiceCall: i.auth?.voice ?? '',
        EmailOTP: i.auth?.email ?? '',
        TAP: i.auth?.tap ?? '',
        Password: i.auth?.password ?? '',
      }));
      const authSheet = XLSX.utils.json_to_sheet(authRows);
      authSheet['!cols'] = [
        {wch:24},{wch:32},{wch:22},{wch:22},{wch:34},{wch:8},
        {wch:8},{wch:14},{wch:14},{wch:8},{wch:10},{wch:10},{wch:8},{wch:10}
      ];
      XLSX.utils.book_append_sheet(workbook, authSheet, '07_Auth_Posture');
    }

    // -----------------------------------------------------------------------
    // Sheet 08 — Recommendations (SOC action items surfaced by the engine)
    // -----------------------------------------------------------------------
    const recListEl = document.querySelectorAll('#recList .rec-item');
    if (recListEl && recListEl.length) {
      const recRows = [...recListEl].map((li, idx) => ({
        '#': idx + 1,
        'Recommendation': li.querySelector('.rec-title')?.textContent?.trim() || '',
        'Details':        li.querySelector('.rec-desc')?.textContent?.trim() || '',
      }));
      const recSheet = XLSX.utils.json_to_sheet(recRows);
      recSheet['!cols'] = [{wch:5},{wch:52},{wch:120}];
      XLSX.utils.book_append_sheet(workbook, recSheet, '08_Recommendations');
    }

    // -----------------------------------------------------------------------
    const timestamp = new Date().toISOString().slice(0,19).replace(/:/g, '-');
    XLSX.writeFile(workbook, `${socFilePrefix()}SOC_RiskyInvestigation_${timestamp}.xlsx`);
    const sheetCount = workbook.SheetNames.length;
    showToast(`✓ Exported full report — ${sheetCount} sheet${sheetCount === 1 ? '' : 's'}.`);
  } catch (err) {
    console.error(err);
    showToast('Export failed: ' + err.message);
  }
}

// ============================================================
// Reset
// ============================================================
function reset() {
  resetCharts();
  state.riskyData = null;
  state.allData = null;
  state.authData = null;
  state.authMap = null;
  state.noniData = null; state.noniFilename = null; state.noniMeta = null;  // v12.3
  state.noniAuthMap = null; state.noniAuthFilename = null;
  state.riskyFilename = null;
  state.allFilename = null;
  state.authFilename = null;
  state.report = null;
  $('pill-risky').classList.add('hidden');
  $('pill-all').classList.add('hidden');
  $('pill-auth').classList.add('hidden');
  $('pill-noni').classList.add('hidden'); $('pill-noniad').classList.add('hidden');   // v12.3
  $('card-risky').classList.remove('loaded');
  $('card-all').classList.remove('loaded');
  $('card-auth').classList.remove('loaded');
  $('card-noni').classList.remove('loaded'); $('card-noniad').classList.remove('loaded');
  $('file-risky').value = '';
  $('file-all').value = '';
  $('file-auth').value = '';
  $('file-noni').value = ''; $('file-noniad').value = '';
  $('report').classList.remove('active');
  $('sec-authposture').style.display = 'none';
  updateAnalyzeReady();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ============================================================
// Clock + session
// ============================================================
function tickClock() {
  const d = new Date();
  $('utcClock').textContent = d.toISOString().replace('T', ' ').slice(0, 19) + 'Z';
}

function makeSession() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += c[Math.floor(Math.random() * c.length)];
  $('sessionId').textContent = s;
}

// ============================================================
// Init
// ============================================================
let _socInitDone = false;
QBR.socInit = function () {
  if (_socInitDone) return;
  _socInitDone = true;
  bindDrop($('drop-risky'), $('file-risky'), 'risky');
  bindDrop($('drop-all'), $('file-all'), 'all');
  bindDrop($('drop-auth'), $('file-auth'), 'auth');
  bindDrop($('drop-noni'), $('file-noni'), 'noni');       // v12.3
  bindDrop($('drop-noniad'), $('file-noniad'), 'noniad'); // v12.3
  $('analyzeBtn').addEventListener('click', runInvestigation);
  $('resetBtn').addEventListener('click', reset);
  $('exportInvestBtn').addEventListener('click', exportInvestigationTable);
  $('exportUserTimelinesBtn').addEventListener('click', exportPerUserTimelines);
  const impBtn = $('soc-import-btn'), impFile = $('soc-import-file');
  if (impBtn && impFile) {
    impBtn.addEventListener('click', () => impFile.click());
    impFile.addEventListener('change', () => { socImportReport(impFile.files[0]); impFile.value = ''; });
  }
  // Upload-stage import card: always visible, no investigation needed first.
  const impDrop = $('drop-socimport'), impCard = $('file-socimport');
  if (impDrop && impCard) {
    impDrop.addEventListener('click', e => { if (e.target !== impCard) impCard.click(); });
    impDrop.addEventListener('dragover', e => { e.preventDefault(); impDrop.classList.add('dragging'); });
    impDrop.addEventListener('dragleave', e => { e.preventDefault(); impDrop.classList.remove('dragging'); });
    impDrop.addEventListener('drop', e => { e.preventDefault(); impDrop.classList.remove('dragging');
      const f = e.dataTransfer.files[0]; if (f) { socImportReport(f); } });
    impCard.addEventListener('change', () => { socImportReport(impCard.files[0]); impCard.value = ''; });
  }

  // v12.1 — collapsible sections. Click a section header to fold/unfold it.
  // Delegated so it also covers sections whose bodies render after analysis.
  $('report').addEventListener('click', e => {
    const h = e.target.closest('.sec-header');
    if (h && h.parentElement.classList.contains('section')) {
      h.parentElement.classList.toggle('collapsed');
    }
  });
  const allSections = () => document.querySelectorAll('#report .section');
  $('collapseAllBtn').addEventListener('click', () => allSections().forEach(s => s.classList.add('collapsed')));
  $('expandAllBtn').addEventListener('click', () => allSections().forEach(s => s.classList.remove('collapsed')));


  // Dashboard port: fetch the vendored xlsx source so the parse worker can
  // prepend it (same as the standalone's inlined copy). If fetch fails
  // (e.g. file://), the worker falls back gracefully to main-thread parsing.
  // v1.32.0: only over http(s) — on file:// fetch is blocked and logs console errors.
  try {
    if (/^https?:$/.test(location.protocol)) fetch('libs/xlsx.full.min.js').then(r => r.text()).then(t => {
      if (t && t.length > 1000) window.__XLSX_SRC__ = t;
    }).catch(() => {});
  } catch (e) {}
  makeSession();
  tickClock();
  setInterval(tickClock, 1000);
};

/* Dashboard port: tenant-coded export filenames.
 * Reads #soc-tenant (e.g. BSCS) and #soc-month (e.g. SEP); returns
 * "BSCS_SEP_" or "" when empty (falls back to legacy naming). */
function socFilePrefix() {
  const t = document.getElementById('soc-tenant');
  const m = document.getElementById('soc-month');
  const clean = v => String(v || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 12);
  const tc = t ? clean(t.value) : '', mo = m ? clean(m.value) : '';
  if (!tc && !mo) return '';
  return (tc ? tc + '_' : '') + (mo ? mo + '_' : '');
}

/* Dashboard port: import a previously exported SOC report (.xlsx) and
 * re-render the full interactive report without re-uploading source files.
 * Rebuilds the analyze() inputs from the export sheets:
 *   04_Risky_Detection_Log -> risky rows
 *   06_Full_Timeline       -> interactive sign-in rows (enriched from 05)
 *   05_RiskyContext_Signins -> username / MFA / CA / browser / OS enrichment
 * Then runs the real analyze() + renderReport(), so scores, charts and all
 * sections are recomputed by the existing engine. */
async function socImportReport(file) {
  if (!file) return;
  showToast('Reading exported report…');
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const byName = {};
    wb.SheetNames.forEach(n => { byName[n] = wb.Sheets[n]; });
    const sheet = n => {
      const ws = byName[n];
      if (!ws) return [];
      return XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
    };
    const det = sheet('04_Risky_Detection_Log');
    const ctx = sheet('05_RiskyContext_Signins');
    const tl  = sheet('06_Full_Timeline');
    if (!det.length || !tl.length) { showToast('Not a SOC report export: missing 04/06 sheets.'); return; }

    // Build enrichment lookup from 05: key = user|date|ip (lowercased)
    const enrich = new Map();
    const ekey = r => [r['User'], r['Date (UTC)'], r['IP Address']].map(v => String(v || '').toLowerCase().trim()).join('|');
    ctx.forEach(r => enrich.set(ekey(r), r));

    const risky = det.map(r => ({
      'Date (UTC)': r['Date (UTC)'], 'User': r['User'], 'Username': r['Username'],
      'IP address': r['IP Address'], 'Location': r['Location'], 'Risk state': r['Risk State'],
    }));
    const all = [];
    tl.forEach(r => {
      if (String(r['Event Type'] || '').toLowerCase().indexOf('sign-in') < 0) return;
      if (String(r['Event Type'] || '').toLowerCase().indexOf('background') >= 0) return; // v12.3 layer: skip on import
      const e = enrich.get([r['User'], r['Date (UTC)'], r['IP Address']].map(v => String(v || '').toLowerCase().trim()).join('|'));
      all.push({
        'Date (UTC)': r['Date (UTC)'], 'User': r['User'],
        'Username': e ? e['Username'] : null,
        'Status': r['Status'], 'IP address': r['IP Address'], 'Location': r['Location'],
        'Authentication method': r['Auth Method'], 'Application': r['Application'],
        'Failure reason': r['Failure Reason'],
        'Multifactor authentication auth method': e ? e['MFA Method'] : null,
        'Multifactor authentication result': e ? e['MFA Result'] : null,
        'Conditional Access': e ? e['Conditional Access'] : null,
        'Browser': e ? e['Browser'] : null, 'Operating System': e ? e['Operating System'] : null,
      });
    });
    if (!risky.length || !all.length) { showToast('No usable rows found in this file.'); return; }

    state.riskyData = risky; state.riskyFilename = file.name + ' (imported)';
    state.allData = all;     state.allFilename = file.name + ' (imported)';
    state.authData = null; state.authMap = null; state.noniData = null; state.noniAuthMap = null;
    // Tenant code back from filename: "<TENANT>_<MON>_SOC_..." -> prefill
    const m = /^([A-Z0-9]{1,12})_([A-Z0-9]{1,12})_SOC_/i.exec(file.name);
    const ti = document.getElementById('soc-tenant'), mi = document.getElementById('soc-month');
    if (m) { if (ti && !ti.value) ti.value = m[1].toUpperCase(); if (mi && !mi.value) mi.value = m[2].toUpperCase(); }
    updateAnalyzeReady();
    runInvestigation();
    showToast(`✓ Imported ${risky.length} detections + ${all.length} sign-ins from ${file.name}`);
    const pill = document.getElementById('pill-socimport');
    if (pill) { pill.classList.remove('hidden');
      pill.querySelector('.filename').textContent = file.name;
      pill.querySelector('.rows').textContent = `${risky.length + all.length} rows`;
    }
  } catch (err) {
    console.error(err);
    showToast('Import failed: ' + (err && err.message ? err.message : err));
  }
}

  // Exposed for testing / dashboard integration (IIFE lesson: host-called
  // functions must live on the shared namespace).
  QBR.socFilePrefix = socFilePrefix;
  QBR.socImportReport = socImportReport;
  // Test-only hook: lets a harness observe/wrap engine functions.
  QBR.socTest = function (name, fn) {
    if (name === "wrapAnalyze") { const o = analyze; analyze = function (r, a, n) { fn(r, a, n); return o(r, a, n); }; }
    if (name === "wrapRender") { const o = renderReport; renderReport = function (r) { fn(r); return o(r); }; }
  };
})();
