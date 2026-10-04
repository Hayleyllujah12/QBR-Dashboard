/* =============================================================================
 * excel-loader.js
 * Parses uploaded Microsoft 365 tenant workbook(s) into a normalized data model.
 * - Detects sheets by name (filename-independent).
 * - Cleans dirty categorical values (typos / casing).
 * - Builds an internal MASTER SCHOOL DIMENSION so every fact sheet
 *   (Risky / Security / Usage / Storage / Canva) resolves to one canonical school.
 * No network, no persistence. Everything runs in-memory in the browser.
 * ===========================================================================*/

var QBR = (window.QBR = window.QBR || {});

// Single source of truth for the app version. The header renders from this, and
// the regression harness asserts it matches the deployment folder name
// (QBR-Dashboard_<VERSION>). Per the versioning convention, in-place fixes are
// dated patches and DO NOT bump this — a new number is cut only for a deliberate
// major release, together with a new folder.
QBR.VERSION = "1.21.0";

/* ---------- small helpers ------------------------------------------------ */

// Canonical join key for a school name: uppercase, strip punctuation/marks,
// collapse whitespace. Used only for matching, never for display.
function schoolKey(name) {
  if (name == null) return "";
  return String(name)
    .replace(/[​-‏‪-‮﻿]/g, "") // stray unicode marks
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Repair UTF-8 text that was mis-decoded as Latin-1 / CP1252 (the classic
// "Ã±" -> "ñ", "Ã©" -> "é" corruption seen in some exported cells). Only touches
// strings that carry the tell-tale Ã / Â / â€ byte sequences, and only keeps the
// result when it re-decodes cleanly — a correctly-encoded name is returned as-is,
// so this can never corrupt good data.
function fixMojibake(s) {
  if (s == null) return s;
  s = String(s);
  if (!/[ÂÃ]|â/.test(s)) return s;
  try {
    var f = decodeURIComponent(escape(s));
    if (f && f.indexOf("�") < 0) return f;
  } catch (e) { /* not a valid UTF-8 byte sequence — leave unchanged */ }
  return s;
}
// Display form of a school name: trimmed + mojibake-repaired. The join key still
// comes from schoolKey() (which strips these marks anyway), so matching is
// unaffected — this only fixes what the user sees.
function dispName(s) { return fixMojibake(String(s == null ? "" : s).trim()); }

// Parse a storage cell like "43.71 TB" / "882.25 GB" / "650.77 GB" -> number of GB.
function toGB(val) {
  if (val == null || val === "") return null;
  const s = String(val).replace(/[​-‏‪-‮﻿,]/g, "").trim();
  const m = s.match(/([\d.]+)\s*(TB|GB|MB)?/i);
  if (!m) return null;
  let n = parseFloat(m[1]);
  if (isNaN(n)) return null;
  const unit = (m[2] || "GB").toUpperCase();
  if (unit === "TB") n *= 1024;
  else if (unit === "MB") n /= 1024;
  return n;
}

// Coerce to number; blanks/dashes -> null (so "no data" != 0).
function toNum(val) {
  if (val == null || val === "" || val === "-") return null;
  const n = typeof val === "number" ? val : parseFloat(String(val).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

const MONTHS = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY",
  "AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
const MONTH_TO_Q = { 0:"Q1",1:"Q1",2:"Q1",3:"Q2",4:"Q2",5:"Q2",
  6:"Q3",7:"Q3",8:"Q3",9:"Q4",10:"Q4",11:"Q4" };

function monthIndex(m) {
  if (m == null) return -1;
  return MONTHS.indexOf(String(m).trim().toUpperCase());
}
function quarterFromMonth(m) {
  const i = monthIndex(m);
  return i >= 0 ? MONTH_TO_Q[i] : null;
}
function normQuarter(q) {
  if (q == null) return null;
  const m = String(q).toUpperCase().match(/Q\s*([1-4])/);
  return m ? "Q" + m[1] : null;
}

/* ---------- canonical value maps (dirty-data normalization) -------------- */

function normDomainHealth(v) {
  if (v == null || v === "" || v === "-") return null;
  const s = String(v).trim().toLowerCase();
  if (s.includes("healh") || s === "healthy") return "Healthy";
  if (s.includes("possible")) return "Possible Service Issues";
  if (s.includes("incomplete")) return "Incomplete Setup";
  if (s.includes("no service")) return "No Services Selected";
  if (s.includes("no admin access")) return "Not Managed";
  if (s.includes("no access")) return "Not Managed";
  if (s.includes("not connected")) return "Not Connected";
  if (s.includes("not applicable") || s === "n/a" || s === "na") return "Not Applicable";
  if (s.includes("contract")) return "End Contract";
  return String(v).trim();
}

function normSecurityDefault(v) {
  if (v == null || v === "") return "UNKNOWN";
  const s = String(v).trim().toUpperCase();
  if (s.includes("CONDITIONAL")) return "CONDITIONAL ACCESS";
  if (s.includes("NO ACCESS")) return "NOT MANAGED";
  if (s.startsWith("DISABL")) return "DISABLED";
  if (s.startsWith("ENABL")) return "ENABLED";
  return s;
}

function yesNo(v) {
  if (v == null || v === "" || v === "-") return null;
  const s = String(v).trim().toUpperCase();
  if (s.startsWith("Y")) return "YES";
  if (s.startsWith("N")) return "NO";
  return s;
}

/* ---------- authentication methods (free-text → canonical tokens) --------
 * The SECURITY_DATA "Authentication Methods" cell is a free-text list, dirty in
 * the usual ways: mixed delimiters/spacing ("MFA,SMS,Email OTP" vs "MFA, SMS,
 * Email OTP"), the "FID02"→FIDO2 typo, and non-method sentinels ("NO ACCESS").
 * QBR.parseAuthMethods tokenizes it into the canonical vocabulary so the app can
 * filter/count by method (the structured MFA/SMS/Email OTP/SSPR columns only
 * cover four of the methods that actually appear here). Pure + node-testable. */
QBR.AUTH_CANON = ["Passkey/FIDO2", "Temporary Access Pass", "Authenticator/OATH", "MFA", "SMS", "Voice", "Email OTP"];
QBR.AUTH_STRONG = { "Passkey/FIDO2": 1, "Temporary Access Pass": 1, "Authenticator/OATH": 1, "MFA": 1 };
QBR.AUTH_PHISH_RESISTANT = { "Passkey/FIDO2": 1, "Temporary Access Pass": 1 };
QBR.parseAuthMethods = function (text) {
  if (text == null) return [];
  var s = String(text).toUpperCase().replace(/FID0?2/g, "FIDO2");   // FID02 typo → FIDO2
  var t = s.trim();
  if (!t || t === "-" || /^N\/?A$/.test(t) || /NO ACCESS|NOT MANAGED|NO ADMIN|NO DATA/.test(s)) return [];
  var out = {};
  if (/PASSKEY|FIDO2|WINDOWS HELLO/.test(s)) out["Passkey/FIDO2"] = 1;
  if (/TEMPORARY ACCESS|\bTAP\b/.test(s)) out["Temporary Access Pass"] = 1;
  if (/AUTHENTICATOR|OATH|\bTOTP\b/.test(s)) out["Authenticator/OATH"] = 1;
  if (/\bMFA\b/.test(s)) out["MFA"] = 1;
  if (/\bSMS\b/.test(s)) out["SMS"] = 1;
  if (/VOICE/.test(s)) out["Voice"] = 1;
  if (/EMAIL OTP|EMAIL ONE|\bOTP\b/.test(s)) out["Email OTP"] = 1;
  return QBR.AUTH_CANON.filter(function (m) { return out[m]; });   // canonical order, deduped
};

/* ---------- header-based column resolver --------------------------------- */

// Given a header row (array), return a function idx(patterns[]) -> column index.
function makeResolver(headerRow) {
  const heads = headerRow.map(h => (h == null ? "" : String(h).trim().toLowerCase()));
  return function idx(patterns) {
    for (const p of patterns) {
      const pl = p.toLowerCase();
      // exact first
      let i = heads.indexOf(pl);
      if (i >= 0) return i;
      // then substring
      i = heads.findIndex(h => h.includes(pl));
      if (i >= 0) return i;
    }
    return -1;
  };
}

// Find the sheet whose (normalized) name matches any candidate.
function findSheet(wb, candidates) {
  const names = wb.SheetNames;
  for (const c of candidates) {
    const cl = c.toLowerCase().replace(/\s+/g, "");
    const hit = names.find(n => n.toLowerCase().replace(/\s+/g, "") === cl);
    if (hit) return hit;
  }
  // fallback: substring
  for (const c of candidates) {
    const cl = c.toLowerCase().replace(/\s+/g, "");
    const hit = names.find(n => n.toLowerCase().replace(/\s+/g, "").includes(cl));
    if (hit) return hit;
  }
  return null;
}

function sheetRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
}

/* ---------- per-sheet parsers -------------------------------------------- */

function parseRisky(wb, model) {
  const name = findSheet(wb, ["RISKY_USERS_AND_DOMAIN", "RISKY USERS AND DOMAIN"]);
  if (!name) return;
  const rows = sheetRows(wb, name);
  const idx = makeResolver(rows[0]);
  const cMonth = idx(["month"]);
  const cSchool = idx(["school_rakso", "school"]);
  const cOrg = idx(["organization"]);
  const cRisky = idx(["total risky users", "risky"]);
  const cHealth = idx(["domain_health", "domain health"]);
  const cErr = idx(["error cause", "error_cause", "remarks"]);
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    const mi = monthIndex(row[cMonth]);
    model.risky.push({
      key: schoolKey(school),
      schoolRaw: dispName(school),
      org: cOrg >= 0 ? (row[cOrg] || null) : null,
      month: mi >= 0 ? MONTHS[mi] : null,
      monthIdx: mi,
      quarter: quarterFromMonth(row[cMonth]),
      risky: toNum(row[cRisky]),
      health: normDomainHealth(cHealth >= 0 ? row[cHealth] : null),
      errorCause: cErr >= 0 ? (row[cErr] != null && String(row[cErr]).trim() ? String(row[cErr]).trim() : null) : null,
    });
  }
  model.sources.risky = true;
}

function parseSecurity(wb, model) {
  const name = findSheet(wb, ["SECURITY_DATA", "SECURITY DATA"]);
  if (!name) return;
  const rows = sheetRows(wb, name);
  const idx = makeResolver(rows[0]);
  const cSchool = idx(["school"]);
  const cGdap = idx(["gdap"]);
  const cDomain = idx(["domain status", "domain"]);
  const cSecDef = idx(["security default"]);
  const cAuth = idx(["authentication methods", "auth"]);
  const cMfa = idx(["mfa"]);
  const cOtp = idx(["email otp"]);
  const cSms = idx(["sms"]);
  const cSspr = idx(["sspr"]);
  // GDAP relationship metadata — pre-wired for optional Expiry / Roles columns
  const cGdapExp = idx(["gdap expiry", "gdap expiration", "gdap expiry date", "gdap expires"]);
  const cGdapRoles = idx(["gdap roles", "gdap role", "gdap admin roles"]);
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    const gdapRaw = cGdap >= 0 ? row[cGdap] : null;
    const gdapStr = gdapRaw != null ? String(gdapRaw).trim() : "";
    // The GDAP cell holds relationship GUID(s); a few tenants have 2-3 stacked.
    // "None"/blank = no delegated admin relationship.
    const gdapIds = (gdapStr && !/^none$/i.test(gdapStr))
      ? gdapStr.split(/[\n;,]+/).map(x => x.trim()).filter(Boolean) : [];
    model.security.push({
      key: schoolKey(school),
      schoolRaw: dispName(school),
      gdap: gdapStr || null,
      hasGdap: gdapIds.length > 0,
      gdapIds: gdapIds,
      gdapExpiryRaw: cGdapExp >= 0 ? (row[cGdapExp] || null) : null,
      gdapRoles: cGdapRoles >= 0 ? (row[cGdapRoles] || null) : null,
      domainStatus: cDomain >= 0 ? (row[cDomain] || null) : null,
      securityDefault: normSecurityDefault(cSecDef >= 0 ? row[cSecDef] : null),
      authMethods: cAuth >= 0 ? (row[cAuth] || null) : null,
      methods: QBR.parseAuthMethods(cAuth >= 0 ? row[cAuth] : null),   // canonical tokens for filtering
      mfa: yesNo(cMfa >= 0 ? row[cMfa] : null),
      emailOtp: yesNo(cOtp >= 0 ? row[cOtp] : null),
      sms: yesNo(cSms >= 0 ? row[cSms] : null),
      sspr: yesNo(cSspr >= 0 ? row[cSspr] : null),
    });
  }
  model.sources.security = true;
}

function parseStorage(wb, model) {
  const name = findSheet(wb, ["STORAGE_DATA", "STORAGE DATA", "STORAGE REPORT"]);
  if (!name) return;
  const rows = sheetRows(wb, name);
  const idx = makeResolver(rows[0]);
  let cQ = idx(["quarter"]);
  if (cQ < 0) cQ = 0; // STORAGE holds the quarter label ("Q1") in the first column
  const cSchool = idx(["school", "client"]);
  const cOrg = idx(["organization"]);
  const cOne = idx(["onedrive"]);
  const cEx = idx(["exchange"]);
  const cSp = idx(["share point", "sharepoint"]);
  const cUsed = idx(["used storage(gb)", "used storage"]);        // numeric GB column (decimal)
  const cTotal = idx(["total storage(gb)", "total storage"]);
  const cPct = idx(["percentage%", "percentage"]);
  const cUsage = idx(["usage"]);                                  // "X TB of Y TB used" text
  let cUsedTxt = idx(["current storage"]);                        // the "used" headline text col
  if (cUsedTxt < 0 && cUsage > 0) cUsedTxt = cUsage - 1;
  // Unit bases: per-service cells are "X TB" text (toGB -> TB*1024, binary GB); the
  // "Used/Total Storage(GB)" columns are decimal GB (TB*1000). We normalise everything
  // to binary GB (display divides by 1024).
  //
  // The source is dirty in BOTH directions: some rows have a corrupt per-service cell
  // (e.g. Exchange "168.83 TB" that should be GB — inflates the service sum), and some
  // rows have a corrupt Used-GB column (e.g. 1250 that should be 12500 — deflates it).
  // No single field is authoritative, so we take the MEDIAN of the three independent
  // "used" estimates (used-text, used-GB-column, service-sum). The median ignores a
  // single bad field in either direction while agreeing with the others on clean rows.
  const decToBin = v => (v == null ? null : v / 1000 * 1024);
  const TBb = 1024;
  // The median only delivers its "ignore one corrupt field" guarantee at THREE
  // values. At two it averages them, so a single bad cell still drags the answer
  // halfway (e.g. 3.9 TB and a mislabeled 168 TB average to a meaningless 86 TB).
  // So: 3+ -> median; 2 -> agree within 25% ? mean : the higher-confidence source;
  // 1 -> that value. `arr` MUST be ordered most-trusted-first (used-text carries its
  // own explicit unit per cell, so it is the most reliable of the three).
  const bestUsed = arr => {
    const ranked = arr.map((x, i) => [x, i]).filter(p => p[0] != null && !isNaN(p[0]) && p[0] >= 0);
    if (!ranked.length) return null;
    if (ranked.length === 1) return ranked[0][0];
    if (ranked.length === 2) {
      const [a, b] = [ranked[0][0], ranked[1][0]];
      const hi = Math.max(a, b), lo = Math.min(a, b);
      if (hi === 0 || (hi - lo) / hi <= 0.25) return (a + b) / 2;  // agree -> average
      // They disagree and there is no third value to break the tie. Do NOT pick a
      // favourite: the "current storage" text is precisely the column that carries
      // unit mislabels ("1.35 GB" for 1.35 TB), and the service sum is the column
      // that carries magnitude typos — either can be the liar. Prefer the larger
      // only when the smaller looks like a TB-labelled-as-GB slip (~1024x), which is
      // the one mismatch we can actually diagnose; otherwise report no reading.
      if (lo > 0 && hi / lo > 900 && hi / lo < 1150) return hi;
      return null;
    }
    const a = ranked.map(p => p[0]).sort((x, y) => x - y);
    return a[a.length >> 1];
  };
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    let od = toGB(cOne >= 0 ? row[cOne] : null);
    let ex = toGB(cEx >= 0 ? row[cEx] : null);
    let sp = toGB(cSp >= 0 ? row[cSp] : null);
    const hasSvc = od != null || ex != null || sp != null;
    const svcRaw = (od || 0) + (ex || 0) + (sp || 0);
    // --- USED: median of the three estimates (all in binary GB) ---
    // The "current storage" text column carries its own explicit unit per row (this
    // sheet mixes "85.92 TB" and "440.62 GB"), so let toGB honour it — do NOT coerce
    // the unit, or the GB-denominated rows get inflated ~1000x.
    const usedTxtBin = cUsedTxt >= 0 ? toGB(row[cUsedTxt]) : null;
    const usedColBin = decToBin(toNum(cUsed >= 0 ? row[cUsed] : null));
    const svcBin = hasSvc ? svcRaw : null;
    // The USAGE cell ("1.35 TB of 100.15 TB used") carries a FOURTH independent
    // reading of "used", with its own explicit unit. It was previously mined only
    // for the capacity on the right of "of". Using the left side too is what lets
    // the median resolve rows where the storage-text unit is mislabeled — without
    // it those rows fall back to a two-source coin-flip. NB: these cells are salted
    // with LRM/RLM marks, so strip them before matching or every parse misses.
    const usageTxt = String(cUsage >= 0 ? (row[cUsage] || "") : "")
      .replace(/[​-‏‪-‮﻿]/g, "");
    const um = usageTxt.match(/([\d.]+)\s*(TB|GB|MB)\s*of/i);
    const usageUsedBin = um ? toGB(um[1] + " " + um[2]) : null;
    const usedGB = bestUsed([usedTxtBin, usedColBin, svcBin, usageUsedBin]);
    // --- TOTAL: prefer the numeric column, guard absurd values, fall back to the
    //     "of Y TB used" text total ---
    let totalGB = decToBin(toNum(cTotal >= 0 ? row[cTotal] : null));
    const tt = (usageTxt.match(/of\s*([\d.]+)\s*TB/i) || [])[1];   // marks already stripped
    const totTxtBin = tt ? parseFloat(tt) * TBb : null;
    if (totalGB == null || totalGB <= 0 || totalGB > 10240000) totalGB = totTxtBin; // >10,000 TB is impossible here
    // --- Sanitise per-service for the composition breakdown: drop any single service
    //     that exceeds capacity, and if the remaining sum still can't reconcile with
    //     the median used, mark the split unreliable rather than draw a false bar. ---
    const cap = totalGB;
    if (cap != null) {
      if (od != null && od > cap * 1.02) od = null;
      if (ex != null && ex > cap * 1.02) ex = null;
      if (sp != null && sp > cap * 1.02) sp = null;
    }
    let svcSuspect = false;
    if (usedGB != null && ((od || 0) + (ex || 0) + (sp || 0)) > usedGB * 1.5 + TBb) {
      od = ex = sp = null; svcSuspect = true;
    }
    // --- PCT: recompute from the corrected used/total when we have a sane capacity,
    //     otherwise fall back to the reported percentage column. ---
    let pct = (() => { const p = toNum(cPct >= 0 ? row[cPct] : null); return p == null ? null : (p <= 1 ? p * 100 : p); })();
    if (usedGB != null && totalGB) pct = usedGB / totalGB * 100;
    if (usedGB == null && !hasSvc) continue;
    model.storage.push({
      key: schoolKey(school),
      schoolRaw: dispName(school),
      quarter: normQuarter(cQ >= 0 ? row[cQ] : null),
      org: cOrg >= 0 ? (row[cOrg] || null) : null,
      onedriveGB: od, exchangeGB: ex, sharepointGB: sp,
      usedGB: usedGB != null ? usedGB : ((od || 0) + (ex || 0) + (sp || 0)),
      totalGB: totalGB,
      pct: pct,
      svcSuspect: svcSuspect,
    });
  }
  model.sources.storage = true;
}

function parseUsage(wb, model) {
  const name = findSheet(wb, ["USAGE_REPORT", "USAGE REPORT"]);
  if (!name) return;
  const rows = sheetRows(wb, name);
  const idx = makeResolver(rows[0]);
  const cQ = idx(["quarter"]);
  const cSchool = idx(["client", "school"]);
  const cOrg = idx(["organization"]);
  const cExA = idx(["exchange active"]);
  const cOneA = idx(["onedrive active"]);
  const cSpA = idx(["sharepoint active"]);
  const cTeamsA = idx(["teams active"]);
  const cO365A = idx(["office 365 active"]);
  const cO365T = idx(["total of office 365", "total office 365"]);
  const cPct = idx(["% of usage report", "usage report", "usage"]);
  // per-service Inactive / Total + license columns (for the per-tenant detail view)
  const cExI = idx(["exchange inactive"]), cExT = idx(["total of exchange", "total exchange"]);
  const cOneI = idx(["onedrive inactive"]), cOneT = idx(["total of onedrive", "total onedrive"]);
  const cSpI = idx(["sharepoint inactive"]), cSpT = idx(["total of sharepoint", "total sharepoint"]);
  const cTeamsI = idx(["teams inactive"]), cTeamsT = idx(["total of teams", "total teams"]);
  const cO365I = idx(["office 365 inactive"]);
  const cLic = idx(["assigned licenses", "number of assigned"]);
  const cAct = idx(["activated"]);
  const cDesk = idx(["desktop apps", "desktop"]);
  const cMob = idx(["mobile apps", "mobile"]);
  const cDate = idx(["date extracted", "report period"]);
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    let pct = toNum(cPct >= 0 ? row[cPct] : null);
    if (pct != null && pct <= 1) pct *= 100;
    const g = (c) => toNum(c >= 0 ? row[c] : null);
    model.usage.push({
      key: schoolKey(school),
      schoolRaw: dispName(school),
      quarter: normQuarter(cQ >= 0 ? row[cQ] : null),
      org: cOrg >= 0 ? (row[cOrg] || null) : null,
      exchangeActive: g(cExA), exchangeInactive: g(cExI), exchangeTotal: g(cExT),
      onedriveActive: g(cOneA), onedriveInactive: g(cOneI), onedriveTotal: g(cOneT),
      sharepointActive: g(cSpA), sharepointInactive: g(cSpI), sharepointTotal: g(cSpT),
      teamsActive: g(cTeamsA), teamsInactive: g(cTeamsI), teamsTotal: g(cTeamsT),
      office365Active: g(cO365A), office365Inactive: g(cO365I), office365Total: g(cO365T),
      assignedLicenses: g(cLic), activated: g(cAct), desktopApps: g(cDesk), mobileApps: g(cMob),
      dateExtracted: cDate >= 0 ? (row[cDate] || null) : null,
      usagePct: pct,
    });
  }
  model.sources.usage = true;
}

function parseCanva(wb, model) {
  const name = findSheet(wb, ["CANVA_STATUS", "CANVA STATUS"]);
  if (!name) return;
  const rows = sheetRows(wb, name);
  // header row 0 has quarter label in col 0 ("Q1") + SCHOOL etc.
  const idx = makeResolver(rows[0]);
  let cQ = idx(["quarter"]);
  if (cQ < 0) cQ = 0; // first col holds Q label
  const cSchool = idx(["school"]);
  const cCert = idx(["certificate status"]);
  const cExp = idx(["certificate update", "certificate"]);
  const cUsers = idx(["total user"]);
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    model.canva.push({
      key: schoolKey(school),
      schoolRaw: dispName(school),
      quarter: normQuarter(row[cQ]),
      certStatus: cCert >= 0 ? (row[cCert] || null) : null,
      certExpiry: cExp >= 0 ? (row[cExp] || null) : null,
      users: toNum(cUsers >= 0 ? row[cUsers] : null),
    });
  }
  model.sources.canva = true;
}

// Domain reputation vocabulary. TWO source dialects are handled:
//  (A) Google Postmaster's rated tiers (older / standalone files): HIGH/MEDIUM/LOW/BAD.
//  (B) Google Postmaster's live status strings (current tracker): "Issues detected",
//      "Verify to see health", "Not enough data". These are NOT the rated tiers and
//      must be preserved as their own states — collapsing them (the old behavior)
//      silently hid every "Issues detected" domain, the key deliverability signal.
function normReputation(v) {
  const s = String(v == null ? "" : v).trim();
  if (s === "" || s === "-") return "No entry";             // blank / dash — no record
  const u = s.toUpperCase();
  if (u.includes("ISSUE")) return "Issues detected";        // actionable: deliverability problem
  if (u.includes("VERIFY TO SEE") || u.includes("VERIFY TO")) return "Verify to see health"; // needs verification
  if (u.includes("NOT ENOUGH")) return "Not enough data";   // verified but too little mail to rate
  if (u.includes("NO DATA")) return "No data to display";
  if (["HIGH", "MEDIUM", "LOW", "BAD"].includes(u)) return u;
  if (u.startsWith("MED")) return "MEDIUM";
  // Unrecognised value: preserve it verbatim (never mask real data), but record it.
  // Downstream colour/badge/severity maps are keyed on the known states, so an
  // unknown string renders as plain grey text and would otherwise vanish silently.
  QBR._repUnknown = QBR._repUnknown || {};
  QBR._repUnknown[s] = (QBR._repUnknown[s] || 0) + 1;
  return s;
}

// Postmaster verification status as carried by the combined tracker's "Status"
// column (Verified / Unverified). Distinct from normPmTool, which reads the
// standalone verification sheet's "Google Postmaster Tool" column.
function normVerifyStatus(v) {
  const s = String(v == null ? "" : v).trim().toUpperCase();
  if (s === "" || s === "-") return null;
  if (s.startsWith("VERIF") && !s.startsWith("UNVERIF")) return "Verified";
  if (s.startsWith("UNVERIF") || s === "NOT VERIFIED") return "Unverified";
  if (s.includes("NOT MANAGED")) return "Not Managed";
  if (s.includes("MANAGED")) return "Managed";
  return null;
}

// Postmaster "Google Postmaster Tool" verification column.
function normPmTool(v) {
  const s = String(v == null ? "" : v).trim().toUpperCase();
  if (s === "VERIFIED") return "Verified";
  if (s === "MANAGED") return "Managed";
  if (s.includes("NOT MANAGED")) return "Not Managed";
  return "Unknown";
}
// Cloudflare / DNS TXT-record column.
function normDns(v) {
  const s = String(v == null ? "" : v).trim().toUpperCase();
  if (s.includes("ADDED")) return "Records added";
  if (s.includes("FOR VERIFICATION")) return "Pending";
  if (s.includes("NOT MANAGED")) return "Not Managed";
  return "Unknown";
}
// Normalize an Organization cell: blank / "-" -> null (so it reads "Unspecified").
function cleanOrg(v) { const s = String(v == null ? "" : v).trim(); return (s === "" || s === "-") ? null : s; }
// Excel serial date -> ISO yyyy-mm-dd (ignores junk / out-of-range values).
function excelDate(v) {
  const n = Number(v);
  if (!isFinite(n) || n < 20000 || n > 80000) return null;
  const d = new Date(Math.round((n - 25569) * 86400000));
  return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}

/* Postmaster is sourced two ways and this handles both:
 *   (A) Standalone GOOGLE POSTMASTERTOOLS.xlsx — a master "…Data" sheet carrying
 *       verification (Postmaster Tool + Cloudflare DNS), plus one tab per month
 *       (JANUARY..DECEMBER) carrying reputation / reference / spam.
 *   (B) Combined ALL TENANT tracker — a single GOOGLE_POSTMASTERTOOLS sheet with a
 *       MONTH column (reputation + spam, no verification).
 * When both are loaded the standalone wins (deduped in loadWorkbooks). Optional
 * Status / Health columns are auto-read if a future file adds them. */
function parsePostmaster(wb, model) {
  wb.SheetNames.forEach(sn => {
    const rows = sheetRows(wb, sn);
    if (!rows || !rows.length) return;
    const idx = makeResolver(rows[0]);
    const cSchool = idx(["school"]);
    const cOrg = idx(["organization"]);
    const cDomain = idx(["domain"]);

    // ---- (1) master verification sheet: has a "Google Postmaster Tool" or "…verification link" column ----
    const cTool = idx(["google postmaster tool", "postmaster tool"]);
    const cVerLink = idx(["domain verification link", "verification link"]);
    if ((cTool >= 0 || cVerLink >= 0) && cSchool >= 0) {
      const cDns = idx(["cloudflare", "dns"]);
      const cDate = idx(["date of verification", "date verified", "verification date"]);
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r], school = row[cSchool];
        if (!school || !schoolKey(school)) continue;
        if (cTool >= 0 && /postmaster tool/i.test(String(row[cTool]))) continue; // repeated header
        model.pmVerify.push({
          key: schoolKey(school), schoolRaw: dispName(school),
          org: cOrg >= 0 ? cleanOrg(row[cOrg]) : null,
          domain: cDomain >= 0 ? (row[cDomain] || null) : null,
          pmTool: normPmTool(cTool >= 0 ? row[cTool] : null),
          dns: normDns(cDns >= 0 ? row[cDns] : null),
          dateVerified: cDate >= 0 ? excelDate(row[cDate]) : null,
        });
      }
      model.sources.postmaster = true;
      return; // a verification sheet is never also a reputation sheet
    }

    // ---- (2) reputation sheet: monthly tab (name = month) OR combined sheet (has MONTH column) ----
    const cRep = idx(["domain reputation", "reputation"]);
    if (cRep < 0 || cSchool < 0) return;
    const tabMonth = monthIndex(sn);          // >=0 when the SHEET NAME is a month
    const cMonth = idx(["month"]);            // present on the combined tracker sheet
    if (tabMonth < 0 && cMonth < 0) return;   // neither layout — skip
    const src = tabMonth >= 0 ? "standalone" : "tracker";
    const cSpam = idx(["spam rate", "spam"]);
    const cStatus = idx(["status"]);          // optional (future files)
    const cHealth = idx(["health"]);          // optional (future files)
    const cRef = idx(["reference", "evidence"]);
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r], school = row[cSchool];
      if (!school || !schoolKey(school)) continue;
      if (/reputation/i.test(String(row[cRep]))) continue; // repeated header block
      const mi = tabMonth >= 0 ? tabMonth : monthIndex(row[cMonth]);
      if (mi < 0) continue;
      const spamRaw = cSpam >= 0 ? row[cSpam] : null;
      model.postmaster.push({
        key: schoolKey(school), schoolRaw: dispName(school),
        org: cOrg >= 0 ? cleanOrg(row[cOrg]) : null,
        month: MONTHS[mi], monthIdx: mi, quarter: MONTH_TO_Q[mi],
        domain: cDomain >= 0 ? (row[cDomain] || null) : null,
        reputation: normReputation(cRep >= 0 ? row[cRep] : null),
        status: cStatus >= 0 ? (row[cStatus] || null) : null,
        verifyStatus: normVerifyStatus(cStatus >= 0 ? row[cStatus] : null),
        health: cHealth >= 0 ? (row[cHealth] || null) : null,
        reference: cRef >= 0 ? (row[cRef] || null) : null,
        spamRate: toNum(spamRaw),
        spamRaw: (spamRaw == null || String(spamRaw).trim() === "-" || /no data/i.test(String(spamRaw))) ? null : String(spamRaw).trim(),
        _src: src,
      });
      model.sources.postmaster = true;
    }
  });
}

// Standalone monthly data wins over the combined-tracker sheet; then dedup by school+month.
function dedupePostmaster(model) {
  if (!model.postmaster.length) return;
  if (model.postmaster.some(r => r._src === "standalone"))
    model.postmaster = model.postmaster.filter(r => r._src !== "tracker");
  const seen = new Map();
  model.postmaster.forEach(r => { const k = r.key + "|" + r.monthIdx; if (!seen.has(k)) seen.set(k, r); });
  model.postmaster = [...seen.values()];
}

function parseUserManagement(wb, model) {
  // Detect SY sheets like "SY 2025-2026" / "SY 2026-2027"
  const syNames = wb.SheetNames.filter(n => /^\s*SY\s*20\d{2}\s*[-–]\s*20\d{2}/i.test(n));
  if (!syNames.length) return;
  syNames.forEach(sn => {
    const sy = sn.trim().replace(/\s+/g, " ").replace("–", "-");
    const rows = sheetRows(wb, sn);
    const idx = makeResolver(rows[0]);
    const cSchool = idx(["school"]);
    const cGrade = idx(["update grade level", "grade level"]);
    const cExtract = idx(["extract users", "extract"]);
    const cRemarks = idx(["remarks"]);
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const school = row[cSchool];
      if (!school || !schoolKey(school)) continue;
      // trim to empty -> null (cells often hold whitespace-only placeholders)
      const clean = (i) => { if (i < 0) return null; const v = row[i]; if (v == null) return null; const s = String(v).trim(); return s === "" ? null : s; };
      const remarks = clean(cRemarks);
      const gradeFile = clean(cGrade);
      const extractFile = clean(cExtract);
      // Strict 2-state: Updated iff REMARKS contains UPDATED, else Pending.
      const status = remarks && /updated/i.test(remarks) ? "Updated" : "Pending";
      model.usermgmt.push({
        key: schoolKey(school), schoolRaw: dispName(school), sy,
        gradeFile, extractFile, remarks,
        gradeDone: !!gradeFile, extractDone: !!extractFile, status,
      });
    }
  });
  if (model.usermgmt.length) model.sources.usermgmt = true;
}

// Parse "YYYY-MM-DD-HH-MM-SS" -> epoch ms (local), or null.
function parseRegTimestamp(v) {
  if (v == null) return null;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[-\sT](\d{1,2})-(\d{1,2})-(\d{1,2}))?/);
  if (!m) return null;
  const [, y, mo, d, hh, mm, ss] = m;
  const dt = new Date(+y, +mo - 1, +d, +(hh || 0), +(mm || 0), +(ss || 0));
  return isNaN(dt.getTime()) ? null : dt.getTime();
}

// Add exactly one calendar year to an epoch-ms timestamp.
function addOneYear(ms) {
  if (ms == null) return null;
  const d = new Date(ms);
  d.setFullYear(d.getFullYear() + 1);
  return d.getTime();
}

/* ---------- domain-registration status (pure, source-first) --------------
 * The tracker already states each domain's Status + Remaining Days and carries
 * ops notes in Remarks; the Registration column is dirty (holds sentinels like
 * "Error"/"PENDING" instead of a date). QBR.domainStatus READS those authored
 * signals as truth and only computes from dates when they're blank, mapping the
 * reality onto one operational vocabulary with an action + KPI group. Pure and
 * node-testable (pass `now` for deterministic tests).
 *
 *   rec = { regMs, expMs, regRaw, srcDays, statusRaw, remarks }
 *   → { status, tier, action, days, group }
 * group ∈ ok | attention | critical | inactive (drives KPI buckets/colours). */
QBR.DR_STATUS_ORDER = ["For Deletion", "Expired", "For Renewal", "Expiring Soon", "Error",
  "Invalid Domain", "Pending", "Active", "End Contract", "Deleted", "Not Registered"];
QBR.domainStatus = function (rec, now) {
  rec = rec || {};
  now = now == null ? Date.now() : now;
  var rl = String(rec.remarks == null ? "" : rec.remarks).toLowerCase();
  var statusRaw = String(rec.statusRaw == null ? "" : rec.statusRaw).trim().toLowerCase();
  var regRaw = String(rec.regRaw == null ? "" : rec.regRaw).trim();
  // a non-date Registration value is a lookup sentinel (Error / PENDING / Expired)
  var regSent = (rec.regMs == null && regRaw && !/^\d{4}-/.test(regRaw)) ? regRaw.toLowerCase() : "";
  var days = rec.srcDays;
  if (days == null && rec.expMs != null) days = Math.floor((rec.expMs - now) / 86400000);
  function ret(status, tier, action, group) {
    return { status: status, tier: tier || null, action: action || null, days: (days == null ? null : days), group: group };
  }
  // 1) Remarks — explicit operational notes win (most specific / terminal first)
  if (/already been deleted|has been deleted|been deleted from|deleted from the edu/.test(rl)) return ret("Deleted", null, null, "inactive");
  if (/for deletion/.test(rl)) return ret("For Deletion", null, "Pay now — deletion pending", "critical");
  if (/end contract|end of contract/.test(rl)) return ret("End Contract", null, null, "inactive");
  if (/not a valid domain|invalid domain/.test(rl)) return ret("Invalid Domain", null, "Fix domain name", "attention");
  if (/for renewal/.test(rl)) return ret("For Renewal", null, "Pay renewal fee", "critical");
  // 2) Sheet Status column
  if (statusRaw === "expired") return ret("Expired", null, "Renew now (expired)", "critical");
  if (/renewal/.test(statusRaw)) return ret("For Renewal", null, "Pay renewal fee", "critical");
  // 3) Registration sentinel (lookup state, not a date)
  if (/error/.test(regSent)) return ret("Error", null, "Re-check registration lookup", "attention");
  if (/pending/.test(regSent)) return ret("Pending", null, "Registration in progress", "attention");
  if (/expired/.test(regSent)) return ret("Expired", null, "Renew now (expired)", "critical");
  // 4) Day-based (source-stated days preferred; two-tier expiry window)
  if (days != null) {
    if (days < 0) return ret("Expired", null, "Renew now (expired)", "critical");
    if (days <= 30) return ret("Expiring Soon", "Urgent", "Renew within " + days + " days", "critical");
    if (days <= 60) return ret("Expiring Soon", "Watch", "Renew within " + days + " days", "attention");
    return ret("Active", null, null, "ok");
  }
  // 5) Marked current but no expiry info
  if (statusRaw === "current") return ret("Active", null, null, "ok");
  // 6) nothing usable
  return ret("Not Registered", null, null, "inactive");
};

// Detect a domain-registration sheet by header SHAPE, so both layouts load:
//  (A) older export: a "Registration status" header (sheet often named CURRENT);
//  (B) per-quarter export: sheet named e.g. "Q3", header "Registration", plus
//      DOMAIN + Expiration / Remaining Days / Status columns.
// Requiring DOMAIN + (Expiration | Remaining Days) avoids matching RISKY
// (Domain_Health) or SECURITY (Domain Status), which have neither.
function findDomainRegSheet(wb) {
  for (const sn of wb.SheetNames) {
    const first = (sheetRows(wb, sn)[0] || []).map(h => String(h == null ? "" : h).toLowerCase());
    const has = p => first.some(h => h.indexOf(p) >= 0);
    if (has("registration status") || (has("domain") && (has("expiration") || has("remaining day")))) return sn;
  }
  return null;
}
// True if a workbook carries a domain-registration sheet (used by the app to
// decide whether an upload should override the built-in default seed).
QBR.hasDomainRegSheet = function (wb) { return !!findDomainRegSheet(wb); };

function parseDomainRegistration(wb, model) {
  const name = findDomainRegSheet(wb);
  if (!name) return;
  const rows = sheetRows(wb, name);
  const idx = makeResolver(rows[0]);
  const cSchool = idx(["schools", "school"]);
  const cDomain = idx(["domain"]);
  const cReg = idx(["registration status", "registration"]);
  const cExp = idx(["expiration"]);
  const cDays = idx(["remaining days", "remaining"]);
  let cStatus = idx(["status"]);
  if (cStatus === cReg) cStatus = -1;                 // "Registration status" is the reg col, not a status col
  const cValidYr = idx(["validy in year", "validity in year", "validity period", "validity"]);
  const cValidPD = idx(["validpd", "valid pd"]);
  const cRemarks = idx(["remarks"]);
  const cCF = idx(["conditional formatting"]);         // signed remaining-days number (sort key)
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const school = row[cSchool];
    if (!school || !schoolKey(school)) continue;
    const regRaw = cReg >= 0 ? row[cReg] : null;
    const regMs = parseRegTimestamp(regRaw);
    // Signed remaining days: prefer the "N days remaining / since expiration" text,
    // else the Conditional-formatting number (already signed).
    let srcDays = null;
    if (cDays >= 0 && row[cDays] != null) {
      const dm = String(row[cDays]).match(/(-?\d+)\s*days?\s*(remaining|since|overdue|expired|past|ago)?/i);
      if (dm) { srcDays = parseInt(dm[1], 10); if (/since|overdue|expired|past|ago/i.test(dm[2] || "")) srcDays = -Math.abs(srcDays); }
    }
    if (srcDays == null && cCF >= 0) { const cf = toNum(row[cCF]); if (cf != null) srcDays = cf; }
    // Expiration: explicit column, else registration + validity years (default 1).
    let expMs = parseRegTimestamp(cExp >= 0 ? row[cExp] : null);
    let validYr = cValidYr >= 0 ? toNum(row[cValidYr]) : null;
    if (validYr == null && cValidPD >= 0) { const vm = String(row[cValidPD] || "").match(/(\d+)/); if (vm) validYr = parseInt(vm[1], 10); }
    if (expMs == null && regMs != null) { let e = regMs; const yrs = validYr || 1; for (let i = 0; i < yrs; i++) e = addOneYear(e); expMs = e; }
    const rec = {
      key: schoolKey(school),
      schoolRaw: dispName(school),
      domain: cDomain >= 0 ? (row[cDomain] || null) : null,
      regMs, expMs,
      regRaw: regRaw != null ? String(regRaw).trim() : null,
      srcDays,
      statusRaw: cStatus >= 0 ? (row[cStatus] || null) : null,
      remarks: cRemarks >= 0 && row[cRemarks] != null && String(row[cRemarks]).trim() ? String(row[cRemarks]).trim() : null,
      validYr: validYr || null,
    };
    const st = QBR.domainStatus(rec);
    rec.status = st.status; rec.tier = st.tier; rec.action = st.action; rec.days = st.days; rec.group = st.group;
    model.domainreg.push(rec);
  }
  model.sources.domainreg = true;
}

/* ---------- master school dimension -------------------------------------- */

function buildMasterSchools(model) {
  const master = new Map(); // key -> { key, name, aliases:Set, inSheets:Set }
  const add = (rec, sheet) => {
    if (!rec.key) return;
    let m = master.get(rec.key);
    if (!m) { m = { key: rec.key, name: rec.schoolRaw, aliases: new Set(), inSheets: new Set() }; master.set(rec.key, m); }
    m.aliases.add(rec.schoolRaw);
    m.inSheets.add(sheet);
    // prefer the longest raw name as canonical display
    if (rec.schoolRaw.length > m.name.length) m.name = rec.schoolRaw;
  };
  model.risky.forEach(r => add(r, "risky"));
  model.security.forEach(r => add(r, "security"));
  model.storage.forEach(r => add(r, "storage"));
  model.usage.forEach(r => add(r, "usage"));
  model.canva.forEach(r => add(r, "canva"));
  model.postmaster.forEach(r => add(r, "postmaster"));
  model.pmVerify.forEach(r => add(r, "postmaster"));
  model.usermgmt.forEach(r => add(r, "usermgmt"));
  model.domainreg.forEach(r => add(r, "domainreg"));
  model.master = master;

  // Verification lookup by school key — prefer a "Verified" record when a school has several.
  model.pmVerifyByKey = {};
  model.pmVerify.forEach(v => {
    const cur = model.pmVerifyByKey[v.key];
    if (!cur || (v.pmTool === "Verified" && cur.pmTool !== "Verified")) model.pmVerifyByKey[v.key] = v;
  });

  // org lookup by school key (first non-null org across all sheets that carry it)
  model.orgByKey = {};
  [].concat(model.risky, model.storage, model.usage, model.postmaster, model.pmVerify).forEach(r => {
    if (r.org && !model.orgByKey[r.key]) model.orgByKey[r.key] = r.org;
  });

  // data-quality: schools missing from each sheet
  model.quality = {
    totalSchools: master.size,
    coverage: {},
    unmatched: {},
  };
  ["risky","security","storage","usage","canva","postmaster","usermgmt","domainreg"].forEach(s => {
    let inS = 0; const missing = [];
    master.forEach(m => { if (m.inSheets.has(s)) inS++; else missing.push(m.name); });
    model.quality.coverage[s] = inS;
    model.quality.unmatched[s] = missing;
  });
}

/* ---------- public API --------------------------------------------------- */

function emptyModel() {
  return {
    risky: [], security: [], storage: [], usage: [], canva: [], postmaster: [], pmVerify: [], usermgmt: [], domainreg: [],
    sources: { risky:false, security:false, storage:false, usage:false, canva:false, postmaster:false, usermgmt:false, domainreg:false },
    master: new Map(), orgByKey: {}, pmVerifyByKey: {}, quality: null,
    raw: {}, // raw header+rows per key sheet, for the Data Quality self-audit (QBR.audit)
  };
}

// Stash the raw (unparsed) rows of the key sheets so the Data Quality tab can audit
// source cells (units, typos, sentinels) without re-reading the workbook. In-memory
// only; later workbooks overwrite a sheet's raw rows (last upload wins), which is what
// we want (e.g. a standalone Postmaster export supersedes the embedded sheet).
function captureRaw(wb, model) {
  const grab = (cands, key) => { const n = findSheet(wb, cands); if (n) model.raw[key] = sheetRows(wb, n); };
  grab(["RISKY_USERS_AND_DOMAIN", "RISKY USERS AND DOMAIN"], "risky");
  grab(["SECURITY_DATA", "SECURITY DATA"], "security");
  grab(["STORAGE_DATA", "STORAGE DATA", "STORAGE REPORT"], "storage");
  grab(["USAGE_REPORT", "USAGE REPORT"], "usage");
  grab(["CANVA_STATUS", "CANVA STATUS"], "canva");
  grab(["GOOGLE_POSTMASTERTOOLS", "GOOGLE POSTMASTER"], "postmaster");
  grab(["PRODUCT_AND_SERVICES", "PRODUCT AND SERVICES"], "product");
}

// Parse one or more ArrayBuffers (each an uploaded workbook) into a merged model.
QBR.loadWorkbooks = function (buffers) {
  const model = emptyModel();
  QBR._repUnknown = {};   // reset the unrecognised-reputation register for this load
  for (const buf of buffers) {
    const wb = XLSX.read(buf, { type: "array" });
    parseRisky(wb, model);
    parseSecurity(wb, model);
    parseStorage(wb, model);
    parseUsage(wb, model);
    parseCanva(wb, model);
    parsePostmaster(wb, model);
    parseUserManagement(wb, model);
    parseDomainRegistration(wb, model);
    captureRaw(wb, model);
  }
  canonicalizeOrgs(model);
  dedupePostmaster(model);
  synthesizePmVerify(model);
  buildMasterSchools(model);
  return model;
};

// The standalone GOOGLE POSTMASTERTOOLS.xlsx carries a dedicated verification sheet
// (-> model.pmVerify). The combined tracker does not, but its GOOGLE_POSTMASTERTOOLS
// sheet has a per-month "Status" (Verified / Unverified) column. When no standalone
// verification exists, synthesize one record per domain from that Status (latest
// month wins) so the Postmaster tab's verification donut / KPIs / action list are
// populated instead of blank. A real standalone sheet always takes precedence.
function synthesizePmVerify(model) {
  if (model.pmVerify && model.pmVerify.length) return;      // standalone data present — leave it
  const byKey = {};
  model.postmaster.forEach(p => {
    if (!p.verifyStatus) return;
    const cur = byKey[p.key];
    if (!cur || p.monthIdx >= cur.monthIdx) byKey[p.key] = p;
  });
  Object.values(byKey).forEach(p => {
    model.pmVerify.push({
      key: p.key, schoolRaw: p.schoolRaw, org: p.org,
      domain: p.domain,
      pmTool: p.verifyStatus,          // "Verified" | "Unverified" | "Managed" | "Not Managed"
      dns: "Unknown",                  // combined tracker has no Cloudflare/DNS column
      dateVerified: null,
      monthIdx: p.monthIdx,
      _synth: true,
    });
  });
  if (model.pmVerify.length) model.sources.postmaster = true;
}

// Organizations are federation codes (BULPRISA, DICES, RCBNES…). Canonicalize to
// UPPERCASE so casing variants merge into one org (e.g. "Bulprisa" -> "BULPRISA",
// "Government" -> "GOVERNMENT"); blank / "-" become null (Unspecified).
function canonicalizeOrgs(model) {
  const canon = o => { const s = String(o == null ? "" : o).trim(); return (s === "" || s === "-") ? null : s.toUpperCase(); };
  [model.risky, model.security, model.storage, model.usage, model.canva, model.postmaster, model.pmVerify, model.usermgmt, model.domainreg]
    .forEach(arr => arr.forEach(r => { if (r && "org" in r) r.org = canon(r.org); }));
}

/* ---------------------------------------------------------------------------
 * QBR.buildCleanRows(model) — the "clean dataset": one corrected, normalized row
 * per school × quarter, joining every source by the master key. This is the tidy
 * analytical feed for the executive snapshot and the app's "Export Clean Data".
 * Corrections already applied by the parsers (median storage, capacity guards,
 * org canonicalization, name normalization) flow straight through. Missing values
 * stay null (rendered as "" / "—"), never 0.
 * ------------------------------------------------------------------------- */
QBR.CLEAN_COLS = [
  "School", "Organization", "Quarter", "Managed", "Domain Status", "Risky Detections",
  "Security Default", "MFA", "SSPR", "GDAP", "O365 Usage %", "O365 Active Users",
  "Storage Used (TB)", "Storage Total (TB)", "Storage Util %", "Canva Users", "Email Reputation"
];
QBR.buildCleanRows = function (model) {
  const QS = ["Q1", "Q2", "Q3", "Q4"];
  const round = (v, d) => (v == null ? null : Math.round(v * Math.pow(10, d)) / Math.pow(10, d));
  // snapshot security by key
  const secByKey = {};
  model.security.forEach(s => { if (!secByKey[s.key]) secByKey[s.key] = s; });
  // helper: index an array by key+quarter (last non-null wins for scalar fields)
  const byKQ = (arr) => { const m = {}; arr.forEach(r => { if (r.quarter) m[r.key + "|" + r.quarter] = r; }); return m; };
  const storIdx = byKQ(model.storage), usageIdx = byKQ(model.usage), canvaIdx = byKQ(model.canva);
  // risky: sum detections + latest-month health, per key+quarter
  const riskyIdx = {};
  model.risky.forEach(r => {
    if (!r.quarter) return; const k = r.key + "|" + r.quarter;
    const e = riskyIdx[k] || (riskyIdx[k] = { risky: 0, health: null, mi: -1, hasRisky: false });
    if (r.risky != null) { e.risky += r.risky; e.hasRisky = true; }
    if (r.health && r.monthIdx >= e.mi) { e.health = r.health; e.mi = r.monthIdx; }
  });
  // postmaster: latest-month reputation per key+quarter
  const pmIdx = {};
  model.postmaster.forEach(p => {
    if (!p.quarter || !p.reputation) return; const k = p.key + "|" + p.quarter;
    const e = pmIdx[k]; if (!e || p.monthIdx >= e.mi) pmIdx[k] = { rep: p.reputation, mi: p.monthIdx };
  });
  const rows = [];
  model.master.forEach((ms, key) => {
    const sec = secByKey[key] || {};
    const managed = sec.gdap ? "Yes" : "No";
    QS.forEach(q => {
      const kq = key + "|" + q;
      const rk = riskyIdx[kq], st = storIdx[kq], us = usageIdx[kq], cv = canvaIdx[kq], pm = pmIdx[kq];
      const populated = (rk && (rk.hasRisky || rk.health)) ||
        (st && (st.usedGB != null || st.totalGB != null)) ||
        (us && us.usagePct != null) || (cv && cv.users != null) || pm;
      if (!populated) return;
      rows.push({
        School: ms.name,
        Organization: model.orgByKey[key] || "",
        Quarter: q,
        Managed: managed,
        "Domain Status": (rk && rk.health) || "",
        "Risky Detections": rk && rk.hasRisky ? rk.risky : null,
        "Security Default": sec.securityDefault || "",
        MFA: sec.mfa || "", SSPR: sec.sspr || "", GDAP: sec.gdap ? "Granted" : "",
        "O365 Usage %": us ? round(us.usagePct, 1) : null,
        "O365 Active Users": us ? us.office365Active : null,
        "Storage Used (TB)": st && st.usedGB != null ? round(st.usedGB / 1024, 2) : null,
        "Storage Total (TB)": st && st.totalGB != null ? round(st.totalGB / 1024, 2) : null,
        "Storage Util %": st && st.pct != null ? round(st.pct, 1) : null,
        "Canva Users": cv ? cv.users : null,
        "Email Reputation": (pm && pm.rep) || ""
      });
    });
  });
  rows.sort((a, b) => a.School.localeCompare(b.School) || a.Quarter.localeCompare(b.Quarter));
  return rows;
};

/* =============================================================================
 * Tenant Health Index — pure, DOM-free, node-testable scorer.
 * -----------------------------------------------------------------------------
 * QBR.scoreTenantHealth(dims) → { thi, scores, present } | null
 *
 * Composite 0–100 health score for ONE tenant from up to six dimensions. Each
 * sub-score is 0–100, or null when that tenant has no signal for the dimension.
 * Missing dimensions are EXCLUDED and the remaining weights renormalized — a
 * data gap must never read as a health failure. A tenant needs ≥2 present
 * dimensions to be scored (never fabricate health from a single signal);
 * otherwise the function returns null.
 *
 *   dims = {
 *     security  : { securityDefault, mfa, sspr } | null,  // security snapshot row
 *     risky     : { sum, has } | null,                    // summed risky users + presence flag
 *     adoption  : <pct number> | null,                    // O365 usage %
 *     domain    : <status string> | null,                 // domain-health status
 *     postmaster: <reputation string> | null,             // email reputation
 *     storage   : <pct number> | null,                    // storage utilization %
 *   }
 *
 * Sub-scorers are exposed as QBR.thiScorers so callers and tests can reuse the
 * exact same math (e.g. selecting a tenant's latest *scorable* reputation).
 * Thresholds are read from QBR.THRESH at call time with defaults equal to the
 * production values, so this runs correctly under Node with only this file
 * loaded (report-generator.js, which defines QBR.THRESH, need not be present)
 * and yields identical scores to the browser app. Live view concerns (band
 * colors) stay in app.js; only min/label live here. */
QBR.THI_WEIGHTS = { security: 25, risky: 25, adoption: 20, domain: 15, postmaster: 10, storage: 5 };
QBR.THI_BANDS = [
  { min: 95, label: "Excellent" },
  { min: 80, label: "Healthy" },
  { min: 60, label: "Attention Needed" },
  { min: 0,  label: "Critical" },
];
QBR.thiBand = function (v) {
  return QBR.THI_BANDS.find(function (b) { return v >= b.min; }) || QBR.THI_BANDS[QBR.THI_BANDS.length - 1];
};
QBR.thiScorers = {
  // security snapshot row → 0–100 (null = no usable security signal)
  security: function (sec) {
    if (!sec || !sec.securityDefault) return null;
    var sd = sec.securityDefault;
    var base = (sd === "ENABLED" || sd === "CONDITIONAL ACCESS") ? 80 : sd === "DISABLED" ? 35 : sd === "NOT MANAGED" ? 10 : null;
    if (base == null) return null;                       // UNKNOWN etc. → no signal
    if (sec.mfa === "YES") base += 10;
    if (sec.sspr === "YES") base += 10;
    return Math.min(100, base);
  },
  // { sum, has } → 0–100 (null = no risky reading for this tenant)
  risky: function (r) {
    if (!r || !r.has) return null;
    var T = QBR.THRESH || {}, HR = T.HIGH_RISK || 50, CR = T.CRITICAL || 100, sum = r.sum || 0;
    return sum === 0 ? 100 : sum <= 15 ? 82 : sum <= HR ? 62 : sum <= CR ? 38 : 15;
  },
  // usage % → 0–100 health curve anchored on the adoption benchmark
  adoption: function (pct) {
    if (pct == null) return null;
    var B = (QBR.THRESH && QBR.THRESH.USAGE_LOW) || 40, x = Math.max(0, Math.min(100, pct));
    return Math.round(x >= B ? 75 + (x - B) / (100 - B) * 25 : 20 + x / B * 55);
  },
  domain: function (status) {
    var map = { "Healthy": 100, "Possible Service Issues": 55, "Incomplete Setup": 45,
      "No Services Selected": 40, "Not Connected": 30, "Not Managed": 10 };
    return (status && map[status] != null) ? map[status] : null;   // N/A / End Contract / No Status → no signal
  },
  postmaster: function (rep) {
    var map = { "HIGH": 100, "MEDIUM": 75, "Verify to see health": 55, "LOW": 45, "Issues detected": 25, "BAD": 15 };
    return (rep && map[rep] != null) ? map[rep] : null;            // Not enough data / No entry → no signal
  },
  storage: function (pct) { return pct == null ? null : pct < 70 ? 100 : pct < 85 ? 85 : pct < 95 ? 55 : 25; },
};
QBR.scoreTenantHealth = function (dims) {
  dims = dims || {};
  var S = QBR.thiScorers, W = QBR.THI_WEIGHTS;
  var scores = {
    security:   S.security(dims.security),
    risky:      S.risky(dims.risky),
    adoption:   S.adoption(dims.adoption),
    domain:     S.domain(dims.domain),
    postmaster: S.postmaster(dims.postmaster),
    storage:    S.storage(dims.storage),
  };
  var wsum = 0, w = 0, present = 0;
  Object.keys(W).forEach(function (k) {
    if (scores[k] != null) { wsum += W[k] * scores[k]; w += W[k]; present++; }
  });
  if (present < 2 || !w) return null;                    // too sparse to score
  return { thi: Math.round(wsum / w), scores: scores, present: present };
};

// expose helpers for other modules / tests
QBR.util = { schoolKey, toGB, toNum, MONTHS, MONTH_TO_Q, quarterFromMonth, normQuarter,
  normDomainHealth, normSecurityDefault };

// Node/test export
if (typeof module !== "undefined" && module.exports) {
  module.exports = QBR;
}
