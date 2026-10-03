#!/usr/bin/env node
/* =============================================================================
 * run-tests.cjs — offline regression harness for the QBR dashboard ENGINE.
 *
 * Runs the real excel-loader.js against a tracker workbook (no browser needed)
 * and asserts the invariants that the app depends on. Purpose: before committing
 * any change to js/excel-loader.js (or the parsers), run this and confirm 0
 * failures — it catches parser / aggregation regressions that are otherwise only
 * visible by uploading the workbook in a browser.
 *
 * Zero install: uses the vendored SheetJS in qbr-app/libs (UMD works under Node).
 *
 * USAGE:
 *   node tests/run-tests.cjs "PATH\TO\ALL TENANT AUTOMATED TRACKER.xlsx"
 *   node tests/run-tests.cjs tests/fixture.xlsx      (synthetic — no sensitive data)
 *   (or set env QBR_TRACKER to the path; defaults to tests/fixture.xlsx)
 *
 * PROFILES: the harness auto-detects which workbook it was handed (the synthetic
 * fixture carries a "Fixture Control Tenant" marker) and picks the matching set
 * of SNAPSHOT expectations, so it runs clean (0 failures / 0 drift) on BOTH the
 * real tracker and the fixture. Generate the fixture with:
 *   node tests/make-fixture.cjs
 *
 * Three kinds of checks:
 *   [INVARIANT] must hold for ANY tracker — a failure is a real engine bug.
 *               Includes the SCHEMA CONTRACT (required sheets + columns) which
 *               fails loud if the source workbook is restructured.
 *   [SNAPSHOT]  pinned to a specific workbook's numbers (per profile) — if the
 *               tracker is updated these may legitimately change; update them.
 * ===========================================================================*/
const fs = require("fs"), path = require("path");
function findFile(cands) { for (const c of cands) { try { fs.accessSync(c); return c; } catch (e) {} } return null; }
const base = __dirname;
const XLSXP = findFile([base + "/../qbr-app/libs/xlsx.full.min.js", base + "/../app/libs/xlsx.full.min.js", base + "/../libs/xlsx.full.min.js"]);
const LOADERP = findFile([base + "/../qbr-app/js/excel-loader.js", base + "/../app/js/excel-loader.js", base + "/../js/excel-loader.js"]);
if (!XLSXP || !LOADERP) { console.error("Could not locate qbr-app/libs/xlsx.full.min.js or qbr-app/js/excel-loader.js next to this script."); process.exit(2); }

const trackerArg = process.argv[2] || process.env.QBR_TRACKER || (base + "/fixture.xlsx");
if (!fs.existsSync(trackerArg)) { console.error("Tracker not found: " + trackerArg + "\n(generate the fixture with: node tests/make-fixture.cjs)"); process.exit(2); }

global.window = global;
global.XLSX = require(path.resolve(XLSXP));
require(path.resolve(LOADERP));

let pass = 0, fail = 0, soft = 0;
function ok(cond, kind, msg, detail) {
  const tag = kind === "INV" ? "INVARIANT" : "SNAPSHOT";
  if (cond) { pass++; console.log(`  ✓ [${tag}] ${msg}`); }
  else if (kind === "SNAP") { soft++; console.log(`  ~ [${tag}] ${msg}  (${detail || "value drifted — update if tracker changed"})`); }
  else { fail++; console.log(`  ✗ [${tag}] ${msg}  (${detail || ""})`); }
}

const model = QBR.loadWorkbooks([fs.readFileSync(trackerArg)]);

// ---- profile: fixture vs the real tracker (the fixture carries a marker) ----
const PROFILE = model.master.has("FIXTURE CONTROL TENANT") ? "fixture" : "tracker";
const MIN = PROFILE === "fixture" ? { master: 5, clean: 8 } : { master: 80, clean: 300 };

console.log("QBR engine regression — " + path.basename(trackerArg) + "   [profile: " + PROFILE + "]\n");

/* ---- sheets found ---- */
["risky", "security", "storage", "usage", "canva", "postmaster"].forEach(s =>
  ok(model.sources[s] === true, "INV", `source sheet found: ${s}`));
ok(model.master && model.master.size > MIN.master, "INV", `master school dimension built (>${MIN.master} tenants)`, model.master && model.master.size);

/* ---- name mojibake repaired (display), source unaffected for matching ---- */
const MOJI = /[ÃÂ]/;
let mojiCount = 0;
["risky", "security", "storage", "usage", "canva", "postmaster", "pmVerify"].forEach(sh =>
  (model[sh] || []).forEach(r => { if (r.schoolRaw && MOJI.test(r.schoolRaw)) mojiCount++; }));
ok(mojiCount === 0, "INV", "no mojibake left in any display name (schoolRaw)", "found " + mojiCount);
const dasmar = [...model.master.values()].find(m => /dasmar/i.test(m.name) && /- Dasmar/i.test(m.name));
ok(!!dasmar && /ñ/.test(dasmar.name), "INV", 'accented name repaired ("Dasmariñas")', dasmar && dasmar.name);

/* ---- null vs 0 kept distinct (missing data never becomes 0) ---- */
const rNull = model.risky.some(r => r.risky === null), rZero = model.risky.some(r => r.risky === 0), rPos = model.risky.some(r => r.risky > 0);
ok(rNull && rZero && rPos, "INV", "risky keeps null / 0 / positive all distinct", `null=${rNull} zero=${rZero} pos=${rPos}`);

/* ---- storage: no physically-impossible used values; median parser sane ---- */
const maxUsedTB = Math.max.apply(null, model.storage.filter(s => s.usedGB != null).map(s => s.usedGB / 1024));
ok(maxUsedTB <= 250, "INV", "no storage row exceeds 250 TB used (impossible-capacity guard)", "max=" + maxUsedTB.toFixed(1) + " TB");

/* ---- reputation vocabulary NOT collapsed (the v1.8.5 Postmaster fix) ---- */
const reps = {}; model.postmaster.forEach(p => reps[p.reputation] = (reps[p.reputation] || 0) + 1);
ok((reps["Issues detected"] || 0) > 0 && (reps["Verify to see health"] || 0) > 0, "INV",
  "postmaster reputation preserves live states (Issues detected / Verify to see health)", JSON.stringify(reps));

/* ---- postmaster verification synthesized from the tracker Status column ---- */
const pv = {}; model.pmVerify.forEach(v => pv[v.pmTool] = (pv[v.pmTool] || 0) + 1);
ok(model.pmVerify.length > 0 && (pv["Verified"] || 0) > 0, "INV", "verification synthesized (Verified present)", JSON.stringify(pv));

/* ---- org resolution: gaps are counted, never crash ---- */
let orgGaps = 0; model.master.forEach(m => { if (!model.orgByKey[m.key]) orgGaps++; });
ok(orgGaps >= 0 && orgGaps < model.master.size, "INV", "org resolution runs; gaps counted", "gaps=" + orgGaps);

/* ---- clean dataset builds ---- */
const clean = QBR.buildCleanRows ? QBR.buildCleanRows(model) : [];
ok(clean.length > MIN.clean, "INV", `clean dataset builds (>${MIN.clean} rows)`, "rows=" + clean.length);

/* ---- SCHEMA CONTRACT: required sheets + columns present (fail loud on a
 *      restructured workbook, instead of silently zeroing a dimension). Mirrors
 *      makeResolver: a header matches a pattern by exact OR substring, case-insensitive.
 *      model.raw.<key> is the raw sheet (header row + rows) captured on load. ---- */
console.log("\n-- schema contract --");
function headerHas(rawRows, group) {
  if (!rawRows || !rawRows[0]) return false;
  const heads = rawRows[0].map(h => (h == null ? "" : String(h).trim().toLowerCase()));
  return group.some(p => { const pl = p.toLowerCase(); return heads.indexOf(pl) >= 0 || heads.some(h => h.includes(pl)); });
}
const CONTRACT = {
  risky:      [["month"], ["school_rakso", "school"], ["total risky users", "risky"], ["domain_health", "domain health"]],
  security:   [["school"], ["security default"], ["mfa"], ["sspr"], ["gdap"], ["authentication methods", "auth"]],
  storage:    [["school", "client"], ["used storage", "current storage", "usage"], ["total storage", "percentage", "usage"]],
  usage:      [["client", "school"], ["office 365 active", "active"], ["% of usage", "usage"]],
  canva:      [["school"], ["total user", "user"]],
  postmaster: [["school"], ["domain reputation", "reputation"], ["month"]],
};
Object.keys(CONTRACT).forEach(sheet => {
  const raw = model.raw && model.raw[sheet];
  ok(!!(raw && raw.length), "INV", `schema: ${sheet} raw sheet captured`, raw ? "" : "missing model.raw." + sheet);
  if (!(raw && raw.length)) return;
  CONTRACT[sheet].forEach(group =>
    ok(headerHas(raw, group), "INV", `schema: ${sheet} has [${group[0]}] column`, "missing any of " + JSON.stringify(group)));
});

/* ---- VERSION consistency: QBR.VERSION is the single source of truth for the
 *      header, and must match the deployment folder name (QBR-Dashboard_<ver>).
 *      When the harness runs from a non-versioned folder (dev sandbox), the
 *      folder check is skipped rather than failed. ---- */
console.log("\n-- version --");
ok(typeof QBR.VERSION === "string" && /^\d+\.\d+(\.\d+)?$/.test(QBR.VERSION), "INV", "QBR.VERSION defined (x.y[.z])", QBR.VERSION);
const folder = path.basename(path.resolve(base, ".."));
const fv = (folder.match(/QBR-?Dashboard[_-]?v?(\d+\.\d+(?:\.\d+)?)/i) || [])[1];
if (fv) ok(fv === QBR.VERSION, "INV", "QBR.VERSION matches deployment folder", "folder=" + folder + " version=" + QBR.VERSION);
else { soft++; console.log(`  ~ [VERSION] folder "${folder}" is not versioned — skipping folder/version match (dev layout)`); }

/* ---- Tenant Health Index scorer (pure QBR.scoreTenantHealth) ------------- */
console.log("\n-- Tenant Health Index scorer --");
const H = QBR.scoreTenantHealth, S = QBR.thiScorers, BND = QBR.thiBand;
ok(typeof H === "function" && S && typeof BND === "function", "INV", "scoreTenantHealth / thiScorers / thiBand exported");
ok(Object.values(QBR.THI_WEIGHTS).reduce((a, b) => a + b, 0) === 100, "INV", "THI weights sum to 100");
ok(QBR.THI_BANDS.every(b => b.color === undefined), "INV", "engine bands carry no color (view concern stays in app.js)");

// gating: needs >=2 present dimensions, never fabricates from one signal
ok(H({}) === null, "INV", "no dimensions → null");
ok(H({ security: { securityDefault: "ENABLED" } }) === null, "INV", "single dimension → null (>=2 required)");
const two = H({ security: { securityDefault: "ENABLED", mfa: "YES", sspr: "YES" }, storage: 50 });
ok(two && two.thi === 100 && two.present === 2, "INV", "two full-mark dims → thi 100, present 2", two && JSON.stringify(two));
// renormalization: only security(25)+adoption(20) present, weights renormalized over 45
ok(H({ security: { securityDefault: "DISABLED" }, adoption: 100 }).thi === 64, "INV",
  "missing dims renormalized, never zeroed (DISABLED+100% adoption → 64)");

// security sub-scores
ok(S.security({ securityDefault: "ENABLED" }) === 80, "INV", "security ENABLED = 80");
ok(S.security({ securityDefault: "ENABLED", mfa: "YES", sspr: "YES" }) === 100, "INV", "ENABLED + MFA + SSPR = 100");
ok(S.security({ securityDefault: "CONDITIONAL ACCESS" }) === 80, "INV", "CONDITIONAL ACCESS = 80");
ok(S.security({ securityDefault: "DISABLED" }) === 35, "INV", "security DISABLED = 35");
ok(S.security({ securityDefault: "NOT MANAGED" }) === 10, "INV", "security NOT MANAGED = 10");
ok(S.security({ securityDefault: "UNKNOWN" }) === null && S.security(null) === null, "INV", "security UNKNOWN / missing = null");
// risky sub-scores (5-band, THRESH-aligned: HIGH_RISK 50 / CRITICAL 100)
ok(S.risky({ sum: 0, has: true }) === 100 && S.risky({ sum: 15, has: true }) === 82 &&
   S.risky({ sum: 50, has: true }) === 62 && S.risky({ sum: 100, has: true }) === 38 &&
   S.risky({ sum: 101, has: true }) === 15, "INV", "risky bands 0/15/50/100/101 → 100/82/62/38/15");
ok(S.risky({ sum: 5, has: false }) === null && S.risky(null) === null, "INV", "risky with no reading = null");
// adoption curve (benchmark USAGE_LOW 40 → 75)
ok(S.adoption(40) === 75 && S.adoption(100) === 100 && S.adoption(0) === 20 && S.adoption(null) === null,
  "INV", "adoption 40/100/0 → 75/100/20; null → null");
// domain / postmaster maps
ok(S.domain("Healthy") === 100 && S.domain("Incomplete Setup") === 45 && S.domain("Not Managed") === 10 &&
   S.domain("Whatever") === null && S.domain(null) === null, "INV", "domain map + unknown/null → null");
ok(S.postmaster("HIGH") === 100 && S.postmaster("Issues detected") === 25 &&
   S.postmaster("Not enough data") === null && S.postmaster(null) === null, "INV", "postmaster map + unknown/null → null");
// storage bands (util %)
ok(S.storage(69) === 100 && S.storage(70) === 85 && S.storage(85) === 55 && S.storage(95) === 25 && S.storage(null) === null,
  "INV", "storage bands 69/70/85/95 → 100/85/55/25; null → null");
// band cutoffs (inclusive lower bound)
ok(BND(95).label === "Excellent" && BND(94).label === "Healthy" && BND(80).label === "Healthy" &&
   BND(79).label === "Attention Needed" && BND(60).label === "Attention Needed" &&
   BND(59).label === "Critical" && BND(0).label === "Critical", "INV", "band cutoffs 95/80/60 (inclusive)");

/* ---- Authentication-methods tokenizer (pure QBR.parseAuthMethods) -------- */
console.log("\n-- authentication methods --");
const AM = QBR.parseAuthMethods;
ok(typeof AM === "function" && Array.isArray(QBR.AUTH_CANON) && QBR.AUTH_CANON.length === 7, "INV", "parseAuthMethods / AUTH_CANON exported");
const eqArr = (a, b) => JSON.stringify(a) === JSON.stringify(b);
ok(eqArr(AM("MFA, Temporary Access Pass, Software OATH Tokens, Email OTP"),
  ["Temporary Access Pass", "Authenticator/OATH", "MFA", "Email OTP"]), "INV", "tokenize rich list (canonical order, OATH→Authenticator)");
ok(eqArr(AM("Passkey (FIDO2),MFA,Email OTP"), ["Passkey/FIDO2", "MFA", "Email OTP"]), "INV", "tokenize no-space delimiters + Passkey");
ok(eqArr(AM("SMS, FID02, Email OTP"), ["Passkey/FIDO2", "SMS", "Email OTP"]), "INV", "FID02 typo normalizes to FIDO2");
ok(eqArr(AM("MFA,SMS,Voice call,Email OTP"), ["MFA", "SMS", "Voice", "Email OTP"]), "INV", "Voice call recognized");
ok(eqArr(AM("Email OTP"), ["Email OTP"]), "INV", "single method");
ok(eqArr(AM("NO ACCESS"), []) && eqArr(AM(""), []) && eqArr(AM("-"), []) && eqArr(AM(null), []),
  "INV", "sentinels / blank → no methods (never fabricated)");
ok(QBR.AUTH_STRONG["MFA"] && QBR.AUTH_STRONG["Authenticator/OATH"] && !QBR.AUTH_STRONG["Email OTP"] &&
   QBR.AUTH_PHISH_RESISTANT["Passkey/FIDO2"] && QBR.AUTH_PHISH_RESISTANT["Temporary Access Pass"] && !QBR.AUTH_PHISH_RESISTANT["MFA"],
   "INV", "strength / phishing-resistant sets correct");
// every security row carries a canonical methods[] (subset of AUTH_CANON)
const canon = new Set(QBR.AUTH_CANON);
ok(model.security.every(s => Array.isArray(s.methods) && s.methods.every(x => canon.has(x))), "INV", "every security row has canonical methods[]");

/* ---- Domain-registration status (pure QBR.domainStatus) ------------------ */
console.log("\n-- domain registration --");
const DS = QBR.domainStatus, DNOW = Date.UTC(2026, 8, 25);
const ds = rec => DS(rec, DNOW);
ok(typeof DS === "function" && Array.isArray(QBR.DR_STATUS_ORDER) && QBR.DR_STATUS_ORDER.length === 11, "domainStatus / DR_STATUS_ORDER exported");
const dsEq = (rec, status, group, tier) => { const r = ds(rec); return r.status === status && r.group === group && (tier === undefined || r.tier === tier); };
ok(dsEq({ statusRaw: "Current", srcDays: 304 }, "Active", "ok"), "INV", "Current + 304d → Active/ok");
ok(dsEq({ statusRaw: "Current", srcDays: 55 }, "Expiring Soon", "attention", "Watch"), "INV", "55d → Expiring Soon/Watch");
ok(dsEq({ statusRaw: "Current", srcDays: 20 }, "Expiring Soon", "critical", "Urgent"), "INV", "20d → Expiring Soon/Urgent");
ok(dsEq({ statusRaw: "Expired", srcDays: -11, remarks: "Registration status: FOR DELETION." }, "For Deletion", "critical"), "INV", "remarks FOR DELETION beats Status Expired");
ok(dsEq({ statusRaw: "For Renewal", srcDays: 16, remarks: "FOR RENEWAL." }, "For Renewal", "critical"), "INV", "For Renewal");
ok(dsEq({ regRaw: "Error" }, "Error", "attention"), "INV", "Registration sentinel 'Error' → Error");
ok(dsEq({ regRaw: "PENDING" }, "Pending", "attention"), "INV", "sentinel 'PENDING' → Pending");
ok(dsEq({ remarks: "end contract", regRaw: "Expired" }, "End Contract", "inactive"), "INV", "end contract → End Contract");
ok(dsEq({ remarks: "…has already been deleted from the EDU.PH database.", srcDays: -380 }, "Deleted", "inactive"), "INV", "deleted remark → Deleted");
ok(dsEq({ remarks: "x.edu.ph is not a valid domain name", regRaw: "Error" }, "Invalid Domain", "attention"), "INV", "invalid domain → Invalid Domain");
ok(dsEq({ regMs: null, statusRaw: "", remarks: null, regRaw: null }, "Not Registered", "inactive"), "INV", "no data → Not Registered");
ok(dsEq({ srcDays: -380 }, "Expired", "critical"), "INV", "negative days (no remarks) → Expired");
ok(dsEq({ expMs: DNOW + 10 * 86400000 }, "Expiring Soon", "critical", "Urgent"), "INV", "expMs vs now fallback → Expiring Urgent");
// when a domain sheet is present, every row must carry a canonical status
if (model.domainreg && model.domainreg.length) {
  const canonDR = new Set(QBR.DR_STATUS_ORDER);
  ok(model.domainreg.every(d => canonDR.has(d.status)), "INV", "every domainreg row has a canonical status");
}

/* ---- SNAPSHOTS (per profile) -------------------------------------------- */
console.log("\n-- snapshots [" + PROFILE + "] --");
const EXP = {
  tracker: { canvaQ1: 52, canvaQ2: 64, orgGaps: 12, pmVerify: 106, thiIndex: 69, thiScored: 109, thiDist: { Excellent: 0, Healthy: 44, "Attention Needed": 44, Critical: 21 },
             amEmailOtp: 90, amMfa: 41, amPhish: 33, amWeakOnly: 48, amOtpOnly: 37 },
  fixture: { canvaQ1: 2,  canvaQ2: 3,  orgGaps: 2,  pmVerify: 7,   thiIndex: 73, thiScored: 8,   thiDist: { Excellent: 0, Healthy: 5,  "Attention Needed": 0,  Critical: 3  },
             amEmailOtp: 7,  amMfa: 4,  amPhish: 3,  amWeakOnly: 2,  amOtpOnly: 2,
             drRows: 8, drActive: 1, drAction: 6 },
}[PROFILE];
const canvaActive = q => { const k = new Set(); model.canva.forEach(c => { if (c.quarter === q && c.users > 0) k.add(c.key); }); return k.size; };
ok(canvaActive("Q1") === EXP.canvaQ1, "SNAP", `Canva active schools Q1 = ${EXP.canvaQ1}`, "got " + canvaActive("Q1"));
ok(canvaActive("Q2") === EXP.canvaQ2, "SNAP", `Canva active schools Q2 = ${EXP.canvaQ2}`, "got " + canvaActive("Q2"));
ok(orgGaps === EXP.orgGaps, "SNAP", `tenants with no org = ${EXP.orgGaps}`, "got " + orgGaps);
ok(model.pmVerify.length === EXP.pmVerify, "SNAP", `synthesized verification records = ${EXP.pmVerify}`, "got " + model.pmVerify.length);

// authentication-method counts + posture presets (drive the Security-detail filter)
const amCount = mth => model.security.filter(s => s.methods && s.methods.indexOf(mth) >= 0).length;
const amPreset = fn => model.security.filter(s => fn(s.methods || [])).length;
ok(amCount("Email OTP") === EXP.amEmailOtp, "SNAP", `schools including Email OTP = ${EXP.amEmailOtp}`, "got " + amCount("Email OTP"));
ok(amCount("MFA") === EXP.amMfa, "SNAP", `schools including MFA = ${EXP.amMfa}`, "got " + amCount("MFA"));
ok(amPreset(set => set.some(x => QBR.AUTH_PHISH_RESISTANT[x])) === EXP.amPhish, "SNAP", `phishing-resistant schools = ${EXP.amPhish}`, "got " + amPreset(set => set.some(x => QBR.AUTH_PHISH_RESISTANT[x])));
ok(amPreset(set => set.length > 0 && !set.some(x => QBR.AUTH_STRONG[x])) === EXP.amWeakOnly, "SNAP", `weak-only schools = ${EXP.amWeakOnly}`, "got " + amPreset(set => set.length > 0 && !set.some(x => QBR.AUTH_STRONG[x])));
ok(amPreset(set => set.length === 1 && set[0] === "Email OTP") === EXP.amOtpOnly, "SNAP", `Email-OTP-only schools = ${EXP.amOtpOnly}`, "got " + amPreset(set => set.length === 1 && set[0] === "Email OTP"));

// domain registration (present only when a domain sheet is loaded — the fixture)
if (EXP.drRows != null) {
  const dr = model.domainreg;
  const drAction = dr.filter(d => d.group === "critical" || d.group === "attention").length;
  ok(dr.length === EXP.drRows, "SNAP", `domain-registration rows = ${EXP.drRows}`, "got " + dr.length);
  ok(dr.filter(d => d.status === "Active").length === EXP.drActive, "SNAP", `active domains = ${EXP.drActive}`, "got " + dr.filter(d => d.status === "Active").length);
  ok(drAction === EXP.drAction, "SNAP", `domains needing action = ${EXP.drAction}`, "got " + drAction);
}

/* ---- Portfolio THI over the workbook (no filters) — SNAPSHOT -------------
 * Replicates app.js computeHealthIndex data-shaping (all quarters/orgs/schools)
 * and scores every tenant through the pure engine. Pins the whole-portfolio
 * result so any scoring/parse regression is caught here, not in the browser. */
(function () {
  function dedupe(rows, mf) { const b = new Map(); rows.forEach(r => { const c = b.get(r.key); if (!c || (mf(r) || 0) > (mf(c) || 0)) b.set(r.key, r); }); return [...b.values()]; }
  const riskySum = {}, riskyHas = {}, domLatest = {};
  model.risky.forEach(r => {
    if (r.risky != null) { riskySum[r.key] = (riskySum[r.key] || 0) + r.risky; riskyHas[r.key] = true; }
    if (r.health) { const e = domLatest[r.key]; if (!e || r.monthIdx >= e.mi) domLatest[r.key] = { s: r.health, mi: r.monthIdx }; }
  });
  const usageByKey = {}; dedupe(model.usage.filter(u => u.usagePct != null), u => u.usagePct).forEach(u => usageByKey[u.key] = u.usagePct);
  const storByKey = {}; dedupe(model.storage.filter(s => s.pct != null), s => s.usedGB).forEach(s => storByKey[s.key] = s.pct);
  const secByKey = {}; model.security.forEach(s => { if (!secByKey[s.key]) secByKey[s.key] = s; });
  // latest SCORABLE reputation per school (mirrors app.js latestScorableReputation; fixed 2026-10-03 —
  // the old one-liner compared monthIdx against e.mi on a bare string, so the FIRST scorable month always won)
  const latestRep = rows => { const out = {}; rows.forEach(p => { if (S.postmaster(p.reputation) == null) return;
    const mi = p.monthIdx == null ? -1 : p.monthIdx, e = out[p.key]; if (!e || mi >= e.mi) out[p.key] = { rep: p.reputation, mi }; }); return out; };
  { const t = latestRep([{ key: "X", reputation: "HIGH", monthIdx: 0 }, { key: "X", reputation: "Issues detected", monthIdx: 2 }, { key: "X", reputation: "Not enough data", monthIdx: 3 }, { key: "Y", reputation: "BAD", monthIdx: null }, { key: "Y", reputation: "MEDIUM", monthIdx: 1 }]);
    ok(t.X.rep === "Issues detected" && t.Y.rep === "MEDIUM", "INV", "THI postmaster uses the LATEST scorable month (not the first)", JSON.stringify(t)); }
  const pmLatest = latestRep(model.postmaster); const pmByKey = {}; Object.keys(pmLatest).forEach(k => pmByKey[k] = pmLatest[k].rep);
  const keys = new Set(); [model.risky, model.usage, model.storage, model.postmaster].forEach(a => a.forEach(r => keys.add(r.key))); model.security.forEach(s => keys.add(s.key));
  const per = []; keys.forEach(key => {
    if (!key) return;
    const res = H({ security: secByKey[key], risky: { sum: riskySum[key] || 0, has: !!riskyHas[key] },
      adoption: usageByKey[key], domain: domLatest[key] && domLatest[key].s, postmaster: pmByKey[key], storage: storByKey[key] });
    if (res) per.push(res.thi);
  });
  const dist = { "Excellent": 0, "Healthy": 0, "Attention Needed": 0, "Critical": 0 };
  per.forEach(v => dist[BND(v).label]++);
  const index = per.length ? Math.round(per.reduce((a, b) => a + b, 0) / per.length) : null;
  ok(per.length > 2 && per.every(v => v >= 0 && v <= 100), "INV", "portfolio THI: every tenant score in 0..100", "scored=" + per.length);
  ok(index === EXP.thiIndex, "SNAP", `portfolio Tenant Health Index = ${EXP.thiIndex}`, "got " + index);
  ok(per.length === EXP.thiScored, "SNAP", `tenants scored (>=2 dims) = ${EXP.thiScored}`, "got " + per.length);
  const d = EXP.thiDist;
  ok(dist.Critical === d.Critical && dist["Attention Needed"] === d["Attention Needed"] && dist.Healthy === d.Healthy && dist.Excellent === d.Excellent,
    "SNAP", `THI band distribution ${d.Excellent}/${d.Healthy}/${d["Attention Needed"]}/${d.Critical}`, JSON.stringify(dist));
})();

console.log(`\nRESULT: ${pass} passed, ${fail} failed, ${soft} snapshot-drift.`);
process.exit(fail ? 1 : 0);
