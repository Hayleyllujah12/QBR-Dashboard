// SOC Risky Investigation upload validation (js/soc-validate.js), plain Node, zero install.
// Header rows are the real Entra / script download headers (column names only); every data row is
// dummy (contoso.edu users, documentation IP ranges). Usage: node tests/soc-validate.cjs
"use strict";
const path = require("path");
const V = require(path.join(__dirname, "..", "js", "soc-validate.js"));
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x) + "]" : "")); };

const SIGNIN_COLS = "Date (UTC)\tRequest ID\tUser agent\tCorrelation ID\tUser ID\tUser\tUsername\tUser type\tCross tenant access type\tIncoming token type\tAuthentication Protocol\tUnique token identifier\tOriginal transfer method\tClient credential type\tToken Protection - Sign In Session\tToken Protection - Sign In Session StatusCode\tApplication\tApplication ID \tApp owner tenant ID\tResource\tResource ID \tResource tenant ID\tResource owner tenant ID\tHome tenant ID\tHome tenant name\tIP address\tLocation\tStatus\tSign-in error code\tFailure reason\tClient app\tDevice ID\tBrowser\tOperating System\tCompliant\tManaged\tJoin Type\tMultifactor authentication result\tMultifactor authentication auth method\tMultifactor authentication auth detail\tAuthentication requirement\tSign-in identifier\tSession ID\tIP address (seen by resource)\tThrough Global Secure Access\tGlobal Secure Access IP address\tAutonomous system  number\tFlagged for review\tToken issuer type\tIncoming token type\tToken issuer name\tLatency\tConditional Access\tManaged Identity type\tAssociated Resource Id\tFederated Token Id\tFederated Token Issuer";
const AUTHDET_COLS = "Request ID\tDate\tAuthentication method\tAuthentication method detail\tSucceeded\tResult detail\tRequirement";
const APP_COLS = "Date (UTC)\tRequest ID\tCorrelation ID\tService principal ID\tService principal name\tCredential key ID\tCredential thumbprint\tApplication\tApplication ID \tApp owner tenant ID\tResource\tResource ID \tResource tenant ID\tResource owner tenant ID\tHome tenant ID\tHome tenant name\tIP address\tLocation\tStatus\tSign-in error code\tFailure reason\tConditional Access";
const RISKY_COLS = "Date (UTC),User,IP address,Location,Risk state";
const RISKYUSERS_COLS = "User ID,User,Username,Risk state,Status,Risk processing state,Risk last updated (UTC)";
const AUTHM_COLS = "Timestamp,Row,UPN,Password,Email,AuthenticatorApp,SMS,VoiceCall,FIDO2,WindowsHello,TemporaryAccessPass,MethodCount,DefaultSignInMethod,Status,Message";
const H = s => V.parseHeaderLine(s);
const SLOTS = ["risky", "all", "auth", "noni", "noniad"];
const nm = (kind, from = "2026-10-22", to = "2026-10-29", pre = "DEMO_OCT_") => `${pre}${kind}_${from}_${to}.csv`;
function matrix(cols, file) { const c = V.classify(H(cols), file); return SLOTS.map(s => { const r = V.slotCheck(s, c); return s + ":" + (r.ok ? "ok" : "x" + (r.moveTo || "")); }).join(" "); }

console.log("== header row");
ok(H("﻿\"Date (UTC)\",\"User\",\"Risk state\"\r\n1,2,3").join("|") === "Date (UTC)|User|Risk state", "BOM + quoted cells + CRLF");
ok(H("a\tb\tc").length === 3 && H("a;b;c").length === 3, "tab and semicolon delimiters");
ok(H('"x ""y"", z",b').join("|") === 'x "y", z|b', "escaped quotes and commas inside quotes");
ok(H(SIGNIN_COLS).length === 57, "the 57-column sign-in header (duplicate Incoming token type kept)");
ok(V.headersOfRows([{ "Incoming token type": 1, "Incoming token type_1": 2, User: 3 }]).join("|") === "Incoming token type|User", "SheetJS _1 duplicate column folded back");

console.log("== which file is it (real headers)");
const cases = [
  [RISKY_COLS, nm("RiskySignIns"), "riskySignins", null, "risky:ok all:xrisky auth:xrisky noni:xrisky noniad:xrisky"],
  [RISKYUSERS_COLS, "DEMO_RiskyUsers.csv", "riskyUsers", null, "risky:x all:x auth:x noni:x noniad:x"],
  [SIGNIN_COLS, nm("InteractiveSignIns"), "signins", "int", "risky:xall all:ok auth:xall noni:xall noniad:xall"],
  [SIGNIN_COLS, nm("NonInteractiveSignIns"), "signins", "non", "risky:xnoni all:xnoni auth:xnoni noni:ok noniad:xnoni"],
  [SIGNIN_COLS, "renamed-log.csv", "signins", null, "risky:xall all:ok auth:xall noni:ok noniad:xall"],
  [AUTHDET_COLS, nm("InteractiveSignIns_AuthDetails"), "authDetails", "int", "risky:x all:x auth:x noni:x noniad:x"],
  [AUTHDET_COLS, nm("NonInteractiveSignIns_AuthDetails"), "authDetails", "non", "risky:xnoniad all:xnoniad auth:xnoniad noni:xnoniad noniad:ok"],
  [AUTHDET_COLS, "details.csv", "authDetails", null, "risky:xnoniad all:xnoniad auth:xnoniad noni:xnoniad noniad:ok"],
  [APP_COLS, nm("ApplicationSignIns"), "appSignins", null, "risky:x all:x auth:x noni:x noniad:x"],
  [APP_COLS, nm("ManagedIdentitySignIns"), "appSignins", null, "risky:x all:x auth:x noni:x noniad:x"],
  [AUTHM_COLS, "DEMO_OCT_AuthMethodLogs_20261030_101500.csv", "authMethods", null, "risky:xauth all:xauth auth:ok noni:xauth noniad:xauth"],
];
for (const [cols, file, kind, variant, m] of cases) {
  const c = V.classify(H(cols), file);
  ok(c.kind === kind && c.variant === variant, file + " → " + kind + (variant ? "/" + variant : ""), [c.kind, c.variant]);
  ok(matrix(cols, file) === m, "  slots: " + m, matrix(cols, file));
}
const rj = (cols, file, slot) => V.slotCheck(slot, V.classify(H(cols), file));
ok(/Risky users list/.test(rj(RISKYUSERS_COLS, "x.csv", "risky").message), "Risky users list in Source 01 is named as such");
ok(/Interactive sign-ins AuthDetails/.test(rj(AUTHDET_COLS, nm("InteractiveSignIns_AuthDetails"), "noniad").message), "Interactive AuthDetails refused from Source 05");
ok(/per-user AuthMethodLogs/.test(rj(AUTHDET_COLS, nm("InteractiveSignIns_AuthDetails"), "auth").message), "AuthDetails in Source 03 → points to the AuthMethodLogs file");
ok(/no user/.test(rj(APP_COLS, nm("ApplicationSignIns"), "all").message), "App sign-ins refused (no user)");
ok(rj(SIGNIN_COLS, nm("NonInteractiveSignIns"), "all").moveTo === "noni", "Non-interactive in Source 02 → Move to Source 04");
ok(V.slotCheck("all", V.classify(H("Date,User,IP address"), "x.csv")).ok, "a reduced sign-in export (user + date) is still a sign-in log");
const unk = V.slotCheck("noniad", V.classify(H("Request ID,Authentication method,Foo"), "x.csv"));
ok(unk.ok && unk.warn, "unrecognised layout with the needed columns loads with a warning");
ok(!V.slotCheck("risky", V.classify(H("Foo,Bar"), "x.csv")).ok, "unrecognised layout without the needed columns is refused");
ok(/empty/.test(V.slotCheck("all", V.classify(H(SIGNIN_COLS), "x.csv"), { size: 0 }).message), "0-byte file refused");

console.log("== file names");
const pn = V.parseName;
ok(pn("DEMO_OCT_NonInteractiveSignIns_AuthDetails_2026-10-22_2026-10-29.csv").variant === "non", "NonInteractive AuthDetails not mistaken for Interactive");
ok(pn("DEMO_OCT_InteractiveSignIns_AuthDetails_2026-10-22_2026-10-29.csv").kind === "authDetails", "Interactive AuthDetails");
let p = pn("ABC_AUGInteractiveSignIns_2026-08-28_2026-09-04.csv");
ok(p.school === "ABC" && p.mon === "AUG" && p.from === "2026-08-28" && p.to === "2026-09-04", "glued prefix ABC_AUGInteractiveSignIns", p);
p = pn("DEMO_SCHOOL_OCT_RiskySignIns_2026-10-22_2026-10-29.csv");
ok(p.school === "DEMO" && p.month === 9, "multi-word school keeps the first word as the code", p);
p = pn("DEMO_OCTOBER_RISKY_SIGNINS.csv");
ok(p.kind === "riskySignins" && p.mon === "OCTOBER" && !p.from, "no dates in the name is fine", p);

console.log("== cross-file checks");
const D = s => new Date(s).toISOString();
function signins(n, from, users, opts = {}) {
  const rows = [];
  for (let i = 0; i < n; i++) {
    const u = users[i % users.length];
    rows.push({ "Date (UTC)": D(Date.parse(from) + (i % 7) * 86400000), "Request ID": (opts.idp || "r") + i, User: u, Username: u.toLowerCase().replace(/ /g, ".") + "@" + (opts.dom || "contoso.edu"),
      "Incoming token type": opts.tok || "None", "IP address": "203.0.113." + (i % 50) });
  }
  return rows;
}
const RU = ["Test One", "Test Two", "Test Three"];
function base(over = {}) {
  const risky = RU.map((u, i) => ({ "Date (UTC)": D("2026-10-23T08:00:00Z"), User: u, "IP address": "198.51.100." + i, Location: "Manila, PH", "Risk state": "At risk" }));
  const f = {
    risky: { name: pn(nm("RiskySignIns")), kind: "riskySignins", stats: V.fileStats("riskySignins", risky) },
    all: { name: pn(nm("InteractiveSignIns")), kind: "signins", variant: "int", stats: V.fileStats("signins", signins(60, "2026-10-22T09:00:00Z", RU)) },
  };
  return Object.assign(f, over);
}
let W = V.crossCheck(base());
ok(Object.keys(W).length === 0, "a clean Source 01 + 02 pair gives no warnings", W);
W = V.crossCheck(base({ risky: { name: pn(nm("RiskySignIns", "2026-10-15", "2026-10-22")), kind: "riskySignins", stats: V.fileStats("riskySignins", [{ "Date (UTC)": D("2026-10-10T08:00:00Z"), User: "Test One", "Risk state": "At risk" }]) } }));
ok((W.risky || []).some(m => /different week/.test(m)), "a file from another week is flagged", W);
W = V.crossCheck(base({ risky: { name: pn(nm("RiskySignIns", "2026-10-22", "2026-10-29", "OTHER_OCT_")), kind: "riskySignins", stats: base().risky.stats } }));
ok((W.risky || []).some(m => /School code OTHER/.test(m)), "a different school code is flagged", W);
W = V.crossCheck(base({ risky: { name: pn(nm("RiskySignIns", "2026-10-22", "2026-10-29", "DEMO_JUN_")), kind: "riskySignins", stats: base().risky.stats } }));
ok((W.risky || []).some(m => /says JUN/.test(m)), "month label that matches neither the dates nor the month before is flagged", W);
W = V.crossCheck(base({ risky: { name: pn(nm("RiskySignIns", "2026-10-22", "2026-10-29", "DEMO_SEP_")), kind: "riskySignins", stats: base().risky.stats } }));
ok(!(W.risky || []).some(m => /says SEP/.test(m)), "a SEP report covering early October is accepted (reporting month)", W);
W = V.crossCheck(base({ all: { name: pn(nm("InteractiveSignIns", "2026-10-01", "2026-10-07")), kind: "signins", stats: base().all.stats } }));
ok((W.all || []).some(m => /file name says 2026-10-01/.test(m)), "rows outside the dates in the file name are flagged", W);
const cap = base(); cap.all.stats = Object.assign({}, cap.all.stats, { totalRows: 100000 });
W = V.crossCheck(cap);
ok((W.all || []).some(m => /100,000 rows .* cut off/.test(m)), "a file at the 100,000-row download limit is flagged", W);
W = V.crossCheck(base({ all: { name: pn(nm("InteractiveSignIns")), kind: "signins", stats: V.fileStats("signins", signins(60, "2026-10-22T09:00:00Z", ["Someone Else"])) } }));
ok((W.all || []).some(m => /Only 0 of 3 at-risk users/.test(m)), "risky users missing from the interactive log are flagged", W);
W = V.crossCheck(base({ all: { name: pn("renamed.csv"), kind: "signins", stats: V.fileStats("signins", signins(60, "2026-10-22T09:00:00Z", RU, { tok: "Refresh token" })) } }));
ok((W.all || []).some(m => /looks like a Non-interactive/.test(m)), "Source 02 that mostly reuses refresh tokens looks non-interactive", W);
const noniInt = base({ noni: { name: pn("renamed2.csv"), kind: "signins", stats: V.fileStats("signins", signins(40, "2026-10-22T09:00:00Z", RU, { idp: "n" })) } });
W = V.crossCheck(noniInt);
ok((W.noni || []).some(m => /looks like an Interactive/.test(m)), "Source 04 with no incoming tokens looks interactive", W);

const noniOk = signins(40, "2026-10-22T09:00:00Z", RU, { tok: "Primary refresh token", idp: "n" });
const withNoni = base({ noni: { name: pn(nm("NonInteractiveSignIns")), kind: "signins", variant: "non", stats: V.fileStats("signins", noniOk) },
  noniad: { name: pn(nm("NonInteractiveSignIns_AuthDetails")), kind: "authDetails", variant: "non", stats: V.fileStats("authDetails", []) } });
const ids = o => { const r = {}; o.forEach(x => { r[x] = 1; }); return r; };
W = V.crossCheck(withNoni, { authMapIds: ids(noniOk.map(r => r["Request ID"])) });
ok(!W.noni && !W.noniad, "Source 04 + matching Source 05: no warnings", W);
W = V.crossCheck(withNoni, { authMapIds: ids(Array.from({ length: 60 }, (_, i) => "r" + i)) });
ok((W.noniad || []).some(m => /match the interactive sign-ins/.test(m)), "Source 05 whose Request IDs match Source 02 is called out", W);
W = V.crossCheck(withNoni, { authMapIds: ids(["n1", "n2"]) });
ok((W.noniad || []).some(m => /only 5%/.test(m)), "low Source 05 coverage of Source 04 is flagged", W);
W = V.crossCheck(base({ noni: { name: pn(nm("NonInteractiveSignIns")), kind: "signins", variant: "non", stats: V.fileStats("signins", signins(30, "2026-10-22T09:00:00Z", RU, { tok: "Refresh token", dom: "other.edu", idp: "n" })) } }));
ok((W.noni || []).some(m => /No email domain in common/.test(m)), "a file whose users are on another email domain is flagged", W);

console.log("== Source 03 (AuthMethodLogs)");
const am = (rows) => ({ name: pn("DEMO_OCT_AuthMethodLogs_20261030_101500.csv"), kind: "authMethods", stats: V.fileStats("authMethods", rows) });
const amRow = (u, st = "Success", ts = "2026-10-30 10:15:00") => ({ Timestamp: ts, UPN: u.toLowerCase().replace(/ /g, ".") + "@contoso.edu", Status: st });
W = V.crossCheck(base({ auth: am(RU.map(u => amRow(u))) }));
ok(!W.auth, "every at-risk user found, extracted the next day: no warnings", W);
W = V.crossCheck(base({ auth: am([amRow("Test One"), amRow("Test Two", "NotFound")]) }));
ok((W.auth || []).some(m => /1 NotFound/.test(m)) && (W.auth || []).some(m => /Only 2 of 3 at-risk users/.test(m)), "NotFound rows and missing risky users are flagged", W);
W = V.crossCheck(base({ auth: am(RU.map(u => amRow(u, "Success", "2026-11-20 09:00:00"))) }));
ok((W.auth || []).some(m => /days after the investigation window/.test(m)), "auth methods extracted weeks after the window are flagged", W);

console.log(pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
