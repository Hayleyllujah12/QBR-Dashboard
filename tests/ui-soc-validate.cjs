// SOC Risky Investigation — upload validation in the real page (headless Chromium, file://).
// Dummy data only: contoso.edu users, documentation IP ranges; header rows are the real Entra columns.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-soc-validate.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const FX = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x) + "]" : "")); };

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-socval-"));
const SIGNIN_COLS = ["Date (UTC)", "Request ID", "User agent", "Correlation ID", "User ID", "User", "Username", "User type", "Cross tenant access type", "Incoming token type", "Authentication Protocol", "Unique token identifier", "Original transfer method", "Client credential type", "Token Protection - Sign In Session", "Token Protection - Sign In Session StatusCode", "Application", "Application ID ", "App owner tenant ID", "Resource", "Resource ID ", "Resource tenant ID", "Resource owner tenant ID", "Home tenant ID", "Home tenant name", "IP address", "Location", "Status", "Sign-in error code", "Failure reason", "Client app", "Device ID", "Browser", "Operating System", "Compliant", "Managed", "Join Type", "Multifactor authentication result", "Multifactor authentication auth method", "Multifactor authentication auth detail", "Authentication requirement", "Sign-in identifier", "Session ID", "IP address (seen by resource)", "Through Global Secure Access", "Global Secure Access IP address", "Autonomous system  number", "Flagged for review", "Token issuer type", "Incoming token type", "Token issuer name", "Latency", "Conditional Access", "Managed Identity type", "Associated Resource Id", "Federated Token Id", "Federated Token Issuer"];
const AUTHDET_COLS = ["Request ID", "Date", "Authentication method", "Authentication method detail", "Succeeded", "Result detail", "Requirement"];
const q = v => /[",\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
function csv(name, cols, rows) { // rows: objects keyed by column name (missing → blank)
  const f = path.join(TMP, name);
  fs.writeFileSync(f, "﻿" + [cols.map(q).join(",")].concat(rows.map(r => cols.map(c => q(r[c] == null ? "" : r[c])).join(","))).join("\r\n") + "\r\n");
  return f;
}
function xlsxFile(name, cols, rows) {
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cols].concat(rows.map(r => cols.map(c => r[c] == null ? "" : r[c])))), "Sheet1");
  const f = path.join(TMP, name); fs.writeFileSync(f, Buffer.from(XLSX.write(wb, { type: "array", bookType: "xlsx" }))); return f;
}
const USERS = ["Test One", "Test Two", "Test Three"];
const upn = u => u.toLowerCase().replace(/ /g, ".") + "@contoso.edu";
const day = (d, h = 9) => `2026-10-${String(d).padStart(2, "0")}T${String(h).padStart(2, "0")}:00:00Z`;
function signinRows(n, tok, idp) {
  return Array.from({ length: n }, (_, i) => { const u = USERS[i % 3];
    return { "Date (UTC)": day(22 + (i % 7), 8 + (i % 9)), "Request ID": idp + i, User: u, Username: upn(u), "Incoming token type": tok,
      Application: "Office 365 Exchange Online", "IP address": "203.0.113." + (i % 40), Location: "Manila, PH", Status: i % 5 ? "Success" : "Failure",
      "Client app": "Browser", Browser: "Edge", "Operating System": "Windows", "Authentication requirement": "Multifactor authentication" }; });
}
const RISKY = csv("DEMO_OCT_RiskySignIns_2026-10-22_2026-10-29.csv", ["Date (UTC)", "User", "IP address", "Location", "Risk state"],
  USERS.map((u, i) => ({ "Date (UTC)": day(23 + i), User: u, "IP address": "198.51.100." + i, Location: "Lagos, NG", "Risk state": "At risk" })));
const RISKY_OLD = csv("DEMO_OCT_RiskySignIns_2026-10-01_2026-10-07.csv", ["Date (UTC)", "User", "IP address", "Location", "Risk state"],
  USERS.map((u, i) => ({ "Date (UTC)": `2026-10-0${2 + i}T08:00:00Z`, User: u, "IP address": "198.51.100." + i, Location: "Lagos, NG", "Risk state": "At risk" })));
const RISKYUSERS = csv("DEMO_RiskyUsers.csv", ["User ID", "User", "Username", "Risk state", "Status", "Risk processing state", "Risk last updated (UTC)"],
  [{ "User ID": "00000000-0000-0000-0000-000000000001", User: "Test One", Username: upn("Test One"), "Risk state": "At risk" }]);
const INTER = csv("DEMO_OCT_InteractiveSignIns_2026-10-22_2026-10-29.csv", SIGNIN_COLS, signinRows(42, "None", "i-"));
const NONI_ROWS = signinRows(36, "Refresh token", "n-");
const NONI = csv("DEMO_OCT_NonInteractiveSignIns_2026-10-22_2026-10-29.csv", SIGNIN_COLS, NONI_ROWS);
const NONIAD = csv("DEMO_OCT_NonInteractiveSignIns_AuthDetails_2026-10-22_2026-10-29.csv", AUTHDET_COLS,
  NONI_ROWS.map(r => ({ "Request ID": r["Request ID"], Date: r["Date (UTC)"], "Authentication method": "Previously satisfied", Succeeded: "true", "Result detail": "MFA requirement satisfied by claim in the token", Requirement: "Multifactor authentication" })));
const INTER_AD_XLSX = xlsxFile("DEMO_OCT_InteractiveSignIns_AuthDetails_2026-10-22_2026-10-29.xlsx", AUTHDET_COLS,
  [{ "Request ID": "i-1", Date: day(22), "Authentication method": "Password", Succeeded: "true", "Result detail": "Correct password", Requirement: "Single-factor authentication" }]);
const APP = csv("DEMO_OCT_ApplicationSignIns_2026-10-22_2026-10-29.csv", ["Date (UTC)", "Request ID", "Correlation ID", "Service principal ID", "Service principal name", "Credential key ID", "Credential thumbprint", "Application"], []);
const AUTHM = csv("DEMO_OCT_AuthMethodLogs_20261030_101500.csv", ["Timestamp", "Row", "UPN", "Password", "Email", "AuthenticatorApp", "SMS", "VoiceCall", "FIDO2", "WindowsHello", "TemporaryAccessPass", "MethodCount", "DefaultSignInMethod", "Status", "Message"],
  USERS.map((u, i) => ({ Timestamp: "2026-10-30 10:15:00", Row: i + 1, UPN: upn(u), Password: "True", AuthenticatorApp: i ? "True" : "False", MethodCount: i ? 2 : 0, DefaultSignInMethod: i ? "Microsoft Authenticator (Push / Number Matching)" : "N/A", Status: i ? "Success" : "NotFound", Message: i ? "" : "User not found in Entra ID" })));
const EMPTY = path.join(TMP, "DEMO_OCT_InteractiveSignIns_2026-10-22_2026-10-29 (empty).csv"); fs.writeFileSync(EMPTY, "");

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1500, height: 950 } });
  const p = await ctx.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.setInputFiles("#file-input", FX); await p.waitForTimeout(2000);
  await p.evaluate(() => document.querySelector('[data-tab="dash-soc"]').click()); await p.waitForTimeout(500);
  const card = slot => p.evaluate(s => { const c = document.getElementById("card-" + s), n = document.getElementById("val-" + s), pill = document.getElementById("pill-" + s);
    return { loaded: c.classList.contains("loaded"), bad: c.classList.contains("val-bad"), warn: c.classList.contains("val-warn"), note: n && !n.hidden ? n.textContent : "",
      move: n && n.querySelector("[data-val-move]") ? n.querySelector("[data-val-move]").textContent : "", pill: pill.classList.contains("hidden") ? "" : pill.querySelector(".filename").textContent }; }, slot);
  const status = () => p.textContent("#analyzeStatus");

  console.log("== wrong file → refused at once, with Move to");
  let t0 = Date.now();
  await p.setInputFiles("#file-all", NONI);
  await p.waitForFunction(() => { const n = document.getElementById("val-all"); return n && !n.hidden; }, null, { timeout: 3000 });
  const ms = Date.now() - t0;
  let c = await card("all");
  ok(c.bad && !c.loaded && !c.pill && /Non-interactive sign-ins download/.test(c.note), "Non-interactive file in Source 02 is refused, nothing loaded", c);
  ok(c.move === "Move to Source 04", "offers Move to Source 04", c.move);
  ok(ms < 1500, "refused straight from the header row (" + ms + " ms)");
  ok(/1 file not loaded/.test(await status()), "status bar counts the refused file");
  await p.click("#val-all [data-val-move]"); await p.waitForTimeout(1500);
  c = await card("all"); const c4 = await card("noni");
  ok(!c.bad && !c.note && c4.loaded && /NonInteractiveSignIns/.test(c4.pill), "Move to Source 04 loads it there and clears the note", [c, c4]);

  await p.setInputFiles("#file-risky", RISKYUSERS); await p.waitForTimeout(600);
  c = await card("risky");
  ok(c.bad && /Risky users list/.test(c.note) && !c.move, "Risky users list refused from Source 01 (no Move — it has no slot)", c);
  await p.setInputFiles("#file-noniad", APP); await p.waitForTimeout(600);
  c = await card("noniad");
  ok(c.bad && /app \/ managed-identity/.test(c.note), "Application sign-ins refused", c);

  console.log("== right files → load, cross-checks");
  await p.setInputFiles("#file-risky", RISKY); await p.waitForTimeout(1500);
  c = await card("risky");
  ok(c.loaded && !c.bad && !c.warn && !c.note, "Risky sign-ins load into Source 01 (refusal note cleared)", c);
  ok(await p.evaluate(() => document.getElementById("soc-tenant").value === "DEMO" && document.getElementById("soc-month").value === "OCT"), "school code + month filled in from the file name");
  await p.setInputFiles("#file-all", INTER); await p.waitForTimeout(1500);
  c = await card("all");
  ok(c.loaded && !c.warn && !c.note, "Interactive sign-ins load into Source 02 with no warnings", c);
  ok(await p.evaluate(() => !document.getElementById("analyzeBtn").disabled), "Analyze enabled");
  await p.setInputFiles("#file-all", EMPTY); await p.waitForTimeout(500);
  c = await card("all");
  ok(c.bad && /empty \(0 bytes\)/.test(c.note) && c.loaded && /InteractiveSignIns_2026/.test(c.pill), "0-byte file refused; the file already loaded stays", c);
  await p.setInputFiles("#file-all", INTER); await p.waitForTimeout(1500);
  ok(!(await card("all")).bad, "re-dropping the right file clears the refusal");

  await p.setInputFiles("#file-noniad", INTER_AD_XLSX); await p.waitForTimeout(1500);
  c = await card("noniad");
  ok(c.bad && /Interactive sign-ins AuthDetails/.test(c.note) && !c.loaded, "Excel file checked after parsing: Interactive AuthDetails refused from Source 05", c);
  await p.setInputFiles("#file-noni", NONI); await p.waitForTimeout(1500);
  await p.setInputFiles("#file-noniad", NONIAD); await p.waitForTimeout(1500);
  c = await card("noniad");
  ok(c.loaded && !c.bad && !c.warn, "Non-interactive AuthDetails load into Source 05, matching Source 04", c);

  await p.setInputFiles("#file-auth", AUTHM); await p.waitForTimeout(1500);
  if (process.env.SHOT) { await p.setInputFiles("#file-all", NONI); await p.waitForTimeout(400); await p.locator("#uploadStage").screenshot({ path: process.env.SHOT }); await p.setInputFiles("#file-all", INTER); await p.waitForTimeout(1500); }
  c = await card("auth");
  ok(c.loaded && c.warn && /1 NotFound/.test(c.note), "AuthMethodLogs load into Source 03 with an amber note for the NotFound user", c);
  ok(/1 warning on Source 03/.test(await status()), "status bar summarises the open warning", await status());
  ok(await p.evaluate(() => !document.getElementById("analyzeBtn").disabled), "warnings never block Analyze");
  await p.click("#analyzeBtn"); await p.waitForTimeout(1800);
  ok(await p.evaluate(() => /Loaded with 1 data warning/.test(document.getElementById("summaryNarrative").textContent)), "report header notes the data warning");

  console.log("== another week + reset");
  await p.setInputFiles("#file-risky", RISKY_OLD); await p.waitForTimeout(1500);
  c = await card("risky");
  ok(c.loaded && c.warn && /different week/.test(c.note), "Source 01 from another week loads with an amber warning", c);
  await p.click("#resetBtn"); await p.waitForTimeout(400);
  const after = await p.evaluate(() => [...document.querySelectorAll("#dash-soc .val-note")].every(n => n.hidden) && !document.querySelector("#dash-soc .val-warn, #dash-soc .val-bad"));
  ok(after, "Reset clears every note and card colour");
  ok(errs.length === 0, "no page errors", errs);
  await b.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL - " + (e && e.stack || e)); process.exit(1); });
