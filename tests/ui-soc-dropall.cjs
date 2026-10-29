// SOC Risky Investigation — "Drop all your files here" (auto-detect + route), headless Chromium, file://.
// Dummy data only: contoso.edu users, documentation IP ranges; header rows are the real Entra columns.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-soc-dropall.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const FX = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  \u2713 " : "  \u2717 ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x) + "]" : "")); };

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-socall-"));
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

const INTER_OLD = csv("DEMO_OCT_InteractiveSignIns_2026-10-15_2026-10-22.csv", SIGNIN_COLS, signinRows(10, "None", "o-"));
const SAVED = path.join(TMP, "DEMO_OCT_SOC_RiskyInvestigation_2026-10-29T09-53-01.xlsx"); fs.copyFileSync(INTER_AD_XLSX, SAVED);
// a month folder with two schools
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-socdir-"));
for (const f of [RISKY, INTER, NONI, NONIAD, AUTHM]) fs.copyFileSync(f, path.join(DIR, path.basename(f)));
fs.copyFileSync(RISKY, path.join(DIR, "OTHER_OCT_RiskySignIns_2026-10-22_2026-10-29.csv"));
fs.writeFileSync(path.join(DIR, "notes.txt"), "not a log");

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1500, height: 950 } });
  const p = await ctx.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.setInputFiles("#file-input", FX); await p.waitForTimeout(2000);
  await p.evaluate(() => document.querySelector('[data-tab="dash-soc"]').click()); await p.waitForTimeout(500);
  const sum = () => p.evaluate(() => { const e = document.getElementById("soc-dropall-sum"); return e.hidden ? "" : e.textContent; });
  const loadedCards = () => p.evaluate(() => ["risky", "all", "auth", "noni", "noniad"].filter(s => document.getElementById("card-" + s).classList.contains("loaded")).join(","));
  const pillOf = s => p.evaluate(x => document.querySelector("#pill-" + x + " .filename").textContent, s);
  const waitSum = re => p.waitForFunction(r => { const e = document.getElementById("soc-dropall-sum"); return !e.hidden && new RegExp(r).test(e.textContent); }, re.source, { timeout: 20000 });

  console.log("== Choose files: a mixed set");
  ok(await p.isVisible("#soc-dropall") && await p.isVisible("#soc-dropall-pick") && await p.isVisible("#soc-dropall-pickdir"), "drop-all bar with Choose files / Choose folder above the cards");
  await p.setInputFiles("#soc-dropall-files", [NONIAD, AUTHM, NONI, INTER_OLD, RISKYUSERS, APP, INTER_AD_XLSX, SAVED, EMPTY, INTER, RISKY]);
  await waitSum(/Loaded \d+ of 11/);
  let s = await sum();
  ok(/Loaded 5 of 11 files · skipped 6/.test(s), "5 of 11 loaded, 6 skipped", s);
  ok(await loadedCards() === "risky,all,auth,noni,noniad", "every card filled", await loadedCards());
  ok(/InteractiveSignIns_2026-10-22/.test(await pillOf("all")), "Source 02 got the file matching the set's dates, not the older week", await pillOf("all"));
  ok(/risky-user rows kept/.test(await p.textContent("#pill-noni .rows")), "Source 04 loaded after Source 01, so it was pre-filtered to the risky users");
  ok(/Risky users list/.test(s) && /app \/ managed-identity/.test(s) && /Interactive AuthDetails/.test(s) && /saved SOC report/.test(s) && /empty file/.test(s) && /another file was picked for Source 02/.test(s), "each skipped file says why", s);
  ok(/Source 01 ← DEMO_OCT_RiskySignIns/.test(s) && /Source 05 ← DEMO_OCT_NonInteractiveSignIns_AuthDetails/.test(s), "summary shows file → card", s);
  ok(await p.evaluate(() => !document.getElementById("analyzeBtn").disabled), "Analyze enabled");
  ok(await p.evaluate(() => document.getElementById("card-auth").classList.contains("val-warn")), "Phase A checks still apply (Source 03 NotFound note)");
  if (process.env.SHOT) await p.locator("#soc-dropall").screenshot({ path: process.env.SHOT });
  await p.click("#soc-dropall [data-dropall-close]");
  ok(await sum() === "", "summary can be dismissed");

  console.log("== Choose folder: two schools → pick one");
  await p.click("#resetBtn"); await p.waitForTimeout(300);
  ok(await loadedCards() === "", "Reset empties the cards");
  let dirOk = true;
  try { await p.setInputFiles("#soc-dropall-dir", DIR); } catch (e) { dirOk = false; console.log("    (folder input not supported by this Playwright: " + e.message.split("\n")[0] + ")"); }
  if (dirOk) {
    await waitSum(/which one/);
    const btns = await p.$$eval("[data-dropall-school]", b => b.map(x => x.textContent));
    ok(btns.join("|") === "DEMO · 5 files|OTHER · 1 file", "asks which school, largest set first", btns);
    ok(await loadedCards() === "", "nothing loads before the school is picked");
    await p.click('[data-dropall-school="DEMO"]');
    await waitSum(/Loaded/);
    s = await sum();
    ok(/Loaded 5 of 6 files · skipped 1 · school DEMO/.test(s) && /another school \(OTHER\)/.test(s), "loads the DEMO set; OTHER skipped; notes.txt ignored", s);
    ok(await loadedCards() === "risky,all,auth,noni,noniad", "every card filled from the folder");
  }

  console.log("== drag and drop");
  await p.click("#resetBtn"); await p.waitForTimeout(300);
  const R = fs.readFileSync(RISKY, "utf8"), I = fs.readFileSync(INTER, "utf8");
  await p.evaluate(([r, i]) => {
    const dt = new DataTransfer();
    dt.items.add(new File([r], "DEMO_OCT_RiskySignIns_2026-10-22_2026-10-29.csv", { type: "text/csv" }));
    dt.items.add(new File([i], "DEMO_OCT_InteractiveSignIns_2026-10-22_2026-10-29.csv", { type: "text/csv" }));
    const bar = document.getElementById("soc-dropall");
    bar.dispatchEvent(new DragEvent("dragover", { dataTransfer: dt, bubbles: true, cancelable: true }));
    bar.dispatchEvent(new DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true }));
  }, [R, I]);
  await waitSum(/Loaded 2 of 2/);
  ok(await loadedCards() === "risky,all", "two dropped files land in Source 01 and 02");
  await p.click("#resetBtn"); await p.waitForTimeout(300);
  ok(await sum() === "", "Reset clears the summary");
  ok(errs.length === 0, "no page errors", errs);
  await b.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL - " + (e && e.stack || e)); process.exit(1); });
