// Content-Security-Policy check (audit 2026-10-31 M1/M4): serves the app over http with the CSP header
// from netlify.toml and drives the heavy paths — every page, MOS PDFs (pdf.js), Scan OCR (Tesseract
// wasm + blob workers), SOC analysis (blob worker), chart image export and the vendored Identity
// widget — failing on any CSP violation or page error. Synthetic data only.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-csp.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), http = require("http"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const TOML = [path.join(ROOT, "netlify.toml"), path.join(APPDIR, "netlify.toml")].find(f => fs.existsSync(f));
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x).slice(0, 400) + "]" : "")); };

const csp = (() => { const m = /Content-Security-Policy\s*=\s*"([^"]+)"/.exec(fs.readFileSync(TOML, "utf8")); return m && m[1]; })();
ok(!!csp, "netlify.toml sets a Content-Security-Policy header");
const swa = [path.join(ROOT, "staticwebapp.config.json"), path.join(APPDIR, "staticwebapp.config.json")].find(f => fs.existsSync(f));
ok(!!swa && (JSON.parse(fs.readFileSync(swa, "utf8")).globalHeaders || {})["Content-Security-Policy"] === csp, "Azure staticwebapp.config.json carries the same CSP");
if (!csp) { console.log(`\nRESULT: ${pass} passed, ${fail} failed`); process.exit(1); }
ok(!/unsafe-eval'/.test(csp.replace(/wasm-unsafe-eval/g, "")) && !/script-src[^;]*'unsafe-inline'/.test(csp), "no 'unsafe-eval' / inline scripts allowed");
ok(!/script-src[^;]*https?:/.test(csp), "scripts load only from the site itself (no third-party script hosts)");

const MIME = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".svg": "image/svg+xml",
  ".wasm": "application/wasm", ".gz": "application/gzip", ".traineddata": "application/octet-stream", ".gif": "image/gif", ".jpg": "image/jpeg" };
function serve() {
  return new Promise(resolve => {
    const srv = http.createServer((req, res) => {
      let u = decodeURIComponent(req.url.split("?")[0]); if (u === "/") u = "/index.html";
      const f = path.join(APPDIR, u);
      if (!f.startsWith(APPDIR)) { res.writeHead(403); res.end(); return; }
      fs.readFile(f, (e, d) => {
        if (e) { res.writeHead(404); res.end(); return; }
        res.writeHead(200, { "Content-Type": MIME[path.extname(f)] || "application/octet-stream", "Content-Security-Policy": csp });
        res.end(d);
      });
    });
    srv.listen(0, "127.0.0.1", () => resolve(srv));
  });
}
function pdfs(dir) { let o = []; for (const f of fs.readdirSync(dir)) { const p = path.join(dir, f); if (fs.statSync(p).isDirectory()) { if (!/^(Archive Files|Certificates)$/.test(f)) o = o.concat(pdfs(p)); } else if (/\.pdf$/i.test(f)) o.push(p); } return o.sort(); }

// minimal synthetic SOC logs (same shape as tests/ui-soc-dropall.cjs; dummy users, documentation IPs)
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-csp-"));
const q = v => /[",\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
function csv(name, cols, rows) { const f = path.join(TMP, name); fs.writeFileSync(f, "﻿" + [cols.map(q).join(",")].concat(rows.map(r => cols.map(c => q(r[c] == null ? "" : r[c])).join(","))).join("\r\n") + "\r\n"); return f; }
const USERS = ["Test One", "Test Two"], upn = u => u.toLowerCase().replace(/ /g, ".") + "@contoso.edu";
const day = (d, h = 9) => `2026-10-${String(d).padStart(2, "0")}T${String(h).padStart(2, "0")}:00:00Z`;
const SIGNIN_COLS = ["Date (UTC)", "Request ID", "User agent", "Correlation ID", "User ID", "User", "Username", "User type", "Cross tenant access type", "Incoming token type", "Authentication Protocol", "Unique token identifier", "Original transfer method", "Client credential type", "Token Protection - Sign In Session", "Token Protection - Sign In Session StatusCode", "Application", "Application ID ", "App owner tenant ID", "Resource", "Resource ID ", "Resource tenant ID", "Resource owner tenant ID", "Home tenant ID", "Home tenant name", "IP address", "Location", "Status", "Sign-in error code", "Failure reason", "Client app", "Device ID", "Browser", "Operating System", "Compliant", "Managed", "Join Type", "Multifactor authentication result", "Multifactor authentication auth method", "Multifactor authentication auth detail", "Authentication requirement", "Sign-in identifier", "Session ID", "IP address (seen by resource)", "Through Global Secure Access", "Global Secure Access IP address", "Autonomous system  number", "Flagged for review", "Token issuer type", "Incoming token type", "Token issuer name", "Latency", "Conditional Access", "Managed Identity type", "Associated Resource Id", "Federated Token Id", "Federated Token Issuer"];
const RISKY = csv("DEMO_OCT_RiskySignIns_2026-10-22_2026-10-29.csv", ["Date (UTC)", "User", "IP address", "Location", "Risk state"],
  USERS.map((u, i) => ({ "Date (UTC)": day(23 + i), User: u, "IP address": "198.51.100." + i, Location: "Lagos, NG", "Risk state": "At risk" })));
const INTER = csv("DEMO_OCT_InteractiveSignIns_2026-10-22_2026-10-29.csv", SIGNIN_COLS, Array.from({ length: 20 }, (_, i) => { const u = USERS[i % 2];
  return { "Date (UTC)": day(22 + (i % 7), 8 + (i % 9)), "Request ID": "i-" + i, User: u, Username: upn(u), "Incoming token type": "None", Application: "Office 365 Exchange Online",
    "IP address": "203.0.113." + (i % 40), Location: "Manila, PH", Status: i % 5 ? "Success" : "Failure", "Client app": "Browser", Browser: "Edge", "Operating System": "Windows",
    "Authentication requirement": "Multifactor authentication" }; }));

(async () => {
  const srv = await serve(); const BASE = "http://127.0.0.1:" + srv.address().port + "/";
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("dialog", d => d.dismiss().catch(() => {}));
  await p.addInitScript(() => {
    window.__csp = [];
    document.addEventListener("securitypolicyviolation", e => window.__csp.push(e.violatedDirective + " " + (e.blockedURI || "inline") + " @" + (e.sourceFile || "").split("/").pop() + ":" + e.lineNumber));
    window.print = () => {};
  });
  const viol = () => p.evaluate(() => window.__csp.slice());
  await p.goto(BASE + "index.html");
  ok(await p.evaluate(() => typeof QBR === "object" && typeof APP === "object"), "app boots under the CSP");

  console.log("== workbooks + every page");
  await p.setInputFiles("#file-input", ["fixture.xlsx", "SAMPLE_Lenovo_Inventory_SYNTH.xlsx", "SAMPLE_AUDIT_RICH.xlsx"].map(f => path.join(__dirname, f)));
  await p.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 90000 });
  await p.waitForTimeout(1500);
  await p.evaluate(() => document.body.classList.add("show-admin"));
  const tabs = await p.$$eval("#app-sidebar .sb-item[data-tab]", x => x.map(y => y.dataset.tab));
  for (const t of tabs) { await p.evaluate(t => document.querySelector('.sb-item[data-tab="' + t + '"]').click(), t); await p.waitForTimeout(250); }
  ok(tabs.length >= 15, "walked " + tabs.length + " pages");
  ok((await viol()).length === 0, "no CSP violation on any page", await viol());

  console.log("== chart image export (offscreen render → PNG data URLs)");
  const png = await p.evaluate(async () => { const id = Object.keys(QBR._configs || {})[0]; const u = id ? await QBR.renderChartPng(id, { width: 400, height: 200, dpr: 1 }) : ""; return String(u).slice(0, 22); });
  ok(/^data:image\/png/.test(png), "chart renders to a PNG", png);

  console.log("== MOS PDFs (pdf.js worker)");
  await p.evaluate(() => document.querySelector('[data-tab="dash-mos"]').click());
  await p.setInputFiles("#mos-pdf", pdfs(path.join(__dirname, "mos-fixture", "01 - Demo High School")));
  await p.waitForFunction(() => /\d/.test((document.getElementById("mos-body") || {}).textContent || "") && !document.getElementById("mos-body").hidden, null, { timeout: 60000 }).catch(() => {});
  await p.waitForTimeout(1500);
  ok(await p.evaluate(() => { const t = (document.getElementById("dash-mos") || {}).textContent || ""; return /9001/.test(t); }), "MOS results parsed from the PDFs");

  console.log("== Scan OCR (Tesseract wasm + blob worker)");
  await p.evaluate(() => document.querySelector('[data-tab="dash-scan"]').click());
  await p.setInputFiles("#scan-files", [path.join(__dirname, "scan-label-clean.png")]);
  await p.waitForFunction(() => QBR._scanUI && QBR._scanUI.rows.length === 1 && /done|error/.test(QBR._scanUI.rows[0].status), null, { timeout: 240000 });
  const sc = await p.evaluate(() => ({ st: QBR._scanUI.rows[0].status, sn: QBR._scanUI.rows[0].serial, ocr: QBR._scanOCR && QBR._scanOCR.mode, failed: QBR._scanOCR && QBR._scanOCR.failed }));
  ok(sc.st === "done" && !!sc.sn && !sc.failed, "label scanned (serial read, OCR engine running)", sc);

  console.log("== SOC analysis (blob worker)");
  await p.evaluate(() => document.querySelector('[data-tab="dash-soc"]').click());
  await p.setInputFiles("#soc-dropall-files", [RISKY, INTER]);
  await p.waitForFunction(() => !document.getElementById("analyzeBtn").disabled, null, { timeout: 60000 });
  await p.click("#analyzeBtn");
  await p.waitForFunction(() => /\d/.test((document.getElementById("kpiGrid") || {}).textContent || ""), null, { timeout: 60000 }).catch(() => {});
  ok(await p.evaluate(() => /\d/.test((document.getElementById("kpiGrid") || {}).textContent || "")), "SOC investigation rendered");

  console.log("== vendored Identity widget (the copy Netlify hosts load)");
  const idw = await p.evaluate(() => new Promise(res => { const s = document.createElement("script"); s.src = "libs/netlify-identity/netlify-identity-widget.js";
    s.onload = () => res(typeof window.netlifyIdentity === "object"); s.onerror = () => res("load error"); document.head.appendChild(s); }));
  ok(idw === true, "widget loads from the site under the CSP", idw);
  ok(await p.evaluate(() => !/identity\.netlify\.com/.test(document.documentElement.outerHTML)), "no reference to the third-party widget host");

  await p.waitForTimeout(500);
  const v = await viol();
  ok(v.length === 0, "no CSP violations overall", v);
  ok(errs.length === 0, "no page errors", errs.slice(0, 3));
  await b.close(); srv.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
