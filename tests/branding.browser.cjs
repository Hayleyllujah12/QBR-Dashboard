/* Branch-badge fallback (2026-10-25): envInfo() returns a "Branch · <label>" badge
 * when hostname detection yields null (file://, local folders, other hosts) and a
 * build-time label (window.QBR_BRANCH) or a ?branch=
 * query override exists. Existing ?main/?beta/?dev flags keep their meaning and
 * win over ?branch=; ?branch= wins over the stamped label.
 * The shipped index.html carries NO stamp (dropped in the MOS-RECON merge, see the merge
 * guide §5); the stamp cases below inject window.QBR_BRANCH with addInitScript before any
 * app script runs, so badge coverage stays without putting the stamp back into the app.
 * Dummy data only. Run: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/branding.browser.cjs */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }
(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  async function load(query, stamp) {
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    if (stamp) await page.addInitScript(l => { window.QBR_BRANCH = l; }, stamp);
    const errs = [];
    page.on("pageerror", e => errs.push(String(e.message)));
    await page.goto(APP + query, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);
    return { page, errs };
  }
  // 0. The app itself ships without a stamp
  ok(!/window\.QBR_BRANCH\s*=/.test(fs.readFileSync(path.join(ROOT, "index.html"), "utf8")), "index.html has no QBR_BRANCH stamp");
  {
    const { page } = await load("");
    ok(await page.evaluate(() => typeof window.QBR_BRANCH === "undefined" && !/Branch/.test((document.getElementById("dev-badge") || {}).textContent || "") && document.title.indexOf("[BRANCH]") !== 0), "unstamped file:// build shows no Branch badge or [BRANCH] title");
    await page.close();
  }
  // 1. Stamped label (injected) fills the file:// null case
  {
    const { page, errs } = await load("", "MOS-RECON");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && b.textContent === "Branch · MOS-RECON"; }), "stamped label -> 'Branch · MOS-RECON' badge");
    ok(await page.evaluate(() => document.title.indexOf("[BRANCH]") === 0), "title prefixed [BRANCH]");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && /not live/.test(b.title); }), "badge tip says not live");
    ok(errs.filter(e => /brand/i.test(e)).length === 0, "stamped label: zero branding errors");
    await page.close();
  }
  // 2. ?branch= override wins over the stamped label
  {
    const { page } = await load("?branch=TEST-X", "MOS-RECON");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && b.textContent === "Branch · TEST-X"; }), "?branch=TEST-X -> 'Branch · TEST-X' badge");
    await page.close();
  }
  // 3. ?branch= value is URL-decoded
  {
    const { page } = await load("?branch=feature%2Ftest");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && b.textContent === "Branch · feature/test"; }), "?branch= URL-decodes the label");
    await page.close();
  }
  // 4. ?main keeps its meaning and wins over ?branch=
  {
    const { page } = await load("?main&branch=TEST-X");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && b.textContent === "Main"; }), "?main wins over ?branch=");
    await page.close();
  }
  // 5. ?beta keeps its meaning and wins over ?branch=
  {
    const { page } = await load("?beta&branch=TEST-X");
    ok(await page.evaluate(() => { const b = document.getElementById("dev-badge"); return !!b && b.textContent === "Beta version"; }), "?beta wins over ?branch=");
    await page.close();
  }
  // 6. No label + no query -> envInfo() null (existing file:// behavior)
  {
    const { page } = await load("");
    ok(await page.evaluate(() => window.RCT_ENV() === null), "no label + no query -> envInfo() null");
    await page.close();
  }
  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
