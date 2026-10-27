// MOS certification tiles — headless Chromium, synthetic attempts only (dummy data).
// Defaults (Excel/Word/PowerPoint), Expert-separate counts, Exam-filter-ignored
// behavior, school-filter updates, Admin toggles + persistence.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-mos-certtiles.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

// Invented fixture: Demo High + Second High, no real people or schools.
function attempts() {
  const mk = (ref, school, exam, score) => ({ ref, school, exam, score, name: "Sample, Student " + ref, date: "2026-10-20" });
  return [
    mk("R1", "Demo High", "Excel", 800), mk("R2", "Demo High", "Excel", 900), mk("R3", "Demo High", "Excel", 500),
    mk("R4", "Demo High", "Excel", null),
    mk("R5", "Demo High", "Word", 750), mk("R6", "Demo High", "Word", 600),
    mk("R7", "Demo High", "PowerPoint", 950),
    mk("R8", "Demo High", "Excel Expert", 850), mk("R9", "Demo High", "Excel Expert", 650),
    mk("R10", "Demo High", "AZ-900", 800),
    mk("R11", "Second High", "Excel", 700), mk("R12", "Second High", "Excel", 400),
  ];
}
async function bootMos(browser) {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);
  // The sidebar (and its MOS tab) only appears after the first file loads.
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx"));
  await page.waitForTimeout(2000);
  await page.click('[data-tab="dash-mos"]');
  await page.waitForTimeout(300);
  await page.evaluate(a => window.QBR.mos._loadAttempts(a, []), attempts());
  await page.waitForTimeout(500);
  return { ctx, page, errs };
}
async function tiles(page) {
  return page.evaluate(() => [...document.querySelectorAll("#mos-certtiles .mos-certtile")].map(t => ({
    label: (t.querySelector(".kpi-lbl") || {}).textContent || "",
    val: (t.querySelector(".kpi-val") || {}).textContent || "",
    sub: (t.querySelector(".mos-certtile-sub") || {}).textContent || "",
    rate: (t.querySelector(".mos-certtile-rate") || {}).textContent || "",
    w: t.getBoundingClientRect().width,
  })));
}
function tileNamed(ts, label) { return ts.filter(t => t.label === label); }

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : { args: ["--no-sandbox"] });

  // A. Defaults: exactly Excel/Word/PowerPoint, in order, correct passed + taker counts
  {
    const { ctx, page, errs } = await bootMos(browser);
    const ts = await tiles(page);
    ok(ts.map(t => t.label).join(",") === "Excel,Word,PowerPoint", "defaults: Excel,Word,PowerPoint only, in order (got " + ts.map(t => t.label).join(",") + ")");
    ok(tileNamed(ts, "Excel")[0].val === "3" && tileNamed(ts, "Excel")[0].sub === "Out of 5 takers \u00b7 60% pass rate", "Excel tile: 3 passed, 'Out of 5 takers \u00b7 60% pass rate'");
    ok(tileNamed(ts, "Excel")[0].rate === "\u00b7 60% pass rate", "Excel tile: rate span present (60%)");
    ok(tileNamed(ts, "Word")[0].val === "1" && tileNamed(ts, "Word")[0].sub === "Out of 2 takers \u00b7 50% pass rate", "Word tile: 1 passed, 'Out of 2 takers \u00b7 50% pass rate'");
    ok(tileNamed(ts, "Word")[0].rate === "\u00b7 50% pass rate", "Word tile: rate span present (50%)");
    ok(tileNamed(ts, "PowerPoint")[0].val === "1" && tileNamed(ts, "PowerPoint")[0].sub === "Out of 1 takers \u00b7 100% pass rate", "PowerPoint tile: 1 passed, 'Out of 1 takers \u00b7 100% pass rate'");
    ok(tileNamed(ts, "PowerPoint")[0].rate === "\u00b7 100% pass rate", "PowerPoint tile: rate span present (100%)");
    ok(tileNamed(ts, "Excel Expert").length === 0 && tileNamed(ts, "AZ-900").length === 0, "Excel Expert + AZ-900 hidden by default");
    ok(ts.length === 3 && Math.max(...ts.map(t => t.w)) - Math.min(...ts.map(t => t.w)) <= 1, "3 tiles share the row equally (widths within 1px)");
    ok(await page.evaluate(() => { const k = document.getElementById("mos-kpi"), c = document.getElementById("mos-certtiles"), t = document.getElementById("mos-ch-trend"); return k && c && t && (k.compareDocumentPosition(c) & 4) && (c.compareDocumentPosition(t) & 4); }), "tile row sits between the 8 KPI tiles and the trend chart");
    ok(errs.length === 0, "section A: zero page errors (got " + errs.length + ")");
    await ctx.close();
  }

  // B. Exam filter ignored by tiles (control: the Attempts KPI does move)
  {
    const { ctx, page, errs } = await bootMos(browser);
    const kpiBefore = await page.evaluate(() => document.querySelector("#mos-kpi .kpi-val").textContent);
    await page.selectOption("#mos-f-exam", "Excel");
    await page.waitForTimeout(400);
    const ts = await tiles(page);
    const kpiAfter = await page.evaluate(() => document.querySelector("#mos-kpi .kpi-val").textContent);
    ok(kpiBefore === "11" && kpiAfter === "5", "control: Attempts KPI moves 11 -> 5 under Exam=Excel (got " + kpiBefore + " -> " + kpiAfter + ")");
    ok(ts.map(t => t.label).join(",") === "Excel,Word,PowerPoint" && tileNamed(ts, "Excel")[0].val === "3" && tileNamed(ts, "Word")[0].val === "1", "tiles unchanged under Exam=Excel");
    ok(errs.length === 0, "section B: zero page errors");
    await ctx.close();
  }

  // C. School filter updates tiles (Second High has Excel rows only)
  {
    const { ctx, page, errs } = await bootMos(browser);
    await page.selectOption("#mos-f-school", "Second High");
    await page.waitForTimeout(400);
    const ts = await tiles(page);
    ok(ts.map(t => t.label).join(",") === "Excel", "school=Second High: only the Excel tile remains");
    ok(tileNamed(ts, "Excel")[0].val === "1" && tileNamed(ts, "Excel")[0].sub === "Out of 2 takers \u00b7 50% pass rate", "Second High Excel tile: 1 passed, 'Out of 2 takers \u00b7 50% pass rate'");
    ok(errs.length === 0, "section C: zero page errors");
    await ctx.close();
  }

  // D. Admin toggles: AZ-900 on via the panel, then Hide all / Show all
  {
    const { ctx, page, errs } = await bootMos(browser);
    await page.dblclick("#app-version");
    await page.waitForTimeout(300);
    await page.fill("#admin-pw-1", "s3cret");
    await page.fill("#admin-pw-2", "s3cret");
    await page.click("#admin-pw-ok");
    await page.waitForTimeout(400);
    ok(await page.evaluate(() => !!document.querySelector(".admin-mos-tiles")), "Admin panel has a MOS certification tiles section");
    ok(await page.evaluate(() => { const r = [...document.querySelectorAll(".admin-mos-tiles .admin-row")].find(x => x.textContent.includes("AZ-900")); return !!r && r.textContent.includes("Hidden"); }), "AZ-900 listed as Hidden in Admin");
    await page.evaluate(() => { const r = [...document.querySelectorAll(".admin-mos-tiles .admin-row")].find(x => (x.querySelector(".admin-row-label") || {}).textContent === "AZ-900"); r.querySelector("button").click(); });
    await page.waitForTimeout(400);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(400);
    let ts = await tiles(page);
    ok(tileNamed(ts, "AZ-900").length === 1 && tileNamed(ts, "AZ-900")[0].val === "1" && tileNamed(ts, "AZ-900")[0].sub === "Out of 1 takers \u00b7 100% pass rate", "AZ-900 tile appears after Admin toggle (1 passed, 'Out of 1 takers \u00b7 100% pass rate')");
    ok(ts.length === 4 && Math.max(...ts.map(t => t.w)) - Math.min(...ts.map(t => t.w)) <= 1, "4 tiles share the same row equally after enabling AZ-900");
    // Hide all -> row hidden; Show all -> all 5 tiles
    await page.click('[data-tab="dash-admin"]');
    await page.waitForTimeout(400);
    await page.evaluate(() => document.querySelector('.admin-mos-tiles [data-mos-tiles-all="0"]').click());
    await page.waitForTimeout(400);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(400);
    ts = await tiles(page);
    ok(ts.length === 0, "Hide all: no tiles shown");
    await page.click('[data-tab="dash-admin"]');
    await page.waitForTimeout(400);
    await page.evaluate(() => document.querySelector('.admin-mos-tiles [data-mos-tiles-all="1"]').click());
    await page.waitForTimeout(400);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(400);
    ts = await tiles(page);
    ok(ts.map(t => t.label).join(",") === "Excel,Word,PowerPoint,AZ-900,Excel Expert", "Show all: all 5 tiles, defaults first then dropdown order");
    ok(tileNamed(ts, "Excel Expert")[0].val === "1" && tileNamed(ts, "Excel Expert")[0].sub === "Out of 2 takers \u00b7 50% pass rate", "Excel Expert tile counts separately: 1 passed, 'Out of 2 takers \u00b7 50% pass rate' (not folded into Excel)");
    ok(errs.length === 0, "section D: zero page errors (got " + errs.length + ")");
    // E. Persistence: reload, re-load attempts, AZ-900 toggle survived
    const stored = await page.evaluate(() => localStorage.getItem("qbr-mos-certtiles-v1"));
    ok(stored && JSON.parse(stored)["AZ-900"] === true, "toggle persisted in localStorage (AZ-900 true)");
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(300);
    await page.evaluate(a => window.QBR.mos._loadAttempts(a, []), attempts());
    await page.waitForTimeout(500);
    ts = await tiles(page);
    ok(tileNamed(ts, "AZ-900").length === 1, "after reload + re-load: AZ-900 tile still shown (persisted toggle)");
    await ctx.close();
  }

  // F. Null-only exam: scored = 0 -> plain "Out of 0 takers", no rate span
  {
    const { ctx, page, errs } = await bootMos(browser);
    const ext = attempts().concat([{ ref: "R13", school: "Demo High", exam: "NullExam", score: null, name: "Sample, Student R13", date: "2026-10-20" }]);
    await page.evaluate(a => window.QBR.mos._loadAttempts(a, []), ext);
    await page.waitForTimeout(500);
    await page.dblclick("#app-version");
    await page.waitForTimeout(300);
    await page.fill("#admin-pw-1", "s3cret");
    await page.fill("#admin-pw-2", "s3cret");
    await page.click("#admin-pw-ok");
    await page.waitForTimeout(400);
    await page.evaluate(() => document.querySelector('.admin-mos-tiles [data-mos-tiles-all="1"]').click());
    await page.waitForTimeout(400);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(400);
    const ts = await tiles(page);
    ok(tileNamed(ts, "NullExam").length === 1, "null-only exam tile appears after Show all");
    ok(tileNamed(ts, "NullExam")[0].val === "0" && tileNamed(ts, "NullExam")[0].sub === "Out of 0 takers", "null-only tile: 0 passed, plain 'Out of 0 takers'");
    ok(tileNamed(ts, "NullExam")[0].rate === "", "null-only tile: no rate span (never NaN%)");
    ok(errs.length === 0, "section F: zero page errors");
    await ctx.close();
  }

  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
