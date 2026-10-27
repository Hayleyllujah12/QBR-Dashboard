// MOS completion certificates — signature auto-save + downscale, and shared
// Results filters surfaced in the Completion view (headless Chromium).
// Dummy data only.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-mos-completion-filters-sig.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs");
const path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

// Failed-twice students: SANTOS (Demo High, Excel, Proctor A, Oct 2026),
// REYES (Demo High, Word, Proctor B, Oct 2026), CRUZ (Second High, Excel,
// Proctor A, Sep 2026). LIM passes once -> excluded.
function attempts() {
  const mk = (ref, school, exam, score, date, month, proctor, name) =>
    ({ ref, school, exam, score, date, month, proctor, name });
  return [
    mk("A1", "Demo High", "Excel", 500, "2026-10-01", "2026-10", "Proctor A", "Maria Santos"),
    mk("A2", "Demo High", "Excel", 600, "2026-10-20", "2026-10", "Proctor A", "Maria Santos"),
    mk("B1", "Demo High", "Word", 400, "2026-10-02", "2026-10", "Proctor B", "Jose Reyes"),
    mk("B2", "Demo High", "Word", 650, "2026-10-21", "2026-10", "Proctor B", "Jose Reyes"),
    mk("C1", "Second High", "Excel", 500, "2026-09-03", "2026-09", "Proctor A", "Ana Cruz"),
    mk("C2", "Second High", "Excel", 600, "2026-09-22", "2026-09", "Proctor A", "Ana Cruz"),
    mk("D1", "Demo High", "PowerPoint", 800, "2026-10-04", "2026-10", "Proctor A", "Jane Roe"),
  ];
}
async function boot(browser) {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message)));
  page.on("dialog", d => d.dismiss());
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx"));
  await page.waitForTimeout(2000);
  await page.click('[data-tab="dash-mos"]');
  await page.waitForTimeout(300);
  await page.evaluate(a => window.QBR.mos._loadAttempts(a, []), attempts());
  await page.waitForTimeout(500);
  await page.click('[data-mos-view="completion"]');
  await page.waitForTimeout(500);
  return { ctx, page, errs };
}
async function batchText(page) {
  return page.evaluate(() => (document.getElementById("mos-view-completion") || {}).textContent || "");
}
async function compMsg(page) {
  return page.evaluate(() => (document.getElementById("mos-comp-msg") || {}).textContent || "");
}
async function sigThumb(page) {
  return page.evaluate(() => !!document.querySelector('#mos-view-completion img[src^="data:image/"]'));
}
async function storedS1Len(page) {
  return page.evaluate(() => { try { return (JSON.parse(localStorage.getItem("qbr-mos-completion-sig-v1") || "{}").s1 || "").length; } catch (e) { return -1; } });
}
async function storedS1IsPng(page) {
  return page.evaluate(() => { try { return (JSON.parse(localStorage.getItem("qbr-mos-completion-sig-v1") || "{}").s1 || "").indexOf("data:image/png;base64,") === 0; } catch (e) { return false; } });
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : { args: ["--no-sandbox"] });

  // A. Signature auto-saves on pick (no Save click) and survives leaving the view.
  {
    const { ctx, page, errs } = await boot(browser);
    let t = await batchText(page);
    ok(/3 certificates/.test(t), "baseline batch: 3 certificates");
    ok(await page.evaluate(() => !!document.getElementById("mos-comp-f-school") && !!document.getElementById("mos-comp-f-proctor") && !!document.getElementById("mos-comp-f-pmode") && !!document.getElementById("mos-comp-f-pval")), "filter controls present (school, proctor, period mode+value)");
    await page.setInputFiles("#mos-comp-s1", path.join(ROOT, "tests", "comp-sig.png"));
    await page.waitForTimeout(800);
    ok(await sigThumb(page), "A: thumbnail appears on pick, without Save");
    ok(/Signature image saved/.test(await compMsg(page)), "A: auto-save message (got " + await compMsg(page) + ")");
    await page.click('[data-mos-view="results"]');
    await page.waitForTimeout(400);
    await page.click('[data-mos-view="completion"]');
    await page.waitForTimeout(500);
    ok(await sigThumb(page), "A: thumbnail survives leaving + returning (old bug: pick was wiped)");
    // B. Save after auto-save keeps the image (no stale-snapshot clobber).
    await page.fill("#mos-comp-n1", "DUMMY ONE");
    await page.click("#mos-comp-sigsave");
    await page.waitForTimeout(300);
    ok(/Signatories saved/.test(await compMsg(page)), "B: save message");
    ok(await sigThumb(page), "B: thumbnail still present after Save");
    ok((await storedS1Len(page)) > 100, "B: stored s1 non-empty after Save");
    ok(errs.length === 0, "sections A-B: zero page errors (got " + errs.length + ")");
    await ctx.close();
  }

  // C. Big phone photo is downscaled and saves (no quota failure).
  {
    const { ctx, page, errs } = await boot(browser);
    await page.setInputFiles("#mos-comp-s1", path.join(ROOT, "tests", "sig-big.png"));
    await page.waitForTimeout(2000);
    const msg = await compMsg(page);
    ok(/Signature image saved/.test(msg), "C: big photo saves (msg: " + msg + ")");
    ok(!/storage may be full/.test(msg), "C: no quota failure");
    const len = await storedS1Len(page);
    ok(len > 1000 && len < 1500000, "C: stored image downscaled, under quota (chars: " + len + ")");
    ok(await storedS1IsPng(page), "C: stored as PNG (transparency preserved)");
    ok(errs.length === 0, "section C: zero page errors");
    await ctx.close();
  }

  // D. Shared filters narrow the batch; Results tab sees the same values.
  {
    const { ctx, page, errs } = await boot(browser);
    await page.selectOption("#mos-comp-f-exam", "Word");
    await page.waitForTimeout(400);
    ok(/1 certificate/.test(await batchText(page)), "D: exam=Word -> 1 certificate");
    await page.click("#mos-comp-f-reset");
    await page.waitForTimeout(400);
    ok(/3 certificates/.test(await batchText(page)), "D: reset -> 3 certificates");
    await page.selectOption("#mos-comp-f-proctor", "Proctor B");
    await page.waitForTimeout(400);
    ok(/1 certificate/.test(await batchText(page)) && /Demo High/.test(await batchText(page)), "D: proctor=Proctor B -> 1 certificate (Demo High)");
    await page.click("#mos-comp-f-reset");
    await page.waitForTimeout(400);
    await page.selectOption("#mos-comp-f-school", "Second High");
    await page.waitForTimeout(400);
    let t = await batchText(page);
    ok(/1 certificate/.test(t) && /Second High/.test(t), "D: school=Second High -> 1 certificate");
    // Shared state: Results tab shows the same school filter.
    await page.click('[data-mos-view="results"]');
    await page.waitForTimeout(500);
    const resSchool = await page.evaluate(() => (document.getElementById("mos-f-school") || {}).value || "");
    ok(resSchool === "Second High", "D: Results tab reflects Completion's school filter (got " + resSchool + ")");
    await page.click('[data-mos-view="completion"]');
    await page.waitForTimeout(500);
    await page.click("#mos-comp-f-reset");
    await page.waitForTimeout(400);
    await page.selectOption("#mos-comp-f-pmode", "month");
    await page.waitForTimeout(400);
    await page.selectOption("#mos-comp-f-pval", "2026-10");
    await page.waitForTimeout(400);
    t = await batchText(page);
    ok(/2 certificates/.test(t), "D: period month=October 2026 -> 2 certificates (got " + (t.match(/\d+ certificates/) || ["?"])[0] + ")");
    ok(errs.length === 0, "section D: zero page errors (got " + errs.length + ")");
    await ctx.close();
  }

  // E. SVG signature rasterizes to PNG (was: passed through, then silently
  // dropped by the PPTX engine because SVG is not a placeable picture).
  {
    const { ctx, page, errs } = await boot(browser);
    await page.setInputFiles("#mos-comp-s1", path.join(ROOT, "tests", "comp-sig.svg"));
    await page.waitForTimeout(2000);
    ok(await storedS1IsPng(page), "E: SVG stored as PNG");
    ok(errs.length === 0, "section E: zero page errors");
    await ctx.close();
  }
  await browser.close();

  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL - " + (e && e.stack || e)); process.exit(1); });
