// MOS completion certificates — PPTX export through the real UI (headless Chromium).
// Boots the app, loads failed-twice fixtures, opens the Completion view,
// saves signatories + a signature image, clicks Download PPTX, and verifies
// the downloaded file: slide count/order, Expert badge removal, filled runs,
// signature picture. Dummy data only.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-mos-completion-pptx.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "libs", "jszip.min.js"));
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

// Invented fixture: Demo High only. A: Excel fail x2. B: Word fail x2.
// C: Excel Expert fail x2. D: PowerPoint pass only -> excluded.
function attempts() {
  const mk = (ref, school, exam, score, date, extra) => Object.assign({ ref, school, exam, score, name: "Sample " + ref, date }, extra || {});
  return [
    mk("A1", "Demo High", "Excel", 500, "2026-10-01", { name: "Maria Santos" }),
    mk("A2", "Demo High", "Excel", 600, "2026-10-20", { name: "Maria Santos" }),
    mk("B1", "Demo High", "Word", 400, "2026-10-02", { name: "Jose Reyes" }),
    mk("B2", "Demo High", "Word", 650, "2026-10-21", { name: "Jose Reyes" }),
    mk("C1", "Demo High", "Excel Expert", 500, "2026-10-03", { name: "Ana Cruz" }),
    mk("C2", "Demo High", "Excel Expert", 600, "2026-10-22", { name: "Ana Cruz" }),
    mk("D1", "Demo High", "PowerPoint", 800, "2026-10-04", { name: "Jane Roe" }),
  ];
}
async function boot(browser) {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message)));
  page.on("dialog", d => d.dismiss()); // log-create confirm, if any: dismissed, generation proceeds
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
function slideRuns(xml, yIn) {
  // Runs of the text box whose frame top is near yIn inches (EMU tolerance).
  const blocks = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)].map(m => m[1]);
  for (const b of blocks) {
    const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(b);
    if (off && Math.abs(+off[2] / 914400 - yIn) < 0.05) {
      return [...b.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1]);
    }
  }
  return null;
}
function hasBadge(xml) {
  return [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)].some(m => {
    const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(m[1]);
    return off && Math.abs(+off[1] / 914400 - 1.242) < 0.05 && Math.abs(+off[2] / 914400 - 1.155) < 0.05 && /<a:blip/.test(m[1]);
  });
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : { args: ["--no-sandbox"] });
  const { ctx, page, errs } = await boot(browser);
  ok(await page.evaluate(() => !!document.getElementById("mos-comp-pptx") && !document.getElementById("mos-comp-pptx").disabled),
    "Download PPTX button present and enabled");
  await page.fill("#mos-comp-date", "2026-03-27");
  await page.dispatchEvent("#mos-comp-date", "change");
  await page.fill("#mos-comp-n1", "DUMMY ONE");
  await page.fill("#mos-comp-t1", "Dummy Title One");
  await page.setInputFiles("#mos-comp-s1", path.join(ROOT, "tests", "comp-sig.png"));
  await page.click("#mos-comp-sigsave");
  await page.waitForTimeout(300);
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 30000 }),
    page.click("#mos-comp-pptx"),
  ]);
  const dlPath = await download.path();
  ok(!!dlPath, "PPTX download fired");
  const bytes = fs.readFileSync(dlPath);
  ok(bytes.length > 100000, "downloaded PPTX non-trivial (" + bytes.length + " bytes)");
  const zip = await JSZip.loadAsync(bytes);
  const slides = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort();
  ok(slides.length === 3, "downloaded PPTX has 3 slides (got " + slides.length + ")");
  const xmls = [];
  for (const n of slides) xmls.push(await zip.file(n).async("string"));
  // Batch order: MOS_FAMS order Excel, Excel Expert, Word.
  const names = xmls.map(x => (slideRuns(x, 3.867) || [])[0]);
  ok(names.join("|") === "SANTOS, MARIA|CRUZ, ANA|REYES, JOSE",
    "slide order Excel, Expert, Word (got " + names.join("|") + ")");
  const exam2 = slideRuns(xmls[1], 4.965);
  ok(exam2 && exam2[1].includes("Excel Expert (Office 365 Apps)"), "Expert slide title");
  ok(!hasBadge(xmls[1]), "Expert slide: NO badge shape");
  ok(hasBadge(xmls[0]) && hasBadge(xmls[2]), "Associate slides: badges present");
  const sig1 = slideRuns(xmls[0], 6.917); // sig row at the fixed-template position (6.917in)
  ok(sig1 && sig1[0] === "DUMMY ONE", "signatory name written (got " + (sig1 && sig1[0]) + ")");
  ok(!!zip.file("ppt/media/sig1.png") || !!zip.file("ppt/media/sig1.jpg"), "signature media part present (png or jpg after downscale)");
  ok(xmls[0].includes("<p:pic>"), "signature <p:pic> on slide 1");
  const msg = await page.evaluate(() => (document.getElementById("mos-comp-msg") || {}).textContent || "");
  ok(/PPTX downloaded/i.test(msg), "UI confirms PPTX download (msg: " + msg + ")");
  ok(errs.length === 0, "zero page errors (got " + errs.length + (errs[0] ? ": " + errs[0] : "") + ")");
  await ctx.close();
  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL - " + (e && e.stack || e)); process.exit(1); });
