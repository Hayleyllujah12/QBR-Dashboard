// MOS completion certificates — signature size/position controls (headless Chromium).
// Sliders + live preview in the signatories editor, saved geometry honored by
// the HTML cert and the PPTX fit box. Dummy data only.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-mos-completion-siggeom.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs");
const path = require("path");
const JSZip = require(path.join(__dirname, "..", "libs", "jszip.min.js"));
global.JSZip = JSZip;
const eng = require(path.join(__dirname, "..", "js", "mos-completion-pptx.js"));
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

function attempts() {
  const mk = (ref, school, exam, score, date, month, proctor, name) =>
    ({ ref, school, exam, score, date, month, proctor, name });
  return [
    mk("A1", "Demo High", "Excel", 500, "2026-10-01", "2026-10", "Proctor A", "Maria Santos"),
    mk("A2", "Demo High", "Excel", 600, "2026-10-20", "2026-10", "Proctor A", "Maria Santos"),
    mk("B1", "Demo High", "Word", 400, "2026-10-02", "2026-10", "Proctor B", "Jose Reyes"),
    mk("B2", "Demo High", "Word", 650, "2026-10-21", "2026-10", "Proctor B", "Jose Reyes"),
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
async function seedSig(page) {
  const b64 = fs.readFileSync(path.join(ROOT, "tests", "comp-sig.png")).toString("base64");
  await page.evaluate(b => {
    localStorage.setItem("qbr-mos-completion-sig-v1", JSON.stringify({
      n1: "DUMMY ONE", t1: "Dummy Title One", n2: "", t2: "Education Technology Director",
      org: "Dummy Org Inc.", s1: "data:image/png;base64," + b, s2: "",
      g1: { size: 100, dx: 0, dy: 0 }, g2: { size: 100, dx: 0, dy: 0 },
    }));
  }, b64);
  await page.click('[data-mos-view="results"]');
  await page.waitForTimeout(300);
  await page.click('[data-mos-view="completion"]');
  await page.waitForTimeout(500);
}
async function setSlider(page, id, val) {
  await page.evaluate(([i, v]) => {
    const el = document.getElementById(i);
    el.value = String(v);
    el.dispatchEvent(new Event("input", { bubbles: true }));
  }, [id, val]);
}
async function prevStyle(page) {
  return page.evaluate(() => {
    const img = document.querySelector("#mos-comp-g1-prev img");
    return img ? img.getAttribute("style") || "" : "";
  });
}
async function prevWrapStyle(page) {
  // The browser serializes styles set via the style property with spaces
  // ("translate(0.1in, -0.05in)"); strip them so checks match the set values.
  return page.evaluate(() => {
    const pv = document.getElementById("mos-comp-g1-prev");
    return pv ? (pv.getAttribute("style") || "").replace(/\s+/g, "") : "";
  });
}
function nameBoxOff(xml, want) {
  for (const b of eng.topSpBlocks(xml)) {
    if (eng.runTexts(b.text)[0] === want) return eng.offOf(b.text);
  }
  return null;
}
async function sliderVals(page) {
  return page.evaluate(() => ["mos-comp-g1-size", "mos-comp-g1-dx", "mos-comp-g1-dy"]
    .map(id => document.getElementById(id).value).join(","));
}
function picFrame(xml) {
  const m = /<p:pic>.*?<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
  return m ? m.slice(1).map(Number) : null;
}
async function downloadPptx(page) {
  const [dl] = await Promise.all([
    page.waitForEvent("download", { timeout: 30000 }),
    page.click("#mos-comp-pptx"),
  ]);
  const p = await dl.path();
  const zip = await JSZip.loadAsync(fs.readFileSync(p));
  return zip.file("ppt/slides/slide1.xml").async("string");
}

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME, args: ["--no-sandbox"] });
  const { ctx, page, errs } = await boot(browser);
  await seedSig(page);

  // Controls + preview present.
  for (const id of ["mos-comp-g1-size", "mos-comp-g1-dx", "mos-comp-g1-dy", "mos-comp-g1-reset", "mos-comp-g1-prev",
    "mos-comp-g2-size", "mos-comp-g2-dx", "mos-comp-g2-dy", "mos-comp-g2-reset", "mos-comp-g2-prev"]) {
    ok(await page.evaluate(i => !!document.getElementById(i), id), "control present: " + id);
  }
  ok((await prevStyle(page)).includes("0.9in"), "preview default 0.9in (got " + await prevStyle(page) + ")");

  // Live preview updates with NO save.
  await setSlider(page, "mos-comp-g1-size", 150);
  ok((await prevStyle(page)).includes("1.35in"), "preview live: 150% -> 1.35in");
  ok(await page.evaluate(() => document.getElementById("mos-comp-g1-size-v").textContent) === "150%", "size readout 150%");
  await setSlider(page, "mos-comp-g1-dx", 0.1);
  await setSlider(page, "mos-comp-g1-dy", -0.05);
  const st1 = await prevStyle(page);
  ok(st1.includes("translate(0.1in,-0.05in)"), "preview live: shift on img (got " + st1 + ")");
  ok(!(await prevWrapStyle(page)).includes("translate"), "preview live: wrapper carries no shift (got " + await prevWrapStyle(page) + ")");

  // Reset restores defaults.
  await page.click("#mos-comp-g1-reset");
  ok((await sliderVals(page)) === "100,0,0", "reset restores 100/0/0 (got " + await sliderVals(page) + ")");
  ok((await prevStyle(page)).includes("0.9in"), "preview back to 0.9in after reset");
  ok((await prevStyle(page)).includes("translate(0in,0in)"), "img shift reset to 0");

  // Save -> view switch keeps geometry (storage round-trip through render).
  await setSlider(page, "mos-comp-g1-size", 150);
  await setSlider(page, "mos-comp-g1-dx", 0.1);
  await setSlider(page, "mos-comp-g1-dy", -0.05);
  await page.click("#mos-comp-sigsave");
  await page.waitForTimeout(500);
  ok((await sliderVals(page)) === "150,0.1,-0.05", "sliders persist after Save (got " + await sliderVals(page) + ")");
  await page.click('[data-mos-view="results"]');
  await page.waitForTimeout(300);
  await page.click('[data-mos-view="completion"]');
  await page.waitForTimeout(500);
  ok((await sliderVals(page)) === "150,0.1,-0.05", "sliders persist across view switch");

  // HTML cert preview honors geometry.
  await page.click("#mos-comp-preview-btn");
  await page.waitForTimeout(500);
  const certStyle = await page.evaluate(() => {
    const img = document.querySelector(".mos-comp-cert .mos-comp-sig img");
    const blk = document.querySelector(".mos-comp-cert .mos-comp-sig");
    return (img ? img.getAttribute("style") || "" : "") + " || " + (blk ? blk.getAttribute("style") || "" : "");
  });
  ok(certStyle.includes("1.35in") && certStyle.includes("translate(0.1in,-0.05in)") && !certStyle.split("||")[1].includes("translate"),
    "cert preview: img honors geometry, block unmoved (got " + certStyle + ")");

  // PPTX: scaled frame vs default frame (same image, same slot).
  const xmlScaled = await downloadPptx(page);
  const fS = picFrame(xmlScaled);
  ok(!!fS, "pptx scaled: pic frame found");
  await page.click("#mos-comp-g1-reset");
  await page.click("#mos-comp-sigsave");
  await page.waitForTimeout(500);
  const xmlDef = await downloadPptx(page);
  const fD = picFrame(xmlDef);
  ok(!!fD, "pptx default: pic frame found");
  if (fS && fD) {
    ok(Math.abs(fS[2] - fD[2] * 1.5) <= 2 && Math.abs(fS[3] - fD[3] * 1.5) <= 2,
      "pptx: 150% scales ext 1.5x (scaled " + fS[2] + "x" + fS[3] + " vs default " + fD[2] + "x" + fD[3] + ")");
    ok(Math.abs((fS[0] + fS[2] / 2) - (fD[0] + fD[2] / 2) - 91440) <= 2, "pptx: dx +0.1in shifts center");
    ok(Math.abs((fS[1] + fS[3]) - (fD[1] + fD[3]) + 45720) <= 2, "pptx: dy -0.05in shifts pic bottom (bottom-anchored)");
    const nbS = nameBoxOff(xmlScaled, "DUMMY ONE"), nbD = nameBoxOff(xmlDef, "DUMMY ONE");
    ok(!!nbS && !!nbD, "pptx: name boxes found");
    if (nbS && nbD) {
      ok(nbD.join(",") === "1699462,6324600", "pptx: default name box = fixed-template position (got " + nbD.join(",") + ")");
      ok(nbS.join(",") === "1699462,6324600", "pptx: shifted geometry leaves name box UNMOVED (got " + nbS.join(",") + ")");
    }
  }

  ok(errs.length === 0, "zero page errors" + (errs.length ? " (" + errs.join(" | ") + ")" : ""));
  await ctx.close();
  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL - " + (e && e.stack || e)); process.exit(1); });
