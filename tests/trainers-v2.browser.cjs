/* Browser test: (A) combobox trainer/school + response-timestamp date auto-detect,
 * (B) in-place write-back with a mock file handle (no download).
 * Usage: PW=/tmp/pw-run/node_modules node tests/trainers-v2.browser.cjs */
"use strict";
const { chromium } = require(process.env.PW + "/playwright-core");
const fs = require("fs"), os = require("os"), path = require("path");
const XLSX = require("/home/hatch/workspace/qbr-trainers-dev/libs/xlsx.full.min.js");
const APP = "file:///home/hatch/workspace/qbr-trainers-dev/index.html";
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

function makeTracker() {
  const rows = [
    ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
    [45000, "RICO, Reneeson III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""]
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Course Code", "Product", "Session Title", "Estimated Duration"],
    ["MS-102A", "Microsoft", "Teams Workshop", 4]
  ]), "03 Course List");
  const p = path.join(os.tmpdir(), "tr-v2-tracker.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
// Eval: filename says 01012024 but ALL response timestamps are 5/7/2026.
function makeEval() {
  const head = ["ID", "Start time", "Knowledge and mastery of the topic",
    "Communication and presentation skills", "Ability to engage and connect with participants",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
    "What did you like best about the session?"];
  const rows = [head];
  for (let i = 0; i < 6; i++)
    rows.push([i + 1, "5/7/2026 9:0" + i + ":00 AM", "Strongly Agree", "Strongly Agree", "Agree", 10, "Great"]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const p = path.join(os.tmpdir(), "tr-v2-eval_01012024_-_AM_-_HCCCI_1-4.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
async function uploadVia(page, pillSel, filePath) {
  const [chooser] = await Promise.all([
    page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
    page.evaluate((s) => { document.querySelector(s).click(); }, pillSel)
  ]);
  if (chooser) await chooser.setFiles(filePath);
}

(async () => {
  const browser = await chromium.launch({ executablePath: "/opt/meta-chromium/chrome" });
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(500);
  await page.evaluate(() => document.getElementById("app-body").classList.remove("d-none"));
  await page.evaluate(() => { window.showOpenFilePicker = undefined; }); // force input fallback in tests
  await page.click('[data-tab="dash-trainers"]');
  await page.waitForSelector("#trainers-body", { state: "attached" });
  await page.waitForTimeout(300);

  const trk = makeTracker(), evf = makeEval();
  await uploadVia(page, '[data-trpill="tracker"]', trk);
  await page.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr").length === 1, null, { timeout: 8000 });

  /* ---- A: date auto-detect + comboboxes ---- */
  // Upload eval with NO session selected -> "Match this eval file" view
  await uploadVia(page, '[data-trpill="eval"]', evf);
  await page.waitForFunction(() => !!document.querySelector(".tr-mapper, .tr-detail h3"), null, { timeout: 8000 });
  if (await page.evaluate(() => !!document.querySelector(".tr-mapper"))) {
    await page.click('[data-traction="save-map"]');
    await page.waitForTimeout(400);
  }
  await page.waitForFunction(() => !!document.querySelector('[data-traction="new-session"]'), null, { timeout: 8000 });
  await page.click('[data-traction="new-session"]');
  await page.waitForSelector("#trns-date", { timeout: 10000 });

  const dateVal = await page.inputValue("#trns-date");
  ok(dateVal === "2026-05-07", "A1 date from response timestamps, not filename (got " + dateVal + ")");
  const hintTxt = await page.evaluate(() => {
    const el = document.querySelector("#trns-date");
    const sib = el && el.parentElement.querySelector(".tr-hint");
    return sib ? sib.textContent : "";
  });
  ok(hintTxt.indexOf("response timestamps") >= 0, "A2 date source hint shown (got '" + hintTxt.trim() + "')");

  ok(await page.evaluate(() => document.querySelector("#trns-trainer").tagName.toLowerCase()) === "input",
    "A3 trainer is a typeable input");
  ok(await page.evaluate(() => {
    const i = document.querySelector("#trns-trainer");
    return i.getAttribute("list") === "trns-trainers" &&
      Array.from(document.querySelectorAll("#trns-trainers option")).some(o => o.value === "RICO, Reneeson III");
  }), "A4 trainer datalist has roster options");
  ok(await page.evaluate(() => {
    const i = document.querySelector("#trns-school");
    return i.tagName.toLowerCase() === "input" && i.getAttribute("list") === "trns-schools" &&
      document.querySelectorAll("#trns-schools option").length > 0;
  }), "A5 school is a typeable input with datalist options");

  await page.fill("#trns-trainer", "RICO, Reneeson III");
  await page.fill("#trns-school", "Typed Academy of Test");
  await page.fill("#trns-code", "MS-102A");
  await page.click('[data-traction="save-new-session"]');
  await page.waitForFunction(() => window.QBR.trainers._state.tracker.addedRows.length === 1, null, { timeout: 10000 });
  const added = await page.evaluate(() => { const t = window.QBR.trainers._state.tracker; return t.sessions[t.sessions.length - 1]; });
  ok(added.client === "Typed Academy of Test", "A6 typed school saved (got '" + added.client + "')");
  ok(added.dateISO === "2026-05-07", "A7 saved date = timestamp date (got '" + added.dateISO + "')");
  /* ---- B: in-place write-back with mock handle ---- */
  // New session saved -> selected with eval attached; generate scores for it.
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 });
  if (await page.evaluate(() => !!document.querySelector(".tr-mapper"))) {
    await page.click('[data-traction="save-map"]');
    await page.waitForTimeout(400);
  }
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"]'), null, { timeout: 8000 });
  await page.click('[data-traction="generate"]');
  await page.waitForFunction(() => !!document.querySelector('[data-traction="writeback"]'), null, { timeout: 8000 });
  await page.evaluate(() => {
    window.__wroteBytes = 0;
    window.__downloaded = [];
    // Mock FileSystemFileHandle
    window.QBR.trainers._state.tracker.handle = {
      createWritable: async () => ({
        write: async (b) => { window.__wroteBytes = b.byteLength || b.length || 0; },
        close: async () => {}
      })
    };
    // Record downloads instead of navigating
    const orig = URL.createObjectURL;
    URL.createObjectURL = function (blob) { window.__downloaded.push(blob.size); return "blob:mock"; };
    window.confirm = () => true;
  });
  await page.click('[data-traction="writeback"]');
  await page.waitForFunction(() => window.__wroteBytes > 0 || window.__downloaded.length > 0, null, { timeout: 15000 });
  const wrote = await page.evaluate(() => window.__wroteBytes);
  const dls = await page.evaluate(() => window.__downloaded.length);
  ok(wrote > 0, "B1 in-place write called with " + wrote + " bytes");
  ok(dls === 0, "B2 no -scored download when handle exists (downloads: " + dls + ")");
  const st = await page.evaluate(() => window.QBR.trainers._state.tracker.sessions.some(s => s.status === "scored"));
  ok(st, "B3 session marked scored after in-place write-back");

  ok(errs.length === 0, "zero page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL", e); process.exit(2); });
