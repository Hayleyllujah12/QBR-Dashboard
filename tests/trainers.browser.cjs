/* Trainers tab browser walkthrough (file://). Synthetic workbooks built in-memory.
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers.browser.cjs */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

function makeTracker() {
  const rows = [
    ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
    [45000, "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""],
    [45000, "SANTOS, Maria", "MS-103A", "Workshop P1", "SPC", 0.54, 0.62, ""]
  ];
  const wb = XLSX.utils.book_new();
  // Used range starts at B3 (regression: write-back must add the sheet offset)
  const ws = XLSX.utils.aoa_to_sheet(rows, { origin: "B3" });
  ws["!ref"] = "B3:I5"; // Excel-style <dimension>: used range, not A1
  XLSX.utils.book_append_sheet(wb, ws, "trial");
  const p = path.join(os.tmpdir(), "tr-fixture-tracker.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
function makeEval() {
  const head = ["ID", "School / Organization", "Knowledge and mastery of the topic",
    "Communication and presentation skills", "Ability to engage and connect with participants",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
    "What did you like best about the session?"];
  const rows = [head];
  const t1 = ["Strongly Agree", "Agree", "Strongly Agree", "Strongly Agree"];
  const nps = [10, 9, 8, 6];
  for (let i = 0; i < 4; i++)
    rows.push([i + 1, "HCCCI", t1[i], "Strongly Agree", "Agree", nps[i], "Great pace"]);
  rows.push([5, "HCCCI", "", "", "", "", ""]); // regression: blank NPS answer is skipped, not a detractor
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const p = path.join(os.tmpdir(), "tr-fixture-eval_01012024_-_AM_-_HCCCI_1-4.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(600);
  // The app keeps #app-body hidden until the main workbook loads; reveal it for tab testing.
  await page.evaluate(() => document.getElementById("app-body").classList.remove("d-none"));

  // 1. Nav button + feature flag registration
  ok(await page.evaluate(() => !!document.querySelector('[data-tab="dash-trainers"]')), "Trainers nav button exists");
  ok(await page.evaluate(() => QBR.features.list().some(f => f.key === "trainers")), "trainers registered in feature flags");

  // 2. Open tab -> empty state with pills
  await page.click('[data-tab="dash-trainers"]');
  await page.waitForSelector("#trainers-body", { state: "attached" });
  await page.waitForTimeout(300);
  ok(await page.evaluate(() => document.querySelectorAll(".tr-pill").length === 2), "Tracker + Eval pills shown");

  // 3. Upload tracker via the pill's file picker
  const trackerPath = makeTracker();
  {
    const [chooser] = await Promise.all([
      page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      page.evaluate(() => { document.querySelector('[data-trpill="tracker"]').click(); })
    ]);
    if (chooser) await chooser.setFiles(trackerPath);
  }
  await page.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr").length === 2, null, { timeout: 8000 })
    .catch(() => {});
  ok(await page.evaluate(() => document.querySelectorAll(".tr-table tbody tr").length) === 2, "2 sessions listed from tracker");
  ok(await page.evaluate(() => document.querySelector(".tr-pill.on .tr-pill-sub").textContent.includes("tr-fixture-tracker")),
    "Tracker pill shows filename");

  // 4. Filter pills
  await page.click('[data-trfilter="scored"]');
  await page.waitForTimeout(300);
  ok(await page.evaluate(() => document.querySelectorAll(".tr-table tbody tr").length) === 0, "Scored filter: 0 rows");
  await page.click('[data-trfilter="all"]');
  await page.waitForTimeout(300);

  // 5. Select session -> detail with upload button
  await page.click('.tr-table tbody tr[data-trsession="0"]');
  await page.waitForTimeout(300);
  ok(await page.evaluate(() => !!document.querySelector(".tr-detail")), "session detail opens");
  ok(await page.evaluate(() => document.querySelector(".tr-detail").textContent.includes("ROE")),
    "detail shows trainer");

  // 6. Upload eval for the session
  const evalPath = makeEval();
  {
    const [chooser] = await Promise.all([
      page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      page.evaluate(() => { document.querySelector('[data-traction="upload-eval"]').click(); })
    ]);
    if (chooser) await chooser.setFiles(evalPath);
  }
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 })
    .catch(() => {});
  // First sight of this layout -> mapper should appear (not saved yet)
  const hasMapper = await page.evaluate(() => !!document.querySelector(".tr-mapper"));
  ok(hasMapper, "column mapper appears for new layout");
  if (hasMapper) {
    await page.click('[data-traction="save-map"]');
    await page.waitForTimeout(400);
  }
  ok(await page.evaluate(() => !!document.querySelector('[data-traction="generate"]')), "Generate button after mapping");

  // 7. Generate -> report with expected numbers
  // NPS: [10,9,8,6] -> promoters 2, passives 1, detractors 1 -> (2-1)/4*100 = 25
  // T-score rows: (5+5+4)/3, (4+5+4)/3, (5+5+4)/3, (5+5+4)/3 -> mean 55/12 = 4.58
  await page.click('[data-traction="generate"]');
  await page.waitForTimeout(500);
  const rep = await page.evaluate(() => document.querySelector(".tr-kpis").textContent.replace(/\s+/g, " "));
  ok(rep.includes("25"), "NPS 25.0 shown (2 promoters, 1 detractor of 4)");
  ok(rep.includes("4.58"), "trainer score 4.58 shown");
  ok(await page.evaluate(() => document.querySelector('.tr-table tbody tr[data-trsession="0"] .tr-status').textContent.includes("Needs verifying")),
    "row marked Needs verifying");

  // 8. Leaderboard
  await page.click('[data-trview="board"]');
  await page.waitForTimeout(400);
  ok(await page.evaluate(() => document.querySelector(".tr-table").textContent.includes("ROE")),
    "leaderboard lists trainer");

  // 9. Write-back: confirm lists the pending row, values land at the sheet offset,
  //    format-safe (untouched zip parts byte-identical), Trainer Score added once.
  await page.click('[data-trview="sessions"]'); await page.waitForTimeout(300);
  await page.click('.tr-table tbody tr[data-trsession="0"]'); await page.waitForTimeout(300);
  const S = require(path.join(ROOT, "js/xlsx-surgical.js"));
  async function doWriteback(tag) {
    let msg = "";
    page.once("dialog", d => { msg = d.message(); d.accept(); });
    const [dl] = await Promise.all([
      page.waitForEvent("download", { timeout: 8000 }).catch(() => null),
      page.click('[data-traction="writeback"]')
    ]);
    const out = dl ? path.join(os.tmpdir(), "tr-out-" + tag + ".xlsx") : null;
    if (dl) await dl.saveAs(out);
    return { msg, out, name: dl ? dl.suggestedFilename() : "" };
  }
  const w1 = await doWriteback(1);
  ok(/Row 4 \(.*ROE.*\): NPS = 25\.0, Trainer Score = 4\.58/.test(w1.msg), "confirm lists the pending row (Excel row 4) with both scores");
  ok(w1.name === "tr-fixture-tracker-scored.xlsx", "write-back downloads <name>-scored.xlsx");
  if (w1.out) {
    const ws = XLSX.read(fs.readFileSync(w1.out)).Sheets.trial;
    ok(ws.I4 && ws.I4.v === 25 && ws.J3 && ws.J3.v === "Trainer Score" && ws.J4 && ws.J4.v === 4.58,
      "NPS → I4, \"Trainer Score\" header J3, score J4 (offset B3 honoured)");
    ok(!ws.I5 || ws.I5.v === "" , "other session row untouched");
    const a = new Uint8Array(fs.readFileSync(trackerPath)), b = new Uint8Array(fs.readFileSync(w1.out));
    const za = S.zipRead(a), zb = S.zipRead(b), diff = [];
    za.entries.forEach(e => { const f = zb.byName.get(e.name); if (!f || f.crc !== e.crc) diff.push(e.name); });
    ok(diff.every(n => /^xl\/(worksheets\/sheet1|workbook)\.xml$/.test(n)) && diff.includes("xl/worksheets/sheet1.xml"),
      "format-safe: only the tracker sheet (+ workbook calc flag) changed (" + diff.join(",") + ")");
  }
  ok(await page.evaluate(() => document.querySelector('.tr-table tbody tr[data-trsession="0"] .tr-status').textContent.includes("Scored")),
    "row marked Scored after write-back");
  // Re-generate + write again: no duplicate "Trainer Score" column
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(400);
  // (a changed value — identical values are a no-op for the format-safe engine)
  await page.evaluate(() => { const c = QBR.trainers._state.tracker.sessions[0].computed; c.nps = 30; QBR.trainers.render(); });
  const w2 = await doWriteback(2);
  if (w2.out) {
    const rows = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(w2.out)).Sheets.trial, { header: 1, defval: "" });
    ok(rows[0].filter(h => h === "Trainer Score").length === 1 && rows[1][7] === 30,
      "second write-back builds on the first: NPS updated, \"Trainer Score\" still exactly once");
  } else ok(false, "second write-back downloaded");

  // 10. Leaderboard keeps NPS 0 above "no NPS" (null-safe sort)
  ok(await page.evaluate(() => {
    const s = QBR.trainers._state.tracker.sessions;
    s[1].nps = 0; s[1].trainerScore = 3; s[0].nps = null; s[0].trainerScore = 4;
    QBR.trainers._state.view = "board"; QBR.trainers.render();
    const first = document.querySelector(".tr-table tbody tr td").textContent;
    return /SANTOS/.test(first);
  }), "leaderboard: NPS 0 ranks above a trainer with no NPS");
  ok(await page.evaluate(() => [...document.querySelectorAll(".tr-table")].every(t => t.getAttribute("data-resizable") === "static")),
    "trainers tables opt out of global column resizing");

  const trErrs = errs.filter(e => /trainer/i.test(e));
  ok(trErrs.length === 0, "zero page errors from trainers code" + (trErrs.length ? ": " + trErrs.join(" | ") : ""));
  if (errs.length) console.log("  (unrelated page errors: " + errs.length + ")");
  // 11. Flag off: trainers:false hides the Training group, render no-ops, no errors
  {
    const ctx = await browser.newContext();
    await ctx.addInitScript(() => { try { localStorage.setItem("qbr-feature-flags", JSON.stringify({ trainers: false })); } catch (e) {} });
    const p2 = await ctx.newPage(); const e2 = [];
    p2.on("pageerror", e => e2.push(String(e.message)));
    await p2.goto(APP, { waitUntil: "domcontentloaded" });
    await p2.waitForTimeout(800);
    ok(await p2.evaluate(() => { const g = document.querySelector('.sb-group[data-feature="trainers"]'); return !!g && getComputedStyle(g).display === "none"; }),
      "flag off: Training group hidden");
    ok(await p2.evaluate(() => { QBR.trainers.render(); return document.getElementById("trainers-body").innerHTML === ""; }),
      "flag off: render is a no-op");
    ok(e2.filter(e => /trainer/i.test(e)).length === 0, "flag off: zero trainers errors");
    await ctx.close();
  }
  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
