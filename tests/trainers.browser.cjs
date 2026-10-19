/* Trainers tab browser walkthrough (file://). Synthetic workbooks built in-memory.
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers.browser.cjs */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const XLSX = require("/home/hatch/workspace/qbr-trainers-dev/libs/xlsx.full.min.js");
const APP = "file:///home/hatch/workspace/qbr-trainers-dev/index.html";
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

function makeTracker() {
  const rows = [
    ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
    [45000, "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""],
    [45000, "SANTOS, Maria", "MS-103A", "Workshop P1", "SPC", 0.54, 0.62, ""]
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
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
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const p = path.join(os.tmpdir(), "tr-fixture-eval_01012024_-_AM_-_HCCCI_1-4.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}

(async () => {
  const browser = await chromium.launch({ executablePath: "/opt/meta-chromium/chrome" });
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

  const trErrs = errs.filter(e => /trainer/i.test(e));
  ok(trErrs.length === 0, "zero page errors from trainers code" + (trErrs.length ? ": " + trErrs.join(" | ") : ""));
  if (errs.length) console.log("  (unrelated page errors: " + errs.length + ")");
  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
