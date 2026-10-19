/* Trainers tab browser walkthrough (file://). Synthetic workbooks built in-memory.
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers.browser.cjs */
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
    [45000, "RICO, Reneeson III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""],
    [45000, "SANTOS, Maria", "MS-103A", "Workshop P1", "SPC", 0.54, 0.62, ""]
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Course Code", "Product", "Session Title", "Estimated Duration"],
    ["MS-102A", "Microsoft", "Teams Workshop", 4],
    ["MS-103A", "Microsoft", "Workshop P1", 3]
  ]), "03 Course List");
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
  await page.evaluate(() => { window.showOpenFilePicker = undefined; }); // force input fallback in tests

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
  ok(await page.evaluate(() => document.querySelector(".tr-detail").textContent.includes("RICO")),
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

  // 8. Flow B regression: eval uploaded with NO session selected and a filename
  //    that matches nothing -> clicking a list row must attach the eval.
  //    (Real-world case: filename date not in the tracker.)
  {
    const b2 = await chromium.launch();
    const p2 = await b2.newPage();
    p2.on("pageerror", e => errs.push("flowB:" + String(e.message).slice(0, 120)));
    await p2.goto(APP, { waitUntil: "domcontentloaded" });
    await p2.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
    await p2.evaluate(() => document.getElementById("app-body").classList.remove("d-none"));
    await p2.evaluate(() => { window.showOpenFilePicker = undefined; });
    await p2.click('[data-tab="dash-trainers"]');
    await p2.waitForTimeout(300);
    let ch = (await Promise.all([
      p2.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      p2.evaluate(() => { document.querySelector('[data-trpill="tracker"]').click(); })
    ]))[0];
    if (ch) await ch.setFiles(trackerPath);
    await p2.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr").length === 2, null, { timeout: 8000 });
    // Upload eval WITHOUT selecting a session; filename date matches no session.
    ch = (await Promise.all([
      p2.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      p2.evaluate(() => { document.querySelector('[data-trpill="eval"]').click(); })
    ]))[0];
    if (ch) await ch.setFiles(evalPath);
    await p2.waitForFunction(() => document.querySelector(".tr-detail").textContent.includes("Match this eval file"), null, { timeout: 8000 });
    ok(await p2.evaluate(() => document.querySelector(".tr-detail").textContent.includes("No session matched")),
      "match panel shown when filename matches nothing");
    // Click a session row -> eval must attach (mapper or generate appears).
    await p2.click('.tr-table tbody tr[data-trsession="1"]');
    await p2.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 });
    ok(true, "clicking a row attaches the waiting eval file");
    await b2.close();
  }

  // 9. Leaderboard
  await page.click('[data-trview="board"]');
  await page.waitForTimeout(400);
  ok(await page.evaluate(() => document.querySelector(".tr-table").textContent.includes("RICO")),
    "leaderboard lists trainer");

  // 10. Add-as-new-session flow (eval with no tracker match)
  {
    const b3 = await chromium.launch();
    const p3 = await b3.newPage();
    p3.on("pageerror", e => errs.push("flowC:" + String(e.message).slice(0, 120)));
    await p3.goto(APP, { waitUntil: "domcontentloaded" });
    await p3.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
    await p3.evaluate(() => document.getElementById("app-body").classList.remove("d-none"));
    await p3.evaluate(() => { window.showOpenFilePicker = undefined; });
    await p3.click('[data-tab="dash-trainers"]');
    await p3.waitForTimeout(300);
    let ch3 = (await Promise.all([
      p3.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      p3.evaluate(() => { document.querySelector('[data-trpill="tracker"]').click(); })
    ]))[0];
    if (ch3) await ch3.setFiles(trackerPath);
    await p3.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr").length === 2, null, { timeout: 8000 });
    ch3 = (await Promise.all([
      p3.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null),
      p3.evaluate(() => { document.querySelector('[data-trpill="eval"]').click(); })
    ]))[0];
    if (ch3) await ch3.setFiles(evalPath);
    await p3.waitForFunction(() => !!document.querySelector('[data-traction="new-session"]'), null, { timeout: 8000 })
      .catch(() => {});
    ok(await p3.evaluate(() => !!document.querySelector('[data-traction="new-session"]')),
      "Add as new session button shown for unmatched eval");
    // Open the form; school should be pre-filled from the filename acronym.
    await p3.click('[data-traction="new-session"]');
    await p3.waitForTimeout(300);
    const schoolVal = await p3.evaluate(() => document.querySelector("#trns-school").value);
    ok(schoolVal === "HCCCI", "school pre-filled from filename (" + schoolVal + ")");
    ok(await p3.evaluate(() => document.querySelector("#trns-school").tagName === "SELECT"),
      "school is a picker, not free text");
    // Two-way course fill: code -> title.
    await p3.evaluate(() => {
      const codeEl = document.querySelector("#trns-code");
      codeEl.value = "MS-102A";
      codeEl.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await p3.waitForTimeout(200);
    ok(await p3.evaluate(() => document.querySelector("#trns-title").value) === "Teams Workshop",
      "picking code auto-fills the session title");
    ok(await p3.evaluate(() => document.querySelector("#trns-codehint").textContent.includes("Microsoft")),
      "product hint shown for the code");
    // Two-way course fill: title -> code (clear both first).
    await p3.evaluate(() => {
      const codeEl = document.querySelector("#trns-code"), titleEl = document.querySelector("#trns-title");
      codeEl.value = ""; titleEl.value = "";
      titleEl.dispatchEvent(new Event("input", { bubbles: true }));
      titleEl.value = "Workshop P1";
      titleEl.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await p3.waitForTimeout(200);
    ok(await p3.evaluate(() => document.querySelector("#trns-code").value) === "MS-103A",
      "picking title auto-fills the course code");
    // Fill required fields and save.
    await p3.evaluate(() => {
      document.querySelector("#trns-date").value = "2024-01-01";
      document.querySelector("#trns-trainer").selectedIndex = 1;
    });
    await p3.click('[data-traction="save-new-session"]');
    await p3.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr").length === 3, null, { timeout: 8000 })
      .catch(() => {});
    ok(await p3.evaluate(() => document.querySelectorAll(".tr-table tbody tr").length) === 3,
      "new session appears in the list");
    ok(await p3.evaluate(() => document.querySelector(".tr-detail").textContent.includes("MS-103A")),
      "new session selected and eval attached");
    await b3.close();
  }

  // 11. Summary view renders KPIs + breakdowns
  await page.click('[data-trview="summary"]');
  await page.waitForTimeout(400);
  ok(await page.evaluate(() => !!document.querySelector(".tr-report")), "summary report renders");
  const kpiText = await page.evaluate(() => document.querySelector(".tr-report").textContent.replace(/\\s+/g, " "));
  ok(kpiText.includes("Sessions") && kpiText.includes("Training hours") && kpiText.includes("Respondents"),
    "summary KPI cards present");
  ok(kpiText.includes("Sessions by school") && kpiText.includes("Sessions by course"),
    "summary breakdowns present");
  ok(await page.evaluate(() => !!document.querySelector('[data-traction="print-report"]')),
    "print report button present");

  const trErrs = errs.filter(e => /trainer/i.test(e));
  ok(trErrs.length === 0, "zero page errors from trainers code" + (trErrs.length ? ": " + trErrs.join(" | ") : ""));
  if (errs.length) console.log("  (unrelated page errors: " + errs.length + ")");
  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
