/* v1.39.0 — Trainers: response counts are written to / read from the tracker, and an eval already written
 * to a row before "Eval Source" existed is recognised (no second session for the same eval).
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers-responses.browser.cjs   (synthetic data only)
 * A: legacy row (same date, same NPS + trainer score, no Eval Source) -> file points at that row, warns,
 *    "Add as a new session" asks first; Replace + generate + write records Eval Source AND Responses there
 * B: reopening the written tracker shows the response count (was "—" for rows scored earlier)
 * C: row encoded with a fingerprint but no Responses -> count-only write, scores untouched */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const TP = require(path.join(ROOT, "js/trainers-parse.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), "trresp-"));

// eval: 12 responses, Oct 16 2025, ASVF
const HEAD = ["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic",
  "Communication and presentation skills", "Ability to engage and connect with participants",
  "Overall, how satisfied were you with the training?",
  "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague", "Why did you give this rating?"];
const ER = [HEAD];
const npsV = [10, 10, 9, 10, 8, 10, 9, 6, 10, 10, 7, 10];
for (let i = 0; i < 12; i++) ER.push([i + 1, 45946.4 + i / 1440, "ASVF", "Strongly Agree", i % 3 ? "Agree" : "Strongly Agree", "Strongly Agree", "Very satisfied", npsV[i], "ok"]);
const det = TP.detectLayout(ER[0]);
const rep = TP.analyzeEval(ER, det.roles);
const NPS = rep.nps.nps, TS = rep.trainer.score, FP = TP.evalFingerprint(ER);
const EVAL = path.join(DIR, "Training Feedback Form - Oct. 16, 2025 [ASVF] - Micro.xlsx");
{ const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ER), "Sheet1"); fs.writeFileSync(EVAL, XLSX.write(wb, { bookType: "xlsx", type: "buffer" })); }

function tracker(file, extraHead, extraCells) {
  const rows = [["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score", "Trainer Score"].concat(extraHead),
    [45946, "DOEÑA, Alex", "MS-101A", "Microsoft 365 Education Onboarding", "Academia de San Valle Fuerte", 0.375, 0.5, NPS, TS].concat(extraCells),
    [45950, "DOEÑA, Alex", "MS-101A", "Microsoft 365 Education Onboarding", "Saint Garnet Academy", 0.375, 0.5, "", ""].concat(extraHead.map(() => ""))];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
  const p = path.join(DIR, file); fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" })); return p;
}
const LEGACY = tracker("legacy-tracker.xlsx", [], []);
const ENCODED = tracker("encoded-tracker.xlsx", ["Eval Source"], [TP.evalSourceText(path.basename(EVAL), FP)]);

async function upload(page, sel, f) {
  const [c] = await Promise.all([page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null), page.click(sel)]);
  if (c) await c.setFiles(f);
  await page.waitForTimeout(600);
}
async function mapIfAsked(page) {
  if (await page.$('[data-traction="q-map"]')) { await page.click('[data-traction="q-map"]'); await page.waitForTimeout(200); }
  if (await page.$(".tr-mapper")) { await page.click('[data-traction="save-map"]'); await page.waitForTimeout(300); }
}
async function fresh(browser) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page._errs = []; page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(400);
  await page.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });
  await page.click('[data-tab="dash-trainers"]'); await page.waitForTimeout(200);
  return page;
}
async function write(page) {
  let msg = ""; page.once("dialog", d => { msg = d.message(); d.accept(); });
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  if (!dl) return { msg, rows: null };
  const out = path.join(DIR, "out-" + Date.now() + ".xlsx"); await dl.saveAs(out);
  return { msg, out, rows: XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(out)).Sheets.trial, { header: 1, defval: "" }) };
}
const listCounts = page => page.$$eval("#tr-list tbody tr[data-trsession]", trs => trs.map(tr => tr.children[4].textContent.trim()));

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});

  /* ---- A ---- */
  let page = await fresh(browser);
  await upload(page, '[data-trpill="tracker"]', LEGACY);
  ok(JSON.stringify(await listCounts(page)) === '["—","—"]', "legacy tracker: no Responses column -> \"—\"");
  ok(await page.$eval("#tr-list tbody tr[data-trsession] td:nth-child(5)", td => /not in the tracker/.test(td.title)), "\"—\" on a scored row explains why (tooltip)");
  await upload(page, '[data-trpill="eval"]', EVAL);
  await mapIfAsked(page);
  const A = await page.evaluate(() => { const q = QBR.trainers._state.queue[0]; return { t: q.target, twin: q.twin, note: q.note }; });
  ok(A.t === 0 && A.twin === 0, "same eval again: points at the row that already has its exact scores");
  ok(/already has these exact scores/.test(A.note), "…and says so (no Eval Source recorded there)");
  let asked = ""; page.once("dialog", d => { asked = d.message(); d.dismiss(); });
  await page.selectOption("[data-trq-target]", "new"); await page.waitForTimeout(300);
  ok(/already has this eval's scores/.test(asked) && await page.evaluate(() => !QBR.trainers._state.newFor), "\"Add as a new session\" asks first; Cancel keeps the row");
  await page.check("[data-trq-replace]"); await page.waitForTimeout(200);
  await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(400);
  let w = await write(page);
  const hd = w.rows ? w.rows[0] : [], ri = hd.indexOf("Responses"), si = hd.indexOf("Eval Source");
  ok(w.rows && ri > 0 && w.rows[1][ri] === 12, "write-back adds a Responses column with the count (12)");
  ok(w.rows && si > 0 && /fp:/.test(w.rows[1][si]) && w.rows.length === 3, "…records Eval Source on the SAME row — no duplicate session row");
  ok(!page._errs.length, "no page errors (A) " + page._errs.join(" | "));
  await page.close();

  /* ---- B ---- */
  page = await fresh(browser);
  await upload(page, '[data-trpill="tracker"]', w.out);
  ok((await listCounts(page))[0] === "12", "reopened tracker shows the stored response count");
  await page.close();

  /* ---- C ---- */
  page = await fresh(browser);
  await upload(page, '[data-trpill="tracker"]', ENCODED);
  await upload(page, '[data-trpill="eval"]', EVAL);
  await mapIfAsked(page);
  ok(await page.evaluate(() => QBR.trainers._state.queue[0].encodedAt === 0), "fingerprinted row: eval shows as already encoded");
  ok(/1 response count/.test(await page.textContent(".tr-pending")), "pending bar offers to record the missing response count");
  w = await write(page);
  const h2 = w.rows ? w.rows[0] : [], r2 = h2.indexOf("Responses"), n2 = h2.indexOf("Net Promoter Score");
  ok(/Responses = 12 \(count only/.test(w.msg), "confirm lists the count-only change");
  ok(w.rows && r2 > 0 && w.rows[1][r2] === 12 && w.rows[1][n2] === NPS && w.rows.length === 3, "count written; NPS untouched; no new row");
  ok(!page._errs.length, "no page errors (C) " + page._errs.join(" | "));
  await page.close();

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
