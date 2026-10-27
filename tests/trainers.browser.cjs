/* Trainers tab browser walkthrough (file://). Synthetic workbooks built in-memory — no real files.
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers.browser.cjs
 * Covers: nav + flag, tracker/eval loading, mapper, scoring, themed KPI tiles, session report,
 * format-safe write-back at a sheet offset, repeat write-back, leaderboard, summary, flag off. */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const SG = require(path.join(ROOT, "js/xlsx-surgical.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

function makeTracker() {
  const rows = [
    ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
    [45000, "ROE, Sample III", "MS-102A", "Teams Workshop", "MRCWI", 0.333, 0.5, ""],
    [45000, "SMITH, Pat", "MS-103A", "Workshop P1", "SPC", 0.54, 0.62, ""]
  ];
  const wb = XLSX.utils.book_new();
  // Used range starts at B3 (regression: write-back must add the sheet offset)
  const ws = XLSX.utils.aoa_to_sheet(rows, { origin: "B3" });
  ws["!ref"] = "B3:I5"; // Excel-style <dimension>: used range, not A1
  XLSX.utils.book_append_sheet(wb, ws, "trial");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Course Code", "Product", "Session Title", "Estimated Duration"],
    ["MS-102A", "Microsoft", "Teams Workshop", 4], ["MS-103A", "Microsoft", "Workshop P1", 3]
  ]), "03 Course List");
  const p = path.join(os.tmpdir(), "tr-fixture-tracker.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
function makeEval() {
  const head = ["ID", "School / Organization", "Knowledge and mastery of the topic",
    "Communication and presentation skills", "Ability to engage and connect with participants",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
    "What did you like best about the session?", "What suggestions do you have to improve?"];
  const rows = [head];
  const t1 = ["Strongly Agree", "Agree", "Strongly Agree", "Strongly Agree"];
  const nps = [10, 9, 8, 6];
  for (let i = 0; i < 4; i++)
    rows.push([i + 1, "MRCWI", t1[i], "Strongly Agree", "Agree", nps[i], "Great pace <b>x</b>", i ? "More practice" : "n/a"]);
  rows.push([5, "MRCWI", "", "", "", "", "", ""]); // regression: blank NPS answer is skipped, not a detractor
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const p = path.join(os.tmpdir(), "tr-fixture-eval_01012024_-_AM_-_MRCWI_1-4.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
async function upload(page, sel, file) {
  const [chooser] = await Promise.all([page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null), page.click(sel)]);
  if (chooser) await chooser.setFiles(file);
}
const kpiVal = (page, scope, label) => page.evaluate(([sc, l]) => {
  const k = [...document.querySelectorAll(sc + " .kpi")].find(x => x.querySelector(".kpi-lbl").textContent.trim() === l);
  return k ? k.querySelector(".kpi-val").textContent.trim() : null;
}, [scope, label]);

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(600);
  await page.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });

  // 1. Nav button + feature flag registration
  ok(await page.evaluate(() => !!document.querySelector('[data-tab="dash-trainers"]')), "Trainers nav button exists");
  ok(await page.evaluate(() => QBR.features.list().some(f => f.key === "trainers")), "trainers registered in feature flags");
  ok(await page.evaluate(() => !document.querySelector('iframe, script[src*="cdnjs"], script[src*="//"]')), "no iframe and no CDN scripts (offline)");

  // 2. Open tab -> themed empty state
  await page.click('[data-tab="dash-trainers"]');
  await page.waitForTimeout(300);
  ok(await page.evaluate(() => document.querySelectorAll("#trainers-body [data-trpill]").length >= 2 && !!document.querySelector("#trainers-body .dash-h")),
    "header + Open tracker / Add eval buttons shown");
  ok(await page.evaluate(() => [...document.querySelectorAll("#trainers-body [data-trpill]")].every(b => b.classList.contains("btn"))), "buttons use the app's Bootstrap .btn style");

  // 3. Upload tracker
  const trackerPath = makeTracker();
  await upload(page, '[data-trpill="tracker"]', trackerPath);
  await page.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr[data-trsession]").length === 2, null, { timeout: 8000 }).catch(() => {});
  ok(await page.evaluate(() => document.querySelectorAll(".tr-table tbody tr[data-trsession]").length) === 2, "2 sessions listed from tracker");
  ok(await page.evaluate(() => document.querySelector('[data-trfile="tracker"]').textContent.includes("tr-fixture-tracker")), "tracker filename shown");
  ok(await kpiVal(page, "#trainers-body", "Sessions") === "2", "KPI tile: label 'Sessions' above value 2 (not swapped)");

  // 4. Filters
  await page.click('[data-trfilter="scored"]'); await page.waitForTimeout(200);
  ok(await page.evaluate(() => document.querySelectorAll(".tr-table tbody tr[data-trsession]").length) === 0, "Scored filter: 0 rows");
  await page.click('[data-trfilter="all"]'); await page.waitForTimeout(200);
  await page.fill("#tr-search", "smith"); await page.waitForTimeout(200);
  ok(await page.evaluate(() => document.querySelectorAll(".tr-table tbody tr[data-trsession]").length === 1 && document.activeElement.id === "tr-search"),
    "search filters rows and keeps focus");
  await page.fill("#tr-search", ""); await page.waitForTimeout(200);

  // 5. Select session -> detail
  await page.click('.tr-table tbody tr[data-trsession="0"]'); await page.waitForTimeout(300);
  ok(await page.evaluate(() => document.querySelector("#tr-detail").textContent.includes("ROE")), "session detail opens with trainer");

  // 6. Upload eval for the session
  const evalPath = makeEval();
  await upload(page, '[data-traction="upload-eval"]', evalPath);
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 }).catch(() => {});
  const hasMapper = await page.evaluate(() => !!document.querySelector(".tr-mapper"));
  ok(hasMapper, "column mapper appears for new layout");
  if (hasMapper) { await page.click('[data-traction="save-map"]'); await page.waitForTimeout(300); }
  ok(await page.evaluate(() => !!document.querySelector('[data-traction="generate"]')), "Generate button after mapping");

  // 7. Generate -> themed report. NPS [10,9,8,6] + blank -> 25.0; T-score 55/12 = 4.58
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(500);
  ok(await kpiVal(page, "#tr-report-session", "Session NPS") === "25.0", "NPS 25.0 (blank answer skipped)");
  ok(await kpiVal(page, "#tr-report-session", "Trainer score") === "4.58", "trainer score 4.58");
  ok(await kpiVal(page, "#tr-report-session", "Responses") === "5", "responses 5 (form rows), NPS from the 4 answered");
  ok(await page.evaluate(() => !!document.querySelector("#tr-report-session .trf-nps-card") && /Interpretation/.test(document.querySelector("#tr-report-session").textContent)), "NPS card with interpretation (standalone layout)");
  ok(await page.evaluate(() => document.querySelectorAll("#tr-report-session .trf-group").length >= 1 && !!document.querySelector("#tr-report-session .trf-dist-bar")), "rating groups with distribution bars");
  ok(await page.evaluate(() => /Great pace <b>x<\/b>/.test(document.querySelector("#tr-report-session").textContent) &&
      !document.querySelector("#tr-report-session .trf-quote b")), "theme quotes shown escaped (no HTML injection)");
  ok(await page.evaluate(() => !/n\/a/.test([...document.querySelectorAll("#tr-report-session .trf-appendix-list li")].map(l => l.textContent).join("|"))), "trivial “n/a” comments left out of the appendix");
  ok(await page.evaluate(() => document.querySelector('.tr-table tbody tr[data-trsession="0"]').textContent.includes("Needs verifying")), "row marked Needs verifying");
  ok(await page.evaluate(() => /1 scored session/.test(document.querySelector(".tr-pending").textContent)), "pending bar shows 1 scored session not written");

  // 8. Leaderboard + summary (unwritten scores count)
  await page.click('[data-trview="board"]'); await page.waitForTimeout(300);
  ok(await page.evaluate(() => document.querySelector(".tr-board").textContent.includes("ROE")), "leaderboard lists trainer");
  await page.click('[data-trview="summary"]'); await page.waitForTimeout(300);
  ok(await kpiVal(page, "#tr-report", "Respondents") === "4", "summary Respondents = 4 (was always 0)");
  ok(await kpiVal(page, "#tr-report", "Avg NPS") === "25.0", "summary Avg NPS includes the generated (unwritten) score");
  ok(await page.evaluate(() => !!QBR.trainers._state.charts["tr-ch-month"]), "summary month chart drawn");

  // 9. Write-back: confirm lists rows, values land at the sheet offset, format-safe, added once.
  await page.click('[data-trview="sessions"]'); await page.waitForTimeout(300);
  await page.click('.tr-table tbody tr[data-trsession="0"]'); await page.waitForTimeout(300);
  async function doWriteback(tag) {
    let msg = "";
    page.once("dialog", d => { msg = d.message(); d.accept(); });
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
    const out = dl ? path.join(os.tmpdir(), "tr-out-" + tag + ".xlsx") : null;
    if (dl) await dl.saveAs(out);
    return { msg, out, name: dl ? dl.suggestedFilename() : "" };
  }
  const w1 = await doWriteback(1);
  ok(/Row 4 \(.*ROE.*\): NPS = 25\.0, Trainer Score = 4\.58/.test(w1.msg), "confirm lists the pending row (Excel row 4) with both scores");
  ok(w1.name === "tr-fixture-tracker-scored.xlsx", "write-back downloads <name>-scored.xlsx");
  if (w1.out) {
    const ws = XLSX.read(fs.readFileSync(w1.out)).Sheets.trial;
    ok(ws.I4 && ws.I4.v === 25 && ws.J3 && ws.J3.v === "Trainer Score" && ws.J4 && ws.J4.v === 4.58, "NPS → I4, \"Trainer Score\" J3/J4 (offset B3 honoured)");
    ok(!ws.I5 || ws.I5.v === "", "other session row untouched");
    const a = new Uint8Array(fs.readFileSync(trackerPath)), b = new Uint8Array(fs.readFileSync(w1.out));
    const za = SG.zipRead(a), zb = SG.zipRead(b), diff = [];
    za.entries.forEach(e => { const f = zb.byName.get(e.name); if (!f || f.crc !== e.crc) diff.push(e.name); });
    ok(diff.every(n => /^xl\/(worksheets\/sheet1|workbook)\.xml$/.test(n)) && diff.includes("xl/worksheets/sheet1.xml"),
      "format-safe: only the tracker sheet (+ workbook calc flag) changed (" + diff.join(",") + ")");
  } else { ok(false, "write-back downloaded"); }
  ok(await page.evaluate(() => document.querySelector('.tr-table tbody tr[data-trsession="0"]').textContent.includes("Scored")), "row marked Scored after write-back");
  ok(await page.evaluate(() => !!document.querySelector("#tr-report-session")), "report stays viewable after writing");
  // Re-attach + regenerate + write again: no duplicate "Trainer Score" column
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(400);
  await page.evaluate(() => { const c = QBR.trainers._state.tracker.sessions[0].computed; c.nps = 30; QBR.trainers.render(); });
  const w2 = await doWriteback(2);
  if (w2.out) {
    const rows = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(w2.out)).Sheets.trial, { header: 1, defval: "" });
    ok(rows[0].filter(h => h === "Trainer Score").length === 1 && rows[1][7] === 30, "second write-back builds on the first; \"Trainer Score\" exactly once");
  } else ok(false, "second write-back downloaded");

  // 10. Leaderboard null-safe sort + table opt-out
  ok(await page.evaluate(() => {
    const s = QBR.trainers._state.tracker.sessions;
    s[1].nps = 0; s[1].trainerScore = 3; s[0].nps = null; s[0].trainerScore = 4; s[0].written = null;
    QBR.trainers._state.view = "board"; QBR.trainers.render();
    return /SMITH/.test(document.querySelector(".tr-board tbody tr td:nth-child(2)").textContent);
  }), "leaderboard: NPS 0 ranks above a trainer with no NPS");
  ok(await page.evaluate(() => [...document.querySelectorAll("#trainers-body table")].every(t => t.getAttribute("data-resizable") === "static")),
    "trainers tables opt out of global column resizing");


  // 13. Live-edit safeguards: (i) a metadata-only touch (OneDrive) must not refuse an
  //     in-place save when the content is identical; (ii) an older tracker copy pauses writing.
  await page.evaluate(() => { QBR.trainers._state.view = "sessions"; QBR.trainers.render(); });
  await page.click('.tr-table tbody tr[data-trsession="1"]'); await page.waitForTimeout(300);
  await upload(page, '[data-traction="upload-eval"]', evalPath);
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 }).catch(() => {});
  if (await page.evaluate(() => !!document.querySelector(".tr-mapper"))) { await page.click('[data-traction="save-map"]'); await page.waitForTimeout(300); }
  const acceptAll = d => d.accept();   // session 2 already carries scores (step 10) + the eval is encoded on row 1: generate asks twice
  page.on("dialog", acceptAll);
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(500);
  page.off("dialog", acceptAll);
  ok(await page.evaluate(() => { const s = QBR.trainers._state.tracker.sessions[1]; return !!(s.computed && s.status === "ready"); }), "safeguards setup: second session scored and pending write");
  await page.evaluate(() => {
    const t = QBR.trainers._state.tracker, snap = t.bytes.slice();
    window.__fake = { bytes: snap, written: null, mtime: (t.mtime || 0) + 3600000 };
    t.handle = {
      getFile: async () => new File([window.__fake.bytes], t.name, { lastModified: window.__fake.mtime }),
      createWritable: async () => ({ write: async b => { window.__fake.written = new Uint8Array(b); }, close: async () => {}, abort: async () => {} })
    };
    t.size = snap.length;              // as last synced
    t.hash = QBR.fsHashOf(snap);       // as opened
    t.mtime = 1000;                    // opened earlier: metadata will differ, content will not
  });
  let dlA = null;
  page.once("dialog", d => d.accept());
  [dlA] = await Promise.all([page.waitForEvent("download", { timeout: 6000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  await page.waitForTimeout(400);
  const inPlace = await page.evaluate(() => ({ written: !!window.__fake.written, st: QBR.trainers._state.tracker.sessions[1].status }));
  ok(!dlA && inPlace.written && inPlace.st === "scored", "metadata-only change (OneDrive): content hash matches → in-place save proceeds, no download fallback");
  // (ii) lineage regression: the browser remembers 99 sessions for this file name; re-opening the 2-session file warns
  await page.evaluate(() => {
    const m = JSON.parse(localStorage.getItem("qbr-trainers-lineage") || "{}");
    m["tr-fixture-tracker.xlsx"] = { sessions: 99, hash: "stale-test-bogus", savedAt: "2026-10-23T00:00:00.000Z" };
    localStorage.setItem("qbr-trainers-lineage", JSON.stringify(m));
  });
  await upload(page, '[data-trpill="tracker"]', trackerPath);
  await page.waitForTimeout(700);
  ok(await page.evaluate(() => !!document.querySelector(".tr-stale") && !!QBR.trainers._state.trackerStale), "older tracker copy: stale banner shows at open, sessions remembered (99) vs this file (2)");
  await page.click('.tr-table tbody tr[data-trsession="0"]'); await page.waitForTimeout(300);
  await upload(page, '[data-traction="upload-eval"]', evalPath);
  await page.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 }).catch(() => {});
  if (await page.evaluate(() => !!document.querySelector(".tr-mapper"))) { await page.click('[data-traction="save-map"]'); await page.waitForTimeout(300); }
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(500);
  let dlB = null;
  page.once("dialog", d => d.accept());
  [dlB] = await Promise.all([page.waitForEvent("download", { timeout: 6000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  await page.waitForTimeout(300);
  const blocked = await page.evaluate(() => ({ status: (document.getElementById("tr-status") || {}).textContent || "", st: QBR.trainers._state.tracker.sessions[0].status }));
  ok(!dlB && blocked.st === "ready" && /looks older than the last one used here/.test(blocked.status), "stale tracker: write-back paused with the warning; nothing written or downloaded");
  await page.click('[data-traction="tracker-anyway"]'); await page.waitForTimeout(200);
  ok(await page.evaluate(() => !document.querySelector(".tr-stale") && !QBR.trainers._state.trackerStale), "Use this tracker anyway: banner clears");
  let dlC = null;
  page.once("dialog", d => d.accept());
  [dlC] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  ok(!!dlC && dlC.suggestedFilename() === "tr-fixture-tracker-scored.xlsx", "after confirming, write-back proceeds");

  const trErrs = errs.filter(e => /trainer/i.test(e));
  ok(trErrs.length === 0, "zero page errors from trainers code" + (trErrs.length ? ": " + trErrs.join(" | ") : ""));

  // 11. Flag off: trainers:false hides the Training group, render no-ops, no errors
  {
    const ctx = await browser.newContext();
    await ctx.addInitScript(() => { try { localStorage.setItem("qbr-feature-flags", JSON.stringify({ trainers: false })); } catch (e) {} });
    const p2 = await ctx.newPage(); const e2 = [];
    p2.on("pageerror", e => e2.push(String(e.message)));
    await p2.goto(APP, { waitUntil: "domcontentloaded" });
    await p2.waitForTimeout(800);
    ok(await p2.evaluate(() => { const g = document.querySelector('.sb-group[data-feature="trainers"]'); return !!g && getComputedStyle(g).display === "none"; }), "flag off: Training group hidden");
    ok(await p2.evaluate(() => { QBR.trainers.render(); return document.getElementById("trainers-body").innerHTML === ""; }), "flag off: render is a no-op");
    ok(e2.filter(e => /trainer/i.test(e)).length === 0, "flag off: zero trainers errors");
    await ctx.close();
  }
  // 12. localStorage blocked: mapper still works (asks every time), no crash
  {
    const ctx = await browser.newContext();
    await ctx.addInitScript(() => { const thrower = () => { throw new Error("blocked"); };
      Object.defineProperty(window, "localStorage", { get: thrower, configurable: true }); });
    const p3 = await ctx.newPage(); const e3 = [];
    p3.on("pageerror", e => e3.push(String(e.message)));
    await p3.goto(APP, { waitUntil: "domcontentloaded" });
    await p3.waitForTimeout(800);
    await p3.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });
    await p3.click('[data-tab="dash-trainers"]'); await p3.waitForTimeout(200);
    await upload(p3, '[data-trpill="eval"]', evalPath);
    await p3.waitForSelector(".tr-mapper", { timeout: 5000 }).catch(() => {});
    await p3.click('[data-traction="save-map"]').catch(() => {}); await p3.waitForTimeout(200);
    await p3.click('[data-traction="generate"]').catch(() => {}); await p3.waitForTimeout(300);
    ok(await kpiVal(p3, "#tr-report-session", "Session NPS") === "25.0", "storage blocked: eval-only report still works");
    ok(e3.filter(e => /trainer/i.test(e)).length === 0, "storage blocked: zero trainers errors");
    await ctx.close();
  }
  // 14. Wired feedback layout: pre-generate title, trainer type-in (no dropdown),
  //     prefilled date/org, write-back kept, export filename rule, toggles
  {
    const ctx = await browser.newContext();
    const p4 = await ctx.newPage(); const e4 = [];
    p4.on("pageerror", e => e4.push(String(e.message)));
    await p4.goto(APP, { waitUntil: "domcontentloaded" });
    await p4.waitForTimeout(800);
    await p4.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });
    await p4.click('[data-tab="dash-trainers"]'); await p4.waitForTimeout(300);
    await upload(p4, '[data-trpill="tracker"]', trackerPath);
    await p4.waitForFunction(() => document.querySelectorAll(".tr-table tbody tr[data-trsession]").length === 2, null, { timeout: 8000 }).catch(() => {});
    await p4.click('.tr-table tbody tr[data-trsession="0"]'); await p4.waitForTimeout(300);
    await upload(p4, '[data-traction="upload-eval"]', evalPath);
    await p4.waitForFunction(() => !!document.querySelector('[data-traction="generate"], .tr-mapper'), null, { timeout: 8000 }).catch(() => {});
    if (await p4.evaluate(() => !!document.querySelector(".tr-mapper"))) { await p4.click('[data-traction="save-map"]'); await p4.waitForTimeout(300); }
    ok(await p4.evaluate(() => { const i = document.querySelector("[data-trf=pretitle]"); return !!i && i.value === "Teams Workshop"; }), "pre-generate title input prefilled with the session title");
    await p4.fill("[data-trf=pretitle]", "My Custom Title");
    await p4.click('[data-traction="generate"]'); await p4.waitForTimeout(500);
    ok(await p4.evaluate(() => { const h = document.querySelector("#tr-report-session [data-trf=title]"); return !!h && h.textContent.trim() === "My Custom Title"; }), "renamed title carries into the generated cover");
    ok(await p4.evaluate(() => { const i = document.querySelector("#tr-report-session [data-trf=trainer]"); return !!i && i.value === "ROE, Sample III"; }), "trainer type-in prefilled from the session");
    ok(await p4.evaluate(() => !document.querySelector(".facilitator-dropdown, [data-trf=trainer-list]")), "no fixed-choice facilitator dropdown anywhere in the report");
    ok(await p4.evaluate(() => { const s = QBR.trainers._state.tracker.sessions[0]; const i = document.querySelector("#tr-report-session [data-trf=date]"); return !!i && i.value === s.dateISO; }), "session date prefilled from the session");
    ok(await p4.evaluate(() => { const i = document.querySelector("#tr-report-session [data-trf=org]"); return !!i && i.value === "MRCWI"; }), "organization prefilled from the session client");
    ok(await p4.evaluate(() => !!document.querySelector("#tr-report-session .tr-writeback")), "write-back bar kept under the new layout for a ready session");
    ok(await p4.evaluate(() => QBR.trainers._exportName("DELA CRUZ JUAN", "2026-10-25", "CSQC") === "DELA CRUZ JUAN-OCTOBER 25, 2026 - CSQC"), "export filename rule: trainer-date-school");
    await p4.fill("#tr-report-session [data-trf=trainer]", "DELA CRUZ JUAN");
    await p4.fill("#tr-report-session [data-trf=date]", "2026-10-25");
    await p4.fill("#tr-report-session [data-trf=org]", "CSQC");
    await p4.waitForTimeout(200);
    ok(await p4.evaluate(() => document.getElementById("tr-report-session").getAttribute("data-title") === "DELA CRUZ JUAN-OCTOBER 25, 2026 - CSQC"), "field edits sync the export filename live");
    await p4.check('[data-trf-toggle="appendix"]'); await p4.waitForTimeout(100);
    ok(await p4.evaluate(() => document.getElementById("tr-report-session").classList.contains("trf-show-appendix")), "appendix toggle reveals the full-responses appendix");
    await p4.uncheck('[data-trf-toggle="brand"]'); await p4.waitForTimeout(100);
    ok(await p4.evaluate(() => document.getElementById("tr-report-session").classList.contains("trf-nobrand")), "branding toggle hides the cover brand header");
    ok(e4.filter(e => /trainer/i.test(e)).length === 0, "wired layout: zero trainers errors");
    await ctx.close();
  }
  await browser.close();
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
