/* Trainers tab — v7 features (browser, file://). Synthetic workbooks only.
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers-v2.browser.cjs
 * A: Add session from an eval (dates from response timestamps, typeable fields)
 * B: new rows are WRITTEN (format-safe): real Excel date, Session ID header, right row, scores on it
 * C: in-place save into a linked file handle; refuses to overwrite a file changed since opening
 * D: manual Add session + remove before writing
 * E: Print / PDF — summary + session report build a print copy (charts as images) and call print()
 * F: eval on its own (no tracker) -> feedback report, nothing to write */
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
    [46100, "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, 70],
    [46101, "SANTOS, Maria", "MS-103A", "Workshop P1", "Hill Crest College of Carmel", 0.54, 0.62, ""]
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Course Code", "Product", "Session Title", "Estimated Duration"],
    ["MS-102A", "Microsoft", "Teams Workshop", 4]]), "03 Course List");
  const p = path.join(os.tmpdir(), "trv2-tracker.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
function makeEval() {
  const head = ["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic",
    "Communication and presentation skills", "Ability to engage and connect with participants",
    "Overall, how satisfied were you with the training?",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
    "Why did you give this rating?", "What did you like best about the session?"];
  const rows = [head];
  for (let i = 0; i < 6; i++)
    rows.push([i + 1, 46149.4 + i / 1440, "HCCCI", "Strongly Agree", i % 2 ? "Agree" : "Strongly Agree", "Neither agree nor disagree",
               i ? "Very satisfied" : "Dissatisfied", [10, 9, 10, 5, 9, 10][i], i === 3 ? "Too fast" : "Good", "Hands-on practice"]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  // filename date deliberately different from the answer timestamps (2026-05-07)
  const p = path.join(os.tmpdir(), "Feedback_-_01012026_-_AM_-_HCCCI_1-6.xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
async function upload(page, sel, f) {
  const [c] = await Promise.all([page.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null), page.click(sel)]);
  if (c) await c.setFiles(f);
  await page.waitForTimeout(500);
}
async function saveMapIfAsked(page) {
  if (await page.$(".tr-mapper")) { await page.click('[data-traction="save-map"]'); await page.waitForTimeout(300); }
}
async function fresh(browser) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page._errs = [];
  page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(500);
  await page.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });
  await page.click('[data-tab="dash-trainers"]'); await page.waitForTimeout(200);
  return page;
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const trk = makeTracker(), evf = makeEval();
  let page = await fresh(browser);

  /* ---- A: add an eval as a new session ---- */
  await upload(page, '[data-trpill="tracker"]', trk);
  await upload(page, '[data-trpill="eval"]', evf);
  await saveMapIfAsked(page);
  await page.waitForSelector('[data-traction="new-session"]', { timeout: 5000 });
  await page.click('[data-traction="new-session"]');
  await page.waitForSelector("#trns-date", { timeout: 5000 });
  ok(await page.inputValue("#trns-date") === "2026-05-07", "A1 date from response timestamps, not filename");
  ok(await page.evaluate(() => /response timestamps/.test(document.querySelector("#trns-date").closest("div").textContent)), "A2 date source hint shown");
  ok(await page.evaluate(() => { const i = document.querySelector("#trns-trainer"); return i.tagName === "INPUT" && [...document.querySelectorAll("#trns-trainers option")].some(o => o.value === "ROE, Sample III"); }),
    "A3 trainer: typeable input with roster suggestions");
  ok(await page.inputValue("#trns-school") === "Hill Crest College of Carmel" || await page.inputValue("#trns-school") === "HCCCI", "A4 school prefilled from filename");
  await page.fill("#trns-title", "Teams Work"); await page.waitForTimeout(100);
  ok(await page.inputValue("#trns-code") === "MS-102A", "A5 course code auto-filled from a unique course-name prefix");
  await page.fill("#trns-trainer", "ROE, Sample III");
  await page.fill("#trns-school", "Typed Academy & <Test>");
  await page.click('[data-traction="save-new-session"]'); await page.waitForTimeout(300);
  const added = await page.evaluate(() => { const t = QBR.trainers._state.tracker; const s = t.sessions[t.sessions.length - 1]; return { s: { client: s.client, dateISO: s.dateISO, id: s.id, rowIdx: s.rowIdx, isNew: s.isNew }, n: t.addedRows.length, att: QBR.trainers._state.eval.attachedTo === t.sessions.length - 1 }; });
  ok(added.n === 1 && added.s.isNew && added.s.client === "Typed Academy & <Test>" && added.s.dateISO === "2026-05-07", "A6 session added (typed school, timestamp date)");
  ok(added.s.id === "SES-2026-0001" && added.s.rowIdx === 3 && added.att, "A7 Session ID issued, placed right after the last row, eval attached");
  ok(await page.evaluate(() => !document.querySelector("#trainers-body table td script, #trainers-body table td test")), "A8 typed school rendered escaped");
  ok(await page.evaluate(() => /1 new session/.test(document.querySelector(".tr-pending").textContent)), "A9 pending bar shows the unwritten new session");

  /* ---- B: generate + write: the new row itself is written ---- */
  await saveMapIfAsked(page);
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(400);
  ok(await page.evaluate(() => { const k = [...document.querySelectorAll("#tr-report-session .kpi")].find(x => /Satisfaction/.test(x.textContent)); return k && /83%/.test(k.textContent); }),
    "B1 satisfaction = top-2 share (5 of 6)");
  ok(await page.evaluate(() => { const r = QBR.trainers._state.tracker.sessions[2].computed.report; return r.ratings.some(x => /engage/.test(x.header) && x.avg === 3); }),
    "B2 'Neither agree nor disagree' rates 3 in the item chart");
  let msg = "";
  page.once("dialog", d => { msg = d.message(); d.accept(); });
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  ok(/New row 4: 2026-05-07 · ROE, Sample III · MS-102A · Typed Academy & <Test> \(SES-2026-0001\)/.test(msg) && /Row 4 \(2026-05-07/.test(msg), "B3 confirm lists the new row and its scores");
  if (dl) {
    const out = path.join(os.tmpdir(), "trv2-out.xlsx"); await dl.saveAs(out);
    const wb = XLSX.read(fs.readFileSync(out), { cellDates: true, cellNF: true }), ws = wb.Sheets.trial;
    ok(ws.A4 && ws.A4.v instanceof Date && ws.A4.v.getFullYear() === 2026 && ws.A4.v.getMonth() === 4 && ws.A4.v.getDate() === 7, "B4 date written as a real Excel date (2026-05-07)");
    ok(ws.B4 && ws.B4.v === "ROE, Sample III" && ws.C4.v === "MS-102A" && ws.D4.v === "Teams Workshop" && ws.E4.v === "Typed Academy & <Test>", "B5 trainer, code, full course title (from the typed prefix), school written on row 4");
    ok(Math.abs(ws.F4.v - 8 / 24) < 1e-3 && Math.abs(ws.G4.v - 12 / 24) < 1e-3, "B6 times written as day fractions");
    ok(ws.I1 && ws.I1.v === "Session ID" && ws.I4.v === "SES-2026-0001" && ws.J1 && ws.J1.v === "Trainer Score", "B7 Session ID + Trainer Score headers added once, in order");
    ok(ws.H4 && ws.H4.v === 66.7 && ws.J4 && typeof ws.J4.v === "number", "B8 NPS (5 promoters, 1 detractor of 6 = 66.7) and trainer score land on the new row");
    ok(ws.H2.v === 70 && ws.B3.v === "SANTOS, Maria", "B9 existing rows untouched");
    const za = SG.zipRead(new Uint8Array(fs.readFileSync(trk))), zb = SG.zipRead(new Uint8Array(fs.readFileSync(out))), diff = [];
    za.entries.forEach(e => { const f = zb.byName.get(e.name); if (!f || f.crc !== e.crc) diff.push(e.name); });
    ok(diff.every(n => /^xl\/(worksheets\/sheet1|workbook|styles)\.xml$/.test(n)), "B10 format-safe: only sheet (+calc flag, date style) parts changed (" + diff.join(",") + ")");
  } else { ok(false, "B4 write-back downloaded"); }
  ok(await page.evaluate(() => { const t = QBR.trainers._state.tracker; return t.addedRows.length === 0 && !t.sessions[2].isNew && t.sessions[2].status === "scored" && !document.querySelector(".tr-pending"); }),
    "B11 after writing: nothing pending, session scored");

  /* ---- C: in-place save via a linked handle ---- */
  await page.evaluate(() => {
    const S = QBR.trainers._state, t = S.tracker;
    window.__w = 0; window.__dl = 0;
    const orig = URL.createObjectURL; URL.createObjectURL = function (b) { window.__dl++; return "blob:mock"; };
    t.mtime = 1000; t.size = t.bytes.length;
    t.handle = { getFile: async () => ({ lastModified: window.__mtime || 1000, size: t.size }),
                 createWritable: async () => ({ write: async (b) => { window.__w = b.byteLength || b.length; }, close: async () => {} }) };
    window.confirm = () => true; window.alert = (m) => { window.__alert = m; };
    S.tracker.sessions[1].computed = { nps: 40, trainerScore: 4, n: 3, promoters: 1, passives: 1, detractors: 1, report: null };
    S.tracker.sessions[1].status = "ready"; QBR.trainers.render();
  });
  await page.click('.tr-pending [data-traction="writeback"]');
  await page.waitForFunction(() => window.__w > 0 || window.__dl > 0, null, { timeout: 8000 }).catch(() => {});
  ok(await page.evaluate(() => window.__w > 0 && window.__dl === 0), "C1 linked + unchanged file: saved in place, no download");
  await page.evaluate(() => { const S = QBR.trainers._state; window.__w = 0; window.__mtime = 99999;
    S.tracker.sessions[1].computed = { nps: 45, trainerScore: 4, n: 3, promoters: 1, passives: 1, detractors: 1, report: null };
    S.tracker.sessions[1].status = "ready"; QBR.trainers.render(); });
  await page.click('.tr-pending [data-traction="writeback"]');
  await page.waitForFunction(() => window.__w > 0 || window.__dl > 0, null, { timeout: 8000 }).catch(() => {});
  ok(await page.evaluate(() => window.__w === 0 && window.__dl === 1 && /changed after you opened it/.test(window.__alert || "")),
    "C2 file changed in Excel since opening: NOT overwritten, copy downloaded + explained");

  /* ---- D: manual add + remove before writing ---- */
  await page.click('[data-traction="add-session"]'); await page.waitForTimeout(200);
  await page.fill("#trns-date", "2026-06-01"); await page.fill("#trns-trainer", "NEW, Trainer"); await page.fill("#trns-code", "X-1"); await page.fill("#trns-school", "SPC");
  await page.click('[data-traction="save-new-session"]'); await page.waitForTimeout(200);
  ok(await page.evaluate(() => { const t = QBR.trainers._state.tracker; const s = t.sessions[t.sessions.length - 1]; return s.isNew && s.id === "SES-2026-0002" && t.roster.indexOf("NEW, Trainer") >= 0; }),
    "D1 manual session added (next Session ID, trainer joins roster)");
  await page.click('[data-traction="remove-new"]'); await page.waitForTimeout(200);
  ok(await page.evaluate(() => { const t = QBR.trainers._state.tracker; return t.addedRows.length === 0 && t.headerAdds.length === 0 && t.sessions.length === 3; }), "D2 removed before writing: nothing pending");

  /* ---- E: print to PDF ---- */
  await page.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; window.__printHtml = document.getElementById("tr-print").innerHTML; window.__theme = document.body.getAttribute("data-theme"); }; document.body.setAttribute("data-theme", "dark"); });
  await page.click('[data-trview="summary"]'); await page.waitForTimeout(400);
  await page.click('[data-traction="print-report"]'); await page.waitForTimeout(800);
  const pr = await page.evaluate(() => ({ n: window.__printed, imgs: (window.__printHtml.match(/<img[^>]+data:image\/png/g) || []).length, canv: /<canvas/.test(window.__printHtml),
    btn: /data-traction/.test(window.__printHtml), theme: window.__theme, cls: document.body.classList.contains("tr-printing") }));
  ok(pr.n === 1, "E1 Summary report (PDF) opens the print dialog (was a no-op)");
  ok(pr.imgs === 2 && !pr.canv, "E2 print copy has both charts as images, no live canvases");
  ok(!pr.btn && pr.theme === "light" && pr.cls, "E3 print copy has no buttons, prints in light theme");
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  ok(await page.evaluate(() => !document.getElementById("tr-print") && !document.body.classList.contains("tr-printing") && document.body.getAttribute("data-theme") === "dark"),
    "E4 afterprint cleans up and restores the theme");
  await page.click('[data-trview="sessions"]'); await page.waitForTimeout(200);
  await page.click('.tr-table tbody tr[data-trsession="2"]'); await page.waitForTimeout(300);
  await page.click('[data-traction="print-session"]'); await page.waitForTimeout(800);
  ok(await page.evaluate(() => window.__printed === 2 && /Participant feedback summary/.test(window.__printHtml) && /Recommendations/.test(window.__printHtml)),
    "E5 Session report (PDF) prints the feedback report");
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  ok(page._errs.length === 0, "zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));
  await page.close();

  /* ---- G: a different eval for an already-generated session drops the old results ---- */
  page = await fresh(browser);
  await upload(page, '[data-trpill="tracker"]', trk);
  await page.click('.tr-table tbody tr[data-trsession="1"]'); await page.waitForTimeout(200);
  await upload(page, '[data-traction="upload-eval"]', evf);
  await saveMapIfAsked(page);
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(300);
  const evB = path.join(os.tmpdir(), "Other_-_01012026_-_PM_-_SPC.xlsx");
  fs.copyFileSync(evf, evB);
  await upload(page, '[data-trpill="eval"]', evB);
  await saveMapIfAsked(page);
  ok(await page.evaluate(() => { const s = QBR.trainers._state.tracker.sessions[1]; return !s.computed && s.status === "unscored" && !!document.querySelector('#tr-detail [data-traction="generate"]') && !document.querySelector(".tr-pending"); }),
    "G1 new eval file on a generated session: old results dropped, Generate shown, nothing pending");
  /* ---- H: after a lossy fallback the linked original is never written again ---- */
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(300);
  await page.evaluate(() => {
    const t = QBR.trainers._state.tracker; window.__w = 0; window.__dl = 0;
    URL.createObjectURL = function () { window.__dl++; return "blob:mock"; };
    t.mtime = 1000; t.size = t.bytes.length;
    t.handle = { getFile: async () => ({ lastModified: 1000, size: t.size }), createWritable: async () => ({ write: async () => { window.__w++; }, close: async () => {} }) };
    window.__origSG = QBR.xlsxSurgical.surgicalSave; QBR.xlsxSurgical.surgicalSave = async () => ({ ok: false, reason: "test" });
    window.confirm = () => true; window.alert = () => {};
  });
  await page.click('.tr-pending [data-traction="writeback"]'); await page.waitForTimeout(800);
  ok(await page.evaluate(() => window.__w === 0 && window.__dl === 1 && QBR.trainers._state.tracker.handle === null), "H1 lossy fallback: copy downloaded, file unlinked");
  await page.evaluate(() => { QBR.xlsxSurgical.surgicalSave = window.__origSG; const s = QBR.trainers._state.tracker.sessions[0];
    s.computed = { nps: 10, trainerScore: 4, n: 2, promoters: 1, passives: 0, detractors: 1, report: null }; s.status = "ready"; QBR.trainers.render(); });
  await page.click('.tr-pending [data-traction="writeback"]'); await page.waitForTimeout(800);
  ok(await page.evaluate(() => window.__w === 0 && window.__dl === 2), "H2 next write-back downloads again — never overwrites the original with the stripped copy");
  ok(page._errs.length === 0, "G/H zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));
  await page.close();

  /* ---- F: eval only, no tracker ---- */
  page = await fresh(browser);
  await upload(page, '[data-trpill="eval"]', evf);
  await saveMapIfAsked(page);
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(400);
  ok(await page.evaluate(() => !!document.querySelector("#tr-report-session") && !document.querySelector('[data-traction="writeback"]') &&
      /no tracker loaded/.test(document.querySelector("#tr-detail").textContent)), "F1 eval-only: feedback report, nothing to write");
  await upload(page, '[data-trpill="tracker"]', trk);
  ok(await page.evaluate(() => !!document.querySelector('[data-traction="attach"]') || !!document.querySelector('[data-traction="new-session"]')),
    "F2 loading a tracker afterwards keeps the eval and offers to match it");
  ok(page._errs.length === 0, "F zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
