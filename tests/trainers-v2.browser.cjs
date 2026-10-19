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
  if (await page.$('[data-traction="q-map"]')) { await page.click('[data-traction="q-map"]'); await page.waitForTimeout(200); }
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
  await page.waitForSelector("[data-trq-target]", { timeout: 5000 });
  ok(await page.evaluate(() => QBR.trainers._state.queue.length === 1 && QBR.trainers._state.queue[0].target === -1), "A0 eval goes to the eval list (no session on that date, nothing auto-picked)");
  await page.selectOption("[data-trq-target]", "new");
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
  const added = await page.evaluate(() => { const t = QBR.trainers._state.tracker; const s = t.sessions[t.sessions.length - 1]; return { s: { client: s.client, dateISO: s.dateISO, id: s.id, rowIdx: s.rowIdx, isNew: s.isNew }, n: t.addedRows.length, att: QBR.trainers._state.queue[0].target === t.sessions.length - 1 }; });
  ok(added.n === 1 && added.s.isNew && added.s.client === "Typed Academy & <Test>" && added.s.dateISO === "2026-05-07", "A6 session added (typed school, timestamp date)");
  ok(added.s.id === "SES-2026-0001" && added.s.rowIdx === 3 && added.att, "A7 Session ID issued, placed right after the last row, eval file set to it");
  ok(await page.evaluate(() => !document.querySelector("#trainers-body table td script, #trainers-body table td test")), "A8 typed school rendered escaped");
  ok(await page.evaluate(() => /1 new session/.test(document.querySelector(".tr-pending").textContent)), "A9 pending bar shows the unwritten new session");

  /* ---- B: generate + write: the new row itself is written ---- */
  await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(400);
  await page.click('[data-traction="q-open"]'); await page.waitForTimeout(400);
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
  ok(pr.imgs === 3 && !pr.canv, "E2 print copy has all 3 charts as images, no live canvases");
  ok(!pr.btn && pr.theme === "light" && pr.cls, "E3 print copy has no buttons, prints in light theme");
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  ok(await page.evaluate(() => !document.getElementById("tr-print") && !document.body.classList.contains("tr-printing") && document.body.getAttribute("data-theme") === "dark"),
    "E4 afterprint cleans up and restores the theme");
  await page.click('[data-trview="sessions"]'); await page.waitForTimeout(200);
  await page.click('.tr-table tbody tr[data-trsession="2"]'); await page.waitForTimeout(300);
  await page.click('[data-traction="print-session"]'); await page.waitForTimeout(800);
  ok(await page.evaluate(() => window.__printed === 2 && /Participant feedback summary/.test(window.__printHtml) && /Recommendations/.test(window.__printHtml)),
    "E5 Session report (PDF) prints the feedback report");
  ok(await page.evaluate(() => /tr-p-(pair|stack)/.test(window.__printHtml) && /tr-p-(keep|flow)/.test(window.__printHtml) && /--tr-tiles: ?5/.test(window.__printHtml)),
    "E6 print layout decided from content (pairs / stacked long tables, 5 tiles per row)");
  ok(await page.evaluate(() => { const f = QBR.trainers._state.lastFit; return !!f && f.pages >= 1 && !document.querySelector("#tr-print .tr-print-foot") && /printed /.test(window.__printHtml); }),
    "E7 page fitting ran before print (measured at paper width); no stray footer block");
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
  ok(await page.evaluate(() => { const S = QBR.trainers._state, s = S.tracker.sessions[1];
      return s.computed && /Feedback_/.test(s.computed.evalName) && S.queue.length === 1 && S.queue[0].target === -1 && /Feedback_/.test(document.querySelector("#tr-detail").textContent); }),
    "G1 'Add eval exports' with a session open: the file goes to the list — the open session keeps its own eval and results (reported bug)");
  await saveMapIfAsked(page);
  await page.selectOption("[data-trq-target]", "pick"); await page.waitForTimeout(200);
  ok(await page.evaluate(() => /Click the session for/.test(document.querySelector(".tr-picking").textContent)), "G1b 'Pick in the session list' asks for a click");
  await page.click('.tr-table tbody tr[data-trsession="1"]'); await page.waitForTimeout(200);
  ok(await page.evaluate(() => /replaces/.test(document.querySelector(".tr-queue").textContent) === false), "G2 (session not yet written: no 'replaces score' warning)");
  await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(300);
  ok(await page.evaluate(() => { const s = QBR.trainers._state.tracker.sessions[1]; return s.computed && /Other_/.test(s.computed.evalName) && s.status === "ready"; }),
    "G3 assigning the new file and generating replaces the session's results with the new file's");
  /* ---- H: after a lossy fallback the linked original is never written again ---- */
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

  /* ---- Q: bulk — many eval files at once, matched to their sessions ---- */
  {
    const rows = [["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
      [46149, "ROE, Sample III", "MS-102A", "Teams AM", "HCCCI", 0.333, 0.5, ""],
      [46149, "SANTOS, Maria", "MS-102A", "Teams PM", "HCCCI", 0.54, 0.7, ""],
      [46150, "LANE, Jun", "MS-103A", "Word", "St. Pine College", "8:00 AM", "12:00 PM", ""],
      [46150, "DEL SAMPLE, Ana", "MS-104A", "Excel", "Pico Catholic School", 0.333, 0.5, ""]];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
    const tq = path.join(os.tmpdir(), "trq-tracker.xlsx"); fs.writeFileSync(tq, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
    const src = fs.readFileSync(evf), dir = fs.mkdtempSync(path.join(os.tmpdir(), "trq-"));
    const names = ["Form_-_05072026_-_AM_-_HCCCI_1-6.xlsx", "Form_-_05072026_-_PM_-_HCCCI_1-6.xlsx", "Form_-_05082026_-_AM_-_SPC_1-6.xlsx",
                   "Form_-_05082026_-_AM_-_Unknown_School_1-6.xlsx", "notes.txt"];
    // different answers per file (same answers = same fingerprint = treated as the same file)
    const evRows = XLSX.utils.sheet_to_json(XLSX.read(src).Sheets.Sheet1, { header: 1 });
    const bufFor = (k) => { const r = evRows.map((row, i) => i ? [row[0] + k * 100].concat(row.slice(1)) : row);
      const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet(r), "Sheet1"); return XLSX.write(w, { bookType: "xlsx", type: "buffer" }); };
    const files = names.map((n, k) => { const f = path.join(dir, n); fs.writeFileSync(f, n.endsWith(".txt") ? "x" : bufFor(k)); return f; });
    page = await fresh(browser);
    await upload(page, '[data-trpill="tracker"]', tq);
    const [ch] = await Promise.all([page.waitForEvent("filechooser"), page.click('[data-trpill="eval"]')]);
    ok(ch.isMultiple(), "Q1 'Add eval exports' lets you select many files");
    await ch.setFiles(files); await page.waitForTimeout(1200);
    const st = () => page.evaluate(() => QBR.trainers._state.queue.map(q => q.target));
    ok(JSON.stringify(await st()) === JSON.stringify([0, 1, 2, -1]), "Q2 4 xlsx files listed (txt ignored): AM→AM session, PM→PM session, SPC→St. Pine College (acronym), unknown school not auto-picked (" + JSON.stringify(await st()) + ")");
    ok(await page.evaluate(() => { const q = QBR.trainers._state.queue[3]; return q.cands.length === 2 && q.cands.indexOf(2) >= 0; }), "Q3 unknown school: same-day sessions offered in its list");
    // the same files again: duplicates skipped
    const [ch2] = await Promise.all([page.waitForEvent("filechooser"), page.click('[data-trpill="eval"]')]);
    await ch2.setFiles(files.slice(0, 2)); await page.waitForTimeout(800);
    ok((await st()).length === 4, "Q4 files already in the list are not added twice");
    // two files set to one session -> neither generated
    await page.selectOption('[data-trq-target="' + (await page.evaluate(() => QBR.trainers._state.queue[3].id)) + '"]', "2"); await page.waitForTimeout(200);
    await saveMapIfAsked(page);
    await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(500);
    ok(await page.evaluate(() => { const S = QBR.trainers._state, s = S.tracker.sessions;
        return !!s[0].computed && !!s[1].computed && !s[2].computed && /same session/.test(S.queue[2].note) && /same session/.test(S.queue[3].note); }),
      "Q5 Generate: 2 sessions scored; two files on one session are held back with a note");
    await page.selectOption('[data-trq-target="' + (await page.evaluate(() => QBR.trainers._state.queue[3].id)) + '"]', "3"); await page.waitForTimeout(200);
    await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(500);
    ok(await page.evaluate(() => QBR.trainers._state.tracker.sessions.every(s => s.computed && s.status === "ready")), "Q6 after fixing the pick: all 4 sessions generated");
    ok(await page.evaluate(() => /4 scored sessions/.test(document.querySelector(".tr-pending").textContent) && document.querySelectorAll(".tr-queue .risk-badge.risk-high").length === 4), "Q7 list shows Generated per file; pending bar: 4 sessions");
    let msg = ""; page.once("dialog", d => { msg = d.message(); d.accept(); });
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
    ok(/Write 4 session scores/.test(msg) && (msg.match(/^Row \d/mg) || []).length === 4, "Q8 one confirm lists all 4 rows");
    if (dl) {
      const out = path.join(os.tmpdir(), "trq-out.xlsx"); await dl.saveAs(out);
      const ws = XLSX.read(fs.readFileSync(out)).Sheets.trial;
      ok([2, 3, 4, 5].every(r => ws["H" + r] && ws["H" + r].v === 66.7 && ws["I" + r] && typeof ws["I" + r].v === "number"), "Q9 all 4 rows written in one file (NPS + Trainer Score)");
    } else ok(false, "Q9 downloaded");
    ok(await page.evaluate(() => document.querySelectorAll(".tr-queue .risk-badge.risk-none").length === 4), "Q10 list shows Written per file");
    // Summary report: trainer tiles, courses conducted, trainer table -> trainer profile
    await page.click('[data-trview="summary"]'); await page.waitForTimeout(400);
    const tile = (l) => page.evaluate((l) => { const k = [...document.querySelectorAll("#tr-report .kpi")].find(x => x.querySelector(".kpi-lbl").textContent.trim() === l); return k ? k.querySelector(".kpi-val").textContent.trim() : null; }, l);
    ok(await tile("Trainers active") === "4" && await tile("Trainings conducted") === "4" && await tile("Courses delivered") === "3" && await tile("Schools served") === "3" && await tile("Eval coverage") === "100%",
      "S1 summary tiles: trainers active, trainings conducted, courses delivered, schools served, eval coverage");
    ok(await page.evaluate(() => { const r = [...document.querySelectorAll(".tr-courses tbody tr")]; return r.length === 3 && /Teams AM|Teams PM/.test(r[0].textContent) && r[0].querySelector("b").textContent === "2"; }),
      "S2 courses conducted: times conducted per course, most first (MS-102A ×2)");
    ok(await page.evaluate(() => document.querySelectorAll(".tr-trainers tbody tr").length === 4 && !!QBR.trainers._state.charts["tr-ch-course"]), "S3 trainer table (one row per trainer) + course chart");
    await page.selectOption("#tr-rep-period", "2026-Q2"); await page.waitForTimeout(300);
    ok(await page.evaluate(() => /vs previous/.test(document.querySelector(".tr-trainers thead").textContent)) && await tile("Trainings conducted") === "4", "S4 period filter (2026 Q2) adds the 'vs previous' column");
    await page.click('[data-trtrainer="LANE, Jun"]'); await page.waitForTimeout(300);
    ok(await page.evaluate(() => QBR.trainers._state.repTrainer === "LANE, Jun" && !document.querySelector(".tr-trainers") && /Facilitator items/.test(document.querySelector("#tr-report").textContent) &&
        /Trainer profile \(PDF\)/.test(document.querySelector('[data-traction="print-report"]').textContent)) && await tile("Trainings conducted") === "1",
      "S5 clicking a trainer opens their profile (tiles, facilitator items, Trainer profile (PDF))");
    await page.evaluate(() => { window.print = () => { window.__ph = document.getElementById("tr-print").innerHTML; }; });
    await page.click('[data-traction="print-report"]'); await page.waitForTimeout(700);
    ok(await page.evaluate(() => /Trainer profile — LANE, Jun/.test(window.__ph) && /Courses conducted/.test(window.__ph) && !/click a name/.test(window.__ph)), "S6 trainer profile prints with its own title");
    await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
    await page.click('[data-traction="rep-clear"]'); await page.waitForTimeout(300);
    ok(await page.evaluate(() => !QBR.trainers._state.repTrainer && !QBR.trainers._state.repPeriod && !!document.querySelector(".tr-trainers")), "S7 Clear filters returns to all trainers");
    ok(page._errs.length === 0, "Q zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));
    await page.close();
  }

  /* ---- R: duplicates + already-scored protection + Eval Source (real filename formats) ---- */
  {
    const rows = [["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score", "Trainer Score"],
      [45946, "DOEÑA, Alex", "MS-101A", "Microsoft", "Academia de San Valle Fuerte", 0.333, 0.5, "", ""],
      [45946, "DOEÑA, Alex", "MS-101A", "Microsoft", "Academia de San Valle Fuerte", 0.5417, 0.7083, "", ""],
      [45989, "LANE, Jun", "MS-102A", "Microsoft", "Saint Garnet Academy", 0.333, 0.5, 80, 4.5]];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
    const tr = path.join(os.tmpdir(), "trr-tracker.xlsx"); fs.writeFileSync(tr, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "trr-"));
    const mk = (name, day, frac, idBase, nps) => {
      const r = [["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic", "Communication and presentation skills",
        "Ability to engage and connect with participants", "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague"]];
      for (let i = 0; i < 4; i++) r.push([idBase + i, day + frac + i / 1440, "School", "Strongly Agree", "Agree", "Strongly Agree", nps[i]]);
      const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet(r), "Sheet1");
      const f = path.join(dir, name); fs.writeFileSync(f, XLSX.write(w, { bookType: "xlsx", type: "buffer" })); return f;
    };
    const fAM = mk("Training Feedback Form - Oct. 16, 2025 [ASVF] - Micro.xlsx", 45946, 0.49, 1, [10, 9, 10, 9]);       // answered ~11:45 AM
    const fPM = mk("Training Feedback Form - Oct. 16, 2025 [ASVF] - Microsoft 3.xlsx", 45946, 0.70, 101, [10, 6, 8, 9]); // answered ~4:50 PM
    const fSGA = mk("Training Feedback Form - Nov. 28, 2025 [SGA] - Microsoft 36.xlsx", 45990, 0.49, 201, [9, 9, 10, 10]); // answered the next day
    const fCopy = path.join(dir, "copy of Micro.xlsx"); fs.copyFileSync(fAM, fCopy);
    page = await fresh(browser);
    await upload(page, '[data-trpill="tracker"]', tr);
    const [c1] = await Promise.all([page.waitForEvent("filechooser"), page.click('[data-trpill="eval"]')]);
    await c1.setFiles([fAM, fPM, fSGA, fCopy]); await page.waitForTimeout(1200);
    const Q = () => page.evaluate(() => QBR.trainers._state.queue.map(q => ({ n: q.name, t: q.target, h: q.hint && q.hint.dateISO, s: q.hint && q.hint.school })));
    let q = await Q();
    ok(q.length === 3, "R1 renamed copy with the same answers is not added (" + q.length + " files)");
    ok(q[0].h === "2025-10-16" && q[0].s === "ASVF" && q[2].h === "2025-11-28" && q[2].s === "SGA", "R2 real filenames read: 'Oct. 16, 2025 [ASVF]', 'Nov. 28, 2025 [SGA]'");
    ok(q[0].t === 0 && q[1].t === 1, "R3 two same-day ASVF sessions split by answer time: Micro → 8 AM session, Microsoft 3 → 1 PM session");
    ok(q[2].t === 2, "R4 SGA matched by the filename date (Nov 28), not the answer date (Nov 29)");
    ok(await page.evaluate(() => { const o = [...document.querySelectorAll("[data-trq-target] option")].map(x => x.textContent); return o.some(x => /8:00 AM–12:00 PM.*row 2/.test(x)) && o.some(x => /1:00 PM–5:00 PM.*row 3/.test(x)); }),
      "R5 session labels show time and row, so same-day sessions differ");
    ok(await page.evaluate(() => /Already scored/.test(document.querySelector('.tr-queue tbody tr:nth-child(3)').textContent) && !!document.querySelector("[data-trq-replace]")), "R6 SGA row: 'Already scored' NPS 80.0 + 'Replace existing score' box");
    await saveMapIfAsked(page);
    ok(await page.evaluate(() => /Generate 2 matched files/.test(document.querySelector('[data-traction="q-generate"]').textContent)), "R7 Generate count leaves out the already-scored session");
    await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(500);
    ok(await page.evaluate(() => { const s = QBR.trainers._state.tracker.sessions; return !!s[0].computed && !!s[1].computed && !s[2].computed && s[2].nps === 80; }), "R8 generated ASVF ×2; SGA score untouched");
    let msg = ""; page.once("dialog", d => { msg = d.message(); d.accept(); });
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
    let out = null;
    if (dl) { out = path.join(os.tmpdir(), "trr-out.xlsx"); await dl.saveAs(out); }
    const ws = out ? XLSX.read(fs.readFileSync(out)).Sheets.trial : {};
    ok(ws.J1 && ws.J1.v === "Eval Source" && /Micro\.xlsx · fp:[0-9a-f]{8}$/.test(ws.J2 && ws.J2.v) && /Microsoft 3\.xlsx · fp:/.test(ws.J3 && ws.J3.v) && !ws.J4, "R9 'Eval Source' column records file + fingerprint per written row");
    // tick Replace for SGA -> confirm shows old -> new
    await page.click("[data-trq-replace]"); await page.waitForTimeout(200);
    await page.click('[data-traction="q-generate"]'); await page.waitForTimeout(400);
    msg = ""; page.once("dialog", d => { msg = d.message(); d.dismiss(); });
    await page.click('.tr-pending [data-traction="writeback"]'); await page.waitForTimeout(400);
    ok(/REPLACES NPS 80\.0 → 100\.0, Trainer Score 4\.50 → /.test(msg), "R10 replacing is opt-in and the confirm shows old → new");
    await page.close();
    // next time (any browser): the written tracker recognises the files
    page = await fresh(browser);
    if (out) await upload(page, '[data-trpill="tracker"]', out);
    const [c2] = await Promise.all([page.waitForEvent("filechooser"), page.click('[data-trpill="eval"]')]);
    await c2.setFiles([fCopy, fPM]); await page.waitForTimeout(1000);
    ok(await page.evaluate(() => { const S = QBR.trainers._state; return S.queue.length === 2 && S.queue[0].target === 0 && S.queue[1].target === 1 &&
        document.querySelectorAll(".tr-queue .risk-badge.risk-none").length === 2 && /Already encoded/.test(document.querySelector(".tr-queue").textContent); }),
      "R11 re-uploading (even a renamed copy) on the written tracker: 'Already encoded' on its row, nothing to generate");
    await saveMapIfAsked(page);
    ok(await page.evaluate(() => document.querySelector('[data-traction="q-generate"]').disabled), "R12 Generate disabled — nothing would be overwritten");
    ok(page._errs.length === 0, "R zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));
    await page.close();
  }

  /* ---- F: eval only, no tracker ---- */
  page = await fresh(browser);
  await upload(page, '[data-trpill="eval"]', evf);
  await saveMapIfAsked(page);
  await page.click('[data-traction="generate"]'); await page.waitForTimeout(400);
  ok(await page.evaluate(() => !!document.querySelector("#tr-report-session") && !document.querySelector('[data-traction="writeback"]') &&
      /no tracker loaded/.test(document.querySelector("#tr-detail").textContent)), "F1 eval-only: feedback report, nothing to write");
  await upload(page, '[data-trpill="tracker"]', trk);
  ok(await page.evaluate(() => QBR.trainers._state.queue.length === 1 && !!document.querySelector("[data-trq-target]") && !QBR.trainers._state.eval),
    "F2 loading a tracker afterwards moves the eval to the list for matching");
  ok(page._errs.length === 0, "F zero page errors" + (page._errs.length ? ": " + page._errs.join(" | ") : ""));

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e.message); process.exit(2); });
