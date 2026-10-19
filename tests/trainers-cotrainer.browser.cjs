/* v1.39.0 — Trainers: co-trained sessions (one row, "NAME / NAME") + duplicate-session warning (browser, file://).
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers-cotrainer.browser.cjs   (synthetic data only)
 * A: duplicate rows (same date, school, course, start) are flagged with their Excel rows
 * B: leaderboard + summary credit each co-trainer; sessions/trainers tiles count the session once
 * C: trainer profile (filter by either name) includes the co-trained session + "co-trained" note
 * D: Add session with a co-trainer writes ONE row with "TRAINER / CO-TRAINER" */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), "trco-"));
const H = ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score", "Trainer Score"];
const ROWS = [H,
  [46276, "CEDO, Anthony / FAYLONA, Franchesca", "CV-AI1", "Reimagine Student Learning with AI-Powered Sessions", "DepEd Philippines", 0.375, 0.5, 80, 4.6],
  [46285, "CEDO, Anthony", "CV-AI1", "Reimagine Student Learning with AI-Powered Sessions", "Paco Catholic School", 0.375, 0.5, 60, 4.2],
  [45946, "AZAÑA, Lito", "MS-101A", "Onboarding", "Academia de San Vicente Ferrer", 0.375, 0.5, 58.3, 4.81],
  [45946, "AZAÑA, Lito", "MS-101A", "Onboarding", "Academia de San Vicente Ferrer", 0.375, 0.5, 58.3, 4.81]];
const TRK = path.join(DIR, "ETG_Training_Tracker_test.xlsx");
{ const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ROWS), "trial"); fs.writeFileSync(TRK, XLSX.write(wb, { bookType: "xlsx", type: "buffer" })); }

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  const errs = []; page.on("pageerror", e => errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(400);
  await page.evaluate(() => { document.getElementById("app-body").classList.remove("d-none"); window.showOpenFilePicker = undefined; });
  await page.click('[data-tab="dash-trainers"]'); await page.waitForTimeout(200);
  const [fc] = await Promise.all([page.waitForEvent("filechooser"), page.click('[data-trpill="tracker"]')]);
  await fc.setFiles(TRK); await page.waitForTimeout(700);

  /* A */
  const dw = await page.textContent(".tr-dupwarn").catch(() => "");
  ok(/Possible duplicate sessions/.test(dw) && /2025-10-16 · Academia de San Vicente Ferrer · MS-101A: rows 4 \(AZAÑA, Lito\), 5/.test(dw), "duplicate rows flagged with Excel rows: " + dw.replace(/\s+/g, " ").slice(0, 160));
  ok(!/DepEd/.test(dw), "a co-trained single row is NOT a duplicate");
  ok(/3 trainers/.test(await page.textContent('[data-trfile="tracker"]')), "tracker pill: 3 trainers (co-trainers split)");

  /* B */
  await page.click('[data-trview="board"]'); await page.waitForTimeout(300);
  const board = await page.$$eval(".tr-board tbody tr", trs => trs.map(tr => [...tr.children].slice(1, 3).map(td => td.textContent.trim())));
  ok(board.some(r => r[0] === "FAYLONA, Franchesca" && r[1] === "1") && board.some(r => r[0] === "CEDO, Anthony" && r[1] === "2"), "leaderboard credits each co-trainer: " + JSON.stringify(board));
  await page.click('[data-trview="summary"]'); await page.waitForTimeout(500);
  const tiles = await page.$$eval("#tr-report .kpi", ks => ks.map(k => k.textContent.replace(/\s+/g, " ").trim()));
  ok(tiles.some(t => /^3\s*Trainers active/i.test(t)) && tiles.some(t => /^4\s*Trainings conducted/i.test(t)), "tiles: 3 trainers, 4 trainings (co-trained session counted once)");
  const trow = await page.$$eval(".tr-trainers tbody tr", trs => trs.map(tr => tr.textContent.replace(/\s+/g, " ")));
  ok(trow.some(t => /FAYLONA, Franchesca.*1 co-trained · with CEDO, Anthony/.test(t)), "Trainers table: co-trained hint");

  /* C */
  await page.selectOption("#tr-rep-trainer", "FAYLONA, Franchesca"); await page.waitForTimeout(500);
  const prof = await page.textContent("#tr-report");
  ok(/1 co-trained session \(with CEDO, Anthony\)/.test(prof), "profile: co-trained note");
  ok(await page.$$eval("#tr-report .kpi", ks => ks.some(k => /^1\s*Trainings conducted/i.test(k.textContent.replace(/\s+/g, " ").trim()))), "profile counts the co-trained session for the second trainer");
  await page.click('[data-traction="rep-clear"]'); await page.waitForTimeout(300);

  /* D */
  await page.click('[data-trview="sessions"]'); await page.waitForTimeout(300);
  await page.click('[data-traction="add-session"]'); await page.waitForSelector("#trns-cotrainer");
  await page.fill("#trns-date", "2026-09-30"); await page.fill("#trns-trainer", "CEDO, Anthony"); await page.fill("#trns-cotrainer", "FAYLONA, Franchesca");
  await page.fill("#trns-code", "CV-AI1"); await page.fill("#trns-school", "DepEd Philippines");
  await page.click('[data-traction="save-new-session"]'); await page.waitForTimeout(300);
  page.once("dialog", d => d.accept());
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('.tr-pending [data-traction="writeback"]')]);
  let rows = null;
  if (dl) { const o = path.join(DIR, "out.xlsx"); await dl.saveAs(o); rows = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(o)).Sheets.trial, { header: 1, defval: "" }); }
  ok(rows && rows.length === 6 && rows[5][1] === "CEDO, Anthony / FAYLONA, Franchesca", "Add session with co-trainer -> one new row, Trainer = \"CEDO, Anthony / FAYLONA, Franchesca\"");
  ok(!errs.length, "no page errors " + errs.join(" | "));

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
