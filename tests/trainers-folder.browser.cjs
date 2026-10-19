/* v1.39.0 — Trainers tab reads its files from a linked folder + GDAP mapping as .xlsx (browser, file://).
 * Usage: PW=<dir with node_modules/playwright> node tests/trainers-folder.browser.cjs
 * Synthetic workbooks only. The folder API (File System Access) is mocked with in-memory handles.
 * A: tab opens with a granted folder -> tracker (saves in place) + eval exports load on their own;
 *    other workbooks (tenant tracker, GDAP, a dated report that isn't an eval) are left alone
 * B: "Load from folder" again -> nothing duplicated
 * C: folder not allowed yet -> no auto-read; the button asks (ask:true) and loads
 * D: GDAP.xlsx (Name / Microsoft ID) uploaded with the workbooks -> gdap pill green + lookup;
 *    a .csv in the list explains why it can't be linked (checkbox disabled) */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

const b64 = (sheets) => {
  const wb = XLSX.utils.book_new();
  Object.keys(sheets).forEach(n => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheets[n]), n));
  return XLSX.write(wb, { bookType: "xlsx", type: "base64" });
};
const TRACKER = b64({ trial: [
  ["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score"],
  [46100, "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""],
  [46101, "SANTOS, Maria", "MS-103A", "Workshop P1", "Hill Crest College of Carmel", 0.54, 0.62, ""]] });
const evalRows = (school, seed) => {
  const head = ["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic",
    "Communication and presentation skills", "Ability to engage and connect with participants",
    "Overall, how satisfied were you with the training?",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
    "Why did you give this rating?"];
  const r = [head];
  for (let i = 0; i < 5; i++) r.push([i + seed, 46100.4 + i / 1440, school, "Strongly Agree", "Agree", "Agree", "Very satisfied", [10, 9, 8, 10, 9][i], "ok"]);
  return { Sheet1: r };
};
const EVAL1 = b64(evalRows("HCCCI", 1));
const EVAL2 = b64(evalRows("Hill Crest College of Carmel", 50));
const OTHER = b64({ RISKY_USERS_AND_DOMAIN: [["Month", "School", "Total Risky Users"], ["Jan", "X", 1]] });
const REPORT = b64({ Sheet1: [["Date", "Amount"], ["2026-01-01", 5]] });
const GDAPX = b64({ Info: [["GDAP export"]], GDAP: [["Granular admin relationships"], [], ["Name", "Microsoft ID", "Status"],
  ["Alpha Test School", "aaaaaaaa-1111-2222-3333-444444444444", "Active"], ["Beta Academy", "bbbbbbbb-1111-2222-3333-444444444444", "Active"]] });

const FOLDER = [
  { name: "ETG_Training_Tracker_2026.xlsx", path: [], data: TRACKER },
  { name: "05072026 - AM - HCCCI(1-44).xlsx", path: ["Evals"], data: EVAL1 },
  { name: "Feedback_-_01022026_-_PM_-_HCC_1-5.xlsx", path: ["Evals"], data: EVAL2 },
  { name: "ALL TENANT AUTOMATED TRACKER.xlsx", path: [], data: OTHER },
  { name: "GDAP.xlsx", path: [], data: GDAPX },
  { name: "Sales report 01012026.xlsx", path: [], data: REPORT }
];

async function fresh(browser, state) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page._errs = [];
  page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForSelector("#app-sidebar .sb-nav .tab-btn", { state: "attached" });
  await page.waitForTimeout(400);
  await page.evaluate(({ files, state }) => {
    const toFile = f => { const bin = atob(f.data), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
      return new File([u], f.name, { lastModified: 1767225600000, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }); };
    window.__fo = { state, asks: 0, reads: [] };
    QBR.fsFolderSupported = () => true;
    QBR.fsFolderList = () => [{ id: "d1", name: "Training", state: window.__fo.state }];
    QBR.fsFolderBooks = async (id, o) => {
      if (window.__fo.state !== "granted") { if (!(o && o.ask)) return null; window.__fo.asks++; window.__fo.state = "granted"; }
      return files.map(f => ({ name: f.name, path: f.path, rel: f.path.concat(f.name), size: 9000, lastModified: 1767225600000,
        handle: { kind: "file", name: f.name, getFile: async () => { window.__fo.reads.push(f.name); return toFile(f); },
                  queryPermission: async () => "granted", createWritable: async () => ({ write: async () => {}, close: async () => {} }) } }));
    };
    document.getElementById("app-body").classList.remove("d-none");
    window.showOpenFilePicker = undefined;
  }, { files: FOLDER, state });
  return page;
}
const status = p => p.evaluate(() => (document.getElementById("tr-status") || {}).textContent || "");

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});

  /* ---- A: auto-read when the tab opens ---- */
  let page = await fresh(browser, "granted");
  await page.click('[data-tab="dash-trainers"]');
  await page.waitForFunction(() => QBR.trainers._state.tracker && QBR.trainers._state.queue.length >= 2, null, { timeout: 8000 }).catch(() => {});
  const A = await page.evaluate(() => { const S = QBR.trainers._state; return { t: S.tracker && S.tracker.name, h: !!(S.tracker && S.tracker.handle), q: S.queue.map(q => q.name).sort(), reads: window.__fo.reads.slice() }; });
  ok(A.t === "ETG_Training_Tracker_2026.xlsx", "tracker found in the linked folder and loaded");
  ok(A.h, "tracker keeps its folder file handle (saves in place)");
  ok(A.q.length === 2 && A.q[0].indexOf("05072026") === 0, "both eval exports (incl. subfolder) loaded: " + A.q.join(", "));
  ok(A.reads.indexOf("ALL TENANT AUTOMATED TRACKER.xlsx") < 0 && A.reads.indexOf("GDAP.xlsx") < 0, "non-training workbooks are not opened");
  ok(!A.q.some(n => /Sales report/.test(n)), "a dated workbook that isn't an MS Forms eval is left out");
  ok(/Training: tracker ETG_Training_Tracker_2026\.xlsx \+ 2 eval exports loaded/.test(await status(page)), "status says what came from the folder: " + await status(page));
  ok(await page.$('[data-traction="folder"]') && /Load from folder/.test(await page.textContent('[data-traction="folder"]')), "\"📁 Load from folder\" button shown");

  /* ---- B: again -> no duplicates ---- */
  await page.click('[data-traction="folder"]'); await page.waitForTimeout(1200);
  ok(await page.evaluate(() => QBR.trainers._state.queue.length) === 2, "loading the folder again adds nothing twice");
  ok(/no new tracker or eval exports/.test(await status(page)), "…and says so");
  ok(!page._errs.length, "no page errors (A/B) " + page._errs.join(" | "));
  await page.close();

  /* ---- C: not allowed yet -> button asks ---- */
  page = await fresh(browser, "prompt");
  await page.click('[data-tab="dash-trainers"]'); await page.waitForTimeout(900);
  ok(await page.evaluate(() => !QBR.trainers._state.tracker && window.__fo.asks === 0), "no auto-read (and no prompt) while the folder isn't allowed");
  await page.click('[data-traction="folder"]');
  await page.waitForFunction(() => QBR.trainers._state.tracker && QBR.trainers._state.queue.length >= 2, null, { timeout: 8000 }).catch(() => {});
  ok(await page.evaluate(() => window.__fo.asks === 1 && !!QBR.trainers._state.tracker && QBR.trainers._state.queue.length === 2), "button asks once for the folder, then loads tracker + evals");
  ok(!page._errs.length, "no page errors (C) " + page._errs.join(" | "));
  await page.close();

  /* ---- D: GDAP.xlsx ---- */
  page = await fresh(browser, "prompt");
  await page.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); QBR.gdapClear(); } catch (e) {} });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "gdapx-"));
  const gx = path.join(dir, "GDAP.xlsx"), gc = path.join(dir, "GranularAdministerRelationship.csv"), fx = path.join(ROOT, "tests", "fixture.xlsx");
  fs.writeFileSync(gx, Buffer.from(GDAPX, "base64"));
  fs.writeFileSync(gc, "Name,Microsoft ID\r\nGamma School,cccccccc-1111-2222-3333-444444444444\r\n");
  await page.setInputFiles("#file-input", [fx, gx]); await page.waitForTimeout(2500);
  ok(await page.evaluate(() => { const c = document.querySelector("[data-gdap-chip]"); return !!c && /success/.test(c.className) && /GDAP\.xlsx/.test(c.title); }), "GDAP.xlsx turns the gdap pill green (tooltip names .xlsx)");
  ok(await page.evaluate(() => { const a = QBR.gdapLookup("alpha test school"), b = QBR.gdapLookup("Beta"); return a && a.tenantId.startsWith("aaaaaaaa") && b && b.tenantId.startsWith("bbbbbbbb"); }), "tenant IDs read from the GDAP sheet (header on row 3)");
  ok(/accept=".xlsx,.xlsm,.csv"/.test(fs.readFileSync(path.join(ROOT, "js/scripts.js"), "utf8")), "Scripts GDAP picker accepts .xlsx");
  await page.evaluate(() => { const box = document.getElementById("loaded-chips"); if (box) box.classList.remove("d-none"); });
  const cb = await page.evaluate(() => [...document.querySelectorAll(".loaded-file-edit")].map(c => ({ n: c.dataset.fname, d: c.disabled, t: c.title })));
  ok(cb.some(c => c.n === "GDAP.xlsx"), "GDAP.xlsx is listed with a link checkbox");
  // legacy CSV path still works; its checkbox explains why it can't be linked
  await page.evaluate(async (txt) => {
    const dt = new DataTransfer(); dt.items.add(new File([txt], "GranularAdministerRelationship.csv", { type: "text/csv" }));
    const inp = document.getElementById("file-input"); inp.files = dt.files; inp.dispatchEvent(new Event("change", { bubbles: true }));
  }, fs.readFileSync(gc, "utf8"));
  await page.waitForTimeout(2500);
  ok(await page.evaluate(() => { const g = QBR.gdapLookup("Gamma School"); return g && g.tenantId.startsWith("cccccccc"); }), "legacy GDAP .csv still parsed");
  const csvBox = await page.evaluate(() => { const c = [...document.querySelectorAll(".loaded-file-edit")].filter(x => /\.csv$/i.test(x.dataset.fname))[0]; return c ? { d: c.disabled, t: c.title } : null; });
  ok(csvBox && csvBox.d && /save it as \.xlsx/.test(csvBox.t), "CSV row: checkbox disabled, tooltip says to save it as .xlsx");
  ok(!page._errs.length, "no page errors (D) " + page._errs.join(" | "));
  await page.close();

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
