// MOS tab (v1.37.0), headless Chromium. Synthetic score reports only (tests/mos-fixture).
// Usage: PW=<dir with node_modules/playwright> node tests/ui-mos.cjs
// Runs twice: over HTTP and from file:// (the offline copy — pdf.js must work without a Worker).
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), http = require("http"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const FX = path.join(__dirname, "mos-fixture");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const MIME = { ".html": "text/html", ".js": "text/javascript", ".json": "application/json", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".gif": "image/gif" };
function serve() {
  return new Promise(resolve => {
    const srv = http.createServer((req, res) => {
      let u = decodeURIComponent(req.url.split("?")[0]); if (u === "/") u = "/index.html";
      const f = path.join(ROOT, u);
      fs.readFile(f, (e, d) => { if (e) { res.writeHead(404); res.end(); return; } res.writeHead(200, { "Content-Type": MIME[path.extname(f)] || "application/octet-stream" }); res.end(d); });
    });
    srv.listen(0, "127.0.0.1", () => resolve(srv));
  });
}
function pdfs(dir) { let o = []; for (const f of fs.readdirSync(dir)) { const p = path.join(dir, f); if (fs.statSync(p).isDirectory()) o = o.concat(pdfs(p)); else if (/\.pdf$/i.test(f)) o.push(p); } return o.sort(); }
function rosterBuffer() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["SUMMARY"]]), "SUMMARY");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["No.", "", "", "NAME", "Email"], ["1", "", "2600001", "DEMO, Alex, Cruz", "a.demo@school.example"], ["2", "", "2600002", "SAMPLE, Bea, Lopez", "bea.sample@example.com"]]), "11A");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["NAME", "Email"], ["TEST, Carlos, Reyes", ""], ["TEST, Carl, Reyes", ""]]), "11B");
  return Buffer.from(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}
const kpiVal = (p, label) => p.evaluate(l => { const k = [...document.querySelectorAll("#mos-kpi .kpi")].find(x => x.querySelector(".kpi-lbl").textContent === l); return k ? k.querySelector(".kpi-val").textContent : null; }, label);
const vis = (p, sel) => p.evaluate(s => { const el = document.querySelector(s); return !!el && el.getClientRects().length > 0 && getComputedStyle(el).display !== "none"; }, sel);

async function run(b, url, label, full) {
  console.log("== MOS tab · " + label);
  const ctx = await b.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  await p.goto(url); await p.waitForTimeout(500);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx")]); await p.waitForTimeout(2000);
  ok(await vis(p, '[data-tab="dash-mos"]') && /Certification/.test(await p.textContent("#sbg-cert")), "Certification › MOS results in the sidebar");
  await p.click('[data-tab="dash-mos"]'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-mos") && await vis(p, "#mos-empty"), "MOS tab opens with the how-it-works card");
  ok(!(await vis(p, ".filter-bar")), "shared QBR filter bar hidden on the MOS tab");
  ok(await p.evaluate(() => !window.pdfjsLib), "pdf.js not loaded until the first PDF is added");

  // PDFs one by one with the typed school (no folder path) for Demo, plus the Sample Academy files
  await p.fill("#mos-school-in", "Demo High School");
  await p.setInputFiles("#mos-pdf", pdfs(path.join(FX, "01 - Demo High School")));
  await p.waitForFunction(() => QBR.mos._state().model && !QBR.mos._state().busy, null, { timeout: 30000 });
  ok(await p.evaluate(() => QBR.mos._state().model.rows.every(r => r.school === "Demo High School")), "typed school applied to files picked one by one");
  await p.fill("#mos-school-in", "Sample Academy");
  await p.setInputFiles("#mos-pdf", pdfs(path.join(FX, "02 - Sample Academy")));
  await p.waitForFunction(() => !QBR.mos._state().busy && QBR.mos._state().files.length === 4, null, { timeout: 30000 });
  await p.waitForTimeout(300);
  ok(errs.length === 0, "no page errors while reading PDFs" + (errs.length ? " (" + errs.join("; ") + ")" : ""));
  ok(await kpiVal(p, "Attempts") === "6", "6 unique attempts (duplicate page ignored)");
  ok(await kpiVal(p, "Pass rate") === "66.7%" && await kpiVal(p, "First-attempt pass rate") === "60.0%", "pass rate 66.7% · first-attempt 60.0%");
  ok(await kpiVal(p, "Highest / lowest") === "1000 / 0", "highest / lowest score");
  ok(/4 PDFs · 7 score pages · 6 unique attempts \(1 duplicate/.test(await p.textContent("#mos-status")), "status line counts files, pages, duplicates");
  ok(await p.evaluate(() => ["mos-ch-trend", "mos-ch-group", "mos-ch-bands", "mos-ch-skills"].every(id => { const c = Chart.getChart(document.getElementById(id)); return c && c.width > 50; })), "4 charts drawn at a real size");

  // roster (school from the input)
  await p.fill("#mos-school-in", "Demo High School");
  await p.setInputFiles("#mos-roster", { name: "roster.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: rosterBuffer() });
  await p.waitForFunction(() => QBR.mos._state().roster.length === 4, null, { timeout: 10000 });
  await p.waitForTimeout(200);
  ok(await p.evaluate(() => { const r = QBR.mos._state().model.rows; const by = {}; r.forEach(x => { by[x.ref] = x; }); return by["90000001"].section === "11A" && by["90000002"].match === "email" && by["90000003"].section === "11B"; }), "roster join: name + email matches, sections assigned");

  // filters
  await p.selectOption("#mos-f-school", "Demo High School"); await p.waitForTimeout(150);
  ok(/Pass rate by section/.test(await p.textContent("#mos-ch-group-h")) && /Summary by section/.test(await p.textContent("#mos-sum-h")), "one school → charts + summary switch to sections");
  ok(await p.$$eval("#mos-f-section option", o => o.map(x => x.value).join(",")) === ",11A,11B", "section filter lists that school's sections");
  await p.selectOption("#mos-f-section", "11A"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "3", "section 11A → 3 attempts");
  await p.selectOption("#mos-f-outcome", "FAIL"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "1" && await kpiVal(p, "Failed") === "1", "fail filter");
  await p.click("#mos-reset"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "6", "reset filters");
  await p.fill("#mos-f-min", "700"); await p.fill("#mos-f-max", "900"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "2", "score range 700–900");
  await p.fill("#mos-f-min", ""); await p.fill("#mos-f-max", "");
  await p.selectOption("#mos-sort", "score-asc"); await p.waitForTimeout(100);
  ok(await p.$$eval("#mos-res tbody tr.mos-row td.num b", b => b.map(x => +x.textContent).join(",")) === "0,600,760,875,920,1000", "sort: lowest score first");
  await p.selectOption("#mos-sort", "score-desc"); await p.waitForTimeout(100);
  ok(await p.$eval("#mos-res tbody tr.mos-row td.num b", x => x.textContent) === "1000", "sort: highest score first");
  await p.selectOption("#mos-f-examno", "9001"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "4", "Exam No. 9001 includes its 9001.1 retake session");
  await p.click("#mos-reset");
  await p.selectOption("#mos-f-pmode", "year"); await p.waitForTimeout(100);
  await p.selectOption("#mos-f-pval", "2025"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "2" && await kpiVal(p, "Pass rate") === "50.0%", "yearly period: 2025 → 2 attempts, 50%");
  await p.selectOption("#mos-f-pmode", "month"); await p.waitForTimeout(100);
  ok(await p.$$eval("#mos-f-pval option", o => o.map(x => x.textContent).join("|")) === "All|Apr 2026|Mar 2026|Sep 2025", "monthly period lists months newest first");
  await p.selectOption("#mos-f-pval", "2026-03"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "3", "month Mar 2026 → 3 attempts");
  await p.click("#mos-reset");
  await p.click('[data-mos-trend="year"]'); await p.waitForTimeout(150);
  ok(await p.evaluate(() => Chart.getChart(document.getElementById("mos-ch-trend")).data.labels.join(",")) === "2025,2026", "trend toggle: yearly");
  await p.click('[data-mos-trend="month"]'); await p.waitForTimeout(150);
  ok(await p.evaluate(() => Chart.getChart(document.getElementById("mos-ch-trend")).data.labels.length) === 3, "trend toggle: monthly");
  await p.fill("#mos-f-q", "bea"); await p.waitForTimeout(150);
  ok(await kpiVal(p, "Attempts") === "2", "student search");
  await p.click("#mos-res tbody tr.mos-row"); await p.waitForTimeout(150);
  ok(await p.$$eval("#mos-res tr.mos-hist tbody tr", x => x.length) === 2, "click a student → full attempt history");
  await p.fill("#mos-f-q", ""); await p.waitForTimeout(100);

  // check & fix: Carl matched; nothing ambiguous; Sample Academy has no roster so it is not listed
  await p.click("#mos-check summary"); await p.waitForTimeout(100);
  const chk = await p.textContent("#mos-check-body");
  ok(/Every attempt is matched/.test(chk) && /Exam 9002 Certificates\.pdf/.test(chk) && /No score-report pages/.test(chk), "Check panel: all matched; certificate PDF listed as skipped");

  if (full) {
    // export
    const [dl] = await Promise.all([p.waitForEvent("download"), p.click("#mos-export")]);
    const fp = path.join(os.tmpdir(), "mos-export-test.xlsx"); await dl.saveAs(fp);
    const wb = XLSX.read(fs.readFileSync(fp), { type: "buffer" });
    const res = XLSX.utils.sheet_to_json(wb.Sheets.Results);
    ok(wb.SheetNames.join(",") === "Summary,Results,By month,By year" && res.length === 6, "export: 4 sheets, 6 result rows");
    ok(!JSON.stringify(res).includes("@") && !JSON.stringify(res).includes("Sample Street"), "export has no emails or addresses");
    fs.unlinkSync(fp);
    // privacy: nothing persisted
    const persisted = await p.evaluate(async () => {
      let s = ""; for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); s += k + "=" + localStorage.getItem(k); }
      for (let i = 0; i < sessionStorage.length; i++) { const k = sessionStorage.key(i); s += k + "=" + sessionStorage.getItem(k); }
      return s;
    });
    ok(!/Alex Cruz|alex\.demo|Bea Lopez|90000001/.test(persisted), "no student data in localStorage / sessionStorage");
    // deck + flags
    ok(await p.evaluate(() => typeof inDeck === "function" && inDeck(document.getElementById("dash-mos")) === false), "MOS is not part of the Export Deck");
    ok(await p.evaluate(() => QBR.features.get("mos") && QBR.features.get("mos").module === "certification"), "feature 'mos' registered (Certification)");
    await p.evaluate(() => { QBR._remoteFlagsTestLoad({ version: 9, flags: { mos: false } }, "live"); QBR.remoteFlagsApplyNav(); });
    await p.waitForTimeout(200);
    ok(!(await vis(p, "#sbg-cert")) && await p.evaluate(() => APP.activeTab !== "dash-mos"), "remote flag mos=false hides the Certification group and leaves the tab");
    await p.evaluate(() => QBR.mos.render());
    ok(/turned off/.test(await p.textContent("#mos-body")), "render refuses while the flag is off");
    // clear
    await p.evaluate(() => { QBR._remoteFlagsTestLoad({ version: 10, flags: {} }, "live"); QBR.remoteFlagsApplyNav(); });
    await p.click('[data-tab="dash-mos"]'); await p.waitForTimeout(200);
    p.once("dialog", d => d.accept());
    await p.click("#mos-clear"); await p.waitForTimeout(200);
    ok(await p.evaluate(() => QBR.mos._state().attempts.length === 0) && await vis(p, "#mos-empty"), "Clear empties the tab");
  }
  ok(errs.length === 0, "no page errors" + (errs.length ? " (" + errs.join("; ") + ")" : ""));
  await ctx.close();
}

(async () => {
  const srv = await serve();
  const b = await chromium.launch();
  try {
    await run(b, "http://127.0.0.1:" + srv.address().port + "/index.html", "http", true);
    await run(b, "file://" + path.join(ROOT, "index.html"), "file:// (offline copy)", false);
  } finally { await b.close(); srv.close(); }
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
