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

    // Certiport results export (.xlsx) + PDFs + roster
    console.log("== Certiport export");
    await p.fill("#mos-school-in", "Demo High School");
    await p.setInputFiles("#mos-pdf", pdfs(path.join(FX, "01 - Demo High School")));
    await p.waitForFunction(() => QBR.mos._state().model && !QBR.mos._state().busy, null, { timeout: 30000 });
    await p.setInputFiles("#mos-roster", [
      { name: "roster.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: rosterBuffer() },
      { name: "Results (4) Plus.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: require("./mos-export-fixture.cjs")(XLSX) }]);
    await p.waitForFunction(() => QBR.mos._state().exports.length === 1 && QBR.mos._state().roster.length === 4, null, { timeout: 10000 });
    await p.waitForTimeout(300);
    const st = await p.textContent("#mos-status");
    ok(/1 Certiport export \(4 rows\)/.test(st) && /1 matched PDF ↔ export/.test(st) && /1 in progress/.test(st), "status: export rows, PDF match, in progress  [" + st + "]");
    ok(await p.evaluate(() => QBR.mos._state().rosterFiles.length === 1), "the export is not mistaken for a roster");
    ok(await kpiVal(p, "Attempts") === "6", "4 PDF attempts + 3 new scored export rows − 1 merged = 6 scored");
    ok(await p.$$eval("#mos-f-proctor option", o => o.map(x => x.value).join(",")) === ",(none),Proctor One,Proctor Two", "proctor filter lists the proctors");
    await p.selectOption("#mos-f-proctor", "Proctor One"); await p.waitForTimeout(150);
    ok(await kpiVal(p, "Attempts") === "2", "proctor filter");
    await p.click("#mos-reset");
    await p.selectOption("#mos-groupby", "proctor"); await p.waitForTimeout(150);
    ok(/Summary by proctor/.test(await p.textContent("#mos-sum-h")) && /Pass rate by proctor/.test(await p.textContent("#mos-ch-group-h")), "group by proctor (summary + chart)");
    await p.selectOption("#mos-f-outcome", "IN PROGRESS"); await p.waitForTimeout(150);
    ok(/In progress/.test(await p.textContent("#mos-res tbody")) && await p.$$eval("#mos-res tbody tr.mos-row", x => x.length) === 1, "in-progress row shown with its own status");
    await p.click("#mos-reset");
    ok(await p.evaluate(() => { const r = QBR.mos._state().model.rows.find(x => x.ref === "90000001"); return r.source === "pdf+export" && r.proctor === "Proctor One" && r.skills.length === 5; }), "PDF + export merged by Exam reference #");
    const [dl2] = await Promise.all([p.waitForEvent("download"), p.click("#mos-export")]);
    const fp2 = path.join(os.tmpdir(), "mos-export-test2.xlsx"); await dl2.saveAs(fp2);
    const res2 = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(fp2), { type: "buffer" }).Sheets.Results);
    ok(res2.length === 7 && res2.some(r => r.Proctor === "Proctor Two" && r.Batch === "DEN-20261008-01") && !JSON.stringify(res2).includes("TEST-VOUC") && !JSON.stringify(res2).includes("2011-01-01"),
      "export: proctor + batch columns, no voucher / date of birth");
    fs.unlinkSync(fp2);

    console.log("== Exam Summary (PDF) + school profile + company logos");
    const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEElEQVR4nGNgYGD4z8DAwAAABQABDQottAAAAABJRU5ErkJggg==", "base64");
    await p.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
    await p.selectOption("#mos-f-school", "Demo High School"); await p.waitForTimeout(150);
    await p.click("#mos-profile"); await p.waitForSelector("#mos-pf-title");
    ok(await p.evaluate(() => document.activeElement && document.activeElement.id === "mos-pf-title"), "profile dialog opens with focus in the first field");
    ok(!(await p.$("#mos-pf-banner")) && await p.evaluate(() => /rakso-ct-logo\.png$/.test(document.getElementById("mos-pf-left-img").getAttribute("src"))), "no banner slot; left company logo = built-in Rakso CT");
    await p.fill("#mos-pf-title", "Demo High School of Example City"); await p.fill("#mos-pf-acr", "dhs");
    await p.setInputFiles("#mos-pf-logo", { name: "logo.png", mimeType: "image/png", buffer: PNG }); await p.waitForSelector("#mos-pf-logo-img");
    await p.setInputFiles("#mos-pf-right", { name: "partner.png", mimeType: "image/png", buffer: PNG }); await p.waitForSelector("#mos-pf-right-img");
    await p.setInputFiles("#mos-pf-logo", { name: "big.png", mimeType: "image/png", buffer: Buffer.alloc(600 * 1024, 1) }); await p.waitForTimeout(200);
    ok(/too large/.test(await p.textContent("#mos-pf-msg")) && !!(await p.$("#mos-pf-logo-img")), "oversized logo refused, previous logo kept");
    ok(await p.evaluate(() => document.getElementById("mos-pf-logoh").value === "40" && document.getElementById("mos-pf-scale").value === "85" && document.getElementById("mos-pf-align").value === "center"),
      "header size defaults: logo 40 px, results box 85%, logos centered");
    await p.$eval("#mos-pf-logoh", el => { el.value = "30"; el.dispatchEvent(new Event("input")); });
    await p.$eval("#mos-pf-scale", el => { el.value = "70"; el.dispatchEvent(new Event("input")); });
    ok(/30 px/.test(await p.textContent("#mos-pf-logoh-v")) && /70%/.test(await p.textContent("#mos-pf-scale-v")) &&
      await p.evaluate(() => document.getElementById("mos-pf-brandprev").style.getPropertyValue("--logo-h") === "18px"), "sliders update their value and the live logo preview");
    await p.click("#mos-pf-save"); await p.waitForTimeout(200);
    ok(!(await p.$(".mos-ov")), "profile saved and dialog closed");
    const prof = await p.evaluate(() => (localStorage.getItem("qbr-mos-school-profiles-v1") || "") + (localStorage.getItem("qbr-mos-branding-v1") || ""));
    ok(/Demo High School of Example City/.test(prof) && /"acronym":"DHS"/.test(prof) && /"right":"data:image\/png/.test(prof) && !/Alex|Bea|Gina|example\.com/.test(prof), "profile + company logos stored in this browser, no student data");
    ok(await p.evaluate(() => QBR.mos._state().model.rows.some(r => r.batchAuto && r.batchAuto.indexOf("DHS-") === 0)), "profile acronym used in batch codes");
    await p.click("#mos-summary"); await p.waitForFunction(() => window.__printed > 0, null, { timeout: 10000 });
    const sm = await p.evaluate(() => { const h = document.getElementById("mos-print"); if (!h) return null; const co = [...h.querySelectorAll(".mos-sum-brand img")];
      return { co: co.length, coTop: co.map(i => Math.round(i.getBoundingClientRect().top)), school: h.querySelector(".mos-sum-box .nm").textContent, logo: !!h.querySelector("img.mos-sum-logo"),
        sub: h.querySelector(".mos-sum-box .sub").textContent, counts: h.querySelector(".mos-sum-counts").textContent, rate: h.querySelector(".mos-sum-box .rate").textContent,
        rows: h.querySelectorAll(".mos-sum-tbl tbody tr").length, txt: h.textContent, title: document.title }; });
    ok(await p.evaluate(() => { const d = document.querySelector("#mos-print .mos-sum-doc"), br = document.querySelector("#mos-print .mos-sum-brand");
      return d.style.getPropertyValue("--logo-h") === "30px" && d.style.getPropertyValue("--hs") === "0.7" && br.classList.contains("mos-al-center"); }), "summary uses the saved header size and centred logos");
    ok(sm && sm.co === 2 && sm.school === "DEMO HIGH SCHOOL OF EXAMPLE CITY" && sm.logo && /^Microsoft Office Specialist — Excel, PowerPoint$/.test(sm.sub), "header: company logos row, school logo + name, exam title");
    ok(sm && sm.rows === 4 && /^4TOTAL/.test(sm.counts) && /3STUDENTSPASSED/.test(sm.counts) && /1STUDENTSFAILED/.test(sm.counts) && /1ELIGIBLEFOR RETAKE/.test(sm.counts) && sm.rate === "75%",
      "summary row: total / passed / failed / eligible for retake + passing rate (Demo High School: 4 attempts, 75%)  [" + (sm && sm.counts) + "]");
    ok(sm && /Failed · retook/.test(sm.txt) && /2nd take/.test(sm.txt), "first-time fail who retook is marked; 2nd take marked");
    ok(sm && !/@|example\.com|alexdemo01/.test(sm.txt), "no emails / Certiport IDs printed");
    ok(sm && /^MOS Exam Summary - DHS - \d{4}-\d\d-\d\d$/.test(sm.title), "document title = PDF file name  [" + (sm && sm.title) + "]");
    await p.emulateMedia({ media: "print" });
    ok(await vis(p, "#mos-print") && !(await vis(p, "#app-body")) && !(await vis(p, "#app-sidebar")), "print shows only the summary");
    const tops = await p.evaluate(() => [...document.querySelectorAll("#mos-print .mos-sum-brand img")].map(i => { const r = i.getBoundingClientRect(); return Math.round(r.top + r.height / 2); }));
    ok(tops.length === 2 && Math.abs(tops[0] - tops[1]) <= 1, "company logos vertically aligned (centres " + tops.join(" / ") + ")");
    await p.emulateMedia({ media: "screen" });
    await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(100);
    ok(await p.evaluate(() => !document.getElementById("mos-print") && !document.body.classList.contains("mos-printing") && !/MOS Exam Summary/.test(document.title)), "after printing the view is removed and the title restored");
    // a pass-only filter would print a misleading 100% → warn first
    await p.selectOption("#mos-f-outcome", "PASS"); await p.waitForTimeout(100);
    let msg = ""; p.once("dialog", d => { msg = d.message(); d.dismiss(); });
    const before = await p.evaluate(() => window.__printed);
    await p.click("#mos-summary"); await p.waitForTimeout(300);
    ok(/leave out 1 attempt/.test(msg) && await p.evaluate(b => window.__printed === b, before), "Result = Pass filter → warning; cancel prints nothing");
    await p.selectOption("#mos-f-outcome", ""); await p.waitForTimeout(100);
    await p.selectOption("#mos-f-exam", "Excel"); await p.selectOption("#mos-f-proctor", ""); await p.waitForTimeout(100);
    await p.evaluate(() => { QBR.mos._state().f.section = ""; });
    const sk = await p.evaluate(() => QBR.mos._summary());
    ok(sk.areas === 5 && /Manage Worksheets and Workbooks/.test(sk.html), "one exam → skill areas I–V as columns with a legend");

    console.log("== resizable columns");
    await p.click("#mos-reset"); await p.waitForTimeout(150);
    ok(await p.evaluate(() => { const t = document.getElementById("mos-res"); return !t.classList.contains("resizable") && getComputedStyle(t).tableLayout === "auto" &&
      [...t.querySelectorAll("tbody td")].every(td => td.scrollWidth <= td.clientWidth + 1); }), "results table: automatic layout, no cell spills into its neighbour (was table-layout:fixed without grips)");
    ok(await p.$$eval("#mos-res thead th .mos-grip", g => g.length) === 11 && await p.$$eval("#mos-sum thead th .mos-grip", g => g.length) > 5, "every column has a resize grip");
    const g1 = await p.$("#mos-res thead th:nth-child(1) .mos-grip"); await g1.scrollIntoViewIfNeeded(); const gb = await g1.boundingBox();
    const w0 = await p.$eval("#mos-res thead th:nth-child(1)", th => th.getBoundingClientRect().width);
    await p.mouse.move(gb.x + 4, gb.y + gb.height / 2); await p.mouse.down(); await p.mouse.move(gb.x + 84, gb.y + gb.height / 2, { steps: 4 }); await p.mouse.up(); await p.waitForTimeout(100);
    const w1 = await p.$eval("#mos-res thead th:nth-child(1)", th => th.getBoundingClientRect().width);
    ok(w1 - w0 > 60 && await p.evaluate(() => document.getElementById("mos-res").classList.contains("mos-fixed")), "drag widens the column (" + Math.round(w0) + " → " + Math.round(w1) + " px)");
    await p.focus("#mos-res thead th:nth-child(1) .mos-grip"); await p.keyboard.press("ArrowLeft"); await p.waitForTimeout(50);
    const w2 = await p.$eval("#mos-res thead th:nth-child(1)", th => th.getBoundingClientRect().width);
    ok(Math.abs((w1 - w2) - 16) <= 2, "arrow key narrows the column by 16 px");
    await p.selectOption("#mos-sort", "name"); await p.waitForTimeout(150);
    ok(await p.evaluate(() => document.getElementById("mos-res").classList.contains("mos-fixed") && document.querySelectorAll("#mos-res .mos-grip").length === 11), "widths and grips survive a re-render (sort)");
    ok(await p.evaluate(() => /mos-res:11/.test(localStorage.getItem("qbr-mos-colw-v1") || "") && !/Alex|Bea/.test(localStorage.getItem("qbr-mos-colw-v1") || "")), "widths kept as a layout preference (no student data)");
    await p.click('[data-mos-colreset="mos-res"]'); await p.waitForTimeout(150);
    ok(await p.evaluate(() => !document.getElementById("mos-res").classList.contains("mos-fixed")), "Reset columns → automatic layout");
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
