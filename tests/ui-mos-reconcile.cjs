// MOS Reconcile view — headless Chromium, synthetic roster + Certiport export only.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-mos-reconcile.cjs   (CHROME=<path> optional)
// Runs over HTTP and from file:// (the offline copy).
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), http = require("http"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const FX = require(path.join(__dirname, "mos-reconcile-fixture.cjs"));
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
const vis = (p, sel) => p.evaluate(s => { const el = document.querySelector(s); return !!el && el.getClientRects().length > 0 && getComputedStyle(el).display !== "none"; }, sel);
const flagOf = (p, namePart) => p.evaluate(n => {
  const tr = [...document.querySelectorAll("#mos-rec-tbl tbody tr")].find(r => r.textContent.includes(n));
  return tr ? tr.textContent.replace(/\s+/g, " ") : "";
}, namePart);

async function run(b, url, label, httpMode) {
  console.log("== MOS Reconcile · " + label);
  const ctx = await b.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  p.on("dialog", d => d.accept());
  await p.goto(url); await p.waitForTimeout(500);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx")]); await p.waitForTimeout(1500);
  await p.click('[data-tab="dash-mos"]'); await p.waitForTimeout(300);

  // open Reconcile from the load-row shortcut (works with nothing loaded)
  await p.click('.mos-load [data-mos-view="reconcile"]'); await p.waitForTimeout(300);
  ok(await vis(p, "#mos-view-reconcile"), "Reconcile view opens with no data loaded");
  ok(await p.evaluate(() => !!document.querySelector("#mos-rec-paste")), "paste box present");

  // paste the roster (roster order is the master list)
  await p.fill("#mos-rec-paste", FX.rosterTSV());
  await p.click('[data-rec="paste"]'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 6), "6 roster rows in the preview");
  ok(await p.evaluate(() => QBR.mosReconcile.count() === 6), "viewbar badge count = 6 students");

  // upload the Certiport export (also loads as attempts in Results)
  const tmp = path.join(os.tmpdir(), "mos-reconcile-export.xlsx");
  fs.writeFileSync(tmp, FX.buildExport(XLSX));
  await p.setInputFiles("#mos-rec-export-file", tmp);
  await p.waitForFunction(() => window.QBR && QBR.mosReconcile._state.exportRows.length === 7, null, { timeout: 15000 });
  await p.waitForTimeout(400);
  ok(true, "export loaded: 7 voucher rows");
  ok(/FILLED/.test(await flagOf(p, "DEMO, Alex")), "Alex: FILLED from Certiport");
  ok(await p.evaluate(() => [...document.querySelectorAll("#mos-view-reconcile thead th")].some(th => th.textContent === "Exam")), "preview table has a dedicated Exam column");
  ok(await p.evaluate(() => { const tr = [...document.querySelectorAll("#mos-rec-tbl tbody tr")].find(r => r.textContent.includes("DEMO, Alex")); return tr && tr.children[4].textContent === "Excel"; }), "Alex row Exam cell = Excel");
  ok(/VERIFIED/.test(await flagOf(p, "SAMPLE, Bea")), "Bea: VERIFIED (already correct)");
  const carlos = await flagOf(p, "TEST, Carlos");
  ok(/CHANGED/.test(carlos) && /500→640/.test(carlos), "Carlos: CHANGED score 500→640");
  ok(/NO CERTIPORT RESULT/.test(await flagOf(p, "EXAMPLE, Dana")), "Dana: NO CERTIPORT RESULT");
  ok(/CONFLICT/.test(await flagOf(p, "TRIAL, Eli")), "Eli: CONFLICT (880/PASSED, no record)");
  ok(/NAME MATCH — CONFIRM/.test(await flagOf(p, "KIM, Fiona")), "Fiona: name candidate awaiting confirmation");

  // accept the name candidate
  await p.click('[data-rec="accept"]');
  await p.waitForFunction(() => document.querySelector("#mos-view-reconcile").textContent.includes("NAME MATCH — CONFIRMED"), null, { timeout: 10000 });
  ok(true, "Fiona: NAME MATCH — CONFIRMED after Accept");
  // exam filter on the preview (options derive from the roster rows' exams)
  ok(await p.evaluate(() => [...document.querySelectorAll("#mos-rec-f-exam option")].map(o => o.value).join(",") === ",Excel"), "Exam filter offers the roster's exam (Excel)");
  await p.selectOption("#mos-rec-f-exam", "Excel"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 6), "Exam filter = Excel: 6 rows");
  ok(await p.evaluate(() => /Showing 6 of 6/.test(document.querySelector("#mos-view-reconcile").textContent)), "filter note shows 6 of 6");
  await p.selectOption("#mos-rec-f-exam", ""); await p.waitForTimeout(200);
  // full filter bar: flag, search, score range, sort, reset
  await p.selectOption("#mos-rec-f-flag", "changed"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 1), "Flag filter = Changed: 1 row");
  await p.selectOption("#mos-rec-f-flag", ""); await p.waitForTimeout(200);
  await p.fill("#mos-rec-q", "fiona"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 1), "Student search 'fiona': 1 row");
  await p.fill("#mos-rec-q", ""); await p.waitForTimeout(200);
  await p.fill("#mos-rec-f-min", "900"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 1), "Score >= 900: 1 row (Bea 940)");
  await p.fill("#mos-rec-f-min", ""); await p.waitForTimeout(200);
  await p.selectOption("#mos-rec-f-sort", "score-desc"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelector("#mos-rec-tbl tbody tr").textContent.includes("SAMPLE, Bea")), "Sort = score desc: Bea (940) first");
  await p.click("#mos-rec-f-reset"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-rec-tbl tbody tr").length === 6 && document.querySelector("#mos-rec-f-sort").value === "roster"), "Reset filters restores all 6 rows in roster order");
  // sections flow into Results: filter + Exam Summary
  await p.click('.mos-viewbar [data-mos-view="results"]'); await p.waitForTimeout(400);
  const secOpts = await p.evaluate(() => [...document.querySelectorAll("#mos-f-section option")].map(o => o.value));
  ok(secOpts.includes("11A"), "Results Section filter offers 11A");
  await p.selectOption("#mos-f-section", "11A"); await p.waitForTimeout(400);
  ok(await p.evaluate(() => QBR.mos._state().model.rows.filter(r => r.section === "11A").length === 4), "4 attempts carry section 11A (3 voucher + 1 confirmed)");
  ok(await p.evaluate(() => /Section: 11A/.test(QBR.mos._summary().html)), "Exam Summary (PDF) shows Section: 11A in its filters line");
  // Retake-eligible filter: failed first takes that have not been retaken (Carlos + Hana)
  await p.selectOption("#mos-f-section", ""); await p.waitForTimeout(300);
  await p.selectOption("#mos-f-outcome", "RETAKE"); await p.waitForTimeout(400);
  ok(await p.evaluate(() => document.querySelectorAll("#mos-res tbody tr[data-sk]").length === 2), "Result filter 'Retake eligible' shows the 2 eligible students");
  await p.selectOption("#mos-f-outcome", ""); await p.waitForTimeout(300);

  // export the filled roster workbook
  await p.click('.mos-viewbar [data-mos-view="reconcile"]'); await p.waitForTimeout(300);
  const [dl] = await Promise.all([p.waitForEvent("download", { timeout: 15000 }), p.click('[data-rec="export"]')]);
  const out = path.join(os.tmpdir(), "mos-reconciled-out.xlsx");
  await dl.saveAs(out);
  const wb = XLSX.read(fs.readFileSync(out), { type: "buffer" });
  ok(JSON.stringify(wb.SheetNames) === JSON.stringify(["Filled Roster", "Paste Sheet (Roster Order)", "Not in Roster"]), "export has the 3 sheets: " + wb.SheetNames.join(" · "));
  const filled = XLSX.utils.sheet_to_json(wb.Sheets["Filled Roster"], { header: 1, defval: "" });
  ok(filled.length === 7 && filled[1][2] === "DEMO, Alex Cruz" && filled[6][2] === "KIM, Fiona", "Filled Roster keeps roster name order");
  ok(filled[3][10] === 640 && filled[3][11] === "FAILED", "Carlos filled with Certiport 640 / FAILED");
  ok(filled[6][10] === 720 && filled[6][11] === "PASSED", "Fiona filled from the confirmed candidate");
  ok(filled[5][10] === 880 && /CONFLICT/.test(String(filled[5][12])), "Eli keeps 880 with a CONFLICT flag");
  const paste = XLSX.utils.sheet_to_json(wb.Sheets["Paste Sheet (Roster Order)"], { header: 1, defval: "" });
  ok(paste.length === 7 && paste[1][0] === "DEMO, Alex Cruz" && paste[4][2] === "" , "Paste Sheet in roster order, no-result rows blank");
  const nir = XLSX.utils.sheet_to_json(wb.Sheets["Not in Roster"], { header: 1, defval: "" });
  ok(nir.length === 4, "Not in Roster holds the 3 unmatched results");

  if (httpMode) {
    // persistence: reload → the saved run resumes (the requester's choice)
    await p.reload(); await p.waitForTimeout(800);
    await p.click('[data-tab="dash-mos"]'); await p.waitForTimeout(400);
    ok(await p.evaluate(() => QBR.mosReconcile._state.blocks.length === 1 && QBR.mosReconcile._state.exportRows.length === 7), "saved run restored after reload");
    ok(await vis(p, "#mos-view-reconcile"), "MOS tab reopens on the Reconcile view when a saved run exists");
    // clear wipes it
    await p.click('[data-rec="clear"]'); await p.waitForTimeout(300);
    ok(await p.evaluate(() => localStorage.getItem("qbr-mos-reconcile-v1") === null), "Clear saved run wipes localStorage");
  }
  ok(errs.length === 0, "no page errors" + (errs.length ? " (" + errs.join("; ") + ")" : ""));
  await ctx.close();
}

(async () => {
  const srv = await serve();
  const port = srv.address().port;
  const b = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : {});
  try {
    try { await run(b, "http://127.0.0.1:" + port + "/index.html", "http", true); }
    catch (e) {
      if (!/ERR_BLOCKED_BY_LOCAL_NETWORK_ACCESS_CHECKS/.test(String(e && e.message))) throw e;
      console.log("  (http run skipped: this environment blocks loopback navigation; file:// run covers the flow)");
    }
    await run(b, "file://" + path.join(ROOT, "index.html"), "file://", false);
  } finally { await b.close(); srv.close(); }
  console.log("\nui-mos-reconcile: " + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
