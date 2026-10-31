// Stored-XSS canary (audit 2026-10-31 C1): every text cell of the synthetic fixtures gets an HTML payload
// (double- and single-quote breakouts + <img onerror>), the poisoned workbooks are loaded into the real app,
// and every page, drill-down, School 360, the THI pane, the Executive report, a filter change, Asset/Ticket 360,
// the file list and the audit editor are rendered. Spreadsheet content is untrusted: nothing may execute or
// become an element. Synthetic data only. Usage: PW=<dir with node_modules/playwright> node tests/ui-xss-canary.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x).slice(0, 500) + "]" : "")); };

const SKIP = /month|quarter|year|date|^\s*sy\b|period|day|time/i;   // keep the time keys parseable
const P = tag => `"><img src=x data-xss="${tag}" onerror="window.__xss&&__xss('${tag}')">'><img src=x data-xss="${tag}s" onerror=__xss&&__xss('${tag}s')>`;
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-xss-"));
let poisoned = 0;
function poison(src) {
  const wb = XLSX.read(fs.readFileSync(path.join(__dirname, src)), { cellStyles: true });
  wb.SheetNames.forEach(sn => {
    const ws = wb.Sheets[sn]; if (!ws["!ref"]) return;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
    const h = rows.findIndex(r => Array.isArray(r) && r.filter(v => v != null && String(v).trim() !== "").length >= 2);
    if (h < 0) return;
    const range = XLSX.utils.decode_range(ws["!ref"]), hdr = rows[h];
    for (let c = 0; c < hdr.length; c++) {
      if (!String(hdr[c] || "").trim() || SKIP.test(String(hdr[c]))) continue;
      for (let r = h + 1; r < rows.length; r++) {
        const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c }), cell = ws[addr];
        if (!cell || cell.t !== "s" || !String(cell.v).trim()) continue;
        cell.v = String(cell.v) + P(sn.slice(0, 6) + ":" + addr); delete cell.w; delete cell.h; delete cell.r; poisoned++;
      }
    }
  });
  const out = path.join(TMP, src.replace(/\.xlsx$/, "_POISON.xlsx"));
  fs.writeFileSync(out, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return out;
}
const FILES = ["fixture.xlsx", "SAMPLE_Lenovo_Inventory_SYNTH.xlsx", "SAMPLE_AUDIT_RICH.xlsx"].map(poison);

(async () => {
  ok(poisoned > 1000, "poisoned " + poisoned + " text cells");
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1440, height: 900 } })).newPage();
  const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("dialog", d => d.dismiss().catch(() => {}));
  await p.addInitScript(() => { window.__hits = []; window.__xss = t => window.__hits.push(t + "@" + (window.__where || "?")); window.print = () => {}; window.open = () => null; });
  await p.goto("file://" + path.join(APPDIR, "index.html"));
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", FILES);
  await p.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 90000 });
  await p.waitForTimeout(2000);
  const nodes = new Map();
  const note = async where => { await p.waitForTimeout(350); for (const x of await p.evaluate(w => { window.__where = w; return [...document.querySelectorAll("[data-xss]")].map(n => n.getAttribute("data-xss") + " in #" + ((n.closest("[id]") || {}).id || "?")); }, where)) if (!nodes.has(x)) nodes.set(x, where); };
  const src = await p.evaluate(() => APP.model.sources);
  ok(src.risky && src.security && src.storage && src.usage && src.canva, "poisoned workbooks still load (core sources found)", src);

  await p.evaluate(() => document.body.classList.add("show-admin"));
  const tabs = await p.$$eval("#app-sidebar .sb-item[data-tab]", x => x.map(y => y.dataset.tab));
  for (const t of tabs) {
    await p.evaluate(t => { window.__where = t; document.querySelector('.sb-item[data-tab="' + t + '"]').click(); }, t); await note(t);
    await p.evaluate(() => document.querySelectorAll(".dash-panel:not(.d-none) button").forEach(bt => { if (/show all|more|expand|per-tenant|details?/i.test(bt.textContent || "") && bt.offsetParent) try { bt.click(); } catch (e) {} }));
    await note(t + " (expanded)");
  }
  for (let i = 0; i < 3; i++) { await p.evaluate(i => { window.__where = "THI pane"; const r = document.querySelectorAll(".thi-row")[i]; r && r.click(); }, i); await note("THI pane"); }
  for (const k of await p.evaluate(() => Object.keys(APP.model.master || {}).slice(0, 8))) { await p.evaluate(k => { window.__where = "School 360"; openSchool360(k); }, k); await note("School 360"); }
  await p.evaluate(() => { window.__where = "filters"; const o = (APP._orgs || [])[1], s = (APP._schools || [])[1]; if (o) APP.filters.org = o; if (s) APP.filters.school = s; renderAll(); }); await note("filters");
  await p.evaluate(() => { APP.filters.org = "ALL"; APP.filters.school = "ALL"; renderAll(); });
  const ids = await p.evaluate(() => { const inv = typeof invModel === "function" ? invModel() : null; return inv ? { a: inv.assets.slice(0, 3).map(a => a.sn), t: (inv.tickets || []).slice(0, 3).map(t => t.no || t.tno || t.id) } : { a: [], t: [] }; });
  for (const sn of ids.a) { await p.evaluate(sn => { window.__where = "Asset 360"; location.hash = "#asset/" + encodeURIComponent(sn); }, sn); await note("Asset 360"); }
  for (const tn of ids.t) { await p.evaluate(tn => { window.__where = "Ticket 360"; location.hash = "#ticket/" + encodeURIComponent(tn); }, tn); await note("Ticket 360"); }
  await p.evaluate(() => { window.__where = "file list"; const u = document.getElementById("upload-status"); u && u.click(); }); await note("file list");
  await p.evaluate(() => { window.__where = "audit editor"; document.querySelector('.sb-item[data-tab="dash-audit"]') && document.querySelector('.sb-item[data-tab="dash-audit"]').click(); const bt = document.querySelector("[data-audit-edit]"); bt && bt.click(); }); await note("audit editor");
  await p.waitForTimeout(800);
  const hits = await p.evaluate(() => window.__hits);
  ok(tabs.length >= 15, "rendered " + tabs.length + " pages + drill-downs");
  ok(hits.length === 0, "no payload executed", hits.slice(0, 8));
  ok(nodes.size === 0, "no payload became an element", [...nodes].slice(0, 8));
  ok(errs.length === 0, "no page errors", errs.slice(0, 3));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
