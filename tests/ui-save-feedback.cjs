// Save feedback + parsed-workbook cache (2026-10-28). Dummy/sample data only.
//  A. Risky sign-ins editor: Save while the linked file is open in Excel (locked) → a visible
//     "Not saved — … open in Excel" warning that SURVIVES the panel re-render; edits kept;
//     after "closing Excel" the next Save writes and shows "Saved ✓".
//  B. Every save outcome has its own message (QBR.saveResultText); Inventory Save shows it too.
//  C. Guided audit "Save all" summary names the locked file's outcome.
//  D. Parsed-workbook cache: a reload re-reads only files whose bytes changed; the model is
//     identical with the cache on and off; a merged save (file changed in Excel) re-reads
//     only that file; the in-place patcher evicts its workbook.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-save-feedback.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const S = require(path.join(ROOT, "js/xlsx-surgical.js"));
const AUD = path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx"), INV = path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"),
  FIX = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const b64 = u8 => Buffer.from(u8).toString("base64");

// Mock Direct-save handle for `file` (lockable), linked for its kinds.
async function link(p, file) {
  await p.evaluate(({ name, data }) => {
    const bin = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    window.__H = window.__H || {};
    const T = window.__H[name] = { bytes: bin, mtime: 1700000000000, writes: 0, lock: false };
    const handle = { name, kind: "file", queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => new File([T.bytes], name, { lastModified: T.mtime }),
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), abort: async () => {},
        close: async () => { if (T.lock) throw new DOMException("The process cannot access the file", "NoModificationAllowedError"); T.bytes = parts[0]; T.mtime += 1000; T.writes++; } }; } };
    const f0 = APP.files.find(f => f.name === name).blob, fp = QBR.fpOf(name, f0);
    QBR._fsLinks = (QBR._fsLinks || []).filter(l => l.name !== name).concat([{ fp, kinds: [...((QBR._kindByFp || {})[fp] || [])], name, size: bin.length, lastModified: T.mtime, hash: QBR.fsHashOf(bin), handle }]);
  }, { name: path.basename(file), data: b64(fs.readFileSync(file)) });
}
async function boot(b, files) {
  const p = await b.newPage({ viewport: { width: 1400, height: 900 } }); const errs = [];
  p.on("pageerror", e => errs.push(e.message)); p.on("dialog", d => d.dismiss());
  await p.goto(APP); await p.waitForTimeout(400);
  await p.evaluate(() => { localStorage.clear(); localStorage.setItem("qbr-remember", "0"); });
  await p.setInputFiles("#file-input", files); await p.waitForTimeout(3000);
  return { p, errs };
}
const badge = (p, kind) => p.evaluate(k => Array.from(document.querySelectorAll('[data-unsaved="' + k + '"]')).map(e => ({ t: e.textContent.trim(), warn: e.classList.contains("qbr-note-warn") })), kind);
// Stable digest of the parsed model (functions dropped, cycles cut).
const digest = p => p.evaluate(() => {
  const seen = new WeakSet();
  return JSON.stringify(APP.model, (k, v) => {
    if (typeof v === "function") return undefined;
    if (v && typeof v === "object") { if (seen.has(v)) return "[cycle]"; seen.add(v); }
    if (v instanceof Set) return Array.from(v);
    return v;
  });
});

(async () => {
  const b = await chromium.launch();

  console.log("== A. Risky editor: Save while open in Excel");
  let { p, errs } = await boot(b, [AUD, FIX]);
  await link(p, AUD);
  const audName = path.basename(AUD);
  await p.evaluate(() => { goToTab("dash-audit"); const a = QBR._audit, ms = Object.values(a.months)[0]; QBR.auditUpdateCell(a.fp, ms.sheet, ms.rows[0].r, "risky", "77"); });
  await p.evaluate(n => { __H[n].lock = true; }, audName);
  await p.click("#audit-save"); await p.waitForTimeout(800);
  let bd = await badge(p, "audit");
  ok(bd.length && /^Not saved — ".+" is open in Excel\. Close it there, then click Save again\. Your edits are kept\.$/.test(bd[0].t) && bd[0].warn, "locked save shows the 'open in Excel' warning (" + (bd[0] && bd[0].t) + ")");
  await p.evaluate(() => renderAll()); await p.waitForTimeout(200);
  bd = await badge(p, "audit");
  ok(bd.length && /open in Excel/.test(bd[0].t), "warning survives another full re-render");
  ok(await p.evaluate(n => __H[n].writes, audName) === 0 && await p.evaluate(() => QBR.journalUnsavedCount("audit")) > 0, "nothing written; edit still pending");
  await p.evaluate(n => { __H[n].lock = false; }, audName);
  await p.click("#audit-save"); await p.waitForTimeout(1200);
  bd = await badge(p, "audit");
  ok(bd.length && /^Saved ✓ SAMPLE_AUDIT_RICH\.xlsx \(1 change\)/.test(bd[0].t) && !bd[0].warn, "after closing Excel: Saved ✓ shown (" + (bd[0] && bd[0].t) + ")");
  ok(await p.evaluate(n => __H[n].writes, audName) === 1 && await p.evaluate(() => QBR.journalUnsavedCount("audit")) === 0, "written once; journal cleared");

  console.log("== B. Every outcome has a message");
  const modes = await p.evaluate(() => ["file", "no-changes", "reloaded", "download", "download-fallback", "download-unsafe", "download-changed", "locked", "write-failed", "busy", "merge-cancelled", "denied", "error", "weird"]
    .map(m => { const t = QBR.saveResultText({ mode: m, name: "X.xlsx", applied: 2, filename: "X-edited.xlsx", reason: "R" }); return [m, t && t.text, t && !!t.warn, t && t.short]; }));
  ok(modes.every(([, t, , s]) => t && s), "text + short summary for every mode");
  ok(modes.filter(([, , w]) => w).map(([m]) => m).join(",") === "download-unsafe,download-changed,locked,write-failed,busy,merge-cancelled,denied,error,weird", "warning styling exactly for the not-saved outcomes");
  ok(await p.evaluate(() => QBR.saveResultText({ mode: "no-data" }) === null), "no-data → handled by the caller's alert");
  await p.close();
  ({ p, errs } = await boot(b, [INV]));
  await link(p, INV);
  await p.evaluate(n => { __H[n].lock = true; goToTab("dash-inventory"); const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "LOCK TEST" }); }, path.basename(INV));
  await p.evaluate(() => QBR.invSave()); await p.waitForTimeout(500);
  bd = await badge(p, "assets");
  ok(bd.length && /open in Excel/.test(bd[0].t) && bd[0].warn, "Inventory Save shows the same warning (" + (bd[0] && bd[0].t) + ")");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await p.close();

  console.log("== C. Guided audit Save all");
  ({ p, errs } = await boot(b, [AUD]));
  await link(p, AUD);
  await p.evaluate(n => { __H[n].lock = true; goToTab("dash-audit"); const a = QBR._audit, ms = Object.values(a.months)[0]; QBR.auditUpdateCell(a.fp, ms.sheet, ms.rows[0].r, "risky", "5"); }, audName);
  await p.evaluate(() => QBR.wizardSaveAll()); await p.waitForTimeout(500);
  bd = await badge(p, "audit");
  ok(bd.length && /^Not saved \(all\) — Risky: open in Excel — close it and save again\. Your unsaved edits are kept\.$/.test(bd[0].t) && bd[0].warn, "Save all summary names the locked file (" + (bd[0] && bd[0].t) + ")");
  await p.close();

  console.log("== D. Parsed-workbook cache");
  const FILES = [AUD, INV, FIX];   // one inventory workbook, so the edit is journaled against INV
  ({ p, errs } = await boot(b, FILES));
  const st = () => p.evaluate(() => Object.assign({}, QBR._wbStats, { size: QBR.wbCacheSize() }));
  const s0 = await st();
  const dg1 = await digest(p);
  // reload: same bytes → zero reads
  await p.evaluate(async () => { QBR._wbStats.reads = 0; QBR._wbStats.hits = 0; await loadItems(APP.files.slice(), {}); });
  let s1 = await st();
  ok(s1.reads === 0 && s1.hits > 0, "reload of unchanged files: 0 workbook reads, " + s1.hits + " cache hits");
  ok(await digest(p) === dg1, "model identical after a cached reload");
  ok(s0.size === FILES.length * 2, "cache holds one plain + one styled read per file (" + s0.size + ")");
  // cache off → same model
  await p.evaluate(async () => { localStorage.setItem("qbr-wb-cache", "off"); QBR._wbStats.reads = 0; await loadItems(APP.files.slice(), {}); });
  s1 = await st();
  ok(s1.reads === FILES.length * 6, "kill switch: every pass reads again (" + s1.reads + ")");
  ok(await digest(p) === dg1, "model identical with the cache off");
  await p.evaluate(() => localStorage.removeItem("qbr-wb-cache"));
  // merged save: Excel changed the inventory file → save merges + reloads; only that file is re-read
  await link(p, INV);
  const inv0 = new Uint8Array(fs.readFileSync(INV));
  const rd = u => XLSX.read(u, { type: "array", cellStyles: true });
  const changed = (await S.surgicalSave(inv0, rd(inv0), (w => { w.Sheets["02 DEVICES"].C3 = { t: "s", v: "EXCEL-MODEL" }; return w; })(rd(inv0)))).bytes;
  await p.evaluate(async ({ n, d }) => {
    await loadItems(APP.files.slice(), {});   // warm the cache for the current set
    const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "DASH SCHOOL" });
    const T = __H[n]; T.bytes = Uint8Array.from(atob(d), c => c.charCodeAt(0)); T.mtime += 1000;
    QBR._wbStats.reads = 0; QBR._wbStats.hits = 0;
  }, { n: path.basename(INV), d: b64(changed) });
  const r = await p.evaluate(async () => { const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, merged: !!r.merged, stats: Object.assign({}, QBR._wbStats) }; });
  ok(r.mode === "file" && r.merged, "merged save into the file changed in Excel (" + r.mode + ")");
  ok(r.stats.reads === 2, "reload after the merged save re-read only the saved file (" + r.stats.reads + " reads, " + r.stats.hits + " hits; " + FILES.length + " files loaded)");
  ok(await p.evaluate(() => APP.model.inventory.assets.some(a => a.model === "EXCEL-MODEL") && APP.model.inventory.assets.some(a => a.client === "DASH SCHOOL")), "dashboard shows Excel's change AND the dashboard edit after the reload");
  ok((await st()).size === FILES.length * 2, "stale entries pruned (cache size " + (await st()).size + ")");
  // in-place patcher evicts the workbook it patches
  const ev = await p.evaluate(() => {
    const a = QBR._audit, wb = QBR._origWb[a.fp], before = QBR.wbCacheSize();
    const ms = Object.values(a.months)[0]; QBR.auditUpdateCell(a.fp, ms.sheet, ms.rows[0].r, "risky", "9");
    QBR.patchWorkbookFromJournal("audit", wb, a.fp);
    return { before, after: QBR.wbCacheSize() };
  });
  ok(ev.after === ev.before - 1, "patchWorkbookFromJournal evicts the patched workbook (" + ev.before + " → " + ev.after + ")");
  await p.evaluate(async () => { QBR._wbStats.reads = 0; await loadItems(APP.files.slice(), {}); });
  ok((await st()).reads === 1, "next load re-reads that file instead of reusing the patched object");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await p.close();
  await b.close();
  console.log("RESULT: " + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL " + (e && e.stack || e)); process.exit(1); });
