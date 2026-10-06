// Journal safety (v1.30.0, js/journal.js), end-to-end in the real app (headless Chromium, file://).
// Synthetic fixtures only. Covers: per-folder namespacing (two dashboard copies in one browser),
// v1 journal migration, no 500-op cap / red badge, full localStorage → IndexedDB backup + recovery,
// beforeunload when at risk, orphaned edits (older copy of a file) re-placed by school/serial,
// parked merge edits, export/import round trip, cleared-batch backups, panel escaping.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-journal.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const INV = path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx");
const AUD = path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const wait = ms => new Promise(r => setTimeout(r, ms));

// Second dashboard folder (simulates QBR-Dashboard_dev next to the live folder).
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-journal-"));
const APP2 = path.join(TMP, "QBR-Dashboard_dev", "qbr-app");
fs.cpSync(APPDIR, APP2, { recursive: true, filter: s => !/libs[\/\\]tesseract/.test(s) });

// "Edited in Excel" copies: same name, different bytes.
function auditShifted() {   // a new school inserted at row 2 of every month sheet → all rows shift down by one
  const wb = XLSX.read(fs.readFileSync(AUD)), out = XLSX.utils.book_new();
  wb.SheetNames.forEach(n => {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null });
    if (n !== "Drop-Down" && aoa.length > 1) { const row = aoa[1].slice(); row[0] = "Inserted In Excel School"; aoa.splice(1, 0, row); }
    XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet(aoa), n);
  });
  return { name: path.basename(AUD), mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(XLSX.write(out, { bookType: "xlsx", type: "array" })) };
}
function invTouched() {     // one unrelated cell changed
  const wb = XLSX.read(fs.readFileSync(INV), { cellStyles: true }), ws = wb.Sheets["SUMMARY"] || wb.Sheets[wb.SheetNames[0]];
  ws["A20"] = { t: "s", v: "edited in Excel" };
  const r = XLSX.utils.decode_range(ws["!ref"]); if (r.e.r < 19) { r.e.r = 19; ws["!ref"] = XLSX.utils.encode_range(r); }
  return { name: path.basename(INV), mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(XLSX.write(wb, { bookType: "xlsx", type: "array" })) };
}

async function open(ctx, dir) {
  const p = await ctx.newPage(); p._errs = [];
  p.on("pageerror", e => p._errs.push(e.message));
  await p.goto("file://" + path.join(dir, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (x) {} });
  return p;
}
async function load(p, file) { await p.setInputFiles("#file-input", file); await p.waitForTimeout(2500); }

(async () => {
  const b = await chromium.launch();

  console.log("== namespacing: two dashboard folders in one browser");
  {
    const ctx = await b.newContext();
    const live = await open(ctx, APPDIR), dev = await open(ctx, APP2);
    const keys = await Promise.all([live, dev].map(p => p.evaluate(() => QBR.JOURNAL_KEY)));
    ok(keys[0] !== keys[1] && keys.every(k => /^qbr-inv-journal-v2:/.test(k)), "each folder has its own journal key");
    await load(live, INV); await load(dev, INV);
    const sn = await live.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "LIVE EDIT" }); return a.sn; });
    ok(await live.evaluate(() => QBR.journalUnsavedCount()) === 1, "live folder: 1 unsaved edit");
    await dev.reload(); await dev.waitForTimeout(400); await load(dev, INV);
    ok(await dev.evaluate(() => QBR.journalUnsavedCount()) === 0, "dev copy of the same file does NOT pick up live's edit");
    const fr = await dev.evaluate(() => QBR.journalForeign());
    ok(fr.length === 1 && fr[0].entries[0].loaded && fr[0].entries[0].count === 1, "dev lists live's edit under 'another dashboard folder'");
    ok(await dev.evaluate(() => !!document.querySelector("#qbr-journal-alert")), "dev shows an alert about it");
    await dev.evaluate(() => QBR.journalClearFp(Object.keys(QBR.journalEntries())[0]));
    ok(await live.evaluate(() => { const s = JSON.parse(localStorage.getItem(QBR.JOURNAL_KEY)); return Object.values(s.files).reduce((n, e) => n + e.ops.length, 0); }) === 1, "clearing in dev leaves live's journal intact");
    const mv = await dev.evaluate(async () => { const f = QBR.journalForeign()[0]; const r = QBR.journalMoveForeign(f.key); await r.done; return { moved: r.moved, n: QBR.journalUnsavedCount() }; });
    ok(mv.moved === 1 && mv.n === 1, "Move to this dashboard brings the edit over and replays it");
    ok(await dev.evaluate(sn => (APP.model.inventory.assets.find(a => a.sn === sn) || {}).cond, sn) === "LIVE EDIT", "moved edit visible in dev's model");
    await ctx.close();
  }

  console.log("== v1 journal migration");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    await load(p, INV);
    const r = await p.evaluate(async () => {
      const f = APP.files[0], fp = QBR.fpOf(f.name, f.blob), a = APP.model.inventory.assets[1];
      localStorage.setItem("qbr-inv-journal-v1", JSON.stringify({ [fp]: { fileName: f.name, ops: [{ id: "opLEGACY1", kind: "assets", op: "invUpdateAsset", args: [a.key, { cond: "FROM V1" }], ts: Date.now() }] },
        "other.xlsx|1|1": { fileName: "other.xlsx", ops: [{ id: "opLEGACY2", kind: "assets", op: "invSetStatus", args: ["x", "y"], ts: Date.now() }] } }));
      await loadItems(APP.files.slice(), {});
      return { cond: APP.model.inventory.assets[1].cond, n: QBR.journalUnsavedCount(), legacy: JSON.parse(localStorage.getItem("qbr-inv-journal-v1") || "null"), foreign: QBR.journalForeign().map(x => x.label) };
    });
    ok(r.cond === "FROM V1" && r.n === 1, "v1 entry for the loaded file is claimed and replayed");
    ok(r.legacy && !Object.keys(r.legacy).some(k => k.indexOf("SAMPLE_") === 0) && r.legacy["other.xlsx|1|1"], "only the claimed entry leaves the v1 key; others stay");
    ok(r.foreign.some(l => /previous dashboard version/.test(l)), "remaining v1 entries are listed as 'previous dashboard version'");
    await ctx.close();
  }

  console.log("== no silent cap; red badge at 400+");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    await load(p, INV);
    await p.evaluate(() => { const a = APP.model.inventory.assets[0]; for (let i = 0; i < 600; i++) QBR.invUpdateAsset(a.key, { cond: "C" + i }); });
    await p.evaluate(() => QBR._journalLastWrite);
    await p.reload(); await p.waitForTimeout(400); await load(p, INV);
    const r = await p.evaluate(() => ({ n: QBR.journalUnsavedCount(), cond: APP.model.inventory.assets[0].cond, red: !!document.querySelector(".qbr-unsaved-badge.bg-danger") || null }));
    ok(r.n === 600, "all 600 edits survive a reload (was capped at 500)");
    ok(r.cond === "C599", "the latest edit is the one applied");
    await p.evaluate(() => { if (typeof goToTab === "function") goToTab("dash-inventory"); QBR.persistRefreshBadge(); });
    await p.waitForTimeout(200);
    ok(await p.evaluate(() => { const el = document.querySelector('[data-unsaved="assets"] .qbr-unsaved-badge'); return !el || el.classList.contains("bg-danger"); }), "unsaved badge turns red above 400 edits");
    const old = await p.evaluate(() => { const fp = Object.keys(QBR.journalEntries())[0]; QBR.journalEntries()[fp].ops[0].ts = Date.now() - 200 * 864e5; return QBR.journalEntries()[fp].ops.length; });
    await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "one more" }); });
    ok(await p.evaluate(() => QBR.journalUnsavedCount()) === old + 1, "a 200-day-old edit is not pruned");
    await ctx.close();
  }

  console.log("== full localStorage → backup, warning, recovery");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    await load(p, INV);
    const sn = await p.evaluate(async () => {
      Storage.prototype._set = Storage.prototype.setItem;
      Storage.prototype.setItem = function () { throw new DOMException("The quota has been exceeded.", "QuotaExceededError"); };
      const a = APP.model.inventory.assets[2]; QBR.invUpdateAsset(a.key, { cond: "QUOTA EDIT" }); await QBR._journalLastWrite; return a.sn;
    });
    const h = await p.evaluate(() => QBR.journalHealth());
    ok(h.local === "failed" && h.idb === "ok", "write failure detected (localStorage failed, backup ok)");
    ok(await p.evaluate(() => { const el = document.querySelector("#qbr-journal-alert"); return !!el && el.classList.contains("jl-danger") && /storage is full/i.test(el.textContent); }), "red 'browser storage is full' alert shown");
    ok(await p.evaluate(() => QBR._journalAtRisk()), "beforeunload guard armed");
    let dialog = null; p.on("dialog", d => { dialog = d.type(); d.accept(); });
    await p.keyboard.press("a");   // beforeunload prompts need a user activation
    await p.close({ runBeforeUnload: true }); await wait(800);
    ok(dialog === "beforeunload", "closing the tab asks first");
    const q = await open(ctx, APPDIR); await load(q, INV);
    const r = await q.evaluate(sn => ({ cond: (APP.model.inventory.assets.find(a => a.sn === sn) || {}).cond, rec: QBR.journalHealth().recovered, alert: (document.querySelector("#qbr-journal-alert") || {}).textContent || "" }), sn);
    ok(r.cond === "QUOTA EDIT", "edit recovered from the IndexedDB backup and replayed after reopening");
    ok(r.rec === 1 && /Recovered 1/.test(r.alert), "'Recovered 1 unsaved edit' notice shown");
    ok(await q.evaluate(() => JSON.parse(localStorage.getItem(QBR.JOURNAL_KEY)).files && QBR.journalHealth().local === "ok"), "recovered edits written back to browser storage");
    await ctx.close();
  }

  console.log("== orphaned edits: file re-uploaded after editing in Excel");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    await load(p, AUD);
    const before = await p.evaluate(() => {
      const a = QBR._audit, ms = Object.values(a.months)[0], rec = ms.rows[1];
      QBR.auditUpdateCell(a.fp, ms.sheet, rec.r, "risky", "77");
      const e = QBR.journalEntries()[a.fp];
      // an op recorded by v1.29 (no row signature) for comparison
      e.ops.push({ id: "opNOSIG", kind: "audit", op: "auditUpdateCell", args: [a.fp, ms.sheet, ms.rows[2].r, "risky", "5"], ts: Date.now() });
      return { sheet: ms.sheet, school: rec.school, r: rec.r, sig: e.ops[0].sig };
    });
    ok(before.sig === before.school, "row edits record a signature (column A = school)");
    await load(p, auditShifted());
    const o = await p.evaluate(() => ({ orph: QBR.journalOrphans(), alert: (document.querySelector("#qbr-journal-alert") || {}).textContent || "", n: QBR.journalUnsavedCount() }));
    ok(o.orph.length === 1 && o.orph[0].count === 2 && o.orph[0].target === "SAMPLE_AUDIT_RICH.xlsx", "edits on the older copy are detected as orphaned (target = loaded file)");
    ok(/older copy/.test(o.alert) && o.n === 0, "alert names the older copy; nothing silently applied");
    await p.evaluate(() => QBR.journalPanel()); await p.waitForTimeout(150);
    ok(await p.evaluate(() => !!document.querySelector('#qbr-journal-panel [data-jp="apply"]')), "panel offers Apply to the loaded file");
    await p.click('#qbr-journal-panel [data-jp="apply"]'); await p.waitForTimeout(300);
    const r = await p.evaluate(b => {
      const a = QBR._audit, ms = Object.values(a.months).find(x => x.sheet === b.sheet), rec = ms.rows.find(x => x.school === b.school);
      const cur = QBR.journalEntries()[a.fp], left = QBR.journalOrphans();
      return { r: rec.r, risky: rec.risky, recorded: cur && cur.ops.map(x => x.args[2]), left: left.length ? left[0].count : 0,
        msg: (document.querySelector("#qbr-journal-panel .jp-msg") || {}).textContent || "" };
    }, before);
    ok(r.r === before.r + 1 && r.risky === "77", "edit re-placed on the school's new row (" + before.r + " → " + r.r + ")");
    ok(r.recorded && r.recorded.length === 1 && r.recorded[0] === before.r + 1, "re-recorded against the loaded file with the new row");
    ok(r.left === 1 && /Applied 1/.test(r.msg) && /before v1\.30/.test(r.msg), "op without a signature is kept, with the reason shown");
    await p.click('#qbr-journal-panel [data-jp="close"]');
    // Inventory (key-based) orphan
    await load(p, INV);
    const sn = await p.evaluate(() => { const a = APP.model.inventory.assets[3]; QBR.invUpdateAsset(a.key, { cond: "ORPHAN INV" }); return a.sn; });
    await load(p, invTouched());
    const inv = await p.evaluate(sn => {
      const o = QBR.journalOrphans().find(x => /Lenovo/.test(x.fileName));
      const res = QBR.journalApplyOrphan(o.fp);
      return { res, cond: (APP.model.inventory.assets.find(a => a.sn === sn) || {}).cond, n: QBR.journalUnsavedCount("assets") };
    }, sn);
    ok(inv.res.applied === 1 && inv.cond === "ORPHAN INV" && inv.n === 1, "inventory edit (by serial) applied to the re-uploaded file");
    const d = await p.evaluate(() => { const o = QBR.journalOrphans()[0]; QBR.journalDiscard(o.fp); return QBR.journalOrphans().length; });
    ok(d === 0, "Discard removes the remaining orphan");
    const tr = await p.evaluate(async () => { await QBR._journalLastTrash; return (await QBR.journalTrash()).map(t => t.reason); });
    ok(tr.includes("discarded") && tr.some(x => /^applied to/.test(x)), "applied/discarded batches are kept in Backups (14 days)");
    ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
    await ctx.close();
  }

  console.log("== parked merge edits, export/import, escaping");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    await load(p, INV);
    await p.evaluate(() => {
      QBR.journalPark("SAMPLE_Lenovo_Inventory_RICH.xlsx", [{ sheet: "02 DEVICES", ref: "F9", header: "Condition", key: "PF0TEST", mine: "<b>x</b>", why: "row \"PF0TEST\" was deleted in Excel" }]);
      const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "EXPORT ME" });
      const fp = Object.keys(QBR.journalEntries())[0]; QBR.journalClearFp(fp);          // a save clears pending…
    });
    await p.evaluate(() => Promise.all([QBR._journalLastTrash, QBR._journalLastWrite]));
    await p.reload(); await p.waitForTimeout(400); await load(p, INV);
    ok(await p.evaluate(() => QBR.journalParked().length === 1), "parked edit survives a save + reload (not dropped)");
    await p.evaluate(() => QBR.journalPanel()); await p.waitForTimeout(150);
    const pan = await p.evaluate(() => { const el = document.getElementById("qbr-journal-panel"); return { html: el.innerHTML, bold: !!el.querySelector("td b b") }; });
    ok(/Couldn't be placed/.test(pan.html) && /&lt;b&gt;x&lt;\/b&gt;/.test(pan.html), "panel lists the parked edit, HTML-escaped");
    await p.click('#qbr-journal-panel [data-jp="close"]');
    const ex = await p.evaluate(async () => {
      const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "EXPORT ME" }); QBR.invUpdateAsset(a.key, { client: "X & <Y>" });
      let text = null; QBR._journalDownloadHook = (n, t) => { text = t; return n; };
      QBR.journalExport();
      const d = JSON.parse(text);
      QBR._journalReset();
      const r1 = QBR.journalImport(text); await r1.done;
      const r2 = QBR.journalImport(text);
      const bad = QBR.journalImport('{"hello":1}');
      return { fmt: d.format, nOps: Object.values(d.files).reduce((n, e) => n + e.ops.length, 0), parked: d.parked.length, r1: { added: r1.added, parked: r1.parked, reloaded: r1.reloaded }, r2: r2.added, bad: bad.ok,
        cond: APP.model.inventory.assets[0].cond, client: APP.model.inventory.assets[0].client };
    });
    ok(ex.fmt === "qbr-unsaved-edits" && ex.nOps === 2 && ex.parked === 1, "export contains pending edits + parked edits");
    ok(ex.r1.added === 2 && ex.r1.parked === 1 && ex.r1.reloaded, "import restores them and re-parses the loaded file");
    ok(ex.cond === "EXPORT ME" && ex.client === "X & <Y>", "imported edits are applied to the model");
    ok(ex.r2 === 0 && ex.bad === false, "re-import is de-duplicated; non-export JSON is rejected");
    const tr = await p.evaluate(async () => { let t = null; QBR._journalDownloadHook = (n, x) => { t = x; return n; }; await QBR.journalExportTrash(); return JSON.parse(t).trash.length; });
    ok(tr >= 1, "saved batches can be exported from Backups (" + tr + ")");
    await p.evaluate(() => QBR.journalUnpark());
    ok(await p.evaluate(() => QBR.journalParked().length === 0), "Dismiss clears parked edits");
    ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
    await ctx.close();
  }

  console.log("== store merge rules (unit)");
  {
    const ctx = await b.newContext(); const p = await open(ctx, APPDIR);
    const u = await p.evaluate(() => {
      const M = QBR._journalMerge, N = QBR._journalNorm;
      const a = N({ v: 2, files: { F: { fileName: "f", ops: [{ id: "1", ts: 2 }, { id: "2", ts: 3 }] } }, gone: { "F\u00012": 9 }, parked: [] });
      const b = N({ v: 2, files: { F: { fileName: "f", ops: [{ id: "2", ts: 3 }, { id: "0", ts: 1 }] }, G: { fileName: "g", ops: [{ id: "9", ts: 1 }] } }, gone: {}, parked: [{ id: "p" }] });
      const m = M(a, b);
      return { F: m.files.F.ops.map(o => o.id).join(","), G: !!m.files.G, parked: m.parked.length, v1: Object.keys(N({ X: { fileName: "x", ops: [{ id: "a" }] } }).files).join() };
    });
    ok(u.F === "0,1" && u.G && u.parked === 1, "union by op id, sorted by time; tombstones beat older copies");
    ok(u.v1 === "X", "v1 map format is read");
    await ctx.close();
  }

  await b.close();
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) {}
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
