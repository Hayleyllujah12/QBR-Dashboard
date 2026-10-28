// Excel write contract — shared scenario test (docs/EXCEL_WRITE_CONTRACT.md, Phase 0).
// One table of workbook KINDS × one list of SCENARIOS. Every kind that writes to Excel through the
// shared save path (QBR.fsSaveKind) must pass every scenario. A new kind joins by adding ONE row to
// KINDS (a dummy sample workbook + one dashboard edit). Dummy/sample data only.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-write-contract.cjs [kind ...]
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const S = require(path.join(ROOT, "js/xlsx-surgical.js"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const b64 = u8 => Buffer.from(u8).toString("base64");
const rd = u8 => XLSX.read(u8, { type: "array", cellStyles: true });

/* ---------------- dummy workbooks for kinds without a sample file ---------------- */
function book(sheets) {
  const wb = XLSX.utils.book_new();
  Object.keys(sheets).forEach(n => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheets[n]), n));
  return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
}
const SUPPLIES = book({
  Items: [["ItemID", "ItemName", "Category", "Unit", "BeginningBalance", "ReorderLevel", "Supplier", "UnitCost", "Location"],
    ["P-PAPER", "Bond paper A4", "Office", "ream", 40, 10, "Sample Supplier", 5, "Shelf 1"],
    ["T-TONER", "Toner black", "Printer", "pc", 6, 2, "Sample Supplier", 60, "Shelf 2"]],
  Transactions: [["Date", "DocNo", "TransactionType", "ItemID", "Quantity", "Department/Project", "RequestedBy", "ApprovedBy", "Notes"],
    ["2026-10-01", "D-1", "Received", "P-PAPER", 10, "Ops", "DOE, John", "ROE, Jane", ""]],
});
const STORAGE = book({
  Q4: [["School (Rakso Managed)", "ORGANIZATION", "ONEDRIVE", "EXCHANGE", "SHAREPOINT", "CURRENT STORAGE", "USAGE", "USED", "TOTAL"],
    ["Sample Academy", "ORG-A", "", "", "", "", "", "", ""], ["Maple Ridge College of Westbrook", "ORG-B", "", "", "", "", "", "", ""]],
});

/* ---------------- KINDS: one row per workbook kind on the shared save path ---------------- */
// edit: page function(arg) that records ONE dashboard edit. append: the edit adds a row (no same-cell conflict).
const KINDS = [
  { kind: "audit", label: "Risky sign-ins editor", name: "SAMPLE_AUDIT_RICH.xlsx", bytes: () => new Uint8Array(fs.readFileSync(path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx"))),
    edit: v => { const a = QBR._audit, ms = Object.values(a.months)[0]; QBR.auditUpdateCell(a.fp, ms.sheet, ms.rows[0].r, "risky", String(v)); } },
  { kind: "storage", label: "Guided audit › storage", name: "STORAGE REPORT.xlsx", bytes: () => STORAGE,
    edit: v => { const s = QBR._storage, q = s.quarters.Q4; QBR.storagePasteRow(s.fp, q.sheet, q.rows[0].r, [String(v), "2", "3", "4", "5", "6", "7"]); } },
  { kind: "assets", label: "Inventory", name: "SAMPLE_Lenovo_Inventory_RICH.xlsx", bytes: () => new Uint8Array(fs.readFileSync(path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"))),
    edit: v => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "DASH " + v }); } },
  { kind: "supplies", label: "Supplies", name: "SAMPLE_SUPPLIES.xlsx", bytes: () => SUPPLIES, append: true,
    edit: v => { QBR.supRecordTransaction({ date: "2026-10-28", type: "Used", itemId: "T-TONER", qty: Number(v) || 1, dept: "Ops", requestedBy: "DOE, John" }); } },
  // Phase 1: a tab-owned "side" file (QBR.fsAttachSide) on the same path — editLog off, as decided for the
  // MOS log / Trainers tracker / domain tracker. Registered by the test itself (no app module uses it yet).
  { kind: "testside", label: "Side file (test adapter)", name: "SIDE LOG.xlsx", side: true, editLog: false,
    bytes: () => book({ LOG: [["Code", "Printed"], ["CODE-A", ""], ["CODE-B", ""], ["CODE-C", ""]] }),
    edit: v => { QBR.journalRecord("testside", "tsSetCell", ["LOG", 2, 1, "P" + v]); } },
];
// Page-side registration of the test adapters (side file + autosave).
const SIDE_SETUP = ({ kind, autosave }) => {
  const set = (wb, fp, a) => { const ws = wb.Sheets[a[0]]; if (!ws) return "skip: sheet"; const ref = XLSX.utils.encode_cell({ r: a[1] - 1, c: a[2] });
    if (ws[ref] && ws[ref].f) return "skip: formula"; if (ws[ref]) { ws[ref].v = a[3]; ws[ref].t = "s"; delete ws[ref].w; } else ws[ref] = { t: "s", v: a[3] }; return "ok"; };
  QBR.wbKinds.register(kind, { label: "test " + kind, source: "side", editLog: false, autosave: !!autosave, ops: { tsSetCell: set } });
  window.QBR.tsSetCell = function () {};   // replay hook (the tab owns its model; nothing to re-apply in the test)
};

/* ---------------- harness ---------------- */
async function open(b, k) {
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message)); p.on("dialog", d => d.dismiss());
  // Fixed clock: saves stamp dates (EDIT LOG rows, docProps); WC_DUMP byte comparisons need them stable.
  await p.clock.setFixedTime(new Date("2026-10-28T01:00:00Z"));
  await p.goto(APP); await p.waitForTimeout(300);
  await p.evaluate(() => { localStorage.clear(); localStorage.setItem("qbr-remember", "0"); });
  const bytes = k.bytes();
  // side kinds: the dashboard loads an unrelated workbook; the side file is attached, not loaded
  await p.evaluate(({ name, data }) => {
    const u = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    return loadItems([{ name, blob: new File([u], name, { lastModified: 1700000000000 }) }], {});
  }, k.side ? { name: "fixture.xlsx", data: b64(new Uint8Array(fs.readFileSync(path.join(__dirname, "fixture.xlsx")))) } : { name: k.name, data: b64(bytes) });
  if (k.side) await p.evaluate(SIDE_SETUP, { kind: k.kind, autosave: !!k.autosave });
  await p.evaluate(({ name, data, kind, side }) => {
    const bin = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    window.__side = side;
    const T = window.__T = { bytes: bin, mtime: 1700000000000, writes: 0, lock: false, deny: false, flip: null, gets: 0 };
    const handle = { name, kind: "file",
      queryPermission: async () => T.deny ? "prompt" : "granted", requestPermission: async () => T.deny ? "denied" : "granted",
      getFile: async () => { T.gets++; if (T.flip) { T.bytes = T.flip[T.gets % 2]; T.mtime += 1; } return new File([T.bytes], name, { lastModified: T.mtime }); },
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), abort: async () => {},
        close: async () => { if (T.lock) throw new DOMException("The process cannot access the file", "NoModificationAllowedError"); T.bytes = parts[0]; T.mtime += 1000; T.writes++; } }; } };
    window.__setFile = d => { T.bytes = Uint8Array.from(atob(d), c => c.charCodeAt(0)); T.mtime += 1000; };
    window.__getFile = () => { let s = ""; for (let i = 0; i < T.bytes.length; i += 32768) s += String.fromCharCode.apply(null, T.bytes.subarray(i, i + 32768)); return btoa(s); };
    window.__handle = handle;
    if (window.__side) return;
    const f0 = APP.files.find(f => f.name === name).blob, fp = QBR.fpOf(name, f0);
    const kinds = [...((QBR._kindByFp || {})[fp] || [])];
    QBR._fsLinks = [{ fp, kinds, name, size: bin.length, lastModified: T.mtime, hash: QBR.fsHashOf(bin), handle }];
    window.__kinds = kinds;
    void kind;
  }, { name: k.name, data: b64(bytes), kind: k.kind, side: k.side });
  if (k.side) await p.evaluate(async kind => {
    window.__sideSaved = [];
    const link = await QBR.fsAttachSide(kind, window.__handle, null, { onSaved: (u8, why) => window.__sideSaved.push(why + ":" + (u8 ? u8.length : 0)) });
    window.__kinds = [...((QBR._fsSideKinds || {})[link.fp] || [])];
  }, k.kind);
  return { p, errs, orig: bytes };
}
const fileNow = async p => new Uint8Array(Buffer.from(await p.evaluate(() => window.__getFile()), "base64"));
const pending = (p, kind) => p.evaluate(k => QBR.journalUnsavedCount(k), kind);
const save = (p, kind) => p.evaluate(async k => { const r = await QBR.fsSaveKind(k); return { mode: r.mode, merged: !!r.merged, writes: __T.writes, applied: r.applied }; }, kind);
// changed cells between two workbooks: ["Sheet!A1", ...]
function diff(a, b) {
  const A = rd(a), B = rd(b), out = [];
  B.SheetNames.forEach(n => {
    const wa = A.Sheets[n] || {}, wb = B.Sheets[n];
    Object.keys(wb).filter(r => r[0] !== "!").forEach(r => { const x = wa[r], y = wb[r]; if (!x || String(x.v) !== String(y.v)) out.push(n + "!" + r); });
  });
  return out;
}
const cellV = (u8, ref) => { const [n, r] = ref.split("!"); const c = (rd(u8).Sheets[n] || {})[r]; return c ? String(c.v) : undefined; };
async function excelSet(u8, ref, v) { const [n, r] = ref.split("!"); const base = rd(u8), work = rd(u8); work.Sheets[n][r] = { t: typeof v === "number" ? "n" : "s", v }; return (await S.surgicalSave(u8, base, work)).bytes; }
// a cell Excel can change that the dashboard edit does not touch: same sheet, last row, column A
function otherRef(u8, editRef) {
  const [n] = editRef.split("!"), ws = rd(u8).Sheets[n], last = XLSX.utils.decode_range(ws["!ref"]).e.r + 1;
  const col = editRef.split("!")[1].replace(/\d+/g, "");
  const r = n + "!" + (col === "B" ? "C" : "B") + last;
  return r === editRef ? n + "!D" + last : r;
}

const DUMP = process.env.WC_DUMP || "";   // dir: write every saved file (byte-identity check across branches)
const dump = (k, scen, u8) => { if (DUMP) { fs.mkdirSync(DUMP, { recursive: true }); fs.writeFileSync(path.join(DUMP, k.kind + "-" + scen + ".xlsx"), Buffer.from(u8)); } };

(async () => {
  const only = process.argv.slice(2);
  const b = await chromium.launch();
  if (!only.length) {   // Phase 1: the registry covers every cell writer
    console.log("== registry");
    const pg = await b.newPage(); await pg.goto(APP); await pg.waitForTimeout(300);
    const reg = await pg.evaluate(() => {
      const kinds = QBR.wbKinds.list(), ops = Object.keys(window.PATCH_OPS || {});
      const owner = op => kinds.filter(k => QBR.wbKinds.opFor(k, op) === window.PATCH_OPS[op]);
      return { kinds, missing: ops.filter(op => owner(op).length !== 1), n: ops.length,
        flags: kinds.map(k => { const a = QBR.wbKinds.get(k); return k + ":" + a.source + "/" + a.editLog + "/" + a.autosave; }) };
    });
    ok(["audit", "storage", "usage", "assets", "supplies"].every(k => reg.kinds.includes(k)), "built-in kinds registered (" + reg.kinds.join(", ") + ")");
    ok(reg.missing.length === 0 && reg.n > 0, "every PATCH_OPS cell writer (" + reg.n + ") belongs to exactly one kind" + (reg.missing.length ? " — not: " + reg.missing.join(", ") : ""));
    ok(reg.flags.every(f => /:loaded\/true\/false$/.test(f)), "built-ins unchanged: loaded files, EDIT LOG on, no autosave (" + reg.flags.join(" ") + ")");
    ok(await pg.evaluate(() => { try { QBR.wbKinds.register("bad", { ops: { x: 1 } }); return false; } catch (e) { return true; } }), "register() rejects a non-function op");
    await pg.close();
  }
  for (const k of KINDS.filter(x => !only.length || only.includes(x.kind))) {
    console.log("== " + k.kind + " (" + k.label + ")");
    // S1 save ok — also learns which cell(s) the edit writes
    let { p, errs, orig } = await open(b, k);
    ok((await p.evaluate(() => window.__kinds)).includes(k.kind), "workbook detected as kind '" + k.kind + "'");
    await p.evaluate(k.edit, 41);
    ok(await pending(p, k.kind) > 0, "edit recorded in the journal before any write");
    let r = await save(p, k.kind), now = await fileNow(p);
    const changed = diff(orig, now).filter(x => !/^EDIT LOG!/.test(x));
    ok(r.mode === "file" && r.writes === 1 && changed.length > 0 && await pending(p, k.kind) === 0, "save ok: written once, journal cleared (" + r.mode + ", " + changed.length + " cell(s): " + changed.slice(0, 4).join(" ") + ")");
    const editRef = changed[0];
    dump(k, "S1-save", now);
    if (k.editLog === false) ok(!rd(now).SheetNames.includes("EDIT LOG"), "editLog off: no EDIT LOG sheet added");
    else ok(rd(now).SheetNames.includes("EDIT LOG"), "editLog on: EDIT LOG sheet appended (unchanged behaviour)");
    if (k.side) ok((await p.evaluate(() => window.__sideSaved)).some(x => /^saved:\d+/.test(x)), "side file: its tab was handed the saved bytes (onSaved)");
    // S8 second save in the same session
    await p.evaluate(k.edit, 42);
    r = await save(p, k.kind);
    ok(r.mode === "file" && r.writes === 2 && await pending(p, k.kind) === 0, "second save in the same session (" + r.mode + ")");
    dump(k, "S8-second", await fileNow(p));
    // fallback copy (download-unsafe path): something real downloads, edits stay pending
    await p.evaluate(k.edit, 49);
    const dl = await p.evaluate(k2 => { let got = null; const w = XLSX.writeFile; XLSX.writeFile = (wb, fn) => { got = { fn, sheets: wb.SheetNames.length, has49: wb.SheetNames.some(n => Object.keys(wb.Sheets[n]).some(r => r[0] !== "!" && /49$/.test(String(wb.Sheets[n][r].v)))) }; };
      try { const r = QBR.fsDownloadKind(k2, "download-unsafe"); return { mode: r.mode, got, pend: QBR.journalUnsavedCount(k2) }; } finally { XLSX.writeFile = w; } }, k.kind);
    const modelCopy = k.kind === "assets" || k.kind === "supplies";   // these rebuild from the model (unchanged); the rest copy THEIR file
    ok(dl.mode === "download-unsafe" && dl.got && dl.got.sheets > 0 && dl.pend > 0 &&
      (modelCopy || (dl.got.fn === k.name.replace(/\.xlsx$/i, "") + " (dashboard copy).xlsx" && dl.got.has49)),
      "fallback copy downloads " + (modelCopy ? "the model export" : "a patched copy of this file") + " and keeps the edit pending (" + JSON.stringify(dl) + ")");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();

    // S2 locked → kept; after unlock → saved
    ({ p, errs } = await open(b, k));
    await p.evaluate(k.edit, 43);
    await p.evaluate(() => { __T.lock = true; });
    r = await save(p, k.kind);
    ok(r.mode === "locked" && r.writes === 0 && await pending(p, k.kind) > 0, "open in Excel (locked): nothing written, edit kept (" + r.mode + ")");
    ok(await p.evaluate(() => /open in Excel/.test(QBR.saveResultText({ mode: "locked", name: "x" }).text)), "locked outcome has the shared message");
    await p.evaluate(() => { __T.lock = false; });
    r = await save(p, k.kind);
    ok(r.mode === "file" && await pending(p, k.kind) === 0, "after Excel closes: saved (" + r.mode + ")");
    await p.close();

    // S3 changed in Excel (another cell) → merged, both kept
    ({ p, errs, orig } = await open(b, k));
    await p.evaluate(k.edit, 44);
    const oref = otherRef(orig, editRef);
    await p.evaluate(d => __setFile(d), b64(await excelSet(orig, oref, "EXCEL-EDIT")));
    r = await save(p, k.kind); now = await fileNow(p);
    dump(k, "S3-merged", now);
    ok(r.mode === "file" && r.merged && cellV(now, oref) === "EXCEL-EDIT" && diff(orig, now).filter(x => x !== oref && !/^EDIT LOG!/.test(x)).length > 0,
      "changed in Excel: merged — Excel's " + oref + " kept and the dashboard edit written (" + r.mode + ")");
    await p.close();

    // S4 same cell changed in Excel → review → Cancel writes nothing
    if (!k.append) {
      ({ p, errs, orig } = await open(b, k));
      await p.evaluate(k.edit, 45);
      await p.evaluate(d => __setFile(d), b64(await excelSet(orig, editRef, "EXCEL-SAME")));
      r = await p.evaluate(async k2 => { QBR._mergeAutoResolve = () => null; const r = await QBR.fsSaveKind(k2); return { mode: r.mode, writes: __T.writes }; }, k.kind);
      ok(r.mode === "merge-cancelled" && r.writes === 0 && await pending(p, k.kind) > 0, "same cell changed in Excel → Cancel: nothing written, edit kept (" + r.mode + ")");
      await p.close();
    } else console.log("  · same-cell conflict: n/a (the edit appends a row)");

    // S5 file keeps changing (sync) → busy, nothing written
    ({ p, errs, orig } = await open(b, k));
    await p.evaluate(k.edit, 46);
    const v1 = await excelSet(orig, otherRef(orig, editRef), "SYNC-1"), v2 = await excelSet(orig, otherRef(orig, editRef), "SYNC-2");
    await p.evaluate(([a, c]) => { const f = s => Uint8Array.from(atob(s), x => x.charCodeAt(0)); __T.flip = [f(a), f(c)]; __T.mtime += 1000; }, [b64(v1), b64(v2)]);
    r = await save(p, k.kind);
    ok(r.mode === "busy" && r.writes === 0 && await pending(p, k.kind) > 0, "file keeps changing (sync): nothing written, edit kept (" + r.mode + ")");
    await p.close();

    // S6 permission denied
    ({ p, errs } = await open(b, k));
    await p.evaluate(k.edit, 47);
    await p.evaluate(() => { __T.deny = true; });
    r = await save(p, k.kind);
    ok(r.mode === "denied" && r.writes === 0 && await pending(p, k.kind) > 0, "permission denied: nothing written, edit kept (" + r.mode + ")");
    await p.close();

    // S7 reload keeps the pending edit; it still saves afterwards
    ({ p, errs, orig } = await open(b, k));
    await p.evaluate(k.edit, 48);
    if (k.side) await p.evaluate(async n => { await QBR.fsReloadFromFile(n, { quiet: true }); }, k.name);
    else await p.evaluate(async () => { await loadItems(APP.files.slice(), {}); });
    ok(await pending(p, k.kind) > 0, "reload: the pending edit survives");
    if (k.side) ok((await p.evaluate(() => window.__sideSaved)).some(x => /^reloaded:/.test(x)), "side reload re-syncs its tab only (onSaved 'reloaded')");
    r = await save(p, k.kind); now = await fileNow(p);
    dump(k, "S7-after-reload", now);
    ok(r.mode === "file" && diff(orig, now).length > 0 && await pending(p, k.kind) === 0, "…and saves after the reload (" + r.mode + ")");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();
  }
  // Autosave (adapter.autosave): an edit saves itself; a locked file keeps it pending with the shared note.
  if (!only.length || only.includes("autosave")) {
    console.log("== autosave (side test adapter, autosave on)");
    const k = Object.assign({}, KINDS.find(x => x.kind === "testside"), { kind: "testauto", autosave: true,
      edit: v => { QBR.journalRecord("testauto", "tsSetCell", ["LOG", 3, 1, "A" + v]); } });
    let { p, errs } = await open(b, k);
    await p.evaluate(k.edit, 1); await p.evaluate(k.edit, 2);   // a burst → one save
    await p.waitForTimeout(900);
    ok(await p.evaluate(() => __T.writes) === 1 && await pending(p, "testauto") === 0 && cellV(await fileNow(p), "LOG!B3") === "A2", "a burst of edits auto-saves once, last value wins");
    await p.evaluate(() => { __T.lock = true; }); await p.evaluate(k.edit, 3); await p.waitForTimeout(900);
    ok(await p.evaluate(() => __T.writes) === 1 && await pending(p, "testauto") > 0, "locked: auto-save writes nothing, the edit stays pending");
    ok(await p.evaluate(() => { const n = QBR.persistNoteOf("testauto"); return !!n && n.warn && /open in Excel/.test(n.text); }), "…and the shared red 'open in Excel' note is set");
    await p.evaluate(() => { __T.lock = false; }); await p.evaluate(k.edit, 4); await p.waitForTimeout(900);
    ok(await p.evaluate(() => __T.writes) === 2 && await pending(p, "testauto") === 0 && cellV(await fileNow(p), "LOG!B3") === "A4", "after Excel closes the next edit saves everything pending");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();
  }
  await b.close();
  console.log("RESULT: " + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL " + (e && e.stack || e)); process.exit(1); });
