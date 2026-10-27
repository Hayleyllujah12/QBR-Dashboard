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
];

/* ---------------- harness ---------------- */
async function open(b, k) {
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message)); p.on("dialog", d => d.dismiss());
  await p.goto(APP); await p.waitForTimeout(300);
  await p.evaluate(() => { localStorage.clear(); localStorage.setItem("qbr-remember", "0"); });
  const bytes = k.bytes();
  await p.evaluate(({ name, data }) => {
    const u = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    return loadItems([{ name, blob: new File([u], name, { lastModified: 1700000000000 }) }], {});
  }, { name: k.name, data: b64(bytes) });
  await p.evaluate(({ name, data, kind }) => {
    const bin = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    const T = window.__T = { bytes: bin, mtime: 1700000000000, writes: 0, lock: false, deny: false, flip: null, gets: 0 };
    const handle = { name, kind: "file",
      queryPermission: async () => T.deny ? "prompt" : "granted", requestPermission: async () => T.deny ? "denied" : "granted",
      getFile: async () => { T.gets++; if (T.flip) { T.bytes = T.flip[T.gets % 2]; T.mtime += 1; } return new File([T.bytes], name, { lastModified: T.mtime }); },
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), abort: async () => {},
        close: async () => { if (T.lock) throw new DOMException("The process cannot access the file", "NoModificationAllowedError"); T.bytes = parts[0]; T.mtime += 1000; T.writes++; } }; } };
    window.__setFile = d => { T.bytes = Uint8Array.from(atob(d), c => c.charCodeAt(0)); T.mtime += 1000; };
    window.__getFile = () => { let s = ""; for (let i = 0; i < T.bytes.length; i += 32768) s += String.fromCharCode.apply(null, T.bytes.subarray(i, i + 32768)); return btoa(s); };
    const f0 = APP.files.find(f => f.name === name).blob, fp = QBR.fpOf(name, f0);
    const kinds = [...((QBR._kindByFp || {})[fp] || [])];
    QBR._fsLinks = [{ fp, kinds, name, size: bin.length, lastModified: T.mtime, hash: QBR.fsHashOf(bin), handle }];
    window.__kinds = kinds;
    void kind;
  }, { name: k.name, data: b64(bytes), kind: k.kind });
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

(async () => {
  const only = process.argv.slice(2);
  const b = await chromium.launch();
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
    // S8 second save in the same session
    await p.evaluate(k.edit, 42);
    r = await save(p, k.kind);
    ok(r.mode === "file" && r.writes === 2 && await pending(p, k.kind) === 0, "second save in the same session (" + r.mode + ")");
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
    await p.evaluate(async () => { await loadItems(APP.files.slice(), {}); });
    ok(await pending(p, k.kind) > 0, "reload: the pending edit survives");
    r = await save(p, k.kind); now = await fileNow(p);
    ok(r.mode === "file" && diff(orig, now).length > 0 && await pending(p, k.kind) === 0, "…and saves after the reload (" + r.mode + ")");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();
  }
  await b.close();
  console.log("RESULT: " + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL " + (e && e.stack || e)); process.exit(1); });
