// Three-way merge (xlsx-merge.js) — unit tests. Node only.
// "Excel" edits are simulated by rewriting the workbook (as Excel does on save).
// Usage: node tests/merge-save.cjs
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const S = require(path.join(APPDIR, "js/xlsx-surgical.js"));
const M = require(path.join(APPDIR, "js/xlsx-merge.js"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const read = b => XLSX.read(b, { type: "array", cellStyles: true });
const INV = new Uint8Array(fs.readFileSync(path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx")));
const AUD = new Uint8Array(fs.readFileSync(path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx")));
const DEV = "02 DEVICES";

// ---- helpers: an "Excel" edit = AOA transform of one sheet, whole file rewritten
function excelEdit(bytes, fnBySheet) {
  const wb = XLSX.read(bytes, { type: "array" });
  for (const [name, fn] of Object.entries(fnBySheet)) {
    if (!wb.Sheets[name]) { wb.SheetNames.push(name); wb.Sheets[name] = XLSX.utils.aoa_to_sheet(fn([])); continue; }
    const old = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(old, { header: 1, defval: null, raw: true });
    // tag every cell with its original address so hyperlinks follow the cell (Excel keeps links when rows/columns move)
    const tagged = aoa.map((row, r) => row.map((v, c) => { const o = old[XLSX.utils.encode_cell({ r, c })]; return o && o.l ? { __v: v, __l: o.l } : v; }));
    const out = fn(tagged); const links = [];
    const plain = out.map((row, r) => (row || []).map((v, c) => { if (v && typeof v === "object" && "__v" in v) { links.push([r, c, v.__l]); return v.__v; } return v; }));
    const ws = XLSX.utils.aoa_to_sheet(plain);
    links.forEach(([r, c, l]) => { const a = XLSX.utils.encode_cell({ r, c }); if (ws[a]) ws[a].l = l; });
    wb.Sheets[name] = ws;
  }
  return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
}
const hdr = (aoa, h) => aoa[0].findIndex(x => String(x).toLowerCase() === h.toLowerCase());
function setByHeader(ws, rowIdx, h, v) { // dashboard edit in "mine" (row index 0-based incl. header)
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const c = hdr(aoa, h); const a = XLSX.utils.encode_cell({ r: rowIdx, c });
  ws[a] = typeof v === "number" ? { t: "n", v } : { t: "s", v };
  const g = XLSX.utils.decode_range(ws["!ref"]); g.e.r = Math.max(g.e.r, rowIdx); g.e.c = Math.max(g.e.c, c); ws["!ref"] = XLSX.utils.encode_range(g);
}
function rowOf(wb, sheet, keyH, key) {
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  const c = hdr(aoa, keyH); return aoa.findIndex(r => r && String(r[c]) === key);
}
function val(wb, sheet, keyH, key, h) {
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  const r = aoa.find(x => x && String(x[hdr(aoa, keyH)]) === key); return r ? r[hdr(aoa, h)] : undefined;
}
// run: base bytes, dashboard edit fn(mineWb), theirs bytes → {res, final bytes, final wb}
async function run(baseBytes, mineFn, theirsBytes, choices) {
  const base = read(baseBytes), mine = read(baseBytes); mineFn(mine);
  const theirs = read(theirsBytes), final = read(theirsBytes);
  const res = M.rebase(base, mine, theirs, final);
  if (choices) M.resolve(final, res.conflicts, choices);
  const out = await S.surgicalSave(theirsBytes, theirs, final);
  const bytes = out.ok ? (out.bytes || theirsBytes) : null;
  return { res, out, wb: bytes ? XLSX.read(bytes, { type: "array" }) : null, bytes, final };
}
const aoaInv = XLSX.utils.sheet_to_json(read(INV).Sheets[DEV], { header: 1, defval: null });
const SN = r => String(aoaInv[r][0]);         // serial in row r
const S1 = SN(1), S2 = SN(2), S3 = SN(3), S9 = SN(9);

(async () => {
  console.log("== 1. Excel fixed another cell; dashboard edited a different row");
  let theirs = excelEdit(INV, { [DEV]: a => { a[2][hdr(a, "Model")] = "EXCEL-FIXED-MODEL"; return a; } });
  let r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Client / Organization", "DASH SCHOOL"), theirs);
  ok(r.res.applied === 1 && !r.res.conflicts.length && !r.res.unresolved.length, "1 edit applied, no conflicts");
  ok(val(r.wb, DEV, "Serial Number", S1, "Client / Organization") === "DASH SCHOOL", "dashboard edit saved");
  ok(val(r.wb, DEV, "Serial Number", S2, "Model") === "EXCEL-FIXED-MODEL", "Excel edit kept");

  console.log("== 2. Same row, different cells");
  theirs = excelEdit(INV, { [DEV]: a => { a[1][hdr(a, "Model")] = "EXCEL-MODEL"; return a; } });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Condition", "For Repair"), theirs);
  ok(r.res.applied === 1 && !r.res.conflicts.length, "no conflict (different cells of the same row)");
  ok(val(r.wb, DEV, "Serial Number", S1, "Model") === "EXCEL-MODEL" && val(r.wb, DEV, "Serial Number", S1, "Condition") === "For Repair", "both edits present");

  console.log("== 3. Same cell changed on both sides");
  theirs = excelEdit(INV, { [DEV]: a => { a[1][hdr(a, "Client / Organization")] = "EXCEL SCHOOL"; return a; } });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Client / Organization", "DASH SCHOOL"), theirs);
  ok(r.res.conflicts.length === 1 && r.res.conflicts[0].theirs === "EXCEL SCHOOL" && r.res.conflicts[0].mine === "DASH SCHOOL", "conflict reported with both values (key " + (r.res.conflicts[0] || {}).key + ")");
  ok(val(r.wb, DEV, "Serial Number", S1, "Client / Organization") === "EXCEL SCHOOL", "default keeps Excel's value");
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Client / Organization", "DASH SCHOOL"), theirs, { 0: "mine" });
  ok(val(r.wb, DEV, "Serial Number", S1, "Client / Organization") === "DASH SCHOOL", "choosing 'dashboard' writes the dashboard value");
  theirs = excelEdit(INV, { [DEV]: a => { a[1][hdr(a, "Client / Organization")] = "SAME VALUE"; return a; } });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Client / Organization", "SAME VALUE"), theirs);
  ok(!r.res.conflicts.length && r.res.skippedSame === 1, "same value on both sides → no conflict");

  console.log("== 4. Excel inserted a column before the edited one");
  theirs = excelEdit(INV, { [DEV]: a => a.map((row, i) => { const x = row.slice(); x.splice(1, 0, i ? "note " + i : "Notes"); return x; }) });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Client / Organization", "DASH SCHOOL"), theirs);
  ok(r.res.applied === 1 && r.res.remapped, "edit remapped to the shifted column");
  ok(val(r.wb, DEV, "Serial Number", S1, "Client / Organization") === "DASH SCHOOL" && val(r.wb, DEV, "Serial Number", S1, "Notes") === "note 1", "lands under its header; Excel's new column untouched");

  console.log("== 5. Excel renamed the edited column's header");
  theirs = excelEdit(INV, { [DEV]: a => { a[0][hdr(a, "Condition")] = "Unit Condition"; return a; } });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Condition", "For Repair"), theirs);
  ok(r.res.unresolved.length === 1 && /renamed or removed/.test(r.res.unresolved[0].why) && r.res.applied === 0, "flagged as unresolved, nothing guessed");

  console.log("== 6. Excel sorted the rows (reverse)");
  theirs = excelEdit(INV, { [DEV]: a => [a[0]].concat(a.slice(1).reverse()) });
  r = await run(INV, m => { setByHeader(m.Sheets[DEV], 1, "Condition", "SORT-A"); setByHeader(m.Sheets[DEV], 9, "Condition", "SORT-B"); }, theirs);
  ok(r.res.applied === 2 && !r.res.conflicts.length && r.res.remapped, "2 edits remapped by serial number");
  ok(val(r.wb, DEV, "Serial Number", S1, "Condition") === "SORT-A" && val(r.wb, DEV, "Serial Number", S9, "Condition") === "SORT-B", "each edit on the right unit");

  console.log("== 7. Excel deleted the edited row");
  theirs = excelEdit(INV, { [DEV]: a => a.filter((x, i) => i !== 3) });
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 3, "Condition", "GONE"), theirs);
  ok(r.res.unresolved.length === 1 && r.res.unresolved[0].key === S3.toUpperCase().trim(), "deleted row → unresolved (" + (r.res.unresolved[0] || {}).why + ")");

  console.log("== 8. Both appended rows; dashboard added a column; Excel added a sheet");
  theirs = excelEdit(INV, { [DEV]: a => a.concat([["EXCEL-NEW-1", "Excel School"]]), "NEW SHEET": () => [["Hello"], ["Excel"]] });
  const last = aoaInv.length;
  r = await run(INV, m => { setByHeader(m.Sheets[DEV], last, "Serial Number", "DASH-NEW-1"); setByHeader(m.Sheets[DEV], last, "Client / Organization", "Dash School");
    const ws = m.Sheets[DEV]; ws[XLSX.utils.encode_cell({ r: 0, c: 16 })] = { t: "s", v: "Batch Code" }; ws[XLSX.utils.encode_cell({ r: 1, c: 16 })] = { t: "s", v: "B-1" };
    const g = XLSX.utils.decode_range(ws["!ref"]); g.e.c = 16; ws["!ref"] = XLSX.utils.encode_range(g); }, theirs);
  ok(!r.res.conflicts.length && !r.res.unresolved.length, "no conflicts");
  ok(rowOf(r.wb, DEV, "Serial Number", "DASH-NEW-1") === rowOf(r.wb, DEV, "Serial Number", "EXCEL-NEW-1") + 1, "dashboard row appended after Excel's new row");
  ok(val(r.wb, DEV, "Serial Number", S1, "Batch Code") === "B-1", "new 'Batch Code' column created and filled");
  ok(r.wb.SheetNames.includes("NEW SHEET"), "Excel's new sheet kept");

  console.log("== 9. Dashboard created a sheet (EDIT LOG) that the file doesn't have");
  r = await run(INV, m => XLSX.utils.book_append_sheet(m, XLSX.utils.aoa_to_sheet([["Timestamp", "Action"], ["2026-10-06", "x"]]), "EDIT LOG"), theirs);
  ok(r.res.sheetsAdded.includes("EDIT LOG") && r.wb.SheetNames.includes("EDIT LOG") && r.wb.SheetNames.includes("NEW SHEET"), "sheet added, Excel's sheets kept");

  console.log("== 10. Audit workbook: Excel added a column + fixed a school; dashboard exempts + adds a link");
  const sheet = "JANUARY";
  theirs = excelEdit(AUD, { [sheet]: a => a.map((row, i) => { const x = row.slice(); x.splice(2, 0, i ? i * 10 : "CASES"); if (i === 2) x[hdr(a, "TOTAL RISKY USERS") + 1] = 999; return x; }) });
  r = await run(AUD, m => { const ws = m.Sheets[sheet];
    ws.K1 = { t: "s", v: "EXEMPT" }; ws.L1 = { t: "s", v: "EXEMPT REASON" }; ws.K2 = { t: "s", v: "Yes" }; ws.L2 = { t: "s", v: "No GDAP" };
    ws.E3 = Object.assign({}, ws.E3 || { t: "s", v: "Open" }, { l: { Target: "https://contoso.sharepoint.com/x?a=1&b=2" } });
    const g = XLSX.utils.decode_range(ws["!ref"]); g.e.c = 11; ws["!ref"] = XLSX.utils.encode_range(g); }, theirs);
  const school1 = String(XLSX.utils.sheet_to_json(read(AUD).Sheets[sheet], { header: 1 })[1][0]), school2 = String(XLSX.utils.sheet_to_json(read(AUD).Sheets[sheet], { header: 1 })[2][0]);
  ok(!r.res.conflicts.length && !r.res.unresolved.length, "no conflicts (" + r.res.applied + " applied)");
  ok(val(r.wb, sheet, "SCHOOL", school1, "EXEMPT") === "Yes" && val(r.wb, sheet, "SCHOOL", school1, "EXEMPT REASON") === "No GDAP", "EXEMPT columns written after Excel's shifted columns");
  ok(val(r.wb, sheet, "SCHOOL", school2, "TOTAL RISKY USERS") === 999 && val(r.wb, sheet, "SCHOOL", school1, "CASES") === 10, "Excel's edits + new column kept");
  const ws2 = r.wb.Sheets[sheet], aoa2 = XLSX.utils.sheet_to_json(ws2, { header: 1 }), refC = hdr(aoa2, "REFERENCES");
  const lk = ws2[XLSX.utils.encode_cell({ r: 2, c: refC })];
  ok(lk && lk.l && S._decodeStable(lk.l.Target) === "https://contoso.sharepoint.com/x?a=1&b=2", "reference link written to the shifted References column");

  console.log("== 11. Formatting survives a merged save (real file, not a rewrite)");
  const theirsFmt = await (async () => { // surgical edit = Excel-like small change that keeps formatting
    const b = read(INV), w = read(INV); w.Sheets[DEV].C3 = { t: "s", v: "EXCEL-SMALL" };
    return (await S.surgicalSave(INV, b, w)).bytes; })();
  r = await run(INV, m => setByHeader(m.Sheets[DEV], 1, "Condition", "For Repair"), theirsFmt);
  const x0 = await S.zipText(S.zipRead(INV), "xl/worksheets/sheet2.xml"), x1 = await S.zipText(S.zipRead(r.bytes), "xl/worksheets/sheet2.xml");
  const frag = (x, t) => (x.match(new RegExp("<" + t + "\\b[\\s\\S]*?</" + t + ">", "g")) || []).join("");
  ok(frag(x0, "conditionalFormatting") === frag(x1, "conditionalFormatting") && frag(x0, "dataValidations") === frag(x1, "dataValidations"), "conditional formatting + validation intact");
  ok(r.wb.Sheets[DEV].C3.v === "EXCEL-SMALL" && val(r.wb, DEV, "Serial Number", S1, "Condition") === "For Repair", "both edits present");

  console.log(`\n${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
