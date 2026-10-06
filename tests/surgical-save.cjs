// Format-safe ("surgical") save — unit + fidelity tests. Node only, no browser.
// Usage: node tests/surgical-save.cjs            (fixtures: make-rich-fixture.py)
// Optional: VALIDATE=1 also opens each output with openpyxl and LibreOffice.
const path = require("path"), fs = require("fs"), cp = require("child_process"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const S = require(path.join(APPDIR, "js/xlsx-surgical.js"));
const FX = n => path.join(__dirname, n);
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const read = b => XLSX.read(b, { type: "array", cellStyles: true });
const td = new TextDecoder();
async function partText(bytes, name) { return S.zipText(S.zipRead(bytes), name); }
function rawSame(a, b, name) { // compressed bytes identical?
  const ea = S.zipRead(a).byName.get(name), eb = S.zipRead(b).byName.get(name);
  if (!ea || !eb) return false;
  const x = a.subarray(ea.dataStart, ea.dataStart + ea.csize), y = b.subarray(eb.dataStart, eb.dataStart + eb.csize);
  return x.length === y.length && x.every((v, i) => v === y[i]) && ea.crc === eb.crc;
}
function untouchedIdentical(a, b, changed) {
  const za = S.zipRead(a); const bad = [];
  za.entries.forEach(e => { if (!changed.includes(e.name) && !rawSame(a, b, e.name)) bad.push(e.name); });
  return bad;
}
const A = XLSX.utils.encode_cell, sheetPath = { "SUMMARY": "xl/worksheets/sheet1.xml", "02 DEVICES": "xl/worksheets/sheet2.xml" };
// mimic patch.js pSet: keep cell object, change t/v
function set(ws, ref, v) {
  const c = ws[ref] || (ws[ref] = {});
  if (v instanceof Date) { c.t = "d"; c.v = v; } else if (typeof v === "number") { c.t = "n"; c.v = v; } else { c.t = "s"; c.v = v; }
  delete c.f;
  const g = XLSX.utils.decode_range(ws["!ref"]), p = XLSX.utils.decode_cell(ref);
  g.e.r = Math.max(g.e.r, p.r); g.e.c = Math.max(g.e.c, p.c); ws["!ref"] = XLSX.utils.encode_range(g);
}

(async () => {
  for (const fx of ["SAMPLE_Lenovo_Inventory_RICH.xlsx", "SAMPLE_Lenovo_Inventory_RICH_LO.xlsx", "SAMPLE_Lenovo_Inventory_RICH.xlsm"]) {
    if (!fs.existsSync(FX(fx))) { console.log("skip (missing) " + fx); continue; }
    console.log("== " + fx);
    const orig = new Uint8Array(fs.readFileSync(FX(fx)));

    // 1. zip round-trip with no changes → identical entries
    const z0 = S.zipRead(orig), rt = await S.zipWrite(z0, new Map());
    ok(untouchedIdentical(orig, rt, []).length === 0, "zip round-trip copies every entry raw");
    ok(XLSX.read(rt, { type: "array" }).SheetNames.length === XLSX.read(orig, { type: "array" }).SheetNames.length, "round-trip still parses");

    // 2. no diff → no-op
    let base = read(orig), work = read(orig);
    let r = await S.surgicalSave(orig, base, work);
    ok(r.ok && r.noChanges, "no changes → nothing written");

    // 3. realistic edit set
    base = read(orig); work = read(orig);
    const dev = work.Sheets["02 DEVICES"], last = XLSX.utils.decode_range(dev["!ref"]).e.r; // 0-based
    const hdr = {}; for (let c = 0; c <= 15; c++) { const h = dev[A({ r: 0, c })]; if (h) hdr[h.v] = c; }
    set(dev, A({ r: 1, c: hdr["Client / Organization"] }), "UPDATED & <Co> \"Ltd\"");
    set(dev, A({ r: 1, c: hdr["Warranty End"] }), new Date(2029, 4, 31));
    set(dev, A({ r: 0, c: 16 }), "Batch Code");                        // new column header (table grows)
    set(dev, A({ r: 1, c: 16 }), "B-20261006-01");
    for (let i = 1; i <= 3; i++) {                                          // appended rows (table grows)
      set(dev, A({ r: last + i, c: 0 }), "PFTEST0" + i);
      set(dev, A({ r: last + i, c: hdr["Client / Organization"] }), "  Lead space school");
      set(dev, A({ r: last + i, c: hdr["Date Delivered"] }), new Date(2026, 9, 6));
      set(dev, A({ r: last + i, c: 16 }), "B-20261006-01");
    }
    const log = XLSX.utils.aoa_to_sheet([["Timestamp", "Action", "Detail"], [new Date(2026, 9, 6, 14, 30), "invUpdateAsset", "a&b"]]);
    XLSX.utils.book_append_sheet(work, log, "EDIT LOG");
    r = await S.surgicalSave(orig, base, work);
    ok(r.ok, "surgical save ok" + (r.ok ? "" : " — " + r.reason));
    if (!r.ok) continue;
    const out = r.bytes;
    { const why = await S.verifySurgical(orig, out, work); ok(!why, "read-back verify passes" + (why ? " — " + why : "")); }
    const changed = r.stats.changedParts;
    const allowed = ["xl/worksheets/_rels/sheet2.xml.rels", "xl/worksheets/sheet2.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml", "xl/styles.xml", "xl/tables/table1.xml"];
    ok(changed.every(p => allowed.includes(p) || /^xl\/worksheets\/sheet\d+\.xml$/.test(p)), "only expected parts changed: " + changed.join(", "));
    const bad = untouchedIdentical(orig, out, changed);
    ok(bad.length === 0, "all other parts byte-identical (" + (S.zipRead(orig).entries.length - changed.filter(c => S.zipRead(orig).byName.has(c)).length) + " parts)" + (bad.length ? " BAD: " + bad : ""));
    const x0 = await partText(orig, "xl/worksheets/sheet2.xml"), x1 = await partText(out, "xl/worksheets/sheet2.xml");
    const frag = (x, tag) => (x.match(new RegExp("<" + tag + "\\b[\\s\\S]*?</" + tag + ">", "g")) || []).join("");
    ["conditionalFormatting", "dataValidations", "cols", "sheetViews"].forEach(t => ok(frag(x0, t) === frag(x1, t) && frag(x0, t).length > 0, t + " preserved verbatim"));
    ok(/<tableParts\b/.test(x1), "tableParts kept");
    const cellOf = (x, ref) => (new RegExp('<c r="' + ref + '"[^>]*?(?:/>|>[\\s\\S]*?</c>)').exec(x) || [""])[0];
    const sOf = c => (/\ss="(\d+)"/.exec(c) || [])[1];
    const b2 = A({ r: 1, c: hdr["Client / Organization"] });
    ok(sOf(cellOf(x0, b2)) === sOf(cellOf(x1, b2)), "edited cell keeps its style index s=" + sOf(cellOf(x0, b2)));
    ok(cellOf(x1, b2).includes("UPDATED &amp; &lt;Co&gt;"), "special characters escaped exactly once");
    const nd = A({ r: last + 1, c: hdr["Date Delivered"] }), od = A({ r: last, c: hdr["Date Delivered"] });
    ok(sOf(cellOf(x1, nd)) === sOf(cellOf(x0, od)), "new row inherits style from row above (date format)");
    ok(/<dimension ref="A1:Q\d+"/.test(x1) && x1.includes(`<dimension ref="A1:Q${last + 4}"`), "dimension grows to A1:Q" + (last + 4));
    const t1 = await partText(out, "xl/tables/table1.xml");
    ok(t1.includes(`ref="A1:Q${last + 4}"`), "table ref grows to cover new rows + Batch Code column");
    ok(/<tableColumn id="\d+" name="Batch Code"\/>/.test(t1) && /tableColumns count="17"/.test(t1), "tableColumn added (count 17)");
    const t0 = await partText(orig, "xl/tables/table1.xml");
    ok((/<autoFilter[^>]*ref="([^"]+)"/.exec(t0) || [])[1] ? t1.includes(`<autoFilter ref="A1:Q${last + 4}"`) : true, "table autoFilter ref matches table ref");
    const wbx = await partText(out, "xl/workbook.xml");
    ok(/fullCalcOnLoad="1"/.test(wbx), "workbook set to recalc on open");
    ok(/<sheet\b[^>]*name="EDIT LOG"/.test(wbx), "new sheet registered in workbook.xml");
    ok(/definedName[^>]*>[^<]*DEVICES/.test(wbx) || /DeviceSerials/.test(wbx), "defined names kept");
    const back = XLSX.read(out, { type: "array", cellDates: true });
    const bd = back.Sheets["02 DEVICES"];
    ok(bd[b2].v === "UPDATED & <Co> \"Ltd\"", "value reads back exactly");
    ok(bd[A({ r: last + 1, c: hdr["Client / Organization"] })].v === "  Lead space school", "leading spaces preserved");
    const we = bd[A({ r: 1, c: hdr["Warranty End"] })].v;
    ok(we instanceof Date && we.getFullYear() === 2029 && we.getMonth() === 4 && we.getDate() === 31, "date reads back as 2029-05-31");
    ok(back.Sheets["EDIT LOG"] && back.Sheets["EDIT LOG"].C2.v === "a&b", "EDIT LOG sheet readable");
    if (/\.xlsm$/.test(fx)) {
      const ct = await partText(out, "[Content_Types].xml");
      ok(ct.includes("sheet.macroEnabled.main+xml"), ".xlsm keeps macro-enabled content type");
      ok(rawSame(orig, out, "xl/vbaProject.bin"), "vbaProject.bin byte-identical");
    }
    // hyperlinks: rels untouched → no &amp;amp; drift
    const po = Object.keys(XLSX.read(orig, { type: "array" }).Sheets);
    const relsNames = S.zipRead(out).entries.map(e => e.name).filter(n => /worksheets\/_rels/.test(n));
    let amp = false; for (const n of relsNames) { if (/&amp;amp;/.test(await partText(out, n))) amp = true; }
    ok(!amp, "hyperlink targets not double-escaped");

    // 4. five consecutive saves → stable
    let cur = out, sizes = [];
    for (let i = 0; i < 5; i++) {
      const b = read(cur), w = read(cur);
      set(w.Sheets["02 DEVICES"], A({ r: 2 + i, c: hdr["Condition"] }), "For Repair");
      const rr = await S.surgicalSave(cur, b, w);
      if (!rr.ok) { ok(false, "save " + (i + 2) + " failed: " + rr.reason); break; }
      const drift = untouchedIdentical(cur, rr.bytes, rr.stats.changedParts);
      if (drift.length) { ok(false, "drift on save " + (i + 2) + ": " + drift); break; }
      cur = rr.bytes; sizes.push(cur.length);
    }
    ok(sizes.length === 5 && Math.max(...sizes) - Math.min(...sizes) < 2000, "5 more saves: no drift, size stable (" + sizes.join(", ") + ")");
    let amp2 = false; for (const n of relsNames) { if (/&amp;amp;/.test(await partText(cur, n) || "")) amp2 = true; }
    ok(!amp2, "still no double-escaped links after 6 saves");

    // 5. legacy comparison (what SheetJS CE rewrite loses)
    const legacy = XLSX.write(work, { bookType: /xlsm$/.test(fx) ? "xlsm" : "xlsx", type: "array", cellStyles: true, bookVBA: true });
    const lz = S.zipRead(new Uint8Array(legacy)), oz = S.zipRead(orig);
    const lost = oz.entries.map(e => e.name).filter(n => !lz.byName.has(n) && !/sharedStrings|calcChain|docProps\/custom/.test(n));
    const l2 = await S.zipText(lz, "xl/worksheets/sheet2.xml") || "";
    const lostFeat = ["conditionalFormatting", "dataValidations", "tableParts"].filter(t => x0.includes("<" + t) && !l2.includes("<" + t));
    console.log("    legacy SheetJS rewrite drops parts: " + (lost.join(", ") || "none") + " | sheet features: " + (lostFeat.join(", ") || "none"));
    ok(lost.length + lostFeat.length > 0, "legacy engine confirmed lossy (baseline for comparison)");

    // 6. independent validators
    if (process.env.VALIDATE) {
      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "surg-")), f = path.join(tmp, "out" + path.extname(fx));
      fs.writeFileSync(f, cur);
      let py = ""; try { py = cp.execFileSync("python3", ["-c", "import openpyxl,sys;wb=openpyxl.load_workbook(sys.argv[1]);ws=wb['02 DEVICES'];print(len(wb.sheetnames),list(ws.tables),ws.conditional_formatting.__len__(),len(ws.data_validations.dataValidation))", f]).toString().trim(); } catch (e) { py = "ERR " + e.message; }
      ok(!/ERR/.test(py), "openpyxl opens output: " + py);
      try { cp.execFileSync("soffice", ["--headless", "--convert-to", "csv", "--outdir", tmp, f], { stdio: "pipe", timeout: 120000 }); ok(fs.existsSync(path.join(tmp, "out.csv")), "LibreOffice opens + converts output"); }
      catch (e) { ok(false, "LibreOffice failed: " + e.message); }
    }
  }

  // 6b. hyperlinks (audit Reference column writes cell.l): add, update, repair
  {
    console.log("== hyperlinks");
    const fx = FX("SAMPLE_Lenovo_Inventory_RICH.xlsx"), orig = new Uint8Array(fs.readFileSync(fx));
    let base = read(orig), work = read(orig);
    const U1 = "https://example.sharepoint.com/:x:/r/sites/a/Doc.aspx?id=9&web=1&e=Q", U2 = "https://example.com/report?a=1&b=<2>";
    work.Sheets["06 PIPELINE"].B2.l = { Target: U1 };                       // sheet without any rels yet
    work.Sheets["02 DEVICES"].H3.l = { Target: U2 };                        // sheet with table+comment rels
    const po = work.Sheets["07 PURCHASE ORDER"], pc = Object.keys(po).find(k => k[0] !== "!" && po[k].l);
    const U3 = "https://example.sharepoint.com/sites/x/New.aspx?id=2&web=1"; po[pc].l = { Target: U3 };  // update existing
    let r = await S.surgicalSave(orig, base, work);
    ok(r.ok && r.stats.links === 3, "3 link edits written (" + (r.ok ? r.stats.links : r.reason) + ")");
    if (r.ok) {
      const why = await S.verifySurgical(orig, r.bytes, work); ok(!why, "read-back verify incl. hyperlinks" + (why ? " — " + why : ""));
      const back = XLSX.read(r.bytes, { type: "array" }), dec = S._decodeStable;
      ok(dec(back.Sheets["06 PIPELINE"].B2.l.Target) === U1, "new link on a sheet with no rels file");
      ok(dec(back.Sheets["02 DEVICES"].H3.l.Target) === U2, "new link on a sheet with table/comment rels (special chars)");
      ok(dec(back.Sheets["07 PURCHASE ORDER"][pc].l.Target) === U3, "existing link updated in place");
      const z1 = S.zipRead(r.bytes), relN = n => S.zipText(z1, n);
      const poRels = await relN("xl/worksheets/_rels/sheet6.xml.rels"), devRels = await relN("xl/worksheets/_rels/sheet2.xml.rels");
      ok(!/&amp;amp;/.test(poRels + devRels) && /web=1&amp;e=/.test(await relN("xl/worksheets/_rels/sheet5.xml.rels") || ""), "targets escaped exactly once");
      ok(/<Relationship[^>]*\/table"/.test(devRels) && /comments/.test(devRels), "table + comment relationships kept");
      const dx = await S.zipText(z1, "xl/worksheets/sheet2.xml");
      ok(dx.indexOf("<hyperlinks>") > dx.indexOf("</sheetData>") && dx.indexOf("<hyperlinks>") < dx.indexOf("<pageMargins") && dx.indexOf("<hyperlinks>") > dx.lastIndexOf("</dataValidations>"), "<hyperlinks> placed in schema order");
      if (process.env.VALIDATE) {
        const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "hl-")), f = path.join(tmp, "hl.xlsx"); fs.writeFileSync(f, r.bytes);
        let py = ""; try { py = cp.execFileSync("python3", ["-c", "import openpyxl,sys;wb=openpyxl.load_workbook(sys.argv[1]);print(wb['06 PIPELINE']['B2'].hyperlink.target,'|',wb['02 DEVICES']['H3'].hyperlink.target)", f]).toString().trim(); } catch (e) { py = "ERR " + e.message.slice(0, 200); }
        ok(py === U1 + " | " + U2, "openpyxl reads the links: " + py.slice(0, 120));
        try { cp.execFileSync("soffice", ["--headless", "--convert-to", "csv", "--outdir", tmp, f], { stdio: "pipe", timeout: 120000 }); ok(fs.existsSync(path.join(tmp, "hl.csv")), "LibreOffice opens it"); } catch (e) { ok(false, "LibreOffice: " + e.message); }
      }
    }
    // repair: a file damaged by the old engine (&amp;amp;) is fixed on the next save
    const z0 = S.zipRead(orig), rn = "xl/worksheets/_rels/sheet6.xml.rels";
    const bad = (await S.zipText(z0, rn)).replace(/&amp;/g, "&amp;amp;amp;");
    const damaged = await S.zipWrite(z0, new Map([[rn, bad]]));
    base = read(damaged); work = read(damaged);
    work.Sheets["02 DEVICES"].B5.v = "REPAIR TEST"; work.Sheets["02 DEVICES"].B5.t = "s";
    r = await S.surgicalSave(damaged, base, work);
    const fixed = r.ok ? await S.zipText(S.zipRead(r.bytes), rn) : "";
    ok(r.ok && r.stats.linksRepaired >= 1 && !/&amp;amp;/.test(fixed) && /&amp;web=1/.test(fixed), "double-encoded links repaired on save (" + (r.ok ? r.stats.linksRepaired : r.reason) + ")");
    const why2 = r.ok ? await S.verifySurgical(damaged, r.bytes, work) : "n/a"; ok(!why2, "verify passes after repair" + (why2 ? " — " + why2 : ""));
  }

  // 7. guard rails
  console.log("== guard rails");
  const sharedMaster = '<worksheet><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c></row></sheetData></worksheet>';
  let threw = ""; try { S._spliceSheet(sharedMaster, [{ r: 0, c: 0, cell: { t: "n", v: 5 } }], { styles: { dateXf: () => null } }); } catch (e) { threw = e.message; }
  ok(/shared-formula master/.test(threw), "refuses to overwrite a shared-formula master");
  threw = ""; try { S._spliceSheet('<worksheet><sheetData><row><c r="A1"/></row></sheetData></worksheet>', [{ r: 0, c: 0, cell: { t: "n", v: 1 } }], {}); } catch (e) { threw = e.message; }
  ok(/without r=/.test(threw), "refuses rows without r= attribute");
  const res = await S.surgicalSave(new Uint8Array([1, 2, 3]), { SheetNames: [] }, { SheetNames: [] });
  ok(!res.ok && /zip/.test(res.reason), "non-zip input rejected, nothing written");
  const ins = S._spliceSheet('<worksheet><sheetData><row r="1"><c r="A1" s="2"><v>1</v></c></row><row r="5"><c r="A5"><v>5</v></c></row></sheetData></worksheet>',
    [{ r: 2, c: 1, cell: { t: "s", v: "mid" } }, { r: 0, c: 2, cell: { t: "n", v: 3 } }], { styles: { dateXf: () => null } }).xml;
  ok(/<row r="1"><c r="A1" s="2"><v>1<\/v><\/c><c r="C1"[^>]*><v>3<\/v><\/c><\/row><row r="3">.*<\/row><row r="5">/.test(ins), "rows/cells inserted in sorted order");
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
