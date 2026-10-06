// Linked "Save to Excel" safety test (2026-10-05 .xlsm corruption fix). Headless Chromium, file://.
// Uses a mocked File System Access handle so the written bytes can be inspected.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-save.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const FIX = { xlsm: path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsm"), xlsx: path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx") };
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const MAIN = { xlsx: "spreadsheetml.sheet.main+xml", xlsm: "sheet.macroEnabled.main+xml" };

async function run(b, ext, externalChange) {
  console.log(`== ${ext}${externalChange ? " (file changed in Excel first)" : ""}`);
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", FIX[ext]); await p.waitForTimeout(2500);
  const name = path.basename(FIX[ext]); const b64 = fs.readFileSync(FIX[ext]).toString("base64");
  const res = await p.evaluate(async ({ name, b64, externalChange }) => {
    const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    let bytes = bin, mtime = 1700000000000, writes = 0;
    const mk = () => new File([bytes], name, { lastModified: mtime });
    const f0 = APP.files.find(f => f.name === name).blob;
    const handle = {
      name, kind: "file",
      queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => mk(),
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), close: async () => { bytes = parts[0]; mtime += 1000; writes++; } }; },
    };
    const fp = QBR.fpOf(name, f0);
    QBR._fsLinks = [{ fp, kinds: ["assets"], name, size: externalChange ? -1 : bin.length, lastModified: externalChange ? 1 : mtime, handle }];
    // baseline the link to the mocked file metadata
    if (!externalChange) { QBR._fsLinks[0].size = bin.length; QBR._fsLinks[0].lastModified = mtime; }
    const inv = QBR._invModel || (typeof invModel === "function" ? invModel() : null);
    const a = inv.assets[0];
    QBR.invUpdateAsset(a.key, { client: "UI-SAVE TEST SCHOOL", batch: "B-20261005-07" });
    QBR.invAddDeployment({ sn: a.sn, sq: "SQ-SAVE-TEST", delivered: "2026-10-01" });
    let dl = 0; const oldW = XLSX.writeFile; XLSX.writeFile = function () { dl++; };
    const r = await QBR.fsSaveKind("assets");
    XLSX.writeFile = oldW;
    let out = ""; for (let i = 0; i < bytes.length; i += 32768) out += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
    return { mode: r.mode, reason: r.reason || null, writes, dl, sn: a.sn, b64: btoa(out) };
  }, { name, b64, externalChange });
  if (externalChange) {
    // 2026-10-06: a file changed outside the dashboard is now MERGED (xlsx-merge.js), not refused
    ok(res.writes === 1 && res.mode === "file", "file changed outside the dashboard → edits merged into it (mode " + res.mode + ")");
    ok(res.dl === 0, "no separate copy downloaded");
  } else {
    ok(res.mode === "file" && res.writes === 1, "saved into the linked file (mode " + res.mode + (res.reason ? ", " + res.reason : "") + ")");
    const out = Buffer.from(res.b64, "base64");
    const zip = XLSX.CFB.read(out, { type: "buffer" });
    const ct = Buffer.from(XLSX.CFB.find(zip, "/[Content_Types].xml").content).toString();
    ok(ct.includes(MAIN[ext]), "content type matches ." + ext + " (Excel will open it)");
    if (ext === "xlsm") ok(!!XLSX.CFB.find(zip, "/xl/vbaProject.bin"), "macros (vbaProject.bin) preserved");
    else ok(!XLSX.CFB.find(zip, "/xl/vbaProject.bin"), "no macro part in .xlsx");
    const orig = XLSX.read(fs.readFileSync(FIX[ext]), { type: "buffer" }), wb = XLSX.read(out, { type: "buffer" });
    ok(orig.SheetNames.every(n => wb.SheetNames.includes(n)), "all original sheets kept (" + orig.SheetNames.length + "; now " + wb.SheetNames.join(", ") + ")");
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["02 DEVICES"], { header: 1, defval: "" });
    ok(rows.some(r => r.includes(res.sn) && r.includes("UI-SAVE TEST SCHOOL")), "edit written into 02 DEVICES for " + res.sn);
    const hdr = rows[0], bc = hdr.indexOf("Batch Code");
    const rk = XLSX.utils.sheet_to_json(wb.Sheets["04 RAKSO INV."], { header: 1, defval: "" });
    ok(rk.some(r => r.includes("SQ-SAVE-TEST") && r.includes(res.sn)), "bulk SQ appended as a row in 04 RAKSO INV.");
    ok(bc === hdr.filter(h => h !== "").length - 1 && rows.some(r => r.includes(res.sn) && r[bc] === "B-20261005-07"), "'Batch Code' column added at the end and filled (col " + bc + ")");
  }
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await p.close();
}
(async () => {
  const b = await chromium.launch();
  await run(b, "xlsm", false); await run(b, "xlsx", false); await run(b, "xlsm", true);
  // verifier rejects a mismatched file (the original bug)
  const p = await b.newPage(); await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  const why = await p.evaluate(() => {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["a"]]), "S");
    return QBR._fsVerifyBytes(XLSX.write(wb, { bookType: "xlsx", type: "array" }), "xlsm", wb);
  });
  console.log("== verifier"); ok(!!why, "xlsx bytes for an .xlsm name are refused (\"" + why + "\")");
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
