// Format-safe save, end-to-end in the real app (headless Chromium, file://).
// Rich fixture (styles/CF/DV/table/chart/comment/hyperlink/macros) → edits via the
// real Inventory API → Save to linked (mocked) file → inspect the written bytes.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-surgical.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const S = require(path.join(APPDIR, "js/xlsx-surgical.js"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

async function session(b, file, engine, rounds) {
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(e => { try { localStorage.setItem("qbr-remember", "0"); localStorage.setItem("qbr-save-engine", e); } catch (x) {} }, engine);
  await p.setInputFiles("#file-input", file); await p.waitForTimeout(2500);
  const name = path.basename(file), b64 = fs.readFileSync(file).toString("base64");
  const res = await p.evaluate(async ({ name, b64, rounds }) => {
    const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    let bytes = bin, mtime = 1700000000000; const outs = [];
    const mk = () => new File([bytes], name, { lastModified: mtime });
    const f0 = APP.files.find(f => f.name === name).blob;
    const handle = { name, kind: "file", queryPermission: async () => "granted", requestPermission: async () => "granted", getFile: async () => mk(),
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), close: async () => { bytes = parts[0]; mtime += 1000; } }; } };
    QBR._fsLinks = [{ fp: QBR.fpOf(name, f0), kinds: ["assets"], name, size: bin.length, lastModified: mtime, handle }];
    const inv = APP.model.inventory, a = inv.assets[0], results = [];
    for (let i = 0; i < rounds; i++) {
      if (i === 0) {
        QBR.invUpdateAsset(a.key, { client: "FMT-SAFE & <TEST> SCHOOL", batch: "B-20261006-01" });
        QBR.invAddDeployment({ sn: a.sn, sq: "SQ-FMT-SAFE", delivered: "2026-10-06" });
        QBR.invIntake([{ sn: "PFFMTSAFE1", client: "FMT-SAFE SCHOOL", model: "ThinkPad E14", delivered: "2026-10-06" }]);
      } else QBR.invUpdateAsset(a.key, { cond: "For Repair " + i });
      const r = await QBR.fsSaveKind("assets");
      results.push({ mode: r.mode, engine: r.engine || null, reason: r.reason || null, changed: r.stats && r.stats.changedParts });
      let s = ""; for (let k = 0; k < bytes.length; k += 32768) s += String.fromCharCode.apply(null, bytes.subarray(k, k + 32768));
      outs.push(btoa(s));
    }
    return { results, outs, sn: a.sn };
  }, { name, b64, rounds });
  await p.close();
  return Object.assign(res, { errs, orig: new Uint8Array(fs.readFileSync(file)), outs: res.outs.map(x => new Uint8Array(Buffer.from(x, "base64"))) });
}
const raw = (u8, n) => { const e = S.zipRead(u8).byName.get(n); return e ? Buffer.from(u8.subarray(e.dataStart, e.dataStart + e.csize)).toString("base64") + e.crc : null; };

(async () => {
  const b = await chromium.launch();
  for (const fx of ["SAMPLE_Lenovo_Inventory_RICH.xlsx", "SAMPLE_Lenovo_Inventory_RICH.xlsm", "SAMPLE_Lenovo_Inventory_RICH_LO.xlsx"]) {
    const file = path.join(__dirname, fx);
    console.log("== " + fx + " (format-safe, 3 saves)");
    const r = await session(b, file, "format-safe", 3);
    ok(r.results.every(x => x.mode === "file" && x.engine === "format-safe"), "all 3 saves written by the format-safe engine (" + r.results.map(x => x.mode + (x.reason ? ":" + x.reason : "")).join(", ") + ")");
    const o = r.orig, z0 = S.zipRead(o), last = r.outs[r.outs.length - 1];
    const keep = z0.entries.map(e => e.name).filter(n => !/worksheets\/sheet\d+\.xml$|workbook\.xml$|workbook\.xml\.rels$|Content_Types|tables\/|styles\.xml/.test(n));
    const drift = keep.filter(n => raw(o, n) !== raw(last, n));
    ok(drift.length === 0, keep.length + " untouched parts byte-identical after 3 saves (charts, drawings, comments, theme, vba, rels…)" + (drift.length ? " DRIFT: " + drift : ""));
    const sx0 = await S.zipText(z0, "xl/worksheets/sheet2.xml"), sx = await S.zipText(S.zipRead(last), "xl/worksheets/sheet2.xml");
    const frag = (x, t) => (x.match(new RegExp("<" + t + "\\b[\\s\\S]*?</" + t + ">", "g")) || []).join("");
    ok(frag(sx0, "conditionalFormatting") === frag(sx, "conditionalFormatting"), "conditional formatting intact");
    ok(frag(sx0, "dataValidations") === frag(sx, "dataValidations"), "data validation intact");
    const tb = await S.zipText(S.zipRead(last), "xl/tables/table1.xml");
    ok(/name="Batch Code"/.test(tb), "Excel table gained the Batch Code column");
    const wb = XLSX.read(last, { type: "array" }), rows = XLSX.utils.sheet_to_json(wb.Sheets["02 DEVICES"], { header: 1, defval: "" });
    const tref = /ref="([A-Z]+\d+:[A-Z]+\d+)"/.exec(tb)[1];
    ok(XLSX.utils.decode_range(tref).e.r === rows.length - 1, "table covers the new intake row (" + tref + ", " + rows.length + " rows)");
    ok(rows.some(x => x.includes(r.sn) && x.includes("FMT-SAFE & <TEST> SCHOOL") && x.includes("B-20261006-01")), "asset edit + batch code written");
    ok(rows.some(x => x.includes("PFFMTSAFE1")), "intake row appended");
    ok(XLSX.utils.sheet_to_json(wb.Sheets["04 RAKSO INV."], { header: 1, defval: "" }).some(x => x.includes("SQ-FMT-SAFE")), "deployment row appended to 04 RAKSO INV.");
    ok(rows.some(x => x.includes(r.sn) && x.includes("For Repair 2")), "later saves applied (round 3 value present)");
    if (/xlsm$/.test(fx)) ok(raw(o, "xl/vbaProject.bin") === raw(last, "xl/vbaProject.bin"), "macros byte-identical");
    ok(!(await Promise.all(S.zipRead(last).entries.filter(e => /_rels/.test(e.name)).map(e => S.zipText(S.zipRead(last), e.name)))).some(t => /&amp;amp;/.test(t)), "no double-escaped hyperlinks");
    ok(r.errs.length === 0, "no page errors" + (r.errs.length ? ": " + r.errs.join(" | ") : ""));
  }
  console.log("== legacy engine still selectable (opt-out)");
  const L = await session(b, path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"), "legacy", 1);
  ok(L.results[0].mode === "file" && !L.results[0].engine, "qbr-save-engine=legacy uses the old rewrite path");
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
