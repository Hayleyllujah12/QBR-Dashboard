// Hyperlink targets must survive read -> write -> read cycles (lesson 2026-10-06; audit 2026-10-31 L1/M3).
// SheetJS keeps targets entity-encoded and re-encodes on write; QBR.sanitizeHyperlinkTargets (excel-loader.js)
// decodes them first. Also checks every plain XLSX.write of a user workbook calls it. Synthetic data only.
// Usage: node tests/hyperlink-sanitize.cjs
"use strict";
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
global.window = global;
global.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
require(path.join(APPDIR, "js/excel-loader.js"));
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x) + "]" : "")); };
const URL0 = "https://example.com/a?x=1&y=2&z=<3>";

ok(typeof QBR.sanitizeHyperlinkTargets === "function" && typeof QBR.decodeEntitiesStable === "function", "sanitizer lives in the engine (no audit.js dependency)");
ok(QBR.decodeEntitiesStable("a&amp;amp;amp;b&lt;c&gt;") === "a&b<c>", "multi-encoded entities collapse");

function cycle(sanitize, n) {
  let wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet([["link"]]); ws.A1.l = { Target: URL0 }; XLSX.utils.book_append_sheet(wb, ws, "S");
  let buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  for (let i = 0; i < n; i++) { wb = XLSX.read(buf, { type: "buffer" }); if (sanitize) QBR.sanitizeHyperlinkTargets(wb); buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" }); }
  return XLSX.read(buf, { type: "buffer" }).Sheets.S.A1.l.Target;   // raw, as SheetJS reads it
}
const once = cycle(true, 0);   // the file as first written: one correct XML encoding
ok(cycle(true, 4) === once, "4 save cycles with the sanitizer keep the link byte-identical", cycle(true, 4));
ok(QBR.decodeEntitiesStable(once) === URL0, "and it still decodes to the original URL");
ok(cycle(false, 2) !== once, "control: without it SheetJS adds an &amp; layer per save (proves the test bites)", cycle(false, 2));

// every module that writes a workbook READ from a user file sanitizes first
const SRC = f => fs.readFileSync(path.join(APPDIR, "js", f), "utf8");
[["mos-certs.js", /var plain = function \(\) \{[\s\S]{0,200}sanitizeHyperlinkTargets\(work\)[\s\S]{0,120}XLSX\.write\(work/],
 ["trainers.js", /sanitizeHyperlinkTargets\(work\);[\s\S]{0,160}XLSX\.write\(work/],
 ["persist.js", /sanitizeHyperlinkTargets\(wb\);[\s\S]{0,40}const bytes = XLSX\.write\(wb/],
 ["audit.js", /sanitizeHyperlinkTargets\(wb\);[\s\S]{0,20}const out = XLSX\.write\(wb/]].forEach(([f, re]) => ok(re.test(SRC(f)), f + ": sanitizes before writing a read workbook"));

console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
