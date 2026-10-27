// Excel write contract — static guard rails (docs/EXCEL_WRITE_CONTRACT.md §2.4). Node, no browser.
//  1. Only js/persist.js opens a file for writing (createWritable). Modules still on their own write
//     code are listed in LEGACY with the phase that migrates them; each migration deletes its entry.
//     A LEGACY entry that no longer writes fails too, so the list can only shrink honestly.
//  2. Every op recorded with QBR.journalRecord(kind, "<op>", …) has a cell writer in PATCH_OPS
//     (js/patch.js) and a replay function QBR.<op> — otherwise the edit could be recorded but never saved.
// Usage: node tests/write-contract.cjs
"use strict";
const fs = require("fs"), path = require("path");
const JS = path.join(__dirname, "..", "js");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

const OWNER = "persist.js";
const LEGACY = {               // file → phase that moves it onto the shared save path
  "mos-certs.js": "Phase 2 (MOS Certificate Log)",
  "domain-edit.js": "Phase 3 (Domain registration)",
  "trainers.js": "Phase 4 (Trainers tracker)",
};
const files = fs.readdirSync(JS).filter(f => /\.js$/.test(f));
const src = {};
files.forEach(f => { src[f] = fs.readFileSync(path.join(JS, f), "utf8"); });
// strip // and /* */ comments so a mention in prose doesn't count
const code = s => s.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:"'\\])\/\/.*$/gm, "$1");

console.log("== 1. who writes files");
const writers = files.filter(f => /\bcreateWritable\s*\(/.test(code(src[f])));
ok(writers.includes(OWNER), OWNER + " owns the write (createWritable)");
const stray = writers.filter(f => f !== OWNER && !LEGACY[f]);
ok(stray.length === 0, "no other module writes files directly" + (stray.length ? " — found in: " + stray.join(", ") + " (use QBR.fsSaveKind; see docs/EXCEL_WRITE_CONTRACT.md)" : ""));
Object.keys(LEGACY).forEach(f => {
  const still = writers.includes(f);
  ok(still, "legacy writer " + f + " still listed for " + LEGACY[f] + (still ? "" : " — it no longer writes: delete its LEGACY entry"));
});
console.log("  · legacy writers left: " + Object.keys(LEGACY).length + " (" + Object.keys(LEGACY).join(", ") + ")");

console.log("== 2. every recorded op can be saved and replayed");
const patch = code(src["patch.js"] || "");
const opsBlock = (patch.match(/var PATCH_OPS\s*=\s*\{([\s\S]*?)\};/) || [])[1] || "";
const patchOps = new Set((opsBlock.match(/^\s*([A-Za-z_$][\w$]*)\s*:/gm) || []).map(s => s.trim().replace(/:$/, "")));
ok(patchOps.size > 0, "PATCH_OPS parsed (" + patchOps.size + " ops)");
const recorded = [];
files.forEach(f => {
  const re = /journalRecord\(\s*"([\w-]+)"\s*,\s*(?:op|"([\w$]+)")/g; let m;
  const c = code(src[f]);
  while ((m = re.exec(c)) !== null) recorded.push({ file: f, kind: m[1], op: m[2] || null });
});
ok(recorded.length > 0, "journalRecord call sites found (" + recorded.length + ")");
// dynamic op names (journalRecord(kind, op, …)) are checked through their known set below
const named = recorded.filter(r => r.op);
const all = code(Object.values(src).join("\n"));
named.forEach(r => {
  ok(patchOps.has(r.op), r.kind + "/" + r.op + " (" + r.file + ") has a PATCH_OPS cell writer");
  ok(new RegExp("QBR\\." + r.op + "\\s*=\\s*(async\\s+)?function").test(all), r.kind + "/" + r.op + " has a replay function QBR." + r.op);
});
// audit.js records via a helper (journalRecord("audit", op, args)); its ops are the audit* PATCH_OPS
const auditOps = [...patchOps].filter(o => /^audit/.test(o));
auditOps.forEach(op => ok(new RegExp("QBR\\." + op + "\\s*=\\s*(async\\s+)?function").test(all), "audit/" + op + " has a replay function QBR." + op));

console.log("RESULT: " + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
