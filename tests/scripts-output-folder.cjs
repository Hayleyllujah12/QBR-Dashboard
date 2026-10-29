// Output-folder convention (2026-10-29): every built-in PowerShell script asks where to save its
// output up front through the ONE shared function Select-QbrOutputFolder (js/scripts.js
// OUTPUT_FOLDER_FN), and a custom script's "# QBR_OUTPUT_FOLDER" line expands to it on Copy.
// Fails if a new built-in script forgets to ask. Zero install (plain Node).
// Usage: node tests/scripts-output-folder.cjs
"use strict";
const path = require("path");
global.window = {};
const store = {};
global.localStorage = { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: k => { delete store[k]; } };
require(path.join(__dirname, "..", "js", "scripts.js"));
const Q = window.QBR;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  \u2713 " + m); } else { fail++; console.log("  \u2717 " + m); } };
const DEF = /^function Select-QbrOutputFolder \{/m, CALL = /Select-QbrOutputFolder -Default /g;

// strip quoted strings + comments, then check (), {}, [] balance — catches a broken paste of the function
function balanced(ps) {
  const code = ps.split("\n").map(l => l.replace(/'[^']*'/g, "''").replace(/"(?:[^"]|"")*"/g, '""').replace(/(^|\s)#.*$/, "$1")).join("\n");
  const st = [], pair = { ")": "(", "}": "{", "]": "[" };
  for (const ch of code) {
    if ("({[".includes(ch)) st.push(ch);
    else if (pair[ch]) { if (st.pop() !== pair[ch]) return false; }
  }
  return st.length === 0;
}

console.log("== built-in scripts ask where to save");
const builtins = Q.scriptList().filter(s => s.builtin);
ok(builtins.length >= 2, builtins.length + " built-in scripts");
for (const s of builtins) {
  const t = Q.scriptCopyText(s.id, null).text;
  const defs = t.split("\n").filter(l => DEF.test(l)).length, calls = (t.match(CALL) || []).length;
  ok(defs === 1, s.title + ": defines Select-QbrOutputFolder exactly once");
  ok(calls >= 1, s.title + ": asks where to save (Select-QbrOutputFolder -Default …)");
  ok(t.search(DEF) < t.search(CALL), s.title + ": the function is defined before it is called");
  const outside = t.replace(/function Select-QbrOutputFolder \{[\s\S]*?\n\}\n/, "");
  ok(!/Read-Host[^\n]*(output|save|folder)/i.test(outside) && !/Read-Value[^\n]*folder/i.test(outside), s.title + ": no second, home-made folder prompt");
  ok(!/\$\{OUTPUT_FOLDER_FN\}|# QBR_OUTPUT_FOLDER\s*$/m.test(t), s.title + ": no unexpanded placeholder");
  ok(!/<#|#>/.test(t), s.title + ": still paste-safe (no block comments)");
  ok(balanced(t), s.title + ": brackets balance");
  // Only the report CSV is written: one Export-Csv, nothing else on disk (temp downloads are deleted).
  const body = outside.split("\n").filter(l => !/^\s*#/.test(l)).join("\n");
  const exports = body.match(/Export-Csv/g) || [];
  ok(exports.length === 1 && /\$Output\s*\|\s*Export-Csv -(Literal)?Path \$OutputPath/.test(body), s.title + ": writes exactly one file, the report CSV (" + exports.length + " Export-Csv)");
  ok(!/Set-Content|Out-File|Start-Transcript|Tee-Object|>\s*\$/.test(body), s.title + ": no other file writes (Set-Content / Out-File / transcript)");
  ok(!/JsonLogFile\s*=\s*Join-Path|raw_\$|_FAILED|\.jsonl/.test(body), s.title + ": no run-log, JSONL, FAILED or raw_ output");
  const tmp = body.match(/-OutputFilePath \$(\w+)/g) || [];
  ok(tmp.every(x => /GetTempPath/.test(body) && new RegExp("Remove-Item -LiteralPath \\$" + x.split("$")[1]).test(body)), s.title + ": Graph CSV downloads go to a temp file that is deleted");
  ok(/M365_(Storage|Usage)Report_\$\{[Pp]eriod\}_\$TimeStamp\.csv/.test(body), s.title + ": report file name M365_<kind>Report_<period>_<timestamp>.csv");
}

console.log("== the shared function");
const fn = (Q.scriptCopyText("usage", null).text.match(/function Select-QbrOutputFolder \{[\s\S]*?\n\}\n/) || [""])[0];
ok(fn && balanced(fn), "function body extracted and balanced");
ok(/param\(\[Parameter\(Mandatory\)\]\[string\]\$Default/.test(fn), "takes -Default (mandatory) and -Purpose");
ok(/Read-Host/.test(fn) && /IsNullOrWhiteSpace\(\$ans\)/.test(fn) && /\$path = \$Default/.test(fn), "Enter keeps the default folder");
ok(/-ieq 'P'/.test(fn) && /FolderBrowserDialog/.test(fn) && /GetApartmentState\(\)/.test(fn), "P opens a folder window (STA-guarded, falls back to typing)");
ok(/':q'/.test(fn) && /return \$null/.test(fn), ":q cancels (returns null, nothing written)");
ok(/New-Item -Path \$path -ItemType Directory/.test(fn), "creates a missing folder");
ok(/WriteAllText\(\$probe/.test(fn), "checks the folder can be written before the run starts");
ok(/\[OK\] Saving to: \$full/.test(fn), "prints the folder it will save to");
ok(!/`/.test(fn), "no backticks (the bodies live in JS template literals)");

console.log("== custom scripts: # QBR_OUTPUT_FOLDER expands on Copy");
Q.scriptUpsert({ id: "t-custom", title: "Dummy custom", body: "Write-Host 'x'\r\n# QBR_OUTPUT_FOLDER\r\n$f = Select-QbrOutputFolder -Default 'C:\\Temp'\r\n" });
const ct = Q.scriptCopyText("t-custom", null).text;
ok(ct.split("\n").filter(l => DEF.test(l)).length === 1 && !/# QBR_OUTPUT_FOLDER\s*$/m.test(ct), "placeholder line replaced by the shared function");
ok(ct.indexOf("Write-Host 'x'") === 0, "rest of the script unchanged");
Q.scriptUpsert({ id: "t-plain", title: "Dummy plain", body: "Write-Host 'y'" });
ok(Q.scriptCopyText("t-plain", null).text === "Write-Host 'y'", "scripts without the placeholder are copied as-is");
Q.scriptUpsert({ id: "t-twice", title: "Dummy twice", body: Q.scriptGet("usage").body + "\n# QBR_OUTPUT_FOLDER\n" });
ok(Q.scriptCopyText("t-twice", null).text.split("\n").filter(l => DEF.test(l)).length === 1, "never inserted twice");
const tid = "11111111-2222-3333-4444-555555555555";
Q.scriptUpsert({ id: "t-tenant", title: "Dummy tenant", body: "$T = 'QBR_TENANT_ID'\n# QBR_OUTPUT_FOLDER\n" });
const tt = Q.scriptCopyText("t-tenant", tid);
ok(tt.prefilled && tt.text.includes(tid) && DEF.test(tt.text), "works together with the tenant-ID pre-fill");

console.log(pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
