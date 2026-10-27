// Folder load per-file skip (2026-10-26) — Node vm harness for js/fsfolder.js.
// Drives the REAL QBR.fsOpenFolder with fake File System Access handles:
// a workbook that is listed by scan() but throws NotFoundError at read time
// (the OneDrive sync/move case the requester reported) must be skipped and NAMED,
// while the rest of the batch still loads. No real files, no real people.
// Usage: node tests/fs-perfile-skip.cjs
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

const NOTFOUND = "A requested file or directory could not be found at the time an operation was processed.";
function makeEntry(name, failRead) {
  return {
    kind: "file", name,
    async getFile() {
      if (failRead) { const e = new Error(NOTFOUND); e.name = "NotFoundError"; throw e; }
      return { name, size: 10, lastModified: 0 };
    },
  };
}
function makeDir(entries) {
  return {
    name: "TestFolder",
    async queryPermission() { return "granted"; },
    async requestPermission() { return "granted"; },
    async isSameEntry() { return false; },
    async *values() { for (const e of entries) yield e; },
  };
}

async function runCase(label, entries, want) {
  // Fresh vm per case so _fsDirs / _fsLinks do not leak between cases.
  let dirQueue = [makeDir(entries)];
  let alerts = [];
  let loadedFiles = [];
  let toastEl = { hidden: true, textContent: "", className: "", setAttribute() {} };
  const sandbox = {
    console,
    setTimeout, clearTimeout,
    Promise, Map, Error, String, Array, Object, Date, Math, JSON,
    alert: (m) => { alerts.push(String(m)); },
    window: {},
    document: {
      readyState: "complete",
      getElementById: (id) => (id === "fo-toast" ? toastEl : null),
      createElement: () => ({ hidden: true, textContent: "", className: "", setAttribute() {}, appendChild() {}, classList: { toggle() {} } }),
      body: { appendChild() {}, classList: { add() {}, remove() {}, toggle() {} } },
      addEventListener() {}, removeEventListener() {},
      dispatchEvent() {},
    },
    localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    APP: { files: [] },
    fileKey: (n) => String(n || "").trim().toLowerCase(),
    persistIdbGet: async () => null,
    persistIdbPut: async () => {},
    fsLinksSave: async () => {},
    fsMemFile: async (f) => ({ file: { name: f.name }, bytes: new Uint8Array([1, 2]) }),
    fsHashBytes: () => "hash",
    loadItems: async (files) => { loadedFiles = files.map((f) => f.name); return true; },
  };
  sandbox.window.showDirectoryPicker = async () => dirQueue.shift();
  sandbox.window.QBR = {};
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  const src = fs.readFileSync(path.join(__dirname, "..", "js", "fsfolder.js"), "utf8");
  vm.runInContext(src, sandbox);
  const QBR = sandbox.window.QBR;
  QBR._folderAutoPick = (books) => books.map((_, i) => i);
  QBR.fpOf = (n) => "fp:" + n;
  QBR._kindByFp = {};
  const result = await QBR.fsOpenFolder();
  ok(result === want.result, label + ": fsOpenFolder returns " + want.result + " (got " + result + ")");
  ok(JSON.stringify(loadedFiles) === JSON.stringify(want.loaded),
    label + ": loadItems got [" + loadedFiles.join(", ") + "] (want [" + want.loaded.join(", ") + "])");
  if (want.alertContains) ok(alerts.some((a) => a.indexOf(want.alertContains) >= 0),
    label + ": alert names '" + want.alertContains + "' (alerts: " + JSON.stringify(alerts) + ")");
  else ok(alerts.length === 0, label + ": no alert (alerts: " + JSON.stringify(alerts) + ")");
  if (want.noteContains) ok(toastEl.textContent.indexOf(want.noteContains) >= 0,
    label + ": final note contains '" + want.noteContains + "' (note: " + JSON.stringify(toastEl.textContent) + ")");
  if (want.noteNotContains) ok(toastEl.textContent.indexOf(want.noteNotContains) < 0,
    label + ": final note does not contain '" + want.noteNotContains + "' (note: " + JSON.stringify(toastEl.textContent) + ")");
}

(async () => {
  // Case 1 — the requester's case: middle workbook vanishes at read time.
  await runCase("mixed", [
    makeEntry("GoodOne.xlsx", false),
    makeEntry("BadOne.xlsx", true),
    makeEntry("GoodTwo.xlsx", false),
  ], { result: true, loaded: ["GoodOne.xlsx", "GoodTwo.xlsx"], noteContains: "failed: BadOne.xlsx" });
  // Case 2 — every picked workbook fails: old behaviour preserved, but named.
  await runCase("all-fail", [
    makeEntry("BadOne.xlsx", true),
    makeEntry("BadTwo.xlsx", true),
  ], { result: false, loaded: [], alertContains: "BadOne.xlsx" });
  // Case 3 — regression: everything reads, note unchanged apart from count.
  await runCase("all-ok", [
    makeEntry("GoodOne.xlsx", false),
    makeEntry("GoodTwo.xlsx", false),
  ], { result: true, loaded: ["GoodOne.xlsx", "GoodTwo.xlsx"], noteContains: "Loaded & linked", noteNotContains: "failed:" });
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error("HARNESS ERROR", e); process.exit(1); });
