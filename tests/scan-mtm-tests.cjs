/* MTM auto-detect tests (Node, no browser). Run: node tests/scan-mtm-tests.cjs
   Sam 2026-10-11 (15 checks) + review additions (data identifiers, shape guard). */
"use strict";
const fs = require("fs"), vm = require("vm"), path = require("path");

const sb = {
  console, APP: { model: {} },
  document: { createElement: () => ({ style: {}, appendChild() {} }) },
  localStorage: { _d: {}, getItem(k) { return this._d[k] || null; }, setItem(k, v) { this._d[k] = String(v); } },
};
sb.window = sb; sb.globalThis = sb;
vm.createContext(sb);
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
vm.runInContext(fs.readFileSync(path.join(APPDIR, "js", "scan.js"), "utf8"), sb, { filename: "scan.js" });

let pass = 0, fail = 0;
function t(name, cond) {
  if (cond) { pass++; }
  else { fail++; console.error("FAIL:", name); }
}
const eq = (a, b) => a === b;

// Robin's label: (S)SN PF62SDPW, (31P)MTM 21M4S5VD00, (1S) 21M4S5VD00PF62SDPW
const SN = "PF62SDPW", MTM = "21M4S5VD00", ONE_S = "21M4S5VD00PF62SDPW";

t("splitMtm extracts head", eq(vm.runInContext(`scanSplitMtm("${ONE_S}")`, sb), MTM));
t("splitMtm null on plain serial", eq(vm.runInContext(`scanSplitMtm("${SN}")`, sb), null));
t("looksLikeMtm true", vm.runInContext(`scanLooksLikeMtm("${MTM}")`, sb) === true);
t("looksLikeMtm false on serial", vm.runInContext(`scanLooksLikeMtm("${SN}")`, sb) === false);
t("looksLikeMtm false on 1S", vm.runInContext(`scanLooksLikeMtm("${ONE_S}")`, sb) === false);

// prefix-correlation: 31P prefixes 1S
t("pickMtm prefix-correlation", eq(
  vm.runInContext(`scanPickMtm(${JSON.stringify([SN, MTM, ONE_S])}, "${SN}")`, sb), MTM));
// 1S only -> head fallback
t("pickMtm 1S-only fallback", eq(
  vm.runInContext(`scanPickMtm(${JSON.stringify([ONE_S])}, "")`, sb), MTM));
// standalone 31P only
t("pickMtm standalone 31P", eq(
  vm.runInContext(`scanPickMtm(${JSON.stringify([MTM])}, "")`, sb), MTM));
// serial excluded even if it prefixes (should not happen, but guard)
t("pickMtm excludes picked serial", eq(
  vm.runInContext(`scanPickMtm(${JSON.stringify([SN, ONE_S])}, "${SN}")`, sb), MTM));
// no MTM present
t("pickMtm empty when none", eq(
  vm.runInContext(`scanPickMtm(${JSON.stringify([SN])}, "${SN}")`, sb), ""));

// end-to-end merge on Robin's full label
const m = vm.runInContext(
  `scanMergeFindings(${JSON.stringify([SN, MTM, ONE_S])}, null, "SN: ${SN} MTM: ${MTM}")`, sb);
t("merge serial is SN", eq(m.serial, SN));
t("merge mtm is MTM", eq(m.mtm, MTM));
t("merge mtm != serial", m.mtm !== m.serial);

// 1S-only label: serial from tail, mtm from head
const m2 = vm.runInContext(`scanMergeFindings(${JSON.stringify([ONE_S])}, null, "")`, sb);
t("1S-only serial from tail", eq(m2.serial, SN));
t("1S-only mtm from head", eq(m2.mtm, MTM));

// review additions (2026-10-11)
const run = e => vm.runInContext(e, sb);
t("1S data identifier stripped before split", eq(run(`scanSplitMtm("1S${ONE_S}")`), MTM));
t("31P data identifier stripped", eq(run(`scanPickMtm(["31P${MTM}"], "${SN}")`), MTM));
t("1S-prefixed + 31P-prefixed label", eq(run(`scanPickMtm(["${SN}", "31P${MTM}", "1S${ONE_S}"], "${SN}")`), MTM));
t("all-digit 10-char value is not an MTM", run(`scanLooksLikeMtm("0196380123")`) === false);
t("serial-length value prefixing a longer code is not an MTM", eq(run(`scanPickMtm(["${SN}", "${SN}XYZ12"], "")`), ""));
const m3 = run(`scanMergeFindings(${JSON.stringify(["1S" + ONE_S])}, null, "")`);
t("1S-prefixed only: mtm from head", eq(m3.mtm, MTM));

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
