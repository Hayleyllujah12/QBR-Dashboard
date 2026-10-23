// MOS Reconcile pure core (js/mos-reconcile.js) — Node tests, synthetic data only.
// Usage: node tests/mos-reconcile.cjs
"use strict";
const path = require("path");
const XLSX = require(path.join(__dirname, "..", "libs", "xlsx.full.min.js"));
const C = require(path.join(__dirname, "..", "js", "mos-reconcile.js"));
const FX = require(path.join(__dirname, "mos-reconcile-fixture.cjs"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

function exportAOA() {
  const wb = XLSX.read(FX.buildExport(XLSX), { type: "array" });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: true });
}

/* roster parsing */
const block = C.parseClassRoster(FX.rosterAOA(), "fixture");
ok(!!block, "class roster detected (name + EXAM VOUCHER header)");
ok(block.students.length === 6, "6 roster students");
ok(block.cols.name === 2 && block.cols.voucher === 7 && block.cols.section === 5 && block.cols.score === 10 && block.cols.result === 11, "columns found by header, not position");
ok(block.students[1].oldScore === 940 && block.students[1].oldResult === "PASSED", "existing score/result read");
ok(block.students[0].oldScore === null && block.students[0].section === "11A", "blank score stays blank (never zero)");

const pasted = C.parseClassRoster(C.parseTSV(FX.rosterTSV()), "paste");
ok(!!pasted && pasted.students.length === 6 && pasted.students[5].voucher === "VCH-DUMMY-999", "TSV paste parses identically");

/* header-position independence: voucher column moved to the front */
const shuffled = FX.rosterAOA().map(r => [r[7], r[2], r[5], r[10], r[11]]);
const shBlock = C.parseClassRoster(shuffled, "shuffled");
ok(!!shBlock && shBlock.students[2].voucher === "VCH-DUMMY-003" && shBlock.students[2].oldScore === 500, "columns survive reordering");

/* export parsing */
const ex = C.parseCertExport(exportAOA());
ok(!!ex && ex.rows.length === 7, "export: 7 voucher rows (footers skipped)");
ok(ex.rows[0].voucher === "VCH-DUMMY-001" && ex.rows[0].ref === "92000001" && ex.rows[0].score === 875 && ex.rows[0].exam === "Excel", "export row fields (voucher/ref/score/exam)");
ok(ex.rows[6].exam === "PowerPoint" && ex.rows[6].result === "Pass", "PowerPoint rows parsed too");

/* reconcile — the six canonical outcomes, in roster order */
const res = C.reconcile([block], ex.rows, {});
const st = res.rows.map(r => r.status);
ok(JSON.stringify(st) === JSON.stringify(["filled", "verified", "changed", "no-result", "conflict", "candidate"]), "statuses in roster order: " + st.join(","));
ok(res.rows.map(r => r.student.name).join("|") === block.students.map(s => s.name).join("|"), "output follows the roster's name order exactly");
ok(res.counts.total === 6 && res.counts.verified === 1 && res.counts.filled === 1 && res.counts.changed === 1 && res.counts.noResult === 1 && res.counts.conflict === 1 && res.counts.candidate === 1, "counts reconcile");
ok(res.counts.notInRoster === 4 && res.counts.exports === 7, "4 unmatched export rows before confirmation (candidate + 3 strangers)");
ok(res.rosterExam === "Excel", "roster exam derived from voucher hits (Excel)");
ok(/score 500→640/.test(res.rows[2].flag) && !/result/.test(res.rows[2].flag), "CHANGED flag shows only the real diff (score), equal results are not a diff");
ok(res.rows[0].finalScore === 875 && res.rows[0].finalResult === "PASSED", "FILLED takes Certiport values");
ok(res.rows[4].finalScore === 880 && res.rows[4].status === "conflict", "CONFLICT keeps roster values, flagged for manual verification");

/* name candidate: accept and reject */
const acc = C.reconcile([block], ex.rows, { "0:5": "92000006" });
ok(acc.rows[5].status === "confirmed" && acc.rows[5].finalScore === 720 && /NAME MATCH — CONFIRMED/.test(acc.rows[5].flag), "accepted candidate fills, flagged NAME MATCH — CONFIRMED");
ok(acc.counts.notInRoster === 3, "accepted row leaves Not in Roster (3 remain)");
const rej = C.reconcile([block], ex.rows, { "0:5": "rejected" });
ok(rej.rows[5].status === "no-result" && /rejected/.test(rej.rows[5].flag), "rejected candidate becomes NO CERTIPORT RESULT (rejected)");

/* name validation on voucher pairs */
ok(C.nameCheck("KIM, Fiona", { first: "Kim", last: "Fiona", alt: "" }).swapped === true, "swapped First/Last detected");
ok(C.nameCheck("KIM, Fiona", { first: "Fiona", last: "Kim", alt: "" }).consistent === true, "normal order consistent");
ok(C.nameCheck("KIM, Fiona", { first: "Fiona", last: "Santos", alt: "" }).consistent === false, "different surname flagged inconsistent");

/* duplicate voucher inside the roster */
const dupBlock = C.parseClassRoster(FX.rosterAOA().concat([[7, "Regular", "COPY, Alex Cruz", "11", "STEM", "11A", "demostud007", "VCH-DUMMY-001", "", "", "", ""]]), "dup");
const dupRes = C.reconcile([dupBlock], ex.rows, {});
ok(dupRes.rows[6].status === "duplicate" && dupRes.counts.duplicate === 1, "second use of a roster voucher flagged DUPLICATE, first still matches");
ok(dupRes.rows[0].status === "filled", "first voucher occurrence unaffected");

/* stacked sections inside one worksheet */
const stk = C.parseClassRosterBlocks(FX.stackedAOA(), "wb · Excel");
ok(stk.length === 2, "stacked worksheet -> 2 blocks (one per section)");
ok(stk[0].students.length === 3 && stk[1].students.length === 2, "students split 3 + 2, headers/titles never become students");
ok(stk[0].students.every(x => x.section === "RM 2- ST.PAUL (HUMSS)") && stk[1].students.every(x => x.section === "RM 3- ST.JOHN (STEM)"), "section taken from the block title row when the column is blank");
ok(/RM 2- ST\.PAUL/.test(stk[0].source), "block label carries its section title");
const mid = C.parseClassRosterBlocks(FX.midTitleAOA(), "wb · G10");
ok(mid.length === 2 && mid[0].students.length === 2 && mid[1].students.length === 2 && mid[1].students[0].section === "RM 4- ST.MARK (ABM)", "mid-sheet title (no repeated header) starts a new section");
const one = C.parseClassRosterBlocks(FX.rosterAOA(), "wb · CLASS ROSTER");
ok(one.length === 1 && one[0].students.length === 6, "single-header worksheet still parses as one block");

/* filterPreview: Results-style preview filtering (export never filters) */
const fp = (f) => C.filterPreview(res, f);
ok(fp({ section: "11A" }).length === 6 && fp({ section: "ZZ" }).length === 0, "filter: section");
ok(fp({ exam: "Excel" }).length === 6 && fp({ exam: "PowerPoint" }).length === 0, "filter: exam (no-result rows use the roster exam)");
ok(fp({ flag: "changed" }).length === 1 && fp({ flag: "changed" })[0].student.name.indexOf("TEST") === 0, "filter: flag = changed -> Carlos only");
ok(fp({ flag: "candidate" }).length === 1 && fp({ flag: "conflict" }).length === 1, "filter: flag = candidate / conflict");
ok(fp({ q: "kim" }).length === 1 && fp({ q: "VCH-DUMMY-003" }).length === 1, "filter: search by name or voucher");
ok(fp({ min: 900 }).length === 1 && fp({ min: 900 })[0].finalScore === 940, "filter: score >= 900 -> Bea only");
ok(fp({ min: 600, max: 900 }).length === 2, "filter: score 600-900 -> Alex + Carlos (unmatched rows excluded)");
ok(fp({ sort: "score-desc" })[0].finalScore === 940 && fp({ sort: "name" })[0].student.name.indexOf("DEMO") === 0 && fp({ sort: "name" })[5].student.name.indexOf("TRIAL") === 0, "sort: score desc / name A-Z");
ok(fp({})[0].idx === 0 && fp({ sort: "name" })[0].idx != null, "rows keep their roster position (idx) under filters");

console.log("\nmos-reconcile: " + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
