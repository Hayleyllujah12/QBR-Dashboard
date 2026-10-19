#!/usr/bin/env node
/* ============================================================================
 * tests/mos-parse.cjs — logic tests for js/mos-parse.js (v1.37.0 MOS tab).
 * Reads the SYNTHETIC score reports in tests/mos-fixture/ with the vendored
 * pdf.js (libs/pdfjs) exactly as the browser does, then checks parsing,
 * file-name hints, roster join, dedupe, attempt numbering, filters and rates.
 * Regenerate fixtures: python tests/make-mos-fixture.py
 * Usage: node tests/mos-parse.cjs
 * ========================================================================== */
"use strict";
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const FX = path.join(__dirname, "mos-fixture");
const M = require(path.join(ROOT, "js/mos-parse.js"));
const _warn = console.warn, _log = console.log;
console.warn = () => {};                       // pdf.js warns about optional node canvas
const pdfjs = require(path.join(ROOT, "libs/pdfjs/pdf.min.js"));
pdfjs.GlobalWorkerOptions.workerSrc = path.join(ROOT, "libs/pdfjs/pdf.worker.min.js");

let pass = 0, fail = 0;
function ok(c, m, d) { if (c) { pass++; _log("  ✓ " + m); } else { fail++; _log("  ✗ " + m + (d !== undefined ? "  (" + JSON.stringify(d) + ")" : "")); } }
function walk(d) { let o = []; for (const f of fs.readdirSync(d)) { const p = path.join(d, f); if (fs.statSync(p).isDirectory()) o = o.concat(walk(p)); else if (/\.pdf$/i.test(f)) o.push(p); } return o.sort(); }

(async () => {
  console.log("== pure helpers");
  ok(M.cleanTitle("Microsoft Excel (Microsoft 365 Apps and O ce 2019)") === "Microsoft Excel (Microsoft 365 Apps and Office 2019)", "ligature fix 'O ce'");
  ok(M.cleanTitle("Microsoft PowerPoint (Oce 2019)") === "Microsoft PowerPoint (Office 2019)", "ligature fix 'Oce'");
  ok(M.cleanTitle("AZ-900: Microsoft Azure Fundamentals Registration ID: ITS-1") === "AZ-900: Microsoft Azure Fundamentals", "Registration ID stripped");
  ok(M.examFamily("Microsoft Excel Expert (Microsoft 365 Apps)") === "Excel Expert", "family Excel Expert");
  ok(M.examFamily("Microsoft PowerPoint (Office 2019)") === "PowerPoint", "family PowerPoint");
  ok(M.examFamily("SC-900: Microsoft Security, Compliance, and Identity Fundamentals") === "SC-900", "family by exam code");
  ok(M.parseDate("September 9, 2025") === "2025-09-09", "date parse");
  ok(M.schoolYear("2025-06-01") === "SY 2025-2026" && M.schoolYear("2026-05-31") === "SY 2025-2026", "school year June–May");
  const fm = M.fileMeta("Course 400/03 - St. Example/SY 25-26/Exam 1234.2 Retake GR11 11B.pdf");
  ok(fm.school === "St. Example" && fm.examNo === "1234.2" && fm.examBase === "1234" && fm.retakeFile && fm.sectionHint === "11B" && fm.gradeHint === "11", "fileMeta hints", fm);
  ok(M.fileMeta("Exam 1119 Results.pdf", "Typed School").school === "Typed School", "school from the input when the path has no school folder");
  ok(M.fileMeta("Exam 1119 Results.pdf").school === "Unassigned school", "unassigned fallback");

  console.log("== synthetic score reports via vendored pdf.js");
  const attempts = [], files = [];
  for (const f of walk(FX)) {
    const rel = path.relative(FX, f).split(path.sep).join("/");
    const meta = M.fileMeta(rel);
    const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(f)), isEvalSupported: false, verbosity: 0 }).promise;
    let n = 0;
    for (let i = 1; i <= doc.numPages; i++) {
      const r = M.parseScorePage((await (await doc.getPage(i)).getTextContent()).items);
      if (!r || r._err) continue;
      Object.assign(r, meta, { page: i }); attempts.push(r); n++;
    }
    files.push({ rel, n });
  }
  ok(attempts.length === 7, "7 score pages read (certificate PDF skipped)", attempts.length);
  ok(files.find(f => /Certificates/.test(f.rel)).n === 0, "certificates file has no score pages");
  const a1 = attempts.find(a => a.ref === "90000001");
  ok(a1 && a1.name === "Alex Cruz Demo" && a1.email === "alex.demo@example.com" && a1.certiportId === "alexdemo01", "candidate name / email / Certiport ID");
  ok(a1 && !JSON.stringify(a1).includes("Sample Street"), "address lines are never captured");
  ok(a1 && a1.examTitle === "Microsoft Excel (Microsoft 365 Apps and Office 2019)" && a1.exam === "Excel", "title across two lines + ligature fix");
  ok(a1 && a1.score === 875 && a1.required === 700 && a1.outcome === "PASS" && a1.date === "2026-03-03", "score / required / outcome / date");
  ok(a1 && a1.skills.length === 5 && a1.skills[3].area === "Perform Operations by using Formulas and Functions" && a1.skills[3].pct === 75, "5 skill areas with %");
  ok(attempts.some(a => a.ref === "90000002" && a.outcome === "FAIL" && a.score === 600), "split 'EXAM S|CORE REPORT' page still parsed (fail 600)");
  ok(attempts.some(a => a.ref === "90000003" && a.exam === "PowerPoint" && a.skills.length === 5), "split 'Exam refe|rence' + gap-spaced 'SECTION ANALYSIS'");
  const d = attempts.find(a => a.ref === "90000005");
  ok(d && d.score === 1000 && d.required === 700 && d.sectionHint === "12B" && d.school === "Sample Academy", "missing 'Required Score' label → unlabeled numbers; section hint from file name");
  ok(attempts.some(a => a.ref === "90000006" && a.score === 0 && a.outcome === "FAIL"), "score 0 is kept (not treated as missing)");

  console.log("== roster + model");
  const roster = M.parseRosterRows("11A", [
    ["No.", "", "", "NAME", "Email", "EXCEL", "Gmetrix"],
    ["1", "", "2600001", "DEMO, Alex, Cruz", "a.demo@school.example", "", "g1"],
    ["2", "", "2600002", "SAMPLE, Bea, Lopez", "bea.sample@example.com", "", "g2"],
  ]).concat(M.parseRosterRows("11B", [
    ["NAME", "Email"], ["TEST, Carl, Reyes", ""], ["TEST, Carla, Reyes", ""]
  ]));
  ok(M.parseRosterRows("SUMMARY", [["NAME"]]).length === 0, "non-section sheets ignored");
  ok(roster.length === 4 && roster[0].studentNo === "2600001" && roster[0].section === "11A", "roster rows (unlabeled student-no column found)");
  roster.forEach(s => { s.school = "Demo High School"; });
  const m = M.buildModel(attempts, roster);
  ok(m.rows.length === 6 && m.duplicates === 1, "duplicate page (same Exam reference #) dropped", [m.rows.length, m.duplicates]);
  const byRef = {}; m.rows.forEach(r => { byRef[r.ref] = r; });
  ok(byRef["90000001"].match === "name" && byRef["90000001"].section === "11A" && byRef["90000001"].studentNo === "2600001", "name match: 'First Middle Surname' ↔ 'SURNAME, First, Middle'");
  ok(byRef["90000002"].match === "email" && byRef["90000002"].section === "11A", "email match wins");
  ok(byRef["90000003"].match === "name" && byRef["90000003"].section === "11B", "first-name token disambiguates Carl vs Carla");
  ok(byRef["90000005"].match === "none" && byRef["90000005"].section === "12B", "other school: no roster match, file-name section kept");
  ok(byRef["90000002"].attemptNo === 1 && byRef["90000004"].attemptNo === 2 && byRef["90000004"].isRetake && byRef["90000004"].pass, "retake numbered as attempt 2 for the same student + exam");
  ok(byRef["90000002"].retakeEligible && byRef["90000002"].retakeUsed && !byRef["90000004"].retakeEligible && byRef["90000006"].retakeEligible && !byRef["90000006"].retakeUsed,
    "retake rule: first-time fail = eligible (Bea retook, Eve not yet); 2nd take is not eligible again");
  ok(M.buildModel([{ ref: "1", name: "A B", email: "a@x.example", exam: "Excel", date: "2026-01-01", score: 700, outcome: "FAIL", school: "S" }], []).rows[0].pass, "exactly 700 passes (score wins over a stray outcome)");
  const m2 = M.buildModel(attempts, roster, { "90000005": { section: "12C" } });
  ok(m2.rows.find(r => r.ref === "90000005").section === "12C", "manual section override");

  console.log("== filters + summaries");
  const s = M.summarize(m.rows);
  ok(s.attempts === 6 && s.passed === 4 && s.failed === 2 && Math.abs(s.passRate - 4 / 6) < 1e-9, "totals + pass rate", s);
  ok(s.firstAttempts === 5 && Math.abs(s.firstPassRate - 3 / 5) < 1e-9, "first-attempt pass rate excludes the retake");
  ok(s.high === 1000 && s.low === 0, "highest / lowest");
  ok(M.filterAttempts(m.rows, { school: "Demo High School" }).length === 4, "filter: school");
  ok(M.filterAttempts(m.rows, { school: "Demo High School", section: "11A" }).length === 3, "filter: section");
  ok(M.filterAttempts(m.rows, { outcome: "FAIL" }).length === 2, "filter: pass/fail");
  ok(M.filterAttempts(m.rows, { min: 700, max: 900 }).length === 2, "filter: score range 700–900");
  ok(M.filterAttempts(m.rows, { examNo: "9001" }).length === 4, "filter: Exam No. includes its .1 retake session");
  ok(M.filterAttempts(m.rows, { year: "2025" }).length === 2 && M.filterAttempts(m.rows, { month: "2026-04" }).length === 1, "filter: year / month");
  ok(M.filterAttempts(m.rows, { sy: "SY 2025-2026" }).length === 6, "filter: school year");
  ok(M.filterAttempts(m.rows, { q: "bea" }).length === 2, "filter: student search");
  ok(M.filterAttempts(m.rows, { firstOnly: true }).length === 5, "filter: first attempts only");
  const mon = M.groupRate(m.rows, r => r.month);
  ok(JSON.stringify(mon.map(g => [g.key, g.attempts, g.passed])) === JSON.stringify([["2025-09", 2, 1], ["2026-03", 3, 2], ["2026-04", 1, 1]]), "monthly pass-rate groups", mon);
  const yr = M.groupRate(m.rows, r => r.year);
  ok(yr.length === 2 && yr[1].key === "2026" && Math.abs(yr[1].passRate - 0.75) < 1e-9, "yearly pass-rate groups");
  const sk = M.skillAreas(m.rows);
  ok(sk[0].avg <= sk[sk.length - 1].avg && sk[0].exam === "Excel", "weakest skill areas sorted ascending");
  ok(M.bandOf(699) === "<700" && M.bandOf(700) === "700–799" && M.bandOf(1000) === "900–1000", "score bands");

  console.log("== Certiport results export (.xlsx)");
  const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
  const xb = require("./mos-export-fixture.cjs")(XLSX);
  const xwb = XLSX.read(xb, { type: "buffer", cellDates: true });
  const xrows = XLSX.utils.sheet_to_json(xwb.Sheets[xwb.SheetNames[0]], { header: 1, defval: "", raw: true });
  ok(M.parseExportRows([["NAME", "Email"], ["X, Y", ""]]) === null, "a roster sheet is not taken for an export");
  const ex = M.parseExportRows(xrows, "results.xlsx");
  ok(ex && ex.attempts.length === 4 && ex.header >= 2 && ex.skipped === 3, "4 rows read; title + 3 footer lines skipped", ex && [ex.attempts.length, ex.header, ex.skipped]);
  const xs = JSON.stringify(ex.attempts);
  ok(!xs.includes("2011-01-01") && !xs.includes("TEST-VOUC"), "date of birth and voucher codes are never read");
  const xa = ex.attempts;
  ok(xa[0].name === "Alex Demo" && xa[0].certiportId === "alexdemo01" && xa[0].studentIdExport === "", "ALL-CAPS name title-cased; '{invalid or no ID}' dropped");
  ok(xa[1].email === "gina.ong@example.com" && xa[1].studentIdExport === "26-0099" && xa[1].proctor === "Proctor One" && xa[1].station === "LAB-2", "email / student ID / proctor / station");
  ok(xa[1].date === "2026-10-08" && xa[1].time === "11:40" && xa[1].timeUsed === 1818 && xa[1].exam === "Word" && xa[1].score === 940 && xa[1].outcome === "PASS", "date, time, time used, exam, score, result");
  ok(xa[3].outcome === "IN PROGRESS" && xa[3].score === null, "'In Progress' kept without a score");
  ok(xa[1].school === "Demo Education Network" && xa[1].schoolFromCenter, "school defaults to the Test Center");
  const roster2 = roster.concat(M.parseRosterRows("11C", [["NAME", "Email"], ["ONG, Gina, Reyes", "g.ong@school.example"]]).map(st => Object.assign(st, { school: "Demo High School" })));
  const mx = M.buildModel(attempts.concat(xa), roster2);
  const bx = {}; mx.rows.forEach(r => { bx[r.ref] = r; });
  ok(mx.rows.length === 9 && mx.merged === 1 && mx.duplicates === 1, "PDF page + export row with the same Exam reference # → one attempt", [mx.rows.length, mx.merged, mx.duplicates]);
  ok(bx["90000001"].source === "pdf+export" && bx["90000001"].proctor === "Proctor One" && bx["90000001"].skills.length === 5 && bx["90000001"].school === "Demo High School" && bx["90000001"].time === "09:15",
    "merged attempt: proctor + time from the export, skill areas + school folder from the PDF");
  ok(bx["91000001"].school === "Demo High School" && bx["91000001"].section === "11C" && bx["91000001"].match === "name", "export row: roster school + section win over the Test Center");
  ok(bx["91000002"].school === "Demo Education Network" && bx["91000002"].batch === "DEN-20261008-01", "no roster match → Test Center school; auto batch ACR-YYYYMMDD-01");
  ok(bx["90000001"].batch === "9001", "PDF Exam No. stays the batch when known");
  const sx = M.summarize(mx.rows);
  ok(sx.pending === 1 && sx.attempts === 8, "in-progress row excluded from scored attempts", sx);
  ok(M.filterAttempts(mx.rows, { proctor: "Proctor Two" }).length === 2 && M.filterAttempts(mx.rows, { outcome: "IN PROGRESS" }).length === 1, "filters: proctor, in progress");
  ok(M.filterAttempts(mx.rows, { examNo: "DEN-20261008-01" }).length === 2, "filter by auto batch code");
  ok(M.groupRate(mx.rows, r => r.proctor).some(g => g.key === "Proctor Two" && g.attempts === 1), "group rates skip in-progress rows");
  ok(M.buildModel(xa, [], {}, { acronyms: { "Demo Education Network": "dnet" } }).rows.every(r => !r.date || r.batch.indexOf("DNET-") === 0), "profile acronym overrides the automatic one");
  ok(M.acronym("Sisters of St Paul of Chartres Education Ministry") === "SSPCEM" && M.acronym("Our Lady of Peace School (OLPS)") === "OLPS", "school acronyms");
  ok(M.parseAnyDate("10/8/2026 1:05 PM").date === "2026-10-08" && M.parseAnyDate("10/8/2026 1:05 PM").time === "13:05" && M.parseAnyDate(new Date(2026, 9, 8, 11, 40)).time === "11:40", "date formats (text, US, Date object)");

  _log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
