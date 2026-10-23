// SYNTHETIC fixtures for the MOS Reconcile tests — every name, voucher and ID is invented.
// rosterAOA(): Pedro's class-roster layout (one sheet, Section column, EXAM VOUCHER column).
// buildExport(XLSX): a Certiport "Results (7) Plus" export with a DISTINCT voucher per row
// (unlike tests/mos-export-fixture.cjs, which reuses one voucher).
"use strict";

function rosterAOA() {
  return [
    ["No.", "Student Type", "Display Name", "Grade Level", "Strand", "Section", "GMETRIX", "EXAM VOUCHER", "DAY 1", "DAY 2", "EXAM SCORE", "RESULT"],
    [1, "Regular", "DEMO, Alex Cruz", "11", "STEM", "11A", "demostud001", "VCH-DUMMY-001", "PRESENT", "PRESENT", "", ""],
    [2, "Regular", "SAMPLE, Bea Lopez", "11", "STEM", "11A", "demostud002", "VCH-DUMMY-002", "PRESENT", "PRESENT", 940, "PASSED"],
    [3, "Regular", "TEST, Carlos Reyes", "11", "STEM", "11A", "demostud003", "VCH-DUMMY-003", "PRESENT", "PRESENT", 500, "FAILED"],
    [4, "Regular", "EXAMPLE, Dana Yu", "11", "STEM", "11A", "demostud004", "VCH-DUMMY-004", "PRESENT", "ABSENT", "", ""],
    [5, "Regular", "TRIAL, Eli Fox", "11", "STEM", "11A", "demostud005", "VCH-DUMMY-005", "PRESENT", "ABSENT", 880, "PASSED"],
    [6, "Regular", "KIM, Fiona", "11", "STEM", "11A", "demostud006", "VCH-DUMMY-999", "PRESENT", "PRESENT", "", ""]
  ];
}

function rosterTSV() {
  return rosterAOA().map(function (r) { return r.join("\t"); }).join("\n");
}

function buildExport(XLSX) {
  const H = ["Test Center ID", "Test Center Name", "First Name", "", "Last Name", "Alternate Name", "", "Student / Employee ID", "Username",
    "Exam Reference Number", "Exam", "Date of Birth", "Program Name", "Exam Level", "Certification Name", "Language", "Exam Date", "Score",
    "Result", "Time Used", "Proctor Name", "Station", "Exam Group", "Teacher", "Payment Type", "Voucher / License"];
  const C = ["99999999", "Demo Education Network"];
  const row = (first, last, sid, user, ref, exam, cert, date, score, result, voucher) =>
    [...C, first, "", last, "", "", sid, user, ref, exam, "2011-01-01", "Microsoft Office Specialist", "Certification", cert, "English",
      date, score, result, 1800, "Proctor One", "LAB-1", "", "", "Voucher", voucher];
  const XL = "Microsoft Excel (Microsoft 365 Apps)", PP = "Microsoft PowerPoint (Microsoft 365 Apps)";
  const aoa = [[], ["Results (7) Plus"], [], H,
    row("Alex", "Demo", "", "alex.demo@example.com", 92000001, XL, "Excel Associate (Microsoft 365 Apps)", "2026-10-08 09:15:00", 875, "Pass", "VCH-DUMMY-001"),
    row("Bea", "Sample", "", "bea.sample@example.com", 92000002, XL, "Excel Associate (Microsoft 365 Apps)", "2026-10-08 09:20:00", 940, "Pass", "VCH-DUMMY-002"),
    row("Carlos", "Test", "", "carlos.test@example.com", 92000003, XL, "Excel Associate (Microsoft 365 Apps)", "2026-10-08 09:25:00", 640, "Fail", "VCH-DUMMY-003"),
    row("Fiona", "Kim", "", "fiona.kim@example.com", 92000006, XL, "Excel Associate (Microsoft 365 Apps)", "2026-10-08 09:30:00", 720, "Pass", "VCH-DUMMY-006"),
    row("Gustavo", "Reyes", "", "gustavo.reyes@example.com", 92000007, XL, "Excel Associate (Microsoft 365 Apps)", "2026-10-08 09:35:00", 810, "Pass", "VCH-DUMMY-007"),
    row("Hana", "Park", "", "hana.park@example.com", 92000008, PP, "PowerPoint Associate (Microsoft 365 Apps)", "2026-10-09 10:00:00", 555, "Fail", "VCH-DUMMY-008"),
    row("Ivan", "Chen", "", "ivan.chen@example.com", 92000009, PP, "PowerPoint Associate (Microsoft 365 Apps)", "2026-10-09 10:05:00", 902, "Pass", "VCH-DUMMY-009"),
    [], ["User: 0"], ["Execution Time: 10/16/2026 8:21:15 AM"], ["© 2026 Sample footer (synthetic test fixture)"]];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Results (7) Plus");
  return Buffer.from(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}


// One worksheet, TWO sections stacked as blocks: [title row] header, students (Section column
// left blank - the title carries it), [title row] header, students.
function stackedAOA() {
  const H = ["No.", "Student Type", "Display Name", "Grade Level", "Strand", "Section", "GMETRIX", "EXAM VOUCHER", "EXAM SCORE", "RESULT"];
  const st = (n, name, gm, v) => [n, "Regular", name, "11", "STEM", "", gm, v, "", ""];
  return [
    ["GRADE 11"],
    ["G11-A ALPHA (HUMSS)"], H,
    st(1, "NORTH, Ana Reyes", "demostud101", "VCH-DUMMY-101"),
    st(2, "SOUTH, Ben Cruz", "demostud102", "VCH-DUMMY-102"),
    st(3, "EAST, Cora Lim", "demostud103", "VCH-DUMMY-103"),
    ["G11-B BETA (STEM)"], H,
    st(1, "WEST, Dan Sy", "demostud104", "VCH-DUMMY-104"),
    st(2, "FIELD, Eva Tan", "demostud105", "VCH-DUMMY-105")
  ];
}
// Same, but the second section has NO repeated header - just a title row between student groups.
function midTitleAOA() {
  const H = ["No.", "Student Type", "Display Name", "Grade Level", "Strand", "Section", "GMETRIX", "EXAM VOUCHER", "EXAM SCORE", "RESULT"];
  const st = (n, name, gm, v) => [n, "Regular", name, "10", "ABM", "", gm, v, "", ""];
  return [H,
    st(1, "HILL, Fay Yu", "demostud201", "VCH-DUMMY-201"),
    st(2, "LAKE, Gil Ong", "demostud202", "VCH-DUMMY-202"),
    ["G10-A GAMMA (ABM)"],
    st(3, "POND, Hal Diaz", "demostud203", "VCH-DUMMY-203"),
    st(4, "REEF, Ida Ko", "demostud204", "VCH-DUMMY-204")];
}
// Workbook with two worksheets: "Excel" (stacked, 2 sections) + "Grade 10" (single header).
function buildRosterWorkbook(XLSX) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(stackedAOA()), "Excel");
  const H = ["No.", "Student Type", "Display Name", "Grade Level", "Strand", "Section", "GMETRIX", "EXAM VOUCHER", "EXAM SCORE", "RESULT"];
  const g10 = [H,
    [1, "Regular", "VALE, Jay Ruiz", "10", "STEM", "G10-C DELTA (STEM)", "demostud301", "VCH-DUMMY-301", "", ""],
    [2, "Regular", "MESA, Kay Ford", "10", "STEM", "G10-C DELTA (STEM)", "demostud302", "VCH-DUMMY-302", "", ""]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(g10), "Grade 10");
  return Buffer.from(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}

module.exports = { rosterAOA: rosterAOA, rosterTSV: rosterTSV, buildExport: buildExport, stackedAOA: stackedAOA, midTitleAOA: midTitleAOA, buildRosterWorkbook: buildRosterWorkbook };
