// SYNTHETIC Certiport "Results (n) Plus" export for the MOS tests — every name, email, ID, date of birth
// and voucher code is invented. Mirrors the real layout: blank row, title row, blank row, header row
// (incl. the empty helper columns after First / Alternate Name), data, then the 4 footer lines.
// require("./mos-export-fixture.cjs")(XLSX) → Buffer (.xlsx)
module.exports = function build(XLSX) {
  const H = ["Test Center ID", "Test Center Name", "First Name", "", "Last Name", "Alternate Name", "", "Student / Employee ID", "Username",
    "Exam Reference Number", "Exam", "Date of Birth", "Program Name", "Exam Level", "Certification Name", "Language", "Exam Date", "Score",
    "Result", "Time Used", "Proctor Name", "Station", "Exam Group", "Teacher", "Payment Type", "Voucher / License"];
  const C = ["99999999", "Demo Education Network"];
  const row = (first, last, sid, user, ref, exam, cert, date, score, result, tu, proctor, station) =>
    [...C, first, "", last, "", "", sid, user, ref, exam, "2011-01-01", "Microsoft Office Specialist", "Certification", cert, "English",
      date, score, result, tu, proctor, station, "", "", "Voucher", "TEST-VOUC-HER0-0000"];
  const XL = "Microsoft Excel (Microsoft 365 Apps and Office 2019)", W = "Microsoft Word (Microsoft 365 Apps)";
  const aoa = [[], ["Results (4) Plus"], [], H,
    row("ALEX", "DEMO", "{invalid or no ID}", "alexdemo01", 90000001, XL, "Excel Associate (Office 2019)", "2026-03-03 09:15:00", 875, "Pass", 1800, "Proctor One", "LAB-1"),
    row("Gina", "Ong", "26-0099", "gina.ong@example.com", 91000001, W, "Word Associate (Microsoft 365 Apps)", "2026-10-08 11:40:00", 940, "Pass", 1818, "Proctor One", "LAB-2"),
    row("Hugo", "Lim", "", "hugo.lim@example.com", 91000002, W, "Word Associate (Microsoft 365 Apps)", "2026-10-08 11:40:00", 640, "Fail", 2990, "Proctor Two", "LAB-3"),
    row("Ivy", "Sy", "{invalid or no ID}", "ivy.sy@example.com", 91000003, W, "", "2026-10-08 02:03:00", "", "In Progress", "", "Proctor Two", "LAB-4"),
    [], ["User: 0"], ["Execution Time: 10/16/2026 8:21:15 AM"], ["© 2026 Sample footer (synthetic test fixture)"]];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Results (4) Plus");
  return Buffer.from(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
};
