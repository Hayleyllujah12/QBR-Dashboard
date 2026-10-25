// MOS completion certificates — logic tests for js/mos-completion.js.
// Pure helpers only (name format, exam titles, ordinals, given text, school
// year, eligibility, grouping). All names invented dummy data; no real people.
// Usage: node tests/mos-completion.cjs
"use strict";
const path = require("path");
const M = require(path.join(__dirname, "..", "js", "mos-completion.js"));
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

// Name format: LAST, FIRST MI — comma form trusted verbatim (compound first
// names survive); non-comma form converted; missing MI omitted.
ok(M.formatCertName("Dela Cruz, Juan P.") === "DELA CRUZ, JUAN P.", "comma form kept");
ok(M.formatCertName("Dela Cruz, Juan") === "DELA CRUZ, JUAN", "comma form, no MI");
ok(M.formatCertName("Dela Cruz, Juan Santos M.") === "DELA CRUZ, JUAN SANTOS M.", "comma form, compound first name kept");
ok(M.formatCertName("Juan Dela Cruz") === "CRUZ, JUAN D.", "non-comma converted");
ok(M.formatCertName("Maria Santos") === "SANTOS, MARIA", "non-comma, no middle");
ok(M.formatCertName("Ana B. Reyes") === "REYES, ANA B.", "non-comma, initial kept");
ok(M.formatCertName("Madonna") === "MADONNA", "single token");
ok(M.formatCertName("") === "", "empty name");

// Exam titles.
ok(M.examTitle("Excel") === "Microsoft Office Specialist \u2013 Excel Associate (Office 365 Apps)", "Excel title");
ok(M.examTitle("Word") === "Microsoft Office Specialist \u2013 Word Associate (Office 365 Apps)", "Word title");
ok(M.examTitle("PowerPoint") === "Microsoft Office Specialist \u2013 PowerPoint Associate (Office 365 Apps)", "PowerPoint title");
ok(M.examTitle("Excel Expert") === "Microsoft Office Specialist \u2013 Excel Expert (Office 365 Apps)", "Expert title (no Associate)");

// Ordinals + given text.
[["01", "1st"], ["02", "2nd"], ["03", "3rd"], ["04", "4th"], ["11", "11th"], ["12", "12th"], ["13", "13th"], ["21", "21st"], ["22", "22nd"], ["23", "23rd"], ["31", "31st"]]
  .forEach(function (pair) { ok(M.ordinal(pair[0]) === pair[1], "ordinal " + pair[0] + " -> " + pair[1]); });
ok(M.givenText("2026-03-27") === "Given this 27th day of March, 2026.", "given text Mar 27");
ok(M.givenText("2026-03-01") === "Given this 1st day of March, 2026.", "given text Mar 1");
ok(M.givenText("") === "", "given text empty");

// School year + eligibility.
ok(M.syText("SY 2025-2026") === "2025\u20132026", "SY prefix stripped, en dash");
ok(M.syText("") === "", "SY empty");
ok(M.isEligibleRow({ noRetakesLeft: true, exam: "Excel" }), "eligible: failed-twice Excel");
ok(M.isEligibleRow({ noRetakesLeft: true, exam: "Excel Expert" }), "eligible: failed-twice Expert");
ok(!M.isEligibleRow({ noRetakesLeft: true, exam: "AZ-900" }), "excluded: non-MOS exam");
ok(!M.isEligibleRow({ noRetakesLeft: false, exam: "Excel" }), "excluded: retakes left");
ok(!M.isEligibleRow({ noRetakesLeft: true, exam: "Outlook" }), "excluded: Outlook (not in MOS-app scope)");

// Grouping: schools alphabetical, exams in Excel/Expert/Word/PowerPoint order,
// names alphabetical within an exam.
const certs = [
  { name: "SANTOS, MARIA", family: "Word", school: "Second High" },
  { name: "REYES, JOSE P.", family: "Excel", school: "Demo High" },
  { name: "CRUZ, ANA", family: "Excel Expert", school: "Demo High" },
  { name: "LIM, PEDRO", family: "Word", school: "Demo High" },
  { name: "TAN, LIZA", family: "Excel", school: "Demo High" },
];
const g = M.groupCerts(certs);
ok(g.map(function (x) { return x.school; }).join(",") === "Demo High,Second High", "groups: schools alphabetical");
ok(g[0].exams.map(function (x) { return x.family; }).join(",") === "Excel,Excel Expert,Word", "groups: exam order Excel/Expert/Word");
ok(g[0].exams[0].certs.map(function (x) { return x.name; }).join(",") === "REYES, JOSE P.,TAN, LIZA", "groups: names alphabetical within exam");
ok(g[0].exams[2].certs.length === 1 && g[1].exams[0].certs.length === 1, "groups: counts");

// Signature geometry (HTML cert).
ok(M.geomOf(null).size === 100 && M.geomOf(null).dx === 0 && M.geomOf(null).dy === 0, "geomOf defaults");
const gmCl = M.geomOf({ size: 500, dx: 2, dy: -9 });
ok(gmCl.size === 150 && gmCl.dx === 0.5 && gmCl.dy === -0.5, "geomOf clamps to 50-150% / +-0.5in");
ok(M.sigImgStyle({ size: 150, dx: 0.1, dy: -0.05 }) === "height:1.35in;transform:translate(0.1in,-0.05in)", "sigImgStyle 150% = height + image shift");
ok(M.sigImgStyle(null) === "height:0.9in;transform:translate(0in,0in)", "sigImgStyle default = today's size, no shift");
const sigH = { n1: "DUMMY ONE", t1: "Dummy Title", n2: "", t2: "", org: "Dummy Org", s1: "data:image/png;base64,AAA", s2: "", g1: { size: 150, dx: 0.1, dy: -0.05 } };
const htmlH = M.certHtml({ name: "X", family: "Excel", title: "T", sy: "2025–2026", given: "Given this 1st day of January, 2026.", partner: "", badge: "" }, sigH);
ok(htmlH.includes('class="mos-comp-sig"'), "certHtml sig block carries no shift (names stay put)");
ok(!htmlH.includes('class="mos-comp-sig" style="transform'), "certHtml sig block has no transform");
ok(htmlH.includes('style="height:1.35in;transform:translate(0.1in,-0.05in)"'), "certHtml sig img honors size + shift");
ok(!htmlH.includes("height:0.9in"), "certHtml sig1 no longer default size when scaled");

// Template-matched background (PDF/Print cert matches the PPTX design).
const htmlBg = M.certHtml({ name: "X", family: "Excel", title: "T", sy: "2025\u20132026", given: "Given this 1st day of January, 2026.", partner: "P", badge: "B", rakso: "R", wave: "W" }, sigH);
ok(htmlBg.includes('class="mos-comp-bg"'), "certHtml background layer present");
ok(htmlBg.includes('class="mos-comp-bar"') && htmlBg.includes('class="mos-comp-stripe"'), "certHtml green bar + stripe present");
ok(htmlBg.includes('class="mos-comp-wave"') && htmlBg.includes("background-image:url(W)"), "certHtml wave art present");
ok(htmlBg.includes('class="mos-comp-rakso"') && htmlBg.includes('src="R"'), "certHtml RAKSO logo present");
ok(htmlBg.indexOf("mos-comp-bg") < htmlBg.indexOf("mos-comp-body"), "certHtml background paints before body");
ok(htmlBg.includes("<u>1st</u> day of"), "certHtml given day underlined (template style)");
const htmlBg0 = M.certHtml({ name: "X", family: "Excel", title: "T", sy: "", given: "Given this 1st day of January, 2026.", partner: "", badge: "" }, sigH);
ok(htmlBg0.includes('class="mos-comp-bar"') && !htmlBg0.includes("mos-comp-wave") && !htmlBg0.includes("mos-comp-rakso"), "certHtml bars render without assets; wave/rakso omitted when absent");

console.log(pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
