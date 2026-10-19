/* Trainers tab logic tests (pure: js/trainers-parse.js). Synthetic fixtures only —
 * no real workbooks. Run: node tests/trainers-parse.test.cjs */
"use strict";
const P = require("../js/trainers-parse.js");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; } else { fail++; console.log("FAIL - " + n); } };
const eq = (a, b, n) => ok(JSON.stringify(a) === JSON.stringify(b), n + " (got " + JSON.stringify(a) + ")");

/* ---- excelDate ---- */
eq(P.excelDate(44721).iso, "2022-06-09", "excel serial 44721 -> 2022-06-09");
eq(P.excelDate("2026-05-07").iso, "2026-05-07", "iso string passthrough");
eq(P.excelDate("").iso, "", "blank -> empty iso");
eq(P.excelDate(null).iso, "", "null -> empty iso");

/* ---- parseTrackerSheet ---- */
const thead = ["Date", "Year", "Month", "Trainer", "New Code", "New Description", "Client",
               "Time Start", "Time End", "Net Promoter Score"];
const trows = [thead,
  [45000, 2023, "Mar", "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.333, 0.5, ""],
  [45000, 2023, "Mar", "ROE, Sample III", "MS-102A", "Teams Workshop", "HCCCI", 0.54, 0.62, 88.6],
  ["", "", "", "", "", "", "", "", "", ""],
  [45002, 2023, "Mar", "SANTOS, Maria", "MS-103A", "Workshop P1", "SPC", 0.54, 0.62, ""]
];
const ss = P.parseTrackerSheet(trows);
eq(ss.length, 3, "blank rows skipped");
eq(ss[0].trainer, "ROE, Sample III", "trainer parsed");
eq(ss[0].code, "MS-102A", "new code -> code");
eq(ss[0].client, "HCCCI", "client parsed");
eq(ss[0].rowIdx, 1, "rowIdx tracks sheet row");
eq(ss[0].status, "unscored", "empty NPS -> unscored");
eq(ss[1].status, "scored", "filled NPS -> scored");
eq(ss[1].nps, 88.6, "nps numeric");
eq(ss[0].timeStart < 0.5, true, "time start fraction kept");

/* ---- parseFilename ---- */
let f = P.parseFilename("Training_Feedback_Form_-_05072026_-_AM_-_HCCCI_1-44.xlsx");
eq(f, { dateISO: "2026-05-07", session: "AM", school: "HCCCI" }, "sample filename parsed");
eq(P.parseFilename("random.xlsx"), null, "non-matching filename -> null");
eq(P.parseFilename("Form_32132026_-_PM_-_SPC.xlsx"), null, "invalid date -> null");
f = P.parseFilename("Eval_-_12122025_-_PM_-_College of San Westbrook.xlsx");
eq(f && f.school, "College of San Westbrook", "multi-word school");

/* ---- suggestSessions ---- */
const sug = P.suggestSessions(ss, { dateISO: ss[0].dateISO, session: "AM", school: "HCCCI" });
eq(sug.length >= 2, true, "date+school matches found");
eq(sug[0].client, "HCCCI", "school match ranked first");

/* ---- detectLayout ---- */
const hrow = ["ID", "School / Organization", "Knowledge and mastery of the topic",
  "Communication and presentation skills", "Ability to engage and connect with participants",
  "Overall, how satisfied are you with the training experience?",
  "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague",
  "What did you like best about the session?"];
const det = P.detectLayout(hrow);
eq(det.roles.nps, 6, "nps column detected");
eq(det.roles.school, 1, "school column detected");
eq(det.roles.t1, 2, "t1 detected"); eq(det.roles.t2, 3, "t2 detected"); eq(det.roles.t3, 4, "t3 detected");
eq(det.roles.c1, 7, "comment col detected");

/* ---- mapLikert (incl. the disagree/dissatisfied substring trap) ---- */
eq(P.mapLikert("Strongly Agree"), 5, "strongly agree=5");
eq(P.mapLikert("Somewhat Agree"), 4, "somewhat agree=4");
eq(P.mapLikert("Disagree"), 2, "disagree=2 (not 4 via /agree/ substring)");
eq(P.mapLikert("Strongly Disagree"), 1, "strongly disagree=1");
eq(P.mapLikert("Dissatisfied"), 2, "dissatisfied=2 (not 4 via /satisfied/ substring)");
eq(P.mapLikert("Strongly Dissatisfied"), 1, "strongly dissatisfied=1");
eq(P.mapLikert("Excellent"), 5, "excellent=5");
eq(P.mapLikert("Very Good"), 4, "very good=4");
eq(P.mapLikert("Neutral"), 3, "neutral=3");
eq(P.mapLikert("3"), 3, "numeric passthrough");
eq(P.mapLikert("???"), null, "unknown -> null");
eq(P.mapLikert(""), null, "blank -> null");

/* ---- calcNPS ---- */
let nps = P.calcNPS([10, 10, 9, 9, 8, 7, 6, 0, "x", "", null]);
eq(nps.n, 8, "non-numeric ignored");
eq(nps.promoters, 4, "promoters 9-10");
eq(nps.passives, 2, "passives 7-8");
eq(nps.detractors, 2, "detractors 0-6");
eq(nps.nps, 25, "(4-2)/8*100 = 25");
eq(P.calcNPS([]).nps, null, "empty -> null nps");

/* ---- calcTrainerScore ---- */
const roles = { t1: 0, t2: 1, t3: 2 };
const drows = [
  ["Strongly Agree", "Agree", "Strongly Agree"],
  ["Agree", "Neutral", "Disagree"],
  ["???", "", ""]
];
const ts = P.calcTrainerScore(drows, roles);
eq(ts.n, 2, "rows with no mappable items skipped");
eq(ts.score, 3.83, "mean of row means: (4.667 + 3)/2 = 3.83");
eq(ts.itemAvgs.t1, 4.5, "t1 avg");

/* ---- findPrevSession ---- */
const sA = { trainer: "Roe, Sample III", code: "MS-102A", dateISO: "2026-05-07", nps: 88.6, trainerScore: 4.8 };
const sB = { trainer: "ROE, Sample III", code: "ms-102a", dateISO: "2026-06-07", nps: null, trainerScore: null };
const sC = { trainer: "Santos, Maria", code: "MS-103A", dateISO: "2026-05-01", nps: 70, trainerScore: 4 };
eq(P.findPrevSession([sA, sB, sC], sB), sA, "prev = same trainer+code (case-insensitive), earlier date");
eq(P.findPrevSession([sB, sC], sB), null, "no earlier scored session -> null");

/* ---- buildRoster ---- */
eq(P.buildRoster([sA, sB, sC]).length, 2, "roster dedups case-insensitively");

/* ---- planWriteback ---- */
const wbHead = ["Date", "Trainer", "New Code", "Net Promoter Score"];
const wss = [{ rowIdx: 5, computed: { nps: 88.6, trainerScore: 4.75 } }, { rowIdx: 9, computed: null }];
const plan = P.planWriteback(wbHead, wss);
eq(plan.edits.length, 1, "only computed sessions produce edits");
eq(plan.npsCol, 3, "nps col found");
eq(plan.addTrainerCol, true, "trainer score col added when missing");
const plan2 = P.planWriteback(["Date", "Net Promoter Score", "Trainer Score"], wss);
eq(plan2.addTrainerCol, false, "no new column when present");

console.log("\n" + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
