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

/* ---- session id column ---- */
const irows = [["Session ID","Date","Trainer"],
  ["SES-2023-0001", 45000, "ROE, Sample III"],
  ["", 45001, "SANTOS, Maria"]];
const iss = P.parseTrackerSheet(irows);
eq(iss[0].id, "SES-2023-0001", "session id parsed");
eq(iss[1].id, "", "blank session id -> empty string");

/* ---- acronym index + next session id ---- */
eq(P.acronymOf("Hill Crest College Carmel, Inc."), "HCCC", "acronym skips Inc.");
eq(P.acronymOf("Academia De San Valle Fuerte"), "ASVF", "acronym skips De");
var aidx = P.buildAcronymIndex(["Hill Crest College Carmel, Inc.", "Saint Garnet Academy", ""]);
eq(aidx["HCCC"], "Hill Crest College Carmel, Inc.", "acronym lookup");
eq(aidx["SGA"], "Saint Garnet Academy", "acronym lookup 2");
eq(P.nextSessionId([{id:"SES-2024-0001"},{id:"SES-2024-0003"},{id:"SES-2023-0009"},{id:""}], "2024"), "SES-2024-0004", "next id = max+1 for the year");
eq(P.nextSessionId([{id:"SES-2024-0001"}], "2026"), "SES-2026-0001", "new year restarts at 0001");

/* ---- buildSummary ---- */
var sumSessions = [
  { dateISO: "2026-05-07", trainer: "A", code: "MS-101", client: "HCCCI", timeStart: 0.333, timeEnd: 0.5, nps: 80, trainerScore: 4.5, n: 44 },
  { dateISO: "2026-05-08", trainer: "A", code: "MS-101", client: "HCCCI", timeStart: 0.333, timeEnd: 0.5, nps: 60, trainerScore: 4.0, n: 30 },
  { dateISO: "2026-06-01", trainer: "B", code: "WD-201", client: "SPC", timeStart: null, timeEnd: null, nps: null, trainerScore: null }
];
var sm = P.buildSummary(sumSessions);
eq(sm.total, 3, "summary total sessions");
eq(sm.hours, 8, "summary training hours (2 x 4h)");
eq(sm.avgNps, 70, "summary avg NPS over scored only");
eq(sm.scoredNps, 2, "summary scored count");
eq(sm.avgScore, 4.25, "summary avg trainer score");
eq(sm.respondents, 74, "summary respondents summed");
eq(sm.respSessions, 2, "summary respondent session count");
eq(sm.repeatRate, 67, "summary repeat-school rate (2 of 3 at HCCCI)");
eq(sm.bySchool.length, 2, "summary two schools");
eq(sm.bySchool[0].key, "HCCCI", "summary schools sorted");
eq(sm.bySchool[0].sessions, 2, "summary HCCCI session count");
eq(sm.monthly.length, 2, "summary two months");
eq(sm.monthly[0].key, "2026-05", "summary months sorted");
eq(sm.byCourse[0].key, "MS-101", "summary courses sorted");

/* ---- buildCourseMap ---- */
var cmap = P.buildCourseMap([["Course Code","Product","Session Title","Estimated Duration"],
  ["MS-101A","Microsoft","M365 Onboarding for Faculty",3],
  ["MS-101B","Microsoft","M365 Onboarding for Students",3],
  ["","Microsoft","No code row",2]]);
eq(cmap.byCode["MS-101A"].title, "M365 Onboarding for Faculty", "course map code->title");
eq(cmap.byCode["MS-101A"].product, "Microsoft", "course map product");
eq(cmap.byCode["MS-101A"].duration, 3, "course map duration");
eq(cmap.byTitle["M365 Onboarding for Students"], "MS-101B", "course map title->code");
eq(!("" in cmap.byCode), true, "course map skips blank codes");

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

/* ---- detectResponseDate ---- */
var tsHead = ["ID", "Start time", "Completion time", "Q1"];
var tsRows = [tsHead,
  [1, "5/7/2026 9:00:00 AM", "5/7/2026 9:12:00 AM", 10],
  [2, "5/7/2026 9:05:00 AM", "5/7/2026 9:18:00 AM", 9],
  [3, "5/8/2026 9:00:00 AM", "5/8/2026 9:15:00 AM", 10]];
eq(P.detectResponseDate(tsRows), "2026-05-07", "modal timestamp date wins");
var tsSerial = [["ID", "Start time", "Q1"], [1, 46149.375, 10], [2, 46149.5, 9]];
eq(P.detectResponseDate(tsSerial), "2026-05-07", "excel serial timestamps parsed");
eq(P.detectResponseDate([["ID", "Q1"], [1, 10]]), null, "no timestamp column -> null");
eq(P.detectResponseDate([["Start time"]]), null, "header only -> null");
var tsBlank = [["ID", "Start time", "Q1"], [1, "", 10], [2, "5/7/2026 9:00 AM", 9]];
eq(P.detectResponseDate(tsBlank), "2026-05-07", "blank cells ignored");

/* ---- regressions (v1.38.0 review) ---- */
// Serial dates must not shift with the viewer's time zone (UTC getters)
eq(P.excelDate(46149).iso, "2026-05-07", "serial 46149 -> 2026-05-07 in any TZ");
// Text dates: M/D/Y (en-US, like the MMDDYYYY filenames); D/M/Y only when day > 12
eq(P.excelDate("05/07/2026").iso, "2026-05-07", "text 05/07/2026 is May 7 (M/D/Y)");
eq(P.excelDate("25/07/2026").iso, "2026-07-25", "text 25/07/2026 -> day-first when first part > 12");
eq(P.excelDate("May 07, 2026").iso, "2026-05-07", "\"May 07, 2026\" (form Event Date) parsed");
eq(P.excelDate("13/13/2026").iso, "", "impossible date -> empty");
// Text times feed the AM/PM tiebreak
eq(P.timeOf("8:00 AM"), 8 / 24, "\"8:00 AM\" -> 0.333");
eq(P.timeOf("1:30 pm"), 13.5 / 24, "\"1:30 pm\" -> afternoon");
eq(P.timeOf("12:15 AM") < 0.5 && P.timeOf("12:15 PM") >= 0.5, true, "12 AM / 12 PM edge");
eq(P.timeOf(""), null, "blank time -> null");
const tt = P.parseTrackerSheet([["Date", "Trainer", "Client", "Time Start"],
  ["05/07/2026", "A", "HCCCI", "1:00 PM"], ["05/07/2026", "B", "HCCCI", "8:00 AM"]]);
eq(P.suggestSessions(tt, { dateISO: "2026-05-07", session: "AM", school: "HCCCI" })[0].trainer, "B",
  "AM/PM tiebreak works with text times");
// Underscored school names (MS Forms filenames)
eq(P.parseFilename("Training_Feedback_Form_-_05072026_-_PM_-_St._Example_Parish_School_1-44.xlsx").school,
  "St. Example Parish School", "underscored school name -> spaces, _1-44 suffix dropped");
eq(P.parseFilename("Form_-_05072026_-_am_-_A&B_College.xlsx").school, "A&B College", "lowercase am + & + trailing name");
// Same-day fallback ranked below school matches
const fb = P.suggestSessions([{ dateISO: "2026-05-07", client: "Hill Crest College", timeStart: 0.3 },
  { dateISO: "2026-05-07", client: "HCCCI", timeStart: 0.6 }, { dateISO: "2026-05-08", client: "X" }],
  { dateISO: "2026-05-07", session: "AM", school: "HCCCI" });
eq(fb.map(s => s.client), ["HCCCI", "Hill Crest College"], "school match first, same-day fallback next, other days excluded");
// Likert neutral
eq(P.mapLikert("Neither agree nor disagree"), 3, "\"Neither agree nor disagree\" = 3 (not 2)");
eq(P.mapLikert("Neither satisfied nor dissatisfied"), 3, "\"Neither satisfied nor dissatisfied\" = 3");
// Missing NPS column is added, not silently skipped
const p3 = P.planWriteback(["Date", "Trainer"], [{ rowIdx: 1, computed: { nps: 50, trainerScore: 4.2 } }]);
eq([p3.npsCol, p3.addNpsCol, p3.addTrainerCol], [-1, true, true], "no NPS column -> addNpsCol");
const p4 = P.planWriteback(["Net Promoter Score"], [{ rowIdx: 1, computed: { nps: null, trainerScore: null } }]);
eq([p4.edits.length, p4.addTrainerCol], [0, false], "empty result -> no edits, no new column");


/* ---- v1.39 review: session feedback report (analyzeEval) + summary ---- */
const H = ["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic", "Communication and presentation skills",
  "Ability to engage and connect", "The pace of the session was just right", "Overall, how satisfied were you with the training?",
  "On a scale of 0 to 10, how likely are you to recommend", "Why did you give this rating?", "What did you like best", "What suggestions do you have to improve"];
const ER = [H,
  [1, 46149.4, "HCCCI", "Strongly Agree", "Strongly Agree", "Agree", "Neutral", "Very satisfied", 10, "great", "Hands-on activities", "More hands-on practice"],
  [2, 46149.4, "Hill Crest College of Carmel, Inc.", "Agree", "Strongly Agree", "Agree", "Somewhat Disagree", "Satisfied", 6, "Too fast pace", "Examples", "Slower pace, more practice"],
  [3, 46149.4, "HCCCI", "Strongly Agree", "Agree", "Strongly Agree", "Agree", "Dissatisfied", "", "", "n/a", "none"],
  ["", "", "", "", "", "", "", "", "", "", "", ""]];
const ER_ROLES = P.detectLayout(H).roles;
const A = P.analyzeEval(ER, ER_ROLES);
eq(A.responses, 3, "analyzeEval: fully blank rows are not responses");
eq([A.nps.n, A.nps.promoters, A.nps.detractors, A.nps.nps], [2, 1, 1, 0], "analyzeEval: blank NPS skipped (not a detractor) — v7 generator bug");
eq(A.ratings.map(r => r.header.slice(0, 12)), ["Knowledge an", "Communicatio", "Ability to e", "The pace of ", "Overall, how"], "rating items detected; NPS / comment columns are not rating items");
eq(A.ratings[3].avg, 3, "pace item: Neutral 3, Somewhat Disagree 2, Agree 4 -> 3.00");
eq(A.satisfaction, 66.7, "satisfaction = top-2 share of the satisfaction item (2 of 3)");
eq(A.schools.map(s => [s.name, s.count]), [["Hill Crest College of Carmel, Inc.", 3]], "acronym HCCCI merged into the full school name");
eq(A.comments.suggest, ["More hands-on practice", "Slower pace, more practice"], "trivial answers (none / n/a) left out of comments");
eq(A.detractorReasons, ["Too fast pace"], "detractor reasons come from NPS 0–6 rows only");
eq(A.keywords.suggest[0], { word: "practice", count: 2 }, "top suggestion theme counted once per answer");
eq(A.recommendations.map(r => r.priority), A.recommendations.map(r => r.priority).slice().sort((a, b) => ({ high: 0, medium: 1, low: 2 })[a] - ({ high: 0, medium: 1, low: 2 })[b]), "recommendations ordered high → low");
ok(A.recommendations.some(r => /pace/i.test(r.title)), "lowest-rated item (pace) gets a recommendation");
eq(A.responseDate, "2026-05-07", "answer date from the Start time column");
eq(P.ratingColumns([["Comment"], ["Excellent session!"], ["Good"], ["Too long"], ["Fair enough I guess"], ["ok"], ["Loved it"], ["Hmm"], ["A"], ["B"]], {}).length, 0,
  "free-text column with many distinct answers is not a rating item");
const AE = P.analyzeEval([H], ER_ROLES);
eq([AE.responses, AE.nps.nps, AE.overall, AE.recommendations.length], [0, null, null, 0], "empty eval: no crash, nulls, no recommendations");
ok(P.isTrivial("N/A") && P.isTrivial(" none. ") && P.isTrivial("wala") && !P.isTrivial("Nothing to improve, great job"), "isTrivial");
const SUM = P.buildSummary([{ client: "A", nps: 50, trainerScore: 4.5, n: 20, timeStart: 0.25, timeEnd: 0.5 }, { client: "A", nps: null, trainerScore: null, n: null }, { client: "B", nps: -10, trainerScore: 3.5, n: 5 }]);
eq([SUM.respondents, SUM.respSessions, SUM.avgNps, SUM.avgScore, SUM.hours, SUM.repeatRate], [25, 2, 20, 4, 6, 67], "buildSummary: respondents from n, negative NPS averaged, hours, repeat rate");

/* ---- bulk matching (matchEval / rankSessions) ---- */
const MS = [{ dateISO: "2026-05-07", client: "HCCCI", timeStart: 0.33 }, { dateISO: "2026-05-07", client: "HCCCI", timeStart: 0.6 },
            { dateISO: "2026-05-08", client: "St. Pine College", timeStart: 0.33 }, { dateISO: "2026-05-08", client: "Pico Catholic School", timeStart: 0.33 }];
ok(P.matchEval(MS, { dateISO: "2026-05-07", session: "PM", school: "HCCCI" }).auto === MS[1], "matchEval: PM file -> PM session");
ok(P.matchEval(MS, { dateISO: "2026-05-08", session: "AM", school: "SPC" }).auto === MS[2], "matchEval: acronym SPC -> St. Pine College");
ok(P.matchEval(MS, { dateISO: "2026-05-08", session: "AM", school: "Unknown" }).auto === null, "matchEval: no school match -> nothing auto-picked");
eq(P.matchEval(MS, { dateISO: "2026-05-08", session: "AM", school: "Unknown" }).candidates.length, 2, "matchEval: same-day sessions still offered");
ok(P.matchEval(MS, { dateISO: "2026-05-07", session: "", school: "HCCCI" }).auto === null, "matchEval: two equal matches (no AM/PM) -> user picks");
ok(P.matchEval(MS, { dateISO: "2026-01-01", session: "AM", school: "SPC" }, "2026-05-08").auto === MS[2], "matchEval: falls back to the answer date when the filename date has no session");
ok(P.matchEval(MS, null, "2026-05-07").auto === null && P.matchEval(MS, null, "2026-05-07").candidates.length === 2, "matchEval: no filename hint -> candidates from answer date only");

/* ---- trainer / course reporting ---- */
const TS = [
  { trainer: "ROE, R", code: "C1", title: "Teams", client: "X", dateISO: "2026-05-01", timeStart: 0.25, timeEnd: 0.5, nps: 50, trainerScore: 4.5, n: 10, itemAvgs: { t1: 4.8, t2: 4.2, t3: 4.5 } },
  { trainer: "roe, r", code: "C1", title: "", client: "Y", dateISO: "2026-04-01", timeStart: 0.25, timeEnd: 0.5, nps: null, trainerScore: null, n: null },
  { trainer: "ROE, R", code: "C2", title: "Excel", client: "X", dateISO: "2026-04-15", nps: -20, trainerScore: 3.5, n: 4 },
  { trainer: "SANTOS, M", code: "C2", title: "Excel", client: "Z", dateISO: "2026-02-01", timeStart: 0.5, timeEnd: 0.625, nps: 10, trainerScore: 4, n: 5 }];
const TT = P.trainerTable(TS, [{ trainer: "Roe, R", nps: 0 }, { trainer: "ROE, R", nps: 10 }]);
eq(TT.map(r => [r.trainer, r.sessions, r.hours, r.courses, r.schools, r.respondents, r.avgNps, r.avgScore, r.noEval, r.last]),
  [["ROE, R", 3, 12, 2, 2, 14, 15, 4, 1, "2026-05-01"], ["SANTOS, M", 1, 3, 1, 1, 5, 10, 4, 0, "2026-02-01"]],
  "trainerTable: trainer names grouped case-insensitively; trainings, hours, courses, schools, respondents, averages, no-eval, last");
eq([TT[0].topCourse.title, TT[0].topCourse.sessions, TT[0].npsChange, TT[1].npsChange], ["Teams", 2, 10, null], "trainerTable: top course + NPS change vs previous period (15 vs avg 5)");
const CT = P.courseTable(TS);
eq(CT.map(c => [c.code, c.title, c.sessions, c.hours, c.schools, c.trainers]), [["C1", "Teams", 2, 12, 2, 1], ["C2", "Excel", 2, 3, 2, 2]], "courseTable: times conducted per course, title from any session");
eq(P.overview(TS), { trainers: 2, sessions: 4, hours: 15, courses: 2, schools: 3, evaluated: 3, coverage: 75 }, "overview: tiles (trainers, trainings, hours, courses, schools, eval coverage)");
eq(P.itemAverages(TS), { n: 1, t1: 4.8, t2: 4.2, t3: 4.5 }, "itemAverages from sessions with eval reports only");
eq(P.periodsOf(TS).map(p => p.key), ["2026", "2026-Q2", "2026-Q1"], "periodsOf: years then quarters, newest first");
ok(P.inPeriod("2026-04-15", "2026-Q2") && !P.inPeriod("2026-03-31", "2026-Q2") && P.inPeriod("2026-03-31", "2026") && P.inPeriod("x", ""), "inPeriod boundaries");
eq([P.prevPeriod("2026"), P.prevPeriod("2026-Q1"), P.prevPeriod("2026-Q3"), P.prevPeriod("")], ["2025", "2025-Q4", "2026-Q2", ""], "prevPeriod");
eq(P.trainerTable([], []), [], "trainerTable: empty");

console.log("\n" + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
