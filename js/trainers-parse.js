/* =============================================================================
 * trainers-parse.js — pure logic for the Trainers tab (trainer evaluations).
 * No DOM, no QBR dependency: exercised by tests/trainers-parse.test.cjs in
 * Node (module.exports) and loaded in the browser as QBR.trainersParse.
 *
 * Sessions come from the training tracker workbook ("trial" sheet style):
 *   Date | Trainer | New Code | New Description | Client | Time Start/End |
 *   Net Promoter Score | (Trainer Score added by write-back)
 * Eval files are MS Forms exports: one file = one session. The form carries
 * no trainer column, so the trainer is inherited from the tracker session
 * the user confirms (filename date/school/AM-PM only *suggests* candidates).
 * ========================================================================== */
(function (root) {
  "use strict";

  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }
  function normKey(s) { return norm(s).toLowerCase(); }

  /* Excel serial (or a date-ish string) -> { date: Date|null, iso: "YYYY-MM-DD"|"" } */
  function excelDate(v) {
    if (v == null || v === "") return { date: null, iso: "" };
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
      return { date: v, iso: isoOf(v) };
    }
    if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) {
      var d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return { date: d, iso: isoOf(d) };
    }
    var s = norm(v);
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/) || s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      var y, mo, da;
      if (m[1].length === 4) { y = +m[1]; mo = +m[2]; da = +m[3]; }
      else { da = +m[1]; mo = +m[2]; y = +m[3]; }
      var dd = new Date(y, mo - 1, da);
      if (!isNaN(dd)) return { date: dd, iso: isoOf(dd) };
    }
    return { date: null, iso: "" };
  }
  function isoOf(d) {
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
  }

  /* ---- tracker sheet (array-of-arrays, header first) -> sessions ---------- */
  var TRACKER_COLS = {
    date: ["date"],
    trainer: ["trainer", "facilitator"],
    code: ["new code", "course code", "code"],
    title: ["new description", "session title", "title", "description"],
    client: ["client", "school", "organization"],
    timeStart: ["time start", "start"],
    timeEnd: ["time end", "end"],
    nps: ["net promoter score", "nps"],
    trainerScore: ["trainer score"]
  };
  function findCol(header, names) {
    for (var i = 0; i < header.length; i++) {
      var h = normKey(header[i]);
      if (!h) continue;
      for (var k = 0; k < names.length; k++) if (h === names[k]) return i;
    }
    return -1;
  }
  function numOrNull(v) {
    if (v == null || v === "") return null;
    var n = Number(String(v).replace(",", "."));
    return isFinite(n) ? n : null;
  }
  function parseTrackerSheet(rows) {
    var out = [];
    if (!rows || rows.length < 2) return out;
    var header = rows[0].map(norm), col = {};
    Object.keys(TRACKER_COLS).forEach(function (k) { col[k] = findCol(header, TRACKER_COLS[k]); });
    if (col.date < 0 || col.trainer < 0) return out; // not a session log
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row) continue;
      var trainer = norm(row[col.trainer]);
      var code = col.code >= 0 ? norm(row[col.code]) : "";
      if (!trainer && !code) continue; // blank row
      var dt = excelDate(row[col.date]);
      var nps = col.nps >= 0 ? numOrNull(row[col.nps]) : null;
      var tsc = col.trainerScore >= 0 ? numOrNull(row[col.trainerScore]) : null;
      out.push({
        rowIdx: r, // 0-based sheet row (header = 0); write-back uses this
        date: dt.date, dateISO: dt.iso,
        trainer: trainer, code: code,
        title: col.title >= 0 ? norm(row[col.title]) : "",
        client: col.client >= 0 ? norm(row[col.client]) : "",
        timeStart: col.timeStart >= 0 ? numOrNull(row[col.timeStart]) : null,
        timeEnd: col.timeEnd >= 0 ? numOrNull(row[col.timeEnd]) : null,
        nps: nps, trainerScore: tsc,
        status: (nps != null ? "scored" : "unscored")
      });
    }
    return out;
  }

  /* ---- eval filename -> { dateISO, session: "AM"|"PM", school } | null -----
   * e.g. "Training_Feedback_Form_-_05072026_-_AM_-_HCCCI_1-44.xlsx"
   * Date is MMDDYYYY (matches the form's "Event Date: May 07, 2026").
   * The trainer is NEVER in the filename. */
  function parseFilename(name) {
    var base = norm(name).replace(/\.[a-z0-9]+$/i, "");
    var m = base.match(/(\d{2})(\d{2})(\d{4})[\s\-_]+(AM|PM)[\s\-_]+([A-Za-z0-9][A-Za-z0-9 .&'\-]*?)(?=[\s\-_]+(?:\d|$)|$)/i);
    if (!m) return null;
    var da = +m[2], mo = +m[1], ye = +m[3];
    if (mo < 1 || mo > 12 || da < 1 || da > 31) return null;
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return { dateISO: ye + "-" + p(mo) + "-" + p(da), session: m[4].toUpperCase(), school: norm(m[5]) };
  }

  /* ---- filename hint -> ranked tracker sessions ---------------------------- */
  function suggestSessions(sessions, hint) {
    if (!hint) return [];
    var hs = normKey(hint.school);
    var scored = [];
    (sessions || []).forEach(function (s) {
      if (s.dateISO !== hint.dateISO) return;
      var c = normKey(s.client);
      var schoolScore = 0;
      if (hs && c) {
        if (c === hs) schoolScore = 2;
        else if (c.indexOf(hs) >= 0 || hs.indexOf(c) >= 0) schoolScore = 1;
      }
      if (!schoolScore) return;
      var sessScore = 0;
      if (hint.session && s.timeStart != null) {
        var isAM = s.timeStart < 0.5;
        if ((hint.session === "AM") === isAM) sessScore = 1;
      }
      scored.push({ s: s, rank: schoolScore * 10 + sessScore });
    });
    scored.sort(function (a, b) { return b.rank - a.rank; });
    return scored.map(function (x) { return x.s; });
  }

  /* ---- MS Forms export layout detection ------------------------------------
   * Returns { sig, roles } where roles maps a role name -> 0-based column
   * index (or -1 when not found). The UI confirms/adjusts on first sight and
   * remembers the mapping per sig in localStorage. */
  var LAYOUT_ROLES = [
    ["nps", /how likely.*recommend|recommend.*colleague/i],
    ["school", /school.*organization|organization/i],
    ["t1", /knowledge.*mastery/i],
    ["t2", /communication.*presentation/i],
    ["t3", /engage.*connect/i],
    ["sat", /overall.*satisf|satisf.*training experience/i],
    ["c1", /lik(e|ed) best/i],
    ["c2", /key insight|skill.*gain/i],
    ["c3", /suggestion.*improve/i],
    ["cnps", /why.*rat/i]
  ];
  function detectLayout(headerRow) {
    var header = (headerRow || []).map(norm);
    var sig = header.map(normKey).join("|");
    var roles = {};
    LAYOUT_ROLES.forEach(function (pair) {
      var idx = -1;
      for (var i = 0; i < header.length; i++) {
        if (pair[1].test(header[i])) { idx = i; break; }
      }
      roles[pair[0]] = idx;
    });
    return { sig: sig, roles: roles };
  }

  /* ---- text Likert -> 1..5 -------------------------------------------------- */
  var LIKERT = [
    [/strongly disagree/, 1], [/somewhat disagree/, 2], [/disagree/, 2],
    [/strongly agree/, 5], [/somewhat agree/, 4], [/^agree$/, 4], [/agree/, 4],
    [/strongly dissatisfied/, 1], [/very dissatisfied/, 1], [/dissatisfied/, 2],
    [/strongly satisfied/, 5], [/very satisfied/, 5], [/^satisfied$/, 4], [/satisfied/, 4],
    [/excellent/, 5], [/very good/, 4], [/^good$/, 3], [/fair/, 2], [/poor/, 1],
    [/neutral/, 3], [/neither/, 3]
  ];
  function mapLikert(text) {
    var t = normKey(text);
    if (!t) return null;
    if (/^[1-5]$/.test(t)) return +t;
    for (var i = 0; i < LIKERT.length; i++) if (LIKERT[i][0].test(t)) return LIKERT[i][1];
    return null;
  }

  /* ---- NPS: standard formula ------------------------------------------------ */
  function calcNPS(scores) {
    var n = 0, pro = 0, pas = 0, det = 0;
    (scores || []).forEach(function (v) {
      if (v == null) return;
      var t = String(v).trim();
      if (!t) return; // blank answers are skipped, NOT scored as 0
      var x = Number(t);
      if (!isFinite(x) || x < 0 || x > 10) return;
      n++;
      if (x >= 9) pro++; else if (x >= 7) pas++; else det++;
    });
    return { n: n, promoters: pro, passives: pas, detractors: det,
             nps: n ? Math.round(((pro - det) / n) * 1000) / 10 : null };
  }

  /* ---- trainer effectiveness: mean of the 3 facilitator items --------------- */
  function calcTrainerScore(dataRows, roles) {
    var keys = ["t1", "t2", "t3"], sums = { t1: 0, t2: 0, t3: 0 }, ns = { t1: 0, t2: 0, t3: 0 };
    var rowScores = [];
    (dataRows || []).forEach(function (r) {
      var vals = [];
      keys.forEach(function (k) {
        var ci = roles[k];
        if (ci == null || ci < 0) return;
        var v = mapLikert(r[ci]);
        if (v != null) { sums[k] += v; ns[k]++; vals.push(v); }
      });
      if (vals.length) rowScores.push(vals.reduce(function (a, b) { return a + b; }, 0) / vals.length);
    });
    var itemAvgs = {};
    keys.forEach(function (k) { itemAvgs[k] = ns[k] ? Math.round((sums[k] / ns[k]) * 100) / 100 : null; });
    var score = rowScores.length
      ? Math.round((rowScores.reduce(function (a, b) { return a + b; }, 0) / rowScores.length) * 100) / 100
      : null;
    return { score: score, n: rowScores.length, itemAvgs: itemAvgs };
  }

  /* ---- previous session: same trainer + same code, most recent before cur --- */
  function findPrevSession(sessions, cur) {
    var ct = normKey(cur.trainer), cc = normKey(cur.code);
    if (!ct) return null;
    var best = null;
    (sessions || []).forEach(function (s) {
      if (s === cur) return;
      if (normKey(s.trainer) !== ct || normKey(s.code) !== cc) return;
      if (!s.dateISO || (cur.dateISO && s.dateISO >= cur.dateISO)) return;
      if (s.nps == null && s.trainerScore == null) return;
      if (!best || s.dateISO > best.dateISO) best = s;
    });
    return best;
  }

  function buildRoster(sessions) {
    var seen = {}, out = [];
    (sessions || []).forEach(function (s) {
      var t = norm(s.trainer);
      if (!t) return;
      var k = normKey(t);
      if (!seen[k]) { seen[k] = 1; out.push(t); }
    });
    out.sort(function (a, b) { return normKey(a) < normKey(b) ? -1 : 1; });
    return out;
  }

  /* ---- write-back plan: edits to apply to the tracker sheet -----------------
   * rows: array-of-arrays (live sheet). Returns { edits, addTrainerCol }.
   * The UI applies edits then downloads the workbook. */
  function planWriteback(headerRow, sessions) {
    var header = (headerRow || []).map(norm);
    var npsCol = findCol(header, TRACKER_COLS.nps);
    var tsCol = findCol(header, TRACKER_COLS.trainerScore);
    var edits = [];
    sessions.forEach(function (s) {
      if (s.computed && (s.computed.nps != null || s.computed.trainerScore != null)) {
        edits.push({ row: s.rowIdx, nps: s.computed.nps, trainerScore: s.computed.trainerScore });
      }
    });
    return { edits: edits, npsCol: npsCol, trainerCol: tsCol, addTrainerCol: tsCol < 0 && edits.length > 0 };
  }

  var api = {
    norm: norm, normKey: normKey, excelDate: excelDate,
    parseTrackerSheet: parseTrackerSheet, parseFilename: parseFilename,
    suggestSessions: suggestSessions, detectLayout: detectLayout,
    mapLikert: mapLikert, calcNPS: calcNPS, calcTrainerScore: calcTrainerScore,
    findPrevSession: findPrevSession, buildRoster: buildRoster,
    planWriteback: planWriteback, TRACKER_COLS: TRACKER_COLS
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else { root.QBR = root.QBR || {}; root.QBR.trainersParse = api; }
})(typeof self !== "undefined" ? self : this);
