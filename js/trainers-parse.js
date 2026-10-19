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

  function norm(s) { return String(s == null ? "" : s).replace(/[\u200b\u200c\ufeff]/g, "").replace(/\s+/g, " ").trim(); }
  function normKey(s) { return norm(s).toLowerCase(); }

  /* Excel serial (or a date-ish string) -> { date: Date|null, iso: "YYYY-MM-DD"|"" } */
  function excelDate(v) {
    if (v == null || v === "") return { date: null, iso: "" };
    if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
      return { date: v, iso: isoOf(v) };
    }
    if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) {
      // Serials are calendar days: read with UTC getters so the ISO date
      // never shifts by the viewer's time zone.
      var u = new Date(Math.round((v - 25569) * 86400 * 1000));
      var d = new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
      return { date: d, iso: isoOf(d) };
    }
    var s = norm(v);
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/) || s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    var mn = !m && s.match(/^([A-Za-z]{3,})\.? (\d{1,2}),? (\d{4})/); // "May 07, 2026"
    if (mn) {
      var mi = MONTHS.indexOf(mn[1].slice(0, 3).toLowerCase());
      if (mi >= 0) m = [s, String(mn[3]), String(mi + 1), mn[2]];
    }
    if (m) {
      var y, mo, da;
      if (m[1].length === 4) { y = +m[1]; mo = +m[2]; da = +m[3]; }
      // Text dates are M/D/Y (en-US, matches the MMDDYYYY eval filenames)
      // unless the first part can only be a day.
      else if (+m[1] > 12) { da = +m[1]; mo = +m[2]; y = +m[3]; }
      else { mo = +m[1]; da = +m[2]; y = +m[3]; }
      if (mo < 1 || mo > 12 || da < 1 || da > 31) return { date: null, iso: "" };
      var dd = new Date(y, mo - 1, da);
      if (!isNaN(dd)) return { date: dd, iso: isoOf(dd) };
    }
    return { date: null, iso: "" };
  }
  var MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  /* Time cell -> day fraction (0..1). Accepts Excel fractions and "8:00 AM" text. */
  function timeOf(v) {
    var n = numOrNull(v);
    if (n != null) return n;
    var m = norm(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])\.?m?\.?$/i) || norm(v).match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) return null;
    var h = +m[1] % (m[3] ? 12 : 24);
    if (m[3] && /p/i.test(m[3])) h += 12;
    return (h * 60 + +m[2]) / 1440;
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
    trainerScore: ["trainer score"],
    id: ["session id", "sessionid"],
    evalSource: ["eval source"],
    responses: ["responses", "no. of responses", "number of responses", "respondents", "response count"]
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
      var src = col.evalSource >= 0 ? norm(row[col.evalSource]) : "";
      var resp = col.responses >= 0 ? numOrNull(row[col.responses]) : null;
      out.push({
        rowIdx: r, // 0-based sheet row (header = 0); write-back uses this
        evalSource: src, evalFp: fpOf(src),
        date: dt.date, dateISO: dt.iso,
        id: col.id >= 0 ? norm(row[col.id]) : "",
        trainer: trainer, code: code,
        title: col.title >= 0 ? norm(row[col.title]) : "",
        client: col.client >= 0 ? norm(row[col.client]) : "",
        timeStart: col.timeStart >= 0 ? timeOf(row[col.timeStart]) : null,
        timeEnd: col.timeEnd >= 0 ? timeOf(row[col.timeEnd]) : null,
        nps: nps, trainerScore: tsc, n: resp,
        status: (nps != null ? "scored" : "unscored")
      });
    }
    return out;
  }

  /* ---- eval filename -> { dateISO, session: "AM"|"PM", school } | null -----
   * e.g. "Training_Feedback_Form_-_05072026_-_AM_-_HCCCI_1-44.xlsx"
   * Date is MMDDYYYY (matches the form's "Event Date: May 07, 2026").
   * The trainer is NEVER in the filename. */
  /* Also seen in real exports:
   *   "Training Feedback Form - 05072026 - AM - HCCCI(1-44).xlsx"
   *   "Training Feedback Form - Nov. 28, 2025 [SGA] - Microsoft 36.xlsx"  (date in words, school in [ ]) */
  var MON3 = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  function parseFilename(name) {
    var base = norm(name).replace(/\.[a-z0-9]+$/i, "");
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    var ok = function (y, mo, da) { return mo >= 1 && mo <= 12 && da >= 1 && da <= 31 && y >= 2000 && y < 2100; };
    var bracket = (base.match(/\[([^\]]+)\]/) || [])[1];
    var m = base.match(/(\d{2})(\d{2})(\d{4})[\s\-_]+(AM|PM)[\s\-_]+([A-Za-z0-9][A-Za-z0-9 _.&'\-]*?)(?=[\s\-_]*[(\[]|[\s\-_]+(?:\d|$)|$)/i);
    if (m && ok(+m[3], +m[1], +m[2]))
      return { dateISO: m[3] + "-" + p(+m[1]) + "-" + p(+m[2]), session: m[4].toUpperCase(), school: norm((bracket || m[5]).replace(/_/g, " ")) };
    var y, mo, da;
    if ((m = base.match(/(?:^|\D)(\d{2})(\d{2})(\d{4})(?!\d)/)) && ok(+m[3], +m[1], +m[2])) { mo = +m[1]; da = +m[2]; y = +m[3]; }
    else if ((m = base.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s_]+(\d{1,2})(?:st|nd|rd|th)?,?[\s_]+(\d{4})\b/i)) &&
             ok(+m[3], MON3.indexOf(m[1].toLowerCase()) + 1, +m[2])) { mo = MON3.indexOf(m[1].toLowerCase()) + 1; da = +m[2]; y = +m[3]; }
    else return null;
    var ses = (base.match(/(?:^|[\s\-_\[(])(AM|PM)(?=$|[\s\-_\])])/i) || [])[1];
    return { dateISO: y + "-" + p(mo) + "-" + p(da), session: ses ? ses.toUpperCase() : "", school: norm((bracket || "").replace(/_/g, " ")) };
  }

  /* ---- filename hint -> ranked tracker sessions ---------------------------- */
  function suggestSessions(sessions, hint) {
    return rankSessions(sessions, hint).map(function (x) { return x.s; });
  }
  /* rank = school (exact or acronym 20, partial 10, none 0) + AM/PM match (1). Same day only. */
  function rankSessions(sessions, hint) {
    if (!hint || !hint.dateISO) return [];
    var hs = normKey(hint.school), hsA = hs.replace(/[^a-z]/g, "");
    var scored = [];
    (sessions || []).forEach(function (s) {
      if (s.dateISO !== hint.dateISO) return;
      var c = normKey(s.client);
      var schoolScore = 0;
      if (hs && c) {
        var acr = acronymOf(s.client).toLowerCase();
        if (c === hs || (hsA.length >= 2 && acr === hsA) ||
            (hsA.length >= 4 && acr.length === hsA.length - 1 && hsA.indexOf(acr) === 0) || // HCCCI = Holy Cross College of Carigara (Inc.)
            (hint.aliases && hint.aliases[hsA] === c)) schoolScore = 2;                    // learned from an earlier pick
        else if (c.indexOf(hs) >= 0 || hs.indexOf(c) >= 0) schoolScore = 1;
      }
      // No school match: still offer same-day sessions, ranked below matches
      // (tracker client names and filename acronyms often differ).
      var sessScore = 0;
      if (hint.session && s.timeStart != null) {
        var isAM = s.timeStart < 0.5;
        if ((hint.session === "AM") === isAM) sessScore = 1;
      }
      scored.push({ s: s, rank: schoolScore * 10 + sessScore });
    });
    scored.sort(function (a, b) { return b.rank - a.rank; });
    return scored;
  }
  /* Bulk matching: candidates for one eval file, and the session to pre-select when it is
   * unambiguous (school matches and no other candidate ranks as high). The filename date is
   * tried first, then the answers' own date. Never auto-picks a school-less same-day match. */
  function matchEval(sessions, hint, responseDate, opts) {
    opts = opts || {};
    var tries = [], base = { session: (hint && hint.session) || opts.session || "", school: hint ? hint.school : "", aliases: opts.aliases || null };
    if (hint && hint.dateISO) tries.push({ dateISO: hint.dateISO, session: base.session, school: base.school, aliases: base.aliases });
    if (responseDate && (!hint || responseDate !== hint.dateISO))
      tries.push({ dateISO: responseDate, session: base.session, school: base.school, aliases: base.aliases });
    for (var i = 0; i < tries.length; i++) {
      var r = rankSessions(sessions, tries[i]);
      if (!r.length) continue;
      var top = r[0], auto = top.rank >= 20 && (r.length === 1 || r[1].rank < top.rank) ? top.s : null;
      return { candidates: r, auto: auto, dateISO: tries[i].dateISO };
    }
    return { candidates: [], auto: null, dateISO: "" };
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
    [/neutral/, 3], [/neither/, 3], // before /disagree/: "Neither agree nor disagree" = 3
    [/strongly disagree/, 1], [/somewhat disagree/, 2], [/disagree/, 2],
    [/strongly agree/, 5], [/somewhat agree/, 4], [/^agree$/, 4], [/agree/, 4],
    [/strongly dissatisfied/, 1], [/very dissatisfied/, 1], [/dissatisfied/, 2],
    [/strongly satisfied/, 5], [/very satisfied/, 5], [/^satisfied$/, 4], [/satisfied/, 4],
    [/excellent/, 5], [/very good/, 4], [/^good$/, 3], [/fair/, 2], [/poor/, 1]
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
    var ct = splitTrainers(cur.trainer).map(normKey), cc = normKey(cur.code);
    if (!ct.length) return null;
    var best = null;
    (sessions || []).forEach(function (s) {
      if (s === cur) return;
      if (normKey(s.code) !== cc || !splitTrainers(s.trainer).some(function (t) { return ct.indexOf(normKey(t)) >= 0; })) return;
      if (!s.dateISO || (cur.dateISO && s.dateISO >= cur.dateISO)) return;
      if (s.nps == null && s.trainerScore == null) return;
      if (!best || s.dateISO > best.dateISO) best = s;
    });
    return best;
  }

  /* ---- co-trainers (v1.39.0) ------------------------------------------------
   * One tracker row per session; several trainers share the Trainer cell:
   * "CEDO, Anthony / FAYLONA, Franchesca" (also ; & + line break " and ").
   * Never split on a comma: names are "LASTNAME, Firstname". */
  var TRAINER_SEP = /\s*(?:\/|;|&|\+|\r?\n|\s+and\s+)\s*/i;
  function splitTrainers(v) {
    var seen = {}, out = [];
    String(v == null ? "" : v).split(TRAINER_SEP).forEach(function (x) {
      var t = norm(x), k = normKey(t);
      if (t && !seen[k]) { seen[k] = 1; out.push(t); }
    });
    return out;
  }
  function hasTrainer(s, name) {
    var k = normKey(name);
    return !!k && splitTrainers(s && s.trainer).some(function (t) { return normKey(t) === k; });
  }
  function joinTrainers(list) { return splitTrainers((list || []).join(" / ")).join(" / "); }
  /* Rows that look like the same session entered twice (same date, school, course, start time),
   * e.g. one row per co-trainer. -> [{ rows: [session…], label }] */
  function duplicateSessions(sessions) {
    var g = {}, order = [];
    (sessions || []).forEach(function (s) {
      if (!s.dateISO || !normKey(s.client)) return;
      var k = [s.dateISO, normKey(s.client), normKey(s.code), s.timeStart == null ? "" : Math.round(s.timeStart * 1440)].join("|");
      if (!g[k]) { g[k] = []; order.push(k); }
      g[k].push(s);
    });
    return order.filter(function (k) { return g[k].length > 1; }).map(function (k) {
      var a = g[k][0];
      return { rows: g[k], label: a.dateISO + " · " + a.client + (a.code ? " · " + a.code : "") };
    });
  }
  function buildRoster(sessions) {
    var seen = {}, out = [];
    (sessions || []).forEach(function (s) {
      splitTrainers(s.trainer).forEach(function (t) {
        var k = normKey(t);
        if (!seen[k]) { seen[k] = 1; out.push(t); }
      });
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
    var srcCol = findCol(header, TRACKER_COLS.evalSource);
    var respCol = findCol(header, TRACKER_COLS.responses);
    var edits = [];
    sessions.forEach(function (s) {
      if (s.computed && (s.computed.nps != null || s.computed.trainerScore != null)) {
        edits.push({ row: s.rowIdx, nps: s.computed.nps, trainerScore: s.computed.trainerScore, evalSource: s.computed.evalSource || "",
                     n: s.computed.n != null ? s.computed.n : null });
      } else if (s.fillN != null && s.n == null) {
        // v1.39.0: response count only (row already encoded from this eval, count never recorded)
        edits.push({ row: s.rowIdx, nps: null, trainerScore: null, evalSource: "", n: s.fillN, countOnly: true });
      }
    });
    return { edits: edits, npsCol: npsCol, trainerCol: tsCol, sourceCol: srcCol, respCol: respCol,
             addRespCol: respCol < 0 && edits.some(function (e) { return e.n != null; }),
             addNpsCol: npsCol < 0 && edits.some(function (e) { return e.nps != null; }),
             addTrainerCol: tsCol < 0 && edits.some(function (e) { return e.trainerScore != null; }),
             addSourceCol: srcCol < 0 && edits.some(function (e) { return e.evalSource; }) };
  }

  /* ---- response-date detection -------------------------------------------
   * The eval's own answer timestamps ("Start time") are ground truth for the
   * session date — more reliable than the filename. Returns the modal
   * YYYY-MM-DD, or null when no timestamp column is found. */
  var TS_COLS = ["start time", "completion time", "submitted", "timestamp", "response time"];
  function detectResponseDate(rows) {
    if (!rows || rows.length < 2) return null;
    var header = rows[0].map(normKey), ci = -1;
    for (var i = 0; i < header.length && ci < 0; i++)
      for (var k = 0; k < TS_COLS.length; k++)
        if (header[i].indexOf(TS_COLS[k]) >= 0) { ci = i; break; }
    if (ci < 0) return null;
    var counts = {};
    for (var r = 1; r < rows.length; r++) {
      var v = rows[r] ? rows[r][ci] : null;
      if (v == null || v === "") continue;
      var iso = null;
      if (typeof v === "number" && isFinite(v)) {
        iso = excelDate(v).iso;
      } else {
        var m = /^(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/.exec(String(v).trim());
        if (m) {
          var a = +m[1], b = +m[2], cc = +m[3];
          var ye = a > 31 ? a : cc, mo = a > 31 ? b : a, da = a > 31 ? cc : b;
          if (ye < 100) ye += 2000;
          iso = ye + "-" + String(mo).padStart(2, "0") + "-" + String(da).padStart(2, "0");
        }
      }
      if (iso) counts[iso] = (counts[iso] || 0) + 1;
    }
    var best = null, bestN = 0;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > bestN) { bestN = counts[k]; best = k; }
    });
    return best;
  }

  /* ---- eval fingerprint: same answers = same file, whatever its name ---------
   * FNV-1a over the response IDs + timestamps (or every cell when those are missing). */
  function fnv(str) {
    var h = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return ("0000000" + h.toString(16)).slice(-8);
  }
  function evalFingerprint(rows) {
    if (!rows || rows.length < 2) return "";
    var header = rows[0].map(normKey), keep = [];
    header.forEach(function (h, i) { if (h === "id" || /start time|completion time|email/.test(h)) keep.push(i); });
    var data = rows.slice(1).filter(function (r) { return r && r.some(function (v) { return norm(v) !== ""; }); });
    var parts = data.map(function (r) {
      return (keep.length ? keep.map(function (i) { return norm(r[i]); }) : r.map(norm)).join("\u0001");
    });
    return fnv(data.length + "\u0002" + parts.join("\u0002"));
  }
  function evalSourceText(name, fp) { return norm(name) + (fp ? " · fp:" + fp : ""); }
  function fpOf(src) { var m = /fp:([0-9a-f]{8})/.exec(src || ""); return m ? m[1] : ""; }
  // AM/PM from the answers' own times (median): forms are answered at the end of the session.
  function detectResponseSession(rows) {
    if (!rows || rows.length < 2) return "";
    var header = rows[0].map(normKey), ci = -1;
    for (var i = 0; i < header.length && ci < 0; i++) if (/start time|completion time|timestamp|submitted/.test(header[i])) ci = i;
    if (ci < 0) return "";
    var fr = [];
    rows.slice(1).forEach(function (r) {
      var v = r && r[ci];
      if (typeof v === "number" && isFinite(v)) { var f = v - Math.floor(v); if (f > 0) fr.push(f); }
      else if (v) { var m = /(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]m)?/i.exec(String(v)); if (m) { var h = +m[1] % (m[3] ? 12 : 24); if (m[3] && /p/i.test(m[3])) h += 12; fr.push((h * 60 + +m[2]) / 1440); } }
    });
    if (!fr.length) return "";
    fr.sort(function (a, b) { return a - b; });
    return fr[Math.floor(fr.length / 2)] < 0.55 ? "AM" : "PM"; // before ~1:12 PM
  }

  var ACR_STOP = /^(of|de|del|la|the|and|for|inc\.?|corp\.?)$/i;
  function acronymOf(name) {
    return norm(name).split(/[^A-Za-z]+/).filter(function (w) { return w && !ACR_STOP.test(w); })
      .map(function (w) { return w[0].toUpperCase(); }).join("");
  }
  function buildAcronymIndex(clients) {
    var idx = {};
    (clients || []).forEach(function (c) {
      c = norm(c);
      if (!c) return;
      var a = acronymOf(c);
      if (a && a.length >= 2 && !idx[a]) idx[a] = c;
    });
    return idx;
  }
  function nextSessionId(sessions, year) {
    var max = 0;
    (sessions || []).forEach(function (s) {
      var m = /^SES-(\d{4})-(\d+)$/.exec(norm(s.id));
      if (m && m[1] === year && +m[2] > max) max = +m[2];
    });
    var p = function (n) { return String(n).padStart(4, "0"); };
    return "SES-" + year + "-" + p(max + 1);
  }

  function r1(n) { return Math.round(n * 10) / 10; }
  function r2(n) { return Math.round(n * 100) / 100; }
  function buildSummary(sessions) {
    var total = sessions.length;
    var hours = 0, npsSum = 0, npsN = 0, tsSum = 0, tsN = 0, respSum = 0, respN = 0;
    var bySchool = {}, byCourse = {}, monthly = {};
    function bucket(map, key) {
      if (!map[key]) map[key] = { key: key, sessions: 0, npsSum: 0, npsN: 0, tsSum: 0, tsN: 0 };
      return map[key];
    }
    (sessions || []).forEach(function (s) {
      if (s.timeStart != null && s.timeEnd != null && s.timeEnd > s.timeStart)
        hours += (s.timeEnd - s.timeStart) * 24;
      if (s.nps != null) { npsSum += s.nps; npsN++; }
      if (s.trainerScore != null) { tsSum += s.trainerScore; tsN++; }
      if (s.n != null) { respSum += s.n; respN++; }
      var b = bucket(bySchool, s.client || "(no school)");
      b.sessions++;
      if (s.nps != null) { b.npsSum += s.nps; b.npsN++; }
      if (s.trainerScore != null) { b.tsSum += s.trainerScore; b.tsN++; }
      var co = bucket(byCourse, s.code || "(no code)");
      co.sessions++;
      if (s.nps != null) { co.npsSum += s.nps; co.npsN++; }
      if (s.trainerScore != null) { co.tsSum += s.trainerScore; co.tsN++; }
      var m = bucket(monthly, s.dateISO ? s.dateISO.slice(0, 7) : "(no date)");
      m.sessions++;
      if (s.nps != null) { m.npsSum += s.nps; m.npsN++; }
    });
    function finish(map) {
      return Object.keys(map).sort().map(function (k) {
        var b = map[k];
        b.avgNps = b.npsN ? r1(b.npsSum / b.npsN) : null;
        b.avgScore = b.tsN ? r2(b.tsSum / b.tsN) : null;
        return b;
      });
    }
    var schools = finish(bySchool);
    var repeatSessions = schools.filter(function (b) { return b.sessions > 1; })
      .reduce(function (a, b) { return a + b.sessions; }, 0);
    return {
      total: total,
      hours: r1(hours),
      avgNps: npsN ? r1(npsSum / npsN) : null, scoredNps: npsN,
      avgScore: tsN ? r2(tsSum / tsN) : null, scoredTs: tsN,
      respondents: respSum, respSessions: respN,
      repeatRate: total ? Math.round(repeatSessions / total * 100) : 0,
      bySchool: schools, byCourse: finish(byCourse), monthly: finish(monthly)
    };
  }

  /* ---- trainer / course reporting (summary report + trainer profile) ----------
   * Input sessions: { trainer, code, title, client, dateISO, timeStart, timeEnd,
   *                   nps, trainerScore, n, itemAvgs? }  (nps/score may be unwritten) */
  function hoursOf(s) {
    return s.timeStart != null && s.timeEnd != null && s.timeEnd > s.timeStart ? (s.timeEnd - s.timeStart) * 24 : 0;
  }
  function quarterOf(iso) { return iso ? iso.slice(0, 4) + "-Q" + (Math.floor((+iso.slice(5, 7) - 1) / 3) + 1) : ""; }
  // Periods present in the data, newest first: years then quarters.
  function periodsOf(sessions) {
    var ys = {}, qs = {};
    (sessions || []).forEach(function (s) { if (s.dateISO) { ys[s.dateISO.slice(0, 4)] = 1; qs[quarterOf(s.dateISO)] = 1; } });
    var lab = function (k) { return k.length === 4 ? k : k.slice(0, 4) + " " + k.slice(5); };
    return Object.keys(ys).sort().reverse().concat(Object.keys(qs).sort().reverse()).map(function (k) { return { key: k, label: lab(k) }; });
  }
  function inPeriod(iso, key) {
    if (!key) return true;
    if (!iso) return false;
    return key.length === 4 ? iso.slice(0, 4) === key : quarterOf(iso) === key;
  }
  function prevPeriod(key) {
    if (!key) return "";
    if (key.length === 4) return String(+key - 1);
    var y = +key.slice(0, 4), q = +key.slice(6);
    return q > 1 ? y + "-Q" + (q - 1) : (y - 1) + "-Q4";
  }
  function avg(list, f, d) {
    var v = list.map(f).filter(function (x) { return x != null && isFinite(x); });
    if (!v.length) return null;
    var p = Math.pow(10, d);
    return Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length * p) / p;
  }
  function groupBy(list, keyFn) {
    var m = {}, order = [];
    list.forEach(function (s) { var k = keyFn(s); if (!m[k]) { m[k] = []; order.push(k); } m[k].push(s); });
    return order.map(function (k) { return { key: k, items: m[k] }; });
  }
  function distinct(list, f) { var o = {}; list.forEach(function (s) { var v = normKey(f(s)); if (v) o[v] = 1; }); return Object.keys(o).length; }
  function distinctTrainers(list) { var o = {}; list.forEach(function (s) { splitTrainers(s.trainer).forEach(function (t) { o[normKey(t)] = 1; }); }); return Object.keys(o).length; }
  // One entry per (trainer, session): a co-trained session counts for each of its trainers.
  function byTrainer(list) {
    var m = {}, order = [];
    (list || []).forEach(function (s) {
      splitTrainers(s.trainer).forEach(function (t) {
        var k = normKey(t);
        if (!m[k]) { m[k] = { key: k, name: t, items: [] }; order.push(k); }
        m[k].items.push(s);
      });
    });
    return order.map(function (k) { return m[k]; });
  }
  // One row per trainer. prev: the same trainers' sessions in the previous period (for the NPS change).
  function trainerTable(sessions, prev) {
    var prevBy = {};
    byTrainer(prev).forEach(function (g) { prevBy[g.key] = avg(g.items, function (s) { return s.nps; }, 1); });
    return byTrainer(sessions).map(function (g) {
      var it = g.items, courses = courseTable(it);
      var a = avg(it, function (s) { return s.nps; }, 1), p = prevBy[g.key];
      var co = {}, coN = 0;
      it.forEach(function (s) {
        var ns = splitTrainers(s.trainer);
        if (ns.length > 1) coN++;
        ns.forEach(function (t) { if (normKey(t) !== g.key) co[normKey(t)] = co[normKey(t)] || t; });
      });
      return {
        trainer: g.name, coTrained: coN, coWith: Object.keys(co).map(function (k) { return co[k]; }).sort(), sessions: it.length, hours: r1(it.reduce(function (x, s) { return x + hoursOf(s); }, 0)),
        courses: courses.length, topCourse: courses[0] || null, schools: distinct(it, function (s) { return s.client; }),
        respondents: it.reduce(function (x, s) { return x + (s.n || 0); }, 0),
        avgNps: a, avgScore: avg(it, function (s) { return s.trainerScore; }, 2),
        npsChange: a != null && p != null ? r1(a - p) : null,
        noEval: it.filter(function (s) { return s.nps == null && s.trainerScore == null; }).length,
        last: it.reduce(function (x, s) { return s.dateISO && s.dateISO > x ? s.dateISO : x; }, "")
      };
    }).sort(function (x, y) { return y.sessions - x.sessions || (x.trainer < y.trainer ? -1 : 1); });
  }
  // One row per course (code; title from the session or the course list), most conducted first.
  function courseTable(sessions) {
    return groupBy((sessions || []).filter(function (s) { return norm(s.code) || norm(s.title); }), function (s) { return normKey(s.code) || "t:" + normKey(s.title); }).map(function (g) {
      var it = g.items, titled = it.filter(function (s) { return norm(s.title); })[0];
      return { code: norm(it[0].code), title: titled ? norm(titled.title) : "", sessions: it.length,
               hours: r1(it.reduce(function (x, s) { return x + hoursOf(s); }, 0)),
               schools: distinct(it, function (s) { return s.client; }), trainers: distinctTrainers(it),
               avgNps: avg(it, function (s) { return s.nps; }, 1), avgScore: avg(it, function (s) { return s.trainerScore; }, 2) };
    }).sort(function (x, y) { return y.sessions - x.sessions || (x.code < y.code ? -1 : 1); });
  }
  // Headline counts for the top tiles.
  function overview(sessions) {
    var list = sessions || [];
    var scored = list.filter(function (s) { return s.nps != null || s.trainerScore != null; }).length;
    return { trainers: distinctTrainers(list), sessions: list.length,
             hours: r1(list.reduce(function (x, s) { return x + hoursOf(s); }, 0)),
             courses: courseTable(list).length, schools: distinct(list, function (s) { return s.client; }),
             evaluated: scored, coverage: list.length ? Math.round(scored / list.length * 100) : null };
  }
  // Facilitator items averaged over sessions that have them (evals scored in the dashboard).
  function itemAverages(sessions) {
    var out = { n: 0 };
    ["t1", "t2", "t3"].forEach(function (k) { out[k] = avg((sessions || []).filter(function (s) { return s.itemAvgs; }), function (s) { return s.itemAvgs[k]; }, 2); });
    out.n = (sessions || []).filter(function (s) { return s.itemAvgs; }).length;
    return out;
  }

  function buildCourseMap(rows) {
    var byCode = {}, byTitle = {};
    (rows || []).slice(1).forEach(function (r) {
      if (!r) return;
      var code = norm(r[0]), title = norm(r[2]);
      if (!code) return;
      byCode[code] = { title: title, product: norm(r[1]),
        duration: (r[3] === '' || r[3] == null) ? null : Number(r[3]) };
      if (title && !byTitle[title]) byTitle[title] = code;
    });
    return { byCode: byCode, byTitle: byTitle };
  }

  /* ---- per-session feedback report (ported from the v7 standalone Feedback
   * Generator, now offline + in-tab). Pure: returns data, the UI escapes it. */
  var TRIVIAL = /^(?:n\.?\s*\/?\s*a\.?|none|nothing|no|no comment|nil|wala|wla|-+|\.+|x)$/i;
  function isTrivial(s) { var t = norm(s).replace(/[.\s]+$/, ""); return !t || TRIVIAL.test(t); }
  var STOP = {};
  ("a an the and or but if of at by for with about to from in on is are was were be been being have has had do does did " +
   "will would should could can may might must shall i me my we our us you your he she it they them their this that these " +
   "those what when where who why how all any both each few more most other some such no nor not only own same so than " +
   "too very just also as am one its really well get got give make made help useful informative great good nice na none " +
   "nothing session training topic topics able because everything found enjoyed liked learned sessions far time today etc " +
   "since much many thank thanks lot").split(" ").forEach(function (w) { STOP[w] = 1; });
  function topKeywords(texts, k) {
    var counts = {};
    (texts || []).forEach(function (t) {
      if (isTrivial(t)) return;
      var seen = {};
      (String(t).toLowerCase().match(/[a-z][a-z'-]+/g) || []).forEach(function (w) {
        w = w.replace(/^['-]+|['-]+$/g, "");
        if (w.length < 4 || STOP[w] || seen[w]) return;
        seen[w] = 1; counts[w] = (counts[w] || 0) + 1;
      });
    });
    return Object.keys(counts).map(function (w) { return { word: w, count: counts[w] }; })
      .sort(function (a, b) { return b.count - a.count || (a.word < b.word ? -1 : 1); }).slice(0, k || 8);
  }
  var RATING_GROUPS = [
    ["Facilitator", /knowledge.*mastery|communication.*presentation|ability to engage|facilitator|presenter/i],
    ["Session content", /scope.*coverage|relevance|applicab|clarity.*organi|learning aids|materials|virtual campus|\blms\b/i],
    ["Learning experience", /helped increase|relevant.*professional|pace|learning experience|understanding/i],
    ["Overall satisfaction", /overall|satisf/i]
  ];
  function ratingGroup(header) {
    for (var i = 0; i < RATING_GROUPS.length; i++) if (RATING_GROUPS[i][1].test(header)) return RATING_GROUPS[i][0];
    return "Other";
  }
  /* A rating column: at most 8 distinct answers, >= half of them text Likert labels. */
  function ratingColumns(rows, skip) {
    var header = (rows[0] || []).map(norm), data = rows.slice(1), out = [];
    header.forEach(function (h, c) {
      if (!h || skip[c]) return;
      var distinct = {}, nd = 0, dist = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 }, labels = {}, n = 0, sum = 0;
      data.forEach(function (r) {
        var v = norm(r && r[c]);
        if (!v) return;
        if (!distinct[v]) { distinct[v] = 1; nd++; }
      });
      if (!nd || nd > 8) return;
      var textual = Object.keys(distinct).filter(function (v) { return /[a-z]/i.test(v) && mapLikert(v) != null; }).length;
      if (textual / nd < 0.5) return;
      data.forEach(function (r) {
        var raw = norm(r && r[c]), s = mapLikert(raw);
        if (s == null) return;
        dist[s]++; n++; sum += s;
        var lb = raw.replace(/^\d+\s*[-–—.)]\s*/, "");
        (labels[s] = labels[s] || {})[lb] = (labels[s][lb] || 0) + 1;
      });
      if (!n) return;
      var lab = {};
      Object.keys(labels).forEach(function (s) {
        lab[s] = Object.keys(labels[s]).sort(function (a, b) { return labels[s][b] - labels[s][a]; })[0];
      });
      out.push({ col: c, header: h, group: ratingGroup(h), n: n, avg: Math.round(sum / n * 100) / 100,
                 top2: Math.round((dist[4] + dist[5]) / n * 1000) / 10, dist: dist, labels: lab });
    });
    return out;
  }
  function orgKey(s) {
    var k = normKey(s).replace(/\s*\([^)]*\)\s*/g, " ").replace(/[.,]/g, " ").replace(/\s*[-‐-―]\s*/g, " ")
      .replace(/\b(inc|incorporated)\b/g, " ").replace(/\s+/g, " ").trim();
    return /^(school|shool|-|na|n\/a)?$/.test(k) ? "" : k;
  }
  function schoolBreakdown(data, ci) {
    if (ci == null || ci < 0) return [];
    var g = {};
    data.forEach(function (r) {
      var raw = norm(r[ci]), k = orgKey(raw);
      if (!k) return;
      var o = g[k] || (g[k] = { count: 0, sp: {}, acr: {} });
      o.count++; o.sp[raw] = (o.sp[raw] || 0) + 1;
      o.acr[first(normKey(raw).replace(/[^a-z ]+/g, " "))] = 1;
      o.acr[acronymOf(raw).toLowerCase()] = 1;
    });
    // "HCCCI" + "Holy Cross College of Carigara, Inc." in the same file: merge the acronym into the long name.
    function first(k) { return k.split(/\s+/).filter(function (t) { return t.length > 2; }).map(function (t) { return t[0]; }).join(""); }
    Object.keys(g).forEach(function (sk) {
      if (!/^[a-z]{2,6}$/.test(sk)) return;
      var lk = Object.keys(g).filter(function (k) { return k !== sk && k.indexOf(" ") > 0 && g[k].acr[sk]; })[0];
      if (!lk) return;
      g[lk].count += g[sk].count;
      Object.keys(g[sk].sp).forEach(function (s) { g[lk].sp[s] = (g[lk].sp[s] || 0) + g[sk].sp[s]; });
      delete g[sk];
    });
    var total = 0; Object.keys(g).forEach(function (k) { total += g[k].count; });
    return Object.keys(g).map(function (k) {
      var sp = Object.keys(g[k].sp).sort(function (a, b) { return b.length - a.length || g[k].sp[b] - g[k].sp[a]; });
      return { name: sp[0], count: g[k].count, pct: total ? Math.round(g[k].count / total * 1000) / 10 : 0 };
    }).sort(function (a, b) { return b.count - a.count; });
  }
  function firstQuote(texts, words) {
    var list = (texts || []).map(norm).filter(function (t) { return !isTrivial(t) && t.length >= 12 && t.length <= 220; });
    for (var i = 0; i < (words || []).length; i++) {
      var re = new RegExp("\\b" + words[i].word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i");
      for (var j = 0; j < list.length; j++) if (re.test(list[j])) return list[j];
    }
    return list[0] || null;
  }
  var ACTIONS = {
    internet: "Run a 15-minute connectivity check before the session and share a dial-in fallback.",
    connect: "Test the meeting link with participants 15 minutes before the session; share a dial-in fallback.",
    hands: "Make at least 25% of the session hands-on: pair every demo with an immediate practice exercise.",
    practice: "Add a practice block every 20 minutes and a take-home exercise sheet.",
    time: "Reserve 25–30% of the time for practice and add short breaks every 20–25 minutes.",
    pace: "Add a mid-session pace check and adjust the second half; break every 25 minutes.",
    engage: "Open with a short icebreaker tied to the topic and run a mid-session poll or Q&A.",
    example: "Use plain-language, school-specific examples before introducing technical terms.",
    material: "Share the materials 24 hours before and a one-page summary after the session.",
    longer: "Review the session length: split dense topics into a follow-up session."
  };
  function actionFor(words) {
    for (var i = 0; i < words.length; i++)
      for (var k in ACTIONS) if (words[i].word.indexOf(k) >= 0) return ACTIONS[k];
    return null;
  }
  function analyzeEval(rows, roles) {
    roles = roles || {};
    var data = (rows || []).slice(1).filter(function (r) { return r && r.some(function (v) { return norm(v) !== ""; }); });
    var col = function (k) { var c = roles[k]; return c == null || c < 0 ? null : c; };
    var texts = function (k) { var c = col(k); return c == null ? [] : data.map(function (r) { return norm(r[c]); }).filter(function (t) { return !isTrivial(t); }); };
    var nps = calcNPS(col("nps") == null ? [] : data.map(function (r) { return r[col("nps")]; }));
    var ts = calcTrainerScore(data, roles);
    var skip = {};
    ["nps", "school", "c1", "c2", "c3", "cnps"].forEach(function (k) { if (col(k) != null) skip[col(k)] = 1; });
    var ratings = ratingColumns([rows[0] || []].concat(data), skip);
    var overall = ratings.length ? Math.round(ratings.reduce(function (a, r) { return a + r.avg; }, 0) / ratings.length * 100) / 100 : null;
    var satCol = ratings.filter(function (r) { return r.col === col("sat"); })[0] || ratings.filter(function (r) { return r.group === "Overall satisfaction"; })[0];
    var c = { best: texts("c1"), insight: texts("c2"), suggest: texts("c3"), why: texts("cnps") };
    var detrWhy = col("nps") == null || col("cnps") == null ? [] : data.filter(function (r) {
      var t = norm(r[col("nps")]); return t !== "" && isFinite(+t) && +t <= 6;
    }).map(function (r) { return norm(r[col("cnps")]); }).filter(function (t) { return !isTrivial(t); });
    var kwBest = topKeywords(c.best, 6), kwSuggest = topKeywords(c.suggest, 6);
    var sorted = ratings.slice().sort(function (a, b) { return b.avg - a.avg; });
    var best = sorted[0], worst = sorted[sorted.length - 1];
    var summary = { verdict: null, wins: [], watches: [] };
    if (overall != null) summary.verdict = overall >= 4.5 ? "overwhelmingly positive" : overall >= 4 ? "highly positive" :
      overall >= 3.5 ? "positive" : overall >= 3 ? "mixed" : "below expectations";
    if (best && worst && best !== worst) {
      summary.wins.push("Highest rated: " + best.header + " (" + best.avg.toFixed(2) + ")");
      if (worst.avg < 4.3) summary.watches.push("Lowest rated: " + worst.header + " (" + worst.avg.toFixed(2) + ")");
    }
    if (nps.n && nps.detractors) summary.watches.push(nps.detractors + " detractor" + (nps.detractors === 1 ? "" : "s") + " (scored 0–6)");
    if (c.suggest.length) summary.watches.push(c.suggest.length + " improvement suggestion" + (c.suggest.length === 1 ? "" : "s"));
    else summary.wins.push("No improvement suggestions raised");
    var recs = [];
    if (c.suggest.length && kwSuggest.length) {
      var top = kwSuggest[0];
      recs.push({ priority: top.count >= 5 ? "high" : top.count >= 2 ? "medium" : "low", effort: "Quick win", owner: "Facilitator",
        title: "Address the top improvement theme: “" + top.word + "”",
        evidence: c.suggest.length + " of " + data.length + " participants left suggestions; “" + top.word + "” is mentioned " + top.count + " time" + (top.count === 1 ? "" : "s") +
          (kwSuggest.length > 1 ? ", then " + kwSuggest.slice(1, 3).map(function (k) { return "“" + k.word + "” (" + k.count + ")"; }).join(", ") : "") + ".",
        quote: firstQuote(c.suggest, kwSuggest),
        action: actionFor(kwSuggest.slice(0, 5)) || "Track the suggestions and address the top three before the next session.",
        metric: "Fewer than 2 mentions of “" + top.word + "” next session" });
    }
    if (worst && worst.avg < 4.3 && ratings.length > 1) {
      recs.push({ priority: worst.avg < 4 ? "high" : "medium", effort: "Strategic", owner: "Program lead",
        title: "Strengthen “" + worst.header + "”",
        evidence: "Scored " + worst.avg.toFixed(2) + " / 5 — the lowest of " + ratings.length + " rated items.",
        quote: null, action: /pace/i.test(worst.header) ? ACTIONS.pace : "Refresh the materials or sequencing for this item and pilot the change with a small group.",
        metric: "4.5+ on this item next cycle" });
    }
    if (nps.n && nps.detractors) {
      var pct = nps.detractors / nps.n * 100;
      recs.push({ priority: pct >= 10 ? "high" : "medium", effort: "Quick win", owner: "Program lead",
        title: "Follow up with detractors",
        evidence: nps.detractors + " of " + nps.n + " respondents (" + pct.toFixed(1) + "%) scored 0–6 on the recommendation question.",
        quote: detrWhy.filter(function (t) { return t.length >= 8 && t.length <= 220; })[0] || null,
        action: "Within a week, ask what disappointed them and what would have made it a 9 or 10; share what changes.",
        metric: "Half of detractors move to passive or promoter next cycle" });
    }
    if (c.best.length >= 5 && kwBest.length >= 2) {
      recs.push({ priority: "low", effort: "Ongoing", owner: "Program lead",
        title: "Keep what participants praised",
        evidence: "Most praised: " + kwBest.slice(0, 3).map(function (k) { return "“" + k.word + "” (" + k.count + ")"; }).join(", ") + ".",
        quote: firstQuote(c.best, kwBest.slice(0, 3)),
        action: "Anchor the next session on these themes and reuse the techniques that earned the praise.",
        metric: "Same or more positive mentions next cycle" });
    }
    var ord = { high: 0, medium: 1, low: 2 };
    recs.sort(function (a, b) { return ord[a.priority] - ord[b.priority]; });
    return {
      responses: data.length, nps: nps, trainer: ts, ratings: ratings, overall: overall,
      satisfaction: satCol ? satCol.top2 : null, schools: schoolBreakdown(data, col("school")),
      comments: c, detractorReasons: detrWhy, keywords: { best: kwBest, suggest: kwSuggest },
      summary: summary, recommendations: recs.slice(0, 5), responseDate: detectResponseDate(rows)
    };
  }

  var api = {
    norm: norm, normKey: normKey, excelDate: excelDate, timeOf: timeOf,
    detectResponseDate: detectResponseDate, detectResponseSession: detectResponseSession,
    evalFingerprint: evalFingerprint, evalSourceText: evalSourceText, fpOf: fpOf,
    acronymOf: acronymOf, buildAcronymIndex: buildAcronymIndex, nextSessionId: nextSessionId,
    buildSummary: buildSummary, trainerTable: trainerTable, courseTable: courseTable, overview: overview, itemAverages: itemAverages,
    periodsOf: periodsOf, inPeriod: inPeriod, prevPeriod: prevPeriod, hoursOf: hoursOf,
    buildCourseMap: buildCourseMap,
    analyzeEval: analyzeEval, ratingColumns: ratingColumns, topKeywords: topKeywords, isTrivial: isTrivial,
    parseTrackerSheet: parseTrackerSheet, parseFilename: parseFilename,
    suggestSessions: suggestSessions, rankSessions: rankSessions, matchEval: matchEval, detectLayout: detectLayout,
    mapLikert: mapLikert, calcNPS: calcNPS, calcTrainerScore: calcTrainerScore,
    findPrevSession: findPrevSession, buildRoster: buildRoster,
    splitTrainers: splitTrainers, hasTrainer: hasTrainer, joinTrainers: joinTrainers, duplicateSessions: duplicateSessions,
    planWriteback: planWriteback, TRACKER_COLS: TRACKER_COLS
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else { root.QBR = root.QBR || {}; root.QBR.trainersParse = api; }
})(typeof self !== "undefined" ? self : this);
