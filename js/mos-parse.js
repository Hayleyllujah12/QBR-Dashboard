/* =============================================================================
 * mos-parse.js — v1.37.0 MOS (Microsoft Office Specialist) results: pure logic
 * -----------------------------------------------------------------------------
 * No DOM, no storage. Used by js/mos.js in the browser (window.QBR.mosParse) and
 * by tests/mos-parse.cjs in Node (module.exports).
 *
 *  - parseScorePage(items)  one Certiport "EXAM SCORE REPORT" page → attempt
 *                           (items = pdf.js text items {str, transform:[..,x,y], width})
 *  - fileMeta(path)         school / Exam No. / retake / section hints from the file path
 *  - parseRosterRows(name, rows)  one roster sheet (array of arrays) → students
 *  - buildModel(attempts, roster) dedupe, attempt numbering, school year, roster join
 *  - filterAttempts / summarize   filter + KPI helpers used by the MOS page
 *
 * PRIVACY: candidate addresses on the score report are never read into the model
 * (only name, email, Certiport ID). Nothing here persists anything.
 * ========================================================================== */
(function (root) {
  "use strict";

  var MONTHS = ["january", "february", "march", "april", "may", "june", "july",
    "august", "september", "october", "november", "december"];

  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }

  // pdf.js text items → [{x, y, w, s}] (non-empty only)
  function toItems(items) {
    var out = [];
    (items || []).forEach(function (it) {
      if (!it) return;
      var s = it.str != null ? it.str : it.s;
      if (s == null || !String(s).trim()) return;
      var x = it.transform ? it.transform[4] : it.x;
      var y = it.transform ? it.transform[5] : it.y;
      out.push({ x: +x || 0, y: +y || 0, w: +(it.width != null ? it.width : it.w) || 0, s: String(s) });
    });
    return runs(out);
  }

  // Merge touching items on the same baseline into runs: exports often split a word
  // into several items ("EXAM S" + "CORE REPORT", "Exam refe" + "rence #: …").
  function runs(items) {
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var out = [];
    sorted.forEach(function (it) {
      var p = out.length ? out[out.length - 1] : null;
      var gap = p ? it.x - (p.x + p.w) : 0;
      // ≤1.5pt: same word split in two; ≤7pt: a word gap drawn as spacing instead of a space char
      if (p && Math.abs(p.y - it.y) <= 1 && gap <= 7 && it.x >= p.x) {
        p.s += (gap > 1.5 && !/\s$/.test(p.s) && !/^\s/.test(it.s) ? " " : "") + it.s; p.w = it.x + it.w - p.x;
      } else out.push({ x: it.x, y: it.y, w: it.w, s: it.s });
    });
    return out;
  }

  // Group items into lines (same baseline ± tol), top → bottom; each line's items left → right.
  function toLines(items, tol) {
    tol = tol == null ? 2.5 : tol;
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var lines = [];
    sorted.forEach(function (it) {
      var ln = lines.length ? lines[lines.length - 1] : null;
      if (ln && Math.abs(ln.y - it.y) <= tol) ln.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    });
    lines.forEach(function (ln) {
      ln.items.sort(function (a, b) { return a.x - b.x; });
      var txt = "", prevEnd = null;
      ln.items.forEach(function (it) {
        if (prevEnd != null) txt += (it.x - prevEnd > 1.5 ? " " : "");
        txt += it.s; prevEnd = it.x + it.w;
      });
      ln.text = norm(txt);
    });
    return lines;
  }

  function find(items, re) { for (var i = 0; i < items.length; i++) if (re.test(norm(items[i].s))) return items[i]; return null; }

  // Exam title clean-up: the score-report font drops the "ffi" ligature ("Oce", "O ce").
  function cleanTitle(t) {
    t = norm(t).replace(/[-]/g, "ffi").replace(/ﬃ/g, "ffi");
    t = t.replace(/\bO\s*(?:ffi)?\s*ce\b/g, "Office");
    t = t.replace(/\s*Registration ID:.*$/i, "");
    return norm(t);
  }

  // "Excel" | "Excel Expert" | "PowerPoint" | "Word" | "Word Expert" | "Outlook" | "Access" | "AZ-900" …
  function examFamily(title) {
    var t = String(title || "");
    var code = t.match(/\b([A-Z]{2,3}-\d{3})\b/);
    if (code) return code[1];
    var m = t.match(/\b(Excel|Word|PowerPoint|Outlook|Access)\b(\s+Expert)?/i);
    if (m) {
      var base = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
      if (/^powerpoint$/i.test(m[1])) base = "PowerPoint";
      return base + (m[2] ? " Expert" : "");
    }
    return t ? t.replace(/^Microsoft\s+/i, "").replace(/\s*\(.*$/, "") : "Other";
  }
  function examVersion(title) {
    var m = String(title || "").match(/\(([^)]+)\)/);
    return m ? m[1] : "";
  }

  // "September 13, 2021" → "2021-09-13" ('' when unknown)
  function parseDate(s) {
    var m = norm(s).match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})/);
    if (!m) return "";
    var mi = MONTHS.indexOf(m[1].toLowerCase());
    if (mi < 0) { mi = MONTHS.map(function (x) { return x.slice(0, 3); }).indexOf(m[1].slice(0, 3).toLowerCase()); }
    if (mi < 0) return "";
    return m[3] + "-" + String(mi + 1).padStart(2, "0") + "-" + String(+m[2]).padStart(2, "0");
  }
  // School year runs June → May: 2025-06-01 … 2026-05-31 = "SY 2025-2026"
  function schoolYear(iso) {
    if (!iso) return "";
    var y = +iso.slice(0, 4), mo = +iso.slice(5, 7);
    var s = mo >= 6 ? y : y - 1;
    return "SY " + s + "-" + (s + 1);
  }

  /* ---- one score-report page ------------------------------------------------ */
  function parseScorePage(rawItems) {
    var items = toItems(rawItems);
    if (!find(items, /EXAM SCORE REPORT/i)) return null;
    var cand = find(items, /^CANDIDATE$/i);
    var examH = null;
    items.forEach(function (it) { if (!examH && /^EXAM$/i.test(norm(it.s)) && (!cand || Math.abs(it.y - cand.y) < 3)) examH = it; });
    var results = find(items, /^RESULTS$/i);
    if (!cand || !examH) return { _err: "no CANDIDATE/EXAM header" };
    var top = cand.y - 2, bottom = results ? results.y + 6 : -Infinity, colX = examH.x - 4;

    var head = items.filter(function (it) { return it.y < top && it.y > bottom; });
    var left = toLines(head.filter(function (it) { return it.x < colX; }));
    var right = toLines(head.filter(function (it) { return it.x >= colX; }));

    var r = { name: left.length ? left[0].text : "", email: "", ref: "", dateText: "", certiportId: "" };
    left.forEach(function (ln) { if (!r.email && /@/.test(ln.text) && !/\s/.test(ln.text)) r.email = ln.text.toLowerCase(); });
    var title = [];
    right.forEach(function (ln) {
      var t = ln.text, m;
      if ((m = t.match(/^Exam reference\s*#\s*:?\s*(\S+)/i))) { r.ref = m[1]; return; }
      if ((m = t.match(/^Date\s*:\s*(.+)$/i))) { r.dateText = m[1]; return; }
      if ((m = t.match(/^ID\s*:\s*(.+)$/i))) { r.certiportId = m[1].trim(); return; }
      if (/^Registration ID/i.test(t)) return;
      if (!r.ref) title.push(t);
    });
    r.examTitle = cleanTitle(title.join(" "));
    r.exam = examFamily(r.examTitle);
    r.examVersion = examVersion(r.examTitle);
    r.date = parseDate(r.dateText);

    // FINAL SCORE column (right of SECTION ANALYSIS): "Required Score  700", "Your Score  875", OUTCOME → Pass/Fail
    var fin = find(items, /^FINAL SCORE$/i);
    var sec = find(items, /^SECTION ANALYSIS$/i);
    var foot = null; items.forEach(function (it) { if (/©|Certiport and the Certiport/.test(it.s) && (!foot || it.y > foot.y)) foot = it; });
    var footY = foot ? foot.y + 2 : -Infinity;
    r.required = null; r.score = null; r.outcome = null;
    if (fin) {
      var finX = fin.x - 4;
      var col = items.filter(function (it) { return it.x >= finX && it.y < fin.y - 1 && it.y > footY; });
      var lines = toLines(col);
      var nums = [];
      lines.forEach(function (ln) {
        var m = ln.text.match(/^(Required Score|Your Score)?\s*(\d{1,4})$/i);
        if (m) nums.push({ label: (m[1] || "").toLowerCase(), v: +m[2] });
      });
      nums.forEach(function (n) {
        if (n.label === "your score") r.score = n.v;
        else if (n.label === "required score") r.required = n.v;
      });
      var unl = nums.filter(function (n) { return !n.label; });
      // unlabeled numbers (label drawn as an image on some exports): top one = required, next = your score
      if (r.required == null && unl.length) r.required = unl.shift().v;
      if (r.score == null && unl.length) r.score = unl.shift().v;
      var oc = null;
      lines.forEach(function (ln, i) { if (/^OUTCOME$/i.test(ln.text)) oc = i; });
      if (oc != null) for (var k = oc + 1; k < lines.length; k++) {
        var mm = lines[k].text.match(/^(Pass|Fail)\b/i);
        if (mm) { r.outcome = mm[1].toUpperCase(); break; }
      }
    }
    if (r.required == null) r.required = 700;
    if (r.outcome == null && r.score != null) r.outcome = r.score >= r.required ? "PASS" : "FAIL";

    // SECTION ANALYSIS (skill areas): "<objective>  NN%" left of FINAL SCORE
    r.skills = [];
    if (sec) {
      var maxX = fin ? fin.x - 4 : Infinity;
      toLines(items.filter(function (it) { return it.y < sec.y - 1 && it.y > footY && it.x < maxX; })).forEach(function (ln) {
        var m = ln.text.match(/^(.*\S)\s+(\d{1,3})%$/);
        if (m) r.skills.push({ area: norm(m[1]), pct: +m[2] });
      });
    }
    if (r.score == null) return { _err: "no score", ref: r.ref };
    return r;
  }

  /* ---- file path → school / Exam No. / retake / section hints ------------- */
  function fileMeta(path, schoolOverride) {
    var parts = String(path || "").replace(/\\/g, "/").split("/").filter(Boolean);
    var base = parts.length ? parts[parts.length - 1] : "";
    var school = "";
    for (var i = 0; i < parts.length - 1; i++) {
      if (/^\d+\s*[-.]\s*\S/.test(parts[i])) { school = parts[i]; break; }
    }
    if (!school && schoolOverride) school = schoolOverride;
    if (!school && parts.length > 1) school = parts[0];
    school = norm(school.replace(/^\d+\s*[-.]\s*/, ""));
    var m = base.match(/Exam\s*(\d{4}(?:\.\d+)?)/i);
    var examNo = m ? m[1] : "";
    var retake = /\.\d/.test(examNo) || /retake/i.test(base);
    var sec = base.match(/\b(1[0-2][A-Z](?:&[A-Z])?)\b/);
    var gr = base.match(/\bGR\s?(1[0-2])\b/i);
    return {
      school: school || "Unassigned school",
      examNo: examNo,
      examBase: examNo.replace(/\..*$/, ""),
      retakeFile: retake,
      sectionHint: sec ? sec[1] : "",
      gradeHint: gr ? gr[1] : (sec ? sec[1].slice(0, 2) : ""),
      file: parts.join("/") || base,
      fileName: base
    };
  }

  /* ---- roster ---------------------------------------------------------------- */
  // A section sheet: a header row containing NAME (+ Email). Section = sheet name ("11A").
  function parseRosterRows(sheetName, rows) {
    var out = [];
    var secM = String(sheetName || "").match(/^\s*(1[0-2]\s*-?\s*[A-Z](?:&[A-Z])?)\s*$/i);
    if (!secM) return out;
    var section = secM[1].replace(/[\s-]/g, "").toUpperCase();
    var hi = -1, cols = {};
    for (var r = 0; r < Math.min(rows.length, 15); r++) {
      var row = rows[r] || [];
      var c = {};
      row.forEach(function (v, j) {
        var h = norm(v).toLowerCase();
        if (h === "name" || /^(student|full)\s*name$/.test(h)) c.name = j;
        else if (/e-?mail/.test(h)) c.email = j;
        else if (/^(student\s*(no|number|id)|id\s*no|lrn)\b/.test(h)) c.studentNo = j;
        else if (/gmetrix/.test(h)) c.gmetrix = j;
      });
      if (c.name != null) { hi = r; cols = c; break; }
    }
    if (hi < 0) return out;
    for (var k = hi + 1; k < rows.length; k++) {
      var rr = rows[k] || [];
      var name = norm(rr[cols.name]);
      if (!name || /^name$/i.test(name)) continue;
      var sn = cols.studentNo != null ? norm(rr[cols.studentNo]) : "";
      if (!sn) { // unlabeled student-number column (e.g. the column before NAME holding 7-digit ids)
        for (var j = 0; j < cols.name; j++) { var v = norm(rr[j]); if (/^\d{6,}$/.test(v)) { sn = v; break; } }
      }
      out.push({
        section: section, grade: section.slice(0, 2), name: name,
        email: cols.email != null ? norm(rr[cols.email]).toLowerCase() : "",
        studentNo: sn
      });
    }
    return out;
  }

  // name keys: roster "SURNAME, First Second, Middle" ↔ report "First Second Middle Surname"
  function tokens(s) {
    return norm(String(s || "").normalize ? String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "") : s)
      .toLowerCase().replace(/[^a-z0-9ñ\s-]/g, " ").split(/[\s-]+/).filter(Boolean);
  }
  function rosterKey(st) {
    var p = String(st.name).split(",");
    var sur = tokens(p[0]), first = tokens(p[1] || "");
    return { sur: sur, first: first, all: tokens(st.name) };
  }
  // Display name "SURNAME, First Middle" for a report name "First Middle Surname" (best effort; roster wins)
  function displayName(n) { return norm(n); }

  function matchRoster(att, roster, idx) {
    if (!roster || !roster.length) return null;
    if (att.email && idx.byEmail[att.email]) return { st: idx.byEmail[att.email], how: "email" };
    var set = {}; tokens(att.name).forEach(function (t) { set[t] = 1; });
    var hits = roster.filter(function (st) {
      if (st.school && att.school && st.school !== att.school) return false;
      var k = st._key || (st._key = rosterKey(st));
      if (!k.sur.length || !k.first.length) return false;
      return k.sur.every(function (t) { return set[t]; }) && set[k.first[0]];
    });
    if (hits.length === 1) return { st: hits[0], how: "name" };
    if (hits.length > 1) return { st: null, how: "ambiguous", candidates: hits };
    return null;
  }

  /* ---- model ---------------------------------------------------------------- */
  function bandOf(score) {
    if (score == null) return "";
    if (score < 700) return "<700";
    if (score < 800) return "700–799";
    if (score < 900) return "800–899";
    return "900–1000";
  }
  var BANDS = ["<700", "700–799", "800–899", "900–1000"];

  // attempts: parsed pages with fileMeta merged in. roster: [{school, section, name, email, studentNo}]
  // overrides: { [ref]: {section, studentKey} } — manual fixes from the Check panel (memory only)
  function buildModel(attempts, roster, overrides) {
    overrides = overrides || {};
    roster = roster || [];
    var idx = { byEmail: {} };
    roster.forEach(function (st) { if (st.email) idx.byEmail[st.email] = st; });
    var seen = {}, dupes = 0, rows = [];
    attempts.forEach(function (a) {
      var key = a.ref || (a.file + "#" + a.page);
      if (seen[key]) { dupes++; return; }
      seen[key] = 1;
      var r = {};
      for (var p in a) if (Object.prototype.hasOwnProperty.call(a, p)) r[p] = a[p];
      r.key = key;
      var m = matchRoster(r, roster, idx);
      r.match = m ? m.how : "none";
      r.candidates = m && m.candidates ? m.candidates : null;
      var st = m && m.st;
      var ov = overrides[key];
      if (ov && ov.student) { st = ov.student; r.match = "manual"; }
      r.studentNo = st ? st.studentNo : "";
      r.rosterName = st ? st.name : "";
      r.section = (ov && ov.section) || (st ? st.section : "") || a.sectionHint || "";
      r.grade = r.section ? r.section.slice(0, 2) : (a.gradeHint || "");
      r.month = r.date ? r.date.slice(0, 7) : "";
      r.year = r.date ? r.date.slice(0, 4) : "";
      r.sy = schoolYear(r.date);
      r.band = bandOf(r.score);
      r.pass = r.outcome === "PASS";
      r.studentKey = st ? ("roster:" + (st.studentNo || st.email || st.name) + "|" + (st.school || r.school))
        : (r.certiportId ? "cid:" + r.certiportId.toLowerCase() : "name:" + tokens(r.name).sort().join(" ") + "|" + r.school);
      rows.push(r);
    });
    // attempt # per student + exam family, ordered by date then ref
    var groups = {};
    rows.forEach(function (r) { var g = r.studentKey + "|" + r.exam; (groups[g] = groups[g] || []).push(r); });
    Object.keys(groups).forEach(function (g) {
      groups[g].sort(function (a, b) { return (a.date || "").localeCompare(b.date || "") || String(a.ref).localeCompare(String(b.ref)); });
      groups[g].forEach(function (r, i) { r.attemptNo = i + 1; r.isRetake = i > 0 || !!r.retakeFile; r.attempts = groups[g].length; });
    });
    rows.sort(function (a, b) {
      return a.school.localeCompare(b.school) || (a.section || "~").localeCompare(b.section || "~") ||
        String(a.rosterName || a.name).localeCompare(String(b.rosterName || b.name)) || (a.date || "").localeCompare(b.date || "");
    });
    return { rows: rows, duplicates: dupes, rosterCount: roster.length };
  }

  /* ---- filters + summaries ----------------------------------------------- */
  // f: {school, section, q, exam, examNo, outcome('PASS'|'FAIL'|''), min, max, year, month, sy, firstOnly}
  function filterAttempts(rows, f) {
    f = f || {};
    var q = norm(f.q).toLowerCase();
    return rows.filter(function (r) {
      if (f.school && r.school !== f.school) return false;
      if (f.section && (r.section || "(none)") !== f.section) return false;
      if (f.exam && r.exam !== f.exam) return false;
      if (f.examNo && r.examBase !== f.examNo && r.examNo !== f.examNo) return false;
      if (f.outcome && r.outcome !== f.outcome) return false;
      if (f.min != null && f.min !== "" && r.score < +f.min) return false;
      if (f.max != null && f.max !== "" && r.score > +f.max) return false;
      if (f.year && r.year !== f.year) return false;
      if (f.month && r.month !== f.month) return false;
      if (f.sy && r.sy !== f.sy) return false;
      if (f.firstOnly && r.attemptNo !== 1) return false;
      if (q && (r.name + " " + r.rosterName + " " + r.email + " " + r.certiportId + " " + r.studentNo).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
  }
  function summarize(rows) {
    var n = rows.length, passed = 0, sum = 0, hi = null, lo = null, first = 0, firstPass = 0;
    rows.forEach(function (r) {
      if (r.pass) passed++;
      sum += r.score;
      if (hi == null || r.score > hi) hi = r.score;
      if (lo == null || r.score < lo) lo = r.score;
      if (r.attemptNo === 1) { first++; if (r.pass) firstPass++; }
    });
    var students = {}; rows.forEach(function (r) { students[r.studentKey] = 1; });
    return {
      attempts: n, passed: passed, failed: n - passed,
      passRate: n ? passed / n : null,
      firstAttempts: first, firstPassRate: first ? firstPass / first : null,
      avg: n ? sum / n : null, high: hi, low: lo,
      students: Object.keys(students).length
    };
  }
  // group → [{key, attempts, passed, passRate}] sorted by key
  function groupRate(rows, keyFn) {
    var g = {};
    rows.forEach(function (r) {
      var k = keyFn(r); if (k == null || k === "") return;
      var o = g[k] || (g[k] = { key: k, attempts: 0, passed: 0, sum: 0 });
      o.attempts++; if (r.pass) o.passed++; o.sum += r.score;
    });
    return Object.keys(g).sort().map(function (k) {
      var o = g[k]; o.passRate = o.attempts ? o.passed / o.attempts : null; o.avg = o.attempts ? o.sum / o.attempts : null; return o;
    });
  }
  // weakest skill areas: average % per (exam, area)
  function skillAreas(rows) {
    var g = {};
    rows.forEach(function (r) {
      (r.skills || []).forEach(function (s) {
        var k = r.exam + " · " + s.area;
        var o = g[k] || (g[k] = { exam: r.exam, area: s.area, n: 0, sum: 0, below50: 0 });
        o.n++; o.sum += s.pct; if (s.pct < 50) o.below50++;
      });
    });
    return Object.keys(g).map(function (k) { var o = g[k]; o.avg = o.sum / o.n; return o; })
      .sort(function (a, b) { return a.avg - b.avg; });
  }

  var api = {
    VERSION: "1.37.0",
    toItems: toItems, toLines: toLines, cleanTitle: cleanTitle, examFamily: examFamily, examVersion: examVersion,
    parseDate: parseDate, schoolYear: schoolYear, parseScorePage: parseScorePage, fileMeta: fileMeta,
    parseRosterRows: parseRosterRows, tokens: tokens, matchRoster: matchRoster, buildModel: buildModel,
    filterAttempts: filterAttempts, summarize: summarize, groupRate: groupRate, skillAreas: skillAreas,
    bandOf: bandOf, BANDS: BANDS, displayName: displayName
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else { root.QBR = root.QBR || {}; root.QBR.mosParse = api; }
})(typeof window !== "undefined" ? window : this);
