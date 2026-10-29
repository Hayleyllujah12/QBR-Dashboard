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
    // fallback: the top folder, unless it is a generic folder name ("Archive SY 2021 - 2022", "Exam Results", "SY 25-26")
    if (!school && parts.length > 1 && !/^(archive|exam\s*results?|results?|sy\s*\d|school\s*year|\d{4}|certificates?|certs?\b)/i.test(parts[0])) school = parts[0];
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
      // export rows carry the Certiport Test Center as school (a group of schools) → match any roster
      if (!att.schoolFromCenter && st.school && att.school && st.school !== att.school) return false;
      var k = st._key || (st._key = rosterKey(st));
      if (!k.sur.length || !k.first.length) return false;
      return k.sur.every(function (t) { return set[t]; }) && set[k.first[0]];
    });
    if (hits.length === 1) return { st: hits[0], how: "name" };
    if (hits.length > 1) return { st: null, how: "ambiguous", candidates: hits };
    return null;
  }

  /* ---- Certiport "Results … Plus" export (.xlsx) ------------------------------ */
  // Sheet layout: title row ("Results (264) Plus"), header row with "Exam Reference Number", data, footer
  // ("User: 0", "Execution Time: …", "© … Pearson"). Date of Birth, Voucher / License and Payment Type are
  // NEVER read (minors' PII / voucher codes).
  var EXPORT_COLS = {
    centerId: /^test\s*center\s*id$/, center: /^test\s*center\s*name$/, first: /^first\s*name$/, last: /^last\s*name$/,
    alt: /^alternate\s*name$/, studentId: /^student\s*\/?\s*employee\s*id$/, username: /^user\s*name$/,
    ref: /^exam\s*reference\s*(number|no\.?|#)$/, exam: /^exam$/, program: /^program\s*name$/, level: /^exam\s*level$/,
    cert: /^certification\s*name$/, date: /^exam\s*date$/, score: /^score$/, result: /^result$/, timeUsed: /^time\s*used$/,
    proctor: /^proctor\s*name$/, station: /^station$/, group: /^exam\s*group$/, teacher: /^teacher$/
  };
  function exportHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 15); r++) {
      var c = {}, row = rows[r] || [];
      row.forEach(function (v, j) {
        var h = norm(v).toLowerCase();
        Object.keys(EXPORT_COLS).forEach(function (k) { if (c[k] == null && EXPORT_COLS[k].test(h)) c[k] = j; });
      });
      if (c.ref != null && c.score != null && c.result != null) return { row: r, cols: c };
    }
    return null;
  }
  function pad2(n) { return String(n).padStart(2, "0"); }
  // Date object | Excel serial | "2026-10-08 11:40:00" | "10/8/2026 11:40 AM" → {date:"YYYY-MM-DD", time:"HH:MM"}
  function parseAnyDate(v) {
    if (v == null || v === "") return { date: "", time: "" };
    if (v instanceof Date && !isNaN(v)) return { date: v.getFullYear() + "-" + pad2(v.getMonth() + 1) + "-" + pad2(v.getDate()), time: pad2(v.getHours()) + ":" + pad2(v.getMinutes()) };
    if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) {
      var ms = Math.round((v - 25569) * 86400000), d = new Date(ms);
      return { date: d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate()), time: pad2(d.getUTCHours()) + ":" + pad2(d.getUTCMinutes()) };
    }
    var t = norm(v), m;
    if ((m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/))) return { date: m[1] + "-" + pad2(m[2]) + "-" + pad2(m[3]), time: m[4] ? pad2(m[4]) + ":" + m[5] : "" };
    if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]M)?)?/i))) {
      var y = +m[3] < 100 ? 2000 + +m[3] : +m[3], h = m[4] ? +m[4] : null;
      if (h != null && m[6]) { if (/pm/i.test(m[6]) && h < 12) h += 12; if (/am/i.test(m[6]) && h === 12) h = 0; }
      return { date: y + "-" + pad2(m[1]) + "-" + pad2(m[2]), time: h != null ? pad2(h) + ":" + m[5] : "" };
    }
    var iso = parseDate(t);
    return { date: iso, time: "" };
  }
  // ALL-CAPS names from the export → "Title Case" for display (mixed-case names are left as typed)
  function nameCase(s) {
    s = norm(s);
    if (!s || s !== s.toUpperCase() || !/[A-Z]/.test(s)) return s;
    return s.toLowerCase().replace(/(^|[\s\-'’.])([a-zñ])/g, function (x, a, b) { return a + b.toUpperCase(); });
  }
  function parseExportRows(rows, fileName) {
    var h = exportHeader(rows);
    if (!h) return null;
    var c = h.cols, out = [], skipped = 0;
    function g(row, k) { return c[k] == null ? "" : row[c[k]]; }
    for (var i = h.row + 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var ref = norm(g(row, "ref"));
      if (!/^\d{5,}$/.test(ref)) { if (row.some(function (v) { return norm(v); })) skipped++; continue; }   // footer / blank
      var first = nameCase(g(row, "first")), last = nameCase(g(row, "last"));
      var user = norm(g(row, "username"));
      var sid = norm(g(row, "studentId"));
      if (/invalid|no id/i.test(sid)) sid = "";
      var dt = parseAnyDate(g(row, "date"));
      var res = norm(g(row, "result")).toLowerCase();
      var sc = norm(g(row, "score")), score = /^\d{1,4}$/.test(sc) ? +sc : null;
      var title = cleanTitle(g(row, "exam"));
      var tu = norm(g(row, "timeUsed"));
      var center = norm(g(row, "center"));
      out.push({
        source: "export", ref: ref,
        name: norm(first + " " + last), firstName: first, lastName: last, altName: norm(g(row, "alt")),
        email: /@/.test(user) ? user.toLowerCase() : "", certiportId: user, studentIdExport: sid,
        examTitle: title, exam: examFamily(title), examVersion: examVersion(title), certName: norm(g(row, "cert")),
        program: norm(g(row, "program")), level: norm(g(row, "level")),
        date: dt.date, time: dt.time, dateText: norm(g(row, "date")),
        score: score, required: 700,
        outcome: /^pass/.test(res) ? "PASS" : /^fail/.test(res) ? "FAIL" : res ? "IN PROGRESS" : (score != null ? (score >= 700 ? "PASS" : "FAIL") : "IN PROGRESS"),
        timeUsed: /^\d+$/.test(tu) ? +tu : null,
        proctor: norm(g(row, "proctor")), station: norm(g(row, "station")),
        examGroup: norm(g(row, "group")), teacher: norm(g(row, "teacher")),
        testCenterId: norm(g(row, "centerId")), testCenter: center,
        school: center || "Unassigned school", schoolFromCenter: true,
        examNo: "", examBase: "", retakeFile: false, sectionHint: "", gradeHint: "",
        skills: [], file: fileName || "", fileName: fileName || ""
      });
    }
    return { attempts: out, skipped: skipped, header: h.row };
  }
  // "Sisters of Silver Pine of Crestview Learning Ministry" → "SSPCLM" (Phase 2: editable per school profile)
  var ACR_SKIP = { of: 1, the: 1, and: 1, "&": 1, de: 1, la: 1, del: 1, for: 1, "-": 1, inc: 1 };
  function acronym(name) {
    var p = String(name || "").replace(/\(([A-Z0-9]{2,})\)/, " $1 ").split(/[\s,.]+/).filter(Boolean);
    var paren = String(name || "").match(/\(([A-Z0-9]{2,8})\)/);
    if (paren) return paren[1];
    var a = p.filter(function (w) { return !ACR_SKIP[w.toLowerCase()]; }).map(function (w) { return /^[A-Z0-9]{2,5}$/.test(w) ? w : w.charAt(0).toUpperCase(); }).join("");
    return (a || "SCH").replace(/[^A-Z0-9]/g, "").slice(0, 8) || "SCH";
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
  // opts.acronyms: { [school]: "ABCD" } from the school profiles (overrides the automatic acronym)
  function buildModel(attempts, roster, overrides, opts) {
    var acrs = (opts && opts.acronyms) || {};
    overrides = overrides || {};
    roster = roster || [];
    var idx = { byEmail: {} };
    roster.forEach(function (st) { if (st.email) idx.byEmail[st.email] = st; });
    var seen = {}, dupes = 0, merged = 0, rows = [], byKey = {};
    // Same Exam reference # from a PDF page and an export row → one attempt: the export gives proctor /
    // station / time / test center; the PDF gives skill areas, Exam No. and the school folder.
    var EXPORT_ONLY = ["proctor", "station", "timeUsed", "examGroup", "teacher", "testCenter", "testCenterId", "certName", "time",
      "studentIdExport", "firstName", "lastName", "altName", "program", "level"];
    var bucket = [];
    attempts.forEach(function (a) {
      var key = a.ref || (a.file + "#" + a.page);
      var prev = byKey[key];
      if (!prev) { byKey[key] = { pdf: a.source === "export" ? null : a, exp: a.source === "export" ? a : null }; bucket.push(key); return; }
      if (a.source === "export" ? prev.exp : prev.pdf) { dupes++; return; }
      if (a.source === "export") prev.exp = a; else prev.pdf = a;
      merged++;
    });
    bucket.forEach(function (key) {
      var e = byKey[key], a = e.pdf || e.exp, r = {}, p;
      for (p in a) if (Object.prototype.hasOwnProperty.call(a, p)) r[p] = a[p];
      if (e.pdf && e.exp) {
        EXPORT_ONLY.forEach(function (k) { if (e.exp[k] != null && e.exp[k] !== "") r[k] = e.exp[k]; });
        if (!r.email && e.exp.email) r.email = e.exp.email;
        if (r.school === "Unassigned school") { r.school = e.exp.school; r.schoolFromCenter = true; } else r.schoolFromCenter = false;
        if (e.exp.score != null) { r.score = e.exp.score; r.outcome = e.exp.outcome; }
        r.source = "pdf+export";
      }
      r.key = key;
      var m = matchRoster(r, roster, idx);
      r.match = m ? m.how : "none";
      r.candidates = m && m.candidates ? m.candidates : null;
      var st = m && m.st;
      var ov = overrides[key];
      if (ov && ov.student) { st = ov.student; r.match = "manual"; }
      r.studentNo = st ? st.studentNo : (r.studentIdExport || "");
      if (st && st.school && r.schoolFromCenter) r.school = st.school;     // roster school beats the Test Center
      r.rosterName = st ? st.name : "";
      r.section = (ov && ov.section) || (st ? st.section : "") || a.sectionHint || "";
      r.grade = r.section ? r.section.slice(0, 2) : (a.gradeHint || "");
      r.month = r.date ? r.date.slice(0, 7) : "";
      r.year = r.date ? r.date.slice(0, 4) : "";
      r.sy = schoolYear(r.date);
      r.band = bandOf(r.score);
      r.pending = r.score == null;                                         // "In Progress" in the export
      // passing = score of 700 or more (exactly 700 passes); the report's own outcome only when there is no score
      r.pass = r.score != null ? r.score >= (r.required || 700) : r.outcome === "PASS";
      // same student across takes: roster > email > Certiport ID > name (emails are how retakes are matched)
      r.studentKey = st ? ("roster:" + (st.studentNo || st.email || st.name) + "|" + (st.school || r.school))
        : r.email ? "email:" + r.email
        : (r.certiportId ? "cid:" + r.certiportId.toLowerCase() : "name:" + tokens(r.name).sort().join(" ") + "|" + r.school);
      rows.push(r);
    });
    // batch: Exam No. from the PDF file name when known, else auto code per school + exam day (ACR-YYYYMMDD-NN)
    rows.forEach(function (r) {
      var acr = String(acrs[r.school] || "").replace(/[^A-Za-z0-9]/g, "").toUpperCase() || acronym(r.school);
      r.batchAuto = r.date ? acr + "-" + r.date.replace(/-/g, "") + "-01" : "";
      r.batch = r.examBase || r.batchAuto;
    });
    // attempt # per student + exam family, ordered by date then ref
    // review fix: a take matched to the roster by name and another take of the same student that only has the
    // email must pair up → rows whose email belongs to a roster-matched row adopt that roster key
    var emailKey = {};
    rows.forEach(function (r) { if (r.email && r.studentKey.indexOf("roster:") === 0) emailKey[r.email] = r.studentKey; });
    rows.forEach(function (r) { if (r.email && r.studentKey.indexOf("roster:") !== 0 && emailKey[r.email]) r.studentKey = emailKey[r.email]; });
    var groups = {};
    rows.forEach(function (r) { var g = r.studentKey + "|" + r.exam; (groups[g] = groups[g] || []).push(r); });
    Object.keys(groups).forEach(function (g) {
      // review fix: only SCORED takes are numbered (an "In Progress" row with no date used to become take #1);
      // undated scored rows sort last
      var scored = groups[g].filter(function (r) { return !r.pending; }), pend = groups[g].filter(function (r) { return r.pending; });
      scored.sort(function (a, b) { return ((a.date || "9999") + (a.time || "")).localeCompare((b.date || "9999") + (b.time || "")) || String(a.ref).localeCompare(String(b.ref)); });
      scored.forEach(function (r, i) {
        r.attemptNo = i + 1; r.isRetake = i > 0 || !!r.retakeFile; r.attempts = scored.length;
        // retake rule (up to 2 takes per exam): a first-time taker who failed may retake once
        r.retakeEligible = r.attemptNo === 1 && !r.pass;
        r.retakeUsed = r.retakeEligible && r.attempts > 1;
        r.noRetakesLeft = r.attemptNo >= 2 && !r.pass;
        r.overLimit = r.attemptNo > 2;
      });
      pend.forEach(function (r, i) {
        r.attemptNo = scored.length + i + 1; r.isRetake = r.attemptNo > 1; r.attempts = scored.length;
        r.retakeEligible = false; r.retakeUsed = false; r.noRetakesLeft = false; r.overLimit = r.attemptNo > 2;
      });
    });
    rows.sort(function (a, b) {
      return a.school.localeCompare(b.school) || (a.section || "~").localeCompare(b.section || "~") ||
        String(a.rosterName || a.name).localeCompare(String(b.rosterName || b.name)) || (a.date || "").localeCompare(b.date || "");
    });
    return { rows: rows, duplicates: dupes, merged: merged, rosterCount: roster.length,
      pending: rows.filter(function (r) { return r.pending; }).length };
  }

  /* ---- certificates (v1.41.0) ----------------------------------------------
   * Certiport certificate exports: one certificate per page, 5 text lines —
   * name · certification ("Excel Associate") · version ("Microsoft 365 Apps") ·
   * date ("April 20, 2023") · verification code (xxxx-xxxx / xxxxx-xxxx, unique per
   * certificate). No email / Exam reference # on a certificate. */
  var CERT_CODE = /^[A-Za-z0-9]{4,6}-[A-Za-z0-9]{4,6}$/;
  var CERT_LEVEL = /\b(Associate|Expert|Master|Specialist|Fundamentals|Professional)\b/i;
  var CERT_APP = /\b(Word|Excel|PowerPoint|Outlook|Access|Office|Azure|Security|Power Platform|[A-Z]{2}-\d{3})\b/;
  function parseCertificatePage(rawItems) {
    var lines = toLines(toItems(rawItems)).filter(function (l) { return l.text; });
    if (lines.length < 4) return null;
    if (lines.some(function (l) { return /EXAM SCORE REPORT|SECTION ANALYSIS|Exam reference/i.test(l.text); })) return null;
    var ci = -1, di = -1, ki = -1, i;
    for (i = lines.length - 1; i >= 0; i--) if (CERT_CODE.test(lines[i].text)) { ci = i; break; }
    if (ci < 0) return null;
    for (i = ci - 1; i >= 0; i--) if (parseDate(lines[i].text)) { di = i; break; }
    if (di < 1) return null;
    for (i = 1; i < di; i++) if (CERT_LEVEL.test(lines[i].text) && CERT_APP.test(lines[i].text)) { ki = i; break; }
    if (ki < 1) return null;                                    // a name line must sit above the certification
    var cert = cleanTitle(lines[ki].text);
    var lv = cert.match(CERT_LEVEL);
    return {
      name: norm(lines.slice(0, ki).map(function (l) { return l.text; }).join(" ")),
      certName: cert, family: examFamily(cert), level: lv ? lv[1].charAt(0).toUpperCase() + lv[1].slice(1).toLowerCase() : "",
      version: norm(lines.slice(ki + 1, di).map(function (l) { return l.text; }).join(" ")),
      dateText: lines[di].text, date: parseDate(lines[di].text), code: lines[ci].text
    };
  }
  function dayDiff(a, b) { return a && b ? Math.round(Math.abs(Date.UTC(+a.slice(0, 4), +a.slice(5, 7) - 1, +a.slice(8, 10)) - Date.UTC(+b.slice(0, 4), +b.slice(5, 7) - 1, +b.slice(8, 10))) / 864e5) : null; }
  // 3 = same name tokens · 2 = one name contains the other (≥2 tokens) · 1 = 2+ tokens shared · 0
  function nameScore(a, b) {
    var A = {}, B = {}; tokens(a).forEach(function (t) { A[t] = 1; }); tokens(b).forEach(function (t) { B[t] = 1; });
    var ka = Object.keys(A), kb = Object.keys(B);
    if (!ka.length || !kb.length) return 0;
    var inter = ka.filter(function (t) { return B[t]; }).length, small = Math.min(ka.length, kb.length);
    if (inter === ka.length && inter === kb.length) return 3;
    if (inter === small && small >= 2) return 2;
    return inter >= 2 ? 1 : 0;
  }
  /* Same school under different spellings: "St. Example College Westbrook (SECW)" (folder) ~ "St Example College Westbrook"
   * (Certiport Test Center) ~ "SECW". Equal words, one name's words inside the other (2+ words), or an acronym. */
  var SCH_STOP = { of: 1, the: 1, and: 1, de: 1, la: 1, del: 1, inc: 1, school: 0 };
  function schoolWords(s) { return tokens(String(s || "").replace(/\(([^)]*)\)/g, " ")).filter(function (t) { return !SCH_STOP[t]; }); }
  function schoolAcr(s) { var m = String(s || "").match(/\(([A-Za-z0-9]{2,10})\)/); return (m ? m[1] : acronym(s)).toLowerCase(); }
  function schoolMatch(a, b) {
    if (!a || !b) return false;
    if (a === b) return true;
    var A = schoolWords(a), B = schoolWords(b);
    if (!A.length || !B.length) return false;
    var inB = A.filter(function (t) { return B.indexOf(t) >= 0; }).length, inA = B.filter(function (t) { return A.indexOf(t) >= 0; }).length;
    if (inB === A.length && inA === B.length) return true;
    if ((inB === A.length && A.length >= 2) || (inA === B.length && B.length >= 2)) return true;
    var aa = schoolAcr(a), ab = schoolAcr(b);
    return aa.length >= 3 && (aa === ab || (A.length === 1 && A[0] === ab) || (B.length === 1 && B[0] === aa));
  }
  /* Loosely related (for choosing which folders to read, never for linking): a Certiport Test Center that groups
   * several schools ("Sisters of St Example of Riverdale …") ~ its school folders ("St. Example College Westbrook") — they
   * share a distinctive word. */
  var SCH_GENERIC = /^(saint|sainte|santa|santo|school|schools|college|colleges|university|academy|education|educational|ministry|sisters|brothers|catholic|parish|high|integrated|elementary|institute|center|centre|foundation|incorporated|system|learning|group|test|campus|main|city|international|christian|montessori)$/;
  function schoolRelated(a, b) {
    if (schoolMatch(a, b)) return true;
    var A = schoolWords(a).filter(function (t) { return t.length >= 4 && !SCH_GENERIC.test(t); });
    var B = schoolWords(b).filter(function (t) { return t.length >= 4 && !SCH_GENERIC.test(t); });
    return A.some(function (t) { return B.indexOf(t) >= 0; });
  }
  /* "…/PRINTED (FEB 13, 2026)/Exam 1343.2 … CERT.pdf" or "(printed)Exam 1339.1 …" → printed (date when the folder has one) */
  function printedFromPath(path) {
    var parts = String(path || "").replace(/\\/g, "/").split("/");
    for (var i = parts.length - 1; i >= 0; i--) {
      var p = parts[i], last = i === parts.length - 1;
      if (last ? /^\((p|printed)\)/i.test(p) : /^printed\b/i.test(p)) {
        var m = p.match(/\(([^)]*\d{4})\)/), iso = m ? parseDate(m[1].replace(/^([A-Za-z]{3})[a-z]*\.?/, "$1")) : "";
        return { printed: true, date: iso, folder: last ? "" : p };
      }
    }
    return null;
  }
  function famKey(f) { return String(f || "").toLowerCase().replace(/\s+/g, " ").trim(); }
  /* certs: parsed certificates (+ fileMeta); rows: buildModel rows; links: { code: attemptKey | "none" } (manual / from the log).
   * → { certs: [cert + {status: linked|ambiguous|unmatched|duplicate, attemptKey, how, cands:[{key, score, days}]}], byAttempt: {key: code} } */
  function matchCertificates(certs, rows, links) {
    links = links || {};
    var byKey = {}; (rows || []).forEach(function (r) { byKey[r.key] = r; });
    var byRef = {}; (rows || []).forEach(function (r) { if (r.ref) byRef[r.ref] = r.key; });
    var out = [], taken = {};
    (certs || []).forEach(function (c) {
      var o = {}, p; for (p in c) if (Object.prototype.hasOwnProperty.call(c, p)) o[p] = c[p];
      o.cands = []; o.attemptKey = ""; o.how = "";
      (rows || []).forEach(function (r) {
        if (!r.pass || r.pending) return;
        if (famKey(r.exam) !== famKey(c.family)) return;
        if (c.school && c.school !== "Unassigned school" && r.school && !r.schoolFromCenter && !schoolMatch(c.school, r.school)) return;
        if (c.examBase && r.examBase && c.examBase !== r.examBase) return;
        var d = dayDiff(c.date, r.date);
        if (d != null && d > 14) return;
        var sc = Math.max(nameScore(c.name, r.name), r.rosterName ? nameScore(c.name, r.rosterName) : 0);
        if (sc < 1) return;
        o.cands.push({ key: r.key, score: sc, days: d });
      });
      o.cands.sort(function (a, b) { return b.score - a.score || (a.days == null ? 99 : a.days) - (b.days == null ? 99 : b.days); });
      var man = links[c.code];
      if (man && man !== "none") { var mk = byKey[man] ? man : byRef[man]; if (mk) { o.attemptKey = mk; o.how = "manual"; } }
      if (man === "none") o.how = "none";
      if (!o.attemptKey && !o.how) {
        var top = o.cands[0], second = o.cands[1];
        if (top && top.score >= 2 && (top.days == null || top.days <= 7) && !(second && second.score === top.score)) { o.attemptKey = top.key; o.how = "auto"; }
      }
      // "none" = you confirmed there is no exam result for it (e.g. results from another school year not loaded)
      o.status = o.attemptKey ? "linked" : o.how === "none" ? "unlinked" : (o.cands.length ? "ambiguous" : "unmatched");
      out.push(o);
    });
    // one certificate per attempt: a second certificate on the same pass (another code) = possible reissue / duplicate
    var byAttempt = {};
    out.forEach(function (o) {
      if (!o.attemptKey) return;
      if (taken[o.attemptKey] && taken[o.attemptKey] !== o.code) { o.status = "duplicate"; o.dupOf = taken[o.attemptKey]; return; }
      taken[o.attemptKey] = o.code; byAttempt[o.attemptKey] = o.code;
    });
    return { certs: out, byAttempt: byAttempt };
  }
  /* ---- certificate log workbook ("MOS Certificate Log.xlsx", sheet "Certificate Log") ---- */
  var LOG_SHEET = "Certificate Log";
  var LOG_COLS = ["Verification Code", "Name", "Certification", "Version", "Certificate Date", "School", "Exam No.", "Source File", "Page",
    "First Seen", "Linked Exam Ref #", "Printed", "First Printed", "Last Printed", "Printed By", "Print Count", "Remarks"];
  function logKey(h) { return norm(h).toLowerCase(); }
  // aoa (header row first) → { header: [...], col: {name: idx}, rows: { code: {r: rowIdx, v: {colName: value}} }, last }
  function parseLogRows(aoa) {
    aoa = aoa || [];
    var hi = -1, i;
    for (i = 0; i < Math.min(10, aoa.length); i++) if ((aoa[i] || []).some(function (h) { return logKey(h) === "verification code"; })) { hi = i; break; }
    var out = { headerRow: hi < 0 ? 0 : hi, header: hi < 0 ? [] : (aoa[hi] || []).map(norm), col: {}, rows: {}, last: hi < 0 ? 0 : hi, order: [] };
    out.header.forEach(function (h, j) { LOG_COLS.forEach(function (n) { if (logKey(h) === logKey(n) && out.col[n] == null) out.col[n] = j; }); });
    if (hi < 0) return out;
    for (i = hi + 1; i < aoa.length; i++) {
      var row = aoa[i] || [], code = out.col["Verification Code"] != null ? norm(row[out.col["Verification Code"]]) : "";
      if (row.some(function (v) { return v !== "" && v != null; })) out.last = i;
      if (!code || out.rows[code]) continue;
      var v = {}; LOG_COLS.forEach(function (n) { if (out.col[n] != null) v[n] = row[out.col[n]]; });
      out.rows[code] = { r: i, v: v }; out.order.push(code);
    }
    return out;
  }
  function logPrinted(v) { return !!v && /^(y|yes|true|1|printed)$/i.test(norm(v.Printed)); }

  /* ---- filters + summaries ----------------------------------------------- */
  // f: {school, section, q, exam, examNo, outcome('PASS'|'FAIL'|''), min, max, year, month, sy, firstOnly}
  function filterAttempts(rows, f) {
    f = f || {};
    var q = norm(f.q).toLowerCase();
    return rows.filter(function (r) {
      if (f.school && r.school !== f.school) return false;
      if (f.section && (r.section || "(none)") !== f.section) return false;
      if (f.exam && r.exam !== f.exam) return false;
      if (f.examNo && r.examBase !== f.examNo && r.examNo !== f.examNo && (r.batch || "(none)") !== f.examNo) return false;
      if (f.proctor && (r.proctor || "(none)") !== f.proctor) return false;
      if (f.outcome === "RETAKE") { if (!(r.retakeEligible && !r.retakeUsed)) return false; }
      else if (f.outcome && r.outcome !== f.outcome) return false;
      if (f.min != null && f.min !== "" && (r.score == null || r.score < +f.min)) return false;
      if (f.max != null && f.max !== "" && (r.score == null || r.score > +f.max)) return false;
      if (f.year && r.year !== f.year) return false;
      if (f.month && r.month !== f.month) return false;
      if (f.sy && r.sy !== f.sy) return false;
      if (f.firstOnly && r.attemptNo !== 1) return false;
      if (q && (r.name + " " + r.rosterName + " " + r.email + " " + r.certiportId + " " + r.studentNo).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
  }
  // "In Progress" rows (no score yet) are counted in `pending` only — never as a fail or a 0.
  function summarize(all) {
    var rows = all.filter(function (r) { return r.score != null; });
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
      students: Object.keys(students).length,
      pending: all.length - rows.length
    };
  }
  // group → [{key, attempts, passed, passRate}] sorted by key
  function groupRate(rows, keyFn) {
    var g = {};
    rows.forEach(function (r) {
      if (r.score == null) return;
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

  // School × exam performance (MOS › Analytics). Same counting as summarize(): only
  // scored attempts count; "In Progress" rows go to `pending`, never to a fail.
  //   passRate      = passed attempts / scored attempts
  //   firstPassRate = passed FIRST takes / first takes (attemptNo 1)
  //   certRate      = exam takers (student × exam) whose BEST take passed / exam takers
  //   small         = fewer than SMALL_GROUP scored attempts — one or two students decide the rate
  // → { schools, exams, cells[school][exam], bySchool[school], byExam[exam], total }
  var SMALL_GROUP = 5;
  function performance(all) {
    function blank() { return { attempts: 0, passed: 0, sum: 0, first: 0, firstPass: 0, pending: 0, students: {}, takers: {} }; }
    function add(c, r) {
      if (r.score == null) { c.pending++; return; }
      c.attempts++; c.sum += r.score; if (r.pass) c.passed++;
      if (r.attemptNo === 1) { c.first++; if (r.pass) c.firstPass++; }
      c.students[r.studentKey] = 1;
      var t = r.studentKey + "|" + r.exam; c.takers[t] = c.takers[t] || !!r.pass;
    }
    function done(c) {
      var takers = Object.keys(c.takers), cert = takers.filter(function (k) { return c.takers[k]; }).length;
      return {
        attempts: c.attempts, passed: c.passed, failed: c.attempts - c.passed, pending: c.pending,
        passRate: c.attempts ? c.passed / c.attempts : null,
        firstAttempts: c.first, firstPassRate: c.first ? c.firstPass / c.first : null,
        avg: c.attempts ? c.sum / c.attempts : null,
        students: Object.keys(c.students).length, takers: takers.length, certified: cert,
        certRate: takers.length ? cert / takers.length : null,
        small: c.attempts > 0 && c.attempts < SMALL_GROUP
      };
    }
    var cells = {}, bySchool = {}, byExam = {}, total = blank();
    (all || []).forEach(function (r) {
      var s = r.school || "Unassigned school", e = r.exam || "(unknown exam)";
      var row = cells[s] || (cells[s] = {});
      add(row[e] || (row[e] = blank()), r);
      add(bySchool[s] || (bySchool[s] = blank()), r);
      add(byExam[e] || (byExam[e] = blank()), r);
      add(total, r);
    });
    var out = { schools: Object.keys(bySchool).sort(), exams: [], cells: {}, bySchool: {}, byExam: {}, total: done(total) };
    out.schools.forEach(function (s) {
      out.bySchool[s] = done(bySchool[s]); out.cells[s] = {};
      Object.keys(cells[s]).forEach(function (e) { out.cells[s][e] = done(cells[s][e]); });
    });
    Object.keys(byExam).forEach(function (e) { out.byExam[e] = done(byExam[e]); });
    // most-taken exam first, so the matrix reads left to right by volume
    out.exams = Object.keys(byExam).sort(function (a, b) { return out.byExam[b].attempts - out.byExam[a].attempts || a.localeCompare(b); });
    return out;
  }

  var api = {
    VERSION: "1.37.0", SMALL_GROUP: SMALL_GROUP, performance: performance,
    toItems: toItems, toLines: toLines, cleanTitle: cleanTitle, examFamily: examFamily, examVersion: examVersion,
    parseDate: parseDate, schoolYear: schoolYear, parseScorePage: parseScorePage, fileMeta: fileMeta,
    parseRosterRows: parseRosterRows, tokens: tokens, matchRoster: matchRoster, buildModel: buildModel,
    filterAttempts: filterAttempts, parseExportRows: parseExportRows, exportHeader: exportHeader,
    parseAnyDate: parseAnyDate, nameCase: nameCase, acronym: acronym, summarize: summarize, groupRate: groupRate, skillAreas: skillAreas,
    bandOf: bandOf, BANDS: BANDS, displayName: displayName,
    parseCertificatePage: parseCertificatePage, matchCertificates: matchCertificates, nameScore: nameScore, dayDiff: dayDiff,
    parseLogRows: parseLogRows, schoolMatch: schoolMatch, schoolRelated: schoolRelated, printedFromPath: printedFromPath, logPrinted: logPrinted, LOG_COLS: LOG_COLS, LOG_SHEET: LOG_SHEET
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else { root.QBR = root.QBR || {}; root.QBR.mosParse = api; }
})(typeof window !== "undefined" ? window : this);
