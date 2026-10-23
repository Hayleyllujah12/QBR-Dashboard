/* =============================================================================
 * mos-reconcile.js — MOS › Reconcile view (Certification › MOS results)
 * -----------------------------------------------------------------------------
 * Matches a class roster (Pedro's format: one sheet, a Section column and an
 * EXAM VOUCHER column) to a Certiport "Results … Plus" export by EXAM VOUCHER,
 * so scores/results can be copied into the roster and every matched student
 * carries their roster Section into the Results view (the score-report PDFs
 * have no section, and the Results export is the only place vouchers live).
 *
 * Matching rules (Pedro, 2026-10-23 — the standalone reconciliation, verified):
 *  - Tier 1: voucher, exact (case-insensitive). Names are validated against the
 *    voucher pair as token sets (First/Last/Alternate, order-free); a swapped
 *    First/Last still matches and is flagged NAME DIFF.
 *  - Tier 2: name fallback (same exam, token sets) — a CANDIDATE only. It is
 *    filled solely when Pedro accepts it in the preview (NAME MATCH — CONFIRMED);
 *    never silently.
 *  - Certiport wins: score/result overwrite the roster's; every change is
 *    flagged old→new. The roster RESULT cell is a formula (=IF(score>=700,…))
 *    upstream, so comparisons use its computed/displayed value and the export
 *    writes the SCORE; the result is written as a value in the new workbook.
 *  - Attendance never gates matching (proctors forget it).
 *  - The roster is the master list: preview and exports follow the roster's
 *    name order exactly; unmatched Certiport rows go to "Not in Roster".
 *
 * NOTE (deliberate, Pedro's spec 2026-10-23): unlike js/mos-parse.js — which
 * never reads the Voucher / License column — this module DOES read it: it is
 * the match key. Per Pedro's choice the LAST RUN is kept in localStorage
 * (qbr-mos-reconcile-v1) so it can be resumed; "Clear saved run" wipes it.
 * Everything stays in the browser; fully offline.
 *
 * Pure core (parseTSV / parseClassRoster / parseCertExport / reconcile) is
 * DOM-free and exported for Node tests (tests/mos-reconcile.cjs).
 * ========================================================================== */
(function (root) {
  "use strict";

  /* ================= pure core ================= */
  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }
  function tokens(s) {
    var t = String(s == null ? "" : s);
    if (t.normalize) t = t.normalize("NFD").replace(/[̀-ͯ]/g, "");
    return t.toLowerCase().replace(/[^a-z0-9ñ\s-]/g, " ").split(/[\s-]+/).filter(Boolean);
  }
  function examFamilyOf(title) {
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
  function mapResult(r) {
    var s = norm(r).toLowerCase();
    if (/^pass/.test(s)) return "PASSED";
    if (/^fail/.test(s)) return "FAILED";
    return norm(r).toUpperCase();
  }
  function cellNum(v) {
    if (typeof v === "number" && isFinite(v)) return v;
    var s = norm(v);
    return /^\d{1,4}$/.test(s) ? +s : null;
  }

  function parseTSV(text) {
    return String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n")
      .map(function (ln) { return ln.split("\t"); })
      .filter(function (r) { return r.some(function (c) { return norm(c); }); });
  }

  /* ---- class roster (Pedro's format) ----
   * A worksheet may hold ONE header row + a filled Section column, or several
   * sections STACKED as blocks: [section title row] header row, students,
   * [next title] header row, students… A header row starts a new segment with
   * its own column map; a mostly-empty row (<=2 filled cells, no voucher /
   * GMetrix) is a section title; a title between student groups (no repeated
   * header) also starts a new segment reusing the layout. Each segment is its
   * own block, labelled "<source> · <section title>" when a title was seen. */
  var ROSTER_COLS = [
    ["name", /^(display\s*name|student\s*name|full\s*name|name)$/i],
    ["voucher", /voucher/i],
    ["section", /^section$/i],
    ["gmetrix", /gmetrix/i],
    ["score", /^exam\s*score$/i],
    ["result", /^result$/i]
  ];
  function headerCols(row) {
    var cols = {};
    (row || []).forEach(function (v, j) {
      var h = norm(v);
      ROSTER_COLS.forEach(function (pair) { if (cols[pair[0]] == null && pair[1].test(h)) cols[pair[0]] = j; });
      if (cols.score == null && /^score$/i.test(h)) cols.score = j;
    });
    return cols.name != null && cols.voucher != null ? cols : null;
  }
  function detectRosterHeader(aoa) {
    for (var r = 0; r < Math.min(aoa.length, 10); r++) {
      var cols = headerCols(aoa[r]);
      if (cols) return { row: r, cols: cols };
    }
    return null;
  }
  function nonEmptyCount(row) { return (row || []).filter(function (v) { return norm(v); }).length; }
  function longestCell(row) {
    var best = "";
    (row || []).forEach(function (v) { var t = norm(v); if (t.length > best.length) best = t; });
    return best;
  }
  function parseClassRosterBlocks(aoa, source) {
    var blocks = [], cur = null, pendingTitle = "";
    function newSegment(cols, headers, title) {
      cur = { source: source + (title ? " · " + title : ""), headers: headers, headerRow: -1, cols: cols, students: [], title: title || "" };
      blocks.push(cur);
    }
    (aoa || []).forEach(function (row, i) {
      var cols = headerCols(row);
      if (cols) { var inherit = pendingTitle || (cur && !cur.students.length ? cur.title : ""); newSegment(cols, (row || []).map(norm), inherit); cur.headerRow = i; pendingTitle = ""; return; }
      var n = nonEmptyCount(row);
      if (!n) return;
      if (n <= 2) {   // title / label row (section heading, sheet heading) — never a student
        var t = longestCell(row);
        if (t) {
          if (cur && cur.students.length) newSegment(cur.cols, cur.headers, t);   // title mid-sheet: new segment, same layout
          else pendingTitle = t;
        }
        return;
      }
      if (!cur) return;
      var c = cur.cols;
      var name = norm(row[c.name]);
      if (!name) return;
      var voucher = c.voucher != null ? norm(row[c.voucher]).toUpperCase() : "";
      var gmetrix = c.gmetrix != null ? norm(row[c.gmetrix]) : "";
      if (!voucher && !gmetrix) return;   // subtotal / note rows are not students
      cur.students.push({
        name: name,
        section: (c.section != null ? norm(row[c.section]) : "") || cur.title || "",
        voucher: voucher,
        oldScore: c.score != null ? cellNum(row[c.score]) : null,
        oldScoreRaw: c.score != null ? norm(row[c.score]) : "",
        oldResult: c.result != null ? mapResult(row[c.result]) : "",
        raw: row.slice()
      });
    });
    return blocks.filter(function (b) { return b.students.length; });
  }
  function parseClassRoster(aoa, source) {
    var blocks = parseClassRosterBlocks(aoa, source);
    return blocks.length ? blocks[0] : null;
  }

  /* ---- Certiport Results export (voucher-reading variant) ---- */
  var EXPORT_COLS = [
    ["voucher", /voucher/i],
    ["ref", /^exam\s*reference\s*(number|no\.?|#)$/i],
    ["first", /^first\s*name$/i],
    ["last", /^last\s*name$/i],
    ["alt", /^alternate\s*name$/i],
    ["exam", /^exam$/i],
    ["date", /^exam\s*date$/i],
    ["score", /^score$/i],
    ["result", /^result$/i]
  ];
  function detectExportHeader(aoa) {
    for (var r = 0; r < Math.min(aoa.length, 15); r++) {
      var row = aoa[r] || [], cols = {};
      row.forEach(function (v, j) {
        var h = norm(v);
        EXPORT_COLS.forEach(function (pair) { if (cols[pair[0]] == null && pair[1].test(h)) cols[pair[0]] = j; });
      });
      if (cols.voucher != null && cols.ref != null && cols.score != null && cols.result != null) return { row: r, cols: cols };
    }
    return null;
  }
  function parseCertExport(aoa) {
    var h = detectExportHeader(aoa || []);
    if (!h) return null;
    var c = h.cols, out = [], noVoucher = 0;
    function g(row, k) { return c[k] == null ? "" : row[c[k]]; }
    for (var i = h.row + 1; i < aoa.length; i++) {
      var row = aoa[i] || [];
      var ref = norm(g(row, "ref"));
      if (!/^\d{5,}$/.test(ref)) continue;
      var voucher = norm(g(row, "voucher")).toUpperCase();
      if (!voucher) { noVoucher++; continue; }
      var title = norm(g(row, "exam"));
      var d = g(row, "date");
      out.push({
        voucher: voucher, ref: ref,
        first: norm(g(row, "first")), last: norm(g(row, "last")), alt: norm(g(row, "alt")),
        name: norm(norm(g(row, "first")) + " " + norm(g(row, "last"))),
        examTitle: title, exam: examFamilyOf(title),
        date: d instanceof Date && !isNaN(d) ? d.toISOString().slice(0, 10) : norm(d).slice(0, 10),
        score: cellNum(g(row, "score")), result: norm(g(row, "result"))
      });
    }
    return { rows: out, noVoucher: noVoucher, headerRow: h.row };
  }

  /* ---- matching ---- */
  function tokenSet(arr) { var o = {}; arr.forEach(function (t) { o[t] = 1; }); return o; }
  function surnameTokens(rosterName) {
    var p = String(rosterName || "").split(",");
    return tokens(p.length > 1 ? p[0] : rosterName);
  }
  function firstTokens(rosterName) {
    var p = String(rosterName || "").split(",");
    return p.length > 1 ? tokens(p.slice(1).join(",")) : tokens(rosterName).slice(0, -1);
  }
  function certAllTokens(row) { return tokenSet(tokens(row.first).concat(tokens(row.last), tokens(row.alt))); }
  // name validation for a voucher pair: surname tokens present + a first-name token present
  function nameCheck(rosterName, row) {
    var all = certAllTokens(row);
    var sur = surnameTokens(rosterName), fst = firstTokens(rosterName);
    var surInLast = sur.every(function (t) { return tokenSet(tokens(row.last))[t]; });
    var surInFirst = sur.length && sur.every(function (t) { return tokenSet(tokens(row.first))[t]; });
    var fstInLast = fst.length && fst.some(function (t) { return tokenSet(tokens(row.last))[t]; });
    return {
      consistent: sur.every(function (t) { return all[t]; }) && (!fst.length || fst.some(function (t) { return all[t]; })),
      swapped: !!(surInFirst && fstInLast && !surInLast)
    };
  }
  function isCandidate(rosterName, row) {
    var all = certAllTokens(row);
    var rt = tokenSet(tokens(rosterName));
    var ct = tokenSet(Object.keys(all));
    var rKeys = Object.keys(rt), cKeys = Object.keys(ct);
    if (rKeys.length && rKeys.length === cKeys.length && rKeys.every(function (t) { return ct[t]; })) return true;
    var sur = surnameTokens(rosterName), fst = firstTokens(rosterName);
    return sur.length > 0 && sur.every(function (t) { return tokenSet(tokens(row.last).concat(tokens(row.alt)))[t]; })
      && fst.length > 0 && !!tokenSet(tokens(row.first))[fst[0]];
  }
  function diffsOf(st, hit) {
    var diffs = [], conflict = false;
    var newScore = hit.score, newResult = mapResult(hit.result);
    if (st.oldScore == null) { if (newScore != null) diffs.push("score (blank)→" + newScore); }
    else if (newScore != null && st.oldScore !== newScore) { diffs.push("score " + st.oldScore + "→" + newScore); conflict = true; }
    if (st.oldResult !== newResult) {
      diffs.push("result " + (st.oldResult || "(blank)") + "→" + newResult);
      if (st.oldResult) conflict = true;
    }
    return { diffs: diffs, conflict: conflict };
  }
  // blocks: [parseClassRoster(...)], decisions: { "<bi>:<si>": "<ref>" | "rejected" }
  function reconcile(blocks, exportRows, decisions) {
    decisions = decisions || {};
    var byVoucher = {}, dupExports = 0;
    (exportRows || []).forEach(function (r) { if (byVoucher[r.voucher]) dupExports++; else byVoucher[r.voucher] = r; });
    var rows = [], matchedRefs = {}, seenVouchers = {};
    // pass 1 — voucher
    (blocks || []).forEach(function (b, bi) {
      b.students.forEach(function (st, si) {
        var o = { bi: bi, si: si, key: bi + ":" + si, student: st, status: "", hit: null, candidates: [], diffs: [], nameDiff: false, swapped: false, flag: "" };
        rows.push(o);
        if (st.voucher && seenVouchers[st.voucher]) { o.status = "duplicate"; o.flag = "DUPLICATE VOUCHER in roster (first used by roster entry " + seenVouchers[st.voucher] + ")"; return; }
        if (st.voucher) seenVouchers[st.voucher] = si + 1;
        var hit = st.voucher ? byVoucher[st.voucher] : null;
        if (hit) {
          o.hit = hit; matchedRefs[hit.ref] = 1;
          var nc = nameCheck(st.name, hit);
          o.nameDiff = !nc.consistent; o.swapped = nc.swapped;
          var d = diffsOf(st, hit);
          o.diffs = d.diffs;
          o.status = !d.diffs.length ? "verified" : d.conflict ? "changed" : "filled";
        }
      });
    });
    // the roster's exam = majority exam among voucher hits (name fallback stays inside it)
    var examCount = {};
    rows.forEach(function (o) { if (o.hit) examCount[o.hit.exam] = (examCount[o.hit.exam] || 0) + 1; });
    var rosterExam = Object.keys(examCount).sort(function (a, b) { return examCount[b] - examCount[a]; })[0] || "";
    // pass 2 — name candidates + no-result / conflict
    rows.forEach(function (o) {
      if (o.hit || o.status === "duplicate") return;
      var st = o.student;
      var cands = (exportRows || []).filter(function (r) {
        return !matchedRefs[r.ref] && (!rosterExam || r.exam === rosterExam) && isCandidate(st.name, r);
      });
      var dec = decisions[o.key];
      var acc = dec && dec !== "rejected" ? cands.filter(function (r) { return r.ref === dec; })[0] : null;
      if (acc) {
        o.hit = acc; o.status = "confirmed"; matchedRefs[acc.ref] = 1;
        o.diffs = diffsOf(st, acc).diffs;
      } else if (dec === "rejected") {
        o.status = "no-result"; o.rejected = true;
      } else if (cands.length) {
        o.status = "candidate"; o.candidates = cands;
      } else {
        o.status = st.oldScore != null && st.oldResult === "PASSED" ? "conflict" : "no-result";
      }
    });
    rows.forEach(function (o, idx) {
      var st = o.student;
      o.idx = idx;
      o.finalScore = o.hit ? o.hit.score : st.oldScore;
      o.finalResult = o.hit ? mapResult(o.hit.result) : st.oldResult;
      if (o.flag) return;
      if (o.status === "verified") o.flag = "VOUCHER MATCH — VERIFIED (already correct)";
      else if (o.status === "filled") o.flag = "VOUCHER MATCH — FILLED: " + o.diffs.join("; ");
      else if (o.status === "changed") o.flag = "VOUCHER MATCH — CHANGED: " + o.diffs.join("; ");
      else if (o.status === "confirmed") o.flag = "NAME MATCH — CONFIRMED" + (o.diffs.length ? ": " + o.diffs.join("; ") : "");
      else if (o.status === "candidate") o.flag = "NAME MATCH — CONFIRM (" + o.candidates.length + " candidate" + (o.candidates.length === 1 ? "" : "s") + ")";
      else if (o.status === "conflict") o.flag = "CONFLICT — roster shows " + st.oldScore + " / PASSED but Certiport has no record — verify manually";
      else if (o.status === "no-result") o.flag = "NO CERTIPORT RESULT";
      if (o.rejected) o.flag = "NO CERTIPORT RESULT (name candidate rejected)";
      if (o.hit && o.swapped) o.flag += " · NAME DIFF: first/last appear swapped in Certiport";
      else if (o.hit && o.nameDiff) o.flag += " · NAME DIFF: Certiport name differs from roster — check";
    });
    var notInRoster = (exportRows || []).filter(function (r) { return !matchedRefs[r.ref]; });
    var counts = { total: rows.length, verified: 0, filled: 0, changed: 0, confirmed: 0, candidate: 0, noResult: 0, conflict: 0, duplicate: 0, notInRoster: notInRoster.length, exports: (exportRows || []).length, dupExports: dupExports };
    rows.forEach(function (o) {
      if (o.status === "verified") counts.verified++;
      else if (o.status === "filled") counts.filled++;
      else if (o.status === "changed") counts.changed++;
      else if (o.status === "confirmed") counts.confirmed++;
      else if (o.status === "candidate") counts.candidate++;
      else if (o.status === "no-result") counts.noResult++;
      else if (o.status === "conflict") counts.conflict++;
      else if (o.status === "duplicate") counts.duplicate++;
    });
    return { rows: rows, notInRoster: notInRoster, counts: counts, rosterExam: rosterExam };
  }

  // Preview-only filtering/sorting (Results-view parity). The workbook export always
  // covers the whole run in roster order — this only narrows what is on screen.
  function filterPreview(res, f) {
    f = f || {};
    var rows = res.rows.filter(function (o) {
      if (f.section && (o.student.section || "(none)") !== f.section) return false;
      if (f.exam) {
        var ex = o.hit ? o.hit.exam : o.candidates.length ? o.candidates[0].exam : res.rosterExam || "";
        if (ex !== f.exam) return false;
      }
      if (f.flag && o.status !== f.flag) return false;
      if (f.q && (o.student.name + " " + o.student.voucher).toLowerCase().indexOf(String(f.q).toLowerCase()) < 0) return false;
      var sc = o.hit ? o.hit.score : null;
      if (f.min !== "" && f.min != null && (sc == null || sc < +f.min)) return false;
      if (f.max !== "" && f.max != null && (sc == null || sc > +f.max)) return false;
      return true;
    });
    var nm = function (o) { return o.student.name.toLowerCase(); };
    var scOf = function (o, dflt) { return o.hit && o.hit.score != null ? o.hit.score : dflt; };
    if (f.sort === "name") rows.sort(function (a, b) { return nm(a).localeCompare(nm(b)); });
    else if (f.sort === "score-desc") rows.sort(function (a, b) { return scOf(b, -1) - scOf(a, -1) || nm(a).localeCompare(nm(b)); });
    else if (f.sort === "score-asc") rows.sort(function (a, b) { return scOf(a, 9999) - scOf(b, 9999) || nm(a).localeCompare(nm(b)); });
    else if (f.sort === "section") rows.sort(function (a, b) { return (a.student.section || "~").localeCompare(b.student.section || "~") || nm(a).localeCompare(nm(b)); });
    return rows;
  }

  var core = {
    norm: norm, tokens: tokens, examFamilyOf: examFamilyOf, mapResult: mapResult,
    parseTSV: parseTSV, detectRosterHeader: detectRosterHeader, parseClassRoster: parseClassRoster, parseClassRosterBlocks: parseClassRosterBlocks,
    detectExportHeader: detectExportHeader, parseCertExport: parseCertExport,
    nameCheck: nameCheck, isCandidate: isCandidate, reconcile: reconcile, filterPreview: filterPreview
  };
  if (typeof module !== "undefined" && module.exports) module.exports = core;
  if (typeof window === "undefined" || !window.document) return;
  root.QBR = root.QBR || {};
  root.QBR.mosReconcileCore = core;

  /* ================= browser UI ================= */
  var QBR = root.QBR;
  var LS_KEY = "qbr-mos-reconcile-v1";
  var FDEF = { section: "", exam: "", flag: "", q: "", min: "", max: "", sort: "roster" };
  var RC = { blocks: [], exportRows: [], exportName: "", decisions: {}, draft: "", msg: "", appliedRefs: [], F: { section: "", exam: "", flag: "", q: "", min: "", max: "", sort: "roster" }, _applied: false };

  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function $(id) { return document.getElementById(id); }

  function persist() {
    try {
      if (!RC.blocks.length && !RC.exportRows.length) { localStorage.removeItem(LS_KEY); return; }
      localStorage.setItem(LS_KEY, JSON.stringify({
        blocks: RC.blocks, exportRows: RC.exportRows, exportName: RC.exportName,
        decisions: RC.decisions, appliedRefs: RC.appliedRefs, savedAt: new Date().toISOString()
      }));
    } catch (e) { /* storage full/blocked: the run still works for this session */ }
  }
  function restore() {
    try {
      var raw = localStorage.getItem(LS_KEY); if (!raw) return false;
      var d = JSON.parse(raw); if (!d) return false;
      RC.blocks = d.blocks || []; RC.exportRows = d.exportRows || []; RC.exportName = d.exportName || "";
      RC.decisions = d.decisions || {}; RC.appliedRefs = d.appliedRefs || [];
      return !!(RC.blocks.length || RC.exportRows.length);
    } catch (e) { return false; }
  }
  var restored = restore();

  function result() { return reconcile(RC.blocks, RC.exportRows, RC.decisions); }

  /* Feed sections into the MOS model: roster entries (name join covers PDF-only
   * attempts) + per-ref overrides for voucher/confirmed matches (exact). */
  function applyToMos() {
    RC._applied = true;
    var MS = QBR.mos && QBR.mos._state ? QBR.mos._state() : null;
    if (!MS) return;
    MS.roster = (MS.roster || []).filter(function (st) { return !st._rec; });
    RC.appliedRefs.forEach(function (ref) { delete MS.overrides[ref]; });
    RC.appliedRefs = [];
    var res = result();
    RC.blocks.forEach(function (b) {
      b.students.forEach(function (st) {
        MS.roster.push({ school: "", section: st.section, name: st.name, email: "", studentNo: "", _rec: true });
      });
    });
    res.rows.forEach(function (o) {
      if (!o.hit) return;
      MS.overrides[o.hit.ref] = {
        section: o.student.section,
        student: { school: "", section: o.student.section, name: o.student.name, email: "", studentNo: "" }
      };
      RC.appliedRefs.push(o.hit.ref);
    });
    persist();
    if (QBR.mos._rebuild) QBR.mos._rebuild();
  }

  async function readWorkbookAOAs(file) {
    var wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: "array", cellDates: true });
    return wb.SheetNames.map(function (sn) {
      return { name: sn, aoa: XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: true }) };
    });
  }
  async function addRosterFile(file) {
    var sheets = await readWorkbookAOAs(file);
    var added = 0;
    sheets.forEach(function (sh) {
      parseClassRosterBlocks(sh.aoa, file.name + " · " + sh.name).forEach(function (b) { RC.blocks.push(b); added += b.students.length; });
    });
    RC.msg = added ? "" : "No class-roster sheet found in " + file.name + " (needs a name column and an EXAM VOUCHER column).";
    return added;
  }
  async function addExportFile(file) {
    var sheets = await readWorkbookAOAs(file);
    var rows = [], found = false;
    sheets.forEach(function (sh) {
      var ex = parseCertExport(sh.aoa);
      if (ex) { found = true; rows = rows.concat(ex.rows); }
    });
    if (!found) { RC.msg = "No Certiport Results sheet found in " + file.name + " (needs Exam Reference Number, Score, Result and Voucher / License columns)."; return 0; }
    RC.exportRows = rows; RC.exportName = file.name; RC.msg = "";
    // the same file also loads as attempts in the Results view (mos loader dedupes by file id)
    if (QBR.mos && QBR.mos.addRoster) { try { await QBR.mos.addRoster([file]); } catch (e) {} }
    return rows.length;
  }

  function rowExam(o, res) { return o.hit ? o.hit.exam : o.candidates.length ? o.candidates[0].exam : (res ? res.rosterExam : "") || ""; }
  function chip(label, n, tone) {
    return '<span class="mos-rec-chip mos-rec-' + tone + '">' + esc(label) + " <b>" + n + "</b></span>";
  }
  var TONES = { verified: "ok", filled: "ok", changed: "warn", confirmed: "ok", candidate: "warn", "no-result": "mut", conflict: "bad", duplicate: "bad" };

  function render(host) {
    if (!host) return;
    var res = result(), c = res.counts;
    var html =
      '<div class="card-box mos-rec">' +
      '<h6 class="mb-1">Reconcile a class roster with the Certiport export</h6>' +
      '<div class="small text-muted mb-2">Paste or upload the class roster (name order is the master list), upload the Certiport <b>Results … Plus</b> export (.xlsx), and matches are made by <b>exam voucher</b>. ' +
      'Certiport wins on conflicts — every change is flagged. Matched students also carry their <b>Section</b> into the Results view, where the Section filter, Excel export and Exam Summary PDF use it.</div>' +
      '<div class="small mos-rec-note">🔒 The last run is saved <b>in this browser only</b> so you can resume it — nothing is uploaded. <button type="button" class="btn btn-link btn-sm p-0" data-rec="clear">Clear saved run</button></div>' +
      (RC.msg ? '<div class="small text-danger mt-1" role="alert">' + esc(RC.msg) + "</div>" : "") +
      '<div class="mos-rec-grid">' +
        '<div><div class="small fw-bold mb-1">1 · Class roster</div>' +
          '<textarea id="mos-rec-paste" class="form-control form-control-sm" rows="4" placeholder="Copy the roster rows in Excel (including the header row) and paste here…">' + esc(RC.draft) + "</textarea>" +
          '<div class="d-flex gap-2 mt-1 flex-wrap">' +
            '<button type="button" class="btn btn-primary btn-sm" data-rec="paste">Load pasted roster</button>' +
            '<label class="btn btn-outline-secondary btn-sm mb-0">Upload roster (.xlsx)…<input id="mos-rec-roster-file" type="file" accept=".xlsx,.xlsm,.xls" hidden></label>' +
          "</div>" +
          (RC.blocks.length ? '<div class="small mt-1">' + RC.blocks.map(function (b, i) {
            var secs = {}; b.students.forEach(function (s) { if (s.section) secs[s.section] = 1; });
            return esc(b.source) + " — " + b.students.length + " students" + (Object.keys(secs).length ? " · section " + esc(Object.keys(secs).join(", ")) : "") +
              ' <button type="button" class="btn btn-link btn-sm p-0" data-rec="rmblock" data-i="' + i + '">remove</button>';
          }).join("<br>") + "</div>" : "") +
        "</div>" +
        '<div><div class="small fw-bold mb-1">2 · Certiport export</div>' +
          '<label class="btn btn-outline-primary btn-sm mb-0">Upload Certiport export (.xlsx)…<input id="mos-rec-export-file" type="file" accept=".xlsx,.xlsm,.xls" hidden></label>' +
          '<div class="small mt-1">' + (RC.exportRows.length
            ? esc(RC.exportName) + " — " + RC.exportRows.length + " results with vouchers" + (res.rosterExam ? " · roster exam: " + esc(res.rosterExam) : "")
            : "No export loaded yet. (The score-report PDFs have no vouchers — the export is the match source.)") + "</div>" +
        "</div>" +
      "</div>";
    var visible = res.rows, shownCount = res.rows.length, anyFilter = false;
    if (RC.blocks.length) {
      var seenEx = {}, seenSec = {}, examOpts = [], secOpts = [];
      res.rows.forEach(function (o) {
        var ex = rowExam(o, res); if (ex && !seenEx[ex]) { seenEx[ex] = 1; examOpts.push(ex); }
        var sv = o.student.section || "(none)"; if (!seenSec[sv]) { seenSec[sv] = 1; secOpts.push(sv); }
      });
      examOpts.sort(); secOpts.sort();
      if (RC.F.exam && !seenEx[RC.F.exam]) RC.F.exam = "";
      if (RC.F.section && !seenSec[RC.F.section]) RC.F.section = "";
      visible = filterPreview(res, RC.F);
      shownCount = visible.length;
      anyFilter = !!(RC.F.section || RC.F.exam || RC.F.flag || RC.F.q || RC.F.min !== "" || RC.F.max !== "" || RC.F.sort !== "roster");
      var fopts = [["", "All flags"], ["verified", "Verified"], ["filled", "Filled"], ["changed", "Changed"], ["confirmed", "Name-confirmed"], ["candidate", "To confirm"], ["no-result", "No result"], ["conflict", "Conflict"], ["duplicate", "Duplicate voucher"]];
      var sopts = [["roster", "Roster order"], ["name", "Name A-Z"], ["score-desc", "Score: highest first"], ["score-asc", "Score: lowest first"], ["section", "Section"]];
      html += '<div class="mos-rec-chips mt-2">' +
        chip("Roster", c.total, "mut") + chip("Verified", c.verified, "ok") + chip("Filled", c.filled, "ok") +
        chip("Changed", c.changed, "warn") + chip("Name-confirmed", c.confirmed, "ok") + chip("To confirm", c.candidate, "warn") +
        chip("No result", c.noResult, "mut") + chip("Conflict", c.conflict, "bad") + chip("Not in roster", c.notInRoster, "mut") + "</div>" +
        '<div class="d-flex gap-2 mt-2 align-items-center"><button type="button" class="btn btn-primary btn-sm" data-rec="export" ' + (RC.exportRows.length ? "" : "disabled") + '>Export filled roster (.xlsx)</button>' +
        (anyFilter ? '<span class="small text-muted">Showing ' + shownCount + " of " + c.total + " students — filters narrow the preview only; the export covers the whole run</span>" : "") +
        '<span class="small text-muted align-self-center">Filled Roster · Paste Sheet (roster order) · Not in Roster</span></div>' +
        '<div class="mos-rec-fbar">' +
          '<label class="mos-fl"><span>Section</span><select id="mos-rec-f-section" class="form-select form-select-sm"><option value="">All sections</option>' + secOpts.map(function (x) { return '<option value="' + esc(x) + '"' + (RC.F.section === x ? " selected" : "") + ">" + esc(x) + "</option>"; }).join("") + "</select></label>" +
          '<label class="mos-fl"><span>Exam</span><select id="mos-rec-f-exam" class="form-select form-select-sm"><option value="">All exams</option>' + examOpts.map(function (x) { return '<option value="' + esc(x) + '"' + (RC.F.exam === x ? " selected" : "") + ">" + esc(x) + "</option>"; }).join("") + "</select></label>" +
          '<label class="mos-fl"><span>Flag</span><select id="mos-rec-f-flag" class="form-select form-select-sm">' + fopts.map(function (x) { return '<option value="' + x[0] + '"' + (RC.F.flag === x[0] ? " selected" : "") + ">" + x[1] + "</option>"; }).join("") + "</select></label>" +
          '<label class="mos-fl"><span>Certiport score</span><span class="mos-range"><input id="mos-rec-f-min" type="number" min="0" max="1000" step="1" class="form-control form-control-sm" placeholder="0" value="' + esc(RC.F.min) + '" aria-label="Lowest Certiport score"> – <input id="mos-rec-f-max" type="number" min="0" max="1000" step="1" class="form-control form-control-sm" placeholder="1000" value="' + esc(RC.F.max) + '" aria-label="Highest Certiport score"></span></label>' +
          '<label class="mos-fl mos-fl-q"><span>Student</span><input id="mos-rec-q" type="search" class="form-control form-control-sm" placeholder="Name or voucher" autocomplete="off" value="' + esc(RC.F.q) + '"></label>' +
          '<label class="mos-fl"><span>Sort</span><select id="mos-rec-f-sort" class="form-select form-select-sm">' + sopts.map(function (x) { return '<option value="' + x[0] + '"' + (RC.F.sort === x[0] ? " selected" : "") + ">" + x[1] + "</option>"; }).join("") + "</select></label>" +
          '<button id="mos-rec-f-reset" class="btn btn-outline-secondary btn-sm" type="button">Reset filters</button>' +
        "</div>" +
        '<div class="table-responsive mt-2"><table id="mos-rec-tbl" class="table table-sm mos-tbl" data-resizable="static"><thead><tr>' +
        "<th>#</th><th>Student (roster order)</th><th>Section</th><th>Voucher</th><th>Exam</th><th>Roster score / result</th><th>Certiport score / result</th><th>Flag</th>" +
        "</tr></thead><tbody>" +
        visible.map(function (o) {
          var st = o.student;
          var certCell = o.hit ? o.hit.score + " / " + esc(mapResult(o.hit.result)) + '<div class="small text-muted">' + esc(o.hit.exam) + " · " + esc(o.hit.date) + " · ref " + esc(o.hit.ref) + "</div>"
            : o.status === "candidate" ? o.candidates.map(function (cd) {
                return '<div>' + cd.score + " / " + esc(mapResult(cd.result)) + ' <span class="text-muted">(' + esc(cd.name) + " · " + esc(cd.date) + ")</span> " +
                  '<button type="button" class="btn btn-success btn-sm py-0" data-rec="accept" data-k="' + o.key + '" data-ref="' + esc(cd.ref) + '">Accept</button> ' +
                  '<button type="button" class="btn btn-outline-secondary btn-sm py-0" data-rec="reject" data-k="' + o.key + '">Reject</button></div>';
              }).join("") : '<span class="text-muted">—</span>';
          var examCell = o.hit ? esc(o.hit.exam) : o.candidates.length ? esc(o.candidates[0].exam) : "—";
          return "<tr><td>" + (o.idx + 1) + "</td><td>" + esc(st.name) + "</td><td>" + esc(st.section || "—") + "</td><td>" + esc(st.voucher || "—") + "</td><td>" + examCell + "</td>" +
            "<td>" + (st.oldScore == null ? "—" : st.oldScore) + " / " + esc(st.oldResult || "—") + "</td><td>" + certCell + "</td>" +
            '<td><span class="mos-rec-chip mos-rec-' + (TONES[o.status] || "mut") + '">' + esc(o.flag) + "</span></td></tr>";
        }).join("") + (visible.length ? "" : '<tr><td colspan="8" class="text-muted small">No students match these filters.</td></tr>') + "</tbody></table></div>";
      if (res.notInRoster.length) {
        html += '<details class="mt-2"><summary class="small"><b>Not in roster</b> (' + res.notInRoster.length + ") — in the export, matched to nobody</summary>" +
          '<div class="table-responsive"><table class="table table-sm" data-resizable="static"><thead><tr><th>Name</th><th>Exam</th><th>Date</th><th>Score</th><th>Result</th><th>Voucher</th><th>Ref</th></tr></thead><tbody>' +
          res.notInRoster.map(function (r) {
            return "<tr><td>" + esc(r.name) + "</td><td>" + esc(r.exam) + "</td><td>" + esc(r.date) + "</td><td>" + (r.score == null ? "—" : r.score) + "</td><td>" + esc(mapResult(r.result)) + "</td><td>" + esc(r.voucher) + "</td><td>" + esc(r.ref) + "</td></tr>";
          }).join("") + "</tbody></table></div></details>";
      }
    }
    html += "</div>";
    host.innerHTML = html;
    wire(host);
  }

  function wire(host) {
    function refilter(focusId) {
      render(host);
      if (focusId) { var el = $(focusId); if (el) { el.focus(); if (el.setSelectionRange && (el.type === "search" || el.type === "text")) { var v = el.value.length; try { el.setSelectionRange(v, v); } catch (e) {} } } }
    }
    function bindF(id, key, ev) { var el = $(id); if (el) el.addEventListener(ev || "change", function () { RC.F[key] = el.value; refilter(id); }); }
    bindF("mos-rec-f-section", "section"); bindF("mos-rec-f-exam", "exam"); bindF("mos-rec-f-flag", "flag"); bindF("mos-rec-f-sort", "sort");
    bindF("mos-rec-q", "q", "input"); bindF("mos-rec-f-min", "min", "input"); bindF("mos-rec-f-max", "max", "input");
    var fr = $("mos-rec-f-reset");
    if (fr) fr.addEventListener("click", function () { RC.F = { section: FDEF.section, exam: FDEF.exam, flag: FDEF.flag, q: FDEF.q, min: FDEF.min, max: FDEF.max, sort: FDEF.sort }; render(host); });
    var ta = $("mos-rec-paste");
    if (ta) ta.addEventListener("input", function () { RC.draft = ta.value; });
    var rf = $("mos-rec-roster-file");
    if (rf) rf.addEventListener("change", function (e) {
      var f = e.target.files[0]; e.target.value = "";
      if (f) addRosterFile(f).then(function () { persist(); applyToMos(); render(host); });
    });
    var ef = $("mos-rec-export-file");
    if (ef) ef.addEventListener("change", function (e) {
      var f = e.target.files[0]; e.target.value = "";
      if (f) addExportFile(f).then(function () { persist(); applyToMos(); render(host); });
    });
    host.querySelectorAll("[data-rec]").forEach(function (b) {
      b.addEventListener("click", function () {
        var a = b.getAttribute("data-rec");
        if (a === "paste") {
          var aoa = parseTSV(RC.draft);
          var blks = parseClassRosterBlocks(aoa, "Pasted roster " + (RC.blocks.length + 1));
          if (blks.length) { blks.forEach(function (b) { RC.blocks.push(b); }); RC.draft = ""; RC.msg = ""; }
          else RC.msg = "That paste doesn't look like a class roster — include the header row (it needs a name column and an EXAM VOUCHER column).";
          persist(); applyToMos(); render(host);
        } else if (a === "rmblock") {
          RC.blocks.splice(+b.getAttribute("data-i"), 1); persist(); applyToMos(); render(host);
        } else if (a === "accept") {
          RC.decisions[b.getAttribute("data-k")] = b.getAttribute("data-ref"); persist(); applyToMos(); render(host);
        } else if (a === "reject") {
          RC.decisions[b.getAttribute("data-k")] = "rejected"; persist(); applyToMos(); render(host);
        } else if (a === "clear") {
          if (window.confirm("Clear the saved reconcile run (rosters, export and confirmations) from this browser?")) { reset(); render(host); }
        } else if (a === "export") exportWorkbook();
      });
    });
  }

  function exportWorkbook() {
    if (typeof XLSX === "undefined") return;
    var res = result();
    var wb = XLSX.utils.book_new();
    // Filled Roster: one sheet per distinct header layout (the common case is one)
    var groups = [], byHdr = {};
    RC.blocks.forEach(function (b, bi) {
      var k = JSON.stringify(b.headers);
      if (!byHdr[k]) { byHdr[k] = { block: b, bis: [] }; groups.push(byHdr[k]); }
      byHdr[k].bis.push(bi);
    });
    groups.forEach(function (g, gi) {
      var b = g.block, aoa = [b.headers.concat(["MATCH FLAG"])];
      g.bis.forEach(function (bi) {
        res.rows.filter(function (o) { return o.bi === bi; }).forEach(function (o) {
          var row = o.student.raw.slice();
          while (row.length < b.headers.length) row.push("");
          if (b.cols.score != null) { var sn = cellNum(row[b.cols.score]); if (sn != null) row[b.cols.score] = sn; }   // pasted rosters arrive as text — keep scores numeric
          if (o.hit) {
            if (b.cols.score != null) row[b.cols.score] = o.finalScore;
            if (b.cols.result != null) row[b.cols.result] = o.finalResult;
          }
          row.push(o.flag);
          aoa.push(row);
        });
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), gi === 0 ? "Filled Roster" : "Filled Roster " + (gi + 1));
    });
    var paste = [["Student", "Section", "Score", "Result", "Voucher", "Exam", "Exam date"]];
    res.rows.forEach(function (o) {
      paste.push([o.student.name, o.student.section, o.hit ? o.finalScore : "", o.hit ? o.finalResult : "",
        o.hit ? o.hit.voucher : o.student.voucher, o.hit ? o.hit.exam : "", o.hit ? o.hit.date : ""]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(paste), "Paste Sheet (Roster Order)");
    var nir = [["First Name", "Last Name", "Exam", "Date", "Score", "Result", "Voucher", "Exam Reference"]];
    res.notInRoster.forEach(function (r) { nir.push([r.first, r.last, r.exam, r.date, r.score == null ? "" : r.score, mapResult(r.result), r.voucher, r.ref]); });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(nir), "Not in Roster");
    XLSX.writeFile(wb, "MOS_Reconciled_Roster_" + new Date().toISOString().slice(0, 10) + ".xlsx");
  }

  // A run restored from localStorage must be re-applied to the MOS model once per page
  // load — otherwise Results shows no sections until the user touches this view again.
  function applyRestored() {
    if (RC._applied) return;
    RC._applied = true;
    if (RC.blocks.length || RC.exportRows.length) applyToMos();
  }
  function reset() {
    RC._applied = false;
    var MS = QBR.mos && QBR.mos._state ? QBR.mos._state() : null;
    if (MS) {
      MS.roster = (MS.roster || []).filter(function (st) { return !st._rec; });
      RC.appliedRefs.forEach(function (ref) { delete MS.overrides[ref]; });
    }
    RC.blocks = []; RC.exportRows = []; RC.exportName = ""; RC.decisions = {}; RC.draft = ""; RC.msg = ""; RC.appliedRefs = [];
    persist();
    if (QBR.mos && QBR.mos._rebuild) QBR.mos._rebuild();
  }

  QBR.mosReconcile = {
    render: render, applyRestored: applyRestored,
    hasSaved: function () { return !!(RC.blocks.length || RC.exportRows.length); },
    count: function () { return RC.blocks.reduce(function (n, b) { return n + b.students.length; }, 0); },
    reset: reset,
    _state: RC, _result: result, _restored: restored
  };
})(typeof window !== "undefined" ? window : this);
