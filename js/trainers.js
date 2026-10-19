/* =============================================================================
 * trainers.js — Trainers tab: trainer evaluations from MS Forms exports.
 * ----------------------------------------------------------------------------
 * Sessions come from the training tracker workbook (ETG_Training_Tracker style):
 * pick a session from the log, upload its MS Forms eval export (one file = one
 * session), generate results, write back to the tracker.
 *
 * Scores per session:
 *   - Session NPS (standard: % Promoters 9-10 minus % Detractors 0-6)
 *   - Trainer effectiveness = mean of the 3 facilitator items
 *     (knowledge/mastery, communication/presentation, ability to engage)
 * The eval file carries NO trainer column; the trainer is inherited from the
 * tracker session the user confirms. The eval filename (date/AM-PM/school)
 * only *suggests* candidate sessions.
 *
 * Write-back: NPS -> tracker's "Net Promoter Score" column, trainer score ->
 * new "Trainer Score" column. Preview-then-confirm; download updated workbook.
 * Everything offline; parsed results in memory, column maps in localStorage.
 * ========================================================================== */
(function () {
  "use strict";
  var QBR = (window.QBR = window.QBR || {});
  var P = function () { return QBR.trainersParse; };

  if (QBR.features && QBR.features.register)
    QBR.features.register({ module: "training", key: "trainers", label: "Trainers",
      desc: "Trainer evaluations: NPS + trainer score from MS Forms exports" });

  var MAP_KEY = "qbr-trainers-maps";
  var LOW_N = 10;

  var S = {
    built: false,
    tracker: null,   // { name, wb, sheetName, headerRow, sessions[], roster[] }
    eval: null,      // { name, rows, header, roles, sig, attachedTo, suggested[] }
    selected: -1,    // index into sessions
    filter: "all",   // all | needs | scored
    q: "",
    view: "sessions",// sessions | board
    maps: {}
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function on() { try { return typeof QBR.adminIsOn !== "function" || QBR.adminIsOn("trainers"); } catch (e) { return true; } }
  function num(v, d) { return v == null || !isFinite(v) ? "—" : Number(v).toFixed(d == null ? 1 : d); }

  function loadMaps() {
    try { S.maps = JSON.parse(localStorage.getItem(MAP_KEY) || "{}") || {}; }
    catch (e) { S.maps = {}; }
  }
  function saveMaps() {
    try { localStorage.setItem(MAP_KEY, JSON.stringify(S.maps)); } catch (e) {}
  }

  /* ---------------- file reading ---------------- */
  function readWorkbook(file) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () {
        try { res(XLSX.read(fr.result, { type: "array" })); }
        catch (e) { rej(e); }
      };
      fr.onerror = function () { rej(new Error("Could not read " + file.name)); };
      fr.readAsArrayBuffer(file);
    });
  }
  function sheetRows(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  }

  /* ---------------- tracker ---------------- */
  function openTracker(file) {
    return readWorkbook(file).then(function (wb) {
      var names = wb.SheetNames || [];
      // Prefer a sheet that parses as a session log; default to "trial".
      var pick = names.indexOf("trial") >= 0 ? "trial" : names[0];
      var sessions = P().parseTrackerSheet(sheetRows(wb, pick));
      if (!sessions.length) {
        // Try every other sheet before giving up.
        for (var i = 0; i < names.length && !sessions.length; i++) {
          if (names[i] === pick) continue;
          sessions = P().parseTrackerSheet(sheetRows(wb, names[i]));
          if (sessions.length) pick = names[i];
        }
      }
      if (!sessions.length) throw new Error("No session log found in " + file.name);
      S.tracker = {
        name: file.name, wb: wb, sheetName: pick,
        headerRow: sheetRows(wb, pick)[0] || [],
        sessions: sessions, roster: P().buildRoster(sessions)
      };
      S.selected = -1; S.eval = null;
      render();
    });
  }

  /* ---------------- eval file ---------------- */
  function openEval(file) {
    return readWorkbook(file).then(function (wb) {
      var rows = sheetRows(wb, wb.SheetNames[0]);
      if (rows.length < 2) throw new Error("No responses in " + file.name);
      var det = P().detectLayout(rows[0]);
      var roles = (S.maps[det.sig] && S.maps[det.sig].roles) || det.roles;
      var hint = P().parseFilename(file.name);
      var suggested = hint && S.tracker ? P().suggestSessions(S.tracker.sessions, hint) : [];
      S.eval = {
        name: file.name, rows: rows, header: rows[0].map(P().norm),
        roles: roles, sig: det.sig, auto: det.roles,
        attachedTo: -1, suggested: suggested, hint: hint, needsMap: !S.maps[det.sig]
      };
      // If a session is already selected, attach; otherwise the user picks
      // from suggestions (or the session list).
      if (S.selected >= 0) S.eval.attachedTo = S.selected;
      render();
    });
  }

  /* ---------------- scoring ---------------- */
  function generateFor(idx) {
    var t = S.tracker, ev = S.eval;
    if (!t || !ev || ev.attachedTo !== idx) return null;
    var roles = ev.roles;
    if (roles.nps == null || roles.nps < 0) return { error: "Map the NPS (0–10) column first." };
    var data = ev.rows.slice(1);
    var scores = data.map(function (r) { return Number(r[roles.nps]); });
    var nps = P().calcNPS(scores);
    var teff = P().calcTrainerScore(data, roles);
    var session = t.sessions[idx];
    var prev = P().findPrevSession(t.sessions, session);
    var computed = {
      n: nps.n, promoters: nps.promoters, passives: nps.passives, detractors: nps.detractors,
      nps: nps.nps, trainerScore: teff.score, itemAvgs: teff.itemAvgs,
      prevNps: prev ? prev.nps : null, prevScore: prev ? prev.trainerScore : null,
      prevLabel: prev ? (prev.dateISO || "") + " · " + prev.client : ""
    };
    session.computed = computed;
    session.status = "ready";
    return computed;
  }

  /* ---------------- write-back ---------------- */
  function writeback() {
    var t = S.tracker;
    if (!t) return;
    var plan = P().planWriteback(t.headerRow, t.sessions);
    if (!plan.edits.length) return;
    var ws = t.wb.Sheets[t.sheetName];
    var npsCol = plan.npsCol, tsCol = plan.trainerCol;
    if (plan.addTrainerCol) {
      tsCol = t.headerRow.length;
      t.headerRow.push("Trainer Score");
      XLSX.utils.sheet_add_aoa(ws, [["Trainer Score"]], { origin: { r: 0, c: tsCol } });
    }
    plan.edits.forEach(function (e) {
      // rowIdx is 0-based with header at 0 -> Excel row = rowIdx + 1
      if (npsCol >= 0 && e.nps != null)
        XLSX.utils.sheet_add_aoa(ws, [[e.nps]], { origin: { r: e.row, c: npsCol } });
      if (tsCol >= 0 && e.trainerScore != null)
        XLSX.utils.sheet_add_aoa(ws, [[e.trainerScore]], { origin: { r: e.row, c: tsCol } });
    });
    t.sessions.forEach(function (s) {
      if (s.computed) {
        if (s.computed.nps != null) s.nps = s.computed.nps;
        if (s.computed.trainerScore != null) s.trainerScore = s.computed.trainerScore;
        s.status = "scored";
        delete s.computed;
      }
    });
    var out = XLSX.write(t.wb, { bookType: "xlsx", type: "array" });
    var blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = t.name.replace(/\.xlsx?$/i, "") + "-scored.xlsx";
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    render();
  }

  /* ---------------- results export ---------------- */
  function exportResults() {
    var t = S.tracker;
    if (!t) return;
    var rows = [["Date", "Trainer", "Course code", "Session title", "School", "Responses",
                 "NPS", "Trainer score", "Promoters", "Passives", "Detractors"]];
    t.sessions.forEach(function (s) {
      var c = s.computed || {};
      var nps = c.nps != null ? c.nps : s.nps, tsc = c.trainerScore != null ? c.trainerScore : s.trainerScore;
      if (nps == null && tsc == null) return;
      rows.push([s.dateISO, s.trainer, s.code, s.title, s.client, c.n || "",
                 nps, tsc, c.promoters != null ? c.promoters : "",
                 c.passives != null ? c.passives : "", c.detractors != null ? c.detractors : ""]);
    });
    if (rows.length < 2) return;
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Trainer results");
    var out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    var blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = "trainer-results.xlsx";
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  /* ---------------- leaderboard ---------------- */
  function leaderboard() {
    var t = S.tracker;
    if (!t) return [];
    var agg = {};
    t.sessions.forEach(function (s) {
      var c = s.computed || {};
      var nps = c.nps != null ? c.nps : s.nps, tsc = c.trainerScore != null ? c.trainerScore : s.trainerScore;
      if (nps == null && tsc == null) return;
      var k = P().normKey(s.trainer);
      var a = agg[k] || (agg[k] = { trainer: s.trainer, n: 0, npsSum: 0, npsN: 0, tsSum: 0, tsN: 0, last: "" });
      a.n++;
      if (nps != null) { a.npsSum += nps; a.npsN++; }
      if (tsc != null) { a.tsSum += tsc; a.tsN++; }
      if (s.dateISO && s.dateISO > a.last) a.last = s.dateISO;
    });
    return Object.keys(agg).map(function (k) {
      var a = agg[k];
      return { trainer: a.trainer, sessions: a.n, last: a.last,
               avgNps: a.npsN ? Math.round((a.npsSum / a.npsN) * 10) / 10 : null,
               avgScore: a.tsN ? Math.round((a.tsSum / a.tsN) * 100) / 100 : null };
    }).sort(function (a, b) { return (b.avgNps || -999) - (a.avgNps || -999); });
  }

  /* ---------------- rendering ---------------- */
  function pillHTML(kind) {
    // kind: "tracker" | "eval"
    var loaded = kind === "tracker" ? S.tracker : S.eval;
    var label = kind === "tracker" ? "Tracker" : "Eval";
    var cls = loaded ? "tr-pill on" : "tr-pill";
    var sub = loaded ? esc(loaded.name) : (kind === "tracker" ? "upload training tracker" : "upload MS Forms export");
    return '<button type="button" class="' + cls + '" data-trpill="' + kind + '">' +
           '<span class="tr-pill-label">' + label + '</span>' +
           '<span class="tr-pill-sub">' + sub + '</span></button>';
  }

  function statusCell(s) {
    if (s.status === "scored") return '<span class="tr-status scored">Scored</span>';
    if (s.status === "ready") return '<span class="tr-status ready">Needs verifying</span>';
    return '<span class="tr-status">Unscored</span>';
  }

  function filteredSessions() {
    var t = S.tracker;
    if (!t) return [];
    var q = P().normKey(S.q);
    return t.sessions.filter(function (s, i) {
      if (S.filter === "needs" && s.status !== "ready") return false;
      if (S.filter === "scored" && s.status !== "scored") return false;
      if (q) {
        var hay = P().normKey(s.trainer + " " + s.code + " " + s.client + " " + s.title + " " + s.dateISO);
        if (hay.indexOf(q) < 0) return false;
      }
      s._idx = i;
      return true;
    });
  }

  function listHTML() {
    var rows = filteredSessions();
    var shown = rows.slice(0, 400);
    var h = '<div class="tr-toolbar">' +
      '<div class="tr-filters">' +
      [["all", "All"], ["needs", "Needs verifying"], ["scored", "Scored"]].map(function (f) {
        return '<button type="button" class="tr-fpill' + (S.filter === f[0] ? " on" : "") +
               '" data-trfilter="' + f[0] + '">' + f[1] + '</button>';
      }).join("") + '</div>' +
      '<input type="search" class="tr-search" id="tr-search" placeholder="Search trainer, school, course…" value="' + esc(S.q) + '">' +
      '<div class="tr-viewbtns">' +
      '<button type="button" class="tr-fpill' + (S.view === "sessions" ? " on" : "") + '" data-trview="sessions">Sessions</button>' +
      '<button type="button" class="tr-fpill' + (S.view === "board" ? " on" : "") + '" data-trview="board">Leaderboard</button>' +
      '<button type="button" class="tr-fpill" data-traction="export">Export</button>' +
      '</div></div>';
    if (S.view === "board") return h + boardHTML();
    h += '<div class="tr-count">' + rows.length + ' session' + (rows.length === 1 ? "" : "s") +
         (rows.length > 400 ? ' (showing 400 — refine search)' : '') + '</div>';
    h += '<table class="table table-sm tr-table"><thead><tr><th>Date</th><th>Trainer</th><th>Course</th><th>School</th>' +
         '<th class="num">n</th><th class="num">NPS</th><th class="num">T-score</th><th>Status</th></tr></thead><tbody>';
    shown.forEach(function (s) {
      var c = s.computed || {};
      var nps = c.nps != null ? c.nps : s.nps, tsc = c.trainerScore != null ? c.trainerScore : s.trainerScore;
      h += '<tr data-trsession="' + s._idx + '" class="' + (s._idx === S.selected ? "sel" : "") + " st-" + s.status + '">' +
        '<td>' + esc(s.dateISO || "—") + '</td><td>' + esc(s.trainer) + '</td>' +
        '<td>' + esc(s.code) + '</td><td>' + esc(s.client) + '</td>' +
        '<td class="num">' + (c.n != null ? c.n : "—") + '</td>' +
        '<td class="num">' + (nps != null ? num(nps, 1) : "—") + '</td>' +
        '<td class="num">' + (tsc != null ? num(tsc, 2) : "—") + '</td>' +
        '<td>' + statusCell(s) + '</td></tr>';
    });
    return h + '</tbody></table>';
  }

  function boardHTML() {
    var rows = leaderboard();
    if (!rows.length) return '<div class="tr-empty">No scored sessions yet.</div>';
    var h = '<table class="table table-sm tr-table"><thead><tr><th>Trainer</th><th class="num">Sessions</th>' +
            '<th class="num">Avg NPS</th><th class="num">Avg T-score</th><th>Last session</th></tr></thead><tbody>';
    rows.forEach(function (r) {
      h += '<tr><td>' + esc(r.trainer) + '</td><td class="num">' + r.sessions + '</td>' +
           '<td class="num">' + (r.avgNps != null ? num(r.avgNps, 1) : "—") + '</td>' +
           '<td class="num">' + (r.avgScore != null ? num(r.avgScore, 2) : "—") + '</td>' +
           '<td>' + esc(r.last || "—") + '</td></tr>';
    });
    return h + '</tbody></table>';
  }

  function detailHTML() {
    var ev = S.eval;
    // Eval uploaded but not attached to any session: show filename suggestions.
    if (ev && ev.attachedTo < 0 && S.tracker) {
      var h = '<div class="tr-detail"><h3>Match this eval file to a session</h3>' +
        '<div class="tr-meta">File: <b>' + esc(ev.name) + '</b> (' + (ev.rows.length - 1) + ' responses)' +
        (ev.hint ? ' · filename: ' + esc(ev.hint.dateISO) + ' · ' + esc(ev.hint.session) + ' · ' + esc(ev.hint.school) : '') + '</div>';
      if (ev.suggested.length) {
        h += '<div class="tr-suggest">' + ev.suggested.slice(0, 5).map(function (s) {
          var idx = S.tracker.sessions.indexOf(s);
          return '<button type="button" class="tr-btn" data-traction="attach" data-idx="' + idx + '">' +
                 esc(s.dateISO || "") + ' · ' + esc(s.client) + ' · ' + esc(s.trainer) + ' (' + esc(s.code) + ')</button>';
        }).join("") + '</div>';
      } else {
        h += '<div class="tr-empty">No session matched the filename — pick one from the list.</div>';
      }
      return h + '</div>';
    }
    if (S.selected < 0 || !S.tracker) return "";
    var s = S.tracker.sessions[S.selected];
    var c = s.computed;
    var h = '<div class="tr-detail"><h3>' + esc(s.title || s.code) + '</h3>' +
      '<div class="tr-meta">' + esc(s.dateISO || "") + ' · ' + esc(s.client) + ' · ' +
      esc(s.code) + ' · Trainer: <b>' + esc(s.trainer) + '</b></div>';
    var ev = S.eval;
    if (!ev || ev.attachedTo !== S.selected) {
      h += '<div class="tr-empty">No eval file attached. ' +
           '<button type="button" class="tr-btn" data-traction="upload-eval">Upload eval for this session</button></div>';
    } else if (ev.needsMap) {
      h += mapperHTML();
    } else {
      // Filename cross-check: warn if the file suggests a different session.
      var hint = ev.hint, mismatch = false;
      if (hint && hint.dateISO && s.dateISO && hint.dateISO !== s.dateISO) mismatch = true;
      if (hint && hint.school && s.client && P().normKey(s.client).indexOf(P().normKey(hint.school)) < 0 &&
          P().normKey(hint.school).indexOf(P().normKey(s.client)) < 0) mismatch = true;
      h += '<div class="tr-evalok">Eval: <b>' + esc(ev.name) + '</b> (' + (ev.rows.length - 1) + ' responses)' +
           (hint ? ' <span class="tr-hint">filename: ' + esc(hint.dateISO) + ' · ' + esc(hint.session) + ' · ' + esc(hint.school) + '</span>' : '') +
           (mismatch ? ' <span class="tr-warn">Filename suggests a different session — double-check before generating.</span>' : '') +
           '</div>';
      if (!c) h += '<button type="button" class="tr-btn primary" data-traction="generate">Generate results</button>';
      else h += reportHTML(s, c);
    }
    return h + '</div>';
  }

  function reportHTML(s, c) {
    var delta = function (cur, prev) {
      if (cur == null || prev == null) return "";
      var d = Math.round((cur - prev) * 10) / 10;
      var cls = d > 0 ? "up" : (d < 0 ? "down" : "flat");
      return ' <span class="tr-delta ' + cls + '">(' + (d > 0 ? "+" : "") + d + ' vs prev)</span>';
    };
    var lowN = c.n < LOW_N;
    var h = '<div class="tr-kpis">' +
      '<div class="tr-kpi"><div class="tr-kpi-v">' + num(c.nps, 1) + '</div><div class="tr-kpi-l">Session NPS' +
        delta(c.nps, c.prevNps) + '</div></div>' +
      '<div class="tr-kpi"><div class="tr-kpi-v">' + num(c.trainerScore, 2) + '</div><div class="tr-kpi-l">Trainer score / 5' +
        delta(c.trainerScore, c.prevScore) + '</div></div>' +
      '<div class="tr-kpi"><div class="tr-kpi-v">' + c.n + '</div><div class="tr-kpi-l">Responses' +
        (lowN ? ' <span class="tr-warn">low-n</span>' : '') + '</div></div>' +
      '<div class="tr-kpi"><div class="tr-kpi-v">' + c.promoters + '/' + c.passives + '/' + c.detractors +
        '</div><div class="tr-kpi-l">Promoters / Passive / Detractors</div></div></div>';
    if (c.prevLabel) h += '<div class="tr-prev">Previous: ' + esc(c.prevLabel) + '</div>';
    // Facilitator item averages
    var labels = { t1: "Knowledge & mastery", t2: "Communication & presentation", t3: "Engage & connect" };
    h += '<div class="tr-items">';
    Object.keys(labels).forEach(function (k) {
      var v = c.itemAvgs[k];
      h += '<div class="tr-item"><span>' + labels[k] + '</span><b>' + (v != null ? num(v, 2) : "—") + '</b></div>';
    });
    h += '</div>';
    // Open-text comments
    var roles = S.eval.roles, rows = S.eval.rows.slice(1);
    var comments = [["c1", "Liked best"], ["c2", "Key insight"], ["c3", "Suggestions"], ["cnps", "Why this rating"]];
    var anyC = false, ch = '<div class="tr-comments">';
    comments.forEach(function (pair) {
      var ci = roles[pair[0]];
      if (ci == null || ci < 0) return;
      var items = [];
      rows.forEach(function (r) { var t = P().norm(r[ci]); if (t) items.push(t); });
      if (!items.length) return;
      anyC = true;
      ch += '<div class="tr-comment"><h4>' + pair[1] + ' (' + items.length + ')</h4><ul>' +
            items.slice(0, 50).map(function (t) { return '<li>' + esc(t) + '</li>'; }).join("") +
            (items.length > 50 ? '<li><i>…and ' + (items.length - 50) + ' more</i></li>' : '') + '</ul></div>';
    });
    h += anyC ? ch + '</div>' : '';
    // Write-back
    if (s.status === "ready") {
      h += '<div class="tr-writeback"><div class="tr-wb-preview">Write back: NPS <b>' + num(c.nps, 1) +
           '</b> → “Net Promoter Score”, T-score <b>' + num(c.trainerScore, 2) +
           '</b> → “Trainer Score” (new column) on row ' + (s.rowIdx + 1) + ' of ' + esc(S.tracker.name) + '.</div>' +
           '<button type="button" class="tr-btn primary" data-traction="writeback">Preview &amp; write to tracker</button></div>';
    } else if (s.status === "scored") {
      h += '<div class="tr-evalok">Written back to the tracker.</div>';
    }
    return h;
  }

  var ROLE_LABELS = [["nps", "NPS score (0–10)"], ["school", "School"], ["t1", "Trainer: knowledge"],
    ["t2", "Trainer: communication"], ["t3", "Trainer: engagement"], ["sat", "Satisfaction"],
    ["c1", "Comments: liked best"], ["c2", "Comments: key insight"], ["c3", "Comments: suggestions"], ["cnps", "Comments: why rating"]];
  function mapperHTML() {
    var ev = S.eval;
    var opts = ev.header.map(function (h, i) { return '<option value="' + i + '">' + esc(h.slice(0, 60)) + '</option>'; }).join("");
    var h = '<div class="tr-mapper"><h4>New form layout — map the columns</h4><p>This layout hasn’t been seen before. Confirm each role (saved for next time).</p>';
    ROLE_LABELS.forEach(function (pair) {
      var cur = ev.roles[pair[0]];
      h += '<label class="tr-maprow"><span>' + pair[1] + '</span><select data-trmap="' + pair[0] + '">' +
           '<option value="-1">— not in this form —</option>' + opts + '</select></label>';
    });
    return h + '<button type="button" class="tr-btn primary" data-traction="save-map">Save mapping</button></div>';
  }

  /* ---------------- assembly + events ---------------- */
  function render() {
    init(); // idempotent; loads saved column maps
    var host = $("trainers-body");
    if (!host) return;
    if (!on()) { host.innerHTML = ""; return; }
    var h = '<div class="tr-head"><h2>Trainers</h2>' +
            '<div class="tr-pills">' + pillHTML("tracker") + pillHTML("eval") + '</div></div>';
    if (!S.tracker) {
      h += '<div class="tr-empty">Upload your training tracker to list sessions. ' +
           '<button type="button" class="tr-btn" data-trpill="tracker">Upload tracker</button></div>';
    } else {
      h += '<div class="tr-sub">From <b>' + esc(S.tracker.name) + '</b> (' + esc(S.tracker.sheetName) + '): ' +
           S.tracker.sessions.length + ' sessions · ' + S.tracker.roster.length + ' trainers</div>';
      h += listHTML() + detailHTML();
    }
    host.innerHTML = h;
    wire(host);
  }

  function wire(host) {
    host.querySelectorAll("[data-trpill]").forEach(function (b) {
      b.addEventListener("click", function () { pickFile(b.getAttribute("data-trpill")); });
    });
    host.querySelectorAll("[data-trfilter]").forEach(function (b) {
      b.addEventListener("click", function () { S.filter = b.getAttribute("data-trfilter"); render(); });
    });
    host.querySelectorAll("[data-trview]").forEach(function (b) {
      b.addEventListener("click", function () { S.view = b.getAttribute("data-trview"); render(); });
    });
    host.querySelectorAll("[data-trsession]").forEach(function (tr) {
      tr.addEventListener("click", function () {
        S.selected = +tr.getAttribute("data-trsession");
        render();
      });
    });
    var search = host.querySelector("#tr-search");
    if (search) search.addEventListener("input", function () { S.q = search.value; renderListOnly(); });
    host.querySelectorAll("[data-traction]").forEach(function (b) {
      b.addEventListener("click", function () { action(b.getAttribute("data-traction"), b); });
    });
    // Pre-select the auto-detected mapping values.
    if (S.eval && S.eval.needsMap) {
      host.querySelectorAll("[data-trmap]").forEach(function (sel) {
        var cur = S.eval.roles[sel.getAttribute("data-trmap")];
        if (cur != null && cur >= 0) sel.value = String(cur);
      });
    }
  }

  function renderListOnly() {
    // Re-render list rows without losing the search box focus.
    var host = $("trainers-body");
    var pos = host.querySelector("#tr-search");
    var val = pos ? pos.value : S.q, start = pos ? pos.selectionStart : 0;
    render();
    var after = $("tr-search");
    if (after) { after.focus(); after.setSelectionRange(start, start); }
    void val;
  }

  function pickFile(kind) {
    if (kind === "eval" && S.selected < 0 && !S.tracker) return;
    var inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx,.xlsm,.xls";
    inp.onchange = function () {
      var f = inp.files && inp.files[0];
      if (!f) return;
      (kind === "tracker" ? openTracker(f) : openEval(f)).catch(function (e) {
        alert("Could not open " + f.name + ": " + e.message);
      });
    };
    inp.click();
  }

  function action(name, el) {
    if (name === "export") { exportResults(); return; }
    if (name === "upload-eval") { pickFile("eval"); return; }
    if (name === "attach") {
      var idx = el ? +el.getAttribute("data-idx") : -1;
      if (idx >= 0 && S.eval) { S.selected = idx; S.eval.attachedTo = idx; }
      render();
      return;
    }
    if (name === "generate") {
      var r = generateFor(S.selected);
      if (r && r.error) alert(r.error);
      render();
      return;
    }
    if (name === "save-map") {
      var host = $("trainers-body"), roles = {};
      host.querySelectorAll("[data-trmap]").forEach(function (sel) {
        roles[sel.getAttribute("data-trmap")] = +sel.value;
      });
      if (roles.nps == null || roles.nps < 0) { alert("The NPS (0–10) column is required."); return; }
      S.eval.roles = roles; S.eval.needsMap = false;
      S.maps[S.eval.sig] = { roles: roles, savedAt: new Date().toISOString() };
      saveMaps(); render();
      return;
    }
    if (name === "writeback") {
      var s = S.tracker.sessions[S.selected], c = s.computed || {};
      var msg = "Write to " + S.tracker.name + "?\n\nRow " + (s.rowIdx + 1) +
                ": Net Promoter Score = " + c.nps + ", Trainer Score = " + c.trainerScore +
                ".\n\nThis downloads an updated copy (-scored.xlsx).";
      if (window.confirm(msg)) writeback();
      return;
    }
  }

  function init() {
    if (S.built) return;
    S.built = true;
    loadMaps();
  }

  QBR.trainers = { render: render, init: init, _state: S };
})();
