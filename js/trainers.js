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
    tracker: null,   // { name, bytes, wb, sheetName, origin, headerRow, sessions[], roster[] }
    eval: null,      // { name, rows, header, roles, sig, attachedTo, suggested[] }
    selected: -1,    // index into sessions
    filter: "all",
    showingNewForm: false,   // all | needs | scored
    manualForm: false,   // standalone Add session (no eval)
    repTrainer: "", repSchool: "",  // summary report filters
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
        try { var wb = XLSX.read(fr.result, { type: "array" }); wb.__bytes = new Uint8Array(fr.result); res(wb); }
        catch (e) { rej(e); }
      };
      fr.onerror = function () { rej(new Error("Could not read " + file.name)); };
      fr.readAsArrayBuffer(file);
    });
  }
  function sheetRows(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) return [];
    // blankrows: row i of the result is sheet row origin.r + i (write-back relies on it)
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: true });
  }
  function sheetOrigin(wb, name) {
    var ws = wb.Sheets[name];
    return ws && ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).s : { r: 0, c: 0 };
  }
  // 1-based Excel row of a session (header = first used row of the sheet)
  function excelRow(s) { return s.rowIdx + 1 + ((S.tracker && S.tracker.origin) ? S.tracker.origin.r : 0); }

  /* ---------------- tracker ---------------- */
  function openTracker(file, handle) {
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
        name: file.name, bytes: wb.__bytes, wb: wb, sheetName: pick, origin: sheetOrigin(wb, pick),
        handle: handle || null,
        headerRow: sheetRows(wb, pick)[0] || [],
        sessions: sessions, roster: P().buildRoster(sessions),
        addedRows: [], acronymIndex: P().buildAcronymIndex(sessions.map(function (s) { return s.client; })),
        headerAdds: [],
        courseMap: courseMap(wb)
      };
      try { S.tracker.lastRow = XLSX.utils.decode_range(wb.Sheets[pick]["!ref"] || "A1").e.r; }
      catch (e) { S.tracker.lastRow = sessions.length; }
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
      var responseDate = P().detectResponseDate(rows);
      var suggested = hint && S.tracker ? P().suggestSessions(S.tracker.sessions, hint) : [];
      S.eval = {
        name: file.name, rows: rows, header: rows[0].map(P().norm),
        roles: roles, sig: det.sig, auto: det.roles,
        attachedTo: -1, suggested: suggested, hint: hint, responseDate: responseDate, needsMap: !S.maps[det.sig]
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
    var scores = data.map(function (r) { return r[roles.nps]; }); // raw: calcNPS skips blanks
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
    // Empty result (no valid answers): show — and offer no write-back.
    if (computed.nps != null || computed.trainerScore != null) session.status = "ready";
    return computed;
  }

  /* ---------------- write-back ---------------- */
  /* Apply the plan to a parsed workbook. Offsets: rows/cols are relative to
   * the sheet's used range (sheet_to_json), so add its start cell. */
  function applyPlan(wb, t, plan) {
    var ws = wb.Sheets[t.sheetName], o = t.origin || { r: 0, c: 0 };
    var npsCol = plan.npsCol, tsCol = plan.trainerCol, next = t.headerRow.length;
    var put = function (r, c, v) { XLSX.utils.sheet_add_aoa(ws, [[v]], { origin: { r: o.r + r, c: o.c + c } }); };
    if (plan.addNpsCol) { npsCol = next++; put(0, npsCol, "Net Promoter Score"); }
    if (plan.addTrainerCol) { tsCol = next++; put(0, tsCol, "Trainer Score"); }
    plan.edits.forEach(function (e) {
      if (npsCol >= 0 && e.nps != null) put(e.row, npsCol, e.nps);
      if (tsCol >= 0 && e.trainerScore != null) put(e.row, tsCol, e.trainerScore);
    });
  }
  function download(bytes, name) {
    var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  /* Format-safe: only the written cells change (QBR.xlsxSurgical, same engine
   * as Save to Excel). Plain SheetJS rewrite only as a confirmed fallback. */
  function writeback() {
    var t = S.tracker;
    if (!t) return Promise.resolve();
    var plan = P().planWriteback(t.headerRow, t.sessions);
    if (!plan.edits.length) return Promise.resolve();
    var SG = QBR.xlsxSurgical, bytes = t.bytes, base = null, work = t.wb;
    try {
      if (SG && bytes) {
        base = XLSX.read(bytes, { type: "array", cellStyles: true });
        work = XLSX.read(bytes, { type: "array", cellStyles: true });
      }
    } catch (e) { base = null; work = t.wb; }
    applyPlan(work, t, plan);
    var stem = t.name.replace(/\.(xlsx|xlsm|xls)$/i, "");
    async function finish(out, ext) { // ext null: tracker already holds these values, nothing to download
      t.bytes = new Uint8Array(out);
      t.wb = XLSX.read(t.bytes, { type: "array" });
      t.headerRow = sheetRows(t.wb, t.sheetName)[0] || [];
      t.origin = sheetOrigin(t.wb, t.sheetName);
      t.sessions.forEach(function (s) {
        if (s.computed && s.status === "ready") {
          if (s.computed.nps != null) s.nps = s.computed.nps;
          if (s.computed.trainerScore != null) s.trainerScore = s.computed.trainerScore;
          s.status = "scored";
          delete s.computed;
        }
      });
      // In-place save when the tracker was linked with a file handle:
      // writes into the original file, no download. Otherwise the -scored download.
      var savedInPlace = false;
      if (ext && t.handle) {
        try {
          var w = await t.handle.createWritable();
          await w.write(t.bytes);
          await w.close();
          savedInPlace = true;
        } catch (e) {
          try { if (w && w.abort) await w.abort(); } catch (x) {}
          var locked = /NoModificationAllowed|InvalidState|InvalidModification/i.test((e && e.name) || "") ||
                       /lock|in use|being used/i.test((e && e.message) || "");
          if (window.console) console.warn("[trainers] in-place save failed:", e);
          if (locked) window.alert("The file is open in Excel — downloaded a copy instead. Close it there and re-link to save in place.");
        }
      }
      if (ext && !savedInPlace) download(t.bytes, stem + "-scored" + ext);
      else if (!ext) window.alert("The tracker already has these values — nothing new to download.");
      else if (typeof QBR.persistNote === "function") QBR.persistNote("trainers", "Saved to " + t.name + " — scores written in place.", 6000);
      render();
    }
    function legacy(why) {
      if (why && !window.confirm("Format-safe write-back isn't possible for this file (" + why + ").\n\n" +
          "Download a plain copy instead? Cell formatting, conditional formatting and tables in the tracker may be lost.")) return;
      finish(XLSX.write(work, { bookType: "xlsx", type: "array" }), ".xlsx");
    }
    if (!base) { legacy(SG ? "file bytes unavailable" : null); return Promise.resolve(); }
    var ext = /\.xlsm$/i.test(t.name) ? ".xlsm" : ".xlsx";
    return Promise.resolve(SG.surgicalSave(bytes, base, work)).then(function (res) {
      if (!res || !res.ok) return legacy((res && res.reason) || "unsupported file");
      if (res.noChanges) return finish(bytes, null);
      return Promise.resolve(SG.verifySurgical(bytes, res.bytes, work)).then(function (why) {
        if (why) return legacy(why);
        finish(res.bytes, ext);
      });
    }).catch(function (e) { legacy(e && e.message || String(e)); });
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
    }).sort(function (a, b) {
      var x = a.avgNps == null ? -Infinity : a.avgNps, y = b.avgNps == null ? -Infinity : b.avgNps;
      return y === x ? 0 : (y > x ? 1 : -1); // NPS 0 / negative are real values
    });
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

  var MONTHS = ["January", "February", "March", "April", "May", "June", "July",
                "August", "September", "October", "November", "December"];
  function findHeader(headerRow, names) {
    for (var i = 0; i < headerRow.length; i++) {
      var h = P().normKey(headerRow[i]);
      for (var k = 0; k < names.length; k++) if (h === names[k]) return i;
    }
    return -1;
  }
  function timeToFrac(hhmm) {
    var m = /^(\d{1,2}):(\d{2})/.exec(hhmm || "");
    if (!m) return null;
    return Math.round(((+m[1]) * 60 + (+m[2])) / 1440 * 10000) / 10000;
  }

  function courseMap(wb) {
    try {
      var ws = wb.Sheets["03 Course List"];
      if (!ws) return { byCode: {}, byTitle: {} };
      return P().buildCourseMap(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }));
    } catch (e) { return { byCode: {}, byTitle: {} }; }
  }
  function tenantSchoolRoster() {
    try {
      var risky = (typeof APP !== "undefined" && APP.model && APP.model.risky) || [];
      if (!risky.length) return null;
      var map = {};
      risky.forEach(function (r) {
        var name = String(r.schoolRaw || "").trim();
        if (!name) return;
        if (!map[name]) map[name] = { name: name, org: "" };
        var org = r.org ? String(r.org).trim() : "";
        if (org && !map[name].org) map[name].org = org;
      });
      return Object.keys(map).sort().map(function (k) { return map[k]; });
    } catch (e) { return null; }
  }

  function schoolOptionsHTML() {
    var t = S.tracker;
    var roster = tenantSchoolRoster();
    var names = [], seen = {};
    var add = function (nm) {
      if (nm && !seen[nm.toLowerCase()]) { seen[nm.toLowerCase()] = true; names.push(nm); }
    };
    if (roster && roster.length) roster.forEach(function (s) { add(s.name); });
    t.sessions.forEach(function (s) { add(s.client); });
    var html = names.sort().map(function (nm) {
      return '<option value="' + esc(nm) + '">';
    }).join("");
    return { html: html, src: (roster && roster.length) ? "ALL TENANTS tracker" : "training tracker" };
  }

  function newSessionFormHTML() {
    var t = S.tracker, ev = S.eval, hint = ev && ev.hint;
    var date = "", dateSrc = "";
    if (ev && ev.responseDate) { date = ev.responseDate; dateSrc = "from response timestamps"; }
    else if (hint && hint.dateISO) { date = hint.dateISO; dateSrc = "from filename"; }
    var school = "";
    if (hint && hint.school) {
      school = t.acronymIndex[hint.school.toUpperCase()] || hint.school;
    }
    var t0 = hint && hint.session === "PM" ? "13:00" : "08:00";
    var t1 = hint && hint.session === "PM" ? "17:00" : "12:00";
    var trainers = t.roster.map(function (r) {
      return '<option value="' + esc(r) + '">';
    }).join("");
    var cm = t.courseMap || { byCode: {}, byTitle: {} };
    var codes = Object.keys(cm.byCode).sort().map(function (c) {
      return '<option value="' + esc(c) + '">';
    }).join("");
    var titles = Object.keys(cm.byTitle).sort().map(function (x) {
      return '<option value="' + esc(x) + '">';
    }).join("");
    var sch = schoolOptionsHTML();
    return '<div class="tr-newform"><h4>Add this as a new session</h4>' +
      '<div class="tr-newgrid">' +
      '<label>Date* <input type="date" id="trns-date" value="' + esc(date) + '">' +
      (dateSrc ? '<span class="tr-hint">' + dateSrc + "</span>" : "") + "</label>" +
      '<label>Trainer* <input id="trns-trainer" list="trns-trainers" placeholder="type or pick"><datalist id="trns-trainers">' + trainers + "</datalist></label>" +
      '<label>Course name <input id="trns-title" list="trns-titles" placeholder="start typing the course name"><datalist id="trns-titles">' + titles + "</datalist></label>" +
      '<label>Course code* <input id="trns-code" list="trns-codes" placeholder="auto-filled from the name"><datalist id="trns-codes">' + codes +
        '</datalist><span class="tr-hint" id="trns-codehint"></span></label>' +
      '<label>School* <input id="trns-school" list="trns-schools" value="' + esc(school) + '" placeholder="type or pick"><datalist id="trns-schools">' + sch.html +
        '</datalist><span class="tr-hint">suggestions from ' + sch.src + "</span></label>" +
      '<label>Time start <input type="time" id="trns-t0" value="' + t0 + '"></label>' +
      '<label>Time end <input type="time" id="trns-t1" value="' + t1 + '"></label>' +
      "</div>" +
      '<button type="button" class="tr-btn primary" data-traction="save-new-session">Save session</button> ' +
      '<button type="button" class="tr-btn" data-traction="' + (S.manualForm ? 'cancel-manual' : 'cancel-new-session') + '">Cancel</button></div>';
  }

  function saveNewSession() {
    var host = $("trainers-body"), t = S.tracker, ev = S.eval;
    if (!t) return;
    var g = function (id) { var el = host.querySelector("#trns-" + id); return el ? el.value.trim() : ""; };
    var date = g("date"), trainer = g("trainer"), code = g("code"), client = g("school");
    if (!date || !trainer || !code || !client) { alert("Date, trainer, course code and school are required."); return; }
    var year = date.slice(0, 4);
    var sid = P().nextSessionId(t.sessions, year);
    // Ensure a Session ID column exists (in-memory; written at write-back).
    var idCol = findHeader(t.headerRow, P().TRACKER_COLS.id);
    if (idCol < 0) { t.headerRow.push("Session ID"); t.headerAdds = t.headerAdds || []; t.headerAdds.push("Session ID"); }
    var row = new Array(t.headerRow.length).fill("");
    var set = function (names, val) {
      var i = findHeader(t.headerRow, names);
      if (i >= 0) row[i] = val;
    };
    var ts = timeToFrac(g("t0")), te = timeToFrac(g("t1"));
    set(["date"], date);
    set(["year"], year);
    set(["month"], MONTHS[+date.slice(5, 7) - 1] || "");
    set(P().TRACKER_COLS.trainer, trainer);
    set(P().TRACKER_COLS.code, code);
    set(P().TRACKER_COLS.title, g("title"));
    set(P().TRACKER_COLS.client, client);
    set(["time start", "start"], ts == null ? "" : ts);
    set(["time end", "end"], te == null ? "" : te);
    if (ts != null && te != null) set(["duration"], Math.round((te - ts) * 24 * 100) / 100);
    set(P().TRACKER_COLS.id, sid);

    var rowIdx = t.lastRow + 1 + t.addedRows.length;
    t.addedRows.push(row);
    var d = new Date(+date.slice(0, 4), +date.slice(5, 7) - 1, +date.slice(8, 10));
    var session = {
      rowIdx: rowIdx, id: sid, date: d, dateISO: date,
      trainer: trainer, code: code, title: g("title"), client: client,
      timeStart: ts, timeEnd: te, nps: null, trainerScore: null,
      status: "unscored", isNew: true
    };
    t.sessions.push(session);
    // Keep the roster complete for future picks.
    if (t.roster.indexOf(trainer) < 0) { t.roster.push(trainer); t.roster.sort(); }
    S.selected = t.sessions.length - 1;
    if (ev) ev.attachedTo = S.selected;
    S.showingNewForm = false;
    S.manualForm = false;
    render();
  }

  function kpiStripHTML() {
    var t = S.tracker, sum = P().buildSummary(t.sessions);
    return '<div class="tr-kpis tr-kpis-top">' +
      kpi(sum.total, "Sessions", "all sessions in tracker") +
      kpi(sum.respondents, "Respondents", sum.respSessions + " session(s) scored") +
      kpi(sum.avgNps == null ? "\u2014" : sum.avgNps, "Avg NPS", sum.scoredNps + " scored session(s)") +
      kpi(sum.avgScore == null ? "\u2014" : sum.avgScore, "Avg trainer score", sum.scoredTs + " scored session(s)") +
      kpi(sum.repeatRate + "%", "Repeat schools", "sessions at repeat schools") +
      '</div>';
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
      '<button type="button" class="tr-fpill' + (S.view === "summary" ? " on" : "") + '" data-trview="summary">Summary</button>' +
      '<button type="button" class="tr-fpill' + (S.view === "gen" ? " on" : "") + '" data-trview="gen">Feedback Generator</button>' +
      '<button type="button" class="tr-fpill" data-traction="export">Export</button>' +
      '</div></div>';
    if (S.view === "board") return h + boardHTML();
    if (S.view === "summary") return h + summaryHTML();
    if (S.view === "gen") return h + '<iframe class="tr-genframe" src="feedback-generator.html" title="Feedback Generator"></iframe>';
    h += kpiStripHTML();
    h += '<div class="tr-addrow"><button type="button" class="tr-btn primary" data-traction="add-session">+ Add session</button>' +
         '<span class="tr-hint">manual entry \u2014 no eval file needed; type or pick every field</span></div>';
    if (S.manualForm) h += newSessionFormHTML();
    h += '<div class="tr-count">' + rows.length + ' session' + (rows.length === 1 ? "" : "s") +
         (rows.length > 400 ? ' (showing 400 — refine search)' : '') + '</div>';
    h += '<table class="table table-sm tr-table" data-resizable="static"><thead><tr><th>Date</th><th>Trainer</th><th>Course</th><th>School</th>' +
         '<th class="num">Responses</th><th class="num">NPS</th><th class="num">Trainer score</th><th>Status</th></tr></thead><tbody>';
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

  function kpi(label, value, sub) {
    return '<div class="tr-kpi"><div class="tr-kpi-v">' + value + '</div>' +
           '<div class="tr-kpi-l">' + label + '</div>' +
           (sub ? '<div class="tr-kpi-s">' + sub + '</div>' : '') + '</div>';
  }

  function summaryHTML() {
    var t = S.tracker;
    var rsess = t.sessions.filter(function (s) {
      if (S.repTrainer && s.trainer !== S.repTrainer) return false;
      if (S.repSchool && s.client !== S.repSchool) return false;
      return true;
    });
    var trainers = t.roster.slice().sort();
    var schools = [];
    t.sessions.forEach(function (s) { if (s.client && schools.indexOf(s.client) < 0) schools.push(s.client); });
    schools.sort();
    var sum = P().buildSummary(rsess);
    var today = new Date();
    var stamp = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" +
                String(today.getDate()).padStart(2, "0");
    var filtDesc = [];
    if (S.repTrainer) filtDesc.push('Trainer: ' + S.repTrainer);
    if (S.repSchool) filtDesc.push('School: ' + S.repSchool);
    var h = '<div class="tr-report" id="tr-report">' +
      '<div class="tr-rpt-filters tr-noprint">' +
      '<label>Trainer <select id="tr-rep-trainer"><option value="">All trainers</option>' +
      trainers.map(function (r) { return '<option value="' + esc(r) + '"' + (S.repTrainer === r ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') +
      '</select></label>' +
      '<label>School <select id="tr-rep-school"><option value="">All schools</option>' +
      schools.map(function (s) { return '<option value="' + esc(s) + '"' + (S.repSchool === s ? ' selected' : '') + '>' + esc(s) + '</option>'; }).join('') +
      '</select></label>' +
      ((S.repTrainer || S.repSchool) ? '<button type="button" class="tr-btn" data-traction="rep-clear">Clear</button>' : '') +
      '</div>' +
      '<div class="tr-rpt-head"><div><h3>Training Summary Report</h3>' +
      '<div class="tr-meta">Generated ' + stamp + ' · ' + esc(t.name) + ' · ' + sum.total + ' session' + (sum.total === 1 ? '' : 's')
      + (filtDesc.length ? ' · ' + esc(filtDesc.join(' · ')) : '') + '</div></div>' +
      '<button type="button" class="tr-btn primary tr-noprint" data-traction="print-report">Print / PDF</button></div>' +
      '<div class="tr-kpis">' +
      kpi(sum.total, "Sessions", "all sessions in tracker") +
      kpi(sum.hours, "Training hours", "sessions with start/end times") +
      kpi(sum.respondents, "Respondents", sum.respSessions + " session(s) scored in the dashboard") +
      kpi(sum.avgNps == null ? "—" : sum.avgNps, "Avg NPS", sum.scoredNps + " scored session(s)") +
      kpi(sum.avgScore == null ? "—" : sum.avgScore, "Avg trainer score", sum.scoredTs + " scored session(s)") +
      kpi(sum.repeatRate + "%", "Repeat-school rate", "sessions at repeat schools") +
      '</div>';
    // Monthly trend (CSS bars, offline-safe)
    if (sum.monthly.length) {
      var maxM = Math.max.apply(null, sum.monthly.map(function (m) { return m.sessions; }));
      h += '<h4>Sessions by month</h4><div class="tr-bars">' +
        sum.monthly.map(function (m) {
          var w = maxM ? Math.round(m.sessions / maxM * 100) : 0;
          return '<div class="tr-bar-row"><span class="tr-bar-k">' + esc(m.key) + '</span>' +
                 '<span class="tr-bar-track"><span class="tr-bar-fill" style="width:' + w + '%"></span></span>' +
                 '<span class="tr-bar-v">' + m.sessions + (m.avgNps != null ? ' · NPS ' + m.avgNps : '') + '</span></div>';
        }).join("") + '</div>';
    }
    // Breakdowns
    function breakTable(title, rows) {
      var b = '<h4>' + title + '</h4><table class="table table-sm tr-table"><thead><tr>' +
              '<th>' + title.replace(" by ", "") + '</th><th>Sessions</th><th>Avg NPS</th><th>Avg trainer score</th>' +
              '</tr></thead><tbody>';
      rows.forEach(function (r) {
        b += '<tr><td>' + esc(r.key) + '</td><td class="num">' + r.sessions + '</td>' +
             '<td class="num">' + (r.avgNps == null ? "—" : r.avgNps) + '</td>' +
             '<td class="num">' + (r.avgScore == null ? "—" : r.avgScore) + '</td></tr>';
      });
      return b + '</tbody></table>';
    }
    if (sum.bySchool.length) h += breakTable("Sessions by school", sum.bySchool);
    if (sum.byCourse.length) h += breakTable("Sessions by course", sum.byCourse);
    return h + '</div>';
  }

  function boardHTML() {
    var rows = leaderboard();
    if (!rows.length) return '<div class="tr-empty">No scored sessions yet.</div>';
    var h = '<table class="table table-sm tr-table" data-resizable="static"><thead><tr><th>Trainer</th><th class="num">Sessions</th>' +
            '<th class="num">Avg NPS</th><th class="num">Avg trainer score</th><th>Last session</th></tr></thead><tbody>';
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
        h += '<div class="tr-empty">No session matched the filename — click a session in the list to attach this file to it.</div>';
      }
      if (S.showingNewForm) {
        h += newSessionFormHTML();
      } else {
        h += '<div><button type="button" class="tr-btn" data-traction="new-session">Add as new session</button> ' +
             '<span class="tr-hint">logs it in the tracker, then scores it like any other session</span></div>';
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
           '</b> → “Net Promoter Score”, trainer score <b>' + num(c.trainerScore, 2) +
           '</b> → “Trainer Score” on row ' + excelRow(s) + ' of ' + esc(S.tracker.name) + '.</div>' +
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
      h += '<div class="tr-viewbtns">' +
        '<button type="button" class="tr-fpill' + (S.view === "gen" ? " on" : "") + '" data-trview="gen">Feedback Generator</button></div>';
      if (S.view === "gen") h += '<iframe class="tr-genframe" src="feedback-generator.html" title="Feedback Generator"></iframe>';
      else h += '<div class="tr-empty">Upload your training tracker to list sessions. ' +
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
    // New-session form: course code <-> session title auto-fill (two-way).
    var codeEl = host.querySelector("#trns-code"), titleEl = host.querySelector("#trns-title");
    if (codeEl && titleEl && S.tracker) {
      var cm = S.tracker.courseMap || { byCode: {}, byTitle: {} };
      var hintEl = host.querySelector("#trns-codehint");
      var showHint = function () {
        if (!hintEl) return;
        var m = cm.byCode[P().norm(codeEl.value)];
        hintEl.textContent = m ? (m.product || "") + (m.duration != null ? " · " + m.duration + "h" : "") : "";
      };
      codeEl.addEventListener("input", function () { codeEl.setAttribute("data-auto", "0"); showHint(); });
      codeEl.addEventListener("change", function () {
        var m = cm.byCode[P().norm(codeEl.value)];
        showHint();
        if (m && m.title && (!titleEl.value || titleEl.getAttribute("data-auto") === "1")) {
          titleEl.value = m.title;
          titleEl.setAttribute("data-auto", "1");
        }
      });
      var findCodeForTitle = function (txt) {
        var key = P().norm(txt);
        if (cm.byTitle[key]) return cm.byTitle[key];
        // Fallback: unique case-insensitive prefix match (handles partial typing).
        var low = key.toLowerCase(), hits = [];
        Object.keys(cm.byTitle).forEach(function (k) {
          if (k.toLowerCase().indexOf(low) === 0) hits.push(cm.byTitle[k]);
        });
        return hits.length === 1 ? hits[0] : null;
      };
      titleEl.addEventListener("change", function () {
        var cc = findCodeForTitle(titleEl.value);
        if (cc && (!codeEl.value || codeEl.getAttribute("data-auto") === "1")) {
          codeEl.value = cc;
          codeEl.setAttribute("data-auto", "1");
          showHint();
        }
      });
      // Live assist while typing: fill when the partial text uniquely matches.
      titleEl.addEventListener("input", function () {
        titleEl.setAttribute("data-auto", "0");
        if (codeEl.value && codeEl.getAttribute("data-auto") !== "1") return;
        var cc = findCodeForTitle(titleEl.value);
        if (cc) { codeEl.value = cc; codeEl.setAttribute("data-auto", "1"); showHint(); }
      });
      showHint();
    }
    var repT = host.querySelector("#tr-rep-trainer"), repS = host.querySelector("#tr-rep-school");
    if (repT) repT.addEventListener("change", function () { S.repTrainer = repT.value; render(); });
    if (repS) repS.addEventListener("change", function () { S.repSchool = repS.value; render(); });
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

  function fsSupported() {
    if (QBR.fsSupported) return QBR.fsSupported();
    return typeof window.showOpenFilePicker === "function";
  }
  async function pickTrackerHandle() {
    var handle;
    try {
      var picks = await window.showOpenFilePicker({
        multiple: false,
        types: [{ description: "Excel workbook", accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] } }],
      });
      handle = picks[0];
    } catch (e) { return; }
    var perm = true;
    try {
      if (typeof QBR.fsEnsurePermission === "function") perm = await QBR.fsEnsurePermission(handle);
      else {
        var q = await handle.queryPermission({ mode: "readwrite" });
        perm = q === "granted" || (await handle.requestPermission({ mode: "readwrite" })) === "granted";
      }
    } catch (e) { perm = false; }
    if (!perm) { alert("Permission was not granted \u2014 the tracker was not linked."); return; }
    var file;
    try { file = await handle.getFile(); }
    catch (e) { alert("Could not read the file: " + (e && e.message || e)); return; }
    openTracker(file, handle).catch(function (e) {
      alert("Could not open " + file.name + ": " + e.message);
    });
  }
  function pickFile(kind) {
    if (kind === "eval" && S.selected < 0 && !S.tracker) return;
    if (kind === "tracker" && fsSupported()) { pickTrackerHandle(); return; }
    var inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx,.xlsm,.xls";
    inp.onchange = function () {
      var f = inp.files && inp.files[0];
      if (!f) return;
      (kind === "tracker" ? openTracker(f, null) : openEval(f)).catch(function (e) {
        alert("Could not open " + f.name + ": " + e.message);
      });
    };
    inp.click();
  }

  function action(name, el) {
    if (name === "export") { exportResults(); return; }
    if (name === "rep-clear") { S.repTrainer = ""; S.repSchool = ""; render(); return; }
    if (name === "print-report") {
      document.body.classList.add("tr-printing");
      document.body.setAttribute("data-print", "dash-trainers");
      var done = function () { document.body.classList.remove("tr-printing");
        document.body.removeAttribute("data-print");
        window.removeEventListener("afterprint", done); };
      return;
    }
    if (name === "upload-eval") { pickFile("eval"); return; }
    if (name === "attach") {
      var idx = el ? +el.getAttribute("data-idx") : -1;
      if (idx >= 0 && S.eval) { S.selected = idx; S.eval.attachedTo = idx; }
      render();
      return;
    }
    if (name === "new-session") { S.showingNewForm = true; render(); return; }
    if (name === "cancel-new-session") { S.showingNewForm = false; render(); return; }
    if (name === "cancel-manual") { S.manualForm = false; render(); return; }
    if (name === "add-session") { S.manualForm = true; render(); return; }
    if (name === "save-new-session") { saveNewSession(); return; }
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
      // List EVERY pending session: write-back writes all "Needs verifying" rows.
      var pend = S.tracker.sessions.filter(function (x) { return x.computed && x.status === "ready"; });
      if (!pend.length) return;
      var lines = pend.slice(0, 12).map(function (x) {
        return "Row " + excelRow(x) + " (" + (x.dateISO || "—") + " · " + x.trainer + " · " + x.client + "): NPS = " +
               num(x.computed.nps, 1) + ", Trainer Score = " + num(x.computed.trainerScore, 2);
      });
      if (pend.length > 12) lines.push("…and " + (pend.length - 12) + " more");
      var inplace = !!S.tracker.handle;
      var msg = "Write " + pend.length + " session" + (pend.length === 1 ? "" : "s") + " to " + S.tracker.name + "?\n\n" +
                lines.join("\n") + "\n\n" + (inplace ? "Saves directly into the linked file." : "Downloads an updated copy (-scored).");
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
