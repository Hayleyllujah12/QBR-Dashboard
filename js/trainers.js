/* =============================================================================
 * trainers.js — Training › Trainers tab: trainer evaluations from MS Forms exports.
 * ----------------------------------------------------------------------------
 * Sessions come from the training tracker workbook (ETG_Training_Tracker style):
 * pick a session from the log, add its MS Forms eval export (one file = one
 * session), generate results, write back to the tracker. An eval can also be
 * opened on its own (no tracker) for a one-off session feedback report.
 *
 * Scores per session:
 *   - Session NPS (standard: % Promoters 9-10 minus % Detractors 0-6, blanks skipped)
 *   - Trainer effectiveness = mean of the 3 facilitator items (5-pt)
 *   - Feedback report: rating items, satisfaction, schools, themes, recommendations
 *     (QBR.trainersParse.analyzeEval — replaces the v7 standalone generator page)
 * The eval file carries NO trainer column; the trainer is inherited from the
 * tracker session the user confirms. The filename (date/AM-PM/school) only suggests.
 *
 * Write-back (format-safe, QBR.xlsxSurgical): NPS -> "Net Promoter Score",
 * trainer score -> "Trainer Score", plus any sessions added here as new rows.
 * Saved in place when the tracker was opened with a file handle and has not
 * changed since; otherwise downloads <name>-scored.xlsx.
 * Reports print to PDF through the browser (offline; charts as light-theme images).
 * Everything offline; eval answers in memory only, column maps in localStorage.
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
    tracker: null,   // { name, bytes, wb, sheetName, origin, headerRow, sessions[], roster[], addedRows[], headerAdds[], handle }
    eval: null,      // { name, rows, header, roles, sig, attachedTo, suggested[], hint, responseDate, needsMap, solo }
    selected: -1,    // index into sessions
    filter: "all",   // all | needs | scored | unscored
    showingNewForm: false, // "add as new session" from an unmatched eval
    manualForm: false,     // standalone Add session (no eval)
    repTrainer: "", repSchool: "",  // summary filters
    q: "",
    view: "sessions",      // sessions | board | summary
    maps: {},
    charts: {}, chartDefs: {}
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function on() { try { return typeof QBR.adminIsOn !== "function" || QBR.adminIsOn("trainers"); } catch (e) { return true; } }
  function num(v, d) { return v == null || !isFinite(v) ? "—" : Number(v).toFixed(d == null ? 1 : d); }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function cssVar(n, fb) { try { var v = getComputedStyle(document.body).getPropertyValue(n).trim(); return v || fb; } catch (e) { return fb; } }

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
  // Last row (relative to the used range) that holds any value: new sessions go right below it.
  function lastDataRow(rows) {
    for (var i = rows.length - 1; i > 0; i--)
      if ((rows[i] || []).some(function (v) { return v !== "" && v != null; })) return i;
    return 0;
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
        for (var i = 0; i < names.length && !sessions.length; i++) {
          if (names[i] === pick) continue;
          sessions = P().parseTrackerSheet(sheetRows(wb, names[i]));
          if (sessions.length) pick = names[i];
        }
      }
      if (!sessions.length) throw new Error("No session log found in " + file.name);
      var rows = sheetRows(wb, pick);
      S.tracker = {
        name: file.name, bytes: wb.__bytes, wb: wb, sheetName: pick, origin: sheetOrigin(wb, pick),
        handle: handle || null, mtime: file.lastModified || 0, size: file.size || 0,
        headerRow: rows[0] || [], lastRow: lastDataRow(rows),
        sessions: sessions, roster: P().buildRoster(sessions),
        addedRows: [], headerAdds: [],
        acronymIndex: P().buildAcronymIndex(sessions.map(function (s) { return s.client; })),
        courseMap: courseMap(wb)
      };
      S.selected = -1; S.manualForm = false; S.showingNewForm = false;
      // Keep an eval that was opened on its own: offer matches from the new tracker.
      if (S.eval) {
        S.eval.solo = null; S.eval.attachedTo = -1;
        S.eval.suggested = S.eval.hint ? P().suggestSessions(sessions, S.eval.hint) : [];
      }
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
        attachedTo: -1, suggested: suggested, hint: hint, responseDate: P().detectResponseDate(rows),
        needsMap: !S.maps[det.sig], solo: null
      };
      S.showingNewForm = false;
      if (S.tracker && S.selected >= 0) attachTo(S.selected);
      render();
    });
  }

  // Attach the eval to a session. Results generated from a DIFFERENT eval file are dropped
  // (they would otherwise be shown and written under the new file's name).
  function attachTo(idx) {
    var s = S.tracker && S.tracker.sessions[idx];
    if (!s || !S.eval) return;
    S.eval.attachedTo = idx;
    if (s.computed && s.computed.evalName !== S.eval.name) {
      delete s.computed;
      s.status = s.nps != null || s.trainerScore != null ? "scored" : "unscored";
    }
  }

  /* ---------------- scoring ---------------- */
  function computeFrom(ev) {
    var roles = ev.roles;
    if (roles.nps == null || roles.nps < 0) return { error: "Map the NPS (0–10) column first." };
    var rep = P().analyzeEval(ev.rows, roles);
    return {
      n: rep.nps.n, promoters: rep.nps.promoters, passives: rep.nps.passives, detractors: rep.nps.detractors,
      nps: rep.nps.nps, trainerScore: rep.trainer.score, itemAvgs: rep.trainer.itemAvgs,
      report: rep, evalName: ev.name
    };
  }
  function generateFor(idx) {
    var t = S.tracker, ev = S.eval;
    if (!ev) return null;
    if (!t) { // eval opened on its own: report only, nothing to write
      var solo = computeFrom(ev);
      if (!solo.error) ev.solo = solo;
      return solo;
    }
    if (ev.attachedTo !== idx) return null;
    var computed = computeFrom(ev);
    if (computed.error) return computed;
    var session = t.sessions[idx];
    var prev = P().findPrevSession(t.sessions, session);
    computed.prevNps = prev ? prev.nps : null;
    computed.prevScore = prev ? prev.trainerScore : null;
    computed.prevLabel = prev ? (prev.dateISO || "") + " · " + prev.client : "";
    session.computed = computed;
    // Empty result (no valid answers): show — and offer no write-back.
    session.status = (computed.nps != null || computed.trainerScore != null) ? "ready" : (session.nps != null ? "scored" : "unscored");
    return computed;
  }
  function pendingScores() {
    return S.tracker ? S.tracker.sessions.filter(function (x) { return x.computed && x.status === "ready"; }) : [];
  }

  /* ---------------- write-back ---------------- */
  /* Apply the plan to a parsed workbook. Offsets: rows/cols are relative to
   * the sheet's used range (sheet_to_json), so add its start cell. */
  function applyPlan(wb, t, plan) {
    var ws = wb.Sheets[t.sheetName], o = t.origin || { r: 0, c: 0 };
    var rng = XLSX.utils.decode_range(ws["!ref"] || "A1");
    var put = function (r, c, v) {
      var R = o.r + r, C = o.c + c, ref = XLSX.utils.encode_cell({ r: R, c: C });
      if (v instanceof Date) ws[ref] = { t: "d", v: v };
      else if (typeof v === "number") ws[ref] = { t: "n", v: v };
      else ws[ref] = { t: "s", v: String(v) };
      if (R > rng.e.r) rng.e.r = R; if (C > rng.e.c) rng.e.c = C;
    };
    var npsCol = plan.npsCol, tsCol = plan.trainerCol, next = t.headerRow.length;
    (t.headerAdds || []).forEach(function (h) {
      var c = t.headerRow.indexOf(h);
      if (c >= 0) put(0, c, h);
    });
    if (plan.addNpsCol) { npsCol = next++; put(0, npsCol, "Net Promoter Score"); }
    if (plan.addTrainerCol) { tsCol = next++; put(0, tsCol, "Trainer Score"); }
    (t.addedRows || []).forEach(function (a) {
      a.cells.forEach(function (v, c) { if (v !== "" && v != null) put(a.rowIdx, c, v); });
    });
    plan.edits.forEach(function (e) {
      if (npsCol >= 0 && e.nps != null) put(e.row, npsCol, e.nps);
      if (tsCol >= 0 && e.trainerScore != null) put(e.row, tsCol, e.trainerScore);
    });
    ws["!ref"] = XLSX.utils.encode_range(rng);
  }
  function download(bytes, name) {
    var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function note(msg) {
    if (typeof QBR.persistNote === "function") { try { QBR.persistNote("trainers", esc(msg), 6000); } catch (e) {} }
    var el = $("tr-status"); if (el) el.textContent = msg;
  }
  /* In place only when the linked file is unchanged since it was opened (Excel or
   * OneDrive may have saved it meanwhile — overwriting would lose those edits). */
  async function saveInPlace(t) {
    if (!t.handle) return "no-handle";
    try {
      var cur = await t.handle.getFile();
      if (cur && ((t.mtime && cur.lastModified !== t.mtime) || (t.size && cur.size !== t.size))) return "changed";
    } catch (e) { return "unreadable"; }
    var w = null;
    try {
      w = await t.handle.createWritable();
      await w.write(t.bytes);
      await w.close();
      try { var f = await t.handle.getFile(); t.mtime = f.lastModified; t.size = f.size; } catch (e) {}
      return "ok";
    } catch (e) {
      try { if (w && w.abort) await w.abort(); } catch (x) {}
      if (window.console) console.warn("[trainers] in-place save failed:", e);
      return "locked";
    }
  }
  /* Format-safe: only the written cells change (QBR.xlsxSurgical, same engine
   * as Save to Excel). Plain SheetJS rewrite only as a confirmed fallback. */
  function writeback() {
    var t = S.tracker;
    if (!t) return Promise.resolve();
    var plan = P().planWriteback(t.headerRow, t.sessions);
    if (!plan.edits.length && !t.addedRows.length) return Promise.resolve();
    var SG = QBR.xlsxSurgical, bytes = t.bytes, base = null, work = null;
    // Always edit a fresh parse — a cancelled fallback must leave the loaded tracker untouched.
    try {
      if (bytes) work = XLSX.read(bytes, { type: "array", cellStyles: true });
      if (SG && bytes) base = XLSX.read(bytes, { type: "array", cellStyles: true });
    } catch (e) { base = null; }
    if (!work) { window.alert("The tracker file can't be read again — open it again and retry."); return Promise.resolve(); }
    applyPlan(work, t, plan);
    var stem = t.name.replace(/\.(xlsx|xlsm|xls)$/i, "");
    async function finish(out, ext) { // ext null: tracker already holds these values
      t.bytes = new Uint8Array(out);
      t.wb = XLSX.read(t.bytes, { type: "array" });
      var rows = sheetRows(t.wb, t.sheetName);
      t.headerRow = rows[0] || [];
      t.origin = sheetOrigin(t.wb, t.sheetName);
      t.lastRow = lastDataRow(rows);
      t.addedRows = []; t.headerAdds = [];
      t.sessions.forEach(function (s) {
        delete s.isNew;
        if (s.computed && s.status === "ready") {
          if (s.computed.nps != null) s.nps = s.computed.nps;
          if (s.computed.trainerScore != null) s.trainerScore = s.computed.trainerScore;
          s.status = "scored";
          s.written = s.computed; // keep the report viewable after writing
          delete s.computed;
        }
      });
      if (!ext) { window.alert("The tracker already has these values — nothing new to write."); render(); return; }
      var r = await saveInPlace(t);
      if (r === "ok") note("Saved to " + t.name + " — changes written in place.");
      else {
        if (r === "changed") window.alert(t.name + " changed after you opened it (Excel or OneDrive saved it). To keep those edits, a copy was downloaded instead — open the tracker again to save in place.");
        else if (r === "locked") window.alert("Couldn't write to " + t.name + " (open in Excel?) — downloaded a copy instead.");
        download(t.bytes, stem + "-scored" + ext);
        note("Downloaded " + stem + "-scored" + ext + ".");
      }
      render();
    }
    function legacy(why) {
      if (why && !window.confirm("Format-safe write-back isn't possible for this file (" + why + ").\n\n" +
          "Download a plain copy instead? Cell formatting, conditional formatting and tables in the tracker may be lost.")) return;
      // The plain copy has lost formatting (and macros): never write it — or anything built on it later —
      // into the linked original. Unlink for good; reopening the tracker links it again.
      t.handle = null;
      return finish(XLSX.write(work, { bookType: "xlsx", type: "array" }), ".xlsx");
    }
    if (!base) return Promise.resolve(legacy(SG ? "file bytes unavailable" : null));
    var ext = /\.xlsm$/i.test(t.name) ? ".xlsm" : ".xlsx";
    // Only the engine's own failures fall back; an error after a successful save is reported, not "fixed" lossy.
    return Promise.resolve().then(function () { return SG.surgicalSave(bytes, base, work); }).then(function (res) {
      if (!res || !res.ok) return { fallback: (res && res.reason) || "unsupported file" };
      if (res.noChanges) return { same: true };
      return Promise.resolve(SG.verifySurgical(bytes, res.bytes, work)).then(function (why) { return why ? { fallback: why } : { out: res.bytes }; });
    }, function (e) { return { fallback: e && e.message || String(e) }; }).then(function (r) {
      if (r.fallback) return legacy(r.fallback);
      return finish(r.same ? bytes : r.out, r.same ? null : ext);
    }).catch(function (e) {
      if (window.console) console.error("[trainers] write-back", e);
      window.alert("Write to tracker stopped: " + (e && e.message || e));
      render();
    });
  }
  function confirmWriteback() {
    var t = S.tracker;
    if (!t) return;
    var pend = pendingScores(), added = t.addedRows;
    if (!pend.length && !added.length) return;
    var lines = [];
    added.forEach(function (a) {
      var s = a.session;
      lines.push("New row " + excelRow(s) + ": " + s.dateISO + " · " + s.trainer + " · " + s.code + " · " + s.client + " (" + s.id + ")");
    });
    pend.forEach(function (x) {
      lines.push("Row " + excelRow(x) + " (" + (x.dateISO || "—") + " · " + x.trainer + " · " + x.client + "): NPS = " +
                 num(x.computed.nps, 1) + ", Trainer Score = " + num(x.computed.trainerScore, 2));
    });
    var shown = lines.slice(0, 14);
    if (lines.length > 14) shown.push("…and " + (lines.length - 14) + " more");
    var what = [];
    if (added.length) what.push(plural(added.length, "new session"));
    if (pend.length) what.push(plural(pend.length, "session score"));
    var msg = "Write " + what.join(" and ") + " to " + t.name + "?\n\n" + shown.join("\n") + "\n\n" +
              (t.handle ? "Saves into the linked file (a copy is downloaded instead if it changed since you opened it)." : "Downloads an updated copy (-scored).");
    if (window.confirm(msg)) return writeback();
  }

  /* ---------------- results export ---------------- */
  function sessionView(s) {
    var c = s.computed || s.written || {};
    return { s: s, c: c, nps: c.nps != null ? c.nps : s.nps, ts: c.trainerScore != null ? c.trainerScore : s.trainerScore, n: c.n != null ? c.n : null };
  }
  function exportResults() {
    var t = S.tracker;
    if (!t) return;
    var rows = [["Date", "Trainer", "Course code", "Session title", "School", "Responses",
                 "NPS", "Trainer score", "Promoters", "Passives", "Detractors"]];
    t.sessions.forEach(function (s) {
      var v = sessionView(s), c = v.c;
      if (v.nps == null && v.ts == null) return;
      rows.push([s.dateISO, s.trainer, s.code, s.title, s.client, v.n != null ? v.n : "",
                 v.nps, v.ts, c.promoters != null ? c.promoters : "",
                 c.passives != null ? c.passives : "", c.detractors != null ? c.detractors : ""]);
    });
    if (rows.length < 2) { window.alert("No scored sessions to export yet."); return; }
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Trainer results");
    download(XLSX.write(wb, { bookType: "xlsx", type: "array" }), "trainer-results.xlsx");
  }

  /* ---------------- leaderboard ---------------- */
  function leaderboard() {
    var t = S.tracker;
    if (!t) return [];
    var agg = {};
    t.sessions.forEach(function (s) {
      var v = sessionView(s);
      if (v.nps == null && v.ts == null) return;
      var k = P().normKey(s.trainer);
      var a = agg[k] || (agg[k] = { trainer: s.trainer, n: 0, npsSum: 0, npsN: 0, tsSum: 0, tsN: 0, last: "" });
      a.n++;
      if (v.nps != null) { a.npsSum += v.nps; a.npsN++; }
      if (v.ts != null) { a.tsSum += v.ts; a.tsN++; }
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

  /* ---------------- add session ---------------- */
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
        if (name && !map[name]) map[name] = 1;
      });
      return Object.keys(map).sort();
    } catch (e) { return null; }
  }
  function schoolOptions() {
    var roster = tenantSchoolRoster(), names = [], seen = {};
    var add = function (nm) { if (nm && !seen[nm.toLowerCase()]) { seen[nm.toLowerCase()] = true; names.push(nm); } };
    if (roster && roster.length) roster.forEach(add);
    S.tracker.sessions.forEach(function (s) { add(s.client); });
    return { html: names.sort().map(function (nm) { return '<option value="' + esc(nm) + '">'; }).join(""),
             src: (roster && roster.length) ? "ALL TENANTS tracker" : "training tracker" };
  }
  function newSessionFormHTML() {
    var t = S.tracker, ev = S.manualForm ? null : S.eval, hint = ev && ev.hint;
    var date = "", dateSrc = "";
    if (ev && ev.responseDate) { date = ev.responseDate; dateSrc = "from response timestamps"; }
    else if (hint && hint.dateISO) { date = hint.dateISO; dateSrc = "from filename"; }
    var school = hint && hint.school ? (t.acronymIndex[hint.school.toUpperCase()] || hint.school) : "";
    var t0 = hint && hint.session === "PM" ? "13:00" : "08:00";
    var t1 = hint && hint.session === "PM" ? "17:00" : "12:00";
    var opt = function (list) { return list.map(function (x) { return '<option value="' + esc(x) + '">'; }).join(""); };
    var cm = t.courseMap || { byCode: {}, byTitle: {} };
    var sch = schoolOptions();
    var fld = function (label, inner, help) {
      return '<div class="col-sm-6 col-lg-4"><label class="form-label small fw-semibold mb-1">' + label + '</label>' + inner +
             (help ? '<div class="form-text">' + help + '</div>' : '') + '</div>';
    };
    return '<div class="card-box tr-newform mt-3"><h6 class="mb-1">' + (S.manualForm ? "Add a session" : "Add this eval as a new session") + '</h6>' +
      '<div class="small text-muted mb-2">Added as a new row at the end of “' + esc(t.sheetName) + '” when you write to the tracker. * required</div>' +
      '<div class="row g-2">' +
      fld("Date *", '<input type="date" class="form-control form-control-sm" id="trns-date" value="' + esc(date) + '">', dateSrc ? '<span class="tr-hint">' + dateSrc + '</span>' : "") +
      fld("Trainer *", '<input class="form-control form-control-sm" id="trns-trainer" list="trns-trainers" placeholder="Type or pick" autocomplete="off"><datalist id="trns-trainers">' + opt(t.roster) + '</datalist>') +
      fld("Course name", '<input class="form-control form-control-sm" id="trns-title" list="trns-titles" placeholder="Start typing the course name" autocomplete="off"><datalist id="trns-titles">' + opt(Object.keys(cm.byTitle).sort()) + '</datalist>') +
      fld("Course code *", '<input class="form-control form-control-sm" id="trns-code" list="trns-codes" placeholder="Filled from the course name" autocomplete="off"><datalist id="trns-codes">' + opt(Object.keys(cm.byCode).sort()) + '</datalist>', '<span class="tr-hint" id="trns-codehint"></span>') +
      fld("School *", '<input class="form-control form-control-sm" id="trns-school" list="trns-schools" value="' + esc(school) + '" placeholder="Type or pick" autocomplete="off"><datalist id="trns-schools">' + sch.html + '</datalist>', '<span class="tr-hint">suggestions from ' + sch.src + '</span>') +
      fld("Time", '<div class="d-flex gap-2"><input type="time" class="form-control form-control-sm" id="trns-t0" value="' + t0 + '" aria-label="Time start"><input type="time" class="form-control form-control-sm" id="trns-t1" value="' + t1 + '" aria-label="Time end"></div>') +
      '</div><div class="d-flex gap-2 mt-3">' +
      '<button type="button" class="btn btn-primary btn-sm" data-traction="save-new-session">Save session</button>' +
      '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="' + (S.manualForm ? "cancel-manual" : "cancel-new-session") + '">Cancel</button></div></div>';
  }
  function saveNewSession() {
    var host = $("trainers-body"), t = S.tracker, ev = S.eval;
    if (!t) return;
    var g = function (id) { var el = host.querySelector("#trns-" + id); return el ? el.value.trim() : ""; };
    var date = g("date"), trainer = g("trainer"), code = g("code"), client = g("school");
    if (!date || !trainer || !code || !client) { alert("Date, trainer, course code and school are required."); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { alert("Pick a valid date."); return; }
    var ts = timeToFrac(g("t0")), te = timeToFrac(g("t1"));
    if (ts != null && te != null && te <= ts) { alert("The end time must be after the start time."); return; }
    var title = g("title"), cm = (t.courseMap && t.courseMap.byCode[code]) || null;
    // A partly typed course name that auto-filled the code is stored as the full course title.
    if (cm && cm.title && (!title || cm.title.toLowerCase().indexOf(title.toLowerCase()) === 0)) title = cm.title;
    var year = date.slice(0, 4);
    var sid = P().nextSessionId(t.sessions, year);
    // A Session ID column is added to the header (written with the rows).
    if (findHeader(t.headerRow, P().TRACKER_COLS.id) < 0) { t.headerRow.push("Session ID"); t.headerAdds.push("Session ID"); }
    var d = new Date(+year, +date.slice(5, 7) - 1, +date.slice(8, 10));
    var row = new Array(t.headerRow.length).fill("");
    var set = function (names, val) { var i = findHeader(t.headerRow, names); if (i >= 0) row[i] = val; };
    set(P().TRACKER_COLS.date, d); // a real Excel date, not text
    set(["year"], +year);
    set(["month"], MONTHS[d.getMonth()]);
    set(P().TRACKER_COLS.trainer, trainer);
    set(P().TRACKER_COLS.code, code);
    set(P().TRACKER_COLS.title, title);
    set(P().TRACKER_COLS.client, client);
    if (ts != null) set(P().TRACKER_COLS.timeStart, ts);
    if (te != null) set(P().TRACKER_COLS.timeEnd, te);
    if (ts != null && te != null) set(["duration"], Math.round((te - ts) * 24 * 100) / 100);
    set(P().TRACKER_COLS.id, sid);
    var rowIdx = t.lastRow + 1 + t.addedRows.length;
    var session = {
      rowIdx: rowIdx, id: sid, date: d, dateISO: date,
      trainer: trainer, code: code, title: title, client: client,
      timeStart: ts, timeEnd: te, nps: null, trainerScore: null,
      status: "unscored", isNew: true
    };
    t.addedRows.push({ rowIdx: rowIdx, cells: row, session: session });
    t.sessions.push(session);
    if (t.roster.indexOf(trainer) < 0) { t.roster.push(trainer); t.roster.sort(); }
    S.selected = t.sessions.length - 1;
    if (ev && !S.manualForm) attachTo(S.selected);
    S.showingNewForm = false; S.manualForm = false;
    render();
  }
  function removeNewSession(idx) {
    var t = S.tracker, s = t && t.sessions[idx];
    if (!s || !s.isNew) return;
    var k = t.addedRows.findIndex(function (a) { return a.session === s; });
    if (k < 0) return;
    t.addedRows.splice(k, 1);
    t.sessions.splice(idx, 1);
    // Later new rows move up one row; Session IDs stay as issued.
    t.addedRows.forEach(function (a, i) { a.rowIdx = a.session.rowIdx = t.lastRow + 1 + i; });
    if (!t.addedRows.length) {
      t.headerAdds.forEach(function (h) { var c = t.headerRow.indexOf(h); if (c >= 0) t.headerRow.splice(c, 1); });
      t.headerAdds = [];
    }
    if (S.eval && S.eval.attachedTo === idx) S.eval.attachedTo = -1;
    else if (S.eval && S.eval.attachedTo > idx) S.eval.attachedTo--;
    S.selected = -1;
    render();
  }

  /* ---------------- summary data ---------------- */
  function summaryData() {
    var t = S.tracker;
    var list = t.sessions.filter(function (s) {
      if (S.repTrainer && s.trainer !== S.repTrainer) return false;
      if (S.repSchool && s.client !== S.repSchool) return false;
      return true;
    }).map(function (s) {
      // Scores generated here but not written yet count too (they are shown as such).
      var v = sessionView(s);
      return { client: s.client, code: s.code, dateISO: s.dateISO, timeStart: s.timeStart, timeEnd: s.timeEnd,
               trainer: s.trainer, nps: v.nps, trainerScore: v.ts, n: v.n };
    });
    return P().buildSummary(list);
  }

  /* ---------------- charts ---------------- */
  function destroyCharts() {
    Object.keys(S.charts).forEach(function (k) { try { S.charts[k].destroy(); } catch (e) {} });
    S.charts = {}; S.chartDefs = {};
  }
  function tones() {
    return { ok: cssVar("--ok-fill", "#107C10"), warn: cssVar("--warn-fill", "#F7630C"), bad: cssVar("--bad-fill", "#D13438"),
             brand: cssVar("--brand", "#0F6CBD"), muted: cssVar("--line-strong", "#D1D1D1"), teal: (QBR.PALETTE || [])[5] || "#038387" };
  }
  // defs: canvas id -> function(tones) returning a Chart.js config (rebuilt for print in light colors)
  function drawCharts() {
    if (typeof Chart === "undefined") return;
    Object.keys(S.chartDefs).forEach(function (id) {
      var cv = $(id); if (!cv) return;
      var cfg = S.chartDefs[id](tones());
      cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false }, cfg.options || {});
      try { S.charts[id] = new Chart(cv, cfg); } catch (e) { if (window.console) console.warn("[trainers] chart", id, e); }
    });
  }
  function npsTone(v) { return v == null ? "blue" : v >= 50 ? "green" : v >= 0 ? "orange" : "red"; }
  function scoreTone(v) { return v == null ? "blue" : v >= 4.5 ? "green" : v >= 3.5 ? "orange" : "red"; }
  function kpi(label, val, tone, sub, title) {
    return '<div class="col"><div class="kpi kpi-' + tone + '"' + (title ? ' title="' + esc(title) + '"' : "") + '>' +
           '<div class="kpi-val">' + val + '</div><div class="kpi-lbl">' + esc(label) + '</div>' +
           (sub ? '<div class="tr-kpi-sub">' + sub + '</div>' : "") + '</div></div>';
  }
  function deltaChip(cur, prev, d) {
    if (cur == null || prev == null) return "";
    var x = Math.round((cur - prev) * Math.pow(10, d)) / Math.pow(10, d);
    if (!x) return '<span class="kpi-delta">± 0 vs previous</span>';
    return '<span class="kpi-delta ' + (x > 0 ? "up" : "down") + '">' + (x > 0 ? "▲ +" : "▼ ") + x + ' vs previous</span>';
  }

  /* ---------------- rendering: shell ---------------- */
  function statusBadge(s) {
    if (s.isNew && s.status !== "ready") return '<span class="risk-badge risk-low">New · not written</span>';
    if (s.status === "scored") return '<span class="risk-badge risk-none">Scored</span>';
    if (s.status === "ready") return '<span class="risk-badge risk-high">Needs verifying</span>';
    return '<span class="risk-badge tr-unscored">Unscored</span>';
  }
  function loadCardHTML() {
    var t = S.tracker, ev = S.eval;
    var pend = pendingScores().length, added = t ? t.addedRows.length : 0;
    var h = '<div class="card-box tr-load">' +
      '<div class="tr-load-row">' +
      '<button type="button" class="btn btn-' + (t ? "outline-primary" : "primary") + ' btn-sm" data-trpill="tracker">' + (t ? "Change tracker…" : "Open training tracker…") + '</button>' +
      '<button type="button" class="btn btn-' + (t && !ev ? "primary" : "outline-primary") + ' btn-sm" data-trpill="eval">' + (ev ? "Add another eval…" : "Add eval export (.xlsx)…") + '</button>' +
      '<span class="tr-files">' +
      '<span class="tr-file' + (t ? " on" : "") + '" data-trfile="tracker"><b>Tracker</b> ' + (t ? esc(t.name) + ' <span class="text-muted">· ' + esc(t.sheetName) + ' · ' + plural(t.sessions.length, "session") + ' · ' + plural(t.roster.length, "trainer") + '</span>' + (t.handle ? ' <span class="badge bg-success" title="Write-back saves into this file">saves in place</span>' : "") : '<span class="text-muted">not loaded</span>') + '</span>' +
      '<span class="tr-file' + (ev ? " on" : "") + '" data-trfile="eval"><b>Eval</b> ' + (ev ? esc(ev.name) + ' <span class="text-muted">· ' + plural(ev.rows.length - 1, "response") + '</span>' : '<span class="text-muted">none</span>') + '</span>' +
      '</span></div>';
    if (pend || added) {
      h += '<div class="tr-pending"><span><b>Not written yet:</b> ' +
        [added ? plural(added, "new session") : "", pend ? plural(pend, "scored session") : ""].filter(Boolean).join(" · ") + '</span>' +
        '<button type="button" class="btn btn-primary btn-sm" data-traction="writeback">Write to tracker…</button></div>';
    }
    return h + '<div id="tr-status" class="small text-muted" role="status" aria-live="polite"></div></div>';
  }
  function viewBarHTML() {
    var b = function (k, l) { return '<button type="button" class="btn btn-outline-secondary' + (S.view === k ? " active" : "") + '" data-trview="' + k + '" aria-pressed="' + (S.view === k) + '">' + l + '</button>'; };
    return '<div class="tr-viewbar"><div class="btn-group btn-group-sm" role="group" aria-label="Trainers view">' +
      b("sessions", "Sessions") + b("board", "Leaderboard") + b("summary", "Summary") + '</div>' +
      '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="export" title="Scored sessions as an Excel workbook">Export results (.xlsx)</button></div>';
  }

  /* ---------------- sessions view ---------------- */
  function filteredSessions() {
    var t = S.tracker;
    if (!t) return [];
    var q = P().normKey(S.q);
    var out = [];
    t.sessions.forEach(function (s, i) {
      if (S.filter === "needs" && s.status !== "ready") return;
      if (S.filter === "scored" && s.status !== "scored") return;
      if (S.filter === "unscored" && s.status !== "unscored") return;
      if (q) {
        var hay = P().normKey(s.trainer + " " + s.code + " " + s.client + " " + s.title + " " + s.dateISO + " " + (s.id || ""));
        if (hay.indexOf(q) < 0) return;
      }
      out.push({ s: s, i: i });
    });
    return out;
  }
  function tableHTML() {
    var rows = filteredSessions(), shown = rows.slice(0, 400);
    var h = '<div class="tr-count small text-muted">' + plural(rows.length, "session") + (rows.length > 400 ? " (showing 400 — refine the search)" : "") + '</div>' +
      '<div class="table-responsive tr-tablewrap"><table class="table table-sm table-hover tr-table" data-resizable="static"><thead><tr><th>Date</th><th>Trainer</th><th>Course</th><th>School</th>' +
      '<th class="num">Responses</th><th class="num">NPS</th><th class="num">Trainer score</th><th>Status</th></tr></thead><tbody>';
    shown.forEach(function (x) {
      var s = x.s, v = sessionView(s);
      h += '<tr data-trsession="' + x.i + '" tabindex="0" class="' + (x.i === S.selected ? "sel " : "") + "st-" + s.status + '">' +
        '<td class="text-nowrap">' + esc(s.dateISO || "—") + '</td><td>' + esc(s.trainer) + '</td>' +
        '<td title="' + esc(s.title) + '">' + esc(s.code) + '</td><td>' + esc(s.client) + '</td>' +
        '<td class="num">' + (v.n != null ? v.n : "—") + '</td>' +
        '<td class="num">' + (v.nps != null ? num(v.nps, 1) : "—") + '</td>' +
        '<td class="num">' + (v.ts != null ? num(v.ts, 2) : "—") + '</td>' +
        '<td>' + statusBadge(s) + '</td></tr>';
    });
    if (!shown.length) h += '<tr class="tr-norow"><td colspan="8" class="text-muted">No sessions match.</td></tr>';
    return h + '</tbody></table></div>';
  }
  function sessionsHTML() {
    var t = S.tracker, all = P().buildSummary(t.sessions.map(function (s) { var v = sessionView(s); return { client: s.client, nps: v.nps, trainerScore: v.ts, n: v.n }; }));
    var h = '<div class="row row-cols-2 row-cols-md-3 row-cols-xl-5 g-3 kpi-row mt-1">' +
      kpi("Sessions", all.total, "blue") +
      kpi("Respondents", all.respondents || "—", "blue", all.respSessions ? plural(all.respSessions, "session") + " with eval results" : "from eval results added here") +
      kpi("Avg NPS", all.avgNps == null ? "—" : num(all.avgNps, 1), npsTone(all.avgNps), plural(all.scoredNps, "scored session")) +
      kpi("Avg trainer score", all.avgScore == null ? "—" : num(all.avgScore, 2), scoreTone(all.avgScore), "out of 5 · " + plural(all.scoredTs, "session")) +
      kpi("Repeat schools", all.repeatRate + "%", "teal", "sessions at schools trained more than once") +
      '</div>';
    var f = function (k, l) { return '<button type="button" class="btn btn-outline-secondary' + (S.filter === k ? " active" : "") + '" data-trfilter="' + k + '" aria-pressed="' + (S.filter === k) + '">' + l + '</button>'; };
    h += '<div class="card-box mt-3"><div class="tr-ch-head"><h6 class="mb-0">Sessions</h6>' +
      '<div class="tr-toolbar"><div class="btn-group btn-group-sm" role="group" aria-label="Status filter">' +
      f("all", "All") + f("needs", "Needs verifying") + f("scored", "Scored") + f("unscored", "Unscored") + '</div>' +
      '<input type="search" class="form-control form-control-sm tr-search" id="tr-search" placeholder="Search trainer, school, course, date…" aria-label="Search sessions" value="' + esc(S.q) + '">' +
      '<button type="button" class="btn btn-outline-primary btn-sm" data-traction="add-session">+ Add session</button></div></div>' +
      '<div id="tr-list">' + tableHTML() + '</div></div>';
    if (S.manualForm) h += newSessionFormHTML();
    return h + detailHTML();
  }

  /* ---------------- session detail + feedback report ---------------- */
  function matchHTML(ev) {
    var h = '<div class="card-box mt-3 tr-detail"><h6 class="mb-1">Match this eval to a session</h6>' +
      '<div class="small text-muted mb-2">' + esc(ev.name) + ' · ' + plural(ev.rows.length - 1, "response") +
      (ev.hint ? ' · filename says ' + esc(ev.hint.dateISO) + ' · ' + esc(ev.hint.session) + ' · ' + esc(ev.hint.school) : "") +
      (ev.responseDate ? ' · answered on ' + esc(ev.responseDate) : "") + '</div>';
    if (ev.suggested.length) {
      h += '<div class="tr-suggest">' + ev.suggested.slice(0, 5).map(function (s) {
        var idx = S.tracker.sessions.indexOf(s);
        return '<button type="button" class="btn btn-outline-primary btn-sm text-start" data-traction="attach" data-idx="' + idx + '">' +
               esc(s.dateISO || "") + ' · ' + esc(s.client) + ' · ' + esc(s.trainer) + ' (' + esc(s.code) + ')</button>';
      }).join("") + '</div>';
    } else {
      h += '<div class="small text-muted">No session matched the filename — click a session in the list to attach this file to it.</div>';
    }
    if (S.showingNewForm) h += newSessionFormHTML();
    else h += '<div class="mt-2"><button type="button" class="btn btn-outline-secondary btn-sm" data-traction="new-session">Add as a new session</button> ' +
              '<span class="tr-hint">logs it in the tracker, then scores it like any other session</span></div>';
    return h + '</div>';
  }
  function detailHTML() {
    var ev = S.eval;
    if (ev && ev.attachedTo < 0 && S.tracker) return matchHTML(ev);
    if (S.selected < 0 || !S.tracker) return "";
    var s = S.tracker.sessions[S.selected];
    if (!s) return "";
    var c = s.computed || s.written;
    var h = '<div class="card-box mt-3 tr-detail" id="tr-detail"><div class="tr-ch-head"><div><h6 class="mb-0">' + esc(s.title || s.code) + '</h6>' +
      '<div class="small text-muted">' + esc(s.dateISO || "") + ' · ' + esc(s.client) + ' · ' + esc(s.code) +
      (s.id ? ' · ' + esc(s.id) : "") + ' · Trainer: <b>' + esc(s.trainer) + '</b> · row ' + excelRow(s) + '</div></div>' +
      (s.isNew ? '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="remove-new" title="Remove this session before it is written">Remove new session</button>' : "") + '</div>';
    if (!ev || ev.attachedTo !== S.selected) {
      if (c && c.report) return h + reportHTML(s, c) + '</div>';
      return h + '<div class="tr-empty small text-muted">No eval attached. ' +
        '<button type="button" class="btn btn-primary btn-sm" data-traction="upload-eval">Add eval for this session…</button></div></div>';
    }
    if (ev.needsMap) return h + mapperHTML() + '</div>';
    var hint = ev.hint, mismatch = false;
    if (hint && hint.dateISO && s.dateISO && hint.dateISO !== s.dateISO) mismatch = true;
    if (hint && hint.school && s.client && P().normKey(s.client).indexOf(P().normKey(hint.school)) < 0 &&
        P().normKey(hint.school).indexOf(P().normKey(s.client)) < 0 &&
        P().normKey(S.tracker.acronymIndex[hint.school.toUpperCase()] || "") !== P().normKey(s.client)) mismatch = true;
    h += '<div class="tr-evalline small">Eval: <b>' + esc(ev.name) + '</b> · ' + plural(ev.rows.length - 1, "response") +
         (hint ? ' <span class="tr-hint">· filename: ' + esc(hint.dateISO) + ' · ' + esc(hint.session) + ' · ' + esc(hint.school) + '</span>' : "") +
         (mismatch ? ' <span class="risk-badge risk-high">Filename suggests a different session — double-check</span>' : "") + '</div>';
    if (!c) h += '<button type="button" class="btn btn-primary btn-sm mt-2" data-traction="generate">Generate results</button>';
    else h += reportHTML(s, c);
    return h + '</div>';
  }
  function soloHTML() {
    var ev = S.eval;
    var h = '<div class="card-box mt-3 tr-detail" id="tr-detail"><div class="tr-ch-head"><div><h6 class="mb-0">Session feedback report</h6>' +
      '<div class="small text-muted">' + esc(ev.name) + ' · ' + plural(ev.rows.length - 1, "response") +
      (ev.responseDate ? ' · answered on ' + esc(ev.responseDate) : "") + ' · no tracker loaded (nothing is written)</div></div></div>';
    if (ev.needsMap) return h + mapperHTML() + '</div>';
    if (!ev.solo) return h + '<button type="button" class="btn btn-primary btn-sm mt-2" data-traction="generate">Generate report</button></div>';
    return h + reportHTML(null, ev.solo) + '</div>';
  }
  function chip(t, tone) { return '<span class="risk-badge risk-' + tone + ' tr-chip">' + esc(t) + '</span>'; }
  function reportHTML(s, c) {
    var R = c.report, lowN = c.n < LOW_N;
    var title = s ? (s.title || s.code) : "Session feedback";
    var h = '<div class="tr-report" id="tr-report-session" data-title="' + esc("Session feedback - " + (s ? (s.client + " - " + (s.dateISO || "")) : c.evalName)) + '">' +
      '<div class="tr-print-only tr-print-head"><div class="tr-ph-title">Participant feedback summary</div>' +
      '<div class="tr-ph-sub">' + esc(title) + (s ? ' · ' + esc(s.client) + ' · ' + esc(s.dateISO || "") + ' · Trainer: ' + esc(s.trainer) : "") +
      ' · ' + plural(R.responses, "response") + '</div></div>' +
      '<div class="tr-rep-actions tr-noprint"><button type="button" class="btn btn-primary btn-sm" data-traction="print-session" title="Opens the print dialog — choose “Save as PDF”">Session report (PDF)</button>' +
      (s && !s.computed && s.written ? ' <span class="small text-muted">Written to the tracker.</span>' +
        (S.eval && S.eval.attachedTo === S.selected ? ' <button type="button" class="btn btn-outline-secondary btn-sm" data-traction="generate">Generate again</button>' : "") : "") + '</div>' +
      '<div class="row row-cols-2 row-cols-md-3 row-cols-xl-5 g-3 kpi-row">' +
      kpi("Session NPS", num(c.nps, 1), npsTone(c.nps), deltaChip(c.nps, c.prevNps, 1) || (c.n ? c.promoters + " promoters · " + c.detractors + " detractors" : "no valid answers")) +
      kpi("Trainer score", num(c.trainerScore, 2), scoreTone(c.trainerScore), deltaChip(c.trainerScore, c.prevScore, 2) || "out of 5 · facilitator items") +
      kpi("Responses", R.responses, lowN ? "orange" : "blue", lowN ? "low response count (under " + LOW_N + ") — read with care" : plural(c.n, "NPS answer")) +
      kpi("Overall rating", num(R.overall, 2), scoreTone(R.overall), R.ratings.length ? "average of " + plural(R.ratings.length, "rated item") : "no rating items found") +
      kpi("Satisfaction", R.satisfaction == null ? "—" : num(R.satisfaction, 0) + "%", R.satisfaction == null ? "blue" : R.satisfaction >= 85 ? "green" : R.satisfaction >= 65 ? "orange" : "red", "top-2 answers on the satisfaction item") +
      '</div>';
    if (c.prevLabel) h += '<div class="small text-muted mt-1">Previous session with this trainer and course: ' + esc(c.prevLabel) + '</div>';
    // charts
    h += '<div class="row g-3 mt-0">' +
      '<div class="col-lg-5"><div class="card-box tr-inner"><h6>Recommendation (NPS)</h6><div class="chart-wrap tr-chart-sm"><canvas id="tr-ch-nps" aria-label="Promoters, passives and detractors"></canvas></div>' +
      '<div class="small text-muted mt-1">Promoters 9–10 · passives 7–8 · detractors 0–6. NPS = % promoters − % detractors.</div></div></div>' +
      '<div class="col-lg-7"><div class="card-box tr-inner"><h6>Rating by item</h6>' +
      (R.ratings.length ? '<div class="chart-wrap" style="height:' + Math.max(200, 34 * R.ratings.length + 60) + 'px"><canvas id="tr-ch-items" aria-label="Average rating per item"></canvas></div>' : '<div class="small text-muted">No Likert rating items found in this form.</div>') +
      '</div></div></div>';
    // executive summary + recommendations
    var sm = R.summary, opening = "Feedback from " + plural(R.responses, "participant") +
      (R.schools.length ? " from " + R.schools.slice(0, 3).map(function (o) { return o.name; }).join(", ") + (R.schools.length > 3 ? " and " + plural(R.schools.length - 3, "other school") : "") : "");
    opening = opening.replace(/\.+$/, "") + ".";
    if (sm.verdict) opening += " Overall sentiment was " + sm.verdict + " (" + num(R.overall, 2) + " / 5" + (R.satisfaction != null ? ", " + num(R.satisfaction, 0) + "% satisfied" : "") + ").";
    if (c.nps != null) opening += " NPS " + num(c.nps, 1) + ": " + c.promoters + " of " + c.n + " respondents are promoters.";
    h += '<div class="row g-3 mt-0"><div class="col-lg-5"><div class="card-box tr-inner"><h6>Executive summary</h6><p class="mb-2 small">' + esc(opening) + '</p>' +
      '<div class="tr-chips">' + sm.wins.map(function (w) { return chip("✓ " + w, "none"); }).join("") + sm.watches.map(function (w) { return chip("⚠ " + w, "high"); }).join("") + '</div></div></div>' +
      '<div class="col-lg-7"><div class="card-box tr-inner"><h6>Recommendations</h6>' + (R.recommendations.length ? '<ol class="tr-recs">' + R.recommendations.map(function (r) {
        var tone = r.priority === "high" ? "critical" : r.priority === "medium" ? "high" : "low";
        return '<li><div class="tr-rec-h">' + chip(r.priority, tone) + ' <b>' + esc(r.title) + '</b> <span class="small text-muted">· ' + esc(r.effort) + ' · ' + esc(r.owner) + '</span></div>' +
          '<div class="small">' + esc(r.evidence) + '</div>' + (r.quote ? '<blockquote class="tr-quote">“' + esc(r.quote) + '”</blockquote>' : "") +
          '<div class="small"><b>Action:</b> ' + esc(r.action) + '</div><div class="small text-muted"><b>Measure:</b> ' + esc(r.metric) + '</div></li>';
      }).join("") + '</ol>' : '<div class="small text-muted">No issues stood out in this session.</div>') + '</div></div></div>';
    // schools + themes
    h += '<div class="row g-3 mt-0">';
    if (R.schools.length) h += '<div class="col-lg-6"><div class="card-box tr-inner"><h6>Participants by school</h6><table class="table table-sm mb-0" data-resizable="static"><thead><tr><th>School</th><th class="num">Responses</th><th class="num">Share</th></tr></thead><tbody>' +
      R.schools.slice(0, 12).map(function (o) { return '<tr><td>' + esc(o.name) + '</td><td class="num">' + o.count + '</td><td class="num">' + num(o.pct, 0) + '%</td></tr>'; }).join("") + '</tbody></table></div></div>';
    var kw = function (list) { return list.length ? list.map(function (k) { return '<span class="badge bg-secondary tr-kw">' + esc(k.word) + ' · ' + k.count + '</span>'; }).join(" ") : '<span class="small text-muted">—</span>'; };
    h += '<div class="col-lg-6"><div class="card-box tr-inner"><h6>Themes in open answers</h6><div class="small fw-semibold mb-1">Liked best</div><div class="mb-2">' + kw(R.keywords.best) + '</div>' +
      '<div class="small fw-semibold mb-1">Suggestions</div><div>' + kw(R.keywords.suggest) + '</div></div></div></div>';
    // comments
    var groups = [["best", "Liked best"], ["insight", "Key insight"], ["suggest", "Suggestions"], ["why", "Why this rating"]];
    var anyC = groups.some(function (g) { return R.comments[g[0]].length; });
    if (anyC) {
      h += '<div class="card-box tr-inner mt-3"><h6>Comments <span class="small text-muted fw-normal">(blank and “n/a” answers left out)</span></h6><div class="tr-comments">';
      groups.forEach(function (g) {
        var items = R.comments[g[0]];
        if (!items.length) return;
        h += '<div class="tr-comment"><div class="small fw-semibold mb-1">' + g[1] + ' (' + items.length + ')</div><ul>' +
             items.slice(0, 50).map(function (t) { return '<li>' + esc(t) + '</li>'; }).join("") +
             (items.length > 50 ? '<li><i>…and ' + (items.length - 50) + ' more</i></li>' : "") + '</ul></div>';
      });
      h += '</div></div>';
    }
    // write-back
    if (s && s.status === "ready") {
      h += '<div class="tr-writeback tr-noprint"><div class="small">Write back: NPS <b>' + num(c.nps, 1) + '</b> → “Net Promoter Score”, trainer score <b>' + num(c.trainerScore, 2) +
           '</b> → “Trainer Score” on row ' + excelRow(s) + ' of ' + esc(S.tracker.name) + '.</div>' +
           '<button type="button" class="btn btn-primary btn-sm" data-traction="writeback">Preview &amp; write to tracker…</button></div>';
    }
    h += '</div>';
    // chart definitions
    S.chartDefs["tr-ch-nps"] = function (T) {
      return { type: "doughnut", data: { labels: ["Promoters", "Passives", "Detractors"],
        datasets: [{ data: [c.promoters || 0, c.passives || 0, c.detractors || 0], backgroundColor: [T.ok, T.muted, T.bad], borderWidth: 1 }] },
        options: { cutout: "62%", plugins: { legend: { position: "right" } } } };
    };
    if (R.ratings.length) S.chartDefs["tr-ch-items"] = function (T) {
      return { type: "bar", data: { labels: R.ratings.map(function (r) { return r.header.length > 46 ? r.header.slice(0, 44) + "…" : r.header; }),
        datasets: [{ label: "Average (1–5)", data: R.ratings.map(function (r) { return r.avg; }),
          backgroundColor: R.ratings.map(function (r) { return r.avg >= 4.5 ? T.ok : r.avg >= 3.5 ? T.brand : T.warn; }) }] },
        options: { indexAxis: "y", scales: { x: { min: 1, max: 5, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false },
          tooltip: { callbacks: { afterLabel: function (ctx) { var r = R.ratings[ctx.dataIndex]; return r.top2 + "% top-2 · n " + r.n + " · " + r.group; } } } } } };
    };
    return h;
  }

  var ROLE_LABELS = [["nps", "NPS score (0–10)"], ["school", "School"], ["t1", "Trainer: knowledge"],
    ["t2", "Trainer: communication"], ["t3", "Trainer: engagement"], ["sat", "Satisfaction"],
    ["c1", "Comments: liked best"], ["c2", "Comments: key insight"], ["c3", "Comments: suggestions"], ["cnps", "Comments: why rating"]];
  function mapperHTML() {
    var ev = S.eval;
    var opts = ev.header.map(function (h, i) { return '<option value="' + i + '">' + esc(h.slice(0, 70)) + '</option>'; }).join("");
    var h = '<div class="tr-mapper"><h6 class="mb-1">New form layout — confirm the columns</h6><p class="small text-muted mb-2">This form layout hasn’t been seen before. Check each role (saved in this browser for next time).</p><div class="row g-2">';
    ROLE_LABELS.forEach(function (pair) {
      h += '<div class="col-md-6"><label class="tr-maprow"><span class="small fw-semibold">' + pair[1] + '</span><select class="form-select form-select-sm" data-trmap="' + pair[0] + '">' +
           '<option value="-1">— not in this form —</option>' + opts + '</select></label></div>';
    });
    return h + '</div><button type="button" class="btn btn-primary btn-sm mt-2" data-traction="save-map">Save mapping</button></div>';
  }

  /* ---------------- leaderboard view ---------------- */
  function boardHTML() {
    var rows = leaderboard();
    if (!rows.length) return '<div class="card-box mt-3 small text-muted">No scored sessions yet. Generate results for a session (or load a tracker that already has scores).</div>';
    S.chartDefs["tr-ch-board"] = function (T) {
      var top = rows.filter(function (r) { return r.avgScore != null; }).slice().sort(function (a, b) { return b.avgScore - a.avgScore; }).slice(0, 15);
      return { type: "bar", data: { labels: top.map(function (r) { return r.trainer; }),
        datasets: [{ label: "Avg trainer score", data: top.map(function (r) { return r.avgScore; }), backgroundColor: top.map(function (r) { return r.avgScore >= 4.5 ? T.ok : r.avgScore >= 3.5 ? T.brand : T.warn; }) }] },
        options: { indexAxis: "y", scales: { x: { min: 1, max: 5 } }, plugins: { legend: { display: false } } } };
    };
    var h = '<div class="row g-3 mt-1"><div class="col-lg-7"><div class="card-box"><h6>Trainer leaderboard</h6><div class="table-responsive"><table class="table table-sm tr-table tr-board" data-resizable="static"><thead><tr><th>#</th><th>Trainer</th><th class="num">Sessions</th>' +
            '<th class="num">Avg NPS</th><th class="num">Avg trainer score</th><th>Last session</th></tr></thead><tbody>';
    rows.forEach(function (r, i) {
      h += '<tr><td class="text-muted">' + (i + 1) + '</td><td>' + esc(r.trainer) + '</td><td class="num">' + r.sessions + '</td>' +
           '<td class="num">' + (r.avgNps != null ? num(r.avgNps, 1) : "—") + '</td>' +
           '<td class="num">' + (r.avgScore != null ? num(r.avgScore, 2) : "—") + '</td>' +
           '<td>' + esc(r.last || "—") + '</td></tr>';
    });
    return h + '</tbody></table></div><div class="small text-muted">Ranked by average NPS. Includes scores generated here that are not written yet.</div></div></div>' +
      '<div class="col-lg-5"><div class="card-box"><h6>Average trainer score</h6><div class="chart-wrap" style="height:' + Math.max(220, 26 * Math.min(15, rows.length) + 60) + 'px"><canvas id="tr-ch-board"></canvas></div></div></div></div>';
  }

  /* ---------------- summary view ---------------- */
  function summaryHTML() {
    var t = S.tracker, sum = summaryData();
    var schools = [];
    t.sessions.forEach(function (s) { if (s.client && schools.indexOf(s.client) < 0) schools.push(s.client); });
    schools.sort();
    var filt = [];
    if (S.repTrainer) filt.push("Trainer: " + S.repTrainer);
    if (S.repSchool) filt.push("School: " + S.repSchool);
    var stamp = new Date().toISOString().slice(0, 10);
    var h = '<div class="card-box mt-3 tr-noprint"><div class="tr-filters-row">' +
      '<label class="small fw-semibold">Trainer <select class="form-select form-select-sm" id="tr-rep-trainer"><option value="">All trainers</option>' +
      t.roster.slice().sort().map(function (r) { return '<option value="' + esc(r) + '"' + (S.repTrainer === r ? " selected" : "") + '>' + esc(r) + '</option>'; }).join("") + '</select></label>' +
      '<label class="small fw-semibold">School <select class="form-select form-select-sm" id="tr-rep-school"><option value="">All schools</option>' +
      schools.map(function (s) { return '<option value="' + esc(s) + '"' + (S.repSchool === s ? " selected" : "") + '>' + esc(s) + '</option>'; }).join("") + '</select></label>' +
      (filt.length ? '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="rep-clear">Clear filters</button>' : "") +
      '<button type="button" class="btn btn-primary btn-sm ms-auto" data-traction="print-report" title="Opens the print dialog — choose “Save as PDF”">Summary report (PDF)</button></div></div>';
    h += '<div class="tr-report" id="tr-report" data-title="' + esc("Training summary - " + stamp) + '">' +
      '<div class="tr-print-only tr-print-head"><div class="tr-ph-title">Training summary report</div>' +
      '<div class="tr-ph-sub">' + esc(t.name) + ' · ' + plural(sum.total, "session") + (filt.length ? ' · ' + esc(filt.join(" · ")) : "") + ' · generated ' + stamp + '</div></div>' +
      '<div class="row row-cols-2 row-cols-md-3 row-cols-xl-6 g-3 kpi-row mt-1">' +
      kpi("Sessions", sum.total, "blue") +
      kpi("Training hours", num(sum.hours, 1), "blue", "sessions with start and end times") +
      kpi("Respondents", sum.respondents || "—", "blue", sum.respSessions ? plural(sum.respSessions, "session") + " with eval results" : "add eval results to count") +
      kpi("Avg NPS", sum.avgNps == null ? "—" : num(sum.avgNps, 1), npsTone(sum.avgNps), plural(sum.scoredNps, "scored session")) +
      kpi("Avg trainer score", sum.avgScore == null ? "—" : num(sum.avgScore, 2), scoreTone(sum.avgScore), "out of 5 · " + plural(sum.scoredTs, "session")) +
      kpi("Repeat schools", sum.repeatRate + "%", "teal", "sessions at schools trained more than once") +
      '</div>';
    if (!sum.total) return h + '<div class="card-box mt-3 small text-muted">No sessions match these filters.</div></div>';
    var months = sum.monthly.filter(function (m) { return m.key !== "(no date)"; });
    h += '<div class="row g-3 mt-0"><div class="col-lg-7"><div class="card-box tr-inner"><h6>Sessions and NPS by month</h6><div class="chart-wrap"><canvas id="tr-ch-month" aria-label="Sessions and average NPS by month"></canvas></div></div></div>' +
      '<div class="col-lg-5"><div class="card-box tr-inner"><h6>Top schools by sessions</h6><div class="chart-wrap"><canvas id="tr-ch-school" aria-label="Sessions per school"></canvas></div></div></div></div>';
    S.chartDefs["tr-ch-month"] = function (T) {
      return { type: "bar", data: { labels: months.map(function (m) { var d = new Date(+m.key.slice(0, 4), +m.key.slice(5, 7) - 1, 1); return isNaN(d) ? m.key : d.toLocaleString("en-US", { month: "short", year: "2-digit" }); }),
        datasets: [
          { type: "line", label: "Avg NPS", data: months.map(function (m) { return m.avgNps; }), borderColor: T.ok, backgroundColor: T.ok, yAxisID: "y2", tension: 0.25, pointRadius: 3, spanGaps: true },
          { type: "bar", label: "Sessions", data: months.map(function (m) { return m.sessions; }), backgroundColor: T.brand + "55", borderColor: T.brand, borderWidth: 1, yAxisID: "y" }] },
        options: { scales: { y: { beginAtZero: true, title: { display: true, text: "Sessions" }, ticks: { precision: 0 } },
          y2: { position: "right", min: -100, max: 100, grid: { drawOnChartArea: false }, title: { display: true, text: "NPS" } } } } };
    };
    var topS = sum.bySchool.slice().sort(function (a, b) { return b.sessions - a.sessions; }).slice(0, 10);
    S.chartDefs["tr-ch-school"] = function (T) {
      return { type: "bar", data: { labels: topS.map(function (b) { return b.key.length > 32 ? b.key.slice(0, 30) + "…" : b.key; }),
        datasets: [{ label: "Sessions", data: topS.map(function (b) { return b.sessions; }), backgroundColor: T.brand }] },
        options: { indexAxis: "y", scales: { x: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { display: false } } } };
    };
    function breakTable(title, first, rows) {
      var b = '<div class="col-lg-6"><div class="card-box tr-inner"><h6>' + title + '</h6><div class="table-responsive"><table class="table table-sm mb-0 tr-break" data-resizable="static"><thead><tr>' +
              '<th>' + first + '</th><th class="num">Sessions</th><th class="num">Avg NPS</th><th class="num">Avg trainer score</th></tr></thead><tbody>';
      rows.forEach(function (r) {
        b += '<tr><td>' + esc(r.key) + '</td><td class="num">' + r.sessions + '</td>' +
             '<td class="num">' + (r.avgNps == null ? "—" : num(r.avgNps, 1)) + '</td>' +
             '<td class="num">' + (r.avgScore == null ? "—" : num(r.avgScore, 2)) + '</td></tr>';
      });
      return b + '</tbody></table></div></div></div>';
    }
    var bySessions = function (a, b) { return b.sessions - a.sessions || (a.key < b.key ? -1 : 1); };
    h += '<div class="row g-3 mt-0">' + breakTable("By school", "School", sum.bySchool.slice().sort(bySessions)) +
         breakTable("By course", "Course code", sum.byCourse.slice().sort(bySessions)) + '</div>';
    return h + '<div class="small text-muted mt-2">Averages include scores generated here that are not written to the tracker yet.</div></div>';
  }

  /* ---------------- assembly + events ---------------- */
  function headHTML() {
    return '<div class="tr-head"><div><h3 class="dash-h mb-1">Trainer evaluations</h3>' +
      '<div class="small text-muted">Training tracker sessions · MS Forms evaluation exports · NPS and trainer scores written back to the tracker</div></div>' +
      '<div class="tr-privacy small" title="Evaluation answers are kept in this browser tab only — nothing is uploaded.">🔒 Answers stay in this tab</div></div>';
  }
  function render() {
    init(); // idempotent; loads saved column maps
    var host = $("trainers-body");
    if (!host) return;
    destroyCharts();
    if (!on()) { host.innerHTML = ""; return; }
    var h = headHTML() + loadCardHTML();
    if (!S.tracker) {
      if (S.eval) h += soloHTML();
      else h += '<div class="card-box mt-3 tr-empty"><h6>Get started</h6><ol class="small mb-0">' +
        '<li><b>Open training tracker…</b> — the session log (e.g. ETG_Training_Tracker_2026.xlsx, “trial” sheet).</li>' +
        '<li>Pick a session and <b>add its eval export</b> (MS Forms .xlsx, one file per session) — or add the eval first and confirm the suggested session.</li>' +
        '<li><b>Generate results</b>, check them, then <b>Write to tracker</b>. Each session gets a feedback report you can save as PDF.</li></ol>' +
        '<div class="small text-muted mt-2">Only have an eval file? Add it on its own for a one-off session report.</div></div>';
    } else {
      h += viewBarHTML();
      if (S.view === "board") h += boardHTML();
      else if (S.view === "summary") h += summaryHTML();
      else h += sessionsHTML();
    }
    host.innerHTML = h;
    wire(host);
    drawCharts();
  }

  function selectRow(i) {
    S.selected = i; S.manualForm = false;
    if (S.eval && S.eval.attachedTo < 0 && S.tracker) attachTo(i); // "click a session to attach this file"
    render(); var d = $("tr-detail"); if (d && d.scrollIntoView) try { d.scrollIntoView({ block: "nearest" }); } catch (e) {} }
  function wireRows(root) {
    root.querySelectorAll("[data-trsession]").forEach(function (tr) {
      tr.addEventListener("click", function () { selectRow(+tr.getAttribute("data-trsession")); });
      tr.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); selectRow(+tr.getAttribute("data-trsession")); } });
    });
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
    wireRows(host);
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
          titleEl.value = m.title; titleEl.setAttribute("data-auto", "1");
        }
      });
      var findCodeForTitle = function (txt) {
        var key = P().norm(txt);
        if (!key) return null;
        if (cm.byTitle[key]) return cm.byTitle[key];
        var low = key.toLowerCase(), hits = [];
        Object.keys(cm.byTitle).forEach(function (k) { if (k.toLowerCase().indexOf(low) === 0) hits.push(cm.byTitle[k]); });
        return hits.length === 1 ? hits[0] : null;
      };
      titleEl.addEventListener("change", function () {
        var cc = findCodeForTitle(titleEl.value);
        if (cc && (!codeEl.value || codeEl.getAttribute("data-auto") === "1")) { codeEl.value = cc; codeEl.setAttribute("data-auto", "1"); showHint(); }
      });
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
    if (S.eval && S.eval.needsMap) {
      host.querySelectorAll("[data-trmap]").forEach(function (sel) {
        var cur = S.eval.roles[sel.getAttribute("data-trmap")];
        if (cur != null && cur >= 0) sel.value = String(cur);
      });
    }
  }
  // Search: redraw only the table (keeps focus, charts and the detail panel intact).
  function renderListOnly() {
    var box = $("tr-list");
    if (!box) { render(); return; }
    box.innerHTML = tableHTML();
    wireRows(box);
  }

  /* ---------------- file pickers ---------------- */
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
          "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] } }]
      });
      handle = picks[0];
    } catch (e) { return; }
    var perm = true;
    try {
      var q = await handle.queryPermission({ mode: "readwrite" });
      perm = q === "granted" || (await handle.requestPermission({ mode: "readwrite" })) === "granted";
    } catch (e) { perm = false; }
    var file;
    try { file = await handle.getFile(); }
    catch (e) { alert("Could not read the file: " + (e && e.message || e)); return; }
    // Without write permission the tracker still opens; write-back downloads a copy.
    openTracker(file, perm ? handle : null).catch(function (e) { alert("Could not open " + file.name + ": " + e.message); });
  }
  function pickFile(kind) {
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

  /* ---------------- print to PDF ----------------
   * Same pattern as MOS: a print-only copy of the report is appended to <body>
   * (body.tr-printing hides everything else). Charts are re-rendered offscreen
   * in the light palette and inserted as images, so they print on any theme. */
  function chartImages(defs) {
    var out = {};
    if (typeof Chart === "undefined") return out;
    var prevTheme = document.body.getAttribute("data-theme");
    var dark = prevTheme && prevTheme !== "light";
    if (dark) { document.body.setAttribute("data-theme", "light"); if (QBR.applyChartTheme) QBR.applyChartTheme(); }
    var box = document.createElement("div");
    box.style.cssText = "position:fixed;left:-10000px;top:0;width:1100px";
    document.body.appendChild(box);
    try {
      Object.keys(defs).forEach(function (id) {
        var live = $(id);
        // Same proportions as on screen, so the image fits its card on paper.
        var wrap = live && live.parentNode;
        var h = wrap ? Math.max(200, Math.min(700, wrap.clientHeight || 260)) : 260;
        var w = wrap ? Math.max(360, Math.min(1100, wrap.clientWidth || 600)) : 600;
        var cv = document.createElement("canvas");
        cv.width = w; cv.height = h; cv.style.width = w + "px"; cv.style.height = h + "px";
        box.appendChild(cv);
        var cfg = defs[id](tones());
        cfg.options = Object.assign({}, cfg.options || {}, { responsive: false, maintainAspectRatio: false, animation: false, devicePixelRatio: 2 });
        try { var ch = new Chart(cv, cfg); out[id] = ch.toBase64Image("image/png", 1); ch.destroy(); } catch (e) {}
      });
    } finally {
      box.remove();
      if (dark) { document.body.setAttribute("data-theme", prevTheme); if (QBR.applyChartTheme) QBR.applyChartTheme(); }
    }
    return out;
  }
  function printReport(srcId) {
    var src = $(srcId);
    if (!src) return false;
    var imgs = chartImages(S.chartDefs);
    var host = $("tr-print"); if (host) host.remove();
    host = document.createElement("div"); host.id = "tr-print";
    var doc = src.cloneNode(true);
    doc.removeAttribute("id");
    doc.querySelectorAll("[id]").forEach(function (el) { if (el.tagName !== "CANVAS") el.removeAttribute("id"); });
    doc.querySelectorAll(".tr-noprint").forEach(function (el) { el.remove(); });
    doc.querySelectorAll("canvas").forEach(function (cv) {
      var url = imgs[cv.id];
      if (!url) { var wrap = cv.closest(".chart-wrap"); if (wrap) wrap.remove(); else cv.remove(); return; }
      var im = document.createElement("img"); im.src = url; im.alt = cv.getAttribute("aria-label") || ""; im.className = "tr-print-chart";
      var w = cv.closest(".chart-wrap");
      if (w) { w.replaceWith(im); } else cv.replaceWith(im);
    });
    host.innerHTML = '<div class="tr-print-doc">' + doc.outerHTML + '<div class="tr-print-foot"><span>RCT OpsDesk · Trainer evaluations</span><span>Generated ' + esc(new Date().toLocaleString()) + '</span></div></div>';
    document.body.appendChild(host);
    document.body.classList.add("tr-printing");
    // Paper is light: print with the light theme tokens (restored afterwards).
    var prevTheme = document.body.getAttribute("data-theme");
    if (prevTheme && prevTheme !== "light") document.body.setAttribute("data-theme", "light");
    var prevTitle = document.title;
    document.title = (src.getAttribute("data-title") || "Trainer report").replace(/[\\/:*?"<>|]+/g, "-");
    var done = false;
    function cleanup() {
      if (done) return; done = true;
      document.body.classList.remove("tr-printing"); document.title = prevTitle;
      if (prevTheme && prevTheme !== "light") document.body.setAttribute("data-theme", prevTheme);
      var h = $("tr-print"); if (h) h.remove();
      window.removeEventListener("afterprint", cleanup);
    }
    window.addEventListener("afterprint", cleanup);
    var pics = Array.prototype.slice.call(host.querySelectorAll("img"));
    S.lastPrint = { charts: Object.keys(imgs).length, html: host.innerHTML.length };
    return Promise.race([Promise.all(pics.map(function (im) { return im.decode ? im.decode().catch(function () {}) : Promise.resolve(); })),
                         new Promise(function (r) { setTimeout(r, 3000); })])
      .then(function () { try { window.print(); } finally { setTimeout(cleanup, 60000); } });
  }

  /* ---------------- actions ---------------- */
  function action(name, el) {
    if (name === "export") { exportResults(); return; }
    if (name === "rep-clear") { S.repTrainer = ""; S.repSchool = ""; render(); return; }
    if (name === "print-report") return printReport("tr-report");
    if (name === "print-session") return printReport("tr-report-session");
    if (name === "upload-eval") { pickFile("eval"); return; }
    if (name === "attach") {
      var idx = el ? +el.getAttribute("data-idx") : -1;
      if (idx >= 0 && S.eval) { S.selected = idx; attachTo(idx); S.showingNewForm = false; }
      render();
      return;
    }
    if (name === "new-session") { S.showingNewForm = true; render(); return; }
    if (name === "cancel-new-session") { S.showingNewForm = false; render(); return; }
    if (name === "cancel-manual") { S.manualForm = false; render(); return; }
    if (name === "add-session") { S.manualForm = true; S.showingNewForm = false; render(); return; }
    if (name === "save-new-session") { saveNewSession(); return; }
    if (name === "remove-new") { removeNewSession(S.selected); return; }
    if (name === "generate") {
      var r = generateFor(S.selected);
      if (r && r.error) alert(r.error);
      render();
      return;
    }
    if (name === "save-map") {
      var host = $("trainers-body"), roles = {};
      host.querySelectorAll("[data-trmap]").forEach(function (sel) { roles[sel.getAttribute("data-trmap")] = +sel.value; });
      if (roles.nps == null || roles.nps < 0) { alert("The NPS (0–10) column is required."); return; }
      S.eval.roles = roles; S.eval.needsMap = false;
      S.maps[S.eval.sig] = { roles: roles, savedAt: new Date().toISOString() };
      saveMaps(); render();
      return;
    }
    if (name === "writeback") return confirmWriteback();
  }

  function init() {
    if (S.built) return;
    S.built = true;
    loadMaps();
  }

  QBR.trainers = { render: render, init: init, _state: S, _print: printReport };
})();
