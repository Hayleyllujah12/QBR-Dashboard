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
  var ALIAS_KEY = "qbr-trainers-aliases"; // filename school (e.g. "bpa") -> tracker client, learned from your picks
  var LOW_N = 10;

  var S = {
    built: false,
    tracker: null,   // { name, bytes, wb, sheetName, origin, headerRow, sessions[], roster[], addedRows[], headerAdds[], handle }
    eval: null,      // { name, rows, header, roles, sig, attachedTo, suggested[], hint, responseDate, needsMap, solo }
    selected: -1,    // index into sessions
    filter: "all",   // all | needs | scored | unscored
    showingNewForm: false, // "add as new session" from an unmatched eval
    manualForm: false,     // standalone Add session (no eval)
    repTrainer: "", repSchool: "", repPeriod: "",  // summary filters (period: "" | "2026" | "2026-Q2")
    q: "",
    view: "sessions",      // sessions | board | summary
    queue: [],             // bulk eval files: [{ ...eval, id, target: idx | -1, auto, cands[] }]
    pickFor: null,         // queue item waiting for a click in the session list
    mapFor: null,          // queue item whose column mapper is open
    newFor: null,          // queue item being added as a new session
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
    try { S.aliases = JSON.parse(localStorage.getItem(ALIAS_KEY) || "{}") || {}; } catch (e) { S.aliases = {}; }
    try { S.maps = JSON.parse(localStorage.getItem(MAP_KEY) || "{}") || {}; }
    catch (e) { S.maps = {}; }
  }
  function learnAlias(q, idx) {
    var s = S.tracker && S.tracker.sessions[idx], k = q && q.hint && q.hint.school ? P().normKey(q.hint.school).replace(/[^a-z]/g, "") : "";
    if (!s || !k || k.length < 2 || !s.client) return;
    S.aliases = S.aliases || {};
    S.aliases[k] = P().normKey(s.client);
    try { localStorage.setItem(ALIAS_KEY, JSON.stringify(S.aliases)); } catch (e) {}
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
      // Evals opened before the tracker (or with a previous tracker) are matched against this one.
      if (S.eval) { S.eval.solo = null; S.queue.push(S.eval); S.eval = null; }
      S.queue.forEach(function (q) { q.solo = null; matchItem(q); });
      dedupeTargets();
      S.pickFor = S.mapFor = S.newFor = null;
      render();
    });
  }

  /* ---------------- eval file ---------------- */
  var evalSeq = 0;
  function makeEval(file, wb) {
    var rows = sheetRows(wb, wb.SheetNames[0]);
    if (rows.length < 2) throw new Error("No responses in " + file.name);
    var det = P().detectLayout(rows[0]);
    var roles = (S.maps[det.sig] && S.maps[det.sig].roles) || det.roles;
    var hint = P().parseFilename(file.name);
    return {
      id: "ev" + (++evalSeq), name: file.name, rows: rows, header: rows[0].map(P().norm),
      roles: roles, sig: det.sig, auto: det.roles,
      attachedTo: -1, suggested: hint && S.tracker ? P().suggestSessions(S.tracker.sessions, hint) : [],
      hint: hint, responseDate: P().detectResponseDate(rows), respSession: P().detectResponseSession(rows),
      fp: P().evalFingerprint(rows), replace: false,
      needsMap: !S.maps[det.sig], solo: null, target: -1, cands: [], note: ""
    };
  }
  // "Add eval for this session" (session detail): one file, attached to the selected session.
  function openEval(file) {
    return readWorkbook(file).then(function (wb) {
      S.eval = makeEval(file, wb);
      S.showingNewForm = false;
      if (S.tracker && S.selected >= 0) attachTo(S.selected);
      render();
    });
  }
  /* "Add eval exports" / drop: one or many files go to the eval list and are matched to
   * sessions; nothing is attached to whatever session happens to be selected. */
  function matchItem(q) {
    q.cands = []; q.auto = false;
    if (!S.tracker) return;
    var m = P().matchEval(S.tracker.sessions, q.hint, q.responseDate, { session: q.respSession, aliases: S.aliases });
    q.cands = m.candidates.slice(0, 8).map(function (x) { return S.tracker.sessions.indexOf(x.s); });
    // Already encoded: the tracker's Eval Source carries this file's fingerprint -> that row, nothing to do.
    q.encodedAt = encodedRow(q);
    q.twin = -1;
    if (q.encodedAt >= 0) { q.target = q.encodedAt; q.auto = true; fillCount(q); return; }
    // v1.39.0: a row scored before "Eval Source" existed has no fingerprint. Same date + same NPS and
    // trainer score = this eval was already written there -> point at that row instead of a new one.
    q.twin = scoreTwin(q);
    if (q.twin >= 0 && !(q.target >= 0 && q.target !== q.twin && S.tracker.sessions[q.target])) {
      q.target = q.twin; q.auto = true;
      q.note = "Row " + excelRow(S.tracker.sessions[q.twin]) + " already has these exact scores (scored before Eval Source was recorded) — most likely from this file. Tick “Replace existing score” to record its Eval Source and Responses there; don't add it as a new session.";
      return;
    }
    if (q.target >= 0 && S.tracker.sessions[q.target]) return; // keep a choice the user made
    q.target = m.auto ? S.tracker.sessions.indexOf(m.auto) : -1;
    q.auto = !!m.auto;
  }
  function quickScore(q) {
    // auto-detected columns are enough here (encoded files never show the column mapper)
    if (!q || !q.roles || q.roles.nps == null || q.roles.nps < 0) return null;
    if (q._quick === undefined) { try { var c = computeFrom(q); q._quick = c.error ? null : c; } catch (e) { q._quick = null; } }
    return q._quick;
  }
  // Encoded row whose response count was never recorded: remember the count, written with the next write.
  function fillCount(q) {
    var s = S.tracker && q.encodedAt >= 0 ? S.tracker.sessions[q.encodedAt] : null;
    if (!s || s.n != null || s.computed) return;
    var c = quickScore(q);
    if (c && c.n != null) s.fillN = c.n;
  }
  function scoreTwin(q) {
    var c = quickScore(q);
    if (!c || c.nps == null) return -1;
    var days = {}; if (q.hint && q.hint.dateISO) days[q.hint.dateISO] = 1; if (q.responseDate) days[q.responseDate] = 1;
    var near = function (a, b, d) { return a == null && b == null || (a != null && b != null && Math.abs(a - b) < d); };
    var hits = [];
    S.tracker.sessions.forEach(function (s, i) {
      if (s.evalFp || s.nps == null || !days[s.dateISO]) return;
      if (near(s.nps, c.nps, 0.051) && near(s.trainerScore, c.trainerScore, 0.0051)) hits.push(i);
    });
    return hits.length === 1 ? hits[0] : -1;
  }
  function encodedRow(q) {
    if (!S.tracker || !q.fp) return -1;
    for (var i = 0; i < S.tracker.sessions.length; i++) if (S.tracker.sessions[i].evalFp === q.fp) return i;
    return -1;
  }
  // Scores already in the tracker for this session (not ones generated here).
  function existingScore(s) { return s && (s.nps != null || s.trainerScore != null) ? { nps: s.nps, ts: s.trainerScore, src: s.evalSource || "" } : null; }
  // Two files on one session: neither is pre-selected (the user picks).
  function dedupeTargets() {
    var seen = {};
    S.queue.forEach(function (q) { if (q.target >= 0) (seen[q.target] = seen[q.target] || []).push(q); });
    Object.keys(seen).forEach(function (k) {
      if (seen[k].length < 2) return;
      seen[k].forEach(function (q) { if (q.auto) { q.target = -1; q.auto = false; q.note = "Several files match the same session — pick one."; } });
    });
  }
  async function openEvals(files, opts) {
    opts = opts || {};
    files = Array.prototype.slice.call(files || []).filter(function (f) { return /\.(xlsx|xlsm|xls)$/i.test(f.name); });
    if (!files.length) return { added: 0, dup: 0, encoded: 0, bad: [] };
    if (!S.tracker && files.length === 1 && !S.queue.length && !opts.auto) { // eval on its own: one-off report
      var wb0 = await readWorkbook(files[0]);
      S.eval = makeEval(files[0], wb0); render(); return;
    }
    var bad = [], dup = 0, encoded = 0;
    for (var i = 0; i < files.length; i++) {
      note("Reading " + (i + 1) + " of " + files.length + ": " + files[i].name + "…");
      try {
        var wb = await readWorkbook(files[i]);
        var q = makeEval(files[i], wb);
        // Same file again (also renamed copies: same answers = same fingerprint).
        if (S.queue.some(function (x) { return (x.name === q.name && x.rows.length === q.rows.length) || (q.fp && x.fp === q.fp); })) { dup++; continue; }
        matchItem(q);
        // Folder scan: evals already encoded in the tracker are old news — leave them out.
        if (opts.auto && q.encodedAt >= 0) { encoded++; continue; }
        S.queue.push(q);
      } catch (e) { bad.push(files[i].name + " (" + (e && e.message || e) + ")"); }
    }
    dedupeTargets();
    render();
    var res = { added: files.length - bad.length - dup - encoded, dup: dup, encoded: encoded, bad: bad };
    if (opts.auto) return res;
    note(plural(res.added, "eval file") + " added" + (dup ? " · " + dup + " already in the list" : "") + (bad.length ? " · couldn't read: " + bad.join(", ") : "") + ".");
    if (bad.length) window.alert("Couldn't read:\n" + bad.join("\n"));
    return res;
  }
  function queueItem(id) { return S.queue.filter(function (q) { return q.id === id; })[0] || null; }
  function itemState(q) {
    var s = S.tracker && q.target >= 0 ? S.tracker.sessions[q.target] : null;
    if (s && s.computed && s.computed.evalId === q.id) return "generated";
    if (s && s.written && s.written.evalId === q.id && !s.computed) return "written";
    if (s && q.fp && s.evalFp === q.fp && !s.computed) return "encoded";
    return "";
  }
  /* Generate every matched, mapped file. Each result is checked before Write to tracker. */
  function generateQueue() {
    var t = S.tracker, done = 0, skipped = [];
    if (!t) return;
    var count = {};
    S.queue.forEach(function (q) { if (q.target >= 0) count[q.target] = (count[q.target] || 0) + 1; });
    S.queue.forEach(function (q) {
      if (q.target < 0 || itemState(q)) return;
      if (count[q.target] > 1) { q.note = "Another file is set to the same session."; skipped.push(q.name); return; }
      if (q.needsMap) { q.note = "Confirm the columns first."; skipped.push(q.name); return; }
      var at = encodedRow(q);
      if (at >= 0 && at !== q.target) { q.note = "This file is already encoded on row " + excelRow(t.sessions[at]) + " — not used for a second session."; skipped.push(q.name); return; }
      if (existingScore(t.sessions[q.target]) && !q.replace) { q.note = "Already scored — tick “Replace existing score” to overwrite it."; skipped.push(q.name); return; }
      var c = computeFrom(q);
      if (c.error) { q.note = c.error; skipped.push(q.name); return; }
      q.note = "";
      applyComputed(q.target, c);
      done++;
    });
    note(plural(done, "session") + " generated" + (skipped.length ? " · " + skipped.length + " need attention" : "") + ". Check them, then Write to tracker.");
    render();
  }

  // Attach the eval to a session. Results generated from a DIFFERENT eval file are dropped
  // (they would otherwise be shown and written under the new file's name).
  function attachTo(idx) {
    var s = S.tracker && S.tracker.sessions[idx];
    if (!s || !S.eval) return;
    S.eval.attachedTo = idx;
    if (s.computed && (s.computed.evalId ? s.computed.evalId !== S.eval.id : s.computed.evalName !== S.eval.name)) {
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
      report: rep, evalName: ev.name, evalId: ev.id, evalFp: ev.fp || "", evalSource: P().evalSourceText(ev.name, ev.fp)
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
    var at = encodedRow(ev), ses = t.sessions[idx];
    if (at >= 0 && at !== idx && !window.confirm("This eval file is already encoded on row " + excelRow(t.sessions[at]) + " (" + sessionLabel(t.sessions[at]) + ").\n\nUse it for this session too?")) return null;
    var ex = existingScore(ses);
    if (ex && !(at === idx) && !window.confirm("This session already has scores in the tracker (NPS " + num(ex.nps, 1) + ", Trainer Score " + num(ex.ts, 2) + (ex.src ? ", from " + ex.src : "") + ").\n\nGenerate new results from " + ev.name + "? They replace the old scores only when you write to the tracker.")) return null;
    var computed = computeFrom(ev);
    if (computed.error) return computed;
    return applyComputed(idx, computed);
  }
  function applyComputed(idx, computed) {
    var t = S.tracker, session = t.sessions[idx];
    var prev = P().findPrevSession(t.sessions, session);
    computed.prevNps = prev ? prev.nps : null;
    computed.prevScore = prev ? prev.trainerScore : null;
    computed.prevLabel = prev ? (prev.dateISO || "") + " · " + prev.client : "";
    session.computed = computed;
    // Empty result (no valid answers): show — and offer no write-back.
    session.status = (computed.nps != null || computed.trainerScore != null) ? "ready" : (session.nps != null ? "scored" : "unscored");
    return computed;
  }
  function pendingFills() {
    return S.tracker ? S.tracker.sessions.filter(function (x) { return x.fillN != null && x.n == null && !x.computed; }) : [];
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
    var srcCol = plan.sourceCol;
    if (plan.addSourceCol) { srcCol = next++; put(0, srcCol, "Eval Source"); }
    var respCol = plan.respCol;
    if (plan.addRespCol) { respCol = next++; put(0, respCol, "Responses"); }
    (t.addedRows || []).forEach(function (a) {
      a.cells.forEach(function (v, c) { if (v !== "" && v != null) put(a.rowIdx, c, v); });
    });
    plan.edits.forEach(function (e) {
      if (npsCol >= 0 && e.nps != null) put(e.row, npsCol, e.nps);
      if (tsCol >= 0 && e.trainerScore != null) put(e.row, tsCol, e.trainerScore);
      if (srcCol >= 0 && e.evalSource) put(e.row, srcCol, e.evalSource);
      if (respCol != null && respCol >= 0 && e.n != null) put(e.row, respCol, e.n);
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
        if (s.fillN != null) { if (s.n == null && !s.computed) s.n = s.fillN; delete s.fillN; }
        if (s.computed && s.status === "ready") {
          if (s.computed.nps != null) s.nps = s.computed.nps;
          if (s.computed.trainerScore != null) s.trainerScore = s.computed.trainerScore;
          if (s.computed.evalSource) { s.evalSource = s.computed.evalSource; s.evalFp = s.computed.evalFp; }
          if (s.computed.n != null) s.n = s.computed.n;
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
    var pend = pendingScores(), added = t.addedRows, fills = pendingFills();
    if (!pend.length && !added.length && !fills.length) return;
    var lines = [];
    added.forEach(function (a) {
      var s = a.session;
      lines.push("New row " + excelRow(s) + ": " + s.dateISO + " · " + s.trainer + " · " + s.code + " · " + s.client + " (" + s.id + ")");
    });
    pend.forEach(function (x) {
      var ex = existingScore(x);
      lines.push("Row " + excelRow(x) + " (" + (x.dateISO || "—") + " · " + x.trainer + " · " + x.client + "): " +
        (ex ? "REPLACES NPS " + num(ex.nps, 1) + " → " + num(x.computed.nps, 1) + ", Trainer Score " + num(ex.ts, 2) + " → " + num(x.computed.trainerScore, 2)
            : "NPS = " + num(x.computed.nps, 1) + ", Trainer Score = " + num(x.computed.trainerScore, 2)));
    });
    fills.forEach(function (x) { lines.push("Row " + excelRow(x) + " (" + (x.dateISO || "—") + " · " + x.client + "): Responses = " + x.fillN + " (count only — scores unchanged)"); });
    var shown = lines.slice(0, 14);
    if (lines.length > 14) shown.push("…and " + (lines.length - 14) + " more");
    var what = [];
    if (added.length) what.push(plural(added.length, "new session"));
    if (pend.length) what.push(plural(pend.length, "session score"));
    if (fills.length) what.push(plural(fills.length, "response count"));
    var msg = "Write " + what.join(" and ") + " to " + t.name + "?\n\n" + shown.join("\n") + "\n\n" +
              (t.handle ? "Saves into the linked file (a copy is downloaded instead if it changed since you opened it)." : "Downloads an updated copy (-scored).");
    if (window.confirm(msg)) return writeback();
  }

  /* ---------------- results export ---------------- */
  function sessionView(s) {
    var c = s.computed || s.written || {};
    return { s: s, c: c, nps: c.nps != null ? c.nps : s.nps, ts: c.trainerScore != null ? c.trainerScore : s.trainerScore, n: c.n != null ? c.n : s.n != null ? s.n : s.fillN != null ? s.fillN : null };
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
      P().splitTrainers(s.trainer).forEach(function (name) { // co-trainers share the session's scores
        var k = P().normKey(name);
        var a = agg[k] || (agg[k] = { trainer: name, n: 0, npsSum: 0, npsN: 0, tsSum: 0, tsN: 0, last: "" });
        a.n++;
        if (v.nps != null) { a.npsSum += v.nps; a.npsN++; }
        if (v.ts != null) { a.tsSum += v.ts; a.tsN++; }
        if (s.dateISO && s.dateISO > a.last) a.last = s.dateISO;
      });
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
    var t = S.tracker, ev = S.manualForm ? null : ((S.newFor && queueItem(S.newFor)) || S.eval), hint = ev && ev.hint;
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
      fld("Co-trainer", '<input class="form-control form-control-sm" id="trns-cotrainer" list="trns-trainers" placeholder="Optional — second speaker" autocomplete="off">', '<span class="tr-hint">one row for the session: saved as “TRAINER / CO-TRAINER”</span>') +
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
    var date = g("date"), trainer = P().joinTrainers([g("trainer"), g("cotrainer")]), code = g("code"), client = g("school");
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
    P().splitTrainers(trainer).forEach(function (n) { if (t.roster.indexOf(n) < 0) t.roster.push(n); });
    t.roster.sort();
    S.selected = t.sessions.length - 1;
    var nq = !S.manualForm && S.newFor && queueItem(S.newFor);
    if (nq) { nq.target = S.selected; nq.auto = false; S.newFor = null; }
    else if (ev && !S.manualForm) attachTo(S.selected);
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
  function summaryRows(period) {
    var t = S.tracker, cm = (t.courseMap && t.courseMap.byCode) || {};
    return t.sessions.filter(function (s) {
      if (S.repTrainer && !P().hasTrainer(s, S.repTrainer)) return false; // co-trained sessions count for each trainer
      if (S.repSchool && s.client !== S.repSchool) return false;
      return P().inPeriod(s.dateISO, period);
    }).map(function (s) {
      // Scores generated here but not written yet count too (the report says so).
      var v = sessionView(s), c = s.computed || s.written || {};
      return { client: s.client, code: s.code, title: s.title || (cm[s.code] && cm[s.code].title) || "", dateISO: s.dateISO,
               timeStart: s.timeStart, timeEnd: s.timeEnd, trainer: s.trainer, nps: v.nps, trainerScore: v.ts, n: v.n,
               itemAvgs: c.itemAvgs || null };
    });
  }
  function summaryData() { return P().buildSummary(summaryRows(S.repPeriod)); }

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
    var pend = pendingScores().length, added = t ? t.addedRows.length : 0, fills = pendingFills().length;
    var h = '<div class="card-box tr-load">' +
      '<div class="tr-load-row">' +
      '<button type="button" class="btn btn-' + (t ? "outline-primary" : "primary") + ' btn-sm" data-trpill="tracker">' + (t ? "Change tracker…" : "Open training tracker…") + '</button>' +
      '<button type="button" class="btn btn-' + (t && !S.queue.length ? "primary" : "outline-primary") + ' btn-sm" data-trpill="eval" title="Pick one or many MS Forms exports (Ctrl/Shift-click). Each file is one session.">Add eval exports (.xlsx)…</button>' +
      folderBtnHTML() +
      '<span class="tr-files">' +
      '<span class="tr-file' + (t ? " on" : "") + '" data-trfile="tracker"><b>Tracker</b> ' + (t ? esc(t.name) + ' <span class="text-muted">· ' + esc(t.sheetName) + ' · ' + plural(t.sessions.length, "session") + ' · ' + plural(t.roster.length, "trainer") + '</span>' + (t.handle ? ' <span class="badge bg-success" title="Write-back saves into this file">saves in place</span>' : "") : '<span class="text-muted">not loaded</span>') + '</span>' +
      '<span class="tr-file' + (ev || S.queue.length ? " on" : "") + '" data-trfile="eval"><b>Evals</b> ' + (S.queue.length ? plural(S.queue.length, "file") + (ev ? " + " + esc(ev.name) : "") : ev ? esc(ev.name) + ' <span class="text-muted">· ' + plural(ev.rows.length - 1, "response") + '</span>' : '<span class="text-muted">none</span>') + '</span>' +
      '</span></div>' +
      '<div class="tr-drop small text-muted" id="tr-drop">or drop eval exports here — many at once is fine</div>';
    if (pend || added || fills) {
      h += '<div class="tr-pending"><span><b>Not written yet:</b> ' +
        [added ? plural(added, "new session") : "", pend ? plural(pend, "scored session") : "", fills ? plural(fills, "response count") : ""].filter(Boolean).join(" · ") + '</span>' +
        '<button type="button" class="btn btn-primary btn-sm" data-traction="writeback">Write to tracker…</button></div>';
    }
    return h + '<div id="tr-status" class="small text-muted" role="status" aria-live="polite"></div></div>';
  }
  function folderBtnHTML() {
    if (!(typeof QBR.fsFolderSupported === "function" && QBR.fsFolderSupported())) return "";
    var d = folders();
    return '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="folder" title="' +
      (d.length ? "Find the training tracker and eval exports in " + esc(d.map(function (x) { return x.name; }).join(", ")) + " (and its subfolders)"
        : "Pick the folder that holds the training tracker and eval exports — they load together") + '">📁 ' +
      (d.length ? "Load from folder" : "Open folder…") + '</button>';
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
        '<td class="num"' + (v.n == null && (v.nps != null || v.ts != null) ? ' title="Response count not in the tracker (scored before Responses was recorded). Add the eval again to record it."' : "") + '>' + (v.n != null ? v.n : "—") + '</td>' +
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
    h += queueHTML();
    var dups = P().duplicateSessions(t.sessions);
    if (dups.length) h += '<div class="tr-dupwarn small mt-3" role="note"><b>Possible duplicate sessions</b> — same date, school, course and start time on more than one row. ' +
      'For co-trained sessions keep ONE row and put both trainers in the Trainer cell (“NAME / NAME”), then delete the extra row in Excel.<ul class="mb-0">' +
      dups.slice(0, 8).map(function (d) {
        return '<li>' + esc(d.label) + ': rows ' + d.rows.map(function (x) { return excelRow(x) + (x.trainer ? " (" + esc(x.trainer) + ")" : ""); }).join(", ") + '</li>';
      }).join("") + (dups.length > 8 ? '<li>…and ' + (dups.length - 8) + ' more</li>' : "") + '</ul></div>';
    h += '<div class="card-box mt-3"><div class="tr-ch-head"><h6 class="mb-0">Sessions</h6>' +
      '<div class="tr-toolbar"><div class="btn-group btn-group-sm" role="group" aria-label="Status filter">' +
      f("all", "All") + f("needs", "Needs verifying") + f("scored", "Scored") + f("unscored", "Unscored") + '</div>' +
      '<input type="search" class="form-control form-control-sm tr-search" id="tr-search" placeholder="Search trainer, school, course, date…" aria-label="Search sessions" value="' + esc(S.q) + '">' +
      '<button type="button" class="btn btn-outline-primary btn-sm" data-traction="add-session">+ Add session</button></div></div>' +
      '<div id="tr-list">' + tableHTML() + '</div></div>';
    if (S.pickFor) h = h.replace('<div id="tr-list">', '<div class="tr-picking small">Click the session for <b>' + esc((queueItem(S.pickFor) || {}).name || "") + '</b> · <button type="button" class="btn btn-link btn-sm p-0" data-traction="q-pick-cancel">Cancel</button></div><div id="tr-list">');
    if (S.manualForm) h += newSessionFormHTML();
    return h + detailHTML();
  }

  /* ---------------- bulk eval list ---------------- */
  function hhmm(f) { if (f == null) return ""; var m = Math.round(f * 1440), h = Math.floor(m / 60) % 24, mm = m % 60; return (h % 12 || 12) + ":" + (mm < 10 ? "0" : "") + mm + (h < 12 ? " AM" : " PM"); }
  // Date · time · school · trainer (code) · row / Session ID — two same-day sessions never look alike.
  function sessionLabel(s) {
    var tm = s.timeStart != null ? hhmm(s.timeStart) + (s.timeEnd != null ? "–" + hhmm(s.timeEnd) : "") : "";
    return (s.dateISO || "—") + (tm ? " · " + tm : "") + " · " + s.client + " · " + s.trainer + " (" + s.code + ")" +
           " · row " + excelRow(s) + (s.id ? " · " + s.id : "") + (s.isNew ? " · new" : "");
  }
  function queueHTML() {
    if (!S.queue.length) return "";
    var t = S.tracker, ready = 0;
    var rows = S.queue.map(function (q) {
      var st = itemState(q), s = t && q.target >= 0 ? t.sessions[q.target] : null;
      var ex = s ? existingScore(s) : null;
      if (t && q.target >= 0 && !st && !q.needsMap && (!ex || q.replace)) ready++;
      var sel = "";
      if (t) {
        var opts = q.cands.slice();
        if (q.target >= 0 && opts.indexOf(q.target) < 0) opts.unshift(q.target);
        sel = '<select class="form-select form-select-sm" data-trq-target="' + q.id + '" aria-label="Session for ' + esc(q.name) + '"' + (st ? " disabled" : "") + '>' +
          '<option value="-1">— choose a session —</option>' +
          opts.map(function (i) { var x = t.sessions[i]; return x ? '<option value="' + i + '"' + (q.target === i ? " selected" : "") + '>' + esc(sessionLabel(x)) + '</option>' : ""; }).join("") +
          '<option value="pick">Pick in the session list…</option><option value="new">＋ Add as a new session…</option></select>' +
          (q.auto && q.target >= 0 && !st ? '<div class="tr-hint">matched from the filename' + (q.hint && q.hint.dateISO && s && s.dateISO !== q.hint.dateISO ? " + answer date" : "") + '</div>' : "") +
          (!q.cands.length && q.target < 0 ? '<div class="tr-hint">no session on ' + esc((q.hint && q.hint.dateISO) || q.responseDate || "that date") + ' — pick one or add it</div>' : "");
      }
      var res;
      var exBlock = t && ex && q.target >= 0 && !st ? '<div class="small"><span class="risk-badge risk-medium">Already scored</span> NPS <b>' + num(ex.nps, 1) + '</b> · ' + num(ex.ts, 2) + '</div>' +
        '<div class="tr-hint">' + (ex.src ? "from " + esc(ex.src.replace(/ · fp:\w+$/, "")) : "eval file not recorded") + '</div>' +
        '<label class="small"><input type="checkbox" data-trq-replace="' + q.id + '"' + (q.replace ? " checked" : "") + '> Replace existing score</label>' : "";
      if (q.needsMap && st !== "encoded") res = '<button type="button" class="btn btn-outline-primary btn-sm" data-traction="q-map" data-q="' + q.id + '">Confirm columns…</button>' + exBlock;
      else if (st === "generated") res = '<span class="risk-badge risk-high">Generated</span> NPS <b>' + num(s.computed.nps, 1) + '</b> · ' + num(s.computed.trainerScore, 2) + ' <button type="button" class="btn btn-link btn-sm p-0" data-traction="q-open" data-q="' + q.id + '">View</button>';
      else if (st === "written") res = '<span class="risk-badge risk-none">Written</span> <button type="button" class="btn btn-link btn-sm p-0" data-traction="q-open" data-q="' + q.id + '">View</button>';
      else if (st === "encoded") res = '<span class="risk-badge risk-none" title="The tracker’s Eval Source for this row has this file’s fingerprint">Already encoded</span> <span class="small text-muted">row ' + excelRow(s) + ' · NPS ' + num(s.nps, 1) + '</span>';
      else if (exBlock) res = exBlock;
      else if (!t) res = '<button type="button" class="btn btn-outline-primary btn-sm" data-traction="q-report" data-q="' + q.id + '">Report</button>';
      else res = q.target >= 0 ? '<span class="small text-muted">ready</span>' + (s && (s.nps != null || s.trainerScore != null) ? ' <span class="risk-badge risk-medium" title="Generating replaces the score already in the tracker">replaces score</span>' : "") : '<span class="small text-muted">needs a session</span>';
      return '<tr data-q="' + q.id + '"><td><div class="tr-qname" title="' + esc(q.name) + '">' + esc(q.name) + '</div><div class="tr-hint">' + plural(q.rows.length - 1, "response") +
        (q.responseDate ? " · answered " + esc(q.responseDate) + (q.respSession && !(q.hint && q.hint.session) ? " (" + q.respSession + ")" : "") : "") +
        (q.hint ? " · filename " + esc(q.hint.dateISO) + (q.hint.session ? " " + esc(q.hint.session) : "") + (q.hint.school ? " " + esc(q.hint.school) : "") : " · filename not recognised") + '</div>' +
        (q.note ? '<div class="small text-danger">' + esc(q.note) + '</div>' : "") + '</td>' +
        (t ? '<td class="tr-qsel">' + sel + '</td>' : "") + '<td class="text-nowrap">' + res + '</td>' +
        '<td class="text-end"><button type="button" class="btn btn-link btn-sm p-0" data-traction="q-remove" data-q="' + q.id + '" aria-label="Remove ' + esc(q.name) + ' from the list">Remove</button></td></tr>';
    }).join("");
    var unmapped = S.queue.filter(function (q) { return q.needsMap && itemState(q) !== "encoded"; });
    var h = '<div class="card-box mt-3 tr-queue"><div class="tr-ch-head"><div><h6 class="mb-0">Eval files (' + S.queue.length + ')</h6>' +
      '<div class="small text-muted">' + (t ? "Each file is matched to a session from its filename (date · AM/PM · school) and answer dates — check the pick, then generate." : "Open the training tracker to match these files to sessions, or open a one-off report per file.") + '</div></div>' +
      '<div class="tr-toolbar">' + (t ? '<button type="button" class="btn btn-primary btn-sm" data-traction="q-generate"' + (ready ? "" : " disabled") + '>Generate ' + (ready ? plural(ready, "matched file") : "matched files") + '</button>' : "") +
      '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="q-clear">Clear list</button></div></div>' +
      (unmapped.length && !S.mapFor ? '<div class="tr-pending mt-0 mb-2"><span>New form layout: confirm its columns once — every file with the same layout uses it.</span>' +
        '<button type="button" class="btn btn-primary btn-sm" data-traction="q-map" data-q="' + unmapped[0].id + '">Confirm columns…</button></div>' : "") +
      '<div class="table-responsive"><table class="table table-sm mb-0 tr-qtable" data-resizable="static"><thead><tr><th>File</th>' + (t ? "<th>Session</th>" : "") + '<th>Result</th><th></th></tr></thead><tbody>' + rows + '</tbody></table></div>';
    var mq = S.mapFor && queueItem(S.mapFor);
    if (mq) h += mapperHTML(mq);
    var nq = S.newFor && queueItem(S.newFor);
    if (nq && t) h += newSessionFormHTML();
    return h + '</div>';
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
  function mapperHTML(ev) {
    ev = ev || S.eval;
    var opts = ev.header.map(function (h, i) { return '<option value="' + i + '">' + esc(h.slice(0, 70)) + '</option>'; }).join("");
    var h = '<div class="tr-mapper"><h6 class="mb-1">New form layout — confirm the columns</h6><p class="small text-muted mb-2">' + esc(ev.name) + ' — this form layout hasn’t been seen before. Check each role (saved in this browser for next time).</p><div class="row g-2">';
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

  /* ---------------- summary view ----------------
   * All trainers: headline tiles, courses conducted, ranked trainer table, trends.
   * One trainer selected: the same report becomes that trainer's profile. */
  function changeChip(d) {
    if (d == null) return "";
    if (!d) return '<span class="kpi-delta">± 0 vs previous period</span>';
    return '<span class="kpi-delta ' + (d > 0 ? "up" : "down") + '">' + (d > 0 ? "▲ +" : "▼ ") + d + ' vs previous period</span>';
  }
  function courseName(c) { return c.title ? c.title + (c.code ? " (" + c.code + ")" : "") : (c.code || "—"); }
  function summaryHTML() {
    var t = S.tracker, P_ = P(), period = S.repPeriod || "";
    var rows = summaryRows(period), prevKey = P_.prevPeriod(period), prevRows = prevKey ? summaryRows(prevKey) : [];
    var sum = P_.buildSummary(rows), ov = P_.overview(rows), courses = P_.courseTable(rows);
    var profile = !!S.repTrainer, trows = profile ? [] : P_.trainerTable(rows, prevRows);
    var periods = P_.periodsOf(t.sessions);
    var schools = [];
    t.sessions.forEach(function (s) { if (s.client && schools.indexOf(s.client) < 0) schools.push(s.client); });
    schools.sort();
    var pLabel = period ? (periods.filter(function (x) { return x.key === period; })[0] || { label: period }).label : "All time";
    var filt = [pLabel];
    if (S.repTrainer) filt.push("Trainer: " + S.repTrainer);
    if (S.repSchool) filt.push("School: " + S.repSchool);
    var stamp = new Date().toISOString().slice(0, 10);
    var sel = function (id, label, all, opts, cur) {
      return '<label class="small fw-semibold">' + label + ' <select class="form-select form-select-sm" id="' + id + '"><option value="">' + all + '</option>' +
        opts.map(function (o) { return '<option value="' + esc(o.key) + '"' + (cur === o.key ? " selected" : "") + '>' + esc(o.label) + '</option>'; }).join("") + '</select></label>';
    };
    var h = '<div class="card-box mt-3 tr-noprint"><div class="tr-filters-row">' +
      sel("tr-rep-period", "Period", "All time", periods, period) +
      sel("tr-rep-trainer", "Trainer", "All trainers", t.roster.slice().sort().map(function (r) { return { key: r, label: r }; }), S.repTrainer) +
      sel("tr-rep-school", "School", "All schools", schools.map(function (r) { return { key: r, label: r }; }), S.repSchool) +
      (S.repTrainer || S.repSchool || period ? '<button type="button" class="btn btn-outline-secondary btn-sm" data-traction="rep-clear">Clear filters</button>' : "") +
      '<button type="button" class="btn btn-primary btn-sm ms-auto" data-traction="print-report" title="Opens the print dialog — choose “Save as PDF”">' +
      (profile ? "Trainer profile (PDF)" : "Summary report (PDF)") + '</button></div></div>';
    var title = profile ? "Trainer profile — " + S.repTrainer : "Training summary report";
    h += '<div class="tr-report" id="tr-report" data-title="' + esc((profile ? "Trainer profile - " + S.repTrainer : "Training summary") + " - " + (period || "all") + " - " + stamp) + '">' +
      '<div class="tr-print-only tr-print-head"><div class="tr-ph-title">' + esc(title) + '</div>' +
      '<div class="tr-ph-sub">' + esc(filt.join(" · ")) + ' · ' + esc(t.name) + ' · generated ' + stamp + '</div></div>' +
      (profile ? '<h5 class="tr-profile-h tr-noprint">' + esc(S.repTrainer) + ' <span class="small text-muted fw-normal">· ' + esc(pLabel) + '</span></h5>' : "");
    if (profile) {
      var me = P_.trainerTable(rows).filter(function (r) { return P_.normKey(r.trainer) === P_.normKey(S.repTrainer); })[0];
      if (me && me.coTrained) h += '<div class="small text-muted mb-1">' + plural(me.coTrained, "co-trained session") + ' (with ' + esc(me.coWith.join(", ")) +
        ') — counted in full for each trainer, with the shared NPS and trainer score.</div>';
    }
    // tiles
    var cov = ov.coverage == null ? "—" : ov.coverage + "%";
    var covTone = ov.coverage == null ? "blue" : ov.coverage >= 80 ? "green" : ov.coverage >= 50 ? "orange" : "red";
    var prevNps = prevKey ? P_.buildSummary(prevRows).avgNps : null;
    var npsChg = sum.avgNps != null && prevNps != null ? Math.round((sum.avgNps - prevNps) * 10) / 10 : null;
    h += '<div class="row row-cols-2 row-cols-md-3 row-cols-xl-5 g-3 kpi-row mt-1">' +
      (profile ? "" : kpi("Trainers active", ov.trainers, "blue")) +
      kpi("Trainings conducted", ov.sessions, "blue", period ? esc(pLabel) : "all sessions in the tracker") +
      kpi("Training hours", num(ov.hours, 1), "blue", "sessions with start and end times") +
      kpi("Courses delivered", ov.courses, "teal", courses[0] ? "most: " + esc(courseName(courses[0])) + " ×" + courses[0].sessions : "") +
      kpi("Schools served", ov.schools, "teal", sum.repeatRate + "% of trainings at repeat schools") +
      kpi("Respondents", sum.respondents || "—", "blue", sum.respSessions ? plural(sum.respSessions, "session") + " with eval results" : "add eval results to count") +
      kpi("Avg NPS", sum.avgNps == null ? "—" : num(sum.avgNps, 1), npsTone(sum.avgNps), changeChip(npsChg) || plural(sum.scoredNps, "scored session")) +
      kpi("Avg trainer score", sum.avgScore == null ? "—" : num(sum.avgScore, 2), scoreTone(sum.avgScore), "out of 5 · " + plural(sum.scoredTs, "session")) +
      kpi("Eval coverage", cov, covTone, ov.evaluated + " of " + plural(ov.sessions, "training") + " have scores") +
      (profile ? kpi("Last session", esc(rows.reduce(function (x, s) { return s.dateISO && s.dateISO > x ? s.dateISO : x; }, "") || "—"), "blue") : "") +
      '</div>';
    if (!rows.length) return h + '<div class="card-box mt-3 small text-muted">No trainings match these filters.</div></div>';
    // profile: facilitator items
    if (profile) {
      var ia = P_.itemAverages(rows), L = { t1: "Knowledge & mastery", t2: "Communication & presentation", t3: "Engage & connect" };
      var have = ["t1", "t2", "t3"].filter(function (k) { return ia[k] != null; }).sort(function (a, b) { return ia[b] - ia[a]; });
      h += '<div class="card-box tr-inner mt-3"><h6>Facilitator items <span class="small text-muted fw-normal">· from ' + plural(ia.n, "eval") + ' scored in the dashboard</span></h6>' +
        (have.length ? '<div class="row row-cols-1 row-cols-md-3 g-3">' + ["t1", "t2", "t3"].map(function (k) {
          var v = ia[k], tag = have.length > 1 && k === have[0] ? chip("strongest", "none") : have.length > 1 && k === have[have.length - 1] ? chip("to develop", "high") : "";
          return '<div class="col"><div class="tr-item"><div class="small fw-semibold">' + L[k] + ' ' + tag + '</div><div class="tr-item-v">' + num(v, 2) + ' <span class="small text-muted">/ 5</span></div>' +
            '<div class="tr-meter"><span style="width:' + (v == null ? 0 : Math.max(0, Math.min(100, (v - 1) / 4 * 100))) + '%"></span></div></div></div>';
        }).join("") + '</div>' : '<div class="small text-muted">No eval files for this trainer were scored in this browser yet — the tracker stores only NPS and the trainer score.</div>') + '</div>';
    }
    // courses conducted
    var topC = courses.slice(0, 10);
    h += '<div class="row g-3 mt-0"><div class="col-lg-7"><div class="card-box tr-inner"><h6>Courses conducted <span class="small text-muted fw-normal">· ' + plural(courses.length, "course") + '</span></h6>' +
      '<div class="table-responsive"><table class="table table-sm mb-0 tr-courses" data-resizable="static"><thead><tr><th>Course</th><th class="num">Times conducted</th><th class="num">Hours</th><th class="num">Schools</th>' +
      (profile ? "" : '<th class="num">Trainers</th>') + '<th class="num">Avg NPS</th></tr></thead><tbody>' +
      courses.map(function (c) {
        return '<tr><td>' + esc(c.title || c.code || "—") + (c.title && c.code ? ' <span class="text-muted small">' + esc(c.code) + '</span>' : "") + '</td><td class="num"><b>' + c.sessions + '</b></td>' +
          '<td class="num">' + num(c.hours, 1) + '</td><td class="num">' + c.schools + '</td>' + (profile ? "" : '<td class="num">' + c.trainers + '</td>') +
          '<td class="num">' + (c.avgNps == null ? "—" : num(c.avgNps, 1)) + '</td></tr>';
      }).join("") + '</tbody></table></div></div></div>' +
      '<div class="col-lg-5"><div class="card-box tr-inner"><h6>Most conducted courses</h6><div class="chart-wrap" style="height:' + Math.max(200, 28 * topC.length + 50) + 'px"><canvas id="tr-ch-course" aria-label="Times each course was conducted"></canvas></div></div></div></div>';
    S.chartDefs["tr-ch-course"] = function (T) {
      return { type: "bar", data: { labels: topC.map(function (c) { var n = c.title || c.code; return n.length > 30 ? n.slice(0, 28) + "…" : n; }),
        datasets: [{ label: "Times conducted", data: topC.map(function (c) { return c.sessions; }), backgroundColor: T.teal }] },
        options: { indexAxis: "y", scales: { x: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { display: false } } } };
    };
    // trainer table (all trainers)
    if (!profile) {
      h += '<div class="card-box tr-inner mt-3' + (trows.length > 6 ? " tr-pb" : "") + '"><h6>Trainers <span class="small text-muted fw-normal">· ranked by trainings conducted<span class="tr-noprint"> · click a name for the trainer profile</span></span></h6>' +
        '<div class="table-responsive"><table class="table table-sm mb-0 tr-trainers" data-resizable="static"><thead><tr><th>#</th><th>Trainer</th><th class="num">Trainings</th><th class="num">Hours</th>' +
        '<th class="num">Courses</th><th class="num">Schools</th><th class="num">Respondents</th><th class="num">Avg NPS</th><th class="num">Avg score</th>' +
        (prevKey ? '<th class="num">vs previous</th>' : "") + '<th class="num">No eval</th><th>Last session</th></tr></thead><tbody>' +
        trows.map(function (r, i) {
          var tone = function (v, f) { return v == null ? "" : ' class="num tr-t-' + f(v) + '"'; };
          return '<tr><td class="text-muted">' + (i + 1) + '</td><td><button type="button" class="btn btn-link btn-sm p-0 tr-tname" data-trtrainer="' + esc(r.trainer) + '">' + esc(r.trainer) + '</button>' +
            (r.topCourse ? '<div class="tr-hint">top: ' + esc(r.topCourse.title || r.topCourse.code) + ' ×' + r.topCourse.sessions + '</div>' : "") +
            (r.coTrained ? '<div class="tr-hint" title="Co-trained sessions count for each trainer, with the session’s shared NPS and score">' + r.coTrained + ' co-trained · with ' + esc(r.coWith.join(", ")) + '</div>' : "") + '</td>' +
            '<td class="num"><b>' + r.sessions + '</b></td><td class="num">' + num(r.hours, 1) + '</td><td class="num">' + r.courses + '</td><td class="num">' + r.schools + '</td>' +
            '<td class="num">' + (r.respondents || "—") + '</td>' +
            '<td' + (tone(r.avgNps, npsTone) || ' class="num"') + '>' + (r.avgNps == null ? "—" : num(r.avgNps, 1)) + '</td>' +
            '<td' + (tone(r.avgScore, scoreTone) || ' class="num"') + '>' + (r.avgScore == null ? "—" : num(r.avgScore, 2)) + '</td>' +
            (prevKey ? '<td class="num">' + (r.npsChange == null ? "—" : (r.npsChange > 0 ? "▲ +" : r.npsChange < 0 ? "▼ " : "") + r.npsChange) + '</td>' : "") +
            '<td class="num">' + (r.noEval ? '<span class="tr-t-orange">' + r.noEval + '</span>' : "0") + '</td><td class="text-nowrap">' + esc(r.last || "—") + '</td></tr>';
        }).join("") + '</tbody></table></div>' +
        (prevKey ? '<div class="small text-muted mt-1">“vs previous” = change in average NPS against ' + esc(prevKey.replace("-", " ")) + '.</div>' : "") + '</div>';
    }
    // trends
    var months = sum.monthly.filter(function (m) { return m.key !== "(no date)"; });
    h += '<div class="row g-3 mt-0' + (profile ? "" : " tr-pb") + '"><div class="col-lg-7"><div class="card-box tr-inner"><h6>Trainings and NPS by month</h6><div class="chart-wrap"><canvas id="tr-ch-month" aria-label="Trainings and average NPS by month"></canvas></div></div></div>' +
      '<div class="col-lg-5"><div class="card-box tr-inner"><h6>Top schools by trainings</h6><div class="chart-wrap"><canvas id="tr-ch-school" aria-label="Trainings per school"></canvas></div></div></div></div>';
    S.chartDefs["tr-ch-month"] = function (T) {
      return { type: "bar", data: { labels: months.map(function (m) { var d = new Date(+m.key.slice(0, 4), +m.key.slice(5, 7) - 1, 1); return isNaN(d) ? m.key : d.toLocaleString("en-US", { month: "short", year: "2-digit" }); }),
        datasets: [
          { type: "line", label: "Avg NPS", data: months.map(function (m) { return m.avgNps; }), borderColor: T.ok, backgroundColor: T.ok, yAxisID: "y2", tension: 0.25, pointRadius: 3, spanGaps: true },
          { type: "bar", label: "Trainings", data: months.map(function (m) { return m.sessions; }), backgroundColor: T.brand + "55", borderColor: T.brand, borderWidth: 1, yAxisID: "y" }] },
        options: { scales: { y: { beginAtZero: true, title: { display: true, text: "Trainings" }, ticks: { precision: 0 } },
          y2: { position: "right", min: -100, max: 100, grid: { drawOnChartArea: false }, title: { display: true, text: "NPS" } } } } };
    };
    var topS = sum.bySchool.slice().sort(function (a, b) { return b.sessions - a.sessions; }).slice(0, 10);
    S.chartDefs["tr-ch-school"] = function (T) {
      return { type: "bar", data: { labels: topS.map(function (b) { return b.key.length > 32 ? b.key.slice(0, 30) + "…" : b.key; }),
        datasets: [{ label: "Trainings", data: topS.map(function (b) { return b.sessions; }), backgroundColor: T.brand }] },
        options: { indexAxis: "y", scales: { x: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { legend: { display: false } } } };
    };
    var bySessions = function (a, b) { return b.sessions - a.sessions || (a.key < b.key ? -1 : 1); };
    h += '<div class="row g-3 mt-0"><div class="col-lg-6"><div class="card-box tr-inner"><h6>By school</h6><div class="table-responsive"><table class="table table-sm mb-0 tr-break" data-resizable="static"><thead><tr>' +
      '<th>School</th><th class="num">Trainings</th><th class="num">Avg NPS</th><th class="num">Avg trainer score</th></tr></thead><tbody>' +
      sum.bySchool.slice().sort(bySessions).map(function (r) {
        return '<tr><td>' + esc(r.key) + '</td><td class="num">' + r.sessions + '</td><td class="num">' + (r.avgNps == null ? "—" : num(r.avgNps, 1)) + '</td>' +
               '<td class="num">' + (r.avgScore == null ? "—" : num(r.avgScore, 2)) + '</td></tr>';
      }).join("") + '</tbody></table></div></div></div>';
    // profile: trainings still without an eval (to chase)
    if (profile) {
      var miss = rows.filter(function (s) { return s.nps == null && s.trainerScore == null; }).sort(function (a, b) { return a.dateISO < b.dateISO ? 1 : -1; });
      h += '<div class="col-lg-6"><div class="card-box tr-inner"><h6>Trainings without an eval <span class="small text-muted fw-normal">· ' + miss.length + '</span></h6>' +
        (miss.length ? '<div class="table-responsive"><table class="table table-sm mb-0" data-resizable="static"><thead><tr><th>Date</th><th>Course</th><th>School</th></tr></thead><tbody>' +
          miss.slice(0, 25).map(function (s) { return '<tr><td class="text-nowrap">' + esc(s.dateISO || "—") + '</td><td>' + esc(s.title || s.code) + '</td><td>' + esc(s.client) + '</td></tr>'; }).join("") +
          (miss.length > 25 ? '<tr><td colspan="3" class="text-muted">…and ' + (miss.length - 25) + ' more</td></tr>' : "") + '</tbody></table></div>'
          : '<div class="small text-muted">Every training has scores.</div>') + '</div></div>';
    }
    h += '</div>';
    return h + '<div class="small text-muted mt-2">Averages include scores generated here that are not written to the tracker yet. Hours come from the tracker’s start and end times.</div></div>';
  }

  /* ---------------- assembly + events ---------------- */
  function headHTML() {
    return '<div class="tr-head"><div><h3 class="dash-h mb-1">Trainer evaluations</h3>' +
      '<div class="small text-muted">Training tracker sessions · MS Forms evaluation exports · NPS and trainer scores written back to the tracker</div></div>' +
      '<div class="tr-privacy small" title="Evaluation answers are kept in this browser tab only — nothing is uploaded.">🔒 Answers stay in this tab</div></div>';
  }
  function render() {
    init(); // idempotent; loads saved column maps
    autoFolder();
    var host = $("trainers-body");
    if (!host) return;
    destroyCharts();
    if (!on()) { host.innerHTML = ""; return; }
    var h = headHTML() + loadCardHTML();
    if (!S.tracker) {
      h += queueHTML();
      if (S.eval) h += soloHTML();
      else if (!S.queue.length) h += '<div class="card-box mt-3 tr-empty"><h6>Get started</h6><ol class="small mb-0">' +
        '<li><b>Open training tracker…</b> — the session log (e.g. ETG_Training_Tracker_2026.xlsx, “trial” sheet).</li>' +
        '<li>Pick a session and <b>add its eval export</b> (MS Forms .xlsx, one file per session) — or add the eval first and confirm the suggested session.</li>' +
        '<li><b>Generate results</b>, check them, then <b>Write to tracker</b>. Each session gets a feedback report you can save as PDF.</li></ol>' +
        '<div class="small text-muted mt-2">Several sessions? Select or drop all their eval exports at once — each file is matched to its session. Only have one eval file? Add it on its own for a one-off session report.</div></div>';
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
    var pq = S.pickFor && queueItem(S.pickFor);
    if (pq) { pq.target = i; pq.auto = false; pq.note = ""; pq.replace = false; learnAlias(pq, i); S.pickFor = null; }
    else if (S.eval && S.eval.attachedTo < 0 && S.tracker) attachTo(i); // "click a session to attach this file"
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
    host.querySelectorAll("[data-trq-target]").forEach(function (sel) {
      sel.addEventListener("change", function () {
        var q = queueItem(sel.getAttribute("data-trq-target")); if (!q) return;
        var v = sel.value;
        if (v === "pick") { S.pickFor = q.id; S.view = "sessions"; render(); var l = $("tr-list"); if (l && l.scrollIntoView) try { l.scrollIntoView({ block: "start" }); } catch (e) {} return; }
        if (v === "new") {
          var tw = q.twin >= 0 ? q.twin : q.encodedAt >= 0 ? q.encodedAt : -1;
          if (tw >= 0 && !window.confirm("Row " + excelRow(S.tracker.sessions[tw]) + " (" + sessionLabel(S.tracker.sessions[tw]) + ") already has this eval's scores.\n\nAdding it as a new session writes the same eval twice. Add a new session anyway?")) { sel.value = String(q.target); return; }
          S.newFor = q.id; S.manualForm = false; S.showingNewForm = false; render(); return;
        }
        q.target = +v; q.auto = false; q.note = ""; q.replace = false; if (q.target >= 0) learnAlias(q, q.target); render();
      });
    });
    host.querySelectorAll("[data-trq-replace]").forEach(function (cb) {
      cb.addEventListener("change", function () { var q = queueItem(cb.getAttribute("data-trq-replace")); if (q) { q.replace = cb.checked; q.note = ""; render(); } });
    });
    var drop = host.querySelector(".tr-load");
    if (drop) {
      drop.addEventListener("dragover", function (e) { e.preventDefault(); drop.classList.add("tr-over"); });
      drop.addEventListener("dragleave", function () { drop.classList.remove("tr-over"); });
      drop.addEventListener("drop", function (e) {
        e.preventDefault(); drop.classList.remove("tr-over");
        openEvals(e.dataTransfer && e.dataTransfer.files).catch(function (er) { alert("Could not open the files: " + er.message); });
      });
    }
    var repT = host.querySelector("#tr-rep-trainer"), repS = host.querySelector("#tr-rep-school");
    if (repT) repT.addEventListener("change", function () { S.repTrainer = repT.value; render(); });
    if (repS) repS.addEventListener("change", function () { S.repSchool = repS.value; render(); });
    var repP = host.querySelector("#tr-rep-period");
    if (repP) repP.addEventListener("change", function () { S.repPeriod = repP.value; render(); });
    host.querySelectorAll("[data-trtrainer]").forEach(function (b) {
      b.addEventListener("click", function () { S.repTrainer = b.getAttribute("data-trtrainer"); render(); var r = $("tr-report"); if (r && r.scrollIntoView) try { r.scrollIntoView({ block: "start" }); } catch (e) {} });
    });
    host.querySelectorAll("[data-traction]").forEach(function (b) {
      b.addEventListener("click", function () { action(b.getAttribute("data-traction"), b); });
    });
    var mapEv = (S.mapFor && queueItem(S.mapFor)) || (S.eval && S.eval.needsMap ? S.eval : null);
    if (mapEv) {
      host.querySelectorAll("[data-trmap]").forEach(function (sel) {
        var cur = mapEv.roles[sel.getAttribute("data-trmap")];
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
  // kind: "tracker" | "eval" (top button: one or many files) | "session-eval" (for the selected session)
  function pickFile(kind) {
    if (kind === "tracker" && fsSupported()) { pickTrackerHandle(); return; }
    var inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx,.xlsm,.xls";
    if (kind === "eval") inp.multiple = true;
    inp.style.display = "none";
    document.body.appendChild(inp); // some browsers ignore click() on a detached input
    inp.onchange = function () {
      var files = inp.files ? Array.prototype.slice.call(inp.files) : [];
      inp.remove();
      if (!files.length) return;
      var f = files[0];
      var p = kind === "tracker" ? openTracker(f, null) : kind === "eval" ? openEvals(files) : openEval(f);
      p.catch(function (e) { alert("Could not open " + (files.length > 1 ? "the files" : f.name) + ": " + e.message); });
    };
    inp.click();
    setTimeout(function () { if (inp.parentNode && !(inp.files && inp.files.length)) inp.remove(); }, 120000);
  }

  /* ---------------- linked folder (v1.39.0) ----------------
   * The folder the dashboard workbooks are linked from (📁 Open folder…) — or one picked
   * here — is scanned for the training tracker and the MS Forms eval exports, so they
   * load without picking each file. The tracker keeps its file handle (saves in place).
   * Only file names that look like a tracker / an eval are opened; the content decides. */
  function folders() { try { return typeof QBR.fsFolderList === "function" ? QBR.fsFolderList() : []; } catch (e) { return []; } }
  function looksTracker(n) { return /train/i.test(n) && /track/i.test(n); }
  function looksEval(n) { return !looksTracker(n) && !/gdap|granular/i.test(n) && (!!P().parseFilename(n) || /eval|feedback|form|survey/i.test(n)); }
  function isEvalWb(wb) {
    var rows = sheetRows(wb, wb.SheetNames[0]);
    if (rows.length < 2) return false;
    var det = P().detectLayout(rows[0]);
    if (S.maps[det.sig]) return true;
    var r = det.roles;
    return (r.nps >= 0 ? 1 : 0) + (r.t1 >= 0 ? 1 : 0) + (r.t2 >= 0 ? 1 : 0) + (r.t3 >= 0 ? 1 : 0) + (r.sat >= 0 ? 1 : 0) >= 2;
  }
  async function fileOf(b) { try { return await b.handle.getFile(); } catch (e) { return null; } }
  async function readFolder(opts) {
    opts = opts || {};
    if (S.folderBusy) return;
    var list = folders();
    if (!list.length) {
      if (!opts.ask || typeof QBR.fsFolderPickOnly !== "function") return;
      var nid = await QBR.fsFolderPickOnly();
      if (!nid) return;
      list = folders();
    }
    S.folderBusy = true;
    try {
      var books = [], used = [];
      for (var i = 0; i < list.length; i++) {
        var b = await QBR.fsFolderBooks(list[i].id, { ask: !!opts.ask });
        if (b) { used.push(list[i].name); books = books.concat(b); }
      }
      if (!used.length) { if (opts.ask) note("The folder wasn't allowed — nothing was read."); return; }
      var msg = [], tName = "";
      if (!S.tracker) {
        var tc = books.filter(function (b) { return looksTracker(b.name); })
          .sort(function (a, b) { return (b.lastModified || 0) - (a.lastModified || 0); });
        for (var j = 0; j < tc.length && !S.tracker; j++) {
          var tf = await fileOf(tc[j]); if (!tf) continue;
          note("Reading " + tf.name + "…");
          var perm = "prompt";
          try { perm = await tc[j].handle.queryPermission({ mode: "readwrite" }); } catch (e) {}
          try { await openTracker(tf, perm === "granted" ? tc[j].handle : null); tName = tf.name; } catch (e) {}
        }
      }
      if (tName) msg.push("tracker " + tName);
      var inList = {};
      S.queue.concat(S.eval ? [S.eval] : []).forEach(function (q) { inList[P().normKey(q.name)] = 1; });
      var ec = books.filter(function (b) { return looksEval(b.name) && !inList[P().normKey(b.name)] && (!b.size || b.size < 20e6); });
      var evals = [];
      for (var k = 0; k < ec.length; k++) {
        var f = await fileOf(ec[k]); if (!f) continue;
        try { if (isEvalWb(await readWorkbook(f))) evals.push(f); } catch (e) {}
      }
      var r = evals.length ? await openEvals(evals, { auto: true }) : { added: 0, encoded: 0, dup: 0, bad: [] };
      if (r.added) msg.push(plural(r.added, "eval export"));
      if (!opts.ask && !tName && !r.added) return; // silent auto-scan: nothing new, say nothing
      render();
      note("📁 " + used.join(", ") + ": " + (msg.length ? msg.join(" + ") + " loaded" : "no new tracker or eval exports found") +
        (r.encoded ? " · " + plural(r.encoded, "eval") + " already encoded (left out)" : "") +
        (r.dup ? " · " + r.dup + " already in the list" : "") + ".");
    } finally { S.folderBusy = false; }
  }
  // Auto-read once per folder grant: when the tab opens, and when a folder is linked/reconnected.
  function autoFolder() {
    if (S.folderTried || !on()) return;
    if (!folders().some(function (d) { return d.state === "granted"; })) return;
    S.folderTried = true;
    setTimeout(function () { readFolder({}).catch(function (e) { console.warn("[QBR] trainers folder scan:", e && e.message); }); }, 0);
  }
  document.addEventListener("qbr:folder-linked", function () {
    S.folderTried = false;
    var p = $("dash-trainers");
    if (p && !p.classList.contains("d-none")) autoFolder();
  });

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
  /* Paper layout, decided from the content: side-by-side cards stay together only while both are
   * short; a row holding a long table is stacked and allowed to run across pages (a flex row can't
   * break across pages, so it was pushed whole to the next page and left the previous one half empty). */
  var PRINT_KEEP_ROWS = 12;
  function tableRows(el) {
    var n = 0;
    el.querySelectorAll("table").forEach(function (t) { n = Math.max(n, t.querySelectorAll("tbody tr").length); });
    // a recommendation is ~4 lines on paper
    return n + el.querySelectorAll(".tr-comment li").length + 4 * el.querySelectorAll(".tr-recs > li").length;
  }
  function layoutForPrint(doc) {
    doc.querySelectorAll(".row").forEach(function (row) {
      if (row.classList.contains("kpi-row")) return;
      row.classList.add(tableRows(row) > PRINT_KEEP_ROWS ? "tr-p-stack" : "tr-p-pair");
    });
    doc.querySelectorAll(".card-box").forEach(function (c) {
      // A card on its own may run on to the next page sooner (its header row repeats) than leave a gap.
      var alone = !c.closest(".row.tr-p-pair");
      c.classList.add(tableRows(c) > (alone ? 6 : PRINT_KEEP_ROWS) ? "tr-p-flow" : "tr-p-keep");
    });
    // Tiles: as many per row as fit on paper (5), fewer only when there are fewer tiles.
    doc.querySelectorAll(".kpi-row").forEach(function (r) {
      var n = r.children.length;
      r.style.setProperty("--tr-tiles", String(n <= 4 ? n : n % 5 === 0 || n > 6 ? 5 : n === 6 ? 3 : 4));
    });
  }
  /* Measure the print copy at paper width and walk it page by page: a block that must stay together
   * but doesn't fit on the rest of the page would leave a gap there. If that gap is large, a card pair is
   * stacked (each card then fits or flows on its own) and a card with a table/list is allowed to run on. */
  var PAGE_W = 720, PAGE_H = 950; // px at 96 dpi: printable area of Letter/A4 with default margins
  function fitToPages(host) {
    var root = host.querySelector(".tr-report");
    if (!root) return { changed: 0 };
    host.classList.add("tr-measure");
    var changed = 0, pages = 1;
    try {
      var y = 0, guard = 0;
      var blocks = Array.prototype.slice.call(root.children);
      for (var i = 0; i < blocks.length && guard++ < 400; i++) {
        var b = blocks[i], cs = getComputedStyle(b);
        if (cs.display === "none") continue;
        var h = b.offsetHeight + (parseFloat(cs.marginTop) || 0) + (parseFloat(cs.marginBottom) || 0);
        var keep = b.classList.contains("tr-p-pair") || b.classList.contains("tr-p-keep") || b.classList.contains("kpi-row") || b.classList.contains("tr-print-head");
        if (!keep || y + h <= PAGE_H) { y += h; while (y > PAGE_H) { y -= PAGE_H; pages++; } continue; }
        var gap = PAGE_H - y, big = gap > PAGE_H * 0.2;
        if (big && b.classList.contains("tr-p-pair")) {          // stack the pair and look at its cards again
          b.classList.remove("tr-p-pair"); b.classList.add("tr-p-stack"); changed++;
          var cards = Array.prototype.slice.call(b.querySelectorAll(":scope > [class*='col-'] > .card-box"));
          cards.forEach(function (c) { if (c.querySelector("table, ul, ol")) { c.classList.remove("tr-p-keep"); c.classList.add("tr-p-flow"); } });
          blocks.splice.apply(blocks, [i + 1, 0].concat(cards.map(function (c) { return c.parentNode; })));
          y += 0; // the row itself has no height of its own now; its columns are walked next
          continue;
        }
        if (big && b.querySelector("table, ul, ol")) {           // let it run on to the next page
          b.classList.remove("tr-p-keep"); b.classList.add("tr-p-flow"); changed++;
          y += h; while (y > PAGE_H) { y -= PAGE_H; pages++; } continue;
        }
        pages++; y = h; while (y > PAGE_H) { y -= PAGE_H; pages++; }  // starts on the next page
      }
    } finally { host.classList.remove("tr-measure"); }
    S.lastFit = { changed: changed, pages: pages };
    return S.lastFit;
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
    layoutForPrint(doc);
    doc.querySelectorAll("canvas").forEach(function (cv) {
      var url = imgs[cv.id];
      if (!url) { var wrap = cv.closest(".chart-wrap"); if (wrap) wrap.remove(); else cv.remove(); return; }
      var im = document.createElement("img"); im.src = url; im.alt = cv.getAttribute("aria-label") || ""; im.className = "tr-print-chart";
      var w = cv.closest(".chart-wrap");
      if (w) { w.replaceWith(im); } else cv.replaceWith(im);
    });
    // No separate footer block: on its own it could spill onto an otherwise empty last page.
    var ph = doc.querySelector(".tr-print-head");
    if (ph) ph.insertAdjacentHTML("beforeend", '<div class="tr-ph-by">RCT OpsDesk · Trainer evaluations · printed ' + esc(new Date().toLocaleString()) + '</div>');
    host.innerHTML = '<div class="tr-print-doc">' + doc.outerHTML + '</div>';
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
      .then(function () { try { fitToPages(host); } catch (e) {} try { window.print(); } finally { setTimeout(cleanup, 60000); } });
  }

  /* ---------------- actions ---------------- */
  function action(name, el) {
    if (name === "export") { exportResults(); return; }
    if (name === "rep-clear") { S.repTrainer = ""; S.repSchool = ""; S.repPeriod = ""; render(); return; }
    if (name === "print-report") return printReport("tr-report");
    if (name === "print-session") return printReport("tr-report-session");
    if (name === "upload-eval") { pickFile("session-eval"); return; }
    if (name === "folder") { readFolder({ ask: true }).catch(function (e) { alert("Could not read the folder: " + (e && e.message || e)); }); return; }
    var qid = el ? el.getAttribute("data-q") : null, qi = qid ? queueItem(qid) : null;
    if (name === "q-generate") { generateQueue(); return; }
    if (name === "q-clear") {
      var pendQ = S.queue.filter(function (q) { return itemState(q) === "generated"; }).length;
      if (pendQ && !window.confirm("Clear the list? The " + plural(pendQ, "generated result") + " stay on their sessions until you write or regenerate them.")) return;
      S.queue = []; S.pickFor = S.mapFor = S.newFor = null; render(); return;
    }
    if (name === "q-remove" && qi) {
      S.queue = S.queue.filter(function (q) { return q !== qi; });
      if (S.pickFor === qi.id) S.pickFor = null; if (S.mapFor === qi.id) S.mapFor = null; if (S.newFor === qi.id) S.newFor = null;
      if (S.eval === qi) S.eval = null;
      render(); return;
    }
    if (name === "q-map" && qi) { S.mapFor = qi.id; render(); return; }
    if (name === "q-open" && qi && qi.target >= 0) { S.view = "sessions"; S.pickFor = null; selectRow(qi.target); return; }
    if (name === "q-report" && qi) { S.eval = qi; render(); return; }
    if (name === "q-pick-cancel") { S.pickFor = null; render(); return; }
    if (name === "attach") {
      var idx = el ? +el.getAttribute("data-idx") : -1;
      if (idx >= 0 && S.eval) { S.selected = idx; attachTo(idx); S.showingNewForm = false; }
      render();
      return;
    }
    if (name === "new-session") { S.showingNewForm = true; render(); return; }
    if (name === "cancel-new-session") { S.showingNewForm = false; S.newFor = null; render(); return; }
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
      var mev = (S.mapFor && queueItem(S.mapFor)) || S.eval;
      mev.roles = roles; mev.needsMap = false;
      S.maps[mev.sig] = { roles: roles, savedAt: new Date().toISOString() };
      // Same form layout in other files: use the same mapping.
      S.queue.concat(S.eval ? [S.eval] : []).forEach(function (q) { if (q.sig === mev.sig) { q.roles = roles; q.needsMap = false; delete q._quick; } });
      S.queue.forEach(function (q) { if (q.sig === mev.sig) matchItem(q); });
      S.mapFor = null;
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
