/* =============================================================================
 * mos.js — v1.37.0 MOS tab (Certification › MOS results)
 * -----------------------------------------------------------------------------
 * Reads Certiport "EXAM SCORE REPORT" PDFs (+ optional school roster .xlsx) in
 * the browser and summarises them School › Section › student.
 *
 *  Inputs    score-report PDFs (files or a whole folder), roster workbook
 *  Parsing   js/mos-parse.js (pure) on text from vendored pdf.js (libs/pdfjs,
 *            loaded on first use; worker runs on the main thread so file:// works)
 *  Page      filters (school, section, student, score range + high/low sort,
 *            pass/fail, Exam No., exam, period: all / school year / year / month),
 *            KPIs, pass-rate trend (monthly / yearly / school year), section
 *            comparison, score bands, weakest skill areas, section summary,
 *            results table with per-student history, Check & fix panel, Excel export.
 *
 * PRIVACY: student names/emails are minors' data. Everything lives in this tab's
 * memory only — no localStorage/IndexedDB, no network. Reload = gone. Addresses on
 * the score report are never read. Not part of the Export Deck (data-deck="skip").
 * Feature flag: "mos" (Certification) — remote flags can hide the whole group.
 * ========================================================================== */
(function () {
  "use strict";
  var QBR = (window.QBR = window.QBR || {});
  var P = function () { return QBR.mosParse; };

  if (QBR.features && QBR.features.register)
    QBR.features.register({ module: "certification", key: "mos", label: "MOS results", desc: "Certiport MOS score reports: pass rates by school, section and student" });

  var S = {
    attempts: [], files: [], exports: [], roster: [], rosterFiles: [], overrides: {},
    model: null, built: false, busy: false, progress: "",
    f: { school: "", section: "", exam: "", examNo: "", proctor: "", outcome: "", min: "", max: "", pmode: "all", pval: "", q: "", firstOnly: false },
    sort: "score-desc", groupBy: "auto", limit: 200, trend: "month", open: {}, charts: {}
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function pct(v, d) { return v == null || !isFinite(v) ? "—" : (v * 100).toFixed(d == null ? 1 : d) + "%"; }
  function num(v) { return v == null || !isFinite(v) ? "—" : Math.round(v).toLocaleString(); }
  function on() { try { return typeof QBR.adminIsOn !== "function" || QBR.adminIsOn("mos"); } catch (e) { return true; } }
  function cssVar(n, fb) { try { var v = getComputedStyle(document.documentElement).getPropertyValue(n).trim(); return v || fb; } catch (e) { return fb; } }
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function monLabel(ym) { return ym && ym.length >= 7 ? MON[+ym.slice(5, 7) - 1] + " " + ym.slice(0, 4) : ym; }
  function fmtDate(iso) { return iso ? MON[+iso.slice(5, 7) - 1] + " " + (+iso.slice(8, 10)) + ", " + iso.slice(0, 4) : "—"; }

  /* ---------------- pdf.js (lazy, offline) ---------------- */
  var _pdf = null;
  function loadScript(src) {
    return new Promise(function (res, rej) {
      var s = document.createElement("script"); s.src = src; s.async = true;
      s.onload = function () { res(); }; s.onerror = function () { rej(new Error("Could not load " + src)); };
      document.head.appendChild(s);
    });
  }
  function pdfLib() {
    if (_pdf) return _pdf;
    _pdf = (window.pdfjsLib ? Promise.resolve() : loadScript("libs/pdfjs/pdf.min.js"))
      // Main-thread worker: a real Worker can't start from file://; this also keeps CSP simple.
      .then(function () { return window.pdfjsWorker ? null : loadScript("libs/pdfjs/pdf.worker.min.js"); })
      .then(function () {
        if (!window.pdfjsLib) throw new Error("pdf.js did not load");
        try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = "libs/pdfjs/pdf.worker.min.js"; } catch (e) {}
        return window.pdfjsLib;
      });
    _pdf.catch(function () { _pdf = null; });
    return _pdf;
  }
  function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }

  async function readPdf(file, schoolInput) {
    var path = file.webkitRelativePath || file._path || file.name;
    var meta = P().fileMeta(path, schoolInput);
    var rec = { file: meta.file, name: meta.fileName, school: meta.school, pages: 0, attempts: 0, skipped: 0, bad: 0, err: "" };
    try {
      var lib = await pdfLib();
      var buf = new Uint8Array(await file.arrayBuffer());
      var doc = await lib.getDocument({ data: buf, isEvalSupported: false, verbosity: 0 }).promise;
      rec.pages = doc.numPages;
      for (var i = 1; i <= doc.numPages; i++) {
        var pg = await doc.getPage(i);
        var tc = await pg.getTextContent();
        var r = P().parseScorePage(tc.items);
        try { pg.cleanup(); } catch (e) {}
        if (!r) { rec.skipped++; continue; }
        if (r._err) { rec.bad++; continue; }
        for (var k in meta) if (Object.prototype.hasOwnProperty.call(meta, k)) r[k] = meta[k];
        r.page = i;
        S.attempts.push(r); rec.attempts++;
        if (i % 4 === 0) { S.progress = "Reading " + rec.name + " · page " + i + "/" + doc.numPages; status(); await tick(); }
      }
      try { await doc.destroy(); } catch (e) {}
    } catch (e) { rec.err = String((e && e.message) || e); }
    return rec;
  }

  async function addPdfs(list) {
    var files = Array.prototype.slice.call(list || []).filter(function (f) { return /\.pdf$/i.test(f.name); });
    if (!files.length || S.busy) return;
    S.busy = true;
    var school = ($("mos-school-in") || {}).value || "";
    var done = 0;
    for (var i = 0; i < files.length; i++) {
      var f = files[i], id = (f.webkitRelativePath || f.name) + "|" + f.size;
      if (S.files.some(function (x) { return x.id === id; })) continue;
      S.progress = "Reading file " + (i + 1) + " of " + files.length + " · " + f.name; status(); await tick();
      var rec = await readPdf(f, school); rec.id = id;
      S.files.push(rec); done++;
    }
    S.busy = false; S.progress = "";
    rebuild();
  }

  async function addRoster(list) {
    var files = Array.prototype.slice.call(list || []).filter(function (f) { return /\.xls[xm]?$/i.test(f.name); });
    if (!files.length || typeof XLSX === "undefined") return;
    var input = ($("mos-school-in") || {}).value || "";
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var meta = P().fileMeta(f.webkitRelativePath || f.name, input);
      var school = meta.school;
      if (school === "Unassigned school") {
        var sc = {}; S.attempts.forEach(function (a) { sc[a.school] = 1; });
        var ks = Object.keys(sc);
        if (S.f.school) school = S.f.school; else if (ks.length === 1) school = ks[0];
      }
      var id = (f.webkitRelativePath || f.name) + "|" + f.size;
      if (S.exports.some(function (x) { return x.id === id; }) || S.rosterFiles.some(function (x) { return x.id === id; })) continue;
      try {
        var wb = XLSX.read(new Uint8Array(await f.arrayBuffer()), { type: "array", cellDates: true });
        var n = 0, exp = null;
        wb.SheetNames.forEach(function (sn) {
          // v1.37.0: a Certiport "Results … Plus" export is recognised by its header row (Exam Reference Number + Score + Result)
          var raw = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: true });
          var ex = P().parseExportRows(raw, f.name);
          if (ex) {
            ex.attempts.forEach(function (a) { S.attempts.push(a); });
            exp = exp || { id: id, name: f.name, sheets: [], attempts: 0, skipped: 0 };
            exp.sheets.push(sn); exp.attempts += ex.attempts.length; exp.skipped += ex.skipped;
            return;
          }
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", raw: false });
          P().parseRosterRows(sn, rows).forEach(function (st) { st.school = school; S.roster.push(st); n++; });
        });
        if (exp) S.exports.push(exp);
        if (!exp || n) S.rosterFiles.push({ id: id, name: f.name, school: school, students: n });
      } catch (e) { S.rosterFiles.push({ id: id, name: f.name, school: school, students: 0, err: String(e.message || e) }); }
    }
    rebuild();
  }

  function clearAll() {
    S.attempts = []; S.files = []; S.exports = []; S.roster = []; S.rosterFiles = []; S.overrides = {}; S.model = null; S.open = {};
    S.f = { school: "", section: "", exam: "", examNo: "", proctor: "", outcome: "", min: "", max: "", pmode: "all", pval: "", q: "", firstOnly: false };
    render();
  }

  function rebuild() {
    S.model = S.attempts.length ? P().buildModel(S.attempts, S.roster, S.overrides, { acronyms: acronymMap() }) : null;
    render();
  }

  /* ---------------- filtering ---------------- */
  function filterObj(extra) {
    var f = S.f, o = {
      school: f.school, section: f.section, exam: f.exam, examNo: f.examNo, proctor: f.proctor, outcome: f.outcome,
      min: f.min, max: f.max, q: f.q, firstOnly: f.firstOnly
    };
    if (f.pmode === "sy" && f.pval) o.sy = f.pval;
    if (f.pmode === "year" && f.pval) o.year = f.pval;
    if (f.pmode === "month" && f.pval) o.month = f.pval;
    if (extra) for (var k in extra) o[k] = extra[k];
    return o;
  }
  function rows() { return S.model ? P().filterAttempts(S.model.rows, filterObj()) : []; }
  function uniq(arr) { var o = {}; arr.forEach(function (v) { if (v != null && v !== "") o[v] = 1; }); return Object.keys(o).sort(); }

  /* ---------------- shell ---------------- */
  function shell(host) {
    host.innerHTML =
      '<div class="mos-head">' +
        '<div><h3 class="dash-h mb-1">MOS certification results</h3>' +
        '<div class="small text-muted">Microsoft Office Specialist score reports exported from Certiport · School › Section › Student</div></div>' +
        '<div class="mos-privacy small" title="Nothing is saved or uploaded. Closing or reloading the tab clears it.">🔒 Student data stays in this tab only</div>' +
      '</div>' +
      '<div class="card-box mos-load">' +
        '<div class="mos-load-row">' +
          '<label class="btn btn-primary btn-sm mb-0">Add score reports (PDF)…<input id="mos-pdf" type="file" accept=".pdf,application/pdf" multiple hidden></label>' +
          '<label class="btn btn-outline-primary btn-sm mb-0" title="Pick a school folder (or the whole MOS folder). Numbered folders like “01 - School name” become the school.">Add folder…<input id="mos-dir" type="file" webkitdirectory multiple hidden></label>' +
          '<label class="btn btn-outline-secondary btn-sm mb-0" title="Certiport results export (Results … Plus .xlsx: proctor, station, exam date) or a school roster (one sheet per section: 11A, 11B… with NAME and Email). The app tells them apart.">Add export / roster (.xlsx)…<input id="mos-roster" type="file" accept=".xlsx,.xlsm,.xls" multiple hidden></label>' +
          '<span class="mos-school-lbl small">School for files picked one by one</span>' +
          '<input id="mos-school-in" class="form-control form-control-sm" list="mos-dl-school" placeholder="e.g. Demo High School" autocomplete="off">' +
          '<datalist id="mos-dl-school"></datalist>' +
          '<button id="mos-clear" class="btn btn-outline-secondary btn-sm" type="button">Clear</button>' +
        '</div>' +
        '<div id="mos-drop" class="mos-drop small text-muted">or drop PDFs, a Certiport export or a roster workbook here</div>' +
        '<div id="mos-status" class="small mt-1" role="status" aria-live="polite"></div>' +
      '</div>' +
      '<div id="mos-empty" class="card-box mos-empty"></div>' +
      '<div id="mos-main" class="d-none">' +
        '<div class="card-box mos-filters" role="search" aria-label="MOS filters">' +
          '<div class="mos-frow">' +
            fsel("mos-f-school", "School") + fsel("mos-f-section", "Section") + fsel("mos-f-exam", "Exam") + fsel("mos-f-examno", "Exam No. / batch") + fsel("mos-f-proctor", "Proctor") +
            '<label class="mos-fl"><span>Result</span><select id="mos-f-outcome" class="form-select form-select-sm"><option value="">Pass &amp; fail</option><option value="PASS">Pass</option><option value="FAIL">Fail</option><option value="IN PROGRESS">In progress</option></select></label>' +
            '<label class="mos-fl mos-fl-score"><span>Score</span><span class="mos-range"><input id="mos-f-min" type="number" min="0" max="1000" step="1" class="form-control form-control-sm" placeholder="0" aria-label="Lowest score"> – <input id="mos-f-max" type="number" min="0" max="1000" step="1" class="form-control form-control-sm" placeholder="1000" aria-label="Highest score"></span></label>' +
          '</div><div class="mos-frow">' +
            '<label class="mos-fl"><span>Period</span><select id="mos-f-pmode" class="form-select form-select-sm"><option value="all">All time</option><option value="sy">School year</option><option value="year">Year</option><option value="month">Month</option></select></label>' +
            fsel("mos-f-pval", "&nbsp;") +
            '<label class="mos-fl mos-fl-q"><span>Student</span><input id="mos-f-q" type="search" class="form-control form-control-sm" placeholder="Name, student no. or Certiport ID" autocomplete="off"></label>' +
            '<label class="mos-fl"><span>Sort</span><select id="mos-sort" class="form-select form-select-sm">' +
              '<option value="score-desc">Score: highest first</option><option value="score-asc">Score: lowest first</option>' +
              '<option value="name">Name A–Z</option><option value="section">Section</option><option value="date-desc">Date: newest first</option></select></label>' +
            '<label class="mos-chk small"><input id="mos-f-first" type="checkbox"> First attempts only</label>' +
            '<button id="mos-reset" class="btn btn-outline-secondary btn-sm" type="button">Reset filters</button>' +
            '<button id="mos-export" class="btn btn-outline-primary btn-sm" type="button" title="Download the filtered results as an Excel workbook (no emails or addresses)">Export (.xlsx)</button>' +
            '<button id="mos-summary" class="btn btn-primary btn-sm" type="button" title="Printable Exam Summary for the current filters (company + school logos). Choose “Save as PDF” in the print dialog.">Exam Summary (PDF)</button>' +
            '<button id="mos-profile" class="btn btn-outline-secondary btn-sm" type="button" title="School logo, name, address and acronym, plus the company logos, used on the Exam Summary (saved in this browser)">School profile…</button>' +
          '</div>' +
        '</div>' +
        '<div id="mos-kpi" class="row row-cols-2 row-cols-md-4 row-cols-xl-8 g-3 kpi-row mt-1"></div>' +
        '<div class="row g-3 mt-1">' +
          '<div class="col-lg-7"><div class="card-box"><div class="mos-ch-head"><h6 class="mb-0">Pass rate over time</h6>' +
            '<div class="btn-group btn-group-sm" role="group" aria-label="Trend granularity">' +
            '<button type="button" class="btn btn-outline-secondary" data-mos-trend="month">Monthly</button>' +
            '<button type="button" class="btn btn-outline-secondary" data-mos-trend="year">Yearly</button>' +
            '<button type="button" class="btn btn-outline-secondary" data-mos-trend="sy">School year</button></div></div>' +
            '<div class="chart-wrap"><canvas id="mos-ch-trend" aria-label="Pass rate over time"></canvas></div></div></div>' +
          '<div class="col-lg-5"><div class="card-box"><h6 id="mos-ch-group-h">Pass rate by section</h6><div class="chart-wrap"><canvas id="mos-ch-group"></canvas></div></div></div>' +
          '<div class="col-lg-5"><div class="card-box"><h6>Score distribution</h6><div class="chart-wrap"><canvas id="mos-ch-bands"></canvas></div><div class="small text-muted mt-1">Passing score is 700 on every MOS exam.</div></div></div>' +
          '<div class="col-lg-7"><div class="card-box"><h6>Weakest skill areas</h6><div class="chart-wrap"><canvas id="mos-ch-skills"></canvas></div><div class="small text-muted mt-1">Average section-analysis score per exam objective (lowest first).</div></div></div>' +
        '</div>' +
        '<div class="card-box mt-3"><div class="mos-ch-head"><h6 class="mb-0" id="mos-sum-h">Summary by section</h6>' +
          '<label class="small mb-0">Group by <select id="mos-groupby" class="form-select form-select-sm d-inline-block w-auto ms-1">' +
          '<option value="auto">Section (or school)</option><option value="school">School</option><option value="section">Section</option>' +
          '<option value="proctor">Proctor</option><option value="exam">Exam</option><option value="batch">Exam No. / batch</option></select></label></div><button type="button" class="btn btn-link btn-sm p-0 mos-colreset d-none" data-mos-colreset="mos-sum">Reset columns</button><div class="table-responsive"><table class="table table-sm mos-tbl" id="mos-sum" data-resizable="mos"></table></div></div>' +
        '<div class="card-box mt-3"><div class="mos-ch-head"><h6 class="mb-0" id="mos-res-h">Results</h6><span class="small text-muted">Click a student to see every attempt.</span></div>' +
          '<button type="button" class="btn btn-link btn-sm p-0 mos-colreset d-none" data-mos-colreset="mos-res">Reset columns</button><div class="table-responsive"><table class="table table-sm mos-tbl mos-res" id="mos-res" data-resizable="mos"></table></div><div id="mos-more" class="mt-2"></div></div>' +
        '<details class="card-box mt-3 mos-check" id="mos-check"><summary><b>Check &amp; fix</b> <span id="mos-check-n" class="small"></span></summary><div id="mos-check-body" class="mt-2"></div></details>' +
      '</div>';
    wire(host);
  }
  function fsel(id, label) { return '<label class="mos-fl"><span>' + label + '</span><select id="' + id + '" class="form-select form-select-sm"></select></label>'; }

  function wire(host) {
    $("mos-pdf").addEventListener("change", function (e) { addPdfs(e.target.files).then(function () { e.target.value = ""; }); });
    $("mos-dir").addEventListener("change", function (e) {
      var fl = Array.prototype.slice.call(e.target.files || []);
      var pdfs = fl.filter(function (f) { return /\.pdf$/i.test(f.name); });
      var xl = fl.filter(function (f) { return /\.xls[xm]?$/i.test(f.name) && !/^~\$/.test(f.name); });
      addPdfs(pdfs).then(function () { return addRoster(xl); }).then(function () { e.target.value = ""; });
    });
    $("mos-roster").addEventListener("change", function (e) { addRoster(e.target.files).then(function () { e.target.value = ""; }); });
    $("mos-clear").addEventListener("click", function () {
      if (!S.attempts.length && !S.roster.length) return;
      if (window.confirm("Clear all loaded MOS results and rosters from this tab?")) clearAll();
    });
    var drop = host.querySelector(".mos-load");
    ["dragover", "dragenter"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); drop.classList.add("mos-over"); }); });
    ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function () { drop.classList.remove("mos-over"); }); });
    drop.addEventListener("drop", function (e) {
      e.preventDefault(); e.stopPropagation();
      var fl = Array.prototype.slice.call((e.dataTransfer && e.dataTransfer.files) || []);
      addPdfs(fl).then(function () { return addRoster(fl.filter(function (f) { return /\.xls[xm]?$/i.test(f.name); })); });
    });
    function bind(id, key, ev, fn) {
      $(id).addEventListener(ev || "change", function (e) {
        S.f[key] = e.target.type === "checkbox" ? e.target.checked : e.target.value;
        if (fn) fn();
        S.limit = 200; update();
      });
    }
    bind("mos-f-school", "school", null, function () { S.f.section = ""; S.f.examNo = ""; });
    bind("mos-f-section", "section"); bind("mos-f-exam", "exam"); bind("mos-f-examno", "examNo"); bind("mos-f-proctor", "proctor");
    bind("mos-f-outcome", "outcome"); bind("mos-f-min", "min", "input"); bind("mos-f-max", "max", "input");
    bind("mos-f-pmode", "pmode", null, function () { S.f.pval = ""; });
    bind("mos-f-pval", "pval"); bind("mos-f-q", "q", "input"); bind("mos-f-first", "firstOnly");
    $("mos-sort").addEventListener("change", function (e) { S.sort = e.target.value; S.limit = 200; renderResults(rows()); });
    $("mos-reset").addEventListener("click", function () {
      S.f = { school: "", section: "", exam: "", examNo: "", proctor: "", outcome: "", min: "", max: "", pmode: "all", pval: "", q: "", firstOnly: false };
      S.limit = 200; update();
    });
    $("mos-export").addEventListener("click", exportXlsx);
    $("mos-summary").addEventListener("click", printSummary);
    $("mos-profile").addEventListener("click", openProfile);
    $("mos-groupby").addEventListener("change", function (e) { S.groupBy = e.target.value; update(); });
    host.querySelectorAll("[data-mos-colreset]").forEach(function (b) { b.addEventListener("click", function () { resetCols(b.getAttribute("data-mos-colreset")); }); });
    host.querySelectorAll("[data-mos-trend]").forEach(function (b) {
      b.addEventListener("click", function () { S.trend = b.getAttribute("data-mos-trend"); update(); });
    });
    $("mos-res").addEventListener("click", function (e) {
      var tr = e.target.closest("tr[data-sk]"); if (!tr) return;
      var k = tr.getAttribute("data-sk"); S.open[k] = !S.open[k]; renderResults(rows());
    });
    $("mos-res").addEventListener("keydown", function (e) {
      if (e.key !== "Enter" && e.key !== " ") return;
      var tr = e.target.closest("tr[data-sk]"); if (!tr) return;
      e.preventDefault(); var k = tr.getAttribute("data-sk"); S.open[k] = !S.open[k]; renderResults(rows());
      var again = $("mos-res").querySelector('tr[data-sk="' + (window.CSS && CSS.escape ? CSS.escape(k) : k.replace(/["\\]/g, "\\$&")) + '"]'); if (again) again.focus();
    });
    $("mos-more").addEventListener("click", function (e) { if (e.target.closest("button")) { S.limit += 500; renderResults(rows()); } });
    $("mos-check-body").addEventListener("change", onCheckEdit);
  }

  function status() {
    var el = $("mos-status"); if (!el) return;
    if (S.progress) { el.innerHTML = '<span class="spinner-border spinner-border-sm me-1" aria-hidden="true"></span>' + esc(S.progress); return; }
    if (!S.files.length && !S.rosterFiles.length && !S.exports.length) { el.textContent = ""; return; }
    var pages = 0, att = 0; S.files.forEach(function (f) { pages += f.pages; att += f.attempts; });
    var t = S.files.length ? S.files.length + " PDF" + (S.files.length === 1 ? "" : "s") + " · " + num(att) + " score pages" : "No PDFs";
    if (S.model) t += " · " + num(S.model.rows.length) + " unique attempts" + (S.model.duplicates ? " (" + num(S.model.duplicates) + " duplicate pages ignored)" : "");
    if (S.exports.length) { var ea = 0; S.exports.forEach(function (x) { ea += x.attempts; }); t += " · " + S.exports.length + " Certiport export" + (S.exports.length === 1 ? "" : "s") + " (" + num(ea) + " rows)"; }
    if (S.model && S.model.merged) t += " · " + num(S.model.merged) + " matched PDF ↔ export";
    if (S.model && S.model.pending) t += " · " + num(S.model.pending) + " in progress (no score yet)";
    if (S.roster.length) t += " · roster: " + num(S.roster.length) + " students";
    el.textContent = t;
  }

  /* ---------------- render ---------------- */
  function render() {
    var host = $("mos-body"); if (!host) return;
    if (!on()) { host.innerHTML = '<div class="card-box small text-muted">The MOS tab is turned off for this site.</div>'; S.built = false; return; }
    if (!S.built || !$("mos-main")) { shell(host); S.built = true; }
    var dl = $("mos-dl-school");
    if (dl) dl.innerHTML = uniq(S.attempts.map(function (a) { return a.school; }).filter(function (s) { return s !== "Unassigned school"; }))
      .map(function (s) { return '<option value="' + esc(s) + '">'; }).join("");
    status();
    var has = !!(S.model && S.model.rows.length);
    $("mos-main").classList.toggle("d-none", !has);
    $("mos-empty").classList.toggle("d-none", has);
    if (!has) { renderEmpty(); destroyCharts(); return; }
    update();
  }

  function renderEmpty() {
    $("mos-empty").innerHTML =
      '<h6>How it works</h6><ol class="small mb-2">' +
      '<li><b>Add score reports</b> — the PDFs exported from Certiport (one page per candidate). Pick a school folder with <b>Add folder…</b> to read every PDF inside it; folders named like <i>“01 - School name”</i> become the school.</li>' +
      '<li><b>Add the Certiport export</b> (Results … Plus .xlsx) — every exam taker with exam, date, score, proctor and station. Rows with the same Exam reference # as a PDF page are merged (the PDF adds the skill areas). Date of birth and voucher codes are never read.</li>' +
      '<li><b>Add the roster</b> (optional) — the school workbook with one sheet per section (11A, 11B…) and NAME / Email columns. Students are matched by email first, then by name.</li>' +
      '<li><b>Filter</b> by school, section, student, score, pass/fail, Exam No. and period, then export.</li></ol>' +
      '<div class="small text-muted">File names like <i>“Exam 1119 Results.pdf”</i> give the Exam No. (batch); <i>“Exam 1119.1”</i> or <i>“… Retake …”</i> marks a retake session. ' +
      'Certificates and summary PDFs are skipped automatically and listed under <b>Check &amp; fix</b>.</div>';
  }

  function setOpts(id, values, cur, allLabel, labelFn) {
    var el = $(id); if (!el) return cur;
    if (cur && values.indexOf(cur) < 0) cur = "";
    el.innerHTML = '<option value="">' + esc(allLabel) + "</option>" + values.map(function (v) {
      return '<option value="' + esc(v) + '"' + (v === cur ? " selected" : "") + ">" + esc(labelFn ? labelFn(v) : v) + "</option>";
    }).join("");
    return cur;
  }

  function update() {
    if (!S.model) return;
    var all = S.model.rows, f = S.f;
    f.school = setOpts("mos-f-school", uniq(all.map(function (r) { return r.school; })), f.school, "All schools");
    var inSchool = f.school ? all.filter(function (r) { return r.school === f.school; }) : all;
    f.section = setOpts("mos-f-section", uniq(inSchool.map(function (r) { return r.section || "(none)"; })), f.section, "All sections");
    f.exam = setOpts("mos-f-exam", uniq(inSchool.map(function (r) { return r.exam; })), f.exam, "All exams");
    var nos = uniq(inSchool.map(function (r) { return r.batch || "(none)"; })).sort(function (a, b) { return (+b || 0) - (+a || 0) || b.localeCompare(a); });
    f.examNo = setOpts("mos-f-examno", nos, f.examNo, "All exam numbers / batches", function (n) { return /^\d+$/.test(n) ? "Exam " + n : n; });
    var pr = uniq(inSchool.map(function (r) { return r.proctor || "(none)"; }));
    f.proctor = setOpts("mos-f-proctor", pr, f.proctor, "All proctors");
    $("mos-f-proctor").closest(".mos-fl").classList.toggle("d-none", pr.length < 2 && !f.proctor && pr[0] === "(none)");
    var pv = [], pl = null;
    if (f.pmode === "sy") pv = uniq(inSchool.map(function (r) { return r.sy; })).reverse();
    else if (f.pmode === "year") pv = uniq(inSchool.map(function (r) { return r.year; })).reverse();
    else if (f.pmode === "month") { pv = uniq(inSchool.map(function (r) { return r.month; })).reverse(); pl = monLabel; }
    var pvEl = $("mos-f-pval"); pvEl.disabled = f.pmode === "all";
    f.pval = setOpts("mos-f-pval", pv, f.pval, f.pmode === "all" ? "—" : "All", pl);
    $("mos-f-outcome").value = f.outcome; $("mos-f-pmode").value = f.pmode; $("mos-sort").value = S.sort;
    $("mos-f-first").checked = !!f.firstOnly;
    if ($("mos-f-min").value !== String(f.min)) $("mos-f-min").value = f.min;
    if ($("mos-f-max").value !== String(f.max)) $("mos-f-max").value = f.max;
    if ($("mos-f-q").value !== f.q) $("mos-f-q").value = f.q;
    document.querySelectorAll("[data-mos-trend]").forEach(function (b) {
      var a = b.getAttribute("data-mos-trend") === S.trend; b.classList.toggle("active", a); b.setAttribute("aria-pressed", String(a));
    });
    var R = rows();
    renderKpis(R); renderCharts(R); renderSummary(R); renderResults(R); renderCheck();
  }

  function kpi(label, val, tone, title) {
    return '<div class="col"><div class="kpi kpi-' + tone + '"' + (title ? ' title="' + esc(title) + '"' : "") + '><div class="kpi-val">' + val + '</div><div class="kpi-lbl">' + esc(label) + "</div></div></div>";
  }
  function rateTone(v) { return v == null ? "blue" : v >= 0.75 ? "green" : v >= 0.5 ? "orange" : "red"; }
  function renderKpis(R) {
    var s = P().summarize(R);
    $("mos-kpi").innerHTML =
      kpi("Attempts", num(s.attempts), "blue") + kpi("Students", num(s.students), "blue") +
      kpi("Passed", num(s.passed), "green") + kpi("Failed", num(s.failed), s.failed ? "red" : "green") +
      kpi("Pass rate", pct(s.passRate), rateTone(s.passRate)) +
      kpi("First-attempt pass rate", pct(s.firstPassRate), rateTone(s.firstPassRate), "Pass rate counting each student's first try at each exam only (retakes excluded)") +
      kpi("Average score", num(s.avg), s.avg >= 700 ? "green" : "orange") +
      kpi("Highest / lowest", (s.high == null ? "—" : s.high) + " / " + (s.low == null ? "—" : s.low), "blue");
  }

  /* ---------------- charts ---------------- */
  function destroyCharts() { Object.keys(S.charts).forEach(function (k) { try { S.charts[k].destroy(); } catch (e) {} }); S.charts = {}; }
  function chart(id, cfg) {
    if (typeof Chart === "undefined") return;
    var cv = $(id); if (!cv) return;
    if (S.charts[id]) { try { S.charts[id].destroy(); } catch (e) {} }
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false }, cfg.options || {});
    S.charts[id] = new Chart(cv, cfg);
  }
  function renderCharts(R) {
    var ok = cssVar("--ok-fill", "#107C10"), bad = cssVar("--bad-fill", "#D13438"), warn = cssVar("--warn-fill", "#F7630C"), brand = cssVar("--brand", "#0F6CBD");
    var P_ = P();
    // 1) trend
    var keyFn = S.trend === "year" ? function (r) { return r.year; } : S.trend === "sy" ? function (r) { return r.sy; } : function (r) { return r.month; };
    var tg = P_.groupRate(R, keyFn);
    chart("mos-ch-trend", {
      type: "bar",
      data: {
        labels: tg.map(function (g) { return S.trend === "month" ? monLabel(g.key) : g.key; }),
        datasets: [
          { type: "line", label: "Pass rate %", data: tg.map(function (g) { return +(g.passRate * 100).toFixed(1); }), borderColor: ok, backgroundColor: ok, yAxisID: "y", tension: 0.25, pointRadius: 3 },
          { type: "bar", label: "Attempts", data: tg.map(function (g) { return g.attempts; }), backgroundColor: brand + "55", borderColor: brand, borderWidth: 1, yAxisID: "y2" }
        ]
      },
      options: {
        interaction: { mode: "index", intersect: false },
        scales: { y: { min: 0, max: 100, title: { display: true, text: "Pass rate %" } }, y2: { position: "right", beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: "Attempts" } } },
        plugins: { tooltip: { callbacks: { afterBody: function (it) { var g = tg[it[0].dataIndex]; return g ? g.passed + " passed of " + g.attempts : ""; } } } }
      }
    });
    // 2) by section (one school) or by school
    var G = groupDef();
    $("mos-ch-group-h").textContent = "Pass rate by " + G.label;
    var gg = P_.groupRate(R, G.key);
    chart("mos-ch-group", {
      type: "bar",
      data: { labels: gg.map(function (g) { return g.key; }), datasets: [{ label: "Pass rate %", data: gg.map(function (g) { return +(g.passRate * 100).toFixed(1); }), backgroundColor: gg.map(function (g) { return g.passRate >= 0.75 ? ok : g.passRate >= 0.5 ? warn : bad; }) }] },
      options: {
        indexAxis: "y", scales: { x: { min: 0, max: 100 } }, plugins: { legend: { display: false },
          tooltip: { callbacks: { label: function (c) { var g = gg[c.dataIndex]; return " " + c.parsed.x + "% · " + g.passed + "/" + g.attempts + " passed · avg " + Math.round(g.avg); } } } }
      }
    });
    // 3) score bands
    var bc = {}; P_.BANDS.forEach(function (b) { bc[b] = 0; }); R.forEach(function (r) { if (bc[r.band] != null) bc[r.band]++; });
    chart("mos-ch-bands", {
      type: "bar",
      data: { labels: P_.BANDS, datasets: [{ label: "Attempts", data: P_.BANDS.map(function (b) { return bc[b]; }), backgroundColor: [bad, warn, brand, ok] }] },
      options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }
    });
    // 4) weakest skill areas
    var sk = P_.skillAreas(R).slice(0, 8);
    chart("mos-ch-skills", {
      type: "bar",
      data: { labels: sk.map(function (s) { var l = s.exam + " · " + s.area; return l.length > 58 ? l.slice(0, 56) + "…" : l; }), datasets: [{ label: "Average %", data: sk.map(function (s) { return +s.avg.toFixed(1); }), backgroundColor: sk.map(function (s) { return s.avg < 50 ? bad : s.avg < 70 ? warn : ok; }) }] },
      options: { indexAxis: "y", scales: { x: { min: 0, max: 100 } }, plugins: { legend: { display: false },
        tooltip: { callbacks: { label: function (c) { var s = sk[c.dataIndex]; return " " + s.avg.toFixed(0) + "% average · " + s.below50 + " of " + s.n + " below 50%"; } } } } }
    });
  }

  /* ---------------- group by (summary table + comparison chart) ---------------- */
  var GROUPS = {
    section: { label: "section", head: "Section", key: function (r) { return r.section || "(none)"; }, filter: "section" },
    school: { label: "school", head: "School", key: function (r) { return r.school; }, filter: "school" },
    proctor: { label: "proctor", head: "Proctor", key: function (r) { return r.proctor || "(none)"; }, filter: "proctor" },
    exam: { label: "exam", head: "Exam", key: function (r) { return r.exam; }, filter: "exam" },
    batch: { label: "Exam No. / batch", head: "Exam No. / batch", key: function (r) { return r.batch || "(none)"; }, filter: "examNo" }
  };
  function groupDef() { return GROUPS[S.groupBy] || (S.f.school ? GROUPS.section : GROUPS.school); }

  /* ---------------- resizable columns (MOS tables) ----------------
   * The app-wide makeTablesResizable() marks a table once and adds grips to its FIRST thead; MOS tables are
   * re-rendered (innerHTML) on every filter change, so the grips vanished while table-layout:fixed stayed →
   * nowrap cells spilled over their neighbours (overlapping text). MOS tables opt out (data-resizable="mos")
   * and get grips after every render here. No widths set → automatic layout (nothing can overlap; the
   * table scrolls sideways). After a drag → fixed layout with every column width kept, long text cut with
   * "…" and shown in full as a tooltip. Widths are a layout preference (localStorage), never student data.
   * Double-click a grip = auto width for that column; "Reset columns" = automatic layout again. */
  var COLW_KEY = "qbr-mos-colw-v1";
  function colWidths() { try { return JSON.parse(localStorage.getItem(COLW_KEY) || "{}") || {}; } catch (e) { return {}; } }
  function saveColWidths(all) { try { localStorage.setItem(COLW_KEY, JSON.stringify(all)); } catch (e) {} }
  function mosCols(id) {
    var tbl = $(id); if (!tbl) return;
    var ths = Array.prototype.slice.call(tbl.querySelectorAll(":scope > thead > tr > th"));
    if (!ths.length) return;
    var key = id + ":" + ths.length, all = colWidths(), w = all[key] || null;
    function apply() {
      var cg = tbl.querySelector(":scope > colgroup"); if (cg) cg.parentNode.removeChild(cg);
      tbl.classList.toggle("mos-fixed", !!w);
      if (w) {
        cg = document.createElement("colgroup");
        ths.forEach(function (th, i) { var c = document.createElement("col"); if (w[i]) c.style.width = w[i] + "px"; cg.appendChild(c); });
        tbl.insertBefore(cg, tbl.firstChild);
        tbl.style.width = w.reduce(function (a, b) { return a + (b || 0); }, 0) + "px";
        // full text as a tooltip wherever "…" cut it
        tbl.querySelectorAll(":scope > tbody > tr > td").forEach(function (td) { if (!td.title && td.scrollWidth > td.clientWidth + 1) td.title = td.textContent.trim(); });
      } else tbl.style.width = "";
    }
    function start(i, x) {
      if (!w) w = ths.map(function (th) { return Math.round(th.getBoundingClientRect().width); });
      return { i: i, x: x, w0: w[i] };
    }
    function set(i, px) { w[i] = Math.max(44, Math.round(px)); all[key] = w; saveColWidths(all); apply(); }
    ths.forEach(function (th, i) {
      var g = document.createElement("span");
      g.className = "mos-grip"; g.setAttribute("role", "separator"); g.setAttribute("aria-orientation", "vertical");
      g.setAttribute("aria-label", "Resize column " + th.textContent.trim()); g.tabIndex = 0;
      g.title = "Drag to resize · double-click: auto width · arrow keys: ±16 px";
      g.addEventListener("click", function (e) { e.stopPropagation(); e.preventDefault(); });
      g.addEventListener("mousedown", function (e) {
        e.preventDefault(); e.stopPropagation();
        var st = start(i, e.clientX); g.classList.add("dragging"); document.body.classList.add("mos-resizing");
        function mv(ev) { set(st.i, st.w0 + ev.clientX - st.x); }
        function up() { g.classList.remove("dragging"); document.body.classList.remove("mos-resizing"); document.removeEventListener("mousemove", mv); document.removeEventListener("mouseup", up); }
        document.addEventListener("mousemove", mv); document.addEventListener("mouseup", up);
      });
      g.addEventListener("keydown", function (e) {
        if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
        e.preventDefault(); var st = start(i, 0); set(i, st.w0 + (e.key === "ArrowRight" ? 16 : -16));
      });
      g.addEventListener("dblclick", function (e) {
        e.stopPropagation();
        if (!w) return;
        tbl.classList.remove("mos-fixed"); tbl.style.width = ""; var cg = tbl.querySelector(":scope > colgroup"); if (cg) cg.parentNode.removeChild(cg);
        var nat = ths.map(function (h) { return Math.round(h.getBoundingClientRect().width); });   // natural width of every column
        w[i] = nat[i]; all[key] = w; saveColWidths(all); apply();
      });
      th.appendChild(g);
    });
    var rs = tbl.parentNode && tbl.parentNode.parentNode ? tbl.parentNode.parentNode.querySelector('[data-mos-colreset="' + id + '"]') : null;
    if (rs) rs.classList.toggle("d-none", !w);
    apply();
  }
  function resetCols(id) { var all = colWidths(); Object.keys(all).forEach(function (k) { if (k.indexOf(id + ":") === 0) delete all[k]; }); saveColWidths(all); update(); }

  /* ---------------- tables ---------------- */
  function rateCell(v) { return '<td class="num"><span class="mos-rate mos-rate-' + rateTone(v) + '">' + pct(v) + "</span></td>"; }
  function renderSummary(R) {
    var G = groupDef();
    $("mos-sum-h").textContent = "Summary by " + G.label + (S.f.school ? " — " + S.f.school : "");
    var gb = $("mos-groupby"); if (gb) gb.value = S.groupBy;
    var g = {};
    R.forEach(function (r) {
      if (r.pending) return;
      var k = G.key(r);
      var o = g[k] || (g[k] = { key: k, n: 0, pass: 0, sum: 0, hi: null, lo: null, st: {}, first: 0, firstPass: 0 });
      o.n++; if (r.pass) o.pass++; o.sum += r.score; o.st[r.studentKey] = 1;
      if (o.hi == null || r.score > o.hi) o.hi = r.score; if (o.lo == null || r.score < o.lo) o.lo = r.score;
      if (r.attemptNo === 1) { o.first++; if (r.pass) o.firstPass++; }
    });
    var list = Object.keys(g).sort().map(function (k) { return g[k]; });
    var tot = P().summarize(R);
    $("mos-sum").innerHTML = "<thead><tr><th>" + esc(G.head) + '</th><th class="num">Students</th><th class="num">Attempts</th><th class="num">Passed</th><th class="num">Failed</th><th class="num">Pass rate</th><th class="num">First-attempt</th><th class="num">Average</th><th class="num">Highest</th><th class="num">Lowest</th></tr></thead><tbody>' +
      list.map(function (o) {
        return '<tr><td><button type="button" class="btn btn-link btn-sm p-0 mos-drill" data-k="' + esc(o.key) + '">' + esc(o.key) + "</button></td>" +
          '<td class="num">' + num(Object.keys(o.st).length) + '</td><td class="num">' + num(o.n) + '</td><td class="num">' + num(o.pass) + '</td><td class="num">' + num(o.n - o.pass) + "</td>" +
          rateCell(o.n ? o.pass / o.n : null) + rateCell(o.first ? o.firstPass / o.first : null) +
          '<td class="num">' + num(o.sum / o.n) + '</td><td class="num">' + o.hi + '</td><td class="num">' + o.lo + "</td></tr>";
      }).join("") + "</tbody>" +
      '<tfoot><tr><th>Total</th><th class="num">' + num(tot.students) + '</th><th class="num">' + num(tot.attempts) + '</th><th class="num">' + num(tot.passed) + '</th><th class="num">' + num(tot.failed) + "</th>" +
      rateCell(tot.passRate).replace("<td", "<th").replace("</td>", "</th>") + rateCell(tot.firstPassRate).replace("<td", "<th").replace("</td>", "</th>") +
      '<th class="num">' + num(tot.avg) + '</th><th class="num">' + (tot.high == null ? "—" : tot.high) + '</th><th class="num">' + (tot.low == null ? "—" : tot.low) + "</th></tr></tfoot>";
    $("mos-sum").querySelectorAll(".mos-drill").forEach(function (b) {
      b.addEventListener("click", function () {
        var k = b.getAttribute("data-k");
        if (G.filter === "school") { S.f.school = k; S.f.section = ""; S.f.examNo = ""; S.f.proctor = ""; if (S.groupBy === "school") S.groupBy = "auto"; }
        else S.f[G.filter] = k;
        S.limit = 200; update();
      });
    });
    mosCols("mos-sum");
  }

  function sc(r, dflt) { return r.score == null ? dflt : r.score; }   // in-progress rows sort last
  function sorted(R) {
    var a = R.slice(), nm = function (r) { return (r.rosterName || r.name || "").toLowerCase(); };
    var cmp = {
      "score-desc": function (x, y) { return sc(y, -1) - sc(x, -1) || nm(x).localeCompare(nm(y)); },
      "score-asc": function (x, y) { return sc(x, 9999) - sc(y, 9999) || nm(x).localeCompare(nm(y)); },
      "name": function (x, y) { return nm(x).localeCompare(nm(y)) || (x.date || "").localeCompare(y.date || ""); },
      "section": function (x, y) { return (x.section || "~").localeCompare(y.section || "~") || nm(x).localeCompare(nm(y)); },
      "date-desc": function (x, y) { return (y.date || "").localeCompare(x.date || "") || nm(x).localeCompare(nm(y)); }
    }[S.sort] || function () { return 0; };
    return a.sort(cmp);
  }
  function skillChips(r) {
    return (r.skills || []).map(function (s) {
      return '<span class="mos-sk' + (s.pct < 50 ? " mos-sk-low" : "") + '" title="' + esc(s.area) + '">' + s.pct + "%</span>";
    }).join("");
  }
  function outcomeCell(r) {
    if (r.pending) return '<td><span class="mos-out mos-out-wait">In progress</span></td>';
    return '<td><span class="mos-out mos-out-' + (r.pass ? "pass" : "fail") + '">' + (r.pass ? "Pass" : "Fail") + "</span></td>";
  }
  function mins(sec) { return sec == null ? "" : Math.floor(sec / 60) + "m " + String(sec % 60).padStart(2, "0") + "s"; }
  function renderResults(R) {
    var list = sorted(R), shown = list.slice(0, S.limit);
    $("mos-res-h").textContent = "Results (" + num(list.length) + (list.length === 1 ? " attempt" : " attempts") + ")";
    var all = S.model ? S.model.rows : [];
    var html = '<thead><tr><th>Student</th><th>School</th><th>Section</th><th>Exam</th><th>Exam No. / batch</th><th>Date</th><th class="num">Score</th><th>Result</th><th class="num">Attempt</th><th>Proctor</th><th>Skill areas</th></tr></thead><tbody>';
    var histDone = {};
    shown.forEach(function (r) {
      var open = !!S.open[r.studentKey];
      html += '<tr data-sk="' + esc(r.studentKey) + '" tabindex="0" aria-expanded="' + open + '" class="mos-row' + (open ? " mos-open" : "") + '">' +
        '<td><span class="mos-name">' + esc(r.rosterName || r.name) + "</span>" + (r.studentNo ? ' <span class="small text-muted">' + esc(r.studentNo) + "</span>" : "") + "</td>" +
        "<td>" + esc(r.school) + "</td><td>" + esc(r.section || "—") + "</td><td>" + esc(r.exam) + "</td><td>" + esc(r.examNo || r.batch || "—") + "</td>" +
        "<td>" + fmtDate(r.date) + '</td><td class="num"><b>' + (r.score == null ? "—" : r.score) + "</b></td>" + outcomeCell(r) +
        '<td class="num">' + r.attemptNo + (r.isRetake ? ' <span class="mos-retake" title="Retake">R</span>' : "") + "</td>" +
        "<td>" + esc(r.proctor || "—") + "</td>" +
        '<td class="mos-sks">' + skillChips(r) + "</td></tr>";
      if (open && !histDone[r.studentKey]) {
        histDone[r.studentKey] = 1;
        var hist = all.filter(function (x) { return x.studentKey === r.studentKey; }).sort(function (a, b) { return (a.date || "").localeCompare(b.date || ""); });
        html += '<tr class="mos-hist"><td colspan="11"><div class="small"><b>All attempts</b> · ' + esc(r.email || "") + (r.certiportId ? " · Certiport ID " + esc(r.certiportId) : "") + "</div>" +
          '<table class="table table-sm mb-0" data-resizable="static"><tbody>' + hist.map(function (h) {
            return "<tr><td>" + fmtDate(h.date) + "</td><td>" + esc(h.exam) + (h.examVersion ? ' <span class="small text-muted">' + esc(h.examVersion) + "</span>" : "") + "</td><td>" + esc(h.examNo ? "Exam " + h.examNo : (h.batch || "—")) + '</td><td class="num"><b>' + (h.score == null ? "—" : h.score) + "</b></td>" + outcomeCell(h) +
              '<td class="small">' + esc([h.proctor ? "Proctor " + h.proctor : "", h.station ? "Station " + h.station : "", h.timeUsed != null ? mins(h.timeUsed) : ""].filter(Boolean).join(" · ")) + "</td>" +
              "<td>" + (h.skills || []).map(function (s) { return '<span class="mos-sk' + (s.pct < 50 ? " mos-sk-low" : "") + '">' + esc(s.area) + " " + s.pct + "%</span>"; }).join(" ") + "</td></tr>";
          }).join("") + "</tbody></table></td></tr>";
      }
    });
    if (!shown.length) html += '<tr><td colspan="11" class="text-muted small">No attempts match these filters.</td></tr>';
    $("mos-res").innerHTML = html + "</tbody>";
    $("mos-more").innerHTML = list.length > shown.length
      ? '<button type="button" class="btn btn-outline-secondary btn-sm">Show more (' + num(list.length - shown.length) + " not shown)</button>" : "";
    mosCols("mos-res");
  }

  /* ---------------- check & fix ---------------- */
  function renderCheck() {
    var m = S.model; if (!m) return;
    var rs = {}; S.roster.forEach(function (st) { rs[st.school] = 1; });   // only schools that have a roster need mapping
    // export rows still on the Test Center (no roster match yet) can belong to any loaded roster
    var needs = m.rows.filter(function (r) { return (rs[r.school] || (r.schoolFromCenter && S.roster.length)) && (r.match === "none" || r.match === "ambiguous"); });
    var noScore = S.files.filter(function (f) { return !f.attempts; });
    var chk = [needs.length ? needs.length + " to map" : "", m.pending ? m.pending + " in progress" : "",
      noScore.length ? noScore.length + " file" + (noScore.length === 1 ? "" : "s") + " without score pages" : ""].filter(Boolean).join(" · ");
    $("mos-check-n").textContent = chk ? "(" + chk + ")" : "";
    var html = "";
    if (!S.roster.length) html += '<div class="small mb-2">No roster loaded — sections come from the file name (e.g. <i>“… 11A …”</i>) when present. Add the school roster to map every student to a section.</div>';
    else {
      if (S.rosterFiles.some(function (r) { return r.students && r.school === "Unassigned school"; }))
        html += '<div class="small mb-1 text-danger">A roster was added without a school, so its students cannot be matched. Pick or type the school, then add the roster again.</div>';
      html += '<div class="small mb-1">Rosters: ' + S.rosterFiles.map(function (r) { return esc(r.name) + (r.err ? " — could not read: " + esc(r.err) : r.students ? " → " + esc(r.school) + " (" + r.students + " students)" : " — no section sheets (11A, 11B…), ignored"); }).join("; ") + "</div>";
      if (needs.length) {
        html += '<div class="small mb-1"><b>' + needs.length + " attempt" + (needs.length === 1 ? "" : "s") + "</b> could not be matched to the roster. Pick the student or type a section.</div>" +
          '<div class="table-responsive mos-check-tbl"><table class="table table-sm" data-resizable="static"><thead><tr><th>Name on score report</th><th>School</th><th>Exam / date</th><th>Roster student</th><th>Section</th></tr></thead><tbody>' +
          needs.slice(0, 300).map(function (r) {
            var pool = r.schoolFromCenter ? S.roster : S.roster.filter(function (s) { return s.school === r.school; });
            var cands = r.candidates || [];
            var opts = '<option value="">— not in roster —</option>' +
              (cands.length ? '<optgroup label="Possible matches">' + cands.map(function (s) { return '<option value="' + esc(rid(s)) + '">' + esc(s.name + " · " + s.section) + "</option>"; }).join("") + "</optgroup>" : "") +
              '<optgroup label="All students">' + pool.slice().sort(function (a, b) { return a.name.localeCompare(b.name); }).map(function (s) { return '<option value="' + esc(rid(s)) + '">' + esc(s.name + " · " + s.section) + "</option>"; }).join("") + "</optgroup>";
            return "<tr><td>" + esc(r.name) + (r.match === "ambiguous" ? ' <span class="badge bg-warning text-dark">several matches</span>' : "") + "</td><td>" + esc(r.school) + "</td><td>" + esc(r.exam) + " · " + fmtDate(r.date) + "</td>" +
              '<td><select class="form-select form-select-sm" data-mos-fix="student" data-ref="' + esc(r.key) + '">' + opts + "</select></td>" +
              '<td><input class="form-control form-control-sm" style="width:80px" data-mos-fix="section" data-ref="' + esc(r.key) + '" value="' + esc(r.section) + '" placeholder="11A"></td></tr>';
          }).join("") + "</tbody></table></div>" + (needs.length > 300 ? '<div class="small text-muted">Showing the first 300.</div>' : "");
      } else html += '<div class="small mb-2 text-success">Every attempt is matched to a roster student.</div>';
    }
    var pend = m.rows.filter(function (r) { return r.pending; });
    if (pend.length) html += '<div class="small mt-2"><b>' + pend.length + " in progress</b> (no score yet — not counted in pass rates): " +
      pend.slice(0, 20).map(function (r) { return esc(r.name) + " · " + esc(r.exam) + " · " + fmtDate(r.date); }).join("; ") + "</div>";
    if (S.exports.length) html += '<div class="small mt-2">Certiport exports: ' + S.exports.map(function (x) {
      return esc(x.name) + " (" + esc(x.sheets.join(", ")) + ": " + num(x.attempts) + " rows" + (x.skipped ? ", " + x.skipped + " footer lines ignored" : "") + ")";
    }).join("; ") + (m.merged ? " · " + num(m.merged) + " matched to PDF pages by Exam reference #" : "") + ". Date of birth and voucher codes are not read.</div>";
    html += '<details class="mt-2"><summary class="small">Files read (' + S.files.length + ")</summary>" +
      '<div class="table-responsive"><table class="table table-sm small" data-resizable="static"><thead><tr><th>File</th><th>School</th><th class="num">Pages</th><th class="num">Score pages</th><th>Note</th></tr></thead><tbody>' +
      S.files.map(function (f) {
        var note = f.err ? "Could not read: " + f.err : !f.attempts ? "No score-report pages (certificate or summary?) — skipped" : (f.skipped ? f.skipped + " other page(s) skipped" : "") + (f.bad ? " · " + f.bad + " page(s) without a score" : "");
        return "<tr" + (!f.attempts ? ' class="table-warning"' : "") + "><td>" + esc(f.file) + "</td><td>" + esc(f.school) + '</td><td class="num">' + f.pages + '</td><td class="num">' + f.attempts + "</td><td>" + esc(note) + "</td></tr>";
      }).join("") + "</tbody></table></div></details>";
    $("mos-check-body").innerHTML = html;
  }
  function rid(s) { return s.school + "|" + s.section + "|" + (s.studentNo || s.email || s.name); }
  function onCheckEdit(e) {
    var el = e.target, kind = el.getAttribute("data-mos-fix"), ref = el.getAttribute("data-ref");
    if (!kind || !ref) return;
    var o = S.overrides[ref] || (S.overrides[ref] = {});
    if (kind === "student") {
      var st = null; S.roster.forEach(function (s) { if (rid(s) === el.value) st = s; });
      o.student = st;
    } else o.section = String(el.value || "").toUpperCase().replace(/\s+/g, "");
    rebuild();
    var d = $("mos-check"); if (d) d.open = true;
  }

  /* ---------------- export ---------------- */
  function exportXlsx() {
    if (typeof XLSX === "undefined" || !S.model) return;
    var R = sorted(rows());
    var data = R.map(function (r) {
      var o = {
        School: r.school, Section: r.section, "Student no.": r.studentNo, Student: r.rosterName || r.name,
        "Name on report": r.name, Exam: r.exam, "Exam version": r.examVersion, Certification: r.certName || "",
        "Exam No.": r.examNo, Batch: r.batch, Date: r.date, Time: r.time || "", "School year": r.sy,
        Score: r.score == null ? "" : r.score, Required: r.required, Result: r.pending ? "In progress" : r.pass ? "Pass" : "Fail",
        Attempt: r.attemptNo, Retake: r.isRetake ? "Yes" : "", Proctor: r.proctor || "", Station: r.station || "",
        "Time used (min)": r.timeUsed == null ? "" : +(r.timeUsed / 60).toFixed(1), "Test center": r.testCenter || "",
        Source: r.source || "pdf", "Exam reference #": r.ref
      };
      (r.skills || []).forEach(function (s, i) { o["Skill " + (i + 1)] = s.area + " — " + s.pct + "%"; });
      return o;
    });
    var s = P().summarize(R);
    var sum = [
      { Metric: "Attempts (scored)", Value: s.attempts }, { Metric: "In progress", Value: s.pending }, { Metric: "Students", Value: s.students },
      { Metric: "Passed", Value: s.passed }, { Metric: "Failed", Value: s.failed },
      { Metric: "Pass rate", Value: s.passRate == null ? "" : +(s.passRate * 100).toFixed(1) + "%" },
      { Metric: "First-attempt pass rate", Value: s.firstPassRate == null ? "" : +(s.firstPassRate * 100).toFixed(1) + "%" },
      { Metric: "Average score", Value: s.avg == null ? "" : Math.round(s.avg) }, { Metric: "Highest", Value: s.high }, { Metric: "Lowest", Value: s.low },
      { Metric: "Filters", Value: describeFilters() }
    ];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sum), "Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), "Results");
    var tg = P().groupRate(R, function (r) { return r.month; }).map(function (g) { return { Month: g.key, Attempts: g.attempts, Passed: g.passed, "Pass rate %": +(g.passRate * 100).toFixed(1) }; });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tg), "By month");
    var yg = P().groupRate(R, function (r) { return r.year; }).map(function (g) { return { Year: g.key, Attempts: g.attempts, Passed: g.passed, "Pass rate %": +(g.passRate * 100).toFixed(1) }; });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(yg), "By year");
    var name = "MOS_Results_" + (S.f.school ? S.f.school.replace(/[^A-Za-z0-9]+/g, "_") + "_" : "") + new Date().toISOString().slice(0, 10) + ".xlsx";
    XLSX.writeFile(wb, name);
  }
  function describeFilters() {
    var f = S.f, p = [];
    if (f.school) p.push("School: " + f.school); if (f.section) p.push("Section: " + f.section);
    if (f.exam) p.push("Exam: " + f.exam); if (f.examNo) p.push("Exam No. / batch: " + f.examNo); if (f.proctor) p.push("Proctor: " + f.proctor);
    if (f.outcome) p.push("Result: " + f.outcome); if (f.min !== "" || f.max !== "") p.push("Score: " + (f.min || 0) + "–" + (f.max || 1000));
    if (f.pmode !== "all" && f.pval) p.push(f.pmode + ": " + f.pval); if (f.q) p.push("Search: " + (/@/.test(f.q) ? "(email)" : f.q)); if (f.firstOnly) p.push("First attempts only");
    return p.join(" · ") || "None";
  }

  /* ---------------- school profiles (logo, name, acronym; this browser only) ----------------
   * { [school as in the data]: { title, acronym, address, logo: dataURL } (+ company logos in qbr-mos-branding-v1) } in localStorage.
   * School images only — no student data is ever stored. */
  var PROF_KEY = "qbr-mos-school-profiles-v1", LOGO_MAX = 512 * 1024;
  function profiles() { try { return JSON.parse(localStorage.getItem(PROF_KEY) || "{}") || {}; } catch (e) { return {}; } }
  function saveProfiles(all) {
    try { localStorage.setItem(PROF_KEY, JSON.stringify(all)); return true; }
    catch (e) { window.alert("Could not save the school images in this browser (storage is full). Try a smaller image."); return false; }
  }
  function profileOf(school) { return profiles()[school] || {}; }
  function acronymMap() { var all = profiles(), o = {}; Object.keys(all).forEach(function (k) { if (all[k].acronym) o[k] = all[k].acronym; }); return o; }
  function readImage(file, max) {
    return new Promise(function (res, rej) {
      if (!file) return rej(new Error("No file"));
      if (!/^image\/(png|jpe?g|gif|webp|svg\+xml)$/i.test(file.type)) return rej(new Error("Use a PNG, JPG, GIF, WebP or SVG image."));
      if (file.size > max) return rej(new Error("Image is too large (" + Math.round(file.size / 1024) + " KB; max " + Math.round(max / 1024) + " KB)."));
      var fr = new FileReader();
      fr.onerror = function () { rej(new Error("Could not read the image.")); };
      fr.onload = function () {
        var img = new Image();
        img.onload = function () { res(fr.result); };
        img.onerror = function () { rej(new Error("That file is not a valid image.")); };
        img.src = fr.result;
      };
      fr.readAsDataURL(file);
    });
  }
  function openProfile() {
    var school = S.f.school || (S.model && uniq(S.model.rows.map(function (r) { return r.school; })).length === 1 ? S.model.rows[0].school : "");
    if (!school) { window.alert("Pick a school in the School filter first."); return; }
    var pr = profileOf(school), work = { title: pr.title || "", acronym: pr.acronym || "", address: pr.address || "", logo: pr.logo || "" };
    var br = branding(), bw = { left: br.left || "", right: br.right || "", logoH: hdrOpt(br).logoH, scale: hdrOpt(br).scale, align: hdrOpt(br).align };
    var ov = document.createElement("div");
    ov.className = "mos-ov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "mos-pf-h");
    function draw(msg) {
      ov.innerHTML = '<div class="mos-ov-box card-box">' +
        '<h5 id="mos-pf-h" class="mb-1">School profile</h5><div class="small text-muted mb-2">' + esc(school) + ' · used on the Exam Summary. Saved in this browser only.</div>' +
        '<label class="mos-fl w-100 mb-2"><span>Name on the summary</span><input id="mos-pf-title" class="form-control form-control-sm" value="' + esc(work.title) + '" placeholder="' + esc(school) + '"></label>' +
        '<div class="d-flex gap-2 mb-2"><label class="mos-fl"><span>Acronym (batch codes)</span><input id="mos-pf-acr" class="form-control form-control-sm" maxlength="8" value="' + esc(work.acronym) + '" placeholder="' + esc(P().acronym(school)) + '"></label>' +
        '<label class="mos-fl flex-grow-1"><span>Address line (optional)</span><input id="mos-pf-addr" class="form-control form-control-sm" value="' + esc(work.address) + '"></label></div>' +
        '<div class="mos-pf-imgs">' +
          '<div><div class="small fw-bold">Logo</div><div class="mos-pf-prev">' + (work.logo ? '<img id="mos-pf-logo-img" alt="School logo" src="' + esc(work.logo) + '">' : '<span class="small text-muted">None</span>') + '</div>' +
          '<label class="btn btn-outline-primary btn-sm mb-0">Upload logo…<input id="mos-pf-logo" type="file" accept="image/*" hidden></label> ' +
          (work.logo ? '<button type="button" class="btn btn-link btn-sm" data-pf-clear="logo">Remove</button>' : "") + '<div class="small text-muted">PNG/JPG/SVG, ≤ 512 KB</div></div>' +
          '<div><div class="small fw-bold">Company logos (every school)</div>' +
            '<div class="mos-pf-co"><div><div class="mos-pf-prev">' + '<img id="mos-pf-left-img" alt="Left company logo" src="' + esc(bw.left || RAKSO_DEFAULT) + '">' + '</div>' +
              '<label class="btn btn-outline-secondary btn-sm mb-0">Left logo…<input id="mos-pf-left" type="file" accept="image/*" hidden></label>' +
              (bw.left ? ' <button type="button" class="btn btn-link btn-sm" data-pf-coclear="left">Use Rakso CT</button>' : '<div class="small text-muted">Rakso CT (built in)</div>') + '</div>' +
            '<div><div class="mos-pf-prev">' + (bw.right ? '<img id="mos-pf-right-img" alt="Right partner logo" src="' + esc(bw.right) + '">' : '<span class="small text-muted">None</span>') + '</div>' +
              '<label class="btn btn-outline-primary btn-sm mb-0">Right logo (GTP)…<input id="mos-pf-right" type="file" accept="image/*" hidden></label>' +
              (bw.right ? ' <button type="button" class="btn btn-link btn-sm" data-pf-coclear="right">Remove</button>' : "") + '</div></div>' +
            '<div class="small text-muted">≤ 512 KB each.</div>' +
            '<div class="mos-pf-hdr"><div class="small fw-bold mt-2">Summary header size</div>' +
              '<label class="mos-fl"><span>Logo height <b id="mos-pf-logoh-v">' + bw.logoH + ' px</b></span><input id="mos-pf-logoh" type="range" class="form-range" min="24" max="80" step="2" value="' + bw.logoH + '"></label>' +
              '<label class="mos-fl"><span>Results box size <b id="mos-pf-scale-v">' + bw.scale + '%</b></span><input id="mos-pf-scale" type="range" class="form-range" min="60" max="120" step="5" value="' + bw.scale + '"></label>' +
              '<label class="mos-fl"><span>Logo position</span><select id="mos-pf-align" class="form-select form-select-sm">' +
                '<option value="center"' + (bw.align === "center" ? " selected" : "") + '>Centered, side by side</option>' +
                '<option value="spread"' + (bw.align === "spread" ? " selected" : "") + '>Left and right edges</option></select></label>' +
              '<button type="button" class="btn btn-link btn-sm p-0" id="mos-pf-hdr-reset">Default size</button>' +
              '<div class="mos-pf-brandprev mos-al-' + bw.align + '" id="mos-pf-brandprev" style="--logo-h:' + Math.round(bw.logoH * 0.6) + 'px" aria-hidden="true">' +
                '<img src="' + esc(bw.left || RAKSO_DEFAULT) + '" alt="">' + (bw.right ? '<img src="' + esc(bw.right) + '" alt="">' : "") + '</div>' +
            '</div></div>' +
        '</div>' +
        '<div id="mos-pf-msg" class="small text-danger mt-1" role="alert">' + esc(msg || "") + '</div>' +
        '<div class="d-flex justify-content-end gap-2 mt-2"><button type="button" class="btn btn-outline-secondary btn-sm" id="mos-pf-cancel">Cancel</button>' +
        '<button type="button" class="btn btn-primary btn-sm" id="mos-pf-save">Save</button></div></div>';
      function keep() { work.title = ov.querySelector("#mos-pf-title").value.trim(); work.acronym = ov.querySelector("#mos-pf-acr").value.trim().toUpperCase(); work.address = ov.querySelector("#mos-pf-addr").value.trim(); }
      ov.querySelector("#mos-pf-logo").addEventListener("change", function (e) {
        keep();
        readImage(e.target.files[0], LOGO_MAX).then(function (url) { work.logo = url; draw(); }, function (err) { draw(err.message); });
      });
      ["left", "right"].forEach(function (k) {
        ov.querySelector("#mos-pf-" + k).addEventListener("change", function (e) {
          keep();
          readImage(e.target.files[0], LOGO_MAX).then(function (url) { bw[k] = url; draw(); }, function (err) { draw(err.message); });
        });
      });
      ov.querySelectorAll("[data-pf-coclear]").forEach(function (b) { b.addEventListener("click", function () { keep(); bw[b.getAttribute("data-pf-coclear")] = ""; draw(); }); });
      function hdrLive() {
        bw.logoH = +ov.querySelector("#mos-pf-logoh").value; bw.scale = +ov.querySelector("#mos-pf-scale").value; bw.align = ov.querySelector("#mos-pf-align").value;
        ov.querySelector("#mos-pf-logoh-v").textContent = bw.logoH + " px"; ov.querySelector("#mos-pf-scale-v").textContent = bw.scale + "%";
        var pv = ov.querySelector("#mos-pf-brandprev"); pv.style.setProperty("--logo-h", Math.round(bw.logoH * 0.6) + "px"); pv.className = "mos-pf-brandprev mos-al-" + bw.align;
      }
      ["#mos-pf-logoh", "#mos-pf-scale"].forEach(function (sel) { ov.querySelector(sel).addEventListener("input", hdrLive); });
      ov.querySelector("#mos-pf-align").addEventListener("change", hdrLive);
      ov.querySelector("#mos-pf-hdr-reset").addEventListener("click", function () { keep(); bw.logoH = HDR_DEF.logoH; bw.scale = HDR_DEF.scale; bw.align = HDR_DEF.align; draw(); });
      ov.querySelectorAll("[data-pf-clear]").forEach(function (b) { b.addEventListener("click", function () { keep(); work[b.getAttribute("data-pf-clear")] = ""; draw(); }); });
      ov.querySelector("#mos-pf-cancel").addEventListener("click", close);
      ov.querySelector("#mos-pf-save").addEventListener("click", function () {
        keep();
        var all = profiles(); all[school] = work;
        if (!work.title && !work.acronym && !work.address && !work.logo) delete all[school];
        if (saveProfiles(all) && saveBranding(bw)) { close(); rebuild(); }
      });
      setTimeout(function () { var t = ov.querySelector("#mos-pf-title"); if (t) t.focus(); }, 20);
    }
    function close() { document.removeEventListener("keydown", onKey, true); if (ov.parentNode) ov.parentNode.removeChild(ov); var b = $("mos-profile"); if (b) b.focus(); }
    function onKey(e) { if (e.key === "Escape") { e.preventDefault(); close(); } }
    document.addEventListener("keydown", onKey, true);
    document.body.appendChild(ov); draw();
  }

  /* ---------------- Exam Summary (print → Save as PDF) ----------------
   * Covers exactly the current filters. Layout (2021 report): company logos row (Rakso CT left, partner badge
   * right), 3-column box [school logo + name | MOS in <exam> · EXAM RESULTS | date + Exam No.] and
   * [total / passed / failed / eligible for retake | big passing rate], then the student table: skill areas as columns I–V when every row is the same exam (legend
   * under the table), red below 50%; scores below 700 red. Opened as a print view; the browser's
   * "Save as PDF" makes the file. No emails, addresses or IDs are printed. */
  function romans(i) { return ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"][i] || String(i + 1); }
  /* company logos on every summary: left = Rakso CT (built in, replaceable), right = partner badge (upload) */
  var BRAND_KEY = "qbr-mos-branding-v1", RAKSO_DEFAULT = "assets/rakso-ct-logo.png";
  // summary header size (School profile › Summary header size): logo height px, results box scale %, logo position
  var HDR_DEF = { logoH: 40, scale: 85, align: "center" };
  function hdrOpt(br) {
    br = br || branding();
    var h = +br.logoH, sc = +br.scale;
    return { logoH: h >= 24 && h <= 80 ? h : HDR_DEF.logoH, scale: sc >= 60 && sc <= 120 ? sc : HDR_DEF.scale, align: br.align === "spread" ? "spread" : "center" };
  }
  function branding() { try { return JSON.parse(localStorage.getItem(BRAND_KEY) || "{}") || {}; } catch (e) { return {}; } }
  function saveBranding(b) {
    try { localStorage.setItem(BRAND_KEY, JSON.stringify(b)); return true; }
    catch (e) { window.alert("Could not save the company logos in this browser (storage is full). Try a smaller image."); return false; }
  }
  function upper(s) { return String(s || "").toUpperCase(); }
  var MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  // "October 8 – 15, 2026" · "September 28 – October 2, 2026" · "December 30, 2025 – January 3, 2026"
  function dateRange(a, b) {
    if (!a) return "";
    if (!b || a === b) return longDate(a);
    var ma = MONTHS_LONG[+a.slice(5, 7) - 1], mb = MONTHS_LONG[+b.slice(5, 7) - 1], da = +a.slice(8, 10), db = +b.slice(8, 10);
    if (a.slice(0, 4) !== b.slice(0, 4)) return longDate(a) + " – " + longDate(b);
    if (ma === mb) return ma + " " + da + " – " + db + ", " + a.slice(0, 4);
    return ma + " " + da + " – " + mb + " " + db + ", " + a.slice(0, 4);
  }
  function longDate(iso) { return iso ? MONTHS_LONG[+iso.slice(5, 7) - 1] + " " + (+iso.slice(8, 10)) + ", " + iso.slice(0, 4) : ""; }
  // filters that hide part of a batch (pass/fail, score range, search, first-attempt) → shown on the summary + warned about
  function narrowingFilters() {
    var f = S.f, p = [];
    if (f.outcome) p.push("Result: " + (f.outcome === "PASS" ? "Pass" : f.outcome === "FAIL" ? "Fail" : "In progress"));
    if (f.min !== "" || f.max !== "") p.push("Score " + (f.min || 0) + "–" + (f.max || 1000));
    if (f.q) p.push("Search: " + (/@/.test(f.q) ? "(email)" : f.q));
    if (f.firstOnly) p.push("First attempts only");
    return p;
  }
  function buildSummary(R) {
    var rows = R.slice().sort(function (a, b) {
      return (a.section || "~").localeCompare(b.section || "~") || String(a.rosterName || a.name).toLowerCase().localeCompare(String(b.rosterName || b.name).toLowerCase()) || (a.date || "").localeCompare(b.date || "");
    });
    var schools = uniq(rows.map(function (r) { return r.school; }));
    var school = schools.length === 1 ? schools[0] : "";
    var pr = school ? profileOf(school) : {}, br = branding();
    var exams = uniq(rows.map(function (r) { return r.exam; }));
    var certs = uniq(rows.map(function (r) { return r.certName; }));
    var dates = uniq(rows.map(function (r) { return r.date; }));
    var batches = uniq(rows.map(function (r) { return r.examNo || r.batch; }));
    var s = P().summarize(rows);
    // eligible = failed first takes that have NOT been retaken yet (a student who already retook is not counted)
    var eligible = rows.filter(function (r) { return r.retakeEligible && !r.retakeUsed; }).length;
    // skill columns I–V: the objectives most rows share (one exam family). Rows from another exam version with the
    // same number of objectives are filled by position and marked †, with their own legend line; others show —.
    var areas = [], altSigs = {}, sigOf = function (r) { return (r.skills || []).map(function (k) { return k.area; }).join("|"); };
    if (exams.length === 1) {
      var cnt = {};
      rows.forEach(function (r) { if (r.skills && r.skills.length) { var g = sigOf(r); cnt[g] = (cnt[g] || 0) + 1; } });
      var top = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; })[0];
      if (top) {
        areas = top.split("|");
        Object.keys(cnt).forEach(function (g) { if (g !== top && g.split("|").length === areas.length) altSigs[g] = 1; });
      }
    }
    var examTitle = exams.length === 1 ? "Microsoft Office Specialist in " + exams[0] : exams.length ? "Microsoft Office Specialist — " + exams.join(", ") : "Microsoft Office Specialist";
    var certLine = certs.length === 1 ? certs[0] : "";
    var dateTxt = dateRange(dates[0], dates[dates.length - 1]);
    var batchTxt = batches.length === 1 ? (/^\d/.test(batches[0]) ? "Exam No.: " + batches[0] : "Batch: " + batches[0]) : batches.length + " batches";
    var batchList = batches.length > 1 && batches.length <= 6 ? batches.join(", ") : "";
    var narrow = narrowingFilters();
    var leftLogo = br.left || RAKSO_DEFAULT;
    var ho = hdrOpt(br);
    var html = '<div class="mos-sum-doc" style="--hs:' + (ho.scale / 100) + ';--logo-h:' + ho.logoH + 'px">' +
      '<div class="mos-sum-brand mos-al-' + ho.align + '"><img class="mos-sum-co" alt="Rakso CT" src="' + esc(leftLogo) + '">' +
        (br.right ? '<img class="mos-sum-co" alt="Partner logo" src="' + esc(br.right) + '">' : "") + "</div>" +
      '<table class="mos-sum-box"><tbody>' +
        "<tr>" +
          '<td class="c1"><div class="mos-sum-school">' + (pr.logo ? '<img class="mos-sum-logo" alt="" src="' + esc(pr.logo) + '">' : "") +
            '<div><div class="nm">' + esc(upper(pr.title || school || "All schools")) + "</div>" + (pr.address ? '<div class="ad">' + esc(pr.address) + "</div>" : "") + "</div></div></td>" +
          '<td class="c2"><div class="sub">' + esc(examTitle) + "</div><div class=\"big\">EXAM RESULTS</div>" + (certLine ? '<div class="ad">' + esc(certLine) + "</div>" : "") + "</td>" +
          '<td class="c3"><div class="dt">' + esc(upper(dateTxt || "—")) + '</div><div class="no">' + esc(batchTxt) + "</div>" + (batchList ? '<div class="ad bl">' + esc(batchList) + "</div>" : "") + "</td>" +
        "</tr><tr>" +
          '<td class="c1"><div class="mos-sum-counts">' +
            '<div><b>' + num(s.attempts) + "</b><span>TOTAL<br>EXAMINEES</span>" + (s.students < s.attempts ? '<i class="st">' + num(s.students) + " students</i>" : "") + "</div>" +
            '<div class="ok"><b>' + num(s.passed) + "</b><span>STUDENTS<br>PASSED</span></div>" +
            '<div class="bad"><b>' + num(s.failed) + "</b><span>STUDENTS<br>FAILED</span></div>" +
            '<div><b>' + num(eligible) + "</b><span>ELIGIBLE<br>FOR RETAKE</span></div>" +
          "</div></td>" +
          '<td class="c23" colspan="2"><span class="rate">' + (s.passRate == null ? "—" : Math.round(s.passRate * 100) + "%") + '</span><span class="rl">STUDENT PASSING RATE</span>' +
            '<div class="ad">Average ' + num(s.avg) + " · Highest " + (s.high == null ? "—" : s.high) + " · Lowest " + (s.low == null ? "—" : s.low) + (s.pending ? " · " + s.pending + " in progress (not counted)" : "") + "</div></td>" +
        "</tr></tbody></table>" +
      (narrow.length ? '<div class="mos-sum-narrow">Filtered: ' + esc(narrow.join(" · ")) + " — this summary does not show the whole batch.</div>" : "") +
      '<table class="mos-sum-tbl"><thead><tr><th class="num">#</th><th>Name</th><th>Section</th>' + (exams.length > 1 ? "<th>Exam</th>" : "") + (dates.length > 1 ? "<th>Date</th>" : "") +
        areas.map(function (a, i) { return '<th class="num" title="' + esc(a) + '">' + romans(i) + "</th>"; }).join("") +
        '<th class="num">Score</th><th>Result</th></tr></thead><tbody>' +
      rows.map(function (r, i) {
        var sg = sigOf(r), alt = areas.length && sg && sg !== areas.join("|"), fit = !alt || altSigs[sg];
        var skCells = areas.map(function (a, j) { var k = fit ? (r.skills || [])[j] : null; return '<td class="num' + (k && k.pct < 50 ? " low" : "") + '">' + (k ? k.pct + "%" : "—") + "</td>"; }).join("");
        var res = r.pending ? "In progress" : r.pass ? "Passed" : r.retakeEligible ? (r.retakeUsed ? "Failed · retook" : "Failed · retake eligible") : "Failed · no retakes left";
        return "<tr" + (r.pending ? ' class="wait"' : !r.pass ? ' class="fail"' : "") + '><td class="num">' + (i + 1) + "</td><td>" + esc(r.rosterName || r.name) + (alt && fit ? " †" : "") + (r.isRetake ? ' <span class="rt">2nd take</span>' : "") +
          "</td><td>" + esc(r.section || "—") + "</td>" + (exams.length > 1 ? "<td>" + esc(r.exam) + "</td>" : "") + (dates.length > 1 ? "<td>" + fmtDate(r.date) + "</td>" : "") + skCells +
          '<td class="num"><b>' + (r.score == null ? "—" : r.score) + "</b></td><td>" + esc(res) + "</td></tr>";
      }).join("") + "</tbody></table>" +
      (areas.length ? '<div class="mos-sum-legend"><b>Skill areas</b> (section analysis %, red below 50%): ' + areas.map(function (a, i) { return "<span><b>" + romans(i) + "</b> " + esc(a) + "</span>"; }).join("") + "</div>" : "") +
      Object.keys(altSigs).map(function (g) { return '<div class="mos-sum-legend"><b>† Other exam version:</b> ' + g.split("|").map(function (a, i) { return "<span><b>" + romans(i) + "</b> " + esc(a) + "</span>"; }).join("") + "</div>"; }).join("") +
      '<div class="mos-sum-foot"><span>Passing score: 700 / 1000 (700 passes). Up to 2 takes per exam: a first-time taker who fails is eligible for one retake (matched by email).</span>' +
        "<span>Source: Certiport score reports" + (S.exports.length ? " and results export" : "") + " · Filters: " + esc(describeFilters()) + "</span><span>Generated " + esc(new Date().toLocaleString()) + " · RCT OpsDesk</span></div>" +
      "</div>";
    return { html: html, rows: rows.length, areas: areas.length, school: school, logo: !!pr.logo, partner: !!br.right, eligible: eligible, narrow: narrow };
  }
  function printSummary() {
    var R = rows();
    if (!R.length) { window.alert("No attempts match the current filters."); return; }
    if (R.length > 600 && !window.confirm(R.length + " attempts will be printed. Narrow the filters to one school / batch for a shorter summary. Continue?")) return;
    var f = S.f;
    if (f.outcome === "PASS" || f.min !== "" || f.max !== "") {
      var wide = P().filterAttempts(S.model.rows, filterObj({ outcome: "", min: "", max: "" }));
      var hidden = wide.length - R.length;
      if (hidden > 0 && !window.confirm("The Result / Score filters leave out " + hidden + " attempt" + (hidden === 1 ? "" : "s") + " of this batch (e.g. failed students). Print the summary anyway?\n\nCancel, then Reset filters, to print the whole batch.")) return;
    }
    var out = buildSummary(R);
    var host = $("mos-print"); if (host) host.parentNode.removeChild(host);
    host = document.createElement("div"); host.id = "mos-print"; host.innerHTML = out.html;
    document.body.appendChild(host);
    document.body.classList.add("mos-printing");
    var prevTitle = document.title;
    document.title = "MOS Exam Summary" + (out.school ? " - " + (profileOf(out.school).acronym || P().acronym(out.school)) : "") + (S.f.examNo ? " - " + S.f.examNo : "") + " - " + new Date().toISOString().slice(0, 10);
    var done = false;
    function cleanup() {
      if (done) return; done = true;
      document.body.classList.remove("mos-printing"); document.title = prevTitle;
      var h = $("mos-print"); if (h && h.parentNode) h.parentNode.removeChild(h);
      window.removeEventListener("afterprint", cleanup);
    }
    window.addEventListener("afterprint", cleanup);
    var imgs = Array.prototype.slice.call(host.querySelectorAll("img"));
    Promise.race([Promise.all(imgs.map(function (im) { return im.decode ? im.decode().catch(function () {}) : Promise.resolve(); })), new Promise(function (r) { setTimeout(r, 3000); })])
      .then(function () { try { window.print(); } finally { setTimeout(cleanup, 60000); } });
    S.lastSummary = out;
  }

  /* ---------------- hooks ---------------- */
  QBR.mos = {
    render: render,
    addPdfs: addPdfs, addRoster: addRoster, clear: clearAll,
    // test seams (no DOM file pickers in headless tests)
    _state: function () { return S; },
    _summary: function () { return buildSummary(rows()); }, printSummary: printSummary, openProfile: openProfile,
    _loadAttempts: function (attempts, roster) { S.attempts = attempts.slice(); S.roster = (roster || []).slice(); rebuild(); }
  };
  document.addEventListener("qbr:flags-ready", function () { if (S.built || $("mos-body")) render(); });
  function boot() { if ($("mos-body")) render(); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot();
})();
