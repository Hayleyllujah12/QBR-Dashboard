/* ==========================================================================
 * mos-certs.js — v1.41.0 MOS › Certificates: which certificates were generated
 * and which were printed.
 *
 * - Certificate pages come from mos.js (readPdf → QBR.mosParse.parseCertificatePage);
 *   one entry per verification code (the certificate's own ID).
 * - Each certificate is linked to a passing attempt (name + certification + date,
 *   QBR.mosParse.matchCertificates); unclear ones are picked by hand.
 * - The log is an Excel workbook in the linked folder — "MOS Certificate Log.xlsx",
 *   one row per verification code — written with the format-safe save engine
 *   (QBR.xlsxSurgical). It is re-read right before every write, so edits made in
 *   Excel / OneDrive in the meantime are kept.
 * - Privacy (MOS rule): no student data in browser storage. Only the staff
 *   "Printed by" name is remembered (localStorage qbr-mos-printedby).
 * ========================================================================== */
(function () {
  "use strict";
  var QBR = (window.QBR = window.QBR || {});
  var P = function () { return QBR.mosParse; };
  var LOG_NAME = "MOS Certificate Log.xlsx";
  var BY_KEY = "qbr-mos-printedby";
  var C = {
    log: null,          // { name, where, handle, dirId, writable, bytes, parsed, dirty }
    links: {}, sel: {}, folderId: "", busy: false, prog: null, msg: "", limit: 300,
    scope: "",          // "" = schools in the loaded results (all when none), "all" = every school folder
    f: { school: "", examNo: "", cert: "", status: "", q: "" }
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function str(v) { return v == null ? "" : String(v).trim(); }
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function fmtDate(iso) { return iso ? MON[+iso.slice(5, 7) - 1] + " " + (+iso.slice(8, 10)) + ", " + iso.slice(0, 4) : "—"; }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function isoOf(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date) return isNaN(v) ? "" : v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
    if (typeof v === "number" && v > 20000 && v < 80000) { var d = new Date(Math.round((v - 25569) * 864e5)); return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate()); }
    var s = String(v).trim(), m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : (P().parseDate(s) || "");
  }
  function dateOf(iso) { return iso ? new Date(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) : ""; }
  function M() { return QBR.mos && QBR.mos._state ? QBR.mos._state() : null; }
  function msg(t) { C.msg = t || ""; var el = $("mos-cert-msg"); if (el) el.textContent = C.msg; }
  // Progress for the folder load: determinate bar where a real fraction
  // exists (file reads), sliding bar where it doesn't (the folder scan).
  // Without it a large folder looked hung and invited spam-clicking.
  function progHtml() {
    var p = C.prog; if (!p) return "";
    if (p.frac == null) return '<div style="height:6px;border-radius:3px;background:rgba(127,127,127,.25);overflow:hidden"><i style="display:block;height:100%;width:40%;border-radius:3px;background:var(--amber,#e8a33d);animation:readslide 1.1s ease-in-out infinite"></i></div>' +
      '<div class="small text-muted mt-1">' + esc(p.label || "Loading\u2026") + "</div>";
    var pct = Math.max(0, Math.min(100, Math.round(p.frac * 100)));
    return '<div class="progress" style="height:16px" role="progressbar" aria-valuenow="' + pct + '" aria-valuemin="0" aria-valuemax="100"><div class="progress-bar progress-bar-striped progress-bar-animated" style="width:' + pct + '%"></div></div>' +
      '<div class="small text-muted mt-1">' + esc(p.label || "") + " \u00b7 " + pct + "%</div>";
  }
  function setProg(label, frac) {
    C.prog = label == null ? null : { label: label, frac: frac };
    var el = $("mos-cert-prog"); if (el) el.innerHTML = progHtml();
  }

  /* ---------------- log workbook ---------------- */
  function findLogSheet(wb) {
    if (wb.Sheets[P().LOG_SHEET]) return P().LOG_SHEET;
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, defval: "", blankrows: true });
      if (aoa.slice(0, 10).some(function (r) { return (r || []).some(function (h) { return str(h).toLowerCase() === "verification code"; }); })) return wb.SheetNames[i];
    }
    return "";
  }
  function parseLog(bytes) {
    var wb = XLSX.read(bytes, { type: "array", cellDates: true });
    var sn = findLogSheet(wb);
    var parsed = sn ? P().parseLogRows(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "", blankrows: true, raw: true }))
      : P().parseLogRows([]);
    parsed.sheet = sn;
    return parsed;
  }
  function newLogBytes() {
    var ws = XLSX.utils.aoa_to_sheet([P().LOG_COLS]);
    ws["!cols"] = P().LOG_COLS.map(function (n) { return { wch: /Name|Source|Remarks|Certification/.test(n) ? 28 : 16 }; });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, P().LOG_SHEET);
    return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
  }
  function useLog(o) {
    o.parsed = parseLog(o.bytes);
    C.log = o;
  }
  /* OneDrive / sync tools touch a file right after it is written: a File snapshot taken a moment earlier then
   * fails with NotReadableError ("…permission problems that have occurred after a reference to a file was
   * acquired"). Take a fresh snapshot and retry a few times before giving up. */
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  async function readHandle(h, tries) {
    tries = tries || 6;
    for (var i = 0; ; i++) {
      try { var f = await h.getFile(); return { file: f, bytes: new Uint8Array(await f.arrayBuffer()) }; }
      catch (e) {
        var busy = e && (e.name === "NotReadableError" || e.name === "InvalidStateError" || e.name === "NoModificationAllowedError");
        if (!busy || i >= tries - 1) throw e;
        await sleep(250 * (i + 1));
      }
    }
  }
  C._readHandle = readHandle;
  async function openLogHandle(h, dirId, where, knownBytes) {
    var rd = knownBytes ? { file: null, bytes: knownBytes } : await readHandle(h);
    var perm = "prompt";
    try { perm = await h.queryPermission({ mode: "readwrite" }); if (perm !== "granted") perm = await h.requestPermission({ mode: "readwrite" }); } catch (e) {}
    useLog({ name: (rd.file && rd.file.name) || h.name || LOG_NAME, where: where || h.name, handle: h, dirId: dirId || "", writable: perm === "granted", bytes: rd.bytes });
  }
  async function pickLog() {
    if (typeof window.showOpenFilePicker === "function") {
      var hs;
      try { hs = await window.showOpenFilePicker({ multiple: false, types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] }); }
      catch (e) { return; }
      await openLogHandle(hs[0], "", hs[0].name);
      render(); return;
    }
    var inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx"; inp.style.display = "none"; document.body.appendChild(inp);
    inp.onchange = async function () {
      var f = inp.files && inp.files[0]; inp.remove(); if (!f) return;
      useLog({ name: f.name, where: f.name, handle: null, writable: false, bytes: new Uint8Array(await f.arrayBuffer()) });
      render();
    };
    inp.click();
  }
  function folders() { try { return QBR.fsFolderList ? QBR.fsFolderList() : []; } catch (e) { return []; } }
  // Which linked folder the certificates live in: the one you chose (remembered — a folder id, no student data),
  // else a folder whose name looks like the MOS / certificates folder, else the first one.
  var FOLDER_KEY = "qbr-mos-cert-folder";
  function remember(id) { try { localStorage.setItem(FOLDER_KEY, id); } catch (e) {} }
  function currentFolder() {
    var l = folders(), saved = ""; try { saved = localStorage.getItem(FOLDER_KEY) || ""; } catch (e) {}
    var by = function (id) { return id ? l.filter(function (d) { return d.id === id; })[0] : null; };
    return by(C.folderId) || by(saved) || l.filter(function (d) { return /mos|certif|microsoft office specialist|course/i.test(d.name); })[0] || l[0] || null;
  }
  function setFolder(id) {
    if (C.log && C.log.dirId && C.log.dirId !== id && !C.log.dirty) C.log = null;   // that folder's own log is found on the next load
    C.folderId = id; remember(id);
  }
  async function chooseFolder() {
    if (!QBR.fsFolderPickOnly) return;
    var id = await QBR.fsFolderPickOnly();
    if (!id) return;
    setFolder(id);
    await loadFromFolder();
  }
  async function createLog() {
    var d = currentFolder(), dir = d && QBR.fsFolderDir ? await QBR.fsFolderDir(d.id, { ask: true }) : null;
    if (dir) {
      var h = null;
      try { h = await dir.getFileHandle(LOG_NAME); } catch (e) { h = null; }      // already there → use it
      var made = null;
      if (!h) {
        h = await dir.getFileHandle(LOG_NAME, { create: true });
        made = newLogBytes();
        var w = await h.createWritable(); await w.write(made); await w.close();
      }
      C.folderId = d.id; remember(d.id);
      // just written → use those bytes (reading it back can hit OneDrive touching the new file)
      await openLogHandle(h, d.id, d.name + "/" + LOG_NAME, made);
      return;
    }
    // no folder access (other browsers): the log lives in this tab; "Download log" saves it
    useLog({ name: LOG_NAME, where: "this tab", handle: null, writable: false, bytes: newLogBytes(), dirty: true });
  }
  function printedBy() {
    try { var u = QBR.users && QBR.users.current && QBR.users.current(); if (u && u.name) return u.name; } catch (e) {}
    var v = ""; try { v = localStorage.getItem(BY_KEY) || ""; } catch (e) {}
    if (!v) {
      v = (window.prompt("Your name for “Printed by” (remembered in this browser):", "") || "").trim();
      if (v) try { localStorage.setItem(BY_KEY, v); } catch (e) {}
    }
    return v;
  }
  /* ops: {t:"add"|"print"|"link", code, row?, at?, by?, reason?, ref?} → new bytes (format-safe) */
  function applyOps(bytes, ops) {
    var cs = { type: "array", cellStyles: true, cellDates: true };
    var base = XLSX.read(bytes, cs), work = XLSX.read(bytes, cs);
    var sn = findLogSheet(work);
    if (!sn) { XLSX.utils.book_append_sheet(work, XLSX.utils.aoa_to_sheet([P().LOG_COLS]), P().LOG_SHEET); sn = P().LOG_SHEET; }
    var ws = work.Sheets[sn];
    var o = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).s : { r: 0, c: 0 };
    var rng = XLSX.utils.decode_range(ws["!ref"] || "A1");
    var put = function (r, c, v) {
      if (c == null || c < 0) return;
      var R = o.r + r, CC = o.c + c, ref = XLSX.utils.encode_cell({ r: R, c: CC });
      if (v instanceof Date) ws[ref] = { t: "d", v: v };
      else if (typeof v === "number") ws[ref] = { t: "n", v: v };
      else ws[ref] = { t: "s", v: String(v) };
      if (R > rng.e.r) rng.e.r = R; if (CC > rng.e.c) rng.e.c = CC;
    };
    var parsed = P().parseLogRows(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: true, raw: true }));
    var col = parsed.col, hr = parsed.headerRow, next = parsed.header.length;
    P().LOG_COLS.forEach(function (n) { if (col[n] == null) { col[n] = next++; put(hr, col[n], n); } });
    var last = parsed.last;
    ops.forEach(function (op) {
      var rec = parsed.rows[op.code];
      if (!rec) {
        if (!op.row) return;
        rec = parsed.rows[op.code] = { r: ++last, v: {} };
        P().LOG_COLS.forEach(function (n) { var v = op.row[n]; if (v !== "" && v != null) { put(rec.r, col[n], v); rec.v[n] = v; } });
      }
      if (op.t === "print") {
        var cnt = (+rec.v["Print Count"] || 0) + 1;
        put(rec.r, col.Printed, "Yes");
        if (!rec.v["First Printed"]) { put(rec.r, col["First Printed"], op.at); rec.v["First Printed"] = op.at; }
        put(rec.r, col["Last Printed"], op.at);
        if (op.by) put(rec.r, col["Printed By"], op.by);
        put(rec.r, col["Print Count"], cnt);
        if (op.reason) { var rm = str(rec.v.Remarks); rm = (rm ? rm + "; " : "") + op.reason; put(rec.r, col.Remarks, rm); rec.v.Remarks = rm; }
        rec.v["Print Count"] = cnt; rec.v.Printed = "Yes";
      }
      if (op.t === "link") { put(rec.r, col["Linked Exam Ref #"], op.ref || "none"); rec.v["Linked Exam Ref #"] = op.ref || "none"; }
    });
    ws["!ref"] = XLSX.utils.encode_range(rng);
    var SG = QBR.xlsxSurgical;
    var plain = function () { return new Uint8Array(XLSX.write(work, { bookType: "xlsx", type: "array", cellDates: true })); };
    if (!SG) return Promise.resolve(plain());
    return Promise.resolve().then(function () { return SG.surgicalSave(bytes, base, work); }).then(function (res) {
      if (!res || !res.ok) return plain();
      if (res.noChanges) return bytes;
      return Promise.resolve(SG.verifySurgical(bytes, res.bytes, work)).then(function (why) { return why ? plain() : res.bytes; });
    }, function () { return plain(); });
  }
  async function commit(ops, label) {
    if (!ops.length || C.busy) return false;
    if (!C.log) {
      var d = currentFolder();
      if (!window.confirm("No certificate log yet.\n\nCreate “" + LOG_NAME + "”" + (d ? " in " + d.name : " (download it afterwards)") + " to record this?")) return false;
      try { await createLog(); } catch (e) { window.alert("Could not create the log: " + (e && e.message || e)); return false; }
      if (!C.log) return false;
    }
    C.busy = true; msg("Saving to " + C.log.name + "…");
    try {
      var bytes = C.log.bytes;
      if (C.log.handle) {                         // latest copy (Excel / OneDrive edits kept)
        try { bytes = (await readHandle(C.log.handle)).bytes; }
        catch (e) {
          if (!(e && e.name === "NotFoundError")) {
            window.alert("Couldn't read the latest " + C.log.name + " (OneDrive may still be syncing it). Nothing was changed — try again in a moment.");
            return false;
          }
        }
      }
      var out = await applyOps(bytes, ops);
      var saved = false;
      if (C.log.handle && C.log.writable) {
        var w = null;
        try { w = await C.log.handle.createWritable(); await w.write(out); await w.close(); saved = true; }
        catch (e) {
          try { if (w && w.abort) await w.abort(); } catch (x) {}
          window.alert("Couldn't write to " + C.log.name + " — is it open in Excel? Close it and try again. Nothing was changed.");
          return false;
        }
      } else C.log.dirty = true;
      C.log.bytes = out; C.log.parsed = parseLog(out);
      msg(label + (saved ? " — saved to " + C.log.name + "." : " — not saved to a file yet: use “Download log”."));
      return true;
    } catch (e) {
      window.alert("Could not update the certificate log: " + (e && e.message || e));
      return false;
    } finally { C.busy = false; render(); }
  }

  /* ---------------- items ---------------- */
  function linksNow() {
    var links = {}, lr = C.log ? C.log.parsed.rows : {};
    Object.keys(lr).forEach(function (code) { var ref = str(lr[code].v["Linked Exam Ref #"]); if (ref) links[code] = ref; });
    Object.keys(C.links).forEach(function (k) { links[k] = C.links[k]; });
    return links;
  }
  function items() {
    var s = M(); if (!s) return [];
    var rows = s.model ? s.model.rows : [], lr = C.log ? C.log.parsed.rows : {};
    var mm = P().matchCertificates(s.certs, rows, linksNow());
    var byKey = {}, byRef = {}; rows.forEach(function (r) { byKey[r.key] = r; if (r.ref) byRef[r.ref] = r; });
    var out = [], seen = {}, covered = {};
    Object.keys(mm.byAttempt).forEach(function (k) { covered[k] = 1; });
    mm.certs.forEach(function (c) {
      seen[c.code] = 1;
      var at = c.attemptKey ? byKey[c.attemptKey] : null, l = lr[c.code];
      out.push({ kind: "cert", code: c.code, name: c.name, cert: c.certName, family: c.family, version: c.version, date: c.date, pf: P().printedFromPath(c.file),
        school: at ? at.school : c.school, examNo: c.examBase || (at ? at.examBase : ""), file: c.file, fileName: c.fileName, page: c.page,
        match: rows.length ? c.status : "noresults", how: c.how, attempt: at, cands: c.cands.map(function (x) { return byKey[x.key]; }).filter(Boolean),
        dupOf: c.dupOf || "", inLog: !!l, log: l ? l.v : null });
    });
    Object.keys(lr).forEach(function (code) {
      if (seen[code]) return;
      var v = lr[code].v, ref = str(v["Linked Exam Ref #"]), at = ref && ref !== "none" ? byRef[ref] : null;
      if (at) covered[at.key] = 1;
      out.push({ kind: "log", code: code, name: str(v.Name), cert: str(v.Certification), family: str(v.Certification) ? P().examFamily(str(v.Certification)) : "",
        version: str(v.Version), date: isoOf(v["Certificate Date"]), school: str(v.School), examNo: str(v["Exam No."]), file: str(v["Source File"]),
        fileName: str(v["Source File"]).split("/").pop(), page: v.Page, match: "logonly", attempt: at, cands: [], inLog: true, log: v });
    });
    rows.forEach(function (r) {
      if (!r.pass || r.pending || covered[r.key]) return;
      out.push({ kind: "nocert", code: "", name: r.rosterName || r.name, cert: r.exam, family: r.exam, version: r.examVersion || "", date: r.date,
        school: r.school, examNo: r.examBase || "", attempt: r, match: "nocert", cands: [] });
    });
    out.forEach(function (it) {
      it.logPrinted = !!(it.log && P().logPrinted(it.log));
      it.printed = it.logPrinted || !!it.pf;                 // a PDF in a "PRINTED (date)" folder counts as printed
      it.status = it.kind === "nocert" ? "nocert" : it.printed ? "printed"
        : (it.match === "ambiguous" || it.match === "unmatched" || it.match === "duplicate") ? "check" : "pending";
    });
    return out;
  }
  function inScope(it, skipStatus) {
    var f = C.f;
    if (f.school && it.school !== f.school) return false;
    if (f.examNo && it.examNo !== f.examNo) return false;
    if (f.cert && P().examFamily(it.cert || it.family) !== f.cert && it.family !== f.cert) return false;
    if (!skipStatus && f.status && it.status !== f.status) return false;
    if (f.q) { var q = f.q.toLowerCase(); if ((it.name + " " + it.code).toLowerCase().indexOf(q) < 0) return false; }
    return true;
  }
  var ORDER = { check: 0, pending: 1, nocert: 2, printed: 3 };
  function byOrder(a, b) { return ORDER[a.status] - ORDER[b.status] || String(a.school).localeCompare(String(b.school)) || String(a.name).localeCompare(String(b.name)); }
  // Rows the select-all checkbox acts on: exactly the rows on screen (same filter, sort and limit as render).
  function selectableShown() { return items().filter(function (i) { return inScope(i, false); }).sort(byOrder).slice(0, C.limit).filter(function (i) { return i.kind !== "nocert"; }); }

  /* ---------------- render ---------------- */
  var BADGE = {
    printed: ["risk-none", "Printed"], pending: ["risk-medium", "Pending print"],
    nocert: ["risk-high", "Not generated"], check: ["risk-high", "Needs checking"]
  };
  function kpi(label, val, tone, title) {
    return '<div class="col"><div class="kpi kpi-' + tone + '"' + (title ? ' title="' + esc(title) + '"' : "") + '><div class="kpi-val">' + val + '</div><div class="kpi-lbl">' + esc(label) + "</div></div></div>";
  }
  function opts(values, cur, all, lab) {
    return '<option value="">' + esc(all) + "</option>" + values.map(function (v) { return '<option value="' + esc(v) + '"' + (v === cur ? " selected" : "") + ">" + esc(lab ? lab(v) : v) + "</option>"; }).join("");
  }
  function uniq(a) { var o = {}; a.forEach(function (v) { if (v) o[v] = 1; }); return Object.keys(o).sort(); }
  function attemptLabel(r) { return (r.rosterName || r.name) + " · " + r.exam + " · " + fmtDate(r.date) + " · " + (r.score == null ? "—" : r.score) + (r.examBase ? " · Exam " + r.examBase : ""); }
  function logBar(all) {
    var notIn = all.filter(function (i) { return i.kind === "cert" && !i.inLog; }).length;
    var fs = typeof QBR.fsFolderSupported === "function" && QBR.fsFolderSupported(), fl = folders(), cf = currentFolder();
    var h = '<div class="mos-load-row">';
    if (C.log) {
      var n = C.log.parsed.order.length, pr = C.log.parsed.order.filter(function (c) { return P().logPrinted(C.log.parsed.rows[c].v); }).length;
      h += '<span>📒 <b>' + esc(C.log.name) + '</b> <span class="text-muted">· ' + plural(n, "certificate") + " · " + pr + " printed · " + esc(C.log.where) + "</span></span> " +
        (C.log.writable ? '<span class="badge bg-success" title="Every change is written into this workbook (format-safe; re-read first, so edits made in Excel are kept)">saves in place</span>'
          : '<span class="badge bg-warning text-dark">not linked to a file — changes stay in this tab</span> <button type="button" class="btn btn-outline-primary btn-sm" data-mc="download">Download log</button>');
    } else h += '<span class="text-muted">No certificate log yet — generated / printed status is kept in <b>' + LOG_NAME + "</b>.</span> " +
      '<button type="button" class="btn btn-outline-primary btn-sm" data-mc="open-log">Open log…</button> <button type="button" class="btn btn-outline-secondary btn-sm" data-mc="create-log">Create log…</button>';
    if (fs) h += '<div class="mos-cert-folder"><span class="small text-muted">Certificates folder</span> ' +
      (fl.length > 1 ? '<select class="form-select form-select-sm d-inline-block w-auto" data-mc="folder-pick" aria-label="Certificates folder">' + fl.map(function (d) { return '<option value="' + esc(d.id) + '"' + (cf && cf.id === d.id ? " selected" : "") + ">📁 " + esc(d.name) + "</option>"; }).join("") + "</select>"
        : cf ? "<b>📁 " + esc(cf.name) + "</b>" : '<span class="small text-muted">none linked yet</span>') +
      ' <button type="button" class="btn btn-outline-secondary btn-sm" data-mc="folder-choose"' + (C.busy ? " disabled" : "") + ' title="Pick the MOS folder (e.g. the course folder with the school folders). It is remembered and kept with your other linked folders.">' + (fl.length ? "Link another folder…" : "Choose folder…") + "</button>" +
      ' <select class="form-select form-select-sm d-inline-block w-auto" data-mc="scope" aria-label="Schools to read" title="Which school folders to read certificate PDFs from">' +
        '<option value=""' + (C.scope !== "all" ? " selected" : "") + '>Schools in my results</option><option value="all"' + (C.scope === "all" ? " selected" : "") + ">All schools</option></select>" +
      ' <button type="button" class="btn btn-primary btn-sm" data-mc="folder"' + (cf && !C.busy ? "" : " disabled") + ' title="Find certificate PDFs (folders and files with “Certificate” in the name, up to 6 levels deep) and the log in ' +
      esc(cf ? cf.name : "the folder") + '">' + (C.busy ? "Loading\u2026" : "Load certificates") + '</button></div>';
    h += "</div>" + '<div id="mos-cert-prog" class="mt-2">' + progHtml() + "</div>";
    var folderOnly = all.filter(function (i) { return i.kind === "cert" && i.inLog && i.pf && !i.logPrinted; }).length;
    if (folderOnly) h += '<div class="mos-cert-pend"><span><b>' + plural(folderOnly, "certificate") + "</b> " + (folderOnly === 1 ? "is" : "are") + " in a PRINTED folder but not marked printed in the log.</span>" +
      '<button type="button" class="btn btn-primary btn-sm" data-mc="record-folder">Record in log</button></div>';
    if (notIn) h += '<div class="mos-cert-pend"><span><b>' + plural(notIn, "certificate") + "</b> loaded here " + (notIn === 1 ? "is" : "are") + " not in the log yet (newly generated).</span>" +
      '<button type="button" class="btn btn-primary btn-sm" data-mc="add-log">Add to log</button></div>';
    return h + '<div id="mos-cert-msg" class="small text-muted mt-1" role="status" aria-live="polite">' + esc(C.msg) + "</div>";
  }
  function render(host) {
    host = host || $("mos-view-certs");
    if (!host || !P()) return;
    var all = items(), scope = all.filter(function (i) { return inScope(i, true); }), shown = scope.filter(function (i) { return inScope(i, false); });
    var n = function (st) { return scope.filter(function (i) { return i.status === st; }).length; };
    var gen = scope.filter(function (i) { return i.kind !== "nocert"; }).length;
    var passed = scope.filter(function (i) { return i.attempt; }).length;
    shown.sort(byOrder);
    var f = C.f;
    var schools = uniq(all.map(function (i) { return i.school; })), nos = uniq(all.map(function (i) { return i.examNo; })), fams = uniq(all.map(function (i) { return P().examFamily(i.cert || i.family); }));
    var files = uniq(all.filter(function (i) { return i.kind === "cert" && i.status === "pending"; }).map(function (i) { return i.file; }));
    var selN = shown.filter(function (i) { return C.sel[i.code] && i.kind !== "nocert"; }).length;
    var h = '<div class="card-box mos-certlog">' + logBar(all) + "</div>" +
      '<div class="card-box mos-filters mt-3" role="search" aria-label="Certificate filters"><div class="mos-frow">' +
      '<label class="mos-fl"><span>School</span><select class="form-select form-select-sm" data-mcf="school">' + opts(schools, f.school, "All schools") + "</select></label>" +
      '<label class="mos-fl"><span>Exam No.</span><select class="form-select form-select-sm" data-mcf="examNo">' + opts(nos, f.examNo, "All", function (v) { return "Exam " + v; }) + "</select></label>" +
      '<label class="mos-fl"><span>Certification</span><select class="form-select form-select-sm" data-mcf="cert">' + opts(fams, f.cert, "All") + "</select></label>" +
      '<label class="mos-fl"><span>Status</span><select class="form-select form-select-sm" data-mcf="status">' +
        '<option value="">All</option>' + ["pending", "printed", "nocert", "check"].map(function (k) { return '<option value="' + k + '"' + (f.status === k ? " selected" : "") + ">" + BADGE[k][1] + "</option>"; }).join("") + "</select></label>" +
      '<label class="mos-fl mos-fl-q"><span>Search</span><input type="search" class="form-control form-control-sm" data-mcf="q" value="' + esc(f.q) + '" placeholder="Name or verification code" autocomplete="off"></label>' +
      '<button type="button" class="btn btn-outline-primary btn-sm" data-mc="export">Export (.xlsx)</button>' +
      "</div></div>" +
      '<div class="row row-cols-2 row-cols-md-3 row-cols-xl-6 g-3 kpi-row mt-1">' +
        kpi("Passed", passed, "blue", "Passing attempts in the loaded score reports") +
        kpi("Generated", gen, "blue", "Certificates found (PDFs loaded here + rows in the log)") +
        kpi("Printed", n("printed"), "green") +
        kpi("Pending print", n("pending"), n("pending") ? "orange" : "green") +
        kpi("Passed, no certificate", n("nocert"), n("nocert") ? "red" : "green", "Passing attempts without a certificate: not generated yet") +
        kpi("Needs checking", n("check"), n("check") ? "red" : "green", "Certificates that could not be linked to exactly one passing attempt") +
      "</div>" +
      '<div class="card-box mt-3"><div class="mos-ch-head"><h6 class="mb-0">Certificates <span class="small text-muted fw-normal">· ' + plural(shown.length, "row") + "</span></h6>" +
        '<div class="d-flex flex-wrap gap-2 align-items-center">' +
        '<button type="button" class="btn btn-primary btn-sm" data-mc="mark-sel"' + (selN ? "" : " disabled") + ">Mark selected as printed" + (selN ? " (" + selN + ")" : "") + "</button>" +
        '<button type="button" class="btn btn-outline-primary btn-sm" data-mc="print-sel"' + (!selN ? " disabled" : "") + ">Print selected" + (selN ? " (" + selN + ")" : "") + "</button>" +
        (files.length ? '<select class="form-select form-select-sm d-inline-block w-auto" data-mc="file-pick" aria-label="Certificate file">' + files.map(function (x) { return '<option value="' + esc(x) + '">' + esc(x.split("/").pop()) + "</option>"; }).join("") + "</select>" +
          '<button type="button" class="btn btn-outline-primary btn-sm" data-mc="mark-file">Mark file as printed</button>' : "") +
        "</div></div>" +
      '<div class="table-responsive"><table class="table table-sm mos-tbl mos-cert-tbl" data-resizable="static"><thead><tr>' +
        '<th><input type="checkbox" data-mc="sel-all" aria-label="Select all shown certificates"></th><th>Name · school · exam</th><th>Certification · date</th><th>Verification code</th><th>Exam result</th><th>Status</th><th>Printed</th></tr></thead><tbody>' +
      shown.slice(0, C.limit).map(rowHTML).join("") +
      (shown.length ? "" : '<tr><td colspan="7" class="text-muted small">No certificates match.</td></tr>') +
      "</tbody></table></div>" + (shown.length > C.limit ? '<button type="button" class="btn btn-link btn-sm" data-mc="more">Show more (' + (shown.length - C.limit) + ")</button>" : "") + "</div>";
    host.innerHTML = h;
    wire(host);
    // Sync the header select-all checkbox with the real selection (it is recreated
    // unchecked on every render otherwise): checked = all shown selected,
    // indeterminate = some shown selected.
    var sa = host.querySelector('[data-mc="sel-all"]');
    if (sa) {
      var elig = selectableShown();
      var onN = elig.filter(function (i) { return C.sel[i.code]; }).length;
      sa.checked = elig.length > 0 && onN === elig.length;
      sa.indeterminate = onN > 0 && onN < elig.length;
    }
  }
  function rowHTML(it) {
    var b = BADGE[it.status], v = it.log || {};
    var res;
    if (it.kind === "nocert") res = '<span class="text-nowrap">' + esc((it.attempt.score == null ? "" : it.attempt.score + " · ") + fmtDate(it.attempt.date)) + "</span>" +
      '<div class="tr-hint">' + esc(whyNoCert(it)) + "</div>";
    else if (it.attempt && it.match !== "duplicate") res = '<span class="text-nowrap">' + esc(it.attempt.score + " · " + fmtDate(it.attempt.date)) + '</span><div class="tr-hint">' + (it.how === "manual" ? (C.links[it.code] ? "picked by hand" : "linked in the log") : it.kind === "log" ? "linked in the log" : "matched") + "</div>";
    else if (it.match === "noresults" || it.match === "logonly") res = '<span class="small text-muted">no score report loaded</span>';
    else if (it.match === "unlinked") res = '<span class="small text-muted">no matching result (confirmed)</span>';
    else {
      var s = M(), pool = it.cands.length ? it.cands : (s && s.model ? s.model.rows.filter(function (r) { return r.pass && (!it.school || it.school === "Unassigned school" || r.school === it.school) && P().examFamily(r.exam) === it.family; }).slice(0, 60) : []);
      res = (it.match === "duplicate" ? '<div class="small text-danger">Same pass as certificate ' + esc(it.dupOf) + " — reissue?</div>" : "") +
        '<select class="form-select form-select-sm" data-mc="link" data-code="' + esc(it.code) + '" aria-label="Exam result for ' + esc(it.name) + '"><option value="">— pick the exam result —</option>' +
        pool.map(function (r) { return '<option value="' + esc(r.key) + '">' + esc(attemptLabel(r)) + "</option>"; }).join("") + '<option value="none">No matching result</option></select>';
    }
    if (it.printed && !it.logPrinted && it.pf) return rowStart(it, b, res) + '<td>' + (it.pf.date ? esc(fmtDate(it.pf.date)) : "Printed") +
      '<div class="tr-hint" title="The PDF is in a folder named PRINTED — Add to log records it">from folder' + (it.pf.folder ? ": " + esc(it.pf.folder) : "") + "</div></td></tr>";
    var pr = it.printed ? esc(fmtDate(isoOf(v["Last Printed"] || v["First Printed"]))) + ((+v["Print Count"] || 0) > 1 ? ' <span class="badge bg-secondary" title="Printed ' + (+v["Print Count"]) + ' times">×' + (+v["Print Count"]) + "</span>" : "") +
      (v["Printed By"] ? '<div class="tr-hint">' + esc(v["Printed By"]) + "</div>" : "") + '<button type="button" class="btn btn-link btn-sm p-0" data-mc="reprint" data-code="' + esc(it.code) + '">Reprint</button>'
      : it.kind === "nocert" ? "—" : '<button type="button" class="btn btn-outline-primary btn-sm text-nowrap" data-mc="print" data-code="' + esc(it.code) + '">Mark printed</button>';
    return rowStart(it, b, res) + "<td>" + pr + "</td></tr>";
  }
  function rowStart(it, b, res) {
    return '<tr class="mos-cert-' + it.status + '">' +
      "<td>" + (it.kind === "nocert"
        ? '<input type="checkbox" disabled title="No certificate for this pass yet — nothing to print" aria-label="' + esc(it.name) + ': no certificate yet">'
        : '<input type="checkbox" data-mc="sel" data-code="' + esc(it.code) + '"' + (C.sel[it.code] ? " checked" : "") + ' aria-label="Select ' + esc(it.name) + '">') + "</td>" +
      "<td>" + esc(it.name) + (it.kind === "cert" && C.log && !it.inLog ? ' <span class="badge bg-info text-dark" title="Not in the log yet">new</span>' : "") +
        '<div class="tr-hint">' + esc([it.school, it.examNo ? "Exam " + it.examNo : ""].filter(Boolean).join(" · ")) + "</div></td>" +
      "<td>" + esc(it.cert) + '<div class="tr-hint" title="' + esc(it.version || "") + '">' + esc(fmtDate(it.date)) + (it.version ? " · " + esc(it.version) : "") + "</div></td>" +
      '<td><code>' + esc(it.code || "—") + "</code>" + (it.fileName ? '<div class="tr-hint mos-cert-file" title="' + esc(it.file + (it.page ? " · page " + it.page : "")) + '">' + esc(it.fileName) + (it.page ? " · p." + esc(it.page) : "") + "</div>" : "") +
        (it.file && it.how !== "auto" ? ' <button type="button" class="btn btn-link btn-sm p-0" data-mc="preview" data-code="' + esc(it.code) + '">Preview</button>' : "") + "</td>" +
      "<td>" + res + "</td>" +
      '<td><span class="risk-badge ' + b[0] + '">' + b[1] + "</span></td>";
  }

  /* ---------------- print certificates (300 DPI, one dialog) ---------------- */
  // Each certificate page is rendered at scale 3 (~300 DPI) IN THIS PAGE, then all of them
  // print from ONE hidden same-origin frame, so one or many certificates = one print dialog.
  // ROOT CAUSE of the "Rendering 1 of 1…" → "render timed out (30s)" bug: a pdf.js render with the
  // default intent ("display") schedules every step with requestAnimationFrame, and a background tab
  // never runs rAF. The old print window opened as a new tab in front, so this tab went to the
  // background and the render froze. Fix: render with intent "print" (pdf.js then continues with
  // microtasks, no rAF — also the right intent for printing), and print from a hidden frame in this
  // page instead of a pop-up, so switching tabs mid-batch can't stall it and no pop-up permission is needed.
  // (Printing the raw PDF inside a frame is blocked cross-origin on file://; "Open original
  // PDF" in the preview stays the true-original path.) Progress + errors show in an overlay.
  var RENDER_TIMEOUT = 30000;
  function printOverlay() {
    var o = document.getElementById("mos-print-prog");
    if (!o) {
      o = document.createElement("div"); o.id = "mos-print-prog"; o.className = "mos-print-prog"; o.setAttribute("role", "status"); o.setAttribute("aria-live", "polite");
      o.innerHTML = '<span class="mos-print-txt"></span> <button type="button" class="btn btn-sm btn-outline-secondary" data-mc-print="cancel">Cancel</button>';
      document.body.appendChild(o);
    }
    return o;
  }
  function withTimeout(task, ms) {
    var t;
    return Promise.race([task.promise, new Promise(function (res, rej) { t = setTimeout(function () { try { task.cancel(); } catch (e) {} rej(new Error("render timed out (" + Math.round(ms / 1000) + "s)")); }, ms); })])
      .then(function (v) { clearTimeout(t); return v; }, function (e) { clearTimeout(t); throw e; });
  }
  function canvasUrl(cv) {
    return new Promise(function (res) {
      if (cv.toBlob) cv.toBlob(function (b) { res(b ? URL.createObjectURL(b) : cv.toDataURL("image/png")); }, "image/png");
      else res(cv.toDataURL("image/png"));
    });
  }
  var printing = null;
  async function printCertItems(list) {
    if (printing) { msg("Already preparing a print — wait for it or press Cancel."); return; }
    var ms = M(), fr = ms && ms.fileRefs ? ms.fileRefs : {};
    var seen = {}, jobs = [];
    list.forEach(function (it) { if (!it || !it.file || it.kind === "nocert") return; var k = it.file + "#" + (it.page || 1); if (seen[k]) return; seen[k] = 1; jobs.push(it); });
    if (!jobs.length) { msg("Nothing to print - no PDF file for the selected certificates."); return; }
    var job0 = printing = { cancel: false }, ov = printOverlay(), txt = ov.querySelector(".mos-print-txt");
    ov.querySelector("[data-mc-print]").onclick = function () { job0.cancel = true; txt.textContent = "Cancelling…"; };
    var say = function (t) { txt.textContent = t; };
    var urls = [], docCache = {}, skipped = [], land = 0;
    var cleanup = function () { printing = null; if (ov.parentNode) ov.remove(); Object.keys(docCache).forEach(function (k) { try { docCache[k].destroy(); } catch (e) {} }); };
    say("Preparing " + plural(jobs.length, "certificate") + " for print…");
    try {
      var lib = await QBR.mos._pdfLib();
      for (var j = 0; j < jobs.length && !job0.cancel; j++) {
        var job = jobs[j];
        say("Rendering " + (j + 1) + " of " + jobs.length + " — " + (job.fileName || job.file) + (job.page ? " p." + job.page : "") + "…");
        await new Promise(function (r) { setTimeout(r, 0); });   // let the progress text paint
        try {
          var fo = fr[job.file];
          if (!fo && job.fileName) { var keys = Object.keys(fr); for (var i = 0; i < keys.length; i++) { if (keys[i].split("/").pop() === job.fileName) { fo = fr[keys[i]]; break; } } }
          if (!fo) throw new Error("PDF not in memory - reload the certificates folder.");
          var doc = docCache[job.file];
          if (!doc) { doc = await lib.getDocument({ data: new Uint8Array(await fo.arrayBuffer()), isEvalSupported: false, verbosity: 0 }).promise; docCache[job.file] = doc; }
          var pg = await doc.getPage(job.page || 1), ok = false, lastErr = null;
          for (var sc = 3; sc >= 2 && !ok && !job0.cancel; sc--) {
            var cv = document.createElement("canvas");
            try {
              var vp = pg.getViewport({ scale: sc });
              cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
              var ctx = cv.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);
              await withTimeout(pg.render({ canvasContext: ctx, viewport: vp, intent: "print" }), QBR.mosCerts._renderTimeout || RENDER_TIMEOUT);
              urls.push(await canvasUrl(cv)); ok = true; if (vp.width > vp.height) land++;
            } catch (e) { lastErr = e; }
            cv.width = cv.height = 0;   // free the backing store before the next page
          }
          if (!ok && !job0.cancel) throw lastErr || new Error("render failed");
        } catch (e) { skipped.push((job.fileName || job.file) + (job.page ? " p." + job.page : "") + ": " + (e && e.message || e)); }
      }
      if (job0.cancel) { urls.forEach(function (u) { if (/^blob:/.test(u)) URL.revokeObjectURL(u); }); cleanup(); msg("Print cancelled."); return; }
      if (!urls.length) { cleanup(); msg("Could not render any certificate — " + skipped.join("; ") + ". Workaround: Preview › Open original PDF, then print it from there (Ctrl+P)."); return; }
      say("Opening the print dialog (" + plural(urls.length, "page") + ")…");
      var old = document.getElementById("mos-print-frame"); if (old) old.remove();
      var fr2 = document.createElement("iframe"); fr2.id = "mos-print-frame"; fr2.setAttribute("aria-hidden", "true"); fr2.tabIndex = -1;
      fr2.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;opacity:0;pointer-events:none";
      document.body.appendChild(fr2);
      var fd = fr2.contentWindow.document;
      fd.open();
      fd.write("<!DOCTYPE html><html><head><meta charset='utf-8'><title>Certificates</title><style>@page{size:" + (land * 2 >= urls.length ? "landscape" : "portrait") + ";margin:0}html,body{margin:0;background:#fff}img{display:block;width:100vw;height:100vh;object-fit:contain;break-after:page;page-break-after:always}img:last-of-type{break-after:auto;page-break-after:auto}</style></head><body>" +
        urls.map(function (u) { return '<img alt="" src="' + u + '">'; }).join("") + "</body></html>");
      fd.close();
      await Promise.all(Array.prototype.map.call(fd.images, function (im) { return im.complete ? Promise.resolve() : new Promise(function (r) { im.onload = im.onerror = r; }); }));
      var gone = false, finish = function () { if (gone) return; gone = true; setTimeout(function () { urls.forEach(function (u) { if (/^blob:/.test(u)) URL.revokeObjectURL(u); }); if (fr2.parentNode) fr2.remove(); }, 1000); };
      try { fr2.contentWindow.addEventListener("afterprint", finish); } catch (e) {}
      cleanup();
      msg("Printing " + plural(urls.length, "certificate") + (skipped.length ? " · skipped " + skipped.length + ": " + skipped.join("; ") : "") + ". After the sheets come out, confirm with “Mark selected as printed”.");
      fr2.contentWindow.focus();
      (QBR.mosCerts._printHook || function (w) { w.print(); })(fr2.contentWindow);
      setTimeout(finish, 120000);   // fallback if afterprint never fires
    } catch (err) {
      cleanup();
      msg("Print error: " + (err && err.message || err) + ". Workaround: Preview › Open original PDF, then print it from there (Ctrl+P).");
    }
  }
  /* ---------------- certificate preview ---------------- */
  async function previewCert(code) {
    try {
    var it = byCode(code);
    if (!it) { msg("Preview: certificate not found (" + code + ")."); return; }
    if (!it.file) { msg("Preview: no PDF file for this certificate."); return; }
    var ms = M(), fr = ms && ms.fileRefs ? ms.fileRefs : {};
    var fo = fr[it.file];
    if (!fo && it.fileName) {
      var keys = Object.keys(fr);
      for (var i = 0; i < keys.length; i++) { if (keys[i].split("/").pop() === it.fileName) { fo = fr[keys[i]]; break; } }
    }
    if (!fo) { msg("Preview: PDF not in memory — reload the certificates folder (" + it.fileName + ")."); return; }
    var bytes;
    try { bytes = new Uint8Array(await fo.arrayBuffer()); }
    catch (e) { msg("Preview: could not read PDF (" + (e && e.message || e) + ")."); return; }
    var s = M(), pool = it.cands.length ? it.cands : (s && s.model ? s.model.rows.filter(function (r) { return r.pass && (!it.school || it.school === "Unassigned school" || r.schoolFromCenter || P().schoolMatch(r.school, it.school)) && P().examFamily(r.exam) === it.family; }).slice(0, 60) : []);
    var old = document.querySelector(".mos-prev-back"); if (old) old.remove();   // one preview at a time
    var h = '<div class="mos-prev-back" data-mc="preview-close">' +
      '<div class="mos-prev-modal" role="dialog" aria-label="Certificate preview">' +
      '<div class="mos-prev-head"><div><b>' + esc(it.name) + '</b> · ' + esc(it.cert) + ' · ' + esc(fmtDate(it.date)) + '<div class="tr-hint">' + esc(it.code) + (it.page ? " · page " + it.page : "") + " · " + esc(it.fileName) + "</div></div>" +
      '<div class="d-flex gap-2"><button type="button" class="btn btn-sm btn-outline-primary" data-mc="preview-open" title="Open the original PDF in a new tab - print from there at full quality">Open original PDF</button><button type="button" class="btn btn-sm btn-outline-primary" data-mc="preview-print" title="Print this certificate (rendered at 300 DPI)">Print</button><button type="button" class="btn btn-sm btn-outline-secondary" data-mc="preview-close">Close</button></div></div>' +
      '<div class="mos-prev-body"><div class="mos-prev-wrap"><canvas id="mos-prev-canvas"></canvas></div>' +
      '<div class="mos-prev-cands"><h6>Exam results — pick the match</h6>' +
      pool.map(function (r) { return '<button type="button" class="btn btn-outline-primary btn-sm d-block w-100 text-start mb-1" data-mc="preview-link" data-code="' + esc(it.code) + '" data-key="' + esc(r.key) + '">' + esc(attemptLabel(r)) + "</button>"; }).join("") +
      '<button type="button" class="btn btn-outline-secondary btn-sm d-block w-100 text-start" data-mc="preview-link" data-code="' + esc(it.code) + '" data-key="none">No matching result</button>' +
      "</div></div></div>";
    var wrap = document.createElement("div"); wrap.innerHTML = h;
    document.body.appendChild(wrap.firstChild);
    var modal = document.querySelector(".mos-prev-modal");
    modal.addEventListener("click", function (e) {
      e.stopPropagation();
      var b = e.target.closest && e.target.closest("button[data-mc]");
      if (!b) return;
      var a = b.getAttribute("data-mc");
      if (a === "preview-close") { var m = document.querySelector(".mos-prev-back"); if (m) m.remove(); } else if (a === "preview-open") { try { var blob = new Blob([bytes], { type: "application/pdf" }), u = URL.createObjectURL(blob); if (!window.open(u, "_blank")) msg("Pop-up blocked - allow pop-ups for this page to open the PDF."); setTimeout(function () { URL.revokeObjectURL(u); }, 60000); } catch (e) { msg("Could not open PDF: " + (e && e.message || e)); } } else if (a === "preview-print") { printCertItems([it]); }
      else if (a === "preview-link") {
        var key = b.getAttribute("data-key"), ccode = b.getAttribute("data-code");
        C.links[ccode] = key;
        var it2 = byCode(ccode);
        if (it2 && it2.inLog) { var s2 = M(), r2 = s2 && s2.model ? s2.model.rows.filter(function (x) { return x.key === key; })[0] : null;
          commit([{ t: "link", code: ccode, ref: key === "none" ? "none" : (r2 && r2.ref) || "" }], "Link saved"); }
        else render();
        var m2 = document.querySelector(".mos-prev-back"); if (m2) m2.remove();
      }
    });
    // Backdrop click or Escape closes
    var back = document.querySelector(".mos-prev-back");
    var onKey = function (e) { if (e.key === "Escape") { back.remove(); } };
    back.addEventListener("click", function (e) {
      if (e.target.classList && e.target.classList.contains("mos-prev-back")) e.target.remove();
    });
    document.addEventListener("keydown", onKey);
    new MutationObserver(function (_, ob) { if (!document.body.contains(back)) { document.removeEventListener("keydown", onKey); ob.disconnect(); } }).observe(document.body, { childList: true });
    var closeBtn = back.querySelector('button[data-mc="preview-close"]'); if (closeBtn) try { closeBtn.focus(); } catch (x) {}
    // Render the PDF page
    QBR.mos._pdfLib().then(function (lib) {
      return lib.getDocument({ data: bytes.slice(), isEvalSupported: false, verbosity: 0 }).promise;   // copy: pdf.js detaches the buffer it gets, and "Open original PDF" still needs it
    }).then(function (doc) {
      return doc.getPage(it.page || 1);
    }).then(function (pg) {
      var cv = document.getElementById("mos-prev-canvas");
      var vp = pg.getViewport({ scale: 1.2 });
      cv.width = vp.width; cv.height = vp.height;
      return pg.render({ canvasContext: cv.getContext("2d"), viewport: vp }).promise;
    }).catch(function (e) { msg("Preview failed: " + (e && e.message || e)); });
    } catch (err) { console.log("[preview] ERROR:", err && err.message || err); msg("Preview error: " + (err && err.message || err)); }
  }
  /* ---------------- actions ---------------- */
  // Why a passing result shows "Not generated": no certificate PDF of that school read yet, or none with that name.
  function whyNoCert(it) {
    var s = M(), sch = it.school;
    var mine = (s ? s.certs : []).filter(function (c) { return c.school === "Unassigned school" || P().schoolMatch(c.school, sch); });
    if (!mine.length) return "no certificate PDF read for this school yet";
    var fam = mine.filter(function (c) { return P().examFamily(c.family) === P().examFamily(it.family); });
    if (!fam.length) return "certificates read for this school, none for " + it.family;
    var best = fam.map(function (c) { return { c: c, s: P().nameScore(c.name, it.name) }; }).sort(function (a, b) { return b.s - a.s; })[0];
    return best && best.s ? "closest certificate: " + best.c.name + " (" + fmtDate(best.c.date) + ")" : "no " + it.family + " certificate with this name";
  }
  function logRow(it) {
    var now = new Date();
    return { "Verification Code": it.code, "Name": it.name, "Certification": it.cert, "Version": it.version || "", "Certificate Date": dateOf(it.date),
      "School": it.school || "", "Exam No.": it.examNo || "", "Source File": it.file || "", "Page": it.page || "", "First Seen": now,
      "Linked Exam Ref #": it.attempt && it.attempt.ref ? it.attempt.ref : (C.links[it.code] === "none" ? "none" : ""),
      "Printed": it.pf ? "Yes" : "No", "Print Count": it.pf ? 1 : 0,
      "First Printed": it.pf && it.pf.date ? dateOf(it.pf.date) : "", "Last Printed": it.pf && it.pf.date ? dateOf(it.pf.date) : "",
      "Remarks": it.pf ? "Printed (from folder" + (it.pf.folder ? " “" + it.pf.folder + "”" : " / file name") + ")" : "" };
  }
  function byCode(code) { return items().filter(function (i) { return i.code === code; })[0]; }
  async function markPrinted(list, opt) {
    list = list.filter(function (i) { return i && i.kind !== "nocert"; });
    if (!list.length) return;
    var by = printedBy(), at = new Date(), reason = "";
    if (opt && opt.reprint) {
      var it = list[0], v = it.log || {};
      if (!window.confirm(it.name + " — " + it.cert + " was already printed (" + fmtDate(isoOf(v["Last Printed"])) + ", " + plural(+v["Print Count"] || 1, "time") + ").\n\nRecord a reprint?")) return;
      reason = (window.prompt("Reason for the reprint (saved in Remarks):", "") || "").trim();
      reason = "Reprint " + isoOf(at) + (reason ? ": " + reason : "");
    } else if (list.length > 1 && !(opt && opt.confirmed) && !window.confirm("Mark " + plural(list.length, "certificate") + " as printed" + (by ? " by " + by : "") + "?")) return;
    var ops = list.map(function (i) { return { t: "print", code: i.code, at: at, by: by, reason: reason, row: i.inLog ? null : logRow(i) }; });
    var ok = await commit(ops, (opt && opt.reprint ? "Reprint recorded" : plural(list.length, "certificate") + " marked printed"));
    if (ok) list.forEach(function (i) { delete C.sel[i.code]; });
  }
  function exportXlsx() {
    var rows = items().filter(function (i) { return inScope(i, false); });
    var aoa = [["Status", "Name", "Certification", "Version", "Certificate Date", "School", "Exam No.", "Verification Code", "Exam score", "Exam date", "Printed", "Last printed", "Print count", "Printed by", "Source file", "Page"]];
    rows.forEach(function (i) {
      var v = i.log || {};
      aoa.push([BADGE[i.status][1], i.name, i.cert, i.version || "", i.date || "", i.school || "", i.examNo || "", i.code || "",
        i.attempt && i.attempt.score != null ? i.attempt.score : "", i.attempt ? i.attempt.date : "", i.printed ? "Yes" : "No",
        isoOf(v["Last Printed"]), +v["Print Count"] || "", str(v["Printed By"]), i.file || "", i.page || ""]);
    });
    var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Certificates");
    download(new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" })), "MOS certificates " + isoOf(new Date()) + ".xlsx");
  }
  function download(bytes, name) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  async function loadFromFolder() {
    if (C.busy) return;
    var list = folders();
    if (!list.length) { var id = QBR.fsFolderPickOnly ? await QBR.fsFolderPickOnly() : null; if (!id) return; setFolder(id); }
    var d = currentFolder(); if (!d) return;
    C.busy = true; msg("Scanning " + d.name + " for certificate PDFs…");
    setProg("Scanning " + d.name + " for certificate PDFs…", null); render();
    // Read only the schools in the loaded results (a whole MOS folder can hold 700+ certificate PDFs);
    // "All schools" reads every school folder.
    var s0 = M(), rs = {};
    (s0 && s0.model ? s0.model.rows : []).forEach(function (r) { if (r.school && r.school !== "Unassigned school") rs[r.school] = 1; });
    var resSchools = Object.keys(rs), scoped = C.scope !== "all" && resSchools.length > 0, other = {};
    var already = {}; (s0 ? s0.files : []).forEach(function (f) { already[f.file] = 1; });
    try {
      var res = await QBR.fsFolderFiles(d.id, {
        ask: true, ext: /\.(pdf|xlsx)$/i, maxDepth: 6, maxEntries: 50000, maxFiles: 5000,
        want: function (rel) {
          var n = rel[rel.length - 1];
          if (/\.xlsx$/i.test(n)) return n.toLowerCase() === LOG_NAME.toLowerCase();
          if (![d.name].concat(rel).some(function (p) { return /certif/i.test(p); })) return false;
          if (!scoped) return true;
          var sch = P().fileMeta(d.name + "/" + rel.join("/")).school;
          if (sch === "Unassigned school" || resSchools.some(function (x) { return P().schoolRelated(sch, x); })) return true;
          other[sch] = (other[sch] || 0) + 1; return false;
        },
        prefer: function (n) { return /certif/i.test(n); }
      });
      if (!res) { msg("The folder wasn't allowed — nothing was read."); return; }
      C.folderId = d.id; remember(d.id);
      var logs = res.files.filter(function (x) { return /\.xlsx$/i.test(x.name); }).sort(function (a, b) { return a.rel.length - b.rel.length; });
      if (logs.length && !C.log) await openLogHandle(logs[0].handle, d.id, d.name + "/" + logs[0].rel.join("/"));
      var pdfs = res.files.filter(function (x) { return /\.pdf$/i.test(x.name); }), fl = [], perSchool = {};
      pdfs.forEach(function (x) { var sc = P().fileMeta(d.name + "/" + x.rel.join("/")).school; perSchool[sc] = (perSchool[sc] || 0) + 1; });
      // Cache File objects for ALL PDFs (even already-read ones) so Preview can read them on demand.
      var _ms = M(), _fr = _ms ? _ms.fileRefs : null;
      if (_fr && pdfs.length) setProg("Reading certificate files…", 0);
      if (_fr) { for (var _i = 0; _i < pdfs.length; _i++) { try { var _fo = (await readHandle(pdfs[_i].handle)).file; var _p = d.name + "/" + pdfs[_i].rel.join("/"); try { _fo._path = _p; } catch (e) {} _fr[_p] = _fo; } catch (e) {} setProg("Reading certificate files \u00b7 " + (_i + 1) + " of " + pdfs.length, pdfs.length ? (_i + 1) / pdfs.length : 1); if (_i % 20 === 19) await sleep(0); } }
      pdfs = pdfs.filter(function (x) { return !already[d.name + "/" + x.rel.join("/")]; });
      if (pdfs.length) setProg("Reading new certificate files…", 0);
      for (var i = 0; i < pdfs.length; i++) {
        try { var fo = (await readHandle(pdfs[i].handle)).file; try { fo._path = d.name + "/" + pdfs[i].rel.join("/"); } catch (e) {} fl.push(fo); } catch (e) {}
        setProg("Reading new certificate files \u00b7 " + (i + 1) + " of " + pdfs.length, pdfs.length ? (i + 1) / pdfs.length : 1); if (i % 20 === 19) await sleep(0);
      }
      var before = (M() || { certs: [] }).certs.length;
      C.busy = false;
      if (fl.length) setProg("Matching " + plural(fl.length, "certificate") + " — watch the progress bar above…", null);
      if (fl.length && QBR.mos && QBR.mos.addCertFiles) await QBR.mos.addCertFiles(fl);
      var added = (M() || { certs: [] }).certs.length - before;
      var otherN = Object.keys(other).reduce(function (a, k) { return a + other[k]; }, 0);
      if (!Object.keys(perSchool).length && !logs.length) {
        msg("📁 " + d.name + ": no certificate PDFs " + (otherN ? "for the schools in your results (" + plural(otherN, "PDF") + " of other schools skipped — choose “All schools”)" : "and no log found") +
          " (" + res.scanned.toLocaleString() + " items in " + plural(res.folders, "folder") + ")." + (otherN ? "" : " Is this the MOS folder? Use “Link another folder…” to pick the one with the school folders."));
        return;
      }
      C.read = perSchool;
      msg("📁 " + d.name + ": " + plural(Object.keys(perSchool).reduce(function (a, k) { return a + perSchool[k]; }, 0), "certificate PDF") +
        " (" + Object.keys(perSchool).sort().map(function (k) { return k + " " + perSchool[k]; }).join(", ") + ")" +
        (pdfs.length ? "" : " · already read") + " · " + plural(added, "new certificate") +
        (otherN ? " · " + plural(otherN, "PDF") + " of " + plural(Object.keys(other).length, "other school") + " not read (not in the loaded results — choose “All schools” to read them)" : "") +
        (logs.length ? " · log: " + logs[0].rel.join("/") : " · no log in this folder yet") +
        " · " + res.scanned.toLocaleString() + " items in " + res.folders.toLocaleString() + " folders scanned" +
        (res.truncated ? " · stopped early (very large folder — link a narrower one if something is missing)" : "") +
        (res.depthCapped ? " · " + res.depthCapped + " folder(s) deeper than 6 levels not opened" : "") + ".");
    } catch (e) { msg("Could not read the folder: " + (e && e.message || e)); }
    finally { setProg(null); C.busy = false; render(); }
  }
  function wire(host) {
    // Filter inputs are recreated on every render — always re-wire them.
    host.querySelectorAll("[data-mcf]").forEach(function (el) {
      el.addEventListener(el.tagName === "INPUT" ? "input" : "change", function () {
        C.f[el.getAttribute("data-mcf")] = el.value; C.limit = 300;
        if (el.tagName === "INPUT") { var pos = el.selectionStart; render(); var again = host.querySelector('[data-mcf="q"]'); if (again) { again.focus(); try { again.setSelectionRange(pos, pos); } catch (e) {} } }
        else render();
      });
    });
    // Host-level delegation: attach once (host persists across renders).
    if (host._mcWired) return; host._mcWired = true;
    host.addEventListener("change", function (e) {
      var t = e.target, a = t.getAttribute && t.getAttribute("data-mc");
      if (a === "sel") { C.sel[t.getAttribute("data-code")] = t.checked; render(); }
      else if (a === "sel-all") { var elig2 = selectableShown(); var target = elig2.some(function (i) { return !C.sel[i.code]; }); elig2.forEach(function (i) { C.sel[i.code] = target; }); render(); }
      else if (a === "folder-pick") { setFolder(t.value); render(); }
      else if (a === "scope") { C.scope = t.value; }
      else if (a === "link") {
        var code = t.getAttribute("data-code"), v = t.value; if (!v) return;
        C.links[code] = v;
        var it = byCode(code);
        if (it && it.inLog) { var s = M(), r = s && s.model ? s.model.rows.filter(function (x) { return x.key === v; })[0] : null; commit([{ t: "link", code: code, ref: v === "none" ? "none" : (r && r.ref) || "" }], "Link saved"); }
        else render();
      }
    });
    host.addEventListener("click", function (e) {
      var b = e.target.closest && e.target.closest("button[data-mc]"); if (!b) return;
      var a = b.getAttribute("data-mc"), code = b.getAttribute("data-code");
      if (a === "folder") loadFromFolder();
      else if (a === "folder-choose") chooseFolder();
      else if (a === "open-log") pickLog();
      else if (a === "create-log") createLog().then(render, function (er) { window.alert("Could not create the log: " + (er && er.message || er)); });
      else if (a === "download") { if (C.log) { download(C.log.bytes, C.log.name); C.log.dirty = false; } }
      else if (a === "add-log") {
        var nw = items().filter(function (i) { return i.kind === "cert" && !i.inLog; });
        commit(nw.map(function (i) { return { t: "add", code: i.code, row: logRow(i) }; }), plural(nw.length, "certificate") + " added to the log");
      }
      else if (a === "record-folder") {
        var fo = items().filter(function (i) { return i.kind === "cert" && i.inLog && i.pf && !i.logPrinted; });
        commit(fo.map(function (i) { return { t: "print", code: i.code, at: i.pf.date ? dateOf(i.pf.date) : new Date(), by: "", reason: "Printed (from folder" + (i.pf.folder ? " “" + i.pf.folder + "”" : " / file name") + ")" }; }),
          plural(fo.length, "folder print") + " recorded");
      }
      else if (a === "print") markPrinted([byCode(code)]);
      else if (a === "preview") previewCert(code);
      else if (a === "preview-close") { var m = document.querySelector(".mos-prev-back"); if (m) m.remove(); }
      else if (a === "preview-link") {
        var key = b.getAttribute("data-key");
        C.links[code] = key;
        var it2 = byCode(code);
        if (it2 && it2.inLog) { var s2 = M(), r2 = s2 && s2.model ? s2.model.rows.filter(function (x) { return x.key === key; })[0] : null;
          commit([{ t: "link", code: code, ref: key === "none" ? "none" : (r2 && r2.ref) || "" }], "Link saved"); }
        else render();
        var m2 = document.querySelector(".mos-prev-back"); if (m2) m2.remove();
      }
      else if (a === "reprint") markPrinted([byCode(code)], { reprint: true });
      else if (a === "mark-sel") markPrinted(items().filter(function (i) { return C.sel[i.code] && inScope(i, false) && i.status !== "printed"; }));
      else if (a === "print-sel") printCertItems(items().filter(function (i) { return C.sel[i.code] && inScope(i, false); }));
      else if (a === "mark-file") {
        var sel = host.querySelector('[data-mc="file-pick"]'), file = sel ? sel.value : "";
        var list = items().filter(function (i) { return i.kind === "cert" && i.file === file && i.status === "pending"; });
        if (list.length && window.confirm("Mark all " + plural(list.length, "pending certificate") + " in " + file.split("/").pop() + " as printed?")) markPrinted(list, { confirmed: true });
      }
      else if (a === "export") exportXlsx();
      else if (a === "more") { C.limit += 500; render(); }
    });
  }
  window.addEventListener("beforeunload", function (e) {
    if (C.log && C.log.dirty && !C.log.writable) { e.preventDefault(); e.returnValue = "The certificate log has changes that are not saved to a file. Use “Download log”."; }
  });

  QBR.mosCerts = {
    render: render, items: items,
    has: function () { var s = M(); return !!(s && s.certs.length) || !!C.log; },
    reset: function () { C.links = {}; C.sel = {}; C.msg = ""; C.f = { school: "", examNo: "", cert: "", status: "", q: "" }; },
    // test seams
    _state: C, _applyOps: applyOps, _parseLog: parseLog, _newLogBytes: newLogBytes,
    _useLog: function (o) { useLog(o); }
  };
})();
