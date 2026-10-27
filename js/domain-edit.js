/* ==================== DOMAIN REGISTRATION INLINE EDIT ====================
 * Edit Registration DATE and VALIDPD directly in the dashboard, writing back
 * to DOMAIN_REGISTRATION_TRACKER.xlsx via format-safe surgical save.
 *
 * Rules:
 * - NEVER raw-XLSX.write() the workbook: QBR.xlsxSurgical only (formulas preserved).
 * - Edit the SOURCES only (DATE, VALIDPD); Expiration is always derived.
 * - Row re-located by school+domain on every write (blank rows may shift indices).
 * - File re-read right before write (OneDrive-safe).
 */
(function () {
  "use strict";
  if (typeof QBR === "undefined") return;
  var XLSX = null;
  try { XLSX = window.XLSX; } catch (e) {}

  var S = {
    handle: null,   // FileSystemFileHandle for the tracker (picked by user)
    name: null,
    busy: false,
  };

  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function escAttr(s) { return esc(s).replace(/'/g, "&#39;"); }

  /* ---------- date helpers ---------- */
  // Parse "YYYY-MO-DD-HR-MIN-SEC" (or prefix) -> { y, mo, d, hh, mm, ss }
  function parseDateParts(v) {
    if (v == null) return null;
    var m = String(v).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[-\sT](\d{1,2})-(\d{1,2})-(\d{1,2}))?/);
    if (!m) return null;
    return { y: +m[1], mo: +m[2], d: +m[3], hh: +(m[4] || 0), mm: +(m[5] || 0), ss: +(m[6] || 0) };
  }
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  // Build "YYYY-MM-DD-HH-MM-SS" preserving the original time-of-day
  function buildDateStr(y, mo, d, origParts) {
    var t = origParts || { hh: 0, mm: 0, ss: 0 };
    return y + "-" + pad2(mo) + "-" + pad2(d) + "-" + pad2(t.hh) + "-" + pad2(t.mm) + "-" + pad2(t.ss);
  }
  function msFromParts(p) {
    var dt = new Date(p.y, p.mo - 1, p.d, p.hh, p.mm, p.ss);
    return isNaN(dt.getTime()) ? null : dt.getTime();
  }

  /* ---------- file picking ---------- */
  function canPick() { return typeof window.showOpenFilePicker === "function"; }

  async function pickTracker() {
    if (!canPick()) throw new Error("File picker not available in this browser.");
    var picks = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx", ".xlsm"] } }]
    });
    var h = picks[0];
    // Validate: must contain the domain sheet
    var f = await h.getFile();
    var bytes = new Uint8Array(await f.arrayBuffer());
    var wb = XLSX.read(bytes, { type: "array" });
    var sheet = findDomainSheet(wb);
    if (!sheet) throw new Error("That file has no Domain Registration sheet.");
    S.handle = h; S.name = h.name;
    return { handle: h, bytes: bytes, wb: wb, sheet: sheet };
  }

  function findDomainSheet(wb) {
    // Mirror the loader's findDomainRegSheet heuristics
    if (typeof QBR.hasDomainRegSheet === "function") {
      for (var i = 0; i < wb.SheetNames.length; i++) {
        var test = { SheetNames: [wb.SheetNames[i]], Sheets: {} };
        test.Sheets[wb.SheetNames[i]] = wb.Sheets[wb.SheetNames[i]];
        try { if (QBR.hasDomainRegSheet(test)) return wb.SheetNames[i]; } catch (e) {}
      }
      return null;
    }
    // Fallback: name match
    for (var j = 0; j < wb.SheetNames.length; j++) {
      if (/domain.*regist|regist.*domain/i.test(wb.SheetNames[j])) return wb.SheetNames[j];
    }
    return null;
  }

  /* ---------- reuse the app's linked-file handle ---------- */
  // The loaded-files list can LINK the tracker (Direct-save checkbox / folder
  // link): persist.js then already holds a writable handle in QBR._fsLinks.
  // Reuse it instead of asking the user to pick the same file again. A link
  // only counts if its file really contains the domain sheet.
  // Only links that carry the domain sheet (kind "domainreg", registered by the loader) or are
  // named like the tracker are opened — never every linked workbook (there can be dozens).
  function isDomainLink(l) {
    var kinds = (l.kinds || []).concat(Array.from(((QBR._kindByFp || {})[l.fp]) || []));
    return kinds.indexOf("domainreg") >= 0 ? 2 : /domain/i.test(l.name || "") ? 1 : 0;
  }
  async function linkedHandle() {
    var links = ((typeof QBR !== "undefined" && QBR._fsLinks) || []).filter(function (l) { return l && l.handle && isDomainLink(l); });
    links.sort(function (a, b) { return isDomainLink(b) - isDomainLink(a); });
    for (var i = 0; i < links.length; i++) {
      var h = links[i].handle;
      try {
        var perm = h.queryPermission ? await h.queryPermission({ mode: "readwrite" }) : "granted";
        if (perm !== "granted" && h.requestPermission) { try { perm = await h.requestPermission({ mode: "readwrite" }); } catch (e) {} }
        if (perm !== "granted") continue;
        var f = await h.getFile();
        var bytes = new Uint8Array(await f.arrayBuffer());
        var wb = XLSX.read(bytes, { type: "array" });
        if (findDomainSheet(wb)) { S.handle = h; S.name = h.name; S.link = links[i]; return true; }
      } catch (e) {}
    }
    return false;
  }

  /* ---------- after a successful write ---------- */
  // The app still holds the PRE-EDIT bytes (loaded-files copy + session cache) and the link
  // still describes the old file: bring both up to date, or a reload silently reverts the edit
  // on screen and the next save thinks the file was changed outside the dashboard.
  // The loaded copy is an IN-MEMORY File of the bytes just written (a File read from disk turns
  // unreadable once Excel or OneDrive touches the file — same rule as persist.js fsAfterSave).
  async function afterWrite(handle, out) {
    try {
      var nf = await handle.getFile();
      var name = handle.name || S.name || "DOMAIN_REGISTRATION_TRACKER.xlsx", nk = String(name).toLowerCase();
      var fpOf = (typeof QBR !== "undefined" && QBR.fpOf) ? QBR.fpOf : null;
      var oldFp = S.link && S.link.fp ? S.link.fp : null, newFp = null, mem = null;
      if (typeof APP !== "undefined" && APP.files) {
        // the loaded copy of this file: same fingerprint as the link, else same file name
        var idx = -1;
        if (oldFp && fpOf) APP.files.forEach(function (it, i) { if (idx < 0 && fpOf(it.name, it.blob) === oldFp) idx = i; });
        if (idx < 0) APP.files.forEach(function (it, i) { if (idx < 0 && String(it.name || "").toLowerCase() === nk) idx = i; });
        if (idx >= 0) {
          var it0 = APP.files[idx];
          if (!oldFp && fpOf) oldFp = fpOf(it0.name, it0.blob);
          mem = new File([out], it0.name, { lastModified: nf.lastModified, type: nf.type || "" });
          APP.files = APP.files.slice(); APP.files[idx] = { name: it0.name, blob: mem };
        }
      }
      if (!mem) mem = new File([out], name, { lastModified: nf.lastModified, type: nf.type || "" });
      newFp = fpOf ? fpOf(mem.name, mem) : null;
      if (oldFp && newFp && oldFp !== newFp) {
        if (QBR._currentFps) QBR._currentFps = QBR._currentFps.map(function (f) { return f === oldFp ? newFp : f; });
        if (QBR._kindByFp && QBR._kindByFp[oldFp]) { QBR._kindByFp[newFp] = QBR._kindByFp[oldFp]; delete QBR._kindByFp[oldFp]; }
        try { if (typeof QBR.journalMove === "function") QBR.journalMove(oldFp, newFp); } catch (e) {}
      }
      if (S.link) {
        S.link.fp = newFp || S.link.fp; S.link.size = mem.size; S.link.lastModified = mem.lastModified;
        try { if (typeof fsHashBytes === "function") S.link.hash = fsHashBytes(out); } catch (e) {}
        try { if (typeof fsSetBase === "function") fsSetBase(S.link.name || mem.name, out); } catch (e) {}   // base for the next three-way merge
        try { if (typeof fsLinksSave === "function") await fsLinksSave(); } catch (e) {}
        try { if (typeof QBR.fsRefreshStatus === "function") QBR.fsRefreshStatus(); } catch (e) {}
      }
      if (typeof cacheSession === "function") cacheSession();
    } catch (e) { try { console.warn("[QBR] domain edit: post-save bookkeeping failed:", e && e.message); } catch (x) {} }
  }

  /* ---------- row location ---------- */
  // Header = first non-empty row (the loader skips blank rows, so a title-less sheet may start lower down).
  function headerIndex(rows) {
    for (var i = 0; i < rows.length; i++) if ((rows[i] || []).some(function (v) { return v != null && String(v).trim() !== ""; })) return i;
    return -1;
  }
  // Find the Excel row (1-based) for a domain record by school+domain match.
  function locateRow(wb, sheet, rec) {
    var ws = wb.Sheets[sheet];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
    if (!rows.length) return -1;
    var o = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).s : { r: 0, c: 0 };   // sheets need not start at A1
    var hi = headerIndex(rows); if (hi < 0) return -1;
    var hdr = rows[hi].map(function (h) { return String(h || "").toLowerCase().trim(); });
    var cSchool = -1, cDomain = -1;
    hdr.forEach(function (h, i) {
      if (cSchool < 0 && /school/.test(h)) cSchool = i;
      if (cDomain < 0 && h === "domain") cDomain = i;
    });
    if (cSchool < 0) return -1;
    var key = function (s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); };
    var wantSchool = key(rec.schoolRaw), wantDomain = key(rec.domain);
    for (var r = hi + 1; r < rows.length; r++) {
      if (key(rows[r][cSchool]) === wantSchool) {
        if (cDomain < 0 || wantDomain === "" || key(rows[r][cDomain]) === wantDomain) return o.r + r + 1; // 1-based Excel row
      }
    }
    return -1;
  }

  function colLetter(idx) {
    var s = "";
    idx++;
    while (idx > 0) { var m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); }
    return s;
  }

  /* ---------- surgical write ---------- */
  async function surgicalWrite(bytes, wb, sheet, row1, colIdx, value) {
    var SG = QBR.xlsxSurgical;
    if (!SG) throw new Error("Format-safe saver not loaded.");
    // base and work MUST be parsed with the same options, or untouched cells (dates…) look changed and get rewritten
    var opt = { type: "array", cellStyles: true };
    wb = XLSX.read(bytes, opt);
    var work = XLSX.read(bytes, opt);
    var ws = work.Sheets[sheet];
    var addr = colLetter(colIdx) + row1;
    var cell = ws[addr] || {};
    if (cell.f) throw new Error("That cell is a formula in the tracker — edit it in Excel.");
    // keep the cell's own type: a number stays a number (e.g. an Excel date serial), text stays text
    ws[addr] = typeof value === "number" ? { t: "n", v: value } : { t: "s", v: String(value) };
    if (cell.z) ws[addr].z = cell.z;
    if (cell.s) ws[addr].s = cell.s;
    var rng = XLSX.utils.decode_range(ws["!ref"] || "A1");
    var r1 = row1 - 1;
    if (r1 > rng.e.r) rng.e.r = r1;
    if (colIdx > rng.e.c) rng.e.c = colIdx;
    ws["!ref"] = XLSX.utils.encode_range(rng);

    var res = await SG.surgicalSave(bytes, wb, work);
    if (res && res.ok && !res.noChanges && res.stats && res.stats.cells > 1) throw new Error("Safety stop: the save would change " + res.stats.cells + " cells.");
    if (!res || !res.ok) throw new Error("Surgical save failed.");
    if (res.noChanges) return bytes;
    var why = await SG.verifySurgical(bytes, res.bytes, work);
    if (why) throw new Error("Surgical verify failed: " + why);
    return res.bytes;
  }

  /* ---------- public: save an edit ---------- */
  // field: "reg" (DATE) or "valid" (VALIDPD). value: new DATE string or years number.
  async function saveEdit(rec, field, value) {
    if (S.busy) return { ok: false, error: "Busy." };
    if (!rec._sheet) return { ok: false, error: "No source coordinates for this row." };
    S.busy = true;
    try {
      var picked = null;
      if (!S.handle && !(await linkedHandle())) {
        if (!window.confirm("To save this edit, pick the DOMAIN_REGISTRATION_TRACKER.xlsx file.\n\nChoose the file now?")) {
          return { ok: false, error: "cancelled" };
        }
        picked = await pickTracker();
      }
      var handle = S.handle;
      // Re-read latest bytes (OneDrive-safe)
      var f = await handle.getFile();
      var bytes = new Uint8Array(await f.arrayBuffer());
      var wb = XLSX.read(bytes, { type: "array" });
      var sheet = findDomainSheet(wb);
      if (!sheet) return { ok: false, error: "Domain sheet not found in the picked file." };

      var row1 = locateRow(wb, sheet, rec);
      if (row1 < 0) return { ok: false, error: "Could not find this school/domain row in the file." };

      // Resolve column indices in the CURRENT file (headers may differ from load time)
      var ws = wb.Sheets[sheet];
      var c0 = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).s.c : 0;
      var allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
      var hdr = (allRows[headerIndex(allRows)] || []).map(function (h) { return String(h || "").toLowerCase().trim(); });
      // same column choice as the loader: exact header first, then contains
      var pick = function (names, not) {
        for (var k = 0; k < names.length; k++) { var e = hdr.indexOf(names[k]); if (e >= 0) return e; }
        for (k = 0; k < names.length; k++) for (var j = 0; j < hdr.length; j++) if (hdr[j].indexOf(names[k]) >= 0 && !(not && not.test(hdr[j]))) return j;
        return -1;
      };
      var cReg = pick(["registration status", "registration"]);
      var cValidPd = pick(["validpd", "valid pd"]);
      var cValidYr = pick(["validy in year", "validity in year", "validity period", "validity"], /valid\s*pd/);
      var here = function (c) { return c < 0 ? -1 : c + c0; };   // header index → sheet column
      cReg = here(cReg); cValidPd = here(cValidPd); cValidYr = here(cValidYr);
      var cur = function (c) { var a = ws[colLetter(c) + row1]; return a || null; };

      var out = bytes, changed = [];
      if (field === "reg") {
        if (cReg < 0) return { ok: false, error: "Registration column not found." };
        var rc = cur(cReg), v = value;
        if (rc && rc.t === "n") { var dp = parseDateParts(value); v = (Date.UTC(dp.y, dp.mo - 1, dp.d, dp.hh, dp.mm, dp.ss) / 864e5) + 25569; }  // the cell holds an Excel date
        out = await surgicalWrite(out, null, sheet, row1, cReg, v);
        changed.push("Registration DATE");
      } else if (field === "valid") {
        var cV = cValidPd >= 0 && !(cur(cValidPd) || {}).f ? cValidPd : cValidYr;
        if (cV < 0) return { ok: false, error: "Validity column not found." };
        var vc = cur(cV), yrs = +value, vv;
        if (vc && vc.t === "n") vv = yrs;
        else {   // keep the tracker's own text format: "VALIDPD: 01" → "VALIDPD: 02", "01" → "02"
          var t = vc && vc.v != null ? String(vc.v) : "", m = t.match(/^(.*?)(\d+)(\D*)$/);
          vv = m ? m[1] + (m[2].length > 1 ? ("0000" + yrs).slice(-m[2].length) : String(yrs)) + m[3] : String(yrs);
        }
        out = await surgicalWrite(out, null, sheet, row1, cV, vv);
        changed.push(cV === cValidPd ? "VALIDPD" : "Validity");
        // a plain (non-formula) "Validity in year" next to VALIDPD must agree, or the dashboard keeps reading the old value
        if (cV === cValidPd && cValidYr >= 0 && cValidYr !== cValidPd) {
          var yc = XLSX.read(out, { type: "array" }).Sheets[sheet][colLetter(cValidYr) + row1];
          if (yc && !yc.f) {
            var yt = String(yc.v == null ? "" : yc.v), ym = yt.match(/^(.*?)(\d+)(\D*)$/);
            var yv = yc.t === "n" ? yrs : ym ? ym[1] + (ym[2].length > 1 ? ("0000" + yrs).slice(-ym[2].length) : String(yrs)) + ym[3] : String(yrs);
            out = await surgicalWrite(out, null, sheet, row1, cValidYr, yv);
            changed.push("Validity in year");
          }
        }
      }
      var ex = pick(["expiration"]), exCell = ex >= 0 ? cur(here(ex)) : null;
      var expTyped = !!(exCell && !exCell.f && exCell.v != null && String(exCell.v).trim());

      // Write back to the file
      var w = null;
      try {
        w = await handle.createWritable();
        await w.write(out);
        await w.close();
        await afterWrite(handle, out);
      } catch (e) {
        try { if (w && w.abort) await w.abort(); } catch (x) {}
        return { ok: false, error: "Couldn't write — is the file open in Excel? Close it and try again." };
      }
      return { ok: true, changed: changed.join(", "), file: S.name || handle.name, expTyped: expTyped };
    } catch (e) {
      return { ok: false, error: (e && e.message) || String(e) };
    } finally {
      S.busy = false;
    }
  }

  /* ---------- inline edit UI ---------- */
  function fmtDateForInput(regRaw) {
    var p = parseDateParts(regRaw);
    if (!p) return "";
    return p.y + "-" + pad2(p.mo) + "-" + pad2(p.d);
  }

  function wireEdit() {
    document.addEventListener("click", function (e) {
      var cell = e.target && e.target.closest ? e.target.closest("[data-dr-edit]") : null;
      if (!cell || cell.querySelector("input")) return; // already editing
      var field = cell.getAttribute("data-dr-edit"); // "reg" | "valid"
      var key = cell.getAttribute("data-dr-key");
      var rec = findRec(key);
      if (!rec) return;
      startEdit(cell, rec, field);
    });
  }

  // key = "i<index>" (position in APP.model.domainreg) or "<school key>|<domain>"; never the school alone
  // (a school can have several domains — the first one used to be edited).
  function findRec(key) {
    var rows = (typeof APP !== "undefined" && APP.model && APP.model.domainreg) || [];
    var m = /^i(\d+)$/.exec(key || "");
    if (m) return rows[+m[1]] || null;
    for (var i = 0; i < rows.length; i++) if (rows[i].key + "|" + (rows[i].domain || "") === key) return rows[i];
    return null;
  }

  function startEdit(cell, rec, field) {
    var orig = cell.innerHTML;
    var input, label;
    if (field === "reg") {
      label = "Registration date";
      input = document.createElement("input");
      input.type = "date";
      input.className = "form-control form-control-sm";
      input.value = fmtDateForInput(rec.regRaw);
    } else {
      label = "Validity (years)";
      input = document.createElement("input");
      input.type = "number";
      input.min = "1"; input.max = "20"; input.step = "1";
      input.className = "form-control form-control-sm";
      input.style.maxWidth = "90px";
      input.value = rec.validYr || "";
    }
    cell.innerHTML = "";
    cell.appendChild(input);
    input.focus();
    if (input.select) input.select();

    var cancelled = false;
    function cleanup() { cell.innerHTML = orig; }
    input.addEventListener("keydown", function (ev) {
      if (ev.key === "Escape") { cancelled = true; cleanup(); }
      else if (ev.key === "Enter") { ev.preventDefault(); commit(); }
    });
    input.addEventListener("blur", function () { if (!cancelled) commit(); });

    async function commit() {
      if (cancelled) return;
      cancelled = true; // prevent double-commit from blur+enter
      var rawVal = input.value.trim();
      if (!rawVal) { cleanup(); return; }
      cell.innerHTML = '<span class="text-muted">Saving…</span>';
      var result;
      if (field === "reg") {
        var parts = rawVal.split("-");
        var origParts = parseDateParts(rec.regRaw);
        var dateStr = buildDateStr(+parts[0], +parts[1], +parts[2], origParts);
        result = await saveEdit(rec, "reg", dateStr);
        if (result.ok) {
          rec.regRaw = dateStr;
          rec.regMs = msFromParts(parseDateParts(dateStr));
          refreshDerived(rec);
        }
      } else {
        var yrs = parseInt(rawVal, 10);
        if (!(yrs >= 1 && yrs <= 20)) { alert("Validity must be 1–20 years."); cleanup(); return; }
        result = await saveEdit(rec, "valid", String(yrs));
        if (result.ok) {
          rec.validYr = yrs;
          refreshDerived(rec);
        }
      }
      if (result.ok) {
        // KPIs, status chart and table all follow the new dates
        if (typeof renderAll === "function") { try { renderAll(); } catch (x) {} }
        else if (typeof renderDrTable === "function") renderDrTable((document.getElementById("dr-search") || {}).value || "");
        if (result.expTyped) alert("Saved (" + result.changed + ").\n\nNote: this row's Expiration in the tracker is a typed value, not a formula, so Excel will keep showing the old expiration. Update that cell in Excel.");
      } else if (result.error !== "cancelled") {
        alert("Could not save: " + result.error);
        cleanup();
      } else {
        cleanup();
      }
    }
  }

  // Recompute expiration + status after an edit (mirrors the loader's derivation)
  function refreshDerived(rec) {
    var e = rec.regMs, yrs = rec.validYr || 1;
    for (var i = 0; i < yrs; i++) {
      var d = new Date(e); d.setFullYear(d.getFullYear() + 1); e = d.getTime();
    }
    // Always recompute here: an Expiration FORMULA in the file only updates when Excel recalculates (on next open)
    if (typeof QBR !== "undefined" && typeof QBR.domainStatus === "function") {
      rec.expMs = e;
      rec.srcDays = null;   // the file's "N days remaining" text is stale after an edit
      var st = QBR.domainStatus(rec);
      rec.status = st.status; rec.tier = st.tier; rec.action = st.action; rec.days = st.days; rec.group = st.group;
    } else {
      rec.expMs = e;
    }
  }

  QBR.domainEdit = {
    save: saveEdit,
    wire: wireEdit,
    status: function () { return { linked: !!S.handle, file: S.name }; },
  };
})();
