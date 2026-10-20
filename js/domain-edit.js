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

  /* ---------- row location ---------- */
  // Find the Excel row (1-based) for a domain record by school+domain match.
  function locateRow(wb, sheet, rec) {
    var ws = wb.Sheets[sheet];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
    if (!rows.length) return -1;
    var hdr = rows[0].map(function (h) { return String(h || "").toLowerCase().trim(); });
    var cSchool = -1, cDomain = -1;
    hdr.forEach(function (h, i) {
      if (cSchool < 0 && /school/.test(h)) cSchool = i;
      if (cDomain < 0 && h === "domain") cDomain = i;
    });
    if (cSchool < 0) return -1;
    var key = function (s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); };
    var wantSchool = key(rec.schoolRaw), wantDomain = key(rec.domain);
    for (var r = 1; r < rows.length; r++) {
      if (key(rows[r][cSchool]) === wantSchool) {
        if (cDomain < 0 || wantDomain === "" || key(rows[r][cDomain]) === wantDomain) return r + 1; // 1-based
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
    // Clone workbook, apply the single-cell change
    var work = XLSX.read(bytes, { type: "array", cellDates: true });
    var ws = work.Sheets[sheet];
    var addr = colLetter(colIdx) + row1;
    var cell = ws[addr] || {};
    // Preserve the cell's number format; write as string to match the tracker's text format
    ws[addr] = { t: "s", v: String(value) };
    if (cell.z) ws[addr].z = cell.z;
    if (cell.s) ws[addr].s = cell.s;
    var rng = XLSX.utils.decode_range(ws["!ref"] || "A1");
    var r1 = row1 - 1;
    if (r1 > rng.e.r) rng.e.r = r1;
    if (colIdx > rng.e.c) rng.e.c = colIdx;
    ws["!ref"] = XLSX.utils.encode_range(rng);

    var res = await SG.surgicalSave(bytes, wb, work);
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
      if (!S.handle) {
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
      var hdr = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true })[0] || [])
        .map(function (h) { return String(h || "").toLowerCase().trim(); });
      var cReg = -1, cValidPd = -1, cValidYr = -1;
      hdr.forEach(function (h, i) {
        if (cReg < 0 && /registration/.test(h)) cReg = i;
        if (cValidPd < 0 && /valid\s*pd|validpd/.test(h)) cValidPd = i;
        if (cValidYr < 0 && /valid.*year|validity/.test(h) && !/valid\s*pd/.test(h)) cValidYr = i;
      });

      var out = bytes, changed = [];
      if (field === "reg") {
        if (cReg < 0) return { ok: false, error: "Registration column not found." };
        out = await surgicalWrite(out, XLSX.read(out, { type: "array" }), sheet, row1, cReg, value);
        changed.push("Registration DATE");
      } else if (field === "valid") {
        var cV = cValidPd >= 0 ? cValidPd : cValidYr;
        if (cV < 0) return { ok: false, error: "Validity column not found." };
        out = await surgicalWrite(out, XLSX.read(out, { type: "array" }), sheet, row1, cV, value);
        changed.push(cValidPd >= 0 ? "VALIDPD" : "Validity");
      }

      // Write back to the file
      var w = null;
      try {
        w = await handle.createWritable();
        await w.write(out);
        await w.close();
      } catch (e) {
        try { if (w && w.abort) await w.abort(); } catch (x) {}
        return { ok: false, error: "Couldn't write — is the file open in Excel? Close it and try again." };
      }
      return { ok: true, changed: changed.join(", "), file: S.name || handle.name };
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

  function findRec(key) {
    var rows = (typeof APP !== "undefined" && APP.model && APP.model.domainreg) || [];
    for (var i = 0; i < rows.length; i++) if (rows[i].key === key) return rows[i];
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
        if (typeof renderDrTable === "function") {
          var q = (document.getElementById("dr-search") || {}).value || "";
          renderDrTable(q);
        }
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
    // Only override if there was no explicit expiration column value
    if (typeof QBR !== "undefined" && typeof QBR.domainStatus === "function") {
      var hadExplicit = rec._hadExplicitExp;
      if (!hadExplicit) rec.expMs = e;
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
