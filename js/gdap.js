/* RCT OpsDesk — GDAP mapping (js/gdap.js)
 *
 * Loads Pedro's GranularAdministerRelationship mapping: .xlsx/.xlsm (v1.39.0) or the
 * legacy .csv (headers: Name, Microsoft ID), and maps an audit school name -> tenant ID for script tenant pre-fill.
 * The file never leaves the browser: parsed rows are kept in localStorage.
 */
var QBR = (window.QBR = window.QBR || {});
(function () {
  "use strict";

  var LS_KEY = "qbr.gdap.v1";

  function splitCsvLine(line) {
    var out = [], cur = "", inQ = false, i, c;
    for (i = 0; i < line.length; i++) {
      c = line[i];
      if (inQ) {
        if (c === '"') {
          if (line[i + 1] === '"') { cur += '"'; i++; }
          else inQ = false;
        } else cur += c;
      } else if (c === '"') inQ = true;
      else if (c === ",") { out.push(cur); cur = ""; }
      else cur += c;
    }
    out.push(cur);
    return out;
  }

  var NAME_HDR = ["name", "school", "school name", "customer", "customer name", "organization"];
  var ID_HDR = ["microsoft id", "tenant id", "tenantid", "customer tenant id", "microsoft tenant id", "directory id"];

  function hdr(h) { return String(h == null ? "" : h).replace(/^\uFEFF/, "").trim().toLowerCase(); }

  /* Rows (array of arrays) -> [{ name, tenantId }], or null when no header row
     with a Name + Microsoft ID column is found in the first 10 rows. */
  function gdapParseAoa(aoa) {
    aoa = aoa || [];
    for (var h = 0; h < Math.min(10, aoa.length); h++) {
      var header = (aoa[h] || []).map(hdr);
      var nameIx = -1, idIx = -1, k;
      for (k = 0; k < NAME_HDR.length && nameIx < 0; k++) nameIx = header.indexOf(NAME_HDR[k]);
      for (k = 0; k < ID_HDR.length && idIx < 0; k++) idIx = header.indexOf(ID_HDR[k]);
      if (nameIx < 0 || idIx < 0) continue;
      var rows = [];
      for (var i = h + 1; i < aoa.length; i++) {
        var cells = aoa[i] || [];
        var name = String(cells[nameIx] == null ? "" : cells[nameIx]).trim();
        var tid = String(cells[idIx] == null ? "" : cells[idIx]).trim();
        if (name && tid) rows.push({ name: name, tenantId: tid });
      }
      return rows;
    }
    return null;
  }

  /* Parse the GDAP CSV text -> [{ name, tenantId }]. Throws on bad headers. */
  function gdapParseCsv(text) {
    var lines = String(text || "").split(/\r?\n/);
    if (!lines.length || !lines[0].trim()) throw new Error("Empty file.");
    var aoa = lines.filter(function (l) { return l.trim(); }).map(splitCsvLine);
    var rows = gdapParseAoa(aoa);
    if (!rows) throw new Error('Expected headers "Name" and "Microsoft ID".');
    return rows;
  }

  /* Parse a GDAP .xlsx/.xlsm (Uint8Array | ArrayBuffer | SheetJS workbook) ->
     [{ name, tenantId }] from the first sheet that has Name + Microsoft ID headers.
     Throws when no sheet has those headers. */
  function gdapParseWorkbook(data) {
    if (typeof XLSX === "undefined") throw new Error("Excel reader not loaded.");
    var wb = (data && data.SheetNames) ? data : XLSX.read(data, { type: "array" });
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var ws = wb.Sheets[wb.SheetNames[i]];
      if (!ws) continue;
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "", blankrows: false });
      var rows = gdapParseAoa(aoa);
      if (rows) return rows;
    }
    throw new Error('No sheet with "Name" and "Microsoft ID" columns.');
  }

  /* File names treated as a GDAP mapping workbook (GDAP.xlsx, GranularAdministerRelationship.xlsx …). */
  function gdapIsMappingName(name) { return /gdap|granular/i.test(String(name || "")); }

  function gdapSave(rows) {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ rows: rows, savedAt: new Date().toISOString() }));
    } catch (e) { /* storage full/blocked */ }
  }

  function gdapLoad() {
    try {
      var d = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      return (d && Array.isArray(d.rows)) ? d : null;
    } catch (e) { return null; }
  }

  function gdapClear() {
    try { localStorage.removeItem(LS_KEY); } catch (e) {}
  }

  function norm(s) { return (s || "").trim().toLowerCase(); }

  /* schoolName -> { name, tenantId } or null.
     Exact (case-insensitive) match first, then a single contains-match. */
  function gdapLookup(schoolName) {
    var d = gdapLoad();
    if (!d || !d.rows || !schoolName) return null;
    var q = norm(schoolName);
    if (!q) return null;
    var i, r, hit = null;
    for (i = 0; i < d.rows.length; i++) {
      r = d.rows[i];
      if (norm(r.name) === q) { hit = r; break; }
    }
    if (!hit) {
      var cands = [];
      for (i = 0; i < d.rows.length; i++) {
        r = d.rows[i];
        var n = norm(r.name);
        if (n.indexOf(q) >= 0 || q.indexOf(n) >= 0) cands.push(r);
      }
      if (cands.length === 1) hit = cands[0];
    }
    return hit ? { name: hit.name, tenantId: hit.tenantId } : null;
  }

  /* Explicitly exposed entry points. */
  QBR.gdapParseCsv = gdapParseCsv;
  QBR.gdapParseAoa = gdapParseAoa;
  QBR.gdapParseWorkbook = gdapParseWorkbook;
  QBR.gdapIsMappingName = gdapIsMappingName;
  QBR.gdapSave = gdapSave;
  QBR.gdapLoad = gdapLoad;
  QBR.gdapClear = gdapClear;
  QBR.gdapLookup = gdapLookup;
})();
