/* RCT OpsDesk — GDAP mapping (js/gdap.js)
 *
 * Loads Pedro's GranularAdministerRelationship CSV (headers: Name,Microsoft ID)
 * and maps an audit school name -> tenant ID for script tenant pre-fill.
 * The CSV never leaves the browser: parsed rows are kept in localStorage.
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

  /* Parse the GDAP CSV text -> [{ name, tenantId }]. Throws on bad headers. */
  function gdapParseCsv(text) {
    var lines = String(text || "").split(/\r?\n/);
    if (!lines.length || !lines[0].trim()) throw new Error("Empty file.");
    var header = splitCsvLine(lines[0]).map(function (h) {
      return h.replace(/^\uFEFF/, "").trim().toLowerCase();
    });
    var nameIx = header.indexOf("name");
    var idIx = header.indexOf("microsoft id");
    if (nameIx < 0 || idIx < 0) throw new Error('Expected headers "Name" and "Microsoft ID".');
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      var cells = splitCsvLine(lines[i]);
      var name = (cells[nameIx] || "").trim();
      var tid = (cells[idIx] || "").trim();
      if (name && tid) rows.push({ name: name, tenantId: tid });
    }
    return rows;
  }

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
  QBR.gdapSave = gdapSave;
  QBR.gdapLoad = gdapLoad;
  QBR.gdapClear = gdapClear;
  QBR.gdapLookup = gdapLookup;
})();
