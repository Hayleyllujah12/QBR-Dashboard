/* =============================================================================
 * xlsx-merge.js — three-way, cell-level merge for linked workbooks (2026-10-06)
 * -----------------------------------------------------------------------------
 * When a linked workbook was changed in Excel after it was loaded, Save used to
 * refuse ("changed outside the dashboard") and the user had to re-link. This
 * module rebases the dashboard's edits onto the file as it is NOW:
 *
 *   base   = the file as loaded / last saved (bytes kept in memory)
 *   mine   = base + the dashboard's journaled edits
 *   theirs = the file on disk now (Excel edits included)
 *
 *   1. diff(base, mine)  → the dashboard's cell edits (base coordinates)
 *   2. map every edit onto theirs:
 *        sheet  → by name
 *        column → by HEADER TEXT (survives inserted / moved columns)
 *        row    → by a KEY column (serial, school, …; survives sorted / inserted
 *                 rows); rows the dashboard appended go after theirs' last row
 *   3. per edit: theirs cell == base  → apply (Excel didn't touch it)
 *                theirs cell == mine  → nothing to do
 *                otherwise            → CONFLICT (both changed) → user decides
 *      edits that can't be placed (row deleted, header renamed) → UNRESOLVED
 *   4. final = theirs + applied edits; xlsx-surgical writes only those cells.
 *
 * Pure (no DOM); Node-testable. Feature code can register row keys:
 *   QBR.mergeKeys.push({ sheet: /^07 PURCHASE/i, headers: ["po#"] })
 * ========================================================================== */
(function (root) {
  "use strict";
  const MERGE_VERSION = "0.1.0";
  const QBR = root.QBR = root.QBR || {};
  const X = () => root.XLSX;

  /* Preferred row-key headers (normalized, first match wins). A sheet-specific
   * registry entry beats these; if none fits, the first column whose values are
   * unique in both versions is used. */
  const DEFAULT_KEYS = ["serial number", "s/n", "serial", "sn", "ticket no", "ticket #", "ticket number", "item id", "item code",
    "school", "school name", "school (rakso managed)", "client", "organization", "po#", "sq", "domain", "timestamp"];
  QBR.mergeKeys = QBR.mergeKeys || [];

  const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();
  const keyNorm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toUpperCase();
  function serial(d) {
    const ms = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
    return (ms - Date.UTC(1899, 11, 30)) / 864e5;
  }
  function decodeStable(t) {
    let p; t = String(t == null ? "" : t);
    do { p = t; t = t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&"); } while (t !== p);
    return t;
  }
  /* Comparable form of a cell: value (dates as serials, numbers rounded),
   * formula, and hyperlink target. Blank = "". */
  function cv(c) {
    if (!c) return "";
    let v = "";
    if (c.f) v = "=" + String(c.f).replace(/^=/, "");
    else if (c.v instanceof Date) v = "n:" + Math.round(serial(c.v) * 1e6) / 1e6;
    else if (c.t === "d" && c.v) v = "n:" + Math.round(serial(new Date(c.v)) * 1e6) / 1e6;
    else if (typeof c.v === "number") v = "n:" + Math.round(c.v * 1e9) / 1e9;
    else if (typeof c.v === "boolean") v = "b:" + c.v;
    else if (c.v != null && c.v !== "" && c.t !== "z") v = "s:" + String(c.v);
    const l = c.l && c.l.Target ? "|l:" + decodeStable(c.l.Target) : "";
    return v + l;
  }
  const show = c => { if (!c) return ""; if (c.f) return "=" + c.f; if (c.v instanceof Date) return c.v.toISOString().slice(0, 10); return c.v == null ? "" : String(c.v); };

  function range(ws) { try { return X().utils.decode_range(ws && ws["!ref"] || "A1"); } catch (e) { return { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; } }
  const A = (r, c) => X().utils.encode_cell({ r, c });
  const cellAt = (ws, r, c) => ws ? ws[A(r, c)] : undefined;

  /* Header row = first row (of the first 10) with ≥2 text cells making up most of the row. */
  function headerRow(ws) {
    if (!ws) return 0;
    const g = range(ws);
    for (let r = g.s.r; r <= Math.min(g.e.r, g.s.r + 9); r++) {
      let text = 0, any = 0;
      for (let c = g.s.c; c <= g.e.c; c++) {
        const x = cellAt(ws, r, c); if (!x || x.v == null || x.v === "") continue;
        any++; if (typeof x.v === "string" && isNaN(Number(x.v))) text++;
      }
      if (text >= 2 && text >= any * 0.6) return r;
    }
    return g.s.r;
  }
  /* header name per column (duplicates get "#2", "#3"…) */
  function headers(ws, hr) {
    const g = range(ws), out = {}, seen = {};
    for (let c = 0; c <= g.e.c; c++) {
      const h = norm(cellAt(ws, hr, c) && cellAt(ws, hr, c).v);
      if (!h) continue;
      seen[h] = (seen[h] || 0) + 1;
      out[c] = seen[h] > 1 ? h + "#" + seen[h] : h;
    }
    return out;
  }
  const colOfHeader = (hmap, h) => { for (const c in hmap) if (hmap[c] === h) return +c; return -1; };
  function lastRow(ws) { // last row holding any value
    const g = range(ws); let last = -1;
    Object.keys(ws || {}).forEach(k => { if (k[0] === "!") return; const x = ws[k]; if (x && (x.v != null && x.v !== "" || x.f)) { const r = X().utils.decode_cell(k).r; if (r > last) last = r; } });
    return Math.max(last, -1, g.s.r - 1);
  }
  function lastCol(ws) {
    let last = -1;
    Object.keys(ws || {}).forEach(k => { if (k[0] === "!") return; const x = ws[k]; if (x && (x.v != null && x.v !== "" || x.f)) { const c = X().utils.decode_cell(k).c; if (c > last) last = c; } });
    return last;
  }

  /* Pick a key column (header name) whose values are unique in BOTH sheets. */
  function pickKey(sheetName, bWs, bHr, bH, tWs, tHr, tH) {
    const reg = (QBR.mergeKeys || []).filter(k => k && k.sheet && k.sheet.test(sheetName)).flatMap(k => (k.headers || []).map(norm));
    const order = reg.concat(DEFAULT_KEYS).concat(Object.keys(bH).sort((a, b) => a - b).map(c => bH[c]));
    const tried = new Set();
    for (const h of order) {
      if (tried.has(h)) continue; tried.add(h);
      const bc = colOfHeader(bH, h), tc = colOfHeader(tH, h);
      if (bc < 0 || tc < 0) continue;
      if (unique(bWs, bHr, bc) && unique(tWs, tHr, tc)) return { header: h, bc, tc };
    }
    return null;
  }
  function unique(ws, hr, c) {
    const last = lastRow(ws), seen = new Set(); let rows = 0, filled = 0;
    for (let r = hr + 1; r <= last; r++) {
      let rowAny = false;
      const g = range(ws);
      for (let cc = 0; cc <= g.e.c && !rowAny; cc++) { const x = cellAt(ws, r, cc); if (x && x.v != null && x.v !== "") rowAny = true; }
      if (!rowAny) continue;
      rows++;
      const v = keyNorm(cellAt(ws, r, c) && cellAt(ws, r, c).v);
      if (!v) continue;
      if (seen.has(v)) return false;
      seen.add(v); filled++;
    }
    return rows > 0 && filled >= rows * 0.8;
  }
  function keyIndex(ws, hr, c) {
    const m = new Map(), last = lastRow(ws);
    for (let r = hr + 1; r <= last; r++) { const v = keyNorm(cellAt(ws, r, c) && cellAt(ws, r, c).v); if (v && !m.has(v)) m.set(v, r); }
    return m;
  }

  function diffCells(bWs, mWs) {
    const keys = new Set(); [bWs || {}, mWs || {}].forEach(ws => Object.keys(ws).forEach(k => { if (k[0] !== "!") keys.add(k); }));
    const out = [];
    keys.forEach(k => { const b = bWs && bWs[k], m = mWs && mWs[k]; if (cv(b) !== cv(m)) { const p = X().utils.decode_cell(k); out.push({ r: p.r, c: p.c, base: b, mine: m }); } });
    return out.sort((a, b) => a.r - b.r || a.c - b.c);
  }
  function setRef(ws, r, c) {
    const g = range(ws); let first = !ws["!ref"];
    if (first) { ws["!ref"] = A(r, c); return; }
    g.e.r = Math.max(g.e.r, r); g.e.c = Math.max(g.e.c, c); ws["!ref"] = X().utils.encode_range(g);
  }
  const copyCell = c => { if (!c) return null; const n = { t: c.t, v: c.v }; if (c.f) n.f = c.f; if (c.l) n.l = { Target: c.l.Target }; if (c.z) n.z = c.z; return n; };

  /* ------------------------------------------------------------------ main */
  /* base, mine, theirs: SheetJS workbooks parsed with the same options.
   * final: a separate parse of theirs that receives the applied edits.
   * Returns { applied, conflicts[], unresolved[], sheetsAdded[], remapped } */
  function rebase(base, mine, theirs, final) {
    const res = { applied: 0, conflicts: [], unresolved: [], sheetsAdded: [], remapped: false, skippedSame: 0 };
    for (const name of mine.SheetNames) {
      const mWs = mine.Sheets[name], bWs = base.Sheets[name] || null, tWs = theirs.Sheets[name] || null;
      const edits = diffCells(bWs, mWs);
      if (!edits.length) continue;
      if (!tWs) {
        if (!bWs) {  // sheet created by the dashboard → add it whole
          X().utils.book_append_sheet(final, JSON.parse(JSON.stringify(mWs)), name);
          res.sheetsAdded.push(name); res.applied += edits.length; continue;
        }
        edits.forEach(e => res.unresolved.push({ sheet: name, ref: A(e.r, e.c), mine: show(e.mine), why: "sheet was renamed or deleted in Excel" }));
        continue;
      }
      const fWs = final.Sheets[name];
      // header rows + header maps (a sheet new to base uses mine's headers)
      const refWs = bWs || mWs;
      const bHr = headerRow(refWs), tHr = headerRow(tWs);
      const bH = headers(refWs, bHr), mH = headers(mWs, bHr), tH = headers(tWs, tHr);
      const key = pickKey(name, refWs, bHr, bH, tWs, tHr, tH);
      const tIdx = key ? keyIndex(tWs, tHr, key.tc) : null;
      const bLast = bWs ? lastRow(bWs) : bHr, tLast = lastRow(tWs), tLastC = lastCol(tWs);
      const sameShape = bWs && bLast - bHr === tLast - tHr;
      const rowMap = new Map(), colMap = new Map();
      let appendAt = tLast, newColAt = tLastC;
      function mapCol(c) {
        if (colMap.has(c)) return colMap.get(c);
        let out, why = null;
        const h = bH[c] || mH[c];
        if (h) {
          out = colOfHeader(tH, h);
          if (out < 0) {
            if (!bH[c]) out = ++newColAt;                 // column the dashboard added (e.g. Batch Code)
            else why = `column "${h}" was renamed or removed in Excel`;
          }
        } else out = c;                                   // no header: same position
        const v = why ? { why } : { c: out };
        colMap.set(c, v); return v;
      }
      function mapRow(r) {
        if (rowMap.has(r)) return rowMap.get(r);
        let v;
        if (r <= bHr) v = { r: r + (tHr - bHr) };          // header / title rows
        else if (key) {
          const kb = keyNorm(cellAt(bWs, r, key.bc) && cellAt(bWs, r, key.bc).v);
          const km = keyNorm(cellAt(mWs, r, colOfHeader(mH, key.header) >= 0 ? colOfHeader(mH, key.header) : key.bc) && cellAt(mWs, r, colOfHeader(mH, key.header) >= 0 ? colOfHeader(mH, key.header) : key.bc).v);
          const k = kb || km;
          if (k && tIdx.has(k)) v = { r: tIdx.get(k) };
          else if (r > bLast) v = { r: ++appendAt, appended: true };
          else v = { why: `row "${k || "row " + (r + 1)}" was deleted in Excel or its ${key.header} changed` };
        } else if (r > bLast) v = { r: ++appendAt, appended: true };
        else if (sameShape) v = { r: r + (tHr - bHr) };
        else v = { why: "rows were added/removed in Excel and this sheet has no unique ID column" };
        rowMap.set(r, v); return v;
      }
      for (const e of edits) {
        const mr = mapRow(e.r), mc = mapCol(e.c);
        const where = { sheet: name, header: bH[e.c] || mH[e.c] || "", key: key ? (keyNorm(cellAt(bWs || mWs, e.r, key.bc) && cellAt(bWs || mWs, e.r, key.bc).v) || keyNorm(cellAt(mWs, e.r, key.bc) && cellAt(mWs, e.r, key.bc).v)) : "" };
        if (mr.why || mc.why) { res.unresolved.push(Object.assign(where, { ref: A(e.r, e.c), mine: show(e.mine), why: mr.why || mc.why })); continue; }
        if (mr.r !== e.r || mc.c !== e.c) res.remapped = true;
        const tCell = cellAt(tWs, mr.r, mc.c), ref = A(mr.r, mc.c);
        const t = cv(tCell), b = cv(e.base), m = cv(e.mine);
        if (t === m) { res.skippedSame++; continue; }
        if (t === b || (mr.appended && t === "")) { apply(fWs, mr.r, mc.c, e.mine); res.applied++; continue; }
        res.conflicts.push(Object.assign(where, { ref, r: mr.r, c: mc.c, base: show(e.base), mine: show(e.mine), theirs: show(tCell), _mine: e.mine }));
      }
    }
    return res;
  }
  function apply(ws, r, c, cell) {
    const ref = A(r, c), n = copyCell(cell);
    if (!n || (n.v == null || n.v === "") && !n.f && !n.l) { if (ws[ref]) { ws[ref].v = ""; ws[ref].t = "s"; delete ws[ref].f; } }
    else { const old = ws[ref]; if (old && old.s) n.s = old.s; ws[ref] = n; }
    setRef(ws, r, c);
  }
  /* choices: { [conflictIndex]: "mine" | "theirs" } (default theirs) */
  function resolve(final, conflicts, choices) {
    let n = 0;
    conflicts.forEach((x, i) => { if ((choices || {})[i] === "mine") { apply(final.Sheets[x.sheet], x.r, x.c, x._mine); n++; } });
    return n;
  }

  const api = { MERGE_VERSION, rebase, resolve, _headerRow: headerRow, _pickKey: pickKey, _cv: cv };
  QBR.xlsxMerge = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
