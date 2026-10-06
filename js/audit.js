/* ============================================================================
 * QBR — Audit editing module (js/audit.js)
 * v1.26.0 — Report 1: Risky Sign-ins (RISKY_USERS_AND_DOMAIN.xlsx)
 *
 * Lets Pedro edit his audit workbooks directly in the dashboard:
 *  - Month picker (defaults to the current month, cycle through all 12)
 *  - Editable school table: Total Risky Users (number/text), Domain Health
 *    (dropdown from the workbook's Drop-Down sheet)
 *  - References renders as a clickable SharePoint link
 *  - Add-school rows; all edits journaled (refresh-proof) and saved via the
 *    patch-in-place flow, preserving formulas and layout
 *  - Background external-change watcher for linked files (OneDrive sync)
 *
 * The module is self-contained behind window.QBR. The host shell calls:
 *   renderAudit()            — renders the dash-audit panel
 * and processBuffers() calls QBR.parseAuditBuffers(buffers).
 * ========================================================================== */
(function () {
  "use strict";

  QBR.AUDIT_VERSION = "1.29.0";
  // Base used to resolve Excel-stored relative hyperlink targets (e.g.
  // "../../../../../../:x:/r/sites/..." -> "file:///C:/:x:/r/sites/..."). The
  // browser cannot see the workbook's local folder, so relative links are
  // resolved against this OneDrive sync drive. Change it if the sync moves.
  QBR.AUDIT_LINK_BASE = QBR.AUDIT_LINK_BASE || "file:///C:/";


  const AUDIT_MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];

  // Column keys used in journal ops (stable even if column order shifts).
  const COL_KEYS = ["school", "org", "risky", "health", "ref"];

  const norm = s => String(s == null ? "" : s).trim().replace(/\s+/g, " ").toUpperCase();

  function normHeader(h) {
    return norm(h).replace(/[()]/g, "").replace(/\s+/g, " ").trim();
  }

  function looksLikeUrl(v) {
    return /^(https?:\/\/|www\.)/i.test(String(v || "").trim());
  }

  // Map a header row (array of raw values) to column indices per COL_KEYS.
  function mapColumns(headerRow) {
    const idx = {};
    (headerRow || []).forEach((h, i) => {
      const n = normHeader(h);
      if (idx.school == null && /SCHOOL/.test(n)) idx.school = i;
      else if (idx.org == null && n === "ORGANIZATION") idx.org = i;
      else if (idx.risky == null && /TOTAL RISKY USERS/.test(n)) idx.risky = i;
      else if (idx.health == null && /DOMAIN HEALTH/.test(n)) idx.health = i;
      else if (idx.ref == null && /REFERENCES/.test(n)) idx.ref = i;
      else if (idx.exempt == null && /^EXEMPT$/.test(n)) idx.exempt = i;
      else if (idx.exemptReason == null && /EXEMPT REASON/.test(n)) idx.exemptReason = i;
    });
    // Exempt columns live far right (K/L) by convention; default there if headers absent.
    if (idx.exempt == null) idx.exempt = 10;
    if (idx.exemptReason == null) idx.exemptReason = 11;
    return idx;
  }

  function isRiskyWorkbook(sheetNames) {
    const names = sheetNames.map(norm);
    const hits = AUDIT_MONTHS.filter(m => names.indexOf(m) >= 0).length;
    return hits >= 6; // at least half the months present
  }

  function findSheet(sheets, candidates) {
    for (const cand of candidates) {
      const hit = Object.keys(sheets).find(n => norm(n) === norm(cand));
      if (hit) return hit;
    }
    return null;
  }

  // Find a school's data row within a month model (case-insensitive).
  function auditFindRow(a, month, school) {
    try {
      const ms = a.months[month];
      if (!ms) return null;
      const key = String(school || "").trim().toLowerCase();
      return ms.rows.find(r => String(r.school || "").trim().toLowerCase() === key) || null;
    } catch (e) { return null; }
  }
  // Previous calendar month name from AUDIT_MONTHS (null for January).
  function auditPrevMonth(month) {
    const i = AUDIT_MONTHS.indexOf(String(month || "").toUpperCase());
    return i > 0 ? AUDIT_MONTHS[i - 1] : null;
  }
  // Month -> quarter (storage/usage workbooks are quarterly).
  function auditMonthToQuarter(month) {
    const m = String(month || "").toUpperCase();
    if (/JAN|FEB|MAR/.test(m)) return "Q1";
    if (/APR|MAY|JUN/.test(m)) return "Q2";
    if (/JUL|AUG|SEP/.test(m)) return "Q3";
    if (/OCT|NOV|DEC/.test(m)) return "Q4";
    return null;
  }
  // "Fully audited" = risky count + health + storage data + usage data.
  // Storage: 7 values in cols C-I (idx 2-8); Usage: 19 values in cols D-V (idx 3-21).
  function auditHasQuarterData(kind, month, school) {
    try {
      const mdl = kind === "storage" ? QBR._storage : QBR._usage;
      const q = auditMonthToQuarter(month);
      if (!mdl || !q || !mdl.quarters[q]) return false;
      const key = String(school || "").trim().toLowerCase();
      const rec = mdl.quarters[q].rows.find(r => String(r.school || "").trim().toLowerCase() === key);
      if (!rec || !rec._row) return false;
      const lo = kind === "storage" ? 2 : 3, hi = kind === "storage" ? 8 : 21;
      for (let i = lo; i <= hi; i++) {
        const v = rec._row[i];
        if (v != null && String(v).trim() !== "") return true;
      }
      return false;
    } catch (e) { return false; }
  }
  function auditIsFullyAudited(a, month, school) {
    try {
      const rec = auditFindRow(a, month, school);
      if (!rec) return false;
      const hasRisky = rec.risky != null && String(rec.risky).trim() !== "";
      const hasHealth = rec.health != null && String(rec.health).trim() !== "";
      if (!hasRisky || !hasHealth) return false;
      return auditHasQuarterData("storage", month, school) && auditHasQuarterData("usage", month, school);
    } catch (e) { return false; }
  }
  // "Audited" (legacy, risky-only) — kept for the progress bar label.
  function auditIsAudited(rec) {
    const v = rec && rec.risky;
    return v != null && String(v).trim() !== "";
  }

  /* ---------- model ------------------------------------------------------ */
  // QBR._audit = { fp, fileName, months: {NAME: {sheet, colIdx, rows: [...]}},
  //                dropdowns: {org: [...], health: [...]}, loadedAt }
  // row = { r: 1-indexed sheet row, school, org, risky, health, ref, refUrl }

  // Scan raw worksheet hyperlink targets for a full https://*.sharepoint.com
  // URL. Excel strips the host from most sharing links (leaving ":x:/r/..."),
  // but usually a few survive intact — their host becomes the base used to
  // rebuild the stripped ones.
  function detectSpBase(raw) {
    try {
      for (const n of Object.keys(raw || {})) {
        const ws = raw[n];
        for (const addr of Object.keys(ws)) {
          if (addr[0] === "!") continue;
          const t = ws[addr].l && ws[addr].l.Target;
          const m = t && String(t).match(/^(https:\/\/[^\/]*\.sharepoint\.com)/i);
          if (m) return m[1];
        }
      }
    } catch (e) {}
    return "";
  }

  // Decode XML entities until stable (undoes any layers of double-encoding).
  function decodeEntitiesStable(t) {
    let prev;
    do {
      prev = t;
      t = t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");
    } while (t !== prev);
    return t;
  }

  // CRITICAL (2026-10-06): SheetJS returns hyperlink targets with XML entities
  // still encoded, and re-encodes on write. Without this, every Save/Export
  // adds another &amp; layer to all links in the workbook (progressive corruption).
  // Call before EVERY XLSX.write of a workbook that may contain hyperlinks.
  QBR.sanitizeHyperlinkTargets = function (wb) {
    try {
      (wb.SheetNames || []).forEach(n => {
        const ws = wb.Sheets[n];
        for (const addr of Object.keys(ws)) {
          if (addr[0] === "!") continue;
          const cell = ws[addr];
          if (cell.l && cell.l.Target) cell.l.Target = decodeEntitiesStable(String(cell.l.Target));
        }
      });
    } catch (e) {}
    return wb;
  };

  // Normalize an Excel hyperlink target for dashboard use:
  //  - SheetJS leaves XML entities encoded in targets; decode first so the
  //    renderer's HTML-escaping does not double-escape them (&amp;amp;).
  //  - Excel stores local-file links relative to the workbook
  //    ("../../../../:x:/r/..."); resolve against QBR.AUDIT_LINK_BASE.
  //  Absolute http(s) links pass through untouched.
  function auditNormUrl(u) {
    let s = String(u == null ? "" : u);
    if (!s) return s;
    s = decodeEntitiesStable(s);
    const m = s.match(/^((?:\.\.\/)+)([\s\S]*)$/);
    if (m) {
      const rest = String(m[2] || "").replace(/^\/*/, "");
      // Stripped SharePoint sharing link (":x:/r/...", ":i:/s/...", …):
      // reattach the tenant host detected from intact links.
      if (/^:[a-z]:\//i.test(rest) && QBR._auditSpBase) {
        s = String(QBR._auditSpBase).replace(/\/*$/, "") + "/" + rest;
      } else {
        const base = String(QBR.AUDIT_LINK_BASE || "file:///C:/").replace(/\/*$/, "/");
        s = base + rest;
      }
    }
    return s;
  }
  QBR.auditNormUrl = auditNormUrl;

  function parseMonthSheet(ws, rows, colIdx) {

    const out = [];
    const cSchool = colIdx.school, cOrg = colIdx.org, cRisky = colIdx.risky,
          cHealth = colIdx.health, cRef = colIdx.ref,
          cExempt = colIdx.exempt, cExemptReason = colIdx.exemptReason;
    rows.slice(1).forEach((row, i) => {
      const r = i + 2; // header is row 1, so first data row is row 2
      const school = cSchool != null ? row[cSchool] : null;
      if (school == null || String(school).trim() === "") return;
      const cell = (c) => (c != null && row[c] != null ? String(row[c]).trim() : "");
      let refUrl = "";
      try {
        if (cRef != null && ws) {
          const addr = XLSX.utils.encode_cell({ r: r - 1, c: cRef });
          const raw = ws[addr];
          if (raw && raw.l && raw.l.Target) refUrl = auditNormUrl(raw.l.Target);
        }
      } catch (e) { /* hyperlink read is best-effort */ }
      const exemptRaw = cell(cExempt).toLowerCase();
      out.push({
        r: r,
        school: cell(cSchool), org: cell(cOrg), risky: cell(cRisky),
        health: cell(cHealth), ref: cell(cRef), refUrl: refUrl,
        exempt: (exemptRaw === "yes" || exemptRaw === "y" || exemptRaw === "true" || exemptRaw === "1"),
        exemptReason: cell(cExemptReason),
      });
    });
    return out;
  }

  function parseDropdowns(sheets) {
    const dd = { org: [], health: [] };
    const name = findSheet(sheets, ["Drop-Down", "Drop Down", "Dropdown", "Lists"]);
    if (!name) return dd;
    const rows = sheets[name] || [];
    const seen = { org: {}, health: {} };
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const o = row[0] != null ? String(row[0]).trim() : "";
      const h = row[1] != null ? String(row[1]).trim() : "";
      // The sheet holds a second table below the lists (a "Tenants"/"Organization"
      // school→org mapping). Stop there so mapping values don't pollute the dropdowns.
      if (o.toLowerCase() === "tenants") break;
      if (o && !seen.org[o]) { seen.org[o] = 1; dd.org.push(o); }
      if (h && !seen.health[h]) { seen.health[h] = 1; dd.health.push(h); }
    }
    return dd;
  }

  QBR.auditParseSheets = function (sheets, rawSheets, opts) {
    opts = opts || {};
    const names = Object.keys(sheets || {});
    if (!isRiskyWorkbook(names)) return null;
    const fp = opts.fp || null;
    const model = { fp: fp, fileName: opts.fileName || "", months: {}, dropdowns: parseDropdowns(sheets) };
    AUDIT_MONTHS.forEach(m => {
      const hit = names.find(n => norm(n) === m);
      if (!hit) return;
      const rows = sheets[hit] || [];
      if (!rows.length) return;
      const colIdx = mapColumns(rows[0]);
      if (colIdx.school == null) return; // not a data month sheet
      model.months[m] = {
        sheet: hit,
        colIdx: colIdx,
        rows: parseMonthSheet(rawSheets ? rawSheets[hit] : null, rows, colIdx),
      };
    });
    if (!Object.keys(model.months).length) return null;
    return model;
  };

  /* Parse raw workbook buffers into the audit model. Mirrors the inventory
   * pattern: per-buffer, filename-independent, sheets by name. */
  QBR.parseAuditBuffers = function (buffers) {
    const fps = QBR._currentFps || [];
    let model = null;
    let rawAll = {};
    (buffers || []).forEach((buf, bi) => {
      try {
        const wb = XLSX.read(buf, { type: "array", cellStyles: true });
        const fp = fps[bi];
        const sheets = {}, raw = {};
        wb.SheetNames.forEach(n => {
          raw[n] = wb.Sheets[n];
          sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false });
        });
        let fileName = "";
        try {
          const list = (typeof persistAppFiles === "function") ? persistAppFiles() : [];
          const f = list.filter(x => QBR.fpOf(x.name, x.blob) === fp)[0];
          if (f) fileName = f.name;
        } catch (e) {}
        Object.keys(raw).forEach(n => { rawAll[n] = raw[n]; });
        try {
          const base = detectSpBase(raw);
          if (base) QBR._auditSpBase = base;
        } catch (e) {}
        const m = QBR.auditParseSheets(sheets, raw, { fp: fp || null, fileName: fileName });
        if (m) {
          if (fp) { (QBR._origWb || (QBR._origWb = {}))[fp] = wb; }
          model = m; // last matching file wins (same convention as inventory)
          if (fp) {
            const set = QBR._kindByFp[fp] || (QBR._kindByFp[fp] = new Set());
            set.add("audit");
          }
        }
      } catch (e) { /* a non-audit workbook simply contributes nothing */ }
    });
    QBR._audit = model;
    QBR._auditMeta = model ? buildMeta(model) : null;
    // Rebuild base for stripped SharePoint links.
    try { QBR._auditSpBase = detectSpBase(rawAll) || QBR._auditSpBase || ""; } catch (e) {}
    return model;
  };

  function buildMeta(model) {
    const meta = { fp: model.fp, sheets: {} };
    Object.keys(model.months).forEach(m => {
      const ms = model.months[m];
      meta.sheets[ms.sheet] = { month: m, colIdx: ms.colIdx, headerRow: 1 };
    });
    return meta;
  }

  function auditModel() { return QBR._audit || null; }
  function monthState(m) {
    const a = auditModel();
    return (a && a.months && a.months[m]) || null;
  }

  /* ---------- journaled ops ---------------------------------------------- */
  // All ops are replayable: QBR.journalReplayFor calls QBR[op](...args).

  function recordOp(op, args) {
    if (typeof QBR.journalRecord === "function") QBR.journalRecord("audit", op, args);
  }

  // Update one cell. colKey in COL_KEYS. For ref, pass {text, url} as value.
  QBR.auditUpdateCell = function (fp, sheet, row, colKey, value, url) {
    const a = auditModel();
    if (!a || a.fp !== fp) return false;
    const ms = Object.values(a.months).find(x => x.sheet === sheet);
    if (!ms) return false;
    const rec = ms.rows.find(x => x.r === row);
    if (!rec) return false;
    if (colKey === "ref") {
      rec.ref = value != null ? String(value) : "";
      if (url !== undefined) rec.refUrl = url != null ? String(url) : "";
      else if (looksLikeUrl(rec.ref) && !rec.refUrl) rec.refUrl = rec.ref.trim();
    } else if (colKey === "exempt") {
      rec.exempt = !!value;
    } else if (colKey === "exemptReason") {
      rec.exemptReason = value != null ? String(value) : "";
    } else {
      rec[colKey] = value != null ? String(value) : "";
    }
    recordOp("auditUpdateCell", [fp, sheet, row, colKey, value, url]);
    return true;
  };

  /* Set/clear exemption for a school. scope: "month" (current) or "all" (every month sheet).
   * reason: free text / dropdown value. Journals auditUpdateCell ops for the
   * EXEMPT (K) and EXEMPT REASON (L) columns; headers are stamped on save. */
  QBR.auditSetExempt = function (fp, school, on, reason, scope) {
    const a = auditModel();
    if (!a || a.fp !== fp) return false;
    const months = scope === "all" ? Object.keys(a.months) : [uiState().month];
    months.forEach(m => {
      const ms = a.months[m];
      if (!ms) return;
      const rec = ms.rows.find(x => String(x.school || "").toLowerCase() === String(school || "").toLowerCase());
      if (!rec) return;
      QBR.auditUpdateCell(fp, ms.sheet, rec.r, "exempt", !!on);
      QBR.auditUpdateCell(fp, ms.sheet, rec.r, "exemptReason", on ? (reason || "") : "");
    });
    renderAudit();
    return true;
  };

  // Append a school row at the end of the month sheet.
  // data = {school, org, risky, health, ref, refUrl}
  QBR.auditAddRow = function (fp, sheet, data) {
    const a = auditModel();
    if (!a || a.fp !== fp) return false;
    const ms = Object.values(a.months).find(x => x.sheet === sheet);
    if (!ms) return false;
    const lastR = ms.rows.reduce((mx, x) => Math.max(mx, x.r), 1);
    const rec = {
      r: lastR + 1,
      school: String((data && data.school) || ""),
      org: String((data && data.org) || ""),
      risky: String((data && data.risky) || ""),
      health: String((data && data.health) || ""),
      ref: String((data && data.ref) || ""),
      refUrl: String((data && data.refUrl) || ""),
    };
    if (!rec.school.trim()) return false;
    ms.rows.push(rec);
    recordOp("auditAddRow", [fp, sheet, data]);
    return true;
  };

  // Paste CSV/TSV values (from Excel copy) into a school row.
  // values: array of up to 5 values mapped to [org, risky, health, ref] —
  // school identity comes from the row being edited. Also accepts the
  // 7-column storage layout; extra values are ignored.
  QBR.auditPasteRow = function (fp, sheet, row, values) {
    const a = auditModel();
    if (!a || a.fp !== fp) return false;
    const ms = Object.values(a.months).find(x => x.sheet === sheet);
    if (!ms) return false;
    const rec = ms.rows.find(x => x.r === row);
    if (!rec || !values || !values.length) return false;
    // Risky layout: pasted values map to org, risky, health, ref(, refUrl)
    const keys = ["org", "risky", "health", "ref"];
    let changed = false;
    keys.forEach((k, i) => {
      if (i < values.length && values[i] != null && String(values[i]).trim() !== "") {
        rec[k] = String(values[i]).trim();
        if (k === "ref" && looksLikeUrl(rec[k])) rec.refUrl = rec[k];
        changed = true;
      }
    });
    if (changed) recordOp("auditPasteRow", [fp, sheet, row, values]);
    return changed;
  };

  /* Parse clipboard text from an Excel copy: tab- or comma-separated, with or
   * without a header row. Returns {values} or {error}. */
  QBR.auditParsePaste = function (text) {
    if (text == null || String(text).trim() === "") return { error: "empty" };
    const lines = String(text).split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return { error: "empty" };
    const splitLine = l => {
      // Prefer tabs (Excel copy); fall back to commas (raw CSV).
      let parts = l.split("\t");
      if (parts.length < 2) parts = l.split(",");
      return parts.map(p => p.trim().replace(/^"|"$/g, "").trim());
    };
    let vals = splitLine(lines[0]);
    // Drop a header row if the first line looks like headers, not data.
    if (vals.length && /^(on[e]?drive|exchange|share[\s_-]?point|current storage|usage|school|org)/i.test(vals[0])) {
      if (lines.length < 2) return { error: "header-only" };
      vals = splitLine(lines[1]);
    }
    vals = vals.filter(v => v !== "");
    if (!vals.length) return { error: "empty" };
    return { values: vals };
  };

  /* ---------- UI state + rendering ---------------------------------------- */
  QBR._auditUI = QBR._auditUI || { month: null, editing: null, adding: false, pasteFor: null };

  function uiState() {
    const u = QBR._auditUI;
    if (!u.month) {
      const d = new Date();
      u.month = AUDIT_MONTHS[d.getMonth()];
    }
    return u;
  }

  function $(id) { return document.getElementById(id); }

  function escOpt(v, sel) {
    const e = String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    return `<option value="${e}"${String(v) === String(sel) ? " selected" : ""}>${e || "—"}</option>`;
  }

  function refCellHtml(rec) {
    const e = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    if (rec.refUrl) {
      const label = rec.ref || rec.refUrl;
      // Safety: only https/mailto become clickable; anything else renders as text.
      if (/^(https:\/\/|mailto:)/i.test(rec.refUrl)) {
        return `<a href="${e(rec.refUrl)}" target="_blank" rel="noopener">${e(label)}</a>`;
      }
      return e(label);
    }
    return rec.ref ? e(rec.ref) : `<span class="text-muted">—</span>`;
  }

  function monthPills(u, months) {
    return `<div class="audit-months" role="tablist" aria-label="Month">` +
      AUDIT_MONTHS.map(m => {
        const has = !!months[m];
        const active = u.month === m;
        return `<button type="button" role="tab" aria-selected="${active}" class="audit-month${active ? " active" : ""}${has ? "" : " empty"}" data-audit-month="${m}" title="${has ? "Open " + m : m + " (no sheet loaded)"}">${m.slice(0, 3)}</button>`;
      }).join("") + `</div>`;
  }

  // Month-over-month delta badge for a school's risky count.
  // Returns "" when there's no previous-month value to compare.
  function auditMomBadge(a, month, rec) {
    try {
      const prev = auditPrevMonth(month);
      if (!prev) return "";
      const prow = auditFindRow(a, prev, rec.school);
      if (!prow) return "";
      const cur = parseFloat(rec.risky), oldV = parseFloat(prow.risky);
      if (isNaN(cur) || isNaN(oldV)) return "";
      const d = cur - oldV;
      if (d === 0) return "";
      const cls = d > 0 ? "text-danger" : "text-success";
      const arrow = d > 0 ? "\u2191" : "\u2193";
      return ` <span class="${cls} small" title="vs ${prev}: ${oldV}">${arrow}${Math.abs(d)}</span>`;
    } catch (e) { return ""; }
  }

  function auditFilterHtml(u) {
    const f = u.filter || "all", q = u.search || "";
    const pill = (key, label) =>
      `<button type="button" class="btn btn-sm ${f === key ? "btn-primary" : "btn-outline-secondary"}" data-audit-filter="${key}">${label}</button>`;
    return `<div class="audit-filterbar">` +
      `<input type="search" id="audit-search" class="form-control form-control-sm" style="max-width:260px" placeholder="Search school…" value="${escHtml(q)}">` +
      `<div class="btn-group btn-group-sm" role="group">` +
      pill("all", "All") + pill("audited", "Audited") + pill("notaudited", "Not audited") + pill("incomplete", "Incomplete") + pill("exempted", "Exempted") +
      `</div></div>`;
  }

  QBR._auditRowPasses = auditRowPasses;
  // Does a row pass the active search + status filter?
  // "Audited" = fully audited (risky + health + storage + usage).
  function auditRowPasses(rec, u, a, month) {
    const q = (u.search || "").trim().toLowerCase();
    if (q && String(rec.school || "").toLowerCase().indexOf(q) < 0) return false;
    const f = u.filter || "all";
    if (f === "all") return true;
    const hasRisky = rec.risky != null && String(rec.risky).trim() !== "";
    const hasHealth = rec.health != null && String(rec.health).trim() !== "";
    if (f === "exempted") return !!rec.exempt;
    if (rec.exempt) return false; // exempted schools never count as not-audited/incomplete
    if (f === "audited") return a ? auditIsFullyAudited(a, month, rec.school) : hasRisky;
    if (f === "notaudited") return a ? !auditIsFullyAudited(a, month, rec.school) : !hasRisky;
    if (f === "incomplete") return hasRisky !== hasHealth;
    return true;
  }

  // Big visual progress bar, rendered into #audit-progress-top (beside the panel heading).
  // Traffic-light colors: red 0-33%, amber 34-66%, blue 67-99%, green 100%.
  function auditProgressTopHtml(a, month) {
    try {
      const ms = a.months[month];
      if (!ms || !ms.rows.length) return "";
      const eligible = ms.rows.filter(rec => !rec.exempt);
      const exemptN = ms.rows.length - eligible.length;
      const total = eligible.length;
      const doneN = eligible.filter(rec => auditIsFullyAudited(a, month, rec.school)).length;
      const pct = total ? Math.round(doneN / total * 100) : 0;
      const grad = pct >= 100 ? "linear-gradient(90deg,#198754,#20c997)"
        : pct > 66 ? "linear-gradient(90deg,#0d6efd,#20c997)"
        : pct > 33 ? "linear-gradient(90deg,#fd7e14,#ffc107)"
        : "linear-gradient(90deg,#dc3545,#e35d6a)";
      const exPill = exemptN ? `<span class="audit-ex-pill">${exemptN} exempted</span>` : "";
      const remain = total - doneN;
      const sub = pct >= 100 ? "All schools audited — done"
        : `${pct}% complete &middot; ${remain} remaining`;
      return `<div class="audit-progress-top-in">` +
        `<span class="audit-progress-month">${escHtml(month)}</span>` +
        `<span class="audit-progress-count">${doneN}<span class="audit-progress-total">/${total}</span></span>` +
        `<span class="audit-progress-lbl">schools audited</span>${exPill}` +
        `<div class="audit-progress-bar"><div class="audit-progress-fill" style="width:${pct}%;background:${grad}"></div></div>` +
        `<div class="audit-progress-sub">${sub}</div></div>`;
    } catch (e) { return ""; }
  }

  function renderAudit() {
    const host = $("audit-body");
    if (!host) return;
    const a = auditModel();
    const u = uiState();

    if (!a) {
      host.innerHTML =
        `<div class="audit-empty">
           <h4>Audit — Risky Sign-ins</h4>
           <p class="text-muted">Upload <code>RISKY_USERS_AND_DOMAIN.xlsx</code> to edit your monthly audit sheets here.</p>
           <p class="text-muted small">Edits are journaled and saved back into the workbook with formatting, formulas and layout preserved. Tick <b>Direct save</b> next to the file in the loaded-files list for direct save + background change detection.</p>
         </div>`;
      try { const pt0 = document.getElementById("audit-progress-top"); if (pt0) pt0.innerHTML = ""; } catch (e) {}
      return;
    }

    if (QBR._auditWizard) { renderWizard(host); return; }

    const ms = monthState(u.month);
    const dd = a.dropdowns || { org: [], health: [] };
    const saveBtn = (typeof QBR.saveButtonHtml === "function") ? QBR.saveButtonHtml("audit", "audit-save") : "";
    const watchBadge = `<span id="audit-watch" class="audit-watch" title="Background change detection (linked files)"></span>`;

    let html =
      `<div class="audit-head">
         <div>
           <h4 class="mb-1">Audit — Risky Sign-ins <span class="badge bg-secondary">v${QBR.AUDIT_VERSION}</span></h4>
           <div class="small text-muted">${escHtml(a.fileName || "audit workbook")} · month sheets: ${Object.keys(a.months).length}/12</div>
         </div>
         <div class="audit-actions">${watchBadge}<button type="button" class="btn btn-sm btn-outline-primary" id="audit-wizard-start" title="Step-by-step: school → risky users → domain health → storage → usage">🧭 Start guided audit</button>${saveBtn}</div>
       </div>` +
      monthPills(u, a.months) + auditFilterHtml(u);
    try {
      const pt = document.getElementById("audit-progress-top");
      if (pt) pt.innerHTML = auditProgressTopHtml(a, u.month);
    } catch (e) {}

    if (!ms) {
      html += `<p class="text-muted mt-3">No <b>${u.month}</b> sheet in this workbook.</p>`;
      host.innerHTML = html;
      wireAuditStatic(host);
      return;
    }

    const isEditing = r => u.editing === r;
    const rows = ms.rows.filter(rec => auditRowPasses(rec, u, a, u.month)).map(rec => {
      if (isEditing(rec.r)) return editRowHtml(rec, dd);
      // Incomplete rows (risky XOR health) get a subtle warning tint.
      const hasRisky = rec.risky != null && String(rec.risky).trim() !== "";
      const hasHealth = rec.health != null && String(rec.health).trim() !== "";
      const incomplete = hasRisky !== hasHealth;
      const e = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      return `<tr data-audit-row="${rec.r}"${rec.exempt ? ' class="audit-exempted"' : (incomplete ? ' class="audit-incomplete"' : "")} title="${incomplete ? "Incomplete: risky count / health mismatch" : ""}">
        <td>${e(rec.school)}${rec.exempt ? " <span class=\"badge bg-secondary\">Exempt" + (rec.exemptReason ? ": " + e(rec.exemptReason) : "") + "</span>" : ""}</td>
        <td>${e(rec.org) || '<span class="text-muted">—</span>'}</td>
        <td>${e(rec.risky) || '<span class="text-muted">—</span>'}${auditMomBadge(a, u.month, rec)}</td>
        <td>${e(rec.health) || '<span class="text-muted">—</span>'}</td>
        <td>${refCellHtml(rec)}</td>
        <td class="audit-rowops"><button type="button" class="btn btn-sm btn-outline-secondary" data-audit-edit="${rec.r}">Edit</button> ${rec.exempt ? `<button type="button" class="btn btn-sm btn-outline-warning" data-audit-unexempt="${rec.r}" title="Remove exemption">Unexempt</button>` : `<button type="button" class="btn btn-sm btn-outline-secondary" data-audit-exempt="${rec.r}" title="Mark exempt from auditing">Exempt</button>`}</td>
      </tr>`;
    }).join("");

    html +=
      `<div class="table-responsive mt-2"><table class="table table-sm table-hover audit-table">
         <thead><tr><th>School</th><th>Organization</th><th>Total Risky Users</th><th>Domain Health</th><th>References</th><th></th></tr></thead>
         <tbody>${rows || `<tr><td colspan="6" class="text-muted">No schools recorded for ${u.month} yet.</td></tr>`}</tbody>
       </table></div>
       <button type="button" class="btn btn-sm btn-outline-primary" id="audit-add">＋ Add school</button>
       <div id="audit-addform"></div>
       <div id="audit-paste"></div>`;

    host.innerHTML = html;
    wireAudit(host, a, ms, dd);
    if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
    auditWatchPaint();
  }

  function escHtml(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function editRowHtml(rec, dd) {
    const e = escHtml;
    const orgOpts = [`<option value="">—</option>`].concat((dd.org || []).map(v => escOpt(v, rec.org))).join("");
    const healthOpts = [`<option value="">—</option>`].concat((dd.health || []).map(v => escOpt(v, rec.health))).join("");
    return `<tr data-audit-row="${rec.r}" class="audit-editing">
      <td><input class="form-control form-control-sm" id="ae-school" value="${e(rec.school)}"></td>
      <td><select class="form-select form-select-sm" id="ae-org">${orgOpts}</select></td>
      <td><input class="form-control form-control-sm" id="ae-risky" value="${e(rec.risky)}" placeholder="0"></td>
      <td><select class="form-select form-select-sm" id="ae-health">${healthOpts}</select></td>
      <td>
        <input class="form-control form-control-sm mb-1" id="ae-reftext" value="${e(rec.ref)}" placeholder="Link text">
        <input class="form-control form-control-sm" id="ae-refurl" value="${e(rec.refUrl)}" placeholder="https://… sharepoint link">
      </td>
      <td class="audit-rowops">
        <button type="button" class="btn btn-sm btn-primary" data-audit-save="${rec.r}">Save</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" data-audit-cancel>Cancel</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" data-audit-paste="${rec.r}" title="Paste Excel copy (tab-separated values)">⧉ Paste</button>
      </td>
    </tr>`;
  }

  function wireAuditStatic(host) {
    host.querySelectorAll("[data-audit-month]").forEach(b =>
      b.addEventListener("click", () => { uiState().month = b.dataset.auditMonth; uiState().editing = null; renderAudit(); }));
    const sv = $("audit-save");
    if (sv) sv.addEventListener("click", auditSave);
    const wz = $("audit-wizard-start");
    if (wz) wz.addEventListener("click", () => { if (typeof QBR.auditWizardStart === "function") QBR.auditWizardStart(); });
    const si = $("audit-search");
    if (si) si.addEventListener("input", () => { uiState().search = si.value; renderAudit(); keepSearchFocus(); });
    host.querySelectorAll("[data-audit-filter]").forEach(b =>
      b.addEventListener("click", () => { uiState().filter = b.dataset.auditFilter; renderAudit(); }));
  }

  // Keep the caret in the search box across re-renders while typing.
  function keepSearchFocus() {
    try {
      const si = $("audit-search");
      if (si && document.activeElement !== si) {
        const v = si.value;
        si.focus();
        si.setSelectionRange(v.length, v.length);
      }
    } catch (e) {}
  }

  var EXEMPT_REASONS = ["No tenant access", "No GDAP", "No admin access", "Other"];

  function showExemptDialog(a, ms, row) {
    const rec = ms.rows.find(x => x.r === row);
    if (!rec) return;
    const host = $("audit-body");
    if (!host) return;
    const opts = EXEMPT_REASONS.map(r => `<option value="${escHtml(r)}">${escHtml(r)}</option>`).join("");
    const ov = document.createElement("div");
    ov.className = "audit-modal-ov";
    ov.innerHTML =
      `<div class="audit-modal" role="dialog" aria-label="Mark school exempt">` +
      `<h5 class="mb-2">Exempt <b>${escHtml(rec.school)}</b> from auditing</h5>` +
      `<label class="form-label small mb-1">Reason</label>` +
      `<select id="ex-reason" class="form-select form-select-sm mb-2">${opts}</select>` +
      `<input id="ex-reason-other" class="form-control form-control-sm mb-2 d-none" placeholder="Specify reason…">` +
      `<label class="form-label small mb-1">Apply to</label>` +
      `<select id="ex-scope" class="form-select form-select-sm mb-3">` +
      `<option value="month">This month (${escHtml(uiState().month)}) only</option>` +
      `<option value="all" selected>All 12 months</option></select>` +
      `<div class="d-flex justify-content-end gap-2">` +
      `<button type="button" class="btn btn-sm btn-outline-secondary" id="ex-cancel">Cancel</button>` +
      `<button type="button" class="btn btn-sm btn-warning" id="ex-ok">Mark exempt</button></div></div>`;
    host.appendChild(ov);
    const rs = ov.querySelector("#ex-reason"), ro = ov.querySelector("#ex-reason-other");
    rs.addEventListener("change", () => ro.classList.toggle("d-none", rs.value !== "Other"));
    ov.querySelector("#ex-cancel").addEventListener("click", () => ov.remove());
    ov.addEventListener("click", e => { if (e.target === ov) ov.remove(); });
    ov.querySelector("#ex-ok").addEventListener("click", () => {
      let reason = rs.value;
      if (reason === "Other") reason = ro.value.trim() || "Other";
      QBR.auditSetExempt(a.fp, rec.school, true, reason, ov.querySelector("#ex-scope").value);
      ov.remove();
    });
  }

  function showUnexemptDialog(a, ms, row) {
    const rec = ms.rows.find(x => x.r === row);
    if (!rec) return;
    if (!window.confirm(`Remove exemption for "${rec.school}"?\n\nThis month only, or all 12 months?\n\nOK = all months · Cancel = this month only`)) {
      QBR.auditSetExempt(a.fp, rec.school, false, "", "month");
    } else {
      QBR.auditSetExempt(a.fp, rec.school, false, "", "all");
    }
  }

  function wireAudit(host, a, ms, dd) {
    wireAuditStatic(host);
    host.querySelectorAll("[data-audit-edit]").forEach(b =>
      b.addEventListener("click", () => { uiState().editing = Number(b.dataset.auditEdit); renderAudit(); }));
    host.querySelectorAll("[data-audit-cancel]").forEach(b =>
      b.addEventListener("click", () => { uiState().editing = null; uiState().adding = false; renderAudit(); }));
    host.querySelectorAll("[data-audit-save]").forEach(b =>
      b.addEventListener("click", () => saveEditRow(a, ms, Number(b.dataset.auditSave))));
    host.querySelectorAll("[data-audit-paste]").forEach(b =>
      b.addEventListener("click", () => showPasteBox(a, ms, Number(b.dataset.auditPaste))));
    host.querySelectorAll("[data-audit-exempt]").forEach(b =>
      b.addEventListener("click", () => showExemptDialog(a, ms, Number(b.dataset.auditExempt))));
    host.querySelectorAll("[data-audit-unexempt]").forEach(b =>
      b.addEventListener("click", () => showUnexemptDialog(a, ms, Number(b.dataset.auditUnexempt))));
    const add = $("audit-add");
    if (add) add.addEventListener("click", () => showAddForm(a, ms, dd));
  }

  function saveEditRow(a, ms, row) {
    const g = id => { const el = $(id); return el ? el.value : ""; };
    const rec = ms.rows.find(x => x.r === row);
    const updates = [
      ["school", g("ae-school")],
      ["org", g("ae-org")],
      ["risky", g("ae-risky")],
      ["health", g("ae-health")],
    ];
    updates.forEach(([k, v]) => {
      if (!rec || String(rec[k] || "") !== String(v || "")) QBR.auditUpdateCell(a.fp, ms.sheet, row, k, v);
    });
    const rt = g("ae-reftext"), ru = g("ae-refurl");
    if (!rec || String(rec.ref || "") !== String(rt || "") || String(rec.refUrl || "") !== String(ru || ""))
      QBR.auditUpdateCell(a.fp, ms.sheet, row, "ref", rt, ru);
    uiState().editing = null;
    renderAudit();
  }

  function showAddForm(a, ms, dd) {
    const host = $("audit-addform");
    if (!host) return;
    const e = escHtml;
    const orgOpts = [`<option value="">—</option>`].concat((dd.org || []).map(v => escOpt(v, "")).join(""));
    const healthOpts = [`<option value="">—</option>`].concat((dd.health || []).map(v => escOpt(v, "")).join(""));
    host.innerHTML =
      `<div class="card card-body mt-2"><h6>Add school — ${e(ms.sheet.trim())}</h6>
       <div class="row g-2">
         <div class="col-md-4"><input class="form-control form-control-sm" id="aa-school" placeholder="School name *"></div>
         <div class="col-md-2"><select class="form-select form-select-sm" id="aa-org">${orgOpts}</select></div>
         <div class="col-md-2"><input class="form-control form-control-sm" id="aa-risky" placeholder="Risky users"></div>
         <div class="col-md-2"><select class="form-select form-select-sm" id="aa-health">${healthOpts}</select></div>
         <div class="col-md-2"><input class="form-control form-control-sm" id="aa-refurl" placeholder="References link"></div>
       </div>
       <div class="mt-2">
         <button type="button" class="btn btn-sm btn-primary" id="aa-save">Add</button>
         <button type="button" class="btn btn-sm btn-outline-secondary" id="aa-cancel">Cancel</button>
       </div></div>`;
    $("aa-cancel").addEventListener("click", () => { host.innerHTML = ""; });
    $("aa-save").addEventListener("click", () => {
      const g = id => { const el = $(id); return el ? el.value.trim() : ""; };
      const school = g("aa-school");
      if (!school) { alert("School name is required."); return; }
      QBR.auditAddRow(a.fp, ms.sheet, {
        school: school, org: g("aa-org"), risky: g("aa-risky"),
        health: g("aa-health"), ref: g("aa-refurl"), refUrl: g("aa-refurl"),
      });
      host.innerHTML = "";
      renderAudit();
    });
    $("aa-school").focus();
  }

  function showPasteBox(a, ms, row) {
    const host = $("audit-paste");
    if (!host) return;
    host.innerHTML =
      `<div class="card card-body mt-2"><h6>Paste Excel copy</h6>
       <p class="small text-muted mb-2">Copy the row in Excel, then paste below. Tab- or comma-separated, headers optional — values map to Organization → Total Risky Users → Domain Health → References.</p>
       <textarea class="form-control form-control-sm" id="ap-text" rows="3" placeholder="Paste here…"></textarea>
       <div id="ap-preview" class="mt-2"></div>
       <div class="mt-2">
         <button type="button" class="btn btn-sm btn-primary" id="ap-apply">Apply to row</button>
         <button type="button" class="btn btn-sm btn-outline-secondary" id="ap-cancel">Cancel</button>
       </div></div>`;
    const ta = $("ap-text"), prev = $("ap-preview");
    ta.addEventListener("input", () => {
      const r = QBR.auditParsePaste(ta.value);
      prev.innerHTML = r.error
        ? `<span class="text-muted small">Waiting for pasted values…</span>`
        : `<span class="small">Will set: <code>${r.values.map(v => escHtml(v)).join("</code> · <code>")}</code></span>`;
    });
    $("ap-cancel").addEventListener("click", () => { host.innerHTML = ""; });
    $("ap-apply").addEventListener("click", () => {
      const r = QBR.auditParsePaste(ta.value);
      if (r.error) { alert("Nothing to paste — copy the row in Excel first."); return; }
      if (QBR.auditPasteRow(a.fp, ms.sheet, row, r.values)) {
        uiState().editing = null;
        host.innerHTML = "";
        renderAudit();
      }
    });
    ta.focus();
  }

  async function auditSave() {
    if (typeof QBR.persistNote === "function") QBR.persistNote("audit", "Saving…", 15000);
    let r;
    try { r = await QBR.auditFamilySave("audit"); }
    catch (e) { r = { mode: "error" }; }
    if (typeof QBR.auditSaveDone === "function") QBR.auditSaveDone(r);
    else if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
  }

  QBR.auditSaveDone = function (r) {
    if (!r) return;
    const note = (typeof QBR.persistNote === "function") ? QBR.persistNote : null;
    if (r.mode === "file") {
      const extra = r.applied ? ` (${r.applied} change${r.applied === 1 ? "" : "s"}${r.skipped ? `, ${r.skipped} skipped` : ""})` : "";
      if (note) note("audit", `Saved ✓ ${r.name}${extra}${r.engine === "format-safe" ? " · formatting kept" : ""}${r.merged ? " · merged with Excel changes" : ""}`, 5000);
    }
    else if (r.mode === "no-changes") { if (note) note("audit", "No changes to save" + (r.kind ? " for " + r.kind : ""), 3000); }
    else if (r.mode === "download" || r.mode === "download-fallback") { if (note) note("audit", `Downloaded ${r.filename || ""} — tick Direct save for the file to save into it`, 5000); }
    else if (r.mode === "download-changed") { if (note) note("audit", `${r.name || "Linked file"} changed in Excel — downloaded a separate copy. Re-link to save directly.`, 12000); }
    else if (r.mode === "denied") { if (note) note("audit", "Permission denied — file not saved", 5000); }
    else if (r.mode === "no-data") alert("Load an audit workbook first.");
    else if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
    if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
  };

  // Exposed for renderAll() in app.js (this file is an IIFE; renderAudit would
  // otherwise be invisible there and the panel would stay blank).
  QBR.renderAudit = renderAudit;

  /* ---------- external-change watcher --------------------------------------
   * For linked audit files (Chrome/Edge): polls the live file's metadata every
   * 30s. When OneDrive syncs a newer version in the background, a badge
   * appears BEFORE save — "File changed externally — Review" — instead of a
   * surprise at save time. Review shows per-row conflicts: rows changed on
   * disk vs rows with pending journaled edits.
   * ------------------------------------------------------------------------ */
  QBR._auditWatch = QBR._auditWatch || { timer: null, changedFp: null, freshRows: null };

  QBR.auditWatchStart = function () {
    const W = QBR._auditWatch;
    if (W.timer || typeof QBR.fsGetLink !== "function") return;
    W.timer = setInterval(() => { QBR.auditWatchTick().catch(() => {}); }, 30000);
  };

  QBR.auditWatchStop = function () {
    const W = QBR._auditWatch;
    if (W.timer) { clearInterval(W.timer); W.timer = null; }
  };

  QBR.auditWatchTick = async function () {
    const W = QBR._auditWatch;
    if (!QBR.fsSupported || !QBR.fsSupported()) return;
    const links = (QBR._fsLinks || []).filter(l => (l.kinds || []).indexOf("audit") >= 0 && l.handle);
    for (const link of links) {
      let cur;
      try { cur = await link.handle.getFile(); }
      catch (e) { continue; } // permission lost or file moved — stay quiet
      if (cur.size !== link.size || cur.lastModified !== link.lastModified) {
        // 2026-10-06: only real content changes count (OneDrive sync touches timestamps)
        try {
          if (link.hash && typeof QBR.fsHashOf === "function" && QBR.fsHashOf(new Uint8Array(await cur.arrayBuffer())) === link.hash) {
            if (typeof QBR.fsRebaseLink === "function") await QBR.fsRebaseLink(link, cur);
            continue;
          }
        } catch (e) {}
        if (W.changedFp !== link.fp) {
          W.changedFp = link.fp;
          try { W.freshBuf = await cur.arrayBuffer(); } catch (e) { W.freshBuf = null; }
          auditWatchPaint();
          if (typeof QBR.persistNote === "function") QBR.persistNote("audit", "⚠ File changed in Excel — Save will merge your edits with it (↻ reloads it)", 15000);
        }
      }
    }
  };

  function auditWatchPaint() {
    const el = document.getElementById("audit-watch");
    if (!el) return;
    const W = QBR._auditWatch;
    if (W.changedFp) {
      el.innerHTML = `<button type="button" class="btn btn-sm btn-warning" id="audit-review">⚠ File changed externally — Review</button>`;
      const b = document.getElementById("audit-review");
      if (b) b.addEventListener("click", auditShowConflicts);
    } else {
      el.innerHTML = "";
    }
  }

  // Re-parse the fresh file bytes and diff against the live model + journal.
  function auditShowConflicts() {
    const W = QBR._auditWatch;
    const a = auditModel();
    const host = document.getElementById("audit-body");
    if (!a || !host || !W.freshBuf) return;
    let fresh = null;
    try {
      const wb = XLSX.read(W.freshBuf, { type: "array" });
      const sheets = {}, raw = {};
      wb.SheetNames.forEach(n => {
        raw[n] = wb.Sheets[n];
        sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false });
      });
      fresh = QBR.auditParseSheets(sheets, raw, { fp: "fresh" });
    } catch (e) { alert("Could not re-read the updated file: " + (e && e.message || e)); return; }
    if (!fresh) { alert("The updated file no longer looks like an audit workbook."); return; }

    // Journaled (pending) edits, keyed by sheet|row|colKey.
    let pending = {};
    try {
      const store = JSON.parse(localStorage.getItem("qbr-inv-journal-v1") || "{}");
      const entry = store[a.fp];
      (entry && entry.ops || []).forEach(o => {
        if (o && o.kind === "audit" && o.op === "auditUpdateCell" && o.args) {
          pending[o.args[1] + "|" + o.args[2] + "|" + o.args[3]] = o.args[4];
        }
      });
    } catch (e) {}

    const u = uiState();
    const ms = a.months[u.month], fm = fresh.months[u.month];
    const e2 = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    let rows = "";
    if (ms && fm) {
      const fmap = {};
      fm.rows.forEach(r => { fmap[r.school] = r; });
      ms.rows.forEach(rec => {
        const fr = fmap[rec.school];
        if (!fr) return;
        const diffs = [];
        ["org", "risky", "health", "ref"].forEach(k => {
          const liveV = String(rec[k] || ""), freshV = String(fr[k] || "");
          const pk = ms.sheet + "|" + rec.r + "|" + k;
          if (liveV !== freshV) {
            diffs.push({ k: k, mine: pending[pk] !== undefined ? pending[pk] : liveV, theirs: freshV, pending: pending[pk] !== undefined });
          }
        });
        if (diffs.length) {
          rows += `<tr><td>${escHtml(rec.school)}</td><td>` +
            diffs.map(d =>
              `<div class="mb-1"><b>${e2(d.k)}</b>:<br>` +
              `<span class="badge ${d.pending ? "bg-primary" : "bg-secondary"}">yours: ${e2(d.mine) || "—"}</span> ` +
              `<span class="badge bg-warning text-dark">file: ${e2(d.theirs) || "—"}</span></div>`
            ).join("") + `</td></tr>`;
        }
      });
    }
    const body =
      `<div class="card card-body mt-3">
         <h5>⚠ External changes detected</h5>
         <p class="small text-muted">The linked file changed outside the dashboard (e.g. OneDrive sync). Rows below differ between your loaded copy and the file on disk for <b>${escHtml(u.month)}</b>.</p>
         ${rows ? `<table class="table table-sm"><thead><tr><th>School</th><th>Differences</th></tr></thead><tbody>${rows}</tbody></table>` : `<p class="text-muted">No row-level differences in ${escHtml(u.month)} — the change may be in another month sheet.</p>`}
         <div class="mt-2">
           <button type="button" class="btn btn-sm btn-primary" id="audit-reload">Reload fresh file</button>
           <button type="button" class="btn btn-sm btn-outline-secondary" id="audit-dismiss">Dismiss (keep editing)</button>
         </div>
         <p class="small text-muted mt-2 mb-0">Reload: re-parse the file from disk. Your journaled edits stay recorded and will re-apply to matching rows on the next save review.</p>
       </div>`;
    const anchor = document.getElementById("audit-paste");
    if (anchor) anchor.innerHTML = body;
    else host.insertAdjacentHTML("beforeend", body);
    document.getElementById("audit-reload").addEventListener("click", () => {
      W.changedFp = null; W.freshBuf = null;
      alert("Untick and re-tick Direct save for the file to load the fresh copy, then re-apply any edits from the conflict list above.");
      auditWatchPaint();
    });
    document.getElementById("audit-dismiss").addEventListener("click", () => {
      W.changedFp = null; auditWatchPaint(); renderAudit();
    });
  }

  /* ---------- export workbook (unlinked fallback) ------------------------- */
  QBR.auditExportWorkbook = function () {
    const a = auditModel();
    if (!a || !QBR._origWb || !QBR._origWb[a.fp]) { alert("Load an audit workbook first."); return; }
    // Patch a copy so the live model keeps working; download the result.
    const wb = QBR._origWb[a.fp];
    const res = (typeof QBR.patchWorkbookFromJournal === "function")
      ? QBR.patchWorkbookFromJournal("audit", wb, a.fp) : { ok: false };
    if (!res.ok) { alert("Could not prepare the workbook: " + (res.error || "unknown")); return; }
    try {
      if (typeof QBR.sanitizeHyperlinkTargets === "function") QBR.sanitizeHyperlinkTargets(wb);
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = (a.fileName || "audit-workbook").replace(/\.xlsx?$/i, "") + "-edited.xlsx";
      document.body.appendChild(link); link.click();
      setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 500);
      if (typeof QBR.journalClearFp === "function") QBR.journalClearFp(a.fp);
      if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
      return { mode: "download", filename: link.download, kind: "audit" };
    } catch (e) { alert("Export failed: " + (e && e.message || e)); }
    return { mode: "error", kind: "audit" };
  };

  /* ---------- storage & usage exporters (unlinked fallback) ------------------ */
  function exportKindWorkbook(kind, fileName) {
    const mdl = kind === "storage" ? QBR._storage : kind === "usage" ? QBR._usage : auditModel();
    if (!mdl || !QBR._origWb || !QBR._origWb[mdl.fp]) { alert("Load the workbook first."); return; }
    const wb = QBR._origWb[mdl.fp];
    const res = (typeof QBR.patchWorkbookFromJournal === "function")
      ? QBR.patchWorkbookFromJournal(kind, wb, mdl.fp) : { ok: false };
    if (!res.ok) { alert("Could not prepare the workbook: " + (res.error || "unknown")); return; }
    try {
      if (typeof QBR.sanitizeHyperlinkTargets === "function") QBR.sanitizeHyperlinkTargets(wb);
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = (fileName || "workbook").replace(/\.xlsx?$/i, "") + "-edited.xlsx";
      document.body.appendChild(link); link.click();
      setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 500);
      // The journaled changes are now in the downloaded copy — clear them.
      if (typeof QBR.journalClearFp === "function") QBR.journalClearFp(mdl.fp);
      if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
      return { mode: "download", filename: link.download, kind: kind };
    } catch (e) { alert("Export failed: " + (e && e.message || e)); }
    return { mode: "error", kind: kind };
  }
  QBR.storageExportWorkbook = function () {
    const m = QBR._storage; if (!m) { alert("Load a storage workbook first."); return; }
    exportKindWorkbook("storage", m.fileName);
  };
  QBR.usageExportWorkbook = function () {
    const m = QBR._usage; if (!m) { alert("Load a usage workbook first."); return; }
    exportKindWorkbook("usage", m.fileName);
  };
  // Route a save click for an audit-family kind: linked → direct patch-in-place
  // save; unlinked → download an edited copy via the kind-specific exporter.
  QBR.auditFamilySave = async function (kind) {
    const link = (typeof QBR.fsGetLink === "function") ? QBR.fsGetLink(kind) : null;
    if (link) return QBR.fsSaveKind(kind);
    if (kind === "storage") return QBR.storageExportWorkbook();
    if (kind === "usage") return QBR.usageExportWorkbook();
    return QBR.auditExportWorkbook();
  };

  /* ================= storage & usage workbooks =================
   * The guided audit wizard writes into two more source workbooks:
   *  - STORAGE_REPORT.xlsx  (quarter sheets Q1-Q4, A1 = "School (Rakso Managed)")
   *  - USAGE_REPORT.xlsx    (quarter sheets Q1-Q4, A1 = "CLIENT")
   * Detection, models and journal ops mirror the risky-signins ("audit") flow.
   * Kinds registered: "storage", "usage". */
  function qNorm(n) { return String(n || "").trim().toUpperCase(); }
  function schoolNorm(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }

  function quarterRows(wb, name, schoolCol) {
    const ws = wb.Sheets[name];
    if (!ws) return null;
    // Keep blank rows so that r = index + 1 is the true sheet row number.
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!rows.length) return null;
    const recs = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const school = r[schoolCol] != null ? String(r[schoolCol]).trim() : "";
      if (!school) continue;
      recs.push({ r: i + 1, school: school, org: r[1] != null ? String(r[1]).trim() : "", _row: r });
    }
    return recs;
  }

  function parseQuarterWorkbook(wb, opts, kind, a1match, schoolCol) {
    const quarters = {};
    wb.SheetNames.forEach(name => {
      const q = qNorm(name);
      if (!/^Q[1-4]$/.test(q)) return;
      const ws = wb.Sheets[name];
      const top = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      if (!top.length) return;
      if (String(top[0][0] || "").trim().toLowerCase() !== a1match) return;
      const recs = quarterRows(wb, name, schoolCol);
      if (recs) quarters[q] = { sheet: name, rows: recs };
    });
    if (!Object.keys(quarters).length) return null;
    return { fp: (opts && opts.fp) || null, fileName: (opts && opts.fileName) || "", quarters: quarters };
  }

  function parseQuarterBuffers(buffers, kind, a1match, schoolCol, parseFn) {
    const fps = QBR._currentFps || [];
    let model = null;
    (buffers || []).forEach((buf, bi) => {
      try {
        const wb = XLSX.read(buf, { type: "array", cellStyles: true });
        const fp = fps[bi];
        let fileName = "";
        try {
          const list = (typeof persistAppFiles === "function") ? persistAppFiles() : [];
          const f = list.filter(x => QBR.fpOf(x.name, x.blob) === fp)[0];
          if (f) fileName = f.name;
        } catch (e) {}
        const m = parseFn(wb, { fp: fp || null, fileName: fileName });
        if (m) {
          if (fp) { (QBR._origWb || (QBR._origWb = {}))[fp] = wb; }
          model = m; // last matching file wins
          if (fp) {
            const set = QBR._kindByFp[fp] || (QBR._kindByFp[fp] = new Set());
            set.add(kind);
          }
        }
      } catch (e) { /* a non-matching workbook simply contributes nothing */ }
    });
    return model;
  }

  QBR.parseStorageBuffers = function (buffers) {
    const m = parseQuarterBuffers(buffers, "storage", "school (rakso managed)", 0,
      (wb, opts) => parseQuarterWorkbook(wb, opts, "storage", "school (rakso managed)", 0));
    QBR._storage = m;
    return m;
  };

  QBR.parseUsageBuffers = function (buffers) {
    const m = parseQuarterBuffers(buffers, "usage", "client", 0,
      (wb, opts) => parseQuarterWorkbook(wb, opts, "usage", "client", 0));
    QBR._usage = m;
    return m;
  };

  function storageModel() { return QBR._storage || null; }
  function usageModel() { return QBR._usage || null; }
  QBR.storageModel = storageModel;
  QBR.usageModel = usageModel;

  // Month ("SEPTEMBER") -> quarter ("Q3").
  QBR.auditQuarterOfMonth = function (month) {
    const i = AUDIT_MONTHS.indexOf(month);
    if (i < 0) return null;
    return "Q" + (Math.floor(i / 3) + 1);
  };

  // Find the sheet-row for a school inside a quarter (normalized match).
  QBR.quarterFindRow = function (model, quarter, school) {
    const q = model && model.quarters && model.quarters[quarter];
    if (!q) return null;
    const n = schoolNorm(school);
    const hit = q.rows.filter(r => schoolNorm(r.school) === n)[0] || null;
    return hit ? { r: hit.r, school: hit.school, org: hit.org, sheet: q.sheet, quarter: quarter } : null;
  };
  QBR.schoolNorm = schoolNorm;

  // Journal ops (model updates are light; the wizard tracks its own saved state).
  QBR.storagePasteRow = function (fp, sheet, row, values) {
    const s = storageModel();
    if (!s || s.fp !== fp) return false;
    QBR.journalRecord("storage", "storagePasteRow", [fp, sheet, row, values]);
    return true;
  };
  QBR.usagePasteRow = function (fp, sheet, row, values, dateStr, period) {
    const u = usageModel();
    if (!u || u.fp !== fp) return false;
    QBR.journalRecord("usage", "usagePasteRow", [fp, sheet, row, values, dateStr, period]);
    return true;
  };

  /* ================= guided audit wizard ================= */
  var WIZ_STEP_NAMES = ["School & month", "Risky users", "Domain health", "Storage report", "Usage report", "Done"];

  function wiz() { return QBR._auditWizard || null; }

  // Schools eligible for the guided wizard: not exempted in the given month.
  function wizEligibleSchools(a, month) {
    const seen = {}, out = [];
    const ms = a.months[month];
    const rows = ms ? (ms.rows || []) : [];
    // Fall back to all months if the wizard month sheet is missing.
    const pool = rows.length ? rows : Object.keys(a.months).reduce((acc, m) => acc.concat(a.months[m].rows || []), []);
    pool.forEach(r => {
      const k = schoolNorm(r.school);
      if (!r.school || seen[k]) return;
      seen[k] = 1;
      // Skip if exempted in the wizard month.
      if (rows.length) {
        const rec = rows.find(x => schoolNorm(x.school) === k);
        if (rec && rec.exempt) return;
      }
      out.push(r.school);
    });
    out.sort((x, y) => String(x).localeCompare(String(y)));
    return out;
  }

  QBR.auditWizardStart = function () {
    const a = auditModel();
    if (!a) return;
    let month = null;
    try { month = uiState().month; } catch (e) {}
    if (!month) {
      const d = new Date();
      month = AUDIT_MONTHS[d.getMonth()];
    }
    const schools = wizEligibleSchools(a, month);
    QBR._auditWizard = { school: schools[0] || "", month: month, step: 0, schools: schools,
      saved: {}, p3: { text: "", parsed: null }, p4: { text: "", parsed: null },
      dateStr: new Date().toISOString().slice(0, 10), period: "30 Days" };
    renderAudit();
  };
  QBR.auditWizardExit = function () { QBR._auditWizard = null; renderAudit(); };

  function wizRec(a, school, month) {
    const ms = a.months[month];
    if (!ms) return null;
    const n = schoolNorm(school);
    const rec = (ms.rows || []).filter(r => schoolNorm(r.school) === n)[0] || null;
    return rec ? { ms: ms, rec: rec } : null;
  }

  // ---- paste parsing ----
  function splitCells(line) {
    let parts = String(line).split("\t");
    if (parts.length < 2) parts = String(line).split(",");
    return parts.map(x => String(x).trim());
  }
  function looksLikeHeader(cells, keywords) {
    return cells.some(c => keywords.some(k => String(c).toLowerCase().indexOf(k) >= 0));
  }
  // Storage: 7 values in PowerShell CSV order -> C..I. Header-aware.
  var STORAGE_HEAD_KEYS = ["onedrive", "exchange", "sharepoint", "current storage", "usage", "used storage", "total storage"];
  QBR.auditParseStoragePaste = function (text) {
    const lines = String(text || "").split(/\r?\n/).map(x => String(x).trim()).filter(x => x);
    if (!lines.length) return { error: "Paste the 7 values copied from Excel." };
    let cells = splitCells(lines[0]);
    if (looksLikeHeader(cells, ["current storage", "used storage"]) && lines.length > 1) cells = splitCells(lines[1]);
    else if (looksLikeHeader(cells, ["current storage", "used storage"])) return { error: "Header row found but no data row below it." };
    if (cells.length < 7) return { error: "Expected 7 values, found " + cells.length + "." };
    return { values: cells.slice(0, 7) };
  };
  QBR.storagePasteHeaders = function () {
    return ["ONEDRIVE (C)", "EXCHANGE (D)", "SHARE POINT (E)", "CURRENT STORAGE (F)", "USAGE (G)", "Used Storage(GB) (H)", "Total Storage(GB) (I)"];
  };
  // Usage: 19 values -> columns D..V (4..22). Header-aware; 24-wide paste tolerated.
  QBR.auditParseUsagePaste = function (text) {
    const lines = String(text || "").split(/\r?\n/).map(x => String(x).trim()).filter(x => x);
    if (!lines.length) return { error: "Paste the 19 values copied from Excel." };
    let cells = splitCells(lines[0]);
    if (looksLikeHeader(cells, ["total number of assigned licenses", "exchange active"]) && lines.length > 1) cells = splitCells(lines[1]);
    else if (looksLikeHeader(cells, ["total number of assigned licenses", "exchange active"])) return { error: "Header row found but no data row below it." };
    if (cells.length >= 24) cells = cells.slice(3, 22); // full-width paste incl. client/org/date
    if (cells.length < 19) return { error: "Expected 19 values, found " + cells.length + "." };
    return { values: cells.slice(0, 19) };
  };
  QBR.usagePasteHeaders = function () {
    return ["Total Assigned Licenses", "Activated", "Desktop Apps", "Mobile Apps",
      "Exchange Active", "Exchange Inactive", "Total of Exchange Users",
      "OneDrive Active", "OneDrive Inactive", "Total of OneDrive Users",
      "SharePoint Active", "SharePoint Inactive", "Total of SharePoint Users",
      "Teams Active", "Teams Inactive", "Total of Teams Users",
      "Office 365 Active", "Office 365 Inactive", "Total of Office 365 Users"];
  };
  // Validate each service's Total = Active + Inactive.
  QBR.auditValidateUsage = function (values) {
    const names = ["Exchange", "OneDrive", "SharePoint", "Teams", "Office 365"];
    return names.map((n, g) => {
      const b = 4 + g * 3;
      const act = parseFloat(values[b]), ina = parseFloat(values[b + 1]), tot = parseFloat(values[b + 2]);
      const ok = isFinite(act) && isFinite(ina) && isFinite(tot) && Math.abs(act + ina - tot) < 0.01;
      return { name: n, ok: ok };
    });
  };

  // ---- wizard rendering ----
  function wizSaveButtons() { return ""; } // v1.27.6: single Save-all lives on the Done step, not the header.
  function renderWizard(host) {
    const w = wiz();
    const a = auditModel();
    const step = w.step;
    const total = 4;
    let body = "";

    const crumbs = WIZ_STEP_NAMES.map((n, i) => {
      if (i === 0 || i === 5) return "";
      const cls = (i - 1) < step ? "wiz-done" : (i - 1) === step ? "wiz-cur" : "";
      const mark = w.saved[i] ? " ✓" : "";
      return `<span class="wiz-crumb ${cls}">${i}. ${n}${mark}</span>`;
    }).join('<span class="wiz-sep">›</span>');

    if (step === 0) body = wizStepSchool(w, a);
    else if (step === 1) body = wizStepRisky(w, a);
    else if (step === 2) body = wizStepHealth(w, a);
    else if (step === 3) body = wizStepStorage(w, a);
    else if (step === 4) body = wizStepUsage(w, a);
    else body = wizStepDone(w, a);

    host.innerHTML =
      `<div class="audit-head">
         <div>
           <h4 class="mb-1">🧭 Guided audit <span class="badge bg-secondary">v${QBR.AUDIT_VERSION}</span></h4>
           <div class="small text-muted">${escHtml(w.school || "")} · ${escHtml(w.month || "")}${step >= 1 && step <= 4 ? " · Step " + step + " of " + total : ""}</div>
           <div class="wiz-crumbs">${crumbs}</div>
         </div>
         <div class="audit-actions">
           <button type="button" class="btn btn-sm btn-outline-secondary" id="wiz-exit">← Back to table</button>
           ${wizSaveButtons()}
         </div>
       </div>` + body;

    document.getElementById("wiz-exit").addEventListener("click", QBR.auditWizardExit);
    const saveAll = document.getElementById("wiz-save-all");
    if (saveAll) saveAll.addEventListener("click", () => QBR.wizardSaveAll());
    // Total pending across the three wizard files, shown on the Save-all button.
    host.querySelectorAll("[data-unsaved-all]").forEach(el => {
      let n = 0;
      ["audit", "storage", "usage"].forEach(k => {
        try { n += QBR.journalUnsavedCount(k); } catch (e) {}
      });
      el.textContent = n ? n + " unsaved" : "nothing pending";
    });
    wireWizardStep(host, w, a, step);
    if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
  }

  function wizNav(step, opts) {
    opts = opts || {};
    let h = `<div class="wiz-nav">`;
    if (step > 0) h += `<button type="button" class="btn btn-outline-secondary" data-wiz="back">← Back</button>`;
    if (opts.save) h += `<button type="button" class="btn btn-primary" data-wiz="save">${opts.saveLabel || "Save"}</button>`;
    if (step < 4) h += `<button type="button" class="btn btn-outline-primary" data-wiz="next">Next →</button>`;
    else if (step === 4) h += `<button type="button" class="btn btn-success" data-wiz="finish">Finish →</button>`;
    h += `</div><div class="wiz-msg small" id="wiz-msg"></div>`;
    return h;
  }

  function wizStepSchool(w, a) {
    const opts = (w.schools || []).map(s => `<option value="${escHtml(s)}"${s === w.school ? " selected" : ""}>${escHtml(s)}</option>`).join("");
    const pills = AUDIT_MONTHS.map(m => `<button type="button" class="btn btn-sm ${m === w.month ? "btn-primary" : "btn-outline-secondary"}" data-wiz-month="${m}">${m.slice(0, 3)}</button>`).join(" ");
    const st = QBR._storage ? "✓" : "○", us = QBR._usage ? "✓" : "○";
    return `<div class="wiz-step">
      <h5>Which school and month?</h5>
      <div class="mb-3"><label class="form-label">School</label>
        <select class="form-select" id="wiz-school" style="max-width:420px">${opts}</select></div>
      <div class="mb-3"><label class="form-label">Month</label><div class="d-flex flex-wrap gap-1">${pills}</div></div>
      <div class="small text-muted mb-3">Storage workbook ${st} · Usage workbook ${us} ${(!QBR._storage || !QBR._usage) ? "— upload the missing file(s) to enable those steps (you can skip them for now)." : ""}</div>
      ${wizNav(0, {})}
    </div>`;
  }

  function wizStepRisky(w, a) {
    const found = wizRec(a, w.school, w.month);
    const cur = found ? found.rec.risky : "—";
    // Month-over-month context: previous month's count + delta.
    let momHtml = "";
    try {
      const prev = auditPrevMonth(w.month);
      const prow = prev ? auditFindRow(a, prev, w.school) : null;
      if (prow && prow.risky != null && String(prow.risky).trim() !== "") {
        momHtml = ` <span class="text-muted small">(${escHtml(prev)}: <b>${escHtml(prow.risky)}</b>${auditMomBadge(a, w.month, { school: w.school, risky: cur === "—" ? "" : cur })})</span>`;
      }
    } catch (e) {}
    return `<div class="wiz-step">
      <h5>Risky users — ${escHtml(w.school)}</h5>
      <p class="text-muted">Current value for ${escHtml(w.month)}: <b>${escHtml(cur)}</b>${momHtml}</p>
      ${found ? `<div class="mb-3"><label class="form-label">New risky-user count</label>
        <input type="number" min="0" class="form-control" id="wiz-risky" style="max-width:200px" value="${escHtml(cur === "—" ? "" : cur)}"></div>
        ${wizNav(1, { save: true, saveLabel: "Save count" })}`
      : `<div class="alert alert-warning">No row for this school in ${escHtml(w.month)}. Add the school in the table first, or pick another month.</div>${wizNav(1, {})}`}
    </div>`;
  }

  function wizStepHealth(w, a) {
    const found = wizRec(a, w.school, w.month);
    const cur = found ? found.rec.health : "";
    const dd = (a.dropdowns && a.dropdowns.health) || [];
    const radios = dd.map(h => `<label class="wiz-radio"><input type="radio" name="wiz-health" value="${escHtml(h)}"${h === cur ? " checked" : ""}> ${escHtml(h)}</label>`).join("");
    return `<div class="wiz-step">
      <h5>Domain health — ${escHtml(w.school)}</h5>
      <p class="text-muted">Current: <b>${escHtml(cur || "—")}</b></p>
      ${found ? `<div class="mb-3 wiz-radios">${radios}</div>${wizNav(2, { save: true, saveLabel: "Save health" })}`
      : `<div class="alert alert-warning">No row for this school in ${escHtml(w.month)}.</div>${wizNav(2, {})}`}
    </div>`;
  }

  function wizStepStorage(w, a) {
    const s = QBR._storage;
    if (!s) return `<div class="wiz-step"><h5>Storage report</h5><div class="alert alert-info">Upload <code>STORAGE_REPORT.xlsx</code> to enable this step.</div>${wizNav(3, {})}</div>`;
    const q = QBR.auditQuarterOfMonth(w.month);
    const row = QBR.quarterFindRow(s, q, w.school);
    if (!row) return `<div class="wiz-step"><h5>Storage report</h5><div class="alert alert-warning">No row for ${escHtml(w.school)} in ${escHtml(q)}.</div>${wizNav(3, {})}</div>`;
    const heads = QBR.storagePasteHeaders();
    let prev = "";
    if (w.p3.parsed && !w.p3.parsed.error) {
      const rows = w.p3.parsed.values.map((v, i) => `<tr><td>${escHtml(heads[i])}</td><td><b>${escHtml(v)}</b></td></tr>`).join("");
      prev = `<table class="table table-sm wiz-prev"><tbody>${rows}</tbody></table>
        <div class="small text-muted mb-2">Writes to ${escHtml(row.sheet)} row ${row.r}, columns C–I. The Percentage% formula is never touched.</div>`;
    } else if (w.p3.parsed && w.p3.parsed.error) {
      prev = `<div class="alert alert-danger">${escHtml(w.p3.parsed.error)}</div>`;
    }
    return `<div class="wiz-step">
      <h5>Storage report — ${escHtml(w.school)} (${escHtml(q)})</h5>
      <p class="text-muted small">Copy the 7 values from your PowerShell CSV row and paste below (tab- or comma-separated, headers optional).</p>
      <textarea class="form-control mb-2" id="wiz-paste3" rows="3" style="max-width:640px" placeholder="1.15 TB&#9;87.25 GB&#9;…">${escHtml(w.p3.text)}</textarea>
      <div class="mb-2"><button type="button" class="btn btn-sm btn-outline-secondary" data-wiz="parse3">Parse &amp; preview</button></div>
      ${prev}
      ${w.p3.parsed && !w.p3.parsed.error ? wizNav(3, { save: true, saveLabel: "Save storage row" }) : wizNav(3, {})}
    </div>`;
  }

  function wizStepUsage(w, a) {
    const u = QBR._usage;
    if (!u) return `<div class="wiz-step"><h5>Usage report</h5><div class="alert alert-info">Upload <code>USAGE_REPORT.xlsx</code> to enable this step.</div>${wizNav(4, {})}</div>`;
    const q = QBR.auditQuarterOfMonth(w.month);
    const row = QBR.quarterFindRow(u, q, w.school);
    if (!row) return `<div class="wiz-step"><h5>Usage report</h5><div class="alert alert-warning">No row for ${escHtml(w.school)} in ${escHtml(q)}.</div>${wizNav(4, {})}</div>`;
    const heads = QBR.usagePasteHeaders();
    let prev = "";
    if (w.p4.parsed && !w.p4.parsed.error) {
      const vals = w.p4.parsed.values;
      const checks = QBR.auditValidateUsage(vals);
      const rows = vals.map((v, i) => {
        let flag = "";
        if (i >= 4) {
          const g = Math.floor((i - 4) / 3), pos = (i - 4) % 3;
          if (pos === 2) flag = checks[g].ok ? ' <span class="text-success">✓</span>' : ' <span class="text-danger">⚠ total ≠ active+inactive</span>';
        }
        return `<tr><td>${escHtml(heads[i])}</td><td><b>${escHtml(v)}</b>${flag}</td></tr>`;
      }).join("");
      prev = `<table class="table table-sm wiz-prev"><tbody>${rows}</tbody></table>
        <div class="small text-muted mb-2">Writes to ${escHtml(row.sheet)} row ${row.r}, columns D–V. Date Extracted = ${escHtml(w.dateStr)}, Report Period = ${escHtml(w.period)}. The % formula column is never touched; calculated Total columns keep their formulas.</div>`;
    } else if (w.p4.parsed && w.p4.parsed.error) {
      prev = `<div class="alert alert-danger">${escHtml(w.p4.parsed.error)}</div>`;
    }
    return `<div class="wiz-step">
      <h5>Usage report — ${escHtml(w.school)} (${escHtml(q)})</h5>
      <p class="text-muted small">Paste the 19 values (tab- or comma-separated, headers optional).</p>
      <textarea class="form-control mb-2" id="wiz-paste4" rows="3" style="max-width:640px" placeholder="5335&#9;2003&#9;…">${escHtml(w.p4.text)}</textarea>
      <div class="mb-2"><button type="button" class="btn btn-sm btn-outline-secondary" data-wiz="parse4">Parse &amp; preview</button></div>
      ${prev}
      ${w.p4.parsed && !w.p4.parsed.error ? wizNav(4, { save: true, saveLabel: "Save usage row" }) : wizNav(4, {})}
    </div>`;
  }

  // One save button for the whole wizard flow: writes every audit-family kind
  // (risky / storage / usage) that has journaled changes, then shows a summary.
  QBR.wizardSaveAll = async function () {
    const note = (typeof QBR.persistNote === "function") ? QBR.persistNote : null;
    const kinds = ["audit", "storage", "usage"].filter(k => {
      try { return QBR.journalUnsavedCount(k) > 0; } catch (e) { return false; }
    });
    if (!kinds.length) { if (note) note("audit", "No changes to save", 3000); return; }
    if (note) note("audit", "Saving…", 15000);
    const done = [];
    for (const k of kinds) {
      try { done.push({ kind: k, r: await QBR.auditFamilySave(k) }); }
      catch (e) { done.push({ kind: k, r: { mode: "error" } }); }
    }
    const names = { audit: "Risky", storage: "Storage", usage: "Usage" };
    const parts = done.map(d => {
      const n = names[d.kind] || d.kind, m = (d.r && d.r.mode) || "error";
      if (m === "file") return `${n}: saved ✓`;
      if (m === "download" || m === "download-fallback") return `${n}: downloaded`;
      if (m === "no-changes") return `${n}: no changes`;
      return `${n}: failed`;
    });
    if (note) note("audit", parts.join(" · "), 8000);
    QBR.renderAudit();
  };

  function wizStepDone(w, a) {
    const items = [];
    if (w.saved[1]) items.push("Risky-user count saved");
    if (w.saved[2]) items.push("Domain health saved");
    if (w.saved[3]) items.push("Storage row saved");
    if (w.saved[4]) items.push("Usage row saved");
    // Incomplete-row warning: risky count without health status, or vice versa.
    let warnHtml = "";
    try {
      const rec = auditFindRow(a, w.month, w.school);
      if (rec) {
        const hasRisky = rec.risky != null && String(rec.risky).trim() !== "";
        const hasHealth = rec.health != null && String(rec.health).trim() !== "";
        if (hasRisky !== hasHealth) {
          warnHtml = `<div class="alert alert-warning small mt-2">Incomplete: ` +
            (hasRisky ? "risky-user count is set but <b>domain health</b> is empty." : "<b>Risky-user count</b> is empty but domain health is set.") +
            `</div>`;
        }
      }
    } catch (e) {}
    const list = (items.length ? `<ul>${items.map(i => `<li>${escHtml(i)}</li>`).join("")}</li></ul>`
      : `<p class="text-muted">Nothing was saved for this school yet.</p>`) + warnHtml;
    return `<div class="wiz-step">
      <h5>Done — ${escHtml(w.school)} · ${escHtml(w.month)}</h5>
      ${list}
      <p><button type="button" class="btn btn-primary" id="wiz-save-all">💾 Save all to Excel <span data-unsaved-all class="badge bg-light text-dark"></span></button></p>
      <p class="small text-muted">Writes every pending change (risky users, storage, usage) into the workbooks.</p>
      <div class="wiz-nav">
        <button type="button" class="btn btn-outline-secondary" data-wiz="back">← Back</button>
        <button type="button" class="btn btn-primary" data-wiz="nextschool">Next school →</button>
        <button type="button" class="btn btn-outline-secondary" id="wiz-exit2">Back to table</button>
      </div>
    </div>`;
  }

  function wizMsg(t, ok) {
    const m = document.getElementById("wiz-msg");
    if (m) { m.innerHTML = t ? `<span class="${ok ? "text-success" : "text-danger"}">${escHtml(t)}</span>` : ""; }
  }

  function wireWizardStep(host, w, a, step) {
    const e2 = document.getElementById("wiz-exit2");
    if (e2) e2.addEventListener("click", QBR.auditWizardExit);
    host.querySelectorAll("[data-wiz-month]").forEach(b => b.addEventListener("click", () => {
      w.month = b.dataset.wizMonth;
      w.schools = wizEligibleSchools(a, w.month);
      if (w.schools.indexOf(w.school) < 0) w.school = w.schools[0] || "";
      w.step = 0; w.saved = {}; w.p3 = { text: "", parsed: null }; w.p4 = { text: "", parsed: null };
      renderAudit();
    }));
    const sch = document.getElementById("wiz-school");
    if (sch) sch.addEventListener("change", () => { w.school = sch.value; w.saved = {}; w.p3.parsed = null; w.p4.parsed = null; renderAudit(); });
    host.querySelectorAll('[data-wiz="parse3"]').forEach(b => b.addEventListener("click", () => {
      w.p3.text = document.getElementById("wiz-paste3").value;
      w.p3.parsed = QBR.auditParseStoragePaste(w.p3.text);
      renderAudit();
    }));
    host.querySelectorAll('[data-wiz="parse4"]').forEach(b => b.addEventListener("click", () => {
      w.p4.text = document.getElementById("wiz-paste4").value;
      w.p4.parsed = QBR.auditParseUsagePaste(w.p4.text);
      renderAudit();
    }));
    host.querySelectorAll('[data-wiz="back"]').forEach(b => b.addEventListener("click", () => { w.step = Math.max(0, w.step - 1); renderAudit(); }));
    host.querySelectorAll('[data-wiz="next"]').forEach(b => b.addEventListener("click", () => { w.step = Math.min(5, w.step + 1); renderAudit(); }));
    host.querySelectorAll('[data-wiz="finish"]').forEach(b => b.addEventListener("click", () => { w.step = 5; renderAudit(); }));
    host.querySelectorAll('[data-wiz="nextschool"]').forEach(b => b.addEventListener("click", () => {
      const i = (w.schools || []).indexOf(w.school);
      w.school = w.schools[(i + 1) % w.schools.length] || w.school;
      w.step = 1; w.saved = {}; w.p3 = { text: "", parsed: null }; w.p4 = { text: "", parsed: null };
      renderAudit();
    }));
    host.querySelectorAll('[data-wiz="save"]').forEach(b => b.addEventListener("click", () => wizSaveStep(w, a, step)));
  }

  function wizSaveStep(w, a, step) {
    if (step === 1) {
      const found = wizRec(a, w.school, w.month);
      if (!found) return wizMsg("No row for this school in " + w.month + ".", false);
      const v = document.getElementById("wiz-risky").value;
      if (v === "" || isNaN(Number(v)) || Number(v) < 0) return wizMsg("Enter a valid count (0 or more).", false);
      if (!QBR.auditUpdateCell(a.fp, found.ms.sheet, found.rec.r, "risky", String(Number(v)))) return wizMsg("Could not journal the change.", false);
      w.saved[1] = true; wizMsg("Risky-user count saved ✓", true);
    } else if (step === 2) {
      const found = wizRec(a, w.school, w.month);
      if (!found) return wizMsg("No row for this school in " + w.month + ".", false);
      const sel = document.querySelector('input[name="wiz-health"]:checked');
      if (!sel) return wizMsg("Pick a domain-health option.", false);
      if (!QBR.auditUpdateCell(a.fp, found.ms.sheet, found.rec.r, "health", sel.value)) return wizMsg("Could not journal the change.", false);
      w.saved[2] = true; wizMsg("Domain health saved ✓", true);
    } else if (step === 3) {
      const s = QBR._storage;
      const q = QBR.auditQuarterOfMonth(w.month);
      const row = s && QBR.quarterFindRow(s, q, w.school);
      if (!row) return wizMsg("No storage row for this school.", false);
      if (!w.p3.parsed || w.p3.parsed.error) return wizMsg("Parse the paste first.", false);
      if (!QBR.storagePasteRow(s.fp, row.sheet, row.r, w.p3.parsed.values)) return wizMsg("Could not journal the change.", false);
      w.saved[3] = true; wizMsg("Storage row saved ✓", true);
    } else if (step === 4) {
      const u = QBR._usage;
      const q = QBR.auditQuarterOfMonth(w.month);
      const row = u && QBR.quarterFindRow(u, q, w.school);
      if (!row) return wizMsg("No usage row for this school.", false);
      if (!w.p4.parsed || w.p4.parsed.error) return wizMsg("Parse the paste first.", false);
      if (!QBR.usagePasteRow(u.fp, row.sheet, row.r, w.p4.parsed.values, w.dateStr, w.period)) return wizMsg("Could not journal the change.", false);
      w.saved[4] = true; wizMsg("Usage row saved ✓", true);
    }
    if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
  }

  // Start the watcher shortly after load (Chrome/Edge only; silent elsewhere).
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => setTimeout(QBR.auditWatchStart, 5000));
  else setTimeout(QBR.auditWatchStart, 5000);
})();
