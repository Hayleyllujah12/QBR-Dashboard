/* ============================================================================
 * QBR Inventory module — laptop/desktop/monitor inventory + support tickets
 * v1.0.0 · 2026-10-04
 *
 * Lives INSIDE the QBR dashboard (new "Inventory" nav pages). Loads BEFORE
 * js/app.js; everything hangs off the shared QBR namespace plus two globals
 * the app shell calls: renderInventory() and renderAsset360Panel().
 *
 * Conventions (mirroring the existing engine):
 *  - Sheets are found by NAME (findSheet) and columns by HEADER PATTERN
 *    (makeResolver) — both borrowed from js/excel-loader.js, which loads first.
 *  - Serial Number is the canonical join key: QBR.invSerialKey() normalizes
 *    (uppercase, trim, collapse whitespace, strip invisible marks). The key is
 *    never displayed; rows keep their raw `sn` for display.
 *  - Missing data renders "—"/"No data", never 0. All user strings go through
 *    esc() before innerHTML.
 *  - Parsers are pure over row-arrays (invParse*Rows) so they run under Node;
 *    only QBR.parseInventoryFile/parseInventoryBuffers touch XLSX/DOM.
 * ==========================================================================*/
var QBR = window.QBR = window.QBR || {};

/* Delivery package version (semver MAJOR.MINOR.PATCH — see VERSIONING.md).
 * Single source of truth for the shipped zip name qbr-inventory-app-<ver>.zip
 * and the version badge on the Inventory page. */
QBR.INV_VERSION = "1.22.0";

/* ---------- Lenovo warranty lookup ---------------------------------------
 * Generic lookup page (per Pedro): paste any serial number. Deep per-unit
 * URLs were tried but Lenovo's product paths don't resolve reliably, so every
 * warranty link goes here. */
QBR.WARRANTY_LOOKUP_URL = "https://pcsupport.lenovo.com/us/en/warranty-lookup#/";
QBR.invWarrantyUrl = function (serial, model) {
  return QBR.WARRANTY_LOOKUP_URL;
};
QBR.INV_THRESH = { AGING_DAYS: 7, LEMON_TICKETS: 3, WARRANTY_WARN_DAYS: 90 };
QBR.INV_STATUS = ["In Stock", "Deployed", "In Repair", "Retired"];

/* ---------- serial key normalization ------------------------------------ */
QBR.invSerialKey = function (s) {
  return String(s == null ? "" : s)
    .replace(/[\u200B-\u200F\u202A-\u202E\u2060-\u2064\uFEFF]/g, "")
    .toUpperCase().trim().replace(/\s+/g, " ");
};

/* ---------- dates -------------------------------------------------------- */
function invDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000)); // Excel serial -> UTC
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  const t = Date.parse(String(v));
  if (!isNaN(t)) { const d = new Date(t); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  return null;
}
function invFmtDate(d) {
  if (!d) return "—";
  // Defensive: coerce strings/serials (e.g. a journal-replayed ISO date that
  // missed normalization) instead of throwing "d.getFullYear is not a function".
  const dd = (d instanceof Date && !isNaN(d)) ? d : invDate(d);
  if (!dd) return "—";
  const p = n => String(n).padStart(2, "0");
  return dd.getFullYear() + "-" + p(dd.getMonth() + 1) + "-" + p(dd.getDate());
}
function invDaysBetween(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 864e5);
}
function invToday() { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

/* ---------- sheet/column resolution ------------------------------------- */
/* Source-coordinate helpers for patch-in-place save (see js/patch.js).
 * Parsers tag each object with _src {fp, sheet, row} and record the resolved
 * column map per sheet in QBR._sheetMeta, so the saver can write cells back
 * into the ORIGINAL workbook instead of rebuilding it. */
function recSheetMeta(meta, cols, kind) {
  if (!meta || !meta.fp || !meta.sheet || typeof QBR === "undefined") return;
  const per = (QBR._sheetMeta || (QBR._sheetMeta = {}));
  (per[meta.fp] || (per[meta.fp] = {}))[meta.sheet] = { kind: kind || null, cols: Object.assign({}, cols) };
}
function srcOf(meta, r) {
  return (meta && meta.fp) ? { fp: meta.fp, sheet: meta.sheet, row: r + 1 } : null;
}
const INV_SHEETS = {
  devices:  ["02 DEVICES", "DEVICES"],
  support:  ["03 TECH SUPPORT LOGS", "TECH SUPPORT LOGS", "SUPPORT LOGS"],
  raksoinv: ["04 RAKSO INV.", "04 RAKSO INV", "RAKSO INV", "04 RAKSO INVENTORY", "RAKSO INVENTORY"],
  pipeline: ["06 PIPELINE", "PIPELINE"],
  po:       ["07 PURCHASE ORDER", "PURCHASE ORDER"],
};
const SENTINELS = new Set(["", "n/a", "na", "-", "—", "–", "none", "null"]);

/* ---------- pure row parsers (Node-testable) ----------------------------- */

function invParseDeviceRows(rows, meta) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = {
    sn: idx(["serial number", "serial"]), client: idx(["client / organization", "client", "organization"]),
    model: idx(["model"]), desc: idx(["description", "specification"]), cat: idx(["category"]),
    brand: idx(["brand"]), supplier: idx(["supplier"]), dr: idx(["dr #", "dr"]),
    delivered: idx(["date delivered"]), wstart: idx(["warranty start"]), wend: idx(["warranty end"]),
    wyears: idx(["warranty years"]), cond: idx(["condition"]), contact: idx(["contact person"]),
    addr: idx(["address"]), phone: idx(["contact details", "contact"]),
  };
  if (c.sn < 0) return out;
  recSheetMeta(meta, c, "devices");
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r], sn = row[c.sn];
    if (sn == null || String(sn).trim() === "") continue;
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : null);
    out.push({
      key: QBR.invSerialKey(sn), sn: String(sn).trim(),
      client: get(c.client), model: get(c.model), desc: get(c.desc), cat: get(c.cat) || "Laptop",
      brand: get(c.brand) || "Lenovo", supplier: get(c.supplier), dr: get(c.dr),
      delivered: invDate(row[c.delivered]), wstart: invDate(row[c.wstart]), wend: invDate(row[c.wend]),
      wyears: row[c.wyears] != null && row[c.wyears] !== "" ? Number(row[c.wyears]) : null,
      cond: get(c.cond) || "No Issue", contact: get(c.contact), addr: get(c.addr), phone: get(c.phone),
      statusOverride: null, _src: srcOf(meta, r),
    });
  }
  return out;
}

/* ---------- ticket numbers: {SERIAL}-{YYYYMMDD} (+ -2, -3 on same-day collisions) -- */
function invTicketNo(sn, date, taken) {
  const d = (date instanceof Date && !isNaN(date)) ? date : invToday();
  const ymd = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
  const base = String(sn).trim().toUpperCase() + "-" + ymd;
  let tno = base, i = 2;
  while (taken.has(tno)) tno = base + "-" + (i++);
  taken.add(tno);
  return tno;
}
/* Assign ticket numbers to parsed tickets missing one, oldest-first so the
 * sequence is deterministic (journal replay depends on it). */
function invBackfillTicketNos(tickets) {
  const taken = new Set(tickets.map(t => t.tno).filter(Boolean));
  tickets
    .filter(t => !t.tno)
    .sort((a, b) => {
      const da = a.repEdtech ? a.repEdtech.getTime() : Infinity;
      const db = b.repEdtech ? b.repEdtech.getTime() : Infinity;
      return da - db;
    })
    .forEach(t => { t.tno = invTicketNo(t.sn, t.repEdtech, taken); });
}
/* Follow-up notes serialize as "YYYY-MM-DD: text" lines in one workbook cell. */
function invParseTicketNotes(v) {
  if (!v) return [];
  return String(v).split(/\n+/).map(l => l.trim()).filter(Boolean).map(l => {
    const m = l.match(/^(\d{4}-\d{2}-\d{2})\s*:\s*(.*)$/);
    return m ? { d: invDate(m[1]) || invToday(), text: m[2] } : { d: null, text: l };
  });
}
function invFmtTicketNotes(notes) {
  return (notes || []).map(n => {
    const d = n.d instanceof Date ? n.d : invDate(n.d);
    const ymd = d ? d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") : "—";
    return ymd + ": " + (n.text || "");
  }).join("\n");
}
function invParseSupportRows(rows, meta) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = {
    tno: idx(["ticket no", "ticket number", "ticket #", "ticket id"]),
    sn: idx(["serial number", "serial"]), client: idx(["client / organization", "client"]),
    model: idx(["model"]), cat: idx(["category"]),
    priority: idx(["priority"]), status: idx(["status"]),
    requester: idx(["requested by", "requester", "reported by", "customer"]),
    repEdtech: idx(["date reported to edtech", "reported to edtech"]),
    repLenovo: idx(["date reported to lenovo", "reported to lenovo"]),
    completed: idx(["date completed"]), pic: idx(["person in-charge", "person in charge"]),
    issue: idx(["issue"]), act: idx(["activities", "troubleshooting"]),
    notes: idx(["follow-up notes", "follow-ups", "followups", "updates", "update notes"]),
    related: idx(["related tickets", "linked tickets", "related ticket"]),
    svcAddr: idx(["service location"]),
  };
  if (c.sn < 0) return out;
  recSheetMeta(meta, c, "support");
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r], sn = row[c.sn];
    if (sn == null || String(sn).trim() === "") continue;
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : null);
    const cleanNA = v => (v != null && SENTINELS.has(v.toLowerCase())) ? null : v;
    const normStatus = s => /^(completed|closed)$/i.test(s || "") ? "Resolved" : (s || "Open");
    out.push({
      key: QBR.invSerialKey(sn), sn: String(sn).trim(),
      tno: cleanNA(get(c.tno)),
      client: cleanNA(get(c.client)), model: cleanNA(get(c.model)), cat: cleanNA(get(c.cat)) || "Hardware",
      priority: cleanNA(get(c.priority)), status: normStatus(cleanNA(get(c.status))),
      requester: cleanNA(get(c.requester)),
      repEdtech: invDate(row[c.repEdtech]), repLenovo: invDate(row[c.repLenovo]),
      completed: invDate(row[c.completed]),
      pic: get(c.pic), issue: get(c.issue), act: get(c.act),
      notes: invParseTicketNotes(get(c.notes)),
      related: (get(c.related) || "").split(/[,;\n]+/).map(s => s.trim()).filter(Boolean),
      svcAddr: get(c.svcAddr), _src: srcOf(meta, r),
    });
  }
  return out;
}

/* 04 RAKSO INV: forward-fill blank SQ/Client (continuation rows) and pivot the
 * three serial columns (S/N=laptop, DEKTOP-S/N=desktop, Monitor-S/N=monitor)
 * into one deployment record per asset. Skips N/A / - / blank sentinels. */
function invParseRaksoRows(rows, meta) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = {
    sq: idx(["sq"]), client: idx(["client"]), sn: idx(["s/n"]),
    lapDesc: idx(["laptop"]), lapQty: idx(["laptop qty"]),
    dsn: idx(["dektop-s/n", "desktop-s/n"]), msn: idx(["monitor- s/n", "monitor-s/n", "monitor s/n"]),
    ddesc: idx(["desktop"]), req: idx(["date requested"]), del: idx(["date delivered"]),
    rem: idx(["remarks"]),
  };
  if (c.sq < 0) return out;
  recSheetMeta(meta, c, "rakso");
  let fSq = null, fClient = null, fReq = null, fDel = null, fRem = null;
  const val = (row, i) => (i >= 0 && row[i] != null ? String(row[i]).trim() : "");
  const usable = v => v && !SENTINELS.has(v.toLowerCase());
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const sq = val(row, c.sq);
    if (usable(sq)) {
      fSq = sq; fClient = val(row, c.client) || null;
      fReq = invDate(row[c.req]); fDel = invDate(row[c.del]); fRem = val(row, c.rem) || null;
    }
    if (!fSq) continue; // no SQ context yet
    const push = (snRaw, type, desc) => {
      const sn = (snRaw || "").trim();
      if (!usable(sn)) return;
      out.push({ key: QBR.invSerialKey(sn), sn, type, desc: desc || null, sq: fSq,
                 client: fClient, req: fReq, delivered: fDel, remarks: fRem, _src: srcOf(meta, r) });
    };
    push(val(row, c.sn), "Laptop", val(row, c.lapDesc));
    push(val(row, c.dsn), "Desktop", val(row, c.ddesc));
    push(val(row, c.msn), "Monitor", null);
  }
  return out;
}

function invParsePipelineRows(rows) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = {
    sq: idx(["sq"]), client: idx(["client"]), qty: idx(["qty"]),
    laptop: idx(["laptop"]), desktop: idx(["desktop"]), proposal: idx(["proposal"]),
    delivery: idx(["delivery"]), date: idx(["date"]), rem: idx(["remarks"]), addr: idx(["address"]),
  };
  if (c.sq < 0) return out;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const sq = row[c.sq];
    if (sq == null || String(sq).trim() === "") continue;
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : null);
    const q = row[c.qty];
    out.push({ sq: String(sq).trim(), client: get(c.client),
      qty: q != null && q !== "" && isFinite(Number(q)) ? Number(q) : null,
      laptop: get(c.laptop), desktop: get(c.desktop), proposal: get(c.proposal),
      delivery: get(c.delivery), date: invDate(row[c.date]), rem: get(c.rem), addr: get(c.addr) });
  }
  return out;
}

function invParsePORows(rows) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = { po: idx(["po#", "po"]), school: idx(["school"]), units: idx(["units"]),
              qty: idx(["qty"]), link: idx(["link"]), date: idx(["date"]) };
  if (c.po < 0) return out;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r], po = row[c.po];
    if (po == null || String(po).trim() === "") continue;
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : null);
    const q = row[c.qty];
    out.push({ po: String(po).trim(), school: get(c.school), units: get(c.units),
      qty: q != null && q !== "" && isFinite(Number(q)) ? Number(q) : null,
      link: get(c.link), date: invDate(row[c.date]) });
  }
  return out;
}

/* Merge one parsed workbook's sheets into the accumulator. */
QBR.invParseSheets = function (sheets, ctx) {
  const fp = ctx && ctx.fp;
  const acc = QBR._invAcc || (QBR._invAcc = { assets: [], tickets: [], deployments: [], pipeline: [], pos: [], sources: {} });
  const grab = key => {
    for (const cand of INV_SHEETS[key]) if (sheets[cand]) return { name: cand, rows: sheets[cand] };
    // substring fallback, mirroring findSheet() semantics in excel-loader.js
    const norm = x => String(x).toLowerCase().replace(/\s+/g, "");
    for (const cand of INV_SHEETS[key]) {
      const hit = Object.keys(sheets).find(n => norm(n).includes(norm(cand)));
      if (hit) return { name: hit, rows: sheets[hit] };
    }
    return null;
  };
  const metaFor = g => (g && fp) ? { fp, sheet: g.name } : null;
  const dev = grab("devices");   if (dev)   { acc.assets.push(...invParseDeviceRows(dev.rows, metaFor(dev)));       acc.sources.devices = true; }
  const sup = grab("support");   if (sup)   { acc.tickets.push(...invParseSupportRows(sup.rows, metaFor(sup)));     acc.sources.support = true; }
  const rinv = grab("raksoinv"); if (rinv)  { acc.deployments.push(...invParseRaksoRows(rinv.rows, metaFor(rinv)));  acc.sources.raksoinv = true; }
  const pl = grab("pipeline");   if (pl)    { acc.pipeline.push(...invParsePipelineRows(pl.rows));    acc.sources.pipeline = true; }
  const po = grab("po");         if (po)    { acc.pos.push(...invParsePORows(po.rows));               acc.sources.po = true; }
  return acc;
};

/* Parse raw workbook buffers (ArrayBuffers) into the inventory model. Mirrors
 * the main loader: per-buffer, filename-independent, sheets by name. */
QBR.parseInventoryBuffers = function (buffers) {
  QBR._invAcc = null;
  const fps = QBR._currentFps || [];
  (buffers || []).forEach((buf, bi) => {
    try {
      const wb = XLSX.read(buf, { type: "array", cellStyles: true }); // cellStyles: keep !cols/styles for patch-in-place save
      const fp = fps[bi];
      // Retain the original workbook for patch-in-place saves (js/patch.js):
      // cell writes go into THESE sheets, preserving formulas and layout.
      if (fp) { (QBR._origWb || (QBR._origWb = {}))[fp] = wb; }
      const sheets = {};
      wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false }); });
      const nBefore = QBR._invAcc ? Object.keys(QBR._invAcc.sources).length : 0;
      QBR.invParseSheets(sheets, { fp: fp || null });
      const nAfter = QBR._invAcc ? Object.keys(QBR._invAcc.sources).length : 0;
      if (nAfter > nBefore && fps[bi]) {
        const set = QBR._kindByFp[fps[bi]] || (QBR._kindByFp[fps[bi]] = new Set());
        set.add("assets");
      }
    } catch (e) { /* a non-inventory workbook simply contributes nothing */ }
  });
  const acc = QBR._invAcc || { assets: [], tickets: [], deployments: [], pipeline: [], pos: [], sources: {} };
  QBR._invAcc = null;
  // de-dupe assets by serial key (newest file wins); tickets/deployments accumulate
  const seen = new Map();
  acc.assets.forEach(a => seen.set(a.key, a));
  acc.assets = [...seen.values()];
  acc.tickets.forEach((t, i) => { t._n = i; });   // legacy index identity (tno is canonical)
  invBackfillTicketNos(acc.tickets);              // oldest-first, deterministic for journal replay
  return acc;
};

/* ---------- ticket status helpers ------------------------------------------ */
const TIX_STATUSES = ["Open", "In Progress", "Waiting for parts", "Escalated", "Resolved"];
const TIX_PRIORITIES = ["Low", "Medium", "High"];
function invTixOpen(t) { return !/^(resolved|completed|closed)$/i.test(String((t && t.status) || "Open")); }
function invTixStatus(t) { return /^(completed|closed)$/i.test(String((t && t.status) || "")) ? "Resolved" : (String((t && t.status) || "Open")); }
function invTixLink(tno, label) {
  return `<a href="#ticket/${encodeURIComponent(tno)}" target="_blank" rel="noopener" class="inv-link"><code>${esc(label || tno)}</code></a>`;
}
function invAssetLink(key, label) {
  return `<a href="#asset/${encodeURIComponent(key)}" target="_blank" rel="noopener" class="inv-link"><code>${esc(label)}</code></a>`;
}
/* 2026-10-05: serials that aren't in 02 DEVICES no longer render as dead links.
 * Loose key (letters+digits only) is used ONLY to suggest close matches —
 * stored serials and the join key are never changed or merged. */
QBR.invLooseKey = function (s) { return QBR.invSerialKey(s).replace(/[^A-Z0-9]/g, ""); };
const _invKeySetCache = new WeakMap();
function invHasAsset(key) {
  const inv = invModel(); if (!inv) return false;
  let set = _invKeySetCache.get(inv.assets);
  if (!set) { set = new Set(inv.assets.map(a => a.key)); _invKeySetCache.set(inv.assets, set); }
  return set.has(key);
}
function invEdit1(a, b) { // true when edit distance <= 1
  if (a === b) return true;
  const la = a.length, lb = b.length; if (Math.abs(la - lb) > 1) return false;
  let i = 0, j = 0, d = 0;
  while (i < la && j < lb) {
    if (a[i] === b[j]) { i++; j++; continue; }
    if (++d > 1) return false;
    if (la > lb) i++; else if (lb > la) j++; else { i++; j++; }
  }
  return d + (la - i) + (lb - j) <= 1;
}
QBR.invSuggestSerials = function (raw, max) {
  const inv = invModel(), lk = QBR.invLooseKey(raw);
  if (!inv || lk.length < 4) return [];
  const exact = inv.assets.filter(a => QBR.invLooseKey(a.sn) === lk);
  const near = lk.length >= 6 ? inv.assets.filter(a => exact.indexOf(a) < 0 && invEdit1(QBR.invLooseKey(a.sn), lk)) : [];
  return exact.concat(near).slice(0, max || 3);
};
function invSerialCell(key, sn) {
  return invHasAsset(key) ? invAssetLink(key, sn)
    : `<code>${esc(sn)}</code> <span class="inv-notinv" title="This serial isn't in 02 DEVICES">Not in inventory</span>`;
}
// Open the Register assets (intake) form prefilled with a serial — from any page.
QBR.invRegisterSerial = function (sn) {
  QBR._invUI.form = "intake";
  if (typeof goToTab === "function") goToTab("dash-inventory");
  try { history.replaceState(null, "", location.pathname + location.search + "#inventory"); } catch (e) {}
  renderInventory();
  const ta = $("in-serials");
  if (ta) { ta.value = String(sn || ""); ta.focus(); }
  const fh = $("inv-form-host");
  if (fh) fh.scrollIntoView({ block: "center" });
};
if (typeof document !== "undefined" && document.addEventListener) {
  document.addEventListener("click", e => {
    const b = e.target && e.target.closest && e.target.closest("[data-inv-register]");
    if (b) { e.preventDefault(); QBR.invRegisterSerial(b.getAttribute("data-inv-register")); }
  });
}
function invPrioPill(p) {
  const tone = { "High": "red", "Medium": "orange", "Low": "blue" }[p] || "gray";
  const colors = { blue: "#0078D4", green: "#107C10", orange: "#FF8C00", red: "#D13438", gray: "#605E5C", purple: "#8764B8" };
  return `<span class="badge" style="background:${colors[tone]};color:#fff">${esc(p || "—")}</span>`;
}
/* ---------- derived status ----------------------------------------------- */
QBR.invAssetStatus = function (a, openByKey) {
  if (a.statusOverride) return a.statusOverride;
  const cond = String(a.cond || "");
  if (/retired|dispos|beyond repair/i.test(cond)) return "Retired";
  if (/repair/i.test(cond)) return "In Repair";
  if (openByKey && openByKey.has(a.key)) return "In Repair";
  if (a.client && a.delivered) return "Deployed";
  return "In Stock";
};

QBR.invOpenByKey = function (tickets) {
  const m = new Map();
  (tickets || []).forEach(t => {
    if (invTixOpen(t)) {
      if (!m.has(t.key)) m.set(t.key, []);
      m.get(t.key).push(t);
    }
  });
  return m;
};

/* ---------- flag engine --------------------------------------------------- */
QBR.invComputeFlags = function (inv, today) {
  today = today || invToday();
  const T = QBR.INV_THRESH;
  const openByKey = QBR.invOpenByKey(inv.tickets);
  const assetByKey = new Map((inv.assets || []).map(a => [a.key, a]));
  const flags = { aging: [], lemons: [], warrantyExpiring: [], warrantyExpired: [], unmatched: [], stalled: [], duplicates: [] };
  (inv.tickets || []).forEach(t => {
    const open = invTixOpen(t);
    if (open && t.repEdtech && invDaysBetween(t.repEdtech, today) > T.AGING_DAYS) flags.aging.push(t);
    if (!assetByKey.has(t.key) && !flags.unmatched.some(u => u.key === t.key))
      flags.unmatched.push({ key: t.key, sn: t.sn, count: inv.tickets.filter(x => x.key === t.key).length });
  });
  const byKey = new Map();
  (inv.tickets || []).forEach(t => { byKey.set(t.key, (byKey.get(t.key) || 0) + 1); });
  byKey.forEach((n, key) => {
    if (n >= T.LEMON_TICKETS && assetByKey.has(key)) flags.lemons.push({ key, sn: assetByKey.get(key).sn, count: n });
  });
  (inv.assets || []).forEach(a => {
    if (!a.wend || QBR.invAssetStatus(a, openByKey) === "Retired") return;
    const d = invDaysBetween(today, a.wend);
    if (d != null && d < 0) flags.warrantyExpired.push(a);
    else if (d != null && d <= T.WARRANTY_WARN_DAYS) flags.warrantyExpiring.push(a);
  });
  (inv.pipeline || []).forEach(p => {
    if (/signed/i.test(String(p.proposal || "")) && !/delivered/i.test(String(p.delivery || "")))
      flags.stalled.push(p);
  });
  /* duplicate serials: same normalized key on 2+ asset rows */
  const seenKeys = new Map();
  (inv.assets || []).forEach(a => {
    const e = seenKeys.get(a.key);
    if (e) e.count++;
    else seenKeys.set(a.key, { key: a.key, sn: a.sn, count: 1 });
  });
  seenKeys.forEach(e => { if (e.count > 1) flags.duplicates.push(e); });
  return flags;
};

/* ---------- edit journal (audit trail for write-back) --------------------- */
QBR.invJournal = [];
QBR.invLog = function (action, detail) {
  if (QBR._replaying) return; // journal replay must not duplicate the edit log
  QBR.invJournal.push({ ts: new Date(), action, detail });
  if (QBR.invJournal.length > 500) QBR.invJournal.splice(0, QBR.invJournal.length - 500);
};

/* ---------- write-back mutations (in-memory + journal + re-render) -------- */
function invModel() { return (typeof APP !== "undefined" && APP.model && APP.model.inventory) || null; }

QBR.invIntake = function (list) {
  const inv = invModel(); if (!inv) return 0;
  let n = 0;
  (list || []).forEach(r => {
    const sn = (r.sn || "").trim();
    if (!sn) return;
    const key = QBR.invSerialKey(sn);
    if (inv.assets.some(a => a.key === key)) return; // no duplicates
    inv.assets.push({
      key, sn, client: r.client || null, model: r.model || null, desc: r.desc || null,
      cat: r.cat || "Laptop", brand: r.brand || "Lenovo", supplier: r.supplier || null,
      dr: r.dr || null, delivered: null, wstart: invDate(r.wstart) || invToday(),
      wend: null, wyears: r.wyears != null ? Number(r.wyears) : 3,
      cond: "No Issue", contact: null, addr: null, phone: null, statusOverride: null,
    });
    const a = inv.assets[inv.assets.length - 1];
    if (a.wstart && a.wyears) a.wend = new Date(a.wstart.getFullYear() + a.wyears, a.wstart.getMonth(), Math.min(a.wstart.getDate(), 28));
    QBR.invLog("intake", sn + (r.client ? " → " + r.client : " (in stock)"));
    n++;
  });
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invIntake", [list]);
  return n;
};

QBR.invDeploy = function (keys, client, dateDelivered, sq) {
  const inv = invModel(); if (!inv) return 0;
  const byKey = new Map(inv.assets.map(a => [a.key, a]));
  let n = 0;
  (keys || []).forEach(k => {
    const a = byKey.get(QBR.invSerialKey(k));
    if (!a) return;
    a.client = client || a.client; a.delivered = invDate(dateDelivered) || invToday();
    a.statusOverride = null;
    if (sq) inv.deployments.push({ key: a.key, sn: a.sn, type: a.cat, desc: a.desc, sq,
      client: a.client, req: null, delivered: a.delivered, remarks: "Deployed via dashboard" });
    QBR.invLog("deploy", a.sn + " → " + (client || "—"));
    n++;
  });
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invDeploy", [keys, client, dateDelivered, sq]);
  return n;
};

QBR.invAddTicket = function (t) {
  const inv = invModel(); if (!inv || !t || !t.sn) return false;
  const taken = new Set(inv.tickets.map(x => x.tno).filter(Boolean));
  const repEdtech = invDate(t.repEdtech) || invToday();
  const ticket = {
    _n: inv.tickets.length,
    tno: t.tno || invTicketNo(t.sn, repEdtech, taken),
    key: QBR.invSerialKey(t.sn), sn: String(t.sn).trim(), client: t.client || null,
    model: t.model || null, cat: t.cat || "Hardware",
    priority: t.priority || null, status: t.status || "Open",
    requester: t.requester || null,
    repEdtech: repEdtech, repLenovo: invDate(t.repLenovo) || null,
    completed: null, pic: t.pic || null, issue: t.issue || null, act: t.act || null,
    notes: (t.notes || []).map(n => ({ d: n.d instanceof Date ? n.d : (invDate(n.d) || invToday()), text: n.text || "" })),
    related: (t.related || []).map(String),
    svcAddr: t.svcAddr || null,
  };
  inv.tickets.push(ticket);
  QBR.invLog("ticket opened", ticket.tno + " — " + t.sn + " — " + (t.issue || "no issue text"));
  t.tno = ticket.tno; // keep the generated number on the journaled args (patch + replay)
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invAddTicket", [t]);
  return ticket.tno;
};
QBR.invFindTicket = function (tno) {
  const inv = invModel(); if (!inv || !tno) return null;
  return inv.tickets.find(t => String(t.tno).toUpperCase() === String(tno).trim().toUpperCase()) || null;
};
/* Edit ticket fields (issue, category, priority, requester, dates, person, …). */
QBR.invUpdateTicket = function (tno, patch) {
  const t = QBR.invFindTicket(tno); if (!t || !patch) return false;
  const d = k => (patch[k] === undefined ? t[k] : (patch[k] === "" ? null : patch[k]));
  ["client", "model", "cat", "priority", "requester", "pic", "issue", "act", "svcAddr"].forEach(k => { t[k] = d(k); });
  if (patch.repEdtech !== undefined) t.repEdtech = invDate(patch.repEdtech);
  if (patch.repLenovo !== undefined) t.repLenovo = invDate(patch.repLenovo);
  QBR.invLog("ticket updated", t.tno);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invUpdateTicket", [tno, patch]);
  return true;
};
/* Append a follow-up note to the ticket timeline. */
QBR.invTicketNote = function (tno, text) {
  const t = QBR.invFindTicket(tno); if (!t || !String(text || "").trim()) return false;
  t.notes = t.notes || [];
  t.notes.push({ d: invToday(), text: String(text).trim() });
  QBR.invLog("ticket note", t.tno + " — " + String(text).trim().slice(0, 80));
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invTicketNote", [tno, String(text).trim()]);
  return true;
};
/* Link another ticket as related (e.g. same issue recurring → new ticket). */

/* ================= MS FORMS TICKET IMPORT ================================
 * Imports an MS Forms "ticket support filing" export (.xlsx) as support
 * tickets. Dedup: a row is skipped when its Forms response ID was imported
 * before OR its full-row fingerprint matches (same data across all columns).
 * Seen IDs/fingerprints persist in localStorage; the Forms ID is also
 * stamped into each ticket's notes so the workbook itself carries the
 * memory (survives browser switches via the journal).
 */
const FORMS_COLS = {
  id: ["id"],
  start: ["start time"], completed: ["completion time"],
  fullName: ["full name"], workEmail: ["email address (work email)", "email address"], phone: ["phone number / mobile number", "phone number", "mobile number"],
  model: ["lenovo model", "model"], sn: ["serial number (sn)", "serial number"],
  usable: ["is the device currently usable?"], category: ["hardware issue category", "issue category"],
  desc: ["detailed description of the issue", "description of the issue"],
  firstSeen: ["when was the issue first observed?"], troubleshooted: ["has any troubleshooting been performed?"],
  trouble: ["troubleshooting performed"], media: ["video or pictures of issues encountered", "video or pictures"],
  priority: ["priority level", "priority"],
};
function formsColIdx(headerRow) {
  const heads = (headerRow || []).map(h => String(h == null ? "" : h).trim().toLowerCase());
  const out = {};
  Object.keys(FORMS_COLS).forEach(k => {
    out[k] = -1;
    for (const p of FORMS_COLS[k]) {
      let i = heads.indexOf(p);
      if (i < 0) i = heads.findIndex(h => h.indexOf(p) >= 0);
      if (i >= 0) { out[k] = i; break; }
    }
  });
  return out;
}
function formsFingerprint(cells) {
  const norm = (cells || []).map(c => String(c == null ? "" : c).trim().replace(/\s+/g, " ")).join("");
  let h = 5381;
  for (let i = 0; i < norm.length; i++) h = (((h << 5) + h) ^ norm.charCodeAt(i)) >>> 0;
  return h.toString(16);
}
function formsSeenLoad() {
  try {
    const s = JSON.parse(localStorage.getItem("qbr-forms-import-v1"));
    if (s && typeof s === "object") return { ids: s.ids || {}, fps: s.fps || {} };
  } catch (e) {}
  return { ids: {}, fps: {} };
}
function formsSeenSave(seen) {
  try { localStorage.setItem("qbr-forms-import-v1", JSON.stringify(seen)); } catch (e) {}
}
/* Forms IDs already in the model (stamped into ticket notes on import). */
function formsImportedIds() {
  const inv = invModel(); const ids = new Set();
  if (!inv) return ids;
  (inv.tickets || []).forEach(t => {
    (t.notes || []).forEach(n => {
      const mch = /Forms response #(\S+)/.exec(String((n && n.text) || ""));
      if (mch) ids.add(mch[1]);
    });
  });
  return ids;
}
/* Parse Forms-export row-arrays into ticket drafts.
 * Returns [{ row, formsId, fp, sn, ticket }]. Pure over rows (Node-testable). */
QBR.formsParseRows = function (rows) {
  if (!rows || rows.length < 2) return [];
  const ci = formsColIdx(rows[0]);
  const inv = invModel();
  const assetClient = sn => {
    if (!inv || !sn) return null;
    const a = inv.assets.find(x => x.key === QBR.invSerialKey(sn));
    return a ? a.client : null;
  };
  const dstr = v => {
    const d = invDate(v);
    if (!(d instanceof Date) || isNaN(d)) return null;
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  };
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (!row.some(c => c != null && String(c).trim() !== "")) continue; // blank row
    const g = k => (ci[k] >= 0 && row[ci[k]] != null ? String(row[ci[k]]).trim() : "");
    const snRaw = g("sn");
    const desc = g("desc"), usable = g("usable");
    const meta = [];
    const fs = dstr(row[ci.firstSeen]);
    if (fs) meta.push("First observed: " + fs);
    if (usable) meta.push("Device usable: " + usable);
    const issue = [desc, meta.length ? "(" + meta.join(" · ") + ")" : ""].filter(Boolean).join("\n");
    const tb = g("trouble");
    const act = tb || (/^no$/i.test(g("troubleshooted")) ? "No troubleshooting performed yet" : null);
    const formsId = g("id");
    const email = g("workEmail"), phone = g("phone"), media = g("media");
    const repDate = invDate(row[ci.completed]) || invDate(row[ci.start]) || null;
    const notes = [{ d: repDate || invToday(), text: "Imported from " + ["Forms response #" + (formsId || "?"), email, phone].filter(Boolean).join(" · ") }];
    if (media) notes.push({ d: notes[0].d, text: "Reporter attachments: " + media });
    out.push({
      row: r + 1, formsId: formsId || null, fp: formsFingerprint(row), sn: snRaw,
      ticket: {
        sn: snRaw, client: assetClient(snRaw), model: g("model") || null,
        cat: g("category") || "Hardware", priority: g("priority") || null, status: "Open",
        requester: g("fullName") || null, repEdtech: repDate,
        issue: issue || null, act: act, notes: notes,
      },
    });
  }
  return out;
};
/* Classify parsed rows: "new" | "duplicate" | "missing-sn".
 * Duplicate = response ID seen before (localStorage or ticket notes), an
 * identical full-row fingerprint seen before, or a repeat within this file. */
QBR.formsClassify = function (parsed) {
  const seen = formsSeenLoad(), noteIds = formsImportedIds();
  const batchIds = new Map(), batchFps = new Map(); // id/fp -> first row number
  return (parsed || []).map(p => {
    let status = "new", reason = "";
    const pastId = p.formsId && (seen.ids[p.formsId] || noteIds.has(p.formsId));
    const pastFp = seen.fps[p.fp];
    const batchRow = (p.formsId && batchIds.get(p.formsId)) || batchFps.get(p.fp);
    if (!p.sn) { status = "missing-sn"; reason = "No serial number"; }
    else if (pastId || pastFp) {
      status = "duplicate";
      reason = pastId ? "Response #" + p.formsId + " already imported" : "Identical row already imported";
    } else if (batchRow) {
      status = "duplicate";
      reason = "Duplicate of row " + batchRow + " in this file";
    }
    if (status !== "missing-sn") {
      if (p.formsId && !batchIds.has(p.formsId)) batchIds.set(p.formsId, p.row);
      if (!batchFps.has(p.fp)) batchFps.set(p.fp, p.row);
    }
    return Object.assign({}, p, { status, reason });
  });
};
/* Import rows classified as new (via invAddTicket → journaled). */
QBR.formsCommit = function (classified) {
  const inv = invModel(); if (!inv) return { imported: [], skipped: 0 };
  const seen = formsSeenLoad();
  const imported = [];
  let skipped = 0;
  (classified || []).forEach(p => {
    if (p.status !== "new") { skipped++; return; }
    const tno = QBR.invAddTicket(Object.assign({}, p.ticket));
    if (tno) {
      imported.push(tno);
      if (p.formsId) seen.ids[p.formsId] = 1;
      seen.fps[p.fp] = 1;
    } else skipped++;
  });
  formsSeenSave(seen);
  return { imported, skipped };
};
QBR.invLinkTicket = function (tno, rel) {
  const t = QBR.invFindTicket(tno); if (!t || !rel) return false;
  const r = String(rel).trim().toUpperCase();
  if (!r || r === String(t.tno).toUpperCase()) return false;
  t.related = t.related || [];
  if (!t.related.map(String).map(s => s.toUpperCase()).includes(r)) t.related.push(r);
  QBR.invLog("ticket linked", t.tno + " ↔ " + r);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invLinkTicket", [tno, r]);
  return true;
};
/* Change status; Resolved stamps the completion date. */
QBR.invSetTicketStatus = function (tno, status, note) {
  const t = QBR.invFindTicket(tno); if (!t || !status) return false;
  t.status = status;
  if (/^resolved$/i.test(status) && !t.completed) t.completed = invToday();
  if (/^open$/i.test(status)) t.completed = null;
  if (note && String(note).trim()) {
    t.notes = t.notes || [];
    t.notes.push({ d: invToday(), text: "Status → " + status + ": " + String(note).trim() });
  }
  QBR.invLog("ticket status", t.tno + " → " + status);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invSetTicketStatus", [tno, status, note || null]);
  return true;
};

QBR.invResolveTicket = function (key, n, note) {
  const inv = invModel(); if (!inv) return false;
  const t = inv.tickets.find(t => t._n === Number(n) && t.key === QBR.invSerialKey(key));
  if (!t) return false;
  return QBR.invSetTicketStatus(t.tno, "Resolved", note);
};

QBR.invSetStatus = function (key, status) {
  const inv = invModel(); if (!inv) return false;
  const a = inv.assets.find(x => x.key === QBR.invSerialKey(key));
  if (!a || QBR.INV_STATUS.indexOf(status) < 0) return false;
  a.statusOverride = status;
  QBR.invLog("status change", a.sn + " → " + status);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invSetStatus", [key, status]);
  return true;
};

/* Update editable asset fields (Asset 360 → "Edit details"). `patch` maps
 * field name → new value; only the keys present are touched. Dates accept
 * anything invDate() parses ("YYYY-MM-DD" from date inputs, Date objects,
 * Excel serials, ISO strings from the journal). Journaled so linked-file
 * saves and session restores keep the change. */
QBR.invUpdateAsset = function (key, patch) {
  const inv = invModel(); if (!inv) return false;
  const a = inv.assets.find(x => x.key === QBR.invSerialKey(key));
  if (!a || !patch || typeof patch !== "object") return false;
  const TEXT = ["client", "model", "desc", "cat", "brand", "supplier", "dr", "cond", "contact", "addr", "phone"];
  const DATES = ["delivered", "wstart", "wend"];
  const applied = {};
  TEXT.forEach(f => {
    if (patch[f] !== undefined) {
      const v = (patch[f] == null || String(patch[f]).trim() === "") ? null : String(patch[f]).trim();
      a[f] = v; applied[f] = v;
    }
  });
  DATES.forEach(f => {
    if (patch[f] !== undefined) { const v = invDate(patch[f]); a[f] = v; applied[f] = v; }
  });
  if (patch.wyears !== undefined) {
    const n = (patch.wyears == null || patch.wyears === "") ? null : Number(patch.wyears);
    a.wyears = (n != null && isFinite(n)) ? n : null; applied.wyears = a.wyears;
  }
  const keys = Object.keys(applied);
  if (!keys.length) return false;
  QBR.invLog("asset update", a.sn + ": " + keys.join(", "));
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("assets", "invUpdateAsset", [key, applied]);
  return true;
};

/* Known warranty END (2026-10-05): sets wend as typed and derives warranty years
 * from the unit's start date (1 decimal; whole number when within ~2 weeks), so
 * the expiry flags (which read wend) are exact. Goes through invUpdateAsset →
 * journaled + saved like any edit. */
QBR.invSetWarrantyEnd = function (key, end) {
  const inv = invModel(); if (!inv) return false;
  const a = inv.assets.find(x => x.key === QBR.invSerialKey(key)); if (!a) return false;
  const we = invDate(end); if (!we) return false;
  const patch = { wend: we };
  if (a.wstart instanceof Date && !isNaN(a.wstart) && we > a.wstart) {
    const yrs = (we - a.wstart) / (365.25 * 864e5), r = Math.round(yrs);
    patch.wyears = Math.abs(yrs - r) < 0.04 ? r : Math.round(yrs * 10) / 10;
  }
  return QBR.invUpdateAsset(a.key, patch);
};

/* ---------- export: regenerate the workbook from the model ---------------- */
// Build the inventory workbook object (shared by Export download and direct save).
QBR.invBuildWorkbook = function () {
  const inv = invModel();
  if (!inv || typeof XLSX === "undefined") return null;
  const openByKey = QBR.invOpenByKey(inv.tickets);
  const dev = [["Serial Number","Client / Organization","Model","Description / Specifications","Category",
    "Brand","Supplier","DR #","Date Delivered","Warranty Start","Warranty End","Warranty Years",
    "Condition","Contact Person","Address","Contact Details","Status (derived)"]];
  inv.assets.forEach(a => dev.push([a.sn, a.client || "", a.model || "", a.desc || "", a.cat, a.brand || "",
    a.supplier || "", a.dr || "", invFmtDate(a.delivered) === "—" ? "" : invFmtDate(a.delivered),
    invFmtDate(a.wstart) === "—" ? "" : invFmtDate(a.wstart), invFmtDate(a.wend) === "—" ? "" : invFmtDate(a.wend),
    a.wyears == null ? "" : a.wyears, a.cond || "", a.contact || "", a.addr || "", a.phone || "",
    QBR.invAssetStatus(a, openByKey)]));
  const sup = [["Ticket No","Serial Number","Client / Organization","Model","Category","Priority","Status","Requested By",
    "Date Reported To EdTech","Date Reported To Lenovo Support","Date Completed","Person In-Charge","Issue/s",
    "Activities / Troubleshooting","Follow-up Notes","Related Tickets","Service Location Address"]];
  inv.tickets.forEach(t => sup.push([t.tno || "", t.sn, t.client || "", t.model || "", t.cat, t.priority || "", invTixStatus(t),
    t.requester || "",
    invFmtDate(t.repEdtech) === "—" ? "" : invFmtDate(t.repEdtech),
    invFmtDate(t.repLenovo) === "—" ? "" : invFmtDate(t.repLenovo),
    invFmtDate(t.completed) === "—" ? "" : invFmtDate(t.completed),
    t.pic || "", t.issue || "", t.act || "",
    invFmtTicketNotes(t.notes), (t.related || []).join(", "), t.svcAddr || ""]));
  const dep = [["SQ","Client","Type","S/N","Description","Date Requested","Date Delivered","Remarks"]];
  inv.deployments.forEach(d => dep.push([d.sq || "", d.client || "", d.type || "", d.sn,
    d.desc || "", invFmtDate(d.req) === "—" ? "" : invFmtDate(d.req),
    invFmtDate(d.delivered) === "—" ? "" : invFmtDate(d.delivered), d.remarks || ""]));
  const pl = [["SQ","Client","QTY","LAPTOP","DESKTOP","PROPOSAL","DELIVERY","DATE","Remarks","Address"]];
  inv.pipeline.forEach(p => pl.push([p.sq, p.client || "", p.qty == null ? "" : p.qty, p.laptop || "",
    p.desktop || "", p.proposal || "", p.delivery || "", invFmtDate(p.date) === "—" ? "" : invFmtDate(p.date),
    p.rem || "", p.addr || ""]));
  const po = [["PO#","SCHOOL","UNITS","QTY","LINK","DATE"]];
  inv.pos.forEach(p => po.push([p.po, p.school || "", p.units || "", p.qty == null ? "" : p.qty,
    p.link || "", invFmtDate(p.date) === "—" ? "" : invFmtDate(p.date)]));
  const log = [["Timestamp","Action","Detail"]];
  QBR.invJournal.forEach(e => log.push([e.ts.toISOString(), e.action, e.detail]));
  const wb = XLSX.utils.book_new();
  [["02 DEVICES", dev], ["03 TECH SUPPORT LOGS", sup], ["04 RAKSO INV.", dep],
   ["06 PIPELINE", pl], ["07 PURCHASE ORDER", po], ["EDIT LOG", log]].forEach(([name, aoa]) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = aoa[0].map(h => ({ wch: Math.min(42, Math.max(12, String(h).length + 2)) }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  return { wb: wb, filename: "Lenovo_Inventory_Export.xlsx" };
};

QBR.invExportWorkbook = function () {
  const b = QBR.invBuildWorkbook();
  if (!b) { alert("Load an inventory workbook first."); return; }
  XLSX.writeFile(b.wb, b.filename, { cellStyles: true });
};

/* ---------- UI state + small view helpers -------------------------------- */
QBR._invUI = { client: "ALL", type: "ALL", status: "ALL", q: "", flag: null, showAll: false, form: null, prefillSn: null, view: "assets" };
QBR._invA360 = null;

function invPill(s) {
  const tone = { "In Stock": "blue", "Deployed": "green", "In Repair": "orange", "Retired": "gray",
                 "Open": "orange", "In Progress": "blue", "Waiting for parts": "purple",
                 "Escalated": "red", "Resolved": "green", "Completed": "green" }[s] || "blue";
  const colors = { blue: "#0078D4", green: "#107C10", orange: "#FF8C00", red: "#D13438", gray: "#605E5C", purple: "#8764B8" };
  return `<span class="badge" style="background:${colors[tone]};color:#fff">${esc(s)}</span>`;
}
function invClients(inv) {
  const s = new Set();
  (inv.assets || []).forEach(a => { if (a.client) s.add(a.client); });
  return [...s].sort();
}
function invAvgResolve(tickets) {
  const ds = tickets.filter(t => t.completed && t.repEdtech)
    .map(t => invDaysBetween(t.repEdtech, t.completed)).filter(d => d != null && d >= 0);
  return ds.length ? ds.reduce((a, b) => a + b, 0) / ds.length : null;
}

/* ---------- top stock overview: remaining-quantity tiles ----------------------
 * At-a-glance tiles in the KPI style: per-supply "Vellum Board / 400 pcs"
 * and per-model "Lenovo ThinkPad E14 Gen 6 / 15 pcs" (In Stock units).
 * Strictly separated: each view shows only its own tiles. */
function invStockOverview(inv, sup, view) {
  const E = (typeof esc === "function") ? esc : (s => s);
  const F = (typeof fmt === "function") ? fmt : (n => (n == null ? "—" : n));
  const tile = (label, value, tone) =>
    `<div class="col"><div class="kpi kpi-${tone}"><div class="kpi-val">${value}</div><div class="kpi-lbl">${E(label)}</div></div></div>`;
  let tiles = "", afterTiles = "";
  if (view === "supplies" && sup && (sup.items.length || sup.transactions.length) &&
      typeof QBR.supComputeSummary === "function") {
    const ssum = QBR.supComputeSummary(sup);
    const rows = [...ssum.per.values()].sort((a, b) => String(a.item.name).localeCompare(String(b.item.name)));
    tiles = rows.map(e => {
      const low = e.ending < 0 || e.restock;
      return tile(e.item.name, `${F(e.ending)} pcs`, low ? "red" : "green");
    }).join("");
  } else if (view === "assets" && inv && inv.assets.length && typeof QBR.invAssetStatus === "function") {
    /* Option C (hybrid): unit-type tiles + bundle-readiness strip + collapsible
     * per-model drill-down. Peripherals (chargers/bags/mice/keyboards) are not
     * tracked in the workbook, so a desktop set = system unit + monitor. */
    const openByKey = QBR.invOpenByKey(inv.tickets);
    const typeOrder = ["Laptop", "Desktop", "Monitor"];
    const byType = new Map(), byModel = new Map();
    inv.assets.forEach(a => {
      const st = QBR.invAssetStatus(a, openByKey);
      const cat = (a.cat || "Laptop").trim() || "Laptop";
      if (!byType.has(cat)) byType.set(cat, { cat, inStock: 0, deployed: 0, total: 0 });
      const gt = byType.get(cat); gt.total++;
      if (st === "In Stock") gt.inStock++;
      if (st === "Deployed") gt.deployed++;
      const label = ((a.brand || "") + " " + (a.model || "Unknown model")).trim().replace(/\s+/g, " ");
      if (!byModel.has(label)) byModel.set(label, { label, cat, inStock: 0, total: 0 });
      const gm = byModel.get(label); gm.total++;
      if (st === "In Stock") gm.inStock++;
    });
    const typeIcon = { Laptop: "💻", Desktop: "🖥️", Monitor: "🖥️" };
    const ordered = [...typeOrder.filter(c => byType.has(c)), ...[...byType.keys()].filter(c => !typeOrder.includes(c))];
    tiles = ordered.map(c => {
      const g = byType.get(c), tone = g.inStock === 0 ? "red" : "green";
      return `<div class="col"><button type="button" data-inv-typetile="${E(c)}" title="Show ${E(c)} assets"
        style="all:unset;display:block;width:100%;cursor:pointer"><div class="kpi kpi-${tone}">
        <div class="kpi-val">${F(g.inStock)} <span style="font-size:16px">pcs</span></div>
        <div class="kpi-lbl">${typeIcon[c] || "📦"} ${E(c)}s <span style="font-weight:400">· Deployed ${F(g.deployed)}</span></div>
      </div></button></div>`;
    }).join("");
    /* bundle strip: desktop set needs a system unit AND a monitor */
    const lap = (byType.get("Laptop") || { inStock: 0 }).inStock;
    const dIn = (byType.get("Desktop") || { inStock: 0 }).inStock;
    const mIn = (byType.get("Monitor") || { inStock: 0 }).inStock;
    const dSets = Math.min(dIn, mIn);
    const dNote = (dIn === 0 && mIn === 0) ? "—"
      : mIn < dIn ? `⚠ ${F(dIn - mIn)} monitor${dIn - mIn === 1 ? "" : "s"} short`
      : dIn < mIn ? `⚠ ${F(mIn - dIn)} system unit${mIn - dIn === 1 ? "" : "s"} short`
      : "✓ complete";
    afterTiles = `<div class="small mt-2 text-muted">Sets ready — 💻 Laptops <b>${F(lap)}</b>` +
      ` · 🖥️ Desktops <b>${F(dSets)}</b> <span class="${dNote.charAt(0) === "⚠" ? "text-danger" : "text-success"}">${dNote}</span></div>`;
    /* per-model drill-down (replaces the wall of per-model tiles) */
    const models = [...byModel.values()].sort((a, b) => b.inStock - a.inStock || b.total - a.total);
    const zeroN = models.filter(m => m.inStock === 0).length;
    afterTiles += `<details class="mt-2"><summary class="small text-muted" style="cursor:pointer">` +
      `${zeroN ? `${F(zeroN)} model${zeroN === 1 ? "" : "s"} out of stock · ` : ""}${F(models.length - zeroN)} in stock — model breakdown</summary>` +
      `<div class="table-responsive mt-1"><table class="table table-sm inv-tbl"><thead><tr><th>Model</th><th>Type</th><th class="text-end">In stock</th><th class="text-end">Total</th></tr></thead><tbody>` +
      models.map(m => `<tr${m.inStock === 0 ? ` class="text-muted"` : ""}><td>${E(m.label)}</td><td>${E(m.cat)}</td><td class="text-end">${F(m.inStock)}</td><td class="text-end">${F(m.total)}</td></tr>`).join("") +
      `</tbody></table></div></details>`;
  }
  if (!tiles) return "";
  return `<div class="row row-cols-2 row-cols-md-4 g-3 kpi-row">${tiles}</div>` + (afterTiles || "");
}

/* ============================ INVENTORY PAGE ============================= */
/* Collapsible card wrapper (assets table, pipeline, POs). State kept in
 * QBR._invUI.collapsed; toggled in-place via invBind (no re-render). */
function invCollapsible(id, titleHtml, bodyHtml) {
  const collapsed = (QBR._invUI.collapsed || {})[id];
  return `<div class="card-box mt-3" id="inv-card-${id}"><div class="d-flex justify-content-between align-items-center gap-2">
      <h6 class="mb-0">${titleHtml}</h6>
      <button type="button" class="btn btn-sm btn-link text-decoration-none p-0" data-inv-col="${id}"
        aria-label="${collapsed ? "Expand" : "Collapse"} section">${collapsed ? "▸" : "▾"}</button></div>
    <div id="inv-col-${id}" class="${collapsed ? "d-none" : ""} mt-2">${bodyHtml}</div></div>`;
}

/* Tickets table card (shared by the tickets view). Sort: ui.tixSort "old"|"new". */
function invTicketsCard(inv, ui, flags, today) {
  const tixQ = (ui.tixQ || "").trim().toUpperCase();
  const dir = ui.tixSort === "new" ? -1 : 1;
  const key = t => t.repEdtech ? t.repEdtech.getTime() : (dir > 0 ? Infinity : -Infinity);
  let tix = inv.tickets.slice().sort((a, b) => (key(a) - key(b)) * dir);
  if (tixQ) tix = tix.filter(t => (t.sn || "").toUpperCase().includes(tixQ) || String(t.tno || "").toUpperCase().includes(tixQ));
  if (ui.flag === "aging") { const s = new Set(flags.aging.map(t => t.sn + "|" + (t.repEdtech || ""))); tix = tix.filter(t => s.has(t.sn + "|" + (t.repEdtech || ""))); }
  if (ui.flag === "unmatched") { const s = new Set(flags.unmatched.map(u => u.key)); tix = tix.filter(t => s.has(t.key)); }
  if (ui.flag === "lemons") { const s = new Set(flags.lemons.map(l => l.key)); tix = tix.filter(t => s.has(t.key)); }
  return `<div class="card-box mt-3"><div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>Ticket No</th><th>Serial</th><th>Client</th><th>Priority</th><th>Status</th><th>Reported</th><th class="text-end">Days open</th><th>Issue</th><th>Person</th><th></th></tr></thead><tbody>` +
    (tix.slice(0, 200).map((t, i) => {
      const open = invTixOpen(t), st = invTixStatus(t);
      const days = open && t.repEdtech ? invDaysBetween(t.repEdtech, today) : (t.completed && t.repEdtech ? invDaysBetween(t.repEdtech, t.completed) : null);
      return `<tr><td>${invTixLink(t.tno)}</td>
        <td>${invSerialCell(t.key, t.sn)}</td>
        <td>${esc(t.client || "—")}</td><td>${invPrioPill(t.priority)}</td><td>${invPill(st)}</td>
        <td>${invFmtDate(t.repEdtech)}</td><td class="text-end">${days == null ? "—" : fmt(days)}</td>
        <td>${esc(t.issue || "—")}</td><td>${esc(t.pic || "—")}</td>
        <td>${open ? `<button type="button" class="btn btn-sm btn-outline-success" data-inv-resolve="${t._n}">Resolve</button>` : ""}</td></tr>`;
    }).join("") || `<tr><td colspan="10" class="text-muted">No tickets match.</td></tr>`) +
    `</tbody></table></div>${tix.length > 200 ? `<div class="small text-muted">Showing ${ui.tixSort === "new" ? "newest" : "oldest"} 200 of ${fmt(tix.length)}.</div>` : ""}</div>`;
}

function renderInventory() {
  const host = $("inv-body");
  if (!host) return;
  const inv = invModel();
  const sup = (typeof supModel === "function") ? supModel() : null;
  const hasAssets = inv && (inv.assets.length || inv.tickets.length || inv.deployments.length);
  const hasSup = sup && (sup.items.length || sup.transactions.length);
  /* Delivery version badge (semver — see VERSIONING.md), visible even with no data. */
  const E0 = (typeof esc === "function") ? esc : (s => s);
  const verBadge = `<div class="d-flex justify-content-end align-items-center gap-2">
    <a href="${E0(QBR.WARRANTY_LOOKUP_URL || "")}" target="_blank" rel="noopener" class="small" title="Open Lenovo's warranty lookup (paste a serial number)">🔍 Lenovo warranty lookup</a>
    <span class="badge bg-light text-muted border" title="Delivery package version">v${E0(QBR.INV_VERSION || "")}</span></div>`;
  if (!hasAssets && !hasSup) {
    host.innerHTML = verBadge + `<div class="card-box"><h6>Inventory</h6>
      <p class="text-muted">No inventory data loaded yet. Upload the Lenovo inventory workbook
      (sheets 02 DEVICES, 03 TECH SUPPORT LOGS, 04 RAKSO INV., 06 PIPELINE, 07 PURCHASE ORDER)
      or the printer supplies workbook (sheets Items, Transactions, Inventory Summary)
      using the upload area above — it merges with any workbooks already loaded.</p></div>`;
    return;
  }
  const ui = QBR._invUI;
  if (!hasAssets && hasSup) ui.view = "supplies"; // only supplies present
  if (ui.view === "supplies" && !hasSup) ui.view = "assets";
  if (ui.view === "tickets" && !hasAssets) ui.view = "assets";
  /* Assets | Support tickets | Supplies sub-nav. */
  const vBtn = (v, label) => `<button type="button" class="btn ${ui.view === v ? "btn-primary" : "btn-outline-primary"}" data-sup-view="${v}">${label}</button>`;
  const viewToggle = (hasAssets || hasSup) ? `<div class="btn-group btn-group-sm mb-3" role="group" aria-label="Inventory view">` +
      (hasAssets ? vBtn("assets", "Assets") : "") +
      (hasAssets ? vBtn("tickets", `Support tickets (${fmt(inv.tickets.length)})`) : "") +
      (hasSup ? vBtn("supplies", "Supplies") : "") +
    `</div>` : "";
  const overview = invStockOverview(inv, sup, ui.view === "supplies" ? "supplies" : "assets");
  if (ui.view === "supplies" && hasSup && typeof QBR.supRender === "function") {
    host.innerHTML = overview + viewToggle + `<div id="sup-body"></div>`;
    host.querySelectorAll("[data-sup-view]").forEach(b => b.addEventListener("click", () => {
      QBR._invUI.view = b.dataset.supView; renderInventory();
    }));
    QBR.supRender($("sup-body"));
    return;
  }
  /* ---- Support tickets: its own page under the Inventory tab ---- */
  if (ui.view === "tickets" && hasAssets) {
    const today = invToday(), flags = QBR.invComputeFlags(inv, today);
    const sortCtl = `<div class="btn-group btn-group-sm" role="group" aria-label="Ticket sort order">
        <button type="button" class="btn ${ui.tixSort === "new" ? "btn-outline-primary" : "btn-primary"}" data-tix-sort="old">Oldest first</button>
        <button type="button" class="btn ${ui.tixSort === "new" ? "btn-primary" : "btn-outline-primary"}" data-tix-sort="new">Newest first</button></div>`;
    const flagChip = ui.flag ? `<button type="button" class="btn btn-sm btn-outline-secondary" data-inv-flag="">Flag: ${esc(ui.flag)} ✕</button>` : "";
    host.innerHTML = verBadge + viewToggle +
      `<div class="d-flex flex-wrap gap-2 align-items-center mt-1"><h6 class="mb-0">Support tickets <span class="text-muted">(${fmt(inv.tickets.length)})</span></h6>${flagChip}
        <span class="flex-grow-1"></span>${sortCtl}
        <input id="inv-tix-q" class="form-control form-control-sm" style="max-width:220px" placeholder="Search serial or ticket no…" value="${esc(ui.tixQ || "")}">
        <button type="button" class="btn btn-sm btn-primary" data-inv-form="ticket">New ticket</button>
        <button type="button" class="btn btn-sm btn-outline-primary" data-inv-form="formsimport" title="Import an MS Forms ticket-support export as tickets">Import Forms file</button></div>
      <div id="inv-form-host" class="mt-2"></div>` +
      invTicketsCard(inv, ui, flags, today);
    invBind(host);
    if (ui.form) invShowForm(ui.form);
    try { if (typeof invApplyHash === "function") invApplyHash(); } catch (e) {}
    return;
  }
  const T = QBR.INV_THRESH, today = invToday();
  const openByKey = QBR.invOpenByKey(inv.tickets);
  const flags = QBR.invComputeFlags(inv, today);
  const withStatus = inv.assets.map(a => ({ a, st: QBR.invAssetStatus(a, openByKey) }));

  /* ---- KPIs ---- */
  const deployed = withStatus.filter(x => x.st === "Deployed").length;
  const inRepair = withStatus.filter(x => x.st === "In Repair").length;
  const openT = inv.tickets.filter(t => invTixOpen(t)).length;
  const avg = invAvgResolve(inv.tickets);
  const kpis = `<div class="row row-cols-2 row-cols-md-3 row-cols-xl-5 g-3 kpi-row">` +
    kpi("Total Inventory", fmt(withStatus.length), "blue") +
    kpi("Deployed", withStatus.length ? Math.round(deployed / withStatus.length * 100) + "%" : "—", "green") +
    kpi("In Repair", fmt(inRepair), inRepair ? "orange" : "green") +
    kpi("Open tickets", fmt(openT), openT ? "orange" : "green") +
    kpi("Avg days to resolve", avg == null ? "—" : avg.toFixed(1), "blue") + `</div>`;

  /* ---- flags ---- */
  const flagDefs = [
    ["aging", "Aging tickets", flags.aging.length, `open > ${T.AGING_DAYS} days`, "orange"],
    ["lemons", "Repeat offenders", flags.lemons.length, `≥ ${T.LEMON_TICKETS} tickets per unit`, "red"],
    ["warrantyExpiring", "Warranty expiring", flags.warrantyExpiring.length, `< ${T.WARRANTY_WARN_DAYS} days left`, "orange"],
    ["warrantyExpired", "Warranty expired", flags.warrantyExpired.length, "past end date", "red"],
    ["unmatched", "Unmatched tickets", flags.unmatched.length, "serial not in registry", "blue"],
    ["stalled", "Stalled pipeline", flags.stalled.length, "signed, not delivered", "blue"],
    ["duplicates", "Duplicate serials", flags.duplicates.length, "same serial 2+ times", "red"],
  ];
  const flagCards = `<div class="row row-cols-2 row-cols-md-3 row-cols-xl-6 g-3 mt-1">` + flagDefs.map(([k, lbl, n, sub, tone]) =>
    `<div class="col"><button type="button" class="card-box inv-flag inv-flag-${tone}${ui.flag === k ? " active" : ""}" data-inv-flag="${k}">
      <div class="inv-flag-n">${fmt(n)}</div><div class="inv-flag-l">${lbl}</div><div class="small text-muted">${sub}</div>
    </button></div>`).join("") + `</div>${ui.flag ? `<div class="mt-2"><button type="button" class="btn btn-sm btn-outline-secondary" data-inv-flag="">Clear flag filter ✕</button></div>` : ""}`;

  /* ---- action buttons + forms ---- */
  const actions = `<div class="d-flex flex-wrap gap-2 mt-3">
      <button type="button" class="btn btn-sm btn-primary" data-inv-form="intake">＋ Register assets</button>
      <button type="button" class="btn btn-sm btn-outline-primary" data-inv-form="deploy">Deploy units</button>
      <button type="button" class="btn btn-sm btn-outline-primary" data-inv-form="ticket">New ticket</button>
      ${(typeof QBR.saveButtonHtml === "function") ? QBR.saveButtonHtml("assets", "inv-save") : ""}
    </div>
    <div class="d-flex flex-wrap gap-2 align-items-center mt-2">
      <label class="form-label small mb-0" for="inv-lookup"><strong>Look up asset</strong></label>
      <div style="position:relative"><input id="inv-lookup" class="form-control form-control-sm" style="--w:220px" placeholder="Type or paste a serial…" autocomplete="off" role="combobox" aria-expanded="false">
      <div id="inv-lookup-dd" class="inv-ac-dd d-none" role="listbox"></div></div>
      <button type="button" class="btn btn-sm btn-primary" id="inv-lookup-go">Open</button>
      <span id="inv-lookup-msg" class="small text-muted"></span>
    </div><div id="inv-form-host" class="mt-2"></div>`;

  /* ---- filters ---- */
  const clients = invClients(inv);
  const clientVal = ui.client && ui.client !== "ALL" ? ui.client : "";
  const filters = `<div class="card-box mt-3"><div class="d-flex flex-wrap gap-2 align-items-end">
      <div><label class="form-label small mb-0" for="inv-f-client">Client</label>
        <div style="position:relative"><input id="inv-f-client" class="form-control form-control-sm" style="--w:220px"
        placeholder="All clients — type to search…" value="${esc(clientVal)}" autocomplete="off"
        role="combobox" aria-expanded="false" aria-label="Filter by client, type to search">
        <div id="inv-f-client-dd" class="inv-ac-dd d-none" role="listbox"></div></div></div>
      <div><label class="form-label small mb-0" for="inv-f-type">Type</label>
        <select id="inv-f-type" class="form-select form-select-sm" style="--w:130px">
        ${["ALL", "Laptop", "Desktop", "Monitor"].map(t => `<option${ui.type === t ? " selected" : ""}>${t === "ALL" ? "All types" : t}</option>`).join("")}</select></div>
      <div><label class="form-label small mb-0" for="inv-f-status">Status</label>
        <select id="inv-f-status" class="form-select form-select-sm" style="--w:130px">
        ${["ALL", ...QBR.INV_STATUS].map(t => `<option${ui.status === t ? " selected" : ""}>${t === "ALL" ? "All statuses" : t}</option>`).join("")}</select></div>
      <div><label class="form-label small mb-0" for="inv-f-q">Search serial</label>
        <input id="inv-f-q" class="form-control form-control-sm" style="--w:200px" placeholder="Type a serial…" value="${esc(ui.q)}"></div>
    </div></div>`;

  /* ---- assets table ---- */
  let rows = withStatus;
  /* client filter: partial, case-insensitive match on the client/school text */
  const clientQ = String(ui.client || "").trim().toLowerCase();
  if (clientQ && clientQ !== "all") rows = rows.filter(x => String(x.a.client || "").toLowerCase().includes(clientQ));
  if (ui.type !== "ALL") rows = rows.filter(x => x.a.cat === ui.type);
  if (ui.status !== "ALL") rows = rows.filter(x => x.st === ui.status);
  if (ui.q) { const q = QBR.invSerialKey(ui.q); rows = rows.filter(x => x.a.key.includes(q)); }
  if (ui.flag === "warrantyExpiring") { const s = new Set(flags.warrantyExpiring.map(a => a.key)); rows = rows.filter(x => s.has(x.a.key)); }
  if (ui.flag === "warrantyExpired") { const s = new Set(flags.warrantyExpired.map(a => a.key)); rows = rows.filter(x => s.has(x.a.key)); }
  if (ui.flag === "lemons") { const s = new Set(flags.lemons.map(l => l.key)); rows = rows.filter(x => s.has(x.a.key)); }
  if (ui.flag === "duplicates") { const s = new Set(flags.duplicates.map(d => d.key)); rows = rows.filter(x => s.has(x.a.key)); }
  const tixByKey = new Map();
  inv.tickets.forEach(t => tixByKey.set(t.key, (tixByKey.get(t.key) || 0) + 1));
  const total = rows.length, shown = ui.showAll ? rows : rows.slice(0, 150);
  const assetTblBody = `<div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>Serial</th><th>Client</th><th>Model</th><th>Type</th><th>Status</th><th>Warranty end</th><th class="text-end">Tickets</th></tr></thead><tbody>` +
    (shown.map(x => `<tr class="inv-row" data-inv-key="${esc(x.a.key)}" tabindex="0" role="button" aria-label="Open asset ${esc(x.a.sn)}">
      <td>${invAssetLink(x.a.key, x.a.sn)}</td><td>${esc(x.a.client || "—")}</td><td>${esc(x.a.model || "—")}</td>
      <td>${esc(x.a.cat)}</td><td>${invPill(x.st)}</td><td>${invFmtDate(x.a.wend)}</td>
      <td class="text-end">${fmt(tixByKey.get(x.a.key) || 0)}</td></tr>`).join("") ||
      `<tr><td colspan="7" class="text-muted">No assets match the current filters.</td></tr>`) +
    `</tbody></table></div>` +
    (!ui.showAll && total > 150 ? `<button type="button" class="btn btn-sm btn-outline-secondary" id="inv-showall">Show all ${fmt(total)}</button>` : "");
  const assetTbl = invCollapsible("assets", `Assets <span class="text-muted">(${fmt(total)})</span>`, assetTblBody);

  /* ---- pipeline + PO ---- */
  const plTblBody = `<div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>SQ</th><th>Client</th><th class="text-end">Qty</th><th>Proposal</th><th>Delivery</th><th>Remarks</th></tr></thead><tbody>` +
    (inv.pipeline.map(p => `<tr class="${/signed/i.test(String(p.proposal || "")) && !/delivered/i.test(String(p.delivery || "")) ? "table-warning" : ""}">
      <td><code>${esc(p.sq)}</code></td><td>${esc(p.client || "—")}</td><td class="text-end">${p.qty == null ? "—" : fmt(p.qty)}</td>
      <td>${esc(p.proposal || "—")}</td><td>${esc(p.delivery || "—")}</td><td>${esc(p.rem || "—")}</td></tr>`).join("") ||
      `<tr><td colspan="6" class="text-muted">No pipeline rows.</td></tr>`) + `</tbody></table></div>`;
  const plTbl = invCollapsible("pipeline", `Pipeline <span class="text-muted">(${fmt(inv.pipeline.length)})</span>`, plTblBody);
  const poTblBody = `<div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>PO#</th><th>School</th><th>Units</th><th class="text-end">Qty</th><th>Date</th></tr></thead><tbody>` +
    (inv.pos.map(p => `<tr><td><code>${esc(p.po)}</code></td><td>${esc(p.school || "—")}</td><td>${esc(p.units || "—")}</td>
      <td class="text-end">${p.qty == null ? "—" : fmt(p.qty)}</td><td>${invFmtDate(p.date)}</td></tr>`).join("") ||
      `<tr><td colspan="5" class="text-muted">No purchase orders.</td></tr>`) + `</tbody></table></div>`;
  const poTbl = invCollapsible("pos", `Purchase orders <span class="text-muted">(${fmt(inv.pos.length)})</span>`, poTblBody);

  /* ---- version badge (delivery version, semver — see VERSIONING.md) ---- */
  host.innerHTML = verBadge + overview + viewToggle + kpis + flagCards + actions + filters + assetTbl + plTbl + poTbl;
  invBind(host);
  if (ui.form) invShowForm(ui.form);
  /* deep link (#ticket/… / #asset/…) may have arrived before the model loaded */
  try { if (typeof invApplyHash === "function") invApplyHash(); } catch (e) {}
}

/* ---------- event wiring for the inventory page -------------------------- */
function invBind(host) {
  const ui = QBR._invUI;
  const rerender = () => renderInventory();
  host.querySelectorAll("[data-sup-view]").forEach(b => b.addEventListener("click", () => {
    ui.view = b.dataset.supView; rerender();
  }));
  host.querySelectorAll("[data-inv-typetile]").forEach(b => b.addEventListener("click", () => {
    ui.type = b.dataset.invTypetile; ui.showAll = false; rerender();
    const at = $("inv-asset-card"); if (at) at.scrollIntoView({ block: "start" });
  }));
  host.querySelectorAll("[data-inv-flag]").forEach(b => b.addEventListener("click", () => {
    ui.flag = b.dataset.invFlag || null; rerender();
  }));
  host.querySelectorAll("[data-inv-form]").forEach(b => b.addEventListener("click", () => {
    ui.form = ui.form === b.dataset.invForm ? null : b.dataset.invForm; rerender();
  }));
  const ex = $("inv-save");
  if (ex) ex.addEventListener("click", () => { if (typeof QBR.invSave === "function") QBR.invSave(); else QBR.invExportWorkbook(); });
  if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
  const lgo = $("inv-lookup-go"), lq = $("inv-lookup");
  if (lgo && lq) {
    lgo.addEventListener("click", invLookupSerial);
    lq.addEventListener("input", invLookupSuggest);
    lq.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); invLookupCloseDd(); invLookupSerial(); }
      else if (e.key === "Escape") invLookupCloseDd();
    });
    lq.addEventListener("blur", () => setTimeout(invLookupCloseDd, 150));
  }
  const sa = $("inv-showall");
  if (sa) sa.addEventListener("click", () => { ui.showAll = true; rerender(); });
  const fc = $("inv-f-client"); if (fc) invClientFilterBind(fc);
  const ft = $("inv-f-type");   if (ft) ft.addEventListener("change", () => { ui.type = ft.value; ui.showAll = false; rerender(); });
  const fs = $("inv-f-status"); if (fs) fs.addEventListener("change", () => { ui.status = fs.value; ui.showAll = false; rerender(); });
  const fq = $("inv-f-q");
  if (fq) {
    /* live table filter: re-render on each keystroke (debounced), then put
     * the cursor back so typing isn't interrupted by the re-render. */
    let deb = 0;
    fq.addEventListener("input", () => {
      clearTimeout(deb);
      deb = setTimeout(() => {
        ui.q = fq.value.trim(); ui.showAll = false;
        renderInventory();
        const nfq = $("inv-f-q");
        if (nfq) { nfq.focus(); try { nfq.setSelectionRange(nfq.value.length, nfq.value.length); } catch (e) {} }
      }, 200);
    });
    fq.addEventListener("keydown", e => {
      if (e.key === "Enter") {
        e.preventDefault(); clearTimeout(deb);
        ui.q = fq.value.trim(); ui.showAll = false; renderInventory();
      }
    });
  }
  const tq = $("inv-tix-q");
  if (tq) tq.addEventListener("change", () => { ui.tixQ = tq.value.trim(); rerender(); });
  host.querySelectorAll("[data-tix-sort]").forEach(b => b.addEventListener("click", () => {
    ui.tixSort = b.dataset.tixSort; rerender();
  }));
  host.querySelectorAll("[data-inv-col]").forEach(b => b.addEventListener("click", () => {
    const id = b.dataset.invCol, body = $("inv-col-" + id);
    if (!body) return;
    const nowHidden = body.classList.toggle("d-none");
    b.textContent = nowHidden ? "▸" : "▾";
    b.setAttribute("aria-label", nowHidden ? "Expand section" : "Collapse section");
    ui.collapsed = ui.collapsed || {};
    ui.collapsed[id] = nowHidden;
  }));
  host.querySelectorAll(".inv-row[data-inv-key]").forEach(r => {
    const go = e => {
      if (e.target && e.target.closest && e.target.closest("a")) return; // real link (new tab) handles it
      e.preventDefault(); openAsset360(r.dataset.invKey);
    };
    r.addEventListener("click", go);
    r.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") go(e); });
  });
  host.querySelectorAll(".inv-link[data-inv-key]").forEach(a => a.addEventListener("click", e => {
    e.preventDefault(); e.stopPropagation(); openAsset360(a.dataset.invKey);
  }));
  host.querySelectorAll("[data-inv-resolve]").forEach(b => b.addEventListener("click", () => {
    const t = inv.tickets.find(x => x._n === Number(b.dataset.invResolve));
    if (!t) return;
    const note = window.prompt("Resolution note (optional):", "");
    if (note === null) return;
    if (QBR.invResolveTicket(t.key, t._n, note.trim())) renderAll();
  }));
}

/* ---------- write-back forms --------------------------------------------- */
function invShowForm(which) {
  const host = $("inv-form-host");
  if (!host) return;
  const ui = QBR._invUI, inv = invModel();
  const clients = inv ? invClients(inv) : [];
  const dl = `<datalist id="dl-inv-clients">${clients.map(c => `<option value="${esc(c)}"></option>`).join("")}</datalist>`;
  const close = `<button type="button" class="btn btn-sm btn-outline-secondary" data-inv-form="">Close ✕</button>`;
  if (which === "intake") {
    host.innerHTML = `<div class="card-box"><h6>Register assets (intake)</h6>
      <div class="row g-2">
        <div class="col-md-4"><label class="form-label small">Serials — one per line *</label>
          <textarea id="in-serials" class="form-control form-control-sm" rows="5" placeholder="PF4ABC123&#10;SGMDEF456"></textarea></div>
        <div class="col-md-8"><div class="row g-2">
          <div class="col-md-6"><label class="form-label small">Model</label>
            <select id="in-model" class="form-select form-select-sm">
              <option>ThinkPad E14 Gen 5</option><option>ThinkCentre M75t Gen 2</option><option>C22-20 21.5"</option><option>Other</option></select></div>
          <div class="col-md-6"><label class="form-label small">Category</label>
            <select id="in-cat" class="form-select form-select-sm"><option>Laptop</option><option>Desktop</option><option>Monitor</option></select></div>
          <div class="col-md-6"><label class="form-label small">Supplier</label><input id="in-supplier" class="form-control form-control-sm" value="Twireless"></div>
          <div class="col-md-3"><label class="form-label small">Warranty years</label>
            <select id="in-wy" class="form-select form-select-sm"><option>1</option><option>2</option><option selected>3</option></select></div>
          <div class="col-md-3"><label class="form-label small">DR #</label><input id="in-dr" class="form-control form-control-sm"></div>
          <div class="col-md-6"><label class="form-label small">Client (leave blank = in stock)</label>
            <input id="in-client" class="form-control form-control-sm" list="dl-inv-clients">${dl}</div>
          <div class="col-md-6"><label class="form-label small">Warranty start</label>
            <input id="in-ws" type="date" class="form-control form-control-sm" value="${invFmtDate(invToday()) === "—" ? "" : (() => { const d = invToday(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); })()}"></div>
          <div class="col-md-6"><label class="form-label small" for="in-we">Warranty end (if known)</label>
            <input id="in-we" type="date" class="form-control form-control-sm" title="Leave blank to count it from warranty start + years"></div>
        </div></div>
      </div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="in-go">Register</button><button type="button" class="btn btn-sm btn-outline-primary" id="in-scan">Scan box label</button>${close}</div>
      <div id="in-msg" class="small mt-1" aria-live="polite"></div></div>`;
    $("in-go").addEventListener("click", () => {
      const serials = $("in-serials").value.split(/\n+/).map(s => s.trim()).filter(Boolean);
      if (!serials.length) { $("in-msg").textContent = "Enter at least one serial number."; return; }
      const had = new Set((invModel() || { assets: [] }).assets.map(a => a.key));
      const n = QBR.invIntake(serials.map(sn => ({ sn, model: $("in-model").value, cat: $("in-cat").value,
        supplier: $("in-supplier").value.trim() || null, wyears: Number($("in-wy").value),
        dr: $("in-dr").value.trim() || null, client: $("in-client").value.trim() || null, wstart: $("in-ws").value || null })));
      const we = $("in-we") ? $("in-we").value : "";
      if (we && n) serials.filter(sn => !had.has(QBR.invSerialKey(sn))).forEach(sn => QBR.invSetWarrantyEnd(sn, we));   // new units only
      $("in-msg").innerHTML = `<span class="text-success">${n} asset(s) registered.</span> <span class="text-muted">Duplicates skipped.</span>`;
      setTimeout(() => { ui.form = null; renderAll(); }, 900);
    });
    const isc = $("in-scan");
    if (isc) isc.addEventListener("click", () => {
      ui.form = null;
      if (typeof QBR.scanSetTarget === "function") QBR.scanSetTarget("intake");
    });
  } else if (which === "deploy") {
    host.innerHTML = `<div class="card-box"><h6>Deploy units</h6>
      <div class="row g-2">
        <div class="col-md-4"><label class="form-label small">Serials — one per line *</label>
          <textarea id="dp-serials" class="form-control form-control-sm" rows="4" placeholder="PF4ABC123"></textarea></div>
        <div class="col-md-8"><div class="row g-2">
          <div class="col-md-6"><label class="form-label small">Client *</label>
            <input id="dp-client" class="form-control form-control-sm" list="dl-inv-clients">${dl}</div>
          <div class="col-md-3"><label class="form-label small">Date delivered</label><input id="dp-date" type="date" class="form-control form-control-sm"></div>
          <div class="col-md-3"><label class="form-label small">SQ (optional)</label><input id="dp-sq" class="form-control form-control-sm"></div>
        </div></div>
      </div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="dp-go">Deploy</button><button type="button" class="btn btn-sm btn-outline-primary" id="dp-scan">Scan unit</button>${close}</div>
      <div id="dp-msg" class="small mt-1" aria-live="polite"></div></div>`;
    $("dp-go").addEventListener("click", () => {
      const serials = $("dp-serials").value.split(/\n+/).map(s => s.trim()).filter(Boolean);
      const client = $("dp-client").value.trim();
      if (!serials.length || !client) { $("dp-msg").textContent = "Serials and client are required."; return; }
      const n = QBR.invDeploy(serials, client, $("dp-date").value || null, $("dp-sq").value.trim() || null);
      $("dp-msg").innerHTML = `<span class="text-success">${n} unit(s) deployed to ${esc(client)}.</span>`;
      setTimeout(() => { ui.form = null; renderAll(); }, 900);
    });
    const dsc = $("dp-scan");
    if (dsc) dsc.addEventListener("click", () => {
      ui.form = null;
      if (typeof QBR.scanSetTarget === "function") QBR.scanSetTarget("deploy");
    });
  } else if (which === "ticket") {
    const pre = ui.prefillSn ? String(ui.prefillSn) : "";
    host.innerHTML = `<div class="card-box"><h6>New support ticket</h6>
      <div class="row g-2">
        <div class="col-md-3"><label class="form-label small">Serial *</label>
          <input id="tk-sn" class="form-control form-control-sm" value="${esc(pre)}" placeholder="PF4ABC123" autocomplete="off">
          <div id="tk-sn-hint" class="small mt-1" aria-live="polite"></div></div>
        <div class="col-md-3"><label class="form-label small">Requested by *</label>
          <input id="tk-req" class="form-control form-control-sm" placeholder="e.g. Juan Dela Cruz"></div>
        <div class="col-md-3"><label class="form-label small">Category</label>
          <select id="tk-cat" class="form-select form-select-sm"><option>Hardware</option><option>Software</option><option>Network</option></select></div>
        <div class="col-md-3"><label class="form-label small">Priority</label>
          <select id="tk-prio" class="form-select form-select-sm"><option value="">—</option><option>Low</option><option selected>Medium</option><option>High</option></select></div>
        <div class="col-md-3"><label class="form-label small">Person in-charge</label><input id="tk-pic" class="form-control form-control-sm"></div>
        <div class="col-md-3"><label class="form-label small">Date reported to EdTech</label><input id="tk-date" type="date" class="form-control form-control-sm"></div>
        <div class="col-md-3"><label class="form-label small">Date reported to Lenovo</label><input id="tk-lendate" type="date" class="form-control form-control-sm"></div>
        <div class="col-md-6"><label class="form-label small">Issue *</label><input id="tk-issue" class="form-control form-control-sm" placeholder="e.g. No power"></div>
        <div class="col-md-6"><label class="form-label small">Initial findings</label><input id="tk-act" class="form-control form-control-sm"></div>
      </div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="tk-go">Open ticket</button>${close}</div>
      <div id="tk-msg" class="small mt-1" aria-live="polite"></div></div>`;
    ui.prefillSn = null;
    // 2026-10-05: live check — is this serial registered? suggest close matches.
    let tkAck = "";
    const tkHint = () => {
      const h = $("tk-sn-hint"), raw = $("tk-sn").value.trim(); if (!h) return;
      tkAck = "";
      if (!raw) { h.innerHTML = ""; return; }
      const inv0 = invModel(), hit = inv0 ? inv0.assets.find(x => x.key === QBR.invSerialKey(raw)) : null;
      if (hit) { h.innerHTML = `<span class="text-success">✓ In inventory</span><span class="text-muted"> · ${esc(hit.client || "in stock")}${hit.model ? " · " + esc(hit.model) : ""}</span>`; return; }
      const sg = QBR.invSuggestSerials(raw);
      h.innerHTML = `<span class="inv-notinv">Not in inventory</span>` +
        (sg.length ? ` Did you mean ${sg.map(a => `<button type="button" class="btn btn-sm btn-link p-0" data-tk-use="${escAttr(a.sn)}"><code>${esc(a.sn)}</code></button>`).join(", ")}?` : "") +
        ` <button type="button" class="btn btn-sm btn-link p-0" data-inv-register="${escAttr(raw)}">Register this unit</button>`;
      h.querySelectorAll("[data-tk-use]").forEach(b => b.addEventListener("click", () => { $("tk-sn").value = b.dataset.tkUse; tkHint(); }));
    };
    $("tk-sn").addEventListener("input", tkHint);
    tkHint();
    $("tk-go").addEventListener("click", () => {
      const sn = $("tk-sn").value.trim(), issue = $("tk-issue").value.trim(), req = $("tk-req").value.trim();
      if (!sn || !issue || !req) { $("tk-msg").textContent = "Serial, requester and issue are required."; return; }
      const invC = invModel(), known = invC && invC.assets.some(x => x.key === QBR.invSerialKey(sn));
      if (!known && tkAck !== sn) {
        tkAck = sn;
        const sg = QBR.invSuggestSerials(sn);
        $("tk-msg").innerHTML = `<span class="text-warning">${esc(sn)} isn't in the inventory${sg.length ? ` — did you mean <code>${esc(sg[0].sn)}</code>?` : "."}
          Click <b>Open ticket</b> again to open it with the serial as typed.</span>`;
        return;
      }
      const inv2 = invModel(), a = inv2 ? inv2.assets.find(x => x.key === QBR.invSerialKey(sn)) : null;
      const tno = QBR.invAddTicket({ sn, client: a ? a.client : null, model: a ? a.model : null,
        cat: $("tk-cat").value, priority: $("tk-prio").value || null, requester: req,
        pic: $("tk-pic").value.trim() || null,
        repEdtech: $("tk-date").value || null, repLenovo: $("tk-lendate").value || null,
        issue, act: $("tk-act").value.trim() || null });
      $("tk-msg").innerHTML = tno
        ? `<span class="text-success">Ticket ${esc(tno)} opened for ${esc(sn)}.</span>`
        : `<span class="text-danger">Could not open ticket.</span>`;
      if (tno) setTimeout(() => { ui.form = null; renderAll(); }, 900);
    });
  } else if (which === "formsimport") {
    host.innerHTML = `<div class="card-box"><h6>Import Forms responses</h6>
      <p class="small text-muted mb-2">Upload the MS Forms ticket-support export (.xlsx). Each new response becomes a support ticket (client auto-matched from the inventory by serial). Responses already imported — same response ID or identical row — are skipped.</p>
      <input type="file" id="fi-file" accept=".xlsx,.xls" class="form-control form-control-sm" style="max-width:340px">
      <div id="fi-preview" class="mt-2"></div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="fi-go" disabled>Import</button>${close}</div>
      <div id="fi-msg" class="small mt-1" aria-live="polite"></div></div>`;
    let parsed = null, classified = null;
    const E = (typeof esc === "function") ? esc : (s => String(s == null ? "" : s));
    const renderPreview = () => {
      const box = $("fi-preview"); if (!box) return;
      const go = $("fi-go");
      if (!classified || !classified.length) {
        box.innerHTML = `<p class="text-muted small mb-0">No response rows found in this file.</p>`;
        go.disabled = true; go.textContent = "Import";
        return;
      }
      const nNew = classified.filter(p => p.status === "new").length;
      const badge = s => s === "new" ? `<span class="badge bg-success">Will import</span>`
        : s === "duplicate" ? `<span class="badge bg-secondary">Duplicate</span>`
        : `<span class="badge bg-warning text-dark">Skipped</span>`;
      box.innerHTML = `<div class="table-responsive"><table class="table table-sm"><thead><tr>
        <th>Row</th><th>Forms ID</th><th>Serial</th><th>Model</th><th>Requested by</th><th>Priority</th><th>Result</th></tr></thead><tbody>` +
        classified.map(p => `<tr><td>${p.row}</td><td>${E(p.formsId || "—")}</td><td><code>${E(p.sn || "—")}</code></td>` +
          `<td>${E(p.ticket.model || "—")}</td><td>${E(p.ticket.requester || "—")}</td><td>${E(p.ticket.priority || "—")}</td>` +
          `<td>${badge(p.status)}${p.reason ? ` <span class="text-muted small">${E(p.reason)}</span>` : ""}</td></tr>`).join("") +
        `</tbody></table></div>`;
      go.disabled = !nNew;
      go.textContent = nNew ? `Import ${nNew} new ticket${nNew === 1 ? "" : "s"}` : "Nothing new to import";
    };
    $("fi-file").addEventListener("change", async () => {
      const f = $("fi-file").files[0];
      $("fi-msg").textContent = "";
      if (!f) return;
      try {
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
        parsed = QBR.formsParseRows(rows);
        classified = QBR.formsClassify(parsed);
        renderPreview();
      } catch (e) { $("fi-msg").textContent = "Could not read that file: " + (e && e.message); }
    });
    $("fi-go").addEventListener("click", () => {
      if (!classified) return;
      const res = QBR.formsCommit(classified);
      $("fi-msg").innerHTML = `<span class="text-success">${res.imported.length} ticket(s) imported.</span>` +
        (res.skipped ? ` <span class="text-muted">${res.skipped} skipped.</span>` : "");
      classified = QBR.formsClassify(parsed); // refresh statuses
      renderPreview();
      setTimeout(() => { ui.form = null; renderAll(); }, 1400);
    });
  }
  /* Close button only — the action-bar buttons are already wired by invBind().
   * (Binding [data-inv-form] here too double-fires: toggle opens the form, then
   * this handler instantly closes it, which made "New ticket" look unresponsive
   * when another form was already open.) */
  host.querySelectorAll("[data-inv-form=\"\"]").forEach(b => b.addEventListener("click", () => { ui.form = null; renderInventory(); }));
}

/* ============================ ASSET 360 PAGE ============================= */
/* Activate a panel directly (mirrors the [data-tab] click handler in app.js);
 * used for Asset 360, which has no sidebar button. */
function invActivateTab(tabId) {
  document.querySelectorAll("[data-tab]").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".dash-panel").forEach(p => p.classList.add("d-none"));
  const panel = document.getElementById(tabId);
  if (panel) panel.classList.remove("d-none");
  if (typeof APP !== "undefined") APP.activeTab = tabId;
  if (typeof applyFilterVisibility === "function") applyFilterVisibility(tabId);
}
function openAsset360(key) {
  QBR._invA360 = key;
  invActivateTab("dash-asset360");
  renderAsset360Panel();
  try { history.replaceState(null, "", "#asset/" + encodeURIComponent(key)); } catch (e) {}
  window.scrollTo(0, 0);
}

/* Edit-details form on Asset 360: prefilled inline card; Save diffs each field
 * against the current values and records only changed fields via
 * QBR.invUpdateAsset (journaled → patch-in-place save). Serial is the join
 * key and can't be changed here. */
function invToggleAssetEdit(a) {
  const host = $("a360-edit-host"); if (!host || !a) return;
  if (host.innerHTML) { host.innerHTML = ""; return; }
  const E = (typeof esc === "function") ? esc : (s => String(s == null ? "" : s));
  const dstr = d => { // Date → "YYYY-MM-DD" for <input type="date">
    if (!(d instanceof Date) || isNaN(d)) return "";
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  };
  const T = (id, label, val, extra) => `<div class="col-md-4"><label class="form-label small" for="${id}">${label}</label>\n    <input id="${id}" class="form-control form-control-sm" value="${E(val == null ? "" : val)}"${extra || ""}></div>`;
  host.innerHTML = `<div class="card-box"><h6>Edit details — <code>${E(a.sn)}</code>\n    <span class="text-muted small fw-normal">(serial number can't be changed)</span></h6>\n    <div class="row g-2">\n      ${T("ae-client", "Client / Organization", a.client, ` list="dl-ae-clients"`)}
      ${T("ae-model", "Model", a.model)}
      <div class="col-md-4"><label class="form-label small" for="ae-cat">Category</label>\n        <select id="ae-cat" class="form-select form-select-sm">${["Laptop", "Desktop", "Monitor"].map(c => `<option${a.cat === c ? " selected" : ""}>${c}</option>`).join("")}</select></div>
      ${T("ae-desc", "Description", a.desc)}
      ${T("ae-brand", "Brand", a.brand)}
      ${T("ae-supplier", "Supplier", a.supplier)}
      ${T("ae-dr", "DR #", a.dr)}
      ${T("ae-delivered", "Date delivered", dstr(a.delivered), ` type="date"`)}
      ${T("ae-wstart", "Warranty start", dstr(a.wstart), ` type="date"`)}
      ${T("ae-wend", "Warranty end", dstr(a.wend), ` type="date"`)}
      ${T("ae-wyears", "Warranty years", a.wyears == null ? "" : a.wyears, ` type="number" min="0" max="10"`)}
      ${T("ae-cond", "Condition", a.cond)}
      ${T("ae-contact", "Contact person", a.contact)}
      ${T("ae-addr", "Address", a.addr)}
      ${T("ae-phone", "Contact details", a.phone)}
    </div>\n    <datalist id="dl-ae-clients">${invClients(invModel() || { assets: [] }).map(c => `<option value="${E(c)}"></option>`).join("")}</datalist>\n    <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="ae-save">Save changes</button>\n    <button type="button" class="btn btn-sm btn-outline-secondary" id="ae-cancel">Cancel</button></div>\n    <div id="ae-msg" class="small mt-1" aria-live="polite"></div></div>`;
  $("ae-cancel").addEventListener("click", () => { host.innerHTML = ""; });
  $("ae-save").addEventListener("click", () => {
    const val = id => $(id).value.trim();
    const patch = {};
    const cur = { client: a.client, model: a.model, desc: a.desc, brand: a.brand, supplier: a.supplier, dr: a.dr, cond: a.cond, contact: a.contact, addr: a.addr, phone: a.phone };
    [["client", "ae-client"], ["model", "ae-model"], ["desc", "ae-desc"], ["brand", "ae-brand"],
     ["supplier", "ae-supplier"], ["dr", "ae-dr"], ["cond", "ae-cond"], ["contact", "ae-contact"],
     ["addr", "ae-addr"], ["phone", "ae-phone"]].forEach(([f, id]) => {
      const v = val(id), c = cur[f] == null ? "" : String(cur[f]);
      if (v !== c) patch[f] = v;
    });
    const cat = $("ae-cat").value;
    if (cat !== a.cat) patch.cat = cat;
    [["delivered", "ae-delivered"], ["wstart", "ae-wstart"], ["wend", "ae-wend"]].forEach(([f, id]) => {
      if ($(id).value !== dstr(a[f])) patch[f] = $(id).value;
    });
    const wy = val("ae-wyears"), wyc = a.wyears == null ? "" : String(a.wyears);
    if (wy !== wyc) patch.wyears = wy;
    if (!Object.keys(patch).length) { $("ae-msg").textContent = "No changes to save."; return; }
    if (QBR.invUpdateAsset(a.key, patch)) { host.innerHTML = ""; renderAsset360Panel(); }
    else $("ae-msg").textContent = "Could not save — please try again.";
  });
}

/* Quick serial lookup from the Inventory page: exact match jumps straight to
 * Asset 360, several matches narrow the Assets table, no match offers Intake
 * with the serial prefilled. */
function invLookupCloseDd() {
  const dd = $("inv-lookup-dd");
  if (dd) { dd.classList.add("d-none"); const lq = $("inv-lookup"); if (lq) lq.setAttribute("aria-expanded", "false"); }
}
/* Autocomplete for the quick-lookup box: while typing, show up to 8 matching
 * serials (partial match) with client/model context. Clicking one opens
 * Asset 360 directly; Enter keeps the existing invLookupSerial behaviour. */
function invLookupSuggest() {
  const lq = $("inv-lookup"); if (!lq) return;
  const dd = $("inv-lookup-dd"); if (!dd) return;
  const inv = invModel();
  const q = QBR.invSerialKey(lq.value.trim());
  if (!inv || q.length < 2) { invLookupCloseDd(); return; }
  const hits = inv.assets.filter(a => a.key.indexOf(q) !== -1).slice(0, 8);
  if (!hits.length) { invLookupCloseDd(); return; }
  dd.innerHTML = hits.map(a =>
    `<button type="button" class="inv-ac-item" data-key="${esc(a.key)}" role="option">` +
    `<code>${esc(a.sn)}</code>` +
    `<span class="text-muted small"> ${esc(a.client || "In stock")}${a.model ? " · " + esc(a.model) : ""}</span></button>`
  ).join("");
  dd.classList.remove("d-none");
  lq.setAttribute("aria-expanded", "true");
  dd.querySelectorAll(".inv-ac-item").forEach(b => b.addEventListener("mousedown", e => {
    e.preventDefault(); /* before blur closes the dropdown */
    invLookupCloseDd();
    openAsset360(b.dataset.key);
  }));
}
/* Searchable client filter: a text input whose dropdown lists clients
 * matching the typed text (partial, case-insensitive). Picking a suggestion,
 * pressing Enter, or leaving the field applies the filter — an empty field
 * means all clients. */
function invClientFilterBind(input) {
  const ui = QBR._invUI;
  const dd = () => $("inv-f-client-dd");
  const clients = invClients(invModel() || { assets: [] });
  const close = () => { const d = dd(); if (d) { d.classList.add("d-none"); input.setAttribute("aria-expanded", "false"); } };
  const apply = () => {
    const v = input.value.trim();
    ui.client = v ? v : "ALL";
    ui.showAll = false;
    renderInventory();
  };
  const show = () => {
    const d = dd(); if (!d) return;
    const q = input.value.trim().toLowerCase();
    const hits = clients.filter(c => !q || c.toLowerCase().indexOf(q) !== -1).slice(0, 12);
    d.innerHTML =
      `<button type="button" class="inv-ac-item" data-v="" role="option"><em>All clients</em></button>` +
      hits.map(c => `<button type="button" class="inv-ac-item" data-v="${esc(c)}" role="option">${esc(c)}</button>`).join("");
    d.classList.remove("d-none");
    input.setAttribute("aria-expanded", "true");
    d.querySelectorAll(".inv-ac-item").forEach(b => b.addEventListener("mousedown", e => {
      e.preventDefault(); /* before blur closes the dropdown */
      input.value = b.dataset.v;
      close();
      apply();
    }));
  };
  input.addEventListener("focus", show);
  input.addEventListener("input", show);
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); close(); apply(); }
    else if (e.key === "Escape") { close(); input.blur(); }
  });
  /* typed-but-unpicked text still applies as a partial filter on blur */
  input.addEventListener("blur", () => setTimeout(() => {
    const d = dd();
    if (d && !d.classList.contains("d-none")) { close(); apply(); }
  }, 150));
}
function invLookupSerial() {
  const inv = invModel();
  const box = $("inv-lookup"), msg = $("inv-lookup-msg");
  if (!inv || !box) return;
  const raw = box.value.trim();
  const say = t => { if (msg) msg.innerHTML = t; };
  if (!raw) { say("Type a serial number first."); box.focus(); return; }
  const q = QBR.invSerialKey(raw);
  const hit = inv.assets.find(x => x.key === q);
  if (hit) { openAsset360(hit.key); return; }
  const subs = inv.assets.filter(x => x.key.indexOf(q) !== -1);
  if (subs.length === 1) { openAsset360(subs[0].key); return; }
  if (subs.length > 1) {
    QBR._invUI.q = raw; QBR._invUI.showAll = false; QBR._invUI.form = null;
    renderInventory();
    const fq = $("inv-f-q");
    if (fq) fq.scrollIntoView({ block: "center" });
    return;
  }
  say(`No asset matches &ldquo;${esc(raw)}&rdquo;. <button type="button" class="btn btn-sm btn-link p-0" id="inv-lookup-add">Register it via intake</button>`);
  const add = $("inv-lookup-add");
  if (add) add.addEventListener("click", () => {
    QBR._invUI.form = "intake";
    renderInventory();
    const ta = $("in-serials");
    if (ta) { ta.value = raw; ta.focus(); }
    const fh = $("inv-form-host");
    if (fh) fh.scrollIntoView({ block: "center" });
  });
}

function renderAsset360Panel() {
  const host = $("a360-body");
  if (!host) return;
  const inv = invModel(), key = QBR._invA360;
  if (!inv || !inv.assets.length) {
    host.innerHTML = `<div class="card-box"><p class="text-muted">Load an inventory workbook to browse assets.</p></div>`;
    return;
  }
  const a = inv.assets.find(x => x.key === key);
  if (!a && key) {
    const sugg = QBR.invSuggestSerials(key);
    host.innerHTML = `<div class="card-box inv-notfound"><h6>Unit not in inventory</h6>
      <p class="mb-1">Serial <code>${esc(key)}</code> isn't in the inventory (sheet <b>02 DEVICES</b>), so there's no unit page to show.</p>
      <p class="small text-muted mb-2">This usually means a ticket was opened for a unit that was never registered, or the serial was typed differently
        (for example with a hyphen or space).</p>
      ${sugg.length ? `<p class="mb-2">Did you mean ${sugg.map(s => `<a href="#asset/${encodeURIComponent(s.key)}" class="inv-link"><code>${esc(s.sn)}</code></a>${s.client ? ` <span class="text-muted small">(${esc(s.client)})</span>` : ""}`).join(", ")}?</p>` : ""}
      <div class="d-flex flex-wrap gap-2">
        <button type="button" class="btn btn-sm btn-primary" data-inv-register="${escAttr(key)}">Register this unit</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="a360-nf-back">Back to Inventory</button></div></div>`;
    $("a360-nf-back").addEventListener("click", () => {
      QBR._invA360 = null;
      try { history.replaceState(null, "", location.pathname + location.search + "#inventory"); } catch (e) {}
      goToTab("dash-inventory");
    });
    return;
  }
  if (!a) {
    host.innerHTML = `<div class="card-box"><h6>Asset 360</h6>
      <p class="text-muted">Pick an asset from the Inventory table, or search by serial below.</p>
      <div class="d-flex gap-2"><input id="a360-q" class="form-control form-control-sm" style="--w:260px" placeholder="Serial number…">
      <button type="button" class="btn btn-sm btn-primary" id="a360-go">Open</button></div>
      <div id="a360-msg" class="small text-muted mt-1"></div></div>`;
    $("a360-go").addEventListener("click", () => {
      const k = QBR.invSerialKey($("a360-q").value);
      const hit = inv.assets.find(x => x.key === k || x.key.includes(k));
      if (hit) openAsset360(hit.key);
      else $("a360-msg").textContent = "No asset matches that serial.";
    });
    return;
  }
  const today = invToday(), openByKey = QBR.invOpenByKey(inv.tickets);
  const st = QBR.invAssetStatus(a, openByKey);
  const tix = inv.tickets.filter(t => t.key === a.key).sort((x, y) => (y.repEdtech || 0) - (x.repEdtech || 0));
  const open = tix.filter(t => !t.completed && !/completed|resolved|closed/i.test(String(t.status || "")));
  const deps = inv.deployments.filter(d => d.key === a.key);
  const wLeft = a.wend ? invDaysBetween(today, a.wend) : null;

  const head = `<div class="card-box" style="border-left:4px solid var(--brand)">
      <div class="d-flex flex-wrap justify-content-between align-items-start gap-2">
        <div><div class="small text-muted">${esc(a.client || "No client on record")}</div>
          <h4 class="mb-1"><code>${esc(a.sn)}</code></h4>
          <div class="small text-muted">${esc(a.brand || "")} ${esc(a.model || "")} · ${esc(a.cat)}</div></div>
        <div class="text-end"><div class="mb-1">${invPill(st)}</div>
          <div class="d-flex gap-2 justify-content-end">
            <select id="a360-status" class="form-select form-select-sm" style="--w:130px" aria-label="Override status">
              ${QBR.INV_STATUS.map(s => `<option${st === s ? " selected" : ""}>${s}</option>`).join("")}</select>
            <button type="button" class="btn btn-sm btn-outline-primary" id="a360-setstatus">Set</button>
            <button type="button" class="btn btn-sm btn-outline-secondary" id="a360-edit" title="Edit this unit's details (warranty dates, client, model…)">✏️ Edit</button>
          </div></div>
      </div></div>`;

  const kpis = `<div class="row row-cols-2 row-cols-md-4 g-3 kpi-row mt-1">` +
    kpi("Tickets", fmt(tix.length), tix.length ? "blue" : "blue") +
    kpi("Open", fmt(open.length), open.length ? "orange" : "green") +
    kpi("Warranty", wLeft == null ? "—" : (wLeft < 0 ? "Expired" : wLeft + " days"), wLeft == null ? "blue" : (wLeft < 0 ? "red" : (wLeft <= QBR.INV_THRESH.WARRANTY_WARN_DAYS ? "orange" : "green"))) +
    kpi("Deployments", fmt(deps.length), "blue") + `</div>`;

  const facts = [["Serial", `<code>${esc(a.sn)}</code>`], ["Client", esc(a.client || "—")], ["Model", esc(a.model || "—")],
    ["Description", esc(a.desc || "—")], ["Category", esc(a.cat)], ["Brand", esc(a.brand || "—")],
    ["Supplier", esc(a.supplier || "—")], ["DR #", esc(a.dr || "—")],
    ["Date delivered", invFmtDate(a.delivered)],
    ["Warranty", (a.wstart || a.wend ? `${invFmtDate(a.wstart)} → ${invFmtDate(a.wend)}${a.wyears ? ` (${a.wyears}y)` : ""}` : "—") +
      ` <a href="${esc(QBR.invWarrantyUrl(a.sn, a.model))}" target="_blank" rel="noopener" class="small" title="Open Lenovo's warranty lookup (paste the serial number)">Check on Lenovo ↗</a>`],
    ["Condition", esc(a.cond || "—")], ["Contact", esc(a.contact || "—")],
    ["Address", esc(a.addr || "—")], ["Contact details", esc(a.phone || "—")]];
  const factCard = `<div class="card-box mt-3"><h6>Asset facts</h6>
    <div class="s360-facts">${facts.map(([k, v]) => `<div class="s360-fact"><div class="s360-fact-k">${k}</div><div class="s360-fact-v">${v}</div></div>`).join("")}</div></div>`;

  const depCard = `<div class="card-box mt-3"><h6>Deployment record</h6>` + (deps.length
    ? `<table class="table table-sm inv-tbl"><thead><tr><th>SQ</th><th>Client</th><th>Date requested</th><th>Date delivered</th><th>Remarks</th></tr></thead><tbody>` +
      deps.map(d => `<tr><td><code>${esc(d.sq || "—")}</code></td><td>${esc(d.client || "—")}</td><td>${invFmtDate(d.req)}</td><td>${invFmtDate(d.delivered)}</td><td>${esc(d.remarks || "—")}</td></tr>`).join("") + `</tbody></table>`
    : `<p class="text-muted s360-empty">No deployment record for this unit.</p>`) + `</div>`;

  const tixCard = `<div class="card-box mt-3"><div class="d-flex justify-content-between align-items-center">
      <h6 class="mb-0">Ticket history <span class="text-muted">(${fmt(tix.length)})</span></h6>
      <button type="button" class="btn btn-sm btn-outline-primary" id="a360-newticket">New ticket</button></div>` +
    (tix.length ? `<ol class="inv-timeline">` + tix.map((t, i) => {
      const isOpen = invTixOpen(t), st = invTixStatus(t);
      return `<li class="inv-tl-item"><div class="d-flex justify-content-between gap-2">
        <b>${invTixLink(t.tno)} — ${esc(t.issue || "—")}</b>${invPill(st)}</div>
        <div class="small text-muted">Reported ${invFmtDate(t.repEdtech)}${t.repLenovo ? ` · Lenovo ${invFmtDate(t.repLenovo)}` : ""}${t.completed ? ` · Resolved ${invFmtDate(t.completed)}` : ""} · ${esc(t.pic || "—")}${t.priority ? ` · Priority: ${esc(t.priority)}` : ""}</div>
        ${t.act ? `<div class="small mt-1">${esc(t.act)}</div>` : ""}
        ${t.svcAddr ? `<div class="small text-muted">${esc(t.svcAddr)}</div>` : ""}
        ${isOpen ? `<div class="mt-1"><button type="button" class="btn btn-sm btn-outline-success" data-a360-resolve="${t._n}">Mark resolved</button></div>` : ""}</li>`;
    }).join("") + `</ol>` : `<p class="text-muted s360-empty">No tickets for this unit.</p>`) + `</div>`;

  /* rule-based recommended actions */
  const acts = [];
  if (wLeft != null && wLeft < 0) acts.push(["High", "Warranty expired — plan replacement or extended coverage."]);
  else if (wLeft != null && wLeft <= QBR.INV_THRESH.WARRANTY_WARN_DAYS) acts.push(["Medium", `Warranty expires in ${wLeft} days — renew or replace before it lapses.`]);
  open.forEach(t => { const d = t.repEdtech ? invDaysBetween(t.repEdtech, today) : null;
    acts.push([d != null && d > QBR.INV_THRESH.AGING_DAYS ? "High" : "Medium",
      `Ticket open${d != null ? ` ${d} days` : ""}: "${t.issue || "—"}" — follow up${t.repLenovo ? " with Lenovo" : ""}.`]); });
  if (tix.length >= QBR.INV_THRESH.LEMON_TICKETS) acts.push(["High", `Repeat offender: ${tix.length} tickets on this unit — consider replacement.`]);
  if (!a.client) acts.push(["Low", "Unit is in stock with no client — assign it or keep as spare."]);
  if (/repair/i.test(String(a.cond || "")) && !open.length) acts.push(["Medium", "Condition is marked for repair but no ticket is open — open one or clear the condition."]);
  const actCard = `<div class="card-box mt-3"><h6>Recommended actions</h6>` + (acts.length
    ? `<ol class="s360-acts">` + acts.map(([p, t]) => `<li class="s360-act s360-act-${p.toLowerCase()}"><span class="s360-pri">${p}</span><span class="s360-act-t">${esc(t)}</span></li>`).join("") + `</ol>`
    : `<p class="text-muted s360-empty">No actions flagged for this unit.</p>`) + `</div>`;

  const bar = `<div class="d-flex gap-2 mt-3">
      <button type="button" class="btn btn-sm btn-outline-secondary" id="a360-back">← Back to Inventory</button></div>`;

  /* sticky top nav so long Asset 360 pages don't require scrolling back up */
  const topbar = `<div class="inv-stickybar">
      <button type="button" class="btn btn-sm btn-outline-secondary" id="a360-back-top">← Back to Inventory</button>
      <code class="ms-2">${esc(a.sn)}</code></div>`;

  host.innerHTML = topbar + head + `<div id="a360-edit-host"></div>` + kpis + factCard + depCard + tixCard + actCard + bar;
  const goBack = () => goToTab("dash-inventory");
  $("a360-back-top").addEventListener("click", goBack);
  $("a360-back").addEventListener("click", goBack);
  $("a360-edit").addEventListener("click", () => invToggleAssetEdit(a));
  $("a360-setstatus").addEventListener("click", () => {
    if (QBR.invSetStatus(a.key, $("a360-status").value)) renderAll();
  });
  $("a360-newticket").addEventListener("click", () => {
    QBR._invUI.prefillSn = a.sn; QBR._invUI.form = "ticket";
    goToTab("dash-inventory");
  });
  host.querySelectorAll("[data-a360-resolve]").forEach(b => b.addEventListener("click", () => {
    const note = window.prompt("Resolution note (optional):", "");
    if (note === null) return;
    if (QBR.invResolveTicket(a.key, b.dataset.a360Resolve, note.trim())) renderAll();
  }));
}

/* ============================ TICKET 360 PAGE ============================ */
QBR._invT360 = null;   // ticket no currently shown in Ticket 360
QBR._invT360Edit = false;

function openTicket360(tno) {
  const t = QBR.invFindTicket(tno);
  if (!t) return;
  QBR._invT360 = t.tno; QBR._invT360Edit = false;
  invActivateTab("dash-ticket360");
  renderTicket360Panel();
  try { history.replaceState(null, "", "#ticket/" + encodeURIComponent(t.tno)); } catch (e) {}
  window.scrollTo(0, 0);
}

function renderTicket360Panel() {
  const host = $("t360-body");
  if (!host) return;
  const inv = invModel(), tno = QBR._invT360;
  if (!inv || !inv.tickets.length) {
    host.innerHTML = `<div class="card-box"><p class="text-muted">Load an inventory workbook to browse tickets.</p></div>`;
    return;
  }
  const t = QBR.invFindTicket(tno);
  if (!t) {
    host.innerHTML = `<div class="card-box"><p class="text-muted">Ticket ${esc(tno || "—")} not found.</p>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="t360-back">← All tickets</button></div>`;
    $("t360-back").addEventListener("click", t360Back);
    return;
  }
  const open = invTixOpen(t), st = invTixStatus(t);
  const a = inv.assets.find(x => x.key === t.key);
  const days = open && t.repEdtech ? invDaysBetween(t.repEdtech, invToday()) : (t.completed && t.repEdtech ? invDaysBetween(t.repEdtech, t.completed) : null);

  const head = `<div class="card-box"><div class="d-flex justify-content-between align-items-start flex-wrap gap-2">
      <div><div class="small text-muted">Support ticket</div>
      <h4 class="mb-1"><code>${esc(t.tno)}</code></h4>
      <div class="d-flex gap-2">${invPill(st)}${invPrioPill(t.priority)}</div></div>
      <div class="text-end small text-muted">Reported ${invFmtDate(t.repEdtech)}<br>${days == null ? "" : fmt(days) + (open ? " days open" : " days to resolve")}</div>
    </div></div>`;

  /* at-a-glance overview: who requested, customer, device, issue summary */
  const facts = [
    ["Requested by", esc(t.requester || "—")], ["Customer", esc(t.client || "—")],
    ["Device", a ? invAssetLink(a.key, a.sn) : `<code>${esc(t.sn)}</code> <span class="inv-notinv">Not in inventory</span>
      <button type="button" class="btn btn-sm btn-link p-0 ms-1" data-inv-register="${escAttr(t.sn)}">Register this unit</button>`],
    ["Model", esc(t.model || (a && a.model) || "—")],
    ["Issue", `<b>${esc(t.issue || "—")}</b>`],
    ["Category", esc(t.cat || "—")], ["Person in-charge", esc(t.pic || "—")],
    ["Service address", esc(t.svcAddr || "—")]];
  const ovCard = `<div class="card-box mt-3"><h6>Overview</h6>
    <div class="s360-facts">${facts.map(([k, v]) => `<div class="s360-fact"><div class="s360-fact-k">${k}</div><div class="s360-fact-v">${v}</div></div>`).join("")}</div>
    ${t.act ? `<div class="small mt-2"><span class="text-muted">Initial findings:</span> ${esc(t.act)}</div>` : ""}</div>`;

  const dates = [["Reported to EdTech", invFmtDate(t.repEdtech)], ["Reported to Lenovo", invFmtDate(t.repLenovo)],
    ["Resolved", invFmtDate(t.completed)]];
  const dateCard = `<div class="card-box mt-3"><h6>Dates</h6>
    <div class="s360-facts">${dates.map(([k, v]) => `<div class="s360-fact"><div class="s360-fact-k">${k}</div><div class="s360-fact-v">${v}</div></div>`).join("")}</div></div>`;

  /* timeline: opened → follow-up notes → resolved */
  const evts = [{ d: t.repEdtech, t: "Ticket opened" + (t.requester ? " by " + t.requester : "") + "." }];
  (t.notes || []).forEach(n => evts.push({ d: n.d instanceof Date ? n.d : invDate(n.d), t: n.text }));
  if (t.completed) evts.push({ d: t.completed, t: "Ticket resolved." });
  evts.sort((x, y) => (x.d ? x.d.getTime() : 0) - (y.d ? y.d.getTime() : 0));
  const tlCard = `<div class="card-box mt-3"><h6>Timeline <span class="text-muted">(${fmt(evts.length)})</span></h6>` +
    `<ol class="inv-timeline">` + evts.map(e =>
      `<li class="inv-tl-item"><div class="small text-muted">${e.d ? invFmtDate(e.d) : "—"}</div><div>${esc(e.t || "")}</div></li>`
    ).join("") + `</ol></div>`;

  const relCard = `<div class="card-box mt-3"><h6>Related tickets</h6>` +
    ((t.related && t.related.length)
      ? `<div class="d-flex flex-wrap gap-2">` + t.related.map(r => invTixLink(r)).join("") + `</div>`
      : `<p class="text-muted s360-empty mb-2">No linked tickets.</p>`) +
    `<div class="d-flex gap-2"><input id="t360-rel" class="form-control form-control-sm" style="max-width:260px" placeholder="Ticket no to link…">
      <button type="button" class="btn btn-sm btn-outline-primary" id="t360-rel-go">Link ticket</button></div>
      <div id="t360-rel-msg" class="small mt-1" aria-live="polite"></div></div>`;

  /* actions */
  const actCard = `<div class="card-box mt-3"><h6>Actions</h6>
    <div class="mb-2"><label class="form-label small mb-1" for="t360-note">Add follow-up note</label>
      <div class="d-flex gap-2"><input id="t360-note" class="form-control form-control-sm" placeholder="e.g. Called client, awaiting parts…">
      <button type="button" class="btn btn-sm btn-outline-primary" id="t360-note-go">Add</button></div></div>
    <div class="d-flex gap-2 flex-wrap align-items-end">
      <div><label class="form-label small mb-1" for="t360-status">Status</label>
        <select id="t360-status" class="form-select form-select-sm">${TIX_STATUSES.map(s => `<option${s === st ? " selected" : ""}>${s}</option>`).join("")}</select></div>
      <button type="button" class="btn btn-sm btn-outline-primary" id="t360-status-go">Update status</button>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="t360-edit">${QBR._invT360Edit ? "Cancel edit" : "Edit details"}</button>
    </div>
    <div id="t360-msg" class="small mt-2" aria-live="polite"></div></div>`;

  /* edit form (hidden until Edit details) */
  const ed = QBR._invT360Edit ? `<div class="card-box mt-3"><h6>Edit ticket details</h6><div class="row g-2">
      ${[["Requester", "t360e-requester", t.requester], ["Customer", "t360e-client", t.client],
         ["Person in-charge", "t360e-pic", t.pic], ["Issue", "t360e-issue", t.issue],
         ["Service address", "t360e-svc", t.svcAddr]].map(([l, id, v]) =>
        `<div class="col-md-4"><label class="form-label small">${l}</label><input id="${id}" class="form-control form-control-sm" value="${esc(v || "")}"></div>`).join("")}
      <div class="col-md-4"><label class="form-label small">Category</label>
        <select id="t360e-cat" class="form-select form-select-sm">${["Hardware", "Software", "Network"].map(c => `<option${c === t.cat ? " selected" : ""}>${c}</option>`).join("")}</select></div>
      <div class="col-md-4"><label class="form-label small">Priority</label>
        <select id="t360e-prio" class="form-select form-select-sm"><option value="">—</option>${TIX_PRIORITIES.map(p => `<option${p === t.priority ? " selected" : ""}>${p}</option>`).join("")}</select></div>
      <div class="col-md-4"><label class="form-label small">Reported to EdTech</label><input id="t360e-edtech" type="date" class="form-control form-control-sm" value="${t.repEdtech ? t.repEdtech.toISOString().slice(0, 10) : ""}"></div>
      <div class="col-md-4"><label class="form-label small">Reported to Lenovo</label><input id="t360e-lenovo" type="date" class="form-control form-control-sm" value="${t.repLenovo ? t.repLenovo.toISOString().slice(0, 10) : ""}"></div>
      <div class="col-12"><label class="form-label small">Initial findings</label><input id="t360e-act" class="form-control form-control-sm" value="${esc(t.act || "")}"></div>
    </div>
    <div class="mt-2"><button type="button" class="btn btn-sm btn-primary" id="t360e-save">Save changes</button></div></div>` : "";

  const bar = `<div class="d-flex gap-2 mt-3">
      <button type="button" class="btn btn-sm btn-outline-secondary" id="t360-back">← All tickets</button></div>`;

  host.innerHTML = head + ovCard + dateCard + tlCard + relCard + actCard + ed + bar;
  $("t360-back").addEventListener("click", t360Back);
  $("t360-note-go").addEventListener("click", () => {
    if (QBR.invTicketNote(t.tno, $("t360-note").value)) { renderAll(); renderTicket360Panel(); }
  });
  $("t360-status-go").addEventListener("click", () => {
    const note = window.prompt("Status note (optional):", "");
    if (note === null) return;
    if (QBR.invSetTicketStatus(t.tno, $("t360-status").value, note.trim())) { renderAll(); renderTicket360Panel(); }
  });
  $("t360-rel-go").addEventListener("click", () => {
    const v = $("t360-rel").value.trim();
    if (!v) return;
    if (!QBR.invFindTicket(v)) { $("t360-rel-msg").innerHTML = `<span class="text-danger">Ticket ${esc(v)} not found.</span>`; return; }
    if (QBR.invLinkTicket(t.tno, v)) { renderTicket360Panel(); }
  });
  $("t360-edit").addEventListener("click", () => { QBR._invT360Edit = !QBR._invT360Edit; renderTicket360Panel(); });
  const sv = $("t360e-save");
  if (sv) sv.addEventListener("click", () => {
    const patch = { requester: $("t360e-requester").value.trim(), client: $("t360e-client").value.trim(),
      pic: $("t360e-pic").value.trim(), issue: $("t360e-issue").value.trim(), svcAddr: $("t360e-svc").value.trim(),
      cat: $("t360e-cat").value, priority: $("t360e-prio").value || null, act: $("t360e-act").value.trim(),
      repEdtech: $("t360e-edtech").value || null, repLenovo: $("t360e-lenovo").value || null };
    if (!patch.issue) { $("t360-msg").innerHTML = `<span class="text-danger">Issue is required.</span>`; return; }
    if (QBR.invUpdateTicket(t.tno, patch)) { QBR._invT360Edit = false; renderAll(); renderTicket360Panel(); }
  });
}
function t360Back() {
  QBR._invT360 = null; QBR._invT360Edit = false;
  try { history.replaceState(null, "", location.pathname + location.search + "#inventory"); } catch (e) {}
  goToTab("dash-inventory");
}

/* ---------- deep links: #ticket/<tno> and #asset/<key> (new-tab support) -- */
function invApplyHash() {
  const h = (location.hash || "").replace(/^#/, "");
  let m = h.match(/^ticket\/(.+)$/);
  if (m) {
    const tno = decodeURIComponent(m[1]);
    if (QBR._invT360 === tno) return true;
    if (QBR.invFindTicket(tno)) { openTicket360(tno); return true; }
    return false;
  }
  m = h.match(/^asset\/(.+)$/);
  if (m) {
    const key = decodeURIComponent(m[1]);
    if (QBR._invA360 === key) return true;
    const inv = invModel();
    if (inv && inv.assets.length) { openAsset360(key); return true; }   // unknown serial → "not in inventory" page
  }
  return false;
}
if (typeof window !== "undefined" && window.addEventListener) {
  window.addEventListener("hashchange", invApplyHash);
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => setTimeout(invApplyHash, 0));
  else setTimeout(invApplyHash, 0);
}

/* ---------- Node export guard (test harness only; inert in browser) ------ */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { QBR, invParseDeviceRows, invParseSupportRows, invParseRaksoRows,
    invParsePipelineRows, invParsePORows, invDate, invFmtDate, invDaysBetween,
    invTicketNo, invBackfillTicketNos, invParseTicketNotes, invFmtTicketNotes,
    invTixOpen, invTixStatus, TIX_STATUSES, TIX_PRIORITIES,
    invStockOverview, invCollapsible, invTicketsCard };
}
