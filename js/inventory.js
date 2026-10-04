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

QBR.INV_VERSION = "1.0.0";
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
  const p = n => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
}
function invDaysBetween(a, b) {
  if (!a || !b) return null;
  return Math.round((b - a) / 864e5);
}
function invToday() { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

/* ---------- sheet/column resolution ------------------------------------- */
const INV_SHEETS = {
  devices:  ["02 DEVICES", "DEVICES"],
  support:  ["03 TECH SUPPORT LOGS", "TECH SUPPORT LOGS", "SUPPORT LOGS"],
  raksoinv: ["04 RAKSO INV.", "04 RAKSO INV", "RAKSO INV", "04 RAKSO INVENTORY", "RAKSO INVENTORY"],
  pipeline: ["06 PIPELINE", "PIPELINE"],
  po:       ["07 PURCHASE ORDER", "PURCHASE ORDER"],
};
const SENTINELS = new Set(["", "n/a", "na", "-", "—", "–", "none", "null"]);

/* ---------- pure row parsers (Node-testable) ----------------------------- */

function invParseDeviceRows(rows) {
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
      statusOverride: null,
    });
  }
  return out;
}

function invParseSupportRows(rows) {
  const out = [];
  if (!rows || !rows.length) return out;
  const idx = makeResolver(rows[0]);
  const c = {
    sn: idx(["serial number", "serial"]), client: idx(["client / organization", "client"]),
    model: idx(["model"]), cat: idx(["category"]), status: idx(["status"]),
    repEdtech: idx(["date reported to edtech", "reported to edtech"]),
    repLenovo: idx(["date reported to lenovo", "reported to lenovo"]),
    completed: idx(["date completed"]), pic: idx(["person in-charge", "person in charge"]),
    issue: idx(["issue"]), act: idx(["activities", "troubleshooting"]), svcAddr: idx(["service location"]),
  };
  if (c.sn < 0) return out;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r], sn = row[c.sn];
    if (sn == null || String(sn).trim() === "") continue;
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : null);
    const cleanNA = v => (v != null && SENTINELS.has(v.toLowerCase())) ? null : v;
    out.push({
      key: QBR.invSerialKey(sn), sn: String(sn).trim(),
      client: cleanNA(get(c.client)), model: cleanNA(get(c.model)), cat: cleanNA(get(c.cat)) || "Hardware",
      status: get(c.status) || "Open",
      repEdtech: invDate(row[c.repEdtech]), repLenovo: invDate(row[c.repLenovo]),
      completed: invDate(row[c.completed]),
      pic: get(c.pic), issue: get(c.issue), act: get(c.act), svcAddr: get(c.svcAddr),
    });
  }
  return out;
}

/* 04 RAKSO INV: forward-fill blank SQ/Client (continuation rows) and pivot the
 * three serial columns (S/N=laptop, DEKTOP-S/N=desktop, Monitor-S/N=monitor)
 * into one deployment record per asset. Skips N/A / - / blank sentinels. */
function invParseRaksoRows(rows) {
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
                 client: fClient, req: fReq, delivered: fDel, remarks: fRem });
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
QBR.invParseSheets = function (sheets) {
  const acc = QBR._invAcc || (QBR._invAcc = { assets: [], tickets: [], deployments: [], pipeline: [], pos: [], sources: {} });
  const grab = key => {
    for (const cand of INV_SHEETS[key]) if (sheets[cand]) return sheets[cand];
    // substring fallback, mirroring findSheet() semantics in excel-loader.js
    const norm = x => String(x).toLowerCase().replace(/\s+/g, "");
    for (const cand of INV_SHEETS[key]) {
      const hit = Object.keys(sheets).find(n => norm(n).includes(norm(cand)));
      if (hit) return sheets[hit];
    }
    return null;
  };
  const dev = grab("devices");   if (dev)   { acc.assets.push(...invParseDeviceRows(dev));       acc.sources.devices = true; }
  const sup = grab("support");   if (sup)   { acc.tickets.push(...invParseSupportRows(sup));     acc.sources.support = true; }
  const rinv = grab("raksoinv"); if (rinv)  { acc.deployments.push(...invParseRaksoRows(rinv));  acc.sources.raksoinv = true; }
  const pl = grab("pipeline");   if (pl)    { acc.pipeline.push(...invParsePipelineRows(pl));    acc.sources.pipeline = true; }
  const po = grab("po");         if (po)    { acc.pos.push(...invParsePORows(po));               acc.sources.po = true; }
  return acc;
};

/* Parse raw workbook buffers (ArrayBuffers) into the inventory model. Mirrors
 * the main loader: per-buffer, filename-independent, sheets by name. */
QBR.parseInventoryBuffers = function (buffers) {
  QBR._invAcc = null;
  (buffers || []).forEach(buf => {
    try {
      const wb = XLSX.read(buf, { type: "array" });
      const sheets = {};
      wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false }); });
      QBR.invParseSheets(sheets);
    } catch (e) { /* a non-inventory workbook simply contributes nothing */ }
  });
  const acc = QBR._invAcc || { assets: [], tickets: [], deployments: [], pipeline: [], pos: [], sources: {} };
  QBR._invAcc = null;
  // de-dupe assets by serial key (newest file wins); tickets/deployments accumulate
  const seen = new Map();
  acc.assets.forEach(a => seen.set(a.key, a));
  acc.assets = [...seen.values()];
  acc.tickets.forEach((t, i) => { t._n = i; });   // stable identity for resolve buttons
  return acc;
};

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
    if (!t.completed && !/completed|resolved|closed/i.test(String(t.status || ""))) {
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
  const flags = { aging: [], lemons: [], warrantyExpiring: [], warrantyExpired: [], unmatched: [], stalled: [] };
  (inv.tickets || []).forEach(t => {
    const open = !t.completed && !/completed|resolved|closed/i.test(String(t.status || ""));
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
  return flags;
};

/* ---------- edit journal (audit trail for write-back) --------------------- */
QBR.invJournal = [];
QBR.invLog = function (action, detail) {
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
  return n;
};

QBR.invAddTicket = function (t) {
  const inv = invModel(); if (!inv || !t || !t.sn) return false;
  inv.tickets.push({
    _n: inv.tickets.length,
    key: QBR.invSerialKey(t.sn), sn: String(t.sn).trim(), client: t.client || null,
    model: t.model || null, cat: t.cat || "Hardware", status: t.status || "Open",
    repEdtech: invDate(t.repEdtech) || invToday(), repLenovo: invDate(t.repLenovo) || null,
    completed: null, pic: t.pic || null, issue: t.issue || null, act: t.act || null,
    svcAddr: t.svcAddr || null,
  });
  QBR.invLog("ticket opened", t.sn + " — " + (t.issue || "no issue text"));
  return true;
};

QBR.invResolveTicket = function (key, n, note) {
  const inv = invModel(); if (!inv) return false;
  const t = inv.tickets.find(t => t._n === Number(n) && t.key === QBR.invSerialKey(key));
  if (!t) return false;
  t.completed = invToday(); t.status = "Completed";
  if (note) t.act = (t.act ? t.act + " | " : "") + note;
  QBR.invLog("ticket resolved", t.sn + (note ? " — " + note : ""));
  return true;
};

QBR.invSetStatus = function (key, status) {
  const inv = invModel(); if (!inv) return false;
  const a = inv.assets.find(x => x.key === QBR.invSerialKey(key));
  if (!a || QBR.INV_STATUS.indexOf(status) < 0) return false;
  a.statusOverride = status;
  QBR.invLog("status change", a.sn + " → " + status);
  return true;
};

/* ---------- export: regenerate the workbook from the model ---------------- */
QBR.invExportWorkbook = function () {
  const inv = invModel();
  if (!inv || typeof XLSX === "undefined") { alert("Load an inventory workbook first."); return; }
  const openByKey = QBR.invOpenByKey(inv.tickets);
  const dev = [["Serial Number","Client / Organization","Model","Description / Specifications","Category",
    "Brand","Supplier","DR #","Date Delivered","Warranty Start","Warranty End","Warranty Years",
    "Condition","Contact Person","Address","Contact Details","Status (derived)"]];
  inv.assets.forEach(a => dev.push([a.sn, a.client || "", a.model || "", a.desc || "", a.cat, a.brand || "",
    a.supplier || "", a.dr || "", invFmtDate(a.delivered) === "—" ? "" : invFmtDate(a.delivered),
    invFmtDate(a.wstart) === "—" ? "" : invFmtDate(a.wstart), invFmtDate(a.wend) === "—" ? "" : invFmtDate(a.wend),
    a.wyears == null ? "" : a.wyears, a.cond || "", a.contact || "", a.addr || "", a.phone || "",
    QBR.invAssetStatus(a, openByKey)]));
  const sup = [["Serial Number","Client / Organization","Model","Category","Status","Date Reported To EdTech",
    "Date Reported To Lenovo Support","Date Completed","Person In-Charge","Issue/s",
    "Activities / Troubleshooting","Service Location Address"]];
  inv.tickets.forEach(t => sup.push([t.sn, t.client || "", t.model || "", t.cat, t.status,
    invFmtDate(t.repEdtech) === "—" ? "" : invFmtDate(t.repEdtech),
    invFmtDate(t.repLenovo) === "—" ? "" : invFmtDate(t.repLenovo),
    invFmtDate(t.completed) === "—" ? "" : invFmtDate(t.completed),
    t.pic || "", t.issue || "", t.act || "", t.svcAddr || ""]));
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
  XLSX.writeFile(wb, "Lenovo_Inventory_Export.xlsx");
};

/* ---------- UI state + small view helpers -------------------------------- */
QBR._invUI = { client: "ALL", type: "ALL", status: "ALL", q: "", flag: null, showAll: false, form: null, prefillSn: null };
QBR._invA360 = null;

function invPill(s) {
  const tone = { "In Stock": "blue", "Deployed": "green", "In Repair": "orange", "Retired": "gray",
                 "Open": "orange", "Completed": "green" }[s] || "blue";
  const colors = { blue: "#0078D4", green: "#107C10", orange: "#FF8C00", red: "#D13438", gray: "#605E5C" };
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

/* ============================ INVENTORY PAGE ============================= */
function renderInventory() {
  const host = $("inv-body");
  if (!host) return;
  const inv = invModel();
  if (!inv || (!inv.assets.length && !inv.tickets.length && !inv.deployments.length)) {
    host.innerHTML = `<div class="card-box"><h6>Inventory</h6>
      <p class="text-muted">No inventory data loaded yet. Upload the Lenovo inventory workbook
      (sheets 02 DEVICES, 03 TECH SUPPORT LOGS, 04 RAKSO INV., 06 PIPELINE, 07 PURCHASE ORDER)
      using the upload area above — it merges with any workbooks already loaded.</p></div>`;
    return;
  }
  const ui = QBR._invUI, T = QBR.INV_THRESH, today = invToday();
  const openByKey = QBR.invOpenByKey(inv.tickets);
  const flags = QBR.invComputeFlags(inv, today);
  const withStatus = inv.assets.map(a => ({ a, st: QBR.invAssetStatus(a, openByKey) }));

  /* ---- KPIs ---- */
  const deployed = withStatus.filter(x => x.st === "Deployed").length;
  const inRepair = withStatus.filter(x => x.st === "In Repair").length;
  const openT = inv.tickets.filter(t => !t.completed && !/completed|resolved|closed/i.test(String(t.status || ""))).length;
  const avg = invAvgResolve(inv.tickets);
  const kpis = `<div class="row row-cols-2 row-cols-md-3 row-cols-xl-5 g-3 kpi-row">` +
    kpi("Fleet size", fmt(withStatus.length), "blue") +
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
      <button type="button" class="btn btn-sm btn-outline-secondary" id="inv-export">Export inventory workbook</button>
    </div>
    <div class="d-flex flex-wrap gap-2 align-items-center mt-2">
      <label class="form-label small mb-0" for="inv-lookup"><strong>Look up asset</strong></label>
      <input id="inv-lookup" class="form-control form-control-sm" style="--w:220px" placeholder="Type or paste a serial…" autocomplete="off">
      <button type="button" class="btn btn-sm btn-primary" id="inv-lookup-go">Open</button>
      <span id="inv-lookup-msg" class="small text-muted"></span>
    </div><div id="inv-form-host" class="mt-2"></div>`;

  /* ---- filters ---- */
  const clients = invClients(inv);
  const filters = `<div class="card-box mt-3"><div class="d-flex flex-wrap gap-2 align-items-end">
      <div><label class="form-label small mb-0" for="inv-f-client">Client</label>
        <select id="inv-f-client" class="form-select form-select-sm" style="--w:220px">
        <option value="ALL">All clients</option>${clients.map(c => `<option${ui.client === c ? " selected" : ""}>${esc(c)}</option>`).join("")}</select></div>
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
  if (ui.client !== "ALL") rows = rows.filter(x => x.a.client === ui.client);
  if (ui.type !== "ALL") rows = rows.filter(x => x.a.cat === ui.type);
  if (ui.status !== "ALL") rows = rows.filter(x => x.st === ui.status);
  if (ui.q) { const q = QBR.invSerialKey(ui.q); rows = rows.filter(x => x.key.includes(q)); }
  if (ui.flag === "warrantyExpiring") { const s = new Set(flags.warrantyExpiring.map(a => a.key)); rows = rows.filter(x => s.has(x.key)); }
  if (ui.flag === "warrantyExpired") { const s = new Set(flags.warrantyExpired.map(a => a.key)); rows = rows.filter(x => s.has(x.key)); }
  if (ui.flag === "lemons") { const s = new Set(flags.lemons.map(l => l.key)); rows = rows.filter(x => s.has(x.key)); }
  const tixByKey = new Map();
  inv.tickets.forEach(t => tixByKey.set(t.key, (tixByKey.get(t.key) || 0) + 1));
  const total = rows.length, shown = ui.showAll ? rows : rows.slice(0, 150);
  const assetTbl = `<div class="card-box mt-3"><h6>Assets <span class="text-muted">(${fmt(total)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>Serial</th><th>Client</th><th>Model</th><th>Type</th><th>Status</th><th>Warranty end</th><th class="text-end">Tickets</th></tr></thead><tbody>` +
    (shown.map(x => `<tr class="inv-row" data-inv-key="${esc(x.a.key)}" tabindex="0" role="button" aria-label="Open asset ${esc(x.a.sn)}">
      <td><code>${esc(x.a.sn)}</code></td><td>${esc(x.a.client || "—")}</td><td>${esc(x.a.model || "—")}</td>
      <td>${esc(x.a.cat)}</td><td>${invPill(x.st)}</td><td>${invFmtDate(x.a.wend)}</td>
      <td class="text-end">${fmt(tixByKey.get(x.a.key) || 0)}</td></tr>`).join("") ||
      `<tr><td colspan="7" class="text-muted">No assets match the current filters.</td></tr>`) +
    `</tbody></table></div>` +
    (!ui.showAll && total > 150 ? `<button type="button" class="btn btn-sm btn-outline-secondary" id="inv-showall">Show all ${fmt(total)}</button>` : "") + `</div>`;

  /* ---- tickets table ---- */
  let tix = inv.tickets.slice().sort((a, b) => (b.repEdtech || 0) - (a.repEdtech || 0));
  if (ui.flag === "aging") { const s = new Set(flags.aging.map(t => t.sn + "|" + (t.repEdtech || ""))); tix = tix.filter(t => s.has(t.sn + "|" + (t.repEdtech || ""))); }
  if (ui.flag === "unmatched") { const s = new Set(flags.unmatched.map(u => u.key)); tix = tix.filter(t => s.has(t.key)); }
  if (ui.flag === "lemons") { const s = new Set(flags.lemons.map(l => l.key)); tix = tix.filter(t => s.has(t.key)); }
  const tixTbl = `<div class="card-box mt-3"><h6>Support tickets <span class="text-muted">(${fmt(tix.length)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>Serial</th><th>Client</th><th>Status</th><th>Reported</th><th class="text-end">Days open</th><th>Issue</th><th>Person</th><th></th></tr></thead><tbody>` +
    (tix.slice(0, 200).map((t, i) => {
      const open = !t.completed && !/completed|resolved|closed/i.test(String(t.status || ""));
      const days = open && t.repEdtech ? invDaysBetween(t.repEdtech, today) : (t.completed && t.repEdtech ? invDaysBetween(t.repEdtech, t.completed) : null);
      return `<tr><td><a href="#" data-inv-key="${esc(t.key)}" class="inv-link"><code>${esc(t.sn)}</code></a></td>
        <td>${esc(t.client || "—")}</td><td>${invPill(open ? "Open" : "Completed")}</td>
        <td>${invFmtDate(t.repEdtech)}</td><td class="text-end">${days == null ? "—" : fmt(days)}</td>
        <td>${esc(t.issue || "—")}</td><td>${esc(t.pic || "—")}</td>
        <td>${open ? `<button type="button" class="btn btn-sm btn-outline-success" data-inv-resolve="${t._n}">Resolve</button>` : ""}</td></tr>`;
    }).join("") || `<tr><td colspan="8" class="text-muted">No tickets match.</td></tr>`) +
    `</tbody></table></div>${tix.length > 200 ? `<div class="small text-muted">Showing latest 200 of ${fmt(tix.length)}.</div>` : ""}</div>`;

  /* ---- pipeline + PO ---- */
  const plTbl = `<div class="card-box mt-3"><h6>Pipeline <span class="text-muted">(${fmt(inv.pipeline.length)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>SQ</th><th>Client</th><th class="text-end">Qty</th><th>Proposal</th><th>Delivery</th><th>Remarks</th></tr></thead><tbody>` +
    (inv.pipeline.map(p => `<tr class="${/signed/i.test(String(p.proposal || "")) && !/delivered/i.test(String(p.delivery || "")) ? "table-warning" : ""}">
      <td><code>${esc(p.sq)}</code></td><td>${esc(p.client || "—")}</td><td class="text-end">${p.qty == null ? "—" : fmt(p.qty)}</td>
      <td>${esc(p.proposal || "—")}</td><td>${esc(p.delivery || "—")}</td><td>${esc(p.rem || "—")}</td></tr>`).join("") ||
      `<tr><td colspan="6" class="text-muted">No pipeline rows.</td></tr>`) + `</tbody></table></div></div>`;
  const poTbl = `<div class="card-box mt-3"><h6>Purchase orders <span class="text-muted">(${fmt(inv.pos.length)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>PO#</th><th>School</th><th>Units</th><th class="text-end">Qty</th><th>Date</th></tr></thead><tbody>` +
    (inv.pos.map(p => `<tr><td><code>${esc(p.po)}</code></td><td>${esc(p.school || "—")}</td><td>${esc(p.units || "—")}</td>
      <td class="text-end">${p.qty == null ? "—" : fmt(p.qty)}</td><td>${invFmtDate(p.date)}</td></tr>`).join("") ||
      `<tr><td colspan="5" class="text-muted">No purchase orders.</td></tr>`) + `</tbody></table></div></div>`;

  host.innerHTML = kpis + flagCards + actions + filters + assetTbl + tixTbl + plTbl + poTbl;
  invBind(host);
  if (ui.form) invShowForm(ui.form);
}

/* ---------- event wiring for the inventory page -------------------------- */
function invBind(host) {
  const ui = QBR._invUI;
  const rerender = () => renderInventory();
  host.querySelectorAll("[data-inv-flag]").forEach(b => b.addEventListener("click", () => {
    ui.flag = b.dataset.invFlag || null; rerender();
  }));
  host.querySelectorAll("[data-inv-form]").forEach(b => b.addEventListener("click", () => {
    ui.form = ui.form === b.dataset.invForm ? null : b.dataset.invForm; rerender();
  }));
  const ex = $("inv-export");
  if (ex) ex.addEventListener("click", () => QBR.invExportWorkbook());
  const lgo = $("inv-lookup-go"), lq = $("inv-lookup");
  if (lgo && lq) {
    lgo.addEventListener("click", invLookupSerial);
    lq.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); invLookupSerial(); } });
  }
  const sa = $("inv-showall");
  if (sa) sa.addEventListener("click", () => { ui.showAll = true; rerender(); });
  const fc = $("inv-f-client"); if (fc) fc.addEventListener("change", () => { ui.client = fc.value; ui.showAll = false; rerender(); });
  const ft = $("inv-f-type");   if (ft) ft.addEventListener("change", () => { ui.type = ft.value; ui.showAll = false; rerender(); });
  const fs = $("inv-f-status"); if (fs) fs.addEventListener("change", () => { ui.status = fs.value; ui.showAll = false; rerender(); });
  const fq = $("inv-f-q");
  if (fq) fq.addEventListener("change", () => { ui.q = fq.value.trim(); ui.showAll = false; rerender(); });
  host.querySelectorAll(".inv-row[data-inv-key]").forEach(r => {
    const go = e => { e.preventDefault(); openAsset360(r.dataset.invKey); };
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
        </div></div>
      </div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="in-go">Register</button><button type="button" class="btn btn-sm btn-outline-primary" id="in-scan">Scan box label</button>${close}</div>
      <div id="in-msg" class="small mt-1" aria-live="polite"></div></div>`;
    $("in-go").addEventListener("click", () => {
      const serials = $("in-serials").value.split(/\n+/).map(s => s.trim()).filter(Boolean);
      if (!serials.length) { $("in-msg").textContent = "Enter at least one serial number."; return; }
      const n = QBR.invIntake(serials.map(sn => ({ sn, model: $("in-model").value, cat: $("in-cat").value,
        supplier: $("in-supplier").value.trim() || null, wyears: Number($("in-wy").value),
        dr: $("in-dr").value.trim() || null, client: $("in-client").value.trim() || null, wstart: $("in-ws").value || null })));
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
          <input id="tk-sn" class="form-control form-control-sm" value="${esc(pre)}" placeholder="PF4ABC123"></div>
        <div class="col-md-3"><label class="form-label small">Category</label>
          <select id="tk-cat" class="form-select form-select-sm"><option>Hardware</option><option>Software</option><option>Network</option></select></div>
        <div class="col-md-3"><label class="form-label small">Person in-charge</label><input id="tk-pic" class="form-control form-control-sm"></div>
        <div class="col-md-3"><label class="form-label small">Date reported</label><input id="tk-date" type="date" class="form-control form-control-sm"></div>
        <div class="col-md-6"><label class="form-label small">Issue *</label><input id="tk-issue" class="form-control form-control-sm" placeholder="e.g. No power"></div>
        <div class="col-md-6"><label class="form-label small">Initial findings</label><input id="tk-act" class="form-control form-control-sm"></div>
      </div>
      <div class="mt-2 d-flex gap-2"><button type="button" class="btn btn-sm btn-primary" id="tk-go">Open ticket</button>${close}</div>
      <div id="tk-msg" class="small mt-1" aria-live="polite"></div></div>`;
    ui.prefillSn = null;
    $("tk-go").addEventListener("click", () => {
      const sn = $("tk-sn").value.trim(), issue = $("tk-issue").value.trim();
      if (!sn || !issue) { $("tk-msg").textContent = "Serial and issue are required."; return; }
      const inv2 = invModel(), a = inv2 ? inv2.assets.find(x => x.key === QBR.invSerialKey(sn)) : null;
      const ok = QBR.invAddTicket({ sn, client: a ? a.client : null, model: a ? a.model : null,
        cat: $("tk-cat").value, pic: $("tk-pic").value.trim() || null,
        repEdtech: $("tk-date").value || null, issue, act: $("tk-act").value.trim() || null });
      $("tk-msg").innerHTML = ok ? `<span class="text-success">Ticket opened for ${esc(sn)}.</span>` : `<span class="text-danger">Could not open ticket.</span>`;
      if (ok) setTimeout(() => { ui.form = null; renderAll(); }, 900);
    });
  }
  host.querySelectorAll("[data-inv-form]").forEach(b => b.addEventListener("click", () => { ui.form = null; renderInventory(); }));
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
  window.scrollTo(0, 0);
}

/* Quick serial lookup from the Inventory page: exact match jumps straight to
 * Asset 360, several matches narrow the Assets table, no match offers Intake
 * with the serial prefilled. */
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
    ["Date delivered", invFmtDate(a.delivered)], ["Warranty", a.wstart || a.wend ? `${invFmtDate(a.wstart)} → ${invFmtDate(a.wend)}${a.wyears ? ` (${a.wyears}y)` : ""}` : "—"],
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
      const isOpen = !t.completed && !/completed|resolved|closed/i.test(String(t.status || ""));
      return `<li class="inv-tl-item"><div class="d-flex justify-content-between gap-2">
        <b>${esc(t.issue || "—")}</b>${invPill(isOpen ? "Open" : "Completed")}</div>
        <div class="small text-muted">Reported ${invFmtDate(t.repEdtech)}${t.repLenovo ? ` · Lenovo ${invFmtDate(t.repLenovo)}` : ""}${t.completed ? ` · Completed ${invFmtDate(t.completed)}` : ""} · ${esc(t.pic || "—")}</div>
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

  host.innerHTML = head + kpis + factCard + depCard + tixCard + actCard + bar;
  $("a360-back").addEventListener("click", () => goToTab("dash-inventory"));
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

/* ---------- Node export guard (test harness only; inert in browser) ------ */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { QBR, invParseDeviceRows, invParseSupportRows, invParseRaksoRows,
    invParsePipelineRows, invParsePORows, invDate, invFmtDate, invDaysBetween };
}
