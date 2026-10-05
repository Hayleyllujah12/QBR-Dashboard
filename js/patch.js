/* ============================================================================
 * QBR Dashboard — patch-in-place workbook saver (js/patch.js)
 * v1.15 — 2026-10-04
 *
 * Problem it solves: the v1.8–v1.14 save flow REBUILT the workbook from
 * parsed data (supBuildWorkbook / invBuildWorkbook) and overwrote the linked
 * file with it. That destroyed the original's formulas (supplies Inventory
 * Summary XLOOKUP/SUMIFS, Items =CONCAT helpers) and its layout, producing
 * #NAME? errors and corrupted cells.
 *
 * New approach: keep the ORIGINAL workbook object per file fingerprint
 * (QBR._origWb, stashed at parse time), and replay the mutation journal as
 * targeted CELL WRITES into those original sheets:
 *   - appends  → new rows at the end of the original sheet (extending !ref);
 *                new rows inherit style/number-format from the row above;
 *                Items!A keeps the =CONCAT(L#,"-",M#) pattern when present.
 *   - updates  → mutate the existing cell's VALUE in place, never replacing
 *                the cell object, so style/number-format survive; a cell that
 *                holds a FORMULA is never overwritten (counted as skipped).
 *   - summary  → untouched; Excel recalculates on open (fullCalcOnLoad) plus
 *                a conservative range-bump: bounded $X$2:$Y$N references that
 *                ended exactly at the old last row are extended to the new one.
 *
 * Source coordinates come from the parsers: every parsed object carries
 * _src {fp, sheet, row} and each sheet's resolved column map lives in
 * QBR._sheetMeta[fp][sheet] = {kind, cols}. Multi-file merges are handled:
 * rows whose _src.fp differs from the linked file are skipped and reported.
 *
 * Depends on: QBR namespace, XLSX, invModel/supModel, invDate, invToday,
 *   invSerialKey, invFmtTicketNotes, QBR.invFindTicket (all runtime-only,
 *   guarded). Loaded AFTER inventory.js + supplies.js, BEFORE persist.js.
 * ==========================================================================*/
var QBR = window.QBR = window.QBR || {};

/* ============================ low-level helpers ========================= */

function pSheet(wb, name) {
  if (!wb || !name) return null;
  const i = (wb.SheetNames || []).indexOf(name);
  return i >= 0 ? wb.Sheets[wb.SheetNames[i]] : null;
}
/* 1-indexed last data row of a sheet (header counts as row 1). */
function pLastRow(ws) {
  if (!ws || !ws["!ref"]) return 1;
  try { return XLSX.utils.decode_range(ws["!ref"]).e.r + 1; } catch (e) { return 1; }
}
function pExtendRef(ws, r, c) {
  let range;
  try { range = XLSX.utils.decode_range(ws["!ref"] || "A1"); }
  catch (e) { range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; }
  range.e.r = Math.max(range.e.r, r - 1);
  range.e.c = Math.max(range.e.c, c);
  ws["!ref"] = XLSX.utils.encode_range(range);
}
/* Normalize a JS value to a SheetJS cell descriptor (no formula). */
function pVal(v) {
  if (v instanceof Date && !isNaN(v)) return { t: "d", v: v };
  if (typeof v === "number" && isFinite(v)) return { t: "n", v: v };
  if (typeof v === "boolean") return { t: "b", v: v };
  if (v == null) return null;
  const s = String(v);
  return s === "" ? null : { t: "s", v: s };
}
/* Write a value into (r, c) — r 1-indexed, c 0-indexed. Updates the existing
 * cell's value IN PLACE (keeps style/number format). Never overwrites a
 * formula cell unless opts.formula is given. Returns "ok"/"skipped-formula". */
function pSet(ws, r, c, v, opts) {
  if (c == null || c < 0 || r < 1) return "skipped-no-col";
  const addr = XLSX.utils.encode_cell({ r: r - 1, c: c });
  const cell = ws[addr];
  if (cell && cell.f && !(opts && opts.formula)) return "skipped-formula";
  if (opts && opts.formula) {
    const nc = { t: opts.t || "s", f: opts.formula, v: opts.v != null ? opts.v : "" };
    if (cell && cell.s) nc.s = cell.s;
    if (cell && cell.z) nc.z = cell.z;
    ws[addr] = nc;
    pExtendRef(ws, r, c);
    return "ok";
  }
  const d = pVal(v);
  if (d == null) { // clear value, keep the cell (style) if it exists
    if (cell) { delete cell.f; cell.t = "s"; cell.v = ""; }
    return "ok";
  }
  if (cell) { cell.t = d.t; cell.v = d.v; delete cell.f; }
  else {
    const nc = { t: d.t, v: d.v };
    // inherit style + number format from the cell above (same column)
    const above = ws[XLSX.utils.encode_cell({ r: r - 2, c: c })];
    if (above) { if (above.s) nc.s = above.s; if (above.z) nc.z = above.z; }
    ws[addr] = nc;
  }
  pExtendRef(ws, r, c);
  return "ok";
}
/* Append one row; `cells` is a sparse array indexed by column. Each entry is
 * a raw value or {formula, v, t}. Returns the new 1-indexed row number. */
function pAppendRow(ws, cells) {
  const r = pLastRow(ws) + 1;
  cells.forEach((entry, c) => {
    if (entry === undefined) return;
    if (entry && typeof entry === "object" && entry.formula) {
      pSet(ws, r, c, null, { formula: entry.formula, v: entry.v, t: entry.t });
    } else {
      pSet(ws, r, c, entry);
    }
  });
  return r;
}
/* Get-or-create a sheet; on create, write the header row and register cols. */
function pEnsureSheet(wb, fp, kind, name, headers) {
  let ws = pSheet(wb, name);
  if (ws) return { name, ws };
  ws = XLSX.utils.aoa_to_sheet([headers]);
  XLSX.utils.book_append_sheet(wb, ws, name);
  const cols = {};
  headers.forEach((h, i) => { cols["c" + i] = i; }); // positional fallback
  return { name, ws, created: true, headers };
}

function sheetMetaFor(fp, kind) {
  const per = ((typeof QBR !== "undefined" && QBR._sheetMeta) || {})[fp] || {};
  for (const name of Object.keys(per)) {
    if (per[name] && per[name].kind === kind) return { name, cols: per[name].cols || {} };
  }
  return null;
}
/* Batch Code column (2026-10-05): when a batch code must be written and the
 * devices sheet has no such column yet, add "Batch Code" right after the last
 * used column of the header row (no existing column moves), copy the header
 * style from its left neighbour, and register it in the shared cols map so
 * later ops in the same save find it. */
function pEnsureCol(ws, cols, key, header) {
  if (!ws || !cols) return -1;
  if (colOf(cols, key) >= 0) return cols[key];
  let last = 0;
  try { last = XLSX.utils.decode_range(ws["!ref"] || "A1").e.c; } catch (e) {}
  const c = last + 1;
  pSet(ws, 1, c, header);
  const left = ws[XLSX.utils.encode_cell({ r: 0, c: last })], me = ws[XLSX.utils.encode_cell({ r: 0, c: c })];
  if (left && me) { if (left.s) me.s = left.s; }
  if (ws["!cols"] && ws["!cols"][last] && !ws["!cols"][c]) ws["!cols"][c] = Object.assign({}, ws["!cols"][last]);
  cols[key] = c;
  return c;
}
function colOf(cols, key) {
  const i = cols ? cols[key] : -1;
  return (typeof i === "number" && i >= 0) ? i : -1;
}

/* ============================ op mappers ================================ */
/* Each mapper: (wb, fp, args) -> "ok" | "noop" | "skip: reason".
 * Appends resolve the model object afterwards to stamp _src (so later ops in
 * the same pass, e.g. intake→deploy, hit the new row). Updates resolve _src
 * first and skip rows that don't live in the linked file. */

function pModel(kind) {
  try {
    if (kind === "supplies" && typeof supModel === "function") return supModel();
    if (kind === "assets" && typeof invModel === "function") return invModel();
  } catch (e) {}
  return null;
}
function pTodayYMD() {
  const n = new Date();
  return n.getFullYear() + "-" + String(n.getMonth() + 1).padStart(2, "0") + "-" + String(n.getDate()).padStart(2, "0");
}

/* ---- supplies: Transactions ---- */
function patchSupRecordTransaction(wb, fp, args) {
  const t = args[0] || {};
  let sm = sheetMetaFor(fp, "trans");
  let ws, cols;
  if (sm) { ws = pSheet(wb, sm.name); cols = sm.cols; }
  if (!ws) {
    const headers = ["Date", "DocNo", "TransactionType", "ItemID", "Quantity", "Department/Project", "RequestedBy", "ApprovedBy", "Notes"];
    const made = pEnsureSheet(wb, fp, "trans", "Transactions", headers);
    ws = made.ws;
    cols = { date: 0, doc: 1, type: 2, item: 3, qty: 4, dept: 5, req: 6, app: 7, notes: 8 };
  }
  const oldLast = pLastRow(ws);
  const date = (typeof invDate === "function" ? invDate(t.date) : null) || new Date();
  const cells = [];
  const put = (k, v) => { const c = colOf(cols, k); if (c >= 0) cells[c] = v; };
  put("date", date);
  put("doc", t.doc || "");
  put("type", t.type || "Used");
  put("item", String(t.itemId || "").trim());
  put("qty", isFinite(Number(t.qty)) ? Number(t.qty) : 0);
  put("dept", t.dept || "");
  put("req", t.requestedBy || "");
  put("app", t.approvedBy || "");
  put("notes", t.notes || "");
  const r = pAppendRow(ws, cells);
  pBumpRanges(wb, sm ? sm.name : "Transactions", oldLast, r);
  // stamp _src on the model transaction (first unmatched with same item+qty)
  try {
    const sup = pModel("supplies");
    const wantId = String(t.itemId || "").trim(), wantQty = isFinite(Number(t.qty)) ? Number(t.qty) : 0;
    const hit = (sup.transactions || []).find(x => !x._src &&
      String(x.itemId || "").trim() === wantId && Number(x.qty) === wantQty);
    if (hit) hit._src = { fp, sheet: sm ? sm.name : "Transactions", row: r };
  } catch (e) {}
  return "ok";
}

/* ---- supplies: Items ---- */
function patchSupAddItem(wb, fp, args) {
  const it = args[0] || {};
  let sm = sheetMetaFor(fp, "items");
  let ws, cols, sheetName;
  if (sm) { ws = pSheet(wb, sm.name); cols = sm.cols; sheetName = sm.name; }
  if (!ws) {
    const headers = ["ItemID", "ItemName", "Category", "Unit", "BeginningBalance", "ReorderLevel", "Supplier", "UnitCost", "Location"];
    const made = pEnsureSheet(wb, fp, "items", "Items", headers);
    ws = made.ws; sheetName = made.name;
    cols = { id: 0, name: 1, cat: 2, unit: 3, beg: 4, reorder: 5, supplier: 6, cost: 7, loc: 8 };
  }
  const name = String(it.name || "").trim();
  if (!name) return "skip: blank item";
  const code = String(it.code || "").trim();
  const id = (code + "-" + name).replace(/^-|-$/g, "") || name;
  // idempotency: already written in a previous patch pass
  try {
    const sup = pModel("supplies");
    const key = (typeof QBR.supItemKey === "function") ? QBR.supItemKey(id) : id;
    const existing = (sup.items || []).find(x => x.key === key);
    if (existing && existing._src && existing._src.fp === fp) return "noop";
  } catch (e) {}
  const namePart = code ? id.slice(code.length + 1) : name;
  const oldLast = pLastRow(ws);
  const r = oldLast + 1;
  const cells = [];
  const put = (k, v) => { const c = colOf(cols, k); if (c >= 0) cells[c] = v; };
  // Column A: keep the workbook's =CONCAT(L#,"-",M#) pattern when the sheet
  // uses it (copy the row above's formula with the row number adjusted);
  // otherwise write the static ID.
  const cId = colOf(cols, "id");
  if (cId >= 0) {
    const aboveA = ws[XLSX.utils.encode_cell({ r: r - 2, c: cId })];
    const af = aboveA && aboveA.f;
    if (af && /^CONCAT\(\s*L\d+\s*,\s*"-"\s*,\s*M\d+\s*\)$/i.test(String(af).trim())) {
      cells[cId] = { formula: "CONCAT(L" + r + ',\"-\",M' + r + ")", v: id, t: "s" };
    } else {
      cells[cId] = id;
    }
  }
  put("name", name);
  put("cat", String(it.cat || "").trim() || "");
  put("unit", String(it.unit || "").trim() || "pcs");
  put("beg", isFinite(Number(it.beginning)) ? Number(it.beginning) : 0);
  put("reorder", isFinite(Number(it.reorder)) ? Number(it.reorder) : 0);
  put("supplier", String(it.supplier || "").trim());
  put("cost", isFinite(Number(it.cost)) ? Number(it.cost) : 0);
  put("loc", String(it.loc || "").trim());
  // helper columns L/M (indices 11/12) when the sheet has them
  if (ws[XLSX.utils.encode_cell({ r: 0, c: 11 })] !== undefined || colOf(cols, "id") === 0) {
    cells[11] = code;
    cells[12] = namePart;
  }
  const newR = pAppendRow(ws, cells);
  pBumpRanges(wb, sheetName, oldLast, newR);
  try {
    const sup = pModel("supplies");
    const key = (typeof QBR.supItemKey === "function") ? QBR.supItemKey(id) : id;
    const hit = (sup.items || []).find(x => x.key === key);
    if (hit) hit._src = { fp, sheet: sheetName, row: newR };
  } catch (e) {}
  return "ok";
}

/* ---- assets: intake (append to 02 DEVICES) ---- */
function patchInvIntake(wb, fp, args) {
  const list = args[0] || [];
  let sm = sheetMetaFor(fp, "devices");
  let ws, cols, sheetName;
  if (sm) { ws = pSheet(wb, sm.name); cols = sm.cols; sheetName = sm.name; }
  if (!ws) {
    const headers = ["Serial Number", "Client / Organization", "Model", "Description / Specifications", "Category", "Brand", "Supplier", "DR #", "Date Delivered", "Warranty Start", "Warranty End", "Warranty Years", "Condition", "Contact Person", "Address", "Contact Details", "Batch Code"];
    const made = pEnsureSheet(wb, fp, "devices", "02 DEVICES", headers);
    ws = made.ws; sheetName = made.name;
    cols = { sn: 0, client: 1, model: 2, desc: 3, cat: 4, brand: 5, supplier: 6, dr: 7, delivered: 8, wstart: 9, wend: 10, wyears: 11, cond: 12, contact: 13, addr: 14, phone: 15, batch: 16 };
  }
  const inv = pModel("assets");
  let n = 0;
  (list || []).forEach(r0 => {
    const sn = String((r0 && r0.sn) || "").trim();
    if (!sn) return;
    let a = null;
    try {
      const key = (typeof QBR.invSerialKey === "function") ? QBR.invSerialKey(sn) : sn;
      a = ((inv && inv.assets) || []).find(x => x.key === key);
    } catch (e) {}
    if (a && a._src && a._src.fp === fp) return; // already written (idempotent)
    const src = a || r0;
    const cells = [];
    const put = (k, v) => { const c = colOf(cols, k); if (c >= 0) cells[c] = v; };
    put("sn", sn);
    put("client", src.client || "");
    put("model", src.model || "");
    put("desc", src.desc || "");
    put("cat", src.cat || "Laptop");
    put("brand", src.brand || "Lenovo");
    put("supplier", src.supplier || "");
    put("dr", src.dr || "");
    // delivered left blank on intake (matches model: delivered = null)
    put("wstart", src.wstart instanceof Date ? src.wstart : null);
    put("wend", src.wend instanceof Date ? src.wend : null);
    put("wyears", src.wyears != null ? Number(src.wyears) : "");
    put("cond", src.cond || "No Issue");
    put("contact", src.contact || "");
    put("addr", src.addr || "");
    put("phone", src.phone || "");
    if (src.batch) { pEnsureCol(ws, cols, "batch", "Batch Code"); put("batch", src.batch); }
    const r = pAppendRow(ws, cells);
    pBumpRanges(wb, sheetName, r - 1, r);
    if (a) a._src = { fp, sheet: sheetName, row: r };
    n++;
  });
  return n ? "ok" : "noop";
}

/* ---- assets: deploy (update device rows + append RAKSO row) ---- */
function patchInvDeploy(wb, fp, args) {
  const keys = args[0] || [], client = args[1], dateDelivered = args[2], sq = args[3];
  const sm = sheetMetaFor(fp, "devices");
  if (!sm) return "skip: no devices sheet";
  const ws = pSheet(wb, sm.name);
  if (!ws) return "skip: no devices sheet";
  const cols = sm.cols;
  const inv = pModel("assets");
  const delDate = (typeof invDate === "function" ? invDate(dateDelivered) : null) || new Date();
  let n = 0, skipped = 0;
  (keys || []).forEach(k => {
    let a = null;
    try {
      const key = (typeof QBR.invSerialKey === "function") ? QBR.invSerialKey(k) : k;
      a = ((inv && inv.assets) || []).find(x => x.key === key);
    } catch (e) {}
    if (!a || !a._src || a._src.fp !== fp) { skipped++; return; }
    const aws = pSheet(wb, a._src.sheet);
    if (!aws) { skipped++; return; }
    const ac = (((QBR._sheetMeta || {})[fp] || {})[a._src.sheet] || {}).cols || cols;
    pSet(aws, a._src.row, colOf(ac, "client"), client || a.client || "");
    pSet(aws, a._src.row, colOf(ac, "delivered"), delDate);
    if (sq) {
      try {
        const dep = ((inv && inv.deployments) || []).filter(d => !d._src && d.key === a.key && d.sq === sq).pop()
          || { key: a.key, sn: a.sn, type: a.cat, desc: a.desc, sq, client: client || a.client, req: null, delivered: delDate, remarks: "Deployed via dashboard" };
        patchAppendDeployment(wb, fp, dep);
      } catch (e) {}
    }
    n++;
  });
  if (!n && skipped) return "skip: rows not in linked file";
  return n ? "ok" : "noop";
}

function patchAppendDeployment(wb, fp, dep) {
  let sm = sheetMetaFor(fp, "rakso");
  let ws, cols, sheetName;
  if (sm) { ws = pSheet(wb, sm.name); cols = sm.cols; sheetName = sm.name; }
  if (!ws) {
    const headers = ["SQ", "Client", "S/N", "Laptop", "Laptop Qty", "Desktop-S/N", "Monitor-S/N", "Desktop", "Date Requested", "Date Delivered", "Remarks"];
    const made = pEnsureSheet(wb, fp, "rakso", "04 RAKSO INV.", headers);
    ws = made.ws; sheetName = made.name;
    cols = { sq: 0, client: 1, sn: 2, lapDesc: 3, lapQty: 4, dsn: 5, msn: 6, ddesc: 7, req: 8, del: 9, rem: 10 };
  }
  const oldLast = pLastRow(ws);
  const cells = [];
  const put = (k, v) => { const c = colOf(cols, k); if (c >= 0) cells[c] = v; };
  put("sq", dep.sq || "");
  put("client", dep.client || "");
  // serial goes into the type-appropriate column of the original pivot layout
  const t = String(dep.type || "Laptop");
  if (/desktop/i.test(t)) { put("dsn", dep.sn || ""); put("ddesc", dep.desc || ""); }
  else if (/monitor/i.test(t)) { put("msn", dep.sn || ""); }
  else { put("sn", dep.sn || ""); put("lapDesc", dep.desc || ""); }
  put("req", dep.req instanceof Date ? dep.req : null);
  put("del", dep.delivered instanceof Date ? dep.delivered : null);
  put("rem", dep.remarks || "");
  const r = pAppendRow(ws, cells);
  pBumpRanges(wb, sheetName, oldLast, r);
  try { dep._src = { fp, sheet: sheetName, row: r }; } catch (e) {}
  return r;
}

/* ---- assets: tickets ---- */
function patchTicketRow(wb, fp, tno) {
  // returns {ws, row, cols} for the ticket's source row, or null
  let ticket = null;
  try { ticket = (typeof QBR.invFindTicket === "function") ? QBR.invFindTicket(tno) : null; } catch (e) {}
  if (!ticket || !ticket._src || ticket._src.fp !== fp) return null;
  const ws = pSheet(wb, ticket._src.sheet);
  if (!ws) return null;
  const cols = (((QBR._sheetMeta || {})[fp] || {})[ticket._src.sheet] || {}).cols || {};
  return { ws, row: ticket._src.row, cols, ticket };
}
function patchAssetRow(wb, fp, key) {
  // returns {ws, row, cols} for the asset's source row in 02 DEVICES, or null
  let asset = null;
  try {
    const inv = pModel("assets");
    asset = inv ? inv.assets.find(a => a.key === key) : null;
  } catch (e) {}
  if (!asset || !asset._src || asset._src.fp !== fp) return null;
  const ws = pSheet(wb, asset._src.sheet);
  if (!ws) return null;
  const cols = (((QBR._sheetMeta || {})[fp] || {})[asset._src.sheet] || {}).cols || {};
  return { ws, row: asset._src.row, cols, asset };
}
function patchInvUpdateAsset(wb, fp, args) {
  const rawKey = args[0], patch = args[1] || {};
  const key = (typeof QBR.invSerialKey === "function") ? QBR.invSerialKey(rawKey) : String(rawKey || "").trim().toUpperCase();
  const ar = patchAssetRow(wb, fp, key);
  if (!ar) return "skip: asset row not in linked file";
  if (patch.batch) pEnsureCol(ar.ws, ar.cols, "batch", "Batch Code");
  const map = { client: "client", model: "model", desc: "desc", cat: "cat", brand: "brand", supplier: "supplier", dr: "dr", cond: "cond", contact: "contact", addr: "addr", phone: "phone", wyears: "wyears", batch: "batch" };
  Object.keys(map).forEach(k => {
    if (patch[k] !== undefined) pSet(ar.ws, ar.row, colOf(ar.cols, map[k]), patch[k] == null ? "" : patch[k]);
  });
  const D = v => (typeof invDate === "function" ? invDate(v) : null);
  ["delivered", "wstart", "wend"].forEach(k => {
    if (patch[k] !== undefined) pSet(ar.ws, ar.row, colOf(ar.cols, k), D(patch[k]));
  });
  return "ok";
}

function patchInvAddTicket(wb, fp, args) {
  const t = args[0] || {};
  let sm = sheetMetaFor(fp, "support");
  let ws, cols, sheetName;
  if (sm) { ws = pSheet(wb, sm.name); cols = sm.cols; sheetName = sm.name; }
  if (!ws) {
    const headers = ["Ticket No", "Serial Number", "Client / Organization", "Model", "Category", "Priority", "Status", "Requested By", "Date Reported To EdTech", "Date Reported To Lenovo Support", "Date Completed", "Person In-Charge", "Issue/s", "Activities / Troubleshooting", "Follow-up Notes", "Related Tickets", "Service Location Address"];
    const made = pEnsureSheet(wb, fp, "support", "03 TECH SUPPORT LOGS", headers);
    ws = made.ws; sheetName = made.name;
    cols = { tno: 0, sn: 1, client: 2, model: 3, cat: 4, priority: 5, status: 6, requester: 7, repEdtech: 8, repLenovo: 9, completed: 10, pic: 11, issue: 12, act: 13, notes: 14, related: 15, svcAddr: 16 };
  }
  // idempotency: ticket already has a source row in this file
  let ticket = null;
  try {
    if (t.tno && typeof QBR.invFindTicket === "function") ticket = QBR.invFindTicket(t.tno);
    if (!ticket && typeof invModel === "function") {
      // fallback for journals recorded before tno was stamped: match an
      // _src-less ticket by serial + issue text
      const inv = invModel();
      ticket = ((inv && inv.tickets) || []).find(x => !x._src &&
        String(x.sn || "") === String(t.sn || "").trim() &&
        String(x.issue || "") === String(t.issue || ""));
    }
  } catch (e) {}
  if (ticket && ticket._src && ticket._src.fp === fp) return "noop";
  const D = v => (typeof invDate === "function" ? invDate(v) : null);
  const oldLast = pLastRow(ws);
  const cells = [];
  const put = (k, v) => { const c = colOf(cols, k); if (c >= 0) cells[c] = v; };
  put("tno", t.tno || "");
  put("sn", String(t.sn || "").trim());
  put("client", t.client || "");
  put("model", t.model || "");
  put("cat", t.cat || "Hardware");
  put("priority", t.priority || "");
  put("status", t.status || "Open");
  put("requester", t.requester || "");
  put("repEdtech", D(t.repEdtech) || new Date());
  put("repLenovo", D(t.repLenovo));
  put("pic", t.pic || "");
  put("issue", t.issue || "");
  put("act", t.act || "");
  put("notes", Array.isArray(t.notes) ? t.notes.map(n => {
    const d = n.d instanceof Date ? n.d : D(n.d);
    return (d ? d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") : "—") + ": " + (n.text || "");
  }).join("\n") : "");
  put("related", Array.isArray(t.related) ? t.related.join(", ") : "");
  put("svcAddr", t.svcAddr || "");
  const r = pAppendRow(ws, cells);
  pBumpRanges(wb, sheetName, oldLast, r);
  if (ticket) ticket._src = { fp, sheet: sheetName, row: r };
  return "ok";
}

function patchInvUpdateTicket(wb, fp, args) {
  const tno = args[0], patch = args[1] || {};
  const tr = patchTicketRow(wb, fp, tno);
  if (!tr) return "skip: ticket row not in linked file";
  const map = { client: "client", model: "model", cat: "cat", priority: "priority", requester: "requester", pic: "pic", issue: "issue", act: "act", svcAddr: "svcAddr" };
  Object.keys(map).forEach(k => {
    if (patch[k] !== undefined) pSet(tr.ws, tr.row, colOf(tr.cols, map[k]), patch[k] == null || patch[k] === "" ? "" : patch[k]);
  });
  const D = v => (typeof invDate === "function" ? invDate(v) : null);
  if (patch.repEdtech !== undefined) pSet(tr.ws, tr.row, colOf(tr.cols, "repEdtech"), D(patch.repEdtech));
  if (patch.repLenovo !== undefined) pSet(tr.ws, tr.row, colOf(tr.cols, "repLenovo"), D(patch.repLenovo));
  return "ok";
}

function patchTicketNoteCell(wb, fp, tno, text) {
  const tr = patchTicketRow(wb, fp, tno);
  if (!tr) return "skip: ticket row not in linked file";
  const c = colOf(tr.cols, "notes");
  if (c < 0) return "skip: no notes column";
  const addr = XLSX.utils.encode_cell({ r: tr.row - 1, c });
  const cell = tr.ws[addr];
  if (cell && cell.f) return "skip: notes cell holds a formula";
  const existing = cell && cell.v != null ? String(cell.v) : "";
  const line = pTodayYMD() + ": " + String(text || "").trim();
  pSet(tr.ws, tr.row, c, existing ? existing + "\n" + line : line);
  return "ok";
}
function patchInvTicketNote(wb, fp, args) { return patchTicketNoteCell(wb, fp, args[0], args[1]); }

function patchInvLinkTicket(wb, fp, args) {
  const tno = args[0], rel = String(args[1] || "").trim().toUpperCase();
  if (!rel) return "skip: blank link";
  const tr = patchTicketRow(wb, fp, tno);
  if (!tr) return "skip: ticket row not in linked file";
  const c = colOf(tr.cols, "related");
  if (c < 0) return "skip: no related column";
  const addr = XLSX.utils.encode_cell({ r: tr.row - 1, c });
  const cell = tr.ws[addr];
  if (cell && cell.f) return "skip: related cell holds a formula";
  const cur = cell && cell.v != null ? String(cell.v).split(/[,;\n]+/).map(s => s.trim()).filter(Boolean) : [];
  if (!cur.map(s => s.toUpperCase()).includes(rel)) cur.push(rel);
  pSet(tr.ws, tr.row, c, cur.join(", "));
  return "ok";
}

function patchInvSetTicketStatus(wb, fp, args) {
  const tno = args[0], status = args[1], note = args[2];
  const tr = patchTicketRow(wb, fp, tno);
  if (!tr) return "skip: ticket row not in linked file";
  pSet(tr.ws, tr.row, colOf(tr.cols, "status"), status || "");
  if (/^resolved$/i.test(String(status || ""))) {
    const done = tr.ticket.completed instanceof Date ? tr.ticket.completed : new Date();
    pSet(tr.ws, tr.row, colOf(tr.cols, "completed"), done);
  }
  if (/^open$/i.test(String(status || ""))) {
    pSet(tr.ws, tr.row, colOf(tr.cols, "completed"), "");
  }
  if (note && String(note).trim()) {
    patchTicketNoteCell(wb, fp, tno, "Status → " + status + ": " + String(note).trim());
  }
  return "ok";
}

/* legacy journal op (pre-1.13): invResolveTicket(key, n, note) */
function patchInvResolveTicket(wb, fp, args) {
  const key = args[0], n = args[1], note = args[2];
  let ticket = null;
  try {
    const inv = pModel("assets");
    const sk = (typeof QBR.invSerialKey === "function") ? QBR.invSerialKey(key) : key;
    ticket = ((inv && inv.tickets) || []).find(t => t._n === Number(n) && t.key === sk);
  } catch (e) {}
  if (!ticket || !ticket.tno) return "skip: ticket not found";
  return patchInvSetTicketStatus(wb, fp, [ticket.tno, "Resolved", note]);
}

/* statusOverride is dashboard-model-only (no sheet column): nothing to write */
function patchInvSetStatus() { return "noop"; }

/* invAddDeployment (2026-10-05 bulk SQ): append one deployment row to 04 RAKSO
 * INV. for a unit that lives in THIS linked file. Never edits existing rows. */
function patchInvAddDeployment(wb, fp, args) {
  const d = args[0] || {};
  const key = (typeof QBR.invSerialKey === "function") ? QBR.invSerialKey(d.sn || "") : String(d.sn || "").toUpperCase();
  const inv = pModel("assets");
  const a = ((inv && inv.assets) || []).find(x => x.key === key);
  if (!a || !a._src || a._src.fp !== fp) return "skip: unit not in linked file";
  const D = v => (typeof invDate === "function" ? invDate(v) : null);
  const dep = ((inv && inv.deployments) || []).filter(x => !x._src && x.key === key && String(x.sq) === String(d.sq)).pop()
    || { key, sn: a.sn, type: a.cat, desc: a.desc, sq: d.sq, client: d.client || a.client, req: null, delivered: D(d.delivered), remarks: d.remarks || "SQ set via bulk edit" };
  if (!(dep.delivered instanceof Date)) dep.delivered = D(dep.delivered);
  patchAppendDeployment(wb, fp, dep);
  return "ok";
}

var PATCH_OPS = {
  supRecordTransaction: patchSupRecordTransaction,
  supAddItem: patchSupAddItem,
  invIntake: patchInvIntake,
  invDeploy: patchInvDeploy,
  invAddTicket: patchInvAddTicket,
  invUpdateTicket: patchInvUpdateTicket,
  invTicketNote: patchInvTicketNote,
  invLinkTicket: patchInvLinkTicket,
  invSetTicketStatus: patchInvSetTicketStatus,
  invResolveTicket: patchInvResolveTicket,
  invSetStatus: patchInvSetStatus,
  invUpdateAsset: patchInvUpdateAsset,
  invAddDeployment: patchInvAddDeployment,
};

/* ============================ edit log ================================== */
function pLogDetail(op, args) {
  try {
    const a = args || [];
    switch (op) {
      case "supRecordTransaction": return "transaction: " + (a[0].type || "Used") + " " + a[0].qty + " × " + a[0].itemId;
      case "supAddItem": return "new item: " + (a[0].code ? a[0].code + "-" : "") + (a[0].name || "");
      case "invIntake": return "intake: " + (a[0] || []).length + " unit(s)";
      case "invDeploy": return "deploy: " + (a[0] || []).length + " unit(s) → " + (a[1] || "—");
      case "invAddTicket": return "ticket opened: " + (a[0].tno || "") + " — " + (a[0].sn || "") + " — " + (a[0].issue || "");
      case "invUpdateTicket": return "ticket updated: " + a[0];
      case "invTicketNote": return "ticket note: " + a[0];
      case "invLinkTicket": return "ticket linked: " + a[0] + " ↔ " + a[1];
      case "invSetTicketStatus": return "ticket status: " + a[0] + " → " + a[1];
      case "invResolveTicket": return "ticket resolved: " + a[0];
      case "invSetStatus": return "status override: " + a[0] + " → " + a[1];
      case "invUpdateAsset": return "asset updated: " + a[0] + " (" + Object.keys(a[1] || {}).join(", ") + ")";
      case "invAddDeployment": return "SQ: " + (a[0].sn || "") + " → " + (a[0].sq || "");
      default: return op;
    }
  } catch (e) { return op; }
}
function pAppendEditLog(wb, entries) {
  if (!entries.length) return;
  let ws = pSheet(wb, "EDIT LOG");
  if (!ws) {
    ws = XLSX.utils.aoa_to_sheet([["Timestamp", "Action", "Detail"]]);
    XLSX.utils.book_append_sheet(wb, ws, "EDIT LOG");
  }
  entries.forEach(([ts, action, detail]) => {
    pAppendRow(ws, [ts instanceof Date ? ts : new Date(), action, detail]);
  });
}

/* ============================ range bumper ============================== */
/* Conservative: in every sheet except `srcSheet`, extend bounded ranges that
 * reference srcSheet and end EXACTLY at oldLast (1-indexed), e.g.
 * Items!$A$2:$A$5 → Items!$A$2:$A$6 after appending row 6. Anything else is
 * left untouched. */
function pBumpRanges(wb, srcSheet, oldLast, newLast) {
  if (!srcSheet || !(newLast > oldLast)) return 0;
  const esc = srcSheet.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  // (^|boundary) guards against matching a longer sheet name (e.g. "XItems")
  const re = new RegExp("(^|[^A-Za-z0-9_.'])('?)" + esc + "('?)!(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+):(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)", "g");
  let bumps = 0;
  (wb.SheetNames || []).forEach(n => {
    if (n === srcSheet) return;
    const ws = wb.Sheets[n];
    Object.keys(ws).forEach(addr => {
      if (addr[0] === "!") return;
      const cell = ws[addr];
      if (!cell || !cell.f) return;
      const f = String(cell.f);
      let hit = false;
      const nf = f.replace(re, (m, bnd, q1, q2, d1, c1, d2, r1, d3, c2, d4, r2) => {
        if (r1 === "2" && Number(r2) === oldLast) {
          hit = true;
          return bnd + q1 + srcSheet + q2 + "!" + d1 + c1 + d2 + r1 + ":" + d3 + c2 + d4 + newLast;
        }
        return m;
      });
      if (hit) { cell.f = nf; bumps++; }
    });
  });
  return bumps;
}

/* ============================ main entry ================================ */
/* Replay the mutation journal for (kind, fp) as cell writes into the ORIGINAL
 * workbook. Returns {ok, applied, skipped, appended, notes[]}. */
QBR.patchWorkbookFromJournal = function (kind, wb, fp) {
  const report = { ok: true, applied: 0, skipped: 0, appended: 0, bumped: 0, notes: [] };
  if (!wb || !fp || typeof XLSX === "undefined") return Object.assign(report, { ok: false, error: "no workbook" });
  let store = {};
  try { store = JSON.parse(localStorage.getItem("qbr-inv-journal-v1")) || {}; }
  catch (e) { return Object.assign(report, { ok: false, error: "no journal" }); }
  const entry = store[fp];
  const ops = entry && entry.ops ? entry.ops.filter(o => o && o.kind === kind) : [];
  const logRows = [];
  ops.forEach(o => {
    const fn = PATCH_OPS[o.op];
    if (typeof fn !== "function") { report.skipped++; report.notes.push(o.op + ": unknown op"); return; }
    let res;
    try { res = fn(wb, fp, o.args || []); }
    catch (err) { res = "skip: " + ((err && err.message) || err); }
    if (res === "ok" || res === "noop") {
      report.applied++;
      if (res === "ok") logRows.push([new Date(), o.op, pLogDetail(o.op, o.args)]);
    } else {
      report.skipped++;
      report.notes.push(o.op + ": " + res);
    }
  });
  if (logRows.length) {
    try { pAppendEditLog(wb, logRows); } catch (e) { report.notes.push("edit-log: " + (e && e.message)); }
  }
  return report;
};

/* After a successful patch save the file bytes changed: carry the retained
 * workbook + sheet metadata to the new fingerprint (coordinates are unchanged;
 * appends only extend sheets). Called from persist.js fsAfterSave. */
QBR.fsCarryPatchState = function (oldFp, newFp) {
  if (!oldFp || !newFp || oldFp === newFp) return;
  try {
    if (QBR._origWb && QBR._origWb[oldFp]) {
      QBR._origWb[newFp] = QBR._origWb[oldFp];
      delete QBR._origWb[oldFp];
    }
    if (QBR._sheetMeta && QBR._sheetMeta[oldFp]) {
      QBR._sheetMeta[newFp] = QBR._sheetMeta[oldFp];
      delete QBR._sheetMeta[oldFp];
    }
  } catch (e) {}
};

/* Drop retained patch state (used when falling back to a full rebuild, whose
 * layout no longer matches recorded coordinates). */
QBR.fsDropPatchState = function (fp) {
  if (!fp) return;
  try {
    if (QBR._origWb) delete QBR._origWb[fp];
    if (QBR._sheetMeta) delete QBR._sheetMeta[fp];
  } catch (e) {}
};

/* ---------- Node export guard (test harness only; inert in browser) ------ */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { QBR, PATCH_OPS, pSet, pAppendRow, pLastRow, pBumpRanges, pSheet };
}
