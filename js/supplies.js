/* ============================================================================
 * QBR Supplies module — consumables / quantity-based inventory (ETG printer)
 * v1.0.0 · 2026-10-04
 *
 * Companion to js/inventory.js, which is SERIALIZED-asset based (one row per
 * physical unit, keyed by serial number). This module is QUANTITY based:
 * items keyed by ItemID, movements recorded in a transaction ledger, stock
 * position recomputed as  Ending = Beginning + Received − Used + Adjustments.
 *
 * Source workbook shape (ETG_PRINTER_INVENTORY.xlsx):
 *   Items               ItemID (=CONCAT(item-code,"-",name)), ItemName,
 *                       Category, Unit, BeginningBalance, ReorderLevel,
 *                       Supplier, UnitCost, Location (+ K=EMPLOYEE / L=ITEM /
 *                       M=name helper columns; M has no header)
 *   Transactions        Date, DocNo, TransactionType (Received/Used/Adjust),
 *                       ItemID, Quantity, Department/Project, RequestedBy,
 *                       ApprovedBy, Notes
 *   Inventory Summary   formula-only in Excel; RECOMPUTED here in JS so the
 *                       dashboard never depends on cached formula values.
 *
 * Loads AFTER js/inventory.js. Hangs off window.QBR. Conventions mirror the
 * inventory module: esc() before innerHTML, "—"/"No data" never 0, pure
 * parsers are Node-testable (guard at bottom). Nothing is persisted; edits
 * live in memory + QBR.invJournal and are exported as a regenerated workbook.
 * ==========================================================================*/
var QBR = window.QBR = window.QBR || {};

QBR.SUP_VERSION = "1.0.0";

/* Model access — guarded: top-level `const APP` is NOT window.APP. */
function supModel() {
  return (typeof APP !== "undefined" && APP.model && APP.model.supplies) || null;
}

/* ---------- sheet resolution -------------------------------------------- */
const SUP_SHEETS = {
  items:   ["ITEMS"],
  trans:   ["TRANSACTIONS"],
  summary: ["INVENTORY SUMMARY", "SUMMARY"],
};
function supFindSheet(sheets, key) {
  const names = Object.keys(sheets || {});
  for (const cand of SUP_SHEETS[key]) {
    const hit = names.find(n => String(n).trim().toUpperCase() === cand);
    if (hit) return { name: hit, rows: sheets[hit] };
  }
  // substring fallback
  for (const cand of SUP_SHEETS[key]) {
    const hit = names.find(n => String(n).trim().toUpperCase().replace(/\s+/g, "").indexOf(cand.replace(/\s+/g, "")) !== -1);
    if (hit) return { name: hit, rows: sheets[hit] };
  }
  return null;
}

/* Source-coordinate helpers (recSheetMeta/srcOf) live in inventory.js, which
 * always loads first — same as invFmtDate/invToday used below. */

/* ---------- keys / normalization ---------------------------------------- */
QBR.supItemKey = function (s) { return String(s == null ? "" : s).trim(); };

function supNormType(t) {
  const v = String(t == null ? "" : t).trim().toLowerCase();
  if (!v) return "";
  if (v.indexOf("receiv") === 0) return "Received";
  if (v.indexOf("adjust") === 0) return "Adjust";
  if (v.indexOf("used") === 0 || v.indexOf("issu") === 0 || v.indexOf("consume") === 0) return "Used";
  return String(t).trim();
}

function supDate(v) {
  if (typeof invDate === "function") return invDate(v);
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  const t = Date.parse(String(v));
  return isNaN(t) ? null : new Date(t);
}

/* ---------- pure row parsers (Node-testable) ----------------------------- */
function supColIndex(header, cands) {
  if (typeof makeResolver === "function") {
    const idx = makeResolver(header);
    return idx(cands);
  }
  const norm = h => String(h == null ? "" : h).toLowerCase().replace(/[^a-z0-9]/g, "");
  for (let i = 0; i < header.length; i++) {
    const h = norm(header[i]);
    if (cands.some(c => h === norm(c) || h.indexOf(norm(c)) !== -1)) return i;
  }
  return -1;
}

function supParseItemRows(rows, meta) {
  const out = [];
  if (!rows || !rows.length) return out;
  const H = rows[0];
  const c = {
    id: supColIndex(H, ["itemid"]), name: supColIndex(H, ["itemname"]),
    cat: supColIndex(H, ["category"]), unit: supColIndex(H, ["unit"]),
    beg: supColIndex(H, ["beginningbalance", "beginning"]),
    reorder: supColIndex(H, ["reorderlevel", "reorder"]),
    supplier: supColIndex(H, ["supplier"]), cost: supColIndex(H, ["unitcost"]),
    loc: supColIndex(H, ["location"]),
  };
  recSheetMeta(meta, c, "items");
  const num = v => (v == null || v === "" ? 0 : (isFinite(Number(v)) ? Number(v) : 0));
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : "");
    const name = get(c.name);
    if (!name) continue; // blank template row
    // ItemID: rebuild from the helper columns (L=ITEM code, M=name part),
    // exactly like the workbook's =CONCAT(L#,"-",M#) formula; fall back to col A.
    const code = row[11] != null ? String(row[11]).trim() : "";
    const namePart = row[12] != null ? String(row[12]).trim() : "";
    let id = get(c.id);
    if (code && namePart) id = code + "-" + namePart;
    if (!id) id = code || name;
    out.push({
      key: QBR.supItemKey(id), id, name,
      cat: get(c.cat) || "—", unit: get(c.unit) || "pcs",
      beginning: num(row[c.beg]), reorder: num(row[c.reorder]),
      supplier: get(c.supplier) || null,
      cost: row[c.cost] != null && row[c.cost] !== "" && isFinite(Number(row[c.cost])) ? Number(row[c.cost]) : 0,
      loc: get(c.loc) || null, _src: srcOf(meta, r),
    });
  }
  return out;
}

function supParseTransRows(rows, meta) {
  const out = [];
  if (!rows || !rows.length) return out;
  const H = rows[0];
  const c = {
    date: supColIndex(H, ["date"]), doc: supColIndex(H, ["docno", "doc no"]),
    type: supColIndex(H, ["transactiontype"]), item: supColIndex(H, ["itemid"]),
    qty: supColIndex(H, ["quantity"]), dept: supColIndex(H, ["department/project", "department"]),
    req: supColIndex(H, ["requestedby"]), app: supColIndex(H, ["approvedby"]),
    notes: supColIndex(H, ["notes"]),
  };
  recSheetMeta(meta, c, "trans");
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const get = i => (i >= 0 && row[i] != null ? String(row[i]).trim() : "");
    const itemId = get(c.item);
    if (!itemId) continue; // blank template row
    const q = Number(row[c.qty]);
    out.push({
      date: supDate(row[c.date]), doc: get(c.doc) || null,
      type: supNormType(row[c.type]) || "Used",
      key: QBR.supItemKey(itemId), itemId,
      qty: isFinite(q) ? q : 0,
      dept: get(c.dept) || null, requestedBy: get(c.req) || null,
      approvedBy: get(c.app) || null, notes: get(c.notes) || null, _src: srcOf(meta, r),
    });
  }
  return out;
}

/* ---------- multi-workbook accumulation ---------------------------------- */
QBR.supParseSheets = function (sheets, ctx) {
  const fp = ctx && ctx.fp;
  const acc = QBR._supAcc || (QBR._supAcc = { items: [], transactions: [], sources: {} });
  const metaFor = g => (g && fp) ? { fp, sheet: g.name } : null;
  const it = supFindSheet(sheets, "items");
  if (it) { acc.items.push(...supParseItemRows(it.rows, metaFor(it))); acc.sources.items = true; }
  const tr = supFindSheet(sheets, "trans");
  if (tr) { acc.transactions.push(...supParseTransRows(tr.rows, metaFor(tr))); acc.sources.trans = true; }
  return acc;
};

QBR.parseSuppliesBuffers = function (buffers) {
  QBR._supAcc = null;
  const fps = QBR._currentFps || [];
  (buffers || []).forEach((buf, bi) => {
    try {
      const wb = XLSX.read(buf, { type: "array", cellStyles: true }); // cellStyles: keep !cols/styles for patch-in-place save
      const fp = fps[bi];
      if (fp) { (QBR._origWb || (QBR._origWb = {}))[fp] = wb; }
      const sheets = {};
      wb.SheetNames.forEach(n => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false }); });
      const nBefore = QBR._supAcc ? Object.keys(QBR._supAcc.sources).length : 0;
      QBR.supParseSheets(sheets, { fp: fp || null });
      const nAfter = QBR._supAcc ? Object.keys(QBR._supAcc.sources).length : 0;
      if (nAfter > nBefore && fps[bi]) {
        const set = (QBR._kindByFp || (QBR._kindByFp = {}))[fps[bi]] ||
          (QBR._kindByFp[fps[bi]] = new Set());
        set.add("supplies");
      }
    } catch (e) { /* a non-supplies workbook simply contributes nothing */ }
  });
  const acc = QBR._supAcc || { items: [], transactions: [], sources: {} };
  QBR._supAcc = null;
  // de-dupe items by key (newest file wins); transactions accumulate
  const seen = new Map();
  acc.items.forEach(it => seen.set(it.key, it));
  acc.items = [...seen.values()];
  return acc;
};

/* ---------- summary engine (recomputed; never trusts Excel caches) ------- */
QBR.supComputeSummary = function (sup) {
  const per = new Map(); // key -> {item, received, used, adjust, ending, totalValue, lastUpdated, restock}
  (sup.items || []).forEach(it => {
    per.set(it.key, { item: it, received: 0, used: 0, adjust: 0, ending: it.beginning,
                      totalValue: 0, lastUpdated: null, restock: false });
  });
  const unmatched = [];
  (sup.transactions || []).forEach(t => {
    const e = per.get(t.key);
    if (!e) {
      if (!unmatched.some(u => u.key === t.key)) unmatched.push({ key: t.key, itemId: t.itemId, count: 0 });
      const u = unmatched.find(x => x.key === t.key); if (u) u.count++;
      return;
    }
    if (t.type === "Received") e.received += t.qty;
    else if (t.type === "Adjust") e.adjust += t.qty;
    else e.used += t.qty;
    if (t.date && (!e.lastUpdated || t.date > e.lastUpdated)) e.lastUpdated = t.date;
  });
  per.forEach(e => {
    e.ending = e.item.beginning + e.received - e.used + e.adjust;
    e.totalValue = e.ending * (e.item.cost || 0);
    e.restock = e.ending < e.item.reorder;
  });
  return { per, unmatched };
};

/* ---------- flags ---------------------------------------------------------- */
QBR.supComputeFlags = function (sup, summary) {
  const flags = { restock: [], negative: [], unknownItems: [] };
  summary.per.forEach(e => {
    if (e.ending < 0) flags.negative.push(e);
    else if (e.restock) flags.restock.push(e);
  });
  flags.unknownItems = summary.unmatched;
  return flags;
};

/* ---------- write-back (in-memory + journal + re-render) ------------------- */
QBR.supRecordTransaction = function (t) {
  const sup = supModel(); if (!sup) return false;
  // Normalize the date: journal replay delivers ISO strings (JSON), the form
  // delivers Dates. A raw string here used to crash the render with
  // "d.getFullYear is not a function".
  const date = (typeof supDate === "function" ? supDate(t.date) : null) ||
    (typeof invToday === "function" ? invToday() : new Date());
  sup.transactions.push({
    date: date,
    doc: t.doc || null, type: supNormType(t.type) || "Used",
    key: QBR.supItemKey(t.itemId), itemId: String(t.itemId || "").trim(),
    qty: isFinite(Number(t.qty)) ? Number(t.qty) : 0,
    dept: t.dept || null, requestedBy: t.requestedBy || null,
    approvedBy: t.approvedBy || null, notes: t.notes || null,
  });
  if (typeof QBR.invLog === "function") QBR.invLog("supplies transaction", (t.type || "Used") + " " + t.qty + " × " + t.itemId);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("supplies", "supRecordTransaction", [t]);
  return true;
};

QBR.supAddItem = function (it) {
  const sup = supModel(); if (!sup) return false;
  const id = ((it.code || "").trim() + "-" + (it.name || "").trim()).replace(/^-|-$/g, "") || (it.name || "").trim();
  if (!id) return false;
  const key = QBR.supItemKey(id);
  if (sup.items.some(x => x.key === key)) return false;
  sup.items.push({
    key, id, name: (it.name || "").trim(), cat: (it.cat || "").trim() || "—",
    unit: (it.unit || "").trim() || "pcs", beginning: Number(it.beginning) || 0,
    reorder: Number(it.reorder) || 0, supplier: (it.supplier || "").trim() || null,
    cost: Number(it.cost) || 0, loc: (it.loc || "").trim() || null,
  });
  if (typeof QBR.invLog === "function") QBR.invLog("supplies new item", id);
  if (typeof QBR.journalRecord === "function") QBR.journalRecord("supplies", "supAddItem", [it]);
  return true;
};

/* ---------- export: regenerate the workbook -------------------------------- */
// Build the supplies workbook object (shared by Export download and direct save).
QBR.supBuildWorkbook = function () {
  const sup = supModel(); if (!sup || typeof XLSX === "undefined") return null;
  const F = (typeof invFmtDate === "function") ? invFmtDate : (d => d ? d.toISOString().slice(0, 10) : "");
  // Items — col A keeps the live =CONCAT(L#,"-",M#) formula (with cached value)
  const itemHead = ["ItemID", "ItemName", "Category", "Unit", "BeginningBalance", "ReorderLevel",
                    "Supplier", "UnitCost", "Location", "", "EMPLOYEE", "ITEM", ""];
  const itemAoa = [itemHead];
  sup.items.forEach((it, i) => {
    const r = i + 2, code = it.id.split("-")[0], nm = it.id.slice(code.length + 1) || it.name;
    itemAoa.push([it.id, it.name, it.cat === "—" ? "" : it.cat, it.unit, it.beginning, it.reorder,
                  it.supplier || "", it.cost, it.loc || "", "", "", code, nm]);
    itemAoa[itemAoa.length - 1]._formulaRow = r;
  });
  const wsI = XLSX.utils.aoa_to_sheet(itemAoa.map(r => r.slice(0, 13)));
  sup.items.forEach((it, i) => {
    const r = i + 2, code = it.id.split("-")[0];
    wsI["A" + r] = { t: "s", f: 'CONCAT(L' + r + ',"-",M' + r + ')', v: it.id };
    wsI["L" + r] = { t: "s", v: code };
  });
  // Transactions
  const trHead = ["Date", "DocNo", "TransactionType", "ItemID", "Quantity", "Department/Project",
                  "RequestedBy", "ApprovedBy", "Notes"];
  const trAoa = [trHead];
  sup.transactions.forEach(t => trAoa.push([F(t.date), t.doc || "", t.type, t.itemId, t.qty,
                                            t.dept || "", t.requestedBy || "", t.approvedBy || "", t.notes || ""]));
  // Inventory Summary — formulas over plain ranges (no Excel Tables needed)
  const nI = sup.items.length, nT = sup.transactions.length;
  const lastI = Math.max(2, nI + 1), lastT = Math.max(2, nT + 1);
  const sumHead = ["ItemID", "ItemName", "Unit", "Beginning", "Received", "Used", "Adjustments",
                   "Ending", "ReorderLevel", "UnitCost", "TotalValue", "LastUpdated", "REMARKS"];
  const sumAoa = [sumHead];
  sup.items.forEach((it, i) => {
    const r = i + 2;
    sumAoa.push([
      { t: "s", f: "Items!A" + r, v: it.id },
      { t: "s", f: 'XLOOKUP($A' + r + ",Items!$A$2:$A$" + lastI + ",Items!$B$2:$B$" + lastI + ',"")', v: it.name },
      { t: "s", f: 'XLOOKUP($A' + r + ",Items!$A$2:$A$" + lastI + ",Items!$D$2:$D$" + lastI + ',"")', v: it.unit },
      { t: "n", f: 'XLOOKUP($A' + r + ",Items!$A$2:$A$" + lastI + ",Items!$E$2:$E$" + lastI + ",0)", v: it.beginning },
      { t: "n", f: 'SUMIFS(Transactions!$E$2:$E$' + lastT + ',Transactions!$C$2:$C$' + lastT + ',"Received",Transactions!$D$2:$D$' + lastT + ",$A" + r + ")", v: 0 },
      { t: "n", f: 'SUMIFS(Transactions!$E$2:$E$' + lastT + ',Transactions!$C$2:$C$' + lastT + ',"Used",Transactions!$D$2:$D$' + lastT + ",$A" + r + ")", v: 0 },
      { t: "n", f: 'SUMIFS(Transactions!$E$2:$E$' + lastT + ',Transactions!$C$2:$C$' + lastT + ',"Adjust",Transactions!$D$2:$D$' + lastT + ",$A" + r + ")", v: 0 },
      { t: "n", f: "=D" + r + "+E" + r + "-F" + r + "+G" + r, v: 0 },
      { t: "n", f: 'XLOOKUP($A' + r + ",Items!$A$2:$A$" + lastI + ",Items!$F$2:$F$" + lastI + ",0)", v: it.reorder },
      { t: "n", f: 'XLOOKUP($A' + r + ",Items!$A$2:$A$" + lastI + ",Items!$H$2:$H$" + lastI + ",0)", v: it.cost },
      { t: "n", f: "=H" + r + "*J" + r, v: 0 },
      { t: "s", f: 'MAXIFS(Transactions!$A$2:$A$' + lastT + ",Transactions!$D$2:$D$" + lastT + ",$A" + r + ")", v: "" },
      { t: "s", f: '=IF($I' + r + ">$H" + r + ',"Need to restock, order now","")', v: "" },
    ]);
  });
  const wsS = XLSX.utils.aoa_to_sheet(sumHead.map(() => ""));
  XLSX.utils.sheet_add_aoa(wsS, [sumHead], { origin: "A1" });
  sumAoa.slice(1).forEach((row, i) => {
    const r = i + 2, cols = "ABCDEFGHIJKLM";
    row.forEach((cell, j) => { wsS[cols[j] + r] = cell; });
  });
  wsS["!ref"] = "A1:M" + (nI + 1); // direct cell assignment doesn't expand !ref; writer needs it
  const wb = XLSX.utils.book_new();
  [["Items", wsI], ["Transactions", XLSX.utils.aoa_to_sheet(trAoa)], ["Inventory Summary", wsS]]
    .forEach(([name, ws]) => XLSX.utils.book_append_sheet(wb, ws, name));
  return { wb: wb, filename: "ETG_PRINTER_INVENTORY_export.xlsx" };
};

QBR.supExportWorkbook = function () {
  const b = QBR.supBuildWorkbook();
  if (!b) return;
  XLSX.writeFile(b.wb, b.filename, { cellStyles: true });
};

/* ---------- UI state -------------------------------------------------------- */
QBR._supUI = { type: "ALL", q: "", form: null, showAllTx: false };

/* ---------- supplies dashboard -------------------------------------------- */
QBR.supRender = function (host) {
  const sup = supModel();
  if (!host) return;
  if (!sup || (!sup.items.length && !sup.transactions.length)) {
    host.innerHTML = `<div class="card-box"><p class="text-muted">No supplies data loaded yet. Upload the printer inventory workbook (sheets Items, Transactions, Inventory Summary).</p></div>`;
    return;
  }
  const ui = QBR._supUI;
  const summary = QBR.supComputeSummary(sup);
  const flags = QBR.supComputeFlags(sup, summary);
  const E = (typeof esc === "function") ? esc : (s => s);
  const F = (typeof fmt === "function") ? fmt : (n => (n == null ? "—" : n));
  const FD = (typeof invFmtDate === "function") ? invFmtDate : (d => (d ? d.toISOString().slice(0, 10) : "—"));

  /* Single alert line — only rendered when something needs attention. */
  const notes = [];
  flags.restock.forEach(e => notes.push(`${E(e.item.name)} — ${F(e.ending)} ${E(e.item.unit)} left (reorder ${F(e.item.reorder)})`));
  flags.negative.forEach(e => notes.push(`${E(e.item.name)} — negative stock (${F(e.ending)})`));
  flags.unknownItems.forEach(u => notes.push(`${E(u.itemId)} — not in Items (${F(u.count)} txn)`));
  const alerts = notes.length
    ? `<div class="alert alert-warning py-2 mt-3" role="alert"><strong>Attention needed:</strong> ${notes.join(" · ")}</div>`
    : "";

  const actions = `<div class="d-flex flex-wrap gap-2 mt-3">
      <button type="button" class="btn btn-sm btn-primary" data-sup-form="txn">＋ Record transaction</button>
      <button type="button" class="btn btn-sm btn-outline-primary" data-sup-form="item">Add item</button>
      ${(typeof QBR.saveButtonHtml === "function") ? QBR.saveButtonHtml("supplies", "sup-save") : ""}
    </div><div id="sup-form-host" class="mt-2"></div>`;

  const filters = `<div class="card-box mt-3"><div class="d-flex flex-wrap gap-2 align-items-end">
      <div><label class="form-label small mb-0" for="sup-f-type">Type</label>
        <select id="sup-f-type" class="form-select form-select-sm" style="--w:150px">
        ${["ALL", "Received", "Used", "Adjust"].map(t => `<option${ui.type === t ? " selected" : ""}>${t === "ALL" ? "All types" : t}</option>`).join("")}</select></div>
      <div><label class="form-label small mb-0" for="sup-f-q">Search item</label>
        <input id="sup-f-q" class="form-control form-control-sm" style="--w:220px" placeholder="Name or ItemID…" value="${E(ui.q)}"></div>
    </div></div>`;

  // stock table
  const q = ui.q.trim().toLowerCase();
  let rows = [...summary.per.values()];
  if (q) rows = rows.filter(e => e.item.name.toLowerCase().includes(q) || e.item.id.toLowerCase().includes(q));
  rows.sort((a, b) => (a.restock === b.restock) ? b.totalValue - a.totalValue : (a.restock ? -1 : 1));
  const stockTbl = `<div class="card-box mt-3"><h6>Stock <span class="text-muted">(${F(rows.length)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>ItemID</th><th>Item</th><th class="text-end">Remaining</th>
    <th class="text-end">Reorder at</th><th>Remarks</th></tr></thead><tbody>` +
    (rows.map(e => `<tr>
      <td><code>${E(e.item.id)}</code></td><td>${E(e.item.name)}</td>
      <td class="text-end"><strong>${F(e.ending)}</strong> ${E(e.item.unit)}</td>
      <td class="text-end">${F(e.item.reorder)}</td>
      <td>${e.ending < 0 ? `<span class="text-danger">Negative stock</span>` : e.restock ? `<span class="text-danger">Need to restock, order now</span>` : "—"}</td></tr>`).join("") ||
      `<tr><td colspan="5" class="text-muted">No items match.</td></tr>`) +
    `</tbody></table></div></div>`;

  // transaction ledger
  let tix = sup.transactions.slice().sort((a, b) => (b.date || 0) - (a.date || 0));
  if (ui.type !== "ALL") tix = tix.filter(t => t.type === ui.type);
  if (q) tix = tix.filter(t => t.itemId.toLowerCase().includes(q));
  const shown = ui.showAllTx ? tix : tix.slice(0, 100);
  const txTbl = `<div class="card-box mt-3"><h6>Transactions <span class="text-muted">(${F(tix.length)})</span></h6>
    <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
    <th>Date</th><th>Doc No</th><th>Type</th><th>Item</th><th class="text-end">Qty</th>
    <th>Dept/Project</th><th>Requested by</th><th>Notes</th></tr></thead><tbody>` +
    (shown.map(t => `<tr><td>${FD(t.date)}</td><td>${E(t.doc || "—")}</td><td>${E(t.type)}</td>
      <td><code>${E(t.itemId)}</code></td><td class="text-end">${F(t.qty)}</td>
      <td>${E(t.dept || "—")}</td><td>${E(t.requestedBy || "—")}</td><td>${E(t.notes || "—")}</td></tr>`).join("") ||
      `<tr><td colspan="8" class="text-muted">No transactions match.</td></tr>`) +
    `</tbody></table></div>` +
    (!ui.showAllTx && tix.length > 100 ? `<button type="button" class="btn btn-sm btn-outline-secondary" id="sup-showalltx">Show all ${F(tix.length)}</button>` : "") + `</div>`;

  host.innerHTML = alerts + actions + filters + stockTbl + txTbl;
  if (ui.form) supShowForm(ui.form);
  QBR.supBind(host);
};

QBR.supBind = function (host) {
  const ui = QBR._supUI;
  host.querySelectorAll("[data-sup-form]").forEach(b => b.addEventListener("click", () => {
    ui.form = ui.form === b.dataset.supForm ? null : b.dataset.supForm; QBR.supRender(host);
  }));
  const ex = host.querySelector("#sup-save");
  if (ex) ex.addEventListener("click", () => { if (typeof QBR.supSave === "function") QBR.supSave(); else QBR.supExportWorkbook(); });
  if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge();
  const sa = host.querySelector("#sup-showalltx");
  if (sa) sa.addEventListener("click", () => { ui.showAllTx = true; QBR.supRender(host); });
  const ft = host.querySelector("#sup-f-type");
  if (ft) ft.addEventListener("change", () => { ui.type = ft.value; ui.showAllTx = false; QBR.supRender(host); });
  const fq = host.querySelector("#sup-f-q");
  if (fq) fq.addEventListener("change", () => { ui.q = fq.value.trim(); ui.showAllTx = false; QBR.supRender(host); });
};

function supItemOptions(sel) {
  const sup = supModel(); if (!sup) return "";
  const E = (typeof esc === "function") ? esc : (s => s);
  return sup.items.map(it => `<option value="${E(it.id)}"${it.id === sel ? " selected" : ""}>${E(it.id)}</option>`).join("");
}

function supShowForm(which) {
  const host = document.getElementById("sup-form-host"); if (!host) return;
  const E = (typeof esc === "function") ? esc : (s => s);
  const today = (typeof invFmtDate === "function" && typeof invToday === "function") ? invFmtDate(invToday()) : new Date().toISOString().slice(0, 10);
  const close = `<button type="button" class="btn btn-sm btn-outline-secondary" data-sup-form="">Close ✕</button>`;
  if (which === "txn") {
    host.innerHTML = `<div class="card-box"><h6>Record transaction</h6>
      <div class="row g-2">
      <div class="col-md-2"><label class="form-label small">Date</label><input id="st-date" type="date" class="form-control form-control-sm" value="${today}"></div>
      <div class="col-md-2"><label class="form-label small">Type</label><select id="st-type" class="form-select form-select-sm"><option>Received</option><option>Used</option><option>Adjust</option></select></div>
      <div class="col-md-4"><label class="form-label small">Item</label><select id="st-item" class="form-select form-select-sm">${supItemOptions()}</select></div>
      <div class="col-md-2"><label class="form-label small">Quantity</label><input id="st-qty" type="number" min="1" class="form-control form-control-sm" value="1"></div>
      <div class="col-md-2"><label class="form-label small">Doc No</label><input id="st-doc" class="form-control form-control-sm"></div>
      <div class="col-md-3"><label class="form-label small">Department / Project</label><input id="st-dept" class="form-control form-control-sm" value="ETG"></div>
      <div class="col-md-3"><label class="form-label small">Requested by</label><input id="st-req" class="form-control form-control-sm"></div>
      <div class="col-md-3"><label class="form-label small">Approved by</label><input id="st-app" class="form-control form-control-sm"></div>
      <div class="col-md-3"><label class="form-label small">Notes</label><input id="st-notes" class="form-control form-control-sm"></div>
      </div><div class="d-flex gap-2 mt-2">
      <button type="button" class="btn btn-sm btn-primary" id="st-go">Record</button>${close}</div>
      <div id="st-msg" class="small text-muted mt-1"></div></div>`;
    host.querySelector("#st-go").addEventListener("click", () => {
      const itemId = host.querySelector("#st-item").value, qty = Number(host.querySelector("#st-qty").value);
      const msg = host.querySelector("#st-msg");
      if (!itemId) { msg.textContent = "Pick an item."; return; }
      if (!isFinite(qty) || qty <= 0) { msg.textContent = "Quantity must be positive."; return; }
      const dv = host.querySelector("#st-date").value;
      QBR.supRecordTransaction({
        date: dv ? new Date(dv + "T00:00:00") : null, type: host.querySelector("#st-type").value,
        itemId, qty, doc: host.querySelector("#st-doc").value.trim(),
        dept: host.querySelector("#st-dept").value.trim(), requestedBy: host.querySelector("#st-req").value.trim(),
        approvedBy: host.querySelector("#st-app").value.trim(), notes: host.querySelector("#st-notes").value.trim(),
      });
      QBR._supUI.form = null; QBR.supRender(document.getElementById("sup-body"));
    });
  } else if (which === "item") {
    host.innerHTML = `<div class="card-box"><h6>Add item</h6>
      <div class="row g-2">
      <div class="col-md-2"><label class="form-label small">Item code</label><input id="si-code" class="form-control form-control-sm" placeholder="ITM-0005"></div>
      <div class="col-md-4"><label class="form-label small">Item name</label><input id="si-name" class="form-control form-control-sm"></div>
      <div class="col-md-2"><label class="form-label small">Category</label><input id="si-cat" class="form-control form-control-sm" value="Consumables"></div>
      <div class="col-md-2"><label class="form-label small">Unit</label><input id="si-unit" class="form-control form-control-sm" value="pcs"></div>
      <div class="col-md-2"><label class="form-label small">Beginning</label><input id="si-beg" type="number" class="form-control form-control-sm" value="0"></div>
      <div class="col-md-2"><label class="form-label small">Reorder level</label><input id="si-reorder" type="number" class="form-control form-control-sm" value="0"></div>
      <div class="col-md-2"><label class="form-label small">Unit cost</label><input id="si-cost" type="number" step="0.01" class="form-control form-control-sm" value="0"></div>
      <div class="col-md-3"><label class="form-label small">Supplier</label><input id="si-sup" class="form-control form-control-sm"></div>
      <div class="col-md-3"><label class="form-label small">Location</label><input id="si-loc" class="form-control form-control-sm" value="ETG Printer"></div>
      </div><div class="d-flex gap-2 mt-2">
      <button type="button" class="btn btn-sm btn-primary" id="si-go">Add item</button>${close}</div>
      <div id="si-msg" class="small text-muted mt-1"></div></div>`;
    host.querySelector("#si-go").addEventListener("click", () => {
      const msg = host.querySelector("#si-msg");
      const ok = QBR.supAddItem({
        code: host.querySelector("#si-code").value, name: host.querySelector("#si-name").value,
        cat: host.querySelector("#si-cat").value, unit: host.querySelector("#si-unit").value,
        beginning: host.querySelector("#si-beg").value, reorder: host.querySelector("#si-reorder").value,
        cost: host.querySelector("#si-cost").value, supplier: host.querySelector("#si-sup").value,
        loc: host.querySelector("#si-loc").value,
      });
      if (!ok) { msg.textContent = "Item name is required (and must not already exist)."; return; }
      QBR._supUI.form = null; QBR.supRender(document.getElementById("sup-body"));
    });
  }
  host.querySelectorAll('[data-sup-form=""]').forEach(b => b.addEventListener("click", () => {
    QBR._supUI.form = null; QBR.supRender(document.getElementById("sup-body"));
  }));
}

/* ---------- Node export guard (test harness only; inert in browser) ------ */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { supParseItemRows, supParseTransRows, supNormType, supColIndex };
}
