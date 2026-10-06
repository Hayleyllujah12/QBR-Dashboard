/* ============================================================================
 * QBR Dashboard — persistence layer (js/persist.js)
 * v1.8 — 2026-10-04
 *
 * Two halves:
 *  A. ENTRY JOURNAL (refresh-proof). Every inventory/supplies mutation is
 *     recorded as a structured, replayable op {kind, op, args} in
 *     localStorage, keyed by the source workbook's fingerprint
 *     (name|size|lastModified). After every parse (fresh upload AND session
 *     restore) the matching ops are replayed onto the freshly built model,
 *     so Ctrl+R no longer wipes recorded entries. Replayed ops are NOT
 *     re-recorded (QBR._replaying guard) and do not duplicate the
 *     human-readable edit log (invLog guard in inventory.js).
 *  B. DIRECT FILE SAVE (Chrome/Edge, File System Access API). The user can
 *     LINK a workbook file; the browser grants per-file read/write
 *     permission via an explicit prompt. "Save" then writes the regenerated
 *     workbook bytes straight back into that same file — no download dance.
 *     File handles are stored in IndexedDB (localStorage cannot hold them).
 *     Non-Chromium browsers, denied permission, or any write error fall
 *     back to the classic Export download.
 *
 * Safety rules:
 *  - A journal op replays only when its exact file fingerprint is loaded.
 *    A changed/replaced file never receives another file's entries.
 *  - After a successful direct save the file's bytes are re-read: the
 *    session cache, the link metadata, and the kind map are all updated,
 *    and the now-redundant journal entries are cleared.
 *  - Before overwriting, the on-disk file is checked for external changes
 *    (size/mtime drift, e.g. edited in Excel) and the user must confirm.
 *
 * Depends on: QBR namespace, XLSX, loadItems/fileKey/cacheSession/renderAll
 * from app.js (all runtime-only, guarded). Loaded AFTER app.js.
 * ==========================================================================*/
var QBR = window.QBR = window.QBR || {};

/* ============================ A. ENTRY JOURNAL ========================== */

QBR.PERSIST_VERSION = "1.9.0";

// Fingerprint: identifies the exact file bytes an entry was recorded against.
QBR.fpOf = function (name, blob) {
  return String(name || "workbook.xlsx") + "|" + (blob && blob.size || 0) + "|" + (blob && blob.lastModified || 0);
};

// APP is a top-level const in app.js: visible as a bare global, NOT as window.APP.
function persistAppFiles() {
  try { return (typeof APP !== "undefined" && APP.files) || []; }
  catch (e) { return []; }
}
var JKEY = "qbr-inv-journal-v1";
var JMAX_OPS = 500, JMAX_AGE = 90 * 24 * 3600 * 1000;

function jLoad() {
  try { return JSON.parse(localStorage.getItem(JKEY)) || {}; }
  catch (e) { return {}; }
}
function jSave(store) {
  try { localStorage.setItem(JKEY, JSON.stringify(store)); } catch (e) { /* quota/full: entries stay in memory */ }
}
function jPrune(store) {
  const cut = Date.now() - JMAX_AGE;
  Object.keys(store).forEach(fp => {
    const e = store[fp];
    if (!e || !Array.isArray(e.ops) || !e.ops.length) { delete store[fp]; return; }
    e.ops = e.ops.filter(o => o && o.ts > cut);
    if (!e.ops.length) delete store[fp];
    else if (e.ops.length > JMAX_OPS) e.ops = e.ops.slice(-JMAX_OPS);
  });
  return store;
}

// Fingerprints currently loaded that contain a given inventory kind.
// QBR._kindByFp is rebuilt by the parsers on every load (see inventory.js /
// supplies.js) as { fp: Set("assets"|"supplies") }.
QBR._fpsForKind = function (kind) {
  const m = QBR._kindByFp || {};
  return Object.keys(m).filter(fp => { try { return m[fp] && m[fp].has(kind); } catch (e) { return false; } });
};

// Record one mutation. Called by the write-back mutations in inventory.js /
// supplies.js. Silent no-op while replaying or when no source file is known.
QBR.journalRecord = function (kind, op, args) {
  if (QBR._replaying) return;
  if (typeof localStorage === "undefined") return;
  const fps = QBR._fpsForKind(kind);
  if (!fps.length) return;
  const fp = fps[0];
  let fileName = "";
  try {
    const f = persistAppFiles().filter(x => QBR.fpOf(x.name, x.blob) === fp)[0];
    if (f) fileName = f.name;
  } catch (e) {}
  const store = jPrune(jLoad());
  const e = store[fp] || (store[fp] = { fileName: fileName, ops: [] });
  e.fileName = fileName || e.fileName;
  let id = "op" + Date.now().toString(36);
  try { id += Math.floor(Math.random() * 1e6).toString(36); } catch (x) {}
  e.ops.push({ id: id, kind: kind, op: op, args: args || [], ts: Date.now() });
  if (e.ops.length > JMAX_OPS) e.ops = e.ops.slice(-JMAX_OPS);
  jSave(store);
  QBR.persistRefreshBadge();
};

// Replay every stored op whose file is currently loaded, onto the freshly
// parsed model. Called from loadItems (app.js) after processBuffers, for both
// fresh uploads and session restores. The model is always rebuilt from raw
// bytes first, so each op applies exactly once — no dedup set needed.
QBR.journalReplayFor = function (files) {
  const store = jLoad();
  let n = 0;
  QBR._replaying = true;
  try {
    (files || []).forEach(f => {
      const fp = QBR.fpOf(f.name, f.blob);
      const e = store[fp];
      if (!e || !e.ops || !e.ops.length) return;
      e.ops.forEach(o => {
        const fn = QBR[o.op];
        if (typeof fn !== "function") return;
        try { fn.apply(null, o.args || []); n++; }
        catch (err) { console.warn("[QBR] journal replay failed:", o.op, err && err.message); }
      });
    });
  } finally {
    QBR._replaying = false;
  }
  QBR.persistRefreshBadge();
  return n;
};

// Drop the journal for one fingerprint (used after its data was saved to disk).
QBR.journalClearFp = function (fp) {
  if (!fp || typeof localStorage === "undefined") return;
  const store = jLoad();
  if (store[fp]) { delete store[fp]; jSave(store); }
  QBR.persistRefreshBadge();
};

// How many unsaved (dashboard-only) entries exist for the loaded files.
QBR.journalUnsavedCount = function (kind) {
  const store = jLoad();
  let fps = [];
  try { fps = persistAppFiles().map(x => QBR.fpOf(x.name, x.blob)); } catch (e) {}
  let n = 0;
  fps.forEach(fp => {
    const e = store[fp];
    if (!e || !e.ops) return;
    e.ops.forEach(o => { if (!kind || o.kind === kind) n++; });
  });
  return n;
};

// Fill every [data-unsaved] badge with the current unsaved-entry count.
QBR.persistRefreshBadge = function () {
  try {
    document.querySelectorAll("[data-unsaved]").forEach(el => {
      if (el.dataset.persistBusy) return;
      const n = QBR.journalUnsavedCount(el.dataset.unsaved);
      el.innerHTML = n
        ? ` <span class="badge bg-warning text-dark" title="Saved in the dashboard, not yet written to the Excel file">● ${n} unsaved</span>`
        : "";
    });
  } catch (e) {}
};

// Transient status text in the same badge slot ("Saving…", "Saved ✓").
QBR.persistNote = function (kind, html, ms) {
  try {
    const el = document.querySelector(`[data-unsaved="${kind}"]`);
    if (!el) return;
    el.dataset.persistBusy = "1";
    el.innerHTML = " " + html;
    setTimeout(() => { delete el.dataset.persistBusy; QBR.persistRefreshBadge(); }, ms || 3500);
  } catch (e) {}
};

/* ====================== B. DIRECT FILE SAVE (FS API) ==================== */

QBR.fsSupported = function () {
  return typeof window.showOpenFilePicker === "function";
};

// --- IndexedDB KV for file handles (handles cannot live in localStorage) ---
function persistIdb() {
  return new Promise((res, rej) => {
    if (!window.indexedDB) return rej(new Error("no-indexeddb"));
    let rq;
    try { rq = indexedDB.open("qbr-cache", 1); } catch (e) { return rej(e); }
    rq.onupgradeneeded = () => { try { rq.result.createObjectStore("session"); } catch (e) {} };
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
  });
}
function persistIdbPut(key, val) {
  return persistIdb().then(db => new Promise((res, rej) => {
    const tx = db.transaction("session", "readwrite");
    tx.objectStore("session").put(val, key);
    tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error);
  }));
}
function persistIdbGet(key) {
  return persistIdb().then(db => new Promise((res, rej) => {
    const tx = db.transaction("session", "readonly");
    const rq = tx.objectStore("session").get(key);
    rq.onsuccess = () => res(rq.result || null);
    rq.onerror = () => rej(rq.error);
  }));
}

var FS_LINKS_KEY = "fs-links";
QBR._fsLinks = []; // [{fp, kinds:[], name, size, lastModified, handle}]

function fsLinksLoad() {
  return persistIdbGet(FS_LINKS_KEY).then(links => {
    QBR._fsLinks = Array.isArray(links) ? links : [];
    QBR.fsRefreshStatus();
    if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
  }).catch(() => { QBR._fsLinks = []; });
}
function fsLinksSave() {
  // Handles are structured-cloneable; plain metadata rides along.
  const slim = QBR._fsLinks.map(l => ({
    fp: l.fp, kinds: l.kinds, name: l.name,
    size: l.size, lastModified: l.lastModified, handle: l.handle,
  }));
  return persistIdbPut(FS_LINKS_KEY, slim).catch(e => console.warn("[QBR] link store failed:", e && e.message));
}
QBR.fsGetLink = function (kind) {
  return (QBR._fsLinks || []).filter(l => l.kinds && l.kinds.indexOf(kind) >= 0)[0] || null;
};

async function fsEnsurePermission(handle) {
  try {
    let p = await handle.queryPermission({ mode: "readwrite" });
    if (p !== "granted") p = await handle.requestPermission({ mode: "readwrite" });
    return p === "granted";
  } catch (e) { return false; }
}

// Link a workbook file: pick it, get read/write permission, load it through
// the normal accumulation path, and remember the handle per detected kind.
QBR.fsLinkFile = async function () {
  if (!QBR.fsSupported()) {
    alert("Direct file linking needs Chrome or Edge. Other browsers use Export download.");
    return;
  }
  let handle;
  try {
    const picks = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"], "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] } }],
    });
    handle = picks[0];
  } catch (e) { return; } // user cancelled
  if (!(await fsEnsurePermission(handle))) {
    alert("Permission was not granted — the file was not linked.");
    return;
  }
  let file;
  try { file = await handle.getFile(); }
  catch (e) { alert("Could not read the file: " + (e && e.message || e)); return; }
  QBR.persistNote("", "Linking…", 10000);
  try {
    // Accumulate exactly like a manual upload (dedup by name, newest wins).
    const map = new Map();
    persistAppFiles().forEach(it => map.set(fileKey(it.name), it));
    map.set(fileKey(file.name || "workbook.xlsx"), { name: file.name || "workbook.xlsx", blob: file });
    const ok = await loadItems([...map.values()], {});
    if (ok) {
      const fp = QBR.fpOf(file.name, file);
      const kinds = [...((QBR._kindByFp || {})[fp] || [])];
      // Replace any older link for the same file name or the same kinds.
      QBR._fsLinks = (QBR._fsLinks || []).filter(l =>
        fileKey(l.name) !== fileKey(file.name) && !(l.kinds || []).some(k => kinds.indexOf(k) >= 0));
      if (kinds.length) {
        QBR._fsLinks.push({ fp: fp, kinds: kinds, name: file.name, size: file.size, lastModified: file.lastModified, handle: handle });
        await fsLinksSave();
      }
      if (typeof cacheSession === "function") cacheSession();
      QBR.fsRefreshStatus();
      if (typeof renderAll === "function") renderAll();
      QBR.persistNote("", kinds.length ? `Linked ✓ ${file.name}` : "Loaded, but no inventory/supplies sheets found", 4000);
    }
  } catch (e) {
    alert("Could not link the file: " + (e && e.message || e));
  }
  QBR.persistRefreshBadge();
};

QBR.fsUnlink = async function (name) {
  if (!confirm(`Stop direct saving to "${name}"? (The dashboard keeps working; Export download still available.)`)) return;
  QBR._fsLinks = (QBR._fsLinks || []).filter(l => l.name !== name);
  await fsLinksSave();
  QBR.fsRefreshStatus();
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
};

// Header status line: which files are linked.
QBR.fsRefreshStatus = function () {
  try {
    const el = document.getElementById("fs-link-status");
    if (!el) return;
    const links = QBR._fsLinks || [];
    el.innerHTML = links.map(l =>
      `<span class="badge bg-success" title="Direct save enabled — click to unlink">🔗 ${String(l.name).replace(/</g, "&lt;")}</span>`
    ).join(" ");
    el.querySelectorAll(".badge").forEach((b, i) => {
      b.style.cursor = "pointer";
      b.onclick = () => QBR.fsUnlink(links[i].name);
    });
    const btn = document.getElementById("btn-linkfile");
    if (btn && !QBR.fsSupported()) btn.classList.add("d-none");
  } catch (e) {}
};

/* ---- Safe write helpers (2026-10-05) ------------------------------------
 * Root cause of "Excel cannot open the file … .xlsm": the linked save always
 * wrote bookType "xlsx" bytes into the user's .xlsm file. Excel validates the
 * workbook content type against the extension and refuses the mismatch; the
 * VBA project was dropped too (never read with bookVBA). Now:
 *  - the output type follows the linked file's extension (.xlsm → "xlsm"),
 *  - the VBA project is carried over (re-read once with bookVBA if needed),
 *  - every write is verified BEFORE touching the file (content type matches
 *    the extension, macros present when the source had them, all sheets
 *    there); any doubt → nothing is written and a separate copy downloads. */
function fsBookType(name) { return /\.xlsm$/i.test(String(name || "")) ? "xlsm" : "xlsx"; }
const FS_MAIN_CT = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
};
async function fsEnsureVba(wb, file) {
  if (!wb || wb.vbaraw || !file) return;
  try {
    const w2 = XLSX.read(await file.arrayBuffer(), { type: "array", bookVBA: true });
    if (w2 && w2.vbaraw) wb.vbaraw = w2.vbaraw;
  } catch (e) { /* verification below decides */ }
}
// Returns null when the bytes are safe to write, else a plain-language reason.
function fsVerifyBytes(bytes, bookType, wb) {
  try {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const zip = XLSX.CFB.read(u8, { type: "array" });
    const ctEntry = XLSX.CFB.find(zip, "/[Content_Types].xml");
    const ct = ctEntry ? new TextDecoder().decode(ctEntry.content) : "";
    if (ct.indexOf(FS_MAIN_CT[bookType]) < 0) return "the file format would not match its ." + bookType + " name";
    if (bookType === "xlsm" && wb && wb.vbaraw && !XLSX.CFB.find(zip, "/xl/vbaProject.bin")) return "the workbook's macros would be lost";
    const back = XLSX.read(u8, { type: "array", bookSheets: true });
    const want = (wb && wb.SheetNames || []).join("\u0001");
    if (want && (back.SheetNames || []).join("\u0001") !== want) return "some sheets would be missing";
    return null;
  } catch (e) { return "the new file could not be read back (" + ((e && e.message) || e) + ")"; }
}
QBR._fsBookType = fsBookType; QBR._fsVerifyBytes = fsVerifyBytes;   // exposed for tests

/* ---------- format-safe save engine (develop prototype, 2026-10-06) ---------- */
QBR.saveEngine = function () {
  let v = null; try { v = localStorage.getItem("qbr-save-engine"); } catch (e) {}
  return (v === "legacy" || !QBR.xlsxSurgical) ? "legacy" : "format-safe";
};
async function fsSaveSurgical(kind, link, cur) {
  const S = QBR.xlsxSurgical;
  const buf = new Uint8Array(await cur.arrayBuffer());
  // Two parses of the SAME bytes with the SAME options: one stays the baseline,
  // the other receives the journal; their difference is exactly what to write.
  const base = XLSX.read(buf, { type: "array", cellStyles: true });
  const work = XLSX.read(buf, { type: "array", cellStyles: true });
  if (typeof QBR.patchWorkbookFromJournal !== "function") return QBR.fsDownloadKind(kind, "download-fallback");
  const res = QBR.patchWorkbookFromJournal(kind, work, link.fp);
  if (!res.ok) { console.warn("[QBR] patch failed:", res.error); return QBR.fsDownloadKind(kind, "download-fallback"); }
  if (res.applied === 0) return { mode: "no-changes", name: link.name };
  const out = await S.surgicalSave(buf, base, work);
  const bookType = fsBookType(link.name);
  let why = !out.ok ? "format-safe save can't handle this file yet (" + out.reason + ")" : null;
  if (!why && out.noChanges) return { mode: "no-changes", name: link.name };
  if (!why) why = await S.verifySurgical(buf, out.bytes, work);
  if (!why) why = fsVerifyBytes(out.bytes, bookType, work);
  if (why) {
    console.warn("[QBR] format-safe save blocked — " + why, out.stats);
    const r = QBR.fsDownloadKind(kind, "download-unsafe");
    return Object.assign(r, { name: link.name, reason: why });
  }
  const w = await link.handle.createWritable();
  await w.write(out.bytes);
  await w.close();
  (QBR._origWb || (QBR._origWb = {}))[link.fp] = work; // retained copy now matches the file
  const fresh = await link.handle.getFile();
  await QBR.fsAfterSave(kind, link, fresh);
  console.info("[QBR] format-safe save:", out.stats);
  return { mode: "file", name: link.name, applied: res.applied, skipped: res.skipped, notes: res.notes, engine: "format-safe", stats: out.stats };
}

/* Rebuild-download path (classic Export). Used when no file is linked, when the
 * browser lacks the File System Access API, and as a fallback. NOTE: this
 * regenerates the workbook from parsed data — formulas/layout of the original
 * are NOT preserved. The linked direct-save path below (patch-in-place)
 * preserves them. */
QBR.fsDownloadKind = function (kind, mode) {
  const build = kind === "supplies" ? QBR.supBuildWorkbook : QBR.invBuildWorkbook;
  if (typeof build !== "function") return { mode: "no-data" };
  const b = build();
  if (!b || !b.wb) return { mode: "no-data" };
  XLSX.writeFile(b.wb, b.filename, { cellStyles: true });
  return { mode: mode || "download", filename: b.filename };
};

/* Last-resort linked write: the on-disk file changed externally and the user
 * confirmed overwrite. Recorded cell coordinates may be stale, so fall back
 * to the full rebuild write, then re-sync the model from the fresh bytes. */
QBR.fsRebuildToLink = async function (kind, link) {
  const build = kind === "supplies" ? QBR.supBuildWorkbook : QBR.invBuildWorkbook;
  if (typeof build !== "function") return { mode: "no-data" };
  const b = build();
  if (!b || !b.wb) return { mode: "no-data" };
  const oldFp = link.fp;
  if (typeof QBR.fsDropPatchState === "function") QBR.fsDropPatchState(oldFp);
  const bookType = fsBookType(link.name);
  if (bookType === "xlsm") return QBR.fsDownloadKind(kind, "download-unsafe");   // a rebuild can't carry macros
  const bytes = XLSX.write(b.wb, { bookType: bookType, type: "array", cellStyles: true });
  if (fsVerifyBytes(bytes, bookType, null)) return QBR.fsDownloadKind(kind, "download-unsafe");
  const w = await link.handle.createWritable();
  await w.write(bytes);
  await w.close();
  const fresh = await link.handle.getFile();
  await QBR.fsAfterSave(kind, link, fresh);
  try {
    if (typeof loadItems === "function" && typeof APP !== "undefined" && APP.files) {
      await loadItems(APP.files, {});
      if (typeof cacheSession === "function") cacheSession();
    }
  } catch (e) { console.warn("[QBR] post-rebuild re-sync failed:", e && e.message); }
  return { mode: "file", name: link.name, rebuilt: true };
};

/* Save one inventory kind. Returns a result descriptor:
 *  file             — bytes written straight into the linked file (patch-in-place)
 *  no-changes       — linked file, but the journal held nothing to write
 *  download         — no link: fell back to Export download
 *  download-fallback— write failed: downloaded instead
 *  cancelled        — user declined the overwrite confirm
 *  denied           — permission not granted (then downloaded? no: aborted)
 *  no-data          — nothing to save */
QBR.fsSaveKind = async function (kind) {
  const link = QBR.fsGetLink(kind);
  if (!link || !QBR.fsSupported()) return QBR.fsDownloadKind(kind);
  try {
    if (!(await fsEnsurePermission(link.handle))) return { mode: "denied", name: link.name };
    const cur = await link.handle.getFile();
    const changed = cur.size !== link.size || cur.lastModified !== link.lastModified;
    if (changed) {
      // 2026-10-05: never rebuild over the linked file (that wiped formulas and
      // layout, and broke .xlsm files). Download a separate copy instead.
      alert(`"${link.name}" changed outside the dashboard since it was linked (for example, edited in Excel).\n\n` +
        `To keep that file safe, nothing will be written to it. Your changes will download as a separate copy.\n` +
        `To save directly again, click "Link file" and pick the updated workbook.`);
      const r = QBR.fsDownloadKind(kind, "download-changed");
      return Object.assign(r, { name: link.name });
    }
    // DEVELOP 2026-10-06: format-safe save (js/xlsx-surgical.js). Writes only the
    // changed cells into the file and copies every other part untouched, so
    // styles, conditional formatting, validation, tables, charts and macros
    // survive. Opt out per browser: localStorage "qbr-save-engine" = "legacy".
    if (QBR.saveEngine() === "format-safe") return await fsSaveSurgical(kind, link, cur);
    // Patch-in-place: replay the journal as cell writes into the ORIGINAL
    // workbook, preserving its formulas, layout and helper columns.
    let wb = (QBR._origWb || {})[link.fp];
    if (!wb) {
      const buf = await cur.arrayBuffer();
      wb = XLSX.read(buf, { type: "array" });
      (QBR._origWb || (QBR._origWb = {}))[link.fp] = wb;
    }
    if (typeof QBR.patchWorkbookFromJournal !== "function") return QBR.fsDownloadKind(kind, "download-fallback");
    const res = QBR.patchWorkbookFromJournal(kind, wb, link.fp);
    if (!res.ok) {
      console.warn("[QBR] patch failed:", res.error, "— falling back to download");
      return QBR.fsDownloadKind(kind, "download-fallback");
    }
    if (res.applied === 0) return { mode: "no-changes", name: link.name };
    // Ask Excel to recalculate on open (patched cells have no calc chain).
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.CalcPr = { fullCalcOnLoad: true };
    try { delete wb.CalcChain; } catch (e) {}
    const bookType = fsBookType(link.name);
    if (bookType === "xlsm") await fsEnsureVba(wb, cur);
    const bytes = XLSX.write(wb, { bookType: bookType, type: "array", cellStyles: true, bookVBA: true });
    const why = fsVerifyBytes(bytes, bookType, wb);
    if (why) {
      console.warn("[QBR] save blocked — " + why);
      const r = QBR.fsDownloadKind(kind, "download-unsafe");
      return Object.assign(r, { name: link.name, reason: why });
    }
    const w = await link.handle.createWritable();
    await w.write(bytes);
    await w.close();
    const fresh = await link.handle.getFile();
    await QBR.fsAfterSave(kind, link, fresh);
    return { mode: "file", name: link.name, applied: res.applied, skipped: res.skipped, notes: res.notes };
  } catch (e) {
    console.warn("[QBR] direct save failed, falling back to download:", e && e.message);
    return QBR.fsDownloadKind(kind, "download-fallback");
  }
};

// After a successful direct write: refresh session bytes, link metadata,
// kind map, and clear the journal entries now baked into the file.
QBR.fsAfterSave = async function (kind, link, freshFile) {
  const oldFp = link.fp;
  const newFp = QBR.fpOf(freshFile.name, freshFile);
  try {
    // 1. session file list now carries the fresh bytes
    try {
      if (typeof APP !== "undefined" && APP.files) {
        APP.files = APP.files.map(it =>
          fileKey(it.name) === fileKey(link.name) ? { name: link.name, blob: freshFile } : it);
      }
    } catch (e) {}
    // 2. link metadata follows the new file state
    link.fp = newFp; link.size = freshFile.size; link.lastModified = freshFile.lastModified;
    await fsLinksSave();
    // 3. kind map: move the entry to the new fingerprint so later mutations
    //    record against the right file
    if (QBR._kindByFp && QBR._kindByFp[oldFp]) {
      QBR._kindByFp[newFp] = QBR._kindByFp[oldFp];
      delete QBR._kindByFp[oldFp];
    }
    // 3b. patch state: the retained workbook was patched in place, so its
    //     coordinates are still valid — carry them to the new fingerprint.
    //     (The rebuild path drops them beforehand via fsDropPatchState.)
    if (typeof QBR.fsCarryPatchState === "function") QBR.fsCarryPatchState(oldFp, newFp);
    // 4. journal entries are now in the file — clear them
    QBR.journalClearFp(oldFp);
    // 5. persist the fresh bytes for session restore
    if (typeof cacheSession === "function") cacheSession();
  } catch (e) { console.warn("[QBR] post-save bookkeeping failed:", e && e.message); }
  QBR.fsRefreshStatus();
};

// Convenience wrappers wired to the Save buttons.
QBR.supSave = async function () {
  QBR.persistNote("supplies", "Saving…", 15000);
  const r = await QBR.fsSaveKind("supplies");
  QBR.supSaveDone(r);
};
QBR.invSave = async function () {
  QBR.persistNote("assets", "Saving…", 15000);
  const r = await QBR.fsSaveKind("assets");
  QBR.invSaveDone(r);
};
QBR.supSaveDone = function (r) {
  if (!r) return;
  if (r.mode === "file") {
    const extra = (r.applied ? ` (${r.applied} change${r.applied === 1 ? "" : "s"}${r.skipped ? `, ${r.skipped} skipped` : ""})` : "");
    QBR.persistNote("supplies", `Saved ✓ ${r.name}${extra}${r.engine === "format-safe" ? " · formatting kept" : ""}`, 5000);
  }
  else if (r.mode === "no-changes") QBR.persistNote("supplies", "No changes to save", 3000);
  else if (r.mode === "download" || r.mode === "download-fallback")
    QBR.persistNote("supplies", `Downloaded ${r.filename || ""} — link a file for direct save`, 5000);
  else if (r.mode === "download-unsafe")
    QBR.persistNote("supplies", `Not saved to ${r.name || "the linked file"} — ${r.reason || "safety check failed"}. Downloaded a separate copy instead.`, 12000);
  else if (r.mode === "download-changed")
    QBR.persistNote("supplies", `${r.name || "Linked file"} changed in Excel — downloaded a separate copy. Re-link to save directly.`, 12000);
  else if (r.mode === "denied") QBR.persistNote("supplies", "Permission denied — file not saved", 5000);
  else if (r.mode === "no-data") alert("Load a supplies workbook first.");
  else QBR.persistRefreshBadge();
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
};
QBR.invSaveDone = function (r) {
  if (!r) return;
  if (r.mode === "file") {
    const extra = (r.applied ? ` (${r.applied} change${r.applied === 1 ? "" : "s"}${r.skipped ? `, ${r.skipped} skipped` : ""})` : "");
    QBR.persistNote("assets", `Saved ✓ ${r.name}${extra}${r.engine === "format-safe" ? " · formatting kept" : ""}`, 5000);
  }
  else if (r.mode === "no-changes") QBR.persistNote("assets", "No changes to save", 3000);
  else if (r.mode === "download" || r.mode === "download-fallback")
    QBR.persistNote("assets", `Downloaded ${r.filename || ""} — link a file for direct save`, 5000);
  else if (r.mode === "download-unsafe")
    QBR.persistNote("assets", `Not saved to ${r.name || "the linked file"} — ${r.reason || "safety check failed"}. Downloaded a separate copy instead.`, 12000);
  else if (r.mode === "download-changed")
    QBR.persistNote("assets", `${r.name || "Linked file"} changed in Excel — downloaded a separate copy. Re-link to save directly.`, 12000);
  else if (r.mode === "denied") QBR.persistNote("assets", "Permission denied — file not saved", 5000);
  else if (r.mode === "no-data") alert("Load an inventory workbook first.");
  else QBR.persistRefreshBadge();
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
};

// Adaptive Save button for an inventory kind: direct "Save to Excel" when a
// file is linked, otherwise the classic "Export workbook" download.
QBR.saveButtonHtml = function (kind, id) {
  const link = (typeof QBR.fsGetLink === "function") ? QBR.fsGetLink(kind) : null;
  const esc = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const badge = `<span data-unsaved="${kind}" class="small text-muted align-self-center"></span>`;
  if (link) {
    return `<button type="button" class="btn btn-sm btn-primary" id="${id}" title="Write directly to ${esc(link.name)}">💾 Save to Excel</button>` + badge;
  }
  return `<button type="button" class="btn btn-sm btn-outline-secondary" id="${id}" title="Download the workbook (use Link file for direct save)">⭳ Export workbook</button>` + badge;
};

// Boot: load remembered links, paint status, hide the Link button where
// the File System Access API does not exist.
QBR.fsInit = function () {
  try {
    const btn = document.getElementById("btn-linkfile");
    if (btn) {
      if (!QBR.fsSupported()) btn.classList.add("d-none");
      else btn.addEventListener("click", () => QBR.fsLinkFile());
    }
  } catch (e) {}
  fsLinksLoad();
  QBR.persistRefreshBadge();
};

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", QBR.fsInit);
else QBR.fsInit();
