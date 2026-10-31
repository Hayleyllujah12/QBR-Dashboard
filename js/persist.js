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
 *     v1.30.0: the journal store moved to js/journal.js (localStorage + IndexedDB
 *     backup, per-folder key, no caps, orphan/parked/export panel).
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

QBR.PERSIST_VERSION = "1.10.0"; // v1.30.0: journal moved to js/journal.js (journal safety)

// Fingerprint: identifies the exact file bytes an entry was recorded against.
QBR.fpOf = function (name, blob) {
  return String(name || "workbook.xlsx") + "|" + (blob && blob.size || 0) + "|" + (blob && blob.lastModified || 0);
};
// cyrb53 -- fast non-crypto 53-bit hash for change detection. Sync, no secure
// context needed (the app also runs over file:// where crypto.subtle may be
// unavailable). Hex string output. v1.25.1: kills OneDrive timestamp-only
// false positives in the linked-save guard.
function fsHashBytes(u8) {
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < u8.length; i++) {
    const ch = u8[i];
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16);
}

// File lineage (live-edit safeguard): the browser remembers, per file name,
// the fingerprint + content hash of the last version IT saved (qbr-lineage-v1).
// Linking a different copy under the same name while journal edits are still
// pending against the saved version earns an informational warning — the
// journal replays by exact fingerprint, so edits can never land on the wrong
// copy; this just stops the user wondering where their edits went. Records are
// written ONLY on successful saves/rebases, never at link time, and every
// access is guarded: a lineage failure must never break linking or saving.
var LINEAGE_KEY = "qbr-lineage-v1";
function lineageMap() { try { return JSON.parse(localStorage.getItem(LINEAGE_KEY) || "{}") || {}; } catch (e) { return {}; } }
function lineageRecord(name, fp, hash) {
  if (!name) return;
  try {
    var m = lineageMap();
    m[name] = { fp: fp || "", hash: hash || "", savedAt: new Date().toISOString() };
    localStorage.setItem(LINEAGE_KEY, JSON.stringify(m));
  } catch (e) {}
}
async function lineageRecordFile(name, file) {
  try {
    var h = "";
    try { h = fsHashBytes(new Uint8Array(await file.arrayBuffer())); } catch (e) {}
    lineageRecord(name, QBR.fpOf(name, file), h);
  } catch (e) {}
}

// APP is a top-level const in app.js: visible as a bare global, NOT as window.APP.
function persistAppFiles() {
  try { return (typeof APP !== "undefined" && APP.files) || []; }
  catch (e) { return []; }
}
// The journal itself (record / replay / clear / badge / panel) lives in
// js/journal.js since v1.30.0 (journal safety). Loaded right after this file.

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
  // Side links (QBR.fsAttachSide) are re-attached by their own tab each session — not stored here.
  const slim = QBR._fsLinks.filter(l => !l.side).map(l => ({
    fp: l.fp, kinds: l.kinds, name: l.name,
    size: l.size, lastModified: l.lastModified, hash: l.hash, handle: l.handle,
    dirId: l.dirId || null, rel: l.rel || null,   // v1.33.0: opened through a linked folder (js/fsfolder.js)
  }));
  return persistIdbPut(FS_LINKS_KEY, slim).catch(e => console.warn("[QBR] link store failed:", e && e.message));
}
QBR.fsGetLink = function (kind) {
  return (QBR._fsLinks || []).filter(l => l.kinds && l.kinds.indexOf(kind) >= 0)[0] || null;
};

async function fsEnsurePermission(handle) {
  try {
    let p = await handle.queryPermission({ mode: "readwrite" });
    if (p === "granted") return true;
    // v1.33.0: a file opened through "Open folder" → re-grant the FOLDER once;
    // that covers every linked workbook in it (no prompt per file).
    if (typeof QBR.fsFolderGrantFor === "function") {
      const h2 = await QBR.fsFolderGrantFor(handle);
      if (h2 === false) return false;
      if (h2) {
        try { if ((await h2.queryPermission({ mode: "readwrite" })) === "granted") return true; } catch (e) {}
      }
    }
    p = await handle.requestPermission({ mode: "readwrite" });
    return p === "granted";
  } catch (e) { return false; }
}

// Link a workbook file: pick it, get read/write permission, load it through
// the normal accumulation path, and remember the handle per detected kind.
QBR.fsLinkFile = async function (expectName, opts) {
  opts = opts || {};
  if (!QBR.fsSupported()) {
    alert("Direct file linking needs Chrome or Edge. Other browsers use Export download.");
    return;
  }
  let handle;
  try {
    // 2026-10-30: open the picker in the folder last used with "Open folder..." (fsfolder.js)
    let startIn = null;
    try { if (typeof QBR.fsFolderStartIn === "function") startIn = QBR.fsFolderStartIn(); } catch (e) {}
    const picks = await window.showOpenFilePicker({
      multiple: false,
      ...(startIn ? { startIn: startIn } : {}),
      types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"], "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] } }],
    });
    handle = picks[0];
  } catch (e) {
    if (!e || e.name === "AbortError") return false; // user cancelled
    // 2026-10-30: the click that started this can expire while the folder is searched (or after a
    // permission prompt) and the browser then refuses the picker - ask for one more click.
    if (opts.retried || typeof QBR.fsAskClick !== "function") return false;
    if (!(await QBR.fsAskClick(`Pick "${expectName || "the workbook"}" to turn on direct save.`))) return false;
    return QBR.fsLinkFile(expectName, Object.assign({}, opts, { retried: true }));
  }
    if (!(await fsEnsurePermission(handle))) {
    alert("Permission was not granted \u2014 the file was not linked.");
    return false;
  }
  let file;
  try { file = await handle.getFile(); }
  catch (e) { alert("Could not read the file: " + (e && e.message || e)); return false; }
  if (expectName && fsNameKey(file.name) !== fsNameKey(expectName)) {
    alert(`You picked "${file.name}" \u2014 expected "${expectName}". No link was created.`);
    return false;
  }
  let fileHash = "";
  try { fileHash = fsHashBytes(new Uint8Array(await file.arrayBuffer())); } catch (e) {} // v1.25.1: content hash for false-positive guard
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
        // 2026-10-30: a file picked inside a remembered folder becomes a folder link (re-allowed with the folder next visit)
        let inDir = null;
        try { if (typeof QBR.fsFolderOf === "function") inDir = await QBR.fsFolderOf(handle); } catch (e) {}
        QBR._fsLinks.push({ fp: fp, kinds: kinds, name: file.name, size: file.size, lastModified: file.lastModified, hash: fileHash, handle: handle,
          dirId: inDir ? inDir.id : null, rel: inDir ? inDir.rel : null });
        await fsLinksSave();
      }
      if (typeof cacheSession === "function") cacheSession();
      QBR.fsRefreshStatus();
      if (typeof renderAll === "function") renderAll();
      QBR.persistNote("", kinds.length ? `Linked ✓ ${file.name}` : "Loaded, but no inventory/supplies sheets found", 4000);
    }
  } catch (e) {
    alert("Could not link the file: " + (e && e.message || e));
    return false;
  }
  QBR.persistRefreshBadge();
  if (typeof renderFileList === "function") { try { renderFileList(); } catch (e2) {} }
      // Lineage warning: this name was saved from this browser before, the
      // linked copy's content differs, and edits recorded against the saved
      // version are still pending — they will not apply to this copy.
      try {
        const rec = lineageMap()[file.name];
        if (rec && rec.hash && fileHash && fileHash !== rec.hash) {
          let pending = 0;
          try { pending = QBR.journalOpsFor(rec.fp).length; } catch (e) {}
          if (pending > 0) {
            QBR.persistNote((((QBR._fsLinks || []).find(l => l.name === file.name) || {}).kinds || [])[0] || "", `⚠ ${file.name}: this copy differs from the version last saved in this browser (${new Date(rec.savedAt).toLocaleString()}). Edits recorded against that version won't apply to this copy.`, 15000);
          }
        }
      } catch (e) {}
  return true;
};

// Per-file edit toggle (checkbox in the loaded-files list).
function fsNameKey(n) {
  try { return String(n || "").trim().toLowerCase(); }
  catch (e) { return ""; }
}
QBR.fsIsLinkedByName = function (name) {
  try { return (QBR._fsLinks || []).some(l => fsNameKey(l.name) === fsNameKey(name)); }
  catch (e) { return false; }
};
QBR.fsToggleFileLink = async function (name, enable) {
  if (enable) {
    // 2026-10-30: first look for the file in a folder opened with "Open folder..." (no picker when found)
    if (typeof QBR.fsAutoLinkByName === "function") {
      let r = null;
      try { r = await QBR.fsAutoLinkByName(name); } catch (e) { console.warn("[QBR] folder auto-link:", e && e.message); r = null; }
      if (r !== null) return r;
    }
    return QBR.fsLinkFile(name);
  }
  await QBR.fsUnlink(name, true, true); // keep the folder remembered so re-ticking finds the file again
  if (typeof renderFileList === "function") { try { renderFileList(); } catch (e) {} }
  return true;
};

QBR.fsUnlink = async function (name, skipConfirm, keepFolder) {
  if (!skipConfirm && !confirm(`Stop direct saving to "${name}"? (The dashboard keeps working; Export download still available.)`)) return;
  QBR._fsLinks = (QBR._fsLinks || []).filter(l => l.name !== name);
  await fsLinksSave();
  try { if (!keepFolder && QBR.fsFolderPrune) QBR.fsFolderPrune(); } catch (e) {}
  QBR.fsRefreshStatus();
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
};

// Header status line: which files are linked.
QBR.fsRefreshStatus = function () {
  try {
    const el = document.getElementById("fs-link-status");
    if (!el) return;
    const links = QBR._fsLinks || [];
    // v1.33.0: an opened folder links every workbook — one summary badge instead of a header full of them
    if (links.length > 2) {
      const names = links.map(l => "• " + l.name).join("\n");
      el.innerHTML = `<span class="badge bg-success" title="Direct save / reload linked:\n${names.replace(/"/g, "&quot;").replace(/</g, "&lt;")}\n\nOpen the Loaded list to unlink one">🔗 ${links.length} linked</span>`;
      return;
    }
    el.innerHTML = links.map(l =>
      `<span class="badge bg-success" title="Direct save enabled — click to unlink">🔗 ${String(l.name).replace(/</g, "&lt;")}</span>`
    ).join(" ");
    el.querySelectorAll(".badge").forEach((b, i) => {
      b.style.cursor = "pointer";
      b.onclick = () => QBR.fsUnlink(links[i].name);
    });
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

/* ---------- three-way merge save (2026-10-06; js/xlsx-merge.js) ----------
 * Base = the file as loaded / last saved (in memory). When the linked file was
 * changed in Excel, the dashboard's edits are rebased onto the current file:
 * columns by header, rows by key; same-cell conflicts go to a review dialog. */
QBR._fsBase = QBR._fsBase || {};
function fsSetBase(name, u8) { try { QBR._fsBase[fileKey(name)] = u8 instanceof Uint8Array ? u8 : new Uint8Array(u8); } catch (e) {} }
function fsBaseOf(name) { try { return QBR._fsBase[fileKey(name)] || null; } catch (e) { return null; } }
QBR.fsRememberBase = function (items, buffers) {   // called by loadItems with the bytes it parsed
  try { (items || []).forEach((it, i) => { if (buffers[i]) fsSetBase(it.name, buffers[i]); }); } catch (e) {}
};
async function fsMemFile(file, name) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  return { bytes, file: new File([bytes], name || file.name, { lastModified: file.lastModified, type: file.type || "" }) };
}
// Write bytes into the linked file. Returns null on success, or a result object
// when the file is locked (open in desktop Excel) — the journal is kept.
async function fsWriteBytes(kind, link, bytes) {
  let w = null;
  try {
    w = await link.handle.createWritable();
    await w.write(bytes);
    await w.close();
    return null;
  } catch (e) {
    try { if (w && w.abort) await w.abort(); } catch (x) {}
    const name = (e && e.name) || "", msg = (e && e.message) || String(e);
    console.warn("[QBR] write failed:", name, msg);
    const locked = /NoModificationAllowed|InvalidState|InvalidModification|NotReadable/i.test(name) || /lock|in use|being used|access/i.test(msg);
    const r = { mode: locked ? "locked" : "write-failed", name: link.name, reason: name || msg, kind: kind };
    try {
      QBR.persistNote(kind, locked
        ? `Not saved — "${link.name}" is open in Excel. Close it there, then click Save again. Your edits are kept.`
        : `Not saved — could not write "${link.name}" (${name || msg}). Your edits are kept.`, 15000);
    } catch (x) {}
    return r;
  }
}
// Re-parse the dashboard from the files as they are now (shows Excel's edits,
// refreshes every row/column position after a merge).
async function fsReloadModel() {
  try { if (QBR._auditWatch) { QBR._auditWatch.changedFp = null; QBR._auditWatch.freshBuf = null; } } catch (e) {}
  if (typeof loadItems === "function" && typeof APP !== "undefined" && APP.files) {
    await loadItems(APP.files.slice(), {});
    if (typeof cacheSession === "function") cacheSession();
  }
}
async function fsSaveMerged(kind, link, cur, attempt) {
  const S = QBR.xlsxSurgical, M = QBR.xlsxMerge;
  const curBuf = new Uint8Array(await cur.arrayBuffer());
  const baseBytes = fsBaseOf(link.name);
  const opt = { type: "array", cellStyles: true };
  const base = XLSX.read(baseBytes, opt), mine = XLSX.read(baseBytes, opt);
  const res = QBR.patchWorkbookFromJournal(kind, mine, link.fp);
  if (!res.ok) return QBR.fsDownloadKind(kind, "download-fallback");
  if (res.applied === 0) {                       // nothing pending: just pick up Excel's version
    await QBR.fsReloadFromFile(link.name, { quiet: true });
    return { mode: "reloaded", name: link.name, kind: kind };
  }
  const theirs = XLSX.read(curBuf, opt), final = XLSX.read(curBuf, opt);
  const m = M.rebase(base, mine, theirs, final);
  if (m.conflicts.length || m.unresolved.length) {
    const choices = await QBR.fsMergeReview(m, link.name);
    if (!choices) {
      QBR.persistNote(kind, "Save cancelled — nothing was written. Your edits are kept.", 8000);
      return { mode: "merge-cancelled", name: link.name, kind: kind, conflicts: m.conflicts.length, unresolved: m.unresolved.length };
    }
    M.resolve(final, m.conflicts, choices);
  }
  const out = await S.surgicalSave(curBuf, theirs, final);
  let why = !out.ok ? "format-safe save can't handle this file yet (" + out.reason + ")" : null;
  if (!why && !out.noChanges) why = (await S.verifySurgical(curBuf, out.bytes, final)) || fsVerifyBytes(out.bytes, fsBookType(link.name), final);
  if (why) {
    console.warn("[QBR] merged save blocked — " + why);
    const r = QBR.fsDownloadKind(kind, "download-unsafe");
    return Object.assign(r, { name: link.name, reason: why });
  }
  // Last-moment check (OneDrive sync / Excel Online can change the file mid-save).
  const again = await link.handle.getFile();
  const againBuf = new Uint8Array(await again.arrayBuffer());
  if (fsHashBytes(againBuf) !== fsHashBytes(curBuf)) {
    if (attempt < 3) return fsSaveMerged(kind, link, again, attempt + 1);
    QBR.persistNote(kind, `Not saved — "${link.name}" keeps changing (sync in progress?). Try again in a moment. Your edits are kept.`, 12000);
    return { mode: "busy", name: link.name, kind: kind };
  }
  if (!out.noChanges) {
    const wr = await fsWriteBytes(kind, link, out.bytes);
    if (wr) return wr;
  }
  // v1.30.0: edits that couldn't be placed are kept (Unsaved edits → "Couldn't be placed"), not dropped.
  if (m.unresolved.length) { try { QBR.journalPark(link.name, m.unresolved); } catch (e) {} }
  const fresh = await link.handle.getFile();
  await QBR.fsAfterSave(kind, link, fresh);       // new base, journal cleared
  if (!link.side) await fsReloadModel();          // positions + Excel's edits now in the dashboard (side: its tab re-parsed in fsAfterSave)
  const merged = { applied: m.applied, conflicts: m.conflicts.length, unresolved: m.unresolved.length, remapped: m.remapped };
  console.info("[QBR] merged save:", merged, out.stats);
  return { mode: "file", name: link.name, applied: res.applied, skipped: res.skipped, notes: res.notes, engine: "format-safe", merged: merged, kind: kind };
}
QBR._fsSaveMerged = fsSaveMerged;
QBR.fsHashOf = fsHashBytes;

/* Reload a linked workbook from disk through its existing link (no picker).
 * Pending edits are saved first (merged with the file's changes). */
QBR.fsReloadFromFile = async function (name, opts) {
  opts = opts || {};
  const link = (QBR._fsLinks || []).find(l => fileKey(l.name) === fileKey(name));
  if (!link) return false;
  if (!(await fsEnsurePermission(link.handle))) return false;
  if (link.side) {   // write contract: a side file re-syncs its own tab, never the dashboard sources
    const cur = await link.handle.getFile(), u8 = new Uint8Array(await cur.arrayBuffer());
    const newFp = QBR.fpOf(link.name, cur);
    try { QBR.journalMove(link.fp, newFp); } catch (e) {}
    fsSideMove(link.fp, newFp);
    link.fp = newFp; link.size = cur.size; link.lastModified = cur.lastModified; link.hash = fsHashBytes(u8);
    fsSetBase(link.name, u8);
    await fsSideNotify(link, u8, "reloaded");
    return true;
  }
  if (!opts.quiet) {
    let pending = 0;
    try { pending = QBR.journalOpsFor(link.fp).length; } catch (e) {}
    if (pending) {
      if (!confirm(`You have ${pending} unsaved edit${pending === 1 ? "" : "s"} for "${link.name}".\n\nOK = save them now (merged with the file's latest changes), then reload.\nCancel = do nothing.`)) return false;
      for (const k of (link.kinds || [])) await QBR.fsSaveKind(k);
      return true;
    }
  }
  const cur = await link.handle.getFile();
  const mem = await fsMemFile(cur, link.name);
  APP.files = (APP.files || []).map(it => fileKey(it.name) === fileKey(link.name) ? { name: link.name, blob: mem.file } : it);
  link.fp = QBR.fpOf(link.name, mem.file); link.size = cur.size; link.lastModified = cur.lastModified; link.hash = fsHashBytes(mem.bytes);
  fsSetBase(link.name, mem.bytes);
  try { await fsLinksSave(); } catch (e) {}
  await fsReloadModel();
  if (!opts.quiet) { try { QBR.persistNote("", `Reloaded ✓ ${link.name}`, 4000); } catch (e) {} }
  return true;
};

/* Review dialog for a merge: same-cell conflicts (pick Excel or Dashboard per
 * cell) and edits that can't be placed. Resolves to {index: "mine"|"theirs"} or
 * null (cancel). Tests can set QBR._mergeAutoResolve = m => choices|null. */
QBR.fsMergeReview = function (m, fileName) {
  if (typeof QBR._mergeAutoResolve === "function") return Promise.resolve(QBR._mergeAutoResolve(m));
  return new Promise(resolve => {
    const e = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    const ov = document.createElement("div");
    ov.className = "merge-ov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "merge-h");
    const rows = m.conflicts.map((c, i) =>
      `<tr><td><b>${e(c.sheet)}</b> ${e(c.ref)}<div class="small text-muted">${e(c.header)}${c.key ? " · " + e(c.key) : ""}</div></td>` +
      `<td><label class="merge-pick"><input type="radio" name="mc${i}" value="theirs" checked> ${e(c.theirs) || "<i>blank</i>"}</label></td>` +
      `<td><label class="merge-pick"><input type="radio" name="mc${i}" value="mine"> ${e(c.mine) || "<i>blank</i>"}</label></td></tr>`).join("");
    const un = m.unresolved.map(u => `<li><b>${e(u.sheet)}</b> ${e(u.ref)} → "${e(u.mine)}" — ${e(u.why)}</li>`).join("");
    ov.innerHTML = `<div class="merge-box">
      <h5 id="merge-h">"${e(fileName)}" was changed in Excel</h5>
      <p class="small">${m.applied} of your edits fit around the Excel changes and will be saved. ${m.conflicts.length ? "These cells were changed in <b>both</b> places — choose which value to keep:" : ""}</p>
      ${m.conflicts.length ? `<div class="merge-tools"><button type="button" class="btn btn-sm btn-outline-secondary" data-all="theirs">Keep all Excel values</button> <button type="button" class="btn btn-sm btn-outline-secondary" data-all="mine">Keep all dashboard values</button></div>
      <div class="merge-scroll"><table class="table table-sm"><thead><tr><th>Cell</th><th>Excel (file)</th><th>Dashboard (yours)</th></tr></thead><tbody>${rows}</tbody></table></div>` : ""}
      ${un ? `<div class="merge-warn"><b>${m.unresolved.length} edit${m.unresolved.length === 1 ? "" : "s"} can't be placed</b> — they'll be kept under <b>Unsaved edits → Couldn't be placed</b> so you can re-enter them:<ul>${un}</ul></div>` : ""}
      <div class="merge-act"><button type="button" class="btn btn-primary" id="merge-ok">Save</button> <button type="button" class="btn btn-outline-secondary" id="merge-cancel">Cancel — save nothing</button></div></div>`;
    document.body.appendChild(ov);
    const done = v => { ov.remove(); resolve(v); };
    ov.querySelectorAll("[data-all]").forEach(b => b.addEventListener("click", () => ov.querySelectorAll(`input[value="${b.dataset.all}"]`).forEach(x => { x.checked = true; })));
    ov.querySelector("#merge-cancel").addEventListener("click", () => done(null));
    ov.addEventListener("keydown", ev => { if (ev.key === "Escape") done(null); });
    ov.querySelector("#merge-ok").addEventListener("click", () => {
      const ch = {}; m.conflicts.forEach((c, i) => { const x = ov.querySelector(`input[name="mc${i}"]:checked`); ch[i] = x ? x.value : "theirs"; });
      done(ch);
    });
    setTimeout(() => { const b = ov.querySelector("#merge-ok"); if (b) b.focus(); }, 0);
  });
};

/* ---------- format-safe save engine (2026-10-06; docs/FORMAT_SAFE_SAVE.md) ---------- */
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
  if (res.applied === 0) return { mode: "no-changes", name: link.name, kind: kind };
  const out = await S.surgicalSave(buf, base, work);
  const bookType = fsBookType(link.name);
  let why = !out.ok ? "format-safe save can't handle this file yet (" + out.reason + ")" : null;
  if (!why && out.noChanges) return { mode: "no-changes", name: link.name, kind: kind };
  if (!why) why = await S.verifySurgical(buf, out.bytes, work);
  if (!why) why = fsVerifyBytes(out.bytes, bookType, work);
  if (why) {
    console.warn("[QBR] format-safe save blocked — " + why, out.stats);
    const r = QBR.fsDownloadKind(kind, "download-unsafe");
    return Object.assign(r, { name: link.name, reason: why });
  }
  // Last-moment check: if OneDrive/Excel changed the file while we worked, merge instead.
  // 2026-10-20 fix: never fall through to the write when the file changed and can't be merged,
  // or when it can't be re-read — that overwrote Excel's edits (docs/FORMAT_SAFE_SAVE.md promise).
  let changedNow = false, reReadFailed = false;
  try {
    const again = new Uint8Array(await (await link.handle.getFile()).arrayBuffer());
    changedNow = fsHashBytes(again) !== fsHashBytes(buf);
  } catch (e) { reReadFailed = true; }
  if (changedNow && QBR.xlsxMerge && fsBaseOf(link.name)) return await fsSaveMerged(kind, link, await link.handle.getFile(), 1);
  if (changedNow || reReadFailed) {
    const why = reReadFailed ? "could not re-check the file just before writing" : "the file changed in Excel while saving and can't be merged";
    console.warn("[QBR] format-safe save blocked — " + why);
    const r = QBR.fsDownloadKind(kind, "download-unsafe");
    return Object.assign(r, { name: link.name, reason: why });
  }
  const wr = await fsWriteBytes(kind, link, out.bytes);
  if (wr) return wr;
  (QBR._origWb || (QBR._origWb = {}))[link.fp] = work; // retained copy now matches the file
  const fresh = await link.handle.getFile();
  await QBR.fsAfterSave(kind, link, fresh);
  console.info("[QBR] format-safe save:", out.stats);
  return { mode: "file", name: link.name, applied: res.applied, skipped: res.skipped, notes: res.notes, engine: "format-safe", stats: out.stats, kind: kind };
}

/* Rebuild-download path (classic Export). Used when no file is linked, when the
 * browser lacks the File System Access API, and as a fallback. NOTE: this
 * regenerates the workbook from parsed data — formulas/layout of the original
 * are NOT preserved. The linked direct-save path below (patch-in-place)
 * preserves them. */
QBR.fsDownloadKind = function (kind, mode) {
  // 2026-10-28 (write contract): the adapter decides what a fallback copy is. Inventory and
  // Supplies rebuild from the model (as before); every other kind downloads a PATCHED COPY
  // of the last-saved bytes with its pending edits (edits stay pending). Before, every kind
  // fell through to the Inventory builder — an audit/storage/usage fallback had nothing to download.
  const ad = QBR.wbKinds && QBR.wbKinds.get(kind);
  if (ad && typeof ad.download === "function") return ad.download(mode);
  if (ad && ad.download !== "model") return fsDownloadPatchedCopy(kind, mode);
  const build = kind === "supplies" ? QBR.supBuildWorkbook : QBR.invBuildWorkbook;
  if (typeof build !== "function") return { mode: "no-data" };
  const b = build();
  if (!b || !b.wb) return { mode: "no-data" };
  XLSX.writeFile(b.wb, b.filename, { cellStyles: true });
  return { mode: mode || "download", filename: b.filename };
};

function fsDownloadPatchedCopy(kind, mode) {
  try {
    const link = QBR.fsGetLink(kind);
    const base = link && fsBaseOf(link.name);
    if (!base || typeof QBR.patchWorkbookFromJournal !== "function") return { mode: "no-data" };
    const wb = XLSX.read(base, { type: "array", cellStyles: true });
    const res = QBR.patchWorkbookFromJournal(kind, wb, link.fp);
    QBR._lastPatch = null;   // nothing was written to the file: no row positions to commit, nothing to clear
    if (!res.ok) return { mode: "no-data" };
    if (typeof QBR.sanitizeHyperlinkTargets === "function") QBR.sanitizeHyperlinkTargets(wb);
    const bookType = fsBookType(link.name);
    const filename = String(link.name).replace(/\.(xlsx|xlsm)$/i, "") + " (dashboard copy)." + (bookType === "xlsm" ? "xlsm" : "xlsx");
    XLSX.writeFile(wb, filename, { bookType: bookType, cellStyles: true, bookVBA: true });
    return { mode: mode || "download", filename: filename };
  } catch (e) { console.warn("[QBR] copy download failed:", e && e.message); return { mode: "no-data" }; }
}
QBR._fsDownloadPatchedCopy = fsDownloadPatchedCopy;

/* 2026-10-28 (write contract) — "side" files: a workbook a tab opens itself (MOS
 * Certificate Log, Trainers tracker) joins the SAME save path as a loaded file:
 * journal (keyed by its fingerprint), merge base, lineage, lock detection,
 * shared messages. QBR.fsAttachSide(kind, handle, bytes, opts) → link;
 * opts.onSaved(bytes, reason) is called after every save / reload of that file so the tab
 * re-parses its own model (side files are not dashboard sources: no loadItems). */
QBR._fsSideKinds = QBR._fsSideKinds || {};
QBR.fsAttachSide = async function (kind, handle, bytes, opts) {
  opts = opts || {};
  const f = await handle.getFile();
  const name = opts.name || f.name || handle.name;
  const u8 = bytes ? (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)) : new Uint8Array(await f.arrayBuffer());
  const fp = QBR.fpOf(name, f);
  QBR.fsDetachSide(kind);
  const link = { fp: fp, kinds: [kind], name: name, size: f.size, lastModified: f.lastModified, hash: fsHashBytes(u8), handle: handle,
    side: true, onSaved: typeof opts.onSaved === "function" ? opts.onSaved : null };
  QBR._fsLinks = (QBR._fsLinks || []).filter(l => !(l.side && fileKey(l.name) === fileKey(name))).concat([link]);
  (QBR._fsSideKinds[fp] || (QBR._fsSideKinds[fp] = new Set())).add(kind);
  fsSetBase(name, u8);
  try { if (typeof QBR.persistRefreshBadge === "function") QBR.persistRefreshBadge(); } catch (e) {}
  return link;
};
QBR.fsDetachSide = function (kind) {
  (QBR._fsLinks || []).filter(l => l.side && (l.kinds || []).indexOf(kind) >= 0).forEach(l => {
    const s = QBR._fsSideKinds[l.fp]; if (s) { s.delete(kind); if (!s.size) delete QBR._fsSideKinds[l.fp]; }
  });
  QBR._fsLinks = (QBR._fsLinks || []).filter(l => !(l.side && (l.kinds || []).indexOf(kind) >= 0));
};
function fsSideMove(oldFp, newFp) {
  if (!oldFp || !newFp || oldFp === newFp || !QBR._fsSideKinds[oldFp]) return;
  QBR._fsSideKinds[newFp] = QBR._fsSideKinds[oldFp]; delete QBR._fsSideKinds[oldFp];
}
async function fsSideNotify(link, u8, reason) {
  if (!link || !link.side || typeof link.onSaved !== "function") return;
  try { await link.onSaved(u8, reason); } catch (e) { console.warn("[QBR] side reload failed:", e && e.message); }
}
// Autosave (adapter.autosave): edits recorded in a burst save once, 300 ms after the last one.
const WB_AUTOSAVE = {};
QBR.wbScheduleSave = function (kind) {
  clearTimeout(WB_AUTOSAVE[kind]);
  WB_AUTOSAVE[kind] = setTimeout(async () => {
    delete WB_AUTOSAVE[kind];
    try { QBR.persistNote(kind, "Saving…", 15000); } catch (e) {}
    let r; try { r = await QBR.fsSaveKind(kind); } catch (e) { r = { mode: "error" }; }
    try { QBR.saveResultNote(kind, r); } catch (e) {}
    try { document.dispatchEvent(new CustomEvent("qbr:autosaved", { detail: { kind: kind, mode: r && r.mode } })); } catch (e) {}
  }, 300);
};

/* 2026-10-31 (audit L2): QBR.fsRebuildToLink removed — it rebuilt the whole workbook from the
 * model over the linked file (formulas/layout lost), bypassing format-safe save; nothing called it.
 * A changed file now always goes through merge-on-save or a downloaded copy. */

/* Save one inventory kind. Returns a result descriptor:
 *  file             — bytes written straight into the linked file (patch-in-place)
 *  no-changes       — linked file, but the journal held nothing to write
 *  download         — no link: fell back to Export download
 *  download-fallback— write failed: downloaded instead
 *  cancelled        — user declined the overwrite confirm
 *  denied           — permission not granted (then downloaded? no: aborted)
 *  no-data          — nothing to save */
/* 2026-10-20: one save at a time per kind, and saves to the same linked file run one after
 * another. A double-click on Save used to start two overlapping saves of the same journal;
 * the slower one could overwrite the faster one's appended rows. A repeat click while a
 * save is running now just waits for that save's result. */
QBR._fsInFlight = QBR._fsInFlight || {};
QBR._fsLinkQueue = QBR._fsLinkQueue || {};
QBR.fsSaveKind = function (kind) {
  if (QBR._fsInFlight[kind]) return QBR._fsInFlight[kind];
  const link = QBR.fsGetLink(kind);
  const qk = link ? "link:" + fileKey(link.name) : "kind:" + kind;
  const prev = QBR._fsLinkQueue[qk] || Promise.resolve();
  const run = prev.catch(() => {}).then(() => fsSaveKindNow(kind));
  QBR._fsInFlight[kind] = run;
  QBR._fsLinkQueue[qk] = run;
  const done = () => {
    if (QBR._fsInFlight[kind] === run) delete QBR._fsInFlight[kind];
    if (QBR._fsLinkQueue[qk] === run) delete QBR._fsLinkQueue[qk];
  };
  run.then(done, done);
  return run;
};
async function fsSaveKindNow(kind) {
  const link = QBR.fsGetLink(kind);
  if (!link || !QBR.fsSupported()) return QBR.fsDownloadKind(kind);
  try {
    if (!(await fsEnsurePermission(link.handle))) return { mode: "denied", name: link.name };
    const cur = await link.handle.getFile();
    const metaChanged = cur.size !== link.size || cur.lastModified !== link.lastModified;
    if (metaChanged) {
      // v1.25.1: same size but different timestamp = possible OneDrive sync noise.
      // Verify the actual content before refusing to write.
      let contentSame = false;
      if (cur.size === link.size && link.hash) {
        try {
          const curHash = fsHashBytes(new Uint8Array(await cur.arrayBuffer()));
          contentSame = (curHash === link.hash);
        } catch (e) { contentSame = false; }
      }
      if (!contentSame && QBR.saveEngine() === "format-safe" && QBR.xlsxMerge && fsBaseOf(link.name)) {
        // 2026-10-06: the file was edited in Excel — merge the dashboard's edits
        // into the current file instead of refusing (no re-link needed).
        return await fsSaveMerged(kind, link, cur, 1);
      }
      if (!contentSame) {
      // 2026-10-05: never rebuild over the linked file (that wiped formulas and
      // layout, and broke .xlsm files). Download a separate copy instead.
      alert(`"${link.name}" changed outside the dashboard since it was linked (for example, edited in Excel).\n\n` +
        `To keep that file safe, nothing will be written to it. Your changes will download as a separate copy.\n` +
        `To save directly again, check the Direct-save box for that file in the loaded-files list and pick the updated workbook.`);
      const r = QBR.fsDownloadKind(kind, "download-changed");
      return Object.assign(r, { name: link.name });
      }
      // Content is byte-identical: rebase the link to the new metadata and save normally.
      await QBR.fsRebaseLink(link, cur);
    }
    // 2026-10-06: format-safe save (js/xlsx-surgical.js). Writes only the
    // changed cells (and hyperlinks) into the file and copies every other part
    // untouched, so styles, conditional formatting, validation, tables, charts
    // and macros survive. Opt out per browser: localStorage "qbr-save-engine" = "legacy".
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
    if (res.applied === 0) return { mode: "no-changes", name: link.name, kind: kind };
    // Ask Excel to recalculate on open (patched cells have no calc chain).
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.CalcPr = { fullCalcOnLoad: true };
    try { delete wb.CalcChain; } catch (e) {}
    const bookType = fsBookType(link.name);
    if (bookType === "xlsm") await fsEnsureVba(wb, cur);
    if (typeof QBR.sanitizeHyperlinkTargets === "function") QBR.sanitizeHyperlinkTargets(wb);
    const bytes = XLSX.write(wb, { bookType: bookType, type: "array", cellStyles: true, bookVBA: true });
    const why = fsVerifyBytes(bytes, bookType, wb);
    if (why) {
      try { delete QBR._origWb[link.fp]; } catch (x) {}   // 2026-10-20: retained copy now holds unsaved rows; re-read next time
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
    // 2026-10-20: the legacy path patches the retained workbook in place; drop it so a retry
    // re-reads the file instead of appending the same rows twice.
    try { if (link && QBR._origWb && QBR.saveEngine() !== "format-safe") delete QBR._origWb[link.fp]; } catch (x) {}
    console.warn("[QBR] direct save failed, falling back to download:", e && e.message);
    return QBR.fsDownloadKind(kind, "download-fallback");
  }
};

/* v1.25.1 — Rebase a link to new file metadata after content-hash verification.
 * The bytes are proven identical, so this is a metadata-only move: every fp-keyed
 * structure (journal, kind map, retained workbook, model fps) migrates from the
 * old fingerprint to the new one. Parsed coordinates stay valid.
 */
QBR.fsRebaseLink = async function (link, freshFile) {
  const oldFp = link.fp;
  const newFp = QBR.fpOf(freshFile.name, freshFile);
  if (!oldFp || newFp === oldFp) return;
  try { QBR.journalMove(oldFp, newFp); } catch (e) {}
  try {
    if (QBR._kindByFp && QBR._kindByFp[oldFp]) {
      QBR._kindByFp[newFp] = QBR._kindByFp[oldFp];
      delete QBR._kindByFp[oldFp];
    }
  } catch (e) {}
  try { if (typeof QBR.fsCarryPatchState === "function") QBR.fsCarryPatchState(oldFp, newFp); } catch (e) {}
  try { fsSideMove(oldFp, newFp); } catch (e) {}
  try { if (QBR._audit && QBR._audit.fp === oldFp) QBR._audit.fp = newFp; } catch (e) {}
  try { if (QBR._auditMeta && QBR._auditMeta.fp === oldFp) QBR._auditMeta.fp = newFp; } catch (e) {}
  link.fp = newFp; link.size = freshFile.size; link.lastModified = freshFile.lastModified;
  try { await fsLinksSave(); } catch (e) {}
  try { await lineageRecordFile(link.name, freshFile); } catch (e) {}
  try { if (typeof QBR.fsRefreshStatus === "function") QBR.fsRefreshStatus(); } catch (e) {}
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
    // 2026-10-06: keep the saved bytes in memory — a File read from disk becomes
    // unreadable (NotReadableError) once Excel changes the file — and make them
    // the base for the next three-way merge.
    try {
      const mem = await fsMemFile(freshFile, link.name);
      if (typeof APP !== "undefined" && APP.files) APP.files = APP.files.map(it => fileKey(it.name) === fileKey(link.name) ? { name: link.name, blob: mem.file } : it);
      fsSetBase(link.name, mem.bytes);
    } catch (e) {}
    // 2. link metadata follows the new file state
    link.fp = newFp; link.size = freshFile.size; link.lastModified = freshFile.lastModified;
    try { link.hash = fsHashBytes(new Uint8Array(await freshFile.arrayBuffer())); } catch (e) {} // v1.25.1: keep the content hash in sync
    await fsLinksSave();
    try { await lineageRecordFile(link.name, freshFile); } catch (e) {} // lineage: this browser saved this version
    // 3. kind map: move the entry to the new fingerprint so later mutations
    //    record against the right file
    if (QBR._kindByFp && QBR._kindByFp[oldFp]) {
      QBR._kindByFp[newFp] = QBR._kindByFp[oldFp];
      delete QBR._kindByFp[oldFp];
    }
    fsSideMove(oldFp, newFp);   // write contract: side files keep their kind across saves
    // 3b. patch state: the retained workbook was patched in place, so its
    //     coordinates are still valid — carry them to the new fingerprint.
    //     (The rebuild path drops them beforehand via fsDropPatchState.)
    // 2026-10-20: the bytes are on disk, so the row positions this save stamped are now real.
    if (typeof QBR.patchCommitStamps === "function") QBR.patchCommitStamps(oldFp, kind);
    if (typeof QBR.fsCarryPatchState === "function") QBR.fsCarryPatchState(oldFp, newFp);
    // 4. journal entries now in the file — clear exactly those (2026-10-13: was the whole
    //    fingerprint, which also dropped edits made during the save, other kinds' pending
    //    edits and skipped ops). Whatever is left follows the file to its new fingerprint.
    const lp = QBR._lastPatch;
    if (lp && lp.fp === oldFp && lp.kind === kind && typeof QBR.journalClearIds === "function") {
      QBR.journalClearIds(oldFp, lp.ids, "saved");
      if (newFp && newFp !== oldFp && typeof QBR.journalMove === "function") QBR.journalMove(oldFp, newFp);
    } else {
      QBR.journalClearFp(oldFp);
    }
    QBR._lastPatch = null;
    // 5. persist the fresh bytes for session restore (side files are not session sources)
    if (!link.side && typeof cacheSession === "function") cacheSession();
    // 6. write contract: a side file's tab re-parses the saved bytes itself
    if (link.side) { const sb = fsBaseOf(link.name); await fsSideNotify(link, sb, "saved"); }
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
// 2026-10-28: one message per save outcome, shared by every Save button
// (Inventory, Supplies, Risky sign-ins editor, guided audit). Before, the
// "open in Excel" / write-failed / sync-busy / merge-cancelled outcomes had no
// message of their own and the badge just went back to "● N unsaved".
// Returns { text, ms, warn, short } — short is the guided-audit "Save all" summary form.
QBR.saveResultText = function (r) {
  r = r || { mode: "error" };
  const nm = r.name || "the linked file";
  const extra = (r.applied ? ` (${r.applied} change${r.applied === 1 ? "" : "s"}${r.skipped ? `, ${r.skipped} skipped` : ""})` : "");
  switch (r.mode) {
    case "file": return { text: `Saved ✓ ${r.name}${extra}${r.engine === "format-safe" ? " · formatting kept" : ""}${r.merged ? " · merged with Excel changes" : ""}`, ms: 6000, short: "saved ✓" };
    case "no-changes": return { text: "No changes to save" + (r.kind && r.kind !== "assets" && r.kind !== "supplies" ? " for " + r.kind : ""), ms: 3000, short: "no changes" };
    case "reloaded": return { text: `No dashboard edits to save — reloaded ${nm} with the changes made in Excel`, ms: 6000, short: "reloaded" };
    case "download": case "download-fallback": return { text: `Downloaded ${r.filename || ""} — tick Direct save for the file to save into it`, ms: 6000, short: "downloaded" };
    case "download-unsafe": return { text: `Not saved to ${nm} — ${r.reason || "safety check failed"}. Downloaded a separate copy instead.`, ms: 15000, warn: true, short: "not saved (copy downloaded)" };
    case "download-changed": return { text: `${nm} changed in Excel — downloaded a separate copy. Re-link to save directly.`, ms: 15000, warn: true, short: "copy downloaded" };
    case "locked": return { text: `Not saved — "${nm}" is open in Excel. Close it there, then click Save again. Your edits are kept.`, ms: 30000, warn: true, short: "open in Excel — close it and save again" };
    case "write-failed": return { text: `Not saved — could not write "${nm}"${r.reason ? " (" + r.reason + ")" : ""}. Your edits are kept.`, ms: 20000, warn: true, short: "not saved (write failed)" };
    case "busy": return { text: `Not saved — "${nm}" keeps changing (sync in progress?). Try again in a moment. Your edits are kept.`, ms: 20000, warn: true, short: "not saved (file busy)" };
    case "merge-cancelled": return { text: "Save cancelled — nothing was written. Your edits are kept.", ms: 10000, warn: true, short: "cancelled" };
    case "denied": return { text: "Not saved — permission to the file was denied. Your edits are kept.", ms: 10000, warn: true, short: "permission denied" };
    case "no-data": return null;
    default: return { text: "Not saved — something went wrong (see the browser console). Your edits are kept.", ms: 15000, warn: true, short: "failed" };
  }
};
QBR.saveResultNote = function (kind, r, noDataMsg) {
  const t = QBR.saveResultText(r);
  if (!t) { if (noDataMsg) alert(noDataMsg); return; }
  QBR.persistNote(kind, t.text, t.ms, { warn: !!t.warn });
};
QBR.supSaveDone = function (r) {
  if (!r) return;
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
  QBR.saveResultNote("supplies", r, "Load a supplies workbook first.");   // after the re-render; the note also survives later ones
};
QBR.invSaveDone = function (r) {
  if (!r) return;
  if (typeof renderAll === "function") { try { renderAll(); } catch (e) {} }
  QBR.saveResultNote("assets", r, "Load an inventory workbook first.");   // after the re-render; the note also survives later ones
};

// Adaptive Save button for an inventory kind: direct "Save to Excel" when a
// file is linked, otherwise the classic "Export workbook" download.
QBR.saveButtonHtml = function (kind, id) {
  const link = (typeof QBR.fsGetLink === "function") ? QBR.fsGetLink(kind) : null;
  const esc = s => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const badge = `<span data-unsaved="${kind}" class="small text-muted align-self-center" role="status" aria-live="polite"></span>`;
  if (link) {
    return `<button type="button" class="btn btn-sm btn-primary" id="${id}" title="Write directly to ${esc(link.name)}">💾 Save to Excel</button>` + badge;
  }
  return `<button type="button" class="btn btn-sm btn-outline-secondary" id="${id}" title="Download the workbook (check Direct-save on the file for direct save)">⭳ Export workbook</button>` + badge;
};

// Boot: load remembered links, paint status, hide the Link button where
// the File System Access API does not exist.
QBR.fsInit = function () {
  fsLinksLoad();
  QBR.persistRefreshBadge();
};

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", QBR.fsInit);
else QBR.fsInit();
