/* ============================================================================
 * QBR Dashboard — unsaved-edits journal (js/journal.js)
 * v2.0.0 — 2026-10-06 (dashboard v1.30.0, "journal safety")
 *
 * Every dashboard edit (inventory, supplies, audit, storage, usage) is recorded
 * as a replayable op {id, kind, op, args, ts[, sig]} keyed by the source file's
 * fingerprint (name|size|lastModified). This module replaces the v1 journal that
 * lived in persist.js. What changed, and why:
 *
 *  1. In-memory store is authoritative for the session. It is written to
 *     localStorage AND mirrored to IndexedDB (auto-backup) on every change.
 *     Startup merges both (union by op id, removals kept as tombstones), so a
 *     full localStorage no longer loses edits silently: the backup still has
 *     them, a banner says so, and closing the tab asks first.
 *  2. No silent caps: the old 500-op cap and 90-day deletion are gone. The
 *     unsaved badge turns red at WARN_OPS edits for one file instead.
 *  3. Namespaced per dashboard folder. On file:// every local page shares ONE
 *     localStorage, so a test copy (QBR-Dashboard_dev) used to share — and
 *     clear — the live folder's pending edits. Each folder now has its own key;
 *     the other folders' edits are listed, never touched, unless you move them.
 *     v1 data (key qbr-inv-journal-v1) is claimed when its file is loaded.
 *  4. Orphaned edits (recorded against an older copy of a file) are surfaced in
 *     the "Unsaved edits" panel with Apply to the loaded file / Export / Discard.
 *     Row-addressed edits carry a signature (column A of the row when recorded)
 *     so they can be re-placed by school/key instead of by row number.
 *  5. Edits a merge couldn't place are PARKED (kept and listed) instead of
 *     being dropped when the journal is cleared after a save.
 *  6. Export / Import of pending edits as JSON. Cleared batches (after a save)
 *     are kept 14 days in IndexedDB and can be exported.
 *
 * Public API (all on QBR): journalRecord, journalReplayFor, journalClearFp,
 * journalUnsavedCount, journalEntries, journalOpsFor, journalMove, journalPark,
 * journalOrphans, journalForeign, journalApplyOrphan, journalDiscard,
 * journalMoveForeign, journalExportData, journalExport, journalImport,
 * journalHealth, journalPanel, persistRefreshBadge, persistNote, _fpsForKind,
 * journalReady (Promise; loadItems waits for it before replaying).
 * ==========================================================================*/
(function () {
  "use strict";
  var QBR = window.QBR = window.QBR || {};
  QBR.JOURNAL_VERSION = "2.0.0";

  var LEGACY_KEY = "qbr-inv-journal-v1";
  var KEY_PREFIX = "qbr-inv-journal-v2:";
  var WARN_OPS = 400;
  var DAY = 24 * 3600 * 1000;
  var GONE_TTL = 30 * DAY, TRASH_TTL = 14 * DAY, TRASH_MAX = 30;
  // Ops addressed by (fp, sheet, row): index of the row argument.
  var ROW_OPS = { auditUpdateCell: 2, auditPasteRow: 2, storagePasteRow: 2, usagePasteRow: 2 };
  var FP_OPS = { auditUpdateCell: 1, auditPasteRow: 1, storagePasteRow: 1, usagePasteRow: 1, auditAddRow: 1 };

  /* ------------------------------ helpers -------------------------------- */
  function strHash(s) {           // cyrb53 over UTF-16 code units
    var h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  }
  function appDir() {
    try {
      var l = window.location, p = decodeURIComponent(l.pathname || "").replace(/[^\/]*$/, "");
      if (l.protocol === "file:") return "file://" + p.toLowerCase();   // Windows paths are case-insensitive
      return l.origin + p;
    } catch (e) { return "unknown"; }
  }
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function nameKey(n) { return String(n || "").trim().toLowerCase(); }
  function normSig(v) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().toLowerCase(); }
  function appFiles() { try { return (typeof APP !== "undefined" && APP.files) || []; } catch (e) { return []; } }
  function fpOf(f) { return QBR.fpOf ? QBR.fpOf(f.name, f.blob) : String(f.name); }
  function loadedFps() { return appFiles().map(fpOf); }
  function gk(fp, id) { return fp + "\u0001" + id; }
  var seq = 0;
  function newId(prefix) { return (prefix || "op") + Date.now().toString(36) + (seq++).toString(36) + Math.floor(Math.random() * 1e6).toString(36); }

  QBR.JOURNAL_NS_LABEL = appDir();
  var NS = strHash(QBR.JOURNAL_NS_LABEL);
  var LS_KEY = KEY_PREFIX + NS;
  QBR.JOURNAL_KEY = LS_KEY;

  /* ----------------------------- the store ------------------------------- */
  function blank() { return { v: 2, ns: QBR.JOURNAL_NS_LABEL, files: {}, gone: {}, parked: [], updated: 0 }; }
  // Accepts a v2 envelope or a v1 map {fp: {fileName, ops}}.
  function norm(x) {
    var b = blank();
    if (!x || typeof x !== "object") return b;
    if (x.v === 2 && x.files) {
      b.ns = x.ns || ""; b.files = x.files || {}; b.gone = x.gone || {};
      b.parked = Array.isArray(x.parked) ? x.parked : []; b.updated = x.updated || 0;
      return b;
    }
    b.ns = "";
    Object.keys(x).forEach(function (fp) {
      var e = x[fp];
      if (e && Array.isArray(e.ops)) b.files[fp] = { fileName: e.fileName || "", ops: e.ops.filter(Boolean) };
    });
    return b;
  }
  // Union of two stores. Tombstones ("gone") always win, so a removal is never undone by an older copy.
  function merge(a, b) {
    var out = blank(); out.ns = a.ns || b.ns || out.ns;
    [a, b].forEach(function (s) { Object.keys(s.gone || {}).forEach(function (k) { out.gone[k] = Math.max(out.gone[k] || 0, s.gone[k]); }); });
    [a, b].forEach(function (s) {
      Object.keys(s.files || {}).forEach(function (fp) {
        var src = s.files[fp]; if (!src || !Array.isArray(src.ops)) return;
        var t = out.files[fp] || (out.files[fp] = { fileName: src.fileName || "", ops: [] });
        if (!t.fileName && src.fileName) t.fileName = src.fileName;
        var have = {}; t.ops.forEach(function (o) { have[o.id] = 1; });
        src.ops.forEach(function (o) {
          if (!o || !o.id || have[o.id] || out.gone[gk(fp, o.id)]) return;
          t.ops.push(o); have[o.id] = 1;
        });
      });
    });
    Object.keys(out.files).forEach(function (fp) {
      var e = out.files[fp];
      e.ops = e.ops.map(function (o, i) { return [o, i]; }).sort(function (x, y) { return ((x[0].ts || 0) - (y[0].ts || 0)) || (x[1] - y[1]); }).map(function (p) { return p[0]; });
      if (!e.ops.length) delete out.files[fp];
    });
    var pid = {};
    [a, b].forEach(function (s) {
      (s.parked || []).forEach(function (p) {
        if (!p || !p.id || pid[p.id] || out.gone[gk("parked", p.id)]) return;
        pid[p.id] = 1; out.parked.push(p);
      });
    });
    out.updated = Math.max(a.updated || 0, b.updated || 0);
    return out;
  }
  QBR._journalMerge = merge; QBR._journalNorm = norm;   // exposed for tests

  function countOps(s, fps) {
    var n = 0;
    Object.keys(s.files).forEach(function (fp) { if (!fps || fps.indexOf(fp) >= 0) n += s.files[fp].ops.length; });
    return n;
  }
  function lsRead(key) { try { var raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : null; } catch (e) { return null; } }

  var MEM = blank();
  var HEALTH = { local: "unknown", idb: "unknown", persisted: null, lastError: "", recovered: 0, bytes: 0, dismissed: false };

  /* ---- IndexedDB mirror (same db/store as the session cache) ---- */
  function idb() {
    return new Promise(function (res, rej) {
      if (!window.indexedDB) return rej(new Error("no-indexeddb"));
      var rq; try { rq = indexedDB.open("qbr-cache", 1); } catch (e) { return rej(e); }
      rq.onupgradeneeded = function () { try { rq.result.createObjectStore("session"); } catch (e) {} };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
    });
  }
  function idbPut(key, val) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction("session", "readwrite");
        tx.objectStore("session").put(val, key);
        tx.oncomplete = function () { res(); }; tx.onerror = function () { rej(tx.error); }; tx.onabort = function () { rej(tx.error); };
      });
    });
  }
  function idbGet(key) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var rq = db.transaction("session", "readonly").objectStore("session").get(key);
        rq.onsuccess = function () { res(rq.result == null ? null : rq.result); }; rq.onerror = function () { rej(rq.error); };
      });
    });
  }
  var IDB_KEY = "journal:" + NS, TRASH_KEY = "journal-trash:" + NS;

  function persist() {
    MEM.updated = Date.now();
    var cut = Date.now() - GONE_TTL;
    Object.keys(MEM.gone).forEach(function (k) { if (MEM.gone[k] < cut) delete MEM.gone[k]; });
    Object.keys(MEM.files).forEach(function (fp) { if (!MEM.files[fp].ops.length) delete MEM.files[fp]; });
    var json = JSON.stringify(MEM);
    HEALTH.bytes = json.length;
    try {
      if (QBR._journalLsFail) throw Object.assign(new Error("simulated"), { name: "QuotaExceededError" });   // test hook
      localStorage.setItem(LS_KEY, json);
      HEALTH.local = "ok";
    } catch (e) {
      HEALTH.local = "failed"; HEALTH.lastError = (e && e.name) || String(e);
      console.warn("[QBR] journal: browser storage write failed (" + HEALTH.lastError + ") — kept in memory + IndexedDB backup");
    }
    var p = idbPut(IDB_KEY, json).then(function () { HEALTH.idb = "ok"; }, function (e) { HEALTH.idb = "failed"; console.warn("[QBR] journal: IndexedDB backup failed", e && e.message); });
    QBR._journalLastWrite = p.then(refreshUi);
    refreshUi();
    return p;
  }

  // Trash: batches removed from the journal (after a save/export/discard), kept 14 days in IndexedDB.
  function trashAdd(fp, entry, reason) {
    if (!entry || !entry.ops || !entry.ops.length) return Promise.resolve();
    var item = { id: newId("tr"), fp: fp, fileName: entry.fileName || "", ops: entry.ops.slice(), reason: reason || "", at: Date.now() };
    var p = idbGet(TRASH_KEY).then(function (raw) {
      var list = []; try { list = raw ? JSON.parse(raw) : []; } catch (e) { list = []; }
      var cut = Date.now() - TRASH_TTL;
      list = list.filter(function (t) { return t && t.at > cut; });
      list.push(item);
      if (list.length > TRASH_MAX) list = list.slice(-TRASH_MAX);
      return idbPut(TRASH_KEY, JSON.stringify(list));
    }).catch(function () {});
    QBR._journalLastTrash = p;
    return p;
  }
  QBR.journalTrash = function () {
    return idbGet(TRASH_KEY).then(function (raw) {
      var list = []; try { list = raw ? JSON.parse(raw) : []; } catch (e) {}
      var cut = Date.now() - TRASH_TTL;
      return list.filter(function (t) { return t && t.at > cut; });
    }).catch(function () { return []; });
  };

  /* ---- startup: localStorage now, IndexedDB backup merged before the first replay ---- */
  (function initSync() {
    try { localStorage.getItem(LS_KEY); HEALTH.local = "ok"; } catch (e) { HEALTH.local = "failed"; HEALTH.lastError = (e && e.name) || String(e); }
    var own = lsRead(LS_KEY);
    MEM = own ? norm(own) : blank();
  })();
  var replayed = {};   // fps already replayed into the current model
  QBR.journalReady = new Promise(function (resolve) {
    var settled = false;
    var done = function () { if (!settled) { settled = true; resolve(); } };
    var t = setTimeout(done, 2500);   // never block loading on a stuck IndexedDB
    idbGet(IDB_KEY).then(function (raw) {
      HEALTH.idb = "ok";
      if (!raw) return;
      var backup = norm(typeof raw === "string" ? JSON.parse(raw) : raw);
      var before = countOps(MEM);
      MEM = merge(MEM, backup);
      var added = countOps(MEM) - before;
      if (added > 0) {
        HEALTH.recovered = added;
        console.info("[QBR] journal: recovered " + added + " edit(s) from the IndexedDB backup");
        persist();
        // Late arrival (after a replay already ran): re-parse so they apply.
        var late = Object.keys(backup.files).some(function (fp) { return replayed[fp]; });
        if (late && settled && typeof loadItems === "function") { try { loadItems(appFiles().slice(), {}); } catch (e) {} }
      }
    }).catch(function () { HEALTH.idb = "failed"; }).then(function () { clearTimeout(t); done(); refreshUi(); });
  });

  var askedPersist = false;
  function requestPersistOnce() {
    if (askedPersist) return; askedPersist = true;
    try {
      if (!navigator.storage || !navigator.storage.persisted) return;
      navigator.storage.persisted().then(function (p) {
        HEALTH.persisted = p;
        if (!p && navigator.storage.persist) return navigator.storage.persist().then(function (q) { HEALTH.persisted = q; });
      }).catch(function () {});
    } catch (e) {}
  }

  /* ------------------------------- core API ------------------------------ */
  // Fingerprints currently loaded that contain a given kind (rebuilt by the parsers).
  QBR._fpsForKind = function (kind) {
    var m = QBR._kindByFp || {};
    return Object.keys(m).filter(function (fp) { try { return m[fp] && m[fp].has(kind); } catch (e) { return false; } });
  };
  function rowSig(fp, op, args) {
    if (!(op in ROW_OPS)) return null;
    try {
      var wb = (QBR._origWb || {})[fp], ws = wb && wb.Sheets && wb.Sheets[args[1]];
      var c = ws && ws["A" + args[ROW_OPS[op]]];
      var v = c ? (c.w != null ? c.w : c.v) : null;
      return v == null || String(v).trim() === "" ? null : String(v);
    } catch (e) { return null; }
  }
  function fileNameOf(fp) {
    var f = appFiles().filter(function (x) { return fpOf(x) === fp; })[0];
    return f ? f.name : "";
  }

  QBR.journalRecord = function (kind, op, args) {
    if (QBR._replaying) return;
    var fps = QBR._fpsForKind(kind), fp = null;
    if (op in FP_OPS && args && fps.indexOf(args[0]) >= 0) fp = args[0];   // row ops name their file
    if (!fp) fp = fps[0];
    if (!fp) return;
    var e = MEM.files[fp] || (MEM.files[fp] = { fileName: "", ops: [] });
    e.fileName = fileNameOf(fp) || e.fileName;
    var o = { id: newId("op"), kind: kind, op: op, args: args || [], ts: Date.now() };
    var sig = rowSig(fp, op, o.args);
    if (sig != null) o.sig = sig;
    e.ops.push(o);
    persist();
    requestPersistOnce();
  };

  // v1 data: claim entries for the files being loaded (others stay listed as "previous version").
  function claimLegacy(fps) {
    var legacy = lsRead(LEGACY_KEY);
    if (!legacy || typeof legacy !== "object") return 0;
    var moved = 0, part = {};
    fps.forEach(function (fp) {
      if (legacy[fp] && Array.isArray(legacy[fp].ops) && legacy[fp].ops.length) { part[fp] = legacy[fp]; moved += legacy[fp].ops.length; delete legacy[fp]; }
    });
    if (!moved) return 0;
    MEM = merge(MEM, norm(part));
    persist();
    try {
      if (Object.keys(legacy).length) localStorage.setItem(LEGACY_KEY, JSON.stringify(legacy));
      else localStorage.removeItem(LEGACY_KEY);
    } catch (e) {}
    console.info("[QBR] journal: moved " + moved + " edit(s) from the v1 journal");
    return moved;
  }

  // Replay stored ops for the loaded files onto the freshly parsed model.
  QBR.journalReplayFor = function (files) {
    var fps = (files || []).map(fpOf);
    claimLegacy(fps);
    var n = 0;
    replayed = {};
    QBR._replaying = true;
    try {
      fps.forEach(function (fp) {
        replayed[fp] = 1;
        var e = MEM.files[fp];
        if (!e || !e.ops.length) return;
        e.ops.forEach(function (o) {
          var fn = QBR[o.op];
          if (typeof fn !== "function") return;
          try { fn.apply(null, o.args || []); n++; }
          catch (err) { console.warn("[QBR] journal replay failed:", o.op, err && err.message); }
        });
      });
    } finally { QBR._replaying = false; }
    refreshUi();
    return n;
  };

  QBR.journalEntries = function () { return MEM.files; };
  QBR.journalOpsFor = function (fp, kind) {
    var e = MEM.files[fp];
    return e ? e.ops.filter(function (o) { return !kind || o.kind === kind; }) : [];
  };
  function removeOps(fp, ids) {
    var e = MEM.files[fp]; if (!e) return;
    var now = Date.now(), set = {};
    ids.forEach(function (id) { set[id] = 1; MEM.gone[gk(fp, id)] = now; });
    e.ops = e.ops.filter(function (o) { return !set[o.id]; });
    if (!e.ops.length) delete MEM.files[fp];
  }
  // Drop the journal for one fingerprint (its edits are now in the file). Kept 14 days in the trash.
  QBR.journalClearFp = function (fp, reason) {
    if (!fp) return;
    var e = MEM.files[fp];
    if (e && e.ops.length) {
      trashAdd(fp, e, reason || "saved");
      removeOps(fp, e.ops.map(function (o) { return o.id; }));
      persist();
    }
    refreshUi();
  };
  // Same bytes, new fingerprint (OneDrive touched the timestamp): move the ops.
  QBR.journalMove = function (oldFp, newFp) {
    if (!oldFp || !newFp || oldFp === newFp) return;
    var e = MEM.files[oldFp]; if (!e) return;
    var t = MEM.files[newFp] || (MEM.files[newFp] = { fileName: e.fileName, ops: [] });
    var have = {}; t.ops.forEach(function (o) { have[o.id] = 1; });
    t.ops = e.ops.filter(function (o) { return !have[o.id]; }).concat(t.ops);
    t.fileName = e.fileName || t.fileName;
    removeOps(oldFp, e.ops.map(function (o) { return o.id; }));
    persist();
  };
  QBR.journalUnsavedCount = function (kind) {
    var n = 0;
    loadedFps().forEach(function (fp) { n += QBR.journalOpsFor(fp, kind).length; });
    return n;
  };
  function maxPerFile() {
    var m = 0; loadedFps().forEach(function (fp) { m = Math.max(m, QBR.journalOpsFor(fp).length); }); return m;
  }

  // Merge edits that could not be placed: kept and listed instead of dropped.
  QBR.journalPark = function (fileName, items) {
    (items || []).forEach(function (u) {
      MEM.parked.push({ id: newId("pk"), fileName: fileName || "", sheet: u.sheet || "", ref: u.ref || "", header: u.header || "",
        key: u.key || "", value: u.mine == null ? "" : String(u.mine), why: u.why || "", at: Date.now() });
    });
    if (items && items.length) persist();
  };
  QBR.journalParked = function () { return MEM.parked.slice(); };
  QBR.journalUnpark = function (ids) {
    var set = {}, now = Date.now();
    (ids || MEM.parked.map(function (p) { return p.id; })).forEach(function (id) { set[id] = 1; MEM.gone[gk("parked", id)] = now; });
    MEM.parked = MEM.parked.filter(function (p) { return !set[p.id]; });
    persist();
  };

  /* --------------------- orphans / other folders -------------------------- */
  function loadedByName(name) {
    var f = appFiles().filter(function (x) { return nameKey(x.name) === nameKey(name); })[0];
    return f ? { name: f.name, fp: fpOf(f) } : null;
  }
  function targetFor(entry) {
    var t = loadedByName(entry.fileName);
    if (t) return t;
    var kinds = {}; entry.ops.forEach(function (o) { kinds[o.kind] = 1; });
    var ks = Object.keys(kinds);
    if (ks.length !== 1) return null;
    var fps = QBR._fpsForKind(ks[0]).filter(function (fp) { return loadedFps().indexOf(fp) >= 0; });
    if (fps.length !== 1) return null;
    return { name: fileNameOf(fps[0]), fp: fps[0], byKind: true };
  }
  function summarize(fp, e) {
    var ts = e.ops.map(function (o) { return o.ts || 0; });
    var t = targetFor(e);
    return { fp: fp, fileName: e.fileName, count: e.ops.length, first: Math.min.apply(null, ts), last: Math.max.apply(null, ts),
      kinds: Object.keys(e.ops.reduce(function (m, o) { m[o.kind] = 1; return m; }, {})), target: t ? t.name : null };
  }
  // Entries in THIS folder's journal whose exact file is not loaded.
  QBR.journalOrphans = function () {
    var fps = loadedFps();
    return Object.keys(MEM.files).filter(function (fp) { return fps.indexOf(fp) < 0 && MEM.files[fp].ops.length; })
      .map(function (fp) { return summarize(fp, MEM.files[fp]); });
  };
  // Journals of OTHER dashboard folders (and the v1 journal) in this browser.
  QBR.journalForeign = function () {
    var out = [], fps = loadedFps(), keys = [];
    try { for (var i = 0; i < localStorage.length; i++) keys.push(localStorage.key(i)); } catch (e) {}
    keys.forEach(function (k) {
      if (k === LS_KEY || !(k === LEGACY_KEY || k.indexOf(KEY_PREFIX) === 0)) return;
      var s = norm(lsRead(k));
      var entries = Object.keys(s.files).map(function (fp) {
        var e = s.files[fp];
        return { fp: fp, fileName: e.fileName, count: e.ops.length, loaded: fps.indexOf(fp) >= 0, sameName: !!loadedByName(e.fileName) };
      }).filter(function (x) { return x.count; });
      if (entries.length || s.parked.length) out.push({ key: k, label: k === LEGACY_KEY ? "previous dashboard version (before v1.30)" : (s.ns || "another dashboard folder"), entries: entries, parked: s.parked.length });
    });
    return out;
  };
  function locateRow(fp, sheet, sig) {
    if (sig == null) return { why: "recorded before v1.30 — its row can't be matched safely" };
    var wb = (QBR._origWb || {})[fp], ws = wb && wb.Sheets && wb.Sheets[sheet];
    if (!ws) return { why: 'sheet "' + sheet + '" is not in the loaded file' };
    var last = 1;
    try { last = XLSX.utils.decode_range(ws["!ref"]).e.r + 1; } catch (e) {}
    var want = normSig(sig), hits = [];
    for (var r = 1; r <= last; r++) {
      var c = ws["A" + r]; if (!c) continue;
      if (normSig(c.w != null ? c.w : c.v) === want) hits.push(r);
    }
    if (hits.length === 1) return { row: hits[0] };
    return { why: hits.length ? '"' + sig + '" appears ' + hits.length + " times in " + sheet : 'no row for "' + sig + '" in ' + sheet };
  }
  // Re-apply an orphaned entry to the loaded copy of its file. Applied ops are
  // re-recorded against the loaded file; the rest stay, with the reason.
  QBR.journalApplyOrphan = function (fp) {
    var e = MEM.files[fp], res = { applied: 0, kept: 0, notes: [] };
    if (!e) return res;
    var target = targetFor(e);
    if (!target || target.fp === fp) { res.kept = e.ops.length; res.notes.push("no loaded file to apply to"); return res; }
    var done = [];
    e.ops.forEach(function (o) {
      var keep = function (why) { o.lastTry = why; res.kept++; if (res.notes.indexOf(why) < 0) res.notes.push(why); };
      var fn = QBR[o.op];
      if (typeof fn !== "function") return keep("unknown action " + o.op);
      var args = (o.args || []).slice();
      if (o.op in FP_OPS) {
        args[0] = target.fp;
        if (o.op in ROW_OPS) {
          var loc = locateRow(target.fp, args[1], o.sig);
          if (loc.why) return keep(loc.why);
          args[ROW_OPS[o.op]] = loc.row;
        }
      }
      var ok;
      try { ok = fn.apply(null, args); } catch (err) { return keep((err && err.message) || String(err)); }
      if (ok === false) return keep("the row or item was not found in the loaded file");
      done.push(o.id); res.applied++;
    });
    if (done.length) { trashAdd(fp, { fileName: e.fileName, ops: e.ops.filter(function (o) { return done.indexOf(o.id) >= 0; }) }, "applied to " + target.name); removeOps(fp, done); }
    persist();
    if (res.applied && typeof renderAll === "function") { try { renderAll(); } catch (x) {} }
    return res;
  };
  QBR.journalDiscard = function (fp) { QBR.journalClearFp(fp, "discarded"); };
  // Move another folder's journal into this one (removed there). Loaded files re-parse so the edits show.
  QBR.journalMoveForeign = function (key) {
    if (key === LS_KEY) return { moved: 0 };
    var s = norm(lsRead(key)), moved = countOps(s) + s.parked.length;
    if (!moved) return { moved: 0 };
    s.gone = {};   // their tombstones are about their copies
    MEM = merge(MEM, s);
    persist();
    try { localStorage.removeItem(key); } catch (e) {}
    var touchesLoaded = Object.keys(s.files).some(function (fp) { return loadedFps().indexOf(fp) >= 0; });
    var p = touchesLoaded && typeof loadItems === "function" ? loadItems(appFiles().slice(), {}) : Promise.resolve();
    return { moved: moved, reloaded: touchesLoaded, done: p };
  };

  /* --------------------------- export / import --------------------------- */
  QBR.journalExportData = function (opts) {
    opts = opts || {};
    var files = {};
    Object.keys(MEM.files).forEach(function (fp) { if (!opts.fps || opts.fps.indexOf(fp) >= 0) files[fp] = MEM.files[fp]; });
    var data = { format: "qbr-unsaved-edits", version: 1, exportedAt: new Date().toISOString(), app: QBR.VERSION || "", source: QBR.JOURNAL_NS_LABEL,
      files: JSON.parse(JSON.stringify(files)), parked: opts.fps ? [] : MEM.parked.slice() };
    if (opts.trash) data.trash = opts.trash;
    return data;
  };
  function download(obj, name) {
    var text = JSON.stringify(obj, null, 1);
    if (typeof QBR._journalDownloadHook === "function") return QBR._journalDownloadHook(name, text);
    var url = URL.createObjectURL(new Blob([text], { type: "application/json" }));
    var a = document.createElement("a"); a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 500);
    return name;
  }
  function stamp() { var d = new Date(), p = function (n) { return String(n).padStart(2, "0"); }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes()); }
  QBR.journalExport = function (opts) { return download(QBR.journalExportData(opts), "qbr-unsaved-edits-" + stamp() + ".json"); };
  QBR.journalExportTrash = function () {
    return QBR.journalTrash().then(function (list) {
      return download({ format: "qbr-unsaved-edits", version: 1, exportedAt: new Date().toISOString(), app: QBR.VERSION || "", source: QBR.JOURNAL_NS_LABEL, files: {}, parked: [], trash: list }, "qbr-saved-batches-" + stamp() + ".json");
    });
  };
  // Import an exported file. Ops are de-duplicated by id; loaded files re-parse to show them.
  QBR.journalImport = function (text) {
    var d;
    try { d = typeof text === "string" ? JSON.parse(text) : text; } catch (e) { return { ok: false, error: "not a JSON file" }; }
    if (!d || d.format !== "qbr-unsaved-edits" || typeof d.files !== "object") return { ok: false, error: "not a QBR unsaved-edits export" };
    var inc = norm({ v: 2, files: d.files || {}, parked: d.parked || [], gone: {} });
    var before = countOps(MEM), pBefore = MEM.parked.length;
    // Imported ops override this folder's tombstones only if the user imports them on purpose.
    Object.keys(inc.files).forEach(function (fp) { inc.files[fp].ops.forEach(function (o) { delete MEM.gone[gk(fp, o.id)]; }); });
    inc.parked.forEach(function (p) { delete MEM.gone[gk("parked", p.id)]; });
    MEM = merge(MEM, inc);
    var added = countOps(MEM) - before, parked = MEM.parked.length - pBefore;
    persist();
    var touchesLoaded = Object.keys(inc.files).some(function (fp) { return loadedFps().indexOf(fp) >= 0; });
    var p = added && touchesLoaded && typeof loadItems === "function" ? loadItems(appFiles().slice(), {}) : Promise.resolve();
    return { ok: true, added: added, parked: parked, skipped: countOps(inc) - added, reloaded: !!(added && touchesLoaded), done: p };
  };

  QBR.journalHealth = function () {
    return { local: HEALTH.local, idb: HEALTH.idb, persisted: HEALTH.persisted, lastError: HEALTH.lastError, recovered: HEALTH.recovered,
      bytes: HEALTH.bytes, key: LS_KEY, ns: QBR.JOURNAL_NS_LABEL, unsaved: QBR.journalUnsavedCount(), orphans: QBR.journalOrphans().length, parked: MEM.parked.length };
  };

  /* ---------------------------------- UI ---------------------------------- */
  // Fill every [data-unsaved] badge with the current unsaved-entry count. Click → Unsaved edits panel.
  QBR.persistRefreshBadge = function () {
    try {
      var big = maxPerFile() >= WARN_OPS;
      document.querySelectorAll("[data-unsaved]").forEach(function (el) {
        if (el.dataset.persistBusy) return;
        var n = QBR.journalUnsavedCount(el.dataset.unsaved);
        el.innerHTML = n
          ? ' <button type="button" class="badge ' + (big ? "bg-danger" : "bg-warning text-dark") + ' qbr-unsaved-badge border-0" title="' +
            (big ? "Many unsaved edits — save to Excel soon. " : "") + 'Saved in the dashboard, not yet written to the Excel file. Click to review or export.">● ' + n + " unsaved</button>"
          : "";
      });
    } catch (e) {}
  };
  // Transient status text in the same badge slot ("Saving…", "Saved ✓").
  QBR.persistNote = function (kind, html, ms) {
    try {
      var el = document.querySelector('[data-unsaved="' + kind + '"]');
      if (!el) return;
      el.dataset.persistBusy = "1";
      el.innerHTML = " " + html;
      setTimeout(function () { delete el.dataset.persistBusy; QBR.persistRefreshBadge(); }, ms || 3500);
    } catch (e) {}
  };

  // One-line alert above the dashboards when edits need attention.
  function alertState() {
    var unsaved = QBR.journalUnsavedCount();
    if (HEALTH.local === "failed" && (unsaved || countOps(MEM))) {
      return { level: "danger", text: (HEALTH.idb === "ok"
        ? "Browser storage is full — " + countOps(MEM) + " unsaved edit(s) are kept only in this tab and its backup. Save to Excel or export them now."
        : "Browser storage is unavailable — " + countOps(MEM) + " unsaved edit(s) exist only in this tab. Save to Excel or export them before closing.") };
    }
    var orph = QBR.journalOrphans().filter(function (o) { return o.target; });
    if (orph.length) {
      var n = orph.reduce(function (s, o) { return s + o.count; }, 0);
      return { level: "warning", text: n + " unsaved edit(s) were made on an older copy of " + orph.map(function (o) { return '"' + esc(o.fileName || "a file") + '"'; }).join(", ") + " and are not applied." };
    }
    var foreign = QBR.journalForeign().filter(function (f) { return f.entries.some(function (e) { return e.loaded || e.sameName; }); });
    if (foreign.length) return { level: "warning", text: "Unsaved edits for a loaded file were found in " + esc(foreign[0].label) + "." };
    if (MEM.parked.length) return { level: "warning", text: MEM.parked.length + " edit(s) couldn't be placed during a merge — re-enter them, then dismiss." };
    if (HEALTH.recovered && !HEALTH.dismissed) return { level: "info", text: "Recovered " + HEALTH.recovered + " unsaved edit(s) from the browser backup.", dismiss: true };
    return null;
  }
  function refreshUi() {
    QBR.persistRefreshBadge();
    try {
      var host = document.getElementById("app-main") || document.body;
      var bar = document.getElementById("qbr-journal-alert");
      var st = alertState();
      if (!st) { if (bar) bar.remove(); return; }
      if (!bar) {
        bar = document.createElement("div"); bar.id = "qbr-journal-alert"; bar.setAttribute("role", "status");
        host.insertBefore(bar, host.firstChild);
      }
      bar.className = "qbr-journal-alert jl-" + st.level;
      bar.innerHTML = '<span class="jl-text">' + st.text + '</span> <button type="button" class="btn btn-sm btn-light" data-jl="review">Review edits</button>' +
        (st.dismiss ? ' <button type="button" class="btn btn-sm btn-link" data-jl="dismiss">Dismiss</button>' : "");
    } catch (e) {}
  }
  QBR.journalRefreshUi = refreshUi;

  function fmtDate(ts) { try { return new Date(ts).toLocaleString(); } catch (e) { return ""; } }
  function describe(o) {
    var a = (o.args || []).slice(o.op in FP_OPS ? 1 : 0), s;
    try { s = JSON.stringify(a); } catch (e) { s = ""; }
    if (s.length > 110) s = s.slice(0, 107) + "…";
    return o.op + (o.sig ? " · " + o.sig : "") + " " + s;
  }
  // The "Unsaved edits" panel: storage status, pending, orphans, other folders, parked, backups.
  QBR.journalPanel = function () {
    var old = document.getElementById("qbr-journal-panel"); if (old) old.remove();
    var ov = document.createElement("div");
    ov.id = "qbr-journal-panel"; ov.className = "merge-ov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "jp-h");
    function render(msg) {
      var h = QBR.journalHealth(), fps = loadedFps();
      var store = h.local === "ok" ? "Saved in this browser" + (h.idb === "ok" ? " + backup" : "") : (h.idb === "ok" ? "<b>Browser storage full</b> — backup only" : "<b>Not stored</b> — this tab only");
      var pend = fps.map(function (fp) { var e = MEM.files[fp]; return e && e.ops.length ? "<li><b>" + esc(e.fileName || fileNameOf(fp)) + "</b> — " + e.ops.length + " edit(s)</li>" : ""; }).join("");
      var orph = QBR.journalOrphans().map(function (o) {
        var e = MEM.files[o.fp];
        var list = e.ops.slice(0, 8).map(function (x) { return "<li>" + esc(describe(x)) + (x.lastTry ? ' <span class="text-muted">— ' + esc(x.lastTry) + "</span>" : "") + "</li>"; }).join("") + (e.ops.length > 8 ? "<li>… " + (e.ops.length - 8) + " more</li>" : "");
        return '<div class="jp-item"><div><b>' + esc(o.fileName || "unknown file") + "</b> — " + o.count + " edit(s), " + esc(fmtDate(o.first)) + " – " + esc(fmtDate(o.last)) + "</div>" +
          '<details><summary class="small">Show edits</summary><ul class="small">' + list + "</ul></details>" +
          '<div class="jp-btns">' + (o.target ? '<button type="button" class="btn btn-sm btn-primary" data-jp="apply" data-fp="' + esc(o.fp) + '">Apply to loaded "' + esc(o.target) + '"</button> ' : '<span class="small text-muted">Load the file to apply these. </span>') +
          '<button type="button" class="btn btn-sm btn-outline-secondary" data-jp="export1" data-fp="' + esc(o.fp) + '">Export</button> ' +
          '<button type="button" class="btn btn-sm btn-outline-danger" data-jp="discard" data-fp="' + esc(o.fp) + '">Discard</button></div></div>';
      }).join("");
      var foreign = QBR.journalForeign().map(function (f) {
        return '<div class="jp-item"><div><b>' + esc(f.label) + "</b></div><ul class=\"small\">" + f.entries.map(function (e) { return "<li>" + esc(e.fileName || "unknown file") + " — " + e.count + " edit(s)" + (e.loaded ? " · <b>matches a loaded file</b>" : "") + "</li>"; }).join("") +
          (f.parked ? "<li>" + f.parked + " unplaced merge edit(s)</li>" : "") + '</ul><div class="jp-btns"><button type="button" class="btn btn-sm btn-outline-primary" data-jp="move" data-key="' + esc(f.key) + '">Move to this dashboard</button></div></div>';
      }).join("");
      var parked = MEM.parked.map(function (p) {
        return "<tr><td>" + esc(p.fileName) + "</td><td><b>" + esc(p.sheet) + "</b> " + esc(p.ref) + (p.key ? '<div class="small text-muted">' + esc(p.header) + " · " + esc(p.key) + "</div>" : "") + "</td><td>" + esc(p.value) + "</td><td class=\"small\">" + esc(p.why) + "</td></tr>";
      }).join("");
      ov.innerHTML = '<div class="merge-box jp-box">' +
        '<h5 id="jp-h">Unsaved edits</h5>' +
        (msg ? '<div class="jp-msg small">' + msg + "</div>" : "") +
        '<p class="small mb-1">Storage: ' + store + (h.persisted === true ? " · protected from browser clean-up" : "") + ' · <span class="text-muted">' + esc(h.ns) + "</span></p>" +
        '<div class="merge-scroll">' +
        "<h6>Waiting to be saved to Excel</h6>" + (pend ? "<ul class=\"small\">" + pend + "</ul>" : '<p class="small text-muted">None for the loaded files.</p>') +
        (orph ? '<h6>Made on an older copy of a file</h6><p class="small">The file changed since these edits were made (for example, re-uploaded after editing in Excel). Apply re-places them by school/serial; edits that can\'t be matched stay here.</p>' + orph : "") +
        (foreign ? '<h6>In another dashboard folder</h6><p class="small">Every dashboard folder keeps its own edits. Move them here only if this is the copy you save from.</p>' + foreign : "") +
        (parked ? '<h6>Couldn\'t be placed during a merge</h6><p class="small">Excel deleted or renamed the row/column. Re-enter these, then dismiss.</p><table class="table table-sm small"><thead><tr><th>File</th><th>Cell</th><th>Your value</th><th>Why</th></tr></thead><tbody>' + parked + '</tbody></table><button type="button" class="btn btn-sm btn-outline-secondary" data-jp="unpark">Dismiss all</button>' : "") +
        '<h6 class="mt-2">Backups</h6><p class="small">Batches cleared after a save are kept 14 days in this browser. <button type="button" class="btn btn-sm btn-link p-0" data-jp="trash">Export saved batches</button></p>' +
        "</div>" +
        '<p class="small text-muted mb-1">Exports contain school names and contacts — store them like the workbook.</p>' +
        '<div class="merge-act"><button type="button" class="btn btn-outline-secondary" data-jp="export">Export all (JSON)</button>' +
        '<label class="btn btn-outline-secondary mb-0">Import…<input type="file" accept=".json,application/json" data-jp="import" hidden></label>' +
        '<button type="button" class="btn btn-primary" data-jp="close">Close</button></div></div>';
    }
    render();
    document.body.appendChild(ov);
    var close = function () { ov.remove(); refreshUi(); };
    ov.addEventListener("keydown", function (ev) { if (ev.key === "Escape") close(); });
    ov.addEventListener("change", function (ev) {
      var t = ev.target; if (!t || t.getAttribute("data-jp") !== "import" || !t.files || !t.files[0]) return;
      t.files[0].text().then(function (txt) {
        var r = QBR.journalImport(txt);
        render(r.ok ? "Imported " + r.added + " edit(s)" + (r.parked ? " and " + r.parked + " unplaced edit(s)" : "") + (r.skipped ? " (" + r.skipped + " already here)" : "") + "." : "Import failed: " + esc(r.error));
      });
    });
    ov.addEventListener("click", function (ev) {
      var b = ev.target.closest && ev.target.closest("[data-jp]"); if (!b || b.tagName === "INPUT") return;
      var a = b.getAttribute("data-jp"), fp = b.getAttribute("data-fp");
      if (a === "close") return close();
      if (a === "export") { QBR.journalExport(); return; }
      if (a === "export1") { QBR.journalExport({ fps: [fp] }); return; }
      if (a === "trash") { QBR.journalExportTrash(); return; }
      if (a === "apply") { var r = QBR.journalApplyOrphan(fp); render("Applied " + r.applied + " edit(s)" + (r.kept ? "; " + r.kept + " kept: " + esc(r.notes.join("; ")) : "") + "."); return; }
      if (a === "discard") { if (confirm("Discard these edits? They stay in Backups for 14 days.")) { QBR.journalDiscard(fp); render("Discarded."); } return; }
      if (a === "unpark") { QBR.journalUnpark(); render("Dismissed."); return; }
      if (a === "move") { var m = QBR.journalMoveForeign(b.getAttribute("data-key")); Promise.resolve(m.done).then(function () { render("Moved " + m.moved + " edit(s) to this dashboard."); }); return; }
    });
    setTimeout(function () { var c = ov.querySelector('[data-jp="close"]'); if (c) c.focus(); }, 0);
    return ov;
  };

  document.addEventListener("click", function (ev) {
    var t = ev.target;
    if (!t || !t.closest) return;
    if (t.closest(".qbr-unsaved-badge") || t.closest('#qbr-journal-alert [data-jl="review"]')) { ev.preventDefault(); QBR.journalPanel(); return; }
    if (t.closest('#qbr-journal-alert [data-jl="dismiss"]')) { HEALTH.dismissed = true; refreshUi(); }
  });

  // Closing the tab while edits exist only in memory → ask first.
  window.addEventListener("beforeunload", function (ev) {
    if (HEALTH.local === "failed" && countOps(MEM) > 0) { ev.preventDefault(); ev.returnValue = ""; return ""; }
  });
  QBR._journalAtRisk = function () { return HEALTH.local === "failed" && countOps(MEM) > 0; };
  QBR._journalReset = function () { MEM = blank(); HEALTH.recovered = 0; persist(); };   // tests only

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refreshUi);
  else refreshUi();
})();
