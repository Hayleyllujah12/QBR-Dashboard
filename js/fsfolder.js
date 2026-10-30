/* ==========================================================================
 * fsfolder.js — v1.33.0 "Open folder": one permission for every workbook.
 *
 * Problem: a plain upload (<input type=file> / drag-drop) gives the page a COPY
 * of each file — never its location — so enabling Direct save meant re-picking
 * every workbook and approving "Allow editing" once per file, every session.
 *
 * Fix: the user picks the PARENT folder once (showDirectoryPicker, readwrite).
 * One prompt covers every file inside it, subfolders included. We scan for
 * Excel workbooks, the user ticks which to load, and each one is loaded AND
 * linked (Direct save + ↻ reload without a picker) in one go.
 *
 * Next session: the folder handle is kept in IndexedDB. The first save/reload
 * (or the "Reconnect folder" button) asks once for the folder, which re-grants
 * every linked file in it. On https (Netlify beta) Chrome/Edge offer
 * "Allow on every visit", which removes even that prompt.
 *
 * Chrome/Edge only (File System Access API). Elsewhere the buttons stay hidden
 * and the classic upload + Export download path is unchanged.
 * Depends on persist.js globals: persistIdbGet/Put, fsLinksSave, fsMemFile,
 * fsHashBytes, fsSetBase; app.js globals: loadItems, fileKey, cacheSession,
 * renderFileList, renderAll.
 * ========================================================================== */
(function () {
  "use strict";
  var QBR = window.QBR = window.QBR || {};
  var DIRS_KEY = "fs-dirs";
  var MAX_DEPTH = 3, MAX_ENTRIES = 1500, MAX_BOOKS = 80;
  QBR._fsDirs = [];          // [{id, name, handle, picked:[relPath]}]
  QBR._fsDirState = {};      // id -> "granted" | "prompt" | "denied"

  QBR.fsFolderSupported = function () {
    return typeof window.showDirectoryPicker === "function";
  };

  function e(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function nk(n) { try { return String(n || "").trim().toLowerCase(); } catch (x) { return ""; } }
  // Small status toast (persistNote targets per-kind badges; this is app-wide).
  var toastT = null;
  function note(msg, ms) {
    try {
      var el = document.getElementById("fo-toast");
      if (!msg) { if (el) el.hidden = true; return; }
      if (!el) {
        el = document.createElement("div"); el.id = "fo-toast"; el.className = "fo-toast";
        el.setAttribute("role", "status"); el.setAttribute("aria-live", "polite");
        document.body.appendChild(el);
      }
      el.textContent = msg; el.hidden = false;
      clearTimeout(toastT);
      toastT = setTimeout(function () { el.hidden = true; }, ms || 5000);
    } catch (x) {}
  }
  QBR._foNote = note;
  function newId() { return "d" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  /* ---------------- storage ---------------- */
  function dirsSave() {
    var slim = QBR._fsDirs.map(function (d) { return { id: d.id, name: d.name, handle: d.handle, picked: d.picked || [], keep: !!d.keep }; });
    return persistIdbPut(DIRS_KEY, slim).catch(function (x) { console.warn("[QBR] folder store failed:", x && x.message); });
  }
  function dirsLoad() {
    return persistIdbGet(DIRS_KEY).then(function (v) {
      QBR._fsDirs = Array.isArray(v) ? v : [];
      return QBR.fsFolderRefreshState();
    }).catch(function () { QBR._fsDirs = []; });
  }
  function dirById(id) { return (QBR._fsDirs || []).filter(function (d) { return d.id === id; })[0] || null; }

  // Query (never request) each folder's permission so the list can show
  // "Reconnect folder" only when it is actually needed.
  QBR.fsFolderRefreshState = async function () {
    for (var i = 0; i < (QBR._fsDirs || []).length; i++) {
      var d = QBR._fsDirs[i];
      try { QBR._fsDirState[d.id] = await d.handle.queryPermission({ mode: "readwrite" }); }
      catch (x) { QBR._fsDirState[d.id] = "prompt"; }
    }
    if (typeof renderFileList === "function") { try { renderFileList(); } catch (x) {} }
  };
  // Folders that have linked files but no live permission yet.
  QBR.fsFoldersNeedingGrant = function () {
    var used = {};
    (QBR._fsLinks || []).forEach(function (l) { if (l.dirId) used[l.dirId] = 1; });
    return (QBR._fsDirs || []).filter(function (d) { return used[d.id] && QBR._fsDirState[d.id] !== "granted"; });
  };

  /* ---------------- scanning ---------------- */
  // 2026-10-25 load bar: onEntry(seen) reports walk progress; the walk yields
  // every 50 entries so the bar paints (getFile awaits already yield per file).
  async function scan(dir, onEntry) {
    var out = [], seen = 0;
    async function walk(h, path, depth) {
      for await (var entry of h.values()) {
        var s = ++seen;
        if (onEntry) { try { onEntry(s); } catch (x) {} }
        if (s % 50 === 0) await new Promise(function (r) { setTimeout(r, 0); });
        if (s > MAX_ENTRIES || out.length >= MAX_BOOKS) return;
        var n = entry.name || "";
        if (n.charAt(0) === "." || n.indexOf("~$") === 0) continue;   // hidden + Excel lock files
        if (entry.kind === "directory") {
          if (depth < MAX_DEPTH) await walk(entry, path.concat(n), depth + 1);
        } else if (/\.(xlsx|xlsm)$/i.test(n)) {
          var f = null; try { f = await entry.getFile(); } catch (x) {}
          out.push({ name: n, path: path, rel: path.concat(n), handle: entry, size: f ? f.size : 0, lastModified: f ? f.lastModified : 0 });
        }
      }
    }
    await walk(dir, [], 0);
    out.sort(function (a, b) { return a.rel.join("/").localeCompare(b.rel.join("/")); });
    return { books: out, truncated: seen > MAX_ENTRIES || out.length >= MAX_BOOKS };
  }
  async function resolveRel(dir, rel) {
    var h = dir;
    for (var i = 0; i < rel.length - 1; i++) h = await h.getDirectoryHandle(rel[i]);
    return h.getFileHandle(rel[rel.length - 1]);
  }

  /* ---------------- picking dialog ---------------- */
  // Pre-tick: workbooks already loaded / linked / picked last time; when none
  // are known, everything (if the folder is small). One file per name.
  function defaultPicks(books, dirRec) {
    var known = {};
    ((typeof APP !== "undefined" && APP.files) || []).forEach(function (it) { known[nk(it.name)] = 1; });
    (QBR._fsLinks || []).forEach(function (l) { known[nk(l.name)] = 1; });
    var lastRel = {};
    ((dirRec && dirRec.picked) || []).forEach(function (r) { lastRel[r] = 1; });
    var anyKnown = books.some(function (b) { return known[nk(b.name)] || lastRel[b.rel.join("/")]; });
    var pick = {}, bestByName = {};
    books.forEach(function (b, i) {
      var want = anyKnown ? (lastRel[b.rel.join("/")] || known[nk(b.name)]) : books.length <= 12;
      if (!want) return;
      var k = nk(b.name), prev = bestByName[k];
      // same name in two subfolders: the one picked last time, else the newest
      if (prev == null || (!lastRel[books[prev].rel.join("/")] && (lastRel[b.rel.join("/")] || b.lastModified > books[prev].lastModified))) bestByName[k] = i;
    });
    Object.keys(bestByName).forEach(function (k) { pick[bestByName[k]] = true; });
    return pick;
  }

  function pickDialog(dirName, books, truncated, pre) {
    if (typeof QBR._folderAutoPick === "function") return Promise.resolve(QBR._folderAutoPick(books, pre));
    return new Promise(function (resolve) {
      var dupName = {};
      books.forEach(function (b) { var k = nk(b.name); dupName[k] = (dupName[k] || 0) + 1; });
      var rows = books.map(function (b, i) {
        var d = b.lastModified ? new Date(b.lastModified).toLocaleDateString() : "";
        return '<tr><td><input type="checkbox" class="fo-pick" data-i="' + i + '"' + (pre[i] ? " checked" : "") + ' aria-label="Load ' + e(b.name) + '"></td>' +
          '<td>' + e(b.name) + (dupName[nk(b.name)] > 1 ? ' <span class="badge bg-warning text-dark" title="Same file name in another subfolder — only one can be loaded">same name</span>' : "") + '</td>' +
          '<td class="small text-muted">' + (b.path.length ? e(b.path.join(" / ")) : "(top folder)") + '</td>' +
          '<td class="small text-muted">' + e(d) + '</td></tr>';
      }).join("");
      var ov = document.createElement("div");
      ov.className = "merge-ov fo-ov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "fo-h");
      ov.innerHTML = '<div class="merge-box">' +
        '<h5 id="fo-h">Open folder: ' + e(dirName) + '</h5>' +
        '<p class="small mb-1">Found <b>' + books.length + '</b> Excel workbook' + (books.length === 1 ? "" : "s") +
        (truncated ? " (stopped scanning early — pick a smaller folder if one is missing)" : "") +
        '. Tick the ones to load. Direct save is enabled for all of them with the permission you just gave — no prompt per file.</p>' +
        '<div class="merge-tools"><button type="button" class="btn btn-sm btn-outline-secondary" data-fo="all">Tick all</button> ' +
        '<button type="button" class="btn btn-sm btn-outline-secondary" data-fo="none">Clear</button></div>' +
        '<div class="merge-scroll"><table class="table table-sm fo-table"><thead><tr><th></th><th>Workbook</th><th>Subfolder</th><th>Modified</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
        '<div class="merge-act"><button type="button" class="btn btn-primary" id="fo-ok">Load &amp; link</button> ' +
        '<button type="button" class="btn btn-outline-secondary" id="fo-cancel">Cancel</button></div></div>';
      document.body.appendChild(ov);
      var boxes = function () { return [].slice.call(ov.querySelectorAll(".fo-pick")); };
      var sync = function () {
        var n = boxes().filter(function (c) { return c.checked; }).length;
        var ok = ov.querySelector("#fo-ok");
        ok.textContent = "Load & link (" + n + ")"; ok.disabled = !n;
      };
      // one file per name: ticking one unticks its same-name twins
      ov.addEventListener("change", function (ev) {
        var t = ev.target; if (!t.classList || !t.classList.contains("fo-pick") || !t.checked) { sync(); return; }
        var k = nk(books[+t.dataset.i].name);
        boxes().forEach(function (c) { if (c !== t && nk(books[+c.dataset.i].name) === k) c.checked = false; });
        sync();
      });
      ov.addEventListener("click", function (ev) {
        var a = ev.target.closest && ev.target.closest("[data-fo]");
        if (a) {
          var on = a.dataset.fo === "all", taken = {};
          boxes().forEach(function (c) {
            var k = nk(books[+c.dataset.i].name);
            c.checked = on && !taken[k]; if (c.checked) taken[k] = 1;
          });
          sync();
        }
      });
      var done = function (v) { ov.remove(); document.removeEventListener("keydown", esc); resolve(v); };
      var esc = function (ev) { if (ev.key === "Escape") done(null); };
      document.addEventListener("keydown", esc);
      ov.querySelector("#fo-cancel").onclick = function () { done(null); };
      ov.querySelector("#fo-ok").onclick = function () {
        done(boxes().filter(function (c) { return c.checked; }).map(function (c) { return +c.dataset.i; }));
      };
      sync();
      setTimeout(function () { try { ov.querySelector("#fo-ok").focus(); } catch (x) {} }, 0);
    });
  }

  /* ---------------- open folder: load + link ---------------- */
  QBR.fsOpenFolder = async function () {
    if (!QBR.fsFolderSupported()) {
      alert("Opening a folder needs Chrome or Edge. Use Upload instead.");
      return false;
    }
    var last = (QBR._fsDirs || [])[0];
    var dir;
    nameScan = {};
    try {
      var opt = { id: "qbr-workbooks", mode: "readwrite" };
      if (last && last.handle) opt.startIn = last.handle;
      dir = await window.showDirectoryPicker(opt);
    } catch (x) {
      if (x && x.name !== "AbortError") alert("Could not open that folder: " + (x.message || x.name) +
        "\n\nThe browser refuses some system folders (for example your whole user folder). Pick the folder that holds the workbooks.");
      return false;
    }
    // mode:"readwrite" asks at pick time; double-check (some builds grant read first)
    try {
      var p = await dir.queryPermission({ mode: "readwrite" });
      if (p !== "granted") p = await dir.requestPermission({ mode: "readwrite" });
      if (p !== "granted") { alert("Edit permission was not granted — nothing was loaded."); return false; }
    } catch (x) { return false; }

    // reuse the record when it is the same folder
    var rec = null;
    for (var i = 0; i < (QBR._fsDirs || []).length; i++) {
      try { if (await QBR._fsDirs[i].handle.isSameEntry(dir)) { rec = QBR._fsDirs[i]; break; } } catch (x) {}
    }
    var LP = (typeof QBR !== "undefined" && QBR.loadProgress) || null;
    var ownScan = LP ? LP.begin() : false;
    var res;
    try { if (ownScan) LP.scan(0); res = await scan(dir, ownScan ? function (n) { LP.scan(n); } : null); }
    catch (x) { if (LP) LP.end(); alert("Could not read the folder: " + (x && x.message || x)); return false; }
    if (LP) LP.end();
    if (!res.books.length) { alert('No .xlsx / .xlsm workbooks found in "' + dir.name + '" (or its subfolders).'); return false; }
    var chosen = await pickDialog(dir.name, res.books, res.truncated, defaultPicks(res.books, rec));
    if (!chosen || !chosen.length) return false;
    var picks = chosen.map(function (i) { return res.books[i]; });

    if (!rec) { rec = { id: newId(), name: dir.name, handle: dir, picked: [] }; QBR._fsDirs.unshift(rec); }
    else { rec.handle = dir; rec.name = dir.name; }
    rec.picked = picks.map(function (b) { return b.rel.join("/"); });
    // Any number of folders can be linked. Folders still used by a link are never dropped;
    // only UNUSED remembered folders are trimmed (keep at most 5 records in total).
    var usedIds = {};
    (QBR._fsLinks || []).forEach(function (l) { if (l.dirId) usedIds[l.dirId] = 1; });
    (QBR._fsDirs || []).forEach(function (d) { if (d.keep) usedIds[d.id] = 1; }); // picked for another tab (Trainers)
    var others = QBR._fsDirs.filter(function (d) { return d !== rec; });
    var used = others.filter(function (d) { return usedIds[d.id]; });
    var spare = others.filter(function (d) { return !usedIds[d.id]; }).slice(0, Math.max(0, 4 - used.length));
    QBR._fsDirs = [rec].concat(others.filter(function (d) { return usedIds[d.id] || spare.indexOf(d) >= 0; }));
    QBR._fsDirState[rec.id] = "granted";

    try {
      var mems = [];
      // 2026-10-26 per-file skip: a workbook that vanished between the picker
      // list and the read (OneDrive sync/move - Chromium NotFoundError) must
      // not abort the whole batch. Collect it, load the rest, name it below.
      var failed = [];
      var ownRead = LP ? LP.begin() : false;
      try {
        for (var j = 0; j < picks.length; j++) {
          if (ownRead) LP.read(j + 1, picks.length, picks[j].name);
          try {
            var f = await picks[j].handle.getFile();
            var mem = await fsMemFile(f, f.name);
            mems.push({ b: picks[j], file: f, mem: mem });
          } catch (x) {
            failed.push({ name: picks[j].name, rel: (picks[j].rel || []).join("/"), err: (x && x.message) || (x && x.name) || String(x) });
          }
        }
      } finally { if (LP) LP.end(); }
      if (!mems.length) {
        alert("Could not load from the folder: " + failed.map(function (d) { return d.rel || d.name; }).join(", ") + (failed[0] ? " - " + failed[0].err : ""));
        return false;
      }
      // accumulate exactly like Upload: dedup by name, newest wins
      var map = new Map();
      ((typeof APP !== "undefined" && APP.files) || []).forEach(function (it) { map.set(fileKey(it.name), it); });
      mems.forEach(function (m) { map.set(fileKey(m.b.name), { name: m.b.name, blob: m.mem.file }); });
      var ok = await loadItems(Array.from(map.values()), {});
      if (!ok) { note("Could not load those workbooks" + (failed.length ? " - skipped: " + failed.map(function (d) { return d.rel || d.name; }).join(", ") : ""), 6000); return false; }
      var names = {};
      mems.forEach(function (m) { names[nk(m.b.name)] = 1; });
      var links = (QBR._fsLinks || []).filter(function (l) { return !names[nk(l.name)]; });
      mems.forEach(function (m) {
        var fp = QBR.fpOf(m.b.name, m.mem.file);
        var kinds = Array.from((QBR._kindByFp || {})[fp] || []);
        // one link per editable kind: the newest link for a kind wins (as fsLinkFile)
        if (kinds.length) links = links.filter(function (l) { return !(l.kinds || []).some(function (k) { return kinds.indexOf(k) >= 0; }); });
        links.push({ fp: fp, kinds: kinds, name: m.b.name, size: m.file.size, lastModified: m.file.lastModified,
          hash: fsHashBytes(m.mem.bytes), handle: m.b.handle, dirId: rec.id, rel: m.b.rel });
      });
      QBR._fsLinks = links;
      await fsLinksSave();
      await dirsSave();
      if (typeof cacheSession === "function") cacheSession();
      if (typeof QBR.fsRefreshStatus === "function") QBR.fsRefreshStatus();
      if (typeof renderAll === "function") { try { renderAll(); } catch (x) {} }
      if (typeof renderFileList === "function") { try { renderFileList(); } catch (x) {} }
      try { QBR.persistRefreshBadge(); } catch (x) {}
      linked(rec.id);
      var loadedMsg = "Loaded & linked ✓ " + mems.length + " workbook" + (mems.length === 1 ? "" : "s") + " from " + dir.name;
      if (failed.length) loadedMsg += " - failed: " + failed.map(function (d) { return d.rel || d.name; }).join(", ");
      note(loadedMsg, failed.length ? 8000 : 5000);
      return true;
    } catch (x) {
      alert("Could not load from the folder: " + (x && x.message || x));
      return false;
    }
  };

  /* ---------------- other tabs: read files in a linked folder (v1.39.0) ----------------
   * The Trainers tab finds its tracker + eval exports in the same folder the dashboard
   * workbooks are linked from, without loading them as dashboard workbooks. */
  function linked(id) {
    try { document.dispatchEvent(new CustomEvent("qbr:folder-linked", { detail: { id: id } })); } catch (x) {}
  }
  QBR.fsFolderList = function () {
    return (QBR._fsDirs || []).map(function (d) { return { id: d.id, name: d.name, state: QBR._fsDirState[d.id] || "prompt" }; });
  };
  // Excel files in a remembered folder: [{name, path, rel, handle, size, lastModified}], or null
  // when the folder isn't allowed. ask:true may prompt (must run inside a click).
  QBR.fsFolderBooks = async function (id, opts) {
    var d = dirById(id); if (!d) return null;
    var p;
    try {
      p = await d.handle.queryPermission({ mode: "readwrite" });
      if (p !== "granted" && opts && opts.ask) p = await d.handle.requestPermission({ mode: "readwrite" });
    } catch (x) { p = "denied"; }
    QBR._fsDirState[id] = p;
    if (p !== "granted") return null;
    try { return (await scan(d.handle)).books; } catch (x) { return null; }
  };
  /* v1.41.0: configurable scan for other file types / deeper trees (MOS certificate PDFs live 4–5 folders deep).
   * opts: { ext: RegExp (default .xlsx/.xlsm), maxDepth (default 3), maxEntries (1500), maxFiles (80),
   *         want(relParts) -> bool (filter files), prefer(name) -> bool (folders scanned first),
   *         skipDir(name) -> bool (folders not entered) }
   * -> { files: [{name, path, rel, handle, size, lastModified}], truncated, scanned, folders, depthCapped } */
  var SKIP_DIRS = /^(images?|photos?|pictures?|videos?|screenshots?|recordings?|node_modules|\$recycle\.bin|system volume information)$/i;
  async function scanFiles(dir, opts) {
    opts = opts || {};
    var re = opts.ext || /\.(xlsx|xlsm)$/i, maxDepth = opts.maxDepth == null ? MAX_DEPTH : opts.maxDepth;
    var maxEntries = opts.maxEntries || MAX_ENTRIES, maxFiles = opts.maxFiles || MAX_BOOKS;
    var out = [], seen = 0, folders = 0, capped = 0, stop = false;
    async function walk(h, path, depth) {
      var dirs = [];
      for await (var entry of h.values()) {
        if (++seen > maxEntries) { stop = true; break; }
        var n = entry.name || "";
        if (n.charAt(0) === "." || n.indexOf("~$") === 0) continue;
        if (entry.kind === "directory") { if (!SKIP_DIRS.test(n) && !(opts.skipDir && opts.skipDir(n))) dirs.push(entry); continue; }
        if (!re.test(n) || out.length >= maxFiles) { if (out.length >= maxFiles && re.test(n)) stop = true; continue; }
        var rel = path.concat(n);
        if (opts.want && !opts.want(rel)) continue;
        var f = null; try { f = await entry.getFile(); } catch (x) {}
        out.push({ name: n, path: path, rel: rel, handle: entry, size: f ? f.size : 0, lastModified: f ? f.lastModified : 0 });
      }
      if (stop) return;
      if (depth >= maxDepth) { capped += dirs.length; return; }
      if (opts.prefer) dirs.sort(function (a, b) { return (opts.prefer(b.name) ? 1 : 0) - (opts.prefer(a.name) ? 1 : 0); });
      for (var i = 0; i < dirs.length && !stop; i++) { folders++; await walk(dirs[i], path.concat(dirs[i].name), depth + 1); }
    }
    await walk(dir, [], 0);
    out.sort(function (a, b) { return a.rel.join("/").localeCompare(b.rel.join("/")); });
    return { files: out, truncated: stop, scanned: seen, folders: folders, depthCapped: capped };
  }
  QBR._fsScanFiles = scanFiles;
  async function grant(id, opts) {
    var d = dirById(id); if (!d) return null;
    var p;
    try {
      p = await d.handle.queryPermission({ mode: "readwrite" });
      if (p !== "granted" && opts && opts.ask) p = await d.handle.requestPermission({ mode: "readwrite" });
    } catch (x) { p = "denied"; }
    QBR._fsDirState[id] = p;
    return p === "granted" ? d : null;
  }
  QBR.fsFolderFiles = async function (id, opts) {
    var d = await grant(id, opts); if (!d) return null;
    try { return await scanFiles(d.handle, opts); } catch (x) { return null; }
  };
  // The folder's own handle (create / write files at its top level, e.g. a log workbook). -> handle | null
  QBR.fsFolderDir = async function (id, opts) { var d = await grant(id, opts); return d ? d.handle : null; };

  // Pick a folder and remember it WITHOUT loading anything into the dashboard. -> id | null
  QBR.fsFolderPickOnly = async function () {
    if (!QBR.fsFolderSupported()) return null;
    var dir, last = (QBR._fsDirs || [])[0];
    try {
      var opt = { id: "qbr-workbooks", mode: "readwrite" };
      if (last && last.handle) opt.startIn = last.handle;
      dir = await window.showDirectoryPicker(opt);
    } catch (x) { return null; }
    var rec = null;
    for (var i = 0; i < (QBR._fsDirs || []).length; i++) {
      try { if (await QBR._fsDirs[i].handle.isSameEntry(dir)) { rec = QBR._fsDirs[i]; break; } } catch (x) {}
    }
    if (rec) { rec.handle = dir; QBR._fsDirs = [rec].concat(QBR._fsDirs.filter(function (d) { return d !== rec; })); }
    else { rec = { id: newId(), name: dir.name, handle: dir, picked: [] }; QBR._fsDirs.unshift(rec); }
    rec.keep = true;
    try { QBR._fsDirState[rec.id] = await dir.queryPermission({ mode: "readwrite" }); } catch (x) { QBR._fsDirState[rec.id] = "prompt"; }
    await dirsSave();
    return rec.id;
  };

  /* ---------------- reconnect (next session) ---------------- */
  // One prompt for the folder; then re-resolve each linked file through it so
  // the file handles inherit the grant. Must run inside a click.
  QBR.fsFolderReconnect = async function (id) {
    var d = dirById(id); if (!d) return false;
    nameScan = {};
    var p;
    try {
      p = await d.handle.queryPermission({ mode: "readwrite" });
      if (p !== "granted") p = await d.handle.requestPermission({ mode: "readwrite" });
    } catch (x) { p = "denied"; }
    QBR._fsDirState[id] = p;
    if (p !== "granted") { if (typeof renderFileList === "function") renderFileList(); return false; }
    var missing = [];
    for (var i = 0; i < (QBR._fsLinks || []).length; i++) {
      var l = QBR._fsLinks[i];
      if (l.dirId !== id || !l.rel) continue;
      try { l.handle = await resolveRel(d.handle, l.rel); }
      catch (x) { missing.push(l.rel.join("/")); }
    }
    await fsLinksSave();
    if (typeof renderFileList === "function") { try { renderFileList(); } catch (x) {} }
    linked(id);
    if (missing.length) note("Reconnected " + d.name + " — not found (moved or renamed?): " + missing.join(", "), 12000);
    else note("Reconnected ✓ " + d.name + " — Direct save is on for every linked workbook in it", 5000);
    return true;
  };
  QBR.fsFolderReconnectAll = async function () {
    var need = QBR.fsFoldersNeedingGrant(), ok = true;
    for (var i = 0; i < need.length; i++) ok = (await QBR.fsFolderReconnect(need[i].id)) && ok;
    return ok;
  };

  // Called by persist.js fsEnsurePermission when a file handle is not granted:
  // if it belongs to a linked folder, re-grant the folder (covers every file).
  QBR.fsFolderGrantFor = async function (handle) {
    var link = (QBR._fsLinks || []).filter(function (l) { return l.handle === handle && l.dirId; })[0];
    if (!link) return null;                      // not a folder link → caller asks for the file as before
    if (!dirById(link.dirId)) return null;       // folder record gone → fall back to asking for the file
    if (!(await QBR.fsFolderReconnect(link.dirId))) return false; // folder declined → don't nag per file
    return link.handle; // re-resolved through the folder
  };

  // Reload every folder-linked workbook from disk (picks up Excel edits).
  // Files with unsaved dashboard edits are skipped — save those first.
  // 2026-10-25 load bar: the outer session owns the display (Reloading i/N);
  // inner per-file loads nest silently. Failures used to vanish in the catch
  // below — they are now collected and named in the end note. Skips (unsaved
  // edits) keep their existing wording and are NOT failures.
  QBR.fsFolderReloadAll = async function () {
    nameScan = {};
    if (!(await QBR.fsFolderReconnectAll())) return false;
    var list = (QBR._fsLinks || []).filter(function (l) { return l.dirId; });
    var LP = (typeof QBR !== "undefined" && QBR.loadProgress) || null;
    var own = (LP && list.length) ? LP.begin() : false;
    var done = 0, skipped = [], failed = [];
    try {
      for (var i = 0; i < list.length; i++) {
        var l = list[i], pending = 0;
        if (own) LP.reload(i + 1, list.length, l.name);
        try { pending = QBR.journalOpsFor(l.fp).length; } catch (x) {}
        if (pending) { skipped.push(l.name); continue; }
        var r = false;
        try { r = await QBR.fsReloadFromFile(l.name, { quiet: true }); } catch (x) { r = false; }
        if (r) done++; else failed.push(l.name);
      }
    } finally { if (LP) LP.end(); }
    note("Reloaded ✓ " + done + " workbook" + (done === 1 ? "" : "s") +
      (skipped.length ? " · skipped (unsaved edits — save first): " + skipped.join(", ") : "") +
      (failed.length ? " · failed: " + failed.join(", ") : ""), (skipped.length || failed.length) ? 12000 : 5000);
    return true;
  };

  // A link is gone → drop folders nobody uses any more (keeps the store tidy).
  QBR.fsFolderPrune = function () {
    var used = {};
    (QBR._fsLinks || []).forEach(function (l) { if (l.dirId) used[l.dirId] = 1; });
    var before = QBR._fsDirs.length;
    // keep the most recent folder even when unused (picker starts there)
    QBR._fsDirs = QBR._fsDirs.filter(function (d, i) { return i === 0 || used[d.id] || d.keep; });
    if (QBR._fsDirs.length !== before) dirsSave();
  };

  /* ---------------- 2026-10-30: per-file checkbox -> link through a remembered folder ----------------
   * Ticking Direct save on an UPLOADED workbook used to always open the file picker (an upload never tells
   * the browser where the file lives) and stored a plain file link that asked again per file next session.
   * Now the tick looks for the same file name in the folders already opened with "Open folder...":
   *  - a folder already allowed -> no prompt; none allowed -> ONE prompt for the most recent folder (inside the click)
   *  - same content as the upload -> linked silently (no reload); different content -> the user chooses
   *  - not found / declined -> null, and the caller opens the file picker as before (starting in that folder)
   * The link is stored as a folder link (dirId + rel), so the next session re-grants it with the folder. */
  var nameScan = {};   // dirId -> books[]; cleared on Open folder / Reconnect / Reload all
  QBR.fsFolderScanReset = function () { nameScan = {}; };

  // Small modal with buttons. buttons: [{id, label, primary}] -> resolves the clicked id, or null (Esc / Cancel).
  function askDialog(title, bodyHtml, buttons) {
    return new Promise(function (resolve) {
      var ov = document.createElement("div");
      ov.className = "merge-ov fo-ov fo-al"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "fo-al-h");
      ov.innerHTML = '<div class="merge-box"><h5 id="fo-al-h">' + e(title) + '</h5>' + bodyHtml +
        '<div class="merge-act">' + buttons.map(function (b) {
          return '<button type="button" class="btn ' + (b.primary ? "btn-primary" : "btn-outline-secondary") + '" data-al="' + e(b.id) + '">' + e(b.label) + '</button>';
        }).join(" ") + '</div></div>';
      document.body.appendChild(ov);
      var done = function (v) { ov.remove(); document.removeEventListener("keydown", esc); resolve(v); };
      var esc = function (ev) { if (ev.key === "Escape") done(null); };
      document.addEventListener("keydown", esc);
      ov.addEventListener("click", function (ev) {
        var b = ev.target.closest && ev.target.closest("[data-al]");
        if (!b) return;
        var v = b.dataset.al;
        if (v === "ok") { var r = ov.querySelector(".fo-al-pick:checked"); v = r ? "i" + r.value : "ok"; }
        done(v === "cancel" ? null : v);
      });
      setTimeout(function () { try { ov.querySelector(".btn-primary").focus(); } catch (x) {} }, 0);
    });
  }
  // The browser only opens a file picker during a click. When searching the folder took too long (or a
  // permission prompt ran first) that click has expired: ask for one more click. -> true when clicked
  QBR.fsAskClick = function (msg) {
    if (typeof QBR._autoAskClick === "function") return Promise.resolve(QBR._autoAskClick(msg));
    return askDialog("Choose the file", '<p class="small mb-2">' + e(msg) + '</p>',
      [{ id: "go", label: "Choose file…", primary: true }, { id: "cancel", label: "Cancel" }]).then(function (v) { return v === "go"; });
  };

  // The remembered folder that contains this file handle -> {id, rel} | null (no permission needed).
  QBR.fsFolderOf = async function (handle) {
    var all = QBR._fsDirs || [];
    for (var i = 0; i < all.length; i++) {
      try {
        if (typeof all[i].handle.resolve !== "function") continue;
        var rel = await all[i].handle.resolve(handle);
        if (rel && rel.length) return { id: all[i].id, rel: rel };
      } catch (x) {}
    }
    return null;
  };
  QBR.fsFolderStartIn = function () { var d = (QBR._fsDirs || [])[0]; return d && d.handle || null; };

  // Allowed folders; when none is allowed, ONE prompt for the most recent folder (must run inside the click).
  async function allowedDirs() {
    var all = QBR._fsDirs || [], ok = [];
    for (var i = 0; i < all.length; i++) {
      var p; try { p = await all[i].handle.queryPermission({ mode: "readwrite" }); } catch (x) { p = "denied"; }
      QBR._fsDirState[all[i].id] = p;
      if (p === "granted") ok.push(all[i]);
    }
    if (ok.length || !all.length) return { dirs: ok, asked: false };
    var granted = await QBR.fsFolderReconnect(all[0].id);   // one prompt; also re-allows its other linked files
    return { dirs: granted ? [all[0]] : [], asked: true, declined: !granted };
  }
  function kb(n) { return n ? Math.max(1, Math.round(n / 1024)).toLocaleString() + " KB" : "—"; }
  function when(t) { return t ? new Date(t).toLocaleString() : "—"; }
  function where(m) { return m.d.name + (m.b.path.length ? " / " + m.b.path.join(" / ") : ""); }

  async function storeLink(link) {
    QBR._fsLinks = (QBR._fsLinks || []).filter(function (l) {
      return fileKey(l.name) !== fileKey(link.name) && !(l.kinds || []).some(function (k) { return link.kinds.indexOf(k) >= 0; });
    });
    QBR._fsLinks.push(link);
    await fsLinksSave();
    if (typeof cacheSession === "function") { try { cacheSession(); } catch (x) {} }
    if (typeof QBR.fsRefreshStatus === "function") QBR.fsRefreshStatus();
    if (typeof renderAll === "function") { try { renderAll(); } catch (x) {} }
  }

  // -> true (linked) | false (cancelled / nothing to link) | null (caller opens the file picker)
  QBR.fsAutoLinkByName = async function (name) {
    if (!QBR.fsFolderSupported() || !(QBR._fsDirs || []).length) return null;
    var item = persistAppFiles().filter(function (it) { return fileKey(it.name) === fileKey(name); })[0];
    if (!item || !item.blob) return null;
    var upFp = QBR.fpOf(item.name, item.blob);
    var kinds = Array.from((QBR._kindByFp || {})[upFp] || []);
    if (!kinds.length) {   // same outcome as before (no link), without the detour through the picker
      note('"' + item.name + '" has no sheets the dashboard edits, so there is nothing to save to it. Use Open folder… to link it for ↻ reload.', 8000);
      return false;
    }
    var g = await allowedDirs();
    if (!g.dirs.length) return null;

    var key = nk(item.name), found = [];
    for (var pass = 0; pass < 2 && !found.length; pass++) {
      var cached = false;
      for (var i = 0; i < g.dirs.length; i++) {
        var d = g.dirs[i], books = null;
        if (pass === 0 && nameScan[d.id]) { books = nameScan[d.id]; cached = true; }
        else { try { books = (await scan(d.handle)).books; } catch (x) { books = []; } nameScan[d.id] = books; }
        for (var j = 0; j < books.length; j++) if (nk(books[j].name) === key) found.push({ d: d, b: books[j] });
      }
      if (!cached) break;   // a fresh scan already ran: no second pass
    }
    if (!found.length) {
      note('"' + item.name + '" was not found in ' + g.dirs.map(function (x) { return x.name; }).join(", ") + " — pick it.", 6000);
      return null;
    }

    // Same file as the upload? name|size|modified first; same size -> compare the bytes.
    var upHash = null;
    var uploadedHash = async function () {
      if (upHash == null) { var u = fsBaseOf(item.name); if (!u) u = new Uint8Array(await item.blob.arrayBuffer()); upHash = fsHashBytes(u); }
      return upHash;
    };
    var same = [], diff = [];
    for (var k = 0; k < found.length; k++) {
      var f; try { f = await found[k].b.handle.getFile(); } catch (x) { continue; }
      found[k].file = f;
      var isSame = QBR.fpOf(item.name, f) === upFp;
      if (!isSame && f.size === item.blob.size) {
        try { isSame = fsHashBytes(new Uint8Array(await f.arrayBuffer())) === (await uploadedHash()); } catch (x) { isSame = false; }
      }
      (isSame ? same : diff).push(found[k]);
    }

    if (same.length) {
      same.sort(function (a, b) { return b.file.lastModified - a.file.lastModified; });
      var m = same[0];
      // Link metadata = the copy that is loaded; a newer timestamp on identical bytes (OneDrive) is
      // rebased by the existing save path (size + hash equal -> fsRebaseLink).
      await storeLink({ fp: upFp, kinds: kinds, name: item.name, size: item.blob.size, lastModified: item.blob.lastModified,
        hash: await uploadedHash(), handle: m.b.handle, dirId: m.d.id, rel: m.b.rel });
      note("Linked ✓ " + item.name + " — found in " + where(m) + ". Direct save is on; next visit the folder's one permission covers it.", 6000);
      return true;
    }
    if (!diff.length) return null;

    diff.sort(function (a, b) { return b.file.lastModified - a.file.lastModified; });
    var pending = 0; try { pending = QBR.journalOpsFor(upFp).length; } catch (x) {}
    var choice;
    if (typeof QBR._autoLinkChoose === "function") choice = QBR._autoLinkChoose(diff.map(function (m) { return { where: where(m), size: m.file.size, lastModified: m.file.lastModified }; }), pending);
    else {
      var rows = diff.map(function (m, i) {
        return '<tr><td><input type="radio" name="fo-al" class="fo-al-pick" value="' + i + '"' + (i === 0 ? " checked" : "") + ' aria-label="' + e(where(m)) + '"></td>' +
          '<td>' + e(where(m)) + '</td><td class="small text-muted">' + e(when(m.file.lastModified)) + '</td><td class="small text-muted">' + e(kb(m.file.size)) + '</td></tr>';
      }).join("");
      choice = await askDialog("Different copy in the folder",
        '<p class="small mb-1">The folder has <b>' + e(item.name) + '</b>, but it is not the same as the copy you uploaded ' +
        '(modified ' + e(when(item.blob.lastModified)) + ', ' + e(kb(item.blob.size)) + ').</p>' +
        '<p class="small mb-1">Link it and reload the dashboard from it, or pick a different file.</p>' +
        (pending ? '<p class="small mb-1 text-danger">⚠ ' + pending + ' unsaved dashboard edit' + (pending === 1 ? "" : "s") +
          ' on the uploaded copy will not carry over to the folder copy.</p>' : "") +
        '<div class="merge-scroll"><table class="table table-sm fo-table"><thead><tr><th></th><th>Location</th><th>Modified</th><th>Size</th></tr></thead><tbody>' + rows + '</tbody></table></div>',
        [{ id: "ok", label: "Link & reload", primary: true }, { id: "picker", label: "Pick another file…" }, { id: "cancel", label: "Cancel" }]);
      if (choice && choice.charAt(0) === "i") choice = +choice.slice(1);
      else if (choice === "ok") choice = 0;
    }
    if (choice === "picker") return null;          // the button click is a fresh gesture for the picker
    if (choice == null || choice === false || !diff[choice]) return false;
    return linkAndReload(item, diff[choice]);
  };

  async function linkAndReload(item, m) {
    var mem;
    try { mem = await fsMemFile(m.file, item.name); }
    catch (x) { alert("Could not read " + where(m) + ": " + ((x && x.message) || x)); return false; }
    var map = new Map();
    persistAppFiles().forEach(function (it) { map.set(fileKey(it.name), it); });
    map.set(fileKey(item.name), { name: item.name, blob: mem.file });
    var ok = await loadItems(Array.from(map.values()), {});
    if (!ok) { note("Could not load " + where(m), 6000); return false; }
    var fp = QBR.fpOf(item.name, mem.file);
    var kinds = Array.from((QBR._kindByFp || {})[fp] || []);
    if (!kinds.length) {
      note("Reloaded " + item.name + " from " + where(m) + ", but it has no sheets the dashboard edits — not linked.", 8000);
      if (typeof cacheSession === "function") { try { cacheSession(); } catch (x) {} }
      return true;
    }
    await storeLink({ fp: fp, kinds: kinds, name: item.name, size: m.file.size, lastModified: m.file.lastModified,
      hash: fsHashBytes(mem.bytes), handle: m.b.handle, dirId: m.d.id, rel: m.b.rel });
    note("Linked ✓ " + item.name + " — reloaded from " + where(m) + ".", 6000);
    return true;
  }


  /* ---------------- toolbar in the loaded-files list ---------------- */
  QBR.fsFolderToolbarHtml = function () {
    if (!QBR.fsFolderSupported()) return "";
    var need = QBR.fsFoldersNeedingGrant();
    var hasFolderLinks = (QBR._fsLinks || []).some(function (l) { return l.dirId; });
    return '<div class="fo-bar">' +
      '<button type="button" class="btn btn-sm btn-outline-primary" data-fo-act="open" title="Pick the folder that holds your workbooks — one permission enables Direct save for all of them">📁 Open folder…</button>' +
      (need.length ? '<button type="button" class="btn btn-sm btn-warning" data-fo-act="reconnect" title="Your browser needs one click to re-allow editing in ' + e(need.map(function (d) { return d.name; }).join(", ")) + '">🔓 Reconnect folder</button>' : "") +
      (hasFolderLinks ? '<button type="button" class="btn btn-sm btn-outline-secondary" data-fo-act="reload" title="Reload every folder-linked workbook from disk (picks up edits made in Excel)">↻ Reload all</button>' : "") +
      '</div>';
  };
  // Delegated clicks (the list is re-rendered often).
  document.addEventListener("click", function (ev) {
    var b = ev.target.closest && ev.target.closest("[data-fo-act]");
    if (!b) return;
    ev.preventDefault(); ev.stopPropagation();
    var a = b.dataset.foAct;
    if (a === "open") QBR.fsOpenFolder();
    else if (a === "reconnect") QBR.fsFolderReconnectAll();
    else if (a === "reload") QBR.fsFolderReloadAll();
  }, true);

  function init() {
    if (!QBR.fsFolderSupported()) return;
    document.body.classList.add("fs-folder-ok");
    dirsLoad();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
