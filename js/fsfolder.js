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
    var slim = QBR._fsDirs.map(function (d) { return { id: d.id, name: d.name, handle: d.handle, picked: d.picked || [] }; });
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
  async function scan(dir) {
    var out = [], seen = 0;
    async function walk(h, path, depth) {
      for await (var entry of h.values()) {
        if (++seen > MAX_ENTRIES || out.length >= MAX_BOOKS) return;
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
    note("Scanning " + dir.name + "…", 15000);
    var res;
    try { res = await scan(dir); } catch (x) { alert("Could not read the folder: " + (x && x.message || x)); return false; }
    if (!res.books.length) { note(""); alert('No .xlsx / .xlsm workbooks found in "' + dir.name + '" (or its subfolders).'); return false; }
    note("");
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
    var others = QBR._fsDirs.filter(function (d) { return d !== rec; });
    var used = others.filter(function (d) { return usedIds[d.id]; });
    var spare = others.filter(function (d) { return !usedIds[d.id]; }).slice(0, Math.max(0, 4 - used.length));
    QBR._fsDirs = [rec].concat(others.filter(function (d) { return usedIds[d.id] || spare.indexOf(d) >= 0; }));
    QBR._fsDirState[rec.id] = "granted";

    note("Loading " + picks.length + " workbook" + (picks.length === 1 ? "" : "s") + "…", 30000);
    try {
      var mems = [];
      for (var j = 0; j < picks.length; j++) {
        var f = await picks[j].handle.getFile();
        var mem = await fsMemFile(f, f.name);
        mems.push({ b: picks[j], file: f, mem: mem });
      }
      // accumulate exactly like Upload: dedup by name, newest wins
      var map = new Map();
      ((typeof APP !== "undefined" && APP.files) || []).forEach(function (it) { map.set(fileKey(it.name), it); });
      mems.forEach(function (m) { map.set(fileKey(m.b.name), { name: m.b.name, blob: m.mem.file }); });
      var ok = await loadItems(Array.from(map.values()), {});
      if (!ok) { note("Could not load those workbooks", 6000); return false; }
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
      note("Loaded & linked ✓ " + picks.length + " workbook" + (picks.length === 1 ? "" : "s") + " from " + dir.name, 5000);
      return true;
    } catch (x) {
      alert("Could not load from the folder: " + (x && x.message || x));
      return false;
    }
  };

  /* ---------------- reconnect (next session) ---------------- */
  // One prompt for the folder; then re-resolve each linked file through it so
  // the file handles inherit the grant. Must run inside a click.
  QBR.fsFolderReconnect = async function (id) {
    var d = dirById(id); if (!d) return false;
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
  QBR.fsFolderReloadAll = async function () {
    if (!(await QBR.fsFolderReconnectAll())) return false;
    var done = 0, skipped = [];
    var list = (QBR._fsLinks || []).filter(function (l) { return l.dirId; });
    for (var i = 0; i < list.length; i++) {
      var l = list[i], pending = 0;
      try { pending = QBR.journalOpsFor(l.fp).length; } catch (x) {}
      if (pending) { skipped.push(l.name); continue; }
      try { if (await QBR.fsReloadFromFile(l.name, { quiet: true })) done++; } catch (x) {}
    }
    note("Reloaded ✓ " + done + " workbook" + (done === 1 ? "" : "s") +
      (skipped.length ? " · skipped (unsaved edits — save first): " + skipped.join(", ") : ""), skipped.length ? 12000 : 5000);
    return true;
  };

  // A link is gone → drop folders nobody uses any more (keeps the store tidy).
  QBR.fsFolderPrune = function () {
    var used = {};
    (QBR._fsLinks || []).forEach(function (l) { if (l.dirId) used[l.dirId] = 1; });
    var before = QBR._fsDirs.length;
    // keep the most recent folder even when unused (picker starts there)
    QBR._fsDirs = QBR._fsDirs.filter(function (d, i) { return i === 0 || used[d.id]; });
    if (QBR._fsDirs.length !== before) dirsSave();
  };

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
