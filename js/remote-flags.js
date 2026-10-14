/* ============================================================================
 * RCT OpsDesk — Remote feature flags (js/remote-flags.js)
 * v1.36.0 (2026-10-14) — NEW feature
 *
 * Lets the HOSTED app be managed remotely. A `config/features.json` file
 * deployed with the site is the source of truth for feature flags:
 *
 *   1. On startup the app fetches "config/features.json" (same origin — the
 *      offline-only rule stays intact; no external calls).
 *   2. The last-known-good copy is cached in localStorage ("qbr-remote-flags")
 *      so offline use and file:// still work.
 *   3. QBR.adminIsOn(key) is the single check every render path uses.
 *      Precedence: REMOTE ALWAYS WINS > per-browser local flags > default ON.
 *      A parent module turned off hides its children ("inventory" off hides
 *      "inventory.supplies").
 *
 * Future features self-declare their flags — no admin.js edits needed:
 *   QBR.features.register({ module: "inventory", key: "inventory.supplies",
 *                           label: "Supplies", desc: "…" });
 *
 * Enforcement is two-layer: nav elements carry data-feature="key" and are
 * hidden when the flag is off (plus a capture-phase click guard, so a
 * deep link / stray click can't reopen a disabled feature), and each render
 * function checks QBR.adminIsOn(key) before drawing its section.
 *
 * Publishing (Admin panel → Remote flags section, Pedro's machine only):
 * flipping toggles and hitting Publish commits the new JSON to the repo via
 * the GitHub Contents API using a personal access token that lives ONLY in
 * that browser's localStorage ("qbr-gh-pat"). No token → read-only view.
 * ============================================================================ */
(function () {
  "use strict";
  var QBR = (window.QBR = window.QBR || {});

  var REMOTE_URL = "config/features.json";
  var CACHE_KEY = "qbr-remote-flags";   // last-known-good remote copy
  var LOCAL_KEY = "qbr-feature-flags";  // existing per-browser module flags
  var PAT_KEY = "qbr-gh-pat";           // GitHub token — Pedro's machine only
  var REPO_KEY = "qbr-gh-repo";          // { owner, repo, branch }
  // Review fix: the default publish branch follows the host the admin is on, so a flag flip on a
  // beta/preview site never edits production by accident. "<branch>--site.netlify.app" -> that
  // branch, the Netlify main address -> develop (beta), anything else (GitHub Pages / Azure) -> main.
  function hostBranch() {
    var h = "";
    try { h = String(location.hostname || "").toLowerCase(); } catch (e) {}
    var m = /^([a-z0-9-]+)--[a-z0-9-]+\.netlify\.app$/.exec(h);
    if (m) return m[1];
    if (/\.netlify\.app$/.test(h)) return "develop";
    return "main";
  }
  var REPO_DEFAULTS = { owner: "Hayleyllujah12", repo: "QBR-Dashboard", branch: hostBranch() };

  /* ---------------- feature registry ---------------- */
  var _reg = [];
  var MODULE_LABELS = {
    inventory: "Inventory",
    security: "Security & identity",
    health: "Tenant health",
    adoption: "Adoption & capacity",
    mail: "Domains & email",
    reporting: "Reporting",
    audit: "Audit",
    data: "Data"
  };

  QBR.features = {
    // { module, key, label, desc }. Re-registering a key replaces it.
    register: function (f) {
      if (!f || typeof f.key !== "string" || !f.key) return false;
      var rec = {
        module: String(f.module || f.key.split(".")[0] || ""),
        key: f.key,
        label: String(f.label || f.key),
        desc: String(f.desc || "")
      };
      for (var i = 0; i < _reg.length; i++) {
        if (_reg[i].key === rec.key) { _reg[i] = rec; return true; }
      }
      _reg.push(rec);
      return true;
    },
    list: function () { return _reg.slice(); },
    get: function (key) {
      for (var i = 0; i < _reg.length; i++) if (_reg[i].key === key) return _reg[i];
      return null;
    },
    moduleLabel: function (m) {
      return MODULE_LABELS[m] || (m ? m.charAt(0).toUpperCase() + m.slice(1) : "");
    }
  };

  /* ---------------- remote load ---------------- */
  var _remote = null;    // parsed { version, updated, flags }
  var _source = "none";  // "live" | "cache" | "none"
  var _ready = false;

  function readCache() {
    try {
      var c = JSON.parse(localStorage.getItem(CACHE_KEY) || "null");
      if (c && c.data && c.data.flags && typeof c.data.flags === "object") return c.data;
    } catch (e) {}
    return null;
  }
  function writeCache(doc) {
    try { localStorage.setItem(CACHE_KEY, JSON.stringify({ savedAt: Date.now(), data: doc })); }
    catch (e) {}
  }
  function validDoc(d) {
    return d && typeof d === "object" && d.flags && typeof d.flags === "object";
  }

  async function loadRemote() {
    try {
      // Review fix: on file:// (the offline copy) there is no site to ask — fetch() fails and Chrome
      // logs "net::ERR_FAILED" on every load. Go straight to the last-known-good copy.
      var proto = "";
      try { proto = String(location.protocol || ""); } catch (e) {}
      if (proto === "file:") throw new Error("offline copy");
      // cache:"no-store" + bust param: a flag flip must reach clients fast.
      var res = await fetch(REMOTE_URL + "?v=" + Date.now(), { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      var doc = await res.json();
      if (!validDoc(doc)) throw new Error("bad shape");
      _remote = doc; _source = "live";
      writeCache(doc);
    } catch (e) {
      var cached = readCache();
      if (cached) { _remote = cached; _source = "cache"; }
      else { _remote = null; _source = "none"; }
    }
    _ready = true;
    applyAll();
    try { document.dispatchEvent(new CustomEvent("qbr:flags-ready")); } catch (e) {}
  }
  // Module groups (admin.js) + data-feature markers. adminApplyFlags calls applyNav itself.
  function applyAll() {
    try { if (typeof QBR.adminApplyFlags === "function") { QBR.adminApplyFlags(); return; } } catch (e) {}
    applyNav();
  }

  QBR.remoteFlagsStatus = function () {
    return { ready: _ready, source: _source, remote: _remote };
  };
  // Test seam: inject a remote doc without network.
  QBR._remoteFlagsTestLoad = function (docOrNull, source) {
    _remote = validDoc(docOrNull) ? docOrNull : null;
    _source = source || (_remote ? "live" : "none");
    _ready = true;
  };

  /* ---------------- effective flag: REMOTE ALWAYS WINS ---------------- */
  function parentOf(key) {
    var i = String(key).indexOf(".");
    return i < 0 ? null : String(key).slice(0, i);
  }
  function adminIsOn(key) {
    key = String(key || "");
    if (!key) return true;
    var p = parentOf(key);
    if (p && !adminIsOn(p)) return false;           // parent off hides children
    if (_remote && _remote.flags &&
        Object.prototype.hasOwnProperty.call(_remote.flags, key))
      return !!_remote.flags[key];                  // remote wins
    try {
      var local = JSON.parse(localStorage.getItem(LOCAL_KEY) || "{}") || {};
      if (Object.prototype.hasOwnProperty.call(local, key)) return !!local[key];
    } catch (e) {}
    return true;                                    // default: on
  }
  QBR.adminIsOn = adminIsOn;

  /* ---------------- nav enforcement ---------------- */
  function firstVisibleTab() {
    try {
      return document.querySelector(
        '.sb-group:not([style*="none"]) .tab-btn:not([style*="none"])');
    } catch (e) { return null; }
  }
  function applyNav() {
    try {
      _reg.forEach(function (f) {
        var on = adminIsOn(f.key);
        document.querySelectorAll('[data-feature="' + f.key + '"]').forEach(function (el) {
          el.style.display = on ? "" : "none";
        });
      });
      // If the active tab just got hidden, fall back to the first visible tab.
      var active = document.querySelector(".tab-btn.active");
      if (active && (active.style.display === "none" ||
          (active.closest(".sb-group") && active.closest(".sb-group").style.display === "none"))) {
        var first = firstVisibleTab();
        if (first && first !== active) first.click();
      }
    } catch (e) {}
  }
  QBR.remoteFlagsApplyNav = applyNav;

  // Capture-phase guard: a disabled feature can't be opened by click/deep-link.
  // Runs before app.js's own [data-tab] handler, so stopPropagation wins.
  function guardClick(e) {
    try {
      var t = e.target;
      var el = t && t.closest ? t.closest("[data-feature]") : null;
      if (!el) return;
      var key = el.getAttribute("data-feature");
      if (adminIsOn(key)) return;
      // Review fix: just block. (Redirecting to the first tab yanked the user off the page they
      // were on; a deep link to a disabled page is handled by the page's own "disabled" notice.)
      e.preventDefault();
      e.stopPropagation();
    } catch (err) {}
  }

  /* ---------------- GitHub publish (token = Pedro's machine only) ---------------- */
  function b64(s) {
    // btoa is Latin-1 only; round-trip through UTF-8 first.
    return btoa(unescape(encodeURIComponent(s)));
  }
  // Review fix: the token can write to the repo that GitHub Pages serves, so it is kept for this
  // browser TAB only (sessionStorage) unless the admin ticks "Remember on this machine".
  function readToken() {
    var t = null;
    try { t = sessionStorage.getItem(PAT_KEY); } catch (e) {}
    if (!t) { try { t = localStorage.getItem(PAT_KEY); } catch (e) {} }
    return t || null;
  }
  QBR.remoteFlagsToken = {
    present: function () { return !!readToken(); },
    remembered: function () { try { return !!localStorage.getItem(PAT_KEY); } catch (e) { return false; } },
    set: function (t, remember) {
      if (!t) return false;
      try {
        if (remember) { localStorage.setItem(PAT_KEY, t); try { sessionStorage.removeItem(PAT_KEY); } catch (e) {} }
        else { sessionStorage.setItem(PAT_KEY, t); try { localStorage.removeItem(PAT_KEY); } catch (e) {} }
        return true;
      } catch (e) { return false; }
    },
    clear: function () {
      try { localStorage.removeItem(PAT_KEY); } catch (e) {}
      try { sessionStorage.removeItem(PAT_KEY); } catch (e) {}
    }
  };
  QBR.remoteFlagsHostBranch = hostBranch;
  QBR.remoteFlagsRepo = {
    get: function () {
      var c = {};
      try { c = JSON.parse(localStorage.getItem(REPO_KEY) || "{}") || {}; } catch (e) {}
      return {
        owner: c.owner || REPO_DEFAULTS.owner,
        repo: c.repo || REPO_DEFAULTS.repo,
        branch: c.branch || REPO_DEFAULTS.branch
      };
    },
    set: function (c) {
      try { localStorage.setItem(REPO_KEY, JSON.stringify(c || {})); return true; }
      catch (e) { return false; }
    }
  };

  // flags: { key: bool }. Commits config/features.json via the Contents API.
  // Creates the file on first publish (no sha). Throws with a readable message.
  QBR.remoteFlagsPublish = async function (flags) {
    var repo = QBR.remoteFlagsRepo.get();
    var token = readToken();
    if (!token) throw new Error("No GitHub token on this machine — publishing is disabled.");
    if (!repo.owner || !repo.repo) throw new Error("Repository not configured.");
    var branch = repo.branch || "main";
    var api = "https://api.github.com/repos/" + encodeURIComponent(repo.owner) + "/" +
      encodeURIComponent(repo.repo) + "/contents/" + REMOTE_URL;
    var headers = { "Accept": "application/vnd.github+json", "Authorization": "Bearer " + token };
    var sha = null, curDoc = null;
    var cur = await fetch(api + "?ref=" + encodeURIComponent(branch), { headers: headers });
    if (cur.status === 200) {
      var meta = await cur.json();
      sha = meta.sha || null;
      // Review fix: read the CURRENT file so keys this build doesn't register (e.g. the "inventory"
      // parent, or flags added by a newer build) survive, and the version can't go backwards when
      // this browser only had a stale cached copy.
      try {
        var raw = decodeURIComponent(escape(atob(String(meta.content || "").replace(/\s/g, ""))));
        var parsed = JSON.parse(raw);
        if (validDoc(parsed)) curDoc = parsed;
      } catch (e) {}
    } else if (cur.status !== 404) {
      throw new Error("GitHub read failed (HTTP " + cur.status + "). Check the token and repo.");
    }
    var vCur = curDoc && typeof curDoc.version === "number" ? curDoc.version : 0;
    var vMine = _remote && typeof _remote.version === "number" ? _remote.version : 0;
    var doc = {
      version: Math.max(vCur, vMine) + 1,
      updated: new Date().toISOString().slice(0, 10),
      flags: Object.assign({}, curDoc ? curDoc.flags : (_remote ? _remote.flags : {}), flags || {})
    };
    var body = {
      message: "Update remote feature flags (RCT OpsDesk admin)",
      content: b64(JSON.stringify(doc, null, 2) + "\n"),
      branch: branch
    };
    if (sha) body.sha = sha;
    var put = await fetch(api, {
      method: "PUT",
      headers: Object.assign({ "Content-Type": "application/json" }, headers),
      body: JSON.stringify(body)
    });
    if (!put.ok) {
      var detail = "";
      try { detail = await put.text(); } catch (e) {}
      throw new Error("GitHub write failed (HTTP " + put.status + "). " + detail.slice(0, 180));
    }
    _remote = doc; _source = "live";
    writeCache(doc);
    applyAll();
    try { document.dispatchEvent(new CustomEvent("qbr:flags-ready")); } catch (e) {}
    return doc;
  };

  /* ---------------- boot ---------------- */
  try { document.addEventListener("click", guardClick, true); } catch (e) {}
  try {
    if (document.readyState === "loading")
      document.addEventListener("DOMContentLoaded", applyNav);
    else
      applyNav();
  } catch (e) {}
  try { loadRemote(); } catch (e) {}

  /* ---------- Node export guard (test harness only; inert in browser) ------ */
  if (typeof module !== "undefined" && module.exports) {
    module.exports = { QBR: QBR, _test: {
      parentOf: parentOf, b64: b64, validDoc: validDoc, hostBranch: hostBranch,
      get remote() { return _remote; }, get source() { return _source; }
    } };
  }
})();
