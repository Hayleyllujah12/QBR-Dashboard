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
 * main / develop → a pull request (branch flags/<stamp>) that a person merges on GitHub;
 * preview branches → a direct commit. Uses a fine-grained GitHub token kept per tab (sessionStorage,
 * 30-min idle lock) or encrypted with a passphrase ("Remember") — see "token security" below.
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
    data: "Data",
    certification: "Certification"
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
  /* ---------------- token security (v1.36.0, hardened 2026-10-14) ----------------
   * The token can write to the repo the live site is built from, so:
   *  - only FINE-GRAINED tokens ("github_pat_…") are accepted. Classic tokens (ghp_/gho_/…) reach
   *    every repo the owner has and are refused.
   *  - by default it lives in this TAB only (sessionStorage) and locks after 30 min without use.
   *  - "Remember on this machine" stores it ENCRYPTED (AES-GCM, key from a passphrase via PBKDF2,
   *    250k rounds). The passphrase is never stored; after a reload the token stays locked until
   *    the passphrase is typed again. Needs https (WebCrypto) — not offered on file://.
   *  - a plain-text token left by an earlier build in localStorage is deleted on load.
   *  - it is only ever sent to https://api.github.com (fixed here; owner/repo are validated).
   * -------------------------------------------------------------------------------- */
  var ENC_KEY = "qbr-gh-pat-enc";
  var IDLE_MS = 30 * 60 * 1000;
  var API = "https://api.github.com";
  var PROTECTED_BRANCHES = ["main", "develop"];   // publish = pull request, never a direct commit
  var _purgedPlain = false;
  try { if (localStorage.getItem(PAT_KEY)) { localStorage.removeItem(PAT_KEY); _purgedPlain = true; } } catch (e) {}

  function now() { return Date.now(); }
  function readSession() {
    var raw = null;
    try { raw = sessionStorage.getItem(PAT_KEY); } catch (e) {}
    if (!raw) return null;
    var o = null;
    try { o = JSON.parse(raw); } catch (e) {}
    if (!o || typeof o.t !== "string") { try { sessionStorage.removeItem(PAT_KEY); } catch (e) {} return null; }
    if (now() - (o.at || 0) > IDLE_MS) { try { sessionStorage.removeItem(PAT_KEY); } catch (e) {} return null; }
    return o.t;
  }
  function writeSession(t) {
    try { sessionStorage.setItem(PAT_KEY, JSON.stringify({ t: t, at: now() })); return true; } catch (e) { return false; }
  }
  function readToken() {
    var t = readSession();
    if (t) writeSession(t);   // use = activity → idle timer restarts
    return t;
  }
  function tokenFormatError(t) {
    t = String(t || "").trim();
    if (!t) return "Paste a token first.";
    if (/^(ghp|gho|ghu|ghs|ghr)_/.test(t))
      return "That is a classic token — it can reach every repository on the account. Create a fine-grained token (starts with github_pat_) for this repository only.";
    if (!/^github_pat_[A-Za-z0-9_]{20,}$/.test(t)) return "That doesn't look like a GitHub fine-grained token (it should start with github_pat_).";
    return "";
  }
  function subtle() {
    try {
      if (typeof crypto !== "undefined" && crypto.subtle && (typeof isSecureContext === "undefined" || isSecureContext !== false)) return crypto.subtle;
    } catch (e) {}
    return null;
  }
  function bytesToB64(u8) { var s = ""; for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return btoa(s); }
  function b64ToBytes(s) { var b = atob(s), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
  async function deriveKey(pass, salt) {
    var sc = subtle();
    var base = await sc.importKey("raw", new TextEncoder().encode(String(pass)), "PBKDF2", false, ["deriveKey"]);
    return sc.deriveKey({ name: "PBKDF2", salt: salt, iterations: 250000, hash: "SHA-256" }, base,
      { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  }
  async function encryptToken(t, pass) {
    var salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
    var key = await deriveKey(pass, salt);
    var ct = new Uint8Array(await subtle().encrypt({ name: "AES-GCM", iv: iv }, key, new TextEncoder().encode(t)));
    return { v: 1, salt: bytesToB64(salt), iv: bytesToB64(iv), ct: bytesToB64(ct), savedAt: now() };
  }
  async function decryptToken(rec, pass) {
    var key = await deriveKey(pass, b64ToBytes(rec.salt));
    var pt = await subtle().decrypt({ name: "AES-GCM", iv: b64ToBytes(rec.iv) }, key, b64ToBytes(rec.ct));
    return new TextDecoder().decode(pt);
  }
  function readEnc() {
    try { var r = JSON.parse(localStorage.getItem(ENC_KEY) || "null"); return r && r.v === 1 && r.ct ? r : null; } catch (e) { return null; }
  }

  QBR.remoteFlagsToken = {
    IDLE_MIN: IDLE_MS / 60000,
    canRemember: function () { return !!subtle(); },
    formatError: tokenFormatError,
    present: function () { return !!readSession(); },          // usable now (doesn't extend the timer)
    remembered: function () { return !!readEnc(); },            // encrypted copy on this machine
    locked: function () { return !readSession() && !!readEnc(); },
    purgedPlain: function () { return _purgedPlain; },
    // opts: { remember: bool, passphrase: string }. Resolves true, or throws a readable Error.
    set: async function (t, opts) {
      opts = opts || {};
      t = String(t || "").trim();
      var bad = tokenFormatError(t);
      if (bad) throw new Error(bad);
      if (opts.remember) {
        if (!subtle()) throw new Error("This browser can't encrypt here (needs the https site) — the token is kept for this tab only.");
        if (!opts.passphrase || String(opts.passphrase).length < 8) throw new Error("Use a passphrase of at least 8 characters.");
        var rec = await encryptToken(t, opts.passphrase);
        try { localStorage.setItem(ENC_KEY, JSON.stringify(rec)); } catch (e) { throw new Error("Could not save in this browser."); }
      } else {
        try { localStorage.removeItem(ENC_KEY); } catch (e) {}
      }
      if (!writeSession(t)) throw new Error("Could not keep the token in this tab.");
      return true;
    },
    unlock: async function (pass) {
      var rec = readEnc();
      if (!rec) throw new Error("No saved token on this machine.");
      if (!subtle()) throw new Error("Unlocking needs the https site.");
      var t;
      try { t = await decryptToken(rec, pass || ""); } catch (e) { throw new Error("Wrong passphrase."); }
      writeSession(t);
      return true;
    },
    lock: function () { try { sessionStorage.removeItem(PAT_KEY); } catch (e) {} },
    clear: function () {
      try { localStorage.removeItem(PAT_KEY); } catch (e) {}
      try { localStorage.removeItem(ENC_KEY); } catch (e) {}
      try { sessionStorage.removeItem(PAT_KEY); } catch (e) {}
    }
  };
  QBR.remoteFlagsHostBranch = hostBranch;
  QBR.remoteFlagsIsProtected = function (b) { return PROTECTED_BRANCHES.indexOf(String(b || "main")) >= 0; };
  var NAME_RE = /^[A-Za-z0-9_.-]{1,100}$/, BRANCH_RE = /^[A-Za-z0-9_.\/-]{1,200}$/;
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

  function gh(token) {
    var headers = { "Accept": "application/vnd.github+json", "Authorization": "Bearer " + token, "X-GitHub-Api-Version": "2022-11-28" };
    return async function (method, path, body) {
      var init = { method: method, headers: headers };
      if (body) { init.headers = Object.assign({ "Content-Type": "application/json" }, headers); init.body = JSON.stringify(body); }
      var res = await fetch(API + path, init);
      var data = null;
      try { data = await res.json(); } catch (e) {}
      return { status: res.status, ok: res.ok, data: data, headers: res.headers };
    };
  }
  function repoPath(repo) {
    if (!NAME_RE.test(repo.owner || "") || !NAME_RE.test(repo.repo || "")) throw new Error("Repository not configured (owner / repo name).");
    return "/repos/" + repo.owner + "/" + repo.repo;
  }
  function ghErr(what, r) {
    var m = r && r.data && r.data.message ? " — " + String(r.data.message).slice(0, 160) : "";
    var hint = r && (r.status === 401) ? " The token is invalid or expired." :
      r && (r.status === 403 || r.status === 404) ? " Check the token has access to this repository (Contents + Pull requests: read & write)." : "";
    return new Error(what + " failed (HTTP " + (r ? r.status : "?") + ")" + m + "." + hint);
  }

  // Checks the token can see the repo; reports its expiry when GitHub exposes it.
  QBR.remoteFlagsCheckToken = async function (tokenOpt) {
    var token = tokenOpt || readToken();
    if (!token) throw new Error("No token.");
    var repo = QBR.remoteFlagsRepo.get();
    var r = await gh(token)("GET", repoPath(repo));
    if (!r.ok) throw ghErr("Token check", r);
    var exp = null;
    try { exp = r.headers && r.headers.get ? r.headers.get("github-authentication-token-expiration") : null; } catch (e) {}
    var days = null;
    if (exp) { var d = Date.parse(String(exp).replace(" UTC", "Z").replace(" ", "T")); if (isFinite(d)) days = Math.round((d - now()) / 86400000); }
    var perms = (r.data && r.data.permissions) || {};
    return { ok: true, expires: exp, daysLeft: days, admin: !!perms.admin };
  };

  async function readFlagsFile(call, rp, ref) {
    var cur = await call("GET", rp + "/contents/" + REMOTE_URL + "?ref=" + encodeURIComponent(ref));
    if (cur.status === 404) return { sha: null, doc: null };
    if (!cur.ok) throw ghErr("Reading the flags file", cur);
    var doc = null;
    // Read the CURRENT file so keys this build doesn't register (the "inventory" parent, flags from a newer
    // build) survive, and the version can't go backwards when this browser only had a stale cached copy.
    try {
      var parsed = JSON.parse(decodeURIComponent(escape(atob(String(cur.data.content || "").replace(/\s/g, "")))));
      if (validDoc(parsed)) doc = parsed;
    } catch (e) {}
    return { sha: cur.data.sha || null, doc: doc };
  }
  function diffLines(oldFlags, newFlags) {
    var keys = {}, out = [];
    Object.keys(oldFlags || {}).concat(Object.keys(newFlags || {})).forEach(function (k) { keys[k] = 1; });
    Object.keys(keys).sort().forEach(function (k) {
      var a = oldFlags && Object.prototype.hasOwnProperty.call(oldFlags, k) ? !!oldFlags[k] : null, b = !!newFlags[k];
      if (a !== b) out.push("- `" + k + "`: " + (a === null ? "(new)" : a ? "on" : "off") + " → **" + (b ? "on" : "off") + "**");
    });
    return out;
  }

  // The flags file as it is NOW on another branch (e.g. main while you are on a preview), so the admin
  // panel can show — and diff against — that branch's real values. Resolves { branch, doc|null }.
  QBR.remoteFlagsReadTarget = async function (branch) {
    var token = readToken();
    if (!token) throw new Error("No GitHub token in this tab.");
    if (!BRANCH_RE.test(String(branch || ""))) throw new Error("Invalid branch.");
    var r = await readFlagsFile(gh(token), repoPath(QBR.remoteFlagsRepo.get()), branch);
    return { branch: branch, doc: r.doc };
  };
  QBR.remoteFlagsDiff = function (oldFlags, newFlags) {
    return diffLines(oldFlags || {}, newFlags || {}).map(function (l) { return l.replace(/`/g, "").replace(/\*\*/g, "").replace(/^- /, ""); });
  };

  // flags: { key: bool }. main / develop → opens a PULL REQUEST (branch flags/<stamp>) that someone merges on
  // GitHub; any other (preview) branch → direct commit. Resolves { doc, pr|null, direct }.
  QBR.remoteFlagsPublish = async function (flags) {
    var token = readToken();
    if (!token) throw new Error(QBR.remoteFlagsToken.locked() ? "The saved token is locked — unlock it first." : "No GitHub token in this tab — publishing is disabled.");
    var repo = QBR.remoteFlagsRepo.get();
    var rp = repoPath(repo);
    var base = repo.branch || "main";
    if (!BRANCH_RE.test(base)) throw new Error("Invalid target branch.");
    var call = gh(token);
    var cur = await readFlagsFile(call, rp, base);
    var vCur = cur.doc && typeof cur.doc.version === "number" ? cur.doc.version : 0;
    var vMine = _remote && typeof _remote.version === "number" ? _remote.version : 0;
    var oldFlags = cur.doc ? cur.doc.flags : (_remote ? _remote.flags : {});
    var doc = {
      version: Math.max(vCur, vMine) + 1,
      updated: new Date().toISOString().slice(0, 10),
      flags: Object.assign({}, oldFlags, flags || {})
    };
    var changes = diffLines(oldFlags, doc.flags);
    var content = b64(JSON.stringify(doc, null, 2) + "\n");
    var message = "Update remote feature flags v" + doc.version + " (RCT OpsDesk admin)";

    if (!QBR.remoteFlagsIsProtected(base)) {
      var put = await call("PUT", rp + "/contents/" + REMOTE_URL, cur.sha ? { message: message, content: content, branch: base, sha: cur.sha } : { message: message, content: content, branch: base });
      if (!put.ok) throw ghErr("Writing the flags file", put);
      _remote = doc; _source = "live";
      writeCache(doc);
      applyAll();
      try { document.dispatchEvent(new CustomEvent("qbr:flags-ready")); } catch (e) {}
      return { doc: doc, pr: null, direct: true };
    }

    // Pull-request flow: nothing reaches `base` until a person merges the PR on GitHub.
    var ref = await call("GET", rp + "/git/ref/heads/" + encodeURIComponent(base));
    if (!ref.ok || !ref.data || !ref.data.object) throw ghErr("Reading branch " + base, ref);
    var d = new Date(), pad = function (n) { return (n < 10 ? "0" : "") + n; };
    var head = "flags/" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "-" + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + "-v" + doc.version;
    var mk = await call("POST", rp + "/git/refs", { ref: "refs/heads/" + head, sha: ref.data.object.sha });
    if (!mk.ok) throw ghErr("Creating branch " + head, mk);
    var put2 = await call("PUT", rp + "/contents/" + REMOTE_URL, cur.sha ? { message: message, content: content, branch: head, sha: cur.sha } : { message: message, content: content, branch: head });
    if (!put2.ok) throw ghErr("Writing the flags file", put2);
    var pr = await call("POST", rp + "/pulls", {
      title: "Feature flags v" + doc.version + " → " + base,
      head: head, base: base,
      body: "Published from RCT OpsDesk › Admin › Remote flags.\n\n" + (changes.length ? changes.join("\n") : "_No flag values changed (version bump only)._") +
        "\n\nMerging this pull request updates `config/features.json` on `" + base + "`; the site picks it up after it rebuilds."
    });
    if (!pr.ok || !pr.data) throw ghErr("Opening the pull request", pr);
    var url = String(pr.data.html_url || "");
    if (url.indexOf("https://github.com/") !== 0) url = "";
    return { doc: doc, pr: { number: pr.data.number, url: url, head: head }, direct: false };
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
