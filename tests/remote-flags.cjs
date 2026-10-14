#!/usr/bin/env node
/* ============================================================================
 * tests/remote-flags.cjs — logic tests for js/remote-flags.js (v1.36.0).
 *
 * Zero install: stubs window/document/localStorage/fetch, loads the real
 * module fresh per scenario (require cache cleared), and asserts:
 *   - feature registry (register/list/get/replace/moduleLabel)
 *   - QBR.adminIsOn precedence: REMOTE > local > default ON; parent off
 *     hides children
 *   - loader fallback chain: live → cached copy → defaults
 *   - publish: GitHub Contents API payload shape (GET sha → PUT), no-token
 *     refusal, version bump
 *   - nav enforcement: data-feature elements hidden when off; capture-phase
 *     click guard blocks disabled features
 *   - inventory.js pilot helpers: invEnabledViews / invFeatureOn
 *
 * Usage: node tests/remote-flags.cjs
 * ========================================================================== */
"use strict";
const path = require("path");
const MOD = path.resolve(__dirname, "../js/remote-flags.js");
const INV = path.resolve(__dirname, "../js/inventory.js");

let pass = 0, fail = 0;
function ok(cond, msg, detail) {
  if (cond) { pass++; console.log("  ✓ " + msg); }
  else { fail++; console.log("  ✗ " + msg + (detail ? "  (" + detail + ")" : "")); }
}
function eq(a, b) { return JSON.stringify(a) === JSON.stringify(b); }

// ---- harness: fresh browser-ish globals per scenario -----------------------
function freshEnv(opts) {
  opts = opts || {};
  delete require.cache[require.resolve(MOD)];
  delete require.cache[require.resolve(INV)];
  const store = Object.assign({}, opts.localStorage || {});
  const fetchCalls = [];
  global.window = global;
  global.document = {
    readyState: "complete",
    _listeners: {},
    addEventListener(type, fn, cap) { this._listeners[type + (cap ? ":cap" : "")] = fn; },
    dispatchEvent() { return true; },
    querySelectorAll: opts.querySelectorAll || (() => []),
    querySelector: opts.querySelector || (() => null),
  };
  const sstore = {};
  global.location = { hostname: opts.hostname || "hayleyllujah12.github.io", pathname: "/", search: "" };
  global.sessionStorage = {
    getItem: k => (k in sstore ? sstore[k] : null),
    setItem: (k, v) => { sstore[k] = String(v); },
    removeItem: k => { delete sstore[k]; },
  };
  global.localStorage = {
    getItem: k => (k in store ? store[k] : null),
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: k => { delete store[k]; },
  };
  global.fetch = async (url, init) => {
    fetchCalls.push({ url: String(url), init: init || {} });
    return opts.fetch(url, init || {});
  };
  const m = require(MOD);
  return { QBR: m.QBR, test: m._test, store, sstore, fetchCalls };
}
function jsonRes(obj, status) {
  return { ok: (status || 200) < 300, status: status || 200, json: async () => obj, text: async () => JSON.stringify(obj) };
}
async function settled(QBR) {
  for (let i = 0; i < 50 && !QBR.remoteFlagsStatus().ready; i++)
    await new Promise(r => setTimeout(r, 10));
}

(async () => {
// ============ 1. registry ====================================================
{
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(QBR);
  ok(QBR.features.register({ module: "inventory", key: "inventory.scan", label: "Scan", desc: "d" }) === true, "register returns true");
  ok(QBR.features.register({}) === false, "register rejects empty record");
  ok(QBR.features.register({ module: "x" }) === false, "register rejects missing key");
  QBR.features.register({ module: "inventory", key: "inventory.scan", label: "Scan v2" });
  const got = QBR.features.get("inventory.scan");
  ok(got && got.label === "Scan v2" && QBR.features.list().length === 1, "re-registering a key replaces it");
  ok(QBR.features.get("nope") === null, "get returns null for unknown key");
  ok(QBR.features.moduleLabel("inventory") === "Inventory", "known module label");
  ok(QBR.features.moduleLabel("whatever") === "Whatever", "unknown module label falls back to capitalized");
  QBR.features.register({ key: "solo" });
  ok(QBR.features.get("solo").module === "solo", "module defaults to key prefix");
}

// ============ 2. adminIsOn precedence =========================================
{
  const live = { version: 2, updated: "2026-10-14", flags: { "inventory.scan": false, "inventory": true } };
  const { QBR } = freshEnv({
    fetch: async () => jsonRes(live),
    localStorage: { "qbr-feature-flags": JSON.stringify({ "inventory.scan": true, "legacy": false }) },
  });
  await settled(QBR);
  const st = QBR.remoteFlagsStatus();
  ok(st.source === "live" && st.ready === true, "loader reports live source when fetch works");
  ok(QBR.adminIsOn("inventory.scan") === false, "remote OFF wins over local ON");
  ok(QBR.adminIsOn("legacy") === false, "local flag honored when remote lacks the key");
  ok(QBR.adminIsOn("inventory.supplies") === true, "unlisted key defaults ON");
  ok(QBR.adminIsOn("") === true, "empty key defaults ON");
}
{
  // parent off hides children even when the child is explicitly on
  const { QBR } = freshEnv({
    fetch: async () => jsonRes({ version: 1, flags: { "inventory": false, "inventory.supplies": true } }),
  });
  await settled(QBR);
  ok(QBR.adminIsOn("inventory") === false, "parent OFF");
  ok(QBR.adminIsOn("inventory.supplies") === false, "parent OFF hides child even when child is ON remotely");
  ok(QBR.adminIsOn("inventory.scan") === false, "parent OFF hides unlisted child too");
}
{
  // remote wins over local for module keys as well
  const { QBR } = freshEnv({
    fetch: async () => jsonRes({ version: 1, flags: { "inventory": false } }),
    localStorage: { "qbr-feature-flags": JSON.stringify({ "inventory": true }) },
  });
  await settled(QBR);
  ok(QBR.adminIsOn("inventory") === false, "remote module flag wins over local module flag");
}

// ============ 3. loader fallback chain ========================================
{
  // fetch fails → cached copy used
  const cached = { version: 5, updated: "2026-10-10", flags: { "inventory.tickets": false } };
  const { QBR, store } = freshEnv({
    fetch: async () => { throw new Error("offline"); },
    localStorage: { "qbr-remote-flags": JSON.stringify({ savedAt: 1, data: cached }) },
  });
  await settled(QBR);
  const st = QBR.remoteFlagsStatus();
  ok(st.source === "cache", "fetch failure falls back to cached copy");
  ok(QBR.adminIsOn("inventory.tickets") === false, "cached flags are honored");
  ok(store["qbr-remote-flags"] !== undefined, "cache entry preserved");
}
{
  // fetch fails, no cache → defaults, nothing explodes
  const { QBR } = freshEnv({ fetch: async () => ({ ok: false, status: 404 }) });
  await settled(QBR);
  const st = QBR.remoteFlagsStatus();
  ok(st.source === "none" && st.remote === null, "no fetch + no cache → source none");
  ok(QBR.adminIsOn("inventory.scan") === true, "defaults to ON when nothing loaded");
}
{
  // fetch ok → cache written for next offline run
  const { QBR, store } = freshEnv({
    fetch: async () => jsonRes({ version: 7, updated: "2026-10-14", flags: { "a.b": false } }),
  });
  await settled(QBR);
  const c = JSON.parse(store["qbr-remote-flags"] || "null");
  ok(c && c.data.version === 7 && c.data.flags["a.b"] === false, "live load writes last-known-good cache");
}
{
  // malformed JSON body → treated as failure, cache used
  const { QBR } = freshEnv({
    fetch: async () => jsonRes({ nope: 1 }),
    localStorage: { "qbr-remote-flags": JSON.stringify({ savedAt: 1, data: { version: 3, flags: {} } }) },
  });
  await settled(QBR);
  ok(QBR.remoteFlagsStatus().source === "cache", "bad-shaped remote doc falls back to cache");
}

// ============ 4. publish ======================================================
{
  let putBody = null;
  const { QBR, fetchCalls } = freshEnv({
    fetch: async (url, init) => {
      if ((init.method || "GET") === "GET") return jsonRes({ sha: "abc123" });
      putBody = JSON.parse(init.body);
      return { ok: true, status: 201, json: async () => ({}), text: async () => "" };
    },
    localStorage: {
      "qbr-gh-pat": "tok123",
      "qbr-gh-repo": JSON.stringify({ owner: "Hayleyllujah12", repo: "QBR-Dashboard", branch: "main" }),
      "qbr-remote-flags": JSON.stringify({ savedAt: 1, data: { version: 4, updated: "2026-10-01", flags: {} } }),
    },
  });
  await settled(QBR);
  fetchCalls.length = 0; // drop the loader's own startup fetch
  const doc = await QBR.remoteFlagsPublish({ "inventory.scan": false, "inventory.supplies": true });
  ok(fetchCalls.length === 2, "publish does GET (sha) then PUT", "calls=" + fetchCalls.length);
  ok(/\/repos\/Hayleyllujah12\/QBR-Dashboard\/contents\/config\/features\.json/.test(fetchCalls[0].url), "targets config/features.json in the configured repo");
  ok(fetchCalls[0].init.headers.Authorization === "Bearer tok123", "token sent as Bearer");
  ok(putBody.sha === "abc123" && putBody.branch === "main", "PUT carries sha + branch");
  ok(/Update remote feature flags/.test(putBody.message), "PUT has a commit message");
  const decoded = JSON.parse(Buffer.from(putBody.content, "base64").toString("utf8"));
  ok(decoded.flags["inventory.scan"] === false && decoded.flags["inventory.supplies"] === true, "PUT content carries the flags");
  ok(decoded.version === 5, "version bumps from the loaded remote doc");
  ok(doc.flags["inventory.scan"] === false, "publish resolves with the new doc");
  ok(QBR.remoteFlagsStatus().source === "live", "local copy refreshed after publish");
}
{
  // first publish (file doesn't exist yet): PUT without sha
  let putBody = null;
  const { QBR } = freshEnv({
    fetch: async (url, init) => {
      if ((init.method || "GET") === "GET") return { ok: false, status: 404, json: async () => ({}), text: async () => "" };
      putBody = JSON.parse(init.body);
      return { ok: true, status: 201, json: async () => ({}), text: async () => "" };
    },
    localStorage: { "qbr-gh-pat": "tok", "qbr-gh-repo": JSON.stringify({ owner: "o", repo: "r", branch: "main" }) },
  });
  await settled(QBR);
  await QBR.remoteFlagsPublish({ "x": true });
  ok(!("sha" in putBody), "first publish omits sha (file creation)");
  ok(JSON.parse(Buffer.from(putBody.content, "base64").toString("utf8")).version === 1, "first publish starts at version 1");
}
{
  // no token on this machine → publish refuses
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("should not be called"); } });
  await settled(QBR);
  let err = null;
  try { await QBR.remoteFlagsPublish({}); } catch (e) { err = e; }
  ok(err && /No GitHub token/.test(err.message), "publish without a token throws");
}
{
  // token helpers: present/set/clear
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(QBR);
  ok(QBR.remoteFlagsToken.present() === false, "no token initially");
  QBR.remoteFlagsToken.set("abc");
  ok(QBR.remoteFlagsToken.present() === true, "token present after set");
  QBR.remoteFlagsToken.clear();
  ok(QBR.remoteFlagsToken.present() === false, "token gone after clear");
  const repo = QBR.remoteFlagsRepo.get();
  ok(repo.owner === "Hayleyllujah12" && repo.repo === "QBR-Dashboard" && repo.branch === "main", "repo defaults to Pedro's repo/main");
}

// ============ 4b. review fixes (v1.36.0) ======================================
{
  // token: tab-only by default, persisted only with "remember"
  const { QBR, store, sstore } = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(QBR);
  QBR.remoteFlagsToken.set("tabtok");
  ok(sstore["qbr-gh-pat"] === "tabtok" && !("qbr-gh-pat" in store), "token is kept for this tab only by default (sessionStorage)");
  ok(QBR.remoteFlagsToken.present() && !QBR.remoteFlagsToken.remembered(), "tab token counts as present, not remembered");
  QBR.remoteFlagsToken.set("keep", true);
  ok(store["qbr-gh-pat"] === "keep" && !("qbr-gh-pat" in sstore), "'Remember on this machine' moves it to localStorage");
  QBR.remoteFlagsToken.clear();
  ok(!("qbr-gh-pat" in store) && !("qbr-gh-pat" in sstore), "Forget clears both stores");
}
{
  // default branch follows the host
  const cases = [["rmt--qbr-dashboard-development.netlify.app", "rmt"], ["develop--qbr-dashboard-development.netlify.app", "develop"],
    ["qbr-dashboard-development.netlify.app", "develop"], ["hayleyllujah12.github.io", "main"], ["x.azurestaticapps.net", "main"]];
  let good = true, bad = "";
  for (const [h, b] of cases) {
    const { QBR } = freshEnv({ hostname: h, fetch: async () => { throw new Error("offline"); } });
    await settled(QBR);
    if (QBR.remoteFlagsRepo.get().branch !== b) { good = false; bad += h + "→" + QBR.remoteFlagsRepo.get().branch + " "; }
  }
  ok(good, "publish branch defaults to the host's branch (preview → its branch, Netlify main → develop, Pages/Azure → main)", bad);
}
{
  // publish merges onto the file on GitHub: unregistered keys kept, version from the file (stale cache)
  let putBody = null;
  const onGh = { version: 9, updated: "2026-10-10", flags: { inventory: true, "future.flag": false, "inventory.scan": true } };
  const { QBR } = freshEnv({
    fetch: async (url, init) => {
      if ((init.method || "GET") === "GET" && /api\.github\.com/.test(url))
        return jsonRes({ sha: "s9", content: Buffer.from(JSON.stringify(onGh)).toString("base64").replace(/(.{60})/g, "$1\n") });
      if ((init.method || "GET") === "GET") throw new Error("offline");
      putBody = JSON.parse(init.body);
      return { ok: true, status: 200, json: async () => ({}), text: async () => "" };
    },
    localStorage: { "qbr-gh-pat": "t",
      "qbr-remote-flags": JSON.stringify({ savedAt: 1, data: { version: 3, flags: {} } }) },
  });
  await settled(QBR);
  await QBR.remoteFlagsPublish({ "inventory.scan": false });
  const d = JSON.parse(Buffer.from(putBody.content, "base64").toString("utf8"));
  ok(d.version === 10, "version bumps from the file on GitHub, not a stale cached copy (" + d.version + ")");
  ok(d.flags.inventory === true && d.flags["future.flag"] === false && d.flags["inventory.scan"] === false,
    "keys this build doesn't register survive a publish " + JSON.stringify(d.flags));
}

// ============ 5. nav enforcement ==============================================
{
  const els = {};
  const mk = key => ({ style: {}, clicks: 0, click() { this.clicks++; }, closest() { return null; } });
  const { QBR, test } = freshEnv({
    fetch: async () => jsonRes({ version: 1, flags: { "inventory.scan": false } }),
    querySelectorAll: sel => {
      const m = /\[data-feature="([^"]+)"\]/.exec(sel);
      if (!m) return [];
      if (!els[m[1]]) els[m[1]] = mk(m[1]);
      return [els[m[1]]];
    },
    querySelector: () => null,
  });
  await settled(QBR);
  QBR.features.register({ module: "inventory", key: "inventory.scan", label: "Scan" });
  QBR.features.register({ module: "inventory", key: "inventory.supplies", label: "Supplies" });
  QBR.remoteFlagsApplyNav();
  ok(els["inventory.scan"].style.display === "none", "data-feature element hidden when flag off");
  ok(els["inventory.supplies"].style.display === "", "data-feature element shown when flag on");

  // capture-phase click guard
  const guard = global.document._listeners["click:cap"];
  ok(typeof guard === "function", "click guard attached in capture phase");
  let prevented = 0, stopped = 0;
  const evtFor = key => ({
    target: { closest: () => ({ getAttribute: () => key }) },
    preventDefault: () => { prevented++; },
    stopPropagation: () => { stopped++; },
  });
  guard(evtFor("inventory.scan"));
  ok(prevented === 1 && stopped === 1, "click on disabled feature is blocked");
  prevented = 0; stopped = 0;
  guard(evtFor("inventory.supplies"));
  ok(prevented === 0 && stopped === 0, "click on enabled feature passes through");
  guard({ target: null, preventDefault: () => { prevented++; }, stopPropagation: () => {} });
  ok(prevented === 0, "guard tolerates null target");
  ok(test.parentOf("inventory.supplies") === "inventory", "parentOf splits child keys");
  ok(test.parentOf("inventory") === null, "parentOf returns null for top-level keys");
  ok(typeof test.b64("✓ unicode") === "string", "b64 handles unicode");
}

// ============ 6. inventory.js pilot helpers ===================================
{
  global.window = global;
  if (!global.document) global.document = { readyState: "complete", addEventListener() {}, dispatchEvent() { return true; }, querySelectorAll: () => [], querySelector: () => null };
  if (!global.localStorage) global.localStorage = { getItem: () => null, setItem() {}, removeItem() {} };
  if (!global.fetch) global.fetch = async () => { throw new Error("offline"); };
  delete require.cache[require.resolve(MOD)];
  delete require.cache[require.resolve(INV)];
  const SCAN = path.resolve(__dirname, "../js/scan.js");
  delete require.cache[require.resolve(SCAN)];
  require(MOD);
  const inv = require(INV);
  let scan = null;
  try { scan = require(SCAN); } catch (e) { scan = null; }
  const QBR = global.QBR;
  ok(typeof inv.invEnabledViews === "function" && typeof inv.invFeatureOn === "function", "pilot helpers exposed");
  ok(eq(inv.invEnabledViews(), ["assets", "tickets", "supplies"]), "all sub-views enabled by default");
  QBR._remoteFlagsTestLoad({ version: 1, flags: { "inventory.supplies": false } }, "live");
  ok(eq(inv.invEnabledViews(), ["assets", "tickets"]), "supplies off → dropped from enabled views");
  ok(inv.invFeatureOn("inventory.supplies") === false, "invFeatureOn reflects remote");
  QBR._remoteFlagsTestLoad({ version: 1, flags: { "inventory": false, "inventory.assets": true } }, "live");
  ok(eq(inv.invEnabledViews(), []), "parent off → no enabled views");
  const keys = QBR.features.list().map(f => f.key).sort();
  ok(eq(keys, ["inventory.assets", "inventory.scan", "inventory.supplies", "inventory.tickets"]), "all four pilot sub-features registered");
}

console.log("\n" + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
})().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
