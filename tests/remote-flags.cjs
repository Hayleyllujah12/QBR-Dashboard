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
  const sstore = Object.assign({}, opts.sessionStorage || {});
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

// ============ 4. publish + token security (hardened 2026-10-14) ===============
const TOK = "github_pat_" + "A1b2C3d4E5".repeat(4);
const sess = t => ({ "qbr-gh-pat": JSON.stringify({ t: t || TOK, at: Date.now() }) });
function ghMock(o) {
  o = o || {};
  const calls = [], files = Object.assign({}, o.files || {});   // branch -> doc
  const res = (status, data, hdr) => ({ ok: status < 300, status, json: async () => data, text: async () => JSON.stringify(data),
    headers: { get: k => (hdr || {})[String(k).toLowerCase()] || null } });
  const handler = async (url, init) => {
    const m = (init.method || "GET"), u = String(url).replace("https://api.github.com", "");
    const body = init.body ? JSON.parse(init.body) : null;
    calls.push({ m, u, body, auth: init.headers && init.headers.Authorization });
    if (!/^https:\/\/api\.github\.com\//.test(String(url))) return res(200, {});        // loader's own fetch etc.
    let r;
    if (m === "GET" && (r = /^\/repos\/[^/]+\/[^/]+$/.exec(u))) return o.repoStatus ? res(o.repoStatus, { message: "Bad credentials" }) : res(200, { permissions: { push: true } }, o.hdr || {});
    if (m === "GET" && (r = /\/contents\/config\/features\.json\?ref=(.+)$/.exec(u))) {
      const d = files[decodeURIComponent(r[1])];
      return d ? res(200, { sha: "sha-" + r[1], content: Buffer.from(JSON.stringify(d)).toString("base64").replace(/(.{60})/g, "$1\n") }) : res(404, { message: "Not Found" });
    }
    if (m === "GET" && (r = /\/git\/ref\/heads\/(.+)$/.exec(u))) return res(200, { object: { sha: "base-" + r[1] } });
    if (m === "POST" && /\/git\/refs$/.test(u)) return res(201, { ref: body.ref });
    if (m === "PUT" && /\/contents\/config\/features\.json$/.test(u)) { files[body.branch] = JSON.parse(Buffer.from(body.content, "base64").toString("utf8")); return res(o.putStatus || 200, {}); }
    if (m === "POST" && /\/pulls$/.test(u)) return res(201, { number: 7, html_url: "https://github.com/o/r/pull/7" });
    return res(404, { message: "unexpected " + m + " " + u });
  };
  return { calls, files, handler };
}
{
  // classic / malformed tokens refused; fine-grained accepted, tab-only by default
  const { QBR, store, sstore } = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(QBR);
  let e1 = null; try { await QBR.remoteFlagsToken.set("ghp_" + "x".repeat(36)); } catch (e) { e1 = e; }
  ok(e1 && /classic token/.test(e1.message), "classic ghp_ token is refused");
  let e2 = null; try { await QBR.remoteFlagsToken.set("hello"); } catch (e) { e2 = e; }
  ok(e2 && /fine-grained/.test(e2.message), "non-GitHub text is refused");
  await QBR.remoteFlagsToken.set(TOK);
  ok(QBR.remoteFlagsToken.present() && !("qbr-gh-pat" in store) && !("qbr-gh-pat-enc" in store), "fine-grained token kept for this tab only (nothing in localStorage)");
  ok(JSON.parse(sstore["qbr-gh-pat"]).t === TOK, "tab copy stores the token with a last-used time");
  QBR.remoteFlagsToken.clear();
  ok(!QBR.remoteFlagsToken.present() && !("qbr-gh-pat" in sstore), "Forget clears the tab copy");
}
{
  // idle timeout
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("offline"); },
    sessionStorage: { "qbr-gh-pat": JSON.stringify({ t: TOK, at: Date.now() - 31 * 60000 }) } });
  await settled(QBR);
  ok(!QBR.remoteFlagsToken.present(), "a tab token unused for 30+ min is dropped (idle lock)");
}
{
  // plain-text token from an earlier build is purged on load
  const { QBR, store } = freshEnv({ fetch: async () => { throw new Error("offline"); }, localStorage: { "qbr-gh-pat": TOK } });
  await settled(QBR);
  ok(!("qbr-gh-pat" in store) && QBR.remoteFlagsToken.purgedPlain() && !QBR.remoteFlagsToken.present(), "plain-text token left in localStorage is deleted on load");
}
{
  // remember = encrypted at rest; locked after reload; unlock needs the passphrase
  const env1 = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(env1.QBR);
  let e0 = null; try { await env1.QBR.remoteFlagsToken.set(TOK, { remember: true, passphrase: "short" }); } catch (e) { e0 = e; }
  ok(e0 && /8 characters/.test(e0.message), "remember needs a passphrase of 8+ characters");
  await env1.QBR.remoteFlagsToken.set(TOK, { remember: true, passphrase: "correct horse" });
  const enc = env1.store["qbr-gh-pat-enc"];
  ok(enc && enc.indexOf(TOK) < 0 && enc.indexOf("github_pat_") < 0, "remembered token is stored encrypted (no plain text at rest)");
  const env2 = freshEnv({ fetch: async () => { throw new Error("offline"); }, localStorage: { "qbr-gh-pat-enc": enc } });
  await settled(env2.QBR);
  ok(env2.QBR.remoteFlagsToken.locked() && !env2.QBR.remoteFlagsToken.present(), "after a reload the saved token is locked");
  let e3 = null; try { await env2.QBR.remoteFlagsToken.unlock("wrong pass!"); } catch (e) { e3 = e; }
  ok(e3 && /Wrong passphrase/.test(e3.message) && !env2.QBR.remoteFlagsToken.present(), "wrong passphrase doesn't unlock");
  await env2.QBR.remoteFlagsToken.unlock("correct horse");
  ok(env2.QBR.remoteFlagsToken.present() && JSON.parse(env2.sstore["qbr-gh-pat"]).t === TOK, "right passphrase unlocks it for this tab");
  env2.QBR.remoteFlagsToken.lock();
  ok(env2.QBR.remoteFlagsToken.locked(), "Lock now drops the tab copy, keeps the encrypted one");
}
{
  // publish refuses without a token / when locked
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("should not be called"); } });
  await settled(QBR);
  let err = null;
  try { await QBR.remoteFlagsPublish({}); } catch (e) { err = e; }
  ok(err && /No GitHub token/.test(err.message), "publish without a token throws");
}
{
  // main → pull request flow; nothing written to main
  const gm = ghMock({ files: { main: { version: 9, updated: "2026-10-10", flags: { inventory: true, "future.flag": false, "inventory.scan": true } } } });
  const { QBR, fetchCalls } = freshEnv({ fetch: gm.handler, sessionStorage: sess(),
    localStorage: { "qbr-gh-repo": JSON.stringify({ owner: "Hayleyllujah12", repo: "QBR-Dashboard", branch: "main" }),
      "qbr-remote-flags": JSON.stringify({ savedAt: 1, data: { version: 3, flags: {} } }) } });
  await settled(QBR);
  gm.calls.length = 0;
  const r = await QBR.remoteFlagsPublish({ "inventory.scan": false });
  const seq = gm.calls.map(c => c.m + " " + c.u.replace(/^\/repos\/[^/]+\/[^/]+/, "").replace(/\?.*$/, "")).join(" | ");
  ok(/GET \/contents\/config\/features\.json \| GET \/git\/ref\/heads\/main \| POST \/git\/refs \| PUT \/contents\/config\/features\.json \| POST \/pulls/.test(seq), "main: read → branch → commit on branch → open PR  (" + seq + ")");
  const put = gm.calls.find(c => c.m === "PUT");
  ok(put && /^flags\/\d{8}-\d{6}-v10$/.test(put.body.branch) && put.body.sha === "sha-main", "commit goes to a new flags/<stamp> branch, not main");
  ok(gm.files.main.version === 9, "main itself is untouched until the PR is merged");
  const pr = gm.calls.find(c => c.u.endsWith("/pulls"));
  ok(pr.body.base === "main" && /inventory\.scan`: on → \*\*off\*\*/.test(pr.body.body), "PR targets main and lists the change");
  ok(r.pr && r.pr.number === 7 && r.pr.url === "https://github.com/o/r/pull/7" && !r.direct, "publish resolves with the PR link");
  const d = gm.files[put.body.branch];
  ok(d.version === 10 && d.flags.inventory === true && d.flags["future.flag"] === false && d.flags["inventory.scan"] === false, "merged onto the file on GitHub: version 9→10, unknown keys kept");
  ok(QBR.remoteFlagsStatus().remote.version === 3, "local flags don't change before the PR is merged");
  ok(gm.calls.every(c => c.auth === "Bearer " + TOK), "token sent as Bearer, only to api.github.com");
  ok(fetchCalls.every(c => !/github/.test(c.url) || /^https:\/\/api\.github\.com\//.test(c.url)), "no other GitHub host is contacted");
}
{
  // develop → PR too; preview branch → direct commit
  const gm = ghMock({});
  const { QBR } = freshEnv({ fetch: gm.handler, sessionStorage: sess(), localStorage: { "qbr-gh-repo": JSON.stringify({ owner: "o", repo: "r", branch: "develop" }) } });
  await settled(QBR);
  const r1 = await QBR.remoteFlagsPublish({ x: true });
  ok(r1.pr && !gm.files.develop, "develop: pull request, not a direct commit");
  QBR.remoteFlagsRepo.set({ owner: "o", repo: "r", branch: "rmt" });
  gm.calls.length = 0;
  const r2 = await QBR.remoteFlagsPublish({ x: false });
  const put = gm.calls.find(c => c.m === "PUT");
  ok(r2.direct && put.body.branch === "rmt" && !("sha" in put.body) && gm.files.rmt.version === 1 && !gm.calls.some(c => /pulls|git\/refs/.test(c.u)), "preview branch: direct commit (first publish creates the file, v1)");
  ok(QBR.remoteFlagsStatus().source === "live" && QBR.remoteFlagsStatus().remote.flags.x === false, "direct publish refreshes the local flags");
}
{
  // repo / branch names can't inject paths
  const gm = ghMock({});
  const { QBR } = freshEnv({ fetch: gm.handler, sessionStorage: sess(), localStorage: { "qbr-gh-repo": JSON.stringify({ owner: "o/../evil", repo: "r", branch: "main" }) } });
  await settled(QBR);
  let err = null; try { await QBR.remoteFlagsPublish({}); } catch (e) { err = e; }
  ok(err && /not configured/.test(err.message) && gm.calls.filter(c => /api\.github/.test("https://api.github.com" + c.u) && c.u.startsWith("/repos")).length === 0, "owner/repo with path characters is refused before any call");
}
{
  // token check: expiry + bad token
  const gm = ghMock({ hdr: { "github-authentication-token-expiration": new Date(Date.now() + 20 * 86400000).toISOString().replace("T", " ").replace(/\..*/, " UTC") } });
  const { QBR } = freshEnv({ fetch: gm.handler, sessionStorage: sess() });
  await settled(QBR);
  const c = await QBR.remoteFlagsCheckToken();
  ok(c.ok && c.daysLeft >= 19 && c.daysLeft <= 20, "token check reads the expiry (" + c.daysLeft + " days)");
  const gm2 = ghMock({ repoStatus: 401 });
  const env = freshEnv({ fetch: gm2.handler, sessionStorage: sess() });
  await settled(env.QBR);
  let err = null; try { await env.QBR.remoteFlagsCheckToken(); } catch (e) { err = e; }
  ok(err && /invalid or expired/.test(err.message), "token check explains a 401");
}
{
  // repo defaults + default branch follows the host
  const { QBR } = freshEnv({ fetch: async () => { throw new Error("offline"); } });
  await settled(QBR);
  const repo = QBR.remoteFlagsRepo.get();
  ok(repo.owner === "Hayleyllujah12" && repo.repo === "QBR-Dashboard" && repo.branch === "main", "repo defaults to Robin's repo/main on the live host");
  const cases = [["rmt--qbr-dashboard-development.netlify.app", "rmt"], ["develop--qbr-dashboard-development.netlify.app", "develop"],
    ["qbr-dashboard-development.netlify.app", "develop"], ["hayleyllujah12.github.io", "main"], ["x.azurestaticapps.net", "main"]];
  let good = true, bad = "";
  for (const [h, b] of cases) {
    const e = freshEnv({ hostname: h, fetch: async () => { throw new Error("offline"); } });
    await settled(e.QBR);
    if (e.QBR.remoteFlagsRepo.get().branch !== b) { good = false; bad += h + "→" + e.QBR.remoteFlagsRepo.get().branch + " "; }
  }
  ok(good, "publish branch defaults to the host's branch (preview → its branch, Netlify main → develop, Pages/Azure → main)", bad);
  ok(QBR.remoteFlagsIsProtected("main") && QBR.remoteFlagsIsProtected("develop") && !QBR.remoteFlagsIsProtected("rmt"), "main + develop are PR-only");
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
