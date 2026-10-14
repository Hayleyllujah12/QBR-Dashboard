// Remote feature flags (v1.36.0), headless Chromium.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-remote-flags.cjs
// Serves the repo over HTTP (the flags loader fetches config/features.json,
// which fails under file://) and exercises the Admin panel's Remote flags
// section plus nav enforcement. Never hits api.github.com.
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), http = require("http");
const ROOT = path.resolve(__dirname, "..");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const vis = (p, sel) => p.evaluate(s => { const el = document.querySelector(s); return !!el && el.getClientRects().length > 0 && getComputedStyle(el).display !== "none"; }, sel);
const MIME = { ".html": "text/html", ".js": "text/javascript", ".json": "application/json", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png" };

function serve() {
  return new Promise(resolve => {
    const srv = http.createServer((req, res) => {
      let u = decodeURIComponent(req.url.split("?")[0]);
      if (u === "/") u = "/index.html";
      const f = path.join(ROOT, u);
      fs.readFile(f, (e, d) => {
        if (e) { res.writeHead(404); res.end(); return; }
        res.writeHead(200, { "Content-Type": MIME[path.extname(f)] || "application/octet-stream" });
        res.end(d);
      });
    });
    srv.listen(0, "127.0.0.1", () => resolve(srv));
  });
}

(async () => {
  const srv = await serve();
  const url = "http://127.0.0.1:" + srv.address().port + "/index.html";
  const b = await chromium.launch();
  console.log("== remote flags UI");
  const ctx = await b.newContext();
  const p = await ctx.newPage();
  p._errs = []; p.on("pageerror", e => p._errs.push(e.message));
  await p.goto(url); await p.waitForTimeout(1200);
  // The admin panel lives in #app-body, which is only shown once a workbook is loaded.
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx")]); await p.waitForTimeout(2000);

  ok(p._errs.length === 0, "no page errors on load" + (p._errs.length ? " (" + p._errs.join("; ") + ")" : ""));
  ok(await p.evaluate(() => QBR.remoteFlagsStatus().source) === "live", "flags loaded from the live site");
  ok(await p.evaluate(() => QBR.features.list().map(f => f.key).sort().join(",")) ===
    "inventory.assets,inventory.scan,inventory.supplies,inventory.tickets", "4 pilot sub-features registered");
  ok(await p.evaluate(() => QBR.adminIsOn("inventory.supplies")) === true, "adminIsOn true by default");
  ok(await p.evaluate(() => document.querySelector('[data-tab="dash-scan"]').getAttribute("data-feature")) === "inventory.scan",
    "Scan nav button carries data-feature");

  // Admin panel: double-click version label, set password.
  await p.dblclick("#app-version"); await p.waitForTimeout(250);
  await p.fill("#admin-pw-1", "t3st"); await p.fill("#admin-pw-2", "t3st");
  await p.click("#admin-pw-ok"); await p.waitForTimeout(700);
  ok(await vis(p, "#admin-body .admin-remote"), "Remote flags section renders in the Admin panel");
  ok(await p.$$eval("#admin-body [data-rf-toggle]", x => x.length) === 4, "4 remote toggles listed");
  ok(await p.$$eval("#admin-body .admin-remote .badge", x => x.length) === 4, "each toggle has a REMOTE badge");
  ok(await p.$$eval("#admin-body [data-rf-toggle]:disabled", x => x.length) === 4, "toggles disabled without a token on this machine");
  ok(/This tab can publish: No/.test(await p.textContent("#admin-body .admin-remote")), "publish indicator says No without token");
  ok(await p.evaluate(() => { const b = document.querySelector("#rf-publish"); return b && b.disabled; }), "Publish button disabled without token");
  ok(/only on your own machine/.test(await p.textContent("#admin-body .admin-remote")), "own-machine token warning shown");

  // Simulate a remote flip (no network publish involved).
  await p.evaluate(() => {
    QBR._remoteFlagsTestLoad({ version: 9, updated: "2026-10-14", flags: { "inventory.scan": false } }, "live");
    QBR.remoteFlagsApplyNav();
  });
  await p.waitForTimeout(150);
  ok(!(await vis(p, '[data-tab="dash-scan"]')), "Scan nav button hidden when its flag is off");
  const before = await p.evaluate(() => APP.activeTab);
  await p.evaluate(() => document.querySelector('[data-tab="dash-scan"]').click());
  await p.waitForTimeout(150);
  ok(await p.evaluate(() => APP.activeTab) === before, "click guard blocks the disabled feature's tab switch");

  // Token flow: enabling publish on this machine only (fake token, never published).
  const TOK = "github_pat_" + "A1b2C3d4E5".repeat(4);
  await p.evaluate(async t => { await QBR.remoteFlagsToken.set(t); QBR.renderAdmin(); }, TOK);
  await p.waitForTimeout(200);
  ok(/This tab can publish: Yes/.test(await p.textContent("#admin-body .admin-remote")), "publish indicator says Yes with token");
  ok(await p.evaluate(() => { const b = document.querySelector("#rf-publish"); return b && !b.disabled; }), "Publish button enabled with token");
  ok(await p.$$eval("#admin-body [data-rf-toggle]:disabled", x => x.length) === 0, "toggles enabled with token");
  await p.evaluate(() => { QBR.remoteFlagsToken.clear(); QBR.renderAdmin(); });
  await p.waitForTimeout(200);
  ok(/This tab can publish: No/.test(await p.textContent("#admin-body .admin-remote")), "forgetting the token returns to read-only");

  console.log("== review fixes (v1.36.0)");
  ok(/Target main.*pull request/.test(await p.textContent("#rf-target")) && /Open pull request/.test(await p.textContent("#rf-publish")),
    "127.0.0.1 → target main → publish opens a pull request (button says so)");
  // classic token refused in the UI, never stored
  await p.fill("#rf-pat", "ghp_" + "x".repeat(36)); await p.click("#rf-pat-save"); await p.waitForTimeout(200);
  ok(/classic token/.test(await p.textContent("#rf-msg")) && await p.evaluate(() => !QBR.remoteFlagsToken.present() && !sessionStorage.getItem("qbr-gh-pat") && !localStorage.getItem("qbr-gh-pat")),
    "classic token rejected with an explanation, nothing stored");
  ok(await p.evaluate(() => !document.querySelector("#rf-pat-remember").disabled), "'Remember (encrypted)' offered on the https/localhost site");
  // Save via the form with "Remember": passphrase modal → encrypted, token checked against a mocked GitHub
  const exp = new Date(Date.now() + 30 * 86400000).toISOString().replace("T", " ").replace(/\..*/, " UTC");
  await p.route(/api\.github\.com/, route => route.fulfill({ status: 200, contentType: "application/json",
    headers: { "access-control-allow-origin": "*", "access-control-expose-headers": "github-authentication-token-expiration", "github-authentication-token-expiration": exp },
    body: JSON.stringify({ permissions: { push: true } }) }));
  await p.fill("#rf-pat", TOK); await p.check("#rf-pat-remember"); await p.click("#rf-pat-save"); await p.waitForTimeout(200);
  await p.fill("#admin-pw-1", "correct horse"); await p.fill("#admin-pw-2", "correct horse"); await p.click("#admin-pw-ok"); await p.waitForTimeout(2000);
  ok(/encrypted on this machine/.test(await p.textContent("#rf-msg")) && /expires in 30 days/.test(await p.textContent("#rf-msg")),
    "Save + Remember: passphrase asked, saved encrypted, expiry reported  [" + (await p.textContent("#rf-msg")).trim() + "]");
  ok(await p.evaluate(() => !!localStorage.getItem("qbr-gh-pat-enc") && QBR.remoteFlagsToken.present()), "…and the tab can publish");
  await p.unroute(/api\.github\.com/);
  await p.evaluate(() => { QBR.remoteFlagsToken.clear(); QBR.renderAdmin(); });
  // encrypted + locked state
  await p.evaluate(async t => { await QBR.remoteFlagsToken.set(t, { remember: true, passphrase: "correct horse" }); QBR.remoteFlagsToken.lock(); QBR.renderAdmin(); }, TOK);
  await p.waitForTimeout(150);
  ok(await vis(p, "#rf-unlock") && /locked/.test(await p.textContent("#admin-body .admin-remote")) && await p.evaluate(() => document.querySelector("#rf-publish").disabled),
    "saved token shows as locked: Unlock button, Publish disabled");
  ok(await p.evaluate(t => (localStorage.getItem("qbr-gh-pat-enc") || "").indexOf(t) < 0, TOK), "nothing readable in localStorage");
  await p.click("#rf-unlock"); await p.waitForTimeout(150);
  await p.fill("#admin-pw-1", "correct horse"); await p.click("#admin-pw-ok"); await p.waitForTimeout(1500);
  ok(await p.evaluate(() => QBR.remoteFlagsToken.present() && !document.querySelector("#rf-publish").disabled), "Unlock with the passphrase enables Publish");
  await p.evaluate(() => { QBR.remoteFlagsToken.clear(); QBR.renderAdmin(); });
  // remote parent "inventory": false hides the whole Inventory sidebar group and locks the local toggle
  await p.evaluate(() => {
    QBR._remoteFlagsTestLoad({ version: 10, flags: { inventory: false } }, "live");
    QBR.adminApplyFlags(); QBR.renderAdmin();
  });
  await p.waitForTimeout(150);
  ok(!(await vis(p, '[data-tab="dash-inventory"]')), "remote 'inventory': false hides the Inventory group (remote wins over local ON)");
  ok(await p.evaluate(() => { const b = document.querySelector('[data-admin-toggle="inventory"]'); return b && b.disabled && /OFF/.test(b.textContent); }),
    "local Inventory module toggle shows OFF + locked (REMOTE)");
  ok(await p.evaluate(() => QBR.adminIsModuleOn("inventory")) === false, "adminIsModuleOn follows the remote value");
  // late arrival: working copy follows a remote load that lands after the panel opened
  await p.evaluate(() => {
    QBR._rfUI.working = null; QBR._rfUI.dirty = false;
    QBR._remoteFlagsTestLoad({ version: 11, flags: { "inventory.supplies": false } }, "live");
    document.dispatchEvent(new CustomEvent("qbr:flags-ready"));
  });
  await p.waitForTimeout(150);
  ok(await p.evaluate(() => QBR._rfUI.working && QBR._rfUI.working["inventory.supplies"] === false),
    "admin working copy refreshes when remote flags arrive after the panel opened");
  ok(await p.evaluate(() => /OFF/.test(document.querySelector('[data-rf-toggle="inventory.supplies"]').textContent)), "…and the toggle redraws OFF");
  ok(p._errs.length === 0, "no page errors" + (p._errs.length ? " (" + p._errs.join("; ") + ")" : ""));

  await b.close(); srv.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
