// Admin panel (feature flags, password gate) + environment badge (v1.31.0), headless Chromium.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-admin.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const WB = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const vis = (p, sel) => p.evaluate(s => { const el = document.querySelector(s); return !!el && el.getClientRects().length > 0 && getComputedStyle(el).display !== "none"; }, sel);
const shortcut = p => p.keyboard.press("Control+Alt+Shift+KeyA");

async function boot(ctx, url) {
  const p = await ctx.newPage(); p._errs = []; p.on("pageerror", e => p._errs.push(e.message));
  await p.goto(url); await p.waitForTimeout(400);
  return p;
}
async function load(p) { await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} }); await p.setInputFiles("#file-input", WB); await p.waitForTimeout(2000); }

(async () => {
  const b = await chromium.launch();
  console.log("== admin panel");
  {
    const ctx = await b.newContext(); const p = await boot(ctx, "file://" + path.join(APPDIR, "index.html")); await load(p);
    ok(!(await vis(p, ".sb-group.sb-admin")), "Admin section hidden by default");
    ok(await p.evaluate(() => { const s = document.getElementById("dash-admin"); return s && s.parentElement.id === "app-body"; }), "dash-admin is a direct child of #app-body (not nested in Audit)");
    ok(await p.evaluate(() => { const u = document.querySelector('[data-tab="dash-admin"] use'); return !!document.querySelector(u.getAttribute("href")); }), "Admin nav icon exists in the sprite");
    await shortcut(p); await p.waitForTimeout(200);
    ok(await p.evaluate(() => { const i = [...document.querySelectorAll(".admin-pw-overlay input")]; return i.length === 2 && i.every(x => x.type === "password"); }), "first use: masked 'set password' modal (2 fields)");
    await p.fill("#admin-pw-1", "s3cret"); await p.fill("#admin-pw-2", "nope"); await p.click("#admin-pw-ok");
    ok(/did not match/.test(await p.textContent("#admin-pw-err")), "mismatch is reported");
    await p.fill("#admin-pw-2", "s3cret"); await p.click("#admin-pw-ok"); await p.waitForTimeout(300);
    ok(await vis(p, ".sb-group.sb-admin") && await p.evaluate(() => APP.activeTab === "dash-admin"), "after setting the password the Admin section shows and opens");
    const stored = await p.evaluate(() => localStorage.getItem("qbr-admin-pw"));
    ok(stored && !/s3cret/.test(stored) && /^[0-9a-f]{64}$/.test(stored), "only a SHA-256 hash is stored");
    ok(await p.$$eval("#admin-body [data-admin-toggle]", x => x.length) === 8, "8 module toggles listed");
    await p.click('[data-admin-toggle="inventory"]'); await p.waitForTimeout(150);
    ok(!(await vis(p, '[data-module="inventory"]')), "turning Inventory OFF hides its sidebar group");
    await p.click('.sb-item[data-tab="dash-risky"]'); await p.waitForTimeout(100);
    await p.evaluate(() => QBR.adminSetFlag("security", false)); await p.waitForTimeout(150);
    ok(await p.evaluate(() => APP.activeTab !== "dash-risky"), "disabling the active page's module moves to a visible page");
    await shortcut(p); await p.waitForTimeout(150);
    ok(!(await vis(p, ".sb-group.sb-admin")), "Ctrl+Shift+A again hides the Admin section");
    await p.reload(); await p.waitForTimeout(400); await load(p);
    ok(!(await vis(p, '[data-module="inventory"]')) && !(await vis(p, '[data-module="security"]')), "flags persist after reload");
    await shortcut(p); await p.waitForTimeout(200);
    ok(await p.$$eval(".admin-pw-overlay input", x => x.length) === 1, "next time: single password prompt");
    await p.fill("#admin-pw-1", "wrong"); await p.click("#admin-pw-ok"); await p.waitForTimeout(200);
    ok(!(await vis(p, ".sb-group.sb-admin")), "wrong password → stays hidden");
    await shortcut(p); await p.waitForTimeout(200); await p.fill("#admin-pw-1", "s3cret"); await p.keyboard.press("Enter"); await p.waitForTimeout(300);
    ok(await vis(p, ".sb-group.sb-admin"), "right password (Enter) → Admin shown");
    await shortcut(p); await p.waitForTimeout(150);
    ok(!(await vis(p, ".sb-group.sb-admin")), "Ctrl+Alt+Shift+A hides it again");
    await p.keyboard.press("Control+Shift+KeyA"); await p.waitForTimeout(150);
    ok(!(await p.$(".admin-pw-overlay")), "Ctrl+Shift+A (browser tab search) no longer triggers the panel");
    await p.evaluate(() => { if (document.activeElement) document.activeElement.blur(); }); await p.keyboard.type("rctadmin"); await p.waitForTimeout(200);
    ok(await p.$$eval(".admin-pw-overlay input", x => x.length) === 1, "typing 'rctadmin' outside a text box opens the password prompt");
    await p.fill("#admin-pw-1", "s3cret"); await p.keyboard.press("Enter"); await p.waitForTimeout(300);
    ok(await vis(p, ".sb-group.sb-admin"), "…and the right password shows Admin");
    await p.evaluate(() => { const i = document.createElement("input"); i.id = "tmp-in"; document.body.appendChild(i); });
    await shortcut(p); await p.waitForTimeout(150); await p.focus("#tmp-in"); await p.keyboard.type("rctadmin"); await p.waitForTimeout(200);
    ok(!(await p.$(".admin-pw-overlay")), "typing 'rctadmin' inside a text box does nothing");
    await shortcut(p); await p.waitForTimeout(200); await p.fill("#admin-pw-1", "s3cret"); await p.keyboard.press("Enter"); await p.waitForTimeout(300);
    await p.click("#admin-reset"); await p.waitForTimeout(150);
    ok(await vis(p, '[data-module="inventory"]') && await vis(p, '[data-module="security"]'), "Reset to defaults turns every module back on");
    ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
    await ctx.close();
  }
  console.log("== appearance");
  {
    const ctx = await b.newContext(); const p = await boot(ctx, "file://" + path.join(APPDIR, "index.html")); await load(p);
    const fsz = (sel) => p.evaluate(s => { const e = document.querySelector(s); return e ? parseFloat(getComputedStyle(e).fontSize) : null; }, sel);
    await p.click('.sb-item[data-tab="dash-risky"]'); await p.waitForTimeout(150);
    const base = { body: await fsz("body"), h: await fsz("#dash-risky .dash-h"), td: await fsz("#dash-risky td") };
    ok(base.body === 14 && base.h === 20 && !(await p.evaluate(() => document.documentElement.hasAttribute("data-qbr-fs"))), "defaults unchanged (body 14px, page heading 20px)");
    await p.evaluate(() => QBR.adminThemeSave({ textScale: 120, headScale: 130 })); await p.waitForTimeout(100);
    const big = { body: await fsz("body"), h: await fsz("#dash-risky .dash-h"), td: await fsz("#dash-risky td") };
    ok(Math.abs(big.body - 16.8) < 0.05 && Math.abs(big.h - 26) < 0.05, "body text 120% → 16.8px, headings 130% → 26px");
    ok(base.td && big.td > base.td * 1.15, "table text scales too (" + base.td + " → " + big.td + "px)");
    await p.evaluate(() => { setTheme("dark"); QBR.adminThemeSave({ pageBg: "#123456", headerText: "#ffcc00" }); }); await p.waitForTimeout(150);
    const col = await p.evaluate(() => ({ bg: getComputedStyle(document.body).backgroundColor, h: getComputedStyle(document.querySelector("#dash-risky .dash-h")).color }));
    ok(col.bg === "rgb(18, 52, 86)", "custom page background applies in DARK mode too (" + col.bg + ")");
    ok(col.h === "rgb(255, 204, 0)", "heading colour applies to headings");
    await p.evaluate(() => { window.dispatchEvent(new Event("beforeprint")); });
    const pr = await p.evaluate(() => ({ fs: document.documentElement.hasAttribute("data-qbr-fs"), bg: document.body.style.getPropertyValue("--bg") }));
    ok(!pr.fs && !pr.bg, "printing/export uses the standard look");
    await p.evaluate(() => { window.dispatchEvent(new Event("afterprint")); });
    ok(await p.evaluate(() => document.documentElement.hasAttribute("data-qbr-fs") && !!document.body.style.getPropertyValue("--bg")), "custom look restored after printing");
    await p.reload(); await p.waitForTimeout(400); await load(p);
    ok(Math.abs((await fsz("body")) - 16.8) < 0.05 && await p.evaluate(() => getComputedStyle(document.body).backgroundColor) === "rgb(18, 52, 86)", "appearance persists after reload");
    // Admin UI
    await p.evaluate(() => { document.body.classList.add("show-admin"); document.querySelector('[data-tab="dash-admin"]').click(); }); await p.waitForTimeout(200);
    ok(await p.$$eval(".admin-theme input[type=color]", x => x.length) === 4 && await p.$$eval(".admin-theme input[type=range]", x => x.length) === 2, "Appearance section: 4 colour pickers + 2 size sliders");
    ok(await p.$$eval("[data-admin-toggle]", x => x.length) === 8, "module toggles still listed once");
    await p.click('[data-admin-toggle="audit"]'); await p.waitForTimeout(150);
    ok(await p.evaluate(() => QBR.adminFlags().audit === false), "one click toggles a module exactly once");
    await p.evaluate(() => QBR.adminThemeSave({ pageBg: "#ffffff", headerText: "#fefefe" })); await p.evaluate(() => QBR.renderAdmin()); await p.waitForTimeout(100);
    ok(/Hard to read/.test(await p.textContent("#admin-body")), "low-contrast colours show a 'Hard to read' warning");
    await p.click("#admin-theme-reset"); await p.waitForTimeout(150);
    ok((await fsz("body")) === 14 && !(await p.evaluate(() => document.body.style.getPropertyValue("--bg"))), "Reset appearance restores defaults");
    await p.click('[data-theme-mode="light"]'); await p.waitForTimeout(200);
    ok(await p.evaluate(() => document.body.getAttribute("data-theme") === "light"), "Theme buttons switch light/dark");
    ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
    await ctx.close();
  }
  { const r = require("child_process").spawnSync(process.execPath, [path.join(__dirname, "build-type-scale.cjs"), "--check"], { encoding: "utf8" });
    ok(r.status === 0, "css/type-scale.css is up to date with styles.css"); }
  console.log("== environment badge");
  {
    const ctx = await b.newContext();
    const badge = async url => { const p = await boot(ctx, url); const r = await p.evaluate(() => { const e = document.getElementById("dev-badge"); return { t: e ? e.textContent : null, c: e ? e.className : "", title: document.title }; }); await p.close(); return r; };
    let r = await badge("file://" + path.join(APPDIR, "index.html"));
    ok(r.t === null && !/^\[/.test(r.title), "live folder / plain URL: no badge");
    r = await badge("file://" + path.join(APPDIR, "index.html") + "?beta");
    ok(r.t === "Beta version" && /env-beta/.test(r.c) && /^\[BETA\]/.test(r.title), "?beta → 'Beta version' badge + [BETA] title");
    const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-env-")), DEV = path.join(TMP, "QBR-Dashboard_dev", "qbr-app");
    fs.mkdirSync(DEV, { recursive: true }); fs.copyFileSync(path.join(APPDIR, "index.html"), path.join(DEV, "index.html"));
    for (const d of ["js", "css", "libs", "assets"]) if (fs.existsSync(path.join(APPDIR, d))) fs.symlinkSync(path.join(APPDIR, d), path.join(DEV, d));
    r = await badge("file://" + path.join(DEV, "index.html"));
    ok(r.t === "Beta version", "QBR-Dashboard_dev test folder → 'Beta version'");
    // Netlify hosts, served from disk
    await ctx.route(/github\.io\//, route => { const u = new URL(route.request().url()); const rel = decodeURIComponent(u.pathname).replace(/^\/QBR-Dashboard\/?/, ""); const f = path.join(APPDIR, rel || "index.html"); route.fulfill(fs.existsSync(f) ? { path: f } : { status: 404, body: "" }); });
    await ctx.route(/netlify\.app\//, route => { const u = new URL(route.request().url()); const f = path.join(APPDIR, u.pathname === "/" ? "index.html" : decodeURIComponent(u.pathname)); route.fulfill(fs.existsSync(f) ? { path: f } : { status: 404, body: "" }); });
    r = await badge("https://develop--rct-opsdesk.netlify.app/");
    ok(r.t === "Beta version", "Netlify develop--<site> URL → 'Beta version'");
    r = await badge("https://feature-admin-panel--rct-opsdesk.netlify.app/");
    ok(r.t === "Preview · feature-admin-panel" && /env-preview/.test(r.c), "Netlify branch URL → 'Preview · <branch>'");
    r = await badge("https://deploy-preview-3--rct-opsdesk.netlify.app/");
    ok(/^Preview · deploy-preview-3/.test(r.t || ""), "Netlify deploy-preview URL → Preview");
    r = await badge("https://qbr-dashboard-development.netlify.app/");
    ok(r.t === "Beta version", "Netlify main address (beta host; live is GitHub Pages) → 'Beta version'");
    r = await badge("https://hayleyllujah12.github.io/QBR-Dashboard/");
    ok(r.t === null, "GitHub Pages (live) → no badge");
    try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) {}
    await ctx.close();
  }
  console.log("== Netlify Identity gate");
  {
    const FAKE = `(function(){var h={},u=localStorage.getItem("fakeUser");function f(e,x){(h[e]||[]).forEach(function(c){c(x)})}
      window.netlifyIdentity={on:function(e,c){(h[e]=h[e]||[]).push(c)},init:function(){setTimeout(function(){f("init",u?{email:u}:null)},10)},
      open:function(){window.__opened=1;localStorage.setItem("fakeUser","pedro@example.com");f("login",{email:"pedro@example.com"})},
      currentUser:function(){var x=localStorage.getItem("fakeUser");return x?{email:x}:null},close:function(){},logout:function(){localStorage.removeItem("fakeUser");f("logout")}};})();`;
    const mk = async (blockWidget) => {
      const ctx = await b.newContext(); const reqs = [];
      ctx.on("request", r => reqs.push(r.url()));
      await ctx.route(/netlify\.app\//, route => { const u = new URL(route.request().url()); const f = path.join(APPDIR, u.pathname === "/" ? "index.html" : decodeURIComponent(u.pathname)); route.fulfill(fs.existsSync(f) ? { path: f } : { status: 404, body: "" }); });
      await ctx.route(/identity\.netlify\.com/, route => blockWidget ? route.abort() : route.fulfill({ contentType: "application/javascript", body: FAKE }));
      return { ctx, reqs };
    };
    const hidden = p => p.evaluate(() => getComputedStyle(document.querySelector(".app-header")).visibility === "hidden" && getComputedStyle(document.querySelector("main")).visibility === "hidden");
    // offline / live: nothing loaded, no gate
    { const { ctx, reqs } = await mk(false); const p = await boot(ctx, "file://" + path.join(APPDIR, "index.html"));
      ok(!(await p.$("#identity-gate")) && !reqs.some(u => /identity\.netlify\.com/.test(u)), "file:// (live/offline): no gate and the Netlify widget is NOT downloaded");
      await ctx.close(); }
    { const { ctx } = await mk(false); const p = await boot(ctx, "https://develop--rct-opsdesk.netlify.app/"); await p.waitForTimeout(300);
      ok(await p.isVisible("#identity-gate") && await hidden(p), "Netlify beta, signed out: login gate shown, dashboard hidden");
      ok(await p.evaluate(() => +getComputedStyle(document.getElementById("identity-gate")).zIndex < 99), "gate sits below the Netlify login modal (widget iframe z-index 99)");
      await p.click("#identity-login-btn"); await p.waitForTimeout(200);
      ok(!(await p.isVisible("#identity-gate")) && !(await hidden(p)) && await p.isVisible("#identity-logout-btn"), "after login: dashboard shown + Log out button");
      await p.reload(); await p.waitForTimeout(400);
      ok(!(await p.isVisible("#identity-gate")), "stays signed in after reload");
      await p.click("#identity-logout-btn"); await p.waitForTimeout(200);
      ok(await p.isVisible("#identity-gate") && await hidden(p) && !(await p.$("#identity-logout-btn")), "Log out → gate back, dashboard hidden");
      ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
      await ctx.close(); }
    { const { ctx } = await mk(true); const p = await boot(ctx, "https://feature-admin-panel--rct-opsdesk.netlify.app/"); await p.waitForTimeout(500);
      ok(await p.isVisible("#identity-gate") && await hidden(p) && /unavailable/i.test(await p.textContent("#identity-note")), "widget blocked → stays locked with 'Login service unavailable'");
      await ctx.close(); }
  }
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
