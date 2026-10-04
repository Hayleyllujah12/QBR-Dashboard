// Headless UI smoke test for the v1.9.0 app shell (navigation, exports, THI, Data Quality,
// print, responsive, deep links). Needs Playwright + Chromium; not part of the zero-install harness.
// Usage: PW=<dir containing node_modules/playwright> node tests/ui-smoke.cjs "<tracker>.xlsx" [shotsDir]
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const APP = "file://" + path.resolve(__dirname, (require("fs").existsSync(path.resolve(__dirname, "../qbr-app/index.html")) ? "../qbr-app/index.html" : "../index.html"));
const WB = process.argv[2];
const OUT = process.argv[3] || "/tmp/shots";
require("fs").mkdirSync(OUT, { recursive: true });
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (x !== undefined ? "  [" + x + "]" : "")); };

(async () => {
  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" }).catch(() => chromium.launch());
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", e => errors.push("pageerror: " + e.message));
  page.on("console", m => { if (m.type() === "error") errors.push("console: " + m.text()); });
  await page.addInitScript(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
  await page.goto(APP);

  console.log("-- empty state");
  ok(await page.isVisible(".app-header"), "top bar visible");
  ok(!(await page.isVisible("#app-sidebar")), "sidebar hidden before data");
  ok((await page.textContent("#app-version")) === "v1.9.0", "version label from QBR.VERSION", await page.textContent("#app-version"));
  ok((await page.evaluate(() => document.querySelector(".app-header").offsetHeight)) === 48, "top bar is 48px");
  await page.screenshot({ path: OUT + "/01_empty.png" });

  console.log("-- upload");
  await page.setInputFiles("#file-input", WB);
  await page.waitForFunction(() => !document.getElementById("app-body").classList.contains("d-none"), null, { timeout: 60000 });
  await page.waitForTimeout(800);
  ok(await page.evaluate(() => document.body.classList.contains("has-data")), "body.has-data set");
  ok(await page.isVisible("#app-sidebar"), "sidebar visible after data");
  const tabs = await page.$$eval("#app-sidebar .sb-item[data-tab]", b => b.map(x => x.dataset.tab));
  // detail pages open from other pages and have no nav item by design (Asset 360 in the inventory build)
  const DETAIL_PAGES = ["dash-asset360"];
  const panels = (await page.$$eval(".dash-panel", p => p.map(x => x.id))).filter(id => !DETAIL_PAGES.includes(id));
  ok(tabs.length >= 15 && tabs.length === panels.length, "nav items: " + tabs.length + " (15 core + optional Inventory/Scan)", tabs.length);
  ok(tabs.every(t => panels.includes(t)) && panels.every(p => tabs.includes(p)), "nav items ↔ panels 1:1 (detail pages excluded)");
  ok((await page.$$("[data-tab]")).length === tabs.length, "only nav items carry data-tab");
  await page.screenshot({ path: OUT + "/02_overview_1440.png" });

  console.log("-- navigation");
  for (const t of tabs) {
    await page.click(`#app-sidebar .sb-item[data-tab="${t}"]`);
    await page.waitForTimeout(60);
    const st = await page.evaluate(t => {
      const vis = Array.from(document.querySelectorAll(".dash-panel")).filter(p => !p.classList.contains("d-none")).map(p => p.id);
      const b = document.querySelector(`.sb-item[data-tab="${t}"]`);
      return { vis, cur: b.getAttribute("aria-current"), act: b.classList.contains("active"), crumb: document.getElementById("crumb-page").textContent, lbl: b.querySelector(".sb-label").textContent, active: APP.activeTab, hash: location.hash };
    }, t);
    ok(st.vis.length === 1 && st.vis[0] === t && st.cur === "page" && st.act && st.crumb === st.lbl && st.active === t, "nav " + t, st.vis.join(",") + " · " + st.crumb + " · " + st.hash);
  }
  // goToTab path (overview quick-jump buttons)
  await page.click('.sb-item[data-tab="dash-overview"]');
  const qj = await page.$("#dash-overview [data-goto]");
  if (qj) { const target = await qj.getAttribute("data-goto"); await qj.click(); await page.waitForTimeout(60);
    ok(await page.evaluate(t => APP.activeTab === t && document.querySelector(`.sb-item[data-tab="${t}"]`).getAttribute("aria-current") === "page", target), "Overview quick-jump → sidebar state follows", target); }

  console.log("-- preserved features");
  await page.click('.sb-item[data-tab="dash-overview"]');
  const thi = (await page.textContent("#thi-score")).trim();
  ok(/^\d+$/.test(thi), "Tenant Health Index renders", thi);
  await page.click("#thi-drill-toggle"); await page.waitForTimeout(100);
  ok((await page.$$("#thi-drill tbody tr")).length > 0, "THI drill-down rows", (await page.$$("#thi-drill tbody tr")).length);
  await page.click('.sb-item[data-tab="dash-quality"]');
  ok((await page.evaluate(() => document.getElementById("dash-quality").innerText.length)) > 500, "Data Quality audit renders");
  const canv = await page.evaluate(() => Object.keys(QBR._charts || {}).length);
  ok(canv >= 30, "charts instantiated", canv);

  console.log("-- export menu");
  await page.click('.sb-item[data-tab="dash-overview"]');
  await page.click("#btn-export");
  ok(await page.isVisible("#export-menu"), "Export menu opens");
  ok((await page.$$eval("#export-menu .menu-item", m => m.map(x => x.id))).join() === "btn-pdf,btn-deck,btn-images,btn-data,btn-clean", "5 export buttons keep ids");
  await page.screenshot({ path: OUT + "/03_export_menu.png" });
  await page.keyboard.press("Escape");
  ok(!(await page.isVisible("#export-menu")), "Escape closes menu");
  for (const id of ["btn-data", "btn-clean"]) {
    await page.click("#btn-export");
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), page.click("#" + id)]);
    ok(!!dl, id + " downloads", dl.suggestedFilename());
    ok(!(await page.isVisible("#export-menu")), id + " closes menu after click");
  }
  await page.click("#btn-export"); await page.click("#btn-pdf"); await page.waitForTimeout(300);
  ok((await page.evaluate(() => window.__printed)) === 1, "Export Tab calls print");
  await page.click("#btn-export"); await page.click("#btn-deck");
  await page.waitForFunction(() => window.__printed >= 2, null, { timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(800);
  ok((await page.evaluate(() => window.__printed)) === 2, "Export Deck runs to print");
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await page.waitForTimeout(800); await page.waitForFunction(() => !document.getElementById("btn-deck").disabled, null, { timeout: 30000 }).catch(() => {}); ok((await page.textContent("#btn-deck")).trim() === "Export Deck" && !(await page.$eval("#btn-deck", b => b.disabled)), "Deck button restored", await page.textContent("#btn-deck"));
  const popP = ctx.waitForEvent("page", { timeout: 8000 }).catch(() => null);
  await page.click("#btn-export"); await page.click("#btn-images");
  const pop = await popP;
  await page.waitForTimeout(1500);
  const imgs = pop ? await pop.$$eval("img", i => i.length).catch(() => 0) : await page.$$eval("img", i => i.length);
  ok(imgs >= 20, "Export Images gallery", (pop ? "popup " : "overlay ") + imgs + " imgs");
  if (pop) await pop.close();
  ok((await page.textContent("#btn-images")).trim() === "Export Images", "Images button restored");

  console.log("-- print media");
  await page.emulateMedia({ media: "print" });
  const pr = await page.evaluate(() => ["#app-sidebar", ".app-header", ".page-head", ".filter-bar"].map(s => getComputedStyle(document.querySelector(s)).display));
  ok(pr.every(d => d === "none"), "print hides top bar, sidebar, page head, filters", pr.join(","));
  ok((await page.evaluate(() => getComputedStyle(document.getElementById("app-main")).marginLeft)) === "0px", "print resets content offset");
  await page.emulateMedia({ media: "screen" });

  console.log("-- rail + theme");
  await page.click("#sb-collapse"); await page.waitForTimeout(300);
  ok((await page.evaluate(() => document.getElementById("app-sidebar").offsetWidth)) === 48, "rail = 48px");
  await page.screenshot({ path: OUT + "/04_rail.png" });
  await page.click("#sb-collapse"); await page.waitForTimeout(300);
  ok((await page.evaluate(() => document.getElementById("app-sidebar").offsetWidth)) === 260, "expanded = 260px");
  await page.click("#btn-theme"); await page.waitForTimeout(600);
  ok((await page.getAttribute("body", "data-theme")) === "dark", "dark theme on");
  ok((await page.getAttribute("#btn-theme", "aria-label")) === "Switch to light mode", "theme button accessible name");
  await page.screenshot({ path: OUT + "/05_dark_overview.png" });
  await page.click('.sb-item[data-tab="dash-risky"]'); await page.waitForTimeout(200);
  await page.screenshot({ path: OUT + "/06_dark_risky.png" });
  await page.click("#btn-theme"); await page.waitForTimeout(400);

  console.log("-- responsive");
  for (const w of [375, 768, 1024, 1440]) {
    await page.setViewportSize({ width: w, height: 860 }); await page.waitForTimeout(250);
    for (const t of ["dash-overview", "dash-sec", "dash-fulldata"]) {
      await page.evaluate(t => { if (document.body.classList.contains("sb-open")) document.getElementById("sb-scrim").click(); document.querySelector(`.sb-item[data-tab="${t}"]`).click(); }, t);
      await page.waitForTimeout(150);
      const sw = await page.evaluate(() => [document.documentElement.scrollWidth, window.innerWidth]);
      ok(sw[0] <= sw[1], `no horizontal scroll @${w} ${t}`, sw.join("/"));
    }
    const hdr = await page.evaluate(() => document.querySelector(".app-header").offsetHeight);
    ok(hdr === 48, `top bar stays 48px @${w}`, hdr);
    await page.evaluate(() => document.querySelector('.sb-item[data-tab="dash-overview"]').click());
    await page.waitForTimeout(150);
    await page.screenshot({ path: OUT + `/07_w${w}.png` });
  }
  await page.setViewportSize({ width: 768, height: 860 }); await page.waitForTimeout(200);
  ok(await page.isVisible("#nav-toggle"), "menu button shown < 1024");
  await page.click("#nav-toggle"); await page.waitForTimeout(350);
  ok(await page.evaluate(() => document.body.classList.contains("sb-open") && !document.getElementById("sb-scrim").hidden), "drawer opens with scrim");
  await page.screenshot({ path: OUT + "/08_drawer_768.png" });
  await page.keyboard.press("Escape"); await page.waitForTimeout(300);
  ok(await page.evaluate(() => !document.body.classList.contains("sb-open")), "Escape closes drawer");
  await page.click("#nav-toggle"); await page.waitForTimeout(300);
  await page.click('.sb-item[data-tab="dash-storage"]'); await page.waitForTimeout(300);
  ok(await page.evaluate(() => APP.activeTab === "dash-storage" && !document.body.classList.contains("sb-open")), "drawer item navigates + closes");

  console.log("-- deep link + restore");
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.goto("about:blank");
  await page.goto(APP + "#risky");
  await page.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(1200);
  const dl = await page.evaluate(() => [document.body.classList.contains("has-data"), typeof APP !== "undefined" && APP.activeTab]);
  ok(dl[0] && dl[1] === "dash-risky", "#risky deep link opens Risky sign-ins after cached restore", dl.join(","));

  console.log("\n-- errors: " + errors.length);
  errors.slice(0, 10).forEach(e => console.log("   " + e));
  ok(errors.length === 0, "zero console / page errors");
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  await browser.close();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
