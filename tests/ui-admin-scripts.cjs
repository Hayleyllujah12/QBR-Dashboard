// 2026-10-09 (v1.34.0 beta patch) Admin › PowerShell scripts (Mars's feature-admin-scripts + feature-scripts-pastesafe), headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-admin-scripts.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const FX = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 } });
  await ctx.grantPermissions(["clipboard-read", "clipboard-write"]);
  const p = await ctx.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", FX); await p.waitForTimeout(2000);

  console.log("== paste-safe built-in scripts");
  const ps = await p.evaluate(() => QBR.scriptList().filter(s => s.builtin).map(s => ({ t: s.title, blk: /<#|#>/.test(s.body),
    hdr: s.body.split("\n").slice(0, 30).filter(l => l.trim() && !/^\s*#/.test(l)).length,
    // every line before the first executable statement must be a # comment
    firstCode: (s.body.split("\n").find(l => l.trim() && !/^\s*#/.test(l)) || "").trim().slice(0, 40) })));
  ok(ps.length === 2, "2 built-in scripts");
  ok(ps.every(s => !s.blk), "no <# … #> block comments left (paste-safe)  " + JSON.stringify(ps.map(s => s.t + ":" + s.blk)));
  ok(ps.every(s => /^(Write-Host|Set-StrictMode)/.test(s.firstCode)), "first non-comment line is real code, so a partial paste can't run doc text  " + JSON.stringify(ps.map(s => s.firstCode)));
  const st = await p.evaluate(() => QBR.scriptList().find(s => /Storage/.test(s.title)).body);
  ok(/TB of \{1:N2\} GB used/.test(st), "storage USAGE format is 'X TB of Y GB used' (Mars's 2026-10-09 change)");

  console.log("== wizard no longer has 'Manage scripts'");
  const wiz = await p.evaluate(() => {
    goToTab("dash-audit");
    const w = QBR._auditWiz || (QBR._audit && QBR._audit.wiz);
    return { src: !!document.querySelector('[data-sact-wiz="managescripts"]') };
  });
  ok(!wiz.src, "no Manage scripts button rendered in the audit page");

  console.log("== Admin panel section");
  await p.dblclick("#app-version"); await p.waitForTimeout(200);
  await p.fill("#admin-pw-1", "s3cret"); await p.fill("#admin-pw-2", "s3cret"); await p.click("#admin-pw-ok"); await p.waitForTimeout(400);
  const ad = await p.evaluate(() => {
    const host = document.getElementById("admin-scripts-host");
    const order = [...document.querySelectorAll("#admin-body h4")].map(h => h.textContent.trim()).filter(t => /^Admin/.test(t));
    return { has: !!host, cards: host ? host.querySelectorAll(".qbr-scripts-card").length : 0,
      close: !!(host && host.querySelector('[data-sact="close"]')), gdap: !!(host && host.querySelector('[data-sfield="gdapfile"]')),
      add: !!(host && host.querySelector('[data-sact="add"]')), overlay: !!document.querySelector(".qbr-scripts-overlay"), order };
  });
  ok(ad.has && ad.cards === 2, "Admin shows 'PowerShell scripts' with the 2 built-in cards");
  ok(/PowerShell scripts/.test(ad.order[ad.order.length - 1] || ""), "section order: flags → theme → PowerShell scripts  " + JSON.stringify(ad.order));
  ok(ad.add && ad.gdap && !ad.close && !ad.overlay, "inline: Add + GDAP mapping, no Close button, no pop-up");

  console.log("== search / add / edit / delete / copy inline");
  await p.fill('#admin-scripts-host [data-sfield="q"]', "usage"); await p.waitForTimeout(150);
  ok(await p.evaluate(() => document.querySelectorAll("#admin-scripts-host .qbr-scripts-card").length === 1), "search filters the cards");
  await p.fill('#admin-scripts-host [data-sfield="q"]', ""); await p.waitForTimeout(100);
  await p.click('#admin-scripts-host [data-sact="add"]'); await p.waitForTimeout(150);
  const fields = await p.evaluate(() => [...document.querySelectorAll("#admin-scripts-host .qbr-scripts-form [data-sfield]")].map(e => e.dataset.sfield));
  const fill = async (f, v) => { const s = `#admin-scripts-host .qbr-scripts-form [data-sfield="${f}"]`; if (await p.$(s)) await p.fill(s, v); };
  await fill("title", "Test Script ZZ"); await fill("tags", "zz, test"); await fill("desc", "temp"); await fill("body", "Write-Host 'hi'");
  await p.click('#admin-scripts-host [data-sact="save"]'); await p.waitForTimeout(200);
  ok(await p.evaluate(() => QBR.scriptList().some(s => s.title === "Test Script ZZ") && document.querySelectorAll("#admin-scripts-host .qbr-scripts-card").length === 3), "add a custom script  " + JSON.stringify(fields));
  // admin re-render (theme change) must NOT wipe an open edit form
  await p.click('#admin-scripts-host [data-sact="add"]'); await p.waitForTimeout(100);
  await fill("title", "half typed");
  await p.evaluate(() => QBR.renderAdmin()); await p.waitForTimeout(150);
  ok(await p.evaluate(() => { const i = document.querySelector('#admin-scripts-host .qbr-scripts-form [data-sfield="title"]'); return !!i && i.value === "half typed"; }), "a panel re-render (flag/theme change) keeps a half-typed script");
  await p.evaluate(() => { const c = document.querySelector('#admin-scripts-host [data-sact="cancel"]'); if (c) c.click(); });
  p.once("dialog", d => d.accept());
  await p.evaluate(() => { const card = [...document.querySelectorAll("#admin-scripts-host .qbr-scripts-card")].find(c => /Test Script ZZ/.test(c.textContent)); card.querySelector('[data-sact="del"]').click(); });
  await p.waitForTimeout(200);
  ok(await p.evaluate(() => !QBR.scriptList().some(s => s.title === "Test Script ZZ")), "delete a custom script");
  await p.evaluate(() => { const card = [...document.querySelectorAll("#admin-scripts-host .qbr-scripts-card")].find(c => /Storage/.test(c.textContent)); card.querySelector('[data-sact="copy"]').click(); });
  await p.waitForTimeout(250);
  ok(/Bulk Extract M365 Storage Report/.test(await p.evaluate(() => navigator.clipboard.readText().catch(() => ""))), "Copy puts the storage script on the clipboard");

  console.log("== pop-up still works (gdap pill) and hands back to the inline view");
  await p.evaluate(() => QBR.scriptsOpen()); await p.waitForTimeout(150);
  ok(await p.evaluate(() => !!document.querySelector(".qbr-scripts-overlay .qbr-scripts-card")), "QBR.scriptsOpen() pop-up still opens (gdap pill)");
  await p.keyboard.press("Escape"); await p.evaluate(() => { const c = document.querySelector('.qbr-scripts-overlay [data-sact="close"]'); if (c) c.click(); }); await p.waitForTimeout(150);
  ok(await p.evaluate(() => !document.querySelector(".qbr-scripts-overlay") && !!document.getElementById("admin-scripts-host")), "closing the pop-up leaves the Admin section in place");
  await p.fill('#admin-scripts-host [data-sfield="q"]', "storage"); await p.waitForTimeout(150);
  ok(await p.evaluate(() => document.querySelectorAll("#admin-scripts-host .qbr-scripts-card").length === 1), "Admin search still works after the pop-up closed");

  console.log("== dark mode legibility");
  await p.evaluate(() => { document.body.setAttribute("data-theme", "dark"); });
  const dk = await p.evaluate(() => { const c = document.querySelector("#admin-scripts-host .qbr-scripts-card"); const cs = getComputedStyle(c);
    return { bg: cs.backgroundColor, fg: cs.color, page: getComputedStyle(document.body).backgroundColor }; });
  ok(dk.bg !== "rgb(255, 255, 255)" && dk.fg !== "rgb(34, 34, 34)", "cards follow the dark theme  " + JSON.stringify(dk));

  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
