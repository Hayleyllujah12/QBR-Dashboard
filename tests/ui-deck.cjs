// 2026-10-14 Export Deck fixes: order, QBR-only scope, export-clone label trimming + value labels. Headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-deck.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1440, height: 900 } })).newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(300);
  await p.evaluate(() => { localStorage.setItem("qbr-remember", "0"); window.print = () => { window.__printed = 1; }; });
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx"), path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx")]); await p.waitForTimeout(2500);
  await p.evaluate(() => exportDeck()); await p.waitForFunction(() => window.__printed, null, { timeout: 30000 });
  await p.emulateMedia({ media: "print" });
  const v = await p.evaluate(() => {
    const shown = [...document.querySelectorAll("#app-body > .dash-panel")].filter(x => getComputedStyle(x).display !== "none").map(x => x.id);
    return { shown };
  });
  ok(v.shown[0] === "dash-overview", "Overview is the first page after the cover (was printed last)  " + v.shown.slice(0, 3));
  ok(!v.shown.some(id => /inventory|scan|todo|gtickets|ticket360|asset360|audit|admin|soc|school|fulldata/.test(id)), "tools / raw data pages are not in the deck  " + JSON.stringify(v.shown));
  const out = path.join(os.tmpdir(), "qbr-deck-test.pdf");
  await p.pdf({ path: out, format: "A4", printBackground: true });
  ok(fs.statSync(out).size > 50000, "deck PDF renders");
  await p.emulateMedia({ media: "screen" });
  await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(300);
  ok(await p.evaluate(() => !document.querySelector(".deck-out")), "panels restored after printing");
  const pol = await p.evaluate(() => {
    const long = "SCHOOLS DIVISION OFFICE (SDO) OF TAGUIG CITY AND PATEROS";
    const bar = { type: "bar", data: { labels: [long, "B"], datasets: [{ data: [3, 0] }] } }, ob = { plugins: {}, scales: { x: { ticks: {} } } };
    _qbrExportPolish(bar, ob);
    const ln = { type: "line", data: { labels: ["J", "F"], datasets: [{ data: [5, 0] }, { data: [1, 2] }] } }, ol = { plugins: {} };
    _qbrExportPolish(ln, ol);
    const busy = { type: "line", data: { labels: Array(30).fill("x"), datasets: [{ data: Array(30).fill(1) }] } }, oy = { plugins: {} };
    _qbrExportPolish(busy, oy);
    const fake = { getLabelForValue: () => long };
    return { cb: ob.scales.x.ticks.callback.call(fake, 0), bl: ob.plugins.qbrLabels, ll: ol.plugins.qbrLabels, busy: oy.plugins.qbrLabels,
      cfgClean: Object.values(QBR._configs).every(c => !c || !c.options || !c.options.plugins || !c.options.plugins.qbrLabels || !c.options.plugins.qbrLabels.__auto) };
  });
  ok(pol.cb.length <= 22 && /…$/.test(pol.cb), "export clone trims long category labels  (" + pol.cb + ")");
  ok(pol.bl && pol.bl.mode === "vbar" && pol.bl.labels[0] === "3", "single-series bars get value labels in exports");
  ok(pol.ll && pol.ll.labels[0][1] === null && pol.ll.labels[1][1] === "2", "line charts: values on points, zeros skipped");
  ok(!pol.busy, "busy charts (30 points) left without labels");
  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
