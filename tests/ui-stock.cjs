// v1.33.0 Inventory "remaining stock" quick tiles + filter (headless Chromium, file://, synthetic workbook).
// Usage: PW=<dir with node_modules/playwright> node tests/ui-stock.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const FIX = path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", FIX); await p.waitForTimeout(2500);
  await p.evaluate(() => { goToTab("dash-inventory"); QBR._invUI.view = "assets"; renderInventory(); }); await p.waitForTimeout(300);

  const tiles = await p.evaluate(() => [...document.querySelectorAll("[data-inv-stocktile]")].map(t => ({
    k: t.dataset.invStocktile, n: +t.querySelector(".kpi-val").textContent.replace(/[^\d]/g, ""), lbl: t.querySelector(".kpi-lbl").textContent.trim() })));
  console.log("== tiles");
  ok(tiles.length >= 2 && tiles[tiles.length - 1].k === "ALL", "category tiles + an 'All in stock' tile last  " + JSON.stringify(tiles.map(t => t.k + ":" + t.n)));
  const sumCats = tiles.filter(t => t.k !== "ALL").reduce((n, t) => n + t.n, 0);
  ok(tiles[tiles.length - 1].n === sumCats, "'All in stock' = sum of the category tiles (" + sumCats + ")");

  const listState = () => p.evaluate(() => {
    const ui = QBR._invUI;
    const g = [...document.querySelectorAll(".inv-stocktbl tbody tr.inv-mgrp")];
    return { type: ui.type, status: ui.status, stock: !!ui.stockView,
      groups: g.length, sum: g.reduce((n, r) => n + +r.cells[2].textContent.replace(/[^\d]/g, ""), 0),
      chip: (document.querySelector(".inv-stockbar .inv-chip") || {}).textContent || "",
      title: (document.querySelector("#inv-card-assets h6") || {}).textContent || "",
      on: [...document.querySelectorAll(".inv-stocktile.on")].map(t => t.dataset.invStocktile),
      typeOpts: [...document.querySelectorAll("#inv-f-type option")].map(o => o.value) };
  });

  console.log("== click a category tile");
  const first = tiles[0];
  await p.click(`[data-inv-stocktile="${first.k}"]`); await p.waitForTimeout(300);
  let s = await listState();
  ok(s.stock && s.type === first.k && s.status === "In Stock", "tile sets Type = " + first.k + " and Status = In Stock");
  ok(s.sum === first.n && s.groups > 0, `grouped-by-model list adds up to the tile (${s.sum} = ${first.n}, ${s.groups} models)`);
  ok(/Remaining stock: /.test(s.chip) && /Remaining stock/.test(s.title), "chip + card title say 'Remaining stock'");
  ok(JSON.stringify(s.on) === JSON.stringify([first.k]), "the clicked tile is highlighted");
  const sorted = await p.evaluate(() => { const v = [...document.querySelectorAll(".inv-stocktbl tr.inv-mgrp")].map(r => +r.cells[2].textContent.replace(/[^\d]/g, "")); return v.every((x, i) => !i || v[i - 1] >= x); });
  ok(sorted, "models sorted by most in stock first");
  ok(await p.evaluate(() => !document.querySelector(".inv-stocktbl th.th-sort")), "grouped table is not click-sortable (keeps each model with its serials)");

  console.log("== expand a model");
  const ex = await p.evaluate(() => {
    const r = document.querySelector(".inv-stocktbl tr.inv-mgrp"); r.click();
    const d = document.getElementById("inv-mgrp-" + r.dataset.invMgrp);
    return { open: !d.classList.contains("d-none") && r.getAttribute("aria-expanded") === "true",
      n: d.querySelectorAll(".inv-sn").length, want: +r.cells[2].textContent.replace(/[^\d]/g, "") };
  });
  ok(ex.open && ex.n === ex.want, `clicking a model lists its ${ex.want} serial(s)`);
  await p.keyboard.press("Tab");
  ok(await p.evaluate(() => { const r = document.querySelector(".inv-stocktbl tr.inv-mgrp"); r.focus(); r.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); return r.getAttribute("aria-expanded") === "false"; }), "Enter on a focused model row toggles it (keyboard)");

  console.log("== Units view");
  await p.click('[data-inv-stockmode="units"]'); await p.waitForTimeout(250);
  const u = await p.evaluate(() => {
    const rows = [...document.querySelectorAll("#inv-card-assets tr.inv-row")];
    return { n: rows.length, allStock: rows.every(r => /In Stock/.test(r.textContent)), chip: !!document.querySelector(".inv-stockbar .inv-chip") };
  });
  ok(u.n === Math.min(first.n, 150) && u.allStock && u.chip, `Units view: ${u.n} rows, all In Stock, chip kept`);
  await p.click('[data-inv-stockmode="model"]'); await p.waitForTimeout(250);
  ok((await listState()).groups > 0, "back to By model");

  console.log("== All in stock tile");
  await p.click('[data-inv-stocktile="ALL"]'); await p.waitForTimeout(300);
  s = await listState();
  ok(s.type === "ALL" && s.status === "In Stock" && s.sum === sumCats, `All in stock lists every category (${s.sum})`);

  console.log("== other filters are cleared so the list matches the tile");
  await p.evaluate(() => { QBR._invUI.q = "ZZZ"; QBR._invUI.client = "Nobody"; renderInventory(); }); await p.waitForTimeout(200);
  await p.click(`[data-inv-stocktile="${first.k}"]`); await p.waitForTimeout(300);
  s = await listState();
  ok(s.sum === first.n, "a stale serial/client filter doesn't hide stock after clicking a tile");

  console.log("== clearing");
  await p.click(`[data-inv-stocktile="${first.k}"]`); await p.waitForTimeout(250);
  s = await listState();
  ok(!s.stock && s.type === "ALL" && s.status === "ALL" && !s.chip, "clicking the active tile again returns to all assets");
  await p.click(`[data-inv-stocktile="${first.k}"]`); await p.waitForTimeout(250);
  await p.click("#inv-stock-x"); await p.waitForTimeout(250);
  s = await listState();
  ok(!s.stock && s.status === "ALL" && /^Assets/.test(s.title.trim()), "✕ on the chip clears it");
  await p.click(`[data-inv-stocktile="${first.k}"]`); await p.waitForTimeout(250);
  await p.selectOption("#inv-f-status", "Deployed"); await p.waitForTimeout(250);
  s = await listState();
  ok(!s.stock && !s.chip, "changing Status away from In Stock leaves the stock view");

  console.log("== Type filter uses the workbook's categories");
  ok(s.typeOpts[0] === "ALL" && ["Laptop", "Desktop", "Monitor"].every(t => s.typeOpts.includes(t)), "Type options: All + Laptop/Desktop/Monitor");
  const extra = await p.evaluate(() => [...document.querySelectorAll("[data-inv-stocktile]")].map(t => t.dataset.invStocktile).filter(k => k !== "ALL" && !["Laptop","Desktop","Monitor"].includes(k)));
  ok(extra.every(k => s.typeOpts.includes(k)), "any extra category with a tile is also in the Type filter  " + JSON.stringify(extra));

  console.log("== print / dark");
  await p.evaluate(() => { document.body.setAttribute("data-theme", "dark"); renderInventory(); });
  ok(await p.evaluate(() => !!document.querySelector("[data-inv-stocktile]")), "renders in dark mode");
  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
