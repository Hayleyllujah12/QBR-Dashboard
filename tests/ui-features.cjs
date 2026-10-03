// Headless checks for the v1.9.0 Phase 3 features: status lists, nav attention counts, KPI deltas,
// sortable tables, Tenant Health detail pane, print snapshots, School 360 (2026-10-03). Usage:
// PW=<dir containing node_modules/playwright> node tests/ui-features.cjs "<tracker>.xlsx" [shotsDir]
const { chromium } = require(process.env.PW + "/playwright");
const WB = process.argv[2];
let pass = 0, fail = 0; const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (x !== undefined ? "  [" + x + "]" : "")); };
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1440, height: 900 } }); const p = await ctx.newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message)); p.on("console", m => { if (m.type() === "error") errs.push(m.text()); });
  await p.addInitScript(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
  await p.goto("file://" + require("path").resolve(__dirname, (require("fs").existsSync(require("path").resolve(__dirname, "../qbr-app/index.html")) ? "../qbr-app/index.html" : "../index.html")));
  await p.setInputFiles("#file-input", WB);
  await p.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 60000 }); await p.waitForTimeout(800);
  console.log("-- status lists");
  const sl = await p.evaluate(() => {
    const a = document.querySelectorAll("#ov-secposture .stl-row"), h = document.querySelectorAll("#ov-health .stl-row");
    return { a: a.length, h: h.length, topA: Math.round(a[0].getBoundingClientRect().top), topH: Math.round(h[0].getBoundingClientRect().top),
      firstA: a[0].querySelector(".stl-lbl").textContent + " " + a[0].querySelector(".stl-num").textContent,
      tot: document.querySelector("#ov-secposture .stl-total").textContent, canv: !!document.getElementById("ch-ov-secposture") };
  });
  ok(sl.a === 4 && sl.h >= 5, "rows rendered", sl.a + "/" + sl.h); ok(sl.topA === sl.topH, "cards aligned row-for-row", sl.topA + "/" + sl.topH);
  ok(/Enabled 76/.test(sl.firstA), "ranked (largest first)", sl.firstA); ok(sl.canv, "hidden canvas kept for Export Images", sl.tot);
  await p.click('#ov-secposture .stl-row:has-text("Not Enabled")'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-sec" && document.getElementById("f-secdef").value === "DISABLED" && APP.filters.secDefault === "DISABLED"), "row drills into Security with filter preset");
  await p.selectOption("#f-secdef", "ALL"); await p.waitForTimeout(300);
  console.log("-- storage capacity + share of total");
  const cs = await p.evaluate(() => {
    const rows = [...document.querySelectorAll("#tbl-stor-cap .cap-row")];
    const pcts = rows.map(r => parseFloat(r.cells[5].textContent));
    const badge = document.querySelector('.sb-item[data-tab="dash-storage"] .sb-count');
    const kpiN = [...document.querySelectorAll("#kpi-storage .kpi")].find(k => /Near Capacity/.test(k.textContent));
    return { n: rows.length, allOver: pcts.every(v => v >= (QBR.THRESH.STORAGE_WARN || 80)), sorted: pcts.every((v, i) => !i || pcts[i - 1] >= v),
      badge: badge ? +badge.title.split(" ")[0] : 0, kpi: kpiN ? +kpiN.querySelector(".kpi-val").textContent : -1,
      riskyNote: document.getElementById("risky-top-note").textContent, shareCells: document.querySelectorAll("#tbl-risky td:nth-child(3)").length };
  });
  ok(cs.allOver && cs.sorted, "Capacity Watch lists only tenants ≥ warning level, worst first", cs.n);
  ok(cs.kpi === cs.n && (cs.n === 0 ? cs.badge === 0 : cs.badge === cs.n), "KPI, nav count and watchlist agree", cs.kpi + "/" + cs.badge + "/" + cs.n);
  ok(/Top 3 = \d+%/.test(cs.riskyNote) && cs.shareCells > 0, "risky share-of-total note + Share column", cs.riskyNote.slice(0, 60));
  console.log("-- nav badges");
  const nb = await p.evaluate(() => Object.fromEntries(Array.from(document.querySelectorAll(".sb-item .sb-count")).map(s => [s.closest(".sb-item").dataset.tab, s.textContent.replace(/[^\d]/g, "")])));
  ok(nb["dash-sec"] === "12" && +nb["dash-risky"] > 0, "attention counts", JSON.stringify(nb));
  console.log("-- KPI deltas");
  await p.click('.sb-item[data-tab="dash-risky"]'); await p.waitForTimeout(200);
  ok((await p.$$("#kpi-risky .kpi-delta")).length === 0, "no deltas on All quarters");
  await p.click('.qchip[data-q="Q2"]'); await p.waitForTimeout(800);
  const d = await p.evaluate(() => ({ risky: Array.from(document.querySelectorAll("#kpi-risky .kpi-delta")).map(x => x.textContent.trim()),
    sec: document.querySelectorAll("#kpi-sec .kpi-delta").length, canva: document.querySelectorAll("#kpi-canva .kpi-delta").length,
    usage: Array.from(document.querySelectorAll("#kpi-usage .kpi-delta")).map(x => x.textContent.trim()), detail: document.querySelectorAll("#kpi-td .kpi-delta, #kpi-sd .kpi-delta").length }));
  ok(d.risky.length >= 1 && /vs Q1/.test(d.risky[0]), "Risky page deltas vs Q1", d.risky.join(" | "));
  ok(d.sec === 0, "snapshot page (Security) has no deltas"); ok(d.canva >= 1, "Canva deltas", d.canva);
  ok(d.usage.length >= 1, "Usage deltas", d.usage.join(" | ")); ok(d.detail === 0, "per-tenant detail cards have no portfolio deltas");
  await p.click('.qchip[data-q="ALL"]'); await p.waitForTimeout(800);
  console.log("-- sortable");
  await p.click('.sb-item[data-tab="dash-sec"]'); await p.waitForTimeout(200);
  const th = await p.$('#dash-sec table th:has-text("Risky Users")');
  if (th) { await th.click(); await p.waitForTimeout(150);
    let r = await p.evaluate(() => { const t = [...document.querySelectorAll("#dash-sec table")].find(t => t.querySelector("th[aria-sort='descending'],th[aria-sort='ascending']")); const i = [...t.tHead.rows[0].cells].findIndex(c => c.getAttribute("aria-sort") !== "none" && c.classList.contains("th-sort")); const v = [...t.tBodies[0].rows].slice(0, 3).map(r => +r.cells[i].textContent.replace(/,/g, "")); return { dir: t.tHead.rows[0].cells[i].getAttribute("aria-sort"), v }; });
    ok(r.dir === "descending" && r.v[0] >= r.v[1], "numeric column sorts descending first", r.v.join(","));
    await th.click(); await p.waitForTimeout(150);
    r = await p.evaluate(() => { const t = [...document.querySelectorAll("#dash-sec table")].find(t => t.querySelector("th[aria-sort='ascending']")); return t ? "asc" : "none"; });
    ok(r === "asc", "second click ascending");
    const sch = await p.$('#dash-sec table th:has-text("School")'); await sch.click(); await p.waitForTimeout(150);
    const names = await p.evaluate(() => { const t = [...document.querySelectorAll("#dash-sec table")].find(t => t.querySelector("th[aria-sort='ascending']")); return [...t.tBodies[0].rows].slice(0, 3).map(r => r.cells[0].textContent.trim()); });
    ok(names[0].localeCompare(names[1]) <= 0, "text column A→Z", names.join(" / "));
    await p.selectOption("#f-authmethod", { index: 2 }); await p.waitForTimeout(400);
    const still = await p.evaluate(() => { const t = [...document.querySelectorAll("#dash-sec table")].find(t => t.querySelector("th[aria-sort='ascending']")); if (!t) return false; const n = [...t.tBodies[0].rows].map(r => r.cells[0].textContent.trim()); return n.every((x, i) => i === 0 || n[i - 1].localeCompare(x, undefined, { numeric: true }) <= 0); });
    ok(still, "sort survives a re-render (auth-method filter)");
    await p.selectOption("#f-authmethod", "ALL"); await p.waitForTimeout(300);
  } else ok(false, "found Risky Users header");
  console.log("-- THI detail pane");
  await p.click('.sb-item[data-tab="dash-overview"]'); await p.waitForTimeout(200);
  await p.click("#thi-drill-toggle"); await p.waitForTimeout(200);
  const nm = await p.$eval("#thi-drill-body .thi-row td:nth-child(2)", e => e.textContent);
  await p.click("#thi-drill-body .thi-row"); await p.waitForTimeout(400);
  const pane = await p.evaluate(() => ({ vis: !document.getElementById("detail-pane").hidden, title: document.getElementById("dp-title").textContent, dims: document.querySelectorAll(".dp-dim").length, focus: document.activeElement.id }));
  ok(pane.vis && pane.title === nm && pane.dims === 6, "pane opens with 6 dimensions", pane.title);
  ok(pane.focus === "dp-close", "focus moves into pane");
  await p.screenshot({ path: "" + (process.argv[3] || "/tmp") + "/p3_pane.png" });
  await p.keyboard.press("Escape"); await p.waitForTimeout(200);
  ok(await p.evaluate(() => document.getElementById("detail-pane").hidden && document.activeElement.classList.contains("thi-row")), "Escape closes, focus returns to row");
  await p.keyboard.press("Enter"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => !document.getElementById("detail-pane").hidden), "Enter on row opens pane");
  await p.click(".dp-link[data-dp-tab]"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.getElementById("detail-pane").hidden && APP.activeTab !== "dash-overview"), "dimension link navigates + closes", await p.evaluate(() => APP.activeTab));
  console.log("-- print snapshots");
  await p.click('.sb-item[data-tab="dash-overview"]'); await p.waitForTimeout(200);
  await p.click("#btn-export"); await p.click("#btn-pdf"); await p.waitForTimeout(300);
  const snap = await p.evaluate(() => ({ imgs: document.querySelectorAll("#dash-overview .print-chart").length, cls: document.body.classList.contains("print-snap") }));
  ok(snap.imgs >= 5 && snap.cls, "Export Tab swaps charts for print images", snap.imgs);
  await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.querySelectorAll(".print-chart").length === 0 && !document.body.classList.contains("print-snap")), "snapshots removed after print");
  console.log("-- Postmaster latest reputation (THI)");
  ok(await p.evaluate(() => { const t = QBR._latestScorableReputation([{ key: "X", reputation: "HIGH", monthIdx: 0 }, { key: "X", reputation: "Issues detected", monthIdx: 2 }, { key: "X", reputation: "Not enough data", monthIdx: 3 }]); return t.X && t.X.rep === "Issues detected" && t.X.mi === 2; }), "latest scorable month wins (HIGH Jan → Issues Mar = Issues)");
  console.log("-- School 360");
  await p.click('.sb-item[data-tab="dash-school"]'); await p.waitForTimeout(500);
  const s1 = await p.evaluate(() => ({ tab: APP.activeTab, crumb: document.getElementById("crumb-page").textContent, name: (document.querySelector("#s360-body .s360-name") || {}).textContent,
    kpis: document.querySelectorAll("#s360-body .kpi").length, cards: document.querySelectorAll("#s360-body .s360-card").length, dl: document.querySelectorAll("#dl-s360 option").length, nm: APP.model.master.size,
    zero: [...document.querySelectorAll("#s360-body .kpi-val")].some(v => v.textContent.trim() === "0" && /Storage|Usage|Reputation|Health/.test(v.closest(".kpi").textContent)) }));
  ok(s1.tab === "dash-school" && /School 360/.test(s1.crumb) && !!s1.name, "nav opens with a default school", s1.name);
  ok(s1.kpis === 6 && s1.cards >= 10 && s1.dl === s1.nm, "6 KPIs, all cards, picker lists every school", s1.kpis + "/" + s1.cards + "/" + s1.dl);
  const pick = await p.evaluate(() => { const ks = [...APP.model.master.keys()]; const k = ks.find(k => APP.model.storage.some(s => s.key === k) && APP.model.risky.some(r => r.key === k && r.risky > 0)); return APP.model.master.get(k).name; });
  await p.fill("#s360-school", pick); await p.dispatchEvent("#s360-school", "change"); await p.waitForTimeout(500);
  const s2 = await p.evaluate(() => ({ name: document.querySelector("#s360-body .s360-name").textContent, chart: !!(QBR._charts && QBR._charts["ch-s360-risky"]), acts: document.querySelectorAll("#s360-body .s360-act, #s360-body .s360-acts + p, #s360-body .s360-card p.s360-empty").length }));
  ok(s2.name === pick && s2.chart, "picker switches school + risky trend chart", s2.name);
  await p.fill("#s360-school", "zzz no such school"); await p.dispatchEvent("#s360-school", "change"); await p.waitForTimeout(300);
  ok(await p.evaluate(n => document.querySelector("#s360-body .s360-name").textContent === n && /No school named/i.test(document.getElementById("s360-note").textContent), pick), "unknown name keeps current school + shows a note");
  await p.click('#s360-body .s360-go[data-s360-tab="dash-storage"]'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-storage"), "card link opens the related page", await p.evaluate(() => APP.activeTab + " " + document.getElementById("s360-note").textContent));
  await p.click('.sb-item[data-tab="dash-school"]'); await p.waitForTimeout(300);
  await p.click("#btn-export"); await p.click("#btn-pdf"); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.body.dataset.print === "dash-school" && document.querySelectorAll("#dash-school .print-chart").length >= 1), "Export Tab prints this school with chart snapshot");
  await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(300);
  ok(await p.evaluate(() => { const g = collectChartImages(); return !g.some(x => x.items.some(i => i.id === "ch-s360-risky")); }), "School 360 excluded from Export Images gallery");
  await p.emulateMedia({ media: "print" });
  ok(await p.evaluate(() => { const el = document.getElementById("dash-school"), hid = el.classList.contains("d-none"); document.body.dataset.print = "ALL"; el.classList.remove("d-none"); const d = getComputedStyle(el).display; if (hid) el.classList.add("d-none"); delete document.body.dataset.print; return d === "none"; }), "School 360 hidden in Export Deck print");
  await p.emulateMedia({ media: "screen" });
  await p.setViewportSize({ width: 1440, height: 2000 }); await p.waitForTimeout(400);
  await p.screenshot({ path: "" + (process.argv[3] || "/tmp") + "/s360_light.png" });
  await p.click("#btn-theme"); await p.waitForTimeout(500);
  await p.screenshot({ path: "" + (process.argv[3] || "/tmp") + "/s360_dark.png" });
  await p.click("#btn-theme"); await p.waitForTimeout(300); await p.setViewportSize({ width: 1440, height: 900 }); await p.waitForTimeout(300);
  console.log("-- THI pane → School 360");
  await p.click('.sb-item[data-tab="dash-overview"]'); await p.waitForTimeout(300);
  if (!(await p.isVisible("#thi-drill-body .thi-row"))) { await p.click("#thi-drill-toggle"); await p.waitForTimeout(200); }
  const nm2 = await p.$eval("#thi-drill-body .thi-row td:nth-child(2)", e => e.textContent);
  await p.click("#thi-drill-body .thi-row"); await p.waitForTimeout(300);
  await p.click("#dp-s360"); await p.waitForTimeout(500);
  ok(await p.evaluate(n => APP.activeTab === "dash-school" && document.getElementById("detail-pane").hidden && document.querySelector("#s360-body .s360-name").textContent === n, nm2), "pane 'Open School 360' jumps to that school", nm2);
  await p.click('.sb-item[data-tab="dash-overview"]'); await p.waitForTimeout(300);
  await p.click("#btn-theme"); await p.waitForTimeout(500);
  await p.screenshot({ path: "" + (process.argv[3] || "/tmp") + "/p3_dark_overview.png" });
  await p.click("#btn-theme"); await p.waitForTimeout(500);
  await p.screenshot({ path: "" + (process.argv[3] || "/tmp") + "/p3_light_overview.png" });
  console.log("errors:", errs.length, errs.slice(0, 5).join(" | "));
  ok(errs.length === 0, "zero console errors");
  console.log(`RESULT: ${pass} passed, ${fail} failed`);
  await b.close();
})().catch(e => { console.error(e); process.exit(2); });
