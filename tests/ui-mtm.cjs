// 2026-10-11 Scan MTM (the contributor's feature-scan-mtm + review fixes) + tenant pre-fill convention. Headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-mtm.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const FIX = path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx");
const TRK = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [FIX, TRK]); await p.waitForTimeout(2500);

  console.log("== tenant pre-fill convention");
  const G = "aaaaaaaa-1111-2222-3333-444444444444";
  const tp = await p.evaluate(G => {
    const s = QBR.scriptCopyText("storage", G), u = QBR.scriptCopyText("usage", G);
    const sn = QBR.scriptCopyText("storage", null), un = QBR.scriptCopyText("usage", null);
    const bad = QBR.scriptCopyText("storage", 'x"; Remove-Item C:\\ -Recurse #');
    return {
      s: s.prefilled && s.text.includes('$InputTenantId = "' + G + '"') && !s.text.includes("QBR_TENANT_ID"),
      u: u.prefilled && u.text.includes('$defTenant = "' + G + '"') && !u.text.includes("QBR_TENANT_ID"),
      sn: !sn.prefilled && sn.text.includes('$InputTenantId = "QBR_TENANT_ID"') && /Read-Host "Enter Tenant ID \(GUID\)"/.test(sn.text),
      un: !un.prefilled && /RunState\.Contains\('Phase2'\)/.test(un.text),
      bad: !bad.prefilled && !bad.text.includes("Remove-Item C:"),
      fallbackHasNoLiteral: (sn.text.match(/QBR_TENANT_ID/g) || []).length === 1,
    };
  }, G);
  ok(tp.s, "storage: GUID swapped in, no placeholder left");
  ok(tp.u, "usage: GUID becomes the Read-Value default");
  ok(tp.sn, "storage, no GDAP match: placeholder kept + Read-Host fallback present");
  ok(tp.un, "usage, no GDAP match: previous default (RunState / Get-MgContext) kept as fallback");
  ok(tp.bad, "a non-GUID tenant value is never pasted into the script");
  ok(tp.fallbackHasNoLiteral, "fallback test doesn't compare against the literal placeholder (all occurrences get replaced)");

  console.log("== scan table");
  await p.click('.sb-item[data-tab="dash-scan"]'); await p.waitForTimeout(300);
  await p.evaluate(png => {
    const ui = QBR._scanUI;
    ui.rows.push({ id: ui.nextId++, file: null, dataUrl: png, fp: "t1", model: "ThinkPad E14", serial: "PFZZ0001", mtm: "21M4S5VD00",
      productKey: "", barcodeRaw: "PFZZ0001 | 21M4S5VD00 | 21M4S5VD00PFZZ0001", status: "done", dup: null });
    scanRenderRows();
  }, PNG);
  const tbl = await p.evaluate(() => {
    const t = document.getElementById("scan-tbody").closest("table");
    const th = [...t.querySelectorAll("thead th")].map(x => x.textContent.trim());
    const tr = t.querySelector("tbody tr");
    return { th, nth: th.length, ntd: tr ? tr.children.length : -1,
      mtmCol: tr ? (tr.children[th.indexOf("MTM")] || {}).innerHTML || "" : "" };
  });
  ok(tbl.nth === tbl.ntd, `header and row have the same column count (${tbl.nth} / ${tbl.ntd})`);
  ok(/data-f="mtm"/.test(tbl.mtmCol), "MTM header sits over the MTM input  " + JSON.stringify(tbl.th));

  console.log("== lightbox keeps a detected MTM");
  await p.evaluate(() => { const r = QBR._scanUI.rows[QBR._scanUI.rows.length - 1]; scanPreview(r.id); });
  await p.waitForTimeout(120);
  ok(await p.inputValue("#scan-lb-mtm") === "21M4S5VD00", "opening the photo shows the detected MTM");
  await p.evaluate(() => scanPreviewSave(true)); await p.waitForTimeout(100);
  ok(await p.evaluate(() => QBR._scanUI.rows[QBR._scanUI.rows.length - 1].mtm === "21M4S5VD00"), "saving the lightbox (serial fix) doesn't wipe the MTM");

  console.log("== batch tag dialog");
  await p.evaluate(() => { const r = QBR._scanUI.rows[QBR._scanUI.rows.length - 1]; r.sel = true; scanBatchTagDialog("deploy"); });
  await p.waitForTimeout(250);
  const bd = await p.evaluate(() => {
    const th = [...document.querySelectorAll("#sb-thead th")].map(x => x.textContent.trim());
    const tr = document.querySelector("#sb-body tr");
    return { th, ntd: tr ? tr.children.length : -1, mtm: tr ? tr.children[th.indexOf("MTM")].textContent.trim() : "" };
  });
  ok(bd.th.length === bd.ntd && bd.mtm === "21M4S5VD00", `batch review: MTM column lines up (${bd.th.length}/${bd.ntd}, "${bd.mtm}")`);
  const cl = await p.evaluate(() => { const s = document.getElementById("sb-client"); if (s) { if (s.tagName === "SELECT") s.selectedIndex = Math.min(1, s.options.length - 1); else s.value = "Test School"; } const d = document.getElementById("sb-date"); if (d && !d.value) d.value = "2026-10-11"; const g = document.getElementById("sb-go"); if (g) g.click(); return !!g; });
  await p.waitForTimeout(300);
  const added = await p.evaluate(() => (APP.model.inventory.assets.find(a => a.sn === "PFZZ0001") || {}).mtm || null);
  ok(cl && added === "21M4S5VD00", "batch Add registers the unit with its MTM (" + added + ")");

  console.log("== intake form + Excel");
  await p.evaluate(() => { goToTab("dash-inventory"); QBR.invRegisterSerial("PFZZ0002"); }); await p.waitForTimeout(250);
  await p.fill("#in-mtm", "21m4s5vd01"); await p.click("#in-go"); await p.waitForTimeout(250);
  const a2 = await p.evaluate(() => (APP.model.inventory.assets.find(a => a.sn === "PFZZ0002") || {}).mtm);
  ok(a2 === "21M4S5VD01", "intake stores MTM (upper-cased)");
  const ex = await p.evaluate(() => {
    const wb = QBR.invBuildWorkbook().wb; const ws = wb.Sheets[wb.SheetNames.find(n => /DEVICES/i.test(n))];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const h = rows[0], r = rows.find(x => x[0] === "PFZZ0002");
    return { lens: rows.every(x => x.length <= h.length), mtm: r[h.indexOf("MTM")], model: r[h.indexOf("Model")] };
  });
  ok(ex.mtm === "21M4S5VD01" && ex.lens, "Export: MTM column after Model, rows aligned with the header");
  const b64 = fs.readFileSync(FIX).toString("base64");
  const sv = await p.evaluate(b64 => {
    const store = QBR.journalEntries(); const fp = Object.keys(store).find(k => (store[k].ops || []).some(o => o.op === "invIntake"));
    const wb = XLSX.read(Uint8Array.from(atob(b64), c => c.charCodeAt(0)), { type: "array", cellDates: true });
    const rep = QBR.patchWorkbookFromJournal("assets", wb, fp);
    const ws = wb.Sheets[wb.SheetNames.find(n => /DEVICES/i.test(n))];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const h = rows[0].map(String), r = rows.find(x => String(x[0]) === "PFZZ0002") || [];
    return { fp: !!fp, ok: rep.ok, hasCol: h.includes("MTM"), v: r[h.indexOf("MTM")] };
  }, b64);
  ok(sv.fp && sv.ok && sv.hasCol && sv.v === "21M4S5VD01", "Save to Excel (journal patch) writes an MTM column into the linked workbook  " + JSON.stringify(sv));

  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
