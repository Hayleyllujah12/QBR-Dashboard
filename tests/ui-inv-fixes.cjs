// Inventory fixes (2026-10-05): unregistered serials (ticket form check, "Not in inventory" tag,
// #asset not-found page) + Scan batch review / stocktake. Headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-inv-fixes.cjs [shotsDir]
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const FIX = path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx");
const SHOTS = process.argv[2] || null;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const shot = async (p, n) => { if (SHOTS) { fs.mkdirSync(SHOTS, { recursive: true }); await p.screenshot({ path: path.join(SHOTS, n + ".png") }); } };

async function load(b) {
  const p = await b.newPage({ viewport: { width: 1366, height: 900 } }); p._errs = [];
  p.on("pageerror", e => p._errs.push(e.message));
  p.on("dialog", d => d.accept());
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", FIX); await p.waitForTimeout(2500);
  return p;
}

async function serialFixes(b) {
  console.log("== unregistered serials");
  const p = await load(b);
  await p.evaluate(() => { goToTab("dash-inventory"); QBR._invUI.view = "assets"; renderInventory(); }); await p.waitForTimeout(300);
  const cnt = () => p.evaluate(() => document.querySelectorAll("#inv-col-assets tbody tr.inv-row").length);
  const all0 = await cnt(); await p.selectOption("#inv-f-status", { index: 1 }); await p.waitForTimeout(150);
  await p.selectOption("#inv-f-status", { index: 0 }); await p.waitForTimeout(150);
  ok((await cnt()) === all0, "choosing 'All statuses' again restores the list (was empty before the fix)");
  const sn = await p.evaluate(() => invModel().assets[0].sn);
  const typo = sn.slice(0, 2) + "-" + sn.slice(2);
  await p.evaluate(() => { goToTab("dash-inventory"); QBR._invUI.view = "tickets"; QBR._invUI.form = "ticket"; renderInventory(); });
  await p.waitForTimeout(300);
  ok(await p.isVisible("#tk-sn"), "New ticket form open");
  await p.fill("#tk-sn", typo); await p.waitForTimeout(150);
  const hint = await p.textContent("#tk-sn-hint");
  ok(/Not in inventory/.test(hint) && hint.includes(sn), `hint flags ${typo} and suggests ${sn}`);
  await p.fill("#tk-req", "Test Requester"); await p.fill("#tk-issue", "No power");
  await shot(p, "ticket-hint");
  await p.click("#tk-go"); await p.waitForTimeout(150);
  ok(/isn't in the inventory/.test(await p.textContent("#tk-msg")), "first click warns instead of opening");
  await p.click("#tk-go"); await p.waitForTimeout(1300);
  const tix = await p.evaluate(t => invModel().tickets.filter(x => x.sn === t).length, typo);
  ok(tix === 1, "second click opens the ticket as typed");
  await p.evaluate(() => { QBR._invUI.view = "tickets"; renderInventory(); });
  await p.waitForTimeout(300);
  const tag = await p.evaluate(t => { const c = [...document.querySelectorAll("#dash-inventory code")].find(x => x.textContent === t); return c ? { link: !!c.closest("a"), tag: !!(c.parentElement && c.parentElement.querySelector(".inv-notinv")) } : null; }, typo);
  ok(tag && !tag.link && tag.tag, "ticket list shows plain serial + 'Not in inventory' (no dead link)");
  await p.evaluate(t => { location.hash = "#asset/" + encodeURIComponent(t); }, typo); await p.waitForTimeout(400);
  const nf = await p.evaluate(() => { const h = document.querySelector("#a360-body"); return h ? h.innerText : ""; });
  ok(/Unit not in inventory/.test(nf) && nf.includes(sn), "#asset/<unknown> shows the not-found page with a suggestion");
  await shot(p, "asset-notfound");
  await p.click("[data-inv-register]"); await p.waitForTimeout(400);
  ok((await p.inputValue("#in-serials")) === typo, "Register this unit opens intake prefilled");
  await p.evaluate(s => { location.hash = "#asset/" + encodeURIComponent(s); }, sn); await p.waitForTimeout(400);
  ok(!/Unit not in inventory/.test(await p.evaluate(() => document.querySelector("#a360-body").innerText)), "known serial still opens Asset 360");
  ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
  await p.close();
}


async function batchTests(b) {
  console.log("== scan batch review");
  const p = await load(b);
  // pick: 2 in-stock or same-school units, 2 units at another school, 3 new serials, 1 duplicate scan
  const pick = await p.evaluate(() => {
    const A = invModel().assets;
    const byClient = {}; A.forEach(a => { if (a.client) (byClient[a.client] = byClient[a.client] || []).push(a); });
    const schools = Object.keys(byClient).sort((x, y) => byClient[y].length - byClient[x].length);
    const target = schools[0], other = schools[1];
    return { target, other, same: byClient[target].slice(0, 2).map(a => a.sn), moving: byClient[other].slice(0, 2).map(a => a.sn),
             moveFrom: other, expected: byClient[target].length, ws: byClient[other][0].wstart ? 1 : 0 };
  });
  const fresh = ["ZZTEST0001", "ZZTEST0002", "ZZTEST0003"];
  const serials = pick.same.concat(pick.moving, fresh, [fresh[0]]);
  await p.evaluate(list => {
    goToTab("dash-scan");
    const ui = QBR._scanUI;
    list.forEach(sn => ui.rows.push({ id: ui.nextId++, file: null, dataUrl: "", fp: "t" + Math.random(), model: "ThinkPad E14 Gen 6",
      serial: sn, productKey: "", barcodeRaw: "", status: "done", dup: null, sel: true }));
    renderScan();
  }, serials);
  await p.waitForTimeout(300);
  await p.click("#scan-batch-tag"); await p.waitForTimeout(200);
  ok(await p.isVisible("#scan-batch-modal"), "Tag batch opens the review dialog");
  await p.fill("#sb-client", pick.target); await p.waitForTimeout(200);
  const sum = await p.textContent("#sb-sum");
  ok(/3 new/.test(sum) && /2 already at/.test(sum) && /2 at another school/.test(sum) && /1 scanned twice/.test(sum), "groups: 3 new · 2 same · 2 other school · 1 duplicate (" + sum.replace(/\s+/g, " ").slice(0, 120) + ")");
  ok(/add 3 · update 4/.test(await p.textContent("#sb-go")), "apply button shows add 3 · update 4");
  const mv = await p.evaluate(() => [...document.querySelectorAll("#sb-body .scan-move")].map(x => x.textContent));
  ok(mv.length === 2 && mv.every(t => t.includes("→")), "moves highlighted as 'from → to'");
  // skip one of the moves
  await p.click(`[data-sb-act="${pick.moving[1].toUpperCase()}|other"][data-v="skip"]`); await p.waitForTimeout(150);
  ok(/add 3 · update 3/.test(await p.textContent("#sb-go")), "per-row Skip updates the totals (update 3)");
  const code = await p.inputValue("#sb-batch");
  ok(/^B-\d{8}-01$/.test(code), "batch code auto-generated (" + code + ")");
  await p.fill("#sb-sq", "SQ-TEST-1"); await p.fill("#sb-owner", "Test Owner"); await p.fill("#sb-dr", "DR-TEST");
  // warranty: known batch end date for all units + one per-row exception
  await p.fill("#sb-wend", "2029-06-30"); await p.waitForTimeout(100);
  ok(await p.isChecked('input[name="sb-wmode"][value="end"]'), "typing an end date selects 'Known end date'");
  await p.selectOption("#sb-wend-scope", "all"); await p.waitForTimeout(100);
  await p.fill(`[data-sb-end="${fresh[1]}"]`, "2027-12-31"); await p.dispatchEvent(`[data-sb-end="${fresh[1]}"]`, "change");
  if (process.env.SHOTS_DIR || SHOTS) await shot(p, "batch-review");
  await p.click("#sb-go"); await p.waitForTimeout(800);
  const after = await p.evaluate(a => {
    const A = invModel().assets, f = sn => A.find(x => x.sn === sn) || {};
    return { fresh: a.fresh.map(s => f(s).client), dupCount: A.filter(x => x.sn === a.fresh[0]).length,
      wend: [a.fresh[0], a.fresh[1], a.same[0], a.moving[0]].map(s => { const d = f(s).wend; return d ? d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") : null; }),
      wyears: f(a.fresh[0]).wyears, skippedEnd: (f(a.moving[1]).wend || "") + "",
      moved: f(a.moving[0]).client, kept: f(a.moving[1]).client, owner: f(a.same[0]).contact, dr: f(a.same[0]).dr,
      j: (JSON.parse(localStorage.getItem("qbr-inv-journal-v1") || "null") ? 1 : 0), log: (QBR._invLog || invModel().log || []).slice(-6).map(l => JSON.stringify(l)).join(" ") };
  }, { fresh, moving: pick.moving, same: pick.same });
  ok(after.fresh.every(c => c === pick.target), "3 new units registered to " + pick.target);
  ok(after.dupCount === 1, "duplicate scan registered once");
  ok(after.moved === pick.target, "moved unit now at " + pick.target);
  ok(after.kept === pick.moveFrom, "skipped unit left at " + pick.moveFrom);
  ok(after.owner === "Test Owner" && after.dr === "DR-TEST", "owner + DR applied to existing units");
  ok(after.wend[0] === "2029-06-30" && after.wend[2] === "2029-06-30" && after.wend[3] === "2029-06-30", "batch warranty end 2029-06-30 set on new + existing units");
  ok(after.wend[1] === "2027-12-31", "per-row warranty end overrides the batch date");
  ok(after.wyears > 2.5 && after.wyears < 3, "warranty years derived from start → end (" + after.wyears + ")");
  ok(!(await p.isVisible("#scan-batch-modal")), "dialog closes after apply");
  const inBatch = await p.evaluate(c => invModel().assets.filter(a => a.batch === c).map(a => a.sn), code);
  ok(inBatch.length === 6 && !inBatch.includes(pick.moving[1]), "batch code stored on the 6 added/updated units (skipped unit excluded)");
  // Batches view + Edit batch
  await p.evaluate(() => { goToTab("dash-inventory"); QBR._invUI.view = "assets"; QBR._invUI.form = null; renderInventory(); });
  await p.waitForTimeout(300);
  ok(await p.isVisible(`[data-inv-batch-edit="${code}"]`), "Batches card lists " + code);
  await p.click(`[data-inv-batch-show="${code}"]`); await p.waitForTimeout(300);
  const shown = await p.evaluate(() => document.querySelectorAll("#inv-col-assets tbody tr.inv-row").length);
  ok(shown === 6, "Show units filters the asset list to the batch (" + shown + ")");
  await p.click(`[data-inv-batch-edit="${code}"]`); await p.waitForTimeout(300);
  ok(await p.isVisible("#eb-go"), "Edit batch form opens");
  await p.fill("#eb-owner", "Batch Owner"); await p.fill("#eb-we", "2030-01-31");
  await p.uncheck(`.eb-u[value="${fresh[2]}"]`);
  await p.click("#eb-go"); await p.waitForTimeout(300);
  const eb = await p.evaluate(a => { const A = invModel().assets, f = sn => A.find(x => x.sn === sn) || {}, d = x => x ? x.toISOString().slice(0, 10) : null;
    return { o0: f(a[0]).contact, e0: d(f(a[0]).wend), o2: f(a[2]).contact, e2: d(f(a[2]).wend) }; }, fresh);
  ok(eb.o0 === "Batch Owner" && eb.e0 && eb.e0.startsWith("2030-01"), "Edit batch updated owner + warranty end");
  ok(eb.o2 !== "Batch Owner", "unticked unit left unchanged");
  if (SHOTS) await shot(p, "edit-batch");

  console.log("== stocktake");
  await p.evaluate(list => {
    goToTab("dash-scan"); const ui = QBR._scanUI; ui.rows = [];
    list.forEach(sn => ui.rows.push({ id: ui.nextId++, file: null, dataUrl: "", fp: "s" + Math.random(), model: "", serial: sn, productKey: "", barcodeRaw: "", status: "done", dup: null, sel: false }));
    renderScan();
  }, [pick.same[0], pick.moving[1], "ZZUNKNOWN9"]);
  await p.waitForTimeout(200);
  ok(!(await p.isDisabled("#scan-stocktake")), "Stocktake button enabled when rows have serials");
  await p.click("#scan-stocktake"); await p.waitForTimeout(200);
  await p.fill("#sb-client", pick.target); await p.waitForTimeout(200);
  const ss = (await p.textContent("#sb-sum")).replace(/\s+/g, " ");
  const expNow = await p.evaluate(t => invModel().assets.filter(a => (a.client || "").toLowerCase() === t.toLowerCase()).length, pick.target);
  ok(new RegExp("1 found").test(ss) && new RegExp((expNow - 1) + " missing").test(ss) && /1 belong elsewhere/.test(ss) && /1 not in inventory/.test(ss), "found 1 · missing " + (expNow - 1) + " · elsewhere 1 · unknown 1");
  const beforeFix = await p.evaluate(sn => invModel().assets.find(a => a.sn === sn).client, pick.moving[1]);
  ok(beforeFix === pick.moveFrom, "stocktake changes nothing by itself");
  if (SHOTS) await shot(p, "stocktake");
  const dl = p.waitForEvent("download", { timeout: 5000 }).catch(() => null);
  await p.click("#sb-export"); const d = await dl;
  ok(d && /^Stocktake_.*\.xlsx$/.test(d.suggestedFilename()), "export downloads " + (d ? d.suggestedFilename() : "nothing"));
  await p.check(`[data-sb-fix="${pick.moving[1].toUpperCase()}"]`); await p.check('[data-sb-fix="ZZUNKNOWN9"]'); await p.waitForTimeout(100);
  await p.click("#sb-fix"); await p.waitForTimeout(500);
  const fixed = await p.evaluate(a => { const A = invModel().assets; return [A.find(x => x.sn === a[0]).client, (A.find(x => x.sn === "ZZUNKNOWN9") || {}).client]; }, [pick.moving[1]]);
  ok(fixed[0] === pick.target && fixed[1] === pick.target, "fixes applied: moved + registered to " + pick.target);
  ok(/Applied: 1 registered, 1 moved/.test(await p.textContent("#sb-msg")), "dialog stays open with a confirmation");
  ok(/3 found/.test(await p.textContent("#sb-sum")), "re-check after fixes: all 3 scanned units found");
  // stocktake by batch code
  await p.evaluate(list => {
    goToTab("dash-scan"); const ui = QBR._scanUI; ui.rows = [];
    list.forEach(sn => ui.rows.push({ id: ui.nextId++, file: null, dataUrl: "", fp: "b" + Math.random(), model: "", serial: sn, productKey: "", barcodeRaw: "", status: "done", dup: null, sel: false }));
    renderScan();
  }, [fresh[0], pick.moving[1]]);
  await p.waitForTimeout(200);
  await p.click("#scan-stocktake"); await p.waitForTimeout(200);
  await p.selectOption("#sb-target", "batch"); await p.waitForTimeout(100);
  await p.fill("#sb-batchq", code); await p.waitForTimeout(200);
  const bs = (await p.textContent("#sb-sum")).replace(/\s+/g, " ");
  ok(/1 found/.test(bs) && /5 missing/.test(bs) && /1 belong elsewhere/.test(bs), "stocktake by batch: 1 found · 5 missing · 1 elsewhere (" + bs.slice(0, 90) + ")");
  ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
  await p.close();
}

async function bulkTests(b) {
  console.log("== bulk select + bulk edit");
  const p = await load(b);
  // make 3 'assigned but not delivered' units (client set, no Date Delivered) — the user's screenshot case
  const keys = await p.evaluate(() => {
    const A = invModel().assets.filter(a => !a.delivered && QBR.invAssetStatus(a, QBR.invOpenByKey(invModel().tickets)) === "In Stock").slice(0, 3);
    A.forEach(a => QBR.invUpdateAsset(a.key, { client: "Bulk Test School" }));
    goToTab("dash-inventory"); QBR._invUI.view = "assets"; QBR._invUI.flag = null; renderInventory();
    return A.map(a => a.key);
  });
  await p.waitForTimeout(300);
  const fl = await p.evaluate(() => { const b = document.querySelector('[data-inv-flag="assignedNoDel"] .inv-flag-n'); return b ? Number(b.textContent.replace(/,/g, "")) : -1; });
  ok(fl >= 3, "'Assigned, not delivered' flag counts the mis-tagged units (" + fl + ")");
  await p.click('[data-inv-flag="assignedNoDel"]'); await p.waitForTimeout(300);
  await p.fill("#inv-f-client", "Bulk Test School"); await p.press("#inv-f-client", "Enter"); await p.waitForTimeout(300);
  const rowsN = await p.evaluate(() => document.querySelectorAll("#inv-col-assets tbody tr.inv-row").length);
  ok(rowsN === 3, "flag + client filter list exactly those units (" + rowsN + ")");
  await p.check("#inv-sel-all"); await p.waitForTimeout(200);
  ok(/3/.test(await p.textContent("#inv-bulk-edit")), "header checkbox selects the 3 shown units");
  await p.click("#inv-bulk-edit"); await p.waitForTimeout(300);
  ok(await p.isVisible("#bk-preview"), "Bulk edit form opens");
  await p.selectOption("#bk-status", "Deployed"); await p.waitForTimeout(100);
  ok(await p.isVisible("#bk-del"), "Deployed shows the Date delivered field");
  await p.fill("#bk-del", "2026-10-01"); await p.fill("#bk-dr", "DR-BULK-9"); await p.fill("#bk-sq", "SQ-BULK-9");
  await p.fill("#bk-batch", "B-20261001-01"); await p.fill("#bk-we", "2029-12-31");
  ok(await p.isDisabled("#bk-go"), "Apply stays disabled until previewed");
  await p.click("#bk-preview"); await p.waitForTimeout(200);
  const prev = await p.textContent("#bk-plan");
  ok(/Status In Stock → Deployed/.test(prev) && /Date delivered — → 2026-10-01/.test(prev) && /SQ — → SQ-BULK-9/.test(prev), "preview shows from → to per field");
  if (SHOTS) await shot(p, "bulk-edit");
  await p.click("#bk-go"); await p.waitForTimeout(500);
  const after = await p.evaluate(k => {
    const inv = invModel(), ob = QBR.invOpenByKey(inv.tickets);
    return k.map(key => { const a = inv.assets.find(x => x.key === key);
      return { st: QBR.invAssetStatus(a, ob), del: invFmtDate(a.delivered), dr: a.dr, batch: a.batch, we: invFmtDate(a.wend), sq: QBR.invCurrentSq(key) }; });
  }, keys);
  ok(after.every(x => x.st === "Deployed" && x.del === "2026-10-01"), "status now Deployed via Date Delivered (saved as data)");
  ok(after.every(x => x.dr === "DR-BULK-9" && x.batch === "B-20261001-01" && x.we === "2029-12-31"), "DR, batch code and warranty end applied");
  ok(after.every(x => x.sq === "SQ-BULK-9"), "SQ recorded as a deployment row");
  ok(await p.evaluate(() => QBR._invUI.sel.size === 0), "selection cleared after apply");

  console.log("== look up by DR / SQ / batch");
  await p.evaluate(() => { QBR._invUI.flag = null; QBR._invUI.client = "ALL"; renderInventory(); }); await p.waitForTimeout(200);
  await p.fill("#inv-lookup", "DR-BULK"); await p.dispatchEvent("#inv-lookup", "input"); await p.waitForTimeout(200);
  const sug = await p.evaluate(() => [...document.querySelectorAll("#inv-lookup-dd .inv-ac-item")].map(x => x.textContent.replace(/\s+/g, " ").trim()));
  ok(sug.some(t => /DR DR-BULK-9 3 units/.test(t)), "suggestions include 'DR DR-BULK-9 · 3 units'");
  await p.dispatchEvent('#inv-lookup-dd [data-gkind="dr"]', "mousedown"); await p.waitForTimeout(300);
  const gN = await p.evaluate(() => document.querySelectorAll("#inv-col-assets tbody tr.inv-row").length);
  ok(gN === 3 && /DR DR-BULK-9/.test(await p.textContent("#inv-bulkbar")), "picking it filters to the 3 units, with a chip");
  await p.click("#inv-group-x"); await p.waitForTimeout(200);
  await p.fill("#inv-lookup", "sq-bulk-9"); await p.click("#inv-lookup-go"); await p.waitForTimeout(300);
  ok((await p.evaluate(() => document.querySelectorAll("#inv-col-assets tbody tr.inv-row").length)) === 3, "Open with an SQ (any case) filters to its units");
  await p.fill("#inv-lookup", "B-20261001-01"); await p.click("#inv-lookup-go"); await p.waitForTimeout(300);
  ok((await p.evaluate(() => (QBR._invUI.group || {}).kind)) === "batch", "Open with a batch code filters by batch");

  console.log("== bulk status → In Stock");
  await p.evaluate(k => { QBR._invUI.sel = new Set([k]); QBR._invUI.form = "bulkedit"; renderInventory(); }, keys[0]); await p.waitForTimeout(300);
  await p.selectOption("#bk-status", "In Stock"); await p.check("#bk-clr");
  await p.click("#bk-preview"); await p.waitForTimeout(150); await p.click("#bk-go"); await p.waitForTimeout(400);
  const back = await p.evaluate(key => { const a = invModel().assets.find(x => x.key === key); return [QBR.invAssetStatus(a, QBR.invOpenByKey(invModel().tickets)), a.delivered, a.client]; }, keys[0]);
  ok(back[0] === "In Stock" && !back[1] && !back[2], "In Stock clears Date Delivered and (ticked) the client");
  ok(p._errs.length === 0, "no page errors" + (p._errs.length ? ": " + p._errs.join(" | ") : ""));
  await p.close();
}

(async () => {
  const b = await chromium.launch();
  await serialFixes(b);
  await batchTests(b);
  await bulkTests(b);
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
