/* End-to-end Scan page test (Playwright): uploads synthetic label photos through the real UI and checks
 * the extracted serial/model, the engine note and console errors. Works on file:// (default) or http://.
 * Usage: PW=<dir with node_modules/playwright> node tests/ui-scan.cjs [http://host/index.html] */
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const root = path.resolve(__dirname, "..");
const appDir = fs.existsSync(path.join(root, "qbr-app/index.html")) ? path.join(root, "qbr-app") : root;
const URL = process.argv[2] || "file://" + path.join(appDir, "index.html");
const WB = [path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx"), path.join(root, "data/SAMPLE_Lenovo_Inventory_SYNTH.xlsx"), path.join(__dirname, "fixture.xlsx")].find(f => fs.existsSync(f));
const IMGS = ["scan-label-clean.png", "scan-label-phone.jpg", "scan-label-lowres.jpg"].map(f => path.join(__dirname, f));
let pass = 0, fail = 0; const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (x !== undefined ? "  [" + x + "]" : "")); };
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1440, height: 1000 } });
  const errs = []; p.on("pageerror", e => errs.push(e.message)); p.on("console", m => { if (m.type() === "error") errs.push(m.text()); });
  await p.goto(URL);
  console.log("-- " + URL.split(":")[0] + "://");
  await p.setInputFiles("#file-input", WB);
  await p.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 60000 });
  await p.click('.sb-item[data-tab="dash-scan"]'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => APP.activeTab === "dash-scan" && !!document.getElementById("scan-zone")), "Scan page opens");
  ok(!/use Chrome\/Edge/.test(await p.textContent("#scan-engine-note")), "no false 'use Chrome/Edge' warning", await p.textContent("#scan-engine-note"));
  const t0 = Date.now();
  await p.setInputFiles("#scan-files", IMGS);
  await p.waitForFunction(n => QBR._scanUI.rows.length === n && QBR._scanUI.rows.every(r => r.status === "done" || r.status === "error"), IMGS.length, { timeout: 240000 });
  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  const rows = await p.evaluate(() => QBR._scanUI.rows.map(r => ({ status: r.status, serial: r.serial, model: r.model, barcodes: r.barcodeRaw })));
  rows.forEach((r, i) => console.log("     " + path.basename(IMGS[i]).padEnd(24) + r.status + "  serial=" + (r.serial || "—") + "  model=" + (r.model || "—") + "  barcodes=" + (r.barcodes || "—")));
  ok(rows.every(r => r.status === "done"), "all photos scanned (no hang)", secs + "s for " + rows.length);
  ok(rows[0].serial === "PF62SDPW" && /E14 Gen 6/.test(rows[0].model), "clean label → serial + model");
  ok(rows[0].barcodes.split(" | ").length >= 5, "clean label → all 5 stacked barcodes decoded", rows[0].barcodes.split(" | ").length);
  ok(rows[1].serial === "PF62SDPW", "tilted/blurred phone photo → serial", rows[1].serial);
  ok(rows[2].serial === "PF62SDPW", "low-res tilted photo → serial", rows[2].serial);
  const st = await p.evaluate(() => ({ ocr: QBR._scanOCR.mode, failed: QBR._scanOCR.failed, bc: QBR._scanEngines.barcode, note: document.getElementById("scan-engine-note").textContent }));
  ok(!st.failed && st.ocr, "OCR engine running", st.ocr + " · " + st.bc);
  ok(/OCR \(on-device/.test(st.note), "engine note updated after scan", st.note);
  // queue: add a photo while another is scanning — it must not stay "Pending"
  await p.setInputFiles("#scan-files", [IMGS[0]]);
  await p.waitForTimeout(150);
  await p.setInputFiles("#scan-files", [IMGS[2]]);
  await p.waitForFunction(() => QBR._scanUI.rows.length === 5 && QBR._scanUI.rows.every(r => r.status === "done" || r.status === "error"), null, { timeout: 120000 });
  ok(await p.evaluate(() => QBR._scanUI.rows.slice(3).every(r => r.status === "done")), "photos added mid-scan are queued and scanned");
  ok(errs.length === 0, "zero console errors", errs.slice(0, 3).join(" | "));
  console.log(`RESULT: ${pass} passed, ${fail} failed`);
  await b.close(); process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
