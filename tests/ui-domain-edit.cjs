/* Domain Registration inline edit (js/domain-edit.js) — browser, file://. Synthetic tracker only (dummy schools).
 * Usage: PW=<dir with node_modules/playwright> node tests/ui-domain-edit.cjs
 * The tracker mirrors the real layout: Registration "YYYY-MM-DD-HH-MM-SS" text, ValidPD "VALIDPD: 01",
 * "Validy in Year" and Expiration as FORMULAS (cached values), one school with TWO domains, a sheet that starts at B2.
 * A edit Registration date → only that cell changes (text, time-of-day kept); formulas + other cells untouched;
 *   the other domain of the same school is NOT edited; KPIs/expiration recomputed in the dashboard
 * B edit Validity → "VALIDPD: 02" (format kept); formula cell untouched; expiration recomputed in the dashboard
 * C a sheet starting at B2 (offset) still targets the right cell · D cancel (Escape) writes nothing */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), "dredit-"));

function tracker(origin, plain) {   // origin: "A1" or "B2"; plain: typed values instead of formulas
  const H = ["Schools", "DOMAIN", "Registration", "ValidPD", "Validy in Year", "Expiration", "Remaining Days", "Status", "Remarks"];
  const rows = [H,
    ["Alpha Academy", "alpha.edu.example", "2025-07-25-09-37-38", "VALIDPD: 01", null, null, "", "Current", ""],
    ["Beta College Inc.", "beta.edu.example", "2025-09-30-10-00-00", "VALIDPD: 01", null, null, "", "Current", ""],
    ["Beta College Inc.", "beta-shs.edu.example", "2025-11-02-08-15-00", "VALIDPD: 01", null, null, "", "Current", ""],
    ["Gamma High School", "gamma.edu.example", "2025-10-15-10-00-00", "VALIDPD: 01", null, null, "", "Current", ""]];
  const ws = XLSX.utils.aoa_to_sheet(rows, { origin });
  const o = XLSX.utils.decode_cell(origin);
  for (let i = 1; i < rows.length; i++) {
    const R = o.r + i, ref = c => XLSX.utils.encode_cell({ r: R, c: o.c + c });
    // cached values the way Excel stores them
    const reg = rows[i][2], y = +reg.slice(0, 4) + 1;
    ws[ref(4)] = plain ? { t: "s", v: "01" } : { t: "s", v: "01", f: 'RIGHT(' + ref(3) + ',2)' };
    ws[ref(5)] = plain ? { t: "s", v: y + reg.slice(4) } : { t: "s", v: y + reg.slice(4), f: 'TEXT(DATE(LEFT(' + ref(2) + ',4)+' + ref(4) + ',MID(' + ref(2) + ',6,2),MID(' + ref(2) + ',9,2)),"yyyy-mm-dd")&MID(' + ref(2) + ',11,9)' };
  }
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Q3");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Notes"], ["keep me"]]), "Notes");
  const p = path.join(DIR, "DOMAIN_REGISTRATION_TRACKER_" + origin + (plain ? "_plain" : "") + ".xlsx");
  fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
  return p;
}
const cells = b => { const wb = XLSX.read(b, { type: "buffer" }); const o = {}; wb.SheetNames.forEach(n => { const ws = wb.Sheets[n]; Object.keys(ws).filter(k => k[0] !== "!").forEach(k => { o[n + "!" + k] = (ws[k].f ? "=" + ws[k].f : "") + "|" + ws[k].v; }); }); return o; };

async function fresh(browser, file) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page._errs = []; page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" }); await page.waitForTimeout(400);
  await page.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await page.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx"), file]); await page.waitForTimeout(1500);
  await page.evaluate((b64) => {
    const s = atob(b64), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i);
    const h = window.__h = { name: "DOMAIN_REGISTRATION_TRACKER.xlsx", bytes: u, writes: 0,
      getFile: async () => new File([h.bytes], h.name), queryPermission: async () => "granted", requestPermission: async () => "granted",
      createWritable: async () => { let buf; return { write: async b => { buf = new Uint8Array(b.buffer ? b.buffer.slice(b.byteOffset || 0, (b.byteOffset || 0) + b.byteLength) : b); }, close: async () => { h.bytes = buf; h.writes++; }, abort: async () => {} }; } };
    window.showOpenFilePicker = async () => [h];
  }, fs.readFileSync(file).toString("base64"));
  page._dlg = []; page.on("dialog", d => { page._dlg.push(d.message()); d.accept(); });
  await page.click('[data-tab="dash-domainreg"]'); await page.waitForTimeout(300);
  return page;
}
const cellFor = (page, domain, field) => page.$(`#tbl-dr tr:has(td:has-text("${domain}")) [data-dr-edit="${field}"]`);
const outBytes = page => page.evaluate(() => Array.from(window.__h.bytes));
async function edit(page, domain, field, value) {
  const c = await cellFor(page, domain, field); await c.click(); await page.waitForTimeout(100);
  const inp = await page.$(`#tbl-dr [data-dr-edit="${field}"] input`);
  await inp.fill(String(value)); await inp.press("Enter");
  await page.waitForFunction(() => !document.querySelector("#tbl-dr input") , null, { timeout: 8000 }).catch(() => {});
  await page.waitForTimeout(600);
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  for (const origin of ["A1", "B2"]) {
    const file = tracker(origin), before = cells(fs.readFileSync(file));
    const page = await fresh(browser, file);
    const o = XLSX.utils.decode_cell(origin), A = (r, c) => "Q3!" + XLSX.utils.encode_cell({ r: o.r + r, c: o.c + c });
    ok(await cellFor(page, "beta-shs.edu.example", "reg"), origin + ": domain rows rendered with editable cells");
    /* A */
    await edit(page, "beta-shs.edu.example", "reg", "2026-03-04");
    let after = cells(Buffer.from(await outBytes(page)));
    const changed = Object.keys(after).filter(k => after[k] !== before[k]);
    ok(changed.length === 1 && changed[0] === A(3, 2), origin + ": registration edit changes exactly one cell (" + changed.join(", ") + ")");
    ok(after[A(3, 2)] === "|2026-03-04-08-15-00", origin + ": written as text, time-of-day kept (" + after[A(3, 2)] + ")");
    ok(after[A(2, 2)] === before[A(2, 2)], origin + ": the same school's OTHER domain is not edited");
    ok(after[A(3, 5)] === before[A(3, 5)] && /^=/.test(after[A(3, 5)]) && after["Notes!A2"] === "|keep me", origin + ": Expiration formula and other sheets untouched");
    const rec = await page.evaluate(() => { const r = APP.model.domainreg.find(x => x.domain === "beta-shs.edu.example"); return { exp: new Date(r.expMs).toISOString().slice(0, 10), reg: r.regRaw }; });
    ok(rec.reg === "2026-03-04-08-15-00" && rec.exp === "2027-03-04", origin + ": dashboard recomputes expiration from the new date (" + JSON.stringify(rec) + ")");
    /* B */
    await edit(page, "gamma.edu.example", "valid", 2);
    after = cells(Buffer.from(await outBytes(page)));
    ok(after[A(4, 3)] === "|VALIDPD: 02", origin + ": validity written in the tracker's own format (" + after[A(4, 3)] + ")");
    ok(after[A(4, 4)] === before[A(4, 4)] && /^=/.test(after[A(4, 4)]), origin + ": “Validy in Year” formula left for Excel to recalculate");
    const g = await page.evaluate(() => { const r = APP.model.domainreg.find(x => x.domain === "gamma.edu.example"); return { yrs: r.validYr, exp: new Date(r.expMs).toISOString().slice(0, 10) }; });
    ok(g.yrs === 2 && g.exp === "2027-10-15", origin + ": dashboard shows 2 yrs and the recomputed expiration (" + JSON.stringify(g) + ")");
    /* D */
    const w0 = await page.evaluate(() => window.__h.writes);
    const c = await cellFor(page, "alpha.edu.example", "reg"); await c.click(); await page.keyboard.press("Escape"); await page.waitForTimeout(300);
    ok(await page.evaluate(w => window.__h.writes === w, w0), origin + ": Escape cancels without writing");
    ok(!page._errs.length, origin + ": no page errors " + page._errs.join(" | "));
    await page.close();
  }
  /* E: typed (non-formula) Validity-in-year + Expiration */
  {
    const file = tracker("A1", true), before = cells(fs.readFileSync(file));
    const page = await fresh(browser, file);
    await edit(page, "gamma.edu.example", "valid", 3);
    const after = cells(Buffer.from(await outBytes(page)));
    const changed = Object.keys(after).filter(k => after[k] !== before[k]).sort();
    ok(JSON.stringify(changed) === '["Q3!D5","Q3!E5"]' && after["Q3!D5"] === "|VALIDPD: 03" && after["Q3!E5"] === "|03", "typed values: VALIDPD and Validy in Year both updated, format kept (" + changed.join(", ") + ")");
    ok(page._dlg.some(m => /Expiration in the tracker is a typed value/.test(m)), "typed Expiration: tells you to update it in Excel");
    ok(!page._errs.length, "no page errors (E) " + page._errs.join(" | "));
    await page.close();
  }

  /* F: stale-copy lineage — the browser remembers 99 rows for this tracker;
   * this file has 4 → saving asks first; declining writes nothing. */
  {
    const file = tracker("A1");
    let acceptLineage = false;
    const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
    page._errs = []; page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
    await page.goto(APP, { waitUntil: "domcontentloaded" }); await page.waitForTimeout(400);
    await page.evaluate(() => localStorage.setItem("qbr-remember", "0"));
    await page.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx"), file]); await page.waitForTimeout(1500);
    await page.evaluate((b64) => {
      const s2 = atob(b64), u = new Uint8Array(s2.length); for (let i = 0; i < s2.length; i++) u[i] = s2.charCodeAt(i);
      const h = window.__h = { name: "DOMAIN_REGISTRATION_TRACKER.xlsx", bytes: u, writes: 0,
        getFile: async () => new File([h.bytes], h.name), queryPermission: async () => "granted", requestPermission: async () => "granted",
        createWritable: async () => { let buf; return { write: async b => { buf = new Uint8Array(b.buffer ? b.buffer.slice(b.byteOffset || 0, (b.byteOffset || 0) + b.byteLength) : b); }, close: async () => { h.bytes = buf; h.writes++; }, abort: async () => {} }; } };
      window.showOpenFilePicker = async () => [h];
      localStorage.setItem("qbr-domainreg-lineage", JSON.stringify({ rows: 99, hash: "stale-test-bogus", savedAt: "2026-10-23T00:00:00.000Z" }));
    }, fs.readFileSync(file).toString("base64"));
    page._dlg = [];
    page.on("dialog", async d => { const m = d.message(); page._dlg.push(m);
      if (/looks older than the last one saved/.test(m) && !acceptLineage) await d.dismiss(); else await d.accept(); });
    await page.click('[data-tab="dash-domainreg"]'); await page.waitForTimeout(300);
    await edit(page, "alpha.edu.example", "reg", "2026-04-05");
    ok(page._dlg.some(m => /looks older than the last one saved from this browser \(99 rows/.test(m)), "F: lineage warning shown for the older copy (99 remembered rows vs this file's 4)");
    ok(page._dlg.some(m => /Could not save: Not saved \u2014 the picked file looks like an older copy/.test(m)), "F: declining the warning refuses the save with the reason");
    ok(await page.evaluate(() => window.__h.writes) === 0, "F: nothing written into the older copy");
    acceptLineage = true;
    await edit(page, "alpha.edu.example", "reg", "2026-04-05");
    ok(await page.evaluate(() => window.__h.writes) === 1, "F: confirming the warning saves the edit into the file");
    const lin = await page.evaluate(() => { const r = JSON.parse(localStorage.getItem("qbr-domainreg-lineage") || "null"); return r && { rows: r.rows, match: r.hash === QBR.fsHashOf(window.__h.bytes) }; });
    ok(!!lin && lin.rows === 4 && lin.match, "F: lineage refreshed to this file (4 rows, hash of the written bytes)");
    page._dlg.length = 0; acceptLineage = false;
    await edit(page, "beta.edu.example", "reg", "2026-05-06");
    ok(await page.evaluate(() => window.__h.writes) === 2 && !page._dlg.some(m => /looks older/.test(m)), "F: once the lineage matches, later saves prompt no more");
    ok(!page._errs.length, "F: no page errors " + page._errs.join(" | "));
    await page.close();
  }
  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
