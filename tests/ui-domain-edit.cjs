/* Domain Registration inline edit (js/domain-edit.js) — browser, file://. Synthetic tracker only (dummy schools).
 * Usage: PW=<dir with node_modules/playwright> node tests/ui-domain-edit.cjs
 * The tracker mirrors the real layout: Registration "YYYY-MM-DD-HH-MM-SS" text, ValidPD "VALIDPD: 01",
 * "Validy in Year" and Expiration as FORMULAS (cached values), one school with TWO domains, a sheet that starts at B2.
 * A edit Registration date → only that cell changes (text, time-of-day kept); formulas + other cells untouched;
 *   the other domain of the same school is NOT edited; KPIs/expiration recomputed in the dashboard
 * B edit Validity → "VALIDPD: 02" (format kept); formula cell untouched; expiration recomputed in the dashboard
 * C a sheet starting at B2 (offset) still targets the right cell · D cancel (Escape) writes nothing
 * F LINKED tracker (Direct save): loader registers kind "domainreg"; the edit reuses the link's handle — no
 *   "pick the file" prompt, unrelated linked workbooks are never opened; loaded copy + link metadata updated
 *   (in-memory bytes, new fingerprint/hash) so a session reload keeps the edit */
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
  /* F: linked tracker */
  {
    const file = tracker("A1"), before = cells(fs.readFileSync(file));
    const page = await fresh(browser, file);
    const k = await page.evaluate(() => {
      const it = APP.files.find(f => /DOMAIN_REGISTRATION/i.test(f.name)), fp = QBR.fpOf(it.name, it.blob);
      return { fp, kinds: Array.from((QBR._kindByFp || {})[fp] || []) };
    });
    ok(k.kinds.includes("domainreg"), "loader registers kind \"domainreg\" for the tracker's fingerprint (linked tracker = editable)");
    await page.evaluate(fp => {
      window.__picks = 0; window.showOpenFilePicker = async () => { window.__picks++; throw new Error("picker must not open"); };
      const decoy = { name: "Inventory.xlsx", reads: 0, getFile: async () => { decoy.reads++; return new File([new Uint8Array(4)], "Inventory.xlsx"); }, queryPermission: async () => "granted" };
      window.__decoy = decoy;
      const it = APP.files.find(f => /DOMAIN_REGISTRATION/i.test(f.name));
      QBR._fsLinks = [{ fp: "x|1|1", kinds: ["assets"], name: "Inventory.xlsx", size: 4, lastModified: 1, handle: decoy },
                      { fp, kinds: ["domainreg"], name: it.name, size: it.blob.size, lastModified: it.blob.lastModified, hash: "old", handle: window.__h }];
    }, k.fp);
    page._dlg.length = 0;
    await edit(page, "gamma.edu.example", "reg", "2026-03-04");
    const r = await page.evaluate(async () => {
      const it = APP.files.find(f => /DOMAIN_REGISTRATION/i.test(f.name)), L = QBR._fsLinks[1];
      const memBytes = new Uint8Array(await it.blob.arrayBuffer()), disk = window.__h.bytes;
      const same = memBytes.length === disk.length && memBytes.every((b, i) => b === disk[i]);
      const newFp = QBR.fpOf(it.name, it.blob);
      return { picks: window.__picks, writes: window.__h.writes, decoyReads: window.__decoy.reads, same, linkFp: L.fp === newFp, hash: L.hash !== "old" && !!L.hash,
               kind: Array.from((QBR._kindByFp || {})[newFp] || []).includes("domainreg"), curFp: (QBR._currentFps || []).includes(newFp),
               status: APP.model.domainreg.find(x => x.domain === "gamma.edu.example").regRaw };
    });
    ok(r.picks === 0 && !page._dlg.some(m => /pick the DOMAIN_REGISTRATION/.test(m)) && r.writes === 1, "edit saved through the linked handle: no pick prompt, one write  " + JSON.stringify({ picks: r.picks, writes: r.writes, dlg: page._dlg.length }));
    ok(r.decoyReads === 0, "unrelated linked workbook (Inventory.xlsx) is never opened");
    const after = cells(Buffer.from(await outBytes(page)));
    ok(after["Q3!C5"] === "|2026-03-04-10-00-00" && Object.keys(after).filter(c => after[c] !== before[c]).length === 1, "linked file: exactly the edited cell changed");
    ok(r.same, "loaded copy (APP.files) now holds the written bytes, in memory");
    ok(r.linkFp && r.hash && r.kind && r.curFp, "link fingerprint + content hash, kind map and current fingerprints follow the new file  " + JSON.stringify(r));
    await page.evaluate(async () => { const it = APP.files.find(f => /DOMAIN_REGISTRATION/i.test(f.name)); await loadItems(APP.files.slice(), {}); });
    await page.waitForTimeout(800);
    const re = await page.evaluate(() => APP.model.domainreg.find(x => x.domain === "gamma.edu.example").regRaw);
    ok(/^2026-03-04/.test(re || ""), "re-parsing the session copy keeps the edit (was: reverted on reload)  [" + re + "]");
    ok(!page._errs.length, "no page errors (F) " + page._errs.join(" | "));
    await page.close();
  }
  /* G: tracker NOT linked, other workbooks linked → picker fallback, the other links are not opened */
  {
    const file = tracker("A1");
    const page = await fresh(browser, file);
    await page.evaluate(() => {
      window.__reads = 0;
      const mk = n => ({ name: n, getFile: async () => { window.__reads++; return new File([new Uint8Array(4)], n); }, queryPermission: async () => "granted" });
      QBR._fsLinks = [1, 2, 3, 4, 5].map(i => ({ fp: "f" + i, kinds: ["assets"], name: "Workbook" + i + ".xlsx", size: 4, lastModified: 1, handle: mk("Workbook" + i + ".xlsx") }));
    });
    page._dlg.length = 0;
    await edit(page, "gamma.edu.example", "reg", "2026-03-04");
    const g = await page.evaluate(() => ({ reads: window.__reads, writes: window.__h.writes }));
    ok(g.reads === 0 && g.writes === 1 && page._dlg.some(m => /pick the DOMAIN_REGISTRATION/.test(m)), "unlinked tracker: asks to pick the file (fallback) and never opens the 5 other linked workbooks  " + JSON.stringify(g));
    await page.close();
  }
  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
