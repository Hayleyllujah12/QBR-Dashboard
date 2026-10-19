/* v1.41.0 — MOS › Certificates (browser, file://). Synthetic data only (tests/mos-fixture, dummy names).
 * Usage: PW=<dir with node_modules/playwright> node tests/ui-mos-certs.cjs
 * The linked folder is an in-memory File System Access tree built from tests/mos-fixture, scanned by the REAL
 * deep scan (QBR._fsScanFiles): certificates 4 folders deep, an image folder and a 7-level folder are not opened.
 * A  load score reports + "Load from folder" → certificates found, KPIs (passed / generated / pending / not generated / check)
 * B  Mark printed with no log → creates MOS Certificate Log.xlsx in the folder; row Printed=Yes, count 1, printed by
 * C  Add to log (new certificates) · D reprint → count 2 + reason in Remarks · E mark a whole file as printed
 * F  "No matching result" saved as a link · G an edit made in Excel meanwhile is kept (log re-read before writing)
 * H  new session: the log alone shows printed / pending (no PDFs needed) · I export */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const FX = path.join(__dirname, "mos-fixture");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };

// folder tree: { name: { dir... } | base64 }
function tree(d) {
  const o = {};
  for (const f of fs.readdirSync(d)) { const p = path.join(d, f); o[f] = fs.statSync(p).isDirectory() ? tree(p) : fs.readFileSync(p).toString("base64"); }
  return o;
}
const T = { "MOS Courses": tree(FX) };
const C4 = T["MOS Courses"];
C4["01 - Demo High School"]["Images"] = { "Exam 9999 Certificates.pdf": C4["02 - Sample Academy"]["Certificates"]["Exam 9002 Excel Certificates.pdf"] };   // skipped folder
let deep = C4; ["a", "b", "c", "d", "e", "f", "Certificates g"].forEach(n => { deep = deep[n] = deep[n] || {}; });
deep["Exam 9998 Certificates.pdf"] = C4["02 - Sample Academy"]["Certificates"]["Exam 9002 Excel Certificates.pdf"];                                 // 7 levels: not opened

async function fresh(browser) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 }, acceptDownloads: true });
  page._errs = []; page.on("pageerror", e => page._errs.push(String(e.message).slice(0, 200)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(500);
  await page.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await page.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx")]); await page.waitForTimeout(1500);
  await page.evaluate((T) => {
    const bin = b => { const s = atob(b), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };
    window.__w = 0;
    function file(name, bytes) {
      const h = { kind: "file", name, bytes,
        getFile: async () => new File([h.bytes], name, { lastModified: 1767225600000 }),
        queryPermission: async () => "granted", requestPermission: async () => "granted",
        createWritable: async () => { let buf = null; return { write: async b => { buf = new Uint8Array(b.buffer ? b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) : b); }, close: async () => { h.bytes = buf; window.__w++; }, abort: async () => {} }; } };
      return h;
    }
    function dir(name, t) {
      const kids = {};
      Object.keys(t).forEach(k => { kids[k] = typeof t[k] === "string" ? file(k, bin(t[k])) : dir(k, t[k]); });
      return { kind: "directory", name, kids,
        values: async function* () { for (const k of Object.keys(kids)) yield kids[k]; },
        getFileHandle: async (n, o) => { if (kids[n]) return kids[n]; if (!(o && o.create)) { const e = new Error("not found"); e.name = "NotFoundError"; throw e; } return (kids[n] = file(n, new Uint8Array(0))); } };
    }
    window.__root = window.__root || dir("MOS Courses", T["MOS Courses"]);
    QBR.fsFolderSupported = () => true;
    QBR.fsFolderList = () => [{ id: "d1", name: "MOS Courses", state: "granted" }];
    QBR.fsFolderFiles = async (id, opts) => { window.__scanOpts = { depth: opts.maxDepth }; return QBR._fsScanFiles(window.__root, opts); };
    QBR.fsFolderDir = async () => window.__root;
    try { localStorage.setItem("qbr-mos-printedby", "Tester One"); } catch (e) {}
  }, T);
  await page.click('[data-tab="dash-mos"]'); await page.waitForTimeout(300);
  return page;
}
const kpis = p => p.$$eval("#mos-view-certs .kpi", ks => Object.fromEntries(ks.map(k => [k.querySelector(".kpi-lbl").textContent.trim(), k.querySelector(".kpi-val").textContent.trim()])));
const logRows = p => p.evaluate(() => {
  const h = window.__root.kids["MOS Certificate Log.xlsx"]; if (!h || !h.bytes.length) return null;
  const wb = XLSX.read(h.bytes, { type: "array", cellDates: true }), ws = wb.Sheets["Certificate Log"];
  return XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
});
const row = (rows, code) => (rows || []).find(r => r["Verification Code"] === code);
async function dialogs(page, answers) { page.removeAllListeners("dialog"); page.on("dialog", d => { const a = answers.shift(); if (a === false) d.dismiss(); else d.accept(a === true || a == null ? undefined : a); }); }
const PDFS = school => fs.readdirSync(path.join(FX, school)).filter(f => /Results.*\.pdf$/i.test(f)).map(f => path.join(FX, school, f));

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  let page = await fresh(browser);

  /* ---- A ---- */
  await page.fill("#mos-school-in", "Demo High School"); await page.setInputFiles("#mos-pdf", PDFS("01 - Demo High School"));
  await page.waitForFunction(() => QBR.mos._state().model && !QBR.mos._state().busy, null, { timeout: 30000 });
  await page.fill("#mos-school-in", "Sample Academy"); await page.setInputFiles("#mos-pdf", PDFS("02 - Sample Academy"));
  await page.waitForFunction(() => !QBR.mos._state().busy && QBR.mos._state().files.length === 3, null, { timeout: 30000 });
  await page.click('[data-mos-view="certs"]'); await page.waitForTimeout(200);
  await page.click('[data-mc="folder"]');
  await page.waitForFunction(() => QBR.mos._state().certs.length >= 4 && !QBR.mosCerts._state.busy, null, { timeout: 30000 });
  await page.waitForTimeout(300);
  const st = await page.evaluate(() => ({ n: QBR.mos._state().certs.length, dup: QBR.mos._state().certDupPages, files: QBR.mos._state().certs.map(c => c.fileName), msg: QBR.mosCerts._state.msg, d: window.__scanOpts.depth }));
  ok(st.n === 4 && st.dup === 1, "deep scan: 4 certificates found (4 folders below the linked folder), repeated page ignored");
  ok(!st.files.some(f => /9999|9998/.test(f)), "Images folder and folders deeper than 6 levels are not opened");
  ok(/3 certificate PDFs · 4 new certificates/.test(st.msg) && /deeper than 6 levels not opened/.test(st.msg) && /no log in this folder yet/.test(st.msg), "scan summary: PDFs, depth note, no log yet — " + st.msg.slice(0, 140));
  let k = await kpis(page);
  ok(k["Passed"] === "4" && k["Generated"] === "4" && k["Printed"] === "0" && k["Pending print"] === "3" && k["Passed, no certificate"] === "1" && k["Needs checking"] === "1", "KPIs " + JSON.stringify(k));
  const statuses = await page.$$eval(".mos-cert-tbl tbody tr", trs => trs.map(t => t.children[1].textContent.trim().split(" ")[0] + ":" + t.querySelector(".risk-badge").textContent));
  ok(statuses.includes("Carl:Not generated") && statuses.includes("Zed:Needs checking") && statuses.includes("Bea:Pending print"), "rows: Carl's pass not generated · unmatched certificate needs checking · retake certificate pending");

  /* ---- B ---- */
  await dialogs(page, [true]);   // "create the log?"
  await page.click('tr:has(td:text-is("Alex Cruz Demo")) [data-mc="print"]');
  await page.waitForFunction(() => window.__w >= 2, null, { timeout: 10000 });   // created + saved
  await page.waitForTimeout(300);
  let L = await logRows(page);
  ok(L && L.length === 1 && row(L, "a1b2-C3d4").Printed === "Yes" && row(L, "a1b2-C3d4")["Print Count"] === "1" && row(L, "a1b2-C3d4")["Printed By"] === "Tester One", "Mark printed (no log yet) → MOS Certificate Log.xlsx created in the folder; Printed Yes · count 1 · printed by");
  ok(row(L, "a1b2-C3d4")["Linked Exam Ref #"] === "90000001" && /Certificates SY 25-26/.test(row(L, "a1b2-C3d4")["Source File"]), "log row keeps the linked Exam ref # and the source file path");
  ok(/saves in place/.test(await page.textContent(".mos-certlog")), "log bar: saves in place");

  /* ---- C ---- */
  ok(/3 certificates/.test(await page.textContent(".mos-cert-pend")), "3 certificates not in the log yet");
  await page.click('[data-mc="add-log"]'); await page.waitForFunction(() => window.__w >= 3); await page.waitForTimeout(200);
  L = await logRows(page);
  ok(L.length === 4 && row(L, "d4e5-F6g7").Printed === "No" && !(await page.$(".mos-cert-pend")), "Add to log: 4 rows, new ones Printed No");

  /* ---- D ---- */
  await dialogs(page, [true, "torn"]);
  await page.click('tr:has(td:text-is("Alex Cruz Demo")) [data-mc="reprint"]');
  await page.waitForFunction(() => window.__w >= 4); await page.waitForTimeout(200);
  L = await logRows(page);
  ok(row(L, "a1b2-C3d4")["Print Count"] === "2" && /Reprint \d{4}-\d{2}-\d{2}: torn/.test(row(L, "a1b2-C3d4").Remarks), "reprint: count 2 + reason in Remarks");
  ok(/×2/.test(await page.textContent('tr:has(td:text-is("Alex Cruz Demo"))')), "row shows ×2");

  /* ---- E ---- */
  await page.selectOption('[data-mc="file-pick"]', { label: "Exam 9002 Excel Certificates.pdf" });
  await dialogs(page, [true]);
  await page.click('[data-mc="mark-file"]'); await page.waitForFunction(() => window.__w >= 5); await page.waitForTimeout(200);
  L = await logRows(page);
  ok(row(L, "d4e5-F6g7").Printed === "Yes" && row(L, "b2c3D-4e5f").Printed === "No", "Mark file as printed: only that file's certificates");

  /* ---- F + G ---- */
  await page.evaluate(() => {   // someone edits the log in Excel meanwhile (remark on Bea's row)
    const h = window.__root.kids["MOS Certificate Log.xlsx"], wb = XLSX.read(h.bytes, { type: "array" }), ws = wb.Sheets["Certificate Log"];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }), r = aoa.findIndex(x => x[0] === "b2c3D-4e5f"), c = aoa[0].indexOf("Remarks");
    ws[XLSX.utils.encode_cell({ r, c })] = { t: "s", v: "edited in Excel" };
    h.bytes = new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
  });
  await page.selectOption('tr:has(td:text-is("Zed Ghost Sample")) [data-mc="link"]', "none");
  await page.waitForFunction(() => window.__w >= 6); await page.waitForTimeout(200);
  L = await logRows(page);
  ok(row(L, "z9y8-X7w6")["Linked Exam Ref #"] === "none", "\"No matching result\" saved in the log");
  ok(row(L, "b2c3D-4e5f").Remarks === "edited in Excel", "an edit made in Excel meanwhile is kept (log re-read before writing)");
  k = await kpis(page);
  ok(k["Needs checking"] === "0" && k["Printed"] === "2" && k["Pending print"] === "2", "after the link: nothing to check · 2 printed · 2 pending " + JSON.stringify(k));

  /* ---- I ---- */
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 8000 }).catch(() => null), page.click('[data-mc="export"]')]);
  let ex = null;
  if (dl) { const o = path.join(require("os").tmpdir(), "mos-certs-export.xlsx"); await dl.saveAs(o); ex = XLSX.utils.sheet_to_json(XLSX.read(fs.readFileSync(o)).Sheets.Certificates, { defval: "" }); }
  ok(ex && ex.length === 5 && ex.some(r => r.Status === "Not generated" && r.Name === "Carl Reyes Test"), "Export: 5 rows incl. the pass without a certificate");
  ok(!page._errs.length, "no page errors (A–I) " + page._errs.join(" | "));
  const rootTree = await page.evaluate(() => { const ser = d => Object.fromEntries(Object.entries(d.kids).map(([k, v]) => [k, v.kind === "directory" ? ser(v) : Array.from(v.bytes).map(b => String.fromCharCode(b)).join("")])); return ser(window.__root); });
  await page.close();

  /* ---- H: new session, log only ---- */
  page = await fresh(browser);
  await page.evaluate((t) => {
    const mk = (name, tt) => { const kids = {}; Object.keys(tt).forEach(k => { kids[k] = typeof tt[k] === "string" ? (() => { const s = tt[k], u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); const h = { kind: "file", name: k, bytes: u, getFile: async () => new File([h.bytes], k), queryPermission: async () => "granted", requestPermission: async () => "granted", createWritable: async () => ({ write: async () => {}, close: async () => {} }) }; return h; })() : mk(k, tt[k]); }); return { kind: "directory", name, kids, values: async function* () { for (const k of Object.keys(kids)) yield kids[k]; }, getFileHandle: async n => kids[n] }; };
    window.__root = mk("MOS Courses", t);
  }, rootTree);
  await page.evaluate(() => { const s = QBR.mos._state(); s.view = "certs"; });
  // open the log through the folder scan, without loading certificate PDFs: remove them from the tree first
  await page.evaluate(() => { const strip = d => { Object.keys(d.kids).forEach(k => { const v = d.kids[k]; if (v.kind === "directory") strip(v); else if (/\.pdf$/i.test(k)) delete d.kids[k]; }); }; strip(window.__root); });
  await page.evaluate(() => QBR.mos.render());
  await page.waitForTimeout(200);
  await page.click('[data-mc="folder"]'); await page.waitForFunction(() => QBR.mosCerts._state.log && !QBR.mosCerts._state.busy, null, { timeout: 10000 }); await page.waitForTimeout(300);
  k = await kpis(page);
  ok(k["Generated"] === "4" && k["Printed"] === "2" && k["Pending print"] === "2", "new session, log only: 4 generated · 2 printed · 2 pending " + JSON.stringify(k));
  ok(/4 certificates · 2 printed/.test(await page.textContent(".mos-certlog")), "log bar counts from the workbook");
  ok(!page._errs.length, "no page errors (H) " + page._errs.join(" | "));
  await page.close();

  await browser.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
