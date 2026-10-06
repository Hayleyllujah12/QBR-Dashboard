// Audit editor (v1.28) + format-safe save, end-to-end (headless Chromium, file://).
// Synthetic SAMPLE_AUDIT_RICH.xlsx (styles, CF, DV, hyperlinks) → audit edits (risky
// count, reference link, exempt K/L columns, add row) → Save → inspect the written bytes.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-audit-save.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const S = require(path.join(APPDIR, "js/xlsx-surgical.js"));
const FX = path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const NEWURL = "https://contoso.sharepoint.com/:x:/r/sites/Test/Shared%20Documents/new.xlsx?d=w9&csf=1&web=1";

(async () => {
  const b = await chromium.launch(); const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (x) {} });
  await p.setInputFiles("#file-input", FX); await p.waitForTimeout(2500);
  const name = path.basename(FX), b64 = fs.readFileSync(FX).toString("base64");
  const res = await p.evaluate(async ({ name, b64, NEWURL }) => {
    const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    let bytes = bin, mtime = 1700000000000; const outs = [];
    const f0 = APP.files.find(f => f.name === name).blob;
    const handle = { name, kind: "file", queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => new File([bytes], name, { lastModified: mtime }),
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), close: async () => { bytes = parts[0]; mtime += 1000; } }; } };
    const a = QBR._audit; if (!a) return { err: "audit workbook not detected" };
    // v1.28 progress bar beside the heading + exempt flow (UI level)
    const ui = {};
    try { if (typeof goToTab === "function") goToTab("dash-audit"); else document.querySelector('[data-tab="dash-audit"]').click(); } catch (e) {}
    await new Promise(r => setTimeout(r, 300));
    const pt = document.getElementById("audit-progress-top");
    ui.ptText = pt ? pt.textContent : null;
    ui.badge = (document.getElementById("dash-audit") || document.body).textContent.includes(QBR.AUDIT_VERSION || "1.28.0");
    const firstMonth = Object.values(a.months)[0], sch = firstMonth.rows[2].school;
    QBR.auditSetExempt(a.fp, sch, true, "No tenant access", "all");
    await new Promise(r => setTimeout(r, 200));
    ui.pill = !!(pt && pt.querySelector(".audit-ex-pill"));
    ui.exRow = !!document.querySelector("tr.audit-exempted");
    QBR.auditSetExempt(a.fp, sch, false, "", "all");
    if (typeof QBR.journalClearFp === "function") QBR.journalClearFp(a.fp);   // UI check only: start the save scenario clean
    QBR._fsLinks = [{ fp: a.fp, kinds: ["audit"], name, size: bin.length, lastModified: mtime, handle }];
    const sheet = a.months.JANUARY ? a.months.JANUARY.sheet : Object.values(a.months)[0].sheet;
    const ms = Object.values(a.months).find(x => x.sheet === sheet), r0 = ms.rows[0].r, r1 = ms.rows[1].r;
    QBR.auditUpdateCell(a.fp, sheet, r0, "risky", "77");
    QBR.auditUpdateCell(a.fp, sheet, r1, "ref", "New evidence", NEWURL);
    QBR.auditUpdateCell(a.fp, sheet, r0, "exempt", true);
    QBR.auditUpdateCell(a.fp, sheet, r0, "exemptReason", "No GDAP");
    QBR.auditAddRow(a.fp, sheet, { school: "Delta Test School", org: "ORG-C", risky: "3", health: "Healthy" });
    const results = [];
    for (let i = 0; i < 2; i++) {
      if (i === 1) QBR.auditUpdateCell(QBR._audit.fp, sheet, r1, "risky", "12");
      const r = await QBR.auditFamilySave("audit");
      results.push({ mode: r.mode, engine: r.engine || null, reason: r.reason || null });
      let s = ""; for (let k = 0; k < bytes.length; k += 32768) s += String.fromCharCode.apply(null, bytes.subarray(k, k + 32768));
      outs.push(btoa(s));
    }
    return { results, outs, sheet, r0, r1, ui };
  }, { name, b64, NEWURL });
  if (res.err) { ok(false, res.err); await b.close(); process.exit(1); }
  ok(res.ui && /schools audited/.test(res.ui.ptText || ""), "progress bar renders beside the Audit heading (\"" + String(res.ui && res.ui.ptText).replace(/\s+/g, " ").trim().slice(0, 60) + "\")");
  ok(res.ui && res.ui.badge, "Audit page shows the v1.28.0 version badge");
  ok(res.ui && res.ui.pill && res.ui.exRow, "exempting a school shows the 'exempted' pill and a muted row");
  ok(res.results.every(x => x.mode === "file" && x.engine === "format-safe"), "2 audit saves written by the format-safe engine (" + res.results.map(x => x.mode + (x.reason ? ":" + x.reason : "")).join(", ") + ")");
  const orig = new Uint8Array(fs.readFileSync(FX)), out = new Uint8Array(Buffer.from(res.outs[1], "base64"));
  const wb = XLSX.read(out, { type: "array" }), ws = wb.Sheets[res.sheet], A = (c, r) => ws[c + r];
  const hdr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" })[0];
  const ci = n => XLSX.utils.encode_col(hdr.findIndex(h => String(h).toUpperCase() === n));
  ok(A(ci("TOTAL RISKY USERS"), res.r0) && String(A(ci("TOTAL RISKY USERS"), res.r0).v) === "77", "risky count written");
  ok(hdr.includes("EXEMPT") && hdr.includes("EXEMPT REASON"), "EXEMPT / EXEMPT REASON headers stamped (" + hdr.join(" | ") + ")");
  ok(String((A(ci("EXEMPT"), res.r0) || {}).v) === "Yes" && String((A(ci("EXEMPT REASON"), res.r0) || {}).v) === "No GDAP", "exempt values written");
  const refCell = A(ci("REFERENCES"), res.r1);
  ok(refCell && S._decodeStable(refCell.l && refCell.l.Target) === NEWURL, "reference hyperlink written (& kept, escaped once)");
  ok(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }).some(r => r.includes("Delta Test School")), "added school row present");
  ok(String(A(ci("TOTAL RISKY USERS"), res.r1).v) === "12", "second save applied");
  const z0 = S.zipRead(orig), z1 = S.zipRead(out);
  const sp = n => { const i = XLSX.read(orig, { type: "array" }).SheetNames.indexOf(n); return "xl/worksheets/sheet" + (i + 1) + ".xml"; };
  const x0 = await S.zipText(z0, sp(res.sheet)), x1 = await S.zipText(z1, sp(res.sheet));
  const frag = (x, t) => (x.match(new RegExp("<" + t + "\\b[\\s\\S]*?</" + t + ">", "g")) || []).join("");
  ok(frag(x0, "conditionalFormatting") && frag(x0, "conditionalFormatting") === frag(x1, "conditionalFormatting"), "conditional formatting intact");
  ok(frag(x0, "dataValidations") && frag(x0, "dataValidations") === frag(x1, "dataValidations"), "drop-down validation intact");
  const other = XLSX.read(orig, { type: "array" }).SheetNames.filter(n => n !== res.sheet).map(sp);
  const raw = (z, u8, n) => { const e = z.byName.get(n); return e ? Buffer.from(u8.subarray(e.dataStart, e.dataStart + e.csize)).toString("base64") : null; };
  ok(other.every(n => raw(z0, orig, n) === raw(z1, out, n)), "the other " + other.length + " month sheets byte-identical");
  const rels = await S.zipText(z1, sp(res.sheet).replace("worksheets/", "worksheets/_rels/") + ".rels");
  ok(!/&amp;amp;/.test(rels) && /csf=1&amp;web=1/.test(rels), "existing links stay single-escaped");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
