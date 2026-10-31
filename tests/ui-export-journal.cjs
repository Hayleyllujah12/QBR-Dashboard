// Unlinked Export (audit / storage / usage) clears only the edits it wrote (audit 2026-10-31 H1).
// Before: journalClearFp() dropped every pending edit for the workbook — other kinds' edits and skipped ops.
// Synthetic data only. Usage: PW=<dir with node_modules/playwright> node tests/ui-export-journal.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m, x) => { c ? pass++ : fail++; console.log((c ? "  ✓ " : "  ✗ ") + m + (!c && x !== undefined ? "  [" + JSON.stringify(x) + "]" : "")); };
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ acceptDownloads: true })).newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message)); p.on("dialog", d => d.dismiss().catch(() => {}));
  await p.goto("file://" + path.join(APPDIR, "index.html"));
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx"));
  await p.waitForFunction(() => document.body.classList.contains("has-data"), null, { timeout: 60000 });
  await p.waitForTimeout(1200);
  const r = await p.evaluate(() => {
    const a = APP.model.audit, fp = a && a.fp;
    const m0 = Object.values(a.months).find(x => x && x.rows && x.rows.length);
    QBR.auditUpdateCell(fp, m0.sheet, m0.rows[0].r, "risky", "7");          // a real edit the export writes
    const e = QBR.journalEntries()[fp];
    e.ops.push({ id: "op-other-kind", kind: "storage", op: "storagePasteRow", args: ["Q4", "Sample Academy", {}], ts: Date.now() });
    e.ops.push({ id: "op-skipped", kind: "audit", op: "auditUpdateCell", args: [fp, "NoSuchMonth", 99, "risky", "5"], ts: Date.now() });
    const before = e.ops.map(o => o.id);
    const res = QBR.auditExportWorkbook();
    const after = ((QBR.journalEntries()[fp] || {}).ops || []).map(o => o.id);
    return { mode: res && res.mode, before, after, written: before[0] };
  });
  ok(r.mode === "download", "export downloads the patched copy", r.mode);
  ok(r.before.length === 3 && !r.after.includes(r.written), "the edit written into the copy is cleared", r);
  ok(r.after.includes("op-other-kind"), "another kind's pending edit on the same workbook is kept", r.after);
  ok(r.after.includes("op-skipped"), "an edit the export could not apply is kept", r.after);
  ok(await p.evaluate(() => QBR._lastPatch === null), "patch bookkeeping reset after export");
  ok(errs.length === 0, "no page errors", errs);
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
