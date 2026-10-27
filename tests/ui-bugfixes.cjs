// Regression tests for the 2026-10-20 code-review fixes (bug-fixes branch).
//  1. A failed save must not leave rows marked "already written" (patch.js _src stamps).
//  2. Overlapping saves (double-click) run once and append each row once.
//  3. Format-safe save refuses to overwrite a file that changed mid-save when it can't merge.
//  4. esc() escapes quotes, so cell/barcode text can't break out of HTML attributes.
//  5. Audit editor: a blank row in a month sheet doesn't shift saves onto the wrong school.
//  6. Report narrative escapes Postmaster school names.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-bugfixes.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os"), cp = require("child_process");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const INV = path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx");
const AUD = path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

async function open(b, file, engine, kinds) {
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(e => { try { localStorage.setItem("qbr-remember", "0"); localStorage.setItem("qbr-save-engine", e); } catch (x) {} }, engine || "format-safe");
  await p.setInputFiles("#file-input", file); await p.waitForTimeout(2500);
  // Linked-file mock: state.lock makes createWritable fail like a file open in Excel;
  // state.mutateOnRead changes the bytes on the Nth getFile() (OneDrive sync mid-save).
  await p.evaluate(({ name, b64, kinds }) => {
    const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    const st = window.__fx = { bytes: bin, mtime: 1700000000000, lock: false, writes: 0, reads: 0, mutateOnRead: 0 };
    const f0 = APP.files.find(f => f.name === name).blob;
    const handle = { name, kind: "file", queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => {
        st.reads++;
        if (st.mutateOnRead && st.reads === st.mutateOnRead) { const c = st.bytes.slice(); c[c.length - 30] ^= 1; st.bytes = c; }
        return new File([st.bytes], name, { lastModified: st.mtime });
      },
      createWritable: async () => {
        if (st.lock) { const e = new Error("locked"); e.name = "NoModificationAllowedError"; throw e; }
        const parts = []; return { write: async d => parts.push(new Uint8Array(d)), close: async () => { st.bytes = parts[0]; st.mtime += 1000; st.writes++; }, abort: async () => {} };
      } };
    QBR._fsLinks = [{ fp: QBR.fpOf(name, f0), kinds, name, size: bin.length, lastModified: st.mtime, handle }];
    QBR.fsDownloadKind = (kind, mode) => ({ mode: mode || "download", filename: "stub.xlsx" }); // no real downloads in tests
  }, { name: path.basename(file), b64: fs.readFileSync(file).toString("base64"), kinds: kinds || ["assets"] });
  return { p, errs };
}
const bytesOf = async p => new Uint8Array(Buffer.from(await p.evaluate(() => {
  const u = window.__fx.bytes; let s = ""; for (let k = 0; k < u.length; k += 32768) s += String.fromCharCode.apply(null, u.subarray(k, k + 32768)); return btoa(s);
}), "base64"));
const rowsOf = (u8, sheet) => XLSX.utils.sheet_to_json(XLSX.read(u8, { type: "array" }).Sheets[sheet], { header: 1, defval: "" });
const count = (rows, v) => rows.filter(r => r.includes(v)).length;

(async () => {
  const b = await chromium.launch();

  for (const engine of ["format-safe", "legacy"]) {
    console.log(`== 1. failed save keeps appended rows pending (${engine})`);
    const { p, errs } = await open(b, INV, engine);
    const r = await p.evaluate(async () => {
      const a = APP.model.inventory.assets[0];
      QBR.invIntake([{ sn: "PFRETRY0001", client: "RETRY SCHOOL", model: "ThinkPad E14" }]);
      QBR.invAddDeployment({ sn: a.sn, sq: "SQ-RETRY-1", delivered: "2026-10-20" });
      window.__fx.lock = true;
      const first = await QBR.fsSaveKind("assets");
      const stamped = APP.model.inventory.assets.filter(x => x.sn === "PFRETRY0001" && x._src).length;
      window.__fx.lock = false;
      const second = await QBR.fsSaveKind("assets");
      return { first: first.mode, second: second.mode, stamped, writes: window.__fx.writes };
    });
    ok(r.first === "locked" || (engine === "legacy" && r.first === "download-fallback"), "first save did not write the locked file (" + r.first + ")");
    ok(r.stamped === 0, "failed save left no 'already written' stamp on the new row");
    ok(r.second === "file" && r.writes === 1, "retry writes the file (" + r.second + ")");
    const out = await bytesOf(p);
    ok(count(rowsOf(out, "02 DEVICES"), "PFRETRY0001") === 1, "intake row present exactly once after the retry");
    ok(count(rowsOf(out, "04 RAKSO INV."), "SQ-RETRY-1") === 1, "deployment row present exactly once after the retry");
    const after = await p.evaluate(() => (APP.model.inventory.assets.find(x => x.sn === "PFRETRY0001") || {})._src || null);
    ok(!!(after && after.row), "successful save commits the row position for later edits");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();
  }

  console.log("== 2. double-click Save runs one save");
  {
    const { p, errs } = await open(b, INV);
    const r = await p.evaluate(async () => {
      QBR.invIntake([{ sn: "PFDOUBLE001", client: "DOUBLE SCHOOL", model: "ThinkPad E14" }]);
      const p1 = QBR.fsSaveKind("assets"), p2 = QBR.fsSaveKind("assets");
      const [a, c] = await Promise.all([p1, p2]);
      return { same: p1 === p2, a: a.mode, c: c.mode, writes: window.__fx.writes };
    });
    ok(r.same && r.writes === 1, "second click waits for the running save (writes=" + r.writes + ")");
    ok(count(rowsOf(await bytesOf(p), "02 DEVICES"), "PFDOUBLE001") === 1, "row appended once");
    ok(errs.length === 0, "no page errors");
    await p.close();
  }

  console.log("== 3. file changed mid-save without a merge base → nothing overwritten");
  {
    const { p } = await open(b, INV);
    const r = await p.evaluate(async () => {
      QBR.invUpdateAsset(APP.model.inventory.assets[0].key, { cond: "CHANGED-MIDSAVE" });
      QBR._fsBase = {};                       // no base → merge impossible
      window.__fx.mutateOnRead = window.__fx.reads + 2; // fsSaveKind reads once, the last-moment check reads again
      const res = await QBR.fsSaveKind("assets");
      return { mode: res.mode, reason: res.reason || "", writes: window.__fx.writes, pending: QBR.journalOpsFor(QBR._fsLinks[0].fp).length };
    });
    ok(r.mode === "download-unsafe" && r.writes === 0, "save refused, linked file not written (" + r.mode + ": " + r.reason + ")");
    ok(r.pending > 0, "edit stays pending in the journal");
    await p.close();
  }

  console.log("== 4. esc() escapes quotes (attribute XSS)");
  {
    const { p, errs } = await open(b, INV);
    const r = await p.evaluate(async () => {
      const payload = 'X" onmouseover="window.__pwned=1" data-x="';
      const e = esc(payload);
      QBR.invUpdateAsset(APP.model.inventory.assets[0].key, { client: payload });
      try { if (typeof renderAll === "function") renderAll(); } catch (x) {}
      await new Promise(r => setTimeout(r, 300));
      return { e, injected: document.querySelectorAll("[onmouseover]").length, attr: escAttr("a\"b") };
    });
    ok(!/"/.test(r.e) && /&quot;/.test(r.e), "esc() output has no raw double quote");
    ok(r.injected === 0, "no element gained an onmouseover attribute from cell text");
    ok(r.attr === "a&quot;b", "escAttr unchanged (no double-escaping)");
    ok(errs.length === 0, "no page errors");
    await p.close();
  }

  console.log("== 5. audit: blank row in a month sheet");
  {
    const tmp = path.join(os.tmpdir(), "qbr-audit-blank-" + process.pid + ".xlsx");
    cp.execFileSync("python3", ["-c", `
import openpyxl,sys
wb=openpyxl.load_workbook(sys.argv[1]); ws=wb["JANUARY"]; ws.insert_rows(3); wb.save(sys.argv[2])`, AUD, tmp]);
    const { p, errs } = await open(b, tmp, "format-safe", ["audit"]);
    const r = await p.evaluate(async () => {
      const a = QBR._audit; if (!a) return { err: "audit not detected" };
      const ms = a.months.JANUARY, rec = ms.rows.find(x => x.school === "Beta Test Academy");
      if (!rec) return { err: "row not found" };
      QBR.auditUpdateCell(a.fp, ms.sheet, rec.r, "risky", "99");
      const res = await QBR.fsSaveKind("audit");
      return { r: rec.r, mode: res.mode, reason: res.reason || "", notes: res.notes || [] };
    });
    ok(!r.err, r.err || "audit workbook parsed");
    ok(r.r === 4, "Beta Test Academy maps to sheet row 4 after the blank row (got " + r.r + ")");
    const rows = rowsOf(await bytesOf(p), "JANUARY");
    const beta = rows.find(x => x[0] === "Beta Test Academy"), gamma = rows.find(x => x[0] === "Gamma Test College");
    ok(r.mode === "file" && beta && String(beta[2]) === "99", "risky count saved on Beta's own row");
    ok(gamma && String(gamma[2]) !== "99" && rows.filter(x => String(x[2]) === "99").length === 1, "no other school's row received the edit");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close(); try { fs.unlinkSync(tmp); } catch (e) {}
  }

  console.log("== 6. report narrative escapes Postmaster names");
  {
    const { p, errs } = await open(b, INV);
    const r = await p.evaluate(() => {
      const agg = Object.assign({}, computeAggregates());
      agg.pmReported = 2; agg.pmIssues = 1; agg.pmBad = 0;
      agg.pmBadSchools = ['<img src=x onerror="window.__pwned=1">'];
      agg.umTotal = 1; agg.umUpdated = 0; agg.umCompletion = 0; agg.umSY = "<b>SY</b>";
      const rep = QBR.generateReport(APP.model, agg);
      const body = rep.sections.map(s => s.body).join(" ");
      return { raw: /<img|<b>/.test(body), esc: /&lt;img/.test(body) && /&lt;b&gt;SY/.test(body) };
    });
    ok(!r.raw && r.esc, "Postmaster school names and school year are HTML-escaped in the narrative");
    ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
    await p.close();
  }

  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
