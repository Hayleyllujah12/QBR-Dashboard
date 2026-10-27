/* Load progress bar (2026-10-25): the shared bar above the app chrome must
 * advance through Reading -> Parsing -> Rendering on the real upload path,
 * reach 100%, and hide on completion AND on a failed load. The bar updates
 * fast, so a MutationObserver records every bar change into window.__lpLog
 * (observer callbacks deliver every record, even when updates come in bursts).
 * Dummy data only (tests/ fixtures + one corrupt buffer).
 * Run: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-load-progress.browser.cjs */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

const RECORDER = `(() => {
  window.__lpLog = [];
  const snap = () => {
    const bar = document.getElementById("qbr-loadbar");
    if (!bar) return;
    const fill = bar.querySelector(".qbr-loadbar-fill");
    const lab = bar.querySelector(".qbr-loadbar-label");
    const det = bar.querySelector(".qbr-loadbar-detail");
    window.__lpLog.push({ label: lab ? lab.textContent : "", detail: det ? det.textContent : "",
      width: fill ? fill.style.width : "", hidden: !!bar.hidden });
  };
  const inside = (t) => {
    if (!t) return false;
    if (t.id === "qbr-loadbar") return true;
    const el = t.nodeType === 3 ? t.parentElement : t;
    if (el && el.closest) return !!el.closest("#qbr-loadbar");
    return false;
  };
  const ob = new MutationObserver((muts) => {
    if (muts.some((m) => inside(m.target) ||
      [...(m.addedNodes || [])].some((n) => n.id === "qbr-loadbar"))) snap();
  });
  ob.observe(document.body, { childList: true, subtree: true, attributes: true, characterData: true });
})();`;

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  async function load() {
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    const errs = [];
    page.on("pageerror", (e) => errs.push(String(e.message)));
    await page.goto(APP, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);
    return { page, errs };
  }

  // 1. Real upload path: phases advance, filenames shown, 100%, hidden at end
  {
    const { page, errs } = await load();
    await page.evaluate(RECORDER);
    await page.setInputFiles("#file-input", [
      path.join(__dirname, "fixture.xlsx"),
      path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx"),
      path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx"),
    ]);
    await page.waitForFunction(
      () => document.getElementById("upload-status").textContent.includes("Loaded:"),
      { timeout: 30000 });
    await page.waitForTimeout(300);
    const log = await page.evaluate(() => window.__lpLog);
    const has = (pred) => log.some(pred);
    ok(has((e) => e.label === "Reading" && /1\/3/.test(e.detail) && /fixture\.xlsx/.test(e.detail)),
      "Reading phase shows 1/3 + first filename");
    ok(has((e) => e.label === "Reading" && /3\/3/.test(e.detail)),
      "Reading phase reaches 3/3");
    ok(has((e) => e.label === "Parsing" && /3\/3/.test(e.detail)),
      "Parsing phase shows per-file detail up to 3/3");
    ok(has((e) => e.label === "Rendering…" && e.width === "100%"),
      "Rendering phase at 100% width");
    ok(has((e) => e.label === "Parsing" && e.width !== "0%" && e.width !== "100%"),
      "Parsing phase shows intermediate widths (bar actually moves)");
    const barHidden = await page.evaluate(() => {
      const b = document.getElementById("qbr-loadbar"); return !!b && !!b.hidden;
    });
    ok(barHidden, "bar hidden after successful load");
    ok(await page.evaluate(() => document.getElementById("upload-status").textContent.includes("Loaded:")),
      "upload completes with Loaded state (parse results intact)");
    ok(errs.length === 0, "zero page errors during upload load (got " + errs.length + ")");
    await page.close();
  }

  // 2. Failed load: loadItems with an unreadable blob rejects (read phase) ->
  // bar must hide via the finally path, not stick mid-load. SheetJS does NOT
  // throw on garbage bytes (it CSV-falls-back to one sheet), so a corrupt
  // file cannot produce this path — the rejecting blob is the honest trigger.
  {
    const { page, errs } = await load();
    await page.evaluate(RECORDER);
    const outcome = await page.evaluate(async () => {
      const bad = { arrayBuffer: () => Promise.reject(new Error("locked")) };
      let err = null;
      try { await loadItems([{ name: "LOCKED.xlsx", blob: bad }], {}); }
      catch (e) { err = String(e.message); }
      await new Promise((r) => setTimeout(r, 100));
      const bar = document.getElementById("qbr-loadbar");
      return { err, barHidden: !!bar && !!bar.hidden, files: (APP.files || []).length };
    });
    const log = await page.evaluate(() => window.__lpLog);
    ok(outcome.err && /locked/.test(outcome.err), "unreadable blob rejects the load (" + outcome.err + ")");
    ok(log.length > 0, "bar was shown during the failed load (log non-empty)");
    ok(outcome.barHidden, "bar hidden after failed load (no stuck bar)");
    ok(outcome.files === 0, "failed load rolls APP.files back (0 files)");
    ok(errs.length === 0, "zero page errors during failed load (got " + errs.length + ")");
    await page.close();
  }

  // 3. Folder link: scan phase (indeterminate) then Reading, bar hidden at end.
  // The OS picker can't be driven headless, so showDirectoryPicker is stubbed
  // with an in-memory folder (same pattern as tests/ui-folder.cjs).
  {
    const fs = require("fs");
    const b64 = (f) => fs.readFileSync(path.join(__dirname, f)).toString("base64");
    const TREE = {
      "Tracker/ALL TENANT AUTOMATED TRACKER.xlsx": b64("fixture.xlsx"),
      "Audit/RISKY_USERS_AND_DOMAIN.xlsx": b64("SAMPLE_AUDIT_RICH.xlsx"),
      "Inventory/Lenovo Inventory.xlsx": b64("SAMPLE_Lenovo_Inventory_SYNTH.xlsx"),
    };
    const MOCK = `(() => {
      const TREE = ${JSON.stringify(TREE)};
      const bytes = (s) => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
      const store = {}; Object.keys(TREE).forEach((k) => store[k] = bytes(TREE[k]));
      class FH {
        constructor(rel) { this.kind = "file"; this.name = rel.split("/").pop(); this._rel = rel; }
        async queryPermission() { return "granted"; }
        async requestPermission() { return "granted"; }
        async isSameEntry(o) { return o && o._rel === this._rel; }
        async getFile() { return new File([store[this._rel]], this.name, { lastModified: 2000 }); }
      }
      class DH {
        constructor(prefix, name) { this.kind = "directory"; this.name = name; this._p = prefix; }
        async queryPermission() { return "granted"; }
        async requestPermission() { return "granted"; }
        async isSameEntry(o) { return o && o._p === this._p && o.kind === "directory"; }
        async *values() {
          const seen = new Set();
          for (const k of Object.keys(store)) {
            if (!k.startsWith(this._p)) continue;
            const rest = k.slice(this._p.length), i = rest.indexOf("/");
            if (i < 0) yield new FH(k);
            else { const d = rest.slice(0, i); if (!seen.has(d)) { seen.add(d); yield new DH(this._p + d + "/", d); } }
          }
        }
      }
      window.showDirectoryPicker = async () => new DH("", "QBR Workbooks");
    })();`;
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const page = await ctx.newPage();
    const errs = [];
    page.on("pageerror", (e) => errs.push(String(e.message)));
    await page.addInitScript({ content: MOCK });
    await page.goto(APP, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);
    await page.evaluate(RECORDER);
    await page.evaluate(() => { QBR._folderAutoPick = (books) => books.map((_, i) => i); });
    const linked = await page.evaluate(async () => {
      try { return await QBR.fsOpenFolder(); } catch (e) { return "THREW " + e.message; }
    });
    await page.waitForTimeout(300);
    const log = await page.evaluate(() => window.__lpLog);
    const has = (pred) => log.some(pred);
    ok(linked === true, "fsOpenFolder links the mocked folder (got " + linked + ")");
    ok(has((e) => e.label === "Scanning…" && /entries found/.test(e.detail)),
      "Scanning phase shows entries-found detail");
    ok(has((e) => e.label === "Reading" && /\.xlsx/.test(e.detail)),
      "Reading phase shows a workbook filename during folder link");
    ok(has((e) => e.label === "Parsing"), "Parsing phase runs during folder link");
    const barHidden = await page.evaluate(() => {
      const b = document.getElementById("qbr-loadbar"); return !!b && !!b.hidden;
    });
    ok(barHidden, "bar hidden after folder link");
    const toast = await page.evaluate(() => {
      const t = document.getElementById("fo-toast"); return t ? t.textContent : "";
    });
    ok(/Loaded & linked/.test(toast), "end toast still reports Loaded & linked (" + toast + ")");
    ok(errs.length === 0, "zero page errors during folder link (got " + errs.length + ")");
    await ctx.close();
  }

  // 4. Reload-all: a throwing per-file reload is named in the end note, not
  // silently skipped; the bar shows Reloading i/N and hides at the end.
  {
    const { page, errs } = await load();
    await page.evaluate(RECORDER);
    const res = await page.evaluate(async () => {
      QBR._fsLinks = [
        { name: "GOOD.xlsx", dirId: "d1", fp: "fp1" },
        { name: "BAD.xlsx", dirId: "d1", fp: "fp2" },
      ];
      QBR.fsFolderReconnectAll = async () => true;
      QBR.fsReloadFromFile = async (name) => {
        if (name === "BAD.xlsx") throw new Error("locked");
        return true;
      };
      const ok = await QBR.fsFolderReloadAll();
      await new Promise((r) => setTimeout(r, 100));
      const t = document.getElementById("fo-toast");
      const bar = document.getElementById("qbr-loadbar");
      return { ok, toast: t ? t.textContent : "", barHidden: !!bar && !!bar.hidden };
    });
    const log = await page.evaluate(() => window.__lpLog);
    const has = (pred) => log.some(pred);
    ok(res.ok === true, "fsFolderReloadAll returns true");
    ok(/Reloaded ✓ 1 workbook/.test(res.toast) && /failed: BAD\.xlsx/.test(res.toast),
      "end note names the failed file (" + res.toast + ")");
    ok(has((e) => e.label === "Reloading" && /1\/2/.test(e.detail)),
      "Reloading phase shows 1/2");
    ok(has((e) => e.label === "Reloading" && /2\/2/.test(e.detail)),
      "Reloading phase shows 2/2");
    ok(res.barHidden, "bar hidden after reload-all");
    ok(errs.length === 0, "zero page errors during reload-all (got " + errs.length + ")");
    await page.close();
  }

  await browser.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
