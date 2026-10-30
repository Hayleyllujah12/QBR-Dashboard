// 2026-10-30: per-file Direct-save checkbox links an UPLOADED workbook through a folder already opened
// with "Open folder…" — no file picker when the same file is found there (headless Chromium, file://).
// Reuses the in-memory folder mock of tests/ui-folder.cjs (+ resolve() and a file-picker mock).
// Synthetic workbooks only (tests/fixture.xlsx + SAMPLE_* fixtures).
// Usage: PW=<dir with node_modules/playwright> node tests/ui-folder-autolink.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const b64 = f => fs.readFileSync(path.join(__dirname, f)).toString("base64");
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const trackerNoDomain = (() => {   // reload-only workbook (no sheets the dashboard edits)
  const wb = XLSX.read(fs.readFileSync(path.join(__dirname, "fixture.xlsx")));
  wb.SheetNames.filter(n => { const r = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 })[0] || []; return r.some(c => /registration|remaining day/i.test(String(c))); })
    .forEach(n => { delete wb.Sheets[n]; wb.SheetNames.splice(wb.SheetNames.indexOf(n), 1); });
  return XLSX.write(wb, { bookType: "xlsx", type: "base64" });
})();
const AUD = b64("SAMPLE_AUDIT_RICH.xlsx");
const INV = b64("SAMPLE_Lenovo_Inventory_SYNTH.xlsx"), INV2 = b64("SAMPLE_Lenovo_Inventory_RICH.xlsx");
const NAME = "Lenovo Inventory.xlsx";
const TREE = {
  "Inventory/Lenovo Inventory.xlsx": { b: INV, t: 2000 },
  "Tracker/ALL TENANT AUTOMATED TRACKER.xlsx": { b: trackerNoDomain, t: 3000 },
};

const MOCK = `(() => {
  const TREE = ${JSON.stringify(TREE)};
  const T = window.__fo = { dirReq: 0, fileReq: 0, writes: 0, picks: 0, session: 1 };
  const bytes = s => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const store = {}; Object.keys(TREE).forEach(k => store[k] = { u: bytes(TREE[k].b), t: TREE[k].t });
  // one grant per picked ROOT folder ("" = first folder, "@2/" = a second folder elsewhere)
  const st = {}; const rootOf = k => k.charAt(0) === "@" ? k.slice(0, k.indexOf("/") + 1) : "";
  const gs = r => st[r] || "prompt";
  class FH {
    constructor(rel, viaDir) { this.kind = "file"; this.name = rel.split("/").pop(); this._rel = rel; this._via = viaDir; this._session = T.session; this._own = "prompt"; }
    _state() { return (this._via && this._session === T.session) ? gs(rootOf(this._rel)) : this._own; }
    async queryPermission() { return this._state(); }
    async requestPermission() { T.fileReq++; this._own = "granted"; return "granted"; }
    async isSameEntry(o) { return o && o._rel === this._rel; }
    async getFile() { const s = store[this._rel]; return new File([s.u], this.name, { lastModified: s.t }); }
    async createWritable() {
      if (this._state() !== "granted") throw new DOMException("not allowed", "NotAllowedError");
      const self = this; let buf = null;
      return { async write(d) { buf = new Uint8Array(d instanceof ArrayBuffer ? d : (d.buffer ? d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength) : d)); },
               async close() { T.writes++; store[self._rel] = { u: buf, t: Date.now() }; }, async abort() {} };
    }
  }
  class DH {
    constructor(prefix, name) { this.kind = "directory"; this.name = name; this._p = prefix; }
    async queryPermission() { return gs(rootOf(this._p)); }
    async requestPermission() { T.dirReq++; const r = rootOf(this._p); T.reqBy[r] = (T.reqBy[r] || 0) + 1; if (T.denyDir) { st[r] = "denied"; return "denied"; } st[r] = "granted"; return "granted"; }
    async isSameEntry(o) { return o && o._p === this._p && o.kind === "directory"; }
    async resolve(h) { if (!h || !h._rel || !h._rel.startsWith(this._p) || rootOf(h._rel) !== rootOf(this._p)) return null; return h._rel.slice(this._p.length).split("/"); }
    async getDirectoryHandle(n) { const p = this._p + n + "/"; if (!Object.keys(store).some(k => k.startsWith(p))) throw new DOMException("nf", "NotFoundError"); return new DH(p, n); }
    async getFileHandle(n) { const k = this._p + n; if (!store[k]) throw new DOMException("nf", "NotFoundError"); return new FH(k, true); }
    async *values() {
      const seen = new Set();
      for (const k of Object.keys(store)) {
        if (!k.startsWith(this._p) || (this._p === "" && k.charAt(0) === "@")) continue;
        const rest = k.slice(this._p.length), i = rest.indexOf("/");
        if (i < 0) yield new FH(k, true);
        else { const d = rest.slice(0, i); if (!seen.has(d)) { seen.add(d); yield new DH(this._p + d + "/", d); } }
      }
    }
  }
  window.showDirectoryPicker = async (opt) => { T.picks++; T.lastOpt = opt && { id: opt.id, mode: opt.mode }; const r = T.nextRoot || ""; if (opt && opt.mode === "readwrite") st[r] = "granted"; return new DH(r, r ? "Domains Share" : "QBR Workbooks"); };
  // next browser session: every grant is forgotten, handles restored from storage start at "prompt"
  T.reqBy = {};
  T.newSession = () => { T.session++; Object.keys(st).forEach(r => st[r] = "prompt"); };
  T.store = store;
  // file picker: T.pickQueue items = "abort" | "security" | {rel, b64?, t?} (adds the file first when b64 given)
  T.filePicks = 0; T.pickQueue = [];
  window.showOpenFilePicker = async (opt) => {
    T.filePicks++; T.lastPickOpt = { startIn: !!(opt && opt.startIn) };
    const a = T.pickQueue.shift();
    if (!a || a === "abort") throw new DOMException("cancelled", "AbortError");
    if (a === "security") throw new DOMException("Must be handling a user gesture", "SecurityError");
    if (a.b64) store[a.rel] = { u: bytes(a.b64), t: a.t || Date.now() };
    return [new FH(a.rel, true)];
  };
})();`;

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1500, height: 900 } });
  const p = await ctx.newPage(); const errs = [], ext = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("request", r => { if (/^https?:/.test(r.url())) ext.push(r.url()); });
  await p.addInitScript({ content: MOCK });
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(() => {
    localStorage.setItem("qbr-remember", "0");
    const orig = window.loadItems; window.__fo.loads = 0;
    window.loadItems = function () { window.__fo.loads++; return orig.apply(this, arguments); };
  });
  // "upload" = a plain File (no location), like drag-drop / Choose files
  const upload = (name, data, t) => p.evaluate(async ([name, data, t]) => {
    const bin = atob(data); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    await handleFiles([new File([u], name, { lastModified: t })]);
  }, [name, data, t]);
  const snap = () => p.evaluate(n => {
    const T = window.__fo, l = (QBR._fsLinks || []).find(x => x.name === n);
    return { linked: !!l, dirId: l ? l.dirId : null, rel: l && l.rel ? l.rel.join("/") : "", lm: l ? l.lastModified : null,
      picks: T.filePicks, dirReq: T.dirReq, fileReq: T.fileReq, loads: T.loads, dirs: (QBR._fsDirs || []).length,
      toast: ((document.getElementById("fo-toast") || {}).textContent || "") };
  }, NAME);
  const toggle = (n, on) => p.evaluate(([n, on]) => QBR.fsToggleFileLink(n, on), [n, on]);

  await upload(NAME, INV, 2000); await upload("ALL TENANT AUTOMATED TRACKER.xlsx", trackerNoDomain, 3000); await p.waitForTimeout(1500);
  ok(await p.evaluate(() => (APP.files || []).length === 2), "two workbooks uploaded (plain upload, no folder)");

  console.log("== no folder remembered → file picker as before");
  let s0 = await snap();
  let r = await toggle(NAME, true); let s = await snap();
  ok(r === false && s.picks === s0.picks + 1 && !s.linked, "tick opens the file picker; cancelling links nothing  " + JSON.stringify(s));

  console.log("== folder remembered + allowed → silent link");
  await p.evaluate(async () => { await QBR.fsFolderPickOnly(); QBR._fsDirs[0].keep = false; });
  s0 = await snap();
  await p.evaluate(n => { const c = document.querySelector('.loaded-file-edit[data-fname="' + n + '"]'); c.click(); }, NAME);
  await p.waitForTimeout(1500); s = await snap();
  ok(s.linked && s.dirId && s.rel === "Inventory/Lenovo Inventory.xlsx", "clicking the checkbox links it through the folder (dirId + subfolder path)  " + JSON.stringify({ rel: s.rel }));
  ok(s.picks === s0.picks && s.dirReq === s0.dirReq && s.fileReq === s0.fileReq, "…with NO file picker and NO permission prompt");
  ok(s.loads === s0.loads, "…and no reload (the uploaded copy is the same file)");
  ok(/found in/.test(s.toast), "toast says where it was found: " + s.toast);
  ok(await p.evaluate(n => document.querySelector('.loaded-file-edit[data-fname="' + n + '"]').checked, NAME), "the box shows ticked");

  console.log("== same bytes, different timestamp (OneDrive) → still the same file");
  await toggle(NAME, false);
  await p.evaluate(() => window.__fo.store["Inventory/Lenovo Inventory.xlsx"].t = 2500);
  s0 = await snap(); r = await toggle(NAME, true); s = await snap();
  ok(r === true && s.linked && s.picks === s0.picks && s.loads === s0.loads, "content hash matches → silent link, no reload");

  console.log("== untick / re-tick");
  await toggle(NAME, false); s = await snap();
  ok(!s.linked && s.dirs === 1, "untick removes the link but the folder stays remembered (was forgotten before)");
  s0 = s; r = await toggle(NAME, true); s = await snap();
  ok(r === true && s.linked && s.picks === s0.picks && s.dirReq === s0.dirReq, "re-tick links again silently");

  console.log("== next session: one folder prompt, no picker");
  await toggle(NAME, false); await p.evaluate(() => window.__fo.newSession());
  s0 = await snap(); r = await toggle(NAME, true); s = await snap();
  ok(r === true && s.linked && s.dirReq === s0.dirReq + 1 && s.picks === s0.picks && s.fileReq === s0.fileReq, "exactly ONE folder prompt, zero file prompts  " + JSON.stringify({ dirReq: s.dirReq - s0.dirReq }));

  console.log("== folder prompt declined → file picker");
  await toggle(NAME, false); await p.evaluate(() => { window.__fo.newSession(); window.__fo.denyDir = true; });
  s0 = await snap(); r = await toggle(NAME, true); s = await snap();
  await p.evaluate(() => { window.__fo.denyDir = false; });
  ok(s.dirReq === s0.dirReq + 1 && s.picks === s0.picks + 1 && !s.linked, "declined → falls back to the picker (cancelled → no link)");
  ok(await p.evaluate(() => window.__fo.lastPickOpt.startIn), "the picker starts in the remembered folder");

  console.log("== folder copy differs from the upload");
  await p.evaluate(b => { const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    window.__fo.store["Inventory/Lenovo Inventory.xlsx"] = { u, t: 7000 }; }, INV2);
  s0 = await snap();
  const tick = p.evaluate(n => QBR.fsToggleFileLink(n, true), NAME);
  await p.waitForSelector(".fo-al", { timeout: 5000 });
  const dlg = await p.evaluate(() => ({ h: document.getElementById("fo-al-h").textContent, rows: document.querySelectorAll(".fo-al .fo-al-pick").length,
    btns: [...document.querySelectorAll(".fo-al [data-al]")].map(x => x.textContent) }));
  ok(/Different copy/.test(dlg.h) && dlg.rows === 1 && dlg.btns.length === 3, "asks instead of reloading on its own  " + JSON.stringify(dlg));
  await p.click('.fo-al [data-al="cancel"]'); r = await tick; s = await snap();
  ok(r === false && !s.linked && s.loads === s0.loads && s.picks === s0.picks, "Cancel → nothing linked, nothing reloaded");
  r = await p.evaluate(n => { QBR._autoLinkChoose = (list, pending) => { window.__fo.choose = list; return 0; }; return QBR.fsToggleFileLink(n, true); }, NAME);
  await p.waitForTimeout(800); s = await snap();
  ok(r === true && s.linked && s.lm === 7000 && s.loads === s0.loads + 1, "'Link & reload' reloads from the folder copy and links it");
  ok(await p.evaluate(() => (APP.files || []).some(f => f.name === "Lenovo Inventory.xlsx" && f.blob.lastModified === 7000)), "the dashboard now holds the folder copy");

  console.log("== several copies: identical one wins, else choose");
  await toggle(NAME, false);
  await p.evaluate(b => { const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    window.__fo.store["Copy/Lenovo Inventory.xlsx"] = { u, t: 9000 }; QBR.fsFolderScanReset(); }, INV);
  s0 = await snap(); r = await toggle(NAME, true); s = await snap();
  ok(r === true && s.rel === "Inventory/Lenovo Inventory.xlsx" && s.loads === s0.loads, "the copy identical to the loaded one is linked, not the newer different one");
  await toggle(NAME, false);
  await p.evaluate(b => { const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    window.__fo.store["Inventory/Lenovo Inventory.xlsx"] = { u, t: 9500 }; QBR.fsFolderScanReset(); }, INV);
  s0 = await snap();
  r = await p.evaluate(n => { window.__fo.choose = null; QBR._autoLinkChoose = list => { window.__fo.choose = list; return "picker"; }; return QBR.fsToggleFileLink(n, true); }, NAME);
  s = await snap();
  const ch = await p.evaluate(() => window.__fo.choose);
  ok(ch && ch.length === 2 && ch[0].lastModified >= ch[1].lastModified, "two different copies → both offered, newest first  " + JSON.stringify(ch && ch.map(c => c.where)));
  ok(s.picks === s0.picks + 1 && s.loads === s0.loads, "'Pick another file…' opens the picker");
  await p.evaluate(() => { QBR._autoLinkChoose = null; });

  console.log("== not in the folder → picker; a file picked inside the folder becomes a folder link");
  await upload("RISKY_USERS_AND_DOMAIN.xlsx", AUD, 4000); await p.waitForTimeout(1200);
  await p.evaluate(b => { window.__fo.pickQueue.push({ rel: "Inventory/RISKY_USERS_AND_DOMAIN.xlsx", b64: b, t: 4000 }); }, AUD);
  s0 = await snap();
  r = await toggle("RISKY_USERS_AND_DOMAIN.xlsx", true); s = await snap();
  const l2 = await p.evaluate(() => { const l = (QBR._fsLinks || []).find(x => x.name === "RISKY_USERS_AND_DOMAIN.xlsx"); return l && { dirId: !!l.dirId, rel: (l.rel || []).join("/") }; });
  ok(r === true && s.picks === s0.picks + 1, "not found → the picker opens once");
  ok(l2 && l2.dirId && l2.rel === "Inventory/RISKY_USERS_AND_DOMAIN.xlsx", "picked inside the remembered folder → stored as a folder link  " + JSON.stringify(l2));
  await toggle("RISKY_USERS_AND_DOMAIN.xlsx", false);

  console.log("== click expired before the picker → one more click");
  await p.evaluate(() => { window.__fo.store["Inventory/RISKY_USERS_AND_DOMAIN.xlsx"] = undefined; delete window.__fo.store["Inventory/RISKY_USERS_AND_DOMAIN.xlsx"]; QBR.fsFolderScanReset();
    window.__fo.pickQueue.push("security", "abort"); window.__fo.asked = 0; QBR._autoAskClick = () => { window.__fo.asked++; return true; }; });
  s0 = await snap();
  r = await toggle("RISKY_USERS_AND_DOMAIN.xlsx", true); s = await snap();
  ok(await p.evaluate(() => window.__fo.asked === 1) && s.picks === s0.picks + 2 && r === false, "refused picker → 'Choose file…' button → picker retried once");
  await p.evaluate(() => { QBR._autoAskClick = null; });

  console.log("== workbook with nothing to edit");
  s0 = await snap();
  r = await toggle("ALL TENANT AUTOMATED TRACKER.xlsx", true); s = await snap();
  ok(r === false && s.picks === s0.picks && s.dirReq === s0.dirReq && /no sheets the dashboard edits/.test(s.toast), "says so straight away — no picker, no prompt, no link");

  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  ok(ext.length === 0, "no external requests (offline)");
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
