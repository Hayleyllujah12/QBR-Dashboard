// v1.33.0 "Open folder": one permission loads + links every workbook (headless Chromium, file://).
// The OS folder picker can't be driven headless, so window.showDirectoryPicker is replaced by an
// in-memory folder that mimics Chrome's permission model: a grant on the folder covers every file
// handle obtained through it; handles restored in a later session start at "prompt".
// Synthetic workbooks only (tests/fixture.xlsx + SAMPLE_* fixtures).
// Usage: PW=<dir with node_modules/playwright> node tests/ui-folder.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const b64 = f => fs.readFileSync(path.join(__dirname, f)).toString("base64");

// folder layout (paths relative to the picked folder)
const TREE = {
  "Tracker/ALL TENANT AUTOMATED TRACKER.xlsx": { b: b64("fixture.xlsx"), t: 3000 },
  "Old/ALL TENANT AUTOMATED TRACKER.xlsx": { b: b64("fixture.xlsx"), t: 1000 },          // same name, older
  "Audit/RISKY_USERS_AND_DOMAIN.xlsx": { b: b64("SAMPLE_AUDIT_RICH.xlsx"), t: 2000 },
  "Inventory/Lenovo Inventory.xlsx": { b: b64("SAMPLE_Lenovo_Inventory_SYNTH.xlsx"), t: 2000 },
  "Tracker/~$ALL TENANT AUTOMATED TRACKER.xlsx": { b: "", t: 1 },                         // Excel lock file
  ".hidden/secret.xlsx": { b: "", t: 1 },
  "A/B/C/D/too-deep.xlsx": { b: "", t: 1 },                                               // depth 4
  "notes.txt": { b: "", t: 1 },
};

const MOCK = `(() => {
  const TREE = ${JSON.stringify(TREE)};
  const T = window.__fo = { dirReq: 0, fileReq: 0, writes: 0, picks: 0, session: 1 };
  const bytes = s => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const store = {}; Object.keys(TREE).forEach(k => store[k] = { u: bytes(TREE[k].b), t: TREE[k].t });
  let dirState = "prompt";
  class FH {
    constructor(rel, viaDir) { this.kind = "file"; this.name = rel.split("/").pop(); this._rel = rel; this._via = viaDir; this._session = T.session; this._own = "prompt"; }
    _state() { return (this._via && this._session === T.session) ? dirState : this._own; }
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
    async queryPermission() { return dirState; }
    async requestPermission() { T.dirReq++; if (T.denyDir) { dirState = "denied"; return "denied"; } dirState = "granted"; return "granted"; }
    async isSameEntry(o) { return o && o._p === this._p && o.kind === "directory"; }
    async getDirectoryHandle(n) { const p = this._p + n + "/"; if (!Object.keys(store).some(k => k.startsWith(p))) throw new DOMException("nf", "NotFoundError"); return new DH(p, n); }
    async getFileHandle(n) { const k = this._p + n; if (!store[k]) throw new DOMException("nf", "NotFoundError"); return new FH(k, true); }
    async *values() {
      const seen = new Set();
      for (const k of Object.keys(store)) {
        if (!k.startsWith(this._p)) continue;
        const rest = k.slice(this._p.length), i = rest.indexOf("/");
        if (i < 0) yield new FH(k, true);
        else { const d = rest.slice(0, i); if (!seen.has(d)) { seen.add(d); yield new DH(this._p + d + "/", d); } }
      }
    }
  }
  window.showDirectoryPicker = async (opt) => { T.picks++; T.lastOpt = opt && { id: opt.id, mode: opt.mode }; if (opt && opt.mode === "readwrite") dirState = "granted"; return new DH("", "QBR Workbooks"); };
  // next browser session: every grant is forgotten, handles restored from storage start at "prompt"
  T.newSession = () => { T.session++; dirState = "prompt"; };
  T.store = store;
})();`;

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1500, height: 900 } });
  const p = await ctx.newPage(); const errs = [], ext = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("request", r => { if (/^https?:/.test(r.url())) ext.push(r.url()); });
  await p.addInitScript({ content: MOCK });
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));

  console.log("== entry points");
  ok(await p.evaluate(() => document.body.classList.contains("fs-folder-ok")), "folder support detected (Chrome/Edge API present)");
  ok(await p.isVisible("#dz-folder"), "empty state shows 'Open folder…' next to Choose files");
  ok(await p.isVisible("#hdr-folder"), "header shows a Folder button");

  console.log("== cancel does nothing");
  await p.click("#dz-folder"); await p.waitForSelector(".fo-ov");
  await p.click("#fo-cancel"); await p.waitForTimeout(200);
  ok(await p.evaluate(() => !(APP.files || []).length && !document.querySelector(".fo-ov")), "Cancel closes the dialog and loads nothing");

  console.log("== scan + pick");
  await p.click("#dz-folder"); await p.waitForSelector(".fo-ov");
  const dlg = await p.evaluate(() => ({
    opt: window.__fo.lastOpt,
    rows: [...document.querySelectorAll(".fo-table tbody tr")].map(r => r.cells[1].textContent.trim() + "|" + r.cells[2].textContent.trim()),
    ticked: [...document.querySelectorAll(".fo-pick")].filter(c => c.checked).map(c => c.closest("tr").cells[2].textContent.trim() + "/" + c.closest("tr").cells[1].textContent.replace("same name", "").trim()),
    btn: document.getElementById("fo-ok").textContent,
  }));
  ok(dlg.opt && dlg.opt.mode === "readwrite" && dlg.opt.id === "qbr-workbooks", "picker asks for read/write once (remembered start folder id)");
  ok(dlg.rows.length === 4, "finds the 4 workbooks in subfolders; skips lock file, hidden folder, >3 levels deep, non-Excel  " + JSON.stringify(dlg.rows));
  ok(dlg.ticked.length === 3 && dlg.ticked.includes("Tracker/ALL TENANT AUTOMATED TRACKER.xlsx") && !dlg.ticked.some(x => x.startsWith("Old/")), "pre-ticks all, but only the NEWEST of two same-name files  " + JSON.stringify(dlg.ticked));
  ok(/\(3\)/.test(dlg.btn), "button shows the count: " + dlg.btn);
  // ticking the old twin unticks the new one
  await p.evaluate(() => { const c = [...document.querySelectorAll(".fo-pick")].find(x => x.closest("tr").cells[2].textContent.trim() === "Old"); c.click(); });
  const twin = await p.evaluate(() => [...document.querySelectorAll(".fo-pick")].filter(c => c.checked).map(c => c.closest("tr").cells[2].textContent.trim()));
  ok(twin.includes("Old") && !twin.includes("Tracker"), "same-name twins are exclusive (one per name)");
  await p.evaluate(() => { const c = [...document.querySelectorAll(".fo-pick")].find(x => x.closest("tr").cells[2].textContent.trim() === "Tracker"); c.click(); });
  await p.click("#fo-ok"); await p.waitForTimeout(3500);

  console.log("== load + link in one go");
  const st = await p.evaluate(() => ({
    files: (APP.files || []).map(f => f.name).sort(),
    links: (QBR._fsLinks || []).map(l => ({ n: l.name, k: l.kinds, d: !!l.dirId, rel: (l.rel || []).join("/") })),
    fileReq: window.__fo.fileReq, dirReq: window.__fo.dirReq,
    boxes: [...document.querySelectorAll(".loaded-file-edit")].map(c => c.checked),
    bar: [...document.querySelectorAll(".fo-bar [data-fo-act]")].map(x => x.dataset.foAct),
    model: !!(APP.model && APP.model.sources && APP.model.sources.risky),
  }));
  ok(st.files.length === 3, "3 workbooks loaded  " + JSON.stringify(st.files));
  ok(st.model, "dashboard model built from the folder (risky source found)");
  ok(st.links.length === 3 && st.links.every(l => l.d), "all 3 are linked through the folder");
  ok(st.links.some(l => l.rel === "Tracker/ALL TENANT AUTOMATED TRACKER.xlsx"), "the picked twin (Tracker/) is the one linked");
  ok(st.links.some(l => (l.k || []).includes("audit")) && st.links.some(l => (l.k || []).includes("assets")), "editable kinds detected (audit + inventory)");
  ok(st.fileReq === 0, "ZERO per-file permission prompts (was one per workbook)");
  ok(st.boxes.length === 3 && st.boxes.every(Boolean), "every Direct-save box in the list is ticked");
  ok(st.bar.includes("open") && st.bar.includes("reload") && !st.bar.includes("reconnect"), "list toolbar: Open folder + Reload all, no Reconnect while granted");
  ok(await p.evaluate(() => { const b = document.querySelectorAll("#fs-link-status .badge"); return b.length === 1 && /3 linked/.test(b[0].textContent); }), "header shows one '🔗 3 linked' badge, not one per file");
  const vo = await p.evaluate(() => { const c = [...document.querySelectorAll(".loaded-file-edit")].find(x => /TRACKER/.test(x.dataset.fname)); return c && c.title; });
  ok(/nothing is ever written/.test(vo || ""), "a workbook without editable sheets is marked reload-only");

  console.log("== next session: one folder prompt re-grants everything");
  await p.evaluate(() => window.__fo.newSession());
  await p.evaluate(() => QBR.fsFolderRefreshState()); await p.waitForTimeout(200);
  ok(await p.evaluate(() => !!document.querySelector('.fo-bar [data-fo-act="reconnect"]')), "'Reconnect folder' appears when the browser forgot the grant");
  const s1 = await p.evaluate(async () => {
    const r = await QBR.fsSaveKind("assets");
    return { mode: r && r.mode, dirReq: window.__fo.dirReq, fileReq: window.__fo.fileReq };
  });
  ok(s1.mode && s1.mode !== "denied", "save to the inventory workbook goes ahead (" + s1.mode + ")");
  ok(s1.dirReq === 1 && s1.fileReq === 0, "…after exactly ONE folder prompt and no file prompt  " + JSON.stringify(s1));
  const s2 = await p.evaluate(async () => {
    const r = await QBR.fsReloadFromFile("RISKY_USERS_AND_DOMAIN.xlsx", { quiet: true });
    return { r, dirReq: window.__fo.dirReq, fileReq: window.__fo.fileReq };
  });
  ok(s2.r === true && s2.dirReq === 1 && s2.fileReq === 0, "another workbook in the folder needs no further prompt  " + JSON.stringify(s2));
  ok(await p.evaluate(() => !document.querySelector('.fo-bar [data-fo-act="reconnect"]')), "Reconnect button disappears once granted");

  console.log("== reconnect button + reload all");
  await p.evaluate(() => window.__fo.newSession());
  await p.evaluate(() => QBR.fsFolderRefreshState()); await p.waitForTimeout(150);
  await p.evaluate(() => document.querySelector('.fo-bar [data-fo-act="reconnect"]').click()); await p.waitForTimeout(400);
  ok(await p.evaluate(() => window.__fo.dirReq === 2 && window.__fo.fileReq === 0), "Reconnect = one folder prompt, zero file prompts");
  // Excel changes a file on disk → Reload all picks it up
  await p.evaluate(() => { const k = "Audit/RISKY_USERS_AND_DOMAIN.xlsx"; window.__fo.store[k].t = 99999; });
  await p.evaluate(() => document.querySelector('.fo-bar [data-fo-act="reload"]').click()); await p.waitForTimeout(2500);
  ok(await p.evaluate(() => (QBR._fsLinks || []).find(l => l.name === "RISKY_USERS_AND_DOMAIN.xlsx").lastModified === 99999), "Reload all re-reads the changed workbook from the folder");
  ok(await p.evaluate(() => /Reloaded/.test((document.getElementById("fo-toast") || {}).textContent || "")), "toast confirms the reload");

  console.log("== reopen same folder: remembers last picks");
  await p.evaluate(() => document.querySelector('.fo-bar [data-fo-act="open"]').click()); await p.waitForSelector(".fo-ov");
  const again = await p.evaluate(() => [...document.querySelectorAll(".fo-pick")].filter(c => c.checked).map(c => c.closest("tr").cells[2].textContent.trim()).sort());
  ok(JSON.stringify(again) === '["Audit","Inventory","Tracker"]', "re-opening pre-ticks last time's choice  " + JSON.stringify(again));
  await p.click("#fo-cancel");
  ok(await p.evaluate(() => (QBR._fsDirs || []).length === 1), "same folder is stored once");

  console.log("== folder declined → no per-file nagging");
  const dn = await p.evaluate(async () => {
    window.__fo.newSession(); window.__fo.denyDir = true;
    const before = window.__fo.fileReq;
    const r = await QBR.fsReloadFromFile("RISKY_USERS_AND_DOMAIN.xlsx", { quiet: true });
    window.__fo.denyDir = false;
    return { r, extra: window.__fo.fileReq - before };
  });
  ok(dn.r === false && dn.extra === 0, "declining the folder prompt stops there (no prompt per file)  " + JSON.stringify(dn));

  console.log("== unlink still works");
  await p.evaluate(() => QBR.fsUnlink("Lenovo Inventory.xlsx", true)); await p.waitForTimeout(150);
  ok(await p.evaluate(() => (QBR._fsLinks || []).length === 2), "unchecking one workbook unlinks only that one");

  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  ok(ext.length === 0, "no external requests (offline)");
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
