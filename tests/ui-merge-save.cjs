// Merge-on-save, end-to-end in the real app (headless Chromium, file://).
// A mocked Direct-save file handle lets the test change the file "in Excel"
// between dashboard edits and Save.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-merge-save.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = globalThis.XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const S = require(path.join(APPDIR, "js/xlsx-surgical.js"));
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
const INV = path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"), AUD = path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx");
const read = b => XLSX.read(b, { type: "array", cellStyles: true });
const b64 = u8 => Buffer.from(u8).toString("base64");
const DEV = "02 DEVICES";

// "Excel" edits: a small in-place change (keeps formatting) or a full rewrite (inserted column)
async function excelCell(bytes, sheet, ref, v) { const b = read(bytes), w = read(bytes); w.Sheets[sheet][ref] = { t: typeof v === "number" ? "n" : "s", v }; return (await S.surgicalSave(bytes, b, w)).bytes; }
function excelInsertCol(bytes, sheet, at, header) {
  const wb = XLSX.read(bytes, { type: "array" }); const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  wb.Sheets[sheet] = XLSX.utils.aoa_to_sheet(aoa.map((r, i) => { const x = r.slice(); x.splice(at, 0, i ? "n" + i : header); return x; }));
  return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
}
const colOf = (wb, sheet, h) => XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1 })[0].findIndex(x => String(x) === h);
const cellBy = (wb, sheet, keyH, key, h) => { const a = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null }); const r = a.find(x => x && String(x[colOf(wb, sheet, keyH)]) === key); return r ? r[colOf(wb, sheet, h)] : undefined; };

async function open(b, file) {
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => { try { localStorage.clear(); localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", file); await p.waitForTimeout(2500);
  const name = path.basename(file);
  await p.evaluate(({ name, data }) => {
    const bin = Uint8Array.from(atob(data), c => c.charCodeAt(0));
    const T = window.__T = { bytes: bin, mtime: 1700000000000, writes: 0, alerts: 0, downloads: 0, lock: false, onGet: null, gets: 0 };
    window.alert = () => { T.alerts++; };
    const oldW = XLSX.writeFile; XLSX.writeFile = function () { T.downloads++; };
    window.__setFile = b64 => { T.bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0)); T.mtime += 1000; };
    window.__getFile = () => { let s = ""; for (let i = 0; i < T.bytes.length; i += 32768) s += String.fromCharCode.apply(null, T.bytes.subarray(i, i + 32768)); return btoa(s); };
    const handle = { name, kind: "file", queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => { T.gets++; if (T.onGet) { const f = T.onGet; T.onGet = null; f(T.gets); } return new File([T.bytes], name, { lastModified: T.mtime }); },
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), abort: async () => {},
        close: async () => { if (T.lock) throw new DOMException("The file is open in another program", "NoModificationAllowedError"); T.bytes = parts[0]; T.mtime += 1000; T.writes++; } }; } };
    const f0 = APP.files.find(f => f.name === name).blob;
    QBR._fsLinks = [{ fp: QBR.fpOf(name, f0), kinds: [...((QBR._kindByFp || {})[QBR.fpOf(name, f0)] || [])], name, size: bin.length, lastModified: T.mtime, hash: QBR.fsHashOf(bin), handle }];
  }, { name, data: b64(fs.readFileSync(file)) });
  return { p, errs, name };
}
const fileNow = async p => new Uint8Array(Buffer.from(await p.evaluate(() => window.__getFile()), "base64"));
const pending = p => p.evaluate(() => { const s = QBR.journalEntries(); return Object.values(s).reduce((n, e) => n + ((e && e.ops) || []).length, 0); });

(async () => {
  const b = await chromium.launch();
  const inv0 = new Uint8Array(fs.readFileSync(INV)), aoa = XLSX.utils.sheet_to_json(read(inv0).Sheets[DEV], { header: 1, defval: null });
  const SN1 = String(aoa[1][0]), SN2 = String(aoa[2][0]);

  console.log("== A. Excel fixed another unit; dashboard edited one → merged, no prompt");
  let { p, errs, name } = await open(b, INV);
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "DASH SCHOOL" }); });
  await p.evaluate(d => window.__setFile(d), b64(await excelCell(inv0, DEV, "C3", "EXCEL-MODEL")));
  let r = await p.evaluate(async () => { const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, merged: r.merged || null, alerts: __T.alerts, downloads: __T.downloads, writes: __T.writes }; });
  ok(r.mode === "file" && r.merged && r.alerts === 0 && r.downloads === 0 && r.writes === 1, "saved into the file with no alert / no download (" + JSON.stringify(r) + ")");
  let wb = read(await fileNow(p));
  ok(cellBy(wb, DEV, "Serial Number", SN1, "Client / Organization") === "DASH SCHOOL" && cellBy(wb, DEV, "Serial Number", SN2, "Model") === "EXCEL-MODEL", "file has the dashboard edit AND the Excel edit");
  ok(await p.evaluate(sn => (APP.model.inventory.assets.find(a => a.sn === sn) || {}).model, SN2) === "EXCEL-MODEL", "dashboard reloaded: shows the Excel edit");
  ok(await pending(p) === 0, "journal cleared");
  // second round in the same session (base = last save)
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "For Repair" }); });
  await p.evaluate(async d => window.__setFile(d), b64(await excelCell(await fileNow(p), DEV, "C4", "EXCEL-2")));
  r = await p.evaluate(async () => (await QBR.fsSaveKind("assets")).mode);
  wb = read(await fileNow(p));
  ok(r === "file" && cellBy(wb, DEV, "Serial Number", SN1, "Condition") === "For Repair" && wb.Sheets[DEV].C4.v === "EXCEL-2" && wb.Sheets[DEV].C3.v === "EXCEL-MODEL", "second merged save in the same session works");
  const cf0 = (await S.zipText(S.zipRead(inv0), "xl/worksheets/sheet2.xml")).match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/)[0];
  ok((await S.zipText(S.zipRead(await fileNow(p)), "xl/worksheets/sheet2.xml")).includes(cf0), "formatting (conditional formatting) still intact");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== B. Same cell changed in Excel and dashboard → review");
  ({ p, errs } = await open(b, INV));
  const clientRef = XLSX.utils.encode_cell({ r: 1, c: colOf(read(inv0), DEV, "Client / Organization") });
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "DASH SCHOOL" }); });
  await p.evaluate(d => window.__setFile(d), b64(await excelCell(inv0, DEV, clientRef, "EXCEL SCHOOL")));
  r = await p.evaluate(async () => { QBR._mergeAutoResolve = m => { window.__seen = m.conflicts.map(c => [c.theirs, c.mine]); return null; }; const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, writes: __T.writes, seen: window.__seen }; });
  ok(r.mode === "merge-cancelled" && r.writes === 0 && JSON.stringify(r.seen) === JSON.stringify([["EXCEL SCHOOL", "DASH SCHOOL"]]), "conflict shown (Excel vs dashboard); Cancel writes nothing");
  ok(await pending(p) > 0, "edits kept after Cancel");
  r = await p.evaluate(async () => { QBR._mergeAutoResolve = () => ({ 0: "mine" }); return (await QBR.fsSaveKind("assets")).mode; });
  ok(r === "file" && cellBy(read(await fileNow(p)), DEV, "Serial Number", SN1, "Client / Organization") === "DASH SCHOOL", "choosing the dashboard value saves it");
  // real dialog renders
  await p.evaluate(() => { delete QBR._mergeAutoResolve; window.__dlg = QBR.fsMergeReview({ applied: 2, conflicts: [{ sheet: "02 DEVICES", ref: "B2", header: "client", key: "PF1", theirs: "A & <b>", mine: "B" }], unresolved: [{ sheet: "02 DEVICES", ref: "M9", mine: "x", why: "row deleted" }] }, "f.xlsx"); });
  ok(await p.locator(".merge-box").isVisible() && (await p.locator(".merge-box").innerText()).includes("A & <b>"), "review dialog renders values as text (escaped)");
  await p.click("[data-all=mine]"); await p.click("#merge-ok");
  ok(JSON.stringify(await p.evaluate(() => window.__dlg)) === JSON.stringify({ 0: "mine" }), "dialog returns the choices ('Keep all dashboard values')");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== C. Excel inserted a column; dashboard edit lands under its header");
  ({ p, errs } = await open(b, INV));
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "DASH SCHOOL", batch: "B-20261006-09" }); });
  await p.evaluate(d => window.__setFile(d), b64(excelInsertCol(inv0, DEV, 1, "Notes")));
  r = await p.evaluate(async () => { const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, merged: r.merged }; });
  wb = read(await fileNow(p));
  ok(r.mode === "file" && r.merged && r.merged.remapped, "merged with remapped columns");
  ok(cellBy(wb, DEV, "Serial Number", SN1, "Client / Organization") === "DASH SCHOOL" && cellBy(wb, DEV, "Serial Number", SN1, "Notes") === "n1" && cellBy(wb, DEV, "Serial Number", SN1, "Batch Code") === "B-20261006-09", "client under 'Client', Excel's 'Notes' kept, 'Batch Code' added");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== D. File open in desktop Excel (locked)");
  ({ p, errs } = await open(b, INV));
  await p.evaluate(() => { __T.lock = true; const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "LOCK TEST" }); });
  r = await p.evaluate(async () => { const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, writes: __T.writes, downloads: __T.downloads, note: (document.body.textContent.match(/open in Excel[^.]*\.[^.]*\./) || [""])[0] }; });
  ok(r.mode === "locked" && r.writes === 0 && r.downloads === 0, "locked → nothing written, no download (" + r.mode + ")");
  ok(/Close it there/.test(r.note), "message tells the user to close it in Excel");
  ok(await pending(p) > 0, "edits kept");
  r = await p.evaluate(async () => { __T.lock = false; return (await QBR.fsSaveKind("assets")).mode; });
  ok(r === "file" && cellBy(read(await fileNow(p)), DEV, "Serial Number", SN1, "Client / Organization") === "LOCK TEST", "after closing Excel, Save works");
  // locked during a MERGED save too
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { cond: "Locked merge" }); });
  await p.evaluate(async d => window.__setFile(d), b64(await excelCell(await fileNow(p), DEV, "C5", "X")));
  r = await p.evaluate(async () => { __T.lock = true; const r = await QBR.fsSaveKind("assets"); __T.lock = false; return r.mode; });
  ok(r === "locked" && await pending(p) > 0, "locked during a merged save → edits kept");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== E. Reload from file (↻) without the file picker");
  ({ p, errs } = await open(b, INV));
  await p.evaluate(d => window.__setFile(d), b64(await excelCell(inv0, DEV, "C3", "RELOADED-MODEL")));
  r = await p.evaluate(async name => { let picker = 0; const oc = HTMLInputElement.prototype.click; HTMLInputElement.prototype.click = function () { if (this.type === "file") picker++; else oc.call(this); };
    document.querySelector("#loaded-toggle") && document.querySelector("#loaded-toggle").click();
    const btn = [...document.querySelectorAll(".loaded-file-r")].find(x => x.closest(".loaded-file").textContent.includes(name)); btn.click();
    await new Promise(r => setTimeout(r, 2500)); HTMLInputElement.prototype.click = oc;
    return { picker, model: (APP.model.inventory.assets[1] || {}).model }; }, name);
  ok(r.picker === 0 && r.model === "RELOADED-MODEL", "↻ reloads the linked file through its link (no picker); Excel edit shown");
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "AFTER RELOAD" }); });
  r = await p.evaluate(async () => { const r = await QBR.fsSaveKind("assets"); return { mode: r.mode, merged: !!r.merged }; });
  ok(r.mode === "file" && !r.merged && cellBy(read(await fileNow(p)), DEV, "Serial Number", SN1, "Client / Organization") === "AFTER RELOAD", "save after reload is a normal save (file is the new base)");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== F. File changes DURING the save (sync) → re-merged");
  ({ p, errs } = await open(b, INV));
  const syncBytes = await excelCell(inv0, DEV, "C6", "SYNCED-IN");
  await p.evaluate(() => { const a = APP.model.inventory.assets[0]; QBR.invUpdateAsset(a.key, { client: "RACE" }); });
  r = await p.evaluate(async d => { __T.onGet = n => { window.__setFile(d); }; const r = await QBR.fsSaveKind("assets"); return r.mode; }, b64(syncBytes));
  wb = read(await fileNow(p));
  ok(r === "file" && wb.Sheets[DEV].C6.v === "SYNCED-IN" && cellBy(wb, DEV, "Serial Number", SN1, "Client / Organization") === "RACE", "change that arrived mid-save is kept, and the edit is saved");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== G. Audit workbook: exempt + risky edit while Excel fixed another school");
  ({ p, errs, name } = await open(b, AUD));
  const aud0 = new Uint8Array(fs.readFileSync(AUD));
  const res = await p.evaluate(() => { const a = QBR._audit, ms = Object.values(a.months)[0]; QBR.auditUpdateCell(a.fp, ms.sheet, ms.rows[0].r, "risky", "77"); QBR.auditSetExempt(a.fp, ms.rows[0].school, true, "No GDAP", "month"); return { sheet: ms.sheet, s0: ms.rows[0].school, s1: ms.rows[1].school }; });
  await p.evaluate(d => window.__setFile(d), b64(excelInsertCol(aud0, res.sheet, 2, "CASES")));
  r = await p.evaluate(async () => { const r = await QBR.auditFamilySave("audit"); return { mode: r.mode, merged: r.merged }; });
  wb = read(await fileNow(p));
  ok(r.mode === "file" && r.merged, "audit save merged (" + JSON.stringify(r.merged) + ")");
  ok(String(cellBy(wb, res.sheet, "SCHOOL", res.s0, "TOTAL RISKY USERS")) === "77" && cellBy(wb, res.sheet, "SCHOOL", res.s0, "CASES") === "n1", "risky count under its header; Excel's new column kept");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  console.log("== H. Row deleted in Excel → that edit is parked (v1.30.0), the rest saved");
  ({ p, errs } = await open(b, INV));
  await p.evaluate(sns => { const A = APP.model.inventory.assets; QBR.invUpdateAsset(A.find(a => a.sn === sns[0]).key, { client: "KEEP ME" }); QBR.invUpdateAsset(A.find(a => a.sn === sns[1]).key, { cond: "ON A DELETED ROW" }); }, [SN1, SN2]);
  { const wbd = XLSX.read(inv0, { type: "array" }); const a2 = XLSX.utils.sheet_to_json(wbd.Sheets[DEV], { header: 1, defval: null }); a2.splice(2, 1);
    wbd.Sheets[DEV] = XLSX.utils.aoa_to_sheet(a2); await p.evaluate(d => window.__setFile(d), b64(new Uint8Array(XLSX.write(wbd, { bookType: "xlsx", type: "array" })))); }
  r = await p.evaluate(async () => { QBR._mergeAutoResolve = m => { window.__un = m.unresolved.length; return {}; }; const r = await QBR.fsSaveKind("assets"); delete QBR._mergeAutoResolve;
    return { mode: r.mode, un: window.__un, parked: QBR.journalParked().map(x => x.value + "|" + x.why) }; });
  ok(r.mode === "file" && cellBy(read(await fileNow(p)), DEV, "Serial Number", SN1, "Client / Organization") === "KEEP ME", "placeable edit saved");
  ok(r.un >= 1 && r.parked.some(x => /ON A DELETED ROW/.test(x)), "unplaceable edit kept under 'Couldn't be placed' (" + r.parked.join("; ") + ")");
  ok(await pending(p) === 0, "pending journal cleared after the save");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : "")); await p.close();

  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
