// Persist-family lineage warning test (live-edit safeguard, fix 3). Headless Chromium, file://.
// Flow (ui-save harness pattern, mocked File System Access handles):
//   1. load the synthetic inventory workbook, link a mock handle, edit, Save to Excel
//      → qbr-lineage-v1 records the saved version (fp + hash + savedAt).
//   2. edit again → journal holds a pending op against the SAVED fingerprint.
//   3. link a DIFFERENT-content copy under the same name via fsLinkFile (stubbed picker)
//      → the assets badge shows the "differs from the version last saved" warning.
//   4. link the same-content copy again → no warning.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-lineage.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const FIX = path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

(async () => {
  const b = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : {});
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(500);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); localStorage.removeItem("qbr-lineage-v1"); } catch (e) {} });
  await p.setInputFiles("#file-input", FIX); await p.waitForTimeout(2500);
  const name = path.basename(FIX); const b64 = fs.readFileSync(FIX).toString("base64");
  const res = await p.evaluate(async ({ name, b64 }) => {
    const out = {};
    const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    let bytes = bin, mtime = 1700000000000, writes = 0;
    const mk = () => new File([bytes], name, { lastModified: mtime });
    const f0 = APP.files.find(f => f.name === name).blob;
    const handle = {
      name, kind: "file",
      queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => mk(),
      createWritable: async () => { const parts = []; return { write: async d => parts.push(new Uint8Array(d)), close: async () => { bytes = parts[0]; mtime += 1000; writes++; } }; },
    };
    QBR._fsLinks = [{ fp: QBR.fpOf(name, f0), kinds: ["assets"], name, size: bin.length, lastModified: mtime, handle }];
    const inv = QBR._invModel || (typeof invModel === "function" ? invModel() : null);
    const a = inv.assets[0];
    QBR.invUpdateAsset(a.key, { client: "LINEAGE TEST SCHOOL" });
    const r = await QBR.fsSaveKind("assets");
    out.saveMode = r.mode; out.writes = writes;
    // 1. lineage recorded by the save
    let rec = null;
    try { rec = (JSON.parse(localStorage.getItem("qbr-lineage-v1") || "{}"))[name] || null; } catch (e) {}
    out.rec = rec;
    out.linkFp = QBR._fsLinks[0].fp; out.linkHash = QBR._fsLinks[0].hash;
    // 2. a fresh edit pends against the SAVED fingerprint
    QBR.invUpdateAsset(a.key, { client: "LINEAGE TEST SCHOOL 2" });
    out.pending = 0;
    try { out.pending = QBR.journalOpsFor(rec.fp).length; } catch (e) {}
    out.badgeExists = !!document.querySelector('[data-unsaved="assets"]');
    // 3. link a different-content copy (extra sheet — still a valid workbook)
    const wb = XLSX.read(bytes.slice(), { type: "array" });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["probe"]]), "Lineage probe");
    const diffBytes = new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
    const mkHandle = (fileBytes, mt) => ({
      name, kind: "file",
      queryPermission: async () => "granted", requestPermission: async () => "granted",
      getFile: async () => new File([fileBytes], name, { lastModified: mt }),
      createWritable: async () => ({ write: async () => {}, close: async () => {}, abort: async () => {} }),
    });
    window.showOpenFilePicker = async () => [mkHandle(diffBytes, mtime + 9999)];
    out.linkedDiff = await QBR.fsLinkFile();
    const badge = document.querySelector('[data-unsaved="assets"]');
    out.warnText = badge ? badge.textContent : "";
    // 4. link the same-content copy (the bytes this browser saved)
    if (badge) badge.textContent = "";
    window.showOpenFilePicker = async () => [mkHandle(bytes, mtime + 5000)];
    out.linkedSame = await QBR.fsLinkFile();
    out.sameText = badge ? badge.textContent : "";
    return out;
  }, { name, b64 });
  ok(res.saveMode === "file" && res.writes === 1, "setup: edit saved into the linked file");
  ok(!!res.rec && !!res.rec.savedAt, "lineage recorded on save (qbr-lineage-v1)");
  ok(res.rec && res.rec.fp === res.linkFp && res.rec.hash === res.linkHash, "lineage fp + hash match the saved version");
  ok(res.pending > 0, "journal holds a pending op against the saved fingerprint (" + res.pending + ")");
  ok(res.badgeExists, "assets unsaved badge present (persistNote target)");
  ok(res.linkedDiff === true && /differs from the version last saved in this browser/.test(res.warnText),
    "different-content copy linked → warning fires: " + (res.warnText || "").trim().slice(0, 110));
  ok(res.linkedSame === true && !/differs from the version last saved/.test(res.sameText),
    "same-content copy linked → no warning");
  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("BROWSER ERROR:", e); process.exit(2); });
