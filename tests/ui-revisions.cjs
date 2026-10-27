// Regression tests for the 2026-10-27 revision list (feature/revisions-2026-10). Dummy data only.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-revisions.cjs
//  A. Guided audit: newest storage/usage workbook wins; a miss says which file/sheet was searched.
//  B. Guided audit step 1: the school field accepts pick OR type (partial, case-insensitive).
//  C. Inventory: PF-4XNZNY and PF4XNZNY are the same serial (key, intake, look-up, filter, tickets search).
//  D. Scan: unreadable MTM falls back to the Inventory record; empty MTM shows "not read", not "MTM".
//  E. User Management: "Updated Schools by Organization" is its own full-width, horizontal, sorted chart.
//  F. Google Postmaster: dropdown filters (verification, DNS, reputation, organization) built from the data.
//  G. My To-Do / General tickets: re-open a saved "<Name>-todo.xlsx" (merge, no duplicates, no deletions).
//  H. Loaded files: grouped, searchable pop-up window (close by ×, Esc, backdrop); ↻/× still work.
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(APPDIR, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

function storageBook(schoolsByQ) {
  const wb = XLSX.utils.book_new();
  ["Q1", "Q2", "Q3", "Q4"].forEach(q => {
    const rows = [["School (Rakso Managed)", "ORGANIZATION", "ONEDRIVE", "EXCHANGE", "SHAREPOINT", "CURRENT STORAGE", "USAGE", "USED", "TOTAL"]];
    (schoolsByQ[q] || []).forEach(s => rows.push([s, "ORG-A", null, null, null, null, null, null, null]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), q);
  });
  return Buffer.from(XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
}
const b64 = buf => buf.toString("base64");

async function boot(b) {
  const p = await b.newPage(); p._errs = [];
  p.on("pageerror", e => p._errs.push(e.message));
  await p.goto(APP); await p.waitForTimeout(400);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  return p;
}
// Load files with explicit lastModified (setInputFiles can't set it).
async function load(p, files) {
  await p.evaluate(async fl => {
    const list = fl.map(f => new File([Uint8Array.from(atob(f.b64), c => c.charCodeAt(0))], f.name, { lastModified: f.mtime }));
    handleFiles(list);
  }, files);
  await p.waitForTimeout(3000);
}

(async () => {
  const b = await chromium.launch();
  const AUD = fs.readFileSync(path.join(__dirname, "SAMPLE_AUDIT_RICH.xlsx"));
  const all = ["Alpha Test School", "Beta Test Academy", "Gamma Test College"];
  const NEW = storageBook({ Q1: all, Q2: all, Q3: all, Q4: all });
  const OLD = storageBook({ Q1: all, Q2: all, Q3: all, Q4: ["Alpha Test School"] });

  console.log("== A. guided audit: storage workbook choice + miss message");
  {
    const p = await boot(b);
    // the OLD copy is read last (old behaviour: last read wins) but is older on disk
    await load(p, [
      { name: "SAMPLE_AUDIT_RICH.xlsx", b64: b64(AUD), mtime: Date.UTC(2026, 9, 20) },
      { name: "STORAGE REPORT.xlsx", b64: b64(NEW), mtime: Date.UTC(2026, 9, 26) },
      { name: "STORAGE REPORT (old copy).xlsx", b64: b64(OLD), mtime: Date.UTC(2026, 6, 1) },
    ]);
    const r = await p.evaluate(() => ({
      primary: QBR._storage && QBR._storage.fileName,
      alts: QBR._storage && (QBR._storage.alternates || []).map(m => m.fileName),
      found: !!QBR.quarterFindRow(QBR._storage, "Q4", "Beta Test Academy"),
    }));
    ok(r.primary === "STORAGE REPORT.xlsx", "newest storage workbook is used (" + r.primary + ")");
    ok(r.alts.length === 1 && /old copy/.test(r.alts[0]), "older copy kept as an alternate");
    ok(r.found, "Q4 row found for a school missing only from the old copy");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }
  {
    const p = await boot(b);
    await load(p, [
      { name: "SAMPLE_AUDIT_RICH.xlsx", b64: b64(AUD), mtime: Date.UTC(2026, 9, 20) },
      { name: "STORAGE REPORT (old copy).xlsx", b64: b64(OLD), mtime: Date.UTC(2026, 9, 27) },
      { name: "STORAGE REPORT.xlsx", b64: b64(NEW), mtime: Date.UTC(2026, 6, 1) },
    ]);
    const msg = await p.evaluate(() => QBR.quarterMissReason(QBR._storage, "Q4", "Beta Test Academy"));
    ok(/Searched <b>STORAGE REPORT \(old copy\)\.xlsx<\/b> › sheet <b>Q4<\/b> \(1 schools\)/.test(msg), "miss names the file + sheet searched");
    ok(/IS in another loaded copy: <b>STORAGE REPORT\.xlsx<\/b>/.test(msg), "miss points to the copy that has the row");
    // the wizard shows it
    const shown = await p.evaluate(async () => {
      document.querySelector('[data-tab="dash-audit"]').click();
      await new Promise(r => setTimeout(r, 200));
      QBR.auditWizardStart(); const w = QBR._auditWizard;
      w.month = "OCTOBER"; w.school = "Beta Test Academy"; w.step = 3;
      QBR.renderAudit();
      await new Promise(r => setTimeout(r, 200));
      const el = document.querySelector(".wiz-step"); return el ? el.textContent : "(no wizard step) " + (document.getElementById("dash-audit") || {}).textContent.slice(0, 200);
    });
    ok(/Searched STORAGE REPORT \(old copy\)\.xlsx › sheet Q4/.test(shown), "wizard storage step shows the reason  [" + shown.replace(/\s+/g, " ").slice(0, 160) + "]");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== B. guided audit: type or pick the school");
  {
    const p = await boot(b);
    await load(p, [{ name: "SAMPLE_AUDIT_RICH.xlsx", b64: b64(AUD), mtime: Date.UTC(2026, 9, 20) }]);
    await p.evaluate(async () => { document.querySelector('[data-tab="dash-audit"]').click(); await new Promise(r => setTimeout(r, 200)); QBR.auditWizardStart(); });
    await p.waitForTimeout(300);
    ok(await p.$eval("#wiz-school", el => el.tagName === "INPUT" && el.getAttribute("list") === "wiz-school-list"), "school field is a searchable input with a list");
    ok(await p.$$eval("#wiz-school-list option", os => os.length >= 3), "list holds the schools");
    const typed = async v => { await p.fill("#wiz-school", v); await p.dispatchEvent("#wiz-school", "change"); await p.waitForTimeout(250); return p.evaluate(() => ({ school: QBR._auditWizard.school, hint: (document.getElementById("wiz-school-hint") || {}).textContent || "" })); };
    let r = await typed("beta");
    ok(r.school === "Beta Test Academy", "partial, lower-case text picks the single match (" + r.school + ")");
    r = await typed("GAMMA TEST COLLEGE");
    ok(r.school === "Gamma Test College", "exact name in other case picks it");
    r = await typed("test");
    ok(r.school === "Gamma Test College" && /schools match/.test(r.hint), "ambiguous text keeps the current school and says how many match  [" + r.hint + "]");
    r = await typed("zzz");
    ok(/No school matches/.test(r.hint), "unknown text explains, no change");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== C. inventory: serial with or without dash");
  {
    const p = await boot(b);
    const INV = fs.readFileSync(path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"));
    await load(p, [{ name: "SAMPLE_Lenovo_Inventory_RICH.xlsx", b64: b64(INV), mtime: Date.UTC(2026, 9, 20) }]);
    const r = await p.evaluate(() => {
      const inv = APP.model.inventory, a = inv.assets.find(x => /^PF[0-9A-Z]{7}$/.test(x.sn));
      const dashed = a.sn.slice(0, 2) + "-" + a.sn.slice(2), spaced = " " + a.sn.slice(0, 2).toLowerCase() + " " + a.sn.slice(2) + " ";
      const n0 = inv.assets.length;
      const added = QBR.invIntake([{ sn: dashed, client: "Sample Academy", model: "ThinkPad E14" }]);
      return { sn: a.sn, dashed, same: QBR.invSerialKey(dashed) === a.key && QBR.invSerialKey(spaced) === a.key && QBR.invSerialKey("PF\u20114X") === QBR.invSerialKey("PF4X"),
        added, n0, n1: inv.assets.length, storedKept: a.sn };
    });
    ok(r.same, "PF-…, pf …, PF‑… (other dash) and PF… share one key");
    ok(r.added === 0 && r.n1 === r.n0, "intake of the dashed form does not create a duplicate unit");
    ok(r.storedKept === r.sn, "stored serial text is unchanged (" + r.sn + ")");
    // Inventory filter + Asset 360 look-up with the dashed form
    await p.evaluate(() => document.querySelector('[data-tab="dash-inventory"]').click());
    await p.waitForTimeout(300);
    await p.fill("#inv-f-q", r.dashed); await p.press("#inv-f-q", "Enter"); await p.waitForTimeout(400);
    ok(await p.evaluate(sn => document.getElementById("dash-inventory").textContent.includes(sn), r.sn), "Inventory serial filter finds " + r.sn + " when typed as " + r.dashed);
    const a360 = await p.evaluate(d => { location.hash = "#asset/" + encodeURIComponent(d); return new Promise(res => setTimeout(() => res((document.getElementById("dash-asset360") || {}).textContent || ""), 500)); }, r.dashed);
    ok(a360.includes(r.sn) && !/not found/i.test(a360.slice(0, 400)), "Asset 360 opens for the dashed serial");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== D. scan: MTM fallback + placeholder");
  {
    const p = await boot(b);
    const INV = fs.readFileSync(path.join(__dirname, "SAMPLE_Lenovo_Inventory_RICH.xlsx"));
    await load(p, [{ name: "SAMPLE_Lenovo_Inventory_RICH.xlsx", b64: b64(INV), mtime: Date.UTC(2026, 9, 20) }]);
    const r = await p.evaluate(async () => {
      const inv = APP.model.inventory;
      const known = inv.assets.find(x => /^PF/.test(x.sn)); known.mtm = "21M4S5VD00";
      const other = inv.assets.find(x => x !== known && !x.mtm);
      document.querySelector('[data-tab="dash-scan"]').click(); await new Promise(r => setTimeout(r, 300));
      const ui = QBR._scanUI;
      QBR.scanImage = async () => ({ barcodes: ["199271075067"], serial: ui._next, mtm: "", productKey: "", model: "ThinkPad E14 Gen 6" });
      const mk = (id, sn) => { ui.rows.push({ id, name: id + ".jpg", dataUrl: "data:image/png;base64,iVBORw0KGgo=", model: "", serial: "", mtm: "", productKey: "", barcodeRaw: "", status: "pending", dup: null }); ui._next = sn; return scanRowScan(id); };
      await mk("t1", known.sn.slice(0, 2) + "-" + known.sn.slice(2));
      await mk("t2", other ? other.sn : "PF00ZZZZ");
      scanRenderRows(); await new Promise(r => setTimeout(r, 100));
      const inp = id => document.querySelector('input[data-f="mtm"][data-id="' + id + '"]');
      return { m1: ui.rows.find(x => x.id === "t1").mtm, src1: ui.rows.find(x => x.id === "t1").mtmSrc, ph2: inp("t2") && inp("t2").placeholder, tip2: inp("t2") && inp("t2").title, v2: inp("t2") && inp("t2").value };
    });
    ok(r.m1 === "21M4S5VD00" && r.src1 === "inventory", "label without MTM -> MTM taken from the Inventory record (dashed serial too)");
    ok(r.v2 === "" && r.ph2 === "not read" && /not read from this photo/.test(r.tip2), "empty MTM shows 'not read' + a tip, not the word 'MTM'");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== E. user management: org chart layout");
  {
    const p = await boot(b);
    const FX = fs.readFileSync(path.join(__dirname, "fixture.xlsx"));
    // a small dummy USER MANAGEMENT workbook (SY sheet), 12 schools over 3 orgs
    const um = XLSX.utils.book_new();
    const rows = [["School", "Status", "Remarks"]];
    for (let i = 0; i < 12; i++) rows.push(["Dummy School " + i, i % 3 ? "Pending" : "Updated", ""]);
    XLSX.utils.book_append_sheet(um, XLSX.utils.aoa_to_sheet(rows), "SY 2026-2027");
    await load(p, [{ name: "fixture.xlsx", b64: b64(FX), mtime: Date.UTC(2026, 9, 20) },
                   { name: "USER MANAGEMENT.xlsx", b64: b64(Buffer.from(XLSX.write(um, { bookType: "xlsx", type: "buffer" }))), mtime: Date.UTC(2026, 9, 20) }]);
    await p.evaluate(() => document.querySelector('[data-tab="dash-usermgmt"]') && document.querySelector('[data-tab="dash-usermgmt"]').click());
    await p.waitForTimeout(500);
    const r = await p.evaluate(() => {
      const c = QBR._charts && QBR._charts["ch-um-org"], card = document.getElementById("ch-um-org").closest(".col-12");
      const labels = c ? c.data.labels : [], tot = labels.map((_, i) => c.data.datasets.reduce((s, d) => s + (d.data[i] || 0), 0));
      return { has: !!c, full: !!card, axis: c && c.options.indexAxis, sorted: tot.every((v, i) => !i || v <= tot[i - 1]), n: labels.length,
        sib: document.getElementById("ch-um-readiness").closest(".col-lg-6") && document.getElementById("ch-um-sy").closest(".col-lg-6") ? true : false };
    });
    ok(r.has && r.full && r.axis === "y", "org chart sits in its own full-width row as horizontal bars");
    ok(r.sorted && r.n >= 1, "organizations sorted largest first (" + r.n + ")");
    ok(r.sib, "the two remaining cards share the row above (half width each)");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== F. postmaster: dropdown filters");
  {
    const p = await boot(b);
    const FX = fs.readFileSync(path.join(__dirname, "fixture.xlsx"));
    await load(p, [{ name: "fixture.xlsx", b64: b64(FX), mtime: Date.UTC(2026, 9, 20) }]);
    await p.evaluate(() => document.querySelector('[data-tab="dash-postmaster"]').click());
    await p.waitForTimeout(500);
    const opts = await p.evaluate(() => ["pm-f-verify", "pm-f-dns", "pm-f-rep", "pm-f-org"].map(id => [...document.getElementById(id).options].map(o => o.value)));
    ok(opts.every(o => o[0] === "ALL" && o.length >= 2), "four dropdowns, each with All + values from the data  " + JSON.stringify(opts.map(o => o.length)));
    ok(opts[2].some(v => /Issues detected|Verify to see health/.test(v)), "reputation list offers Google's live states");
    const rows = () => p.$$eval("#tbl-pm tr", tr => tr.length);
    const total = await rows();
    const pick = opts[2].find(v => v !== "ALL");
    await p.selectOption("#pm-f-rep", pick); await p.waitForTimeout(200);
    const r = await p.evaluate(v => ({ n: document.querySelectorAll("#tbl-pm tr").length, cells: [...document.querySelectorAll("#tbl-pm tr td:nth-child(6)")].map(td => td.textContent.trim()), count: document.getElementById("pm-count").textContent }), pick);
    ok(r.n > 0 && r.n <= total && r.cells.every(c => c === pick || (pick === "No entry" && c === "—")), "reputation filter keeps only '" + pick + "' rows (" + r.n + "/" + total + ")");
    ok(/of \d+ domains match/.test(r.count), "count line shows filtered vs total  [" + r.count + "]");
    await p.evaluate(() => { const el = document.querySelector('[data-tab="dash-overview"]'); el && el.click(); });
    await p.evaluate(() => document.querySelector('[data-tab="dash-postmaster"]').click()); await p.waitForTimeout(300);
    ok(await p.$eval("#pm-f-rep", (el, v) => el.value === v, pick), "choice survives a re-render");
    await p.click("#pm-f-reset"); await p.waitForTimeout(200);
    ok(await rows() === total && await p.$eval("#pm-f-rep", el => el.value === "ALL"), "Reset clears all filters");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== G. to-do profile: open a saved file");
  {
    const p = await boot(b);
    const FX = fs.readFileSync(path.join(__dirname, "fixture.xlsx"));
    await load(p, [{ name: "fixture.xlsx", b64: b64(FX), mtime: Date.UTC(2026, 9, 20) }]);
    const r = await p.evaluate(async () => {
      Object.keys(localStorage).filter(k => /^qbr-user/.test(k)).forEach(k => localStorage.removeItem(k));
      QBR.users.register("Jane Roe");
      const d = QBR.users.myData();
      d.todos.push({ id: "t1", title: "Renew dummy domain", product: "general", due: "2026-11-02", priority: "high", notes: "call ORG-A", done: false, createdAt: "2026-10-20T00:00:00.000Z" });
      d.tickets.push({ tno: "GT-0001", title: "Projector not detected", requester: "ROE, Jane", client: "Sample Academy", priority: "High", status: "Open", due: "2026-10-30", serial: null, notes: [{ d: "2026-10-21", text: "asked for photo" }], createdAt: "2026-10-21T00:00:00.000Z" });
      d.history.push({ doneAt: "2026-10-19T00:00:00.000Z", kind: "todo", title: "Old task", product: "general" });
      QBR.users.saveMyData(d);
      let wbOut = null; const orig = XLSX.writeFile; XLSX.writeFile = (wb, fn) => { wbOut = { wb, fn }; }; QBR.users.exportWorkbook(); XLSX.writeFile = orig;
      const buf = new Uint8Array(XLSX.write(wbOut.wb, { bookType: "xlsx", type: "array" }));
      // 1. lost local data -> re-open restores it
      const u = QBR.users.current(); localStorage.setItem("qbr-userdata-" + u.id, JSON.stringify({ todos: [], tickets: [], history: [] }));
      const a = QBR.users.importWorkbook(buf, wbOut.fn); const d1 = QBR.users.myData();
      // 2. open the same file again -> updates, no duplicates
      const b2 = QBR.users.importWorkbook(buf, wbOut.fn); const d2 = QBR.users.myData(); const n2 = [d2.todos.length, d2.tickets.length, d2.history.length];
      // 3. a newer local item survives an import of an older file
      d2.todos.push({ id: "t2", title: "Local only", product: "general", done: false, createdAt: new Date().toISOString() }); QBR.users.saveMyData(d2);
      QBR.users.importWorkbook(buf, wbOut.fn); const d3 = QBR.users.myData();
      // 4. no user picked yet -> the name in the file name is registered
      Object.keys(localStorage).filter(k => /^qbr-user/.test(k)).forEach(k => localStorage.removeItem(k));
      const c = QBR.users.importWorkbook(buf, wbOut.fn);
      // 5. a random workbook is refused
      const junk = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(junk, XLSX.utils.aoa_to_sheet([["a"]]), "Sheet1");
      const j = QBR.users.importWorkbook(new Uint8Array(XLSX.write(junk, { bookType: "xlsx", type: "array" })), "x.xlsx");
      return { fn: wbOut.fn, a: a.counts, t1: d1.todos[0], k1: d1.tickets[0], h1: d1.history.length, b2: b2.counts, n2,
        n3: d3.todos.map(t => t.id), cUser: c.ok && c.user.name, cCounts: c.counts, j: j.ok, jErr: j.error };
    });
    ok(r.fn === "Jane-Roe-todo.xlsx", "export file name " + r.fn);
    ok(r.a.todosAdded === 1 && r.a.ticketsAdded === 1 && r.a.history === 1, "re-opened file restores to-dos, tickets and history " + JSON.stringify(r.a));
    ok(r.t1.title === "Renew dummy domain" && r.t1.due === "2026-11-02" && r.t1.priority === "high" && r.t1.notes === "call ORG-A", "to-do fields round-trip");
    ok(r.k1.tno === "GT-0001" && r.k1.status === "Open" && r.k1.notes.length === 1 && r.k1.notes[0].d === "2026-10-21" && r.k1.notes[0].text === "asked for photo", "ticket fields + dated notes round-trip");
    ok(r.b2.todosUpdated === 1 && r.b2.ticketsUpdated === 1 && r.b2.todosAdded === 0 && JSON.stringify(r.n2) === "[1,1,1]", "opening the same file twice updates, never duplicates " + JSON.stringify([r.b2, r.n2]));
    ok(r.n3.includes("t2") && r.n3.includes("t1"), "items only in the browser are kept (merge, no deletions)");
    ok(r.cUser === "Jane Roe" && r.cCounts.todosAdded === 1, "with no user picked, the name in the file name is registered");
    ok(r.j === false && /isn't a My To-Do export/.test(r.jErr), "a non-export workbook is refused with a clear message");
    await p.evaluate(() => document.querySelector('[data-tab="dash-todo"]') && document.querySelector('[data-tab="dash-todo"]').click()); await p.waitForTimeout(300);
    ok(await p.$('[data-todo-act="import"]') !== null, "My To-Do has an 'Open saved file…' button");
    await p.evaluate(() => document.querySelector('[data-tab="dash-gtickets"]') && document.querySelector('[data-tab="dash-gtickets"]').click()); await p.waitForTimeout(300);
    ok(await p.$('[data-gt-act="open"]') !== null, "General tickets has an 'Open saved file…' button");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  console.log("== H. loaded files pop-up");
  {
    const p = await boot(b);
    const FX = fs.readFileSync(path.join(__dirname, "fixture.xlsx"));
    const blank = () => { const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([["x"]]), "Sheet1"); return b64(Buffer.from(XLSX.write(w, { bookType: "xlsx", type: "buffer" }))); };
    const names = ["Sample Training Feedback Form - 01022026 - AM.xlsx", "Sample Training Feedback Form - 02022026 - PM.xlsx", "Sample Canva Onboarding 2026.xlsx",
      "Sample Printer Inventory.xlsx", "Sample DOMAIN REGISTRATION TRACKER.xlsx", "Misc notes.xlsx"];
    await load(p, [{ name: "fixture.xlsx", b64: b64(FX), mtime: 1 }].concat(names.map(n => ({ name: n, b64: blank(), mtime: 1 }))));
    await p.click("#loaded-toggle"); await p.waitForTimeout(200);
    const r = await p.evaluate(() => {
      const box = document.getElementById("loaded-chips"), cs = getComputedStyle(box);
      const groups = [...box.querySelectorAll(".loaded-group")].map(g => [g.dataset.group, +g.querySelector(".loaded-group-n").textContent]);
      return { pos: cs.position, open: !box.classList.contains("d-none"), scrim: !document.getElementById("loaded-scrim").classList.contains("d-none"), groups,
        dialog: box.getAttribute("role") === "dialog", focus: document.activeElement && document.activeElement.id };
    });
    ok(r.open && r.pos === "fixed" && r.scrim && r.dialog, "toggle opens a fixed pop-up window with a backdrop");
    const g = Object.fromEntries(r.groups);
    ok(g.trainers === 2 && g.canva === 1 && g.inv === 1 && g.domain === 1 && g.other >= 2, "files grouped by type with counts  " + JSON.stringify(r.groups));
    ok(r.focus === "loaded-search", "search box gets focus");
    await p.fill("#loaded-search", "feedback"); await p.waitForTimeout(150);
    const vis = await p.evaluate(() => [...document.querySelectorAll("#loaded-files .loaded-file")].filter(x => !x.classList.contains("d-none") && !x.closest(".loaded-group").classList.contains("d-none")).length);
    ok(vis === 2, "search narrows the list to matching files (" + vis + ")");
    await p.fill("#loaded-search", ""); await p.waitForTimeout(100);
    await p.keyboard.press("Escape"); await p.waitForTimeout(100);
    ok(await p.$eval("#loaded-chips", el => el.classList.contains("d-none")), "Escape closes it");
    await p.click("#loaded-toggle"); await p.click("#loaded-close"); await p.waitForTimeout(100);
    ok(await p.$eval("#loaded-chips", el => el.classList.contains("d-none")), "× closes it");
    await p.click("#loaded-toggle"); await p.mouse.click(5, 700); await p.waitForTimeout(100);
    ok(await p.$eval("#loaded-chips", el => el.classList.contains("d-none")), "clicking the backdrop closes it");
    // remove one file from inside a group still works
    await p.click("#loaded-toggle");
    await p.evaluate(() => { window.confirm = () => true; const row = [...document.querySelectorAll(".loaded-file")].find(x => /Misc notes/.test(x.textContent)); row.querySelector(".loaded-file-x").click(); });
    await p.waitForTimeout(2500);
    ok(await p.evaluate(() => !(APP.files || []).some(f => /Misc notes/.test(f.name))), "× on a grouped row still removes that workbook");
    ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
    await p.close();
  }

  await b.close();
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
