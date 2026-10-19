// 2026-10-13 overnight review: regression checks for the fixes made that night. Headless Chromium, file://.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-review-1013.cjs   (run also with TZ=Asia/Manila)
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };
(async () => {
  const b = await chromium.launch();
  const p = await (await b.newContext({ timezoneId: "Asia/Manila", viewport: { width: 1400, height: 900 } })).newPage();
  const errs = []; p.on("pageerror", e => errs.push(e.message));
  let dialogs = 0; p.on("dialog", d => { dialogs++; d.dismiss(); });
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(300);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [path.join(__dirname, "fixture.xlsx"), path.join(__dirname, "SAMPLE_Lenovo_Inventory_SYNTH.xlsx")]); await p.waitForTimeout(2500);

  console.log("== XSS");
  const x = await p.evaluate(() => {
    const evil = '<img src=x onerror="window.__xss=1">';
    APP.model.risky.forEach(r => { r.schoolRaw = evil; });
    goToTab("dash-report"); if (typeof renderReport === "function") renderReport();
    const imgs = document.querySelectorAll("#report-body img").length;
    QBR.persistNote("assets", "Linked ✓ " + evil, 50);
    const noteImgs = document.querySelectorAll("[data-unsaved] img").length;
    return { imgs, noteImgs, text: /&lt;img|<img/.test(document.getElementById("report-body").textContent) };
  });
  await p.waitForTimeout(200);
  ok(x.imgs === 0 && x.text, "report: a school named <img onerror> shows as text, no element created");
  ok(x.noteImgs === 0 && !(await p.evaluate(() => window.__xss)), "save note: a crafted file name can't inject HTML");

  console.log("== dates / warranty");
  const d = await p.evaluate(() => {
    const A = (s, y) => { const r = QBR.invAddYears(new Date(s[0], s[1], s[2]), y); return r.getFullYear() + "-" + (r.getMonth() + 1) + "-" + r.getDate(); };
    const t = APP.model.inventory.tickets.find(x => x.repEdtech instanceof Date);
    let shown = null;
    if (t) { openTicket360(t.tno); QBR._invT360Edit = true; renderTicket360Panel(); const i = document.getElementById("t360e-edtech"); shown = i && i.value; }
    const want = t ? t.repEdtech.getFullYear() + "-" + String(t.repEdtech.getMonth() + 1).padStart(2, "0") + "-" + String(t.repEdtech.getDate()).padStart(2, "0") : null;
    return { jan31: A([2026, 0, 31], 3), feb29: A([2024, 1, 29], 1), half: A([2026, 0, 15], 1.5), shown, want };
  });
  ok(d.jan31 === "2029-1-31" && d.feb29 === "2025-2-28" && d.half === "2027-7-15", "warranty end: same day / month-end / 1.5 years  " + JSON.stringify(d));
  ok(d.shown === null || d.shown === d.want, `Ticket 360 edit shows the stored date in UTC+8 (${d.shown} = ${d.want})`);

  console.log("== audit paste");
  const a = await p.evaluate(() => [QBR.auditParsePaste("Diocese A\t\tHealthy\thttps://x/ev"), QBR.auditParsePaste('"Pinewood Catholic School, Inc.",12,Healthy'), QBR.auditParsePaste("\t\t")]);
  ok(JSON.stringify(a[0].values) === JSON.stringify(["Diocese A", "", "Healthy", "https://x/ev"]), "blank Risky cell keeps Health in its column");
  ok(a[1].values && a[1].values[0] === "Pinewood Catholic School, Inc." && a[1].values[1] === "12", "quoted comma in a CSV name stays one value");
  ok(a[2].error === "empty", "all-blank paste is refused");

  console.log("== To-Do / General tickets");
  const g = await p.evaluate(() => {
    if (!QBR.users.ensure()) QBR.users.register("RevA");
    const sk = Object.keys(localStorage);
    APP.model.inventory.tickets.push({ tno: "PF12345678", issue: "Case number only", status: "Open", sn: "X1" });
    const t = QBR.todo.all().find(x => x.tno === "PF12345678");
    const pv = QBR.gtickets && QBR.gtickets.previewImport ? QBR.gtickets.previewImport([{ date: "2026-10-01", time: "", title: "Same", requester: "" }, { date: "2026-10-01", time: "", title: "Same", requester: "" }]) : null;
    return { age: t ? t.age : "missing", pv: pv && pv.map(x => x.dupe) };
  });
  ok(g.age === null, "numeric case number is not read as a date (age unknown, not 287,807 days)");
  ok(!g.pv || JSON.stringify(g.pv) === "[false,true]", "same row twice in one import file → second marked duplicate");

  console.log("== deploy duplicate serial");
  const dd = await p.evaluate(() => {
    const inv = APP.model.inventory, a = inv.assets.find(x => !x.client);
    if (!a) return "skip";
    const twin = Object.assign({}, a, { sn: a.sn.toLowerCase(), _src: null });
    inv.assets.push(twin);
    QBR.invDeploy([a.sn], "Review School", new Date(2026, 9, 1), null);
    const first = inv.assets.find(x => x.key === a.key);
    const r = first.client === "Review School" && !twin.client;
    inv.assets.pop();
    return r;
  });
  ok(dd === true || dd === "skip", "deploy updates the same row Asset 360 shows (first match)");
  ok(errs.length === 0, "no page errors  " + JSON.stringify(errs.slice(0, 3)));
  await b.close();
  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
