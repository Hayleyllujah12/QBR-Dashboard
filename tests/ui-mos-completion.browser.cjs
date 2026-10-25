// MOS completion certificates — headless Chromium, synthetic attempts only (dummy data).
// Failed-twice eligibility, MOS-only scope, current-filter batching, grouping,
// cert content (name format, titles, ordinal date, Expert no-badge), signatory
// persistence, and the issued-log separate column block.
// Usage: PW=~/workspace/.pwtest/node_modules CHROME=/opt/meta-chromium/chrome node tests/ui-mos-completion.browser.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path");
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }

// Invented fixture: Demo High + Second High, no real people or schools.
// A: Excel failed twice (export name, no roster) -> SANTOS, MARIA.
// B: Word failed twice, roster email match -> REYES, JOSE P. (roster wins over export name).
// C: Excel Expert failed twice -> Expert cert, NO badge.
// D: PowerPoint pass only -> excluded. E: AZ-900 failed twice -> excluded (non-MOS).
// H: Excel fail then pass on retake -> excluded (passed, no longer failed-twice).
// F: Second High Word failed twice -> grouping + school filter.
// G: Excel pending + single fail -> excluded (not failed twice).
function attempts() {
  const mk = (ref, school, exam, score, date, extra) => Object.assign({ ref, school, exam, score, name: "Sample " + ref, date }, extra || {});
  return [
    mk("A1", "Demo High", "Excel", 500, "2026-10-01", { name: "Maria Santos" }),
    mk("A2", "Demo High", "Excel", 600, "2026-10-20", { name: "Maria Santos" }),
    mk("B1", "Demo High", "Word", 400, "2026-10-02", { name: "Jose Reyes", email: "jose@example.test" }),
    mk("B2", "Demo High", "Word", 650, "2026-10-21", { name: "Jose Reyes", email: "jose@example.test" }),
    mk("C1", "Demo High", "Excel Expert", 500, "2026-10-03", { name: "Ana Cruz" }),
    mk("C2", "Demo High", "Excel Expert", 600, "2026-10-22", { name: "Ana Cruz" }),
    mk("D1", "Demo High", "PowerPoint", 800, "2026-10-04", { name: "Pedro Lim" }),
    mk("E1", "Demo High", "AZ-900", 300, "2026-10-06", { name: "Liza Tan" }),
    mk("E2", "Demo High", "AZ-900", 400, "2026-10-23", { name: "Liza Tan" }),
    mk("F1", "Second High", "Word", 500, "2026-10-07", { name: "Nina Park" }),
    mk("F2", "Second High", "Word", 600, "2026-10-24", { name: "Nina Park" }),
    mk("G1", "Demo High", "Excel", null, "2026-10-08", { name: "Omar Diaz" }),
    mk("G2", "Demo High", "Excel", 500, "2026-10-09", { name: "Omar Diaz" }),
    mk("H1", "Demo High", "Excel", 500, "2026-10-10", { name: "Rosa Vega" }),
    mk("H2", "Demo High", "Excel", 750, "2026-10-25", { name: "Rosa Vega" }),
  ];
}
function roster() {
  return [{ name: "REYES, JOSE P.", email: "jose@example.test", studentNo: "1", school: "Demo High", section: "" }];
}
async function bootMos(browser) {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx"));
  await page.waitForTimeout(2000);
  await page.click('[data-tab="dash-mos"]');
  await page.waitForTimeout(300);
  await page.evaluate(a => window.QBR.mos._loadAttempts(a, window.__rost || []), attempts());
  await page.waitForTimeout(500);
  return { ctx, page, errs };
}
async function bootMosRoster(browser) {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", e => errs.push(String(e.message)));
  await page.goto(APP, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx"));
  await page.waitForTimeout(2000);
  await page.click('[data-tab="dash-mos"]');
  await page.waitForTimeout(300);
  await page.evaluate(args => window.QBR.mos._loadAttempts(args.a, args.r), { a: attempts(), r: roster() });
  await page.waitForTimeout(500);
  return { ctx, page, errs };
}
async function gotoCompletion(page) {
  await page.click('[data-mos-view="completion"]');
  await page.waitForTimeout(500);
}
async function batchText(page) {
  return page.evaluate(() => (document.getElementById("mos-view-completion") || {}).textContent || "");
}

(async () => {
  const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME, args: ["--no-sandbox"] } : { args: ["--no-sandbox"] });

  // A. View + batch: 4 certs (A Excel, B Word, C Expert, F Word), grouped, badge count.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await gotoCompletion(page);
    const t = await batchText(page);
    ok(/4 certificates/.test(t), "batch: 4 certificates (got " + (t.match(/\d+ certificates/) || ["?"])[0] + ")");
    ok(/Demo High/.test(t) && /Second High/.test(t), "batch: both schools grouped");
    ok(/Excel 1/.test(t) && /Excel Expert 1/.test(t) && /Word 1/.test(t), "batch: Demo High groups Excel/Expert/Word x1");
    const badge = await page.evaluate(() => (document.getElementById("mos-comp-n") || {}).textContent || "");
    ok(badge === "4", "viewbar badge: 4 (got " + badge + ")");
    ok(await page.evaluate(() => !!document.getElementById("mos-comp-print") && !!document.getElementById("mos-comp-pdf") && !document.getElementById("mos-comp-print").disabled), "Print + Save as PDF present and enabled");
    ok(errs.length === 0, "section A: zero page errors (got " + errs.length + ")");
    await ctx.close();
  }

  // B. Preview first cert: Demo High Excel first -> SANTOS, MARIA, Excel Associate title,
  // ordinal given date, exam badge + partner badge present.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await gotoCompletion(page);
    await page.fill("#mos-comp-date", "2026-03-27");
    await page.dispatchEvent("#mos-comp-date", "change");
    await page.click("#mos-comp-preview-btn");
    await page.waitForTimeout(300);
    const pv = await page.evaluate(() => {
      const b = document.getElementById("mos-comp-preview");
      const q = (sel) => b ? b.querySelector(sel) : null;
      const bar = q(".mos-comp-bar");
      return { html: b ? b.innerHTML : "", vis: b && !b.classList.contains("d-none"),
        text: b ? b.innerText : "",
        bg: !!q(".mos-comp-bg"), bar: !!bar, stripe: !!q(".mos-comp-stripe"),
        wave: !!q(".mos-comp-wave"), rakso: !!q(".mos-comp-rakso"),
        barColor: bar ? getComputedStyle(bar).backgroundColor : "",
        examWeight: (q(".mos-comp-exam b") ? getComputedStyle(q(".mos-comp-exam b")).fontWeight : ""),
        waveLeft: (q(".mos-comp-wave") ? parseFloat(getComputedStyle(q(".mos-comp-wave")).left) : -1),
        partnerLeft: (q(".mos-comp-partner") ? parseFloat(getComputedStyle(q(".mos-comp-partner")).left) : -1),
        partnerWidth: (q(".mos-comp-partner") ? parseFloat(getComputedStyle(q(".mos-comp-partner")).width) : -1),
        waveTop: (q(".mos-comp-wave") ? parseFloat(getComputedStyle(q(".mos-comp-wave")).top) : -1) };
    });
    ok(pv.vis, "preview visible after Preview click");
    ok(/SANTOS, MARIA/.test(pv.html), "preview: export name formatted SANTOS, MARIA");
    ok(/Excel Associate\s*(<br>)?\s*\(Office 365 Apps\)/.test(pv.html), "preview: Excel Associate title");
    ok(/Given this 27th day of March, 2026\./.test(pv.text), "preview: ordinal given date");
    ok(/<u>27th<\/u> day of/.test(pv.html), "preview: given day underlined (template style)");
    ok(!/for the School Year \./.test(pv.html), "preview: no dangling empty School Year line");
    ok(/mos-comp-badge/.test(pv.html) && /mos-comp-partner/.test(pv.html), "preview: exam badge + partner badge present");
    ok(pv.bg && pv.bar && pv.stripe, "preview: template background (bar + stripe) present");
    ok(pv.barColor === "rgb(4, 134, 66)", "preview: bar is template green #048642 (got " + pv.barColor + ")");
    ok(pv.wave, "preview: green wave art present");
    ok(pv.rakso, "preview: RAKSO logo present");
    ok(pv.examWeight === "700", "preview: exam title is bold (got " + pv.examWeight + ")");
    ok(Math.abs(pv.waveLeft - 4.255 * 96) < 2 && Math.abs(pv.waveTop - 0.769 * 96) < 2, "preview: wave moved to fixed-template frame (got " + pv.waveLeft + "," + pv.waveTop + ")");
    ok(Math.abs(pv.partnerLeft - 9.629 * 96) < 2 && Math.abs(pv.partnerWidth - 1.224 * 96) < 2, "preview: partner badge positioned with gap + print margin (got " + pv.partnerLeft + "," + pv.partnerWidth + ")");
    ok(errs.length === 0, "section B: zero page errors");
    await ctx.close();
  }

  // C. Roster name wins + Expert has NO badge: filter to Word (Demo High) -> REYES, JOSE P.;
  // filter to Excel Expert -> Expert title, no badge image.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await page.selectOption("#mos-f-exam", "Word");
    await page.waitForTimeout(400);
    await gotoCompletion(page);
    await page.click("#mos-comp-preview-btn");
    await page.waitForTimeout(300);
    const pvW = await page.evaluate(() => (document.getElementById("mos-comp-preview") || {}).innerHTML || "");
    ok(/REYES, JOSE P\./.test(pvW), "roster name wins: REYES, JOSE P. (not export name)");
    await ctx.close();
  }
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await page.selectOption("#mos-f-exam", "Excel Expert");
    await page.waitForTimeout(400);
    await gotoCompletion(page);
    const t = await batchText(page);
    ok(/1 certificate/.test(t), "Expert filter: 1 certificate");
    await page.click("#mos-comp-preview-btn");
    await page.waitForTimeout(300);
    const pv = await page.evaluate(() => (document.getElementById("mos-comp-preview") || {}).innerHTML || "");
    ok(/Excel Expert\s*(<br>)?\s*\(Office 365 Apps\)/.test(pv), "Expert cert title (no Associate)");
    ok(!/mos-comp-badge/.test(pv), "Expert cert: NO badge image");
    ok(/mos-comp-partner/.test(pv), "Expert cert: partner badge still present");
    ok(errs.length === 0, "section C: zero page errors");
    await ctx.close();
  }

  // D. Filters narrow the batch: school = Second High -> 1 cert (F Word);
  // exam = PowerPoint -> 0 certs, buttons disabled.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await page.selectOption("#mos-f-school", "Second High");
    await page.waitForTimeout(400);
    await gotoCompletion(page);
    const t = await batchText(page);
    ok(/1 certificate/.test(t) && /Second High/.test(t), "school filter: Second High batch = 1");
    await ctx.close();
  }
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await page.selectOption("#mos-f-exam", "PowerPoint");
    await page.waitForTimeout(400);
    await gotoCompletion(page);
    const dis = await page.evaluate(() => document.getElementById("mos-comp-print").disabled && document.getElementById("mos-comp-pdf").disabled);
    ok(dis, "empty batch: Print + Save as PDF disabled");
    ok(errs.length === 0, "section D: zero page errors");
    await ctx.close();
  }

  // E. Signatories persist across reload (same context) + signature upload thumb.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    await gotoCompletion(page);
    await page.fill("#mos-comp-n1", "DUMMY ONE");
    await page.fill("#mos-comp-t1", "Dummy Title One");
    await page.setInputFiles("#mos-comp-s1", path.join(ROOT, "tests", "comp-sig.png"));
    await page.click("#mos-comp-sigsave");
    await page.waitForTimeout(300);
    const msg = await page.evaluate(() => (document.getElementById("mos-comp-msg") || {}).textContent || "");
    ok(/saved in this browser/i.test(msg), "signatories saved (msg: " + msg + ")");
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.waitForTimeout(800);
    await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx"));
    await page.waitForTimeout(2000);
    await page.click('[data-tab="dash-mos"]');
    await page.waitForTimeout(300);
    await page.evaluate(args => window.QBR.mos._loadAttempts(args.a, args.r), { a: attempts(), r: roster() });
    await page.waitForTimeout(500);
    await gotoCompletion(page);
    const n1 = await page.evaluate(() => (document.getElementById("mos-comp-n1") || {}).value || "");
    ok(n1 === "DUMMY ONE", "signatory name persists across reload (got " + n1 + ")");
    const thumb = await page.evaluate(() => !!document.querySelector('#mos-view-completion img[src^="data:image/"]'));
    ok(thumb, "uploaded signature thumbnail shown after reload");
    ok(errs.length === 0, "section E: zero page errors");
    await ctx.close();
  }

  // F. Issued log: separate completion block in the same sheet, passer cols untouched, dedup.
  {
    const { ctx, page, errs } = await bootMosRoster(browser);
    const aoa = await page.evaluate(async () => {
      const b0 = window.QBR.mosCerts._newLogBytes();
      const b1 = await window.QBR.mosCerts._applyOps(b0, [{ t: "add", code: "TEST-0001", row: { "Verification Code": "TEST-0001", "Name": "Dummy Passer", "Certification": "Excel Associate", "School": "Demo High" } }]);
      const rec = { t: "completion", name: "SANTOS, MARIA", exam: "Excel", school: "Demo High", section: "", given: "2026-03-27", at: new Date("2026-03-28") };
      const b2 = await window.QBR.mosCerts._applyOps(b1, [rec]);
      const b3 = await window.QBR.mosCerts._applyOps(b2, [Object.assign({}, rec, { at: new Date("2026-03-29") })]);
      const wb = XLSX.read(b3, { type: "array" });
      return XLSX.utils.sheet_to_json(wb.Sheets["Certificate Log"], { header: 1, defval: "", blankrows: true });
    });
    const hdr = aoa[0] || [];
    ok(hdr.indexOf("Verification Code") === 0, "log: passer header intact at col 0");
    ok(hdr.indexOf("Completion Name") === 18, "log: completion block starts col 18 (got " + hdr.indexOf("Completion Name") + ")");
    ok(hdr[17] === "", "log: separator col 17 blank");
    ok(hdr.indexOf("Completion Count") === 24, "log: completion block ends col 24");
    const compRows = aoa.filter(r => r[18] === "SANTOS, MARIA");
    ok(compRows.length === 1, "log: re-issue dedups to one row (got " + compRows.length + ")");
    ok(compRows[0] && compRows[0][24] === 2, "log: re-issue bumps count to 2 (got " + (compRows[0] || [])[24] + ")");
    ok(compRows[0] && compRows[0][19] === "Excel" && compRows[0][20] === "Demo High", "log: completion exam/school stored");
    const passer = aoa.filter(r => r[0] === "TEST-0001");
    ok(passer.length === 1 && passer[0][1] === "Dummy Passer" && passer[0][18] === "", "log: passer row untouched, no completion cols leaked");
    ok(errs.length === 0, "section F: zero page errors");
    await ctx.close();
  }

  await browser.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL", e); process.exit(2); });
