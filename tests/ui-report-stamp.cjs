/* Admin › Reports "report generated" stamp toggle + loading progress bars (v1.42.0). Synthetic data only.
 * Usage: PW=<dir with node_modules/playwright> node tests/ui-report-stamp.cjs
 * A Admin › Reports: toggle default ON, OFF persists (qbr-report-stamp), QBR.reportStampOn()
 * B deck cover / image gallery meta (exportMeta): "Generated" row only when ON
 * C Trainers report print header: "printed <date>" only when ON, attribution kept
 * D MOS Exam Summary footer: "Generated <date>" only when ON, "RCT OpsDesk" kept
 * E progress bars: MOS PDF loading (determinate, reaches 100%), Trainers eval upload, Risky sign-ins pill CSS */
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), os = require("os"), path = require("path");
const ROOT = path.resolve(__dirname, "..");
const XLSX = require(path.join(ROOT, "libs/xlsx.full.min.js"));
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
const FX = path.join(__dirname, "mos-fixture");
let pass = 0, fail = 0;
const ok = (c, n) => { if (c) { pass++; console.log("ok - " + n); } else { fail++; console.log("FAIL - " + n); } };
const DIR = fs.mkdtempSync(path.join(os.tmpdir(), "rstamp-"));
const pdfs = dir => { let o = []; for (const f of fs.readdirSync(dir)) { const p = path.join(dir, f); if (fs.statSync(p).isDirectory()) { if (!/^(Archive Files|Certificates)$/.test(f)) o = o.concat(pdfs(p)); } else if (/\.pdf$/i.test(f)) o.push(p); } return o; };

function tracker() {
  const rows = [["Date", "Trainer", "New Code", "New Description", "Client", "Time Start", "Time End", "Net Promoter Score", "Trainer Score"],
    [46100, "DOE, John", "MS-101A", "Microsoft 365 Education Onboarding", "Sample Academy", 0.375, 0.5, 70, 4.5],
    [46101, "ROE, Jane", "MS-102A", "Teams for Classrooms", "Bright Path Academy", 0.54, 0.62, 60, 4.2]];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "trial");
  const p = path.join(DIR, "Training Tracker.xlsx"); fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" })); return p;
}
function evalFile(n, label) {
  const head = ["ID", "Start time", "School / Organization", "Knowledge and mastery of the topic", "Communication and presentation skills",
    "Ability to engage and connect with participants", "Overall, how satisfied were you with the training?",
    "On a scale of 0 to 10, how likely are you to recommend this training session to a colleague", "Why did you give this rating?"];
  const rows = [head]; for (let i = 0; i < n; i++) rows.push([i + 1, 46100.4 + i / 1440, "SA", "Strongly Agree", "Agree", "Strongly Agree", "Very satisfied", 9, "ok"]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const p = path.join(DIR, "Training Feedback Form - " + label + ".xlsx"); fs.writeFileSync(p, XLSX.write(wb, { bookType: "xlsx", type: "buffer" })); return p;
}
async function boot(ctx) {
  const p = await ctx.newPage(); p._errs = []; p.on("pageerror", e => p._errs.push(String(e.message).slice(0, 200)));
  await p.goto(APP); await p.waitForTimeout(400);
  await p.evaluate(() => { try { localStorage.setItem("qbr-remember", "0"); } catch (e) {} });
  await p.setInputFiles("#file-input", path.join(__dirname, "fixture.xlsx")); await p.waitForTimeout(1800);
  return p;
}
const meta = p => p.evaluate(() => exportMeta().map(r => r[0]));

(async () => {
  const b = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
  const ctx = await b.newContext({ viewport: { width: 1440, height: 950 } });
  let p = await boot(ctx);

  /* ---- A + B ---- */
  ok(await p.evaluate(() => QBR.reportStampOn() === true && localStorage.getItem("qbr-report-stamp") === null), "default: stamp ON (no setting stored)");
  ok((await meta(p)).includes("Generated"), "deck cover / gallery meta has a Generated row by default");
  await p.dblclick("#app-version"); await p.waitForTimeout(200);
  await p.fill("#admin-pw-1", "s3cret"); await p.fill("#admin-pw-2", "s3cret"); await p.click("#admin-pw-ok"); await p.waitForTimeout(400);
  const btn = '#admin-body [data-report-stamp]';
  ok(await p.$(btn) && /Reports/.test(await p.textContent("#admin-body .admin-reports h4")) && (await p.textContent(btn)).trim() === "ON", "Admin › Reports section with the toggle ON");
  await p.click(btn); await p.waitForTimeout(200);
  ok((await p.textContent(btn)).trim() === "OFF" && await p.$eval(btn, el => el.getAttribute("aria-pressed") === "false") &&
     await p.evaluate(() => localStorage.getItem("qbr-report-stamp") === "0" && QBR.reportStampOn() === false), "toggle → OFF, stored, QBR.reportStampOn() false");
  const m = await meta(p);
  ok(!m.includes("Generated") && m.includes("Schools in scope"), "OFF: no Generated row; other cover rows kept");
  await p.evaluate(() => buildPrintCover());
  ok(!/Generated/.test(await p.textContent("#cover-meta")), "OFF: printed deck cover has no Generated row");
  await p.reload(); await p.waitForTimeout(400); await p.setInputFiles("#file-input", path.join(__dirname, "fixture.xlsx")); await p.waitForTimeout(1800);
  ok(await p.evaluate(() => QBR.reportStampOn() === false), "OFF persists after reload");

  /* ---- C: Trainers print header ---- */
  const trHead = async () => {
    await p.evaluate(() => { window.__ph = null; window.print = () => { const e = document.querySelector(".tr-ph-by"); window.__ph = e ? e.textContent : ""; }; });
    await p.click('[data-traction="print-report"]');
    await p.waitForFunction(() => window.__ph !== null, null, { timeout: 10000 }).catch(() => {});
    const t = await p.evaluate(() => window.__ph);
    await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(300);
    return t;
  };
  await p.evaluate(() => { window.showOpenFilePicker = undefined; });
  await p.click('[data-tab="dash-trainers"]'); await p.waitForTimeout(300);
  {
    const [c] = await Promise.all([p.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null), p.click('[data-trpill="tracker"]')]);
    if (c) await c.setFiles(tracker()); await p.waitForTimeout(700);
  }
  await p.click('#dash-trainers button:text-is("Summary")'); await p.waitForTimeout(700);
  let ph = await trHead();
  ok(ph === "RCT OpsDesk · Trainer evaluations", "OFF: Trainers report header keeps the attribution, no printed date  [" + ph + "]");
  await p.evaluate(() => localStorage.setItem("qbr-report-stamp", "1"));
  ph = await trHead();
  ok(/^RCT OpsDesk · Trainer evaluations · printed .*\d/.test(ph || ""), "ON: printed date back  [" + ph + "]");

  /* ---- E1: Trainers eval upload progress ---- */
  await p.evaluate(() => { window.__trp = []; const el = document.getElementById("tr-status");
    new MutationObserver(() => { const pb = el.querySelector(".progress"); if (pb) window.__trp.push(+pb.getAttribute("aria-valuenow")); }).observe(el, { childList: true, subtree: true }); });
  {
    const [c] = await Promise.all([p.waitForEvent("filechooser", { timeout: 5000 }).catch(() => null), p.click('[data-trpill="eval"]')]);
    if (c) await c.setFiles([evalFile(6, "Mar. 21, 2026 [SA] - AM"), evalFile(8, "Mar. 22, 2026 [BPA] - PM")]); await p.waitForTimeout(1200);
  }
  const trp = await p.evaluate(() => window.__trp);
  ok(trp.length >= 2 && trp[0] === 0 && trp.includes(50), "Trainers eval upload shows a determinate bar per file  " + JSON.stringify(trp));
  ok(!(await p.$("#tr-status .progress")) && /eval file/.test(await p.textContent("#tr-status")), "bar replaced by the summary note when done");

  /* ---- D + E2: MOS ---- */
  await p.click('[data-tab="dash-mos"]'); await p.waitForTimeout(300);
  await p.evaluate(() => { window.__mp = []; const el = document.getElementById("mos-status");
    new MutationObserver(() => { const pb = el.querySelector(".progress"); if (pb) window.__mp.push(+pb.getAttribute("aria-valuenow")); }).observe(el, { childList: true, subtree: true }); });
  await p.fill("#mos-school-in", "Demo High School");
  await p.setInputFiles("#mos-pdf", pdfs(path.join(FX, "01 - Demo High School")));
  await p.waitForFunction(() => QBR.mos._state().model && !QBR.mos._state().busy, null, { timeout: 30000 });
  const mp = await p.evaluate(() => window.__mp);
  ok(mp.length >= 2 && mp[0] === 0 && mp.every((v, i) => !i || v >= mp[i - 1]) && Math.max(...mp) > 0, "MOS PDF loading: determinate bar, never goes backwards  " + JSON.stringify(mp.slice(0, 12)));
  ok(!(await p.$("#mos-status .progress")) && /PDF/.test(await p.textContent("#mos-status")), "MOS bar replaced by the file summary when done");
  const foot = async () => {
    await p.evaluate(() => { window.__ft = null; window.print = () => { const f = document.querySelector("#mos-print .mos-sum-foot"); window.__ft = f ? f.textContent : ""; }; });
    await p.click("#mos-summary");
    await p.waitForFunction(() => window.__ft !== null, null, { timeout: 10000 }).catch(() => {});
    const t = await p.evaluate(() => window.__ft);
    await p.evaluate(() => window.dispatchEvent(new Event("afterprint"))); await p.waitForTimeout(300);
    return t || "";
  };
  let ft = await foot();
  ok(/Generated \d.* · RCT OpsDesk/.test(ft), "ON: MOS summary footer has Generated <date> · RCT OpsDesk");
  await p.evaluate(() => localStorage.setItem("qbr-report-stamp", "0"));
  ft = await foot();
  ok(!/Generated/.test(ft) && /RCT OpsDesk$/.test(ft.trim()), "OFF: MOS summary footer keeps RCT OpsDesk, no date");

  /* ---- E3: Risky sign-ins reading pill ---- */
  const css = fs.readFileSync(path.join(ROOT, "css/styles.css"), "utf8"), soc = fs.readFileSync(path.join(ROOT, "js/soc.js"), "utf8");
  ok(/#dash-soc \.readbar i\{[^}]*animation: readslide/.test(css) && !/#dash-soc \.spin\b/.test(css) && /class="readbar" role="progressbar"/.test(soc), "Risky sign-ins file pill uses the sliding bar (spinner CSS removed)");

  ok(!p._errs.length, "no page errors " + p._errs.join(" | "));
  await b.close();
  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
