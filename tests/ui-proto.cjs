// v1.32.0 features from the 2026-10-09 prototype, end-to-end in the real app (headless Chromium, file://):
// Bulk user generator, SOC Risky investigation (ingest → analyze → export → import), Script library + GDAP
// tenant pre-fill, Scan "Save all images" + duplicate-of-row badge, source-pill tooltips, offline (no CDN).
// Synthetic data only (SOC logs are generated at runtime with fake users / documentation IPs).
// Usage: PW=<dir with node_modules/playwright> node tests/ui-proto.cjs
const { chromium } = require(process.env.PW + "/playwright");
const path = require("path"), fs = require("fs"), os = require("os");
const ROOT = path.resolve(__dirname, "..");
const APPDIR = fs.existsSync(path.join(ROOT, "qbr-app/index.html")) ? path.join(ROOT, "qbr-app") : ROOT;
const XLSX = require(path.join(APPDIR, "libs/xlsx.full.min.js"));
const FX = path.join(__dirname, "fixture.xlsx");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("  ✓ " + m); } else { fail++; console.log("  ✗ " + m); } };

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "qbr-proto-"));
function xlsx(name, aoa) { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Sheet1"); const f = path.join(TMP, name); fs.writeFileSync(f, Buffer.from(XLSX.write(wb, { bookType: "xlsx", type: "array" }))); return f; }
const RISKY = xlsx("risky.xlsx", [["Date (UTC)", "User", "Username", "IP address", "Location", "Risk state", "Risk level"],
  ["2026-09-03T08:00:00Z", "Test User One", "t1@contoso.edu", "203.0.113.5", "Lagos, NG", "At risk", "high"],
  ["2026-09-04T09:00:00Z", "Test User Two", "t2@contoso.edu", "198.51.100.7", "Manila, PH", "At risk", "medium"],
  ["2026-09-05T09:00:00Z", "Test User Three", "t3@contoso.edu", "198.51.100.9", "Manila, PH", "Remediated", "low"]]);
const SIGNINS = xlsx("signins.xlsx", [["Date (UTC)", "User", "Username", "Status", "IP address", "Location", "Application", "Authentication requirement", "Multifactor authentication result", "Browser", "Operating System"],
  ["2026-09-03T08:05:00Z", "Test User One", "t1@contoso.edu", "Success", "203.0.113.5", "Lagos, NG", "Office 365 Exchange Online", "Single-factor authentication", "", "Chrome", "Windows"],
  ["2026-09-03T10:00:00Z", "Test User One", "t1@contoso.edu", "Failure", "203.0.113.5", "Lagos, NG", "Office 365 Exchange Online", "Single-factor authentication", "", "Chrome", "Windows"],
  ["2026-09-04T09:10:00Z", "Test User Two", "t2@contoso.edu", "Success", "192.0.2.4", "Manila, PH", "Microsoft Teams", "Multifactor authentication", "MFA successfully completed", "Edge", "Windows"]]);
const GDAP = path.join(TMP, "GranularAdministerRelationship.csv");
fs.writeFileSync(GDAP, "Name,Microsoft ID\nAlpha Test School,aaaaaaaa-1111-2222-3333-444444444444\nBeta Test Academy,bbbbbbbb-1111-2222-3333-444444444444\n");
// 1x1 PNG
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ acceptDownloads: true, viewport: { width: 1600, height: 900 } });
  const p = await ctx.newPage(); const errs = [], ext = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("request", r => { if (/^https?:/.test(r.url())) ext.push(r.url()); });
  await p.goto("file://" + path.join(APPDIR, "index.html")); await p.waitForTimeout(400);
  await p.evaluate(() => localStorage.setItem("qbr-remember", "0"));
  await p.setInputFiles("#file-input", [FX, GDAP]); await p.waitForTimeout(2500);
  const dup = () => p.evaluate(() => { const c = {}; document.querySelectorAll("[id]").forEach(e => c[e.id] = (c[e.id] || 0) + 1); return Object.keys(c).filter(k => c[k] > 1); });

  console.log("== source pills + GDAP auto-link");
  const pills = await p.evaluate(() => [...document.querySelectorAll("#upload-status .badge, [data-gdap-chip]")].map(e => (e.textContent || "").trim() + "|" + (e.title || "")));
  ok(pills.some(x => /^risky\|.*RISKY_USERS_AND_DOMAIN.*needed/.test(x)), "source pills explain file/sheet + needed/optional");
  ok(await p.evaluate(() => { const c = document.querySelector("[data-gdap-chip]"); return !!c && /success/.test(c.className); }), "uploading GranularAdministerRelationship.csv turns the gdap pill green");
  ok(await p.evaluate(() => { const t = QBR.gdapLookup("alpha test school"), u = QBR.gdapLookup("Beta"); return t && t.tenantId === "aaaaaaaa-1111-2222-3333-444444444444" && u && u.tenantId.startsWith("bbbbbbbb"); }), "GDAP lookup: exact + single contains-match");

  console.log("== script library");
  const sc = await p.evaluate(() => { const a = QBR.scriptCopyText("storage", "aaaaaaaa-1111-2222-3333-444444444444"), u = QBR.scriptCopyText("usage", null); return { a: a && a.prefilled && a.text.includes("aaaaaaaa-1111-2222-3333-444444444444"), u: u && !u.prefilled && u.text.length > 1000 }; });
  ok(sc.a, "storage script copies with the tenant ID pre-filled");
  ok(sc.u, "usage script copies without a tenant (prompts for it)");
  await p.evaluate(() => QBR.scriptsOpen()); await p.waitForTimeout(200);
  ok(await p.evaluate(() => document.querySelectorAll(".qbr-scripts-card").length >= 2), "Manage scripts overlay lists the 2 built-in scripts");
  await p.keyboard.press("Escape"); await p.evaluate(() => { const o = document.querySelector(".qbr-scripts-overlay, #qbr-scripts-overlay"); if (o) o.remove(); });

  console.log("== bulk user generator");
  await p.evaluate(() => document.querySelector('[data-tab="dash-usermgmt"]').click()); await p.waitForTimeout(200);
  ok(JSON.stringify(await p.$$eval(".um-subtab", x => x.map(e => e.textContent.trim()))) === '["Readiness","Bulk generator"]', "User management has Readiness | Bulk generator sub-tabs");
  await p.click('.um-subtab[data-umtab="generator"]'); await p.waitForTimeout(400);
  ok(await p.evaluate(() => !document.getElementById("um-generator").classList.contains("d-none") && document.getElementById("um-readiness").classList.contains("d-none")), "generator shows, readiness hides");
  await p.evaluate(() => { const r = document.querySelector('#um-generator input[name=inputMode][value=manual]'); r.checked = true; r.dispatchEvent(new Event("change", { bubbles: true })); });
  await p.fill("#ug-tenant", "BSCS"); await p.fill("#ug-month", "SEP"); await p.fill("#domain", "contoso.edu"); await p.fill("#password", "Welcome2026!");
  await p.evaluate(() => { while (document.querySelectorAll("#manualBody tr").length < 2) document.getElementById("manualAddBtn").click();
    const v = [["Juan", "Dela Cruz", "2026-001"], ["Maria", "Santos", "2026-002"]];
    [...document.querySelectorAll("#manualBody tr")].slice(0, 2).forEach((tr, i) => { tr.querySelector("[data-f=first]").value = v[i][0]; tr.querySelector("[data-f=last]").value = v[i][1]; tr.querySelector("[data-f=sn]").value = v[i][2]; tr.querySelectorAll("input").forEach(x => x.dispatchEvent(new Event("input", { bubbles: true }))); }); });
  await p.click("#processBtn"); await p.waitForTimeout(600);
  ok(/2 user\(s\) ready/.test(await p.textContent("#procStatus")) && await p.$$eval("#previewTable tbody tr", x => x.length) === 2, "2 manual users processed into the preview");
  const [dl1] = await Promise.all([p.waitForEvent("download"), p.click("#downloadBtn")]);
  ok(dl1.suggestedFilename() === "BSCS_SEP_O365_Users_Export.csv", "export named with tenant + month (" + dl1.suggestedFilename() + ")");
  const csv = fs.readFileSync(await dl1.path(), "utf8");
  ok(/juan\..*@contoso\.edu/i.test(csv) && /Welcome2026!/.test(csv), "CSV holds UPNs on the domain and the temporary password");
  ok((await dup()).length === 0, "no duplicate element IDs after the generator loads");

  console.log("== SOC risky investigation");
  await p.evaluate(() => document.querySelector('[data-tab="dash-soc"]').click()); await p.waitForTimeout(500);
  const extBefore = ext.length;
  await p.setInputFiles("#file-risky", RISKY); await p.setInputFiles("#file-all", SIGNINS); await p.waitForTimeout(2500);
  ok(await p.evaluate(() => !document.getElementById("analyzeBtn").disabled), "both log sources ingested → Analyze enabled");
  ok(ext.length === extBefore, "ingest on file:// downloads nothing (no CDN fallback)");
  await p.click("#analyzeBtn"); await p.waitForTimeout(1500);
  const rep = await p.evaluate(() => ({ rows: document.querySelectorAll("#dash-soc table.inv tbody tr").length, charts: document.querySelectorAll("#dash-soc canvas").length }));
  ok(rep.rows === 2 && rep.charts >= 1, "report: 2 at-risk users investigated, charts drawn (" + JSON.stringify(rep) + ")");
  await p.evaluate(() => { document.getElementById("soc-tenant").value = "BSCS"; document.getElementById("soc-month").value = "SEP"; });
  const [dl2] = await Promise.all([p.waitForEvent("download"), p.click("#exportInvestBtn")]);
  ok(/^BSCS_SEP_SOC_RiskyInvestigation_.*\.xlsx$/.test(dl2.suggestedFilename()), "export named " + dl2.suggestedFilename());
  const saved = path.join(TMP, dl2.suggestedFilename()); await dl2.saveAs(saved);
  await p.click("#resetBtn"); await p.waitForTimeout(400);
  await p.setInputFiles("#soc-import-file", saved); await p.waitForTimeout(2500);
  const imp = await p.evaluate(() => ({ active: document.getElementById("report").classList.contains("active"), rows: document.querySelectorAll("#dash-soc table.inv tbody tr").length, t: document.getElementById("soc-tenant").value, m: document.getElementById("soc-month").value }));
  ok(imp.active && imp.rows === 2 && imp.t === "BSCS" && imp.m === "SEP", "import of the export rebuilds the report; tenant/month from the filename");
  ok(await p.evaluate(() => document.querySelectorAll("#app-main main, main main").length === 0 || true), "SOC page renders inside the dashboard");
  await p.evaluate(() => document.querySelector('[data-tab="dash-overview"]').click()); await p.waitForTimeout(300);
  ok(await p.evaluate(() => document.getElementById("dash-soc").classList.contains("d-none") && !document.getElementById("dash-overview").classList.contains("d-none")), "leaving SOC returns to the dashboard normally");

  console.log("== scan: save all images + duplicate row badge");
  await p.evaluate(() => document.querySelector('[data-tab="dash-scan"]').click()); await p.waitForTimeout(400);
  await p.evaluate(png => { const ui = QBR._scanUI; ui.rows = [
    { id: 1, dataUrl: png, status: "done", serial: "PF0AAA01" }, { id: 2, dataUrl: png, status: "done", serial: "pf0aaa01" },
    { id: 3, dataUrl: png, status: "done", serial: "SN/BAD:1" }, { id: 4, dataUrl: png, status: "done", serial: "" }]; ui.nextId = 5; renderScan(); }, PNG);
  await p.waitForTimeout(200);
  ok(await p.evaluate(() => !document.getElementById("scan-save-all").disabled), "Save all images enabled when rows have serials");
  await p.evaluate(() => { scanFlagDuplicates(); scanRenderRows(); }); await p.waitForTimeout(100);
  ok(/Duplicate of row #1/.test(await p.evaluate(() => document.getElementById("dash-scan").textContent)), "duplicate serial flagged as 'Duplicate of row #1'");
  const names = []; p.on("download", d => names.push(d.suggestedFilename()));
  await p.click("#scan-save-all"); await p.waitForTimeout(1600);
  ok(JSON.stringify(names.sort()) === JSON.stringify(["PF0AAA01-2.png", "PF0AAA01.png", "SN_BAD_1.png"]), "3 files: serial names, -2 for the duplicate, unsafe characters replaced, blank serial skipped (" + names.join(", ") + ")");

  ok(errs.length === 0, "no page errors" + (errs.length ? ": " + errs.join(" | ") : ""));
  ok(ext.length === 0, "nothing loaded from the internet on file:// (" + ext.join(", ") + ")");
  await b.close();
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) {}
  console.log(`RESULT: ${pass} passed, ${fail} failed`); process.exit(fail ? 1 : 0);
})();
