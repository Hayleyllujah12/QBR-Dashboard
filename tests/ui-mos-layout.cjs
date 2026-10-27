// MOS completion certificate layout (2026-10-27): Phase A shared centre line +
// Phase B "Adjust layout" (nudge / size / centre / reset / drag + snap /
// arrow keys / persistence), the same layout in the HTML print and the PPTX
// export. Dummy data only.
// Usage: PW=<dir with node_modules/playwright> node tests/ui-mos-layout.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), path = require("path");
global.JSZip = require(path.join(__dirname, "..", "libs", "jszip.min.js"));
const JSZip = global.JSZip;
const L = require(path.join(__dirname, "..", "js", "mos-completion-layout.js"));
const eng = require(path.join(__dirname, "..", "js", "mos-completion-pptx.js"));
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
const IN = 914400, CX = 5.971;
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("ok - " + m); } else { fail++; console.log("FAIL - " + m); } };
const near = (a, b, t) => Math.abs(a - b) <= t;

const tplB64 = (/b64: "([^"]+)"/.exec(fs.readFileSync(path.join(ROOT, "js", "mos-completion-template.js"), "utf8")) || [])[1];
const cert = (name, family) => ({ name, family,
  title: family === "Excel Expert" ? "Microsoft Office Specialist – Excel Expert (Office 365 Apps)"
    : "Microsoft Office Specialist – " + family + " Associate (Office 365 Apps)",
  sy: "2025–2026", given: "Given this 27th day of March, 2026." });
const SIG = { n1: "DUMMY ONE", t1: "Dummy Title One", n2: "DUMMY TWO", t2: "Dummy Title Two", org: "Dummy Org Inc.", s1: "", s2: "" };

// Shape frames of a slide by cNvPr name: {name: [x, y, cx, cy]} (first xfrm after the name).
function frames(xml) {
  const out = {}, re = /<p:cNvPr id="\d+" name="([^"]*)"[\s\S]*?<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g;
  let m; while ((m = re.exec(xml)) !== null) if (!out[m[1]]) out[m[1]] = m.slice(2).map(Number);
  return out;
}
async function slide1(lay, sig) {
  const bytes = await eng.generate(tplB64, [cert("ROE, JANE", "Excel")], Object.assign({}, sig || SIG, lay === undefined ? {} : { layout: lay }));
  const z = await JSZip.loadAsync(bytes);
  return z.file("ppt/slides/slide1.xml").async("string");
}

(async () => {
  // ---- A. Layout model (Node) ----
  const d = L.layoutOf(null);
  ok(L.ORDER.length === 8 && L.ORDER.every(id => d[id].dx === 0 && d[id].dy === 0 && d[id].size === 100), "A: defaults = 8 objects at 0/0/100%");
  ok(["title", "certify", "name", "exam"].every(id => near(L.baseLeft(id) + L.BOX[id].w / 2, CX, 1e-4)), "A: title/certify/name/exam centred on 5.971in");
  ok(["title", "certify", "name", "exam"].every(id => L.centreDx(id) === 0), "A: centred objects need no Centre nudge");
  const cl = L.layoutOf({ name: { dx: 50, dy: -9, size: 999 }, badge: { size: 999 }, title: { dx: "x", size: "y" } });
  const nb = L.boxOf("name", cl);
  ok(cl.name.dx === 5 && near(nb.t + nb.h / 2, 0, 0.006) && cl.name.size === 150, "A: nudge clamps to ±5in with the box centre kept on the page; font 150% (dx " + cl.name.dx + ", dy " + cl.name.dy + ")");
  ok(cl.badge.size === 200 && L.layoutOf({ badge: { size: 1 } }).badge.size === 50, "A: images clamp size 50–200%");
  ok(cl.title.dx === 0 && cl.title.size === 100, "A: garbage values fall back to defaults");
  ok(L.layoutOf({ name: { dx: 0.123456 } }).name.dx === 0.12, "A: nudges round to 0.01in");
  ok(L.outside("logos", L.layoutOf({ logos: { dy: -0.5 } })) && !L.outside("logos", d), "A: 'outside the page' flags a logo pushed off the top edge");
  ok(L.isDefault(d) && !L.isDefault({ sig1: { dy: 0.1 } }), "A: isDefault");
  ok(L.centreDx("badge") === Math.round((CX - (1.242 + 1.931 / 2)) * 100) / 100, "A: Centre puts the badge on the shared line");

  // ---- B. PPTX: Phase A + layout ----
  const fD = frames(await slide1(undefined));
  for (const [nm, id] of [["TextBox 12", "title"], ["TextBox 13", "certify"], ["TextBox 15", "name"], ["TextBox 14", "exam"]]) {
    const f = fD[nm];
    ok(near(f[0] + f[2] / 2, CX * IN, 1), "B: default " + id + " box centred on 5.971in (got " + ((f[0] + f[2] / 2) / IN).toFixed(4) + ")");
  }
  const xD = await slide1(undefined);
  const blockOf = (xml, nm) => { const i = xml.indexOf('name="' + nm + '"'), s = xml.lastIndexOf("<p:sp>", i); return xml.slice(s, xml.indexOf("</p:sp>", i)); };
  ok(!/algn="l"/.test(blockOf(xD, "TextBox 12")) && !/algn="l"/.test(blockOf(xD, "TextBox 13")), "B: title + certify paragraphs centred (algn=ctr)");
  ok(JSON.stringify(frames(await slide1(L.defaults()))) === JSON.stringify(fD), "B: explicit default layout = no layout");
  const LAY = { name: { dx: -0.2, dy: 0.1, size: 110 }, badge: { dx: 0.3, size: 120 }, logos: { dy: 0.2, size: 80 }, sig1: { dy: 0.15 }, title: { size: 90 } };
  const png = fs.readFileSync(path.join(__dirname, "comp-sig.png")).toString("base64");
  const sigP = Object.assign({}, SIG, { s1: "data:image/png;base64," + png });
  const xP0 = await slide1(undefined, sigP), xL = await slide1(LAY, sigP), fL = frames(xL), fP0 = frames(xP0);
  ok(fL["TextBox 15"][0] - fD["TextBox 15"][0] === -182880 && fL["TextBox 15"][1] - fD["TextBox 15"][1] === 91440, "B: name nudged -0.2in / +0.1in (exact EMU)");
  ok(/ sz="4399"/.test(blockOf(xL, "TextBox 15")), "B: name font 110% (39.99pt → 43.99pt)");
  ok(/ sz="4496"/.test(blockOf(xL, "TextBox 12")) && / sz="5301"/.test(blockOf(xL, "TextBox 12")), "B: title font 90% on both lines");
  const b0 = fD["Freeform 11"], b1 = fL["Freeform 11"];
  ok(near(b1[2], b0[2] * 1.2, 1) && near(b1[0] + b1[2] / 2, b0[0] + b0[2] / 2 + 0.3 * IN, 1) && near(b1[1] + b1[3] / 2, b0[1] + b0[3] / 2, 1), "B: badge 120% around its centre, +0.3in right");
  const g0 = fD["Group 17"], g1 = fL["Group 17"];
  ok(near(g1[2], g0[2] * 0.8, 1) && near(g1[1] + g1[3] / 2, g0[1] + g0[3] / 2 + 0.2 * IN, 1), "B: logo group 80% + 0.2in down (one group = both logos)");
  ok(/<a:chOff x="6878694" y="0"\/><a:chExt cx="3115357" cy="1027043"\/>/.test(xL), "B: group child frame untouched (children follow the group)");
  ok(fL["TextBox 16"][1] - fP0["TextBox 16"][1] === 137160 && fL["Signature 1"][1] - fP0["Signature 1"][1] === 137160, "B: signatory 1 + its signature picture move together (+0.15in)");
  ok(fL["TextBox 17"].join() === fP0["TextBox 17"].join(), "B: signatory 2 untouched");
  const zx = await JSZip.loadAsync(await eng.generate(tplB64, [cert("ROE, JANE", "Excel Expert")], Object.assign({}, SIG, { layout: LAY })));
  ok(!frames(await zx.file("ppt/slides/slide1.xml").async("string"))["Freeform 11"], "B: Expert slide (no badge) tolerates a badge layout");
  for (const fam of ["PowerPoint", "Word"]) {
    const z = await JSZip.loadAsync(await eng.generate(tplB64, [cert("ROE, JANE", fam)], Object.assign({}, SIG, { layout: LAY })));
    const f = frames(await z.file("ppt/slides/slide1.xml").async("string"));
    const nameBox = Object.keys(f).map(k => f[k]).find(v => v[2] === 7131547);
    ok(nameBox && near(nameBox[0] + nameBox[2] / 2, (CX - 0.2) * IN, 1), "B: " + fam + " slide (different shape ids) gets the same layout");
  }

  // ---- C. Browser: panel, preview, drag, keys, persistence, print + PPTX ----
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  const page = await ctx.newPage(); const errs = [];
  page.on("pageerror", e => errs.push(e.message)); page.on("dialog", dg => dg.dismiss());
  async function openCompletion() {
    await page.click('[data-tab="dash-mos"]'); await page.waitForTimeout(300);
    const mk = (ref, school, exam, score, date, month, proctor, name) => ({ ref, school, exam, score, date, month, proctor, name });
    await page.evaluate(a => window.QBR.mos._loadAttempts(a, []), [
      mk("A1", "Demo High", "Excel", 500, "2026-10-01", "2026-10", "Proctor A", "Jane Roe"),
      mk("A2", "Demo High", "Excel", 600, "2026-10-20", "2026-10", "Proctor A", "Jane Roe")]);
    await page.waitForTimeout(400);
    await page.click('[data-mos-view="completion"]'); await page.waitForTimeout(400);
  }
  await page.goto(APP, { waitUntil: "domcontentloaded" }); await page.waitForTimeout(800);
  await page.evaluate(() => localStorage.removeItem("qbr-mos-completion-layout-v1"));
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx")); await page.waitForTimeout(2000);
  await openCompletion();
  ok(await page.locator("#mos-comp-layout tr[data-lay-row]").count() === 8, "C: Adjust layout panel lists 8 objects");
  ok(await page.locator("#mos-comp-layout table[data-nosort]").count() === 1, "C: layout table opts out of column sorting");
  await page.click("#mos-comp-preview-btn"); await page.waitForTimeout(300);
  const centres = () => page.evaluate(() => {
    const c = document.querySelector("#mos-comp-preview .mos-comp-cert").getBoundingClientRect(), o = {};
    document.querySelectorAll("#mos-comp-preview [data-lay]").forEach(e => {
      const r = e.getBoundingClientRect(), id = e.getAttribute("data-lay") + (e.classList.contains("mos-comp-partner") ? "P" : "");
      o[id] = { cx: (r.left - c.left + r.width / 2) / 96, cy: (r.top - c.top + r.height / 2) / 96, w: r.width / 96, l: (r.left - c.left) / 96, t: (r.top - c.top) / 96 };
    });
    return o;
  });
  let c0 = await centres();
  ok(["title", "certify", "name", "exam"].every(id => near(c0[id].cx, CX, 0.005)), "C: preview title/certify/name/exam on one centre line (" +
    ["title", "certify", "name", "exam"].map(id => c0[id].cx.toFixed(3)).join(", ") + ")");
  const tt = await page.evaluate(() => { const c = document.querySelector("#mos-comp-preview .mos-comp-cert").getBoundingClientRect(), r = document.createRange();
    r.selectNodeContents(document.querySelector("#mos-comp-preview .mos-comp-title")); return (r.getBoundingClientRect().left - c.left) / 96; });
  ok(tt > 1.242 + 1.931, "C: centred title text clears the exam badge (title starts " + tt.toFixed(2) + "in, badge ends 3.17in)");

  // Nudge + size via the inputs → saved + preview moved.
  await page.fill("#mos-lay-name-dx", "-0.2"); await page.dispatchEvent("#mos-lay-name-dx", "change"); await page.waitForTimeout(150);
  await page.fill("#mos-lay-badge-size", "120"); await page.dispatchEvent("#mos-lay-badge-size", "change"); await page.waitForTimeout(150);
  const st1 = await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")));
  ok(st1 && st1.name.dx === -0.2 && st1.badge.size === 120, "C: input changes save at once (this browser)");
  let c1 = await centres();
  ok(near(c1.name.cx, CX - 0.2, 0.005), "C: preview name moved -0.2in");
  ok(near(c1.badge.w, c0.badge.w * 1.2, 0.01) && near(c1.badge.cx, c0.badge.cx, 0.005), "C: preview badge 120% around its centre");
  // Font size %.
  await page.fill("#mos-lay-exam-size", "90"); await page.dispatchEvent("#mos-lay-exam-size", "change"); await page.waitForTimeout(150);
  const fs90 = await page.evaluate(() => parseFloat(getComputedStyle(document.querySelector("#mos-comp-preview .mos-comp-exam")).fontSize));
  ok(near(fs90, 16 * 0.9 * 96 / 72, 0.2), "C: exam font 90% (" + fs90.toFixed(2) + "px)");
  // Clamp shown back in the input.
  await page.fill("#mos-lay-certify-dx", "70"); await page.dispatchEvent("#mos-lay-certify-dx", "change"); await page.waitForTimeout(150);
  ok(await page.inputValue("#mos-lay-certify-dx") === "5", "C: out-of-range nudge clamps (shown back in the input: " + await page.inputValue("#mos-lay-certify-dx") + ")");
  // Arrow keys in an input: 0.01 / Shift 0.1.
  await page.click("#mos-lay-certify-reset"); await page.waitForTimeout(100);
  await page.focus("#mos-lay-certify-dy"); await page.keyboard.press("ArrowUp"); await page.keyboard.press("Shift+ArrowUp"); await page.waitForTimeout(100);
  ok(await page.inputValue("#mos-lay-certify-dy") === "0.11", "C: arrow 0.01in + Shift-arrow 0.1in (got " + await page.inputValue("#mos-lay-certify-dy") + ")");
  // Centre + per-object reset + warnings.
  await page.click("#mos-lay-badge-centre"); await page.waitForTimeout(150);
  c1 = await centres();
  ok(near(c1.badge.cx, CX, 0.006), "C: Centre puts the badge on the centre line");
  await page.click("#mos-lay-badge-reset"); await page.waitForTimeout(150);
  c1 = await centres();
  ok(near(c1.badge.cx, c0.badge.cx, 0.005) && near(c1.badge.w, c0.badge.w, 0.005), "C: per-object Reset");
  await page.fill("#mos-lay-logos-dy", "-0.5"); await page.dispatchEvent("#mos-lay-logos-dy", "change"); await page.waitForTimeout(150);
  ok((await page.textContent("#mos-lay-logos-warn")) === "Outside the page", "C: 'Outside the page' warning");
  c1 = await centres();
  ok(near(c1.logos.cy - c0.logos.cy, -0.5, 0.006) && near(c1.logosP.cy - c0.logosP.cy, -0.5, 0.006), "C: RAKSO + partner logos move together");
  await page.fill("#mos-lay-logos-size", "80"); await page.dispatchEvent("#mos-lay-logos-size", "change"); await page.waitForTimeout(150);
  c1 = await centres();
  const gc0 = 7.523 + 3.407 / 2;
  ok(near(c1.logosP.cx - gc0, (c0.logosP.cx - gc0) * 0.8, 0.01) && near(c1.logos.cx - gc0, (c0.logos.cx - gc0) * 0.8, 0.01), "C: logos scale around the group centre (same as the PPTX group)");

  // HTML ↔ PPTX: the saved layout lands identically in both.
  const savedLay = await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")));
  const fx = frames(await slide1(savedLay));
  c1 = await centres();
  ok(near(c1.name.cx, (fx["TextBox 15"][0] + fx["TextBox 15"][2] / 2) / IN, 0.005), "C: name centre equal in HTML and PPTX");
  ok(near(c1.badge.l, fx["Freeform 11"][0] / IN, 0.005) && near(c1.badge.w, fx["Freeform 11"][2] / IN, 0.005), "C: badge frame equal in HTML and PPTX");
  const gx = fx["Group 17"];
  ok(near(c1.logosP.cx - (gx[0] + gx[2] / 2) / IN, (9.629 + 1.224 / 2 - gc0) * 0.8, 0.01), "C: partner logo position equal in HTML and PPTX group");

  // Reset all.
  await page.click("#mos-lay-resetall"); await page.waitForTimeout(150);
  c1 = await centres();
  ok(JSON.stringify(Object.keys(c0).map(k => c0[k].cx.toFixed(3))) === JSON.stringify(Object.keys(c1).map(k => c1[k].cx.toFixed(3))), "C: Reset all restores every object");
  ok((await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")))).name.dx === 0, "C: Reset all saved");

  // Drag with snap-to-centre, then arrow keys on the selection.
  await page.check("#mos-lay-drag"); await page.waitForTimeout(100);
  await page.locator("#mos-comp-preview [data-lay=name]").scrollIntoViewIfNeeded();
  const nm = await page.locator("#mos-comp-preview [data-lay=name]").boundingBox();
  await page.mouse.move(nm.x + nm.width / 2, nm.y + nm.height / 2); await page.mouse.down();
  await page.mouse.move(nm.x + nm.width / 2 + 48, nm.y + nm.height / 2 + 24, { steps: 4 });
  const guideOn1 = await page.evaluate(() => document.querySelector("#mos-comp-preview .mos-comp-guide").classList.contains("on"));
  await page.mouse.up(); await page.waitForTimeout(150);
  let s2 = await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")));
  ok(near(s2.name.dx, 0.5, 0.02) && near(s2.name.dy, 0.25, 0.02) && !guideOn1, "C: drag moves the name (dx " + s2.name.dx + ", dy " + s2.name.dy + ") and saves");
  ok(await page.inputValue("#mos-lay-name-dx") === String(s2.name.dx), "C: drag updates the panel inputs");
  await page.locator("#mos-comp-preview [data-lay=name]").scrollIntoViewIfNeeded();
  const nm2 = await page.locator("#mos-comp-preview [data-lay=name]").boundingBox();
  await page.mouse.move(nm2.x + nm2.width / 2, nm2.y + nm2.height / 2); await page.mouse.down();
  await page.mouse.move(nm2.x + nm2.width / 2 - 46, nm2.y + nm2.height / 2, { steps: 4 });
  const guideOn2 = await page.evaluate(() => document.querySelector("#mos-comp-preview .mos-comp-guide").classList.contains("on"));
  await page.mouse.up(); await page.waitForTimeout(150);
  s2 = await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")));
  ok(s2.name.dx === 0 && guideOn2, "C: dragging near the centre line snaps onto it (guide shown)");
  await page.keyboard.press("ArrowRight"); await page.keyboard.press("Shift+ArrowDown"); await page.waitForTimeout(100);
  s2 = await page.evaluate(() => JSON.parse(localStorage.getItem("qbr-mos-completion-layout-v1")));
  ok(s2.name.dx === 0.01 && near(s2.name.dy, 0.35, 1e-9), "C: arrow keys nudge the selected object (0.01 / Shift 0.1)");
  ok(await page.evaluate(() => !!document.querySelector("#mos-comp-preview [data-lay=name].mos-lay-sel")), "C: selected object outlined");

  // Persistence across reload.
  await page.reload({ waitUntil: "domcontentloaded" }); await page.waitForTimeout(800);
  await page.setInputFiles("#file-input", path.join(ROOT, "tests", "fixture.xlsx")); await page.waitForTimeout(2000);
  await openCompletion();
  ok(await page.inputValue("#mos-lay-name-dx") === "0.01" && await page.inputValue("#mos-lay-name-dy") === "0.35", "C: layout survives a reload");

  // Print host carries the layout.
  await page.evaluate(() => { window.print = function () { window.__printed = document.getElementById("mos-completion-print").innerHTML; }; });
  await page.click("#mos-comp-print"); await page.waitForTimeout(1500);
  const printed = await page.evaluate(() => window.__printed || "");
  ok(/data-lay="name" style="transform:translate\(0\.01in,0\.35in\)/.test(printed), "C: printed certificate uses the saved layout");
  // PPTX download carries the layout.
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 20000 }), page.click("#mos-comp-pptx")]);
  const zd = await JSZip.loadAsync(fs.readFileSync(await dl.path()));
  const fd = frames(await zd.file("ppt/slides/slide1.xml").async("string"));
  ok(near(fd["TextBox 15"][0] + fd["TextBox 15"][2] / 2, (CX + 0.01) * IN, 1) && fd["TextBox 15"][1] === 3536271 + 0.35 * IN, "C: downloaded PPTX uses the saved layout");
  ok(errs.length === 0, "C: no page errors" + (errs.length ? " (" + errs.join(" | ") + ")" : ""));
  await b.close();
  console.log(pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log("FATAL - " + (e && e.stack || e)); process.exit(1); });
