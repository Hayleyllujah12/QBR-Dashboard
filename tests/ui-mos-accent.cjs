// MOS completion certificate accent colour follows the exam's badge (2026-10-27 revision list #5).
// HTML print preview: bar + stripe + wave use Excel green / PowerPoint orange-red / Word blue,
// the same colours as the PPTX template; PPTX export: PowerPoint/Word wave art recoloured too.
// Dummy data only. Usage: PW=<dir with node_modules/playwright> node tests/ui-mos-accent.cjs
"use strict";
const { chromium } = require(process.env.PW + "/playwright");
const fs = require("fs"), path = require("path");
global.JSZip = require(path.join(__dirname, "..", "libs", "jszip.min.js"));
const eng = require(path.join(__dirname, "..", "js", "mos-completion-pptx.js"));
const ROOT = path.resolve(__dirname, "..");
const APP = "file://" + path.join(ROOT, "index.html").replace(/\\/g, "/");
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; console.log("ok - " + m); } else { fail++; console.log("FAIL - " + m); } };
const COL = { Excel: "#048642", "Excel Expert": "#048642", PowerPoint: "#B8482B", Word: "#25589B" };

(async () => {
  // ---- HTML preview (browser) ----
  const b = await chromium.launch();
  const p = await b.newPage(); const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  await p.goto(APP); await p.waitForTimeout(800);
  const r = await p.evaluate(COL => {
    const M = QBR.mosCompletion, out = {};
    Object.keys(COL).forEach(fam => {
      const c = { name: "ROE, JANE", family: fam, title: "Microsoft Office Specialist – " + fam + " Associate (Office 365 Apps)",
        sy: "2025–2026", given: "Given this 27th day of March, 2026.", accent: M._accentOf(fam), wave: M._waveArt(fam) };
      const host = document.createElement("div"); host.innerHTML = M._certHtml(c, {}); document.body.appendChild(host);
      const sec = host.querySelector(".mos-comp-cert"), bar = host.querySelector(".mos-comp-bar"), st = host.querySelector(".mos-comp-stripe");
      const wave = c.wave && /base64,/.test(c.wave) ? atob(c.wave.split("base64,")[1]) : "";
      out[fam] = { accent: M._accentOf(fam), bar: getComputedStyle(bar).backgroundColor, stripe: getComputedStyle(st).backgroundColor,
        fam: sec.dataset.family, waveHas: wave ? (wave.match(/#[0-9A-Fa-f]{6}/g) || []).map(x => x.toUpperCase()) : null };
      host.remove();
    });
    return out;
  }, COL);
  const rgb = h => "rgb(" + [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16)).join(", ") + ")";
  Object.keys(COL).forEach(fam => {
    const x = r[fam];
    ok(x.accent === COL[fam] && x.bar === rgb(COL[fam]) && x.stripe === rgb(COL[fam]), `${fam}: left bar + stripe are ${COL[fam]} (${x.bar})`);
    if (x.waveHas) ok(x.waveHas.includes(COL[fam]) && (COL[fam] === "#048642" || !x.waveHas.includes("#048642")), `${fam}: wave art recoloured to ${COL[fam]}`);
  });
  ok(!errs.length, "no page errors " + errs.join(" | "));
  await b.close();

  // ---- PPTX export (Node engine) ----
  const tplSrc = fs.readFileSync(path.join(__dirname, "..", "js", "mos-completion-template.js"), "utf8");
  const tplB64 = (/b64: "([^"]+)"/.exec(tplSrc) || [])[1];
  const SIG = { n1: "DUMMY ONE", t1: "Dummy Title One", n2: "DUMMY TWO", t2: "Dummy Title Two", org: "Dummy Org Inc.", s1: "", s2: "" };
  const cert = fam => ({ name: "ROE, JANE", family: fam, title: "Microsoft Office Specialist – " + fam + " Associate (Office 365 Apps)", sy: "2025–2026", given: "Given this 27th day of March, 2026." });
  const out = await eng.generate(tplB64, [cert("Excel"), cert("PowerPoint"), cert("Word")], SIG);
  const zip = await JSZip.loadAsync(out);
  for (const [i, fam] of [[1, "Excel"], [2, "PowerPoint"], [3, "Word"]]) {
    const xml = await zip.file(`ppt/slides/slide${i}.xml`).async("string");
    const rels = await zip.file(`ppt/slides/_rels/slide${i}.xml.rels`).async("string");
    const svgName = (rels.match(/Target="\.\.\/media\/([^"]+\.svg)"/) || [])[1];
    const svg = svgName ? await zip.file("ppt/media/" + svgName).async("string") : "";
    const want = COL[fam].slice(1);
    ok(new RegExp('srgbClr val="' + want + '"').test(xml), `PPTX slide ${i} (${fam}): bar/stripe in ${COL[fam]}`);
    ok(svg.toUpperCase().includes("#" + want) && (fam === "Excel" || !/#048642/i.test(svg)), `PPTX slide ${i} (${fam}): wave art in ${COL[fam]}`);
  }
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
