// MOS completion certificates — PPTX export engine (Node).
// Verifies template fill, slide cloning, Expert badge removal, signature
// pictures, slide assembly (sldIdLst / rels / content types), geometry
// preservation, and the scrubbed-template placeholders. Dummy data only.
// Usage: node tests/mos-completion-pptx.cjs
"use strict";
const fs = require("fs");
const path = require("path");
global.JSZip = require(path.join(__dirname, "..", "libs", "jszip.min.js"));
const eng = require(path.join(__dirname, "..", "js", "mos-completion-pptx.js"));
const JSZip = global.JSZip;

let pass = 0, fail = 0;
function ok(c, msg) { if (c) { pass++; console.log("ok - " + msg); } else { fail++; console.log("FAIL - " + msg); } }
function eq(a, b, msg) { ok(a === b, msg + " (got " + JSON.stringify(a) + ", want " + JSON.stringify(b) + ")"); }

const tplSrc = fs.readFileSync(path.join(__dirname, "..", "js", "mos-completion-template.js"), "utf8");
const tplB64 = (/b64: "([^"]+)"/.exec(tplSrc) || [])[1];
ok(!!tplB64 && tplB64.length > 100000, "template module carries base64 template");

const GIVEN = "Given this 27th day of March, 2026.";
const SIG = { n1: "DUMMY ONE", t1: "Dummy Title One", n2: "DUMMY TWO", t2: "Dummy Title Two", org: "Dummy Org Inc.", s1: "", s2: "" };
const cert = (name, family, sy) => ({
  name, family,
  title: family === "Excel Expert"
    ? "Microsoft Office Specialist \u2013 Excel Expert (Office 365 Apps)"
    : "Microsoft Office Specialist \u2013 " + family + " Associate (Office 365 Apps)",
  sy: sy == null ? "2025\u20132026" : sy, given: GIVEN,
});
const NAME_AT = [1699462, 3536271], EXAM_AT = [1219962, 4539577];
const BADGE_AT = [1135740, 1056323];

function boxRuns(xml, at) {
  for (const b of eng.topSpBlocks(xml)) {
    const off = eng.offOf(b.text);
    if (off && Math.abs(off[0] - at[0]) <= 45720 && Math.abs(off[1] - at[1]) <= 45720) return eng.runTexts(b.text);
  }
  return null;
}
function hasBadge(xml) {
  return eng.topSpBlocks(xml).some(b => {
    const off = eng.offOf(b.text);
    return off && Math.abs(off[0] - BADGE_AT[0]) <= 45720 && Math.abs(off[1] - BADGE_AT[1]) <= 45720 && /<a:blip/.test(b.text);
  });
}
function geomList(xml) {
  const out = [], re = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g;
  let m; while ((m = re.exec(xml)) !== null) out.push(m.slice(1).join(","));
  return out;
}
async function slidesOf(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort();
  const out = { zip, names, xml: {}, rels: {} };
  for (const n of names) {
    out.xml[n] = await zip.file(n).async("string");
    out.rels[n] = await zip.file(n.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels").async("string");
  }
  out.pres = await zip.file("ppt/presentation.xml").async("string");
  out.presRels = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  out.ct = await zip.file("[Content_Types].xml").async("string");
  return out;
}
async function templateSlides() {
  const zip = await JSZip.loadAsync(tplB64, { base64: true });
  const out = {};
  for (let s = 1; s <= 3; s++) {
    out[s] = await zip.file("ppt/slides/slide" + s + ".xml").async("string");
    out["rels" + s] = await zip.file("ppt/slides/_rels/slide" + s + ".xml.rels").async("string");
  }
  return out;
}

(async () => {
  // Scrub: placeholders in, dummy names out.
  ok(tplSrc.includes("LAST NAME, FIRST NAME M.I."), "template: student placeholder present");
  ok(tplSrc.includes("SIGNATORY NAME"), "template: signatory placeholder present");
  for (const bad of ["DELA CRUZ", "MARIA CLARA", "JUAN DELA CRUZ"]) {
    ok(!tplSrc.includes(bad), "template: no dummy name " + bad);
  }
  const tpl = await templateSlides();
  // Partner badge, positioned 2026-10-26: right edge 10.853in (0.147in print
  // margin) with a 0.100in gap to the RAKSO logo (was 0.071in margin, clipped
  // when printing; the first move to 9.555in left the logos almost touching).
  const BADGE_X = 8804758, BADGE_EXT = [1119477, 1027043];
  for (let sIdx = 1; sIdx <= 3; sIdx++) {
    const bx = tpl[sIdx].indexOf('Freeform 6');
    const bsp = tpl[sIdx].lastIndexOf('<p:sp>', bx);
    const bblk = tpl[sIdx].slice(bsp, tpl[sIdx].indexOf('</p:sp>', bsp));
    const boff = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(bblk);
    ok(!!boff, "template slide" + sIdx + ": partner badge present");
    eq(boff[1], String(BADGE_X), "template slide" + sIdx + ": badge positioned (right edge 10.853in, gap 0.100in)");
    eq(boff[3] + "," + boff[4], BADGE_EXT.join(","), "template slide" + sIdx + ": badge size unchanged");
  }
  // Wave frame: moved 2026-10-26 from the fixed output template (right/down/
  // smaller) so the accent no longer crosses the signatory text.
  const WAVE_AT = [3890321, 702918], WAVE_EXT = [9673279, 10310659];
  for (let sIdx = 1; sIdx <= 3; sIdx++) {
    let wf = null;
    for (const b of eng.topSpBlocks(tpl[sIdx])) { if (b.text.includes('name="Freeform 5"')) { wf = b; break; } }
    ok(!!wf, "template slide" + sIdx + ": wave shape present");
    eq(eng.offOf(wf.text).join(","), WAVE_AT.join(","), "template slide" + sIdx + ": wave moved right/down");
    eq(eng.extOf(wf.text).join(","), WAVE_EXT.join(","), "template slide" + sIdx + ": wave resized");
  }

  // A. Three exams: fill, order, assembly.
  const certsA = [cert("SANTOS, MARIA L.", "Excel"), cert("REYES, JOSE A.", "Word"), cert("CRUZ, ANA M.", "PowerPoint")];
  const outA = await slidesOf(await eng.generate(tplB64, certsA, SIG));
  eq(outA.names.length, 3, "A: 3 slides for 3 certs");
  eq((outA.pres.match(/<p:sldId /g) || []).length, 3, "A: sldIdLst has 3 entries");
  eq((outA.ct.match(/\/ppt\/slides\/slide\d+\.xml"/g) || []).length, 3, "A: 3 slide content-type overrides");
  eq((outA.presRels.match(/\/relationships\/slide"/g) || []).length, 3, "A: 3 slide presentation rels");
  const namesA = outA.names.map(n => boxRuns(outA.xml[n], NAME_AT)[0]);
  eq(namesA.join("|"), "SANTOS, MARIA L.|REYES, JOSE A.|CRUZ, ANA M.", "A: slide order matches cert order");
  const examA = boxRuns(outA.xml[outA.names[0]], EXAM_AT);
  eq(examA[0], "has successfully completed the ", "A: exam run 0 kept");
  ok(examA[1].startsWith("Microsoft Office Specialist \u2013 Excel Associate"), "A: exam title filled");
  ok(examA[1].endsWith(" "), "A: exam title keeps trailing space before SY run");
  eq(examA[2], "for the School Year 2025\u20132026.", "A: SY run filled");
  eq(examA[3], "Given this ", "A: given run 0 kept");
  eq(examA[4], "27th ", "A: given day run (underlined slot)");
  eq(examA[5], "day of March, 2026.", "A: given rest run");
  // Underline lives on the day run's rPr.
  const dayBlock = /<a:r>[\s\S]*?<a:t>27th <\/a:t>[\s\S]*?<\/a:r>/.exec(outA.xml[outA.names[0]]);
  ok(!!dayBlock && /u="sng"/.test(dayBlock[0]), "A: day run carries underline (rPr u=sng)");
  const titleBlock = /<a:r>[\s\S]*?<a:t>Microsoft Office Specialist[^<]*<\/a:t>[\s\S]*?<\/a:r>/.exec(outA.xml[outA.names[0]]);
  ok(!!titleBlock && /b="1"/.test(titleBlock[0]), "A: exam title run is bold (rPr b=1)");
  // Geometry + fonts untouched (off/ext pairs identical to template slide).
  eq(geomList(outA.xml[outA.names[0]]).join(";"), geomList(tpl[1]).join(";"), "A: Excel slide geometry identical to template");
  eq(geomList(outA.xml[outA.names[1]]).join(";"), geomList(tpl[3]).join(";"), "A: Word slide geometry identical to template");
  ok(hasBadge(outA.xml[outA.names[0]]) && hasBadge(outA.xml[outA.names[1]]) && hasBadge(outA.xml[outA.names[2]]), "A: badges present on all Associate slides");
  // Signatories written.
  const sigA = boxRuns(outA.xml[outA.names[0]], [1699462, 6324600]);
  eq(sigA[0], "DUMMY ONE", "A: signatory 1 name");
  eq(sigA[1], "Dummy Title One Dummy Org Inc.", "A: signatory 1 title + org");
  // No dummy names anywhere in the output.
  const allA = outA.names.map(n => outA.xml[n]).join("");
  for (const bad of ["DELA CRUZ", "MARIA CLARA", "SIGNATORY NAME", "LAST NAME, FIRST NAME"]) {
    ok(!allA.includes(bad), "A: output has no placeholder/dummy name " + bad);
  }

  // B. Expert: Excel slide minus badge, Expert title.
  const outB = await slidesOf(await eng.generate(tplB64, [cert("LIM, PEDRO S.", "Excel Expert")], SIG));
  eq(outB.names.length, 1, "B: 1 slide for 1 Expert cert");
  ok(!hasBadge(outB.xml[outB.names[0]]), "B: Expert slide has NO badge shape");
  const examB = boxRuns(outB.xml[outB.names[0]], EXAM_AT);
  ok(examB[1].includes("Excel Expert (Office 365 Apps)"), "B: Expert title (no Associate)");
  const tplGeom = geomList(tpl[1]), outGeom = geomList(outB.xml[outB.names[0]]);
  eq(outGeom.length, tplGeom.length - 1, "B: geometry list is template minus exactly the badge");
  eq(outGeom.join(";"), tplGeom.filter(g => g !== "1135740,1056323,1765573,1765573").join(";"), "B: all other geometry identical");
  const relCount = s => (s.match(/<Relationship /g) || []).length;
  eq(relCount(outB.rels[outB.names[0]]), relCount(tpl.rels1) - 1, "B: badge rel dropped from slide rels");

  // C. No school year: SY run empties, no dangling text.
  const outC = await slidesOf(await eng.generate(tplB64, [cert("TAN, LIZA R.", "Word", "")], SIG));
  const examC = boxRuns(outC.xml[outC.names[0]], EXAM_AT);
  eq(examC[2], "", "C: empty SY leaves SY run empty");

  // D. Signature image: media part, rel, pic element with fit box.
  const sigPng = fs.readFileSync(path.join(__dirname, "comp-sig.png"));
  const sigD = Object.assign({}, SIG, { s1: "data:image/png;base64," + sigPng.toString("base64") });
  const outD = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sigD));
  ok(!!outD.zip.file("ppt/media/sig1.png"), "D: sig1.png media part present");
  ok(/sig1\.png/.test(outD.rels[outD.names[0]]), "D: slide rels reference sig1.png");
  const pic = /<p:pic>.*?<\/p:pic>/.exec(outD.xml[outD.names[0]]);
  ok(!!pic, "D: signature <p:pic> inserted");
  if (pic) {
    const g = /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(pic[0]);
    ok(!!g, "D: pic has frame");
    if (g) {
      ok(+g[3] <= 1737360 && +g[4] <= 457200, "D: pic fits 1.9x0.5in box (got " + g[3] + "x" + g[4] + ")");
      eq(+g[2] + +g[4], 6289548, "D: pic bottom sits at 6.878in (fixed template), just above the name");
    }
  }

  // E. Mixed 5-cert batch keeps input order.
  const certsE = [cert("A ONE", "Word"), cert("B TWO", "Excel"), cert("C THREE", "Excel Expert"),
    cert("D FOUR", "PowerPoint"), cert("E FIVE", "Excel")];
  const outE = await slidesOf(await eng.generate(tplB64, certsE, SIG));
  eq(outE.names.length, 5, "E: 5 slides for 5 certs");
  eq(outE.names.map(n => boxRuns(outE.xml[n], NAME_AT)[0]).join("|"), "A ONE|B TWO|C THREE|D FOUR|E FIVE", "E: input order preserved");

  // G. Signature geometry: explicit 100%/0/0 reproduces the default frame;
  // size scales the fit box around its center; nudges shift by exact EMU inches.
  const sigG0 = Object.assign({}, sigD, { g1: { size: 100, dx: 0, dy: 0 } });
  const outG0 = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sigG0));
  const picD = /<p:pic>.*?<\/p:pic>/.exec(outD.xml[outD.names[0]]);
  const picG0 = /<p:pic>.*?<\/p:pic>/.exec(outG0.xml[outG0.names[0]]);
  const fr = (m) => /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(m[0]).slice(1).map(Number);
  const fD = fr(picD), fG0 = fr(picG0);
  eq(fG0.join(","), fD.join(","), "G: explicit default geometry = default frame");
  const sigG = Object.assign({}, sigD, { g1: { size: 150, dx: 0.1, dy: -0.05 } });
  const outG = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sigG));
  const fG = fr(/<p:pic>.*?<\/p:pic>/.exec(outG.xml[outG.names[0]]));
  ok(Math.abs(fG[2] - fD[2] * 1.5) <= 2 && Math.abs(fG[3] - fD[3] * 1.5) <= 2, "G: 150% scales ext 1.5x (got " + fG[2] + "x" + fG[3] + " vs " + fD[2] + "x" + fD[3] + ")");
  ok(Math.abs((fG[0] + fG[2] / 2) - (fD[0] + fD[2] / 2) - 91440) <= 2, "G: dx +0.1in shifts center +91440 EMU");
  ok(Math.abs((fG[1] + fG[3]) - (fD[1] + fD[3]) + 45720) <= 2, "G: dy -0.05in shifts pic bottom -45720 EMU (bottom-anchored)");
  const gmE = eng.geomOf({ size: 500, dx: 2, dy: -9 });
  ok(gmE.size === 150 && gmE.dx === 0.5 && gmE.dy === -0.5, "G: engine geomOf clamps to 150");

  // H. Signature media content types: every sig part gets a Default with
  // the correct MIME (a JPEG sig used to ship as sig1.jpg with NO Default —
  // the template covers "jpeg", not "jpg" — and PowerPoint demanded repair).
  const sigJ = Object.assign({}, SIG, { s1: "data:image/jpeg;base64,/9j/4AAQSkZJRg==" });
  const outJ = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sigJ));
  ok(!!outJ.zip.file("ppt/media/sig1.jpg"), "H: sig1.jpg media part present");
  ok(outJ.ct.includes('Extension="jpg" ContentType="image/jpeg"'), "H: jpg Default with image/jpeg");
  ok(outJ.ct.includes('Extension="png" ContentType="image/png"'), "H: png Default still present");
  const sigGif = Object.assign({}, SIG, { s1: "data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==" });
  const outGif = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sigGif));
  ok(outGif.ct.includes('Extension="gif" ContentType="image/gif"'), "H: gif Default with image/gif");

  // I. Name boxes NEVER move with the picture (image independence):
  // default and shifted geometry give the identical name-box position.
  function nameBoxOff(xml, want) {
    for (const b of eng.topSpBlocks(xml)) {
      if (eng.runTexts(b.text)[0] === want) return eng.offOf(b.text);
    }
    return null;
  }
  const nbD = nameBoxOff(outD.xml[outD.names[0]], "DUMMY ONE");
  const nbG = nameBoxOff(outG.xml[outG.names[0]], "DUMMY ONE");
  ok(!!nbD && !!nbG, "I: name boxes found");
  if (nbD && nbG) {
    eq(nbD.join(","), "1699462,6324600", "I: default name box = fixed-template position");
    eq(nbG.join(","), "1699462,6324600", "I: shifted geometry leaves the name box UNMOVED");
  }

  // J. Picture structure matches what PowerPoint writes (regression:
  // omitting <p:nvPr/> made PowerPoint refuse the whole file while
  // LibreOffice still opened it).
  const sig2 = Object.assign({}, SIG, { s1: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", s2: "data:image/jpeg;base64,/9j/4AAQSkZJRg==" });
  const outJ2 = await slidesOf(await eng.generate(tplB64, [cert("SANTOS, MARIA L.", "Excel")], sig2));
  const picJ = outJ2.xml[outJ2.names[0]].match(/<p:pic>[\s\S]*?<\/p:pic>/g) || [];
  ok(picJ.length === 2, "J: two signature pictures on the slide");
  ok(picJ.every(x => x.includes("<p:nvPr/>")), "J: every pic carries <p:nvPr/>");
  ok(picJ[0].includes('name="Signature 1"') && picJ[1].includes('name="Signature 2"'), "J: pic names unique per slot");
  ok(picJ[0].includes('descr="sig1.png"') || picJ[0].includes('descr="sig1.jpg"'), "J: pic descr names its media");

  // F. Errors fail loudly.
  try { await eng.generate(tplB64, [], SIG); ok(false, "F: empty batch throws"); }
  catch (e) { ok(/No certificates/.test(e.message), "F: empty batch throws"); }
  try { await eng.generate(tplB64, [Object.assign(cert("X", "Excel"), { given: "sometime" })], SIG); ok(false, "F: bad given throws"); }
  catch (e) { ok(/cannot parse given/.test(e.message), "F: bad given throws"); }
  try { await eng.generate(tplB64, [cert("X", "AZ-900")], SIG); ok(false, "F: unknown family throws"); }
  catch (e) { ok(/Unsupported exam family/.test(e.message), "F: unknown family throws"); }
  try { await eng.generate("", [cert("X", "Excel")], SIG); ok(false, "F: missing template throws"); }
  catch (e) { ok(/not installed/.test(e.message), "F: missing template throws"); }

  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error("FATAL - " + (e && e.stack || e)); process.exit(1); });
