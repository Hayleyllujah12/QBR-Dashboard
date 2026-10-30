/* MOS > Completion certificates — layout model (shared by the HTML
 * certificate and the PPTX export).
 *
 * Phase A (2026-10-27): the title, "This is to certify that", the student
 * name and the exam/given block share ONE centre line — the centre of the
 * content area right of the accent bars (0.942in → 11in = 5.971in). The
 * template boxes had four different centre lines (5.62 / 5.76 / 5.84 /
 * 6.75in), so the lines visibly drifted. The certify line is now centred
 * (it was a left-aligned box), and the title is centred on the same line —
 * it still clears the exam badge (title text ≈5.4in wide → starts ≈3.24in,
 * badge ends 3.173in).
 *
 * Phase B: a per-object layout the user adjusts in Completion → "Adjust
 * layout": x/y nudge in inches and a size % (font size for text, picture
 * size for images). One layout per browser (localStorage, like the
 * signatories). Values are clamped so a stale or hand-edited record can
 * never throw an object off the page. The SAME numbers drive the HTML
 * print and the PPTX export (mos-completion-pptx.js), so both match.
 */
(function (root) {
  root.QBR = root.QBR || {};
  var QBR = root.QBR;
  var KEY = "qbr-mos-completion-layout-v1";
  var PAGE_W = 11, PAGE_H = 8.5;
  var CENTER_X = 5.971; // (0.942 + 11) / 2 — centre of the area right of the bars
  var NUDGE = 5; // max x/y nudge, inches (the box centre is also kept on the page)

  // Template boxes in inches (master template, slide 1). `at` = EMU anchor
  // the PPTX engine uses to find the shape. `centred` = on the shared line.
  var BOX = {
    title:   { label: "Title", kind: "text", l: 3.473, t: 1.225, w: 6.544, h: 1.861, at: [3175802, 1120085], centred: true },
    certify: { label: "“This is to certify that”", kind: "text", l: 4.494, t: 3.212, w: 2.251, h: 0.507, at: [4109171, 2937405], centred: true },
    name:    { label: "Student name", kind: "text", l: 1.859, t: 3.867, w: 7.799, h: 0.785, at: [1699462, 3536271], centred: true },
    exam:    { label: "Exam + given date", kind: "text", l: 1.334, t: 4.965, w: 9.014, h: 0.912, at: [1219962, 4539577], centred: true },
    badge:   { label: "MOS badge", kind: "image", l: 1.242, t: 1.155, w: 1.931, h: 1.931, at: [1135740, 1056323] },
    logos:   { label: "RAKSO + partner logos", kind: "image", l: 7.523, t: 0, w: 3.407, h: 1.123, at: [6878694, 0], group: true },
    sig1:    { label: "Signatory 1 (name block)", kind: "text", l: 1.859, t: 6.917, w: 2.559, h: 0.668, at: [1699462, 6324600] },
    sig2:    { label: "Signatory 2 (name block)", kind: "text", l: 7.048, t: 6.917, w: 2.803, h: 0.668, at: [6444720, 6324600] }
  };
  var ORDER = ["title", "badge", "certify", "name", "exam", "sig1", "sig2", "logos"];
  function sizeRange(id) { return BOX[id].kind === "image" ? [50, 200] : [50, 150]; }

  // Phase A base left edge (inches): centred boxes sit on CENTER_X.
  function baseLeft(id) {
    var b = BOX[id];
    return b.centred ? Math.round((CENTER_X - b.w / 2) * 10000) / 10000 : b.l;
  }
  function r2(x) { return Math.round(x * 100) / 100; }
  function num(x, fb) { var n = parseFloat(x); return isFinite(n) ? n : fb; }
  function cl(x, lo, hi, fb) { return Math.min(hi, Math.max(lo, num(x, fb))); }

  function objOf(id, v) {
    v = v || {};
    var sr = sizeRange(id), b = BOX[id], s = cl(v.size, sr[0], sr[1], 100);
    // Nudge limit, then keep the box centre on the page.
    var dx = r2(cl(v.dx, -NUDGE, NUDGE, 0)), dy = r2(cl(v.dy, -NUDGE, NUDGE, 0));
    var cx = baseLeft(id) + b.w / 2 + dx, cy = b.t + b.h / 2 + dy;
    if (cx < 0) dx = r2(dx - cx); if (cx > PAGE_W) dx = r2(dx - (cx - PAGE_W));
    if (cy < 0) dy = r2(dy - cy); if (cy > PAGE_H) dy = r2(dy - (cy - PAGE_H));
    return { dx: dx, dy: dy, size: Math.round(s) };
  }
  function defaults() {
    var o = {}; ORDER.forEach(function (id) { o[id] = { dx: 0, dy: 0, size: 100 }; }); return o;
  }
  function layoutOf(v) {
    v = v && typeof v === "object" ? v : {};
    var o = {}; ORDER.forEach(function (id) { o[id] = objOf(id, v[id]); }); return o;
  }
  function isDefault(lay) {
    lay = layoutOf(lay);
    return ORDER.every(function (id) { var o = lay[id]; return !o.dx && !o.dy && o.size === 100; });
  }
  function load() {
    var s = null;
    try { s = JSON.parse(root.localStorage.getItem(KEY) || "null"); } catch (e) { s = null; }
    return layoutOf(s);
  }
  function save(lay) {
    try { root.localStorage.setItem(KEY, JSON.stringify(layoutOf(lay))); return true; }
    catch (e) { return false; }
  }
  // Final box (inches) after Phase A + the user's nudge/size (size scales
  // images around their centre; text keeps its box, the font scales).
  function boxOf(id, lay) {
    var b = BOX[id], o = layoutOf(lay)[id];
    var l = baseLeft(id) + o.dx, t = b.t + o.dy, w = b.w, h = b.h;
    if (b.kind === "image") {
      var k = o.size / 100;
      l += w * (1 - k) / 2; t += h * (1 - k) / 2; w *= k; h *= k;
    }
    return { l: l, t: t, w: w, h: h };
  }
  // True when any part of the object would print outside the page.
  function outside(id, lay) {
    var x = boxOf(id, lay), e = 0.001;
    return x.l < -e || x.t < -e || x.l + x.w > PAGE_W + e || x.t + x.h > PAGE_H + e;
  }
  // dx that puts the object's centre on the shared centre line.
  function centreDx(id) { return r2(CENTER_X - (baseLeft(id) + BOX[id].w / 2)); }

  // Inline style for the HTML element(s). Text: translate + --fs font
  // scale (the CSS font sizes are calc(<pt> * var(--fs,1))). Images:
  // translate + scale around the object's centre (logos: the group centre,
  // so the RAKSO logo and the partner badge move and resize together, the
  // same way the PPTX scales their group).
  function htmlStyle(id, lay, part) {
    var b = BOX[id], o = layoutOf(lay)[id], tr = "translate(" + o.dx + "in," + o.dy + "in)";
    if (b.kind === "text") return (o.dx || o.dy ? "transform:" + tr + ";" : "") + (o.size !== 100 ? "--fs:" + o.size / 100 + ";" : "");
    if (!o.dx && !o.dy && o.size === 100) return "";
    var origin = "50% 50%";
    if (id === "logos" && part) {
      var gcx = b.l + b.w / 2, gcy = b.t + b.h / 2;
      origin = r4(gcx - part.l) + "in " + r4(gcy - part.t) + "in";
    }
    return "transform:" + tr + " scale(" + o.size / 100 + ");transform-origin:" + origin + ";";
  }
  function r4(x) { return Math.round(x * 10000) / 10000; }
  // Logo parts inside the logos group (inches, template positions).
  var LOGO_PARTS = { rakso: { l: 7.523, t: 0.336 }, partner: { l: 9.629, t: 0 } };

  var api = { KEY: KEY, CENTER_X: CENTER_X, PAGE_W: PAGE_W, PAGE_H: PAGE_H, NUDGE: NUDGE, BOX: BOX, ORDER: ORDER,
    LOGO_PARTS: LOGO_PARTS, sizeRange: sizeRange, baseLeft: baseLeft, defaults: defaults, layoutOf: layoutOf,
    isDefault: isDefault, load: load, save: save, boxOf: boxOf, outside: outside, centreDx: centreDx, htmlStyle: htmlStyle };
  QBR.mosCompletionLayout = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : this);
