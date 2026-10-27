/* MOS completion certificates — failed-twice students, bulk generated.
 * New MOS sub-view (Results | Certificates | Reconcile | Completion):
 * the current Results filters select the batch (no per-student tick step),
 * one certificate per failed exam, grouped by school and exam type.
 * Dynamic parts: exam badge (Associate apps; Excel Expert has NO badge),
 * student name (LAST, FIRST MI), exam title, batch given date.
 * Signatory names ship BLANK (titles default to the two roles only); the requester
 * enters names once in his browser, where they persist in localStorage.
 * No personal data in this file: names in tests are invented dummy data.
 */
(function (root) {
  root.QBR = root.QBR || {};
  var QBR = root.QBR;
  function esc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function $(id) { return document.getElementById(id); }
  var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  function ordinal(d) {
    d = +d;
    if (d % 100 >= 11 && d % 100 <= 13) return d + "th";
    if (d % 10 === 1) return d + "st";
    if (d % 10 === 2) return d + "nd";
    if (d % 10 === 3) return d + "rd";
    return d + "th";
  }
  function isoToday() {
    var t = new Date();
    return t.getFullYear() + "-" + String(t.getMonth() + 1).padStart(2, "0") + "-" + String(t.getDate()).padStart(2, "0");
  }
  function givenText(iso) {
    if (!iso) return "";
    var y = iso.slice(0, 4), m = +iso.slice(5, 7), d = +iso.slice(8, 10);
    if (!y || !m || !d) return "";
    return "Given this " + ordinal(d) + " day of " + MONTHS[m - 1] + ", " + y + ".";
  }
  function syText(sy) {
    var s = String(sy || "").replace(/^SY\s+/i, "").trim();
    return s.replace(/-/g, "\u2013");
  }
  // Student name → "LAST, FIRST MI". Comma form is already in target shape,
  // so it is trusted verbatim (case/spacing normalized only) — this keeps
  // compound first names intact. Non-comma form ("First [Middle] Last") is
  // converted; a missing middle initial is omitted, never placeholdered.
  function formatCertName(raw) {
    var s = String(raw || "").replace(/\s+/g, " ").trim();
    if (!s) return "";
    if (s.indexOf(",") >= 0) {
      var p = s.split(",");
      var last = p[0].trim().toUpperCase();
      var rest = p.slice(1).join(" ").replace(/\s+/g, " ").trim().toUpperCase();
      return rest ? last + ", " + rest : last;
    }
    var t = s.split(" ").filter(function (x) { return !!x; });
    if (t.length === 1) return t[0].toUpperCase();
    var last2 = t[t.length - 1].toUpperCase(), first2 = t[0].toUpperCase();
    var mid = t.slice(1, -1), mi = "";
    if (mid.length) {
      var tok = mid[mid.length - 1].replace(/\./g, "");
      mi = (tok.length === 1 ? tok : mid[0].charAt(0)).toUpperCase();
    }
    return last2 + ", " + first2 + (mi ? " " + mi + "." : "");
  }
  var MOS_FAMS = ["Excel", "Excel Expert", "Word", "PowerPoint"];
  function familyOf(exam) {
    try {
      if (QBR.mosParse && QBR.mosParse.examFamily) return QBR.mosParse.examFamily(exam);
    } catch (e) {}
    var m = String(exam || "").match(/\b(Excel|Word|PowerPoint)\b(\s+Expert)?/i);
    if (!m) return String(exam || "");
    var base = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
    if (/^powerpoint$/i.test(m[1])) base = "PowerPoint";
    return base + (m[2] ? " Expert" : "");
  }
  function isMos(fam) { return MOS_FAMS.indexOf(fam) >= 0; }
  function examTitle(fam) {
    if (fam === "Excel Expert") return "Microsoft Office Specialist \u2013 Excel Expert (Office 365 Apps)";
    return "Microsoft Office Specialist \u2013 " + fam + " Associate (Office 365 Apps)";
  }
  function badgeFor(fam) {
    var L = QBR.mosCompletionLogos || {};
    if (fam === "Excel") return L.excel || "";
    if (fam === "Word") return L.word || "";
    if (fam === "PowerPoint") return L.powerpoint || "";
    return ""; // Excel Expert: NO logo, by the requester's explicit call — never fall back to Excel.
  }
  function partnerBadge() { return (QBR.mosCompletionLogos || {}).partner || ""; }
  function raksoLogo() { return (QBR.mosCompletionLogos || {}).rakso || ""; }
  /* 2026-10-27: accent colour follows the exam's badge — the same colours the PPTX template uses for the
   * left bar + stripe (Excel green, PowerPoint orange-red, Word blue); the wave art is recoloured to match. */
  var ACCENT = { "Excel": "#048642", "Excel Expert": "#048642", "PowerPoint": "#B8482B", "Word": "#25589B" };
  var WAVE_GREEN = /#048642/gi;
  function accentOf(fam) { return ACCENT[fam] || ACCENT.Excel; }
  var _waveCache = {};
  function waveArt(fam) {
    var w = (QBR.mosCompletionLogos || {}).wave || "", col = accentOf(fam);
    if (!w || col === ACCENT.Excel) return w;
    if (_waveCache[col]) return _waveCache[col];
    try {
      var m = /^data:image\/svg\+xml;base64,(.*)$/.exec(w);
      var svg = m ? atob(m[1]) : (/^data:image\/svg\+xml[,;]/.test(w) ? decodeURIComponent(w.slice(w.indexOf(",") + 1)) : null);
      if (svg == null) return w;
      return (_waveCache[col] = "data:image/svg+xml;base64," + btoa(svg.replace(WAVE_GREEN, col)));
    } catch (e) { return w; }
  }
  function isEligibleRow(r) { return !!r && !!r.noRetakesLeft && isMos(familyOf(r.exam)); }

  /* ---------------- signatories (this browser only) ---------------- */
  var SIG_KEY = "qbr-mos-completion-sig-v1";
  function sigDefaults() {
    return { n1: "", t1: "Certifications Manager", n2: "", t2: "Education Technology Director",
      org: "Rakso Computer Technology, Inc.", s1: "", s2: "",
      g1: { size: 100, dx: 0, dy: 0 }, g2: { size: 100, dx: 0, dy: 0 } };
  }
  function loadSig() {
    var d = sigDefaults(), s = {};
    try { s = JSON.parse(localStorage.getItem(SIG_KEY) || "{}") || {}; } catch (e) { s = {}; }
    var o = {};
    Object.keys(d).forEach(function (k) { o[k] = s[k] != null && s[k] !== "" ? s[k] : d[k]; });
    // Names ship blank: an empty stored name must stay empty, not revert to a default.
    o.n1 = s.n1 || ""; o.n2 = s.n2 || "";
    o.g1 = geomOf(s.g1); o.g2 = geomOf(s.g2);
    return o;
  }
  function saveSig(s) {
    try { localStorage.setItem(SIG_KEY, JSON.stringify(s)); return true; }
    catch (e) { return false; }
  }
  // Signature geometry: size % of the default box, plus x/y nudges in inches.
  // Same values drive the HTML cert (0.9in image height) and the PPTX fit box.
  function geomOf(v) {
    function num(x, fb) { var n = parseFloat(x); return isFinite(n) ? n : fb; }
    function cl(x, lo, hi, fb) { return Math.min(hi, Math.max(lo, num(x, fb))); }
    v = v || {};
    return { size: cl(v.size, 50, 150, 100), dx: cl(v.dx, -0.5, 0.5, 0), dy: cl(v.dy, -0.5, 0.5, 0) };
  }
  // Signature image style: size + shift live on the IMAGE ONLY. The
  // signatory name/title/org block never moves — moving or resizing the
  // signature cannot disturb the signatories (the requester's call; an image dragged
  // over a name overlaps it by placement, it never pushes the name away).
  function sigImgStyle(gm) {
    gm = geomOf(gm);
    var h = Math.round(0.9 * gm.size / 100 * 1000) / 1000;
    return "height:" + h + "in;transform:translate(" + gm.dx + "in," + gm.dy + "in)";
  }
  // Read the editor sliders (fall back to stored geometry when absent).
  function readGeom(idx) {
    function num(id, fb) { var el = $(id); var n = el ? parseFloat(el.value) : NaN; return isFinite(n) ? n : fb; }
    var cur = loadSig(), d = idx === 1 ? cur.g1 : cur.g2;
    return geomOf({ size: num("mos-comp-g" + idx + "-size", d.size),
      dx: num("mos-comp-g" + idx + "-dx", d.dx), dy: num("mos-comp-g" + idx + "-dy", d.dy) });
  }
  // Surgical preview refresh: updates only the preview img + readouts, never
  // re-renders the view (a render would destroy picked-but-unsaved files).
  function refreshGeomPreview(idx) {
    var gm = readGeom(idx);
    var pv = $("mos-comp-g" + idx + "-prev"); if (!pv) return;
    var img = pv.querySelector("img"); if (img) img.setAttribute("style", sigImgStyle(gm));
    var sv = $("mos-comp-g" + idx + "-size-v"); if (sv) sv.textContent = gm.size + "%";
    var xv = $("mos-comp-g" + idx + "-dx-v"); if (xv) xv.textContent = (gm.dx > 0 ? "+" : "") + gm.dx + "in";
    var yv = $("mos-comp-g" + idx + "-dy-v"); if (yv) yv.textContent = (gm.dy > 0 ? "+" : "") + gm.dy + "in";
  }
  function readSigFile(file, slot, done) {
    if (!file) { done(""); return; }
    var rd = new FileReader();
    rd.onload = function () { done(String(rd.result || "")); };
    rd.onerror = function () { done(""); };
    rd.readAsDataURL(file);
  }
  // Shrink a signature image to fit browser storage (~5 MB per origin).
  // A phone photo (several MB) becomes at most 1000px on its long side —
  // typically tens of KB. Transparency is PRESERVED: PNG/GIF/WebP stay PNG
  // (no white fill), because the PPTX certificate has a green graphic behind
  // the right signatory and a JPEG white box would show on it. JPEG/BMP
  // (no alpha) stay JPEG q0.85. SVG is rasterized to PNG when its size can
  // be read (width/height or viewBox); anything undecodable passes through
  // untouched and the old quota message still applies if it then fails.
  function downscaleSig(dataUrl, done) {
    if (!/^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml)/i.test(dataUrl || "")) { done(dataUrl); return; }
    var isSvg = /^data:image\/svg\+xml/i.test(dataUrl || "");
    var keepAlpha = isSvg || /^data:image\/(png|gif|webp)/i.test(dataUrl || "");
    // PNG output is capped at ~800 KB: a real signature (sparse ink) is
    // tens of KB at 1000px and never shrinks, but a noisy photo PNG can be
    // megabytes — then redraw smaller (1000 -> 750 -> 562 -> 421px) until
    // it fits, instead of failing on the storage quota. JPEG needs no cap.
    function emitPng(cv, ctx, maxDim, tries) {
      if (!cv.toBlob) { done(dataUrl); return; }
      cv.toBlob(function (b) {
        if (!b) { done(dataUrl); return; }
        if (b.size > 800000 && tries < 3 && maxDim > 350) {
          drawAt(Math.round(maxDim * 0.75), tries + 1);
          return;
        }
        var rd2 = new FileReader();
        rd2.onload = function () { done(String(rd2.result || dataUrl)); };
        rd2.onerror = function () { done(dataUrl); };
        rd2.readAsDataURL(b);
      }, "image/png");
    }
    function emitJpeg(cv) {
      if (!cv.toBlob) { done(dataUrl); return; }
      cv.toBlob(function (b) {
        if (!b) { done(dataUrl); return; }
        var rd2 = new FileReader();
        rd2.onload = function () { done(String(rd2.result || dataUrl)); };
        rd2.onerror = function () { done(dataUrl); };
        rd2.readAsDataURL(b);
      }, "image/jpeg", 0.85);
    }
    function svgDims() {
      try {
        var txt = decodeURIComponent((dataUrl.split(",", 2)[1] || ""));
        if (!/^\s*</.test(txt)) txt = String(atob(dataUrl.split(",", 2)[1] || ""));
        var w = /width\s*=\s*"?([0-9.]+)/i.exec(txt), h = /height\s*=\s*"?([0-9.]+)/i.exec(txt);
        if (w && h) return [+w[1], +h[1]];
        var vb = /viewBox\s*=\s*"?[\-0-9.]+\s+[\-0-9.]+\s+([0-9.]+)\s+([0-9.]+)/i.exec(txt);
        if (vb) return [+vb[1], +vb[2]];
      } catch (x) {}
      return null;
    }
    try {
      var img = new Image();
      var cv = document.createElement("canvas");
      function drawAt(maxDim, tries) {
        var sc = Math.min(1, maxDim / Math.max(imgW, imgH));
        cv.width = Math.max(1, Math.round(imgW * sc)); cv.height = Math.max(1, Math.round(imgH * sc));
        var ctx = cv.getContext("2d");
        if (!keepAlpha) { ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, cv.width, cv.height); }
        ctx.drawImage(img, 0, 0, cv.width, cv.height);
        if (keepAlpha) emitPng(cv, ctx, maxDim, tries);
        else emitJpeg(cv);
      }
      var imgW = 0, imgH = 0;
      img.onload = function () {
        try {
          imgW = img.naturalWidth || img.width; imgH = img.naturalHeight || img.height;
          if ((!imgW || !imgH) && isSvg) { var d = svgDims(); if (d) { imgW = d[0]; imgH = d[1]; } }
          if (!imgW || !imgH) { done(dataUrl); return; }
          drawAt(1000, 0);
        } catch (x) { done(dataUrl); }
      };
      img.onerror = function () { done(dataUrl); };
      img.src = dataUrl;
    } catch (x) { done(dataUrl); }
  }
  function processSigFile(file, done) {
    readSigFile(file, 0, function (d) { if (!d) { done(""); return; } downscaleSig(d, done); });
  }
  // Save a picked signature image immediately, without waiting for the Save
  // button. The old flow lost the pick whenever the view re-rendered (leaving
  // Completion and coming back rebuilt the file input empty), then Save
  // reported success with no image. Now picking IS saving (image only);
  // names/titles/org still save via the Save button.
  function stageSigFile(file, slot, host) {
    if (!file) return;
    processSigFile(file, function (d) {
      if (!d) { setMsg("Could not read that signature image — try a PNG or JPEG photo."); return; }
      var cur = loadSig(); cur["s" + slot] = d;
      var saved = saveSig(cur);
      // Preserve typed-but-unsaved names/titles: render rebuilds the inputs
      // from storage, which would otherwise wipe what the user just typed
      // (the same wipe the auto-save itself was built to fix, one step over).
      var keep = {};
      ["n1", "t1", "n2", "t2", "org", "g1-size", "g1-dx", "g1-dy", "g2-size", "g2-dx", "g2-dy"].forEach(function (k) { var el = $("mos-comp-" + k); if (el) keep[k] = el.value; });
      render(host);
      Object.keys(keep).forEach(function (k) { var el = $("mos-comp-" + k); if (el) el.value = keep[k]; });
      refreshGeomPreview(1); refreshGeomPreview(2);
      setMsg(saved ? "Signature image saved in this browser." : "Could not save signature image — browser storage may be full (try a smaller image).");
    });
  }

  /* ---------------- batch ---------------- */
  var batchDate = "";
  function mosState() {
    try { return QBR.mos && QBR.mos._state ? QBR.mos._state() : null; } catch (e) { return null; }
  }
  function currentFilter() {
    var S = mosState(); if (!S) return null;
    var f = S.f || {};
    var o = { school: f.school || "", section: f.section || "", exam: f.exam || "", examNo: f.examNo || "",
      proctor: f.proctor || "", outcome: f.outcome || "", min: f.min || "", max: f.max || "",
      q: f.q || "", firstOnly: !!f.firstOnly };
    if (f.pmode === "sy" && f.pval) o.sy = f.pval;
    if (f.pmode === "year" && f.pval) o.year = f.pval;
    if (f.pmode === "month" && f.pval) o.month = f.pval;
    return o;
  }
  function eligible() {
    var S = mosState(); if (!S || !S.model) return [];
    var f = currentFilter();
    var rows = QBR.mosParse.filterAttempts(S.model.rows, f);
    return rows.filter(isEligibleRow);
  }
  function certFor(row, given) {
    var fam = familyOf(row.exam);
    return { name: formatCertName(row.rosterName || row.name), family: fam, title: examTitle(fam),
      school: row.school || "", section: row.section || "", sy: syText(row.sy), given: given,
      badge: badgeFor(fam), partner: partnerBadge(), rakso: raksoLogo(), wave: waveArt(fam), accent: accentOf(fam),
      key: (row.studentKey || "") + "|" + fam };
  }
  function groupCerts(certs) {
    var bySchool = {};
    certs.forEach(function (c) { (bySchool[c.school] = bySchool[c.school] || []).push(c); });
    return Object.keys(bySchool).sort().map(function (sch) {
      var byExam = {};
      bySchool[sch].forEach(function (c) { (byExam[c.family] = byExam[c.family] || []).push(c); });
      var exams = MOS_FAMS.filter(function (f) { return byExam[f]; }).map(function (f) {
        return { family: f, certs: byExam[f].sort(function (a, b) { return a.name.localeCompare(b.name); }) };
      });
      return { school: sch, exams: exams };
    });
  }
  function buildBatch() {
    if (!batchDate) batchDate = isoToday();
    var given = givenText(batchDate);
    var certs = eligible().map(function (r) { return certFor(r, given); });
    var groups = groupCerts(certs), flat = [];
    groups.forEach(function (g) { g.exams.forEach(function (e) { e.certs.forEach(function (c) { flat.push(c); }); }); });
    return { groups: groups, certs: flat, given: given, date: batchDate };
  }
  function count() { try { return eligible().length; } catch (e) { return 0; } }

  /* ---------------- certificate HTML (absolute template positions) ----------------
   * Landscape letter page, every block absolutely positioned at the fixed
   * template's coordinates (title, certify, name, exam paragraph, signatories),
   * so the printed/PDF output matches the template's design positions. The
   * exam paragraph flows like the template's exam block (title bold, given
   * day underlined). Partner badge top-right, exam badge left (Associate
   * apps only — Expert has none, and the text block keeps its position so the
   * layout stays identical), centered title/text, two signatories at the foot.
   * No school name on the certificate (the examples show none). */
  function sigBlock(sig, idx) {
    var nm = idx === 1 ? sig.n1 : sig.n2, ti = idx === 1 ? sig.t1 : sig.t2, im = idx === 1 ? sig.s1 : sig.s2;
    var gm = geomOf(idx === 1 ? sig.g1 : sig.g2);
    return '<div class="mos-comp-sig">' +
      (im ? '<img src="' + esc(im) + '" alt="" style="' + sigImgStyle(gm) + '">' : "") +
      '<div class="mos-comp-sig-name">' + esc(nm) + "</div>" +
      '<div class="mos-comp-sig-role">' + esc(ti) + "</div>" +
      '<div class="mos-comp-sig-org">' + esc(sig.org) + "</div></div>";
  }
  // Template style: the given line underlines the DAY only (the template's
  // day run carries u="sng"; the issued examples underline day + month, but
  // the filler follows the template — so does the HTML cert).
  function givenHtml(given) {
    return esc(given).replace(/^(Given this )(\S+)( day of )/, "$1<u>$2</u>$3");
  }
  // Citation break matches the PPTX: "(Office 365 Apps)" starts the next
  // line. The natural wrap differs by renderer (Edge vs Chromium vs
  // PowerPoint metrics), so the break is explicit, not left to metrics.
  function titleHtml(title) {
    var t = esc(title), i = t.indexOf(" (Office 365 Apps)");
    return i >= 0 ? t.slice(0, i + 1) + "<br>" + t.slice(i + 1) : t;
  }
  function certHtml(c, sig) {
    // Background layer FIRST (paints behind everything): the template's
    // green left bar + thin stripe (pure CSS) and the green wave art
    // (stretched to the template's Freeform frame, clipped by the cert's
    // overflow:hidden exactly like the slide edge clips it in PowerPoint).
    // The RAKSO logo sits at its template spot, left of the partner badge.
    return '<section class="mos-comp-cert" data-family="' + esc(c.family || "") + '" style="--mos-accent:' + esc(c.accent || "#048642") + '">' +
      '<div class="mos-comp-bg"><div class="mos-comp-bar"></div><div class="mos-comp-stripe"></div>' +
      (c.wave ? '<div class="mos-comp-wave" style="background-image:url(' + esc(c.wave) + ')"></div>' : "") + "</div>" +
      (c.rakso ? '<img class="mos-comp-rakso" src="' + esc(c.rakso) + '" alt="">' : "") +
      (c.partner ? '<img class="mos-comp-partner" src="' + esc(c.partner) + '" alt="">' : "") +
      (c.badge ? '<img class="mos-comp-badge" src="' + esc(c.badge) + '" alt="">' : "") +
      '<div class="mos-comp-body">' +
      '<div class="mos-comp-title">CERTIFICATE OF<br>COMPLETION</div>' +
      '<div class="mos-comp-certify">This is to certify that</div>' +
      '<div class="mos-comp-name">' + esc(c.name) + "</div>" +
      '<div class="mos-comp-exam">has successfully completed the <b>' + titleHtml(c.title) + "</b>" +
      (c.sy ? " for the School Year " + esc(c.sy) + "." : "") + " <br>" + givenHtml(c.given) + "</div>" +
      "</div>" +
      '<div class="mos-comp-sigs">' + sigBlock(sig, 1) + sigBlock(sig, 2) + "</div>" +
      "</section>";
  }

  function compUniq(arr) { var o = {}; arr.forEach(function (v) { if (v != null && v !== "") o[v] = 1; }); return Object.keys(o).sort(); }
  function compMonLabel(ym) { var M = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]; return ym && ym.length >= 7 ? M[+ym.slice(5, 7) - 1] + " " + ym.slice(0, 4) : ym; }
  // Completion filter row: the SAME shared Results filters (S.f), surfaced
  // here so the batch can be narrowed without leaving the view. Option lists
  // mirror mos.js update(): school cascades into exam/proctor/period, and a
  // current value that is no longer offered resets to "" (same as setOpts).
  // Distinct ids (mos-comp-f-*) — the Results controls (mos-f-*) must never
  // share an id with these, or document lookups break.
  function compFilterHtml() {
    var S = mosState(); if (!S || !S.model) return "";
    var f = S.f || {}, rows = S.model.rows || [];
    var schools = compUniq(rows.map(function (r) { return r.school; }));
    var inSchool = f.school ? rows.filter(function (r) { return r.school === f.school; }) : rows;
    var exams = compUniq(inSchool.map(function (r) { return r.exam; }));
    var proctors = compUniq(inSchool.map(function (r) { return r.proctor || "(none)"; }));
    if (f.exam && exams.indexOf(f.exam) < 0) f.exam = "";
    if (f.proctor && proctors.indexOf(f.proctor) < 0) f.proctor = "";
    var pv = [], pl = null;
    if (f.pmode === "sy") pv = compUniq(inSchool.map(function (r) { return r.sy; })).reverse();
    else if (f.pmode === "year") pv = compUniq(inSchool.map(function (r) { return r.year; })).reverse();
    else if (f.pmode === "month") { pv = compUniq(inSchool.map(function (r) { return r.month; })).reverse(); pl = compMonLabel; }
    if (f.pval && pv.indexOf(f.pval) < 0) f.pval = "";
    function opts(values, cur, allLabel, labelFn) {
      return '<option value="">' + esc(allLabel) + "</option>" + values.map(function (v) {
        return '<option value="' + esc(v) + '"' + (v === cur ? " selected" : "") + ">" + esc(labelFn ? labelFn(v) : v) + "</option>";
      }).join("");
    }
    var proctorHidden = proctors.length < 2 && !f.proctor && proctors[0] === "(none)";
    return '<label class="mos-fl"><span>School</span><select id="mos-comp-f-school" class="form-select form-select-sm">' + opts(schools, f.school, "All schools") + "</select></label>" +
      '<label class="mos-fl"><span>Exam</span><select id="mos-comp-f-exam" class="form-select form-select-sm">' + opts(exams, f.exam, "All exams") + "</select></label>" +
      '<label class="mos-fl' + (proctorHidden ? " d-none" : "") + '"><span>Proctor</span><select id="mos-comp-f-proctor" class="form-select form-select-sm">' + opts(proctors, f.proctor, "All proctors") + "</select></label>" +
      '<label class="mos-fl"><span>Period</span><select id="mos-comp-f-pmode" class="form-select form-select-sm">' +
        '<option value="all"' + (f.pmode === "all" ? " selected" : "") + ">All time</option>" +
        '<option value="sy"' + (f.pmode === "sy" ? " selected" : "") + ">School year</option>" +
        '<option value="year"' + (f.pmode === "year" ? " selected" : "") + ">Year</option>" +
        '<option value="month"' + (f.pmode === "month" ? " selected" : "") + ">Month</option></select></label>" +
      '<label class="mos-fl"><span>&nbsp;</span><select id="mos-comp-f-pval" class="form-select form-select-sm"' + (f.pmode === "all" ? " disabled" : "") + ">" + opts(pv, f.pval, f.pmode === "all" ? "—" : "All", pl) + "</select></label>" +
      '<button type="button" class="btn btn-sm btn-outline-secondary" id="mos-comp-f-reset">Reset filters</button>';
  }

  /* ---------------- view ---------------- */
  function filterSummary() {
    var S = mosState(); if (!S) return "No results loaded.";
    var f = S.f || {}, parts = [];
    if (f.school) parts.push("School: " + f.school);
    if (f.section) parts.push("Section: " + f.section);
    if (f.exam) parts.push("Exam: " + f.exam);
    if (f.examNo) parts.push("Exam No.: " + f.examNo);
    if (f.proctor) parts.push("Proctor: " + f.proctor);
    if (f.pmode !== "all" && f.pval) parts.push("Period: " + f.pval);
    if (f.outcome) parts.push("Result: " + f.outcome);
    if (f.q) parts.push("Search: " + f.q);
    return parts.length ? parts.join(" \u00b7 ") : "All schools, all exams, all periods";
  }
  function setMsg(t) { var el = $("mos-comp-msg"); if (el) el.textContent = t || ""; }
  function sigEditorHtml(sig) {
    function field(idx, label, nm, ti, im, gm) {
      gm = geomOf(gm);
      function fmt(v) { return (v > 0 ? "+" : "") + v + "in"; }
      return '<div class="mos-comp-sigedit"><div class="fw-bold mb-1">' + label + "</div>" +
        '<label class="small d-block mb-1">Name <input type="text" class="form-control form-control-sm" id="mos-comp-n' + idx + '" value="' + esc(nm) + '" placeholder="(blank until you enter it)"></label>' +
        '<label class="small d-block mb-1">Title <input type="text" class="form-control form-control-sm" id="mos-comp-t' + idx + '" value="' + esc(ti) + '"></label>' +
        '<label class="small d-block">Signature image <input type="file" class="form-control form-control-sm" id="mos-comp-s' + idx + '" accept="image/*"></label>' +
        (im ? '<div class="mt-1"><img src="' + esc(im) + '" alt="" style="height:44px;width:auto;max-width:100%;object-fit:contain;border:1px solid var(--line);background:#fff"> <button type="button" class="btn btn-sm btn-outline-secondary" id="mos-comp-rm' + idx + '">Remove signature</button></div>' : "") +
        '<div class="mt-2"><div class="small fw-bold">Signature size &amp; position <span class="text-muted fw-normal">(image only — names stay put)</span></div>' +
        '<label class="small d-block">Size <input type="range" id="mos-comp-g' + idx + '-size" min="50" max="150" step="1" value="' + gm.size + '"> <span id="mos-comp-g' + idx + '-size-v">' + gm.size + '%</span></label>' +
        '<label class="small d-block">Shift left/right <input type="range" id="mos-comp-g' + idx + '-dx" min="-0.5" max="0.5" step="0.05" value="' + gm.dx + '"> <span id="mos-comp-g' + idx + '-dx-v">' + fmt(gm.dx) + '</span></label>' +
        '<label class="small d-block">Shift up/down <input type="range" id="mos-comp-g' + idx + '-dy" min="-0.5" max="0.5" step="0.05" value="' + gm.dy + '"> <span id="mos-comp-g' + idx + '-dy-v">' + fmt(gm.dy) + '</span></label>' +
        '<button type="button" class="btn btn-sm btn-outline-secondary" id="mos-comp-g' + idx + '-reset">Reset size &amp; position</button></div>' +
        '<div class="mt-2" id="mos-comp-g' + idx + '-prev" style="border:1px solid var(--line);background:#fff;padding:8px"><div class="mos-comp-sig">' +
        (im ? '<img src="' + esc(im) + '" alt="" style="' + sigImgStyle(gm) + '">' : '<div class="small text-muted">Pick a signature image to preview it here.</div>') +
        '<div class="mos-comp-sig-name">' + esc(nm) + '</div>' +
        '<div class="mos-comp-sig-role">' + esc(ti) + '</div>' +
        '<div class="mos-comp-sig-org">' + esc(sig.org) + '</div></div></div>' +
        "</div>";
    }
    return '<div class="fw-bold mb-1">Signatories (saved in this browser)</div>' +
      '<div class="small text-muted mb-2">Names ship blank — enter them once here. Signature images save as soon as you pick them (shrunk to fit browser storage) and are stored in this browser only, never uploaded. Size and position sliders adjust the signature on both the printed certificate and the PPTX export — the preview under each signature updates as you move them.</div>' +
      '<label class="small d-block mb-2">Organization (both signatories) <input type="text" class="form-control form-control-sm" id="mos-comp-org" value="' + esc(sig.org) + '"></label>' +
      '<div class="row g-2"><div class="col-md-6">' + field(1, "Signatory 1", sig.n1, sig.t1, sig.s1, sig.g1) + '</div>' +
      '<div class="col-md-6">' + field(2, "Signatory 2", sig.n2, sig.t2, sig.s2, sig.g2) + "</div></div>" +
      '<button type="button" class="btn btn-sm btn-outline-primary mt-2" id="mos-comp-sigsave">Save signatories</button>';
  }
  function render(host) {
    if (!host) return;
    if (!mosState() || !mosState().model) {
      host.innerHTML = '<div class="card-box small text-muted">Load MOS results first — completion certificates are generated from failed-twice results.</div>';
      return;
    }
    if (!batchDate) batchDate = isoToday();
    var sig = loadSig(), fHtml = compFilterHtml(), batch = buildBatch();
    var groupsHtml = batch.groups.length ? batch.groups.map(function (g) {
      return "<div><b>" + esc(g.school) + "</b>: " + g.exams.map(function (e) { return esc(e.family) + " " + e.certs.length; }).join(" \u00b7 ") + "</div>";
    }).join("") : '<div class="text-muted">No failed-twice MOS students match the current filters.</div>';
    host.innerHTML =
      '<div class="card-box"><h5 class="mb-1">Completion certificates</h5>' +
      '<div class="small text-muted mb-2">Failed-twice students (two failed takes, no retakes left), MOS apps only. The batch follows the same filters as the Results view — adjust them below or there.</div>' +
      '<div class="small mb-1"><b>Filters:</b> ' + esc(filterSummary()) + "</div>" +
      '<div class="mos-frow mb-2">' + fHtml + "</div>" +
      '<div class="small mb-2"><b>Batch:</b> ' + batch.certs.length + " certificate" + (batch.certs.length === 1 ? "" : "s") + "</div>" +
      '<div class="small mb-3">' + groupsHtml + "</div>" +
      '<div class="d-flex gap-2 align-items-end flex-wrap mb-2">' +
      '<label class="small mb-0">Batch date <input type="date" class="form-control form-control-sm" id="mos-comp-date" value="' + esc(batchDate) + '"></label>' +
      '<button type="button" class="btn btn-sm btn-outline-secondary" id="mos-comp-preview-btn">Preview first certificate</button>' +
      '<button type="button" class="btn btn-sm btn-primary" id="mos-comp-print" ' + (batch.certs.length ? "" : "disabled") + ">Print\u2026</button>" +
      '<button type="button" class="btn btn-sm btn-primary" id="mos-comp-pdf" ' + (batch.certs.length ? "" : "disabled") + ">Save as PDF\u2026</button>" +
      '<button type="button" class="btn btn-sm btn-primary" id="mos-comp-pptx" ' + (batch.certs.length ? "" : "disabled") + ">Download PPTX\u2026</button>" +
      "</div>" +
      '<div class="small text-muted mb-2">Print opens the browser print dialog (pick a printer). Save as PDF opens the same dialog with the batch filename preset — pick \u201cSave as PDF\u201d as the destination. Download PPTX builds a PowerPoint file from the master template (one slide per certificate) \u2014 open it in PowerPoint, then Save as PDF for the fixed-layout export. Issuance is logged when you start any of these.</div>' +
      '<div class="small mb-2" id="mos-comp-msg"></div>' +
      '<div class="card-box mt-2">' + sigEditorHtml(sig) + "</div>" +
      '<div id="mos-comp-preview" class="mos-comp-preview mt-3 d-none"></div>' +
      "</div>";
    $("mos-comp-date").addEventListener("change", function (e) { batchDate = e.target.value || isoToday(); refreshPreview(); });
    $("mos-comp-preview-btn").addEventListener("click", function () { preview(); });
    $("mos-comp-print").addEventListener("click", function () { output("print"); });
    $("mos-comp-pdf").addEventListener("click", function () { output("pdf"); });
    $("mos-comp-pptx").addEventListener("click", function () { outputPptx(); });
    $("mos-comp-sigsave").addEventListener("click", function () {
      // Fresh read: the render-time snapshot goes stale once a picked image
      // auto-saves (see the change listeners below) — saving from it would
      // clobber the new image with the old one.
      var cur = loadSig();
      var s = { n1: $("mos-comp-n1").value.trim(), t1: $("mos-comp-t1").value.trim() || "Certifications Manager",
        n2: $("mos-comp-n2").value.trim(), t2: $("mos-comp-t2").value.trim() || "Education Technology Director",
        org: $("mos-comp-org").value.trim(), s1: cur.s1, s2: cur.s2, g1: readGeom(1), g2: readGeom(2) };
      var f1 = $("mos-comp-s1").files[0], f2 = $("mos-comp-s2").files[0];
      var after = function () {
        var saved = saveSig(s);
        render(host); // recreates the msg element, so set the message after it
        setMsg(saved ? "Signatories saved in this browser." : "Could not save signatories — browser storage may be full (try smaller signature images).");
      };
      // Fallback for a file picked but not yet auto-saved (change handler
      // still processing): same downscale path, so both routes store the
      // same bytes and last-write-wins stays harmless.
      if (f1) processSigFile(f1, function (d) { if (d) s.s1 = d; if (f2) processSigFile(f2, function (d2) { if (d2) s.s2 = d2; after(); }); else after(); });
      else if (f2) processSigFile(f2, function (d2) { if (d2) s.s2 = d2; after(); });
      else after();
    });
    var rm1 = $("mos-comp-rm1"), rm2 = $("mos-comp-rm2");
    if (rm1) rm1.addEventListener("click", function () { var s = loadSig(); s.s1 = ""; saveSig(s); render(host); });
    if (rm2) rm2.addEventListener("click", function () { var s = loadSig(); s.s2 = ""; saveSig(s); render(host); });
    var s1El = $("mos-comp-s1"), s2El = $("mos-comp-s2");
    if (s1El) s1El.addEventListener("change", function () { stageSigFile(s1El.files[0], 1, host); });
    if (s2El) s2El.addEventListener("change", function () { stageSigFile(s2El.files[0], 2, host); });
    [1, 2].forEach(function (idx) {
      ["size", "dx", "dy"].forEach(function (k) { var el = $("mos-comp-g" + idx + "-" + k); if (el) el.addEventListener("input", function () { refreshGeomPreview(idx); }); });
      var rs = $("mos-comp-g" + idx + "-reset");
      if (rs) rs.addEventListener("click", function () {
        var sz = $("mos-comp-g" + idx + "-size"), dx = $("mos-comp-g" + idx + "-dx"), dy = $("mos-comp-g" + idx + "-dy");
        if (sz) sz.value = "100"; if (dx) dx.value = "0"; if (dy) dy.value = "0";
        refreshGeomPreview(idx);
      });
    });
    function compSet(key, val) { var S = mosState(); if (!S) return; S.f[key] = val; render(host); }
    var cfS = $("mos-comp-f-school"); if (cfS) cfS.addEventListener("change", function (e) { var S = mosState(); if (!S) return; S.f.school = e.target.value; S.f.section = ""; S.f.examNo = ""; render(host); });
    var cfE = $("mos-comp-f-exam"); if (cfE) cfE.addEventListener("change", function (e) { compSet("exam", e.target.value); });
    var cfP = $("mos-comp-f-proctor"); if (cfP) cfP.addEventListener("change", function (e) { compSet("proctor", e.target.value); });
    var cfM = $("mos-comp-f-pmode"); if (cfM) cfM.addEventListener("change", function (e) { var S = mosState(); if (!S) return; S.f.pmode = e.target.value; S.f.pval = ""; render(host); });
    var cfV = $("mos-comp-f-pval"); if (cfV) cfV.addEventListener("change", function (e) { compSet("pval", e.target.value); });
    var cfR = $("mos-comp-f-reset"); if (cfR) cfR.addEventListener("click", function () { var S = mosState(); if (!S) return; S.f.school = ""; S.f.exam = ""; S.f.proctor = ""; S.f.pmode = "all"; S.f.pval = ""; S.f.section = ""; S.f.examNo = ""; render(host); });
  }
  function preview() {
    var batch = buildBatch(), box = $("mos-comp-preview");
    if (!box) return;
    if (!batch.certs.length) { box.classList.add("d-none"); setMsg("Nothing to preview — the current filters match no failed-twice MOS student."); return; }
    box.innerHTML = certHtml(batch.certs[0], loadSig());
    box.classList.remove("d-none");
    setMsg("Previewing the first certificate of " + batch.certs.length + " (" + batch.certs[0].name + " — " + batch.certs[0].family + ").");
  }
  function refreshPreview() {
    var box = $("mos-comp-preview");
    if (box && !box.classList.contains("d-none")) preview();
  }
  function batchFileName(batch) {
    var schools = [];
    batch.groups.forEach(function (g) { if (schools.indexOf(g.school) < 0) schools.push(g.school); });
    var s = "Completion Certificates" + (schools.length === 1 ? " - " + schools[0] : schools.length ? " - " + schools.length + " schools" : "") + " - " + batch.date;
    return s.replace(/[^A-Za-z0-9 _\-]+/g, "").replace(/\s+/g, " ").trim();
  }
  function logIssuance(batch) {
    var seam = QBR.mosCerts && QBR.mosCerts._commitCompletion;
    if (!seam) return Promise.resolve("Printed, but NOT logged — the certificate log is unavailable in this tab.");
    var recs = batch.certs.map(function (c) {
      return { name: c.name, exam: c.family, school: c.school, section: c.section, given: batch.date, at: new Date() };
    });
    try {
      var r = seam(recs, batch.certs.length + " completion certificate" + (batch.certs.length === 1 ? "" : "s") + " logged");
      return Promise.resolve(r).then(function (ok) {
        return ok ? "Issuance logged to the MOS Certificate Log (separate completion columns)."
          : "Printed, but NOT logged — the log was not saved (view-only log, or log creation declined). Use the Certificates tab to record it.";
      });
    } catch (e) { return Promise.resolve("Printed, but NOT logged — " + (e && e.message || e)); }
  }
  function output(kind) {
    var batch = buildBatch();
    if (!batch.certs.length) { setMsg("Nothing to output — the current filters match no failed-twice MOS student."); return; }
    var sig = loadSig();
    // Log first: if no log exists yet, the create-log confirm resolves before
    // the print dialog opens. Browsers cannot confirm paper-out, so logging
    // records issuance as triggered (same honesty as the mark flow).
    logIssuance(batch).then(function (note) {
      setMsg(note);
      var host = $("mos-completion-print"); if (host) host.parentNode.removeChild(host);
      host = document.createElement("div"); host.id = "mos-completion-print";
      host.innerHTML = batch.certs.map(function (c) { return certHtml(c, sig); }).join("");
      document.body.appendChild(host);
      document.body.classList.add("mos-completion-printing");
      var prevTitle = document.title;
      document.title = batchFileName(batch);
      var style = document.createElement("style"); style.id = "mos-completion-page-style";
      style.textContent = "@page{size:letter landscape;margin:0}";
      document.head.appendChild(style);
      var done = false;
      function cleanup() {
        if (done) return; done = true;
        document.body.classList.remove("mos-completion-printing"); document.title = prevTitle;
        var h = $("mos-completion-print"); if (h && h.parentNode) h.parentNode.removeChild(h);
        var st = $("mos-completion-page-style"); if (st) st.remove();
        window.removeEventListener("afterprint", cleanup);
      }
      window.addEventListener("afterprint", cleanup);
      var imgs = Array.prototype.slice.call(host.querySelectorAll("img"));
      Promise.race([Promise.all(imgs.map(function (im) { return im.decode ? im.decode().catch(function () {}) : Promise.resolve(); })), new Promise(function (r) { setTimeout(r, 3000); })])
        .then(function () { try { window.print(); } finally { setTimeout(cleanup, 60000); } });
    });
  }

  function downloadBlob(blob, name) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    // 60s, not 1s: revoking too early can cut off a slow download start
    // and leave a truncated file PowerPoint cannot read at all.
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 60000);
  }
  function outputPptx() {
    var batch = buildBatch();
    if (!batch.certs.length) { setMsg("Nothing to output \u2014 the current filters match no failed-twice MOS student."); return; }
    var btn = $("mos-comp-pptx");
    if (btn) btn.disabled = true;
    setMsg("Generating PPTX\u2026");
    var tpl = QBR.mosCompletionTemplate && QBR.mosCompletionTemplate.b64;
    var eng = QBR.mosCompletionPptx && QBR.mosCompletionPptx.generate;
    logIssuance(batch).then(function (note) {
      setMsg(note);
      if (!eng) throw new Error("PPTX engine is not loaded (js/mos-completion-pptx.js).");
      if (!tpl) throw new Error("PPTX template is not installed (js/mos-completion-template.js).");
      var sig0 = loadSig(), dropped = [];
      try {
        var pm = QBR.mosCompletionPptx._test && QBR.mosCompletionPptx._test.parseSigMedia;
        if (pm) { var md = pm(sig0); if (sig0.s1 && !md[0]) dropped.push(1); if (sig0.s2 && !md[1]) dropped.push(2); }
      } catch (x) {}
      return eng(tpl, batch.certs, sig0).then(function (bytes) {
        if (dropped.length) setMsg("Note: signatory " + dropped.join(" & ") + "'s signature format is not supported in PowerPoint export and was left out — use a PNG or JPEG signature.");
        return bytes;
      });
    }).then(function (bytes) {
      var blob = new Blob([bytes], { type: QBR.mosCompletionPptx.MIME });
      downloadBlob(blob, batchFileName(batch) + ".pptx");
      if (btn) btn.disabled = false;
      setMsg("PPTX downloaded: " + batch.certs.length + " certificate" + (batch.certs.length === 1 ? "" : "s") + ". Open it in PowerPoint, then Save as PDF for the fixed-layout export.");
    }).catch(function (e) {
      if (btn) btn.disabled = false;
      setMsg("PPTX failed: " + (e && e.message || e));
    });
  }

  QBR.mosCompletion = { render: render, count: count,
    _formatName: formatCertName, _examTitle: examTitle, _ordinal: ordinal, _givenText: givenText,
    _syText: syText, _familyOf: familyOf, _accentOf: accentOf, _waveArt: waveArt, _isEligible: isEligibleRow, _group: groupCerts,
    _certHtml: certHtml, _buildBatch: buildBatch, _sig: loadSig, _batchFileName: batchFileName, _geomOf: geomOf, _sigImgStyle: sigImgStyle };
  if (typeof module !== "undefined" && module.exports) module.exports = {
    formatCertName: formatCertName, examTitle: examTitle, ordinal: ordinal, givenText: givenText,
    syText: syText, familyOf: familyOf, isMos: isMos, isEligibleRow: isEligibleRow,
    groupCerts: groupCerts, certFor: certFor, certHtml: certHtml,
    geomOf: geomOf, sigImgStyle: sigImgStyle };
})(typeof window !== "undefined" ? window : this);
