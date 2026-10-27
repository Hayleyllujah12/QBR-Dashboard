/* MOS completion certificates — PPTX export (template fill + slide clone).
 * Builds one .pptx per batch from the scrubbed master template
 * (js/mos-completion-template.js): one slide per certificate, in batch order
 * (school, then exam, then name — the order QBR.mosCompletion.buildBatch
 * already produces). Exam slide map: Excel -> slide 1, PowerPoint -> slide 2,
 * Word -> slide 3; Excel Expert clones the Excel slide with the badge shape
 * removed (no badge, by the requester's explicit call). Signatory names, titles and
 * organization come from the saved signatories; signature images (when the
 * user uploaded any) are added as pictures above each signatory name.
 * Signature geometry (size/dx/dy) moves and scales the IMAGE ONLY — the
 * signatory name boxes never move (the requester's call).
 *
 * Implementation notes (read before touching this file):
 * - Slide XML is edited by string surgery, NOT DOMParser, so this exact file
 *   runs in the browser and in the Node tests with no extra dependency.
 * - Boxes are identified by their shape-frame offset (EMU) in the master
 *   template, with a 0.05in tolerance. If the master geometry changes, the
 *   box anchors below must move with it — the tests pin them.
 * - Only text runs change inside a box; geometry, fonts and fills are never
 *   touched, so the template layout survives the fill byte-for-byte in
 *   shape terms (verified by the tests: geometry + fonts identical).
 * - Zip work uses JSZip (libs/jszip.min.js, MIT). No personal data in this
 *   file: names in tests are invented dummy data.
 */
(function (root) {
  root.QBR = root.QBR || {};
  var QBR = root.QBR;
  var EMU_IN = 914400;
  var TOL = 45720; // 0.05in — box-anchor tolerance
  // Master-template box anchors (EMU, from the fixed master, slide 1):
  var NAME_AT = [1699462, 3536271]; // student name box (1.859in, 3.867in)
  var EXAM_AT = [1219962, 4539577]; // exam block (1.334in, 4.965in)
  var SIG_LEFT = [1699462, 6324600]; // signatory 1 (1.859in, 6.917in)
  var SIG_RIGHT = [6444720, 6324600]; // signatory 2 (7.048in, 6.917in)
  var BADGE_AT = [1135740, 1056323]; // exam badge frame (1.242in, 1.155in)
  var SRC_SLIDE = { "Excel": 1, "PowerPoint": 2, "Word": 3, "Excel Expert": 1 };
  var PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  var SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  var REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
  var REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  var SIG_MAX_W = 1737360; // 1.9in
  var SIG_MAX_H = 457200; // 0.5in
  var SIG_BOTTOM = 6289548; // 6.878in — signature bottom, just above the name (fixed template)
  var ROW_DROP_EMU = 0; // retired 2026-10-26: the fixed template now carries the
  // final signatory position, so output boxes equal template boxes (no drop).
  // Signature geometry from the Completion editor: size % of the fit box,
  // plus x/y nudges in inches. Clamped so a stale or hand-edited record
  // can never throw the picture off the slide.
  function geomOf(v) {
    function num(x, fb) { var n = parseFloat(x); return isFinite(n) ? n : fb; }
    function cl(x, lo, hi, fb) { return Math.min(hi, Math.max(lo, num(x, fb))); }
    v = v || {};
    return { size: cl(v.size, 50, 150, 100), dx: cl(v.dx, -0.5, 0.5, 0), dy: cl(v.dy, -0.5, 0.5, 0) };
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function near(a, b) { return Math.abs(a - b) <= TOL; }
  function bytesFromB64(b64) {
    if (typeof Buffer !== "undefined") return Buffer.from(b64, "base64");
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function u16BE(b, o) { return (b[o] << 8) | b[o + 1]; }
  function u32BE(b, o) { return (b[o] * 16777216) + (b[o + 1] << 16) + (b[o + 2] << 8) + b[o + 3]; }
  function u16LE(b, o) { return b[o] | (b[o + 1] << 8); }

  // Image aspect (w/h) from raw bytes; PNG, JPEG, GIF, WebP and BMP.
  // Falls back to 3 (a wide signature) when the format cannot be read.
  function imgAspect(b64, ext) {
    try {
      var b = bytesFromB64(b64);
      if (ext === "png" && b.length > 24 && b[0] === 137 && b[1] === 80) {
        var w = u32BE(b, 16), h = u32BE(b, 20);
        if (w > 0 && h > 0) return w / h;
      }
      if ((ext === "jpg" || ext === "jpeg") && b.length > 4 && b[0] === 0xFF && b[1] === 0xD8) {
        var i = 2;
        while (i + 9 < b.length) {
          if (b[i] !== 0xFF) break;
          var mk = b[i + 1];
          if (mk === 0xD9 || mk === 0xDA) break;
          if (mk >= 0xC0 && mk <= 0xCF && mk !== 0xC4 && mk !== 0xC8 && mk !== 0xCC) {
            var hh = u16BE(b, i + 5), ww = u16BE(b, i + 7);
            if (ww > 0 && hh > 0) return ww / hh;
            break;
          }
          i += 2 + u16BE(b, i + 2);
        }
      }
      if (ext === "gif" && b.length > 10 && b[0] === 71 && b[1] === 73) {
        var gw = u16LE(b, 6), gh = u16LE(b, 8);
        if (gw > 0 && gh > 0) return gw / gh;
      }
      if (ext === "webp" && b.length > 30 && b[0] === 82 && b[1] === 73 && b[2] === 70 && b[3] === 70 &&
          b[8] === 87 && b[9] === 69 && b[10] === 66 && b[11] === 80) {
        // Byte compare (not Buffer.toString): this also runs on Uint8Array
        // in the browser, where toString takes no encoding arguments.
        var fourCC = String.fromCharCode(b[12], b[13], b[14], b[15]);
        if (fourCC === "VP8L" && b.length > 25) {
          var packed = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24);
          var ww = (packed & 0x3FFF) + 1, hh = ((packed >> 14) & 0x3FFF) + 1;
          if (ww > 0 && hh > 0) return ww / hh;
        }
        if (fourCC === "VP8X" && b.length > 30) {
          var xw = (b[24] | (b[25] << 8) | (b[26] << 16)) + 1, xh = (b[27] | (b[28] << 8) | (b[29] << 16)) + 1;
          if (xw > 0 && xh > 0) return xw / xh;
        }
        if (fourCC === "VP8 " && b.length > 30) {
          var lw = ((b[26] | (b[27] << 8)) & 0x3FFF), lh = ((b[28] | (b[29] << 8)) & 0x3FFF);
          if (lw > 0 && lh > 0) return lw / lh;
        }
      }
      if (ext === "bmp" && b.length > 26 && b[0] === 0x42 && b[1] === 0x4D) {
        var bw = b.readInt32LE ? b.readInt32LE(18) : (b[18] | (b[19] << 8) | (b[20] << 16) | (b[21] << 24));
        var bh = b.readInt32LE ? b.readInt32LE(22) : (b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24));
        if (bw !== 0 && bh !== 0) return Math.abs(bw) / Math.abs(bh);
      }
    } catch (e) {}
    return 3;
  }

  // Top-level <p:sp> blocks of a slide, skipping shapes nested inside
  // <p:grpSp> groups (the top-right logo group lives in one — its members
  // must never match the box anchors below).
  function topSpBlocks(xml) {
    var blocks = [], i = 0, n = xml.length;
    while (i < n) {
      var sp = xml.indexOf("<p:sp>", i), grp = xml.indexOf("<p:grpSp>", i);
      if (sp < 0) break;
      if (grp >= 0 && grp < sp) {
        var depth = 1, j = grp + 9;
        while (depth > 0 && j < n) {
          var og = xml.indexOf("<p:grpSp>", j), cg = xml.indexOf("</p:grpSp>", j);
          if (cg < 0) { j = n; break; }
          if (og >= 0 && og < cg) { depth++; j = og + 9; }
          else { depth--; j = cg + 10; } // "</p:grpSp>" is 10 chars
        }
        i = j; continue;
      }
      var end = xml.indexOf("</p:sp>", sp);
      if (end < 0) break;
      blocks.push({ start: sp, end: end + 7, text: xml.slice(sp, end + 7) });
      i = end + 7;
    }
    return blocks;
  }
  function offOf(block) {
    var m = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(block);
    return m ? [+m[1], +m[2]] : null;
  }
  function extOf(block) {
    var m = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(block);
    return m ? [+m[1], +m[2]] : null;
  }
  function runTexts(block) {
    var out = [], re = /<a:t>([^<]*)<\/a:t>/g, m;
    while ((m = re.exec(block)) !== null) out.push(m[1]);
    return out;
  }
  // Replace run texts in order; fn(i, oldText) returns the new text, or
  // null/undefined to keep the old text. Throws on an unexpected run count
  // so template drift fails loudly instead of producing a wrong certificate.
  function replaceRuns(block, expected, fn) {
    var runs = runTexts(block);
    if (runs.length !== expected) throw new Error("unexpected run count " + runs.length + " (expected " + expected + ")");
    var idx = 0;
    return block.replace(/<a:t>([^<]*)<\/a:t>/g, function (whole, old) {
      var nw = fn(idx++, old);
      return "<a:t>" + esc(nw == null ? old : nw) + "</a:t>";
    });
  }
  function shiftBoxOff(xml, at, dxEmu, dyEmu) {
    var blocks = topSpBlocks(xml);
    for (var k = 0; k < blocks.length; k++) {
      var off = offOf(blocks[k].text);
      if (off && near(off[0], at[0]) && near(off[1], at[1])) {
        var shifted = blocks[k].text.replace(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/,
          '<a:off x="' + (off[0] + dxEmu) + '" y="' + (off[1] + dyEmu) + '"/>');
        return xml.slice(0, blocks[k].start) + shifted + xml.slice(blocks[k].end);
      }
    }
    throw new Error("box not found at " + at[0] + "," + at[1]);
  }
  function patchBox(xml, at, fn) {
    var blocks = topSpBlocks(xml);
    for (var k = 0; k < blocks.length; k++) {
      var off = offOf(blocks[k].text);
      if (off && near(off[0], at[0]) && near(off[1], at[1])) {
        return xml.slice(0, blocks[k].start) + fn(blocks[k].text) + xml.slice(blocks[k].end);
      }
    }
    throw new Error("box not found at " + at[0] + "," + at[1]);
  }
  function findBox(xml, at) {
    var blocks = topSpBlocks(xml);
    for (var k = 0; k < blocks.length; k++) {
      var off = offOf(blocks[k].text);
      if (off && near(off[0], at[0]) && near(off[1], at[1])) return blocks[k];
    }
    return null;
  }
  function parseGiven(given) {
    var m = /^Given this (\d+(?:st|nd|rd|th)) (day of .*)$/.exec(String(given || ""));
    if (!m) throw new Error("cannot parse given date: " + given);
    return { day: m[1] + " ", rest: m[2] };
  }

  function fillSlide(xml, cert, sig) {
    var parts = parseGiven(cert.given);
    xml = patchBox(xml, NAME_AT, function (blk) {
      return replaceRuns(blk, 1, function () { return cert.name; });
    });
    xml = patchBox(xml, EXAM_AT, function (blk) {
      return replaceRuns(blk, 6, function (i) {
        if (i === 1) return cert.title + " "; // trailing space: the SY run starts with "for"
        if (i === 2) return cert.sy ? "for the School Year " + cert.sy + "." : "";
        if (i === 4) return parts.day; // underlined run — the underline lives in rPr, untouched
        if (i === 5) return parts.rest;
        return null;
      });
    });
    function sigPatch(idx) {
      var nm = idx === 1 ? sig.n1 : sig.n2, ti = idx === 1 ? sig.t1 : sig.t2;
      var role = [ti || "", sig.org || ""].filter(function (x) { return !!x; }).join(" ");
      return function (blk) {
        return replaceRuns(blk, 2, function (i) {
          if (i === 0) return nm || "";
          if (i === 1) return role;
          return null;
        });
      };
    }
    xml = patchBox(xml, SIG_LEFT, sigPatch(1));
    xml = patchBox(xml, SIG_RIGHT, sigPatch(2));
    return xml;
  }

  function removeBadge(xml, relsXml) {
    var blocks = topSpBlocks(xml);
    for (var k = 0; k < blocks.length; k++) {
      var off = offOf(blocks[k].text);
      if (off && near(off[0], BADGE_AT[0]) && near(off[1], BADGE_AT[1]) && /<a:blip/.test(blocks[k].text)) {
        var rid = /<a:blip[^>]*r:embed="(rId\d+)"/.exec(blocks[k].text);
        xml = xml.slice(0, blocks[k].start) + xml.slice(blocks[k].end);
        if (rid) relsXml = relsXml.replace(new RegExp('<Relationship[^>]*Id="' + rid[1] + '"[^>]*/>'), "");
        return { xml: xml, relsXml: relsXml };
      }
    }
    throw new Error("badge shape not found (Expert slide)");
  }

  function maxId(xml, re) {
    var m, mx = 0;
    while ((m = re.exec(xml)) !== null) mx = Math.max(mx, +m[1]);
    return mx;
  }
  // Picture element mirrors the structure PowerPoint itself writes
  // (verified against a python-pptx file PowerPoint accepts): cNvPr with a
  // UNIQUE name + descr, cNvPicPr with picLocks, and the trailing <p:nvPr/>
  // that every non-visual block in the master template carries. An earlier
  // version omitted <p:nvPr/> — LibreOffice tolerated it, PowerPoint
  // refused the whole file ("can't read"). Never drop it again.
  function addSigPic(xml, relsXml, mediaName, aspect, boxOff, boxExt, geom, slot) {
    geom = geomOf(geom);
    var scale = Math.min(SIG_MAX_W / (SIG_MAX_H * aspect), 1);
    var H = Math.round(SIG_MAX_H * scale), W = Math.round(H * aspect);
    if (W > SIG_MAX_W) { W = SIG_MAX_W; H = Math.round(W / aspect); }
    // Bottom-anchored scaling: the pic bottom stays at anchor+dy while the
    // name box stays fixed, so resizing/moving the image never disturbs the
    // signatories. At 100%/0/0 this is the old placement shifted down by
    // ROW_DROP_EMU.
    var W2 = Math.round(W * geom.size / 100), H2 = Math.round(H * geom.size / 100);
    var cx = boxOff[0] + boxExt[0] / 2, anchor = SIG_BOTTOM + ROW_DROP_EMU;
    var X = Math.round(cx - W2 / 2 + geom.dx * EMU_IN), Y = Math.round(anchor - H2 + geom.dy * EMU_IN);
    var cnvId = maxId(xml, /<p:cNvPr id="(\d+)"/g) + 1;
    var ridNum = maxId(relsXml, /Id="rId(\d+)"/g) + 1;
    var rid = "rId" + ridNum;
    var pic = '<p:pic><p:nvPicPr><p:cNvPr id="' + cnvId + '" name="Signature ' + (slot + 1) + '" descr="' + mediaName + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="' + X + '" y="' + Y + '"/><a:ext cx="' + W2 + '" cy="' + H2 + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    if (xml.indexOf("</p:spTree>") < 0) throw new Error("spTree close tag not found");
    xml = xml.replace("</p:spTree>", pic + "</p:spTree>");
    relsXml = relsXml.replace("</Relationships>",
      '<Relationship Id="' + rid + '" Type="' + REL_IMAGE + '" Target="../media/' + mediaName + '"/></Relationships>');
    return { xml: xml, relsXml: relsXml };
  }
  function parseSigMedia(sig) {
    function slot(du, n) {
      if (!du) return null;
      var m = /^data:image\/([A-Za-z0-9.+-]+);base64,([\s\S]*)$/.exec(du);
      if (!m) return null;
      var ext = m[1].toLowerCase();
      if (ext === "jpeg") ext = "jpg";
      if (["png", "jpg", "gif", "webp", "bmp"].indexOf(ext) < 0) return null;
      return { name: "sig" + n + "." + ext, data: m[2], ext: ext, aspect: imgAspect(m[2], ext) };
    }
    return [slot(sig.s1, 1), slot(sig.s2, 2)];
  }

  // certs: [{name, family, title, sy, given}] in output order.
  // sig: {n1,t1,n2,t2,org,s1,s2} with s1/s2 data URLs (or ""), plus
  // optional g1/g2 geometry {size,dx,dy} from the Completion editor.
  // Returns a Promise<Uint8Array> of the .pptx bytes.
  async function generate(templateB64, certs, sig) {
    // NOTE: do not name this local JSZip — var hoisting would shadow the
    // global and break the typeof check below in Node (browser path works
    // only because root.JSZip short-circuits first).
    var ZipImpl = root.JSZip || (typeof JSZip !== "undefined" ? JSZip : null);
    if (!ZipImpl) throw new Error("JSZip is not loaded (libs/jszip.min.js).");
    if (!templateB64) throw new Error("PPTX template is not installed (js/mos-completion-template.js).");
    if (!certs || !certs.length) throw new Error("No certificates in this batch.");
    var zip = await ZipImpl.loadAsync(templateB64, { base64: true });
    var srcXml = [], srcRels = [];
    for (var s = 1; s <= 3; s++) {
      srcXml.push(await zip.file("ppt/slides/slide" + s + ".xml").async("string"));
      srcRels.push(await zip.file("ppt/slides/_rels/slide" + s + ".xml.rels").async("string"));
    }
    var media = parseSigMedia(sig || {});
    media.forEach(function (md) {
      if (md) zip.file("ppt/media/" + md.name, md.data, { base64: true });
    });
    for (var i = 0; i < certs.length; i++) {
      var c = certs[i];
      var si = SRC_SLIDE[c.family];
      if (si == null) throw new Error("Unsupported exam family: " + c.family);
      var xml = srcXml[si - 1], rels = srcRels[si - 1];
      xml = fillSlide(xml, c, sig || {});
      if (c.family === "Excel Expert") {
        var rb = removeBadge(xml, rels);
        xml = rb.xml; rels = rb.relsXml;
      }
      for (var slot = 0; slot < 2; slot++) {
        var md = media[slot];
        if (!md) continue;
        var box = findBox(xml, slot === 0 ? SIG_LEFT : SIG_RIGHT);
        if (!box) throw new Error("signatory box not found");
        var off = offOf(box.text), ext = extOf(box.text);
        if (!off || !ext) throw new Error("signatory box geometry not found");
        var gm = geomOf(slot === 0 ? sig.g1 : sig.g2);
        var added = addSigPic(xml, rels, md.name, md.aspect, off, ext, gm, slot);
        xml = added.xml; rels = added.relsXml;
        // Names stay put: the template carries the final row position, so
        // the row drop is 0 — never the user's image nudges (image
        // independence — the requester's call).
        xml = shiftBoxOff(xml, slot === 0 ? SIG_LEFT : SIG_RIGHT, 0, ROW_DROP_EMU);
      }
      zip.file("ppt/slides/slide" + (i + 1) + ".xml", xml);
      zip.file("ppt/slides/_rels/slide" + (i + 1) + ".xml.rels", rels);
    }
    for (var k = certs.length + 1; k <= 3; k++) {
      zip.remove("ppt/slides/slide" + k + ".xml");
      zip.remove("ppt/slides/_rels/slide" + k + ".xml.rels");
    }
    var pres = await zip.file("ppt/presentation.xml").async("string");
    var sldIds = "";
    for (var j = 0; j < certs.length; j++) sldIds += '<p:sldId id="' + (256 + j) + '" r:id="rId' + (100 + j) + '"/>';
    if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(pres)) {
      pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, "<p:sldIdLst>" + sldIds + "</p:sldIdLst>");
    } else throw new Error("sldIdLst not found in presentation.xml");
    zip.file("ppt/presentation.xml", pres);
    var presRels = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
    presRels = presRels.replace(/<Relationship[^>]*\/relationships\/slide"[^>]*\/>/g, "");
    var add = "";
    for (var q = 0; q < certs.length; q++) {
      add += '<Relationship Id="rId' + (100 + q) + '" Type="' + REL_SLIDE + '" Target="slides/slide' + (q + 1) + '.xml"/>';
    }
    presRels = presRels.replace("</Relationships>", add + "</Relationships>");
    zip.file("ppt/_rels/presentation.xml.rels", presRels);
    var ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace(/<Override[^>]*\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, "");
    var ctAdd = "";
    for (var t = 0; t < certs.length; t++) {
      ctAdd += '<Override PartName="/ppt/slides/slide' + (t + 1) + '.xml" ContentType="' + SLIDE_CT + '"/>';
    }
    ct = ct.replace("</Types>", ctAdd + "</Types>");
    var CT_MIME = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png",
      gif: "image/gif", webp: "image/webp", bmp: "image/bmp" };
    media.forEach(function (md) {
      if (md && CT_MIME[md.ext] && ct.indexOf('Extension="' + md.ext + '"') < 0) {
        ct = ct.replace("</Types>", '<Default Extension="' + md.ext + '" ContentType="' + CT_MIME[md.ext] + '"/></Types>');
      }
    });
    zip.file("[Content_Types].xml", ct);
    try {
      var appXml = await zip.file("docProps/app.xml").async("string");
      if (/<Slides>\d+<\/Slides>/.test(appXml)) {
        zip.file("docProps/app.xml", appXml.replace(/<Slides>\d+<\/Slides>/, "<Slides>" + certs.length + "</Slides>"));
      }
    } catch (x) {}
    return await zip.generateAsync({ type: "uint8array", mimeType: PPTX_MIME,
      compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  QBR.mosCompletionPptx = { generate: generate, MIME: PPTX_MIME,
    _test: { topSpBlocks: topSpBlocks, offOf: offOf, extOf: extOf, runTexts: runTexts,
      replaceRuns: replaceRuns, patchBox: patchBox, findBox: findBox, parseGiven: parseGiven,
      fillSlide: fillSlide, removeBadge: removeBadge, addSigPic: addSigPic,
      parseSigMedia: parseSigMedia, imgAspect: imgAspect, geomOf: geomOf, shiftBoxOff: shiftBoxOff } };
  if (typeof module !== "undefined" && module.exports) module.exports = {
    generate: generate, MIME: PPTX_MIME,
    topSpBlocks: topSpBlocks, offOf: offOf, extOf: extOf, runTexts: runTexts,
    replaceRuns: replaceRuns, patchBox: patchBox, findBox: findBox, parseGiven: parseGiven,
    fillSlide: fillSlide, removeBadge: removeBadge, addSigPic: addSigPic,
    parseSigMedia: parseSigMedia, imgAspect: imgAspect, geomOf: geomOf, shiftBoxOff: shiftBoxOff };
})(typeof window !== "undefined" ? window : this);
