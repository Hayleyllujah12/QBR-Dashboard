/* =============================================================================
 * xlsx-surgical.js — format-safe save for .xlsx / .xlsm  (PROTOTYPE, develop)
 * -----------------------------------------------------------------------------
 * SheetJS CE rewrites the whole workbook on save and drops what it doesn't model
 * (styles, conditional formatting, data validation, tables, charts, pivots,
 * Power Query, comments…). This module instead:
 *   1. diffs the patched SheetJS workbook against a fresh parse of the file;
 *   2. writes ONLY the changed cells into the sheet XML (keeps each cell's s=);
 *   3. copies every other zip entry RAW (compressed bytes untouched).
 * Zero dependencies: native DecompressionStream/CompressionStream('deflate-raw')
 * (Chrome/Edge 103+, Node 20.12+), own zip reader/writer + CRC-32.
 * Anything it can't handle safely returns {ok:false, reason} — callers must not
 * write the file in that case.
 * ========================================================================== */
(function (root) {
  "use strict";
  const SURGICAL_VERSION = "0.2.0"; // 0.2.0: hyperlink writes + &amp;amp; repair
  const TD = new TextDecoder("utf-8"), TE = new TextEncoder();

  /* ------------------------------ CRC-32 --------------------------------- */
  const CRC_T = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c; } return t; })();
  function crc32(u8) { let c = -1; for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; }

  /* --------------------------- deflate helpers --------------------------- */
  async function streamAll(u8, ts) {
    const out = new Response(new Blob([u8]).stream().pipeThrough(ts));
    return new Uint8Array(await out.arrayBuffer());
  }
  async function inflateRaw(u8) {
    if (typeof DecompressionStream === "function") return streamAll(u8, new DecompressionStream("deflate-raw"));
    if (root.XLSX && XLSX.CFB && XLSX.CFB.utils._inflateRaw) return XLSX.CFB.utils._inflateRaw(u8, 0);
    throw new Error("no inflate available");
  }
  async function deflateRaw(u8) {
    if (typeof CompressionStream === "function") return { method: 8, data: await streamAll(u8, new CompressionStream("deflate-raw")) };
    return { method: 0, data: u8 }; // stored: valid zip, just bigger
  }

  /* ------------------------------ zip read ------------------------------- */
  function zipRead(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("not a zip (no end record)");
    if (eocd >= 20 && dv.getUint32(eocd - 20, true) === 0x07064b50) throw new Error("zip64 not supported");
    const n = dv.getUint16(eocd + 10, true), cdSize = dv.getUint32(eocd + 12, true), cdOff = dv.getUint32(eocd + 16, true);
    if (n === 0xFFFF || cdOff === 0xFFFFFFFF) throw new Error("zip64 not supported");
    const entries = []; let p = cdOff;
    for (let i = 0; i < n; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("bad central directory");
      const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const e = {
        flags: dv.getUint16(p + 8, true), method: dv.getUint16(p + 10, true),
        crc: dv.getUint32(p + 16, true), csize: dv.getUint32(p + 20, true), usize: dv.getUint32(p + 24, true),
        lho: dv.getUint32(p + 42, true), name: TD.decode(u8.subarray(p + 46, p + 46 + nl)),
        cd: u8.subarray(p, p + 46 + nl + xl + cl),
      };
      if (e.csize === 0xFFFFFFFF || e.usize === 0xFFFFFFFF || e.lho === 0xFFFFFFFF) throw new Error("zip64 not supported");
      if (e.flags & 1) throw new Error("encrypted zip entry");
      entries.push(e); p += 46 + nl + xl + cl;
    }
    // raw span of each entry = local header .. next entry's local header (covers data descriptors)
    const offs = entries.map(e => e.lho).concat([cdOff]).sort((a, b) => a - b);
    entries.forEach(e => {
      const next = offs.find(o => o > e.lho);
      e.span = u8.subarray(e.lho, next);
      const lnl = dv.getUint16(e.lho + 26, true), lxl = dv.getUint16(e.lho + 28, true);
      e.dataStart = e.lho + 30 + lnl + lxl;
    });
    const commentLen = dv.getUint16(eocd + 20, true);
    return { u8, entries, comment: u8.subarray(eocd + 22, eocd + 22 + commentLen), byName: new Map(entries.map(e => [e.name, e])) };
  }
  async function zipText(z, name) {
    const e = z.byName.get(name); if (!e) return null;
    const raw = z.u8.subarray(e.dataStart, e.dataStart + e.csize);
    const data = e.method === 0 ? raw : e.method === 8 ? await inflateRaw(raw) : null;
    if (!data) throw new Error("unsupported compression in " + name);
    return TD.decode(data);
  }

  /* ------------------------------ zip write ------------------------------ */
  function dosTime(d) {
    return { time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
             date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate() };
  }
  /* changes: Map name -> string|Uint8Array (replace/add) or null (delete). */
  async function zipWrite(z, changes) {
    const parts = [], cds = []; let off = 0;
    const now = dosTime(new Date());
    const push = b => { parts.push(b); off += b.length; };
    async function writeNew(name, content, cdTemplate) {
      const raw = typeof content === "string" ? TE.encode(content) : content;
      const crc = crc32(raw), { method, data } = await deflateRaw(raw), nb = TE.encode(name);
      const lh = new Uint8Array(30 + nb.length), lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true);
      lv.setUint16(8, method, true); lv.setUint16(10, now.time, true); lv.setUint16(12, now.date, true);
      lv.setUint32(14, crc, true); lv.setUint32(18, data.length, true); lv.setUint32(22, raw.length, true);
      lv.setUint16(26, nb.length, true); lh.set(nb, 30);
      const cd = new Uint8Array(46 + nb.length), cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, method, true); cv.setUint16(12, now.time, true); cv.setUint16(14, now.date, true);
      cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, raw.length, true);
      cv.setUint16(28, nb.length, true);
      if (cdTemplate) cv.setUint32(38, new DataView(cdTemplate.buffer, cdTemplate.byteOffset).getUint32(38, true), true); // keep ext attrs
      cv.setUint32(42, off, true); cd.set(nb, 46);
      push(lh); push(data); cds.push(cd);
    }
    for (const e of z.entries) {
      if (changes.has(e.name)) {
        const c = changes.get(e.name);
        if (c != null) await writeNew(e.name, c, e.cd);
        continue; // null = delete
      }
      const cd = new Uint8Array(e.cd); new DataView(cd.buffer).setUint32(42, off, true);
      push(e.span); cds.push(cd); // RAW copy: compressed bytes untouched
    }
    for (const [name, c] of changes) if (c != null && !z.byName.has(name)) await writeNew(name, c, null);
    const cdOff = off; cds.forEach(push);
    const end = new Uint8Array(22 + z.comment.length), ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, cds.length, true); ev.setUint16(10, cds.length, true);
    ev.setUint32(12, off - cdOff, true); ev.setUint32(16, cdOff, true); ev.setUint16(20, z.comment.length, true); end.set(z.comment, 22);
    push(end);
    const out = new Uint8Array(off); let q = 0; parts.forEach(b => { out.set(b, q); q += b.length; });
    return out;
  }

  /* ------------------------------ XML utils ------------------------------ */
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const attr = (tag, a) => { const m = new RegExp("\\s" + a + "=\"([^\"]*)\"").exec(tag); return m ? m[1] : null; };
  const setAttr = (tag, a, v) => new RegExp("\\s" + a + "=\"").test(tag)
    ? tag.replace(new RegExp("(\\s" + a + "=\")[^\"]*\""), "$1" + v + "\"")
    : tag.replace(/^<([\w:]+)/, "<$1 " + a + "=\"" + v + "\"");
  const dropAttr = (tag, a) => tag.replace(new RegExp("\\s" + a + "=\"[^\"]*\""), "");
  function colIdx(ref) { let c = 0; const m = /^([A-Z]+)/.exec(ref); for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
  function colName(c) { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
  const addr = (r, c) => colName(c) + (r + 1);
  function decodeRange(ref) {
    const [a, b] = String(ref).split(":"); const p = x => ({ r: +/\d+/.exec(x)[0] - 1, c: colIdx(x) });
    return { s: p(a), e: p(b || a) };
  }
  const encodeRange = g => addr(g.s.r, g.s.c) + ":" + addr(g.e.r, g.e.c);
  function relsOf(path) { const i = path.lastIndexOf("/"); return path.slice(0, i) + "/_rels/" + path.slice(i + 1) + ".rels"; }
  function resolveTarget(base, t) {
    if (t[0] === "/") return t.slice(1);
    const parts = base.split("/"); parts.pop();
    t.split("/").forEach(s => { if (s === "..") parts.pop(); else if (s !== ".") parts.push(s); });
    return parts.join("/");
  }
  function parseRels(xml) {
    const out = []; (xml || "").replace(/<Relationship\b[^>]*>/g, t => { out.push({ id: attr(t, "Id"), type: attr(t, "Type") || "", target: attr(t, "Target") || "", mode: attr(t, "TargetMode") }); });
    return out;
  }

  /* ------------------------------ values -------------------------------- */
  function excelSerial(d, date1904) {
    const ms = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
    let v = (ms - Date.UTC(1899, 11, 30)) / 864e5;
    if (date1904) v -= 1462;
    return v;
  }
  const isDateCell = c => c && (c.t === "d" || c.v instanceof Date);
  function sameCell(a, b) {
    if (!a && !b) return true;
    const blank = c => !c || ((c.v == null || c.v === "") && !c.f);
    if (blank(a) && blank(b)) return true;
    if (!a || !b) return false;
    if ((a.f || "") !== (b.f || "")) return false;
    if (a.f) return true; // formula unchanged → cached value irrelevant
    const av = a.v instanceof Date ? a.v.getTime() : a.v, bv = b.v instanceof Date ? b.v.getTime() : b.v;
    return av === bv && (a.t === b.t || (a.t === "d") === (b.t === "d"));
  }
  function diffSheet(base, work) {
    const keys = new Set(); [base || {}, work || {}].forEach(ws => Object.keys(ws).forEach(k => { if (k[0] !== "!") keys.add(k); }));
    const out = [];
    keys.forEach(k => {
      const a = base && base[k], b = work && work[k];
      if (!sameCell(a, b)) { const r = decodeRange(k).s; out.push({ r: r.r, c: r.c, cell: b || null }); }
    });
    return out;
  }

  /* --------------------------- styles (dates) ---------------------------- */
  function makeStyleHelper(stylesXml) {
    let xml = stylesXml, dirty = false; const cache = {};
    return {
      dateXf(withTime) {
        const fmt = withTime ? 22 : 14;
        if (cache[fmt] != null) return cache[fmt];
        if (!xml) return null;
        const m = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml); if (!m) return null;
        const xfs = m[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
        let i = xfs.findIndex(x => attr(x, "numFmtId") === String(fmt) && (attr(x, "fontId") || "0") === "0" && (attr(x, "fillId") || "0") === "0" && (attr(x, "borderId") || "0") === "0");
        if (i < 0) {
          const nx = `<xf numFmtId="${fmt}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
          const open = /<cellXfs\b[^>]*>/.exec(xml)[0];
          const newOpen = setAttr(open, "count", String(xfs.length + 1));
          xml = xml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, newOpen + m[1] + nx + "</cellXfs>");
          i = xfs.length; dirty = true;
        }
        return (cache[fmt] = i);
      },
      get xml() { return xml; }, get dirty() { return dirty; },
    };
  }

  /* ----------------------------- cell XML ------------------------------- */
  function cellXml(r, c, cell, s, ctx) {
    const ref = addr(r, c), sa = s != null && s !== "" ? ` s="${s}"` : "";
    if (!cell || ((cell.v == null || cell.v === "") && !cell.f)) return `<c r="${ref}"${sa}/>`;
    if (cell.f) {
      const f = `<f>${esc(String(cell.f).replace(/^=/, ""))}</f>`;
      if (cell.v == null || cell.v === "") return `<c r="${ref}"${sa}>${f}</c>`;
      if (typeof cell.v === "number") return `<c r="${ref}"${sa}>${f}<v>${cell.v}</v></c>`;
      if (typeof cell.v === "boolean") return `<c r="${ref}"${sa} t="b">${f}<v>${cell.v ? 1 : 0}</v></c>`;
      return `<c r="${ref}"${sa} t="str">${f}<v>${esc(cell.v)}</v></c>`;
    }
    if (isDateCell(cell)) {
      const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
      const withTime = d.getHours() || d.getMinutes() || d.getSeconds();
      if (!sa || s === "0") { const x = ctx.styles.dateXf(withTime); if (x != null) return `<c r="${ref}" s="${x}"><v>${excelSerial(d, ctx.date1904)}</v></c>`; }
      return `<c r="${ref}"${sa}><v>${excelSerial(d, ctx.date1904)}</v></c>`;
    }
    if (cell.t === "n" || typeof cell.v === "number") return `<c r="${ref}"${sa}><v>${cell.v}</v></c>`;
    if (cell.t === "b" || typeof cell.v === "boolean") return `<c r="${ref}"${sa} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
    const t = String(cell.v), sp = /^\s|\s$|\n/.test(t) ? ' xml:space="preserve"' : "";
    return `<c r="${ref}"${sa} t="inlineStr"><is><t${sp}>${esc(t)}</t></is></c>`;
  }

  /* --------------------------- sheet splicing ---------------------------- */
  const ROW_RE = /<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
  const CELL_RE = /<c(?=[\s>\/])[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
  function parseRow(rowXml) {
    const open = /^<row\b[^>]*?\/?>/.exec(rowXml)[0];
    const cells = rowXml.endsWith("/>") && open === rowXml ? [] : (rowXml.slice(open.length).match(CELL_RE) || []);
    return { open: open.replace(/\/>$/, ">"), cells: cells.map(x => {
      const r = attr(x.slice(0, x.indexOf(">") + 1), "r");
      if (!r) throw new Error("cell without r= attribute");
      return { c: colIdx(r), xml: x };
    }) };
  }
  function cellStyle(xml) { return attr(xml.slice(0, xml.indexOf(">") + 1), "s"); }

  function spliceSheet(xml, edits, ctx) {
    const sd = /<sheetData\s*\/>|<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
    if (!sd) throw new Error("no <sheetData> (unsupported sheet XML)");
    const body = sd[1] || "", rows = [];
    let m; ROW_RE.lastIndex = 0;
    while ((m = ROW_RE.exec(body))) {
      const r = attr(m[0].slice(0, m[0].indexOf(">") + 1), "r");
      if (!r) throw new Error("row without r= attribute");
      rows.push({ r: +r - 1, xml: m[0], start: m.index, end: m.index + m[0].length });
    }
    const byRow = new Map(rows.map(x => [x.r, x]));
    const origLast = rows.length ? rows[rows.length - 1].r : -1;
    // style template per column = nearest original cell above with s=
    const tmplCache = new Map();
    function templateS(r, c) {
      const k = r + ":" + c; if (tmplCache.has(k)) return tmplCache.get(k);
      let s = null;
      for (let i = rows.length - 1; i >= 0 && s == null; i--) {
        if (rows[i].r >= r) continue;
        if (r - rows[i].r > 50) break;
        const pr = parseRow(rows[i].xml).cells.find(x => x.c === c); if (pr) s = cellStyle(pr.xml);
        if (pr) break;
      }
      tmplCache.set(k, s); return s;
    }
    const editsByRow = new Map();
    edits.forEach(e => { if (!editsByRow.has(e.r)) editsByRow.set(e.r, []); editsByRow.get(e.r).push(e); });
    let maxR = -1, maxC = -1, formulas = false;
    const newRowXml = new Map();
    for (const [r, list] of editsByRow) {
      const ex = byRow.get(r);
      const pr = ex ? parseRow(ex.xml) : { open: `<row r="${r + 1}">`, cells: [] };
      list.forEach(e => {
        const i = pr.cells.findIndex(x => x.c === e.c);
        const old = i >= 0 ? pr.cells[i].xml : null;
        if (old && /<f\b[^>]*\bt="shared"[^>]*\bref="/.test(old)) throw new Error(`${addr(r, e.c)} is a shared-formula master`);
        if (old && /<f\b[^>]*\bt="array"/.test(old)) throw new Error(`${addr(r, e.c)} is an array formula`);
        if (e.cell && e.cell.f) formulas = true;
        const s = old ? cellStyle(old) : templateS(r, e.c);
        const nx = { c: e.c, xml: cellXml(r, e.c, e.cell, s, ctx) };
        if (i >= 0) pr.cells[i] = nx; else pr.cells.push(nx);
        maxR = Math.max(maxR, r); maxC = Math.max(maxC, e.c);
      });
      pr.cells.sort((a, b) => a.c - b.c);
      const open = dropAttr(pr.open, "spans");
      newRowXml.set(r, open + pr.cells.map(x => x.xml).join("") + "</row>");
    }
    // rebuild sheetData body: keep untouched text, replace edited rows, insert new rows in order
    let out = "", pos = 0;
    const newOnly = [...newRowXml.keys()].filter(r => !byRow.has(r)).sort((a, b) => a - b);
    let ni = 0;
    rows.forEach(row => {
      while (ni < newOnly.length && newOnly[ni] < row.r) { out += body.slice(pos, row.start); pos = row.start; out += newRowXml.get(newOnly[ni++]); }
      if (newRowXml.has(row.r)) { out += body.slice(pos, row.start) + newRowXml.get(row.r); pos = row.end; }
    });
    out += body.slice(pos);
    while (ni < newOnly.length) out += newRowXml.get(newOnly[ni++]);
    let nxml = xml.slice(0, sd.index) + "<sheetData>" + out + "</sheetData>" + xml.slice(sd.index + sd[0].length);
    if (sd[0].startsWith("<sheetData") && !sd[0].endsWith("/>")) {
      const openTag = /^<sheetData\b[^>]*>/.exec(sd[0])[0];
      nxml = xml.slice(0, sd.index) + openTag + out + "</sheetData>" + xml.slice(sd.index + sd[0].length);
    }
    // <dimension>
    const dm = /<dimension\b[^>]*\/>/.exec(nxml);
    if (dm && maxR >= 0) {
      const g = decodeRange(attr(dm[0], "ref") || "A1");
      const ng = { s: g.s, e: { r: Math.max(g.e.r, maxR), c: Math.max(g.e.c, maxC) } };
      nxml = nxml.replace(dm[0], setAttr(dm[0], "ref", encodeRange(ng)));
    }
    return { xml: nxml, origLast, formulas, maxR, maxC };
  }

  /* ------------------------- tables (ListObjects) ------------------------ */
  function growTable(txml, info, workWs) {
    const open = /<table\b[^>]*>/.exec(txml)[0], ref = attr(open, "ref"); if (!ref) return { xml: txml, changed: false };
    const g = decodeRange(ref); let changed = false, x = txml;
    const totals = +(attr(open, "totalsRowCount") || 0);
    // rows appended directly under the table (no totals row)
    if (!totals && g.e.r === info.origLast && info.maxR > g.e.r) { g.e.r = info.maxR; changed = true; }
    // header written in the column right after the table → add a tableColumn
    const hdr = workWs && workWs[addr(g.s.r, g.e.c + 1)];
    if (hdr && hdr.v != null && hdr.v !== "" && info.newCols.has(g.e.c + 1)) {
      const tc = /<tableColumns\b[^>]*>([\s\S]*?)<\/tableColumns>/.exec(x);
      if (tc) {
        const cols = tc[1].match(/<tableColumn\b/g) || [], ids = (tc[1].match(/\bid="(\d+)"/g) || []).map(s => +/\d+/.exec(s)[0]);
        const nid = Math.max(0, ...ids) + 1, tcOpen = /<tableColumns\b[^>]*>/.exec(tc[0])[0];
        x = x.replace(tc[0], setAttr(tcOpen, "count", String(cols.length + 1)) + tc[1] + `<tableColumn id="${nid}" name="${esc(hdr.v)}"/>` + "</tableColumns>");
        g.e.c += 1; changed = true;
      }
    }
    if (!changed) return { xml: txml, changed };
    const nr = encodeRange(g);
    x = x.replace(open, setAttr(open, "ref", nr));
    x = x.replace(/<autoFilter\b[^>]*?ref="[^"]*"/, s => setAttr(s, "ref", encodeRange({ s: g.s, e: { r: g.e.r, c: g.e.c } })));
    return { xml: x, changed };
  }

  /* ----------------------------- new sheet ------------------------------- */
  function newSheetXml(ws, ctx) {
    const cells = Object.keys(ws).filter(k => k[0] !== "!").map(k => ({ ...decodeRange(k).s, cell: ws[k] }));
    const byR = new Map(); cells.forEach(x => { if (!byR.has(x.r)) byR.set(x.r, []); byR.get(x.r).push(x); });
    const rows = [...byR.keys()].sort((a, b) => a - b).map(r => `<row r="${r + 1}">` +
      byR.get(r).sort((a, b) => a.c - b.c).map(x => cellXml(r, x.c, x.cell, null, ctx)).join("") + "</row>").join("");
    const ref = ws["!ref"] || "A1";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="${ref}"/><sheetData>${rows}</sheetData></worksheet>`;
  }


  /* ----------------------------- hyperlinks ------------------------------ */
  // SheetJS hands back hyperlink targets with XML entities still encoded and
  // re-encodes on write, so files saved by the old engine can carry &amp;amp;.
  // Compare and write targets in a decoded-until-stable form, escaped once.
  const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const HL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
  function decodeStable(t) {
    let prev; t = String(t == null ? "" : t);
    do { prev = t; t = t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&"); } while (t !== prev);
    return t;
  }
  const linkOf = c => (c && c.l && c.l.Target != null && c.l.Target !== "") ? decodeStable(c.l.Target) : "";
  function diffLinks(base, work) {
    const out = [];
    Object.keys(work || {}).forEach(k => {
      if (k[0] === "!") return;
      const a = linkOf(base && base[k]), b = linkOf(work[k]);
      if (a !== b) out.push({ ref: k, target: b });
    });
    return out;
  }
  // Schema order: <hyperlinks> sits after dataValidations and before these.
  const AFTER_HL = ["printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
  function applyLinks(xml, relsXml, linkEdits) {
    const root = /<worksheet\b[^>]*>/.exec(xml)[0];
    const pm = new RegExp('xmlns:([\\w]+)="' + REL_NS.replace(/[./]/g, "\\$&") + '"').exec(root);
    const pfx = pm ? pm[1] : "r", nsDecl = pm ? "" : ` xmlns:r="${REL_NS}"`;
    let rels = relsXml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const relIds = new Set((rels.match(/\bId="([^"]+)"/g) || []).map(x => /"([^"]+)"/.exec(x)[1]));
    let nid = 1; const newId = () => { while (relIds.has("rIdQL" + nid)) nid++; const id = "rIdQL" + nid; relIds.add(id); return id; };
    const hm = /<hyperlinks\b[^>]*>([\s\S]*?)<\/hyperlinks>|<hyperlinks\s*\/>/.exec(xml);
    const items = hm ? ((hm[1] || "").match(/<hyperlink\b[^>]*?\/>|<hyperlink\b[^>]*>[\s\S]*?<\/hyperlink>/g) || []) : [];
    const ridAttr = t => (new RegExp("\\s[\\w]+:id=\"([^\"]*)\"").exec(t) || [])[1];
    linkEdits.forEach(e => {
      const i = items.findIndex(t => attr(t, "ref") === e.ref);
      const old = i >= 0 ? items[i] : null, oldRid = old ? ridAttr(old) : null;
      if (!e.target) { if (i >= 0) items.splice(i, 1); return; }           // link removed
      if (e.target[0] === "#") {                                           // in-workbook location
        const t = `<hyperlink ref="${e.ref}" location="${esc(e.target.slice(1))}"/>`;
        if (i >= 0) items[i] = t; else items.push(t); return;
      }
      const tgt = esc(e.target);
      const relRe = oldRid ? new RegExp(`<Relationship\\b[^>]*\\bId="${oldRid}"[^>]*/>`) : null;
      if (relRe && relRe.test(rels)) {                                     // update the existing relationship
        rels = rels.replace(relRe, t => setAttr(t, "Target", tgt));
        return;
      }
      const id = newId();
      rels = rels.replace(/<\/Relationships>/, `<Relationship Id="${id}" Type="${HL_TYPE}" Target="${tgt}" TargetMode="External"/></Relationships>`);
      const t = `<hyperlink${nsDecl} ref="${e.ref}" ${pfx}:id="${id}"/>`;
      if (i >= 0) items[i] = t; else items.push(t);
    });
    const block = items.length ? "<hyperlinks>" + items.join("") + "</hyperlinks>" : "";
    if (hm) xml = xml.slice(0, hm.index) + block + xml.slice(hm.index + hm[0].length);
    else if (block) {
      let at = -1;
      for (const tag of AFTER_HL) { const m = new RegExp("<" + tag + "\\b").exec(xml); if (m && (at < 0 || m.index < at)) at = m.index; }
      if (at < 0) at = xml.lastIndexOf("</worksheet>");
      xml = xml.slice(0, at) + block + xml.slice(at);
    }
    return { xml, rels };
  }
  // Collapse multi-encoded hyperlink targets (&amp;amp; → &amp;) left by older saves.
  function repairRels(rels) {
    let n = 0;
    const out = rels.replace(/<Relationship\b[^>]*>/g, t => {
      if (!/\/hyperlink"/.test(t)) return t;
      const cur = attr(t, "Target"); if (cur == null || !/&amp;(amp;|lt;|gt;|quot;)/.test(cur)) return t;
      n++; return setAttr(t, "Target", esc(decodeStable(cur)));
    });
    return { rels: out, n };
  }

  /* ------------------------------- main --------------------------------- */
  /* bytes: original file (Uint8Array|ArrayBuffer). base: fresh SheetJS parse of
   * those bytes. work: same parse after patching. Returns
   * {ok, bytes, stats:{cells, sheets, rowsAdded, newSheets, tables}, reason} */
  async function surgicalSave(bytes, base, work, opts) {
    opts = opts || {};
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const stats = { cells: 0, sheets: [], newSheets: [], tables: 0, removedCalcChain: false, engine: "surgical " + SURGICAL_VERSION };
    try {
      const z = zipRead(u8);
      const changes = new Map();
      const wbXml = await zipText(z, "xl/workbook.xml");
      if (!wbXml) throw new Error("xl/workbook.xml missing");
      const wbRelsXml = await zipText(z, "xl/_rels/workbook.xml.rels");
      const rels = parseRels(wbRelsXml);
      const sheetTags = wbXml.match(/<sheet\b[^>]*\/?>/g) || [];
      const ridAttr = (sheetTags[0] && (/\s([\w]+:id)="/.exec(sheetTags[0]) || [])[1]) || "r:id";
      const unXml = s => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
      const pathOf = {};
      sheetTags.forEach(t => {
        const rel = rels.find(r => r.id === attr(t, ridAttr));
        if (rel) pathOf[unXml(attr(t, "name"))] = resolveTarget("xl/workbook.xml", rel.target);
      });
      const stylesPath = (rels.find(r => /\/styles$/.test(r.type)) || {}).target;
      const stylesFull = stylesPath ? resolveTarget("xl/workbook.xml", stylesPath) : "xl/styles.xml";
      const styles = makeStyleHelper(await zipText(z, stylesFull));
      const ctx = { styles, date1904: !!(base.Workbook && base.Workbook.WBProps && base.Workbook.WBProps.date1904) };
      let anyFormula = false, anyChange = false, needRelsDefault = false;

      // 1. existing sheets
      for (const name of base.SheetNames) {
        const edits = diffSheet(base.Sheets[name], work.Sheets[name]);
        const linkEdits = diffLinks(base.Sheets[name], work.Sheets[name]);
        if (!edits.length && !linkEdits.length) continue;
        const p = pathOf[name]; if (!p) throw new Error(`sheet "${name}" not found in the file`);
        if (/\.bin$/.test(p)) throw new Error("binary .xlsb sheets not supported");
        const xml = await zipText(z, p);
        const before = parseRowsLight(xml);
        const res = edits.length ? spliceSheet(xml, edits, ctx) : { xml, origLast: -1, maxR: -1, maxC: -1, formulas: false };
        res.newCols = new Set(edits.filter(e => !before.cols.has(e.c)).map(e => e.c));
        if (linkEdits.length) {
          const rp = relsOf(p), lk = applyLinks(res.xml, changes.get(rp) || await zipText(z, rp), linkEdits);
          res.xml = lk.xml; changes.set(rp, lk.rels); stats.links = (stats.links || 0) + linkEdits.length;
          if (!z.byName.has(rp)) needRelsDefault = true;
        }
        changes.set(p, res.xml);
        stats.cells += edits.length; stats.sheets.push(name + " (" + edits.length + ")");
        anyFormula = anyFormula || res.formulas; anyChange = true;
        // tables on this sheet
        const srels = parseRels(await zipText(z, relsOf(p)));
        for (const tr of srels.filter(r => /\/table$/.test(r.type))) {
          const tp = resolveTarget(p, tr.target), txml = await zipText(z, tp); if (!txml) continue;
          const g = growTable(txml, res, work.Sheets[name]);
          if (g.changed) { changes.set(tp, g.xml); stats.tables++; }
        }
      }
      // 2. new sheets
      const added = work.SheetNames.filter(n => !base.SheetNames.includes(n));
      if (added.length) {
        let wx = changes.get("xl/workbook.xml") || wbXml, rx = wbRelsXml, ct = await zipText(z, "[Content_Types].xml");
        const ids = (wx.match(/\bsheetId="(\d+)"/g) || []).map(s => +/\d+/.exec(s)[0]);
        const rids = rels.map(r => +((/\d+$/.exec(r.id) || [0])[0]));
        let sid = Math.max(0, ...ids), rid = Math.max(0, ...rids), n = 1;
        for (const name of added) {
          while (z.byName.has(`xl/worksheets/sheet${n}.xml`) || changes.has(`xl/worksheets/sheet${n}.xml`)) n++;
          const path = `xl/worksheets/sheet${n}.xml`, id = "rIdS" + (++rid);
          changes.set(path, newSheetXml(work.Sheets[name], ctx));
          const pfx = ridAttr.split(":")[0]; // declare the prefix inline: some producers only declare it per element
          wx = wx.replace(/<\/sheets>/, `<sheet xmlns:${pfx}="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${esc(name)}" sheetId="${++sid}" ${ridAttr}="${id}"/></sheets>`);
          rx = rx.replace(/<\/Relationships>/, `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/></Relationships>`);
          ct = ct.replace(/<\/Types>/, `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
          stats.newSheets.push(name); stats.cells += Object.keys(work.Sheets[name]).filter(k => k[0] !== "!").length;
        }
        changes.set("xl/workbook.xml", wx); changes.set("xl/_rels/workbook.xml.rels", rx); changes.set("[Content_Types].xml", ct);
        anyChange = true;
      }
      if (!anyChange) return { ok: true, bytes: null, stats, noChanges: true };
      // 2b. repair double-encoded hyperlink targets on every sheet (older saves)
      for (const name of Object.keys(pathOf)) {
        const rp = relsOf(pathOf[name]), cur = changes.get(rp) || await zipText(z, rp);
        if (!cur || !/&amp;amp;/.test(cur)) continue;
        const fx = repairRels(cur);
        if (fx.n) { changes.set(rp, fx.rels); stats.linksRepaired = (stats.linksRepaired || 0) + fx.n; anyChange = true; }
      }
      if (needRelsDefault) {
        const ct = changes.get("[Content_Types].xml") || await zipText(z, "[Content_Types].xml");
        if (!/Extension="rels"/i.test(ct)) changes.set("[Content_Types].xml", ct.replace(/<\/Types>/, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'));
      }
      // 3. styles (date xf added)
      if (styles.dirty) changes.set(stylesFull, styles.xml);
      // 4. recalc on open; drop calcChain when formulas were written
      let wx = changes.get("xl/workbook.xml") || wbXml;
      if (/<calcPr\b[^>]*\/?>/.test(wx)) wx = wx.replace(/<calcPr\b[^>]*?(\/?)>/, (t, sl) => setAttr(t.replace(/\/?>$/, ""), "fullCalcOnLoad", "1") + (sl ? "/>" : ">"));
      else wx = wx.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>');
      changes.set("xl/workbook.xml", wx);
      const cc = rels.find(r => /\/calcChain$/.test(r.type));
      if (anyFormula && cc) {
        const ccp = resolveTarget("xl/workbook.xml", cc.target);
        changes.set(ccp, null);
        const rx = changes.get("xl/_rels/workbook.xml.rels") || wbRelsXml;
        changes.set("xl/_rels/workbook.xml.rels", rx.replace(new RegExp(`<Relationship\\b[^>]*Id="${cc.id}"[^>]*/>`), ""));
        const ct = changes.get("[Content_Types].xml") || await zipText(z, "[Content_Types].xml");
        changes.set("[Content_Types].xml", ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${ccp.replace(/[.]/g, "\\.")}"[^>]*/>`), ""));
        stats.removedCalcChain = true;
      }
      const out = await zipWrite(z, changes);
      stats.changedParts = [...changes.keys()];
      return { ok: true, bytes: out, stats };
    } catch (e) {
      return { ok: false, reason: (e && e.message) || String(e), stats };
    }
  }
  function parseRowsLight(xml) {
    const cols = new Set(); (xml.match(/<c\b[^>]*\br="([A-Z]+)\d+"/g) || []).forEach(t => cols.add(colIdx(/r="([A-Z]+)/.exec(t)[1])));
    return { cols };
  }

  /* Read-back check: re-parse the output and confirm every diffed cell holds the
   * expected value, every original part still exists, and macros are intact. */
  async function verifySurgical(origBytes, outBytes, work) {
    const zo = zipRead(origBytes instanceof Uint8Array ? origBytes : new Uint8Array(origBytes)), zn = zipRead(outBytes);
    for (const e of zo.entries) {
      if (!zn.byName.has(e.name) && !/calcChain\.xml$/.test(e.name)) return "part missing after save: " + e.name;
    }
    const vo = zo.byName.get("xl/vbaProject.bin"), vn = zn.byName.get("xl/vbaProject.bin");
    if (vo && (!vn || vo.crc !== vn.crc)) return "macros (vbaProject.bin) changed";
    const ct0 = await zipText(zo, "[Content_Types].xml"), ct1 = await zipText(zn, "[Content_Types].xml");
    const main = s => (/PartName="\/xl\/workbook\.xml"\s+ContentType="([^"]+)"/.exec(s) || [])[1];
    if (main(ct0) !== main(ct1)) return "workbook content type changed";
    let back; try { back = XLSX.read(outBytes, { type: "array", cellStyles: true }); } /* same options as the diff: keeps value-less formula cells */ catch (e) { return "re-read failed: " + e.message; }
    for (const n of work.SheetNames) {
      const a = work.Sheets[n], b = back.Sheets[n]; if (!b) return "sheet missing after save: " + n;
      for (const k of Object.keys(a)) {
        if (k[0] === "!") continue;
        const x = a[k], y = b[k];
        if (linkOf(x) !== linkOf(y)) return `hyperlink mismatch ${n}!${k}`;
        if (x.f) { if (!y || String(y.f || "") !== String(x.f).replace(/^=/, "")) return `formula mismatch ${n}!${k}`; continue; }
        if (x.v == null || x.v === "" || x.t === "z") continue;
        if (isDateCell(x)) { const want = excelSerial(x.v instanceof Date ? x.v : new Date(x.v), false); if (!y || Math.abs((y.v instanceof Date ? excelSerial(y.v) : y.v) - want) > 1e-6) return `date mismatch ${n}!${k}`; continue; }
        if (!y || String(y.v) !== String(x.v)) return `value mismatch ${n}!${k}: "${x.v}" vs "${y && y.v}"`;
      }
    }
    return null;
  }

  const api = { SURGICAL_VERSION, surgicalSave, verifySurgical, zipRead, zipWrite, zipText, crc32, _spliceSheet: spliceSheet, _cellXml: cellXml, _diffSheet: diffSheet, _applyLinks: applyLinks, _repairRels: repairRels, _decodeStable: decodeStable };
  root.QBR = root.QBR || {};
  root.QBR.xlsxSurgical = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
