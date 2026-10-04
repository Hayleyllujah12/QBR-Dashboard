/* ============================================================================
 * QBR Scan module — barcode/QR + OCR label extraction, fully offline
 * v1.2.0 · 2026-10-04 (fix: barcodes on Windows, OCR on file://, tilted photos)
 *
 * Adapts the standalone Asset Extractor (photo upload → native
 * window.BarcodeDetector → Tesseract.js OCR → serial / product-key / model
 * parsing → editable table). Two deliberate fixes vs the standalone:
 *  - NO bogus "ZXing" script tags (the original pointed at unrelated
 *    packages); barcode detection uses the browser's native BarcodeDetector
 *    only — zero new dependencies for that path.
 *  - NO CDN: Tesseract.js v5.1.1 (Apache-2.0) is vendored under
 *    libs/tesseract/ (tesseract.min.js, worker.min.js, the four
 *    tesseract-core wasm variants, fast eng.traineddata + .gz).
 *
 * Loads AFTER js/inventory.js. Everything hangs off the shared QBR
 * namespace. Conventions mirror the existing engine: esc() before innerHTML,
 * "—"/"No data" never 0, pure parsers are Node-testable (guard at bottom).
 *
 * Two entry points (per user decision):
 *  (A) dedicated "Scan" page (dash-scan) under the Inventory nav;
 *  (B) scan buttons inside the Intake ("Scan box label") and Deployment
 *      ("Scan unit") forms, which route through the Scan page in target mode.
 * Post-scan routing: serial normalized via QBR.invSerialKey(); if the serial
 * already exists → toast warning first, then jump to its Asset 360 page;
 * otherwise "Add to intake" prefills the intake form, "Tag to client" opens
 * the tag dialog (Client*, Date*, Purchase location, SQ number, DR #,
 * Assigned to).
 * ==========================================================================*/
var QBR = window.QBR = window.QBR || {};

QBR.SCAN_VERSION = "1.2.0";
/* v1.2.0 fixes (2026-10-04):
 *  1. Barcodes on Windows/Linux: Chrome/Edge only ship window.BarcodeDetector on macOS, ChromeOS and
 *     Android, so the barcode path was dead on Windows PCs. ZXing (vendored, Apache-2.0, pure JS —
 *     works from file://) now decodes when the native detector is missing or finds nothing, scanning
 *     the label in overlapping horizontal bands so every stacked 1D barcode is found, not just one.
 *  2. OCR on file://: Tesseract's worker set-up hung forever (it fetches its core + language files,
 *     which browsers block on file://), so scans never finished and the "OCR unavailable" fallback
 *     never fired. On file:// the engine now loads from libs/tesseract/offline/*.js (plain scripts)
 *     and starts from a blob URL; every start is time-boxed so it can no longer hang.
 *  3. Tilted photos: OCR now deskews (rotateAuto), upscales small images, and retries in sparse-text
 *     mode when no serial was read.
 *  4. Serial parser: a bare "S" no longer counts as a serial label (it grabbed "5VD00…" out of the
 *     MTM "21M4S5VD00"); only SN / S/N / Serial [Number|No|#] labels qualify.
 *  5. Photos added while a scan is running are queued and scanned (they used to stay "Pending").
 *     The detection note shows the engines actually in use and updates after each scan. */

/* ============================ pure parsers =============================== */
/* Adapted from the standalone extractor. Pure over strings — Node-testable. */

function scanParseOCRText(text) {
  const result = { serial: "", productKey: "", model: "" };
  if (!text) return result;

  const pkPatterns = [
    /\b([A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5})\b/gi,
    /\b(\d{5}-\d{5}-\d{4})\b/g,
    /\b(\d{5}-\d{3}-\d{7}-\d{5})\b/g,
  ];
  for (const rx of pkPatterns) {
    const m = text.match(rx);
    if (m) { result.productKey = m[0]; break; }
  }

  /* Serial: needs an explicit label — SN, S/N, (S) SN, Serial, Serial Number/No/# — that is not glued
   * to a preceding letter/digit, so the "S" inside an MTM like "21M4S5VD00" can't start a match.
   * All labelled hits are collected; one with the unit-serial shape wins, else the first. */
  const snRx = /(?:^|[^A-Z0-9])(?:S\s*\/\s*N|SN|SERIAL(?:\s*(?:NUMBER|NO\.?|#))?)\s*[:.=#]?\s*([A-Z0-9][A-Z0-9\-]{5,25})/gi;
  const snHits = [];
  let sm;
  while ((sm = snRx.exec(text)) !== null) {
    const v = sm[1].trim();
    if (!/^(NUMBER|SERIAL)$/i.test(v)) snHits.push(v);
  }
  result.serial = snHits.find(scanLooksLikeSerial) || snHits[0] || "";

  const brands = [
    "ThinkPad", "ThinkCentre", "ThinkStation", "IdeaPad", "Yoga", "Legion",
    "Latitude", "Precision", "OptiPlex", "EliteBook", "ProBook", "ZBook",
    "Inspiron", "XPS", "Vostro", "Surface", "MacBook", "ProLiant",
  ];
  for (const b of brands) {
    const rx = new RegExp(b + "\\s*[A-Z0-9][A-Z0-9\\- ]{1,25}", "gi");
    const m = text.match(rx);
    if (m) { result.model = m[0].trim(); break; }
  }
  /* Fallback: OCR often mangles the brand word ("ThinkPad" -> "Milian") but
   * reads the model number cleanly. Capture the model-number fragment so the
   * row still gets a useful, honest guess the user can refine. */
  if (!result.model) {
    const m = text.match(/\b([A-Z]{1,4}\d{2,4}[A-Z]?\s*Gen\s*\d)/i);
    if (m) result.model = m[1].replace(/\s+/g, " ").trim();
  }
  return result;
}

/* Classify a raw barcode value: product "key", "serial" candidate, or "other"
 * (UPC/EAN digit strings, UUIDs, MAC addresses — never the unit serial). */
function scanClassifyBarcode(bc) {
  const v = String(bc == null ? "" : bc).trim();
  if (/^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$/.test(v) ||
      /^\d{5}-\d{5}-\d{4}$/.test(v)) return "key";
  if (/^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$/i.test(v)) return "other";
  if (/^\d{8}$/.test(v) || /^\d{12,14}$/.test(v)) return "other"; // EAN-8 / UPC-A / EAN-13
  if (/^[0-9A-F]{12}$/i.test(v) && /[A-F]/i.test(v) && /\d/.test(v)) return "other"; // MAC-ish
  return "serial";
}

/* Lenovo (1S) labels concatenate MTM + serial, e.g. "21M4S5VD00PF62SDPW".
 * Split off the trailing 8-char serial so it can win on its own merits. */
function scanSplitMtmSn(v) {
  const m = /^([A-Z0-9]{10,12})([A-Z0-9]{8})$/i.exec(String(v == null ? "" : v).trim());
  return m ? m[2].toUpperCase() : null;
}

/* Lenovo unit serial shape: 7-10 alphanumerics mixing letters and digits. */
function scanLooksLikeSerial(v) {
  return /^(?=.*[A-Z])(?=.*\d)[A-Z0-9]{7,10}$/.test(String(v || ""));
}

/* Pick the serial by consensus across barcode votes and OCR text.
 * Robust on labels with many barcodes, where "first barcode wins" could grab
 * a battery code, MTM, or UPC instead of the unit serial. */
function scanPickSerial(cands, ocrText) {
  const ocr = String(ocrText || "").toUpperCase();
  const votes = new Map(); // value -> {n, first}
  (cands || []).forEach((raw, idx) => {
    const v = String(raw == null ? "" : raw).toUpperCase().trim();
    if (!v) return;
    if (!votes.has(v)) votes.set(v, { n: 0, first: idx });
    votes.get(v).n++;
  });
  let best = "", bestScore = -1, bestFirst = Infinity;
  votes.forEach((e, v) => {
    let s = e.n * 3;
    if (scanLooksLikeSerial(v)) s += 3;
    if (v.length >= 7 && ocr.indexOf(v) !== -1) s += 4;
    if (s > bestScore || (s === bestScore && e.first < bestFirst)) {
      best = v; bestScore = s; bestFirst = e.first;
    }
  });
  return best;
}

/* Merge barcode + OCR findings into one extraction, preferring consensus. */
function scanMergeFindings(barcodes, parsed, ocrText) {
  let bcKey = "";
  const cands = [];
  (barcodes || []).forEach(bc => {
    const v = String(bc == null ? "" : bc).trim();
    if (!v) return;
    const kind = scanClassifyBarcode(v);
    if (kind === "key") { if (!bcKey) bcKey = v; }
    else if (kind === "serial") {
      const tail = scanSplitMtmSn(v);
      cands.push(tail || v); // MTM+SN concatenation votes for its tail
    }
  });
  if (parsed && parsed.serial) cands.push(parsed.serial);
  let serial = scanPickSerial(cands, ocrText) || "";
  let productKey = bcKey || (parsed && parsed.productKey) || "";
  const model = (parsed && parsed.model) || "";
  serial = serial.replace(/^S\/?N\s*[:=]?\s*/i, "").trim();
  return { serial, productKey, model };
}

/* ====================== browser-only detection =========================== */
/* Needs DOM (Image, canvas) and possibly Tesseract — never called under Node. */

QBR._scanOCR = { worker: null, failed: false, starting: null, mode: "", reason: "" };
QBR._scanEngines = { barcode: "", ocr: "" }; // what actually ran on the last scan (for the page note)

/* Classic <script> injection works from file:// (fetch/XHR do not); loads are cached. */
const _scanScripts = {};
function scanLoadScript(src) {
  if (!_scanScripts[src]) {
    _scanScripts[src] = new Promise((res, rej) => {
      const el = document.createElement("script");
      el.src = src; el.async = true;
      el.onload = () => res();
      el.onerror = () => { delete _scanScripts[src]; rej(new Error("could not load " + src)); };
      document.head.appendChild(el);
    });
  }
  return _scanScripts[src];
}
function scanWithTimeout(p, ms, label) {
  return Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(label + " timed out")), ms))]);
}
function scanLoadImage(dataUrl) {
  return new Promise((res, rej) => { const img = new Image(); img.onload = () => res(img); img.onerror = () => rej(new Error("image could not be read")); img.src = dataUrl; });
}
function scanCanvasFrom(src, w, h) {
  const c = document.createElement("canvas");
  c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(h));
  const ctx = c.getContext("2d", { willReadFrequently: true });
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(src, 0, 0, c.width, c.height);
  return c;
}

const SCAN_BD_FORMATS = ["code_128", "code_39", "code_93", "codabar", "ean_13", "ean_8", "itf",
  "qr_code", "upc_a", "upc_e", "pdf417", "aztec", "data_matrix"];

/* Native detector (macOS / ChromeOS / Android Chrome & Edge). Full image, then a hard-threshold pass. */
async function scanNativeDetect(canvas) {
  const out = [];
  if (typeof window === "undefined" || !window.BarcodeDetector) return out;
  const det = new BarcodeDetector({ formats: SCAN_BD_FORMATS });
  const run = async c => { const ds = await det.detect(c); ds.forEach(d => { if (d.rawValue && !out.includes(d.rawValue)) out.push(d.rawValue); }); };
  try { await run(canvas); } catch (e) { /* unsupported format on this platform etc. */ }
  if (!out.length) {
    try {
      const c = scanCanvasFrom(canvas, canvas.width, canvas.height), ctx = c.getContext("2d");
      const id = ctx.getImageData(0, 0, c.width, c.height), px = id.data;
      for (let i = 0; i < px.length; i += 4) {
        const v = (px[i] * 0.299 + px[i + 1] * 0.587 + px[i + 2] * 0.114) > 128 ? 255 : 0;
        px[i] = px[i + 1] = px[i + 2] = v;
      }
      ctx.putImageData(id, 0, 0);
      await run(c);
    } catch (e) { /* best effort */ }
  }
  return out;
}

/* ZXing (vendored libs/zxing/zxing.min.js, lazy-loaded). ZXing returns one code per decode, and box
 * labels stack ~5-10 1D barcodes, so after a TRY_HARDER pass on the whole image we decode overlapping
 * horizontal bands (each band usually holds one barcode row). Typical cost: 0.4-1.5 s per photo. */
async function scanZxingDetect(canvas) {
  await scanLoadScript("libs/zxing/zxing.min.js");
  const Z = window.ZXing;
  if (!Z || !Z.MultiFormatReader) throw new Error("ZXing not available");
  const F = Z.BarcodeFormat;
  const linear = [F.CODE_128, F.CODE_39, F.CODE_93, F.CODABAR, F.EAN_13, F.EAN_8, F.ITF, F.UPC_A, F.UPC_E];
  const reader = (hard, formats) => {
    const h = new Map();
    h.set(Z.DecodeHintType.POSSIBLE_FORMATS, formats);
    if (hard) h.set(Z.DecodeHintType.TRY_HARDER, true);
    const r = new Z.MultiFormatReader(); r.setHints(h); return r;
  };
  const rFull = reader(true, linear.concat([F.QR_CODE, F.DATA_MATRIX, F.PDF_417, F.AZTEC]));
  const rBand = reader(false, linear);
  const sc = Math.min(1, 2000 / Math.max(canvas.width, canvas.height)); // detail matters; cap the cost
  const full = sc < 1 ? scanCanvasFrom(canvas, canvas.width * sc, canvas.height * sc) : canvas;
  const out = [];
  const decode = (rd, c, bins) => {
    for (const Bin of bins) {
      try {
        const r = rd.decodeWithState(new Z.BinaryBitmap(new Bin(new Z.HTMLCanvasElementLuminanceSource(c))));
        const v = r && r.getText();
        if (v && !out.includes(v)) out.push(v);
        return;
      } catch (e) { /* NotFound / checksum — try the next binarizer or band */ }
    }
  };
  decode(rFull, full, [Z.HybridBinarizer]);
  const W = full.width, H = full.height;
  const band = document.createElement("canvas"), bctx = band.getContext("2d", { willReadFrequently: true });
  [7, 12].forEach(n => {
    const bh = Math.ceil(H / n);
    for (let y = 0; y < H; y += Math.ceil(bh / 2)) {
      const h = Math.min(bh, H - y);
      if (h < 16) continue;
      band.width = W; band.height = h;
      bctx.drawImage(full, 0, y, W, h, 0, 0, W, h);
      decode(rBand, band, [Z.GlobalHistogramBinarizer, Z.HybridBinarizer]);
    }
  });
  return out;
}

async function scanDetectBarcodes(dataUrl) {
  const img = await scanLoadImage(dataUrl);
  const canvas = scanCanvasFrom(img, img.naturalWidth || img.width, img.naturalHeight || img.height);
  let codes = await scanNativeDetect(canvas), engine = codes.length ? "native" : "";
  if (!codes.length) {
    try { codes = await scanZxingDetect(canvas); engine = "zxing"; }
    catch (e) { console.warn("[QBR scan] ZXing barcode decode unavailable:", e && e.message); engine = engine || "none"; }
  }
  QBR._scanEngines.barcode = (typeof window !== "undefined" && window.BarcodeDetector) ? "native+zxing" : (engine === "none" ? "none" : "zxing");
  return codes;
}

/* ---- OCR engine start-up ----
 * http(s): the standard Tesseract.js loader (worker/core/lang fetched from libs/tesseract/).
 * file://: browsers refuse worker scripts, importScripts and fetch for file:// URLs, so the engine is
 *   assembled from libs/tesseract/offline/*.js (plain scripts): core + worker source go into ONE blob
 *   (a blob worker can't importScripts a second blob on an opaque origin), the English model is
 *   handed over in memory. Each start is time-boxed; a failure leaves barcode scanning working. */
function scanWasmSimd() {
  try { return WebAssembly.validate(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0, 1, 5, 1, 96, 0, 1, 123, 3, 2, 1, 0, 10, 10, 1, 8, 0, 65, 0, 253, 15, 253, 98, 11])); }
  catch (e) { return false; }
}
async function scanStartOCRHosted() {
  return Tesseract.createWorker("eng", 1, {
    workerPath: "libs/tesseract/worker.min.js",
    corePath: "libs/tesseract",
    langPath: "libs/tesseract",
    cacheMethod: "none",
    logger: function () {},
  });
}
async function scanStartOCROffline() {
  const core = scanWasmSimd() ? "core-simd-lstm" : "core-lstm";
  await scanLoadScript("libs/tesseract/offline/worker.js");
  await scanLoadScript("libs/tesseract/offline/" + core + ".js");
  await scanLoadScript("libs/tesseract/offline/eng.js");
  const S = window.__QBR_TESS || {};
  if (!S.worker || !S[core] || !S.eng) throw new Error("offline OCR files incomplete");
  const blobUrl = URL.createObjectURL(new Blob([S[core], "\n;\n", S.worker], { type: "application/javascript" }));
  const bin = atob(S.eng), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const w = await Tesseract.createWorker([{ code: "eng", data: bytes }], 1, {
    workerBlobURL: false, workerPath: blobUrl, corePath: "inline.js", // core is already inside the blob
    cacheMethod: "none", logger: function () {},
  });
  // the ~6 MB of source strings are now inside the worker — release the page copies
  delete S.worker; delete S["core-simd-lstm"]; delete S["core-lstm"]; delete S.eng;
  return w;
}
async function scanGetOCRWorker() {
  const st = QBR._scanOCR;
  if (st.worker) return st.worker;
  if (st.failed || typeof Tesseract === "undefined") { if (!st.reason) st.reason = "engine not loaded"; st.failed = true; return null; }
  if (!st.starting) {
    st.starting = (async () => {
      const isFile = typeof location !== "undefined" && location.protocol === "file:";
      const tries = isFile ? [["offline", scanStartOCROffline]] : [["hosted", scanStartOCRHosted], ["offline", scanStartOCROffline]];
      for (const [mode, start] of tries) {
        try {
          st.worker = await scanWithTimeout(start(), 60000, "OCR start");
          st.mode = mode;
          return st.worker;
        } catch (e) {
          st.reason = (e && e.message) || "could not start";
          console.warn("[QBR scan] OCR " + mode + " start failed:", st.reason);
        }
      }
      st.failed = true;
      return null;
    })();
  }
  return st.starting;
}

/* OCR one label photo: upscale small images, deskew (rotateAuto), and if no labelled serial came back,
 * retry in sparse-text mode. Both passes' text is returned so the serial vote sees everything. */
async function scanRunOCR(dataUrl) {
  const w = await scanGetOCRWorker();
  if (!w) { QBR._scanEngines.ocr = "unavailable"; return { text: "", ocrOk: false }; }
  try {
    const img = await scanLoadImage(dataUrl);
    const iw = img.naturalWidth || img.width, ih = img.naturalHeight || img.height;
    const up = Math.max(iw, ih) < 1200 ? 1600 / Math.max(iw, ih) : 1;
    const src = up > 1 ? scanCanvasFrom(img, iw * up, ih * up) : img;
    const pass = async psm => {
      await w.setParameters({ tessedit_pageseg_mode: psm });
      const r = await scanWithTimeout(w.recognize(src, { rotateAuto: true }), 150000, "OCR");
      return (r && r.data && r.data.text) || "";
    };
    let text = await pass("3");
    if (!scanParseOCRText(text).serial) text += "\n" + await pass("11");
    QBR._scanEngines.ocr = "ok";
    return { text, ocrOk: true };
  } catch (e) {
    console.warn("[QBR scan] OCR failed:", e && e.message);
    QBR._scanEngines.ocr = "error";
    return { text: "", ocrOk: false };
  } finally {
    try { await w.setParameters({ tessedit_pageseg_mode: "3" }); } catch (e) { /* worker gone */ }
  }
}

/* Full pipeline for one image. Browser-only. */
QBR.scanImage = async function (dataUrl) {
  const barcodes = await scanDetectBarcodes(dataUrl);
  const ocr = await scanRunOCR(dataUrl);
  const merged = scanMergeFindings(barcodes, scanParseOCRText(ocr.text), ocr.text);
  return { barcodes, serial: merged.serial, productKey: merged.productKey,
           model: merged.model, ocrOk: ocr.ocrOk, barcodeOk: barcodes.length > 0 };
};

/* ============================ UI state =================================== */
QBR._scanUI = { rows: [], nextId: 1, processing: false, target: null }; // target: {mode:"intake"|"deploy"}

function scanInv() { return (typeof APP !== "undefined" && APP.model && APP.model.inventory) || null; }
function scanKey(s) { return (typeof QBR.invSerialKey === "function" ? QBR.invSerialKey(s) : String(s || "").toUpperCase().trim()); }

/* Dashboard-styled toast (app.js has no global toast; keep it local). */
let _scanToastTimer = null;
function scanToast(msg, ms) {
  let el = document.getElementById("scan-toast");
  if (!el) {
    el = document.createElement("div");
    el.id = "scan-toast";
    el.setAttribute("role", "status");
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(_scanToastTimer);
  _scanToastTimer = setTimeout(() => el.classList.remove("show"), ms || 3200);
}

/* Called by the Intake/Deployment form buttons (inventory.js). */
QBR.scanSetTarget = function (mode) {
  QBR._scanUI.target = { mode };
  if (typeof goToTab === "function") goToTab("dash-scan");
  scanToast(mode === "intake"
    ? "Scan a box label — the first result fills the intake form."
    : "Scan a unit label — the serial fills the deployment form.");
};

/* ============================ SCAN PAGE ================================== */
function scanEngineNote() {
  const hasBD = typeof window !== "undefined" && !!window.BarcodeDetector;
  const st = QBR._scanOCR;
  const bc = hasBD ? "barcode/QR (built-in detector, ZXing backup)" : "barcode/QR (ZXing, offline)";
  const ocr = (st.failed || typeof Tesseract === "undefined")
    ? '<b>OCR unavailable</b>' + (st.reason ? " (" + esc(st.reason) + ")" : "") + " — barcode scanning still works"
    : st.worker ? "label text OCR (on-device" + (st.mode === "offline" ? ", offline engine" : "") + ")"
    : "label text OCR (on-device; starts on the first scan, a few seconds)";
  return "Detection: " + bc + " · " + ocr + ".";
}
function scanUpdateEngineNote() { const n = $("scan-engine-note"); if (n) n.innerHTML = scanEngineNote(); }

function renderScan() {
  const host = $("scan-body");
  if (!host) return;
  const ui = QBR._scanUI;
  const ocrDead = QBR._scanOCR.failed || typeof Tesseract === "undefined";

  const targetBar = ui.target
    ? `<div class="alert alert-info py-2 small mb-2">Scan mode: <b>${ui.target.mode === "intake" ? "Intake — first result fills the intake form" : "Deployment — first result fills the deployment form"}</b>
       <button type="button" class="btn btn-sm btn-outline-secondary ms-2" id="scan-target-cancel">Cancel</button></div>`
    : "";
  const engineNote = `<div class="small text-muted mb-2" id="scan-engine-note">${scanEngineNote()}</div>`;

  host.innerHTML =
    `<div class="card-box"><h6>Scan labels</h6>
      <p class="text-muted small mb-2">Upload photos of box or unit labels. Each image is scanned for barcodes/QR codes${ocrDead ? "" : " and label text (OCR)"}; the serial number, model and product key are extracted into the table below.</p>
      ${targetBar}${engineNote}
      <div class="scan-zone" id="scan-zone" role="button" tabindex="0" aria-label="Upload label photos">
        <input type="file" id="scan-files" multiple accept="image/*" aria-hidden="true" tabindex="-1">
        <div class="scan-zone-icon">📷</div>
        <div><b>Drop label photos here or click to browse</b></div>
        <div class="small text-muted">JPG / PNG / WebP — one label per photo works best</div>
      </div>
      <div class="d-flex flex-wrap gap-2 align-items-center mt-2">
        <button type="button" class="btn btn-sm btn-primary" id="scan-all" ${ui.rows.some(r => r.status !== "done" && r.dataUrl) && !ui.processing ? "" : "disabled"}>Scan all</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="scan-clear" ${ui.rows.length ? "" : "disabled"}>Clear</button>
        <span class="small text-muted" id="scan-count">${ui.rows.length} image${ui.rows.length === 1 ? "" : "s"}</span>
      </div></div>
    <div class="card-box mt-3"><div class="d-flex flex-wrap align-items-center gap-2 mb-2">
        <h6 class="mb-0">Extraction results</h6>
        <button type="button" class="btn btn-sm btn-primary" id="scan-batch-tag" disabled>Tag batch</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="scan-batch-clear" disabled>Clear selection</button>
        <span class="small text-muted" id="scan-batch-count"></span>
      </div>
      <div class="table-responsive"><table class="table table-sm inv-tbl"><thead><tr>
        <th style="width:36px"><input type="checkbox" id="scan-sel-all" title="Select all"></th>
        <th style="width:52px">Photo</th><th>Model</th><th>Serial number</th><th>Product key</th>
        <th>Raw barcodes</th><th style="width:90px">Status</th><th style="width:210px">Actions</th>
      </tr></thead><tbody id="scan-tbody"></tbody></table></div>
      <p class="text-muted small mb-0" id="scan-empty"${ui.rows.length ? " hidden" : ""}>No scans yet — upload a photo to start.</p>
    </div>
    <div id="scan-tag-host"></div>
    <div class="modal-overlay" id="scan-lightbox"><div class="scan-lb-dialog" role="dialog" aria-label="Label photo and manual entry">
      <button class="modal-close" id="scan-lb-close" aria-label="Close">✕</button>
      <div class="scan-lb-photo"><img id="scan-lb-img" alt="Label photo preview"></div>
      <div class="scan-lb-form">
        <h6>Manual entry</h6>
        <label class="form-label small mb-1" for="scan-lb-serial">Serial number</label>
        <input id="scan-lb-serial" class="form-control form-control-sm" style="font-family:monospace" placeholder="Type serial from photo…" autocomplete="off">
        <label class="form-label small mb-1 mt-2" for="scan-lb-model">Model</label>
        <input id="scan-lb-model" class="form-control form-control-sm" placeholder="Model" autocomplete="off">
        <div class="d-flex gap-2 mt-3">
          <button type="button" class="btn btn-sm btn-primary" id="scan-lb-save">Save</button>
          <button type="button" class="btn btn-sm btn-outline-secondary" id="scan-lb-rescan">↻ Rescan</button>
        </div>
        <p class="text-muted small mt-2 mb-0" id="scan-lb-hint"></p>
      </div>
    </div></div>`;

  scanRenderRows();
  scanBind(host);
}

function scanRowStatus(row) {
  const map = { done: ["Done", "status-done"], pending: ["Pending", "status-pending"],
                processing: ["Scanning…", "status-processing"], error: ["Error", "status-error"] };
  const [lbl, cls] = map[row.status] || map.pending;
  const dup = row.dup === "inventory"
    ? ` <span class="badge bg-warning text-dark" title="This serial is already in the inventory database">⚠ In inventory</span>`
    : row.dup === "queue"
    ? ` <span class="badge bg-warning text-dark" title="This serial appears more than once in this scan batch">⚠ Duplicate</span>`
    : "";
  return `<span class="status ${cls}">${lbl}</span>${dup}`;
}

function scanRenderRows() {
  const tb = $("scan-tbody");
  if (!tb) return;
  const ui = QBR._scanUI;
  tb.innerHTML = ui.rows.map(row => {
    const thumb = row.dataUrl
      ? `<img src="${row.dataUrl}" class="scan-thumb" alt="label photo" data-scan-act="preview" data-id="${row.id}">`
      : `<span class="text-muted">—</span>`;
    const checked = row.sel ? " checked" : "";
    return `<tr data-id="${row.id}">
      <td><input type="checkbox" class="scan-sel" data-id="${row.id}"${checked} title="Select for batch tag"></td>
      <td>${thumb}</td>
      <td><input class="scan-input" data-f="model" data-id="${row.id}" value="${esc(row.model)}" placeholder="Model"></td>
      <td><input class="scan-input" data-f="serial" data-id="${row.id}" value="${esc(row.serial)}" placeholder="Serial" style="font-family:monospace"></td>
      <td><input class="scan-input" data-f="productKey" data-id="${row.id}" value="${esc(row.productKey)}" placeholder="—"></td>
      <td class="small text-muted" style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(row.barcodeRaw)}">${esc(row.barcodeRaw) || "—"}</td>
      <td>${scanRowStatus(row)}</td>
      <td><div class="d-flex flex-wrap gap-1">
        <button type="button" class="btn btn-sm btn-outline-primary" data-scan-act="intake" data-id="${row.id}" title="Add to intake">Intake</button>
        <button type="button" class="btn btn-sm btn-outline-primary" data-scan-act="tag" data-id="${row.id}" title="Tag to client">Tag</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" data-scan-act="rescan" data-id="${row.id}" title="Rescan">↻</button>
        ${row.dataUrl ? `<button type="button" class="btn btn-sm btn-outline-secondary" data-scan-act="dl" data-id="${row.id}" title="Download photo">⭳</button>` : ""}
        <button type="button" class="btn btn-sm btn-outline-danger" data-scan-act="del" data-id="${row.id}" title="Delete">✕</button>
      </div></td></tr>`;
  }).join("");
  const sa = $("scan-all"), sc = $("scan-clear"), cn = $("scan-count"), em = $("scan-empty");
  if (sa) sa.disabled = ui.processing || !ui.rows.some(r => r.status !== "done" && r.dataUrl);
  if (sc) sc.disabled = !ui.rows.length;
  if (cn) cn.textContent = ui.rows.length + " image" + (ui.rows.length === 1 ? "" : "s");
  if (em) em.hidden = !!ui.rows.length;
  scanRefreshBatchUI();
}

function scanBind(host) {
  const ui = QBR._scanUI;
  const zone = $("scan-zone"), fi = $("scan-files");
  if (zone && fi) {
    const open = e => { if (e.target !== fi) fi.click(); };
    zone.addEventListener("click", open);
    zone.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fi.click(); } });
    zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragover"); });
    zone.addEventListener("dragleave", e => { e.preventDefault(); zone.classList.remove("dragover"); });
    zone.addEventListener("drop", e => { e.preventDefault(); zone.classList.remove("dragover"); scanAddFiles(e.dataTransfer.files); });
    fi.addEventListener("change", () => { scanAddFiles(fi.files); fi.value = ""; });
  }
  const sa = $("scan-all");
  if (sa) sa.addEventListener("click", () => {
    ui.rows.forEach(r => { if (r.status !== "done" && r.status !== "processing" && r.dataUrl) r.status = "pending"; });
    scanPump();
  });
  const sc = $("scan-clear");
  if (sc) sc.addEventListener("click", () => { ui.rows = []; ui.nextId = 1; renderScan(); });
  const tc = $("scan-target-cancel");
  if (tc) tc.addEventListener("click", () => { ui.target = null; renderScan(); });

  const tb = $("scan-tbody");
  if (tb) {
    tb.addEventListener("input", e => {
      if (!e.target.matches(".scan-input")) return;
      const row = ui.rows.find(r => r.id === +e.target.dataset.id);
      if (row) row[e.target.dataset.f] = e.target.value;
    });
    tb.addEventListener("change", e => {
      if (!e.target.matches(".scan-input")) return;
      if (e.target.dataset.f !== "serial") return;
      const r = ui.rows.find(x => x.id === +e.target.dataset.id);
      if (r) r.dupAck = false;
      scanFlagDuplicates();
      scanRenderRows();
    });
    tb.addEventListener("change", e => {
      if (!e.target.matches(".scan-sel")) return;
      const r = ui.rows.find(x => x.id === +e.target.dataset.id);
      if (r) r.sel = e.target.checked;
      scanRefreshBatchUI();
    });
    tb.addEventListener("click", e => {
      const el = e.target.closest("[data-scan-act]");
      if (!el) return;
      const id = +el.dataset.id, act = el.dataset.scanAct;
      if (act === "del") { ui.rows = ui.rows.filter(r => r.id !== id); scanRenderRows(); }
      else if (act === "rescan") { const r = ui.rows.find(x => x.id === id); if (r && r.status !== "processing") { r.status = "pending"; scanRenderRows(); scanPump(); } }
      else if (act === "dl") scanDownload(id);
      else if (act === "preview") scanPreview(id);
      else if (act === "intake" || act === "tag") {
        const row = ui.rows.find(r => r.id === id);
        if (row) scanRoute(row, act);
      }
    });
  }
  const lb = $("scan-lightbox");
  if (lb) {
    lb.addEventListener("click", e => { if (e.target === lb || e.target.id === "scan-lb-close") lb.classList.remove("open"); });
  }
  if (!QBR._scanBatchBound) {
    QBR._scanBatchBound = true;
    document.addEventListener("click", function (e) {
      if (e.target && e.target.id === "scan-batch-tag") scanBatchTagDialog();
      if (e.target && e.target.id === "scan-batch-clear") {
        QBR._scanUI.rows.forEach(r => { r.sel = false; });
        scanRenderRows();
      }
    });
    document.addEventListener("change", function (e) {
      if (e.target && e.target.id === "scan-sel-all") {
        const on = e.target.checked;
        QBR._scanUI.rows.forEach(r => { r.sel = on; });
        scanRenderRows();
      }
    });
  }
  if (!QBR._scanLbBound) {
    QBR._scanLbBound = true;
    document.addEventListener("click", function (e) {
      if (e.target && e.target.id === "scan-lb-save") scanPreviewSave(true);
      if (e.target && e.target.id === "scan-lb-rescan") {
        var l = $("scan-lightbox"), rid = l && l.dataset.rowId;
        scanPreviewSave(false);
        if (l) l.classList.remove("open");
        var row = QBR._scanUI.rows.find(function (r) { return String(r.id) === String(rid); });
        if (row && row.dataUrl) { row.status = "pending"; row.serial = ""; row.model = ""; scanRenderRows(); scanPump(); }
      }
    });
    /* Enter in the serial box saves; focus starts in the serial box. */
    document.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && e.target && e.target.id === "scan-lb-serial") {
        e.preventDefault(); scanPreviewSave(true);
      }
    });
  }
  if (!QBR._scanEscBound) {
    QBR._scanEscBound = true;
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") { const l = $("scan-lightbox"); if (l) l.classList.remove("open"); }
    });
  }
}

/* Fast fingerprint of a photo's bytes (sampled FNV-1a, two passes) so the
 * exact same file uploaded twice is caught without re-scanning it. */
function scanPhotoHash(dataUrl) {
  const s = String(dataUrl || "");
  const step = Math.max(1, Math.floor(s.length / 4096));
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  for (let i = 0; i < s.length; i += step) {
    h1 = Math.imul(h1 ^ s.charCodeAt(i), 0x01000193) >>> 0;
    h2 = Math.imul(h2 ^ s.charCodeAt(s.length - 1 - i), 0x01000193) >>> 0;
  }
  return h1.toString(36) + "-" + h2.toString(36);
}

/* Flag rows whose serial is already in the inventory database ("inventory")
 * or appears more than once in this scan batch ("queue"). Runs after every
 * scan and after manual serial edits. */
function scanFlagDuplicates() {
  const ui = QBR._scanUI;
  const seen = new Map();
  ui.rows.forEach(r => {
    r.dup = null;
    const sn = (r.serial || "").trim().toUpperCase();
    if (!sn || r.status === "pending" || r.status === "processing") return;
    if (scanFindAsset(sn)) { r.dup = "inventory"; return; }
    if (seen.has(sn)) { r.dup = "queue"; seen.get(sn).dup = "queue"; }
    else seen.set(sn, r);
  });
}

function scanAddFiles(files) {
  const ui = QBR._scanUI;
  const imgs = [...(files || [])].filter(f => f.type && f.type.startsWith("image/"));
  if (!imgs.length) { scanToast("No image files found."); return; }
  let n = 0;
  imgs.forEach(f => {
    const rd = new FileReader();
    rd.onload = () => {
      n++;
      const fp = scanPhotoHash(rd.result);
      const already = ui.rows.find(r => r.fp === fp);
      if (already) {
        scanToast("This photo is already in the list — skipped.");
      } else {
        ui.rows.push({ id: ui.nextId++, file: f, dataUrl: rd.result, fp: fp,
          model: "", serial: "", productKey: "", barcodeRaw: "", status: "pending", dup: null });
      }
      scanRenderRows();
      if (n === imgs.length) {
        const sa = $("scan-all"); if (sa) sa.disabled = false;
        scanPump(); // queue: also picks up photos added while another batch is still scanning
      }
    };
    rd.readAsDataURL(f);
  });
}

/* One scan at a time, in order; keeps going until no row is pending (rows added mid-run included). */
async function scanPump() {
  const ui = QBR._scanUI;
  if (ui.processing || ui._pumping) return;
  ui._pumping = true;
  try {
    let next;
    while ((next = ui.rows.find(r => r.status === "pending" && r.dataUrl))) await scanRowScan(next.id);
  } finally { ui._pumping = false; }
}

async function scanRowScan(id) {
  const ui = QBR._scanUI;
  const row = ui.rows.find(r => r.id === id);
  if (!row || !row.dataUrl || row.status === "processing" || ui.processing) return;
  ui.processing = true; row.status = "processing"; scanRenderRows();
  try {
    const res = await QBR.scanImage(row.dataUrl);
    row.barcodeRaw = res.barcodes.join(" | ");
    // never overwrite a hand-edited field
    if (!row.serial) row.serial = res.serial;
    if (!row.productKey) row.productKey = res.productKey;
    if (!row.model) row.model = res.model;
    row.status = "done";
    row.dupAck = false;
    scanFlagDuplicates();
    const dupMsg = row.dup === "inventory" ? " — already in inventory!" :
                   row.dup === "queue" ? " — duplicate in this batch!" : "";
    scanToast(row.serial ? `Extracted serial ${row.serial}${dupMsg}` : "Scan done — no serial found, edit the row manually.");
  } catch (e) {
    row.status = "error";
    scanToast("Scan failed: " + (e && e.message ? e.message : "unknown error"));
  } finally {
    ui.processing = false; scanRenderRows(); scanUpdateEngineNote();
  }
  // target mode (scan buttons inside Intake/Deployment forms): first good
  // result routes through the standard post-scan routing (exists-check first)
  if (ui.target && row.serial) {
    const t = ui.target; ui.target = null;
    scanRoute(row, t.mode === "deploy" ? "deploy-fill" : "intake");
  } else if (ui.target) {
    renderScan(); // keep target bar visible so the user can try another photo
  }
}

function scanDownload(id) {
  const row = QBR._scanUI.rows.find(r => r.id === id);
  if (!row || !row.dataUrl) return;
  const m = /^data:image\/(\w+)/.exec(row.dataUrl || "");
  const ext = m ? (m[1].toLowerCase() === "jpeg" ? "jpg" : m[1].toLowerCase()) : "jpg";
  const a = document.createElement("a");
  a.href = row.dataUrl;
  a.download = (row.serial || ("label_" + row.id)).replace(/[<>:"/\\|?*]/g, "_") + "." + ext;
  a.click();
}

function scanPreview(id) {
  const row = QBR._scanUI.rows.find(r => r.id === id);
  const lb = $("scan-lightbox");
  if (!row || !row.dataUrl || !lb) return;
  lb.dataset.rowId = id;
  const img = $("scan-lb-img"), se = $("scan-lb-serial"),
        mo = $("scan-lb-model"), hint = $("scan-lb-hint");
  if (img) img.src = row.dataUrl;
  if (se) se.value = row.serial || "";
  if (mo) mo.value = row.model || "";
  if (hint) hint.textContent = row.serial
    ? "Detected — correct it here if wrong."
    : "Nothing detected — read the label and type the serial.";
  lb.classList.add("open");
  setTimeout(function () { var s = $("scan-lb-serial"); if (s) { s.focus(); s.select(); } }, 60);
}

/* Save the manual entry back to the row (serial normalized to uppercase,
 * like extracted values) and refresh the table. */
function scanPreviewSave(close) {
  const lb = $("scan-lightbox");
  const rid = lb && lb.dataset.rowId;
  const row = QBR._scanUI.rows.find(function (r) { return String(r.id) === String(rid); });
  if (row) {
    const se = $("scan-lb-serial"), mo = $("scan-lb-model");
    if (se) row.serial = se.value.trim().toUpperCase();
    if (mo) row.model = mo.value.trim();
    scanRenderRows();
  }
  if (close !== false && lb) lb.classList.remove("open");
}

/* ====================== post-scan routing ================================ */
function scanFindAsset(sn) {
  const inv = scanInv();
  if (!inv || !sn) return null;
  const key = scanKey(sn);
  return inv.assets.find(a => a.key === key) || null;
}

/* Decision 3: existing serial → warn FIRST, then jump to its Asset 360 page. */
function scanWarnAndJump(asset) {
  scanToast(`Serial ${asset.sn} is already in inventory — opening its Asset 360 page.`);
  setTimeout(() => { try { openAsset360(asset.key); } catch (e) { /* noop */ } }, 800);
}

/* ---- batch selection ---- */
function scanSelRows() {
  return QBR._scanUI.rows.filter(r => r.sel);
}
function scanRefreshBatchUI() {
  const ui = QBR._scanUI;
  const n = ui.rows.filter(r => r.sel).length;
  const bt = $("scan-batch-tag"), bc = $("scan-batch-clear"), cc = $("scan-batch-count");
  if (bt) { bt.disabled = !n; bt.textContent = n ? `Tag batch (${n})` : "Tag batch"; }
  if (bc) bc.disabled = !n;
  if (cc) cc.textContent = n ? `${n} selected` : "";
  const sa = $("scan-sel-all");
  if (sa) {
    const all = ui.rows.length > 0 && ui.rows.every(r => r.sel);
    sa.checked = all;
    sa.indeterminate = !all && n > 0;
  }
}

/* ---- batch tag: one school / SQ / DR# / owner for many serials ---- */
function scanBatchTagDialog() {
  const host = $("scan-tag-host");
  const rows = scanSelRows().filter(r => (r.serial || "").trim());
  if (!host || !rows.length) { scanToast("Select at least one row with a serial first."); return; }
  const inv = scanInv();
  const clients = inv ? [...new Set(inv.assets.map(a => a.client).filter(Boolean))].sort() : [];
  const today = (() => { const d = new Date(); return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); })();
  const already = rows.filter(r => scanFindAsset(r.serial.trim()));
  const fresh = rows.filter(r => !scanFindAsset(r.serial.trim()));
  const noSerial = scanSelRows().length - rows.length;
  host.innerHTML =
    `<div class="modal-overlay open" id="scan-batch-modal"><div class="scan-dialog card-box" role="dialog" aria-modal="true" aria-label="Tag batch to school">
      <h6>Tag ${rows.length} serial${rows.length === 1 ? "" : "s"} to one school</h6>
      <p class="small text-muted mb-2"><code>${esc(fresh.slice(0, 6).map(r => r.serial.trim()).join(", "))}</code>${fresh.length > 6 ? ` <span class="text-muted">+${fresh.length - 6} more</span>` : ""}</p>
      ${already.length ? `<p class="small text-warning mb-1">⚠ ${already.length} already in inventory — will be skipped.</p>` : ""}
      ${noSerial ? `<p class="small text-muted mb-1">${noSerial} selected row${noSerial === 1 ? "" : "s"} without a serial — will be skipped.</p>` : ""}
      <div class="row g-2 mt-1">
        <div class="col-md-6"><label class="form-label small">School / Client *</label>
          <input id="btg-client" class="form-control form-control-sm" list="dl-scan-bclients" autocomplete="off">
          <datalist id="dl-scan-bclients">${clients.map(c => `<option value="${esc(c)}"></option>`).join("")}</datalist></div>
        <div class="col-md-6"><label class="form-label small">Date *</label>
          <input id="btg-date" type="date" class="form-control form-control-sm" value="${today}"></div>
        <div class="col-md-6"><label class="form-label small">Purchase location</label>
          <input id="btg-loc" class="form-control form-control-sm" placeholder="e.g. Twireless Megamall"></div>
        <div class="col-md-6"><label class="form-label small">SQ number</label>
          <input id="btg-sq" class="form-control form-control-sm" placeholder="e.g. 121823-002"></div>
        <div class="col-md-6"><label class="form-label small">DR #</label>
          <input id="btg-dr" class="form-control form-control-sm"></div>
        <div class="col-md-6"><label class="form-label small">Owner (assigned to)</label>
          <input id="btg-assigned" class="form-control form-control-sm" placeholder="Person receiving the units"></div>
      </div>
      <div class="d-flex gap-2 mt-2">
        <button type="button" class="btn btn-sm btn-primary" id="btg-go">Tag ${fresh.length} asset${fresh.length === 1 ? "" : "s"}</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="btg-cancel">Cancel</button>
      </div>
      <div id="btg-msg" class="small mt-1" aria-live="polite"></div>
    </div></div>`;
  const close = () => { host.innerHTML = ""; };
  $("btg-cancel").addEventListener("click", close);
  $("scan-batch-modal").addEventListener("click", e => { if (e.target.id === "scan-batch-modal") close(); });
  $("btg-go").addEventListener("click", () => {
    const f = {
      client: $("btg-client").value.trim(), date: $("btg-date").value,
      location: $("btg-loc").value.trim(), sq: $("btg-sq").value.trim(),
      dr: $("btg-dr").value.trim(), assigned: $("btg-assigned").value.trim(),
    };
    if (!f.client || !f.date) { $("btg-msg").textContent = "School/client and date are required."; return; }
    // re-check: something may have been registered while the dialog was open
    const todo = fresh.filter(r => !scanFindAsset(r.serial.trim()));
    const skipped = fresh.length - todo.length;
    if (!todo.length) { $("btg-msg").textContent = "Nothing left to tag — all selected serials are already in inventory."; return; }
    const list = todo.map(r => ({ sn: r.serial.trim(), model: r.model || null, cat: "Laptop",
      dr: f.dr || null, client: f.client, wstart: f.date }));
    QBR.invIntake(list);
    const inv2 = scanInv();
    todo.forEach(r => {
      const a = inv2.assets.find(x => x.key === scanKey(r.serial.trim()));
      if (a && f.assigned) a.contact = f.assigned;
    });
    QBR.invDeploy(todo.map(r => r.serial.trim()), f.client, f.date, f.sq || null);
    todo.forEach(r => {
      const dep = inv2.deployments.slice().reverse().find(d => d.key === scanKey(r.serial.trim()));
      if (dep) dep.remarks = "Tagged via scan (batch)" + (f.location ? " · Purchased: " + f.location : "");
    });
    QBR.invLog("scan batch tag", todo.length + " → " + f.client + (f.sq ? " (SQ " + f.sq + ")" : ""));
    QBR._scanUI.rows.forEach(r => { r.sel = false; });
    scanFlagDuplicates();
    scanRenderRows();
    close();
    if (typeof renderAll === "function") renderAll();
    let msg = `Tagged ${todo.length} asset${todo.length === 1 ? "" : "s"} to ${f.client}.`;
    if (already.length + skipped) msg += ` ${already.length + skipped} skipped (already in inventory).`;
    if (noSerial) msg += ` ${noSerial} skipped (no serial).`;
    scanToast(msg);
  });
  setTimeout(() => { const c = $("btg-client"); if (c) c.focus(); }, 60);
}

function scanRoute(row, action) {
  const sn = (row.serial || "").trim();
  if (!sn) { scanToast("No serial extracted yet — edit the row or rescan first."); return; }
  scanFlagDuplicates();
  if (row.dup && !row.dupAck) { scanDupDialog(row, action); return; }
  scanRouteGo(row, action);
}

function scanRouteGo(row, action) {
  if (action === "tag") scanOpenTagDialog(row);
  else if (action === "deploy-fill") scanPrefillDeploy(row);
  else scanPrefillIntake(row); // default + "intake" action
}

/* Duplicate serial confirmation: "already in inventory" offers to view the
 * existing asset (Asset 360) or add anyway; "duplicate in batch" offers to
 * add anyway or cancel. Reuses the tag dialog host. */
function scanDupDialog(row, action) {
  const host = $("scan-tag-host");
  if (!host) { scanRouteGo(row, action); return; }
  const sn = (row.serial || "").trim().toUpperCase();
  const asset = scanFindAsset(sn);
  const qCount = QBR._scanUI.rows.filter(r => r !== row &&
    (r.serial || "").trim().toUpperCase() === sn).length;
  const invLine = asset
    ? `<p class="mb-1">Serial <code>${esc(sn)}</code> is <strong>already in inventory</strong>.</p>
       <p class="small text-muted mb-2">Client: ${esc(asset.client || "—")} · Status: ${esc(asset.status || "—")} · Model: ${esc(asset.model || "—")}</p>`
    : "";
  const qLine = qCount
    ? `<p class="mb-1">Serial <code>${esc(sn)}</code> appears <strong>${qCount + 1} times</strong> in this scan batch.</p>`
    : "";
  host.innerHTML =
    `<div class="modal-overlay open" id="scan-dup-modal"><div class="scan-dialog card-box" role="dialog" aria-modal="true" aria-label="Duplicate serial">
      <h6>⚠ Duplicate serial</h6>
      ${invLine}${qLine}
      <div class="d-flex flex-wrap gap-2 mt-2">
        ${asset ? `<button type="button" class="btn btn-sm btn-outline-primary" id="dup-view">View existing asset</button>` : ""}
        <button type="button" class="btn btn-sm btn-primary" id="dup-go">Add anyway</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="dup-cancel">Cancel</button>
      </div>
    </div></div>`;
  const close = () => { host.innerHTML = ""; };
  $("dup-cancel").addEventListener("click", close);
  $("scan-dup-modal").addEventListener("click", e => { if (e.target.id === "scan-dup-modal") close(); });
  const vw = $("dup-view");
  if (vw) vw.addEventListener("click", () => { close(); scanWarnAndJump(asset); });
  $("dup-go").addEventListener("click", () => { close(); row.dupAck = true; scanRouteGo(row, action); });
}

function scanModelOption(model) {
  const m = String(model || "").toLowerCase();
  if (/thinkpad/.test(m)) return 0;
  if (/thinkcentre/.test(m)) return 1;
  if (/c22|monitor/.test(m)) return 2;
  return 3; // "Other"
}

function scanPrefillIntake(row) {
  if (typeof QBR._invUI === "undefined") return;
  QBR._invUI.form = "intake";
  if (typeof goToTab === "function") goToTab("dash-inventory");
  if (typeof renderAll === "function") renderAll();
  setTimeout(() => {
    const ta = $("in-serials");
    if (ta) ta.value = row.serial || "";
    const ms = $("in-model");
    if (ms && row.model) ms.selectedIndex = scanModelOption(row.model);
    scanToast("Scanned details filled into the intake form — review and click Register.");
    const first = $("in-serials"); if (first) first.focus();
  }, 80);
}

function scanPrefillDeploy(row) {
  if (typeof QBR._invUI === "undefined") return;
  QBR._invUI.form = "deploy";
  if (typeof goToTab === "function") goToTab("dash-inventory");
  if (typeof renderAll === "function") renderAll();
  setTimeout(() => {
    const ta = $("dp-serials");
    if (ta) ta.value = row.serial || "";
    const hit = scanFindAsset(row.serial || "");
    scanToast(hit ? `Serial ${row.serial} is already registered — fill in the client and deploy.`
                  : "Serial filled into the deployment form. Register the unit first if it is new.");
  }, 80);
}

/* Tag dialog: Client*, Date*, Purchase location, SQ number, DR #, Assigned to. */
function scanOpenTagDialog(row) {
  const host = $("scan-tag-host");
  if (!host) return;
  const inv = scanInv();
  const clients = inv ? [...new Set(inv.assets.map(a => a.client).filter(Boolean))].sort() : [];
  const today = (() => { const d = new Date(); return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); })();
  host.innerHTML =
    `<div class="modal-overlay open" id="scan-tag-modal"><div class="scan-dialog card-box" role="dialog" aria-modal="true" aria-label="Tag asset to client">
      <h6>Tag <code>${esc(row.serial)}</code> to a client</h6>
      <div class="row g-2 mt-1">
        <div class="col-md-6"><label class="form-label small">Client *</label>
          <input id="tg-client" class="form-control form-control-sm" list="dl-scan-clients" autocomplete="off">
          <datalist id="dl-scan-clients">${clients.map(c => `<option value="${esc(c)}"></option>`).join("")}</datalist></div>
        <div class="col-md-6"><label class="form-label small">Date *</label>
          <input id="tg-date" type="date" class="form-control form-control-sm" value="${today}"></div>
        <div class="col-md-6"><label class="form-label small">Purchase location</label>
          <input id="tg-loc" class="form-control form-control-sm" placeholder="e.g. Twireless Megamall"></div>
        <div class="col-md-6"><label class="form-label small">SQ number</label>
          <input id="tg-sq" class="form-control form-control-sm" placeholder="e.g. 121823-002"></div>
        <div class="col-md-6"><label class="form-label small">DR #</label>
          <input id="tg-dr" class="form-control form-control-sm"></div>
        <div class="col-md-6"><label class="form-label small">Assigned to</label>
          <input id="tg-assigned" class="form-control form-control-sm" placeholder="Person receiving the unit"></div>
      </div>
      <div class="small text-muted mt-1">Model: ${esc(row.model || "—")} · creates the asset, a deployment record and a journal entry.</div>
      <div class="d-flex gap-2 mt-2">
        <button type="button" class="btn btn-sm btn-primary" id="tg-go">Tag asset</button>
        <button type="button" class="btn btn-sm btn-outline-secondary" id="tg-cancel">Cancel</button>
      </div>
      <div id="tg-msg" class="small mt-1" aria-live="polite"></div>
    </div></div>`;
  const close = () => { host.innerHTML = ""; };
  $("tg-cancel").addEventListener("click", close);
  $("scan-tag-modal").addEventListener("click", e => { if (e.target.id === "scan-tag-modal") close(); });
  $("tg-go").addEventListener("click", () => {
    const f = {
      client: $("tg-client").value.trim(), date: $("tg-date").value,
      location: $("tg-loc").value.trim(), sq: $("tg-sq").value.trim(),
      dr: $("tg-dr").value.trim(), assigned: $("tg-assigned").value.trim(),
    };
    if (!f.client || !f.date) { $("tg-msg").textContent = "Client and date are required."; return; }
    // re-check: someone may have registered it while the dialog was open
    const dup = scanFindAsset(row.serial);
    if (dup) { close(); scanWarnAndJump(dup); return; }
    const sn = row.serial.trim();
    const n = QBR.invIntake([{ sn, model: row.model || null, cat: "Laptop",
      dr: f.dr || null, client: f.client, wstart: f.date }]);
    if (!n) { $("tg-msg").textContent = "Could not register the asset."; return; }
    const inv2 = scanInv();
    const a = inv2.assets.find(x => x.key === scanKey(sn));
    if (a && f.assigned) a.contact = f.assigned;
    QBR.invDeploy([sn], f.client, f.date, f.sq || null);
    const dep = inv2.deployments.slice().reverse().find(d => d.key === scanKey(sn));
    if (dep) dep.remarks = "Tagged via scan" + (f.location ? " · Purchased: " + f.location : "");
    QBR.invLog("scan tag", sn + " → " + f.client + (f.location ? " (" + f.location + ")" : ""));
    close();
    scanToast(`Asset ${sn} tagged to ${f.client}.`);
    if (typeof renderAll === "function") renderAll();
    setTimeout(() => { const hit2 = scanFindAsset(sn); if (hit2) { try { openAsset360(hit2.key); } catch (e) {} } }, 700);
  });
  setTimeout(() => { const c = $("tg-client"); if (c) c.focus(); }, 60);
}

/* ---------- Node export guard (test harness only; inert in browser) ------ */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { QBR, scanParseOCRText, scanClassifyBarcode, scanMergeFindings, scanPickSerial, scanSplitMtmSn, scanLooksLikeSerial };
}
