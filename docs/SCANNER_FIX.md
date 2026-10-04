# Barcode & OCR Scanner — Fix Notes and Dependencies

**Module:** `qbr-app/js/scan.js` · **Version:** 1.1.0 → **1.2.0** · **Date:** 2026-10-04
**Applies to:** the QBR Dashboard Inventory/Scan build (based on v1.9.0)
**Status:** fixed and tested; not yet merged into the live `QBR-Dashboard_1.9.0` folder or the GitHub repo.

---

## 1. Summary

The Scan page reads photos of laptop/desktop box labels and pulls out the **serial number, model and product key**. It does this two ways:

- **Barcode decoding**: reads the barcodes printed on the label.
- **OCR**: reads the printed text.

Before the fix, neither worked on the team's Windows PCs when the app was opened by double-clicking `index.html`:

| # | Problem | Effect |
|---|---|---|
| 1 | The barcode reader used Chrome/Edge's built-in `BarcodeDetector`, which **does not exist on Windows or Linux** | No barcode was ever decoded; the page wrongly said "use Chrome/Edge" |
| 2 | The OCR engine (Tesseract) tried to load its files in a way browsers **block on `file://`** | Scans **froze** on "Scanning…"; the "OCR unavailable" fallback never appeared |
| 3 | Photos went to OCR as-is, with no straightening or enlarging | Tilted or low-resolution photos returned nothing |
| 4 | The serial-number rule accepted a lone letter "S" | It could pick part of the model number (`21M4S5VD00` → `5VD00…`) |
| 5 | UI: no scan queue, stale status text | Photos added during a scan stayed "Pending"; the counter showed "0 images" |

After the fix, all five are resolved. Scanning works fully offline from `file://` and from a web server.

---

## 2. Libraries and dependencies

All dependencies are **vendored** (copied into the app folder). There is **no CDN, no internet access and no install step**, which preserves the dashboard's offline, zero-build design.

### 2.1 Third-party libraries

| Library | Version | License | Location in app | Size | Role | Loaded when |
|---|---|---|---|---|---|---|
| **ZXing** (`@zxing/library`, UMD build) | 0.23.0 | Apache-2.0 | `libs/zxing/zxing.min.js` (+ `LICENSE`) | 0.36 MB | Barcode/QR decoding where the browser has no built-in detector (Windows, Linux) | Lazily, on the first scan |
| **Tesseract.js** | 5.1.1 | Apache-2.0 | `libs/tesseract/tesseract.min.js`, `worker.min.js` | 0.19 MB | OCR engine API + Web Worker | `tesseract.min.js` at page load (already in the original build) |
| **tesseract.js-core** (WebAssembly OCR engine) | 5.1.1 | Apache-2.0 | `libs/tesseract/tesseract-core-simd-lstm.wasm.js`, `tesseract-core-lstm.wasm.js` | 3.9 MB each | The compiled Tesseract engine (wasm embedded in the `.js`) | On the first scan; one variant, chosen by CPU feature test |
| **English model** (`eng.traineddata`, "fast" variant, from tesseract-ocr/tessdata_fast) | 4.0.0 | Apache-2.0 | `libs/tesseract/eng.traineddata.gz` | 2.0 MB | Language data for OCR | On the first scan |

**New in this fix:** ZXing. Tesseract.js, its core and the English model were already in the build; the fix changes **how** they are loaded (see §3.2).

### 2.2 Generated files (built from the vendored Tesseract files)

| File | Size | Content |
|---|---|---|
| `libs/tesseract/offline/worker.js` | 0.12 MB | Tesseract worker source (with one bug fix, §3.2) wrapped as a plain script |
| `libs/tesseract/offline/core-simd-lstm.js` | 3.8 MB | SIMD core, wrapped as a plain script |
| `libs/tesseract/offline/core-lstm.js` | 3.8 MB | Non-SIMD core (older CPUs), wrapped as a plain script |
| `libs/tesseract/offline/eng.js` | 2.5 MB | English model (base64), wrapped as a plain script |

Created by `node tests/build-tesseract-offline.cjs`. Re-run it whenever anything in `libs/tesseract/` is updated.

### 2.3 Browser features used (built in, nothing to install)

| Feature | Used for |
|---|---|
| `BarcodeDetector` (Shape Detection API) | Fast barcode path on macOS / ChromeOS / Android, used when present |
| Canvas 2D | Image loading, scaling, banding, grayscale |
| Web Workers + Blob URLs | Running OCR off the main thread; blob URLs make this work on `file://` |
| WebAssembly (+ SIMD feature test) | Running the Tesseract engine |

### 2.4 Test-only tools (not shipped with the app)

| Tool | Used for |
|---|---|
| Node.js | `tests/scan-tests.cjs` (parser unit tests), `tests/build-tesseract-offline.cjs` |
| Playwright + Chromium | `tests/ui-scan.cjs` (end-to-end Scan page test) |
| Python `python-barcode` + Pillow | Generating the synthetic test labels (one-off; images are committed in `tests/`) |

### 2.5 Explicitly not used

- **No CDN or network calls.** Everything loads from the app folder.
- **No "ZXing" script tags from the original standalone extractor.** Those pointed at unrelated npm packages (`@aspect-build/aspect-workflows`, `@nicolo-ribaudo/chokidar-2`) and were never real ZXing. This fix vendors the genuine `@zxing/library`.

---

## 3. How each problem was fixed

### 3.1 Barcodes on Windows: ZXing added

**Root cause.** Chrome and Edge ship `window.BarcodeDetector` only on macOS, ChromeOS and Android ([Chrome docs](https://developer.chrome.com/docs/capabilities/shape-detection)). On Windows it is `undefined`, so the old code returned zero barcodes every time.

**Fix.** `scanDetectBarcodes()` now works in two steps:

1. If the browser has `BarcodeDetector`, use it (fast path). Then retry once on a black-and-white version.
2. If there is no detector, or it found nothing, use **ZXing** (`scanZxingDetect()`).

ZXing decodes **one** barcode per attempt, but a Lenovo box label stacks 5–10 barcodes (serial, MTM, MTM+serial, UPC, MAC, …). So the image is scanned in pieces:

```
whole image  → 1 attempt with TRY_HARDER (also finds QR / Data Matrix / PDF417 / Aztec)
7 bands      → overlapping horizontal strips, 50% overlap
12 bands     → finer strips, 50% overlap
each strip   → GlobalHistogram binarizer, then Hybrid binarizer
```

- Images are capped at 2000 px on the long side to keep it fast: **about 0.4–1.5 s per photo**.
- Formats: Code 128, Code 39, Code 93, Codabar, EAN-13, EAN-8, ITF, UPC-A, UPC-E (strips); QR, Data Matrix, PDF417, Aztec (whole image).
- ZXing is loaded with a normal `<script>` tag injected on first use, which browsers allow on `file://`.

After decoding, the existing logic picks the serial. Barcodes are classified as key / serial / other (UPC, MAC, UUID are ignored), Lenovo's MTM+serial barcode is split, and the candidates are scored by agreement with the OCR text.

### 3.2 OCR froze on `file://`: offline engine loader

**Root cause.** Tesseract.js starts a Web Worker that **fetches** its engine and language files. On `file://` the browser blocks those requests. Tesseract didn't report an error; its start-up promise simply never finished. So the "OCR unavailable" fallback never ran and the scan froze.

**Fix.** `scanGetOCRWorker()` now has two start paths. Each is limited to 60 seconds, so it can never hang again.

| Opened via | Start path |
|---|---|
| `http://` / `https://` | Standard Tesseract loader (files fetched from `libs/tesseract/`). Falls back to the offline path if it fails. |
| `file://` (double-click) | **Offline loader** (`scanStartOCROffline()`), below |

The offline loader works around three browser restrictions:

1. **Files can't be fetched on `file://`, but `<script>` tags still load.** The engine, worker and model were pre-wrapped as plain scripts (`libs/tesseract/offline/*.js`, §2.2), and the page injects them.
2. **A worker can't load a second blob on a `file://` page.** So the engine source and worker source are joined into **one** blob, and the worker is started from it. Because the engine is already defined inside the worker, Tesseract skips its own (blocked) loader.
3. **The model can't be fetched.** It is decoded in the page and handed to the worker in memory as `{ code: "eng", data: <bytes> }`.

The loader also checks whether the CPU supports WebAssembly SIMD and picks the matching engine file. After start-up it frees the ~10 MB of source text from page memory.

**Tesseract.js 5.1.1 bug found and worked around.** When language data is passed in memory, the worker builds the language name from the raw bytes (`t.data`) instead of the code (`t.code`) and hangs. The build script fixes that one expression in the **offline copy only**. It stops with an error if the expression isn't found exactly once, which protects against a silent mis-patch after a library upgrade. The original `worker.min.js` is untouched.

### 3.3 Tilted and low-resolution photos: image preparation

`scanRunOCR()` now:

1. **Enlarges** images whose long side is under 1200 px to 1600 px.
2. **Straightens** the text (`recognize(image, { rotateAuto: true })`). Testing showed this was the single biggest improvement.
3. **Retries** in sparse-text mode (PSM 11) when the first pass finds no labelled serial. Both passes' text is passed to the serial vote.

### 3.4 Serial-number rule

| | Pattern |
|---|---|
| **Before** | `/(?:S(?:erial)?[\s.]*(?:N(?:o\|umber)?)?\|S\/N\|SN)[\s.:=]*([A-Z0-9][\w\-]{5,25})/gi`, where a lone "S" counted as a label |
| **After** | `/(?:^\|[^A-Z0-9])(?:S\s*\/\s*N\|SN\|SERIAL(?:\s*(?:NUMBER\|NO\.?\|#))?)\s*[:.=#]?\s*([A-Z0-9][A-Z0-9\-]{5,25})/gi` |

- A label is required (`SN`, `S/N`, `(S) SN`, `Serial`, `Serial Number/No/#`), and it must not be glued to a preceding letter or digit. So the "S" inside `21M4S5VD00` can no longer start a match.
- All labelled matches are collected. The one shaped like a unit serial (7–10 characters mixing letters and digits) wins.

### 3.5 Scan page fixes

- **Queue** (`scanPump()`): one scan at a time, in order, until nothing is pending. Photos added mid-scan, "Scan all" and ↻ all go through it.
- **Engine note** (`#scan-engine-note`): shows what is actually in use, e.g. "barcode/QR (ZXing, offline) · label text OCR (on-device, offline engine)". If OCR fails it shows the reason. It refreshes after every scan.
- **Counter and empty message** now update as photos are added or removed.

---

## 4. Files changed

| File | Change |
|---|---|
| `qbr-app/js/scan.js` | **Modified**: everything in §3. Only this app file was edited. |
| `qbr-app/libs/zxing/zxing.min.js`, `LICENSE` | **New**: vendored ZXing 0.23.0 |
| `qbr-app/libs/tesseract/offline/*.js` (4 files) | **New**: generated offline OCR assets |
| `tests/build-tesseract-offline.cjs` | **New**: generates the offline assets and applies the guarded bug fix |
| `tests/scan-tests.cjs` | **New**: 23 parser / vote unit tests (Node) |
| `tests/ui-scan.cjs` | **New**: end-to-end Scan page test (Playwright), runs on `file://` or `http://` |
| `tests/scan-label-clean.png`, `scan-label-phone.jpg`, `scan-label-lowres.jpg` | **New**: synthetic test labels (fictional serial `PF62SDPW`) |
| `tests/ui-smoke.cjs` | **Updated**: nav count now 15 core pages + optional Inventory/Scan; Asset 360 treated as a detail page |
| `DEV_NOTES.md`, `qbr-dashboard-rebuild-guide.md` | **Updated**: scanner section |

Not changed: `index.html`, `app.js`, `inventory.js`, `styles.css` and all dashboard engine files.

---

## 5. Test results

Run in headless Chromium on synthetic Lenovo-style labels.

| Test label | Before | After (`file://` and `http://`) |
|---|---|---|
| Clean label | Serial by OCR only on `http://`; 0 barcodes | Serial + model; all 5 barcodes decoded |
| Tilted, blurred phone shot (1468 × 1652) | Nothing | Serial + model; 2 barcodes |
| Low-res tilted photo (415 × 474) | Nothing | Serial + model |
| Opened by double-click (`file://`) | Froze | 3 photos in about 6.6 s, offline engine |

| Suite | Result |
|---|---|
| `node tests/scan-tests.cjs` | 23 / 23 |
| `tests/ui-scan.cjs` on `file://` and on `http://` | 11 / 11 each, zero console errors |
| Dashboard regression: engine harness / UI smoke / UI features | 104 & 100 / 71 / 37, all passing |

How to run:

```powershell
node tests/scan-tests.cjs
$env:PW="<folder containing node_modules\playwright>"; node tests/ui-scan.cjs                          # file://
$env:PW="<folder containing node_modules\playwright>"; node tests/ui-scan.cjs http://localhost:8000/index.html
```

---

## 6. Maintenance

- **Updating Tesseract.js:** replace the files in `libs/tesseract/`, then run `node tests/build-tesseract-offline.cjs`. If the build reports that the language-join expression wasn't found, check whether the new version fixed the `t.data` / `t.code` bug. If it did, remove that patch step from the script.
- **Updating ZXing:** replace `libs/zxing/zxing.min.js` with the new `@zxing/library` UMD build (`umd/index.min.js`) and run `tests/ui-scan.cjs`.
- **Optional clean-up (~24 MB):** these files are never used, because the app runs Tesseract in LSTM-only mode and the `.wasm.js` cores embed their wasm. Both tests still passed 11 / 11 with them removed. They are not deleted yet:
  - `tesseract-core.wasm`, `tesseract-core.wasm.js`, `tesseract-core-simd.wasm`, `tesseract-core-simd.wasm.js`
  - `tesseract-core-lstm.wasm`, `tesseract-core-simd-lstm.wasm`
  - `eng.traineddata` (uncompressed copy)

---

## 7. Known limits

- Real phone photos vary (glare, motion blur, steep angles). Results stay **editable**, and serial characters are never auto-corrected, so a person confirms what was read.
- The first scan in a session takes a few extra seconds while the OCR engine starts. Later scans reuse it.
- Very blurry 1D barcodes may not decode. The OCR text and the serial vote usually still recover the serial.
- Photos stay in browser memory only; nothing is saved or uploaded.
