# Inventory Module — Dev Notes (v1.1.0 · 2026-10-04)

Full-flow laptop/desktop/monitor inventory + support-ticket system, built **inside**
the QBR dashboard as new pages (not a standalone app).

## How to run

1. Open `qbr-app/index.html` in a browser (double-click; `file://` works — fully offline).
2. Upload `data/SAMPLE_Lenovo_Inventory_SYNTH.xlsx` via the upload area (it merges
   with any workbooks already loaded, like the existing accumulation flow).
3. Click **Inventory** in the left sidebar (new "Inventory" nav group).

## What was added

**`data/SAMPLE_Lenovo_Inventory_SYNTH.xlsx`** (new) — synthetic fixture, all names
fictional (seeded RNG for reproducibility). Sheets mirror the real workbook layout:
- `02 DEVICES` — 600 assets (380 laptops / 120 desktops / 100 monitors; ThinkPad E14
  Gen 5, ThinkCentre M75t Gen 2, C22-20 monitors; ~8% unassigned = In Stock;
  warranty dates spanning 2023–2026 with edge cases: ~15 expiring <90d, ~10 expired)
- `03 TECH SUPPORT LOGS` — 85 tickets (76 Completed / 9 Open; 6 open >7 days for the
  aging flag; 3 "lemon" serials with 4 tickets each; 4 tickets with serials NOT in
  DEVICES for the unmatched flag)
- `04 RAKSO INV.` — 25 SQs with blank-SQ continuation rows (forward-fill) and mixed
  laptop/desktop/monitor serial columns, incl. N/A and `-` sentinels
- `06 PIPELINE` — 20 rows, several SIGNED-but-not-DELIVERED
- `07 PURCHASE ORDER` — 8 rows

**`qbr-app/js/inventory.js`** (new, ~830 lines, loaded before `app.js`) —
- Parsers (header-pattern based, reusing `makeResolver`/`findSheet` from
  `excel-loader.js`): `invParseDeviceRows`, `invParseSupportRows`,
  `invParseRaksoRows` (SQ/client forward-fill + pivot of S/N / DEKTOP-S/N /
  Monitor-S/N into one row per asset; skips N/A/-/blank sentinels),
  `invParsePipelineRows`, `invParsePORows`. Entry: `QBR.parseInventoryBuffers(buffers)`.
- Serial key normalization: `QBR.invSerialKey` (uppercase/trim/collapse whitespace/
  strip invisible marks) — canonical join key, never displayed.
- Derived status (`QBR.invAssetStatus`): In Stock / Deployed / In Repair (open ticket
  or condition) / Retired, plus manual override via UI.
- Flag engine (`QBR.invComputeFlags`): aging tickets (>7d open), lemons (≥3 tickets),
  warranty expiring <90d / expired, unmatched ticket serials, stalled pipeline
  (SIGNED not DELIVERED). Thresholds on `QBR.INV_THRESH`.
- Write-back: `QBR.invIntake` (register), `QBR.invDeploy` (assign), `QBR.invAddTicket`,
  `QBR.invResolveTicket`, `QBR.invSetStatus` — all mutate the in-memory model,
  append to `QBR.invJournal` (audit trail), and re-render.
- Export: `QBR.invExportWorkbook()` regenerates the .xlsx from the model via the
  vendored SheetJS (mirrors `exportCleanData`), including an EDIT LOG sheet.
- UI: `renderInventory()` (KPIs, clickable flag cards, filters, assets/tickets/
  pipeline/PO tables, intake/deploy/ticket forms) and `renderAsset360Panel()` /
  `openAsset360(key)` — per-serial 360° page imitating School 360 (facts, warranty
  timeline, deployment record, ticket history, rule-based recommended actions).
- Node export guard exposes the pure parsers for testing (inert in browser).

**`qbr-app/index.html`** — sidebar "Inventory" nav group (`data-tab="dash-inventory"`);
`<section id="dash-inventory">` and `<section id="dash-asset360" data-deck="skip">`
panels; `<script src="js/inventory.js">` before `app.js`.

**`qbr-app/js/app.js`** (4 additive hooks, existing pages untouched) —
- `processBuffers`: parses inventory sheets from the same buffers into
  `APP.model.inventory` (try/catch — inventory failure can never break the main flow)
- `renderAll`: calls `renderInventory()` + `renderAsset360Panel()` (typeof-guarded)
- `TAB_FILTERS`: `dash-inventory: []`, `dash-asset360: []` (own in-panel filters)
- `updateNavBadges`: open-ticket count on the Inventory nav item

**`qbr-app/css/styles.css`** — print rules (`dash-inventory` in the print list;
`dash-asset360` excluded from the deck like School 360) + §20 component styles
(flag cards, asset tables, ticket timeline).

## Scanner integration (v1.1.0 · 2026-10-04)

Barcode/QR + OCR label extraction integrated from the standalone Asset Extractor,
per user decisions: (1) both a dedicated **Scan** page under the Inventory nav AND
scan buttons inside the Intake ("Scan box label") and Deployment ("Scan unit")
forms; (2) photo upload only, no live camera; (3) existing serial → warn first,
then jump to Asset 360; (4) tag mode captures Client*, Date*, Purchase location,
SQ number, DR #, Assigned to (*required); (5) fully offline — vendor open-source
libs (Tesseract.js) into the app.

**Vendored libs — `qbr-app/libs/tesseract/` (34.6 MB total, Apache-2.0):**
- `tesseract.min.js` + `worker.min.js` (Tesseract.js v5.1.1, from jsdelivr)
- `tesseract-core{,-lstm,-simd,-simd-lstm}.wasm{,.js}` ×4 variants (5.1.1) — the
  worker picks simd/non-simd × lstm/plain at runtime, so all four ship
- `eng.traineddata` (4.1 MB, "fast" variant from tesseract-ocr/tessdata_fast) +
  `eng.traineddata.gz` (2 MB) — the worker requests the `.gz` by default
- Loaded via `<script src="libs/tesseract/tesseract.min.js">` before `js/scan.js`;
  worker/core/lang paths are passed explicitly (`libs/tesseract`).

**IMPORTANT — no ZXing:** the standalone's `<script>` tags labeled "ZXing" were
bogus (they pointed at `@aspect-build/aspect-workflows` and
`@nicolo-ribaudo/chokidar-2`, not ZXing). They were NOT copied. Barcode detection
uses the browser's native `window.BarcodeDetector` only (Chrome/Edge) — zero new
dependencies on that path.

**`qbr-app/js/scan.js`** (new, loaded after `inventory.js`, before `app.js`) —
- Pure, Node-testable: `scanParseOCRText` (serial S/N patterns, Windows-key
  patterns, brand-name model list incl. ThinkCentre), `scanClassifyBarcode`
  (key vs serial), `scanMergeFindings` (barcodes win over OCR; strips `S/N:` prefix).
- Browser detection: `QBR.scanImage(dataUrl)` → `{barcodes, serial, productKey,
  model, ocrOk, barcodeOk}`. Native BarcodeDetector (13 formats) with a
  binarized-contrast fallback pass; OCR via a lazily-created Tesseract worker
  (reused across scans, `cacheMethod: "none"`).
- Scan page (`renderScan()`): upload zone (drag-drop + file input, multi-image,
  auto-scan on add), editable results table (thumbnail, model, serial, product
  key, raw barcodes, status), per-row actions — **Intake** (prefill intake form),
  **Tag** (tag dialog), rescan, photo download, delete, thumbnail lightbox.
- Post-scan routing: serial normalized via `QBR.invSerialKey()` → if the serial
  exists → toast warning first, then `openAsset360()`; else Intake prefills the
  intake form (serial + model matched to the model dropdown), Tag opens the tag
  dialog (Client*, Date*, Purchase location, SQ number, DR #, Assigned to) →
  confirm creates the asset (`invIntake`), deployment record (`invDeploy`,
  remarks amended with purchase location), contact = Assigned to, journal entry,
  then jumps to the new Asset 360 page.
- Scan buttons: `QBR.scanSetTarget("intake"|"deploy")` routes through the Scan
  page in target mode; the first successful scan auto-applies through the same
  routing (exists-check included). Photos stay in memory only (never persisted
  to IndexedDB — avoids bloat); per-photo download kept.
- OCR failure is graceful: if the Tesseract worker can't start (e.g. `file://`
  worker restrictions in some browsers — works fine over http/https), the page
  shows an "OCR unavailable, barcode detection still works" note and scanning
  continues barcode-only.
- Node export guard exposes the pure parsers (inert in browser).

**Wiring (additive):** `index.html` — `fi-scan`/`fi-scan-f` sprite symbols
(barcode glyph), Scan nav button in the Inventory group, `<section id="dash-scan"
data-deck="skip">`; `js/app.js` — `TAB_FILTERS["dash-scan"] = []`, `renderAll()`
calls `renderScan()` (typeof-guarded); `js/inventory.js` — "Scan box label" /
"Scan unit" buttons in the intake/deploy forms; `css/styles.css` — §21 scan
styles, single-tab print rule for `dash-scan`, deck exclusion like Asset 360.

## Verification (2026-10-04) — inventory module

- `node --check` passes on all 7 JS files (incl. a fix for a try/catch split introduced
  mid-edit in `processBuffers`, since corrected and re-checked).
- Node smoke test (28 assertions, real `makeResolver`/`findSheet` from excel-loader.js
  executed as a classic script): parsers, forward-fill, sentinel skipping, serial-key
  normalization, status derivation, all six flags, intake/deploy/ticket/resolve/status
  mutations, journal — **ALL PASS**.
- End-to-end through the real vendored SheetJS (`XLSX.read` on the actual .xlsx,
  `QBR.parseInventoryBuffers`): 600 assets / 85 tickets / 133 deployments / 20
  pipeline / 8 POs, Excel-serial dates parsed, flags correct — **ALL PASS**.
- Fixed during testing: (1) sheet-name grab needed substring fallback for
  `04 RAKSO INV.`; (2) ticket Resolve buttons now use stable `_n` ids instead of
  display-order indexes; (3) Asset 360 activates via `invActivateTab()` since it has
  no sidebar button (`goToTab` requires one).

## Verification (2026-10-04) — scanner

- `node --check` passes on `js/scan.js`, `js/inventory.js`, `js/app.js`
  (one stray-quote typo in scan.js caught and fixed during the check).
- Node smoke test (7 groups, pure logic): OCR text parsing (serial / product
  key / model incl. `S/N:` variants and OEM key formats), barcode
  classification (key vs serial), merge precedence (barcodes win, prefix
  stripped), `QBR.scanImage`/`scanSetTarget` exposure, version + UI state —
  **ALL PASS**.
- Wiring verified by grep: Scan nav button + panel + sprite symbols + script
  tags in `index.html`; `in-scan`/`dp-scan` buttons in `inventory.js`;
  `TAB_FILTERS` + `renderAll` hook in `app.js`; print rules in `styles.css`.
- Bugs fixed during review: (1) batch auto-scan ran `forEach` without awaiting,
  so only the first image scanned — now sequential; (2) target-mode routing
  bypassed the exists-check — now routes through the standard warn-then-jump
  path; (3) Escape-key listener stacked on every re-render — now bound once.
- NOT verifiable headless: native `BarcodeDetector` and the Tesseract worker
  need a real browser. Live detection must be smoke-tested in Chrome/Edge by
  uploading a label photo on the Scan page.

## Known limitations / next steps

- Inventory re-reads every uploaded buffer (workbooks are parsed twice — once by the
  M365 loader, once by the inventory parser). Fine at current sizes; revisit if the
  tracker grows very large.
- Edits live in memory + the edit journal; the journal is not yet persisted to
  IndexedDB (see the earlier write-back design — cache the journal alongside the
  workbook bytes so edits survive reload).
- No charts yet on the Inventory page (tables + flags only) — candidates: tickets by
  category, warranty-expiry histogram, deployments over time.
- The synthetic fixture is redacted-shape only; drop in the real workbook and the same
  parsers apply (header-pattern based, column-order free).

## v1.2 — 2026-10-04 — scanner hardened against Pedro's real label photos

Pedro supplied 10 CamScanner photos of ThinkPad E14 Gen 6 box labels and asked
why the QBR Scan page couldn't extract serials like his standalone extractor.

Findings from a Node test harness (real vendored Tesseract files + real
`scan.js` parse functions, barcode decode via @zxing/library):
- The scans' 1D barcodes are undecodable: CamScanner downsampling destroyed the
  thin bars (verified visually + ZXing fails on tight crops). Native
  BarcodeDetector may do slightly better in-browser, but OCR is the reliable
  path for these images — matching what Pedro's standalone showed (its serials
  came from OCR; its Raw Barcodes column was empty).
- Vendored Tesseract OCR reads the labels fine; `scanParseOCRText` got 9/10
  serials exactly right. One tilted photo misread chars (PF62P5LJ -> PFe2P5LU);
  the Scan row stays editable for that case.
- Model extraction missed 3/10 because OCR mangled "ThinkPad" ("Milian",
  "iim", "iia") while reading "E14 Gen 6" cleanly.

Fixes in `js/scan.js` (all covered by 15 passing Node assertions):
- `scanParseOCRText`: model fallback captures the model-number fragment
  (`E14 Gen 6` / `E14 Gen6`) when the brand word is OCR-mangled. Now 10/10
  labels yield a useful model guess.
- `scanClassifyBarcode` now returns key/serial/other: UPC-A/EAN-8/EAN-13 digit
  strings, UUIDs, and MAC-like 12-hex values are "other" (previously a 12-digit
  UPC or Product Key ID could be picked as the serial).
- `scanSplitMtmSn`: Lenovo (1S) labels concatenate MTM+serial
  ("21M4S5VD00PF62SDPW"); the trailing 8-char serial now votes as its own
  candidate instead of the 18-char blob winning.
- `scanMergeFindings` replaced "first barcode wins" with consensus scoring:
  votes per occurrence x3, Lenovo serial shape +3, presence in OCR text +4.
  Verified on the full 10-barcode label set incl. adversarial ordering (battery
  code / key ID / UPC listed first) — still picks the true serial.
- `scanMergeFindings(barcodes, parsed, ocrText)` takes the raw OCR text as a
  third arg; call site in `QBR.scanImage` updated.

Reminder for Pedro: OCR needs the app served over http(s) (Web Workers are
blocked on file://) — `python3 -m http.server` locally or the Azure Static Web
App. Barcode-only mode remains the graceful fallback with an on-page note.

## v1.3 — 2026-10-04 — quick serial lookup on the Inventory page

Pedro asked for a Search feature to look up info by serial number. The page
only had a table filter ("Search serial") with no direct jump.
- New "Look up asset" row in the Inventory actions area: type/paste a serial,
  Enter or Open. Exact or single-substring match jumps straight to Asset 360;
  several matches narrow the Assets table via the existing filter and scroll to
  it; no match shows an inline "Register it via intake" action that opens the
  intake form with the serial prefilled in the serials box.
- `invLookupSerial()` in `js/inventory.js`, wired in `invBind()`; normalizes
  with `QBR.invSerialKey`; 12 headless assertions pass (exact / substring /
  multi / none / register-prefill / empty).

## Scanner fixes — scan.js v1.2.0 (2026-10-04)

**Problems found (reproduced headless in Chromium):**
1. **Barcodes never decoded on Windows.** Chrome/Edge ship `window.BarcodeDetector` only on macOS,
   ChromeOS and Android — not Windows or Linux. On the team's PCs the barcode path was dead, and the page
   told users "barcode unavailable — use Chrome/Edge" while they were already on Chrome/Edge.
2. **OCR hung on `file://`.** Opening `index.html` by double-click (the normal way) made Tesseract's worker
   start-up fetch its core/language files, which browsers block on `file://`. The promise never settled,
   so `failed` was never set, the "OCR unavailable" fallback never appeared, and scans stayed "Scanning…".
3. **Tilted / low-res photos returned nothing.** No deskew or upscaling before OCR; barcode text dominated.
4. **Serial parser bug.** The SN regex accepted a bare `S`, so `MTM: 21M4S5VD00` could yield `5VD00…`.
5. **UI:** photos added while a scan ran stayed "Pending"; the engine note never updated; the image
   counter / "No scans yet" text went stale.

**Fixes (`js/scan.js` only + vendored files; no other app file changed):**
- **ZXing** `@zxing/library` 0.23.0 (Apache-2.0, pure JS) vendored at `libs/zxing/zxing.min.js` (+LICENSE),
  lazy-loaded on first scan. Used when the native detector is missing or finds nothing. ZXing returns one
  code per decode, so `scanZxingDetect` does a TRY_HARDER pass on the whole image, then overlapping
  horizontal bands (7 and 12 bands, 50% overlap; GlobalHistogram then Hybrid binarizer) — finds every
  stacked 1D barcode. Image capped at 2000 px. ~0.4–1.5 s per photo.
- **Offline OCR on `file://`**: `tests/build-tesseract-offline.cjs` wraps the vendored worker, the two LSTM
  core variants and `eng.traineddata.gz` as plain scripts in `libs/tesseract/offline/` (classic `<script>`
  loads work on `file://`). On `file://`, `scanStartOCROffline` injects them, puts **core + worker source in
  one blob** (a blob worker can't `importScripts` a second blob on an opaque origin), passes the model in
  memory as `{code:"eng", data}`, picks SIMD/non-SIMD core by feature test, then frees the page copies.
  The build script also fixes one Tesseract.js 5.1.1 bug in the offline worker copy (initialize joined
  `t.data` instead of `t.code` for in-memory languages and hung); the original `worker.min.js` is untouched.
  On http(s) the standard loader is used, with the offline path as fallback. Every start is time-boxed (60 s).
- **OCR quality**: upscale images whose long side is < 1200 px to 1600 px; `recognize(…, {rotateAuto:true})`
  deskews; if no labelled serial is read, a second pass in sparse-text mode (PSM 11); both texts feed the vote.
- **Parser**: serial requires an explicit `SN` / `S/N` / `Serial [Number|No|#]` label not glued to a preceding
  letter/digit; all labelled hits collected, the one with the unit-serial shape wins.
- **UI**: `scanPump()` queue (one scan at a time, picks up rows added mid-run; Scan all / ↻ go through it);
  `#scan-engine-note` shows the engines actually in use and refreshes after each scan; counter/empty text update.

**Verified (headless Chromium, synthetic Lenovo-style labels in `tests/scan-label-*.{png,jpg}`):**

| | Before | After (`file://` and `http://`) |
|---|---|---|
| Clean label | serial via OCR (http only), 0 barcodes | serial + model, all 5 barcodes |
| Tilted, blurred phone shot | nothing | serial + model, 2 barcodes |
| Low-res tilted photo | nothing | serial + model |
| `file://` | hangs forever | 3 photos in ~6.6 s, offline engine |

`tests/scan-tests.cjs` 23/23 (parser + vote, incl. the MTM regression) · `tests/ui-scan.cjs` 11/11 on file:// and
http:// (zero console errors, queue test) · dashboard suites: harness 104/100, ui-smoke 71/71 (now layout-aware:
15 core nav + Inventory/Scan; Asset 360 is a detail page), ui-features 37/37.

**Size:** `libs/tesseract/offline/` adds ~10 MB, `libs/zxing/` 0.4 MB. Not used by the app at all (OEM 1 = LSTM-only
and the `.wasm.js` files inline their wasm): `tesseract-core.wasm(.js)`, `tesseract-core-simd.wasm(.js)`, the four
bare `.wasm` files and uncompressed `eng.traineddata` (~24 MB) — safe to delete if folder size matters (not done).

**Still true:** real phone photos vary — rows stay editable; don't auto-correct serial characters.
