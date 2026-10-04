# QBR Dashboard — Inventory & Scanner Rebuild Guide

**Base app:** QBR Dashboard v1.9.0 (offline-first, zero-build: plain HTML/CSS/JS, no bundler, no CDN, no fetch).
**What this documents:** how two features were added on top of v1.9.0, in enough detail for another AI to reimplement them from scratch:
1. **Inventory Management / Ticket Tracking** (new "Inventory" nav section)
2. **Barcode/Label Scanner** (new "Scan" page + scan buttons inside inventory forms)

**Delivery history** (zip versions; in-code constants are `QBR.INV_VERSION="1.0.0"`, `QBR.SCAN_VERSION="1.2.0"`):

| Delivery | Change |
|---|---|
| v1.0 (initial) | Inventory module + scanner integration built and verified |
| v1.1 | Bugfix: `window.APP` → guarded `typeof APP` check (see §5.1) |
| v1.2 | Scanner hardened against real label photos: barcode classifier (key/serial/other), MTM+serial split, consensus serial scoring, model fallback |
| v1.3 | Quick serial lookup ("Look up asset") on the Inventory page |
| v1.4 | Scanner fixes (scan.js 1.2.0): ZXing barcode decoding (native `BarcodeDetector` does not exist on Windows), fully offline OCR on `file://` (was hanging), deskew/upscale for tilted photos, serial-label regex fix, scan queue (see §3.3, §5.2) |

---

## 1. Host-app architecture you must know first

### 1.1 Globals and namespaces

- `window.QBR` — the shared namespace object. Every module attaches constants and functions to it (`QBR.invSerialKey`, `QBR.scanImage`, …). Modules are plain `<script>` tags; there is no import system.
- `const APP` — declared at top level in `js/app.js`. **Critical gotcha:** a top-level `const` does NOT become `window.APP`. Any module that needs the app model must use the guarded form:
  ```js
  function invModel() { return (typeof APP !== "undefined" && APP.model && APP.model.inventory) || null; }
  ```
  Never `window.APP` — it is always `undefined` (this exact bug shipped once; §5.1).
- `XLSX` — SheetJS global, already loaded by the host app. Reuse it for parsing and export.
- `esc()` — global HTML-escape helper. **Every user-derived string inserted into HTML must pass through it.**
- Display convention: missing values render as `—` or `No data`, never `0`.

### 1.2 Tab / panel system

- Sidebar buttons: `<button class="tab-btn" data-tab="dash-xyz" data-group="Inventory">`. The `data-group` attribute visually groups nav items.
- Panels: `<section id="dash-xyz" class="dash-panel d-none">`. Panels with `data-deck="skip"` are excluded from the printable deck.
- `goToTab("dash-xyz")` / `renderAll()` — global navigation and re-render entry points (defined in `app.js`; call them guarded: `if (typeof renderAll === "function") renderAll();`).
- `TAB_FILTERS` (in `app.js`) maps each tab id to the global filters that apply to it. The Inventory tab uses `[]` (it has its own in-panel filters).
- Nav badges: `app.js` computes a `[count, tone, label]` tuple per tab. The Inventory badge shows the **open ticket count**.

### 1.3 Model loading flow (where inventory hooks in)

When workbooks are uploaded, `app.js` runs the main M365 parser and then, in the same flow:

```js
try { APP.model.inventory = (typeof QBR.parseInventoryBuffers === "function")
        ? QBR.parseInventoryBuffers(buffers) : null; }
catch (e) { console.warn("[QBR] inventory parse failed:", e && e.message); APP.model.inventory = null; }
```

- `buffers` = the raw uploaded `ArrayBuffer`s. Inventory parsing is **filename-independent** and **additive**: every uploaded workbook is scanned for inventory sheets and merged.
- `renderInventory()` is called from the main render path (`if (typeof renderInventory === "function") renderInventory();`).
- Nothing is persisted: the model lives in memory for the session. (Session restore re-parses the saved workbook bytes.)

### 1.4 Conventions to preserve

- Zero-build, no CDN, no `fetch`. Everything must work from `file://` except where noted (§5.2).
- Pure logic (parsers, scoring) must be DOM-free and Node-testable; browser-only code goes in clearly-marked sections. Each module ends with a Node export guard (inert in browser) so logic can be unit-tested headless.
- Add features additively; do not modify existing v1.9.0 pages.

---

## 2. Feature 1 — Inventory Management / Ticket Tracking

### 2.1 Files

| File | Change |
|---|---|
| `js/inventory.js` | **New** (~900 lines). All inventory logic + UI. Loads after `excel-loader.js` (it borrows `makeResolver`), before `scan.js`. |
| `index.html` | Add nav buttons (`dash-inventory`, `dash-scan` group "Inventory"), panels (`dash-inventory`, `dash-asset360`, `dash-scan`), `<script src="js/inventory.js">`. |
| `js/app.js` | 4 integration points (§2.9). |

### 2.2 Public API added to `QBR`

```js
QBR.INV_VERSION = "1.0.0";
QBR.INV_THRESH  = { AGING_DAYS: 7, LEMON_TICKETS: 3, WARRANTY_WARN_DAYS: 90 };
QBR.INV_STATUS  = ["In Stock", "Deployed", "In Repair", "Retired"];

QBR.invSerialKey(s)        // canonical serial: uppercase, trim, collapse whitespace
QBR.invParseSheets(sheets) // {SheetName: rows[][]} -> accumulator (multi-workbook)
QBR.parseInventoryBuffers(buffers) // ArrayBuffer[] -> {assets, tickets, deployments, pipeline, pos, sources}
QBR.invAssetStatus(a, openByKey)    // derived status string
QBR.invOpenByKey(tickets)           // Map key -> open tickets[]
QBR.invComputeFlags(inv, today)     // {aging, lemons, warrantyExpiring, warrantyExpired, unmatched, stalled}
QBR.invJournal = []                 // audit trail entries {ts, action, detail}
QBR.invLog(action, detail)
QBR.invIntake(list)                 // [{sn, client, model, ...}] -> creates assets
QBR.invDeploy(keys, client, dateDelivered, sq)
QBR.invAddTicket(t)
QBR.invResolveTicket(key, n, note)
QBR.invSetStatus(key, status)       // sets a.statusOverride
QBR.invExportWorkbook()             // regenerates .xlsx download
QBR._invUI  = { client:"ALL", type:"ALL", status:"ALL", q:"", flag:null, showAll:false, form:null, prefillSn:null }
QBR._invA360 = null                 // key of the asset shown on the Asset 360 page
```

### 2.3 Data model

```js
asset      = { key, sn, client, model, desc, cat, brand, supplier, dr,
               delivered, wstart, wend, wyears, cond, contact, addr, phone,
               statusOverride }
ticket     = { _n, key, sn, client, model, cat, status, repEdtech, repLenovo,
               completed, pic, issue, act, svcAddr }
deployment = { key, sn, type, desc, sq, client, req, delivered, remarks }
pipeline   = { ...row fields incl. proposal, delivery ... }
po         = { ...purchase-order row fields... }
```

- `key = QBR.invSerialKey(sn)` is the **canonical join key** across all collections. Tickets link to assets by normalized serial — never by Excel lookup formulas.
- `SENTINELS = {"", "n/a", "na", "-", "—", "–", "none", "null"}` — values treated as empty.

### 2.4 Workbook parsing

**Sheet resolution** (`INV_SHEETS`): each logical sheet has a candidate name list, e.g. devices → `["02 DEVICES", "DEVICES"]`. Resolution tries exact names first, then a substring fallback (normalized lowercase, whitespace removed) mirroring `findSheet()` in `excel-loader.js`.

| Logical sheet | Candidates | Parser |
|---|---|---|
| devices | `02 DEVICES`, `DEVICES` | `invParseDeviceRows` |
| support | `03 TECH SUPPORT LOGS`, `TECH SUPPORT LOGS`, `SUPPORT LOGS` | `invParseSupportRows` |
| raksoinv | `04 RAKSO INV.`, `04 RAKSO INV`, `RAKSO INV`, … | `invParseRaksoRows` |
| pipeline | `06 PIPELINE`, `PIPELINE` | `invParsePipelineRows` |
| po | `07 PURCHASE ORDER`, `PURCHASE ORDER` | `invParsePORows` |

**Column resolution:** `makeResolver(headerRow)` (borrowed from `excel-loader.js`) maps a list of candidate header patterns to a column index by normalized substring match. Each parser declares its columns as pattern lists, e.g. devices: `sn: ["serial number", "serial"]`, `repEdtech: ["date reported to edtech", "reported to edtech"]`. Column order never matters; a missing serial column aborts that sheet's parser.

**Dates** (`invDate`): accepts Excel serial numbers and date strings; returns a `Date` or `null`.

**RAKSO sheet specifics** (the tricky one):
- Rows are grouped by SQ: a row with a usable SQ value starts a group; blank-SQ rows **forward-fill** (`fSq`, `fClient`, `fReq`, `fDel`, `fRem`).
- Three serial columns are **pivoted into one record per physical asset**: `s/n` → type Laptop, `desktop-s/n` → Desktop, `monitor-s/n` → Monitor, each with its own description column.

**Assembly** (`QBR.parseInventoryBuffers`): reads each buffer with `XLSX.read(buf, {type:"array"})`, converts every sheet with `sheet_to_json(..., {header:1, defval:null, blankrows:false})`, accumulates across workbooks, then de-dupes assets by key (newest file wins) and assigns stable ticket ids (`t._n = index`, used by resolve buttons).

### 2.5 Derived status & flags

`QBR.invAssetStatus(a, openByKey)` precedence:
1. `a.statusOverride` (manual, from write-back) wins
2. condition mentions retired/disposed/beyond repair → `Retired`
3. condition mentions repair, OR asset has an open ticket → `In Repair`
4. has client + delivered date → `Deployed`
5. else → `In Stock`

`QBR.invComputeFlags(inv, today)` — six flags, thresholds from `QBR.INV_THRESH`:
- `aging`: open ticket with `repEdtech` older than 7 days
- `lemons`: asset with ≥ 3 tickets (must exist in registry)
- `warrantyExpiring`: `0 ≤ (wend − today) ≤ 90` days, not retired
- `warrantyExpired`: `wend < today`, not retired
- `unmatched`: ticket serials with no asset in the registry (deduped)
- `stalled`: pipeline rows where proposal matches /signed/i but delivery doesn't match /delivered/i

### 2.6 Inventory page UI (`renderInventory`)

Rendered into `#inv-body`. Sections, top to bottom:
1. **KPI row**: Fleet size · Deployed % · In Repair · Open tickets · Avg days to resolve (missing data → `—`).
2. **Flag cards** (6, clickable): each sets `ui.flag`, which filters the tables below; a "Clear flag filter ✕" button resets.
3. **Action buttons**: Register assets (intake) · Deploy units · New ticket · Export inventory workbook — plus the **Look up asset** row (v1.3, §2.8).
4. **Filters**: Client / Type / Status dropdowns + "Search serial" text filter (substring on normalized key).
5. **Tables**: Assets (Serial, Client, Model, Type, Status pill, Warranty end, Tickets count — first 150 rows + "Show all N"), Support tickets (with Resolve buttons), Pipeline, Purchase orders.
6. Clicking an asset row opens its **Asset 360** page.

**Asset 360** (`openAsset360(key)`, panel `dash-asset360`, `data-deck="skip"`): per-serial dossier — asset details, deployment history, all tickets, warranty state. Has its own fallback serial search when opened with no asset selected.

### 2.7 Write-back (in-memory + journal + export)

All mutations update the in-memory model, append to `QBR.invJournal`, and re-render. A browser cannot silently overwrite the user's local file, so changes are exported as a **regenerated workbook**:

- `QBR.invIntake(list)` — creates asset records (defaults: cat "Laptop", brand "Lenovo", cond "No Issue").
- `QBR.invDeploy(keys, client, dateDelivered, sq)` — sets client/delivered/sq on assets (clears `statusOverride`).
- `QBR.invAddTicket(t)` / `QBR.invResolveTicket(key, n, note)` — ticket lifecycle; resolve prompts for an optional note.
- `QBR.invSetStatus(key, status)` — manual override (`In Stock/Deployed/In Repair/Retired`).
- `QBR.invExportWorkbook()` — rebuilds `02 DEVICES`, `03 TECH SUPPORT LOGS`, `04 RAKSO INV.`, `06 PIPELINE`, `07 PURCHASE ORDER` from the model plus an `EDIT LOG` sheet (timestamp, action, detail per journal entry); auto-sizes columns; triggers download via SheetJS `XLSX.writeFile`.

### 2.8 Quick serial lookup (v1.3)

`invLookupSerial()` + `#inv-lookup` input / `#inv-lookup-go` button in the actions area, wired in `invBind()` (Enter key supported):
- exact key match → `openAsset360(key)`
- single substring match → `openAsset360`
- multiple matches → sets `ui.q`, re-renders, scrolls to the Assets filter
- no match → inline message with a "Register it via intake" button that opens the intake form with the serial prefilled in `#in-serials`

### 2.9 `app.js` integration points (exact)

1. **Model**: in the workbook-load flow —
   `APP.model.inventory = QBR.parseInventoryBuffers(buffers)` (guarded try/catch, null on failure).
2. **Badge**: `dash-inventory` badge tuple counts open tickets (`!completed` and status not matching /completed|resolved|closed/i).
3. **Render**: `if (typeof renderInventory === "function") renderInventory();` in the main render path.
4. **Filters**: `TAB_FILTERS["dash-inventory"] = []` (panel owns its filters).

### 2.10 `index.html` integration points

- Nav: two `data-tab` buttons (`dash-inventory`, `dash-scan`) with `data-group="Inventory"`, using the `#fi-storage` / `#fi-scan` SVG sprite icons.
- Panels: `<section id="dash-inventory" class="dash-panel d-none">` (with `#inv-body`), `<section id="dash-asset360" class="dash-panel d-none" data-deck="skip">`, `<section id="dash-scan" class="dash-panel d-none" data-deck="skip">`.
- Scripts (order matters): `js/inventory.js`, then `libs/tesseract/tesseract.min.js`, then `js/scan.js`.

---

## 3. Feature 2 — Barcode / Label Scanner

### 3.1 Files

| File | Change |
|---|---|
| `js/scan.js` | **New** (~600 lines). Detection pipeline, parsers, Scan page UI, routing. Loads after `inventory.js` + Tesseract. |
| `libs/tesseract/` | **Vendored** Tesseract.js v5.1.1 (Apache-2.0): `tesseract.min.js`, `worker.min.js`, `tesseract-core*.wasm(.js)` variants, `eng.traineddata.gz`. Only the two **LSTM `.wasm.js`** cores and the `.gz` model are actually used (OEM 1; `.wasm.js` inlines its wasm). |
| `libs/tesseract/offline/` | **Generated** by `tests/build-tesseract-offline.cjs`: worker (with one bug fix), both LSTM cores and the model wrapped as plain `<script>` files, for OCR on `file://` (~10 MB). |
| `libs/zxing/` | **Vendored** `@zxing/library` 0.23.0 UMD (Apache-2.0, pure JS, 0.4 MB) + LICENSE — barcode decoding where the native detector is missing (Windows/Linux). Lazy-loaded. |
| `index.html` | Tesseract script tag + Scan nav/panel (see §2.10). |

No CDN anywhere. Barcodes: the browser's **native** `BarcodeDetector` where it exists (macOS / ChromeOS / Android only — **not Windows or Linux**), otherwise the vendored **ZXing**. OCR: the vendored Tesseract. (The original standalone's "ZXing" script tags pointed at unrelated npm packages and were dropped; v1.4 vendors the real `@zxing/library`.)

### 3.2 Entry points (user decision)

- **(A)** Dedicated **Scan** page (`dash-scan`) under the Inventory nav group.
- **(B)** Scan buttons inside the inventory forms: "Scan box label" in Intake, "Scan unit" in Deployment. These call `QBR.scanSetTarget("intake"|"deploy")`, which routes through the Scan page in target mode; after a successful scan the results prefill the originating form.

### 3.3 Detection pipeline (`QBR.scanImage(dataUrl)`)

```
photo upload (no live camera, per user decision)
  → scanDetectBarcodes(dataUrl)      # native BarcodeDetector (13 formats) + hard-threshold retry
      → if no native detector or zero results: scanZxingDetect — TRY_HARDER pass on the whole
        image (≤2000 px), then overlapping horizontal bands (7 and 12, 50% overlap) so each
        stacked 1D barcode is decoded (ZXing returns one code per decode)
  → scanRunOCR(dataUrl)              # Tesseract; upscale <1200 px images to 1600, rotateAuto deskew,
                                     # PSM 11 retry when no labelled serial; {text, ocrOk}
  → scanMergeFindings(barcodes, scanParseOCRText(ocr.text), ocr.text)
  → { barcodes, serial, productKey, model, ocrOk, barcodeOk }
```

**Tesseract worker** (`scanGetOCRWorker`): single reused worker, two start paths:
```js
// http(s): standard loader
Tesseract.createWorker("eng", 1, { workerPath: "libs/tesseract/worker.min.js",
  corePath: "libs/tesseract", langPath: "libs/tesseract", cacheMethod: "none", logger(){} });
// file:// (and fallback): inject libs/tesseract/offline/{worker,core-(simd-)lstm,eng}.js, then
const blobUrl = URL.createObjectURL(new Blob([coreSrc, "\n;\n", workerSrc]));   // ONE blob
Tesseract.createWorker([{ code: "eng", data: engBytes }], 1, {
  workerBlobURL: false, workerPath: blobUrl, corePath: "inline.js", cacheMethod: "none" });
```
- Lazy, cached in `QBR._scanOCR = {worker, failed, starting, mode, reason}`; each start time-boxed (60 s).
- Why one blob: on `file://` the page origin is opaque, and a blob worker cannot `importScripts` a second blob URL. With `TesseractCore` already defined, the worker skips its own core loader.
- Tesseract.js 5.1.1 bug: with in-memory languages the worker initializes with `t.data` (the bytes) instead of `t.code` and hangs. The build script patches that single expression in the offline copy only, and asserts it matched exactly once.
- **Failure is non-fatal**: barcode-only with the reason shown in `#scan-engine-note`. OCR has a 150 s timeout race.
- Photos stay in memory; nothing is persisted.

**Scan page UI**: upload area → per-image rows in an editable table (Image · Raw barcodes · Serial · Model · actions). Rows are editable; each row offers **Add to intake** (prefill), **Tag to client** (dialog), and per the scan decision, existing serials warn first.

### 3.4 OCR text parsing (`scanParseOCRText`, pure/Node-testable)

Input: raw OCR text. Output: `{serial, productKey, model}`.

1. **Product key** — first match of:
   - `\b([A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5})\b`
   - `\b(\d{5}-\d{5}-\d{4})\b`, `\b(\d{5}-\d{3}-\d{7}-\d{5})\b`
2. **Serial** — explicit label required, not glued to a preceding letter/digit (covers `SN:`, `S/N:`, `(S)SN:`, `Serial Number/No/#`); all hits collected, the one with the unit-serial shape wins:
   - `/(?:^|[^A-Z0-9])(?:S\s*\/\s*N|SN|SERIAL(?:\s*(?:NUMBER|NO\.?|#))?)\s*[:.=#]?\s*([A-Z0-9][A-Z0-9\-]{5,25})/gi`
   - (v1.3 and earlier accepted a bare `S`, which pulled `5VD00…` out of `MTM: 21M4S5VD00`.)
3. **Model** — first brand-list hit (`ThinkPad|ThinkCentre|…|Latitude|…|EliteBook|…`) followed by the model fragment; **fallback** (v1.2): if the brand word is OCR-mangled (real case: "ThinkPad" → "Milian"), capture the model-number fragment `/\b([A-Z]{1,4}\d{2,4}[A-Z]?\s*Gen\s*\d)/i` → e.g. `E14 Gen 6`.

### 3.5 Barcode classification & consensus (v1.2, pure/Node-testable)

Lenovo box labels carry ~10 barcodes (serial, MTM, MTM+serial concatenated, UUID, UPC, 2× MAC, Product Key ID, battery). "First barcode wins" is unsafe, so:

- `scanClassifyBarcode(v)` → `"key"` (Windows-key shapes), `"other"` (UUID, 8-digit EAN-8, 12–14-digit UPC/EAN digit strings, 12-hex MAC-like values — never the unit serial), or `"serial"`.
- `scanSplitMtmSn(v)` — `/^([A-Z0-9]{10,12})([A-Z0-9]{8})$/` splits Lenovo's `(1S)` MTM+serial concatenation so the trailing 8-char serial votes on its own.
- `scanPickSerial(candidates, ocrText)` — consensus scoring per unique value: **+3 per occurrence vote, +3 for Lenovo serial shape** (`^(?=.*[A-Z])(?=.*\d)[A-Z0-9]{7,10}$`), **+4 if present in the OCR text**; highest wins, ties broken by first occurrence.
- `scanMergeFindings(barcodes, parsed, ocrText)` — collects serial candidates (barcode serials incl. MTM tails + OCR serial), picks by consensus; product key prefers barcode classification, falls back to OCR; strips a leading `S/N` prefix from the serial.

### 3.6 Post-scan routing (user decisions)

Serials are normalized with `QBR.invSerialKey()` before any lookup.

- **Existing serial → warn FIRST, then jump** to its Asset 360 page (`scanWarnAndJump`: toast, then `openAsset360` after ~800 ms).
- **New serial → "Add to intake"** prefills the intake form (`QBR._invUI.form="intake"`, serial into `#in-serials`); **"Tag to client"** opens the tag dialog which creates the asset + deployment record + journal entry. Tag dialog fields: **Client\*, Date\*, Purchase location, SQ number, DR #, Assigned to**.
- In target mode (launched from Intake/Deployment forms), the scan result prefills the originating form instead.

### 3.7 UI state

`QBR._scanUI = { rows: [], nextId: 1, processing: false, target: null }`, `QBR.scanSetTarget(mode)`. Scans run **sequentially through `scanPump()`**: one at a time, in order, until no row is "pending" — so photos added mid-run, "Scan all" and ↻ all go through the same queue (the parallel `forEach` version only ever scanned the first image; the v1.3 per-batch loop left mid-run additions stuck on "Pending").

---

## 4. Rebuild checklist (ordered)

**Inventory**
1. Create `js/inventory.js`: constants → serial/date/sentinel utils → sheet/column resolvers → 5 row parsers → `invParseSheets`/`parseInventoryBuffers` → status/flags → journal + mutations → export → UI state → `renderInventory` + Asset 360 + lookup → `invBind` → Node guard.
2. `index.html`: nav buttons, 3 panels, script tag (after `excel-loader.js`).
3. `app.js`: model assignment, badge, render hook, `TAB_FILTERS`.
4. Test: synthetic workbook (600 assets / 85 tickets / deployments / pipeline / POs) asserting counts; flag/mutation assertions; `node --check`.

**Scanner**
5. Vendor Tesseract.js v5.1.1 files under `libs/tesseract/` (script tag **before** `js/scan.js`), run `node tests/build-tesseract-offline.cjs` to generate `libs/tesseract/offline/`, and vendor `@zxing/library` UMD as `libs/zxing/zxing.min.js`.
6. Create `js/scan.js`: pure parsers (`scanParseOCRText`, classifier, MTM split, consensus) → browser detection (`scanDetectBarcodes` + contrast fallback, `scanGetOCRWorker` + `scanRunOCR`) → `QBR.scanImage` → Scan page UI → routing (`scanRoute`, warn-then-jump, prefill intake/deploy, tag dialog).
7. `index.html`: Scan nav button + `dash-scan` panel.
8. Inventory forms: add "Scan box label"/"Scan unit" buttons calling `QBR.scanSetTarget`.
9. Test: `node tests/scan-tests.cjs` (parsers/classifier/consensus incl. the MTM regression); `tests/ui-scan.cjs` (Playwright) on **both `file://` and `http://`**: synthetic clean / tilted-phone / low-res labels must all yield the serial, with no hang and zero console errors; real label photos (expect ~9/10 serials exact; rows stay editable).

---

## 5. Gotchas & lessons learned

### 5.1 `window.APP` is always undefined
`app.js` declares top-level `const APP`; `const` never attaches to `window`. The first scanner build checked `window.APP` and silently saw no model while the badge (using bare `APP`) worked. Rule: cross-module model access always uses `typeof APP !== "undefined" && APP.model …`.

### 5.2 `file://` and Windows — both solved in v1.4
- **OCR on `file://`:** the standard Tesseract loader fetches worker/core/model files, which browsers block on `file://`, and the v1.3 code **hung** rather than failing. v1.4 loads the engine from generated plain scripts and starts it from a blob (§3.3). Over http(s) the standard loader is used.
- **Barcodes on Windows:** `BarcodeDetector` exists only on macOS, ChromeOS and Android. ZXing (pure JS) covers Windows/Linux and works on `file://`.
- Lesson: wrap every engine start in a timeout; a promise that never settles is worse than an error.

### 5.3 Scan quality beats detector choice
On real CamScanner label photos the 1D barcodes were undecodable even by ZXing (thin bars destroyed by downsampling) while the printed text OCR'd cleanly. Design for OCR-first on labels; keep barcode detection as the fast path for clean shots.

### 5.4 Real-world OCR confusions to expect
`ThinkPad`→`Milian`, `6`→`e`, `J`→`U` (tilted photos), `Gen6` without space. The model fallback and editable rows exist because of these. Don't auto-"correct" serial characters — surface the guess and let the human confirm.

### 5.5 Workbook parsing robustness
Sheets by name (with substring fallback), columns by header pattern (order-free), serials by normalized key, RAKSO continuation rows by SQ forward-fill, one record per physical serial. These five rules are what make the parser survive real-world workbook edits.

### 5.6 Write-back honesty
A browser page cannot silently overwrite the user's local `.xlsm`. The design journals every mutation in memory and regenerates the workbook on Export. If multi-user or persistence is ever needed, that's a backend project — not a frontend patch.
