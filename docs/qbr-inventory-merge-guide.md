# QBR Dashboard — Inventory Module Merge Guide
**Delivery version:** v1.21.0 · `QBR.INV_VERSION = "1.21.0"` · 2026-10-05

This document lists every **working** change in the Inventory module so it can be merged
into the live QBR dashboard file. Superseded patches and dead-end experiments are
excluded (see §9). The host app shell is v1.9.0; `QBR.VERSION` and `QBR.SCAN_VERSION`
are independent of `QBR.INV_VERSION`.

Working copy: `~/workspace/qbr-inventory-dev/qbr-app`
Original: `~/workspace/user/files/qbr-app.zip`
Engineering guide: `~/workspace/user/files/QBR_Dashboard.md`

---

## 1. New report sections → modules (the map you asked for)

Every new report section in the dashboard, and the JS modules that power it:

| Report section | Tab id | Entry function | Modules (copy these files) |
|---|---|---|---|
| **Inventory** — Assets sub-page (KPIs, unit-type tiles, bundle strip, flags, asset table, intake/deployment forms, Search serial, Look up asset) | `dash-inventory` | `renderInventory()` in `js/inventory.js` | `js/inventory.js` |
| **Inventory** — Support tickets sub-page (ticket table, new-ticket form, Import Forms file) | `dash-inventory` | `renderInventory()` (view=`"tickets"`) | `js/inventory.js` |
| **Inventory** — Supplies sub-page (printer consumables: items, transactions, summary, flags) | `dash-inventory` | `QBR.supRender()` | `js/inventory.js` + `js/supplies.js` |
| **Asset 360** — single-unit profile (detail, ✏️ Edit, Lenovo warranty link, ticket history) | `dash-asset360` | `renderAsset360Panel()` in `js/inventory.js` | `js/inventory.js` |
| **Ticket 360** — single-ticket detail (notes timeline, status, related tickets) | `dash-ticket360` | `renderTicket360Panel()` in `js/inventory.js` | `js/inventory.js` |
| **Scan** — barcode/OCR label scanner page + scan buttons in Intake/Deployment | `dash-scan` | `js/scan.js` (own page renderer; plugs into intake/deploy via `QBR.scanSetTarget()`) | `js/scan.js` + `libs/tesseract/**` + `libs/zxing/**` |
| **School 360 → Hardware** — Hardware KPI tile + Hardware section on the school page | (extension of existing school page) | `s360Hardware()` in `js/app.js` | `js/app.js` (patch only) |

Shared plumbing every section above depends on:
- `js/persist.js` — journal persistence (refresh-proof), File System Access "Link file / Save to Excel"
- `js/patch.js` — patch-in-place saver (writes journal ops as targeted cell edits into the original workbook)
- `css/styles.css` — inventory styles (autocomplete dropdown, sticky Asset 360 bar, flag cards, tile styles)

Do **not** copy `tests/`, `VERSIONING.md`, or `~/workspace/qbr-inventory-dev/` scratch files into the live app —
they are dev-only (keep them alongside your repo for future work).

---

## 2. Merge checklist (files)

### 2a. New files — copy as-is
- `js/inventory.js` — the entire Inventory module (Assets, tickets, Ticket 360, Asset 360, Forms import)
- `js/supplies.js` — printer consumables (Supplies sub-page)
- `js/scan.js` — offline scanner page + intake/deploy scan integration
- `js/patch.js` — patch-in-place workbook saver
- `js/persist.js` — journal persistence + direct file save (File System Access)
- `libs/tesseract/**` — offline OCR engine + `eng.traineddata` (required by `scan.js` over `file://`)
- `libs/zxing/**` — vendored ZXing 0.23.0 (Windows/Linux barcode fallback)
- `tests/**` — 9 Node test files, 199 assertions (dev-only, not shipped to users)

### 2b. `index.html` — merge these blocks
1. **Sidebar nav** — the `Inventory` group buttons (`data-tab="dash-inventory"` and `data-tab="dash-scan"`)
   around line 129–131 of the dev copy.
2. **Panels** — the `<section>` blocks with ids `dash-inventory` (`#inv-body`),
   `dash-asset360` (`#a360-body`), `dash-ticket360`, and `dash-scan` (`#scan-body`)
   around line 585–607. Note `data-deck="skip"` on Asset 360 / Ticket 360 / Scan
   (keeps them out of the PPT deck and gallery export).
3. **Script tags**, in this order, after the existing generators and before/around `js/app.js`:
   ```html
   <script src="js/inventory.js"></script>
   <script src="js/supplies.js"></script>
   <script src="libs/tesseract/tesseract.min.js"></script>
   <script src="js/scan.js"></script>
   <script src="js/app.js"></script>
   <script src="js/patch.js"></script>    <!-- v1.15: patch-in-place workbook saver -->
   <script src="js/persist.js"></script>  <!-- v1.8: journal persistence + direct file save -->
   ```
   (`js/shell.js` stays last if present.)

### 2c. `js/app.js` — merge these integration points
- Tab registry: `"dash-inventory"`, `"dash-asset360"`, `"dash-ticket360"`, `"dash-scan"` entries
  (tab open/close logic, filter scoping — asset360/ticket360/scan declare empty shared-filter lists).
- `#asset/<serial-key>` and `#ticket/<tno>` deep-link routing (opens full dashboard in new tab).
- `goToTab("dash-inventory")` call sites (e.g. from Inventory action buttons).
- `s360Hardware()` — School 360 hardware tile + section (v1.20.0), including the alias-Set
  normalization fix (v1.20.1).
- Any `QBR.renderAll()` / render-dispatch wiring that calls `renderInventory()`.

### 2d. `css/styles.css` — merge the inventory additions
Autocomplete dropdown (serial lookup), sticky Asset 360 top bar, flag cards
(`Duplicate serials` red card), KPI tile styles, hardware tile, import-preview panel,
tag-batch dialog styles.

---

## 3. Working features by version

### v1.0 — Inventory module (base)
- New **Inventory** nav section: asset registry (laptops/desktops/monitors), KPIs,
  6-flag monitoring engine (aging tickets, lemons, warranty <90d, unmatched serials, stalled pipeline, …).
- Intake → deployment → support ticket flow.
- **Asset 360** per-serial detail pages.
- Edit journal + workbook export.

### v1.2 — Offline scanner
- Dedicated **Scan** page + scan buttons in Intake/Deployment forms; photo upload.
- Fully offline via vendored Tesseract.js (works over `file://`).
- Duplicate serial → warns, then opens Asset 360.
- Tag mode captures Client, Date, Purchase location, SQ number, DR #, Assigned to.
- OCR verified on 10 supplied Lenovo label photos (9/10 serials exact).

### v1.2.0 (scanner rebuild, per SCANNER_FIX.md)
- Vendored ZXing 0.23.0 build for Windows/Linux where native `BarcodeDetector` is missing.
- Offline Tesseract loader (engine/model as plain scripts, worker from a single blob).
- Image upscaling + auto-rotation before OCR; tightened serial regex requiring a real SN label.
- Scan-page queue/counter/status fixes. Real decode verified (PF62SDPW synthetic label; 4/4 barcodes via banded scan).

### v1.3 — Serial quick-lookup
- "Look up asset" under Inventory action buttons: Enter/Open jumps to Asset 360;
  no match offers intake with the serial pre-filled.

### v1.4 — Printer supplies (consumables)
- New `js/supplies.js`; Assets | Supplies toggle; quantity-based model
  (Items/Transactions sheets, JS-recomputed summary, restock/negative/unknown-item flags,
  record-transaction/add-item write-back, formula-preserving export).
- Verified against `ETG_PRINTER_INVENTORY.xlsx`.

### v1.8 — Persistence
- Refresh-proof **entry journal**: mutations recorded as replayable ops in
  `localStorage` (`qbr-inv-journal-v1`, keyed by file fingerprint), replayed after every
  parse incl. session restore.
- **Direct Excel save** via File System Access API (Chrome/Edge): "🔗 Link file" →
  per-file permission → "💾 Save to Excel" writes straight into the linked workbook;
  conflict confirm on external edits. Other browsers fall back to Export download.

### v1.9.3 — Versioning
- Semver adopted (MAJOR=breaking, MINOR=feature, PATCH=bugfix); zip named
  `qbr-inventory-app-<ver>.zip`. `QBR.INV_VERSION` badge atop the Inventory page;
  `QBR.VERSION` (host shell) and `QBR.SCAN_VERSION` stay independent.

### v1.10.0 — Scan photo dialog
- Clicking a Scan-page photo thumbnail opens a two-pane dialog: label photo left,
  manual serial/model entry form right (serial auto-focused), Save + Rescan buttons.
  Enter saves normalized to uppercase; Escape/outside-click dismisses.

### v1.12.0 — Scan batch tagging
- Checkbox column + **Tag batch** button on the Scan page: one dialog tags many serials to a
  single school/client, date, SQ, DR#, owner — skips already-registered and blank rows
  with a summary toast.

### v1.13.0 — Ticket 360 overhaul
- Fixed "New ticket" freeze (double-bound buttons).
- Ticket numbers `{SERIAL}-{YYYYMMDD}` with collision suffixes.
- 5 statuses: Open / In Progress / Waiting for parts / Escalated / Resolved.
- Priority + requester fields; follow-up notes timeline; related-ticket linking.
- `#ticket/<tno>` and `#asset/<serial-key>` deep links (open full dashboard in new tabs).
- Ticket table oldest-first with serial search; workbook round-trip for all new fields.

### v1.14.0 — Inventory summary overhaul
- **Unit-type tiles** (clickable, In Stock headline + Deployed subtext) replaced 28 per-model tiles.
- Bundle strip: deployable desktop sets = min(systems, monitors) with shortfall callout.
- Per-model breakdown now an expandable drill-down.
- Tickets moved to their own sub-page (Assets | Support tickets | Supplies) with
  oldest/newest sort toggle + search; Assets/Pipeline/PO cards collapsible.

### v1.15.0 — Patch-in-place save
- Linked-file saves replay the mutation journal as **targeted cell writes into the original
  workbook** (`js/patch.js`): formulas/layout/helper columns preserved, `=CONCAT` pattern
  extended for new Items rows, summary ranges bumped, Excel recalculates on open;
  `cellStyles:true` on read/write preserves column widths.
- Rebuild-download kept for unlinked/export.
- ⚠️ Known limitation: the free SheetJS build cannot serialize cell *styling*
  (bold/colors/borders flatten on any save); formulas and layout are preserved.

### v1.15.1 — Link-crash fix
- "Could not link the file: `d.getFullYear` is not a function" — journal-replayed supplies
  transactions carried ISO-string dates (JSON round-trip). Now normalized via `supDate`;
  `invFmtDate` defensively coerces non-Dates.

### v1.16.0 — Lenovo warranty links
- "🔍 Lenovo warranty lookup" link in the Inventory header (`pcsupport.lenovo.com/us/en/warranty-lookup#/`
  via `QBR.WARRANTY_LOOKUP_URL`); per-serial "Check on Lenovo ↗" on Asset 360.
- No official serial-prefill URL exists; full auto-fetch isn't feasible keyless
  (CORS-blocked from `file://`; official API needs enterprise ClientID) — deep links are the
  ToS-clean maximum.

### v1.17.0 — Usability
- All warranty links go to the generic US lookup page (per-model deep URLs removed — they
  failed to resolve).
- Serial quick-lookup shows up to 8 partial matches while typing (click opens Asset 360).
- Client filter: searchable text field with suggestion dropdown, partial case-insensitive
  match; blank = all clients.
- Asset 360: sticky "← Back to Inventory" top bar (bottom button kept).

### v1.17.2 — Search/render fix
- "Search serial" table filter was dead: asset-table row wrappers are `{a, st}` but the
  serial/flag filters read `x.key` (undefined) — now `x.a.key`. Covered by
  `tests/inv-render.test.js`, which runs the real render.

### v1.18.0 — Asset 360 ✏️ Edit
- Edit any unit's details inline: client, model, description, category, brand, supplier,
  DR #, delivered/warranty dates, warranty years, condition, contact, address, contact details.
- Serial locked as the canonical join key. Save diffs changed fields only via
  `QBR.invUpdateAsset` (journaled; replayed as targeted cell writes by `patchInvUpdateAsset`).

### v1.19.0 — Duplicate serials flag
- Red **Duplicate serials** flag card: counts normalized serials on 2+ asset rows;
  clicking filters the asset table to those rows.

### v1.19.1 — KPI rename
- **Fleet size** → **Total Inventory**.

### v1.20.0 — School 360 hardware
- Per approved spec: All-tenant-Status/Risky Users Report school naming is source of truth.
- Counts all laptops/desktops ever linked to the school (any unit status); monitors ignored
  in the School 360 overview.
- One **Hardware** KPI tile (laptops · desktops · open tickets) + **Hardware** section
  (laptops, desktops, total units purchased, open tickets, warranties expiring <90d).
- Tile and section hidden entirely when the school has no hardware data.
- "Open page ›" jumps to Inventory pre-filtered to the school.
- Matching case-insensitive incl. known school aliases.

### v1.20.1 — School 360 upload crash fix
- `aliases.map is not a function`: the loader stores master school aliases as a `Set`;
  `s360Hardware()` now normalizes array/Set/single-value aliases (realm-safe duck typing).

### v1.21.0 — MS Forms ticket import
- **Import Forms file** button on Support tickets; preview-then-confirm panel for MS Forms
  exports (`INTERNAL__LENOVO_TICKET_SUPPORT*.xlsx`).
- Maps: SN/model/category/priority direct; Full Name → Requested By; detailed issue +
  first-observed date + usable Yes/No → Issue/s; troubleshooting → Activities;
  Completion time → Date Reported; work email, phone, Forms ID, attachment link → notes.
  Client/School auto-matched from inventory by serial (blank if unknown). New tickets start Open.
- Ticket numbers `{SERIAL}-{YYYYMMDD}`; created via `invAddTicket` (journaled → patch-in-place save).
- **Dedup**: skips rows whose Forms response ID was already imported
  (`localStorage` `qbr-forms-import-v1` + ID stamped in ticket notes so the workbook
  remembers across browsers), or whose full-row fingerprint matches — including repeats
  within the same file. Re-importing the same export yields "Nothing new to import".

---

## 4. Data model & persistence

### Workbook sheets consumed
- **Assets** (inventory registry: serial, model, client, status, dates, warranty, condition, …)
- **Support tickets** (ticket no, serial, status, priority, requester, issues, notes/timeline, …)
- **Supplies**: `Items` + `Transactions` (+ JS-recomputed summary)

### Journal (refresh-proof persistence)
- Mutations recorded as replayable ops in `localStorage` key `qbr-inv-journal-v1`
  (keyed by file fingerprint), replayed after every parse including session restore.
- API surface: `QBR.journalRecord`, `QBR.invJournal`; domain ops
  `invIntake`, `invDeploy`, `invAddTicket`, `invUpdateTicket`, `invSetTicketStatus`,
  `invLinkTicket`, `invUpdateAsset`, `supRecordTransaction`, etc.

### Save paths
- **Linked file** (Chrome/Edge): `patchWorkbookFromJournal()` (`js/patch.js`) applies the
  journal as targeted cell writes into the **original** workbook object — formulas, layout,
  helper columns, `=CONCAT` extension, summary range bumps, and column widths survive.
- **Unlinked / export**: rebuild-download via `QBR.invBuildWorkbook` / `QBR.invExportWorkbook`.
- Missing values render as `—` / `No data` (never `0`); user-derived HTML is escaped.

### Other localStorage keys
- `qbr-inv-journal-v1` — mutation journal
- `qbr-forms-import-v1` — Forms import dedup memory (response IDs + row fingerprints)

---

## 5. Deep links (inventory contract — keep as specified)
- Ticket 360 opens in a **new tab with the full dashboard**: `#ticket/<ticket-number>`
- Asset 360 opens in a **new tab with the full dashboard**: `#asset/<serial-key>`

---

## 6. Test suite (dev-only)
Run: `cd ~/workspace/qbr-inventory-dev && node --test tests/`
**199 passed, 0 failed** at v1.21.0.
- `tests/forms-import.test.js` — 27 assertions (Forms parse/map/dedup)
- `tests/inv-render.test.js` — real-render regression (v1.17.2)
- `tests/inv-ux-1.17.test.js`, `tests/inv-ux-behavior.test.js` — 48 assertions (v1.17.0)
- `tests/patch-in-place.test.js`, `tests/patch-edge.test.js` — 52 assertions (v1.15.0)
- plus scanner parser tests (24) and earlier inventory tests.

---

## 7. Working copy vs live file
- The live file (Azure Static Web Apps, GitHub `Hayleylujah12/QBR-Dashboard`) is the v1.9.0
  base plus whatever has been merged so far.
- Merge = §2 checklist (new files + index.html blocks + app.js integration points + CSS).
  The Inventory module is self-contained behind `window.QBR` and the four panel ids; it does
  not modify host data logic (host v1.9.0 `js/app.js` except the listed integration points).
- Pedro's decision (2026-10-04): **the SharePoint part will not be implemented** — the app
  stays local/browser-based (Excel loading, local persistence, direct file save, export).

---

## 8. Known limitations (working as designed, be aware when merging)
1. **Cell styling flattens on save** — the free SheetJS build cannot serialize bold/colors/borders;
   formulas, layout, helper columns, and column widths are preserved. (ExcelJS writer was offered
   as a follow-up; not built.)
2. **No serial pre-fill on Lenovo warranty lookup** — no official URL supports it.
3. **Forms-import fingerprint memory is browser-local** — cross-browser dedup relies on the
   Forms ID stamped in ticket notes, which requires the tickets to be loaded.
4. Live browser testing of the BarcodeDetector/Tesseract extraction path is still an open item
   (Playwright e2e was written but not run — Chromium download was throttled).

## 9. Excluded from this guide (not working / superseded — do not merge)
- Per-model Lenovo warranty deep URLs (removed in v1.17.0; they failed to resolve).
- v1.17.1's `change`-listener search fix (superseded by the v1.17.2 root-cause fix).
- Preserved-but-untouched files per Pedro's constraint: his original scanner/OCR files
  (`~/workspace/user/files/Working_SCANNER_qbr-app_0_qtvi.zip`, `qbr-app/js/scan.js`,
  `qbr-app/libs/tesseract/offline/*.js`) — kept unless he approves replacement.
