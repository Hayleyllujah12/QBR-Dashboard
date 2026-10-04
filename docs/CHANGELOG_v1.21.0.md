# QBR Dashboard — Release v1.21.0 (Inventory module) · 2026-10-05

**Host app:** v1.9.0 (unchanged) · **Inventory:** `QBR.INV_VERSION` 1.0.0 → **1.21.0** · **Scanner:** `QBR.SCAN_VERSION` 1.2.0 (fixes kept) ·
**New:** `QBR.SUP_VERSION` 1.0.0, `QBR.PERSIST_VERSION` 1.8.0
**Release:** promoted to live on 2026-10-05 as **app version 1.21.0** (`QBR.VERSION` 1.9.0 → 1.21.0, aligned with the Inventory
number) in the new folder `QBR-Dashboard_1.21.0`; **`QBR-Dashboard_1.9.0` is the rollback** (banner only, code untouched).
**Compared against:** the previous live merge (2026-10-04: Inventory 1.0.0 + Scanner 1.2.0, repo commit `16af281`).
**Feature-by-feature history:** `docs/qbr-inventory-merge-guide.md` (supplied with this release).

---

## 1. What changed, in short

| Area | Change |
|---|---|
| **Inventory → Assets** | Unit-type tiles (In Stock / Deployed), desktop-bundle strip, per-model drill-down, "Total Inventory" KPI (was "Fleet size"), **Duplicate serials** flag, searchable client filter, serial lookup with live suggestions, collapsible cards |
| **Inventory → Support tickets** (own sub-page) | Ticket numbers `{SERIAL}-{YYYYMMDD}`, 5 statuses, priority and requester, notes timeline, related tickets, oldest/newest sort, **Import Forms file** (MS Forms export, preview, de-duplicated) |
| **Inventory → Supplies** (new sub-page) | Printer consumables: items, transactions ledger, recomputed stock summary, restock / negative / unknown-item flags |
| **Asset 360** | ✏️ Edit unit details inline (serial locked), sticky "Back to Inventory" bar, Lenovo warranty lookup link |
| **Ticket 360** (new detail page) | Single-ticket view: timeline, status, related tickets; opens via `#ticket/<ticket-no>` |
| **Deep links** | `#ticket/<no>` and `#asset/<serial>` open the full dashboard on that page (new-tab friendly) |
| **School 360** | New **Hardware** KPI tile and section (laptops, desktops, open tickets, warranties expiring), hidden when a school has no hardware |
| **Scan page** | Photo dialog with manual entry, batch tagging (one dialog for many serials), duplicate-photo skip, duplicate-serial warning |
| **Saving** | Changes survive a page refresh (entry journal). **🔗 Link file / 💾 Save to Excel** (Chrome/Edge) writes changes straight into the original workbook as cell edits, keeping its formulas and layout. Export download is still available |

## 2. Files

| File | Status | Lines | Role |
|---|---|---|---|
| `js/inventory.js` | modified | +1051 / −111 | Inventory, tickets, Ticket 360, Asset 360 edit, Forms import, deep links |
| `js/scan.js` | modified | +291 / −11 | Photo dialog, batch tag, duplicate photo/serial handling (1.2.0 fixes intact) |
| `js/app.js` | modified | +95 / −1 | `dash-ticket360` filters, supplies parse, file fingerprints + journal replay, School 360 `s360Hardware()` |
| `index.html` | modified | +11 | Link-file button, `dash-ticket360` panel, `supplies.js` / `patch.js` / `persist.js` script tags |
| `css/styles.css` | modified | +22 | Scan photo dialog, autocomplete dropdown, sticky Asset 360 bar, hardware tile |
| `js/supplies.js` | **new** | 29 KB | Supplies (consumables) module |
| `js/patch.js` | **new** | 31 KB | Patch-in-place saver: journal → targeted cell writes in the original workbook |
| `js/persist.js` | **new** | 23 KB | Refresh-proof journal (localStorage) + File System Access direct save |
| `js/excel-loader.js` | modified | 1 line | `QBR.VERSION = "1.21.0"` (the only change; the harness checks it matches the folder name) |
| `README.txt` | updated | — | New user-facing entry for 2026-10-05; the 2026-10-04 entry, which the delivery zip had dropped, is restored |
| `tests/ui-smoke.cjs`, `tests/ui-scan.cjs` (+2 images) | updated | — | Ticket 360 is a detail page; version check reads `QBR.VERSION` instead of a hard-coded "v1.9.0"; queue test uses new photos (identical photos are now skipped) + a duplicate-photo check |

Unchanged: the dashboard engine (`excel-loader.js` apart from the version line, `report-generator.js`, `chart-generator.js`, `data-quality.js`, `shell.js`), `libs/tesseract/**`, `libs/zxing/**`.

Script order (index.html): `… inventory.js → supplies.js → libs/tesseract/tesseract.min.js → scan.js → app.js → patch.js → persist.js → shell.js`.

## 3. Verification (2026-10-05, headless Chromium, `file://`)

| Suite | Result |
|---|---|
| Engine harness (`tests/run-tests.cjs`), real tracker / synthetic fixture | 100 / 104 passed, 0 failed (only the version-folder check skips outside the live folder) |
| Scanner parser (`tests/scan-tests.cjs`) | 23 / 23 |
| UI smoke (`tests/ui-smoke.cjs`) | 71 / 71 (17 nav items; Asset 360 and Ticket 360 are detail pages) |
| UI features (`tests/ui-features.cjs`) | 37 / 37, zero console errors |
| Scan end-to-end (`tests/ui-scan.cjs`) | 12 / 12. This is the live browser test the merge guide (§8.4) listed as still open: synthetic labels read correctly, the offline engine starts, the queue works, an identical photo is skipped |
| Inventory smoke (sample + tracker) | 600 assets / 85 tickets, Total Inventory KPI, Supplies loaded, `#ticket/…` → Ticket 360, `#asset/…` → Asset 360, Tenant Health Index unchanged (69), zero console errors |

## 4. Privacy / PII review

**Code and app files — no personal data found.** Scanned all non-library files for e-mail addresses, phone numbers, tenant IDs (GUIDs), personal names, user names and local paths: **none**. (One "cedo" match was a false positive inside `officedocument`.) The single external link (Lenovo warranty lookup) is a plain link that sends no serial number or other data.

**Flagged:**

| # | Finding | Where | Risk | Recommendation |
|---|---|---|---|---|
| 1 | **Entries are stored in the browser for up to 90 days.** The journal (`localStorage` `qbr-inv-journal-v1`) holds every recorded change: serials, clients, ticket requester names, contact person, address, contact details, Forms-imported **work e-mail and phone** (in ticket notes). | `js/persist.js`, `js/inventory.js` | Medium on **shared PCs**: anyone using the same browser profile can read it | Use on personal work profiles only; clear the site data when done on a shared PC. Optional: add a "Clear saved entries" button and a shorter retention |
| 2 | **Forms import copies requester e-mail and phone into ticket notes**, which are written into the workbook on save | `js/inventory.js` (Forms import) | Low–medium: personal contact data enters the inventory workbook | Confirm this is intended. Otherwise keep only the name and Forms ID |
| 3 | **Client (school) names and tenant-specific figures in `README.txt`** (older changelog entries, e.g. storage corrections per school) | `README.txt` | Low: business-confidential, not personal. The repo is private | Keep the repo private. Optionally anonymise the old entries |
| 4 | **Synthetic sample workbook** uses realistic-looking +63 mobile numbers (sequential, e.g. +63 917 111 2233), fictional people and addresses, and some school names that may exist in real life (e.g. "Corpus Christi College", "San Isidro Integrated School") | `tests/SAMPLE_Lenovo_Inventory_SYNTH.xlsx` | Low | Optionally switch to clearly invalid numbers (+63 900 000 00xx) and obviously fictional school names |
| 5 | **Merge guide names a staff member ("Pedro")** and dev workspace paths | `docs/qbr-inventory-merge-guide.md` | Low (internal, private repo) | Keep or redact as you prefer |

## 5. Known limitations (from the merge guide, still true)

1. Saving into the linked file keeps formulas, layout and column widths, but **cell styling (bold, colours, borders) is flattened**. This is a limit of the free SheetJS build.
2. The Lenovo warranty lookup can't pre-fill the serial number; no official URL supports it.
3. Forms-import de-duplication memory is browser-local. Across browsers it relies on the Forms ID stamped in ticket notes.
4. Direct "Save to Excel" works only in Chrome/Edge (File System Access API). Other browsers use the Export download.
