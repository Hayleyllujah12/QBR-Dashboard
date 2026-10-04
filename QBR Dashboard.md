# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> ## ⚠️ CURRENT STATE (2026-09) — read `claude/OPERATING_NOTES.md` first
> The prose below was written for **1.8** and is partly stale. Corrections that override it:
> - **THIS FOLDER IS `QBR-Dashboard_1.21.0` — the LIVE version** (released + promoted 2026-10-05; `QBR.VERSION`
>   = "1.21.0", aligned with the Inventory delivery number). **`QBR-Dashboard_1.9.0` is the rollback — do not edit it**
>   (`1.8.5` is an older archive). Patch here in place (dated
>   entries, no number bump). Phase 2 (2026-09-30) consolidated `css/styles.css` into ONE tokenized layer
>   (16 numbered sections; the old v1.8 "premium polish" layer, `.tab-nav`/`.tab-dd` rules and gradients are gone).
>   **Phase 3 (2026-09-30, app.js additive):** page-KPI QoQ deltas via a silent previous-quarter *capture pass* in
>   `renderAll()` (`captureKpiPrev`, `KPI_DELTA` map keyed `"<page>|<label>"`, page scope from `withKpiPage()`
>   wrappers rebinding the renderer declarations; tenant-detail renderers wrapped with `noKpiPage`); Overview
>   posture/health → `renderStatusList()` (hidden canvases kept for Export Images); `updateNavBadges()`;
>   Tenant Health detail pane (`openThiPane/refreshThiPane/closeThiPane`, `#detail-pane` outside `#app-body`);
>   **print snapshots** — `swapChartsForPrint()` replaces every in-scope canvas with a fixed-size PNG
>   (`QBR.renderChartPng` @700px) for Export Tab + Export Deck, removed on afterprint; print CSS turns non-KPI
>   `.row`s into blocks so cards paginate. Sortable tables live in `js/shell.js` (MutationObserver re-applies
>   the sort after re-renders). Feature tests: `tests/ui-features.cjs` (37 checks as of 2026-10-03). Phase 1 = **app shell only**: `index.html` (48px top bar, left
>   sidebar, breadcrumb page header, one Export menu), a Fluent 2 token layer appended to `css/styles.css`,
>   and new `js/shell.js` (nav state/aria-current, rail + drawer, Export menu, theme-button label, `#tab`
>   deep links). **app.js, excel-loader.js (except the VERSION constant), chart-generator.js,
>   report-generator.js and data-quality.js are byte-identical to 1.8.5.** Shell contract:
>   nav items are `button.tab-btn.sb-item[data-tab]` (app.js binds every `[data-tab]`; only nav items may
>   carry it); panels stay DIRECT children of `#app-body` (applyDeckOrder); export buttons keep their ids
>   inside `#export-menu`; `.app-header` height uses `--topbar-h`, never the app.js-measured `--hdr-h`
>   (measuring a hidden header during print once collapsed it to 1px). The old "Adding a tab" step 1 is now:
>   add a `.sb-item` button in the right `.sb-group` of `#app-sidebar`. UI regression test:
>   `tests/ui-smoke.cjs` (Playwright; 69 checks).
>   **2026-10-03:** (1) THI Postmaster dim now uses the latest *scorable* month per school —
>   `latestScorableReputation(rows)` → `{key:{rep,mi}}` (exposed as `QBR._latestScorableReputation`); the old
>   one-liner compared `p.monthIdx >= e.mi` on a stored *string*, so `e.mi` was undefined and the FIRST month
>   always won. Harness replica fixed + INV test; snapshot unchanged (THI 69). (2) **School 360** page
>   `#dash-school` (`data-deck="skip"`, top-level nav after Overview): `renderSchool360()` → `#s360-body`,
>   picker `#s360-school`/`#dl-s360`, `APP.school360` (default: global school filter → weakest THI tenant),
>   `openSchool360(key)` (also from the THI pane `#dp-s360`), rule-based action list, CSS §19. Quarter
>   chips only (`TAB_FILTERS`); Export Tab yes, deck/gallery no. Feature tests now 37.
>   **2026-10-04 — Inventory + Scanner merged (additive).** New `js/inventory.js` (parsers, flags, write-back,
>   `renderInventory`, Asset 360 `dash-asset360` detail page — no nav item, `data-deck="skip"`) and `js/scan.js` 1.2.0
>   (Scan page `dash-scan`). `app.js` 4 hooks only: `APP.model.inventory = QBR.parseInventoryBuffers(buffers)` in
>   processBuffers (try/catch), Inventory nav badge = open tickets, `renderInventory/renderAsset360Panel/renderScan`
>   in renderAll (typeof-guarded), `TAB_FILTERS` `[]` for the 3 pages. `index.html`: Inventory nav group + 3 panels +
>   scripts (`inventory.js`, `libs/tesseract/tesseract.min.js`, `scan.js` before `app.js`). CSS §20/§21. Vendored:
>   `libs/tesseract/` (Tesseract.js 5.1.1 + generated `offline/` assets for file:// — rebuild with
>   `node tests/build-tesseract-offline.cjs`), `libs/zxing/` (@zxing/library 0.23.0; native BarcodeDetector is absent on
>   Windows). Never `window.APP` (top-level const) — use `typeof APP !== "undefined"`. Docs: `docs/SCANNER_FIX.md`,
>   `docs/INVENTORY_DEV_NOTES.md`, `docs/qbr-dashboard-rebuild-guide.md`. Tests: `tests/scan-tests.cjs` (23),
>   `tests/ui-scan.cjs` (11, file:// or http://), ui-smoke now layout-aware (15 core nav + Inventory/Scan = 17).
>   **2026-10-05 — Inventory v1.21.0** (see `docs/CHANGELOG_v1.21.0.md` + `docs/qbr-inventory-merge-guide.md`): new
>   `js/supplies.js` (consumables), `js/patch.js` (journal → targeted cell writes into the ORIGINAL workbook; never
>   overwrites formula cells), `js/persist.js` (localStorage journal `qbr-inv-journal-v1` keyed by file fingerprint,
>   replayed after every parse; File System Access link/save, handles in IndexedDB `qbr-cache`). New detail page
>   `dash-ticket360` (no nav item, `data-deck="skip"`), deep links `#ticket/<no>` / `#asset/<key>`. app.js adds:
>   supplies parse, `QBR._currentFps` + `QBR.journalReplayFor(APP.files)` after load, `s360Hardware()` tile/section.
>   Script order: inventory → supplies → tesseract → scan → app → patch → persist → shell. **PII note:** the journal
>   and Forms import hold requester names / contact details / work e-mail & phone in browser storage (90 days).
> - **Live / production version is `QBR-Dashboard_1.8.5`** — app label **1.8.5** (promoted
>   2026-09-17). In-place fixes are dated patches, **not version bumps** — the label stays 1.8.5.
>   Recent behavior: security-default status shows as "Not Enabled" (was "Disabled"; canonical model
>   values stay `ENABLED`/`DISABLED`/…, `SD_LABEL` relabels display, the `f-secdef` filter keeps the
>   value). `QBR.chart.bar` forces `scales.x.ticks.autoSkip=false` so every category
>   label shows (Chart.js was dropping crowded rotated org names), and `renderStorageDetail` scopes
>   the tenant's rows to `qActive(s.quarter)` so the storage snapshot/KPIs/composition follow the
>   Quarter chips instead of always showing the latest quarter. The three "by
>   Organization" charts (`ch-usage-org`, `ch-risky-org`, `ch-health-org`) resolve a blank org via
>   `orgOf(key)` (USAGE_REPORT leaves Organization blank for 168/436 rows). Counting is accurate:
>   unattributed-quarter rows are in scope only under "All"; per-school de-duplication runs in every
>   filter scope; storage `Used` is a median of four estimates. The Security & Risk Detail table also
>   has an **auth-method filter** (`QBR.parseAuthMethods` → canonical methods + posture presets).
> - ⚠️ **Multiple editors on the live folder** (OneDrive edtech-la `C:` ↔ admin `J:`, plus a
>   parallel Cowork session on the other machine). Before committing, `device_stage_files` the live
>   JS and diff (normalize CRLF first) — merge onto what's there, never blind-overwrite.
>   **`QBR-Dashboard_1.8.4` is the rollback.** Rule: **patch these files in place, no zip** —
>   in-place fixes are dated patches that keep the same folder AND the same label (1.8.5, never
>   bumped per patch); the dated
>   README / CHANGE_QUEUE blocks are the history. A whole new folder is cut only on a large
>   release: copy the `qbr-app` tree into a new `QBR-Dashboard_<next>` and bump properly
>   (`→ 1.9.0` for a large feature set). **Never use a 4-part number** like `1.8.4.1`.
> - **Postmaster:** reads Google's live states — `Issues detected` / `Verify to see
>   health` / `Not enough data` — not just HIGH/MED/LOW/BAD; verification is synthesized from the
>   tracker's `Status` (Verified/Unverified) column when no standalone sheet is present. School
>   display names are mojibake-repaired (`fixMojibake`/`dispName`); the join key is unaffected.
> - `excel-loader.js`, `app.js`, `chart-generator.js`, `report-generator.js` are **no longer
>   byte-identical to 1.7** — the engine has changed materially since.
> - **Quarter filter is multi-select:** `APP.filters.quarters` is an **array** (`[]` = All), not
>   the scalar `APP.filters.quarter` described below. Never compare `r.quarter` directly —
>   go through `qActive()` / `selQuarters()` / `isAllQ()`. A row with `quarter == null` is
>   **unattributed**: in scope only under "All", never under a specific quarter.
>   `dedupeBySchool()` must run in every scope — do not re-add a single-quarter bypass.
> - **Storage `Used` = median-of-FOUR:** current-storage text, `Used Storage(GB)`
>   column, OneDrive+Exchange+SharePoint, **and the used figure left of "of" in the USAGE
>   cell**. Four readings let the median out-vote a unit-mislabeled cell. At two readings
>   that disagree >25% the app resolves only the diagnosable ~1024x TB-as-GB case and
>   otherwise returns null — it does not average them. Strip LRM/RLM marks before parsing
>   the USAGE cell; they silently defeated the capacity fallback previously. Capacity is guarded against
>   impossible values; the storage text column honors each cell's own TB/GB unit.
> - **Domain Status lists only tenants with a `RISKY_USERS_AND_DOMAIN` record** (Canva-only
>   tenants excluded).
> - **Risk is a 5-band scale** (0 · 1–15 · 16–50 · 51–100 · >100); `QBR.THRESH` =
>   `{HIGH_RISK:50, CRITICAL:100, USAGE_LOW:40}`.
> - **New "Data Quality" self-audit (2026-09-17, Phase 1A):** `js/data-quality.js` →
>   `QBR.audit(model)` runs on every upload — a source-hygiene score, trust banners (Q3
>   capacity, Postmaster source), a coverage matrix, a storage bad-cell register with cell
>   refs + suggested fixes, an "Export Fix List (.xlsx)" button, plus vocabulary / sentinel /
>   org-gap checks. `excel-loader.js` now stashes raw sheet rows on `model.raw` for it
>   (parsing outputs unchanged). This makes the app **five** scripts — `data-quality.js`
>   loads before `app.js`. See "Data Quality self-audit" below.
> - **Tenant Health Index (Overview, 2026-09-22):** a composite 0–100 health score per tenant.
>   The scoring engine is **pure and lives in `excel-loader.js`** — `QBR.scoreTenantHealth(dims)`
>   plus `QBR.thiScorers` (six sub-scorers), `QBR.THI_WEIGHTS` (Security 25 / Risky 25 / Adoption 20
>   / Domain 15 / Postmaster 10 / Storage 5), `QBR.THI_BANDS` (min+label only) and `QBR.thiBand(v)`.
>   Missing dims are excluded and weights renormalize (a gap is never scored 0); ≥2 dims required or
>   the tenant is unscored. Thresholds default to production when `QBR.THRESH` is absent, so it runs
>   under Node. `app.js` `computeHealthIndex()` only shapes per-tenant dims and delegates; **band
>   colors stay in `app.js`** (`THI_BAND_COLOR`). The Overview card shows the portfolio index, band,
>   distribution bar, and a collapsible **Per-Tenant Health Detail** drill-down (ranked weakest-first,
>   six sub-scores, weakest dim outlined, band-filter chips, click-a-segment-to-filter). Scorer is
>   unit-tested + a portfolio snapshot (index 69 / 109 scored / 0-44-44-21) is pinned in `tests/run-tests.cjs`.
> - **Version + testable-without-real-data (2026-09-22):** `QBR.VERSION` in `excel-loader.js` is the
>   single source of truth — the header renders from it and the harness asserts it matches the
>   `QBR-Dashboard_<ver>` folder (kills label/folder drift). The harness now also enforces a **schema
>   contract** (required sheets + column headers resolve, else fail loud) and ships a **synthetic
>   fixture**: `node tests/make-fixture.cjs` writes `tests/fixture.xlsx`, and `node tests/run-tests.cjs
>   tests/fixture.xlsx` runs the full engine with ZERO sensitive data (71/71). It auto-detects the
>   workbook (fixture vs real tracker) and picks the matching snapshot profile.
> - **Always update the changelog when you ship:** `README.txt` (user-facing) **and**
>   `claude/CHANGE_QUEUE.md` (engineering), in the same step as the code.
>
> `claude/OPERATING_NOTES.md` holds the full rules set, session workflow, and current
> data-reading reference. The sections below remain accurate for architecture, exports,
> pinned nav, theming, and the "adding a tab" checklist.

## What this is

`QBR-Dashboard_1.8.5/qbr-app/` is a **fully offline, zero-build browser app**: a Quarterly Business Review dashboard for Rakso Education's managed Microsoft 365 school tenants, built for the Education Technology Technical Operations team. Users drop Excel workbooks on it and get filtered dashboards across its tabs, a rule-based executive narrative, and a print-to-PDF export.

There is no package.json, no bundler, no test runner, no git repo. "Running it" = opening `QBR-Dashboard_1.8.5/qbr-app/index.html` in a browser (double-click; `file://` works — everything in `libs/` is vendored: Chart.js, SheetJS/XLSX, Bootstrap 5). "Testing a change" = reload the page and re-upload a workbook. Data is in-memory only; nothing is uploaded, cached, or written to disk, and reloading clears everything. The one exception is the theme preference in `localStorage` (`qbr-theme`) — never store workbook data there.

### Versioning convention — 1.8.5 is live; patch in place

Versions are snapshots in numbered folders (no git), named
`QBR-Dashboard_<MAJOR>.<MINOR>.<PATCH>`. Progression to date: `1.7 → 1.8 → … → 1.8.4 → 1.8.5`. Rules:

- **`QBR-Dashboard_1.8.5/qbr-app/` is the live/production version** (promoted 2026-09-17). Small
  fixes and even sizeable feature work are patched **in place** under 1.8.5 (no new folder, **no
  number change, no label bump**); the dated blocks in `README.txt` + `claude/CHANGE_QUEUE.md`
  are the history. `QBR.VERSION` in `excel-loader.js` is the single source of truth for the label,
  and the regression harness asserts it matches the folder name.
- **A new version is cut only on a deliberate release of a major change** — copy the whole
  `qbr-app` tree into a new `QBR-Dashboard_<next>` folder and bump the number **properly**
  (increment PATCH for a normal release, or MINOR for a large feature set, `→ 1.9.0`). Update the
  header label (`QBR.VERSION`), the `README.txt` title, and the CHANGE_QUEUE "Live version" line
  together, in the same step.
- **Do NOT use 4-part numbers** (e.g. `1.8.4.1`) and do NOT bump the number/label for in-place
  patches. `QBR-Dashboard_1.8.4` is the rollback. Earlier `QBR-Dashboard_1.7 … 1.8.3` folders are archives.

Note: the older claim that `excel-loader.js`/`report-generator.js` are byte-identical to 1.7
is **no longer true** — the loader (storage parsing, postmaster, master dimension) and app
(multi-quarter filter, drill-downs, domain-status scoping, data-quality audit) have changed
materially. See `claude/CHANGE_QUEUE.md` for the per-change engineering log and `README.txt`
for the changelog.

## Architecture

Five plain scripts, loaded in order by `index.html`, all attaching to one global `window.QBR` namespace. Load order matters — `app.js` must be last:

1. **`js/excel-loader.js`** — Excel → normalized in-memory model. Exposes `QBR.loadWorkbooks(buffers)`, `QBR.buildCleanRows`, the pure engines `QBR.scoreTenantHealth` (Tenant Health Index), `QBR.parseAuthMethods` (auth-method tokenizer) and `QBR.domainStatus` (domain-registration status), plus `QBR.VERSION` and `QBR.util`. Also stashes raw sheet rows on `model.raw` (`captureRaw`) for the Data Quality audit.
2. **`js/chart-generator.js`** — `QBR.chart.{line,bar,hbar,pie,doughnut}` over Chart.js. Exposes `QBR.COLORS` / `QBR.PALETTE`.
3. **`js/report-generator.js`** — `QBR.generateReport(model, agg)`, deterministic prose. Exposes `QBR.THRESH`.
4. **`js/data-quality.js`** — `QBR.audit(model)` + `QBR.auditFixRows(audit)`, the offline source-hygiene self-audit rendered on the Data Quality tab. Loads before `app.js`.
5. **`js/app.js`** — the `APP` singleton (`model`, `filters`, `activeTab`), aggregation, all `render*()` functions, filter wiring, upload handling.

### Data sources

`loadWorkbooks` merges **any number of uploaded workbooks into one model**. Sheets are located by *name*, never by filename, so the master tracker and split exports can be uploaded together or separately.

| Workbook | Sheet(s) | Parser | Model array |
|---|---|---|---|
| ALL TENANT AUTOMATED TRACKER.xlsx | `RISKY_USERS_AND_DOMAIN` | `parseRisky` | `risky` |
| ″ | `SECURITY_DATA` | `parseSecurity` | `security` |
| ″ | `STORAGE_DATA` | `parseStorage` | `storage` |
| ″ | `USAGE_REPORT` | `parseUsage` | `usage` |
| ″ | `CANVA_STATUS` | `parseCanva` | `canva` |
| ″ | `GOOGLE_POSTMASTERTOOLS` | `parsePostmaster` | `postmaster` |
| USER MANAGEMENT.xlsx | `SY 2025-2026`, `SY 2026-2027` (regex `SY 20xx-20xx`) | `parseUserManagement` | `usermgmt` |
| DOMAIN REGISTRATION workbook (e.g. `QBR_domainreg.xlsx`) | the domain-reg sheet, found by header SHAPE — `Registration status`, OR `DOMAIN` + (`Expiration`/`Remaining Days`) — so both the old `CURRENT` layout and the per-quarter `Q3` export load | `parseDomainRegistration` | `domainreg` |

Each source sets a `model.sources.<name>` boolean, surfaced as badges after upload.

### Time dimensions (mixed, by design)

This is the single most important thing to internalize before touching aggregation:

- **Risky Users / Domain Health / Postmaster** — **monthly** (`MONTH` column); quarter is *derived* (Q1 = Jan–Mar).
- **Storage / Usage / Canva** — **quarterly** (`QUARTER` column); no monthly detail.
- **Security** — a point-in-time **snapshot** with no time column, joined to monthly risky data by school key.

### Master School Dimension

School names vary across sheets (`DR. YANGA'S COLLEGES` vs `Dr Yanga's Colleges`). Every fact row carries `key` (canonical join key) and `schoolRaw` (display). `buildMasterSchools()` builds `model.master` (key → `{name, aliases, inSheets}`), `model.orgByKey`, and `model.quality`. Sheets lacking an Organization column (Canva, User Management, Domain Registration) resolve theirs through `orgByKey`.

**Source data is dirty by design.** All defensive normalization lives in the loader; extend it there rather than pushing cleanup downstream:
- `schoolKey()` — uppercase, strip zero-width/bidi unicode and punctuation, collapse whitespace. Join key only, never displayed. Canonical display name = the *longest* raw variant seen.
- `normDomainHealth` / `normSecurityDefault` / `normReputation` / `yesNo` — map typos and casing onto fixed vocabularies (`"healh"` catches the real `Healhty` typo).
- `toGB` parses `"43.71 TB"` / `"882.25 GB"` into GB; `toNum` returns **`null` for blanks and `"-"`** so "no data" never becomes `0`.
- `makeResolver(headerRow)` matches columns by header text (exact, then substring) so column order can shift.

### Business rules

- Risky users: **> 50 HIGH RISK**, **> 100 CRITICAL**; O365 low-adoption benchmark **< 40%**. All three live in `QBR.THRESH` (`{HIGH_RISK:50, CRITICAL:100, USAGE_LOW:40}`) — change them there, nowhere else.
- Storage is parsed from mixed text into GB and rolled up to **TB** for display.
- Domain registration is **source-first**: `parseDomainRegistration` reads the sheet's own `Status` + `Remaining Days` (and `Remarks`) as authoritative, computing from dates only as a fallback. The `Registration` column is dirty (may hold `Error`/`PENDING` sentinels, not a date). Pure `QBR.domainStatus(rec, now)` maps everything onto a canonical operational vocabulary — Active / Expiring Soon (Urgent ≤30, Watch ≤60) / Expired / For Renewal / For Deletion / Deleted / End Contract / Invalid Domain / Pending / Error / Not Registered — with an action + KPI group. Precedence: Remarks → Status column → Registration sentinel → signed days → fallback. Domain registration is upload-based like every other source (no built-in default).

### Filters and rendering

`APP.filters` is global state: `quarters` (an **array**, `[]` = All — never the old scalar `quarter`), `month, org, school, secDefault, usageCat, sy, status, authMethod`. Any change calls `renderAll()`, which re-renders **every** dashboard (charts self-destroy and rebuild via `QBR.draw`). `TAB_FILTERS` declares which shared filters apply per tab; `applyFilterVisibility()` hides the rest. Tab-specific filters (Security Default + Authentication Method; School Year + Status; Domain Registration status) live inside their own panels.

Quarter/Month/Security-Default/SY/Status are `<select>`s bound on `change`. **School and Organization are `<input list=...>` + `<datalist>` type-to-search**, bound via `bindSearch()`: the typed value is validated against `APP._schools` / `APP._orgs` and **anything empty or non-matching silently resolves to `"ALL"`** — a typo shows unfiltered data rather than an error. They fire on `change` (blur/commit), not `input`.

Org resolution has a subtlety worth preserving: `applyFilters` falls back through `r.org || APP.model.orgByKey[r.key] || "—"`. Sheets with a blank Organization column (Canva, User Management, Domain Registration) get theirs from the Master School Dimension; without that fallback they vanish whenever an Organization is selected.

Two correctness patterns to preserve when adding views:
- **Count per school, not per row.** `risky` and `postmaster` are monthly, so any status/reputation KPI collapses to one representative record per school — the *latest month with a non-null status*. Counting raw rows inflates totals ~12x (the difference between "71 of 110 schools" and "458 of 589 records").
- **`dedupeBySchool(rows, metricFn)`** collapses a school's quarterly rows to its max-metric row, but only when `quarter === "ALL"`; with a quarter selected the rows are already unique.

The `month` filter is deliberately narrow: it scopes KPIs and tables, but the risky-users monthly trend chart always spans all 12 months (it re-filters with `hasMonth: false`).

### Pinned navigation bars

The header, filter bar, and tab strip are three stacked `position:sticky` bars, so each one's `top` is the summed height of those above it. **Do not hardcode those offsets** (the filter bar previously used a magic `top:70px`): the filter bar wraps on narrow windows and switching tabs hides some of its controls, so its height is not fixed. `syncStickyOffsets()` measures the real heights into the `--hdr-h` / `--filter-h` CSS variables, driven by a `ResizeObserver`; the CSS values are fallbacks only. Z-order must stay header (30) > filter bar (20) > tab strip (12) so lower bars slide *under* higher ones.

Two constraints that came out of measuring it:

- **The tab strip must not wrap.** Eleven tabs wrapped to 2–4 rows cost 98–141px; as one `overflow-x:auto` row it is a flat 65px at every width. Wrapped, the pinned stack reached 51% of a 600px-wide viewport.
- **`body.no-sticky` is a safety valve.** When the bars exceed `STICKY_MAX_VIEWPORT_FRACTION` (40%) of the viewport they revert to `position:static`, because pinning them on a short window leaves no room for the dashboard. Toggling the class only changes `position`, never the bars' heights, so the measurement it depends on cannot oscillate.

The tab strip is inset by a 16px margin, so pinning it would let dashboards scroll through the side gutters. The first `box-shadow` on `.tab-nav` is a 16px ring of `var(--bg)` acting as a mask — it paints outside the border box without affecting layout. Keep it theme-variable-driven, not a literal color.

All three bars are `display:none` in `@media print`, so none of this touches either export.

### Theming and chart colors

`css/styles.css` defines the palette as CSS custom properties on `:root`, overridden under `[data-theme="dark"]` — including `--axis` and `--grid`, which exist specifically so Chart.js can read them. `setTheme()` stamps `data-theme` on `<body>`, persists to `localStorage`, calls `QBR.applyChartTheme()`, then **re-runs `renderAll()`** because Chart.js bakes colors in at construction time and cannot re-theme in place.

`QBR.applyChartTheme()` reads `--axis`/`--grid` off `getComputedStyle(document.body)` and writes them into `Chart.defaults`. It must run *before* charts are drawn — hence the call in `initTheme()` during `initShell`.

**When styling anything new, use the CSS variables, never hardcoded hex** — a literal color will look correct in light mode and break in dark. Two helpers keep single-series charts colorful: `QBR.paletteColors(n)` cycles the categorical palette, and `QBR.QUARTER_COLORS` fixes Q1–Q4 so quarters read identically everywhere. Single-series bars pass `backgroundColor` as an *array* and disable the legend (a per-bar legend is meaningless).

### Report and PDF

`computeAggregates()` in `app.js` produces the `agg` object; `report-generator.js` turns it into sections plus priority-ranked recommendations (Critical/High/Medium/Low) using only `QBR.THRESH`. No LLM, no network. If you add an aggregate, add it in `computeAggregates` and consume it in `generateReport`.

Single-tab PDF export is `window.print()` after stamping `document.body[data-print] = APP.activeTab`; `css/styles.css` has one `@media print` rule per panel id.

### Presentation export (`exportDeck` / `exportImages`)

Three buttons: **Export Tab** (active panel only), **Export Deck** (all dashboards → paginated PDF), **Export Images** (all charts → 3x PNG gallery for pasting into PowerPoint).

**The two exports use completely different strategies, and mixing them up is how this feature broke twice.**

**Images — never touch the live canvas.** `QBR.renderChartPng(canvasId, {width, height, dpr})` rebuilds the chart *offscreen* from `QBR._configs[canvasId]` (a pristine JSON clone that `QBR.draw` stashes before Chart.js mutates the config), renders it with `responsive:false, animation:false`, and returns a PNG. Nothing on screen is disturbed, so hidden panels are irrelevant and no rAF is involved. Verified: exporting leaves all 33 live canvases byte-identical.

The earlier approach — resizing the live canvases and calling `toDataURL()` — *appeared* to work in a hidden tab and failed in a real browser, because Chart.js's rAF-driven update re-runs and clears the canvas you just painted. Do not go back to it.

**Deck — never set fixed pixel sizes, and never use flexbox.** `prepareExport({dpr, deckOrder:true})` reveals every panel, forces the light theme, reorders the DOM, and raises resolution *only* via Chart.js's `devicePixelRatio` option, which grows the backing store while leaving the CSS size responsive. Writing `canvas.style.width` in px is what made the printed deck overflow the page and collide with the card below it: Chart.js sizes canvases to the *screen* container (~874px), which is wider than a printed page. `@media print` additionally forces `width:100%; height:auto !important` on deck canvases so a stale inline size can never overflow. **Always call `restore()`, including on the error path.**

Three ordering rules the deck depends on, each learned from a real failure:

- **Deck order is applied by moving DOM nodes (`applyDeckOrder()`), never by CSS `order`.** A flex container does not fragment across printed pages — Chrome overlaps the panels instead of breaking between them, which printed the Executive Report on top of the next dashboard. `#app-body` must stay `display:block` in print.
- **Reorder *before* the charts are painted.** Moving a canvas in the DOM makes Chart.js's ResizeObserver fire, which clears the backing store and defers the repaint to rAF; reordering after painting leaves every chart blank. `prepareExport` therefore moves nodes first, then resizes, then paints.
- **The final paint is a synchronous `update("none")` + `draw()`.** Print does not wait for rAF, so relying on Chart.js's deferred render can spool blank canvases.
- `#print-cover` is also a `<section>`, so `.dash-panel:first-of-type` silently matches nothing. Use `#print-cover + .dash-panel` to suppress the duplicate page break after the cover.

Three more constraints worth keeping:

- **A chart inside a `d-none` panel has a 0×0 backing store** — measured, not theoretical. Only the deck path cares (it reveals panels first); the image path sidesteps it entirely.
- **Do not pre-multiply `dpr` into `canvas.width` *and* pass `devicePixelRatio`.** That squares the scale (960×540 @2x became 3840×2160); 33 canvases that large exhaust memory and silently fail to decode, which surfaces as blank cards in the gallery.
- **`afterFrames()` races a timeout**, because rAF never fires in a hidden or backgrounded tab and the deck export would otherwise hang with every panel revealed.

`forceLightTheme()` writes `data-theme` directly rather than calling `setTheme()`, deliberately: `setTheme()` persists to `localStorage`, and an export must not overwrite the user's saved preference.

Chart PNGs are composited onto white — Chart.js canvases are transparent, and a transparent PNG on a colored slide loses all its axis text. `openGallery()` prefers a popup but falls back to an in-page overlay, since pop-ups are commonly blocked and the app's normal home is a `file://` page.

Deck order lives in exactly one place: the `DECK_ORDER` map in `app.js`, which drives both `applyDeckOrder()` (printed deck) and the image gallery's grouping.

## Data Quality self-audit (`js/data-quality.js`)

`QBR.audit(model)` runs on every upload (called by `renderQuality()`), reading the raw sheet
rows on `model.raw` (stashed by the loader's `captureRaw`) plus the master dimension. Offline
and deterministic. It returns:

- a source-hygiene **score** (0–100, graded) — describes the RAW workbook; the dashboard still
  corrects everything at read time, so a low score is expected and not a dashboard defect.
- **trust banners**, auto-detected: "Q3 storage capacity unreliable" (≥3 impossible Q3
  capacities) and "Postmaster reputation source not loaded" (embedded sheet has no
  HIGH/MED/LOW/BAD — the standalone `GOOGLE POSTMASTERTOOLS.xlsx` is needed).
- a **coverage matrix** — distinct schools per sheet vs the master union, "missing from sheet"
  lists, and the Canva-only tenants.
- a **bad-cell register** for `STORAGE_DATA` — types: unit mislabel / magnitude typo /
  impossible capacity / missing capacity / service check — each with an exact cell ref, the
  current value, and a suggested fix. Impossible-capacity threshold = >250 TB or >3× the
  school's own typical capacity. `QBR.auditFixRows(audit)` feeds `exportAuditFixList()`, which
  writes `QBR_Source_Fix_List.xlsx` via SheetJS (a browser download).
- **controlled-vocabulary violations**, **status-words-in-a-numeric-column** (e.g. "Not
  Connected" in `TOTAL RISKY USERS`), **organization gaps**, **name variants**, and
  **sentinel-row counts** per sheet.

The register is advisory — the engine already tolerates every issue at read time; the tab
exists to drive SOURCE cleanup. `model.raw` is in-memory only; `renderQuality()` stores the
last result on `APP._audit` for the export handler.

## Adding a dashboard tab

Edits across three files; missing any one fails silently:

1. `index.html` — sidebar item in the right `.sb-group` of `#app-sidebar`:
   `<button class="tab-btn sb-item" type="button" data-tab="dash-x" data-group="<Group label>">` with the two
   `fi-*` / `fi-*-f` sprite icons + `<span class="sb-label">`. Only nav items may carry `data-tab` (in-panel
   links use their own attribute, e.g. `data-s360-tab`, and call `goToTab()`); update the nav count in
   `tests/ui-smoke.cjs`.
2. `index.html` — panel `<section id="dash-x" class="dash-panel d-none">` with the KPI/chart/table element ids the renderer targets.
3. `app.js` — a `renderX()` that writes into those ids.
4. `app.js` — call it from `renderAll()`.
5. `app.js` — add `"dash-x"` to `TAB_FILTERS` (omitting it silently shows *all* global filters).
6. `styles.css` — add `body[data-print="dash-x"] #dash-x` to the `@media print` rule, **or the tab exports as a blank page**.
7. `app.js` — add the panel to `DECK_ORDER`, or it lands at the end of both the printed deck and the image gallery. (`DECK_ORDER` is the single source of truth for deck order; the print CSS no longer carries `order:` values.)
8. **Per-school / non-portfolio pages** (e.g. School 360): mark the panel `data-deck="skip"` (skipped by
   `collectChartImages()`) AND add `body[data-print="ALL"] #dash-x{display:none !important}` so Export Deck
   leaves it out. If the page has KPIs, wrap its renderer with `withKpiPage("dash-x", fn)`.

## Debugging

`APP` and `QBR` are plain globals, so everything is live in the browser console: `APP.model` (parsed data), `APP.model.sources` (which sheets were found), `APP.model.raw` (raw rows per sheet for the audit), `APP.model.quality` (coverage gaps), `APP._audit` (last Data Quality audit result), `APP.filters` (current filter state), `QBR._charts` (canvasId → Chart instance). Run `QBR.audit(APP.model)` directly to re-audit. There is no other debugging affordance.

**A synthetic fixture exists for testing without sensitive data.** `tests/make-fixture.cjs` generates `tests/fixture.xlsx` (a small fabricated tracker covering every dirty-data edge case), so `node tests/run-tests.cjs tests/fixture.xlsx` verifies the engine end-to-end with no real tenant data (71/71). The real `data/` folder is still empty (the live tracker is confidential) — run the harness against the real workbook too when available, but the fixture is enough to catch parser/aggregation regressions anywhere.

## Known behaviors that look like bugs

- `parse*` sets `model.sources.<name> = true` when the sheet is *found*, not when rows are extracted — a badly-shaped sheet shows a green badge with zero data (the Data Quality tab surfaces this). `parseUserManagement` is the only parser that guards against it.

(Earlier caveats now **resolved** and removed: null-`quarter` rows leaking into every quarter filter — they're now unattributed and in scope only under "All"; and `wireFilters()` double-binding on every upload — it now runs once in `initShell()`.)

## Conventions

- ES5/ES2015-era browser JS, no modules, no `import`/`export` (the trailing `module.exports` guards exist only for Node-side testing and are otherwise inert).
- All user-derived strings go through `esc()` before landing in `innerHTML`; tables are built as HTML strings, not DOM nodes.
- Keep it offline: no CDN links, no fetch, no network. New third-party code goes into `libs/` as a vendored file.
- Design system is Microsoft / Power BI: Blue `#0078D4`, Green `#107C10`, Orange `#FF8C00`, Red `#D13438` — in `QBR.COLORS` for charts, and as CSS variables for everything else. Both themes must stay legible.
- Missing data renders as "No data" or `—`, never `0` — preserve this when adding views.
- `README.txt` is the end-user doc and reads as a changelog — update it when tabs, filters, or supported workbooks change.
