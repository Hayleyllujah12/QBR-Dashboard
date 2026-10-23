# Spec — Auto-routing: app-level uploads feed their designated tab

**Status:** SPEC ONLY (Pedro approved scope 2026-10-23: all modules, incl. preview flows). No code yet.
**Base:** `origin/develop` (build on a feature branch per the usual flow; minor version bump at merge — new feature).
**Problem it fixes:** files loaded at the app level (Upload button, Folder picker, drag-and-drop, session restore) only build the QBR workbook model. Modules with their own ingestion (MOS, Trainers, ticket imports) never see them, so the same file must be loaded a second time inside its tab — and in a fresh window every one of those tabs starts empty even though the workbooks restored.

## Design in one paragraph

`loadItems(items, meta)` in `app.js` is the single funnel for Upload, Folder picks, and `restoreSession()`. Add one step at its end — a **router**, modelled on the existing `detectGdap(items, buffers)` precedent — that offers every loaded `.xlsx`/`.xlsm` file to registered module sniffers. A file is offered to **every** sniffer whose signature matches (never first-match-wins). Each module ingests through its **existing** entry point; the router adds no new parsers. Because restore replays `loadItems`, every routed module also rebuilds itself in a fresh window with no extra work.

## The router (app.js)

- New function `routeModuleFiles(items, meta)` called from `loadItems` after `processBuffers` succeeds and after `detectGdap`. Wrapped so a routing failure can never fail or roll back the main load (same isolation as the `processBuffers` sub-parsers: per-module `try/catch`, `console.warn`).
- Async and sequential per module, in this order: **MOS Results → MOS Reconcile → Trainers → General tickets (stage) → Inventory Forms (stage)**.
- Only `.xlsx` / `.xlsm` items are offered. PDFs, images, CSV, JSON are never routed (the app door is `.xlsx`-only; CSV ticket imports stay manual in v1).
- Dedup is each module's own existing mechanism (below) — the router itself keeps no state, so re-loads and restore replays are naturally idempotent for the silent modules.
- **Feedback:** the router returns a per-file summary; `loadItems` appends one line to the upload-status area for fresh loads, e.g. `Also routed: Results_Export_….xlsx → MOS results · Class_Lists_….xlsx → MOS roster + Reconcile · Training_Feedback_….xlsx → Trainers`. On session-restore replays (`meta.savedAt` set) the note is skipped — restore stays quiet — but badges/counts still update.

## Module contracts

### 1. MOS Results — SILENT route
- Entry: `QBR.mos.addRoster(files)` (already exposed; used today by Reconcile).
- Signature: its own — per sheet, Certiport export header signature (`parseExportRows`) else roster rows (`parseRosterRows`, NAME/Email). Non-matching workbooks contribute nothing.
- Dedup: built in (file id = name|size across `S.exports` / `S.rosterFiles`).
- School attribution: unchanged — the "school for files picked one by one" box is empty in this path, so the existing fallback applies (single loaded school wins). Acceptable for v1; noted, not changed.
- Privacy: no new storage. MOS stays memory-only; nothing about the tab's "student data stays in this tab" posture changes.

### 2. MOS Reconcile — SILENT route
- Entry: `QBR.mosReconcile.addFiles(files)` — **new thin seam** over the existing per-type loaders: a file that sniffs as a Certiport export (voucher-bearing variant, `parseCertExport`) loads as the run's export; a file that sniffs as a class roster (`parseClassRosterBlocks` header signature) loads as roster block(s).
- **New dedup (required):** Reconcile currently has no cross-load file dedup. Record seen file ids (name|size) for routed *and* manually loaded roster/export files in the module state and in the persisted run payload (`filesSeen`); routing (or re-adding) an already-seen file is a no-op. Duplicate *vouchers inside* a newly routed roster keep the existing DUPLICATE flag behaviour.
- A routed roster must not clobber an in-progress run: blocks append, export replaces only if none is loaded (if one is loaded, the routed export is skipped and the summary line says so).
- Persistence unchanged (last run in localStorage, per Pedro's earlier choice).

### 3. Trainers — SILENT route
- Entries: `openTracker(file, null)` and `openEvals(files, { auto: true })` — the exact calls the linked-folder flow already makes (`fsFolderBooks` path), so Upload and Folder behaviour finally match.
- Classification: reuse the module's own filename heuristics — `looksTracker(name)` (train + track) and `looksEval(name)` (filename parse or eval/feedback/form/survey, excluding GDAP). Content sniffing is a v2 question (see Open decisions).
- Dedup: the module's own (eval content-hash duplicates are already counted and skipped in `openEvals`; tracker reload replaces by design).
- Note: Trainers already reads the linked folder directly (v1.39.0 `fsFolderBooks` flow). That path stays as-is; routing only adds the Upload/restore side.

### 4. General tickets — STAGED route (preview-then-confirm, never silent)
- Entry today: `importDialog(file)` → parse → `previewImport` (fingerprint dedup vs existing tickets) → user confirms.
- Routing must **never** auto-open a modal or auto-import. Instead: a new seam `QBR.gentickets.stageImport(file)` stores the file, sets a **nav badge** on the General tickets tab ("import ready"), and the tab shows a banner — "1 uploaded file is ready to import — Review / Dismiss" — which calls the existing `importDialog(file)`. If the user is already on the tab when routing happens, the banner simply appears; the dialog still opens only on click.
- One staged file per module; a newer matching file replaces the staged one (noted on the banner). Dismissal is remembered per file id (name|size) in `localStorage` (ids only, no content) so session-restore replays don't re-badge a dismissed file forever.
- Signature: reuse the import parser's own header check on the first sheet — if the parser finds no ticket-shaped rows, the file is not staged (the router learns this from a cheap `canImport(file)` probe the module exposes, which parses headers only).

### 5. Inventory — Forms ticket import — STAGED route (same rules as §4)
- Entry today: `#fi-file` change handler inside Ticket 360 → parses the MS Forms export → renders the preview into `#fi-preview` → user applies.
- New seam: `QBR.inventory.stageFormsFile(file)` — same badge + banner pattern (banner lives at the top of the Forms-import panel: "Forms export uploaded — Review preview / Dismiss"). Reviewing runs the existing preview render with the staged file; apply stays manual.
- Dedup: the module already tracks imported Forms row ids (`formsImportedIds()`); the preview keeps flagging already-imported rows. Staged dismissal uses the same `localStorage` id list approach as §4 (shared helper, per-module keys).
- Signature: the Forms parser's own column signature (Forms export headers) via a header-only probe, same as §4.

## Explicitly out of scope
- MOS score-report PDFs and Certificate PDFs (app door is `.xlsx`-only; the MOS tab's folder grants stay as they are).
- User Generator (one-shot tool), Scripts (JSON import), Journal (JSON import), Scan (images).
- CSV routing (General tickets keeps its manual CSV path).
- Any change to what the core pipeline parsers accept, or to write-back behaviour of any module.

## Verification standard (Pedro's rule — exercise the changed flow)
1. **Logic tests:** router classification with dummy fixtures — one file per module signature, one unrelated workbook (routes nowhere), one class-lists workbook (routes to MOS roster **and** Reconcile), one GDAP-named workbook (still GDAP only).
2. **Browser run-through:** (a) fresh window → top-level upload of a Certiport export → MOS Results populated without touching the MOS tab; (b) reload → session restore rebuilds MOS + Reconcile automatically; (c) upload a class-lists workbook → Results roster + Reconcile blocks both appear; (d) upload a Forms export → Inventory badge appears, banner stages it, preview opens on Review, nothing imported before Apply; (e) same for a general-tickets file; (f) re-upload the same files → no duplicates anywhere; (g) upload a plain QBR workbook → loads exactly as today, router summary empty.
3. Existing suites stay green (mos-parse, mos-reconcile, ui-mos-reconcile, trainers, ui-bugfixes, domain/folder flows).

## Open decisions (defaults chosen; flag if you disagree)
1. **Trainers classification = filename heuristics** (consistent with its folder flow). Content sniffing would be stricter but diverges from the module's own convention — v2 candidate.
2. **Staged-import dismissal persists** via a `localStorage` list of file ids (names+sizes only). Alternative — session-only — re-badges after every reload until imported.
3. **Router summary line shows on fresh loads only**, not on restore replays (restore already reports "Saved …").
