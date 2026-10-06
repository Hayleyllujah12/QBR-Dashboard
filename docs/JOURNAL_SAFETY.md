# Journal safety — unsaved-edits store (v1.30.0, `js/journal.js`)

Status: built and tested on `feature/journal-safety` (2026-10-06). Replaces the v1 journal in `persist.js`.

## 1. Problem (risk audit, fix #1)
The v1 journal (`localStorage["qbr-inv-journal-v1"]`) could lose pending edits without telling anyone:

| # | Loss path | v1 behaviour |
|---|---|---|
| 1 | localStorage full (~4.75 MB, **shared by every file:// page**) | `jSave` swallowed `QuotaExceededError`; edits lived only in the tab |
| 2 | > 500 ops per file | oldest ops sliced off |
| 3 | ops older than 90 days | deleted on the next edit |
| 4 | workbook re-uploaded after an Excel edit (new fingerprint) | old ops never replayed, invisible |
| 5 | merge edits that can't be placed | listed once, then cleared with the journal |
| 6 | live folder and `_dev` copy in the same browser | same key + same fingerprint (Explorer copies keep mtime) → shared ops; saving in `_dev` cleared live's journal |

Measured (Chromium, file://): localStorage **and** IndexedDB are one origin (`file://`) for all local files;
`navigator.storage.persisted()` is `false` by default (best-effort, evictable); `isSecureContext` is true.

## 2. Design
- **Authoritative in-memory store**, envelope `{v:2, ns, files:{fp:{fileName, ops[]}}, gone:{}, parked:[], updated}`.
- **Write-through to two places** on every change: localStorage key `qbr-inv-journal-v2:<hash(app folder)>` and IndexedDB
  `qbr-cache/session/"journal:<ns>"` (auto-backup). A localStorage failure sets `health.local="failed"` → red bar + `beforeunload` prompt.
- **Startup**: localStorage read synchronously; the IndexedDB copy is merged in `QBR.journalReady` (≤2.5 s); `loadItems` awaits it before
  replaying. Merge = union by op id, ordered by `ts`; **tombstones** (`gone["fp\u0001id"]`, kept 30 days) make removals stick.
- **No caps**: no op limit, no age prune. Badge turns red at 400 ops/file.
- **Namespacing** by app folder (`file://` path, lowercased). Other folders' stores are read-only and listed ("Move to this dashboard"
  merges and removes them there). v1 entries are claimed when their exact file is loaded; the rest are listed as "previous version".
- **Row signatures**: row-addressed ops (`auditUpdateCell`, `auditPasteRow`, `storagePasteRow`, `usagePasteRow`) record `sig` = column A
  of that row from `QBR._origWb[fp]`. **Orphan apply** re-targets fp → loaded file (same name, else the only loaded file of that kind) and
  row → the unique column-A match; key-based inventory/supplies ops replay as-is. Unmatched/ambiguous/no-sig ops stay with a reason.
- **Parked**: `fsSaveMerged` calls `QBR.journalPark(name, m.unresolved)` before clearing; listed until dismissed.
- **Trash**: `journalClearFp` (save, export, discard, apply) copies the batch to IndexedDB `journal-trash:<ns>` (14 days, ≤30 batches).
- **Export/Import**: `{format:"qbr-unsaved-edits", version:1, files, parked[, trash]}`; import de-duplicates by op id and re-parses loaded files.
- `navigator.storage.persist()` is requested once, on the first recorded edit.
- All readers use the API: `patchWorkbookFromJournal` → `QBR.journalEntries()`, audit conflict review → `QBR.journalOpsFor()`,
  `fsReloadFromFile` → `journalOpsFor`, `fsRebaseLink` → `QBR.journalMove()`.

## 3. Tests
`tests/ui-journal.cjs` (44): two dashboard folders in one browser context, v1 migration, 600 ops survive reload, 200-day-old op kept,
red badge, simulated `QuotaExceededError` (Storage.prototype.setItem throws) → red bar, `beforeunload` dialog, recovery from IndexedDB on
reopen, audit orphan re-placed after a row insert in Excel (row 3 → 4), no-sig op kept with reason, inventory orphan by serial,
discard + trash, parked edit survives save + reload, panel escaping, export/import round trip + dedupe + rejection, merge unit rules.
`tests/ui-merge-save.cjs` H (+4): row deleted in Excel → that edit parked, the rest saved. `feature-manifest` +2.

## 4. Known limits / follow-ups
- Ops recorded before v1.30 have no row signature: their orphans can be exported but not auto-re-placed.
- Key-based ops (inventory, supplies) are applied as recorded; if Excel already contains the same change, applying is harmless
  (same value), but an `invIntake` of a serial that now exists may duplicate — review the "Show edits" list first.
- The session cache (`qbr-cache/session/workbooks`) and Direct-save links (`fs-links`) are still shared by all file:// folders. A `_dev`
  copy can therefore restore live's session and links — next fix: namespace them the same way.
- Trash is best-effort (IndexedDB write is async; closing the tab within milliseconds of a save can skip it).
