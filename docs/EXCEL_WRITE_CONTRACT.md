# Excel write contract — scope (2026-10-28)

Status: **Phase 0 built** (PR #31). **Phase 1 built** (branch `refactor/excel-write-contract-p1`, see §7). Phases 2–5 not built. Decisions: §4. Branch `refactor/excel-write-contract` (from `develop` 1353168, after PR #30).
Goal: every feature that edits an Excel file goes through ONE save path, so every module behaves the same way when:
- the file is open in Excel (locked);
- the file was changed in Excel (merge);
- a sync is running;
- permission is missing;
- a safety check fails.

It also gives a future feature a checklist instead of a new save routine.

## 1. Where we are

| Module | Workbook | Records edits first? | Writes through | When the file is open in Excel |
|---|---|---|---|---|
| Risky sign-ins editor | RISKY_USERS_AND_DOMAIN | yes (journal, kind `audit`) | shared `QBR.fsSaveKind` | red note, edits kept (PR #30) |
| Guided audit (storage/usage) | STORAGE / USAGE reports | yes (`storage`, `usage`) | shared | same |
| Inventory (+ Scan intake, tickets, deployments) | Lenovo Inventory | yes (`assets`) | shared | same |
| Supplies | supplies workbook | yes (`supplies`) | shared | same |
| Domain registration (inline date / validity) | DOMAIN_REGISTRATION_TRACKER | **no** — writes each edit at once | own `createWritable` in `js/domain-edit.js` | alert, change dropped |
| MOS › Certificates (print / reprint / completion) | MOS Certificate Log | **no** — own op list applied at once | own `createWritable` in `js/mos-certs.js` | alert, change dropped |
| Trainers (write sessions to tracker) | trainers tracker | **no** — plan applied at once | own `createWritable` in `js/trainers.js` | alert, **downloads a copy** |

How the shared path works today:
1. `QBR.journalRecord(kind, op, args)` records the edit (it survives reloads; IndexedDB backup).
2. On replay, `QBR[op](...args)` re-applies it to the model.
3. `QBR.fsSaveKind(kind)` saves:
   - checks permission (folder grant);
   - if the file changed, does a three-way merge (`xlsx-merge.js`);
   - applies the ops to a fresh workbook with `PATCH_OPS[op](wb, fp, args)` (`js/patch.js`);
   - writes only the changed cells (`xlsx-surgical.js`) and verifies;
   - writes the file, detecting a lock;
   - clears exactly the saved ops and reloads (only changed files re-read, PR #30);
   - shows the message (`QBR.saveResultText`).

The three outliers each re-implement part of this: their own lock detection, their own "changed in Excel" check, their own lineage.

## 2. Target design

### 2.1 One adapter per workbook kind
`QBR.wbKinds.register(kind, adapter)` (new, in `js/patch.js` or a new `js/wb-kinds.js`):

```
adapter = {
  label:      "Domain registration",
  detect(wb)  -> bool                // is this my workbook? (sheet/header shape, not file name)
  ops: {                             // op name -> cell writer, same signature as PATCH_OPS today
    drSetRegistration(wb, fp, args) -> "ok" | "noop" | "skip: <why>",
    ...
  },
  rowKey(op, args) -> string|null    // row identity for merge/re-location (domain, verification code, school+month…)
  editLog: true|false                // append the "EDIT LOG" sheet? (false for the tracker / MOS log / trainers tracker)
  source: "loaded" | "side",         // see 2.2
  autosave: true|false               // save right after each edit (keeps today's "instant" feel)
}
```

- The existing kinds (`audit`, `storage`, `usage`, `assets`, `supplies`) are registered by wrapping today's `PATCH_OPS`. There is no behaviour change for them.
- `pPatchPass` looks up `wbKinds[kind].ops[op]` first, then falls back to `PATCH_OPS` during the move.

### 2.2 Two ways a workbook is attached
- **loaded**: the file is in `APP.files` (Upload / Open folder) and linked in `QBR._fsLinks`. Domain registration is already like this (kind `domainreg` is registered by the loader since the bug-fixes PR).
- **side**: the tab opens the file itself (MOS Certificate Log, Trainers tracker) and keeps its own handle.
  - The contract adds `QBR.fsAttachSide(kind, handle, bytes, name)`. It registers a link `{kind, handle, fp, hash, side: true}`, so the journal, merge base, lineage and Save work exactly like for a loaded file.
  - After a save, the tab gets `onSaved(bytes)` to re-parse its own model. There is no `loadItems` (side files are not dashboard sources).

### 2.3 Behaviour rules (all kinds)
1. Every edit is `journalRecord` first.
   - With `autosave`, the module then calls `QBR.fsSaveKind(kind)` itself. The feel is unchanged (saved in a second), but a locked file now keeps the edit pending with the red note instead of dropping it.
2. Only `persist.js` opens a file for writing. No module calls `createWritable` or shows its own "couldn't write" alert.
3. Messages come from `QBR.saveResultText`, shown in the module's `[data-unsaved="<kind>"]` badge.
4. "Changed in Excel" = merge (same review dialog), never "download a copy" — except when the safety check fails (`download-unsafe`, as today).
5. Cell writes keep formulas (never overwrite a formula cell), styles and number formats (format-safe engine).

### 2.4 Guard rails so it doesn't drift again
- **Lint test** `tests/write-contract.cjs`: fails if any `js/*.js` except `persist.js` contains `createWritable(`.
  - Phase 0 ships it with an allowlist of the three current outliers. Each migration removes its entry.
- **Shared scenario test** `tests/ui-write-contract.cjs`: one table of kinds × scenarios, all with dummy sample workbooks:
  - save ok;
  - locked → edits kept + note;
  - changed in Excel → merged;
  - same-cell conflict → Cancel writes nothing;
  - file keeps changing → busy;
  - permission denied;
  - reload keeps pending edits;
  - a second save in the same session.

  A new kind passes by adding one row (sample file + one edit + the expected cell).
- **CLAUDE.md checklist** "Adding an Excel write feature" (next to "Adding a dashboard tab"):
  1. register the adapter;
  2. record ops;
  3. badge `data-unsaved`;
  4. add a sample workbook + one row in the scenario test;
  5. README entry.

## 3. Phases (each = its own PR into `develop`, beta check, then on)

| Phase | Content | Size | Notes |
|---|---|---|---|
| **0** | This doc · lint test with allowlist · scenario test for the 4 existing kinds · CLAUDE.md checklist | S | No app behaviour change |
| **1** | `wbKinds` registry + `fsAttachSide`; existing kinds wrapped; `editLog` flag | M | Must be byte-identical saves for existing kinds (compare outputs before/after on the sample workbooks) |
| **2** | MOS Certificate Log → kind `moslog` (side, autosave, editLog off); its `applyOps` becomes the adapter ops keyed by verification code | S–M | Non-Chrome keeps "Download log" (no handle) |
| **3** | Domain registration → kind `domainreg` (loaded, autosave, editLog off); ops `drSetRegistration`, `drSetValidity` keyed by domain; keeps the "typed Expiration" warning | M | Removes its own lineage/lock code |
| **4** | Trainers tracker → kind `trainers` (side, editLog off); `planWriteback` output becomes ops keyed by session ID | M | Locked/changed no longer download a copy; pending sessions show in the badge |
| **5** | Remove the allowlist; delete the dead code paths | S | Lint test then enforces the rule |

## 4. Owner decisions (2026-10-28)
1. **Autosave**: YES for Domain registration and the MOS log. They save right after each edit; a locked file keeps the edit pending with the red note.
2. **Trainers**: keep pending + merge (no more "download a copy" when the tracker is locked or changed in Excel).
3. **EDIT LOG sheet**: OFF for the domain tracker, the MOS Certificate Log and the trainers tracker (`editLog: false`).
4. **PRs**: one PR per phase (0 → 5), each beta-checked on `develop` before the next starts.

## 5. Risks
- **Phase 1 touches the save path of 5 kinds.** Mitigation: byte-for-byte comparison of saves before/after, the full suite, and a hosted beta pass before Phase 2.
- **Side files and the journal.** Journal entries are keyed by file fingerprint. A side file that is renamed or moved shows as an orphan in "Unsaved edits", with the existing recovery flow. This is acceptable, but it must be tested.
- **Trainers semantics change.** Sessions become "pending until saved", so the Trainers tab needs the same badge and a Save button. That is a small UI addition.
- **Version.** These are patches; no `QBR.VERSION` bump unless the owner decides otherwise.

## 6. Checklist — adding an Excel write feature (Phase 0)
1. Record every edit with `QBR.journalRecord(kind, op, args)` — never write the file from the module.
2. Add the cell writer `PATCH_OPS[op](wb, fp, args)` (Phase 1+: in the kind's adapter) and the replay function `QBR[op]`.
   - Never overwrite a formula cell.
   - Key rows by identity, not position, where possible.
3. Save only with `QBR.fsSaveKind(kind)`, and show the result with `QBR.saveResultNote(kind, r)` AFTER any re-render.
4. Put a `[data-unsaved="<kind>"]` badge next to the Save button (`QBR.saveButtonHtml(kind, id)` does it).
5. Add ONE row to `KINDS` in `tests/ui-write-contract.cjs`: a dummy sample workbook + one edit. All scenarios must pass.
6. `node tests/write-contract.cjs` must pass (no stray `createWritable`; every recorded op has a writer + replay).
7. README.txt dated entry; claude/CHANGE_QUEUE.md engineering note.

Guard-rail tests (Phase 0):
- `tests/write-contract.cjs` (Node): only `persist.js` writes; LEGACY allowlist = mos-certs.js (Phase 2), domain-edit.js (Phase 3), trainers.js (Phase 4); every recorded op has a PATCH_OPS writer + QBR replay.
- `tests/ui-write-contract.cjs` (browser): kinds audit · storage · assets · supplies × scenarios: save ok, second save, locked → kept → saved, changed in Excel → merged, same cell → Cancel, file keeps changing → busy, permission denied, reload keeps pending → saves.
  - Baseline on develop 1353168: 59/59. Phase 1 must keep this byte-for-byte; see §3.

## 7. Phase 1 — what was built and how it was verified (2026-10-28)
**Built** (`js/patch.js`, `js/persist.js`, `js/journal.js`; no module behaviour change except one fallback bug fix):
- **`QBR.wbKinds`** (`register` / `get` / `list` / `opFor`).
  - The five built-in kinds are registered by wrapping the existing `PATCH_OPS` writers: same functions, same order.
  - `pPatchPass` asks the adapter first. `PATCH_OPS` stays as the fallback until Phase 5.
  - `editLog: false` skips the EDIT LOG sheet.
- **`QBR.fsAttachSide(kind, handle, bytes, {onSaved})` / `fsDetachSide(kind)`**: a tab-owned file joins the shared path.
  - What it gets: journal by fingerprint, merge base, lock / merge / busy / denied handling, shared messages.
  - Side links are kept in `QBR._fsSideKinds` (this survives `loadItems`). They are not stored as dashboard links, are not session sources, and their tab is told via `onSaved(bytes, "saved"|"reloaded")`.
  - The journal treats attached side files as loaded (unsaved count, file name).
- **Autosave** (`adapter.autosave`): `journalRecord` → `QBR.wbScheduleSave(kind)`.
  - Debounced 300 ms; then `fsSaveKind` → `saveResultNote` → `qbr:autosaved` event.
  - A locked file keeps the edit pending with the red note.
- **Fallback copy fix** (`fsDownloadKind`): Inventory/Supplies still rebuild from the model. Every other kind now downloads a patched copy of ITS OWN file ("<name> (dashboard copy).xlsx"), and its edits stay pending.
  - Before: an audit/storage/usage fallback fell through to the Inventory builder and downloaded `Lenovo_Inventory_Export.xlsx`.

**Verified:**
- **Byte-identical saves** for the four existing kinds. `WC_DUMP=<dir> node tests/ui-write-contract.cjs audit storage assets supplies` (fixed clock) was run on the Phase 0 tip and on Phase 1.
  - All 16 saved files (save, second save, merged save, save after reload × 4 kinds) have identical SHA-256.
- **`tests/ui-write-contract.cjs`: 95/95.**
  - Registry coverage: every PATCH_OPS writer belongs to exactly one kind; the built-ins' flags are unchanged.
  - The 4 built-in kinds × all scenarios.
  - A side-file test adapter × all scenarios (with `onSaved`, no EDIT LOG).
  - Autosave: a burst saves once; locked keeps the edit pending with the red note; the next edit saves everything.
  - Fallback copy: Phase 0 fails the new check for `audit` (it downloaded the Inventory export); Phase 1 passes for every kind.
- `tests/write-contract.cjs` 38/38, and the full repo suite (see the PR).
