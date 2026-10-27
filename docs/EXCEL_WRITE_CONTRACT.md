# Excel write contract — scope (2026-10-28)

Status: **SCOPE ONLY. Not built.** Branch `refactor/excel-write-contract` (from `develop` 1353168, after PR #30).
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

## 4. Decisions needed from the owner
1. **Autosave** for Domain registration and the MOS log (save right after each edit, as today, but keep the edit pending when the file is locked). *Recommended: yes.*
2. **Trainers**: replace "download a copy" on locked / changed-in-Excel with "keep pending + merge". *Recommended: yes.*
3. **EDIT LOG sheet**: keep it off for the domain tracker, the MOS log and the trainers tracker (they're shared files; the MOS log already records who/when itself). *Recommended: off.*
4. **PR granularity**: one PR per phase (recommended), or Phase 0+1 together.

## 5. Risks
- **Phase 1 touches the save path of 5 kinds.** Mitigation: byte-for-byte comparison of saves before/after, the full suite, and a hosted beta pass before Phase 2.
- **Side files and the journal.** Journal entries are keyed by file fingerprint. A side file that is renamed or moved shows as an orphan in "Unsaved edits", with the existing recovery flow. This is acceptable, but it must be tested.
- **Trainers semantics change.** Sessions become "pending until saved", so the Trainers tab needs the same badge and a Save button. That is a small UI addition.
- **Version.** These are patches; no `QBR.VERSION` bump unless the owner decides otherwise.
