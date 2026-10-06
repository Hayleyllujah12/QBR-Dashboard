# Format-safe save — design, alternatives, test strategy

Status: v0.2.0 **live on main** (2026-10-06; 0.2.0 adds hyperlinks for the v1.28 audit editor). Module `js/xlsx-surgical.js`, wired in `persist.js` (`fsSaveSurgical`).
Opt out per browser: `localStorage.setItem("qbr-save-engine","legacy")`.

## 1. Problem
SheetJS CE rebuilds the entire workbook from its object model on save. Whatever it doesn't model is lost:
cell styles, conditional formatting (CF), data validation (DV), Excel tables, charts and drawings, comments,
pivots and Power Query. It also double-escapes hyperlink targets (`&amp;amp;`). Excel has no setting that
prevents this, because the loss happens in the writer.

## 2. Design: diff → splice → raw copy
1. Read the linked file's bytes. Parse them twice with identical options (`cellStyles:true`) into `base` and `work`.
2. Replay the journal into `work` (existing `patchWorkbookFromJournal`; patch logic unchanged).
3. `diffSheet(base, work)` per sheet: the cells whose value or formula changed are exactly what gets written.
4. For each changed sheet, rewrite only the affected `<row>`s in the sheet XML:
   - Keep each cell's `s=` style index. New cells inherit `s=` from the nearest cell above in the same column.
   - Strings are written as `inlineStr`, so `sharedStrings.xml` is never touched.
   - Dates become serials (honours `date1904`). A date xf is appended to `styles.xml` only if a cell has none.
   - Drop the `spans=` attribute on rows that are edited. Update `<dimension>`.
5. Excel tables on the sheet grow when rows are appended directly underneath. They gain a `tableColumn` when a header is written in the next column. `autoFilter` ref is synced.
6. New sheets (e.g. EDIT LOG) are added to workbook.xml, its rels and `[Content_Types].xml`.
6b. **Hyperlinks (v0.2.0)**: compares each cell's link target with entities decoded until stable.
    - A new link adds a `<hyperlink>` (placed in schema order) and an External relationship in the sheet's `.rels` (the file is created if missing).
    - A changed link updates its relationship `Target`, escaped once. `#Sheet!A1` targets use `location=`.
    - Any sheet `.rels` with multi-encoded targets (`&amp;amp;`, left by the old engine) is repaired on save.
    - The read-back gate also checks links.
7. `fullCalcOnLoad="1"` is set. If formulas were written, `calcChain.xml` is removed along with its rel and content-type entry.
8. Own zip writer. Untouched entries are copied **raw**, with the same compressed bytes and CRC; only the central-directory offsets change. Changed entries are deflated with native `CompressionStream('deflate-raw')` (Chrome/Edge 103+). Zero new libraries.
9. Gates (nothing is written unless all pass):
   - `verifySurgical`: every original part is still present, `vbaProject.bin` CRC is identical, the main content type is unchanged, and every edited cell reads back with the expected value.
   - The existing `fsVerifyBytes`.

**Refusals.** These return `{ok:false, reason}`; the save then downloads a separate copy and leaves the linked file untouched:
- zip64 or encrypted entries;
- `.xlsb`;
- rows or cells without `r=`;
- overwriting a shared-formula master or an array formula;
- missing `<sheetData>`.

## 3. Alternatives evaluated (sources checked 2026-10-06)
| Option | Fidelity | Offline | Verdict |
|---|---|---|---|
| ExcelJS 4.4.0 | Rebuilds from its own model. Drops charts (#2607); tables trigger Excel "repair" (#2585). Last release 2023-10. | yes | Reject |
| SheetJS Pro | Claims full-fidelity "Edit" build. Pricing via sales. | yes | Only paid drop-in; needs a trial on real files |
| xlsx-populate 1.21 | Edits XML in place (right idea) | yes | Abandoned since 2020 |
| hucre 1.2 | openXlsx/saveXlsx copies unknown parts | yes | Too new (first release 2026-03) |
| xlsx-template / docxtemplater | Placeholder templating | yes | Wrong fit |
| Microsoft Graph Excel API | Perfect (Excel does the edit) | **no** (online, SharePoint/OneDrive, auth) | Breaks offline rule |
| Office.js add-in | Perfect | **no** (manifest needs HTTPS hosting) | Different product |
| **Surgical (this)** | Untouched parts byte-identical | yes, 0 KB deps | **Chosen** |

## 4. Test strategy
Pyramid:

| Layer | File | What | Count |
|---|---|---|---|
| Unit + fidelity (Node, ~5 s) | `tests/surgical-save.cjs` | Zip round-trip raw copy; no-op; edit/append/new column/new sheet; escaping; dates; style inheritance; table/dimension growth; content type + macros; five more saves with no drift; legacy-loss baseline; guard rails | 106 (114 with `VALIDATE=1`, incl. hyperlinks) |
| Independent validators | same, `VALIDATE=1` | openpyxl load (tables/CF/DV counted) + LibreOffice headless convert | +6 |
| E2E audit (Playwright) | `tests/ui-audit-save.cjs` | v1.28 Audit editor on `SAMPLE_AUDIT_RICH.xlsx`: progress bar + version badge + exempt pill/row (UI); risky count, reference link, EXEMPT K/L headers + values, add row, 2 saves; CF/DV kept, other month sheets byte-identical | 15 |
| E2E (Playwright, real app) | `tests/ui-surgical.cjs` | Real Inventory API edits → Save ×3 via mocked file handle → byte-level inspection; legacy opt-out | 38 |
| Regression | `ui-save`, `ui-inv-fixes`, `run-tests`, `ui-smoke`, `ui-scan`, `scan-tests`, `feature-manifest` | Unchanged behaviour | all green |
| Manual UAT (required before main) | checklist below | Real Excel desktop + real workbooks | — |

Fixtures (synthetic, no PII), from `python3 tests/make-rich-fixture.py`:
- `SAMPLE_Lenovo_Inventory_RICH.xlsx`: styles, number formats, widths, freeze panes, CF, DV, table, chart, formulas, defined name, merged cells, comment, and a hyperlink containing `&`.
- `SAMPLE_Lenovo_Inventory_RICH.xlsm`: the same file plus `vbaProject.bin`.
- `SAMPLE_Lenovo_Inventory_RICH_LO.xlsx`: a LibreOffice re-save, so it uses shared strings and a different producer.

Known gaps / next tests:
- No real Excel in CI. Excel's "repair" dialog can only be checked by hand (UAT).
- No fixture yet with calcChain, pivots, Power Query, or shared-formula **children** being edited.
- The Supplies workbook (ETG_PRINTER_INVENTORY) is not yet exercised end-to-end.
- Performance with large sheets (>50k rows) is unmeasured.

Pre-existing fix included: `QBR.fsCarryPatchState` now re-points model rows' `_src.fp` after a save. Before this, a second save in the same session skipped edits to existing rows. This affects live/main too.

## 5. Manual UAT checklist (QBR-Dashboard_dev, on COPIES of real workbooks)
1. Link a copy of the real Lenovo Inventory `.xlsx` and make a single-cell edit. Save, then open it in Excel desktop. Expect: no repair prompt, formatting identical, and the edit present.
2. Run a batch tag of 5+ units, which appends rows. Expect: the table/filters cover the new rows, and the new rows look like their neighbours.
3. Set a batch code on a file that has no Batch Code column. Expect: a column is added with a header in the same style.
4. Repeat 1–3 on the `.xlsm` copy. Expect: macros still run.
5. Save three times in one session without reloading. Expect: each save shows "Saved ✓ … · formatting kept".
6. Open the saved file in Excel Online/SharePoint. Expect: it opens without errors and links work (no `&amp;`).
7. Supplies workbook: record a transaction and save. Expect: same checks as step 1.
8. Edit the file in Excel while it's linked, then Save. Expect: the dashboard refuses and downloads a copy.
9. Escape hatch: set `qbr-save-engine=legacy`, then save. Expect: the old behaviour.
