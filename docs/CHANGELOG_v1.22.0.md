# RCT OpsDesk / QBR Dashboard — v1.22.0

Released 2026-10-30 from `main` (PRs #26–#38 via `develop`). Rollback: `QBR-Dashboard_1.21.0`.
User-facing entries are in `README.txt` (dated UPDATE blocks); this file groups them for the release.

## Numbers that change after upgrading
- **"All quarters" totals** (storage, Canva, Microsoft 365 usage, and the Tenant Health Index usage/storage inputs)
  use each school's **latest quarter with data**, not its highest quarter. Totals show the current state and are
  usually lower than in 1.21.0.
- **Managed tenants** counts schools present in at least one M365 sheet; Canva-only schools are no longer counted.
  The deck cover row "Schools in scope" is now "Managed tenants".
- Security Defaults `NOT MANAGED` (tracker value NO ACCESS) displays as **"No access"**.
- Tracker regression snapshot re-recorded: portfolio Tenant Health Index **67** (was 69), bands
  Excellent 0 / Healthy 30 / Attention Needed 56 / Critical 23.

## Overview and charts (#33)
Health-index hero, "Needs attention" list, KPI strip (sparkline, meter), always-visible per-tenant table with band
filter, top-5 lists; sidebar in five sections; months without data drawn as gaps; posture comparisons per tenant
with group sizes; label collision fixes; per-tenant table fits the printed page.

## Save to Excel (#27, #30, #31, #32)
Never overwrites a file that changed mid-save; failed appends stay pending; every save outcome stays visible
("Not saved — open in Excel"…); parsed-workbook cache (reload re-reads only changed files); Excel write contract
Phases 0–1 (workbook-kind registry, side files, autosave, guard-rail tests; saves byte-identical to 1.21.0).

## MOS (#29, #34, #35)
Completion certificates: shared centre line, Adjust layout panel (hidden until needed), accent colour per exam;
roster reconciliation; Analytics view (performance by school and exam).

## Risky Investigation (#37, #38)
Each Source card recognises the dropped file (real Entra headers) and refuses a wrong one at once with
"Move to Source 0X"; amber notes for another week/school, month label, 100,000-row cap, coverage, AuthDetails pairing,
interactive/non-interactive mix-ups, Source 03 NotFound/age. "Drop all your files here" routes a whole set or a month
folder (asks for the school) and loads in 01→05 order.

## Scripts (#36)
Shared "WHERE TO SAVE" question for every built-in script (Enter / path / P = folder window); the Usage script
writes only `M365_UsageReport_<period>_<timestamp>.csv`.

## Revision list and review fixes (#27, #28)
Guided audit (newest storage/usage workbook, pick-or-type school, real Excel row mapping), inventory serial
matching with/without dash, Scan MTM variants, Postmaster dropdown filters, User Management full-width chart,
grouped searchable "Loaded sources & files" pop-up, My To-Do "Open saved file…", quote escaping in `esc()`.

## CI (#26)
Azure Static Web Apps workflow: close job uses the token; main→develop sync PRs skipped.
