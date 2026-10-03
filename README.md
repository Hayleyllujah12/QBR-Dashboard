# Rakso QBR Dashboard

Offline, zero-build browser dashboard for Quarterly Business Reviews of Rakso-managed Microsoft 365 school tenants.
Open `index.html` in a browser (double-click works), upload the tracker workbook(s), and everything is processed
in the browser — no server, no network, nothing uploaded or stored.

- **Current version:** 1.9.0 (in-place fixes are dated patches; see `README.txt` for the user-facing changelog)
- **Libraries (vendored in `libs/`):** Chart.js, SheetJS (xlsx), Bootstrap 5

## Layout

| Path | What |
|---|---|
| `index.html` | App shell (sidebar, pages) |
| `css/styles.css` | Single tokenized stylesheet (light + dark) |
| `js/excel-loader.js` | Workbook → data model; pure scoring engines; `QBR.VERSION` |
| `js/app.js` | Aggregation + every page renderer, filters, exports |
| `js/chart-generator.js`, `js/report-generator.js`, `js/data-quality.js`, `js/shell.js` | Charts, narrative, audit, chrome |
| `tests/` | Regression harness + UI tests (synthetic `fixture.xlsx` only) |

## Tests

```powershell
node tests/run-tests.cjs                       # engine harness on the synthetic fixture
node tests/run-tests.cjs "<path>\ALL TENANT AUTOMATED TRACKER.xlsx"   # on the real tracker (never commit it)
```
UI tests (`tests/ui-smoke.cjs`, `tests/ui-features.cjs`) need Playwright: `PW=<dir with node_modules/playwright>`.

## Rules

- **Never commit real workbooks** — `.gitignore` blocks `*.xlsx` except `tests/fixture.xlsx`.
- Run the harness before every commit; one commit per change with a clear message.
- The OneDrive folder `SCRIPTS\DASHBOARD\QBR-Dashboard_1.9.0\qbr-app\` is the published copy people open;
  this repo is the source of truth.
