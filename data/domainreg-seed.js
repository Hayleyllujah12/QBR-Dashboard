/* DISABLED 2026-09-25 — the built-in DEFAULT domain workbook was removed so the
 * Domain Registration tab behaves like every other source: populated only when a
 * domain workbook is uploaded, and blank (chip grey) when it is removed.
 * This file is intentionally empty (no QBR.SEED_DOMAINREG) and is no longer loaded
 * by index.html.
 *
 * To re-enable a built-in default later:
 *   1) node tests/embed-domainreg.cjs "PATH\TO\QBR_domainreg.xlsx"   (regenerates this file)
 *   2) re-add  <script src="data/domainreg-seed.js"></script>  before js/app.js in index.html
 *   3) restore the seed hooks in app.js (seedDRBuffer + processBuffers/restore/removeFile)
 *      — see the 2026-09-25 entries in claude/CHANGE_QUEUE.md. */
