RAKSO EDUCATION — M365 TENANT QBR DASHBOARD (Offline)   ·   v1.21.0
==================================================================

** LIVE / production version: 1.21.0  (folder: QBR-Dashboard_1.21.0), promoted 2026-10-05.
   QBR-Dashboard_1.9.0 is now the ROLLBACK — don't edit it. Fixes are applied IN PLACE
   here and recorded as the dated entries below; the number is NOT bumped per patch. A new
   number is cut only for a deliberate major release. **

UPDATE — 2026-10-06 — v1.29.0 Save to Excel works after the file was edited in Excel (no re-linking)
  - If a Direct-save workbook was changed in Excel after you loaded it, Save no longer refuses
    and downloads a copy. The dashboard reads the file as it is now and puts your edits into it:
      · Edits made in Excel are kept.
      · Your edits go to the right place even if columns were added or moved in Excel (matched
        by the column heading) or rows were sorted or added (matched by serial number, school,
        item ID, ticket number…).
      · If the SAME cell was changed in Excel and in the dashboard, a short list asks which
        value to keep (Excel or dashboard) — or Cancel to save nothing.
      · An edit that can't be placed (its row was deleted, or its column heading renamed in
        Excel) is listed and not saved, so nothing lands in the wrong cell.
  - After saving, the dashboard reloads the file, so you see the changes made in Excel too.
  - ↻ next to a Direct-save workbook now reloads it straight from the file — no file picker.
  - If the file is open in desktop Excel, Save says "close it there, then click Save again";
    nothing is written and your edits are kept (no extra download).
  - If OneDrive updates the file while you save, the dashboard notices and merges again.
  - Tip: avoid saving from the dashboard while someone is typing in the same file in Excel
    Online — OneDrive may then keep both versions as a conflict copy.

UPDATE — 2026-10-06 — Audit editor, guided audit wizard, exempt schools (v1.26–1.28)
  - NEW PAGE: Audit › Risky sign-ins editor. Load RISKY_USERS_AND_DOMAIN.xlsx (12 month sheets):
      · Edit each school's organization, risky-user count, domain health and reference link.
      · Add school rows, or paste a row copied from Excel.
      · Search, filter (All / Audited / Not audited / Incomplete / Exempted) and month-over-month
        changes in risky users.
      · A large progress bar beside the Audit heading shows "N / total schools audited" for the
        month, colour-coded red → amber → blue → green as it fills, with the number exempted.
      · If the workbook is changed in Excel while you work, the dashboard warns you and lists the
        rows to review before saving.
  - GUIDED AUDIT WIZARD: pick school + month, then step through risky users, domain health,
    storage (paste 7 values) and usage (paste 19 values; totals are checked), then the next school.
    One "Save all to Excel" at the end writes every workbook you changed.
  - EXEMPT SCHOOLS: mark a school exempt with a reason (No tenant access / No GDAP / No admin
    access / Other) for this month or all 12 months. Saved in columns K–L (EXEMPT, EXEMPT REASON).
    Exempt schools don't count against progress and the wizard skips them.
  - DIRECT SAVE is now a checkbox next to each loaded file (replaces the top "Link file" button).
    Ticking it the first time asks you to pick the file; unticking stops saving into it.
  - Fewer false "changed outside the dashboard" warnings: the file's content is compared, so a
    OneDrive sync that only touches the date no longer blocks saving.
  - Links in saved workbooks no longer pick up extra "&amp;" each time. Links already damaged by
    older saves are repaired on the next save, and reference links you add are written correctly.
  - The audit workbooks save with their formatting kept, like every other "Save to Excel".
  - The consolidated ALL TENANT AUTOMATED TRACKER is still never written by the dashboard.

UPDATE — 2026-10-06 — "Save to Excel" keeps your workbook's formatting
  - Save to Excel now writes only the cells you changed and leaves the rest of the file as it
    was. Colours, fonts, column widths, conditional formatting, drop-down lists, Excel tables,
    charts, comments, links and macros (.xlsm) are kept. The save message ends in
    "· formatting kept" when this engine was used.
  - New rows placed directly under an Excel table extend the table, and a new "Batch Code"
    column is added to the table.
  - If the file contains something this engine can't edit safely, nothing is written to it.
    A separate copy downloads instead, and the message gives the reason.
  - To switch back to the old save (browser console):
    localStorage.setItem("qbr-save-engine","legacy")
  - Also fixed: a second "Save to Excel" in the same session could report "No changes to save"
    and leave edits to existing rows unsaved until the page was reloaded.

UPDATE — 2026-10-05 — Edit batch / Bulk edit now show every field
  - Edit batch and Bulk edit use one form with all the unit fields, grouped as:
      · Status & deployment: Status, School/client, Date delivered, DR #, SQ, Batch code
      · Unit details: Model, Category, Description, Brand, Supplier, Condition
      · Warranty: Warranty start, Warranty end, Warranty years (end = start + years when no end
        date is typed)
      · Contact: Contact person/owner, Address, Contact details
  - Blank fields are left as they are; type a single "-" to clear a text field.
  - Edit batch now also has "Preview changes" before Apply, like Bulk edit.
  - Asset 360 › Edit (one unit) now also has Batch code and SQ.

UPDATE — 2026-10-05 — Bulk select & bulk edit, "Assigned, not delivered" flag, look up by DR / SQ / batch
  - INVENTORY › ASSETS now has a tick box on every row. Tick units, or filter the list (client,
    flag, status, batch, DR, SQ) and click "Select all N shown", then "Bulk edit".
  - Bulk edit can change: Status, Date delivered, School/client, Owner, DR #, SQ, Batch code,
    Warranty start/end and Condition. Blank fields are left as they are.
      · "Preview changes" lists every unit with "from → to" for each field. Nothing is saved
        until you click Apply. Changes are kept and go into the workbook on "Save to Excel".
      · STATUS is worked out from the data, so it's saved through the real columns:
        Deployed = client + Date Delivered (set to the date you pick when it's empty) ·
        In Stock = Date Delivered cleared (tick "Also clear the client" if needed) ·
        In Repair / Retired = Condition. A unit with an open ticket stays In Repair; the preview
        notes it.
      · SQ adds a new row to the 04 RAKSO INV. sheet; older rows stay as history.
  - NEW FLAG "Assigned, not delivered": units that have a client but no Date Delivered, so they
    still show as In Stock. Click it to list them, select all, and fix them with Bulk edit.
  - LOOK UP (top of Inventory) now finds DR #, SQ and batch codes as well as serials. Pick a
    suggestion, or type the full value and click Open, to list those units; the "✕" chip clears it.

UPDATE — 2026-10-05 — Batch codes: find and bulk-update the units of one batch
  - Every Tag batch now gets a batch code, e.g. B-20261005-01 (date + running number for that day).
    It's filled in automatically; type your own, or pick an existing code to add units to that
    batch, or clear it for no batch.
  - The code is saved on each unit. The first time you use "Save to Excel", a "Batch Code"
    column is added at the end of the 02 DEVICES sheet (no other column moves).
  - INVENTORY › ASSETS has a new Batches list (code, date, school, number of units):
      · "Show units" filters the asset list to that batch (also via the new Batch filter).
      · "Edit batch" changes school, owner, DR #, warranty start/end or condition for all units
        of the batch at once. Leave a field blank to keep it; untick units to leave them out.
  - STOCKTAKE can now check a whole batch: choose "Compare against: A batch code".
  - Asset 360 shows the unit's batch code.
  - Fixed: choosing "All types" or "All statuses" again in the asset filters showed an empty list.

UPDATE — 2026-10-05 — Inventory fixes: safe Save to Excel, batch tagging, stocktake, serial links
  - SAVE TO EXCEL NO LONGER BREAKS .xlsm FILES. Saving into a linked macro workbook (.xlsm) used to
    write it in the wrong format, so Excel refused to open it ("file format or file extension is
    not valid") and the macros were lost. It now saves in the right format and keeps the macros.
    Every save is checked first; if anything looks wrong, the linked file is NOT touched and a
    separate copy downloads instead.
  - If the linked file was changed in Excel after you linked it, the dashboard no longer overwrites
    it. Your changes download as a separate copy; click "Link file" again to keep saving directly.
  - A file damaged by the old version: restore it from OneDrive/SharePoint Version history (keeps
    macros), or copy it and rename the copy from .xlsm to .xlsx to see the data.
  - SCAN › TAG BATCH now shows one review for the whole batch: new units are registered, units
    already in inventory (in stock, same school or another school) are updated instead of skipped,
    and units scanned twice count once. Units moving from another school are highlighted, and each
    row can be set to Skip. School, SQ, DR #, owner and date are entered once for all of them.
    Warranty start is filled only where it's empty, unless you tick "overwrite".
  - KNOWN WARRANTY END DATE: in Tag batch, choose "Known end date", pick the date and whether it
    applies to all units or only those without one. A unit that ends on a different date gets its
    own date in that row's "Warranty end" box. Register assets also has a "Warranty end (if known)"
    field. Warranty years are worked out from the start date, so expiry alerts stay correct.
  - SCAN › STOCKTAKE (new, check only): scan everything at a school and compare it with the
    inventory — found, missing, belongs to another school, not in inventory. Nothing changes unless
    you apply a fix (move here / register). "Export stocktake" downloads the list as Excel.
  - "Add anyway" for a serial that's already in inventory is now under "Advanced".
  - TICKETS: the New ticket form warns when a serial isn't in inventory and suggests close matches
    (e.g. PF-4J4PRJ → PF4J4PRJ). Serials that aren't registered show a "Not in inventory" tag
    instead of a link that went nowhere, and opening such a link shows a clear message with
    "Register this unit".

UPDATE — 2026-10-05 — New name and logo: RCT OpsDesk
  - The app is now called RCT OpsDesk. The top bar shows the new logo and "RCT OpsDesk", and the
    browser tab reads "RCT OpsDesk — IT Operations Hub" with the logo as its icon.
  - CHANGE THE LOGO YOURSELF: click the logo (top-left) → "Upload new logo…" and pick an SVG,
    PNG, JPG or WebP (max 512 KB; square works best). "Reset to default" brings the standard logo
    back. An uploaded logo is saved in this browser only — other PCs keep the standard one.
  - To change the standard logo for everyone, replace qbr-app\assets\rct-opsdesk-logo.svg with the
    new file (same name).
  - Nothing else changed: pages, filters, Inventory, Scan, exports and the version number are the same.

UPDATE — 2026-10-05 — Welcome screen now shows which files to upload:
  - Before any data is loaded, the start screen lists what to drop:
      1. Start here (required): ALL TENANT AUTOMATED TRACKER.xlsx — the master data for the
         Overview, Security, Risky sign-ins, Adoption, Storage, Canva and Postmaster pages.
      2. Optional: USER MANAGEMENT, DOMAIN REGISTRATION TRACKER, the Lenovo Inventory workbook,
         ETG_PRINTER_INVENTORY — each card shows the sheet names it needs.
      3. Loaded inside the app instead: the MS Forms ticket export (Support tickets › Import
         Forms file), label photos (Scan) and Link file / Save to Excel.
  - New "Choose files…" button in the drop area (same as Upload at the top right).
  - Files are still recognised by their sheets, so renamed files keep working. Old separate
    exports (Usage, Storage, Canva, Security, Postmaster) still load but are no longer listed.

UPDATE — 2026-10-05 — v1.21.0 is now the live version (1.9.0 becomes the rollback):
  - Open QBR-Dashboard_1.21.0\qbr-app\index.html from now on (update any shortcut or bookmark
    that still points at the 1.9.0 folder). The header now shows v1.21.0.
  - If your remembered data doesn't reload automatically, upload the workbooks once. Entries
    saved in the browser are kept per workbook file and reappear when that same file is loaded.

UPDATE — 2026-10-05 — Inventory v1.21.0: tickets, supplies, saving to Excel, School 360 hardware
  - INVENTORY now has three sub-pages: Assets | Support tickets | Supplies.
      · Assets: tiles per unit type (in stock / deployed), deployable desktop sets, a per-model
        breakdown, "Total Inventory" KPI, a red "Duplicate serials" flag, a searchable client
        filter, and serial lookup that suggests matches as you type.
      · Support tickets: ticket numbers like PF62SDPW-20251005, statuses (Open, In Progress,
        Waiting for parts, Escalated, Resolved), priority, requester, follow-up notes, related
        tickets. "Import Forms file" loads the MS Forms ticket export — you see a preview first,
        and rows already imported are skipped.
      · Supplies (new): printer consumables from the ETG printer inventory workbook — items,
        received/used transactions, stock levels, and flags for items to reorder.
  - TICKET 360 (new): click a ticket to see its full timeline, status and related tickets.
  - ASSET 360: "✏️ Edit" to correct a unit's details (the serial can't be changed), a sticky
    "Back to Inventory" bar, and a "Check on Lenovo ↗" warranty link.
  - Links like …index.html#ticket/<ticket-no> or #asset/<serial> open that page directly.
  - SCHOOL 360: a Hardware tile and section (laptops, desktops, open tickets, warranties
    expiring soon) when the school has inventory records.
  - SCAN: click a photo to type the serial while looking at it; tick several rows and use
    "Tag batch" to assign them to one school at once; the same photo added twice is skipped.
  - SAVING:
      · Your entries now survive a page refresh (kept in this browser).
      · Chrome/Edge: "🔗 Link file" once, then "💾 Save to Excel" writes your changes straight
        into the original workbook — formulas and layout stay; bold/colours may be lost.
      · Other browsers: use Export to download an updated copy.
  - PRIVACY: entries kept in the browser include names and contact details from tickets.
    On a shared PC, clear this site's data (browser settings) when you're done.

UPDATE — 2026-10-04 — Inventory & ticket tracking + label Scanner (new "Inventory" section)
  - NEW "INVENTORY" PAGE (left navigation, new "Inventory" group): laptops, desktops and
    monitors with their support tickets, built from an inventory workbook you upload together
    with (or instead of) the tracker. Sheets are found by name: 02 DEVICES, 03 TECH SUPPORT LOGS,
    04 RAKSO INV., 06 PIPELINE, 07 PURCHASE ORDER.
      · KPIs: fleet size, % deployed, in repair, open tickets, average days to resolve.
      · Flags you can click to filter: tickets open > 7 days, "lemon" units (3+ tickets),
        warranty expiring within 90 days / expired, ticket serials not in the device list,
        signed-but-not-delivered pipeline items.
      · Register assets, deploy units to a client, log and resolve tickets, change a status.
        Changes stay in this session; "Export inventory workbook" downloads an updated .xlsx
        with an EDIT LOG sheet (the dashboard never overwrites your original file).
      · "Look up asset" finds a serial; clicking any asset opens its ASSET 360 page
        (details, warranty, deployment, ticket history, recommended actions).
      · The Inventory item in the navigation shows the number of open tickets.
  - NEW "SCAN" PAGE: upload photos of box or unit labels — the serial number, model and
    product key are read from the barcodes and the printed text. Results stay editable.
      · Works fully offline and when the dashboard is opened by double-click.
      · Barcodes are read on Windows too (built-in browser detectors only exist on Mac/Android).
      · Straightens tilted photos and enlarges small ones before reading the text.
      · Existing serial → warning, then its Asset 360 page. New serial → "Intake" fills the
        register form, "Tag" assigns it to a client (Client, Date, Purchase location, SQ, DR #,
        Assigned to).
      · Intake and Deploy forms have "Scan box label" / "Scan unit" buttons.
      · The first scan in a session takes a few extra seconds while the reader starts.
      · Photos stay on your computer, in memory only — nothing is uploaded or saved.
  - Asset 360 and Scan are left out of Export Deck and Export Images, like School 360.

UPDATE — 2026-10-03 — School 360 page + Tenant Health email-score fix
  - NEW "SCHOOL 360" PAGE (left navigation, just under Overview): everything about ONE
    school on a single page — handy for customer calls and school-level reviews.
      · Pick a school by typing in the School box (it opens on the school you filtered to,
        otherwise the school with the weakest Tenant Health score).
      · Header: organization, Tenant Health score and band, health rank, weakest area.
      · Six KPI cards: Tenant Health, Risky Users, Security Defaults, M365 Usage, Storage
        Used, Email Reputation.
      · Risky sign-ins with a monthly trend chart, share of the portfolio and rank.
      · RECOMMENDED ACTIONS for that school, worst first (Critical / High / Medium / Low) —
        e.g. Security Defaults not enabled, high risky sign-ins, storage at 80%/90%, domain
        expiring, email reputation issues, no GDAP, weak sign-in methods, SSPR off, low
        usage, Canva certificate expiring within 60 days.
      · Cards for Security & identity, Tenant & domain, Microsoft 365 usage, Storage
        (with per-quarter history), Canva, Email reputation, User management, and a Data
        coverage card showing which workbooks have this school ("No data" is never shown
        as 0). Each card's "Open page ›" jumps to the full page.
      · The Quarter chips apply; Organization/School filters don't (you pick the school here).
      · "Print this school" = Export Tab for this page. "Show this school across the
        dashboard" sets the School filter and returns to Overview.
      · From Overview → Tenant Health detail panel, "Open School 360" jumps straight here.
      · School 360 is NOT included in Export Deck or Export Images (those stay portfolio-wide).
  - FIX — Tenant Health Index, email (Postmaster) score: it now uses each school's LATEST
    month with a usable reputation. Before, it kept the FIRST such month, so a school that
    went HIGH → "Issues detected" could still score as HIGH. On the current tracker no
    school was affected (index unchanged at 69); it matters as new months are added.

UPDATE — 2026-10-01 — Storage capacity alerts + "share of total" analytics
  - STORAGE CAPACITY FLAGS: tenants using 80% or more of their pooled storage are now
    flagged "Near capacity" (orange) and 90% or more "At capacity" (red), based on each
    tenant's latest snapshot in the current filter. You'll see them in:
      · a new "Capacity Watch" table at the top of the Storage page (worst first — click a
        row to open that tenant's detail),
      · a "Near Capacity (≥80%)" KPI card and a status pill in the storage tables,
      · the Storage count in the left navigation and an Overview priority flag.
    Note: if the Data Quality audit flags a quarter's capacity figures, verify before acting.
  - SHARE OF TOTAL on ranked reports:
      · Risky sign-ins — bars read "603 · 38%", a Share column, and a line such as "The
        largest school alone holds 38%. Top 3 = 56% · Top 10 = 78% of 1,598 risky users";
        organization bars show each org's share. Overview's Top 5 table has a Share column.
      · Storage — top consumers show their share of all managed storage (+ Share column).
      · Canva — top schools show their share of all Canva users.
      · Microsoft 365 usage — organization bars and the top/bottom tables show share of all
        active O365 users.
  - Risk pills (Critical / High / Medium Risk…) use a drawn dot instead of emoji.

UPDATE — 2026-09-30 — Phase 3: status lists, change-vs-last-quarter, sorting, tenant detail, clean printing
  - PRINTING FIXED (Export Tab and Export Deck): charts no longer blow up or overlap the
    cards below them. Each chart is printed from a fixed-size copy made at paper width, and
    rows of cards now break cleanly between pages. Export Tab also prints light charts even if
    you're in dark mode. (This problem existed in 1.8.5 too.)
  - OVERVIEW — "Security Posture" and "Domain Health" are now ranked lists: each status on its
    own row with its count, % and a bar, largest first, no legend to decode, and the two cards
    line up. Click a row to open that page already filtered to that status.
  - CHANGE VS LAST QUARTER on every page: pick a single quarter (e.g. Q2) and the KPI cards on
    Risky sign-ins, Tenant status, Canva, Microsoft 365 usage, Storage and Google Postmaster
    show a green/red chip such as "▼ 37% vs Q1" (green = better, red = worse, grey = neutral
    or no change). Snapshot pages (Security defaults, GDAP, User management, Domain
    registration) don't change by quarter, so they have no chip.
  - ATTENTION COUNTS in the left navigation: Risky sign-ins (schools over 100 risky users),
    Security defaults (not enabled), Tenant status (not managed), Domain registration
    (needing action) and Google Postmaster (reputation issues).
  - SORT ANY TABLE: click a column header (or press Enter on it) — click again to reverse.
    The sort stays while you filter or search.
  - TENANT DETAIL PANEL: on Overview → Tenant Health Index → Show detail, click a school to
    open a side panel with its score, band, weakest area and all six sub-scores, with links to
    the matching pages and a "Show this school across the dashboard" button. Esc closes it.
  - Overview "Priority Flags": the email item now counts every domain with a reputation
    problem ("Issues detected" as well as BAD) — it read 0 before because the tracker uses
    "Issues detected". The navigation item is now simply "Security defaults".
  - Wording: "Domain Status" is now "Tenant Status" everywhere on that page and in the Data
    explorer; the remaining emoji in buttons and messages were removed.

UPDATE — 2026-09-30 — v1.9.0 is now the live version (1.8.5 becomes the rollback):
  - Open QBR-Dashboard_1.9.0\qbr-app\index.html from now on (update any shortcut or
    bookmark that still points at the 1.8.5 folder). If your remembered data doesn't
    reload automatically in 1.9.0, upload the workbooks once.

UPDATE — 2026-09-30 — New look for cards, KPIs, tables and badges (Phase 2):
  Same numbers, same charts, same exports — cleaner Microsoft 365 styling throughout.
  - KPI CARDS: the label sits above the number, and a small status icon in the corner shows
    the meaning — green check (healthy), orange warning (attention), red alert (critical),
    blue chart (metric). The colored side bar is gone; the "vs Q1" change chips are pills.
  - TABLES: the column headers stay visible while you scroll down a long list; rows
    highlight on hover; numbers line up in columns.
  - BADGES & STATUS PILLS (Healthy, Critical Risk, Granted, Not Enabled, …) are now soft
    tinted pills that are easier to read, including the Tenant Health band and the "69" score.
  - ALERTS: the Data Quality warning banner and the blue notes are Microsoft-style message
    bars with an icon. The Overview "Priority Flags" use status icons instead of emoji, and
    the Quick Actions have an arrow.
  - Every text color now meets the accessibility contrast standard in light and dark mode.
  - Printed deck is more compact (about 45 pages instead of 50 on the current tracker).

RELEASE — 2026-09-30 — v1.9.0 · New app layout (Phase 1: shell)
  Same data, same charts, same exports — the frame around them is new.
  - LEFT NAVIGATION replaces the tab strip and the hover-only "Microsoft" menu.
    Every page is now one click away, grouped by what you're checking:
      Overview
      Security & identity .... Risky sign-ins · Security defaults & auth · GDAP partner access
      Tenant health .......... Tenant status · User management
      Adoption & capacity .... Microsoft 365 usage · Canva Education · Storage
      Domains & email ........ Domain registration · Google Postmaster
      Reporting .............. Executive report
      Data ................... Data quality audit · Data explorer
    "Domain Status" is now called "Tenant status" (it's tenant health, not the
    registrar) and "Full Data" is now "Data explorer". The page you're on is
    highlighted and shown in the breadcrumb above the filters.
  - Collapse the navigation to a slim icon rail with "Collapse" at the bottom of
    the sidebar (your choice is remembered). On a small window or tablet the
    navigation opens from the menu button at the top left.
  - ONE "Export" MENU at the top right holds all five exports — Export Tab,
    Export Deck, Export Images, Export Data, Export Clean — each with a one-line
    description. They work exactly as before.
  - Slimmer top bar (48px) so more of the dashboard fits on screen. The light /
    dark switch is the moon / sun icon.
  - Links to a page: add #risky, #sec, #storage, … to the address to open that
    page directly (e.g. index.html#risky).
  - Cleaner Microsoft 365-style look: lighter background, Fluent icons (no more
    emoji in the navigation), stronger text contrast, and an accessible orange for
    "attention" items. KPI rows and filters now reflow on narrow screens instead
    of squeezing.
  - Nothing changed in how data is read, counted, or scored: the Tenant Health
    Index, Data Quality audit, charts and exports are the same as 1.8.5.

UPDATE — 2026-09-25 — Domain Registration status chart is now a horizontal bar chart:
  - The "Registration Status" chart on the Domain Registration tab changed from a
    donut to a horizontal bar chart — one bar per status, sorted from most to
    least common, with the count on each bar. It's much easier to read than the
    donut's many small slices.

UPDATE — 2026-09-25 — Domain Registration uses the uploaded workbook (built-in default removed):
  - The Domain Registration tab now works like every other source: it fills in
    only when you upload the domain workbook (e.g. QBR_domainreg.xlsx), the
    DOMAINREG badge turns green when it's loaded, and the tab goes blank / the
    badge greys out when you remove the file.
  - (An earlier build tried baking the domain file in as an always-on default;
    that made the DOMAINREG badge stay green even after removing files, so it was
    removed in favor of the consistent upload behavior above.)

UPDATE — 2026-09-25 — Domain Registration report rebuilt (reads the tracker's own status):
  - The Domain Registration tab now reads each domain's Status and Remaining Days
    straight from the workbook (as your team sees them in Excel) instead of
    recomputing dates — so it can't drift from the source. It also now loads the
    per-quarter export whose sheet is named "Q3" (the old version only recognized
    a sheet with a "Registration status" header, and would have shown blank).
  - New status cards at the top: Total Domains, Active, Expiring ≤60 days, Expired,
    Renew / Delete, and Errors / Invalid / Pending — plus a status donut.
  - Every domain now gets a clear operational status instead of "No Data":
    Active, Expiring Soon (Urgent ≤30 days / Watch ≤60), Expired, For Renewal,
    For Deletion, Deleted, End Contract, Invalid Domain, Pending, Error, or Not
    Registered — with an Action / Remarks column that tells you what to do
    ("Pay now — deletion pending", "Renew within N days", "Re-check lookup",
    "Fix domain name"). "Error"/"PENDING" values in the Registration column are
    now understood as lookup states rather than blanks.
  - Filter the table to "⚑ Action Needed" (or any single status), and it sorts
    worst / soonest-to-expire first. The overview alert now reads "Domains Needing
    Action". The "Days" column shows negatives (overdue) in red.

UPDATE — 2026-09-23 — Filter the Security & Risk Detail table by authentication method:
  - On the Security Defaults tab, the "Security & Risk Detail" table now has an
    "Auth method" dropdown at its top-right. Pick a method (MFA, Passkey/FIDO2,
    Email OTP, SMS, Temporary Access Pass, Authenticator, Voice) to show only the
    schools that use it, or use a posture preset:
      • Phishing-resistant  — schools with Passkey/FIDO2 or Temporary Access Pass
      • Weak-only           — schools whose only factors are weak (no MFA/passkey/authenticator)
      • Email OTP only      — schools relying solely on Email OTP
      • No MFA              — schools with methods but no MFA
    The filter narrows this table only; the KPI cards and charts above stay
    portfolio-wide. Use the "Show all" toggle to page past the first 25 rows.
  - The "Security Default" column in this table now reads "Not Enabled" (instead
    of the raw "DISABLED"), matching the neutral wording used everywhere else.

UPDATE — 2026-09-22 — Per-Tenant Health Detail (drill-down under the Health Index):
  - The Tenant Health Index card now has an expandable "Per-Tenant Health Detail"
    table. Click "Show detail" to see every scored tenant ranked weakest-first, with
    its overall index, band, and the six underlying scores (Security, Risky, Adoption,
    Domain, Email Reputation, Storage). Each score is color-coded and the tenant's
    weakest area is outlined, so you can see at a glance what is pulling a school down.
  - Filter the list to one band with the chips (All / Healthy / Attention Needed /
    Critical), or click a segment of the distribution bar to jump straight to that band.
    A "—" means that signal is missing for the tenant (left out of the score, not zero).
    The panel follows the Q1–Q4 / organization / school filters like the rest of Overview.

UPDATE — 2026-09-22 — Chart labels no longer overlap:
  - On charts with long category names (e.g. "Health by Organization", where full
    org titles like "SCHOOLS DIVISION OFFICE (SDO) OF TAGUIG…" ran into each other),
    long labels are now trimmed with a "…" so every bar still gets a label but they
    no longer collide. Hover a bar to see the full name. Short labels (Q1–Q4, service
    names) are unchanged.

UPDATE — 2026-09-22 — Tenant Health Index (Overview):
  - New card at the top of the Overview tab: a single 0–100 "Tenant Health Index"
    that blends six signals per school — Security, Risky Sign-ins, O365 Adoption,
    Domain Health, Email Reputation and Storage — into one executive number, with a
    band (Excellent ≥95 / Healthy 80–94 / Attention Needed 60–79 / Critical <60) and a
    bar showing how many tenants fall in each band.
  - If a school is missing a signal, that signal is left out of its score (weights
    re-balance) rather than counted as zero — a data gap never looks like a failure. A
    tenant needs at least two signals to be scored. The card follows the Q1–Q4 / org /
    school filters like the rest of Overview.

UPDATE — 2026-09-18 — Neutral wording for Security Defaults ("Not Enabled"):
  - Across the Security Defaults tab, the Risky Users comparison, the Executive Report
    and the filter dropdown, the status previously shown as "Disabled" now reads
    "Not Enabled", and value-laden labels ("Protected / Unprotected", "OFF") were
    replaced with plain descriptions. "Enabled", "Conditional Access" and "Not Managed"
    are unchanged. This is wording only — the numbers, colors, and the Security Default
    filter all behave exactly as before.

UPDATE — 2026-09-18 — Chart labels + storage detail follows the quarter filter:
  - Bar charts now show EVERY column's category label. Long, rotated names (e.g. on
    "O365 Active by Organization") were being silently dropped by the chart library
    when they got crowded — RCAMES, SPCEM, MAPSA, CEAP and others went unlabeled.
    All labels are now forced on across every bar chart in the app.
  - The "Per-Tenant Storage Detail" (its KPIs + "Storage Composition vs Capacity")
    now follows the Q1–Q4 filter. Before, it always showed the latest quarter's
    snapshot no matter which quarter you picked. Now: pick Q1 and you see that
    tenant's Q1 storage; "All" shows the most recent quarter as before. If a tenant
    has no storage rows in the selected quarter, it says so instead of showing stale
    numbers.

UPDATE — 2026-09-18 — "By Organization" charts no longer dump into "Unspecified":
  - The "O365 Active by Organization" chart (Usage tab) was throwing a large
    "Unspecified" bar. Cause: the USAGE_REPORT sheet leaves the Organization column
    blank for 168 of its 436 rows, and this one chart grouped by that blank cell alone
    instead of falling back to the school's known organization (the way every other
    view already does). All 168 blank-org rows are in fact resolvable, so "Unspecified"
    now disappears and those users land under their real orgs (Adamson, ALCU, MAPSA,
    LASSAI, …).
  - Applied the same organization fallback to the two sibling charts — "Risky Users by
    Organization" and "Health by Organization" — so the identical bug can't surface
    there either.
  - Note: this fixes the CHART only; the underlying blank Organization cells in
    USAGE_REPORT are still worth filling in at the source (the Data Quality tab lists
    the tenants with no organization).

UPDATE — 2026-09-18 — Counting-accuracy fixes + storage parsing:
  - Rows with a blank or unreadable Quarter cell are no longer counted in every
    quarter at once. Previously such a row matched Q1, Q2, Q3 AND Q4 simultaneously,
    inflating storage, usage and Canva totals whenever a quarter chip was selected.
    They now appear only under "All" quarters, and the Data Quality tab reports how
    many there are so the source cells can be fixed.
  - Fixed the related double-count: with a single quarter selected, the app skipped
    its per-school de-duplication (it assumed one row per school), so any repeated
    school row was counted twice in totals and Top-10 rankings. De-duplication now
    runs in every filter scope.
  - CHECKED AGAINST THE CURRENT TRACKER: both of the above are latent on today's
    data — that workbook set has no unattributed-quarter rows and no repeated school
    rows, so every headline total is UNCHANGED. The fixes stop the miscount the first
    time a blank Quarter cell or a duplicated row reaches the tracker.
  - Storage "Used" now reads a FOURTH estimate: the used figure on the left of the
    USAGE cell ("1.35 TB of 100.15 TB used"), which carries its own explicit unit.
    With four independent readings the median can out-vote a mislabeled cell instead
    of being stuck between two. Concretely — Our Lady of Guadalupe Minor Seminary Q2,
    whose "current storage" cell says "1.35 GB" but means 1.35 TB, read as 0.68 TB
    and now reads the correct 1.35 TB. Q2 storage rises 171.11 TB -> 171.79 TB.
  - Where only TWO estimates exist and they disagree by more than 25%, the app no
    longer averages them (one bad cell used to drag the answer halfway). It resolves
    the case it can actually diagnose — a TB value labelled GB, a ~1024x gap — and
    otherwise reports no reading rather than a confident wrong number. No row in the
    current tracker loses its reading as a result.
  - Fixed invisible text marks (left-to-right marks) in the USAGE column defeating
    the capacity fallback: rows where the numeric Total column was missing or absurd
    silently got no capacity at all, because "of 100.15 TB" never matched.
  - A Postmaster reputation value we do not recognise is still shown verbatim, but is
    now listed in the Data Quality tab's vocabulary violations — previously it slipped
    past every colour and severity map and read as unremarkable grey text.
  - The Data Quality tab now says "Audit incomplete" when the workbook's raw rows are
    unavailable, instead of reporting a clean bill of health it never actually checked.
  - Housekeeping: two unreferenced duplicate scripts (js/app-1.js, js/excel-loader-1.js)
    were moved out of the app to ../_superseded-source/. They were never loaded, but one
    held a stale pre-Postmaster-fix palette that could mislead future reviews.

UPDATE — 2026-09-17 — Google Postmaster tab fixed + name spelling repair:
  - Google Postmaster tab now actually works with the tracker. Before, every domain
    showed as "Not enough data" and the verification/health cards were blank. The tab
    now reads the real Google states:
      • "Issues Detected" — Google flags a mail-deliverability problem (the key
        signal; 27 domains this period). Review SPF/DKIM/DMARC & sending practices.
      • "Verify to See Health" / "Not Verified" — the domain must be verified in
        Postmaster before Google will report its reputation.
      • "Not enough data" — normal for low-volume domains.
  - Verified / Unverified counts, the reputation distribution, the watchlist (worst
    domains first, with month-by-month history) and the executive report now populate
    from the tracker's own Status + Reputation columns — no separate file needed.
  - The "Cloudflare DNS" card shows a short note when a workbook has no DNS column
    (that data ships only with the standalone Postmaster export).
  - Fixed a school-name spelling glitch: "De La Salle University - DasmariÃ±as" now
    renders correctly as "Dasmariñas" everywhere (a text-encoding repair). The Data
    Quality tab also flags any such source names so the tracker itself can be cleaned.

UPDATE — 2026-09-17 — v1.8.5 — Canva "Active Schools by Quarter" tile:
  - New 4th tile on the Canva Adoption tab (below "Adoption by Quarter") — a column chart
    of the number of ACTIVE SCHOOLS per quarter (schools with Canva users > 0), labeled
    above each column, e.g. Q1 = 52, Q2 = 64. It responds to the Q1–Q4 filter chips: pick
    one quarter and it shows that quarter's count as a stat; pick several and it shows the
    columns plus a quarter-over-quarter growth line ("Q1 → Q2: +12 schools, +23%").
  - A quarter that is only partially entered (data still coming in) is labelled "in
    progress" and a quarter with nothing recorded shows "no data yet" — neither is drawn as
    a misleading 0, so the growth story stays honest.

UPDATE — 2026-09-17 — v1.8.5 — Phase 1B: clean-data export + Executive Snapshot:
  - New "Export Clean" button (header) → downloads QBR_Clean_Dataset.xlsx: one tidy,
    CORRECTED row per tenant × quarter (median storage, resolved organization, coverage
    flags, blanks for "no data" — never 0), joining every sheet by the master key. This is
    the analytical feed for the executive layer. Engine: QBR.buildCleanRows(model) /
    QBR.CLEAN_COLS in excel-loader.js.
  - Executive Snapshot (separate self-contained file, Rakso_QBR_Executive_Snapshot.html):
    an offline, shareable one-page executive dashboard built from the clean dataset —
    quarter selector, 6 headline KPIs with quarter-over-quarter deltas, risky-by-quarter,
    security-posture + domain-health donuts, Top-10 storage / adoption / Canva, and a
    sortable tenant table. Chart.js is inlined and the tenant data is embedded, so it opens
    offline and nothing leaves the device.
  - v1.8.5 also contains the Phase 1A Data Quality self-audit (below) — this is the first
    numbered release since 1.8.4 (the 1.8.4.1 folder was a working/staging copy).

UPDATE — 2026-09-17 — Data Quality self-audit tab:
  - The "Data Quality" tab now runs a full source audit LIVE on every upload
    (new js/data-quality.js → QBR.audit(model)). It shows:
      · a Source-Hygiene score (0–100, graded) — reflects the RAW workbook; the
        dashboard still corrects everything at read time.
      · trust banners that auto-detect "Q3 storage capacity unreliable" and
        "Postmaster reputation source not loaded" (embedded sheet has no HIGH/MED/
        LOW/BAD — upload the standalone GOOGLE POSTMASTERTOOLS.xlsx).
      · a coverage matrix (distinct schools per sheet vs the 122 master union) with
        the "missing from sheet" lists and the Canva-only tenants.
      · a "Source cells to fix" register with exact cell refs, current value, a
        suggested fix and a type tag (unit mislabel / magnitude typo / impossible
        capacity / missing capacity / service check), plus a one-click
        "Export Fix List (.xlsx)" button (writes QBR_Source_Fix_List.xlsx).
      · controlled-vocabulary violations (e.g. "Healhty" ×6), status-words-in-a-
        numeric-column (72 risky cells), organization gaps, name variants, and
        placeholder/sentinel row counts per sheet.
  - Engine: excel-loader now stashes raw sheet rows on model.raw (in-memory only)
    so the audit can inspect source cells; parsing outputs are unchanged.

UPDATE — 2026-09-17 — Storage data integrity + Domain Status scoping:
  - Storage "Used" is now the MEDIAN of three source estimates — the "current storage"
    text, the "Used Storage(GB)" column, and OneDrive+Exchange+SharePoint. The median
    ignores a single corrupt cell in EITHER direction:
      · Our Lady of Lourdes Q1 — mislabeled Exchange "168.83 TB" (should be GB) no longer
        inflates Used to 172.6 TB; now correctly 3.89 TB.
      · Dr Yanga's Q2 — "Used Storage(GB)" typo 1250 (should be 12500) now reads 12.5 TB.
  - Capacity (Pooled) guarded against impossible values (7–8 digit Q3 totals); Utilization
    recomputed from the corrected Used/Total; any per-service value larger than capacity is
    dropped so the composition bar can't draw a false segment.
  - Storage text column now honors each cell's OWN unit (the sheet mixes "85.92 TB" and
    "440.62 GB") instead of coercing GB→TB.
  - Domain Status tab is now scoped to tenants that actually have a RISKY_USERS_AND_DOMAIN
    record. Canva-only tenants (which reach the Master School Dimension via the Canva sheet)
    are excluded rather than listed as "Unassigned / No Status". Reconciliation caption
    updated; count 122 → 110 domain-tracked tenants. Consistent with the Overview donut.
  - Source cleanup: 13 mislabeled/impossible cells identified in STORAGE_DATA (see
    "Storage_Cells_To_Fix" checklist). Dashboard tolerates them via the median + guards.

UPDATE — 2026-09-14 — Multi-quarter filter, Postmaster, risk & data-quality passes:
  - Global Quarter filter is now a MULTI-SELECT toggle-chip group [All][Q1][Q2][Q3][Q4].
    Flow metrics (risky sign-ins) SUM across the selected quarters; snapshot metrics
    (Canva, storage, security, reputation) de-duplicate to one row per school so they
    are not double-counted.
  - Multi-file upload persistence: uploading workbooks one-by-one accumulates them
    (dedupe by filename) instead of replacing; each source has a per-file Refresh.
  - Google Postmaster tab rebuilt: reputation mapping (blank / "-" → "Not enough data",
    "No data to display" kept verbatim), verification + DNS-status donuts,
    reputation-by-organization, plus Watchlist and Action tables.
  - Risky Users: new "Security Defaults vs Risky Users" section above the monthly trend —
    Enabled vs Disabled lines that reconcile to the monthly total, actual numbers on each
    point, green = Enabled / red = Disabled, with an Enabled/Disabled filter.
  - Terminology: "No Access" → "Not Managed" across every report.
  - Domain Status reconciled so every in-scope tenant lands in exactly one status bucket.
  - Organization casing auto-merged (e.g. Bulprisa → BULPRISA, Government → GOVERNMENT).
  - Tenant Management classification (GDAP-based: Managed / access-only / Not Connected /
    not onboarded).
  - Per-tenant drill-downs on Office 365 Usage and Storage (click a Top-10 bar or pick a
    tenant): detailed adoption / composition + per-quarter stacked bars with totals; the
    vague "Usage Distribution" chart replaced by "Average Adoption by Service".

UPDATE — v1.8.4 — Lowest Risk Schools fix + rankings, risk bands, nav grouping:
  - "Lowest Risk Schools" now shows the 10 lowest schools with at least 1 risky user
    (previously the bottom 10 were all 0, rendering as blank bars). A caption reports how
    many schools had 0 risky sign-ins (fully clean).
  - Top-10 / Lowest charts: rankings 1-10 prefixed, school names left-aligned.
  - Risky Users tab: charts renamed ("Top 10 Risky Users Per School", "Lowest Risky Users
    Per School"); "Risk Flags" table -> "Risky Users (Top Schools)" with a new auto-suggested
    Action column.
  - App-wide 5-band risk scale (0 No Risk · 1-15 Low · 16-50 Medium · 51-100 High · >100
    Critical) with colored 🔴🟠🟡🔵🟢 badges — used on the Risky tab, Overview, and Report.
  - Navigation grouped: six Microsoft views collapsed under a "Microsoft ▾" dropdown
    (Risky Users, Domain Status, Security Defaults, Office 365 Usage, Storage, User Management).
  - "Domain Health" renamed to "Domain Status" (tab + heading).
  - Domain Status table: new "Remarks" column sourced from the workbook's "Error Cause" column;
    blank Organization shows "Not connected" when the domain has No Access (no admin access),
    otherwise "Unassigned".
  - Security Defaults: "Disabled Schools Ranked by Risk" -> "Schools with Security Defaults
    OFF — highest risk first".
  - New "GDAP Access" tab (under Microsoft): lists every tenant with a Granted/None badge
    (default filter = Granted), relationship-ID count + the GUID(s), coverage donut and KPIs.
    Sourced from SECURITY_DATA "GDAP" column. Roles and Expiry columns are pre-wired — add
    "GDAP Roles" and/or "GDAP Expiry" columns to SECURITY_DATA and they populate automatically
    (Expiry shows Active / Expiring Soon (<=60d) / Expired badges).
  - Storage: added "Bottom 10 Storage Consumers — Lowest Usage" and a "Pooled Storage (TB)"
    column; Office 365 Usage: added "Bottom 5 Usage Report — Lowest Adoption".

HOW TO RUN
  1. Unzip anywhere.
  2. Double-click index.html (opens in your browser). No server/internet needed.
  3. Click "Upload Workbook(s)" or drag-drop your Excel file(s).
     - Works with ALL TENANT AUTOMATED TRACKER.xlsx and/or split exports
       (Usage, Storage, Canva, Security). Sheets are detected by name.
  4. Use the global filters (Quarter drives all views; Month filters Risky Users only).
  5. "Export PDF" prints the currently open dashboard (use browser "Save as PDF").

NOTES
  - Data is processed in-memory only; nothing is uploaded or saved. Re-upload each session.
  - Storage shown in TB (source GB rolled up). Thresholds: >25 HIGH RISK, >100 CRITICAL.
  - Dirty values (typos/casing) are auto-normalized; a Master School Dimension
    joins all sheets by school name. See the "Data Quality" tab for coverage.

FILES
  index.html · css/styles.css
  js/excel-loader.js  (parsing + master school dimension)
  js/chart-generator.js  (Chart.js wrappers)
  js/report-generator.js (rule-based executive narrative)
  js/app.js  (dashboards, filters, aggregation, PDF export)
  libs/  (Chart.js, SheetJS, Bootstrap — all local/offline)

UPDATE — new tabs:
  - Google Postmaster: domain reputation (HIGH/MEDIUM/LOW/BAD) + spam rate per school (respects Month filter).
  - User Management: upload USER MANAGEMENT.xlsx (SY 2025-2026 / SY 2026-2027). School-year readiness,
    Updated vs Pending, by-organization, readiness-by-SY, detailed status, and cross-module correlation score.
  - Two new filters (School Year, Status) apply to the User Management tab.
  - Tip: select BOTH the tracker and USER MANAGEMENT.xlsx together in the upload dialog.

UPDATE — Domain Registration tab:
  - Upload DOMAIN REGISTRATION TRACKER.xlsx (sheet CURRENT).
  - Parses "Registration status: CURRENT" timestamp (YYYY-MM-DD-HH-MM-SS).
  - Expiration = registration + 1 year (unless a Validity Period column exists).
  - Status vs your system clock: Registered (>30 days, green), Expiring Soon (<=30 days, red),
    Expired (past, dark red), No Data (no registration date, gray).
  - Table columns: School, Domain, Registration Date, Expiration Date, Days Remaining, Status.

UPDATE — polish pass:
  - Fixed: Canva (and other blank-org rows) now filter correctly by Organization via the Master School Dimension.
  - School & Organization filters are now type-to-search (start typing; non-match = All).
  - Single-series bar/column charts use distinct per-category colors (quarters fixed Q1-Q4 palette).
  - Darker light background + new Dark/Light theme toggle (top-right; remembers your choice).
  - Tab bar redesigned: bordered strip with a filled active-tab chip.

UPDATE — presentation export (three buttons, top-right):
  - "Export Tab"    : prints only the dashboard you are viewing (the old Export PDF).
  - "Export Deck"   : prints EVERY dashboard as one paginated QBR document, in
                      presentation order - cover page first (stamped with the filters
                      that were active, school count and date), then the Executive
                      Report, then risk -> security -> email -> adoption -> storage ->
                      readiness, with Data Quality last as an appendix. One dashboard
                      per page. Use the browser's "Save as PDF" in the print dialog.
                      Long tables are capped at 25 rows so they don't run for pages.
  - "Export Images" : renders all 33 charts to PNG at 3x resolution and opens a
                      gallery. RIGHT-CLICK any chart -> Copy image -> paste straight
                      onto a PowerPoint slide, or use its Download link to save a file.
                      Images are composited on white, so they sit cleanly on any
                      slide background.
  Notes:
  - Both exports temporarily switch to the light theme (dark charts are unreadable
    on paper and on a projector) and restore your theme afterwards. Your saved
    Dark/Light preference is not changed.
  - If your browser blocks the gallery pop-up, it opens inside the app instead.
  - Nothing is uploaded; images are generated in the browser like everything else.

UPDATE — v1.8.3 — table filters, stacked labels, export & UI fixes:
  - Domain Health "Health by Organization": each stacked segment now shows its count
    (tiny segments skipped to stay readable).
  - "All Schools" tables now have a status dropdown beside the search, styled uniformly:
      · Domain Health  -> filter by Domain Status (Healthy / Possible Service Issues / …).
      · Google Postmaster -> filter by Reputation (HIGH / MEDIUM / LOW / BAD / No Data).
      · User Management  -> filter Detailed Status by Updated / Pending.
  - User Management: Grade-Level Update & Extract Users "Pending" now render as bold plain
    text (no yellow fill); only the Status column keeps the colored badge.
  - Header "Loaded" files collapsed into a one-line toggle ("Loaded: N sources ▾") that
    expands the source chips on click — no more wrapping/overflow.
  - Export Images: the School-Year Readiness gauge now draws its centre value ON the canvas,
    so the "% Updated" text appears in the exported PNG (previously an HTML overlay was lost).
  - Domain Registration already includes its status filter (Expired / Expiring Soon / …).

UPDATE — v1.8.2 — Executive Overview + in-chart data labels:
  - New "Overview" tab (first / default landing): 6 hero KPI cards with quarter-over-quarter
    deltas, risky-users trend, Security Posture & Domain Health donuts, School-Year Readiness
    gauge, Top-5 Storage & Canva bars, Top-5 Risky table, clickable Priority Flags, Quick Actions.
  - In-chart data labels on EVERY tab: donuts/pies show value + %, bars show the value at the
    bar end / above the column, lines show the value at each point. Slices too thin for an inside
    label get a leader line to the number outside the ring (offline; no extra library).
  - Stacked / multi-series charts (Health-by-Org, Reputation-by-Org, Enabled-vs-Disabled trend,
    Updated/Pending-by-Org) are intentionally left unlabeled to stay readable.
  - Premium UI pass: layered shadows, hover lift, soft brand tints on KPIs, glassy header,
    gradient accents. Prints flat/clean. Labels also appear in Export Deck / Export Images.

UPDATE — pinned navigation:
  - The header, filter bar and tab strip now stay frozen at the top of the window
    while you scroll a dashboard, so you can switch tabs or change a filter from
    anywhere in a long report without scrolling back up.
  - The tab strip is now a single scrollable row (drag or shift+scroll it sideways
    if your window is too narrow to show all 11 tabs) instead of wrapping onto
    2-4 lines and eating the screen.
  - On a very short or very narrow window the bars would cover most of the view,
    so they automatically go back to scrolling normally. Widen or maximise the
    window to get them pinned again.
  - Printing and both exports are unaffected — the bars never appear on paper.
