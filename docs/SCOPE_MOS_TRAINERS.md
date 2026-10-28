# Scope — MOS issuance & analytics, Trainer schedule & portfolio

**Status:** SCOPE ONLY (decisions below confirmed with the requester 2026-10-28). No code yet.
**Base:** `origin/develop`. Features 1 and 4 write to Excel, so they build on the Excel write contract
(`docs/EXCEL_WRITE_CONTRACT.md`, branch `refactor/excel-write-contract-p1`) once it lands — not on the old save path.

## Decisions recorded

| Question | Decision |
|---|---|
| What's missing after printing? | Release/claim tracking, batch transmittal form, reprint/replacement log |
| "Per program" in MOS analytics means | Exam / certification (Word, Excel, PowerPoint, Excel Expert…) |
| Schedule source | `ETG Training Tracker` › `02 Sessions Tracker`; the team will enter upcoming sessions there |
| Schedule shows | Upcoming + past |
| MOS vs Onboarding split | Software = `MOS` (training + proctoring) vs Software = `Microsoft`; Canva and the rest → "Other" |
| Portfolio content | Courses they can teach, their certifications, delivery track record |
| Qualifications/certifications source | New sheet in the ETG tracker (`04 Trainer Profiles`), editable in the app |
| Where issuance is recorded | New columns on `MOS Certificate Log.xlsx` (one row per certificate stays the rule) |

## What exists today (so nothing is rebuilt)

- **MOS › Certificates** — which certificates were generated and printed; log columns already include
  `Printed`, `First/Last Printed`, `Printed By`, `Print Count`, `Remarks`.
- **MOS › Results** — school › section › student analytics, pass-rate trend, section comparison, weakest skill areas.
- **Trainers** — sessions from the tracker + MS Forms evals → NPS and Trainer Score (Sessions / Board / Summary views),
  trainer profile PDF.

---

## 1 · MOS › Printing & issuance

**Lifecycle per certificate:** Generated → Printed → Released → Claimed (Reprinted is an event, not a state).

**New Certificate Log columns** (appended to the right; existing columns untouched):
`Release Status` (Pending / Released / Claimed) · `Released To` (school contact or student) · `Release Date` ·
`Released By` · `Transmittal No.` · `Claimed Date` · `Reprint Reason` (last reason; every reprint also bumps the existing `Print Count`).

**Screens (inside MOS › Certificates):**
- *Issuance summary* — per school: printed / released / claimed / pending counts; filter by status.
- *Release a batch* — select printed, unreleased certificates (by school) → assigns one `Transmittal No.`
  (format `TR-YYYYMMDD-<SCHOOL CODE>-NN`) → writes Released To / Date / By → opens the **transmittal form**.
- *Transmittal form* (print to PDF) — school, transmittal no., date, table of name · certification · verification code ·
  certificate date, "received by / date / signature" lines for both sides.
- *Mark claimed* — per certificate or per whole transmittal.
- *Reprint* — requires a reason (damaged, wrong name, lost, other); increments `Print Count`, logs the reason.

**Rules:** the log is re-read before every write (as today), so edits made in Excel are kept. No student data goes into
browser storage (MOS privacy rule) — the transmittal is printed, not stored.

**Done when:** a batch for one school can go printed → released (with transmittal PDF) → claimed, a reprint with reason
shows in the log, and the summary counts match the log rows.

## 2 · MOS › Performance analytics (per school, per exam)

- **School × exam matrix** — attempts, students, pass rate, average score, **first-attempt** pass rate; cells link to the
  filtered Results table.
- **Per-exam view** — pass-rate trend by month / school year, score distribution, weakest skill areas for that exam.
- **Per-school view** — the same, with the school's exams side by side and its rank among schools.
- **Counting rules (must be fixed before build):** retakes — pass rate on *first attempts* and on *best attempt*, shown
  separately; a cell with fewer than 5 attempts is shown but marked "small group" (same rule as the Security posture fix).
- Uses only the score reports already loaded in Results — no new input file.

**Done when:** the matrix totals equal the Results table for the same filters, and every number can be clicked through
to the student rows behind it.

## 3 · Trainers › Trainer's schedule

**Source:** `02 Sessions Tracker` rows (Date, Trainer, Course Code, Software, Category, Client, Time Start/End, Duration,
Meeting Link). Grouping: **MOS** = Software `MOS` (incl. Category `Proctoring`), **Onboarding** = Software `Microsoft`,
**Other** = everything else.

**Views:** month/week calendar and an agenda list, **upcoming first**, past below; filters for trainer, group, client.
Per-trainer load: sessions and hours per week/month. **Clash check:** the same trainer booked in overlapping times on the
same day is flagged.

**Prerequisite (team process, not code):** upcoming sessions must be entered in `02 Sessions Tracker` ahead of time —
today the sheet has none dated after 28 Oct 2026.

**Data clean-up the parser must handle** (found in the 2026 tracker): 98 rows with blank Software, one row dated 1900,
course codes containing an invisible zero-width space (`MS-102​A`), Duration stored as text with spaces.

**Done when:** an upcoming session entered in the tracker appears in the calendar under the right trainer and group,
and a deliberate double-booking is flagged.

## 4 · Trainers › Trainer's portfolio

**New sheet `04 Trainer Profiles`** in the ETG tracker (one row per item, long format so it stays easy to edit in Excel):
`Trainer` · `Type` (Qualification / Certification) · `Course Code` (qualifications, from `03 Course List`) ·
`Certification` (e.g. MOS Expert, MCE, MIE Expert, MCT) · `Credential ID` · `Issued` · `Expires` · `Proof Link`.
Editable from the app (add / edit / remove rows), written with the format-safe save.

**Portfolio page per trainer:**
- *Can teach* — qualified courses grouped MOS / Microsoft / Other (titles from `03 Course List`).
- *Certifications* — with expiry warnings (expired, ≤ 60 days).
- *Track record* — per course: sessions, hours, clients, average NPS and Trainer Score (from existing eval data);
  courses delivered but **not** listed as qualified are flagged for the profile to be updated.
- Extends the existing trainer profile PDF with these sections.

**Done when:** a trainer's portfolio shows their qualified courses, certifications with expiry state, and per-course
delivery numbers that match the Sessions view for the same trainer.

---

## Open items (not blocking the first build)

1. **MOS pass rate per trainer** — linking students' exam results to the trainer who prepared them needs a rule
   (school + exam + date window after the session?). Left out of the track record until agreed.
2. **`03 Course List` clean-up** — its Product column mixes product names with session titles; the portfolio needs a
   clean Product value per course code.
3. **Who may mark "Claimed"** — anyone with the log, or only the person who released it?

## Suggested order and size

| # | Feature | Size | Depends on |
|---|---|---|---|
| 1 | MOS performance analytics | M | nothing new (read-only) |
| 2 | Trainer's schedule | M | team entering upcoming sessions |
| 3 | MOS printing & issuance | L | Excel write contract |
| 4 | Trainer's portfolio | L | Excel write contract, `04 Trainer Profiles` sheet, Course List clean-up |
