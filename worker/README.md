# NY Alumni Pilot Worker

This worker reads a **Legacy Advancement `.xlsx` export**, filters to your rollout
cohort (region + class years), and optionally enriches rows via the **RocketReach
API**. It is isolated from `index.html`, `matchGate.html`, `code.gs`, and the live
Google Sheet.

## How to scrape from Excel (team workflow)

This matches the agreed pilot focus:

| Priority | Field |
|----------|--------|
| Core columns | **Name**, **graduation year** (`Reunion Year`), **company**, **industry** |
| Outreach | **Email** first (RocketReach email when found, else Advancement email on file) |
| De-prioritized | Title (legacy titles are often 7–8 years stale) and phone |

### Two ways to enrich

1. **Batch from Excel (this worker — recommended for tri-state)**  
   Uses the RocketReach **API** (`worker/pilot.py`). One lookup ≈ one credit. Safe for
   automation, daily caps, and a audit trail in the output workbook.

2. **RocketReach website (manual “page scrape”)**  
   The team’s browser workflow on rocketreach.co. That path is **not** in this repo
   (an old `AlumniScraper` package exists elsewhere with hard-coded credentials — do
   not commit or run it from Git). Use the site only with your shared Pro account and
   stay within ~**500 lookups/day** until you confirm limits with RocketReach.

The dashboard (`index.html`) is for **review after** enrichment: load cleaned rows
into the Google Sheet **Records** tab, or upload a CSV for a one-off preview — it
does not batch-scrape an entire Excel file.

### Step 1 — Get a real Excel file

Use a **Legacy**-style export (same column names as Advancement), saved as a true
`.xlsx` (not an Excel Online HTML file renamed to `.xlsx`).

Required columns (exact header text):

- `Account Name` (or first + last name columns)
- `Person Account: Home State/Province`
- `Reunion Year`
- `Person Account: Current Employer Name` (helps RocketReach match)

Optional: `Person Account: Business Email`, `Person Account: Preferred Email`, `Industry`

Keep files **outside** this git repo (e.g. `C:\RA_C\private-data\`).

### Step 2 — Dry run (no credits)

Count how many rows match **tri-state** and **1980–1999** before spending lookups:

```powershell
cd C:\RA_C\MatchGate
.\.venv\Scripts\Activate.ps1
python .\worker\pilot.py `
  --input "C:\RA_C\private-data\TriState_Legacy.xlsx" `
  --limit 5000 `
  --region tri-state `
  --min-reunion-year 1980 `
  --max-reunion-year 1999
```

You should see `DRY RUN ONLY` and `People loaded: …`.

### Step 3 — Live batch (~500/day)

Start with **1 row**, then **5**, then up to **500** per day:

```powershell
python .\worker\pilot.py `
  --input "C:\RA_C\private-data\TriState_Legacy.xlsx" `
  --output "C:\RA_C\private-data\TriState_Scrape_2026-09-23.xlsx" `
  --limit 500 `
  --region tri-state `
  --min-reunion-year 1980 `
  --max-reunion-year 1999 `
  --sleep-seconds 1 `
  --live `
  --confirm-credits YES
```

- **`--limit 500`** — cap one run near the discussed daily lookup target.  
- **`--sleep-seconds 1`** — small pause between API calls.  
- Re-run the next day with a higher `--limit` or split input files by state if you
  track progress manually.

### Step 4 — Output sheets

After a **live** run, the output workbook includes:

| Tab | Purpose |
|-----|---------|
| **Run Summary** | Counts and **percentages** (`pct_rr_success`, `pct_rr_failed`, `pct_needs_attention`, review queues) |
| **Success** | Rows where RocketReach returned a match (`outcome` = FOUND) |
| **Needs review** | Not found, API error, or no RR match |
| **Scrape Focus** | All rows — meeting columns + `lookup_status`, `review_queue`, `outcome` |
| **Pilot Results** | Full legacy row + lookup metadata for audit |

Use **Run Summary** for Vani/Julianne stats. Filter **Scrape Focus** in Excel if you need finer buckets (e.g. `QUICK_REVIEW` vs `NEEDS_RESEARCH`).

**Website:** do not upload this `.xlsx` to bulk-run Refresh. Import cleaned rows into the Google Sheet **Records** tab; open **`index.html`** (live dashboard). Refresh is **one person at a time** for exceptions after batch enrichment.

Manual verification (~9 hours per ~4k records in the meeting estimate) happens in
Excel on **Scrape Focus** (or Pilot Results): fix company/industry/email, mark
`human_decision` / `human_comment` on the full sheet if you use that column.

### Step 5 — Into MatchGate

Evaluate still happens in the live dashboard (`index.html`): **Refresh via RocketReach** plus Accept / Flag / Reject. This worker is **not** the live RocketReach or review path yet.

- Import verified rows into the Google Sheet **Records** tab (production path), **or**
- Export **Scrape Focus** as CSV and use the dashboard upload for a local preview only.

### Rate limits

Plan for ~**500 lookups/day** (~2,500/week). If you hit caps, pause and upgrade the
RocketReach plan only after you have usage data from this worker’s `lookup_seconds`
summary at the end of each live run.

## Important safety notes

- The default mode is validation only. It makes no RocketReach calls.
- Live mode requires two explicit flags and may consume credits.
- It never writes back to the source workbook, UI, or Google Sheet.
- Keep alumni workbooks outside this repository.
- The imported `AlumniScraper` package contains hard-coded credentials. Do not
  execute or commit it. Rotate those credentials before reuse.

## Obtain a real workbook

The file currently named `NY_Legacy_Sample_100 .xlsx` is an HTML copy of Excel
Online, not an Excel workbook.

In Excel Online:

1. Open the 100-row NY sample.
2. Use **File > Save As > Download a Copy**.
3. Save the downloaded file outside Git, for example:
   `C:\RA_C\private-data\NY_Legacy_Sample_100.xlsx`.

## Setup (PowerShell)

```powershell
cd C:\RA_C\MatchGate
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r .\worker\requirements.txt
Copy-Item .\worker\.env.example .\worker\.env
```

Edit `worker\.env` and replace the placeholder with the RocketReach API key.
Do not add the RocketReach website username or password; this worker uses only
the API.

## Validate without spending credits

Tri-state cohort, 1980s–1990s reunion years (matches current scraping focus):

```powershell
python .\worker\pilot.py `
  --input "C:\RA_C\private-data\NY_Legacy_Sample_100.xlsx" `
  --limit 100 `
  --region tri-state `
  --min-reunion-year 1980 `
  --max-reunion-year 1999
```

Use `--region all` to skip state filtering, or `massachusetts` / `west-coast` for later rollout phases.

Expected:

- `People loaded` matches rows that pass region + year filters
- `Validation result: PASS`
- `DRY RUN ONLY`

After a live run, the script prints average RocketReach lookup seconds per record (API time only; add manual verification separately).

## Controlled live test

Start with one person:

```powershell
python .\worker\pilot.py `
  --input "C:\RA_C\private-data\NY_Legacy_Sample_100.xlsx" `
  --output "C:\RA_C\private-data\NY_Pilot_Results_1.xlsx" `
  --limit 1 `
  --live `
  --confirm-credits YES
```

Review that output before changing `--limit` to 5 and then 100.

The result workbook preserves the original columns and adds the RocketReach
candidate, lookup timing, old-vs-new change flags, and an empty human decision
field.
