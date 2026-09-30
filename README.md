# MatchIQ Dashboard - Updates & Maintenance Guide

**Author:** Nisha Sirdesai  
**Last updated:** September 2026  
**Status:** Working prototype with live RocketReach integration, tested on real production data

---

## Quick start - how to run this project

This is **not** a Node/Python app with a local database. The three pieces are:

| Piece | What it is | Where it lives |
|-------|------------|----------------|
| **Frontend** | Live dashboard | `index.html`. Static prototype (localStorage only): `matchGate.html` |
| **Backend** | Google Apps Script Web App | `code.gs` (deployed to Google) |
| **Database** | Google Sheet | tabs `Sheet1` (reviews) + `Records` (people) |

### 1. Prerequisites (one-time)

1. A Google account with access to the MatchIQ Google Sheet and Apps Script project.
2. `ROCKETREACH_API_KEY` set in Apps Script  ->  **Project Settings  ->  Script Properties**.
3. Apps Script deployed as a Web App: **Execute as Me**, **Who has access: Anyone**.
4. The deployment URL copied into `REVIEWS_API_URL` near the top of `index.html`'s `<script>` section (already set in this repo).

### 2. Run the frontend locally

From the project folder:

```powershell
cd c:\RA_C\MatchGate
python -m http.server 8080
```

Then open: [http://localhost:8080](http://localhost:8080)

**Sign-in:** there is no password / SSO login. On first visit (or after Logout), use **Settings** and enter your name (and optional role). That name is stored in the browser and attached to every review you save on the shared Google Sheet.

Alternative (if Python is unavailable):

```powershell
npx --yes serve -l 8080
```

Or simply open `index.html` in a browser (file://). Local HTTP is preferred so `fetch` to Apps Script behaves consistently.

### 3. Confirm backend + DB are connected

In PowerShell:

```powershell
# Records (Google Sheet  ->  Records tab)
Invoke-WebRequest "https://script.google.com/macros/s/AKfycbyI4HSVEi9ohVAz1-EBZ7U2SpJ2Aa5S-CGSp6PKwhaUXkS_X77ZLcUquAJ6uvQ_gSQ/exec?type=records" -UseBasicParsing | Select-Object StatusCode, @{N='Bytes';E={$_.Content.Length}}

# Reviews (Google Sheet  ->  Sheet1 tab)
Invoke-WebRequest "https://script.google.com/macros/s/AKfycbyI4HSVEi9ohVAz1-EBZ7U2SpJ2Aa5S-CGSp6PKwhaUXkS_X77ZLcUquAJ6uvQ_gSQ/exec" -UseBasicParsing | Select-Object StatusCode, @{N='Bytes';E={$_.Content.Length}}
```

Both should return **StatusCode 200**. In the browser DevTools Network tab you should also see successful calls to that same Apps Script URL when the dashboard loads.

**Verified (Sep 2026):** Records API about  404 people; Reviews API returns shared decisions.

### 4. Redeploy backend after editing `code.gs`

1. Paste updated `code.gs` into the Apps Script editor.
2. **Deploy  ->  Manage deployments  ->  edit (pencil)  ->  Version: New version  ->  Deploy**.
3. URL stays the same - no change needed in `index.html`.

---

## Data (Google Sheet - not this repo)

Alumni records and reviews live in the MatchIQ Google Sheet (`Records` and `Sheet1` tabs). Keep raw Advancement exports in Drive/Sheets; do not add Excel dumps to this repo. To change who appears in the dashboard, edit the **Records** tab and refresh.

---

## What changed, and why

The original dashboard (`index.html` / `matchGate.html`) was a fully static prototype:

- All 373 records were hardcoded directly into the HTML as a JavaScript array (`DEFAULT_DATA`)
- Review decisions (Accept/Reject/etc.) were saved to the browser's `localStorage` - meaning they only existed on one person's computer, in one browser, and were never shared with anyone else reviewing the same data
- All matching was manual - no live connection to RocketReach or any other enrichment source

Four fixes were made to turn this into a genuinely usable, shared, live tool:

1. **Shared review storage** - review decisions now sync to a Google Sheet, visible to everyone, with a record of who reviewed what and when, plus a warning if someone is about to overwrite an existing review
2. **Dynamic record loading** - the master list of people to review now loads from a separate Google Sheet tab, so adding a new person just means adding a row to a spreadsheet, no code editing required
3. **Human-in-the-loop filters + export** - added Degree, State, and "graduated before year" filters, plus a CSV export button, so someone like Julie Anne can self-serve spot-checks on specific subsets without needing an engineer
4. **Live RocketReach integration** - a "Refresh via RocketReach" button on each profile calls RocketReach's API in real time to pull current title/company/location/LinkedIn, with the API key kept entirely server-side (never exposed in the HTML or GitHub repo)

---

## Architecture overview

The dashboard in the browser talks to the Google Apps Script web app. That script reads and writes the Google Sheet and calls RocketReach.


**No server hosting required** - the "backend" is entirely a Google Sheet + a free Apps Script Web App deployment. This was a deliberate choice to keep things maintainable by non-engineers (Julie Anne is already comfortable with spreadsheet-like tools).

**Why the API key is safe:** the RocketReach API key is stored as a Script Property inside Apps Script (Project Settings  ->  Script Properties), not in any code file. The browser never sees it - it only talks to the Apps Script Web App, which makes the actual RocketReach call server-side. This means the key is never exposed in `index.html`, never committed to GitHub, and never visible via "View Page Source" on the live GitHub Pages site.

---

## How to add a new person to review

1. Open the Google Sheet, go to the **`Records`** tab
2. Add a new row with the same columns as the existing rows (55+ fields - see a CSV export from the dashboard for the full schema)
3. If you only have basic info (name, degree, employer from the legacy Advancement export) and want RocketReach to fill in the rest, leave the `rr_*` fields blank - the dashboard will show "N/A" for those until someone clicks "Refresh via RocketReach" on that profile
4. Refresh the dashboard - the new person will appear automatically. No code changes needed.

---

## How reviewing works now

1. Click into a profile, choose Accept/Flag/Reject, add a comment if needed
2. You'll be prompted for your name (required - the save won't go through without it)
3. If someone else already reviewed this record, you'll see a warning showing who reviewed it, what they decided, and their comment - you can choose to overwrite or cancel
4. The review is saved to the **`Sheet1`** tab of the Google Sheet, with columns: `id, action, comment, timestamp, name, conf, score, reviewer`

**Known limitation:** if two people review the same record, it's "last write wins" - the second review overwrites the first, though the overwrite warning at least makes this visible rather than silent. There's no automatic conflict resolution beyond that warning.

---

## The filter/export feature

New controls added to the dashboard header:

- **Degree filter** - dropdown, auto-populated from whatever degrees exist in the current `Records` data
- **State filter** - dropdown, auto-populated the same way
- **Region** - rollout presets: Tri-State (NY/NJ/CT), West Coast, Massachusetts
- **Class from / Class to** - filters on `legacy_Reunion Year` / `legacy_Preferred Class Year` (e.g. 1980-1999 for the C-suite scraping cohort)
- **Export CSV** - exports whatever is currently visible/filtered (not the full dataset), as a downloadable CSV file

These combine with each other and with the existing search box and status filters (All / Auto-Accept / Accepted / Review / Rejected).

---

## The RocketReach "Refresh" feature

Each profile now has a **"Refresh via RocketReach"** button that:

1. Asks for confirmation first ("This will use 1 RocketReach lookup credit. Continue?") so credits aren't burned by accident
2. Sends the person's name and best-available employer (RocketReach's own `rr_company` if already known, otherwise falls back to the Advancement system's `legacy_Person Account: Current Employer Name`) to the backend
3. The backend calls RocketReach's `/person/lookup` endpoint
4. If RocketReach needs time to search (common - its API is asynchronous), the backend automatically polls the `/person/checkStatus` endpoint up to 4 times, 2 seconds apart, before giving up and telling the user to try again shortly
5. Displays fresh title, company, location, LinkedIn URL, and **validated email** on the profile
6. **Email validation** (same evaluate screen) rejects Postmark / automated / official inboxes (`noreply@`, `info@`, `official@`, `*.postmarkapp.com`, `*.mtasv.net`, role-based or SMTP-invalid addresses). A usable person email is labeled **Usable**; catch-all/unknown SMTP is **Needs review**; the rest are skipped and listed so reviewers can see what RocketReach returned.
7. Shows a **"Save this update to Records"** button - click it to write `rr_title`, `rr_company`, `rr_location`, `rr_linkedin_url`, and if present `rr_email` / `rr_email_status` back to the Google Sheet **Records** tab (via Apps Script `update_record`). Nothing is auto-saved; the reviewer must opt in after checking the fresh data. Add the two email columns to **Records** if they are not there yet - missing columns are skipped and the rest still save.

RocketReach + evaluate still use this dashboard path (`index.html`  ->  Apps Script `rocketreach_lookup`  ->  Accept/Flag/Reject). The `worker/pilot.py` batch path is not the live review process yet.

**Rate limit awareness:** each confirmed "Refresh via RocketReach" that finds a match consumes one lookup credit (RocketReach does not charge credits for "no match found", per their docs). There is a confirmation dialog, but still no org-wide usage counter in the dashboard - worth discussing before wider rollout.

---

## Real-data testing (as of August 2026)

The dashboard was tested against actual DMSB Advancement data (`All_DMSB_Cut2_Sep8.xlsx`), not just the original 373-record mock dataset:

- The file contains two sheets: **Ascend** (60,773 rows, 60,028 unique people - a newer CRM export with major/college detail) and **Legacy** (67,401 rows, 67,023 unique people - an older Salesforce-style export matching the schema this dashboard's `legacy_` fields are based on)
- **58,443 people appear in both sheets.** Total unique people across both is approximately **68,600** - closely matching the "~65,000 record" figure referenced throughout this project
- Of the Legacy sheet, **33,239 people (about half) have both a job title and employer already on file** - the other half would need RocketReach enrichment (or manual research) just to get basic employment info, let alone verified/current info
- Ascend has **723 people with duplicate rows** - a minor data quality issue worth flagging, likely from multiple gift/activity records per person rather than true duplicates

**A real 30-person sample was pulled** (people with existing job title + employer, sorted by most recent activity) and successfully tested through the full pipeline:

- Loaded dynamically via the Records sheet 
- RocketReach lookup succeeded for at least one person (Rita B. Allen - President, Rita B. Allen Associates) with accurate, current data returned 
- RocketReach lookup correctly failed-gracefully for at least one person (Jennifer Abraham) with a clean "not found" message rather than a crash - this validates Vani's original point that not everyone is RocketReach-findable, and manual review remains necessary for a meaningful portion of the dataset

---

## Backend setup (if this needs to be redeployed or moved)

1. **Google Sheet**: needs two tabs  - 
   - `Sheet1`: columns `id | action | comment | timestamp | name | conf | score | reviewer`
   - `Records`: same column structure as the original dataset (see a dashboard CSV export). Optional for email save: `rr_email`, `rr_email_status`

2. **Apps Script Script Properties**: `ROCKETREACH_API_KEY` must be set (Project Settings  ->  Script Properties) - the script will fail to authenticate with RocketReach without this.

3. **Apps Script code**: see `code.gs` in this repo for the complete script. Handles:
   - reviews (`doPost` / default `doGet`)
   - records (`doGet` with `?type=records`)
   - RocketReach lookups (`doPost` with `type: 'rocketreach_lookup'`), including email validation on the returned addresses
   - saving RR updates to Records (`doPost` with `type: 'update_record'`)

4. **Deployment**: Deploy  ->  New deployment  ->  Web app  ->  Execute as "Me"  ->  Who has access "Anyone". Copy the resulting URL into the `REVIEWS_API_URL` constant near the top of `index.html`'s `<script>` section.

5. **Redeploying after script changes**: Deploy  ->  Manage deployments  ->  edit (pencil icon)  ->  Version: New version  ->  Deploy. This keeps the same URL, so `index.html` doesn't need to be touched again.

---

## What's NOT built yet / open questions for the team

- **Org-wide RocketReach usage meter** - there is a per-click confirmation dialog, but no shared credit counter or soft quota in the UI.
- **Auto-save vs opt-in save** - opt-in **"Save this update to Records"** is implemented. Still open: should saves overwrite `rr_*` only (current behavior), or also track a separate "verified" / audit column?
- **Full 65K-68K rollout** - everything tested so far uses a small real sample plus the original ~373 mock set (live Sheet currently ~404 records). Running the full dataset or the ~5,764 matched file through RocketReach at scale needs budget/credit planning first.
- **Data quality cleanup** - the 723 duplicate rows in Ascend, and the ~50% of Legacy records missing job title/employer entirely, are real gaps that exist independent of anything this dashboard does.
- **Automated quality scoring improvements** - per Vani's team, similarity/confidence metrics alone aren't reliable enough to replace human review; this remains a human-in-the-loop process by design, not something meant to be fully automated.
