# WebPT Scheduler History Automation — Phase 1 (Scrape → Sheet)

This is the GitHub Actions half of the system: given a date range, it
logs into WebPT, pulls scheduler-history for every clinic, keeps
**every** event type (Created, Cancelled, Checked In, Checked Out,
Deleted, Edited, No Show), looks up each `CREATOR USER` in the `Names`
tab to tag a `GROUP` / `GROUP 2` (department), and appends only the
genuinely new rows into the `Master Data` tab of your Google Sheet.
Event-type and department **filtering happens later**, in the Reports
step (Phase 2, the Apps Script web app) — this phase's job is just to
keep Master Data complete and de-duplicated.

## 1. Repo secrets to set (Settings → Secrets and variables → Actions)

| Secret | What it is |
|---|---|
| `WEBPT_USERNAME` | Your WebPT login username |
| `WEBPT_PASSWORD` | Your WebPT login password |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The full contents (paste as-is) of a Google Cloud service account key JSON file, with the Sheets API enabled |
| `SHEET_ID` | The ID from your Sheet's URL: `.../spreadsheets/d/`**`THIS_PART`**`/edit` |

## 2. Google Sheet setup

Share the Sheet with the service account's email (found in the JSON
key as `client_email`) as **Editor**.

Required tabs (you already have these per our conversation):

- **`Names`** — must have a `Scheduler Name` column (matched against
  the scraped `CREATOR USER` field) and `Group` / `Group 2` columns.
- **`Groups`** — reference list of Group / Group 2 pairs (used later
  by the Reports dropdowns, not read by this scrape step).
- **`Emails`** — column A = recipient addresses (used later by the
  Reports step).

This workflow creates these tabs automatically on first run if
they're missing:

- **`Master Data`** — the accumulating table of every scraped row,
  plus `GROUP` and `GROUP 2`.
- **`Runs`** — one row per scrape run: timestamp, date range, rows
  scraped, rows newly added, running total.

## 3. Running it manually (before the web app exists)

Go to **Actions → WebPT Scheduler History Scrape → Run workflow**,
enter a start and end date as `MM/DD/YYYY`, and run it. Once it
finishes, check the `Master Data` and `Runs` tabs.

## 4. What's next (Phase 2)

The Apps Script web app will call this same workflow via GitHub's
`workflow_dispatch` REST API (using a GitHub Personal Access Token
stored in Apps Script's Script Properties, never in the Sheet itself)
instead of you clicking "Run workflow" by hand — plus the Reports
section (event-type + department filters, pivot table, email).
