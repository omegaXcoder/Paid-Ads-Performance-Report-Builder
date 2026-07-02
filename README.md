# Ads Performance Report Builder

Joins `Ads_Raw_Metrics` + `WhatConverts_Raw_Leads` (via `Account_Mapping`,
which now includes each client's Target Cost / Qualified Lead) and writes
the result into your report template, producing one dated tab per day
inside a monthly file:

```
Ads Performance Reports - July 2026   (Drive file, one per month)
├── Dashboard                    ← pristine template, never touched by the script
├── Last Year Performance        ← pristine template, never touched by the script
├── Account Mapping (Source)     ← seeded once, when the month file is created
├── July 1, 2026                 ← daily tab (duplicated from Dashboard)
├── July 1, 2026 Previous Year   ← daily tab (duplicated from Last Year Performance)
├── July 2, 2026
├── July 2, 2026 Previous Year
└── ...
```

**Every day's numbers are frozen as plain values at write time** —
including the Target-CPL comparison that drives the color coding. A
later change to a client's Target CPL will never retroactively repaint
a day you've already generated.

**Always defaults to a dry run.** Locally, it creates and writes nothing
unless you pass `--live`. It **always emails a log** (dry run or live) to
`omega@kudos.marketing`, subject line prefixed `AUTOMATION LOGGING:` for
your email filter. On GitHub Actions, scheduled runs are always live;
manual runs from the Actions tab default to dry run unless you tick the
"live" box.

## 1. One-time setup: convert your template to a real Google Sheet

This script duplicates tabs via the Sheets API, which only works on native
Google Sheets — not `.xlsx` files sitting in Drive.

1. Upload your `.xlsx` template to Google Drive (the same folder as your
   raw data spreadsheet, or wherever you'd like monthly report files to live)
2. Right-click it → **Open with → Google Sheets** — this converts it in place
3. Open the new Google Sheet, copy its file ID from the URL
   (`https://docs.google.com/spreadsheets/d/THIS_PART/edit`)
4. That's your `TEMPLATE_SPREADSHEET_ID`

**Recommended:** rename the `Dashboard` and `Last Year Performance` tabs in
this template file to something like `TEMPLATE - Dashboard` so they're
visually distinct from real dated tabs if you ever open the template
directly — the script looks them up by exact name (`Dashboard` /
`Last Year Performance`), so if you rename them, update `TEMPLATE_CURRENT_TAB`
/ `TEMPLATE_PRIOR_YEAR_TAB` at the top of `main.py` to match.

## 2. Local setup

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env`:
- `RAW_SPREADSHEET_ID` — already pre-filled
- `TEMPLATE_SPREADSHEET_ID` — from step 1 above
- `REPORTS_FOLDER_ID` — leave blank to auto-use the same folder as your raw
  data spreadsheet, or set explicitly if you want monthly files somewhere else
- `GOOGLE_SERVICE_ACCOUNT_FILE` — reuse the same service account JSON from
  your other two automations, as long as it has edit access to both the raw
  data spreadsheet AND the template spreadsheet (share the template with
  the service account's email, same as you did for the raw sheet)
- `SMTP_*` / `EMAIL_FROM` / `EMAIL_TO` — for the run log email

## 3. GitHub Actions setup

**Settings → Secrets and variables → Actions**, add:

| Secret | Value |
|---|---|
| `RAW_SPREADSHEET_ID` | `1v9pqP0IQPsHLF45pTqlHvzkTtg1fHTYNVuNvSzbOU54` |
| `TEMPLATE_SPREADSHEET_ID` | Your template's Google Sheet ID |
| `REPORTS_FOLDER_ID` | Leave unset to auto-detect, or set explicitly |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Full contents of your service account JSON key |
| `SMTP_HOST` / `SMTP_PORT` / `SMTP_USERNAME` / `SMTP_PASSWORD` | Your SMTP creds |
| `EMAIL_FROM` | `automation@kudos.marketing` |
| `EMAIL_TO` | `omega@kudos.marketing` |

Scheduled to run daily at **11:45 UTC** — after the Ads export (~11:00) and
WhatConverts export (~11:20) should both have finished, giving this job
fresh raw data to read.

## Commands / modes

| What you want | How |
|---|---|
| **Dry run locally** | `python main.py` |
| **Live run locally** | `python main.py --live` |
| **Manual dry run on GitHub** | Actions tab → "Ads Performance Report Builder" → Run workflow → leave "live" unchecked |
| **Manual live run on GitHub** | Actions tab → Run workflow → check "live" |
| **Change the schedule** | Edit the `cron` line in `.github/workflows/report_builder.yml` |

## What gets written

- **New month, first run:** creates `Ads Performance Reports - [Month Year]`
  in the reports folder (copied from the template), seeds
  `Account Mapping (Source)` from the live `Account_Mapping` tab
- **Every run:** duplicates `Dashboard` → renames to today's date;
  duplicates `Last Year Performance` → renames to today's date + " Previous Year";
  writes the 3 period tables (Last 7 Days, Previous Full Week, Month to Date)
  into each, with the Target-CPL columns frozen as values
- **Footnote cell (A47) is cleared** on every generated tab — no boilerplate text

## Known limitations

- **Ads vs. WhatConverts date alignment.** The Ads export runs in Google
  Ads Scripts (account timezone) and the WhatConverts export runs on
  GitHub Actions (UTC by default) — for "Last 7 Days" and "Month to Date,"
  their exact day boundaries can be offset by about a day depending on time
  of day. Rows are joined by *period type* (Last 7 Days / Previous Full
  Week / Month to Date + current-vs-prior-year), not by exact date range,
  so this doesn't break the join — but the displayed date range in each
  tab's header comes from whichever source had data for that client
  (preferring Ads), and won't always perfectly describe both sources'
  exact windows. "Previous Full Week (Mon-Sun)" is unaffected since it's
  always a clean calendar week.
- **New client onboarding.** A client only appears in the report once
  they exist in `Account_Mapping` with both a WhatConverts Profile ID and
  Google Ads Customer ID filled in.
- **Target CPL missing:** a client with no Target CPL set shows the
  "gray — no target set" color state, same as the template's original
  design intent.
- **10-client row range is fixed** by the template's layout (rows 7–16,
  21–30, 35–44 per table). Adding an 11th client requires expanding the
  template's row ranges and updating `PERIOD_ROWS` in `main.py` to match.
