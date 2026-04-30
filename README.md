
# Jira Export Helper

This tool calls a Jira Search API endpoint, expands changelog history, and exports:

1. `issue_summary_YYYYMMDD_HHMMSS.csv`  
   One row per issue with current snapshot fields, SLA/resolution metrics, and workflow-agnostic status transition columns.

2. `issue_activity_YYYYMMDD_HHMMSS.csv`  
   One row per changelog item with:
   - Issue Key
   - Change Date
   - Author
   - Field
   - From
   - To

3. `jira_exports_YYYYMMDD_HHMMSS.zip`  
   Bundle containing both CSVs and run metadata.

## Features

- Multiple projects
- Multiple issue types
- Current status filter
- Assignee filter
- Label filter
- Date-field selection (`created`, `updated`, `resolutiondate`)
- Date/time range
- Time block filter on changelog events
- Optional custom JQL override
- Pagination
- Optional exclusion of admin workflow events
- Wide status transition columns (`Status 1..30 From/To/Timestamp/Author`)
- `Status Transition Count`, `Status Overflow Count`, and `Status Path`
- Simple HTML UI

## Setup

Create a virtual environment and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
pip install flask requests
```

Set Jira credentials as environment variables.

macOS / Linux:
```bash
export JIRA_USERNAME="your_username"
export JIRA_PASSWORD="your_password"
```

or
```bash
export JIRA_EMAIL="you@example.com"
export JIRA_API_TOKEN="your_token"
```

Windows (PowerShell) - current session only:
```powershell
$env:JIRA_USERNAME="your_username"
$env:JIRA_PASSWORD="your_password"
```

or
```powershell
$env:JIRA_EMAIL="you@example.com"
$env:JIRA_API_TOKEN="your_token"
```

Windows (PowerShell) - persist for future sessions:
```powershell
[Environment]::SetEnvironmentVariable("JIRA_USERNAME", "your_username", "User")
[Environment]::SetEnvironmentVariable("JIRA_PASSWORD", "your_password", "User")
```

or
```powershell
[Environment]::SetEnvironmentVariable("JIRA_EMAIL", "you@example.com", "User")
[Environment]::SetEnvironmentVariable("JIRA_API_TOKEN", "your_token", "User")
```

Windows (Command Prompt / cmd.exe) - current session only:
```cmd
set JIRA_USERNAME=your_username
set JIRA_PASSWORD=your_password
```

or
```cmd
set JIRA_EMAIL=you@example.com
set JIRA_API_TOKEN=your_token
```

Windows (Command Prompt / cmd.exe) - persist for future sessions:
```cmd
setx JIRA_USERNAME "your_username"
setx JIRA_PASSWORD "your_password"
```

or
```cmd
setx JIRA_EMAIL "you@example.com"
setx JIRA_API_TOKEN "your_token"
```

After setting persistent variables, restart PowerShell/Cursor terminal.

## Run

```bash
python app.py
```

Then open:

```text
http://127.0.0.1:5001
```

### Windows quick start (recommended)

In PowerShell, set credentials and start the app from the same terminal session:

```powershell
# from the project folder
& ".\.venv\Scripts\Activate.ps1"

# recommended: API token auth
$env:JIRA_EMAIL="you@example.com"
$env:JIRA_API_TOKEN="your_token"

# start the app using the venv python (avoids Windows Store python alias issues)
.\.venv\Scripts\python.exe .\app.py
```

If you see “Login Required” / “anonymous user” errors after setting env vars, restart the app so it picks up the variables.

## UI Fields (What to Enter)

Use the form on the home page to define your export request.

Required / recommended:
- `Jira Search Endpoint` (`base_url`): Jira search API URL.  
  Example: `https://your-jira-domain/rest/api/2/search`
- `Projects` (`projects`): Comma-separated project keys.  
  Example: `CSSD,ABC`

Common filters (optional):
- `Issue Types` (`issue_types`): `Bug,Task,Story`
- `Current Statuses` (`statuses`): `Open,In Progress,Closed`
- `Assignees` (`assignees`): Jira usernames/account identifiers, comma separated
- `Labels` (`labels`): Comma-separated labels

Date and time controls:
- `Date Field` (`date_field`): One of `created`, `updated`, `resolutiondate`
- `Start Date/Time` (`start_dt`): Lower bound for selected date field
- `End Date/Time` (`end_dt`): Upper bound for selected date field
- `Time Block Start` (`time_block_start`): Optional time-of-day start (HH:MM)
- `Time Block End` (`time_block_end`): Optional time-of-day end (HH:MM)

Query and limits:
- `Page Size` (`page_size`): Jira page size per request (default `50`, max `100`)
- `Max Issues` (`max_issues`): `0` means no cap; otherwise stop at this many issues
- `Extra JQL` (`extra_jql`): Extra clause appended with `AND`
- `Custom JQL Override` (`custom_jql`): Full JQL; when set, it replaces builder inputs
  - Example: `issuekey = CSSD-123` (do not prefix with `custom_jql =`)

Checkbox options:
- `Include comments in summary metrics` (`include_comments`)
- `Include workflow admin changes in activity export` (`include_workflow_events`)
- `Verify SSL` (`verify_ssl`)

Important:
- Jira credentials are **not entered in the UI**; they come from environment variables (`JIRA_USERNAME`/`JIRA_PASSWORD` or `JIRA_EMAIL`/`JIRA_API_TOKEN`).

## Tabs

The app now has two report tabs:

- `Legacy Reports`: existing export workflow (`/preview-jql`, `/run-export`) preserved for backward compatibility.
- `CSMS Executive Incident Summary`: new executive dashboard comparing two rolling periods.

### CSMS Executive Incident Summary Inputs

- `Jira Search Endpoint` (`base_url`)
- `Projects` (`projects`) default `CSSD,CSD,CDF`
- `Report Generation Date/Time` (`report_datetime`)
- `Period Length (days)` (`period_length`, default `15`)
- `Issue Types` (`issue_types`)
- `Status Filters` (`statuses`)
- `Components` (`components`)
- `Page Size` (`page_size`)
- `Max Issues` (`max_issues`)
- `Process Alignment %` (`process_alignment_pct`, default `60`)

### CSMS Period Windows

For report date `R` and period length `N`:

- `Period 2` = `R-(N-1)` through `R`
- `Period 1` = the preceding `N` days ending one day before `Period 2`

Example for `R=2026-04-28`, `N=15`:
- `Period 2`: `2026-04-14` to `2026-04-28`
- `Period 1`: `2026-03-30` to `2026-04-13`

### CSMS Business Rules

- CSSD final status: `Closed`
- CSD final status: `Ready For Production Users`
- Backlog counts issues not yet in project final status.

### CSMS Export Artifacts

After `Refresh from Jira API`, the dashboard exposes:

- CSV export (ZIP): `raw_period1`, `raw_period2`, and KPI table CSVs
- Excel export: one workbook with separate sheets for period raw rows and KPI metrics
- PDF export: executive summary snapshot (KPIs + narratives)

### CSMS Troubleshooting

- Jira 400/500 API errors are returned with server response details in JSON.
- If Jira returns a referral ID in an error page, include it when contacting Jira administrators.
- For large result sets, increase `page_size` and keep `max_issues=0` to allow full pagination.

## Summary Output (Current Schema)

`issue_summary_*.csv` includes:
- Core issue snapshot fields (`Issue Key`, `Summary`, `Issue Type`, `Priority`, `Assignee`, etc.)
- SLA/resolution fields:
  - `First Response Date`
  - `Resolution`
  - `Resolution Date`
  - `Time to First Response`
  - `First Response SLA Breached`
  - `Time to Resolution`
  - `Resolution SLA Breached`
- Workflow-agnostic status movement fields:
  - `Status Transition Count`
  - `Status Overflow Count` (`max(0, transition_count - 30)`)
  - `Status Path` (single serialized path)
  - `Status 1..30 From/To/Timestamp/Author`

`issue_activity_*.csv` remains the detailed source of truth with:
- `Issue Key`, `Change Date`, `Author`, `Field`, `From`, `To`

## Batch Reports Process

### UI batch process (manual, repeatable)

Use this when you want controlled project/date-window runs from the app UI.

1. Start app and confirm credentials are available in the same terminal session.
2. For each batch window, set:
   - `Projects` (example: `CSSD`)
   - `Date Field` (usually `created` or `updated`)
   - `Start Date/Time`
   - `End Date/Time`
   - Optional filters (`Issue Types`, `Statuses`, `Labels`, `Assignees`)
3. Run export and download generated files.
4. Repeat for the next window/project.

Suggested project/date-window batch set (example):
- Run A: Project `CSSD`, `created`, `2026-04-01 00:00` to `2026-04-07 23:59`
- Run B: Project `CSSD`, `created`, `2026-04-08 00:00` to `2026-04-14 23:59`
- Run C: Project `CSSD`, `created`, `2026-04-15 00:00` to `2026-04-21 23:59`
- Run D: Project `ABC`, `created`, `2026-04-01 00:00` to `2026-04-21 23:59`

Because filenames are timestamped, each run produces unique artifacts.

### Scripted batch process (PowerShell)

Use this for repeatable local automation against the running app.

```powershell
$batch = @(
  @{ projects = "CSSD"; date_field = "created"; start_dt = "2026-04-01T00:00"; end_dt = "2026-04-07T23:59" },
  @{ projects = "CSSD"; date_field = "created"; start_dt = "2026-04-08T00:00"; end_dt = "2026-04-14T23:59" },
  @{ projects = "ABC";  date_field = "updated"; start_dt = "2026-04-01T00:00"; end_dt = "2026-04-21T23:59" }
)

foreach ($item in $batch) {
  $payload = @{
    base_url = "https://jira.mdthink.maryland.gov/rest/api/2/search"
    projects = $item.projects
    date_field = $item.date_field
    start_dt = $item.start_dt
    end_dt = $item.end_dt
    page_size = 50
    max_issues = 0
    include_comments = $true
    include_workflow_events = $false
    verify_ssl = $true
    extra_jql = ""
    custom_jql = ""
  } | ConvertTo-Json

  $resp = Invoke-RestMethod -Uri "http://127.0.0.1:5001/run-export" -Method Post -ContentType "application/json" -Body $payload
  $resp | Select-Object jql, issue_count, summary_rows, activity_rows, files
}
```

### Batch troubleshooting

- If responses look anonymous or empty unexpectedly, restart app after setting `JIRA_*` env vars.
- `custom_jql` overrides all builder filters (`projects`, dates, statuses, etc.).
- Timestamped filenames are expected and help separate batch artifacts.

## Endpoint example

Default base URL:
```text
https://jira.mdthink.maryland.gov/rest/api/2/search
```

Example generated JQL:
```text
project in ("CSSD", "CSD", "CDF") AND issuetype in ("Bug") AND created >= "2026-04-01 00:00" AND created <= "2026-04-21 23:59"
```

## Notes

- First response/resolution SLA columns can be blank when Jira SLA custom fields have no completed cycles for that issue/request type.
- Workflow admin change records can be excluded from the activity export.
- If your Jira server uses self-signed certificates, uncheck **Verify SSL** in the UI or set up trusted certs.

## Direct API pattern used

The app calls the search endpoint with parameters similar to:

```text
/rest/api/2/search?jql=project in (CSSD, CSD, CDF)&expand=changelog&maxResults=50
```

and paginates until all results are collected or `Max Issues` is reached.
