# CSMS Operations Dashboard (Jira Export Helper)

Single-file Flask app (`app.py`) that serves the **CSMS Operations** web UI and Jira-backed APIs: executive CSMS summary, team posture, legacy CSV exports, trends dashboard, auth diagnostics, and downloads (CSV, Excel, PDF, ZIP).

Default dev URL: `http://127.0.0.1:5001`

---

## Navigation (sidebar)

| Tab | Purpose |
|-----|---------|
| **Executive Report** | CSMS Application — KPIs, narratives, health panel, stuck ticket drill-down, Chart.js (daily trend, top category, status). |
| **Operations Team** | Per-member ticket posture — metric cards, status summary, oldest open detail, label pie chart, CSV preview, exports. |
| **Ticket trend** | Legacy dashboard — created/updated/resolved trends and current status distribution charts; shares **Report Settings** with Data Exports. |
| **Notes** | In-app explainer for CSMS KPIs, Team Posture definitions, usage, and auth/data quality tips. |
| **U** (Auth) | Jira API auth diagnostics (`/myself`, visible projects, CSSD/CSD access flags). |
| **Theme** | Light / dark / system theme toggle (persisted in the browser). |

**Variables & Settings** are collapsible cards per scope: Legacy export form, CSMS parameters, and Team Posture parameters (shown/hidden based on the active tab).

---

## HTTP API (summary)

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `/` | Main HTML UI (`render_template_string`). |
| `POST` | `/preview-jql` | Returns built JQL for the legacy export form. |
| `POST` | `/run-export` | Runs full legacy export; returns paths + download links. |
| `POST` | `/run-csms-exec-summary` | CSMS executive payload + optional CSV ZIP / Excel / PDF exports. |
| `POST` | `/run-legacy-dashboard` | Ticket trend: charts + KPIs including TTFR/TTR CSSD/CSD (configurable median, mean, or p90). |
| `POST` | `/run-team-posture` | Team posture JSON for one member; includes `jql`, `broad_jql`, `metrics`, `warnings`, exports. |
| `POST` | `/run-team-posture-board-export` | Team board export for all submitted members; includes dashboard-bucket tagging rows for matched tickets. |
| `POST` | `/auth-status` | Auth / visibility diagnostics JSON. |
| `GET` | `/download?path=...` | Legacy export file download (server-side path). |
| `GET` | `/download-csms-export?export_id=&kind=` | Cached CSMS export (`csv_zip`, `excel`, `pdf`). |
| `GET` | `/download-team-posture-export?export_id=&kind=` | Cached team export (`csv`, `excel`). |
| `POST` | `/run-pipeline-backlog-count` | Fast **Pipeline Backlog** count (Jira search `total`, `maxResults=0`). |
| `POST` | `/run-team-board-metrics` | **Team Closed** board metric (pass `skip_pipeline: true` for closed-only; omit for both). |
| `POST` | `/run-team-posture-refresh` | **Refresh all roster members:** one broad Jira fetch, in-memory metrics per member, board metrics; returns `pool_cache_id` for export. |
| `POST` | `/run-team-posture-board-export` | **Slim team CSV** (`Member Name`, `Dashboard Bucket`, `Issue Key`, `Summary`); pass `pool_cache_id` after refresh to avoid re-fetching Jira. |
| `GET` | `/snapshots/list-options?report_id=` | Dropdown options for official saved reports (`exec`, `ops`, `legacy`). |
| `GET` | `/snapshots/<id>` | Full snapshot record including saved `params`. |
| `DELETE` | `/snapshots/<id>?report_id=` | Remove one saved snapshot (`report_id` optional guard). |
| `GET` | `/snapshots/<id>/display` | Hydrate dashboard from a saved snapshot (no Jira); includes `params` for form restore. |
| `POST` | `/snapshots` | Manually save an official report snapshot. |
| `GET` | `/snapshots/compare?report_id=&snapshot_id=` | Between-report metric deltas vs previous snapshot or manual baseline. |
| `POST` | `/snapshots/compare-live` | Compare current live metrics to last snapshot / manual baseline. |
| `GET` | `/snapshots/trends?report_id=&metric_key=` | Time series for a metric (optional API). |
| `GET` | `/snapshots/label-trends?report_id=&member_username=` | Label trends: `ops` requires `member_username`; `legacy` uses whole-report `charts.label_distribution` from saves. |
| `GET` | `/snapshots/audit?report_id=` | Data-quality audit of saved snapshots (issues + warnings per row). |
| `POST` | `/snapshots/migrate` | Backfill `trend.*` from `view` data (`{"report_id":"legacy"}` or `{"snapshot_id":5}`). |
| `POST` / `GET` | `/manual-baselines` | Manual comparison fallback values when no prior snapshot exists. |

---

## Official report snapshots (`data/snapshots.db`)

SQLite database (stdlib only) stores **manually saved** dashboard runs. Refreshing from Jira does **not** auto-save.

- **Archive mode:** On tab open, the latest official snapshot loads so KPIs/cards work without Jira.
- **Report dropdown:** Select any past save or switch to **Live**. Live Jira data loads only when you use **Refresh from Jira** (in Team Posture settings), not when selecting Live alone.
- **Selecting a saved report** loads archived **metrics only** (KPIs/cards/charts). It does not change the settings form until you click **Load saved settings**.
- **Load saved settings / Rerun with saved settings:** Pick a saved report (not Live), then **Load saved settings** copies `params` into the report form (Operations Team also restores `team_members` when the snapshot was saved with a roster). **Rerun** loads settings, switches to Live, and runs the tab’s refresh (CSMS submit, Ticket trend refresh, or Refresh All Member Metrics).
- **Delete snapshot:** Pick a saved report (not Live), click **Delete snapshot**, confirm. Removes that row from SQLite and switches to Live. Does not remove manual baselines (`manual_baselines` table).
- **Download Team CSV:** Uses in-memory cache when every member has ticket rows; otherwise calls `POST /run-team-posture-board-export` against Jira. Archived snapshots do not store ticket rows — rerun live first.
- **Operations Team cards:** Each metric card shows a **delta %** vs the prior official report (or manual baseline). Sparklines on cards are disabled.
- **Pipeline Backlog:** Uses the **CSMS Prod pipeline JQL** (default: project CSMS Defect Management, Phase Reported = Prod, status in New / In Progress / Reopened, issue-type and label exclusions, `created >=` configurable date default 2021-11-08). Override via **Pipeline Backlog JQL** in Team settings. Count = Jira search **`total`** for that JQL (no full issue download). Loaded via `/run-pipeline-backlog-count` first during refresh so it is not blocked by Team Closed.
- **Team Closed:** CSSD/CSD issues in **Closed** status (optional `updated` window from Team Start/End), capped fetch with changelog for roster attribution.

Backup `jira_export_app/data/snapshots.db` with the app folder. The file is gitignored by default.

**Audit / repair:** `python audit_snapshots.py` lists gaps (missing label_distribution, trend fields, params). `python audit_snapshots.py --migrate` backfills `trend` from `view` without re-running Jira. New saves run the same normalization automatically. API: `GET /snapshots/audit`, `POST /snapshots/migrate`.

---

## Ticket trend (Legacy dashboard tab)

Open **Ticket trend** in the sidebar, then **Show Report Variables & Settings** → **Refresh Dashboard**.

### KPI cards

| Card | Source | Default status gate | Default aggregate |
|------|--------|---------------------|-------------------|
| **TTFR CSSD** | `customfield_10318` elapsedTime (+ optional stop−created fallback) | *(blank = any ticket with TTFR SLA)* | Median |
| **TTFR CSD** | Linked **CSSD** `customfield_10318` if linked; else CSD SLA (+ optional fallback) | *(blank)* | Median |
| **TTR CSSD** | `customfield_10317` elapsedTime (+ optional `resolutiondate − created`) | `Closed` | Median |
| **TTR CSD** | Same as TTR CSSD on CSD tickets | `Ready For Production Users` | Median |

Card subtitle shows the chosen aggregate and ticket count (e.g. `mean · 42 ticket(s)`). Saved snapshots store the numeric rollup under `ttfr_*_median_hours` / `ttr_*_median_hours` regardless of aggregate name.

### Report settings (SLA section)

**Status gates** (comma-separated Jira status names; filter which tickets enter each rollup):

- `ttr_status_cssd`, `ttr_status_csd`, `ttfr_status_cssd`, `ttfr_status_csd`

**Aggregates** (how to combine per-ticket hours into one card value):

- `ttfr_cssd_aggregate`, `ttfr_csd_aggregate`, `ttr_cssd_aggregate`, `ttr_csd_aggregate` — `median`, `mean`, or `p90`

**JQL scope:** Same as the export form (projects, dates, filters). Include **CSSD** and **CSD** in projects for SLA cards to populate.

### Ticket Labels

**Created / Updated / Resolved Trends** is a **line chart** by day. **Ticket Labels** (below status charts): horizontal bar for top **15** labels; **label trend line chart** for top **10** labels across saved official reports. **TTFR / TTR** cards show **▲/▼ % vs prior saved report** (green = faster / lower hours, red = slower / higher hours). Save snapshots after refresh so comparisons and label trends have a baseline.

### Snapshots

Official report id `legacy`. Trend keys include the four SLA hour metrics for compare-over-time when you save snapshots. Label distribution is stored under `charts.label_distribution` in each save.

---

## Legacy data export (Trends tab + Data Exports)

End-to-end Jira Search with changelog expansion, producing:

1. **`issue_summary_YYYYMMDD_HHMMSS.csv`** — One row per issue: snapshot fields, SLA/resolution columns, status transition slots (`Status 1..30` From/To/Timestamp/Author), transition count, overflow, status path.  
2. **`issue_activity_YYYYMMDD_HHMMSS.csv`** — One row per changelog item (issue key, date, author, field, from, to).  
3. **`jira_exports_YYYYMMDD_HHMMSS.zip`** — Bundle + run metadata.

**Features:** multiple projects/issue types/statuses/assignees/labels; `date_field` (`created` / `updated` / `resolutiondate`); date/time range; optional time-of-day filter on changelog events; `extra_jql` / `custom_jql`; pagination; optional workflow admin events; optional comments in summary metrics; SSL verify toggle.

**Endpoints:** `POST /preview-jql`, `POST /run-export`.

---

## CSMS Executive Incident Summary (Executive tab)

- **Rolling periods:** configurable report datetime, period length (days), KPI comparison (backlog, new created, resolved), trend %, narratives, process alignment health, stuck ticket highlight.  
- **Elapsed time:** optional **Last Report Timestamp** drives the “time since last report” sentence.  
- **Charts:** daily created/updated/resolved, top category, status mix.  
- **Business rules:** CSSD final status `Closed`; CSD final status `Ready For Production Users` (used for backlog/open vs final).  
- **Exports:** CSV ZIP (raw period 1 & 2 + KPI CSV), multi-sheet Excel, PDF summary — via `POST /run-csms-exec-summary` then `/download-csms-export`.

**Form fields (typical):** `base_url`, `projects`, `report_datetime`, `last_report_timestamp`, `last_report_backlog_tickets`, `last_report_new_created`, `last_report_resolved_tickets`, `period_length`, `issue_types`, `statuses`, `components`, `page_size`, `max_issues`, `process_alignment_pct`, `verify_ssl`.

---

## Team Member Ticket Posture (Team tab)

- **Live refresh:** **Refresh from Jira** (all roster members + board metrics) and **Refresh selected member** live in **Team Posture Variables & Settings** (collapsible card). Status hint (`#teamDataModeHint`) shows Live vs archive and cached member count.
- **Refresh behavior:** **Refresh from Jira** calls `POST /run-team-posture-refresh` — **one** broad JQL fetch for the whole team, then per-member metrics in memory (no ticket rows). Pipeline + Team Closed load with the same refresh. Use **Save snapshot** to persist metrics to SQLite.
- **Export:** **Download Team CSV** / per-member **Download CSV** build slim rows on demand from the server issue pool (`pool_cache_id` from the last refresh when still in server memory).
- **Mobile:** Hamburger sidebar and responsive metric/chart grids at narrow widths.
- **Team roster:** add/remove members (display name + **Assignee username**); roster stored in **browser localStorage**.  
- **Jira queries:**  
  - **`jql`** — same filters as broad query **plus** an `assignee in (...)` clause for the selected member’s Jira username.  
  - **`broad_jql`** — project, issue type, **`created`** date range only (no assignee). Used for worked-on, reopened-in-scope, contributed resolved, and SLA scope over member-linked issues.  
- **Ownership:** assignee match **or** (CSD project + configured **CSD Assigned Developer** field, default `customfield_14700`) when that field identifies the developer.  
- **Team header rollups** (above member pills): **Team Queue Backlog**, **Team In Progress**, **Team Resolved (Report Period)**, **Team Closed** — summed or deduped from cached live refresh (see board metrics API). Delta % vs prior official snapshot when compare data exists.
- **Metric cards:**  
  - **Queue Backlog** — CSSD: Under QA Analysis; CSD: New (owned tickets only).  
  - **In Progress** — CSSD: open, not New, not Under QA Analysis; CSD: open, not New.  
  - **Worked Status (Last 8 Hours)** — owned tickets where you authored a **status** changelog entry in the last 8 hours.  
  - **Worked Status (Others, Last 8 Hours)** — tickets owned by someone else where you authored a **status** changelog entry in the last 8 hours.  
  - **Resolved (Owned)** — owned tickets whose status matches resolved-style rollups (e.g. resolved, closed, ready for production users, completed, duplicate, dev-completed).  
  - **Resolved (Last 8 Hours)** — owned tickets resolved within the last 8 hours.  
  - **Resolved (Contributed)** — same status rollups, member authored ≥1 **status** changelog transition, **not** current owner.  
  - **Assigned Open** — owned CSSD/CSD only: CSSD not Resolved/Closed; CSD not Ready For Production Users.  
  - **Reopened** — current status name contains `reopened` / `re-opened` / `re opened`, and member is owner **or** (not owner but has status transitions as author).  
  - **Worked On (Assigned to Others)** — status-change author is member; current owner ≠ member (any time in changelog).  
  - **SLA Breach Count** — **24h from `created`:** open tickets past 24h; closed tickets prefer Jira **Resolution SLA Breached**-style custom field (discovered via `/rest/api/2/field`), else fallback elapsed created → resolutiondate (or updated). Counts include relevant open + closed breaches in member scope.  
  - **Open &lt; 8h to SLA breach** — open tickets with under 8 hours remaining before the 24h window.  
  - **Oldest open** — key + age (days) + detail JSON.  
- **Other UI:** ticket count by status (owned issues), **Ticket Labels** bar chart (current) + **Trend** chart from saved snapshots (`/snapshots/label-trends`), CSV preview (first rows of raw export).  
- **Exports:** per-member **CSV** + **Excel** (raw tickets + summary metrics including queue/in-progress/worked-status counts). **Download Team CSV** — session cache when all members have `raw_rows`; otherwise `POST /run-team-posture-board-export` (dashboard bucket tags, one row per issue per matching bucket). Partial cache download on export failure when some members are cached.  
- **Warnings:** if Jira returns 500 with `changelog` expansion, the app may retry without changelog (reopen / worked-on / contributed paths may be incomplete).

**Form fields:** `base_url`, `projects`, `start_dt`, `end_dt`, `issue_types`, `csd_assigned_dev_field`, `page_size`, `max_issues`, `pipeline_backlog_created_since`, `pipeline_backlog_jql` (optional override), `verify_ssl`, plus per-request `assignee_username` and `member_name`. Board metrics requests may include `member_usernames`, `skip_pipeline`, or `skip_closed`.

---

## Jira authentication

Credentials are **not** entered in the UI. Set environment variables:

- `JIRA_USERNAME` + `JIRA_PASSWORD`, **or**  
- `JIRA_EMAIL` + `JIRA_API_TOKEN`  

Use **Auth** tab (`POST /auth-status`) to confirm `myself`, visible projects, and CSSD/CSD access.

---

## Setup

```bash
python -m venv .venv
# Windows: .\.venv\Scripts\Activate.ps1
# macOS/Linux: source .venv/bin/activate
pip install flask requests openpyxl reportlab
```

---

## Run

```bash
python app.py
```

Default port **5001** (see `app.run` at bottom of `app.py`).

### Windows (recommended)

Use the venv interpreter to avoid the Windows Store `python.exe` shim:

```powershell
cd path\to\jira_export_app
.\.venv\Scripts\Activate.ps1
$env:JIRA_EMAIL="you@example.com"
$env:JIRA_API_TOKEN="your_token"
.\.venv\Scripts\python.exe .\app.py
```

---

## Legacy UI field reference (export form)

- **Required / common:** `base_url`, `projects`  
- **Filters:** `issue_types`, `statuses`, `assignees`, `labels`  
- **Dates:** `date_field`, `start_dt`, `end_dt`, `time_block_start`, `time_block_end`  
- **Limits:** `page_size` (1–100), `max_issues` (`0` = no cap)  
- **JQL:** `extra_jql`, `custom_jql` (override replaces builder)  
- **Checkboxes:** `include_comments`, `include_workflow_events`, `verify_ssl`

---

## CSMS period windows (reference)

For report date `R` and period length `N` days:

- **Period 2:** `R-(N-1)` through `R`  
- **Period 1:** the preceding `N` days ending the day before Period 2 starts  

---

## Summary CSV schema (legacy export)

Includes core snapshot columns, SLA fields (`First Response Date`, `Resolution`, `Resolution Date`, time-to-response/resolution, SLA breached flags), `Status Transition Count`, `Status Overflow Count`, `Status Path`, and `Status 1..30` From/To/Timestamp/Author. Activity CSV remains row-per-changelog-item.

---

## Batch / automation

- **UI:** repeat export runs with different project/date windows; filenames are timestamped.  
- **Scripted:** `POST http://127.0.0.1:5001/run-export` with JSON body (see previous README examples); same pattern for `/run-csms-exec-summary`, `/run-team-posture`, etc.

---

## Troubleshooting

- **401 / anonymous:** restart the app after setting `JIRA_*` env vars; confirm Auth tab.  
- **400/500 from Jira:** JSON error responses often include server `details`.  
- **Team counts look wrong:** compare `broad_jql` from the API response to Issue Navigator — **`created`** window must include the tickets you expect; member **username** must match a Jira user field (`name`, `key`, `emailAddress`, `displayName`, …).  
- **SSL:** uncheck **Verify SSL** only if appropriate for your environment.

---

## Default Jira search URL (example)

```text
https://jira.mdthink.maryland.gov/rest/api/2/search
```

Search requests use `expand=changelog` where applicable and paginate until `max_issues` or end of results.
