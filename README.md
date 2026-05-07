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
| `POST` | `/run-legacy-dashboard` | Legacy trends + chart data (JSON). |
| `POST` | `/run-team-posture` | Team posture JSON for one member; includes `jql`, `broad_jql`, `metrics`, `warnings`, exports. |
| `POST` | `/run-team-posture-board-export` | Team board export for all submitted members; includes dashboard-bucket tagging rows for matched tickets. |
| `POST` | `/auth-status` | Auth / visibility diagnostics JSON. |
| `GET` | `/download?path=...` | Legacy export file download (server-side path). |
| `GET` | `/download-csms-export?export_id=&kind=` | Cached CSMS export (`csv_zip`, `excel`, `pdf`). |
| `GET` | `/download-team-posture-export?export_id=&kind=` | Cached team export (`csv`, `excel`). |

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

- **Team roster:** add/remove members (display name + **Assignee username**); roster stored in **browser localStorage**.  
- **Jira queries:**  
  - **`jql`** — same filters as broad query **plus** an `assignee in (...)` clause for the selected member’s Jira username.  
  - **`broad_jql`** — project, issue type, **`created`** date range only (no assignee). Used for worked-on, reopened-in-scope, contributed resolved, and SLA scope over member-linked issues.  
- **Ownership:** assignee match **or** (CSD project + configured **CSD Assigned Developer** field, default `customfield_14700`) when that field identifies the developer.  
- **Metric cards:**  
  - **Resolved (Owned)** — owned tickets whose status matches resolved-style rollups (e.g. resolved, closed, ready for production users, completed, duplicate, dev-completed).  
  - **Resolved (Last 8 Hours)** — owned tickets resolved within the last 8 hours.  
  - **Resolved (Contributed)** — same status rollups, member authored ≥1 **status** changelog transition, **not** current owner.  
  - **Assigned Open** — open by project-specific final status.  
  - **Reopened** — current status name contains `reopened` / `re-opened` / `re opened`, and member is owner **or** (not owner but has status transitions as author).  
  - **Worked On (Assigned to Others)** — status-change author is member; current owner ≠ member.  
  - **SLA Breach Count** — **24h from `created`:** open tickets past 24h; closed tickets prefer Jira **Resolution SLA Breached**-style custom field (discovered via `/rest/api/2/field`), else fallback elapsed created → resolutiondate (or updated). Counts include relevant open + closed breaches in member scope.  
  - **Open &lt; 8h to SLA breach** — open tickets with under 8 hours remaining before the 24h window.  
  - **Oldest open** — key + age (days) + detail JSON.  
- **Other UI:** ticket count by status (owned issues), **Ticket Labels** pie (Chart.js) over **member scope** issues, CSV preview (first rows of raw export).  
- **Exports:** per-member **CSV** + **Excel** (raw tickets + summary metrics); **Download Team CSV** — `POST /run-team-posture-board-export` with all members and row-level dashboard bucket tags (one row per issue per matching bucket).  
- **Warnings:** if Jira returns 500 with `changelog` expansion, the app may retry without changelog (reopen / worked-on / contributed paths may be incomplete).

**Form fields:** `base_url`, `projects`, `start_dt`, `end_dt`, `issue_types`, `csd_assigned_dev_field`, `page_size`, `max_issues`, `verify_ssl`, plus per-request `assignee_username` and `member_name`.

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
