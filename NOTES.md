# CSMS Reporting Notes

## 2026-06-04

- **Snapshot audit / migrate:** `snapshots_db.audit_all_snapshots()`, `migrate_all_snapshots()`, `normalize_snapshot_metrics()` on every save. CLI: `python audit_snapshots.py` / `--migrate`. API: `GET /snapshots/audit`, `POST /snapshots/migrate`.

## 2026-05-27

- **Operations Team performance:** `POST /run-team-posture-refresh` fetches **one** broad team JQL (with changelog when Jira allows), computes all member dashboard metrics in memory, and caches the issue pool server-side (`pool_cache_id`). Refresh no longer downloads ticket rows or writes per-member export files.
- **Export on demand:** `POST /run-team-posture-board-export` returns slim CSV columns only: Member Name, Dashboard Bucket, Issue Key, Summary. Reuses `pool_cache_id` from the last refresh when possible.

## 2026-05-25

- **Official reports — Load saved settings:** Choosing a snapshot in the dropdown shows archived numbers only; **Load saved settings** (or **Rerun with saved settings**) is required to copy `params` into the form. `GET /snapshots/<id>/display` includes `params` for exec, ops, and legacy.
- **Official reports — Delete snapshot:** **Delete snapshot** on Executive, Operations Team, and Ticket trend tabs; confirms then `DELETE /snapshots/<id>?report_id=`. Switches to Live after delete. Manual comparison baselines are unchanged.
- **Operations Team — 8-hour cards:** **Worked Status (Last 8 Hours)**, **Worked Status (Others, Last 8 Hours)**, and **Resolved (Last 8 Hours)** do not show ▲/▼ trend markers (still compared internally for other uses where applicable).

## 2026-05-21

- **Operations Team — status rollups (team header):** After refresh, sums cached member metrics: **Team Queue Backlog**, **Team In Progress**, **Team Resolved (Report Period)**. Partial roster shows a note until **Refresh from Jira** completes for all members.
- **Operations Team — per-member cards:** **Queue Backlog** (CSSD: Under QA Analysis; CSD: New), **In Progress** (CSSD: open, not New/Under QA Analysis; CSD: open, not New), **Worked Status (Last 8 Hours)** (owned tickets with your status changelog in last 8h), **Worked Status (Others, Last 8 Hours)** (not your ticket; you changed status in last 8h). Changelog required; Jira 500 fallback may zero worked-status counts.
- **Resolved (Report Period):** Still computed per member (`resolutiondate` within Team Start/End among created-window issues) for **team rollup** and summary export — **no** per-member card on the grid.
- **Download Team CSV:** Uses session cache (`teamPayloadByMemberId` / `raw_rows`) when every roster member is cached; otherwise `POST /run-team-posture-board-export`. On error, partial cache download when available. Archived snapshots do not store ticket rows — refresh live first.

## 2026-05-20

- **Ticket trend SLA KPIs:** Four cards (TTFR/TTR × CSSD/CSD) on **Refresh Dashboard** (`POST /run-legacy-dashboard`).
- **Status gates:** `ttr_status_cssd` (default Closed), `ttr_status_csd` (default Ready For Production Users), `ttfr_status_*` optional.
- **Aggregates:** Per-card dropdowns `*_aggregate` — `median` (default), `mean`, `p90`.
- **TTR rule:** Prefer `customfield_10317` elapsedTime; optional calendar fallback `resolutiondate − created` via **Include calendar fallbacks** toggle.
- **TTFR rule:** Prefer `customfield_10318` elapsedTime; optional fallback SLA stop − created. CSD may inherit from linked CSSD. **Force hours view** keeps KPI cards in hours.
- **TTFR CSD:** Uses linked CSSD ticket’s `customfield_10318` when `issuelinks` point to CSSD; otherwise CSD’s own SLA field.
- **Snapshots:** `LEGACY_TREND_KEYS` extended with the four SLA hour metrics; aggregate choice stored in snapshot `params`.

## 2026-05-19

- **Pipeline Backlog:** Count uses Jira search `total` (`POST /run-pipeline-backlog-count`), not full issue pagination. Refresh loads pipeline before Team Closed (`POST /run-team-board-metrics` with `skip_pipeline: true` for closed-only).
- **Operations Team UI:** Refresh from Jira / Refresh selected member moved into **Team Posture Variables & Settings**. Metric card sparklines removed; delta % vs prior snapshot/baseline remains.
- **Live vs archive:** Official report dropdown; Live mode requires explicit refresh. Archive hydrates from `GET /snapshots/<id>/display` without Jira.
- **Team Closed:** Narrow JQL (`project in (CSSD, CSD)`, `status = Closed`, optional report `updated` window); capped issue fetch with changelog for contributed tickets.
- **Label trends:** `GET /snapshots/label-trends` for member label history across saved ops snapshots; Ticket Labels section shows Current + Trend charts.
- **Bulk refresh:** Serialized Jira queue; pipeline count runs in parallel with per-member posture during Refresh from Jira.
- **Mobile:** Responsive layout and hamburger nav for small screens.

## 2026-05-07

- Updated sidebar and page naming to match current UI: `Executive Report`, `Operations Team`, and `Ticket trend`, with app title shown as `CSMS Reporting`.
- Added optional Executive settings baselines for prior report KPIs: backlog, new created, and resolved ticket counts.
- Added Operations Team metric for `Resolved (Last 8 Hours)`.
- Added report-period labels on Operations Team and Ticket trend sections based on query start/end inputs.
- Expanded high-level metric/chart tooltips across Executive, Team, and Ticket trend pages.
- Updated team board CSV export behavior to include dashboard bucket tagging rows (one row per issue per matching bucket).
- Improved team export/download flow resilience around fetch/network failures.
