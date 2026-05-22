# CSMS Reporting Notes

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
