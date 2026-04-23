
import csv
import io
import json
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
from flask import Flask, Response, jsonify, render_template_string, request, send_file
from requests.auth import HTTPBasicAuth

app = Flask(__name__)
STATUS_TRANSITION_SLOTS = 30

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Jira Export Helper</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root {
      --bg: #0b1020;
      --panel: #121933;
      --panel-2: #1a2448;
      --text: #e8ecff;
      --muted: #9fb0e5;
      --accent: #79a8ff;
      --accent-2: #4fd1c5;
      --danger: #ff7b7b;
      --border: #2a376a;
      --body-grad-start: #09101d;
      --body-grad-end: #0d1530;
      --card-bg: rgba(18, 25, 51, 0.92);
      --pre-bg: #0b1228;
      --pre-text: #dbe7ff;
      --download-bg: #18234a;
      --download-text: #ffffff;
    }
    :root[data-theme="light"] {
      --bg: #f3f7ff;
      --panel: #ffffff;
      --panel-2: #eef3ff;
      --text: #12213f;
      --muted: #4c5f86;
      --accent: #2f6fe8;
      --accent-2: #29b7a8;
      --border: #d4def0;
      --body-grad-start: #f7faff;
      --body-grad-end: #edf3ff;
      --card-bg: rgba(255, 255, 255, 0.96);
      --pre-bg: #eef3ff;
      --pre-text: #20345e;
      --download-bg: #e6edff;
      --download-text: #1b2d52;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, Arial, sans-serif;
      background: linear-gradient(180deg, var(--body-grad-start) 0%, var(--body-grad-end) 100%);
      color: var(--text);
    }
    .wrap {
      max-width: 1180px;
      margin: 0 auto;
      padding: 28px 18px 48px;
    }
    .hero {
      display: grid;
      grid-template-columns: 1.15fr .85fr;
      gap: 18px;
      margin-bottom: 18px;
    }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 18px;
      padding: 18px;
      box-shadow: 0 10px 30px rgba(0,0,0,.24);
    }
    .hero-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 10px;
    }
    h1 { margin: 0 0 10px; font-size: 30px; }
    h2 { margin: 0 0 12px; font-size: 18px; }
    p { color: var(--muted); line-height: 1.45; }
    form {
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 14px;
    }
    .field { grid-column: span 4; }
    .field.wide { grid-column: span 6; }
    .field.full { grid-column: span 12; }
    label {
      display: block;
      font-size: 12px;
      color: var(--muted);
      margin-bottom: 6px;
      text-transform: uppercase;
      letter-spacing: .06em;
    }
    input, select, textarea {
      width: 100%;
      background: var(--panel-2);
      border: 1px solid var(--border);
      color: var(--text);
      border-radius: 12px;
      padding: 11px 12px;
      font-size: 14px;
    }
    textarea { min-height: 88px; resize: vertical; }
    .row {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }
    .check {
      display: flex;
      align-items: center;
      gap: 8px;
      background: var(--panel-2);
      border: 1px solid var(--border);
      padding: 10px 12px;
      border-radius: 12px;
    }
    .check input { width: auto; }
    button {
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      border: 0;
      color: #08101f;
      font-weight: 700;
      padding: 12px 16px;
      border-radius: 14px;
      cursor: pointer;
    }
    .muted-btn {
      background: transparent;
      color: var(--text);
      border: 1px solid var(--border);
    }
    .theme-toggle {
      padding: 8px 12px;
      border-radius: 12px;
      font-size: 13px;
      font-weight: 600;
      white-space: nowrap;
    }
    pre {
      background: var(--pre-bg);
      padding: 14px;
      border-radius: 12px;
      color: var(--pre-text);
      white-space: pre-wrap;
      word-break: break-word;
      border: 1px solid var(--border);
    }
    .downloads a {
      display: inline-block;
      margin-right: 10px;
      margin-top: 10px;
      color: var(--download-text);
      text-decoration: none;
      padding: 10px 12px;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: var(--download-bg);
    }
    .small {
      font-size: 12px;
      color: var(--muted);
    }
    @media (max-width: 900px) {
      .hero { grid-template-columns: 1fr; }
      .field, .field.wide { grid-column: span 12; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div class="card">
        <div class="hero-head">
          <h1>Jira Insights Exporter</h1>
          <button type="button" class="muted-btn theme-toggle" id="themeToggle">Switch to Light</button>
        </div>
        <p>
          Build and run Jira exports with flexible filters, then download clean summary and
          activity datasets in seconds.
        </p>
        <p class="small">
          Project-ready exports with status-path analytics, SLA visibility, batch-friendly output,
          and CSV/ZIP delivery.
        </p>
      </div>
      <div class="card">
        <h2>Notes</h2>
        <pre id="notes">
Base endpoint example:
https://jira.mdthink.maryland.gov/rest/api/2/search

Status movement output in summary export:
- Status Transition Count
- Status Overflow Count
- Status Path
- Status 1..30 From/To/Timestamp/Author

Credentials are read from:
JIRA_USERNAME / JIRA_PASSWORD
or
JIRA_EMAIL / JIRA_API_TOKEN
        </pre>
      </div>
    </div>

    <div class="card">
      <form id="exportForm">
        <div class="field wide">
          <label>Jira Search Endpoint</label>
          <input name="base_url" value="https://jira.mdthink.maryland.gov/rest/api/2/search" placeholder="https://your-jira-domain/rest/api/2/search" />
        </div>

        <div class="field wide">
          <label>Projects (comma separated)</label>
          <input name="projects" placeholder="CSSD,ABC (project keys, not names)" />
        </div>

        <div class="field">
          <label>Issue Types (comma separated)</label>
          <input name="issue_types" placeholder="Bug, Task, Story (exact Jira issue type names)" />
        </div>

        <div class="field">
          <label>Current Statuses (comma separated)</label>
          <input name="statuses" placeholder="Open, In Progress, Closed (use exact status names)" />
        </div>

        <div class="field">
          <label>Assignees (comma separated)</label>
          <input name="assignees" placeholder="username1,username2 (as Jira expects in JQL)" />
        </div>

        <div class="field">
          <label>Labels (comma separated)</label>
          <input name="labels" placeholder="defect, uat" />
        </div>

        <div class="field">
          <label>Date Field</label>
          <select name="date_field">
            <option value="created">created</option>
            <option value="updated">updated</option>
            <option value="resolutiondate">resolutiondate</option>
          </select>
        </div>

        <div class="field">
          <label>Start Date/Time</label>
          <input type="datetime-local" name="start_dt" />
        </div>

        <div class="field">
          <label>End Date/Time</label>
          <input type="datetime-local" name="end_dt" />
        </div>

        <div class="field">
          <label>Time Block Start</label>
          <input type="time" name="time_block_start" />
        </div>

        <div class="field">
          <label>Time Block End</label>
          <input type="time" name="time_block_end" />
        </div>

        <div class="field">
          <label>Page Size</label>
          <input type="number" name="page_size" value="50" min="1" max="100" />
        </div>

        <div class="field">
          <label>Max Issues (0 = all)</label>
          <input type="number" name="max_issues" value="0" min="0" />
        </div>

        <div class="field full">
          <label>Extra JQL (AND appended)</label>
          <textarea name="extra_jql" placeholder='Example: component = "Establishment" AND priority = Medium'></textarea>
        </div>

        <div class="field full">
          <label>Custom JQL Override (optional)</label>
          <textarea name="custom_jql" placeholder='Example: issuekey = CSSD-123 (when set, this replaces all builder fields)'></textarea>
        </div>

        <div class="field full">
          <div class="row">
            <label class="check"><input type="checkbox" name="include_comments" checked /> Include comments in summary metrics</label>
            <label class="check"><input type="checkbox" name="include_workflow_events" /> Include workflow admin changes in activity export</label>
            <label class="check"><input type="checkbox" name="verify_ssl" checked /> Verify SSL</label>
          </div>
        </div>

        <div class="field full">
          <div class="row">
            <button type="submit">Run export</button>
            <button type="button" class="muted-btn" id="previewBtn">Preview JQL</button>
          </div>
        </div>
      </form>
    </div>

    <div class="card" style="margin-top:18px;">
      <h2>Result</h2>
      <pre id="result">Ready.</pre>
      <div class="downloads" id="downloads"></div>
    </div>
  </div>

<script>
const THEME_PARAMS = {
  light: {
    "--bg": "#f3f7ff",
    "--panel": "#ffffff",
    "--panel-2": "#eef3ff",
    "--text": "#12213f",
    "--muted": "#4c5f86",
    "--accent": "#2f6fe8",
    "--accent-2": "#29b7a8",
    "--border": "#d4def0",
    "--body-grad-start": "#f7faff",
    "--body-grad-end": "#edf3ff",
    "--card-bg": "rgba(255, 255, 255, 0.96)",
    "--pre-bg": "#eef3ff",
    "--pre-text": "#20345e",
    "--download-bg": "#e6edff",
    "--download-text": "#1b2d52"
  },
  dark: {
    "--bg": "#0b1020",
    "--panel": "#121933",
    "--panel-2": "#1a2448",
    "--text": "#e8ecff",
    "--muted": "#9fb0e5",
    "--accent": "#79a8ff",
    "--accent-2": "#4fd1c5",
    "--border": "#2a376a",
    "--body-grad-start": "#09101d",
    "--body-grad-end": "#0d1530",
    "--card-bg": "rgba(18, 25, 51, 0.92)",
    "--pre-bg": "#0b1228",
    "--pre-text": "#dbe7ff",
    "--download-bg": "#18234a",
    "--download-text": "#ffffff"
  }
};

function applyTheme(theme) {
  const chosen = THEME_PARAMS[theme] ? theme : "dark";
  const vars = THEME_PARAMS[chosen];
  document.documentElement.setAttribute("data-theme", chosen);
  for (const [key, value] of Object.entries(vars)) {
    document.documentElement.style.setProperty(key, value);
  }
  localStorage.setItem("theme", chosen);
  const toggle = document.getElementById("themeToggle");
  toggle.textContent = chosen === "dark" ? "Switch to Light" : "Switch to Dark";
}

document.getElementById("themeToggle").addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "dark";
  applyTheme(current === "dark" ? "light" : "dark");
});

applyTheme(localStorage.getItem("theme") || "dark");

function formToObject(form) {
  const obj = {};
  const fd = new FormData(form);
  for (const [key, value] of fd.entries()) {
    if (obj[key] !== undefined) continue;
    obj[key] = value;
  }
  obj.include_comments = form.include_comments.checked;
  obj.include_workflow_events = form.include_workflow_events.checked;
  obj.verify_ssl = form.verify_ssl.checked;
  return obj;
}

document.getElementById("previewBtn").addEventListener("click", async () => {
  const payload = formToObject(document.getElementById("exportForm"));
  const res = await fetch("/preview-jql", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(payload)
  });
  const data = await res.json();
  document.getElementById("result").textContent = JSON.stringify(data, null, 2);
  document.getElementById("downloads").innerHTML = "";
});

document.getElementById("exportForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = formToObject(e.target);
  document.getElementById("result").textContent = "Running...";
  document.getElementById("downloads").innerHTML = "";

  const res = await fetch("/run-export", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(payload)
  });
  const data = await res.json();
  document.getElementById("result").textContent = JSON.stringify(data, null, 2);

  if (data.downloads) {
    const parts = [];
    for (const item of data.downloads) {
      parts.push(`<a href="${item.url}">${item.label}</a>`);
    }
    document.getElementById("downloads").innerHTML = parts.join("");
  }
});
</script>
</body>
</html>
"""


def parse_csv_list(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


def normalize_dt_local(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip()
    # Converts "2026-04-21T10:30" to Jira-friendly quoted value.
    return value.replace("T", " ")


def parse_time_value(value: Optional[str]) -> Optional[time]:
    if not value:
        return None
    return datetime.strptime(value, "%H:%M").time()


def jql_quote(val: str) -> str:
    escaped = val.replace('"', '\\"')
    return f'"{escaped}"'


def list_clause(field: str, values: List[str]) -> Optional[str]:
    if not values:
        return None
    quoted = ", ".join(jql_quote(v) for v in values)
    return f'{field} in ({quoted})'


def build_jql(params: Dict[str, Any]) -> str:
    custom_jql = (params.get("custom_jql") or "").strip()
    if custom_jql:
        return custom_jql

    clauses: List[str] = []
    projects = parse_csv_list(params.get("projects"))
    issue_types = parse_csv_list(params.get("issue_types"))
    statuses = parse_csv_list(params.get("statuses"))
    assignees = parse_csv_list(params.get("assignees"))
    labels = parse_csv_list(params.get("labels"))

    for clause in [
        list_clause("project", projects),
        list_clause("issuetype", issue_types),
        list_clause("status", statuses),
        list_clause("assignee", assignees),
        list_clause("labels", labels),
    ]:
        if clause:
            clauses.append(clause)

    date_field = params.get("date_field") or "created"
    start_dt = normalize_dt_local(params.get("start_dt"))
    end_dt = normalize_dt_local(params.get("end_dt"))
    if start_dt:
        clauses.append(f'{date_field} >= "{start_dt}"')
    if end_dt:
        clauses.append(f'{date_field} <= "{end_dt}"')

    extra_jql = (params.get("extra_jql") or "").strip()
    if extra_jql:
        clauses.append(f"({extra_jql})")

    return " AND ".join(clauses) if clauses else "order by created desc"


def get_auth() -> Optional[HTTPBasicAuth]:
    username = os.getenv("JIRA_USERNAME") or os.getenv("JIRA_EMAIL")
    password = os.getenv("JIRA_PASSWORD") or os.getenv("JIRA_API_TOKEN")
    if username and password:
        return HTTPBasicAuth(username, password)
    return None


def flatten_user(user_obj: Optional[Dict[str, Any]]) -> str:
    if not user_obj:
        return ""
    return user_obj.get("displayName") or user_obj.get("name") or user_obj.get("emailAddress") or ""


def option_value(obj: Any) -> str:
    if obj is None:
        return ""
    if isinstance(obj, dict):
        return obj.get("value") or obj.get("name") or obj.get("displayName") or obj.get("key") or json.dumps(obj)
    if isinstance(obj, list):
        vals = [option_value(item) for item in obj]
        return "; ".join([v for v in vals if v])
    return str(obj)


def pick_severity(fields: Dict[str, Any]) -> str:
    for key in ("customfield_10707", "customfield_Severity", "severity"):
        if key in fields:
            return option_value(fields.get(key))
    return ""


def extract_sla(fields: Dict[str, Any], key: str) -> Tuple[str, str, str]:
    data = fields.get(key) or {}
    cycles = data.get("completedCycles") or []
    if not cycles:
        return "", "", ""
    cycle = cycles[0]
    elapsed = (cycle.get("elapsedTime") or {}).get("friendly", "")
    breached = str(cycle.get("breached", "")).title() if "breached" in cycle else ""
    stop = (cycle.get("stopTime") or {}).get("jira") or (cycle.get("stopTime") or {}).get("iso8601", "")
    return elapsed, breached, stop


def is_workflow_event(field_name: str) -> bool:
    return field_name.strip().lower() == "workflow"


def filter_by_time_block(dt_str: str, block_start: Optional[time], block_end: Optional[time]) -> bool:
    if not dt_str or (block_start is None and block_end is None):
        return True
    try:
        dt = datetime.strptime(dt_str[:19], "%Y-%m-%dT%H:%M:%S")
    except Exception:
        try:
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        except Exception:
            return True

    t = dt.time()
    if block_start and block_end:
        if block_start <= block_end:
            return block_start <= t <= block_end
        return t >= block_start or t <= block_end
    if block_start:
        return t >= block_start
    if block_end:
        return t <= block_end
    return True


def extract_comments_count(fields: Dict[str, Any]) -> int:
    comment = fields.get("comment") or {}
    return int(comment.get("total") or len(comment.get("comments") or []))


def fetch_issues(base_url: str, jql: str, page_size: int, max_issues: int, verify_ssl: bool) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    start_at = 0
    auth = get_auth()
    session = requests.Session()

    while True:
        payload = {
            "jql": jql,
            "expand": "changelog",
            "startAt": start_at,
            "maxResults": page_size,
        }
        resp = session.get(base_url, params=payload, auth=auth, verify=verify_ssl, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        batch = data.get("issues", [])
        issues.extend(batch)

        total = int(data.get("total", 0))
        start_at += len(batch)

        if not batch:
            break
        if max_issues and len(issues) >= max_issues:
            issues = issues[:max_issues]
            break
        if start_at >= total:
            break

    return issues


def make_activity_rows(issue: Dict[str, Any], include_workflow_events: bool, block_start: Optional[time], block_end: Optional[time]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    key = issue.get("key", "")
    histories = ((issue.get("changelog") or {}).get("histories") or [])

    for history in histories:
        change_date = history.get("created", "")
        if not filter_by_time_block(change_date, block_start, block_end):
            continue
        author = flatten_user(history.get("author"))
        for item in history.get("items", []):
            field_name = item.get("field", "")
            if not include_workflow_events and is_workflow_event(field_name):
                continue
            rows.append({
                "Issue Key": key,
                "Change Date": change_date,
                "Author": author,
                "Field": field_name,
                "From": item.get("fromString") or "",
                "To": item.get("toString") or "",
            })
    return rows


def ordered_status_transitions(activity_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return sorted(
        [row for row in activity_rows if row["Field"].lower() == "status"],
        key=lambda r: r["Change Date"],
    )


def first_status_transition(activity_rows: List[Dict[str, Any]], target_status: str) -> Dict[str, Any]:
    target = target_status.strip().lower()
    for row in ordered_status_transitions(activity_rows):
        if (row["To"] or "").strip().lower() == target:
            return row
    return {
        "Change Date": "",
        "Author": "",
        "From": "",
        "To": "",
    }


def first_status_date(activity_rows: List[Dict[str, Any]], target_status: str) -> str:
    return first_status_transition(activity_rows, target_status)["Change Date"]


def wide_status_transition_fields(status_transitions: List[Dict[str, Any]], slots: int = STATUS_TRANSITION_SLOTS) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}
    for index in range(1, slots + 1):
        row = status_transitions[index - 1] if index <= len(status_transitions) else None
        fields[f"Status {index} From"] = (row or {}).get("From", "")
        fields[f"Status {index} To"] = (row or {}).get("To", "")
        fields[f"Status {index} Timestamp"] = (row or {}).get("Change Date", "")
        fields[f"Status {index} Author"] = (row or {}).get("Author", "")
    return fields


def wide_status_transition_columns(slots: int = STATUS_TRANSITION_SLOTS) -> List[str]:
    columns: List[str] = []
    for index in range(1, slots + 1):
        columns.extend([
            f"Status {index} From",
            f"Status {index} To",
            f"Status {index} Timestamp",
            f"Status {index} Author",
        ])
    return columns


def first_resolution_date(activity_rows: List[Dict[str, Any]]) -> str:
    for row in sorted(activity_rows, key=lambda r: r["Change Date"]):
        if row["Field"].lower() == "resolution" and row["To"]:
            return row["Change Date"]
    return ""


def build_summary_row(issue: Dict[str, Any], activity_rows: List[Dict[str, Any]], include_comments: bool) -> Dict[str, Any]:
    fields = issue.get("fields") or {}

    ttr_elapsed, ttr_breached, _ = extract_sla(fields, "customfield_10317")
    ttfr_elapsed, ttfr_breached, ttfr_stop = extract_sla(fields, "customfield_10318")

    comments_total = extract_comments_count(fields)
    current_status = option_value(fields.get("status"))
    resolution = option_value(fields.get("resolution"))
    status_transitions = ordered_status_transitions(activity_rows)
    transition_overflow_count = max(0, len(status_transitions) - STATUS_TRANSITION_SLOTS)
    status_path = " || ".join(
        f'{row["Change Date"]}|{row["Author"]}|{row["From"]}->{row["To"]}'
        for row in status_transitions
    )

    summary_row = {
        "Issue Key": issue.get("key", ""),
        "Summary": fields.get("summary", ""),
        "Issue Type": option_value(fields.get("issuetype")),
        "Priority": option_value(fields.get("priority")),
        "Severity": pick_severity(fields),
        "Assignee": flatten_user(fields.get("assignee")),
        "Reporter": flatten_user(fields.get("reporter")),
        "Component": option_value(fields.get("components")),
        "Label(s)": option_value(fields.get("labels")),
        "Created Date": fields.get("created", ""),
        "First Response Date": ttfr_stop,
        "Resolution": resolution,
        "Resolution Date": first_resolution_date(activity_rows),
        "Time to First Response": ttfr_elapsed,
        "First Response SLA Breached": ttfr_breached,
        "Time to Resolution": ttr_elapsed,
        "Resolution SLA Breached": ttr_breached,
        "Current Status": current_status,
        "Updated Date": fields.get("updated", ""),
        "Comment Count": comments_total if include_comments else "",
        "Status Transition Count": len(status_transitions),
        "Status Overflow Count": transition_overflow_count,
        "Status Path": status_path,
    }
    summary_row.update(wide_status_transition_fields(status_transitions))
    return summary_row


def write_csv(path: Path, rows: List[Dict[str, Any]], columns: List[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def export_run(params: Dict[str, Any]) -> Dict[str, Any]:
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")

    page_size = int(params.get("page_size") or 50)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    include_comments = bool(params.get("include_comments", True))
    include_workflow_events = bool(params.get("include_workflow_events", False))

    jql = build_jql(params)
    block_start = parse_time_value(params.get("time_block_start"))
    block_end = parse_time_value(params.get("time_block_end"))

    issues = fetch_issues(base_url, jql, page_size, max_issues, verify_ssl)

    summary_rows: List[Dict[str, Any]] = []
    activity_rows: List[Dict[str, Any]] = []

    for issue in issues:
        issue_activity = make_activity_rows(issue, include_workflow_events, block_start, block_end)
        activity_rows.extend(issue_activity)
        summary_rows.append(build_summary_row(issue, issue_activity, include_comments))

    out_dir = Path(tempfile.mkdtemp(prefix="jira_export_"))
    generated_at = datetime.now().astimezone()
    generated_at_iso = generated_at.isoformat(timespec="seconds")
    generated_stamp = generated_at.strftime("%Y%m%d_%H%M%S")
    summary_path = out_dir / f"issue_summary_{generated_stamp}.csv"
    activity_path = out_dir / f"issue_activity_{generated_stamp}.csv"
    zip_path = out_dir / f"jira_exports_{generated_stamp}.zip"
    meta_path = out_dir / f"run_metadata_{generated_stamp}.json"

    summary_columns = [
        "Issue Key", "Summary", "Issue Type", "Priority", "Severity", "Assignee", "Reporter",
        "Component", "Label(s)", "Created Date", "First Response Date", "Resolution",
        "Resolution Date", "Time to First Response", "First Response SLA Breached",
        "Time to Resolution", "Resolution SLA Breached", "Current Status", "Updated Date",
        "Comment Count", "Status Transition Count", "Status Overflow Count", "Status Path",
    ]
    summary_columns.extend(wide_status_transition_columns())
    activity_columns = ["Issue Key", "Change Date", "Author", "Field", "From", "To"]

    write_csv(summary_path, summary_rows, summary_columns)
    write_csv(activity_path, activity_rows, activity_columns)

    metadata = {
        "generated_at": generated_at_iso,
        "jql": jql,
        "issue_count": len(issues),
        "summary_row_count": len(summary_rows),
        "activity_row_count": len(activity_rows),
        "time_block_start": params.get("time_block_start") or "",
        "time_block_end": params.get("time_block_end") or "",
    }
    meta_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(summary_path, summary_path.name)
        zf.write(activity_path, activity_path.name)
        zf.write(meta_path, meta_path.name)

    return {
        "jql": jql,
        "issue_count": len(issues),
        "summary_rows": len(summary_rows),
        "activity_rows": len(activity_rows),
        "files": {
            "summary": str(summary_path),
            "activity": str(activity_path),
            "zip": str(zip_path),
            "metadata": str(meta_path),
        }
    }


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/preview-jql", methods=["POST"])
def preview_jql():
    params = request.get_json(force=True)
    return jsonify({"jql": build_jql(params)})


@app.route("/run-export", methods=["POST"])
def run_export():
    try:
        result = export_run(request.get_json(force=True))
        downloads = [
            {"label": "Download summary CSV", "url": f"/download?path={result['files']['summary']}"},
            {"label": "Download activity CSV", "url": f"/download?path={result['files']['activity']}"},
            {"label": "Download ZIP bundle", "url": f"/download?path={result['files']['zip']}"},
        ]
        result["downloads"] = downloads
        return jsonify(result)
    except requests.HTTPError as exc:
        try:
            details = exc.response.text
        except Exception:
            details = str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/download")
def download():
    path = request.args.get("path", "")
    if not path:
        return Response("Missing path", status=400)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5001)
