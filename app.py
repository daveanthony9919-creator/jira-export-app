
import csv
import io
import json
import os
import re
import tempfile
import zipfile
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
from flask import Flask, Response, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from requests.auth import HTTPBasicAuth

app = Flask(__name__)
STATUS_TRANSITION_SLOTS = 30
CSMS_EXPORT_CACHE: Dict[str, Dict[str, str]] = {}

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Jira Export Helper</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root {
      --bg: #07132c;
      --panel: #0f1d3a;
      --panel-2: #1e2f4e;
      --text: #e8ecff;
      --muted: #95a6cc;
      --accent: #36e0d0;
      --accent-2: #2bb5ff;
      --danger: #ff7098;
      --border: #21385f;
      --body-grad-start: #07122a;
      --body-grad-end: #0a1734;
      --card-bg: rgba(14, 28, 57, 0.95);
      --pre-bg: #0b1228;
      --pre-text: #dbe7ff;
      --download-bg: #18234a;
      --download-text: #ffffff;
      --btn-bg-start: #435f8e;
      --btn-bg-end: #2b3d62;
      --btn-border: #5c7ab0;
      --btn-text: #f3f8ff;
      --btn-shadow: rgba(2, 10, 27, 0.35);
      --btn-muted-bg-start: #223250;
      --btn-muted-bg-end: #1b2a47;
      --btn-muted-border: #4d6a9e;
      --btn-muted-text: #dce8ff;
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
      --btn-bg-start: #5d8dff;
      --btn-bg-end: #3f74e6;
      --btn-border: #2f61cd;
      --btn-text: #ffffff;
      --btn-shadow: rgba(31, 70, 148, 0.2);
      --btn-muted-bg-start: #f5f8ff;
      --btn-muted-bg-end: #e8efff;
      --btn-muted-border: #c4d4f4;
      --btn-muted-text: #16315f;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, Arial, sans-serif;
      background: linear-gradient(180deg, var(--body-grad-start) 0%, var(--body-grad-end) 100%);
      color: var(--text);
    }
    .wrap {
      max-width: 1320px;
      margin: 0 auto;
      padding: 0 20px 48px;
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
      border-radius: 14px;
      padding: 16px;
      box-shadow: 0 10px 30px rgba(0,0,0,.24);
    }
    .hero-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 10px;
    }
    h1 { margin: 0 0 10px; font-size: 44px; line-height: 1.05; }
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
      background: linear-gradient(180deg, var(--btn-bg-start), var(--btn-bg-end));
      border: 1px solid var(--btn-border);
      color: var(--btn-text);
      font-weight: 700;
      font-size: 13px;
      letter-spacing: .01em;
      padding: 10px 14px;
      border-radius: 10px;
      cursor: pointer;
      transition: transform .12s ease, box-shadow .15s ease, border-color .15s ease, filter .15s ease;
      box-shadow: 0 4px 10px var(--btn-shadow);
    }
    button:hover {
      transform: translateY(-1px);
      border-color: var(--btn-border);
      box-shadow: 0 7px 14px var(--btn-shadow);
      filter: brightness(1.03);
    }
    button:active {
      transform: translateY(0);
      box-shadow: 0 3px 7px var(--btn-shadow);
    }
    .muted-btn {
      background: linear-gradient(180deg, var(--btn-muted-bg-start), var(--btn-muted-bg-end));
      color: var(--btn-muted-text);
      border: 1px solid var(--btn-muted-border);
    }
    .theme-toggle {
      width: 40px;
      height: 40px;
      padding: 0;
      border-radius: 50%;
      font-size: 16px;
      font-weight: 700;
      display: inline-flex;
      align-items: center;
      justify-content: center;
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
    .app-nav {
      position: sticky;
      top: 0;
      z-index: 10;
      margin-bottom: 18px;
      backdrop-filter: blur(8px);
      width: 100vw;
      max-width: 100vw;
      margin-left: calc(50% - 50vw);
      margin-right: calc(50% - 50vw);
      border-radius: 0;
    }
    .app-nav-shell {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
    }
    .app-nav-title {
      font-size: 40px;
      font-weight: 800;
      letter-spacing: .01em;
      margin: 0;
      color: var(--text);
      white-space: nowrap;
    }
    .app-nav-actions {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }
    .auth-icon-btn {
      width: 40px;
      height: 40px;
      padding: 0;
      border-radius: 50%;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
      line-height: 1;
    }
    .nav-link.active,
    .report-tab.active {
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      color: #08101f;
      border-color: transparent;
      box-shadow: 0 6px 14px rgba(24, 204, 206, 0.25);
    }
    .app-section {
      scroll-margin-top: 94px;
      margin-bottom: 14px;
    }
    .two-col {
      display: grid;
      grid-template-columns: 1.2fr .95fr;
      gap: 14px;
    }
    .kpi-grid .kpi-card {
      min-height: 126px;
      background: var(--panel-2);
    }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(180px, 1fr));
      gap: 14px;
    }
    .tooltip-panel {
      margin-top: 10px;
      padding: 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel-2);
      color: var(--muted);
      font-size: 13px;
    }
    .legacy-chart-wrap {
      width: 100%;
      position: relative;
      overflow: hidden;
      border-radius: 10px;
    }
    .legacy-chart-wrap.daily {
      height: 200px; /* ~40% of prior ~500px */
    }
    .legacy-chart-wrap.status {
      height: 500px; /* ~40% of prior ~1246px */
    }
    .csms-chart-wrap {
      width: 100%;
      position: relative;
      overflow: hidden;
      border-radius: 10px;
      background: rgba(9, 19, 43, 0.45);
      border: 1px solid var(--border);
      padding: 6px;
    }
    .csms-chart-wrap.daily {
      height: 180px; /* ~40% smaller visual footprint */
    }
    .csms-chart-wrap.small {
      height: 150px; /* ~40% smaller visual footprint */
    }
    .csms-chart-grid {
      display: grid;
      grid-template-columns: 1.25fr .75fr;
      gap: 10px;
      margin-top: 10px;
    }
    .report-scope-csms .section-title { font-size: 24px; }
    .report-scope-csms .section-subtitle { font-size: 14px; letter-spacing: .05em; }
    .report-scope-csms .status-pill { font-size: 16px; min-width: 220px; padding: 8px 12px; }
    .report-scope-csms .kpi-grid { grid-template-columns: repeat(4, minmax(140px, 1fr)); gap: 10px; }
    .report-scope-csms .kpi-grid .kpi-card { min-height: 96px; padding: 10px; }
    .report-scope-csms .kpi-label { font-size: 14px; }
    .report-scope-csms .kpi-number { font-size: 22px; margin: 4px 0; }
    .report-scope-csms .kpi-trend { font-size: 13px; }
    .report-scope-csms .large-pre { min-height: 165px; font-size: 14px; line-height: 1.4; }
    .report-scope-csms .health-panel { min-height: 165px; font-size: 16px; }
    .report-scope-csms .health-panel h3 { font-size: 22px; }
    .report-scope-csms .elapsed-note { font-size: 12px; color: var(--muted); margin: 6px 0 0; }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }
    .tab-btn.active {
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      color: #08101f;
    }
    .kpi-card {
      background: var(--panel-2);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 14px;
      min-width: 180px;
      flex: 1;
    }
    .kpi-number { font-size: 52px; font-weight: 800; margin: 8px 0 6px; color: var(--accent); }
    .kpi-label { color: var(--muted); font-size: 23px; letter-spacing: .01em; text-transform: uppercase; }
    .kpi-trend { font-size: 22px; font-weight: 700; }
    .trend-pos { color: #4fd1c5; }
    .trend-neg { color: #ff7098; }
    .section-title {
      font-size: 34px;
      margin: 0;
      letter-spacing: .01em;
    }
    .section-subtitle {
      font-size: 26px;
      letter-spacing: .08em;
      color: var(--muted);
      text-transform: uppercase;
      margin: 8px 0 0;
    }
    .elapsed-note {
      margin-top: 8px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.35;
    }
    .status-pill {
      background: var(--panel-2);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 12px 16px;
      font-size: 24px;
      color: var(--muted);
      min-width: 300px;
      text-align: right;
    }
    .status-pill strong { color: var(--text); }
    .large-pre {
      min-height: 220px;
      font-size: 25px;
      line-height: 1.45;
    }
    .health-panel {
      min-height: 220px;
      font-size: 24px;
    }
    .health-panel h3 {
      margin: 0 0 10px;
      font-size: 30px;
      color: var(--accent);
    }
    .progress-wrap {
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel-2);
      height: 14px;
      overflow: hidden;
      margin: 8px 0;
    }
    .progress-bar {
      height: 100%;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
    }
    .collapse-card { padding-top: 14px; }
    .collapse-toggle {
      width: 100%;
      text-align: left;
      font-weight: 700;
    }
    .collapse-body {
      margin-top: 14px;
      border-top: 1px solid var(--border);
      padding-top: 14px;
    }
    @media (max-width: 900px) {
      .hero { grid-template-columns: 1fr; }
      .field, .field.wide { grid-column: span 12; }
      .two-col { grid-template-columns: 1fr; }
      .kpi-grid { grid-template-columns: repeat(2, minmax(160px, 1fr)); }
      .kpi-number { font-size: 36px; }
      .section-title { font-size: 32px; }
      .section-subtitle { font-size: 18px; }
      .csms-chart-grid { grid-template-columns: 1fr; }
      .app-nav-title { font-size: 28px; }
      .app-nav-shell { align-items: flex-start; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card app-nav">
      <div class="app-nav-shell">
        <h1 class="app-nav-title">CSMS Operations</h1>
        <div class="app-nav-actions">
          <button type="button" class="muted-btn report-tab active" data-report="csms">CSMS Report</button>
          <button type="button" class="muted-btn report-tab" data-report="legacy">General Report</button>
          <button type="button" class="muted-btn" id="notesInfoBtn" aria-expanded="false" aria-controls="notesTooltip">Notes</button>
          <button type="button" class="muted-btn theme-toggle" id="themeToggle" aria-label="Toggle theme" title="Switch theme">◐</button>
          <button type="button" class="muted-btn report-tab auth-icon-btn" data-report="auth" title="Auth diagnostics" aria-label="Auth diagnostics">U</button>
        </div>
      </div>
      <div id="notesTooltip" class="tooltip-panel" hidden>
        <strong>Quick Notes</strong><br />
        Base endpoint: `https://jira.mdthink.maryland.gov/rest/api/2/search`<br />
        Summary output includes status transition metrics and path columns.<br />
        Credentials come from env vars: `JIRA_USERNAME/JIRA_PASSWORD` or `JIRA_EMAIL/JIRA_API_TOKEN`.
      </div>
    </div>

    <section id="csmsDashboardSection" class="app-section report-scope-csms">
      <div class="card">
        <div class="hero-head">
          <div>
            <h1 class="section-title">CSMS Support Portfolio</h1>
            <p id="csmsSubtitle" class="section-subtitle">Executive Incident Summary</p>
            <p id="csmsElapsed" class="elapsed-note">Elapsed time sentence appears after a report run.</p>
          </div>
          <div class="status-pill">Operational Status: <strong id="csmsOpStatus">Monitoring</strong></div>
        </div>
        <div class="row kpi-grid" id="csmsKpis"></div>
      </div>
      <div class="two-col" style="margin-top:18px;">
        <div class="card">
          <h2>Top 3 Trends & Critical Issues</h2>
          <pre id="csmsNarratives" class="large-pre">Run the report to generate narratives.</pre>
        </div>
        <div class="card">
          <div id="csmsHealth" class="health-panel"></div>
          <h3 style="margin:16px 0 8px;">Stuck Ticket Deep-Dive</h3>
          <pre id="csmsStuck">No ticket identified yet.</pre>
        </div>
      </div>
      <div class="card" style="margin-top:18px;">
        <h2>Charts</h2>
        <div class="csms-chart-wrap daily">
          <canvas id="dailyTrendChart"></canvas>
        </div>
        <div class="csms-chart-grid">
          <div class="csms-chart-wrap small">
            <canvas id="topCategoryChart"></canvas>
          </div>
          <div class="csms-chart-wrap small">
            <canvas id="statusChart"></canvas>
          </div>
        </div>
      </div>
    </section>

    <section id="legacyDashboardSection" class="app-section report-scope-legacy" hidden>
      <div class="card" style="margin-top:18px;">
        <h2>Trends</h2>
        <h3 style="margin:0 0 8px;">Created / Updated / Resolved Trends</h3>
        <div class="legacy-chart-wrap daily">
          <canvas id="legacyDailyChart"></canvas>
        </div>
        <h3 style="margin:14px 0 8px;">Current Status Distribution</h3>
        <div class="legacy-chart-wrap status">
          <canvas id="legacyStatusChart"></canvas>
        </div>
      </div>
      <div class="two-col" style="margin-top:18px;">
        <div class="card">
          <h2>Insights</h2>
          <pre id="legacyInsights">Run a legacy dashboard refresh.</pre>
        </div>
        <div class="card">
          <h2>Status Summary</h2>
          <pre id="legacyStatusSummary">Run a legacy dashboard refresh.</pre>
        </div>
      </div>
    </section>

    <section id="settingsSection" class="app-section">
      <div id="legacySettingsCard" class="card collapse-card report-scope-legacy" hidden>
        <button type="button" class="muted-btn collapse-toggle" data-collapse-target="legacySettings" aria-expanded="false" aria-controls="legacySettings">Show Report Variables & Settings</button>
        <div id="legacySettings" class="collapse-body" hidden>
          <h2>Report Settings</h2>
          <form id="exportForm" style="margin-top:14px;">
            <div class="field wide">
              <label>Jira Search Endpoint</label>
              <input name="base_url" value="https://jira.mdthink.maryland.gov/rest/api/2/search" placeholder="https://your-jira-domain/rest/api/2/search" />
            </div>
            <div class="field wide"><label>Projects (comma separated)</label><input name="projects" placeholder="CSSD,ABC (project keys, not names)" /></div>
            <div class="field"><label>Issue Types</label><input name="issue_types" placeholder="Bug, Task, Story" /></div>
            <div class="field"><label>Current Statuses</label><input name="statuses" placeholder="Open, In Progress, Closed" /></div>
            <div class="field"><label>Assignees</label><input name="assignees" placeholder="username1,username2" /></div>
            <div class="field"><label>Labels</label><input name="labels" placeholder="defect, uat" /></div>
            <div class="field"><label>Date Field</label><select name="date_field"><option value="created">created</option><option value="updated">updated</option><option value="resolutiondate">resolutiondate</option></select></div>
            <div class="field"><label>Start Date/Time</label><input type="datetime-local" name="start_dt" /></div>
            <div class="field"><label>End Date/Time</label><input type="datetime-local" name="end_dt" /></div>
            <div class="field"><label>Time Block Start</label><input type="time" name="time_block_start" /></div>
            <div class="field"><label>Time Block End</label><input type="time" name="time_block_end" /></div>
            <div class="field"><label>Page Size</label><input type="number" name="page_size" value="50" min="1" max="100" /></div>
            <div class="field"><label>Max Issues (0 = all)</label><input type="number" name="max_issues" value="0" min="0" /></div>
            <div class="field full"><label>Extra JQL (AND appended)</label><textarea name="extra_jql" placeholder='Example: component = "Establishment" AND priority = Medium'></textarea></div>
            <div class="field full"><label>Custom JQL Override (optional)</label><textarea name="custom_jql" placeholder='Example: issuekey = CSSD-123'></textarea></div>
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
                <button type="button" class="muted-btn" id="legacyRefreshBtn">Refresh Dashboard</button>
              </div>
            </div>
          </form>
        </div>
      </div>

      <div id="csmsSettingsCard" class="card collapse-card report-scope-csms" style="margin-top:18px;">
        <button type="button" class="muted-btn collapse-toggle" data-collapse-target="csmsSettings" aria-expanded="false" aria-controls="csmsSettings">Show CSMS Variables & Settings</button>
        <div id="csmsSettings" class="collapse-body" hidden>
          <h2>CSMS Executive Incident Summary Parameters</h2>
          <form id="csmsForm">
            <div class="field wide"><label>Jira Search Endpoint</label><input name="base_url" value="https://jira.mdthink.maryland.gov/rest/api/2/search" /></div>
            <div class="field wide"><label>Projects (comma separated)</label><input name="projects" value="CSSD,CSD,CDF" /></div>
            <div class="field"><label>Report Generation Date/Time</label><input type="datetime-local" name="report_datetime" /></div>
            <div class="field"><label>Last Report Timestamp</label><input type="datetime-local" name="last_report_timestamp" /></div>
            <div class="field"><label>Period Length (days)</label><input type="number" name="period_length" value="15" min="1" /></div>
            <div class="field"><label>Issue Types</label><input name="issue_types" placeholder="Bug, Task" /></div>
            <div class="field"><label>Status Filters</label><input name="statuses" placeholder="New, Open, Closed" /></div>
            <div class="field"><label>Components</label><input name="components" placeholder="Financial Management, Case Management" /></div>
            <div class="field"><label>Page Size</label><input type="number" name="page_size" value="100" min="1" max="500" /></div>
            <div class="field"><label>Max Issues (0 = all)</label><input type="number" name="max_issues" value="0" min="0" /></div>
            <div class="field"><label>Process Alignment %</label><input type="number" name="process_alignment_pct" value="60" min="0" max="100" /></div>
            <div class="field full">
              <div class="row">
                <label class="check"><input type="checkbox" name="verify_ssl" checked /> Verify SSL</label>
                <button type="submit">Refresh from Jira API</button>
                <button type="button" class="muted-btn" id="csmsExportCsv">Export CSV</button>
                <button type="button" class="muted-btn" id="csmsExportExcel">Export Excel</button>
                <button type="button" class="muted-btn" id="csmsExportPdf">Export PDF</button>
              </div>
            </div>
          </form>
        </div>
      </div>
    </section>

    <section id="dataExportsSection" class="app-section report-scope-legacy" hidden>
      <div class="card">
        <h2>Data Exports</h2>
        <pre id="result">Ready.</pre>
        <div class="downloads" id="downloads"></div>
      </div>
    </section>

    <section id="authSection" class="app-section report-scope-auth" hidden>
      <div class="card">
        <div class="hero-head">
          <div>
            <h2>Jira API Auth Diagnostics</h2>
            <p class="small">Validate app credentials, current Jira user identity, and project visibility.</p>
          </div>
          <button type="button" class="muted-btn" id="authCheckBtn">Run Auth Check</button>
        </div>
        <div class="row">
          <label class="check"><input type="checkbox" id="authVerifySsl" checked /> Verify SSL</label>
        </div>
        <pre id="authResult" style="margin-top:12px;">Run auth check to inspect current API identity and visible projects.</pre>
      </div>
    </section>
  </div>

<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
  const next = chosen === "dark" ? "light" : "dark";
  toggle.textContent = chosen === "dark" ? "☀" : "☾";
  toggle.setAttribute("title", `Switch to ${next}`);
  toggle.setAttribute("aria-label", `Switch to ${next}`);
}

document.getElementById("themeToggle").addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "dark";
  applyTheme(current === "dark" ? "light" : "dark");
});

applyTheme(localStorage.getItem("theme") || "dark");

function setCollapseState(toggle, expanded) {
  const targetId = toggle.getAttribute("data-collapse-target");
  const body = document.getElementById(targetId);
  if (!body) return;
  toggle.setAttribute("aria-expanded", expanded ? "true" : "false");
  toggle.textContent = expanded ? "Hide Variables & Settings" : "Show Variables & Settings";
  body.hidden = !expanded;
}

document.querySelectorAll(".collapse-toggle").forEach((toggle) => {
  setCollapseState(toggle, false);
  toggle.addEventListener("click", () => {
    const expanded = toggle.getAttribute("aria-expanded") === "true";
    setCollapseState(toggle, !expanded);
  });
});

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
  await refreshLegacyDashboard(payload);
});

let csmsCharts = { daily: null, status: null, top: null };
let legacyCharts = { status: null, daily: null };
let latestCsmsPayload = null;

function setActiveReport(report) {
  document.querySelectorAll(".report-scope-csms").forEach((el) => { el.hidden = report !== "csms"; });
  document.querySelectorAll(".report-scope-legacy").forEach((el) => { el.hidden = report !== "legacy"; });
  document.querySelectorAll(".report-scope-auth").forEach((el) => { el.hidden = report !== "auth"; });
  document.querySelectorAll(".report-tab").forEach((btn) => {
    btn.classList.toggle("active", btn.getAttribute("data-report") === report);
  });
  window.scrollTo({ top: 0, behavior: "smooth" });
}

document.querySelectorAll(".report-tab").forEach((btn) => {
  btn.addEventListener("click", () => setActiveReport(btn.getAttribute("data-report")));
});

const notesInfoBtn = document.getElementById("notesInfoBtn");
const notesTooltip = document.getElementById("notesTooltip");
notesInfoBtn.addEventListener("click", () => {
  const expanded = notesInfoBtn.getAttribute("aria-expanded") === "true";
  notesInfoBtn.setAttribute("aria-expanded", expanded ? "false" : "true");
  notesTooltip.hidden = expanded;
});
document.addEventListener("click", (evt) => {
  if (!notesTooltip.hidden && !notesTooltip.contains(evt.target) && evt.target !== notesInfoBtn) {
    notesTooltip.hidden = true;
    notesInfoBtn.setAttribute("aria-expanded", "false");
  }
});
document.addEventListener("keydown", (evt) => {
  if (evt.key === "Escape") {
    notesTooltip.hidden = true;
    notesInfoBtn.setAttribute("aria-expanded", "false");
  }
});

function csmsFormToObject(form) {
  const obj = {};
  const fd = new FormData(form);
  for (const [key, value] of fd.entries()) {
    obj[key] = value;
  }
  obj.verify_ssl = form.verify_ssl.checked;
  return obj;
}

function trendClass(val) {
  return val >= 0 ? "trend-pos" : "trend-neg";
}

function trendTone(metricKey, trendValue) {
  const upIsGood = {
    backlog: false,
    new_created: false,
    resolved: true,
  };
  const isUp = Number(trendValue) >= 0;
  const good = upIsGood[metricKey] ? isUp : !isUp;
  return good ? "trend-pos" : "trend-neg";
}

function renderCsmsKpis(kpis) {
  const longest = kpis.longest_open || {};
  const cards = [
    ["backlog", "Backlog Tickets", kpis.backlog.period2, kpis.backlog.trend],
    ["new_created", "New Created", kpis.new_created.period2, kpis.new_created.trend],
    ["resolved", "Resolved Tickets", kpis.resolved.period2, kpis.resolved.trend],
    ["longest_open", "Longest Open", `${longest.age_days || 0} days`, null, `${longest.issue_key || "N/A"}`],
  ];
  const html = cards.map(([metricKey, title, value, trend, subtext]) => `
    <div class="kpi-card">
      <div class="kpi-label">${title}</div>
      <div class="kpi-number">${value}</div>
      ${subtext ? `<div class="small">${subtext}</div>` : ""}
      ${trend === null ? "" : `<div class="kpi-trend ${trendTone(metricKey, trend)}">${Math.abs(trend).toFixed(2)}% ${trend >= 0 ? "↗" : "↘"}</div>`}
    </div>
  `).join("");
  document.getElementById("csmsKpis").innerHTML = html;
}

function renderCsmsNarratives(narratives) {
  const text = [
    narratives.ticket_trends || "",
    narratives.pipeline_flow || "",
    narratives.csd_backlog || "",
  ].join("\\n\\n");
  document.getElementById("csmsNarratives").textContent = text;
}

function renderCsmsHealth(data) {
  const pct = Number(data.process_alignment_pct || 60);
  const gap = data.process_gap_identified || "No";
  document.getElementById("csmsHealth").innerHTML = `
    <h3>Operational Health & Readiness</h3>
    <div>Process Alignment: CSMS ${pct.toFixed(1)}% Complete</div>
    <div class="progress-wrap"><div class="progress-bar" style="width:${Math.max(0, Math.min(100, pct))}%"></div></div>
    <div>Process Gap Identified: ${gap}</div>
  `;
}

function destroyChart(instance) {
  if (instance) instance.destroy();
}

function renderLegacyKpis(kpis) {
  const container = document.getElementById("legacyKpis");
  if (!container) return;
  const cards = [
    ["Issue Count", kpis.issue_count || 0],
    ["Status Transitions", kpis.transition_count || 0],
    ["Comment Volume", kpis.comment_count || 0],
    ["Date Window Days", kpis.date_window_days || 0],
  ];
  container.innerHTML = cards.map(([title, value]) => `
    <div class="kpi-card">
      <div>${title}</div>
      <div class="kpi-number">${value}</div>
    </div>
  `).join("");
}

function renderLegacyCharts(charts) {
  const dailyCtx = document.getElementById("legacyDailyChart").getContext("2d");
  const statusCtx = document.getElementById("legacyStatusChart").getContext("2d");
  destroyChart(legacyCharts.status);
  destroyChart(legacyCharts.daily);

  const statusEntries = Object.entries(charts.status_distribution || {});
  legacyCharts.status = new Chart(statusCtx, {
    type: "doughnut",
    data: {
      labels: statusEntries.map(([k]) => k),
      datasets: [{ data: statusEntries.map(([, v]) => v) }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });

  legacyCharts.daily = new Chart(dailyCtx, {
    type: "bar",
    data: {
      labels: (charts.created_daily || {}).dates || [],
      datasets: [
        { label: "Created", data: (charts.created_daily || {}).created_counts || [] },
        { label: "Updated", data: (charts.created_daily || {}).updated_counts || [] },
        { label: "Resolved", data: (charts.created_daily || {}).resolved_counts || [] },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

function renderLegacyStatusSummary(charts) {
  const statusEntries = Object.entries(charts.status_distribution || {});
  if (!statusEntries.length) {
    document.getElementById("legacyStatusSummary").textContent = "No status data available.";
    return;
  }
  const total = statusEntries.reduce((acc, [, v]) => acc + Number(v || 0), 0) || 1;
  const lines = statusEntries
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .map(([name, count]) => {
      const pct = ((Number(count) / total) * 100).toFixed(1);
      return `${name}: ${count} (${pct}%)`;
    });
  document.getElementById("legacyStatusSummary").textContent = lines.join("\\n");
}

async function refreshLegacyDashboard(payload) {
  const res = await fetch("/run-legacy-dashboard", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok) {
    document.getElementById("legacyInsights").textContent = JSON.stringify(data, null, 2);
    return;
  }
  renderLegacyKpis(data.kpis || {});
  renderLegacyCharts(data.charts || {});
  renderLegacyStatusSummary(data.charts || {});
  const warningLines = data.warnings || [];
  const insightLines = data.insights || [];
  document.getElementById("legacyInsights").textContent = [...warningLines, ...insightLines].join("\\n");
}

function renderCsmsCharts(charts) {
  const dailyCtx = document.getElementById("dailyTrendChart").getContext("2d");
  const statusCtx = document.getElementById("statusChart").getContext("2d");
  const topCtx = document.getElementById("topCategoryChart").getContext("2d");

  destroyChart(csmsCharts.daily);
  destroyChart(csmsCharts.status);
  destroyChart(csmsCharts.top);

  csmsCharts.daily = new Chart(dailyCtx, {
    type: "line",
    data: {
      labels: charts.daily_by_component.dates || [],
      datasets: (charts.daily_by_component.series || []).map((s) => ({ label: s.name, data: s.data, tension: 0.2 }))
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { title: { display: true, text: "Daily Ticket Trend by Component" } },
    }
  });

  const statusEntries = Object.entries(charts.status_distribution || {});
  csmsCharts.status = new Chart(statusCtx, {
    type: "doughnut",
    data: {
      labels: statusEntries.map(([k]) => k),
      datasets: [{ data: statusEntries.map(([,v]) => v) }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: { display: false },
        legend: { display: false },
      },
    }
  });

  const topRows = charts.top_categories || [];
  csmsCharts.top = new Chart(topCtx, {
    type: "bar",
    data: {
      labels: topRows.map((r) => r.name),
      datasets: [{ label: "Count", data: topRows.map((r) => r.count) }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { title: { display: true, text: "Top 5 Issue Type or Component Trends" } },
    }
  });
}

document.getElementById("csmsForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const payload = csmsFormToObject(e.target);
  document.getElementById("csmsNarratives").textContent = "Running...";
  const res = await fetch("/run-csms-exec-summary", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await res.json();
  if (!res.ok) {
    document.getElementById("csmsNarratives").textContent = JSON.stringify(data, null, 2);
    return;
  }
  latestCsmsPayload = data;
  document.getElementById("csmsSubtitle").textContent = `Executive Incident Summary | ${data.periods.period2.label}`;
  document.getElementById("csmsElapsed").textContent = data.elapsed_time_sentence || "Elapsed time sentence not available.";
  renderCsmsKpis(data.kpis);
  renderCsmsNarratives(data.narratives);
  renderCsmsHealth(data.operational_health || {});
  document.getElementById("csmsStuck").textContent = JSON.stringify(data.kpis.longest_open || {}, null, 2);
  renderCsmsCharts(data.charts || {});
});

function openCsmsExport(kind) {
  if (!latestCsmsPayload || !latestCsmsPayload.exports) return;
  const url = latestCsmsPayload.exports[kind];
  if (url) window.open(url, "_blank");
}

document.getElementById("csmsExportCsv").addEventListener("click", () => openCsmsExport("csv_zip"));
document.getElementById("csmsExportExcel").addEventListener("click", () => openCsmsExport("excel"));
document.getElementById("csmsExportPdf").addEventListener("click", () => openCsmsExport("pdf"));
document.getElementById("legacyRefreshBtn").addEventListener("click", async () => {
  const payload = formToObject(document.getElementById("exportForm"));
  document.getElementById("legacyInsights").textContent = "Refreshing dashboard...";
  await refreshLegacyDashboard(payload);
});

document.getElementById("authCheckBtn").addEventListener("click", async () => {
  const base = document.querySelector('#csmsForm input[name="base_url"]')?.value
    || document.querySelector('#exportForm input[name="base_url"]')?.value
    || "https://jira.mdthink.maryland.gov/rest/api/2/search";
  const verifySsl = document.getElementById("authVerifySsl").checked;
  document.getElementById("authResult").textContent = "Running auth check...";
  const res = await fetch("/auth-status", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ base_url: base, verify_ssl: verifySsl }),
  });
  const data = await res.json();
  document.getElementById("authResult").textContent = JSON.stringify(data, null, 2);
});

renderCsmsKpis({
  backlog: { period2: "--", trend: 0 },
  new_created: { period2: "--", trend: 0 },
  resolved: { period2: "--", trend: 0 },
  longest_open: { age_days: "--", issue_key: "" },
});
document.getElementById("csmsElapsed").textContent = "Provide Last Report Timestamp and run CSMS refresh to compute elapsed time.";
renderCsmsHealth({ process_alignment_pct: 60, process_gap_identified: "Pending run" });
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


def calculate_percent_trend(period1: int, period2: int) -> float:
    if period1 == 0:
        if period2 == 0:
            return 0.0
        return 100.0
    return ((period2 - period1) / period1) * 100.0


def get_period_windows(report_date: Optional[str], period_length: int) -> Dict[str, Dict[str, str]]:
    period_length = max(1, int(period_length or 15))
    if report_date:
        try:
            parsed = datetime.fromisoformat(report_date.replace("Z", "+00:00"))
        except ValueError:
            parsed = datetime.strptime(report_date[:16], "%Y-%m-%dT%H:%M")
    else:
        parsed = datetime.now().astimezone()

    report_day = parsed.date()
    p2_start_day = report_day - timedelta(days=period_length - 1)
    p1_end_day = p2_start_day - timedelta(days=1)
    p1_start_day = p1_end_day - timedelta(days=period_length - 1)

    return {
        "period1": {
            "start": f"{p1_start_day.isoformat()} 00:00",
            "end": f"{p1_end_day.isoformat()} 23:59",
            "label": f"{p1_start_day.isoformat()} to {p1_end_day.isoformat()}",
        },
        "period2": {
            "start": f"{p2_start_day.isoformat()} 00:00",
            "end": f"{report_day.isoformat()} 23:59",
            "label": f"{p2_start_day.isoformat()} to {report_day.isoformat()}",
        },
    }


def parse_report_datetime_value(value: Optional[str]) -> datetime:
    if value:
        try:
            dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            dt = datetime.strptime(value[:16], "%Y-%m-%dT%H:%M")
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
        return dt
    return datetime.now().astimezone()


def build_elapsed_time_sentence(last_report_ts: Optional[str], report_dt: datetime) -> str:
    if not last_report_ts:
        return ""
    last_dt = parse_report_datetime_value(last_report_ts)
    delta = report_dt - last_dt
    if delta.total_seconds() < 0:
        return "Last report timestamp is after the report generation time."
    days = delta.days
    hours = delta.seconds // 3600
    minutes = (delta.seconds % 3600) // 60
    return (
        f"The time elapsed between the last report timestamp ({last_dt.strftime('%Y-%m-%d %H:%M')}) "
        f"and report generation time is: {days} days, {hours} hours, and {minutes} minutes."
    )


def build_jql_for_period(projects: List[str], start_date: str, end_date: str, filters: Dict[str, Any]) -> str:
    clauses: List[str] = []
    for clause in [
        list_clause("project", projects),
        list_clause("issuetype", parse_csv_list(filters.get("issue_types"))),
        list_clause("status", parse_csv_list(filters.get("statuses"))),
        list_clause("component", parse_csv_list(filters.get("components"))),
    ]:
        if clause:
            clauses.append(clause)
    clauses.append(f'created >= "{start_date}"')
    clauses.append(f'created <= "{end_date}"')
    extra_jql = (filters.get("extra_jql") or "").strip()
    if extra_jql:
        clauses.append(f"({extra_jql})")
    return " AND ".join(clauses) + " ORDER BY created DESC"


def parse_jira_datetime(value: str) -> Optional[datetime]:
    if not value:
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except Exception:
        return None


def fetch_jira_issues(
    base_url: str,
    jql: str,
    page_size: int,
    max_issues: int,
    verify_ssl: bool,
    include_changelog: bool = True,
) -> List[Dict[str, Any]]:
    return fetch_issues(base_url, jql, page_size, max_issues, verify_ssl, include_changelog=include_changelog)


def get_issue_project_key(issue: Dict[str, Any]) -> str:
    fields = issue.get("fields") or {}
    project = fields.get("project") or {}
    return (project.get("key") or "").upper()


def get_issue_status(issue: Dict[str, Any]) -> str:
    fields = issue.get("fields") or {}
    status = fields.get("status") or {}
    return (status.get("name") or "").strip()


def get_project_final_status(project_key: str, project_rules: Dict[str, str]) -> str:
    key = (project_key or "").upper()
    if key in project_rules:
        return project_rules[key]
    return "Closed"


def group_by_component(issues: List[Dict[str, Any]]) -> Dict[str, int]:
    counts: Counter = Counter()
    for issue in issues:
        fields = issue.get("fields") or {}
        components = fields.get("components") or []
        if not components:
            counts["Unspecified"] += 1
            continue
        for component in components:
            name = (component.get("name") or "Unspecified").strip() or "Unspecified"
            counts[name] += 1
    return dict(counts)


def group_by_status(issues: List[Dict[str, Any]]) -> Dict[str, int]:
    counts: Counter = Counter()
    for issue in issues:
        name = get_issue_status(issue) or "Unspecified"
        counts[name] += 1
    return dict(counts)


def get_backlog_count(issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    backlog = 0
    for issue in issues:
        status = get_issue_status(issue)
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules)
        if status.lower() != final_status.lower():
            backlog += 1
    return backlog


def get_finalized_count(issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    count = 0
    for issue in issues:
        status = get_issue_status(issue)
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules)
        if status.lower() == final_status.lower():
            count += 1
    return count


def get_oldest_open_ticket(issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> Dict[str, Any]:
    oldest_issue: Optional[Dict[str, Any]] = None
    oldest_dt: Optional[datetime] = None
    now = datetime.now().astimezone()

    for issue in issues:
        fields = issue.get("fields") or {}
        project_key = get_issue_project_key(issue)
        status = get_issue_status(issue)
        final_status = get_project_final_status(project_key, project_rules)
        if status.lower() == final_status.lower():
            continue
        created_dt = parse_jira_datetime(fields.get("created") or "")
        if not created_dt:
            continue
        if oldest_dt is None or created_dt < oldest_dt:
            oldest_dt = created_dt
            oldest_issue = issue

    if not oldest_issue or not oldest_dt:
        return {}

    fields = oldest_issue.get("fields") or {}
    age_days = max(0, (now - oldest_dt).days)
    return {
        "issue_key": oldest_issue.get("key", ""),
        "created_date": fields.get("created", ""),
        "current_status": get_issue_status(oldest_issue),
        "age_days": age_days,
        "workflow_gap": f'Workflow gap between {get_issue_status(oldest_issue)} -> {get_project_final_status(get_issue_project_key(oldest_issue), project_rules)}',
    }


def get_component_daily_series(issues: List[Dict[str, Any]], top_n: int = 10) -> Dict[str, Any]:
    component_totals = Counter(group_by_component(issues))
    top_components = [name for name, _ in component_totals.most_common(top_n)]
    date_component: Dict[str, Counter] = defaultdict(Counter)

    for issue in issues:
        fields = issue.get("fields") or {}
        created_dt = parse_jira_datetime(fields.get("created") or "")
        if not created_dt:
            continue
        created_day = created_dt.date().isoformat()
        components = fields.get("components") or []
        comp_names = [(c.get("name") or "Unspecified").strip() or "Unspecified" for c in components] or ["Unspecified"]
        for name in comp_names:
            bucket = name if name in top_components else "Other"
            date_component[created_day][bucket] += 1

    dates = sorted(date_component.keys())
    categories = top_components + (["Other"] if any("Other" in counts for counts in date_component.values()) else [])
    series = []
    for category in categories:
        series.append({
            "name": category,
            "data": [date_component[d].get(category, 0) for d in dates],
        })
    return {"dates": dates, "series": series}


def get_top_issue_type_or_component(issues: List[Dict[str, Any]], top_n: int = 5) -> List[Dict[str, Any]]:
    counts: Counter = Counter()
    for issue in issues:
        fields = issue.get("fields") or {}
        components = fields.get("components") or []
        if components:
            for component in components:
                name = (component.get("name") or "Unspecified").strip() or "Unspecified"
                counts[name] += 1
        else:
            issue_type = ((fields.get("issuetype") or {}).get("name") or "Unspecified").strip()
            counts[issue_type or "Unspecified"] += 1
    total = sum(counts.values()) or 1
    top_rows = []
    for name, value in counts.most_common(top_n):
        pct = (value / total) * 100.0
        top_rows.append({"name": name, "count": value, "percent": round(pct, 2)})
    return top_rows


def get_component_concentration(top_components: List[Dict[str, Any]], total: int) -> float:
    if total <= 0:
        return 0.0
    return round((sum(row["count"] for row in top_components) / total) * 100.0, 2)


def generate_executive_narratives(metrics: Dict[str, Any]) -> Dict[str, str]:
    top_components = metrics.get("top_components", [])
    concentration = metrics.get("component_concentration", 0.0)
    pipeline_trend = metrics.get("resolved_trend", 0.0)
    csd_backlog_trend = metrics.get("csd_backlog_trend", 0.0)

    ticket_trends = (
        f"Ticket Trends: The Throughput Bottleneck Burden: With approximately {concentration:.1f}% of ticket demand "
        f"concentrated in top components ({', '.join(row['name'] for row in top_components[:3]) or 'N/A'}), "
        "the system is heavily driven by a few high-volume areas."
    )
    pipeline_flow = (
        f"Pipeline Flow: The Final-Stage Bottleneck Burden: Final-stage throughput changed by {pipeline_trend:+.2f}% "
        "between periods. If intake remains steady while closure slows, tickets will accumulate at final stages."
    )
    csd_backlog = (
        f"CSD Backlog: The Capacity Imbalance Burden: CSD backlog shifted by {csd_backlog_trend:+.2f}% "
        "between periods, indicating where capacity balancing or workflow intervention is needed."
    )
    return {
        "ticket_trends": ticket_trends,
        "pipeline_flow": pipeline_flow,
        "csd_backlog": csd_backlog,
    }


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


def fetch_issues(
    base_url: str,
    jql: str,
    page_size: int,
    max_issues: int,
    verify_ssl: bool,
    include_changelog: bool = True,
) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    start_at = 0
    auth = get_auth()
    session = requests.Session()

    while True:
        payload = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": page_size,
        }
        if include_changelog:
            payload["expand"] = "changelog"
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


def issue_to_raw_row(issue: Dict[str, Any]) -> Dict[str, Any]:
    fields = issue.get("fields") or {}
    components = "; ".join((c.get("name") or "") for c in (fields.get("components") or []))
    return {
        "Issue Key": issue.get("key", ""),
        "Summary": fields.get("summary", ""),
        "Issue Type": option_value(fields.get("issuetype")),
        "Priority": option_value(fields.get("priority")),
        "Severity": pick_severity(fields),
        "Project": option_value(fields.get("project")),
        "Component/s": components,
        "Subcomponent": option_value(fields.get("customfield_subcomponent") or fields.get("subcomponent")),
        "Created Date": fields.get("created", ""),
        "Updated Date": fields.get("updated", ""),
        "Current Status": option_value(fields.get("status")),
        "Resolution": option_value(fields.get("resolution")),
        "Assignee": flatten_user(fields.get("assignee")),
        "Reporter": flatten_user(fields.get("reporter")),
    }


def write_excel(path: Path, sheets: Dict[str, List[Dict[str, Any]]]) -> None:
    wb = Workbook()
    first = True
    for sheet_name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = sheet_name[:31]
        if not rows:
            ws.append(["No data"])
            continue
        headers = list(rows[0].keys())
        ws.append(headers)
        for row in rows:
            ws.append([row.get(h, "") for h in headers])
    wb.save(path)


def write_pdf_summary(path: Path, payload: Dict[str, Any]) -> None:
    c = canvas.Canvas(str(path), pagesize=letter)
    width, height = letter
    y = height - 50
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "CSMS Executive Incident Summary")
    y -= 24
    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Period 2: {payload['periods']['period2']['label']}")
    y -= 20
    for label, value in [
        ("Backlog Tickets", payload["kpis"]["backlog"]["period2"]),
        ("New Created", payload["kpis"]["new_created"]["period2"]),
        ("Resolved Tickets", payload["kpis"]["resolved"]["period2"]),
    ]:
        c.drawString(40, y, f"{label}: {value}")
        y -= 16
    y -= 8
    for key in ("ticket_trends", "pipeline_flow", "csd_backlog"):
        text = payload["narratives"].get(key, "")
        c.drawString(40, y, text[:120])
        y -= 16
        if y < 60:
            c.showPage()
            y = height - 50
    c.save()


def build_csms_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")
    projects = parse_csv_list(params.get("projects")) or ["CSSD", "CSD", "CDF"]
    period_length = int(params.get("period_length") or 15)
    report_dt = parse_report_datetime_value(params.get("report_datetime"))
    periods = get_period_windows(report_dt.isoformat(), period_length)
    page_size = int(params.get("page_size") or 100)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    elapsed_time_sentence = build_elapsed_time_sentence(params.get("last_report_timestamp"), report_dt)

    project_rules = {"CSSD": "Closed", "CSD": "Ready For Production Users"}
    for item in parse_csv_list(params.get("project_final_status_rules")):
        if ":" in item:
            k, v = item.split(":", 1)
            if k.strip() and v.strip():
                project_rules[k.strip().upper()] = v.strip()

    p1_jql = build_jql_for_period(projects, periods["period1"]["start"], periods["period1"]["end"], params)
    p2_jql = build_jql_for_period(projects, periods["period2"]["start"], periods["period2"]["end"], params)
    period1_issues = fetch_jira_issues(base_url, p1_jql, page_size, max_issues, verify_ssl)
    period2_issues = fetch_jira_issues(base_url, p2_jql, page_size, max_issues, verify_ssl)

    p1_backlog = get_backlog_count(period1_issues, project_rules)
    p2_backlog = get_backlog_count(period2_issues, project_rules)
    p1_resolved = get_finalized_count(period1_issues, project_rules)
    p2_resolved = get_finalized_count(period2_issues, project_rules)
    p1_new = len(period1_issues)
    p2_new = len(period2_issues)
    oldest_open = get_oldest_open_ticket(period2_issues, project_rules)

    p2_component_counts = group_by_component(period2_issues)
    top_components = [{"name": k, "count": v} for k, v in Counter(p2_component_counts).most_common(5)]
    total_volume = sum(p2_component_counts.values())
    component_concentration = get_component_concentration(top_components, total_volume)

    p1_csd = [issue for issue in period1_issues if get_issue_project_key(issue) == "CSD"]
    p2_csd = [issue for issue in period2_issues if get_issue_project_key(issue) == "CSD"]
    csd_backlog_trend = calculate_percent_trend(
        get_backlog_count(p1_csd, project_rules), get_backlog_count(p2_csd, project_rules)
    )

    metrics = {
        "top_components": top_components,
        "component_concentration": component_concentration,
        "resolved_trend": calculate_percent_trend(p1_resolved, p2_resolved),
        "csd_backlog_trend": csd_backlog_trend,
    }

    payload = {
        "periods": periods,
        "elapsed_time_sentence": elapsed_time_sentence,
        "jql": {"period1": p1_jql, "period2": p2_jql},
        "kpis": {
            "backlog": {"period1": p1_backlog, "period2": p2_backlog, "trend": calculate_percent_trend(p1_backlog, p2_backlog)},
            "new_created": {"period1": p1_new, "period2": p2_new, "trend": calculate_percent_trend(p1_new, p2_new)},
            "resolved": {"period1": p1_resolved, "period2": p2_resolved, "trend": calculate_percent_trend(p1_resolved, p2_resolved)},
            "longest_open": oldest_open,
        },
        "charts": {
            "daily_by_component": get_component_daily_series(period2_issues, top_n=10),
            "status_distribution": group_by_status(period2_issues),
            "top_categories": get_top_issue_type_or_component(period2_issues, top_n=5),
        },
        "narratives": generate_executive_narratives(metrics),
        "operational_health": {
            "process_alignment_pct": float(params.get("process_alignment_pct") or 60),
            "process_gap_identified": "Yes" if metrics["resolved_trend"] < 0 else "No",
        },
        "raw_rows": {
            "period1": [issue_to_raw_row(issue) for issue in period1_issues],
            "period2": [issue_to_raw_row(issue) for issue in period2_issues],
        },
    }
    return payload


def build_legacy_dashboard_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")

    page_size = int(params.get("page_size") or 50)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    include_comments = bool(params.get("include_comments", True))
    include_workflow_events = bool(params.get("include_workflow_events", False))
    block_start = parse_time_value(params.get("time_block_start"))
    block_end = parse_time_value(params.get("time_block_end"))
    jql = build_jql(params)
    warnings: List[str] = []
    try:
        issues = fetch_jira_issues(
            base_url, jql, page_size, max_issues, verify_ssl, include_changelog=True
        )
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            issues = fetch_jira_issues(
                base_url, jql, page_size, max_issues, verify_ssl, include_changelog=False
            )
            warnings.append(
                "Jira returned 500 with changelog expansion; loaded legacy dashboard without changelog details."
            )
        else:
            raise

    activity_rows: List[Dict[str, Any]] = []
    summary_rows: List[Dict[str, Any]] = []
    for issue in issues:
        issue_activity = make_activity_rows(issue, include_workflow_events, block_start, block_end)
        activity_rows.extend(issue_activity)
        summary_rows.append(build_summary_row(issue, issue_activity, include_comments))

    created_daily: Counter = Counter()
    updated_daily: Counter = Counter()
    resolved_daily: Counter = Counter()
    issue_type_counts: Counter = Counter()
    for issue in issues:
        fields = issue.get("fields") or {}
        created_dt = parse_jira_datetime(fields.get("created") or "")
        if created_dt:
            created_daily[created_dt.date().isoformat()] += 1
        updated_dt = parse_jira_datetime(fields.get("updated") or "")
        if updated_dt:
            updated_daily[updated_dt.date().isoformat()] += 1
        resolved_dt = parse_jira_datetime(fields.get("resolutiondate") or "")
        if resolved_dt:
            resolved_daily[resolved_dt.date().isoformat()] += 1
        issue_type_counts[option_value(fields.get("issuetype")) or "Unspecified"] += 1

    total_comments = 0
    total_transitions = 0
    for row in summary_rows:
        total_comments += int(row.get("Comment Count") or 0)
        total_transitions += int(row.get("Status Transition Count") or 0)

    status_counts = group_by_status(issues)
    top_components = [{"name": k, "count": v} for k, v in Counter(group_by_component(issues)).most_common(5)]
    insights = [
        f"Workload distribution is led by {top_components[0]['name']} ({top_components[0]['count']} issues)." if top_components else "No component concentration detected.",
        f"Average transitions per issue: {(total_transitions / max(1, len(issues))):.2f}.",
        f"Most common current status: {max(status_counts, key=status_counts.get) if status_counts else 'N/A'}.",
    ]

    all_days = sorted(set(created_daily.keys()) | set(updated_daily.keys()) | set(resolved_daily.keys()))
    return {
        "jql": jql,
        "warnings": warnings,
        "kpis": {
            "issue_count": len(issues),
            "transition_count": total_transitions,
            "comment_count": total_comments,
            "date_window_days": len(created_daily),
        },
        "charts": {
            "status_distribution": status_counts,
            "issue_type_distribution": dict(issue_type_counts),
            "created_daily": {
                "dates": all_days,
                "created_counts": [created_daily.get(d, 0) for d in all_days],
                "updated_counts": [updated_daily.get(d, 0) for d in all_days],
                "resolved_counts": [resolved_daily.get(d, 0) for d in all_days],
            },
        },
        "insights": insights,
    }


def get_jira_root_url(search_url: str) -> str:
    url = (search_url or "").strip()
    if not url:
        return ""
    marker = "/rest/api/"
    idx = url.lower().find(marker)
    if idx > 0:
        return url[:idx]
    return url.rstrip("/")


def get_auth_diagnostics(params: Dict[str, Any]) -> Dict[str, Any]:
    base_url = (params.get("base_url") or "").strip()
    jira_root = get_jira_root_url(base_url)
    if not jira_root:
        raise ValueError("base_url is required")
    verify_ssl = bool(params.get("verify_ssl", True))
    auth = get_auth()
    auth_user = os.getenv("JIRA_EMAIL") or os.getenv("JIRA_USERNAME") or ""
    session = requests.Session()

    result: Dict[str, Any] = {
        "jira_root": jira_root,
        "using_env_auth": bool(auth_user and auth is not None),
        "auth_user_hint": auth_user,
        "myself": {},
        "visible_projects": [],
        "errors": [],
    }

    try:
        resp = session.get(f"{jira_root}/rest/api/2/myself", auth=auth, verify=verify_ssl, timeout=30)
        resp.raise_for_status()
        me = resp.json()
        result["myself"] = {
            "name": me.get("name") or me.get("key") or "",
            "displayName": me.get("displayName") or "",
            "emailAddress": me.get("emailAddress") or "",
            "active": bool(me.get("active", True)),
        }
    except Exception as exc:
        result["errors"].append(f"myself check failed: {exc}")

    try:
        resp = session.get(f"{jira_root}/rest/api/2/project", auth=auth, verify=verify_ssl, timeout=30)
        resp.raise_for_status()
        projects = resp.json() or []
        result["visible_projects"] = [
            {"key": p.get("key", ""), "name": p.get("name", "")}
            for p in projects
        ]
    except Exception as exc:
        result["errors"].append(f"project visibility check failed: {exc}")

    keys = {p["key"].upper() for p in result["visible_projects"] if p.get("key")}
    result["can_access_cssd"] = "CSSD" in keys
    result["can_access_csd"] = "CSD" in keys
    return result


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


@app.route("/run-csms-exec-summary", methods=["POST"])
def run_csms_exec_summary():
    try:
        params = request.get_json(force=True)
        payload = build_csms_payload(params)
        out_dir = Path(tempfile.mkdtemp(prefix="csms_exec_"))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_zip = out_dir / f"csms_exec_export_{stamp}.zip"
        excel_path = out_dir / f"csms_exec_export_{stamp}.xlsx"
        pdf_path = out_dir / f"csms_exec_summary_{stamp}.pdf"
        p1_csv = out_dir / f"raw_period1_{stamp}.csv"
        p2_csv = out_dir / f"raw_period2_{stamp}.csv"
        kpi_csv = out_dir / f"kpis_{stamp}.csv"

        p1_rows = payload["raw_rows"]["period1"]
        p2_rows = payload["raw_rows"]["period2"]
        write_csv(p1_csv, p1_rows, list(p1_rows[0].keys()) if p1_rows else ["Issue Key"])
        write_csv(p2_csv, p2_rows, list(p2_rows[0].keys()) if p2_rows else ["Issue Key"])
        kpi_rows = [
            {"Metric": "Backlog Tickets", "Period 1": payload["kpis"]["backlog"]["period1"], "Period 2": payload["kpis"]["backlog"]["period2"], "Trend %": payload["kpis"]["backlog"]["trend"]},
            {"Metric": "New Created", "Period 1": payload["kpis"]["new_created"]["period1"], "Period 2": payload["kpis"]["new_created"]["period2"], "Trend %": payload["kpis"]["new_created"]["trend"]},
            {"Metric": "Resolved Tickets", "Period 1": payload["kpis"]["resolved"]["period1"], "Period 2": payload["kpis"]["resolved"]["period2"], "Trend %": payload["kpis"]["resolved"]["trend"]},
        ]
        write_csv(kpi_csv, kpi_rows, ["Metric", "Period 1", "Period 2", "Trend %"])

        with zipfile.ZipFile(csv_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(p1_csv, p1_csv.name)
            zf.write(p2_csv, p2_csv.name)
            zf.write(kpi_csv, kpi_csv.name)

        write_excel(
            excel_path,
            {
                "Raw Period 1": p1_rows,
                "Raw Period 2": p2_rows,
                "KPI Metrics": kpi_rows,
            },
        )
        write_pdf_summary(pdf_path, payload)

        export_id = uuid.uuid4().hex
        CSMS_EXPORT_CACHE[export_id] = {
            "csv_zip": str(csv_zip),
            "excel": str(excel_path),
            "pdf": str(pdf_path),
        }
        payload["exports"] = {
            "csv_zip": f"/download-csms-export?export_id={export_id}&kind=csv_zip",
            "excel": f"/download-csms-export?export_id={export_id}&kind=excel",
            "pdf": f"/download-csms-export?export_id={export_id}&kind=pdf",
        }
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/run-legacy-dashboard", methods=["POST"])
def run_legacy_dashboard():
    try:
        payload = build_legacy_dashboard_payload(request.get_json(force=True))
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/auth-status", methods=["POST"])
def auth_status():
    try:
        payload = get_auth_diagnostics(request.get_json(force=True))
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/download-csms-export")
def download_csms_export():
    export_id = request.args.get("export_id", "")
    kind = request.args.get("kind", "")
    if not export_id or export_id not in CSMS_EXPORT_CACHE:
        return Response("Invalid export id", status=400)
    path = CSMS_EXPORT_CACHE[export_id].get(kind)
    if not path:
        return Response("Invalid export kind", status=400)
    return send_file(path, as_attachment=True)


@app.route("/download")
def download():
    path = request.args.get("path", "")
    if not path:
        return Response("Missing path", status=400)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=5001)
