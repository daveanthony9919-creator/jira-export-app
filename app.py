
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
from datetime import datetime, time, timedelta, timezone
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
TEAM_EXPORT_CACHE: Dict[str, Dict[str, str]] = {}
RESOLUTION_SLA_FIELD_CACHE: Dict[str, str] = {}

# Status substrings rolled up as "resolved" style outcomes for Team Posture resolved KPIs.
TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS: Tuple[str, ...] = (
    "resolved",
    "dev-completed",
    "closed",
    "ready for production users",
    "completed",
    "duplicate",
)

# Default Team tab roster when localStorage has no saved members.
# "username" is the JQL assignee literal passed through assignee in ("...") (see build_team_posture_jql / jql_quote).
# Use the same token your Jira accepts in Issue Navigator; swap for accountId or legacy username if needed.
TEAM_DEFAULT_MEMBERS: List[Dict[str, str]] = [
    {"id": "default_akanksha_mittal", "name": "Akanksha", "username": "akanksha.mittal@maryland.gov"},
    {"id": "default_anshuli_chaturvedi", "name": "Anshuli", "username": "anshuli.chaturvedi@maryland.gov"},
    {"id": "default_brahmendra_pathuri", "name": "Brahmendra", "username": "brahmendra.pathuri@maryland.gov"},
    {"id": "default_dustin_motley", "name": "Dustin", "username": "dustin.motley@maryland.gov"},
    {"id": "default_mathivathana_sakthipondu", "name": "Mathivathana", "username": "mathivathana.sakthipondu@maryland.gov"},
    {"id": "default_naga_neppalli", "name": "Naga", "username": "naga.neppalli@maryland.gov"},
    {"id": "default_nischay_modi", "name": "Nischay", "username": "nischay.modi@maryland.gov"},
    {"id": "default_pooja_parbadia", "name": "Pooja", "username": "pooja.parbadia@maryland.gov"},
    {"id": "default_sravanthi_kopalli", "name": "Sravanthi", "username": "sravanthirajendra.kopalli@maryland.gov"},
    {"id": "default_sulabh_kukreja", "name": "Sulabh", "username": "sulabh.kukreja@maryland.gov"},
    {"id": "default_swapnil_bante", "name": "Swapnil", "username": "swapnil.bante@maryland.gov"},
]

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>CSMS Operations Dashboard</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root {
      --bg: #f3f7fb;
      --panel: #ffffff;
      --panel-2: #f7f9fc;
      --text: #0f1f3a;
      --muted: #637799;
      --accent: #11b6ad;
      --accent-2: #2f7af8;
      --danger: #dc4b55;
      --border: #d9e2f0;
      --body-grad-start: #f7fafe;
      --body-grad-end: #eef4fb;
      --card-bg: #ffffff;
      --pre-bg: #f8fbff;
      --pre-text: #1d3358;
      --download-bg: #eef4ff;
      --download-text: #1b3360;
      --btn-bg-start: #1f6ff2;
      --btn-bg-end: #185bc7;
      --btn-border: #2456a8;
      --btn-text: #ffffff;
      --btn-shadow: rgba(20, 61, 131, 0.22);
      --btn-muted-bg-start: #ffffff;
      --btn-muted-bg-end: #f5f8fe;
      --btn-muted-border: #d6e0ef;
      --btn-muted-text: #1a335b;
      --side-bg: #082745;
      --side-title: #f2fbff;
      --side-btn-text: #d4e8ff;
      --side-btn-active-start: #0ea7a2;
      --side-btn-active-end: #0b8a86;
      --side-btn-active-border: #0b8e89;
      --action-btn-bg-start: #f9fbff;
      --action-btn-bg-end: #edf3fd;
      --action-btn-border: #d3e0f1;
      --action-btn-text: #1a335b;
      --heading-strong: #1c365f;
      --kpi-strong: #0f2343;
      --kpi-subtle: #08101f;
      --trend-good: #0da38e;
      --trend-bad: #dc4b55;
      --member-icon-start: #7fd1ff;
      --member-icon-end: #5db9ff;
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
      --side-bg: #0d1c39;
      --side-title: #d9e7ff;
      --side-btn-text: #c8daf8;
      --side-btn-active-start: #3a6fdf;
      --side-btn-active-end: #2a57b8;
      --side-btn-active-border: #3f6bc5;
      --action-btn-bg-start: rgba(255, 255, 255, 0.12);
      --action-btn-bg-end: rgba(255, 255, 255, 0.04);
      --action-btn-border: rgba(202, 226, 255, 0.24);
      --action-btn-text: #c9daf7;
      --heading-strong: #1c365f;
      --kpi-strong: #0f2343;
      --kpi-subtle: #08101f;
      --trend-good: #148a78;
      --trend-bad: #c23f4f;
      --member-icon-start: #8fd7ff;
      --member-icon-end: #73c6ff;
    }
    :root[data-theme="dark"] {
      --side-bg: #081a35;
      --side-title: #dce8ff;
      --side-btn-text: #c9daf7;
      --side-btn-active-start: #6f96ea;
      --side-btn-active-end: #4b72ca;
      --side-btn-active-border: #6c8fd9;
      --action-btn-bg-start: rgba(255, 255, 255, 0.12);
      --action-btn-bg-end: rgba(255, 255, 255, 0.04);
      --action-btn-border: rgba(202, 226, 255, 0.24);
      --action-btn-text: #c9daf7;
      --heading-strong: #cddcff;
      --kpi-strong: #f1f5ff;
      --kpi-subtle: #dce8ff;
      --trend-good: #4fd1c5;
      --trend-bad: #ff7098;
      --member-icon-start: #8fd7ff;
      --member-icon-end: #73c6ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, Arial, sans-serif;
      background: linear-gradient(180deg, var(--body-grad-start) 0%, var(--body-grad-end) 100%);
      color: var(--text);
    }
    .wrap {
      max-width: 1800px;
      margin: 0;
      padding: 14px 18px 48px 124px;
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
      border-radius: 8px;
      padding: 14px;
      box-shadow: 0 1px 4px rgba(16, 45, 92, 0.08);
    }
    .hero-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 10px;
    }
    h1 { margin: 0 0 10px; font-size: 44px; line-height: 1.05; }
    h2 {
      margin: 0 0 10px;
      font-size: 12px;
      color: var(--heading-strong);
      text-transform: uppercase;
      letter-spacing: .06em;
      font-weight: 700;
    }
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
      background: linear-gradient(180deg, var(--action-btn-bg-start), var(--action-btn-bg-end));
      border: 1px solid var(--action-btn-border);
      color: var(--action-btn-text);
      font-weight: 700;
      font-size: 12px;
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
      background: linear-gradient(180deg, var(--action-btn-bg-start), var(--action-btn-bg-end));
      color: var(--action-btn-text);
      border: 1px solid var(--action-btn-border);
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
      padding: 12px;
      border-radius: 8px;
      color: var(--pre-text);
      white-space: pre-wrap;
      word-break: break-word;
      border: 1px solid var(--border);
      margin: 0;
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
      position: fixed;
      left: 0;
      top: 0;
      bottom: 0;
      width: 108px;
      z-index: 30;
      margin: 0;
      border-radius: 0;
      border-left: 0;
      border-top: 0;
      border-bottom: 0;
      padding: 14px 10px;
      background: var(--side-bg);
    }
    .app-nav-shell {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
      gap: 12px;
      height: 100%;
    }
    .app-nav-title {
      font-size: 10px;
      font-weight: 800;
      letter-spacing: .06em;
      margin: 0;
      color: var(--side-title);
      text-align: center;
      white-space: normal;
      line-height: 1.25;
    }
    .app-nav-actions {
      display: flex;
      flex-direction: column;
      align-items: stretch;
      gap: 10px;
      width: 100%;
      margin-top: 6px;
    }
    .app-nav .muted-btn {
      width: 100%;
      min-height: 38px;
      padding: 8px 6px;
      font-size: 11px;
      border-radius: 8px;
      text-align: center;
      color: var(--side-btn-text);
      background: linear-gradient(180deg, rgba(255, 255, 255, 0.12), rgba(255, 255, 255, 0.04));
      border-color: rgba(202, 226, 255, 0.24);
      font-weight: 700;
      letter-spacing: .02em;
    }
    .app-nav .theme-toggle,
    .app-nav .auth-icon-btn {
      font-size: 14px;
      width: 100%;
      height: 38px;
      border-radius: 8px;
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
    .report-tab.active {
      background: linear-gradient(180deg, var(--side-btn-active-start), var(--side-btn-active-end));
      color: var(--btn-text);
      border-color: var(--side-btn-active-border);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.18);
    }
    .app-section {
      scroll-margin-top: 94px;
      margin-bottom: 12px;
    }
    .member-grid {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 10px;
    }
    .member-pill {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 14px;
      border: 1px solid var(--border);
      border-radius: 999px;
      background: linear-gradient(180deg, var(--panel), var(--panel-2));
      color: var(--text);
      cursor: pointer;
      box-shadow: 0 3px 8px rgba(24, 63, 130, 0.14);
      font-weight: 700;
    }
    .member-pill.active {
      border-color: var(--accent-2);
      box-shadow: 0 0 0 2px color-mix(in srgb, var(--accent-2) 30%, transparent), 0 6px 14px rgba(24, 63, 130, 0.2);
    }
    .member-icon {
      width: 26px;
      height: 26px;
      border-radius: 50%;
      background: linear-gradient(90deg, var(--member-icon-start), var(--member-icon-end));
      color: var(--kpi-subtle);
      font-weight: 800;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 12px;
    }
    .team-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(120px, 1fr));
      gap: 10px;
      margin-top: 12px;
    }
    .team-metric-card {
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel-2);
      padding: 12px;
    }
    .team-metric-card .label {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .04em;
    }
    .team-metric-card .value {
      font-size: 26px;
      font-weight: 800;
      margin-top: 6px;
      color: var(--accent);
    }
    .two-col {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
    }
    .kpi-grid .kpi-card {
      min-height: 104px;
      background: var(--panel-2);
      padding: 12px;
    }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(180px, 1fr));
      gap: 10px;
    }
    .tooltip-panel {
      margin-top: 10px;
      padding: 12px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel-2);
      color: var(--muted);
      font-size: 13px;
      max-width: 360px;
      margin-left: 96px;
    }
    .notes-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(260px, 1fr));
      gap: 12px;
      margin-top: 12px;
    }
    .notes-card {
      border: 1px solid var(--border);
      border-radius: 8px;
      background: var(--panel-2);
      padding: 12px;
    }
    .notes-card h3 {
      margin: 0 0 8px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: var(--heading-strong);
    }
    .notes-card p, .notes-card li {
      font-size: 13px;
      color: var(--text);
      line-height: 1.4;
    }
    .notes-card ul {
      margin: 0;
      padding-left: 18px;
    }
    .legacy-chart-wrap {
      width: 100%;
      position: relative;
      overflow: hidden;
      border-radius: 10px;
    }
    .legacy-chart-wrap.daily {
      height: 240px;
    }
    .legacy-chart-wrap.status {
      height: 240px;
    }
    .csms-chart-wrap {
      width: 100%;
      position: relative;
      overflow: hidden;
      border-radius: 8px;
      background: var(--panel);
      border: 1px solid var(--border);
      padding: 8px;
    }
    .csms-chart-wrap.daily {
      height: 170px;
    }
    .csms-chart-wrap.small {
      height: 155px;
    }
    .csms-chart-grid {
      display: grid;
      grid-template-columns: 1.25fr .75fr;
      gap: 8px;
      margin-top: 8px;
    }
    .report-scope-csms .section-title { font-size: 36px; font-weight: 800; letter-spacing: 0; }
    .report-scope-csms .section-subtitle { font-size: 14px; letter-spacing: .01em; text-transform: none; margin-top: 4px; }
    .report-scope-csms .status-pill { font-size: 11px; min-width: 180px; padding: 8px 10px; border-radius: 8px; }
    .report-scope-csms .kpi-grid { grid-template-columns: repeat(4, minmax(140px, 1fr)); gap: 10px; }
    .report-scope-csms .kpi-grid .kpi-card { min-height: 100px; padding: 12px; }
    .report-scope-csms .kpi-label { font-size: 11px; }
    .report-scope-csms .kpi-number { font-size: 34px; margin: 2px 0 4px; line-height: 1; }
    .report-scope-csms .kpi-trend { font-size: 12px; }
    .report-scope-csms .large-pre { min-height: 178px; font-size: 13px; line-height: 1.35; }
    .report-scope-csms .health-panel { min-height: 178px; font-size: 12px; }
    .report-scope-csms .health-panel h3 { font-size: 12px; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 8px; color: var(--heading-strong); }
    .report-scope-csms .elapsed-note { font-size: 11px; color: var(--muted); margin: 3px 0 0; }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }
    .tab-btn.active {
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      color: var(--kpi-subtle);
    }
    .kpi-card {
      background: var(--panel-2);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 14px;
      min-width: 180px;
      flex: 1;
    }
    .kpi-number { font-size: 36px; font-weight: 800; margin: 8px 0 6px; color: var(--kpi-strong); line-height: 1; }
    .kpi-label { color: var(--muted); font-size: 11px; letter-spacing: .06em; text-transform: uppercase; }
    .kpi-trend { font-size: 12px; font-weight: 700; }
    #teamMetricsGrid .team-metric-card,
    #csmsKpis .kpi-card,
    #legacyKpis .kpi-card {
      cursor: help;
    }
    .trend-pos { color: var(--trend-good); }
    .trend-neg { color: var(--trend-bad); }
    .section-title {
      font-size: 36px;
      margin: 0;
      letter-spacing: .01em;
    }
    .section-subtitle {
      font-size: 14px;
      letter-spacing: .02em;
      color: var(--muted);
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
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 11px;
      color: var(--muted);
      min-width: 170px;
      text-align: right;
    }
    .status-pill strong { color: var(--text); }
    .large-pre {
      min-height: 180px;
      font-size: 13px;
      line-height: 1.35;
    }
    .health-panel {
      min-height: 180px;
      font-size: 12px;
    }
    .health-panel h3 {
      margin: 0 0 10px;
      font-size: 12px;
      color: var(--heading-strong);
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
      .app-nav-title { font-size: 12px; }
      .wrap { padding-left: 16px; }
      .app-nav { position: static; width: 100%; height: auto; }
      .app-nav-actions { flex-direction: row; flex-wrap: wrap; justify-content: center; }
      .team-grid { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card app-nav">
      <div class="app-nav-shell">
        <h1 class="app-nav-title">CSMS Reporting</h1>
        <div class="app-nav-actions">
          <button type="button" class="muted-btn report-tab active" data-report="csms" title="Executive Report">Executive Report</button>
          <button type="button" class="muted-btn report-tab" data-report="team" title="Operations Team">Operations Team</button>
          <button type="button" class="muted-btn report-tab" data-report="legacy" title="Ticket trend">Ticket trend</button>
          <button type="button" class="muted-btn report-tab" data-report="notes" title="Notes & Guides">Notes</button>
          <button type="button" class="muted-btn report-tab auth-icon-btn" data-report="auth" title="Auth diagnostics" aria-label="Auth diagnostics">U</button>
          <button type="button" class="muted-btn theme-toggle" id="themeToggle" aria-label="Toggle theme" title="Switch theme">◐</button>
        </div>
      </div>
    </div>

    <section id="csmsDashboardSection" class="app-section report-scope-csms">
      <div class="card">
        <div class="hero-head">
          <div>
            <h1 class="section-title">CSMS Application</h1>
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
        <div class="csms-chart-wrap daily" title="Daily created and updated ticket counts for the current period, split by top components.">
          <canvas id="dailyTrendChart"></canvas>
        </div>
        <div class="csms-chart-grid">
          <div class="csms-chart-wrap small" title="Top five issue types or components by ticket count in the current period.">
            <canvas id="topCategoryChart"></canvas>
          </div>
          <div class="csms-chart-wrap small" title="How tickets in the current period snapshot are spread across current statuses.">
            <canvas id="statusChart"></canvas>
          </div>
        </div>
      </div>
    </section>

    <section id="teamPostureSection" class="app-section report-scope-team" hidden>
      <div class="card">
        <div class="hero-head">
          <div>
            <h2>Team Member Ticket Posture</h2>
            <p class="small">Click a team member to load stand-up and EOD posture metrics.</p>
            <p id="teamReportPeriod" class="small" style="margin-top:8px;color:var(--muted);">Report period: set Start and End in Team Posture settings.</p>
            <p id="teamRollupNote" class="small" style="margin-top:8px;color:var(--muted);"></p>
          </div>
        </div>
        <div id="teamRollupGrid" class="team-grid" style="margin-bottom:12px;">
          <div class="team-metric-card" title="Sum of Queue Backlog counts across team members with cached data: CSSD tickets in Under QA Analysis plus CSD tickets in New. Other projects are excluded per member."><div class="label">Team Queue Backlog</div><div id="teamRollupQueueBacklog" class="value">--</div></div>
          <div class="team-metric-card" title="Sum of In Progress counts across cached members: open CSSD tickets not in New or Under QA Analysis, and open CSD tickets not in New."><div class="label">Team In Progress</div><div id="teamRollupInProgress" class="value">--</div></div>
          <div class="team-metric-card" title="Sum across cached members: owned tickets in resolved-like status whose resolution time falls between Team Start and End."><div class="label">Team Resolved (Period)</div><div id="teamRollupResolvedPeriod" class="value">--</div></div>
        </div>
        <div id="teamMemberGrid" class="member-grid"></div>
        <div id="teamMetricsGrid" class="team-grid">
          <div class="team-metric-card" title="Open tickets tied to this member as assignee or CSD Assigned Developer when configured. Done means Closed on CSSD and Ready For Production Users on CSD."><div class="label">Assigned Open Tickets</div><div id="teamOpenCount" class="value">--</div></div>
          <div class="team-metric-card" title="CSSD: Under QA Analysis. CSD: New. Other projects: not counted. Uses ownership rules for this member."><div class="label">Queue Backlog</div><div id="teamQueueBacklogCount" class="value">--</div></div>
          <div class="team-metric-card" title="CSSD: open, not New, not Under QA Analysis. CSD: open and not New. Other projects: not counted."><div class="label">In Progress</div><div id="teamInProgressCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets you own where you authored a Jira status change in the changelog within the last eight hours from when this report ran. Requires changelog data from Jira."><div class="label">Worked Status (Last 8 Hours)</div><div id="teamWorkedStatusLast8hCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets in the date window whose status sounds reopened, where the member owns the ticket or authored at least one status change."><div class="label">Reopened Tickets</div><div id="teamReopenedCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets assigned to this member whose status looks finished, such as resolved, closed, completed, or duplicate."><div class="label">Resolved (Owned)</div><div id="teamResolvedOwnedCount" class="value">--</div></div>
          <div class="team-metric-card" title="Finished tickets owned by someone else where this member changed the status at least once."><div class="label">Resolved (Contributed)</div><div id="teamResolvedContributedCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets you own that have a Jira resolution time in the rolling last eight hours from when this report ran. Uses the same finished-status rules as Resolved (Owned)."><div class="label">Resolved (Last 8 Hours)</div><div id="teamResolvedLast8hCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets owned by someone else where this member still changed the status at least once."><div class="label">Worked On (Assigned to Others)</div><div id="teamWorkedOtherCount" class="value">--</div></div>
          <div class="team-metric-card" title="Tickets in this member scope that missed the 24-hour expectation from created time, using the Jira resolution SLA breached field when available, otherwise time to finish."><div class="label">SLA Breach Count</div><div id="teamSlaBreachCount" class="value">--</div></div>
          <div class="team-metric-card" title="Open tickets still within 24 hours from created but with under eight hours left before that window ends."><div class="label">Open Tickets &lt; 8h to SLA Breach</div><div id="teamSlaNearCount" class="value">--</div></div>
          <div class="team-metric-card" title="Among open tickets assigned to this member, the issue key that has waited the longest since creation."><div class="label">Oldest Open Ticket</div><div id="teamOldestTicket" class="value">--</div></div>
          <div class="team-metric-card" title="How many days that oldest open ticket has been waiting."><div class="label">Oldest Open Age (days)</div><div id="teamOldestAge" class="value">--</div></div>
        </div>
      </div>
      <div class="two-col" style="margin-top:18px;">
        <div class="card">
          <h2>Ticket Count by Status</h2>
          <pre id="teamStatusSummary">Select a member and refresh.</pre>
        </div>
        <div class="card">
          <h2>Oldest Open Detail</h2>
          <pre id="teamOldestDetail">Select a member and refresh.</pre>
        </div>
      </div>
      <div class="card" style="margin-top:18px;" title="How tickets in this member scope are split across Jira labels.">
        <h2>Ticket Labels</h2>
        <div class="legacy-chart-wrap daily">
          <canvas id="teamLabelsChart"></canvas>
        </div>
      </div>
    </section>

    <section id="legacyDashboardSection" class="app-section report-scope-legacy" hidden>
      <div class="card" style="margin-top:18px;">
        <h2>Trends</h2>
        <p id="legacyReportPeriod" class="small" style="margin:4px 0 12px;color:var(--muted);">Report period: set Start and End in Report Settings.</p>
        <div class="row kpi-grid" id="legacyKpis"></div>
        <h3 style="margin:0 0 8px;">Created / Updated / Resolved Trends</h3>
        <div class="legacy-chart-wrap daily" title="Tickets from your legacy query: how many were created, updated, or resolved on each day in the chart window.">
          <canvas id="legacyDailyChart"></canvas>
        </div>
        <h3 style="margin:14px 0 8px;">Current Status Distribution</h3>
        <div class="legacy-chart-wrap status" title="Share of tickets from your legacy query by their current Jira status.">
          <canvas id="legacyStatusChart"></canvas>
        </div>
      </div>
      <div class="two-col" style="margin-top:18px;">
        <div class="card" title="Short automated observations after you refresh the legacy dashboard.">
          <h2>Insights</h2>
          <pre id="legacyInsights">Run a legacy dashboard refresh.</pre>
        </div>
        <div class="card" title="Each status from your legacy query and how many tickets are in it.">
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
            <div class="field"><label>Last Report Backlog Tickets (optional)</label><input type="number" name="last_report_backlog_tickets" min="0" placeholder="e.g. 42" /></div>
            <div class="field"><label>Last Report New Created (optional)</label><input type="number" name="last_report_new_created" min="0" placeholder="e.g. 15" /></div>
            <div class="field"><label>Last Report Resolved Tickets (optional)</label><input type="number" name="last_report_resolved_tickets" min="0" placeholder="e.g. 18" /></div>
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

      <div id="teamSettingsCard" class="card collapse-card report-scope-team" style="margin-top:18px;" hidden>
        <button type="button" class="muted-btn collapse-toggle" data-collapse-target="teamSettings" aria-expanded="false" aria-controls="teamSettings">Show Team Posture Variables & Settings</button>
        <div id="teamSettings" class="collapse-body" hidden>
          <h2>Team Posture Settings</h2>
          <form id="teamPostureForm">
            <div class="field wide"><label>Jira Search Endpoint</label><input name="base_url" value="https://jira.mdthink.maryland.gov/rest/api/2/search" /></div>
            <div class="field wide"><label>Projects (comma separated)</label><input name="projects" value="CSSD,CSD,CDF" /></div>
            <div class="field"><label>Start Date/Time</label><input type="datetime-local" name="start_dt" /></div>
            <div class="field"><label>End Date/Time</label><input type="datetime-local" name="end_dt" /></div>
            <div class="field"><label>Issue Types</label><input name="issue_types" placeholder="Bug, Task" /></div>
            <div class="field"><label>CSD Assigned Developer Field Key</label><input name="csd_assigned_dev_field" value="customfield_14700" placeholder="customfield_12345" /></div>
            <div class="field"><label>Page Size</label><input type="number" name="page_size" value="50" min="1" max="100" /></div>
            <div class="field"><label>Max Issues (0 = all)</label><input type="number" name="max_issues" value="0" min="0" /></div>
            <div class="field full">
              <div class="row">
                <label class="check"><input type="checkbox" name="verify_ssl" checked /> Verify SSL</label>
              </div>
            </div>
            <div class="field full">
              <h3 style="margin:0 0 8px;">Team Members</h3>
              <div class="row">
                <input id="teamMemberNameInput" placeholder="Display name" />
                <input id="teamMemberUsernameInput" placeholder="Assignee username" />
                <button type="button" class="muted-btn" id="teamAddMemberBtn">Add Member</button>
                <button type="button" class="muted-btn" id="teamRemoveMemberBtn">Remove Selected</button>
              </div>
            </div>
            <div class="field full">
              <div class="row">
                <button type="button" class="muted-btn" id="teamRefreshBtn">Refresh All Member Metrics</button>
                <button type="button" class="muted-btn" id="teamExportCsvBtn">Download CSV</button>
                <button type="button" class="muted-btn" id="teamExportExcelBtn">Download Excel</button>
                <button type="button" class="muted-btn" id="teamExportAllBtn">Download Team CSV</button>
              </div>
            </div>
          </form>
        </div>
      </div>

      <div id="teamCsvPreviewCard" class="card report-scope-team" style="margin-top:18px;" hidden>
        <h2>CSV Preview</h2>
        <pre id="teamCsvPreview">Run Team Posture refresh to preview CSV rows.</pre>
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
      <div class="card" title="Runs a credential and project visibility check against your Jira endpoint.">
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

    <section id="notesSection" class="app-section report-scope-notes" hidden>
      <div class="card" title="Static help text for metrics, team posture, and how to use the dashboard.">
        <div class="hero-head">
          <div>
            <h2>Notes & Guides</h2>
            <p class="small">Reference content for dashboard metrics, charts, and daily usage.</p>
          </div>
        </div>
        <div class="notes-grid">
          <div class="notes-card">
            <h3>CSMS Dashboard Explainer</h3>
            <ul>
              <li><strong>Backlog Tickets:</strong> Open tickets based on project-specific final status rules.</li>
              <li><strong>New Created:</strong> Tickets created in the most recent comparison period.</li>
              <li><strong>Resolved Tickets:</strong> Tickets currently in final/resolved states.</li>
              <li><strong>Longest Open:</strong> Oldest active ticket with age, issue key, and workflow gap.</li>
            </ul>
          </div>
          <div class="notes-card">
            <h3>Team Posture Explainer</h3>
            <ul>
              <li><strong>Resolved (Owned):</strong> Resolved/final statuses for tickets owned by the member (assignee plus CSD &quot;Assigned Developer&quot; when configured).</li>
              <li><strong>Resolved (Contributed):</strong> Resolved/final statuses for tickets where the member appears as a status-transition author but is not the current owner.</li>
              <li><strong>Assigned Open:</strong> Current open workload assigned to selected member.</li>
              <li><strong>Reopened Tickets:</strong> Includes assigned tickets and worked-on tickets with reopen history.</li>
              <li><strong>Worked On (Assigned to Others):</strong> Status-change author matches member, but assignee differs.</li>
              <li><strong>Ticket Count by Status:</strong> Distribution for assigned tickets in current filter window.</li>
            </ul>
          </div>
          <div class="notes-card">
            <h3>How To Use</h3>
            <ul>
              <li>Set date range and project filters in Variables & Settings.</li>
              <li>Use Team icons to view per-member metrics.</li>
              <li>Click Refresh All Member Metrics for team-wide refresh.</li>
              <li>Use Download CSV/Excel for selected member, or Download Team CSV for all members.</li>
            </ul>
          </div>
          <div class="notes-card">
            <h3>Auth & Data Quality</h3>
            <ul>
              <li>Use Auth page to validate credentials and project visibility.</li>
              <li>If Jira returns 500 with changelog, fallback may reduce reopen/worked-on completeness.</li>
              <li>Username must match Jira account key/name for team calculations.</li>
              <li>Base endpoint: `https://jira.mdthink.maryland.gov/rest/api/2/search`</li>
            </ul>
          </div>
        </div>
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
    "--download-text": "#1b2d52",
    "--action-btn-bg-start": "#f9fbff",
    "--action-btn-bg-end": "#edf3fd",
    "--action-btn-border": "#d3e0f1",
    "--action-btn-text": "#1a335b",
    "--side-bg": "#0d1c39",
    "--side-title": "#dce8ff",
    "--side-btn-text": "#c9daf7",
    "--side-btn-active-start": "#3a6fdf",
    "--side-btn-active-end": "#2a57b8",
    "--side-btn-active-border": "#3f6bc5",
    "--heading-strong": "#1c365f",
    "--kpi-strong": "#0f2343",
    "--kpi-subtle": "#08101f",
    "--trend-good": "#148a78",
    "--trend-bad": "#c23f4f",
    "--member-icon-start": "#7fd1ff",
    "--member-icon-end": "#5db9ff"
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
    "--download-text": "#ffffff",
    "--action-btn-bg-start": "rgba(255, 255, 255, 0.12)",
    "--action-btn-bg-end": "rgba(255, 255, 255, 0.04)",
    "--action-btn-border": "rgba(202, 226, 255, 0.24)",
    "--action-btn-text": "#c9daf7",
    "--side-bg": "#081a35",
    "--side-title": "#dce8ff",
    "--side-btn-text": "#c9daf7",
    "--side-btn-active-start": "#6f96ea",
    "--side-btn-active-end": "#4b72ca",
    "--side-btn-active-border": "#6c8fd9",
    "--heading-strong": "#cddcff",
    "--kpi-strong": "#f1f5ff",
    "--kpi-subtle": "#dce8ff",
    "--trend-good": "#4fd1c5",
    "--trend-bad": "#ff7098",
    "--member-icon-start": "#8fd7ff",
    "--member-icon-end": "#73c6ff"
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
let teamCharts = { labels: null };
let latestCsmsPayload = null;
let teamMembers = [];
let activeTeamMemberId = null;
let latestTeamPosturePayload = null;
let teamPayloadByMemberId = {};

const DEFAULT_TEAM_MEMBERS = __DEFAULT_TEAM_ROSTER_JSON__;

function teamStorageKey() {
  return "team-posture-members-v1";
}

function loadTeamMembersFromStorage() {
  try {
    const raw = localStorage.getItem(teamStorageKey());
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed;
    }
  } catch (e) {}
  return JSON.parse(JSON.stringify(DEFAULT_TEAM_MEMBERS));
}

function saveTeamMembersToStorage(members) {
  localStorage.setItem(teamStorageKey(), JSON.stringify(members));
}

function initialsFor(name) {
  const parts = String(name || "").trim().split(/\\s+/).filter(Boolean);
  if (!parts.length) return "U";
  return parts.slice(0, 2).map((p) => p[0].toUpperCase()).join("");
}

function renderTeamMemberIcons() {
  const grid = document.getElementById("teamMemberGrid");
  if (!grid) return;
  grid.innerHTML = teamMembers.map((member) => `
    <button type="button" class="member-pill ${member.id === activeTeamMemberId ? "active" : ""}" data-member-id="${member.id}">
      <span class="member-icon">${initialsFor(member.name)}</span>
      <span>${member.name}</span>
    </button>
  `).join("");
  grid.querySelectorAll(".member-pill").forEach((btn) => {
    btn.addEventListener("click", () => {
      activeTeamMemberId = btn.getAttribute("data-member-id");
      renderTeamMemberIcons();
      const cached = teamPayloadByMemberId[activeTeamMemberId];
      if (cached) {
        latestTeamPosturePayload = cached;
        renderTeamPostureMetrics(cached);
      } else {
        refreshTeamPosture();
      }
    });
  });
}

function activeTeamMember() {
  return teamMembers.find((m) => m.id === activeTeamMemberId) || null;
}

function teamFormToObject() {
  const form = document.getElementById("teamPostureForm");
  const obj = {};
  const fd = new FormData(form);
  for (const [key, value] of fd.entries()) obj[key] = value;
  obj.verify_ssl = form.verify_ssl.checked;
  const member = activeTeamMember();
  obj.assignee_username = member ? member.username : "";
  obj.member_name = member ? member.name : "";
  return obj;
}

function updateTeamRollupHeader() {
  const qEl = document.getElementById("teamRollupQueueBacklog");
  const pEl = document.getElementById("teamRollupInProgress");
  const rEl = document.getElementById("teamRollupResolvedPeriod");
  const noteEl = document.getElementById("teamRollupNote");
  if (!qEl || !pEl || !rEl) return;
  if (!teamMembers.length) {
    qEl.textContent = "--";
    pEl.textContent = "--";
    rEl.textContent = "--";
    if (noteEl) noteEl.textContent = "";
    return;
  }
  let qb = 0;
  let ip = 0;
  let rs = 0;
  let cached = 0;
  for (const m of teamMembers) {
    const pl = teamPayloadByMemberId[m.id];
    if (!pl || !pl.metrics) continue;
    cached += 1;
    qb += Number(pl.metrics.queue_backlog_count ?? 0);
    ip += Number(pl.metrics.in_progress_count ?? 0);
    rs += Number(pl.metrics.resolved_in_period_count ?? 0);
  }
  if (cached === 0) {
    qEl.textContent = "--";
    pEl.textContent = "--";
    rEl.textContent = "--";
    if (noteEl) noteEl.textContent = "Refresh members to load team totals.";
    return;
  }
  qEl.textContent = String(qb);
  pEl.textContent = String(ip);
  rEl.textContent = String(rs);
  if (noteEl) {
    noteEl.textContent = cached < teamMembers.length
      ? `Team totals include ${cached}/${teamMembers.length} members with cached data. Run Refresh All Member Metrics for the full roster.`
      : "";
  }
}

function renderTeamPostureMetrics(payload) {
  const metrics = payload.metrics || {};
  const oldest = payload.oldest_open || {};
  const resolvedOwned = metrics.resolved_owned_count ?? metrics.resolved_count ?? 0;
  document.getElementById("teamResolvedOwnedCount").textContent = String(resolvedOwned);
  document.getElementById("teamResolvedContributedCount").textContent = String(metrics.resolved_contributed_count ?? 0);
  document.getElementById("teamResolvedLast8hCount").textContent = String(metrics.resolved_last_8h_count ?? 0);
  document.getElementById("teamOpenCount").textContent = String(metrics.assigned_open_count ?? 0);
  const qbEl = document.getElementById("teamQueueBacklogCount");
  const ipEl = document.getElementById("teamInProgressCount");
  const wsEl = document.getElementById("teamWorkedStatusLast8hCount");
  if (qbEl) qbEl.textContent = String(metrics.queue_backlog_count ?? 0);
  if (ipEl) ipEl.textContent = String(metrics.in_progress_count ?? 0);
  if (wsEl) wsEl.textContent = String(metrics.worked_status_last_8h_count ?? 0);
  document.getElementById("teamReopenedCount").textContent = String(metrics.reopened_count ?? 0);
  document.getElementById("teamWorkedOtherCount").textContent = String(metrics.worked_on_assigned_others_count ?? 0);
  document.getElementById("teamSlaBreachCount").textContent = String(metrics.sla_breach_count ?? 0);
  document.getElementById("teamSlaNearCount").textContent = String(metrics.open_near_sla_breach_8h_count ?? 0);
  document.getElementById("teamOldestTicket").textContent = oldest.issue_key || "N/A";
  document.getElementById("teamOldestAge").textContent = String(oldest.age_days ?? "--");

  const status = payload.status_distribution || {};
  const statusLines = Object.entries(status)
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .map(([k, v]) => `${k}: ${v}`);
  document.getElementById("teamStatusSummary").textContent = statusLines.join("\\n") || "No status counts.";
  document.getElementById("teamOldestDetail").textContent = JSON.stringify(oldest || {}, null, 2);
  renderTeamLabelsChart(payload.label_distribution || {});
  renderTeamCsvPreview(payload.raw_rows || []);
  updateTeamRollupHeader();
}

function renderTeamLabelsChart(labelDistribution) {
  const el = document.getElementById("teamLabelsChart");
  if (!el) return;
  const entries = Object.entries(labelDistribution || {}).sort((a, b) => Number(b[1]) - Number(a[1]));
  const ctx = el.getContext("2d");
  destroyChart(teamCharts.labels);
  teamCharts.labels = new Chart(ctx, {
    type: "pie",
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{ data: entries.map(([, v]) => v) }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "right" },
      },
    },
  });
}

function toCsv(rows, headers) {
  const esc = (val) => {
    const s = String(val ?? "");
    if (s.includes(",") || s.includes("\\\"") || s.includes("\\n")) {
      return `"${s.replace(/"/g, "\\\"\\\"")}"`;
    }
    return s;
  };
  const allHeaders = headers && headers.length ? headers : (rows[0] ? Object.keys(rows[0]) : []);
  const lines = [allHeaders.join(",")];
  for (const row of rows) {
    lines.push(allHeaders.map((h) => esc(row[h])).join(","));
  }
  return lines.join("\\n");
}

function renderTeamCsvPreview(rawRows) {
  const previewRows = (rawRows || []).slice(0, 8);
  if (!previewRows.length) {
    document.getElementById("teamCsvPreview").textContent = "No rows available for preview.";
    return;
  }
  const headers = Object.keys(previewRows[0]);
  document.getElementById("teamCsvPreview").textContent = toCsv(previewRows, headers);
}

async function refreshTeamPosture() {
  updateTeamReportPeriodLabel();
  const member = activeTeamMember();
  if (!member) {
    document.getElementById("teamStatusSummary").textContent = "Add and select a member first.";
    document.getElementById("teamCsvPreview").textContent = "Add and select a member first.";
    return;
  }
  const payload = teamFormToObject();
  try {
    const res = await fetch("/run-team-posture", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) {
      document.getElementById("teamStatusSummary").textContent = JSON.stringify(data, null, 2);
      document.getElementById("teamCsvPreview").textContent = JSON.stringify(data, null, 2);
      return;
    }
    latestTeamPosturePayload = data;
    if (member && member.id) {
      teamPayloadByMemberId[member.id] = data;
    }
    renderTeamPostureMetrics(data);
    updateTeamReportPeriodLabel();
  } catch (err) {
    const msg = `Network error calling /run-team-posture: ${err && err.message ? err.message : String(err)}. If this happens after edits, wait for Flask reload or restart app.py and retry.`;
    document.getElementById("teamStatusSummary").textContent = msg;
    document.getElementById("teamCsvPreview").textContent = msg;
  }
}

async function refreshAllTeamMembers() {
  updateTeamReportPeriodLabel();
  if (!teamMembers.length) {
    document.getElementById("teamStatusSummary").textContent = "Add and select a member first.";
    document.getElementById("teamCsvPreview").textContent = "Add and select a member first.";
    return;
  }
  const base = teamFormToObject();
  let successCount = 0;
  for (const member of teamMembers) {
    const payload = { ...base, assignee_username: member.username, member_name: member.name };
    const res = await fetch("/run-team-posture", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) {
      continue;
    }
    teamPayloadByMemberId[member.id] = data;
    successCount += 1;
  }
  if (!activeTeamMemberId && teamMembers[0]) {
    activeTeamMemberId = teamMembers[0].id;
  }
  const activePayload = teamPayloadByMemberId[activeTeamMemberId];
  if (activePayload) {
    latestTeamPosturePayload = activePayload;
    renderTeamPostureMetrics(activePayload);
  }
  if (!activePayload) {
    document.getElementById("teamStatusSummary").textContent = `No member data loaded (${successCount}/${teamMembers.length} successful).`;
  }
  updateTeamReportPeriodLabel();
  updateTeamRollupHeader();
}

function formatReportDatetimeLocal(value) {
  const s = (value || "").trim();
  if (!s) return "";
  return s.replace("T", " ");
}

function updateTeamReportPeriodLabel() {
  const form = document.getElementById("teamPostureForm");
  const el = document.getElementById("teamReportPeriod");
  if (!form || !el) return;
  const start = form.querySelector('input[name="start_dt"]')?.value || "";
  const end = form.querySelector('input[name="end_dt"]')?.value || "";
  if (!start && !end) {
    el.textContent = "Report period: set Start and End in Team Posture settings (created date range for the team JQL).";
    return;
  }
  const a = formatReportDatetimeLocal(start);
  const b = formatReportDatetimeLocal(end);
  if (start && end) {
    el.textContent = `Report period (created): ${a} → ${b}`;
  } else if (start) {
    el.textContent = `Report period (created): from ${a}`;
  } else {
    el.textContent = `Report period (created): through ${b}`;
  }
}

function updateLegacyReportPeriodLabel() {
  const form = document.getElementById("exportForm");
  const el = document.getElementById("legacyReportPeriod");
  if (!form || !el) return;
  const start = form.querySelector('input[name="start_dt"]')?.value || "";
  const end = form.querySelector('input[name="end_dt"]')?.value || "";
  const dateField = form.querySelector('select[name="date_field"]')?.value || "created";
  if (!start && !end) {
    el.textContent = `Report period: set Start and End in Report Settings (JQL uses the selected date field: ${dateField}).`;
    return;
  }
  const a = formatReportDatetimeLocal(start);
  const b = formatReportDatetimeLocal(end);
  if (start && end) {
    el.textContent = `Report period (${dateField}): ${a} → ${b}`;
  } else if (start) {
    el.textContent = `Report period (${dateField}): from ${a}`;
  } else {
    el.textContent = `Report period (${dateField}): through ${b}`;
  }
}

function setActiveReport(report) {
  document.querySelectorAll(".report-scope-csms").forEach((el) => { el.hidden = report !== "csms"; });
  document.querySelectorAll(".report-scope-team").forEach((el) => { el.hidden = report !== "team"; });
  document.querySelectorAll(".report-scope-legacy").forEach((el) => { el.hidden = report !== "legacy"; });
  document.querySelectorAll(".report-scope-auth").forEach((el) => { el.hidden = report !== "auth"; });
  document.querySelectorAll(".report-scope-notes").forEach((el) => { el.hidden = report !== "notes"; });
  document.querySelectorAll(".report-tab").forEach((btn) => {
    btn.classList.toggle("active", btn.getAttribute("data-report") === report);
  });
  window.scrollTo({ top: 0, behavior: "smooth" });
  if (report === "team") updateTeamReportPeriodLabel();
  if (report === "legacy") updateLegacyReportPeriodLabel();
}

document.querySelectorAll(".report-tab").forEach((btn) => {
  btn.addEventListener("click", () => setActiveReport(btn.getAttribute("data-report")));
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

const CSMS_KPI_TITLES = {
  backlog: "Open workload at the end of the current reporting window: tickets not yet in each project done status (CSSD uses Closed, CSD uses Ready For Production Users). Trend is percent change versus the previous window.",
  new_created: "How many tickets matched your filters and fell in the current period. Trend compares to the previous period of the same length.",
  resolved: "Tickets counted as finished in the current period using each project done status. Trend compares to the previous period.",
  longest_open: "Oldest ticket that was still open in the current period snapshot: key plus age in days from created.",
};

function renderCsmsKpis(kpis) {
  const longest = kpis.longest_open || {};
  const cards = [
    ["backlog", "Backlog Tickets", kpis.backlog.period2, kpis.backlog.trend],
    ["new_created", "New Created", kpis.new_created.period2, kpis.new_created.trend],
    ["resolved", "Resolved Tickets", kpis.resolved.period2, kpis.resolved.trend],
    ["longest_open", "Longest Open", `${longest.age_days || 0} days`, null, `${longest.issue_key || "N/A"}`],
  ];
  const html = cards.map(([metricKey, title, value, trend, subtext]) => `
    <div class="kpi-card" title="${CSMS_KPI_TITLES[metricKey] || ""}">
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
    <div title="How complete the CSMS process alignment appears for this run compared to your target percent.">
      <div>Process Alignment: CSMS ${pct.toFixed(1)}% Complete</div>
      <div class="progress-wrap"><div class="progress-bar" style="width:${Math.max(0, Math.min(100, pct))}%"></div></div>
    </div>
    <div title="Whether this run flagged a process gap from resolved performance.">
      <div>Process Gap Identified: ${gap}</div>
    </div>
  `;
}

function destroyChart(instance) {
  if (instance) instance.destroy();
}

const LEGACY_KPI_TITLES = [
  "Number of issues returned by your current legacy filters and caps.",
  "Total status changes counted across those issues.",
  "Total comments counted on those issues.",
  "Number of calendar days in this result set that have at least one created ticket.",
];

function renderLegacyKpis(kpis) {
  const container = document.getElementById("legacyKpis");
  if (!container) return;
  const cards = [
    ["Issue Count", kpis.issue_count || 0],
    ["Status Transitions", kpis.transition_count || 0],
    ["Comment Volume", kpis.comment_count || 0],
    ["Date Window Days", kpis.date_window_days || 0],
  ];
  container.innerHTML = cards.map(([title, value], i) => `
    <div class="kpi-card" title="${LEGACY_KPI_TITLES[i] || ""}">
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
  updateLegacyReportPeriodLabel();
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
  updateLegacyReportPeriodLabel();
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

document.getElementById("teamRefreshBtn").addEventListener("click", async () => {
  await refreshAllTeamMembers();
});

document.getElementById("teamAddMemberBtn").addEventListener("click", () => {
  const nameInput = document.getElementById("teamMemberNameInput");
  const usernameInput = document.getElementById("teamMemberUsernameInput");
  const name = (nameInput.value || "").trim();
  const username = (usernameInput.value || "").trim();
  if (!name || !username) return;
  const member = { id: `m_${Date.now()}`, name, username };
  teamMembers.push(member);
  activeTeamMemberId = member.id;
  saveTeamMembersToStorage(teamMembers);
  teamPayloadByMemberId = {};
  nameInput.value = "";
  usernameInput.value = "";
  renderTeamMemberIcons();
  updateTeamRollupHeader();
});

document.getElementById("teamRemoveMemberBtn").addEventListener("click", () => {
  if (!activeTeamMemberId) return;
  teamMembers = teamMembers.filter((m) => m.id !== activeTeamMemberId);
  if (teamMembers.length) {
    activeTeamMemberId = teamMembers[0].id;
  } else {
    activeTeamMemberId = null;
  }
  saveTeamMembersToStorage(teamMembers);
  teamPayloadByMemberId = {};
  renderTeamMemberIcons();
  updateTeamRollupHeader();
});

function openTeamExport(kind) {
  if (!latestTeamPosturePayload || !latestTeamPosturePayload.exports) return;
  const url = latestTeamPosturePayload.exports[kind];
  if (url) window.open(url, "_blank");
}

document.getElementById("teamExportCsvBtn").addEventListener("click", () => openTeamExport("csv"));
document.getElementById("teamExportExcelBtn").addEventListener("click", () => openTeamExport("excel"));
document.getElementById("teamExportAllBtn").addEventListener("click", () => {
  if (!teamMembers.length) {
    document.getElementById("teamStatusSummary").textContent = "Add and select a member first.";
    return;
  }

  const exportPayloads = [];
  const missingMembers = [];

  // Client-only: merge whatever is already in memory from Refresh / per-member loads. No network calls.
  for (const member of teamMembers) {
    const cached = teamPayloadByMemberId[member.id];
    if (cached) {
      exportPayloads.push(cached);
    } else {
      missingMembers.push(member);
    }
  }

  if (!exportPayloads.length) {
    document.getElementById("teamStatusSummary").textContent =
      "No cached team data yet. Click Refresh All Member Metrics (or select each member once), then try Download Team CSV again.";
    return;
  }

  // Merge member-level raw rows into one board export dataset.
  const allRows = [];
  for (const payload of exportPayloads) {
    const memberName = payload?.member?.name ?? "";
    const assigneeUsername = payload?.member?.assignee_username ?? "";
    const rawRows = Array.isArray(payload.raw_rows) ? payload.raw_rows : [];
    for (const raw of rawRows) {
      allRows.push({
        "Member Name": memberName,
        "Assignee Username": assigneeUsername,
        ...raw,
      });
    }
  }

  if (!allRows.length) {
    document.getElementById("teamStatusSummary").textContent =
      "Cached members have no dashboard ticket rows to export. Refresh metrics after changing filters, or confirm members have matching tickets.";
    return;
  }

  const headers = Object.keys(allRows[0]);
  const csv = toCsv(allRows, headers);
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `team_posture_board_${stamp}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  const statusParts = [
    `Team CSV downloaded (${allRows.length} rows) from cached data only — no server request.`,
    missingMembers.length
      ? `Not in cache (skipped): ${missingMembers.map((m) => m.name || m.username).join(", ")}. Use Refresh All Member Metrics to cache them.`
      : "All roster members were included.",
  ];
  document.getElementById("teamStatusSummary").textContent = statusParts.join(" ");
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

const teamPostureFormEl = document.getElementById("teamPostureForm");
if (teamPostureFormEl) {
  teamPostureFormEl.addEventListener("input", updateTeamReportPeriodLabel);
  teamPostureFormEl.addEventListener("change", updateTeamReportPeriodLabel);
}
const exportFormEl = document.getElementById("exportForm");
if (exportFormEl) {
  exportFormEl.addEventListener("input", updateLegacyReportPeriodLabel);
  exportFormEl.addEventListener("change", updateLegacyReportPeriodLabel);
}

teamMembers = loadTeamMembersFromStorage();
activeTeamMemberId = teamMembers[0] ? teamMembers[0].id : null;
renderTeamMemberIcons();
updateTeamReportPeriodLabel();
updateLegacyReportPeriodLabel();
updateTeamRollupHeader();

renderCsmsKpis({
  backlog: { period2: "--", trend: 0 },
  new_created: { period2: "--", trend: 0 },
  resolved: { period2: "--", trend: 0 },
  longest_open: { age_days: "--", issue_key: "" },
});
document.getElementById("csmsElapsed").textContent = "Provide Last Report Timestamp and run CSMS refresh to compute elapsed time.";
renderCsmsHealth({ process_alignment_pct: 60, process_gap_identified: "Pending run" });
renderLegacyKpis({ issue_count: 0, transition_count: 0, comment_count: 0, date_window_days: 0 });
</script>
</body>
</html>
"""


def parse_csv_list(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


def parse_optional_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except (TypeError, ValueError):
        return None


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
    # Optional explicit prior report KPI baselines.
    last_report_backlog = parse_optional_int(params.get("last_report_backlog_tickets"))
    last_report_new_created = parse_optional_int(params.get("last_report_new_created"))
    last_report_resolved = parse_optional_int(params.get("last_report_resolved_tickets"))
    trend_base_backlog = last_report_backlog if last_report_backlog is not None else p1_backlog
    trend_base_new_created = last_report_new_created if last_report_new_created is not None else p1_new
    trend_base_resolved = last_report_resolved if last_report_resolved is not None else p1_resolved

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
        "resolved_trend": calculate_percent_trend(trend_base_resolved, p2_resolved),
        "csd_backlog_trend": csd_backlog_trend,
    }

    payload = {
        "periods": periods,
        "elapsed_time_sentence": elapsed_time_sentence,
        "jql": {"period1": p1_jql, "period2": p2_jql},
        "kpis": {
            "backlog": {"period1": trend_base_backlog, "period2": p2_backlog, "trend": calculate_percent_trend(trend_base_backlog, p2_backlog)},
            "new_created": {"period1": trend_base_new_created, "period2": p2_new, "trend": calculate_percent_trend(trend_base_new_created, p2_new)},
            "resolved": {"period1": trend_base_resolved, "period2": p2_resolved, "trend": calculate_percent_trend(trend_base_resolved, p2_resolved)},
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


def build_team_posture_jql(params: Dict[str, Any], assignee_username: str, include_assignee: bool = True) -> str:
    projects = parse_csv_list(params.get("projects"))
    issue_types = parse_csv_list(params.get("issue_types"))
    clauses: List[str] = []
    for clause in [
        list_clause("project", projects),
        list_clause("issuetype", issue_types),
        list_clause("assignee", [assignee_username] if include_assignee and assignee_username else []),
    ]:
        if clause:
            clauses.append(clause)
    start_dt = normalize_dt_local(params.get("start_dt"))
    end_dt = normalize_dt_local(params.get("end_dt"))
    if start_dt:
        clauses.append(f'created >= "{start_dt}"')
    if end_dt:
        clauses.append(f'created <= "{end_dt}"')
    return (" AND ".join(clauses) if clauses else "order by created desc") + " ORDER BY created DESC"


def count_reopened_issues(issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    reopened = 0
    reopen_targets = {
        "open", "in progress", "new", "selected for development", "under qa analysis",
    }
    for issue in issues:
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules).lower()
        histories = ((issue.get("changelog") or {}).get("histories") or [])
        was_reopened = False
        for history in histories:
            for item in history.get("items", []):
                if (item.get("field") or "").lower() != "status":
                    continue
                from_status = (item.get("fromString") or "").strip().lower()
                to_status = (item.get("toString") or "").strip().lower()
                if from_status == final_status and (to_status in reopen_targets or to_status != final_status):
                    was_reopened = True
                    break
            if was_reopened:
                break
        if was_reopened:
            reopened += 1
    return reopened


def count_by_status_keywords(status_distribution: Dict[str, int], keywords: List[str]) -> int:
    total = 0
    wanted = [k.lower() for k in keywords]
    for status_name, count in status_distribution.items():
        lowered = (status_name or "").lower()
        if any(keyword in lowered for keyword in wanted):
            total += int(count or 0)
    return total


def issue_status_matches_keywords(issue: Dict[str, Any], keywords: Iterable[str]) -> bool:
    status = (get_issue_status(issue) or "").strip().lower()
    if not status:
        return False
    wanted = [k.lower() for k in keywords]
    return any(k in status for k in wanted)


def count_resolved_contributed_for_member(
    issues: List[Dict[str, Any]],
    assignee_username: str,
    csd_assigned_dev_field: str,
    keywords: Tuple[str, ...] = TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS,
) -> int:
    """Resolved-like tickets worked (status changelog) by member who is not the current owner."""
    target = (assignee_username or "").strip().lower()
    if not target:
        return 0
    count = 0
    for issue in issues:
        if issue_owner_username(issue, csd_assigned_dev_field) == target:
            continue
        if not member_has_status_change(issue, target):
            continue
        if issue_status_matches_keywords(issue, keywords):
            count += 1
    return count


def group_by_labels(issues: List[Dict[str, Any]]) -> Dict[str, int]:
    counts: Counter = Counter()
    for issue in issues:
        fields = issue.get("fields") or {}
        labels = fields.get("labels") or []
        if not labels:
            counts["No Label"] += 1
            continue
        for label in labels:
            val = (str(label) if label is not None else "").strip()
            counts[val or "No Label"] += 1
    return dict(counts)


def issue_current_assignee_username(issue: Dict[str, Any]) -> str:
    fields = issue.get("fields", {}) or {}
    assignee = fields.get("assignee") or {}
    return (assignee.get("name") or assignee.get("key") or assignee.get("emailAddress") or "").strip().lower()


def issue_assignee_matches(issue: Dict[str, Any], username: str) -> bool:
    fields = issue.get("fields", {}) or {}
    assignee = fields.get("assignee") or {}
    if isinstance(assignee, dict):
        return user_matches_username(assignee, username)
    if isinstance(assignee, str):
        return assignee.strip().lower() == (username or "").strip().lower()
    return False


def user_matches_username(user_obj: Dict[str, Any], username: str) -> bool:
    target = (username or "").strip().lower()
    if not target:
        return False
    candidates = [
        user_obj.get("name"),
        user_obj.get("key"),
        user_obj.get("emailAddress"),
        user_obj.get("displayName"),
        user_obj.get("accountId"),
        user_obj.get("value"),
    ]
    for item in candidates:
        if (item or "").strip().lower() == target:
            return True
    return False


def issue_owner_username(issue: Dict[str, Any], csd_assigned_dev_field: str) -> str:
    project_key = get_issue_project_key(issue)
    fields = issue.get("fields", {}) or {}
    if project_key == "CSD" and csd_assigned_dev_field:
        assigned_dev = fields.get(csd_assigned_dev_field)
        if isinstance(assigned_dev, dict):
            owner = (
                assigned_dev.get("name")
                or assigned_dev.get("key")
                or assigned_dev.get("emailAddress")
                or assigned_dev.get("accountId")
                or assigned_dev.get("value")
                or assigned_dev.get("displayName")
                or ""
            ).strip().lower()
            if owner:
                return owner
        if isinstance(assigned_dev, str):
            owner = assigned_dev.strip().lower()
            if owner:
                return owner
        if isinstance(assigned_dev, list):
            for item in assigned_dev:
                if isinstance(item, dict):
                    owner = (
                        item.get("name")
                        or item.get("key")
                        or item.get("emailAddress")
                        or item.get("accountId")
                        or item.get("value")
                        or item.get("displayName")
                        or ""
                    ).strip().lower()
                else:
                    owner = (str(item) if item is not None else "").strip().lower()
                if owner:
                    return owner
    return issue_current_assignee_username(issue)


def member_has_status_change(issue: Dict[str, Any], assignee_username: str) -> bool:
    target = (assignee_username or "").strip().lower()
    if not target:
        return False
    histories = ((issue.get("changelog") or {}).get("histories") or [])
    for history in histories:
        author = history.get("author") or {}
        if not user_matches_username(author, target):
            continue
        for item in history.get("items", []):
            if (item.get("field") or "").lower() == "status":
                return True
    return False


def issues_owned_by_member(issues: List[Dict[str, Any]], assignee_username: str, csd_assigned_dev_field: str) -> List[Dict[str, Any]]:
    target = (assignee_username or "").strip().lower()
    if not target:
        return []
    owned: List[Dict[str, Any]] = []
    for issue in issues:
        # Keep prior behavior: assignee match always counts as owned.
        if issue_assignee_matches(issue, target):
            owned.append(issue)
            continue
        # Add CSD Assigned Developer match as additional ownership path.
        project_key = get_issue_project_key(issue)
        if project_key == "CSD" and csd_assigned_dev_field:
            fields = issue.get("fields", {}) or {}
            assigned_dev = fields.get(csd_assigned_dev_field)
            if isinstance(assigned_dev, dict) and user_matches_username(assigned_dev, target):
                owned.append(issue)
                continue
            if isinstance(assigned_dev, str) and assigned_dev.strip().lower() == target:
                owned.append(issue)
                continue
            if isinstance(assigned_dev, list):
                matched = False
                for item in assigned_dev:
                    if isinstance(item, dict) and user_matches_username(item, target):
                        matched = True
                        break
                    if not isinstance(item, dict) and (str(item) if item is not None else "").strip().lower() == target:
                        matched = True
                        break
                if matched:
                    owned.append(issue)
    return owned


def issues_in_member_scope(issues: List[Dict[str, Any]], assignee_username: str, csd_assigned_dev_field: str) -> List[Dict[str, Any]]:
    target = (assignee_username or "").strip().lower()
    if not target:
        return []
    scoped: List[Dict[str, Any]] = []
    for issue in issues:
        owned = issue_owner_username(issue, csd_assigned_dev_field) == target
        worked_on = member_has_status_change(issue, target)
        if owned or worked_on:
            scoped.append(issue)
    return scoped


def get_issue_due_datetime(issue: Dict[str, Any], due_field_key: str = "duedate") -> Optional[datetime]:
    fields = issue.get("fields", {}) or {}
    raw = fields.get(due_field_key)
    if not raw:
        return None
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return None
        parsed = parse_jira_datetime(text)
        if parsed:
            return parsed
        try:
            # Jira duedate may be date-only (YYYY-MM-DD).
            d = datetime.strptime(text, "%Y-%m-%d")
            return d.replace(tzinfo=timezone.utc)
        except ValueError:
            return None
    return None


def get_issue_created_datetime(issue: Dict[str, Any]) -> Optional[datetime]:
    fields = issue.get("fields", {}) or {}
    return parse_jira_datetime(fields.get("created") or "")


def get_issue_finalized_datetime(issue: Dict[str, Any]) -> Optional[datetime]:
    fields = issue.get("fields", {}) or {}
    finalized = parse_jira_datetime(fields.get("resolutiondate") or "")
    if finalized:
        return finalized
    # Fallback for tickets in final status that may not have resolutiondate populated.
    return parse_jira_datetime(fields.get("updated") or "")


def _coerce_boolish(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "yes", "y", "1", "breached"}:
            return True
        if lowered in {"false", "no", "n", "0", "not breached"}:
            return False
    return None


def find_resolution_sla_field_key(base_url: str, verify_ssl: bool) -> Optional[str]:
    jira_root = get_jira_root_url(base_url)
    if not jira_root:
        return None
    if jira_root in RESOLUTION_SLA_FIELD_CACHE:
        return RESOLUTION_SLA_FIELD_CACHE[jira_root]
    auth = get_auth()
    session = requests.Session()
    resp = session.get(f"{jira_root}/rest/api/2/field", auth=auth, verify=verify_ssl, timeout=30)
    resp.raise_for_status()
    fields = resp.json() or []
    for item in fields:
        name = (item.get("name") or "").strip().lower()
        if "resolution sla breached" in name:
            key = (item.get("id") or "").strip()
            if key:
                RESOLUTION_SLA_FIELD_CACHE[jira_root] = key
                return key
    RESOLUTION_SLA_FIELD_CACHE[jira_root] = ""
    return None


def issue_resolution_sla_breached(issue: Dict[str, Any], resolution_sla_field_key: Optional[str]) -> Optional[bool]:
    if not resolution_sla_field_key:
        return None
    fields = issue.get("fields", {}) or {}
    raw = fields.get(resolution_sla_field_key)
    if isinstance(raw, dict):
        for key in ("value", "name"):
            parsed = _coerce_boolish(raw.get(key))
            if parsed is not None:
                return parsed
    if isinstance(raw, list):
        for item in raw:
            if isinstance(item, dict):
                parsed = _coerce_boolish(item.get("value") or item.get("name"))
            else:
                parsed = _coerce_boolish(item)
            if parsed is not None:
                return parsed
        return None
    return _coerce_boolish(raw)


def get_open_issues(issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> List[Dict[str, Any]]:
    open_issues: List[Dict[str, Any]] = []
    for issue in issues:
        status = (get_issue_status(issue) or "").strip().lower()
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules).strip().lower()
        if status != final_status:
            open_issues.append(issue)
    return open_issues


def is_issue_open_for_project(issue: Dict[str, Any], project_rules: Dict[str, str]) -> bool:
    project_key = get_issue_project_key(issue)
    status = (get_issue_status(issue) or "").strip().lower()
    final_status = get_project_final_status(project_key, project_rules).strip().lower()
    return status != final_status


def is_queue_backlog_issue(issue: Dict[str, Any], project_rules: Dict[str, str]) -> bool:
    """CSSD: Under QA Analysis. CSD: New. Other projects: not counted."""
    project_key = get_issue_project_key(issue)
    status = (get_issue_status(issue) or "").strip().lower()
    if project_key == "CSSD":
        return "under qa analysis" in status
    if project_key == "CSD":
        return status == "new"
    return False


def is_in_progress_issue(issue: Dict[str, Any], project_rules: Dict[str, str]) -> bool:
    """
    CSSD: open, not New, not Under QA Analysis.
    CSD: open, not New.
    Other projects: not counted.
    """
    project_key = get_issue_project_key(issue)
    if project_key not in ("CSSD", "CSD"):
        return False
    if not is_issue_open_for_project(issue, project_rules):
        return False
    status = (get_issue_status(issue) or "").strip().lower()
    if project_key == "CSSD":
        if status == "new":
            return False
        if "under qa analysis" in status:
            return False
        return True
    if project_key == "CSD":
        return status != "new"
    return False


def count_owned_queue_backlog(owned_issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    return sum(1 for issue in owned_issues if is_queue_backlog_issue(issue, project_rules))


def count_owned_in_progress(owned_issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    return sum(1 for issue in owned_issues if is_in_progress_issue(issue, project_rules))


def issue_has_member_status_change_after(
    issue: Dict[str, Any], assignee_username: str, cutoff_utc: datetime
) -> bool:
    """True if the member authored a status transition in changelog at or after cutoff (UTC)."""
    target = (assignee_username or "").strip().lower()
    if not target:
        return False
    histories = ((issue.get("changelog") or {}).get("histories") or [])
    for history in histories:
        change_dt = parse_jira_datetime(history.get("created") or "")
        if not change_dt:
            continue
        change_dt = _as_utc(change_dt)
        if change_dt < cutoff_utc:
            continue
        author = history.get("author") or {}
        if not user_matches_username(author, target):
            continue
        for item in history.get("items", []):
            if (item.get("field") or "").lower() == "status":
                return True
    return False


def count_owned_status_changes_in_last_hours(
    owned_issues: List[Dict[str, Any]], assignee_username: str, hours: float = 8.0
) -> int:
    """Distinct owned issues with a status changelog entry authored by the member in the last `hours`."""
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=hours)
    n = 0
    for issue in owned_issues:
        if issue_has_member_status_change_after(issue, assignee_username, cutoff):
            n += 1
    return n


def parse_team_form_datetime(value: Optional[str]) -> Optional[datetime]:
    """Parse datetime-local style strings from Team Posture form for report-window comparisons."""
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        try:
            dt = datetime.strptime(text[:16], "%Y-%m-%dT%H:%M")
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)
    return dt


def count_owned_resolved_in_report_window(
    owned_issues: List[Dict[str, Any]],
    keywords: Tuple[str, ...],
    start_dt: Optional[datetime],
    end_dt: Optional[datetime],
) -> int:
    """Owned tickets in resolved-like status whose resolutiondate falls in [start_dt, end_dt] (inclusive)."""
    if start_dt is None and end_dt is None:
        return 0
    start_utc = _as_utc(start_dt) if start_dt else None
    end_utc = _as_utc(end_dt) if end_dt else None
    n = 0
    for issue in owned_issues:
        if not issue_status_matches_keywords(issue, keywords):
            continue
        fields = issue.get("fields", {}) or {}
        resolved_dt = parse_jira_datetime(fields.get("resolutiondate") or "")
        if not resolved_dt:
            continue
        resolved_dt = _as_utc(resolved_dt)
        if start_utc is not None and resolved_dt < start_utc:
            continue
        if end_utc is not None and resolved_dt > end_utc:
            continue
        n += 1
    return n


def compute_sla_metrics(
    issues: List[Dict[str, Any]],
    project_rules: Dict[str, str],
    base_url: str,
    verify_ssl: bool,
    sla_hours: float = 24.0,
) -> Dict[str, int]:
    now = datetime.now(timezone.utc)
    breached = 0
    near_breach = 0
    resolution_sla_field_key: Optional[str] = None
    try:
        resolution_sla_field_key = find_resolution_sla_field_key(base_url, verify_ssl)
    except Exception:
        resolution_sla_field_key = None

    for issue in issues:
        created_dt = get_issue_created_datetime(issue)
        if not created_dt:
            continue
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules).strip().lower()
        current_status = (get_issue_status(issue) or "").strip().lower()
        is_open = current_status != final_status
        if is_open:
            elapsed_hours = (now - created_dt).total_seconds() / 3600.0
            remaining_hours = sla_hours - elapsed_hours
            if remaining_hours < 0:
                breached += 1
            elif remaining_hours < 8:
                near_breach += 1
            continue

        # Closed/finalized tickets: use Resolution SLA Breached indicator first.
        flagged = issue_resolution_sla_breached(issue, resolution_sla_field_key)
        if flagged is True:
            breached += 1
            continue
        if flagged is False:
            continue

        # Fallback: compute breach from created -> finalized timestamp.
        finalized_dt = get_issue_finalized_datetime(issue)
        if not finalized_dt:
            continue
        elapsed_to_final_hours = (finalized_dt - created_dt).total_seconds() / 3600.0
        if elapsed_to_final_hours > sla_hours:
            breached += 1
    return {
        "sla_breach_count": breached,
        "open_near_sla_breach_8h_count": near_breach,
    }


def count_worked_on_assigned_to_others(issues: List[Dict[str, Any]], assignee_username: str, csd_assigned_dev_field: str) -> int:
    target = (assignee_username or "").strip().lower()
    if not target:
        return 0
    count = 0
    for issue in issues:
        current_assignee = issue_owner_username(issue, csd_assigned_dev_field)
        if current_assignee == target:
            continue
        if member_has_status_change(issue, target):
            count += 1
    return count


def count_owned_resolved_in_last_hours(
    owned_issues: List[Dict[str, Any]],
    keywords: Tuple[str, ...],
    hours: float = 8.0,
) -> int:
    """Owned tickets whose Jira resolutiondate falls within the last `hours` from now (UTC)."""
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=hours)
    n = 0
    for issue in owned_issues:
        if not issue_status_matches_keywords(issue, keywords):
            continue
        fields = issue.get("fields", {}) or {}
        resolved_dt = parse_jira_datetime(fields.get("resolutiondate") or "")
        if not resolved_dt:
            continue
        if resolved_dt.tzinfo is None:
            resolved_dt = resolved_dt.replace(tzinfo=timezone.utc)
        if resolved_dt >= cutoff:
            n += 1
    return n


def count_reopened_for_member(issues: List[Dict[str, Any]], assignee_username: str, project_rules: Dict[str, str], csd_assigned_dev_field: str) -> int:
    target = (assignee_username or "").strip().lower()
    if not target:
        return 0
    reopened_status_keywords = ("reopened", "re-opened", "re opened")
    reopened = 0
    for issue in issues:
        current_owner = issue_owner_username(issue, csd_assigned_dev_field)
        current_status = (get_issue_status(issue) or "").strip().lower()
        if not any(k in current_status for k in reopened_status_keywords):
            continue
        # Include reopened tickets currently owned by member.
        if current_owner == target:
            reopened += 1
            continue
        # Include reopened tickets not owned by member but worked by member
        # through at least one status change history entry.
        if member_has_status_change(issue, target):
            reopened += 1
    return reopened


def _as_utc(dt: Optional[datetime]) -> Optional[datetime]:
    if not dt:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def build_member_dashboard_tagged_rows(
    broad_issues: List[Dict[str, Any]],
    assignee_username: str,
    csd_assigned_dev_field: str,
    project_rules: Dict[str, str],
    sla_hours: float = 24.0,
) -> List[Dict[str, Any]]:
    """
    Build ticket-level rows tagged with dashboard buckets expected by Operations Team:
    assigned open (CSSD/CSD), reopened, open <8h to SLA, oldest open.
    """
    target = (assignee_username or "").strip().lower()
    if not target:
        return []
    now = datetime.now(timezone.utc)
    reopened_status_keywords = ("reopened", "re-opened", "re opened")

    by_key: Dict[str, Dict[str, Any]] = {}
    for issue in broad_issues:
        key = str(issue.get("key") or "").strip()
        if key:
            by_key[key] = issue

    # Use the same ownership/scope logic as Team posture cards.
    owned_issues = issues_owned_by_member(broad_issues, assignee_username, csd_assigned_dev_field)
    member_scope_issues = issues_in_member_scope(broad_issues, assignee_username, csd_assigned_dev_field)
    open_owned_issues = get_open_issues(owned_issues, project_rules)

    tags_by_key: Dict[str, set] = {}

    # Assigned open CSSD/CSD/general tags.
    for issue in open_owned_issues:
        key = str(issue.get("key") or "").strip()
        if not key:
            continue
        project_key = get_issue_project_key(issue)
        if project_key == "CSSD":
            tag = "assigned_open_cssd"
        elif project_key == "CSD":
            tag = "assigned_open_csd"
        else:
            tag = "assigned_open"
        tags_by_key.setdefault(key, set()).add(tag)

    # Reopened tag using the same predicate as reopened card.
    for issue in broad_issues:
        key = str(issue.get("key") or "").strip()
        if not key:
            continue
        current_owner = issue_owner_username(issue, csd_assigned_dev_field)
        current_status = (get_issue_status(issue) or "").strip().lower()
        if not any(k in current_status for k in reopened_status_keywords):
            continue
        if current_owner == target or member_has_status_change(issue, target):
            tags_by_key.setdefault(key, set()).add("reopened")

    # Open tickets <8h to SLA tag based on same scope + 24h window logic.
    for issue in member_scope_issues:
        key = str(issue.get("key") or "").strip()
        if not key:
            continue
        project_key = get_issue_project_key(issue)
        final_status = get_project_final_status(project_key, project_rules).strip().lower()
        current_status = (get_issue_status(issue) or "").strip().lower()
        if current_status == final_status:
            continue
        created_dt = _as_utc(get_issue_created_datetime(issue))
        if not created_dt:
            continue
        remaining_hours = sla_hours - ((now - created_dt).total_seconds() / 3600.0)
        if 0 <= remaining_hours < 8:
            tags_by_key.setdefault(key, set()).add("open_lt_8h_to_sla")

    # Oldest open tag from same owned/open logic as oldest-open card.
    oldest_open = get_oldest_open_ticket(owned_issues, project_rules)
    oldest_key = str(oldest_open.get("issue_key") or "").strip()
    if oldest_key:
        tags_by_key.setdefault(oldest_key, set()).add("oldest_open")

    # Expanded export shape: one row per ticket per dashboard bucket.
    rows: List[Dict[str, Any]] = []
    for key in sorted(tags_by_key.keys()):
        issue = by_key.get(key)
        if not issue:
            continue
        raw = issue_to_raw_row(issue)
        for tag in sorted(tags_by_key[key]):
            rows.append({"Dashboard Bucket": tag, **raw})
    return rows


def build_team_posture_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")
    assignee_username = (params.get("assignee_username") or "").strip()
    if not assignee_username:
        raise ValueError("assignee_username is required")
    member_name = (params.get("member_name") or assignee_username).strip()

    page_size = int(params.get("page_size") or 50)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    csd_assigned_dev_field = (params.get("csd_assigned_dev_field") or "").strip()
    project_rules = {"CSSD": "Closed", "CSD": "Ready For Production Users"}
    jql = build_team_posture_jql(params, assignee_username, include_assignee=True)
    broad_jql = build_team_posture_jql(params, assignee_username, include_assignee=False)
    warnings: List[str] = []

    try:
        issues = fetch_jira_issues(base_url, jql, page_size, max_issues, verify_ssl, include_changelog=True)
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            issues = fetch_jira_issues(base_url, jql, page_size, max_issues, verify_ssl, include_changelog=False)
            warnings.append("Jira returned 500 with changelog expansion; reopened count may be incomplete.")
        else:
            raise

    try:
        broad_issues = fetch_jira_issues(base_url, broad_jql, page_size, max_issues, verify_ssl, include_changelog=True)
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            broad_issues = fetch_jira_issues(base_url, broad_jql, page_size, max_issues, verify_ssl, include_changelog=False)
            warnings.append("Jira returned 500 on broad team query with changelog expansion; worked-on/reopened may be incomplete.")
        else:
            raise

    owned_issues = issues_owned_by_member(broad_issues, assignee_username, csd_assigned_dev_field)
    member_scope_issues = issues_in_member_scope(broad_issues, assignee_username, csd_assigned_dev_field)
    sla_metrics = compute_sla_metrics(member_scope_issues, project_rules, base_url, verify_ssl, 24.0)
    status_distribution = group_by_status(owned_issues)
    label_distribution = group_by_labels(member_scope_issues)
    # Resolved/Open remain aligned to owner-matched assigned-ticket status distribution.
    resolved_owned_count = count_by_status_keywords(
        status_distribution,
        list(TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS),
    )
    resolved_contributed_count = count_resolved_contributed_for_member(
        broad_issues, assignee_username, csd_assigned_dev_field, TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS
    )
    open_count = get_backlog_count(owned_issues, project_rules)
    oldest_open = get_oldest_open_ticket(owned_issues, project_rules)
    reopened_count = count_reopened_for_member(broad_issues, assignee_username, project_rules, csd_assigned_dev_field)
    worked_on_assigned_others_count = count_worked_on_assigned_to_others(broad_issues, assignee_username, csd_assigned_dev_field)
    resolved_last_8h_count = count_owned_resolved_in_last_hours(
        owned_issues, TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS, hours=8.0
    )
    report_start = parse_team_form_datetime(params.get("start_dt"))
    report_end = parse_team_form_datetime(params.get("end_dt"))
    queue_backlog_count = count_owned_queue_backlog(owned_issues, project_rules)
    in_progress_count = count_owned_in_progress(owned_issues, project_rules)
    worked_status_last_8h_count = count_owned_status_changes_in_last_hours(
        owned_issues, assignee_username, hours=8.0
    )
    resolved_in_period_count = count_owned_resolved_in_report_window(
        owned_issues,
        TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS,
        report_start,
        report_end,
    )
    raw_rows = build_member_dashboard_tagged_rows(
        broad_issues,
        assignee_username,
        csd_assigned_dev_field,
        project_rules,
        sla_hours=24.0,
    )

    return {
        "member": {"name": member_name, "assignee_username": assignee_username},
        "jql": jql,
        "broad_jql": broad_jql,
        "warnings": warnings,
        "metrics": {
            # Back-compat: historically "resolved_count" counted owned resolved tickets only.
            "resolved_count": resolved_owned_count,
            "resolved_owned_count": resolved_owned_count,
            "resolved_contributed_count": resolved_contributed_count,
            "assigned_open_count": open_count,
            "reopened_count": reopened_count,
            "worked_on_assigned_others_count": worked_on_assigned_others_count,
            "resolved_last_8h_count": resolved_last_8h_count,
            "queue_backlog_count": queue_backlog_count,
            "in_progress_count": in_progress_count,
            "worked_status_last_8h_count": worked_status_last_8h_count,
            "resolved_in_period_count": resolved_in_period_count,
            "sla_breach_count": sla_metrics["sla_breach_count"],
            "open_near_sla_breach_8h_count": sla_metrics["open_near_sla_breach_8h_count"],
        },
        "status_distribution": status_distribution,
        "label_distribution": label_distribution,
        "oldest_open": oldest_open,
        "raw_rows": raw_rows,
    }


def build_team_board_export_rows(params: Dict[str, Any], members: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for member in members:
        username = (member.get("username") or "").strip()
        if not username:
            continue
        member_params = dict(params)
        member_params["assignee_username"] = username
        member_params["member_name"] = (member.get("name") or username).strip()
        payload = build_team_posture_payload(member_params)
        for raw in payload.get("raw_rows") or []:
            rows.append(
                {
                    "Member Name": payload["member"]["name"],
                    "Assignee Username": payload["member"]["assignee_username"],
                    **raw,
                }
            )
    return rows


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
    page = HTML.replace("__DEFAULT_TEAM_ROSTER_JSON__", json.dumps(TEAM_DEFAULT_MEMBERS))
    return render_template_string(page)


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


@app.route("/run-team-posture", methods=["POST"])
def run_team_posture():
    try:
        params = request.get_json(force=True)
        payload = build_team_posture_payload(params)
        out_dir = Path(tempfile.mkdtemp(prefix="team_posture_"))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_member = re.sub(r"[^A-Za-z0-9_-]+", "_", payload["member"]["name"])[:40] or "member"
        csv_path = out_dir / f"team_posture_{safe_member}_{stamp}.csv"
        summary_csv = out_dir / f"team_posture_summary_{safe_member}_{stamp}.csv"
        excel_path = out_dir / f"team_posture_{safe_member}_{stamp}.xlsx"

        raw_rows = payload["raw_rows"]
        write_csv(csv_path, raw_rows, list(raw_rows[0].keys()) if raw_rows else ["Issue Key"])
        mets = payload.get("metrics") or {}
        summary_rows = [
            {"Metric": "Resolved (Owned)", "Value": mets.get("resolved_owned_count", mets.get("resolved_count", 0))},
            {"Metric": "Resolved (Contributed)", "Value": mets.get("resolved_contributed_count", 0)},
            {"Metric": "Resolved (Last 8 Hours)", "Value": mets.get("resolved_last_8h_count", 0)},
            {"Metric": "Queue Backlog", "Value": mets.get("queue_backlog_count", 0)},
            {"Metric": "In Progress", "Value": mets.get("in_progress_count", 0)},
            {"Metric": "Worked Status (Last 8 Hours)", "Value": mets.get("worked_status_last_8h_count", 0)},
            {"Metric": "Resolved (Report Period)", "Value": mets.get("resolved_in_period_count", 0)},
            {"Metric": "Assigned Open Tickets", "Value": payload["metrics"]["assigned_open_count"]},
            {"Metric": "Reopened Tickets", "Value": payload["metrics"]["reopened_count"]},
            {"Metric": "Worked On (Assigned to Others)", "Value": payload["metrics"]["worked_on_assigned_others_count"]},
            {"Metric": "SLA Breach Count", "Value": payload["metrics"]["sla_breach_count"]},
            {"Metric": "Open Tickets < 8h to SLA Breach", "Value": payload["metrics"]["open_near_sla_breach_8h_count"]},
            {"Metric": "Oldest Open Ticket", "Value": payload["oldest_open"].get("issue_key", "")},
            {"Metric": "Oldest Open Age Days", "Value": payload["oldest_open"].get("age_days", "")},
        ]
        write_csv(summary_csv, summary_rows, ["Metric", "Value"])
        write_excel(
            excel_path,
            {
                "Raw Tickets": raw_rows,
                "Summary": summary_rows,
            },
        )
        export_id = uuid.uuid4().hex
        TEAM_EXPORT_CACHE[export_id] = {
            "csv": str(csv_path),
            "excel": str(excel_path),
        }
        payload["exports"] = {
            "csv": f"/download-team-posture-export?export_id={export_id}&kind=csv",
            "excel": f"/download-team-posture-export?export_id={export_id}&kind=excel",
        }
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/run-team-posture-board-export", methods=["POST"])
def run_team_posture_board_export():
    try:
        params = request.get_json(force=True)
        members = params.get("team_members") or []
        if not isinstance(members, list) or not members:
            return jsonify({"error": "team_members is required"}), 400
        out_dir = Path(tempfile.mkdtemp(prefix="team_posture_board_"))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = out_dir / f"team_posture_board_{stamp}.csv"
        rows = build_team_board_export_rows(params, members)
        if not rows:
            return jsonify({"error": "No valid team members provided"}), 400
        headers = list(rows[0].keys())
        write_csv(csv_path, rows, headers)
        export_id = uuid.uuid4().hex
        TEAM_EXPORT_CACHE[export_id] = {
            "csv": str(csv_path),
        }
        return jsonify(
            {
                "rows": rows,
                "exports": {
                    "csv": f"/download-team-posture-export?export_id={export_id}&kind=csv",
                },
            }
        )
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


@app.route("/download-team-posture-export")
def download_team_posture_export():
    export_id = request.args.get("export_id", "")
    kind = request.args.get("kind", "")
    if not export_id or export_id not in TEAM_EXPORT_CACHE:
        return Response("Invalid export id", status=400)
    path = TEAM_EXPORT_CACHE[export_id].get(kind)
    if not path:
        return Response("Invalid export kind", status=400)
    return send_file(path, as_attachment=True)


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
    # Keep debug diagnostics but disable auto-reloader to avoid transient
    # connection resets/address-in-use during frequent file edits.
    app.run(debug=True, use_reloader=False, port=5001)
