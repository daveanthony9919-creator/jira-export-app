
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

import snapshots_db as snap_db

app = Flask(__name__)
snap_db.init_db()
STATUS_TRANSITION_SLOTS = 30
CSMS_EXPORT_CACHE: Dict[str, Dict[str, str]] = {}
TEAM_EXPORT_CACHE: Dict[str, Dict[str, str]] = {}
TEAM_ISSUE_POOL_CACHE: Dict[str, Dict[str, Any]] = {}
TEAM_ISSUE_POOL_CACHE_MAX = 3
RESOLUTION_SLA_FIELD_CACHE: Dict[str, str] = {}

TEAM_EXPORT_SLIM_COLUMNS = (
    "Member Name",
    "Dashboard Bucket",
    "Issue Key",
    "Summary",
)

# Status substrings rolled up as "resolved" style outcomes for Team Posture resolved KPIs.
TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS: Tuple[str, ...] = (
    "resolved",
    "dev-completed",
    "closed",
    "ready for production users",
    "completed",
    "duplicate",
)

# Operations board: CSSD/CSD only, status must be Closed (exact name, case-insensitive).
TEAM_OPS_CLOSED_PROJECT_KEYS: Tuple[str, ...] = ("CSSD", "CSD")

# Default Pipeline Backlog JQL (matches Jira UI filter: Prod phase, early statuses, type/label exclusions).
TEAM_PIPELINE_BACKLOG_JQL_DEFAULT = (
    'project = "CSMS Defect Management" AND "Phase Reported" = Prod '
    'AND created >= "{created_since}" '
    'AND status in (New, "In Progress", Reopened) '
    'AND issuetype not in (Enhancement, "Enhancement Request", Story, Gap) '
    "AND (labels is EMPTY OR labels not in (Enhancement)) "
    "ORDER BY created ASC"
)
TEAM_PIPELINE_BACKLOG_CREATED_SINCE_DEFAULT = "2021-11-08"

# Jira SLA custom fields (legacy export + Ticket trend TTFR/TTR cards).
JIRA_SLA_TTR_FIELD = "customfield_10317"
JIRA_SLA_TTFR_FIELD = "customfield_10318"
LEGACY_SLA_SEARCH_FIELDS = (
    "project,status,created,updated,resolutiondate,issuetype,issuelinks,labels,"
    f"{JIRA_SLA_TTR_FIELD},{JIRA_SLA_TTFR_FIELD}"
)
DEFAULT_LEGACY_TTR_STATUS_CSSD = "Closed"
DEFAULT_LEGACY_TTR_STATUS_CSD = "Ready For Production Users"
DEFAULT_LEGACY_SLA_AGGREGATE = "median"
LEGACY_SLA_AGGREGATE_OPTIONS = ("median", "mean", "p90")

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
      overflow-x: hidden;
    }
    body.nav-open { overflow: hidden; }
    .wrap {
      max-width: 1800px;
      margin: 0;
      padding: 14px 18px 48px 124px;
      width: 100%;
      min-width: 0;
    }
    .app-mobile-bar,
    .nav-overlay {
      display: none;
    }
    .app-mobile-bar {
      align-items: center;
      gap: 12px;
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      height: 52px;
      padding: 0 12px;
      z-index: 100;
      background: var(--side-bg);
      border-bottom: 1px solid var(--border);
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.12);
    }
    .app-menu-toggle {
      width: 44px;
      height: 44px;
      padding: 0;
      flex-shrink: 0;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 22px;
      line-height: 1;
      border-radius: 10px;
    }
    .app-mobile-title {
      flex: 1;
      min-width: 0;
      font-size: 15px;
      font-weight: 800;
      color: var(--side-title);
      letter-spacing: .02em;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .nav-overlay {
      position: fixed;
      inset: 0;
      z-index: 110;
      background: rgba(8, 16, 32, 0.55);
      backdrop-filter: blur(2px);
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
    .muted-btn.danger-btn {
      color: var(--danger);
      border-color: rgba(220, 75, 85, 0.45);
    }
    .muted-btn.danger-btn:hover {
      background: rgba(220, 75, 85, 0.08);
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
      width: 100%;
      min-height: 38px;
      padding: 8px 12px;
      border-radius: 8px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 13px;
      line-height: 1;
    }
    .auth-tab-label { display: inline; }
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
      grid-auto-rows: min-content;
      gap: 10px;
      margin-top: 12px;
      align-items: start;
    }
    #teamRollupGrid.team-grid {
      grid-template-columns: repeat(5, minmax(130px, 1fr));
    }
    .team-metric-card {
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel-2);
      padding: 12px;
      height: auto;
      min-height: 0;
      align-self: start;
      overflow: hidden;
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
    .legacy-chart-wrap,
    .csms-chart-wrap {
      width: 100%;
      max-width: 100%;
      min-width: 0;
      position: relative;
      overflow: hidden;
      border-radius: 10px;
      background: var(--panel);
      border: 1px solid var(--border);
      padding: 8px;
      box-sizing: border-box;
    }
    .legacy-chart-wrap > canvas,
    .csms-chart-wrap > canvas {
      display: block;
      width: 100% !important;
      height: 100% !important;
      max-width: 100%;
    }
    .legacy-chart-wrap.daily {
      height: clamp(200px, 32vh, 280px);
    }
    .legacy-chart-wrap.labels-bar {
      min-height: 200px;
      height: clamp(200px, 36vh, 320px);
    }
    .legacy-chart-wrap.labels-trend {
      min-height: 220px;
      height: clamp(220px, 36vh, 320px);
    }
    .team-labels-charts {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(min(100%, 300px), 1fr));
      gap: 14px;
      margin-top: 8px;
      width: 100%;
      min-width: 0;
    }
    .team-labels-charts > div {
      min-width: 0;
      max-width: 100%;
    }
    .team-labels-charts h3 {
      margin: 0 0 8px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: var(--heading-strong);
    }
    .team-labels-hint {
      margin: 6px 0 0;
      font-size: 12px;
      color: var(--muted);
    }
    .legacy-labels-block {
      margin-top: 4px;
      scroll-margin-top: 72px;
    }
    .legacy-chart-wrap.status {
      height: clamp(200px, 32vh, 280px);
    }
    .csms-chart-wrap.daily {
      height: clamp(150px, 28vh, 220px);
    }
    .csms-chart-wrap.small {
      height: clamp(140px, 26vh, 200px);
    }
    .csms-chart-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(min(100%, 280px), 1fr));
      gap: 8px;
      margin-top: 8px;
      width: 100%;
      min-width: 0;
    }
    .csms-chart-grid > .csms-chart-wrap {
      min-width: 0;
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
      .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .kpi-number { font-size: 36px; }
      .section-title { font-size: 32px; }
      .section-subtitle { font-size: 18px; }
      .csms-chart-grid { grid-template-columns: 1fr; }
      .team-grid,
      #teamRollupGrid.team-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
      .hero-head {
        flex-direction: column;
        align-items: flex-start;
      }
      .status-pill {
        width: 100%;
        max-width: 100%;
      }
      .snapshot-settings-block .snapshot-toolbar,
      .snapshot-settings-block .baseline-toolbar {
        display: flex;
        flex-direction: column;
        align-items: stretch;
      }
      .snapshot-settings-block .snapshot-toolbar > label,
      .snapshot-settings-block .snapshot-toolbar select,
      .snapshot-settings-block .snapshot-toolbar input[type="text"],
      .snapshot-settings-block .snapshot-toolbar button,
      .snapshot-settings-block .baseline-toolbar input[type="text"],
      .snapshot-settings-block .baseline-toolbar input[type="number"],
      .snapshot-settings-block .baseline-toolbar button {
        grid-column: auto;
        width: 100%;
      }
      pre { max-width: 100%; overflow-x: auto; }
      .card:not(.app-nav) { overflow-x: clip; }
      .legacy-chart-wrap,
      .csms-chart-wrap {
        height: clamp(180px, 42vw, 260px);
      }
      .legacy-chart-wrap.labels-bar {
        height: clamp(200px, 50vw, 360px);
      }
    }
    @media (max-width: 768px) {
      .app-mobile-bar {
        display: flex;
      }
      .nav-overlay:not([hidden]) {
        display: block;
      }
      .app-nav {
        transform: translateX(-105%);
        transition: transform 0.22s ease;
        width: min(300px, 88vw);
        z-index: 120;
        padding-top: 58px;
        overflow-y: auto;
        -webkit-overflow-scrolling: touch;
      }
      .app-nav.is-open {
        transform: translateX(0);
        box-shadow: 8px 0 24px rgba(0, 0, 0, 0.25);
      }
      .app-nav-title {
        font-size: 13px;
        margin-bottom: 4px;
      }
      .app-nav .muted-btn {
        min-height: 44px;
        font-size: 13px;
        padding: 10px 12px;
      }
      .wrap {
        padding: 62px 12px 28px;
        max-width: 100%;
      }
      h1, h1.section-title { font-size: 26px; }
      .section-subtitle { font-size: 16px; }
      .member-grid {
        gap: 8px;
      }
      .member-pill {
        min-width: 72px;
        padding: 8px 10px;
      }
    }
    @media (max-width: 480px) {
      .team-grid,
      #teamRollupGrid.team-grid,
      .kpi-grid,
      .report-scope-csms .kpi-grid {
        grid-template-columns: 1fr;
      }
      .notes-grid {
        grid-template-columns: 1fr;
      }
      button, .muted-btn {
        width: 100%;
      }
      .row > button,
      .row > .muted-btn {
        width: auto;
        flex: 1 1 auto;
      }
      .snapshot-toolbar button {
        width: 100%;
      }
    }
    .archive-banner {
      background: var(--download-bg);
      color: var(--download-text);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 10px 14px;
      margin-bottom: 12px;
      font-size: 13px;
    }
    .snapshot-settings-block {
      margin-top: 16px;
      padding-top: 14px;
      border-top: 1px solid var(--border);
    }
    .snapshot-settings-block h3 {
      margin: 0 0 10px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .06em;
      color: var(--heading-strong);
    }
    .snapshot-toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      margin-bottom: 10px;
      padding: 0;
    }
    .snapshot-toolbar label { font-size: 12px; color: var(--muted); margin-right: 4px; }
    .snapshot-settings-block .snapshot-toolbar {
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 14px;
      align-items: end;
    }
    .snapshot-settings-block .snapshot-toolbar > label {
      grid-column: span 12;
      margin: 0;
    }
    .snapshot-settings-block .snapshot-toolbar select {
      grid-column: span 5;
    }
    .snapshot-settings-block .snapshot-toolbar input[type="text"] {
      grid-column: span 4;
    }
    .snapshot-settings-block .snapshot-toolbar button {
      grid-column: span 3;
    }
    .snapshot-settings-block .snapshot-toolbar select,
    .snapshot-settings-block .snapshot-toolbar input[type="text"] {
      width: 100%;
      min-width: 0;
      padding: 11px 12px;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: var(--panel-2);
      color: var(--text);
      font-size: 14px;
    }
    .snapshot-settings-block .baseline-toolbar {
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 14px;
      align-items: end;
      margin-top: 8px;
    }
    .snapshot-settings-block .baseline-toolbar input[type="text"] {
      grid-column: span 6;
    }
    .snapshot-settings-block .baseline-toolbar input[type="number"] {
      grid-column: span 3;
    }
    .snapshot-settings-block .baseline-toolbar button {
      grid-column: span 3;
    }
    .snapshot-status { font-size: 12px; color: var(--muted); margin: 8px 0 0; }
    .team-metric-card .metric-trend-sub {
      font-size: 11px;
      font-weight: 700;
      margin-top: 4px;
      min-height: 14px;
    }
    .team-metric-card .metric-spark-wrap {
      display: none !important;
    }
    .kpi-card .metric-trend-sub { font-size: 11px; font-weight: 700; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="nav-overlay" id="navOverlay" hidden aria-hidden="true"></div>
  <header class="app-mobile-bar">
    <button type="button" class="muted-btn app-menu-toggle" id="navMenuToggle" aria-expanded="false" aria-controls="appNav" aria-label="Open navigation menu" title="Menu">☰</button>
    <span class="app-mobile-title" id="appMobileTitle">CSMS Reporting</span>
  </header>
  <nav class="card app-nav" id="appNav" aria-label="Main navigation">
    <div class="app-nav-shell">
      <h1 class="app-nav-title">CSMS Reporting</h1>
      <div class="app-nav-actions">
        <button type="button" class="muted-btn report-tab active" data-report="csms" title="Executive Report">Executive Report</button>
        <button type="button" class="muted-btn report-tab" data-report="team" title="Operations Team">Operations Team</button>
        <button type="button" class="muted-btn report-tab" data-report="legacy" title="Ticket trend">Ticket trend</button>
        <button type="button" class="muted-btn report-tab" data-report="notes" title="Notes & Guides">Notes</button>
        <button type="button" class="muted-btn report-tab auth-icon-btn" data-report="auth" title="Auth diagnostics" aria-label="Auth diagnostics"><span class="auth-tab-label">Auth</span></button>
        <button type="button" class="muted-btn theme-toggle" id="themeToggle" aria-label="Toggle theme" title="Switch theme">◐</button>
      </div>
    </div>
  </nav>

  <div class="wrap">
    <section id="csmsDashboardSection" class="app-section report-scope-csms">
      <div class="card">
        <div id="csmsArchiveBanner" class="archive-banner" hidden></div>
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
          <div class="field full snapshot-settings-block report-scope-csms">
            <h3>Official reports &amp; snapshots</h3>
            <div class="snapshot-toolbar">
              <label>Official report</label>
              <select id="csmsSnapshotSelect" data-report-id="exec"></select>
              <input type="text" id="csmsSnapshotNote" placeholder="Note for save (e.g. Monday exec)" />
              <button type="button" class="muted-btn" id="csmsSaveSnapshotBtn">Save snapshot</button>
              <button type="button" class="muted-btn" id="csmsLoadSnapshotParamsBtn" title="Copy report variables from the selected saved report into the form">Load saved settings</button>
              <button type="button" class="muted-btn" id="csmsRerunSnapshotBtn" title="Load saved settings, switch to Live, and run the executive report from Jira">Rerun with saved settings</button>
              <button type="button" class="muted-btn danger-btn" id="csmsDeleteSnapshotBtn" title="Permanently remove the selected saved report from SQLite">Delete snapshot</button>
            </div>
            <p id="csmsSnapshotStatus" class="snapshot-status small"></p>
          </div>
        </div>
      </div>
    </section>

    <section id="teamPostureSection" class="app-section report-scope-team" hidden>
      <div class="card">
        <div id="teamArchiveBanner" class="archive-banner" hidden></div>
        <div class="hero-head">
          <div>
            <h2>Team Member Ticket Posture</h2>
            <p id="teamReportPeriod" class="small" style="margin-top:8px;color:var(--muted);">Report period: set Start and End in Team Posture settings.</p>
          </div>
        </div>
        <div id="teamRollupGrid" class="team-grid" style="margin-bottom:12px;">
          <div class="team-metric-card" data-metric-scope="board" data-metric-key="pipeline_backlog_count" title="CSSD Prod defects in New, In Progress, or Reopened (Phase Reported = Prod, issue-type and label exclusions). Uses Pipeline Backlog JQL in Team settings."><div class="label">Pipeline Backlog</div><div id="teamPipelineBacklogCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="board" data-metric-key="queue_backlog_count" title="Sum of Queue Backlog counts across team members with cached data: CSSD tickets in Under QA Analysis plus CSD tickets in New. Other projects are excluded per member."><div class="label">Team Queue Backlog</div><div id="teamRollupQueueBacklog" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="board" data-metric-key="in_progress_count" title="Sum of In Progress counts across cached members: open CSSD tickets not in New or Under QA Analysis, and open CSD tickets not in New."><div class="label">Team In Progress</div><div id="teamRollupInProgress" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="board" data-metric-key="resolved_in_period_count" title="Sum across cached members: owned tickets in resolved-like status whose resolution time falls between Team Start and End."><div class="label">Team Resolved (Report Period)</div><div id="teamRollupResolvedPeriod" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="board" data-metric-key="closed_cssd_csd_team_count" title="Unique CSSD/CSD tickets in Closed status where any roster member is assignee/CSD dev owner or contributed a status change. Deduped across the team."><div class="label">Team Closed</div><div id="teamRollupClosedCssdCsd" class="value">--</div><div class="metric-trend-sub"></div></div>
        </div>
        <div id="teamMemberGrid" class="member-grid"></div>
        <div id="teamMetricsGrid" class="team-grid">
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="assigned_open_count" title="Owned CSSD/CSD only: CSSD not Resolved or Closed; CSD not Ready For Production Users (assignee or CSD Assigned Developer)."><div class="label">Assigned Open Tickets</div><div id="teamOpenCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="queue_backlog_count" title="CSSD: Under QA Analysis. CSD: New. Other projects: not counted. Uses ownership rules for this member."><div class="label">Queue Backlog</div><div id="teamQueueBacklogCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="in_progress_count" title="CSSD: open, not New, not Under QA Analysis. CSD: open and not New. Other projects: not counted."><div class="label">In Progress</div><div id="teamInProgressCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="worked_status_last_8h_count" title="Tickets you own where you authored a Jira status change in the changelog within the last eight hours from when this report ran. Requires changelog data from Jira."><div class="label">Worked Status (Last 8 Hours)</div><div id="teamWorkedStatusLast8hCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="worked_status_last_8h_assigned_others_count" title="Tickets owned by someone else under the same ownership rules as Worked On (Assigned to Others), where you authored a status change in the changelog within the last eight hours. Requires changelog data from Jira."><div class="label">Worked Status (Others, Last 8 Hours)</div><div id="teamWorkedStatusOthersLast8hCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="reopened_count" title="Tickets in the date window whose status sounds reopened, where the member owns the ticket or authored at least one status change."><div class="label">Reopened Tickets</div><div id="teamReopenedCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="resolved_owned_count" title="Tickets assigned to this member whose status looks finished, such as resolved, closed, completed, or duplicate."><div class="label">Resolved (Owned)</div><div id="teamResolvedOwnedCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="resolved_contributed_count" title="Finished tickets owned by someone else where this member changed the status at least once."><div class="label">Resolved (Contributed)</div><div id="teamResolvedContributedCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="closed_cssd_csd_count" title="CSSD and CSD tickets in Closed status only, where you are assignee/CSD Assigned Developer or you contributed a status change (combined, no double count)."><div class="label">Closed</div><div id="teamClosedCssdCsdCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="resolved_last_8h_count" title="Tickets you own that have a Jira resolution time in the rolling last eight hours from when this report ran. Uses the same finished-status rules as Resolved (Owned)."><div class="label">Resolved (Last 8 Hours)</div><div id="teamResolvedLast8hCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="worked_on_assigned_others_count" title="Tickets owned by someone else where this member still changed the status at least once."><div class="label">Worked On (Assigned to Others)</div><div id="teamWorkedOtherCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="sla_breach_count" title="Tickets in this member scope that missed the 24-hour expectation from created time, using the Jira resolution SLA breached field when available, otherwise time to finish."><div class="label">SLA Breach Count</div><div id="teamSlaBreachCount" class="value">--</div><div class="metric-trend-sub"></div></div>
          <div class="team-metric-card" data-metric-scope="member" data-metric-key="open_near_sla_breach_8h_count" title="Open tickets still within 24 hours from created but with under eight hours left before that window ends."><div class="label">Open Tickets &lt; 8h to SLA Breach</div><div id="teamSlaNearCount" class="value">--</div><div class="metric-trend-sub"></div></div>
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
      <div class="card" style="margin-top:18px;" title="Jira label counts for the selected member: current bar chart and trends from saved Operations snapshots.">
        <h2>Ticket Labels</h2>
        <p id="teamLabelsHint" class="team-labels-hint">Select a member and refresh. Save snapshots to build label trends over time.</p>
        <div class="team-labels-charts">
          <div>
            <h3>Current (top labels)</h3>
            <div class="legacy-chart-wrap labels-bar">
              <canvas id="teamLabelsBarChart"></canvas>
            </div>
          </div>
          <div>
            <h3>Trend</h3>
            <div class="legacy-chart-wrap labels-trend">
              <canvas id="teamLabelsTrendChart"></canvas>
            </div>
          </div>
        </div>
      </div>
      <div id="teamSettingsCard" class="card collapse-card report-scope-team" style="margin-top:18px;">
        <button type="button" class="muted-btn collapse-toggle" data-collapse-target="teamSettings" aria-expanded="false" aria-controls="teamSettings">Show Team Posture Variables & Settings</button>
        <div id="teamSettings" class="collapse-body" hidden>
          <p id="teamDataModeHint" class="small" style="margin:0 0 10px;color:var(--muted);"></p>
          <div class="row" style="margin-bottom:14px;flex-wrap:wrap;gap:8px;">
            <button type="button" class="muted-btn" id="teamRefreshBtnMain">Refresh from Jira</button>
            <button type="button" class="muted-btn" id="teamRefreshActiveBtn">Refresh selected member</button>
          </div>
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
            <div class="field"><label>Pipeline Backlog Created Since</label><input type="date" name="pipeline_backlog_created_since" value="2021-11-08" title="Used in default Pipeline Backlog JQL as created &gt;= this date." /></div>
            <div class="field full">
              <label>Pipeline Backlog JQL (optional override)</label>
              <textarea name="pipeline_backlog_jql" rows="4" placeholder="Leave blank to use the default CSMS Prod filter (New / In Progress / Reopened)."></textarea>
            </div>
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
          <div class="field full snapshot-settings-block report-scope-team">
            <h3>Official reports &amp; snapshots</h3>
            <div class="snapshot-toolbar">
              <label>Official report</label>
              <select id="teamSnapshotSelect" data-report-id="ops"></select>
              <input type="text" id="teamSnapshotNote" placeholder="Note for save (e.g. AM ops)" />
              <button type="button" class="muted-btn" id="teamSaveSnapshotBtn">Save snapshot</button>
              <button type="button" class="muted-btn" id="teamLoadSnapshotParamsBtn" title="Copy team report variables (and roster when saved) from the selected official report">Load saved settings</button>
              <button type="button" class="muted-btn" id="teamRerunSnapshotBtn" title="Load saved settings, switch to Live, and refresh all members from Jira">Rerun with saved settings</button>
              <button type="button" class="muted-btn danger-btn" id="teamDeleteSnapshotBtn" title="Permanently remove the selected saved report from SQLite">Delete snapshot</button>
            </div>
            <p id="teamSnapshotStatus" class="snapshot-status small"></p>
            <details class="small" style="margin-top:10px;">
              <summary>Manual comparison baselines (fallback when no prior snapshot)</summary>
              <div class="baseline-toolbar">
                <input type="text" id="teamBaselineMetric" placeholder="metric_key e.g. pipeline_backlog_count" />
                <input type="number" id="teamBaselineValue" placeholder="Value" />
                <button type="button" class="muted-btn" id="teamBaselineSaveBtn">Save baseline</button>
              </div>
            </details>
          </div>
        </div>
      </div>
      <div id="teamCsvPreviewCard" class="card report-scope-team" style="margin-top:18px;">
        <h2>CSV Preview</h2>
        <pre id="teamCsvPreview">Run Team Posture refresh to preview CSV rows.</pre>
      </div>
    </section>

    <section id="legacyDashboardSection" class="app-section report-scope-legacy" hidden>
      <div class="card" style="margin-top:18px;">
        <div id="legacyArchiveBanner" class="archive-banner" hidden></div>
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
        <div id="legacyTicketLabelsSection" class="legacy-labels-block" title="Jira label counts for tickets in your Ticket trend query.">
          <h2 style="margin:20px 0 6px;padding-top:16px;border-top:1px solid var(--border);">Ticket Labels &amp; label trends</h2>
          <p id="legacyLabelsHint" class="team-labels-hint">Refresh the dashboard to see labels. Save snapshots to build label trends over time.</p>
          <div class="team-labels-charts">
            <div>
              <h3>Current (top labels)</h3>
              <div class="legacy-chart-wrap labels-bar">
                <canvas id="legacyLabelsBarChart"></canvas>
              </div>
            </div>
            <div>
              <h3>Trend (line chart)</h3>
              <div class="legacy-chart-wrap labels-trend" title="Top labels per saved official report; Live mode adds a Now column from the current refresh.">
                <canvas id="legacyLabelsTrendChart"></canvas>
              </div>
            </div>
          </div>
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
      <div id="legacySettingsCard" class="card collapse-card report-scope-legacy" style="margin-top:18px;">
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
            <div class="field full" style="margin-top:12px;">
              <h3 style="margin:0 0 8px;">SLA status gates (Ticket trend cards)</h3>
              <p class="small" style="margin:0 0 10px;color:var(--muted);">Comma-separated Jira status names. TTR uses customfield_10317 or resolutiondate − created. TTFR uses customfield_10318; CSD can inherit from linked CSSD.</p>
            </div>
            <div class="field"><label>TTR CSSD status</label><input name="ttr_status_cssd" value="Closed" title="Only CSSD tickets in these statuses count toward TTR CSSD." /></div>
            <div class="field"><label>TTR CSD status</label><input name="ttr_status_csd" value="Ready For Production Users" title="Only CSD tickets in these statuses count toward TTR CSD." /></div>
            <div class="field"><label>TTFR CSSD status</label><input name="ttfr_status_cssd" placeholder="Blank = any with TTFR SLA" title="Leave blank to include any CSSD ticket with Time to First Response SLA data." /></div>
            <div class="field"><label>TTFR CSD status</label><input name="ttfr_status_csd" placeholder="Blank = linked CSSD or own SLA" title="Leave blank for linked CSSD TTFR or CSD own SLA when no link." /></div>
            <div class="field full" style="margin-top:8px;">
              <p class="small" style="margin:0 0 8px;color:var(--muted);">Rollup across matching tickets: median (typical), mean (average), or 90th percentile.</p>
            </div>
            <div class="field"><label>TTFR CSSD aggregate</label>
              <select name="ttfr_cssd_aggregate" title="How to combine TTFR CSSD ticket hours into one card value.">
                <option value="median" selected>Median</option>
                <option value="mean">Mean (average)</option>
                <option value="p90">90th percentile</option>
              </select>
            </div>
            <div class="field"><label>TTFR CSD aggregate</label>
              <select name="ttfr_csd_aggregate">
                <option value="median" selected>Median</option>
                <option value="mean">Mean (average)</option>
                <option value="p90">90th percentile</option>
              </select>
            </div>
            <div class="field"><label>TTR CSSD aggregate</label>
              <select name="ttr_cssd_aggregate">
                <option value="median" selected>Median</option>
                <option value="mean">Mean (average)</option>
                <option value="p90">90th percentile</option>
              </select>
            </div>
            <div class="field"><label>TTR CSD aggregate</label>
              <select name="ttr_csd_aggregate">
                <option value="median" selected>Median</option>
                <option value="mean">Mean (average)</option>
                <option value="p90">90th percentile</option>
              </select>
            </div>
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
          <div class="field full snapshot-settings-block report-scope-legacy">
            <h3>Official reports &amp; snapshots</h3>
            <div class="snapshot-toolbar">
              <label>Official report</label>
              <select id="legacySnapshotSelect" data-report-id="legacy"></select>
              <input type="text" id="legacySnapshotNote" placeholder="Note for save" />
              <button type="button" class="muted-btn" id="legacySaveSnapshotBtn">Save snapshot</button>
              <button type="button" class="muted-btn" id="legacyLoadSnapshotParamsBtn" title="Copy Ticket trend report variables from the selected saved report">Load saved settings</button>
              <button type="button" class="muted-btn" id="legacyRerunSnapshotBtn" title="Load saved settings, switch to Live, and refresh the Ticket trend dashboard">Rerun with saved settings</button>
              <button type="button" class="muted-btn danger-btn" id="legacyDeleteSnapshotBtn" title="Permanently remove the selected saved report from SQLite">Delete snapshot</button>
            </div>
            <p id="legacySnapshotStatus" class="snapshot-status small"></p>
          </div>
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
              <li><strong>Pipeline Backlog:</strong> Count from the CSMS Prod pipeline JQL in Team settings (Jira search total, not a full export).</li>
              <li><strong>Team rollup cards:</strong> Sum or dedupe across cached members after refresh; delta % compares to the prior saved official report or manual baseline.</li>
              <li><strong>Resolved (Owned):</strong> Resolved/final statuses for tickets owned by the member (assignee plus CSD &quot;Assigned Developer&quot; when configured).</li>
              <li><strong>Resolved (Contributed):</strong> Resolved/final statuses for tickets where the member appears as a status-transition author but is not the current owner.</li>
              <li><strong>Assigned Open:</strong> Current open workload assigned to selected member.</li>
              <li><strong>Reopened Tickets:</strong> Includes assigned tickets and worked-on tickets with reopen history.</li>
              <li><strong>Worked On (Assigned to Others):</strong> Status-change author matches member, but assignee differs.</li>
              <li><strong>Ticket Count by Status:</strong> Distribution for assigned tickets in current filter window.</li>
            </ul>
          </div>
          <div class="notes-card">
            <h3>Ticket Trend SLA Cards</h3>
            <ul>
              <li><strong>Created / Updated / Resolved:</strong> Line chart by day. <strong>Ticket Labels:</strong> Current bar + label trend line chart from saved reports.</li>
              <li><strong>TTFR / TTR cards:</strong> Subline shows ▲/▼ % vs the <strong>prior saved official report</strong> (lower hours = green).</li>
              <li><strong>TTFR / TTR CSSD &amp; CSD:</strong> Configure status gates and aggregate (median, mean, or 90th percentile) under Report Settings, then <strong>Refresh Dashboard</strong>.</li>
              <li><strong>TTR:</strong> Jira Time to Resolution SLA (<code>customfield_10317</code>) or resolution date minus created; only tickets in your TTR status list.</li>
              <li><strong>TTFR CSD:</strong> Uses linked CSSD first-response time when a CSSD link exists.</li>
            </ul>
          </div>
          <div class="notes-card">
            <h3>How To Use</h3>
            <ul>
              <li>Open <strong>Team Posture Variables &amp; Settings</strong> for filters, roster, <strong>Refresh from Jira</strong>, and <strong>Refresh selected member</strong>.</li>
              <li>Open <strong>Report Variables &amp; Settings</strong> on Ticket trend for JQL, SLA gates, aggregates, and <strong>Refresh Dashboard</strong>.</li>
              <li>Choose an <strong>Official report</strong> snapshot or <strong>Live</strong>; Live still requires refresh to pull Jira.</li>
              <li>Use <strong>Save snapshot</strong> to store an official report in SQLite (refresh does not auto-save).</li>
              <li>Use Team member icons to switch per-member metrics.</li>
              <li><strong>Refresh from Jira</strong> loads dashboard metrics only (one team-wide Jira fetch). <strong>Download CSV</strong> / <strong>Download Team CSV</strong> build slim ticket rows on demand (Member Name, Dashboard Bucket, Issue Key, Summary) from the cached issue pool after refresh.</li>
              <li>Official reports: <strong>Load saved settings</strong> restores form variables (and team roster when saved); <strong>Rerun with saved settings</strong> runs a live refresh with those values; <strong>Delete snapshot</strong> removes the selected saved report (confirmation required).</li>
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
let csmsCharts = { daily: null, status: null, top: null };
let legacyCharts = { status: null, daily: null, labelsBar: null, labelsTrend: null };
let lastLegacyLabelDistribution = null;
let teamCharts = { labelsBar: null, labelsTrend: null };
let lastTeamLabelDistribution = null;
let latestCsmsPayload = null;
let latestLegacyPayload = null;
let latestBoardMetrics = {};
let latestBoardMetricsLoaded = false;
let latestPipelineBacklogLoaded = false;
let boardMetricsRefreshInFlight = null;
let teamBulkRefreshInFlight = null;
let teamPoolCacheId = null;
let teamJiraQueue = Promise.resolve();

/** Run team Jira API calls one at a time so Flask is not overloaded. */
function enqueueTeamJira(task) {
  const run = teamJiraQueue.then(() => task());
  teamJiraQueue = run.catch(() => {});
  return run;
}

let teamMembers = [];
let activeTeamMemberId = null;
let latestTeamPosturePayload = null;
let teamPayloadByMemberId = {};
const snapshotViewMode = { exec: "archive", ops: "archive", legacy: "archive" };
const activeSnapshotId = { exec: null, ops: null, legacy: null };
const metricSparkCharts = {};

function destroyChart(instance) {
  if (instance) instance.destroy();
}

function chartInstancesForResize() {
  return [
    ...Object.values(csmsCharts),
    ...Object.values(legacyCharts),
    ...Object.values(teamCharts),
    ...Object.values(metricSparkCharts),
  ].filter(Boolean);
}

function isChartContainerVisible(chart) {
  const canvas = chart && chart.canvas;
  if (!canvas) return false;
  const wrap = canvas.closest(".legacy-chart-wrap, .csms-chart-wrap, .metric-spark-wrap");
  if (!wrap || wrap.offsetParent === null) return false;
  const section = canvas.closest(".app-section");
  if (section && section.hidden) return false;
  return wrap.clientWidth > 0 && wrap.clientHeight > 0;
}

function resizeAllCharts() {
  for (const chart of chartInstancesForResize()) {
    if (!isChartContainerVisible(chart)) continue;
    try {
      chart.resize();
    } catch (_) { /* chart may be mid-destroy */ }
  }
}

function scheduleChartResize() {
  if (scheduleChartResize._timer) clearTimeout(scheduleChartResize._timer);
  scheduleChartResize._timer = setTimeout(() => {
    scheduleChartResize._timer = null;
    resizeAllCharts();
  }, 80);
}

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
  refreshChartsForTheme();
}

function getChartTheme() {
  const styles = getComputedStyle(document.documentElement);
  const pick = (name, fallback) => (styles.getPropertyValue(name).trim() || fallback);
  return {
    text: pick("--text", "#e8ecff"),
    muted: pick("--muted", "#9fb0e5"),
    border: pick("--border", "#2a376a"),
    panel: pick("--panel", "#121933"),
    heading: pick("--heading-strong", "#cddcff"),
  };
}

/** Distinct slice colors for pie/doughnut charts (readable on light and dark backgrounds). */
const CHART_SLICE_PALETTE = [
  "#4e9af5", "#f28e2b", "#59c9a5", "#e76f8b", "#9b7ede",
  "#f4b740", "#5ec8d8", "#ff8c69", "#7cb342", "#ba68c8",
  "#26a69a", "#ef5350", "#42a5f5", "#ab47bc", "#ffa726",
  "#66bb6a", "#29b6f6", "#ec407a", "#8d6e63", "#78909c",
  "#7e57c2", "#26c6da", "#d4e157", "#ff7043", "#5c6bc0",
];

function chartSliceColors(count) {
  const colors = [];
  for (let i = 0; i < count; i++) {
    colors.push(CHART_SLICE_PALETTE[i % CHART_SLICE_PALETTE.length]);
  }
  return colors;
}

function chartPieDataset(values, theme) {
  const colors = chartSliceColors(values.length);
  return {
    data: values,
    backgroundColor: colors,
    borderColor: theme.panel,
    borderWidth: 1,
    hoverBorderColor: theme.text,
    hoverBorderWidth: 1,
  };
}

function chartThemePlugins(theme, extraPlugins) {
  const extra = extraPlugins || {};
  const legendExtra = extra.legend || {};
  const titleExtra = extra.title || {};
  return {
    legend: {
      display: legendExtra.display !== false,
      position: legendExtra.position || "top",
      labels: {
        color: theme.text,
        boxWidth: 12,
        padding: 10,
        font: { size: 11 },
        usePointStyle: true,
        pointStyle: "rectRounded",
        ...(legendExtra.labels || {}),
      },
    },
    tooltip: {
      backgroundColor: theme.panel,
      titleColor: theme.heading,
      bodyColor: theme.text,
      borderColor: theme.border,
      borderWidth: 1,
      ...(extra.tooltip || {}),
    },
    title: titleExtra.display
      ? {
          display: true,
          text: titleExtra.text || "",
          color: theme.heading,
          font: { size: 13, weight: "600" },
          ...(titleExtra.font ? { font: titleExtra.font } : {}),
        }
      : { display: false },
  };
}

function chartThemeScales(theme) {
  const axis = {
    ticks: { color: theme.muted },
    grid: { color: theme.border },
    border: { color: theme.border },
  };
  return { x: { ...axis }, y: { ...axis } };
}

function refreshChartsForTheme() {
  const teamSection = document.getElementById("teamPostureSection");
  if (lastTeamLabelDistribution && teamSection && !teamSection.hidden) {
    refreshTeamLabelCharts(lastTeamLabelDistribution);
  }
  const csmsSection = document.getElementById("csmsDashboardSection");
  if (latestCsmsPayload && latestCsmsPayload.charts && csmsSection && !csmsSection.hidden) {
    renderCsmsCharts(latestCsmsPayload.charts);
  }
  const legacySection = document.getElementById("legacyDashboardSection");
  if (latestLegacyPayload && latestLegacyPayload.charts && legacySection && !legacySection.hidden) {
    renderLegacyCharts(latestLegacyPayload.charts);
    if (lastLegacyLabelDistribution) {
      void refreshLegacyLabelCharts(lastLegacyLabelDistribution);
    }
  }
  scheduleChartResize();
}

document.getElementById("themeToggle").addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "dark";
  applyTheme(current === "dark" ? "light" : "dark");
});

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

const snapshotParamsById = {};

const SNAPSHOT_PARAM_SKIP = new Set([
  "assignee_username", "member_name", "member_usernames", "fetch_board_metrics",
  "skip_pipeline", "skip_closed", "team_members",
]);

function objectToForm(form, params, skipKeys) {
  if (!form || !params || typeof params !== "object") return 0;
  const skip = skipKeys || SNAPSHOT_PARAM_SKIP;
  let applied = 0;
  for (const [key, value] of Object.entries(params)) {
    if (skip.has(key)) continue;
    const el = form.elements[key] || form.querySelector(`[name="${CSS.escape(key)}"]`);
    if (!el) continue;
    if (el.type === "checkbox") {
      el.checked = Boolean(value);
    } else if (el.tagName === "SELECT" || el.tagName === "INPUT" || el.tagName === "TEXTAREA") {
      el.value = value == null ? "" : String(value);
    } else {
      continue;
    }
    applied += 1;
  }
  return applied;
}

function applyTeamMembersFromParams(params) {
  const saved = params && params.team_members;
  if (!Array.isArray(saved) || !saved.length) return false;
  const merged = saved.map((m, i) => ({
    id: (m.id || `snap-${i}-${(m.username || "").toLowerCase()}`).toString(),
    name: (m.name || m.username || "").trim(),
    username: (m.username || "").trim(),
  })).filter((m) => m.username);
  if (!merged.length) return false;
  teamMembers = merged;
  saveTeamMembersToStorage(teamMembers);
  if (!activeTeamMemberId || !teamMembers.some((m) => m.id === activeTeamMemberId)) {
    activeTeamMemberId = teamMembers[0].id;
  }
  renderTeamMemberIcons();
  return true;
}

function snapshotSelectForReport(reportUiKey) {
  const id = reportUiKey === "csms" ? "csmsSnapshotSelect"
    : reportUiKey === "team" ? "teamSnapshotSelect" : "legacySnapshotSelect";
  return document.getElementById(id);
}

function snapshotStatusElForReport(reportUiKey) {
  const id = reportUiKey === "csms" ? "csmsSnapshotStatus"
    : reportUiKey === "team" ? "teamSnapshotStatus" : "legacySnapshotStatus";
  return document.getElementById(id);
}

async function getSelectedSnapshotParams(reportUiKey) {
  const sel = snapshotSelectForReport(reportUiKey);
  if (!sel || sel.value === "live") return null;
  const snapId = parseInt(sel.value, 10);
  if (!snapId) return null;
  if (snapshotParamsById[snapId]) return snapshotParamsById[snapId];
  const reportId = REPORT_ID_MAP[reportUiKey];
  const display = await loadSnapshotDisplay(reportId, snapId);
  return display && display.params ? display.params : null;
}

function ensureLiveModeForReport(reportUiKey) {
  snapshotViewMode[reportUiKey] = "live";
  const reportId = REPORT_ID_MAP[reportUiKey];
  activeSnapshotId[reportId] = null;
  const sel = snapshotSelectForReport(reportUiKey);
  if (sel) sel.value = "live";
  setArchiveBanner(reportUiKey, false, "");
  if (reportUiKey === "team") {
    latestBoardMetricsLoaded = false;
    latestPipelineBacklogLoaded = false;
    latestBoardMetrics = {};
    teamPayloadByMemberId = {};
    updateTeamRollupHeader();
    updateTeamDataModeHint();
  }
}

function hydrateFormFromSnapshotParams(reportUiKey, params) {
  if (!params || typeof params !== "object") return { applied: 0, roster: false };
  let applied = 0;
  let roster = false;
  if (reportUiKey === "csms") {
    applied = objectToForm(document.getElementById("csmsForm"), params);
  } else if (reportUiKey === "legacy") {
    applied = objectToForm(document.getElementById("exportForm"), params);
    updateLegacyReportPeriodLabel();
  } else if (reportUiKey === "team") {
    roster = applyTeamMembersFromParams(params);
    applied = objectToForm(document.getElementById("teamPostureForm"), params);
    updateTeamReportPeriodLabel();
  }
  return { applied, roster };
}

async function loadSavedReportSettings(reportUiKey) {
  const statusEl = snapshotStatusElForReport(reportUiKey);
  const params = await getSelectedSnapshotParams(reportUiKey);
  if (!params) {
    const msg = "Select a saved official report (not Live), then click Load saved settings.";
    if (statusEl) statusEl.textContent = msg;
    else alert(msg);
    return false;
  }
  const { applied, roster } = hydrateFormFromSnapshotParams(reportUiKey, params);
  const parts = [`Loaded ${applied} setting(s) from saved report.`];
  if (reportUiKey === "team" && roster) parts.push(`Roster: ${teamMembers.length} member(s).`);
  if (reportUiKey === "team" && !roster) {
    parts.push("No roster in snapshot — using current team members in the browser.");
  }
  if (statusEl) statusEl.textContent = parts.join(" ");
  return true;
}

async function rerunWithSavedReportSettings(reportUiKey) {
  const statusEl = snapshotStatusElForReport(reportUiKey);
  const loaded = await loadSavedReportSettings(reportUiKey);
  if (!loaded) return;
  ensureLiveModeForReport(reportUiKey);
  if (statusEl) statusEl.textContent = (statusEl.textContent || "") + " Running live refresh…";
  if (reportUiKey === "csms") {
    await runCsmsExecutiveReport();
  } else if (reportUiKey === "legacy") {
    document.getElementById("legacyInsights").textContent = "Refreshing dashboard…";
    await refreshLegacyDashboard(formToObject(document.getElementById("exportForm")));
  } else {
    await onTeamRefreshAllClick();
  }
}

async function deleteSelectedSnapshot(reportUiKey) {
  const statusEl = snapshotStatusElForReport(reportUiKey);
  const sel = snapshotSelectForReport(reportUiKey);
  if (!sel || sel.value === "live") {
    const msg = "Select a saved official report to delete (not Live).";
    if (statusEl) statusEl.textContent = msg;
    else alert(msg);
    return;
  }
  const snapId = parseInt(sel.value, 10);
  if (!snapId) return;
  const label = sel.options[sel.selectedIndex]?.textContent || `Report #${snapId}`;
  if (!window.confirm(`Delete this saved report permanently?\n\n${label}\n\nThis cannot be undone.`)) {
    return;
  }
  const reportId = REPORT_ID_MAP[reportUiKey];
  const res = await fetch(
    `/snapshots/${snapId}?report_id=${encodeURIComponent(reportId)}`,
    { method: "DELETE" }
  );
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = data.error || "Delete failed";
    if (statusEl) statusEl.textContent = err;
    else alert(err);
    return;
  }
  delete snapshotParamsById[snapId];
  if (activeSnapshotId[reportId] === snapId) {
    activeSnapshotId[reportId] = null;
  }
  await loadSnapshotOptions(reportId, sel);
  sel.value = "live";
  ensureLiveModeForReport(reportUiKey);
  if (statusEl) {
    statusEl.textContent = `Deleted saved report #${snapId}. Switched to Live — refresh from Jira when ready.`;
  }
}

async function runCsmsExecutiveReport() {
  const form = document.getElementById("csmsForm");
  const payload = csmsFormToObject(form);
  document.getElementById("csmsNarratives").textContent = "Running...";
  const res = await fetch("/run-csms-exec-summary", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
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

  if (!res.ok) {
    document.getElementById("downloads").innerHTML = "";
    const msg = data.error || data.details || "Export failed.";
    document.getElementById("legacyInsights").textContent = `Export error: ${msg}`;
    return;
  }

  if (data.downloads) {
    const parts = [];
    for (const item of data.downloads) {
      parts.push(`<a href="${item.url}">${item.label}</a>`);
    }
    document.getElementById("downloads").innerHTML = parts.join("");
  }
  await refreshLegacyDashboard(payload);
});

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

function updateTeamDataModeHint() {
  const el = document.getElementById("teamDataModeHint");
  if (!el) return;
  if (snapshotViewMode.team === "live") {
    const n = Object.keys(teamPayloadByMemberId).length;
    el.textContent = n
      ? `Live — ${n} member(s) cached in this session. Use Refresh from Jira to update.`
      : "Live — no Jira data loaded yet. Use Refresh from Jira or Refresh selected member.";
    return;
  }
  const snapId = activeSnapshotId.ops;
  el.textContent = snapId
    ? `Archived report #${snapId} — member metrics from the saved snapshot (no Jira calls).`
    : "Archived report — select a saved official report from settings.";
}

function clearTeamMemberMetricCards() {
  const ids = [
    "teamOpenCount", "teamQueueBacklogCount", "teamInProgressCount",
    "teamWorkedStatusLast8hCount", "teamWorkedStatusOthersLast8hCount",
    "teamReopenedCount", "teamResolvedOwnedCount", "teamResolvedContributedCount",
    "teamClosedCssdCsdCount", "teamResolvedLast8hCount", "teamWorkedOtherCount",
    "teamSlaBreachCount", "teamSlaNearCount", "teamOldestTicket", "teamOldestAge",
  ];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (el) el.textContent = "--";
  }
  document.querySelectorAll("#teamMetricsGrid .metric-trend-sub").forEach((sub) => {
    sub.textContent = "—";
    sub.className = "metric-trend-sub";
    sub.removeAttribute("title");
  });
}

function refreshOpsMetricTrendsForCurrentMode() {
  if (snapshotViewMode.team === "live") {
    return refreshOpsMetricTrends(null, "live");
  }
  if (activeSnapshotId.ops) {
    return refreshOpsMetricTrends(activeSnapshotId.ops, "archive");
  }
  return Promise.resolve();
}

function selectTeamMember(memberId) {
  if (teamBulkRefreshInFlight) return;
  activeTeamMemberId = memberId;
  renderTeamMemberIcons();
  updateTeamDataModeHint();
  const member = teamMembers.find((m) => m.id === memberId);
  if (!member) return;

  const cached = teamPayloadByMemberId[memberId];
  if (cached) {
    latestTeamPosturePayload = cached;
    renderTeamPostureMetrics(cached);
    updateTeamRollupHeader();
    void refreshOpsMetricTrendsForCurrentMode();
    return;
  }

  if (snapshotViewMode.team === "live") {
    void refreshTeamPosture();
    return;
  }

  clearTeamMemberMetricCards();
  const statusEl = document.getElementById("teamStatusSummary");
  const detailEl = document.getElementById("teamOldestDetail");
  const msg = `${member.name} is not in this saved report. Choose another member, or switch Official report to Live and refresh from Jira.`;
  if (statusEl) statusEl.textContent = msg;
  if (detailEl) detailEl.textContent = msg;
  destroyChart(teamCharts.labelsBar);
  destroyChart(teamCharts.labelsTrend);
  teamCharts.labelsBar = null;
  teamCharts.labelsTrend = null;
  const hintEl = document.getElementById("teamLabelsHint");
  if (hintEl) hintEl.textContent = msg;
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
    btn.disabled = Boolean(teamBulkRefreshInFlight);
    btn.addEventListener("click", () => {
      selectTeamMember(btn.getAttribute("data-member-id"));
    });
  });
}

function ensureLiveModeForTeamRefresh() {
  snapshotViewMode.team = "live";
  activeSnapshotId.ops = null;
  const sel = document.getElementById("teamSnapshotSelect");
  if (sel) sel.value = "live";
  setArchiveBanner("team", false, "");
  updateTeamDataModeHint();
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

function buildOpsBoardRollupFromCache() {
  const board = {
    pipeline_backlog_count: (latestPipelineBacklogLoaded || latestBoardMetricsLoaded)
      ? Number(latestBoardMetrics.pipeline_backlog_count ?? 0)
      : null,
    closed_cssd_csd_team_count: latestBoardMetricsLoaded
      ? Number(latestBoardMetrics.closed_cssd_csd_team_count ?? 0)
      : null,
    queue_backlog_count: 0,
    in_progress_count: 0,
    resolved_in_period_count: 0,
    sla_breach_count: 0,
    open_near_sla_breach_8h_count: 0,
  };
  let cached = 0;
  for (const m of teamMembers) {
    const pl = teamPayloadByMemberId[m.id]
      || (m.id === activeTeamMemberId ? latestTeamPosturePayload : null);
    if (!pl || !pl.metrics) continue;
    cached += 1;
    board.queue_backlog_count += Number(pl.metrics.queue_backlog_count ?? 0);
    board.in_progress_count += Number(pl.metrics.in_progress_count ?? 0);
    board.resolved_in_period_count += Number(pl.metrics.resolved_in_period_count ?? 0);
    board.sla_breach_count += Number(pl.metrics.sla_breach_count ?? 0);
    board.open_near_sla_breach_8h_count += Number(pl.metrics.open_near_sla_breach_8h_count ?? 0);
  }
  if (!latestBoardMetricsLoaded && cached > 0) {
    board.closed_cssd_csd_team_count = teamMembers.reduce((sum, m) => {
      const pl = teamPayloadByMemberId[m.id];
      return sum + Number(pl?.metrics?.closed_cssd_csd_count ?? 0);
    }, 0);
  }
  board._cachedMembers = cached;
  return board;
}

function updateTeamRollupHeader(boardOverride) {
  const qEl = document.getElementById("teamRollupQueueBacklog");
  const pEl = document.getElementById("teamRollupInProgress");
  const rEl = document.getElementById("teamRollupResolvedPeriod");
  const closedEl = document.getElementById("teamRollupClosedCssdCsd");
  const pipeEl = document.getElementById("teamPipelineBacklogCount");
  if (!qEl || !pEl || !rEl) return;
  if (boardOverride) {
    if (pipeEl) pipeEl.textContent = String(boardOverride.pipeline_backlog_count ?? "--");
    if (closedEl) closedEl.textContent = String(boardOverride.closed_cssd_csd_team_count ?? "--");
    qEl.textContent = String(boardOverride.queue_backlog_count ?? "--");
    pEl.textContent = String(boardOverride.in_progress_count ?? "--");
    rEl.textContent = String(boardOverride.resolved_in_period_count ?? "--");
    return;
  }
  if (!teamMembers.length) {
    if (pipeEl) pipeEl.textContent = "--";
    if (closedEl) closedEl.textContent = "--";
    qEl.textContent = "--";
    pEl.textContent = "--";
    rEl.textContent = "--";
    return;
  }
  const board = buildOpsBoardRollupFromCache();
  const cached = board._cachedMembers;
  if (pipeEl) {
    pipeEl.textContent = board.pipeline_backlog_count == null ? "--" : String(board.pipeline_backlog_count);
  }
  if (cached === 0) {
    if (closedEl) closedEl.textContent = "--";
    qEl.textContent = "--";
    pEl.textContent = "--";
    rEl.textContent = "--";
    return;
  }
  if (closedEl) {
    closedEl.textContent = board.closed_cssd_csd_team_count == null ? "--" : String(board.closed_cssd_csd_team_count);
  }
  qEl.textContent = String(board.queue_backlog_count);
  pEl.textContent = String(board.in_progress_count);
  rEl.textContent = String(board.resolved_in_period_count);
}

function teamBoardMetricsPayload() {
  const payload = teamFormToObject();
  delete payload.assignee_username;
  delete payload.member_name;
  payload.member_usernames = teamMembers.map((m) => m.username).filter(Boolean);
  return payload;
}

function applyPipelineBacklogFromResponse(data) {
  const count = data?.pipeline_backlog_count ?? data?.board_metrics?.pipeline_backlog_count;
  if (count == null) return false;
  latestBoardMetrics = { ...latestBoardMetrics, pipeline_backlog_count: Number(count) };
  latestPipelineBacklogLoaded = true;
  updateTeamRollupHeader();
  return true;
}

function applyBoardMetricsFromResponse(data, merge) {
  if (!data || !data.board_metrics) return false;
  latestBoardMetrics = merge
    ? { ...latestBoardMetrics, ...data.board_metrics }
    : data.board_metrics;
  if (data.board_metrics.pipeline_backlog_count != null) {
    latestPipelineBacklogLoaded = true;
  }
  if (data.board_metrics.closed_cssd_csd_team_count != null) {
    latestBoardMetricsLoaded = true;
  } else if (!merge) {
    latestBoardMetricsLoaded = true;
  }
  updateTeamRollupHeader();
  return true;
}

function ensureTeamBoardMetrics() {
  if (boardMetricsRefreshInFlight) return boardMetricsRefreshInFlight;
  boardMetricsRefreshInFlight = enqueueTeamJira(() => refreshTeamBoardMetrics()).finally(() => {
    boardMetricsRefreshInFlight = null;
  });
  return boardMetricsRefreshInFlight;
}

async function refreshPipelineBacklogCount() {
  const payload = teamBoardMetricsPayload();
  const pipeEl = document.getElementById("teamPipelineBacklogCount");
  if (pipeEl && !latestPipelineBacklogLoaded) pipeEl.textContent = "…";
  try {
    const res = await fetch("/run-pipeline-backlog-count", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(90000),
    });
    let data = {};
    try {
      data = await res.json();
    } catch (_) {
      console.warn("Pipeline backlog: invalid JSON response", res.status);
      if (pipeEl && !latestPipelineBacklogLoaded) pipeEl.textContent = "--";
      return false;
    }
    if (res.ok && applyPipelineBacklogFromResponse(data)) {
      if ((data.warnings || []).length) console.warn("Pipeline backlog:", data.warnings.join(" "));
      return true;
    }
    console.warn("Pipeline backlog failed:", data);
    if (pipeEl && !latestPipelineBacklogLoaded) pipeEl.textContent = "--";
    return false;
  } catch (e) {
    console.warn("Pipeline backlog error:", e);
    if (pipeEl && !latestPipelineBacklogLoaded) pipeEl.textContent = "--";
    return false;
  }
}

async function refreshTeamClosedBoardMetrics() {
  const payload = teamBoardMetricsPayload();
  payload.skip_pipeline = true;
  const closedEl = document.getElementById("teamRollupClosedCssdCsd");
  if (closedEl && !latestBoardMetricsLoaded) closedEl.textContent = "…";
  try {
    const res = await fetch("/run-team-board-metrics", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(300000),
    });
    let data = {};
    try {
      data = await res.json();
    } catch (_) {
      console.warn("Team closed metrics: invalid JSON response", res.status);
      if (closedEl && !latestBoardMetricsLoaded) closedEl.textContent = "--";
      return false;
    }
    if (res.ok && applyBoardMetricsFromResponse(data, true)) {
      latestBoardMetricsLoaded = true;
      updateTeamRollupHeader();
      if ((data.warnings || []).length) console.warn("Team closed metrics:", data.warnings.join(" "));
      return true;
    }
    console.warn("Team closed metrics failed:", data);
    if (closedEl && !latestBoardMetricsLoaded) closedEl.textContent = "--";
    return false;
  } catch (e) {
    console.warn("Team closed metrics error:", e);
    if (closedEl && !latestBoardMetricsLoaded) closedEl.textContent = "--";
    const statusEl = document.getElementById("teamStatusSummary");
    if (statusEl && e && e.name === "TimeoutError") {
      statusEl.textContent = "Team closed metrics timed out. Pipeline backlog above may still be valid — retry refresh.";
    }
    return false;
  }
}

async function refreshTeamBoardMetrics() {
  const pipeOk = await refreshPipelineBacklogCount();
  const closedOk = await refreshTeamClosedBoardMetrics();
  return pipeOk || closedOk;
}

async function fetchTeamPosture(payload) {
  const res = await fetch("/run-team-posture", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    signal: AbortSignal.timeout(600000),
  });
  let data;
  try {
    data = await res.json();
  } catch (_) {
    throw new Error(`Invalid server response (HTTP ${res.status})`);
  }
  if (!res.ok) {
    throw new Error((data && data.error) ? String(data.error) : "Request failed");
  }
  return data;
}

function formatTeamPostureStatus(payload) {
  const meta = payload.query_meta || {};
  const broadN = Number(meta.broad_issue_count ?? (payload.raw_rows || []).length);
  const assigneeN = Number(meta.owned_issue_count ?? meta.assignee_issue_count ?? 0);
  if (broadN === 0) {
    return [
      `Jira returned 0 tickets in the broad metrics query for ${(payload.member && payload.member.name) || "this member"}.`,
      "",
      "Broad JQL (used for posture metrics):",
      payload.broad_jql || "(not set)",
      "",
      "Assignee JQL:",
      payload.jql || "(not set)",
      "",
      "If Start Date/Time is set in Team settings, only tickets created on/after that time are included.",
      "Clear Start or use an earlier date (e.g. start of month) to include more tickets.",
    ].join("\\n");
  }
  const status = payload.status_distribution || {};
  const lines = Object.entries(status)
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .map(([k, v]) => `${k}: ${v}`);
  lines.push("", `${broadN} ticket(s) in team pool (${assigneeN} owned by this member).`);
  return lines.join("\\n");
}

function renderTeamPostureMetrics(payload) {
  const metrics = payload.metrics || {};
  const oldest = payload.oldest_open || {};
  const resolvedOwned = metrics.resolved_owned_count ?? metrics.resolved_count ?? 0;
  document.getElementById("teamResolvedOwnedCount").textContent = String(resolvedOwned);
  document.getElementById("teamResolvedContributedCount").textContent = String(metrics.resolved_contributed_count ?? 0);
  const closedEl = document.getElementById("teamClosedCssdCsdCount");
  if (closedEl) closedEl.textContent = String(metrics.closed_cssd_csd_count ?? 0);
  document.getElementById("teamResolvedLast8hCount").textContent = String(metrics.resolved_last_8h_count ?? 0);
  document.getElementById("teamOpenCount").textContent = String(metrics.assigned_open_count ?? 0);
  const qbEl = document.getElementById("teamQueueBacklogCount");
  const ipEl = document.getElementById("teamInProgressCount");
  const wsEl = document.getElementById("teamWorkedStatusLast8hCount");
  const wsoEl = document.getElementById("teamWorkedStatusOthersLast8hCount");
  if (qbEl) qbEl.textContent = String(metrics.queue_backlog_count ?? 0);
  if (ipEl) ipEl.textContent = String(metrics.in_progress_count ?? 0);
  if (wsEl) wsEl.textContent = String(metrics.worked_status_last_8h_count ?? 0);
  if (wsoEl) wsoEl.textContent = String(metrics.worked_status_last_8h_assigned_others_count ?? 0);
  document.getElementById("teamReopenedCount").textContent = String(metrics.reopened_count ?? 0);
  document.getElementById("teamWorkedOtherCount").textContent = String(metrics.worked_on_assigned_others_count ?? 0);
  document.getElementById("teamSlaBreachCount").textContent = String(metrics.sla_breach_count ?? 0);
  document.getElementById("teamSlaNearCount").textContent = String(metrics.open_near_sla_breach_8h_count ?? 0);
  document.getElementById("teamOldestTicket").textContent = oldest.issue_key || "N/A";
  document.getElementById("teamOldestAge").textContent = String(oldest.age_days ?? "--");

  document.getElementById("teamStatusSummary").textContent = formatTeamPostureStatus(payload);
  document.getElementById("teamOldestDetail").textContent = JSON.stringify(oldest || {}, null, 2);
  refreshTeamLabelCharts(payload.label_distribution || {});
  const previewRows = payload.raw_rows || [];
  if (previewRows.length) {
    renderTeamCsvPreview(previewRows);
  } else {
    document.getElementById("teamCsvPreview").textContent =
      "Dashboard refresh does not load ticket rows. Use Download CSV or Download Team CSV for Member Name, Dashboard Bucket, Issue Key, and Summary.";
  }
  updateTeamRollupHeader();
  updateTeamDataModeHint();
  if (!teamBulkRefreshInFlight && !payload._fromArchive) {
    void refreshOpsMetricTrendsForCurrentMode();
  }
}

const TEAM_LABELS_BAR_TOP_N = 15;
const TEAM_LABELS_TREND_TOP_N = 10;

function renderTeamLabelsBarChart(labelDistribution) {
  const el = document.getElementById("teamLabelsBarChart");
  if (!el) return;
  lastTeamLabelDistribution = labelDistribution || {};
  const entries = Object.entries(lastTeamLabelDistribution)
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .slice(0, TEAM_LABELS_BAR_TOP_N);
  const theme = getChartTheme();
  const wrap = el.closest(".legacy-chart-wrap");
  if (wrap) {
    const h = Math.min(480, Math.max(200, 80 + entries.length * 22));
    wrap.style.height = `${h}px`;
    wrap.style.maxHeight = "min(480px, 55vh)";
  }
  const ctx = el.getContext("2d");
  destroyChart(teamCharts.labelsBar);
  if (!entries.length) {
    teamCharts.labelsBar = null;
    return;
  }
  const colors = chartSliceColors(entries.length);
  teamCharts.labelsBar = new Chart(ctx, {
    type: "bar",
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{
        label: "Tickets",
        data: entries.map(([, v]) => Number(v)),
        backgroundColor: colors,
        borderColor: theme.panel,
        borderWidth: 1,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: chartThemePlugins(theme, { legend: { display: false } }),
      scales: {
        x: {
          ticks: { color: theme.muted, precision: 0 },
          grid: { color: theme.border },
          border: { color: theme.border },
          title: { display: true, text: "Ticket count", color: theme.muted },
        },
        y: {
          ticks: { color: theme.text, autoSkip: false, font: { size: 10 } },
          grid: { display: false },
          border: { color: theme.border },
        },
      },
    },
  });
  scheduleChartResize();
}

async function renderTeamLabelsTrendChart() {
  const el = document.getElementById("teamLabelsTrendChart");
  const hintEl = document.getElementById("teamLabelsHint");
  if (!el) return;
  const member = activeTeamMember();
  const theme = getChartTheme();
  destroyChart(teamCharts.labelsTrend);
  teamCharts.labelsTrend = null;

  if (!member || !member.username) {
    if (hintEl) hintEl.textContent = "Select a member and refresh to see label charts.";
    return;
  }

  let url = `/snapshots/label-trends?report_id=ops&member_username=${encodeURIComponent(member.username)}&top=${TEAM_LABELS_TREND_TOP_N}`;
  if (snapshotViewMode.team !== "live" && activeSnapshotId.ops) {
    url += `&to_snapshot_id=${encodeURIComponent(activeSnapshotId.ops)}`;
  }
  let series = { time_labels: [], datasets: [], snapshot_count: 0 };
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (res.ok) series = data;
  } catch (e) {
    if (hintEl) hintEl.textContent = `Label trend load failed: ${e && e.message ? e.message : String(e)}`;
    return;
  }

  const timeLabels = [...(series.time_labels || [])];
  const datasets = (series.datasets || []).map((ds) => ({
    label: ds.label,
    data: [...(ds.data || [])],
  }));

  if (snapshotViewMode.team === "live" && lastTeamLabelDistribution && Object.keys(lastTeamLabelDistribution).length) {
    timeLabels.push("Now");
    for (const ds of datasets) {
      ds.data.push(Number(lastTeamLabelDistribution[ds.label] || 0));
    }
  }

  if (!timeLabels.length) {
    if (hintEl) {
      hintEl.textContent = snapshotViewMode.team === "live"
        ? "No saved snapshots with label data for this member yet. Click Save snapshot after refresh to start tracking trends."
        : "No label history in this archived report for this member.";
    }
    return;
  }

  if (hintEl) {
    const pts = series.snapshot_count || timeLabels.length;
    hintEl.textContent = `Trend uses top ${TEAM_LABELS_TREND_TOP_N} labels across ${pts} saved report(s)${snapshotViewMode.team === "live" ? "; “Now” is the current refresh." : "."}`;
  }

  const ctx = el.getContext("2d");
  teamCharts.labelsTrend = new Chart(ctx, {
    type: "line",
    data: {
      labels: timeLabels,
      datasets: datasets.map((ds, i) => {
        const color = CHART_SLICE_PALETTE[i % CHART_SLICE_PALETTE.length];
        return {
          label: ds.label,
          data: ds.data,
          borderColor: color,
          backgroundColor: color + "33",
          tension: 0.25,
          fill: false,
          pointRadius: 3,
          pointHoverRadius: 5,
        };
      }),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: chartThemePlugins(theme, {
        legend: { position: "bottom", labels: { boxWidth: 10, padding: 8, font: { size: 10 } } },
      }),
      scales: chartThemeScales(theme),
      layout: { padding: { left: 4, right: 8, top: 4, bottom: 4 } },
    },
  });
  scheduleChartResize();
}

async function refreshTeamLabelCharts(labelDistribution) {
  renderTeamLabelsBarChart(labelDistribution);
  await renderTeamLabelsTrendChart();
  scheduleChartResize();
}

const LEGACY_LABELS_BAR_TOP_N = 15;
const LEGACY_LABELS_TREND_TOP_N = 10;

function renderLegacyLabelsBarChart(labelDistribution) {
  const el = document.getElementById("legacyLabelsBarChart");
  if (!el) return;
  lastLegacyLabelDistribution = labelDistribution || {};
  const entries = Object.entries(lastLegacyLabelDistribution)
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .slice(0, LEGACY_LABELS_BAR_TOP_N);
  const theme = getChartTheme();
  const wrap = el.closest(".legacy-chart-wrap");
  if (wrap) {
    const h = Math.min(480, Math.max(200, 80 + entries.length * 22));
    wrap.style.height = `${h}px`;
    wrap.style.maxHeight = "min(480px, 55vh)";
  }
  const ctx = el.getContext("2d");
  destroyChart(legacyCharts.labelsBar);
  if (!entries.length) {
    legacyCharts.labelsBar = null;
    const hintEl = document.getElementById("legacyLabelsHint");
    if (hintEl && !Object.keys(lastLegacyLabelDistribution || {}).length) {
      hintEl.textContent = snapshotViewMode.legacy === "live"
        ? "Refresh the dashboard to see label counts. Save snapshots to build the trend bar chart over time."
        : "No label data in this archived report.";
    }
    return;
  }
  const colors = chartSliceColors(entries.length);
  legacyCharts.labelsBar = new Chart(ctx, {
    type: "bar",
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{
        label: "Tickets",
        data: entries.map(([, v]) => Number(v)),
        backgroundColor: colors,
        borderColor: theme.panel,
        borderWidth: 1,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: chartThemePlugins(theme, { legend: { display: false } }),
      scales: {
        x: {
          ticks: { color: theme.muted, precision: 0 },
          grid: { color: theme.border },
          border: { color: theme.border },
          title: { display: true, text: "Ticket count", color: theme.muted },
        },
        y: {
          ticks: { color: theme.text, autoSkip: false, font: { size: 10 } },
          grid: { display: false },
          border: { color: theme.border },
        },
      },
    },
  });
  scheduleChartResize();
}

async function renderLegacyLabelsTrendChart() {
  const el = document.getElementById("legacyLabelsTrendChart");
  const hintEl = document.getElementById("legacyLabelsHint");
  if (!el) return;
  const theme = getChartTheme();
  destroyChart(legacyCharts.labelsTrend);
  legacyCharts.labelsTrend = null;

  let url = `/snapshots/label-trends?report_id=legacy&top=${LEGACY_LABELS_TREND_TOP_N}`;
  if (snapshotViewMode.legacy !== "live" && activeSnapshotId.legacy) {
    url += `&to_snapshot_id=${encodeURIComponent(activeSnapshotId.legacy)}`;
  }
  let series = { time_labels: [], datasets: [], snapshot_count: 0 };
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (res.ok) series = data;
  } catch (e) {
    if (hintEl) hintEl.textContent = `Label trend load failed: ${e && e.message ? e.message : String(e)}`;
    return;
  }

  const timeLabels = [...(series.time_labels || [])];
  const datasets = (series.datasets || []).map((ds) => ({
    label: ds.label,
    data: [...(ds.data || [])],
  }));

  if (snapshotViewMode.legacy === "live" && lastLegacyLabelDistribution && Object.keys(lastLegacyLabelDistribution).length) {
    timeLabels.push("Now");
    for (const ds of datasets) {
      ds.data.push(Number(lastLegacyLabelDistribution[ds.label] || 0));
    }
  }

  if (!timeLabels.length) {
    if (hintEl) {
      hintEl.textContent = snapshotViewMode.legacy === "live"
        ? "No saved Ticket trend snapshots with label data yet. Save snapshot after refresh to start tracking trends."
        : "No label history in this archived report.";
    }
    return;
  }

  if (hintEl) {
    const pts = series.snapshot_count || timeLabels.length;
    hintEl.textContent = `Label trend line chart: top ${LEGACY_LABELS_TREND_TOP_N} labels across ${pts} saved report(s)${snapshotViewMode.legacy === "live" ? "; “Now” is the current refresh." : "."}`;
  }

  const ctx = el.getContext("2d");
  legacyCharts.labelsTrend = new Chart(ctx, {
    type: "line",
    data: {
      labels: timeLabels,
      datasets: datasets.map((ds, i) => {
        const color = CHART_SLICE_PALETTE[i % CHART_SLICE_PALETTE.length];
        return {
          label: ds.label,
          data: ds.data,
          borderColor: color,
          backgroundColor: color + "33",
          tension: 0.25,
          fill: false,
          pointRadius: 3,
          pointHoverRadius: 5,
        };
      }),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: chartThemePlugins(theme, {
        legend: { position: "bottom", labels: { boxWidth: 10, padding: 8, font: { size: 10 } } },
      }),
      scales: chartThemeScales(theme),
      layout: { padding: { left: 4, right: 8, top: 4, bottom: 4 } },
    },
  });
  scheduleChartResize();
}

async function refreshLegacyLabelCharts(labelDistribution) {
  renderLegacyLabelsBarChart(labelDistribution);
  await renderLegacyLabelsTrendChart();
  scheduleChartResize();
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
  if (teamBulkRefreshInFlight) {
    const statusEl = document.getElementById("teamStatusSummary");
    if (statusEl) statusEl.textContent = "Bulk refresh in progress — wait for it to finish.";
    return;
  }
  ensureLiveModeForTeamRefresh();
  return enqueueTeamJira(async () => {
    updateTeamReportPeriodLabel();
    const member = activeTeamMember();
    if (!member) {
      document.getElementById("teamStatusSummary").textContent = "Add and select a member first.";
      document.getElementById("teamCsvPreview").textContent = "Add and select a member first.";
      return;
    }
    const payload = teamFormToObject();
    payload.fetch_board_metrics = false;
    payload.include_raw_rows = false;
    if (teamPoolCacheId) payload.pool_cache_id = teamPoolCacheId;
    const statusEl = document.getElementById("teamStatusSummary");
    if (statusEl) {
      statusEl.textContent = teamPoolCacheId
        ? `Refreshing ${member.name} from cached team pool…`
        : `Refreshing ${member.name} (loading team pool from Jira)…`;
    }
    try {
      const data = await fetchTeamPosture(payload);
      if (data.pool_cache_id) teamPoolCacheId = data.pool_cache_id;
      teamPayloadByMemberId[member.id] = data;
      if (!latestBoardMetricsLoaded) {
        await ensureTeamBoardMetrics();
      }
      selectTeamMember(member.id);
      updateTeamReportPeriodLabel();
    } catch (err) {
      const msg = `Network error calling /run-team-posture: ${err && err.message ? err.message : String(err)}. Wait a few seconds and retry.`;
      if (statusEl) statusEl.textContent = msg;
      document.getElementById("teamCsvPreview").textContent = msg;
    }
  });
}

function teamRefreshRequestBody() {
  const base = teamFormToObject();
  delete base.assignee_username;
  delete base.member_name;
  base.team_members = teamMembers.map((m) => ({
    id: m.id,
    name: m.name,
    username: m.username,
  }));
  return base;
}

function applyTeamRefreshResponse(data) {
  if (data.pool_cache_id) teamPoolCacheId = data.pool_cache_id;
  if (data.board_metrics) {
    latestBoardMetrics = { ...latestBoardMetrics, ...data.board_metrics };
    if (data.board_metrics.pipeline_backlog_count != null) latestPipelineBacklogLoaded = true;
    if (data.board_metrics.closed_cssd_csd_team_count != null) latestBoardMetricsLoaded = true;
  }
  const byUsername = {};
  for (const payload of data.members || []) {
    const uname = (payload.member && payload.member.assignee_username || "").toLowerCase();
    if (uname) byUsername[uname] = payload;
  }
  let successCount = 0;
  for (const member of teamMembers) {
    const payload = byUsername[(member.username || "").toLowerCase()];
    if (payload) {
      teamPayloadByMemberId[member.id] = payload;
      successCount += 1;
    }
  }
  return successCount;
}

async function refreshAllTeamMembers() {
  if (teamBulkRefreshInFlight) return teamBulkRefreshInFlight;
  teamBulkRefreshInFlight = (async () => {
    const statusEl = document.getElementById("teamStatusSummary");
    const previewEl = document.getElementById("teamCsvPreview");
    const refreshBtn = document.getElementById("teamRefreshBtn");
    const refreshBtnMain = document.getElementById("teamRefreshBtnMain");
    const refreshActiveBtn = document.getElementById("teamRefreshActiveBtn");
    let lastError = "";
    try {
      ensureLiveModeForTeamRefresh();
      updateTeamReportPeriodLabel();
      if (!teamMembers.length) {
        if (statusEl) statusEl.textContent = "Add and select a member first.";
        if (previewEl) previewEl.textContent = "Add and select a member first.";
        return;
      }
      if (refreshBtn) refreshBtn.disabled = true;
      if (refreshBtnMain) refreshBtnMain.disabled = true;
      if (refreshActiveBtn) refreshActiveBtn.disabled = true;
      if (statusEl) {
        statusEl.textContent = `Refreshing team (one Jira fetch for all ${teamMembers.length} members)…`;
      }
      latestPipelineBacklogLoaded = false;
      latestBoardMetricsLoaded = false;
      latestBoardMetrics = {};
      teamPoolCacheId = null;
      updateTeamRollupHeader();
      const body = teamRefreshRequestBody();
      const data = await enqueueTeamJira(async () => {
        const res = await fetch("/run-team-posture-refresh", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
          signal: AbortSignal.timeout(3600000),
        });
        const json = await res.json();
        if (!res.ok) {
          throw new Error((json && json.error) ? String(json.error) : "Team refresh failed");
        }
        return json;
      });
      const successCount = applyTeamRefreshResponse(data);
      if ((data.warnings || []).length) {
        console.warn("Team refresh warnings:", data.warnings.join(" "));
      }
      if (!activeTeamMemberId && teamMembers[0]) {
        activeTeamMemberId = teamMembers[0].id;
      }
      if (activeTeamMemberId) {
        selectTeamMember(activeTeamMemberId);
      }
      const poolN = data.query_meta && data.query_meta.broad_issue_count;
      if (statusEl) {
        statusEl.textContent = successCount === teamMembers.length
          ? `All ${successCount} member(s) refreshed (${poolN != null ? poolN + " tickets in pool" : "shared pool"}). Use Download Team CSV for ticket rows.`
          : `Loaded ${successCount}/${teamMembers.length} member(s).`;
      }
      updateTeamReportPeriodLabel();
      updateTeamRollupHeader();
      try {
        await refreshOpsMetricTrendsForCurrentMode();
      } catch (err) {
        console.warn("Metric trends refresh skipped:", err);
      }
    } catch (err) {
      lastError = err && err.message ? err.message : String(err);
      if (statusEl) statusEl.textContent = lastError || "Refresh stopped due to a connection error.";
      updateTeamRollupHeader();
      updateTeamDataModeHint();
    } finally {
      if (refreshBtn) refreshBtn.disabled = false;
      if (refreshBtnMain) refreshBtnMain.disabled = false;
      if (refreshActiveBtn) refreshActiveBtn.disabled = false;
      teamBulkRefreshInFlight = null;
      renderTeamMemberIcons();
      updateTeamDataModeHint();
    }
  })();
  return teamBulkRefreshInFlight;
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

const REPORT_ID_MAP = { csms: "exec", team: "ops", legacy: "legacy" };

function formatDeltaLine(delta) {
  if (!delta || delta.change === undefined) return "—";
  const ch = Number(delta.change);
  const arrow = ch > 0 ? "▲" : ch < 0 ? "▼" : "•";
  if (delta.pct_change != null && delta.pct_change !== "") {
    const pct = Number(delta.pct_change);
    const sign = pct > 0 ? "+" : "";
    return `${arrow} ${sign}${pct}%`;
  }
  if (ch === 0) return `${arrow} 0%`;
  return "—";
}

function deltaTone(metricKey, change) {
  const upBad = new Set([
    "backlog", "new_created", "pipeline_backlog_count", "queue_backlog_count",
    "assigned_open_count", "reopened_count", "sla_breach_count",
    "ttfr_cssd_median_hours", "ttfr_csd_median_hours", "ttr_cssd_median_hours", "ttr_csd_median_hours",
  ]);
  const upGood = new Set([
    "resolved", "resolved_in_period_count", "resolved_owned_count",
    "resolved_contributed_count", "closed_cssd_csd_count", "closed_cssd_csd_team_count",
  ]);
  const ch = Number(change);
  if (ch === 0) return "trend-pos";
  const up = ch > 0;
  if (upGood.has(metricKey)) return up ? "trend-pos" : "trend-neg";
  return (upBad.has(metricKey) ? !up : up) ? "trend-pos" : "trend-neg";
}

function setArchiveBanner(reportKey, visible, text) {
  const id = reportKey === "csms" ? "csmsArchiveBanner" : reportKey === "team" ? "teamArchiveBanner" : "legacyArchiveBanner";
  const el = document.getElementById(id);
  if (!el) return;
  el.hidden = !visible;
  el.textContent = text || "";
}

async function loadSnapshotOptions(reportId, selectEl) {
  if (!selectEl) return null;
  const res = await fetch(`/snapshots/list-options?report_id=${encodeURIComponent(reportId)}`);
  const data = await res.json();
  if (!res.ok) return null;
  const prev = selectEl.value;
  selectEl.innerHTML = "";
  const liveOpt = document.createElement("option");
  liveOpt.value = "live";
  liveOpt.textContent = "Live (refresh from Jira)";
  selectEl.appendChild(liveOpt);
  for (const opt of data.options || []) {
    const o = document.createElement("option");
    o.value = String(opt.id);
    o.textContent = opt.label;
    selectEl.appendChild(o);
  }
  if (prev && [...selectEl.options].some((o) => o.value === prev)) {
    selectEl.value = prev;
  } else if (data.latest_id) {
    selectEl.value = String(data.latest_id);
    activeSnapshotId[reportId] = data.latest_id;
    snapshotViewMode[reportId === "exec" ? "csms" : reportId === "ops" ? "team" : "legacy"] = "archive";
  } else {
    selectEl.value = "live";
    snapshotViewMode[reportId === "exec" ? "csms" : reportId === "ops" ? "team" : "legacy"] = "live";
  }
  const statusEl = selectEl.id === "csmsSnapshotSelect" ? document.getElementById("csmsSnapshotStatus")
    : selectEl.id === "teamSnapshotSelect" ? document.getElementById("teamSnapshotStatus")
    : document.getElementById("legacySnapshotStatus");
  if (statusEl) {
    const cadence = data.suggested_cadence ? `Suggested cadence: ${data.suggested_cadence}. ` : "";
    statusEl.textContent = data.options?.length
      ? `${cadence}${data.options.length} saved report(s). Load/Rerun saved settings, or Delete snapshot to remove one.`
      : `${cadence}No saved reports yet — run live and Save snapshot.`;
  }
  return data;
}

async function loadSnapshotDisplay(reportId, snapshotId) {
  const res = await fetch(`/snapshots/${snapshotId}/display`);
  const data = await res.json();
  if (!res.ok) return null;
  if (data.params && typeof data.params === "object") {
    snapshotParamsById[snapshotId] = data.params;
  }
  return data;
}

function hydrateCsmsFromDisplay(data) {
  document.getElementById("csmsSubtitle").textContent = `Executive Incident Summary | ${(data.periods && data.periods.period2 && data.periods.period2.label) || "saved report"}`;
  document.getElementById("csmsElapsed").textContent = data.elapsed_time_sentence || "Archived report.";
  renderCsmsKpis(data.kpis || {});
  renderCsmsNarratives({ narratives: data.narratives || {} });
  renderCsmsHealth(data.operational_health || {});
  document.getElementById("csmsStuck").textContent = JSON.stringify((data.kpis && data.kpis.longest_open) || {}, null, 2);
  renderCsmsCharts(data.charts || {});
}

function hydrateLegacyFromDisplay(data) {
  renderLegacyKpis(data.kpis || {});
  renderLegacyCharts(data.charts || {});
  renderLegacyStatusSummary(data.charts || {});
  void refreshLegacyLabelCharts((data.charts || {}).label_distribution || {});
  const lines = [...(data.warnings || []), ...(data.insights || [])];
  document.getElementById("legacyInsights").textContent = lines.join("\\n") || "Archived report.";
}

const LEGACY_SLA_METRIC_KEYS = [
  "ttfr_cssd_median_hours",
  "ttfr_csd_median_hours",
  "ttr_cssd_median_hours",
  "ttr_csd_median_hours",
];

async function applyDeltaToLegacySlaCard(card, snapshotId, liveKpis) {
  const sub = card.querySelector(".metric-trend-sub");
  const metricKey = card.getAttribute("data-metric-key");
  if (!sub || !metricKey) return;
  const valueEl = card.querySelector(".kpi-number");
  if (valueEl && (valueEl.textContent.trim() === "--" || !valueEl.textContent.trim())) {
    sub.textContent = "—";
    sub.className = "metric-trend-sub";
    sub.removeAttribute("title");
    return;
  }
  let deltas = [];
  let baselineSource = "prior saved report";
  if (snapshotId) {
    const res = await fetch(
      `/snapshots/compare?report_id=legacy&snapshot_id=${encodeURIComponent(snapshotId)}`
    );
    const cmp = await res.json();
    if (res.ok) {
      deltas = (cmp.deltas || []).filter((d) => LEGACY_SLA_METRIC_KEYS.includes(d.metric_key));
      baselineSource = (cmp.baseline && cmp.baseline.source === "manual") ? "manual baseline" : "prior saved report";
    }
  } else if (liveKpis) {
    const res = await fetch("/snapshots/compare-live", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ report_id: "legacy", metrics: { kpis: liveKpis } }),
    });
    const cmp = await res.json();
    if (res.ok) {
      deltas = (cmp.deltas || []).filter((d) => LEGACY_SLA_METRIC_KEYS.includes(d.metric_key));
      baselineSource = (cmp.baseline && cmp.baseline.source === "manual") ? "manual baseline" : "prior saved report";
    }
  }
  const d = deltas.find((x) => x.metric_key === metricKey);
  if (!d) {
    sub.textContent = "—";
    sub.className = "metric-trend-sub";
    sub.title = "No prior saved report with this metric (save a snapshot after refresh).";
    return;
  }
  sub.textContent = formatDeltaLine(d);
  sub.className = `metric-trend-sub ${deltaTone(metricKey, d.change)}`;
  const ch = Number(d.change);
  const pct = d.pct_change != null ? `${d.pct_change}%` : "";
  const hoursNote = ch > 0 ? "slower" : ch < 0 ? "faster" : "unchanged";
  sub.title = `${hoursNote} vs ${baselineSource}: ${ch > 0 ? "+" : ""}${ch.toFixed(1)}h${pct ? ` (${pct})` : ""}`;
}

async function refreshLegacySlaTrends(snapshotId, mode) {
  const cards = document.querySelectorAll("#legacyKpis .legacy-sla-kpi-card[data-metric-key]");
  const liveKpis = mode === "live" && latestLegacyPayload ? latestLegacyPayload.kpis : null;
  for (const card of cards) {
    try {
      if (mode === "archive" && snapshotId) {
        await applyDeltaToLegacySlaCard(card, snapshotId, null);
      } else if (mode === "live" && liveKpis) {
        await applyDeltaToLegacySlaCard(card, null, liveKpis);
      } else {
        const sub = card.querySelector(".metric-trend-sub");
        if (sub) {
          sub.textContent = "—";
          sub.className = "metric-trend-sub";
        }
      }
    } catch (err) {
      console.warn("Legacy SLA trend failed:", card.getAttribute("data-metric-key"), err);
    }
  }
}

function opsMemberPayloadFromArchive(memberView) {
  return {
    member: { name: memberView.name, assignee_username: memberView.username },
    metrics: memberView.metrics || {},
    status_distribution: memberView.status_distribution || {},
    label_distribution: memberView.label_distribution || {},
    oldest_open: memberView.oldest_open || {},
    raw_rows: [],
    _fromArchive: true,
  };
}

function hydrateOpsFromDisplay(data) {
  const board = data.board || {};
  latestBoardMetrics = {
    pipeline_backlog_count: board.pipeline_backlog_count,
    closed_cssd_csd_team_count: board.closed_cssd_csd_team_count,
    queue_backlog_count: board.queue_backlog_count,
    in_progress_count: board.in_progress_count,
    resolved_in_period_count: board.resolved_in_period_count,
  };
  latestBoardMetricsLoaded = true;
  latestPipelineBacklogLoaded = board.pipeline_backlog_count != null;
  board._archiveNote = data.note ? `Archived: ${data.note}` : "Archived official report (not live Jira).";
  updateTeamRollupHeader(board);
  const members = data.members || [];
  teamPayloadByMemberId = {};
  for (const memberView of members) {
    const uname = (memberView.username || "").toLowerCase();
    const local = teamMembers.find((m) => (m.username || "").toLowerCase() === uname);
    if (!local) continue;
    teamPayloadByMemberId[local.id] = opsMemberPayloadFromArchive(memberView);
  }
  if (members.length && !activeTeamMemberId) {
    activeTeamMemberId = teamMembers[0] ? teamMembers[0].id : null;
  }
  const activeMember = activeTeamMember();
  let memberView = null;
  if (activeMember) {
    memberView = members.find((m) => (m.username || "").toLowerCase() === (activeMember.username || "").toLowerCase());
  }
  updateTeamDataModeHint();
  if (activeTeamMemberId) {
    selectTeamMember(activeTeamMemberId);
  } else if (teamMembers[0]) {
    selectTeamMember(teamMembers[0].id);
  }
}

async function applySnapshotSelection(reportUiKey) {
  const reportId = REPORT_ID_MAP[reportUiKey];
  const selectId = reportUiKey === "csms" ? "csmsSnapshotSelect" : reportUiKey === "team" ? "teamSnapshotSelect" : "legacySnapshotSelect";
  const sel = document.getElementById(selectId);
  if (!sel) return;
  const val = sel.value;
  if (val === "live") {
    snapshotViewMode[reportUiKey] = "live";
    activeSnapshotId[reportId] = null;
    setArchiveBanner(reportUiKey, false, "");
    if (reportUiKey === "team") {
      latestBoardMetricsLoaded = false;
      latestPipelineBacklogLoaded = false;
      latestBoardMetrics = {};
      updateTeamRollupHeader();
      updateTeamDataModeHint();
      if (activeTeamMemberId) {
        selectTeamMember(activeTeamMemberId);
      }
    }
    return;
  }
  const snapId = parseInt(val, 10);
  if (!snapId) return;
  snapshotViewMode[reportUiKey] = "archive";
  activeSnapshotId[reportId] = snapId;
  const display = await loadSnapshotDisplay(reportId, snapId);
  if (!display) return;
  const bannerText = `Official saved report — ${display.captured_at || ""}${display.note ? " — " + display.note : ""} — not live Jira`;
  setArchiveBanner(reportUiKey, true, bannerText);
  if (reportUiKey === "csms") hydrateCsmsFromDisplay(display);
  if (reportUiKey === "legacy") {
    hydrateLegacyFromDisplay(display);
  }
  if (reportUiKey === "team") {
    hydrateOpsFromDisplay(display);
    await refreshOpsMetricTrends(snapId, "archive");
  }
}

async function initReportSnapshots(reportUiKey) {
  const reportId = REPORT_ID_MAP[reportUiKey];
  const selectId = reportUiKey === "csms" ? "csmsSnapshotSelect" : reportUiKey === "team" ? "teamSnapshotSelect" : "legacySnapshotSelect";
  const data = await loadSnapshotOptions(reportId, document.getElementById(selectId));
  if (data && data.latest_id) {
    await applySnapshotSelection(reportUiKey);
  } else {
    snapshotViewMode[reportUiKey] = "live";
    setArchiveBanner(reportUiKey, false, "");
  }
}

function clearSparklineWrap(wrap) {
  if (!wrap) return;
  const key = wrap.dataset.sparkKey;
  if (key) destroyChart(metricSparkCharts[key]);
  wrap.classList.remove("has-chart");
  wrap.innerHTML = "";
  delete wrap.dataset.sparkKey;
}

async function renderMetricSparkline(wrap, reportId, metricKey, memberUsername, toSnapshotId) {
  if (!wrap || typeof Chart === "undefined") return;
  const key = `${reportId}:${metricKey}:${memberUsername || "board"}`;
  wrap.dataset.sparkKey = key;
  destroyChart(metricSparkCharts[key]);
  clearSparklineWrap(wrap);

  let url = `/snapshots/trends?report_id=${encodeURIComponent(reportId)}&metric_key=${encodeURIComponent(metricKey)}`;
  if (memberUsername) url += `&member_username=${encodeURIComponent(memberUsername)}`;
  if (toSnapshotId) url += `&to_snapshot_id=${toSnapshotId}`;
  let res;
  let series;
  try {
    res = await fetch(url);
    series = await res.json();
  } catch (err) {
    console.warn("Sparkline fetch failed:", metricKey, err);
    metricSparkCharts[key] = null;
    return;
  }
  const points = series.data || [];
  if (!res.ok || points.length < 2) {
    metricSparkCharts[key] = null;
    return;
  }

  wrap.classList.add("has-chart");
  const canvas = document.createElement("canvas");
  canvas.className = "metric-sparkline";
  wrap.appendChild(canvas);
  const w = Math.max(wrap.clientWidth || 160, 80);
  canvas.width = w;
  canvas.height = 36;
  metricSparkCharts[key] = new Chart(canvas.getContext("2d"), {
    type: "line",
    data: {
      labels: series.labels || [],
      datasets: [{
        data: points,
        borderColor: "rgba(47, 122, 248, 0.9)",
        borderWidth: 1.5,
        pointRadius: 0,
        tension: 0.2,
        fill: false,
      }],
    },
    options: {
      responsive: false,
      animation: false,
      plugins: { legend: { display: false } },
      scales: { x: { display: false }, y: { display: false } },
    },
  });
}

async function applyDeltaToCard(card, reportId, metricKey, memberUsername, snapshotId, liveBoard) {
  const sub = card.querySelector(".metric-trend-sub");
  if (!sub) return;
  const valueEl = card.querySelector(".value");
  if (valueEl && valueEl.textContent.trim() === "--") {
    sub.textContent = "—";
    sub.className = "metric-trend-sub";
    sub.removeAttribute("title");
    return;
  }
  let deltas = [];
  let baselineSource = "prior report";
  if (snapshotId) {
    let url = `/snapshots/compare?report_id=${encodeURIComponent(reportId)}&snapshot_id=${snapshotId}`;
    if (memberUsername) url += `&member_username=${encodeURIComponent(memberUsername)}`;
    const res = await fetch(url);
    const cmp = await res.json();
    if (res.ok) {
      deltas = cmp.deltas || [];
      baselineSource = (cmp.baseline && cmp.baseline.source) || "prior report";
    }
  } else if (liveBoard) {
    const metricsBody = memberUsername
      ? { members: [{ username: memberUsername, metrics: liveBoard.memberMetrics || {} }] }
      : { board: liveBoard };
    const res = await fetch("/snapshots/compare-live", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ report_id: reportId, metrics: metricsBody, member_username: memberUsername || null }),
    });
    const cmp = await res.json();
    if (res.ok) {
      deltas = cmp.deltas || [];
      baselineSource = (cmp.baseline && cmp.baseline.source) || "prior report";
    }
  }
  const d = deltas.find((x) => x.metric_key === metricKey);
  if (!d) {
    sub.textContent = "—";
    sub.className = "metric-trend-sub";
    return;
  }
  d._source = baselineSource === "manual" ? "manual baseline" : "prior report";
  sub.textContent = formatDeltaLine(d);
  sub.className = `metric-trend-sub ${deltaTone(metricKey, d.change)}`;
  const ch = Number(d.change);
  const pct = d.pct_change != null ? `${d.pct_change}%` : "";
  sub.title = `Change ${ch > 0 ? "+" : ""}${ch}${pct ? ` (${pct})` : ""} vs ${d._source}`;
}

async function refreshOpsMetricTrends(snapshotId, mode) {
  const reportId = "ops";
  const member = activeTeamMember();
  const memberUsername = member ? member.username : null;
  const liveBoard = mode === "live" ? buildOpsBoardRollupFromCache() : null;
  if (mode === "live" && liveBoard) {
    liveBoard.memberMetrics = (latestTeamPosturePayload && latestTeamPosturePayload.metrics) || {};
  }
  const cards = document.querySelectorAll("#teamRollupGrid .team-metric-card[data-metric-key], #teamMetricsGrid .team-metric-card[data-metric-key]");
  const noTrendKeys = new Set([
    "worked_status_last_8h_count",
    "worked_status_last_8h_assigned_others_count",
    "resolved_last_8h_count",
  ]);
  for (const card of cards) {
    const metricKey = card.getAttribute("data-metric-key");
    const scope = card.getAttribute("data-metric-scope");
    const uname = scope === "member" ? memberUsername : null;
    if (metricKey && noTrendKeys.has(metricKey)) {
      const sub = card.querySelector(".metric-trend-sub");
      if (sub) {
        sub.textContent = "";
        sub.className = "metric-trend-sub";
        sub.removeAttribute("title");
      }
      continue;
    }
    if (scope === "member" && !uname) {
      const sub = card.querySelector(".metric-trend-sub");
      if (sub) { sub.textContent = "—"; sub.className = "metric-trend-sub"; }
      continue;
    }
    try {
      if (mode === "archive" && snapshotId) {
        await applyDeltaToCard(card, reportId, metricKey, uname, snapshotId, null);
      } else if (mode === "live") {
        await applyDeltaToCard(card, reportId, metricKey, uname, null, scope === "board" ? liveBoard : { memberMetrics: (latestTeamPosturePayload && latestTeamPosturePayload.metrics) || {} });
      }
    } catch (err) {
      console.warn("Metric card trend failed:", metricKey, err);
    }
  }
}

function buildOpsSavePayload() {
  const board = buildOpsBoardRollupFromCache();
  const members = [];
  for (const m of teamMembers) {
    const pl = teamPayloadByMemberId[m.id];
    if (!pl) continue;
    members.push({
      username: m.username,
      name: m.name,
      metrics: pl.metrics || {},
      status_distribution: pl.status_distribution || {},
      label_distribution: pl.label_distribution || {},
      oldest_open: pl.oldest_open || {},
    });
  }
  return { board, members };
}

async function saveReportSnapshot(reportUiKey) {
  const reportId = REPORT_ID_MAP[reportUiKey];
  const noteId = reportUiKey === "csms" ? "csmsSnapshotNote" : reportUiKey === "team" ? "teamSnapshotNote" : "legacySnapshotNote";
  const note = (document.getElementById(noteId) || {}).value || "";
  let metrics = null;
  let params = null;
  if (reportUiKey === "csms") {
    if (!latestCsmsPayload) { alert("Run the executive report first."); return; }
    params = csmsFormToObject(document.getElementById("csmsForm"));
    metrics = { view: {
      kpis: latestCsmsPayload.kpis,
      periods: latestCsmsPayload.periods,
      operational_health: latestCsmsPayload.operational_health,
      narratives: latestCsmsPayload.narratives,
      charts: latestCsmsPayload.charts,
      elapsed_time_sentence: latestCsmsPayload.elapsed_time_sentence,
    }, trend: {
      backlog: latestCsmsPayload.kpis?.backlog?.period2,
      new_created: latestCsmsPayload.kpis?.new_created?.period2,
      resolved: latestCsmsPayload.kpis?.resolved?.period2,
    }};
  } else if (reportUiKey === "legacy") {
    if (!latestLegacyPayload) { alert("Refresh the legacy dashboard first."); return; }
    params = formToObject(document.getElementById("exportForm"));
    metrics = { view: {
      kpis: latestLegacyPayload.kpis,
      charts: latestLegacyPayload.charts,
      insights: latestLegacyPayload.insights,
      warnings: latestLegacyPayload.warnings,
    }, trend: {
      issue_count: latestLegacyPayload.kpis?.issue_count,
      transition_count: latestLegacyPayload.kpis?.transition_count,
      comment_count: latestLegacyPayload.kpis?.comment_count,
      date_window_days: latestLegacyPayload.kpis?.date_window_days,
      ttfr_cssd_median_hours: latestLegacyPayload.kpis?.ttfr_cssd_median_hours,
      ttfr_csd_median_hours: latestLegacyPayload.kpis?.ttfr_csd_median_hours,
      ttr_cssd_median_hours: latestLegacyPayload.kpis?.ttr_cssd_median_hours,
      ttr_csd_median_hours: latestLegacyPayload.kpis?.ttr_csd_median_hours,
    }};
  } else {
    const savePayload = buildOpsSavePayload();
    if (!savePayload.members.length) { alert("Refresh team members before saving."); return; }
    params = teamFormToObject();
    delete params.assignee_username;
    delete params.member_name;
    params.team_members = teamMembers.map((m) => ({
      id: m.id,
      name: m.name,
      username: m.username,
    }));
    const boardTrend = {};
    for (const k of ["pipeline_backlog_count","closed_cssd_csd_team_count","queue_backlog_count","in_progress_count","resolved_in_period_count","sla_breach_count","open_near_sla_breach_8h_count"]) {
      if (savePayload.board[k] != null) boardTrend[k] = savePayload.board[k];
    }
    metrics = {
      view: { board: savePayload.board, members: savePayload.members },
      trend: {
        board: boardTrend,
        members: savePayload.members.map((m) => ({ username: m.username, name: m.name, metrics: m.metrics })),
      },
    };
  }
  const res = await fetch("/snapshots", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ report_id: reportId, metrics, params, note }),
  });
  const data = await res.json();
  if (!res.ok) {
    alert(data.error || "Save failed");
    return;
  }
  const selectId = reportUiKey === "csms" ? "csmsSnapshotSelect" : reportUiKey === "team" ? "teamSnapshotSelect" : "legacySnapshotSelect";
  await loadSnapshotOptions(reportId, document.getElementById(selectId));
  const sel = document.getElementById(selectId);
  if (sel && data.id) sel.value = String(data.id);
  await applySnapshotSelection(reportUiKey);
}

const REPORT_MOBILE_TITLES = {
  csms: "Executive Report",
  team: "Operations Team",
  legacy: "Ticket trend",
  notes: "Notes",
  auth: "Auth",
};

function setNavOpen(open) {
  const nav = document.getElementById("appNav");
  const overlay = document.getElementById("navOverlay");
  const toggle = document.getElementById("navMenuToggle");
  if (!nav || !toggle) return;
  nav.classList.toggle("is-open", open);
  if (overlay) {
    overlay.hidden = !open;
    overlay.setAttribute("aria-hidden", open ? "false" : "true");
  }
  toggle.setAttribute("aria-expanded", open ? "true" : "false");
  toggle.setAttribute("aria-label", open ? "Close navigation menu" : "Open navigation menu");
  toggle.textContent = open ? "✕" : "☰";
  toggle.title = open ? "Close menu" : "Menu";
  document.body.classList.toggle("nav-open", open);
}

function closeMobileNav() {
  setNavOpen(false);
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
  const mobileTitle = document.getElementById("appMobileTitle");
  if (mobileTitle) {
    mobileTitle.textContent = REPORT_MOBILE_TITLES[report] || "CSMS Reporting";
  }
  closeMobileNav();
  window.scrollTo({ top: 0, behavior: "smooth" });
  requestAnimationFrame(() => scheduleChartResize());
  if (report === "team") {
    updateTeamReportPeriodLabel();
    initReportSnapshots("team");
  }
  if (report === "legacy") {
    updateLegacyReportPeriodLabel();
    initReportSnapshots("legacy");
    const labelDist = latestLegacyPayload?.charts?.label_distribution
      || lastLegacyLabelDistribution
      || {};
    void refreshLegacyLabelCharts(labelDist);
    void refreshLegacySlaTrendsForCurrentMode();
  }
  if (report === "csms") initReportSnapshots("csms");
}

document.getElementById("navMenuToggle")?.addEventListener("click", () => {
  const nav = document.getElementById("appNav");
  setNavOpen(!nav?.classList.contains("is-open"));
});
document.getElementById("navOverlay")?.addEventListener("click", closeMobileNav);
window.addEventListener("resize", () => {
  if (window.innerWidth > 768) closeMobileNav();
  scheduleChartResize();
});
if (typeof ResizeObserver !== "undefined") {
  const chartResizeObserver = new ResizeObserver(() => scheduleChartResize());
  const wrapEl = document.querySelector(".wrap");
  if (wrapEl) chartResizeObserver.observe(wrapEl);
}
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeMobileNav();
});

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

const LEGACY_KPI_TITLES = [
  "Number of issues returned by your current legacy filters and caps.",
  "Total status changes counted across those issues.",
  "Total comments counted on those issues.",
  "Number of calendar days in this result set that have at least one created ticket.",
  "TTFR CSSD: rollup of per-ticket hours (aggregate chosen in settings).",
  "TTFR CSD: linked CSSD SLA when present, else CSD customfield_10318.",
  "TTR CSSD: customfield_10317 else resolutiondate − created, for tickets in TTR status gate.",
  "TTR CSD: customfield_10317 else resolutiondate − created, for tickets in TTR status gate.",
];

function legacySlaAggregateLabel(method) {
  const m = (method || "median").toLowerCase();
  if (m === "mean") return "mean";
  if (m === "p90") return "p90";
  return "median";
}

function formatLegacySlaHours(hours) {
  if (hours == null || Number.isNaN(Number(hours))) return "--";
  const h = Number(hours);
  if (h < 48) return `${h.toFixed(1)}h`;
  return `${(h / 24).toFixed(1)}d`;
}

function renderLegacyKpis(kpis) {
  const container = document.getElementById("legacyKpis");
  if (!container) return;
  const slaSub = (count, aggregate) => {
    const agg = legacySlaAggregateLabel(aggregate);
    const n = count != null && count > 0 ? `${count} ticket(s)` : "no matching tickets";
    return `${agg} · ${n}`;
  };
  const cards = [
    { title: "Issue Count", value: kpis.issue_count || 0, sub: null, metricKey: null },
    { title: "Status Transitions", value: kpis.transition_count || 0, sub: null, metricKey: null },
    { title: "Comment Volume", value: kpis.comment_count || 0, sub: null, metricKey: null },
    { title: "Date Window Days", value: kpis.date_window_days || 0, sub: null, metricKey: null },
    { title: "TTFR CSSD", value: formatLegacySlaHours(kpis.ttfr_cssd_median_hours), sub: slaSub(kpis.ttfr_cssd_count, kpis.ttfr_cssd_aggregate), metricKey: "ttfr_cssd_median_hours" },
    { title: "TTFR CSD", value: formatLegacySlaHours(kpis.ttfr_csd_median_hours), sub: slaSub(kpis.ttfr_csd_count, kpis.ttfr_csd_aggregate), metricKey: "ttfr_csd_median_hours" },
    { title: "TTR CSSD", value: formatLegacySlaHours(kpis.ttr_cssd_median_hours), sub: slaSub(kpis.ttr_cssd_count, kpis.ttr_cssd_aggregate), metricKey: "ttr_cssd_median_hours" },
    { title: "TTR CSD", value: formatLegacySlaHours(kpis.ttr_csd_median_hours), sub: slaSub(kpis.ttr_csd_count, kpis.ttr_csd_aggregate), metricKey: "ttr_csd_median_hours" },
  ];
  container.innerHTML = cards.map((card, i) => `
    <div class="kpi-card${card.metricKey ? " legacy-sla-kpi-card" : ""}"${card.metricKey ? ` data-metric-key="${card.metricKey}"` : ""} title="${LEGACY_KPI_TITLES[i] || ""}">
      <div class="kpi-label">${card.title}</div>
      <div class="kpi-number">${card.value}</div>
      ${card.sub ? `<div class="small" style="margin-top:4px;color:var(--muted);">${card.sub}</div>` : ""}
      ${card.metricKey ? `<div class="metric-trend-sub">—</div>` : ""}
    </div>
  `).join("");
  void refreshLegacySlaTrendsForCurrentMode();
}

function refreshLegacySlaTrendsForCurrentMode() {
  if (snapshotViewMode.legacy === "live") {
    return refreshLegacySlaTrends(null, "live");
  }
  if (activeSnapshotId.legacy) {
    return refreshLegacySlaTrends(activeSnapshotId.legacy, "archive");
  }
  return Promise.resolve();
}

function renderLegacyCharts(charts) {
  const dailyCtx = document.getElementById("legacyDailyChart").getContext("2d");
  const statusCtx = document.getElementById("legacyStatusChart").getContext("2d");
  const theme = getChartTheme();
  destroyChart(legacyCharts.status);
  destroyChart(legacyCharts.daily);

  const statusEntries = Object.entries(charts.status_distribution || {});
  legacyCharts.status = new Chart(statusCtx, {
    type: "doughnut",
    data: {
      labels: statusEntries.map(([k]) => k),
      datasets: [chartPieDataset(statusEntries.map(([, v]) => v), theme)],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: chartThemePlugins(theme, {
        legend: { position: "bottom" },
      }),
    },
  });

  const dailyColors = [
    "rgba(47, 122, 248, 0.9)",
    "rgba(236, 72, 153, 0.9)",
    "rgba(245, 158, 11, 0.9)",
  ];
  legacyCharts.daily = new Chart(dailyCtx, {
    type: "line",
    data: {
      labels: (charts.created_daily || {}).dates || [],
      datasets: [
        { label: "Created", data: (charts.created_daily || {}).created_counts || [], borderColor: dailyColors[0], backgroundColor: dailyColors[0] + "22", tension: 0.2, fill: false, pointRadius: 2 },
        { label: "Updated", data: (charts.created_daily || {}).updated_counts || [], borderColor: dailyColors[1], backgroundColor: dailyColors[1] + "22", tension: 0.2, fill: false, pointRadius: 2 },
        { label: "Resolved", data: (charts.created_daily || {}).resolved_counts || [], borderColor: dailyColors[2], backgroundColor: dailyColors[2] + "22", tension: 0.2, fill: false, pointRadius: 2 },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: chartThemePlugins(theme),
      scales: chartThemeScales(theme),
    },
  });
  scheduleChartResize();
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
  latestLegacyPayload = data;
  renderLegacyKpis(data.kpis || {});
  renderLegacyCharts(data.charts || {});
  renderLegacyStatusSummary(data.charts || {});
  await refreshLegacyLabelCharts((data.charts || {}).label_distribution || {});
  await refreshLegacySlaTrends(null, "live");
  const warningLines = data.warnings || [];
  const insightLines = data.insights || [];
  document.getElementById("legacyInsights").textContent = [...warningLines, ...insightLines].join("\\n");
  updateLegacyReportPeriodLabel();
}

function renderCsmsCharts(charts) {
  const dailyCtx = document.getElementById("dailyTrendChart").getContext("2d");
  const statusCtx = document.getElementById("statusChart").getContext("2d");
  const topCtx = document.getElementById("topCategoryChart").getContext("2d");
  const theme = getChartTheme();

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
      plugins: chartThemePlugins(theme, { title: { display: true, text: "Daily Ticket Trend by Component" } }),
      scales: chartThemeScales(theme),
    }
  });

  const statusEntries = Object.entries(charts.status_distribution || {});
  csmsCharts.status = new Chart(statusCtx, {
    type: "doughnut",
    data: {
      labels: statusEntries.map(([k]) => k),
      datasets: [chartPieDataset(statusEntries.map(([,v]) => v), theme)],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: chartThemePlugins(theme, { title: { display: false }, legend: { display: false } }),
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
      plugins: chartThemePlugins(theme, { title: { display: true, text: "Top 5 Issue Type or Component Trends" } }),
      scales: chartThemeScales(theme),
    }
  });
  scheduleChartResize();
}

document.getElementById("csmsForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  await runCsmsExecutiveReport();
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

async function onTeamRefreshAllClick() {
  try {
    await refreshAllTeamMembers();
  } catch (err) {
    console.error("Refresh all failed:", err);
    const statusEl = document.getElementById("teamStatusSummary");
    if (statusEl) {
      statusEl.textContent = `Refresh failed: ${err && err.message ? err.message : String(err)}`;
    }
    teamBulkRefreshInFlight = null;
    renderTeamMemberIcons();
  }
}

document.getElementById("teamRefreshBtn")?.addEventListener("click", onTeamRefreshAllClick);
document.getElementById("teamRefreshBtnMain")?.addEventListener("click", onTeamRefreshAllClick);
document.getElementById("teamRefreshActiveBtn")?.addEventListener("click", () => {
  void refreshTeamPosture();
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

async function downloadMemberExportCsv() {
  const member = activeTeamMember();
  const statusEl = document.getElementById("teamStatusSummary");
  if (!member) {
    if (statusEl) statusEl.textContent = "Select a member first.";
    return;
  }
  if (snapshotViewMode.team !== "live") {
    if (statusEl) statusEl.textContent = "Switch to Live and refresh before exporting ticket rows.";
    return;
  }
  if (statusEl) statusEl.textContent = `Building CSV for ${member.name}…`;
  const payload = teamFormToObject();
  delete payload.assignee_username;
  delete payload.member_name;
  payload.team_members = [{ name: member.name, username: member.username }];
  if (teamPoolCacheId) payload.pool_cache_id = teamPoolCacheId;
  try {
    const res = await fetch("/run-team-posture-board-export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(600000),
    });
    const data = await res.json();
    if (!res.ok) {
      if (statusEl) statusEl.textContent = data.error || "Export failed.";
      return;
    }
    if (data.pool_cache_id) teamPoolCacheId = data.pool_cache_id;
    if (Array.isArray(data.rows) && data.rows.length) {
      downloadTeamCsvBlob(data.rows);
      if (statusEl) statusEl.textContent = `CSV downloaded (${data.rows.length} rows) for ${member.name}.`;
      return;
    }
    if (data.exports && data.exports.csv) {
      window.open(data.exports.csv, "_blank");
      if (statusEl) statusEl.textContent = `CSV download started for ${member.name}.`;
    }
  } catch (err) {
    if (statusEl) statusEl.textContent = `Export error: ${err && err.message ? err.message : String(err)}`;
  }
}

document.getElementById("teamExportCsvBtn").addEventListener("click", () => { void downloadMemberExportCsv(); });
document.getElementById("teamExportExcelBtn").addEventListener("click", () => { void downloadMemberExportCsv(); });

function mergeTeamExportRowsFromCache() {
  const exportPayloads = [];
  const missingMembers = [];
  for (const member of teamMembers) {
    const cached = teamPayloadByMemberId[member.id];
    if (cached && Array.isArray(cached.raw_rows) && cached.raw_rows.length) {
      exportPayloads.push(cached);
    } else if (cached) {
      missingMembers.push({ member, reason: "no rows" });
    } else {
      missingMembers.push({ member, reason: "not cached" });
    }
  }
  const allRows = [];
  for (const payload of exportPayloads) {
    const memberName = payload?.member?.name ?? "";
    const assigneeUsername = payload?.member?.assignee_username ?? "";
    for (const raw of payload.raw_rows) {
      allRows.push({
        "Member Name": memberName,
        "Assignee Username": assigneeUsername,
        ...raw,
      });
    }
  }
  return { allRows, missingMembers };
}

function downloadTeamCsvBlob(allRows) {
  const slimHeaders = ["Member Name", "Dashboard Bucket", "Issue Key", "Summary"];
  const headers = allRows[0] && slimHeaders.every((h) => h in allRows[0])
    ? slimHeaders
    : Object.keys(allRows[0]);
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
}

async function downloadTeamBoardCsv() {
  const statusEl = document.getElementById("teamStatusSummary");
  if (!teamMembers.length) {
    if (statusEl) statusEl.textContent = "Add team members first.";
    return;
  }

  if (snapshotViewMode.team !== "live") {
    const { allRows: archiveRows } = mergeTeamExportRowsFromCache();
    if (!archiveRows.length) {
      if (statusEl) {
        statusEl.textContent =
          "Archived reports do not include ticket rows. Use Load saved settings → Rerun with saved settings (or switch to Live and Refresh All), then Download Team CSV.";
      }
      return;
    }
  }

  const { allRows: cachedRows, missingMembers } = mergeTeamExportRowsFromCache();
  const allMembersCached = missingMembers.length === 0 && cachedRows.length > 0;
  if (allMembersCached) {
    downloadTeamCsvBlob(cachedRows);
    if (statusEl) {
      statusEl.textContent = `Team CSV downloaded (${cachedRows.length} rows) from session cache.`;
    }
    return;
  }

  if (statusEl) {
    statusEl.textContent = teamPoolCacheId
      ? "Building team CSV from cached issue pool…"
      : "Building team CSV (one Jira fetch for all members)…";
  }
  const payload = teamFormToObject();
  delete payload.assignee_username;
  delete payload.member_name;
  payload.team_members = teamMembers.map((m) => ({
    name: m.name,
    username: m.username,
  }));
  if (teamPoolCacheId) payload.pool_cache_id = teamPoolCacheId;
  try {
    const res = await fetch("/run-team-posture-board-export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (data.pool_cache_id) teamPoolCacheId = data.pool_cache_id;
    if (!res.ok) {
      if (statusEl) statusEl.textContent = data.error || data.details || "Team export failed.";
      if (cachedRows.length) {
        downloadTeamCsvBlob(cachedRows);
        if (statusEl) {
          statusEl.textContent += ` Partial cache download: ${cachedRows.length} row(s).`;
        }
      }
      return;
    }
    if (data.exports && data.exports.csv) {
      window.open(data.exports.csv, "_blank");
      const n = Array.isArray(data.rows) ? data.rows.length : "";
      if (statusEl) {
        statusEl.textContent = n
          ? `Team CSV ready from Jira (${n} rows).`
          : "Team CSV download started.";
      }
      return;
    }
    if (Array.isArray(data.rows) && data.rows.length) {
      downloadTeamCsvBlob(data.rows);
      if (statusEl) statusEl.textContent = `Team CSV downloaded (${data.rows.length} rows) from Jira.`;
      return;
    }
    if (statusEl) statusEl.textContent = "Export returned no rows for the current filters and roster.";
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    if (statusEl) statusEl.textContent = `Team export error: ${msg}`;
    if (cachedRows.length) {
      downloadTeamCsvBlob(cachedRows);
      if (statusEl) statusEl.textContent += ` Downloaded partial cache (${cachedRows.length} rows).`;
    }
  }
}

document.getElementById("teamExportAllBtn").addEventListener("click", () => {
  void downloadTeamBoardCsv();
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
updateTeamDataModeHint();

renderCsmsKpis({
  backlog: { period2: "--", trend: 0 },
  new_created: { period2: "--", trend: 0 },
  resolved: { period2: "--", trend: 0 },
  longest_open: { age_days: "--", issue_key: "" },
});
document.getElementById("csmsElapsed").textContent = "Provide Last Report Timestamp and run CSMS refresh to compute elapsed time.";
renderCsmsHealth({ process_alignment_pct: 60, process_gap_identified: "Pending run" });
renderLegacyKpis({ issue_count: 0, transition_count: 0, comment_count: 0, date_window_days: 0 });

["csmsSnapshotSelect", "teamSnapshotSelect", "legacySnapshotSelect"].forEach((id) => {
  const el = document.getElementById(id);
  if (!el) return;
  const key = id.startsWith("csms") ? "csms" : id.startsWith("team") ? "team" : "legacy";
  el.addEventListener("change", () => applySnapshotSelection(key));
});
document.getElementById("csmsSaveSnapshotBtn")?.addEventListener("click", () => saveReportSnapshot("csms"));
document.getElementById("teamSaveSnapshotBtn")?.addEventListener("click", () => saveReportSnapshot("team"));
document.getElementById("legacySaveSnapshotBtn")?.addEventListener("click", () => saveReportSnapshot("legacy"));
document.getElementById("csmsLoadSnapshotParamsBtn")?.addEventListener("click", () => loadSavedReportSettings("csms"));
document.getElementById("teamLoadSnapshotParamsBtn")?.addEventListener("click", () => loadSavedReportSettings("team"));
document.getElementById("legacyLoadSnapshotParamsBtn")?.addEventListener("click", () => loadSavedReportSettings("legacy"));
document.getElementById("csmsRerunSnapshotBtn")?.addEventListener("click", () => rerunWithSavedReportSettings("csms"));
document.getElementById("teamRerunSnapshotBtn")?.addEventListener("click", () => rerunWithSavedReportSettings("team"));
document.getElementById("legacyRerunSnapshotBtn")?.addEventListener("click", () => rerunWithSavedReportSettings("legacy"));
document.getElementById("csmsDeleteSnapshotBtn")?.addEventListener("click", () => deleteSelectedSnapshot("csms"));
document.getElementById("teamDeleteSnapshotBtn")?.addEventListener("click", () => deleteSelectedSnapshot("team"));
document.getElementById("legacyDeleteSnapshotBtn")?.addEventListener("click", () => deleteSelectedSnapshot("legacy"));
document.getElementById("teamBaselineSaveBtn")?.addEventListener("click", async () => {
  const metric_key = (document.getElementById("teamBaselineMetric")?.value || "").trim();
  const value = document.getElementById("teamBaselineValue")?.value;
  if (!metric_key || value === "") { alert("Enter metric_key and value."); return; }
  const member = activeTeamMember();
  const scope = document.querySelector(`#teamMetricsGrid .team-metric-card[data-metric-key="${metric_key}"]`);
  const isMember = scope && scope.getAttribute("data-metric-scope") === "member";
  await fetch("/manual-baselines", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      report_id: "ops",
      metric_key,
      value: Number(value),
      member_username: isMember && member ? member.username : null,
    }),
  });
  if (snapshotViewMode.team === "live") await refreshOpsMetricTrends(null, "live");
  else if (activeSnapshotId.ops) await refreshOpsMetricTrends(activeSnapshotId.ops, "archive");
});

initReportSnapshots("csms");
applyTheme(localStorage.getItem("theme") || "dark");
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

    if not clauses:
        raise ValueError(
            "JQL requires at least one filter: projects, issue types, statuses, assignees, "
            "labels, start/end dates, extra JQL, or a custom JQL override."
        )
    return " AND ".join(clauses) + " ORDER BY created DESC"


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
    fields: Optional[str] = None,
) -> List[Dict[str, Any]]:
    return fetch_issues(
        base_url,
        jql,
        page_size,
        max_issues,
        verify_ssl,
        include_changelog=include_changelog,
        fields=fields,
    )


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


def parse_sla_elapsed_hours(fields: Dict[str, Any], field_key: str) -> Optional[float]:
    """Parse Jira SLA completedCycles elapsed time to hours."""
    data = fields.get(field_key) or {}
    cycles = data.get("completedCycles") or []
    if not cycles:
        return None
    cycle = cycles[0]
    elapsed = cycle.get("elapsedTime") or {}
    millis = elapsed.get("millis")
    if millis is not None:
        try:
            return float(millis) / 3600000.0
        except (TypeError, ValueError):
            pass
    return None


def parse_legacy_status_gate(value: Optional[str], default_when_empty: str) -> List[str]:
    """Comma-separated status names; blank uses default_when_empty."""
    text = (value or "").strip()
    if not text:
        text = default_when_empty
    return [s.strip().lower() for s in parse_csv_list(text) if s.strip()]


def issue_status_matches_gate(issue: Dict[str, Any], allowed_statuses: List[str]) -> bool:
    if not allowed_statuses:
        return True
    status = (get_issue_status(issue) or "").strip().lower()
    return status in allowed_statuses


def ttfr_hours_from_fields(fields: Dict[str, Any]) -> Optional[float]:
    hours = parse_sla_elapsed_hours(fields, JIRA_SLA_TTFR_FIELD)
    if hours is not None:
        return hours
    _, _, stop = extract_sla(fields, JIRA_SLA_TTFR_FIELD)
    if not stop:
        return None
    created_raw = fields.get("created") or ""
    stop_dt = parse_jira_datetime(stop)
    created_dt = parse_jira_datetime(created_raw)
    if stop_dt and created_dt:
        return (stop_dt - created_dt).total_seconds() / 3600.0
    return None


def ttr_hours_from_issue(issue: Dict[str, Any]) -> Optional[float]:
    fields = issue.get("fields") or {}
    hours = parse_sla_elapsed_hours(fields, JIRA_SLA_TTR_FIELD)
    if hours is not None:
        return hours
    created_dt = get_issue_created_datetime(issue)
    resolved_dt = parse_jira_datetime(fields.get("resolutiondate") or "")
    if created_dt and resolved_dt:
        return (resolved_dt - created_dt).total_seconds() / 3600.0
    return None


def find_linked_cssd_key(issue: Dict[str, Any]) -> Optional[str]:
    for link in (issue.get("fields") or {}).get("issuelinks") or []:
        for side in ("outwardIssue", "inwardIssue"):
            other = link.get(side) or {}
            key = (other.get("key") or "").strip().upper()
            if key.startswith("CSSD-"):
                return key
    return None


def fetch_issues_by_keys(
    base_url: str,
    keys: List[str],
    verify_ssl: bool,
    fields: str = LEGACY_SLA_SEARCH_FIELDS,
) -> Dict[str, Dict[str, Any]]:
    """Batch-fetch issues by key; returns map key -> issue."""
    unique = sorted({k.strip().upper() for k in keys if (k or "").strip()})
    if not unique:
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    chunk_size = 50
    for i in range(0, len(unique), chunk_size):
        chunk = unique[i : i + chunk_size]
        jql = "key in (" + ", ".join(jql_quote(k) for k in chunk) + ")"
        batch = fetch_issues(
            base_url,
            jql,
            page_size=min(50, len(chunk)),
            max_issues=len(chunk),
            verify_ssl=verify_ssl,
            include_changelog=False,
            fields=fields,
        )
        for issue in batch:
            key = (issue.get("key") or "").strip().upper()
            if key:
                out[key] = issue
    return out


def median_hours(values: List[float]) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    n = len(ordered)
    mid = n // 2
    if n % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2.0


def parse_legacy_aggregate(value: Optional[str], default: str = DEFAULT_LEGACY_SLA_AGGREGATE) -> str:
    chosen = (value or default).strip().lower()
    if chosen in LEGACY_SLA_AGGREGATE_OPTIONS:
        return chosen
    return DEFAULT_LEGACY_SLA_AGGREGATE


def percentile_hours(values: List[float], percentile: float) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    rank = (percentile / 100.0) * (len(ordered) - 1)
    low = int(rank)
    high = min(low + 1, len(ordered) - 1)
    frac = rank - low
    return ordered[low] * (1.0 - frac) + ordered[high] * frac


def aggregate_hours(values: List[float], method: str) -> Optional[float]:
    if not values:
        return None
    agg = parse_legacy_aggregate(method)
    if agg == "mean":
        return sum(values) / len(values)
    if agg == "p90":
        return percentile_hours(values, 90.0)
    return median_hours(values)


def legacy_aggregate_display_name(method: str) -> str:
    agg = parse_legacy_aggregate(method)
    if agg == "mean":
        return "mean"
    if agg == "p90":
        return "90th percentile"
    return "median"


def compute_legacy_sla_kpis(
    issues: List[Dict[str, Any]],
    params: Dict[str, Any],
    base_url: str,
    verify_ssl: bool,
) -> Tuple[Dict[str, Any], List[str]]:
    """TTFR/TTR rollups for CSSD and CSD using configurable status gates and aggregates."""
    warnings: List[str] = []
    ttr_cssd_gate = parse_legacy_status_gate(
        params.get("ttr_status_cssd"), DEFAULT_LEGACY_TTR_STATUS_CSSD
    )
    ttr_csd_gate = parse_legacy_status_gate(
        params.get("ttr_status_csd"), DEFAULT_LEGACY_TTR_STATUS_CSD
    )
    ttfr_cssd_gate = parse_legacy_status_gate(params.get("ttfr_status_cssd"), "")
    ttfr_csd_gate = parse_legacy_status_gate(params.get("ttfr_status_csd"), "")
    ttfr_cssd_agg = parse_legacy_aggregate(params.get("ttfr_cssd_aggregate"))
    ttfr_csd_agg = parse_legacy_aggregate(params.get("ttfr_csd_aggregate"))
    ttr_cssd_agg = parse_legacy_aggregate(params.get("ttr_cssd_aggregate"))
    ttr_csd_agg = parse_legacy_aggregate(params.get("ttr_csd_aggregate"))

    csd_needing_cssd: List[str] = []
    csd_issues: List[Dict[str, Any]] = []
    for issue in issues:
        if get_issue_project_key(issue) != "CSD":
            continue
        csd_issues.append(issue)
        cssd_key = find_linked_cssd_key(issue)
        if cssd_key:
            csd_needing_cssd.append(cssd_key)

    cssd_by_key = fetch_issues_by_keys(base_url, csd_needing_cssd, verify_ssl)

    ttfr_cssd_hours: List[float] = []
    ttfr_csd_hours: List[float] = []
    ttr_cssd_hours: List[float] = []
    ttr_csd_hours: List[float] = []
    ttfr_csd_from_link = 0

    for issue in issues:
        project = get_issue_project_key(issue)
        fields = issue.get("fields") or {}

        if project == "CSSD":
            if issue_status_matches_gate(issue, ttfr_cssd_gate):
                hours = ttfr_hours_from_fields(fields)
                if hours is not None:
                    ttfr_cssd_hours.append(hours)
            if issue_status_matches_gate(issue, ttr_cssd_gate):
                hours = ttr_hours_from_issue(issue)
                if hours is not None:
                    ttr_cssd_hours.append(hours)

        elif project == "CSD":
            if issue_status_matches_gate(issue, ttfr_csd_gate):
                hours = None
                cssd_key = find_linked_cssd_key(issue)
                if cssd_key and cssd_key in cssd_by_key:
                    hours = ttfr_hours_from_fields(cssd_by_key[cssd_key].get("fields") or {})
                    if hours is not None:
                        ttfr_csd_from_link += 1
                if hours is None:
                    hours = ttfr_hours_from_fields(fields)
                if hours is not None:
                    ttfr_csd_hours.append(hours)
            if issue_status_matches_gate(issue, ttr_csd_gate):
                hours = ttr_hours_from_issue(issue)
                if hours is not None:
                    ttr_csd_hours.append(hours)

    if csd_issues and not ttfr_csd_hours and csd_needing_cssd:
        warnings.append(
            "No CSD TTFR values: check linked CSSD tickets and TTFR CSD status gate."
        )
    if not ttfr_cssd_hours:
        warnings.append(
            "No CSSD TTFR values in this result set (status gate or missing customfield_10318)."
        )

    kpis = {
        "ttfr_cssd_median_hours": aggregate_hours(ttfr_cssd_hours, ttfr_cssd_agg),
        "ttfr_cssd_aggregate": ttfr_cssd_agg,
        "ttfr_cssd_count": len(ttfr_cssd_hours),
        "ttfr_csd_median_hours": aggregate_hours(ttfr_csd_hours, ttfr_csd_agg),
        "ttfr_csd_aggregate": ttfr_csd_agg,
        "ttfr_csd_count": len(ttfr_csd_hours),
        "ttfr_csd_from_linked_cssd_count": ttfr_csd_from_link,
        "ttr_cssd_median_hours": aggregate_hours(ttr_cssd_hours, ttr_cssd_agg),
        "ttr_cssd_aggregate": ttr_cssd_agg,
        "ttr_cssd_count": len(ttr_cssd_hours),
        "ttr_csd_median_hours": aggregate_hours(ttr_csd_hours, ttr_csd_agg),
        "ttr_csd_aggregate": ttr_csd_agg,
        "ttr_csd_count": len(ttr_csd_hours),
        "sla_status_gates": {
            "ttr_cssd": ttr_cssd_gate,
            "ttr_csd": ttr_csd_gate,
            "ttfr_cssd": ttfr_cssd_gate or ["(any with TTFR SLA)"],
            "ttfr_csd": ttfr_csd_gate or ["(linked CSSD or own SLA)"],
        },
        "sla_aggregates": {
            "ttfr_cssd": ttfr_cssd_agg,
            "ttfr_csd": ttfr_csd_agg,
            "ttr_cssd": ttr_cssd_agg,
            "ttr_csd": ttr_csd_agg,
        },
    }
    return kpis, warnings


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
    fields: Optional[str] = None,
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
        if fields:
            payload["fields"] = fields
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


def count_jira_issues(
    base_url: str,
    jql: str,
    verify_ssl: bool,
    timeout: int = 120,
) -> int:
    """Return Jira search `total` without downloading issue pages (maxResults=0)."""
    auth = get_auth()
    session = requests.Session()
    resp = session.get(
        base_url,
        params={"jql": jql, "startAt": 0, "maxResults": 0},
        auth=auth,
        verify=verify_ssl,
        timeout=timeout,
    )
    resp.raise_for_status()
    return int(resp.json().get("total", 0))


def team_board_issue_fetch_cap(max_issues: int) -> int:
    """Cap board-level issue scans when the form sends 0 (unlimited)."""
    if max_issues and max_issues > 0:
        return max_issues
    return 500


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

    try:
        issues = fetch_issues(base_url, jql, page_size, max_issues, verify_ssl, include_changelog=True)
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            issues = fetch_issues(
                base_url, jql, page_size, max_issues, verify_ssl, include_changelog=False
            )
        else:
            raise

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
    search_fields = LEGACY_SLA_SEARCH_FIELDS
    try:
        issues = fetch_jira_issues(
            base_url,
            jql,
            page_size,
            max_issues,
            verify_ssl,
            include_changelog=True,
            fields=search_fields,
        )
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            issues = fetch_jira_issues(
                base_url,
                jql,
                page_size,
                max_issues,
                verify_ssl,
                include_changelog=False,
                fields=search_fields,
            )
            warnings.append(
                "Jira returned 500 with changelog expansion; loaded legacy dashboard without changelog details."
            )
        else:
            raise

    sla_kpis, sla_warnings = compute_legacy_sla_kpis(issues, params, base_url, verify_ssl)
    warnings.extend(sla_warnings)

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
    label_distribution = group_by_labels(issues)
    top_components = [{"name": k, "count": v} for k, v in Counter(group_by_component(issues)).most_common(5)]
    insights = [
        f"Workload distribution is led by {top_components[0]['name']} ({top_components[0]['count']} issues)." if top_components else "No component concentration detected.",
        f"Average transitions per issue: {(total_transitions / max(1, len(issues))):.2f}.",
        f"Most common current status: {max(status_counts, key=status_counts.get) if status_counts else 'N/A'}.",
    ]
    def _sla_insight(label: str, hours_key: str, count_key: str, agg_key: str, extra: str = "") -> str:
        count = sla_kpis.get(count_key) or 0
        agg_name = legacy_aggregate_display_name(sla_kpis.get(agg_key) or DEFAULT_LEGACY_SLA_AGGREGATE)
        hours = sla_kpis.get(hours_key)
        suffix = f" ({count} tickets{extra})."
        if hours is None:
            return f"{label} ({agg_name} n/a{suffix}"
        return f"{label} {agg_name}: {hours:.1f}h{suffix}"

    if sla_kpis.get("ttfr_cssd_count"):
        insights.append(_sla_insight("TTFR CSSD", "ttfr_cssd_median_hours", "ttfr_cssd_count", "ttfr_cssd_aggregate"))
    if sla_kpis.get("ttfr_csd_count"):
        link_n = sla_kpis.get("ttfr_csd_from_linked_cssd_count") or 0
        extra = f"; {link_n} from linked CSSD" if link_n else ""
        insights.append(_sla_insight("TTFR CSD", "ttfr_csd_median_hours", "ttfr_csd_count", "ttfr_csd_aggregate", extra))
    if sla_kpis.get("ttr_cssd_count"):
        insights.append(_sla_insight("TTR CSSD", "ttr_cssd_median_hours", "ttr_cssd_count", "ttr_cssd_aggregate"))
    if sla_kpis.get("ttr_csd_count"):
        insights.append(_sla_insight("TTR CSD", "ttr_csd_median_hours", "ttr_csd_count", "ttr_csd_aggregate"))
    insights = [line for line in insights if line]

    all_days = sorted(set(created_daily.keys()) | set(updated_daily.keys()) | set(resolved_daily.keys()))
    kpis: Dict[str, Any] = {
        "issue_count": len(issues),
        "transition_count": total_transitions,
        "comment_count": total_comments,
        "date_window_days": len(created_daily),
    }
    kpis.update(sla_kpis)
    return {
        "jql": jql,
        "warnings": warnings,
        "kpis": kpis,
        "charts": {
            "status_distribution": status_counts,
            "label_distribution": label_distribution,
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


def build_team_posture_broad_jql(params: Dict[str, Any]) -> str:
    """Team-wide JQL (no assignee): one fetch for all roster members."""
    return build_team_posture_jql(params, "", include_assignee=False)


def _store_team_issue_pool(pool: Dict[str, Any]) -> str:
    while len(TEAM_ISSUE_POOL_CACHE) >= TEAM_ISSUE_POOL_CACHE_MAX:
        oldest_key = next(iter(TEAM_ISSUE_POOL_CACHE))
        del TEAM_ISSUE_POOL_CACHE[oldest_key]
    cache_id = uuid.uuid4().hex
    TEAM_ISSUE_POOL_CACHE[cache_id] = pool
    return cache_id


def fetch_team_issue_pool(params: Dict[str, Any]) -> Dict[str, Any]:
    """Download all issues for the team report window once (with changelog when possible)."""
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")
    page_size = int(params.get("page_size") or 50)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    broad_jql = build_team_posture_broad_jql(params)
    warnings: List[str] = []
    changelog_included = True
    try:
        broad_issues = fetch_jira_issues(
            base_url, broad_jql, page_size, max_issues, verify_ssl, include_changelog=True
        )
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            broad_issues = fetch_jira_issues(
                base_url, broad_jql, page_size, max_issues, verify_ssl, include_changelog=False
            )
            changelog_included = False
            warnings.append(
                "Jira returned 500 with changelog expansion; worked-on/reopened metrics may be incomplete."
            )
        else:
            raise
    cap = team_board_issue_fetch_cap(max_issues)
    if max_issues and max_issues > 0 and len(broad_issues) >= max_issues:
        warnings.append(f"Broad team query capped at {max_issues} issues (max_issues).")
    elif not max_issues and cap and len(broad_issues) >= cap:
        warnings.append(f"Broad team query capped at {cap} issues.")
    return {
        "issues": broad_issues,
        "broad_jql": broad_jql,
        "warnings": warnings,
        "changelog_included": changelog_included,
        "issue_count": len(broad_issues),
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


def build_team_board_jql(params: Dict[str, Any]) -> str:
    """Board-level JQL for team closed metrics: project/issue type (no created window)."""
    projects = parse_csv_list(params.get("projects"))
    issue_types = parse_csv_list(params.get("issue_types"))
    clauses: List[str] = []
    for clause in [list_clause("project", projects), list_clause("issuetype", issue_types)]:
        if clause:
            clauses.append(clause)
    return (" AND ".join(clauses) if clauses else "order by created desc") + " ORDER BY created DESC"


def build_team_closed_board_jql(params: Dict[str, Any]) -> str:
    """Narrow JQL for TEAM CLOSED: CSSD/CSD in Closed status, optional report updated window."""
    clauses: List[str] = [
        list_clause("project", list(TEAM_OPS_CLOSED_PROJECT_KEYS)) or 'project in (CSSD, CSD)',
        'status = Closed',
    ]
    start_dt = normalize_dt_local(params.get("start_dt"))
    end_dt = normalize_dt_local(params.get("end_dt"))
    if start_dt:
        clauses.append(f'updated >= "{start_dt}"')
    if end_dt:
        clauses.append(f'updated <= "{end_dt}"')
    return " AND ".join(clauses) + " ORDER BY updated DESC"


def build_pipeline_backlog_jql(params: Dict[str, Any]) -> str:
    """Pipeline Backlog: official CSMS Prod filter (Jira UI parity) unless overridden in settings."""
    custom = (params.get("pipeline_backlog_jql") or "").strip()
    if custom:
        return custom
    created_since = (params.get("pipeline_backlog_created_since") or TEAM_PIPELINE_BACKLOG_CREATED_SINCE_DEFAULT).strip()
    if "T" in created_since:
        created_since = created_since.split("T", 1)[0]
    return TEAM_PIPELINE_BACKLOG_JQL_DEFAULT.format(created_since=created_since)


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


def is_team_ops_closed_issue(issue: Dict[str, Any], _project_rules: Optional[Dict[str, str]] = None) -> bool:
    """CSSD/CSD tickets whose Jira status is Closed (not RFPU or other done states)."""
    project_key = get_issue_project_key(issue)
    if project_key not in TEAM_OPS_CLOSED_PROJECT_KEYS:
        return False
    status = (get_issue_status(issue) or "").strip().lower()
    return status == "closed"


def count_closed_cssd_csd_combined_for_member(
    issues: List[Dict[str, Any]],
    assignee_username: str,
    csd_assigned_dev_field: str,
    project_rules: Dict[str, str],
) -> int:
    """CSSD/CSD in closed/done status where member owns the ticket or contributed a status change."""
    target = (assignee_username or "").strip().lower()
    if not target:
        return 0
    count = 0
    for issue in issues:
        if get_issue_project_key(issue) not in TEAM_OPS_CLOSED_PROJECT_KEYS:
            continue
        if not is_team_ops_closed_issue(issue, project_rules):
            continue
        owned = issue_owner_username(issue, csd_assigned_dev_field) == target
        if owned:
            count += 1
            continue
        if member_has_status_change(issue, target):
            count += 1
    return count


def count_closed_cssd_csd_team_deduped(
    issues: List[Dict[str, Any]],
    member_usernames: Iterable[str],
    csd_assigned_dev_field: str,
    project_rules: Dict[str, str],
) -> int:
    """Unique CSSD/CSD closed/done tickets any roster member owns or contributed to."""
    targets = {(u or "").strip().lower() for u in member_usernames if (u or "").strip()}
    if not targets:
        return 0
    seen: set = set()
    for issue in issues:
        if get_issue_project_key(issue) not in TEAM_OPS_CLOSED_PROJECT_KEYS:
            continue
        if not is_team_ops_closed_issue(issue, project_rules):
            continue
        issue_key = issue.get("key")
        if not issue_key or issue_key in seen:
            continue
        for target in targets:
            owned = issue_owner_username(issue, csd_assigned_dev_field) == target
            if owned or member_has_status_change(issue, target):
                seen.add(issue_key)
                break
    return len(seen)


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


def is_owned_assigned_open_cssd_csd(issue: Dict[str, Any], project_rules: Dict[str, str]) -> bool:
    """Owned CSSD/CSD still open: CSSD not Resolved/Closed; CSD not Ready For Production Users."""
    del project_rules
    project_key = get_issue_project_key(issue)
    if project_key not in TEAM_OPS_CLOSED_PROJECT_KEYS:
        return False
    status = (get_issue_status(issue) or "").strip().lower()
    if not status:
        return False
    if project_key == "CSSD":
        if status == "closed" or status == "resolved":
            return False
        if "resolved" in status:
            return False
        return True
    if project_key == "CSD":
        return status != "ready for production users"
    return False


def count_owned_assigned_open_cssd_csd(owned_issues: List[Dict[str, Any]], project_rules: Dict[str, str]) -> int:
    return sum(1 for issue in owned_issues if is_owned_assigned_open_cssd_csd(issue, project_rules))


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


def is_unassigned_issue(issue: Dict[str, Any]) -> bool:
    return not issue_current_assignee_username(issue)


def count_pipeline_backlog_from_jira(issues: List[Dict[str, Any]]) -> int:
    """Pipeline backlog count = issues returned by the dedicated Pipeline Backlog JQL."""
    return len(issues)


def build_pipeline_backlog_count_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    """Fast pipeline backlog: Jira search total only (no issue pages)."""
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")
    verify_ssl = bool(params.get("verify_ssl", True))
    pipeline_jql = build_pipeline_backlog_jql(params)
    warnings: List[str] = []
    try:
        pipeline_count = count_jira_issues(base_url, pipeline_jql, verify_ssl)
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            pipeline_count = count_jira_issues(base_url, pipeline_jql, verify_ssl)
            warnings.append("Pipeline backlog loaded after Jira 500 retry.")
        else:
            raise
    print(
        f"[pipeline-backlog-count] count={pipeline_count} jql={pipeline_jql!r}",
        flush=True,
    )
    return {
        "pipeline_backlog_count": pipeline_count,
        "pipeline_backlog_jql": pipeline_jql,
        "jql": pipeline_jql,
        "warnings": warnings,
        "board_metrics": {"pipeline_backlog_count": pipeline_count},
    }


def build_team_closed_board_metrics_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    """Team closed rollup (may fetch capped issue pages with changelog)."""
    base_url = (params.get("base_url") or "").strip()
    if not base_url:
        raise ValueError("base_url is required")
    page_size = int(params.get("page_size") or 50)
    max_issues = int(params.get("max_issues") or 0)
    verify_ssl = bool(params.get("verify_ssl", True))
    project_rules = {"CSSD": "Closed", "CSD": "Ready For Production Users"}
    closed_jql = build_team_closed_board_jql(params)
    closed_fetch_cap = team_board_issue_fetch_cap(max_issues)
    warnings: List[str] = []
    closed_issues: List[Dict[str, Any]] = []
    try:
        closed_issues = fetch_jira_issues(
            base_url,
            closed_jql,
            page_size,
            closed_fetch_cap,
            verify_ssl,
            include_changelog=True,
        )
    except requests.HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else None
        if status_code == 500:
            closed_issues = fetch_jira_issues(
                base_url,
                closed_jql,
                page_size,
                closed_fetch_cap,
                verify_ssl,
                include_changelog=True,
            )
            warnings.append("Team closed metrics loaded after Jira 500 retry.")
        else:
            raise
    if closed_fetch_cap and len(closed_issues) >= closed_fetch_cap:
        warnings.append(
            f"Team closed count uses at most {closed_fetch_cap} most recently updated "
            "CSSD/CSD Closed issues (changelog attribution)."
        )
    csd_assigned_dev_field = (params.get("csd_assigned_dev_field") or "customfield_14700").strip()
    member_usernames = params.get("member_usernames") or []
    if isinstance(member_usernames, str):
        member_usernames = parse_csv_list(member_usernames)
    board_metrics: Dict[str, Any] = {}
    if member_usernames:
        board_metrics["closed_cssd_csd_team_count"] = count_closed_cssd_csd_team_deduped(
            closed_issues, member_usernames, csd_assigned_dev_field, project_rules
        )
    return {
        "closed_team_jql": closed_jql,
        "warnings": warnings,
        "board_metrics": board_metrics,
    }


def build_team_board_metrics_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    skip_pipeline = bool(params.get("skip_pipeline"))
    skip_closed = bool(params.get("skip_closed"))
    warnings: List[str] = []
    board_metrics: Dict[str, Any] = {}
    pipeline_jql = ""
    closed_jql = ""

    if not skip_pipeline:
        pipe_payload = build_pipeline_backlog_count_payload(params)
        pipeline_jql = pipe_payload.get("pipeline_backlog_jql") or ""
        warnings.extend(pipe_payload.get("warnings") or [])
        board_metrics.update(pipe_payload.get("board_metrics") or {})

    if not skip_closed:
        closed_payload = build_team_closed_board_metrics_payload(params)
        closed_jql = closed_payload.get("closed_team_jql") or ""
        warnings.extend(closed_payload.get("warnings") or [])
        board_metrics.update(closed_payload.get("board_metrics") or {})

    return {
        "jql": pipeline_jql,
        "pipeline_backlog_jql": pipeline_jql,
        "closed_team_jql": closed_jql,
        "warnings": warnings,
        "board_metrics": board_metrics,
    }


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


def count_worked_status_last_8h_assigned_to_others(
    broad_issues: List[Dict[str, Any]],
    assignee_username: str,
    csd_assigned_dev_field: str,
    hours: float = 8.0,
) -> int:
    """Tickets not owned by the member where they authored a status changelog entry in the last `hours`."""
    target = (assignee_username or "").strip().lower()
    if not target:
        return 0
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=hours)
    n = 0
    for issue in broad_issues:
        if issue_owner_username(issue, csd_assigned_dev_field) == target:
            continue
        if issue_has_member_status_change_after(issue, assignee_username, cutoff):
            n += 1
    return n


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
    *,
    slim: bool = False,
    member_name: str = "",
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
    tags_by_key: Dict[str, set] = {}

    # Assigned open: CSSD/CSD only (not Closed / not Ready For Production Users).
    for issue in owned_issues:
        if not is_owned_assigned_open_cssd_csd(issue, project_rules):
            continue
        key = str(issue.get("key") or "").strip()
        if not key:
            continue
        project_key = get_issue_project_key(issue)
        tag = "assigned_open_cssd" if project_key == "CSSD" else "assigned_open_csd"
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

    # One row per ticket per dashboard bucket.
    rows: List[Dict[str, Any]] = []
    for key in sorted(tags_by_key.keys()):
        issue = by_key.get(key)
        if not issue:
            continue
        fields = issue.get("fields") or {}
        for tag in sorted(tags_by_key[key]):
            if slim:
                rows.append(
                    {
                        "Member Name": member_name,
                        "Dashboard Bucket": tag,
                        "Issue Key": key,
                        "Summary": fields.get("summary") or "",
                    }
                )
            else:
                raw = issue_to_raw_row(issue)
                rows.append({"Dashboard Bucket": tag, **raw})
    return rows


def build_team_member_posture_from_pool(
    broad_issues: List[Dict[str, Any]],
    params: Dict[str, Any],
    assignee_username: str,
    member_name: str,
    *,
    broad_jql: str = "",
    pool_warnings: Optional[List[str]] = None,
    include_raw_rows: bool = False,
    slim_export: bool = True,
) -> Dict[str, Any]:
    """Compute dashboard metrics for one member from a shared issue pool (no Jira fetch)."""
    base_url = (params.get("base_url") or "").strip()
    verify_ssl = bool(params.get("verify_ssl", True))
    csd_assigned_dev_field = (params.get("csd_assigned_dev_field") or "").strip()
    project_rules = {"CSSD": "Closed", "CSD": "Ready For Production Users"}
    jql = build_team_posture_jql(params, assignee_username, include_assignee=True)
    if not broad_jql:
        broad_jql = build_team_posture_broad_jql(params)
    warnings: List[str] = list(pool_warnings or [])

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
    open_count = count_owned_assigned_open_cssd_csd(owned_issues, project_rules)
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
    worked_status_last_8h_assigned_others_count = count_worked_status_last_8h_assigned_to_others(
        broad_issues, assignee_username, csd_assigned_dev_field, hours=8.0
    )
    resolved_in_period_count = count_owned_resolved_in_report_window(
        owned_issues,
        TEAM_POSTURE_RESOLVED_STATUS_KEYWORDS,
        report_start,
        report_end,
    )
    closed_cssd_csd_count = count_closed_cssd_csd_combined_for_member(
        broad_issues, assignee_username, csd_assigned_dev_field, project_rules
    )
    raw_rows: List[Dict[str, Any]] = []
    if include_raw_rows:
        raw_rows = build_member_dashboard_tagged_rows(
            broad_issues,
            assignee_username,
            csd_assigned_dev_field,
            project_rules,
            sla_hours=24.0,
            slim=slim_export,
            member_name=member_name,
        )

    return {
        "member": {"name": member_name, "assignee_username": assignee_username},
        "jql": jql,
        "broad_jql": broad_jql,
        "query_meta": {
            "assignee_issue_count": len(owned_issues),
            "owned_issue_count": len(owned_issues),
            "broad_issue_count": len(broad_issues),
        },
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
            "worked_status_last_8h_assigned_others_count": worked_status_last_8h_assigned_others_count,
            "resolved_in_period_count": resolved_in_period_count,
            "closed_cssd_csd_count": closed_cssd_csd_count,
            "sla_breach_count": sla_metrics["sla_breach_count"],
            "open_near_sla_breach_8h_count": sla_metrics["open_near_sla_breach_8h_count"],
        },
        "status_distribution": status_distribution,
        "label_distribution": label_distribution,
        "oldest_open": oldest_open,
        "raw_rows": raw_rows,
    }


def build_team_posture_payload(params: Dict[str, Any]) -> Dict[str, Any]:
    """Single-member posture: uses shared pool cache or fetches the team pool once."""
    assignee_username = (params.get("assignee_username") or "").strip()
    if not assignee_username:
        raise ValueError("assignee_username is required")
    member_name = (params.get("member_name") or assignee_username).strip()
    include_raw_rows = bool(params.get("include_raw_rows"))
    slim_export = params.get("slim_export", True) is not False

    pool_cache_id = (params.get("pool_cache_id") or "").strip()
    pool: Optional[Dict[str, Any]] = None
    if pool_cache_id and pool_cache_id in TEAM_ISSUE_POOL_CACHE:
        pool = TEAM_ISSUE_POOL_CACHE[pool_cache_id]
    if pool is None:
        pool = fetch_team_issue_pool(params)
        pool_cache_id = _store_team_issue_pool(pool)

    payload = build_team_member_posture_from_pool(
        pool["issues"],
        params,
        assignee_username,
        member_name,
        broad_jql=pool.get("broad_jql") or "",
        pool_warnings=pool.get("warnings") or [],
        include_raw_rows=include_raw_rows,
        slim_export=slim_export,
    )
    payload["pool_cache_id"] = pool_cache_id
    return payload


def build_team_posture_refresh_payload(
    params: Dict[str, Any], members: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Fetch team issue pool once, compute dashboard metrics for every roster member."""
    if not members:
        raise ValueError("team_members is required")
    print("[run-team-posture-refresh] fetching team issue pool", flush=True)
    pool = fetch_team_issue_pool(params)
    pool_cache_id = _store_team_issue_pool(pool)
    broad_jql = pool.get("broad_jql") or ""
    warnings: List[str] = list(pool.get("warnings") or [])

    member_payloads: List[Dict[str, Any]] = []
    for member in members:
        username = (member.get("username") or "").strip()
        if not username:
            continue
        name = (member.get("name") or username).strip()
        print(f"[run-team-posture-refresh] metrics member={name}", flush=True)
        member_payloads.append(
            build_team_member_posture_from_pool(
                pool["issues"],
                params,
                username,
                name,
                broad_jql=broad_jql,
                pool_warnings=warnings,
                include_raw_rows=False,
            )
        )

    board_metrics: Dict[str, Any] = {}
    pipeline_jql = ""
    closed_jql = ""
    try:
        board_params = dict(params)
        board_params["member_usernames"] = [
            (m.get("username") or "").strip()
            for m in members
            if (m.get("username") or "").strip()
        ]
        board_payload = build_team_board_metrics_payload(board_params)
        board_metrics = board_payload.get("board_metrics") or {}
        pipeline_jql = board_payload.get("pipeline_backlog_jql") or ""
        closed_jql = board_payload.get("closed_team_jql") or ""
        warnings.extend(board_payload.get("warnings") or [])
    except Exception as board_exc:
        warnings.append(f"Board metrics: {board_exc}")

    return {
        "pool_cache_id": pool_cache_id,
        "broad_jql": broad_jql,
        "query_meta": {
            "broad_issue_count": pool.get("issue_count", 0),
            "member_count": len(member_payloads),
            "changelog_included": pool.get("changelog_included", True),
        },
        "warnings": warnings,
        "members": member_payloads,
        "board_metrics": board_metrics,
        "pipeline_backlog_jql": pipeline_jql,
        "closed_team_jql": closed_jql,
    }


def build_team_board_export_rows(
    params: Dict[str, Any],
    members: List[Dict[str, Any]],
    pool: Optional[Dict[str, Any]] = None,
    *,
    slim: bool = True,
) -> List[Dict[str, Any]]:
    """Build export rows from a shared issue pool (one Jira fetch unless pool provided)."""
    if pool is None:
        pool = fetch_team_issue_pool(params)
    broad_issues = pool.get("issues") or []
    csd_assigned_dev_field = (params.get("csd_assigned_dev_field") or "").strip()
    project_rules = {"CSSD": "Closed", "CSD": "Ready For Production Users"}
    rows: List[Dict[str, Any]] = []
    for member in members:
        username = (member.get("username") or "").strip()
        if not username:
            continue
        name = (member.get("name") or username).strip()
        member_rows = build_member_dashboard_tagged_rows(
            broad_issues,
            username,
            csd_assigned_dev_field,
            project_rules,
            slim=slim,
            member_name=name,
        )
        rows.extend(member_rows)
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
    try:
        params = request.get_json(force=True)
        return jsonify({"jql": build_jql(params)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


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


@app.route("/run-team-posture-refresh", methods=["POST"])
def run_team_posture_refresh():
    try:
        params = request.get_json(force=True)
        members = params.get("team_members") or []
        if not isinstance(members, list) or not members:
            return jsonify({"error": "team_members is required"}), 400
        print(f"[run-team-posture-refresh] start roster={len(members)}", flush=True)
        payload = build_team_posture_refresh_payload(params, members)
        print("[run-team-posture-refresh] done", flush=True)
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        print(f"[run-team-posture-refresh] error: {exc}", flush=True)
        return jsonify({"error": str(exc)}), 400


@app.route("/run-team-posture", methods=["POST"])
def run_team_posture():
    try:
        params = request.get_json(force=True)
        member_label = (params.get("member_name") or params.get("assignee_username") or "?").strip()
        include_raw_rows = bool(params.get("include_raw_rows"))
        print(f"[run-team-posture] start member={member_label} raw_rows={include_raw_rows}", flush=True)
        payload = build_team_posture_payload(params)
        print(f"[run-team-posture] done member={member_label}", flush=True)

        if include_raw_rows:
            out_dir = Path(tempfile.mkdtemp(prefix="team_posture_"))
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_member = re.sub(r"[^A-Za-z0-9_-]+", "_", payload["member"]["name"])[:40] or "member"
            csv_path = out_dir / f"team_posture_{safe_member}_{stamp}.csv"
            summary_csv = out_dir / f"team_posture_summary_{safe_member}_{stamp}.csv"
            excel_path = out_dir / f"team_posture_{safe_member}_{stamp}.xlsx"

            raw_rows = payload.get("raw_rows") or []
            headers = list(TEAM_EXPORT_SLIM_COLUMNS) if raw_rows else ["Issue Key"]
            if raw_rows and "Member Name" not in raw_rows[0]:
                headers = list(raw_rows[0].keys())
            write_csv(csv_path, raw_rows, headers)
            mets = payload.get("metrics") or {}
            summary_rows = [
                {"Metric": "Resolved (Owned)", "Value": mets.get("resolved_owned_count", mets.get("resolved_count", 0))},
                {"Metric": "Resolved (Contributed)", "Value": mets.get("resolved_contributed_count", 0)},
                {"Metric": "Resolved (Last 8 Hours)", "Value": mets.get("resolved_last_8h_count", 0)},
                {"Metric": "Queue Backlog", "Value": mets.get("queue_backlog_count", 0)},
                {"Metric": "In Progress", "Value": mets.get("in_progress_count", 0)},
                {"Metric": "Worked Status (Last 8 Hours)", "Value": mets.get("worked_status_last_8h_count", 0)},
                {"Metric": "Worked Status (Others, Last 8 Hours)", "Value": mets.get("worked_status_last_8h_assigned_others_count", 0)},
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
        if params.get("fetch_board_metrics"):
            try:
                board_params = dict(params)
                board_params["member_usernames"] = params.get("member_usernames") or []
                if not board_params["member_usernames"]:
                    roster = params.get("team_members") or []
                    if isinstance(roster, list):
                        board_params["member_usernames"] = [
                            (m.get("username") or "").strip()
                            for m in roster
                            if (m.get("username") or "").strip()
                        ]
                board_payload = build_team_board_metrics_payload(board_params)
                payload["board_metrics"] = board_payload.get("board_metrics")
                payload["pipeline_backlog_jql"] = board_payload.get("pipeline_backlog_jql")
                if board_payload.get("warnings"):
                    payload["warnings"] = list(payload.get("warnings") or []) + board_payload["warnings"]
            except Exception as board_exc:
                payload["board_metrics_error"] = str(board_exc)
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        print(f"[run-team-posture] error: {exc}", flush=True)
        return jsonify({"error": str(exc)}), 400


@app.route("/run-team-posture-board-export", methods=["POST"])
def run_team_posture_board_export():
    try:
        params = request.get_json(force=True)
        members = params.get("team_members") or []
        if not isinstance(members, list) or not members:
            return jsonify({"error": "team_members is required"}), 400
        pool_cache_id = (params.get("pool_cache_id") or "").strip()
        pool: Optional[Dict[str, Any]] = None
        if pool_cache_id and pool_cache_id in TEAM_ISSUE_POOL_CACHE:
            pool = TEAM_ISSUE_POOL_CACHE[pool_cache_id]
            print(f"[run-team-posture-board-export] using pool cache {pool_cache_id}", flush=True)
        else:
            print("[run-team-posture-board-export] fetching team issue pool", flush=True)
            pool = fetch_team_issue_pool(params)
            pool_cache_id = _store_team_issue_pool(pool)
        out_dir = Path(tempfile.mkdtemp(prefix="team_posture_board_"))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = out_dir / f"team_posture_board_{stamp}.csv"
        rows = build_team_board_export_rows(params, members, pool=pool, slim=True)
        if not rows:
            return jsonify({"error": "No export rows for the current filters and roster"}), 400
        headers = list(TEAM_EXPORT_SLIM_COLUMNS)
        write_csv(csv_path, rows, headers)
        export_id = uuid.uuid4().hex
        TEAM_EXPORT_CACHE[export_id] = {
            "csv": str(csv_path),
        }
        return jsonify(
            {
                "rows": rows,
                "pool_cache_id": pool_cache_id,
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


@app.route("/run-pipeline-backlog-count", methods=["POST"])
def run_pipeline_backlog_count():
    try:
        print("[run-pipeline-backlog-count] start", flush=True)
        payload = build_pipeline_backlog_count_payload(request.get_json(force=True))
        print("[run-pipeline-backlog-count] done", flush=True)
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        print(f"[run-pipeline-backlog-count] HTTP error: {exc}", flush=True)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        print(f"[run-pipeline-backlog-count] error: {exc}", flush=True)
        return jsonify({"error": str(exc)}), 400


@app.route("/run-team-board-metrics", methods=["POST"])
def run_team_board_metrics():
    try:
        print("[run-team-board-metrics] start", flush=True)
        payload = build_team_board_metrics_payload(request.get_json(force=True))
        print("[run-team-board-metrics] done", flush=True)
        return jsonify(payload)
    except requests.HTTPError as exc:
        details = exc.response.text if exc.response is not None else str(exc)
        print(f"[run-team-board-metrics] HTTP error: {exc}", flush=True)
        return jsonify({"error": f"HTTP error: {exc}", "details": details}), 400
    except Exception as exc:
        print(f"[run-team-board-metrics] error: {exc}", flush=True)
        return jsonify({"error": str(exc)}), 400


@app.route("/snapshots/list-options", methods=["GET"])
def snapshots_list_options():
    report_id = (request.args.get("report_id") or "").strip()
    if not report_id:
        return jsonify({"error": "report_id required"}), 400
    return jsonify(snap_db.list_snapshot_options(report_id))


@app.route("/snapshots/<int:snapshot_id>", methods=["GET", "DELETE"])
def snapshots_one(snapshot_id: int):
    if request.method == "DELETE":
        report_id = (request.args.get("report_id") or "").strip() or None
        if snap_db.delete_snapshot(snapshot_id, report_id=report_id):
            return jsonify({"deleted": True, "id": snapshot_id})
        return jsonify({"error": "not found"}), 404
    snap = snap_db.get_snapshot(snapshot_id)
    if not snap:
        return jsonify({"error": "not found"}), 404
    return jsonify(snap)


@app.route("/snapshots/<int:snapshot_id>/display", methods=["GET"])
def snapshots_display(snapshot_id: int):
    display = snap_db.snapshot_to_display(snapshot_id)
    if not display:
        return jsonify({"error": "not found"}), 404
    return jsonify(display)


@app.route("/snapshots", methods=["GET", "POST"])
def snapshots_collection():
    if request.method == "GET":
        report_id = (request.args.get("report_id") or "").strip()
        if not report_id:
            return jsonify({"error": "report_id required"}), 400
        rows = snap_db.list_snapshots(
            report_id,
            from_iso=request.args.get("from"),
            to_iso=request.args.get("to"),
            limit=int(request.args.get("limit") or 500),
        )
        return jsonify({"snapshots": rows})

    body = request.get_json(force=True)
    report_id = (body.get("report_id") or "").strip()
    metrics = body.get("metrics")
    if not report_id or not isinstance(metrics, dict):
        return jsonify({"error": "report_id and metrics required"}), 400
    try:
        snap_id = snap_db.save_snapshot(
            report_id,
            metrics,
            params=body.get("params"),
            note=body.get("note"),
        )
        return jsonify({"id": snap_id, "report_id": report_id})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/snapshots/compare", methods=["GET"])
def snapshots_compare():
    report_id = (request.args.get("report_id") or "").strip()
    snapshot_id = request.args.get("snapshot_id")
    if not report_id or not snapshot_id:
        return jsonify({"error": "report_id and snapshot_id required"}), 400
    member_username = (request.args.get("member_username") or "").strip() or None
    result = snap_db.compare_snapshots(
        report_id,
        int(snapshot_id),
        compare_to=request.args.get("compare_to") or "previous",
        member_username=member_username,
    )
    if member_username and result.get("deltas"):
        result["deltas"] = [d for d in result["deltas"] if d.get("member_username") == member_username or d.get("member_username") is None]
    return jsonify(result)


@app.route("/snapshots/compare-live", methods=["POST"])
def snapshots_compare_live():
    body = request.get_json(force=True)
    report_id = (body.get("report_id") or "").strip()
    metrics = body.get("metrics") or {}
    member_username = (body.get("member_username") or "").strip() or None
    if not report_id:
        return jsonify({"error": "report_id required"}), 400
    return jsonify(snap_db.compare_live_to_baseline(report_id, metrics, member_username=member_username))


@app.route("/snapshots/trends", methods=["GET"])
def snapshots_trends():
    report_id = (request.args.get("report_id") or "").strip()
    metric_key = (request.args.get("metric_key") or "").strip()
    if not report_id or not metric_key:
        return jsonify({"error": "report_id and metric_key required"}), 400
    member_username = (request.args.get("member_username") or "").strip() or None
    to_snapshot_id = request.args.get("to_snapshot_id")
    return jsonify(
        snap_db.build_trend_series(
            report_id,
            metric_key,
            member_username=member_username,
            to_snapshot_id=int(to_snapshot_id) if to_snapshot_id else None,
        )
    )


@app.route("/snapshots/label-trends", methods=["GET"])
def snapshots_label_trends():
    report_id = (request.args.get("report_id") or "ops").strip()
    member_username = (request.args.get("member_username") or "").strip()
    if report_id != "legacy" and not member_username:
        return jsonify({"error": "member_username required for ops label trends"}), 400
    top_n = int(request.args.get("top") or 10)
    limit = int(request.args.get("limit") or 20)
    to_snapshot_id = request.args.get("to_snapshot_id")
    return jsonify(
        snap_db.build_label_trend_series(
            report_id,
            member_username,
            top_n=top_n,
            limit=limit,
            to_snapshot_id=int(to_snapshot_id) if to_snapshot_id else None,
        )
    )


@app.route("/manual-baselines", methods=["GET", "POST"])
def manual_baselines():
    if request.method == "GET":
        report_id = (request.args.get("report_id") or "").strip()
        if not report_id:
            return jsonify({"error": "report_id required"}), 400
        member_username = (request.args.get("member_username") or "").strip() or None
        return jsonify({"baselines": snap_db.list_manual_baselines(report_id, member_username)})

    body = request.get_json(force=True)
    report_id = (body.get("report_id") or "").strip()
    metric_key = (body.get("metric_key") or "").strip()
    value = body.get("value")
    if not report_id or not metric_key or value is None:
        return jsonify({"error": "report_id, metric_key, and value required"}), 400
    snap_db.upsert_manual_baseline(
        report_id,
        metric_key,
        float(value),
        note=body.get("note"),
        member_username=(body.get("member_username") or "").strip() or None,
    )
    return jsonify({"ok": True})


if __name__ == "__main__":
    # Keep debug diagnostics but disable auto-reloader to avoid transient
    # connection resets/address-in-use during frequent file edits.
    app.run(debug=True, use_reloader=False, threaded=True, port=5001)
