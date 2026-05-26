"""SQLite storage for official dashboard report snapshots and trend analysis."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "snapshots.db"

REPORT_DEFINITIONS = [
    ("exec", "Executive Report", "weekly"),
    ("ops", "Operations Team", "2x_daily"),
    ("legacy", "Ticket trend", "manual"),
]

OPS_BOARD_METRIC_KEYS = (
    "pipeline_backlog_count",
    "closed_cssd_csd_team_count",
    "queue_backlog_count",
    "in_progress_count",
    "resolved_in_period_count",
    "sla_breach_count",
    "open_near_sla_breach_8h_count",
)

OPS_MEMBER_METRIC_KEYS = (
    "assigned_open_count",
    "queue_backlog_count",
    "in_progress_count",
    "worked_status_last_8h_count",
    "worked_status_last_8h_assigned_others_count",
    "reopened_count",
    "resolved_owned_count",
    "resolved_contributed_count",
    "closed_cssd_csd_count",
    "resolved_last_8h_count",
    "worked_on_assigned_others_count",
    "sla_breach_count",
    "open_near_sla_breach_8h_count",
)

EXEC_TREND_KEYS = ("backlog", "new_created", "resolved")
LEGACY_TREND_KEYS = (
    "issue_count",
    "transition_count",
    "comment_count",
    "date_window_days",
    "ttfr_cssd_median_hours",
    "ttfr_csd_median_hours",
    "ttr_cssd_median_hours",
    "ttr_csd_median_hours",
)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _json_loads(text: Optional[str], default: Any) -> Any:
    if not text:
        return default
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return default


@contextmanager
def db_connection():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def init_db() -> None:
    with db_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS report_definitions (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                suggested_cadence TEXT,
                enabled INTEGER NOT NULL DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id TEXT NOT NULL,
                captured_at TEXT NOT NULL,
                trigger_type TEXT NOT NULL DEFAULT 'manual',
                params_json TEXT,
                metrics_json TEXT NOT NULL,
                note TEXT,
                FOREIGN KEY (report_id) REFERENCES report_definitions(id)
            );
            CREATE TABLE IF NOT EXISTS manual_baselines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id TEXT NOT NULL,
                metric_key TEXT NOT NULL,
                value REAL NOT NULL,
                effective_at TEXT NOT NULL,
                note TEXT,
                member_username TEXT,
                UNIQUE(report_id, metric_key, member_username)
            );
            CREATE INDEX IF NOT EXISTS idx_snapshots_report_time
                ON snapshots (report_id, captured_at);
            """
        )
        for row in REPORT_DEFINITIONS:
            conn.execute(
                """
                INSERT OR IGNORE INTO report_definitions (id, name, suggested_cadence)
                VALUES (?, ?, ?)
                """,
                row,
            )
        conn.commit()


def _row_to_snapshot(row: sqlite3.Row) -> Dict[str, Any]:
    return {
        "id": row["id"],
        "report_id": row["report_id"],
        "captured_at": row["captured_at"],
        "trigger_type": row["trigger_type"],
        "params": _json_loads(row["params_json"], {}),
        "metrics": _json_loads(row["metrics_json"], {}),
        "note": row["note"] or "",
    }


def save_snapshot(
    report_id: str,
    metrics: Dict[str, Any],
    params: Optional[Dict[str, Any]] = None,
    note: Optional[str] = None,
    trigger_type: str = "manual",
) -> int:
    captured_at = _utc_now_iso()
    with db_connection() as conn:
        cur = conn.execute(
            """
            INSERT INTO snapshots (report_id, captured_at, trigger_type, params_json, metrics_json, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                report_id,
                captured_at,
                trigger_type,
                json.dumps(params or {}, default=str),
                json.dumps(metrics, default=str),
                note or "",
            ),
        )
        conn.commit()
        return int(cur.lastrowid)


def get_snapshot(snapshot_id: int) -> Optional[Dict[str, Any]]:
    with db_connection() as conn:
        row = conn.execute("SELECT * FROM snapshots WHERE id = ?", (snapshot_id,)).fetchone()
    return _row_to_snapshot(row) if row else None


def delete_snapshot(snapshot_id: int, report_id: Optional[str] = None) -> bool:
    """Delete one snapshot row. If report_id is set, the row must match that report."""
    with db_connection() as conn:
        if report_id:
            cur = conn.execute(
                "DELETE FROM snapshots WHERE id = ? AND report_id = ?",
                (snapshot_id, report_id),
            )
        else:
            cur = conn.execute("DELETE FROM snapshots WHERE id = ?", (snapshot_id,))
        conn.commit()
        return cur.rowcount > 0


def get_latest_snapshot(report_id: str) -> Optional[Dict[str, Any]]:
    with db_connection() as conn:
        row = conn.execute(
            """
            SELECT * FROM snapshots WHERE report_id = ?
            ORDER BY captured_at DESC, id DESC LIMIT 1
            """,
            (report_id,),
        ).fetchone()
    return _row_to_snapshot(row) if row else None


def list_snapshots(
    report_id: str,
    from_iso: Optional[str] = None,
    to_iso: Optional[str] = None,
    limit: int = 500,
) -> List[Dict[str, Any]]:
    sql = "SELECT * FROM snapshots WHERE report_id = ?"
    args: List[Any] = [report_id]
    if from_iso:
        sql += " AND captured_at >= ?"
        args.append(from_iso)
    if to_iso:
        sql += " AND captured_at <= ?"
        args.append(to_iso)
    sql += " ORDER BY captured_at DESC, id DESC LIMIT ?"
    args.append(limit)
    with db_connection() as conn:
        rows = conn.execute(sql, args).fetchall()
    return [_row_to_snapshot(r) for r in rows]


def _format_option_label(captured_at: str, note: str, snap_id: int) -> str:
    try:
        dt = datetime.fromisoformat(captured_at.replace("Z", "+00:00"))
        stamp = dt.strftime("%Y-%m-%d %H:%M UTC")
    except ValueError:
        stamp = captured_at
    suffix = f" — {note}" if note else ""
    return f"{stamp}{suffix} (#{snap_id})"


def list_snapshot_options(report_id: str) -> Dict[str, Any]:
    snaps = list_snapshots(report_id, limit=200)
    latest_id = snaps[0]["id"] if snaps else None
    options = [
        {
            "id": s["id"],
            "captured_at": s["captured_at"],
            "note": s["note"],
            "label": _format_option_label(s["captured_at"], s["note"], s["id"]),
        }
        for s in snaps
    ]
    cadence = None
    with db_connection() as conn:
        row = conn.execute(
            "SELECT suggested_cadence FROM report_definitions WHERE id = ?",
            (report_id,),
        ).fetchone()
        if row:
            cadence = row["suggested_cadence"]
    return {"report_id": report_id, "latest_id": latest_id, "suggested_cadence": cadence, "options": options}


def _coerce_metric_float(value: Any) -> Optional[float]:
    """Convert a metric value to float; None and invalid values mean metric unavailable."""
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _get_metric_value(
    metrics: Dict[str, Any],
    metric_key: str,
    member_username: Optional[str] = None,
) -> Optional[float]:
    if member_username:
        members = metrics.get("trend", {}).get("members") or metrics.get("view", {}).get("members") or []
        if not isinstance(members, list):
            members = []
        target = member_username.strip().lower()
        for m in members:
            uname = (m.get("username") or "").strip().lower()
            if uname != target:
                continue
            vals = m.get("metrics") or m
            if metric_key in vals:
                return _coerce_metric_float(vals[metric_key])
            return None
        return None
    board = metrics.get("trend", {}).get("board") or metrics.get("view", {}).get("board") or {}
    if metric_key in board:
        return _coerce_metric_float(board[metric_key])
    trend = metrics.get("trend") or {}
    if metric_key in trend:
        return _coerce_metric_float(trend[metric_key])
    kpis = (metrics.get("view") or {}).get("kpis") or {}
    if metric_key in kpis:
        return _coerce_metric_float(kpis[metric_key])
    return None


def _previous_snapshot(report_id: str, snapshot_id: int) -> Optional[Dict[str, Any]]:
    current = get_snapshot(snapshot_id)
    if not current:
        return None
    with db_connection() as conn:
        row = conn.execute(
            """
            SELECT * FROM snapshots
            WHERE report_id = ? AND captured_at < ?
            ORDER BY captured_at DESC, id DESC LIMIT 1
            """,
            (report_id, current["captured_at"]),
        ).fetchone()
    return _row_to_snapshot(row) if row else None


def _get_manual_baseline(
    report_id: str,
    metric_key: str,
    member_username: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    with db_connection() as conn:
        row = conn.execute(
            """
            SELECT * FROM manual_baselines
            WHERE report_id = ? AND metric_key = ?
              AND (member_username IS ? OR (member_username IS NULL AND ? IS NULL))
            ORDER BY effective_at DESC, id DESC LIMIT 1
            """,
            (report_id, metric_key, member_username, member_username),
        ).fetchone()
    if not row:
        return None
    return {
        "value": row["value"],
        "effective_at": row["effective_at"],
        "note": row["note"] or "",
    }


def upsert_manual_baseline(
    report_id: str,
    metric_key: str,
    value: float,
    note: Optional[str] = None,
    member_username: Optional[str] = None,
) -> None:
    with db_connection() as conn:
        conn.execute(
            """
            INSERT INTO manual_baselines (report_id, metric_key, value, effective_at, note, member_username)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(report_id, metric_key, member_username) DO UPDATE SET
                value = excluded.value,
                effective_at = excluded.effective_at,
                note = excluded.note
            """,
            (report_id, metric_key, value, _utc_now_iso(), note or "", member_username),
        )
        conn.commit()


def list_manual_baselines(
    report_id: str,
    member_username: Optional[str] = None,
) -> List[Dict[str, Any]]:
    sql = "SELECT * FROM manual_baselines WHERE report_id = ?"
    args: List[Any] = [report_id]
    if member_username is not None:
        sql += " AND member_username = ?"
        args.append(member_username)
    with db_connection() as conn:
        rows = conn.execute(sql + " ORDER BY metric_key", args).fetchall()
    return [
        {
            "id": r["id"],
            "report_id": r["report_id"],
            "metric_key": r["metric_key"],
            "value": r["value"],
            "effective_at": r["effective_at"],
            "note": r["note"] or "",
            "member_username": r["member_username"],
        }
        for r in rows
    ]


def _pct_change(current: float, baseline: float) -> Optional[float]:
    if baseline == 0:
        return None if current == 0 else 100.0
    return round(((current - baseline) / baseline) * 100.0, 1)


def compare_snapshots(
    report_id: str,
    snapshot_id: int,
    compare_to: str = "previous",
    metric_keys: Optional[List[str]] = None,
    member_username: Optional[str] = None,
    member_only: bool = False,
) -> Dict[str, Any]:
    current = get_snapshot(snapshot_id)
    if not current:
        return {"error": "snapshot not found"}

    metrics = current["metrics"]
    if member_username:
        keys = metric_keys or list(OPS_MEMBER_METRIC_KEYS)
    else:
        keys = metric_keys or _default_metric_keys(report_id, None)
    baseline_snap = _previous_snapshot(report_id, snapshot_id) if compare_to == "previous" else None

    baseline_info: Dict[str, Any] = {"source": None}
    deltas: List[Dict[str, Any]] = []

    for key in keys:
        cur_val = _get_metric_value(metrics, key, member_username)
        if cur_val is None:
            continue
        base_val: Optional[float] = None
        if baseline_snap:
            base_val = _get_metric_value(baseline_snap["metrics"], key, member_username)
            if base_val is not None:
                baseline_info = {
                    "source": "snapshot",
                    "id": baseline_snap["id"],
                    "captured_at": baseline_snap["captured_at"],
                    "note": baseline_snap["note"],
                }
        if base_val is None:
            manual = _get_manual_baseline(report_id, key, member_username)
            if manual:
                base_val = float(manual["value"])
                baseline_info = {
                    "source": "manual",
                    "effective_at": manual["effective_at"],
                    "note": manual["note"],
                }
        if base_val is None:
            continue
        change = cur_val - base_val
        deltas.append(
            {
                "metric_key": key,
                "member_username": member_username,
                "current": cur_val,
                "baseline": base_val,
                "change": change,
                "pct_change": _pct_change(cur_val, base_val),
            }
        )

    return {
        "current": {
            "id": current["id"],
            "captured_at": current["captured_at"],
            "note": current["note"],
        },
        "baseline": baseline_info,
        "deltas": deltas,
    }


def _live_metric_value(
    live_metrics: Dict[str, Any],
    metric_key: str,
    member_username: Optional[str] = None,
) -> Optional[float]:
    if member_username:
        target = member_username.strip().lower()
        for m in live_metrics.get("members") or []:
            if (m.get("username") or "").strip().lower() != target:
                continue
            metrics = m.get("metrics") or {}
            if metric_key in metrics:
                return _coerce_metric_float(metrics[metric_key])
        return None
    board = live_metrics.get("board") or live_metrics
    if metric_key in board:
        return _coerce_metric_float(board[metric_key])
    kpis = live_metrics.get("kpis")
    if isinstance(kpis, dict) and metric_key in kpis:
        return _coerce_metric_float(kpis[metric_key])
    return None


def compare_live_to_baseline(
    report_id: str,
    live_metrics: Dict[str, Any],
    metric_keys: Optional[List[str]] = None,
    member_username: Optional[str] = None,
) -> Dict[str, Any]:
    """Compare unsaved live metrics to latest snapshot or manual baseline."""
    latest = get_latest_snapshot(report_id)
    keys = metric_keys or _default_metric_keys(report_id, member_username)
    baseline_info: Dict[str, Any] = {"source": None}
    deltas: List[Dict[str, Any]] = []

    for key in keys:
        cur_val = _live_metric_value(live_metrics, key, member_username)
        if cur_val is None:
            continue

        base_val: Optional[float] = None
        if latest:
            base_val = _get_metric_value(latest["metrics"], key, member_username)
            if base_val is not None:
                baseline_info = {
                    "source": "snapshot",
                    "id": latest["id"],
                    "captured_at": latest["captured_at"],
                    "note": latest["note"],
                }
        if base_val is None:
            manual = _get_manual_baseline(report_id, key, member_username)
            if manual:
                base_val = float(manual["value"])
                baseline_info = {"source": "manual", "effective_at": manual["effective_at"], "note": manual["note"]}
        if base_val is None:
            continue
        change = cur_val - base_val
        deltas.append(
            {
                "metric_key": key,
                "member_username": member_username,
                "current": cur_val,
                "baseline": base_val,
                "change": change,
                "pct_change": _pct_change(cur_val, base_val),
            }
        )

    return {"current": {"source": "live"}, "baseline": baseline_info, "deltas": deltas}


def _default_metric_keys(report_id: str, member_username: Optional[str]) -> List[str]:
    if report_id == "exec":
        return list(EXEC_TREND_KEYS)
    if report_id == "legacy":
        return list(LEGACY_TREND_KEYS)
    if report_id == "ops":
        return list(OPS_MEMBER_METRIC_KEYS if member_username else OPS_BOARD_METRIC_KEYS)
    return []


def build_trend_series(
    report_id: str,
    metric_key: str,
    member_username: Optional[str] = None,
    to_snapshot_id: Optional[int] = None,
    limit: int = 30,
) -> Dict[str, Any]:
    snaps = list_snapshots(report_id, limit=500)
    snaps = sorted(snaps, key=lambda s: s["captured_at"])
    if to_snapshot_id:
        cutoff = None
        for s in snaps:
            if s["id"] == to_snapshot_id:
                cutoff = s["captured_at"]
                break
        if cutoff:
            snaps = [s for s in snaps if s["captured_at"] <= cutoff]
    snaps = snaps[-limit:]

    labels: List[str] = []
    data: List[Optional[float]] = []
    points: List[Dict[str, Any]] = []

    for s in snaps:
        val = _get_metric_value(s["metrics"], metric_key, member_username)
        if val is None:
            continue
        try:
            dt = datetime.fromisoformat(s["captured_at"].replace("Z", "+00:00"))
            label = dt.strftime("%m/%d")
        except ValueError:
            label = s["captured_at"][:10]
        labels.append(label)
        data.append(val)
        points.append({"id": s["id"], "captured_at": s["captured_at"], "value": val, "note": s["note"]})

    return {
        "report_id": report_id,
        "metric_key": metric_key,
        "member_username": member_username,
        "labels": labels,
        "data": data,
        "points": points,
    }


def _member_label_distribution_from_snapshot(
    metrics: Dict[str, Any], member_username: str
) -> Dict[str, Any]:
    target = (member_username or "").strip().lower()
    if not target:
        return {}
    members = (metrics.get("view") or {}).get("members") or []
    if not isinstance(members, list):
        return {}
    for m in members:
        if (m.get("username") or "").strip().lower() == target:
            dist = m.get("label_distribution") or {}
            return dist if isinstance(dist, dict) else {}
    return {}


def _legacy_label_distribution_from_snapshot(metrics: Dict[str, Any]) -> Dict[str, Any]:
    charts = (metrics.get("view") or {}).get("charts") or {}
    dist = charts.get("label_distribution") or {}
    return dist if isinstance(dist, dict) else {}


def _label_distribution_from_snapshot(
    report_id: str, metrics: Dict[str, Any], member_username: str
) -> Dict[str, Any]:
    if report_id == "legacy":
        return _legacy_label_distribution_from_snapshot(metrics)
    return _member_label_distribution_from_snapshot(metrics, member_username)


def build_label_trend_series(
    report_id: str,
    member_username: str = "",
    top_n: int = 10,
    limit: int = 20,
    to_snapshot_id: Optional[int] = None,
) -> Dict[str, Any]:
    """Time series of top Jira label counts across saved snapshots (ops: per member; legacy: whole report)."""
    target = (member_username or "").strip().lower()
    if report_id != "legacy" and not target:
        return {
            "report_id": report_id,
            "member_username": member_username,
            "time_labels": [],
            "datasets": [],
            "top_labels": [],
            "snapshot_count": 0,
        }

    snaps = list_snapshots(report_id, limit=500)
    snaps = sorted(snaps, key=lambda s: s["captured_at"])
    if to_snapshot_id:
        cutoff = None
        for s in snaps:
            if s["id"] == to_snapshot_id:
                cutoff = s["captured_at"]
                break
        if cutoff:
            snaps = [s for s in snaps if s["captured_at"] <= cutoff]

    rows: List[Tuple[str, str, Dict[str, Any]]] = []
    label_totals: Dict[str, int] = {}
    for s in snaps:
        dist = _label_distribution_from_snapshot(report_id, s["metrics"], target)
        if not dist:
            continue
        try:
            dt = datetime.fromisoformat(s["captured_at"].replace("Z", "+00:00"))
            time_label = dt.strftime("%m/%d %H:%M")
        except ValueError:
            time_label = (s["captured_at"] or "")[:16]
        rows.append((time_label, s["captured_at"], dist))
        for key, val in dist.items():
            label_totals[str(key)] = label_totals.get(str(key), 0) + int(val or 0)

    rows = rows[-limit:]
    top_labels = [
        k
        for k, _ in sorted(label_totals.items(), key=lambda item: (-item[1], item[0]))[: max(1, top_n)]
    ]

    datasets: List[Dict[str, Any]] = []
    for label in top_labels:
        datasets.append(
            {
                "label": label,
                "data": [int(dist.get(label, 0) or 0) for _, _, dist in rows],
            }
        )

    return {
        "report_id": report_id,
        "member_username": member_username,
        "time_labels": [r[0] for r in rows],
        "datasets": datasets,
        "top_labels": top_labels,
        "snapshot_count": len(rows),
    }


def extract_metrics_for_save(report_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    if report_id == "exec":
        kpis = payload.get("kpis") or {}
        return {
            "view": {
                "kpis": kpis,
                "periods": payload.get("periods"),
                "operational_health": payload.get("operational_health"),
                "narratives": payload.get("narratives"),
                "charts": payload.get("charts"),
                "elapsed_time_sentence": payload.get("elapsed_time_sentence"),
            },
            "trend": {
                "backlog": (kpis.get("backlog") or {}).get("period2"),
                "new_created": (kpis.get("new_created") or {}).get("period2"),
                "resolved": (kpis.get("resolved") or {}).get("period2"),
            },
        }
    if report_id == "legacy":
        kpis = payload.get("kpis") or {}
        return {
            "view": {
                "kpis": kpis,
                "charts": payload.get("charts"),
                "insights": payload.get("insights"),
                "warnings": payload.get("warnings"),
            },
            "trend": {
                "issue_count": kpis.get("issue_count"),
                "transition_count": kpis.get("transition_count"),
                "comment_count": kpis.get("comment_count"),
                "date_window_days": kpis.get("date_window_days"),
            },
        }
    if report_id == "ops":
        return payload
    raise ValueError(f"Unknown report_id: {report_id}")


def build_ops_metrics_from_client(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Payload from client: { board: {...}, members: [...] }."""
    board = payload.get("board") or {}
    members = payload.get("members") or []
    member_trends = []
    view_members = []
    for m in members:
        metrics = m.get("metrics") or {}
        uname = m.get("username") or m.get("assignee_username") or ""
        name = m.get("name") or ""
        trend_m = {k: metrics.get(k) for k in OPS_MEMBER_METRIC_KEYS if k in metrics}
        member_trends.append({"username": uname, "name": name, "metrics": trend_m})
        view_members.append(
            {
                "username": uname,
                "name": name,
                "metrics": metrics,
                "status_distribution": m.get("status_distribution") or {},
                "label_distribution": m.get("label_distribution") or {},
                "oldest_open": m.get("oldest_open") or {},
            }
        )
    board_trend = {k: board.get(k) for k in OPS_BOARD_METRIC_KEYS if k in board}
    return {
        "view": {"board": board, "members": view_members},
        "trend": {"board": board_trend, "members": member_trends},
    }


def snapshot_to_display(snapshot_id: int) -> Optional[Dict[str, Any]]:
    snap = get_snapshot(snapshot_id)
    if not snap:
        return None
    report_id = snap["report_id"]
    metrics = snap["metrics"]
    view = metrics.get("view") or {}

    saved_params = snap.get("params") or {}

    if report_id == "exec":
        return {
            "report_id": "exec",
            "snapshot_id": snap["id"],
            "captured_at": snap["captured_at"],
            "note": snap["note"],
            "params": saved_params,
            "periods": view.get("periods"),
            "kpis": view.get("kpis"),
            "charts": view.get("charts") or {},
            "narratives": view.get("narratives") or {},
            "operational_health": view.get("operational_health") or {},
            "elapsed_time_sentence": view.get("elapsed_time_sentence") or "",
        }
    if report_id == "legacy":
        return {
            "report_id": "legacy",
            "snapshot_id": snap["id"],
            "captured_at": snap["captured_at"],
            "note": snap["note"],
            "params": saved_params,
            "kpis": view.get("kpis") or {},
            "charts": view.get("charts") or {},
            "insights": view.get("insights") or [],
            "warnings": view.get("warnings") or [],
        }
    if report_id == "ops":
        return {
            "report_id": "ops",
            "snapshot_id": snap["id"],
            "captured_at": snap["captured_at"],
            "note": snap["note"],
            "params": saved_params,
            "board": view.get("board") or {},
            "members": view.get("members") or [],
        }
    return None
