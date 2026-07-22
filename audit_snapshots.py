"""CLI audit / migrate for snapshots.db. Run: python audit_snapshots.py [--migrate] [--json]"""
from __future__ import annotations

import argparse
import json
import sys

import snapshots_db as snap_db


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit (and optionally migrate) saved snapshots")
    parser.add_argument("--report-id", choices=["exec", "ops", "legacy"])
    parser.add_argument("--migrate", action="store_true", help="Backfill trend fields from view data")
    parser.add_argument("--json", action="store_true", help="Output JSON only")
    args = parser.parse_args()

    snap_db.init_db()

    if args.migrate:
        result = snap_db.migrate_all_snapshots(report_id=args.report_id)
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"Migrated {result['updated']} of {result['scanned']} snapshot(s).")
            for u in result.get("details") or []:
                print(f"  #{u['id']} {u['report_id']}: {', '.join(u.get('changes') or [])}")
        return 0

    report = snap_db.audit_all_snapshots(report_id=args.report_id)
    if args.json:
        print(json.dumps(report, indent=2))
        return 0 if report["summary"].get("error", 0) == 0 else 1

    s = report["summary"]
    print(
        f"Audited {s['total']} snapshot(s): "
        f"{s.get('ok', 0)} ok, {s.get('warn', 0)} warn, {s.get('error', 0)} error"
    )
    for row in report["snapshots"]:
        flags = row["issues"] + row["warnings"]
        line = f"#{row['id']} {row['report_id']} {row['captured_at'][:19]} [{row['status'].upper()}]"
        if row["note"]:
            line += f" — {row['note'][:40]}"
        print(line)
        for f in flags:
            print(f"    - {f}")
    return 0 if s.get("error", 0) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
