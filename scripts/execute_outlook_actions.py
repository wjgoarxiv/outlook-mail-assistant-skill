from __future__ import annotations

import argparse
import csv
from datetime import datetime
import json
from pathlib import Path
import sqlite3

from outlook_mail_assistant.outlook_actions import (
    create_outlook_appointment,
    create_outlook_task,
    log_outlook_action,
)
from outlook_mail_assistant.storage import initialize_database


def _read_candidates(path: Path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8", newline="") as handle:
            return list(csv.DictReader(handle))
    raise RuntimeError("Only CSV input is supported for execute_outlook_actions.py")


def _get_outlook_application():
    import win32com.client  # type: ignore

    return win32com.client.Dispatch("Outlook.Application")


def _apply_action_overrides(item, *, start_at=None, end_at=None):
    updated = dict(item)
    if start_at:
        updated["start_at"] = start_at
    if end_at:
        updated["end_at"] = end_at
    return updated


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Outlook tasks/calendar items from extracted candidates.")
    parser.add_argument("--input-csv", required=True)
    parser.add_argument("--audit-db")
    parser.add_argument("--include-needs-confirmation", action="store_true")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--kind", choices=("task", "meeting"))
    parser.add_argument("--subject-contains")
    parser.add_argument("--start-at")
    parser.add_argument("--end-at")
    args = parser.parse_args()

    candidates = _read_candidates(Path(args.input_csv))
    if not args.include_needs_confirmation:
        candidates = [item for item in candidates if str(item.get("needs_confirmation")).lower() not in {"true", "1", "yes"}]
    if args.kind:
        candidates = [item for item in candidates if item.get("kind") == args.kind]
    if args.subject_contains:
        needle = args.subject_contains.lower()
        candidates = [item for item in candidates if needle in str(item.get("subject", "")).lower()]
    if args.limit is not None:
        candidates = candidates[: args.limit]

    app = None if not args.apply else _get_outlook_application()
    connection = None
    if args.audit_db:
        db_path = initialize_database(Path(args.audit_db))
        connection = sqlite3.connect(db_path)

    try:
        results = []
        for item in candidates:
            item = _apply_action_overrides(item, start_at=args.start_at, end_at=args.end_at)
            kind = item.get("kind")
            if kind == "meeting":
                result = create_outlook_appointment(item, outlook_app=app, dry_run=not args.apply)
            else:
                result = create_outlook_task(item, outlook_app=app, dry_run=not args.apply)
            results.append(result)
            if connection is not None:
                log_outlook_action(
                    connection,
                    action_type=result["action_type"],
                    status=result["status"],
                    details=result["payload"] if "payload" in result else result,
                )
    finally:
        if connection is not None:
            connection.close()

    print(json.dumps({"processed": len(results), "applied": bool(args.apply)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
