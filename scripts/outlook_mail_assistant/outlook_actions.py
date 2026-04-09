from __future__ import annotations

import json
from datetime import datetime
from typing import Any


TASK_ITEM_KIND = 3
APPOINTMENT_ITEM_KIND = 1


def _parse_datetime(value: Any):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    return value


def build_task_item_payload(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "subject": item.get("subject") or "(no subject)",
        "body": "\n".join(
            [
                "Created by Outlook Mail Assistant",
                f"Source: {item.get('sender_email')}",
                f"Received: {item.get('received_at')}",
                f"Folder: {item.get('folder_path')}",
                "",
                item.get("snippet") or "",
            ]
        ).strip(),
        "due_date": item.get("due_at"),
    }


def build_calendar_item_payload(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "subject": item.get("subject") or "(no subject)",
        "body": "\n".join(
            [
                "Meeting/Calendar candidate created from Outlook Mail Assistant",
                f"Source: {item.get('sender_email')}",
                f"Received: {item.get('received_at')}",
                f"Folder: {item.get('folder_path')}",
                "",
                item.get("snippet") or "",
            ]
        ).strip(),
        "start_at": item.get("start_at"),
        "end_at": item.get("end_at"),
    }


def create_outlook_task(item: dict[str, Any], *, outlook_app=None, dry_run: bool = True):
    payload = build_task_item_payload(item)
    if dry_run:
        return {"status": "dry_run", "action_type": "create_task", "payload": payload}
    if outlook_app is None:
        raise RuntimeError("outlook_app is required for non-dry-run task creation")
    task = outlook_app.CreateItem(TASK_ITEM_KIND)
    task.Subject = payload["subject"]
    task.Body = payload["body"]
    if payload.get("due_date"):
        task.DueDate = _parse_datetime(payload["due_date"])
    task.Save()
    return {"status": "created", "action_type": "create_task", "payload": payload}


def create_outlook_appointment(item: dict[str, Any], *, outlook_app=None, dry_run: bool = True):
    payload = build_calendar_item_payload(item)
    missing = [field for field in ("start_at", "end_at") if not payload.get(field)]
    if missing:
        return {
            "status": "blocked",
            "action_type": "create_appointment",
            "missing_fields": missing,
            "payload": payload,
        }
    if dry_run:
        return {"status": "dry_run", "action_type": "create_appointment", "payload": payload}
    if outlook_app is None:
        raise RuntimeError("outlook_app is required for non-dry-run appointment creation")
    appointment = outlook_app.CreateItem(APPOINTMENT_ITEM_KIND)
    appointment.Subject = payload["subject"]
    appointment.Body = payload["body"]
    appointment.Start = _parse_datetime(payload["start_at"])
    appointment.End = _parse_datetime(payload["end_at"])
    appointment.BusyStatus = 2
    appointment.Save()
    return {"status": "created", "action_type": "create_appointment", "payload": payload}


def log_outlook_action(connection, *, action_type: str, status: str, details: dict[str, Any]) -> int:
    connection.execute(
        """
        INSERT INTO audit_log (action_type, status, details_json)
        VALUES (?, ?, ?)
        """,
        (action_type, status, json.dumps(details, ensure_ascii=False)),
    )
    connection.commit()
    row_id = connection.execute("SELECT last_insert_rowid()").fetchone()[0]
    return int(row_id)


# Backward-compatible names
def create_task_item(item: dict[str, Any], *, outlook_app, dry_run: bool = True):
    result = create_outlook_task(item, outlook_app=outlook_app, dry_run=dry_run)
    return {"dry_run": result["status"] == "dry_run", **result["payload"]}


def create_calendar_item(item: dict[str, Any], *, outlook_app, dry_run: bool = True):
    result = create_outlook_appointment(item, outlook_app=outlook_app, dry_run=dry_run)
    if result["status"] == "blocked":
        return {"dry_run": True, **result["payload"]}
    return {"dry_run": result["status"] == "dry_run", **result["payload"]}


def record_audit_event(connection, *, action_type: str, status: str, details: dict[str, Any]) -> int:
    return log_outlook_action(connection, action_type=action_type, status=status, details=details)
