import json
import sqlite3

from outlook_mail_assistant.outlook_actions import (
    create_outlook_appointment,
    create_outlook_task,
    log_outlook_action,
)


class _FakeItem:
    def __init__(self):
        self.Subject = None
        self.Body = None
        self.DueDate = None
        self.Start = None
        self.End = None
        self.BusyStatus = None
        self.saved = False

    def Save(self):
        self.saved = True


class _FakeApp:
    def __init__(self):
        self.created = []

    def CreateItem(self, kind):
        item = _FakeItem()
        self.created.append((kind, item))
        return item


def _sample_task_candidate():
    return {
        "message_id": "m-1",
        "store_name": "Primary Mailbox",
        "folder_path": "Inbox",
        "subject": "Please review the spec",
        "sender_email": "owner@example.com",
        "received_at": "2026-04-01T10:00:00Z",
        "kind": "task",
        "confidence": "explicit",
        "reason": "request-like wording detected",
        "snippet": "Could you review the attached spec by Friday?",
        "needs_confirmation": False,
    }


def _sample_meeting_candidate():
    return {
        "message_id": "m-2",
        "store_name": "Primary Mailbox",
        "folder_path": "Inbox",
        "subject": "Weekly meeting reminder",
        "sender_email": "events@example.com",
        "received_at": "2026-04-01T12:00:00Z",
        "kind": "meeting",
        "confidence": "inferred",
        "reason": "meeting wording with action cue detected",
        "snippet": "Please attend the weekly meeting.",
        "needs_confirmation": True,
        "start_at": "2026-04-03T10:00:00Z",
        "end_at": "2026-04-03T11:00:00Z",
    }


def test_create_outlook_task_dry_run_returns_payload_only():
    result = create_outlook_task(_sample_task_candidate(), dry_run=True)

    assert result["status"] == "dry_run"
    assert result["payload"]["subject"] == "Please review the spec"


def test_create_outlook_task_creates_and_saves_item():
    app = _FakeApp()

    result = create_outlook_task(_sample_task_candidate(), outlook_app=app, dry_run=False)

    assert result["status"] == "created"
    kind, item = app.created[0]
    assert kind == 3
    assert item.Subject == "Please review the spec"
    assert item.saved is True


def test_create_outlook_appointment_requires_start_and_end():
    candidate = _sample_meeting_candidate()
    candidate.pop("start_at")

    result = create_outlook_appointment(candidate, dry_run=True)

    assert result["status"] == "blocked"
    assert "start_at" in result["missing_fields"]


def test_create_outlook_appointment_creates_and_saves_item():
    app = _FakeApp()

    result = create_outlook_appointment(_sample_meeting_candidate(), outlook_app=app, dry_run=False)

    assert result["status"] == "created"
    kind, item = app.created[0]
    assert kind == 1
    assert item.Subject == "Weekly meeting reminder"
    assert item.saved is True


def test_log_outlook_action_writes_audit_row(tmp_path):
    db_path = tmp_path / "audit.sqlite3"
    connection = sqlite3.connect(db_path)
    try:
        connection.executescript(
            """
            CREATE TABLE audit_log (
                id INTEGER PRIMARY KEY,
                action_type TEXT NOT NULL,
                status TEXT NOT NULL,
                details_json TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        row_id = log_outlook_action(
            connection,
            action_type="create_task",
            status="dry_run",
            details={"message_id": "m-1"},
        )
        stored = connection.execute("SELECT details_json FROM audit_log WHERE id = ?", (row_id,)).fetchone()[0]
    finally:
        connection.close()

    assert json.loads(stored)["message_id"] == "m-1"
