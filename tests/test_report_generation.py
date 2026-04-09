from pathlib import Path

from openpyxl import load_workbook

from outlook_mail_assistant.report_generation import (
    write_task_report_csv,
    write_task_report_markdown,
    write_task_report_xlsx,
)


def _sample_items():
    return [
        {
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
        },
        {
            "message_id": "m-2",
            "store_name": "Primary Mailbox",
            "folder_path": "Inbox/Subfolder",
            "subject": "Webinar reminder",
            "sender_email": "events@example.com",
            "received_at": "2026-04-01T12:00:00Z",
            "kind": "meeting",
            "confidence": "inferred",
            "reason": "meeting wording detected",
            "snippet": "Reminder for tomorrow webinar.",
            "needs_confirmation": True,
        },
    ]


def test_write_task_report_markdown_creates_summary_and_rows(tmp_path: Path):
    output = tmp_path / "task-report.md"

    write_task_report_markdown(_sample_items(), output)

    text = output.read_text(encoding="utf-8")
    assert "# Task Report" in text
    assert "Please review the spec" in text
    assert "Webinar reminder" in text


def test_write_task_report_csv_writes_header_and_rows(tmp_path: Path):
    output = tmp_path / "task-report.csv"

    write_task_report_csv(_sample_items(), output)

    text = output.read_text(encoding="utf-8")
    assert "message_id,store_name,folder_path,subject" in text
    assert "m-1" in text
    assert "m-2" in text


def test_write_task_report_xlsx_creates_workbook(tmp_path: Path):
    output = tmp_path / "task-report.xlsx"

    write_task_report_xlsx(_sample_items(), output)

    workbook = load_workbook(output)
    sheet = workbook.active
    assert sheet["A1"].value == "message_id"
    assert sheet["D2"].value == "Please review the spec"
    workbook.close()
