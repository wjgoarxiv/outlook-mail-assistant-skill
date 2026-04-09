from pathlib import Path

from outlook_mail_assistant.summary_reports import build_mail_summary_markdown


def _sample_records():
    return [
        {
            "message_id": "m-1",
            "store_name": "Primary Mailbox",
            "folder_path": "Inbox",
            "subject": "Please review the spec",
            "sender_email": "owner@example.com",
            "received_at": "2026-03-31T10:00:00Z",
            "body_text": "Could you review the attached spec by Friday?",
        },
        {
            "message_id": "m-2",
            "store_name": "Primary Mailbox",
            "folder_path": "Inbox/Subfolder",
            "subject": "Weekly meeting reminder",
            "sender_email": "events@example.com",
            "received_at": "2026-04-01T12:00:00Z",
            "body_text": "Reminder for the weekly meeting.",
        },
        {
            "message_id": "m-3",
            "store_name": "Shared Mailbox",
            "folder_path": "Shared/Inbox",
            "subject": "Monthly budget review",
            "sender_email": "finance@example.com",
            "received_at": "2026-04-01T15:00:00Z",
            "body_text": "Please check the budget review notes.",
        },
    ]


def _sample_candidates():
    return [
        {
            "message_id": "m-1",
            "store_name": "Primary Mailbox",
            "folder_path": "Inbox",
            "subject": "Please review the spec",
            "sender_email": "owner@example.com",
            "received_at": "2026-03-31T10:00:00Z",
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
            "subject": "Weekly meeting reminder",
            "sender_email": "events@example.com",
            "received_at": "2026-04-01T12:00:00Z",
            "kind": "meeting",
            "confidence": "inferred",
            "reason": "meeting wording detected",
            "snippet": "Reminder for the weekly meeting.",
            "needs_confirmation": True,
        },
    ]


def test_build_mail_summary_markdown_contains_office_style_sections():
    markdown = build_mail_summary_markdown(
        records=_sample_records(),
        candidates=_sample_candidates(),
        title="Outlook Weekly and Monthly Summary Report",
    )

    assert "# Outlook Weekly and Monthly Summary Report" in markdown
    assert "## Executive Summary" in markdown
    assert "## Weekly Summary" in markdown
    assert "## Weekly Key Issues" in markdown
    assert "## Weekly Work Summary" in markdown
    assert "## Monthly Summary" in markdown
    assert "## Monthly Key Issues" in markdown
    assert "## Monthly Work Summary" in markdown
    assert "## Priority Follow-Ups" in markdown
    assert "## Communication Trends" in markdown


def test_build_mail_summary_markdown_includes_key_counts_and_subjects():
    markdown = build_mail_summary_markdown(
        records=_sample_records(),
        candidates=_sample_candidates(),
        title="Outlook Weekly and Monthly Summary Report",
    )

    assert "Total messages reviewed: 3" in markdown
    assert "Action candidates identified: 2" in markdown
    assert "Please review the spec" in markdown
    assert "Weekly meeting reminder" in markdown
    assert "Monthly budget review" in markdown
