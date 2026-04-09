import sqlite3

from outlook_mail_assistant.extraction import (
    classify_work_item_signal,
    extract_task_candidates,
    normalize_subject,
    persist_task_candidates,
)


def test_classify_work_item_signal_marks_explicit_requests():
    result = classify_work_item_signal(
        subject="Please review the spec",
        body_text="Could you review the attached spec by Friday?",
    )

    assert result["kind"] == "task"
    assert result["confidence"] == "explicit"
    assert result["needs_confirmation"] is False


def test_classify_work_item_signal_marks_inferred_deadlines():
    result = classify_work_item_signal(
        subject="Reminder: webinar on Friday",
        body_text="Please confirm attendance by Friday. Meeting starts at 10:00.",
    )

    assert result["kind"] in {"deadline", "meeting"}
    assert result["confidence"] == "inferred"
    assert result["needs_confirmation"] is True


def test_classify_work_item_signal_ignores_social_or_notice_messages():
    result = classify_work_item_signal(
        subject="사용자A 님이 [TEAM] Internal Collaboration Hub 님을 멘션했습니다.",
        body_text="알림 메일입니다.",
    )

    assert result is None


def test_classify_work_item_signal_treats_event_announcements_as_meeting_review():
    result = classify_work_item_signal(
        subject="Announcement & request: Int. Expert Workshop on Design and Safety of Next-Generation Ships",
        body_text="Workshop schedule and registration notice. Please register before Friday.",
    )

    assert result["kind"] == "meeting"
    assert result["confidence"] == "inferred"


def test_classify_work_item_signal_ignores_plain_event_notice_without_action_cue():
    result = classify_work_item_signal(
        subject="Technical Expo 2026",
        body_text="Expo schedule and venue 안내입니다.",
    )

    assert result is None


def test_classify_work_item_signal_flags_event_with_registration_or_attendance_request():
    result = classify_work_item_signal(
        subject="Workshop registration request",
        body_text="Please register and attend the workshop by Friday.",
    )

    assert result["kind"] == "meeting"
    assert result["needs_confirmation"] is True


def test_classify_work_item_signal_ignores_condolence_notice():
    result = classify_work_item_signal(
        subject="◆訃告◆ 연구팀 구성원 가족상 안내",
        body_text="부고 안내입니다.",
    )

    assert result is None


def test_normalize_subject_strips_re_fw_and_tags():
    assert normalize_subject("RE: FW: [EXTERNAL] [REMIND] Marine Systems Specifications Review") == "Marine Systems Specifications Review"


def test_classify_work_item_signal_ignores_weak_check_word_without_request_phrase():
    result = classify_work_item_signal(
        subject="진행상황 확인",
        body_text="회의 참고용으로 공유드립니다.",
    )

    assert result is None


def test_extract_task_candidates_collapses_repeated_normalized_subjects():
    records = [
        {
            "message_id": "a1",
            "store_name": "Primary",
            "folder_path": "Inbox",
            "subject": "RE: Marine Systems Specifications Review_Request for input",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T10:00:00Z",
            "body_text": "Please review the attached document.",
        },
        {
            "message_id": "a2",
            "store_name": "Primary",
            "folder_path": "Inbox",
            "subject": "FW: RE: Marine Systems Specifications Review_Request for input",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T11:00:00Z",
            "body_text": "Please review the attached document.",
        },
    ]

    items = extract_task_candidates(records)

    assert len(items) == 1
    assert items[0]["message_id"] == "a2"


def test_extract_task_candidates_returns_only_actionable_records():
    records = [
        {
            "message_id": "a",
            "store_name": "Primary",
            "folder_path": "Inbox",
            "subject": "Please review the spec",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T10:00:00Z",
            "body_text": "Could you review the attached spec by Friday?",
        },
        {
            "message_id": "b",
            "store_name": "Primary",
            "folder_path": "Inbox",
            "subject": "Newsletter",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T12:00:00Z",
            "body_text": "Monthly news update.",
        },
    ]

    items = extract_task_candidates(records)

    assert len(items) == 1
    assert items[0]["message_id"] == "a"
    assert items[0]["confidence"] == "explicit"
    assert items[0]["kind"] == "task"
    assert items[0]["needs_confirmation"] is False


def test_persist_task_candidates_inserts_into_tasks_table(tmp_path):
    db_path = tmp_path / "tasks.sqlite3"
    connection = sqlite3.connect(db_path)
    try:
        connection.executescript(
            """
            CREATE TABLE tasks (
                id INTEGER PRIMARY KEY,
                message_id TEXT,
                store_name TEXT,
                folder_path TEXT,
                subject TEXT,
                sender_email TEXT,
                received_at TEXT,
                kind TEXT,
                confidence TEXT,
                reason TEXT,
                snippet TEXT,
                needs_confirmation INTEGER
            );
            """
        )
        inserted = persist_task_candidates(
            connection,
            [
                {
                    "message_id": "a",
                    "store_name": "Primary",
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
            ],
        )
        count = connection.execute("SELECT COUNT(*) FROM tasks").fetchone()[0]
    finally:
        connection.close()

    assert inserted == 1
    assert count == 1
