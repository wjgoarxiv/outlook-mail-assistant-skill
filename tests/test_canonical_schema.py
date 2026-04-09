from datetime import datetime, timezone

import pytest

from outlook_mail_assistant.canonical_schema import CanonicalMessage, canonical_dedupe_hash


def test_canonical_message_normalizes_datetimes_to_utc_isoformat():
    message = CanonicalMessage(
        source_type="msg",
        source_path="fixtures/sample.msg",
        message_id="abc-123",
        subject="Quarterly review",
        sender_email="owner@example.com",
        received_at=datetime(2026, 4, 2, 1, 30, tzinfo=timezone.utc),
    )

    payload = message.to_record()

    assert payload["received_at"] == "2026-04-02T01:30:00Z"


def test_canonical_message_requires_core_fields():
    with pytest.raises(ValueError):
        CanonicalMessage(
            source_type="",
            source_path="fixtures/sample.msg",
            message_id="abc-123",
            subject="Quarterly review",
            sender_email="owner@example.com",
        )


def test_canonical_dedupe_hash_is_stable_for_same_identity_fields():
    left = canonical_dedupe_hash(
        message_id="abc-123",
        subject="Quarterly review",
        sender_email="owner@example.com",
        received_at="2026-04-02T01:30:00Z",
    )
    right = canonical_dedupe_hash(
        message_id="abc-123",
        subject="Quarterly review",
        sender_email="owner@example.com",
        received_at="2026-04-02T01:30:00Z",
    )

    assert left == right


def test_canonical_dedupe_hash_changes_when_identity_changes():
    left = canonical_dedupe_hash(
        message_id="abc-123",
        subject="Quarterly review",
        sender_email="owner@example.com",
        received_at="2026-04-02T01:30:00Z",
    )
    right = canonical_dedupe_hash(
        message_id="abc-124",
        subject="Quarterly review",
        sender_email="owner@example.com",
        received_at="2026-04-02T01:30:00Z",
    )

    assert left != right
