import json

from outlook_mail_assistant.synthetic_dataset import generate_synthetic_mailset


def test_generate_synthetic_mailset_hits_target_size():
    base_records = [
        {
            "source_type": "outlook-com",
            "source_path": "Inbox/Message 1",
            "message_id": "id-1",
            "subject": "Please review the spec",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T10:00:00Z",
            "body_text": "Could you review the attached spec by Friday?",
            "dedupe_hash": "hash-1",
        }
    ]

    generated = generate_synthetic_mailset(base_records, target_count=5)

    assert len(generated) == 5
    assert len({record["message_id"] for record in generated}) == 5


def test_generate_synthetic_mailset_preserves_required_fields():
    base_records = [
        {
            "source_type": "outlook-com",
            "source_path": "Inbox/Message 1",
            "message_id": "id-1",
            "subject": "Please review the spec",
            "sender_email": "owner@example.com",
            "received_at": "2026-04-01T10:00:00Z",
            "body_text": "Could you review the attached spec by Friday?",
            "dedupe_hash": "hash-1",
        }
    ]

    generated = generate_synthetic_mailset(base_records, target_count=3)

    for record in generated:
        dumped = json.dumps(record)
        assert record["source_type"] == "synthetic"
        assert "message_id" in record
        assert "subject" in record
        assert dumped
