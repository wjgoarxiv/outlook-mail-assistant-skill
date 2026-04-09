import json
import sqlite3

from outlook_mail_assistant.export_pipeline import (
    append_messages_to_db,
    dedupe_records,
    export_records_to_jsonl_chunks,
    write_export_artifacts,
)


def _sample_records():
    return [
        {
            "source_type": "outlook-com",
            "source_path": f"Inbox/Message {i}",
            "message_id": f"id-{i}",
            "subject": f"Message {i}",
            "sender_email": "owner@example.com",
            "received_at": f"2026-04-0{i}T01:00:00Z",
            "body_text": "Please review the attached spec by Friday.",
            "dedupe_hash": f"hash-{i}",
        }
        for i in range(1, 4)
    ]


def test_export_records_to_jsonl_chunks_splits_by_chunk_size(tmp_path):
    output_dir = tmp_path / "export"
    files = export_records_to_jsonl_chunks(
        _sample_records(),
        output_dir=output_dir,
        chunk_size=2,
    )

    assert len(files) == 2
    first_lines = files[0].read_text(encoding="utf-8").strip().splitlines()
    second_lines = files[1].read_text(encoding="utf-8").strip().splitlines()

    assert len(first_lines) == 2
    assert len(second_lines) == 1
    assert json.loads(first_lines[0])["message_id"] == "id-1"


def test_append_messages_to_db_inserts_canonical_records(tmp_path):
    db_path = tmp_path / "mail-index.sqlite3"
    connection = sqlite3.connect(db_path)
    try:
        connection.executescript(
            """
            CREATE TABLE messages (
                id INTEGER PRIMARY KEY,
                message_id TEXT NOT NULL,
                subject TEXT NOT NULL,
                sender_email TEXT NOT NULL,
                received_at TEXT,
                dedupe_hash TEXT NOT NULL UNIQUE,
                payload_json TEXT NOT NULL
            );
            """
        )
        inserted = append_messages_to_db(connection, _sample_records())
        count = connection.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    finally:
        connection.close()

    assert inserted == 3
    assert count == 3


def test_dedupe_records_prefers_message_id_plus_store_name():
    records = _sample_records()
    records[0]["store_name"] = "Primary"
    records[1]["store_name"] = "Primary"
    records[2]["store_name"] = "Primary"
    duplicate = dict(records[0])
    duplicate["subject"] = "Changed title"
    records.append(duplicate)

    deduped, skipped = dedupe_records(records)

    assert len(deduped) == 3
    assert skipped == 1


def test_write_export_artifacts_creates_config_and_summary(tmp_path):
    records = _sample_records()
    for record in records:
        record["store_name"] = "Primary"
        record["folder_path"] = "Inbox"

    result = write_export_artifacts(
        workspace_root=tmp_path / "workspace",
        records=records,
        config={"store_scope": "primary-shared", "recursive": True},
        chunk_size=2,
        sqlite_enabled=True,
    )

    assert result["run_dir"].exists()
    assert result["config_path"].exists()
    assert result["summary_path"].exists()
    summary = json.loads(result["summary_path"].read_text(encoding="utf-8"))
    assert summary["exported_records"] == 3
    assert summary["chunk_files"] == 2
    assert summary["duplicate_skipped"] == 0
