from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3

from .storage import bootstrap_workspace, initialize_database


def dedupe_records(records):
    unique = []
    seen = set()
    skipped = 0
    for record in records:
        key = (
            (record.get("message_id") or "").strip(),
            (record.get("store_name") or "").strip(),
        )
        if not key[0]:
            key = ((record.get("dedupe_hash") or "").strip(), "")
        if key in seen:
            skipped += 1
            continue
        seen.add(key)
        unique.append(record)
    return unique, skipped


def export_records_to_jsonl_chunks(records, *, output_dir: Path, chunk_size: int = 1000):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    files = []

    for start in range(0, len(records), chunk_size):
        chunk = records[start : start + chunk_size]
        path = output_dir / f"messages-{(start // chunk_size) + 1:05d}.jsonl"
        with path.open("w", encoding="utf-8") as handle:
            for record in chunk:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
        files.append(path)

    return files


def append_messages_to_db(connection: sqlite3.Connection, records) -> int:
    inserted = 0
    for record in records:
        before = connection.total_changes
        connection.execute(
            """
            INSERT OR IGNORE INTO messages (
                message_id, subject, sender_email, received_at, dedupe_hash, payload_json
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                record["message_id"],
                record["subject"],
                record["sender_email"],
                record.get("received_at"),
                record["dedupe_hash"],
                json.dumps(record, ensure_ascii=False),
            ),
        )
        inserted += int(connection.total_changes > before)
    connection.commit()
    return inserted


def append_task_candidates_to_db(connection: sqlite3.Connection, candidates) -> int:
    inserted = 0
    for candidate in candidates:
        before = connection.total_changes
        connection.execute(
            """
            INSERT INTO tasks (
                message_id, store_name, folder_path, subject, sender_email,
                received_at, kind, confidence, reason, snippet, needs_confirmation
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                candidate.get("message_id"),
                candidate.get("store_name"),
                candidate.get("folder_path"),
                candidate.get("subject"),
                candidate.get("sender_email"),
                candidate.get("received_at"),
                candidate.get("kind"),
                candidate.get("confidence"),
                candidate.get("reason"),
                candidate.get("snippet"),
                1 if candidate.get("needs_confirmation") else 0,
            ),
        )
        inserted += int(connection.total_changes > before)
    connection.commit()
    return inserted


def build_export_summary(
    *,
    records_exported: int,
    duplicate_skipped: int,
    stores_scanned: int,
    folders_scanned: int,
    excluded_folders: int,
    chunk_files: int,
    elapsed_seconds: float,
):
    return {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "exported_records": records_exported,
        "duplicate_skipped": duplicate_skipped,
        "stores_scanned": stores_scanned,
        "folders_scanned": folders_scanned,
        "excluded_folders": excluded_folders,
        "chunk_files": chunk_files,
        "elapsed_seconds": elapsed_seconds,
    }


def write_export_artifacts(
    *,
    workspace_root: Path,
    records,
    config: dict,
    chunk_size: int = 1000,
    sqlite_enabled: bool = False,
):
    workspace = bootstrap_workspace(Path(workspace_root))
    run_name = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    run_dir = workspace.exports / run_name
    run_dir.mkdir(parents=True, exist_ok=True)

    deduped, duplicate_skipped = dedupe_records(records)
    config_path = run_dir / "export-config.json"
    summary_path = run_dir / "export-summary.json"
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    files = export_records_to_jsonl_chunks(deduped, output_dir=run_dir, chunk_size=chunk_size)

    sqlite_inserted = 0
    db_path = None
    if sqlite_enabled:
        db_path = initialize_database(run_dir / "mail-index.sqlite3")
        connection = sqlite3.connect(db_path)
        try:
            sqlite_inserted = append_messages_to_db(connection, deduped)
        finally:
            connection.close()

    summary = build_export_summary(
        records_exported=len(deduped),
        duplicate_skipped=duplicate_skipped,
        stores_scanned=int(config.get("stores_scanned", 0)),
        folders_scanned=int(config.get("folders_scanned", 0)),
        excluded_folders=int(config.get("excluded_folders", 0)),
        chunk_files=len(files),
        elapsed_seconds=float(config.get("elapsed_seconds", 0.0)),
    )
    summary["sqlite_enabled"] = sqlite_enabled
    summary["sqlite_inserted"] = sqlite_inserted
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "run_dir": run_dir,
        "config_path": config_path,
        "summary_path": summary_path,
        "summary": summary,
        "files": files,
        "db_path": db_path,
    }
