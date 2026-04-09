from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import sqlite3


REQUIRED_TABLES = {"messages", "tasks", "decisions", "audit_log"}


@dataclass(slots=True)
class WorkspacePaths:
    root: Path
    raw: Path
    normalized: Path
    reports: Path
    exports: Path
    logs: Path
    manifest_path: Path


def bootstrap_workspace(root: Path) -> WorkspacePaths:
    root = Path(root)
    raw = root / "raw"
    normalized = root / "normalized"
    reports = root / "reports"
    exports = root / "exports"
    logs = root / "logs"

    for path in (root, raw, normalized, reports, exports, logs):
        path.mkdir(parents=True, exist_ok=True)

    manifest_path = root / "workspace.json"
    if not manifest_path.exists():
        manifest_path.write_text(
            json.dumps(
                {
                    "version": 1,
                    "workspace_name": root.name,
                },
                indent=2,
            ),
            encoding="utf-8",
        )

    return WorkspacePaths(
        root=root,
        raw=raw,
        normalized=normalized,
        reports=reports,
        exports=exports,
        logs=logs,
        manifest_path=manifest_path,
    )


def initialize_database(db_path: Path) -> Path:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(db_path) as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY,
                message_id TEXT NOT NULL,
                subject TEXT NOT NULL,
                sender_email TEXT NOT NULL,
                received_at TEXT,
                dedupe_hash TEXT NOT NULL UNIQUE,
                payload_json TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS tasks (
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
                needs_confirmation INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS decisions (
                id INTEGER PRIMARY KEY,
                message_ref TEXT,
                summary TEXT NOT NULL,
                decided_at TEXT
            );

            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY,
                action_type TEXT NOT NULL,
                status TEXT NOT NULL,
                details_json TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            """
        )

    return db_path
