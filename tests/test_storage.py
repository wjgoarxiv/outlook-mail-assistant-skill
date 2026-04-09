import sqlite3

from outlook_mail_assistant.storage import bootstrap_workspace, initialize_database


def test_bootstrap_workspace_creates_expected_directories(tmp_path):
    workspace = bootstrap_workspace(tmp_path / "workspace")

    assert workspace.root.exists()
    assert workspace.raw.exists()
    assert workspace.normalized.exists()
    assert workspace.reports.exists()
    assert workspace.exports.exists()
    assert workspace.logs.exists()


def test_bootstrap_workspace_writes_manifest(tmp_path):
    workspace = bootstrap_workspace(tmp_path / "workspace")

    assert workspace.manifest_path.exists()
    manifest = workspace.manifest_path.read_text(encoding="utf-8")

    assert '"version": 1' in manifest
    assert '"workspace_name": "workspace"' in manifest


def test_initialize_database_creates_expected_tables(tmp_path):
    db_path = tmp_path / "mail-index.sqlite3"

    initialize_database(db_path)

    with sqlite3.connect(db_path) as connection:
        tables = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }

    assert {"messages", "tasks", "decisions", "audit_log"} <= tables


def test_initialize_database_is_idempotent(tmp_path):
    db_path = tmp_path / "mail-index.sqlite3"

    initialize_database(db_path)
    initialize_database(db_path)

    with sqlite3.connect(db_path) as connection:
        tables = list(
            connection.execute("SELECT name FROM sqlite_master WHERE type = 'table'")
        )

    assert tables
