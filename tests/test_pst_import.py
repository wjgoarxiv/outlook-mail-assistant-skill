from datetime import datetime, timezone

import pytest

from outlook_mail_assistant.pst_import import (
    AsposePstParser,
    LibratomParser,
    MissingDependencyError,
    parse_pst_messages,
)


class StubPstParser:
    def iter_messages(self, path):
        return [
            {
                "source_path": str(path),
                "message_id": "pst-001",
                "subject": "Archived note",
                "sender_email": "archive@example.com",
                "received_at": datetime(2026, 3, 1, 8, 0, tzinfo=timezone.utc),
                "body_text": "Archived body",
                "store_name": "archive",
                "folder_path": "Inbox",
            }
        ]


def test_parse_pst_messages_normalizes_parser_output(tmp_path):
    pst_path = tmp_path / "archive.pst"
    pst_path.write_text("stub", encoding="utf-8")

    records = parse_pst_messages(pst_path, parser=StubPstParser())

    assert len(records) == 1
    assert records[0]["source_type"] == "pst"
    assert records[0]["message_id"] == "pst-001"
    assert records[0]["received_at"] == "2026-03-01T08:00:00Z"
    assert records[0]["body_text"] == "Archived body"
    assert records[0]["store_name"] == "archive"


def test_parse_pst_messages_raises_clear_error_without_dependency(tmp_path):
    pst_path = tmp_path / "archive.pst"
    pst_path.write_text("stub", encoding="utf-8")

    with pytest.raises(MissingDependencyError):
        LibratomParser().iter_messages(pst_path)


class _FakeInfo:
    def __init__(self):
        self.entry_id = "entry-id"
        self.entry_id_string = "entry-id-string"
        self.subject = "PST Subject"


class _FakeMapiMessage:
    def __init__(self):
        self.subject = "PST Subject"
        self.sender_email_address = "/EX"
        self.sender_smtp_address = "pst@example.com"
        self.delivery_time = datetime(2026, 3, 2, 9, 0, tzinfo=timezone.utc)
        self.body = "PST body"
        self.body_html = "<p>PST body</p>"


class _FakeFolder:
    def __init__(self, name, children=None):
        self.display_name = name
        self._children = children or []

    def enumerate_folders(self):
        return self._children

    def enumerate_messages(self):
        return [_FakeInfo()]

    def retrieve_full_path(self):
        return self.display_name


class _FakePersonalStorage:
    def __init__(self):
        self.root_folder = _FakeFolder("Root", children=[_FakeFolder("Inbox")])

    def extract_message(self, _entry_id):
        return _FakeMapiMessage()


def test_aspose_pst_parser_iter_messages_yields_contract_payload(monkeypatch, tmp_path):
    pst_path = tmp_path / "archive.pst"
    pst_path.write_text("stub", encoding="utf-8")

    monkeypatch.setattr(AsposePstParser, "_open_personal_storage", lambda self, _ae, _path: _FakePersonalStorage())
    parser = AsposePstParser()
    payloads = list(parser.iter_messages(pst_path))

    assert payloads
    assert payloads[0]["subject"] == "PST Subject"
    assert payloads[0]["sender_email"] == "pst@example.com"
    assert payloads[0]["body_text"] == "PST body"
