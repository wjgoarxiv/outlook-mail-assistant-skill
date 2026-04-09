from datetime import datetime, timezone

import pytest

from outlook_mail_assistant.msg_import import MissingDependencyError, parse_msg_message


class StubMessage:
    def __init__(self, *, subject="Quarterly review", sender="owner@example.com", body="Need follow-up", date=None):
        self.subject = subject
        self.sender = sender
        self.body = body
        self.date = date or datetime(2026, 4, 2, 1, 30, tzinfo=timezone.utc)


class StubParser:
    def __init__(self, message):
        self._message = message

    def open_message(self, path):
        return self._message


def test_parse_msg_message_maps_stub_parser_to_canonical_record(tmp_path):
    path = tmp_path / "sample.msg"
    path.write_text("stub", encoding="utf-8")

    record = parse_msg_message(path, parser=StubParser(StubMessage()))

    assert record["source_type"] == "msg"
    assert record["source_path"] == str(path)
    assert record["subject"] == "Quarterly review"
    assert record["sender_email"] == "owner@example.com"
    assert record["received_at"] == "2026-04-02T01:30:00Z"


def test_parse_msg_message_prefers_transport_date_when_present(tmp_path):
    path = tmp_path / "sample.msg"
    path.write_text("stub", encoding="utf-8")

    record = parse_msg_message(
        path,
        parser=StubParser(
            StubMessage(date=datetime(2026, 4, 3, 5, 0, tzinfo=timezone.utc))
        ),
    )

    assert record["received_at"] == "2026-04-03T05:00:00Z"


def test_parse_msg_message_raises_clear_error_when_dependency_missing(tmp_path):
    path = tmp_path / "sample.msg"
    path.write_text("stub", encoding="utf-8")

    with pytest.raises(MissingDependencyError):
        parse_msg_message(path)
