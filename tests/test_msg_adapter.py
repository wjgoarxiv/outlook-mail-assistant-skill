from datetime import datetime, timezone
import types

import pytest

from outlook_mail_assistant.msg_adapter import load_msg_via_outlook


class _FakeItem:
    def __init__(self):
        self.EntryID = "entry-123"
        self.Subject = "MSG Subject"
        self.SenderEmailAddress = "sender@example.com"
        self.ReceivedTime = datetime(2026, 4, 2, 1, 30, tzinfo=timezone.utc)
        self.Body = "Plain body"
        self.HTMLBody = "<p>Plain body</p>"


class _FakeApplication:
    def OpenSharedItem(self, path):
        assert path.endswith(".msg")
        return _FakeItem()


def test_load_msg_via_outlook_returns_canonical_message():
    message = load_msg_via_outlook(
        "C:/mail/sample.msg",
        outlook_app=_FakeApplication(),
    )

    record = message.to_record()

    assert record["source_type"] == "msg"
    assert record["source_path"].endswith("sample.msg")
    assert record["message_id"] == "entry-123"
    assert record["subject"] == "MSG Subject"
    assert record["sender_email"] == "sender@example.com"
    assert record["received_at"] == "2026-04-02T01:30:00Z"


def test_load_msg_via_outlook_raises_clear_error_when_pywin32_missing(monkeypatch):
    fake_importlib = types.SimpleNamespace(find_spec=lambda _: None)

    monkeypatch.setattr("outlook_mail_assistant.msg_adapter.importlib.util", fake_importlib)

    with pytest.raises(RuntimeError, match="pywin32"):
        load_msg_via_outlook("C:/mail/sample.msg")
