from datetime import datetime, timezone

from outlook_mail_assistant.outlook_com import OutlookComSession, import_outlook_messages


class StubSession:
    def list_messages(
        self,
        *,
        folder=None,
        limit=None,
        received_after=None,
        received_before=None,
        recursive=False,
        all_stores=False,
        store_scope="primary-shared",
        include_system=False,
        store_name=None,
    ):
        messages = [
            {
                "source_path": "Inbox/Quarterly review",
                "message_id": "live-001",
                "subject": "Quarterly review",
                "sender_email": "owner@example.com",
                "received_at": datetime(2026, 4, 2, 2, 0, tzinfo=timezone.utc),
            },
            {
                "source_path": "Inbox/Follow up",
                "message_id": "live-002",
                "subject": "Follow up",
                "sender_email": "manager@example.com",
                "received_at": datetime(2026, 4, 2, 3, 0, tzinfo=timezone.utc),
            },
        ]
        if received_after is not None:
            messages = [
                message
                for message in messages
                if message["received_at"] >= received_after
            ]
        if received_before is not None:
            messages = [
                message
                for message in messages
                if message["received_at"] < received_before
            ]
        if limit is not None:
            messages = messages[:limit]
        return messages


def test_import_outlook_messages_normalizes_live_records():
    records = import_outlook_messages(session=StubSession(), limit=1)

    assert len(records) == 1
    assert records[0]["source_type"] == "outlook-com"
    assert records[0]["message_id"] == "live-001"
    assert records[0]["received_at"] == "2026-04-02T02:00:00Z"


def test_import_outlook_messages_filters_to_messages_received_before_cutoff():
    records = import_outlook_messages(
        session=StubSession(),
        received_before=datetime(2026, 4, 2, 2, 30, tzinfo=timezone.utc),
    )

    assert len(records) == 1
    assert records[0]["message_id"] == "live-001"


class _FakeMessage:
    def __init__(self, subject, sender, received_at):
        self.EntryID = f"id-{subject}"
        self.Subject = subject
        self.SenderEmailAddress = sender
        self.ReceivedTime = received_at
        self.Body = f"Body for {subject}"
        self.HTMLBody = f"<p>Body for {subject}</p>"


class _FakeItems(list):
    pass


class _FakeFolder:
    def __init__(self, name, folders=None, items=None):
        self.Name = name
        self.Folders = folders or []
        self.Items = items or _FakeItems()


class _FakeStore:
    def __init__(self, name, root_folder, *, is_data_file_store=False, file_path=""):
        self.DisplayName = name
        self._root_folder = root_folder
        self.IsDataFileStore = is_data_file_store
        self.FilePath = file_path

    def GetRootFolder(self):
        return self._root_folder


class _FakeNamespace:
    def __init__(self):
        subfolder_items = _FakeItems(
            [
                _FakeMessage("Sub task", "sub@example.com", datetime(2026, 4, 2, 4, 0, tzinfo=timezone.utc)),
            ]
        )
        inbox_items = _FakeItems(
            [
                _FakeMessage("Quarterly review", "owner@example.com", datetime(2026, 4, 2, 2, 0, tzinfo=timezone.utc)),
                _FakeMessage("Follow up", "manager@example.com", datetime(2026, 4, 2, 3, 0, tzinfo=timezone.utc)),
            ]
        )
        self._default_folder = _FakeFolder(
            "Inbox",
            folders=[_FakeFolder("Subfolder", items=subfolder_items)],
            items=inbox_items,
        )
        shared_root = _FakeFolder(
            "SharedRoot",
            folders=[
                _FakeFolder(
                    "Inbox",
                    items=_FakeItems(
                        [
                            _FakeMessage(
                                "Shared follow up",
                                "shared@example.com",
                                datetime(2026, 4, 2, 5, 0, tzinfo=timezone.utc),
                            )
                        ]
                    ),
                ),
            ],
            items=_FakeItems(),
        )
        archive_root = _FakeFolder(
            "ArchiveRoot",
            folders=[_FakeFolder("Inbox", items=_FakeItems([_FakeMessage("Archived", "archive@example.com", datetime(2026, 4, 2, 6, 0, tzinfo=timezone.utc))]))],
            items=_FakeItems(),
        )
        self.Stores = [
            _FakeStore("Primary Mailbox", self._default_folder),
            _FakeStore("Shared Mailbox", shared_root),
            _FakeStore("Archive Store", archive_root, is_data_file_store=True, file_path="C:/archive.pst"),
        ]
        self.DefaultStore = self.Stores[0]

    def GetDefaultFolder(self, _folder_id):
        return self._default_folder


def test_outlook_com_session_lists_messages_from_default_inbox():
    session = OutlookComSession(namespace=_FakeNamespace())

    messages = session.list_messages(limit=1)

    assert len(messages) == 1
    assert messages[0]["source_path"] == "Inbox/Quarterly review"
    assert messages[0]["message_id"] == "id-Quarterly review"
    assert messages[0]["subject"] == "Quarterly review"
    assert messages[0]["sender_email"] == "owner@example.com"


def test_outlook_com_session_filters_messages_before_cutoff():
    session = OutlookComSession(namespace=_FakeNamespace())

    messages = session.list_messages(
        received_before=datetime(2026, 4, 2, 2, 30, tzinfo=timezone.utc)
    )

    assert len(messages) == 1
    assert messages[0]["message_id"] == "id-Quarterly review"


def test_outlook_com_session_lists_messages_recursively():
    session = OutlookComSession(namespace=_FakeNamespace())

    messages = session.list_messages(recursive=True)

    assert len(messages) == 3
    assert any(message["source_path"] == "Inbox/Subfolder/Sub task" for message in messages)


def test_list_stores_defaults_to_primary_and_shared_only():
    session = OutlookComSession(namespace=_FakeNamespace())

    stores = session.list_stores()

    assert [store["name"] for store in stores] == ["Primary Mailbox", "Shared Mailbox"]


def test_import_outlook_messages_can_scan_all_primary_and_shared_stores():
    records = import_outlook_messages(
        session=OutlookComSession(namespace=_FakeNamespace()),
        recursive=True,
        store_scope="primary-shared",
        all_stores=True,
    )

    ids = {record["message_id"] for record in records}
    assert "id-Shared follow up" in ids
    assert "id-Archived" not in ids
    assert all("store_name" in record for record in records)
    assert all("folder_path" in record for record in records)


def test_list_stores_all_scope_includes_archive_store():
    session = OutlookComSession(namespace=_FakeNamespace())

    stores = session.list_stores(store_scope="all")

    assert [store["name"] for store in stores] == [
        "Primary Mailbox",
        "Shared Mailbox",
        "Archive Store",
    ]


def test_import_outlook_messages_can_include_archive_store_when_scope_is_all():
    records = import_outlook_messages(
        session=OutlookComSession(namespace=_FakeNamespace()),
        recursive=True,
        store_scope="all",
        all_stores=True,
    )

    ids = {record["message_id"] for record in records}
    assert "id-Archived" in ids
