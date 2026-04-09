from datetime import datetime, timezone

from outlook_mail_assistant.outlook_com import OutlookComSession


class _FakeMessage:
    def __init__(self, subject, sender):
        self.EntryID = f"id-{subject}"
        self.Subject = subject
        self.SenderEmailAddress = sender
        self.ReceivedTime = datetime(2026, 4, 2, 2, 0, tzinfo=timezone.utc)
        self.Body = f"Body for {subject}"
        self.HTMLBody = f"<p>Body for {subject}</p>"


class _FakeFolder:
    def __init__(self, name, folders=None, items=None):
        self.Name = name
        self.Folders = folders or []
        self.Items = items or []


class _FakeStore:
    def __init__(self, name, *, is_archive=False, root_folder=None):
        self.DisplayName = name
        self.FilePath = f"C:/mail/{name}.pst" if is_archive else ""
        self.ExchangeStoreType = 3 if is_archive else 0
        self._root_folder = root_folder or _FakeFolder(name)

    def GetRootFolder(self):
        return self._root_folder


class _FakeNamespace:
    def __init__(self):
        primary_root = _FakeFolder(
            "Primary",
            folders=[_FakeFolder("Inbox", items=[_FakeMessage("Primary task", "p@example.com")])],
        )
        shared_root = _FakeFolder(
            "Shared Team",
            folders=[_FakeFolder("Inbox", items=[_FakeMessage("Shared task", "s@example.com")])],
        )
        archive_root = _FakeFolder(
            "Archive",
            folders=[_FakeFolder("Inbox", items=[_FakeMessage("Archived task", "a@example.com")])],
        )
        self.Stores = [
            _FakeStore("Primary", root_folder=primary_root),
            _FakeStore("Shared Team", root_folder=shared_root),
            _FakeStore("Archive", is_archive=True, root_folder=archive_root),
        ]

    def GetDefaultFolder(self, _folder_id):
        return self.Stores[0].GetRootFolder().Folders[0]


def test_list_stores_defaults_to_primary_and_shared_only():
    session = OutlookComSession(namespace=_FakeNamespace())

    stores = session.list_stores()

    assert [store["store_name"] for store in stores] == ["Primary", "Shared Team"]


def test_list_messages_all_stores_excludes_archive_by_default():
    session = OutlookComSession(namespace=_FakeNamespace())

    messages = session.list_messages(recursive=True, all_stores=True)

    subjects = {message["subject"] for message in messages}
    assert "Primary task" in subjects
    assert "Shared task" in subjects
    assert "Archived task" not in subjects
