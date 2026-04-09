from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import PurePosixPath
import importlib.util

from .canonical_schema import CanonicalMessage


@dataclass(slots=True)
class OutlookComSession:
    namespace: object | None = None

    def list_stores(
        self,
        *,
        store_scope: str = "primary-shared",
        include_system: bool = False,
    ):
        namespace = self.namespace or _get_outlook_namespace()
        raw_stores = list(getattr(namespace, "Stores", []))
        stores = []

        for index, store in enumerate(raw_stores):
            file_path = str(getattr(store, "FilePath", "") or "")
            is_archive = file_path.lower().endswith(".pst")

            if store_scope == "primary-only" and index > 0:
                continue
            if store_scope == "primary-shared" and is_archive:
                continue

            stores.append(
                {
                    "name": str(getattr(store, "DisplayName", f"Store {index + 1}")),
                    "store_name": str(getattr(store, "DisplayName", f"Store {index + 1}")),
                    "store": store,
                    "file_path": file_path,
                    "is_archive": is_archive,
                    "include_system": include_system,
                }
            )

        return stores

    def list_messages(
        self,
        *,
        store_name: str | None = None,
        folder: str | None = None,
        limit: int | None = None,
        received_after: datetime | None = None,
        received_before: datetime | None = None,
        recursive: bool = False,
        all_stores: bool = False,
        store_scope: str = "primary-shared",
        include_system: bool = False,
    ):
        messages = []
        stores = self.list_stores(store_scope=store_scope, include_system=include_system)
        if store_name is not None:
            stores = [store for store in stores if store["store_name"] == store_name]
            if not stores:
                raise RuntimeError(f"Outlook store not found: {store_name}")
        if not all_stores:
            stores = stores[:1]

        for store_info in stores:
            target_folder = self._resolve_folder(folder, store_info=store_info)
            for folder_path, item in self._iter_folder_items(
                target_folder,
                recursive=recursive,
                include_system=include_system,
            ):
                received_at = getattr(item, "ReceivedTime", None)
                if received_at is not None and received_at.tzinfo is None:
                    received_at = received_at.replace(tzinfo=timezone.utc)

                floor = received_after
                if floor is not None and floor.tzinfo is None:
                    floor = floor.replace(tzinfo=timezone.utc)
                if floor is not None and received_at is not None and received_at < floor:
                    continue

                cutoff = received_before
                if cutoff is not None and cutoff.tzinfo is None:
                    cutoff = cutoff.replace(tzinfo=timezone.utc)
                if cutoff is not None and received_at is not None and received_at >= cutoff:
                    continue

                subject = str(getattr(item, "Subject", "") or "").strip() or "(no subject)"
                entry_id = str(getattr(item, "EntryID", "") or "")
                sender_email = str(getattr(item, "SenderEmailAddress", "") or "")
                source_path = str(folder_path / (subject or entry_id or "message"))
                messages.append(
                    {
                        "source_path": source_path,
                        "message_id": entry_id or source_path,
                        "subject": subject,
                        "sender_email": sender_email or "unknown@local",
                        "received_at": received_at,
                        "body_text": getattr(item, "Body", None),
                        "body_html": getattr(item, "HTMLBody", None),
                        "store_name": store_info["store_name"],
                        "folder_path": str(folder_path),
                        "mailbox_root": str(getattr(target_folder, "Name", "")),
                        "source_account": store_info["store_name"],
                    }
                )
                if limit is not None and len(messages) >= limit:
                    return messages

        return messages

    def _resolve_folder(self, folder: str | None, *, store_info=None):
        namespace = self.namespace or _get_outlook_namespace()
        if store_info is None:
            current = namespace.GetDefaultFolder(6)
        else:
            current = store_info["store"].GetRootFolder()
        if not folder:
            return current
        for part in [piece for piece in folder.split("/") if piece]:
            current = self._find_child_folder(current, part)
        return current

    @staticmethod
    def _find_child_folder(parent, name: str):
        for child in getattr(parent, "Folders", []):
            if str(getattr(child, "Name", "")) == name:
                return child
        raise RuntimeError(f"Outlook folder not found: {name}")

    def _iter_folder_items(self, folder, *, recursive: bool, include_system: bool):
        folder_path = PurePosixPath(str(getattr(folder, "Name", "Inbox")))
        for item in getattr(folder, "Items", []):
            yield folder_path, item
        if recursive:
            for child in getattr(folder, "Folders", []):
                child_name = str(getattr(child, "Name", "Folder"))
                if not include_system and self._is_system_folder(child_name):
                    continue
                for child_path, item in self._iter_child_items(
                    child,
                    folder_path / child_name,
                    include_system=include_system,
                ):
                    yield child_path, item

    def _iter_child_items(self, folder, folder_path: PurePosixPath, *, include_system: bool):
        for item in getattr(folder, "Items", []):
            yield folder_path, item
        for child in getattr(folder, "Folders", []):
            child_name = str(getattr(child, "Name", "Folder"))
            if not include_system and self._is_system_folder(child_name):
                continue
            for nested_path, item in self._iter_child_items(
                child,
                folder_path / child_name,
                include_system=include_system,
            ):
                yield nested_path, item

    @staticmethod
    def _is_system_folder(name: str) -> bool:
        normalized = name.strip().lower()
        return normalized in {
            "deleted items",
            "junk email",
            "sync issues",
            "conversation history",
            "rss feeds",
        }


def _get_outlook_namespace():
    if importlib.util.find_spec("win32com.client") is None:
        raise RuntimeError("pywin32 is required for Outlook COM access")

    import win32com.client  # type: ignore

    application = win32com.client.Dispatch("Outlook.Application")
    return application.GetNamespace("MAPI")


def _normalize_live_message(payload: dict) -> dict[str, str | None]:
    canonical = CanonicalMessage(
        source_type="outlook-com",
        source_path=payload["source_path"],
        message_id=payload["message_id"],
        subject=payload["subject"],
        sender_email=payload["sender_email"],
        received_at=payload.get("received_at"),
        body_text=payload.get("body_text"),
        body_html=payload.get("body_html"),
        store_name=payload.get("store_name"),
        folder_path=payload.get("folder_path"),
        mailbox_root=payload.get("mailbox_root"),
        source_account=payload.get("source_account"),
    )
    return canonical.to_record()


def import_outlook_messages(
    *,
    session: OutlookComSession,
    store_name: str | None = None,
    folder: str | None = None,
    limit: int | None = None,
    received_after: datetime | None = None,
    received_before: datetime | None = None,
    recursive: bool = False,
    all_stores: bool = False,
    store_scope: str = "primary-shared",
    include_system: bool = False,
) -> list[dict[str, str | None]]:
    return [
        _normalize_live_message(message)
        for message in session.list_messages(
            store_name=store_name,
            folder=folder,
            limit=limit,
            received_after=received_after,
            received_before=received_before,
            recursive=recursive,
            all_stores=all_stores,
            store_scope=store_scope,
            include_system=include_system,
        )
    ]
