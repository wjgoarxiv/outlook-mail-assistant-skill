from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import tempfile

from .canonical_schema import CanonicalMessage


class MissingDependencyError(RuntimeError):
    """Raised when no usable .pst parser dependency is available."""


@dataclass(slots=True)
class LibratomParser:
    def iter_messages(self, path: Path):
        try:
            import libratom  # type: ignore  # noqa: F401
        except Exception as exc:  # pragma: no cover - environment-specific
            raise MissingDependencyError(
                "PST ingestion via libratom is unavailable in this environment. "
                "On Python 3.13, prefer installing Aspose.Email-for-Python-via-NET as the active fallback."
            ) from exc
        raise MissingDependencyError(
            "Libratom support is not wired yet in this environment. Aspose fallback should be used."
        )


@dataclass(slots=True)
class AsposePstParser:
    def iter_messages(self, path: Path):
        try:
            import aspose.email as ae  # type: ignore
        except Exception as exc:  # pragma: no cover - environment-specific
            raise MissingDependencyError(
                "PST ingestion requires libratom or Aspose.Email-for-Python-via-NET."
            ) from exc

        temp_copy = None
        try:
            pst = self._open_personal_storage(ae, path)
            root = pst.root_folder
            for folder in self._walk_folders(root):
                folder_name = str(getattr(folder, "display_name", "") or getattr(folder, "DisplayName", "") or "")
                folder_path = self._safe_folder_path(folder)
                for info in folder.enumerate_messages():
                    message = pst.extract_message(info.entry_id)
                    sender_email = (
                        getattr(message, "sender_smtp_address", None)
                        or getattr(message, "sender_email_address", None)
                        or "unknown@example.invalid"
                    )
                    yield {
                        "source_path": f"{path}:{folder_path}/{info.subject or info.entry_id_string}",
                        "message_id": str(info.entry_id_string or info.entry_id),
                        "subject": str(info.subject or "(no subject)"),
                        "sender_email": str(sender_email),
                        "received_at": getattr(message, "delivery_time", None),
                        "body_text": getattr(message, "body", None),
                        "body_html": getattr(message, "body_html", None),
                        "store_name": path.stem,
                        "folder_path": folder_path or folder_name,
                        "mailbox_root": str(getattr(root, "display_name", "") or path.stem),
                        "source_account": path.stem,
                    }
        finally:
            if temp_copy and temp_copy.exists():
                try:
                    temp_copy.unlink()
                except OSError:
                    pass

    def _open_personal_storage(self, ae, path: Path):
        try:
            return ae.storage.pst.PersonalStorage.from_file(str(path))
        except RuntimeError as exc:
            if "being used by another process" not in str(exc):
                raise
            temp_dir = Path(tempfile.mkdtemp(prefix="oma-pst-"))
            temp_copy = temp_dir / path.name
            shutil.copy2(path, temp_copy)
            return ae.storage.pst.PersonalStorage.from_file(str(temp_copy))

    def _walk_folders(self, folder):
        yield folder
        for child in folder.enumerate_folders():
            yield from self._walk_folders(child)

    @staticmethod
    def _safe_folder_path(folder) -> str:
        path = getattr(folder, "retrieve_full_path", None)
        if callable(path):
            try:
                return str(path())
            except Exception:
                return str(getattr(folder, "display_name", "") or "")
        return str(getattr(folder, "display_name", "") or "")


def choose_pst_parser():
    try:
        import libratom  # type: ignore  # noqa: F401
        return LibratomParser()
    except Exception:
        pass
    try:
        import aspose.email  # type: ignore  # noqa: F401
        return AsposePstParser()
    except Exception as exc:
        raise MissingDependencyError(
            "PST ingestion requires libratom or Aspose.Email-for-Python-via-NET."
        ) from exc


def parse_pst_messages(path: Path, parser=None) -> list[dict[str, str | None]]:
    path = Path(path)
    parser = parser or choose_pst_parser()

    records = []
    for payload in parser.iter_messages(path):
        canonical = CanonicalMessage(
            source_type="pst",
            source_path=payload.get("source_path", str(path)),
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
        records.append(canonical.to_record())
    return records
