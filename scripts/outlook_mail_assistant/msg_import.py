from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .canonical_schema import CanonicalMessage


class MissingDependencyError(RuntimeError):
    """Raised when the .msg parser dependency is not available."""


@dataclass(slots=True)
class ExtractMsgParser:
    def open_message(self, path: Path):
        try:
            import extract_msg  # type: ignore
        except ModuleNotFoundError as exc:  # pragma: no cover - exercised via caller
            raise MissingDependencyError(
                "MSG ingestion requires the 'extract-msg' package."
            ) from exc
        return extract_msg.Message(str(path))


def _message_date(message) -> datetime | None:
    value = getattr(message, "date", None)
    if isinstance(value, datetime):
        return value
    return None


def parse_msg_message(path: Path, parser: ExtractMsgParser | None = None) -> dict[str, str | None]:
    path = Path(path)
    parser = parser or ExtractMsgParser()
    message = parser.open_message(path)

    canonical = CanonicalMessage(
        source_type="msg",
        source_path=str(path),
        message_id=str(path),
        subject=getattr(message, "subject", "") or path.stem,
        sender_email=getattr(message, "sender", "") or "unknown@example.invalid",
        received_at=_message_date(message),
    )
    return canonical.to_record()
