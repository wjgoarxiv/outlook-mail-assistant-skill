from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import hashlib


def _to_utc_isoformat(value: datetime | str | None) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def canonical_dedupe_hash(
    *,
    message_id: str | None,
    subject: str,
    sender_email: str,
    received_at: str | None,
) -> str:
    payload = "||".join(
        [
            (message_id or "").strip().lower(),
            subject.strip(),
            sender_email.strip().lower(),
            (received_at or "").strip(),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


@dataclass(slots=True)
class CanonicalMessage:
    source_type: str
    source_path: str
    message_id: str
    subject: str
    sender_email: str
    received_at: datetime | str | None = None
    body_text: str | None = None
    store_name: str | None = None
    folder_path: str | None = None
    mailbox_root: str | None = None
    conversation_id: str | None = None
    sender_name: str | None = None
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    sent_at: datetime | str | None = None
    body_html: str | None = None
    attachment_manifest: list[dict] = field(default_factory=list)
    source_account: str | None = None
    source_folder: str | None = None
    imported_at: datetime | str | None = None

    def __post_init__(self) -> None:
        for name in ("source_type", "source_path", "message_id", "subject", "sender_email"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"{name} is required")

    def to_record(self) -> dict:
        payload = asdict(self)
        payload["received_at"] = _to_utc_isoformat(self.received_at)
        payload["sent_at"] = _to_utc_isoformat(self.sent_at)
        payload["imported_at"] = _to_utc_isoformat(self.imported_at)
        payload["dedupe_hash"] = canonical_dedupe_hash(
            message_id=self.message_id,
            subject=self.subject,
            sender_email=self.sender_email,
            received_at=payload["received_at"],
        )
        return payload
