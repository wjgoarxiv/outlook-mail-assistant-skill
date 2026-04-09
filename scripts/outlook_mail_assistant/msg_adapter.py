from __future__ import annotations

import importlib.util

from .canonical_schema import CanonicalMessage


def _get_outlook_application():
    if importlib.util.find_spec("win32com.client") is None:
        raise RuntimeError("pywin32 is required for Outlook COM access")

    import win32com.client  # type: ignore

    return win32com.client.Dispatch("Outlook.Application")


def load_msg_via_outlook(path: str, outlook_app=None) -> CanonicalMessage:
    app = outlook_app or _get_outlook_application()
    item = app.OpenSharedItem(path)

    return CanonicalMessage(
        source_type="msg",
        source_path=path,
        message_id=str(getattr(item, "EntryID", path)),
        subject=str(getattr(item, "Subject", "")),
        sender_email=str(getattr(item, "SenderEmailAddress", "")),
        received_at=getattr(item, "ReceivedTime", None),
        body_text=getattr(item, "Body", None),
        body_html=getattr(item, "HTMLBody", None),
    )
