import os

import pytest

from outlook_mail_assistant.outlook_com import OutlookComSession, import_outlook_messages


pytestmark = pytest.mark.skipif(
    os.getenv("OMA_ENABLE_LIVE_OUTLOOK_TEST") != "1",
    reason="set OMA_ENABLE_LIVE_OUTLOOK_TEST=1 to run live Outlook smoke test",
)


def test_live_outlook_smoke_extracts_at_least_one_message():
    folder = os.getenv("OMA_OUTLOOK_FOLDER")
    limit = int(os.getenv("OMA_OUTLOOK_LIMIT", "1"))

    session = OutlookComSession()
    records = import_outlook_messages(
        session=session,
        folder=folder,
        limit=limit,
        all_stores=True,
        store_scope="primary-shared",
        recursive=True,
    )

    assert records
    assert records[0]["source_type"] == "outlook-com"
    assert records[0]["message_id"]
    assert records[0]["subject"] is not None
