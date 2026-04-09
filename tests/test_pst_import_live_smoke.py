import os

import pytest

from outlook_mail_assistant.pst_import import parse_pst_messages


pytestmark = pytest.mark.skipif(
    os.getenv("OMA_ENABLE_LIVE_PST_TEST") != "1",
    reason="set OMA_ENABLE_LIVE_PST_TEST=1 and OMA_PST_PATH to run live PST smoke test",
)


def test_live_pst_smoke_extracts_at_least_one_message():
    pst_path = os.environ["OMA_PST_PATH"]
    records = parse_pst_messages(pst_path)

    assert records
    assert records[0]["source_type"] == "pst"
    assert records[0]["message_id"]
