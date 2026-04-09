from pathlib import Path

from execute_outlook_actions import _apply_action_overrides, _read_candidates


def test_apply_action_overrides_injects_calendar_times():
    item = {
        "subject": "Weekly meeting reminder",
        "kind": "meeting",
    }

    updated = _apply_action_overrides(
        item,
        start_at="2026-04-03T10:00:00",
        end_at="2026-04-03T11:00:00",
    )

    assert updated["start_at"] == "2026-04-03T10:00:00"
    assert updated["end_at"] == "2026-04-03T11:00:00"


def test_read_candidates_reads_csv(tmp_path: Path):
    csv_path = tmp_path / "candidates.csv"
    csv_path.write_text(
        "message_id,subject,kind,needs_confirmation\nm1,Weekly meeting reminder,meeting,False\n",
        encoding="utf-8",
    )

    rows = _read_candidates(csv_path)

    assert len(rows) == 1
    assert rows[0]["message_id"] == "m1"
