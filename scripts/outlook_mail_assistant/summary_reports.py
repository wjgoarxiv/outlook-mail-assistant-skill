from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime, timezone


def _parse_received(value: str | None):
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def _top_counts(records, key, limit=8):
    counts = Counter(record.get(key) or "unknown" for record in records)
    return counts.most_common(limit)


def _top_subjects(records, limit=8):
    counts = Counter(record.get("subject") or "Untitled" for record in records)
    return counts.most_common(limit)


def _build_period_stats(records, candidates, period_key):
    stats = defaultdict(lambda: {"messages": 0, "candidates": 0, "explicit": 0, "needs_confirmation": 0})

    for record in records:
        received = _parse_received(record.get("received_at"))
        if not received:
            continue
        stats[period_key(received)]["messages"] += 1

    for item in candidates:
        received = _parse_received(item.get("received_at"))
        if not received:
            continue
        bucket = stats[period_key(received)]
        bucket["candidates"] += 1
        if item.get("needs_confirmation"):
            bucket["needs_confirmation"] += 1
        else:
            bucket["explicit"] += 1
    return stats


def _section_summary_lines(label: str, period_stats):
    lines = [f"## {label}", "", f"| {label.split()[0]} | Messages | Candidates | Immediate Actions | Needs Confirmation |", "|---|---:|---:|---:|---:|"]
    for key in sorted(period_stats.keys(), reverse=True):
        stats = period_stats[key]
        lines.append(
            f"| {key} | {stats['messages']} | {stats['candidates']} | {stats['explicit']} | {stats['needs_confirmation']} |"
        )
    return lines


def _issue_lines(label: str, records, limit=8):
    lines = [f"## {label}", ""]
    for subject, count in _top_subjects(records, limit=limit):
        lines.append(f"- {subject} ({count})")
    if len(lines) == 2:
        lines.append("- No dominant issue clusters detected.")
    return lines


def _work_summary_lines(label: str, candidates):
    explicit = [item for item in candidates if not item.get("needs_confirmation")]
    inferred = [item for item in candidates if item.get("needs_confirmation")]
    kind_counts = Counter(item.get("kind") or "unknown" for item in candidates)

    lines = [
        f"## {label}",
        "",
        f"- Immediate actions identified: {len(explicit)}",
        f"- Items requiring confirmation: {len(inferred)}",
    ]
    for kind, count in sorted(kind_counts.items()):
        lines.append(f"- {kind}: {count}")
    return lines


def build_mail_summary_markdown(*, records, candidates, title: str) -> str:
    records = sorted(records, key=lambda item: item.get("received_at") or "", reverse=True)
    candidates = sorted(candidates, key=lambda item: item.get("received_at") or "", reverse=True)

    weekly = _build_period_stats(records, candidates, lambda dt: f"{dt.isocalendar().year}-W{dt.isocalendar().week:02d}")
    monthly = _build_period_stats(records, candidates, lambda dt: dt.strftime("%Y-%m"))

    top_senders = _top_counts(records, "sender_email", limit=6)
    top_folders = _top_counts(records, "folder_path", limit=6)
    kind_counts = Counter(item.get("kind") or "unknown" for item in candidates)

    weekly_records = records[: min(len(records), 50)]
    monthly_records = records
    weekly_candidates = candidates[: min(len(candidates), 50)]
    monthly_candidates = candidates

    lines = [
        f"# {title}",
        "",
        f"Generated: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}",
        "",
        "## Executive Summary",
        f"- Total messages reviewed: {len(records)}",
        f"- Action candidates identified: {len(candidates)}",
        f"- Immediate actions: {sum(1 for item in candidates if not item.get('needs_confirmation'))}",
        f"- Needs confirmation: {sum(1 for item in candidates if item.get('needs_confirmation'))}",
    ]
    for kind, count in sorted(kind_counts.items()):
        lines.append(f"- {kind}: {count}")

    lines.extend([""] + _section_summary_lines("Weekly Summary", weekly))
    lines.extend([""] + _issue_lines("Weekly Key Issues", weekly_records))
    lines.extend([""] + _work_summary_lines("Weekly Work Summary", weekly_candidates))
    lines.extend([""] + _section_summary_lines("Monthly Summary", monthly))
    lines.extend([""] + _issue_lines("Monthly Key Issues", monthly_records))
    lines.extend([""] + _work_summary_lines("Monthly Work Summary", monthly_candidates))

    lines.extend(
        [
            "",
            "## Priority Follow-Ups",
            "",
            "| Subject | Sender | Received | Kind | Confidence | Folder |",
            "|---|---|---|---|---|---|",
        ]
    )
    for item in candidates[:20]:
        confidence = "Needs confirmation" if item.get("needs_confirmation") else "Immediate action"
        lines.append(
            f"| {item['subject']} | {item['sender_email']} | {item['received_at']} | {item['kind']} | {confidence} | {item.get('folder_path') or '-'} |"
        )

    lines.extend(["", "## Communication Trends", "", "### Top Senders", ""])
    for sender, count in top_senders:
        lines.append(f"- {sender}: {count}")

    lines.extend(["", "### Top Folders", ""])
    for folder, count in top_folders:
        lines.append(f"- {folder}: {count}")

    lines.extend(
        [
            "",
            "## Notes",
            "- This report is optimized for management review and office-style distribution.",
            "- The workbook companion contains the full review queue and confirmation split.",
        ]
    )
    return "\n".join(lines) + "\n"
