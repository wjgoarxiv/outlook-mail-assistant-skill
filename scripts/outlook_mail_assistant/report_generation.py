from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook

from .docx_export import convert_markdown_to_docx


REPORT_COLUMNS = [
    "message_id",
    "store_name",
    "folder_path",
    "subject",
    "sender_email",
    "received_at",
    "kind",
    "confidence",
    "reason",
    "snippet",
    "needs_confirmation",
]


def write_task_report_markdown(items, output_path: Path):
    output_path = Path(output_path)
    counts = {}
    for item in items:
        counts[item["kind"]] = counts.get(item["kind"], 0) + 1

    lines = [
        "# Task Report",
        "",
        f"- Total candidates: {len(items)}",
    ]
    for kind, count in sorted(counts.items()):
        lines.append(f"- {kind}: {count}")
    lines.extend(
        [
            "",
            "| message_id | kind | confidence | subject | sender_email | received_at | needs_confirmation |",
            "|---|---|---|---|---|---|---|",
        ]
    )
    for item in items:
        lines.append(
            "| {message_id} | {kind} | {confidence} | {subject} | {sender_email} | {received_at} | {needs_confirmation} |".format(
                **item
            )
        )
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output_path


def write_task_report_csv(items, output_path: Path):
    output_path = Path(output_path)
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        writer.writerows(items)
    return output_path


def write_task_report_xlsx(items, output_path: Path):
    output_path = Path(output_path)
    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:
        raise RuntimeError("Failed to create XLSX worksheet")
    sheet.title = "task_candidates"
    sheet.append(REPORT_COLUMNS)
    for item in items:
        sheet.append([item.get(column) for column in REPORT_COLUMNS])
    resolved = output_path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(str(resolved))
    workbook.close()
    return resolved


def write_task_report_docx(markdown_path: Path, output_path: Path):
    return convert_markdown_to_docx(markdown_path, output_path)
