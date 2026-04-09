from __future__ import annotations

import argparse
import json
from pathlib import Path

from outlook_mail_assistant.extraction import extract_task_candidates
from outlook_mail_assistant.report_exports import (
    export_markdown_to_docx,
    export_task_candidates_to_csv,
    export_task_candidates_to_markdown,
    export_task_candidates_to_xlsx,
)
from outlook_mail_assistant.summary_reports import build_mail_summary_markdown


def _read_jsonl_dir(path: Path):
    records = []
    for file in sorted(path.glob("*.jsonl")):
        with file.open("r", encoding="utf-8", errors="replace") as handle:
            for line in handle:
                if line.strip():
                    records.append(json.loads(line))
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="Export task candidate reports from JSONL exports.")
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--docx", action="store_true")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    records = _read_jsonl_dir(Path(args.input_dir))
    candidates = extract_task_candidates(records)

    md_path = export_task_candidates_to_markdown(candidates, output_dir / "task-candidates.md")
    export_task_candidates_to_csv(candidates, output_dir / "task-candidates.csv")
    export_task_candidates_to_xlsx(candidates, output_dir / "task-candidates.xlsx")

    summary_md = output_dir / "mail-summary.md"
    summary_md.write_text(
        build_mail_summary_markdown(
            records=records,
            candidates=candidates,
            title="Outlook Weekly and Monthly Summary Report",
        ),
        encoding="utf-8",
    )
    if args.docx:
        export_markdown_to_docx(summary_md, output_dir / "mail-summary.docx")

    print(f"task_candidates={len(candidates)}")


if __name__ == "__main__":
    main()
