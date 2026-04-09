from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize benchmark result JSON files into markdown.")
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-md", required=True)
    args = parser.parse_args()

    rows = []
    for file in sorted(Path(args.input_dir).glob("*.json")):
        rows.append(json.loads(file.read_text(encoding="utf-8")))

    lines = ["# Benchmark Results", "", "| Records | Metric | Value |", "|---|---:|---:|"]
    for row in rows:
        record_count = row.get("records", row.get("exported_records"))
        if record_count is None:
            continue
        if "mails_per_second" in row:
            lines.append(f"| {record_count} | mails_per_second | {row['mails_per_second']} |")
        if "records_per_second" in row:
            lines.append(f"| {record_count} | records_per_second | {row['records_per_second']} |")
        if "peak_memory_mb" in row:
            lines.append(f"| {record_count} | peak_memory_mb | {row['peak_memory_mb']} |")

    Path(args.output_md).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(args.output_md)


if __name__ == "__main__":
    main()
