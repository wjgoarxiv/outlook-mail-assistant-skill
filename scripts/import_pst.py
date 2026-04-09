from __future__ import annotations

import argparse
from pathlib import Path

from outlook_mail_assistant.export_pipeline import write_export_artifacts
from outlook_mail_assistant.pst_import import parse_pst_messages


def main() -> None:
    parser = argparse.ArgumentParser(description="Import PST into chunked JSONL and optional SQLite.")
    parser.add_argument("--pst", required=True)
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--chunk-size", type=int, default=1000)
    parser.add_argument("--sqlite", action="store_true")
    parser.add_argument("--limit", type=int)
    args = parser.parse_args()

    records = parse_pst_messages(Path(args.pst))
    if args.limit is not None:
        records = records[: args.limit]
    result = write_export_artifacts(
        workspace_root=Path(args.workspace),
        records=records,
        config={
            "source_type": "pst",
            "pst_path": args.pst,
            "limit": args.limit,
        },
        chunk_size=args.chunk_size,
        sqlite_enabled=args.sqlite,
    )
    print(f"exported_records={result['summary']['exported_records']}")
    print(f"chunk_files={result['summary']['chunk_files']}")
    print(f"duplicate_skipped={result['summary']['duplicate_skipped']}")


if __name__ == "__main__":
    main()
