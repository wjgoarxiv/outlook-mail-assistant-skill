from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3
import time
import tracemalloc

from outlook_mail_assistant.export_pipeline import (
    append_messages_to_db,
    dedupe_records,
    export_records_to_jsonl_chunks,
)
from outlook_mail_assistant.storage import initialize_database


def _read_jsonl_dir(path: Path):
    records = []
    for file in sorted(path.glob("*.jsonl")):
        with file.open("r", encoding="utf-8", errors="replace") as handle:
            for line in handle:
                if line.strip():
                    records.append(json.loads(line))
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark JSONL export and optional SQLite ingestion.")
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--chunk-size", type=int, default=1000)
    parser.add_argument("--sqlite", action="store_true")
    args = parser.parse_args()

    records = _read_jsonl_dir(Path(args.input_dir))
    tracemalloc.start()
    started = time.perf_counter()
    deduped, duplicate_skipped = dedupe_records(records)
    files = export_records_to_jsonl_chunks(deduped, output_dir=Path(args.output_dir), chunk_size=args.chunk_size)
    sqlite_inserted = 0
    if args.sqlite:
        db_path = initialize_database(Path(args.output_dir) / "mail-index.sqlite3")
        connection = sqlite3.connect(db_path)
        try:
            sqlite_inserted = append_messages_to_db(connection, deduped)
        finally:
            connection.close()
    elapsed = time.perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    result = {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "records": len(deduped),
        "duplicate_skipped": duplicate_skipped,
        "chunk_size": args.chunk_size,
        "chunk_files": len(files),
        "sqlite": args.sqlite,
        "sqlite_inserted": sqlite_inserted,
        "elapsed_seconds": round(elapsed, 4),
        "mails_per_second": round(len(deduped) / elapsed, 2) if elapsed else None,
        "peak_memory_mb": round(peak / (1024 * 1024), 2),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
