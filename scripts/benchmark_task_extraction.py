from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import time
import tracemalloc

from outlook_mail_assistant.extraction import extract_task_candidates


def _read_jsonl_dir(path: Path):
    records = []
    for file in sorted(path.glob("*.jsonl")):
        with file.open("r", encoding="utf-8", errors="replace") as handle:
            for line in handle:
                if line.strip():
                    records.append(json.loads(line))
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark task extraction from exported JSONL records.")
    parser.add_argument("--input-dir", required=True)
    args = parser.parse_args()

    records = _read_jsonl_dir(Path(args.input_dir))
    tracemalloc.start()
    started = time.perf_counter()
    tasks = extract_task_candidates(records)
    elapsed = time.perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    result = {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "records": len(records),
        "task_candidates": len(tasks),
        "elapsed_seconds": round(elapsed, 4),
        "records_per_second": round(len(records) / elapsed, 2) if elapsed else None,
        "peak_memory_mb": round(peak / (1024 * 1024), 2),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
