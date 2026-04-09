from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import time
import tracemalloc

from outlook_mail_assistant.export_pipeline import write_export_artifacts
from outlook_mail_assistant.outlook_com import OutlookComSession, import_outlook_messages


def _parse_iso_datetime(raw: str | None):
    if not raw:
        return None
    value = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark live Outlook export.")
    parser.add_argument("--workspace", required=True)
    parser.add_argument("--all-stores", action="store_true")
    parser.add_argument("--store")
    parser.add_argument("--folder")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--received-after")
    parser.add_argument("--received-before")
    parser.add_argument("--chunk-size", type=int, default=1000)
    parser.add_argument("--sqlite", action="store_true")
    parser.add_argument("--recursive", action="store_true")
    parser.add_argument("--include-system", action="store_true")
    parser.add_argument(
        "--store-scope",
        choices=("primary-shared", "all", "primary-only"),
        default="primary-shared",
    )
    args = parser.parse_args()

    tracemalloc.start()
    started = time.perf_counter()
    records = import_outlook_messages(
        session=OutlookComSession(),
        store_name=args.store,
        folder=args.folder,
        limit=args.limit,
        received_after=_parse_iso_datetime(args.received_after),
        received_before=_parse_iso_datetime(args.received_before),
        recursive=args.recursive,
        include_system=args.include_system,
        store_scope=args.store_scope,
        all_stores=args.all_stores,
    )
    elapsed = time.perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    store_names = sorted({record.get("store_name") for record in records if record.get("store_name")})
    folder_paths = sorted({record.get("folder_path") for record in records if record.get("folder_path")})
    result = write_export_artifacts(
        workspace_root=Path(args.workspace),
        records=records,
        config={
            "all_stores": args.all_stores,
            "store_scope": args.store_scope,
            "store": args.store,
            "folder": args.folder,
            "recursive": args.recursive,
            "received_after": args.received_after,
            "received_before": args.received_before,
            "stores_scanned": len(store_names),
            "folders_scanned": len(folder_paths),
            "excluded_folders": 0 if args.include_system else 3,
            "elapsed_seconds": round(elapsed, 4),
            "peak_memory_mb": round(peak / (1024 * 1024), 2),
        },
        chunk_size=args.chunk_size,
        sqlite_enabled=args.sqlite,
    )
    summary = dict(result["summary"])
    summary["peak_memory_mb"] = round(peak / (1024 * 1024), 2)
    summary["mails_per_second"] = round(summary["exported_records"] / elapsed, 2) if elapsed else None
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
