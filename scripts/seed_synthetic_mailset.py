from __future__ import annotations

import argparse
import json
from pathlib import Path

from outlook_mail_assistant.export_pipeline import export_records_to_jsonl_chunks
from outlook_mail_assistant.synthetic_dataset import generate_synthetic_mailset


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a synthetic mail dataset from a base JSON export.")
    parser.add_argument("--input-json", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--target-count", type=int, required=True)
    parser.add_argument("--chunk-size", type=int, default=1000)
    args = parser.parse_args()

    base_records = json.loads(Path(args.input_json).read_text(encoding="utf-8"))
    generated = generate_synthetic_mailset(base_records, target_count=args.target_count)
    files = export_records_to_jsonl_chunks(
        generated,
        output_dir=Path(args.output_dir),
        chunk_size=args.chunk_size,
    )
    print(f"generated_records={len(generated)}")
    print(f"chunk_files={len(files)}")


if __name__ == "__main__":
    main()
