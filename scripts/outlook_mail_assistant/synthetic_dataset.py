from __future__ import annotations

import copy


def generate_synthetic_mailset(base_records, *, target_count: int):
    if not base_records:
        return []

    generated = []
    index = 0
    while len(generated) < target_count:
        source = copy.deepcopy(base_records[index % len(base_records)])
        ordinal = len(generated) + 1
        source["source_type"] = "synthetic"
        source["source_path"] = f"synthetic/Message {ordinal}"
        source["message_id"] = f"synthetic-{ordinal:05d}"
        source["dedupe_hash"] = f"synthetic-hash-{ordinal:05d}"
        source["subject"] = f"{source.get('subject', 'Synthetic message')} #{ordinal}"
        generated.append(source)
        index += 1
    return generated
