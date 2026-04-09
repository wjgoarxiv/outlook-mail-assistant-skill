from __future__ import annotations

import tracemalloc


def measure_peak_memory_mb(fn) -> float:
    tracemalloc.start()
    try:
        fn()
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    return peak / (1024 * 1024)
