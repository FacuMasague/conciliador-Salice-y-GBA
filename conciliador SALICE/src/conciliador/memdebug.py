from __future__ import annotations

import os
import resource
import time
from typing import Any, Dict, List, Optional


def is_mem_debug_enabled(mem_debug: Optional[bool] = None) -> bool:
    if mem_debug is not None:
        return bool(mem_debug)
    return os.getenv("CONCILIADOR_MEM_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}


def process_rss_mb() -> float:
    # macOS reports ru_maxrss in bytes; Linux reports KB.
    rss = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    if rss <= 0:
        return 0.0
    if rss > 10_000_000:
        return round(rss / (1024.0 * 1024.0), 2)
    return round(rss / 1024.0, 2)


def mem_debug_recorder(enabled: bool) -> tuple[List[Dict[str, Any]], Any]:
    start = time.perf_counter()
    stages: List[Dict[str, Any]] = []

    def record(stage: str, extra: Optional[Dict[str, Any]] = None) -> None:
        if not enabled:
            return
        payload: Dict[str, Any] = {
            "stage": stage,
            "rss_mb": process_rss_mb(),
            "elapsed_s": round(time.perf_counter() - start, 3),
        }
        if extra:
            payload.update(extra)
        stages.append(payload)

    return stages, record
