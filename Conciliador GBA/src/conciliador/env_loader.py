from __future__ import annotations

import os
from pathlib import Path

_LOADED_ENV_FILES: set[str] = set()


def _strip_env_value(raw: str) -> str:
    value = raw.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1]
    return value


def load_project_env(env_path: str | Path | None = None, *, override: bool = True) -> Path | None:
    path = Path(env_path) if env_path is not None else Path(__file__).resolve().parents[2] / ".env"
    path = path.resolve()
    if not path.exists() or not path.is_file():
        return None

    cache_key = str(path)
    if cache_key in _LOADED_ENV_FILES:
        return path

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key:
            continue
        value = _strip_env_value(value)
        if override or key not in os.environ:
            os.environ[key] = value

    _LOADED_ENV_FILES.add(cache_key)
    return path
