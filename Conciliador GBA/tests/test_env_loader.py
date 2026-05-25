from __future__ import annotations

import os

from src.conciliador.env_loader import load_project_env


def test_load_project_env_reads_file_and_overrides_existing_env_by_default(monkeypatch, tmp_path):
    env_path = tmp_path / ".env"
    env_path.write_text(
        "\n".join(
            [
                "# comentario",
                "API_MODE_ENABLED=true",
                "GESI_API_USERNAME=archivo_user",
                "export GESI_API_PASSWORD='archivo_pass'",
            ]
        ),
        encoding="utf-8",
    )

    monkeypatch.delenv("API_MODE_ENABLED", raising=False)
    monkeypatch.setenv("GESI_API_USERNAME", "shell_user")
    monkeypatch.delenv("GESI_API_PASSWORD", raising=False)

    loaded = load_project_env(env_path)

    assert loaded == env_path.resolve()
    assert os.environ["API_MODE_ENABLED"] == "true"
    assert os.environ["GESI_API_USERNAME"] == "archivo_user"
    assert os.environ["GESI_API_PASSWORD"] == "archivo_pass"


def test_load_project_env_can_preserve_existing_env_when_override_is_false(monkeypatch, tmp_path):
    env_path = tmp_path / ".env"
    env_path.write_text("GESI_API_USERNAME=archivo_user\n", encoding="utf-8")

    monkeypatch.setenv("GESI_API_USERNAME", "shell_user")

    load_project_env(env_path, override=False)

    assert os.environ["GESI_API_USERNAME"] == "shell_user"
