from __future__ import annotations

import asyncio
import os
import signal
import subprocess
import sys
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Iterable

import httpx
from fastapi import FastAPI, Request, Response

ROOT = Path(__file__).resolve().parent
GBA_PORT = int(os.getenv("GBA_INTERNAL_PORT", "8011"))
SALICE_PORT = int(os.getenv("SALICE_INTERNAL_PORT", "8012"))
DEFAULT_APP = os.getenv("DEFAULT_APP", "gba").strip().lower()

HOP_BY_HOP_HEADERS = {
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailer",
    "transfer-encoding",
    "upgrade",
}

processes: list[subprocess.Popen] = []
client: httpx.AsyncClient | None = None


def _split_hosts(raw: str, defaults: Iterable[str]) -> list[str]:
    values = [part.strip().lower() for part in raw.split(",") if part.strip()]
    return values or [value.lower() for value in defaults]


GBA_HOSTS = _split_hosts(os.getenv("GBA_HOSTS", ""), ["gba"])
SALICE_HOSTS = _split_hosts(os.getenv("SALICE_HOSTS", ""), ["salice"])


def _env_for_app(prefix: str) -> dict[str, str]:
    env = os.environ.copy()
    for key, value in os.environ.items():
        marker = f"{prefix}_"
        if key.startswith(marker) and value:
            env[key[len(marker):]] = value
    return env


def _start_app(name: str, cwd: Path, port: int, env_prefix: str) -> subprocess.Popen:
    cmd = [
        sys.executable,
        "-m",
        "uvicorn",
        "app:app",
        "--host",
        "127.0.0.1",
        "--port",
        str(port),
    ]
    return subprocess.Popen(cmd, cwd=str(cwd), env=_env_for_app(env_prefix))


async def _wait_until_ready(port: int, name: str) -> None:
    deadline = asyncio.get_running_loop().time() + 60
    url = f"http://127.0.0.1:{port}/health"
    async with httpx.AsyncClient(timeout=2.0) as probe:
        while True:
            try:
                response = await probe.get(url)
                if response.status_code < 500:
                    return
            except Exception:
                pass
            if asyncio.get_running_loop().time() > deadline:
                raise RuntimeError(f"{name} no inicio en el puerto {port}")
            await asyncio.sleep(0.5)


@asynccontextmanager
async def lifespan(app: FastAPI):
    global client
    processes.append(_start_app("gba", ROOT / "Conciliador GBA", GBA_PORT, "GBA"))
    processes.append(_start_app("salice", ROOT / "conciliador SALICE", SALICE_PORT, "SALICE"))
    await asyncio.gather(
        _wait_until_ready(GBA_PORT, "GBA"),
        _wait_until_ready(SALICE_PORT, "Salice"),
    )
    client = httpx.AsyncClient(timeout=None, follow_redirects=False)
    try:
        yield
    finally:
        if client is not None:
            await client.aclose()
            client = None
        for proc in processes:
            if proc.poll() is None:
                proc.send_signal(signal.SIGTERM)
        for proc in processes:
            try:
                proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                proc.kill()


app = FastAPI(title="Conciliador Salice y GBA Router", lifespan=lifespan)


def _target_from_host(host_header: str) -> tuple[str, int]:
    host = host_header.split(":", 1)[0].lower()
    if any(token in host for token in SALICE_HOSTS):
        return "salice", SALICE_PORT
    if any(token in host for token in GBA_HOSTS):
        return "gba", GBA_PORT
    if DEFAULT_APP == "salice":
        return "salice", SALICE_PORT
    return "gba", GBA_PORT


@app.get("/health")
async def health() -> dict[str, object]:
    return {
        "ok": True,
        "apps": {
            "gba": processes[0].poll() is None if len(processes) > 0 else False,
            "salice": processes[1].poll() is None if len(processes) > 1 else False,
        },
    }


@app.api_route("/{path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
async def proxy(path: str, request: Request) -> Response:
    if client is None:
        return Response("Router no inicializado", status_code=503)

    app_name, port = _target_from_host(request.headers.get("host", ""))
    body = await request.body()
    query = request.url.query
    target_url = f"http://127.0.0.1:{port}/{path}"
    if query:
        target_url += f"?{query}"

    headers = {
        key: value
        for key, value in request.headers.items()
        if key.lower() not in HOP_BY_HOP_HEADERS and key.lower() != "host"
    }
    headers["x-forwarded-host"] = request.headers.get("host", "")
    headers["x-forwarded-proto"] = request.url.scheme
    headers["x-conciliador-app"] = app_name

    upstream = await client.request(
        request.method,
        target_url,
        headers=headers,
        content=body,
    )
    response_headers = {
        key: value
        for key, value in upstream.headers.items()
        if key.lower() not in HOP_BY_HOP_HEADERS
    }
    return Response(
        content=upstream.content,
        status_code=upstream.status_code,
        headers=response_headers,
        media_type=upstream.headers.get("content-type"),
    )
