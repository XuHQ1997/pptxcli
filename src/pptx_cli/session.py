from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any
from urllib import error, request

STATE_FILE_ENV = "PPTXCLI_STATE_FILE"
STATE_FILE_NAME = ".pptxcli-session.json"
SERVER_STARTUP_TIMEOUT = 8.0
REQUEST_TIMEOUT = 5.0


class SessionError(RuntimeError):
    pass


def resolve_state_file_path(cli_argv0: str | None = None) -> Path:
    env_path = os.environ.get(STATE_FILE_ENV)
    if env_path:
        return Path(env_path).expanduser().resolve()

    argv0 = cli_argv0 or sys.argv[0]
    if argv0:
        return Path(argv0).resolve().parent / STATE_FILE_NAME
    return Path.cwd() / STATE_FILE_NAME


def load_session_state(state_file: Path) -> dict[str, Any]:
    if not state_file.exists():
        raise SessionError(
            f"no active session state file found: {state_file}. Run `pptxcli init --origin_file <file>` first."
        )
    return json.loads(state_file.read_text(encoding="utf-8"))


def save_session_state(state_file: Path, payload: dict[str, Any]) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    temp_file = state_file.with_suffix(f"{state_file.suffix}.tmp")
    temp_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    temp_file.replace(state_file)


def remove_session_state(state_file: Path) -> None:
    try:
        state_file.unlink()
    except FileNotFoundError:
        return


def read_session_log_tail(state_file: Path, max_lines: int = 20) -> str:
    log_file = state_file.with_suffix(".log")
    if not log_file.exists():
        return ""
    lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
    return "\n".join(lines[-max_lines:])


def session_request(
    *,
    state_file: Path,
    method: str,
    route: str,
    payload: dict[str, Any] | None = None,
    timeout: float = REQUEST_TIMEOUT,
) -> dict[str, Any]:
    state = load_session_state(state_file)
    server_url = state.get("server_url")
    if not isinstance(server_url, str) or not server_url:
        raise SessionError(f"invalid session state file: {state_file}")

    url = server_url.rstrip("/") + route
    body = None
    headers = {"Accept": "application/json"}
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    req = request.Request(url=url, data=body, method=method, headers=headers)
    try:
        with request.urlopen(req, timeout=timeout) as response:
            content = response.read().decode("utf-8")
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise SessionError(f"session server returned HTTP {exc.code}: {detail}") from exc
    except error.URLError as exc:
        raise SessionError(
            f"session server is unavailable at {server_url}. Run `pptxcli init --origin_file <file>` again or clean stale state {state_file}."
        ) from exc

    if not content:
        return {}
    return json.loads(content)


def cleanup_stale_session_state(state_file: Path) -> bool:
    if not state_file.exists():
        return False
    try:
        session_request(state_file=state_file, method="GET", route="/health", timeout=1.0)
    except SessionError:
        remove_session_state(state_file)
        return True
    return False


def assert_no_active_session(state_file: Path) -> None:
    if not state_file.exists():
        return
    try:
        health = session_request(
            state_file=state_file,
            method="GET",
            route="/health",
            timeout=1.0,
        )
    except SessionError:
        remove_session_state(state_file)
        return

    origin_file = health.get("origin_file", "unknown")
    raise SessionError(
        f"an active session already exists for {origin_file}. Run `pptxcli finish` before starting another session."
    )


def start_session_server(
    *,
    origin_file: Path,
    state_file: Path,
) -> dict[str, Any]:
    assert_no_active_session(state_file)

    log_file = state_file.with_suffix(".log")
    env = os.environ.copy()
    src_root = Path(__file__).resolve().parents[1]
    existing_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = (
        f"{src_root}{os.pathsep}{existing_pythonpath}"
        if existing_pythonpath
        else str(src_root)
    )

    with log_file.open("w", encoding="utf-8") as log_handle:
        process = subprocess.Popen(
            [
                sys.executable,
                "-m",
                "pptx_cli.session_server",
                "--origin-file",
                str(origin_file),
                "--state-file",
                str(state_file),
            ],
            stdin=subprocess.DEVNULL,
            stdout=log_handle,
            stderr=log_handle,
            start_new_session=True,
            env=env,
        )

    deadline = time.time() + SERVER_STARTUP_TIMEOUT
    while time.time() < deadline:
        if state_file.exists():
            try:
                return session_request(
                    state_file=state_file,
                    method="GET",
                    route="/health",
                    timeout=1.0,
                )
            except SessionError:
                pass

        exit_code = process.poll()
        if exit_code is not None:
            remove_session_state(state_file)
            log_tail = read_session_log_tail(state_file)
            detail = log_tail or f"server process exited with code {exit_code}"
            raise SessionError(f"failed to start session server: {detail}")
        time.sleep(0.1)

    process.terminate()
    raise SessionError("session server startup timed out before becoming ready.")
