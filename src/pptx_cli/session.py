from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from copy import deepcopy
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
    payload = load_session_state_if_exists(state_file)
    if payload is None:
        raise SessionError(
            f"no active session state file found: {state_file}. Start with `pptxcli template create --from <file> --name <name>` first."
        )
    return payload


def load_session_state_if_exists(state_file: Path) -> dict[str, Any] | None:
    if not state_file.exists():
        return None
    payload = json.loads(state_file.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise SessionError(f"expected JSON object in session state file: {state_file}")
    return payload


def save_session_state(state_file: Path, payload: dict[str, Any]) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    temp_file = state_file.with_suffix(f"{state_file.suffix}.tmp")
    temp_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    temp_file.replace(state_file)


def merge_session_state(state_file: Path, updates: dict[str, Any]) -> dict[str, Any]:
    current = load_session_state_if_exists(state_file) or {}
    merged = deepcopy(current)
    merged.update(updates)
    save_session_state(state_file, merged)
    return merged


def update_session_mode(
    state_file: Path,
    *,
    mode: str,
    active_template: str | None = None,
    edit_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    updates: dict[str, Any] = {"mode": mode}
    if active_template is not None:
        updates["active_template"] = active_template
    if edit_context is not None:
        updates["edit_context"] = edit_context
    elif "edit_context" in (load_session_state_if_exists(state_file) or {}):
        updates["edit_context"] = None
    return merge_session_state(state_file, updates)


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
        mode = str(state.get("mode", "idle")).strip() or "idle"
        raise SessionError(
            f"current session mode is `{mode}` and has no live session server. Start a new workflow with `pptxcli template create --from <file> --name <name>`."
        )

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
            f"session server is unavailable at {server_url}. Re-run `pptxcli template create --from <file> --name <name>` or clean stale state {state_file}."
        ) from exc

    if not content:
        return {}
    return json.loads(content)


def cleanup_stale_session_state(state_file: Path) -> bool:
    if not state_file.exists():
        return False
    state = load_session_state_if_exists(state_file)
    if not state:
        remove_session_state(state_file)
        return True
    server_url = state.get("server_url")
    if not isinstance(server_url, str) or not server_url:
        return False
    try:
        session_request(state_file=state_file, method="GET", route="/health", timeout=1.0)
    except SessionError:
        cleanup_session_artifacts(state_file)
        remove_session_state(state_file)
        return True
    return False


def assert_no_active_session(state_file: Path) -> None:
    if not state_file.exists():
        return
    state = load_session_state_if_exists(state_file)
    if not state:
        remove_session_state(state_file)
        return
    server_url = state.get("server_url")
    if not isinstance(server_url, str) or not server_url:
        mode = str(state.get("mode", "idle")).strip() or "idle"
        raise SessionError(
            f"an active local session already exists in mode `{mode}`. Reuse it or wait for timeout before starting another session."
        )
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
        f"an active session already exists for {origin_file}. Reuse it or wait for timeout before starting another session."
    )


def cleanup_session_artifacts(state_file: Path) -> None:
    state = load_session_state_if_exists(state_file)
    if not state:
        return
    edit_context = state.get("edit_context")
    if not isinstance(edit_context, dict):
        return
    working_path = edit_context.get("working_pptx_path")
    if not isinstance(working_path, str) or not working_path.strip():
        return
    try:
        Path(working_path).resolve().unlink()
    except FileNotFoundError:
        return


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


def ensure_session_for_origin(
    *,
    origin_file: Path,
    state_file: Path,
) -> dict[str, Any]:
    resolved_origin = origin_file.expanduser().resolve()
    if not resolved_origin.exists():
        raise SessionError(f"origin file does not exist: {resolved_origin}")

    state = load_session_state_if_exists(state_file)
    if state is None:
        return start_session_server(origin_file=resolved_origin, state_file=state_file)

    server_url = state.get("server_url")
    if isinstance(server_url, str) and server_url.strip():
        try:
            health = session_request(
                state_file=state_file,
                method="GET",
                route="/health",
                timeout=1.0,
            )
        except SessionError:
            cleanup_session_artifacts(state_file)
            remove_session_state(state_file)
            return start_session_server(origin_file=resolved_origin, state_file=state_file)

        active_origin = Path(str(health.get("origin_file", ""))).resolve()
        if active_origin != resolved_origin:
            raise SessionError(
                f"an active session already exists for {active_origin}. Reuse it or wait for timeout before switching to {resolved_origin}."
            )
        return health

    cleanup_session_artifacts(state_file)
    remove_session_state(state_file)
    return start_session_server(origin_file=resolved_origin, state_file=state_file)
