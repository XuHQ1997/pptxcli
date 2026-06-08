from __future__ import annotations

import argparse
import json
import os
from datetime import UTC, datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from threading import Thread
from typing import Any

from .inspect import build_candidates_payload, build_inspect_payload, load_presentation
from .session import remove_session_state, save_session_state
from .show import build_show_payload


class SessionRuntime:
    def __init__(self, origin_file: Path) -> None:
        self.origin_file = origin_file.resolve()
        self.presentation = load_presentation(self.origin_file)
        self.inspect_cache: dict[int, dict[str, Any]] = {}
        self.candidate_cache: dict[int, dict[str, Any]] = {}
        self.context: dict[str, Any] = {}

    def inspect(
        self,
        *,
        slide_index: int,
        output_path: Path | None,
        command_name: str,
    ) -> dict[str, Any]:
        cache_key = slide_index
        cached = self.inspect_cache.get(cache_key)
        if cached is None:
            cached = build_inspect_payload(
                command_name=command_name,
                input_path=self.origin_file,
                slide_index=slide_index,
                presentation=self.presentation,
            )
            self.inspect_cache[cache_key] = cached

        payload = json.loads(json.dumps(cached))
        if output_path is not None:
            resolved_output = output_path.resolve()
            resolved_output.parent.mkdir(parents=True, exist_ok=True)
            resolved_output.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            payload["output_path"] = str(resolved_output)
        return payload

    def show(
        self,
        *,
        slide_index: int,
        annotate: bool,
        output_path: Path | None,
        candidates_out: Path | None,
        command_name: str,
    ) -> dict[str, Any]:
        return build_show_payload(
            command_name=command_name,
            input_path=self.origin_file,
            slide_index=slide_index,
            annotate=annotate,
            output_path=output_path,
            candidates_out=candidates_out,
            presentation=self.presentation,
        )

    def template_candidates(
        self,
        *,
        slide_index: int,
        command_name: str,
    ) -> dict[str, Any]:
        cache_key = slide_index
        cached = self.candidate_cache.get(cache_key)
        if cached is None:
            cached = build_candidates_payload(
                command_name=command_name,
                input_path=self.origin_file,
                slide_index=slide_index,
                presentation=self.presentation,
            )
            self.candidate_cache[cache_key] = cached
        return json.loads(json.dumps(cached))


class SessionHTTPServer(HTTPServer):
    def __init__(
        self,
        server_address: tuple[str, int],
        runtime: SessionRuntime,
        state_file: Path,
    ) -> None:
        super().__init__(server_address, SessionRequestHandler)
        self.runtime = runtime
        self.state_file = state_file


class SessionRequestHandler(BaseHTTPRequestHandler):
    server: SessionHTTPServer

    def do_GET(self) -> None:  # noqa: N802
        try:
            if self.path == "/health":
                self._send_json(
                    {
                        "status": "ok",
                        "server_url": f"http://127.0.0.1:{self.server.server_address[1]}",
                        "origin_file": str(self.server.runtime.origin_file),
                        "slide_count": len(self.server.runtime.presentation.slides),
                        "cached_inspections": len(self.server.runtime.inspect_cache),
                    }
                )
                return
            self._send_json(
                {"status": "error", "detail": f"unknown route: {self.path}"},
                status=404,
            )
        except Exception as exc:
            self._send_json({"status": "error", "detail": str(exc)}, status=500)

    def do_POST(self) -> None:  # noqa: N802
        try:
            payload = self._read_json()
            if self.path == "/inspect":
                response = self.server.runtime.inspect(
                    slide_index=int(payload.get("slide_index", 0)),
                    output_path=_optional_path(payload.get("output_path")),
                    command_name=str(payload.get("command_name", "inspect")),
                )
                self._send_json(response)
                return

            if self.path in {"/show", "/template/detect"}:
                response = self.server.runtime.show(
                    slide_index=int(payload.get("slide_index", 0)),
                    annotate=bool(payload.get("annotate", False)),
                    output_path=_optional_path(payload.get("output_path")),
                    candidates_out=_optional_path(payload.get("candidates_out")),
                    command_name=str(payload.get("command_name", "show")),
                )
                self._send_json(response)
                return

            if self.path == "/template/candidates":
                response = self.server.runtime.template_candidates(
                    slide_index=int(payload.get("slide_index", 0)),
                    command_name=str(payload.get("command_name", "template candidates")),
                )
                self._send_json(response)
                return

            if self.path == "/shutdown":
                self._send_json({"status": "shutting_down"})
                Thread(target=self.server.shutdown, daemon=True).start()
                return

            self._send_json(
                {"status": "error", "detail": f"unknown route: {self.path}"},
                status=404,
            )
        except Exception as exc:
            self._send_json({"status": "error", "detail": str(exc)}, status=500)

    def log_message(self, format: str, *args: object) -> None:
        return

    def _read_json(self) -> dict[str, Any]:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return {}
        raw = self.rfile.read(content_length).decode("utf-8")
        return json.loads(raw)

    def _send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def _optional_path(value: object) -> Path | None:
    if value in {None, ""}:
        return None
    return Path(str(value))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="pptx_cli session server")
    parser.add_argument("--origin-file", required=True)
    parser.add_argument("--state-file", required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    origin_file = Path(args.origin_file).resolve()
    state_file = Path(args.state_file).resolve()

    runtime = SessionRuntime(origin_file)
    server = SessionHTTPServer(("127.0.0.1", 0), runtime, state_file)

    state_payload = {
        "version": 1,
        "pid": os.getpid(),
        "server_url": f"http://127.0.0.1:{server.server_address[1]}",
        "origin_file": str(origin_file),
        "created_at": datetime.now(UTC).isoformat(),
    }
    save_session_state(state_file, state_payload)

    try:
        server.serve_forever(poll_interval=0.2)
    finally:
        server.server_close()
        remove_session_state(state_file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
