from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any

from . import __version__
from .inspect import cmd_inspect
from .session import (
    SessionError,
    cleanup_stale_session_state,
    remove_session_state,
    resolve_state_file_path,
    session_request,
    start_session_server,
)
from .show import cmd_show
from .template_ops import (
    add_slide_to_template_draft,
    build_template_package,
    create_template_draft,
    save_template_package,
)

ROOT = Path(__file__).resolve().parents[2]
DEMO_FORM_PATH = ROOT / "examples" / "demo-form.json"

FUTURE_COMMANDS = {
    "preview": "Planned for task 006: export slide previews for validation.",
}

FUTURE_TEMPLATE_COMMANDS = {
    "show": "Planned for task 004: show template schema or form metadata.",
    "fill": "Planned for task 005: fill a template with JSON form data.",
    "modify": "Planned for task 006: replace or insert slides from template data.",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptxcli",
        description=(
            "Agent-friendly CLI for PPT template extraction, filling, and preview."
        ),
    )
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("version", help="Print CLI version")
    subparsers.add_parser("tech", help="Print current technical decisions")
    _add_inspect_parser(
        subparsers.add_parser(
            "inspect",
            help="Inspect slide candidates and emit JSON",
        )
    )
    _add_init_parser(subparsers.add_parser("init", help="Start a single-session server"))
    subparsers.add_parser("finish", help="Stop the active session server")
    subparsers.add_parser("preview", help="Placeholder for slide preview export")
    _add_show_parser(subparsers.add_parser("show", help="Render a slide preview"))

    template_parser = subparsers.add_parser("template", help="Template operations")
    template_subparsers = template_parser.add_subparsers(dest="template_command")
    _add_show_parser(
        template_subparsers.add_parser(
            "detect",
            help="Detect text/image candidates and optionally render annotated preview",
        )
    )
    _add_template_create_parser(
        template_subparsers.add_parser(
            "create",
            help="Create an empty template draft JSON from the current session or input PPTX",
        )
    )
    _add_template_add_slide_parser(
        template_subparsers.add_parser(
            "add_slide",
            help="Append one slide and its confirmed fields into the template draft JSON",
        )
    )
    _add_template_save_parser(
        template_subparsers.add_parser(
            "save",
            help="Crop the original PPTX to selected slides and generate manifest.json",
        )
    )
    _add_template_save_parser(
        template_subparsers.add_parser(
            "build",
            help="Alias of template save",
        )
    )
    for command_name in FUTURE_TEMPLATE_COMMANDS:
        template_subparsers.add_parser(
            command_name,
            help=f"Placeholder for template {command_name}",
        )

    demo_parser = subparsers.add_parser("demo", help="Demo helpers")
    demo_subparsers = demo_parser.add_subparsers(dest="demo_command")
    demo_subparsers.add_parser("form", help="Print a minimal fill-form JSON example")

    return parser


def _add_show_parser(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--input",
        help="Path to the input PPTX file; omit to reuse the active session",
    )
    parser.add_argument("--slide", type=int, default=0, help="Zero-based slide index")
    parser.add_argument(
        "--annotate",
        action="store_true",
        help="Overlay object-detection style candidate boxes onto the rendered preview",
    )
    parser.add_argument(
        "--output",
        help="Path to the output PNG file; defaults to <input>.slide-<n>.(preview|annotated).png",
    )
    parser.add_argument(
        "--candidates-out",
        help="Optional path to write candidate JSON",
    )


def _add_template_create_parser(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--name", required=True, help="Template name; .json/.pptx suffix is optional")


def _add_template_add_slide_parser(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--slide", type=int, required=True, help="Zero-based slide index")
    parser.add_argument(
        "--field",
        "-f",
        action="append",
        required=True,
        help='Field selector in the form "index:description"; repeat for multiple fields',
    )
    parser.add_argument(
        "--replace",
        action="store_true",
        help="Replace an existing slide entry for the same source slide index",
    )


def _add_template_save_parser(parser: argparse.ArgumentParser) -> None:
    del parser


def _add_inspect_parser(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--input",
        help="Path to the input PPTX file; omit to reuse the active session",
    )
    parser.add_argument("--slide", type=int, default=0, help="Zero-based slide index")
    parser.add_argument(
        "--output",
        help="Optional path to write the JSON payload",
    )


def _add_init_parser(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--origin_file",
        "--origin-file",
        dest="origin_file",
        required=True,
        help="Path to the source PPTX file for the active session",
    )


def cmd_version() -> int:
    print(__version__)
    return 0


def cmd_tech() -> int:
    message = {
        "language": "Python 3.11+",
        "dependency_management": "pyproject.toml + editable install",
        "cli_framework": "argparse",
        "ppt_read_write": "python-pptx (planned)",
        "preview": "LibreOffice headless (planned)",
        "data_model": "dataclass internal model + versioned JSON contract",
    }
    print(json.dumps(message, indent=2, ensure_ascii=False))
    return 0


def cmd_demo_form() -> int:
    with DEMO_FORM_PATH.open("r", encoding="utf-8") as fh:
        payload: Any = json.load(fh)
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


def cmd_future(command_name: str, detail: str) -> int:
    message = {
        "command": command_name,
        "status": "not_implemented",
        "detail": detail,
    }
    print(json.dumps(message, indent=2, ensure_ascii=False))
    return 0


def cmd_error(command_name: str, detail: str) -> int:
    message = {
        "command": command_name,
        "status": "error",
        "detail": detail,
    }
    print(json.dumps(message, indent=2, ensure_ascii=False), file=sys.stderr)
    return 1


def cmd_success(payload: dict[str, Any]) -> int:
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


def run_show_command(args: argparse.Namespace, command_name: str) -> int:
    if not args.input:
        state_file = resolve_state_file_path()
        try:
            payload = session_request(
                state_file=state_file,
                method="POST",
                route="/show" if command_name == "show" else "/template/detect",
                payload={
                    "command_name": command_name,
                    "slide_index": args.slide,
                    "annotate": args.annotate,
                    "output_path": args.output,
                    "candidates_out": args.candidates_out,
                },
            )
        except SessionError as exc:
            cleanup_stale_session_state(state_file)
            return cmd_error(command_name, str(exc))
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        return cmd_error(command_name, f"input file does not exist: {input_path}")

    try:
        return cmd_show(
            input_path=input_path,
            slide_index=args.slide,
            annotate=args.annotate,
            output_path=Path(args.output) if args.output else None,
            candidates_out=Path(args.candidates_out) if args.candidates_out else None,
            command_name=command_name,
        )
    except Exception as exc:
        return cmd_error(command_name, str(exc))


def run_inspect_command(args: argparse.Namespace, command_name: str) -> int:
    if not args.input:
        state_file = resolve_state_file_path()
        try:
            payload = session_request(
                state_file=state_file,
                method="POST",
                route="/inspect",
                payload={
                    "command_name": command_name,
                    "slide_index": args.slide,
                    "output_path": args.output,
                },
            )
        except SessionError as exc:
            cleanup_stale_session_state(state_file)
            return cmd_error(command_name, str(exc))
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        return cmd_error(command_name, f"input file does not exist: {input_path}")

    try:
        return cmd_inspect(
            input_path=input_path,
            slide_index=args.slide,
            output_path=Path(args.output) if args.output else None,
            command_name=command_name,
        )
    except Exception as exc:
        return cmd_error(command_name, str(exc))


def cmd_init(args: argparse.Namespace) -> int:
    origin_file = Path(args.origin_file).resolve()
    if not origin_file.exists():
        return cmd_error("init", f"origin file does not exist: {origin_file}")

    state_file = resolve_state_file_path()
    try:
        health = start_session_server(origin_file=origin_file, state_file=state_file)
    except SessionError as exc:
        return cmd_error("init", str(exc))

    payload = {
        "command": "init",
        "status": "ready",
        "origin_file": str(origin_file),
        "state_file": str(state_file),
        "server_url": health.get("server_url"),
        "slide_count": health.get("slide_count"),
    }
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


def cmd_finish() -> int:
    state_file = resolve_state_file_path()
    if not state_file.exists():
        return cmd_error("finish", "no active session found.")

    try:
        state = session_request(state_file=state_file, method="GET", route="/health", timeout=1.0)
        session_request(state_file=state_file, method="POST", route="/shutdown", payload={})
        deadline = time.time() + 5.0
        while state_file.exists() and time.time() < deadline:
            time.sleep(0.1)
    except SessionError:
        remove_session_state(state_file)
        payload = {
            "command": "finish",
            "status": "cleaned_stale_session",
            "state_file": str(state_file),
        }
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    remove_session_state(state_file)
    payload = {
        "command": "finish",
        "status": "stopped",
        "origin_file": state.get("origin_file"),
        "state_file": str(state_file),
    }
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "version":
        return cmd_version()
    if args.command == "tech":
        return cmd_tech()
    if args.command == "init":
        return cmd_init(args)
    if args.command == "finish":
        return cmd_finish()
    if args.command == "inspect":
        return run_inspect_command(args, "inspect")
    if args.command == "show":
        return run_show_command(args, "show")
    if args.command in FUTURE_COMMANDS:
        return cmd_future(args.command, FUTURE_COMMANDS[args.command])
    if args.command == "template":
        if args.template_command == "detect":
            return run_show_command(args, "template detect")
        if args.template_command == "create":
            try:
                payload = create_template_draft(
                    name=args.name,
                )
            except Exception as exc:
                return cmd_error("template create", str(exc))
            return cmd_success(payload)
        if args.template_command == "add_slide":
            try:
                payload = add_slide_to_template_draft(
                    slide_index=args.slide,
                    field_specs=args.field,
                    replace=args.replace,
                )
            except Exception as exc:
                return cmd_error("template add_slide", str(exc))
            return cmd_success(payload)
        if args.template_command == "save":
            try:
                payload = save_template_package()
            except Exception as exc:
                return cmd_error("template save", str(exc))
            return cmd_success(payload)
        if args.template_command == "build":
            try:
                payload = build_template_package()
            except Exception as exc:
                return cmd_error("template build", str(exc))
            return cmd_success(payload)
        if args.template_command in FUTURE_TEMPLATE_COMMANDS:
            return cmd_future(
                f"template {args.template_command}",
                FUTURE_TEMPLATE_COMMANDS[args.template_command],
            )
        parser.error("template requires a subcommand")
    if args.command == "demo":
        if args.demo_command == "form":
            return cmd_demo_form()
        parser.error("demo requires a subcommand")

    parser.print_help()
    return 0
