from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .show import cmd_show

ROOT = Path(__file__).resolve().parents[2]
DEMO_FORM_PATH = ROOT / "examples" / "demo-form.json"

FUTURE_COMMANDS = {
    "inspect": "Planned for task 002: parse PPT structure into an intermediate model.",
    "preview": "Planned for task 006: export slide previews for validation.",
}

FUTURE_TEMPLATE_COMMANDS = {
    "build": "Planned for task 004: build a template package and manifest.",
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
    subparsers.add_parser("inspect", help="Placeholder for PPT structure inspection")
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
    parser.add_argument("--input", required=True, help="Path to the input PPTX file")
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


def run_show_command(args: argparse.Namespace, command_name: str) -> int:
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
        )
    except Exception as exc:
        return cmd_error(command_name, str(exc))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "version":
        return cmd_version()
    if args.command == "tech":
        return cmd_tech()
    if args.command == "show":
        return run_show_command(args, "show")
    if args.command in FUTURE_COMMANDS:
        return cmd_future(args.command, FUTURE_COMMANDS[args.command])
    if args.command == "template":
        if args.template_command == "detect":
            return run_show_command(args, "template detect")
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
