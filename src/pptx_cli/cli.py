from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from . import __version__

ROOT = Path(__file__).resolve().parents[2]
DEMO_FORM_PATH = ROOT / "examples" / "demo-form.json"

FUTURE_COMMANDS = {
    "inspect": "Planned for task 002: parse PPT structure into an intermediate model.",
    "preview": "Planned for task 006: export slide previews for validation.",
}

FUTURE_TEMPLATE_COMMANDS = {
    "detect": "Planned for task 003: detect candidate fields and create annotations.",
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

    template_parser = subparsers.add_parser("template", help="Template operations")
    template_subparsers = template_parser.add_subparsers(dest="template_command")
    for command_name in FUTURE_TEMPLATE_COMMANDS:
        template_subparsers.add_parser(
            command_name,
            help=f"Placeholder for template {command_name}",
        )

    demo_parser = subparsers.add_parser("demo", help="Demo helpers")
    demo_subparsers = demo_parser.add_subparsers(dest="demo_command")
    demo_subparsers.add_parser("form", help="Print a minimal fill-form JSON example")

    return parser


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


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "version":
        return cmd_version()
    if args.command == "tech":
        return cmd_tech()
    if args.command in FUTURE_COMMANDS:
        return cmd_future(args.command, FUTURE_COMMANDS[args.command])
    if args.command == "template":
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
