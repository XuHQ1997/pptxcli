from __future__ import annotations

import json
import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pptx import Presentation

from .session import (
    SessionError,
    cleanup_stale_session_state,
    load_session_state,
    resolve_state_file_path,
    session_request,
)

TEMPLATE_ROOT_ENV = "PPTXCLI_TEMPLATE_ROOT"
TEMPLATE_DRAFT_VERSION = 1
MANIFEST_VERSION = 1
TEMPLATE_STATE_FILE_NAME = ".pptxcli-template-state.json"


def resolve_template_root(cli_argv0: str | None = None) -> Path:
    env_path = os.environ.get(TEMPLATE_ROOT_ENV)
    if env_path:
        return Path(env_path).expanduser().resolve()
    return resolve_state_file_path(cli_argv0).parent / "templates"


def resolve_template_state_file() -> Path:
    return resolve_template_root() / TEMPLATE_STATE_FILE_NAME


def create_template_draft(
    *,
    name: str,
) -> dict[str, Any]:
    template_name = normalize_template_name(name)
    origin_file = _resolve_origin_file()
    draft_path = resolve_template_root() / f"{template_name}.json"
    if draft_path.exists():
        raise ValueError(f"template draft already exists: {draft_path}")

    payload = {
        "draft_version": TEMPLATE_DRAFT_VERSION,
        "template_name": template_name,
        "origin_file": str(origin_file),
        "created_at": _now_iso(),
        "slides": [],
    }
    _write_json(draft_path, payload)
    _save_template_state({"active_template": template_name})
    return {
        "command": "template create",
        "status": "created",
        "template_name": template_name,
        "origin_file": str(origin_file),
        "draft_path": str(draft_path),
    }


def add_slide_to_template_draft(
    *,
    slide_index: int,
    field_specs: list[str],
    replace: bool,
) -> dict[str, Any]:
    template_name = _resolve_template_name()
    draft_path = resolve_template_root() / f"{template_name}.json"
    draft = _load_template_draft(draft_path)
    inspect_payload = _load_slide_candidates(slide_index=slide_index)
    slide_record = _build_slide_record(
        slide_index=slide_index,
        field_specs=field_specs,
        inspect_payload=inspect_payload,
    )

    existing_index = _find_slide_record_index(
        draft=draft,
        source_slide_index=slide_index,
    )
    if existing_index is not None:
        if not replace:
            raise ValueError(
                "slide already exists in template draft; pass --replace to overwrite it"
            )
        draft["slides"][existing_index] = slide_record
    else:
        draft["slides"].append(slide_record)

    draft["updated_at"] = _now_iso()
    _write_json(draft_path, draft)
    return {
        "command": "template add_slide",
        "status": "updated",
        "template_name": template_name,
        "draft_path": str(draft_path),
        "slide_name": slide_record["slide_name"],
        "source_slide_index": slide_index,
        "field_count": len(slide_record["fields"]),
    }


def save_template_package(
) -> dict[str, Any]:
    template_name = _resolve_template_name()
    draft_path = resolve_template_root() / f"{template_name}.json"
    draft = _load_template_draft(draft_path)
    slides = draft.get("slides", [])
    if not slides:
        raise ValueError("template draft has no slides; add at least one slide before save")

    origin_file = Path(str(draft["origin_file"])).resolve()
    if not origin_file.exists():
        raise ValueError(f"origin file does not exist: {origin_file}")

    package_dir = resolve_template_root() / template_name
    package_dir.mkdir(parents=True, exist_ok=True)
    template_pptx_path = package_dir / "template.pptx"
    manifest_path = package_dir / "manifest.json"

    source_slide_indexes = [int(slide["source_slide_index"]) for slide in slides]
    _save_cropped_presentation(
        origin_file=origin_file,
        slide_indexes=source_slide_indexes,
        output_path=template_pptx_path,
    )

    manifest = _build_manifest(
        draft=draft,
        draft_path=draft_path,
        template_pptx_path=template_pptx_path,
    )
    _write_json(manifest_path, manifest)

    return {
        "command": "template save",
        "status": "saved",
        "template_name": template_name,
        "draft_path": str(draft_path),
        "package_dir": str(package_dir),
        "template_pptx_path": str(template_pptx_path),
        "manifest_path": str(manifest_path),
        "slide_count": manifest["slide_count"],
        "field_count": manifest["field_count"],
    }


def build_template_package(
) -> dict[str, Any]:
    payload = save_template_package()
    payload["command"] = "template build"
    return payload


def normalize_template_name(name: str) -> str:
    normalized = Path(name.strip()).name
    if normalized.endswith(".json") or normalized.endswith(".pptx"):
        normalized = Path(normalized).stem
    if not normalized:
        raise ValueError("template name cannot be empty")
    return normalized


def _resolve_template_name() -> str:
    state = _load_template_state()
    active_template = str(state.get("active_template", "")).strip()
    if not active_template:
        raise ValueError(
            "no active template found. Run `pptxcli template create --name <name>` first."
        )
    return normalize_template_name(active_template)


def _resolve_origin_file() -> Path:
    state_file = resolve_state_file_path()
    try:
        session_request(state_file=state_file, method="GET", route="/health", timeout=1.0)
        state = load_session_state(state_file)
    except SessionError as exc:
        cleanup_stale_session_state(state_file)
        raise ValueError(str(exc)) from exc

    origin_file = Path(str(state.get("origin_file", ""))).resolve()
    if not origin_file.exists():
        raise ValueError(f"origin file does not exist: {origin_file}")
    return origin_file


def _load_slide_candidates(*, slide_index: int) -> dict[str, Any]:
    state_file = resolve_state_file_path()
    try:
        return session_request(
            state_file=state_file,
            method="POST",
            route="/template/candidates",
            payload={
                "command_name": "template add_slide",
                "slide_index": slide_index,
            },
        )
    except SessionError as exc:
        cleanup_stale_session_state(state_file)
        raise ValueError(str(exc)) from exc


def _build_slide_record(
    *,
    slide_index: int,
    field_specs: list[str],
    inspect_payload: dict[str, Any],
) -> dict[str, Any]:
    candidates = inspect_payload.get("candidates")
    if not isinstance(candidates, list):
        raise ValueError("invalid inspect payload: missing candidates")
    candidates_by_index = {
        int(candidate.get("index")): candidate
        for candidate in candidates
        if isinstance(candidate, dict) and candidate.get("index") is not None
    }

    if not field_specs:
        raise ValueError("at least one --field is required")

    seen_indexes: set[int] = set()
    fields: list[dict[str, Any]] = []
    for raw_field_spec in field_specs:
        selected_index, description = _parse_field_spec(raw_field_spec)
        if selected_index in seen_indexes:
            raise ValueError(f"duplicate field index on slide {slide_index}: {selected_index}")
        seen_indexes.add(selected_index)

        candidate = candidates_by_index.get(selected_index)
        if candidate is None:
            raise ValueError(
                f"field index {selected_index} does not exist on slide {slide_index}"
            )

        field_type = str(candidate.get("kind", "")).strip()
        field_record: dict[str, Any] = {
            "index": selected_index,
            "description": description,
            "type": field_type,
            "shape_id": int(candidate["shape_id"]),
            "role_hint": candidate.get("role_hint"),
            "bbox": candidate.get("bbox"),
        }
        if candidate.get("text_excerpt") is not None:
            field_record["text_excerpt"] = candidate["text_excerpt"]
        if candidate.get("image_name") is not None:
            field_record["image_name"] = candidate["image_name"]
        fields.append(field_record)

    return {
        "slide_name": f"slide_{slide_index}",
        "source_slide_index": slide_index,
        "slide_size": inspect_payload.get("slide_size"),
        "fields": fields,
        "added_at": _now_iso(),
    }


def _find_slide_record_index(
    *,
    draft: dict[str, Any],
    source_slide_index: int,
) -> int | None:
    for index, slide in enumerate(draft.get("slides", [])):
        existing_source_slide = int(slide.get("source_slide_index", -1))
        if existing_source_slide == source_slide_index:
            return index
    return None


def _build_manifest(
    *,
    draft: dict[str, Any],
    draft_path: Path,
    template_pptx_path: Path,
) -> dict[str, Any]:
    slides_payload: list[dict[str, Any]] = []
    all_fields: list[dict[str, Any]] = []

    for template_slide_index, slide in enumerate(draft["slides"]):
        fields_payload: list[dict[str, Any]] = []
        for field in slide["fields"]:
            field_payload = {
                "index": field["index"],
                "description": field["description"],
                "type": field["type"],
                "shape_id": field["shape_id"],
                "role_hint": field.get("role_hint"),
                "bbox": field.get("bbox"),
                "binding": {
                    "kind": "shape",
                    "shape_id": field["shape_id"],
                },
            }
            if field.get("text_excerpt") is not None:
                field_payload["text_excerpt"] = field["text_excerpt"]
            if field.get("image_name") is not None:
                field_payload["image_name"] = field["image_name"]
            fields_payload.append(field_payload)
            all_fields.append(
                {
                    "slide_name": slide["slide_name"],
                    **field_payload,
                }
            )

        slides_payload.append(
            {
                "slide_name": slide["slide_name"],
                "template_slide_index": template_slide_index,
                "source_slide_index": slide["source_slide_index"],
                "slide_size": slide.get("slide_size"),
                "field_count": len(fields_payload),
                "fields": fields_payload,
            }
        )

    return {
        "manifest_version": MANIFEST_VERSION,
        "template_name": draft["template_name"],
        "origin_file": draft["origin_file"],
        "draft_path": str(draft_path),
        "template_file": template_pptx_path.name,
        "created_at": _now_iso(),
        "slide_count": len(slides_payload),
        "field_count": len(all_fields),
        "slides": slides_payload,
        "fields": all_fields,
    }


def _save_cropped_presentation(
    *,
    origin_file: Path,
    slide_indexes: list[int],
    output_path: Path,
) -> None:
    if len(set(slide_indexes)) != len(slide_indexes):
        raise ValueError("template draft contains duplicate source_slide_index values")

    presentation = Presentation(str(origin_file))
    total_slides = len(presentation.slides)
    for slide_index in slide_indexes:
        if slide_index < 0 or slide_index >= total_slides:
            raise ValueError(
                f"slide index {slide_index} out of range; total slides: {total_slides}"
            )

    slide_id_list = presentation.slides._sldIdLst
    original_slide_ids = list(slide_id_list)
    selected_by_index = {slide_index: original_slide_ids[slide_index] for slide_index in slide_indexes}

    for slide_index in range(len(original_slide_ids) - 1, -1, -1):
        slide_id = original_slide_ids[slide_index]
        if slide_index in selected_by_index:
            continue
        presentation.part.drop_rel(slide_id.rId)
        slide_id_list.remove(slide_id)

    for slide_id in list(slide_id_list):
        slide_id_list.remove(slide_id)
    for slide_index in slide_indexes:
        slide_id_list.append(selected_by_index[slide_index])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(output_path))


def _load_template_draft(draft_path: Path) -> dict[str, Any]:
    payload = _load_json_file(draft_path)
    if int(payload.get("draft_version", 0)) != TEMPLATE_DRAFT_VERSION:
        raise ValueError(
            f"unsupported template draft version in {draft_path}: {payload.get('draft_version')}"
        )
    return payload


def _parse_field_spec(value: str) -> tuple[int, str]:
    raw_value = value.strip()
    if ":" not in raw_value:
        raise ValueError(f'invalid --field value `{value}`; expected "index:description"')
    raw_index, raw_description = raw_value.split(":", 1)
    try:
        index = int(raw_index.strip())
    except ValueError as exc:
        raise ValueError(f'invalid --field index in `{value}`') from exc
    if index <= 0:
        raise ValueError(f"field index must be >= 1: {index}")
    description = raw_description.strip()
    if not description:
        raise ValueError(f"field description cannot be empty: `{value}`")
    return index, description


def _load_json_file(path: Path) -> dict[str, Any]:
    resolved = path.resolve()
    if not resolved.exists():
        raise ValueError(f"file does not exist: {resolved}")
    payload = json.loads(resolved.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"expected JSON object in file: {resolved}")
    return payload


def _load_template_state() -> dict[str, Any]:
    state_file = resolve_template_state_file()
    if not state_file.exists():
        return {}
    payload = json.loads(state_file.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"expected JSON object in file: {state_file}")
    return payload


def _save_template_state(payload: dict[str, Any]) -> None:
    _write_json(resolve_template_state_file(), payload)


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def _now_iso() -> str:
    return datetime.now(UTC).isoformat()
