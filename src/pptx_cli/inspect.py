from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .models import BBox, Candidate, SlideCandidates


def inspect_slide(input_path: Path, slide_index: int) -> SlideCandidates:
    presentation = Presentation(str(input_path))
    if slide_index < 0 or slide_index >= len(presentation.slides):
        raise ValueError(
            f"slide index {slide_index} out of range; total slides: {len(presentation.slides)}"
        )

    slide = presentation.slides[slide_index]
    candidates: list[Candidate] = []

    for shape in slide.shapes:
        candidate = _shape_to_candidate(shape=shape, next_index=len(candidates) + 1)
        if candidate is not None:
            candidates.append(candidate)

    return SlideCandidates(
        slide_index=slide_index,
        slide_width=presentation.slide_width,
        slide_height=presentation.slide_height,
        candidates=candidates,
    )


def _shape_to_candidate(shape: object, next_index: int) -> Candidate | None:
    bbox = BBox(x=shape.left, y=shape.top, w=shape.width, h=shape.height)
    shape_id = int(shape.shape_id)
    candidate_id = f"slide-shape-{shape_id}"

    if _is_picture(shape):
        image = getattr(shape, "image", None)
        image_name = getattr(image, "filename", None)
        return Candidate(
            index=next_index,
            candidate_id=candidate_id,
            shape_id=shape_id,
            kind="image",
            role_hint="image",
            bbox=bbox,
            image_name=image_name,
        )

    if getattr(shape, "has_text_frame", False):
        text = _normalize_text(shape.text)
        if text:
            return Candidate(
                index=next_index,
                candidate_id=candidate_id,
                shape_id=shape_id,
                kind="text",
                role_hint=_infer_text_role_hint(text),
                bbox=bbox,
                text_excerpt=text[:80],
            )

    return None


def _is_picture(shape: object) -> bool:
    shape_type = getattr(shape, "shape_type", None)
    return shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image")


def _normalize_text(text: str) -> str:
    return " ".join(text.split())


def _infer_text_role_hint(text: str) -> str:
    lower_text = text.lower()
    if len(text) <= 40:
        return "title"
    if any(token in lower_text for token in ("summary", "overview", "agenda")):
        return "section_heading"
    if len(text) >= 120:
        return "body"
    return "text"
