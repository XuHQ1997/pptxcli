from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(slots=True)
class BBox:
    x: int
    y: int
    w: int
    h: int

    def to_dict(self) -> dict[str, int]:
        return asdict(self)


@dataclass(slots=True)
class Candidate:
    index: int
    candidate_id: str
    shape_id: int
    kind: str
    role_hint: str
    bbox: BBox
    text_excerpt: str | None = None
    image_name: str | None = None

    def to_dict(self) -> dict[str, object]:
        payload = asdict(self)
        payload["bbox"] = self.bbox.to_dict()
        return payload


@dataclass(slots=True)
class SlideCandidates:
    slide_index: int
    slide_width: int
    slide_height: int
    candidates: list[Candidate]

    def to_dict(self) -> dict[str, object]:
        return {
            "slide_index": self.slide_index,
            "slide_size": {
                "width": self.slide_width,
                "height": self.slide_height,
                "unit": "emu",
            },
            "candidates": [candidate.to_dict() for candidate in self.candidates],
        }
