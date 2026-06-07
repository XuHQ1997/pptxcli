from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont

from .inspect import inspect_slide
from .models import Candidate

ANNOTATION_COLORS = [
    "#FF3B30",
    "#007AFF",
    "#34C759",
    "#FF9500",
    "#AF52DE",
    "#FF2D55",
]
DEFAULT_RENDER_SCALE = 2.0
MIN_LABEL_FONT_SIZE = 12
MAX_LABEL_FONT_SIZE = 72
FONT_CANDIDATES = [
    "DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
]


def cmd_show(
    *,
    input_path: Path,
    slide_index: int,
    annotate: bool,
    output_path: Path | None,
    candidates_out: Path | None,
) -> int:
    slide_data = inspect_slide(input_path=input_path, slide_index=slide_index)
    image = render_slide_preview(input_path=input_path, slide_index=slide_index)

    if annotate:
        image = annotate_candidates(
            image=image,
            candidates=slide_data.candidates,
            slide_width=slide_data.slide_width,
            slide_height=slide_data.slide_height,
        )

    resolved_output = _resolve_output_path(
        input_path=input_path,
        slide_index=slide_index,
        annotate=annotate,
        output_path=output_path,
    )
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    image.save(resolved_output)

    if candidates_out is not None:
        candidates_out.parent.mkdir(parents=True, exist_ok=True)
        candidates_out.write_text(
            json.dumps(slide_data.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    payload = slide_data.to_dict()
    payload.update(
        {
            "command": "show",
            "input": str(input_path),
            "annotated": annotate,
            "image_path": str(resolved_output),
        }
    )
    if candidates_out is not None:
        payload["candidates_path"] = str(candidates_out)

    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


def render_slide_preview(input_path: Path, slide_index: int) -> Image.Image:
    office_binary = shutil.which("libreoffice") or shutil.which("soffice")
    if office_binary is None:
        raise RuntimeError(
            "LibreOffice/soffice not found. Install LibreOffice to enable high-fidelity slide preview."
        )

    with tempfile.TemporaryDirectory(prefix="pptxcli-show-") as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)
        subprocess.run(
            [
                office_binary,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_dir),
                str(input_path),
            ],
            check=True,
            capture_output=True,
            text=True,
        )

        pdf_path = tmp_dir / f"{input_path.stem}.pdf"
        if not pdf_path.exists():
            raise RuntimeError(f"LibreOffice did not produce expected PDF: {pdf_path}")

        pdf = pdfium.PdfDocument(str(pdf_path))
        if slide_index < 0 or slide_index >= len(pdf):
            raise ValueError(
                f"slide index {slide_index} out of range for rendered PDF; total pages: {len(pdf)}"
            )

        page = pdf[slide_index]
        bitmap = page.render(scale=DEFAULT_RENDER_SCALE)
        image = bitmap.to_pil()
        page.close()
        pdf.close()
        return image


def annotate_candidates(
    *,
    image: Image.Image,
    candidates: list[Candidate],
    slide_width: int,
    slide_height: int,
) -> Image.Image:
    annotated = image.convert("RGB").copy()
    draw = ImageDraw.Draw(annotated)

    x_scale = annotated.width / slide_width
    y_scale = annotated.height / slide_height

    for candidate in candidates:
        color = ANNOTATION_COLORS[(candidate.index - 1) % len(ANNOTATION_COLORS)]
        left = candidate.bbox.x * x_scale
        top = candidate.bbox.y * y_scale
        right = (candidate.bbox.x + candidate.bbox.w) * x_scale
        bottom = (candidate.bbox.y + candidate.bbox.h) * y_scale
        draw.rectangle((left, top, right, bottom), outline=color, width=4)
        _draw_label(
            draw=draw,
            text=str(candidate.index),
            box=(left, top, right, bottom),
            color=color,
        )

    return annotated


def _draw_label(
    *,
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[float, float, float, float],
    color: str,
) -> None:
    left, top, _, _ = box
    layout = _fit_label_layout(draw=draw, text=text, box=box)

    label_left = left + 2
    label_top = top + 2

    draw.rectangle(
        (
            label_left,
            label_top,
            label_left + layout["label_width"],
            label_top + layout["label_height"],
        ),
        fill=color,
    )
    draw.text(
        (
            label_left + layout["text_left"],
            label_top + layout["text_top"],
        ),
        text=text,
        fill="white",
        font=layout["font"],
    )


def _fit_label_layout(
    *,
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[float, float, float, float],
) -> dict[str, object]:
    left, top, right, bottom = box
    box_width = max(1, int(right - left))
    box_height = max(1, int(bottom - top))
    max_label_width = max(20, min(box_width - 4, int(box_width * 0.45)))
    max_label_height = max(16, min(box_height - 4, int(box_height * 0.3)))
    max_font_size = max(
        MIN_LABEL_FONT_SIZE,
        min(MAX_LABEL_FONT_SIZE, int(min(box_width, box_height) * 0.55)),
    )

    chosen_font: ImageFont.ImageFont | ImageFont.FreeTypeFont = ImageFont.load_default()
    chosen_text_box = draw.textbbox((0, 0), text=text, font=chosen_font)
    chosen_padding_x = 4
    chosen_padding_y = 2

    for font_size in range(max_font_size, MIN_LABEL_FONT_SIZE - 1, -1):
        font = _load_annotation_font(font_size)
        text_box = draw.textbbox((0, 0), text=text, font=font)
        text_width = text_box[2] - text_box[0]
        text_height = text_box[3] - text_box[1]
        padding_x = max(4, int(font_size * 0.35))
        padding_y = max(2, int(font_size * 0.2))
        label_width = text_width + padding_x * 2
        label_height = text_height + padding_y * 2

        if label_width <= max_label_width and label_height <= max_label_height:
            chosen_font = font
            chosen_text_box = text_box
            chosen_padding_x = padding_x
            chosen_padding_y = padding_y
            break
    else:
        chosen_font = _load_annotation_font(MIN_LABEL_FONT_SIZE)
        chosen_text_box = draw.textbbox((0, 0), text=text, font=chosen_font)

    text_width = chosen_text_box[2] - chosen_text_box[0]
    text_height = chosen_text_box[3] - chosen_text_box[1]
    label_width = min(max_label_width, text_width + chosen_padding_x * 2)
    label_height = min(max_label_height, text_height + chosen_padding_y * 2)
    text_left = max(
        1.0,
        (label_width - text_width) / 2 - chosen_text_box[0],
    )
    text_top = max(
        1.0,
        (label_height - text_height) / 2 - chosen_text_box[1],
    )

    return {
        "font": chosen_font,
        "label_width": label_width,
        "label_height": label_height,
        "text_left": text_left,
        "text_top": text_top,
    }


def _load_annotation_font(
    size: int,
) -> ImageFont.ImageFont | ImageFont.FreeTypeFont:
    for font_name in FONT_CANDIDATES:
        try:
            return ImageFont.truetype(font_name, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _resolve_output_path(
    *,
    input_path: Path,
    slide_index: int,
    annotate: bool,
    output_path: Path | None,
) -> Path:
    if output_path is not None:
        return output_path.resolve()
    suffix = "annotated" if annotate else "preview"
    return input_path.resolve().with_name(
        f"{input_path.stem}.slide-{slide_index}.{suffix}.png"
    )
