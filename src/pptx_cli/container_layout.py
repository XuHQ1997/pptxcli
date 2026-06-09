from __future__ import annotations

import io
import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

from .models import BBox

LAYOUT_TYPES = {"vertical", "horizontal", "grid"}
LEAF_TYPES = {"text", "image", "svg"}
ALIGN_VALUES = {"start", "center", "end", "stretch"}


@dataclass(slots=True)
class Insets:
    top: int = 0
    right: int = 0
    bottom: int = 0
    left: int = 0

    def to_dict(self) -> dict[str, int]:
        return {
            "top": self.top,
            "right": self.right,
            "bottom": self.bottom,
            "left": self.left,
        }


@dataclass(slots=True)
class ContentNode:
    node_type: str
    width: int | None = None
    height: int | None = None
    min_width: int = 0
    min_height: int = 0
    max_width: int | None = None
    max_height: int | None = None
    grow: float = 0.0
    align: str = "start"
    cross_align: str = "stretch"
    fit: str = "fill"
    name: str | None = None
    bbox: BBox | None = None
    layout: str | None = None
    padding: Insets = field(default_factory=Insets)
    gap: int = 0
    columns: int | None = None
    rows: int | None = None
    column_weights: list[float] | None = None
    row_weights: list[float] | None = None
    row: int | None = None
    column: int | None = None
    row_span: int = 1
    col_span: int = 1
    text: str | None = None
    path: str | None = None
    style: dict[str, Any] = field(default_factory=dict)
    children: list["ContentNode"] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "type": self.node_type,
            "width": self.width,
            "height": self.height,
            "min_width": self.min_width,
            "min_height": self.min_height,
            "max_width": self.max_width,
            "max_height": self.max_height,
            "grow": self.grow,
            "align": self.align,
            "cross_align": self.cross_align,
            "fit": self.fit,
            "name": self.name,
            "style": self.style,
        }
        if self.bbox is not None:
            payload["bbox"] = self.bbox.to_dict()
        if self.layout is not None:
            payload["layout"] = self.layout
            payload["padding"] = self.padding.to_dict()
            payload["gap"] = self.gap
            payload["columns"] = self.columns
            payload["rows"] = self.rows
            payload["column_weights"] = self.column_weights
            payload["row_weights"] = self.row_weights
            payload["children"] = [child.to_dict() for child in self.children]
        if self.row is not None:
            payload["row"] = self.row
        if self.column is not None:
            payload["column"] = self.column
        if self.row_span != 1:
            payload["row_span"] = self.row_span
        if self.col_span != 1:
            payload["col_span"] = self.col_span
        if self.text is not None:
            payload["text"] = self.text
        if self.path is not None:
            payload["path"] = self.path
        return payload


@dataclass(slots=True)
class ResolvedContent:
    node_type: str
    bbox: BBox
    name: str | None = None
    text: str | None = None
    path: str | None = None
    style: dict[str, Any] = field(default_factory=dict)
    fit: str = "fill"
    children: list["ResolvedContent"] = field(default_factory=list)
    debug: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "type": self.node_type,
            "bbox": self.bbox.to_dict(),
            "name": self.name,
            "fit": self.fit,
            "style": self.style,
            "debug": self.debug,
            "children": [child.to_dict() for child in self.children],
        }
        if self.text is not None:
            payload["text"] = self.text
        if self.path is not None:
            payload["path"] = self.path
        return payload


def parse_content_spec(
    raw_value: str,
    *,
    slide_width: int,
    slide_height: int,
) -> ContentNode:
    try:
        payload = json_loads(raw_value)
    except ValueError as exc:
        raise ValueError(f"invalid --content JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("--content expects a JSON object")
    node = _parse_node(payload, is_root=True)
    if node.bbox is None:
        node.bbox = BBox(0, 0, slide_width, slide_height)
    return node


def solve_content_layout(
    node: ContentNode,
    *,
    slide_width: int,
    slide_height: int,
) -> ResolvedContent:
    root_bbox = node.bbox or BBox(0, 0, slide_width, slide_height)
    bounded_root = _clip_bbox(root_bbox, BBox(0, 0, slide_width, slide_height))
    return _resolve_node(node=node, bbox=bounded_root)


def flatten_resolved_leaves(node: ResolvedContent) -> list[ResolvedContent]:
    if node.node_type in LEAF_TYPES:
        return [node]
    leaves: list[ResolvedContent] = []
    for child in node.children:
        leaves.extend(flatten_resolved_leaves(child))
    return leaves


def render_resolved_content(*, slide: Any, resolved: ResolvedContent) -> list[dict[str, Any]]:
    rendered: list[dict[str, Any]] = []
    for leaf in flatten_resolved_leaves(resolved):
        if leaf.node_type == "text":
            shape = _render_text(slide=slide, leaf=leaf)
        elif leaf.node_type in {"image", "svg"}:
            shape = _render_image(slide=slide, leaf=leaf)
        else:
            raise ValueError(f"unsupported resolved content type: {leaf.node_type}")
        rendered.append(
            {
                "type": leaf.node_type,
                "shape_id": int(shape.shape_id),
                "bbox": leaf.bbox.to_dict(),
                "name": leaf.name,
            }
        )
    return rendered


def json_loads(raw_value: str) -> Any:
    import json

    return json.loads(raw_value)


def _parse_node(payload: dict[str, Any], *, is_root: bool) -> ContentNode:
    raw_type = str(payload.get("type") or "").strip().lower()
    raw_layout = str(payload.get("layout") or "").strip().lower()
    if raw_type in LAYOUT_TYPES and not raw_layout:
        raw_layout = raw_type
        raw_type = "container"
    if not raw_type:
        raw_type = "container" if raw_layout else ""
    if raw_type not in {"container", *LEAF_TYPES}:
        raise ValueError(f"unsupported content node type: {raw_type or '<empty>'}")

    bbox = None
    raw_bbox = payload.get("bbox")
    if raw_bbox is not None:
        if not isinstance(raw_bbox, dict):
            raise ValueError("content bbox must be an object")
        bbox = BBox(
            x=_read_int(raw_bbox, "x", default=0),
            y=_read_int(raw_bbox, "y", default=0),
            w=_read_int(raw_bbox, "w"),
            h=_read_int(raw_bbox, "h"),
        )
    elif is_root and all(key in payload for key in ("x", "y", "w", "h")):
        bbox = BBox(
            x=_read_int(payload, "x", default=0),
            y=_read_int(payload, "y", default=0),
            w=_read_int(payload, "w"),
            h=_read_int(payload, "h"),
        )

    node = ContentNode(
        node_type=raw_type,
        width=_read_optional_int(payload, "width"),
        height=_read_optional_int(payload, "height"),
        min_width=_read_int(payload, "min_width", default=0),
        min_height=_read_int(payload, "min_height", default=0),
        max_width=_read_optional_int(payload, "max_width"),
        max_height=_read_optional_int(payload, "max_height"),
        grow=_read_float(payload, "grow", default=0.0),
        align=_read_align(payload.get("align"), default="start"),
        cross_align=_read_align(payload.get("cross_align"), default="stretch"),
        fit=str(payload.get("fit") or "fill").strip().lower(),
        name=_read_optional_str(payload, "name"),
        bbox=bbox,
        row=_read_optional_int(payload, "row"),
        column=_read_optional_int(payload, "column"),
        row_span=max(_read_int(payload, "row_span", default=1), 1),
        col_span=max(_read_int(payload, "col_span", default=1), 1),
        style=_read_style(payload.get("style")),
    )

    if raw_type == "container":
        if raw_layout not in LAYOUT_TYPES:
            raise ValueError("container node requires layout: vertical, horizontal, or grid")
        children = payload.get("children")
        if not isinstance(children, list):
            raise ValueError("container node requires children array")
        node.layout = raw_layout
        node.padding = _read_insets(payload.get("padding"))
        node.gap = _read_int(payload, "gap", default=0)
        node.columns = _read_optional_int(payload, "columns")
        node.rows = _read_optional_int(payload, "rows")
        node.column_weights = _read_optional_float_list(payload.get("column_weights"))
        node.row_weights = _read_optional_float_list(payload.get("row_weights"))
        node.children = [
            _parse_node(child, is_root=False) if isinstance(child, dict) else _invalid_child(child)
            for child in children
        ]
        return node

    if raw_type == "text":
        node.text = _read_leaf_value(payload, ("text", "value", "content"))
        return node

    node.path = _resolve_asset_path(_read_leaf_value(payload, ("path", "src", "value")))
    return node


def _invalid_child(value: Any) -> ContentNode:
    raise ValueError(f"container child must be an object, got: {type(value).__name__}")


def _resolve_node(*, node: ContentNode, bbox: BBox) -> ResolvedContent:
    bounded = _enforce_node_constraints(node=node, bbox=bbox)
    if node.node_type != "container":
        return ResolvedContent(
            node_type=node.node_type,
            bbox=bounded,
            name=node.name,
            text=node.text,
            path=node.path,
            style=node.style,
            fit=node.fit,
            debug={"width": bounded.w, "height": bounded.h},
        )

    inner = _apply_padding(bounded, node.padding)
    if node.layout in {"vertical", "horizontal"}:
        children, overflow = _resolve_linear_children(node=node, inner=inner)
    elif node.layout == "grid":
        children, overflow = _resolve_grid_children(node=node, inner=inner)
    else:
        raise ValueError(f"unsupported container layout: {node.layout}")
    return ResolvedContent(
        node_type="container",
        bbox=bounded,
        name=node.name,
        style=node.style,
        children=children,
        debug={
            "layout": node.layout,
            "gap": node.gap,
            "padding": node.padding.to_dict(),
            "overflow": overflow,
        },
    )


def _resolve_linear_children(
    *,
    node: ContentNode,
    inner: BBox,
) -> tuple[list[ResolvedContent], int]:
    is_vertical = node.layout == "vertical"
    child_count = len(node.children)
    if child_count == 0:
        return [], 0
    gap_total = node.gap * max(child_count - 1, 0)
    inner_main = inner.h if is_vertical else inner.w
    inner_cross = inner.w if is_vertical else inner.h

    main_sizes: list[int | None] = []
    minimum_sizes: list[int] = []
    weights: list[float] = []
    fixed_total = gap_total
    for child in node.children:
        fixed_main = child.height if is_vertical else child.width
        minimum_main = max(child.min_height if is_vertical else child.min_width, 0)
        fixed_resolved = None if fixed_main is None else max(int(fixed_main), minimum_main)
        main_sizes.append(fixed_resolved)
        minimum_sizes.append(minimum_main)
        if fixed_resolved is not None:
            fixed_total += fixed_resolved
            weights.append(0.0)
            continue
        if child.grow > 0:
            weights.append(child.grow)
            fixed_total += minimum_main
            continue
        weights.append(1.0)
        fixed_total += minimum_main

    free_space = inner_main - fixed_total
    overflow = max(-free_space, 0)
    free_space = max(free_space, 0)
    weight_total = sum(weight for size, weight in zip(main_sizes, weights, strict=False) if size is None)
    remaining = free_space
    allocated_sizes: list[int] = []
    pending_indexes = [index for index, size in enumerate(main_sizes) if size is None]
    for index, size in enumerate(main_sizes):
        if size is not None:
            allocated_sizes.append(size)
            continue
        base = minimum_sizes[index]
        if not pending_indexes or weight_total <= 0:
            extra = 0
        elif index == pending_indexes[-1]:
            extra = remaining
        else:
            extra = int(round(free_space * (weights[index] / weight_total)))
            extra = min(extra, remaining)
        remaining -= extra
        allocated_sizes.append(base + extra)

    content_main = sum(allocated_sizes) + gap_total
    spare_space = max(inner_main - content_main, 0)
    offset = _resolve_offset(align=node.align, available=spare_space)

    resolved_children: list[ResolvedContent] = []
    cursor = offset
    for child, allocated_main in zip(node.children, allocated_sizes, strict=False):
        fixed_cross = child.width if is_vertical else child.height
        minimum_cross = max(child.min_width if is_vertical else child.min_height, 0)
        available_cross = inner_cross if fixed_cross is None else max(int(fixed_cross), minimum_cross)
        cross_size = min(available_cross, inner_cross) if inner_cross > 0 else available_cross
        child_bbox = _build_linear_child_bbox(
            inner=inner,
            is_vertical=is_vertical,
            main_offset=cursor,
            main_size=allocated_main,
            cross_size=cross_size,
            cross_align=child.cross_align,
            available_cross=inner_cross,
        )
        resolved_children.append(_resolve_node(node=child, bbox=child_bbox))
        cursor += allocated_main + node.gap
    return resolved_children, overflow


def _resolve_grid_children(
    *,
    node: ContentNode,
    inner: BBox,
) -> tuple[list[ResolvedContent], int]:
    child_count = len(node.children)
    if child_count == 0:
        return [], 0
    columns = node.columns
    rows = node.rows
    if columns is None and rows is None:
        columns = max(int(math.ceil(math.sqrt(child_count))), 1)
    if columns is None:
        rows = max(rows or 1, 1)
        columns = max(int(math.ceil(child_count / rows)), 1)
    if rows is None:
        columns = max(columns, 1)
        rows = max(int(math.ceil(child_count / columns)), 1)

    column_weights = _normalize_weights(node.column_weights, count=columns)
    row_weights = _normalize_weights(node.row_weights, count=rows)
    column_widths = _distribute_track_sizes(inner.w, columns, node.gap, column_weights)
    row_heights = _distribute_track_sizes(inner.h, rows, node.gap, row_weights)
    column_positions = _build_track_positions(inner.x, column_widths, node.gap)
    row_positions = _build_track_positions(inner.y, row_heights, node.gap)

    occupied: set[tuple[int, int]] = set()
    placements: list[ResolvedContent] = []
    overflow = 0
    auto_row = 0
    auto_col = 0
    for child in node.children:
        if child.row is not None and child.column is not None:
            row = child.row
            column = child.column
        else:
            row, column, auto_row, auto_col = _next_grid_slot(
                rows=rows,
                columns=columns,
                occupied=occupied,
                start_row=auto_row,
                start_column=auto_col,
            )
        row_span = min(child.row_span, rows - row)
        col_span = min(child.col_span, columns - column)
        if row_span <= 0 or col_span <= 0:
            overflow += 1
            continue
        for row_index in range(row, row + row_span):
            for column_index in range(column, column + col_span):
                occupied.add((row_index, column_index))

        cell_x = column_positions[column]
        cell_y = row_positions[row]
        cell_w = sum(column_widths[column : column + col_span]) + node.gap * max(col_span - 1, 0)
        cell_h = sum(row_heights[row : row + row_span]) + node.gap * max(row_span - 1, 0)
        child_bbox = _fit_child_into_cell(
            child=child,
            cell=BBox(cell_x, cell_y, cell_w, cell_h),
        )
        placements.append(_resolve_node(node=child, bbox=child_bbox))
    return placements, overflow


def _build_linear_child_bbox(
    *,
    inner: BBox,
    is_vertical: bool,
    main_offset: int,
    main_size: int,
    cross_size: int,
    cross_align: str,
    available_cross: int,
) -> BBox:
    cross_offset = _resolve_offset(align=cross_align, available=max(available_cross - cross_size, 0))
    if is_vertical:
        return BBox(
            x=inner.x + cross_offset,
            y=inner.y + main_offset,
            w=cross_size,
            h=main_size,
        )
    return BBox(
        x=inner.x + main_offset,
        y=inner.y + cross_offset,
        w=main_size,
        h=cross_size,
    )


def _fit_child_into_cell(*, child: ContentNode, cell: BBox) -> BBox:
    width = child.width if child.width is not None else cell.w
    height = child.height if child.height is not None else cell.h
    width = max(min(int(width), cell.w), child.min_width)
    height = max(min(int(height), cell.h), child.min_height)
    x = cell.x + _resolve_offset(align=child.align, available=max(cell.w - width, 0))
    y = cell.y + _resolve_offset(align=child.cross_align, available=max(cell.h - height, 0))
    return BBox(x=x, y=y, w=width, h=height)


def _enforce_node_constraints(*, node: ContentNode, bbox: BBox) -> BBox:
    width = bbox.w
    height = bbox.h
    if node.width is not None:
        width = min(width, int(node.width))
    if node.height is not None:
        height = min(height, int(node.height))
    width = max(width, node.min_width)
    height = max(height, node.min_height)
    if node.max_width is not None:
        width = min(width, node.max_width)
    if node.max_height is not None:
        height = min(height, node.max_height)
    x = bbox.x + _resolve_offset(align=node.align, available=max(bbox.w - width, 0))
    y = bbox.y + _resolve_offset(align=node.cross_align, available=max(bbox.h - height, 0))
    return BBox(x=x, y=y, w=width, h=height)


def _apply_padding(bbox: BBox, padding: Insets) -> BBox:
    width = max(bbox.w - padding.left - padding.right, 0)
    height = max(bbox.h - padding.top - padding.bottom, 0)
    return BBox(
        x=bbox.x + padding.left,
        y=bbox.y + padding.top,
        w=width,
        h=height,
    )


def _resolve_offset(*, align: str, available: int) -> int:
    if align == "center":
        return max(available // 2, 0)
    if align == "end":
        return max(available, 0)
    return 0


def _clip_bbox(bbox: BBox, bounds: BBox) -> BBox:
    x = min(max(bbox.x, bounds.x), bounds.x + bounds.w)
    y = min(max(bbox.y, bounds.y), bounds.y + bounds.h)
    max_w = max(bounds.x + bounds.w - x, 0)
    max_h = max(bounds.y + bounds.h - y, 0)
    return BBox(
        x=x,
        y=y,
        w=min(max(bbox.w, 0), max_w),
        h=min(max(bbox.h, 0), max_h),
    )


def _normalize_weights(values: list[float] | None, *, count: int) -> list[float]:
    if values is None:
        return [1.0] * count
    if len(values) != count:
        raise ValueError(f"weight count mismatch: expected {count}, got {len(values)}")
    sanitized = [value if value > 0 else 1.0 for value in values]
    if not any(sanitized):
        return [1.0] * count
    return sanitized


def _distribute_track_sizes(total: int, count: int, gap: int, weights: list[float]) -> list[int]:
    available = max(total - gap * max(count - 1, 0), 0)
    weight_total = sum(weights) or float(count)
    sizes: list[int] = []
    consumed = 0
    for index in range(count):
        if index == count - 1:
            size = max(available - consumed, 0)
        else:
            size = int(round(available * (weights[index] / weight_total)))
            size = min(size, max(available - consumed, 0))
        consumed += size
        sizes.append(size)
    return sizes


def _build_track_positions(start: int, sizes: list[int], gap: int) -> list[int]:
    positions: list[int] = []
    cursor = start
    for size in sizes:
        positions.append(cursor)
        cursor += size + gap
    return positions


def _next_grid_slot(
    *,
    rows: int,
    columns: int,
    occupied: set[tuple[int, int]],
    start_row: int,
    start_column: int,
) -> tuple[int, int, int, int]:
    row = start_row
    column = start_column
    while row < rows:
        while column < columns:
            if (row, column) not in occupied:
                next_row = row
                next_column = column + 1
                if next_column >= columns:
                    next_row += 1
                    next_column = 0
                return row, column, next_row, next_column
            column += 1
        row += 1
        column = 0
    raise ValueError("grid container overflow: children exceed configured rows/columns")


def _render_text(*, slide: Any, leaf: ResolvedContent) -> Any:
    shape = slide.shapes.add_textbox(leaf.bbox.x, leaf.bbox.y, leaf.bbox.w, leaf.bbox.h)
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    style = leaf.style
    margin = style.get("margin")
    if isinstance(margin, int):
        text_frame.margin_top = margin
        text_frame.margin_right = margin
        text_frame.margin_bottom = margin
        text_frame.margin_left = margin
    else:
        margins = _read_insets(style.get("margin"))
        text_frame.margin_top = margins.top
        text_frame.margin_right = margins.right
        text_frame.margin_bottom = margins.bottom
        text_frame.margin_left = margins.left
    text_frame.vertical_anchor = _map_vertical_align(style.get("vertical_align"))
    lines = (leaf.text or "").split("\n")
    first_paragraph = text_frame.paragraphs[0]
    first_paragraph.clear()
    for index, line in enumerate(lines):
        paragraph = first_paragraph if index == 0 else text_frame.add_paragraph()
        run = paragraph.add_run()
        run.text = line
        paragraph.alignment = _map_paragraph_align(style.get("align"))
        _apply_run_style(run=run, style=style)
    return shape


def _render_image(*, slide: Any, leaf: ResolvedContent) -> Any:
    if leaf.path is None:
        raise ValueError("image/svg leaf missing path")
    path = Path(leaf.path)
    left, top, width, height = _resolve_image_box(path=path, bbox=leaf.bbox, fit=leaf.fit)
    return slide.shapes.add_picture(str(path), left, top, width, height)


def _resolve_image_box(*, path: Path, bbox: BBox, fit: str) -> tuple[int, int, int, int]:
    if fit not in {"fill", "contain", "cover"}:
        fit = "fill"
    if fit == "fill" or path.suffix.lower() == ".svg":
        return bbox.x, bbox.y, bbox.w, bbox.h
    with Image.open(path) as image:
        image_width, image_height = image.size
    if image_width <= 0 or image_height <= 0 or bbox.w <= 0 or bbox.h <= 0:
        return bbox.x, bbox.y, bbox.w, bbox.h
    image_ratio = image_width / image_height
    box_ratio = bbox.w / bbox.h
    if (fit == "contain" and image_ratio >= box_ratio) or (fit == "cover" and image_ratio <= box_ratio):
        width = bbox.w
        height = int(round(width / image_ratio))
    else:
        height = bbox.h
        width = int(round(height * image_ratio))
    left = bbox.x + max((bbox.w - width) // 2, 0)
    top = bbox.y + max((bbox.h - height) // 2, 0)
    return left, top, width, height


def _apply_run_style(*, run: Any, style: dict[str, Any]) -> None:
    font = run.font
    if "font_name" in style:
        font.name = str(style["font_name"])
    if "font_size" in style:
        font.size = Pt(float(style["font_size"]))
    if "bold" in style:
        font.bold = bool(style["bold"])
    if "italic" in style:
        font.italic = bool(style["italic"])
    if "underline" in style:
        font.underline = bool(style["underline"])
    color = style.get("color")
    rgb = _parse_color(color)
    if rgb is not None:
        font.color.rgb = rgb


def _map_paragraph_align(value: Any) -> PP_ALIGN | None:
    mapping = {
        "start": PP_ALIGN.LEFT,
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "end": PP_ALIGN.RIGHT,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    if value is None:
        return None
    return mapping.get(str(value).strip().lower())


def _map_vertical_align(value: Any) -> MSO_ANCHOR | None:
    mapping = {
        "top": MSO_ANCHOR.TOP,
        "start": MSO_ANCHOR.TOP,
        "middle": MSO_ANCHOR.MIDDLE,
        "center": MSO_ANCHOR.MIDDLE,
        "bottom": MSO_ANCHOR.BOTTOM,
        "end": MSO_ANCHOR.BOTTOM,
    }
    if value is None:
        return None
    return mapping.get(str(value).strip().lower())


def _parse_color(value: Any) -> RGBColor | None:
    if value is None:
        return None
    raw = str(value).strip().lstrip("#")
    if len(raw) != 6:
        raise ValueError(f"invalid hex color: {value}")
    try:
        return RGBColor.from_string(raw.upper())
    except ValueError as exc:
        raise ValueError(f"invalid hex color: {value}") from exc


def _resolve_asset_path(raw_value: str) -> str:
    candidate = Path(raw_value).expanduser().resolve()
    if not candidate.exists():
        raise ValueError(f"content asset does not exist: {candidate}")
    return str(candidate)


def _read_leaf_value(payload: dict[str, Any], keys: tuple[str, ...]) -> str:
    for key in keys:
        value = payload.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    raise ValueError(f"content node requires one of: {', '.join(keys)}")


def _read_style(value: Any) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError("content style must be an object")
    return dict(value)


def _read_insets(value: Any) -> Insets:
    if value is None:
        return Insets()
    if isinstance(value, (int, float)):
        inset = int(value)
        return Insets(top=inset, right=inset, bottom=inset, left=inset)
    if not isinstance(value, dict):
        raise ValueError("padding/margin must be a number or object")
    return Insets(
        top=_read_int(value, "top", default=0),
        right=_read_int(value, "right", default=0),
        bottom=_read_int(value, "bottom", default=0),
        left=_read_int(value, "left", default=0),
    )


def _read_align(value: Any, *, default: str) -> str:
    if value is None:
        return default
    align = str(value).strip().lower()
    if align not in ALIGN_VALUES:
        raise ValueError(f"unsupported alignment: {value}")
    return align


def _read_optional_str(payload: dict[str, Any], key: str) -> str | None:
    value = payload.get(key)
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _read_optional_int(payload: dict[str, Any], key: str) -> int | None:
    value = payload.get(key)
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"{key} must be an integer")
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str) and value.strip():
        try:
            return int(float(value))
        except ValueError as exc:
            raise ValueError(f"{key} must be an integer") from exc
    raise ValueError(f"{key} must be an integer")


def _read_int(payload: dict[str, Any], key: str, default: int | None = None) -> int:
    value = payload.get(key, default)
    if value is None:
        raise ValueError(f"missing required integer field: {key}")
    result = _read_optional_int({key: value}, key)
    if result is None:
        raise ValueError(f"missing required integer field: {key}")
    return result


def _read_float(payload: dict[str, Any], key: str, default: float = 0.0) -> float:
    value = payload.get(key, default)
    if isinstance(value, bool):
        raise ValueError(f"{key} must be a number")
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.strip():
        try:
            return float(value)
        except ValueError as exc:
            raise ValueError(f"{key} must be a number") from exc
    raise ValueError(f"{key} must be a number")


def _read_optional_float_list(value: Any) -> list[float] | None:
    if value is None:
        return None
    if not isinstance(value, list):
        raise ValueError("weight list must be an array")
    return [float(item) for item in value]
