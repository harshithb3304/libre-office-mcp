"""Impress: text boxes, shapes, lines, images and shape arrangement."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from com.sun.star.awt import Point, Size

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.ops.impress.formatting import apply_text_format
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, file_url, pv

_PX_TO_MM100 = 26.46  # one pixel at 96 DPI, in 1/100 mm

#: tool-level shape name -> LibreOffice custom-shape preset ("rectangle" is a plain
#: RectangleShape so that it supports corner radius).
SHAPE_TYPES = {
    "rectangle": None,
    "rounded_rectangle": "round-rectangle",
    "ellipse": "ellipse",
    "triangle": "isosceles-triangle",
    "right_triangle": "right-triangle",
    "diamond": "diamond",
    "pentagon": "pentagon",
    "hexagon": "hexagon",
    "octagon": "octagon",
    "parallelogram": "parallelogram",
    "trapezoid": "trapezoid",
    "star": "star5",
    "heart": "heart",
    "cross": "cross",
    "right_arrow": "right-arrow",
    "left_arrow": "left-arrow",
    "up_arrow": "up-arrow",
    "down_arrow": "down-arrow",
    "left_right_arrow": "left-right-arrow",
    "chevron": "chevron",
    "callout": "rectangular-callout",
    "speech_bubble": "round-callout",
    "cylinder": "can",
    "cube": "cube",
    "sun": "sun",
    "moon": "moon",
    "smiley": "smiley",
    "lightning_bolt": "lightning",
    "thought_bubble": "cloud-callout",
    "ring": "ring",
    "up_down_arrow": "up-down-arrow",
    "flowchart_process": "flowchart-process",
    "flowchart_decision": "flowchart-decision",
    "flowchart_terminator": "flowchart-terminator",
    "flowchart_document": "flowchart-document",
    "flowchart_data": "flowchart-data",
}


def _anchor(name: str | None) -> Any:
    if name not in ("top", "middle", "bottom"):
        raise OpError("vertical_anchor must be 'top', 'middle' or 'bottom'.")
    return enum(
        "com.sun.star.drawing.TextVerticalAdjust",
        {"top": "TOP", "middle": "CENTER", "bottom": "BOTTOM"}[name],
    )


def _style_text(
    s: Session,
    shape: Any,
    font: str | None,
    font_size: float | None,
    bold: bool | None,
    italic: bool | None,
    color: str | None,
    alignment: str | None,
    bullets: str | None = None,
) -> None:
    if any(v is not None for v in (font, font_size, bold, italic, color, alignment, bullets)):
        apply_text_format(
            s.doc, shape, font=font, size=font_size, bold=bold, italic=italic, color=color,
            alignment=alignment, bullets=bullets,
        )  # fmt: skip


# -- text boxes ---------------------------------------------------------------------------


@op("impress.add_text_box", kind="impress", mutates=True)
def add_text_box(
    s: Session,
    slide_index: int,
    x_mm: float,
    y_mm: float,
    width_mm: float,
    height_mm: float = 20,
    text: str | None = None,
    paragraphs: list[Any] | None = None,
    font: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    alignment: str | None = None,
    vertical_anchor: str | None = None,
    bullets: str | None = None,
    fill_color: str | None = None,
    line_color: str | None = None,
    name: str | None = None,
) -> dict[str, Any]:
    """Add a text box. The box grows downwards if the text needs more room."""
    if (text is None) == (paragraphs is None):
        raise OpError("Give either 'text' or 'paragraphs' (not both).")
    items = h.normalize_items(paragraphs if paragraphs is not None else text, "paragraphs")
    slide = h.slide_at(s.doc, slide_index)
    shape = h.create_shape(s, "com.sun.star.drawing.TextShape", slide)
    h.set_geometry(shape, x_mm, y_mm, width_mm, height_mm)
    h.set_props(shape, {"TextWordWrap": True, "TextAutoGrowHeight": True})
    h.set_paragraphs(shape, items)
    if vertical_anchor is not None:
        shape.TextVerticalAdjust = _anchor(vertical_anchor)
    _style_text(s, shape, font, font_size, bold, italic, color, alignment, bullets)
    h.apply_look(shape, fill_color=fill_color, line_color=line_color)
    h.apply_alt_text(shape, None, name)
    return {
        "slide_index": slide_index,
        "shape_index": slide.getCount() - 1,
        **h.geometry_mm(shape),
    }


# -- shapes -------------------------------------------------------------------------------


@op("impress.add_shape", kind="impress", mutates=True)
def add_shape(
    s: Session,
    slide_index: int,
    shape: str,
    x_mm: float,
    y_mm: float,
    width_mm: float,
    height_mm: float,
    text: str | None = None,
    fill_color: str | None = None,
    line_color: str | None = None,
    line_width_mm: float | None = None,
    no_line: bool = False,
    gradient_end_color: str | None = None,
    transparency: int | None = None,
    shadow: bool | None = None,
    text_color: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    alignment: str | None = "center",
    name: str | None = None,
) -> dict[str, Any]:
    """Add a rectangle, ellipse, arrow, star, callout... optionally with text inside."""
    if shape not in SHAPE_TYPES:
        raise OpError(f"Unknown shape {shape!r}. Valid shapes: {', '.join(SHAPE_TYPES)}.")
    slide = h.slide_at(s.doc, slide_index)
    preset = SHAPE_TYPES[shape]
    obj = h.create_shape(
        s, "com.sun.star.drawing.CustomShape" if preset else "com.sun.star.drawing.RectangleShape",
        slide,
    )  # fmt: skip
    h.set_geometry(obj, x_mm, y_mm, width_mm, height_mm)
    if preset:
        obj.CustomShapeGeometry = (pv("Type", preset),)
    h.apply_look(
        obj, fill_color=fill_color, line_color=line_color, line_width_mm=line_width_mm,
        no_line=no_line, gradient_end_color=gradient_end_color, transparency=transparency,
        shadow=shadow,
    )  # fmt: skip
    if text:
        obj.setString(text)
        # Explicit centring: LibreOffice centres shape text by default but does not write
        # that to .pptx, where the text would come back left-aligned.
        _style_text(s, obj, None, font_size, bold, None, text_color, alignment)
    h.apply_alt_text(obj, None, name)
    return {"slide_index": slide_index, "shape_index": slide.getCount() - 1, **h.geometry_mm(obj)}


@op("impress.add_line", kind="impress", mutates=True)
def add_line(
    s: Session,
    slide_index: int,
    x1_mm: float,
    y1_mm: float,
    x2_mm: float,
    y2_mm: float,
    color: str = "#333333",
    width_mm: float = 0.6,
    dash: str = "solid",
    arrow_start: bool = False,
    arrow_end: bool = False,
    name: str | None = None,
) -> dict[str, Any]:
    """Add a straight line or arrow from (x1, y1) to (x2, y2)."""
    if (x1_mm, y1_mm) == (x2_mm, y2_mm):
        raise OpError("The line's start and end points are identical.")
    if dash not in ("solid", "dashed", "dotted"):
        raise OpError("dash must be 'solid', 'dashed' or 'dotted'.")
    slide = h.slide_at(s.doc, slide_index)
    line = h.create_shape(s, "com.sun.star.drawing.LineShape", slide)
    # A LineShape's Position/Size only describe a top-left -> bottom-right diagonal; the
    # polygon is what lets a line run in any direction.
    line.PolyPolygon = (
        (Point(h.mm(x1_mm), h.mm(y1_mm)), Point(h.mm(x2_mm), h.mm(y2_mm))),
    )  # fmt: skip
    h.apply_look(line, line_color=color, line_width_mm=width_mm, dash=dash)
    arrow_size = max(300, round(h.mm(width_mm) * 5))
    if arrow_start:
        h.set_props(line, {"LineStartName": "Arrow", "LineStartWidth": arrow_size})
    if arrow_end:
        h.set_props(line, {"LineEndName": "Arrow", "LineEndWidth": arrow_size})
    h.apply_alt_text(line, None, name)
    return {"slide_index": slide_index, "shape_index": slide.getCount() - 1, **h.geometry_mm(line)}


# -- images -------------------------------------------------------------------------------


def load_graphic(s: Session, image_path: str) -> tuple[Any, float, float]:
    """Load an image file; returns (graphic, native width, native height) in 1/100 mm."""
    provider = s.service("com.sun.star.graphic.GraphicProvider")
    graphic = provider.queryGraphic((pv("URL", file_url(image_path)),))
    if graphic is None:
        raise OpError(
            f"{Path(image_path).name} is not an image LibreOffice can read (try PNG, JPG, GIF "
            "or SVG).",
            protocol.E_UNSUPPORTED,
        )
    size = graphic.Size100thMM
    if size.Width > 0 and size.Height > 0:
        return graphic, float(size.Width), float(size.Height)
    pixels = graphic.SizePixel  # no DPI in the file: assume 96 DPI
    return graphic, pixels.Width * _PX_TO_MM100, pixels.Height * _PX_TO_MM100


def insert_image(
    s: Session,
    slide: Any,
    image_path: str,
    area: tuple[int, int, int, int],
    x_mm: float | None = None,
    y_mm: float | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    alt_text: str | None = None,
    name: str | None = None,
) -> Any:
    """Embed a picture. With no size it is fitted (aspect kept) into ``area``; with one
    dimension the other follows; x/y default to centring within ``area``."""
    graphic, native_w, native_h = load_graphic(s, image_path)
    ratio = native_h / native_w
    if width_mm and height_mm:
        width, height = h.mm(width_mm), h.mm(height_mm)
    elif width_mm:
        width = h.mm(width_mm)
        height = round(width * ratio)
    elif height_mm:
        height = h.mm(height_mm)
        width = round(height / ratio)
    else:
        _, _, width, height = h.fit_into(native_w, native_h, area)
    ax, ay, aw, ah = area
    x = h.mm(x_mm) if x_mm is not None else ax + (aw - width) // 2
    y = h.mm(y_mm) if y_mm is not None else ay + (ah - height) // 2
    picture = h.create_shape(s, "com.sun.star.drawing.GraphicObjectShape", slide)
    picture.Graphic = graphic  # embeds the pixels in the document (not a link to the file)
    picture.Size = Size(width, height)
    picture.Position = Point(x, y)
    h.apply_alt_text(picture, alt_text, name)
    return picture


@op("impress.add_image", kind="impress", mutates=True)
def add_image(
    s: Session,
    slide_index: int,
    image_path: str,
    x_mm: float | None = None,
    y_mm: float | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    alt_text: str | None = None,
    name: str | None = None,
) -> dict[str, Any]:
    slide = h.slide_at(s.doc, slide_index)
    area = h.content_area(s.doc, slide)
    picture = insert_image(
        s, slide, image_path, area, x_mm, y_mm, width_mm, height_mm, alt_text, name
    )
    return {
        "slide_index": slide_index,
        "shape_index": slide.getCount() - 1,
        **h.geometry_mm(picture),
    }


# -- arranging ----------------------------------------------------------------------------


@op("impress.set_shape_geometry", kind="impress", mutates=True)
def set_shape_geometry(
    s: Session,
    slide_index: int,
    shape_index: int,
    x_mm: float | None = None,
    y_mm: float | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    rotation_deg: float | None = None,
    keep_aspect_ratio: bool = False,
) -> dict[str, Any]:
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    if all(v is None for v in (x_mm, y_mm, width_mm, height_mm, rotation_deg)):
        raise OpError("Give at least one of x_mm, y_mm, width_mm, height_mm, rotation_deg.")
    for label, value in (("width_mm", width_mm), ("height_mm", height_mm)):
        if value is not None and value <= 0:
            raise OpError(f"{label} must be positive.")
    if keep_aspect_ratio and (width_mm is None) != (height_mm is None):
        size = shape.Size
        if width_mm is not None:
            height_mm = h.to_mm(h.mm(width_mm) * size.Height / size.Width)
        else:
            width_mm = h.to_mm(h.mm(height_mm) * size.Width / size.Height)
    h.set_geometry(shape, x_mm, y_mm, width_mm, height_mm)
    if rotation_deg is not None:
        # Positive = clockwise (the convention people expect). LibreOffice counts counter-clockwise.
        shape.RotateAngle = round(-rotation_deg * 100) % 36000
    return {"slide_index": slide_index, "shape_index": shape_index, **h.geometry_mm(shape)}


@op("impress.set_shape_order", kind="impress", mutates=True)
def set_shape_order(s: Session, slide_index: int, shape_index: int, action: str) -> dict[str, Any]:
    """Change the stacking order. Returns the shape's new index."""
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    if action not in ("front", "back", "forward", "backward"):
        raise OpError("action must be front, back, forward or backward.")
    top = slide.getCount() - 1
    target = {
        "front": top,
        "back": 0,
        "forward": min(top, shape_index + 1),
        "backward": max(0, shape_index - 1),
    }[action]
    if target == shape_index:
        return {"slide_index": slide_index, "shape_index": shape_index, "_noop": True}
    shape.ZOrder = target
    return {"slide_index": slide_index, "shape_index": target}


@op("impress.delete_shapes", kind="impress", mutates=True)
def delete_shapes(s: Session, slide_index: int, shape_indexes: list[int]) -> dict[str, Any]:
    """Delete shapes by index (all indexes refer to the slide as it is now)."""
    slide = h.slide_at(s.doc, slide_index)
    if not shape_indexes:
        raise OpError("shape_indexes must not be empty.")
    wanted = sorted(set(shape_indexes), reverse=True)  # back to front keeps indexes valid
    shapes = [h.shape_at(slide, i, slide_index) for i in wanted]
    for shape in shapes:
        slide.remove(shape)
    return {"deleted": len(shapes), "shapes_remaining": slide.getCount()}


@op("impress.duplicate_shape", kind="impress", mutates=True)
def duplicate_shape(
    s: Session,
    slide_index: int,
    shape_index: int,
    target_slide_index: int | None = None,
    offset_x_mm: float | None = None,
    offset_y_mm: float | None = None,
    count: int = 1,
) -> dict[str, Any]:
    """Copy a shape (to the same slide, offset, or to another slide at the same position)."""
    doc = s.doc
    source = h.slide_at(doc, slide_index)
    shape = h.shape_at(source, shape_index, slide_index)
    target = source if target_slide_index is None else h.slide_at(doc, target_slide_index)
    if not 1 <= count <= 50:
        raise OpError("count must be between 1 and 50.")
    same = target is source or target_slide_index == slide_index
    step_x = offset_x_mm if offset_x_mm is not None else (10.0 if same else 0.0)
    step_y = offset_y_mm if offset_y_mm is not None else (10.0 if same else 0.0)
    new_indexes = []
    for n in range(1, count + 1):
        (copy,) = h.copy_shapes(s, [shape], source, target)
        pos = copy.Position
        copy.Position = Point(pos.X + h.mm(step_x * n), pos.Y + h.mm(step_y * n))
        new_indexes.append(target.getCount() - 1)
    return {
        "slide_index": slide_index if target_slide_index is None else target_slide_index,
        "new_shape_indexes": new_indexes,
    }


@op("impress.align_shapes", kind="impress", mutates=True)
def align_shapes(
    s: Session,
    slide_index: int,
    shape_indexes: list[int],
    align: str | None = None,
    distribute: str | None = None,
    relative_to: str = "selection",
) -> dict[str, Any]:
    """Align shapes to each other (or the slide) and/or space them evenly."""
    slide = h.slide_at(s.doc, slide_index)
    if align is None and distribute is None:
        raise OpError("Give align and/or distribute.")
    if align not in (None, "left", "center", "right", "top", "middle", "bottom"):
        raise OpError("align must be left, center, right, top, middle or bottom.")
    if distribute not in (None, "horizontal", "vertical"):
        raise OpError("distribute must be 'horizontal' or 'vertical'.")
    if relative_to not in ("selection", "slide"):
        raise OpError("relative_to must be 'selection' or 'slide'.")
    indexes = list(dict.fromkeys(shape_indexes))
    if len(indexes) < (3 if distribute else 2) and not (align and relative_to == "slide"):
        need = "three" if distribute else "two"
        raise OpError(f"Give at least {need} different shape indexes (or relative_to='slide').")
    shapes = [h.shape_at(slide, i, slide_index) for i in indexes]
    boxes = [(sh.Position.X, sh.Position.Y, sh.Size.Width, sh.Size.Height) for sh in shapes]
    if relative_to == "slide":
        left, top, right, bottom = 0, 0, slide.Width, slide.Height
    else:
        left = min(b[0] for b in boxes)
        top = min(b[1] for b in boxes)
        right = max(b[0] + b[2] for b in boxes)
        bottom = max(b[1] + b[3] for b in boxes)
    places = [[x, y] for x, y, _, _ in boxes]
    for place, (_, _, w, hh) in zip(places, boxes, strict=True):
        if align == "left":
            place[0] = left
        elif align == "right":
            place[0] = right - w
        elif align == "center":
            place[0] = left + (right - left - w) // 2
        elif align == "top":
            place[1] = top
        elif align == "bottom":
            place[1] = bottom - hh
        elif align == "middle":
            place[1] = top + (bottom - top - hh) // 2
    if distribute:
        axis = 0 if distribute == "horizontal" else 1
        extent = 2 if axis == 0 else 3
        order = sorted(range(len(boxes)), key=lambda i: places[i][axis])
        first, last = order[0], order[-1]
        span_start = places[first][axis]
        span_end = places[last][axis] + boxes[last][extent]
        gap = (span_end - span_start - sum(boxes[i][extent] for i in order)) / (len(order) - 1)
        cursor = span_start
        for i in order:
            places[i][axis] = round(cursor)
            cursor += boxes[i][extent] + gap
    for shape, (x, y) in zip(shapes, places, strict=True):
        shape.Position = Point(int(x), int(y))
    return {"slide_index": slide_index, "shapes_moved": len(shapes)}


@op("impress.set_shape_description", kind="impress", mutates=True)
def set_shape_description(
    s: Session,
    slide_index: int,
    shape_index: int,
    name: str | None = None,
    alt_text: str | None = None,
) -> dict[str, Any]:
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    if name is None and alt_text is None:
        raise OpError("Give a name and/or alt_text.")
    h.apply_alt_text(shape, alt_text, name)
    return {"slide_index": slide_index, "shape_index": shape_index}
