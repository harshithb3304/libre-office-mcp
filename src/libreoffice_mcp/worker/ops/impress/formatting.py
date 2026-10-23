"""Impress: text formatting, shape styling and hyperlinks."""

from __future__ import annotations

from typing import Any

import uno

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color, pv, struct_any

_ARABIC = 4  # com.sun.star.style.NumberingType.ARABIC


# -- bullets -----------------------------------------------------------------------


def _outline_rules(doc: Any, numbered: bool) -> Any:
    """A NumberingRules object based on the master's level-1 outline style (so bullets keep
    the deck's look), optionally switched to '1. 2. 3.' numbering."""
    master = doc.getMasterPages().getByIndex(0).Name
    rules = doc.getStyleFamilies().getByName(master).getByName("outline1").NumberingRules
    if numbered:
        for level in range(rules.getCount()):
            # pyuno turns a plain tuple into sequence<any>; the rules object insists on
            # sequence<PropertyValue>, and 16-bit properties need an explicit "short".
            change = struct_any(
                "[]com.sun.star.beans.PropertyValue",
                (
                    pv("NumberingType", uno.Any("short", _ARABIC)),
                    pv("Suffix", "."),
                    pv("StartWith", uno.Any("short", 1)),
                    pv("BulletRelSize", uno.Any("short", 100)),  # digits as big as the text
                ),
            )
            uno.invoke(rules, "replaceByIndex", (level, change))
    return rules


def apply_bullets(doc: Any, paragraphs: list[Any], kind: str, in_placeholder: bool) -> None:
    """Switch paragraphs to bullets / numbers / plain text.

    Body placeholders get their bullets from the master's outline style, so "bullets" only has
    to switch the style's numbering back on; free text boxes have no list style and get an
    explicit one copied from the outline style.
    """
    if kind not in ("bullets", "numbered", "none"):
        raise OpError("bullets must be 'bullets', 'numbered' or 'none'.")
    rules = None
    if kind == "numbered" or (kind == "bullets" and not in_placeholder):
        rules = _outline_rules(doc, numbered=kind == "numbered")
    for para in paragraphs:
        if kind == "none":
            para.NumberingIsNumber = False
            continue
        if rules is not None:
            para.NumberingRules = rules
        if para.NumberingLevel is None:  # no level at all means "no bullet": start at level 0
            para.NumberingLevel = 0
        para.NumberingIsNumber = True


# -- text formatting --------------------------------------------------------------------


def _char_values(
    font: str | None,
    size: float | None,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    strikethrough: bool | None,
    color: str | None,
    highlight_color: str | None,
) -> dict[str, Any]:
    values: dict[str, Any] = {}
    if font is not None:
        values["CharFontName"] = font
    if size is not None:
        if not 1 <= size <= 1000:
            raise OpError("size must be between 1 and 1000 points.")
        values["CharHeight"] = float(size)
    if bold is not None:
        values["CharWeight"] = 150.0 if bold else 100.0
    if italic is not None:
        values["CharPosture"] = enum("com.sun.star.awt.FontSlant", "ITALIC" if italic else "NONE")
    if underline is not None:
        values["CharUnderline"] = 1 if underline else 0
    if strikethrough is not None:
        values["CharStrikeout"] = 1 if strikethrough else 0
    if color is not None:
        values["CharColor"] = parse_color(color)
    if highlight_color is not None:
        values["CharBackColor"] = parse_color(highlight_color)
        values["CharBackTransparent"] = False
    return values


def _para_values(
    alignment: str | None,
    line_spacing_percent: float | None,
    space_before_mm: float | None,
    space_after_mm: float | None,
    level: int | None,
) -> dict[str, Any]:
    values: dict[str, Any] = {}
    if alignment is not None:
        if alignment not in h.ALIGN_TO_UNO:
            raise OpError(f"alignment must be one of: {', '.join(h.ALIGN_TO_UNO)}.")
        values["ParaAdjust"] = enum("com.sun.star.style.ParagraphAdjust", h.ALIGN_TO_UNO[alignment])
    if line_spacing_percent is not None:
        if not 50 <= line_spacing_percent <= 400:
            raise OpError("line_spacing_percent must be between 50 and 400 (100 = single).")
        spacing = uno.createUnoStruct("com.sun.star.style.LineSpacing")
        spacing.Mode = 0  # PROP: Height is a percentage
        spacing.Height = round(line_spacing_percent)
        values["ParaLineSpacing"] = spacing
    if space_before_mm is not None:
        values["ParaTopMargin"] = h.mm(space_before_mm)
    if space_after_mm is not None:
        values["ParaBottomMargin"] = h.mm(space_after_mm)
    if level is not None:
        if not 0 <= level <= 8:
            raise OpError("level must be 0 (top level) .. 8.")
        values["NumberingLevel"] = level
    return values


def _ranges_for_text(shape: Any, find: str) -> list[Any]:
    """Text cursors covering each occurrence of ``find`` (case-sensitive) in the shape."""
    text = shape.getString()
    ranges = []
    start = text.find(find)
    while start != -1:
        cursor = shape.createTextCursor()
        cursor.gotoStart(False)
        _advance(cursor, start, False)
        _advance(cursor, len(find), True)
        ranges.append(cursor)
        start = text.find(find, start + len(find))
    return ranges


def _advance(cursor: Any, count: int, select: bool) -> None:
    """Move a cursor right by count characters (optionally extending the selection)."""
    while count > 0:  # goRight takes a 16-bit count
        step = min(count, 30000)
        cursor.goRight(step, select)
        count -= step


def apply_text_format(
    doc: Any,
    shape: Any,
    *,
    find: str | None = None,
    paragraph_index: int | None = None,
    font: str | None = None,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    color: str | None = None,
    highlight_color: str | None = None,
    alignment: str | None = None,
    line_spacing_percent: float | None = None,
    space_before_mm: float | None = None,
    space_after_mm: float | None = None,
    bullets: str | None = None,
    level: int | None = None,
) -> dict[str, Any]:
    """Format text in a shape: all of it, one paragraph, or every occurrence of ``find``."""
    if not h.shape_has_text_support(shape):
        raise OpError(
            f"This shape ({h.shape_kind(shape)}) has no text to format. For tables use "
            "impress_set_table_cells."
        )
    char = _char_values(font, size, bold, italic, underline, strikethrough, color, highlight_color)
    para = _para_values(alignment, line_spacing_percent, space_before_mm, space_after_mm, level)
    if not char and not para and bullets is None:
        raise OpError("Nothing to change: give at least one formatting option.")
    text = shape.getString()
    if not text:
        raise OpError("The shape has no text yet. Set its text first (impress_set_shape_text).")
    result: dict[str, Any] = {}
    if find is not None:
        if para or bullets is not None:
            raise OpError(
                "With 'find' only character formatting (font, size, bold, italic, underline, "
                "color...) can be applied; paragraph options (alignment, spacing, bullets, "
                "level) apply to whole paragraphs, so omit 'find' for those."
            )
        ranges = _ranges_for_text(shape, find)
        if not ranges:
            raise OpError(
                f"The text {find!r} was not found in this shape (search is case-sensitive)."
            )
        for cursor in ranges:
            h.set_props(cursor, char)
        result["occurrences_formatted"] = len(ranges)
        return result
    paragraphs = h.paragraphs_of(shape)
    if paragraph_index is not None:
        if not 0 <= paragraph_index < len(paragraphs):
            raise OpError(
                f"paragraph_index {paragraph_index} is out of range: the shape has "
                f"{len(paragraphs)} paragraph(s) (valid: 0..{len(paragraphs) - 1})."
            )
        paragraphs = [paragraphs[paragraph_index]]
    for paragraph in paragraphs:
        h.set_props(paragraph, {**char, **para})
    if bullets is not None:
        apply_bullets(doc, paragraphs, bullets, h.shape_kind(shape) == "body")
    result["paragraphs_formatted"] = len(paragraphs)
    return result


@op("impress.format_text", kind="impress", mutates=True)
def format_text(
    s: Session,
    slide_index: int,
    shape_index: int,
    find: str | None = None,
    paragraph_index: int | None = None,
    font: str | None = None,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    color: str | None = None,
    highlight_color: str | None = None,
    alignment: str | None = None,
    line_spacing_percent: float | None = None,
    space_before_mm: float | None = None,
    space_after_mm: float | None = None,
    bullets: str | None = None,
    level: int | None = None,
) -> dict[str, Any]:
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    result = apply_text_format(
        s.doc, shape, find=find, paragraph_index=paragraph_index, font=font, size=size,
        bold=bold, italic=italic, underline=underline, strikethrough=strikethrough, color=color,
        highlight_color=highlight_color, alignment=alignment,
        line_spacing_percent=line_spacing_percent, space_before_mm=space_before_mm,
        space_after_mm=space_after_mm, bullets=bullets, level=level,
    )  # fmt: skip
    return {"slide_index": slide_index, "shape_index": shape_index, **result}


# -- shape styling -----------------------------------------------------------------------


@op("impress.format_shape", kind="impress", mutates=True)
def format_shape(
    s: Session,
    slide_index: int,
    shape_index: int,
    fill_color: str | None = None,
    no_fill: bool = False,
    gradient_end_color: str | None = None,
    gradient_angle: int | None = None,
    transparency: int | None = None,
    line_color: str | None = None,
    line_width_mm: float | None = None,
    no_line: bool = False,
    dash: str | None = None,
    shadow: bool | None = None,
    corner_radius_mm: float | None = None,
    text_anchor: str | None = None,
    margin_mm: float | None = None,
) -> dict[str, Any]:
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    kind = h.shape_kind(shape)
    if kind in ("table", "group", "object", "chart"):
        raise OpError(f"Shapes of type '{kind}' cannot be styled with this tool.")
    if dash is not None and dash not in ("solid", "dashed", "dotted"):
        raise OpError("dash must be 'solid', 'dashed' or 'dotted'.")
    if fill_color is not None and no_fill:
        raise OpError("Give fill_color or no_fill, not both.")
    if line_color is not None and no_line:
        raise OpError("Give line_color or no_line, not both.")
    changed = h.apply_look(
        shape, fill_color=fill_color, no_fill=no_fill, line_color=line_color,
        line_width_mm=line_width_mm, no_line=no_line, gradient_end_color=gradient_end_color,
        gradient_angle=gradient_angle, transparency=transparency, shadow=shadow, dash=dash,
    )  # fmt: skip
    extra: dict[str, Any] = {}
    if corner_radius_mm is not None:
        if shape.ShapeType == "com.sun.star.drawing.RectangleShape":
            extra["CornerRadius"] = h.mm(corner_radius_mm)
            changed.append("corner_radius")
        else:
            raise OpError(
                "corner_radius_mm only applies to plain rectangles; add a shape of type "
                "'rounded_rectangle' for rounded corners."
            )
    if text_anchor is not None:
        if text_anchor not in ("top", "middle", "bottom"):
            raise OpError("text_anchor must be 'top', 'middle' or 'bottom'.")
        extra["TextVerticalAdjust"] = enum(
            "com.sun.star.drawing.TextVerticalAdjust",
            {"top": "TOP", "middle": "CENTER", "bottom": "BOTTOM"}[text_anchor],
        )
        changed.append("text_anchor")
    if margin_mm is not None:
        margin = h.mm(margin_mm)
        extra.update(TextLeftDistance=margin, TextRightDistance=margin)
        extra.update(TextUpperDistance=margin, TextLowerDistance=margin)
        changed.append("margin")
    h.set_props(shape, extra)
    if not changed:
        raise OpError("Nothing to change: give at least one styling option.")
    return {"slide_index": slide_index, "shape_index": shape_index, "changed": changed}


# -- hyperlinks ----------------------------------------------------------------------------


@op("impress.set_hyperlink", kind="impress", mutates=True)
def set_hyperlink(
    s: Session,
    slide_index: int,
    shape_index: int,
    url: str,
    find: str | None = None,
) -> dict[str, Any]:
    """Link text (``find``) or make the whole shape clickable; url '' removes a shape link."""
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    if find is None:
        click = "DOCUMENT" if url else "NONE"
        h.set_props(
            shape,
            {"OnClick": enum("com.sun.star.presentation.ClickAction", click), "Bookmark": url},
        )
        return {"slide_index": slide_index, "shape_index": shape_index, "shape_link": url or None}
    if not url:
        raise OpError("To remove a text link, retype the text with impress_set_shape_text.")
    if not h.shape_has_text_support(shape):
        raise OpError(f"This shape ({h.shape_kind(shape)}) has no text to link.")
    ranges = _ranges_for_text(shape, find)
    if not ranges:
        raise OpError(f"The text {find!r} was not found in this shape (search is case-sensitive).")
    # Inserting a field replaces the selected text and shifts later offsets, so go back to
    # front: earlier ranges are still valid.
    for cursor in reversed(ranges):
        field = s.new("com.sun.star.text.TextField.URL")
        field.URL = url
        field.Representation = cursor.getString()
        field.TargetFrame = "_blank"
        shape.insertTextContent(cursor, field, True)
    return {"slide_index": slide_index, "shape_index": shape_index, "links_added": len(ranges)}
