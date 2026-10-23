"""Impress: tables on slides."""

from __future__ import annotations

from typing import Any

import uno
from com.sun.star.awt import Point, Size

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color

TABLE_SERVICE = "com.sun.star.drawing.TableShape"
HEADER_FILL = "#1F4E79"
BAND_FILL = "#EAF0F7"
BORDER_COLOR = "#BFC9D6"
_MAX_CELLS = 2000


def table_shape_at(slide: Any, shape_index: int, slide_index: int) -> Any:
    shape = h.shape_at(slide, shape_index, slide_index)
    if h.shape_kind(shape) != "table":
        raise OpError(
            f"Shape {shape_index} on slide {slide_index} is a {h.shape_kind(shape)}, not a "
            "table. Use impress_get_slide to find tables."
        )
    return shape


def table_grid(shape: Any, max_rows: int | None = None) -> list[list[str]]:
    """Cell text as a row-major grid, optionally only the first ``max_rows`` rows."""
    model = shape.Model
    rows, cols = model.Rows.Count, model.Columns.Count
    return [
        [model.getCellByPosition(c, r).getString() for c in range(cols)]
        for r in range(rows if max_rows is None else min(rows, max_rows))
    ]


def _border(color: int) -> Any:
    line = uno.createUnoStruct("com.sun.star.table.BorderLine")
    line.Color = color
    line.OuterLineWidth = 1  # 1/100 mm hairline
    return line


def row_min_mm(font_size: float) -> float:
    """Height of a one-line row: line height plus LibreOffice's default cell padding."""
    return font_size * 0.3528 * 1.25 + 3.0


def _auto_font_size(rows: int, height_mm100: int) -> float:
    """Largest comfortable font size at which ``rows`` rows still fit in the given height."""
    for size in (16.0, 14.0, 12.0, 11.0, 10.0, 9.0):
        if rows * h.mm(row_min_mm(size)) <= height_mm100:
            return size
    return 8.0


def style_cell(
    cell: Any,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    fill: str | None = None,
    font_size: float | None = None,
    alignment: str | None = None,
    vertical: str | None = None,
) -> None:
    values: dict[str, Any] = {}
    if bold is not None:
        values["CharWeight"] = 150.0 if bold else 100.0
    if italic is not None:
        values["CharPosture"] = enum("com.sun.star.awt.FontSlant", "ITALIC" if italic else "NONE")
    if color is not None:
        values["CharColor"] = parse_color(color)
    if font_size is not None:
        values["CharHeight"] = float(font_size)
    if fill is not None:
        values["FillStyle"] = enum("com.sun.star.drawing.FillStyle", "SOLID")
        values["FillColor"] = parse_color(fill)
    if alignment is not None:
        if alignment not in h.ALIGN_TO_UNO:
            raise OpError(f"alignment must be one of: {', '.join(h.ALIGN_TO_UNO)}.")
        values["ParaAdjust"] = enum("com.sun.star.style.ParagraphAdjust", h.ALIGN_TO_UNO[alignment])
    if vertical is not None:
        if vertical not in ("top", "middle", "bottom"):
            raise OpError("vertical must be 'top', 'middle' or 'bottom'.")
        values["TextVerticalAdjust"] = enum(
            "com.sun.star.drawing.TextVerticalAdjust",
            {"top": "TOP", "middle": "CENTER", "bottom": "BOTTOM"}[vertical],
        )
    h.set_props(cell, values)


def create_table(
    s: Session,
    slide: Any,
    data: list[list[Any]],
    area: tuple[int, int, int, int],
    *,
    header: bool = True,
    column_widths: list[float] | None = None,
    font_size: float | None = None,
    header_color: str = HEADER_FILL,
    banded: bool = True,
) -> Any:
    """Insert a styled table filling ``area`` (x, y, width, height in 1/100 mm)."""
    if not data or not isinstance(data, list) or not all(isinstance(r, list) for r in data):
        raise OpError("data must be a non-empty list of rows, each a list of cell texts.")
    rows = len(data)
    cols = max(len(r) for r in data)
    if cols == 0:
        raise OpError("data rows must not be empty.")
    if rows * cols > _MAX_CELLS:
        raise OpError(f"Tables are limited to {_MAX_CELLS} cells; split the data over slides.")
    if column_widths is not None and len(column_widths) != cols:
        raise OpError(
            f"column_widths has {len(column_widths)} entries but the table has {cols} columns."
        )
    x, y, width, height = area
    shape = h.create_shape(s, TABLE_SERVICE, slide)
    shape.Size = Size(width, height)
    shape.Position = Point(x, y)
    model = shape.Model
    # A new table shape holds a single cell; grow it to the requested grid.
    if cols > 1:
        model.Columns.insertByIndex(0, cols - 1)
    if rows > 1:
        model.Rows.insertByIndex(0, rows - 1)
    weights = column_widths or [1.0] * cols
    total = float(sum(weights))
    for c in range(cols):
        model.Columns.getByIndex(c).Width = round(width * weights[c] / total)
    size = font_size or _auto_font_size(rows, height)
    row_height = max(round(height / rows), h.mm(row_min_mm(size)))
    border = _border(parse_color(BORDER_COLOR))
    text_color = "#222222"
    for r in range(rows):
        model.Rows.getByIndex(r).Height = row_height
        for c in range(cols):
            cell = model.getCellByPosition(c, r)
            value = data[r][c] if c < len(data[r]) else ""
            cell.setString("" if value is None else str(value))
            is_header = header and r == 0
            fill = (
                header_color
                if is_header
                else (BAND_FILL if banded and r % 2 == (1 if header else 0) else "#FFFFFF")
            )
            style_cell(
                cell, bold=is_header, color="#FFFFFF" if is_header else text_color,
                fill=fill, font_size=size, vertical="middle",
            )  # fmt: skip
            h.set_props(
                cell,
                {"TopBorder": border, "BottomBorder": border, "LeftBorder": border,
                 "RightBorder": border},
            )  # fmt: skip
    return shape


@op("impress.add_table", kind="impress", mutates=True)
def add_table(
    s: Session,
    slide_index: int,
    data: list[list[Any]],
    x_mm: float | None = None,
    y_mm: float | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    header: bool = True,
    column_widths: list[float] | None = None,
    font_size: float | None = None,
    header_color: str | None = None,
    banded: bool = True,
) -> dict[str, Any]:
    """Add a table. Unspecified geometry fills the slide's content area."""
    slide = h.slide_at(s.doc, slide_index)
    ax, ay, aw, ah = h.content_area(s.doc, slide)
    # Default height: as tall as the rows need, not the whole content area.
    # Without an explicit height the table is only as tall as its rows need (and never
    # taller than the content area), instead of stretching a 3-row table over the slide.
    size = font_size or _auto_font_size(len(data), ah)
    default_h = min(ah, h.mm(row_min_mm(size)) * len(data))
    area = (
        h.mm(x_mm) if x_mm is not None else ax,
        h.mm(y_mm) if y_mm is not None else ay,
        h.mm(width_mm) if width_mm is not None else aw,
        h.mm(height_mm) if height_mm is not None else default_h,
    )
    shape = create_table(
        s, slide, data, area, header=header, column_widths=column_widths, font_size=font_size,
        header_color=header_color or HEADER_FILL, banded=banded,
    )  # fmt: skip
    return {
        "slide_index": slide_index,
        "shape_index": slide.getCount() - 1,
        "rows": shape.Model.Rows.Count,
        "columns": shape.Model.Columns.Count,
        **h.geometry_mm(shape),
    }


@op("impress.set_table_cells", kind="impress", mutates=True)
def set_table_cells(
    s: Session, slide_index: int, shape_index: int, cells: list[dict[str, Any]]
) -> dict[str, Any]:
    """Set text and/or styling of individual cells ({row, column, text, bold, ...})."""
    slide = h.slide_at(s.doc, slide_index)
    shape = table_shape_at(slide, shape_index, slide_index)
    model = shape.Model
    rows, cols = model.Rows.Count, model.Columns.Count
    for n, spec in enumerate(cells):
        where = f"cells[{n}]"
        try:
            row, col = int(spec["row"]), int(spec["column"])
        except (KeyError, TypeError, ValueError) as exc:
            raise OpError(f"{where}: every cell needs integer 'row' and 'column'.") from exc
        if not (0 <= row < rows and 0 <= col < cols):
            raise OpError(
                f"{where}: row {row}, column {col} is outside the table ({rows} rows x {cols} "
                "columns, 0-based)."
            )
        cell = model.getCellByPosition(col, row)
        if "text" in spec and spec["text"] is not None:
            cell.setString(str(spec["text"]))
        try:
            style_cell(
                cell, bold=spec.get("bold"), italic=spec.get("italic"), color=spec.get("color"),
                fill=spec.get("fill"), font_size=spec.get("font_size"),
                alignment=spec.get("alignment"),
            )  # fmt: skip
        except OpError as exc:
            raise OpError(f"{where}: {exc.message}", exc.code) from exc
    return {"slide_index": slide_index, "shape_index": shape_index, "cells_updated": len(cells)}


@op("impress.modify_table", kind="impress", mutates=True)
def modify_table(
    s: Session,
    slide_index: int,
    shape_index: int,
    action: str,
    index: int,
    count: int = 1,
) -> dict[str, Any]:
    """Insert or delete rows / columns of a table."""
    slide = h.slide_at(s.doc, slide_index)
    shape = table_shape_at(slide, shape_index, slide_index)
    if action not in ("insert_rows", "insert_columns", "delete_rows", "delete_columns"):
        raise OpError("action must be insert_rows, insert_columns, delete_rows or delete_columns.")
    if count < 1:
        raise OpError("count must be at least 1.")
    model = shape.Model
    axis = model.Rows if action.endswith("rows") else model.Columns
    total = axis.Count
    what = "row" if action.endswith("rows") else "column"
    if action.startswith("insert"):
        if not 0 <= index <= total:
            raise OpError(f"index {index} is out of range: insert at 0..{total} ({total} = end).")
        axis.insertByIndex(index, count)
    else:
        if not 0 <= index < total or index + count > total:
            raise OpError(
                f"Cannot delete {count} {what}(s) from index {index}: the table has {total} "
                f"{what}s (0-based)."
            )
        if count >= total:
            raise OpError("A table needs at least one row and one column.")
        axis.removeByIndex(index, count)
    return {
        "slide_index": slide_index,
        "shape_index": shape_index,
        "rows": model.Rows.Count,
        "columns": model.Columns.Count,
    }
