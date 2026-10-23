"""Writer: tables. Cells are addressed like a spreadsheet: ``A1`` is column A, row 1."""

from __future__ import annotations

from typing import Any

from com.sun.star.table import BorderLine2

from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.ops.writer.helpers import cell_name, parse_cell
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import parse_color

HEADER_BACKGROUND = "#D9D9D9"
_PT_TO_MM100 = 35.28


def parse_range(text: str) -> tuple[int, int, int, int]:
    """'A1:C2' or 'B2' -> (col1, row1, col2, row2)."""
    first, _, last = text.partition(":")
    c1, r1 = parse_cell(first)
    c2, r2 = parse_cell(last) if last else (c1, r1)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def _check_in_table(table: Any, col: int, row: int) -> None:
    rows, cols = table.getRows().getCount(), table.getColumns().getCount()
    if not (0 <= col < cols and 0 <= row < rows):
        raise OpError(
            f"Cell {cell_name(col, row)} is outside the table ({rows} rows x {cols} columns)."
        )


def _set_widths(table: Any, widths: list[float]) -> None:
    columns = table.getColumns().getCount()
    if len(widths) != columns:
        raise OpError(f"column_widths needs {columns} numbers (one per column), got {len(widths)}.")
    if any(w <= 0 for w in widths):
        raise OpError("column_widths must be positive numbers.")
    separators = list(table.TableColumnSeparators)
    total = table.TableColumnRelativeSum
    running = 0.0
    for i, width in enumerate(widths[:-1]):
        running += width / sum(widths)
        separators[i].Position = int(running * total)
    table.TableColumnSeparators = tuple(separators)


def _style_header_row(table: Any) -> None:
    table.HeaderRowCount = 1  # repeat the header on each page
    for col in range(table.getColumns().getCount()):
        cell = table.getCellByPosition(col, 0)
        cell.BackColor = parse_color(HEADER_BACKGROUND)
        cursor = cell.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        cursor.CharWeight = 150.0


def build_table(
    slot: h.Slot,
    data: list[list[Any]],
    header: bool = True,
    column_widths: list[float] | None = None,
) -> Any:
    """Insert a table at ``slot`` sized to ``data`` (a list of rows) and fill it."""
    if not data or not all(isinstance(row, list) for row in data):
        raise OpError("Table data must be a non-empty list of rows, each a list of cell values.")
    rows, columns = len(data), max(len(row) for row in data)
    if columns == 0:
        raise OpError("Table rows are empty.")
    table = slot.table(rows, columns)
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            if value is not None and value != "":
                table.getCellByPosition(c, r).setString(str(value))
    if header and rows > 1:
        _style_header_row(table)
    if column_widths:
        _set_widths(table, column_widths)
    return table


@op("writer.insert_table", kind="writer", mutates=True)
def insert_table(
    s: Session,
    data: list[list[Any]] | None = None,
    rows: int | None = None,
    columns: int | None = None,
    header: bool = True,
    column_widths: list[float] | None = None,
    position: str = "end",
    block_index: int | None = None,
) -> dict[str, Any]:
    if data is None:
        if not rows or not columns:
            raise OpError("Give either 'data' (a list of rows) or both 'rows' and 'columns'.")
        data = [["" for _ in range(columns)] for _ in range(rows)]
    slot = h.Slot(s.doc, position, block_index)
    table = build_table(slot, data, header, column_widths)
    return {
        "table_index": len([b for b in h.all_blocks(s.doc)[: slot.after + 1] if h.is_table(b)]) - 1,
        "block_index": slot.after,
        "rows": table.getRows().getCount(),
        "columns": table.getColumns().getCount(),
    }


@op("writer.set_table_cells", kind="writer", mutates=True)
def set_table_cells(
    s: Session, table_index: int, data: list[list[Any]], start_cell: str = "A1"
) -> dict[str, Any]:
    """Write a block of values starting at ``start_cell`` (overwrites existing text)."""
    table = h.get_table(s.doc, table_index)
    col0, row0 = parse_cell(start_cell)
    written = 0
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            _check_in_table(table, col0 + c, row0 + r)
            table.getCellByPosition(col0 + c, row0 + r).setString(
                "" if value is None else str(value)
            )
            written += 1
    return {"cells_written": written}


@op("writer.add_table_rows", kind="writer", mutates=True)
def add_table_rows(
    s: Session, table_index: int, count: int = 1, after_row: int | None = None
) -> dict[str, Any]:
    """Insert empty rows after 0-based row ``after_row`` (default: at the end)."""
    table = h.get_table(s.doc, table_index)
    rows = table.getRows()
    position = rows.getCount() if after_row is None else after_row + 1
    if not 0 <= position <= rows.getCount() or count < 1:
        raise OpError(f"after_row must be within 0..{rows.getCount() - 1} and count >= 1.")
    rows.insertByIndex(position, count)
    return {"rows": rows.getCount()}


@op("writer.delete_table_rows", kind="writer", mutates=True)
def delete_table_rows(
    s: Session, table_index: int, start_row: int, count: int = 1
) -> dict[str, Any]:
    table = h.get_table(s.doc, table_index)
    rows = table.getRows()
    if start_row < 0 or count < 1 or start_row + count > rows.getCount():
        raise OpError(
            f"Rows {start_row}..{start_row + count - 1} are outside the table ({rows.getCount()} rows)."
        )
    if count == rows.getCount():
        raise OpError("That would delete every row; use writer_delete_table instead.")
    rows.removeByIndex(start_row, count)
    return {"rows": rows.getCount()}


@op("writer.add_table_columns", kind="writer", mutates=True)
def add_table_columns(
    s: Session, table_index: int, count: int = 1, after_column: int | None = None
) -> dict[str, Any]:
    table = h.get_table(s.doc, table_index)
    columns = table.getColumns()
    position = columns.getCount() if after_column is None else after_column + 1
    if not 0 <= position <= columns.getCount() or count < 1:
        raise OpError(f"after_column must be within 0..{columns.getCount() - 1} and count >= 1.")
    columns.insertByIndex(position, count)
    return {"columns": columns.getCount()}


@op("writer.delete_table_columns", kind="writer", mutates=True)
def delete_table_columns(
    s: Session, table_index: int, start_column: int, count: int = 1
) -> dict[str, Any]:
    table = h.get_table(s.doc, table_index)
    columns = table.getColumns()
    if start_column < 0 or count < 1 or start_column + count > columns.getCount():
        raise OpError(f"Columns {start_column}..{start_column + count - 1} are outside the table.")
    if count == columns.getCount():
        raise OpError("That would delete every column; use writer_delete_table instead.")
    columns.removeByIndex(start_column, count)
    return {"columns": columns.getCount()}


@op("writer.merge_table_cells", kind="writer", mutates=True)
def merge_table_cells(s: Session, table_index: int, cell_range: str) -> dict[str, Any]:
    """Merge a rectangular range such as 'A1:C1' into one cell."""
    table = h.get_table(s.doc, table_index)
    col1, row1, col2, row2 = parse_range(cell_range)
    _check_in_table(table, col2, row2)
    cursor = table.createCursorByCellName(cell_name(col1, row1))
    if col2 > col1:
        cursor.goRight(col2 - col1, True)
    if row2 > row1:
        cursor.goDown(row2 - row1, True)
    cursor.mergeRange()
    return {"merged": cell_range}


@op("writer.delete_table", kind="writer", mutates=True)
def delete_table(s: Session, table_index: int) -> dict[str, Any]:
    table = h.get_table(s.doc, table_index)
    s.doc.getText().removeTextContent(table)
    return {"deleted_table": table_index}


@op("writer.format_table", kind="writer", mutates=True)
def format_table(
    s: Session,
    table_index: int,
    border_width_pt: float | None = None,
    border_color: str | None = None,
    background_color: str | None = None,
    header_row: bool | None = None,
    header_background: str | None = None,
    alignment: str | None = None,
    column_widths: list[float] | None = None,
) -> dict[str, Any]:
    """Borders, background, header row, table alignment and column widths."""
    table = h.get_table(s.doc, table_index)
    changed = []
    if border_width_pt is not None or border_color is not None:
        line = BorderLine2()
        width = 0 if border_width_pt == 0 else int((border_width_pt or 0.5) * _PT_TO_MM100)
        line.LineWidth = width
        line.OuterLineWidth = width
        line.LineStyle = 0 if width else 0x7FFF  # SOLID; width 0 means "no border"
        line.Color = parse_color(border_color or "#000000")
        border = table.TableBorder2
        for side in ("Top", "Bottom", "Left", "Right", "Horizontal", "Vertical"):
            setattr(border, f"{side}Line", line)
            setattr(border, f"Is{side}LineValid", True)
        table.TableBorder2 = border
        changed.append("borders")
    if background_color is not None:
        table.BackColor = parse_color(background_color)
        table.BackTransparent = False
        changed.append("background_color")
    if header_row is not None:
        if header_row:
            _style_header_row(table)
            if header_background:
                for col in range(table.getColumns().getCount()):
                    table.getCellByPosition(col, 0).BackColor = parse_color(header_background)
        else:
            table.HeaderRowCount = 0
        changed.append("header_row")
    if alignment is not None:
        orient = {"left": 3, "center": 2, "right": 1, "full": 6}.get(alignment)  # HoriOrientation
        if orient is None:
            raise OpError("alignment must be left, center, right or full.")
        table.HoriOrient = orient
        changed.append("alignment")
    if column_widths:
        _set_widths(table, column_widths)
        changed.append("column_widths")
    if not changed:
        raise OpError("No formatting requested.")
    return {"table_index": table_index, "changed": changed}


@op("writer.format_table_cells", kind="writer", mutates=True)
def format_table_cells(
    s: Session,
    table_index: int,
    cell_range: str,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    font: str | None = None,
    size: float | None = None,
    background_color: str | None = None,
    alignment: str | None = None,
    vertical_align: str | None = None,
) -> dict[str, Any]:
    """Format cells in a range like 'A1:C1' (text style, background, alignment)."""
    table = h.get_table(s.doc, table_index)
    col1, row1, col2, row2 = parse_range(cell_range)
    _check_in_table(table, col2, row2)
    vertical = {"top": 1, "center": 2, "bottom": 3}  # VertOrientation
    if vertical_align is not None and vertical_align not in vertical:
        raise OpError("vertical_align must be top, center or bottom.")
    count = 0
    for r in range(row1, row2 + 1):
        for c in range(col1, col2 + 1):
            cell = table.getCellByPosition(c, r)
            if background_color is not None:
                cell.BackColor = parse_color(background_color)
            if vertical_align is not None:
                cell.VertOrient = vertical[vertical_align]
            cursor = cell.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            h.apply_char_format(cursor, bold=bold, italic=italic, color=color, font=font, size=size)
            if alignment is not None:
                h.apply_para_format(cursor, alignment=alignment)
            count += 1
    return {"cells_formatted": count}
