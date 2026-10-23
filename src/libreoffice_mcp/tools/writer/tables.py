"""Writer tools for tables. Cells use spreadsheet notation: A1 = first column, first row."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import (
    Alignment,
    AnchorBlock,
    CellRange,
    Color,
    DocPath,
    Position,
    TableIndex,
)

W = "writer"
Count = Annotated[int, Field(ge=1, le=1000)]


@tool(W, "edit", title="Insert table")
async def writer_insert_table(
    path: DocPath,
    data: Annotated[
        list[list[str]] | None,
        Field(description="Rows of cell text (first row = header). Sets the table size."),
    ] = None,
    rows: Annotated[
        int | None, Field(ge=1, le=1000, description="Only if no data: empty table rows.")
    ] = None,
    columns: Annotated[
        int | None, Field(ge=1, le=100, description="Only if no data: empty table columns.")
    ] = None,
    header: Annotated[
        bool,
        Field(description="Style the first row as a header (bold, shaded, repeats on each page)."),
    ] = True,
    column_widths: Annotated[
        list[float] | None, Field(description="Relative widths, e.g. [1, 3, 1].")
    ] = None,
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Insert a table. Returns its table_index for later table_* calls."""
    return await edit(
        path, "writer.insert_table", data=data, rows=rows, columns=columns, header=header,
        column_widths=column_widths, position=position, block_index=block_index,
    )  # fmt: skip


@tool(W, "edit", title="Set table cells")
async def writer_set_table_cells(
    path: DocPath,
    table_index: TableIndex,
    data: Annotated[list[list[str]], Field(description="Rows of values to write.")],
    start_cell: Annotated[str, Field(description="Top-left cell to write to, e.g. 'B2'.")] = "A1",
) -> dict[str, Any]:
    """Write a block of text into a table starting at start_cell (overwrites existing text)."""
    return await edit(
        path, "writer.set_table_cells", table_index=table_index, data=data, start_cell=start_cell
    )


@tool(W, "edit", title="Add table rows")
async def writer_add_table_rows(
    path: DocPath,
    table_index: TableIndex,
    count: Count = 1,
    after_row: Annotated[
        int | None, Field(ge=0, description="0-based row to insert after; default = at the end.")
    ] = None,
) -> dict[str, Any]:
    """Insert empty rows into a table."""
    return await edit(
        path, "writer.add_table_rows", table_index=table_index, count=count, after_row=after_row
    )


@tool(W, "destructive", title="Delete table rows")
async def writer_delete_table_rows(
    path: DocPath, table_index: TableIndex, start_row: Annotated[int, Field(ge=0)], count: Count = 1
) -> dict[str, Any]:
    """Delete rows from a table (0-based start_row)."""
    return await edit(
        path, "writer.delete_table_rows", table_index=table_index, start_row=start_row, count=count
    )


@tool(W, "edit", title="Add table columns")
async def writer_add_table_columns(
    path: DocPath,
    table_index: TableIndex,
    count: Count = 1,
    after_column: Annotated[
        int | None, Field(ge=0, description="0-based column to insert after; default = at the end.")
    ] = None,
) -> dict[str, Any]:
    """Insert empty columns into a table."""
    return await edit(
        path,
        "writer.add_table_columns",
        table_index=table_index,
        count=count,
        after_column=after_column,
    )


@tool(W, "destructive", title="Delete table columns")
async def writer_delete_table_columns(
    path: DocPath,
    table_index: TableIndex,
    start_column: Annotated[int, Field(ge=0)],
    count: Count = 1,
) -> dict[str, Any]:
    """Delete columns from a table (0-based start_column)."""
    return await edit(
        path,
        "writer.delete_table_columns",
        table_index=table_index,
        start_column=start_column,
        count=count,
    )


@tool(W, "edit", title="Merge table cells")
async def writer_merge_table_cells(
    path: DocPath, table_index: TableIndex, cell_range: CellRange
) -> dict[str, Any]:
    """Merge a rectangular range (e.g. 'A1:C1' for a full-width title row) into one cell."""
    return await edit(
        path, "writer.merge_table_cells", table_index=table_index, cell_range=cell_range
    )


@tool(W, "destructive", title="Delete table")
async def writer_delete_table(path: DocPath, table_index: TableIndex) -> dict[str, Any]:
    """Delete an entire table."""
    return await edit(path, "writer.delete_table", table_index=table_index)


@tool(W, "edit", title="Format table")
async def writer_format_table(
    path: DocPath,
    table_index: TableIndex,
    border_width_pt: Annotated[
        float | None, Field(ge=0, le=20, description="0 removes borders.")
    ] = None,
    border_color: Color | None = None,
    background_color: Color | None = None,
    header_row: Annotated[
        bool | None, Field(description="True: bold + shaded first row that repeats on each page.")
    ] = None,
    header_background: Color | None = None,
    alignment: Literal["left", "center", "right", "full"] | None = None,
    column_widths: Annotated[
        list[float] | None, Field(description="Relative widths, one per column.")
    ] = None,
) -> dict[str, Any]:
    """Table-wide formatting: borders, background, header row, alignment, column widths."""
    return await edit(
        path, "writer.format_table", table_index=table_index, border_width_pt=border_width_pt,
        border_color=border_color, background_color=background_color, header_row=header_row,
        header_background=header_background, alignment=alignment, column_widths=column_widths,
    )  # fmt: skip


@tool(W, "edit", title="Format table cells")
async def writer_format_table_cells(
    path: DocPath,
    table_index: TableIndex,
    cell_range: CellRange,
    bold: bool | None = None,
    italic: bool | None = None,
    color: Color | None = None,
    font: str | None = None,
    size: Annotated[float | None, Field(gt=0, le=1000)] = None,
    background_color: Color | None = None,
    alignment: Alignment | None = None,
    vertical_align: Literal["top", "center", "bottom"] | None = None,
) -> dict[str, Any]:
    """Format the cells in a range: text style, shading, horizontal/vertical alignment."""
    return await edit(
        path, "writer.format_table_cells", table_index=table_index, cell_range=cell_range, bold=bold,
        italic=italic, color=color, font=font, size=size, background_color=background_color,
        alignment=alignment, vertical_align=vertical_align,
    )  # fmt: skip
