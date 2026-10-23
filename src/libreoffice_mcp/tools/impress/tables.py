"""Impress table tools."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.types import (
    FontSize,
    I,
    OptHeight,
    OptWidth,
    OptX,
    OptY,
    ShapeIndex,
    SlideIndex,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Alignment, Color, DocPath


class CellEdit(BaseModel):
    """Change to one table cell; omitted fields stay as they are."""

    model_config = ConfigDict(extra="forbid")

    row: int = Field(ge=0, description="0-based row (0 = header row).")
    column: int = Field(ge=0, description="0-based column.")
    text: str | int | float | None = None
    bold: bool | None = None
    italic: bool | None = None
    color: Color | None = Field(None, description="Text colour.")
    fill: Color | None = Field(None, description="Cell background colour.")
    font_size: FontSize = None
    alignment: Alignment | None = None


@tool(I, "edit", title="Add table")
async def impress_add_table(
    path: DocPath,
    slide_index: SlideIndex,
    data: Annotated[
        list[list[str | int | float]],
        Field(min_length=1, description="Rows of cell values; the first row is the header."),
    ],
    x_mm: OptX = None,
    y_mm: OptY = None,
    width_mm: OptWidth = None,
    height_mm: OptHeight = None,
    header: Annotated[bool, Field(description="Style the first row as a header.")] = True,
    column_widths: Annotated[
        list[float] | None, Field(description="Relative column widths, e.g. [3, 1, 1].")
    ] = None,
    font_size: FontSize = None,
    header_color: Annotated[Color | None, Field(description="Header row fill colour.")] = None,
    banded: Annotated[bool, Field(description="Alternate row shading.")] = True,
) -> dict[str, Any]:
    """Add a styled table (header row, banded rows, thin borders). Without a position and size
    it spans the slide's content area below the title. Edit cells later with
    impress_set_table_cells. Returns the shape_index."""
    return await edit(
        path, "impress.add_table", slide_index=slide_index, data=data, x_mm=x_mm, y_mm=y_mm,
        width_mm=width_mm, height_mm=height_mm, header=header, column_widths=column_widths,
        font_size=font_size, header_color=header_color, banded=banded,
    )  # fmt: skip


@tool(I, "edit", title="Edit table cells")
async def impress_set_table_cells(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    cells: Annotated[list[CellEdit], Field(min_length=1, max_length=500)],
) -> dict[str, Any]:
    """Change the text and/or styling (bold, italic, text colour, fill, font size, alignment)
    of individual table cells. Rows and columns are 0-based; find the table's shape_index with
    impress_get_slide and read its cells with impress_get_shape."""
    return await edit(
        path, "impress.set_table_cells", slide_index=slide_index, shape_index=shape_index,
        cells=[c.model_dump(exclude_none=True) for c in cells],
    )  # fmt: skip


@tool(I, "destructive", title="Insert or delete table rows/columns")
async def impress_modify_table(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    action: Literal["insert_rows", "insert_columns", "delete_rows", "delete_columns"],
    index: Annotated[
        int,
        Field(ge=0, description="Insert before this row/column (n = append), or first to delete."),
    ],
    count: Annotated[int, Field(ge=1, le=100)] = 1,
) -> dict[str, Any]:
    """Insert or delete rows / columns of an existing table. New cells are empty; fill them
    with impress_set_table_cells."""
    return await edit(
        path, "impress.modify_table", slide_index=slide_index, shape_index=shape_index,
        action=action, index=index, count=count,
    )  # fmt: skip
