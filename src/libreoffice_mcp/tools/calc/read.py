"""Calc tools that only read (never modify the file)."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import read
from libreoffice_mcp.tools.calc.types import CellA1, SheetName, WholeRangeA1
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath

C = "calc"

OptionalSheet = Annotated[str | None, Field(description="Only this sheet. Default: all sheets.")]


@tool(C, "read", title="List sheets")
async def calc_list_sheets(path: DocPath) -> dict[str, Any]:
    """List every sheet of a spreadsheet: name, position (0-based index), used range (e.g.
    'A1:F120'), size, visibility and tab colour. Call this first when working with an existing
    file: it tells you the exact sheet names and how much data each sheet holds."""
    return await read(path, "calc.list_sheets")


@tool(C, "read", title="Get range values")
async def calc_get_range(
    path: DocPath,
    sheet: SheetName = None,
    range: Annotated[
        WholeRangeA1 | None,
        Field(description="Range to read. Default: the sheet's whole used range."),
    ] = None,
    include_formulas: Annotated[
        bool, Field(description="Also return the formulas of formula cells, keyed by cell.")
    ] = False,
    formatted: Annotated[
        bool,
        Field(
            description="Also return the text as displayed (dates, currencies, percentages as "
            "formatted). Limited to 2000 cells per call."
        ),
    ] = False,
    max_cells: Annotated[
        int, Field(ge=1, le=10_000, description="Cap on cells returned per call (rows are paged).")
    ] = 5000,
) -> dict[str, Any]:
    """Read cell values as a 2-D list (rows of cells). Numbers are returned as numbers, text as
    strings, empty cells as ''. Dates are stored as day numbers (e.g. 45306); pass
    formatted=true to see them as displayed. Large ranges are paged by rows: if `truncated` is
    true, call again with range=next_range. Formulas (with include_formulas) use LibreOffice
    syntax: ';' between arguments and 'Sheet2.A1' references."""
    return await read(
        path, "calc.get_range", sheet=sheet, range=range, include_formulas=include_formulas,
        formatted=formatted, max_cells=max_cells,
    )  # fmt: skip


@tool(C, "read", title="Get cell info")
async def calc_get_cell_info(
    path: DocPath, cell: CellA1, sheet: SheetName = None
) -> dict[str, Any]:
    """Detailed description of one cell: value, formula, displayed text, number format, font,
    background, alignment, merged range, hyperlink, comment and data validation. Use it to
    inspect formatting before changing it."""
    return await read(path, "calc.get_cell_info", sheet=sheet, cell=cell)


@tool(C, "read", title="Find text in sheets")
async def calc_find_text(
    path: DocPath,
    query: str,
    sheet: Annotated[
        str | None, Field(description="Limit the search to this sheet. Default: all sheets.")
    ] = None,
    regex: Annotated[bool, Field(description="Treat query as a regular expression.")] = False,
    match_case: bool = False,
    whole_cell: Annotated[
        bool, Field(description="Only match cells whose entire content equals the query.")
    ] = False,
    search_in: Annotated[
        Literal["values", "formulas"],
        Field(description="Search the cell values (default) or the formula text."),
    ] = "values",
    max_results: Annotated[int, Field(ge=1, le=500)] = 100,
) -> dict[str, Any]:
    """Find cells containing text or numbers. Returns each hit's sheet, cell address and value
    (plus the formula for formula cells when search_in='formulas')."""
    return await read(
        path, "calc.find_text", query=query, sheet=sheet, regex=regex, match_case=match_case,
        whole_cell=whole_cell, search_in=search_in, max_results=max_results,
    )  # fmt: skip


@tool(C, "read", title="Get sheet layout")
async def calc_get_sheet_layout(path: DocPath, sheet: SheetName = None) -> dict[str, Any]:
    """Describe a sheet's layout: used range, frozen rows/columns, auto-filter range, column
    widths (mm), hidden rows/columns, custom row heights and merged ranges."""
    return await read(path, "calc.get_sheet_layout", sheet=sheet)


@tool(C, "read", title="List named ranges")
async def calc_list_named_ranges(path: DocPath) -> dict[str, Any]:
    """List the named ranges (defined names) and what they refer to."""
    return await read(path, "calc.list_named_ranges")


@tool(C, "read", title="List conditional formats")
async def calc_list_conditional_formats(path: DocPath, sheet: SheetName = None) -> dict[str, Any]:
    """List the conditional formatting rules of a sheet: the cells they cover and each rule
    (comparison, formula or data bar)."""
    return await read(path, "calc.list_conditional_formats", sheet=sheet)


@tool(C, "read", title="List charts")
async def calc_list_charts(path: DocPath, sheet: OptionalSheet = None) -> dict[str, Any]:
    """List charts: sheet, name (needed by calc_update_chart / calc_delete_chart), type, title,
    data range, position and size."""
    return await read(path, "calc.list_charts", sheet=sheet)


@tool(C, "read", title="List images")
async def calc_list_images(path: DocPath, sheet: OptionalSheet = None) -> dict[str, Any]:
    """List pictures placed on sheets: name (needed by calc_delete_image), anchor cell,
    position and size in mm."""
    return await read(path, "calc.list_images", sheet=sheet)


@tool(C, "read", title="List comments")
async def calc_list_comments(path: DocPath, sheet: OptionalSheet = None) -> dict[str, Any]:
    """List cell comments (notes): sheet, cell, text and author."""
    return await read(path, "calc.list_comments", sheet=sheet)
