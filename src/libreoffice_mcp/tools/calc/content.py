"""Calc tools that write or rearrange cell content."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.calc.types import (
    CellA1,
    CellValue,
    Grid,
    ParseText,
    RangeA1,
    SheetName,
    WholeRangeA1,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath

C = "calc"
Count = Annotated[int, Field(ge=1, le=100_000)]


@tool(C, "edit", title="Set range values")
async def calc_set_range_values(
    path: DocPath,
    data: Grid,
    sheet: SheetName = None,
    start_cell: Annotated[
        CellA1, Field(description="Top-left cell where the first value goes, e.g. 'B2'.")
    ] = "A1",
    parse_text: ParseText = False,
) -> dict[str, Any]:
    """Write a block of values starting at start_cell, overwriting what is there (null clears a
    cell). Strings starting with '=' are formulas; they may use Excel syntax (commas between
    arguments, Sheet2!A1 references) or LibreOffice syntax. The result lists `formula_errors`
    for formulas that evaluate to errors, so check it. To build a whole table with a styled
    header in one go, use calc_write_table instead."""
    return await edit(
        path, "calc.set_range_values", data=data, sheet=sheet, start_cell=start_cell,
        parse_text=parse_text,
    )  # fmt: skip


@tool(C, "edit", title="Set one cell")
async def calc_set_cell(
    path: DocPath,
    cell: CellA1,
    value: Annotated[
        CellValue,
        Field(description="Number, text, boolean, null (clear) or a formula such as '=A1*B1'."),
    ],
    sheet: SheetName = None,
    parse_text: ParseText = False,
) -> dict[str, Any]:
    """Set the content of a single cell (value or formula). Returns the displayed result, which
    shows immediately whether a formula worked."""
    return await edit(path, "calc.set_cell", cell=cell, value=value, sheet=sheet,
                      parse_text=parse_text)  # fmt: skip


@tool(C, "edit", title="Write table")
async def calc_write_table(
    path: DocPath,
    headers: Annotated[
        list[str], Field(min_length=1, description="Column headings, one per column.")
    ],
    rows: Annotated[
        list[list[CellValue]],
        Field(
            description="Data rows. Numbers stay numbers; strings starting with '=' are formulas."
        ),
    ],
    sheet: SheetName = None,
    start_cell: Annotated[
        CellA1, Field(description="Where the header row starts, e.g. 'A1'.")
    ] = "A1",
    header_style: Annotated[
        bool, Field(description="Make the header row bold with a light background and a border.")
    ] = True,
    autofit: Annotated[bool, Field(description="Fit column widths to the content.")] = True,
    column_formats: Annotated[
        dict[str, str] | None,
        Field(
            description="Number formats per column, keyed by header text or column letter, e.g. "
            "{'Revenue': '#,##0.00', 'Growth': '0.0%', 'Date': 'yyyy-mm-dd'}."
        ),
    ] = None,
    autofilter: Annotated[bool, Field(description="Add filter drop-downs to the header.")] = False,
    freeze_header: Annotated[
        bool, Field(description="Freeze everything down to the header row so it stays visible.")
    ] = False,
    parse_text: ParseText = False,
) -> dict[str, Any]:
    """Create a complete table in ONE call: a header row plus data rows, with optional header
    styling, number formats, column auto-width, filter buttons and a frozen header. This is the
    most efficient way to build a spreadsheet; follow with calc_add_chart for a chart. (In a
    .csv file only the cell text is stored, so styling and filters have no effect.)"""
    return await edit(
        path, "calc.write_table", headers=headers, rows=rows, sheet=sheet, start_cell=start_cell,
        header_style=header_style, autofit=autofit, column_formats=column_formats,
        autofilter=autofilter, freeze_header=freeze_header, parse_text=parse_text,
    )  # fmt: skip


@tool(C, "destructive", title="Clear range")
async def calc_clear_range(
    path: DocPath,
    range: WholeRangeA1,
    sheet: SheetName = None,
    what: Annotated[
        Literal["contents", "formats", "all"],
        Field(
            description="'contents' removes values and formulas but keeps formatting (default); "
            "'formats' removes formatting but keeps values; 'all' removes both and comments. "
            "Charts and images are never touched."
        ),
    ] = "contents",
) -> dict[str, Any]:
    """Clear the contents and/or formatting of a range. The cells themselves stay (use
    calc_delete_rows / calc_delete_columns to remove rows or columns)."""
    return await edit(path, "calc.clear_range", range=range, sheet=sheet, what=what)


@tool(C, "edit", title="Fill range")
async def calc_fill_range(
    path: DocPath,
    range: Annotated[
        RangeA1,
        Field(
            description="The source cell(s) AND the cells to fill, e.g. 'D2:D50' fills D3:D50 "
            "from D2 (direction 'down')."
        ),
    ],
    sheet: SheetName = None,
    direction: Annotated[Literal["down", "right", "up", "left"], Field()] = "down",
    mode: Annotated[
        Literal["autofill", "copy", "linear", "growth"],
        Field(
            description="'autofill' continues a series like dragging the fill handle (1,2,3..; "
            "Mon,Tue..; adjusts formulas); 'copy' repeats the source cell as is (formulas "
            "adjust their relative references); 'linear' / 'growth' build a number series "
            "from the first cell using `step`."
        ),
    ] = "autofill",
    step: Annotated[
        float, Field(description="Increment (linear) or factor (growth) for those modes.")
    ] = 1.0,
) -> dict[str, Any]:
    """Fill a range from its first cell (or last, for up/left): copy a formula down a column,
    extend a series, or repeat a value."""
    return await edit(
        path, "calc.fill_range", range=range, sheet=sheet, direction=direction, mode=mode, step=step
    )


@tool(C, "edit", title="Copy or move range")
async def calc_copy_range(
    path: DocPath,
    source_range: RangeA1,
    target_cell: Annotated[
        CellA1,
        Field(description="Top-left cell of the destination (existing content is overwritten)."),
    ],
    sheet: SheetName = None,
    target_sheet: Annotated[
        str | None, Field(description="Destination sheet. Default: the same sheet.")
    ] = None,
    move: Annotated[
        bool, Field(description="Move instead of copy: the source is emptied.")
    ] = False,
) -> dict[str, Any]:
    """Copy (or move) a block of cells with their values, formulas and formatting. Relative
    references in copied formulas adjust to the new position."""
    return await edit(
        path, "calc.copy_range", source_range=source_range, target_cell=target_cell, sheet=sheet,
        target_sheet=target_sheet, move=move,
    )  # fmt: skip


@tool(C, "destructive", title="Replace text in cells")
async def calc_replace_text(
    path: DocPath,
    search: str,
    replace: Annotated[str, Field(description="Replacement text ('' deletes the match).")],
    sheet: Annotated[str | None, Field(description="Only this sheet. Default: all sheets.")] = None,
    range: Annotated[
        WholeRangeA1 | None, Field(description="Only this range (requires `sheet`).")
    ] = None,
    regex: Annotated[
        bool, Field(description="Regular expression; $1 in `replace` refers to group 1.")
    ] = False,
    match_case: bool = False,
    whole_cell: Annotated[
        bool, Field(description="Only replace cells whose entire content equals `search`.")
    ] = False,
    search_in: Annotated[
        Literal["values", "formulas"],
        Field(description="'values' matches displayed text; 'formulas' matches raw formula text."),
    ] = "values",
) -> dict[str, Any]:
    """Find and replace text in cells. Returns the number of cells changed (nothing is saved
    when no cell matched)."""
    return await edit(
        path, "calc.replace_text", search=search, replace=replace, sheet=sheet, range=range,
        regex=regex, match_case=match_case, whole_cell=whole_cell, search_in=search_in,
    )  # fmt: skip


class SortKey(BaseModel):
    column: str | int = Field(
        description="Column to sort by: a letter like 'C', or a header text if has_header is true."
    )
    ascending: bool = Field(default=True, description="false = descending.")


@tool(C, "edit", title="Sort range")
async def calc_sort_range(
    path: DocPath,
    range: Annotated[
        RangeA1, Field(description="The whole table to sort, including the header row if any.")
    ],
    sort_by: Annotated[
        list[SortKey],
        Field(min_length=1, max_length=3, description="Sort keys, most important first."),
    ],
    sheet: SheetName = None,
    has_header: Annotated[
        bool, Field(description="The first row is a header and stays on top.")
    ] = True,
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Sort the rows of a range by up to three columns. Whole rows of the range move together.
    Formulas that refer to cells outside the range may change meaning when rows move."""
    return await edit(
        path, "calc.sort_range", range=range, sort_by=[k.model_dump() for k in sort_by],
        sheet=sheet, has_header=has_header, case_sensitive=case_sensitive,
    )  # fmt: skip


@tool(C, "edit", title="Insert rows")
async def calc_insert_rows(
    path: DocPath,
    before_row: Annotated[
        int, Field(ge=1, description="1-based row number; new rows are inserted above it.")
    ],
    count: Count = 1,
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Insert empty rows. Existing rows move down and formulas/ranges that refer to them adjust."""
    return await edit(path, "calc.insert_rows", before_row=before_row, count=count, sheet=sheet)


@tool(C, "destructive", title="Delete rows")
async def calc_delete_rows(
    path: DocPath,
    start_row: Annotated[int, Field(ge=1, description="First row to delete (1-based).")],
    count: Count = 1,
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Delete whole rows including their content. Rows below move up; formulas that referred to
    deleted cells become #REF!."""
    return await edit(path, "calc.delete_rows", start_row=start_row, count=count, sheet=sheet)


@tool(C, "edit", title="Insert columns")
async def calc_insert_columns(
    path: DocPath,
    before_column: Annotated[
        str, Field(description="Column letter, e.g. 'C'; new columns are inserted to its left.")
    ],
    count: Count = 1,
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Insert empty columns. Existing columns move right and formulas that refer to them adjust."""
    return await edit(
        path, "calc.insert_columns", before_column=before_column, count=count, sheet=sheet
    )


@tool(C, "destructive", title="Delete columns")
async def calc_delete_columns(
    path: DocPath,
    start_column: Annotated[str, Field(description="First column to delete, e.g. 'C'.")],
    count: Count = 1,
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Delete whole columns including their content. Columns to the right move left."""
    return await edit(
        path, "calc.delete_columns", start_column=start_column, count=count, sheet=sheet
    )


@tool(C, "edit", title="Merge cells")
async def calc_merge_cells(
    path: DocPath, range: RangeA1, sheet: SheetName = None
) -> dict[str, Any]:
    """Merge a range into one cell (e.g. a title across 'A1:E1'). Only the top-left cell's
    content survives; the result warns if other cells had content."""
    return await edit(path, "calc.merge_cells", range=range, sheet=sheet)


@tool(C, "edit", title="Unmerge cells")
async def calc_unmerge_cells(
    path: DocPath, range: RangeA1, sheet: SheetName = None
) -> dict[str, Any]:
    """Split merged cells in a range back into individual cells."""
    return await edit(path, "calc.unmerge_cells", range=range, sheet=sheet)
