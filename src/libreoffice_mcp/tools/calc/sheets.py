"""Calc tools for sheets, named ranges, freeze panes, auto-filter and recalculation."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.calc.types import RangeA1, RequiredSheet, SheetName
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath

C = "calc"
Position = Annotated[
    int | None,
    Field(ge=0, description="0-based position among the sheet tabs (0 = first). Default: last."),
]


@tool(C, "edit", title="Add sheet")
async def calc_add_sheet(
    path: DocPath,
    name: Annotated[
        str, Field(max_length=31, description="New sheet name (unique, max 31 chars).")
    ],
    position: Position = None,
) -> dict[str, Any]:
    """Add an empty sheet with the given name, at a tab position (default: after the last one)."""
    return await edit(path, "calc.add_sheet", name=name, position=position)


@tool(C, "edit", title="Rename sheet")
async def calc_rename_sheet(path: DocPath, sheet: RequiredSheet, new_name: str) -> dict[str, Any]:
    """Rename a sheet. Formulas that refer to it are updated automatically."""
    return await edit(path, "calc.rename_sheet", sheet=sheet, new_name=new_name)


@tool(C, "destructive", title="Delete sheet")
async def calc_delete_sheet(path: DocPath, sheet: RequiredSheet) -> dict[str, Any]:
    """Delete a sheet and everything on it. Formulas in other sheets that referred to it
    become #REF!. A spreadsheet must keep at least one sheet."""
    return await edit(path, "calc.delete_sheet", sheet=sheet)


@tool(C, "edit", title="Move sheet")
async def calc_move_sheet(
    path: DocPath, sheet: RequiredSheet, position: Annotated[int, Field(ge=0)]
) -> dict[str, Any]:
    """Move a sheet to a new tab position (0 = first)."""
    return await edit(path, "calc.move_sheet", sheet=sheet, position=position)


@tool(C, "edit", title="Copy sheet")
async def calc_copy_sheet(
    path: DocPath, sheet: RequiredSheet, new_name: str, position: Position = None
) -> dict[str, Any]:
    """Duplicate a sheet (content, formatting, charts) under a new name."""
    return await edit(path, "calc.copy_sheet", sheet=sheet, new_name=new_name, position=position)


@tool(C, "edit", title="Set sheet properties")
async def calc_set_sheet_properties(
    path: DocPath,
    sheet: RequiredSheet,
    visible: Annotated[
        bool | None, Field(description="false hides the sheet tab, true shows it.")
    ] = None,
    tab_color: Annotated[
        str | None, Field(description="Tab colour '#RRGGBB' or a name; 'none' removes it.")
    ] = None,
) -> dict[str, Any]:
    """Show/hide a sheet and/or set its tab colour."""
    return await edit(
        path, "calc.set_sheet_properties", sheet=sheet, visible=visible, tab_color=tab_color
    )


@tool(C, "edit", title="Add named range")
async def calc_add_named_range(
    path: DocPath,
    name: Annotated[str, Field(description="Name, e.g. 'SalesData' (letters, digits, '_', '.').")],
    range: RangeA1,
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Give a range a name that formulas can use, e.g. =SUM(SalesData)."""
    return await edit(path, "calc.add_named_range", name=name, range=range, sheet=sheet)


@tool(C, "destructive", title="Delete named range")
async def calc_delete_named_range(path: DocPath, name: str) -> dict[str, Any]:
    """Delete a named range (the cells stay; formulas using the name will show #NAME?)."""
    return await edit(path, "calc.delete_named_range", name=name)


@tool(C, "edit", title="Freeze panes")
async def calc_freeze_panes(
    path: DocPath,
    sheet: SheetName = None,
    rows: Annotated[
        int, Field(ge=0, le=1000, description="Number of top rows to keep visible (0 = none).")
    ] = 0,
    columns: Annotated[
        int, Field(ge=0, le=100, description="Number of left columns to keep visible (0 = none).")
    ] = 0,
) -> dict[str, Any]:
    """Freeze the top rows and/or left columns so they stay visible while scrolling (typically
    rows=1 to keep a header). rows=0 and columns=0 removes the freeze. Only available for .xlsx
    and .xls files: LibreOffice cannot save freeze panes into .ods files without a window."""
    return await edit(path, "calc.freeze_panes", sheet=sheet, rows=rows, columns=columns)


@tool(C, "edit", title="Set auto-filter")
async def calc_set_autofilter(
    path: DocPath,
    range: Annotated[
        RangeA1 | None,
        Field(description="Table range including the header row. Default: the used range."),
    ] = None,
    sheet: SheetName = None,
    enabled: Annotated[bool, Field(description="false removes the filter buttons.")] = True,
) -> dict[str, Any]:
    """Show (or remove) filter drop-down buttons on a table's header row. One filter per sheet."""
    return await edit(path, "calc.set_autofilter", range=range, sheet=sheet, enabled=enabled)


@tool(C, "edit", title="Recalculate")
async def calc_recalculate(path: DocPath) -> dict[str, Any]:
    """Recalculate every formula and store the fresh results. Useful after editing a file that
    was produced by another program."""
    return await edit(path, "calc.recalculate")
