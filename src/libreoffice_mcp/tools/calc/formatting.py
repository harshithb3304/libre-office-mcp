"""Calc tools for formatting: cells, borders, sizes, conditional formats, data validation."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.calc.types import (
    CellValue,
    Color,
    HAlign,
    RangeA1,
    SheetName,
    VAlign,
    WholeRangeA1,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath

C = "calc"


@tool(C, "edit", title="Format range")
async def calc_format_range(
    path: DocPath,
    range: WholeRangeA1,
    sheet: SheetName = None,
    font_name: Annotated[str | None, Field(description="e.g. 'Arial', 'Calibri'.")] = None,
    font_size: Annotated[float | None, Field(ge=1, le=400, description="Points.")] = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_color: Color | None = None,
    background: Annotated[
        str | None,
        Field(description="Fill colour '#RRGGBB' or a name; 'none' removes the fill."),
    ] = None,
    number_format: Annotated[
        str | None,
        Field(
            description="Number format code: '0.00', '#,##0', '0%', '0.0%', 'yyyy-mm-dd', "
            "'dd/mm/yyyy hh:mm', '$#,##0.00', '[$EUR] #,##0.00', '@' (text), 'General'."
        ),
    ] = None,
    h_align: HAlign | None = None,
    v_align: VAlign | None = None,
    wrap_text: bool | None = None,
    rotation: Annotated[
        int | None, Field(ge=-360, le=360, description="Text rotation in degrees (90 = up).")
    ] = None,
    cell_style: Annotated[
        str | None,
        Field(description="Name of a cell style to apply first (e.g. 'Heading 1', 'Good', 'Bad')."),
    ] = None,
) -> dict[str, Any]:
    """Format cells: font, colours, number format, alignment, wrapping. Only the options you
    pass are changed. Accepts whole columns ('A:C') and rows ('1:1'). Use calc_set_borders for
    borders and calc_set_column_width for widths."""
    return await edit(
        path, "calc.format_range", range=range, sheet=sheet, font_name=font_name,
        font_size=font_size, bold=bold, italic=italic, underline=underline,
        strikethrough=strikethrough, font_color=font_color, background=background,
        number_format=number_format, h_align=h_align, v_align=v_align, wrap_text=wrap_text,
        rotation=rotation, cell_style=cell_style,
    )  # fmt: skip


@tool(C, "edit", title="Set borders")
async def calc_set_borders(
    path: DocPath,
    range: RangeA1,
    sheet: SheetName = None,
    sides: Annotated[
        list[
            Literal[
                "all",
                "outer",
                "inner",
                "top",
                "bottom",
                "left",
                "right",
                "horizontal",
                "vertical",
                "none",
            ]
        ]
        | None,
        Field(
            description="Which borders: 'all' (every cell edge, the default), 'outer' (around the "
            "range), 'inner', single sides, 'horizontal'/'vertical' inner lines, or 'none' to "
            "remove all."
        ),
    ] = None,
    style: Literal["solid", "dashed", "dotted", "double"] = "solid",
    width_pt: Annotated[float, Field(ge=0.1, le=10)] = 0.75,
    color: Color = "#000000",
) -> dict[str, Any]:
    """Draw (or remove) cell borders. Sides not listed keep their current border."""
    return await edit(
        path, "calc.set_borders", range=range, sheet=sheet, sides=sides, style=style,
        width_pt=width_pt, color=color,
    )  # fmt: skip


@tool(C, "edit", title="Set column width")
async def calc_set_column_width(
    path: DocPath,
    columns: Annotated[str, Field(description="Column letter 'B' or a span 'A:D'.")],
    sheet: SheetName = None,
    width_mm: Annotated[
        float | None, Field(ge=0.5, le=1000, description="Width in mm (default column is ~22.6).")
    ] = None,
    auto_fit: Annotated[bool, Field(description="Fit the width to the content.")] = False,
    hidden: Annotated[bool | None, Field(description="true hides, false shows.")] = None,
) -> dict[str, Any]:
    """Set the width of columns, fit them to their content, or hide/show them."""
    return await edit(
        path, "calc.set_column_width", columns=columns, sheet=sheet, width_mm=width_mm,
        auto_fit=auto_fit, hidden=hidden,
    )  # fmt: skip


@tool(C, "edit", title="Set row height")
async def calc_set_row_height(
    path: DocPath,
    rows: Annotated[str, Field(description="Row number '3' (1-based) or a span '2:10'.")],
    sheet: SheetName = None,
    height_mm: Annotated[
        float | None, Field(ge=0.5, le=400, description="Height in mm (default row is ~4.5).")
    ] = None,
    auto_fit: Annotated[bool, Field(description="Fit the height to the content.")] = False,
    hidden: Annotated[bool | None, Field(description="true hides, false shows.")] = None,
) -> dict[str, Any]:
    """Set the height of rows, fit them to their content (e.g. after wrapping text), or
    hide/show them."""
    return await edit(
        path, "calc.set_row_height", rows=rows, sheet=sheet, height_mm=height_mm,
        auto_fit=auto_fit, hidden=hidden,
    )  # fmt: skip


@tool(C, "edit", title="Add conditional format")
async def calc_add_conditional_format(
    path: DocPath,
    range: RangeA1,
    rule: Annotated[
        Literal[
            "greater",
            "greater_equal",
            "less",
            "less_equal",
            "equal",
            "not_equal",
            "between",
            "not_between",
            "formula",
            "duplicate",
            "unique",
        ],
        Field(
            description="Condition on each cell's value. 'formula' evaluates an expression written "
            "for the top-left cell of the range (relative references shift for other cells)."
        ),
    ],
    sheet: SheetName = None,
    value: Annotated[
        CellValue,
        Field(
            description="Number or text to compare with ('between': the lower bound). For "
            "'formula': the expression, e.g. '=$B2>$C2' or '=ISBLANK(A1)'. A text starting with "
            "'=' is a cell reference or expression, e.g. '=$F$1'."
        ),
    ] = None,
    value2: Annotated[CellValue, Field(description="Upper bound for 'between'.")] = None,
    background: Annotated[
        str | None, Field(description="Fill colour when the rule matches.")
    ] = None,
    font_color: Color | None = None,
    bold: bool = False,
) -> dict[str, Any]:
    """Highlight cells automatically when a condition holds, e.g. red fill for values < 0. Give
    at least one of background / font_color / bold as the look to apply. Rules accumulate; use
    calc_clear_conditional_formats to remove them."""
    return await edit(
        path, "calc.add_conditional_format", range=range, rule=rule, sheet=sheet, value=value,
        value2=value2, background=background, font_color=font_color, bold=bold,
    )  # fmt: skip


@tool(C, "edit", title="Add data bars")
async def calc_add_data_bar(
    path: DocPath,
    range: RangeA1,
    sheet: SheetName = None,
    color: Color = "#638EC6",
    show_value: Annotated[bool, Field(description="Show the number next to the bar.")] = True,
    gradient: Annotated[bool, Field(description="Gradient fill (false = solid).")] = True,
) -> dict[str, Any]:
    """Show an in-cell bar whose length reflects each number relative to the range's min/max."""
    return await edit(
        path,
        "calc.add_data_bar",
        range=range,
        sheet=sheet,
        color=color,
        show_value=show_value,
        gradient=gradient,
    )


@tool(C, "destructive", title="Clear conditional formats")
async def calc_clear_conditional_formats(
    path: DocPath,
    sheet: SheetName = None,
    range: Annotated[
        WholeRangeA1 | None,
        Field(description="Remove rules touching this range. Default: all rules on the sheet."),
    ] = None,
) -> dict[str, Any]:
    """Remove conditional formatting rules (a rule that covers several ranges is removed as a
    whole)."""
    return await edit(path, "calc.clear_conditional_formats", sheet=sheet, range=range)


@tool(C, "edit", title="Set data validation")
async def calc_set_data_validation(
    path: DocPath,
    range: WholeRangeA1,
    type: Annotated[
        Literal["list", "whole_number", "decimal", "date", "text_length", "custom", "any"],
        Field(description="Kind of rule; 'any' removes validation from the range."),
    ],
    sheet: SheetName = None,
    items: Annotated[
        list[str] | None,
        Field(description="For type 'list': the allowed choices, e.g. ['Open','Closed']."),
    ] = None,
    source_range: Annotated[
        RangeA1 | None,
        Field(description="For type 'list': take the choices from these cells instead of items."),
    ] = None,
    source_sheet: Annotated[
        str | None, Field(description="Sheet of source_range. Default: the same sheet.")
    ] = None,
    operator: Annotated[
        Literal[
            "between",
            "not_between",
            "equal",
            "not_equal",
            "greater",
            "greater_equal",
            "less",
            "less_equal",
        ],
        Field(description="For number/date/length rules."),
    ] = "between",
    min_value: Annotated[
        CellValue,
        Field(description="Lower bound, or the comparison value. Dates as 'YYYY-MM-DD'."),
    ] = None,
    max_value: Annotated[CellValue, Field(description="Upper bound for 'between'.")] = None,
    formula: Annotated[
        str | None, Field(description="For type 'custom': a formula that must be TRUE.")
    ] = None,
    allow_blank: bool = True,
    error_message: Annotated[
        str | None, Field(description="Message shown when invalid data is entered.")
    ] = None,
    error_title: str | None = None,
    error_style: Annotated[
        Literal["stop", "warning", "info"],
        Field(description="'stop' rejects bad input; 'warning'/'info' only warn."),
    ] = "stop",
    input_message: Annotated[
        str | None, Field(description="Hint shown when the cell is selected.")
    ] = None,
    input_title: str | None = None,
) -> dict[str, Any]:
    """Restrict what can be typed into cells: a drop-down list, a number/date range, a text
    length, or a custom formula. Existing values are not checked or changed."""
    return await edit(
        path, "calc.set_data_validation", range=range, type=type, sheet=sheet, items=items,
        source_range=source_range, source_sheet=source_sheet, operator=operator,
        min_value=min_value, max_value=max_value, formula=formula, allow_blank=allow_blank,
        error_message=error_message, error_title=error_title, error_style=error_style,
        input_message=input_message, input_title=input_title,
    )  # fmt: skip
