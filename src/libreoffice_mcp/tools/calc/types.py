"""Parameter types shared by the Calc tools.

Addressing convention used by every calc_* tool: the sheet is given by *name* in a separate
``sheet`` argument (default: the first sheet) and cells/ranges use A1 notation (``B2``,
``A1:C10``). Rows are 1-based (as in Excel), columns are letters.
"""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field

SheetName = Annotated[
    str | None,
    Field(
        description="Sheet name, e.g. 'Sheet1' (see calc_list_sheets). Default: the first sheet."
    ),
]

RequiredSheet = Annotated[str, Field(description="Sheet name (see calc_list_sheets).")]

RangeA1 = Annotated[
    str,
    Field(
        description="Range in A1 notation, e.g. 'B2' or 'A1:C10'. Do not include the sheet name."
    ),
]

WholeRangeA1 = Annotated[
    str,
    Field(
        description="Range in A1 notation: 'B2', 'A1:C10', or whole columns 'A:C' / rows '2:5'. "
        "Do not include the sheet name."
    ),
]

CellA1 = Annotated[str, Field(description="A single cell in A1 notation, e.g. 'C3'.")]

#: What a cell can hold when writing. Strings starting with '=' are formulas.
CellValue = str | int | float | bool | None

Grid = Annotated[
    list[list[CellValue]],
    Field(
        description="Rows of cell values (a list of rows, each a list). Numbers become numbers, "
        "strings become text, strings starting with '=' become formulas (e.g. '=SUM(B2:B9)'), "
        "booleans become TRUE/FALSE and null leaves the cell empty."
    ),
]

ParseText = Annotated[
    bool,
    Field(
        description="If true, strings are interpreted as if typed into the cell: '12' becomes a "
        "number, '2024-01-31' a date (shown as typed), '50%' a percentage. Default false: "
        "strings stay text, so IDs like '007' keep their zeros. To write dates, either use "
        "this option or a formula such as '=DATE(2024;1;31)' plus a date number format."
    ),
]

Color = Annotated[
    str, Field(description="Colour as '#RRGGBB' (e.g. '#1F4E79') or a basic name like 'red'.")
]

HAlign = Literal["left", "center", "right", "justify"]
VAlign = Literal["top", "middle", "bottom"]

ChartType = Literal["column", "bar", "line", "pie", "scatter", "area", "radar"]
LegendPosition = Literal["top", "bottom", "left", "right"]
