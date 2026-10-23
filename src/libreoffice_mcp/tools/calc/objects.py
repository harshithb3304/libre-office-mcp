"""Calc tools for charts, images, hyperlinks and comments."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.calc.types import (
    CellA1,
    ChartType,
    Color,
    LegendPosition,
    RangeA1,
    SheetName,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath

C = "calc"

ChartName = Annotated[str, Field(description="Chart name from calc_list_charts / calc_add_chart.")]
Millimetres = Annotated[float | None, Field(gt=0, le=2000)]
Titles = Annotated[str | None, Field(description="Text; '' removes it.")]


@tool(C, "edit", title="Add chart")
async def calc_add_chart(
    path: DocPath,
    data_range: Annotated[
        RangeA1,
        Field(
            description="Cells to plot, normally including headers: 'A1:C7' = category labels in "
            "column A, series names in row 1, one series per further column."
        ),
    ],
    chart_type: ChartType = "column",
    sheet: SheetName = None,
    name: Annotated[str | None, Field(description="Chart name. Default: Chart1, Chart2...")] = None,
    title: Titles = None,
    subtitle: Titles = None,
    x_axis_title: Titles = None,
    y_axis_title: Titles = None,
    anchor_cell: Annotated[
        CellA1 | None,
        Field(description="Top-left cell of the chart. Default: right of the data."),
    ] = None,
    width_mm: Annotated[float, Field(gt=20, le=1000)] = 140.0,
    height_mm: Annotated[float, Field(gt=20, le=1000)] = 80.0,
    first_row_is_header: Annotated[
        bool, Field(description="The first row holds series names (not data).")
    ] = True,
    first_column_is_labels: Annotated[
        bool, Field(description="The first column holds category labels (not data).")
    ] = True,
    series_in_rows: Annotated[
        bool | None, Field(description="true = each ROW is a series (default: each column).")
    ] = None,
    stacked: Annotated[
        bool | None, Field(description="Stack the series (column/bar/area).")
    ] = None,
    legend: Annotated[bool | None, Field(description="Show or hide the legend.")] = None,
    legend_position: LegendPosition | None = None,
    series_colors: Annotated[
        list[Color] | None, Field(description="Colour per series, in order.")
    ] = None,
) -> dict[str, Any]:
    """Create a chart from a data range and place it on the sheet. Types: column, bar
    (horizontal), line, pie, scatter (first column = X values), area, radar. Returns the chart
    name to use with calc_update_chart / calc_delete_chart."""
    return await edit(
        path, "calc.add_chart", data_range=data_range, chart_type=chart_type, sheet=sheet,
        name=name, title=title, subtitle=subtitle, x_axis_title=x_axis_title,
        y_axis_title=y_axis_title, anchor_cell=anchor_cell, width_mm=width_mm,
        height_mm=height_mm, first_row_is_header=first_row_is_header,
        first_column_is_labels=first_column_is_labels, series_in_rows=series_in_rows,
        stacked=stacked, legend=legend, legend_position=legend_position,
        series_colors=series_colors,
    )  # fmt: skip


@tool(C, "edit", title="Update chart")
async def calc_update_chart(
    path: DocPath,
    chart: ChartName,
    sheet: SheetName = None,
    chart_type: ChartType | None = None,
    title: Titles = None,
    subtitle: Titles = None,
    x_axis_title: Titles = None,
    y_axis_title: Titles = None,
    stacked: bool | None = None,
    series_in_rows: bool | None = None,
    legend: bool | None = None,
    legend_position: LegendPosition | None = None,
    series_colors: list[Color] | None = None,
) -> dict[str, Any]:
    """Change an existing chart's type, titles, legend, stacking or series colours. Only the
    options you pass change. To change the plotted data, delete the chart and add a new one."""
    return await edit(
        path, "calc.update_chart", chart=chart, sheet=sheet, chart_type=chart_type, title=title,
        subtitle=subtitle, x_axis_title=x_axis_title, y_axis_title=y_axis_title, stacked=stacked,
        series_in_rows=series_in_rows, legend=legend, legend_position=legend_position,
        series_colors=series_colors,
    )  # fmt: skip


@tool(C, "destructive", title="Delete chart")
async def calc_delete_chart(
    path: DocPath, chart: ChartName, sheet: SheetName = None
) -> dict[str, Any]:
    """Delete a chart (the data it plotted stays)."""
    return await edit(path, "calc.delete_chart", chart=chart, sheet=sheet)


@tool(C, "edit", title="Insert image")
async def calc_insert_image(
    path: DocPath,
    image_path: Annotated[str, Field(description="Path to a PNG/JPG/GIF/SVG... file on disk.")],
    sheet: SheetName = None,
    anchor_cell: Annotated[
        CellA1, Field(description="Cell for the image's top-left corner.")
    ] = "A1",
    width_mm: Millimetres = None,
    height_mm: Millimetres = None,
    name: Annotated[str | None, Field(description="Image name. Default: Image1, Image2...")] = None,
) -> dict[str, Any]:
    """Place a picture on a sheet, anchored at a cell. Give one of width_mm/height_mm to keep the
    aspect ratio."""
    return await edit(
        path, "calc.insert_image", image_path=image_path, sheet=sheet, anchor_cell=anchor_cell,
        width_mm=width_mm, height_mm=height_mm, name=name,
    )  # fmt: skip


@tool(C, "destructive", title="Delete image")
async def calc_delete_image(
    path: DocPath,
    image: Annotated[str, Field(description="Image name from calc_list_images.")],
    sheet: SheetName = None,
) -> dict[str, Any]:
    """Remove a picture from a sheet."""
    return await edit(path, "calc.delete_image", image=image, sheet=sheet)


@tool(C, "edit", title="Set hyperlink")
async def calc_set_hyperlink(
    path: DocPath,
    cell: CellA1,
    url: Annotated[str, Field(description="Target, e.g. 'https://example.com' or 'mailto:a@b.c'.")],
    sheet: SheetName = None,
    text: Annotated[
        str | None,
        Field(
            description="Link text shown in the cell. Default: the cell's current text, or the URL."
        ),
    ] = None,
) -> dict[str, Any]:
    """Turn a cell into a clickable hyperlink (replaces the cell's content)."""
    return await edit(path, "calc.set_hyperlink", cell=cell, url=url, sheet=sheet, text=text)


@tool(C, "edit", title="Add comment")
async def calc_add_comment(
    path: DocPath,
    cell: CellA1,
    text: str,
    sheet: SheetName = None,
    visible: Annotated[
        bool, Field(description="Keep the comment always shown (default: only on hover).")
    ] = False,
) -> dict[str, Any]:
    """Attach a comment (note) to a cell; an existing comment on that cell is replaced."""
    return await edit(path, "calc.add_comment", cell=cell, text=text, sheet=sheet, visible=visible)


@tool(C, "destructive", title="Delete comment")
async def calc_delete_comment(
    path: DocPath, cell: CellA1, sheet: SheetName = None
) -> dict[str, Any]:
    """Remove the comment from a cell."""
    return await edit(path, "calc.delete_comment", cell=cell, sheet=sheet)
