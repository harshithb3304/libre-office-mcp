"""Calc: charts, images, hyperlinks and cell comments."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from com.sun.star.awt import Point, Rectangle, Size

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.ops.calc import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, file_url, pv

# -- charts --------------------------------------------------------------------

_CHART_SERVICES = {
    "column": "com.sun.star.chart.BarDiagram",
    "bar": "com.sun.star.chart.BarDiagram",
    "line": "com.sun.star.chart.LineDiagram",
    "pie": "com.sun.star.chart.PieDiagram",
    "scatter": "com.sun.star.chart.XYDiagram",
    "area": "com.sun.star.chart.AreaDiagram",
    "radar": "com.sun.star.chart.NetDiagram",
}
_SERVICE_TO_TYPE = {
    "com.sun.star.chart.LineDiagram": "line",
    "com.sun.star.chart.PieDiagram": "pie",
    "com.sun.star.chart.XYDiagram": "scatter",
    "com.sun.star.chart.AreaDiagram": "area",
    "com.sun.star.chart.NetDiagram": "radar",
}
_LEGEND_POSITIONS = {"top": "TOP", "bottom": "BOTTOM", "left": "LEFT", "right": "RIGHT"}
DEFAULT_CHART_MM = (140.0, 80.0)


def _chart_type_name(diagram: Any) -> str:
    kind = diagram.getDiagramType()
    if kind == "com.sun.star.chart.BarDiagram":
        return "bar" if diagram.Vertical else "column"
    return _SERVICE_TO_TYPE.get(kind, kind.rsplit(".", 1)[-1])


def _chart_list(sheet: Any) -> list[tuple[str, str, Any]]:
    """``(display name, internal name, drawing shape)`` for each chart on the sheet.

    The shape's own name is what survives a save as .xlsx (the internal name becomes
    "Object 1"), so tools show and accept that one; the internal name is the fallback.
    """
    internal = set(sheet.getCharts().getElementNames())
    page = sheet.getDrawPage()
    found = []
    for i in range(page.getCount()):
        shape = page.getByIndex(i)
        if shape.getShapeType().endswith("OLE2Shape") and shape.PersistName in internal:
            found.append((shape.Name or shape.PersistName, shape.PersistName, shape))
    return found


def _find_chart(sheet: Any, name: str) -> tuple[Any, str, Any]:
    """Resolve a chart by name -> (table chart, internal name, shape)."""
    charts = _chart_list(sheet)
    for display, internal, shape in charts:
        if name in (display, internal):
            return sheet.getCharts().getByName(internal), internal, shape
    raise OpError(
        f"No chart named {name!r} on sheet {sheet.Name!r}. Charts here: "
        f"{', '.join(repr(c[0]) for c in charts) or 'none'}. Use calc_list_charts.",
        protocol.E_NOT_FOUND,
    )


def _configure_chart(
    emb: Any,
    chart_type: str | None,
    title: str | None,
    subtitle: str | None,
    x_axis_title: str | None,
    y_axis_title: str | None,
    legend: bool | None,
    legend_position: str | None,
    stacked: bool | None,
    series_in_rows: bool | None,
    series_colors: list[str] | None,
) -> None:
    if chart_type is not None:
        emb.setDiagram(emb.createInstance(_CHART_SERVICES[chart_type]))
    diagram = emb.getDiagram()
    if chart_type == "bar":
        diagram.Vertical = True  # horizontal bars
    elif chart_type == "column":
        diagram.Vertical = False
    if chart_type == "scatter":
        diagram.Lines = False  # points only; connecting lines make a scatter plot unreadable
    if stacked is not None and hasattr(diagram, "Stacked"):
        diagram.Stacked = stacked
    if series_in_rows is not None:
        diagram.DataRowSource = enum(
            "com.sun.star.chart.ChartDataRowSource", "ROWS" if series_in_rows else "COLUMNS"
        )
    if title is not None:
        emb.HasMainTitle = bool(title)
        if title:
            emb.Title.String = title
    if subtitle is not None:
        emb.HasSubTitle = bool(subtitle)
        if subtitle:
            emb.SubTitle.String = subtitle
    for text, flag, holder in (
        (x_axis_title, "HasXAxisTitle", "XAxisTitle"),
        (y_axis_title, "HasYAxisTitle", "YAxisTitle"),
    ):
        if text is not None and hasattr(diagram, flag):
            setattr(diagram, flag, bool(text))
            if text:
                getattr(diagram, holder).String = text
    if legend is not None:
        emb.HasLegend = legend
    if legend_position and emb.HasLegend:
        emb.Legend.Alignment = enum(
            "com.sun.star.chart.ChartLegendPosition", _LEGEND_POSITIONS[legend_position]
        )
    if series_colors:
        for index, color in enumerate(series_colors):
            series = diagram.getDataRowProperties(index)
            info = series.getPropertySetInfo()
            # bars/areas are filled ("FillColor"); lines and symbols use "Color"
            for prop in ("FillColor", "Color"):
                if info.hasPropertyByName(prop):
                    series.setPropertyValue(prop, h.color_value(color))


def _validate_chart_args(chart_type: str | None, legend_position: str | None) -> None:
    if chart_type is not None and chart_type not in _CHART_SERVICES:
        raise OpError(f"chart_type must be one of: {', '.join(_CHART_SERVICES)}.")
    if legend_position is not None and legend_position not in _LEGEND_POSITIONS:
        raise OpError(f"legend_position must be one of: {', '.join(_LEGEND_POSITIONS)}.")


@op("calc.add_chart", kind="calc", mutates=True)
@h.warn_in_csv
def add_chart(
    s: Session,
    data_range: str,
    chart_type: str = "column",
    sheet: str | None = None,
    name: str | None = None,
    title: str | None = None,
    subtitle: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    anchor_cell: str | None = None,
    width_mm: float = DEFAULT_CHART_MM[0],
    height_mm: float = DEFAULT_CHART_MM[1],
    first_row_is_header: bool = True,
    first_column_is_labels: bool = True,
    series_in_rows: bool | None = None,
    stacked: bool | None = None,
    legend: bool | None = None,
    legend_position: str | None = None,
    series_colors: list[str] | None = None,
) -> dict[str, Any]:
    """Create a chart from a data range and place it at a cell."""
    _validate_chart_args(chart_type, legend_position)
    sh, rng, box = h.get_range(s, sheet, data_range)
    if box[2] == box[0] and box[3] == box[1]:
        raise OpError("The chart data range must cover at least two cells, e.g. 'A1:B6'.")
    charts = sh.getCharts()
    taken = {c[0] for c in _chart_list(sh)} | set(charts.getElementNames())
    if name is None:
        n = 1
        while f"Chart{n}" in taken:
            n += 1
        name = f"Chart{n}"
    elif name in taken:
        raise OpError(f"A chart named {name!r} already exists on this sheet; choose another name.")
    if anchor_cell is None:
        used = h.used_box(sh)
        # default: next to the data, two columns to the right of the used area
        anchor = (max(box[2], used[2] if used else 0) + 2, box[1])
    else:
        anchor = h.parse_cell(anchor_cell)
    x, y = h.cell_position_mm100(sh, *anchor)
    rect = Rectangle(x, y, h.mm(width_mm), h.mm(height_mm))
    charts.addNewByName(
        name, rect, (rng.getRangeAddress(),), first_row_is_header, first_column_is_labels
    )
    _, _, shape = _find_chart(sh, name)
    shape.Name = name  # survives .xlsx, unlike the internal object name
    if legend is None:  # a legend only helps when there is more than one series to tell apart
        across = (
            (box[3] - box[1] + 1 - first_row_is_header)
            if series_in_rows
            else (box[2] - box[0] + 1 - first_column_is_labels)
        )
        legend = chart_type == "pie" or across > 1
    emb = charts.getByName(name).getEmbeddedObject()
    _configure_chart(
        emb, chart_type, title, subtitle, x_axis_title, y_axis_title, legend, legend_position,
        stacked, series_in_rows, series_colors,
    )  # fmt: skip
    return {
        "sheet": sh.Name,
        "chart": name,
        "type": chart_type,
        "data_range": h.box_name(box),
        "anchor_cell": h.cell_name(*anchor),
    }


@op("calc.update_chart", kind="calc", mutates=True)
@h.warn_in_csv
def update_chart(
    s: Session,
    chart: str,
    sheet: str | None = None,
    chart_type: str | None = None,
    title: str | None = None,
    subtitle: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    stacked: bool | None = None,
    series_in_rows: bool | None = None,
    legend: bool | None = None,
    legend_position: str | None = None,
    series_colors: list[str] | None = None,
) -> dict[str, Any]:
    """Change type, titles, legend, stacking or colours of an existing chart."""
    _validate_chart_args(chart_type, legend_position)
    sh = h.get_sheet(s, sheet)
    emb = _find_chart(sh, chart)[0].getEmbeddedObject()
    _configure_chart(
        emb, chart_type, title, subtitle, x_axis_title, y_axis_title, legend, legend_position,
        stacked, series_in_rows, series_colors,
    )  # fmt: skip
    return {"sheet": sh.Name, "chart": chart, "type": _chart_type_name(emb.getDiagram())}


@op("calc.list_charts", kind="calc")
def list_charts(s: Session, sheet: str | None = None) -> dict[str, Any]:
    """Charts with type, title, data ranges, anchor cell and size."""
    out = []
    for sh in h.sheets_for(s, sheet):
        for display, _, shape in _chart_list(sh):
            table_chart = _find_chart(sh, display)[0]
            emb = table_chart.getEmbeddedObject()
            x, y = shape.Position.X, shape.Position.Y
            out.append(
                {
                    "sheet": sh.Name,
                    "chart": display,
                    "type": _chart_type_name(emb.getDiagram()),
                    "title": emb.Title.String if emb.HasMainTitle else None,
                    "legend": bool(emb.HasLegend),
                    "data_ranges": [
                        h.box_name(h.box_of_address(a)) for a in table_chart.getRanges()
                    ],
                    "anchor_cell": h.cell_name(*h.cell_at_position(sh, x, y)),
                    "position_mm": [round(x / 100, 1), round(y / 100, 1)],
                    "size_mm": [
                        round(shape.Size.Width / 100, 1),
                        round(shape.Size.Height / 100, 1),
                    ],
                }
            )
    return {"count": len(out), "charts": out}


@op("calc.delete_chart", kind="calc", mutates=True)
def delete_chart(s: Session, chart: str, sheet: str | None = None) -> dict[str, Any]:
    """Remove a chart; its data stays."""
    sh = h.get_sheet(s, sheet)
    _, internal, _ = _find_chart(sh, chart)
    sh.getCharts().removeByName(internal)
    return {"sheet": sh.Name, "deleted_chart": chart}


# -- images --------------------------------------------------------------------


def _images(sheet: Any) -> list[Any]:
    page = sheet.getDrawPage()
    shapes = [page.getByIndex(i) for i in range(page.getCount())]
    return [x for x in shapes if x.getShapeType().endswith("GraphicObjectShape")]


@op("calc.insert_image", kind="calc", mutates=True)
@h.warn_in_csv
def insert_image(
    s: Session,
    image_path: str,
    sheet: str | None = None,
    anchor_cell: str = "A1",
    width_mm: float | None = None,
    height_mm: float | None = None,
    name: str | None = None,
) -> dict[str, Any]:
    """Embed a picture anchored at a cell, keeping its aspect ratio."""
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(anchor_cell)
    provider = s.service("com.sun.star.graphic.GraphicProvider")
    graphic = provider.queryGraphic((pv("URL", file_url(image_path)),))
    if graphic is None:
        raise OpError(
            f"{Path(image_path).name} is not an image LibreOffice can read.", protocol.E_UNSUPPORTED
        )
    size = graphic.Size100thMM
    if size.Width > 0 and size.Height > 0:
        native_w, native_h = size.Width, size.Height
    else:  # no DPI information in the file: assume 96 DPI
        native_w = graphic.SizePixel.Width * 26.46
        native_h = graphic.SizePixel.Height * 26.46
    ratio = native_h / native_w
    if width_mm and height_mm:
        width, height = h.mm(width_mm), h.mm(height_mm)
    elif width_mm:
        width = h.mm(width_mm)
        height = int(width * ratio)
    elif height_mm:
        height = h.mm(height_mm)
        width = int(height / ratio)
    else:  # keep the native size but never wider than 160 mm
        width = int(min(native_w, 16000))
        height = int(width * ratio)
    existing = {x.Name for x in _images(sh)}
    if name is None:
        n = 1
        while f"Image{n}" in existing:
            n += 1
        name = f"Image{n}"
    elif name in existing:
        raise OpError(f"An image named {name!r} already exists on this sheet.")
    shape = s.new("com.sun.star.drawing.GraphicObjectShape")
    shape.Graphic = graphic  # embeds the picture (GraphicURL would only link to the file)
    sh.getDrawPage().add(shape)
    x, y = h.cell_position_mm100(sh, col, row)
    shape.Position = Point(x, y)
    shape.Size = Size(width, height)
    shape.Name = name
    shape.setPropertyValue("Anchor", sh.getCellByPosition(col, row))  # move with the cell
    return {
        "sheet": sh.Name,
        "image": name,
        "anchor_cell": h.cell_name(col, row),
        "width_mm": round(width / 100, 1),
        "height_mm": round(height / 100, 1),
    }


@op("calc.list_images", kind="calc")
def list_images(s: Session, sheet: str | None = None) -> dict[str, Any]:
    """Pictures on the sheets with anchor cell and size."""
    out = []
    for sh in h.sheets_for(s, sheet):
        for shape in _images(sh):
            info: dict[str, Any] = {
                "sheet": sh.Name,
                "image": shape.Name,
                "position_mm": [round(shape.Position.X / 100, 1), round(shape.Position.Y / 100, 1)],
                "size_mm": [round(shape.Size.Width / 100, 1), round(shape.Size.Height / 100, 1)],
            }
            info["anchor_cell"] = h.cell_name(
                *h.cell_at_position(sh, shape.Position.X, shape.Position.Y)
            )
            out.append(info)
    return {"count": len(out), "images": out}


@op("calc.delete_image", kind="calc", mutates=True)
def delete_image(s: Session, image: str, sheet: str | None = None) -> dict[str, Any]:
    """Remove a picture by name."""
    sh = h.get_sheet(s, sheet)
    images = _images(sh)
    for shape in images:
        if shape.Name == image:
            sh.getDrawPage().remove(shape)
            return {"sheet": sh.Name, "deleted_image": image}
    raise OpError(
        f"No image named {image!r} on sheet {sh.Name!r}. Images here: "
        f"{', '.join(repr(x.Name) for x in images) or 'none'}. Use calc_list_images.",
        protocol.E_NOT_FOUND,
    )


# -- hyperlinks ----------------------------------------------------------------


@op("calc.set_hyperlink", kind="calc", mutates=True)
@h.warn_in_csv
def set_hyperlink(
    s: Session, cell: str, url: str, sheet: str | None = None, text: str | None = None
) -> dict[str, Any]:
    """Make a cell a URL field with the given text."""
    if not url.strip():
        raise OpError("url must not be empty.")
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(cell)
    target = sh.getCellByPosition(col, row)
    label = text if text is not None else (target.getString() or url)
    field = s.new("com.sun.star.text.TextField.URL")
    field.URL = url
    field.Representation = label
    target.setString("")  # drop any previous content or link
    target.getText().insertTextContent(target.getText().createTextCursor(), field, False)
    return {"sheet": sh.Name, "cell": h.cell_name(col, row), "url": url, "text": label}


# -- comments ------------------------------------------------------------------


def _comment_at(sheet: Any, col: int, row: int) -> Any:
    annotations = sheet.getAnnotations()
    for i in range(annotations.getCount()):
        annotation = annotations.getByIndex(i)
        pos = annotation.getPosition()
        if pos.Column == col and pos.Row == row:
            return i, annotation
    return None


@op("calc.add_comment", kind="calc", mutates=True)
@h.warn_in_csv
def add_comment(
    s: Session, cell: str, text: str, sheet: str | None = None, visible: bool = False
) -> dict[str, Any]:
    """Attach (or replace) a comment on a cell."""
    if not text.strip():
        raise OpError("text must not be empty; use calc_delete_comment to remove a comment.")
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(cell)
    annotations = sh.getAnnotations()
    existing = _comment_at(sh, col, row)
    if existing is not None:
        annotations.removeByIndex(existing[0])
    annotations.insertNew(h.cell_address(sh, col, row), text)
    _comment_at(sh, col, row)[1].setIsVisible(visible)
    return {
        "sheet": sh.Name,
        "cell": h.cell_name(col, row),
        "replaced": existing is not None,
    }


@op("calc.delete_comment", kind="calc", mutates=True)
def delete_comment(s: Session, cell: str, sheet: str | None = None) -> dict[str, Any]:
    """Remove a cell's comment."""
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(cell)
    existing = _comment_at(sh, col, row)
    if existing is None:
        return {"sheet": sh.Name, "cell": h.cell_name(col, row), "deleted": False, "_noop": True}
    sh.getAnnotations().removeByIndex(existing[0])
    return {"sheet": sh.Name, "cell": h.cell_name(col, row), "deleted": True}
