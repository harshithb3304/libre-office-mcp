"""Calc: read-only operations (``calc.list_*`` / ``calc.get_*`` / ``calc.find_text``)."""

from __future__ import annotations

import re
from typing import Any

from libreoffice_mcp.worker.ops.calc import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session

MAX_FORMATTED_CELLS = 2000
MAX_FIND_RESULTS = 500


def _hex(color: int) -> str | None:
    return None if color is None or color < 0 else f"#{color & 0xFFFFFF:06X}"


def _cell_value(cell: Any) -> Any:
    kind = cell.getType().value
    if kind == "EMPTY":
        return None
    if kind == "TEXT":
        return cell.getString()
    if kind == "VALUE":
        return h.clean_number(cell.getValue())
    # FORMULA: error, text or number result
    if cell.getError():
        return cell.getString()
    return h.clean_number(cell.getValue()) if cell.FormulaResultType2 == 1 else cell.getString()


# -- sheets --------------------------------------------------------------------


@op("calc.list_sheets", kind="calc")
def list_sheets(s: Session) -> dict[str, Any]:
    """Every sheet with its used range, size and visibility."""
    sheets = s.doc.getSheets()
    out = []
    for index in range(sheets.getCount()):
        sheet = sheets.getByIndex(index)
        box = h.used_box(sheet)
        info: dict[str, Any] = {
            "index": index,
            "name": sheet.Name,
            "visible": bool(sheet.IsVisible),
            "used_range": h.box_name(box) if box else None,
            "rows": box[3] - box[1] + 1 if box else 0,
            "columns": box[2] - box[0] + 1 if box else 0,
            "charts": sheet.getCharts().getCount(),
        }
        tab_color = _hex(sheet.TabColor)
        if tab_color:
            info["tab_color"] = tab_color
        out.append(info)
    return {"count": len(out), "sheets": out}


@op("calc.get_range", kind="calc")
def get_range(
    s: Session,
    sheet: str | None = None,
    range: str | None = None,
    include_formulas: bool = False,
    formatted: bool = False,
    max_cells: int = 5000,
) -> dict[str, Any]:
    """Cell values of a range as a 2-D list (paged by rows)."""
    sh = h.get_sheet(s, sheet)
    if range is None:
        box = h.used_box(sh)
        if box is None:
            return {"sheet": sh.Name, "range": None, "rows": 0, "columns": 0, "values": []}
    else:
        box = h.parse_range(range, whole=True)
        if box[3] == h.MAX_ROW or box[2] == h.MAX_COL:  # whole rows/columns: clip to the data
            used = h.used_box(sh)
            if used is None:
                return {"sheet": sh.Name, "range": range, "rows": 0, "columns": 0, "values": []}
            box = (
                max(box[0], used[0]) if box[2] == h.MAX_COL else box[0],
                max(box[1], used[1]) if box[3] == h.MAX_ROW else box[1],
                min(box[2], used[2]),
                min(box[3], used[3]),
            )
    c1, r1, c2, r2 = box
    width = c2 - c1 + 1
    max_cells = max(1, min(max_cells, h.MAX_CELLS))
    if formatted and width * (r2 - r1 + 1) > MAX_FORMATTED_CELLS:
        max_cells = min(max_cells, MAX_FORMATTED_CELLS)
    rows_fit = max(1, max_cells // width)
    last = min(r2, r1 + rows_fit - 1)
    rng = h.range_of(sh, (c1, r1, c2, last))
    data = rng.getDataArray()
    values = [[h.clean_number(v) if isinstance(v, float) else v for v in row] for row in data]
    result: dict[str, Any] = {
        "sheet": sh.Name,
        "range": h.box_name((c1, r1, c2, last)),
        "rows": last - r1 + 1,
        "columns": width,
        "values": values,
    }
    if formatted:
        if width * (last - r1 + 1) > MAX_FORMATTED_CELLS:
            raise OpError(
                f"formatted=true is limited to {MAX_FORMATTED_CELLS} cells per call; "
                "request a smaller range."
            )
        result["formatted"] = [
            [sh.getCellByPosition(c, r).getString() for c in h.span(c1, c2)]
            for r in h.span(r1, last)
        ]
    if include_formulas:
        formulas = {}
        for r, row in enumerate(rng.getFormulaArray()):
            for c, text in enumerate(row):
                if text.startswith("="):
                    formulas[h.cell_name(c1 + c, r1 + r)] = text
        result["formulas"] = formulas
    if last < r2:
        result["truncated"] = True
        result["next_range"] = h.box_name((c1, last + 1, c2, r2))
    return result


@op("calc.get_cell_info", kind="calc")
def get_cell_info(s: Session, cell: str, sheet: str | None = None) -> dict[str, Any]:
    """Value, formula, number format, font, fill, alignment and attachments of one cell."""
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(cell)
    try:
        c = sh.getCellByPosition(col, row)
    except Exception as exc:
        raise OpError(f"Cell {cell!r} is outside the sheet.") from exc
    kind = c.getType().value
    formats = s.doc.getNumberFormats()
    info: dict[str, Any] = {
        "sheet": sh.Name,
        "cell": h.cell_name(col, row),
        "type": kind.lower(),
        "value": _cell_value(c),
        "displayed": c.getString(),
    }
    if kind == "FORMULA":
        info["formula"] = c.getFormula()
        if c.getError():
            info["error_code"] = c.getError()
    try:
        info["number_format"] = formats.getByKey(c.NumberFormat).FormatString
    except Exception:
        info["number_format"] = None
    info["font"] = {
        "name": c.CharFontName,
        "size": c.CharHeight,
        "bold": c.CharWeight >= 150,
        "italic": c.CharPosture.value != "NONE",
        "underline": c.CharUnderline != 0,
        "color": _hex(c.CharColor) or "#000000",
    }
    info["background"] = None if c.IsCellBackgroundTransparent else _hex(c.CellBackColor)
    info["h_align"] = c.HoriJustify.value.lower()
    info["v_align"] = {1: "top", 2: "middle", 3: "bottom"}.get(c.VertJustify, "bottom")
    info["wrap_text"] = bool(c.IsTextWrapped)
    if c.CellStyle:
        info["cell_style"] = c.CellStyle
    # merged?
    cursor = sh.createCursorByRange(c)
    cursor.collapseToMergedArea()
    merged = h.box_of(cursor)
    if merged != (col, row, col, row):
        info["merged_range"] = h.box_name(merged)
    # hyperlink / comment / validation
    url = _cell_hyperlink(c)
    if url:
        info["hyperlink"] = url
    annotation = _annotation_at(sh, col, row)
    if annotation is not None:
        info["comment"] = annotation.getString()
    validation = c.Validation
    if validation.Type.value != "ANY":
        info["validation"] = {
            "type": validation.Type.value.lower(),
            "formula1": validation.getFormula1(),
            "formula2": validation.getFormula2(),
        }
    return info


def _annotation_at(sheet: Any, col: int, row: int) -> Any:
    annotations = sheet.getAnnotations()
    for i in range(annotations.getCount()):
        a = annotations.getByIndex(i)
        pos = a.getPosition()
        if pos.Column == col and pos.Row == row:
            return a
    return None


def _cell_hyperlink(cell: Any) -> str | None:
    if cell.getType().value != "TEXT":
        return None
    paragraphs = cell.getText().createEnumeration()
    while paragraphs.hasMoreElements():
        portions = paragraphs.nextElement().createEnumeration()
        while portions.hasMoreElements():
            portion = portions.nextElement()
            if portion.TextPortionType == "TextField":
                field = portion.TextField
                if field.supportsService("com.sun.star.text.TextField.URL"):
                    return field.URL
    return None


# -- searching -----------------------------------------------------------------


def _matcher(query: str, regex: bool, match_case: bool, whole_cell: bool):
    flags = 0 if match_case else re.IGNORECASE
    if regex:
        try:
            pattern = re.compile(f"^(?:{query})$" if whole_cell else query, flags)
        except re.error as exc:
            raise OpError(f"Invalid regular expression {query!r}: {exc}") from exc
        return lambda text: pattern.search(text) is not None
    needle = query if match_case else query.lower()

    def plain(text: str) -> bool:
        hay = text if match_case else text.lower()
        return hay == needle if whole_cell else needle in hay

    return plain


@op("calc.find_text", kind="calc")
def find_text(
    s: Session,
    query: str,
    sheet: str | None = None,
    regex: bool = False,
    match_case: bool = False,
    whole_cell: bool = False,
    search_in: str = "values",
    max_results: int = 100,
) -> dict[str, Any]:
    """Find cells whose value (or formula) matches; searches all sheets unless one is named."""
    if not query:
        raise OpError("query must not be empty.")
    matches = _matcher(query, regex, match_case, whole_cell)
    sheets = h.sheets_for(s, sheet)
    max_results = max(1, min(max_results, MAX_FIND_RESULTS))
    hits: list[dict[str, Any]] = []
    total = 0
    for sh in sheets:
        box = h.used_box(sh)
        if box is None:
            continue
        c1, r1, c2, r2 = box
        width = c2 - c1 + 1
        step = max(1, h.MAX_CELLS // width)
        for start in range(r1, r2 + 1, step):
            end = min(r2, start + step - 1)
            rng = h.range_of(sh, (c1, start, c2, end))
            data = rng.getDataArray()
            forms = rng.getFormulaArray() if search_in == "formulas" else None
            for dr, row in enumerate(data):
                for dc, value in enumerate(row):
                    if value == "":
                        continue
                    if forms is not None:
                        text = forms[dr][dc]
                    else:
                        text = str(h.clean_number(value)) if isinstance(value, float) else value
                    if matches(text):
                        total += 1
                        if len(hits) < max_results:
                            hit = {
                                "sheet": sh.Name,
                                "cell": h.cell_name(c1 + dc, start + dr),
                                "value": h.clean_number(value)
                                if isinstance(value, float)
                                else value,
                            }
                            if forms is not None and forms[dr][dc].startswith("="):
                                hit["formula"] = forms[dr][dc]
                            hits.append(hit)
    return {
        "query": query,
        "total_matches": total,
        "returned": len(hits),
        "truncated": total > len(hits),
        "matches": hits,
    }


# -- named ranges / layout -----------------------------------------------------


@op("calc.list_named_ranges", kind="calc")
def list_named_ranges(s: Session) -> dict[str, Any]:
    """Defined names and what they refer to."""
    named = s.doc.NamedRanges
    out = []
    for name in named.getElementNames():
        item = named.getByName(name)
        out.append({"name": name, "refers_to": item.getContent().replace("$", "")})
    return {"count": len(out), "named_ranges": out}


@op("calc.get_sheet_layout", kind="calc")
def get_sheet_layout(s: Session, sheet: str | None = None) -> dict[str, Any]:
    """Used range, freeze panes, auto-filter, column widths, hidden rows/columns, merged cells."""
    sh = h.get_sheet(s, sheet)
    box = h.used_box(sh)
    frozen_cols, frozen_rows = h.get_freeze(s.doc, h.sheet_index(sh))
    layout: dict[str, Any] = {
        "sheet": sh.Name,
        "used_range": h.box_name(box) if box else None,
        "frozen_columns": frozen_cols,
        "frozen_rows": frozen_rows,
        "auto_filter": _autofilter_range(s, sh),
    }
    page = sh.getDrawPage()
    layout["images"] = sum(
        1
        for i in range(page.getCount())
        if page.getByIndex(i).getShapeType().endswith("GraphicObjectShape")
    )
    layout["charts"] = sh.getCharts().getCount()
    if box is None:
        return layout
    c1, r1, c2, r2 = box
    widths: dict[str, float] = {}
    hidden_cols: list[str] = []
    columns = sh.getColumns()
    for col in range(c1, min(c2, c1 + 59) + 1):
        column = columns.getByIndex(col)
        if not column.IsVisible:
            hidden_cols.append(h.col_name(col))
        widths[h.col_name(col)] = round(column.Width / 100, 1)
    layout["column_widths_mm"] = widths
    layout["hidden_columns"] = hidden_cols
    rows = sh.getRows()
    hidden_rows: list[int] = []
    heights: dict[str, float] = {}
    for r in range(r1, min(r2, r1 + 199) + 1):
        row = rows.getByIndex(r)
        if not row.IsVisible:
            hidden_rows.append(r + 1)
        elif not row.OptimalHeight:
            heights[str(r + 1)] = round(row.Height / 100, 1)
    layout["hidden_rows"] = hidden_rows
    layout["custom_row_heights_mm"] = heights
    layout["merged_ranges"] = _merged_ranges(sh, box)
    return layout


def _autofilter_range(s: Session, sheet: Any) -> str | None:
    index = h.sheet_index(sheet)
    databases = s.doc.DatabaseRanges
    for name in databases.getElementNames():
        db = databases.getByName(name)
        area = db.getDataArea()
        if area.Sheet == index and db.AutoFilter:
            return h.box_name((area.StartColumn, area.StartRow, area.EndColumn, area.EndRow))
    return None


def _merged_ranges(sheet: Any, box: h.Box, limit: int = 100) -> list[str]:
    """Merged areas inside ``box``.

    Calc has no "list merges" call. Merged cells carry their own cell attributes, so the
    uniform-format ranges of the used area split around them; only those pieces are checked.
    """
    merged: list[str] = []
    seen: set[h.Box] = set()
    formats = h.range_of(sheet, box).getCellFormatRanges()
    for i in range(min(formats.getCount(), 2000)):
        piece = formats.getByIndex(i)
        pbox = h.box_of(piece)
        if h.box_cells(pbox) > 4000 or not piece.getIsMerged():
            continue
        cursor = sheet.createCursorByRange(sheet.getCellByPosition(pbox[0], pbox[1]))
        cursor.collapseToMergedArea()
        area = h.box_of(cursor)
        if area not in seen:
            seen.add(area)
            merged.append(h.box_name(area))
            if len(merged) >= limit:
                break
    return merged


# -- conditional formats / validation ------------------------------------------

#: ScConditionMode values used by the conditional-format API (see conditional.py).
CONDITION_NAMES = {
    0: "equal", 1: "less", 2: "greater", 3: "less_equal", 4: "greater_equal", 5: "not_equal",
    6: "between", 7: "not_between", 8: "duplicate", 9: "unique", 24: "formula",
}  # fmt: skip


@op("calc.list_conditional_formats", kind="calc")
def list_conditional_formats(s: Session, sheet: str | None = None) -> dict[str, Any]:
    """Conditional formats of a sheet: ranges and rules."""
    sh = h.get_sheet(s, sheet)
    out = []
    for fmt in sh.ConditionalFormats.getConditionalFormats():
        ranges = [h.box_name(h.box_of_address(a)) for a in fmt.Range.getRangeAddresses()]
        rules = []
        for i in range(fmt.getCount()):
            entry = fmt.getByIndex(i)
            names = {p.Name for p in entry.getPropertySetInfo().getProperties()}
            if "ColorScaleEntries" in names:
                rules.append({"type": "color_scale"})
            elif "DataBarEntries" in names:
                rules.append({"type": "data_bar", "color": _hex(entry.Color)})
            elif "IconSetEntries" in names:
                rules.append({"type": "icon_set"})
            else:
                rule: dict[str, Any] = {
                    "type": CONDITION_NAMES.get(entry.Operator, f"operator_{entry.Operator}"),
                    "value": entry.Formula1,
                    "style": entry.StyleName,
                }
                if entry.Formula2:
                    rule["value2"] = entry.Formula2
                rules.append(rule)
        out.append({"ranges": ranges, "rules": rules})
    return {"sheet": sh.Name, "count": len(out), "conditional_formats": out}


@op("calc.list_comments", kind="calc")
def list_comments(s: Session, sheet: str | None = None) -> dict[str, Any]:
    """Cell comments on one or all sheets."""
    sheets = h.sheets_for(s, sheet)
    out = []
    for sh in sheets:
        annotations = sh.getAnnotations()
        for i in range(annotations.getCount()):
            a = annotations.getByIndex(i)
            pos = a.getPosition()
            out.append(
                {
                    "sheet": sh.Name,
                    "cell": h.cell_name(pos.Column, pos.Row),
                    "text": a.getString(),
                    "author": a.getAuthor(),
                    "visible": bool(a.getIsVisible()),
                }
            )
    return {"count": len(out), "comments": out}
