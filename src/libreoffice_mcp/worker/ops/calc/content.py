"""Calc: writing and rearranging cell content."""

from __future__ import annotations

import re
from typing import Any

import uno
from com.sun.star.table import BorderLine2, TableSortField

from libreoffice_mcp.worker.ops.calc import helpers as h
from libreoffice_mcp.worker.ops.calc.sheets import enable_autofilter
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum

MAX_WRITE_CELLS = 100_000

# com.sun.star.sheet.CellFlags
_VALUE, _DATETIME, _STRING, _ANNOTATION, _FORMULA = 1, 2, 4, 8, 16
_HARDATTR, _STYLES, _EDITATTR = 32, 64, 256
_CONTENTS = _VALUE | _DATETIME | _STRING | _FORMULA
_FORMATS = _HARDATTR | _STYLES | _EDITATTR


# -- writing values ------------------------------------------------------------


def _validate_grid(data: Any) -> tuple[int, int]:
    if not isinstance(data, list) or not data or not all(isinstance(r, list) for r in data):
        raise OpError(
            "data must be a non-empty list of rows, each row a list of cell values, "
            'e.g. [["Name", "Qty"], ["Apples", 3]].'
        )
    cols = max(len(r) for r in data)
    if cols == 0:
        raise OpError("data rows are empty.")
    if len(data) * cols > MAX_WRITE_CELLS:
        raise OpError(
            f"data has {len(data) * cols:,} cells; at most {MAX_WRITE_CELLS:,} per call. "
            "Split it into several calls."
        )
    return len(data), cols


def write_grid(
    s: Session, sheet: Any, col: int, row: int, data: list[list[Any]], parse_text: bool = False
) -> dict[str, Any]:
    """Write a 2-D list at (col, row). One UNO call for the values plus one per formula column run.

    Rules (documented in the tools): numbers become numbers, text stays text (``"007"`` is
    not converted), a string starting with ``=`` is a formula, ``None`` clears the cell and
    booleans become ``=TRUE()`` / ``=FALSE()``. With ``parse_text`` every string is parsed as if
    typed into the cell ("12" -> number, "2024-01-31" -> date, "50%" -> percentage).
    """
    rows, cols = _validate_grid(data)
    box = (col, row, col + cols - 1, row + rows - 1)
    rng = h.range_of(sheet, box)
    values: list[list[Any]] = []
    formulas: dict[tuple[int, int], str] = {}
    for r, source in enumerate(data):
        out_row: list[Any] = []
        for c in range(cols):
            value = source[c] if c < len(source) else None
            if isinstance(value, bool):
                formulas[(c, r)] = "=TRUE()" if value else "=FALSE()"
                out_row.append("")
            elif isinstance(value, (int, float)):
                out_row.append(float(value))
            elif value is None:
                out_row.append("")
            elif isinstance(value, str):
                if value.startswith("="):
                    formulas[(c, r)] = h.normalize_formula(value)
                    out_row.append("")
                else:
                    out_row.append(value)
            else:
                raise OpError(
                    f"data[{r}][{c}] is a {type(value).__name__}; cell values must be a number, "
                    "string, boolean or null."
                )
        values.append(out_row)
    if parse_text:
        # setFormulaArray parses every string like keyboard input ("12" -> 12, dates...).
        grid = [[_as_input(v) for v in out_row] for out_row in values]
        rng.setFormulaArray(tuple(tuple(r) for r in grid))
        _apply_detected_formats(s, sheet, col, row, grid)
    else:
        rng.setDataArray(tuple(tuple(r) for r in values))
    _write_formula_runs(sheet, col, row, formulas)
    result: dict[str, Any] = {
        "sheet": sheet.Name,
        "range": h.box_name(box),
        "rows": rows,
        "columns": cols,
        "cells_written": rows * cols,
        "formulas": len(formulas),
    }
    if formulas or parse_text:
        errors = h.formula_errors(sheet, rng)
        if errors:
            result["formula_errors"] = errors
            result["warning"] = (
                "Some formulas evaluate to errors (see formula_errors). Check function names "
                "and references; LibreOffice has no XLOOKUP/LET before version 24.8."
            )
    return result


_NEEDS_FORMAT = re.compile(
    r"^\s*(?:[-+]?\d[\d.,]*\s*%"  # percentages
    r"|\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?"  # dates (+ time)
    r"|\d{1,2}:\d{2}(?::\d{2})?)\s*$"  # times
)
MAX_DETECTED_FORMATS = 2000


def _apply_detected_formats(
    s: Session, sheet: Any, col: int, row: int, grid: list[list[str]]
) -> None:
    """Give typed-in dates, times and percentages the number format keyboard input would get.

    ``setFormulaArray`` parses "2024-01-31" into a day number but leaves the cell formatted as
    a plain number; the formatter's ``detectNumberFormat`` returns what typing it would do.
    """
    formatter = s.service("com.sun.star.util.NumberFormatter")
    formatter.attachNumberFormatsSupplier(s.doc)
    found = 0
    for r, grid_row in enumerate(grid):
        for c, text in enumerate(grid_row):
            if not _NEEDS_FORMAT.match(text):
                continue
            key = formatter.detectNumberFormat(0, text)
            if key:
                sheet.getCellByPosition(col + c, row + r).NumberFormat = key
            found += 1
            if found >= MAX_DETECTED_FORMATS:
                return


def _as_input(value: Any) -> str:
    if isinstance(value, float):
        return repr(h.clean_number(value))
    return str(value)


#: Columns with at least this many formulas are written by filling down from the first one.
FILL_THRESHOLD = 200


def _write_formula_runs(
    sheet: Any, col: int, row: int, formulas: dict[tuple[int, int], str]
) -> None:
    """Write formulas with one call per vertical run of formula cells in a column."""
    by_col: dict[int, list[int]] = {}
    for c, r in formulas:
        by_col.setdefault(c, []).append(r)
    for c, rows in by_col.items():
        rows.sort()
        start = prev = rows[0]
        for r in [*rows[1:], None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            _set_formula_run(
                sheet, col + c, row + start, [formulas[(c, i)] for i in range(start, prev + 1)]
            )
            if r is not None:
                start = prev = r


def _set_formula_run(sheet: Any, col: int, row: int, run: list[str]) -> None:
    """Set a vertical run of formulas.

    ``setFormulaArray`` is quadratic in the number of formulas that refer to neighbouring cells
    (20,000 rows of ``=B2*C2`` took 35 s). Filling down from the first cell is instantaneous
    and gives the same relative references, so long runs are filled, verified by reading the
    formulas back, and only the rows that differ are written explicitly.
    """
    last = row + len(run) - 1
    rng = h.range_of(sheet, (col, row, col, last))
    if len(run) < FILL_THRESHOLD:
        rng.setFormulaArray(tuple((f,) for f in run))
        return
    sheet.getCellByPosition(col, row).setFormula(run[0])
    rng.fillAuto(enum("com.sun.star.sheet.FillDirection", "TO_BOTTOM"), 1)
    got = rng.getFormulaArray()
    i = 0
    while i < len(run):
        if h.same_formula(got[i][0], run[i]):
            i += 1
            continue
        j = i
        while j + 1 < len(run) and not h.same_formula(got[j + 1][0], run[j + 1]):
            j += 1
        h.range_of(sheet, (col, row + i, col, row + j)).setFormulaArray(
            tuple((f,) for f in run[i : j + 1])
        )
        i = j + 1


@op("calc.set_range_values", kind="calc", mutates=True)
def set_range_values(
    s: Session,
    data: list[list[Any]],
    sheet: str | None = None,
    start_cell: str = "A1",
    parse_text: bool = False,
) -> dict[str, Any]:
    """Write a 2-D list of values/formulas at start_cell."""
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(start_cell)
    return write_grid(s, sh, col, row, data, parse_text)


@op("calc.set_cell", kind="calc", mutates=True)
def set_cell(
    s: Session, cell: str, value: Any, sheet: str | None = None, parse_text: bool = False
) -> dict[str, Any]:
    """Set one cell and return what it displays."""
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(cell)
    result = write_grid(s, sh, col, row, [[value]], parse_text)
    c = sh.getCellByPosition(col, row)
    result["cell"] = h.cell_name(col, row)
    result["value"] = c.getString()
    return result


# -- bulk table ----------------------------------------------------------------

HEADER_BACKGROUND = "#D9E2F3"


@op("calc.write_table", kind="calc", mutates=True)
def write_table(
    s: Session,
    headers: list[str],
    rows: list[list[Any]],
    sheet: str | None = None,
    start_cell: str = "A1",
    header_style: bool = True,
    autofit: bool = True,
    column_formats: dict[str, str] | None = None,
    autofilter: bool = False,
    freeze_header: bool = False,
    parse_text: bool = False,
) -> dict[str, Any]:
    """Header + rows in one call, with optional header style, formats, filter, freeze."""
    if not headers or not all(isinstance(x, str) for x in headers):
        raise OpError("headers must be a non-empty list of strings (one per column).")
    if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
        raise OpError("rows must be a list of rows, each a list of cell values.")
    for i, r in enumerate(rows):
        if len(r) > len(headers):
            raise OpError(
                f"rows[{i}] has {len(r)} values but there are only {len(headers)} headers."
            )
    sh = h.get_sheet(s, sheet)
    col, row = h.parse_cell(start_cell)
    result = write_grid(s, sh, col, row, [list(headers), *rows], parse_text)
    last_col, last_row = col + len(headers) - 1, row + len(rows)
    if header_style:
        header = h.range_of(sh, (col, row, last_col, row))
        h.set_props(
            header,
            {
                "CharWeight": 150.0,
                "CellBackColor": h.color_value(HEADER_BACKGROUND),
                "IsCellBackgroundTransparent": False,
            },
        )
        _bottom_border(header)
    if column_formats:
        _apply_column_formats(s, sh, col, row, headers, len(rows), column_formats)
    if autofit:
        h.range_of(sh, (col, row, last_col, row)).getColumns().OptimalWidth = True
    if autofilter:
        enable_autofilter(s, sh, (col, row, last_col, last_row))
        result["auto_filter"] = True
    if freeze_header:
        if h.freeze_supported(s):
            h.set_freeze(s.doc, h.sheet_index(sh), 0, row + 1)
            result["frozen_rows"] = row + 1
        else:
            result["warning"] = "freeze_header was not applied. " + h.freeze_unsupported_message(s)
    result["header_row"] = row + 1
    result["data_rows"] = len(rows)
    return result


def _bottom_border(rng: Any) -> None:
    line = BorderLine2()
    line.OuterLineWidth = 0
    line.LineWidth = 26
    line.Color = 0x808080
    rng.BottomBorder = line


def _apply_column_formats(
    s: Session, sheet: Any, col: int, row: int, headers: list[str], n_rows: int, formats: dict
) -> None:
    if n_rows == 0:
        return
    for key, code in formats.items():
        if key in headers:
            c = col + headers.index(key)
        else:
            try:
                c = h.column_letter_or_index(key)
            except OpError:
                c = -1
            if not col <= c < col + len(headers):
                raise OpError(
                    f"column_formats key {key!r} is neither a header ({', '.join(headers)}) "
                    "nor a column letter inside the table."
                ) from None
        h.range_of(sheet, (c, row + 1, c, row + n_rows)).NumberFormat = h.number_format_key(s, code)


# -- clear / fill / copy -------------------------------------------------------


@op("calc.clear_range", kind="calc", mutates=True)
def clear_range(
    s: Session, range: str, sheet: str | None = None, what: str = "contents"
) -> dict[str, Any]:
    """Clear contents, formats or both (charts and images are untouched)."""
    sh, rng, box = h.get_range(s, sheet, range, whole=True)
    flags = {"contents": _CONTENTS, "formats": _FORMATS, "all": _CONTENTS | _FORMATS | _ANNOTATION}
    if what not in flags:
        raise OpError("what must be 'contents', 'formats' or 'all'.")
    if what == "contents" and rng.queryContentCells(_CONTENTS).getCount() == 0:
        return {"sheet": sh.Name, "range": h.box_name(box), "cleared": what, "_noop": True}
    rng.clearContents(flags[what])
    return {"sheet": sh.Name, "range": h.box_name(box), "cleared": what}


_DIRECTIONS = {"down": "TO_BOTTOM", "right": "TO_RIGHT", "up": "TO_TOP", "left": "TO_LEFT"}


@op("calc.fill_range", kind="calc", mutates=True)
def fill_range(
    s: Session,
    range: str,
    sheet: str | None = None,
    direction: str = "down",
    mode: str = "autofill",
    step: float = 1.0,
) -> dict[str, Any]:
    """Fill a range from its first (or last) row/column: autofill, copy or series."""
    sh, rng, box = h.get_range(s, sheet, range)
    if direction not in _DIRECTIONS:
        raise OpError("direction must be 'down', 'right', 'up' or 'left'.")
    c1, r1, c2, r2 = box
    vertical = direction in ("down", "up")
    if (r2 - r1 if vertical else c2 - c1) < 1:
        raise OpError(
            f"Range {h.box_name(box)} has nothing to fill into. Give the source cell(s) plus the "
            f"cells to fill, e.g. 'C2:C20' to fill {direction} from C2."
        )
    h.check_cell_budget(box, "fill range")
    if mode == "copy":
        _copy_fill(s, sh, box, direction)
    elif mode == "autofill":
        rng.fillAuto(enum("com.sun.star.sheet.FillDirection", _DIRECTIONS[direction]), 1)
    elif mode in ("linear", "growth"):
        rng.fillSeries(
            enum("com.sun.star.sheet.FillDirection", _DIRECTIONS[direction]),
            enum("com.sun.star.sheet.FillMode", mode.upper()),
            enum("com.sun.star.sheet.FillDateMode", "FILL_DATE_DAY"),
            float(step),
            1e308,
        )
    else:
        raise OpError("mode must be 'autofill', 'copy', 'linear' or 'growth'.")
    result: dict[str, Any] = {"sheet": sh.Name, "range": h.box_name(box), "mode": mode}
    errors = h.formula_errors(sh, rng)
    if errors:
        result["formula_errors"] = errors
    return result


def _copy_fill(s: Session, sheet: Any, box: h.Box, direction: str) -> None:
    """Repeat the first (or, for up/left, last) row/column over the whole range.

    ``copyRange`` copies one block; doubling the filled part each step needs only
    log2(n) calls and keeps relative references right (``=A1*2`` becomes ``=A2*2`` ...).
    """
    c1, r1, c2, r2 = box
    vertical = direction in ("down", "up")
    total = (r2 - r1 + 1) if vertical else (c2 - c1 + 1)
    filled = 1
    while filled < total:
        chunk = min(filled, total - filled)
        if direction == "down":
            src, dest = (c1, r1, c2, r1 + chunk - 1), (c1, r1 + filled)
        elif direction == "up":
            src, dest = (c1, r2 - chunk + 1, c2, r2), (c1, r2 - filled - chunk + 1)
        elif direction == "right":
            src, dest = (c1, r1, c1 + chunk - 1, r2), (c1 + filled, r1)
        else:
            src, dest = (c2 - chunk + 1, r1, c2, r2), (c2 - filled - chunk + 1, r1)
        sheet.copyRange(h.cell_address(sheet, *dest), h.range_of(sheet, src).getRangeAddress())
        filled += chunk


@op("calc.copy_range", kind="calc", mutates=True)
def copy_range(
    s: Session,
    source_range: str,
    target_cell: str,
    sheet: str | None = None,
    target_sheet: str | None = None,
    move: bool = False,
) -> dict[str, Any]:
    """Copy or move a block of cells, possibly to another sheet."""
    src_sheet, src, box = h.get_range(s, sheet, source_range)
    dst_sheet = h.get_sheet(s, target_sheet) if target_sheet else src_sheet
    col, row = h.parse_cell(target_cell)
    c1, r1, c2, r2 = box
    target_box = (col, row, col + (c2 - c1), row + (r2 - r1))
    h.range_of(dst_sheet, target_box)  # bounds check
    dest = h.cell_address(dst_sheet, col, row)
    if move:
        src_sheet.moveRange(dest, src.getRangeAddress())
    else:
        src_sheet.copyRange(dest, src.getRangeAddress())
    return {
        "from": f"{src_sheet.Name}!{h.box_name(box)}",
        "to": f"{dst_sheet.Name}!{h.box_name(target_box)}",
        "moved": move,
    }


# -- find & replace, sort ------------------------------------------------------


@op("calc.replace_text", kind="calc", mutates=True)
def replace_text(
    s: Session,
    search: str,
    replace: str,
    sheet: str | None = None,
    range: str | None = None,
    regex: bool = False,
    match_case: bool = False,
    whole_cell: bool = False,
    search_in: str = "values",
) -> dict[str, Any]:
    """Find & replace across cells (LibreOffice's own replace, so regex works)."""
    if not search:
        raise OpError("search must not be empty.")
    if search_in not in ("values", "formulas"):
        raise OpError("search_in must be 'values' or 'formulas'.")
    sheets = h.sheets_for(s, sheet)
    total = 0
    for sh in sheets:
        target = sh
        if range:
            if not sheet:
                raise OpError("Give 'sheet' together with 'range'.")
            target = h.range_of(sh, h.parse_range(range, whole=True))
        descriptor = target.createReplaceDescriptor()
        descriptor.SearchString = search
        descriptor.ReplaceString = replace
        descriptor.SearchRegularExpression = regex
        descriptor.SearchCaseSensitive = match_case
        descriptor.SearchWords = whole_cell
        # SearchType: 0 = formulas (the raw cell content), 1 = values (what is displayed).
        descriptor.SearchType = 1 if search_in == "values" else 0
        found = target.findAll(descriptor)
        count = (
            sum(h.box_cells(h.box_of_address(a)) for a in found.getRangeAddresses()) if found else 0
        )
        if count:
            target.replaceAll(descriptor)
            total += count
    result: dict[str, Any] = {"replaced_cells": total}
    if total == 0:
        result["_noop"] = True
    return result


@op("calc.sort_range", kind="calc", mutates=True)
def sort_range(
    s: Session,
    range: str,
    sort_by: list[dict[str, Any]],
    sheet: str | None = None,
    has_header: bool = True,
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Sort rows of a range by up to three keys (column letter or header name)."""
    sh, rng, box = h.get_range(s, sheet, range)
    c1, r1, c2, r2 = box
    if not sort_by:
        raise OpError("sort_by needs at least one key, e.g. [{'column': 'B', 'ascending': false}].")
    if r1 == r2:
        raise OpError("The range has only one row; there is nothing to sort.")
    header = (
        [str(v) for v in h.range_of(sh, (c1, r1, c2, r1)).getDataArray()[0]] if has_header else []
    )
    fields = []
    for key in sort_by:
        column = key.get("column")
        if column is None:
            raise OpError("Each sort_by item needs a 'column' (letter like 'B' or a header name).")
        if str(column) in header:
            index = c1 + header.index(str(column))
        else:
            try:
                index = h.column_letter_or_index(column)
            except OpError:
                raise OpError(
                    f"Sort column {column!r} is not a column letter or a header in "
                    f"{h.box_name(box)} (headers: {', '.join(header) or 'none - has_header is false'})."
                ) from None
        if not c1 <= index <= c2:
            raise OpError(
                f"Sort column {h.col_name(index)} is outside the range {h.box_name(box)} "
                f"(columns {h.col_name(c1)}..{h.col_name(c2)})."
            )
        field = TableSortField()
        field.Field = index - c1  # relative to the range
        field.IsAscending = bool(key.get("ascending", True))
        field.IsCaseSensitive = case_sensitive
        fields.append(field)
    descriptor = {p.Name: p for p in rng.createSortDescriptor()}
    descriptor["SortFields"].Value = uno.Any("[]com.sun.star.table.TableSortField", tuple(fields))
    descriptor["ContainsHeader"].Value = has_header
    descriptor["IsSortColumns"].Value = False
    rng.sort(tuple(descriptor.values()))
    return {"sheet": sh.Name, "range": h.box_name(box), "sorted_by": len(fields)}


# -- rows / columns / merging --------------------------------------------------


@op("calc.insert_rows", kind="calc", mutates=True)
def insert_rows(
    s: Session, before_row: int, count: int = 1, sheet: str | None = None
) -> dict[str, Any]:
    """Insert empty rows above a 1-based row number."""
    sh = h.get_sheet(s, sheet)
    _check_row(before_row, count, inserting=True)
    sh.getRows().insertByIndex(before_row - 1, count)
    return {"sheet": sh.Name, "inserted_rows": count, "first_new_row": before_row}


@op("calc.delete_rows", kind="calc", mutates=True)
def delete_rows(
    s: Session, start_row: int, count: int = 1, sheet: str | None = None
) -> dict[str, Any]:
    """Delete rows starting at a 1-based row number."""
    sh = h.get_sheet(s, sheet)
    _check_row(start_row, count)
    sh.getRows().removeByIndex(start_row - 1, count)
    return {"sheet": sh.Name, "deleted_rows": count, "start_row": start_row}


@op("calc.insert_columns", kind="calc", mutates=True)
def insert_columns(
    s: Session, before_column: str, count: int = 1, sheet: str | None = None
) -> dict[str, Any]:
    """Insert empty columns to the left of a column letter."""
    sh = h.get_sheet(s, sheet)
    index = h.column_letter_or_index(before_column)
    _check_col(index, count, inserting=True)
    sh.getColumns().insertByIndex(index, count)
    return {"sheet": sh.Name, "inserted_columns": count, "first_new_column": h.col_name(index)}


@op("calc.delete_columns", kind="calc", mutates=True)
def delete_columns(
    s: Session, start_column: str, count: int = 1, sheet: str | None = None
) -> dict[str, Any]:
    """Delete columns starting at a column letter."""
    sh = h.get_sheet(s, sheet)
    index = h.column_letter_or_index(start_column)
    _check_col(index, count)
    sh.getColumns().removeByIndex(index, count)
    return {"sheet": sh.Name, "deleted_columns": count, "start_column": h.col_name(index)}


def _check_row(first: int, count: int, inserting: bool = False) -> None:
    if first < 1 or count < 1 or first - 1 + count > h.MAX_ROW + 1:
        raise OpError(
            f"Rows {first}..{first + count - 1} are outside the sheet (1..{h.MAX_ROW + 1})."
        )


def _check_col(index: int, count: int, inserting: bool = False) -> None:
    if count < 1 or index + count - 1 > h.MAX_COL:
        raise OpError(
            f"Columns {h.col_name(index)}.. ({count}) are outside the sheet (A..{h.col_name(h.MAX_COL)})."
        )


@op("calc.merge_cells", kind="calc", mutates=True)
@h.warn_in_csv
def merge_cells(s: Session, range: str, sheet: str | None = None) -> dict[str, Any]:
    """Merge a range; only the top-left cell's content survives."""
    sh, rng, box = h.get_range(s, sheet, range)
    if h.box_cells(box) < 2:
        raise OpError("Select at least two cells to merge, e.g. 'A1:C1'.")
    if rng.getIsMerged():
        return {"sheet": sh.Name, "range": h.box_name(box), "merged": True, "_noop": True}
    others = sum(
        1
        for r, row in enumerate(rng.getDataArray())
        for c, v in enumerate(row)
        if (r, c) != (0, 0) and v != ""
    )
    rng.merge(True)
    result: dict[str, Any] = {"sheet": sh.Name, "range": h.box_name(box), "merged": True}
    if others:
        result["warning"] = (
            f"{others} non-empty cell(s) other than the top-left one were discarded by the merge."
        )
    return result


@op("calc.unmerge_cells", kind="calc", mutates=True)
def unmerge_cells(s: Session, range: str, sheet: str | None = None) -> dict[str, Any]:
    """Split merged cells back into single cells."""
    sh, rng, box = h.get_range(s, sheet, range)
    if not rng.getIsMerged():
        # a range that merely contains merged areas reports False too: check the top-left cell
        cursor = sh.createCursorByRange(sh.getCellByPosition(box[0], box[1]))
        cursor.collapseToMergedArea()
        if h.box_cells(h.box_of(cursor)) == 1:
            return {"sheet": sh.Name, "range": h.box_name(box), "unmerged": False, "_noop": True}
    rng.merge(False)
    return {"sheet": sh.Name, "range": h.box_name(box), "unmerged": True}
