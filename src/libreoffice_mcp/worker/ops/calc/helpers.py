"""Shared building blocks for the Calc operations.

Vocabulary used throughout:

* **sheet**  - addressed by name (``sheet`` argument; ``None`` means the first sheet).
* **range**  - A1 notation: ``B2``, ``A1:C10``; whole columns ``A:C`` and rows ``2:5`` are
  accepted where a tool says so. Columns and rows are 0-based *inside* the worker and 1-based
  (rows) / letters (columns) everywhere the AI sees them.
"""

from __future__ import annotations

import functools
import re
from collections.abc import Callable
from typing import Any

import uno

from libreoffice_mcp import formats, protocol
from libreoffice_mcp.worker.registry import OpError
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color

MM100 = 100  # UNO lengths are 1/100 mm
MAX_COL = 16383  # 16384 columns (A..XFD) in LibreOffice >= 7.4 and in .xlsx
MAX_ROW = 1048575  # 1,048,576 rows
#: Largest number of cells one tool call will return or write in a single UNO round trip.
MAX_CELLS = 10_000

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW = re.compile(r"^\$?(\d+)$")

Box = tuple[int, int, int, int]  # (col1, row1, col2, row2), 0-based, inclusive


# -- A1 notation ---------------------------------------------------------------


def span(first: int, last: int) -> range:
    """Inclusive range (the tool argument called ``range`` shadows the builtin in some ops)."""
    return range(first, last + 1)


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def col_name(index: int) -> str:
    letters = ""
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_name(col: int, row: int) -> str:
    return f"{col_name(col)}{row + 1}"


def box_name(box: Box) -> str:
    c1, r1, c2, r2 = box
    if (c1, r1) == (c2, r2):
        return cell_name(c1, r1)
    return f"{cell_name(c1, r1)}:{cell_name(c2, r2)}"


def _bad_a1(text: str, hint: str = "") -> OpError:
    return OpError(
        f"Invalid cell range {text!r}. Use A1 notation such as 'B2' or 'A1:C10'"
        f"{hint}. Give the sheet name in the separate 'sheet' argument."
    )


def _parse_cell(text: str, original: str) -> tuple[int, int]:
    m = _CELL.match(text.strip())
    if not m:
        raise _bad_a1(original)
    col, row = col_index(m.group(1)), int(m.group(2)) - 1
    if row < 0 or row > MAX_ROW or col > MAX_COL:
        raise OpError(
            f"{text!r} is outside the sheet (columns A..{col_name(MAX_COL)}, rows 1..{MAX_ROW + 1})."
        )
    return col, row


def parse_cell(text: str) -> tuple[int, int]:
    """'C3' -> (2, 2)."""
    if ":" in text:
        raise OpError(f"Expected a single cell like 'C3', got the range {text!r}.")
    return _parse_cell(text, text)


def parse_range(text: str, whole: bool = False) -> Box:
    """'B2', 'A1:C10' (and with ``whole=True`` also 'A:C' / '2:5') -> (col1, row1, col2, row2)."""
    if not isinstance(text, str) or not text.strip():
        raise _bad_a1(text)
    if "!" in text or text.count(":") > 1:
        raise _bad_a1(text)
    first, _, last = text.strip().partition(":")
    if not last:
        c, r = _parse_cell(first, text)
        return c, r, c, r
    if whole:
        if _COL.match(first) and _COL.match(last):
            c1, c2 = col_index(first.lstrip("$")), col_index(last.lstrip("$"))
            if max(c1, c2) > MAX_COL:
                raise _bad_a1(text)
            return min(c1, c2), 0, max(c1, c2), MAX_ROW
        if _ROW.match(first) and _ROW.match(last):
            r1, r2 = int(first.lstrip("$")) - 1, int(last.lstrip("$")) - 1
            if min(r1, r2) < 0 or max(r1, r2) > MAX_ROW:
                raise _bad_a1(text)
            return 0, min(r1, r2), MAX_COL, max(r1, r2)
    c1, r1 = _parse_cell(first, text)
    c2, r2 = _parse_cell(last, text)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def parse_columns(text: str) -> tuple[int, int]:
    """'B', 'A:C' -> (first, last) 0-based column indexes."""
    first, _, last = str(text).strip().partition(":")
    if not _COL.match(first) or (last and not _COL.match(last)):
        raise OpError(f"Invalid column spec {text!r}. Use a letter like 'B' or a range like 'A:C'.")
    c1 = col_index(first.lstrip("$"))
    c2 = col_index(last.lstrip("$")) if last else c1
    if max(c1, c2) > MAX_COL:
        raise OpError(f"Column {text!r} is beyond the last column {col_name(MAX_COL)}.")
    return min(c1, c2), max(c1, c2)


def parse_rows(text: str) -> tuple[int, int]:
    """'5', '2:9' (1-based) -> (first, last) 0-based row indexes."""
    first, _, last = str(text).strip().partition(":")
    if not _ROW.match(first) or (last and not _ROW.match(last)):
        raise OpError(f"Invalid row spec {text!r}. Use a number like '5' or a range like '2:9'.")
    r1 = int(first.lstrip("$")) - 1
    r2 = int(last.lstrip("$")) - 1 if last else r1
    if min(r1, r2) < 0 or max(r1, r2) > MAX_ROW:
        raise OpError(f"Row {text!r} is outside the sheet (rows 1..{MAX_ROW + 1}).")
    return min(r1, r2), max(r1, r2)


def column_letter_or_index(value: Any) -> int:
    """Column given as a letter ('C') or a 1-based number (3) -> 0-based index."""
    if isinstance(value, int) and not isinstance(value, bool):
        if not 1 <= value <= MAX_COL + 1:
            raise OpError(f"Column number {value} is out of range (1..{MAX_COL + 1}).")
        return value - 1
    text = str(value).strip()
    if text.isdigit():
        return column_letter_or_index(int(text))
    return parse_columns(text)[0]


# -- sheets and ranges ---------------------------------------------------------


def sheet_names(s: Session) -> list[str]:
    return list(s.doc.getSheets().getElementNames())


def sheet_index(sheet: Any) -> int:
    return sheet.getRangeAddress().Sheet


def get_sheet(s: Session, name: str | None) -> Any:
    """Sheet by name (exact, then case-insensitive); ``None`` = the first sheet."""
    sheets = s.doc.getSheets()
    if name is None or name == "":
        return sheets.getByIndex(0)
    if sheets.hasByName(name):
        return sheets.getByName(name)
    names = list(sheets.getElementNames())
    lowered = [n for n in names if n.lower() == name.lower()]
    if len(lowered) == 1:
        return sheets.getByName(lowered[0])
    raise OpError(
        f"No sheet named {name!r}. Sheets in this file: {', '.join(repr(n) for n in names)}. "
        "Use calc_list_sheets to see them.",
        protocol.E_NOT_FOUND,
    )


def sheets_for(s: Session, name: str | None) -> list[Any]:
    """The named sheet, or every sheet when ``name`` is None."""
    if name:
        return [get_sheet(s, name)]
    sheets = s.doc.getSheets()
    return [sheets.getByIndex(i) for i in range(sheets.getCount())]


def get_range(s: Session, sheet_name: str | None, text: str, whole: bool = False):
    """Resolve ``(sheet, range)`` -> ``(sheet_obj, range_obj, box)`` with helpful errors."""
    sheet = get_sheet(s, sheet_name)
    box = parse_range(text, whole=whole)
    return sheet, range_of(sheet, box), box


def range_of(sheet: Any, box: Box) -> Any:
    c1, r1, c2, r2 = box
    try:
        return sheet.getCellRangeByPosition(c1, r1, c2, r2)
    except Exception as exc:  # IndexOutOfBoundsException
        raise OpError(
            f"Range {box_name(box)} is outside the sheet (columns A..{col_name(MAX_COL)}, "
            f"rows 1..{MAX_ROW + 1})."
        ) from exc


def box_of(rng: Any) -> Box:
    a = rng.getRangeAddress()
    return a.StartColumn, a.StartRow, a.EndColumn, a.EndRow


def box_of_address(a: Any) -> Box:
    return a.StartColumn, a.StartRow, a.EndColumn, a.EndRow


def box_cells(box: Box) -> int:
    return (box[2] - box[0] + 1) * (box[3] - box[1] + 1)


def check_cell_budget(box: Box, what: str = "range") -> None:
    if box_cells(box) > MAX_CELLS * 10:
        raise OpError(
            f"The {what} {box_name(box)} covers {box_cells(box):,} cells, which is too large for "
            "one call. Use a smaller range (for example the used range from calc_list_sheets)."
        )


def used_box(sheet: Any) -> Box | None:
    """Bounding box of cells that hold content; ``None`` for an empty sheet."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end = cursor.getRangeAddress()
    cursor.gotoStartOfUsedArea(False)
    start = cursor.getRangeAddress()
    box = (start.StartColumn, start.StartRow, end.EndColumn, end.EndRow)
    if box[2] < box[0] or box[3] < box[1]:  # a sheet holding only pictures reports an empty box
        return None
    if box == (0, 0, 0, 0) and sheet.getCellByPosition(0, 0).getType().value == "EMPTY":
        return None
    return box


# -- values --------------------------------------------------------------------


def clean_number(value: float) -> int | float:
    """UNO returns every number as a float; show whole numbers as ints."""
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return int(value)
    return value


def normalize_formula(formula: str) -> str:
    """Standard spreadsheet formula -> the syntax ``setFormula`` understands.

    ``XCell.setFormula`` uses LibreOffice's API grammar: ``;`` separates arguments and a
    sheet reference is ``$Sheet.A1``. AI models naturally write ``=SUM(A1,B1)`` and
    ``=Sheet2!A1``, so we translate those (outside of string literals) and leave everything
    else alone. Formulas with array constants (``{``) are passed through unchanged.
    """
    if "{" in formula:
        return formula
    parts = re.split(r'("(?:[^"]|"")*")', formula)  # odd items are "string literals"
    out = []
    for i, part in enumerate(parts):
        if i % 2 == 0:
            part = re.sub(r"((?:'(?:[^']|'')+')|(?:[A-Za-z_][A-Za-z0-9_.]*))!(?!=)", r"$\1.", part)
            part = _commas_to_semicolons(part)
        out.append(part)
    return "".join(out)


def same_formula(read_back: str, wanted: str) -> bool:
    """Is a formula read back from a cell the one we asked for?

    LibreOffice normalises function-name case and spacing, so those are ignored; text inside
    string literals must match exactly.
    """
    a = re.split(r'("(?:[^"]|"")*")', read_back)
    b = re.split(r'("(?:[^"]|"")*")', wanted)
    if len(a) != len(b):
        return False
    for i, (x, y) in enumerate(zip(a, b, strict=True)):
        if i % 2:
            if x != y:
                return False
        elif x.replace(" ", "").upper() != y.replace(" ", "").upper():
            return False
    return True


def _commas_to_semicolons(text: str) -> str:
    """Replace commas that are not inside a quoted sheet name ('...')."""
    pieces = re.split(r"('(?:[^']|'')+')", text)
    return "".join(p if i % 2 else p.replace(",", ";") for i, p in enumerate(pieces))


def formula_errors(sheet: Any, rng: Any, limit: int = 20) -> list[dict[str, str]]:
    """Formula cells in ``rng`` that evaluate to an error, e.g. ``{"cell": "C3", "error": "#NAME?"}``.

    One ``queryFormulaCells`` round trip finds the bad cells so callers can warn the AI right
    away instead of letting it discover ``#NAME?`` later.
    """
    result = rng.queryFormulaCells(4)  # com.sun.star.sheet.FormulaResult.ERROR
    errors: list[dict[str, str]] = []
    for address in result.getRangeAddresses():
        for row in range(address.StartRow, address.EndRow + 1):
            for col in range(address.StartColumn, address.EndColumn + 1):
                if len(errors) >= limit:
                    return errors
                cell = sheet.getCellByPosition(col, row)
                errors.append({"cell": cell_name(col, row), "error": cell.getString()})
    return errors


# -- formatting helpers --------------------------------------------------------


def mm(value: float) -> int:
    return round(value * MM100)


def set_props(obj: Any, values: dict[str, Any]) -> None:
    """Set several properties in one round trip (names must be sorted for the multi-setter)."""
    if not values:
        return
    if not hasattr(obj, "setPropertyValues"):  # e.g. TableColumns / TableRows
        for name, value in values.items():
            obj.setPropertyValue(name, value)
        return
    names = sorted(values)
    obj.setPropertyValues(tuple(names), tuple(values[n] for n in names))


def color_value(value: str) -> int:
    return parse_color(value)


def number_format_key(s: Session, code: str) -> int:
    """Key for a number-format code such as ``0.00`` or ``yyyy-mm-dd`` (created if new).

    Codes are interpreted in the en-US locale so the AI can rely on one syntax whatever
    language the host runs in.
    """
    locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    locale.Language, locale.Country = "en", "US"
    formats = s.doc.getNumberFormats()
    key = formats.queryKey(code, locale, False)
    if key == -1:
        try:
            key = formats.addNew(code, locale)
        except Exception as exc:
            raise OpError(
                f"Invalid number format {code!r}. Examples: '0.00', '#,##0', '0%', "
                "'yyyy-mm-dd', '$#,##0.00', '0.0E+00', '@' (text)."
            ) from exc
    return key


def standard_format_key(s: Session, kind: int) -> int:
    """Built-in format key for a ``com.sun.star.util.NumberFormat`` type (e.g. 1024 = boolean)."""
    locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    locale.Language, locale.Country = "en", "US"
    return s.doc.getNumberFormats().getStandardFormat(kind, locale)


def h_align(name: str) -> Any:
    table = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "BLOCK"}
    return enum("com.sun.star.table.CellHoriJustify", table[name])


def v_align(name: str) -> int:
    """VertJustify is a plain long (CellVertJustify2), unlike HoriJustify which is an enum."""
    return {"top": 1, "middle": 2, "bottom": 3}[name]


def cell_address(sheet: Any, col: int, row: int) -> Any:
    address = uno.createUnoStruct("com.sun.star.table.CellAddress")
    address.Sheet, address.Column, address.Row = sheet_index(sheet), col, row
    return address


def cell_at_position(sheet: Any, x: int, y: int) -> tuple[int, int]:
    """The (col, row) whose top-left corner is the last one at or before (x, y), in 1/100 mm.

    Shapes remember where they are, not which cell they were inserted at; a binary search
    over cell positions (about 34 calls) recovers it.
    """

    def search(limit: int, coordinate: str) -> int:
        low, high = 0, limit
        while low < high:
            mid = (low + high + 1) // 2
            cell = (
                sheet.getCellByPosition(mid, 0)
                if coordinate == "X"
                else sheet.getCellByPosition(0, mid)
            )
            if getattr(cell.Position, coordinate) <= (x if coordinate == "X" else y):
                low = mid
            else:
                high = mid - 1
        return low

    return search(MAX_COL, "X"), search(MAX_ROW, "Y")


def cell_position_mm100(sheet: Any, col: int, row: int) -> tuple[int, int]:
    """Top-left corner of a cell on the sheet's drawing layer, in 1/100 mm."""
    pos = sheet.getCellByPosition(col, row).Position
    return pos.X, pos.Y


# -- freeze panes (view settings) ----------------------------------------------
#
# Calc stores "freeze panes" in the *view*, and a hidden document's view has no window size, so
# ``controller.freezeAtPosition`` silently does nothing. The view's settings string can still be
# rewritten with ``restoreViewData`` (this is what loading a file does), and the result is saved
# by the .ods/.xlsx filters. Format (ScViewData::WriteUserData):
#   "<zoom>;<active tab>;tw:<n>;<tab 0>;<tab 1>;..." where every tab is
#   "curX/curY/hMode/vMode/hPos/vPos/activePane/posX0/posX1/posY0/posY1"; mode 2 = frozen,
#   hPos/vPos = number of frozen columns/rows.

_SPLIT_FIXED = "2"
#: Formats whose writer reads the live view. The OpenDocument writer takes view settings from a
#: different source that does not exist without a window, so freeze panes would be lost there.
FREEZE_FORMATS = ("xlsx", "xls")


def freeze_supported(s: Session) -> bool:
    return formats.ext_of(str(s.path)) in FREEZE_FORMATS


def freeze_unsupported_message(s: Session) -> str:
    return (
        f"Freeze panes cannot be saved in .{formats.ext_of(str(s.path))} files by this server "
        "(LibreOffice only writes them for .xlsx/.xls when running without a window). "
        "Use an .xlsx file, or convert with office_convert_document."
    )


def _view_parts(doc: Any) -> tuple[Any, list[str], int]:
    controller = doc.getCurrentController()
    if controller is None:
        raise OpError("This document has no view, so freeze panes cannot be read or changed.")
    parts = controller.getViewData().split(";")
    first_tab = next((i + 1 for i, p in enumerate(parts) if p.startswith("tw:")), 2)
    return controller, parts, first_tab


def get_freeze(doc: Any, sheet_idx: int) -> tuple[int, int]:
    """(frozen columns, frozen rows) of a sheet."""
    try:
        _, parts, first = _view_parts(doc)
        fields = parts[first + sheet_idx].split("/")
        cols = int(fields[4]) if fields[2] == _SPLIT_FIXED else 0
        rows = int(fields[5]) if fields[3] == _SPLIT_FIXED else 0
        return cols, rows
    except (OpError, IndexError, ValueError):
        return 0, 0


def set_freeze(doc: Any, sheet_idx: int, columns: int, rows: int) -> None:
    controller, parts, first = _view_parts(doc)
    if first + sheet_idx >= len(parts):
        raise OpError("Could not update this sheet's freeze panes (unexpected view layout).")
    h_mode = _SPLIT_FIXED if columns else "0"
    v_mode = _SPLIT_FIXED if rows else "0"
    active = 3 if columns and rows else 1 if columns else 2
    parts[first + sheet_idx] = "/".join(
        [
            "0",
            "0",
            h_mode,
            v_mode,
            str(columns),
            str(rows),
            str(active),
            "0",
            str(columns),
            "0",
            str(rows),
        ]
    )
    controller.restoreViewData(";".join(parts))


# -- CSV ------------------------------------------------------------------------
#
# A .csv file is a calc document that stores only the first sheet's cell text. Everything else
# (formatting, charts, images, comments, extra sheets) would silently disappear on save, so
# operations either refuse (extra sheets) or add a warning to their result.


def is_csv(s: Session) -> bool:
    return formats.ext_of(str(s.path)) == "csv"


def require_not_csv(s: Session, what: str) -> None:
    if is_csv(s):
        raise OpError(
            f"{what} is not possible in a .csv file: CSV holds a single sheet of plain text. "
            "Use an .xlsx or .ods file instead (office_convert_document can convert it)."
        )


def warn_in_csv(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator for ops whose effect CSV cannot store: add a ``warning`` to the result."""

    @functools.wraps(fn)
    def wrapper(s: Session, *args: Any, **kwargs: Any) -> Any:
        result = fn(s, *args, **kwargs)
        if is_csv(s) and isinstance(result, dict) and not result.get("_noop"):
            result["warning"] = (
                "This is a .csv file, which stores only cell text: this change will not be "
                "saved. Use an .xlsx or .ods file to keep formatting, charts and comments."
            )
        return result

    return wrapper
