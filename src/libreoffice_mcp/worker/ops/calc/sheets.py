"""Calc: sheets, named ranges, freeze panes, auto-filter, recalculation."""

from __future__ import annotations

import re
from typing import Any

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.ops.calc import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session

# Characters Calc refuses in sheet names.
_BAD_SHEET_CHARS = re.compile(r"[\[\]*?:/\\]")


def _check_sheet_name(name: str) -> str:
    if not name or not name.strip():
        raise OpError("The sheet name must not be empty.")
    if len(name) > 31:
        raise OpError(f"Sheet name {name!r} is longer than 31 characters (Excel's limit).")
    if _BAD_SHEET_CHARS.search(name) or name.startswith("'") or name.endswith("'"):
        raise OpError(
            f"Invalid sheet name {name!r}: it may not contain [ ] * ? : / \\ or start/end with "
            "an apostrophe."
        )
    return name


def _check_new_name(s: Session, name: str) -> None:
    _check_sheet_name(name)
    if any(n.lower() == name.lower() for n in h.sheet_names(s)):
        raise OpError(
            f"A sheet named {name!r} already exists. Sheets: "
            f"{', '.join(repr(n) for n in h.sheet_names(s))}."
        )


def _clamp_position(s: Session, position: int | None) -> int:
    count = s.doc.getSheets().getCount()
    if position is None:
        return count
    if not 0 <= position <= count:
        raise OpError(f"position {position} is out of range (0..{count}); 0 = first sheet.")
    return position


@op("calc.add_sheet", kind="calc", mutates=True)
def add_sheet(s: Session, name: str, position: int | None = None) -> dict[str, Any]:
    """Add an empty sheet."""
    h.require_not_csv(s, "Adding a sheet")
    _check_new_name(s, name)
    index = _clamp_position(s, position)
    s.doc.getSheets().insertNewByName(name, index)
    return {"sheet": name, "position": index, "sheets": h.sheet_names(s)}


@op("calc.rename_sheet", kind="calc", mutates=True)
def rename_sheet(s: Session, sheet: str, new_name: str) -> dict[str, Any]:
    """Rename a sheet (LibreOffice updates formulas that refer to it)."""
    sh = h.get_sheet(s, sheet)
    if sh.Name == new_name:
        return {"sheet": new_name, "_noop": True}
    _check_new_name(s, new_name)
    old = sh.Name
    sh.Name = new_name  # formulas referring to the sheet are updated by LibreOffice
    return {"sheet": new_name, "previous_name": old}


@op("calc.delete_sheet", kind="calc", mutates=True)
def delete_sheet(s: Session, sheet: str) -> dict[str, Any]:
    """Delete a sheet; a spreadsheet must keep at least one."""
    sh = h.get_sheet(s, sheet)
    sheets = s.doc.getSheets()
    if sheets.getCount() == 1:
        raise OpError("A spreadsheet must keep at least one sheet; add another sheet first.")
    name = sh.Name
    sheets.removeByName(name)
    return {"deleted_sheet": name, "sheets": h.sheet_names(s)}


@op("calc.move_sheet", kind="calc", mutates=True)
def move_sheet(s: Session, sheet: str, position: int) -> dict[str, Any]:
    """Move a sheet to a 0-based tab position."""
    sh = h.get_sheet(s, sheet)
    count = s.doc.getSheets().getCount()
    if not 0 <= position < count:
        raise OpError(f"position {position} is out of range (0..{count - 1}); 0 = first sheet.")
    if h.sheet_index(sh) == position:
        return {"sheet": sh.Name, "position": position, "_noop": True}
    # moveByName(name, n) inserts *before* index n of the current order; n == count appends.
    current = h.sheet_index(sh)
    s.doc.getSheets().moveByName(sh.Name, position + 1 if position > current else position)
    return {"sheet": sh.Name, "position": position, "sheets": h.sheet_names(s)}


@op("calc.copy_sheet", kind="calc", mutates=True)
def copy_sheet(
    s: Session, sheet: str, new_name: str, position: int | None = None
) -> dict[str, Any]:
    """Duplicate a sheet under a new name."""
    h.require_not_csv(s, "Copying a sheet")
    sh = h.get_sheet(s, sheet)
    _check_new_name(s, new_name)
    index = _clamp_position(s, position)
    s.doc.getSheets().copyByName(sh.Name, new_name, index)
    return {"sheet": new_name, "copied_from": sh.Name, "sheets": h.sheet_names(s)}


@op("calc.set_sheet_properties", kind="calc", mutates=True)
@h.warn_in_csv
def set_sheet_properties(
    s: Session, sheet: str, visible: bool | None = None, tab_color: str | None = None
) -> dict[str, Any]:
    """Show/hide a sheet and set its tab colour."""
    sh = h.get_sheet(s, sheet)
    changed = []
    if visible is not None and bool(sh.IsVisible) != visible:
        if not visible:
            shown = [n for n in h.sheet_names(s) if s.doc.getSheets().getByName(n).IsVisible]
            if shown == [sh.Name]:
                raise OpError("The last visible sheet cannot be hidden; show another sheet first.")
        sh.IsVisible = visible
        changed.append("visible")
    if tab_color is not None:
        value = -1 if tab_color.lower() in ("none", "") else h.color_value(tab_color)
        if sh.TabColor != value:
            sh.TabColor = value
            changed.append("tab_color")
    if not changed:
        return {"sheet": sh.Name, "changed": [], "_noop": True}
    return {"sheet": sh.Name, "changed": changed}


# -- named ranges --------------------------------------------------------------

_NAME_OK = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_LOOKS_LIKE_CELL = re.compile(r"^[A-Za-z]{1,3}\d+$")


def _absolute_ref(sheet_name: str, box: h.Box) -> str:
    quoted = (
        "'" + sheet_name.replace("'", "''") + "'" if re.search(r"\W", sheet_name) else sheet_name
    )
    c1, r1, c2, r2 = box
    first = f"${h.col_name(c1)}${r1 + 1}"
    if (c1, r1) == (c2, r2):
        return f"${quoted}.{first}"
    return f"${quoted}.{first}:${h.col_name(c2)}${r2 + 1}"


@op("calc.add_named_range", kind="calc", mutates=True)
@h.warn_in_csv
def add_named_range(s: Session, name: str, range: str, sheet: str | None = None) -> dict[str, Any]:
    """Define a name for a range."""
    if not _NAME_OK.match(name) or _LOOKS_LIKE_CELL.match(name):
        raise OpError(
            f"Invalid name {name!r}: use letters, digits, '_' or '.', start with a letter or "
            "'_', and do not look like a cell address (e.g. 'A1')."
        )
    named = s.doc.NamedRanges
    if named.hasByName(name):
        raise OpError(
            f"A named range called {name!r} already exists. Delete it first with "
            "calc_delete_named_range or pick another name."
        )
    sh, _, box = h.get_range(s, sheet, range)
    named.addNewByName(name, _absolute_ref(sh.Name, box), h.cell_address(sh, 0, 0), 0)
    return {"name": name, "refers_to": f"{sh.Name}!{h.box_name(box)}"}


@op("calc.delete_named_range", kind="calc", mutates=True)
def delete_named_range(s: Session, name: str) -> dict[str, Any]:
    """Remove a defined name."""
    named = s.doc.NamedRanges
    if not named.hasByName(name):
        raise OpError(
            f"No named range {name!r}. Existing: {', '.join(named.getElementNames()) or 'none'}.",
            protocol.E_NOT_FOUND,
        )
    named.removeByName(name)
    return {"deleted": name}


# -- freeze panes / recalculation ----------------------------------------------


@op("calc.freeze_panes", kind="calc", mutates=True)
def freeze_panes(
    s: Session, sheet: str | None = None, rows: int = 0, columns: int = 0
) -> dict[str, Any]:
    """Freeze top rows / left columns (xlsx/xls only; see helpers)."""
    sh = h.get_sheet(s, sheet)
    if not (0 <= rows <= 1000 and 0 <= columns <= 100):
        raise OpError("rows must be 0..1000 and columns 0..100 (0 = do not freeze).")
    index = h.sheet_index(sh)
    if (rows or columns) and not h.freeze_supported(s):
        raise OpError(h.freeze_unsupported_message(s))
    if h.get_freeze(s.doc, index) == (columns, rows):
        return {"sheet": sh.Name, "frozen_rows": rows, "frozen_columns": columns, "_noop": True}
    h.set_freeze(s.doc, index, columns, rows)
    return {"sheet": sh.Name, "frozen_rows": rows, "frozen_columns": columns}


@op("calc.recalculate", kind="calc", mutates=True)
def recalculate(s: Session) -> dict[str, Any]:
    """Hard-recalculate every formula."""
    s.doc.calculateAll()
    return {"recalculated": True}


# -- auto-filter ---------------------------------------------------------------

_FILTER_PREFIX = "mcp_filter_"


def _filter_ranges(s: Session, sheet: Any) -> list[Any]:
    index = h.sheet_index(sheet)
    databases = s.doc.DatabaseRanges
    found = []
    for name in databases.getElementNames():
        db = databases.getByName(name)
        if db.getDataArea().Sheet == index and db.AutoFilter:
            found.append((name, db))
    return found


def enable_autofilter(s: Session, sheet: Any, box: h.Box) -> None:
    """Turn on filter buttons for ``box`` (first row = headers); replaces an existing filter."""
    area = h.range_of(sheet, box).getRangeAddress()
    existing = _filter_ranges(s, sheet)
    if existing:
        _, db = existing[0]
        db.setDataArea(area)
        db.ContainsHeader = True
        return
    databases = s.doc.DatabaseRanges
    n = 1
    while databases.hasByName(f"{_FILTER_PREFIX}{n}"):
        n += 1
    databases.addNewByName(f"{_FILTER_PREFIX}{n}", area)
    db = databases.getByName(f"{_FILTER_PREFIX}{n}")
    db.ContainsHeader = True
    db.AutoFilter = True


@op("calc.set_autofilter", kind="calc", mutates=True)
@h.warn_in_csv
def set_autofilter(
    s: Session, range: str | None = None, sheet: str | None = None, enabled: bool = True
) -> dict[str, Any]:
    """Show or remove filter buttons on a table's header row."""
    sh = h.get_sheet(s, sheet)
    if not enabled:
        existing = _filter_ranges(s, sh)
        if not existing:
            return {"sheet": sh.Name, "auto_filter": None, "_noop": True}
        for name, db in existing:
            db.AutoFilter = False
            if name.startswith(_FILTER_PREFIX):
                s.doc.DatabaseRanges.removeByName(name)
        return {"sheet": sh.Name, "auto_filter": None}
    if range is None:
        box = h.used_box(sh)
        if box is None:
            raise OpError("The sheet is empty; there is nothing to filter. Give a range.")
    else:
        box = h.parse_range(range)
    if box[3] == box[1]:
        raise OpError(
            "An auto-filter needs a header row plus at least one data row; "
            f"{h.box_name(box)} has only one row."
        )
    enable_autofilter(s, sh, box)
    return {"sheet": sh.Name, "auto_filter": h.box_name(box)}
