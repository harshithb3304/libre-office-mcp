"""Calc: cell formatting, borders, sizes, conditional formatting and data validation."""

from __future__ import annotations

import re
from typing import Any

from com.sun.star.table import BorderLine2

from libreoffice_mcp.worker.ops.calc import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum

_PT_TO_MM100 = 35.28


# -- cell formatting -----------------------------------------------------------


@op("calc.format_range", kind="calc", mutates=True)
@h.warn_in_csv
def format_range(
    s: Session,
    range: str,
    sheet: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_color: str | None = None,
    background: str | None = None,
    number_format: str | None = None,
    h_align: str | None = None,
    v_align: str | None = None,
    wrap_text: bool | None = None,
    rotation: int | None = None,
    cell_style: str | None = None,
) -> dict[str, Any]:
    """Apply font, colour, number format, alignment and wrapping to a range."""
    sh, rng, box = h.get_range(s, sheet, range, whole=True)
    values: dict[str, Any] = {}
    if font_name is not None:
        values["CharFontName"] = font_name
    if font_size is not None:
        if not 1 <= font_size <= 400:
            raise OpError("font_size must be between 1 and 400 points.")
        values["CharHeight"] = float(font_size)
    if bold is not None:
        values["CharWeight"] = 150.0 if bold else 100.0
    if italic is not None:
        values["CharPosture"] = enum("com.sun.star.awt.FontSlant", "ITALIC" if italic else "NONE")
    if underline is not None:
        values["CharUnderline"] = 1 if underline else 0
    if strikethrough is not None:
        values["CharStrikeout"] = 1 if strikethrough else 0
    if font_color is not None:
        values["CharColor"] = h.color_value(font_color)
    if background is not None:
        if background.lower() in ("none", "transparent"):
            values["IsCellBackgroundTransparent"] = True
        else:
            values["CellBackColor"] = h.color_value(background)
            values["IsCellBackgroundTransparent"] = False
    if number_format is not None:
        values["NumberFormat"] = h.number_format_key(s, number_format)
    if h_align is not None:
        values["HoriJustify"] = h.h_align(h_align)
    if v_align is not None:
        values["VertJustify"] = h.v_align(v_align)
    if wrap_text is not None:
        values["IsTextWrapped"] = wrap_text
    if rotation is not None:
        values["RotateAngle"] = int(rotation) * 100
    if cell_style is not None:
        styles = s.doc.getStyleFamilies().getByName("CellStyles")
        if not styles.hasByName(cell_style):
            raise OpError(
                f"No cell style named {cell_style!r}. Available: "
                f"{', '.join(styles.getElementNames()[:40])}."
            )
        rng.CellStyle = cell_style  # set first: applying a style resets direct formatting
    if not values and cell_style is None:
        raise OpError(
            "Nothing to do: give at least one formatting option (font_size, bold, background, "
            "number_format, h_align, ...)."
        )
    h.set_props(rng, values)
    return {"sheet": sh.Name, "range": h.box_name(box), "applied": sorted([*values, *(["CellStyle"] if cell_style else [])])}  # fmt: skip


_LINE_STYLES = {"solid": 0, "dashed": 1, "dotted": 2, "double": 3}
_SIDES = ("top", "bottom", "left", "right")


def _line(color: int, width_pt: float, style: str) -> BorderLine2:
    line = BorderLine2()
    line.Color = color
    line.LineStyle = _LINE_STYLES[style]
    line.LineWidth = max(1, round(width_pt * _PT_TO_MM100))
    return line


@op("calc.set_borders", kind="calc", mutates=True)
@h.warn_in_csv
def set_borders(
    s: Session,
    range: str,
    sheet: str | None = None,
    sides: list[str] | None = None,
    style: str = "solid",
    width_pt: float = 0.75,
    color: str = "#000000",
) -> dict[str, Any]:
    """Set or remove borders on chosen sides of a range."""
    sh, rng, box = h.get_range(s, sheet, range)
    sides = sides or ["all"]
    allowed = {"all", "outer", "inner", "top", "bottom", "left", "right", "horizontal",
               "vertical", "none"}  # fmt: skip
    bad = [x for x in sides if x not in allowed]
    if bad:
        raise OpError(f"Unknown border side(s) {bad}. Allowed: {', '.join(sorted(allowed))}.")
    if style not in _LINE_STYLES:
        raise OpError(f"style must be one of {', '.join(_LINE_STYLES)}.")
    if not 0.1 <= width_pt <= 10:
        raise OpError("width_pt must be between 0.1 and 10.")
    border = rng.TableBorder2
    none = BorderLine2()
    none.LineStyle = 0x7FFF  # BorderLineStyle.NONE
    chosen = set(sides)
    names = {"top": "Top", "bottom": "Bottom", "left": "Left", "right": "Right",
             "horizontal": "Horizontal", "vertical": "Vertical"}  # fmt: skip
    if "none" in chosen:
        targets, line = set(names), none
    else:
        line = _line(h.color_value(color), width_pt, style)
        if "all" in chosen:
            chosen |= {"outer", "inner"}
        if "outer" in chosen:
            chosen |= set(_SIDES)
        if "inner" in chosen:
            chosen |= {"horizontal", "vertical"}
        targets = chosen & set(names)
    for side, name in names.items():
        # Is*LineValid tells Calc which sides to change; the others keep their border.
        setattr(border, f"Is{name}LineValid", side in targets)
        if side in targets:
            setattr(border, f"{name}Line", line)
    rng.TableBorder2 = border
    return {"sheet": sh.Name, "range": h.box_name(box), "sides": sorted(set(sides))}


# -- sizes ---------------------------------------------------------------------


@op("calc.set_column_width", kind="calc", mutates=True)
@h.warn_in_csv
def set_column_width(
    s: Session,
    columns: str,
    sheet: str | None = None,
    width_mm: float | None = None,
    auto_fit: bool = False,
    hidden: bool | None = None,
) -> dict[str, Any]:
    """Set width, auto-fit or hide/show columns."""
    sh = h.get_sheet(s, sheet)
    c1, c2 = h.parse_columns(columns)
    if width_mm is None and not auto_fit and hidden is None:
        raise OpError("Give width_mm, auto_fit=true or hidden.")
    if width_mm is not None and not 0.5 <= width_mm <= 1000:
        raise OpError("width_mm must be between 0.5 and 1000 (a default column is about 22.6).")
    cols = sh.getCellRangeByPosition(c1, 0, c2, 0).getColumns()
    values: dict[str, Any] = {}
    if width_mm is not None:
        values["Width"] = h.mm(width_mm)
    if auto_fit:
        values["OptimalWidth"] = True
    if hidden is not None:
        values["IsVisible"] = not hidden
    h.set_props(cols, values)
    result: dict[str, Any] = {"sheet": sh.Name, "columns": f"{h.col_name(c1)}:{h.col_name(c2)}"}
    if width_mm is not None or auto_fit:
        result["width_mm"] = round(sh.getColumns().getByIndex(c1).Width / 100, 1)
    return result


@op("calc.set_row_height", kind="calc", mutates=True)
@h.warn_in_csv
def set_row_height(
    s: Session,
    rows: str,
    sheet: str | None = None,
    height_mm: float | None = None,
    auto_fit: bool = False,
    hidden: bool | None = None,
) -> dict[str, Any]:
    """Set height, auto-fit or hide/show rows."""
    sh = h.get_sheet(s, sheet)
    r1, r2 = h.parse_rows(rows)
    if height_mm is None and not auto_fit and hidden is None:
        raise OpError("Give height_mm, auto_fit=true or hidden.")
    if height_mm is not None and not 0.5 <= height_mm <= 400:
        raise OpError("height_mm must be between 0.5 and 400 (a default row is about 4.5).")
    target = sh.getCellRangeByPosition(0, r1, 0, r2).getRows()
    values: dict[str, Any] = {}
    if height_mm is not None:
        values["Height"] = h.mm(height_mm)
    if auto_fit:
        values["OptimalHeight"] = True
    if hidden is not None:
        values["IsVisible"] = not hidden
    h.set_props(target, values)
    result: dict[str, Any] = {"sheet": sh.Name, "rows": f"{r1 + 1}:{r2 + 1}"}
    if height_mm is not None or auto_fit:
        result["height_mm"] = round(sh.getRows().getByIndex(r1).Height / 100, 1)
    return result


# -- conditional formatting ----------------------------------------------------

#: ScConditionMode integers understood by the conditional-format API (verified against
#: LibreOffice by round-tripping the older ConditionOperator enum).
CONDITION_MODES = {
    "equal": 0, "less": 1, "greater": 2, "less_equal": 3, "greater_equal": 4,
    "not_equal": 5, "between": 6, "not_between": 7, "duplicate": 8, "unique": 9,
    "formula": 24,
}  # fmt: skip
_NEEDS_VALUE = {"equal", "less", "greater", "less_equal", "greater_equal", "not_equal", "formula"}
_NUMERIC = re.compile(r"^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$")


def _operand(value: Any) -> str:
    """Condition operand as the formula text Calc expects: numbers raw, text quoted."""
    if isinstance(value, bool):
        raise OpError("value must be a number or text, not a boolean.")
    if isinstance(value, (int, float)):
        return repr(h.clean_number(float(value)))
    text = str(value).strip()
    if _NUMERIC.match(text):
        return text
    if text.startswith("="):  # a cell reference or expression, e.g. "=$F$1"
        return h.normalize_formula(text[1:])
    return '"' + text.replace('"', '""') + '"'


def _cf_style(s: Session, background: str | None, font_color: str | None, bold: bool) -> str:
    """Find or create the cell style that a conditional rule applies."""
    key = f"{(background or '-').lstrip('#').upper()}_{(font_color or '-').lstrip('#').upper()}_{int(bold)}"
    name = f"mcp_cf_{key}"
    styles = s.doc.getStyleFamilies().getByName("CellStyles")
    if not styles.hasByName(name):
        style = s.doc.createInstance("com.sun.star.style.CellStyle")
        styles.insertByName(name, style)
        values: dict[str, Any] = {}
        if background:
            values["CellBackColor"] = h.color_value(background)
            values["IsCellBackgroundTransparent"] = False
        if font_color:
            values["CharColor"] = h.color_value(font_color)
        if bold:
            values["CharWeight"] = 150.0
        h.set_props(style, values)
    return name


def _new_format(s: Session, sheet: Any, box: h.Box) -> Any:
    ranges = s.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
    ranges.addRangeAddress(h.range_of(sheet, box).getRangeAddress(), False)
    formats = sheet.ConditionalFormats
    new_id = formats.createByRange(ranges)
    return next(f for f in formats.getConditionalFormats() if new_id == f.ID)


@op("calc.add_conditional_format", kind="calc", mutates=True)
@h.warn_in_csv
def add_conditional_format(
    s: Session,
    range: str,
    rule: str,
    sheet: str | None = None,
    value: Any = None,
    value2: Any = None,
    background: str | None = None,
    font_color: str | None = None,
    bold: bool = False,
) -> dict[str, Any]:
    """Add a value/formula rule that applies a cell style when it matches."""
    sh, _, box = h.get_range(s, sheet, range)
    if rule not in CONDITION_MODES:
        raise OpError(f"Unknown rule {rule!r}. Use one of: {', '.join(CONDITION_MODES)}.")
    if rule in _NEEDS_VALUE and value is None:
        raise OpError(f"rule '{rule}' needs a value.")
    if rule in ("between", "not_between") and (value is None or value2 is None):
        raise OpError(f"rule '{rule}' needs both value (lower bound) and value2 (upper bound).")
    if not (background or font_color or bold):
        raise OpError("Give at least one of background, font_color or bold: the look to apply.")
    fmt = _new_format(s, sh, box)
    fmt.createEntry(0, 0)  # ConditionEntryType.CONDITION
    entry = fmt.getByIndex(0)
    entry.Operator = CONDITION_MODES[rule]
    if rule == "formula":
        entry.Formula1 = h.normalize_formula(str(value).lstrip("="))
    elif value is not None:
        entry.Formula1 = _operand(value)
    if value2 is not None:
        entry.Formula2 = _operand(value2)
    entry.StyleName = _cf_style(s, background, font_color, bold)
    return {"sheet": sh.Name, "range": h.box_name(box), "rule": rule}


@op("calc.add_data_bar", kind="calc", mutates=True)
@h.warn_in_csv
def add_data_bar(
    s: Session,
    range: str,
    sheet: str | None = None,
    color: str = "#638EC6",
    show_value: bool = True,
    gradient: bool = True,
) -> dict[str, Any]:
    """Add an in-cell data bar (colour scales cannot be created through UNO)."""
    sh, _, box = h.get_range(s, sheet, range)
    fmt = _new_format(s, sh, box)
    fmt.createEntry(2, 0)  # ConditionEntryType.DATABAR; min/max default to automatic
    bar = fmt.getByIndex(0)
    bar.Color = h.color_value(color)
    bar.ShowValue = show_value
    bar.UseGradient = gradient
    return {"sheet": sh.Name, "range": h.box_name(box), "rule": "data_bar"}


@op("calc.clear_conditional_formats", kind="calc", mutates=True)
def clear_conditional_formats(
    s: Session, sheet: str | None = None, range: str | None = None
) -> dict[str, Any]:
    """Remove conditional formats touching a range (or the whole sheet)."""
    sh = h.get_sheet(s, sheet)
    target = h.parse_range(range, whole=True) if range else None
    formats = sh.ConditionalFormats
    removed = 0
    for fmt in formats.getConditionalFormats():
        boxes = [h.box_of_address(a) for a in fmt.Range.getRangeAddresses()]
        if target is None or any(_intersects(b, target) for b in boxes):
            formats.removeByID(fmt.ID)
            removed += 1
    result: dict[str, Any] = {"sheet": sh.Name, "removed": removed}
    if not removed:
        result["_noop"] = True
    return result


def _intersects(a: h.Box, b: h.Box) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


# -- data validation -----------------------------------------------------------

_VALIDATION_TYPES = {
    "list": "LIST", "whole_number": "WHOLE", "decimal": "DECIMAL", "date": "DATE",
    "text_length": "TEXT_LEN", "custom": "CUSTOM", "any": "ANY",
}  # fmt: skip
_OPERATORS = {
    "between": "BETWEEN", "not_between": "NOT_BETWEEN", "equal": "EQUAL",
    "not_equal": "NOT_EQUAL", "greater": "GREATER", "greater_equal": "GREATER_EQUAL",
    "less": "LESS", "less_equal": "LESS_EQUAL",
}  # fmt: skip
_ISO_DATE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$")


def _bound(kind: str, value: Any) -> str:
    """Validation bound as formula text; ISO dates become DATE(y;m;d)."""
    if isinstance(value, str):
        m = _ISO_DATE.match(value.strip())
        if m:
            return f"DATE({int(m.group(1))};{int(m.group(2))};{int(m.group(3))})"
        return h.normalize_formula(value.lstrip("="))
    return repr(h.clean_number(float(value)))


@op("calc.set_data_validation", kind="calc", mutates=True)
@h.warn_in_csv
def set_data_validation(
    s: Session,
    range: str,
    type: str,
    sheet: str | None = None,
    items: list[str] | None = None,
    source_range: str | None = None,
    source_sheet: str | None = None,
    operator: str = "between",
    min_value: Any = None,
    max_value: Any = None,
    formula: str | None = None,
    allow_blank: bool = True,
    error_message: str | None = None,
    error_title: str | None = None,
    error_style: str = "stop",
    input_message: str | None = None,
    input_title: str | None = None,
) -> dict[str, Any]:
    """Set a list / number / date / length / custom validation rule."""
    sh, rng, box = h.get_range(s, sheet, range)
    if type not in _VALIDATION_TYPES:
        raise OpError(f"type must be one of: {', '.join(_VALIDATION_TYPES)}.")
    validation = rng.Validation
    validation.Type = enum("com.sun.star.sheet.ValidationType", _VALIDATION_TYPES[type])
    if type == "any":
        validation.ShowErrorMessage = False
        validation.ShowInputMessage = False
    elif type == "list":
        if items:
            validation.setFormula1(";".join('"' + str(i).replace('"', '""') + '"' for i in items))
        elif source_range:
            src_sheet, _, src_box = h.get_range(s, source_sheet or sheet, source_range)
            validation.setFormula1(_absolute_source(src_sheet.Name, src_box))
        else:
            raise OpError("A list validation needs 'items' (e.g. ['Yes','No']) or 'source_range'.")
        validation.ShowList = 1
    elif type == "custom":
        if not formula:
            raise OpError("A custom validation needs 'formula', e.g. '=COUNTIF($A:$A;A1)=1'.")
        validation.setFormula1(h.normalize_formula(formula.lstrip("=")))
    else:
        if operator not in _OPERATORS:
            raise OpError(f"operator must be one of: {', '.join(_OPERATORS)}.")
        if min_value is None:
            raise OpError("Give min_value (and max_value for 'between' / 'not_between').")
        if operator in ("between", "not_between") and max_value is None:
            raise OpError(f"operator '{operator}' needs both min_value and max_value.")
        validation.setOperator(enum("com.sun.star.sheet.ConditionOperator", _OPERATORS[operator]))
        validation.setFormula1(_bound(type, min_value))
        if max_value is not None:
            validation.setFormula2(_bound(type, max_value))
    if type != "any":
        validation.IgnoreBlankCells = allow_blank
        validation.ShowErrorMessage = True
        validation.ErrorAlertStyle = enum(
            "com.sun.star.sheet.ValidationAlertStyle",
            {"stop": "STOP", "warning": "WARNING", "info": "INFO"}.get(error_style, "STOP"),
        )
        if error_message is not None:
            validation.ErrorMessage = error_message
        if error_title is not None:
            validation.ErrorTitle = error_title
        if input_message is not None or input_title is not None:
            validation.ShowInputMessage = True
            validation.InputMessage = input_message or ""
            validation.InputTitle = input_title or ""
    rng.Validation = validation
    return {"sheet": sh.Name, "range": h.box_name(box), "validation": type}


def _absolute_source(sheet_name: str, box: h.Box) -> str:
    quoted = (
        "'" + sheet_name.replace("'", "''") + "'" if re.search(r"\W", sheet_name) else sheet_name
    )
    c1, r1, c2, r2 = box
    return f"${quoted}.${h.col_name(c1)}${r1 + 1}:${h.col_name(c2)}${r2 + 1}"
