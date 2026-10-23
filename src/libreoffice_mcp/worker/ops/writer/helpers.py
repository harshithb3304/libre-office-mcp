"""Shared building blocks for the Writer operations.

Vocabulary used throughout:

* **block**  - a top-level item of the document body: a paragraph (headings and list
  items are paragraphs too) or a table. Blocks are addressed by 0-based index in
  document order. Paragraphs *inside* table cells are not blocks; use the table ops.
* **slot**   - a position where new blocks are inserted (:class:`Slot`).
"""

from __future__ import annotations

import difflib
import re
from typing import Any

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.registry import OpError
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color

TEXT_TABLE = "com.sun.star.text.TextTable"
DEFAULT_STYLE = "Standard"  # the default paragraph style; keeps .docx round-trips natural

ALIGNMENTS = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "BLOCK"}
_ALIGN_NAMES = {0: "left", 1: "right", 2: "justify", 3: "center"}

MM100 = 100  # UNO lengths are in 1/100 mm


def mm(value: float) -> int:
    return round(value * MM100)


# -- blocks --------------------------------------------------------------------


def is_table(block: Any) -> bool:
    return bool(block.supportsService(TEXT_TABLE))


def all_blocks(doc: Any) -> list[Any]:
    enumeration = doc.getText().createEnumeration()
    blocks = []
    while enumeration.hasMoreElements():
        blocks.append(enumeration.nextElement())
    return blocks


def check_index(blocks: list[Any], index: int, what: str = "block_index") -> Any:
    if not 0 <= index < len(blocks):
        raise OpError(
            f"{what} {index} is out of range: the document has {len(blocks)} blocks "
            f"(valid: 0..{len(blocks) - 1}). Use writer_get_structure to see them."
        )
    return blocks[index]


def paragraph_at(blocks: list[Any], index: int, what: str = "block_index") -> Any:
    block = check_index(blocks, index, what)
    if is_table(block):
        raise OpError(f"Block {index} is a table, not a paragraph. Use the writer_*table* tools.")
    return block


def table_blocks(blocks: list[Any]) -> list[Any]:
    return [b for b in blocks if is_table(b)]


def get_table(doc: Any, table_index: int) -> Any:
    tables = table_blocks(all_blocks(doc))
    if not 0 <= table_index < len(tables):
        raise OpError(
            f"table_index {table_index} is out of range: the document has {len(tables)} "
            "tables (valid: 0..n-1). Use writer_get_tables to list them.",
            protocol.E_NOT_FOUND,
        )
    return tables[table_index]


_CELL = re.compile(r"^([A-Za-z]+)(\d+)$")


def parse_cell(name: str) -> tuple[int, int]:
    """Writer cell name -> (col, row), both 0-based: 'A1' -> (0, 0), 'C3' -> (2, 2)."""
    m = _CELL.match(name.strip())
    if not m:
        raise OpError(f"Invalid cell name {name!r}; use e.g. 'A1' or 'B3'.")
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return col - 1, int(m.group(2)) - 1


def cell_name(col: int, row: int) -> str:
    """(col, row) -> Writer cell name: (0, 0) -> 'A1'; columns past Z continue AA, AB..."""
    letters = ""
    n = col + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row + 1}"


def table_cells(table: Any) -> dict[tuple[int, int], Any]:
    """All real cells of a table keyed by (col, row).

    Built from the cell names rather than rows x columns counts, because after cells are
    merged LibreOffice's column count no longer matches the grid.
    """
    cells = {}
    for name in table.getCellNames():
        if _CELL.match(name):  # skips split-cell names like 'A1.1.1'
            cells[parse_cell(name)] = table.getCellByName(name)
    return cells


def heading_level(paragraph: Any) -> int:
    """Outline level of a paragraph (1-10 for headings, 0 for body text)."""
    try:
        return int(paragraph.OutlineLevel)
    except Exception:
        return 0


def block_info(index: int, block: Any, max_chars: int = 80) -> dict[str, Any]:
    if is_table(block):
        return {
            "index": index,
            "type": "table",
            "rows": block.getRows().getCount(),
            "columns": block.getColumns().getCount(),
            "name": block.getName(),
        }
    text = block.getString()
    info: dict[str, Any] = {"index": index, "style": block.ParaStyleName}
    level = heading_level(block)
    if level:
        info["type"] = "heading"
        info["level"] = level
    elif block.NumberingStyleName:
        info["type"] = "list_item"
        info["list_style"] = block.NumberingStyleName
        info["list_level"] = int(block.NumberingLevel)
    else:
        info["type"] = "paragraph"
    info["text"] = text if len(text) <= max_chars else text[:max_chars] + "..."
    info["length"] = len(text)
    return info


def resolve_style(doc: Any, family: str, name: str) -> str:
    """Return the exact style name (case-insensitive match) or raise with suggestions."""
    styles = doc.getStyleFamilies().getByName(family)
    if styles.hasByName(name):
        return name
    names = list(styles.getElementNames())
    for candidate in names:
        if candidate.lower() == name.lower():
            return candidate
    close = difflib.get_close_matches(name, names, n=6, cutoff=0.4)
    hint = f" Similar: {', '.join(close)}." if close else ""
    raise OpError(
        f"Style '{name}' does not exist.{hint} Use writer_list_styles to see all styles.",
        protocol.E_NOT_FOUND,
    )


def reset_formatting(paragraph: Any) -> None:
    """Drop all direct character and paragraph formatting from ``paragraph``.

    Paragraph objects lack ``setAllPropertiesToDefault``; a text cursor over the same
    range has it. (The paragraph *style* is a property too, so callers reapply it.)
    """
    cursor = paragraph.getText().createTextCursorByRange(paragraph)
    cursor.gotoStartOfParagraph(False)
    cursor.gotoEndOfParagraph(True)
    cursor.setAllPropertiesToDefault()
    paragraph.NumberingStyleName = ""  # list membership is not reset by the call above


# -- insertion slots -----------------------------------------------------------


class Slot:
    """A moving insertion point: ``position`` is ``end``, ``start``, ``before`` or ``after``.

    Each ``paragraph()``/``table()`` call inserts at the slot and advances it, so a list
    of blocks lands in order. For ``end`` an empty trailing paragraph (every new
    document has one, and Writer keeps one after a final table) is *reused* instead of
    leaving a blank line behind.

    Performance note: every UNO call is a round trip to LibreOffice (about 1 ms). The
    document is enumerated exactly once, in ``__init__``; after that the slot only moves a
    live text cursor, so inserting n blocks costs O(n) calls, not O(n^2). ``self.after`` is
    kept as a plain counter (index of the last block before the slot) for the same reason.
    """

    def __init__(self, doc: Any, position: str, block_index: int | None):
        self.doc = doc
        self.text = doc.getText()
        #: Cursor at the END of the paragraph the next block goes after (None: see _before).
        self._cur: Any = None
        #: Cursor at the START of the paragraph the next block goes before.
        self._before: Any = None
        #: Next paragraph() fills the empty paragraph at ``_cur`` instead of splitting.
        self._reuse = False
        blocks = all_blocks(doc)  # the one and only enumeration
        if not blocks:
            raise OpError("The document has no paragraphs.")
        if position == "end":
            last = blocks[-1]
            self._cur = self.text.createTextCursorByRange(last.getEnd())
            self._reuse = not is_table(last) and not last.getString()
            self.after = len(blocks) - 2 if self._reuse else len(blocks) - 1
        elif position in ("start", "before", "after"):
            if position != "start":
                if block_index is None:
                    raise OpError(f"position='{position}' needs block_index.")
                check_index(blocks, block_index)
            index = -1 if position == "start" else block_index - (position == "before")
            self.after = index
            if index >= 0 and not is_table(blocks[index]):
                self._cur = self.text.createTextCursorByRange(blocks[index].getEnd())
            else:  # document start, or right after a table: split before the next paragraph
                following = blocks[index + 1] if index + 1 < len(blocks) else None
                if following is None or is_table(following):
                    raise OpError(
                        "Cannot insert here (between two tables, or before a table at the "
                        "very start of the document). Insert after a paragraph instead."
                    )
                self._before = self.text.createTextCursorByRange(following.getStart())
        else:
            raise OpError(f"position must be end, start, before or after (got {position!r}).")

    def paragraph(self, text: str = "", style: str | None = None) -> Any:
        """Insert a paragraph at the slot and return a cursor covering it.

        The result supports the same paragraph/character properties as a paragraph object,
        plus ``getStart()``/``getEnd()``; callers format it directly. It is reset to
        ``style`` (default Standard) with no direct formatting or list membership.
        """
        from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

        if self._reuse:
            self._reuse = False
        else:
            if self._cur is None:
                # Split at the start of the following paragraph: the new empty paragraph
                # appears *before* it, and the cursor is left inside the old one.
                self._cur, self._before = self._before, None
                self.text.insertControlCharacter(self._cur, PARAGRAPH_BREAK, False)
                self._cur.goLeft(1, False)  # step back into the new empty paragraph
            else:
                self.text.insertControlCharacter(self._cur, PARAGRAPH_BREAK, False)
        self.after += 1
        para = self.text.createTextCursorByRange(self._cur)
        para.gotoStartOfParagraph(False)
        para.gotoEndOfParagraph(True)
        # A split copies the neighbour's formatting; start from a clean slate.
        para.setAllPropertiesToDefault()
        para.NumberingStyleName = ""  # list membership is not reset by the call above
        para.ParaStyleName = style or DEFAULT_STYLE
        if text:
            para.setString(text)
        self._cur.gotoEndOfParagraph(False)
        return para

    def table(self, rows: int, columns: int) -> Any:
        """Insert an empty ``rows`` x ``columns`` table at the slot and return it.

        Writer always keeps a paragraph after a table. We create the table in front of a
        fresh empty paragraph, then drop that paragraph again unless it is the last block.
        """
        anchor = self.paragraph("")
        index = self.after  # the anchor's index; the table will take this index
        table = self.doc.createInstance(TEXT_TABLE)
        table.initialize(rows, columns)
        self.text.insertTextContent(anchor.getStart(), table, False)
        self.after = index  # now the table

        probe = self.text.createTextCursorByRange(self._cur)
        if not probe.gotoNextParagraph(False):
            # The anchor is the last paragraph: keep it, and let the next append fill it.
            self._reuse = True
            return table
        try:
            following_is_table = probe.TextTable is not None
        except Exception:
            following_is_table = False
        if following_is_table:  # another table follows: the empty paragraph must stay
            self._reuse = True
            return table
        # A normal paragraph follows: delete the empty anchor so it directly follows the table.
        doomed = self.text.createTextCursorByRange(self._cur)
        doomed.gotoStartOfParagraph(False)
        doomed.gotoNextParagraph(True)  # selects the anchor including its paragraph break
        doomed.setString("")
        self._before = self.text.createTextCursorByRange(self._cur)
        self._before.gotoStartOfParagraph(False)
        self._cur = None
        return table


# -- formatting ----------------------------------------------------------------


def apply_char_format(
    target: Any,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    color: str | None = None,
    background_color: str | None = None,
    font: str | None = None,
    size: float | None = None,
    superscript: bool | None = None,
    subscript: bool | None = None,
) -> list[str]:
    """Apply character formatting to anything with Char* properties (text range, paragraph,
    style). ``None`` leaves a property alone; ``False`` switches it off. Returns what changed."""
    changed = []
    if bold is not None:
        target.CharWeight = 150.0 if bold else 100.0
        changed.append("bold")
    if italic is not None:
        target.CharPosture = enum("com.sun.star.awt.FontSlant", "ITALIC" if italic else "NONE")
        changed.append("italic")
    if underline is not None:
        target.CharUnderline = 1 if underline else 0  # FontUnderline.SINGLE / NONE
        changed.append("underline")
    if strikethrough is not None:
        target.CharStrikeout = 1 if strikethrough else 0  # FontStrikeout.SINGLE / NONE
        changed.append("strikethrough")
    if color is not None:
        target.CharColor = parse_color(color)
        changed.append("color")
    if background_color is not None:
        target.CharBackColor = parse_color(background_color)
        target.CharBackTransparent = False
        changed.append("background_color")
    if font is not None:
        target.CharFontName = font
        changed.append("font")
    if size is not None:
        if size <= 0:
            raise OpError("size must be positive (points).")
        target.CharHeight = float(size)
        changed.append("size")
    if superscript:
        target.CharEscapement, target.CharEscapementHeight = 33, 58
        changed.append("superscript")
    elif subscript:
        target.CharEscapement, target.CharEscapementHeight = -33, 58
        changed.append("subscript")
    elif superscript is False or subscript is False:
        target.CharEscapement, target.CharEscapementHeight = 0, 100
        changed.append("baseline")
    return changed


def apply_para_format(
    target: Any,
    *,
    alignment: str | None = None,
    line_spacing_percent: float | None = None,
    space_before_mm: float | None = None,
    space_after_mm: float | None = None,
    left_indent_mm: float | None = None,
    right_indent_mm: float | None = None,
    first_line_indent_mm: float | None = None,
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    background_color: str | None = None,
) -> list[str]:
    """Apply paragraph formatting to a paragraph, text range or paragraph style."""
    from com.sun.star.style import LineSpacing

    changed = []
    if alignment is not None:
        if alignment not in ALIGNMENTS:
            raise OpError(f"alignment must be one of {sorted(ALIGNMENTS)}.")
        target.ParaAdjust = enum("com.sun.star.style.ParagraphAdjust", ALIGNMENTS[alignment])
        changed.append("alignment")
    if line_spacing_percent is not None:
        spacing = LineSpacing()
        spacing.Mode = 0  # PROP: height is a percentage
        spacing.Height = int(line_spacing_percent)
        target.ParaLineSpacing = spacing
        changed.append("line_spacing_percent")
    for name, value, prop in (
        ("space_before_mm", space_before_mm, "ParaTopMargin"),
        ("space_after_mm", space_after_mm, "ParaBottomMargin"),
        ("left_indent_mm", left_indent_mm, "ParaLeftMargin"),
        ("right_indent_mm", right_indent_mm, "ParaRightMargin"),
        ("first_line_indent_mm", first_line_indent_mm, "ParaFirstLineIndent"),
    ):
        if value is not None:
            setattr(target, prop, mm(value))
            changed.append(name)
    if keep_with_next is not None:
        target.ParaKeepTogether = bool(keep_with_next)
        changed.append("keep_with_next")
    if page_break_before is not None:
        target.BreakType = enum(
            "com.sun.star.style.BreakType", "PAGE_BEFORE" if page_break_before else "NONE"
        )
        changed.append("page_break_before")
    if background_color is not None:
        target.ParaBackColor = parse_color(background_color)
        target.ParaBackTransparent = False
        changed.append("background_color")
    return changed


def alignment_name(value: Any) -> str:
    return _ALIGN_NAMES.get(int(value), "left")


def first_page_style(doc: Any) -> str:
    """Name of the page style the document starts with (``Standard`` for new documents)."""
    enumeration = doc.getText().createEnumeration()
    if enumeration.hasMoreElements():
        try:
            return enumeration.nextElement().PageStyleName or "Standard"
        except Exception:
            pass
    return "Standard"


def used_page_styles(doc: Any) -> list[Any]:
    """Page styles the document uses.

    ``isInUse()`` is the obvious test, but on some LibreOffice versions it reports False for
    every style in a freshly loaded headless document, so fall back to the first page's style.
    """
    styles = doc.getStyleFamilies().getByName("PageStyles")
    used = [styles.getByName(n) for n in styles.getElementNames() if styles.getByName(n).isInUse()]
    return used or [styles.getByName(first_page_style(doc))]


# -- searching -----------------------------------------------------------------


def search_descriptor(
    doc: Any,
    query: str,
    *,
    regex: bool = False,
    case_sensitive: bool = False,
    whole_word: bool = False,
) -> Any:
    if not query:
        raise OpError("The search text is empty.")
    if regex:
        try:
            re.compile(query)
        except re.error as exc:
            raise OpError(f"Invalid regular expression: {exc}") from exc
    desc = doc.createSearchDescriptor()
    desc.SearchString = query
    desc.SearchRegularExpression = regex
    desc.SearchCaseSensitive = case_sensitive
    desc.SearchWords = whole_word
    return desc


def find_ranges(doc: Any, query: str, **kwargs: Any) -> list[Any]:
    found = doc.findAll(search_descriptor(doc, query, **kwargs))
    return [found.getByIndex(i) for i in range(found.getCount())]


def target_ranges(
    doc: Any,
    *,
    find: str | None,
    block_index: int | None,
    end_block_index: int | None = None,
    occurrence: str = "all",
    case_sensitive: bool = False,
) -> list[Any]:
    """Ranges to act on: occurrences of ``find`` (optionally within blocks) or whole blocks."""
    blocks = all_blocks(doc)
    if find:
        ranges = find_ranges(doc, find, case_sensitive=case_sensitive)
        if block_index is not None:
            last = end_block_index if end_block_index is not None else block_index
            check_index(blocks, block_index)
            check_index(blocks, last, "end_block_index")
            wanted = [
                paragraph_at(blocks, i)
                for i in range(block_index, last + 1)
                if not is_table(blocks[i])
            ]
            ranges = [r for r in ranges if any(_inside(doc, r, p) for p in wanted)]
        if not ranges:
            raise OpError(f"Text {find!r} was not found.", protocol.E_NOT_FOUND)
        return ranges[:1] if occurrence == "first" else ranges
    if block_index is None:
        raise OpError("Give either 'find' (text to look for) or 'block_index'.")
    last = end_block_index if end_block_index is not None else block_index
    check_index(blocks, block_index)
    check_index(blocks, last, "end_block_index")
    if last < block_index:
        raise OpError("end_block_index must be >= block_index.")
    return [paragraph_at(blocks, i) for i in range(block_index, last + 1)]


def _inside(doc: Any, found: Any, paragraph: Any) -> bool:
    """True if range ``found`` starts inside ``paragraph``."""
    text = doc.getText()
    try:
        return (
            text.compareRegionStarts(paragraph.getStart(), found.getStart()) >= 0
            and text.compareRegionEnds(found.getStart(), paragraph.getEnd()) >= 0
        )
    except Exception:
        return False


def require_writer(session: Session) -> None:
    session.require("writer")
