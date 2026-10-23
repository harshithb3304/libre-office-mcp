"""Writer: page setup, headers/footers and the table of contents."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session

# Paper sizes in 1/100 mm (portrait width x height)
PAPER = {
    "A3": (29700, 42000),
    "A4": (21000, 29700),
    "A5": (14800, 21000),
    "Letter": (21590, 27940),
    "Legal": (21590, 35560),
}


@op("writer.set_page_setup", kind="writer", mutates=True)
def set_page_setup(
    s: Session,
    paper: str | None = None,
    orientation: str | None = None,
    top_mm: float | None = None,
    bottom_mm: float | None = None,
    left_mm: float | None = None,
    right_mm: float | None = None,
) -> dict[str, Any]:
    """Paper size, orientation and margins for all pages in use."""
    if paper is not None and paper not in PAPER:
        raise OpError(f"paper must be one of {', '.join(PAPER)}.")
    if orientation is not None and orientation not in ("portrait", "landscape"):
        raise OpError("orientation must be 'portrait' or 'landscape'.")
    changed: set[str] = set()
    for style in h.used_page_styles(s.doc):
        if paper is not None or orientation is not None:
            width, height = PAPER[paper] if paper else (style.Width, style.Height)
            landscape = orientation == "landscape" if orientation else bool(style.IsLandscape)
            if landscape != (width > height):
                width, height = height, width
            style.IsLandscape = landscape
            style.Width, style.Height = width, height
            changed |= ({"paper"} if paper else set()) | ({"orientation"} if orientation else set())
        for prop, value, label in (
            ("TopMargin", top_mm, "top_mm"),
            ("BottomMargin", bottom_mm, "bottom_mm"),
            ("LeftMargin", left_mm, "left_mm"),
            ("RightMargin", right_mm, "right_mm"),
        ):
            if value is not None:
                setattr(style, prop, h.mm(value))
                changed.add(label)
    if not changed:
        raise OpError("Nothing to change: give paper, orientation or a margin.")
    return {"changed": sorted(changed)}


def _fill(doc: Any, text_obj: Any, text: str, page_numbers: str, alignment: str) -> None:
    """Write ``text`` (and optional page-number fields) into a header/footer text object."""
    text_obj.setString(text)
    cursor = text_obj.createTextCursor()
    cursor.gotoEnd(False)
    if page_numbers != "none":
        if text:
            text_obj.insertString(cursor, "  |  ", False)
        text_obj.insertString(cursor, "Page ", False)
        number = doc.createInstance("com.sun.star.text.TextField.PageNumber")
        number.NumberingType = 4  # ARABIC
        text_obj.insertTextContent(cursor, number, False)
        if page_numbers == "page_of_total":
            text_obj.insertString(cursor, " of ", False)
            count = doc.createInstance("com.sun.star.text.TextField.PageCount")
            count.NumberingType = 4
            text_obj.insertTextContent(cursor, count, False)
    whole = text_obj.createTextCursor()
    whole.gotoStart(False)
    whole.gotoEnd(True)
    h.apply_para_format(whole, alignment=alignment)


@op("writer.set_header_footer", kind="writer", mutates=True)
def set_header_footer(
    s: Session,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: str = "none",
    alignment: str = "center",
) -> dict[str, Any]:
    """Set header/footer text and optional page numbers. '' removes a header/footer."""
    if page_numbers not in ("none", "page", "page_of_total"):
        raise OpError("page_numbers must be 'none', 'page' or 'page_of_total'.")
    if header_text is None and footer_text is None and page_numbers == "none":
        raise OpError("Nothing to change: give header_text, footer_text or page_numbers.")
    changed = []
    for style in h.used_page_styles(s.doc):
        if header_text is not None:
            style.HeaderIsOn = header_text != ""
            if header_text:
                _fill(s.doc, style.HeaderText, header_text, "none", alignment)
            changed.append("header")
        if footer_text is not None or page_numbers != "none":
            wants_footer = bool(footer_text) or page_numbers != "none"
            style.FooterIsOn = wants_footer
            if wants_footer:
                _fill(s.doc, style.FooterText, footer_text or "", page_numbers, alignment)
            changed.append("footer")
    return {"changed": sorted(set(changed))}


@op("writer.insert_toc", kind="writer", mutates=True)
def insert_toc(
    s: Session,
    title: str = "Table of Contents",
    levels: int = 3,
    position: str = "start",
    block_index: int | None = None,
) -> dict[str, Any]:
    """Insert a table of contents built from the headings (updates page numbers)."""
    if not 1 <= levels <= 10:
        raise OpError("levels must be 1-10.")
    slot = h.Slot(s.doc, position, block_index)
    anchor = slot.paragraph("")
    toc = s.doc.createInstance("com.sun.star.text.ContentIndex")
    toc.CreateFromOutline = True
    toc.Level = levels
    toc.Title = title
    s.doc.getText().insertTextContent(anchor.getStart(), toc, False)
    toc.update()
    return {"inserted": True, "title": title}
