"""Writer: read-only operations (``writer.get_*`` / ``writer.find_text``)."""

from __future__ import annotations

import re
from typing import Any

from libreoffice_mcp.worker.ops.common import iso_datetime
from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session

_FAMILIES = {
    "paragraph": "ParagraphStyles",
    "character": "CharacterStyles",
    "page": "PageStyles",
    "list": "NumberingStyles",
    "frame": "FrameStyles",
}


def table_grid(table: Any, max_rows: int | None = None) -> list[list[str]]:
    """Cell text as a row-major grid (merged-away cells are ''), optionally the first rows only."""
    cells = h.table_cells(table)
    if not cells:
        return []
    rows = max(r for _, r in cells) + 1
    cols = max(c for c, _ in cells) + 1
    grid = []
    for r in range(rows if max_rows is None else min(rows, max_rows)):
        grid.append([cells[(c, r)].getString() if (c, r) in cells else "" for c in range(cols)])
    return grid


def _block_text(block: Any) -> str:
    if h.is_table(block):
        return "\n".join("\t".join(row) for row in table_grid(block))
    return block.getString()


@op("writer.get_text", kind="writer")
def get_text(s: Session, start_block: int = 0, max_chars: int = 20000) -> dict[str, Any]:
    """Document text, paragraph per line, tables as tab-separated rows; paginated by block."""
    blocks = h.all_blocks(s.doc)
    if start_block and start_block >= len(blocks):
        raise OpError(f"start_block {start_block} is past the end ({len(blocks)} blocks).")
    parts: list[str] = []
    size = 0
    index = start_block
    table_no = len([b for b in blocks[:start_block] if h.is_table(b)])
    while index < len(blocks):
        block = blocks[index]
        text = _block_text(block)
        if h.is_table(block):
            text = f"[Table {table_no}]\n{text}"
            table_no += 1
        if parts and size + len(text) > max_chars:
            break
        parts.append(text)
        size += len(text) + 1
        index += 1
    truncated = index < len(blocks)
    return {
        "text": "\n".join(parts),
        "first_block": start_block,
        "blocks_returned": index - start_block,
        "total_blocks": len(blocks),
        "next_block": index if truncated else None,
    }


@op("writer.get_structure", kind="writer")
def get_structure(
    s: Session, start_block: int = 0, limit: int = 100, text_chars: int = 80
) -> dict[str, Any]:
    """One summary line per body block: index, type, style and a text preview."""
    blocks = h.all_blocks(s.doc)
    chosen = blocks[start_block : start_block + limit]
    table_no = len([b for b in blocks[:start_block] if h.is_table(b)])
    out = []
    for offset, block in enumerate(chosen):
        info = h.block_info(start_block + offset, block, text_chars)
        if info["type"] == "table":
            info["table_index"] = table_no
            table_no += 1
        out.append(info)
    end = start_block + len(chosen)
    return {
        "total_blocks": len(blocks),
        "blocks": out,
        "next_block": end if end < len(blocks) else None,
    }


@op("writer.get_outline", kind="writer")
def get_outline(s: Session) -> dict[str, Any]:
    """Headings (level, text, block index) - the document's table of contents."""
    out = []
    for index, block in enumerate(h.all_blocks(s.doc)):
        if not h.is_table(block):
            level = h.heading_level(block)
            if level:
                out.append({"block_index": index, "level": level, "text": block.getString()})
    return {"headings": out, "count": len(out)}


@op("writer.get_paragraph", kind="writer")
def get_paragraph(s: Session, block_index: int) -> dict[str, Any]:
    """Full text and formatting of one paragraph, including formatted runs."""
    block = h.paragraph_at(h.all_blocks(s.doc), block_index)
    info = h.block_info(block_index, block, max_chars=10**9)
    info["alignment"] = h.alignment_name(block.ParaAdjust)
    info["space_before_mm"] = round(block.ParaTopMargin / 100, 1)
    info["space_after_mm"] = round(block.ParaBottomMargin / 100, 1)
    info["left_indent_mm"] = round(block.ParaLeftMargin / 100, 1)
    info["font"] = block.CharFontName
    info["size"] = block.CharHeight
    runs = []
    portions = block.createEnumeration()
    while portions.hasMoreElements():
        portion = portions.nextElement()
        if portion.TextPortionType not in ("Text", "TextField"):
            continue
        text = portion.getString()
        if not text:
            continue
        run: dict[str, Any] = {"text": text}
        if portion.CharWeight >= 150:
            run["bold"] = True
        if portion.CharPosture.value == "ITALIC":
            run["italic"] = True
        if portion.CharUnderline:
            run["underline"] = True
        if portion.CharStrikeout:
            run["strikethrough"] = True
        if portion.CharColor not in (-1, 0):
            run["color"] = f"#{portion.CharColor & 0xFFFFFF:06X}"
        if portion.HyperLinkURL:
            run["link"] = portion.HyperLinkURL
        run["font"] = portion.CharFontName
        run["size"] = portion.CharHeight
        runs.append(run)
    info["runs"] = runs
    return info


@op("writer.find_text", kind="writer")
def find_text(
    s: Session,
    query: str,
    regex: bool = False,
    case_sensitive: bool = False,
    whole_word: bool = False,
    max_results: int = 50,
    context_chars: int = 40,
) -> dict[str, Any]:
    """Find text in paragraphs and table cells; returns block index, offset and context."""
    if not query:
        raise OpError("The search text is empty.")
    pattern = query if regex else re.escape(query)
    if whole_word:
        pattern = rf"\b(?:{pattern})\b"
    try:
        compiled = re.compile(pattern, 0 if case_sensitive else re.IGNORECASE)
    except re.error as exc:
        raise OpError(f"Invalid regular expression: {exc}") from exc

    matches: list[dict[str, Any]] = []
    total = 0

    def scan(text: str, where: dict[str, Any]) -> None:
        nonlocal total
        for m in compiled.finditer(text):
            if m.end() == m.start():
                continue
            total += 1
            if len(matches) < max_results:
                lo, hi = max(0, m.start() - context_chars), min(len(text), m.end() + context_chars)
                matches.append(
                    {**where, "offset": m.start(), "match": m.group(0), "context": text[lo:hi]}
                )

    table_no = 0
    for index, block in enumerate(h.all_blocks(s.doc)):
        if h.is_table(block):
            grid = table_grid(block)
            for r, row in enumerate(grid):
                for c, text in enumerate(row):
                    scan(
                        text,
                        {"block_index": index, "table_index": table_no, "cell": h.cell_name(c, r)},
                    )
            table_no += 1
        else:
            scan(block.getString(), {"block_index": index})
    return {"total_matches": total, "returned": len(matches), "matches": matches}


@op("writer.list_styles", kind="writer")
def list_styles(
    s: Session, family: str = "paragraph", contains: str | None = None
) -> dict[str, Any]:
    """Style names in a family (paragraph, character, page, list, frame)."""
    if family not in _FAMILIES:
        raise OpError(f"family must be one of {sorted(_FAMILIES)}.")
    styles = s.doc.getStyleFamilies().getByName(_FAMILIES[family])
    out = []
    for name in styles.getElementNames():
        if contains and contains.lower() not in name.lower():
            continue
        style = styles.getByName(name)
        entry: dict[str, Any] = {"name": name, "in_use": bool(style.isInUse())}
        if style.isUserDefined():
            entry["user_defined"] = True
        parent = getattr(style, "ParentStyle", "")
        if parent:
            entry["parent"] = parent
        out.append(entry)
    return {"family": family, "count": len(out), "styles": out}


@op("writer.get_tables", kind="writer")
def get_tables(s: Session) -> dict[str, Any]:
    """List tables with size and a preview of the first row."""
    out = []
    table_no = 0
    for index, block in enumerate(h.all_blocks(s.doc)):
        if h.is_table(block):
            first_row = table_grid(block, max_rows=1)
            out.append(
                {
                    "table_index": table_no,
                    "block_index": index,
                    "rows": block.getRows().getCount(),
                    "columns": block.getColumns().getCount(),
                    "first_row": first_row[0] if first_row else [],
                }
            )
            table_no += 1
    return {"count": len(out), "tables": out}


@op("writer.get_table", kind="writer")
def get_table(s: Session, table_index: int, max_rows: int = 200) -> dict[str, Any]:
    """All cell text of one table as a 2-D array (row-major)."""
    table = h.get_table(s.doc, table_index)
    grid = table_grid(table)
    return {
        "table_index": table_index,
        "rows": len(grid),
        "columns": len(grid[0]) if grid else 0,
        "cells": grid[:max_rows],
        "truncated": len(grid) > max_rows,
    }


@op("writer.get_images", kind="writer")
def get_images(s: Session) -> dict[str, Any]:
    """Images in the document: size (mm), alt text and name."""
    images = s.doc.getGraphicObjects()
    out = []
    for i in range(images.getCount()):
        img = images.getByIndex(i)
        out.append(
            {
                "image_index": i,
                "name": img.getName(),
                "width_mm": round(img.Width / 100, 1),
                "height_mm": round(img.Height / 100, 1),
                "alt_text": img.Description,
                "title": img.Title,
            }
        )
    return {"count": len(out), "images": out}


@op("writer.get_comments", kind="writer")
def get_comments(s: Session) -> dict[str, Any]:
    """Comments (annotations): author, date, text, and the text they are attached to."""
    out = []
    fields = s.doc.getTextFields().createEnumeration()
    while fields.hasMoreElements():
        field = fields.nextElement()
        if field.supportsService("com.sun.star.text.TextField.Annotation"):
            out.append(
                {
                    "comment_index": len(out),
                    "author": field.Author,
                    "date": iso_datetime(field.DateTimeValue),
                    "text": field.Content,
                    "anchor_text": field.getAnchor().getString()[:80],
                    "resolved": bool(getattr(field, "Resolved", False)),
                }
            )
    return {"count": len(out), "comments": out}


@op("writer.get_bookmarks", kind="writer")
def get_bookmarks(s: Session) -> dict[str, Any]:
    marks = s.doc.getBookmarks()
    out = []
    for name in marks.getElementNames():
        out.append({"name": name, "text": marks.getByName(name).getAnchor().getString()[:80]})
    return {"count": len(out), "bookmarks": out}


@op("writer.get_hyperlinks", kind="writer")
def get_hyperlinks(s: Session) -> dict[str, Any]:
    """Hyperlinks: the linked text, its URL and the block it is in."""
    out = []
    for index, block in enumerate(h.all_blocks(s.doc)):
        if h.is_table(block):
            continue
        portions = block.createEnumeration()
        while portions.hasMoreElements():
            portion = portions.nextElement()
            url = getattr(portion, "HyperLinkURL", "")
            if url:
                out.append({"block_index": index, "text": portion.getString(), "url": url})
    return {"count": len(out), "hyperlinks": out}


@op("writer.get_footnotes", kind="writer")
def get_footnotes(s: Session) -> dict[str, Any]:
    notes = s.doc.getFootnotes()
    out = [
        {
            "footnote_index": i,
            "label": notes.getByIndex(i).Label,
            "text": notes.getByIndex(i).getString(),
        }
        for i in range(notes.getCount())
    ]
    return {"count": len(out), "footnotes": out}


def first_page_style(doc: Any) -> str:
    enumeration = doc.getText().createEnumeration()
    if enumeration.hasMoreElements():
        try:
            return enumeration.nextElement().PageStyleName or "Standard"
        except Exception:
            pass
    return "Standard"


@op("writer.get_page_setup", kind="writer")
def get_page_setup(s: Session) -> dict[str, Any]:
    """Paper size, orientation and margins of the document's first page style."""
    name = first_page_style(s.doc)
    st = s.doc.getStyleFamilies().getByName("PageStyles").getByName(name)
    return {
        "page_style": name,
        "width_mm": st.Width / 100,
        "height_mm": st.Height / 100,
        "landscape": bool(st.IsLandscape),
        "margins_mm": {
            "top": st.TopMargin / 100,
            "bottom": st.BottomMargin / 100,
            "left": st.LeftMargin / 100,
            "right": st.RightMargin / 100,
        },
    }


@op("writer.get_headers_footers", kind="writer")
def get_headers_footers(s: Session) -> dict[str, Any]:
    """Header and footer text of every page style in use."""
    styles = s.doc.getStyleFamilies().getByName("PageStyles")
    out = []
    for name in styles.getElementNames():
        st = styles.getByName(name)
        if not st.isInUse():
            continue
        out.append(
            {
                "page_style": name,
                "header": st.HeaderText.getString() if st.HeaderIsOn else None,
                "footer": st.FooterText.getString() if st.FooterIsOn else None,
            }
        )
    return {"page_styles": out}


@op("writer.get_track_changes", kind="writer")
def get_track_changes(s: Session, limit: int = 100) -> dict[str, Any]:
    """Whether change tracking is on, plus recorded changes (type, author, date)."""
    redlines = s.doc.getRedlines()
    out = []
    enumeration = redlines.createEnumeration()
    while enumeration.hasMoreElements() and len(out) < limit:
        r = enumeration.nextElement()
        out.append(
            {
                "type": r.getPropertyValue("RedlineType"),
                "author": r.getPropertyValue("RedlineAuthor"),
                "date": iso_datetime(r.getPropertyValue("RedlineDateTime")),
            }
        )
    return {
        "tracking_enabled": bool(s.doc.RecordChanges),
        "count": redlines.getCount(),
        "changes": out,
    }
