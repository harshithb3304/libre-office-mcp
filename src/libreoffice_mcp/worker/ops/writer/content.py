"""Writer: adding, editing and removing body content (paragraphs, headings, lists...)."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.ops.writer import objects, tables
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum

BULLET_LIST = "List 1"
NUMBERED_LIST = "Numbering 123"
MAX_HEADING = 9


def _format_new_paragraph(para: Any, alignment: str | None, char: dict[str, Any]) -> None:
    if alignment:
        h.apply_para_format(para, alignment=alignment)
    char = {k: v for k, v in char.items() if v is not None}
    if char:
        h.apply_char_format(para, **char)


def _add_paragraph(
    slot: h.Slot,
    text: str,
    style: str | None = None,
    alignment: str | None = None,
    **char: Any,
) -> dict[str, Any]:
    doc = slot.doc
    style_name = h.resolve_style(doc, "ParagraphStyles", style) if style else None
    para = slot.paragraph(text, style_name)
    _format_new_paragraph(para, alignment, char)
    return {"block_index": slot.after}


def _add_heading(
    slot: h.Slot, text: str, level: int = 1, alignment: str | None = None
) -> dict[str, Any]:
    if not 1 <= level <= MAX_HEADING:
        raise OpError(f"Heading level must be 1-{MAX_HEADING} (got {level}).")
    para = slot.paragraph(text, f"Heading {level}")
    _format_new_paragraph(para, alignment, {})
    return {"block_index": slot.after, "level": level}


def _add_list(slot: h.Slot, items: list[Any], ordered: bool = False) -> dict[str, Any]:
    if not items:
        raise OpError("The list has no items.")
    style = NUMBERED_LIST if ordered else BULLET_LIST
    first = slot.after + 1
    for n, item in enumerate(items):
        text, level = (
            (item, 0) if isinstance(item, str) else (item.get("text", ""), item.get("level", 0))
        )
        if not 0 <= level <= 8:
            raise OpError("List item level must be 0-8.")
        para = slot.paragraph(str(text))
        para.NumberingStyleName = style
        para.NumberingLevel = level
        if ordered and n == 0:  # start counting at 1 even if another numbered list precedes
            para.ParaIsNumberingRestart = True
            para.NumberingStartValue = 1
    return {"first_block_index": first, "items": len(items)}


@op("writer.add_paragraph", kind="writer", mutates=True)
def add_paragraph(
    s: Session,
    text: str,
    style: str | None = None,
    alignment: str | None = None,
    position: str = "end",
    block_index: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font: str | None = None,
    size: float | None = None,
    color: str | None = None,
) -> dict[str, Any]:
    slot = h.Slot(s.doc, position, block_index)
    return _add_paragraph(
        slot, text, style, alignment,
        bold=bold, italic=italic, underline=underline, font=font, size=size, color=color,
    )  # fmt: skip


@op("writer.add_heading", kind="writer", mutates=True)
def add_heading(
    s: Session,
    text: str,
    level: int = 1,
    alignment: str | None = None,
    position: str = "end",
    block_index: int | None = None,
) -> dict[str, Any]:
    return _add_heading(h.Slot(s.doc, position, block_index), text, level, alignment)


@op("writer.add_list", kind="writer", mutates=True)
def add_list(
    s: Session,
    items: list[Any],
    ordered: bool = False,
    position: str = "end",
    block_index: int | None = None,
) -> dict[str, Any]:
    return _add_list(h.Slot(s.doc, position, block_index), items, ordered)


_BLOCK_TYPES = ("heading", "paragraph", "bullets", "numbered", "table", "image", "page_break")


@op("writer.add_content", kind="writer", mutates=True)
def add_content(
    s: Session,
    blocks: list[dict[str, Any]],
    position: str = "end",
    block_index: int | None = None,
) -> dict[str, Any]:
    """Insert many blocks in order with one call (see the tool description for block types)."""
    slot = h.Slot(s.doc, position, block_index)
    first = len(h.all_blocks(s.doc)) if position == "end" else None
    for n, spec in enumerate(blocks):
        kind = spec.get("type")
        where = f"blocks[{n}] ({kind})"
        try:
            if kind == "heading":
                _add_heading(slot, spec["text"], spec.get("level", 1), spec.get("alignment"))
            elif kind == "paragraph":
                char = {
                    k: spec.get(k) for k in ("bold", "italic", "underline", "font", "size", "color")
                }
                _add_paragraph(slot, spec["text"], spec.get("style"), spec.get("alignment"), **char)
            elif kind in ("bullets", "numbered"):
                _add_list(slot, spec["items"], ordered=kind == "numbered")
            elif kind == "table":
                tables.build_table(
                    slot, spec["data"], spec.get("header", True), spec.get("column_widths")
                )
            elif kind == "image":
                objects.insert_image_at(
                    s,
                    slot,
                    spec["image_path"],
                    spec.get("width_mm"),
                    spec.get("height_mm"),
                    spec.get("alt_text"),
                )
            elif kind == "page_break":
                slot.paragraph("").BreakType = enum("com.sun.star.style.BreakType", "PAGE_BEFORE")
            else:
                raise OpError(f"unknown block type {kind!r}; use one of {', '.join(_BLOCK_TYPES)}")
        except KeyError as exc:
            raise OpError(f"{where}: missing required field {exc}") from exc
        except OpError as exc:
            raise OpError(f"{where}: {exc.message}", exc.code) from exc
    return {"blocks_added": len(blocks), "last_block_index": slot.after, "first_new_hint": first}


@op("writer.edit_paragraph", kind="writer", mutates=True)
def edit_paragraph(s: Session, block_index: int, text: str) -> dict[str, Any]:
    """Replace the text of a paragraph, keeping its style and paragraph formatting."""
    para = h.paragraph_at(h.all_blocks(s.doc), block_index)
    para.setString(text)
    return {"block_index": block_index}


@op("writer.delete_blocks", kind="writer", mutates=True)
def delete_blocks(s: Session, start_block: int, end_block: int | None = None) -> dict[str, Any]:
    """Delete paragraphs and/or tables from start_block to end_block (inclusive)."""
    blocks = h.all_blocks(s.doc)
    end = start_block if end_block is None else end_block
    h.check_index(blocks, start_block, "start_block")
    h.check_index(blocks, end, "end_block")
    if end < start_block:
        raise OpError("end_block must be >= start_block.")
    text = s.doc.getText()
    for block in reversed(blocks[start_block : end + 1]):  # back to front keeps indexes valid
        try:
            text.removeTextContent(block)
        except Exception:
            if h.is_table(block):
                raise
            block.setString("")  # the final paragraph cannot be removed; empty it instead
    return {"deleted": end - start_block + 1, "blocks_remaining": len(h.all_blocks(s.doc))}


@op("writer.replace_text", kind="writer", mutates=True)
def replace_text(
    s: Session,
    search: str,
    replace: str,
    regex: bool = False,
    case_sensitive: bool = False,
    whole_word: bool = False,
    first_only: bool = False,
) -> dict[str, Any]:
    """Replace text everywhere (body, tables, headers...). With regex, $1 refers to groups."""
    doc = s.doc
    if first_only:
        found = doc.findFirst(
            h.search_descriptor(
                doc, search, regex=regex, case_sensitive=case_sensitive, whole_word=whole_word
            )
        )
        if found is None:
            return {"replacements": 0, "_noop": True}
        found.setString(replace)
        return {"replacements": 1}
    desc = h.search_descriptor(
        doc, search, regex=regex, case_sensitive=case_sensitive, whole_word=whole_word
    )
    desc_replace = doc.createReplaceDescriptor()
    desc_replace.SearchString = desc.SearchString
    desc_replace.ReplaceString = replace
    desc_replace.SearchRegularExpression = regex
    desc_replace.SearchCaseSensitive = case_sensitive
    desc_replace.SearchWords = whole_word
    count = doc.replaceAll(desc_replace)
    return {"replacements": count, **({"_noop": True} if count == 0 else {})}


@op("writer.insert_page_break", kind="writer", mutates=True)
def insert_page_break(
    s: Session, position: str = "end", block_index: int | None = None
) -> dict[str, Any]:
    """Start a new page: sets 'page break before' on the paragraph that should begin it."""
    doc = s.doc
    blocks = h.all_blocks(doc)
    page_before = enum("com.sun.star.style.BreakType", "PAGE_BEFORE")
    if position == "start":
        target = 0
    elif position == "before":
        if block_index is None:
            raise OpError("position='before' needs block_index.")
        h.check_index(blocks, block_index)
        target = block_index
    elif position == "after":
        if block_index is None:
            raise OpError("position='after' needs block_index.")
        h.check_index(blocks, block_index)
        target = block_index + 1
    elif position == "end":
        target = len(blocks)
    else:
        raise OpError("position must be end, start, before or after.")
    if target >= len(blocks):  # nothing follows yet: make an empty paragraph to start the page
        slot = h.Slot(doc, "end", None)
        slot.paragraph("").BreakType = page_before
        return {"page_break_before_block": slot.after}
    blocks[target].BreakType = page_before
    return {"page_break_before_block": target}


@op("writer.set_list", kind="writer", mutates=True)
def set_list(
    s: Session,
    block_index: int,
    end_block_index: int | None = None,
    kind: str = "bullets",
    level: int = 0,
) -> dict[str, Any]:
    """Turn existing paragraphs into a bullet/numbered list, or remove list formatting."""
    if kind not in ("bullets", "numbered", "none"):
        raise OpError("kind must be 'bullets', 'numbered' or 'none'.")
    blocks = h.all_blocks(s.doc)
    last = block_index if end_block_index is None else end_block_index
    h.check_index(blocks, block_index)
    h.check_index(blocks, last, "end_block_index")
    for i in range(block_index, last + 1):
        para = h.paragraph_at(blocks, i)
        if kind == "none":
            para.NumberingStyleName = ""
        else:
            para.NumberingStyleName = NUMBERED_LIST if kind == "numbered" else BULLET_LIST
            para.NumberingLevel = level
            if kind == "numbered" and i == block_index:
                para.ParaIsNumberingRestart = True
                para.NumberingStartValue = 1
    return {"paragraphs": last - block_index + 1, "kind": kind}
