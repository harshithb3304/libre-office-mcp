"""Writer: character, paragraph and style formatting."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session


@op("writer.format_text", kind="writer", mutates=True)
def format_text(
    s: Session,
    find: str | None = None,
    block_index: int | None = None,
    end_block_index: int | None = None,
    occurrence: str = "all",
    case_sensitive: bool = False,
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
) -> dict[str, Any]:
    """Character formatting for occurrences of ``find``, or for whole paragraphs."""
    if occurrence not in ("all", "first"):
        raise OpError("occurrence must be 'all' or 'first'.")
    ranges = h.target_ranges(
        s.doc,
        find=find,
        block_index=block_index,
        end_block_index=end_block_index,
        occurrence=occurrence,
        case_sensitive=case_sensitive,
    )
    changed: list[str] = []
    for target in ranges:
        changed = h.apply_char_format(
            target,
            bold=bold, italic=italic, underline=underline, strikethrough=strikethrough,
            color=color, background_color=background_color, font=font, size=size,
            superscript=superscript, subscript=subscript,
        )  # fmt: skip
    if not changed:
        raise OpError("No formatting requested (set bold, italic, color, size, ...).")
    return {"ranges_formatted": len(ranges), "applied": changed}


@op("writer.format_paragraph", kind="writer", mutates=True)
def format_paragraph(
    s: Session,
    block_index: int,
    end_block_index: int | None = None,
    style: str | None = None,
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
) -> dict[str, Any]:
    """Paragraph style, alignment, spacing and indents for one or a range of paragraphs."""
    blocks = h.all_blocks(s.doc)
    last = block_index if end_block_index is None else end_block_index
    h.check_index(blocks, block_index)
    h.check_index(blocks, last, "end_block_index")
    style_name = h.resolve_style(s.doc, "ParagraphStyles", style) if style else None
    changed: list[str] = []
    for i in range(block_index, last + 1):
        para = h.paragraph_at(blocks, i)
        if style_name:
            para.ParaStyleName = style_name
            changed = ["style"]
        changed += h.apply_para_format(
            para,
            alignment=alignment, line_spacing_percent=line_spacing_percent,
            space_before_mm=space_before_mm, space_after_mm=space_after_mm,
            left_indent_mm=left_indent_mm, right_indent_mm=right_indent_mm,
            first_line_indent_mm=first_line_indent_mm, keep_with_next=keep_with_next,
            page_break_before=page_break_before, background_color=background_color,
        )  # fmt: skip
        changed = list(dict.fromkeys(changed))
    if not changed:
        raise OpError("No formatting requested.")
    return {"paragraphs": last - block_index + 1, "applied": changed}


@op("writer.clear_formatting", kind="writer", mutates=True)
def clear_formatting(
    s: Session, block_index: int, end_block_index: int | None = None
) -> dict[str, Any]:
    """Remove direct formatting (bold, colours, alignment...) so the paragraph style rules."""
    blocks = h.all_blocks(s.doc)
    last = block_index if end_block_index is None else end_block_index
    h.check_index(blocks, block_index)
    h.check_index(blocks, last, "end_block_index")
    for i in range(block_index, last + 1):
        para = h.paragraph_at(blocks, i)
        style = para.ParaStyleName
        h.reset_formatting(para)
        para.ParaStyleName = style  # keep the style itself
    return {"paragraphs": last - block_index + 1}


@op("writer.set_style", kind="writer", mutates=True)
def set_style(
    s: Session,
    name: str,
    family: str = "paragraph",
    parent: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    background_color: str | None = None,
    font: str | None = None,
    size: float | None = None,
    alignment: str | None = None,
    line_spacing_percent: float | None = None,
    space_before_mm: float | None = None,
    space_after_mm: float | None = None,
    left_indent_mm: float | None = None,
    first_line_indent_mm: float | None = None,
) -> dict[str, Any]:
    """Create a paragraph/character style, or change it if it already exists."""
    if family not in ("paragraph", "character"):
        raise OpError("family must be 'paragraph' or 'character' (page setup has its own tool).")
    group = "ParagraphStyles" if family == "paragraph" else "CharacterStyles"
    styles = s.doc.getStyleFamilies().getByName(group)
    created = not styles.hasByName(name)
    if created:
        style = s.doc.createInstance(
            "com.sun.star.style.ParagraphStyle"
            if family == "paragraph"
            else "com.sun.star.style.CharacterStyle"
        )
        styles.insertByName(name, style)
    style = styles.getByName(name)
    changed = []
    if parent is not None:
        style.ParentStyle = h.resolve_style(s.doc, group, parent)
        changed.append("parent")
    changed += h.apply_char_format(
        style, bold=bold, italic=italic, underline=underline, color=color,
        background_color=background_color, font=font, size=size,
    )  # fmt: skip
    if family == "paragraph":
        changed += h.apply_para_format(
            style, alignment=alignment, line_spacing_percent=line_spacing_percent,
            space_before_mm=space_before_mm, space_after_mm=space_after_mm,
            left_indent_mm=left_indent_mm, first_line_indent_mm=first_line_indent_mm,
        )  # fmt: skip
    return {"style": name, "created": created, "changed": changed}
