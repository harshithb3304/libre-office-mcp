"""Writer tools that add, change, format and remove body content."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import (
    Alignment,
    AnchorBlock,
    BlockIndex,
    Color,
    DocPath,
    Position,
)
from libreoffice_mcp.tools.writer.blocks import ContentBlock, dump_blocks

W = "writer"

FontSize = Annotated[float | None, Field(gt=0, le=1000, description="Font size in points.")]
EndBlock = Annotated[
    int | None,
    Field(ge=0, description="Last block of a range (inclusive). Omit for a single block."),
]


# -- adding --------------------------------------------------------------------


@tool(W, "edit", title="Add paragraph")
async def writer_add_paragraph(
    path: DocPath,
    text: str,
    style: Annotated[
        str | None, Field(description="Paragraph style, e.g. 'Title', 'Quotations'.")
    ] = None,
    alignment: Alignment | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font: str | None = None,
    size: FontSize = None,
    color: Color | None = None,
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Add one paragraph (appended at the end by default). Returns its block_index."""
    return await edit(
        path, "writer.add_paragraph", text=text, style=style, alignment=alignment, bold=bold,
        italic=italic, underline=underline, font=font, size=size, color=color,
        position=position, block_index=block_index,
    )  # fmt: skip


@tool(W, "edit", title="Add heading")
async def writer_add_heading(
    path: DocPath,
    text: str,
    level: Annotated[int, Field(ge=1, le=9, description="1 = top-level heading.")] = 1,
    alignment: Alignment | None = None,
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Add a heading (style 'Heading N'). Headings feed the outline and table of contents."""
    return await edit(
        path, "writer.add_heading", text=text, level=level, alignment=alignment,
        position=position, block_index=block_index,
    )  # fmt: skip


@tool(W, "edit", title="Add list")
async def writer_add_list(
    path: DocPath,
    items: Annotated[
        list[str],
        Field(
            min_length=1, description="List item texts (use writer_add_content for nested items)."
        ),
    ],
    ordered: Annotated[bool, Field(description="True for 1. 2. 3., false for bullets.")] = False,
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Add a bulleted or numbered list."""
    return await edit(
        path,
        "writer.add_list",
        items=items,
        ordered=ordered,
        position=position,
        block_index=block_index,
    )


@tool(W, "edit", title="Add content (many blocks)")
async def writer_add_content(
    path: DocPath,
    blocks: Annotated[list[ContentBlock], Field(min_length=1, max_length=200)],
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Build a whole section in ONE call: a list of blocks inserted in order. Each block has a
    `type`: heading, paragraph, bullets, numbered, table, image or page_break. This is the
    fastest way to write a document - prefer it over many single-paragraph calls.
    Example: [{"type":"heading","text":"Report","level":1},
    {"type":"paragraph","text":"Intro..."},
    {"type":"table","data":[["Name","Qty"],["Apple","3"]]}]"""
    return await edit(
        path,
        "writer.add_content",
        blocks=dump_blocks(blocks),
        position=position,
        block_index=block_index,
    )


@tool(W, "edit", title="Insert page break")
async def writer_insert_page_break(
    path: DocPath, position: Position = "end", block_index: AnchorBlock = None
) -> dict[str, Any]:
    """Start a new page at the given position."""
    return await edit(path, "writer.insert_page_break", position=position, block_index=block_index)


@tool(W, "edit", title="Insert another document")
async def writer_insert_file(
    path: DocPath,
    source_path: Annotated[
        str, Field(description="Document to insert (docx, odt, rtf, txt, html...).")
    ],
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Insert the contents of another document file (merging documents)."""
    return await edit(
        path,
        "writer.insert_file",
        source_path=source_path,
        position=position,
        block_index=block_index,
    )


# -- changing text -----------------------------------------------------------------


@tool(W, "edit", title="Edit paragraph text")
async def writer_edit_paragraph(
    path: DocPath, block_index: BlockIndex, text: str
) -> dict[str, Any]:
    """Replace the whole text of one paragraph, keeping its style and paragraph formatting.
    (Character formatting inside the paragraph, such as a bold word, is lost.)"""
    return await edit(path, "writer.edit_paragraph", block_index=block_index, text=text)


@tool(W, "destructive", title="Replace text")
async def writer_replace_text(
    path: DocPath,
    search: str,
    replace: Annotated[str, Field(description="Replacement. Use '' to delete the found text.")],
    regex: Annotated[
        bool, Field(description="Search is a regular expression; $1 in replace = group 1.")
    ] = False,
    case_sensitive: bool = False,
    whole_word: bool = False,
    first_only: bool = False,
) -> dict[str, Any]:
    """Find and replace text everywhere in the document (body, tables, headers). Returns the
    number of replacements; 0 means nothing matched and the file was not changed."""
    return await edit(
        path, "writer.replace_text", search=search, replace=replace, regex=regex,
        case_sensitive=case_sensitive, whole_word=whole_word, first_only=first_only,
    )  # fmt: skip


@tool(W, "destructive", title="Delete blocks")
async def writer_delete_blocks(
    path: DocPath, start_block: BlockIndex, end_block: EndBlock = None
) -> dict[str, Any]:
    """Delete paragraphs and/or tables from start_block to end_block (inclusive). Indexes of
    later blocks shift down, so re-run writer_get_structure afterwards."""
    return await edit(path, "writer.delete_blocks", start_block=start_block, end_block=end_block)


@tool(W, "edit", title="Set list formatting")
async def writer_set_list(
    path: DocPath,
    block_index: BlockIndex,
    end_block_index: EndBlock = None,
    kind: Literal["bullets", "numbered", "none"] = "bullets",
    level: Annotated[int, Field(ge=0, le=8)] = 0,
) -> dict[str, Any]:
    """Turn existing paragraphs into a bullet/numbered list, or remove list formatting (kind='none')."""
    return await edit(
        path, "writer.set_list", block_index=block_index, end_block_index=end_block_index,
        kind=kind, level=level,
    )  # fmt: skip


# -- formatting ------------------------------------------------------------------


@tool(W, "edit", title="Format text")
async def writer_format_text(
    path: DocPath,
    find: Annotated[
        str | None, Field(description="Format every occurrence of this text...")
    ] = None,
    block_index: Annotated[
        int | None, Field(ge=0, description="...or format whole paragraph(s) from here.")
    ] = None,
    end_block_index: EndBlock = None,
    occurrence: Literal["all", "first"] = "all",
    case_sensitive: bool = False,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    color: Color | None = None,
    background_color: Color | None = None,
    font: str | None = None,
    size: FontSize = None,
    superscript: bool | None = None,
    subscript: bool | None = None,
) -> dict[str, Any]:
    """Character formatting (bold, italic, colour, font, size...). Target either specific text
    (`find`, optionally limited to block_index..end_block_index) or whole paragraphs
    (`block_index` alone). Pass false to switch a property off."""
    return await edit(
        path, "writer.format_text", find=find, block_index=block_index, end_block_index=end_block_index,
        occurrence=occurrence, case_sensitive=case_sensitive, bold=bold, italic=italic,
        underline=underline, strikethrough=strikethrough, color=color, background_color=background_color,
        font=font, size=size, superscript=superscript, subscript=subscript,
    )  # fmt: skip


@tool(W, "edit", title="Format paragraph")
async def writer_format_paragraph(
    path: DocPath,
    block_index: BlockIndex,
    end_block_index: EndBlock = None,
    style: Annotated[
        str | None, Field(description="Apply this paragraph style, e.g. 'Heading 2'.")
    ] = None,
    alignment: Alignment | None = None,
    line_spacing_percent: Annotated[
        float | None, Field(ge=50, le=500, description="100 = single, 150 = 1.5 lines.")
    ] = None,
    space_before_mm: Annotated[float | None, Field(ge=0)] = None,
    space_after_mm: Annotated[float | None, Field(ge=0)] = None,
    left_indent_mm: float | None = None,
    right_indent_mm: float | None = None,
    first_line_indent_mm: float | None = None,
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    background_color: Color | None = None,
) -> dict[str, Any]:
    """Paragraph-level formatting for one paragraph or a range: style, alignment, spacing, indents."""
    return await edit(
        path, "writer.format_paragraph", block_index=block_index, end_block_index=end_block_index,
        style=style, alignment=alignment, line_spacing_percent=line_spacing_percent,
        space_before_mm=space_before_mm, space_after_mm=space_after_mm, left_indent_mm=left_indent_mm,
        right_indent_mm=right_indent_mm, first_line_indent_mm=first_line_indent_mm,
        keep_with_next=keep_with_next, page_break_before=page_break_before, background_color=background_color,
    )  # fmt: skip


@tool(W, "destructive", title="Clear formatting")
async def writer_clear_formatting(
    path: DocPath, block_index: BlockIndex, end_block_index: EndBlock = None
) -> dict[str, Any]:
    """Remove direct formatting (bold, colours, alignment, list membership) so the paragraph
    style alone decides how the paragraph looks. The style itself is kept."""
    return await edit(
        path, "writer.clear_formatting", block_index=block_index, end_block_index=end_block_index
    )


@tool(W, "edit", title="Create or modify style")
async def writer_set_style(
    path: DocPath,
    name: Annotated[
        str, Field(description="Style name. Existing styles are updated, new names are created.")
    ],
    family: Literal["paragraph", "character"] = "paragraph",
    parent: Annotated[
        str | None, Field(description="Style to inherit from, e.g. 'Standard'.")
    ] = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: Color | None = None,
    background_color: Color | None = None,
    font: str | None = None,
    size: FontSize = None,
    alignment: Alignment | None = None,
    line_spacing_percent: float | None = None,
    space_before_mm: float | None = None,
    space_after_mm: float | None = None,
    left_indent_mm: float | None = None,
    first_line_indent_mm: float | None = None,
) -> dict[str, Any]:
    """Define a reusable style (e.g. change what 'Heading 1' looks like for the whole
    document, or create a 'Callout' style). Apply it with writer_format_paragraph(style=...)."""
    return await edit(
        path, "writer.set_style", name=name, family=family, parent=parent, bold=bold, italic=italic,
        underline=underline, color=color, background_color=background_color, font=font, size=size,
        alignment=alignment, line_spacing_percent=line_spacing_percent, space_before_mm=space_before_mm,
        space_after_mm=space_after_mm, left_indent_mm=left_indent_mm, first_line_indent_mm=first_line_indent_mm,
    )  # fmt: skip
