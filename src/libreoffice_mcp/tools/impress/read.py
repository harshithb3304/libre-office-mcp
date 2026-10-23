"""Impress tools that only read (never modify the file)."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from libreoffice_mcp.tools.base import read
from libreoffice_mcp.tools.impress.types import I, ShapeIndex, SlideIndex
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath


@tool(I, "read", title="Get presentation overview")
async def impress_get_overview(path: DocPath) -> dict[str, Any]:
    """Overview of a presentation: slide count, slide size in mm, master slides, and one entry
    per slide (index, title, layout, shape count, hidden / has_notes flags). Call this first
    when working on an existing deck: `slide_index` everywhere else is the number listed here."""
    return await read(path, "impress.get_overview")


@tool(I, "read", title="Get slide")
async def impress_get_slide(
    path: DocPath,
    slide_index: SlideIndex,
    text_chars: Annotated[
        int, Field(ge=20, le=20_000, description="Truncate each shape's text to this length.")
    ] = 300,
) -> dict[str, Any]:
    """List everything on one slide: layout, master, background, transition, speaker notes
    and every shape with its `index` (the `shape_index` the editing tools take), type
    (title, subtitle, body, text_box, shape, line, image, table...), position and size in mm,
    text, fill/line colours, link and alt text. Empty placeholders are flagged `empty`; body
    placeholders carry the `content_index` that impress_set_body takes."""
    return await read(path, "impress.get_slide", slide_index=slide_index, text_chars=text_chars)


@tool(I, "read", title="Get shape details")
async def impress_get_shape(
    path: DocPath, slide_index: SlideIndex, shape_index: ShapeIndex
) -> dict[str, Any]:
    """Full detail of one shape: complete text split into paragraphs with level, font, size,
    bold/italic/underline, colour, alignment and bullet flag; fill, outline and shadow; for
    tables the full cell grid. Use it to inspect formatting before changing it."""
    return await read(path, "impress.get_shape", slide_index=slide_index, shape_index=shape_index)


@tool(I, "read", title="Get presentation text")
async def impress_get_text(
    path: DocPath,
    start_slide: Annotated[int, Field(ge=0)] = 0,
    max_chars: Annotated[
        int, Field(ge=100, le=200_000, description="Page size in characters.")
    ] = 20_000,
    include_notes: Annotated[bool, Field(description="Include speaker notes.")] = True,
) -> dict[str, Any]:
    """Read the deck as a text outline: per slide its title, bullets (indented by level), other
    text, tables and speaker notes. Long decks are paged: if `next_start_slide` is present,
    call again with start_slide=next_start_slide."""
    return await read(
        path, "impress.get_text", start_slide=start_slide, max_chars=max_chars,
        include_notes=include_notes,
    )  # fmt: skip


@tool(I, "read", title="Get speaker notes")
async def impress_get_notes(
    path: DocPath,
    slide_index: Annotated[
        int | None, Field(ge=0, description="One slide; omit to list all slides with notes.")
    ] = None,
) -> dict[str, Any]:
    """Read speaker notes of one slide, or of every slide that has notes."""
    return await read(path, "impress.get_notes", slide_index=slide_index)


@tool(I, "read", title="Find text in slides")
async def impress_find_text(
    path: DocPath,
    query: str,
    case_sensitive: bool = False,
    include_notes: bool = True,
) -> dict[str, Any]:
    """Find text in the presentation (shapes, table cells, speaker notes). Returns slide_index,
    shape_index and a snippet per match (first 100 matches)."""
    return await read(
        path, "impress.find_text", query=query, case_sensitive=case_sensitive,
        include_notes=include_notes,
    )  # fmt: skip


@tool(I, "read", title="List master slides and layouts")
async def impress_list_masters(path: DocPath) -> dict[str, Any]:
    """List the master slides (name, how many slides use it, background colour, title and body
    fonts/sizes/colours), the slide layouts you can use (with descriptions) and the built-in
    themes for impress_apply_theme."""
    return await read(path, "impress.list_masters")
