"""Writer tools that only read (never modify the file)."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import read
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import BlockIndex, DocPath, TableIndex

W = "writer"


@tool(W, "read", title="Get document text")
async def writer_get_text(
    path: DocPath,
    start_block: Annotated[int, Field(ge=0)] = 0,
    max_chars: Annotated[
        int, Field(ge=100, le=200_000, description="Page size in characters.")
    ] = 20_000,
) -> dict[str, Any]:
    """Read the document's text: one paragraph per line, tables as tab-separated rows.
    Long documents are paged: if `next_block` is not null, call again with
    start_block=next_block. Use writer_get_structure first to understand layout and styles."""
    return await read(path, "writer.get_text", start_block=start_block, max_chars=max_chars)


@tool(W, "read", title="Get document structure")
async def writer_get_structure(
    path: DocPath,
    start_block: Annotated[int, Field(ge=0)] = 0,
    limit: Annotated[int, Field(ge=1, le=500)] = 100,
) -> dict[str, Any]:
    """List the document body block by block: index, type (heading / paragraph / list_item /
    table), style and a text preview. The block `index` is how all other writer_* tools refer to
    a paragraph or table - call this before editing existing documents."""
    return await read(path, "writer.get_structure", start_block=start_block, limit=limit)


@tool(W, "read", title="Get outline")
async def writer_get_outline(path: DocPath) -> dict[str, Any]:
    """List the headings (level, text, block index): the document's table of contents."""
    return await read(path, "writer.get_outline")


@tool(W, "read", title="Get paragraph details")
async def writer_get_paragraph(path: DocPath, block_index: BlockIndex) -> dict[str, Any]:
    """Full text and formatting of one paragraph, including formatted runs (bold, italic,
    colour, links) - use it to inspect formatting before changing it."""
    return await read(path, "writer.get_paragraph", block_index=block_index)


@tool(W, "read", title="Find text")
async def writer_find_text(
    path: DocPath,
    query: str,
    regex: Annotated[
        bool, Field(description="Treat query as a (Python-style) regular expression.")
    ] = False,
    case_sensitive: bool = False,
    whole_word: bool = False,
    max_results: Annotated[int, Field(ge=1, le=500)] = 50,
) -> dict[str, Any]:
    """Search paragraphs and table cells. Returns each match with its block index, offset
    and surrounding context."""
    return await read(
        path, "writer.find_text", query=query, regex=regex, case_sensitive=case_sensitive,
        whole_word=whole_word, max_results=max_results,
    )  # fmt: skip


@tool(W, "read", title="List styles")
async def writer_list_styles(
    path: DocPath,
    family: Literal["paragraph", "character", "page", "list", "frame"] = "paragraph",
    contains: Annotated[str | None, Field(description="Only names containing this text.")] = None,
) -> dict[str, Any]:
    """List style names available in the document (e.g. 'Heading 1', 'Title', 'Quotations')."""
    return await read(path, "writer.list_styles", family=family, contains=contains)


@tool(W, "read", title="List tables")
async def writer_get_tables(path: DocPath) -> dict[str, Any]:
    """List the tables: table_index, block index, size and a preview of the first row."""
    return await read(path, "writer.get_tables")


@tool(W, "read", title="Get table")
async def writer_get_table(
    path: DocPath, table_index: TableIndex, max_rows: Annotated[int, Field(ge=1, le=1000)] = 200
) -> dict[str, Any]:
    """Read all cell text of one table as a row-major 2-D array."""
    return await read(path, "writer.get_table", table_index=table_index, max_rows=max_rows)


@tool(W, "read", title="List images")
async def writer_get_images(path: DocPath) -> dict[str, Any]:
    """List images: index, size in mm and alt text."""
    return await read(path, "writer.get_images")


@tool(W, "read", title="List comments")
async def writer_get_comments(path: DocPath) -> dict[str, Any]:
    """List review comments: author, date, text and the text each is attached to."""
    return await read(path, "writer.get_comments")


@tool(W, "read", title="List bookmarks")
async def writer_get_bookmarks(path: DocPath) -> dict[str, Any]:
    """List bookmarks and the text they mark."""
    return await read(path, "writer.get_bookmarks")


@tool(W, "read", title="List hyperlinks")
async def writer_get_hyperlinks(path: DocPath) -> dict[str, Any]:
    """List hyperlinks: linked text, URL and block index."""
    return await read(path, "writer.get_hyperlinks")


@tool(W, "read", title="List footnotes")
async def writer_get_footnotes(path: DocPath) -> dict[str, Any]:
    """List footnotes and their text."""
    return await read(path, "writer.get_footnotes")


@tool(W, "read", title="Get page setup")
async def writer_get_page_setup(path: DocPath) -> dict[str, Any]:
    """Paper size, orientation and margins (mm)."""
    return await read(path, "writer.get_page_setup")


@tool(W, "read", title="Get headers and footers")
async def writer_get_headers_footers(path: DocPath) -> dict[str, Any]:
    """Header and footer text for each page style in use (null = none)."""
    return await read(path, "writer.get_headers_footers")


@tool(W, "read", title="Get tracked changes")
async def writer_get_track_changes(path: DocPath) -> dict[str, Any]:
    """Whether change tracking is on, and the recorded changes (type, author, date)."""
    return await read(path, "writer.get_track_changes")
