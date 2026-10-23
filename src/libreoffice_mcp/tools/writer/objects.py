"""Writer tools for images, links, notes, comments, page layout and change tracking."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Alignment, AnchorBlock, DocPath, Position

W = "writer"

Find = Annotated[str | None, Field(description="Text to attach to (first occurrence)...")]
InBlock = Annotated[int | None, Field(ge=0, description="...or a whole paragraph by block index.")]
ImageIndex = Annotated[int, Field(ge=0, description="0-based image number from writer_get_images.")]
Millimetres = Annotated[float | None, Field(gt=0, le=2000)]


# -- images --------------------------------------------------------------------


@tool(W, "edit", title="Insert image")
async def writer_insert_image(
    path: DocPath,
    image_path: Annotated[str, Field(description="Path to a PNG/JPG/GIF/SVG... file on disk.")],
    width_mm: Millimetres = None,
    height_mm: Millimetres = None,
    alt_text: Annotated[str | None, Field(description="Accessibility description.")] = None,
    alignment: Alignment | None = None,
    position: Position = "end",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Embed an image in its own paragraph. Give one of width_mm/height_mm to keep the aspect
    ratio; with neither, it is sized to fit the page width."""
    return await edit(
        path, "writer.insert_image", image_path=image_path, width_mm=width_mm, height_mm=height_mm,
        alt_text=alt_text, alignment=alignment, position=position, block_index=block_index,
    )  # fmt: skip


@tool(W, "edit", title="Resize image / set alt text")
async def writer_set_image(
    path: DocPath,
    image_index: ImageIndex,
    width_mm: Millimetres = None,
    height_mm: Millimetres = None,
    alt_text: str | None = None,
) -> dict[str, Any]:
    """Resize an image (aspect ratio kept if only one size is given) and/or set its alt text."""
    return await edit(
        path,
        "writer.set_image",
        image_index=image_index,
        width_mm=width_mm,
        height_mm=height_mm,
        alt_text=alt_text,
    )


@tool(W, "destructive", title="Delete image")
async def writer_delete_image(path: DocPath, image_index: ImageIndex) -> dict[str, Any]:
    """Remove an image from the document."""
    return await edit(path, "writer.delete_image", image_index=image_index)


# -- links, bookmarks, notes, comments ------------------------------------------------


@tool(W, "edit", title="Add hyperlink")
async def writer_add_hyperlink(
    path: DocPath,
    url: str,
    find: Annotated[str | None, Field(description="Turn existing text into a link...")] = None,
    text: Annotated[str | None, Field(description="...or append this new linked text.")] = None,
    block_index: Annotated[
        int | None, Field(ge=0, description="Limit `find` to / append `text` in this paragraph.")
    ] = None,
    occurrence: Literal["first", "all"] = "first",
) -> dict[str, Any]:
    """Create a hyperlink on existing text (`find`) or add new linked `text` to a paragraph
    (or as a new last paragraph when block_index is omitted)."""
    return await edit(
        path,
        "writer.add_hyperlink",
        url=url,
        find=find,
        text=text,
        block_index=block_index,
        occurrence=occurrence,
    )


@tool(W, "edit", title="Add bookmark")
async def writer_add_bookmark(
    path: DocPath, name: str, find: Find = None, block_index: InBlock = None
) -> dict[str, Any]:
    """Name a location (text or paragraph) so it can be linked or found later."""
    return await edit(path, "writer.add_bookmark", name=name, find=find, block_index=block_index)


@tool(W, "destructive", title="Delete bookmark")
async def writer_delete_bookmark(path: DocPath, name: str) -> dict[str, Any]:
    """Delete a bookmark (the marked text stays)."""
    return await edit(path, "writer.delete_bookmark", name=name)


@tool(W, "edit", title="Add footnote")
async def writer_add_footnote(
    path: DocPath,
    note_text: str,
    find: Find = None,
    block_index: InBlock = None,
    endnote: Annotated[
        bool, Field(description="Place the note at the end of the document instead of the page.")
    ] = False,
) -> dict[str, Any]:
    """Add a footnote (or endnote) at the end of the found text / paragraph."""
    return await edit(
        path,
        "writer.add_footnote",
        note_text=note_text,
        find=find,
        block_index=block_index,
        endnote=endnote,
    )


@tool(W, "edit", title="Add comment")
async def writer_add_comment(
    path: DocPath,
    comment: str,
    find: Find = None,
    block_index: InBlock = None,
    author: str = "AI Assistant",
) -> dict[str, Any]:
    """Add a review comment (margin note) on the found text / paragraph."""
    return await edit(
        path,
        "writer.add_comment",
        comment=comment,
        find=find,
        block_index=block_index,
        author=author,
    )


@tool(W, "destructive", title="Delete comment")
async def writer_delete_comment(
    path: DocPath,
    comment_index: Annotated[int, Field(ge=0, description="From writer_get_comments.")],
) -> dict[str, Any]:
    """Delete one review comment."""
    return await edit(path, "writer.delete_comment", comment_index=comment_index)


# -- page layout -----------------------------------------------------------------


@tool(W, "edit", title="Set page setup")
async def writer_set_page_setup(
    path: DocPath,
    paper: Literal["A3", "A4", "A5", "Letter", "Legal"] | None = None,
    orientation: Literal["portrait", "landscape"] | None = None,
    top_mm: Annotated[float | None, Field(ge=0, le=200)] = None,
    bottom_mm: Annotated[float | None, Field(ge=0, le=200)] = None,
    left_mm: Annotated[float | None, Field(ge=0, le=200)] = None,
    right_mm: Annotated[float | None, Field(ge=0, le=200)] = None,
) -> dict[str, Any]:
    """Set paper size, orientation and margins for the whole document."""
    return await edit(
        path, "writer.set_page_setup", paper=paper, orientation=orientation, top_mm=top_mm,
        bottom_mm=bottom_mm, left_mm=left_mm, right_mm=right_mm,
    )  # fmt: skip


@tool(W, "edit", title="Set header and footer")
async def writer_set_header_footer(
    path: DocPath,
    header_text: Annotated[
        str | None,
        Field(description="Header text; '' removes the header; omit to leave unchanged."),
    ] = None,
    footer_text: Annotated[
        str | None, Field(description="Footer text; '' removes it; omit to leave unchanged.")
    ] = None,
    page_numbers: Annotated[
        Literal["none", "page", "page_of_total"],
        Field(description="Add 'Page N' or 'Page N of M' to the footer."),
    ] = "none",
    alignment: Alignment = "center",
) -> dict[str, Any]:
    """Set the page header/footer text and optional page numbers for all pages."""
    return await edit(
        path, "writer.set_header_footer", header_text=header_text, footer_text=footer_text,
        page_numbers=page_numbers, alignment=alignment,
    )  # fmt: skip


@tool(W, "edit", title="Insert table of contents")
async def writer_insert_toc(
    path: DocPath,
    title: str = "Table of Contents",
    levels: Annotated[
        int, Field(ge=1, le=10, description="How many heading levels to include.")
    ] = 3,
    position: Position = "start",
    block_index: AnchorBlock = None,
) -> dict[str, Any]:
    """Insert a table of contents built from the document's headings."""
    return await edit(
        path,
        "writer.insert_toc",
        title=title,
        levels=levels,
        position=position,
        block_index=block_index,
    )


# -- change tracking -------------------------------------------------------------


@tool(W, "edit", title="Turn change tracking on/off")
async def writer_set_track_changes(path: DocPath, enabled: bool) -> dict[str, Any]:
    """Enable or disable change tracking. While on, later edits are recorded as changes a
    human can review (accept/reject) in LibreOffice."""
    return await edit(path, "writer.set_track_changes", enabled=enabled)


@tool(W, "destructive", title="Accept all tracked changes")
async def writer_accept_all_changes(path: DocPath) -> dict[str, Any]:
    """Accept every tracked change, making the edits permanent."""
    return await edit(path, "writer.accept_all_changes")


@tool(W, "destructive", title="Reject all tracked changes")
async def writer_reject_all_changes(path: DocPath) -> dict[str, Any]:
    """Reject every tracked change, reverting the document to its pre-edit text."""
    return await edit(path, "writer.reject_all_changes")
