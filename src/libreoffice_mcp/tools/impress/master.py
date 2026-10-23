"""Impress tools for master slides: the look of the whole deck."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.types import FontSize, I, OptionalSlideIndex, ThemeName
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Color, DocPath

MasterName = Annotated[
    str | None,
    Field(description="Master slide name (impress_list_masters); default: ALL master slides."),
]


@tool(I, "edit", title="Set master slide style")
async def impress_set_master_style(
    path: DocPath,
    master_name: MasterName = None,
    background: Annotated[
        Color | None, Field(description="Background colour of all slides.")
    ] = None,
    title_font: str | None = None,
    title_size: FontSize = None,
    title_color: Color | None = None,
    title_bold: bool | None = None,
    body_font: str | None = None,
    body_size: Annotated[
        float | None,
        Field(gt=0, le=200, description="Level-1 body size in points; deeper levels shrink."),
    ] = None,
    body_color: Color | None = None,
) -> dict[str, Any]:
    """Re-theme the whole deck in one call by changing the master slide: background colour,
    title font/size/colour/bold and body (bullet) font/size/colour. All slides using the master
    follow, including slides added later; text that has its own direct formatting and slides
    with their own background do not."""
    return await edit(
        path, "impress.set_master_style", master_name=master_name, background=background,
        title_font=title_font, title_size=title_size, title_color=title_color,
        title_bold=title_bold, body_font=body_font, body_size=body_size, body_color=body_color,
    )  # fmt: skip


@tool(I, "edit", title="Apply built-in theme")
async def impress_apply_theme(
    path: DocPath, theme: ThemeName, master_name: MasterName = None
) -> dict[str, Any]:
    """Apply a ready-made look to the deck's master: 'clean' (white, dark grey), 'corporate'
    (white, navy titles), 'dark' (near-black background, white text), 'warm' (cream background,
    brown serif titles) or 'modern' (light grey, teal titles). Fine-tune afterwards with
    impress_set_master_style."""
    return await edit(path, "impress.apply_theme", theme=theme, master_name=master_name)


@tool(I, "edit", title="Assign master slide")
async def impress_set_slide_master(
    path: DocPath,
    master_name: Annotated[str, Field(description="Name from impress_list_masters.")],
    slide_index: OptionalSlideIndex = None,
) -> dict[str, Any]:
    """Make one slide (or all slides) use a different master slide."""
    return await edit(
        path, "impress.set_slide_master", master_name=master_name, slide_index=slide_index
    )
