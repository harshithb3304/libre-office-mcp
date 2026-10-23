"""Impress tools for text formatting, shape styling and hyperlinks."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.types import FontSize, I, ShapeIndex, SlideIndex
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Alignment, Color, DocPath


@tool(I, "edit", title="Format text")
async def impress_format_text(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    find: Annotated[
        str | None,
        Field(
            description="Format only this text (every occurrence, case-sensitive). Allows "
            "character options only. Omit to format whole paragraphs."
        ),
    ] = None,
    paragraph_index: Annotated[
        int | None,
        Field(ge=0, description="Format only this paragraph (0-based); omit for all text."),
    ] = None,
    font: Annotated[str | None, Field(description="Font family, e.g. 'Calibri', 'Arial'.")] = None,
    size: FontSize = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    color: Annotated[Color | None, Field(description="Text colour.")] = None,
    highlight_color: Annotated[Color | None, Field(description="Text background colour.")] = None,
    alignment: Alignment | None = None,
    line_spacing_percent: Annotated[
        float | None, Field(ge=50, le=400, description="100 = single, 150 = one and a half.")
    ] = None,
    space_before_mm: Annotated[float | None, Field(ge=0, le=100)] = None,
    space_after_mm: Annotated[float | None, Field(ge=0, le=100)] = None,
    bullets: Annotated[
        Literal["bullets", "numbered", "none"] | None,
        Field(description="Turn paragraphs into a bulleted / numbered list, or remove the bullet."),
    ] = None,
    level: Annotated[
        int | None, Field(ge=0, le=8, description="Nesting level of the paragraph(s).")
    ] = None,
) -> dict[str, Any]:
    """Format text in a shape: all of it, one paragraph, or specific words (`find`). Options
    left out are not changed; booleans can switch a style off (bold=false). Works on
    placeholders, text boxes and shapes with text; for tables use impress_set_table_cells."""
    return await edit(
        path, "impress.format_text", slide_index=slide_index, shape_index=shape_index, find=find,
        paragraph_index=paragraph_index, font=font, size=size, bold=bold, italic=italic,
        underline=underline, strikethrough=strikethrough, color=color,
        highlight_color=highlight_color, alignment=alignment,
        line_spacing_percent=line_spacing_percent, space_before_mm=space_before_mm,
        space_after_mm=space_after_mm, bullets=bullets, level=level,
    )  # fmt: skip


@tool(I, "edit", title="Format shape (fill, outline, shadow)")
async def impress_format_shape(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    fill_color: Color | None = None,
    no_fill: Annotated[bool, Field(description="Make the shape transparent (no fill).")] = False,
    gradient_end_color: Annotated[
        Color | None, Field(description="With fill_color: gradient from fill_color to this.")
    ] = None,
    gradient_angle: Annotated[
        int | None, Field(ge=0, le=359, description="Gradient direction in degrees.")
    ] = None,
    transparency: Annotated[
        int | None, Field(ge=0, le=100, description="Fill transparency in percent.")
    ] = None,
    line_color: Color | None = None,
    line_width_mm: Annotated[float | None, Field(ge=0, le=100)] = None,
    no_line: Annotated[bool, Field(description="Remove the outline.")] = False,
    dash: Literal["solid", "dashed", "dotted"] | None = None,
    shadow: bool | None = None,
    corner_radius_mm: Annotated[
        float | None, Field(ge=0, le=500, description="Rounded corners (plain rectangles).")
    ] = None,
    text_anchor: Annotated[
        Literal["top", "middle", "bottom"] | None,
        Field(description="Vertical position of the text inside the shape."),
    ] = None,
    margin_mm: Annotated[
        float | None, Field(ge=0, le=100, description="Inner padding between text and edge.")
    ] = None,
) -> dict[str, Any]:
    """Style a shape, text box or picture: fill colour/gradient/transparency, outline colour,
    width and dash, shadow, corner radius, text anchoring and inner margin. Only the options you
    give are changed. (Use impress_format_text for the text's own font and colour.)"""
    return await edit(
        path, "impress.format_shape", slide_index=slide_index, shape_index=shape_index,
        fill_color=fill_color, no_fill=no_fill, gradient_end_color=gradient_end_color,
        gradient_angle=gradient_angle, transparency=transparency, line_color=line_color,
        line_width_mm=line_width_mm, no_line=no_line, dash=dash, shadow=shadow,
        corner_radius_mm=corner_radius_mm, text_anchor=text_anchor, margin_mm=margin_mm,
    )  # fmt: skip


@tool(I, "edit", title="Add hyperlink")
async def impress_set_hyperlink(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    url: Annotated[
        str, Field(description="Web address (https://...) or file path; '' removes a shape link.")
    ],
    find: Annotated[
        str | None,
        Field(
            description="Turn this text inside the shape into a link (every occurrence, "
            "case-sensitive). Omit to make the whole shape clickable."
        ),
    ] = None,
) -> dict[str, Any]:
    """Add a hyperlink: either on specific text in a shape (`find`) or on the whole shape
    (clickable in slide show mode)."""
    return await edit(
        path, "impress.set_hyperlink", slide_index=slide_index, shape_index=shape_index, url=url,
        find=find,
    )  # fmt: skip
