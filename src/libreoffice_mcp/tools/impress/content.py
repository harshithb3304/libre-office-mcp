"""Impress tools that set text and speaker notes and add text boxes, shapes, lines, pictures."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.types import (
    Bullets,
    FontSize,
    Height,
    I,
    OptHeight,
    OptWidth,
    OptX,
    OptY,
    ShapeIndex,
    ShapeType,
    SlideIndex,
    Width,
    X,
    Y,
    plain_items,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Alignment, Color, DocPath

Text = Annotated[str | None, Field(description="The text. Use \\n for several paragraphs.")]
Paragraphs = Annotated[
    Bullets | None,
    Field(description="Alternative to `text`: paragraphs, optionally with nesting levels."),
]


@tool(I, "edit", title="Set slide title")
async def impress_set_title(
    path: DocPath,
    slide_index: SlideIndex,
    text: Annotated[str | None, Field(description="Title text.")] = None,
    subtitle: Annotated[
        str | None, Field(description="Subtitle text (slides with the title_slide layout).")
    ] = None,
) -> dict[str, Any]:
    """Set the text of a slide's title (and subtitle) placeholder. The slide's layout must
    have one - otherwise change it first with impress_set_slide_layout."""
    return await edit(
        path, "impress.set_title", slide_index=slide_index, text=text, subtitle=subtitle
    )


@tool(I, "edit", title="Set slide body text")
async def impress_set_body(
    path: DocPath,
    slide_index: SlideIndex,
    items: Bullets,
    content_index: Annotated[
        int,
        Field(ge=0, description="Which content area: 0 = first/left, 1 = second/right, ..."),
    ] = 0,
) -> dict[str, Any]:
    """Replace the bullets in a slide's content placeholder. Items are strings, or
    {"text": ..., "level": 1} for nested sub-bullets. Formatting comes from the master."""
    return await edit(
        path, "impress.set_body", slide_index=slide_index, items=plain_items(items),
        content_index=content_index,
    )  # fmt: skip


@tool(I, "edit", title="Set shape text")
async def impress_set_shape_text(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    text: Text = None,
    paragraphs: Paragraphs = None,
) -> dict[str, Any]:
    """Replace all text of any shape that holds text (placeholder, text box, shape). Give
    `text` or `paragraphs`. Run-level formatting is reset; use impress_format_text afterwards."""
    return await edit(
        path, "impress.set_shape_text", slide_index=slide_index, shape_index=shape_index,
        text=text, paragraphs=plain_items(paragraphs),
    )  # fmt: skip


@tool(I, "edit", title="Set speaker notes")
async def impress_set_notes(
    path: DocPath,
    slide_index: SlideIndex,
    text: Annotated[str, Field(description="Notes text; '' clears the notes.")],
) -> dict[str, Any]:
    """Set (replace) the speaker notes of a slide."""
    return await edit(path, "impress.set_notes", slide_index=slide_index, text=text)


@tool(I, "destructive", title="Replace text")
async def impress_replace_text(
    path: DocPath,
    search: str,
    replace: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    include_notes: bool = True,
) -> dict[str, Any]:
    """Find and replace text on all slides (titles, bullets, text boxes, shapes and speaker
    notes), keeping the formatting. Table cells are not searched: use impress_set_table_cells.
    Returns the number of replacements."""
    return await edit(
        path, "impress.replace_text", search=search, replace=replace,
        case_sensitive=case_sensitive, whole_word=whole_word,
        include_notes=include_notes,
    )  # fmt: skip


@tool(I, "edit", title="Add text box")
async def impress_add_text_box(
    path: DocPath,
    slide_index: SlideIndex,
    x_mm: X,
    y_mm: Y,
    width_mm: Width,
    height_mm: Annotated[float, Field(gt=0, le=5000, description="Minimum height in mm.")] = 20,
    text: Text = None,
    paragraphs: Paragraphs = None,
    font: str | None = None,
    font_size: FontSize = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: Color | None = None,
    alignment: Alignment | None = None,
    vertical_anchor: Literal["top", "middle", "bottom"] | None = None,
    bullets: Annotated[
        Literal["bullets", "numbered"] | None, Field(description="Turn paragraphs into a list.")
    ] = None,
    fill_color: Color | None = None,
    line_color: Color | None = None,
    name: Annotated[str | None, Field(description="Optional shape name.")] = None,
) -> dict[str, Any]:
    """Add a free text box at a position (mm from the slide's top-left corner; the default
    slide is 280 x 157.5 mm). Give `text` or `paragraphs`. The box grows downward if the text
    needs more room. Returns its shape_index and final size."""
    return await edit(
        path, "impress.add_text_box", slide_index=slide_index, x_mm=x_mm, y_mm=y_mm,
        width_mm=width_mm, height_mm=height_mm, text=text, paragraphs=plain_items(paragraphs),
        font=font, font_size=font_size, bold=bold, italic=italic, color=color,
        alignment=alignment, vertical_anchor=vertical_anchor, bullets=bullets,
        fill_color=fill_color, line_color=line_color, name=name,
    )  # fmt: skip


@tool(I, "edit", title="Add shape")
async def impress_add_shape(
    path: DocPath,
    slide_index: SlideIndex,
    shape: ShapeType,
    x_mm: X,
    y_mm: Y,
    width_mm: Width,
    height_mm: Height,
    text: Annotated[str | None, Field(description="Text inside the shape.")] = None,
    fill_color: Color | None = None,
    line_color: Color | None = None,
    line_width_mm: Annotated[float | None, Field(ge=0, le=100)] = None,
    no_line: Annotated[bool, Field(description="Remove the outline.")] = False,
    gradient_end_color: Annotated[
        Color | None, Field(description="With fill_color: makes a gradient to this colour.")
    ] = None,
    transparency: Annotated[
        int | None, Field(ge=0, le=100, description="Fill transparency in percent.")
    ] = None,
    shadow: bool | None = None,
    text_color: Color | None = None,
    font_size: FontSize = None,
    bold: bool | None = None,
    alignment: Annotated[
        Alignment, Field(description="Text alignment inside the shape.")
    ] = "center",
    name: Annotated[str | None, Field(description="Optional shape name.")] = None,
) -> dict[str, Any]:
    """Add a geometric shape (rectangle, rounded_rectangle, ellipse, triangle, diamond, star,
    arrows, chevron, callout, cylinder...) at a position in mm, optionally filled, outlined and
    with text. Colours are '#RRGGBB'. Returns its shape_index."""
    return await edit(
        path, "impress.add_shape", slide_index=slide_index, shape=shape, x_mm=x_mm, y_mm=y_mm,
        width_mm=width_mm, height_mm=height_mm, text=text, fill_color=fill_color,
        line_color=line_color, line_width_mm=line_width_mm, no_line=no_line,
        gradient_end_color=gradient_end_color, transparency=transparency, shadow=shadow,
        text_color=text_color, font_size=font_size, bold=bold, alignment=alignment, name=name,
    )  # fmt: skip


@tool(I, "edit", title="Add line or arrow")
async def impress_add_line(
    path: DocPath,
    slide_index: SlideIndex,
    x1_mm: X,
    y1_mm: Y,
    x2_mm: X,
    y2_mm: Y,
    color: Color = "#333333",
    width_mm: Annotated[float, Field(gt=0, le=50)] = 0.6,
    dash: Literal["solid", "dashed", "dotted"] = "solid",
    arrow_start: bool = False,
    arrow_end: bool = False,
    name: str | None = None,
) -> dict[str, Any]:
    """Add a straight line between two points (mm), optionally with arrow heads."""
    return await edit(
        path, "impress.add_line", slide_index=slide_index, x1_mm=x1_mm, y1_mm=y1_mm, x2_mm=x2_mm,
        y2_mm=y2_mm, color=color, width_mm=width_mm, dash=dash, arrow_start=arrow_start,
        arrow_end=arrow_end, name=name,
    )  # fmt: skip


@tool(I, "edit", title="Add image")
async def impress_add_image(
    path: DocPath,
    slide_index: SlideIndex,
    image_path: Annotated[str, Field(description="Path to a PNG/JPG/GIF/SVG file on disk.")],
    x_mm: OptX = None,
    y_mm: OptY = None,
    width_mm: OptWidth = None,
    height_mm: OptHeight = None,
    alt_text: Annotated[str | None, Field(description="Accessibility description.")] = None,
    name: str | None = None,
) -> dict[str, Any]:
    """Embed a picture. With no size it is fitted into the slide's content area (below the
    title) keeping its aspect ratio; with only a width or only a height the other follows.
    Without x/y it is centred in the content area. Returns its shape_index and final geometry."""
    return await edit(
        path, "impress.add_image", slide_index=slide_index, image_path=image_path, x_mm=x_mm,
        y_mm=y_mm, width_mm=width_mm, height_mm=height_mm, alt_text=alt_text, name=name,
    )  # fmt: skip
