"""Impress tools that create and arrange slides and set slide-level properties."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.blocks import SlideBlock, dump_slides
from libreoffice_mcp.tools.impress.types import (
    Bullets,
    I,
    OptionalSlideIndex,
    SlideIndex,
    SlideLayout,
    TransitionEffect,
    plain_items,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import Color, DocPath

AtIndex = Annotated[
    int | None,
    Field(ge=0, description="Insert before the slide with this index. Omit to add at the end."),
]


@tool(I, "edit", title="Add slides (many at once)")
async def impress_add_slides(
    path: DocPath,
    slides: Annotated[list[SlideBlock], Field(min_length=1, max_length=100)],
    at_index: AtIndex = None,
) -> dict[str, Any]:
    """Build many slides in ONE call - the fastest way to create a deck; prefer it over many
    impress_add_slide calls. Each slide is an object with optional fields: layout, title,
    subtitle, bullets (strings or {text, level} for nested points), right_bullets, image
    ({image_path, alt_text}), table ({data: rows, header}), notes (speaker notes),
    background_color. A brand-new presentation's empty first slide is reused for the first
    slide. Returns the index of each new slide.
    Example: [{"title": "Q3 Review", "subtitle": "Finance team"},
    {"title": "Highlights", "bullets": ["Revenue up 12%", {"text": "EU +18%", "level": 1}],
    "notes": "Mention the new customers"},
    {"title": "Numbers", "table": {"data": [["Region", "Sales"], ["EU", 10], ["US", 20]]}}]"""
    return await edit(path, "impress.add_slides", slides=dump_slides(slides), at_index=at_index)


@tool(I, "edit", title="Add slide")
async def impress_add_slide(
    path: DocPath,
    layout: Annotated[
        SlideLayout | None,
        Field(
            description="Layout; omitted: title_slide if subtitle is given, two_content if "
            "right_bullets is given, title_content if bullets are given, else title_only."
        ),
    ] = None,
    title: str | None = None,
    subtitle: Annotated[str | None, Field(description="Layout 'title_slide' only.")] = None,
    bullets: Bullets | None = None,
    right_bullets: Annotated[
        Bullets | None, Field(description="Second content area of two-column layouts.")
    ] = None,
    notes: Annotated[str | None, Field(description="Speaker notes.")] = None,
    at_index: AtIndex = None,
    background_color: Color | None = None,
) -> dict[str, Any]:
    """Add one slide (at the end by default) and fill its placeholders. For several slides use
    impress_add_slides. Layouts: title_slide, title_content, two_content, title_only, blank,
    centered_text, content_over_content, content_left_two_right, two_left_content_right,
    four_content, six_content. Pictures, tables and shapes are added afterwards (impress_add_image,
    impress_add_table, impress_add_shape) or via impress_add_slides."""
    return await edit(
        path, "impress.add_slide", layout=layout, title=title, subtitle=subtitle,
        bullets=plain_items(bullets), right_bullets=plain_items(right_bullets), notes=notes,
        at_index=at_index, background_color=background_color,
    )  # fmt: skip


@tool(I, "destructive", title="Delete slides")
async def impress_delete_slides(
    path: DocPath,
    start_slide: SlideIndex,
    end_slide: Annotated[
        int | None, Field(ge=0, description="Last slide to delete (inclusive); omit for one.")
    ] = None,
) -> dict[str, Any]:
    """Delete one slide or a range of slides. At least one slide must remain. Later slides
    move up: their indexes change."""
    return await edit(path, "impress.delete_slides", start_slide=start_slide, end_slide=end_slide)


@tool(I, "edit", title="Duplicate slide")
async def impress_duplicate_slide(
    path: DocPath,
    slide_index: SlideIndex,
    count: Annotated[int, Field(ge=1, le=50, description="Number of copies.")] = 1,
) -> dict[str, Any]:
    """Copy a slide with all its content, notes, background and transition. The copies are
    inserted directly after the original; their indexes are returned."""
    return await edit(path, "impress.duplicate_slide", slide_index=slide_index, count=count)


@tool(I, "edit", title="Move slide")
async def impress_move_slide(
    path: DocPath,
    from_index: SlideIndex,
    to_index: Annotated[
        int, Field(ge=0, description="Index the slide should have afterwards (others shift).")
    ],
) -> dict[str, Any]:
    """Reorder: move a slide so it ends up at to_index. Shapes, notes, background, transition
    and footer settings move with it; slide animations are not preserved."""
    return await edit(path, "impress.move_slide", from_index=from_index, to_index=to_index)


@tool(I, "edit", title="Set slide layout")
async def impress_set_slide_layout(
    path: DocPath, slide_index: SlideIndex, layout: SlideLayout
) -> dict[str, Any]:
    """Change a slide's layout. Existing text stays in matching placeholders. Returns the
    slide's shapes afterwards (their indexes may have changed)."""
    return await edit(path, "impress.set_slide_layout", slide_index=slide_index, layout=layout)


@tool(I, "edit", title="Set slide size")
async def impress_set_slide_size(
    path: DocPath,
    preset: Annotated[
        Literal["16:9", "16:10", "4:3", "a4_landscape", "letter_landscape"] | None,
        Field(description="Standard size. Or give both width_mm and height_mm."),
    ] = None,
    width_mm: Annotated[float | None, Field(ge=20, le=5000)] = None,
    height_mm: Annotated[float | None, Field(ge=20, le=5000)] = None,
    scale_content: Annotated[
        bool, Field(description="Scale shapes and masters proportionally (recommended).")
    ] = True,
) -> dict[str, Any]:
    """Change the slide size of the whole presentation. New presentations are 16:9
    (280 x 157.5 mm). Font sizes are not scaled."""
    return await edit(
        path, "impress.set_slide_size", preset=preset, width_mm=width_mm, height_mm=height_mm,
        scale_content=scale_content,
    )  # fmt: skip


@tool(I, "edit", title="Hide or show slide")
async def impress_set_slide_hidden(
    path: DocPath, slide_index: SlideIndex, hidden: bool = True
) -> dict[str, Any]:
    """Hide a slide from the slide show (it stays in the file), or unhide it (hidden=false)."""
    return await edit(path, "impress.set_slide_hidden", slide_index=slide_index, hidden=hidden)


@tool(I, "edit", title="Set slide background")
async def impress_set_slide_background(
    path: DocPath,
    color: Annotated[
        Color | None, Field(description="Background colour (omit with reset).")
    ] = None,
    slide_index: OptionalSlideIndex = None,
    gradient_end_color: Annotated[
        Color | None, Field(description="Second colour: makes a top-to-bottom gradient.")
    ] = None,
    reset: Annotated[
        bool, Field(description="Remove the slide's own background (use the master's).")
    ] = False,
) -> dict[str, Any]:
    """Set the background of one slide, or of all slides when slide_index is omitted. To
    change the look of the whole deck including new slides use impress_set_master_style."""
    return await edit(
        path, "impress.set_slide_background", color=color, slide_index=slide_index,
        gradient_end_color=gradient_end_color, reset=reset,
    )  # fmt: skip


@tool(I, "edit", title="Set slide transition")
async def impress_set_slide_transition(
    path: DocPath,
    effect: Annotated[TransitionEffect | None, Field(description="Transition effect.")] = None,
    slide_index: OptionalSlideIndex = None,
    direction: Annotated[
        Literal["from_left", "from_right", "from_top", "from_bottom"],
        Field(description="Direction for wipe, push and cover."),
    ] = "from_left",
    speed: Literal["slow", "medium", "fast"] = "medium",
    advance_after_seconds: Annotated[
        float | None,
        Field(ge=0, description="Auto-advance after this many seconds; 0 = advance on click."),
    ] = None,
) -> dict[str, Any]:
    """Set the slide-show transition (none, fade, dissolve, wipe, push, cover, random) of one
    slide or all slides, and optionally automatic advancing."""
    return await edit(
        path, "impress.set_slide_transition", effect=effect, slide_index=slide_index,
        direction=direction, speed=speed, advance_after_seconds=advance_after_seconds,
    )  # fmt: skip


@tool(I, "edit", title="Set footer, slide number and date")
async def impress_set_footer(
    path: DocPath,
    slide_index: OptionalSlideIndex = None,
    footer_text: Annotated[
        str | None, Field(description="Footer text; '' hides the footer; omit to leave as is.")
    ] = None,
    show_slide_number: bool | None = None,
    show_date: bool | None = None,
    date_text: Annotated[
        str | None, Field(description="Fixed date text, e.g. '2026-03-01' (also shows the date).")
    ] = None,
) -> dict[str, Any]:
    """Show or hide the footer text, slide number and date on one slide or all slides. They
    appear in the master's footer area."""
    return await edit(
        path, "impress.set_footer", slide_index=slide_index, footer_text=footer_text,
        show_slide_number=show_slide_number, show_date=show_date, date_text=date_text,
    )  # fmt: skip
