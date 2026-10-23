"""Impress tools that arrange shapes: move/resize/rotate, stacking order, delete, duplicate."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import Field

from libreoffice_mcp.tools.base import edit
from libreoffice_mcp.tools.impress.types import (
    I,
    OptHeight,
    OptWidth,
    OptX,
    OptY,
    ShapeIndex,
    SlideIndex,
)
from libreoffice_mcp.tools.registry import tool
from libreoffice_mcp.tools.types import DocPath


@tool(I, "edit", title="Move, resize or rotate shape")
async def impress_set_shape_geometry(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    x_mm: OptX = None,
    y_mm: OptY = None,
    width_mm: OptWidth = None,
    height_mm: OptHeight = None,
    rotation_deg: Annotated[
        float | None, Field(ge=-360, le=360, description="Rotation in degrees, clockwise.")
    ] = None,
    keep_aspect_ratio: Annotated[
        bool, Field(description="With only width or only height: scale the other to match.")
    ] = False,
) -> dict[str, Any]:
    """Change a shape's position (top-left corner, mm), size and/or rotation (about its
    centre). Only the values you give are changed."""
    return await edit(
        path, "impress.set_shape_geometry", slide_index=slide_index, shape_index=shape_index,
        x_mm=x_mm, y_mm=y_mm, width_mm=width_mm, height_mm=height_mm, rotation_deg=rotation_deg,
        keep_aspect_ratio=keep_aspect_ratio,
    )  # fmt: skip


@tool(I, "edit", title="Change stacking order")
async def impress_set_shape_order(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    action: Literal["front", "back", "forward", "backward"],
) -> dict[str, Any]:
    """Bring a shape to the front / send to the back, or move it one step. Shape indexes follow
    the stacking order (0 = bottom), so the shape's index changes: the new one is returned."""
    return await edit(
        path, "impress.set_shape_order", slide_index=slide_index, shape_index=shape_index,
        action=action,
    )  # fmt: skip


@tool(I, "destructive", title="Delete shapes")
async def impress_delete_shapes(
    path: DocPath,
    slide_index: SlideIndex,
    shape_indexes: Annotated[
        list[int],
        Field(min_length=1, max_length=200, description="Indexes as listed by impress_get_slide."),
    ],
) -> dict[str, Any]:
    """Delete one or more shapes (placeholders, text boxes, pictures, tables...). All indexes
    refer to the slide as it is now; remaining shapes are renumbered afterwards."""
    return await edit(
        path, "impress.delete_shapes", slide_index=slide_index, shape_indexes=shape_indexes
    )


@tool(I, "edit", title="Duplicate shape")
async def impress_duplicate_shape(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    target_slide_index: Annotated[
        int | None, Field(ge=0, description="Copy onto this slide; omit for the same slide.")
    ] = None,
    offset_x_mm: Annotated[
        float | None, Field(description="Shift of the copy (default 10 mm on the same slide).")
    ] = None,
    offset_y_mm: Annotated[float | None, Field(description="Vertical shift of the copy.")] = None,
    count: Annotated[
        int, Field(ge=1, le=50, description="Number of copies, each shifted once more.")
    ] = 1,
) -> dict[str, Any]:
    """Copy a shape with all its formatting, on the same slide (offset so it is visible) or onto
    another slide (same position). Returns the new shape indexes."""
    return await edit(
        path, "impress.duplicate_shape", slide_index=slide_index, shape_index=shape_index,
        target_slide_index=target_slide_index, offset_x_mm=offset_x_mm, offset_y_mm=offset_y_mm,
        count=count,
    )  # fmt: skip


@tool(I, "edit", title="Align or distribute shapes")
async def impress_align_shapes(
    path: DocPath,
    slide_index: SlideIndex,
    shape_indexes: Annotated[list[int], Field(min_length=1, max_length=200)],
    align: Annotated[
        Literal["left", "center", "right", "top", "middle", "bottom"] | None,
        Field(description="left/center/right align horizontally; top/middle/bottom vertically."),
    ] = None,
    distribute: Annotated[
        Literal["horizontal", "vertical"] | None,
        Field(description="Even spacing between at least three shapes."),
    ] = None,
    relative_to: Annotated[
        Literal["selection", "slide"],
        Field(description="Align to the group's bounding box or to the slide."),
    ] = "selection",
) -> dict[str, Any]:
    """Align shapes to each other (or to the slide edges/centre) and/or distribute them evenly.
    Align and distribute may be combined, e.g. align='middle' with distribute='horizontal'
    lays shapes out in an even row."""
    return await edit(
        path, "impress.align_shapes", slide_index=slide_index, shape_indexes=shape_indexes,
        align=align, distribute=distribute, relative_to=relative_to,
    )  # fmt: skip


@tool(I, "edit", title="Set shape name and alt text")
async def impress_set_shape_description(
    path: DocPath,
    slide_index: SlideIndex,
    shape_index: ShapeIndex,
    name: Annotated[str | None, Field(description="Shape name (shown in get_slide).")] = None,
    alt_text: Annotated[
        str | None, Field(description="Accessibility description (important for pictures).")
    ] = None,
) -> dict[str, Any]:
    """Set a shape's name and/or its alt text (accessibility description)."""
    return await edit(
        path, "impress.set_shape_description", slide_index=slide_index, shape_index=shape_index,
        name=name, alt_text=alt_text,
    )  # fmt: skip
