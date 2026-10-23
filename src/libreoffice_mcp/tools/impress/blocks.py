"""Typed slide blocks for ``impress_add_slides``.

Pydantic models (rather than loose dicts) give the AI an exact JSON schema per slide, which
makes malformed calls far less likely. ``extra="forbid"`` turns typos such as ``bullet`` into
a clear validation error instead of silently dropping the content.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import BaseModel, ConfigDict, Field

from libreoffice_mcp.state import state
from libreoffice_mcp.tools.impress.types import BulletItem, SlideLayout, plain_items
from libreoffice_mcp.tools.types import Color


class ImageSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    image_path: str = Field(description="Path to a PNG/JPG/GIF/SVG file on disk.")
    alt_text: str | None = Field(None, description="Accessibility description of the picture.")
    width_mm: Annotated[float | None, Field(gt=0, le=5000)] = None
    height_mm: Annotated[float | None, Field(gt=0, le=5000)] = None


class TableSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    data: list[list[str | int | float]] = Field(
        min_length=1, description="Rows of cell values; the first row is the header."
    )
    header: bool = Field(True, description="Style the first row as a header.")
    column_widths: list[float] | None = Field(
        None, description="Relative widths, one per column, e.g. [2, 1, 1]."
    )
    font_size: Annotated[float | None, Field(gt=0, le=200)] = Field(
        None, description="Cell font size in points (default: chosen from the row count)."
    )


class SlideBlock(BaseModel):
    """One slide. Everything is optional; unused placeholders stay empty."""

    model_config = ConfigDict(extra="forbid")

    layout: SlideLayout | None = Field(
        None,
        description="Slide layout. Omitted: 'title_slide' if a subtitle is given, 'two_content' "
        "if right_bullets is given, 'title_content' if bullets are given, else 'title_only'.",
    )
    title: str | None = Field(None, description="Title text (needs a layout with a title).")
    subtitle: str | None = Field(None, description="Subtitle text (layout 'title_slide').")
    bullets: list[str | BulletItem] | None = Field(
        None,
        description="Body text: strings, or {text, level} for nested bullets. Goes into the "
        "first content area.",
    )
    right_bullets: list[str | BulletItem] | None = Field(
        None, description="Second content area (layout 'two_content' and similar)."
    )
    image: ImageSpec | None = Field(
        None,
        description="A picture placed in the content area (beside the bullets if there are any).",
    )
    table: TableSpec | None = Field(
        None, description="A table placed in the content area (beside the bullets if any)."
    )
    notes: str | None = Field(None, description="Speaker notes for this slide.")
    background_color: Color | None = Field(None, description="Solid background colour.")
    master: str | None = Field(None, description="Master slide name (see impress_list_masters).")


def dump_slides(slides: list[SlideBlock]) -> list[dict[str, Any]]:
    """Validated models -> plain dicts for the worker (dropping unset fields).

    Image paths nested inside slide blocks are resolved and sandbox-checked here, because the
    generic ``*_path`` handling in ``tools.base`` only sees top-level arguments.
    """
    out = []
    for slide in slides:
        data = slide.model_dump(exclude_none=True)
        for key in ("bullets", "right_bullets"):
            if key in data:
                data[key] = plain_items(data[key])
        if "image" in data:
            data["image"]["image_path"] = str(
                state.paths.existing_file(data["image"]["image_path"])
            )
        out.append(data)
    return out
