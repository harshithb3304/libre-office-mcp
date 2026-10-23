"""Typed content blocks for ``writer_add_content``.

Using Pydantic models (rather than a loose ``dict``) gives the AI model an exact JSON
schema for each kind of block, which makes malformed calls far less likely.
"""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from libreoffice_mcp.state import state
from libreoffice_mcp.tools.types import Alignment, Color


class HeadingBlock(BaseModel):
    type: Literal["heading"]
    text: str
    level: int = Field(1, ge=1, le=9, description="1 = chapter title ... 9 = smallest.")
    alignment: Alignment | None = None


class ParagraphBlock(BaseModel):
    type: Literal["paragraph"]
    text: str
    style: str | None = Field(None, description="Paragraph style, e.g. 'Title', 'Quotations'.")
    alignment: Alignment | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    font: str | None = None
    size: float | None = Field(None, description="Font size in points.")
    color: Color | None = None


class ListItem(BaseModel):
    text: str
    level: int = Field(0, ge=0, le=8, description="Nesting depth, 0 = top level.")


class ListBlock(BaseModel):
    type: Literal["bullets", "numbered"]
    items: list[str | ListItem] = Field(
        min_length=1, description="Plain strings, or {text, level} for nested items."
    )


class TableBlock(BaseModel):
    type: Literal["table"]
    data: list[list[str]] = Field(description="Rows of cell text; the first row is the header.")
    header: bool = Field(True, description="Style the first row as a header.")
    column_widths: list[float] | None = Field(
        None, description="Relative widths, one per column, e.g. [1, 3, 1]."
    )


class ImageBlock(BaseModel):
    type: Literal["image"]
    image_path: str = Field(description="Path to a PNG/JPG/GIF/SVG... file on disk.")
    width_mm: float | None = None
    height_mm: float | None = None
    alt_text: str | None = None


class PageBreakBlock(BaseModel):
    type: Literal["page_break"]


ContentBlock = Annotated[
    HeadingBlock | ParagraphBlock | ListBlock | TableBlock | ImageBlock | PageBreakBlock,
    Field(discriminator="type"),
]


def dump_blocks(blocks: list[Any]) -> list[dict[str, Any]]:
    """Validated models -> plain dicts for the worker (dropping unset fields).

    Image paths nested inside blocks are resolved and sandbox-checked here, because the
    generic ``*_path`` handling in ``tools.base`` only sees top-level arguments.
    """
    out = []
    for block in blocks:
        data = block.model_dump(exclude_none=True)
        if "items" in data:  # ListItem models -> plain strings or dicts
            data["items"] = [
                i if isinstance(i, str) else {"text": i["text"], "level": i["level"]}
                for i in data["items"]
            ]
        if "image_path" in data:
            data["image_path"] = str(state.paths.existing_file(data["image_path"]))
        out.append(data)
    return out
