"""Shared parameter types.

``Annotated[..., Field(description=...)]`` is how a parameter's description reaches the
AI model, so reuse these aliases to keep wording consistent across tools.
"""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import Field

DocPath = Annotated[
    str,
    Field(
        description="Path to the document. Absolute, or relative to the default documents "
        "folder. Supports ~ for the home folder."
    ),
]

BlockIndex = Annotated[
    int,
    Field(
        ge=0,
        description="0-based index of a body block (paragraph or table), as listed by "
        "writer_get_structure.",
    ),
]

Alignment = Literal["left", "center", "right", "justify"]

Color = Annotated[
    str,
    Field(description="Colour as '#RRGGBB' (e.g. '#1F4E79') or a basic name like 'red'."),
]

Position = Annotated[
    Literal["end", "start", "before", "after"],
    Field(
        description="Where to insert: 'end' of the document (default), 'start', or 'before' / "
        "'after' the block given in block_index."
    ),
]

AnchorBlock = Annotated[
    int | None,
    Field(
        ge=0, description="Reference block index; required when position is 'before' or 'after'."
    ),
]

TableIndex = Annotated[
    int,
    Field(ge=0, description="0-based table number, as listed by writer_get_tables."),
]

CellRange = Annotated[
    str,
    Field(description="Cell or range in spreadsheet notation, e.g. 'B2' or 'A1:C1'."),
]
