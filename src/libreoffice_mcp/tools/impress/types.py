"""Parameter types shared by the Impress tools (descriptions are what the AI model reads)."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field

I = "impress"  # noqa: E741  (the app name every tool registers under)

SlideIndex = Annotated[
    int,
    Field(
        ge=0,
        description="0-based slide number (0 = first slide), as listed by impress_get_overview.",
    ),
]
OptionalSlideIndex = Annotated[
    int | None,
    Field(
        ge=0,
        description="0-based slide number. Omit to apply to ALL slides.",
    ),
]
ShapeIndex = Annotated[
    int,
    Field(
        ge=0,
        description="0-based shape number on that slide (stacking order, bottom first), as "
        "listed by impress_get_slide.",
    ),
]
Mm = Annotated[float, Field(ge=-2000, le=5000)]
PositiveMm = Annotated[float, Field(gt=0, le=5000)]
X = Annotated[float, Field(ge=-2000, le=5000, description="Left edge, mm from the slide's left.")]
Y = Annotated[float, Field(ge=-2000, le=5000, description="Top edge, mm from the slide's top.")]
Width = Annotated[float, Field(gt=0, le=5000, description="Width in millimetres.")]
Height = Annotated[float, Field(gt=0, le=5000, description="Height in millimetres.")]
OptX = Annotated[float | None, Field(ge=-2000, le=5000, description="Left edge in mm.")]
OptY = Annotated[float | None, Field(ge=-2000, le=5000, description="Top edge in mm.")]
OptWidth = Annotated[float | None, Field(gt=0, le=5000, description="Width in mm.")]
OptHeight = Annotated[float | None, Field(gt=0, le=5000, description="Height in mm.")]
FontSize = Annotated[float | None, Field(gt=0, le=1000, description="Font size in points.")]
Percent = Annotated[int | None, Field(ge=0, le=100)]

SlideLayout = Literal[
    "title_slide",
    "title_content",
    "two_content",
    "title_only",
    "blank",
    "centered_text",
    "content_over_content",
    "content_left_two_right",
    "two_left_content_right",
    "four_content",
    "six_content",
]

ShapeType = Literal[
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "triangle",
    "right_triangle",
    "diamond",
    "pentagon",
    "hexagon",
    "octagon",
    "parallelogram",
    "trapezoid",
    "star",
    "heart",
    "cross",
    "right_arrow",
    "left_arrow",
    "up_arrow",
    "down_arrow",
    "left_right_arrow",
    "chevron",
    "callout",
    "speech_bubble",
    "cylinder",
    "cube",
    "sun",
    "moon",
    "smiley",
    "lightning_bolt",
    "thought_bubble",
    "ring",
    "up_down_arrow",
    "flowchart_process",
    "flowchart_decision",
    "flowchart_terminator",
    "flowchart_document",
    "flowchart_data",
]

TransitionEffect = Literal["none", "fade", "dissolve", "wipe", "push", "cover", "random"]
ThemeName = Literal["clean", "corporate", "dark", "warm", "modern"]


class BulletItem(BaseModel):
    """One bullet with an explicit nesting level."""

    model_config = ConfigDict(extra="forbid")

    text: str
    level: int = Field(0, ge=0, le=8, description="Nesting depth: 0 = top level, 1 = sub-bullet...")


#: Bullets / paragraphs: plain strings, or {text, level} objects for nested points.
Bullets = Annotated[
    list[str | BulletItem],
    Field(
        description='Paragraphs, top to bottom. Each is a string, or {"text": ..., "level": 1} '
        "for a nested sub-bullet (level 0 = top level)."
    ),
]


def plain_items(items: list[Any] | None) -> list[Any] | None:
    """Validated bullets -> JSON-ready list (the worker accepts strings and {text, level})."""
    if items is None:
        return None
    out: list[Any] = []
    for item in items:
        if isinstance(item, str):
            out.append(item)
        elif isinstance(item, BulletItem):
            out.append({"text": item.text, "level": item.level})
        else:
            out.append({"text": item["text"], "level": item.get("level", 0)})
    return out
