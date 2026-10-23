"""Impress: slide-level operations - delete, duplicate, move, layout, size, visibility,
background, transitions and footers. (Creating slides lives in content.py.)"""

from __future__ import annotations

from typing import Any

from com.sun.star.awt import Point, Size

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color

# -- delete / duplicate / move ---------------------------------------------------


@op("impress.delete_slides", kind="impress", mutates=True)
def delete_slides(s: Session, start_slide: int, end_slide: int | None = None) -> dict[str, Any]:
    """Delete slides start_slide..end_slide (inclusive)."""
    end = start_slide if end_slide is None else end_slide
    h.slide_at(s.doc, start_slide, "start_slide")
    h.slide_at(s.doc, end, "end_slide")
    if end < start_slide:
        raise OpError("end_slide must be >= start_slide.")
    pages = s.doc.getDrawPages()
    if end - start_slide + 1 >= pages.getCount():
        raise OpError(
            "A presentation must keep at least one slide, and this would delete all of them. "
            "Delete fewer slides, or empty the last one instead."
        )
    for i in range(end, start_slide - 1, -1):  # back to front keeps indexes valid
        pages.remove(pages.getByIndex(i))
    return {"deleted": end - start_slide + 1, "slide_count": pages.getCount()}


@op("impress.duplicate_slide", kind="impress", mutates=True)
def duplicate_slide(s: Session, slide_index: int, count: int = 1) -> dict[str, Any]:
    """Copy a slide (shapes, notes, background, transition) directly after the original."""
    slide = h.slide_at(s.doc, slide_index)
    if not 1 <= count <= 50:
        raise OpError("count must be between 1 and 50.")
    for _ in range(count):
        s.doc.duplicate(slide)  # always inserted right after ``slide``
    return {
        "new_slide_indexes": list(range(slide_index + 1, slide_index + 1 + count)),
        "slide_count": h.slide_count(s.doc),
    }


@op("impress.move_slide", kind="impress", mutates=True)
def move_slide(s: Session, from_index: int, to_index: int) -> dict[str, Any]:
    """Move a slide so that it ends up at to_index (the others shift)."""
    doc = s.doc
    source = h.slide_at(doc, from_index, "from_index")
    h.slide_at(doc, to_index, "to_index")
    if from_index == to_index:
        return {"slide_index": to_index, "_noop": True}
    # Copy to the target position, then delete the original (see copy_slide_content).
    # Moving down, the copy sits one place too far right until the original is gone.
    target, _ = h.insert_slide(s, to_index + 1 if to_index > from_index else to_index)
    h.copy_slide_content(s, source, target)
    pages = doc.getDrawPages()
    pages.remove(source)
    return {"slide_index": to_index, "slide_count": pages.getCount()}


# -- layout / size / visibility ----------------------------------------------------


@op("impress.set_slide_layout", kind="impress", mutates=True)
def set_slide_layout(s: Session, slide_index: int, layout: str) -> dict[str, Any]:
    """Change a slide's layout; existing text stays in the matching placeholders."""
    slide = h.slide_at(s.doc, slide_index)
    number = h.layout_id(layout)
    if slide.Layout == number:
        return {"slide_index": slide_index, "layout": layout, "_noop": True}
    slide.Layout = number
    return {
        "slide_index": slide_index,
        "layout": layout,
        "shapes": [{"index": i, "type": h.shape_kind(sh)} for i, sh in enumerate(h.all_shapes(slide))],
    }  # fmt: skip


SIZE_PRESETS = {
    "16:9": (280.0, 157.5),
    "16:10": (280.0, 175.0),
    "4:3": (280.0, 210.0),
    "a4_landscape": (297.0, 210.0),
    "letter_landscape": (279.4, 215.9),
}


def _geometry_of(pages: list[Any]) -> list[tuple[Any, tuple[int, int, int, int]]]:
    """(shape, (x, y, w, h)) for every shape on the given pages, read before a resize."""
    return [
        (shape, (shape.Position.X, shape.Position.Y, shape.Size.Width, shape.Size.Height))
        for page in pages
        for shape in h.all_shapes(page)
    ]


def _restore_geometry(
    boxes: list[tuple[Any, tuple[int, int, int, int]]], fx: float, fy: float
) -> None:
    for shape, (x, y, w, hh) in boxes:
        shape.Size = Size(max(1, round(w * fx)), max(1, round(hh * fy)))
        shape.Position = Point(round(x * fx), round(y * fy))


@op("impress.set_slide_size", kind="impress", mutates=True)
def set_slide_size(
    s: Session,
    preset: str | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    scale_content: bool = True,
) -> dict[str, Any]:
    """Resize all slides (and masters). Shapes are scaled proportionally unless told not to."""
    if preset is not None:
        if preset not in SIZE_PRESETS:
            raise OpError(
                f"Unknown preset {preset!r}. Use one of {', '.join(SIZE_PRESETS)}, or give "
                "width_mm and height_mm."
            )
        width_mm, height_mm = SIZE_PRESETS[preset]
    if width_mm is None or height_mm is None:
        raise OpError("Give a preset, or both width_mm and height_mm.")
    if not (20 <= width_mm <= 5000 and 20 <= height_mm <= 5000):
        raise OpError("width_mm and height_mm must be between 20 and 5000.")
    doc = s.doc
    first = h.slide_at(doc, 0)
    new_w, new_h = h.mm(width_mm), h.mm(height_mm)
    old_w, old_h = first.Width, first.Height
    if (old_w, old_h) == (new_w, new_h):
        return {"width_mm": width_mm, "height_mm": height_mm, "_noop": True}
    masters = doc.getMasterPages()
    pages = doc.getDrawPages()
    all_pages = [masters.getByIndex(i) for i in range(masters.getCount())]
    all_pages += [pages.getByIndex(i) for i in range(pages.getCount())]
    # Placeholders on slides follow their master's geometry, and LibreOffice may move other
    # shapes on its own, so the reliable way to a predictable result is: remember every
    # shape's box, resize, then write each box back - scaled, or unchanged.
    boxes = _geometry_of(all_pages)
    # The size is a document-wide setting: setting it on one page changes all of them.
    first.Width, first.Height = new_w, new_h
    for page in all_pages:  # defensive: make sure every page agrees
        if (page.Width, page.Height) != (new_w, new_h):
            page.Width, page.Height = new_w, new_h
    if scale_content:
        _restore_geometry(boxes, new_w / old_w, new_h / old_h)
    else:
        _restore_geometry(boxes, 1.0, 1.0)
    return {"width_mm": width_mm, "height_mm": height_mm, "scaled_content": scale_content}


def target_slides(doc: Any, slide_index: int | None) -> list[Any]:
    if slide_index is None:
        pages = doc.getDrawPages()
        return [pages.getByIndex(i) for i in range(pages.getCount())]
    return [h.slide_at(doc, slide_index)]


@op("impress.set_slide_hidden", kind="impress", mutates=True)
def set_slide_hidden(s: Session, slide_index: int, hidden: bool = True) -> dict[str, Any]:
    """Hide a slide from the slide show (it stays in the file) or unhide it."""
    slide = h.slide_at(s.doc, slide_index)
    if bool(slide.Visible) == (not hidden):
        return {"slide_index": slide_index, "hidden": hidden, "_noop": True}
    slide.Visible = not hidden
    return {"slide_index": slide_index, "hidden": hidden}


@op("impress.set_slide_background", kind="impress", mutates=True)
def set_slide_background(
    s: Session,
    color: str | None = None,
    slide_index: int | None = None,
    gradient_end_color: str | None = None,
    reset: bool = False,
) -> dict[str, Any]:
    """Set a solid (or two-colour gradient) background on one slide / all slides, or reset
    slides to their master's background."""
    if not reset and color is None:
        raise OpError("Give a colour, or reset=true to go back to the master's background.")
    slides = target_slides(s.doc, slide_index)
    for slide in slides:
        # A background object must not be shared between slides, so build one per slide.
        slide.Background = None if reset else make_background(s, color, gradient_end_color)
    return {"slides_changed": len(slides), "reset": reset}


def make_background(s: Session, color: str, gradient_end_color: str | None) -> Any:
    """A ``drawing.Background`` with a solid colour or a top-to-bottom two-colour gradient."""
    start = parse_color(color)
    background = s.new("com.sun.star.drawing.Background")
    if gradient_end_color is not None:
        h.set_props(
            background,
            {
                "FillStyle": enum("com.sun.star.drawing.FillStyle", "GRADIENT"),
                "FillGradient": h.make_gradient(start, parse_color(gradient_end_color), 0),
            },
        )
    else:
        h.set_props(
            background,
            {"FillStyle": enum("com.sun.star.drawing.FillStyle", "SOLID"), "FillColor": start},
        )
    return background


# -- transitions -------------------------------------------------------------------

_TRANSITION_TYPE = "com.sun.star.animations.TransitionType."
_TRANSITION_SUBTYPE = "com.sun.star.animations.TransitionSubType."
_DIRECTIONS = ("left", "right", "top", "bottom")
#: effect -> legacy FadeEffect name pattern. LibreOffice maps these onto the modern
#: transition types (PowerPoint's Wipe / Push / Cover) and writes them to .pptx.
_DIRECTIONAL = {"wipe": "FADE_FROM_{}", "push": "ROLL_FROM_{}", "cover": "MOVE_FROM_{}"}
TRANSITION_EFFECTS = ("none", "fade", "dissolve", "wipe", "push", "cover", "random")
_SPEEDS = {"slow": 2.0, "medium": 1.0, "fast": 0.5}


def transition_info(slide: Any) -> dict[str, Any] | None:
    """Describe a slide's transition for the read tools (None when it has none)."""
    p = h.get_props(slide, ("TransitionType", "Effect", "TransitionDuration", "Change", "Duration"))
    info: dict[str, Any] = {}
    ttype = p["TransitionType"] or 0
    effect = str(getattr(p["Effect"], "value", "NONE"))
    if ttype == 37:
        info["effect"] = "fade"
    elif ttype == 40 or effect == "DISSOLVE":
        info["effect"] = "dissolve"
    elif effect == "RANDOM" or ttype == 42:
        info["effect"] = "random"
    else:
        for name, pattern in _DIRECTIONAL.items():
            for direction in _DIRECTIONS:
                if effect == pattern.format(direction.upper()):
                    info["effect"] = name
                    info["direction"] = f"from_{direction}"
        if "effect" not in info and ttype:
            info["effect"] = f"other (type {ttype})"
    if info:
        duration = p["TransitionDuration"]
        info["duration_seconds"] = round(duration, 2) if duration else None
    if p["Change"] == 1:
        info["advance_after_seconds"] = p["Duration"]
    return info or None


@op("impress.set_slide_transition", kind="impress", mutates=True)
def set_slide_transition(
    s: Session,
    effect: str | None = None,
    slide_index: int | None = None,
    direction: str = "from_left",
    speed: str = "medium",
    advance_after_seconds: float | None = None,
) -> dict[str, Any]:
    """Set the slide-show transition of one slide or all slides."""
    if effect is None and advance_after_seconds is None:
        raise OpError("Give an effect and/or advance_after_seconds.")
    if effect is not None and effect not in TRANSITION_EFFECTS:
        raise OpError(f"Unknown effect {effect!r}. Use one of: {', '.join(TRANSITION_EFFECTS)}.")
    if speed not in _SPEEDS:
        raise OpError(f"speed must be one of {', '.join(_SPEEDS)}.")
    side = direction.removeprefix("from_")
    if side not in _DIRECTIONS:
        raise OpError("direction must be from_left, from_right, from_top or from_bottom.")
    slides = target_slides(s.doc, slide_index)
    fade_effect = "com.sun.star.presentation.FadeEffect"
    for slide in slides:
        if effect is not None:
            if effect == "none":
                slide.TransitionType = 0
                slide.TransitionSubtype = 0
                slide.Effect = enum(fade_effect, "NONE")
            elif effect == "fade":
                import uno

                slide.TransitionType = uno.getConstantByName(_TRANSITION_TYPE + "FADE")
                slide.TransitionSubtype = uno.getConstantByName(_TRANSITION_SUBTYPE + "CROSSFADE")
            else:
                name = (
                    _DIRECTIONAL[effect].format(side.upper())
                    if effect in _DIRECTIONAL
                    else effect.upper()
                )
                slide.Effect = enum(fade_effect, name)
            if effect != "none":
                slide.TransitionDuration = _SPEEDS[speed]
        if advance_after_seconds is not None:
            if advance_after_seconds <= 0:
                slide.Change = 0  # advance on click
            else:
                slide.Change = 1  # advance automatically after ``Duration`` seconds
                slide.Duration = max(1, round(advance_after_seconds))
    return {"slides_changed": len(slides), "effect": effect, "advance_after_seconds": advance_after_seconds}  # fmt: skip
