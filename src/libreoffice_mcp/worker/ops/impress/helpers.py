"""Shared building blocks for the Impress operations.

Vocabulary used throughout:

* **slide**  - a draw page, addressed by 0-based index in presentation order.
* **shape**  - anything on a slide (placeholder, text box, picture, table...), addressed by its
  0-based index in the slide's z-order (bottom first), exactly as ``impress_get_slide`` lists
  it. Indexes shift when shapes are deleted or re-ordered; mutating ops return fresh ones.
* **placeholder** - a shape the layout created (title, subtitle, body/content). They inherit
  fonts and bullets from the master slide, which is what makes a deck look consistent.

UNO lengths are 1/100 mm; the tool level speaks millimetres.
"""

from __future__ import annotations

from typing import Any

from com.sun.star.awt import Point, Size

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.registry import OpError
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, parse_color

MM100 = 100


def mm(value: float) -> int:
    return round(value * MM100)


def to_mm(value: int | float) -> float:
    return round(value / MM100, 1)


def hex_color(value: Any) -> str | None:
    """UNO colour int -> '#RRGGBB' (None for the 'automatic' colour -1 / unset)."""
    if value is None or not isinstance(value, int) or value < 0:
        return None
    return f"#{value & 0xFFFFFF:06X}"


# -- layouts -------------------------------------------------------------------

#: name -> (LibreOffice AutoLayout number, description). Verified against LibreOffice 24.2;
#: numbers not listed here (charts, clip-art, vertical text...) are deliberately not exposed.
LAYOUTS: dict[str, tuple[int, str]] = {
    "title_slide": (0, "Title and subtitle (for the first slide or a section opener)"),
    "title_content": (1, "Title and one body/content area for bullets (the usual slide)"),
    "two_content": (3, "Title and two side-by-side content areas"),
    "title_only": (19, "Title only; add your own text boxes, pictures, tables or shapes"),
    "blank": (20, "Nothing; add everything yourself"),
    "centered_text": (32, "One large centred text area, no title"),
    "content_over_content": (14, "Title and two content areas stacked top / bottom"),
    "content_left_two_right": (12, "Title, one tall area on the left, two stacked on the right"),
    "two_left_content_right": (15, "Title, two stacked areas on the left, one tall on the right"),
    "four_content": (18, "Title and a 2 x 2 grid of content areas"),
    "six_content": (34, "Title and a 3 x 2 grid of content areas"),
}
LAYOUT_BY_ID = {number: name for name, (number, _) in LAYOUTS.items()}


def layout_id(name: str) -> int:
    if name not in LAYOUTS:
        raise OpError(f"Unknown layout {name!r}. Valid layouts: {', '.join(LAYOUTS)}.")
    return LAYOUTS[name][0]


def layout_name(slide: Any) -> str:
    number = slide.Layout
    return LAYOUT_BY_ID.get(number, f"other({number})")


# -- slides --------------------------------------------------------------------


def slide_count(doc: Any) -> int:
    return doc.getDrawPages().getCount()


def slide_at(doc: Any, index: int, what: str = "slide_index") -> Any:
    pages = doc.getDrawPages()
    count = pages.getCount()
    if not 0 <= index < count:
        raise OpError(
            f"{what} {index} is out of range: the presentation has {count} slide(s) "
            f"(valid: 0..{count - 1}). Use impress_get_overview to see them.",
            protocol.E_NOT_FOUND,
        )
    return pages.getByIndex(index)


def slide_index_of(doc: Any, slide: Any) -> int:
    """Presentation-order index of ``slide`` (the ``Number`` property is 1-based)."""
    return int(slide.Number) - 1


def insert_slide(s: Session, at_index: int | None) -> tuple[Any, int]:
    """Insert an empty slide so that it ends up at ``at_index`` (None = at the end)."""
    pages = s.doc.getDrawPages()
    count = pages.getCount()
    index = count if at_index is None else at_index
    if not 0 <= index <= count:
        raise OpError(
            f"at_index {index} is out of range: the presentation has {count} slide(s), so a new "
            f"slide can go at 0..{count} ({count} = at the end)."
        )
    if index > 0:
        # insertNewByIndex(n) inserts the new slide *after* slide n (n is clamped to the last).
        return pages.insertNewByIndex(index - 1), index
    # Nothing can be inserted before the first slide. Insert after it, move the first slide's
    # content into the new slide and recycle the (emptied) first slide as the new one.
    first = pages.getByIndex(0)
    copy = pages.insertNewByIndex(0)
    copy_slide_content(s, first, copy)
    reset_slide(first)
    return first, 0


# Slide properties carried over when a slide is rebuilt (move / insert at the front).
_SLIDE_PROPS = (
    "Visible", "IsBackgroundVisible", "IsBackgroundObjectsVisible", "IsFooterVisible",
    "FooterText", "IsPageNumberVisible", "IsDateTimeVisible", "IsDateTimeFixed", "DateTimeText",
    "TransitionType", "TransitionSubtype", "TransitionDuration", "TransitionDirection",
    "TransitionFadeColor", "Speed", "Change", "Duration",
)  # fmt: skip
_BACKGROUND_PROPS = ("FillStyle", "FillColor", "FillGradient", "FillTransparence")


def copy_slide_content(s: Session, source: Any, target: Any) -> None:
    """Make ``target`` a copy of ``source``: shapes, layout, master, background, transition,
    footer settings and notes. Whatever ``target`` held before is removed. Animations are not
    carried over (they are not reachable through the API).

    LibreOffice has no "move slide" call outside its GUI slide sorter, so moving a slide is
    done by copying it to the new position and deleting the original.
    """
    for shape in reversed(all_shapes(target)):
        target.remove(shape)
    target.MasterPage = source.MasterPage
    shapes = all_shapes(source)
    layout = source.Layout
    target.Layout = layout_id("blank")  # no placeholders yet: the pasted shapes replace them
    if shapes:
        copy_shapes(s, shapes, source, target)
        target.Layout = layout  # re-attaches pasted title/body shapes as placeholders
        _drop_surplus_placeholders(target, shapes)
    else:
        target.Layout = layout
    set_props(target, {k: v for k, v in get_props(source, _SLIDE_PROPS).items() if v is not None})
    background = source.Background
    target.Background = None
    if background is not None:
        copy = s.new("com.sun.star.drawing.Background")
        wanted = get_props(background, _BACKGROUND_PROPS)
        set_props(copy, {k: v for k, v in wanted.items() if v is not None})
        target.Background = copy
    notes_shape(target).setString(notes_shape(source).getString())


def reset_slide(slide: Any) -> None:
    """Return a slide to the state of a freshly inserted one (no shapes, notes, transition)."""
    for shape in reversed(all_shapes(slide)):
        slide.remove(shape)
    slide.Layout = layout_id("blank")
    set_props(slide, {"Visible": True, "TransitionType": 0, "TransitionSubtype": 0, "Change": 0})
    slide.Background = None
    notes_shape(slide).setString("")


def _drop_surplus_placeholders(slide: Any, original_shapes: list[Any]) -> None:
    """Re-applying a layout re-creates placeholders the author had deleted; remove the extras
    (they are empty) so a copied slide matches the original."""
    wanted: dict[str, int] = {}
    for shape in original_shapes:
        kind = shape_kind(shape)
        if kind in ("title", "subtitle", "body"):
            wanted[kind] = wanted.get(kind, 0) + 1
    for kind in ("title", "subtitle", "body"):
        present = shapes_of_kind(slide, kind)
        surplus = len(present) - wanted.get(kind, 0)
        for _, shape in reversed(present):
            if surplus <= 0:
                break
            if not shape_text(shape):
                slide.remove(shape)
                surplus -= 1


def master_by_name(doc: Any, name: str | None) -> Any:
    masters = doc.getMasterPages()
    if name is None:
        return masters.getByIndex(0)
    for i in range(masters.getCount()):
        master = masters.getByIndex(i)
        if master.Name == name:
            return master
    names = [masters.getByIndex(i).Name for i in range(masters.getCount())]
    raise OpError(
        f"No master slide named {name!r}. Available: {', '.join(names)}. Use impress_list_masters.",
        protocol.E_NOT_FOUND,
    )


# -- shapes --------------------------------------------------------------------

_PLACEHOLDER_ROLES = {
    "com.sun.star.presentation.TitleTextShape": "title",
    "com.sun.star.presentation.SubtitleShape": "subtitle",
    "com.sun.star.presentation.OutlinerShape": "body",
    "com.sun.star.presentation.NotesShape": "notes",
    "com.sun.star.presentation.DateTimeShape": "date",
    "com.sun.star.presentation.FooterShape": "footer",
    "com.sun.star.presentation.SlideNumberShape": "slide_number",
}
_SHAPE_KINDS = {
    "com.sun.star.drawing.TextShape": "text_box",
    "com.sun.star.drawing.RectangleShape": "shape",
    "com.sun.star.drawing.EllipseShape": "shape",
    "com.sun.star.drawing.CustomShape": "shape",
    "com.sun.star.drawing.PolyPolygonShape": "shape",
    "com.sun.star.drawing.ClosedBezierShape": "shape",
    "com.sun.star.drawing.LineShape": "line",
    "com.sun.star.drawing.PolyLineShape": "line",
    "com.sun.star.drawing.OpenBezierShape": "line",
    "com.sun.star.drawing.ConnectorShape": "connector",
    "com.sun.star.drawing.GraphicObjectShape": "image",
    "com.sun.star.presentation.GraphicObjectShape": "image",
    "com.sun.star.drawing.TableShape": "table",
    "com.sun.star.presentation.TableShape": "table",
    "com.sun.star.drawing.GroupShape": "group",
    "com.sun.star.drawing.OLE2Shape": "object",
    "com.sun.star.presentation.OLE2Shape": "object",
    "com.sun.star.presentation.ChartShape": "chart",
}


def shape_kind(shape: Any) -> str:
    """Normalised kind: title, subtitle, body, text_box, shape, line, image, table, group..."""
    shape_type = shape.ShapeType
    role = _PLACEHOLDER_ROLES.get(shape_type)
    if role:
        return role
    return _SHAPE_KINDS.get(shape_type, shape_type.rsplit(".", 1)[-1])


def all_shapes(slide: Any) -> list[Any]:
    return [slide.getByIndex(i) for i in range(slide.getCount())]


def shape_at(slide: Any, index: int, slide_index: int | None = None) -> Any:
    count = slide.getCount()
    if not 0 <= index < count:
        where = f"slide {slide_index}" if slide_index is not None else "the slide"
        raise OpError(
            f"shape_index {index} is out of range: {where} has {count} shape(s) "
            f"(valid: 0..{count - 1}). Use impress_get_slide to list them.",
            protocol.E_NOT_FOUND,
        )
    return slide.getByIndex(index)


def shapes_of_kind(slide: Any, kind: str) -> list[tuple[int, Any]]:
    return [(i, sh) for i, sh in enumerate(all_shapes(slide)) if shape_kind(sh) == kind]


def get_props(obj: Any, names: tuple[str, ...]) -> dict[str, Any]:
    """Read several properties in one round trip (falls back to one-by-one when a name is
    unknown to this shape type, which makes the bulk call fail as a whole)."""
    try:
        return dict(zip(names, obj.getPropertyValues(names), strict=True))
    except Exception:
        out: dict[str, Any] = {}
        for name in names:
            try:
                out[name] = obj.getPropertyValue(name)
            except Exception:
                out[name] = None
        return out


def set_props(obj: Any, values: dict[str, Any]) -> None:
    """Set several properties in one round trip (names must be sorted for the bulk call).

    Some objects (slide backgrounds, styles) are plain property sets without the bulk
    interface; they get one call per property.
    """
    if not values:
        return
    names = tuple(sorted(values))
    try:
        bulk = obj.setPropertyValues
    except AttributeError:
        for name in names:
            obj.setPropertyValue(name, values[name])
        return
    bulk(names, tuple(values[n] for n in names))


def geometry_mm(shape: Any) -> dict[str, float]:
    pos, size = shape.Position, shape.Size
    return {
        "x_mm": to_mm(pos.X),
        "y_mm": to_mm(pos.Y),
        "width_mm": to_mm(size.Width),
        "height_mm": to_mm(size.Height),
    }


def set_geometry(
    shape: Any,
    x_mm: float | None = None,
    y_mm: float | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
) -> None:
    """Move/resize a shape. Size is applied before position so the top-left corner lands
    where asked even for shapes that grow around their centre."""
    if width_mm is not None or height_mm is not None:
        cur = shape.Size
        shape.Size = Size(
            mm(width_mm) if width_mm is not None else cur.Width,
            mm(height_mm) if height_mm is not None else cur.Height,
        )
    if x_mm is not None or y_mm is not None:
        cur_pos = shape.Position
        shape.Position = Point(
            mm(x_mm) if x_mm is not None else cur_pos.X,
            mm(y_mm) if y_mm is not None else cur_pos.Y,
        )


def content_area(doc: Any, slide: Any) -> tuple[int, int, int, int]:
    """(x, y, width, height) in 1/100 mm of the region below the title, where body content
    belongs. Taken from the master's body placeholder so it follows the deck's design;
    falls back to sensible margins for masters without one."""
    master = slide.MasterPage
    for i in range(master.getCount()):
        sh = master.getByIndex(i)
        if sh.ShapeType == "com.sun.star.presentation.OutlinerShape":
            pos, size = sh.Position, sh.Size
            return pos.X, pos.Y, size.Width, size.Height
    width, height = slide.Width, slide.Height
    return (
        round(width * 0.05),
        round(height * 0.24),
        round(width * 0.9),
        round(height * 0.62),
    )


def fit_into(
    native_w: float, native_h: float, area: tuple[int, int, int, int]
) -> tuple[int, int, int, int]:
    """Scale a (w, h) box to fit inside ``area`` keeping its aspect ratio; centre it."""
    ax, ay, aw, ah = area
    scale = min(aw / native_w, ah / native_h)
    w, h = round(native_w * scale), round(native_h * scale)
    return ax + (aw - w) // 2, ay + (ah - h) // 2, w, h


# -- text ----------------------------------------------------------------------

ALIGN_TO_UNO = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "BLOCK"}
ALIGN_FROM_UNO = {0: "left", 1: "right", 2: "justify", 3: "center"}


def paragraphs_of(shape: Any) -> list[Any]:
    enumeration = shape.createEnumeration()
    paragraphs = []
    while enumeration.hasMoreElements():
        paragraphs.append(enumeration.nextElement())
    return paragraphs


def para_level(paragraph: Any) -> int:
    """Outline level of a paragraph (0 = top level); LibreOffice returns None for level 0."""
    try:
        return int(paragraph.NumberingLevel or 0)
    except Exception:
        return 0


def normalize_items(items: Any, what: str = "items") -> list[tuple[str, int]]:
    """Accept a string, or a list of strings / {text, level} dicts; return (text, level) pairs.
    Newlines inside an item start further paragraphs at the same level."""
    if isinstance(items, str):
        items = [items]
    if not isinstance(items, (list, tuple)):
        raise OpError(f"{what} must be a string or a list of strings / {{text, level}} objects.")
    out: list[tuple[str, int]] = []
    for item in items:
        if isinstance(item, str):
            text, level = item, 0
        elif isinstance(item, dict) and "text" in item:
            text, level = str(item["text"]), int(item.get("level") or 0)
        else:
            raise OpError(
                f"Each entry of {what} must be a string or an object like "
                f'{{"text": "...", "level": 1}}; got {item!r}.'
            )
        if not 0 <= level <= 8:
            raise OpError(f"Bullet level {level} is out of range (0 = top level .. 8).")
        out.extend((line, level) for line in text.split("\n"))
    return out


def set_paragraphs(shape: Any, items: list[tuple[str, int]]) -> None:
    """Replace a shape's text with paragraphs at the given outline levels.

    In a body placeholder every paragraph needs an explicit level: LibreOffice leaves new
    paragraphs without one (``NumberingLevel`` is void) and such paragraphs get no bullet
    at all. Level 0 is the first bullet level.
    """
    shape.setString("\n".join(text for text, _ in items))
    explicit = shape_kind(shape) == "body"
    if not explicit and not any(level for _, level in items):
        return
    for para, (_, level) in zip(paragraphs_of(shape), items, strict=False):
        if explicit or level:
            para.NumberingLevel = level


def shape_text(shape: Any) -> str:
    try:
        return shape.getString()
    except Exception:
        return ""


def shape_has_text_support(shape: Any) -> bool:
    return shape_kind(shape) not in ("line", "image", "group", "table", "object", "chart")


def notes_shape(slide: Any) -> Any:
    notes = slide.getNotesPage()
    for i in range(notes.getCount()):
        sh = notes.getByIndex(i)
        if sh.ShapeType == "com.sun.star.presentation.NotesShape":
            return sh
    raise OpError("This slide has no notes area.", protocol.E_UNSUPPORTED)


def find_placeholder(slide: Any, kind: str, nth: int = 0) -> tuple[int, Any] | None:
    matches = shapes_of_kind(slide, kind)
    return matches[nth] if nth < len(matches) else None


# -- drawing objects -----------------------------------------------------------


def copy_shapes(s: Session, shapes: list[Any], source: Any, target: Any) -> list[Any]:
    """Copy shapes to another slide (or the same one) through the clipboard.

    UNO has no "clone shape" call, and ``XDrawPageDuplicator`` only copies whole slides. The
    ``.uno:Copy`` / ``.uno:Paste`` dispatches work on the hidden document's view and keep
    everything (text formatting, hyperlinks, tables, pictures). The pasted shapes land on top
    of the target's z-order, in their original order.
    """
    controller = s.doc.getCurrentController()
    frame = controller.getFrame()
    dispatcher = s.service("com.sun.star.frame.DispatchHelper")
    selection = s.service("com.sun.star.drawing.ShapeCollection")
    for shape in shapes:
        selection.add(shape)
    controller.setCurrentPage(source)
    controller.select(selection)
    dispatcher.executeDispatch(frame, ".uno:Copy", "", 0, ())
    controller.setCurrentPage(target)
    before = target.getCount()
    dispatcher.executeDispatch(frame, ".uno:Paste", "", 0, ())
    created = [target.getByIndex(i) for i in range(before, target.getCount())]
    if len(created) != len(shapes):
        raise OpError(
            "LibreOffice could not copy the shape(s) (clipboard unavailable).",
            protocol.E_OFFICE,
        )
    return created


def create_shape(s: Session, service: str, slide: Any) -> Any:
    shape = s.new(service)
    slide.add(shape)  # a shape must be on a slide before most of its properties work
    return shape


def apply_look(
    shape: Any,
    fill_color: str | None = None,
    no_fill: bool = False,
    line_color: str | None = None,
    line_width_mm: float | None = None,
    no_line: bool = False,
    gradient_end_color: str | None = None,
    gradient_angle: int | None = None,
    transparency: int | None = None,
    line_transparency: int | None = None,
    shadow: bool | None = None,
    dash: str | None = None,
) -> list[str]:
    """Fill / outline / shadow styling shared by shape tools. Returns what was changed."""
    changed: list[str] = []
    values: dict[str, Any] = {}
    if no_fill:
        values["FillStyle"] = enum("com.sun.star.drawing.FillStyle", "NONE")
        changed.append("fill")
    elif fill_color is not None:
        color = parse_color(fill_color)
        values["FillColor"] = color
        if gradient_end_color is not None:
            gradient = make_gradient(color, parse_color(gradient_end_color), gradient_angle or 0)
            values["FillStyle"] = enum("com.sun.star.drawing.FillStyle", "GRADIENT")
            values["FillGradient"] = gradient
        else:
            values["FillStyle"] = enum("com.sun.star.drawing.FillStyle", "SOLID")
        changed.append("fill")
    if transparency is not None:
        if not 0 <= transparency <= 100:
            raise OpError("transparency is a percentage from 0 (opaque) to 100 (invisible).")
        values["FillTransparence"] = transparency
        changed.append("transparency")
    if no_line:
        values["LineStyle"] = enum("com.sun.star.drawing.LineStyle", "NONE")
        changed.append("line")
    elif line_color is not None or line_width_mm is not None or dash is not None:
        style = "DASH" if dash in ("dashed", "dotted") else "SOLID"
        values["LineStyle"] = enum("com.sun.star.drawing.LineStyle", style)
        if line_color is not None:
            values["LineColor"] = parse_color(line_color)
        if line_width_mm is not None:
            values["LineWidth"] = mm(line_width_mm)
        if dash in ("dashed", "dotted"):
            values["LineDash"] = _dash(dash)
        changed.append("line")
    if line_transparency is not None:
        if not 0 <= line_transparency <= 100:
            raise OpError("line_transparency is a percentage from 0 to 100.")
        values["LineTransparence"] = line_transparency
    if shadow is not None:
        values["Shadow"] = bool(shadow)
        if shadow:
            values.update(ShadowColor=0x808080, ShadowXDistance=200, ShadowYDistance=200)
            values["ShadowTransparence"] = 50
        changed.append("shadow")
    set_props(shape, values)
    return changed


def make_gradient(start: int, end: int, angle: int) -> Any:
    import uno

    gradient = uno.createUnoStruct("com.sun.star.awt.Gradient")
    gradient.Style = enum("com.sun.star.awt.GradientStyle", "LINEAR")
    gradient.StartColor, gradient.EndColor = start, end
    gradient.Angle = int(angle) * 10  # tenths of a degree
    gradient.Border = 0
    gradient.XOffset = gradient.YOffset = 0
    gradient.StartIntensity = gradient.EndIntensity = 100
    gradient.StepCount = 0
    return gradient


def _dash(kind: str) -> Any:
    import uno

    dash = uno.createUnoStruct("com.sun.star.drawing.LineDash")
    dash.Style = enum("com.sun.star.drawing.DashStyle", "RECT")
    if kind == "dotted":
        dash.Dots, dash.DotLen, dash.Dashes, dash.DashLen, dash.Distance = 1, 100, 0, 0, 150
    else:
        dash.Dots, dash.DotLen, dash.Dashes, dash.DashLen, dash.Distance = 0, 0, 1, 400, 200
    return dash


def apply_alt_text(shape: Any, alt_text: str | None, name: str | None = None) -> None:
    if alt_text:
        set_props(shape, {"Description": alt_text, "Title": alt_text})
    if name:
        shape.Name = name
