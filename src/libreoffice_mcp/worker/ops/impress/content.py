"""Impress: creating slides and filling placeholders, shape text and speaker notes."""

from __future__ import annotations

from typing import Any

from com.sun.star.awt import Point, Size

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.ops.impress.master import ensure_theme
from libreoffice_mcp.worker.ops.impress.shapes import insert_image
from libreoffice_mcp.worker.ops.impress.slides import make_background
from libreoffice_mcp.worker.ops.impress.tables import create_table
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum

_SLIDE_KEYS = (
    "layout", "title", "subtitle", "bullets", "right_bullets", "image", "table", "notes",
    "background_color", "master",
)  # fmt: skip
_BODY_SHARE = 0.46  # width share of the bullets when a picture / table sits beside them


# -- speaker notes / placeholders ---------------------------------------------------------


@op("impress.set_notes", kind="impress", mutates=True)
def set_notes(s: Session, slide_index: int, text: str) -> dict[str, Any]:
    """Set (or with '' clear) the speaker notes of a slide."""
    slide = h.slide_at(s.doc, slide_index)
    shape = h.notes_shape(slide)
    if shape.getString() == text:
        return {"slide_index": slide_index, "_noop": True}
    shape.setString(text)
    return {"slide_index": slide_index, "characters": len(text)}


def shrink_to_fit(shape: Any) -> None:
    """Let long text shrink instead of overflowing its box (titles do not do this by default;
    body placeholders already do)."""
    # Shrinking only works on a box of fixed height; an auto-growing one just gets taller.
    shape.TextAutoGrowHeight = False
    shape.TextFitToSize = enum("com.sun.star.drawing.TextFitToSizeType", "AUTOFIT")


def _missing(slide_index: int, slide: Any, what: str) -> OpError:
    return OpError(
        f"Slide {slide_index} (layout '{h.layout_name(slide)}') has no {what}. Change the "
        "layout with impress_set_slide_layout (e.g. 'title_content' for a title + bullets), "
        "or add a text box with impress_add_text_box."
    )


@op("impress.set_title", kind="impress", mutates=True)
def set_title(
    s: Session, slide_index: int, text: str | None = None, subtitle: str | None = None
) -> dict[str, Any]:
    """Set the title (and, on a title slide, the subtitle) placeholder text."""
    if text is None and subtitle is None:
        raise OpError("Give text (the title) and/or subtitle.")
    slide = h.slide_at(s.doc, slide_index)
    ensure_theme(s)
    result: dict[str, Any] = {"slide_index": slide_index}
    for kind, value in (("title", text), ("subtitle", subtitle)):
        if value is None:
            continue
        found = h.find_placeholder(slide, kind)
        if found is None:
            raise _missing(slide_index, slide, f"{kind} placeholder")
        found[1].setString(value)
        if kind == "title":
            shrink_to_fit(found[1])
        result[f"{kind}_shape_index"] = found[0]
    return result


@op("impress.set_body", kind="impress", mutates=True)
def set_body(
    s: Session, slide_index: int, items: list[Any], content_index: int = 0
) -> dict[str, Any]:
    """Fill a content placeholder with bullets (strings or {text, level})."""
    slide = h.slide_at(s.doc, slide_index)
    ensure_theme(s)
    bodies = h.shapes_of_kind(slide, "body") or h.shapes_of_kind(slide, "subtitle")
    if not bodies:
        raise _missing(slide_index, slide, "content placeholder")
    if not 0 <= content_index < len(bodies):
        raise OpError(
            f"content_index {content_index} is out of range: slide {slide_index} has "
            f"{len(bodies)} content placeholder(s) (valid: 0..{len(bodies) - 1})."
        )
    shape_index, shape = bodies[content_index]
    h.set_paragraphs(shape, h.normalize_items(items, "items"))
    return {"slide_index": slide_index, "shape_index": shape_index}


@op("impress.set_shape_text", kind="impress", mutates=True)
def set_shape_text(
    s: Session,
    slide_index: int,
    shape_index: int,
    text: str | None = None,
    paragraphs: list[Any] | None = None,
) -> dict[str, Any]:
    """Replace all text of a shape (placeholder, text box, shape with text)."""
    if (text is None) == (paragraphs is None):
        raise OpError("Give either 'text' or 'paragraphs' (not both).")
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    if not h.shape_has_text_support(shape):
        extra = " Use impress_set_table_cells for tables." if h.shape_kind(shape) == "table" else ""
        raise OpError(f"A {h.shape_kind(shape)} shape has no text.{extra}")
    items = h.normalize_items(paragraphs if paragraphs is not None else text, "paragraphs")
    h.set_paragraphs(shape, items)
    return {"slide_index": slide_index, "shape_index": shape_index, "paragraphs": len(items)}


@op("impress.replace_text", kind="impress", mutates=True)
def replace_text(
    s: Session,
    search: str,
    replace: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    include_notes: bool = True,
) -> dict[str, Any]:
    """Replace text on every slide (and in speaker notes). Formatting of the text is kept."""
    if not search:
        raise OpError("search must not be empty.")
    total = 0
    for i in range(h.slide_count(s.doc)):
        slide = h.slide_at(s.doc, i)
        # The document itself has no replace API in Impress; each page (and notes page) does.
        targets = [slide, slide.getNotesPage()] if include_notes else [slide]
        for page in targets:
            descriptor = page.createReplaceDescriptor()
            descriptor.SearchString = search
            descriptor.ReplaceString = replace
            descriptor.SearchCaseSensitive = case_sensitive
            descriptor.SearchWords = whole_word
            total += page.replaceAll(descriptor)
    return {"replacements": total, **({"_noop": True} if total == 0 else {})}


# -- building slides ----------------------------------------------------------------------


def _is_pristine(slide: Any) -> bool:
    """True for the untouched first slide of a new presentation (empty title + subtitle)."""
    if slide.Layout != h.layout_id("title_slide"):
        return False
    shapes = h.all_shapes(slide)
    return all(
        h.shape_kind(sh) in ("title", "subtitle") and not h.shape_text(sh) for sh in shapes
    ) and not h.shape_text(h.notes_shape(slide))


def _infer_layout(spec: dict[str, Any]) -> str:
    if spec.get("subtitle") is not None:
        return "title_slide"
    if spec.get("right_bullets"):
        return "two_content"
    if spec.get("bullets"):
        return "title_content"
    if spec.get("title") is not None:
        return "title_only"
    return "blank"


def build_slide(
    s: Session, spec: dict[str, Any], at_index: int | None, reuse: Any | None = None
) -> dict[str, Any]:
    """Create one slide from a spec dict (see ``impress_add_slides``) and return a summary."""
    unknown = sorted(set(spec) - set(_SLIDE_KEYS))
    if unknown:
        raise OpError(
            f"Unknown slide field(s): {', '.join(unknown)}. Valid fields: {', '.join(_SLIDE_KEYS)}."
        )
    if spec.get("image") and spec.get("table"):
        raise OpError("A slide can have an image or a table in its content area, not both.")
    doc = s.doc
    layout = spec.get("layout") or _infer_layout(spec)
    number = h.layout_id(layout)  # validate before touching the document
    if reuse is not None:
        slide, index = reuse, h.slide_index_of(doc, reuse)
    else:
        slide, index = h.insert_slide(s, at_index)
    if spec.get("master"):
        slide.MasterPage = h.master_by_name(doc, spec["master"])
    slide.Layout = number
    placeholders = h.all_shapes(slide)
    kinds = [h.shape_kind(sh) for sh in placeholders]

    def pick(kind: str) -> list[Any]:
        return [sh for sh, k in zip(placeholders, kinds, strict=True) if k == kind]

    titles, subtitles, bodies = pick("title"), pick("subtitle"), pick("body")
    if layout == "centered_text":
        bodies = subtitles  # that layout's single text area is a subtitle-type shape

    def need(found: list[Any], position: int, field: str, what: str) -> Any:
        if len(found) <= position:
            raise OpError(
                f"The '{layout}' layout has no {what} for '{field}'. "
                f"Valid layouts and their areas: {_layout_help()}"
            )
        return found[position]

    if spec.get("title") is not None:
        title_shape = need(titles, 0, "title", "title placeholder")
        title_shape.setString(spec["title"])
        shrink_to_fit(title_shape)
    if spec.get("subtitle") is not None:
        need(subtitles, 0, "subtitle", "subtitle placeholder").setString(spec["subtitle"])
    if layout == "title_slide":
        _center_title_slide(slide, titles, subtitles)
    for field, position in (("bullets", 0), ("right_bullets", 1)):
        if spec.get(field):
            items = h.normalize_items(spec[field], field)
            h.set_paragraphs(need(bodies, position, field, "content area"), items)
    result: dict[str, Any] = {"slide_index": index, "layout": layout}

    extra = spec.get("image") or spec.get("table")
    if extra:
        empties = [sh for sh in bodies if not h.shape_text(sh)]
        if empties:  # the picture / table takes the place of an unused content area
            target = empties[0]
            area = (target.Position.X, target.Position.Y, target.Size.Width, target.Size.Height)
            slide.remove(target)
        elif len(bodies) == 1:  # bullets on the left, the picture / table beside them
            body = bodies[0]
            ax, ay, aw, ah = h.content_area(doc, slide)
            body.Position = Point(ax, ay)
            body.Size = Size(round(aw * _BODY_SHARE), ah)
            area = (ax + round(aw * 0.5), ay, round(aw * 0.5), ah)
        else:
            area = h.content_area(doc, slide)
        if spec.get("image"):
            image = spec["image"]
            if "image_path" not in image:
                raise OpError("image needs an 'image_path'.")
            insert_image(
                s, slide, image["image_path"], area, alt_text=image.get("alt_text"),
                width_mm=image.get("width_mm"), height_mm=image.get("height_mm"),
            )  # fmt: skip
            result["image_shape_index"] = slide.getCount() - 1
        else:
            table = spec["table"]
            if "data" not in table:
                raise OpError("table needs 'data' (a list of rows).")
            create_table(
                s, slide, table["data"], area, header=table.get("header", True),
                column_widths=table.get("column_widths"), font_size=table.get("font_size"),
            )  # fmt: skip
            result["table_shape_index"] = slide.getCount() - 1
    if spec.get("notes"):
        h.notes_shape(slide).setString(spec["notes"])
    if spec.get("background_color"):
        slide.Background = make_background(s, spec["background_color"], None)
    result["shape_count"] = slide.getCount()
    return result


def _center_title_slide(slide: Any, titles: list[Any], subtitles: list[Any]) -> None:
    """LibreOffice's title-slide layout keeps the title at the top edge like a normal slide;
    move title and subtitle to the middle of the slide so it reads as an opening slide."""
    width, height = slide.Width, slide.Height
    left, box_w = round(width * 0.1), round(width * 0.8)
    if titles:
        titles[0].Position = Point(left, round(height * 0.26))
        titles[0].Size = Size(box_w, round(height * 0.24))
    if subtitles:
        subtitles[0].Position = Point(left, round(height * 0.54))
        subtitles[0].Size = Size(box_w, round(height * 0.2))


def _layout_help() -> str:
    return "; ".join(f"{name} ({desc})" for name, (_, desc) in h.LAYOUTS.items())


def build_slides(s: Session, specs: list[dict[str, Any]], at_index: int | None) -> dict[str, Any]:
    if not specs:
        raise OpError("slides must contain at least one slide.")
    doc = s.doc
    ensure_theme(s)
    reuse = None
    if at_index is None and h.slide_count(doc) == 1 and _is_pristine(h.slide_at(doc, 0)):
        reuse = h.slide_at(doc, 0)  # a brand-new deck starts with one empty title slide
    built = []
    for n, spec in enumerate(specs):
        position = None if at_index is None else at_index + n
        try:
            built.append(build_slide(s, spec, position, reuse if n == 0 else None))
        except OpError as exc:
            raise OpError(f"slides[{n}]: {exc.message}", exc.code) from exc
    result: dict[str, Any] = {
        "slides_added": len(built),
        "first_slide_index": built[0]["slide_index"],
        "last_slide_index": built[-1]["slide_index"],
        "slide_count": h.slide_count(doc),
        "slides": built,
    }
    if reuse is not None:
        result["reused_empty_first_slide"] = True
    return result


@op("impress.add_slide", kind="impress", mutates=True)
def add_slide(
    s: Session,
    layout: str | None = None,
    title: str | None = None,
    subtitle: str | None = None,
    bullets: list[Any] | None = None,
    right_bullets: list[Any] | None = None,
    notes: str | None = None,
    at_index: int | None = None,
    master: str | None = None,
    background_color: str | None = None,
) -> dict[str, Any]:
    fields = {
        "layout": layout,
        "title": title,
        "subtitle": subtitle,
        "bullets": bullets,
        "right_bullets": right_bullets,
        "notes": notes,
        "master": master,
        "background_color": background_color,
    }
    result = build_slides(s, [{k: v for k, v in fields.items() if v is not None}], at_index)
    single = result["slides"][0]
    single["slide_count"] = result["slide_count"]
    if result.get("reused_empty_first_slide"):
        single["reused_empty_first_slide"] = True
    return single


@op("impress.add_slides", kind="impress", mutates=True)
def add_slides(
    s: Session, slides: list[dict[str, Any]], at_index: int | None = None
) -> dict[str, Any]:
    return build_slides(s, slides, at_index)
