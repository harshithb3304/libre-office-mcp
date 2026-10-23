"""Impress: read-only operations (``impress.get_*`` / ``impress.find_text``)."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.ops.impress.slides import transition_info
from libreoffice_mcp.worker.ops.impress.tables import table_grid
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session

_SHAPE_PROPS = (
    "Name", "RotateAngle", "IsPresentationObject", "FillStyle", "FillColor",
    "LineStyle", "LineColor", "Description", "OnClick", "Bookmark",
)  # fmt: skip
_TEXT_KINDS = ("title", "subtitle", "body", "text_box", "shape")


def title_of(slide: Any) -> str:
    """The slide's title text; falls back to a subtitle (title slides) or empty string."""
    for kind in ("title", "subtitle"):
        found = h.find_placeholder(slide, kind)
        if found and (text := h.shape_text(found[1]).strip()):
            return text.replace("\n", " ")
    return ""


def _clip(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[:limit] + "..."


def shape_summary(index: int, shape: Any, text_chars: int = 300) -> dict[str, Any]:
    """One compact dict describing a shape (what ``impress_get_slide`` lists)."""
    kind = h.shape_kind(shape)
    p = h.get_props(shape, _SHAPE_PROPS)
    info: dict[str, Any] = {"index": index, "type": kind}
    if p["Name"]:
        info["name"] = p["Name"]
    pos, size = shape.Position, shape.Size
    info.update(
        x_mm=h.to_mm(pos.X), y_mm=h.to_mm(pos.Y),
        width_mm=h.to_mm(size.Width), height_mm=h.to_mm(size.Height),
    )  # fmt: skip
    if p["RotateAngle"]:
        info["rotation_deg"] = round((-p["RotateAngle"] % 36000) / 100, 1)
    if p["IsPresentationObject"]:
        info["placeholder"] = True
    if kind == "table":
        grid = table_grid(shape, max_rows=6)
        info["rows"] = shape.Model.Rows.Count
        info["columns"] = shape.Model.Columns.Count
        info["preview"] = grid
        return info
    if kind in _TEXT_KINDS:
        text = h.shape_text(shape)
        if text:
            info["text"] = _clip(text, text_chars)
            if kind == "body":
                levels = [h.para_level(para) for para in h.paragraphs_of(shape)]
                if any(levels):
                    info["levels"] = levels
        elif p["IsPresentationObject"]:
            info["empty"] = True
    fill_style = str(getattr(p["FillStyle"], "value", p["FillStyle"]))
    line_style = str(getattr(p["LineStyle"], "value", p["LineStyle"]))
    if kind in ("shape", "text_box", "image"):
        if fill_style != "NONE" and kind != "image":
            info["fill"] = h.hex_color(p["FillColor"])
        if line_style != "NONE":
            info["line"] = h.hex_color(p["LineColor"])
    elif kind == "line":
        info["line"] = h.hex_color(p["LineColor"])
    if shape.ShapeType == "com.sun.star.drawing.CustomShape":
        geometry = {x.Name: x.Value for x in shape.CustomShapeGeometry}
        info["geometry"] = geometry.get("Type")
    if kind == "image" and p["Description"]:
        info["alt_text"] = p["Description"]
    if str(getattr(p["OnClick"], "value", p["OnClick"])) == "DOCUMENT" and p["Bookmark"]:
        info["link"] = p["Bookmark"]
    return info


def _notes_text(slide: Any) -> str:
    try:
        return h.shape_text(h.notes_shape(slide))
    except OpError:
        return ""


def slide_details(doc: Any, slide: Any, index: int, text_chars: int = 300) -> dict[str, Any]:
    props = h.get_props(slide, ("Visible", "Background"))
    background = props["Background"]
    shapes = h.all_shapes(slide)
    info: dict[str, Any] = {
        "slide_index": index,
        "layout": h.layout_name(slide),
        "master": slide.MasterPage.Name,
        "hidden": not props["Visible"],
        "shape_count": len(shapes),
        "shapes": [shape_summary(i, sh, text_chars) for i, sh in enumerate(shapes)],
    }
    content_index = 0
    for entry in info["shapes"]:  # which "content area" impress_set_body would address
        if entry["type"] == "body":
            entry["content_index"] = content_index
            content_index += 1
    if background is not None and str(background.FillStyle.value) == "SOLID":
        info["background"] = h.hex_color(background.FillColor)
    transition = transition_info(slide)
    if transition:
        info["transition"] = transition
    notes = _notes_text(slide)
    if notes:
        info["notes"] = notes
    return info


@op("impress.get_overview", kind="impress")
def get_overview(s: Session) -> dict[str, Any]:
    """Slide count, slide size, masters and one line per slide (title, layout, shape count)."""
    doc = s.doc
    pages = doc.getDrawPages()
    count = pages.getCount()
    slides = []
    for i in range(count):
        slide = pages.getByIndex(i)
        shapes = h.all_shapes(slide)
        title = title_of(slide)
        entry: dict[str, Any] = {
            "slide_index": i,
            "title": title,
            "layout": h.layout_name(slide),
            "shape_count": len(shapes),
        }
        if not title:  # e.g. a blank slide with only a text box: show what is on it
            for sh in shapes:
                text = h.shape_text(sh).strip() if h.shape_has_text_support(sh) else ""
                if text:
                    entry["first_text"] = _clip(text.replace("\n", " "), 60)
                    break
        if not slide.Visible:
            entry["hidden"] = True
        if _notes_text(slide):
            entry["has_notes"] = True
        slides.append(entry)
    width, height = pages.getByIndex(0).Width, pages.getByIndex(0).Height
    masters = doc.getMasterPages()
    return {
        "slide_count": count,
        "slide_width_mm": h.to_mm(width),
        "slide_height_mm": h.to_mm(height),
        "aspect_ratio": round(width / height, 3),
        "masters": [masters.getByIndex(i).Name for i in range(masters.getCount())],
        "slides": slides,
    }


@op("impress.get_slide", kind="impress")
def get_slide(s: Session, slide_index: int, text_chars: int = 300) -> dict[str, Any]:
    """Every shape on a slide, with the indexes the editing tools need."""
    return slide_details(s.doc, h.slide_at(s.doc, slide_index), slide_index, text_chars)


@op("impress.get_shape", kind="impress")
def get_shape(s: Session, slide_index: int, shape_index: int) -> dict[str, Any]:
    """Full detail of one shape: complete text with per-paragraph formatting, fill, line, ..."""
    slide = h.slide_at(s.doc, slide_index)
    shape = h.shape_at(slide, shape_index, slide_index)
    info = shape_summary(shape_index, shape, text_chars=100000)
    info["slide_index"] = slide_index
    kind = info["type"]
    if kind == "table":
        info["data"] = table_grid(shape)
        info.pop("preview", None)
        return info
    extra = h.get_props(
        shape, ("FillTransparence", "LineWidth", "Shadow", "Title", "LineTransparence")
    )
    if kind not in ("group", "object", "chart"):
        info["fill_transparency"] = extra["FillTransparence"]
        info["line_width_mm"] = h.to_mm(extra["LineWidth"] or 0)
        info["shadow"] = bool(extra["Shadow"])
    if h.shape_has_text_support(shape) and h.shape_text(shape):
        info["paragraphs"] = [_paragraph_info(para, kind) for para in h.paragraphs_of(shape)]
    return info


_RUN_PROPS = (
    "CharFontName",
    "CharHeight",
    "CharWeight",
    "CharPosture",
    "CharColor",
    "CharUnderline",
)


def _run_format(portion: Any) -> dict[str, Any]:
    p = h.get_props(portion, _RUN_PROPS)
    return {
        "font": p["CharFontName"],
        "size_pt": p["CharHeight"],
        "bold": (p["CharWeight"] or 100) > 100,
        "italic": str(getattr(p["CharPosture"], "value", "NONE")) != "NONE",
        "underline": bool(p["CharUnderline"]),
        "color": h.hex_color(p["CharColor"]),
    }


def _paragraph_info(para: Any, kind: str) -> dict[str, Any]:
    """Level, alignment and character formatting of one paragraph.

    Character formatting is read from the text *runs* (portions): the paragraph object
    itself only reports paragraph-level defaults once a file has been reloaded. When the
    runs differ (partly bold, a link...) they are listed under ``runs``.
    """
    p = h.get_props(para, ("ParaAdjust", "NumberingIsNumber", "NumberingLevel"))
    runs = []
    enumeration = para.createEnumeration()
    while enumeration.hasMoreElements():
        portion = enumeration.nextElement()
        run: dict[str, Any] = {"text": portion.getString(), **_run_format(portion)}
        if portion.TextPortionType == "TextField":
            run["link"] = portion.TextField.URL
        runs.append(run)
    first = runs[0] if runs else _run_format(para)
    formats = [{k: v for k, v in r.items() if k not in ("text", "link")} for r in runs]
    entry: dict[str, Any] = {
        "text": para.getString(),
        "level": h.para_level(para),
        **{k: first[k] for k in ("font", "size_pt", "bold", "italic", "underline", "color")},
        "alignment": h.ALIGN_FROM_UNO.get(p["ParaAdjust"], "left"),
    }
    if len(runs) > 1 and (any(f != formats[0] for f in formats) or any("link" in r for r in runs)):
        entry["runs"] = runs
    if kind in ("body", "text_box"):
        # In a placeholder a paragraph without a level (None) shows no bullet.
        has_level = kind != "body" or p["NumberingLevel"] is not None
        entry["bullet"] = bool(p["NumberingIsNumber"]) and has_level
    return entry


def _outline_lines(slide: Any, index: int, include_notes: bool) -> list[str]:
    lines = [f"Slide {index}: {title_of(slide) or '(no title)'}"]
    title_done = False
    for si, sh in enumerate(h.all_shapes(slide)):
        kind = h.shape_kind(sh)
        if kind == "table":
            lines.append(f"  [table, shape {si}]")
            lines.extend("    " + " | ".join(row) for row in table_grid(sh))
            continue
        if kind not in _TEXT_KINDS:
            continue
        text = h.shape_text(sh)
        if not text.strip():
            continue
        if kind == "title" and not title_done:
            title_done = True
            continue  # already in the heading line
        if kind == "subtitle" and not h.find_placeholder(slide, "title"):
            continue
        if kind == "body":
            for para in h.paragraphs_of(sh):
                lines.append("  " + "  " * h.para_level(para) + "- " + para.getString())
        elif kind == "subtitle":
            lines.append(f"  (subtitle) {text}")
        else:
            lines.append(f"  [{kind}, shape {si}] {text}")
    if include_notes and (notes := _notes_text(slide).strip()):
        lines.append(f"  Notes: {notes}")
    return lines


@op("impress.get_text", kind="impress")
def get_text(
    s: Session, start_slide: int = 0, max_chars: int = 20000, include_notes: bool = True
) -> dict[str, Any]:
    """Text outline of the whole deck: per slide the title, bullets (indented by level), other
    text, tables and speaker notes. Paginated by slide."""
    count = h.slide_count(s.doc)
    if start_slide and start_slide >= count:
        raise OpError(f"start_slide {start_slide} is past the end ({count} slides).")
    out: list[str] = []
    size = 0
    index = start_slide
    while index < count:
        lines = _outline_lines(h.slide_at(s.doc, index), index, include_notes)
        chunk = "\n".join(lines)
        if out and size + len(chunk) > max_chars:
            break
        out.append(chunk)
        size += len(chunk) + 1
        index += 1
    result: dict[str, Any] = {
        "slide_count": count,
        "slides_returned": index - start_slide,
        "outline": "\n".join(out),
    }
    if index < count:
        result["next_start_slide"] = index
    return result


@op("impress.get_notes", kind="impress")
def get_notes(s: Session, slide_index: int | None = None) -> dict[str, Any]:
    """Speaker notes of one slide, or of every slide that has notes."""
    if slide_index is not None:
        text = _notes_text(h.slide_at(s.doc, slide_index))
        return {"slide_index": slide_index, "notes": text}
    found = []
    for i in range(h.slide_count(s.doc)):
        text = _notes_text(h.slide_at(s.doc, i))
        if text.strip():
            found.append({"slide_index": i, "notes": text})
    return {"count": len(found), "slides": found}


@op("impress.find_text", kind="impress")
def find_text(
    s: Session, query: str, case_sensitive: bool = False, include_notes: bool = True
) -> dict[str, Any]:
    """Find text on slides (shapes, table cells and optionally notes)."""
    if not query:
        raise OpError("query must not be empty.")
    needle = query if case_sensitive else query.lower()

    def hit(text: str) -> bool:
        return needle in (text if case_sensitive else text.lower())

    matches: list[dict[str, Any]] = []
    for i in range(h.slide_count(s.doc)):
        slide = h.slide_at(s.doc, i)
        for si, sh in enumerate(h.all_shapes(slide)):
            kind = h.shape_kind(sh)
            if kind == "table":
                for r, row in enumerate(table_grid(sh)):
                    for c, cell in enumerate(row):
                        if hit(cell):
                            matches.append(_match(i, si, kind, cell, row=r, column=c))
            elif h.shape_has_text_support(sh) and hit(text := h.shape_text(sh)):
                matches.append(_match(i, si, kind, text))
        if include_notes and hit(notes := _notes_text(slide)):
            matches.append(_match(i, None, "notes", notes))
    return {"query": query, "total_matches": len(matches), "matches": matches[:100]}


def _match(slide: int, shape: int | None, kind: str, text: str, **extra: Any) -> dict[str, Any]:
    out: dict[str, Any] = {"slide_index": slide, "type": kind, "text": _clip(text, 120)}
    if shape is not None:
        out["shape_index"] = shape
    out.update(extra)
    return out
