"""Impress: master slides - look of the whole deck (fonts, colours, background).

.pptx note: LibreOffice 24.2 writes a master's text styles into .pptx in a form it cannot read
back, so after the next load of the file the styles are LibreOffice's defaults again. To keep
decks consistent we (a) remember the chosen look in a document property that does survive,
(b) re-apply it before ops that create placeholder text (``ensure_theme``), and (c) write it
directly into the placeholders that already exist (``bake_theme``).
"""

from __future__ import annotations

import json
from typing import Any

from libreoffice_mcp.worker.ops.impress import helpers as h
from libreoffice_mcp.worker.ops.impress.slides import make_background, target_slides
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import parse_color

#: Relative size of outline levels 1..9 compared with level 1.
_LEVEL_SCALE = (1.0, 0.85, 0.75, 0.7, 0.65, 0.6, 0.6, 0.6, 0.6)
_THEME_PROPERTY = "LibreOfficeMcpTheme"

#: Built-in looks for ``impress_apply_theme``. Fonts use the Office names; the Docker image
#: renders them with metric-compatible fonts (Calibri -> Carlito, Cambria -> Caladea).
THEMES: dict[str, dict[str, Any]] = {
    "clean": dict(
        background="#FFFFFF", title_color="#1F2937", body_color="#374151",
        title_font="Calibri", body_font="Calibri", title_size=40, body_size=26, title_bold=True,
    ),
    "corporate": dict(
        background="#FFFFFF", title_color="#1F4E79", body_color="#333333",
        title_font="Calibri", body_font="Calibri", title_size=40, body_size=26, title_bold=True,
    ),
    "dark": dict(
        background="#1B1F2A", title_color="#FFFFFF", body_color="#D6DBE4",
        title_font="Calibri", body_font="Calibri", title_size=40, body_size=26, title_bold=True,
    ),
    "warm": dict(
        background="#FFF7EA", title_color="#7A3E00", body_color="#4A3B2A",
        title_font="Cambria", body_font="Calibri", title_size=40, body_size=26, title_bold=True,
    ),
    "modern": dict(
        background="#F4F7FA", title_color="#0F766E", body_color="#1F2937",
        title_font="Arial", body_font="Arial", title_size=36, body_size=24, title_bold=True,
    ),
}  # fmt: skip


def _style(doc: Any, master: Any, name: str) -> Any:
    return doc.getStyleFamilies().getByName(master.Name).getByName(name)


def _all_masters(doc: Any) -> list[Any]:
    masters = doc.getMasterPages()
    return [masters.getByIndex(i) for i in range(masters.getCount())]


def _text_values(theme: dict[str, Any], kind: str, level: int = 0) -> dict[str, Any]:
    """Character properties a theme prescribes for a placeholder kind (title/subtitle/body)."""
    values: dict[str, Any] = {}
    if kind == "title":
        prefix, size, bold = "title", theme.get("title_size"), theme.get("title_bold")
        if bold is not None:
            values["CharWeight"] = 150.0 if bold else 100.0
    else:
        prefix, size = "body", theme.get("body_size")
        if size is not None:
            scale = 1.1 if kind == "subtitle" else _LEVEL_SCALE[level]
            size = max(10, round(size * scale))
    if theme.get(f"{prefix}_font") is not None:
        values["CharFontName"] = theme[f"{prefix}_font"]
    if theme.get(f"{prefix}_color") is not None:
        values["CharColor"] = parse_color(theme[f"{prefix}_color"])
    if size is not None:
        values["CharHeight"] = float(size)
    return values


def restyle_master(s: Session, master: Any, theme: dict[str, Any]) -> list[str]:
    """Apply a theme (see ``_SETTING_NAMES``) to one master: background, title / body text
    styles and the footer, date and slide-number placeholders. Returns what changed."""
    changed: list[str] = []
    if theme.get("background") is not None:
        master.Background = make_background(s, theme["background"], None)
        changed.append("background")
    if title := _text_values(theme, "title"):
        h.set_props(_style(s.doc, master, "title"), title)
        changed.append("title style")
    body_given = any(theme.get(k) is not None for k in ("body_font", "body_size", "body_color"))
    if body_given:
        for level in range(1, 10):
            h.set_props(
                _style(s.doc, master, f"outline{level}"), _text_values(theme, "body", level - 1)
            )
        h.set_props(_style(s.doc, master, "subtitle"), _text_values(theme, "subtitle"))
        # The footer / date / number areas are plain shapes with their own (serif, black)
        # defaults; without this they vanish on a dark theme and clash with the fonts.
        footer = {
            k: v
            for k, v in _text_values(theme, "body").items()
            if k in ("CharFontName", "CharColor")
        }
        for shape in h.all_shapes(master):
            if h.shape_kind(shape) in ("date", "footer", "slide_number") and footer:
                h.set_props(shape, footer)
        changed.append("body styles")
    return changed


# -- .pptx persistence ----------------------------------------------------------------------


def _is_pptx(s: Session) -> bool:
    return s.path.suffix.lower() == ".pptx"


def _stored_theme(s: Session) -> dict[str, Any]:
    props = s.doc.getDocumentProperties().getUserDefinedProperties()
    try:
        return json.loads(props.getPropertyValue(_THEME_PROPERTY))
    except Exception:
        return {}


def _store_theme(s: Session, theme: dict[str, Any]) -> None:
    props = s.doc.getDocumentProperties().getUserDefinedProperties()
    text = json.dumps(theme)
    if props.getPropertySetInfo().hasPropertyByName(_THEME_PROPERTY):
        props.setPropertyValue(_THEME_PROPERTY, text)
    else:
        props.addProperty(_THEME_PROPERTY, 0, text)


def bake_theme(s: Session, theme: dict[str, Any]) -> None:
    """Write the theme's fonts / colours / sizes directly into existing title, subtitle and
    body text, so slides keep their look even though .pptx forgets the master styles."""
    for i in range(h.slide_count(s.doc)):
        for shape in h.all_shapes(h.slide_at(s.doc, i)):
            kind = h.shape_kind(shape)
            if kind not in ("title", "subtitle", "body") or not h.shape_text(shape):
                continue
            for para in h.paragraphs_of(shape):
                level = min(h.para_level(para), 8) if kind == "body" else 0
                h.set_props(para, _text_values(theme, kind, level))


def ensure_theme(s: Session) -> None:
    """Re-apply the remembered theme to masters that lost it (only matters for .pptx)."""
    if not _is_pptx(s):
        return
    theme = _stored_theme(s)
    title = _text_values(theme, "title")
    if not title:
        return
    for master in _all_masters(s.doc):
        style = _style(s.doc, master, "title")
        if all(style.getPropertyValue(k) == v for k, v in title.items() if k != "CharWeight"):
            continue
        restyle_master(s, master, theme)


def _remember(s: Session, update: dict[str, Any]) -> None:
    """Merge ``update`` into the remembered theme and, for .pptx, bake it into the slides."""
    if not _is_pptx(s):
        return
    theme = _stored_theme(s)
    theme.update(update)
    _store_theme(s, theme)
    bake_theme(s, theme)


# -- operations -------------------------------------------------------------------------------


def _color_of(background: Any) -> str | None:
    try:
        if background is not None and str(background.FillStyle.value) == "SOLID":
            return h.hex_color(background.FillColor)
    except Exception:
        pass
    return None


@op("impress.list_masters", kind="impress")
def list_masters(s: Session) -> dict[str, Any]:
    """Master slides with their background and title/body text styles, plus the layouts."""
    doc = s.doc
    pages = doc.getDrawPages()
    usage: dict[str, int] = {}
    for i in range(pages.getCount()):
        name = pages.getByIndex(i).MasterPage.Name
        usage[name] = usage.get(name, 0) + 1
    listed = []
    for master in _all_masters(doc):
        entry: dict[str, Any] = {
            "name": master.Name,
            "used_by_slides": usage.get(master.Name, 0),
            "background": _color_of(master.Background),
        }
        for label, style_name in (
            ("title", "title"),
            ("body_level_1", "outline1"),
            ("body_level_2", "outline2"),
        ):
            st = h.get_props(
                _style(doc, master, style_name), ("CharFontName", "CharHeight", "CharColor")
            )
            entry[label] = {
                "font": st["CharFontName"],
                "size_pt": st["CharHeight"],
                "color": h.hex_color(st["CharColor"]),
            }
        listed.append(entry)
    return {
        "masters": listed,
        "layouts": [{"name": n, "description": d} for n, (_, d) in h.LAYOUTS.items()],
        "themes": sorted(THEMES),
    }


def _target_masters(doc: Any, master_name: str | None) -> list[Any]:
    """One named master, or all of them: a deck imported from .pptx has one master per layout,
    so re-theming "the deck" means restyling every master."""
    if master_name is not None:
        return [h.master_by_name(doc, master_name)]
    return _all_masters(doc)


@op("impress.set_master_style", kind="impress", mutates=True)
def set_master_style(
    s: Session,
    master_name: str | None = None,
    background: str | None = None,
    title_font: str | None = None,
    title_size: float | None = None,
    title_color: str | None = None,
    title_bold: bool | None = None,
    body_font: str | None = None,
    body_size: float | None = None,
    body_color: str | None = None,
) -> dict[str, Any]:
    settings = {
        "background": background, "title_font": title_font, "title_size": title_size,
        "title_color": title_color, "title_bold": title_bold, "body_font": body_font,
        "body_size": body_size, "body_color": body_color,
    }  # fmt: skip
    given = {k: v for k, v in settings.items() if v is not None}
    if not given:
        raise OpError("Nothing to change: give at least one style option.")
    masters = _target_masters(s.doc, master_name)
    changed: list[str] = []
    for master in masters:
        changed = restyle_master(s, master, given)
    _remember(s, given)
    return {"masters": [m.Name for m in masters], "changed": changed}


@op("impress.apply_theme", kind="impress", mutates=True)
def apply_theme(s: Session, theme: str, master_name: str | None = None) -> dict[str, Any]:
    """Apply a built-in look (colours + fonts) to a master slide."""
    if theme not in THEMES:
        raise OpError(f"Unknown theme {theme!r}. Available: {', '.join(sorted(THEMES))}.")
    masters = _target_masters(s.doc, master_name)
    look = THEMES[theme]
    for master in masters:
        restyle_master(s, master, look)
    _remember(s, look)
    return {"masters": [m.Name for m in masters], "theme": theme, "settings": look}


@op("impress.set_slide_master", kind="impress", mutates=True)
def set_slide_master(
    s: Session, master_name: str, slide_index: int | None = None
) -> dict[str, Any]:
    master = h.master_by_name(s.doc, master_name)
    if slide_index is None:
        pages = s.doc.getDrawPages()
        slides = [pages.getByIndex(i) for i in range(pages.getCount())]
    else:
        slides = [h.slide_at(s.doc, slide_index)]
    for slide in slides:
        slide.MasterPage = master
    return {"master": master_name, "slides_changed": len(slides)}


# -- footer ----------------------------------------------------------------------


@op("impress.set_footer", kind="impress", mutates=True)
def set_footer(
    s: Session,
    slide_index: int | None = None,
    footer_text: str | None = None,
    show_slide_number: bool | None = None,
    show_date: bool | None = None,
    date_text: str | None = None,
) -> dict[str, Any]:
    """Footer text, slide number and date for one slide or all slides."""
    ensure_theme(s)
    values: dict[str, Any] = {}
    if footer_text is not None:
        values.update(FooterText=footer_text, IsFooterVisible=bool(footer_text))
    if show_slide_number is not None:
        values["IsPageNumberVisible"] = bool(show_slide_number)
    if date_text is not None:
        values.update(DateTimeText=date_text, IsDateTimeFixed=True, IsDateTimeVisible=True)
    if show_date is not None:
        values["IsDateTimeVisible"] = bool(show_date)
    if not values:
        raise OpError("Give footer_text, show_slide_number, show_date and/or date_text.")
    slides = target_slides(s.doc, slide_index)
    for slide in slides:
        h.set_props(slide, values)
    return {"slides_changed": len(slides)}
