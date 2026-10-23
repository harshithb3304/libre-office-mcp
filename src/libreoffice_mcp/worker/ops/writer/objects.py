"""Writer: images, hyperlinks, bookmarks, footnotes, comments and merging in other files."""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any

from libreoffice_mcp import protocol
from libreoffice_mcp.worker.ops.writer import helpers as h
from libreoffice_mcp.worker.ops.writer.read import first_page_style
from libreoffice_mcp.worker.registry import OpError, op
from libreoffice_mcp.worker.session import Session
from libreoffice_mcp.worker.util import enum, file_url, pv

_PX_TO_MM100 = 26.46  # one pixel at 96 DPI, in 1/100 mm


# -- images --------------------------------------------------------------------


def _text_width_mm100(doc: Any) -> int:
    style = doc.getStyleFamilies().getByName("PageStyles").getByName(first_page_style(doc))
    return int(style.Width - style.LeftMargin - style.RightMargin)


def insert_image_at(
    s: Session,
    slot: h.Slot,
    image_path: str,
    width_mm: float | None = None,
    height_mm: float | None = None,
    alt_text: str | None = None,
    alignment: str | None = None,
) -> Any:
    """Embed an image in a new paragraph at ``slot``. Keeps the aspect ratio when only one
    dimension is given; unsized images are scaled to fit the text width."""
    provider = s.service("com.sun.star.graphic.GraphicProvider")
    graphic = provider.queryGraphic((pv("URL", file_url(image_path)),))
    if graphic is None:
        raise OpError(
            f"{Path(image_path).name} is not an image LibreOffice can read.", protocol.E_UNSUPPORTED
        )
    size = graphic.Size100thMM
    if size.Width > 0 and size.Height > 0:
        native_w, native_h = size.Width, size.Height
    else:  # no DPI info in the file: assume 96 DPI
        native_w = graphic.SizePixel.Width * _PX_TO_MM100
        native_h = graphic.SizePixel.Height * _PX_TO_MM100
    ratio = native_h / native_w
    if width_mm and height_mm:
        width, height = h.mm(width_mm), h.mm(height_mm)
    elif width_mm:
        width = h.mm(width_mm)
        height = int(width * ratio)
    elif height_mm:
        height = h.mm(height_mm)
        width = int(height / ratio)
    else:
        width = min(int(native_w), _text_width_mm100(s.doc))
        height = int(width * ratio)

    image = s.new("com.sun.star.text.TextGraphicObject")
    image.Graphic = graphic  # embeds the pixels in the document (not a link to the file)
    image.AnchorType = enum("com.sun.star.text.TextContentAnchorType", "AS_CHARACTER")
    image.Width, image.Height = width, height
    if alt_text:
        image.Description = alt_text
        image.Title = alt_text
    para = slot.paragraph("")
    if alignment:
        h.apply_para_format(para, alignment=alignment)
    s.doc.getText().insertTextContent(para.getStart(), image, False)
    return image


@op("writer.insert_image", kind="writer", mutates=True)
def insert_image(
    s: Session,
    image_path: str,
    width_mm: float | None = None,
    height_mm: float | None = None,
    alt_text: str | None = None,
    alignment: str | None = None,
    position: str = "end",
    block_index: int | None = None,
) -> dict[str, Any]:
    slot = h.Slot(s.doc, position, block_index)
    image = insert_image_at(s, slot, image_path, width_mm, height_mm, alt_text, alignment)
    return {
        "block_index": slot.after,
        "width_mm": round(image.Width / 100, 1),
        "height_mm": round(image.Height / 100, 1),
    }


def _get_image(doc: Any, image_index: int) -> Any:
    images = doc.getGraphicObjects()
    if not 0 <= image_index < images.getCount():
        raise OpError(
            f"image_index {image_index} is out of range ({images.getCount()} images). "
            "Use writer_get_images.",
            protocol.E_NOT_FOUND,
        )
    return images.getByIndex(image_index)


@op("writer.set_image", kind="writer", mutates=True)
def set_image(
    s: Session,
    image_index: int,
    width_mm: float | None = None,
    height_mm: float | None = None,
    alt_text: str | None = None,
) -> dict[str, Any]:
    """Resize an image (aspect ratio kept if only one size is given) or set its alt text."""
    image = _get_image(s.doc, image_index)
    if width_mm and height_mm:
        image.Width, image.Height = h.mm(width_mm), h.mm(height_mm)
    elif width_mm:
        ratio = image.Height / image.Width
        image.Width = h.mm(width_mm)
        image.Height = int(image.Width * ratio)
    elif height_mm:
        ratio = image.Width / image.Height
        image.Height = h.mm(height_mm)
        image.Width = int(image.Height * ratio)
    if alt_text is not None:
        image.Description = alt_text
        image.Title = alt_text
    return {"width_mm": round(image.Width / 100, 1), "height_mm": round(image.Height / 100, 1)}


@op("writer.delete_image", kind="writer", mutates=True)
def delete_image(s: Session, image_index: int) -> dict[str, Any]:
    image = _get_image(s.doc, image_index)
    image.getAnchor().getText().removeTextContent(image)
    return {"deleted_image": image_index}


# -- links, bookmarks, notes, comments ------------------------------------------------


def _anchor_range(doc: Any, find: str | None, block_index: int | None) -> Any:
    """The first range matching ``find`` (inside ``block_index`` if given), or a whole paragraph."""
    target = h.target_ranges(doc, find=find, block_index=block_index, occurrence="first")[0]
    # A paragraph object is not a valid range for insertTextContent; a cursor over it is.
    return target.getText().createTextCursorByRange(target)


@op("writer.add_hyperlink", kind="writer", mutates=True)
def add_hyperlink(
    s: Session,
    url: str,
    find: str | None = None,
    text: str | None = None,
    block_index: int | None = None,
    occurrence: str = "first",
) -> dict[str, Any]:
    """Link existing text (``find``), or append new linked ``text`` to a paragraph / the end."""
    doc = s.doc
    char_style = (
        "Internet link"
        if doc.getStyleFamilies().getByName("CharacterStyles").hasByName("Internet link")
        else None
    )
    if find:
        ranges = h.target_ranges(doc, find=find, block_index=block_index, occurrence=occurrence)
        for r in ranges:
            r.HyperLinkURL = url
            if char_style:
                r.CharStyleName = char_style
        return {"links_added": len(ranges)}
    if not text:
        raise OpError("Give 'find' (text to turn into a link) or 'text' (new linked text).")
    if block_index is None:
        para = h.Slot(doc, "end", None).paragraph("")
    else:
        para = h.paragraph_at(h.all_blocks(doc), block_index)
    cursor = doc.getText().createTextCursorByRange(para.getEnd())
    doc.getText().insertString(cursor, text, False)
    cursor.goLeft(len(text), True)
    cursor.HyperLinkURL = url
    if char_style:
        cursor.CharStyleName = char_style
    return {"links_added": 1}


@op("writer.add_bookmark", kind="writer", mutates=True)
def add_bookmark(
    s: Session, name: str, find: str | None = None, block_index: int | None = None
) -> dict[str, Any]:
    """Create a named bookmark covering the found text or a paragraph."""
    if s.doc.getBookmarks().hasByName(name):
        raise OpError(f"A bookmark named '{name}' already exists.", protocol.E_EXISTS)
    target = _anchor_range(s.doc, find, block_index)
    bookmark = s.new("com.sun.star.text.Bookmark")
    bookmark.setName(name)
    target.getText().insertTextContent(target, bookmark, True)
    return {"bookmark": name}


@op("writer.delete_bookmark", kind="writer", mutates=True)
def delete_bookmark(s: Session, name: str) -> dict[str, Any]:
    bookmarks = s.doc.getBookmarks()
    if not bookmarks.hasByName(name):
        raise OpError(
            f"No bookmark named '{name}'. Use writer_get_bookmarks.", protocol.E_NOT_FOUND
        )
    bookmarks.getByName(name).dispose()
    return {"deleted_bookmark": name}


@op("writer.add_footnote", kind="writer", mutates=True)
def add_footnote(
    s: Session,
    note_text: str,
    find: str | None = None,
    block_index: int | None = None,
    endnote: bool = False,
) -> dict[str, Any]:
    """Attach a footnote (or endnote) at the end of the found text / paragraph."""
    target = _anchor_range(s.doc, find, block_index)
    note = s.new("com.sun.star.text.Endnote" if endnote else "com.sun.star.text.Footnote")
    where = target.getText().createTextCursorByRange(target.getEnd())
    target.getText().insertTextContent(where, note, False)
    note.setString(note_text)  # must be set after insertion
    return {"footnote_index": s.doc.getFootnotes().getCount() - 1, "endnote": endnote}


def _now() -> Any:
    from com.sun.star.util import DateTime

    now = datetime.datetime.now()
    return DateTime(0, now.second, now.minute, now.hour, now.day, now.month, now.year, False)


@op("writer.add_comment", kind="writer", mutates=True)
def add_comment(
    s: Session,
    comment: str,
    find: str | None = None,
    block_index: int | None = None,
    author: str = "AI Assistant",
) -> dict[str, Any]:
    """Add a review comment anchored to the found text / paragraph."""
    target = _anchor_range(s.doc, find, block_index)
    annotation = s.new("com.sun.star.text.textfield.Annotation")
    annotation.Author = author
    annotation.Content = comment
    annotation.DateTimeValue = _now()
    text = target.getText()
    before = _count_comments(s.doc)
    # A range anchor highlights the commented text, but ranges that already carry attributes
    # (a hyperlink, say) refuse it - and a refused insert can leave a ghost annotation behind.
    if not getattr(target, "HyperLinkURL", ""):
        try:
            text.insertTextContent(target, annotation, True)
        except Exception:
            _drop_ghosts(s.doc, before)
    if _count_comments(s.doc) == before:
        text.insertTextContent(text.createTextCursorByRange(target.getEnd()), annotation, False)
    return {"comment_index": _count_comments(s.doc) - 1}


def _annotations(doc: Any) -> list[Any]:
    out = []
    fields = doc.getTextFields().createEnumeration()
    while fields.hasMoreElements():
        field = fields.nextElement()
        if field.supportsService("com.sun.star.text.TextField.Annotation"):
            out.append(field)
    return out


def _count_comments(doc: Any) -> int:
    return len(_annotations(doc))


def _drop_ghosts(doc: Any, expected: int) -> None:
    """Remove annotations beyond ``expected`` that a failed insert left behind."""
    extras = _annotations(doc)[expected:]
    for ghost in extras:
        ghost.dispose()


@op("writer.delete_comment", kind="writer", mutates=True)
def delete_comment(s: Session, comment_index: int) -> dict[str, Any]:
    comments = _annotations(s.doc)
    if not 0 <= comment_index < len(comments):
        raise OpError(
            f"comment_index {comment_index} is out of range ({len(comments)} comments). "
            "Use writer_get_comments.",
            protocol.E_NOT_FOUND,
        )
    comments[comment_index].dispose()
    return {"deleted_comment": comment_index}


@op("writer.insert_file", kind="writer", mutates=True)
def insert_file(
    s: Session, source_path: str, position: str = "end", block_index: int | None = None
) -> dict[str, Any]:
    """Insert the contents of another document (any format Writer can open)."""
    slot = h.Slot(s.doc, position, block_index)
    para = slot.paragraph("")
    cursor = s.doc.getText().createTextCursorByRange(para.getStart())
    try:
        cursor.insertDocumentFromURL(file_url(source_path), ())
    except Exception as exc:
        raise OpError(
            f"Could not insert {Path(source_path).name}: {exc}", protocol.E_OFFICE
        ) from exc
    return {"inserted_from": Path(source_path).name}
