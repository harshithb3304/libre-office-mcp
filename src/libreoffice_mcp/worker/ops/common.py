"""Operations that work on any document type (``doc.*``)."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp import formats
from libreoffice_mcp.worker.registry import op
from libreoffice_mcp.worker.session import Session


def iso_datetime(value: Any) -> str | None:
    """UNO ``DateTime`` struct -> ISO-8601 string (None when unset)."""
    if value is None or getattr(value, "Year", 0) == 0:
        return None
    return (
        f"{value.Year:04d}-{value.Month:02d}-{value.Day:02d}"
        f"T{value.Hours:02d}:{value.Minutes:02d}:{value.Seconds:02d}"
    )


def _count(collection_getter: Any) -> int | None:
    try:
        return collection_getter().getCount()
    except Exception:
        return None


def _stats(s: Session) -> dict[str, Any]:
    doc = s.doc
    stats: dict[str, Any] = {}
    if s.kind == formats.WRITER:
        stats["paragraphs"] = doc.ParagraphCount
        stats["words"] = doc.WordCount
        stats["characters"] = doc.CharacterCount
        stats["tables"] = _count(doc.getTextTables)
        stats["images"] = _count(doc.getGraphicObjects)
        stats["pages"] = _writer_page_count(doc)
    elif s.kind == formats.CALC:
        names = list(doc.getSheets().getElementNames())
        stats["sheets"] = names
    elif s.kind in (formats.IMPRESS, formats.DRAW):
        stats["slides" if s.kind == formats.IMPRESS else "pages"] = doc.getDrawPages().getCount()
    return stats


def _writer_page_count(doc: Any) -> int | None:
    """Page count via the (hidden) view cursor; None if layout is unavailable."""
    try:
        cursor = doc.getCurrentController().getViewCursor()
        cursor.jumpToLastPage()
        return cursor.getPage()
    except Exception:
        return None


@op("doc.info")
def info(s: Session) -> dict[str, Any]:
    """Document kind, metadata and basic statistics."""
    props = s.doc.getDocumentProperties()
    return {
        "file": s.path.name,
        "kind": s.kind,
        "format": formats.ext_of(str(s.path)),
        "title": props.Title,
        "author": props.Author,
        "subject": props.Subject,
        "description": props.Description,
        "keywords": list(props.Keywords or ()),
        "created": iso_datetime(props.CreationDate),
        "modified": iso_datetime(props.ModificationDate),
        "statistics": _stats(s),
    }


@op("doc.set_metadata", mutates=True)
def set_metadata(
    s: Session,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    description: str | None = None,
    keywords: list[str] | None = None,
) -> dict[str, Any]:
    """Set document properties; omitted fields are left unchanged."""
    props = s.doc.getDocumentProperties()
    changed = []
    for name, value in (
        ("Title", title),
        ("Author", author),
        ("Subject", subject),
        ("Description", description),
    ):
        if value is not None:
            setattr(props, name, value)
            changed.append(name.lower())
    if keywords is not None:
        props.Keywords = tuple(keywords)
        changed.append("keywords")
    return {"updated": changed}
