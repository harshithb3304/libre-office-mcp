"""File-format knowledge: extension -> document kind, and LibreOffice filter names.

Stdlib-only; shared by the server (path validation, tool help) and the worker
(load/store filters).
"""

from __future__ import annotations

from dataclasses import dataclass

# Document kinds, named after the LibreOffice application that edits them.
WRITER = "writer"
CALC = "calc"
IMPRESS = "impress"
DRAW = "draw"

APPS = (WRITER, CALC, IMPRESS)

FACTORY_URL = {
    WRITER: "private:factory/swriter",
    CALC: "private:factory/scalc",
    IMPRESS: "private:factory/simpress",
    DRAW: "private:factory/sdraw",
}


@dataclass(frozen=True)
class Format:
    ext: str
    kind: str
    label: str
    #: Filter used when *storing* a document to this format (None: cannot export).
    filter: str | None
    #: Extra ``FilterOptions`` for load/store (e.g. CSV tokens).
    filter_options: str | None = None
    #: True when we can create/edit documents in this format (not export-only).
    editable: bool = True


_FORMATS: list[Format] = [
    # Writer
    Format("odt", WRITER, "OpenDocument Text", "writer8"),
    Format("ott", WRITER, "OpenDocument Text Template", "writer8_template"),
    Format("fodt", WRITER, "Flat OpenDocument Text", "OpenDocument Text Flat XML"),
    Format("docx", WRITER, "DOCX", "MS Word 2007 XML"),
    Format("doc", WRITER, "DOC", "MS Word 97"),
    Format("rtf", WRITER, "Rich Text", "Rich Text Format"),
    Format("txt", WRITER, "Plain text", "Text (encoded)", "UTF8,LF,,,", True),
    Format("html", WRITER, "HTML", "HTML (StarWriter)", None, False),
    Format("epub", WRITER, "EPUB", "EPUB", None, False),
    # Calc
    Format("ods", CALC, "OpenDocument Spreadsheet", "calc8"),
    Format("ots", CALC, "OpenDocument Spreadsheet Template", "calc8_template"),
    Format("fods", CALC, "Flat OpenDocument Spreadsheet", "OpenDocument Spreadsheet Flat XML"),
    Format("xlsx", CALC, "XLSX", "Calc MS Excel 2007 XML"),
    Format("xls", CALC, "XLS", "MS Excel 97"),
    Format(
        "csv",
        CALC,
        "CSV",
        "Text - txt - csv (StarCalc)",
        "44,34,76,1,,0,false,true,false,false",
        True,
    ),
    # Impress
    Format("odp", IMPRESS, "OpenDocument Presentation", "impress8"),
    Format("otp", IMPRESS, "OpenDocument Presentation Template", "impress8_template"),
    Format("fodp", IMPRESS, "Flat OpenDocument Presentation", "OpenDocument Presentation Flat XML"),
    Format("pptx", IMPRESS, "PPTX", "Impress MS PowerPoint 2007 XML"),
    Format("ppt", IMPRESS, "PPT", "MS PowerPoint 97"),
    # Draw
    Format("odg", DRAW, "OpenDocument Drawing", "draw8"),
]

#: Extension -> format for formats that can be opened and edited.
BY_EXT: dict[str, Format] = {f.ext: f for f in _FORMATS}

# Export-only targets, keyed by (source kind, extension)
_EXPORT_ONLY: dict[tuple[str, str], str] = {
    (WRITER, "pdf"): "writer_pdf_Export",
    (CALC, "pdf"): "calc_pdf_Export",
    (IMPRESS, "pdf"): "impress_pdf_Export",
    (DRAW, "pdf"): "draw_pdf_Export",
    (CALC, "html"): "HTML (StarCalc)",
    (IMPRESS, "html"): "impress_html_Export",
    (WRITER, "png"): "writer_png_Export",
    (CALC, "png"): "calc_png_Export",
    (IMPRESS, "png"): "impress_png_Export",
    (DRAW, "png"): "draw_png_Export",
    (IMPRESS, "svg"): "impress_svg_Export",
    (DRAW, "svg"): "draw_svg_Export",
    (WRITER, "xhtml"): "XHTML Writer File",
}


def ext_of(path: str) -> str:
    """Lower-case extension without the dot ('' when there is none)."""
    name = path.replace("\\", "/").rsplit("/", 1)[-1]
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def kind_for(path: str) -> str | None:
    """Document kind for an editable path, or None when the extension is unknown."""
    fmt = BY_EXT.get(ext_of(path))
    return fmt.kind if fmt else None


def export_filter(source_kind: str, target_ext: str) -> tuple[str, str | None] | None:
    """Return ``(filter_name, filter_options)`` for storing ``source_kind`` as ``target_ext``.

    A document can only be stored as a format its application supports; Calc
    cannot write ``.docx``. Returns None when that combination is unsupported.
    """
    ext = target_ext.lower().lstrip(".")
    fmt = BY_EXT.get(ext)
    if fmt is not None and fmt.filter is not None:
        # Same-application formats, plus the Draw/Impress family which share storage.
        same_family = fmt.kind == source_kind or {fmt.kind, source_kind} <= {IMPRESS, DRAW}
        if same_family:
            return fmt.filter, fmt.filter_options
    flt = _EXPORT_ONLY.get((source_kind, ext))
    if flt:
        return flt, None
    return None


def supported_exports(source_kind: str) -> list[str]:
    """Extensions a document of ``source_kind`` can be converted to."""
    exts = {e for e in BY_EXT if export_filter(source_kind, e) is not None}
    exts |= {e for (k, e) in _EXPORT_ONLY if k == source_kind}
    return sorted(exts)


def editable_extensions() -> list[str]:
    return sorted(e for e, f in BY_EXT.items() if f.editable)
