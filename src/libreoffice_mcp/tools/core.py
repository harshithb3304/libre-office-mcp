"""Core tools: work on any document type and are always available.

(Tools that need no LibreOffice at all - listing and copying files - run
directly in the server process and so stay instant.)
"""

from __future__ import annotations

import asyncio
import shutil
from datetime import datetime
from typing import Annotated, Any

from pydantic import BaseModel, Field

from libreoffice_mcp import formats, protocol
from libreoffice_mcp.errors import OfficeError
from libreoffice_mcp.state import state
from libreoffice_mcp.tools.base import batch, read
from libreoffice_mcp.tools.registry import CORE, tool
from libreoffice_mcp.tools.types import DocPath

_MAX_LISTED = 200
_MAX_BATCH = 100


@tool(CORE, "create", title="Create document")
async def office_create_document(
    path: DocPath,
    title: str | None = None,
    author: str | None = None,
    overwrite: Annotated[bool, Field(description="Replace the file if it already exists.")] = False,
) -> dict[str, Any]:
    """Create a new, empty document. The type is chosen by the file extension:
    .odt/.docx (Writer text), .ods/.xlsx (Calc spreadsheet), .odp/.pptx (Impress presentation).
    Use the writer_*/calc_*/impress_* tools afterwards to fill it."""
    target = state.paths.resolve(path)
    ops = []
    if title or author:
        args = {k: v for k, v in (("title", title), ("author", author)) if v}
        ops.append({"op": "doc.set_metadata", "args": args})
    reply = await state.office.call(
        protocol.KIND_CREATE, target, ops, options={"overwrite": overwrite}
    )
    return {"file": str(target), "kind": reply["kind"], "created": True}


@tool(CORE, "read", title="Get document info")
async def office_get_document_info(path: DocPath) -> dict[str, Any]:
    """Return a document's type, metadata (title, author, dates) and statistics
    (pages, words, tables... for Writer; sheet names for Calc; slide count for Impress)."""
    return await read(path, "doc.info")


@tool(CORE, "edit", title="Set document properties")
async def office_set_document_properties(
    path: DocPath,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    description: str | None = None,
    keywords: list[str] | None = None,
) -> dict[str, Any]:
    """Set document metadata. Fields you omit are left unchanged."""
    return await batch(
        path,
        [("doc.set_metadata", dict(title=title, author=author, subject=subject,
                                   description=description, keywords=keywords))],
    )  # fmt: skip


@tool(CORE, "read", title="List documents")
async def office_list_documents(
    directory: Annotated[str, Field(description="Folder to scan (absolute or relative).")],
    recursive: bool = False,
) -> dict[str, Any]:
    """List LibreOffice/Office documents (odt, docx, ods, xlsx, odp, pptx, csv...) in a folder."""
    folder = state.paths.existing_dir(directory)

    def scan() -> list[dict[str, Any]]:
        entries = []
        iterator = folder.rglob("*") if recursive else folder.iterdir()
        for p in iterator:
            name = p.name
            # Skip hidden files, Office "~$" owner files and LibreOffice ".~lock" files.
            if name.startswith((".", "~$")) or not p.is_file():
                continue
            fmt = formats.BY_EXT.get(formats.ext_of(name))
            if fmt is None:
                continue
            st = p.stat()
            entries.append(
                {
                    "path": str(p),
                    "type": fmt.kind,
                    "size_bytes": st.st_size,
                    "modified": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                }
            )
        entries.sort(key=lambda e: e["path"].lower())
        return entries

    entries = await asyncio.to_thread(scan)
    return {
        "folder": str(folder),
        "count": len(entries),
        "truncated": len(entries) > _MAX_LISTED,
        "documents": entries[:_MAX_LISTED],
    }


@tool(CORE, "create", title="Copy document")
async def office_copy_document(
    path: DocPath,
    destination: Annotated[str, Field(description="Where to write the copy.")],
    overwrite: bool = False,
) -> dict[str, Any]:
    """Copy a document file (a byte-for-byte copy; format is unchanged)."""
    source = state.paths.existing_file(path)
    target = state.paths.resolve(destination)
    if target.exists() and not overwrite:
        raise OfficeError(f"{target} already exists. Pass overwrite=true to replace it.",
                          protocol.E_EXISTS)  # fmt: skip
    target.parent.mkdir(parents=True, exist_ok=True)
    await asyncio.to_thread(shutil.copy2, source, target)
    return {"source": str(source), "copy": str(target)}


@tool(CORE, "create", title="Convert / export document")
async def office_convert_document(
    path: DocPath,
    output_file: Annotated[
        str,
        Field(
            description="Destination file. The extension picks the format, e.g. report.pdf, "
            "report.docx, sheet.xlsx, slides.pptx, page.html, text.txt."
        ),
    ],
    overwrite: bool = False,
    page_range: Annotated[
        str | None, Field(description="PDF only: pages to export, e.g. '1-3' or '2,5'.")
    ] = None,
    image_quality: Annotated[
        int | None, Field(ge=1, le=100, description="PDF only: JPEG quality 1-100.")
    ] = None,
) -> dict[str, Any]:
    """Convert a document to another format (PDF export, docx<->odt, xlsx<->ods, pptx<->odp,
    HTML, plain text...). The source file is not modified. A document can only be converted
    to formats of its own application, plus PDF."""
    source = state.paths.supported_file(path)
    target = state.paths.resolve(output_file)
    filter_data = {}
    if page_range:
        filter_data["PageRange"] = page_range
    if image_quality:
        filter_data["Quality"] = image_quality
    reply = await state.office.call(
        protocol.KIND_EXPORT,
        source,
        target=target,
        options={"overwrite": overwrite, "filter_data": filter_data},
    )
    return {"output": reply["output"], "bytes": reply["bytes"]}


class Operation(BaseModel):
    op: str = Field(
        description="Operation name: the tool name with the first '_' replaced by '.', "
        "e.g. tool writer_add_paragraph -> 'writer.add_paragraph'."
    )
    args: dict[str, Any] = Field(
        default_factory=dict,
        description="Arguments, exactly as the matching tool takes them (without 'path').",
    )


@tool(CORE, "edit", title="Batch edit")
async def office_batch_edit(
    path: DocPath,
    operations: Annotated[list[Operation], Field(min_length=1, max_length=_MAX_BATCH)],
) -> dict[str, Any]:
    """Apply several edits to ONE document in a single step: the file is opened once, all
    operations run in order, and it is saved once. If any operation fails, NOTHING is saved.
    Much faster than separate tool calls - prefer it when making 3+ changes to a document.

    Each operation is {"op": "<app>.<name>", "args": {...}} where the name and arguments match
    the editing tool of the same name, e.g.
    {"op": "writer.add_heading", "args": {"text": "Intro", "level": 1}}."""
    allowed = {"doc", *state.settings.apps}
    ops = []
    for item in operations:
        app = item.op.split(".", 1)[0]
        if "." not in item.op or app not in allowed:
            raise OfficeError(
                f"Unknown operation '{item.op}'. Use '<app>.<name>' with app in {sorted(allowed)}."
            )
        ops.append((item.op, item.args))
    return await batch(path, ops)


@tool(CORE, "read", title="Server status")
async def office_status() -> dict[str, Any]:
    """Report the LibreOffice backend's health: version, uptime, requests served and the
    active settings. Useful when something seems wrong."""
    s = state.settings
    return {
        "worker": state.office.status(),
        "apps": list(s.apps),
        "read_only": s.read_only,
        "default_dir": str(s.default_dir),
        "allowed_roots": [str(r) for r in s.roots] or "unrestricted",
        "editable_formats": formats.editable_extensions(),
    }
