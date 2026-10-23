"""Opening, saving and closing documents safely.

Safety model (the important part of this file):

* We never let LibreOffice touch the user's file directly. Every request works on
  a *scratch copy* in a private temp dir. LibreOffice therefore never creates
  ``.~lock.*#`` files next to the user's document, and a crash or a failed
  operation cannot leave a half-written file behind.
* A successful edit is written back with :func:`atomic_replace` (copy to a temp
  file in the destination folder, then ``os.replace``). Readers see either the
  old file or the new one, never a partial one.
* Macros in documents are never executed (``MacroExecutionMode = NEVER_EXECUTE``).
"""

from __future__ import annotations

import logging
import os
import shutil
import time
import uuid
from pathlib import Path
from typing import Any

from libreoffice_mcp import formats, protocol
from libreoffice_mcp.worker.registry import OpError
from libreoffice_mcp.worker.util import file_url, prop_dict, props

log = logging.getLogger("worker.session")

# com.sun.star.document.MacroExecMode.NEVER_EXECUTE / UpdateDocMode.NO_UPDATE
_NEVER_EXECUTE = 0
_NO_UPDATE = 0

_SERVICE_TO_KIND = (
    ("com.sun.star.text.TextDocument", formats.WRITER),
    ("com.sun.star.text.WebDocument", formats.WRITER),
    ("com.sun.star.sheet.SpreadsheetDocument", formats.CALC),
    ("com.sun.star.presentation.PresentationDocument", formats.IMPRESS),
    ("com.sun.star.drawing.DrawingDocument", formats.DRAW),
)


class Session:
    """An open document plus the helpers operations need.

    Operations receive a ``Session`` as their first argument. They read/modify
    ``session.doc`` and never save or close it - the runner does that.
    """

    def __init__(self, office: Any, doc: Any, path: Path, read_only: bool):
        self.office = office
        self.doc = doc
        self.path = path
        self.read_only = read_only
        self.kind = detect_kind(doc)
        self.mutated = False

    def require(self, kind: str) -> None:
        if self.kind != kind:
            raise OpError(
                f"This operation needs a {kind} document, but {self.path.name} is a "
                f"{self.kind} document.",
                protocol.E_WRONG_DOC_TYPE,
            )

    def new(self, service: str) -> Any:
        """Create a document-bound UNO object, e.g. ``com.sun.star.text.TextTable``."""
        return self.doc.createInstance(service)

    def service(self, name: str) -> Any:
        """Create a global UNO service, e.g. ``com.sun.star.frame.DispatchHelper``."""
        return self.office.context.ServiceManager.createInstanceWithContext(
            name, self.office.context
        )


def detect_kind(doc: Any) -> str:
    for service, kind in _SERVICE_TO_KIND:
        if doc.supportsService(service):
            return kind
    raise OpError("Unsupported document type", protocol.E_UNSUPPORTED)


# -- loading / closing ---------------------------------------------------------


def load_document(office: Any, file: Path, *, read_only: bool, password: str | None = None) -> Any:
    fmt = formats.BY_EXT.get(formats.ext_of(str(file)))
    load_props: dict[str, Any] = {
        "Hidden": True,
        "ReadOnly": read_only,
        "MacroExecutionMode": _NEVER_EXECUTE,
        "UpdateDocMode": _NO_UPDATE,
    }
    if fmt and fmt.filter_options and fmt.ext in ("csv", "txt"):
        load_props["FilterOptions"] = fmt.filter_options
    if password:
        load_props["Password"] = password
    try:
        doc = office.desktop.loadComponentFromURL(
            file_url(str(file)), "_blank", 0, props(**load_props)
        )
    except Exception as exc:
        raise OpError(
            f"LibreOffice could not open the file: {describe(exc)}", protocol.E_OFFICE
        ) from exc
    if doc is None:
        raise OpError(
            "LibreOffice could not open the file (unsupported, corrupt, or password protected).",
            protocol.E_OFFICE,
        )
    return doc


def close_document(doc: Any) -> None:
    if doc is None:
        return
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            log.warning("could not close document", exc_info=True)


def describe(exc: BaseException) -> str:
    """Readable message for UNO exceptions (which often have an empty ``str``)."""
    message = getattr(exc, "Message", "") or str(exc)
    return f"{type(exc).__name__}: {message}".strip(": ")


# -- saving --------------------------------------------------------------------


def store_in_place(doc: Any, path: Path) -> None:
    """Save ``doc`` back to the file it was loaded from (``path``), keeping its format."""
    ext = formats.ext_of(str(path))
    fmt = formats.BY_EXT.get(ext)
    if fmt and ext in ("csv", "txt"):
        # store() would use default text options; be explicit so round-trips are lossless.
        doc.storeToURL(
            file_url(str(path)),
            props(FilterName=fmt.filter, FilterOptions=fmt.filter_options, Overwrite=True),
        )
        return
    # Reuse the filter the document was loaded with (e.g. "MS Word 2007 XML" for .docx).
    filter_name = prop_dict(doc.getArgs()).get("FilterName")
    if filter_name:
        doc.storeToURL(file_url(str(path)), props(FilterName=filter_name, Overwrite=True))
    else:
        doc.store()


def atomic_replace(source: Path, destination: Path) -> None:
    """Move ``source`` over ``destination`` so readers never observe a partial file.

    The data is first copied to a hidden temp file *in the destination folder*
    (``os.replace`` is only atomic within one filesystem), then renamed over the
    target. Bind-mounted host folders (Docker Desktop) can briefly lock files, so we retry.
    """
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp = destination.with_name(f".{destination.name}.{uuid.uuid4().hex[:8]}.tmp")
    shutil.copyfile(source, temp)
    try:
        if destination.exists():
            shutil.copymode(destination, temp)
        last: OSError | None = None
        for _ in range(10):
            try:
                os.replace(temp, destination)
                return
            except PermissionError as exc:  # file briefly in use - wait and retry
                last = exc
                time.sleep(0.2)
        raise OpError(
            f"Could not write {destination}: the file is in use by another program ({last}).",
            protocol.E_OFFICE,
        )
    finally:
        temp.unlink(missing_ok=True)


def copy_with_retry(source: Path, destination: Path) -> None:
    """Copy the user's file into scratch, tolerating a concurrent atomic replace."""
    last: OSError | None = None
    for _ in range(10):
        try:
            shutil.copyfile(source, destination)
            return
        except PermissionError as exc:
            last = exc
            time.sleep(0.2)
    raise OpError(f"Could not read {source}: {last}", protocol.E_OFFICE)
