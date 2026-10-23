"""Execute one ``run`` request: open -> run operations -> (save) -> close.

Four request kinds (see :mod:`libreoffice_mcp.protocol`):

``read``    open a scratch copy read-only and run read-only ops
``edit``    open a scratch copy, run ops, save atomically *only if all ops succeeded*
``create``  start from a blank document, run ops, save to a new file
``export``  open a scratch copy read-only and store it in another format

A batch of ops is all-or-nothing: if op #3 fails, nothing is written.
"""

from __future__ import annotations

import logging
import tempfile
import time
from pathlib import Path
from typing import Any

from libreoffice_mcp import formats, protocol
from libreoffice_mcp.worker import session as sess
from libreoffice_mcp.worker.registry import OPS, OpError
from libreoffice_mcp.worker.util import file_url, props, struct_any

log = logging.getLogger("worker.runner")


def execute(office: Any, request: dict[str, Any]) -> dict[str, Any]:
    kind = request.get("kind")
    path = Path(request["path"])
    ops = request.get("ops") or []
    options = request.get("options") or {}
    started = time.monotonic()

    if kind == protocol.KIND_READ:
        result = _read_or_edit(office, path, ops, options, edit=False)
    elif kind == protocol.KIND_EDIT:
        result = _read_or_edit(office, path, ops, options, edit=True)
    elif kind == protocol.KIND_CREATE:
        result = _create(office, path, ops, options)
    elif kind == protocol.KIND_EXPORT:
        result = _export(office, path, Path(request["target"]), options)
    else:
        raise OpError(f"Unknown request kind {kind!r}", protocol.E_INTERNAL)

    result["elapsed_ms"] = int((time.monotonic() - started) * 1000)
    return result


# -- helpers -------------------------------------------------------------------


def _require_file(path: Path) -> None:
    if not path.is_file():
        raise OpError(f"File not found: {path}", protocol.E_NOT_FOUND)
    if formats.ext_of(str(path)) not in formats.BY_EXT:
        raise OpError(
            f"Unsupported file type '.{formats.ext_of(str(path))}'. "
            f"Supported: {', '.join(sorted(formats.BY_EXT))}",
            protocol.E_UNSUPPORTED,
        )


def run_ops(session: sess.Session, ops: list[dict[str, Any]]) -> list[Any]:
    """Run operations in order; the first failure aborts the whole batch."""
    results: list[Any] = []
    for index, item in enumerate(ops):
        name = item.get("op", "")
        definition = OPS.get(name)
        try:
            if definition is None:
                raise OpError(f"Unknown operation {name!r}", protocol.E_UNSUPPORTED)
            if definition.doc_kind:
                session.require(definition.doc_kind)
            if definition.mutates and session.read_only:
                raise OpError(f"Operation {name} modifies the document", protocol.E_INTERNAL)
            result = definition.fn(session, **(item.get("args") or {}))
        except OpError as exc:
            exc.op_index, exc.op_name = index, name  # type: ignore[attr-defined]
            raise
        except TypeError as exc:  # wrong/missing keyword arguments
            err = OpError(f"Invalid arguments for {name}: {exc}")
            err.op_index, err.op_name = index, name  # type: ignore[attr-defined]
            raise err from exc
        except Exception as exc:
            log.exception("operation %s failed", name)
            err = OpError(f"{name} failed: {sess.describe(exc)}", protocol.E_OFFICE)
            err.op_index, err.op_name = index, name  # type: ignore[attr-defined]
            raise err from exc
        if definition.mutates:
            session.mutated = True
        results.append(result)
    return results


# -- request kinds -------------------------------------------------------------


def _read_or_edit(
    office: Any, path: Path, ops: list[dict[str, Any]], options: dict[str, Any], *, edit: bool
) -> dict[str, Any]:
    _require_file(path)
    with tempfile.TemporaryDirectory(prefix="lomcp_scratch_") as scratch_dir:
        scratch = Path(scratch_dir) / path.name
        sess.copy_with_retry(path, scratch)
        doc = sess.load_document(
            office, scratch, read_only=not edit, password=options.get("password")
        )
        try:
            session = sess.Session(office, doc, path, read_only=not edit)
            results = run_ops(session, ops)
            saved = False
            if edit and session.mutated:
                sess.store_in_place(doc, scratch)
                saved = True
        finally:
            sess.close_document(doc)
        if saved:
            sess.atomic_replace(scratch, path)
    return {"results": results, "saved": saved}


def _create(
    office: Any, path: Path, ops: list[dict[str, Any]], options: dict[str, Any]
) -> dict[str, Any]:
    fmt = formats.BY_EXT.get(formats.ext_of(str(path)))
    if fmt is None or not fmt.editable:
        raise OpError(
            f"Cannot create '.{formats.ext_of(str(path))}' files. "
            f"Use one of: {', '.join(formats.editable_extensions())}",
            protocol.E_UNSUPPORTED,
        )
    if path.exists() and not options.get("overwrite"):
        raise OpError(
            f"{path} already exists. Pass overwrite=true to replace it.", protocol.E_EXISTS
        )
    try:
        doc = office.desktop.loadComponentFromURL(
            formats.FACTORY_URL[fmt.kind], "_blank", 0, props(Hidden=True)
        )
    except Exception as exc:
        raise OpError(
            f"Could not create a document: {sess.describe(exc)}", protocol.E_OFFICE
        ) from exc
    with tempfile.TemporaryDirectory(prefix="lomcp_scratch_") as scratch_dir:
        scratch = Path(scratch_dir) / path.name
        try:
            session = sess.Session(office, doc, path, read_only=False)
            results = run_ops(session, ops)
            store_props = {"FilterName": fmt.filter, "Overwrite": True}
            if fmt.filter_options and fmt.ext in ("csv", "txt"):
                store_props["FilterOptions"] = fmt.filter_options
            doc.storeToURL(file_url(str(scratch)), props(**store_props))
        finally:
            sess.close_document(doc)
        sess.atomic_replace(scratch, path)
    return {"results": results, "saved": True, "created": str(path), "kind": fmt.kind}


# Keys we allow callers to pass through to the PDF filter's FilterData.
PDF_FILTER_DATA_KEYS = {
    "PageRange", "Quality", "ReduceImageResolution", "MaxImageResolution", "ExportBookmarks",
    "ExportNotes", "UseTaggedPDF", "SelectPdfVersion", "EncryptFile", "DocumentOpenPassword",
}  # fmt: skip


def _export(office: Any, path: Path, target: Path, options: dict[str, Any]) -> dict[str, Any]:
    _require_file(path)
    if target.exists() and not options.get("overwrite"):
        raise OpError(
            f"{target} already exists. Pass overwrite=true to replace it.", protocol.E_EXISTS
        )
    with tempfile.TemporaryDirectory(prefix="lomcp_scratch_") as scratch_dir:
        scratch_in = Path(scratch_dir) / path.name
        scratch_out = Path(scratch_dir) / ("out_" + target.name)
        sess.copy_with_retry(path, scratch_in)
        doc = sess.load_document(
            office, scratch_in, read_only=True, password=options.get("password")
        )
        try:
            kind = sess.detect_kind(doc)
            found = formats.export_filter(kind, formats.ext_of(str(target)))
            if found is None:
                raise OpError(
                    f"A {kind} document cannot be converted to '.{formats.ext_of(str(target))}'. "
                    f"Supported: {', '.join(formats.supported_exports(kind))}",
                    protocol.E_UNSUPPORTED,
                )
            filter_name, filter_options = found
            store: dict[str, Any] = {"FilterName": filter_name, "Overwrite": True}
            if filter_options:
                store["FilterOptions"] = filter_options
            filter_data = {
                k: v for k, v in (options.get("filter_data") or {}).items()
                if k in PDF_FILTER_DATA_KEYS
            }  # fmt: skip
            if filter_data:
                store["FilterData"] = struct_any(
                    "[]com.sun.star.beans.PropertyValue", props(**filter_data)
                )
            try:
                doc.storeToURL(file_url(str(scratch_out)), props(**store))
            except Exception as exc:
                raise OpError(
                    f"Conversion failed: {sess.describe(exc)}", protocol.E_OFFICE
                ) from exc
        finally:
            sess.close_document(doc)
        if not scratch_out.is_file():
            raise OpError("LibreOffice produced no output file.", protocol.E_OFFICE)
        sess.atomic_replace(scratch_out, target)
    return {
        "saved": True,
        "output": str(target),
        "bytes": target.stat().st_size,
        "source_kind": kind,
    }
