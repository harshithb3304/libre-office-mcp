"""Writer: change tracking."""

from __future__ import annotations

from typing import Any

from libreoffice_mcp.worker.registry import op
from libreoffice_mcp.worker.session import Session


@op("writer.set_track_changes", kind="writer", mutates=True)
def set_track_changes(s: Session, enabled: bool) -> dict[str, Any]:
    """Turn change tracking on/off. While on, later edits are recorded as reviewable changes."""
    s.doc.RecordChanges = bool(enabled)
    return {"tracking_enabled": bool(enabled)}


def _dispatch(s: Session, command: str) -> None:
    helper = s.service("com.sun.star.frame.DispatchHelper")
    helper.executeDispatch(s.doc.getCurrentController().getFrame(), command, "", 0, ())


@op("writer.accept_all_changes", kind="writer", mutates=True)
def accept_all_changes(s: Session) -> dict[str, Any]:
    """Accept every tracked change (the edits become permanent)."""
    before = s.doc.getRedlines().getCount()
    _dispatch(s, ".uno:AcceptAllTrackedChanges")
    return {"accepted": before, **({"_noop": True} if before == 0 else {})}


@op("writer.reject_all_changes", kind="writer", mutates=True)
def reject_all_changes(s: Session) -> dict[str, Any]:
    """Reject every tracked change (reverts the edits)."""
    before = s.doc.getRedlines().getCount()
    _dispatch(s, ".uno:RejectAllTrackedChanges")
    return {"rejected": before, **({"_noop": True} if before == 0 else {})}
