"""Small UNO helpers shared by all operations."""

from __future__ import annotations

import re
from typing import Any

import uno
from com.sun.star.beans import PropertyValue

from libreoffice_mcp.worker.registry import OpError


def pv(name: str, value: Any) -> PropertyValue:
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def props(**kwargs: Any) -> tuple[PropertyValue, ...]:
    return tuple(pv(k, v) for k, v in kwargs.items())


def prop_dict(seq: Any) -> dict[str, Any]:
    return {p.Name: p.Value for p in (seq or ())}


def file_url(path: str) -> str:
    return uno.systemPathToFileUrl(path)


def struct_any(type_name: str, value: Any) -> Any:
    """Wrap ``value`` as a typed UNO Any (needed for sequence-valued properties)."""
    return uno.Any(type_name, value)


def const(name: str) -> Any:
    """Look up a UNO constant/enum value by its full dotted name."""
    return uno.getConstantByName(name)


def enum(type_name: str, value: str) -> Any:
    return uno.Enum(type_name, value)


_NAMED_COLORS = {
    "black": 0x000000, "white": 0xFFFFFF, "red": 0xFF0000, "green": 0x008000,
    "blue": 0x0000FF, "yellow": 0xFFFF00, "orange": 0xFFA500, "purple": 0x800080,
    "gray": 0x808080, "grey": 0x808080, "lightgray": 0xD3D3D3, "lightgrey": 0xD3D3D3,
    "darkgray": 0xA9A9A9, "darkgrey": 0xA9A9A9, "cyan": 0x00FFFF, "magenta": 0xFF00FF,
    "brown": 0xA52A2A, "pink": 0xFFC0CB, "navy": 0x000080, "teal": 0x008080,
}  # fmt: skip
_HEX = re.compile(r"^#?([0-9a-fA-F]{6})$")
_HEX3 = re.compile(r"^#([0-9a-fA-F]{3})$")


def parse_color(value: Any) -> int:
    """Accept ``#RRGGBB``, ``#RGB``, a CSS-ish colour name, or a raw integer."""
    if isinstance(value, bool):
        raise OpError(f"Invalid colour {value!r}")
    if isinstance(value, int):
        return value & 0xFFFFFF
    if isinstance(value, str):
        text = value.strip()
        m = _HEX.match(text)
        if m:
            return int(m.group(1), 16)
        m = _HEX3.match(text)
        if m:
            return int("".join(c * 2 for c in m.group(1)), 16)
        named = _NAMED_COLORS.get(text.lower())
        if named is not None:
            return named
    raise OpError(f"Invalid colour {value!r}: use '#RRGGBB' or a basic colour name")


def check_range(value: int, low: int, high: int, what: str) -> int:
    if not low <= value <= high:
        raise OpError(f"{what} {value} is out of range ({low}..{high})")
    return value
