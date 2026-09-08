"""Shared OOXML namespace constants and qualified-name helpers.

Kept dependency-free (no imports from sibling modules) so every other
module in this package -- ``template``, ``richtext``, ``media``,
``introspection`` -- can import from here without risking a circular
import.
"""

from __future__ import annotations

from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {
    "x": MAIN_NS,
    "r": REL_NS,
    "pr": PACKAGE_REL_NS,
    "xdr": DRAWING_NS,
    "a": DRAWING_MAIN_NS,
}

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)
ET.register_namespace("xdr", DRAWING_NS)
ET.register_namespace("a", DRAWING_MAIN_NS)


def q(name: str) -> str:
    """Qualify ``name`` in the main spreadsheetML namespace."""
    return f"{{{MAIN_NS}}}{name}"


def qn(namespace: str, name: str) -> str:
    """Qualify ``name`` in an arbitrary namespace."""
    return f"{{{namespace}}}{name}"
