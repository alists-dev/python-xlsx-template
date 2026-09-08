"""Render a boolean as a checkbox glyph inside a template expression.

Backs the ``{% yn %}`` tag: ``{% yn approved %}`` becomes a checked or
unchecked box character, picked from the cell's existing font -- no font
injection, unlike a Wingdings-specific checkmark. Any font with Unicode
"Ballot Box" support (Calibri, Segoe UI, Arial Unicode, ...) renders these
correctly.
"""

from __future__ import annotations

CHECKED = "\u2611"  # BALLOT BOX WITH CHECK (looks like a checked checkbox)
UNCHECKED = "\u2610"  # BALLOT BOX (looks like an empty checkbox)


def render_checkbox(value: object, invert: bool = False) -> str:
    """Return the checkbox glyph for ``value``, optionally inverted."""
    truthy = bool(value)
    if invert:
        truthy = not truthy
    return CHECKED if truthy else UNCHECKED
