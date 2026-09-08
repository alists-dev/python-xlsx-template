"""Excel rich-text runs, built from Python and dropped into a template.

Mirrors docxtpl's ``RichText``: build one, put it straight into a template
expression (``{{ rt }}``), and xlsx-jinja replaces the whole cell with the
styled runs. The object must be the entire content of its cell -- the same
constraint ``{% xv %}`` already has, since the cell's value type has to be
picked before Jinja renders down to plain text.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Callable
from xml.etree import ElementTree as ET

from ._xml import q

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
RICHTEXT_TOKEN_RE = re.compile(r"__XLSX_JINJA_RICHTEXT_(\d+)__")


def _to_argb(color: str) -> str:
    color = color.lstrip("#").upper()
    return color if len(color) == 8 else f"FF{color}"


@dataclass
class _Run:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool | str = False
    strike: bool = False
    color: str | None = None
    size: int | float | None = None
    font: str | None = None
    subscript: bool = False
    superscript: bool = False

    def to_element(self) -> ET.Element:
        run = ET.Element(q("r"))
        properties = ET.Element(q("rPr"))
        if self.bold:
            ET.SubElement(properties, q("b"))
        if self.italic:
            ET.SubElement(properties, q("i"))
        if self.strike:
            ET.SubElement(properties, q("strike"))
        if self.underline:
            attrs = {"val": self.underline} if isinstance(self.underline, str) else {}
            ET.SubElement(properties, q("u"), attrs)
        if self.subscript:
            ET.SubElement(properties, q("vertAlign"), {"val": "subscript"})
        if self.superscript:
            ET.SubElement(properties, q("vertAlign"), {"val": "superscript"})
        if self.size:
            ET.SubElement(properties, q("sz"), {"val": str(self.size)})
        if self.color:
            ET.SubElement(properties, q("color"), {"rgb": _to_argb(self.color)})
        if self.font:
            ET.SubElement(properties, q("rFont"), {"val": self.font})
        if len(properties):
            run.append(properties)
        text_node = ET.SubElement(run, q("t"))
        if self.text[:1].isspace() or self.text[-1:].isspace():
            text_node.set(_XML_SPACE, "preserve")
        text_node.text = self.text
        return run


class RichText:
    """A sequence of differently-styled text runs for one cell."""

    def __init__(self, text: str | None = None, **run_kwargs: object) -> None:
        self._runs: list[_Run] = []
        if text is not None:
            self.add(text, **run_kwargs)

    def add(
        self,
        text: object,
        *,
        bold: bool = False,
        italic: bool = False,
        underline: bool | str = False,
        strike: bool = False,
        color: str | None = None,
        size: int | float | None = None,
        font: str | None = None,
        subscript: bool = False,
        superscript: bool = False,
    ) -> "RichText":
        """Append one styled run. Returns ``self`` for chaining."""
        self._runs.append(
            _Run(
                text=str(text),
                bold=bold,
                italic=italic,
                underline=underline,
                strike=strike,
                color=color,
                size=size,
                font=font,
                subscript=subscript,
                superscript=superscript,
            )
        )
        return self

    def to_elements(self) -> list[ET.Element]:
        """Build the ``<r>`` run elements for this rich text."""
        if not self._runs:
            raise ValueError("RichText has no content; call add() first.")
        return [run.to_element() for run in self._runs]

    def __str__(self) -> str:
        return "".join(run.text for run in self._runs)

    def __repr__(self) -> str:
        return f"RichText({str(self)!r})"


def make_richtext_finalize(
    sink: list[RichText], inner: Callable[[Any], Any] | None = None
) -> Callable[[Any], Any]:
    """Build a Jinja ``finalize`` hook that tokenizes ``RichText`` values.

    A ``RichText`` placed directly in ``{{ ... }}`` is captured into ``sink``
    and replaced by a placeholder token, so the worksheet renderer can swap
    the whole cell for real ``<r>`` runs once the row's final position is
    known. Any other value is passed through -- to ``inner`` if the caller
    (or a user-supplied ``jinja_env``) already had its own ``finalize``.
    """

    def finalize(value: Any) -> Any:
        if isinstance(value, RichText):
            sink.append(value)
            return f"__XLSX_JINJA_RICHTEXT_{len(sink) - 1}__"
        return inner(value) if inner else value

    return finalize


def apply_richtext_cell(
    cell: ET.Element, text: str, richtext_values: list[RichText]
) -> bool:
    """Replace ``cell``'s inline string with rich-text runs if it holds one.

    Returns ``True`` when a replacement was made. Raises ``ValueError`` when
    a richtext token is mixed with other text -- a ``RichText`` value must be
    the entire content of its cell, the same constraint ``{% xv %}`` has.
    """
    stripped = text.strip()
    match = RICHTEXT_TOKEN_RE.fullmatch(stripped)
    if match is None:
        if RICHTEXT_TOKEN_RE.search(text):
            raise ValueError(
                "RichText must be the entire content of its cell, "
                "not mixed with other text."
            )
        return False
    rich_text = richtext_values[int(match.group(1))]
    for child in list(cell):
        if child.tag != q("f"):
            cell.remove(child)
    inline = ET.SubElement(cell, q("is"))
    for run in rich_text.to_elements():
        inline.append(run)
    return True
