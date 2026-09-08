"""Regex patterns for xlsx-jinja's custom template-tag syntax.

Split out of ``template.py`` so the rendering pipeline and the read-only
``introspection`` module (which needs to recognize the same tags without
rendering anything) share one definition instead of two drifting copies.
"""

from __future__ import annotations

import re

JINJA_RE = re.compile(r"({[{%#].*?[}%#]})", re.DOTALL)
XV_RE = re.compile(r"{%\s*xv\s+(.+?)\s*%}", re.DOTALL | re.IGNORECASE)
IMAGE_TAG_RE = re.compile(
    r"{%\s*(img|insert_img)\s+(.+?)\s*%}", re.DOTALL | re.IGNORECASE
)
STRUCTURAL_ROW_RE = re.compile(r"^\s*{%\s*r\s+(.+?)\s*%}\s*$", re.DOTALL | re.IGNORECASE)
BLOCK_MARKER_RE = re.compile(
    r"^\s*{%\s*b\s+([A-Z]{1,3}:[A-Z]{1,3})\s+(.+?)\s*%}\s*$",
    re.DOTALL | re.IGNORECASE,
)
TR_RE = re.compile(r"{%\s*tr\s*%}", re.IGNORECASE)


def normalize_custom_tags(text: str) -> str:
    """Rewrite xlsx-jinja-only tags into plain Jinja syntax.

    Used for introspection (``meta.find_undeclared_variables``), which only
    understands standard Jinja tags. Not used by the render pipeline, which
    has its own token-based handling for ``{% xv %}``/``{% img %}`` because
    it also needs to know *where* to write the resulting value back.
    """
    match = STRUCTURAL_ROW_RE.match(text)
    if match:
        return "{% " + match.group(1) + " %}"
    match = BLOCK_MARKER_RE.match(text)
    if match:
        return "{% " + match.group(2) + " %}"
    text = TR_RE.sub("", text)
    text = XV_RE.sub(r"{{ \1 }}", text)
    text = IMAGE_TAG_RE.sub(r"{{ \2 }}", text)
    return text
