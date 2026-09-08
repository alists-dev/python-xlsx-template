"""Read-only discovery of Jinja variables referenced by a template.

Mirrors docxtpl's ``get_undeclared_template_variables``.
"""

from __future__ import annotations

from typing import Iterable

from jinja2 import Environment, TemplateSyntaxError, meta

from ._tags import JINJA_RE, normalize_custom_tags


def find_undeclared_variables(
    sheet_texts: Iterable[Iterable[str]],
    jinja_env: Environment | None = None,
) -> set[str]:
    """Collect variable names referenced across worksheet cell text.

    ``sheet_texts`` groups cell text by sheet, in document order, so tags
    split across cells/rows (loops, conditions) still balance when parsed
    together. A sheet whose tags don't balance into valid Jinja (e.g. a
    ``{%r for %}`` marker considered on its own) is skipped rather than
    failing the whole scan.
    """
    env = jinja_env or Environment()
    variables: set[str] = set()
    for texts in sheet_texts:
        relevant = [text for text in texts if JINJA_RE.search(text)]
        if not relevant:
            continue
        source = "\n".join(normalize_custom_tags(text) for text in relevant)
        try:
            variables |= meta.find_undeclared_variables(env.parse(source))
        except TemplateSyntaxError:
            continue
    return variables
