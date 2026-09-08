"""Swap embedded media by CRC32, without touching styling or positions.

Mirrors docxtpl's ``replace_media``: useful when the only thing that
changes between renders is an image (e.g. a per-tenant logo) and a full
render pass isn't needed just to update ``xl/media/*`` bytes.
"""

from __future__ import annotations

import binascii


def compute_crc32(data: bytes) -> int:
    return binascii.crc32(data) & 0xFFFFFFFF


def apply_media_replacements(
    entries: dict[str, bytes],
    replacements: dict[int, bytes],
    media_prefix: str = "xl/media/",
) -> dict[str, bytes]:
    """Return a copy of ``entries`` with matching media parts swapped.

    A part is swapped when its current CRC32 matches a key in
    ``replacements`` (as registered by ``XlsxTemplate.replace_media``).
    """
    if not replacements:
        return entries
    updated = dict(entries)
    for name, data in entries.items():
        if not name.startswith(media_prefix):
            continue
        crc = compute_crc32(data)
        if crc in replacements:
            updated[name] = replacements[crc]
    return updated
