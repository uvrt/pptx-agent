"""Rendering slides in a worker process: the function the tool layer's pool runs.

Module-level, so a spawned worker can import it; it gets the deck's bytes and returns PNGs.
"""

from __future__ import annotations

import struct


def render_slides(data: bytes, numbers: list[int], width: int) -> list[bytes]:
    """PNG bytes of the 1-based slides ``numbers`` of the deck ``data``, ``width`` px wide."""
    from ..edit.document import Document

    deck = Document.open(data)
    return list(deck.render_png(list(numbers), width=int(width)))


def png_size(png: bytes) -> tuple[int, int]:
    """A PNG's width and height, read from its header."""
    if png[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    return struct.unpack(">II", png[16:24])
