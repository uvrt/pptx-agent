"""Rendering slides in a worker process: the function the tool layer's pool runs.

Module-level, so a spawned worker can import it; it gets the deck's bytes and returns PNGs.
"""

from __future__ import annotations

import struct
import warnings


def render_slides(data: bytes, numbers: list[int], width: int) -> list[bytes]:
    """PNG bytes of the 1-based slides ``numbers`` of the deck ``data``, ``width`` px wide."""
    return render_slides_reporting(data, numbers, width)[0]


def render_slides_reporting(data: bytes, numbers: list[int], width: int
                            ) -> tuple[list[bytes], list[list[dict]]]:
    """:func:`render_slides`, and for each slide the text its image leaves out: pptx2svg's
    ``glyphs-missing`` warnings -- no font the renderer has can draw it -- as
    ``{"face", "script", "sample"}``, reported on the first slide that has them."""
    from ..edit.document import Document

    deck = Document.open(data)
    found: list = []
    with warnings.catch_warnings():
        # The same facts come back in ``found``; the worker's stderr is nobody's to read.
        warnings.filterwarnings("ignore", message=".*no font this render loads")
        images = list(deck.render_png(list(numbers), width=int(width), warnings=found))
    by_slide: dict[int, list[dict]] = {}
    for warning in found:
        detail = getattr(warning, "detail", None)
        if getattr(warning, "code", None) != "glyphs-missing" or detail is None:
            continue
        by_slide.setdefault(warning.slide_number, []).append(
            {"face": detail.face, "script": detail.script, "sample": detail.sample})
    return images, [by_slide.get(number, []) for number in numbers]


def png_size(png: bytes) -> tuple[int, int]:
    """A PNG's width and height, read from its header."""
    if png[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    return struct.unpack(">II", png[16:24])
