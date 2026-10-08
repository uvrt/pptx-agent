"""Test pictures, generated here rather than committed: every byte of them is ours."""

from __future__ import annotations

import struct
import zlib
from typing import Callable


def png(width: int, height: int, pixel: Callable[[int, int], tuple[int, int, int]]) -> bytes:
    """An RGB PNG whose pixel ``(x, y)`` is ``pixel(x, y)``."""
    raw = b"".join(
        b"\x00" + b"".join(bytes(pixel(x, y)) for x in range(width)) for y in range(height)
    )

    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


#: A 64x32 gradient with a red block in its top-left corner: asymmetric both ways, so a flip
#: or a quarter turn shows up in a render.
MARKER = png(64, 32, lambda x, y: (255, 0, 0) if x < 16 and y < 12 else (x * 4, y * 8, 160))

#: A plain orange 48x48 square -- a second, different image to replace with.
ORANGE = png(48, 48, lambda x, y: (240, 140, 20))
