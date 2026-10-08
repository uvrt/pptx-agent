"""A picture's native size, and replacing its image without distorting it (the trial's
finding 11: both P5 runs added a throwaway picture to learn a logo's proportions)."""

from __future__ import annotations

import struct
import zlib

import pytest

from images import MARKER, ORANGE, png
from pptx_agent import Document, ImageSize

WIDE = png(60, 24, lambda x, y: (20, 60, 200))          # 2.5:1, a logo's shape


def _with_phys(data: bytes, dpi: float) -> bytes:
    ppm = round(dpi / 0.0254)
    body = struct.pack(">IIB", ppm, ppm, 1)
    chunk = (struct.pack(">I", len(body)) + b"pHYs" + body
             + struct.pack(">I", zlib.crc32(b"pHYs" + body) & 0xFFFFFFFF))
    end_of_header = 8 + 25
    return data[:end_of_header] + chunk + data[end_of_header:]


def _picture(image=ORANGE):
    deck = Document.new()
    slide = deck.add_slide("Blank")
    picture = slide.add_picture(image, 9000000, 300000, width=731520, height=731520)
    return deck, picture


def test_image_size_gives_pixels_and_dpi():
    deck, picture = _picture(_with_phys(WIDE, 144))
    size = picture.image_size
    assert isinstance(size, ImageSize)
    assert (size.width, size.height) == (60, 24) and size.aspect == 2.5
    assert size.dpi == pytest.approx((144, 144), abs=0.1)
    _, plain = _picture(MARKER)
    assert plain.image_size == ImageSize(64, 32, None)
    box = deck.slides[0].add_textbox(0, 0, 914400, 914400, "x")
    assert box.image_size is None


def test_replace_image_keeps_the_frame_by_default():
    deck, picture = _picture()
    frame = (picture.left, picture.top, picture.width, picture.height)
    assert picture.replace_image(WIDE) is picture
    assert (picture.left, picture.top, picture.width, picture.height) == frame
    assert picture.image_size.width == 60


def test_keep_height_anchored_top_right_keeps_the_corner():
    deck, picture = _picture()
    right, top, height = picture.left + picture.width, picture.top, picture.height
    picture.replace_image(WIDE, keep="height", anchor="top_right")
    assert picture.height == height and picture.top == top
    assert picture.width == round(height * 2.5)
    assert picture.left + picture.width == right
    assert deck.validate() == []
    deck.undo()                                  # image and frame: one step
    again = deck.slides[0].shapes[0]
    assert (again.width, again.image_size.width) == (731520, 48)


def test_keep_width_and_none():
    deck, picture = _picture()
    picture.replace_image(WIDE, keep="width", anchor="center")
    assert picture.width == 731520 and picture.height == round(731520 / 2.5)
    assert picture.top == 300000 + (731520 - picture.height) // 2
    picture.replace_image(MARKER, keep="none")
    assert (picture.width, picture.height) == (64 * 9525, 32 * 9525)
    assert (picture.left, picture.top) == (9000000, picture.top)


def test_bad_arguments_are_refused():
    _, picture = _picture()
    with pytest.raises(ValueError):
        picture.replace_image(WIDE, keep="ratio")
    with pytest.raises(ValueError):
        picture.replace_image(WIDE, keep="height", anchor="middle")
