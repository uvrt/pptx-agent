"""Fills and outlines: solid, gradient, image and none; width, dash, arrowheads; theme colours."""

from __future__ import annotations

import io
import struct
import zipfile
import zlib

import pytest

from pptx_agent import Document
from pptx_agent.edit.color import Color
from pptx_agent.edit.fill import Arrowhead
from pptx_agent.oxml.xml import qn


def reopen(document: Document) -> Document:
    return Document.open(document.to_bytes())


def _autoshape(document: Document):
    return next(s for s in document.slides[0].shapes if s.kind == "shape")


def png(rgb=(200, 30, 30), size=8) -> bytes:
    """A tiny valid PNG, so the tests need no binary fixture."""
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))

    header = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()}


# -- shape fills ---------------------------------------------------------------------------


def test_solid_fill_hex(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.fill = "#00B050"
    fill = reopen(document).shape(shape.id).fill
    assert (fill.kind, fill.color) == ("solid", Color.rgb("00B050"))


def test_theme_fill_is_written_as_a_scheme_reference(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.fill = "accent1 lumMod=60% lumOff=40%"

    fill = reopen(document).shape(shape.id).fill
    assert fill.kind == "solid" and fill.color.is_theme
    assert fill.color == Color.theme("accent1", lum_mod=0.6, lum_off=0.4)
    xml = document.package.read(document.slides[0].part_path)
    assert b'<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/>' in xml


def test_gradient_fill(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.set_gradient_fill([(0, "accent1"), (0.5, "#FFFFFF"), (1, "accent1 lumMod=50%")], angle=45)

    fill = reopen(document).shape(shape.id).fill
    assert fill.kind == "gradient"
    assert fill.angle == pytest.approx(45)
    assert [(s.position, str(s.color)) for s in fill.stops] == [
        (0, "accent1"), (0.5, "#FFFFFF"), (1, "accent1 lumMod=50000")]


def test_radial_gradient_with_even_stops(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.set_gradient_fill(["accent2", "accent4", "accent6"], path="circle")
    fill = reopen(document).shape(shape.id).fill
    assert fill.path == "circle" and fill.angle is None
    assert [s.position for s in fill.stops] == [0, 0.5, 1]


def test_no_fill_and_inherit(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.fill = "none"
    assert reopen(document).shape(shape.id).fill.kind == "none"
    shape.fill = None
    assert reopen(document).shape(shape.id).fill is None


def test_a_new_fill_replaces_the_old_one(product_page):
    """EG_FillProperties is a choice: two fills in one spPr is a repair prompt."""
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.fill = "accent1"
    shape.set_gradient_fill(["accent1", "accent2"])
    shape.fill = "#123456"
    properties = shape._element.find(qn("p:spPr"))
    fills = [c for c in properties if c.tag.rpartition("}")[2].endswith("Fill")]
    assert len(fills) == 1


def test_fill_lands_between_geometry_and_outline(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.line.width = 12700
    shape.fill = "accent3"
    names = [c.tag.rpartition("}")[2] for c in shape._element.find(qn("p:spPr"))]
    assert names.index("solidFill") < names.index("ln")
    if "prstGeom" in names:
        assert names.index("prstGeom") < names.index("solidFill")


def test_image_fill_adds_the_media_once(product_page):
    document = Document.open(str(product_page))
    shapes = [s for s in document.slides[0].shapes if s.kind == "shape"][:2]
    picture = png()
    for shape in shapes:
        shape.set_image_fill(picture)

    reread = reopen(document)
    fills = [reread.shape(s.id).fill for s in shapes]
    assert all(f.kind == "image" for f in fills)
    assert fills[0].image_part == fills[1].image_part
    assert document.package.read(fills[0].image_part) == picture

    media = [n for n in _entries(document.to_bytes()) if n.startswith("ppt/media/")]
    assert sum(_entries(document.to_bytes())[n] == picture for n in media) == 1


def test_image_fill_undo_removes_the_media(product_page):
    document = Document.open(str(product_page))
    original = _entries(document.to_bytes())
    _autoshape(document).set_image_fill(png())
    assert set(_entries(document.to_bytes())) - set(original)

    assert document.undo()
    assert _entries(document.to_bytes()) == original
    assert document.redo()
    assert _autoshape(document).fill.kind == "image"


def test_image_fill_declares_a_content_type_when_the_deck_has_none(financial_report):
    """A deck with no PNGs gains a Default for png, or PowerPoint refuses the part."""
    from lxml import etree

    document = Document.open(str(financial_report))
    _autoshape(document).set_image_fill(png())
    types = etree.fromstring(_entries(document.to_bytes())["[Content_Types].xml"])
    defaults = {n.get("Extension").lower(): n.get("ContentType")
                for n in types if n.tag.endswith("Default")}
    assert defaults.get("png") == "image/png"


def test_unsupported_images_are_refused(product_page):
    document = Document.open(str(product_page))
    with pytest.raises(ValueError, match="unsupported image"):
        _autoshape(document).set_image_fill(b"not an image")
    assert not document.history.can_undo()


def test_graphic_frames_have_no_fill(financial_report):
    document = Document.open(str(financial_report))
    frame = next(s for slide in document.slides for s in slide.shapes if s.kind == "graphic_frame")
    with pytest.raises(ValueError):
        frame.fill = "accent1"


# -- outlines ------------------------------------------------------------------------------


def test_line_width_colour_dash(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    line = shape.line
    line.width = 38100
    line.color = "accent2"
    line.dash = "dashDot"
    line.cap = "rnd"

    reread = reopen(document).shape(shape.id).line
    assert (reread.width, reread.color, reread.dash, reread.cap) == (38100, "accent2", "dashDot", "rnd")


def test_arrowheads(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.line.head = "oval"
    shape.line.set_arrowhead("tail", "triangle", width="lg", length="sm")

    reread = reopen(document).shape(shape.id).line
    assert reread.head == Arrowhead("oval")
    assert reread.tail == Arrowhead("triangle", "lg", "sm")

    shape.line.tail = None
    assert reopen(document).shape(shape.id).line.tail is None


def test_line_children_stay_in_schema_order(product_page):
    document = Document.open(str(product_page))
    line = _autoshape(document).line
    line.tail = "arrow"
    line.head = "diamond"
    line.dash = "sysDot"
    line.color = "accent1"
    element = _autoshape(document)._element.find(qn("p:spPr")).find(qn("a:ln"))
    names = [c.tag.rpartition("}")[2] for c in element]
    assert names == ["solidFill", "prstDash", "headEnd", "tailEnd"]


def test_hidden_and_inherited_outline(product_page):
    document = Document.open(str(product_page))
    shape = _autoshape(document)
    shape.line.visible = False
    assert reopen(document).shape(shape.id).line.visible is False
    shape.line.clear()
    assert reopen(document).shape(shape.id).line.exists is False


def test_bad_line_values_are_refused(product_page):
    document = Document.open(str(product_page))
    line = _autoshape(document).line
    for attribute, value in (("dash", "wiggly"), ("head", "anchor"), ("width", -1)):
        with pytest.raises(ValueError):
            setattr(line, attribute, value)
    assert not document.history.can_undo()


def test_connector_arrowheads(pptx_path):
    document = Document.open(str(pptx_path))
    connector = next((s for slide in document.slides for s in slide.shapes
                      if s.kind == "connector"), None)
    if connector is None:
        pytest.skip("no connector in this fixture")
    connector.line.tail = "stealth"
    assert reopen(document).shape(connector.id).line.tail.type == "stealth"
