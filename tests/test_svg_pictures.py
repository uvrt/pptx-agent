"""SVG pictures (round 3, I1): a PNG in ``a:blip``, the SVG behind its ``asvg:svgBlip``
extension, as PowerPoint 2016 and later store an inserted SVG (measured in PowerPoint for
Mac 16: see :mod:`pptx_agent.edit.svgpicture`)."""

from __future__ import annotations

import importlib.util
import io
import json
import struct
import sys
import zipfile

import pytest
from lxml import etree

from images import ORANGE
from pptx_agent import Document
from pptx_agent.edit import svgpicture
from pptx_agent.fullstate.safe import FullStateError, UnsafeInput

#: A hand-written icon: a disc and a tick, 48 px square.
ICON = (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="48" height="48">'
        b'<circle cx="12" cy="12" r="10" fill="#1F5FAD"/>'
        b'<path d="M7 12.5l3.2 3.2L17 9" fill="none" stroke="#fff" stroke-width="2.2"/></svg>')
#: The same tick on a green disc: a recoloured icon.
GREEN = ICON.replace(b"#1F5FAD", b"#2E8B57")
WIDE = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100"/></svg>'

EMU = 914400 // 96          # one pixel at 96 dpi
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _slide():
    deck = Document.new()
    return deck, deck.add_slide("Blank")


def _parts(deck):
    archive = zipfile.ZipFile(io.BytesIO(deck.to_bytes()))
    return {name: archive.read(name) for name in archive.namelist()}


def _png_size(data: bytes) -> tuple[int, int]:
    return struct.unpack(">II", data[16:24])


def test_an_svg_picture_is_written_as_powerpoint_writes_one():
    deck, slide = _slide()
    picture = slide.add_picture(ICON, 914400, 914400, description="Done")
    assert (picture.width, picture.height) == (48 * EMU, 48 * EMU)    # its own size, 96 dpi
    assert picture.vector == ICON and picture.kind == "picture"
    assert picture.name == f"Graphic {int(picture.id.split('.')[1]) - 1}"
    assert deck.validate() == []
    parts = _parts(deck)
    root = etree.fromstring(parts["ppt/slides/slide1.xml"])
    pic = root.find(f".//{{{P}}}pic")
    assert pic.find(f".//{{{A}}}picLocks").attrib == {}          # aspect unlocked, as PowerPoint
    blip = pic.find(f".//{{{A}}}blip")
    ext = blip.find(f"{{{A}}}extLst/{{{A}}}ext")
    assert ext.get("uri") == "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
    svg_blip = ext[0]
    assert svg_blip.tag == f"{{{svgpicture.ASVG_NS}}}svgBlip"
    assert b'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"' \
        in parts["ppt/slides/slide1.xml"]
    rels = {r.get("Id"): r for r in etree.fromstring(parts["ppt/slides/_rels/slide1.xml.rels"])}
    png_rel, svg_rel = rels[blip.get(f"{{{R}}}embed")], rels[svg_blip.get(f"{{{R}}}embed")]
    assert png_rel.get("Type") == svg_rel.get("Type") == f"{R}/image"   # as PowerPoint writes
    assert svg_rel.get("Target").endswith(".svg") and png_rel.get("Target").endswith(".png")
    svg_part = "ppt/" + svg_rel.get("Target").removeprefix("../")
    assert parts[svg_part] == ICON                                     # the bytes given
    assert b'Extension="svg" ContentType="image/svg+xml"' in parts["[Content_Types].xml"]
    # The fallback: the SVG drawn at its placed size, 0.5 in, at 300 dpi.
    fallback = parts["ppt/" + png_rel.get("Target").removeprefix("../")]
    assert _png_size(fallback) == (150, 150)


def test_native_size_from_width_height_or_the_viewbox():
    deck, slide = _slide()
    assert slide.add_picture(WIDE).bounds[2:] == (200 * EMU, 100 * EMU)
    points = WIDE.replace(b"<svg ", b'<svg width="72pt" height="36pt" ')
    assert slide.add_picture(points).bounds[2:] == (96 * EMU, 48 * EMU)
    one_side = WIDE.replace(b"<svg ", b'<svg width="1in" ')
    assert slide.add_picture(one_side).bounds[2:] == (96 * EMU, 48 * EMU)
    assert slide.add_picture(WIDE, width=914400).bounds[2:] == (914400, 457200)
    sizeless = b'<svg xmlns="http://www.w3.org/2000/svg"><rect width="5" height="5"/></svg>'
    with pytest.raises(ValueError, match="states no size"):
        slide.add_picture(sizeless)
    assert slide.add_picture(sizeless, width=914400, height=914400).vector == sizeless


def _fallback_pixel(picture_svg: bytes, at: tuple[float, float]) -> tuple[int, int, int]:
    """The fallback PNG's pixel at a fraction of its size, for a 1 x 1 inch picture."""
    image_module = pytest.importorskip("PIL.Image")
    deck, slide = _slide()
    picture = slide.add_picture(picture_svg, 0, 0, width=914400, height=914400, fallback_dpi=96)
    parts = _parts(deck)
    rels = etree.fromstring(parts["ppt/slides/_rels/slide1.xml.rels"])
    target = next(r.get("Target") for r in rels if r.get("Target").endswith(".png"))
    image = image_module.open(io.BytesIO(parts["ppt/" + target.removeprefix("../")]))
    assert picture.vector == picture_svg
    image = image.convert("RGB")
    return image.getpixel((int(at[0] * image.width), int(at[1] * image.height)))


@pytest.mark.parametrize("size", ['width="72pt" height="72pt"', 'width="2.54cm" height="1in"',
                                  'width="6pc" height="25.4mm"'])
def test_the_fallback_draws_an_svg_sized_in_any_unit(size):
    """resvg-py draws nothing for a root of width="72pt"; the root reaches it in pixels."""
    if importlib.util.find_spec("pptx2svg") is None:
        pytest.skip("pptx2svg is not installed")
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" {size} viewBox="0 0 2 2">'
           '<rect width="1" height="2" fill="#f00"/><rect x="1" width="1" height="2" fill="#00f"/>'
           "</svg>").encode()
    assert _fallback_pixel(svg, (0.25, 0.5)) == (255, 0, 0)
    assert _fallback_pixel(svg, (0.75, 0.5)) == (0, 0, 255)


def test_the_fallback_of_an_svg_with_no_viewbox_is_drawn_at_powerpoints_size():
    """No viewBox, ``6em``: PowerPoint's em is 12 px, so the viewport is 72 px and a 36 px
    square fills half of it, stretched to the picture (measured, PowerPoint 16)."""
    if importlib.util.find_spec("pptx2svg") is None:
        pytest.skip("pptx2svg is not installed")
    svg = (b'<svg xmlns="http://www.w3.org/2000/svg" width="6em" height="6em">'
           b'<rect width="1000" height="1000" fill="#f00"/><rect width="36" height="36" fill="#00f"/>'
           b"</svg>")
    assert _fallback_pixel(svg, (0.4, 0.4)) == (0, 0, 255)
    assert _fallback_pixel(svg, (0.6, 0.6)) == (255, 0, 0)


def test_identical_svgs_are_stored_once_and_a_given_fallback_is_used_as_is():
    deck, slide = _slide()
    first = slide.add_picture(ICON, 0, 0)
    second = slide.add_picture(ICON, 914400, 0, width=1828800, fallback=ORANGE)
    assert first.vector == second.vector == ICON
    parts = _parts(deck)
    assert sum(name.endswith(".svg") for name in parts) == 1
    assert ORANGE in parts.values()
    with pytest.raises(ValueError, match="not an SVG"):
        slide.add_picture(ICON, fallback=ICON)
    with pytest.raises(ValueError, match="fallback is for an SVG picture"):
        slide.add_picture(ORANGE, fallback=ORANGE)


def test_the_fallback_follows_the_dpi_and_a_path_works(tmp_path):
    deck, slide = _slide()
    path = tmp_path / "icon.svg"
    path.write_bytes(ICON)
    picture = slide.add_picture(path, 0, 0, width=914400, height=914400, fallback_dpi=96)
    assert (picture.image_size.width, picture.image_size.height) == (96, 96)
    assert picture.vector == ICON


@pytest.mark.parametrize("svg, error, message", [
    (b'<!DOCTYPE svg [<!ENTITY x "boom">]><svg xmlns="http://www.w3.org/2000/svg" '
     b'width="9" height="9">&x;</svg>', UnsafeInput, "document type"),
    (b'<svg xmlns="http://www.w3.org/2000/svg" width="9" height="9"><text>&x;</text></svg>',
     FullStateError, "not well-formed"),
    (b'<svg width="9" height="9"><rect/></svg>', FullStateError, "not an SVG document"),
    (b'<svg xmlns="http://www.w3.org/2000/svg" width="9" height="9"', FullStateError,
     "not well-formed"),
])
def test_the_svg_is_parsed_as_untrusted_data(svg, error, message):
    deck, slide = _slide()
    with pytest.raises(error, match=message):
        slide.add_picture(svg, 0, 0, fallback=ORANGE)
    assert slide.shapes == [] and not any(n.endswith(".svg") for n in _parts(deck))


def test_an_oversized_svg_is_refused(monkeypatch):
    monkeypatch.setattr(svgpicture, "SVG_LIMIT", 100)
    deck, slide = _slide()
    with pytest.raises(UnsafeInput, match="the limit is 100"):
        slide.add_picture(ICON, 0, 0, fallback=ORANGE)


def test_without_the_png_extra_the_error_names_it(monkeypatch):
    monkeypatch.setitem(sys.modules, "pptx2svg.png", None)
    deck, slide = _slide()
    with pytest.raises(ImportError, match=r"pptx-agent\[png\]"):
        slide.add_picture(ICON, 0, 0)
    assert slide.add_picture(ICON, 0, 0, fallback=ORANGE).vector == ICON


def test_set_vector_attaches_replaces_and_removes_in_one_undo_step_each():
    deck, slide = _slide()
    picture = slide.add_picture(ORANGE, 0, 0)
    assert picture.vector is None
    picture.set_vector(ICON)
    assert deck.resolve(picture.id).vector == ICON
    picture = deck.resolve(picture.id)
    picture.set_vector(GREEN)
    assert deck.resolve(picture.id).vector == GREEN
    assert sum(n.endswith(".svg") for n in _parts(deck)) == 1         # the blue one is reaped
    deck.undo()
    assert deck.resolve(picture.id).vector == ICON
    deck.resolve(picture.id).set_vector(None)
    assert deck.resolve(picture.id).vector is None
    assert not any(n.endswith(".svg") for n in _parts(deck)) and deck.validate() == []
    deck.undo()
    assert deck.resolve(picture.id).vector == ICON
    with pytest.raises(ValueError, match="not an SVG"):
        deck.resolve(picture.id).set_vector(ORANGE)
    box = deck.slides[0].add_textbox(0, 0, 914400, 914400, "x")
    with pytest.raises(ValueError, match="no image"):
        box.set_vector(ICON)


def test_replace_image_keeps_both_images_in_step():
    deck, slide = _slide()
    picture = slide.add_picture(ICON, 0, 0)
    picture.replace_image(ORANGE, svg=GREEN)                     # a recoloured icon
    assert picture.vector == GREEN and picture.image_size.width == 48
    picture.replace_image(WIDE, keep="height")                   # an SVG, its fallback made
    assert picture.vector == WIDE and picture.width == 2 * picture.height
    assert (picture.image_size.width, picture.image_size.height) == (300, 150)
    picture.replace_image(ORANGE)                                # a bitmap: no SVG left
    assert picture.vector is None
    assert not any(n.endswith(".svg") for n in _parts(deck)) and deck.validate() == []
    with pytest.raises(ValueError, match="SVG already"):
        picture.replace_image(ICON, svg=ICON)


def test_copies_keep_the_vector_and_deleting_reaps_both_parts():
    deck, slide = _slide()
    picture = slide.add_picture(ICON, 0, 0)
    copy = picture.duplicate(914400, 0)
    assert copy.vector == ICON
    other = Document.new()
    pasted = slide.copy_shapes([picture], other.add_slide("Blank")).shapes
    assert pasted[0].vector == ICON and other.validate() == []
    copy.delete()
    deck.resolve(picture.id).delete()
    assert not any("media" in n for n in _parts(deck)) and deck.validate() == []


def test_the_outline_marks_an_svg_picture_and_reads_the_mark_back():
    from pptx_agent.outline.draft import _hint

    deck, slide = _slide()
    picture = slide.add_picture(ICON, 0, 0, description="Done")
    slide.add_picture(ORANGE, 914400, 0)
    outline = deck.to_outline()
    assert f"<!-- {picture.id} picture svg -->" in outline
    assert outline.count("picture svg") == 1
    assert _hint(f"{picture.id} picture svg").kind == "picture"


def test_pptx2svg_draws_the_vector():
    deck, slide = _slide()
    slide.add_picture(ICON, 0, 0)
    assert "image/svg+xml" in slide.render_svg()


# -- the tools -----------------------------------------------------------------------------------


def test_ppt_add_picture_takes_an_svg_blob_to_insert_and_to_replace():
    from ooxml_edit.tools import Toolbox
    from pptx_agent.tools import FORMAT, GROUPS, TOOLS

    deck = Document.new()
    deck.add_slide("Title Only")
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session()
        session.open(deck.to_bytes(), name="icons.pptx")
        icon = session.add_blob(ICON, "icon.svg", "image/svg+xml")
        made = toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "image": icon, "slide": "s:256",
            "box": {"x": 72, "y": 72, "w": 36, "h": 36}, "ref": "icon"})
        assert made.ok, made.to_json()
        assert made.data["svg"] is True and made.data["native_px"] == [48, 48]
        green = session.add_blob(GREEN, "green.svg", "image/svg+xml")
        swapped = toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "image": green, "target": "$icon"})
        assert swapped.ok, swapped.to_json()
        shape = session.entry("d1").document.shape(made.data["address"])
        assert shape.vector == GREEN and shape.bounds[2] == 36 * 12700
        geometry = toolbox.dispatch(session, "ppt_read_slides", {"doc": "d1",
                                                                 "detail": "geometry"})
        entry = next(s for s in geometry.data if s["id"] == shape.id)
        assert entry["image"]["svg"] is True
        bad = session.add_blob(b'<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>',
                               "bad.svg", "image/svg+xml")
        refused = toolbox.dispatch(session, "ppt_add_picture", {"doc": "d1", "image": bad,
                                                                "slide": "s:256"})
        assert not refused.ok and "document type" in json.dumps(refused.to_json())
