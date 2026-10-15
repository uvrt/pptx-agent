"""Round 5, T20: an SVG picture recoloured as PowerPoint's Graphics Fill does it, and an
application's asset catalogue (a folder and an index) searched and placed by id.

What PowerPoint writes was measured in PowerPoint for Mac 16 (2026-10-15) on a picture of
four paths -- two filled black, one with no fill, one ``currentColor`` -- recoloured accent 2
and FF6600: the ``p:pic`` stays, the SVG part is rewritten in place (a class rule and
``class``, ``fill`` and ``fill-opacity="1"`` on every path for the theme colour; ``fill``
and ``fill-opacity`` alone for the RGB one) and the PNG fallback is drawn again in the new
colour.  The icons here are hand-written, the keywords made up."""

from __future__ import annotations

import io
import json
import struct
import zipfile
import zlib

import pytest
from lxml import etree

from images import ORANGE
from ooxml_edit.tools import FolderCatalogue, Toolbox
from pptx_agent import Document
from pptx_agent.edit import svgpicture
from pptx_agent.tools import FORMAT, GROUPS, TOOLS, deck_format

pytest.importorskip("resvg_py", reason="the PNG fallback is drawn by pptx2svg[png]")

SVG = "http://www.w3.org/2000/svg"
#: Four squares: black, black, no fill (black by default), currentColor.
QUAD = (b'<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24">\n'
        b'<path fill="#000000" d="M2 2h9v9H2z"/>\n'
        b'<path fill="#000000" d="M13 13h9v9h-9z"/>\n'
        b'<path d="M13 2h9v9h-9z"/>\n'
        b'<path fill="currentColor" d="M2 13h9v9H2z"/>\n</svg>')
#: An outline icon: nothing to fill.
OUTLINE = (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24" '
           b'fill="none" stroke="#000"><path d="M4 12h16"/><circle cx="12" cy="12" r="9"/></svg>')
#: A disc with an outlined tick on it, styled by a class rule, clipped by a clip path.
STYLED = (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">'
          b'<style>.ink{fill:#123456}</style>'
          b'<clipPath id="c"><rect width="24" height="24" fill="#000"/></clipPath>'
          b'<circle class="ink" cx="12" cy="12" r="10" style="fill:#654321;opacity:.9" '
          b'clip-path="url(#c)"/>'
          b'<path d="M7 12l3 3 7-7" fill="none" stroke="#fff"/></svg>')
PLACED = 1828800                    # two inches


def _picture(svg=QUAD, count=1):
    deck = Document.new()
    slide = deck.add_slide("Blank")
    pictures = [slide.add_picture(svg, 914400 + n * 2743200, 914400, width=PLACED)
                for n in range(count)]
    return deck, slide, pictures


def _paths(svg: bytes):
    root = etree.fromstring(svg)
    return root, root.findall(f"{{{SVG}}}path")


def _media(deck):
    return sorted(name for name in deck.package.part_names if name.startswith("ppt/media/"))


def _pixel(png: bytes, fx: float, fy: float) -> tuple[int, ...]:
    """The pixel at a fraction of an 8-bit RGB or RGBA PNG's width and height."""
    width, height, depth, colour = struct.unpack(">IIBB", png[16:26])
    assert depth == 8 and colour in (2, 6)
    channels = 4 if colour == 6 else 3
    data, at = b"", 8
    while at < len(png):
        length, kind = struct.unpack(">I4s", png[at:at + 8])
        if kind == b"IDAT":
            data += png[at + 8:at + 8 + length]
        at += 12 + length
    raw = zlib.decompress(data)
    stride = width * channels
    rows, previous, offset = [], bytearray(stride), 0
    for _ in range(height):
        kind, line = raw[offset], bytearray(raw[offset + 1:offset + 1 + stride])
        offset += 1 + stride
        for i in range(stride):
            left = line[i - channels] if i >= channels else 0
            up, corner = previous[i], previous[i - channels] if i >= channels else 0
            if kind == 1:
                line[i] = (line[i] + left) & 255
            elif kind == 2:
                line[i] = (line[i] + up) & 255
            elif kind == 3:
                line[i] = (line[i] + (left + up) // 2) & 255
            elif kind == 4:
                p = left + up - corner
                pa, pb, pc = abs(p - left), abs(p - up), abs(p - corner)
                line[i] = (line[i] + (left if pa <= pb and pa <= pc else up if pb <= pc
                                      else corner)) & 255
        rows.append(line)
        previous = line
    x, y = int(width * fx), int(height * fy)
    return tuple(rows[y][x * channels:(x + 1) * channels])


def _fallback(deck, picture) -> bytes:
    return deck.package.read(picture.image_part)


# -- the library ------------------------------------------------------------------------------


def test_accent2_is_written_as_powerpoint_writes_it_in_place():
    deck, slide, (picture,) = _picture()
    blip = etree.tostring(picture._element.find(".//{*}blipFill"))
    media = _media(deck)
    picture.recolor("Accent 2")                       # Office's name for accent2
    assert _media(deck) == media                      # rewritten in place, as PowerPoint does
    assert etree.tostring(picture._element.find(".//{*}blipFill")) == blip
    root, paths = _paths(picture.vector)
    style = root[0]
    assert style.tag == f"{{{SVG}}}style"
    assert style.text == "\n.MsftOfcThm_Accent2_Fill_v2 {\n fill:#E97132; \n}\n"
    assert len(paths) == 4
    for path in paths:                                # no-fill and currentColor ones too
        assert path.get("class") == "MsftOfcThm_Accent2_Fill_v2"
        assert path.get("fill") == "#E97132" and path.get("fill-opacity") == "1"
        assert path.get("stroke") is None            # strokes are left as they were
    assert [p.get("d") for p in paths][2] == "M13 2h9v9h-9z"   # path data not normalised
    png = _fallback(deck, picture)
    for corner in ((0.25, 0.25), (0.75, 0.25), (0.25, 0.75), (0.75, 0.75)):
        assert _pixel(png, *corner)[:3] == (233, 113, 50)
    assert deck.validate() == []


def test_rgb_gets_fill_and_opacity_alone():
    deck, slide, (picture,) = _picture()
    picture.recolor("#FF6600")
    root, paths = _paths(picture.vector)
    assert root.find(f"{{{SVG}}}style") is None
    assert all(p.get("class") is None and p.get("fill") == "#FF6600"
               and p.get("fill-opacity") == "1" for p in paths)
    assert _pixel(_fallback(deck, picture), 0.25, 0.75)[:3] == (255, 102, 0)


def test_another_theme_colour_is_its_value_until_its_class_is_measured():
    deck, slide, (picture,) = _picture()
    picture.recolor("accent1")
    root, paths = _paths(picture.vector)
    assert root.find(f"{{{SVG}}}style") is None and paths[0].get("class") is None
    value = paths[0].get("fill")
    from pptx_agent.edit.color import Color

    assert value == Color.parse("accent1").resolve(picture)
    picture.recolor("text1")                           # through the slide's colour map
    assert _paths(picture.vector)[1][0].get("fill") == Color.parse("tx1").resolve(picture)
    picture.recolor("accent2 lumMod=75%")              # a shade: its value, no class
    assert _paths(picture.vector)[1][0].get("class") is None


def test_recolouring_again_drops_the_earlier_class_and_rule():
    deck, slide, (picture,) = _picture()
    picture.recolor("accent2")
    picture.recolor("#336699")
    root, paths = _paths(picture.vector)
    assert root.find(f"{{{SVG}}}style") is None
    assert all(p.get("class") is None and p.get("fill") == "#336699" for p in paths)
    picture.recolor("accent2")
    root, paths = _paths(picture.vector)
    assert len(root.findall(f"{{{SVG}}}style")) == 1


def test_a_picture_sharing_its_images_gets_its_own_and_the_other_keeps_them():
    deck, slide, (first, second) = _picture(count=2)
    assert first.image_part == second.image_part      # stored once
    first.recolor("accent2")
    assert second.vector == QUAD
    assert first.image_part != second.image_part
    assert deck.validate() == []
    archive = zipfile.ZipFile(io.BytesIO(deck.to_bytes()))
    assert sum(name.endswith(".svg") for name in archive.namelist()) == 2


def test_undo_takes_a_recolour_back():
    deck, slide, (picture,) = _picture()
    before = deck.to_bytes()
    picture.recolor("accent2")
    deck.undo()
    assert deck.to_bytes() == before


def test_what_cannot_be_recoloured_is_refused():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    bitmap = slide.add_picture(ORANGE, 0, 0)
    with pytest.raises(ValueError, match="bitmap picture"):
        bitmap.recolor("accent2")
    outline = slide.add_picture(OUTLINE, 0, 0)
    with pytest.raises(ValueError, match="outlines only"):
        outline.recolor("accent2")
    _, _, (icon,) = _picture()
    with pytest.raises(ValueError, match="without alpha"):
        icon.recolor("accent2 alpha=50%")
    with pytest.raises(ValueError, match="neither a hex colour nor a theme colour"):
        icon.recolor("teal-ish")
    box = slide.add_shape("rect", 0, 0, 100, 100)
    with pytest.raises(ValueError, match="no image"):
        box.recolor("accent2")


def test_outlines_and_clip_paths_keep_their_colours_and_class_rules_are_overridden():
    deck, slide, (picture,) = _picture(STYLED)
    with pytest.warns(svgpicture.RecolorWarning, match="1 shape"):
        picture.recolor("#FF6600")
    root = etree.fromstring(picture.vector)
    circle = root.find(f"{{{SVG}}}circle")
    # The SVG's own class rule would win over the attribute: the fill goes in its style too.
    assert circle.get("fill") == "#FF6600" and circle.get("class") == "ink"
    assert circle.get("style") == "opacity:.9;fill:#FF6600;fill-opacity:1"
    tick = root.find(f"{{{SVG}}}path")
    assert tick.get("fill") == "none" and tick.get("stroke") == "#fff"
    assert root.find(f"{{{SVG}}}clipPath/{{{SVG}}}rect").get("fill") == "#000"


# -- the tools --------------------------------------------------------------------------------


ICONS = {
    "check-circle": (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" '
                     b'height="24"><circle cx="12" cy="12" r="10"/></svg>',
                     "Check in a circle", ["done", "complete", "tick"]),
    "flag": (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" '
             b'height="24"><path d="M5 3h12l-3 4 3 4H7v10H5z"/></svg>',
             "Flag", ["milestone", "goal"]),
    "stack": (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 48" width="24" '
              b'height="48"><rect x="2" y="2" width="20" height="44"/></svg>',
              "Stacked blocks", ["layers", "complete set"]),
}


@pytest.fixture
def catalogue(tmp_path):
    index = []
    for asset_id, (svg, name, keywords) in ICONS.items():
        (tmp_path / f"{asset_id}.svg").write_bytes(svg)
        index.append({"id": asset_id, "file": f"{asset_id}.svg", "name": name,
                      "keywords": keywords})
    (tmp_path / "index.json").write_text(json.dumps(index), encoding="utf-8")
    return FolderCatalogue(tmp_path)


def _session(toolbox):
    session = toolbox.session()
    deck = Document.new()
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")
    return session, f"s:{deck.slides[0].slide_id}"


def _ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def test_set_shape_fill_recolours_an_svg_picture_and_fills_behind_a_bitmap():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session, slide = _session(toolbox)
        svg = session.add_blob(QUAD, "quad.svg")
        png = session.add_blob(ORANGE, "orange.png")
        icon = _ok(toolbox.dispatch(session, "ppt_add_picture",
                                    {"doc": "d1", "image": svg, "slide": slide}))
        photo = _ok(toolbox.dispatch(session, "ppt_add_picture",
                                     {"doc": "d1", "image": png, "slide": slide}))
        _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": icon.data["address"], "fill": "accent2"}]}))
        deck = session.documents["d1"].document
        assert b"MsftOfcThm_Accent2_Fill_v2" in deck.shape(icon.data["address"]).vector
        photo_png = deck.package.read(deck.shape(photo.data["address"]).image_part)
        _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": photo.data["address"], "fill": "accent2"}]}))     # its background
        photo_shape = deck.shape(photo.data["address"])
        assert str(photo_shape.fill.color) == "accent2"
        assert deck.package.read(photo_shape.image_part) == photo_png    # image untouched
        _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": photo.data["address"], "fill": "none"}]}))       # as before
        outline = session.add_blob(STYLED, "styled.svg")
        made = _ok(toolbox.dispatch(session, "ppt_add_picture",
                                    {"doc": "d1", "image": outline, "slide": slide}))
        warned = _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": made.data["address"], "fill": "#FF6600"}]}))
        assert any("keep their colours" in str(w) for w in warned.warnings)


def test_without_a_catalogue_nothing_is_offered_for_one():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        assert "ppt_find_asset" not in toolbox.tools
        assert "asset" not in toolbox.tools["ppt_add_picture"].schema["properties"]
        assert toolbox.tools["ppt_add_picture"].schema["required"] == ["doc", "image"]
    with Toolbox(TOOLS, formats=[deck_format(min_text_size=10)], groups=GROUPS,
                 workers=0) as toolbox:
        assert "ppt_find_asset" not in toolbox.tools
    with pytest.raises(ValueError, match="search"):
        deck_format(assets=object())


def test_a_catalogue_is_searched_and_its_assets_placed_by_id(catalogue):
    with Toolbox(TOOLS, formats=[deck_format(assets=catalogue)], groups=GROUPS,
                 workers=0) as toolbox:
        finder = toolbox.tools["ppt_find_asset"]
        assert finder.group == "ppt_objects" and not finder.mutates
        assert "asset" in toolbox.tools["ppt_add_picture"].schema["properties"]
        session, slide = _session(toolbox)
        found = _ok(toolbox.dispatch(session, "ppt_find_asset", {"query": "complete"}))
        assert [a["id"] for a in found.data["assets"]] == ["check-circle", "stack"]
        assert found.data["assets"][0] == {"id": "check-circle", "name": "Check in a circle",
                                           "keywords": ["done", "complete", "tick"]}
        assert _ok(toolbox.dispatch(session, "ppt_find_asset",
                                    {"query": "complete", "limit": 1})).data["assets"][0][
            "id"] == "check-circle"
        none = _ok(toolbox.dispatch(session, "ppt_find_asset", {"query": "rocket"}))
        assert none.data == {"assets": []} and "other words" in none.summary
        placed = _ok(toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "asset": "flag", "slide": slide,
            "box": {"x": 72, "y": 72, "w": 72, "h": 72}}))
        assert placed.data["svg"] is True
        deck = session.documents["d1"].document
        address = placed.data["address"]
        assert deck.shape(address).vector == ICONS["flag"][0]
        swapped = _ok(toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "asset": "stack", "target": address, "keep": "height"}))
        assert swapped.changed == [address]
        assert deck.shape(address).vector == ICONS["stack"][0]
        assert swapped.data["box"]["w"] == 36             # half as wide as high, height kept
        missing = toolbox.dispatch(session, "ppt_add_picture",
                                   {"doc": "d1", "asset": "rocket", "slide": slide})
        assert missing.error.code == "not_found" and "ppt_find_asset" in missing.error.message
        both = toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "asset": "flag", "image": "b1", "slide": slide})
        assert both.error.code == "invalid_arguments"
        neither = toolbox.dispatch(session, "ppt_add_picture", {"doc": "d1", "slide": slide})
        assert neither.error.code == "invalid_arguments"
        blob = session.add_blob(ORANGE, "orange.png")
        _ok(toolbox.dispatch(session, "ppt_add_picture",
                             {"doc": "d1", "image": blob, "slide": slide}))


def test_a_catalogue_costs_its_tool_and_one_parameter_and_only_with_it(catalogue):
    def estimate(toolbox):
        return sum(len(json.dumps(d, separators=(",", ":"))) for d in
                   toolbox.definitions("anthropic", defer=False)) / 3.5

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as plain, \
            Toolbox(TOOLS, formats=[deck_format(assets=catalogue)], groups=GROUPS,
                    workers=0) as configured:
        added = estimate(configured) - estimate(plain)
        assert 0 < added <= 150, added
        core = [t["name"] for t in configured.definitions("anthropic", groups="core")]
        assert core == [t["name"] for t in plain.definitions("anthropic", groups="core")]


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
def test_with_a_catalogue_the_definitions_meet_each_providers_rules(catalogue, provider):
    from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

    with Toolbox(TOOLS, formats=[deck_format(assets=catalogue)], groups=GROUPS,
                 workers=0) as toolbox:
        for groups in ("core", None, ["ppt_objects"]):
            definitions = toolbox.definitions(provider, groups=groups)
            problems = (anthropic_problems(definitions) if provider == "anthropic" else
                        openai_problems(definitions, chat=provider == "openai-chat"))
            assert problems == []
