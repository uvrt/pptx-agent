"""Round 5, T20b: an SVG picture recoloured as PowerPoint writes it, measured with PowerPoint
for Mac 16 on 2026-10-15 (oracle2/results.md, kept outside the repository with the other
oracle runs; cited by section as "§n").  The SVGs here are made up; the patterns checked --
class names, attribute values and their order -- are PowerPoint's."""

from __future__ import annotations

import json
import re

import pytest
from lxml import etree

from images import ORANGE
from ooxml_edit.tools import Toolbox
from pptx_agent import Document
from pptx_agent.edit import svgpicture
from pptx_agent.edit.color import Color
from pptx_agent.oxml.xml import qn
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from test_recolor_and_assets import _pixel

pytest.importorskip("resvg_py", reason="the PNG fallback is drawn by pptx2svg[png]")

SVG = "http://www.w3.org/2000/svg"
INCH = 914400
#: A triangle and a bar.
SOLID = (b'<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 40 40">'
         b'<path fill="#202020" d="M4 36L20 4l16 32z"/>'
         b'<rect fill="#5A5A5A" x="4" y="30" width="32" height="6"/></svg>')
#: Outlines: an open frame, a capped line, a filled square with a stroke, a red ring.
LINES = (b'<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48">'
         b'<rect fill="none" stroke="#111111" stroke-width="3" x="4" y="4" width="16" '
         b'height="16"/>'
         b'<path fill="none" stroke="#111111" stroke-width="3" stroke-linecap="square" '
         b'd="M28 12h16"/>'
         b'<rect fill="#2266AA" stroke="#111111" stroke-width="3" x="28" y="28" width="16" '
         b'height="16"/>'
         b'<circle fill="none" stroke="#AA2222" stroke-width="3" cx="12" cy="36" r="8"/></svg>')
#: Styled by a style sheet (a class, a descendant rule, a group's class) and style
#: attributes.
STYLED = (b'<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 40 40">'
          b'<style>/* made up */ .q1{fill:#0A0A0A} .q2{fill:#3C3C3C;stroke:#123123;'
          b'stroke-width:2} g.ring circle{stroke:#456456} .ring{stroke-width:1.5}</style>'
          b'<path class="q1" d="M2 2h16v16H2z"/>'
          b'<path class="q2" d="M22 2h16v16H22z"/>'
          b'<path style="fill:#AB0000" d="M2 22h16v16H2z"/>'
          b'<g class="ring"><circle style="fill:#00AB00;fill-opacity:0.4" cx="30" cy="30" '
          b'r="7"/></g></svg>')
#: A gradient-filled bar, and a symbol shown twice by ``use`` (``xlink:href`` and ``href``).
REUSED = (b'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
          b'width="40" height="40" viewBox="0 0 40 40"><defs>'
          b'<linearGradient id="fade"><stop offset="0" stop-color="#FFFF00"/>'
          b'<stop offset="1" stop-color="#00FFFF"/></linearGradient>'
          b'<symbol id="dot" viewBox="0 0 8 8"><circle cx="4" cy="4" r="4" fill="#777777"/>'
          b'</symbol></defs>'
          b'<rect fill="url(#fade)" x="2" y="2" width="36" height="14"/>'
          b'<use xlink:href="#dot" x="4" y="22" width="12" height="12"/>'
          b'<use href="#dot" x="24" y="22" width="12" height="12"/></svg>')

#: §1 (PowerPoint's default theme, which Document.new() has): each slot's class and value.
SLOTS = [
    ("dk1", "MainDark1", "#000000"), ("Dark 1", "MainDark1", "#000000"),
    ("lt1", "MainLight1", "#FFFFFF"), ("dk2", "MainDark2", "#0E2841"),
    ("lt2", "MainLight2", "#E8E8E8"), ("accent1", "Accent1", "#156082"),
    ("accent2", "Accent2", "#E97132"), ("Accent 3", "Accent3", "#196B24"),
    ("accent4", "Accent4", "#0F9ED5"), ("accent5", "Accent5", "#A02B93"),
    ("accent6", "Accent6", "#4EA72E"), ("hlink", "Hyperlink", "#467886"),
    ("Hyperlink", "Hyperlink", "#467886"), ("folHlink", "FollowedHyperlink", "#96607D"),
    ("followedHyperlink", "FollowedHyperlink", "#96607D"), ("tx1", "Text1", "#000000"),
    ("text1", "Text1", "#000000"), ("bg1", "Background1", "#FFFFFF"),
    ("background1", "Background1", "#FFFFFF"), ("tx2", "Text2", "#0E2841"),
    ("Text 2", "Text2", "#0E2841"), ("bg2", "Background2", "#E8E8E8"),
    ("background2", "Background2", "#E8E8E8"),
]
#: §1: what every element gets, between its class and its fill, when it had no stroke.
NO_STROKE = ('stroke="none" stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" '
             'stroke-miterlimit="4"')


def _picture(svg, size=INCH, deck=None, slide=None):
    deck = deck or Document.new()
    slide = slide or deck.add_slide("Blank")
    return deck, slide, slide.add_picture(svg, INCH, INCH, width=size)


def _rule(name, prop, value):
    return f".{name} {{\n {prop}:{value}; \n}}\n"


def _elements(svg: bytes, tag: str):
    return etree.fromstring(svg).iter(f"{{{SVG}}}{tag}")


# -- §1: every theme slot ---------------------------------------------------------------------


@pytest.mark.parametrize("name, slot, value", SLOTS)
def test_every_theme_slot_is_written_with_powerpoints_class(name, slot, value):
    deck, slide, picture = _picture(SOLID)
    picture.recolor(name)
    svg = picture.vector.decode()
    css = f"MsftOfcThm_{slot}_Fill_v2"
    assert f"<style>\n{_rule(css, 'fill', value)}</style>\n<path " in svg
    assert svg.count(f'class="{css}" {NO_STROKE} fill="{value}" fill-opacity="1"') == 2
    assert Color.parse(name).resolve(picture) == value           # the slide's theme


def test_text_and_main_slots_are_different_classes_for_one_colour():
    deck, slide, picture = _picture(SOLID)
    picture.recolor("text1")
    text = picture.vector
    picture.recolor("dk1")
    assert b"MsftOfcThm_MainDark1_Fill_v2" in picture.vector
    assert text.replace(b"Text1", b"MainDark1") == picture.vector


def test_a_tint_or_shade_is_its_value_alone():
    deck, slide, picture = _picture(SOLID)
    picture.recolor("accent2 lumMod=75%")
    svg = picture.vector
    value = Color.parse("accent2 lumMod=75%").resolve(picture)
    assert b"<style" not in svg and b"class=" not in svg
    assert svg.count(f'{NO_STROKE} fill="{value}" fill-opacity="1"'.encode()) == 2


# -- §2: outlines, Graphics Outline ------------------------------------------------------------


def test_a_fill_fills_outline_shapes_and_keeps_their_strokes():
    deck, slide, picture = _picture(LINES)
    picture.recolor("accent4")
    svg = picture.vector.decode()
    tail = ('stroke-linejoin="miter" stroke-miterlimit="4" stroke-opacity="1" '
            'fill="#0F9ED5" fill-opacity="1"')
    css = 'class="MsftOfcThm_Accent4_Fill_v2"'
    assert svg.count(f'{css} stroke="#111111" stroke-width="3" {tail}') == 2
    assert f'{css} stroke="#111111" stroke-width="3" stroke-linecap="square" {tail}' in svg
    assert f'{css} stroke="#AA2222" stroke-width="3" {tail}' in svg
    assert 'fill="none"' not in svg


def test_an_outline_colour_strokes_every_shape_and_leaves_fills():
    deck, slide, picture = _picture(LINES)
    picture.recolor(line="accent6")
    svg = picture.vector.decode()
    css = "MsftOfcThm_Accent6_Stroke_v2"
    assert f"<style>\n{_rule(css, 'stroke', '#4EA72E')}</style>" in svg
    common = (f'class="{css}" stroke="#4EA72E" stroke-width="3" ')
    end = 'stroke-linejoin="miter" stroke-miterlimit="4" stroke-opacity="1"'
    assert svg.count(f'{common}{end} fill="none"/>') == 2
    assert f'{common}stroke-linecap="square" {end} fill="none"/>' in svg
    assert f'{common}{end} fill="#2266AA"/>' in svg
    assert "fill-opacity" not in svg
    picture.recolor(line="#FF6600")                   # RGB: no style, no class
    svg = picture.vector.decode()
    assert "<style" not in svg and "class=" not in svg
    assert svg.count('stroke="#FF6600" stroke-width="3"') == 4


@pytest.mark.parametrize("name, slot, value", SLOTS)
def test_every_theme_slot_has_powerpoints_stroke_class(name, slot, value):
    # PowerPoint for Mac, 2026-10-15, oracle2/results.md E1.
    deck, slide, picture = _picture(LINES)
    picture.recolor(line=name)
    svg = picture.vector.decode()
    css = f"MsftOfcThm_{slot}_Stroke_v2"
    assert f"<style>\n{_rule(css, 'stroke', value)}</style>" in svg
    assert svg.count(f'class="{css}" stroke="{value}" stroke-width="3"') == 4


def test_an_outline_colour_adds_no_stroke_to_shapes_without_one():
    # E2: the colour alone leaves an SVG of unstroked shapes as it was drawn: no stroke,
    # no class, no rule; each shape's fill-opacity stated.
    deck, slide, picture = _picture(SOLID)
    picture.recolor(line="accent2")
    svg = picture.vector.decode()
    assert "<style" not in svg and "class=" not in svg
    assert f'{NO_STROKE} fill="#202020" fill-opacity="1"' in svg
    assert f'{NO_STROKE} fill="#5A5A5A" fill-opacity="1"' in svg


def test_an_outline_weight_strokes_unstroked_shapes_in_text_1():
    # E2: a weight alone strokes every shape in PowerPoint's default outline, Text 1.
    deck, slide, picture = _picture(SOLID)
    picture.recolor(line_width=12700)       # 1 pt: 4/3 px, 40 units at 96 px -> 0.5556
    svg = picture.vector.decode()
    css = "MsftOfcThm_Text1_Stroke_v2"
    assert f"<style>\n{_rule(css, 'stroke', '#000000')}</style>" in svg
    pattern = (f'class="{css}" stroke="#000000" stroke-width="0.5556" stroke-linecap="butt" '
               'stroke-linejoin="miter" stroke-miterlimit="4" fill="{}" fill-opacity="1"')
    assert pattern.format("#202020") in svg and pattern.format("#5A5A5A") in svg
    # A colour with it: the weight adds the stroke, the colour sets it (the two measured
    # rules composed; not measured together).
    _, _, other = _picture(SOLID, deck=deck, slide=slide)
    other.recolor(line="accent5", line_width=12700)
    svg = other.vector.decode()
    assert "Text1" not in svg
    assert svg.count('class="MsftOfcThm_Accent5_Stroke_v2" stroke="#A02B93" '
                     'stroke-width="0.5556" stroke-linecap="butt"') == 2


def test_an_outline_transparency_is_stroke_opacity_and_drawn_once():
    # E3: stroke-opacity, the class kept for a theme colour, none for RGB.
    deck, slide, picture = _picture(LINES)
    picture.recolor(line="#FF6600 alpha=50%")
    svg = picture.vector.decode()
    assert "class=" not in svg
    assert svg.count('stroke="#FF6600" stroke-width="3"') == 4
    assert svg.count('stroke-miterlimit="4" stroke-opacity="0.5"') == 4
    # The open frame's left edge (x 2.5-5.5 of 48 units), outside any fill: alpha once.
    pixel = _pixel(deck.package.read(picture.image_part), 4 / 48, 12 / 48)
    assert all(abs(a - b) <= 1 for a, b in zip(pixel, (255, 102, 0, 128)))
    picture.recolor(line="accent2 alpha=50%")
    assert picture.vector.decode().count(
        'class="MsftOfcThm_Accent2_Stroke_v2" stroke="#E97132"') == 4


def test_an_outline_weight_is_converted_to_the_svgs_units_at_the_placed_size():
    # 48 units placed at 1 in (96 px): 2 px a unit; 1.5 pt is 2 px, so 1 unit (§2: 3 pt
    # over 24 units placed at 192 px is 0.5).
    deck, slide, picture = _picture(LINES)
    picture.recolor(line_width=19050)
    svg = picture.vector.decode()
    assert svg.count('stroke-width="1"') == 4 and "<style" not in svg
    assert 'stroke="#AA2222" stroke-width="1"' in svg            # colours kept
    picture.width = picture.height = 2 * INCH
    picture.recolor(line_width=19050)
    assert picture.vector.decode().count('stroke-width="0.5"') == 4


def test_fill_and_outline_share_one_style_and_both_classes():
    deck, slide, picture = _picture(LINES)
    picture.recolor("accent2", line="accent5")
    svg = picture.vector.decode()
    fill, stroke = "MsftOfcThm_Accent2_Fill_v2", "MsftOfcThm_Accent5_Stroke_v2"
    assert (f"<style>\n{_rule(fill, 'fill', '#E97132')}{_rule(stroke, 'stroke', '#A02B93')}"
            "</style>") in svg
    assert svg.count(f'class="{fill} {stroke}" stroke="#A02B93" stroke-width="3"') == 4
    assert svg.count('fill="#E97132" fill-opacity="1"') == 4
    combined = picture.vector
    _, _, other = _picture(LINES, deck=deck, slide=slide)
    other.recolor("accent2")
    other.recolor(line="accent5")                    # in two steps: the same SVG
    assert other.vector == combined
    other.recolor("#336699")                         # a new fill keeps the outline's class
    svg = other.vector.decode()
    assert fill not in svg and svg.count(f'class="{stroke}"') == 4


# -- §3: style sheets and style attributes ------------------------------------------------------


def test_a_style_sheet_is_replaced_and_what_it_set_carried_over():
    deck, slide, picture = _picture(STYLED)
    picture.recolor("accent2")
    svg = picture.vector.decode()
    root = etree.fromstring(picture.vector)
    assert len(root.findall(f"{{{SVG}}}style")) == 1 and "made up" not in svg
    assert "q1" not in svg and "q2" not in svg and "style=" not in svg
    tail = ('stroke-width="1" stroke-linecap="butt" stroke-linejoin="miter" '
            'stroke-miterlimit="4" fill="#E97132" fill-opacity="1"')
    css = 'class="MsftOfcThm_Accent2_Fill_v2"'
    paths = [p for p in re.findall(r"<path [^>]*>", svg)]
    assert f'{css} {tail}' in paths[0]              # no stroke: no stroke attribute
    assert f'{css} stroke="#123123" stroke-width="2" stroke-linecap="butt"' in paths[1]
    assert f'{css} {tail}' in paths[2]
    group = root.find(f"{{{SVG}}}g")
    assert group.get("class") == "ring" and group.get("stroke-width") == "1.5"
    circle = group[0]
    assert circle.get("stroke") == "#456456" and circle.get("stroke-width") == "1.5"
    assert circle.get("fill-opacity") == "1"          # the style attribute's 0.4 goes


def test_an_outline_alone_carries_the_style_sheets_fills_over():
    deck, slide, picture = _picture(STYLED)
    picture.recolor(line="#FF6600")
    paths = list(_elements(picture.vector, "path"))
    assert [p.get("fill") for p in paths] == ["#0A0A0A", "#3C3C3C", "#AB0000"]
    # Only the shape the style sheet stroked takes the colour (E2).
    assert [p.get("stroke") for p in paths] == [None, "#FF6600", None]


@pytest.mark.parametrize("sheet, message", [
    (b"path:first-child{fill:red}", "selector"),
    (b"@media print{.a{fill:red}}", "at-rule"),
    (b".a{font:12px serif}", "'font'"),
])
def test_styles_that_cannot_be_carried_over_are_refused(sheet, message):
    svg = (b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10" width="10" '
           b'height="10"><style>' + sheet + b'</style><path class="a" d="M0 0h9v9z"/></svg>')
    deck, slide, picture = _picture(svg)
    before = picture.vector
    with pytest.raises(ValueError, match=message):
        picture.recolor("accent2")
    assert picture.vector == before


# -- §4: a shared SVG part ----------------------------------------------------------------------


def test_a_shared_picture_is_split_as_powerpoint_splits_it():
    deck, slide, first = _picture(SOLID)
    _, _, second = _picture(SOLID, deck=deck, slide=slide)

    def embeds(picture):
        blip = picture._element.find(".//{*}blip")
        return [node.get(qn("r:embed")) for node in (blip, svgpicture.svg_blip(blip))]

    shared = embeds(first)
    assert embeds(second) == shared                    # one pair of relationships
    svg_part = deck.package.related_part(slide.part_path, shared[1])
    png_part = first.image_part
    before = deck.to_bytes()
    first.recolor("accent2")
    assert embeds(first) == shared                     # the recoloured one keeps them
    assert deck.package.related_part(slide.part_path, shared[1]) == svg_part
    assert first.image_part == png_part
    assert b"MsftOfcThm_Accent2_Fill_v2" in first.vector
    assert set(embeds(second)).isdisjoint(shared)      # the other gets new ones...
    assert second.vector == SOLID                      # ...to copies of the originals
    assert second.image_part not in (png_part, None)
    assert deck.validate() == []
    deck.undo()                                        # one step, the split included
    assert deck.to_bytes() == before


def test_a_picture_shared_with_another_slide_gets_parts_of_its_own():
    deck, slide, first = _picture(SOLID)
    _, other_slide, second = _picture(SOLID, deck=deck)
    part = first.image_part
    first.recolor("accent2")
    assert second.vector == SOLID and second.image_part == part
    assert first.image_part != part
    assert deck.validate() == []


# -- §5: transparency ---------------------------------------------------------------------------


def test_a_transparency_is_fill_opacity_and_drawn_once():
    deck, slide, picture = _picture(SOLID)
    picture.recolor("accent2 alpha=50%")
    svg = picture.vector.decode()
    assert svg.count('class="MsftOfcThm_Accent2_Fill_v2" ' + NO_STROKE +
                     ' fill="#E97132" fill-opacity="0.5"') == 2
    assert "opacity=" not in svg.replace("fill-opacity=", "")
    pixel = _pixel(deck.package.read(picture.image_part), 0.5, 0.5)
    assert all(abs(a - b) <= 1 for a, b in zip(pixel, (233, 113, 50, 128)))
    picture.recolor("#FF6600 alpha=50%")               # RGB: no class; alpha once, not twice
    svg = picture.vector.decode()
    assert "class=" not in svg and svg.count('fill="#FF6600" fill-opacity="0.5"') == 2
    assert abs(_pixel(deck.package.read(picture.image_part), 0.5, 0.5)[3] - 128) <= 1


# -- §6: gradients, use and symbol --------------------------------------------------------------


def test_gradients_symbols_and_uses_are_filled():
    deck, slide, picture = _picture(REUSED)
    picture.recolor("accent2")
    root = etree.fromstring(picture.vector)
    paint = ('class="MsftOfcThm_Accent2_Fill_v2" ' + NO_STROKE +
             ' fill="#E97132" fill-opacity="1"')
    svg = picture.vector.decode()
    assert svg.count(paint) == 4                       # the rect, the circle, both uses
    assert root.find(f".//{{{SVG}}}linearGradient").get("id") == "fade"   # left, unused
    assert "url(#fade)" not in svg
    assert root.find(f".//{{{SVG}}}symbol/{{{SVG}}}circle").get("fill") == "#E97132"


# -- the tool -----------------------------------------------------------------------------------


def _ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def test_set_shape_line_is_a_graphics_outline_on_an_svg_and_a_border_on_a_bitmap():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        deck = Document.new()
        deck.add_slide("Blank")
        session.open(deck.to_bytes(), "deck.pptx")
        slide = f"s:{deck.slides[0].slide_id}"
        icon = _ok(toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "image": session.add_blob(LINES, "lines.svg"), "slide": slide,
            "box": {"x": 72, "y": 72, "w": 72, "h": 72}})).data["address"]
        photo = _ok(toolbox.dispatch(session, "ppt_add_picture", {
            "doc": "d1", "image": session.add_blob(ORANGE, "orange.png"),
            "slide": slide})).data["address"]
        _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": icon, "fill": "accent2", "line": {"color": "accent5", "width": 1.5}},
            {"target": photo, "line": {"color": "accent5", "width": 1.5}}]}))
        document = session.documents["d1"].document
        svg = document.shape(icon).vector.decode()
        assert ('class="MsftOfcThm_Accent2_Fill_v2 MsftOfcThm_Accent5_Stroke_v2" '
                'stroke="#A02B93" stroke-width="1"') in svg
        assert document.shape(icon).line.visible is not True   # no border drawn
        border = document.shape(photo).line
        assert str(border.color) == "accent5" and border.width == 19050
        refused = toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": icon, "line": {"color": "accent1", "dash": "dash"}}]})
        assert not refused.ok and "color and width only" in refused.error.message
        _ok(toolbox.dispatch(session, "ppt_set_shape", {"doc": "d1", "items": [
            {"target": icon, "line": {"dash": "dash"}}]}))         # the border, as before
