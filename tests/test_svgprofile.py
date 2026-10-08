"""The SVG authoring profile (spike): apply_svg_graphic turns restricted SVG into native shapes."""

from __future__ import annotations

import math

import pytest

from pptx_agent import Document
from pptx_agent.edit.svgprofile import (SvgProfileError, apply_svg_graphic, content_area,
                                        measure_svg_text)
from pptx_agent.edit.svgprofile import parse_color, parse_path, parse_transform

PT = 12700
BOX = (914400, 914400, 100 * PT, 50 * PT)   # 100 x 50 pt
HEAD = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50">'


def slide():
    deck = Document.new()
    return deck, deck.add_slide("Blank")


def apply(body, **kw):
    deck, s = slide()
    result = apply_svg_graphic(s, HEAD + body + "</svg>", box=kw.get("box", BOX))
    return deck, s, result


# -- units ---------------------------------------------------------------------------------


@pytest.mark.parametrize("value, expected", [
    ("accent1", "accent1"),
    ("dk2/tint40", "dk2 lumMod=60% lumOff=40%"),
    ("accent2/shade25", "accent2 lumMod=75%"),
    ("#abc", "#AABBCC"),
    ("#0b6e79", "#0B6E79"),
    ("none", None),
])
def test_colours(value, expected):
    assert parse_color(value, "x") == expected


def test_colour_opacity_is_alpha():
    assert parse_color("accent1", "x", 0.5) == "accent1 alpha=50%"


@pytest.mark.parametrize("bad", ["red", "accent9", "accent1/tint", "accent1/light20", "#12"])
def test_bad_colours(bad):
    with pytest.raises(SvgProfileError):
        parse_color(bad, "<rect>")


def test_transforms_compose():
    m = parse_transform("translate(10 20) scale(2) rotate(90)", "x")
    x, y = m.apply(1, 0)
    assert (round(x, 6), round(y, 6)) == (10, 22)
    with pytest.raises(SvgProfileError, match="skewX"):
        parse_transform("skewX(10)", "<g>")


def test_rotate_about_a_point():
    m = parse_transform("rotate(180 5 5)", "x")
    assert tuple(round(v, 6) for v in m.apply(0, 0)) == (10, 10)


def test_path_commands():
    subs = parse_path("M0 0 h10 v10 H0 z m 20 0 l 5 5 q 5 5 10 0 t 10 0 "
                      "c 1 1 2 2 3 3 s 4 4 5 5 A 5 5 0 0 1 60 20", "x")
    assert len(subs) == 2
    assert subs[0][-1] == ("Z",)
    kinds = [seg[0] for seg in subs[1]]
    assert kinds[:2] == ["M", "L"] and kinds.count("C") >= 5
    assert subs[1][-1][-1] == pytest.approx((60, 20))


def test_arc_is_a_circle():
    (sub,) = parse_path("M 0 0 A 10 10 0 0 1 20 0", "x")
    # a half circle of radius 10 around (10, 0): the middle point lies at distance 10
    c = sub[-1]
    mid = c[1]
    assert sub[-1][-1] == pytest.approx((20, 0))
    assert all(abs(math.hypot(p[0] - 10, p[1]) - 10) < 3 for p in [s[-1] for s in sub[1:]])
    del mid


# -- the viewBox -----------------------------------------------------------------------------


def test_viewbox_maps_to_box():
    _, s, r = apply('<rect id="a" x="10" y="5" width="20" height="10" fill="accent1"/>')
    a = s.shape(r.ids["a"])
    assert a.bounds == (914400 + 10 * PT, 914400 + 5 * PT, 20 * PT, 10 * PT)


def test_viewbox_meet_centres_a_different_aspect():
    deck, s = slide()
    r = apply_svg_graphic(s, '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 50 50">'
                          '<rect id="a" x="0" y="0" width="50" height="50"/></svg>', box=BOX)
    a = s.shape(r.ids["a"])
    assert a.bounds == (914400 + 25 * PT, 914400, 50 * PT, 50 * PT)
    assert r.warnings and "aspect" in r.warnings[0]


def test_default_box_is_the_content_area():
    deck = Document.new()
    s = deck.add_slide("Title and Content")
    left, top, width, height = content_area(s)
    body = next(p for p in s.layout.placeholders if p.type is None)
    assert (left, top, width, height) == tuple(body.bounds)


# -- element kinds ----------------------------------------------------------------------------


def test_rect_preset_and_round_rect():
    _, s, r = apply('<rect id="a" width="20" height="10"/>'
                    '<rect id="b" x="30" width="20" height="10" rx="2"/>'
                    '<rect id="c" x="60" width="30" height="10" data-preset="chevron" '
                    'data-adj-adj="30000"/>')
    assert s.shape(r.ids["a"]).preset == "rect"
    b = s.shape(r.ids["b"])
    assert b.preset == "roundRect" and b.adjustments["adj"] == 20000
    c = s.shape(r.ids["c"])
    assert c.preset == "chevron" and c.adjustments["adj"] == 30000


def test_default_fill_is_theme_dark_and_no_outline():
    _, s, r = apply('<rect id="a" width="20" height="10"/>')
    a = s.shape(r.ids["a"])
    assert str(a.fill.color) == "tx1" and a.line.visible is False


def test_ellipse_and_circle():
    _, s, r = apply('<ellipse id="e" cx="20" cy="20" rx="10" ry="5"/><circle id="c" cx="50" cy="20" r="5"/>')
    e, c = s.shape(r.ids["e"]), s.shape(r.ids["c"])
    assert e.preset == c.preset == "ellipse"
    assert e.bounds == (914400 + 10 * PT, 914400 + 15 * PT, 20 * PT, 10 * PT)


def test_path_with_preset_uses_its_extent():
    _, s, r = apply('<path id="h" data-preset="homePlate" d="M10 10 H40 L45 15 L40 20 H10 Z"/>')
    h = s.shape(r.ids["h"])
    assert h.preset == "homePlate" and h.bounds == (914400 + 10 * PT, 914400 + 10 * PT, 35 * PT, 10 * PT)


def test_plain_path_polygon_polyline_are_custom_geometry():
    _, s, r = apply('<path id="p" d="M0 0 L10 0 L5 10 Z" fill="accent2"/>'
                    '<polygon id="g" points="20,0 30,0 25,10"/>'
                    '<polyline id="l" points="40,0 45,10 50,0" stroke="accent3" fill="none"/>')
    for key in "pgl":
        assert s.shape(r.ids[key]).preset == "custom"
    assert s.shape(r.ids["l"]).fill.kind == "none"
    assert s.shape(r.ids["p"]).bounds == (914400, 914400, 10 * PT, 10 * PT)


def test_line_stroke_dash_arrow():
    _, s, r = apply('<line id="l" x1="0" y1="10" x2="50" y2="10" stroke="accent1" '
                    'stroke-width="1.5" stroke-dasharray="6 3" data-arrow-end="triangle"/>')
    line = s.shape(r.ids["l"])
    assert line.kind == "connector" and line.line.width == round(1.5 * PT)
    assert line.line.dash in ("dash", "sysDash") and line.line.tail.type == "triangle"
    assert str(line.line.color) == "accent1"


def test_explicit_dash():
    _, s, r = apply('<line id="l" x1="0" y1="10" x2="50" y2="10" stroke="tx1" data-dash="sysDot"/>')
    assert s.shape(r.ids["l"]).line.dash == "sysDot"


def test_opacity_multiplies():
    _, s, r = apply('<g opacity="0.5"><rect id="a" width="10" height="10" fill="accent1" opacity="0.5"/>'
                    '<rect width="5" height="5" x="20"/></g>')
    assert str(s.shape(r.ids["a"]).fill.color) == "accent1 alpha=25000"


def test_gradient():
    _, s, r = apply('<defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1">'
                    '<stop offset="0" stop-color="accent1"/><stop offset="100%" stop-color="accent1/shade50"/>'
                    '</linearGradient></defs><rect id="a" width="10" height="10" fill="url(#g)"/>')
    assert s.shape(r.ids["a"]).fill.kind == "gradient"


def test_transforms_on_shapes():
    _, s, r = apply('<rect id="a" x="0" y="0" width="20" height="10" transform="rotate(90 10 5)"/>'
                    '<g transform="translate(50 0) scale(2)"><rect id="b" width="10" height="5"/></g>'
                    '<rect id="c" x="0" y="20" width="10" height="10" transform="scale(-1 1) translate(-30 0)"/>')
    a = s.shape(r.ids["a"])
    assert a.rotation == 90 and a.bounds == (914400, 914400, 20 * PT, 10 * PT)
    assert s.shape(r.ids["b"]).bounds == (914400 + 50 * PT, 914400, 20 * PT, 10 * PT)
    c = s.shape(r.ids["c"])
    assert c.flip_h and c.bounds[0] == 914400 + 20 * PT


def test_skew_on_a_preset_is_refused():
    with pytest.raises(SvgProfileError, match="skew"):
        apply('<rect width="10" height="10" transform="matrix(1 0 0.5 1 0 0)"/>')


# -- text -----------------------------------------------------------------------------------


def test_text_box_group_paragraphs_runs_bullets():
    _, s, r = apply(
        '<g id="t" data-text-box="" data-anchor="top" data-inset="4 6">'
        '<rect width="100" height="50" fill="lt2"/>'
        '<text font-size="10" fill="tx2" text-anchor="middle">'
        '<tspan font-weight="bold" font-size="12">Diagnose</tspan>'
        '<tspan data-bullet="" data-space-before="3">Build the <tspan fill="accent1" font-style="italic">'
        'spend</tspan>   cube</tspan>'
        '<tspan data-bullet="–" data-level="1" data-align="left">Sub point</tspan>'
        '</text></g>')
    t = s.shape(r.ids["t"])
    frame = t.text_frame
    assert frame.anchor == "top" and frame.autofit == "none" and frame.wrap is True
    assert frame.insets == (6 * PT, 4 * PT, 6 * PT, 4 * PT)
    p0, p1, p2 = frame.paragraphs
    assert p0.text == "Diagnose" and p0.runs[0].bold and p0.runs[0].size == 12
    assert p0.alignment == "center"
    assert p1.text == "Build the spend cube"
    assert [run.text for run in p1.runs] == ["Build the ", "spend", " cube"]
    assert str(p1.runs[1].color) == "accent1" and p1.runs[1].italic
    assert p1.bullet is not None and p1.indent < 0 and p1.space_before == 3 * PT
    assert p2.alignment == "left" and p2.margin_left > p1.margin_left
    assert r.text_fits[t.id].lines == (1, 2, 1)   # measured as laid out


def test_free_text_needs_a_box():
    with pytest.raises(SvgProfileError, match="data-width"):
        apply('<text x="0" y="10">Floating</text>')


def test_free_text_box_is_exact():
    _, s, r = apply('<text id="t" x="10" y="10" data-width="40" data-height="10" font-size="9">Label</text>')
    t = s.shape(r.ids["t"])
    assert t.bounds == (914400 + 10 * PT, 914400 + 10 * PT, 40 * PT, 10 * PT)
    assert t.text_frame.autofit == "none" and t.text_frame.anchor == "top" and t.fill.kind == "none"


def test_shape_with_a_text_child():
    _, s, r = apply('<ellipse id="e" cx="20" cy="20" rx="15" ry="10" fill="accent1">'
                    '<text fill="bg1" text-anchor="middle">Hub</text></ellipse>')
    e = s.shape(r.ids["e"])
    assert e.text == "Hub" and e.preset == "ellipse" and str(e.text_frame.paragraph(0).run(0).color) == "bg1"


def test_overflow_is_reported():
    _, s, r = apply('<text id="t" x="0" y="0" data-width="20" data-height="5" font-size="12">'
                    'Far too much text for a tiny box</text>')
    assert r.text_fits[r.ids["t"]].overflows
    assert any(o.kind == "text" for o in r.overflows)
    assert not r.ok


def test_loose_text_between_paragraphs_is_refused():
    with pytest.raises(SvgProfileError, match="loose text"):
        apply('<text data-width="10" data-height="10">a<tspan>b</tspan></text>')


# -- connectors and groups -------------------------------------------------------------------


def test_connectors_are_glued_and_pick_sides():
    _, s, r = apply('<line data-from="b" data-to="a" stroke="tx1"/>'
                    '<rect id="a" x="0" y="0" width="20" height="10"/>'
                    '<rect id="b" x="40" y="0" width="20" height="10"/>'
                    '<rect id="c" x="40" y="30" width="20" height="10"/>'
                    '<path d="M0 0 L1 1" data-from="a" data-to="c" data-from-side="bottom" '
                    'data-to-side="left" stroke="tx1"/>')
    conns = [sh for sh in s.shapes if sh.kind == "connector"]
    first, second = conns
    assert first.begin_connection[0].id == r.ids["b"] and first.end_connection[0].id == r.ids["a"]
    assert first.preset == "straightConnector1"           # sides line up: straight
    assert second.preset.startswith("bentConnector")       # they do not: elbow
    order = [sh.id for sh in s.shapes]
    assert order.index(first.id) < order.index(r.ids["a"])   # z-order kept: behind
    s.shape(r.ids["a"]).move_by(dy=PT)
    assert first.begin_connection is not None              # still glued after a move


def test_unknown_connector_target():
    with pytest.raises(SvgProfileError, match="names no element"):
        apply('<line data-from="nope" data-to="a" stroke="tx1"/><rect id="a" width="1" height="1"/>')


def test_groups_nest_and_single_member_groups_flatten():
    _, s, r = apply('<g id="outer"><g id="inner"><rect width="5" height="5"/><rect x="10" width="5" height="5"/></g>'
                    '<rect x="20" width="5" height="5"/></g><g id="solo"><rect x="40" width="5" height="5"/></g>')
    outer = s.shape(r.ids["outer"])
    assert outer.kind == "group"
    kinds = sorted(child.kind for child in outer.children)
    assert kinds == ["group", "shape"]
    assert "solo" not in r.ids


def test_ids_twice():
    with pytest.raises(SvgProfileError, match="twice"):
        apply('<rect id="a" width="1" height="1"/><rect id="a" width="1" height="1"/>')


# -- unsupported features ---------------------------------------------------------------------


@pytest.mark.parametrize("body, needle", [
    ('<filter id="f"/>', "filter"),
    ('<rect width="1" height="1" filter="url(#f)"/>', "filter"),
    ('<mask id="m"/>', "mask"),
    ('<defs><pattern id="p"/></defs>', "pattern"),
    ('<defs><radialGradient id="r"/></defs>', "linearGradient"),
    ('<rect class="box" width="1" height="1"/>', "CSS classes"),
    ('<style>rect{fill:red}</style>', "CSS"),
    ('<foreignObject width="1" height="1"/>', "foreignObject"),
    ('<image href="x.png" width="1" height="1"/>', "add_picture"),
    ('<line x1="0" y1="0" x2="1" y2="1" stroke="tx1" marker-end="url(#a)"/>', "data-arrow-end"),
    ('<rect width="1" height="1" style="filter: blur(2px)"/>', "style property"),
    ('<rect width="1" height="1" fill="red"/>', "not in the profile"),
])
def test_unsupported_features_name_the_element(body, needle):
    with pytest.raises(SvgProfileError, match=needle) as info:
        apply(body)
    assert "line" in str(info.value) or "<" in str(info.value)


def test_nothing_changes_on_error():
    deck, s = slide()
    before = len(s.shapes)
    with pytest.raises(SvgProfileError):
        apply_svg_graphic(s, HEAD + '<rect width="1" height="1"/><rect class="x" width="1" height="1"/></svg>',
                          box=BOX)
    assert len(s.shapes) == before


def test_entities_are_refused():
    deck, s = slide()
    with pytest.raises(SvgProfileError):
        apply_svg_graphic(s, '<!DOCTYPE svg [<!ENTITY x "y">]><svg xmlns="http://www.w3.org/2000/svg" '
                          'viewBox="0 0 1 1"><text data-width="1" data-height="1">&x;</text></svg>')


def test_one_undo_step_and_valid():
    deck, s, r = apply('<rect width="10" height="10"/><g data-text-box=""><rect x="20" width="30" height="10"/>'
                       '<text>Hi</text></g><line x1="0" y1="20" x2="10" y2="20" stroke="tx1"/>')
    assert deck.validate() == []
    assert len(s.shapes) == 3
    deck.undo()
    assert len(deck.slides[0].shapes) == 0


# -- spike round 2: escaping, error locations, insets on the shape, tree sides, measuring ----


from pptx_agent.edit.svgprofile import auto_sides, repair_markup  # noqa: E402


@pytest.mark.parametrize("source, expected", [
    ("R&D", "R&amp;D"),
    ("Q&A &amp; more", "Q&amp;A &amp; more"),
    ("&lt;&gt;&quot;&apos;&#8211;&#x2013;", "&lt;&gt;&quot;&apos;&#8211;&#x2013;"),
    ("Spend & baseline", "Spend &amp; baseline"),
    ("a &ndash; b&nbsp;c", "a &#8211; b&#160;c"),
    ("&foo; bar", "&amp;foo; bar"),
    ("x < 5% and y <= 3", "x &lt; 5% and y &lt;= 3"),
    ("<!-- a & b --><![CDATA[R&D]]>", "<!-- a & b --><![CDATA[R&D]]>"),
])
def test_repair_markup(source, expected):
    assert repair_markup(source)[0] == expected


def test_bare_ampersand_in_text_and_attribute():
    _, s, r = apply('<g id="t" data-text-box="" data-name="R&D box"><rect width="80" height="20" fill="lt2"/>'
                    '<text font-size="9">Spend & baseline, R&D &amp; Q&A &ndash; done</text></g>')
    shape = s.shape(r.ids["t"])
    assert shape.text == "Spend & baseline, R&D & Q&A – done"
    assert shape.name == "R&D box"
    assert any("bare '&'" in w for w in r.warnings)


def test_entities_still_refused_after_repair():
    deck, s = slide()
    with pytest.raises(SvgProfileError, match="DTD"):
        apply_svg_graphic(s, '<!DOCTYPE svg [<!ENTITY x "y">]><svg xmlns="http://www.w3.org/2000/svg" '
                          'viewBox="0 0 1 1"/>', box=BOX)


def test_parse_error_names_the_nearest_element_by_id():
    deck, s = slide()
    svg = (HEAD + '\n<rect id="q1" width="10" height="10"/>\n'
           '<g id="box2" data-text-box=""><rect width="10" height="10"/><text>oops</txt></g></svg>')
    with pytest.raises(SvgProfileError) as info:
        apply_svg_graphic(s, svg, box=BOX)
    message = str(info.value)
    assert "<text> #1" in message and "line 3" in message


def test_parse_error_names_the_element_with_an_id():
    deck, s = slide()
    svg = HEAD + '<rect id="q1" width="10" height="10"/><rect id="q2" width="10" height=10/></svg>'
    with pytest.raises(SvgProfileError, match="<rect id='q2'>"):
        apply_svg_graphic(s, svg, box=BOX)


def test_profile_errors_without_an_id_give_the_tag_index():
    with pytest.raises(SvgProfileError, match=r"<rect> #3"):
        apply('<rect width="1" height="1"/><rect width="1" height="1"/><rect width="1" height="1" fill="red"/>')


def test_inset_on_the_shape_inside_a_text_box_group():
    _, s, r = apply('<g id="t" data-text-box=""><rect width="80" height="20" data-inset="2 5" '
                    'data-anchor="top"/><text>Hi</text></g>')
    frame = s.shape(r.ids["t"]).text_frame
    assert frame.insets == (5 * PT, 2 * PT, 5 * PT, 2 * PT)
    assert frame.anchor == "top"


def test_most_specific_frame_attribute_wins():
    _, s, r = apply('<g id="t" data-text-box="" data-inset="9"><rect width="80" height="20" data-inset="4"/>'
                    '<text data-inset="1">Hi</text></g>')
    assert s.shape(r.ids["t"]).text_frame.insets == (PT, PT, PT, PT)
    _, s, r = apply('<g id="t" data-text-box="" data-inset="9"><rect width="80" height="20" data-inset="4"/>'
                    '<text>Hi</text></g>')
    assert s.shape(r.ids["t"]).text_frame.insets == (4 * PT,) * 4


def test_adjustment_names_are_checked():
    with pytest.raises(SvgProfileError, match=r"<rect id='c'>.*data-adj-adj \(default 50000\)"):
        apply('<rect id="c" data-preset="chevron" data-adj-adj1="30000" width="20" height="10"/>')
    _, s, r = apply('<rect id="a" data-preset="rightArrow" data-adj-adj1="40000" data-adj-adj2="60000" '
                    'width="20" height="10"/>')
    assert s.shape(r.ids["a"]).preset == "rightArrow"


@pytest.mark.parametrize("target, sides", [
    ((0, 40, 20, 10), ("bottom", "top")),       # far left, wholly below: a tree
    ((80, 40, 20, 10), ("bottom", "top")),      # far right, wholly below
    ((40, -30, 20, 10), ("top", "bottom")),     # wholly above
    ((80, 2, 20, 10), ("right", "left")),       # beside, overlapping vertically
    ((0, 5, 20, 10), ("left", "right")),
])
def test_auto_sides(target, sides):
    assert auto_sides((40, 0, 20, 10), target) == sides


def test_org_chart_connectors_go_bottom_to_top():
    _, s, r = apply('<line data-from="lead" data-to="w1"/><line data-from="lead" data-to="w4"/>'
                    '<rect id="lead" x="40" y="0" width="20" height="8"/>'
                    '<rect id="w1" x="0" y="30" width="15" height="8"/>'
                    '<rect id="w4" x="85" y="30" width="15" height="8"/>')
    lead = s.shape(r.ids["lead"])
    for conn in (sh for sh in s.shapes if sh.kind == "connector"):
        begin, end = conn.begin_connection, conn.end_connection
        assert begin[0].id == lead.id
        assert begin[1] == lead.connection_site("bottom")[1]
        assert end[1] == end[0].connection_site("top")[1]
        assert conn.preset.startswith("bentConnector")


def test_one_explicit_side_keeps_the_automatic_other():
    _, s, r = apply('<line data-from="a" data-to="b" data-from-side="right"/>'
                    '<rect id="a" x="0" y="0" width="20" height="8"/><rect id="b" x="40" y="30" width="20" height="8"/>')
    conn = next(sh for sh in s.shapes if sh.kind == "connector")
    b = s.shape(r.ids["b"])
    assert conn.end_connection[1] == b.connection_site("top")[1]


def test_measure_matches_apply_exactly():
    body = ('<g id="t" data-text-box=""><rect width="60" height="12" data-inset="1 3" fill="accent1"/>'
            '<text font-size="9" font-weight="bold">Programme Management Office and reporting</text></g>'
            '<text id="n" x="0" y="20" data-width="40" data-height="30"><tspan data-bullet="">one two three</tspan>'
            '<tspan data-bullet="">four</tspan></text>'
            '<line data-from="t" data-to="n"/>')
    deck, s = slide()
    measured = measure_svg_text(s, HEAD + body + "</svg>", box=BOX)
    assert len(s.shapes) == 0                          # nothing changed
    assert set(measured) == {"t", "n"}
    _, s2, r = apply(body)
    for key in ("t", "n"):
        fit = r.text_fits[r.ids[key]]
        m = measured[key]
        assert m.needed == pytest.approx(fit.needed / PT)
        assert m.fits == (not fit.overflows)
        assert m.lines == tuple(fit.lines)


def test_measure_a_fragment_and_height_to_fit():
    deck, s = slide()
    fragment = ('<g id="t" data-text-box="" data-inset="2"><rect width="100" height="20"/>'
                '<text font-size="14" font-weight="bold">Programme lead and R&amp;D office</text></g>')
    m = measure_svg_text(s, fragment)["t"]
    assert not m.fits and m.lines == (3,)
    assert m.needed == pytest.approx(3 * 14 * 1.2 + 4, abs=0.6)
    fixed = fragment.replace('height="20"', f'height="{m.height_to_fit}"')
    again = measure_svg_text(s, fixed)["t"]
    assert again.fits
    assert len(s.shapes) == 0


def test_measure_bold_is_measured_bold():
    deck, s = slide()
    text = "Workstream lead accountable for delivery"
    width = None
    for w in range(60, 300, 2):          # the narrowest width at which plain text is one line
        if measure_svg_text(s, f'<text id="x" data-width="{w}" data-height="20">{text}</text>')["x"].lines == (1,):
            width = w
            break
    assert width is not None
    bold = measure_svg_text(s, f'<text id="x" data-width="{width}" data-height="20" '
                               f'font-weight="bold">{text}</text>')["x"]
    assert bold.lines[0] > 1
