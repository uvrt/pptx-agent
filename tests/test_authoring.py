"""Shapes built from nothing (E5): autoshapes, text boxes, connectors, tables, z-order steps.

Each creator goes through the gates every edit goes through -- create, save, reopen and
read back; the validity checks, schema order included; undo to the original bytes and redo
to the edited ones; a pptx2svg render with the API's ids -- on its own, and all of them
together on a new slide of every fixture (:mod:`authoring`).  The focused tests below pin
the XML to what PowerPoint was measured to write (see :mod:`pptx_agent.edit.authoring`).
"""

from __future__ import annotations

import re
from typing import Callable

import pytest

from authoring import acceptance_edits
from pptx_agent import Document
from pptx_agent.edit.authoring import DEFAULT_TABLE_STYLE, SHAPE_NAMES
from pptx_agent.edit.ids import read_stamp
from pptx_agent.oxml.xml import PRESET_GEOMETRIES, qn
from svgedit import P, FullStateSvg
from test_validity import assert_valid

EMU = 914400
Verify = Callable[[Document], None]


def _xml(shape) -> str:
    from lxml import etree

    return etree.tostring(shape._element).decode()


# ------------------------------------------------------------------------------------------
# One creator at a time
# ------------------------------------------------------------------------------------------


def make_shape(document: Document) -> Verify:
    slide = document.slides[0]
    shape = slide.add_shape("roundRect", EMU, EMU, 2 * EMU, EMU, text="Made",
                            adjustments={"adj": 25000}, fill="accent2",
                            line={"width": 25400, "color": "tx1", "dash": "dash"})

    def verify(reopened: Document) -> None:
        again = reopened.shape(shape.id)
        assert again.name == shape.name and again.preset == "roundRect"
        assert again.text == "Made" and again.adjustments["adj"] == 25000
        assert again.fill.color.value == "accent2" and again.line.dash == "dash"
        assert (again.left, again.top, again.width, again.height) == (EMU, EMU, 2 * EMU, EMU)

    return verify


def make_textbox(document: Document) -> Verify:
    box = document.slides[0].add_textbox(EMU, 2 * EMU, 3 * EMU, EMU // 2, "Typed\nTwo lines")

    def verify(reopened: Document) -> None:
        again = reopened.shape(box.id)
        assert again.text == "Typed\nTwo lines" and again.name.startswith("TextBox ")

    return verify


def make_connectors(document: Document) -> Verify:
    slide = document.slides[0]
    a = slide.add_shape("rect", EMU // 2, 3 * EMU, EMU, EMU // 2)
    b = slide.add_shape("ellipse", 5 * EMU, 4 * EMU, EMU, EMU)
    made = [slide.add_connector(kind, (a, 3), (b, 2), line={"tail": "triangle"}).id
            for kind in ("straight", "elbow", "curved")]
    document.shape(b.id).move_by(-2 * EMU, EMU // 2)

    def verify(reopened: Document) -> None:
        for identifier in made:
            connector = reopened.shape(identifier)
            assert connector.kind == "connector" and connector.line.tail.type == "triangle"
            assert connector.begin_connection[0].id == a.id
            assert (connector.end_connection[0].id, connector.end_connection[1]) == (b.id, 2)
        from test_connectors import assert_attached

        for identifier in made:
            assert_attached(reopened, identifier)

    return verify


def make_table(document: Document) -> Verify:
    frame = document.slides[0].add_table(3, 4, EMU, 4 * EMU, 6 * EMU, EMU)
    table = frame.table
    table.cell(0, 0).text = "Head"
    table.insert_row(1)
    table.insert_column(4)
    table.merge(1, 1, 2, 2)
    table.cell(3, 4).text = "Corner"

    def verify(reopened: Document) -> None:
        again = reopened.shape(frame.id).table
        assert (again.rows, again.columns) == (4, 5)
        assert again.cell(0, 0).text == "Head" and again.cell(3, 4).text == "Corner"
        assert again.cell(1, 1).span == (2, 2)

    return verify


def make_in_group(document: Document) -> Verify:
    slide = document.slides[0]
    a = slide.add_shape("rect", EMU, EMU, EMU, EMU)
    b = slide.add_shape("rect", 3 * EMU, 2 * EMU, EMU, EMU)
    group = slide.group([a, b])
    child = group.add_shape("star5", 2 * EMU, 3 * EMU, EMU, EMU, text="In")

    def verify(reopened: Document) -> None:
        again = reopened.shape(child.id)
        assert again.parent_group.id == group.id and again.text == "In"
        assert again.slide_bounds == (2 * EMU, 3 * EMU, EMU, EMU)

    return verify


def make_z_steps(document: Document) -> Verify:
    slide = document.slides[0]
    first = slide.add_shape("rect", 0, 0, EMU, EMU)
    second = slide.add_shape("ellipse", 0, 0, EMU, EMU)
    first.bring_forward()

    def verify(reopened: Document) -> None:
        order = [shape.id for shape in reopened.slides[0].shapes]
        assert order.index(second.id) < order.index(first.id)

    return verify


CREATORS: dict[str, Callable[[Document], Verify]] = {
    "shape": make_shape,
    "textbox": make_textbox,
    "connectors": make_connectors,
    "table": make_table,
    "in_group": make_in_group,
    "z_steps": make_z_steps,
}
creator = pytest.mark.parametrize("name", list(CREATORS))


@creator
def test_each_creator_round_trips(name, product_page):
    document = Document.open(product_page.read_bytes())
    verify = CREATORS[name](document)
    saved = document.to_bytes()
    verify(Document.open(saved))
    assert Document.open(saved).to_bytes() == saved


@creator
def test_each_creator_keeps_the_deck_valid(name, pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    CREATORS[name](document)
    assert_valid(document.to_bytes(), original)


@creator
def test_each_creator_undoes_byte_for_byte(name, product_page):
    document = Document.open(product_page.read_bytes())
    document.history._max_depth = 10_000
    before = document.to_bytes()
    CREATORS[name](document)
    edited = document.to_bytes()
    assert edited != before
    while document.undo():
        pass
    assert document.to_bytes() == before
    while document.redo():
        pass
    assert document.to_bytes() == edited


def test_every_creation_is_one_undo_step(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    a = slide.add_shape("rect", 0, 0, EMU, EMU)
    b = slide.add_shape("rect", 2 * EMU, 0, EMU, EMU)
    for make in (
        lambda: slide.add_shape("star5", 0, 0, EMU, EMU, text="x", fill="accent1",
                                line="none", adjustments={"adj": 10000}),
        lambda: slide.add_textbox(0, 0, EMU, EMU, "x\ny", fill="bg1", line="tx1"),
        lambda: slide.add_connector("elbow", (a, 3), (b, 1), line={"tail": "arrow"}),
        lambda: slide.add_table(2, 2, 0, 0, EMU, EMU),
    ):
        before = document.to_bytes()
        make()
        assert document.undo()
        assert document.to_bytes() == before
        assert document.redo()


# ------------------------------------------------------------------------------------------
# Everything together, on every fixture
# ------------------------------------------------------------------------------------------


def test_the_acceptance_slide_is_valid(pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    acceptance_edits(document)
    assert_valid(document.to_bytes(), original)


def test_the_acceptance_slide_round_trips(pptx_path):
    document = Document.open(pptx_path.read_bytes())
    expected = acceptance_edits(document)
    reopened = Document.open(document.to_bytes())
    slide = reopened.slide(expected["slide_id"])
    texts = " ".join(shape.text for shape in slide.shapes if shape.kind == "shape")
    for sentinel in ("E5BOX", "E5KID", "E5FROM"):
        assert sentinel in texts
    table = next(shape for shape in slide.shapes if shape.has_table).table
    assert table.cell(1, 0).text == "E5CELL" and (table.rows, table.columns) == (4, 2)
    from test_connectors import assert_attached

    for identifier in expected["connectors"]:
        assert_attached(reopened, identifier)
    assert reopened.shape(expected["child"]).parent_group.id == expected["group"]


def test_the_acceptance_slide_undoes_byte_for_byte(pptx_path):
    original = pptx_path.read_bytes()
    document = Document.open(original)
    document.history._max_depth = 10_000
    before = document.to_bytes()
    acceptance_edits(document)
    edited = document.to_bytes()
    while document.undo():
        pass
    assert document.to_bytes() == before
    while document.redo():
        pass
    assert document.to_bytes() == edited


def test_the_acceptance_slide_renders_with_matching_ids(pptx_path):
    pytest.importorskip("pptx2svg")
    document = Document.open(pptx_path.read_bytes())
    expected = acceptance_edits(document)
    slide = document.slide(expected["slide_id"])
    svg = slide.render_svg()
    drawn = {i.partition("/")[0] for i in re.findall(r'data-pptx-id="([^"]*)"', svg)}
    made = [shape for shape in slide.shapes if shape.placeholder is None]  # empty ones hide
    assert len(made) == 16
    for shape in made:
        assert shape.id in drawn, shape


def test_the_acceptance_slide_survives_its_full_state_svg(pptx_path):
    """Emitted and applied to a fresh copy of the saved deck, nothing changes; and the SVG
    without its vocabulary is the plain render."""
    pytest.importorskip("pptx2svg")
    document = Document.open(pptx_path.read_bytes())
    expected = acceptance_edits(document)
    saved = document.to_bytes()
    slide = Document.open(saved).slide(expected["slide_id"])
    svg = slide.render_svg(full_state=True)
    copy = Document.open(saved)
    report = copy.apply_svg(svg)
    assert not report
    assert copy.to_bytes() == saved
    from lxml import etree
    from test_fullstate_emit import _strip_state

    assert _strip_state(svg) == etree.tostring(etree.fromstring(slide.render_svg().encode()),
                                               method="c14n")


# ------------------------------------------------------------------------------------------
# What PowerPoint writes (measured)
# ------------------------------------------------------------------------------------------


@pytest.fixture
def deck(product_page):
    return Document.open(product_page.read_bytes())


def test_a_new_shape_is_written_as_powerpoint_writes_one(deck):
    shape = deck.slides[0].add_shape("rect", EMU, EMU, EMU, EMU)
    xml = _xml(shape)
    assert '<p:cNvSpPr/>' in xml and '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' in xml
    assert ('<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/>'
            '</a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/>'
            '</a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>') in xml
    assert ('<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p>'
            '<a:pPr algn="ctr"/><a:endParaRPr lang="') in xml
    assert shape.fill is None and not shape.line.exists  # both come from the theme


def test_a_new_text_box_is_written_as_powerpoint_writes_one(deck):
    box = deck.slides[0].add_textbox(EMU, EMU, 2 * EMU, EMU // 2)
    xml = _xml(box)
    assert '<p:cNvSpPr txBox="1"/>' in xml and "<a:noFill/>" in xml and "p:style" not in xml
    assert ('<a:bodyPr vert="horz" wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>'
            '<a:lstStyle/><a:p><a:endParaRPr lang="') in xml
    assert box.text == ""


def test_a_new_connector_is_written_as_powerpoint_writes_one(deck):
    slide = deck.slides[0]
    a = slide.add_shape("rect", 0, 0, EMU, EMU)
    b = slide.add_shape("rect", 3 * EMU, 2 * EMU, EMU, EMU)
    connector = slide.add_connector("elbow", (a, 3), (b, 1))
    raw_a = a._element.find(".//" + qn("p:cNvPr")).get("id")
    raw_b = b._element.find(".//" + qn("p:cNvPr")).get("id")
    xml = _xml(connector)
    assert (f'<p:cNvCxnSpPr><a:cxnSpLocks/><a:stCxn id="{raw_a}" idx="3"/>'
            f'<a:endCxn id="{raw_b}" idx="1"/></p:cNvCxnSpPr>') in xml
    assert ('<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>'
            '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
            '<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>') in xml
    assert '<a:gd name="adj1" fmla="val 50000"/>' in xml  # a Z, bent halfway
    free = slide.add_connector("straight", (0, 0), (EMU, EMU))
    assert "<p:cNvCxnSpPr/>" in _xml(free)


def test_a_new_table_is_written_as_powerpoint_writes_one(deck):
    frame = deck.slides[0].add_table(2, 3, EMU, EMU, 1000001, 3 * EMU)
    xml = _xml(frame)
    assert '<a:graphicFrameLocks noGrp="1"/>' in xml
    assert (f'<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{DEFAULT_TABLE_STYLE}'
            '</a:tableStyleId></a:tblPr>') in xml
    table = frame.table
    assert sum(table.column_widths) == 1000001 and max(table.column_widths) - min(table.column_widths) <= 1
    assert table.row_heights == [3 * EMU // 2] * 2
    assert xml.count("<a16:colId") == 3 and xml.count("<a16:rowId") == 2
    assert xml.count("<a:tcPr/>") == 6
    other = deck.slides[0].add_table(1, 1, 0, 0, EMU, EMU,
                                     style="{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
    assert "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}" in _xml(other)
    with pytest.raises(ValueError):
        deck.slides[0].add_table(1, 1, 0, 0, EMU, EMU, style="Medium Style 2")
    with pytest.raises(ValueError):
        deck.slides[0].add_table(0, 1, 0, 0, EMU, EMU)


def test_names_are_numbered_as_powerpoint_numbers_them(deck):
    slide = deck.slides[0]
    shapes = [slide.add_shape("rect", 0, 0, EMU, EMU), slide.add_textbox(0, 0, EMU, EMU),
              slide.add_table(1, 1, 0, 0, EMU, EMU)]
    shapes.append(slide.add_connector("straight", (shapes[0], 0), (shapes[1], 2)))
    shapes.append(slide.add_connector("elbow", (0, 0), (EMU, EMU)))
    shapes.append(slide.add_connector("curved", (0, 0), (EMU, EMU)))
    bases = ["Rectangle", "TextBox", "Table", "Straight Arrow Connector", "Elbow Connector",
             "Curved Connector"]
    raw_ids = []
    for shape, base in zip(shapes, bases):
        raw = int(shape._element.find(".//" + qn("p:cNvPr")).get("id"))
        raw_ids.append(raw)
        assert shape.name == f"{base} {raw - 1}"
        assert shape.id == f"{slide.slide_id}.{raw}"
        assert read_stamp(shape._element.find(".//" + qn("p:cNvPr"))) == str(raw)
    every = [node.get("id") for node in slide._sp_tree().iter(qn("p:cNvPr"))]
    assert len(every) == len(set(every))
    named = slide.add_shape("ellipse", 0, 0, EMU, EMU, name="Mine")
    assert named.name == "Mine"


def test_every_preset_can_be_made(deck):
    slide = deck.slides[0]
    original = deck.to_bytes()
    for index, preset in enumerate(sorted(PRESET_GEOMETRIES)):
        shape = slide.add_shape(preset, (index % 16) * EMU // 2, (index // 16) * EMU // 2,
                                EMU // 2, EMU // 2)
        assert shape.preset == preset
        assert shape.name.rsplit(" ", 1)[0] == SHAPE_NAMES.get(preset, "Shape")
        assert shape.adjustments.names == tuple(n for n, _ in shape.adjustments.defaults.items())
    assert_valid(deck.to_bytes(), original)
    pptx2svg = pytest.importorskip("pptx2svg")
    assert len(pptx2svg.convert_pptx_to_svg(deck.to_bytes(), pptx2svg.ConvertOptions(
        slide_numbers=[1]))) == 1


def test_shape_creation_checks_its_arguments(deck):
    slide = deck.slides[0]
    with pytest.raises(ValueError):
        slide.add_shape("rectangle", 0, 0, EMU, EMU)
    with pytest.raises(ValueError):
        slide.add_shape("rect", 0, 0, -1, EMU)
    with pytest.raises(KeyError):
        slide.add_shape("rect", 0, 0, EMU, EMU, adjustments={"adj": 1})
    with pytest.raises(TypeError):
        slide.add_shape("rect", 0, 0, EMU, EMU, line=3)
    picture = next((s for s in slide.shapes if s.kind == "picture"), None)
    if picture is not None:
        with pytest.raises(ValueError):
            picture.add_shape("rect", 0, 0, EMU, EMU)


def test_the_text_language_is_the_decks(pptx_path):
    from pptx_agent.edit.authoring import default_language

    document = Document.open(pptx_path.read_bytes())
    shape = document.slides[0].add_shape("rect", 0, 0, EMU, EMU)
    assert f'lang="{default_language(document.package)}"' in _xml(shape)


# ------------------------------------------------------------------------------------------
# Adjustments
# ------------------------------------------------------------------------------------------


def test_adjustments_read_write_and_reset(deck):
    shape = deck.slides[0].add_shape("rightArrow", 0, 0, 2 * EMU, EMU)
    adjustments = shape.adjustments
    assert adjustments.names == ("adj1", "adj2") and dict(adjustments.items()) == {
        "adj1": 50000, "adj2": 50000}
    assert adjustments.explicit() == {}
    adjustments["adj2"] = 70000
    # Once one is set, all are written, defaults filled in, as PowerPoint writes them.
    assert adjustments.explicit() == {"adj1": 50000, "adj2": 70000}
    adjustments[0] = 30000
    assert adjustments.explicit() == {"adj1": 30000, "adj2": 70000}
    assert re.search(r'name="adj1".*name="adj2"', _xml(deck.shape(shape.id)))
    del adjustments["adj1"]
    assert adjustments["adj1"] == 50000 and adjustments.explicit() == {"adj1": 50000,
                                                                       "adj2": 70000}
    before = deck.to_bytes()
    adjustments.set(adj1=10000, adj2=20000)
    assert deck.undo() and deck.to_bytes() == before
    adjustments.reset()
    assert adjustments.explicit() == {}
    assert "adj1" in adjustments and len(adjustments) == 2
    with pytest.raises(KeyError):
        adjustments["adj3"] = 1
    with pytest.raises(TypeError):
        adjustments["adj1"] = 0.5
    reopened = Document.open(deck.to_bytes())
    assert reopened.shape(shape.id).adjustments.explicit() == {}


def test_a_star_gets_all_its_adjust_values(deck):
    """What PowerPoint wrote for a five-point star with its first adjustment set -- and
    without ``hf``/``vf`` PowerPoint repairs the file."""
    star = deck.slides[0].add_shape("star5", 0, 0, EMU, EMU, adjustments={"adj": 15000})
    assert ('<a:avLst><a:gd name="adj" fmla="val 15000"/><a:gd name="hf" fmla="val 105146"/>'
            '<a:gd name="vf" fmla="val 110557"/></a:avLst>') in _xml(star)
    plain = deck.slides[0].add_shape("star5", 0, 0, EMU, EMU)
    assert "<a:avLst/>" in _xml(plain)
    plain.adjustments["vf"] = 100000
    assert deck.shape(plain.id).adjustments.explicit() == {"adj": 19098, "hf": 105146,
                                                           "vf": 100000}


def test_setting_an_unchanged_adjustment_changes_nothing(deck):
    shape = deck.slides[0].add_shape("roundRect", 0, 0, EMU, EMU, adjustments={"adj": 5000})
    before = deck.to_bytes()
    depth = len(deck.history._undo)
    shape.adjustments["adj"] = 5000
    assert deck.to_bytes() == before and len(deck.history._undo) == depth


def test_only_preset_geometry_has_adjustments(deck):
    slide = deck.slides[0]
    table = slide.add_table(1, 1, 0, 0, EMU, EMU)
    with pytest.raises(ValueError):
        table.adjustments
    shape = slide.add_shape("rect", 0, 0, EMU, EMU)
    from pptx_agent.oxml.xml import make, replace_choice

    custom = make("a:custGeom")
    replace_choice(shape._element.find(qn("p:spPr")), ("a:custGeom", "a:prstGeom"), custom)
    with pytest.raises(ValueError):
        deck.shape(shape.id).adjustments.names


# ------------------------------------------------------------------------------------------
# Z-order steps
# ------------------------------------------------------------------------------------------


def test_bring_forward_and_send_backward_step_once(deck):
    slide = deck.slides[0]
    made = [slide.add_shape("rect", i * EMU, 0, EMU // 2, EMU // 2).id for i in range(4)]

    def order():
        ids = [shape.id for shape in deck.slides[0].shapes]
        return [identifier for identifier in ids if identifier in made]

    deck.shape(made[0]).bring_forward()       # past a shape it does not overlap, as measured
    assert order() == [made[1], made[0], made[2], made[3]]
    deck.shape(made[3]).send_backward()
    assert order() == [made[1], made[0], made[3], made[2]]
    before, depth = deck.to_bytes(), len(deck.history._undo)
    deck.shape(made[2]).bring_forward()       # frontmost already: nothing, not even an undo step
    assert deck.to_bytes() == before and len(deck.history._undo) == depth
    first = deck.slides[0].shapes[0]
    first.send_backward()
    assert deck.to_bytes() == before
    deck.shape(made[1]).send_backward()       # past the deck's own shapes
    assert deck.undo() and deck.to_bytes() == before


def test_z_steps_inside_a_group(deck):
    slide = deck.slides[0]
    a = slide.add_shape("rect", 0, 0, EMU, EMU)
    b = slide.add_shape("rect", EMU, 0, EMU, EMU)
    group = slide.group([a, b])
    deck.shape(a.id).bring_forward()
    assert [child.id for child in deck.shape(group.id).children] == [b.id, a.id]


# ------------------------------------------------------------------------------------------
# Groups
# ------------------------------------------------------------------------------------------


def test_a_shape_added_to_a_scaled_group_lands_in_child_space(deck):
    slide = deck.slides[0]
    a = slide.add_shape("rect", EMU, EMU, EMU, EMU)
    b = slide.add_shape("rect", 3 * EMU, 3 * EMU, EMU, EMU)
    group = slide.group([a, b])
    group.width = group.width * 2              # child space is now stretched two to one
    group = deck.shape(group.id)
    before = {child.id: child.slide_bounds for child in group.children}
    child = group.add_shape("ellipse", 2 * EMU, 4 * EMU, EMU, EMU, text="c")
    # Child units: two slide EMU across for every one, from the group's left edge.
    assert child.slide_bounds == (EMU + 2 * (2 * EMU - EMU), 4 * EMU, 2 * EMU, EMU)
    after = {c.id: c.slide_bounds for c in deck.shape(group.id).children if c.id in before}
    assert after == before                     # the re-fit moved nothing
    with pytest.raises(ValueError):
        deck.shape(group.id).add_table(1, 1, 0, 0, EMU, EMU)
    text = deck.shape(group.id).add_textbox(EMU, EMU, EMU, EMU, "t")
    connector = deck.shape(group.id).add_connector("straight", (a, 3), (b, 1))
    assert text.parent_group.id == group.id and connector.parent_group.id == group.id


# ------------------------------------------------------------------------------------------
# The full-state SVG
# ------------------------------------------------------------------------------------------


def test_new_shapes_carry_their_state_in_the_svg(deck):
    pytest.importorskip("pptx2svg")
    slide = deck.slides[0]
    a = slide.add_shape("roundRect", EMU, EMU, EMU, EMU, text="SVG", adjustments={"adj": 1})
    b = slide.add_shape("ellipse", 4 * EMU, 3 * EMU, EMU, EMU)
    connector = slide.add_connector("elbow", (a, 3), (b, 2))
    svg = FullStateSvg(slide.render_svg(full_state=True))
    assert svg.get(a.id, "geom") == "roundRect" and svg.get(a.id, "kind") == "shape"
    assert svg.get(connector.id, "kind") == "connector"
    assert svg.json(a.id, "text")["p"][0]["c"][0]["t"] == "SVG"
    # Edit the new shape and the connector through typed attributes.
    svg.set(a.id, "x", str(2 * EMU))
    svg.set_fill(b.id, "solid", scheme="accent2")
    report = slide.apply_svg(str(svg))
    assert set(report.edited) == {a.id, b.id}
    from test_connectors import assert_attached

    assert_attached(deck, connector.id)  # the move went through the API: it followed


def test_typed_attributes_apply_to_shapes_added_through_the_svg(deck):
    """A new shape's raw XML (here a copy of one made by the API) plus typed attributes
    that differ from it: the shape is added, then the typed edits go on top."""
    pytest.importorskip("pptx2svg")
    slide = deck.slides[0]
    model = slide.add_shape("star5", EMU, EMU, EMU, EMU, text="Model")
    svg = FullStateSvg(slide.render_svg(full_state=True))
    raw = svg.xml(model.id)
    group = svg.add(raw, identifier="from-svg")
    group.set(P + "x", str(5 * EMU))
    group.set(P + "fill", "solid")
    group.set(P + "fill-scheme", "accent6")
    group.set(P + "geom", "star7")
    group.set(P + "text", '{"p": [{"algn": "ctr", "c": [{"t": "Added", "b": true}]}]}')
    report = slide.apply_svg(str(svg))
    added = deck.shape(report.added["from-svg"])
    assert added.id != model.id and added.left == 5 * EMU and added.preset == "star7"
    assert added.fill.color.value == "accent6" and added.text == "Added"
    assert added.paragraphs[0].runs[0].bold is True
    assert deck.shape(model.id).left == EMU
    assert_valid(deck.to_bytes(), deck.to_bytes())


# ------------------------------------------------------------------------------------------
# Connections and adjust values in the full-state SVG
# ------------------------------------------------------------------------------------------


@pytest.fixture
def wired(deck):
    """Two shapes and an elbow connector between them, on the first slide."""
    pytest.importorskip("pptx2svg")
    slide = deck.slides[0]
    a = slide.add_shape("roundRect", EMU, EMU, EMU, EMU, adjustments={"adj": 25000})
    b = slide.add_shape("ellipse", 5 * EMU, 3 * EMU, EMU, EMU)
    c = slide.add_shape("rect", 2 * EMU, 4 * EMU, EMU, EMU)
    connector = slide.add_connector("elbow", (a, 3), (b, 2))
    return deck, slide, a, b, c, connector


def test_connections_and_adjust_values_are_typed_attributes(wired):
    deck, slide, a, b, c, connector = wired
    free = slide.add_connector("straight", (0, 0), (EMU, EMU))
    star = slide.add_shape("star5", EMU, EMU, EMU, EMU)
    star.adjustments["adj"] = 15000
    svg = FullStateSvg(slide.render_svg(full_state=True))
    assert svg.get(connector.id, "cxn-begin") == f"{a.id} 3"
    assert svg.get(connector.id, "cxn-end") == f"{b.id} 2"
    assert (svg.get(free.id, "cxn-begin"), svg.get(free.id, "cxn-end")) == ("none", "none")
    assert svg.get(a.id, "adj") == "adj=25000"
    assert svg.get(b.id, "adj") == ""
    assert svg.get(star.id, "adj") == "adj=15000 hf=105146 vf=110557"
    assert svg.get(connector.id, "adj") == "adj1=50000"
    assert svg.get(a.id, "cxn-begin") is None  # only connectors have ends


def test_a_connection_changed_in_the_svg_is_made_through_the_api(wired):
    from test_connectors import assert_attached

    deck, slide, a, b, c, connector = wired
    before = deck.to_bytes()
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(connector.id, "cxn-end", f"{c.id} 0")
    report = slide.apply_svg(str(svg))
    assert report.edited == {connector.id: ["cxn-end"]}
    again = deck.shape(connector.id)
    assert again.end_connection[0].id == c.id and again.end_connection[1] == 0
    assert again.begin_connection[0].id == a.id
    assert_attached(deck, connector.id)
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(connector.id, "cxn-begin", "none")
    slide.apply_svg(str(svg))
    assert deck.shape(connector.id).begin_connection is None
    deck.undo()
    deck.undo()
    assert deck.to_bytes() == before


def test_adjust_values_changed_in_the_svg_follow_the_api(wired):
    from test_connectors import assert_attached

    deck, slide, a, b, c, connector = wired
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(a.id, "adj", "adj=50000")
    svg.set(b.id, "adj", "")                                    # unchanged: nothing to do
    star = slide.add_shape("star5", 7 * EMU, EMU, EMU, EMU)
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(a.id, "adj", "adj=50000")
    svg.set(star.id, "adj", "hf=105146")                        # the default: no change
    report = slide.apply_svg(str(svg))
    assert report.edited == {a.id: ["adj"]}
    assert deck.shape(a.id).adjustments["adj"] == 50000
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(star.id, "adj", "adj=12000")
    slide.apply_svg(str(svg))
    assert deck.shape(star.id).adjustments.explicit() == {"adj": 12000, "hf": 105146,
                                                         "vf": 110557}
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(star.id, "adj", "")
    slide.apply_svg(str(svg))
    assert deck.shape(star.id).adjustments.explicit() == {}
    assert_attached(deck, connector.id)


def test_a_preset_changed_in_the_svg_ignores_the_old_presets_values(wired):
    deck, slide, a, b, c, connector = wired
    arrow = slide.add_shape("rightArrow", EMU, 5 * EMU, EMU, EMU,
                            adjustments={"adj1": 30000, "adj2": 70000})
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(arrow.id, "geom", "star5")                           # adj left as it was
    slide.apply_svg(str(svg))
    assert deck.shape(arrow.id).preset == "star5"
    assert deck.shape(arrow.id).adjustments.explicit() == {}
    svg = FullStateSvg(slide.render_svg(full_state=True))
    svg.set(arrow.id, "geom", "roundRect")
    svg.set(arrow.id, "adj", "adj=40000")                        # an opinion for the new one
    slide.apply_svg(str(svg))
    assert deck.shape(arrow.id).adjustments.explicit() == {"adj": 40000}


def test_a_connector_added_through_the_svg_attaches_to_a_new_shape(wired):
    """A new shape and a new connector between it and an existing one, both from raw XML:
    the connector names the new shape by its id in the SVG, is attached to it (whatever
    cNvPr id the shape ends up with), and is routed."""
    from svgedit import red_square
    from test_connectors import assert_attached

    deck, slide, a, b, c, connector = wired
    svg = FullStateSvg(slide.render_svg(full_state=True))
    taken = int(svg.get(a.id, "cnvpr-id"))
    svg.add(red_square(8 * EMU, 4 * EMU, EMU, shape_id=taken), identifier="new-box")
    link = svg.add(svg.xml(connector.id), identifier="new-link")
    link.set(P + "cxn-begin", "new-box 1")
    link.set(P + "cxn-end", f"{b.id} 0")
    report = slide.apply_svg(str(svg))
    box, made = deck.shape(report.added["new-box"]), deck.shape(report.added["new-link"])
    assert box.id != a.id
    assert made.begin_connection[0].id == box.id and made.begin_connection[1] == 1
    assert made.end_connection[0].id == b.id and made.end_connection[1] == 0
    assert_attached(deck, made.id)
    assert_attached(deck, connector.id)
    assert_valid(deck.to_bytes(), deck.to_bytes())


def test_a_shape_moved_in_its_raw_xml_takes_its_connectors_along(wired):
    from test_connectors import assert_attached

    deck, slide, a, b, c, connector = wired
    svg = FullStateSvg(slide.render_svg(full_state=True))
    raw = svg.xml(b.id)
    raw.find(".//{*}off").set("y", str(1 * EMU))
    svg.set_xml(b.id, raw)
    report = slide.apply_svg(str(svg))
    assert report.replaced == [b.id]
    assert deck.shape(b.id).top == EMU
    assert_attached(deck, connector.id)


def test_a_new_deck_carries_its_connections_through_the_svg():
    """On a deck made from nothing too, the SVG of a wired slide applies to a copy
    without a change."""
    pytest.importorskip("pptx2svg")
    document = Document.new()
    slide = document.add_slide("Blank")
    a = slide.add_shape("rect", EMU, EMU, EMU, EMU)
    b = slide.add_shape("ellipse", 5 * EMU, 3 * EMU, EMU, EMU)
    slide.add_connector("curved", (a, 2), (b, 4))
    saved = document.to_bytes()
    copy = Document.open(saved)
    assert not copy.apply_svg(Document.open(saved).slides[0].render_svg(full_state=True))
    assert copy.to_bytes() == saved
