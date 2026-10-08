"""What a layout offers, without making a slide to find out (the trial's finding 8): five
runs added a scratch slide per layout to see its placeholders."""

from __future__ import annotations

import dataclasses

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document, Layout, LayoutPlaceholder, LayoutShape, Placeholder

TEMPLATE = FIXTURE_DIR / "generated" / "trial" / "company-template.potx"


def test_placeholder_is_a_named_tuple_and_still_a_tuple():
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    title, body = slide.shapes
    assert title.placeholder == Placeholder("title", None) == ("title", None)
    assert body.placeholder.type is None and body.placeholder.idx == 1
    kind, idx = body.placeholder
    assert (kind, idx) == (None, 1)
    assert slide.add_textbox(0, 0, 914400, 914400, "x").placeholder is None


def test_layout_placeholders_are_what_a_slide_gets():
    deck = Document.new()
    layout = deck.layout("Two Content")
    offered = layout.placeholders
    assert all(isinstance(ph, LayoutPlaceholder) for ph in offered)
    assert [ph.type for ph in offered] == ["title", None, None, "dt", "ftr", "sldNum"]
    assert [ph.idx for ph in offered][:3] == [None, 1, 2]
    slide = deck.add_slide(layout)
    for shape in slide.shapes:      # inherits its geometry from the layout's placeholder
        match = next(ph for ph in offered if ph.placeholder == shape.placeholder)
        assert match.bounds == (shape.left, shape.top, shape.width, shape.height)
    left, right = offered[1], offered[2]
    assert left.bounds[0] < right.bounds[0]


def test_default_sizes_per_level_follow_the_master():
    deck = Document.new()
    title, body = deck.layout("Title and Content").placeholders[:2]
    assert title.sizes[0] == 44.0
    assert body.sizes[:5] == (28.0, 24.0, 20.0, 18.0, 18.0)
    assert len(body.sizes) == 9


def test_a_template_with_its_own_names():
    with pytest.warns(Warning):
        deck = Document.open(TEMPLATE)
    title_only = deck.layout("TITLE_ONLY")
    title = next(ph for ph in title_only.placeholders if ph.type == "title")
    assert title.bounds == (729450, 1318650, 7688400, 535200)
    assert title.bottom == 1318650 + 535200
    assert title.name == "Google Shape;45;p6"


def test_layout_shapes_are_read_only():
    deck = Document.new()
    layout = deck.layout("Title Slide")
    shapes = layout.shapes
    assert all(isinstance(shape, LayoutShape) for shape in shapes)
    assert [shape.placeholder.type for shape in shapes] == ["ctrTitle", "subTitle", "dt", "ftr",
                                                            "sldNum"]
    with pytest.raises(dataclasses.FrozenInstanceError):
        shapes[0].name = "x"
    with pytest.raises(AttributeError):
        layout.shapes = []
    assert isinstance(layout, Layout) and layout == deck.layouts[0]


def test_a_hand_made_layout_says_where_to_get_one():
    with pytest.raises(ValueError, match="deck.layouts"):
        Layout("x", "ppt/slideLayouts/slideLayout1.xml", "ppt/slideMasters/slideMaster1.xml") \
            .placeholders
