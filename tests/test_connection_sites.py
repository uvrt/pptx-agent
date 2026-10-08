"""Connection sites by side, the documented site order, and start/end arrowheads (the
trial's finding 12: both P4 runs reverse-engineered the order and the head/tail naming)."""

from __future__ import annotations

import pytest

from pptx_agent import Document


def _slide():
    deck = Document.new()
    return deck, deck.add_slide("Blank")


#: The order the docstring of Shape.connection_sites states.
ORDER = {
    "rect": ["top", "left", "bottom", "right"],
    "roundRect": ["top", "left", "bottom", "right"],
    "diamond": ["top", "left", "bottom", "right"],
    "flowChartProcess": ["top", "left", "bottom", "right"],
    "flowChartDecision": ["top", "left", "bottom", "right"],
    "flowChartTerminator": ["top", "left", "bottom", "right"],
    "rightArrow": ["top", "left", "bottom", "right"],
    "trapezoid": ["top", "left", "bottom", "right"],
}


@pytest.mark.parametrize("preset", sorted(ORDER))
def test_the_documented_order_holds(preset):
    _, slide = _slide()
    shape = slide.add_shape(preset, 914400, 914400, 1828800, 914400)
    assert [shape.connection_site(side)[1] for side in ORDER[preset]] == [0, 1, 2, 3]


def test_the_ellipse_goes_round_from_the_top():
    _, slide = _slide()
    ellipse = slide.add_shape("ellipse", 0, 0, 1000000, 1000000)
    sites = ellipse.connection_sites
    assert len(sites) == 8 and sites[0] == (500000, 0) and sites[2] == (0, 500000)
    assert sites[4] == (500000, 1000000) and sites[6] == (1000000, 500000)
    assert [ellipse.connection_site(s)[1] for s in ("top", "left", "bottom", "right")] == \
        [0, 2, 4, 6]


def test_a_side_is_as_drawn_and_names_work_in_add_connector():
    deck, slide = _slide()
    box = slide.add_shape("rect", 914400, 914400, 1828800, 914400)
    goal = slide.add_shape("ellipse", 5486400, 914400, 1371600, 1371600)
    assert box.connection_site("right") == (box, 3)
    assert box.connection_site(2) == (box, 2)
    arrow = slide.add_connector("elbow", box.connection_site("right"), (goal, "left"),
                                line={"end": "triangle"})
    assert arrow.begin_connection[1] == 3 and arrow.end_connection[1] == 2
    assert arrow.line.tail.type == "triangle" and arrow.line.end == arrow.line.tail
    box.rotation = 90                               # its right is now at the bottom
    assert box.connection_site("bottom")[1] == 3
    with pytest.raises(ValueError):
        box.connection_site("middle")
    with pytest.raises(IndexError):
        box.connection_site(4)
    assert deck.validate() == []


def test_start_and_end_are_head_and_tail():
    _, slide = _slide()
    line = slide.add_connector("straight", (0, 0), (914400, 0)).line
    line.start = "oval"
    line.end = "triangle"
    assert line.head.type == "oval" and line.tail.type == "triangle"
    assert line.start == line.head and line.end == line.tail
    assert line.set_arrowhead("start", "diamond") is line
    assert line.head.type == "diamond"
    line.set_arrowhead("end", None)
    assert line.tail is None
    with pytest.raises(ValueError):
        line.set_arrowhead("middle", "oval")
