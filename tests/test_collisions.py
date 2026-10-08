"""Collisions: text over text, lines through text (the trial's N3).

p8's agents wrote their own pairwise box tests and found three milestone labels on top of
each other; p4's elbow ran through a box and p8 run 2's dashed line over the bars' text,
and nothing in the library said so.  ``overflows()`` reported only text fit, placeholders
and the slide's edge.
"""

from __future__ import annotations

import pytest

from conftest import fixture_paths
from pptx_agent import Document, Overflow

pytest.importorskip("pptx2svg")

EMU = 914400


def _slide():
    deck = Document.new()
    return deck, deck.add_slide("Blank")


def _overlaps(slide):
    return {(o.detail, o.shape, o.other) for o in slide.collisions()}


def test_two_labels_on_top_of_each_other_collide():
    deck, slide = _slide()
    first = slide.add_textbox(EMU, EMU, 2 * EMU, EMU // 2, "Kick-off")
    second = slide.add_textbox(EMU + EMU // 4, EMU, 2 * EMU, EMU // 2, "SteerCo 1")
    (problem,) = slide.collisions()
    assert isinstance(problem, Overflow)
    assert (problem.kind, problem.detail) == ("overlap", "text")
    assert {problem.shape, problem.other} == {first.id, second.id} and problem.amount > 0
    assert problem in deck.overflows()


def test_labels_whose_boxes_touch_but_text_does_not_are_fine():
    """A text box with no fill and no outline shows only its text: two boxes that overlap
    where neither draws anything are not a collision."""
    deck, slide = _slide()
    slide.add_textbox(EMU, EMU, 3 * EMU, 2 * EMU, "Top label")          # text at the top
    slide.add_textbox(EMU, 2 * EMU, 3 * EMU, EMU // 2, "Below it")
    assert slide.collisions() == []


def test_a_label_on_a_bar_is_layering_not_a_collision():
    deck, slide = _slide()
    bar = slide.add_shape("rect", EMU, EMU, 4 * EMU, EMU // 2, text="Savings tracking")
    slide.add_textbox(EMU + EMU // 4, EMU, EMU, EMU // 2, "Q1")
    assert slide.collisions() == []
    bar.bring_to_front()                       # now the bar hides the label
    assert {detail for detail, _, _ in _overlaps(slide)} == {"text"}


def test_a_line_over_text_is_caught_and_a_line_behind_a_bar_is_not():
    """p8 run 2: the Board-update line was drawn over the bars' text; run 1 sent it back."""
    deck, slide = _slide()
    bar = slide.add_shape("rect", EMU, 2 * EMU, 4 * EMU, EMU // 2, text="Savings tracking")
    line = slide.add_connector("straight", (3 * EMU, EMU), (3 * EMU, 4 * EMU),
                               line={"dash": "dash"})
    (problem,) = slide.collisions()
    assert (problem.detail, problem.shape, problem.other) == ("line", line.id, bar.id)
    assert abs(problem.amount - (EMU // 2 - 2 * 45720)) <= 2      # through the text area
    assert "crosses the text of" in str(problem)
    line.send_to_back()
    assert slide.collisions() == []


def test_a_line_through_a_label_with_no_fill_is_caught_whatever_the_order():
    deck, slide = _slide()
    line = slide.add_connector("straight", (EMU, 2 * EMU), (5 * EMU, 2 * EMU))
    label = slide.add_textbox(2 * EMU, 2 * EMU - EMU // 4, EMU, EMU // 2, "Pass")
    label.send_to_back()
    assert _overlaps(slide) == {("line", line.id, label.id)}


def test_an_elbow_through_a_box_is_caught():
    """p4's first layout: pick-and-pack's bottom to ship-order's top, through the box
    between them."""
    deck, slide = _slide()
    pick = slide.add_shape("rect", 8 * EMU, EMU, 2 * EMU, EMU, text="Pick and pack")
    notify = slide.add_shape("rect", 4 * EMU, 2 * EMU + EMU // 4, 3 * EMU, EMU,
                             text="Notify customer")
    ship = slide.add_shape("rect", EMU, 4 * EMU, 2 * EMU, EMU, text="Ship order")
    elbow = slide.add_connector("elbow", (pick, "bottom"), (ship, "top"))
    assert ("line", elbow.id, notify.id) in _overlaps(slide)
    notify.top = 3 * EMU + EMU // 2 - EMU // 4 + EMU // 2   # below the elbow's run
    assert all(o.other != notify.id for o in slide.collisions())


def test_a_connector_ending_on_a_box_does_not_cross_it():
    deck, slide = _slide()
    a = slide.add_shape("rect", EMU, EMU, 2 * EMU, EMU, text="Order received")
    b = slide.add_shape("rect", 5 * EMU, EMU, 2 * EMU, EMU, text="Credit check")
    slide.add_connector("straight", (a, "right"), (b, "left"))
    slide.add_connector("elbow", (a, "bottom"), (b, "bottom"))
    assert slide.collisions() == []


def test_a_group_is_looked_into_not_compared_with_its_members():
    deck, slide = _slide()
    bar = slide.add_shape("rect", EMU, EMU, 4 * EMU, EMU // 2, text="Bar")
    label = slide.add_textbox(EMU + EMU // 4, EMU, EMU, EMU // 2, "Q1")
    slide.group([bar, label])
    assert slide.collisions() == []
    other = slide.add_textbox(EMU + EMU // 4, EMU, EMU, EMU // 2, "Q2")
    assert {problem.other for problem in slide.collisions()} | \
        {problem.shape for problem in slide.collisions()} >= {other.id}


@pytest.mark.parametrize("path", fixture_paths(), ids=lambda path: path.stem)
def test_the_real_decks_have_no_collisions(path):
    """Before these rules, overflows() reported 6 placeholder overlaps on the basic-theme
    deck, none of them visible; the corpus is laid out cleanly, so any report here is a
    false positive."""
    deck = Document.open(path)
    assert [str(o) for o in deck.overflows() if o.kind == "overlap"] == []


def test_box_mode_reports_label_boxes_that_overlap_where_their_text_does_not():
    from pptx_agent import Document, Pt

    deck = Document.new()
    slide = deck.add_slide("Blank")
    first = slide.add_textbox(Pt(100), Pt(160), Pt(94), Pt(29), "Kick-off", autofit="none")
    second = slide.add_textbox(Pt(191), Pt(160), Pt(94), Pt(29), "SteerCo 1", autofit="none")
    assert slide.collisions() == []
    (problem,) = slide.collisions(boxes=True)
    assert (problem.detail, problem.shape, problem.other) == ("box", second.id, first.id)
    assert [p.detail for p in deck.overflows(boxes=True)] == ["box"]
    second.left = Pt(200)
    assert slide.collisions(boxes=True) == []
