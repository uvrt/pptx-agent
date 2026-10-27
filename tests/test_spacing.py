"""Spacing measurements (round 12, Q1): the tightest gaps, padding and frame clearances on a
slide, as drawn, nearest first -- :meth:`Slide.spacing` and ``check``'s ``spacing``."""

from __future__ import annotations

import datetime as dt
import json

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from pptx_agent.tools.shared_tools import SPACING_PER_SLIDE

pytest.importorskip("pptx2svg")

PT = 12700
CLOCK = lambda: dt.datetime(2026, 10, 21, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def _summary(heading_top=239, chips_top=488, text_top=403):
    """A made-up summary slide: a block heading 4 pt under a filled panel; a card whose
    text is drawn 3 pt above two chips that sit 4 pt above the card's bottom edge.  The
    card's text box is taller than its four lines, so its box runs into the chips while
    its drawn text ends 3 pt above them."""
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    slide.title = "Quarter in brief"
    panel = slide.add_shape("rect", 72 * PT, 159 * PT, 516 * PT, 76 * PT, fill="#F4F4F3",
                            line="none")
    heading = slide.add_textbox(106 * PT, heading_top * PT, 470 * PT, 22 * PT,
                                "Block heading", autofit="none")
    heading.text_frame.insets = (0, 0, 0, 0)
    card = slide.add_shape("rect", 616 * PT, 378 * PT, 272 * PT, 132 * PT, fill="#F4F4F3",
                           line="none")
    text = slide.add_textbox(636 * PT, text_top * PT, 244 * PT, 99 * PT,
                             "Orchard deliveries rose in the third quarter while the cellar "
                             "stayed full and the presses ran late into autumn.",
                             autofit="none")
    text.text_frame.insets = (0, 0, 0, 0)
    chips = [slide.add_shape("roundRect", 636 * PT, chips_top * PT, 94 * PT, 18 * PT,
                             text="Apples", fill="accent1"),
             slide.add_shape("roundRect", 738 * PT, chips_top * PT, 76 * PT, 18 * PT,
                             text="Pears", fill="accent2")]
    for chip in chips:
        chip.text_frame.insets = (45720, 0, 45720, 0)
        chip.text_frame.paragraphs[0].runs[0].size = 10
    return deck, slide, (panel, heading, card, text, *chips)


def _entries(slide, **kwargs):
    return [entry.to_json() for entry in slide.spacing(**kwargs).entries]


def test_the_repro_reports_the_heading_gap_and_the_chip_spacings_with_their_numbers():
    deck, slide, (panel, heading, card, text, apples, pears) = _summary()
    report = slide.spacing()
    assert report.drawn and report.below == 8 * PT
    found = [entry.to_json() for entry in report.entries]
    # Nearest first: the chips 3 pt under the drawn text (4 lines; the box runs on), the
    # heading 4 pt under the panel, the chips 4 pt above the card's bottom edge.
    assert [(e.get("between"), e.get("inside"), e.get("shape")) for e in found] == [
        ([text.id, apples.id], None, None), ([text.id, pears.id], None, None),
        ([panel.id, heading.id], None, None), (None, card.id, apples.id)]
    assert found[0]["axis"] == found[1]["axis"] == found[2]["axis"] == "y"
    assert abs(found[0]["gap"] - 3) <= 0.5 and found[0]["gap"] == found[1]["gap"]
    assert found[2]["gap"] == 4.0
    assert found[3] == {"inside": card.id, "shape": apples.id, "edge": "bottom",
                        "padding": 4.0, "also": [pears.id]}
    # The text's box overlaps the chips' boxes; what is drawn does not, and check agrees.
    assert text.drawn_bounds[1] + text.drawn_bounds[3] > apples.drawn_bounds[1]
    assert slide.collisions() == []


def test_a_well_spaced_slide_reports_nothing_below_the_threshold():
    deck, slide, _ = _summary(heading_top=251, chips_top=483, text_top=390)
    assert _entries(slide) == []
    # The threshold is the caller's: at 20 pt the same slide's tightest come back.
    wider = _entries(slide, below=20 * PT)
    assert wider and all(e.get("gap", e.get("padding")) < 20 for e in wider)
    assert [e.get("gap", e.get("padding")) for e in wider] == sorted(
        e.get("gap", e.get("padding")) for e in wider)


def test_only_each_shapes_nearest_neighbour_and_overlaps_are_left_out():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    a = slide.add_shape("rect", 100 * PT, 100 * PT, 100 * PT, 40 * PT, fill="accent1")
    b = slide.add_shape("rect", 100 * PT, 142 * PT, 100 * PT, 40 * PT, fill="accent1")
    c = slide.add_shape("rect", 100 * PT, 186 * PT, 100 * PT, 40 * PT, fill="accent1")
    d = slide.add_shape("rect", 205 * PT, 100 * PT, 60 * PT, 40 * PT, fill="accent2")
    slide.add_shape("rect", 400 * PT, 100 * PT, 60 * PT, 60 * PT, fill="accent3")
    slide.add_shape("rect", 430 * PT, 130 * PT, 60 * PT, 60 * PT, fill="accent4")  # overlaps
    found = _entries(slide)
    assert found == [{"between": [a.id, b.id], "axis": "y", "gap": 2.0},
                     {"between": [b.id, c.id], "axis": "y", "gap": 4.0},
                     {"between": [a.id, d.id], "axis": "x", "gap": 5.0}]


def test_backgrounds_lines_and_empty_boxes_are_left_out():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    width, height = deck.slide_size
    slide.add_shape("rect", 0, 0, width, height, fill="#FAFAF8", line="none")
    box = slide.add_shape("rect", 2 * PT, 2 * PT, 100 * PT, 40 * PT, fill="accent1")
    slide.add_connector("straight", (2 * PT, 45 * PT), (102 * PT, 45 * PT))
    slide.add_textbox(2 * PT, 44 * PT, 100 * PT, 20 * PT, "", autofit="none")   # empty
    assert _entries(slide) == []
    other = slide.add_shape("rect", 2 * PT, 46 * PT, 100 * PT, 40 * PT, fill="accent2")
    assert _entries(slide) == [{"between": [box.id, other.id], "axis": "y", "gap": 4.0}]


def test_rotated_and_grouped_shapes_are_measured_where_they_are_drawn():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    # A 100 x 20 bar turned a quarter: drawn 20 wide, 100 tall, centred where it was.
    bar = slide.add_shape("rect", 100 * PT, 200 * PT, 100 * PT, 20 * PT, fill="accent1")
    bar.rotation = 90
    right = slide.add_shape("rect", 163 * PT, 180 * PT, 40 * PT, 40 * PT, fill="accent2")
    assert _entries(slide) == [{"between": [bar.id, right.id], "axis": "x", "gap": 3.0}]
    first = slide.add_shape("rect", 300 * PT, 300 * PT, 50 * PT, 30 * PT, fill="accent3")
    second = slide.add_shape("rect", 356 * PT, 300 * PT, 50 * PT, 30 * PT, fill="accent3")
    group = slide.group([first, second])
    members = [shape.id for shape in group.children]
    assert {"between": members, "axis": "x", "gap": 6.0} in _entries(slide)


def test_clearance_to_the_footer_and_past_the_content_area():
    deck = Document.new()
    deck.add_slide("Title and Content").title = "Frame"
    deck.set_footer("Orchard review", slide_number=True)
    slide = deck.slides[0]
    footer = next(s for s in slide.shapes if s.placeholder and s.placeholder.type == "ftr")
    area = slide.content_area
    top = footer.drawn_bounds[1] - 5.5 * PT - 15 * PT
    note = slide.add_shape("rect", 330 * PT, round(top), 200 * PT, 15 * PT, fill="accent1")
    found = _entries(slide)
    past = round((top + 15 * PT - (area[1] + area[3])) / PT, 1)
    assert {"clearance": "footer", "shape": note.id, "gap": 5.5} in found
    assert {"clearance": "content_area", "shape": note.id, "edge": "bottom",
            "gap": -past} in found
    assert found[0]["gap"] == -past                           # nearest first: past the area
    # The body placeholder where its layout puts it is the area itself: not measured.
    assert all(e.get("shape") != slide.shapes[1].id for e in found)


# -- check's spacing ----------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _check(toolbox, deck, **arguments):
    session = toolbox.session(clock=CLOCK)
    session.open(deck.to_bytes(), "summary.pptx")
    result = toolbox.dispatch(session, "check", {"doc": "d1", **arguments})
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def test_check_reports_the_spacings_on_request(toolbox):
    deck, slide, (panel, heading, card, text, apples, pears) = _summary()
    result = _check(toolbox, deck, slides=[1], spacing=8)
    data = result.data
    assert data["spacing"] == _entries(slide)
    assert data["spacing"][2] == {"between": [panel.id, heading.id], "axis": "y", "gap": 4.0}
    assert "Cite the gap" in data["spacing_note"]
    assert result.summary.endswith("4 spacing(s) below 8 pt")
    assert data["collisions"] == []
    # Without spacing, none is measured; the threshold is the caller's.
    assert "spacing" not in _check(toolbox, deck).data
    assert _check(toolbox, deck, spacing=3.5).data["spacing"] == data["spacing"][:2]


def test_check_lists_a_slides_nearest_and_counts_the_rest(toolbox):
    deck = Document.new()
    slide = deck.add_slide("Blank")
    for row in range(4):
        for column in range(5):
            slide.add_shape("rect", (60 + 52 * column) * PT, (60 + 32 * row) * PT, 50 * PT,
                            30 * PT, fill="accent1")
    data = _check(toolbox, deck, spacing=8).data
    total = len(slide.spacing().entries)
    assert total > SPACING_PER_SLIDE
    assert len(data["spacing"]) == SPACING_PER_SLIDE
    assert data["spacing_left_out"] == total - SPACING_PER_SLIDE


def test_a_decks_check_refuses_a_documents_layout_measures(toolbox):
    """Round 16, F8: ``min_size`` and ``styles`` measure a document's layout (ooxml-edit refuses
    them for a deck, whose handler takes neither).  ``layout`` in include runs for a deck
    since round 18 (Q4, empty paragraphs between text): none here."""
    deck, _, _ = _summary()
    session = toolbox.session(clock=CLOCK)
    session.open(deck.to_bytes(), "summary.pptx")
    for name, value in (("min_size", 9), ("styles", ["Normal"])):
        refused = toolbox.dispatch(session, "check", {"doc": "d1", name: value})
        assert not refused.ok and refused.error.code == "invalid_arguments" and refused.error.field == name
    noted = _check(toolbox, deck, include=["fit", "layout"]).data
    assert "notes" not in noted and noted["layout"] == {"empty_paragraphs": []}
