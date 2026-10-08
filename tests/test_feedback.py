"""Per-edit layout facts: near-alignment, uneven gaps, outlier text sizes and far labels, each
with the change that resolves it -- and silence where a difference has a reason."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.feedback import layout_facts
from pptx_agent.edit.fit import _flatten
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

PT = 12700
FIXTURES = Path(__file__).parent / "fixtures"


def _box(slide, x, y, w, h, text, preset="rect"):
    shape = slide.add_shape(preset, int(x * PT), int(y * PT), int(w * PT), int(h * PT))
    shape.text = text
    return shape


@pytest.fixture
def slide():
    deck = Document.new()
    page = deck.add_slide("Title Only")
    page.title = "Plan"
    return page


def test_a_box_just_off_its_rows_line(slide):
    boxes = [_box(slide, x, y, 150, 60, t) for x, y, t in
             ((60, 150, "A"), (230, 150, "B"), (400, 150, "C"), (570, 152, "D"))]
    (fact,) = layout_facts(slide, [boxes[3].id])
    assert fact.kind == "near_alignment" and fact.fix == {"y": 150.0}
    assert fact.to_json()["fix"] == f"ppt_set_shape {boxes[3].id} y=150"
    assert layout_facts(slide, [boxes[0].id]) == []          # only what the edit touched


def test_one_gap_unlike_the_others(slide):
    boxes = [_box(slide, x, 150, 150, 60, t) for x, t in
             ((60, "A"), (230, "B"), (400, "C"), (574, "D"))]
    (fact,) = layout_facts(slide, [boxes[3].id])
    assert fact.kind == "uneven_gap" and fact.fix == {"x": 570.0}
    three = [_box(slide, x, 300, 150, 60, t) for x, t in ((60, "E"), (230, "F"), (406, "G"))]
    (fact,) = layout_facts(slide, [three[2].id])
    assert fact.kind == "uneven_gap" and fact.fix == {"x": 400.0}


def test_a_gap_on_purpose_and_data_positions_say_nothing(slide):
    wide = [_box(slide, x, 150, 100, 40, t) for x, t in
            ((60, "A"), (180, "B"), (300, "C"), (560, "D"))]          # a 140 pt gap: a choice
    assert layout_facts(slide, [b.id for b in wide]) == []
    dots = [_box(slide, x, 400, 20, 20, "", preset="ellipse") for x in (100, 147, 230)]
    assert layout_facts(slide, [d.id for d in dots]) == []           # no text: data points


def test_a_size_unlike_its_like_boxes(slide):
    boxes = [_box(slide, x, 150, 150, 60, t) for x, t in ((60, "A"), (230, "B"), (400, "C"))]
    boxes[2].text_frame.paragraphs[0].runs[0].size = 24
    assert layout_facts(slide, [boxes[2].id]) == []        # set where the others inherit
    for box in boxes[:2]:
        box.text_frame.paragraphs[0].runs[0].size = 14
    (fact,) = layout_facts(slide, [boxes[2].id])
    assert fact.kind == "text_size" and fact.fix == {"size": 14.0}
    assert fact.to_json()["fix"] == f"ppt_format_text {boxes[2].id} size=14"


def test_a_label_far_from_its_marker(slide):
    for n, x in enumerate((100, 250, 400, 550)):
        _box(slide, x, 300, 20, 20, "", preset="ellipse")
        label = slide.add_textbox(int((x + 24) * PT), int(302 * PT), 80 * PT, 16 * PT)
        label.text = f"Label {n}"
    far = slide.add_textbox(int(724 * PT), int(302 * PT), 80 * PT, 16 * PT)
    far.text = "Label 4"
    _box(slide, 700, 300, 20, 20, "", preset="ellipse")
    assert layout_facts(slide, [far.id]) == []                      # 4 pt: like the others
    far.left = int(760 * PT)
    (fact,) = layout_facts(slide, [far.id])
    assert fact.kind == "label_distance" and fact.fix["x"] < 760


def test_quiet_on_decks_people_made():
    for path in sorted(FIXTURES.glob("*.pptx")):
        deck = Document.open(path)
        for page in deck.slides:
            assert layout_facts(page, [s.id for s in _flatten(page.shapes)]) == [], path.name


def test_the_tools_return_the_facts_with_their_fixes():
    deck = Document.new()
    deck.add_slide("Title Only").title = "Plan"
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: dt.datetime(2026, 10, 7, tzinfo=dt.timezone.utc))
        session.open(deck.to_bytes(), "plan.pptx")
        made = toolbox.dispatch(session, "ppt_add_shape", {"doc": "d1", "slide": "s:256", "items": [
            {"preset": "rect", "box": {"x": x, "y": y, "w": 150, "h": 60}, "text": t}
            for x, y, t in ((60, 150, "A"), (230, 150, "B"), (400, 150, "C"),
                            (570, 152, "D"))]})
        assert made.checks["layout"] == [{
            "kind": "near_alignment",
            "fact": "256.6: top edge 2.0 pt below the line 256.3, 256.4, 256.5 are on",
            "fix": "ppt_set_shape 256.6 y=150"}]
        fixed = toolbox.dispatch(session, "ppt_set_shape",
                                 {"doc": "d1", "items": [{"target": "256.6", "y": 150}]})
        assert "layout" not in fixed.checks
        sized = toolbox.dispatch(session, "ppt_format_text", {"doc": "d1", "items": [
            {"target": "256.3", "size": 14}, {"target": "256.4", "size": 14},
            {"target": "256.5", "size": 16}, {"target": "256.6", "size": 14}]})
        assert sized.checks["layout"] == [{
            "kind": "text_size", "fact": "256.5: text 16 pt where its 3 like shapes have 14 pt",
            "fix": "ppt_format_text 256.5 size=14"}]
        notes = toolbox.dispatch(session, "ppt_set_text",
                                 {"doc": "d1", "items": [{"target": "256/notes", "text": "x"}]})
        assert "layout" not in notes.checks
