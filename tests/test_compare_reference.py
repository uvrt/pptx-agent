"""Copies compared against the slide they were made from (production feedback round 17, P2):
``ppt_compare_slides``' ``reference`` and ``Document.compare_slides(reference=...)``.

The majority is the wrong yardstick for copies of a sample: when two of three copies drift
the same way, the third -- the one left alone -- is the odd one out.  With a reference its
values are the expected ones."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.errors import ArgumentError
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from test_compare_slides import _run

E = 12700


def _entry(data, name):
    return next(entry for entry in data["shapes"] if entry.get("name") == name)


def _sample(deck: Document):
    """A sample slide: a title, a named header box in accent1, a named note."""
    slide = deck.add_slide("Title Only")
    slide.title = "Sample record"
    header = slide.add_shape("rect", 40 * E, 100 * E, 400 * E, 40 * E, text="Header")
    header.name = "Header"
    header.fill = "accent1"
    note = slide.add_textbox(40 * E, 200 * E, 300 * E, 30 * E, "Note")
    note.name = "Note"
    return slide


def _copies():
    """A deck of samples and a report holding three copies of its sample, two of which
    drifted the same way: the header filled light grey and moved 1 pt down."""
    samples = Document.new()
    sample = _sample(samples)
    report = Document.new()
    copies = [report.import_slide(sample).slide for _ in range(3)]
    for copy in copies[1:]:
        header = next(shape for shape in copy.shapes if shape.name == "Header")
        header.fill = "#F2F2F2"
        header.move_by(dy=E)
    return samples, report, copies


def test_without_a_reference_the_majority_wins_and_the_untouched_copy_is_reported():
    _, report, copies = _copies()
    data = report.compare_slides(copies).to_json()
    fill = next(item for item in _entry(data, "Header")["differ"] if item["property"] == "fill")
    assert fill["value"] == "#F2F2F2"
    assert fill["other"][0]["shapes"] == [f"{copies[0].slide_id}.3"]


def test_with_a_reference_every_copy_that_drifted_is_reported_with_a_fix_toward_it():
    samples, report, copies = _copies()
    comparison = report.compare_slides(copies, reference=samples.slides[0])
    data = comparison.to_json()
    assert data["reference"] == f"s:{samples.slides[0].slide_id}"
    entry = _entry(data, "Header")
    assert entry["reference_shape"] == f"{samples.slides[0].slide_id}.3"
    drifted = [f"{copy.slide_id}.3" for copy in copies[1:]]
    by_property = {item["property"]: item for item in entry["differ"]}
    fill = by_property["fill"]
    assert fill["reference"] == fill["value"] and fill["reference"].startswith("#")
    assert fill["on"] == 1
    assert fill["other"] == [{"value": "#F2F2F2", "shapes": drifted}]
    assert fill["fix"] == {"tool": "ppt_set_shape",
                           "items": [{"target": target, "fill": "accent1"} for target in drifted]}
    y = by_property["y"]
    assert y["reference"] == 100 and y["near"] == [{"value": 101, "shapes": drifted}]
    assert y["fix"]["items"] == [{"target": target, "y": 100} for target in drifted]
    assert data["only"] == []
    assert all(entry.get("name") != "Note" for entry in data["shapes"])


def test_a_shape_missing_on_a_copy_and_one_not_on_the_reference_are_reported():
    samples, report, copies = _copies()
    next(shape for shape in copies[0].shapes if shape.name == "Note").delete()
    extra = copies[2].add_textbox(500 * E, 300 * E, 100 * E, 30 * E, "Extra")
    extra.name = "Extra"
    data = report.compare_slides(copies, reference=samples.slides[0]).to_json()
    assert _entry(data, "Note")["missing_on"] == [f"s:{copies[0].slide_id}"]
    assert data["only"] == [{"shape": extra.id, "name": "Extra", "by": "position",
                             "shapes": [extra.id]}]


def test_one_copy_is_enough_and_the_reference_is_not_among_the_slides():
    samples, report, copies = _copies()
    data = report.compare_slides(copies[:1], reference=samples.slides[0]).to_json()
    assert data["shapes"] == []                          # the untouched copy agrees
    with pytest.raises(ArgumentError, match="the reference"):
        report.compare_slides(copies, reference=copies[0])
    with pytest.raises(ArgumentError, match="at least two"):
        report.compare_slides(copies[:1])


def test_a_reference_in_the_same_deck_breaks_the_majority_of_a_planted_run():
    deck = _run(4, drift=(1, 2, 3))
    slides = deck.slides
    data = deck.compare_slides(slides[1:], reference=slides[0]).to_json()
    caption = _entry(data, "Caption")
    y = next(item for item in caption["differ"] if item["property"] == "y")
    assert y["reference"] == 450 and y["on"] == 0
    assert y["near"][0]["shapes"] == [f"{s.slide_id}.4" for s in slides[1:]]


# -- the tool ---------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def test_the_tool_compares_against_another_deck_s_slide_and_its_fixes_close_the_gap(toolbox):
    samples, report, copies = _copies()
    session = toolbox.session()
    d1 = session.open(report.to_bytes(), "report.pptx")
    d2 = session.open(samples.to_bytes(), "samples.pptx")
    ids = [f"s:{copy.slide_id}" for copy in copies]
    sample = f"s:{samples.slides[0].slide_id}"
    result = toolbox.dispatch(session, "ppt_compare_slides",
                              {"doc": d1, "slides": ids,
                               "reference": {"doc": d2, "slide": sample}})
    assert result.ok, result.error
    assert result.data["reference"] == f"{d2}/{sample}"
    assert "against" in result.summary and "not on the reference" in result.summary
    entry = _entry(result.data, "Header")
    assert entry["reference_shape"].startswith(f"{d2}/")
    for item in entry["differ"]:
        fix = item["fix"]
        assert toolbox.dispatch(session, fix["tool"], {"doc": d1, "items": fix["items"]}).ok
    again = toolbox.dispatch(session, "ppt_compare_slides",
                             {"doc": d1, "slides": ids,
                              "reference": {"doc": d2, "slide": sample}})
    assert again.ok and again.data["shapes"] == []
    # the reference in the compared deck: doc defaults to it
    inside = toolbox.dispatch(session, "ppt_compare_slides",
                              {"doc": d1, "slides": ids[1:], "reference": {"slide": ids[0]}})
    assert inside.ok and inside.data["reference"] == ids[0]
    among = toolbox.dispatch(session, "ppt_compare_slides",
                             {"doc": d1, "slides": ids, "reference": {"slide": ids[0]}})
    assert not among.ok and "reference" in among.error.message
    missing = toolbox.dispatch(session, "ppt_compare_slides",
                               {"doc": d1, "slides": ids,
                                "reference": {"doc": d2, "slide": "s:999"}})
    assert not missing.ok and missing.error.field == "reference.slide"
