"""A run of like slides compared on its text's layout (round 14, S1): ``ppt_compare_slides``
lists where a matched shape's paragraphs (alignment, line spacing, space before and after,
bullet -- per level) or its frame (anchor, insets, autofit, text direction) differ from the
majority, as drawn after inheritance, each difference with the ``ppt_format_text`` items
that give the majority's value where that tool sets it.  Every deck here is made up.
"""

from __future__ import annotations

import json

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.oxml.xml import find, qn
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, EMU


def _run(count=3):
    """``count`` like slides on one layout: a title, a heading text box and a two-line text
    box, both named."""
    deck = Document.new(size="16:9")
    for number in range(count):
        slide = deck.add_slide("Title Only")
        slide.title = f"Orchard {number + 1}"
        slide.add_textbox(60 * EMU, 120 * EMU, 300 * EMU, 40 * EMU,
                          "Pear varieties").name = "Heading"
        slide.add_textbox(60 * EMU, 170 * EMU, 300 * EMU, 200 * EMU,
                          "Conference\nComice").name = "Notes"
    return deck


def _session(deck):
    toolbox = Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


def _ids(deck):
    return [f"s:{slide.slide_id}" for slide in deck.slides]


def _compare(toolbox, session, doc, deck):
    result = toolbox.dispatch(session, "ppt_compare_slides", {"doc": doc, "slides": _ids(deck)})
    assert result.ok, json.dumps(result.to_json())
    return result.data


def _apply(toolbox, session, doc, data):
    """Every fix of a comparison, dispatched: each must be valid against the tool's strict
    schema."""
    for entry in data["shapes"]:
        for difference in entry.get("differ", ()):
            fix = difference.get("fix")
            if fix is None:
                continue
            applied = toolbox.dispatch(session, fix["tool"], {"doc": doc, "items": fix["items"]})
            assert applied.ok, json.dumps(applied.to_json())


def _named(deck, index, name):
    return next(shape for shape in deck.slides[index].shapes if shape.name == name)


def test_a_heading_centred_on_one_slide_is_reported_with_its_fix():
    deck = _run(3)
    toolbox, session, doc = _session(deck)
    heading = _named(deck, 1, "Heading").id
    assert _compare(toolbox, session, doc, deck)["shapes"] == []
    toolbox.dispatch(session, "ppt_format_text",
                     {"doc": doc, "items": [{"target": heading, "align": "center"}]})
    data = _compare(toolbox, session, doc, deck)
    assert data["shapes"] == [{
        "shape": _named(deck, 0, "Heading").id, "name": "Heading",
        "differ": [{"property": "align", "value": "left", "on": 2,
                    "other": [{"value": "center", "shapes": [heading]}],
                    "fix": {"tool": "ppt_format_text",
                            "items": [{"target": heading, "align": "left"}]}}]}]
    _apply(toolbox, session, doc, data)
    assert _compare(toolbox, session, doc, deck)["shapes"] == []


def test_every_layout_property_differs_and_the_fixes_bring_it_back():
    deck = _run(5)
    for index in (3, 4):
        frame = _named(deck, index, "Notes").text_frame
        for paragraph in frame.paragraphs:
            paragraph.alignment = "right"
            paragraph.line_spacing = 1.5
            paragraph.space_before = 76200                      # 6 pt
            paragraph.space_after = 38100                       # 3 pt
            paragraph.set_numbering()
        frame.anchor = "middle"
        frame.insets = (0, 0, 0, 0)
        frame.autofit = "normal"                                # a text box's is "shape"
        frame._set_frame_attributes({"vert": "vert270"})
    toolbox, session, doc = _session(deck)
    data = _compare(toolbox, session, doc, deck)
    (entry,) = data["shapes"]
    assert entry["name"] == "Notes"
    by = {difference["property"]: difference for difference in entry["differ"]}
    assert [(name, by[name]["value"], by[name]["other"][0]["value"]) for name in by] == [
        ("align", "left", "right"), ("line_spacing", 1, 1.5), ("space_before", 0, 6),
        ("space_after", 0, 3), ("bullet", "none", "autonum arabicPeriod"),
        ("anchor", "top", "middle"), ("autofit", "shape", "normal"),
        ("direction", "horz", "vert270"), ("insets", [7.2, 3.6, 7.2, 3.6], [0, 0, 0, 0])]
    changed = [_named(deck, index, "Notes").id for index in (3, 4)]
    assert by["space_before"]["fix"]["items"] == [{"target": address, "space_before": 0}
                                                  for address in changed]
    assert by["insets"]["fix"]["items"][0] == {
        "target": changed[0], "insets": {"left": 7.2, "top": 3.6, "right": 7.2, "bottom": 3.6}}
    assert "fix" not in by["direction"]                         # no tool sets it
    _apply(toolbox, session, doc, data)
    left = _compare(toolbox, session, doc, deck)["shapes"]
    assert [d["property"] for d in left[0]["differ"]] == ["direction"]


def test_values_equal_after_inheritance_do_not_differ():
    deck = _run(3)
    frame = _named(deck, 1, "Notes").text_frame
    for paragraph in frame.paragraphs:
        paragraph.alignment = "left"                            # stated, as inherited
        paragraph.line_spacing = 1.0
        paragraph.space_before = 0
        paragraph.set_no_bullet()
    frame.anchor = "top"
    frame.insets = (91440, 45720, 91440, 45720)
    frame.autofit = "shape"
    frame._set_frame_attributes({"vert": "horz"})
    title = deck.slides[2].shapes[0].text_frame
    title.paragraphs[0].alignment = "left"
    title.anchor = deck.slides[0].shapes[0].text_frame.anchor  # the layout's, stated
    assert deck.compare_slides(deck.slides).to_json()["shapes"] == []


def test_a_placeholder_against_its_layout_and_master():
    deck = _run(3)
    layout = deck.layout("Title Only")
    root = deck.package.tree(layout.part_path)
    for shape in find(root, "p:cSld/p:spTree"):
        ph = find(shape, "p:nvSpPr/p:nvPr/p:ph")
        if ph is not None and ph.get("type") == "title":
            styles = shape.find(qn("p:txBody")).find(qn("a:lstStyle"))
            for child in list(styles):
                styles.remove(child)
            styles.append(styles.makeelement(qn("a:lvl1pPr"), {"algn": "ctr"}))
    deck.package.mark_dirty(layout.part_path)
    deck.slides[0].shapes[0].text_frame.paragraphs[0].alignment = "center"   # as the layout
    deck.slides[2].shapes[0].text_frame.paragraphs[0].alignment = "left"
    data = deck.compare_slides(deck.slides).to_json()
    title = data["shapes"][0]
    assert title["name"] == "Title 1"
    assert title["differ"] == [{
        "property": "align", "value": "center", "on": 2,
        "other": [{"value": "left", "shapes": [deck.slides[2].shapes[0].id]}],
        "fix": {"tool": "ppt_format_text",
                "items": [{"target": deck.slides[2].shapes[0].id, "align": "center"}]}}]


def test_paragraphs_that_differ_within_a_shape_are_mixed_and_a_level_on_its_own():
    deck = _run(3)
    for slide in deck.slides:
        frame = next(shape for shape in slide.shapes if shape.name == "Notes").text_frame
        frame.add_paragraph("Williams")
        frame.paragraphs[2].level = 1
    notes = _named(deck, 2, "Notes")
    notes.text_frame.paragraphs[1].alignment = "center"          # level 0 now mixed
    deeper = _named(deck, 1, "Notes").text_frame.paragraphs[2]
    deeper.space_after = 127000                                  # 10 pt, level 1 only
    data = deck.compare_slides(deck.slides).to_json()
    differ = data["shapes"][0]["differ"]
    assert [(d["property"], d.get("level"), d["other"][0]["value"]) for d in differ] == [
        ("align", None, "mixed"), ("space_after", 1, 10)]
    assert differ[0]["fix"]["items"] == [
        {"target": p.address, "align": "left"} for p in notes.text_frame.paragraphs[:2]]
    assert differ[1]["fix"]["items"] == [{"target": deeper.address, "space_after": 0}]
    toolbox, session, doc = _session(deck)
    data = _compare(toolbox, session, doc, deck)
    _apply(toolbox, session, doc, data)
    assert _compare(toolbox, session, doc, deck)["shapes"] == []


def test_a_mixed_majority_and_values_no_tool_writes_carry_no_fix():
    deck = _run(4)
    for index in (0, 1, 2):
        _named(deck, index, "Notes").text_frame.paragraphs[1].alignment = "center"
    odd = _named(deck, 3, "Notes").text_frame
    for paragraph in odd.paragraphs:
        paragraph.line_spacing_points = 18
        paragraph.set_bullet("–")
    for index in (0, 1, 2):
        frame = _named(deck, index, "Heading").text_frame
        frame.paragraphs[0].alignment = "distributed"
        frame.anchor = "distributed"
    data = deck.compare_slides(deck.slides).to_json()
    notes = next(entry for entry in data["shapes"] if entry["name"] == "Notes")
    heading = next(entry for entry in data["shapes"] if entry["name"] == "Heading")
    by = {d["property"]: d for d in notes["differ"]}
    assert by["align"]["value"] == "mixed" and "fix" not in by["align"]
    assert by["line_spacing"]["other"] == [{"value": "18pt", "shapes": [odd.address]}]
    assert by["bullet"]["other"][0]["value"] == "char –"
    assert by["line_spacing"]["fix"]["items"] == [{"target": odd.address, "line_spacing": 1}]
    assert by["bullet"]["fix"]["items"] == [{"target": odd.address, "bullet": "none"}]
    assert {d["property"]: d["value"] for d in heading["differ"]} == {
        "align": "distributed", "anchor": "distributed"}
    assert all("fix" not in d for d in heading["differ"])    # ppt_format_text has no such


def test_shapes_without_text_have_no_layout():
    deck = _run(3)
    for index, slide in enumerate(deck.slides):
        box = slide.add_textbox(500 * EMU, 120 * EMU, 100 * EMU, 40 * EMU, "")
        box.name = "Empty"
        if index == 2:
            box.text_frame.anchor = "bottom"
            box.text_frame.paragraphs[0].alignment = "right"
    assert deck.compare_slides(deck.slides).to_json()["shapes"] == []
