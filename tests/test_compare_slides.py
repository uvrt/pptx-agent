"""A run of like slides compared shape by shape (round 7, L3): ``ppt_compare_slides`` and
:meth:`Document.compare_slides` match corresponding shapes -- placeholders by type and idx,
then unique names, then kind and nearest box -- and list, per property, the majority's
value and the shapes that differ, near-misses (with a one-call fix) apart from larger
differences.  Measurement only.  Every deck here is made up.
"""

from __future__ import annotations

import json

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.compare import match
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, EMU
from test_size_overrides import _layout_title_size


def _run(count=9, *, drift=(), sizes=(), missing=(), only=(), reorder=(), recolour=(),
         moved=()):
    """``count`` like slides on a layout whose title is 28 pt: a title, a body, a named
    caption box, an unnamed marker square and an unnamed picture-like rectangle.  The
    keyword arguments plant differences by slide position (0-based): ``drift`` the caption
    1 pt lower, ``sizes`` the title set to 26 pt on the slide, ``missing`` no caption,
    ``only`` an extra callout, ``reorder`` the unnamed shapes added in the other order and
    named differently, ``recolour`` the marker a shade off, ``moved`` the marker far
    away."""
    deck = Document.new(size="16:9")
    _layout_title_size(deck, "Title and Content", 2800)
    for number in range(count):
        slide = deck.add_slide("Title and Content")
        slide.title = f"Harbour {number + 1}"
        slide.shapes[1].set_text("Moorings\nTide tables")
        if number in sizes:
            slide.shapes[0].text_frame.paragraphs[0].runs[0].size = 26
        if number not in missing:
            caption = slide.add_textbox(60 * EMU, (450 + (1 if number in drift else 0)) * EMU,
                                        400 * EMU, 30 * EMU, "Made-up figures")
            caption.name = "Caption"

        def marker():
            shape = slide.add_shape("rect", (800 if number not in moved else 100) * EMU,
                                    40 * EMU, 60 * EMU, 30 * EMU)
            shape.fill = "#204080" if number not in recolour else "#224182"
            return shape

        def panel():
            return slide.add_shape("roundRect", 600 * EMU, 300 * EMU, 300 * EMU, 150 * EMU)

        first, second = (panel, marker) if number in reorder else (marker, panel)
        a, b = first(), second()
        a.name, b.name = f"Shape {number * 2 + 10}", f"Shape {number * 2 + 11}"
        if number in only:
            slide.add_textbox(300 * EMU, 480 * EMU, 200 * EMU, 30 * EMU, "Callout").name = \
                "Callout"
    return deck


def _session(deck):
    toolbox = Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


def _ids(deck):
    return [f"s:{slide.slide_id}" for slide in deck.slides]


def _entry(data, name):
    return next(entry for entry in data["shapes"] if entry.get("name") == name)


def test_a_run_that_agrees_reports_nothing_to_look_at():
    deck = _run(5)
    data = deck.compare_slides(_ids(deck)).to_json()
    assert data == {"slides": _ids(deck), "matched": 5, "agree": 5, "shapes": [], "only": []}


def test_a_one_point_drift_is_a_near_miss_against_the_majority_with_its_fix():
    deck = _run(9, drift=(6, 7, 8))
    data = deck.compare_slides(_ids(deck)).to_json()
    caption = _entry(data, "Caption")
    drifted = [deck.slides[n].shapes[2].id for n in (6, 7, 8)]
    assert caption["differ"] == [{
        "property": "y", "value": 450, "on": 6,
        "near": [{"value": 451, "shapes": drifted}],
        "fix": {"tool": "ppt_set_shape",
                "items": [{"target": address, "y": 450} for address in drifted]}}]
    assert "by" not in caption                                   # matched by name
    assert data["agree"] == data["matched"] - 1


def test_a_size_set_on_the_slide_against_the_layouts_says_where_each_comes_from():
    deck = _run(9, sizes=(0, 7, 8))
    data = deck.compare_slides(_ids(deck)).to_json()
    title = data["shapes"][0]
    titles = [deck.slides[n].shapes[0].id for n in (0, 7, 8)]
    assert title["shape"] == deck.slides[0].shapes[0].id and title["name"] == "Title 1"
    assert title["differ"] == [{
        "property": "size", "value": 28, "on": 6, "from": "layout",
        "near": [{"value": 26, "from": "slide", "shapes": titles}],
        "fix": {"tool": "ppt_format_text",
                "items": [{"target": address, "size": 0} for address in titles]}}]


def test_a_missing_shape_and_one_only_on_a_slide():
    deck = _run(6, missing=(2,), only=(4,))
    data = deck.compare_slides(_ids(deck)).to_json()
    assert _entry(data, "Caption") == {"shape": deck.slides[0].shapes[2].id,
                                       "name": "Caption",
                                       "missing_on": [f"s:{deck.slides[2].slide_id}"]}
    callout = deck.slides[4].shapes[-1]
    assert data["only"] == [{"shape": callout.id, "name": "Callout", "by": "position",
                             "shapes": [callout.id]}]


def test_reordered_and_renamed_shapes_are_matched_by_kind_and_position():
    deck = _run(5, reorder=(1, 3))
    groups = match(deck.slides)
    assert [group.by for group in groups] == ["placeholder", "placeholder", "name",
                                              "position", "position"]
    marker, panel = groups[3], groups[4]
    assert {shape.preset for shape in marker.shapes.values()} == {"rect"}
    assert {shape.preset for shape in panel.shapes.values()} == {"roundRect"}
    assert len(marker.shapes) == len(panel.shapes) == 5
    data = deck.compare_slides(deck.slides).to_json()
    assert data["shapes"] == [] and data["agree"] == 5


def test_a_shade_off_is_a_near_miss_and_a_far_move_is_another_difference():
    deck = _run(5, recolour=(2,))
    data = deck.compare_slides(deck.slides).to_json()
    marker = next(e for e in data["shapes"] if e["by"] == "position")
    assert marker["differ"] == [{
        "property": "fill", "value": "#204080", "on": 4,
        "near": [{"value": "#224182", "shapes": [deck.slides[2].shapes[3].id]}],
        "fix": {"tool": "ppt_set_shape",
                "items": [{"target": deck.slides[2].shapes[3].id, "fill": "#204080"}]}}]
    # Moved 700 pt: too far to match by position, so it is a shape of its own and the
    # marker is missing there.
    deck = _run(5, moved=(4,))
    data = deck.compare_slides(deck.slides).to_json()
    assert [e.get("missing_on") for e in data["shapes"]] == [[f"s:{deck.slides[4].slide_id}"]]
    assert [o["shapes"] for o in data["only"]] == [[deck.slides[4].shapes[3].id]]


def test_a_larger_difference_is_listed_apart_and_carries_no_fix():
    deck = _run(4)
    deck.slides[3].shapes[2].top = 470 * EMU
    data = deck.compare_slides(deck.slides, within=2).to_json()
    caption = _entry(data, "Caption")
    assert caption["differ"] == [{"property": "y", "value": 450, "on": 3,
                                  "other": [{"value": 470,
                                             "shapes": [deck.slides[3].shapes[2].id]}]}]
    wider = deck.compare_slides(deck.slides, within=25).to_json()
    assert "near" in _entry(wider, "Caption")["differ"][0]


def test_text_colour_and_body_sizes_per_level():
    deck = _run(4)
    body = deck.slides[1].shapes[1]
    body.text_frame.paragraphs[1].level = 1
    deck.slides[2].shapes[2].text_frame.paragraphs[0].runs[0].color = "#C00000"
    data = deck.compare_slides(deck.slides).to_json()
    caption = _entry(data, "Caption")["differ"]
    assert caption == [{"property": "color", "value": "#000000", "on": 3,
                        "other": [{"value": "#C00000",
                                   "shapes": [deck.slides[2].shapes[2].id]}]}]
    # The body with a second level: the level's size is compared only where it is.
    assert all(entry["shape"] != deck.slides[0].shapes[1].id for entry in data["shapes"])


def test_the_tool_measures_pages_and_the_fixes_resolve_the_near_misses():
    deck = _run(9, drift=(6, 7, 8), sizes=(0, 7, 8))
    toolbox, session, doc = _session(deck)
    result = toolbox.dispatch(session, "ppt_compare_slides",
                              {"doc": doc, "slides": list(range(1, 10))})
    assert result.ok, json.dumps(result.to_json())
    assert result.summary.startswith("Compared 9 slides: 5 shape(s) matched, 2 differ")
    assert result.total == 2 and result.next_cursor is None and not result.changed
    for entry in result.data["shapes"]:
        fix = entry["differ"][0]["fix"]
        applied = toolbox.dispatch(session, fix["tool"], {"doc": doc, "items": fix["items"]})
        assert applied.ok, json.dumps(applied.to_json())
    again = toolbox.dispatch(session, "ppt_compare_slides",
                             {"doc": doc, "slides": _ids(deck)})
    assert again.data["shapes"] == [] and again.data["agree"] == 5


def test_the_tool_refuses_fewer_than_two_slides_a_repeat_counting_once():
    deck = _run(3)
    toolbox, session, doc = _session(deck)
    one = toolbox.dispatch(session, "ppt_compare_slides", {"doc": doc, "slides": [1]})
    assert not one.ok and one.error.message.startswith("compare at least two slides")
    twice = toolbox.dispatch(session, "ppt_compare_slides",
                             {"doc": doc, "slides": [1, f"s:{deck.slides[0].slide_id}"]})
    assert not twice.ok and twice.error.message.startswith("compare at least two slides")


def test_a_long_comparison_pages_by_shape():
    deck = Document.new(size="16:9")
    for number in range(3):
        slide = deck.add_slide("Blank")
        for k in range(50):
            box = slide.add_shape("rect", k * 18 * EMU, (10 + number) * EMU, 12 * EMU,
                                  12 * EMU)
            box.name = f"Tile {k}"
    toolbox, session, doc = _session(deck)
    first = toolbox.dispatch(session, "ppt_compare_slides", {"doc": doc, "slides": [1, 2, 3]})
    assert first.total == 50 and len(first.data["shapes"]) == 40 and first.next_cursor
    assert "only" in first.data
    rest = toolbox.dispatch(session, "ppt_compare_slides",
                            {"doc": doc, "slides": [1, 2, 3], "cursor": first.next_cursor})
    assert len(rest.data["shapes"]) == 10 and rest.next_cursor is None
    assert "only" not in rest.data


def test_a_level_past_the_first_is_compared_on_its_own():
    deck = _run(4)
    for slide in deck.slides:
        slide.shapes[1].text_frame.paragraphs[1].level = 1
    deeper = deck.slides[3].shapes[1].text_frame.paragraphs[1]
    deeper.runs[0].size = 23
    data = deck.compare_slides(deck.slides).to_json()
    body = data["shapes"][0]
    assert body["shape"] == deck.slides[0].shapes[1].id
    differ = body["differ"]
    assert [(d["property"], d.get("level"), d["near"][0]["shapes"]) for d in differ] == \
        [("size", 1, [deck.slides[3].shapes[1].id])]
    # The master gives level 1 the majority's size: removing the slide's restores it.
    assert differ[0]["fix"]["items"] == [{"target": deck.slides[3].shapes[1].id, "size": 0}]
    # With the first level's size set on the slide too, size 0 would undo that as well: the
    # fix sets the level's paragraphs instead.
    deck.slides[3].shapes[1].text_frame.paragraphs[0].runs[0].size = 30
    data = deck.compare_slides(deck.slides).to_json()
    fix = next(d for d in data["shapes"][0]["differ"] if d.get("level") == 1)["fix"]
    assert fix["items"] == [{"target": deeper.address, "size": 24}]


def test_the_library_refuses_one_slide_and_a_slide_given_twice():
    import pytest

    deck = _run(2)
    with pytest.raises(ValueError, match="at least two"):
        deck.compare_slides(deck.slides[:1])
    with pytest.raises(ValueError, match="given twice"):
        deck.compare_slides([deck.slides[0], f"s:{deck.slides[0].slide_id}"])


# -- renamed shapes (round 8, M4) -----------------------------------------------------------


def _caption(deck, index):
    return next(shape for shape in deck.slides[index].shapes if shape.name == "Caption")


def test_a_shape_renamed_on_a_rebuilt_slide_is_matched_in_place_and_the_name_differs():
    deck = _run(4)
    renamed = _caption(deck, 3)
    renamed.name = "Source line"
    data = deck.compare_slides(deck.slides).to_json()
    assert data["only"] == [] and data["matched"] == 5
    entry = _entry(data, "Caption")
    assert entry == {"shape": _caption(deck, 0).id, "name": "Caption", "differ": [
        {"property": "name", "value": "Caption", "on": 3,
         "other": [{"value": "Source line", "shapes": [renamed.id]}]}]}
    # Its other properties are compared as before: a drift on the renamed one is a near-miss.
    renamed.top = 451 * EMU
    differ = _entry(deck.compare_slides(deck.slides).to_json(), "Caption")["differ"]
    assert [d["property"] for d in differ] == ["name", "y"]
    assert differ[1]["near"] == [{"value": 451, "shapes": [renamed.id]}]


def test_a_renamed_first_slide_still_reports_the_majoritys_name():
    deck = _run(4)
    _caption(deck, 0).name = "Source line"
    data = deck.compare_slides(deck.slides).to_json()
    (entry,) = data["shapes"]
    assert entry["shape"] == deck.slides[0].shapes[2].id
    assert entry["differ"] == [{"property": "name", "value": "Caption", "on": 3, "other": [
        {"value": "Source line", "shapes": [deck.slides[0].shapes[2].id]}]}]


def test_half_the_run_under_another_name_is_one_group():
    deck = _run(4)
    renamed = [_caption(deck, 2), _caption(deck, 3)]
    for shape in renamed:
        shape.name = "Source line"
    groups = match(deck.slides)
    assert [group.by for group in groups] == ["placeholder", "placeholder", "name",
                                              "position", "position"]
    assert list(groups[2].shapes) == [slide.slide_id for slide in deck.slides]
    entry = _entry(deck.compare_slides(deck.slides).to_json(), "Caption")
    # A tie: the first slide's name is the majority's.
    assert entry["differ"][0] == {"property": "name", "value": "Caption", "on": 2, "other": [
        {"value": "Source line", "shapes": [shape.id for shape in renamed]}]}


def test_a_renamed_shape_far_away_is_still_missing_and_only():
    deck = _run(4)
    far = _caption(deck, 3)
    far.name, far.left, far.top = "Source line", 500 * EMU, 40 * EMU
    data = deck.compare_slides(deck.slides).to_json()
    assert _entry(data, "Caption")["missing_on"] == [f"s:{deck.slides[3].slide_id}"]
    assert "differ" not in _entry(data, "Caption")
    assert data["only"] == [{"shape": far.id, "name": "Source line", "by": "position",
                             "shapes": [far.id]}]


def test_two_renamed_shapes_each_join_the_nearest_name_group_once():
    deck = Document.new(size="16:9")
    for number in range(4):
        slide = deck.add_slide("Blank")
        left = slide.add_shape("rect", 100 * EMU, 100 * EMU, 200 * EMU, 60 * EMU)
        right = slide.add_shape("rect", 130 * EMU, 110 * EMU, 200 * EMU, 60 * EMU)
        if number == 2:            # rebuilt: other names, added in the other order
            left, right = right, left
            left.left, right.left = 100 * EMU, 130 * EMU
            left.top, right.top = 100 * EMU, 110 * EMU
            left.name, right.name = "Valve 1", "Valve 2"
        else:
            left.name, right.name = "Inlet", "Outlet"
    for _ in range(2):             # the same answer each time
        groups = match(deck.slides)
        assert [group.by for group in groups] == ["name", "name"]
        inlet, outlet = groups
        third = deck.slides[2].slide_id
        assert inlet.shapes[third].name == "Valve 1" and outlet.shapes[third].name == "Valve 2"
        assert len({id(shape) for group in groups for shape in group.shapes.values()}) == 8
    data = deck.compare_slides(deck.slides).to_json()
    assert [entry["differ"][0]["other"][0]["value"] for entry in data["shapes"]] == \
        ["Valve 1", "Valve 2"]
    assert data["only"] == []


def test_the_tool_reports_the_renamed_shape():
    deck = _run(4)
    renamed = _caption(deck, 1)
    renamed.name = "Source line"
    toolbox, session, doc = _session(deck)
    result = toolbox.dispatch(session, "ppt_compare_slides", {"doc": doc, "slides": _ids(deck)})
    assert result.ok, json.dumps(result.to_json())
    assert result.total == 1 and result.data["only"] == []
    assert result.data["shapes"][0]["differ"][0]["other"] == [
        {"value": "Source line", "shapes": [renamed.id]}]
