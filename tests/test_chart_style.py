"""A chart's look on a slide: axis scales, gridlines, series and point colours and markers,
combination charts, chart text and point labels (ooxml-edit's ``Chart`` formatting, tested
there), filled radar and bubble charts -- through the library and through ``edit_chart``.

Here: the deck stays valid, labels and names are found as the deck finds them, every call is
one undo step, and the tool's arguments reach the library.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
EMU = 12700
CATEGORIES = ["Governance", "People", "Data", "Technology", "Procure"]
CURRENT = [2.13, 2.10, 2.17, 2.33, 2.08]
TARGET = [4, 3, 4, 3, 4]


def gap_chart():
    deck = Document.new()
    slide = deck.add_slide(deck.layouts[5])
    frame = slide.add_chart("column", CATEGORIES, [{"name": "Current", "values": CURRENT},
                                                   {"name": "Target", "values": TARGET}],
                            40 * EMU, 120 * EMU, 880 * EMU, 380 * EMU)
    return deck, deck.shape(frame.id).chart


def test_the_maturity_gap_chart_through_the_library():
    deck, chart = gap_chart()
    chart.set_axis_scale("value", minimum=0, maximum=5, major_unit=1)
    chart.format_series("Current", fill="tx2")
    chart.format_point("Current", "Procure", fill="accent2")
    chart.set_series_type("Target", "line")
    chart.format_series("target", line=None, marker="diamond", marker_size=10,
                        marker_fill="accent6")  # names found as series_named finds them
    chart.set_text_style(size=10)
    chart.set_gridlines("value", major=False)
    assert chart.chart_types == ["bar", "line"]
    assert chart.series_format("Current")["points"] == {4: {"fill": "accent2", "line": None}}
    assert chart.workbook_values()["series"][1] == {"name": "Target", "values": TARGET}
    assert deck.validate() == []
    reopened = Document.open(deck.to_bytes())
    again = reopened.shape(chart.address).chart
    assert again.scales["value"]["maximum"] == 5
    for _ in range(7):
        deck.undo()
    assert deck.shape(chart.address).chart.chart_types == ["bar"]


def test_a_category_label_no_chart_has_is_a_label_error():
    from pptx_agent import LabelError

    _, chart = gap_chart()
    with pytest.raises(LabelError):
        chart.format_point("Current", "Nowhere", fill="accent2")


@pytest.mark.parametrize("kind", ["radar_filled", "bubble"])
def test_the_new_kinds_on_a_slide(kind):
    deck = Document.new()
    slide = deck.add_slide(deck.layouts[5])
    if kind == "bubble":
        frame = slide.add_chart("bubble", [1, 2, 4], [{"name": "Recs", "values": [4.5, 3, 4],
                                                       "size": [3, 2, 1]}],
                                0, 0, 400 * EMU, 300 * EMU, bubble_scale=60)
    else:
        frame = slide.add_chart("radar_filled", ["Speed", "Cost", "Risk"],
                                [{"name": "Now", "values": [3, 2, 4]}], 0, 0, 400 * EMU,
                                300 * EMU)
    chart = deck.shape(frame.id).chart
    assert chart.chart_type == ("bubble" if kind == "bubble" else "radar")
    assert deck.validate() == []


def test_a_ranked_bar_chart_reads_top_down():
    deck = Document.new()
    slide = deck.add_slide(deck.layouts[5])
    frame = slide.add_chart("bar", ["Worst", "Middle", "Best"],
                            [{"name": "Score", "values": [1.8, 2.4, 3.1]}], 0, 0, 400 * EMU,
                            300 * EMU, category_order="reversed", gridlines=False)
    chart = deck.shape(frame.id).chart
    assert chart.axis_scale("category")["reverse"]
    assert chart.categories == ["Worst", "Middle", "Best"]
    assert chart.workbook_values()["categories"] == ["Worst", "Middle", "Best"]
    assert not list(chart._root().iter(f"{C}majorGridlines"))


# -- the tool -------------------------------------------------------------------------------------


@pytest.fixture()
def session():
    deck = Document.new()
    deck.add_slide(deck.layouts[5])
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        current = box.session()
        current.open(deck.to_bytes(), "deck.pptx")
        yield box, current


def run(session, **arguments):
    box, current = session
    result = box.dispatch(current, "edit_chart", {"doc": "d1", **arguments})
    assert result.ok, result.to_json()
    return result


def refused(session, **arguments):
    box, current = session
    result = box.dispatch(current, "edit_chart", {"doc": "d1", **arguments})
    assert not result.ok
    return result.error


BOX = {"x": 40, "y": 120, "w": 880, "h": 380}


def test_edit_chart_builds_the_maturity_gap_chart(session):
    run(session, target="s:256", action="add", chart_type="column", categories=CATEGORIES,
        data=[{"name": "Current", "values": CURRENT},
              {"name": "Target", "values": TARGET, "type": "line"}],
        box=BOX, axis="value", scale={"min": 0, "max": 5, "major_unit": 1}, gridlines="none",
        ref="gap")
    run(session, target="$gap", action="format_series", series="Current",
        style={"fill": "tx2"})
    run(session, target="$gap", action="format_point", series="Current", category="Procure",
        style={"fill": "accent2"})
    run(session, target="$gap", action="format_series", series="Target",
        style={"line": "none", "marker": "diamond", "marker_size": 10, "marker_fill": "accent6"})
    styled = run(session, target="$gap", action="set_text_style", style={"size": 10})
    assert styled.data["types"] == ["bar", "line"]
    read = run(session, target="$gap", action="read")
    assert read.data["scales"]["value"] == {"minimum": 0, "maximum": 5, "major_unit": 1,
                                            "minor_unit": None, "number_format": None,
                                            "reverse": False, "major_gridlines": False,
                                            "minor_gridlines": False}
    assert read.data["series_format"]["Target"]["marker"] == "diamond"
    assert read.data["series_format"]["Current"]["points"] == {4: {"fill": "accent2",
                                                                   "line": None}}
    box, current = session
    assert current.entry("d1").document.validate() == []


def test_edit_chart_reverses_categories_and_labels_points(session):
    added = run(session, target="s:256", action="add", chart_type="bar",
                categories=["Worst", "Middle", "Best"],
                data=[{"name": "Score", "values": [1.8, 2.4, 3.1]}], box=BOX, axis="category",
                scale={"reverse": True}, ref="r")
    read = run(session, target="$r", action="read")
    assert read.data["scales"]["category"]["reverse"] is True
    assert added.data["categories"] == ["Worst", "Middle", "Best"]
    run(session, target="s:256", action="add", chart_type="bubble", values=[1, 2, 4],
        data=[{"name": "Recs", "values": [4.5, 3, 4], "size": [3, 2, 1]}], box=BOX,
        position="none", ref="b")
    labelled = run(session, target="$b", action="set_point_labels", series="Recs",
                   labels=["R1", "R2", "R3"], position="right", cells=True)
    assert labelled.data["point_labels"] == {"Recs": ["R1", "R2", "R3"]}
    run(session, target="$b", action="set_axis_scale", axis="x",
        scale={"min": 0, "max": 5, "auto": ["major_unit"]})


@pytest.mark.parametrize("arguments, words", [
    ({"action": "set_series_type", "series": "Score", "chart_type": "line"}, "horizontal bars"),
    ({"action": "format_series", "series": "Score", "style": {"size": 10}}, "not for"),
    ({"action": "format_series", "series": "Score", "style": {"fill": "purple"}}, "neither"),
    ({"action": "set_axis_scale", "axis": "category", "scale": {"max": 3}}, "no numeric"),
    ({"action": "set_point_labels", "series": "Score", "labels": ["a", "b", "c"],
      "position": "right"}, "center, inside_end"),
    ({"action": "set_gridlines", "gridlines": "both"}, "needs axis: category or value"),
    ({"action": "set_axis_scale", "scale": {"reverse": True}}, "needs axis: category or value"),
    ({"action": "set_value", "series": "Score", "value": 2},
     "set_value needs category: 'A', 'B' or 'C' (or a number from 0 to 2)"),
    ({"action": "rename_series", "text": "Now"}, "rename_series needs series: 'Score' (or 0)"),
    ({"action": "rename_series", "series": "Score"}, "rename_series needs text: the new name"),
    ({"action": "set_values", "series": "Score"}, "needs values: one number per category (3)"),
    ({"action": "set_axis_title", "text": "x"}, "set_axis_title needs axis: category or value"),
    ({"action": "format_point", "series": "Score", "style": {"fill": "tx2"}},
     "format_point needs category: 'A', 'B' or 'C'"),
])
def test_edit_chart_refuses_what_cannot_be_written(session, arguments, words):
    run(session, target="s:256", action="add", chart_type="bar", categories=["A", "B", "C"],
        data=[{"name": "Score", "values": [1, 2, 3]}], box=BOX, ref="r")
    error = refused(session, target="$r", **arguments)
    assert error.code == "invalid_arguments" and words in error.message, error.message


def test_a_scale_without_an_axis_is_the_value_axis_s(session):
    run(session, target="s:256", action="add", chart_type="column", categories=["A", "B"],
        data=[{"name": "Score", "values": [1, 2]}], box=BOX, scale={"min": 0, "max": 4},
        ref="c")
    assert run(session, target="$c", action="read").data["scales"]["value"]["maximum"] == 4
    run(session, target="$c", action="set_axis_scale", scale={"major_unit": 1})
    assert run(session, target="$c", action="read").data["scales"]["value"]["major_unit"] == 1


def test_a_scale_two_axes_could_take_is_refused_and_nothing_is_made(session):
    box, current = session
    before = len(current.entry("d1").document.slides[0].shapes)
    error = refused(session, target="s:256", action="add", chart_type="scatter",
                    values=[1, 2], data=[{"name": "Y", "values": [3, 4]}], box=BOX,
                    scale={"max": 5})
    assert error.message.startswith("add with scale needs axis: x or y"), error.message
    assert len(current.entry("d1").document.slides[0].shapes) == before


@pytest.mark.parametrize("arguments, words", [
    ({"categories": ["A"], "data": [{"name": "S", "values": [1]}], "box": BOX},
     "add needs chart_type: one of column, stacked_column"),
    ({"chart_type": "column", "categories": ["A"], "box": BOX}, "add needs data: each series"),
    ({"chart_type": "column", "categories": ["A"], "data": [{"name": "S", "values": [1]}]},
     "add needs box: the frame as {x, y, w, h} in pt"),
    ({"chart_type": "column", "data": [{"name": "S", "values": [1]}], "box": BOX},
     "add needs categories: the category labels"),
])
def test_what_add_needs_is_named(session, arguments, words):
    error = refused(session, target="s:256", action="add", **arguments)
    assert words in error.message, error.message


def test_a_label_position_is_not_a_legend_position(session):
    error = refused(session, target="s:256", action="add", chart_type="bar",
                    categories=["A"], data=[{"name": "S", "values": [1]}], box=BOX,
                    position="center")
    assert error.field == "position"


# -- the dot chart (round 6, R10) ---------------------------------------------------------------

AREAS = ["North", "South", "East", "West"]


def test_edit_chart_adds_a_dot_chart_on_a_fixed_scale_in_one_call(session):
    added = run(session, target="s:256", action="add", chart_type="dot", categories=AREAS,
                data=[{"name": "Mean", "values": [3.2, 4.1, 2.7, 3.9]}], box=BOX,
                axis="value", scale={"min": 1, "max": 5, "major_unit": 1}, data_labels=True,
                number_format="0.0", colors=["tx2"], position="none", ref="dots")
    assert added.data["type"] == "dot"
    read = run(session, target="$dots", action="read")
    assert read.data["type"] == "dot" and read.data["types"] == ["line"]
    assert read.data["scales"]["value"]["minimum"] == 1
    assert read.data["scales"]["value"]["maximum"] == 5
    assert read.data["series_format"]["Mean"]["marker_fill"] == "tx2"
    run(session, target="$dots", action="set_values", series="Mean", values=[1, 2, 3, 4])
    grown = run(session, target="$dots", action="add_series", text="Goal", values=[4, 4, 4, 4])
    assert grown.data["type"] == "dot"
    box, current = session
    deck = current.entry("d1").document
    assert deck.validate() == []
    chart = deck.shape(read.data["address"]).chart
    assert chart.is_dot and chart.chart_type == "line"
    assert all(entry["shown"] for entry in chart.data_labels[:1])


def test_a_line_chart_made_line_less_reads_back_as_dot(session):
    run(session, target="s:256", action="add", chart_type="line", categories=AREAS,
        data=[{"name": "Mean", "values": [3.2, 4.1, 2.7, 3.9]}], box=BOX, ref="l")
    assert run(session, target="$l", action="read").data["type"] == "line"
    changed = run(session, target="$l", action="format_series", series="Mean",
                  style={"line": "none", "marker": "circle"})
    assert changed.data["type"] == "dot"
    box, current = session
    listed = box.dispatch(current, "ppt_read_slides",
                          {"doc": "d1", "slides": [1], "detail": "compact"})
    assert "dot, 1 series" in str(listed.to_json())
    outline = box.dispatch(current, "ppt_read_slides",
                           {"doc": "d1", "slides": [1], "detail": "outline"})
    assert "[chart: dot;" in str(outline.to_json())


# N4 (review): one lookup in both formats, ooxml-edit's; a name wins over a position.
# docx-agent's tests/test_chart_style.py makes the same calls on a document.


@pytest.mark.parametrize("names, which, changed", [
    (["1", "2"], "1", 0),          # the series called "1"
    (["2", "1"], "1", 1),
    (["North", "South"], "1", 1),  # no series called "1": position 1
])
def test_a_series_name_wins_over_a_position(session, names, which, changed):
    run(session, target="s:256", action="add", chart_type="column", categories=["1", "Q2"],
        data=[{"name": name, "values": [1, 2]} for name in names], box=BOX, ref="c")
    data = run(session, target="$c", action="set_value", series=which, category="1",
               value=9).data
    assert [series["values"] for series in data["series"]] == [
        [1 if k != changed else 9, 2] for k in range(2)]   # category "1" by name, too
    error = refused(session, target="$c", action="set_value", series="9", value=1,
                    category="Q2")
    assert error.code == "not_found" and error.field == "series"
    assert error.valid_options == names
    error = refused(session, target="$c", action="rename_series", series="East", text="x")
    assert error.code == "label_not_found" and error.valid_options == names
