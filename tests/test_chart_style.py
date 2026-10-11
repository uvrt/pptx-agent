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
    ({"action": "set_gridlines", "gridlines": "both"}, "needs axis"),
])
def test_edit_chart_refuses_what_cannot_be_written(session, arguments, words):
    run(session, target="s:256", action="add", chart_type="bar", categories=["A", "B", "C"],
        data=[{"name": "Score", "values": [1, 2, 3]}], box=BOX, ref="r")
    error = refused(session, target="$r", **arguments)
    assert error.code == "invalid_arguments" and words in error.message, error.message


def test_a_label_position_is_not_a_legend_position(session):
    error = refused(session, target="s:256", action="add", chart_type="bar",
                    categories=["A"], data=[{"name": "S", "values": [1]}], box=BOX,
                    position="center")
    assert error.field == "position"
