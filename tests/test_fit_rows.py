"""Rows fitted to their text from a tool, and the frame's height kept (round 54, R114).

A summary slide's table has one row per item; callers copy a sample row and replace its text,
and each copy keeps the sample's stored height -- a minimum, which PowerPoint keeps when the
text is replaced by shorter text (measured: ``tools/table_fit_probe.py``,
``tests/fixtures/table-fit-probe.json``) -- so a short item leaves a wider gap than a long
one.  ``ppt_edit_table``'s ``fit_rows`` sets the rows (all, or those named) to their text's
height and the frame to the rows' sum; ``check`` and a changing call's checks name tables
whose rows are stored taller than their text with uneven gaps (``row_slack``), with that
call as the fix; and any row height change keeps the frame's stored height at the rows' sum
as drawn, as PowerPoint writes it (rows stored 0 included).  Every deck here is made up.

Rules (agent-work/tool-design-rules.md): 1 (``rows: []`` as strict mode fills it means every
row), 2 (a row that is not there is refused with the table's rows), 3 (fitting rows that fit
is no change), 4 (the rows changed are listed up to 40, the rest counted), 5 (one call for
every row, one undo step), 7 (the uneven gaps a person sees are a check fact), 8 (the heights
match PowerPoint's, measured).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.fit import row_slack, row_text_heights
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, EMU

pytest.importorskip("pptx2svg")

PROBE = json.loads((Path(__file__).parent / "fixtures" / "table-fit-probe.json").read_text())
LONG = ("Move the spare parts to the north shelf, label every bin with its new number, and "
        "check the counts against last month's sheet before Friday")
MID = "Order two more ladders for the back room and return the broken one"
#: The probe's texts, in its order: PowerPoint drew them 28.8, 72, 50.4, 28.8 and 28.8 pt.
TEXTS = ["Shelf count closed", LONG, MID, "Done", "Shelf count closed"]
#: The summary slide: the sample row first, four rows copied from it.
SUMMARY = ["Shelf count closed", LONG, "Done", MID, "Grades in"]


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def run(toolbox):
    session = toolbox.session(clock=CLOCK)
    deck = Document.new(size="16:9")
    deck.add_slide("Title Only").title = "Weekly summary"
    doc = session.open(deck.to_bytes(), "summary.pptx")

    def call(name, **arguments):
        result = toolbox.dispatch(session, name, {"doc": doc, **arguments})
        assert result.ok, json.dumps(result.to_json(), indent=1)
        return result

    def refused(name, **arguments):
        result = toolbox.dispatch(session, name, {"doc": doc, **arguments})
        assert not result.ok
        return result.error

    call.refused = refused
    call.deck = lambda: session.entry(doc).document
    return call


def _points(values):
    return [round(value / EMU, 1) for value in values]


def _summary(run, texts=SUMMARY):
    """The made-up summary table: a 62 pt sample row, the rest copied from it, then filled."""
    run("ppt_add_table", slide="s:256", box={"x": 60, "y": 120, "w": 470, "h": 62}, rows=1,
        columns=2, data=[["1", texts[0]]], style="No Style, No Grid", widths=[50, 420])
    run("ppt_edit_table", target="256.3", action="set_heights", sizes=[62])
    for _ in texts[1:]:
        run("ppt_edit_table", target="256.3", action="insert_row", like=0)
    return run("ppt_edit_table", target="256.3", action="set_cells",
               cells=[{"row": row, "col": col, "text": value}
                      for row, text in enumerate(texts)
                      for col, value in enumerate((str(row + 1), text))])


def _probe_table(heights):
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Blank")
    table = slide.add_table(5, 2, 60 * EMU, 60 * EMU, 470 * EMU, 0,
                            style="No Style, No Grid").table
    table.set_column_width(0, 50 * EMU)
    table.set_column_width(1, 420 * EMU)
    for row, text in enumerate(TEXTS):
        table.cell(row, 0).text = str(row + 1)
        table.cell(row, 1).text = text
    table.set_row_heights([round(h * EMU) for h in heights])
    return deck, table


# -- rule 8: PowerPoint's own heights --------------------------------------------------------


def test_rows_text_heights_and_drawn_heights_are_powerpoints():
    for case in PROBE["cases"]:
        if "drawn_pt" not in case:
            continue
        deck, table = _probe_table(case["written_h_pt"])      # the heights as PowerPoint has them
        frame = deck.shape(table.address)
        assert _points(row_text_heights(frame)) == PROBE["text_pt"]
        assert _points(table.drawn_row_heights) == case["drawn_pt"], case["name"]


def test_the_frame_takes_the_rows_drawn_sum_as_powerpoint_writes_it():
    """PowerPoint writes the frame's cy as the rows' drawn sum, also for rows stored 0 (the
    height-0 case left it at 0 before) and for rows set lower than their text."""
    for case in PROBE["cases"]:
        if "drawn_pt" not in case:
            continue
        deck, table = _probe_table(case["written_h_pt"])
        assert _points(table.row_heights) == case["written_h_pt"]
        assert round(deck.shape(table.address).height / EMU, 1) == case["written_cy_pt"], \
            case["name"]


def test_fit_rows_writes_what_powerpoint_draws():
    deck, table = _probe_table([62] * 5)
    table.fit_rows()
    assert _points(table.row_heights) == _points(table.drawn_row_heights) == PROBE["text_pt"]
    assert round(deck.shape(table.address).height / EMU, 1) == round(sum(PROBE["text_pt"]), 1)


# -- the tool --------------------------------------------------------------------------------


def test_before_copied_rows_keep_the_sample_height_and_check_names_them(run):
    result = _summary(run)
    assert result.data["heights"] == [62.0] * 5
    # Rule 7: the uneven gaps, with the call that fixes them.
    assert result.checks["row_slack"] == [{
        "shape": "256.3",
        "rows": [{"row": 0, "h": 62.0, "text": 28.8}, {"row": 2, "h": 62.0, "text": 28.8},
                 {"row": 3, "h": 62.0, "text": 50.4}, {"row": 4, "h": 62.0, "text": 28.8}],
        "fix": {"tool": "ppt_edit_table", "target": "256.3", "action": "fit_rows"}}]
    assert "row_slack" in result.checks["row_slack_note"]
    checked = run("check", slides=[1])
    assert checked.data["row_slack"][0]["shape"] == "256.3"


def test_fit_rows_sets_every_row_to_its_text_and_the_frame_to_their_sum(run):
    _summary(run)
    result = run("ppt_edit_table", target="256.3", action="fit_rows")
    assert result.summary == ("fit_rows on table 256.3: 5 row(s) set to their text's height; "
                              "frame 310 -> 208.8 pt")
    assert result.data["fit"] == [{"row": 0, "from": 62.0, "to": 28.8},
                                  {"row": 1, "from": 62.0, "to": 72.0},
                                  {"row": 2, "from": 62.0, "to": 28.8},
                                  {"row": 3, "from": 62.0, "to": 50.4},
                                  {"row": 4, "from": 62.0, "to": 28.8}]
    assert result.data["frame_h"] == {"from": 310.0, "to": 208.8}
    assert result.data["heights"] == [28.8, 72.0, 28.8, 50.4, 28.8]
    table = run.deck().shape("256.3").table
    assert _points(table.drawn_row_heights) == result.data["heights"]
    assert "row_slack" not in result.checks
    assert run.deck().shape("256.3").height == sum(table.row_heights)


def test_fit_rows_is_one_undo_step(run):
    _summary(run)
    run("ppt_edit_table", target="256.3", action="fit_rows")
    run("undo")
    assert _points(run.deck().shape("256.3").table.row_heights) == [62.0] * 5


def test_fit_rows_named_rows_leaves_the_others(run):
    _summary(run)
    result = run("ppt_edit_table", target="256.3", action="fit_rows", rows=[2, 4])
    assert result.data["heights"] == [62.0, 62.0, 28.8, 62.0, 28.8]
    # The frame is the rows as drawn: row 1's text grows it past its 62 pt.
    assert result.data["frame_h"] == {"from": 310.0, "to": 253.6}


def test_fit_rows_on_rows_that_fit_is_no_change(run):
    _summary(run)
    run("ppt_edit_table", target="256.3", action="fit_rows")
    again = run("ppt_edit_table", target="256.3", action="fit_rows")
    assert again.changed == []
    assert again.summary == ("fit_rows on table 256.3: no change, every row already fits its "
                             "text")


def test_fit_rows_takes_an_empty_rows_list_as_every_row(run):
    _summary(run)
    result = run("ppt_edit_table", target="256.3", action="fit_rows", rows=[])
    assert len(result.data["fit"]) == 5


def test_fit_rows_refuses_a_row_that_is_not_there_naming_the_rows(run):
    _summary(run)
    error = run.refused("ppt_edit_table", target="256.3", action="fit_rows", rows=[1, 7])
    assert error.code == "not_found" and error.field == "rows"
    assert "no row 7; the table has 5 (rows 0-4)" in error.message


def test_fit_rows_lists_forty_rows_and_counts_the_rest(run):
    texts = [f"Item {n}" for n in range(45)]
    _summary(run, texts)
    result = run("ppt_edit_table", target="256.3", action="fit_rows")
    assert len(result.data["fit"]) == 40
    assert result.data["rows_left_out"] == 5


# -- the frame and the check fact ------------------------------------------------------------


def test_set_heights_to_zero_keeps_the_frame_at_the_rows_drawn_sum(run):
    _summary(run)
    result = run("ppt_edit_table", target="256.3", action="set_heights", sizes=[0] * 5)
    assert result.data["heights"] == [0.0] * 5
    assert round(run.deck().shape("256.3").height / EMU, 1) == 208.8


def test_rows_sharing_a_box_evenly_are_not_slack(run):
    data = [["Part", "Note"]] + [[f"Part {n}", "Short"] for n in range(1, 5)]
    result = run("ppt_add_table", slide="s:256", box={"x": 40, "y": 110, "w": 640, "h": 276},
                 rows=5, columns=2, data=data, size=10, fill_box=True)
    assert "row_slack" not in result.checks
    assert row_slack(run.deck().shape("256.3")) == []


def test_rows_of_one_height_whose_text_differs_are_slack_rows_even_ones_are_not():
    deck, table = _probe_table([62] * 5)
    frame = deck.shape(table.address)
    assert [row for row, _, _ in row_slack(frame)] == [0, 2, 3, 4]
    for row, text in enumerate(TEXTS):
        table.cell(row, 1).text = "Done"
    assert row_slack(frame) == []                  # every gap the same: meant


def test_fit_rows_in_the_library_takes_rows_or_a_box_not_both():
    from pptx_agent.edit.errors import ArgumentError

    deck, table = _probe_table([62] * 5)
    with pytest.raises(ArgumentError, match="give one or the other"):
        table.fit_rows(276 * EMU, fill=True, rows=[1])
    table.fit_rows(rows=[0, 3])
    assert _points(table.row_heights) == [28.8, 62.0, 62.0, 28.8, 62.0]
