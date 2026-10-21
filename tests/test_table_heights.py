"""Table row heights in a run of slides (round 11, P1-P3).

P1: a row's height is a minimum, so ``ppt_add_table``'s box ``h`` is an upper bound: each
row takes its text's height (:meth:`Table.fit_rows`) and the result says how much of the
box the table takes; ``fill_box`` shares the room left among the body rows, the header
keeping its text's height.  P2: ``insert_row`` takes its ``like`` row's height and
``delete_row`` takes its row's out of the frame; no other row changes.  P3:
``ppt_compare_slides`` compares matched tables inside -- header height, fill and text
size, body row height (row counts as context), column widths, insets, borders -- with the
fixing call.  Every deck and table here is made up.
"""

from __future__ import annotations

import json

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, EMU

pytest.importorskip("pptx2svg")

BOX = {"x": 40, "y": 110, "w": 640, "h": 276}
DATA = [["Part", "Crew", "Due"]] + [[f"Part {n}", "Crew", "Fri"] for n in range(1, 5)]
#: A 10 pt line and the cell's 3.6 pt top and bottom margins.
LINE = 19.2


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def run(toolbox):
    session = toolbox.session(clock=CLOCK)
    deck = Document.new(size="16:9")
    deck.add_slide("Title Only").title = "Parts"
    doc = session.open(deck.to_bytes(), "parts.pptx")

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


def _rows(call, address):
    table = call.deck().shape(address).table
    return _points(table.drawn_row_heights)


def _frame(call, address):
    return round(call.deck().shape(address).height / EMU, 1)


# -- P1 -------------------------------------------------------------------------------------


def test_rows_take_their_texts_height_and_the_result_says_what_of_the_box_they_take(run):
    result = run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA,
                 size=10)
    assert result.summary == ("Added a 5x3 table 256.3: it takes 96 pt of the 276 pt box, "
                              "each row its text's height (header row 19.2 pt); fill_box: "
                              "true shares the rest among the body rows")
    assert result.data["heights"] == [LINE] * 5
    table = run.deck().shape("256.3").table
    assert _points(table.row_heights) == _points(table.drawn_row_heights) == [LINE] * 5
    assert _frame(run, "256.3") == 96


def test_fill_box_shares_the_room_among_the_body_rows_the_header_at_its_texts_height(run):
    result = run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA,
                 size=10, fill_box=True)
    assert result.summary == ("Added a 5x3 table 256.3: its body rows share the 276 pt box "
                              "(header row 19.2 pt)")
    assert _rows(run, "256.3") == [LINE] + [64.2] * 4
    assert run.deck().shape("256.3").height == 276 * EMU


def test_rows_taller_than_the_box_run_past_it_and_the_result_says_by_how_much(run):
    data = [["Part", "Note"]] + [[f"Part {n}", "a note long enough to wrap onto a few lines "
                                  "in a narrow column of this made-up table"]
                                 for n in range(1, 9)]
    result = run("ppt_add_table", slide="s:256", box={**BOX, "w": 300, "h": 120}, rows=9,
                 columns=2, data=data, size=12)
    taken = sum(run.deck().shape("256.3").table.row_heights) / EMU
    assert taken > 120
    assert ", each row its text's height; continue: true goes on over the next slides" \
        in result.summary
    assert "pt past the 120 pt box" in result.summary


def test_a_box_of_no_height_takes_the_texts_and_says_so(run):
    result = run("ppt_add_table", slide="s:256", box={**BOX, "h": 0}, rows=5, columns=3,
                 data=DATA, size=10)
    assert result.summary.endswith(": it takes 96 pt, each row its text's height (header "
                                   "row 19.2 pt)")
    assert _frame(run, "256.3") == 96


def test_fill_box_is_for_one_table(run):
    for extra in ({"continue": True}, {"measure": True}):
        error = run.refused("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3,
                            data=DATA, fill_box=True, **extra)
        assert error.field == "fill_box"


def test_fit_rows_in_the_library_is_one_undo_step():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    table = slide.add_table(5, 3, 40 * EMU, 110 * EMU, 640 * EMU, 276 * EMU).table
    for row, values in enumerate(DATA):
        for column, text in enumerate(values):
            table.cell(row, column).text = text
    table.set_text_size(10)
    assert _points(table.row_heights) == [55.2] * 5           # the frame shared evenly
    table.fit_rows()
    assert _points(table.row_heights) == [LINE] * 5
    assert deck.shape(table.address).height == sum(table.row_heights)
    table.fit_rows(276 * EMU, fill=True)
    assert _points(table.row_heights) == [LINE] + [64.2] * 4
    assert sum(table.row_heights) == 276 * EMU
    deck.undo()
    assert _points(table.row_heights) == [LINE] * 5


# -- P2 -------------------------------------------------------------------------------------


def test_insert_and_delete_row_keep_the_other_rows_heights_and_the_frame_follows(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10,
        fill_box=True)
    result = run("ppt_edit_table", target="256.3", action="delete_row", index=4)
    assert result.data["heights"] == [LINE] + [64.2] * 3
    assert _frame(run, "256.3") == round(LINE + 3 * 64.2, 1)
    result = run("ppt_edit_table", target="256.3", action="insert_row", index=1, like=0)
    assert result.data["heights"] == [LINE, LINE] + [64.2] * 3     # like the header
    run("ppt_edit_table", target="256.3", action="insert_row")      # like the row above
    assert _rows(run, "256.3") == [LINE, LINE] + [64.2] * 4
    assert _frame(run, "256.3") == round(2 * LINE + 4 * 64.2, 1)


def test_insert_and_delete_column_keep_the_other_widths(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10)
    run("ppt_edit_table", target="256.3", action="set_widths", columns=[0], sizes=[100])
    run("ppt_edit_table", target="256.3", action="insert_column", like=0)
    table = run.deck().shape("256.3").table
    widths = _points(table.column_widths)
    assert widths[0] == widths[3] == 100 and widths[1] == widths[2]
    run("ppt_edit_table", target="256.3", action="delete_column", index=1)
    shape = run.deck().shape("256.3")
    assert _points(shape.table.column_widths) == [100, widths[2], 100]
    assert shape.width == sum(shape.table.column_widths)


def test_set_heights_and_set_widths_take_the_rows_or_columns_listed(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10)
    result = run("ppt_edit_table", target="256.3", action="set_heights", rows=[0, 2],
                 sizes=[30, 40])
    assert result.data["heights"] == [30, LINE, 40, LINE, LINE]
    error = run.refused("ppt_edit_table", target="256.3", action="set_heights", rows=[0],
                        sizes=[30, 40])
    assert error.field == "sizes" and "2 sizes for 1 rows" in error.message
    error = run.refused("ppt_edit_table", target="256.3", action="set_heights", rows=[7],
                        sizes=[30])
    assert error.code == "not_found" and error.field == "rows"
    error = run.refused("ppt_edit_table", target="256.3", action="set_heights", sizes=[30])
    assert "list the rows in rows" in error.message


def test_set_shape_h_on_a_table_warns_that_its_rows_were_rescaled(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10)
    result = run("ppt_set_shape", items=[{"target": "256.3", "h": 192}])
    assert result.warnings == ["256.3 is a table: h rescaled its rows in proportion, the "
                               "header row 19.2 -> 38.4 pt; a row stays at least its text's "
                               "height, and ppt_edit_table set_heights sets rows"]
    assert not run("ppt_set_shape", items=[{"target": "256.3", "x": 50}]).warnings


# -- P3 and the reported run ------------------------------------------------------------------


def _copies(run, edits):
    """The table slide duplicated once per edit, each copy's table edited so."""
    tables = ["256.3"]
    for edit in edits:
        copy = run("ppt_manage_slides", action="duplicate", slide="s:256").data["slide"]
        address = f"{copy[2:]}.3"
        for action, arguments in edit:
            run("ppt_edit_table", target=address, action=action, **arguments)
        tables.append(address)
    return tables


def _compare(run, count):
    return run("ppt_compare_slides", slides=list(range(1, count + 1)))


def test_the_reported_run_agrees_and_reports_drift_only_when_there_is_some(run):
    """Built in a tall box, duplicated, rows deleted and inserted on the copies: every
    table's header and body rows stay at their text's height, so the comparison finds
    nothing; one copy stretched back to the box (``ppt_set_shape`` ``h``) is reported, with
    the calls that put its rows back."""
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10)
    tables = _copies(run, [[("delete_row", {"index": 4})],
                           [("insert_row", {"like": 1})],
                           [("delete_row", {"index": 3}), ("delete_row", {"index": 2})]])
    heights = {address: _rows(run, address) for address in tables}
    assert {len(rows): rows[0] for rows in heights.values()} == {5: LINE, 4: LINE, 6: LINE,
                                                                3: LINE}
    assert all(set(rows[1:]) == {LINE} for rows in heights.values())
    result = _compare(run, 4)
    assert result.data["shapes"] == [] and result.data["agree"] == 2

    run("ppt_set_shape", items=[{"target": tables[2], "h": 276}])
    result = _compare(run, 4)
    (entry,) = result.data["shapes"]
    assert entry["shape"] == "256.3"
    header, body = entry["differ"]
    assert header == {"property": "header_h", "value": LINE, "on": 3,
                      "other": [{"value": 46, "shapes": [tables[2]]}],
                      "fix": {"tool": "ppt_edit_table", "calls": [
                          {"target": tables[2], "action": "set_heights", "rows": [0],
                           "sizes": [LINE]}]}}
    assert body == {"property": "row_h", "value": LINE, "on": 3, "rows": [2, 3, 4],
                    "other": [{"value": 46, "rows": [5], "shapes": [tables[2]]}],
                    "fix": {"tool": "ppt_edit_table", "calls": [
                        {"target": tables[2], "action": "set_heights",
                         "rows": [1, 2, 3, 4, 5], "sizes": [LINE] * 5}]}}
    for item in (header, body):
        for arguments in item["fix"]["calls"]:
            run(item["fix"]["tool"], **arguments)
    assert _compare(run, 4).data["shapes"] == []


def test_a_tables_box_height_is_not_compared_its_rows_are(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=5, columns=3, data=DATA, size=10)
    _copies(run, [[("delete_row", {"index": 4})], [("delete_row", {"index": 4})]])
    result = _compare(run, 3)
    assert result.data["shapes"] == []         # 96 pt and 76.8 pt frames, the same rows


def test_header_fill_size_widths_insets_and_borders_are_compared_with_fixes(run):
    run("ppt_add_table", slide="s:256", box=BOX, rows=4, columns=3, data=DATA[:4], size=10)
    t1, t2, t3 = _copies(run, [[], []])[0:3]
    run("ppt_edit_table", target=t2, action="format_cells", rows=[0], fill="#166183")
    run("ppt_format_text", items=[{"target": f"{t2}/row0", "size": 11},
                                  {"target": t3, "insets": {"left": 8, "top": 3.6,
                                                            "right": 7.2, "bottom": 3.6}}])
    run("ppt_edit_table", target=t3, action="set_widths", columns=[0], sizes=[214.33])
    run("ppt_edit_table", target=t3, action="format_cells",
        borders=[{"side": "all", "width": 2, "color": "#000000"}])
    entry = _compare(run, 3).data["shapes"][0]
    by = {item["property"]: item for item in entry["differ"]}
    assert set(by) == {"w", "header_h", "header_fill", "header_size", "widths", "insets",
                       "borders"}
    style_fill = by["header_fill"]["value"]          # the default style's header fill
    assert style_fill == "#156082"
    assert by["header_fill"]["near"] == [{"value": "#166183", "shapes": [t2]}]
    assert by["header_fill"]["fix"] == {"tool": "ppt_edit_table", "calls": [
        {"target": t2, "action": "format_cells", "rows": [0], "fill": style_fill}]}
    assert by["header_size"]["fix"] == {"tool": "ppt_format_text",
                                        "items": [{"target": f"{t2}/row0", "size": 10}]}
    assert by["insets"]["value"] == [7.2, 3.6, 7.2, 3.6]
    assert by["insets"]["fix"]["items"] == [{"target": t3, "insets": {
        "left": 7.2, "top": 3.6, "right": 7.2, "bottom": 3.6}}]
    assert by["widths"]["fix"]["calls"][0]["action"] == "set_widths"
    assert "fix" not in by["borders"] and by["borders"]["other"][0]["value"] == "2 #000000"
    for name in ("header_fill", "header_size", "insets", "widths", "header_h"):
        fix = by[name]["fix"]
        if "calls" in fix:
            for arguments in fix["calls"]:
                run(fix["tool"], **arguments)
        else:
            run(fix["tool"], items=fix["items"])
    left = {item["property"] for item in _compare(run, 3).data["shapes"][0]["differ"]}
    assert left == {"borders"}
