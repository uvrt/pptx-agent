"""Numbers no value holds, on a deck (round 36): ``check``'s ``include: ["numbers"]`` reads
each shape's and table cell's own text (``FORMAT.number_texts``) and compares its numbers
with the session's datasets and the tables of the open documents that list them.

The data is made up: the toy dataset of :mod:`test_tables_from_data` and a grid whose
tables are listed as values, each column with a number format.
"""

from __future__ import annotations

import pytest

from ooxml_common.locales import Locale
from ooxml_edit.tools import Dataset, DocumentFormat, NumberRules, Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import (CLOCK, COLUMNS, ROWS, Grid, _grid_table, call, ds_put,
                                   failed, ok)

NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")
BOX = {"x": 40, "y": 40, "w": 600, "h": 80}
FORMATS = (None, "0.0", None, None, None)


def _formatted(grid, address):
    found = _grid_table(grid, address)
    return Dataset(found.columns, found.rows, sources=found.sources, keys=found.keys,
                   formats=FORMATS)


GRID = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _formatted},
                      value_sources=lambda grid: [f"t:{name}" for name in grid.tables])


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox, **options):
    session = toolbox.session(clock=CLOCK, locale=NL_BE, **options)
    deck = Document.new()
    deck.add_slide("Blank")
    deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")                                  # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")       # d2
    return session


def box(toolbox, session, text, slide="s:256"):
    return ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide=slide, items=[
        {"preset": "textbox", "box": BOX, "text": text}])).created[0]


def check(toolbox, session, **arguments):
    return ok(call(toolbox, session, "check", doc="d1", include=["numbers"], **arguments))


def test_typed_and_unverified_numbers_are_listed_per_shape(toolbox):
    session = _session(toolbox)
    shape = box(toolbox, session, "Alpha rated 3,5/5, Gamma 1,3 and Delta 2,9")
    result = check(toolbox, session)
    data = result.data["numbers"]
    # nothing else asked: nothing else run, but what saving would refuse (round 53)
    assert set(result.data) == {"numbers", "validation"}
    (text,) = data["texts"]
    assert (text["at"], text["slide"]) == (shape, "s:256")
    assert text["unverified"] == [{"text": "2,9", "near": [
        {"value": "3,5", "from": "t:checks/Alpha/rating", "sources": ["G!B2"]},
        {"value": "4,0", "from": "t:checks/Beta/rating", "sources": ["G!B3"]}]}]
    rated, gamma = text["typed"]
    assert rated["from"] == "t:checks/Alpha/rating"
    assert rated["fix"] == {"tool": "replace_text", "arguments": {
        "doc": "d1", "find": "Alpha rated 3,5/5, Gamma 1,3", "replace":
            "Alpha rated {{t:checks/Alpha/rating:0.0}}/5, Gamma 1,3", "expect": "one",
        "slides": ["s:256"]}}
    assert gamma["from"] == "t:checks/Gamma/rating"
    assert data["left_aside"] == {"whole": 1}
    assert result.summary.startswith("1 unverified and 2 typed number(s) in 1 text(s)")


def test_the_fix_writes_the_value_by_reference_and_it_is_then_left_out(toolbox):
    session = _session(toolbox)
    shape = box(toolbox, session, "Alpha rated 3,5/5")
    fix = check(toolbox, session).data["numbers"]["texts"][0]["typed"][0]["fix"]
    ok(call(toolbox, session, fix["tool"], **fix["arguments"]))
    deck = session.entry("d1").document
    assert deck.shape(shape).text_frame.text == "Alpha rated 3,5/5"
    assert deck.provenance()[shape]["facts"][0]["source"] == "t:checks/Alpha/rating"
    data = check(toolbox, session).data["numbers"]
    assert data["texts"] == [] and data["typed"] == 0


def test_several_fixes_go_in_one_batch(toolbox):
    session = _session(toolbox)
    shape = box(toolbox, session, "Alpha 3,5 and Gamma 1,3 then Beta 4,0")
    data = check(toolbox, session).data["numbers"]
    fixes = [entry["fix"] for entry in data["texts"][0]["typed"]]
    assert data["fix_hint"].startswith("3 fixes: send them as one batch")
    ok(call(toolbox, session, "batch", ops=[{"tool": fix["tool"], "arguments": fix["arguments"]}
                                           for fix in fixes]))
    deck = session.entry("d1").document
    assert deck.shape(shape).text_frame.text == "Alpha 3,5 and Gamma 1,3 then Beta 4,0"
    assert [f["source"] for f in deck.provenance()[shape]["facts"]] == [
        "t:checks/Alpha/rating", "t:checks/Gamma/rating", "t:checks/Beta/rating"]
    assert check(toolbox, session).data["numbers"]["typed"] == 0


def test_a_dataset_is_compared_too_and_the_value_its_table_holds_counts_once(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    box(toolbox, session, "Gamma 1,25")
    typed = check(toolbox, session).data["numbers"]["texts"][0]["typed"]
    assert typed == [{"text": "1,25", "from": "d:scores/Gamma/rating", "sources": ["G!B4"],
                      "fix": typed[0]["fix"]}]
    assert typed[0]["fix"]["arguments"]["replace"] == "Gamma {{d:scores/Gamma/rating:0.00}}"


def test_text_by_reference_and_a_table_placed_from_data_are_left_out(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    box(toolbox, session, "Gamma {{d:scores/Gamma/rating}}")
    ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
            **{"from": "d:scores", "from_columns": ["check", "rating"]}))
    assert check(toolbox, session).data["numbers"]["texts"] == []
    typed = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    rows=1, columns=2, data=[["Delta", "2,9"]])).created[0]
    (text,) = check(toolbox, session).data["numbers"]["texts"]
    assert text["at"] == f"{typed}/cell0,1"


def test_slides_scope_the_check_and_pages_or_range_are_refused(toolbox):
    session = _session(toolbox)
    box(toolbox, session, "Delta 2,9", slide="s:256")
    box(toolbox, session, "Delta 2,8", slide="s:257")
    data = check(toolbox, session, slides=["s:257"]).data["numbers"]
    assert [t["slide"] for t in data["texts"]] == ["s:257"]
    failed(call(toolbox, session, "check", doc="d1", include=["numbers"], pages=[1]),
           "invalid_arguments")
    failed(call(toolbox, session, "check", doc="d1", include=["numbers"], range="p:1"),
           "invalid_arguments")


def test_other_facts_come_with_numbers_and_numbers_never_come_by_default(toolbox):
    session = _session(toolbox)
    box(toolbox, session, "Delta 2,9")
    both = ok(call(toolbox, session, "check", doc="d1", include=["numbers", "fit"]))
    assert "numbers" in both.data and "overflows" in both.data
    assert "numbers" not in ok(call(toolbox, session, "check", doc="d1")).data


def test_the_applications_ignore_list(toolbox):
    session = _session(toolbox, numbers=NumberRules(ignore=("ISO 9001.5",)))
    box(toolbox, session, "Per ISO 9001.5, Delta 2,9")
    data = check(toolbox, session).data["numbers"]
    assert data["left_aside"] == {"ignored": 1}
    assert [e["text"] for e in data["texts"][0]["unverified"]] == ["2,9"]


def test_decks_are_offered_numbers():
    sent = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0)
    try:
        include = next(d for d in sent.definitions() if d["name"] == "check")
        enum = include["input_schema"]["properties"]["include"]["items"]["enum"]
        assert "numbers" in enum
    finally:
        sent.close()
