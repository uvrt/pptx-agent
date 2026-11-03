"""Round 28, P2: ``check`` measures in a worker process on a snapshot of the deck, with the
deck's lock let go, so checks of one deck run side by side and an edit does not wait.

What it reports must be what the check on the deck itself reports (inside a batch it still
runs there, under the lock), for the same version, every kind of fact asked for."""

from __future__ import annotations

import json

import pytest

from ooxml_edit.tools import Limits, Toolbox

from pptx_agent import Document
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from pptx_agent.tools.shared_tools import CHECK_SLIDES, check_timeout

pytest.importorskip("pptx2svg")

EMU = 12700
#: Every kind of fact a check reports on a deck.
EVERYTHING = {"include": ["fit", "collisions", "validate", "facts", "design", "layout"],
              "boxes": True, "spacing": 6, "contrast": 4.5}


def _deck() -> bytes:
    deck = Document.new()
    for number in range(1, 5):
        slide = deck.add_slide("Title Only")
        slide.title = f"Part {number}: routes and margins for the coming quarter"
        table = slide.add_table(4, 3, 40 * EMU, 110 * EMU, 300 * EMU, 120 * EMU)
        for row in range(4):
            for col in range(3):
                table.table.cell(row, col).text_frame.text = (
                    "Route" if row == 0 else f"{row * 1234 + col:,} crew north ridge")
        slide.add_textbox(300 * EMU, 100 * EMU, 200 * EMU, 40 * EMU,
                          "A box over the table, with a long line of text that will not fit")
        slide.add_textbox(40 * EMU, 400 * EMU, 400 * EMU, 30 * EMU, "\n\nSmall print")
    return deck.to_bytes()


@pytest.fixture(scope="module")
def deck_bytes() -> bytes:
    return _deck()


@pytest.fixture(scope="module", params=[0, 2], ids=["in-process", "workers"])
def toolbox(request):
    with Toolbox(TOOLS, formats=[deck_format(min_text_size=12)], groups=GROUPS, font_dirs=[],
                 workers=request.param, limits=Limits(image_budget=None)) as box:
        yield box


def _plain(result) -> dict:
    return json.loads(json.dumps(result.to_json(), default=str))


@pytest.mark.parametrize("arguments", [{"slides": [1, 2]}, {"slides": [3, 4], **EVERYTHING}],
                         ids=["default", "everything"])
def test_a_check_on_a_snapshot_reports_what_the_check_on_the_deck_does(toolbox, deck_bytes,
                                                                        arguments):
    session = toolbox.session()
    doc = session.open(deck_bytes, "deck.pptx")
    alone = toolbox.dispatch(session, "check", {"doc": doc, **arguments})
    # Inside a batch the check runs on the deck under the lock: the reference.  Another
    # session, so neither reads the other's cache.
    other = toolbox.session()
    second = other.open(deck_bytes, "deck.pptx")
    batched = toolbox.dispatch(other, "batch", {"ops": [
        {"tool": "check", "arguments": {"doc": second, **arguments}}]})
    assert alone.ok and batched.ok, (alone.to_json(), batched.to_json())
    op = batched.data["ops"][0] if "ops" in (batched.data or {}) else None
    reference = op["data"] if op and "data" in op else batched.data
    assert _plain(alone)["data"] == json.loads(json.dumps(reference, default=str))
    assert alone.version == session.entry(doc).version
    assert alone.warnings == []


def test_checks_of_one_deck_in_one_turn_all_come_back_of_one_version(toolbox, deck_bytes):
    session = toolbox.session()
    doc = session.open(deck_bytes, "deck.pptx")
    version = session.entry(doc).version
    results = toolbox.dispatch_many(session, [("check", {"doc": doc, "slides": [n]})
                                              for n in (1, 2, 3, 4)])
    assert all(r.ok for r in results), [r.to_json() for r in results]
    assert [r.data["slides"] for r in results] == [
        [f"s:{s.slide_id}"] for s in session.entry(doc).document.slides]
    assert {r.version for r in results} == {version}
    assert all(r.warnings == [] for r in results)


def test_an_edit_after_a_check_in_one_turn_leaves_the_check_of_the_deck_before(toolbox,
                                                                               deck_bytes):
    session = toolbox.session()
    doc = session.open(deck_bytes, "deck.pptx")
    entry = session.entry(doc)
    version = entry.version
    title = entry.document.slides[0].shapes[0].id
    check, edit = toolbox.dispatch_many(session, [
        ("check", {"doc": doc, "slides": [1]}),
        ("ppt_set_text", {"doc": doc, "items": [
            {"target": title, "text": "Short"}]})])
    assert check.ok and edit.ok, (check.to_json(), edit.to_json())
    assert check.version == version and edit.version == version + 1
    # The turn's own later edit moved the deck: its order says so, no warning.
    assert check.warnings == []


def test_a_checks_deadline_grows_with_its_slides():
    limits = Limits()
    assert check_timeout(limits, 1) == limits.layout_timeout
    assert check_timeout(limits, CHECK_SLIDES) == limits.layout_timeout
    assert check_timeout(limits, CHECK_SLIDES + 1) == 2 * limits.layout_timeout
