"""Round 53, R109: a deck's ``check`` lists what ``save_document`` would refuse.

ooxml-edit's ``check`` carries the save gate's own comparison as ``validation``
(:mod:`ooxml_edit.tools.savegate`), whatever the check's ``slides`` or ``include``; the
deck's ``validate`` delta is read through the same function.  A made-up deck is given
three relationship ids its slides lack by a test tool, then repaired by undo."""

from __future__ import annotations

import datetime as dt

import pytest

pytest.importorskip("ooxml_edit.tools.savegate", reason="needs ooxml-edit 0.111.0")

from ooxml_edit.tools import Result, Toolbox, string, tool  # noqa: E402

from pptx_agent import Document  # noqa: E402
from pptx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 11, 12, tzinfo=dt.timezone.utc)  # noqa: E731
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


@tool("dangle", "Test only: name relationship ids the slides lack.",
      {"doc": string("Document id.")}, kind="pptx", mutates=True)
def dangle(call, doc):
    deck = call.document
    deck.history.checkpoint()
    for number, slide in enumerate(deck.slides, 1):
        root = deck.package.tree(slide.part_path)
        root[0].set(R + "embed", f"rId9{number}")
        deck.package.mark_dirty(slide.part_path)
    return Result(summary="dangled", changed=True)


@pytest.fixture
def box():
    with Toolbox([*TOOLS, dangle], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box):
    deck = Document.new()
    for title in ("Shelf", "Top", "Low"):
        deck.add_slide("Title Only").title = title
    session = box.session(clock=CLOCK)
    session.open(deck.to_bytes(), "shelf.pptx")
    return session


def _ok(result):
    assert result.ok, result.error and result.error.to_json()
    return result


def test_check_lists_what_the_save_refuses_and_agrees_with_validate(box):
    session = _session(box)
    _ok(box.dispatch(session, "dangle", {"doc": "d1"}))
    narrow = _ok(box.dispatch(session, "check", {"doc": "d1", "slides": [1],
                                                 "include": ["fit"]}))
    full = _ok(box.dispatch(session, "check", {"doc": "d1"}))
    refused = box.dispatch(session, "save_document", {"doc": "d1", "name": "o.pptx",
                                                      "format": "pptx"})
    assert refused.error.code == "refused"
    assert len(refused.error.valid_options) == 3
    assert narrow.data["validation"]["new"] == refused.error.valid_options
    assert full.data["validation"]["new"] == full.data["validate"]["new"] \
        == refused.error.valid_options
    assert "saving would refuse 3 new validation problem(s)" in narrow.summary
    _ok(box.dispatch(session, "undo", {"doc": "d1"}))
    after = _ok(box.dispatch(session, "check", {"doc": "d1", "include": ["fit"]}))
    assert after.data["validation"] == {"new": [], "baseline": 0}
    assert _ok(box.dispatch(session, "save_document", {"doc": "d1", "name": "o.pptx",
                                                       "format": "pptx"}))
