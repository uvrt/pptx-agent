"""A slide's title, and a slide by its title (the trial's N14)."""

from __future__ import annotations

import pytest

from pptx_agent import Document, LabelError


def _deck():
    deck = Document.new()
    cover = deck.add_slide("Title Slide")
    body = deck.add_slide("Title and Content")
    blank = deck.add_slide("Blank")
    return deck, cover, body, blank


def test_a_title_is_read_and_written_through_its_placeholder():
    deck, cover, body, blank = _deck()
    assert cover.title == "" and blank.title is None
    cover.title = "Plan"                                 # a title slide's ctrTitle
    body.title = "Budget"
    assert (cover.title, body.title) == ("Plan", "Budget")
    assert body.shapes[0].text == "Budget"
    with pytest.raises(ValueError, match="no title placeholder"):
        blank.title = "Nothing to hold it"


def test_a_slide_is_found_by_its_title():
    deck, cover, body, blank = _deck()
    cover.title = "Plan"
    body.title = "Budget\vreview"                        # a line break in the title
    assert deck.slide_titled("Budget review").slide_id == body.slide_id
    assert deck.slide_titled("budget REVIEW").slide_id == body.slide_id
    with pytest.raises(LabelError, match="'Plan'"):
        deck.slide_titled("Risks")
    deck.add_slide("Title Only").title = "Plan"
    with pytest.raises(LabelError, match="2"):
        deck.slide_titled("Plan")
