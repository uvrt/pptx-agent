"""Slide.content_area: the layout's body, or the band under the title."""

from __future__ import annotations

from pptx_agent import Document


def test_a_layout_with_a_body_gives_its_box():
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    body = deck.layout("Title and Content").placeholders[1]
    assert slide.content_area == tuple(body.bounds)


def test_two_content_placeholders_give_their_union():
    deck = Document.new()
    slide = deck.add_slide("Two Content")
    boxes = [p.bounds for p in deck.layout("Two Content").placeholders if p.type is None]
    left, top, width, height = slide.content_area
    assert left == min(b[0] for b in boxes) and left + width == max(b[0] + b[2] for b in boxes)


def test_title_only_is_the_band_under_the_title():
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    title = deck.layout("Title Only").placeholders[0].bounds
    left, top, width, height = slide.content_area
    assert (left, width) == (title[0], title[2])
    assert top == title[1] + title[3] + 137160
    assert top + height <= deck.slide_size[1] - 365760


def test_blank_has_a_margin_all_round():
    deck = Document.new()
    left, top, width, height = deck.add_slide("Blank").content_area
    assert left == 365760 and top == 1097280 + 137160
    assert left + width == deck.slide_size[0] - 365760
