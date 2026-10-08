"""Measuring text before building (the trial's "measure call"), and the height a box needs.

p7 run 1 and both p8 runs went round build -> overflows() -> resize five or six times
because nothing measured text before a shape existed.
"""

from __future__ import annotations

import pytest

from pptx_agent import Document, TextMeasure, measure_text

pytest.importorskip("pptx2svg")

EMU = 914400
TEXTS = ["Order received", "Pass", "Signed wave-1 savings by mid-February, with four SteerCo "
         "decisions keeping the programme on track", "Line one\nA second paragraph"]


@pytest.mark.parametrize("text", TEXTS)
@pytest.mark.parametrize("font, size, width", [("Aptos", 18, 1371600), ("Calibri", 14, 2743200),
                                               ("Arial", 24, 4572000)])
def test_measure_text_is_text_fit_before_the_shape_exists(text, font, size, width):
    measured = measure_text(text, font=font, size=size, width=width)
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, width, EMU, text)
    box.text_frame.format(size=size, typeface=font)
    fit = box.text_fit()
    assert isinstance(measured, TextMeasure)
    assert measured.height == fit.needed
    assert measured.line_count == sum(fit.lines)
    assert measured.margin_to_wrap == fit.margin_to_wrap


def test_lines_and_breaks():
    text = "Signed wave-1 savings by mid-February"
    measured = measure_text(text, font="Aptos", size=18, width=1828800)
    assert len(measured.lines) > 1
    for line, start in zip(measured.lines, measured.breaks):
        assert text[start:start + len(line)] == line
    assert " ".join(measured.lines) == text
    two = measure_text("One\nTwo", font="Aptos", size=18, width=EMU * 4)
    assert two.lines == ("One", "Two") and two.breaks == (0, 4)
    broken = measure_text("One\vTwo", font="Aptos", size=18, width=EMU * 4)
    assert broken.lines == ("One", "Two") and broken.height == two.height
    # An empty line between two breaks starts after the first, at the second.
    gap = measure_text("a\v\vb", font="Aptos", size=18, width=EMU * 4)
    assert gap.lines == ("a", "", "b") and gap.breaks == (0, 2, 3)


def test_the_deck_gives_its_body_font_and_a_shape_its_formatting():
    deck = Document.new()
    assert measure_text("Plan", size=18, width=EMU, deck_or_shape=deck) == \
        measure_text("Plan", size=18, width=EMU, font=deck.theme.fonts.minor)
    box = deck.add_slide("Blank").add_textbox(EMU, EMU, 2 * EMU, EMU, "Old")
    box.text_frame.format(size=24, typeface="Arial")
    box.text_frame.insets = (0, 0, 0, 0)
    box.text_frame.paragraph(0).line_spacing = 1.5
    new = "A longer text that will need more than one line"
    before = measure_text(new, width=box.width, deck_or_shape=box)
    box.set_text(new)
    fit = box.text_fit()
    assert (before.height, before.line_count) == (fit.needed, sum(fit.lines))


def test_line_spacing_and_insets_count():
    single = measure_text("One\nTwo", font="Aptos", size=18, width=4 * EMU)
    double = measure_text("One\nTwo", font="Aptos", size=18, width=4 * EMU, line_spacing=2.0)
    tight = measure_text("One\nTwo", font="Aptos", size=18, width=4 * EMU, insets=(0, 0, 0, 0))
    assert double.height - 2 * 45720 == 2 * (single.height - 2 * 45720)
    assert tight.height == single.height - 2 * 45720


def test_what_measure_text_needs():
    with pytest.raises(ValueError, match="font"):
        measure_text("Plan", size=18, width=EMU)
    with pytest.raises(ValueError, match="size"):
        measure_text("Plan", font="Aptos", width=EMU)
    with pytest.raises(TypeError):
        measure_text("Plan", size=18, width=EMU, deck_or_shape="deck")


def test_fit_height_is_the_height_the_text_needs():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    box = slide.add_textbox(EMU, EMU, 2 * EMU, 100000, TEXTS[2])
    box.height = box.fit_height()
    fit = box.text_fit()
    assert fit.needed == fit.available and not fit.overflows
    oval = slide.add_shape("ellipse", EMU, 3 * EMU, 3 * EMU, 200000, text=TEXTS[2])
    oval.height = oval.fit_height()
    fit = oval.text_fit()
    assert not fit.overflows and abs(fit.available - fit.needed) <= 0.02 * fit.needed
