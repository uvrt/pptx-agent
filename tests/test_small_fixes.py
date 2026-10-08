"""LP20's small fixes: the preset list, unit helpers, the fit allowance, the drawn fill."""

from __future__ import annotations

import pytest

from pptx_agent import PRESETS, Cm, Document, Inches, Pt, to_pt
from pptx_agent.edit.color import Color


def test_every_preset_is_listed_with_its_adjustment_names():
    assert len(PRESETS) > 180
    assert PRESETS["rect"] == ()
    assert PRESETS["roundRect"] == ("adj",)
    assert PRESETS["rightArrow"] == ("adj1", "adj2")
    deck = Document.new()
    shape = deck.add_slide("Blank").add_shape("rightArrow", 0, 0, Inches(2), Inches(1))
    assert tuple(shape.adjustments.names) == PRESETS["rightArrow"]


def test_unit_helpers_are_emu_ints():
    assert Inches(1) == 914400 and Cm(2.54) == 914400 and Pt(72) == 914400
    assert isinstance(Inches(0.5), int) and repr(Cm(1)) == "Cm(1)"
    assert to_pt(914400) == 72.0 and to_pt(12700 * 1.234) == 1.23


def test_a_fit_states_its_allowance():
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(0, 0, Inches(3), Pt(10), "One line",
                                              autofit="none")
    fit = box.text_fit()
    assert fit.slack > 0
    assert fit.overflows == (fit.overflow > fit.slack)


def test_the_drawn_fill_of_a_new_shape_is_its_styles():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    box = slide.add_shape("rect", 0, 0, Inches(1), Inches(1))
    assert box.fill is None
    assert box.effective_fill.kind == "solid" and box.effective_fill.color == Color.parse("accent1")
    box.fill = "accent2"
    assert box.effective_fill.color == Color.parse("accent2")
    label = slide.add_textbox(0, 0, Inches(1), Inches(1), "x")
    assert label.effective_fill.kind == "none"
    title = deck.add_slide("Title Only").shapes[0]
    assert title.effective_fill is None
