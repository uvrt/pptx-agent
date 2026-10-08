"""The text spec, and one measuring model: what measure_text measures is what is built."""

from __future__ import annotations

import pytest

from pptx_agent import (Document, ParagraphSpec, Pt, RunSpec, TextSpec, fit_box,
                        measure_text)

SPECS = {
    "plain": TextSpec.from_text("Order received and checked against the purchase order"),
    "bold-title": TextSpec([ParagraphSpec([RunSpec("Discover", bold=True, size=20)],
                                          align="center"),
                            ParagraphSpec([RunSpec("Interviews, data audit, "
                                                   "stakeholder map")], space_before=Pt(4))]),
    "mixed-runs": TextSpec([ParagraphSpec([RunSpec("Revenue "), RunSpec("4,310", bold=True,
                                                                        size=24),
                                           RunSpec(" 億円", size=12)])]),
    "bullets": TextSpec([ParagraphSpec([RunSpec("Scope", bold=True)]),
                         ParagraphSpec([RunSpec("Agree the scope with the steering "
                                                "committee")], bullet="bullet", level=1),
                         ParagraphSpec([RunSpec("Baseline the spend")], bullet="bullet",
                                       level=1),
                         ParagraphSpec([RunSpec("Pick the waves")], bullet="number")]),
    "insets-and-break": TextSpec([ParagraphSpec([RunSpec("Kick-off\vNov 2", size=11)])],
                                 insets=(Pt(2), Pt(2), Pt(2), Pt(2)), anchor="top"),
    "spacing": TextSpec([ParagraphSpec([RunSpec("Line one of a paragraph that wraps "
                                                "several times in a narrow box")],
                                       line_spacing=0.9, space_after=Pt(6)),
                         ParagraphSpec([RunSpec("Second", italic=True, font="Georgia")])]),
}
PRESETS = ["rect", "roundRect", "ellipse", "chevron", "homePlate", "textbox"]


def _build(slide, preset, width, height, spec):
    if preset == "textbox":
        return slide.add_textbox(Pt(36), Pt(36), width, height, text=spec, autofit="none")
    return slide.add_shape(preset, Pt(36), Pt(36), width, height, text=spec)


@pytest.mark.parametrize("preset", PRESETS)
@pytest.mark.parametrize("name", sorted(SPECS))
def test_the_measured_height_is_the_built_height(name, preset):
    deck = Document.new()
    slide = deck.add_slide("Blank")
    spec = SPECS[name]
    width = Pt(160)
    measured = measure_text(spec, width=width, preset=preset, deck_or_shape=slide,
                            height=Pt(100))
    shape = _build(slide, preset, width, Pt(100), spec)
    fit = shape.text_fit()
    assert measured.height == fit.needed
    assert measured.paragraph_lines == fit.lines
    # Sized as the measurement says, the built shape's text fits.
    shape.height = measured.box_height
    fit = shape.text_fit()
    assert fit.needed <= fit.available
    # Measured without a height, it is measured at the height it needs, and built at
    # that height it says the same.
    sized = measure_text(spec, width=width, preset=preset, deck_or_shape=slide)
    shape.height = sized.box_height
    fit = shape.text_fit()
    assert sized.height == fit.needed and fit.needed <= fit.available
    if preset in ("rect", "roundRect", "textbox", "ellipse"):
        # A geometry whose text area does not narrow as it grows has one answer.
        assert abs(sized.box_height - measured.box_height) <= 6350


def test_fit_box_sizes_a_box_its_text_fits():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    spec = SPECS["bullets"]
    height = fit_box(spec, Pt(200), preset="roundRect", deck_or_shape=slide)
    shape = slide.add_shape("roundRect", 0, 0, Pt(200), height, text=spec)
    fit = shape.text_fit()
    assert not fit.overflows and fit.needed <= fit.available + fit.slack


def test_measuring_like_a_placeholder_uses_its_inheritance():
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    body = slide.shapes[1]
    spec = TextSpec.from_text("\n".join(f"Point {k}" for k in range(1, 9)))
    measured = measure_text(spec, width=body.width, like=body)
    body.set_text(spec)
    assert measured.height == body.text_fit().needed
    assert measured.paragraph_lines == body.text_fit().lines


def test_bold_is_measured_bold_in_both_paths():
    deck = Document.new()
    text = "Programme Management Office"
    plain = measure_text(text, size=18, width=Pt(150), deck_or_shape=deck)
    bold = measure_text(text, size=18, width=Pt(150), deck_or_shape=deck, bold=True)
    built = measure_text(text, size=18, width=Pt(150), deck_or_shape=deck, bold=True,
                         preset="textbox")
    assert bold.widest > plain.widest
    assert built.lines == bold.lines


def test_a_spec_builds_runs_bullets_and_frame():
    deck = Document.new()
    shape = _build(deck.add_slide("Blank"), "rect", Pt(200), Pt(100), SPECS["bullets"])
    paragraphs = shape.text_frame.paragraphs
    assert [p.text for p in paragraphs] == ["Scope", "Agree the scope with the steering "
                                            "committee", "Baseline the spend", "Pick the waves"]
    assert paragraphs[0].run(0).bold is True
    assert paragraphs[1].bullet.char == "•" and paragraphs[1].level == 1
    assert paragraphs[3].bullet.kind == "number"
    box = _build(deck.add_slide("Blank"), "rect", Pt(200), Pt(100), SPECS["insets-and-break"])
    assert box.text_frame.insets == (Pt(2),) * 4 and box.text_frame.anchor == "top"
    assert box.text_frame.text == "Kick-off\vNov 2"
    mixed = _build(deck.add_slide("Blank"), "rect", Pt(200), Pt(100), SPECS["mixed-runs"])
    runs = mixed.text_frame.paragraph(0).runs
    assert [r.text for r in runs] == ["Revenue ", "4,310", " 億円"]
    assert (runs[1].bold, runs[1].size, runs[2].size) == (True, 24.0, 12.0)


def test_a_spec_on_an_existing_shape_is_one_undo_step():
    deck = Document.new()
    shape = _build(deck.add_slide("Blank"), "rect", Pt(200), Pt(100), SPECS["plain"])
    before = deck.to_bytes()
    shape.set_text(SPECS["bullets"])
    deck.undo()
    assert deck.to_bytes() == before


def test_a_spec_refuses_what_it_cannot_be():
    with pytest.raises(ValueError):
        ParagraphSpec([RunSpec("x")], bullet="star")
    with pytest.raises(TypeError):
        TextSpec.from_text("x", weight="bold")
    with pytest.raises(ValueError):
        measure_text(SPECS["plain"], width=Pt(100), preset="blob", deck_or_shape=Document.new())
