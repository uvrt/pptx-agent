"""Does the text fit, as PowerPoint shows it (the trial's findings 2 and 9)?

P3 run 2 checked fit on pptx2svg's render, which shrank an overflowing ``normAutofit`` body
that stores no ``fontScale``; PowerPoint draws it full size, and the slide overflowed.
``text_fit`` lays the text out at the stored scale, with the measurement pptx2svg draws
with, and the oracle test holds its verdicts to PowerPoint's PDF.
"""

from __future__ import annotations

import os
import warnings
from pathlib import Path

import pytest

import oracle as oracle_helper
from conftest import FIXTURE_DIR
from pptx_agent import Document, Overflow, TextFit, measure_text

pytest.importorskip("pptx2svg")

PILOT = FIXTURE_DIR / "generated" / "trial" / "pilot-retrospective.pptx"
SPLIT = """\
# Lessons learned (1 of 2)

- Planning
  - The scope was agreed two weeks after the pilot started
  - Estimates for data migration were 40% too low
  - Weekly planning sessions worked well once they started
- Delivery
  - Releasing every two weeks kept stakeholders engaged
  - Manual testing slowed down the last three releases
"""


def _pilot() -> Document:
    return Document.open(PILOT)


def test_the_trial_slide_overflows_at_full_size():
    """normAutofit with no stored fontScale: PowerPoint draws 28 and 24 pt, and it does not
    fit -- whatever a renderer that shrinks it shows."""
    fit = _pilot().shape("258.3").text_fit()
    assert isinstance(fit, TextFit)
    assert (fit.autofit, fit.font_scale) == ("normal", 1.0)
    assert fit.overflows and fit.overflow == fit.needed - fit.available > 914400
    assert fit.sizes[0] == (28.0,) and fit.sizes[1] == (24.0,) and len(fit.sizes) == 20
    assert fit.smallest_size == 24.0
    assert fit.lines == (1,) * 20
    title = _pilot().shape("258.2").text_fit()
    assert not title.overflows and title.overflow == 0 and title.sizes == ((44.0,),)


def test_a_stored_scale_is_what_is_measured():
    deck = _pilot()
    full = deck.shape("258.3").text_fit()
    frame = deck.shape("258.3").text_frame
    frame.font_scale = 0.5
    fit = deck.shape("258.3").text_fit()
    assert fit.font_scale == 0.5 and fit.sizes[:2] == ((14.0,), (12.0,))
    # Spacing in points does not scale with the font (lnSpcReduction is what tightens it),
    # so half the size is not half the height.
    assert full.needed * 0.5 < fit.needed < full.needed * 0.7


def test_the_split_slide_fits():
    deck = Document.new()
    (slide,) = deck.insert_outline(SPLIT)
    body = slide.shapes[1]
    fit = body.text_fit()
    assert not fit.overflows and fit.needed < fit.available
    assert fit.smallest_size >= 18
    assert deck.overflows() == []


def test_overflows_lists_text_overlap_and_off_slide():
    deck = _pilot()
    assert [(o.kind, o.shape) for o in deck.overflows()] == [("text", "258.3")]
    (problem,) = deck.overflows(slides=["s:258"])
    assert isinstance(problem, Overflow) and problem.fit.overflows
    assert str(problem).startswith("[text] s:258 258.3:")

    slide = deck.slide(256)
    title = slide.shapes[0]
    # Over the title's text -- centred at the bottom of its frame -- not just its frame.
    box = slide.add_textbox(title.left + title.width // 2 - 457200,
                            title.top + title.height - 457200, 914400, 457200, "over the title")
    clear = slide.add_textbox(title.left, title.top, 914400, 457200, "above the title")
    off = slide.add_shape("rect", deck.slide_size[0] - 457200, 0, 914400, 457200)
    found = {(o.kind, o.shape, o.other) for o in deck.overflows(slides=[slide])}
    assert ("overlap", box.id, title.id) in found
    assert not any(o.kind == "overlap" and clear.id in (o.shape, o.other) for o in
                   deck.overflows(slides=[slide]))
    assert ("off_slide", off.id, None) in found


def test_text_fit_needs_text():
    deck = _pilot()
    slide = deck.slides[0]
    picture = slide.add_picture(_png(), 0, 0)
    with pytest.raises(ValueError):
        picture.text_fit()
    empty = slide.add_shape("rect", 0, 0, 914400, 914400)
    assert empty.text_fit().needed == 0 and not empty.text_fit().overflows


def _png() -> bytes:
    import struct
    import zlib

    raw = b"\x00\xff\x00\x00"
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(
            ">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


requires_powerpoint = pytest.mark.skipif(
    not oracle_helper.available(),
    reason="needs macOS with Microsoft PowerPoint and the pptx2svg oracle script",
)


def _probe_deck() -> tuple[Document, list[tuple[int, str]]]:
    """A slide per case, each with one text body: the trial's slide 3 and the same body
    with 10 and 11 of its paragraphs (the last to fit and the first not to), and a 20 pt
    paragraph in boxes half a line short, exactly as tall as measured, and a little taller."""
    deck = _pilot()
    source = deck.shape("258.3").text_frame
    lines = source.text.split("\n")
    levels = [paragraph.level for paragraph in source.paragraphs]
    cases = [(258, "258.3")]
    for count in (10, 11):
        copy = deck.slide(258).duplicate(index=len(deck.slides))
        body = copy.shapes[1]
        body.set_text("\n".join(lines[:count]))
        for paragraph, level in zip(body.text_frame.paragraphs, levels):
            paragraph.level = level
        cases.append((copy.slide_id, body.id))
    words = ("The shared test environment was often down and monitoring dashboards "
             "arrived too late to help anyone")
    probe = deck.add_slide("Blank").add_textbox(914400, 914400, 4572000, 100000, words)
    probe.text_frame.format(size=20)
    measured = probe.text_fit()
    per_line = measured.needed / sum(measured.lines)
    deck.delete_slide(probe._slide)
    for extra in (-0.5, 0.0, 0.3):
        slide = deck.add_slide("Blank")
        box = slide.add_textbox(914400, 914400, 4572000,
                                round(per_line * (sum(measured.lines) + extra)), words)
        box.text_frame.format(size=20)
        cases.append((slide.slide_id, box.id))
    for slide in list(deck.slides):
        if slide.slide_id not in {case[0] for case in cases}:
            deck.delete_slide(slide)
    return deck, cases


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_agrees_with_text_fit():
    """Each case's verdict against PowerPoint's PDF: the text overflows when its lowest ink
    is below the frame's bottom edge."""
    pypdfium2 = pytest.importorskip("pypdfium2")
    deck, cases = _probe_deck()
    verdicts = [deck.shape(shape).text_fit().overflows for _, shape in cases]
    assert verdicts == [True, False, True, True, False, False]
    home = Path(os.path.expanduser("~"))
    path, pdf = home / "pptx-agent-fit.pptx", home / "pptx-agent-fit.pdf"
    deck.save(path)
    try:
        result = oracle_helper.export_pdf(path, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        document = pypdfium2.PdfDocument(str(pdf))
        try:
            drawn = []
            for (_, shape_id), page in zip(cases, document):
                shape = deck.shape(shape_id)
                height = page.get_height()
                scale = height / (deck.slide_size[1] / 12700)
                text = page.get_textpage()
                lows = [text.get_charbox(i)[1] for i in range(text.count_chars())
                        if text.get_text_range(i, 1).strip()]
                bottom = height - (shape.top + shape.height) / 12700 * scale
                drawn.append(min(lows) < bottom - 0.5)
        finally:
            document.close()
        assert drawn == verdicts
    finally:
        oracle_helper.cleanup(path, pdf)


# -- the wrap boundary (trial 2, N1) -----------------------------------------------------------


def _label(text: str, width: int, font: str = "Aptos", size: float = 18):
    deck = Document.new()
    box = deck.add_slide("Blank").add_textbox(7487000, 2575000, width, 450000, text)
    box.text_frame.format(size=size, typeface=font)
    return box


def test_the_trial_pass_label_breaks_where_powerpoint_breaks_it():
    """p4 run 2: "Pass" at 18 pt Aptos in a 650,000 EMU box.  PowerPoint broke it
    "Pas / s": it kerns with Aptos's legacy ``kern`` table, which lacks the ``ss`` pair
    only GPOS holds (-33/2048 em), and so does the measurement now -- the word is 467,804
    EMU, not 464,121 (tools/wrap_boundary_probe.py)."""
    from pptx_agent.edit.fit import WRAP_MARGIN

    fit = _label("Pass", 650000).text_fit()
    assert fit.lines == (2,)
    assert measure_text("Pass", size=18, width=650000, font="Aptos").lines == ("Pas", "s")
    close = _label("Pass", 652000).text_fit()
    assert close.lines == (1,) and not close.overflows
    assert 0 < close.margin_to_wrap < WRAP_MARGIN == 2032 and close.near_wrap
    assert abs(close.margin_to_wrap - (652000 - 2 * 91440 - 467804)) <= 2
    roomy = _label("Pass", 652000 + 2 * 12700).text_fit()
    assert roomy.margin_to_wrap > WRAP_MARGIN and not roomy.near_wrap


def test_the_margin_is_the_tightest_line_and_none_without_wrapping():
    box = _label("Order received and checked", 1371600)
    fit = box.text_fit()
    assert fit.lines[0] >= 2 and fit.margin_to_wrap is not None
    box.text_frame.wrap = False
    assert box.text_fit().margin_to_wrap is None and not box.text_fit().near_wrap
    empty = _label("", 1371600)
    assert empty.text_fit().margin_to_wrap is None
