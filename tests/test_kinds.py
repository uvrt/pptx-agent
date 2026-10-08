"""A property a shape's kind does not have reads None, not an error (the trial's N13).

Both p5 agents walked every shape reading ``text_frame`` and ``line`` and hit ValueError
on a picture and a graphic frame.
"""

from __future__ import annotations

from conftest import FIXTURE_DIR
from pptx_agent import Document

EMU = 914400


def test_every_shape_can_be_read_without_asking_its_kind():
    for path in sorted(FIXTURE_DIR.glob("*.pptx")):
        deck = Document.open(path)
        for slide in deck.slides:
            for shape in slide.shapes:
                (shape.text_frame, shape.line, shape.table, shape.chart, shape.diagram,
                 shape.paragraphs, shape.text, shape.image_size, shape.route)


def test_what_each_kind_reads():
    deck = Document.open(FIXTURE_DIR / "real-financial-report.pptx")
    table = deck.shape("257.3#5")
    assert table.kind == "graphic_frame" and table.table is not None
    assert (table.text_frame, table.line, table.chart, table.diagram) == (None, None, None, None)
    chart = deck.shape("257.25")
    assert chart.chart is not None and chart.table is None
    fresh = Document.new()
    slide = fresh.add_slide("Blank")
    box = slide.add_shape("rect", EMU, EMU, EMU, EMU, text="x")
    assert box.text_frame is not None and box.table is None and box.route is None
    group = slide.group([box, slide.add_shape("rect", 3 * EMU, EMU, EMU, EMU)])
    assert group.line is None and group.text_frame is None and group.paragraphs == []
