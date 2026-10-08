"""The text API: frames, paragraphs, runs, positional addresses, and mixed-formatting edits.

Every mutation is checked the same way: edit, save, reopen, read the value back.
"""

from __future__ import annotations

import pytest

from pptx_agent import Document, Pt
from pptx_agent.edit.color import Color
from pptx_agent.oxml.xml import qn


def reopen(document: Document) -> Document:
    return Document.open(document.to_bytes())


def _text_shape(document: Document):
    return next(s for s in document.slides[0].shapes if s.kind == "shape" and s.text)


def _mixed(document: Document, text: str = "Revenue grew 12% this year"):
    """A shape whose single paragraph is three runs: plain, bold, plain."""
    shape = _text_shape(document)
    frame = shape.text_frame
    frame.set_text("Revenue grew ")
    paragraph = frame.paragraph(0)
    paragraph.run(0).bold = False
    paragraph.add_run("12%", bold=True, color="accent2")
    paragraph.add_run(" this year", bold=False)
    assert frame.text == text
    return shape


# -- reading -------------------------------------------------------------------------------


def test_frame_paragraphs_and_runs_read_the_xml(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    frame = shape.text_frame
    assert frame.text.replace("\v", "\n") == shape.text
    assert len(frame) == len(frame.paragraphs) >= 1
    run = frame.paragraph(0).run(0)
    assert run.text and run.text in shape.text
    assert run.size is not None


def test_reading_never_dirties(pptx_path):
    document = Document.open(str(pptx_path))
    for slide in document.slides:
        for shape in slide.shapes:
            if shape.kind != "shape":
                continue
            for paragraph in shape.text_frame.paragraphs:
                _ = (paragraph.alignment, paragraph.level, paragraph.bullet, paragraph.space_before)
                for run in paragraph.runs:
                    _ = (run.bold, run.italic, run.underline, run.size, run.typeface, run.color)
    assert document.package.dirty_parts == frozenset()


def test_a_shape_without_text_has_no_text_frame(financial_report):
    document = Document.open(str(financial_report))
    frame = next(s for slide in document.slides for s in slide.shapes if s.kind == "graphic_frame")
    assert frame.text_frame is None and frame.paragraphs == []
    with pytest.raises(ValueError, match="cannot hold text"):
        frame.set_text("x")


# -- run formatting ------------------------------------------------------------------------


@pytest.mark.parametrize(
    "attribute, value",
    [
        ("bold", True),
        ("bold", False),
        ("italic", True),
        ("underline", True),
        ("underline", "dbl"),
        ("underline", False),
        ("strike", True),
        ("size", 27.5),
        ("typeface", "Georgia"),
    ],
)
def test_run_property_round_trips(product_page, attribute, value):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    setattr(run, attribute, value)
    assert getattr(document.resolve(run.address), attribute) == value
    assert getattr(reopen(document).resolve(run.address), attribute) == value


@pytest.mark.parametrize("attribute", ["bold", "italic", "underline", "size", "typeface", "color"])
def test_setting_none_inherits_again(product_page, attribute):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    setattr(run, attribute, None)
    assert getattr(reopen(document).resolve(run.address), attribute) is None


def test_run_colour_hex(product_page):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.color = "#c00000"
    color = reopen(document).resolve(run.address).color
    assert color == Color.rgb("C00000")
    assert color.kind == "rgb"


def test_run_colour_stays_a_theme_colour(product_page):
    """The point of the colour model: accent1 is written as a reference, never as hex."""
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.color = "accent1 lumMod=75% lumOff=25%"

    reread = reopen(document).resolve(run.address)
    assert reread.color.kind == "scheme"
    assert reread.color.value == "accent1"
    assert reread.color.transform("lumMod") == 75000
    assert reread.color.transform("lumOff") == 25000

    xml = document.package.read(document.slides[0].part_path)
    assert b'<a:schemeClr val="accent1"><a:lumMod val="75000"/><a:lumOff val="25000"/>' in xml


def test_theme_colour_keyword_form(product_page):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.color = Color.theme("accent6", lum_mod=0.6, lum_off=0.4)
    assert reopen(document).resolve(run.address).color == "accent6 lumMod=60000 lumOff=40000"


def test_format_is_one_undo_step(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()
    run = _text_shape(document).text_frame.paragraph(0).run(0)

    run.format(bold=True, italic=True, size=40, color="accent3")

    reread = reopen(document).resolve(run.address)
    assert (reread.bold, reread.italic, reread.size, reread.color) == (True, True, 40, "accent3")
    assert document.undo()
    assert document.to_bytes() == original


def test_bad_values_are_refused_before_anything_changes(product_page):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    with pytest.raises(ValueError):
        run.underline = "squiggly"
    with pytest.raises(ValueError):
        run.color = "not-a-colour"
    with pytest.raises(ValueError):
        run.size = 0
    assert not document.history.can_undo()


def test_rpr_children_stay_in_schema_order(product_page):
    """Setting colour and typeface on a run with neither must order solidFill before latin."""
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.typeface = "Georgia"
    run.color = "accent1"
    run.typeface = "Arial"

    properties = run._element().find(qn("a:rPr"))
    names = [child.tag.rpartition("}")[2] for child in properties]
    assert names.index("solidFill") < names.index("latin")


# -- paragraphs ----------------------------------------------------------------------------


@pytest.mark.parametrize("value", ["left", "center", "right", "justify", "distributed"])
def test_alignment_round_trips(product_page, value):
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)
    paragraph.alignment = value
    assert reopen(document).resolve(paragraph.address).alignment == value


def test_level_and_indents_round_trip(product_page):
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)
    paragraph.level = 2
    paragraph.margin_left = 457200
    paragraph.indent = -228600
    reread = reopen(document).resolve(paragraph.address)
    assert (reread.level, reread.margin_left, reread.indent) == (2, 457200, -228600)


def test_spacing_round_trips(product_page):
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)
    paragraph.space_before = Pt(6)
    paragraph.space_after = 158750                 # 12.5 pt, in EMU like every length
    paragraph.line_spacing = 1.5
    reread = reopen(document).resolve(paragraph.address)
    assert (reread.space_before, reread.space_after, reread.line_spacing) == (76200, 158750, 1.5)
    assert paragraph._element().find(f"{qn('a:pPr')}/{qn('a:spcAft')}/{qn('a:spcPts')}") \
        .get("val") == "1250"

    paragraph.line_spacing_points = 18
    reread = reopen(document).resolve(paragraph.address)
    assert reread.line_spacing_points == 18 and reread.line_spacing is None


def test_bullets_round_trip(product_page):
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)

    paragraph.set_bullet("–", font="Arial", color="accent1", size=0.8)
    bullet = reopen(document).resolve(paragraph.address).bullet
    assert (bullet.kind, bullet.char, bullet.font, bullet.color, bullet.size) == (
        "char", "–", "Arial", "accent1", 0.8)

    paragraph.set_numbering("romanUcPeriod", start_at=3)
    bullet = reopen(document).resolve(paragraph.address).bullet
    assert (bullet.kind, bullet.scheme, bullet.start_at) == ("number", "romanUcPeriod", 3)

    paragraph.bullet = "none"
    assert reopen(document).resolve(paragraph.address).bullet.kind == "none"

    paragraph.bullet = None
    assert reopen(document).resolve(paragraph.address).bullet is None


def test_ppr_children_stay_in_schema_order(product_page):
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)
    paragraph.set_bullet("•")
    paragraph.space_after = Pt(3)
    paragraph.line_spacing = 1.2
    paragraph.space_before = Pt(3)

    names = [c.tag.rpartition("}")[2] for c in paragraph._element().find(qn("a:pPr"))]
    order = ["lnSpc", "spcBef", "spcAft", "buChar"]
    assert [n for n in names if n in order] == order


# -- insert and delete ---------------------------------------------------------------------


def test_add_and_delete_paragraphs(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    frame = shape.text_frame
    before = frame.text

    frame.add_paragraph("appended")
    frame.add_paragraph("first", index=0)
    assert reopen(document).shape(shape.id).text_frame.text == "first\n" + before + "\nappended"

    frame.delete_paragraph(0)
    frame.paragraph(-1).delete()
    assert reopen(document).shape(shape.id).text_frame.text == before


def test_a_new_paragraph_continues_its_neighbours_formatting(product_page):
    document = Document.open(str(product_page))
    frame = _text_shape(document).text_frame
    template = frame.paragraph(0).run(-1)
    expected = (template.bold, template.size, template.color)

    new = frame.add_paragraph("continued")

    run = reopen(document).resolve(new.address + "/r0")
    assert (run.bold, run.size, run.color) == expected


def test_add_paragraph_with_formatting(product_page):
    document = Document.open(str(product_page))
    frame = _text_shape(document).text_frame
    new = frame.add_paragraph("loud", bold=True, size=30)
    run = reopen(document).resolve(new.address).run(0)
    assert (run.text, run.bold, run.size) == ("loud", True, 30)


def test_deleting_the_last_paragraph_empties_it(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    frame = shape.text_frame
    while len(frame) > 1:
        frame.delete_paragraph(0)
    frame.delete_paragraph(0)
    reread = reopen(document).shape(shape.id).text_frame
    assert len(reread) == 1 and reread.text == ""


def test_insert_and_delete_runs(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    paragraph = shape.text_frame.paragraph(0)
    text = paragraph.text

    paragraph.add_run(" (new)", italic=True)
    paragraph.add_run(">> ", index=0, bold=True)
    reread = reopen(document).resolve(paragraph.address)
    assert reread.text == ">> " + text + " (new)"
    assert reread.run(0).bold is True and reread.run(-1).italic is True

    paragraph.delete_run(0)
    paragraph.run(-1).delete()
    assert reopen(document).resolve(paragraph.address).text == text


def test_runs_stay_before_end_paragraph_properties(financial_report):
    """A run appended to a paragraph with a:endParaRPr must land before it."""
    document = Document.open(str(financial_report))
    frame = next(s for s in document.slides[0].shapes if s.kind == "shape" and s.text).text_frame
    paragraph = next(p for p in frame.paragraphs
                     if p._element().find(qn("a:endParaRPr")) is not None)
    paragraph.add_run("tail")
    assert paragraph._element()[-1].tag == qn("a:endParaRPr")


def test_run_text_round_trips(product_page):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.text = "swapped"
    assert reopen(document).resolve(run.address).text == "swapped"
    with pytest.raises(ValueError):
        run.text = "two\nlines"


def test_line_breaks_are_vertical_tabs(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    shape.set_text("one\vtwo\nthree")
    frame = reopen(document).shape(shape.id).text_frame
    assert frame.text == "one\vtwo\nthree"
    assert len(frame) == 2
    assert frame.paragraph(0)._element().find(qn("a:br")) is not None
    # Shape.text keeps E0's reading: breaks as newlines.
    assert reopen(document).shape(shape.id).text == "one\ntwo\nthree"


# -- addressing ----------------------------------------------------------------------------


def test_addresses_are_positional_and_resolve(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    run = shape.text_frame.paragraph(0).run(0)
    assert run.address == f"{shape.id}/p0/r0"
    assert document.resolve(run.address).text == run.text
    assert document.resolve(f"{shape.id}/p0").text == shape.text_frame.paragraph(0).text
    assert document.resolve(shape.id).id == shape.id


def test_an_address_is_re_resolved_on_every_call(product_page):
    """A Run names a position: insert before it and it names the run now in that position."""
    document = Document.open(str(product_page))
    paragraph = _text_shape(document).text_frame.paragraph(0)
    first = paragraph.run(0)
    original = first.text

    paragraph.add_run("inserted ", index=0)

    assert first.text == "inserted "
    assert paragraph.run(1).text == original


def test_facades_survive_undo(product_page):
    document = Document.open(str(product_page))
    run = _text_shape(document).text_frame.paragraph(0).run(0)
    run.bold = False
    run.size = 50
    document.undo()
    assert run.size != 50  # the same facade reads the restored tree
    run.size = 51
    assert reopen(document).resolve(run.address).size == 51


def test_stale_addresses_fail_loudly(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    with pytest.raises(IndexError):
        document.resolve(f"{shape.id}/p99")
    with pytest.raises(IndexError):
        document.resolve(f"{shape.id}/p0/r99")
    with pytest.raises(KeyError):
        document.resolve(f"{shape.id}/x1")


# -- set_text keeps mixed formatting -------------------------------------------------------


def test_changing_one_figure_keeps_the_sentence_formatting(product_page):
    document = Document.open(str(product_page))
    shape = _mixed(document)

    shape.set_text("Revenue grew 15% this year")

    paragraph = reopen(document).shape(shape.id).text_frame.paragraph(0)
    assert [(r.text, r.bold) for r in paragraph.runs] == [
        ("Revenue grew ", False), ("15%", True), (" this year", False)]
    assert paragraph.run(1).color == "accent2"


def test_editing_around_a_formatted_word_keeps_it(product_page):
    document = Document.open(str(product_page))
    shape = _mixed(document)

    shape.set_text("Net revenue grew 12% in fiscal 2026")

    runs = reopen(document).shape(shape.id).text_frame.paragraph(0).runs
    bold = [r.text for r in runs if r.bold]
    assert bold == ["12%"]
    assert "".join(r.text for r in runs) == "Net revenue grew 12% in fiscal 2026"


def test_unchanged_paragraphs_are_not_touched(product_page):
    from lxml import etree

    document = Document.open(str(product_page))
    shape = _text_shape(document)
    frame = shape.text_frame
    frame.set_text("alpha\nbeta\ngamma")
    frame.paragraph(1).run(0).italic = True
    before = etree.tostring(frame.paragraph(1)._element())

    frame.set_text("ALPHA\nbeta\ngamma and more")

    assert etree.tostring(frame.paragraph(1)._element()) == before
    assert reopen(document).shape(shape.id).text_frame.paragraph(1).run(0).italic is True


def test_inserting_a_paragraph_keeps_the_following_ones(product_page):
    document = Document.open(str(product_page))
    shape = _text_shape(document)
    frame = shape.text_frame
    frame.set_text("one\ntwo\nthree")
    frame.paragraph(2).run(0).bold = True

    frame.set_text("one\ninserted\ntwo\nthree")

    reread = reopen(document).shape(shape.id).text_frame
    assert reread.text == "one\ninserted\ntwo\nthree"
    assert reread.paragraph(3).run(0).bold is True


def test_retyping_everything_takes_the_first_runs_formatting(product_page):
    """The E0 behaviour, which is still right for "retype this box"."""
    document = Document.open(str(product_page))
    shape = _mixed(document)
    first = shape.text_frame.paragraph(0).run(0)
    expected = (first.bold, first.size, first.color)

    shape.set_text("Completely different words")

    runs = reopen(document).shape(shape.id).text_frame.paragraph(0).runs
    assert len(runs) == 1
    assert (runs[0].bold, runs[0].size, runs[0].color) == expected


def test_paragraph_text_setter_keeps_formatting(product_page):
    document = Document.open(str(product_page))
    shape = _mixed(document)
    paragraph = shape.text_frame.paragraph(0)

    paragraph.text = "Revenue grew 99% this year"

    assert [r.text for r in reopen(document).resolve(paragraph.address).runs if r.bold] == ["99%"]


def test_set_text_is_one_undo_step(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()
    _text_shape(document).set_text("new\nlines\vhere")
    assert document.undo()
    assert document.to_bytes() == original


def test_segment_recuts_runs_and_keeps_text_and_formatting(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape" and len(s.text) > 6)
    paragraph = shape.text_frame.paragraph(0)
    paragraph.run(0).bold = True
    text = paragraph.text

    paragraph.segment([text[:3], text[3:]])

    assert [run.text for run in paragraph.runs] == [text[:3], text[3:]]
    assert all(run.bold for run in paragraph.runs)
    paragraph.run(1).bold = False
    paragraph.segment([text])  # merging keeps the first piece's formatting
    assert [run.text for run in paragraph.runs] == [text]
    assert paragraph.run(0).bold is True


def test_segment_refuses_a_text_change(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape" and s.text)
    with pytest.raises(ValueError):
        shape.text_frame.paragraph(0).segment(["something else"])


def test_segment_makes_and_keeps_line_breaks(product_page):
    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape" and s.text)
    frame = shape.text_frame
    frame.set_text("ab\vcd")
    paragraph = frame.paragraph(0)
    assert paragraph._pieces() == ["ab", "\v", "cd"]

    paragraph.segment(["a", "b", "\v", "cd"])

    assert paragraph._pieces() == ["a", "b", "\v", "cd"]
    assert frame.text == "ab\vcd"
