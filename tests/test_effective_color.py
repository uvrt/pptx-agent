"""Effective text colour and inherited text sizes (production feedback round 5, T12).

``Run.effective_color`` resolves a run's colour through the chain ``effective_size`` reads
(run, paragraph, list styles, layout and master placeholders, the master's text styles, the
presentation's default) plus a plain shape's ``fontRef`` and a table cell's table style,
through the slide's colour map and theme.  ``Run.effective_source`` says where a size or
colour comes from, ``Slide.text_below`` lists text under a minimum, and the tool layer's
``deck_format(min_text_size=...)`` makes ``check`` report it.  Every deck here is made up.
"""

from __future__ import annotations

import datetime as dt

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Color, Document
from pptx_agent.edit.tablestyles import styles_part
from pptx_agent.tools import FORMAT, GROUPS, TOOLS, deck_format

EMU = 914400
CLOCK = lambda: dt.datetime(2027, 1, 4, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
CONTENT = ('<p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>'
           '<a:lstStyle/>')
#: The layout's content placeholder: 9 pt in accent 2 at the first level.
SMALL_LEVEL = ('<a:lstStyle><a:lvl1pPr><a:defRPr sz="900"><a:solidFill>'
               '<a:schemeClr val="accent2"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>')
#: A slide's colour map with text and background swapped (a dark slide).
DARK = ('<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
        'accent6="accent6" hlink="hlink" folHlink="folHlink"/>')


def _patch(deck: Document, part: str, old: str, new: str) -> None:
    xml = deck.package.read(part).decode("utf-8")
    assert old in xml, f"{part}: {old!r} not found"
    deck.package.replace_part(part, xml.replace(old, new, 1).encode("utf-8"))


def small_layout_deck() -> Document:
    """A new deck whose 'Title and Content' layout sets its content placeholder to 9 pt
    accent 2, with one slide on it whose text says neither."""
    deck = Document.new(size="16:9")
    _patch(deck, deck.layout("Title and Content").part_path, CONTENT,
           CONTENT.replace("<a:lstStyle/>", SMALL_LEVEL))
    deck = Document.open(deck.to_bytes())
    slide = deck.add_slide("Title and Content")
    slide.shapes[0].set_text("A title")
    slide.shapes[1].set_text("Inherited body")
    return deck


def first_run(shape):
    return shape.paragraphs[0].runs[0]


# -- sizes and where they come from ------------------------------------------------------------


def test_a_placeholder_inherits_its_size_and_colour_from_the_layout():
    deck = small_layout_deck()
    slide = deck.slides[0]
    run = first_run(slide.shapes[1])
    assert (run.size, run.color) == (None, None)
    assert run.effective_size == 9.0
    assert run.effective_source() == "layout 'Title and Content' content placeholder"
    assert run.effective_color == Color.parse("accent2").resolve(slide)
    assert run.effective_color.kind == "rgb"
    assert run.effective_source("color") == "layout 'Title and Content' content placeholder"
    # The title says nothing either: the master's title style.
    title = first_run(slide.shapes[0])
    assert title.effective_source() == "the master's titleStyle"
    assert title.effective_source("color") == "the master's titleStyle"


def test_a_size_of_its_own_is_the_runs_and_an_autofit_scale_is_named():
    deck = small_layout_deck()
    body = deck.slides[0].shapes[1]
    run = first_run(body)
    run.size = 14
    assert run.effective_source() == "the run"
    body.text_frame.set_autofit("normal", font_scale=0.5)
    assert run.effective_size == 7.0
    assert run.effective_source() == "the run at 50% autofit"


def test_effective_source_names_size_or_color():
    run = first_run(small_layout_deck().slides[0].shapes[1])
    with pytest.raises(ValueError, match="'size' or 'color'"):
        run.effective_source("font")


def test_text_below_lists_inherited_small_text_with_its_source():
    deck = small_layout_deck()
    slide = deck.slides[0]
    found = slide.text_below(12)
    assert [(t.address, t.size, t.source) for t in found] == [
        (slide.shapes[1].id, 9.0, "layout 'Title and Content' content placeholder")]
    assert found[0].to_json() == {"shape": slide.shapes[1].id, "size": 9.0,
                                  "from": "layout 'Title and Content' content placeholder"}
    assert slide.text_below(9) == []


def test_text_below_looks_into_table_cells():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    table = slide.add_table(2, 2, EMU, EMU, 4 * EMU, EMU).table
    table.cell(1, 1).text = "small"
    first_run(table.cell(1, 1).text_frame).size = 8
    found = slide.text_below(10)
    assert [(t.address, t.size, t.source) for t in found] == [
        (table.cell(1, 1).address, 8.0, "the run")]


# -- colours -------------------------------------------------------------------------------


def test_a_theme_colour_goes_through_the_slides_colour_map():
    deck = small_layout_deck()
    slide = deck.slides[0]
    _patch(deck, slide.part_path, "<a:masterClrMapping/>", DARK)
    deck = Document.open(deck.to_bytes())
    slide = deck.slides[0]
    title = first_run(slide.shapes[0])
    # The master's title style says tx1; on this slide tx1 is lt1.
    theme = deck.theme
    assert title.effective_color == Color.rgb(theme.colors["lt1"])
    assert title.effective_color != Color.rgb(theme.colors["dk1"])


def test_a_runs_own_theme_colour_is_resolved_with_its_modifiers():
    deck = small_layout_deck()
    slide = deck.slides[0]
    run = first_run(slide.shapes[1])
    run.color = "accent1 lumMod=75%"
    assert run.effective_source("color") == "the run"
    assert run.effective_color == Color.rgb(Color.parse("accent1 lumMod=75%").resolve(slide))
    assert run.effective_color != Color.rgb(Color.parse("accent1").resolve(slide))
    run.color = "#123456"
    assert run.effective_color == Color.rgb("123456")


def test_a_plain_shape_takes_its_styles_font_colour_and_a_placeholder_does_not():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    box = slide.add_shape("rect", EMU, EMU, 2 * EMU, EMU)
    box.set_text("on accent")
    run = first_run(box)
    assert run.effective_source("color") == "the shape's style (fontRef)"
    assert run.effective_color == Color.rgb(deck.theme.colors["lt1"])
    run.color = "accent3"
    assert run.effective_color == Color.rgb(deck.theme.colors["accent3"])
    title = slide.shapes[0]
    title.set_text("Title")
    assert first_run(title).effective_source("color") == "the master's titleStyle"


def test_a_hyperlink_without_a_colour_is_the_themes_hlink():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    box = slide.add_shape("rect", EMU, EMU, 2 * EMU, EMU)
    box.set_text("example")
    run = first_run(box)
    run.set_hyperlink("https://example.org")
    assert run.effective_color == Color.rgb(deck.theme.colors["hlink"])
    assert run.effective_source("color") == "the hyperlink (theme hlink)"


def test_text_in_no_solid_colour_has_none():
    deck = small_layout_deck()
    slide = deck.slides[0]
    _patch(deck, slide.part_path, '<a:r><a:rPr lang="en-US"',
           '<a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t>x</a:t></a:r><a:r>'
           '<a:rPr lang="en-US"')
    deck = Document.open(deck.to_bytes())
    runs = [run for shape in deck.slides[0].shapes for run in shape.paragraphs[0].runs]
    hidden = next(run for run in runs if run.text == "x")
    assert hidden.effective_color is None
    assert hidden.effective_source("color") == "the run (noFill)"


def test_the_paragraph_reports_its_first_runs_colour():
    deck = small_layout_deck()
    slide = deck.slides[0]
    assert slide.shapes[1].paragraphs[0].effective.color == Color.parse("accent2").resolve(slide)


# -- table cells --------------------------------------------------------------------------


def test_a_cell_takes_its_built_in_table_styles_text_colour():
    """PowerPoint does not write a built-in style into tableStyles.xml; the header's white
    comes from the colours pptx2svg measured (Medium Style 2 - Accent 1, the default)."""
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    table = slide.add_table(3, 2, EMU, EMU, 4 * EMU, EMU).table
    table.cell(0, 0).text = "Header"
    table.cell(1, 0).text = "Body"
    header = first_run(table.cell(0, 0).text_frame)
    body = first_run(table.cell(1, 0).text_frame)
    assert header.effective_color == Color.rgb(deck.theme.colors["lt1"])
    assert header.effective_source("color") == "table style 'Medium Style 2 - Accent 1' firstRow"
    assert body.effective_color == Color.rgb(deck.theme.colors["dk1"])
    assert body.effective_source("color") == "table style 'Medium Style 2 - Accent 1' wholeTbl"
    table.set_flags(first_row=False)
    assert first_run(table.cell(0, 0).text_frame).effective_source("color").endswith("wholeTbl")
    # The cell's own colour still wins.
    body.color = "accent4"
    assert body.effective_color == Color.rgb(deck.theme.colors["accent4"])


ACME = "{0B5C6B1E-3D41-4C2A-9E57-6A2F1D8C4B90}"
ACME_STYLE = (f'<a:tblStyle styleId="{ACME}" styleName="Acme Grid"><a:wholeTbl><a:tcTxStyle>'
              '<a:schemeClr val="accent6"/></a:tcTxStyle></a:wholeTbl><a:band1H><a:tcTxStyle>'
              '<a:schemeClr val="accent5"/></a:tcTxStyle></a:band1H><a:firstRow><a:tcTxStyle '
              'b="on"><a:srgbClr val="FFEE00"/></a:tcTxStyle></a:firstRow></a:tblStyle>')


def test_a_cell_takes_the_decks_own_table_style_by_region():
    deck = Document.new(size="16:9")
    part = styles_part(deck)
    xml = deck.package.read(part).decode("utf-8")
    xml = (xml.replace("/>", f">{ACME_STYLE}</a:tblStyleLst>", 1) if xml.rstrip().endswith("/>")
           else xml.replace("</a:tblStyleLst>", f"{ACME_STYLE}</a:tblStyleLst>"))
    deck.package.replace_part(part, xml.encode("utf-8"))
    deck = Document.open(deck.to_bytes())
    slide = deck.add_slide("Title Only")
    table = slide.add_table(4, 2, EMU, EMU, 4 * EMU, 2 * EMU, style=ACME).table
    for row in range(4):
        table.cell(row, 0).text = f"r{row}"
    colors = [first_run(table.cell(row, 0).text_frame).effective_color for row in range(4)]
    theme = deck.theme.colors
    # Header, then bands 1, 2, 1 (the header row does not advance the banding).
    assert colors == [Color.rgb("FFEE00"), Color.rgb(theme["accent5"]),
                      Color.rgb(theme["accent6"]), Color.rgb(theme["accent5"])]
    assert first_run(table.cell(2, 0).text_frame).effective_source("color") \
        == "table style 'Acme Grid' wholeTbl"


def test_a_table_style_no_one_knows_gives_no_text_colour():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    frame = slide.add_table(2, 2, EMU, EMU, 4 * EMU, EMU)
    table = frame.table
    tbl = table._table()
    from pptx_agent.edit.tablestyles import write_style_id
    from pptx_agent.oxml.xml import qn

    write_style_id(tbl.find(qn("a:tblPr")), "{00000000-0000-0000-0000-0000000000AA}")
    table.cell(0, 0).text = "x"
    assert first_run(table.cell(0, 0).text_frame).effective_source("color") \
        == "the master's otherStyle"


# -- the tool layer -----------------------------------------------------------------------


def _session(fmt):
    toolbox = Toolbox(TOOLS, formats=[fmt], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(small_layout_deck().to_bytes(), "small.pptx")


def test_check_reports_inherited_small_text_against_the_applications_minimum():
    toolbox, session, doc = _session(deck_format(min_text_size=12))
    result = toolbox.dispatch(session, "check", {"doc": doc})
    data = result.data
    slide = session.documents[doc].document.slides[0]
    assert data["small_text"] == [{"shape": slide.shapes[1].id, "size": 9.0,
                                   "from": "layout 'Title and Content' content placeholder"}]
    assert "minimum of 12 pt" in data["small_text_note"]
    # The checks after a change report it too, and the fix clears it.
    target = slide.shapes[1].id
    changed = toolbox.dispatch(session, "ppt_format_text",
                               {"doc": doc, "items": [{"target": target, "bold": True}]})
    assert changed.checks["small_text"][0]["shape"] == target
    changed = toolbox.dispatch(session, "ppt_format_text",
                               {"doc": doc, "items": [{"target": target, "size": 12}]})
    assert "small_text" not in changed.checks
    again = toolbox.dispatch(session, "check", {"doc": doc})
    assert "small_text" not in again.data


def test_without_a_minimum_check_reports_no_small_text():
    toolbox, session, doc = _session(FORMAT)
    assert "small_text" not in toolbox.dispatch(session, "check", {"doc": doc}).data


def test_deck_format_refuses_a_minimum_out_of_range():
    with pytest.raises(ValueError, match="min_text_size"):
        deck_format(min_text_size=0)
    assert deck_format().kind == FORMAT.kind
