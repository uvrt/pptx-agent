"""Empty placeholders (production feedback round 3, A3): what PowerPoint shows as "Click to
add text" while editing and draws no prompt for in a show or an export (measured in
PowerPoint 16: a whitespace-only one shows a stray bullet instead, and a placeholder's fill
or outline is drawn either way).

``Slide.empty_placeholders`` names them, ``delete_empty_placeholders`` removes them in one
undo step (an empty title stays unless asked), and the tools report them: in every
changing call's checks, in ``check`` and in ``save_document``'s ``unresolved``.
"""

from __future__ import annotations

import copy

from ooxml_edit.tools import Toolbox

from images import MARKER
from pptx_agent import Document
from pptx_agent.oxml.xml import find, make, qn
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

EMU = 914400


def _ids(slide) -> list[str]:
    return [shape.id for shape in slide.empty_placeholders()]


def _as_placeholder(shape, kind: str | None, idx: int | None) -> None:
    """Make a picture or graphic frame a placeholder, as PowerPoint writes one that took it."""
    container = {"pic": "p:nvPicPr", "graphicFrame": "p:nvGraphicFramePr"}[
        shape._element.tag.split("}")[1]]
    nv = find(shape._element, f"{container}/p:nvPr")
    holder = make("p:ph")
    if kind:
        holder.set("type", kind)
    if idx is not None:
        holder.set("idx", str(idx))
    nv.insert(0, holder)
    shape._slide._touch()


def _clone_as(slide, source, kind: str) -> str:
    """A copy of ``source`` (a placeholder) as a ``kind`` placeholder -- dt, ftr, sldNum."""
    element = copy.deepcopy(source._element)
    find(element, "p:nvSpPr/p:nvPr/p:ph").set("type", kind)
    find(element, "p:nvSpPr/p:cNvPr").set("id", str(90 + len(kind)))
    source._element.getparent().append(element)
    slide._invalidate()
    return f"{slide.slide_id}.{90 + len(kind)}"


def test_what_counts_as_empty():
    deck = Document.new()
    cover = deck.add_slide("Title Slide")
    assert _ids(cover) == ["256.2", "256.3"]            # the title and subtitle count
    cover.title = "Q3 review"
    assert _ids(cover) == ["256.3"]

    slide = deck.add_slide("Title and Content")
    title, body = slide.shapes
    body.set_text("   \v ")                              # whitespace only
    assert _ids(slide) == [title.id, body.id]
    body.set_text("\n\n")                                # empty paragraphs only
    assert body.id in _ids(slide)
    body.set_text("Revenue grew")
    assert _ids(slide) == [title.id]
    body.set_text("")
    # A field is text PowerPoint fills in: not empty.
    paragraph = find(body._element, "p:txBody/a:p")
    field = make("a:fld")
    field.set("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}")
    field.set("type", "slidenum")
    field.append(make("a:t"))
    paragraph.insert(0, field)
    slide._touch()
    assert body.id not in _ids(slide)

    # Date, footer and slide number are PowerPoint's to fill; a hidden one does not show.
    third = deck.add_slide("Title and Content")
    third.title = "Footers"
    holder = third.shapes[1]
    added = [_clone_as(third, holder, kind) for kind in ("dt", "ftr", "sldNum")]
    assert all(third.shape(address).placeholder.type in ("dt", "ftr", "sldNum")
               for address in added)
    assert _ids(third) == [holder.id]
    find(holder._element, "p:nvSpPr/p:cNvPr").set("hidden", "1")
    third._touch()
    assert _ids(third) == []


def test_a_placeholder_holding_a_picture_or_a_table_is_not_empty():
    deck = Document.new()
    slide = deck.add_slide("Picture with Caption")
    title, picture_holder, caption = slide.shapes
    assert _ids(slide) == [title.id, picture_holder.id, caption.id]
    picture_holder.delete()
    picture = slide.add_picture(MARKER, EMU, EMU, 2 * EMU, EMU)
    _as_placeholder(picture, "pic", 1)
    other = deck.add_slide("Title and Content")
    other.shapes[1].delete()
    table = other.add_table(2, 2, EMU, EMU, 4 * EMU, EMU)
    _as_placeholder(table, None, 1)
    assert picture.placeholder == ("pic", 1) and table.placeholder == (None, 1)
    assert _ids(slide) == [title.id, caption.id]
    assert _ids(other) == [other.shapes[0].id]


def test_deleting_them_is_one_undo_step_and_keeps_the_title():
    deck = Document.new()
    slide = deck.add_slide("Two Content")
    before = deck.to_bytes()
    title = slide.shapes[0].id
    assert len(_ids(slide)) == 3
    deleted = slide.delete_empty_placeholders()
    assert len(deleted) == 2 and title not in deleted
    assert _ids(slide) == [title]
    assert deck.validate() == []
    assert deck.undo()
    assert deck.to_bytes() == before
    slide = deck.slides[0]
    assert slide.delete_empty_placeholders(titles=True) == [title] + deleted
    assert _ids(slide) == [] and slide.title is None


def test_one_that_draws_a_box_is_reported_but_kept():
    """PowerPoint draws an empty placeholder's fill and outline in the export (measured):
    deleting it would change the slide."""
    deck = Document.new()
    slide = deck.add_slide("Two Content")
    title, left, right = slide.shapes
    left.fill = "#FF0000"
    right.line.color = "accent1"
    right.line.width = 12700
    assert _ids(slide) == [title.id, left.id, right.id]
    assert slide.delete_empty_placeholders() == []
    left.fill = None
    left._element.find(qn("p:spPr")).append(make("a:noFill"))   # says "no fill" outright
    slide._touch()
    assert slide.delete_empty_placeholders() == [left.id]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        doc = session.open(deck.to_bytes(), "deck.pptx")
        kept = toolbox.dispatch(session, "ppt_manage_slides", {
            "doc": doc, "action": "delete_empty_placeholders", "slide": "s:256"})
        assert kept.removed == [] and kept.data["kept_drawn"] == [right.id]
        assert kept.data["empty_title"] == [title.id] and "fill or outline" in kept.summary


def test_the_tools_report_them_and_delete_them():
    # The round-3 repro: a chart over a content layout's body, which stays empty.
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title and Content")
    slide.title = "Spend by category"
    body = slide.shapes[1].id
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        doc = session.open(deck.to_bytes(), "deck.pptx")
        added = toolbox.dispatch(session, "edit_chart", {
            "doc": doc, "target": "s:256", "action": "add", "chart_type": "column",
            "categories": ["A", "B"], "data": [{"name": "Spend", "values": [3, 5]}],
            "box": {"x": 72, "y": 126, "w": 472, "h": 315}})
        assert added.ok, added.to_json()
        wanted = [{"shape": body, "type": "content", "idx": 1}]
        assert added.checks["empty_placeholders"] == wanted
        assert "Click to add" in added.checks["empty_placeholders_note"]
        found = toolbox.dispatch(session, "check", {"doc": doc}).data
        assert found["empty_placeholders"] == wanted
        assert "empty_placeholders" not in toolbox.dispatch(
            session, "check", {"doc": doc, "include": ["validate"]}).data
        saved = toolbox.dispatch(session, "save_document", {"doc": doc, "name": "a.pptx",
                                                            "format": "pptx"})
        assert saved.data["unresolved"]["empty_placeholders"] == wanted
        assert "1 empty placeholders" in saved.summary
        removed = toolbox.dispatch(session, "ppt_manage_slides", {
            "doc": doc, "action": "delete_empty_placeholders", "slide": "s:256"})
        assert removed.ok and removed.removed == [body], removed.to_json()
        assert "empty_placeholders" not in removed.checks
        saved = toolbox.dispatch(session, "save_document", {"doc": doc, "name": "b.pptx",
                                                            "format": "pptx"})
        assert "unresolved" not in saved.data
        assert toolbox.dispatch(session, "undo", {"doc": doc}).ok
        assert toolbox.dispatch(session, "check", {"doc": doc}).data[
            "empty_placeholders"] == wanted


def test_an_empty_title_is_kept_and_said():
    deck = Document.new()
    deck.add_slide("Title and Content")
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        doc = session.open(deck.to_bytes(), "deck.pptx")
        removed = toolbox.dispatch(session, "ppt_manage_slides", {
            "doc": doc, "action": "delete_empty_placeholders", "slide": "s:256"})
        assert removed.removed == ["256.3"] and removed.data["empty_title"] == ["256.2"]
        assert "kept the empty title 256.2" in removed.summary
        assert removed.checks["empty_placeholders"] == [
            {"shape": "256.2", "type": "title", "idx": None}]


def test_the_renderer_draws_nothing_for_them():
    """PowerPoint's export draws no prompt for an empty placeholder (measured); neither does
    pptx2svg."""
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    svg = slide.render_svg()
    assert "Click to add" not in svg and "<text" not in svg


def test_a_whitespace_one_is_reported_and_drawn_with_its_stray_bullet():
    """PowerPoint draws the bullet of a placeholder that holds only whitespace, editing and
    exported (measured: a space, a tab, a no-break space, several; ooxml-common's
    ``tests/test_blank_paragraphs.py``), and nothing for one of empty paragraphs; pptx2svg
    draws the same, so the render shows what ``empty_placeholders`` warns about."""
    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    title, body = slide.shapes
    title.set_text("Q3 review")
    for blank in ("   ", "\t", " "):
        body.set_text(blank)
        assert _ids(slide) == [body.id]
        svg = slide.render_svg()
        assert svg.count("•") == 1 and "Click to add" not in svg, repr(blank)
    body.set_text("\n\n")
    assert _ids(slide) == [body.id]
    assert "•" not in slide.render_svg()
