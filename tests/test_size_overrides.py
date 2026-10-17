"""Sizes set on a slide that override its layout's (round 7, L4): ``check`` and the checks
after a change list each placeholder whose slide states a text size other than the one its
layout (or master, through the placeholder's cascade) gives -- one entry per shape, a fact,
not a fault; nothing is changed by itself.  Every deck here is made up.
"""

from __future__ import annotations

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.effective import size_overrides_json
from pptx_agent.oxml.xml import find, qn
from pptx_agent.tools import GROUPS, TOOLS, deck_format
from test_min_text_kinds import CLOCK, EMU


def _layout_title_size(deck, layout_name: str, hundredths: int) -> None:
    """Give the layout's title placeholder a size of its own, as a run-specific layout does."""
    layout = deck.layout(layout_name)
    root = deck.package.tree(layout.part_path)
    for shape in find(root, "p:cSld/p:spTree"):
        ph = find(shape, "p:nvSpPr/p:nvPr/p:ph")
        if ph is not None and ph.get("type") == "title":
            body = shape.find(qn("p:txBody"))
            styles = body.find(qn("a:lstStyle"))
            for child in list(styles):
                styles.remove(child)
            level = styles.makeelement(qn("a:lvl1pPr"), {})
            level.append(level.makeelement(qn("a:defRPr"), {"sz": str(hundredths)}))
            styles.append(level)
    deck.package.mark_dirty(layout.part_path)


def _run_of_slides(sizes):
    """Like slides on a layout whose title is 28 pt; each title set to ``sizes[n]`` (``None``
    leaves it inherited)."""
    deck = Document.new(size="16:9")
    _layout_title_size(deck, "Title Only", 2800)
    slides = []
    for number, size in enumerate(sizes):
        slide = deck.add_slide("Title Only")
        slide.shapes[0].set_text(f"Region {number + 1}")
        if size is not None:
            slide.shapes[0].text_frame.paragraphs[0].runs[0].size = size
        slides.append(slide)
    return deck, slides


def _session(deck):
    toolbox = Toolbox(TOOLS, formats=[deck_format()], groups=GROUPS, workers=0)
    session = toolbox.session(clock=CLOCK)
    return toolbox, session, session.open(deck.to_bytes(), "made-up.pptx")


def test_check_lists_the_titles_set_off_the_layouts_size_one_entry_each():
    deck, slides = _run_of_slides([28, 26, None, 26])
    toolbox, session, doc = _session(deck)
    data = toolbox.dispatch(session, "check", {"doc": doc}).data
    # The slide restating 28 and the one inheriting it draw the layout's size: not listed.
    assert data["overrides"] == [
        {"shape": slides[1].shapes[0].id, "overrides": "size", "set": 26, "layout": 28,
         "fix": {"target": slides[1].shapes[0].id, "size": 0}},
        {"shape": slides[3].shapes[0].id, "overrides": "size", "set": 26, "layout": 28,
         "fix": {"target": slides[3].shapes[0].id, "size": 0}}]
    note = data["overrides_note"]
    assert note.startswith("overrides: a placeholder's text whose slide sets a size")
    assert "A fact, not a fault" in note and "size 0 removes every size" in note


def test_the_layout_size_comes_through_the_master_where_the_layout_states_none():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title Only")
    slide.shapes[0].set_text("A title")
    master = deck.layout("Title Only").placeholders[0].sizes[0]
    slide.shapes[0].text_frame.paragraphs[0].runs[0].size = master - 4
    found = slide.size_overrides()
    assert [(o.address, o.level, o.set, o.layout) for o in found] == \
        [(slide.shapes[0].id, 0, (master - 4,), master)]


def test_mixed_sizes_in_one_shape_are_one_entry_largest_first():
    deck, slides = _run_of_slides([None])
    title = slides[0].shapes[0]
    title.set_text("North and South")
    runs = title.text_frame.paragraphs[0]
    runs.add_run(" regions").size = 24
    runs.runs[0].size = 26
    assert size_overrides_json(slides[0].size_overrides()) == [
        {"shape": title.id, "overrides": "size", "set": [26, 24], "layout": 28,
         "fix": {"target": title.id, "size": 0}}]


def test_levels_with_different_layout_sizes_are_listed_by_level():
    deck = Document.new(size="16:9")
    slide = deck.add_slide("Title and Content")
    body = slide.shapes[1]
    body.set_text("one\ntwo")
    first, second = body.text_frame.paragraphs
    second.level = 1
    first.runs[0].size = 20
    second.runs[0].size = 20
    sizes = deck.layout("Title and Content").placeholders[1].sizes
    assert size_overrides_json(slide.size_overrides()) == [
        {"shape": body.id, "overrides": "size",
         "levels": [{"level": 0, "set": 20, "layout": int(sizes[0])},
                    {"level": 1, "set": 20, "layout": int(sizes[1])}],
         "fix": {"target": body.id, "size": 0}}]


def _defaults(holder, tag: str, hundredths: int) -> None:
    """``<tag><a:defRPr sz=.../></tag>`` first in ``holder``."""
    node = holder.makeelement(qn(tag), {})
    node.append(node.makeelement(qn("a:defRPr"), {"sz": str(hundredths)}))
    holder.insert(0, node)


def test_a_paragraph_or_the_shapes_list_style_setting_it_counts_too():
    deck, slides = _run_of_slides([None, None])
    _defaults(slides[0].shapes[0].text_frame.paragraphs[0]._element(), "a:pPr", 2600)
    _defaults(slides[1].shapes[0]._element.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"),
              "a:lvl1pPr", 3000)
    assert [(o.set, o.layout) for o in slides[0].size_overrides()] == [((26.0,), 28.0)]
    assert [(o.set, o.layout) for o in slides[1].size_overrides()] == [((30.0,), 28.0)]


def test_text_boxes_stored_autofit_and_whitespace_are_not_overrides():
    deck, slides = _run_of_slides([None])
    slide = slides[0]
    box = slide.add_textbox(40 * EMU, 200 * EMU, 300 * EMU, 40 * EMU)
    box.set_text("A footnote")
    box.text_frame.paragraphs[0].runs[0].size = 9          # no layout gives a text box a size
    title = slide.shapes[0]
    title.text_frame.autofit = "normal"
    title.text_frame.font_scale = 0.9                      # PowerPoint's shrink, not a set size
    title.text_frame.paragraphs[0].add_run("   ").size = 12  # draws no ink
    assert slide.size_overrides() == []


def test_check_keeps_its_slides_and_include_scoping():
    deck, slides = _run_of_slides([26, 26])
    toolbox, session, doc = _session(deck)
    one = toolbox.dispatch(session, "check", {"doc": doc, "slides": [2]}).data
    assert [entry["shape"] for entry in one["overrides"]] == [slides[1].shapes[0].id]
    collisions = toolbox.dispatch(session, "check",
                                  {"doc": doc, "include": ["collisions"]}).data
    assert "overrides" not in collisions


def test_the_checks_after_a_change_say_it_and_save_does_not_call_it_unresolved():
    deck, slides = _run_of_slides([None])
    toolbox, session, doc = _session(deck)
    title = slides[0].shapes[0].id
    result = toolbox.dispatch(session, "ppt_format_text",
                              {"doc": doc, "items": [{"target": title, "size": 26}]})
    assert result.checks["overrides"] == [
        {"shape": title, "overrides": "size", "set": 26, "layout": 28,
         "fix": {"target": title, "size": 0}}]
    # Setting the layout's size draws it: the override is gone.
    again = toolbox.dispatch(session, "ppt_format_text",
                             {"doc": doc, "items": [{"target": title, "size": 28}]})
    assert "overrides" not in again.checks
    toolbox.dispatch(session, "ppt_format_text",
                     {"doc": doc, "items": [{"target": title, "size": 26}]})
    saved = toolbox.dispatch(session, "save_document",
                             {"doc": doc, "name": "out.pptx", "format": "pptx"})
    assert saved.ok, saved.summary
    assert "overrides" not in (saved.data or {}).get("unresolved", {})
    assert "unresolved" not in saved.summary


def _effective_sizes(session, doc, address):
    shape = session.documents[doc].document.shape(address)
    return [run.effective_size for paragraph in shape.text_frame.paragraphs
            for run in paragraph.runs]


def test_each_fix_applied_clears_its_entry_and_the_layouts_size_is_drawn():
    deck, slides = _run_of_slides([26, None])
    # The second title sets 30 three ways: a run, its paragraph's defaults, its list style.
    title = slides[1].shapes[0]
    title.set_text("North and South")
    title.text_frame.paragraphs[0].add_run(" regions").size = 30
    _defaults(title.text_frame.paragraphs[0]._element(), "a:pPr", 3000)
    _defaults(title._element.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"), "a:lvl1pPr", 3000)
    body = deck.add_slide("Title and Content").shapes[1]
    body.set_text("one\ntwo")
    body.text_frame.paragraphs[1].level = 1
    body.text_frame.format(size=20)
    toolbox, session, doc = _session(deck)
    found = toolbox.dispatch(session, "check", {"doc": doc}).data["overrides"]
    assert len(found) == 3
    result = toolbox.dispatch(session, "ppt_format_text",
                              {"doc": doc, "items": [entry["fix"] for entry in found]})
    assert result.ok, result.summary
    assert "overrides" not in result.checks
    assert "overrides" not in toolbox.dispatch(session, "check", {"doc": doc}).data
    assert _effective_sizes(session, doc, slides[0].shapes[0].id) == [28.0]
    assert _effective_sizes(session, doc, title.id) == [28.0, 28.0]
    sizes = deck.layout("Title and Content").placeholders[1].sizes
    assert _effective_sizes(session, doc, body.id) == [sizes[0], sizes[1]]


def test_size_0_on_a_run_or_paragraph_clears_only_that_and_below_1_is_refused():
    deck, slides = _run_of_slides([None])
    title = slides[0].shapes[0]
    title.set_text("North and South")
    paragraph = title.text_frame.paragraphs[0]
    paragraph.runs[0].size = 26
    paragraph.add_run(" regions").size = 24
    toolbox, session, doc = _session(deck)
    toolbox.dispatch(session, "ppt_format_text",
                     {"doc": doc, "items": [{"target": f"{title.id}/p0/r0", "size": 0}]})
    assert _effective_sizes(session, doc, title.id) == [28.0, 24.0]
    toolbox.dispatch(session, "ppt_format_text",
                     {"doc": doc, "items": [{"target": f"{title.id}/p0", "size": 0}]})
    assert _effective_sizes(session, doc, title.id) == [28.0, 28.0]
    refused = toolbox.dispatch(session, "ppt_format_text",
                               {"doc": doc, "items": [{"target": title.id, "size": 0.5}]})
    assert not refused.ok and "0 to remove the size set here" in refused.summary + str(
        refused.data)
