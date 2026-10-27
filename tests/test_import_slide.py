"""A slide imported from another open deck (production feedback round 17, P1):
Document.import_slide and ppt_manage_slides' import.

The decks are made here from one made-up template: a deck of samples and a report deck."""

from __future__ import annotations

from pathlib import Path

import pytest
from ooxml_edit.tools import Toolbox

from images import png
from pptx_agent import Document
from pptx_agent.edit.errors import ArgumentError
from pptx_agent.oxml.xml import qn
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

FIXTURES = Path(__file__).resolve().parent / "fixtures"
E = 12700
RED = png(4, 4, lambda x, y: (200, 30, 30))
BLUE = png(4, 4, lambda x, y: (30, 30, 200))


def _template() -> bytes:
    """A made-up template: the Office layouts and one of its own, "Sample Item", with a
    picture on it."""
    template = Document.new()
    layout = template.add_layout("Sample Item")
    layout.add_picture(BLUE, 20 * E, 20 * E, 30 * E, 30 * E)
    return template.to_bytes()


def _samples(template: bytes) -> Document:
    """The deck of samples: one slide on "Sample Item" with a named header, a picture, a
    chart and notes."""
    deck = Document.new(template=template)
    slide = deck.add_slide("Sample Item")
    slide.title = "Sample record"
    header = slide.add_shape("rect", 40 * E, 100 * E, 400 * E, 40 * E, text="Header")
    header.name = "Header"
    slide.add_picture(RED, 500 * E, 100 * E, 40 * E, 40 * E)
    slide.add_picture(RED, 560 * E, 100 * E, 40 * E, 40 * E)
    slide.add_chart("column", ["A", "B"], [{"name": "Count", "values": [3, 5]}],
                    40 * E, 160 * E, 300 * E, 200 * E, name="Counts")
    slide.notes = "Read the sample aloud."
    return deck


def _report(template: bytes) -> Document:
    deck = Document.new(template=template)
    deck.add_slide("Title Slide").title = "Report"
    deck.slides[0].add_picture(RED, 0, 0, 10 * E, 10 * E)
    return deck


@pytest.fixture()
def decks():
    template = _template()
    return _samples(template), _report(template)


def _names(slide):
    return [(shape.name, shape.kind) for shape in slide.shapes]


def test_a_slide_is_imported_with_its_parts_its_layout_by_name_and_its_shape_names(decks):
    samples, report = decks
    source = samples.slides[0]
    samples_before = samples.to_bytes()
    result = report.import_slide(source)
    slide = result.slide
    assert slide.index == 1 and result.layout.name == "Sample Item"
    assert slide.layout.part_path == report.layout("Sample Item").part_path
    assert _names(slide) == _names(source)
    assert result.mapping == {
        old.id: f"{slide.slide_id}.{old.id.split('.', 1)[1]}" for old in source.shapes}
    assert slide.notes == "Read the sample aloud."
    chart = next(shape for shape in slide.shapes if shape.name == "Counts")
    assert chart.has_chart and chart.chart.series[0].values == [3, 5]
    assert not result.theme_differs and result.brought == []
    package = report.package
    # one picture part for both pictures, and the report's own red picture shared
    pictures = {shape.image_part for shape in slide.shapes if shape.kind == "picture"}
    assert pictures == {report.slides[0].shapes[-1].image_part}
    charts = [name for name in package.part_names if name.startswith("ppt/charts/chart")]
    workbooks = [name for name in package.part_names if name.startswith("ppt/embeddings/")]
    assert len(charts) == 1 and len(workbooks) == 1
    assert report.validate() == []
    assert samples.to_bytes() == samples_before          # the source is only read
    reopened = Document.open(report.to_bytes())
    assert reopened.validate() == [] and reopened.slides[1].notes == "Read the sample aloud."


def test_the_import_is_one_undo_step_that_takes_its_parts_away(decks):
    samples, report = decks
    before = report.to_bytes()
    parts = set(report.package.part_names)
    report.import_slide(samples.slides[0], index=0)
    assert report.slides[0].title == "Sample record"
    report.undo()
    assert set(report.package.part_names) == parts
    assert report.to_bytes() == before


def test_a_missing_layout_is_refused_with_the_target_s_layouts(decks):
    samples, _ = decks
    plain = Document.new()
    with pytest.raises(ArgumentError, match="no layout 'Sample Item'.*'Title Slide'"):
        plain.import_slide(samples.slides[0])
    assert len(plain.slides) == 0


def test_bring_layout_copies_the_layout_onto_the_master_of_the_same_name(decks):
    samples, _ = decks
    plain = Document.new()
    masters = len(plain.masters)
    result = plain.import_slide(samples.slides[0], bring_layout=True)
    assert result.brought == ["layout"] and len(plain.masters) == masters
    layout = plain.layout("Sample Item")
    assert result.slide.layout.part_path == layout.part_path
    assert layout.master.part_path == plain.masters[0].part_path
    assert any(shape.kind == "picture" for shape in layout.shapes)
    assert plain.validate() == []
    # a second import finds the layout now there
    again = plain.import_slide(samples.slides[0], bring_layout=True)
    assert again.brought == [] and again.slide.layout.part_path == layout.part_path
    assert Document.open(plain.to_bytes()).validate() == []


def test_bring_layout_brings_the_master_when_the_target_has_none_of_its_name(decks):
    samples, _ = decks
    other = Document.new()
    root = other.package.tree(other.masters[0].part_path)
    root.find(qn("p:cSld")).set("name", "Other Master")
    other.package.mark_dirty(other.masters[0].part_path)
    other.theme.set_colors({"accent1": "#2E8B57"})
    result = other.import_slide(samples.slides[0], bring_layout=True)
    assert result.brought == ["master", "layout"]
    assert [master.name for master in other.masters] == ["Other Master", "Office Theme"]
    assert [lay.name for lay in other.masters[1].layouts] == ["Sample Item"]
    assert result.slide.theme.colors["accent1"] == samples.slides[0].theme.colors["accent1"]
    assert not result.theme_differs                     # the source's theme came along
    assert other.validate() == []
    assert Document.open(other.to_bytes()).validate() == []
    other.undo()
    assert len(other.masters) == 1 and len(other.slides) == 0


def test_a_target_with_other_theme_colours_keeps_the_references_and_says_so(decks):
    samples, report = decks
    source = samples.slides[0]
    header = next(shape for shape in source.shapes if shape.name == "Header")
    header.fill = "accent2"
    report.theme.set_colors({"accent2": "#2E8B57"})
    result = report.import_slide(source)
    copied = next(shape for shape in result.slide.shapes if shape.name == "Header")
    assert result.theme_differs
    assert {"color": "accent2", "was": source.theme.colors["accent2"],
            "now": "#2E8B57"} in result.theme_changes
    assert copied.fill.color.resolve(copied).upper().lstrip("#") == "2E8B57"


def test_a_jump_to_another_source_slide_and_comments_are_left_behind(decks):
    samples, report = decks
    second = samples.add_slide("Title Only")
    source = samples.slides[0]
    header = next(shape for shape in source.shapes if shape.name == "Header")
    header.text_frame.paragraphs[0].runs[0].set_hyperlink(second)
    run_props = header._element.find(f".//{qn('a:rPr')}")
    assert run_props.find(qn("a:hlinkClick")) is not None
    result = report.import_slide(source)
    copied = next(shape for shape in result.slide.shapes if shape.name == "Header")
    assert result.dropped_links == [copied.id]
    assert copied.text == "Header"
    assert not list(copied._element.iter(qn("a:hlinkClick")))
    assert report.validate() == []
    with pytest.raises(ArgumentError, match="another deck"):
        samples.import_slide(source)


def test_a_slide_of_a_real_deck_with_smartart_comes_with_its_diagram():
    source = Document.open(FIXTURES / "powerpoint-smartart.pptx")
    target = Document.new()
    result = target.import_slide(source.slide(256), bring_layout=True)
    diagram = target.shape(result.mapping["256.3"])
    assert diagram.has_diagram
    assert any(name.startswith("ppt/diagrams/data") for name in target.package.part_names)
    assert target.validate() == []
    assert Document.open(target.to_bytes()).validate() == []


# -- the tool ---------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def test_the_tool_imports_and_reports_the_new_ids(toolbox, decks):
    samples, report = decks
    session = toolbox.session()
    d1 = session.open(report.to_bytes(), "report.pptx")
    d2 = session.open(samples.to_bytes(), "samples.pptx")
    result = toolbox.dispatch(session, "ppt_manage_slides",
                              {"doc": d1, "action": "import", "source_doc": d2,
                               "slide": "s:256", "to": 1, "ref": "first"})
    assert result.ok, result.error
    new = result.created[0]
    sid = new.split(":")[1]
    assert result.data["layout"] == "Sample Item" and result.data["order"][0] == new
    header = next(s for s in samples.slides[0].shapes if s.name == "Header")
    assert result.renamed[f"{d2}/{header.id}"] == f"{sid}.{header.id.split('.')[1]}"
    assert result.refs == {"first": new}
    undone = toolbox.dispatch(session, "undo", {"doc": d1})
    assert undone.ok and len(session.documents[d1].document.slides) == 1


def test_the_tool_refuses_a_missing_layout_and_a_deck_copied_into_itself(toolbox, decks):
    samples, _ = decks
    session = toolbox.session()
    d1 = session.open(Document.new().to_bytes(), "plain.pptx")
    d2 = session.open(samples.to_bytes(), "samples.pptx")
    missing = toolbox.dispatch(session, "ppt_manage_slides",
                               {"doc": d1, "action": "import", "source_doc": d2,
                                "slide": "s:256"})
    assert not missing.ok and missing.error.code == "not_found"
    assert 'layout: "source"' in missing.error.message
    assert "Title Slide" in missing.error.valid_options
    brought = toolbox.dispatch(session, "ppt_manage_slides",
                               {"doc": d1, "action": "import", "source_doc": d2,
                                "slide": "s:256", "layout": "source"})
    assert brought.ok, brought.error
    assert brought.data["brought"][0].endswith("(Sample Item)")
    other = toolbox.dispatch(session, "ppt_manage_slides",
                             {"doc": d1, "action": "import", "source_doc": d2,
                              "slide": "s:256", "layout": "Title Slide"})
    assert not other.ok and other.error.field == "layout"
    assert other.error.valid_options == ["source"]
    itself = toolbox.dispatch(session, "ppt_manage_slides",
                              {"doc": d2, "action": "import", "source_doc": d2,
                               "slide": "s:256"})
    assert not itself.ok and itself.error.field == "source_doc"
    nothing = toolbox.dispatch(session, "ppt_manage_slides",
                               {"doc": d1, "action": "import", "slide": "s:256"})
    assert not nothing.ok and nothing.error.field == "source_doc"


def test_the_tool_reads_at_as_to(toolbox, decks):
    samples, report = decks
    session = toolbox.session()
    d1 = session.open(report.to_bytes(), "report.pptx")
    d2 = session.open(samples.to_bytes(), "samples.pptx")
    result = toolbox.dispatch(session, "ppt_manage_slides",
                              {"doc": d1, "action": "import", "source_doc": d2,
                               "slide": "s:256", "at": 1})
    assert result.ok and result.data["order"][0] == result.created[0]



def test_import_keeps_every_strict_tool_strict(toolbox):
    strict = [d["name"] for d in toolbox.definitions("anthropic", defer=False)
              if d.get("strict")]
    assert strict == ["replace_text", "set_properties", "ppt_add_slide", "ppt_draft_slides",
                      "ppt_manage_slides", "edit_smartart"]


# -- notes on the copy (production feedback round 18, Q2) ---------------------------------


def test_notes_replace_the_imported_copy_s_as_on_a_duplicate_one_undo_step(decks):
    samples, report = decks
    before, parts = report.to_bytes(), set(report.package.part_names)
    made = report.import_slide(samples.slides[0], notes="Say the totals first.")
    assert made.slide.notes == "Say the totals first."
    assert samples.slides[0].notes == "Read the sample aloud."     # the source only read
    report.undo()                                 # the copy and its notes: one step
    assert set(report.package.part_names) == parts and report.to_bytes() == before
    report.redo()
    assert report.slides[-1].notes == "Say the totals first."
    reopened = Document.open(report.to_bytes())
    assert reopened.validate() == [] and reopened.slides[-1].notes == "Say the totals first."


def test_empty_notes_empty_the_copy_s_and_notes_go_on_a_slide_without_any(decks):
    samples, report = decks
    emptied = report.import_slide(samples.slides[0], notes="").slide
    copied = samples.duplicate_slide(samples.slides[0], notes="")
    assert (emptied.has_notes, emptied.notes) == (copied.has_notes, copied.notes) \
        == (True, "")
    plain = report.slides[0]                      # no notes page of its own
    assert not plain.has_notes
    samples_without = Document.new(template=_template())
    samples_without.add_slide("Title Slide").title = "No notes"
    added = report.import_slide(samples_without.slides[0], notes="Only on the copy").slide
    assert added.notes == "Only on the copy" and not samples_without.slides[0].has_notes
    untouched = report.import_slide(samples_without.slides[0], notes="").slide
    assert not untouched.has_notes                # as duplicate: "" adds no notes page
    assert report.validate() == []


def test_the_tool_s_import_takes_notes_as_duplicate_does_and_undoes_them(toolbox, decks):
    samples, report = decks
    session = toolbox.session()
    d1 = session.open(report.to_bytes(), "report.pptx")
    d2 = session.open(samples.to_bytes(), "samples.pptx")
    result = toolbox.dispatch(session, "ppt_manage_slides",
                              {"doc": d1, "action": "import", "source_doc": d2,
                               "slide": "s:256", "at": 1, "notes": "Say the totals first."})
    assert result.ok, result.error
    deck = session.documents[d1].document
    assert deck.slides[0].notes == "Say the totals first."
    assert session.documents[d2].document.slides[0].notes == "Read the sample aloud."
    duplicated = toolbox.dispatch(session, "ppt_manage_slides",
                                  {"doc": d1, "action": "duplicate", "slide": result.created[0],
                                   "notes": ""})
    assert duplicated.ok and deck.slides[1].notes == ""
    assert toolbox.dispatch(session, "undo", {"doc": d1, "steps": 2}).ok
    assert len(deck.slides) == 1 and not deck.slides[0].has_notes
    assert toolbox.dispatch(session, "undo", {"doc": d1, "redo": True}).ok
    assert deck.slides[0].notes == "Say the totals first."
    saved = Document.open(deck.to_bytes())
    assert [slide.notes for slide in saved.slides] == ["Say the totals first.", ""]


def test_a_reviewer_is_sent_no_parameter_of_the_actions_it_lacks(toolbox):
    """Tool.action_fields: a reviewer's ppt_manage_slides has find_by_title only, so none of
    the copying, moving or layout parameters; an editor's has those, not find_by_title's."""
    def properties(roles):
        (sent,) = [d for d in toolbox.definitions("anthropic", defer=False, roles=roles)
                   if d["name"] == "ppt_manage_slides"]
        return set(sent["input_schema"]["properties"])

    assert properties(["review"]) == {"doc", "action", "title"}
    assert properties(["edit"]) == {"doc", "action", "slide", "to", "notes", "ref", "layout",
                                    "source_doc"}
    assert properties(["edit", "review"]) == properties(["edit"]) | {"title"}
