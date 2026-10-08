"""Speaker notes of an existing slide: read, write, address, find, duplicate with new notes.

Six of the full trial's twelve PowerPoint runs needed them and found no way in: notes could
only be drafted with ``insert_outline``, never read as text, edited, found or addressed.
"""

from __future__ import annotations

import os
import warnings
from pathlib import Path

import pytest

import oracle as oracle_helper
from conftest import FIXTURE_DIR
from pptx_agent import Document, TextFrame
from pptx_agent.edit import notes as _notes

SAMPLE = FIXTURE_DIR / "sample.pptx"           # slide 257 has notes; the deck has a master


def _deck() -> Document:
    return Document.open(SAMPLE)


def test_existing_notes_read_as_raw_text():
    deck = _deck()
    assert deck.slide(257).notes == "ここがプレゼンターノートになります"
    assert deck.slide(257).has_notes
    without = [slide for slide in deck.slides if not slide.has_notes]
    assert without and all(slide.notes == "" for slide in without)
    assert isinstance(deck.slide(257).notes_frame, TextFrame)
    assert deck.slide(257).notes_frame.address == "257/notes"


def test_setting_notes_keeps_the_formatting_that_survives():
    deck = _deck()
    slide = deck.slide(257)
    frame = slide.notes_frame
    frame.set_text("Revenue grew 11%\nThen the outlook")
    frame.paragraph(0).segment(["Revenue grew ", "11%"])
    frame.paragraph(0).run(1).bold = True
    slide.notes = "Revenue grew 12%\nThen the outlook"
    runs = slide.notes_frame.paragraph(0).runs
    assert [run.text for run in runs] == ["Revenue grew ", "12%"]
    assert runs[1].bold is True
    assert slide.notes == "Revenue grew 12%\nThen the outlook"


def test_every_notes_edit_is_one_undo_step():
    deck = _deck()
    slide = deck.slide(257)
    before = slide.notes
    slide.notes = "First"
    slide.notes_frame.add_paragraph("Second")
    assert deck.slide(257).notes == "First\nSecond"
    assert deck.undo() and deck.slide(257).notes == "First"
    assert deck.undo() and deck.slide(257).notes == before
    assert deck.redo() and deck.slide(257).notes == "First"


def test_notes_on_a_slide_without_any_are_made_as_powerpoint_makes_them():
    """A new deck has no notes master: the first notes add one, with its own theme, exactly
    as insert_outline's drafted notes do (E6, measured on PowerPoint)."""
    by_setter = Document.new()
    slide = by_setter.add_slide("Title and Content")
    parts_before = set(by_setter.package.part_names)
    slide.notes = "Lead with the margin.\nThen the outlook."
    added = set(by_setter.package.part_names) - parts_before
    assert {p.rpartition("/")[0] for p in added} >= {
        "ppt/notesMasters", "ppt/notesSlides", "ppt/theme"}

    by_outline = Document.new()
    by_outline.insert_outline("# \n\nNotes:\n\nLead with the margin.\n\nThen the outlook.\n")
    master_a = _notes.notes_master(by_setter.package)
    master_b = _notes.notes_master(by_outline.package)
    assert by_setter.package.read(master_a) == by_outline.package.read(master_b)
    page = _notes.notes_part(by_setter.package, slide.part_path)
    drafted = _notes.notes_part(by_outline.package, by_outline.slides[0].part_path)
    assert _shape_skeleton(by_setter, page) == _shape_skeleton(by_outline, drafted)
    assert by_setter.validate() == []
    assert by_setter.slides[0].notes == by_outline.slides[0].notes

    by_setter.undo()
    assert not by_setter.slides[0].has_notes
    assert set(by_setter.package.part_names) == parts_before


def _shape_skeleton(deck: Document, part: str) -> list:
    root = deck.package.tree(part)
    from pptx_agent.oxml.xml import qn

    return [(shape.find(f".//{qn('p:cNvPr')}").get("name"),
             dict(shape.find(f".//{qn('p:ph')}").attrib))
            for shape in root.iter(qn("p:sp"))]


def test_empty_notes_on_a_slide_without_any_change_nothing():
    deck = Document.new()
    slide = deck.add_slide("Title Slide")
    data = deck.to_bytes()
    slide.notes = ""
    assert deck.to_bytes() == data and not slide.has_notes


def test_the_notes_address_resolves():
    deck = _deck()
    assert deck.resolve("257/notes").text == "ここがプレゼンターノートになります"
    assert deck.resolve("s:257/notes/p0").text == "ここがプレゼンターノートになります"
    deck.resolve("257/notes/p0/r0").text = "差し替え"
    assert deck.slide(257).notes == "差し替え"
    with pytest.raises(KeyError):
        deck.resolve("999/notes")


def test_notes_are_found_and_listed():
    deck = _deck()
    (block,) = deck.find_text("プレゼンターノート")
    assert (block.address, block.kind, block.slide) == ("257/notes", "notes", 2)
    assert deck.resolve(block.address).text == block.text
    blocks = deck.outline_blocks(slides=[2])
    assert blocks[-1].kind == "notes"


def test_the_outline_shows_the_notes_address_and_still_drafts_them():
    deck = _deck()
    outline = deck.to_outline(slides=["s:257"])
    assert "<!-- 257/notes -->\nNotes:\n\nここがプレゼンターノートになります\n" in outline
    assert "257/notes" not in deck.to_outline(slides=["s:257"], ids=False)
    fresh = Document.new()
    (slide,) = fresh.insert_outline(outline)
    assert slide.notes == "ここがプレゼンターノートになります"


def test_duplicate_with_new_notes():
    deck = _deck()
    original = deck.slide(257)
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        copy = deck.duplicate_slide(original, notes="Engineering works from the Utrecht office.")
    assert copy.notes == "Engineering works from the Utrecht office."
    assert deck.slide(257).notes == "ここがプレゼンターノートになります"
    assert copy.index == original.index + 1
    assert deck.validate() == Document.open(SAMPLE).validate()
    deck.undo()                                   # the copy and its notes: one step
    assert [s.slide_id for s in deck.slides] == [s.slide_id for s in _deck().slides]

    other = Document.new()
    plain = other.add_slide("Title Slide")        # no notes page of its own
    second = plain.duplicate(notes="Only on the copy")
    assert second.notes == "Only on the copy" and not other.slide(plain.slide_id).has_notes
    assert other.validate() == []


requires_powerpoint = pytest.mark.skipif(
    not oracle_helper.available(),
    reason="needs macOS with Microsoft PowerPoint and the pptx2svg oracle script",
)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_keeps_notes_written_by_slide_notes():
    """Notes written on a slide that had none, on one that had some, and on a duplicate:
    PowerPoint opens the deck unprompted, and what it saves back has the same notes.  (The
    export script draws slides, not notes pages, so the notes are read back from
    PowerPoint's own save.)"""
    home = Path(os.path.expanduser("~"))
    deck_path = home / "pptx-agent-notes.pptx"
    resaved = home / "pptx-agent-notes-resaved.pptx"
    pdf = home / "pptx-agent-notes.pdf"
    deck = Document.new()
    first = deck.add_slide("Title Slide")
    first.shapes[0].set_text("Notes probe")
    first.notes = "Welcome everyone.\nToday we ask for a decision."
    second = deck.add_slide("Title and Content")
    second.notes = "Draft"
    second.notes = "The cost figure comes from the finance review."
    copy = deck.duplicate_slide(second, notes="Engineering works from the Utrecht office.")
    wanted = [first.notes, deck.slide(second.slide_id).notes, copy.notes]
    deck.save(deck_path)
    try:
        exported = oracle_helper.export_pdf(deck_path, pdf)
        assert exported.ok, f"PowerPoint {exported.outcome}: {exported.detail}"
        saved = oracle_helper.save_as_pptx(deck_path, resaved)
        assert saved.ok, f"PowerPoint {saved.outcome}: {saved.detail}"
        back = Document.open(resaved)
        assert [slide.notes for slide in back.slides] == wanted
    finally:
        oracle_helper.cleanup(deck_path, resaved, pdf)
