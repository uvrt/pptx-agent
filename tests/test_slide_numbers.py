"""Slide numbers by reference (round 6, R8): ``{{slide:s:258}}`` in a text tool's text is
written as that slide's number and kept current -- after adding, deleting, moving and
duplicating slides, on save, and through a save and a reopen -- as a text fact in the
shape's tags part; a reference to a deleted slide keeps its last number and is listed.

Every deck here is made up: four blank slides, s:256 to s:259, an agenda box on the first.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit import provenance as kept
from pptx_agent.edit import slidenumbers
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import call, failed, ok

BOX = {"x": 40, "y": 40, "w": 400, "h": 80}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def four_slides(first: int | None = None) -> bytes:
    deck = Document.new()
    for _ in range(4):
        deck.add_slide("Blank")
    if first is not None:
        path = deck.package.presentation_part()
        deck.package.tree(path).set("firstSlideNum", str(first))
        deck.package.mark_dirty(path)
    return deck.to_bytes()


@pytest.fixture
def session(toolbox):
    session = toolbox.session()
    session.open(four_slides(), "deck.pptx")
    return session


def deck_of(session):
    return session.entry("d1").document


def agenda(toolbox, session, text, slide="s:256"):
    result = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide=slide,
                     items=[{"preset": "textbox", "box": BOX, "text": text}]))
    return result.created[0], result


def text_of(session, shape):
    return deck_of(session).shape(shape).text_frame.text


def manage(toolbox, session, action, slide, **arguments):
    return ok(call(toolbox, session, "ppt_manage_slides", doc="d1", action=action,
                   slide=slide, **arguments))


# -- writing --------------------------------------------------------------------------------------


def test_a_reference_is_written_as_the_slides_number_and_kept(toolbox, session):
    shape, result = agenda(toolbox, session, "Results: slide {{slide:s:258}}")
    assert text_of(session, shape) == "Results: slide 3"
    assert result.data["facts"] == [{"from": "slide:s:258", "text": "3"}]
    assert kept.facts(deck_of(session).shape(shape)) == [
        {"source": "slide:s:258", "text": "3", "slide": 258, "p": 0, "at": 15}]


def test_a_ref_names_the_slide(toolbox, session):
    added = ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Blank",
                    ref="outro")).created[0]
    shape, _ = agenda(toolbox, session, "Questions: {{slide:$outro}}")
    assert text_of(session, shape) == "Questions: 5"
    assert kept.facts(deck_of(session).shape(shape))[0]["source"] == f"slide:{added}"


def test_a_number_a_shape_or_a_missing_slide_is_refused(toolbox, session):
    error = failed(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": BOX, "text": "{{slide:3}}"}]), "invalid_arguments")
    assert "slide 3 is s:258: write {{slide:s:258}}" in error.message
    assert error.field == "items[0].text"
    error = failed(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": BOX, "text": "{{slide:257.2}}"}]), "invalid_arguments")
    assert "is a shape" in error.message and "{{slide:s:257}}" in error.message
    error = failed(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": BOX, "text": "{{slide:s:999}}"}]), "not_found")
    assert error.valid_options == ["s:256", "s:257", "s:258", "s:259"]
    assert len(deck_of(session).slide(256).shapes) == 0


def test_first_slide_num_is_the_first_number(toolbox):
    session = toolbox.session()
    session.open(four_slides(first=5), "deck.pptx")
    shape, _ = agenda(toolbox, session, "{{slide:s:256}}-{{slide:s:259}}")
    assert text_of(session, shape) == "5-8"


# -- kept current -------------------------------------------------------------------------------


def test_adding_a_slide_before_it_renumbers_and_the_checks_say_so(toolbox, session):
    shape, _ = agenda(toolbox, session, "See {{slide:s:258}}, then {{slide:s:259}}.")
    result = ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Blank", at=2))
    assert text_of(session, shape) == "See 4, then 5."
    assert result.checks["slide_numbers"] == [
        {"shape": shape, "slide": "s:258", "text": "4", "was": "3"},
        {"shape": shape, "slide": "s:259", "text": "5", "was": "4"}]
    assert [f["at"] for f in kept.facts(deck_of(session).shape(shape))] == [4, 12]


def test_moving_and_undoing_take_the_number_along(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    manage(toolbox, session, "move", "s:259", to=2)
    assert text_of(session, shape) == "Close on 2"
    ok(call(toolbox, session, "undo", doc="d1"))
    assert text_of(session, shape) == "Close on 4"
    assert kept.facts(deck_of(session).shape(shape))[0]["text"] == "4"


def test_deleting_a_slide_before_it_and_duplicating_renumber(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    manage(toolbox, session, "delete", "s:257")
    assert text_of(session, shape) == "Close on 3"
    copy = manage(toolbox, session, "duplicate", "s:256").data["slide"]
    assert text_of(session, shape) == "Close on 4"
    # The copy's agenda box keeps its own fact and follows the slides too.
    copied = next(s for s in deck_of(session).slide(copy).shapes if s.text_frame is not None)
    assert copied.text_frame.text == "Close on 4"
    manage(toolbox, session, "move", "s:259", to=1)
    assert text_of(session, shape) == copied.text_frame.text == "Close on 1"


def test_a_hidden_slide_keeps_counting(toolbox, session):
    shape, _ = agenda(toolbox, session, "{{slide:s:259}}")
    deck = deck_of(session)
    hidden = deck.slide(258)
    deck.package.tree(hidden.part_path).set("show", "0")
    deck.package.mark_dirty(hidden.part_path)
    manage(toolbox, session, "move", "s:257", to=4)
    assert text_of(session, shape) == "3"           # s:258 (hidden) still counts


def test_a_deleted_target_keeps_its_last_number_and_is_listed(toolbox, session):
    shape, _ = agenda(toolbox, session, "Appendix: {{slide:s:258}}")
    result = manage(toolbox, session, "delete", "s:258")
    assert text_of(session, shape) == "Appendix: 3"
    assert result.checks["slide_numbers"] == [
        {"shape": shape, "slide": "s:258", "text": "3", "gone": True}]
    manage(toolbox, session, "move", "s:259", to=1)
    assert text_of(session, shape) == "Appendix: 3"
    checked = ok(call(toolbox, session, "check", doc="d1"))
    assert checked.data["missing_slides"] == [
        {"shape": shape, "slide": "s:258", "text": "3",
         "message": "slide s:258 no longer exists: the text keeps its last number, 3"}]
    saved = ok(toolbox.dispatch(session, "save_document",
                                {"doc": "d1", "name": "out.pptx", "format": "pptx"}))
    assert saved.data["unresolved"]["missing_slides"][0]["slide"] == "s:258"
    assert "1 missing slides" in saved.summary


def test_a_gone_slide_is_never_taken_for_a_new_one_with_its_id():
    deck = Document.open(four_slides())
    shape = deck.slide(256).add_textbox(0, 0, 914400, 914400, "x")
    kept.facts_written(shape, None, "x", "Last: 4",
                       [{"source": "slide:s:259", "text": "4", "slide": 259, "p": 0,
                         "at": 6}])
    shape.text_frame.set_text("Last: 4")
    deck.delete_slide(259)
    assert kept.facts(deck.shape(shape.id))[0]["gone"] is True
    reopened = Document.open(deck.to_bytes())
    again = reopened.add_slide("Blank", index=0)     # the top id is handed out again
    assert again.slide_id == 259
    reopened.update_slide_numbers()
    assert reopened.shape(shape.id).text_frame.text == "Last: 4"


# -- save and reopen ---------------------------------------------------------------------------


def test_the_facts_survive_a_save_and_a_reopen(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    reopened = Document.open(deck_of(session).to_bytes())
    reopened.move_slide("s:259", to=0)
    assert reopened.shape(shape).text_frame.text == "Close on 1"
    assert reopened.slide_numbers_kept[-1]["was"] == "4"


def test_saving_renumbers_a_deck_reordered_elsewhere(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    data = deck_of(session).to_bytes()
    # As PowerPoint would leave it: the slide list reordered, the text as it was.
    deck = Document.open(data)
    path = deck.package.presentation_part()
    root = deck.package.tree(path)
    entries = root.find("{*}sldIdLst")
    entries.insert(0, entries[-1])
    deck.package.mark_dirty(path)
    edited = deck.package.to_bytes()
    session.open(edited, "edited.pptx")                                       # d2
    assert session.entry("d2").document.shape(shape).text_frame.text == "Close on 4"
    saved = ok(toolbox.dispatch(session, "save_document",
                                {"doc": "d2", "name": "again.pptx", "format": "pptx"}))
    assert saved.data["slide_numbers"] == [
        {"shape": shape, "slide": "s:259", "text": "1", "was": "4"}]
    assert session.entry("d2").document.shape(shape).text_frame.text == "Close on 1"


def test_library_save_renumbers(tmp_path):
    deck = Document.open(four_slides())
    shape = deck.slide(256).add_textbox(0, 0, 914400, 914400, "2")
    kept.facts_written(shape, None, "2", "2", [{"source": "slide:s:257", "text": "2",
                                                "slide": 257, "p": 0, "at": 0}])
    path = deck.package.presentation_part()
    entries = deck.package.tree(path).find("{*}sldIdLst")
    entries.append(entries[1])                      # s:257 last, behind the library's back
    deck.package.mark_dirty(path)
    deck.save(tmp_path / "out.pptx")
    assert Document.open(tmp_path / "out.pptx").shape(shape.id).text_frame.text == "4"


# -- with other text and facts -----------------------------------------------------------------


def test_a_longer_number_moves_the_facts_after_it_and_keeps_formatting(toolbox):
    session = toolbox.session()
    deck = Document.new()
    for _ in range(10):
        deck.add_slide("Blank")
    session.open(deck.to_bytes(), "deck.pptx")
    shape = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": BOX, "paragraphs": [
            {"runs": [{"text": "Part "}, {"text": "{{slide:s:265}}", "bold": True},
                      {"text": " and {{slide:s:264}}"}]}]}])).created[0]
    assert text_of(session, shape) == "Part 10 and 9"
    ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Blank", at=1))
    assert text_of(session, shape) == "Part 11 and 10"
    runs = deck_of(session).shape(shape).text_frame.paragraphs[0].runs
    assert [(r.text, bool(r.bold)) for r in runs][:2] == [("Part ", False), ("11", True)]
    assert [(f["text"], f["at"]) for f in kept.facts(deck_of(session).shape(shape))] == [
        ("11", 5), ("10", 12)]


def test_a_table_cell_keeps_its_number_current(toolbox, session):
    frame = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256", box=BOX,
                    rows=2, columns=2,
                    data=[["Topic", "Slide"], ["Close", "{{slide:s:259}}"]])).created[0]
    manage(toolbox, session, "move", "s:259", to=2)
    assert deck_of(session).shape(frame).table.cell(1, 1).text == "2"


def test_an_edited_number_is_left_alone(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    ok(call(toolbox, session, "ppt_set_text", doc="d1",
            items=[{"target": shape, "text": "Close on 9"}]))
    manage(toolbox, session, "move", "s:259", to=1)
    assert text_of(session, shape) == "Close on 9"


def test_copied_to_another_deck_the_text_stays_and_the_fact_goes(toolbox, session):
    shape, _ = agenda(toolbox, session, "Close on {{slide:s:259}}")
    other = Document.open(four_slides())
    source = deck_of(session).slide(256)
    copied = source.copy_shapes([shape], other.slide(257)).shapes[0]
    assert copied.text_frame.text == "Close on 4"
    assert kept.facts(copied) == []
    same = source.copy_shapes([shape], deck_of(session).slide(257)).shapes[0]
    assert kept.facts(same)[0]["slide"] == 259


def test_order_reads_ids_and_the_first_number():
    deck = Document.open(four_slides(first=0))
    assert slidenumbers.order(deck) == (0, (256, 257, 258, 259))
    assert slidenumbers.numbers(deck)[259] == 3


def test_the_prompt_names_the_syntax():
    from pptx_agent.tools import PROMPT

    assert "{{slide:s:<id>}}" in PROMPT
