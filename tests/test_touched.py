"""What a call changed, in a deck's addresses (round 13, R2): ``Result.touched``, for the
application -- a live view re-renders what it names -- never the model."""

from __future__ import annotations

import datetime as dt
import json

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.touched import touched
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 10, 22, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _deck() -> bytes:
    deck = Document.new()
    deck.add_slide("Title Only").title = "Plan"
    deck.add_slide("Title and Content").title = "Detail"
    deck.add_slide("Title and Content").title = "More"
    return deck.to_bytes()


@pytest.fixture
def session(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(_deck(), "plan.pptx")
    return session


def call(toolbox, session, tool_name, **arguments):
    result = toolbox.dispatch(session, tool_name, arguments)
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def _ids(session):
    deck = session.entry("d1").document
    return [slide.slide_id for slide in deck.slides]


def _shape(session, position: int, index: int = 0) -> str:
    deck = session.entry("d1").document
    return str(deck.slides[position].shapes[index].id)


def test_a_text_edit_names_its_slide_and_shape_and_the_model_reads_none_of_it(
        toolbox, session):
    first, *_ = _ids(session)
    title = _shape(session, 0)
    result = call(toolbox, session, "ppt_set_text", doc="d1",
                  items=[{"target": title, "text": "Plan, revised"}])
    assert result.touched == [{"doc": "d1", "slide": f"s:{first}", "shapes": [title]}]
    assert "touched" not in result.to_json() and '"shapes"' not in result.to_text()
    again = call(toolbox, session, "ppt_set_text", doc="d1",
                 items=[{"target": title, "text": "Plan, revised"}])
    assert again.touched == []                      # the same text: no bytes changed


def test_a_copy_reports_what_it_made_never_what_it_read(toolbox, session):
    first, second, third = _ids(session)
    source = _shape(session, 1)
    made = call(toolbox, session, "ppt_arrange", doc="d1", targets=[source],
                action="duplicate")
    copy = made.created[0]
    assert made.touched == [{"doc": "d1", "slide": f"s:{second}", "shapes": [copy],
                             "created": [copy]}]
    measured = call(toolbox, session, "ppt_add_shape", doc="d1", slide=f"s:{third}",
                    measure=True, items=[{"preset": "textbox", "like": source, "text": "a",
                                          "box": {"x": 0, "y": 0, "w": 200, "h": 0}}])
    assert measured.touched == []
    added = call(toolbox, session, "ppt_add_shape", doc="d1", slide=f"s:{third}",
                 items=[{"preset": "rect", "like": source, "text": "b",
                         "box": {"x": 10, "y": 10, "w": 100, "h": 40}}])
    assert added.touched == [{"doc": "d1", "slide": f"s:{third}", "shapes": added.created,
                              "created": added.created}]
    slide = call(toolbox, session, "ppt_manage_slides", doc="d1", action="duplicate",
                 slide=f"s:{first}")
    new = _ids(session)[1]
    assert slide.touched == [
        {"doc": "d1", "slide": f"s:{new}", "created": [f"s:{new}"]},
        {"doc": "d1", "deck": ["order"], "affects": [f"s:{second}", f"s:{third}"]}]


def test_slides_made_moved_and_deleted(toolbox, session):
    first, second, third = _ids(session)
    added = call(toolbox, session, "ppt_add_slide", doc="d1", layout="Title Only",
                 title="End")
    new = _ids(session)[-1]
    assert added.touched == [{"doc": "d1", "slide": f"s:{new}", "created": [f"s:{new}"]}]
    moved = call(toolbox, session, "ppt_manage_slides", doc="d1", action="move",
                 slide=f"s:{third}", to=1)
    assert moved.touched == [{"doc": "d1", "deck": ["order"],
                              "affects": [f"s:{third}", f"s:{first}", f"s:{second}"]}]
    deleted = call(toolbox, session, "ppt_manage_slides", doc="d1", action="delete",
                   slide=f"s:{first}")
    assert deleted.touched == [                     # third, first, second, new before
        {"doc": "d1", "slide": f"s:{first}", "removed": [f"s:{first}"]},
        {"doc": "d1", "deck": ["order"], "affects": [f"s:{second}", f"s:{new}"]}]


def test_shapes_grouped_reordered_and_deleted(toolbox, session):
    second = _ids(session)[1]
    title, body = _shape(session, 1, 0), _shape(session, 1, 1)
    extra = call(toolbox, session, "ppt_add_shape", doc="d1", slide=f"s:{second}",
                 items=[{"preset": "rect", "box": {"x": 10, "y": 10, "w": 50, "h": 50}},
                        {"preset": "ellipse", "box": {"x": 90, "y": 10, "w": 50, "h": 50}},
                        {"preset": "rect", "box": {"x": 170, "y": 10, "w": 50, "h": 50}}]
                 ).created
    back = call(toolbox, session, "ppt_arrange", doc="d1", targets=[extra[2]], action="back")
    assert back.touched == [{"doc": "d1", "slide": f"s:{second}", "shapes": [extra[2]]}]
    grouped = call(toolbox, session, "ppt_arrange", doc="d1", targets=extra[:2],
                   action="group")
    group = grouped.created[0]
    entry = grouped.touched[0]
    assert set(entry["shapes"]) == {*extra[:2], group} and entry["created"] == [group]
    deleted = call(toolbox, session, "ppt_arrange", doc="d1", targets=[title],
                   action="delete")
    assert deleted.touched == [{"doc": "d1", "slide": f"s:{second}", "shapes": [title],
                                "removed": [title]}]
    assert body not in json.dumps(deleted.touched)


def test_a_layout_edit_names_the_layout_and_the_slides_that_show_it(toolbox, session):
    first, second, third = _ids(session)
    deck = session.entry("d1").document
    layout = deck.layout("Title and Content")
    prompt = str(layout.placeholder("title").id)
    result = call(toolbox, session, "ppt_edit_master", doc="d1", action="set_prompt",
                  target=prompt, text="Say what it is")
    assert result.touched == [{"doc": "d1", "layout": layout.id, "shapes": [prompt],
                               "affects": [f"s:{second}", f"s:{third}"]}]
    on_master = call(toolbox, session, "ppt_add_shape", doc="d1", slide="M0",
                     items=[{"preset": "rect", "box": {"x": 0, "y": 0, "w": 20, "h": 20}}])
    assert on_master.touched == [{"doc": "d1", "master": "M0", "shapes": on_master.created,
                                  "created": on_master.created,
                                  "affects": [f"s:{first}", f"s:{second}", f"s:{third}"]}]
    copied = call(toolbox, session, "ppt_edit_master", doc="d1", action="duplicate_layout",
                  target=layout.id, text="Mine")
    made = copied.created[0] if copied.created else copied.touched[0]["layout"]
    assert copied.touched == [{"doc": "d1", "layout": made, "created": [made]}]


def test_the_theme_and_a_slide_put_on_another_layout(toolbox, session):
    first, second, third = _ids(session)
    theme = call(toolbox, session, "ppt_set_theme", doc="d1",
                 colors=[{"slot": "accent1", "hex": "#2A6F97"}])
    assert theme.touched == [{"doc": "d1", "theme": "M0",
                              "affects": [f"s:{first}", f"s:{second}", f"s:{third}"]}]
    moved = call(toolbox, session, "ppt_manage_slides", doc="d1", action="set_layout",
                 slide=f"s:{third}", layout="Title Only")
    entry = moved.touched[0]
    assert entry["slide"] == f"s:{third}" and entry["whole"] is True
    assert len(moved.touched) == 1


def test_notes_and_comments_are_flags_on_their_slide(toolbox, session):
    first = _ids(session)[0]
    notes = call(toolbox, session, "ppt_set_text", doc="d1",
                 items=[{"target": f"{first}/notes", "text": "Say it slowly"}])
    # The deck's first notes make its notes master too: a deck-level part, named.
    assert notes.touched == [{"doc": "d1", "slide": f"s:{first}", "notes": True},
                             {"doc": "d1", "deck": ["notesMaster"]}]
    more = call(toolbox, session, "ppt_set_text", doc="d1",
                items=[{"target": f"{first}/notes", "text": "Say it slower"}])
    assert more.touched == [{"doc": "d1", "slide": f"s:{first}", "notes": True}]
    comment = call(toolbox, session, "ppt_comments", doc="d1", action="add", author="Ana",
                   items=[{"target": _shape(session, 0), "text": "Shorter?"}])
    assert comment.touched[0] == {"doc": "d1", "slide": f"s:{first}", "comments": True}
    assert all(entry.get("slide") != f"s:{first}" or "shapes" not in entry
               for entry in comment.touched)


def test_a_chart_edit_in_its_own_part_names_the_frame_that_draws_it(toolbox, session):
    second = _ids(session)[1]
    added = call(toolbox, session, "edit_chart", doc="d1", target=f"s:{second}", action="add",
                 chart_type="bar", categories=["North", "South"],
                 data=[{"name": "Units", "values": [3, 5]}],
                 box={"x": 40, "y": 120, "w": 400, "h": 250})
    frame = added.created[0]
    titled = call(toolbox, session, "edit_chart", doc="d1", target=frame, action="set_title",
                  text="Units by region")
    assert titled.touched == [{"doc": "d1", "slide": f"s:{second}", "shapes": [frame]}]


def test_a_batch_reports_the_union_and_each_op_and_undo_and_redo_theirs(toolbox, session):
    first, second, _ = _ids(session)
    a, b = _shape(session, 0), _shape(session, 1)
    result = call(toolbox, session, "batch", ops=[
        {"tool": "ppt_set_text", "arguments": {"doc": "d1", "items": [{"target": a,
                                                                       "text": "One"}]}},
        {"tool": "ppt_set_shape", "arguments": {"doc": "d1", "items": [{"target": "999.9",
                                                                        "x": 10}]}},
        {"tool": "ppt_set_shape", "arguments": {"doc": "d1", "items": [{"target": b,
                                                                        "x": 30}]}}])
    assert result.data["failed"] == [1]
    union = [{"doc": "d1", "slide": f"s:{first}", "shapes": [a]},
             {"doc": "d1", "slide": f"s:{second}", "shapes": [b]}]
    assert result.touched == union
    assert result.touched_by_op == {0: [union[0]], 2: [union[1]]}
    assert call(toolbox, session, "undo", doc="d1").touched == union
    assert call(toolbox, session, "undo", doc="d1", redo=True).touched == union


def test_the_library_maps_a_change_it_watched():
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    slide.title = "Plan"
    title = str(slide.shapes[0].id)
    with deck.history.watch() as change:
        with deck.batch():
            slide.shapes[0].text = "Plan B"
    assert touched(deck, change) == [{"slide": f"s:{slide.slide_id}", "shapes": [title]}]
    with deck.history.watch() as nothing:
        deck.slides                                  # read
    assert touched(deck, nothing) == []
