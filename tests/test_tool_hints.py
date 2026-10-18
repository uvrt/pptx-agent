"""Round 9 (ooxml-edit 0.47.0): a name that is no deck tool names the tool and action that
do what it says (N5), with the deck's own table (``TOOL_HINTS``) for what ooxml-edit's
reading of names misses; and a batch op's ``comment`` (N4).  Made-up calls."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.hints import tool_hints

from pptx_agent.tools import FORMAT, GROUPS, TOOL_HINTS, TOOLS

from test_tools import CLOCK, _deck


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _session(toolbox, roles=None):
    session = toolbox.session(clock=CLOCK, **({"roles": roles} if roles else {}))
    session.open(_deck(), "plan.pptx")
    return session


@pytest.mark.parametrize("name, meant", [
    # read from the tools' names and actions
    ("ppt_add_chart", "edit_chart with action add"),
    ("ppt_delete_slide", "ppt_manage_slides with action delete"),
    ("ppt_copy_slide", "ppt_manage_slides with action duplicate"),
    ("ppt_insert_picture", "ppt_add_picture"),
    ("ppt_add_row", "ppt_edit_table with action insert_row"),
    ("ppt_group_shapes", "ppt_arrange with action group"),
    ("ppt_add_comment", "ppt_comments with action add"),
    ("ppt_change_layout", "ppt_manage_slides with action set_layout"),
    ("ppt_render_slide", "render"),
    # the deck's own table
    ("ppt_delete_shape", "ppt_arrange with action delete"),
    ("ppt_remove_pictures", "ppt_arrange with action delete"),
    ("ppt_move_shape", "ppt_set_shape"),            # not ppt_manage_slides' move
    ("ppt_set_title", "ppt_set_text"),              # not edit_chart's set_title
    ("ppt_add_textbox", "ppt_add_shape"),
    ("ppt_insert_image", "ppt_add_picture"),
    ("ppt_add_slide_number", "ppt_edit_master with action set_footer"),
])
def test_what_a_name_means(toolbox, name, meant):
    assert [str(hint) for hint in tool_hints(name, toolbox.tools, toolbox.tool_hints)] == [meant]


def test_every_entry_names_a_tool_and_action_the_deck_has(toolbox):
    for name, target in TOOL_HINTS.items():
        tool, _, action = target.partition(" ")
        assert tool in toolbox.tools, name
        assert not action or action in toolbox.tools[tool].actions, name
        assert name not in toolbox.tools


def test_a_batch_names_the_tool_and_action_for_an_unknown_op(toolbox):
    session = _session(toolbox)
    result = toolbox.dispatch(session, "batch", {"atomic": True, "ops": [
        {"tool": "ppt_set_text", "arguments": {"doc": "d1", "items": []},
         "comment": "fix title"},
        {"tool": "ppt_add_chart", "arguments": {"doc": "d1"}}]})
    assert result.error.message == ("ops[1]: no tool 'ppt_add_chart': use edit_chart with "
                                    "action add")
    assert result.error.valid_options[0] == "edit_chart"
    assert len(result.error.valid_options) <= 5


def test_a_reviewer_is_told_the_role_an_edit_needs(toolbox):
    session = _session(toolbox, ["review"])
    result = toolbox.dispatch(session, "ppt_delete_slide", {"doc": "d1"})
    assert result.error.message == ("no tool 'ppt_delete_slide': ppt_manage_slides with action "
                                    "delete does that, and needs the edit role; this session "
                                    "has review")


def test_an_action_of_another_tool(toolbox):
    session = _session(toolbox)
    result = toolbox.dispatch(session, "edit_chart", {"doc": "d1", "action": "delete_slide",
                                                      "target": "256.2"})
    assert result.error.message.endswith("; use ppt_manage_slides with action delete")


def test_a_batch_op_comment_runs_nothing_and_is_logged(toolbox):
    session = _session(toolbox)
    result = toolbox.dispatch(session, "batch", {"ops": [
        {"tool": "ppt_manage_slides", "arguments": {"doc": "d1", "action": "duplicate",
                                                    "slide": "s:256"},
         "comment": "a copy to rework", "note": 3}]})
    assert result.ok, result.error
    assert session.log[-1].comments == [{"op": 0, "comment": "a copy to rework", "note": "3"}]
