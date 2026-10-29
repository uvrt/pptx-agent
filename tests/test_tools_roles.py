"""The deck tools by role (ooxml-edit 0.45.0, round 7, item 9): what each role set is sent,
what it may call, and that all roles -- the default -- is the toolbox as without roles.

Offline numbers (characters / 3.5, as test_tools.py's guards; the online count is in
test_tools_online.py).  The default is unchanged: core 3,556, every definition 12,926.
A reviewer, ``["review"]``: the core 2,204 (about 3,262 counted by the 1.48 proxy; ten
tools: open_document, save_document, find_text, render, check, describe, ppt_read_slides,
batch, and ppt_compare_slides and ppt_comments, loaded in the core of a set without edit),
every definition 4,452 (15 tools).  ``["review", "resolve"]``: 2,212 and 4,460 (resolve,
reopen and delete in ppt_comments' action enum).  Editing a comment's text is every role's
(the user's decision, 2026-10-17), so ppt_comments offers ``edit`` in every set.  No deck
tool settles tracked changes: ``accept`` adds nothing (``["review", "accept"]`` is
``["review"]``'s definitions; ``["accept"]`` alone gets ppt_comments for its text edits,
1,097).  ``["edit"]``: core 2,400, every definition 11,289.  Claude's strict set for a reviewer is ppt_comments and
ppt_manage_slides (its find_by_title action alone); the default's and a builder's
(``["review", "edit"]``) are unchanged.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import ROLES, Toolbox, adapters
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from pptx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tools import (CLOCK, COUNTED_PER_ESTIMATED, CORE_COUNTED_BUDGET, _deck,
                        estimate_tokens)

#: A reviewer's core (``["review", "resolve"]``, the larger of the two), estimated: a
#: regression guard as ALL_GUARD is.  Measured 2,212 (2,210 before comment edits were every
#: role's); round 8, M3 and M4 (layouts and masters in ppt_read_slides' ``slides``, ``name``
#: in ppt_compare_slides' description) 2,220; round 9, N1 (render's ``sheet``, +64) 2,284;
#: round 11, P3 (tables in ppt_compare_slides' description, +14) 2,298; round 12, Q1
#: (check's ``spacing``, +35) 2,333; round 14, S1 (text layout in ppt_compare_slides'
#: description, +19) 2,352; round 15, T2 and T3 (check's ``contrast`` +37, ppt_read_slides'
#: description +12) 2,402; round 16, F11 (save_document's ``clean``, +42) 2,444; round 16,
#: F8 (check's ``layout``, ``min_size`` and ``styles``, documents' only, +67) 2,511; round 17,
#: P2 (ppt_compare_slides' ``reference``, +82) 2,593; round 23, V3 (check, find_text,
#: replace_text, render and save_document without a document's parameters and values) 2,360.
REVIEW_CORE_GUARD = 2365
#: Every definition a reviewer is sent, estimated.  Measured 4,460 (4,458 before); round 8,
#: M3 and M4, 4,469; round 9, N1, 4,532; round 11, P3, 4,546; round 12, Q1, 4,581; round
#: 14, S1, 4,601; round 15, T2 and T3, 4,650; round 16, F11, 4,692; F8, 4,759; round 17,
#: P1 (ppt_manage_slides' import), 4,823; P2 (ppt_compare_slides' ``reference``), 4,905;
#: round 18, ppt_manage_slides' ``action_fields`` (a reviewer, with find_by_title only, is
#: not sent slide, to, notes, ref, layout or source_doc; -164), 4,741; round 23, V3 (the
#: shared tools without a document's parameters), 4,509.
REVIEW_ALL_GUARD = 4510

REVIEW_CORE = {"open_document", "save_document", "find_text", "render", "check", "describe",
               "ppt_read_slides", "batch", "ppt_compare_slides", "ppt_comments"}
REVIEW_TOOLS = REVIEW_CORE | {"close_document", "read_blob", "ppt_design_facts", "edit_chart",
                              "ppt_manage_slides"}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _session(toolbox, roles):
    session = toolbox.session(clock=CLOCK, roles=roles)
    session.open(_deck(), "plan.pptx")
    return session


def _sent(toolbox, roles, **options):
    return {d["name"]: d for d in toolbox.definitions("anthropic", roles=roles, **options)
            if "name" in d}


def _enum(definition):
    return definition["input_schema"]["properties"]["action"]["enum"]


# -- the default ---------------------------------------------------------------------------------


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
def test_all_roles_sends_what_the_toolbox_sent_without_roles(toolbox, provider):
    """The definitions as the adapter makes them from every tool, as before roles."""
    everything = list(toolbox.tools.values())
    for groups in (None, "core"):
        before = adapters.definitions_for(provider, everything if groups is None else
                                          [t for t in everything if t.core],
                                          **_options(toolbox, provider, groups))
        assert toolbox.definitions(provider, groups=groups) == before
        assert toolbox.definitions(provider, groups=groups, roles=ROLES) == before
    assert "Role:" not in toolbox.system_prompt()


def _options(toolbox, provider, groups):
    defer = groups is None and provider in ("anthropic", "openai-responses")
    if provider == "anthropic":
        return {"strict_first": toolbox.strict_first, "defer": defer}
    if provider == "openai-responses":
        return {"groups": toolbox.groups.values(), "namespaces": defer, "defer": defer}
    return {}


# -- what each role set is sent ------------------------------------------------------------------


def test_a_reviewer_is_sent_the_review_tools_and_actions(toolbox):
    sent = _sent(toolbox, ["review"], defer=False)
    assert set(sent) == REVIEW_TOOLS
    core = {d["name"] for d in toolbox.definitions("anthropic", groups="core",
                                                   roles=["review"])}
    assert core == REVIEW_CORE
    assert _enum(sent["ppt_comments"]) == ["list", "add", "reply", "edit"]
    assert _enum(sent["edit_chart"]) == ["read"]
    assert _enum(sent["ppt_manage_slides"]) == ["find_by_title"]
    assert {name for name, d in sent.items() if d["strict"]} == {"ppt_comments",
                                                                 "ppt_manage_slides"}
    second = _sent(toolbox, ["review", "resolve"], defer=False)
    assert set(second) == REVIEW_TOOLS
    assert _enum(second["ppt_comments"]) == ["list", "add", "reply", "resolve", "reopen",
                                             "edit", "delete"]
    assert _sent(toolbox, ["review", "accept"], defer=False) == sent


def test_a_builder_without_resolve_keeps_the_default_core_and_strict_set(toolbox):
    default = _sent(toolbox, None, defer=False)
    builder = _sent(toolbox, ["review", "edit"], defer=False)
    assert set(builder) == set(default)
    assert {n for n, d in builder.items() if d["strict"]} == \
        {n for n, d in default.items() if d["strict"]}
    assert _enum(builder["ppt_comments"]) == ["list", "add", "reply", "edit"]
    assert [d["name"] for d in toolbox.definitions("anthropic", groups="core",
                                                   roles=["review", "edit"])] == \
        [d["name"] for d in toolbox.definitions("anthropic", groups="core")]


def test_an_editor_alone_is_sent_the_changing_tools(toolbox):
    sent = _sent(toolbox, ["edit"], defer=False)
    assert {"ppt_set_text", "ppt_add_shape", "undo", "new_document", "save_document",
            "batch"} <= set(sent)
    assert not {"describe", "ppt_read_slides", "render", "check", "ppt_compare_slides"} & \
        set(sent)
    assert "read" not in _enum(sent["edit_chart"])


@pytest.mark.parametrize("roles", [["review"], ["review", "resolve"], ["resolve"], ["edit"],
                                   ["review", "edit"], ["accept"], ["review", "accept"]],
                         ids="+".join)
def test_each_role_set_s_definitions_meet_each_provider_s_rules(toolbox, roles):
    for groups in ("core", None):
        assert anthropic_problems(toolbox.definitions("anthropic", groups=groups,
                                                      roles=roles)) == []
        assert openai_problems(toolbox.definitions("openai-chat", groups=groups,
                                                   roles=roles), chat=True) == []
    assert anthropic_problems(toolbox.definitions("anthropic", roles=roles)) == []
    assert openai_problems(toolbox.definitions("openai-responses", roles=roles)) == []


def test_a_reviewer_s_definitions_are_within_their_guards(toolbox):
    for roles in (["review"], ["review", "resolve"], ["review", "resolve", "accept"]):
        core = estimate_tokens(toolbox.definitions("anthropic", groups="core", roles=roles))
        everything = estimate_tokens(toolbox.definitions("anthropic", defer=False,
                                                         roles=roles))
        assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, (roles, core)
        assert core <= REVIEW_CORE_GUARD, (roles, core)
        assert everything <= REVIEW_ALL_GUARD, (roles, everything)


def test_a_reviewer_s_prompt_has_the_role_line(toolbox):
    prompt = toolbox.system_prompt(roles=["review"])
    assert prompt.startswith(toolbox.system_prompt().rstrip("\n"))
    assert "Role: review. Do not change the document" in prompt
    assert "resolve" not in prompt.split("Role: review.")[1]
    assert "You may also resolve" in toolbox.system_prompt(roles=["review", "resolve"])


# -- what each role set may call -----------------------------------------------------------------


def test_a_reviewer_reads_comments_and_saves_but_does_not_edit(toolbox):
    session = _session(toolbox, ["review"])
    run = lambda tool_, **arguments: toolbox.dispatch(session, tool_, arguments)  # noqa: E731
    assert run("describe", doc="d1").ok and run("check", doc="d1").ok
    assert run("ppt_read_slides", doc="d1", detail="outline").ok
    found = run("ppt_manage_slides", doc="d1", action="find_by_title", title="Plan")
    assert found.ok
    added = run("ppt_comments", doc="d1", action="add", author="Rev",
                items=[{"target": "s:256", "text": "Title: one line?", "ref": "c1"}])
    assert added.ok
    assert run("ppt_comments", doc="d1", action="reply", author="Rev",
               items=[{"comment": "$c1", "text": "Also the date."}]).ok
    assert run("ppt_comments", doc="d1", action="list").ok
    saved = run("save_document", doc="d1", name="reviewed.pptx", format="pptx")
    assert saved.ok and [o.name for o in session.take_outputs()] == ["reviewed.pptx"]

    text = run("ppt_set_text", doc="d1", items=[{"target": "256.2", "text": "New"}])
    assert text.error.code == "refused" and text.error.details["role"] == "edit"
    assert text.error.message.startswith("ppt_set_text needs the edit role; this session has "
                                         "review.")
    chart = run("edit_chart", doc="d1", target="s:256", action="add", chart_type="column",
                categories=["a"], data=[{"name": "s", "values": [1]}])
    assert chart.error.code == "refused" and chart.error.valid_options == ["read"]
    delete = run("ppt_manage_slides", doc="d1", action="delete", slide="s:257")
    assert delete.error.code == "refused" and "ppt_manage_slides delete" in \
        delete.error.message
    measure = run("ppt_add_shape", doc="d1", slide="s:256", measure=True,
                  items=[{"preset": "rect", "box": {"x": 0, "y": 0, "w": 100, "h": 50},
                          "text": "x"}])
    assert measure.error.code == "refused"          # measuring stays with the edit tool
    for name, arguments in (("undo", {"doc": "d1"}), ("new_document", {"kind": "pptx"})):
        assert toolbox.dispatch(session, name, arguments).error.code == "refused", name
    resolve = run("ppt_comments", doc="d1", action="resolve", items=[{"comment": "$c1"}])
    assert resolve.error.code == "refused" and "resolve role" in resolve.error.message
    closing = run("ppt_comments", doc="d1", action="reply", author="Rev",
                  items=[{"comment": "$c1", "text": "Done.", "resolve": True}])
    assert closing.error.code == "refused" and "resolve role" in closing.error.message
    assert run("ppt_comments", doc="d1", action="list").total == 1


def test_a_second_pass_reviewer_resolves_and_deletes_comments(toolbox):
    session = _session(toolbox, ["review", "resolve"])
    run = lambda tool_, **arguments: toolbox.dispatch(session, tool_, arguments)  # noqa: E731
    added = run("ppt_comments", doc="d1", action="add", author="Rev",
                items=[{"target": "s:256", "text": "a", "ref": "a"},
                       {"target": "s:257", "text": "b", "ref": "b"}])
    assert added.ok
    assert run("ppt_comments", doc="d1", action="resolve", items=[{"comment": "$a"}]).ok
    assert run("ppt_comments", doc="d1", action="reopen", items=[{"comment": "$a"}]).ok
    assert run("ppt_comments", doc="d1", action="reply", author="Rev",
               items=[{"comment": "$a", "text": "Fixed.", "resolve": True}]).ok
    assert run("ppt_comments", doc="d1", action="delete", items=[{"comment": "$b"}]).ok
    assert run("ppt_comments", doc="d1", action="edit", items=[{"comment": "$a",
                                                                "text": "x"}]).ok
    assert run("ppt_set_shape", doc="d1", items=[{"target": "256.2"}]).error.code == "refused"


def test_a_review_batch_applies_the_comments_and_refuses_the_edits(toolbox):
    session = _session(toolbox, ["review"])
    result = toolbox.dispatch(session, "batch", {"ops": [
        {"tool": "ppt_comments", "arguments": {
            "doc": "d1", "action": "add", "author": "Rev",
            "items": [{"target": "s:256", "text": "Shorter?"}]}},
        {"tool": "ppt_set_text", "arguments": {
            "doc": "d1", "items": [{"target": "256.2", "text": "Shorter"}]}},
        {"tool": "check", "arguments": {"doc": "d1"}}]})
    assert result.ok and result.data["applied"] == 2 and result.data["failed"] == [1]
    assert "ppt_set_text needs the edit role" in result.data["ops"][1]["error"]["message"]
    atomic = toolbox.dispatch(session, "batch", {"atomic": True, "ops": [
        {"tool": "ppt_comments", "arguments": {
            "doc": "d1", "action": "add", "author": "Rev",
            "items": [{"target": "s:256", "text": "Again?"}]}},
        {"tool": "ppt_comments", "arguments": {"doc": "d1", "action": "delete",
                                               "items": [{"comment": "c:00000000"}]}}]})
    assert atomic.error.code == "refused" and atomic.error.field == "ops[1].arguments.action"
    listed = toolbox.dispatch(session, "ppt_comments", {"doc": "d1", "action": "list"})
    assert listed.total == 1


def test_every_role_by_default_runs_every_tool(toolbox):
    session = _session(toolbox, None)
    run = lambda tool_, **arguments: toolbox.dispatch(session, tool_, arguments)  # noqa: E731
    assert run("ppt_set_text", doc="d1", items=[{"target": "256.2", "text": "New"}]).ok
    added = run("ppt_comments", doc="d1", action="add", author="Rev",
                items=[{"target": "s:256", "text": "a", "ref": "a"}])
    assert run("ppt_comments", doc="d1", action="delete", items=[{"comment": "$a"}]).ok
    assert added.ok and run("undo", doc="d1").ok


# -- several agents on one deck (ooxml-edit 0.46.0, round 8, M1 and M2) ------------------------


def test_a_reviewer_view_and_a_builder_view_share_one_deck_at_once(toolbox):
    """One session, every role; each agent a view of it with its own roles and context."""
    import threading

    session = _session(toolbox, None)
    reviewer = session.view(roles=["review"], context={"agent": "reviewer-1"})
    builder = session.view(roles=["review", "edit"], context={"agent": "builder-1"})
    assert reviewer.entry("d1") is builder.entry("d1") is session.entry("d1")
    results = {}

    def review():
        results["comment"] = toolbox.dispatch(reviewer, "ppt_comments", {
            "doc": "d1", "action": "add", "author": "Rev",
            "items": [{"target": "s:256", "text": "Shorter title?", "ref": "c1"}]})
        results["refused"] = toolbox.dispatch(reviewer, "ppt_set_text", {
            "doc": "d1", "items": [{"target": "256.2", "text": "Mine"}]})

    def build():
        results["text"] = toolbox.dispatch(builder, "ppt_set_text", {
            "doc": "d1", "items": [{"target": "256.2", "text": "Shorter"}]})

    threads = [threading.Thread(target=review), threading.Thread(target=build)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=60)
    assert results["comment"].ok and results["text"].ok
    refused = results["refused"]
    assert refused.error.code == "refused" and refused.error.details == {
        "role": "edit", "roles": ["review"]}
    # Both agents' work is in the one deck; the reviewer reads the builder's text.
    read = toolbox.dispatch(reviewer, "ppt_read_slides", {"doc": "d1", "detail": "outline",
                                                         "slides": ["s:256"]})
    assert read.ok and "Shorter" in str(read.data)
    assert toolbox.dispatch(builder, "ppt_comments", {"doc": "d1", "action": "list"}).total == 1
    # The builder may not resolve the reviewer's thread; the session's log names who did what.
    resolve = toolbox.dispatch(builder, "ppt_comments", {
        "doc": "d1", "action": "resolve", "items": [{"comment": "$c1"}]})
    assert resolve.error.code == "refused" and "resolve role" in resolve.error.message
    agents = {(record.tool, record.context["agent"]) for record in session.log
              if record.context and record.ok}
    assert {("ppt_comments", "reviewer-1"), ("ppt_set_text", "builder-1")} <= agents
    # A reviewer's batch is held to its view's roles.
    batch = toolbox.dispatch(reviewer, "batch", {"ops": [
        {"tool": "check", "arguments": {"doc": "d1"}},
        {"tool": "ppt_set_text", "arguments": {
            "doc": "d1", "items": [{"target": "256.2", "text": "Again"}]}}]})
    assert batch.ok and batch.data["failed"] == [1]
    with pytest.raises(ValueError, match="never widens"):
        reviewer.view(roles=["review", "edit"])


def test_a_reviewer_s_prompt_follows_the_tools_sent(toolbox):
    sent = toolbox.definitions("anthropic", groups="core", roles=["review"])
    # The reviewer's core has ppt_comments: the line says it may comment.
    assert toolbox.system_prompt(roles=["review"], tools=sent) == \
        toolbox.system_prompt(roles=["review"])
    reading = toolbox.system_prompt(roles=["review"], tools=[
        d for d in sent if d.get("name") != "ppt_comments"])
    assert "Role: review." in reading and "comment" not in reading.split("Role: review.")[1]
    assert toolbox.system_prompt(roles=["review"], groups="core") == \
        toolbox.system_prompt(roles=["review"])
