"""The PowerPoint tools (pptx_agent.tools) through the toolbox, as a model calls them."""

from __future__ import annotations

import datetime as dt
import json
import math
from pathlib import Path

import pytest

from conftest import FIXTURE_DIR
from ooxml_edit.tools import Toolbox, shared
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, PROMPT, TOOLS

CLOCK = lambda: dt.datetime(2026, 10, 7, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
FIXTURES = Path(__file__).parent / "fixtures"
#: The spike's p8 run A1: two milestone label boxes that overlap where their text does not.
SPIKE_P8_A1 = (Path.home() / "Documents/agent-work/spike/mirror/p8-engagement-plan/A1/work"
               / "halden-proposal-plan.pptx")

#: The budgets (tool roadmap, "Budgets", after T5's rationalisation).  What binds is what a
#: request loads: the core, at most 6,000 tokens as Anthropic's count-tokens counts them
#: (test_tools_online.py, marked provider).  Offline, characters / 3.5 of the compact JSON
#: stands in for it, times 1.48: the ratio of the counted to the estimated size T5 and T5b
#: measured on Sonnet 5.5 (the deck core: 3,352 estimated, 5,097 counted; every definition:
#: 12,747 and 18,848).  Every definition, deferred ones included, is held at what the
#: rationalisation left plus a small margin: a regression guard, not a budget.  T4 raised it
#: by 400 for edit_chart's add, data-label and gap-width actions (12,812 -> 13,158), which
#: took the place of a separate ppt_add_chart; the core is unchanged.  Post-T4 lowered it:
#: ppt_layout, ppt_scale, ppt_copy, ppt_align and place went (no model called them) and
#: ppt_comments came, list_documents went (13,158 -> 10,221 estimated; the core unchanged
#: at 3,352, about 4,961 counted by the 1.48 proxy).  The radar chart type in edit_chart's
#: enum added 3 (10,224); the guard stays.  ooxml-edit 0.12's optional scope on the shared
#: undo (undo by slide) added 34 to the core (3,352 -> 3,386, about 5,011 counted by the
#: proxy, within 5,500) and to every definition (10,257); budget and guard stay.  ooxml-edit
#: 0.14's slide ids in render's and check's ``slides`` (``s:256``, ``$ref`` or a number; the
#: items a string) added 12 to the core (3,386 -> 3,398, about 5,029 counted by the proxy,
#: within 5,500) and to every definition (10,269); budget and guard stay.  Round 3's table
#: styles and Table Design flags (ppt_add_table's style, ppt_edit_table's set_style and
#: set_flags) and ppt_manage_slides' delete_empty_placeholders added 198 to every definition
#: (10,467), none to the core (3,398, about 5,029 counted); budget and guard stay.  Round 3's
#: chart formatting in the deferred edit_chart (axis scales, gridlines, series and point
#: colours and markers, series type, chart text, point labels; the radar_filled and bubble
#: types; ``scale``, ``style``, ``gridlines``, ``part``, ``labels``, ``cells``, ``secondary``)
#: added 646 to every definition (11,113), none to the core; the guard rises to 11,200 rather
#: than squeeze the descriptions.  Round 3's align and distribute in the deferred ppt_arrange
#: (``edge``, ``to``, ``axis``, ``within``, ``gap``; house rules that require equal spacing)
#: and SVG pictures in ppt_add_picture's description added 196 to every definition (11,309),
#: none to the core (3,398, about 5,029 counted); the guard rises to 11,400.  With eight
#: optional parameters ppt_arrange no longer fits Claude's strict limit of 24 optional
#: parameters beside the others, so it is sent non-strict (the dispatcher still validates
#: it) and ppt_add_table takes its place: 29 free-text strings and 24 optional parameters
#: over the strict tools, within the limits of 32 and 24.  Round 3's master and layout
#: editing (G5) adds the deferred ppt_edit_master (prompts, Header & Footer, the layout
#: list; about 430), layout and master addresses in the prompt and decks' ``stories``
#: (layouts, masters, all) on find_text and replace_text: 11,779 for every definition, the
#: core 3,440 (about 5,091 counted); the guard rises to 11,800.  The strict set is unchanged:
#: ``stories`` was already an optional parameter of replace_text, only its enum grew.  On
#: 2026-10-11 the user raised the core budget from 5,500 to 6,000 counted tokens per format.
#: ooxml-edit 0.18's workbook kind in the shared tools (``describe``'s and ``find_text``'s
#: descriptions, find_text's ``range``, ``xlsx`` in new_document's ``kind``, ``xlsx``,
#: ``xlsm`` and ``xltx`` in save_document's ``format``) added 64 to the core (3,440 -> 3,504,
#: about 5,186 counted by the proxy, within 6,000) and to every definition (11,844); the
#: guard rises to 11,900.  The strict set is unchanged: no new parameter, find_text's
#: ``range`` only gained a description.
#: ooxml-edit 0.22 (O3) put ``from`` and ``columns`` on the deferred edit_chart (11,897).
#: X1, tables from data (``from``, ``from_columns`` and ``header`` on ppt_add_table and
#: ppt_edit_table, rows and columns optional with ``from``) added 196 to every definition
#: (12,093), none to the core (3,504, about 5,186 counted); the guard rises to 12,100.  With
#: nine optional parameters ppt_add_table no longer fits Claude's strict limit of 24
#: optional parameters beside the others, so it is sent non-strict (the dispatcher still
#: validates it) and ppt_set_theme takes its place: 27 free-text strings and 24 optional
#: parameters over the strict tools.
#: ooxml-edit 0.26.0 (round 5, T1): ``batch`` applies its ops one by one by default and
#: takes ``atomic`` for all or none; its description and the new parameter add 22 to the
#: core (3,504 -> 3,526, about 5,218 counted by the proxy) and to every definition
#: (12,096 -> 12,119); the guard rises to 12,120.  The strict set is unchanged: ``batch``
#: is never strict.
CHARS_PER_TOKEN = 3.5
COUNTED_PER_ESTIMATED = 1.48
CORE_COUNTED_BUDGET = 6000
ALL_GUARD = 12120


def estimate_tokens(definitions) -> int:
    return math.ceil(sum(len(json.dumps(d, separators=(",", ":"), ensure_ascii=False))
                         for d in definitions) / CHARS_PER_TOKEN)


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _deck() -> bytes:
    deck = Document.new()
    deck.add_slide("Title Only").title = "Plan"
    deck.add_slide("Title and Content").title = "Detail"
    return deck.to_bytes()


@pytest.fixture
def session(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(_deck(), "plan.pptx")
    return session


def call(toolbox, session, tool_name, **arguments):
    return toolbox.dispatch(session, tool_name, arguments)


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


# -- definitions ---------------------------------------------------------------------------------


EXPECTED_CORE = {"open_document", "new_document", "save_document", "undo", "find_text",
                 "replace_text", "render", "check", "batch", "describe", "ppt_read_slides",
                 "ppt_set_text", "ppt_set_shape"}


def test_the_tool_set_and_its_groups(toolbox):
    names = set(toolbox.tools)
    assert {t.name for t in toolbox.select("core")} == EXPECTED_CORE
    assert names >= EXPECTED_CORE | {
        "close_document", "set_properties", "ppt_comments",
        "ppt_format_text", "ppt_add_shape", "ppt_add_connector",
        "ppt_arrange", "ppt_add_picture", "ppt_add_table", "ppt_edit_table",
        "ppt_add_slide", "ppt_draft_slides", "ppt_manage_slides",
        "ppt_set_theme"}
    groups = {}
    for tool in toolbox.tools.values():
        groups.setdefault(tool.group, []).append(tool.name)
    assert all(len(members) < 10 for group, members in groups.items() if group != "core")
    assert set(groups) - {"core"} <= {group.name for group in GROUPS}


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
def test_the_definitions_meet_each_providers_rules(toolbox, provider):
    for groups in ("core", None):
        definitions = toolbox.definitions(provider, groups=groups)
        if provider == "anthropic":
            assert anthropic_problems(definitions) == []
        else:
            assert openai_problems(definitions, chat=provider == "openai-chat") == []
    deferred = toolbox.definitions("anthropic", defer=True)
    assert anthropic_problems(deferred) == []
    spaced = toolbox.definitions("openai-responses", defer=True, namespaces=True)
    assert openai_problems(spaced) == []


def test_the_loaded_core_is_within_budget_and_every_definition_within_its_guard(toolbox):
    core = estimate_tokens(toolbox.definitions("anthropic", groups="core"))
    everything = estimate_tokens(toolbox.definitions("anthropic", defer=False))
    assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, core
    assert everything <= ALL_GUARD, everything


def test_no_tool_takes_a_path_and_the_prompt_has_no_house_style():
    from ooxml_edit.tools.schema import is_path_name, walk_properties

    for tool in TOOLS:
        for dotted, _, _ in walk_properties(tool.canonical):
            assert not is_path_name(dotted.split(".")[-1].rstrip("[]")), (tool.name, dotted)
    words = PROMPT.lower()
    for rule in ("palette", "rainbow", "legend", "font size must", "at most two"):
        assert rule not in words


def test_the_shared_tools_are_the_shared_definitions():
    for tool in TOOLS:
        if tool.name in shared.SPECS:
            assert tool.same_definition(shared.definition(tool.name))


# -- reading -------------------------------------------------------------------------------------


def test_describe_lists_slides_theme_and_layouts(toolbox, session):
    data = ok(call(toolbox, session, "describe", doc="d1")).data
    assert [s["id"] for s in data["slides"]] == ["s:256", "s:257"]
    assert data["slides"][0]["title"] == "Plan" and data["size"] == {"w": 960.0, "h": 540.0}
    assert data["slides"][1]["content_area"]["w"] > 0
    assert data["theme"]["roles"]["primary"] == "dk2"
    assert "accent1" in data["theme"]["tints"]
    assert any(layout["name"] == "Blank" for layout in data["layouts"])


def test_read_slides_outline_and_geometry_page(toolbox, session):
    outline = ok(call(toolbox, session, "ppt_read_slides", doc="d1", detail="outline")).data
    assert "<!-- 256.2 -->" in outline
    ops = {"slide": "s:256", "items": [{"preset": "rect",
                                        "box": {"x": 10 * i, "y": 300, "w": 8, "h": 8}}
                                       for i in range(70)]}
    ok(call(toolbox, session, "ppt_add_shape", doc="d1", **ops))
    first = ok(call(toolbox, session, "ppt_read_slides", doc="d1", detail="geometry",
                    slides=[1]))
    assert first.total == 71 and first.next_cursor
    rest = ok(call(toolbox, session, "ppt_read_slides", doc="d1", detail="geometry",
                   slides=[1], cursor=first.next_cursor))
    assert len(first.data) + len(rest.data) == 71 and rest.next_cursor is None
    assert first.data[1]["box"] == {"x": 0.0, "y": 300.0, "w": 8.0, "h": 8.0}
    assert first.data[1]["fill"]["color"] == "accent1"


# -- building ------------------------------------------------------------------------------------


def test_forty_shapes_and_their_connectors_by_ref_in_one_call(toolbox, session):
    entry = session.entry("d1")
    before = entry.version
    shapes = [{"preset": "roundRect", "box": {"x": 20 + (i % 10) * 90, "y": 160 + (i // 10) * 80,
                                              "w": 70, "h": 40},
               "text": f"Step {i + 1}", "ref": f"s{i}"} for i in range(40)]
    links = [{"kind": "straight", "from": {"shape": f"$s{i}"}, "to": {"shape": f"$s{i + 1}"},
              "line": {"end": "triangle"}} for i in range(39)]
    result = ok(call(toolbox, session, "batch", ops=[
        {"tool": "ppt_add_shape", "arguments": {"doc": "d1", "slide": "s:256", "items": shapes}},
        {"tool": "ppt_add_connector", "arguments": {"doc": "d1", "items": links}}]))
    assert entry.version == before + 1
    assert len(result.created) == 79 and len(result.refs) == 40
    deck = entry.document
    link = deck.shape(result.created[40])
    assert link.begin_connection[0].id == result.refs["s0"]
    assert link.end_connection[0].id == result.refs["s1"]
    assert result.checks["slides"] == ["s:256"] and "overflows" in result.checks
    ok(call(toolbox, session, "undo", doc="d1"))
    assert len(deck.slides[0].shapes) == 1


def test_a_batch_skips_a_failing_op_and_applies_the_rest_as_one_step(toolbox, session):
    # Round 5, T1: a bad id or enum in one op no longer loses the whole batch.
    entry = session.entry("d1")
    deck = entry.document
    before = entry.version
    result = ok(call(toolbox, session, "batch", ops=[
        {"tool": "ppt_add_shape", "arguments": {"doc": "d1", "slide": "s:256", "items": [
            {"preset": "rect", "box": {"x": 40, "y": 160, "w": 120, "h": 50}, "ref": "a"}]}},
        {"tool": "ppt_add_shape", "arguments": {"doc": "d1", "slide": "s:256", "items": [
            {"preset": "hexagon-ish", "box": {"x": 300, "y": 160, "w": 60, "h": 60},
             "ref": "b"}]}},
        {"tool": "ppt_add_connector", "arguments": {"doc": "d1", "items": [
            {"kind": "straight", "from": {"shape": "$a"}, "to": {"shape": "$b"}}]}},
        {"tool": "ppt_set_text", "arguments": {"doc": "d1", "items": [
            {"target": "256.999", "text": "nowhere"}]}},
        {"tool": "ppt_set_text", "arguments": {"doc": "d1", "items": [
            {"target": "$a", "text": "Kept"}]}}]))
    assert result.data["applied"] == 2 and result.data["failed"] == [1, 2, 3]
    errors = {op["op"]: op["error"] for op in result.data["ops"] if "error" in op}
    assert errors[1]["code"] == "invalid_arguments" and "rect" in errors[1]["valid_options"]
    assert errors[2]["message"].startswith("$b was to be made by ops[1], which failed")
    assert errors[3]["code"] == "not_found"
    assert "ops[1] (ppt_add_shape)" in result.summary and "ops[3] (ppt_set_text)" in result.summary
    assert entry.version == before + 1 and set(result.refs) == {"a"}
    assert deck.shape(result.refs["a"]).text_frame.text == "Kept"
    assert len(deck.slides[0].shapes) == 2
    ok(call(toolbox, session, "undo", doc="d1"))          # one step takes back what applied
    assert len(deck.slides[0].shapes) == 1 and entry.refs == {}


def test_an_atomic_batch_is_all_or_none(toolbox, session):
    entry = session.entry("d1")
    before = entry.version
    result = call(toolbox, session, "batch", atomic=True, ops=[
        {"tool": "ppt_add_shape", "arguments": {"doc": "d1", "slide": "s:256", "items": [
            {"preset": "rect", "box": {"x": 40, "y": 160, "w": 120, "h": 50}}]}},
        {"tool": "ppt_set_text", "arguments": {"doc": "d1", "items": [
            {"target": "256.999", "text": "nowhere"}]}}])
    assert result.error.code == "not_found" and "nothing in the batch" in result.error.message
    assert entry.version == before and len(entry.document.slides[0].shapes) == 1


def test_a_failing_item_changes_nothing(toolbox, session):
    before = session.entry("d1").document.to_bytes()
    result = call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 0, "y": 0, "w": 50, "h": 50}},
        {"preset": "other", "preset_name": "blob", "box": {"x": 0, "y": 0, "w": 50, "h": 50}}])
    assert result.error.code == "invalid_arguments" and "roundRect" in result.error.valid_options
    assert session.entry("d1").document.to_bytes() == before


def test_measure_then_build_agree(toolbox, session):
    spec = [{"runs": [{"text": "Discover", "bold": True, "size": 16}], "align": "center"},
            {"runs": [{"text": "Interviews, data audit, stakeholder map"}],
             "bullet": "bullet", "level": 1}]
    for preset in ("rect", "roundRect", "chevron", "textbox"):
        measuring = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256",
                            measure=True, items=[{"preset": preset, "paragraphs": spec,
                                                  "box": {"x": 0, "y": 0, "w": 180, "h": 0}}]))
        assert not measuring.created and measuring.version == 0      # nothing was added
        measured = measuring.data["measured"][0]
        made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
            {"preset": preset, "box": {"x": 40, "y": 160, "w": 180, "h": measured["box_height"]},
             "paragraphs": spec}]))
        fit = session.entry("d1").document.shape(made.created[0]).text_fit()
        assert round(fit.needed / 12700, 2) == measured["text_height"]
        assert fit.needed <= fit.available
        ok(call(toolbox, session, "undo", doc="d1"))


def test_measuring_checks_a_height_measures_like_a_shape_and_needs_text(toolbox, session):
    made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 30}, "text": "A",
         "ref": "a"}]))
    measured = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", measure=True,
                       items=[{"preset": "rect", "text": "One\nTwo\nThree\nFour",
                               "box": {"x": 0, "y": 0, "w": 200, "h": 30}},
                              {"preset": "rect", "like": "$a", "text": "Short",
                               "box": {"x": 0, "y": 0, "w": 200, "h": 300}}])).data["measured"]
    assert measured[0]["fits"] is False and len(measured[0]["lines"]) == 4
    assert measured[1]["fits"] is True and measured[1]["lines"] == ["Short"]
    assert len(session.entry("d1").document.slides[0].shapes) == len(made.created) + 1
    bare = call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", measure=True,
                items=[{"preset": "rect", "box": {"x": 0, "y": 0, "w": 200, "h": 0}}])
    assert bare.error.code == "invalid_arguments" and bare.error.field == "items[0].text"


def test_set_text_keeps_formatting_and_reports_fit(toolbox, session):
    made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60},
         "paragraphs": [{"runs": [{"text": "Revenue 4,285", "bold": True}]}], "ref": "kpi"}]))
    result = ok(call(toolbox, session, "ppt_set_text", doc="d1",
                     items=[{"target": "$kpi", "text": "Revenue 4,310"}]))
    shape = session.entry("d1").document.shape(made.refs["kpi"])
    assert shape.text_frame.paragraph(0).run(0).bold is True
    assert result.data["fits"][made.refs["kpi"]]["fits"] is True


def test_format_text_set_shape_and_arrange(toolbox, session):
    made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60}, "text": "A",
         "ref": "a"},
        {"preset": "ellipse", "box": {"x": 300, "y": 160, "w": 60, "h": 60}, "ref": "b"}]))
    ok(call(toolbox, session, "ppt_format_text", doc="d1", items=[
        {"target": "$a", "bold": True, "size": 20, "align": "left", "anchor": "top",
         "insets": {"left": 2, "top": 2, "right": 2, "bottom": 2}}]))
    deck = session.entry("d1").document
    box = deck.shape(made.refs["a"])
    assert box.text_frame.paragraph(0).run(0).size == 20.0
    assert box.text_frame.insets == (25400, 25400, 25400, 25400)
    ok(call(toolbox, session, "ppt_set_shape", doc="d1", items=[
        {"target": "$b", "x": 400, "w": 80, "fill": "accent2 lumMod=75%",
         "line": {"color": "none"}}]))
    oval = deck.shape(made.refs["b"])
    assert (oval.left, oval.width) == (400 * 12700, 80 * 12700)
    assert oval.fill.color == "accent2 lumMod=75%"
    geometry = ok(call(toolbox, session, "ppt_read_slides", doc="d1", detail="geometry")).data
    assert next(s for s in geometry if s["id"] == oval.id)["fill"]["color"] == \
        "accent2 lumMod=75%"
    grouped = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=["$a", "$b"],
                      action="group", ref="g"))
    copy = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=["$g"],
                   action="duplicate", dy=100))
    ids = [s.id for s in deck.slides[0].shapes]
    assert len(ids) == len(set(ids)) and copy.created[0] not in grouped.created


def test_arrange_aligns_and_distributes_in_one_undo_step_each(toolbox, session):
    """Round 3, A5: house rules that ask for equal spacing get one call that guarantees it."""
    entry = session.entry("d1")
    made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:257", items=[
        {"preset": "roundRect", "box": {"x": x, "y": 200 + i * 7, "w": 80, "h": 60},
         "ref": f"t{i}"} for i, x in enumerate([60, 170, 330, 420])]))
    tiles = [f"$t{i}" for i in range(4)]
    before = entry.version
    aligned = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles,
                      action="align", edge="top"))
    assert entry.version == before + 1
    assert aligned.changed == made.created[1:] and aligned.data["moved"] == aligned.changed
    assert {box["y"] for box in aligned.data["boxes"].values()} == {200.0}
    spread = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles,
                     action="distribute", axis="horizontal"))
    xs = sorted(box["x"] for box in spread.data["boxes"].values())
    assert xs == [60.0, 180.0, 300.0, 420.0]                   # the outer two stay put
    deck = entry.document
    area = deck.slides[1].content_area
    across = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles,
                     action="distribute", axis="horizontal", within="content_area"))
    boxes = sorted(across.data["boxes"].values(), key=lambda box: box["x"])
    assert boxes[0]["x"] == round(area[0] / 12700, 2)
    assert boxes[-1]["x"] + 80 == pytest.approx((area[0] + area[2]) / 12700, abs=0.01)
    gaps = {round(b["x"] - a["x"] - a["w"], 1) for a, b in zip(boxes, boxes[1:])}
    assert len(gaps) == 1
    fixed = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles,
                    action="distribute", axis="horizontal", gap=12))
    boxes = sorted(fixed.data["boxes"].values(), key=lambda box: box["x"])
    assert [b["x"] - a["x"] - a["w"] for a, b in zip(boxes, boxes[1:])] == [12.0] * 3
    middle = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles,
                     action="align", edge="center", to="slide"))
    assert {box["x"] for box in middle.data["boxes"].values()} == {440.0}
    before = entry.version
    ok(call(toolbox, session, "undo", doc="d1"))
    assert len({box["x"] for box in fixed.data["boxes"].values()}) == 4
    assert sorted(round(deck_shape.left / 12700, 2) for deck_shape in
                  (entry.document.shape(a) for a in made.created)) == \
        [box["x"] for box in boxes]
    for arguments, field in (({"action": "align"}, "edge"),
                             ({"action": "distribute"}, "axis"),
                             ({"action": "align", "edge": "left", "gap": 4}, "gap"),
                             ({"action": "group", "edge": "left"}, "edge")):
        refused = call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles, **arguments)
        assert not refused.ok and refused.error.field == field
    two = call(toolbox, session, "ppt_arrange", doc="d1", targets=tiles[:2],
               action="distribute", axis="vertical")
    assert not two.ok and "three shapes" in two.error.message


def test_connector_sides_face_each_other(toolbox, session):
    made = ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 300, "y": 150, "w": 100, "h": 40}, "ref": "top"},
        {"preset": "rect", "box": {"x": 100, "y": 300, "w": 100, "h": 40}, "ref": "low"}]))
    link = ok(call(toolbox, session, "ppt_add_connector", doc="d1", items=[
        {"kind": "elbow", "from": {"shape": "$top"}, "to": {"shape": "$low"}}]))
    connector = session.entry("d1").document.shape(link.created[0])
    assert connector.begin_connection[1] == 2 and connector.end_connection[1] == 0


def test_tables(toolbox, session):
    made = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256",
                   box={"x": 40, "y": 160, "w": 400, "h": 90}, rows=3, columns=3,
                   data=[["Item", "Q1", "Q2"], ["Revenue", "10", "12"], ["Margin", "3", "4"]],
                   ref="t"))
    assert made.data["row_labels"] == ["Item", "Revenue", "Margin"]
    edited = ok(call(toolbox, session, "ppt_edit_table", doc="d1", target="$t",
                     action="set_cells", cells=[{"row_label": "Margin", "col_label": "Q2",
                                                 "text": "5"}]))
    assert edited.data["cells"][2][2] == "5"
    missing = call(toolbox, session, "ppt_edit_table", doc="d1", target="$t",
                   action="set_cells", cells=[{"row_label": "EBIT", "col_label": "Q2",
                                               "text": "5"}])
    assert missing.error.code == "label_not_found"
    assert missing.error.valid_options == ["Item", "Revenue", "Margin"]
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target="$t", action="insert_row",
            like=2))
    ok(call(toolbox, session, "ppt_edit_table", doc="d1", target="$t", action="format_cells",
            rows=[0], fill="accent1",
            borders=[{"side": "bottom", "width": 1.5, "color": "accent2"}]))
    ok(call(toolbox, session, "ppt_format_text", doc="d1",
            items=[{"target": f"$t/cell0,{column}", "bold": True} for column in range(3)]))
    table = session.entry("d1").document.shape(made.created[0]).table
    assert table.rows == 4 and table.cell(0, 1).text_frame.paragraph(0).run(0).bold is True
    assert table.cell(0, 1).fill.color == "accent1"
    empty = call(toolbox, session, "ppt_edit_table", doc="d1", target="$t",
                 action="format_cells")
    assert empty.error.code == "invalid_arguments"


def test_slides_notes_and_theme(toolbox, session):
    added = ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Title and Content",
                    at=1, title="Agenda", body="Scope\n  - in\n  - out\nPlan", notes="Open.",
                    ref="agenda"))
    deck = session.entry("d1").document
    assert deck.slides[0].title == "Agenda" and deck.slides[0].notes == "Open."
    body = deck.slides[0].shapes[1]
    assert [p.level for p in body.text_frame.paragraphs] == [0, 1, 1, 0]
    ok(call(toolbox, session, "ppt_set_text", doc="d1",
            items=[{"target": "$agenda/notes", "text": "Then close."}]))
    assert deck.slides[0].notes == "Then close."
    moved = ok(call(toolbox, session, "ppt_manage_slides", doc="d1", action="move",
                    slide="$agenda", to=3))
    assert moved.data["order"][-1] == added.data["slide"]
    found = ok(call(toolbox, session, "ppt_manage_slides", doc="d1", action="find_by_title",
                    title="detail"))
    assert found.data["slide"] == "s:257"
    drafted = ok(call(toolbox, session, "ppt_draft_slides", doc="d1",
                      markdown="# Risks\n\n- Scope creep\n- Budget\n"))
    assert len(drafted.created) == 1
    theme = ok(call(toolbox, session, "ppt_set_theme", doc="d1",
                    colors=[{"slot": "accent1", "hex": "#0B6E79"}],
                    fonts={"major": "Georgia"}))
    assert theme.data["colors"]["accent1"] == "#0B6E79"
    assert theme.data["fonts"]["major"] == "Georgia"


def test_pictures_from_blobs(toolbox, session):
    png = Document.open(_deck()).slides[0].render_png(width=200)
    handle = session.add_blob(png, "logo.png")
    made = ok(call(toolbox, session, "ppt_add_picture", doc="d1", slide="s:256", image=handle,
                   box={"x": 800, "y": 20, "w": 100, "h": 56.25}, ref="logo"))
    assert made.data["native_px"] == [200, 113]
    wide = Document.open(_deck()).slides[0].render_png(width=300, height=120)
    other = session.add_blob(wide, "wide.png")
    ok(call(toolbox, session, "ppt_add_picture", doc="d1", target="$logo", image=other,
            keep="height", anchor="top_right"))
    picture = session.entry("d1").document.shape(made.created[0])
    assert picture.height == round(56.25 * 12700) and picture.image_size.width == 300


# -- the shared tools on a deck ------------------------------------------------------------------


def test_find_and_replace(toolbox, session):
    found = ok(call(toolbox, session, "find_text", doc="d1", text="Plan"))
    assert found.total == 1 and found.data[0]["address"] == "256.2"
    many = call(toolbox, session, "replace_text", doc="d1", find="a", replace="A",
                expect="one")
    assert many.error.code == "ambiguous" and many.error.valid_options
    done = ok(call(toolbox, session, "replace_text", doc="d1", find="Plan", replace="Roadmap",
                   expect="one"))
    assert done.changed == ["256.2"] and done.checks["slides"] == ["s:256"]
    wrong = call(toolbox, session, "find_text", doc="d1", text="x", range="p:1")
    assert wrong.error.code == "invalid_arguments" and wrong.error.field == "range"


def test_undo_gives_back_the_original_bytes(toolbox, session):
    entry = session.entry("d1")
    original = entry.document.to_bytes()
    ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60}, "text": "A"}]))
    ok(call(toolbox, session, "ppt_set_text", doc="d1",
            items=[{"target": "256.2", "text": "Changed"}]))
    ok(call(toolbox, session, "undo", doc="d1", steps=2))
    assert entry.document.to_bytes() == original


def test_undo_by_slide_leaves_the_other_slides_changes(toolbox, session):
    entry = session.entry("d1")
    original = entry.document.to_bytes()
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": "256.2", "text": "A"}]))
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": "257.2", "text": "B"}]))
    ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60}, "text": "C"}]))
    undone = ok(call(toolbox, session, "undo", doc="d1", scope="257"))
    assert undone.data["scope"] == "257"
    deck = entry.document
    assert [s.title for s in deck.slides] == ["A", "Detail"] and len(deck.slides[0].shapes) == 2
    ok(call(toolbox, session, "undo", doc="d1", scope="s:256", steps=2))
    assert entry.document.to_bytes() == original
    ok(call(toolbox, session, "undo", doc="d1", scope="257", redo=True))
    assert [s.title for s in entry.document.slides] == ["Plan", "B"]
    missing = call(toolbox, session, "undo", doc="d1", scope="s:999")
    assert missing.error.code == "not_found" and missing.error.field == "scope"


def test_undo_by_slide_refuses_a_change_entangled_with_a_later_one(toolbox, session):
    first = ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Title Only", title="X"))
    ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Title Only", title="Y"))
    refused = call(toolbox, session, "undo", doc="d1", scope=str(first.data["slide"]))
    assert refused.error.code == "entangled"
    assert any(part.endswith("presentation.xml") for part in refused.error.details["shared"])
    assert len(session.entry("d1").document.slides) == 4


def test_a_key_makes_a_retry_safe(toolbox, session):
    arguments = {"doc": "d1", "slide": "s:256", "key": "k1", "items": [
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60}}]}
    first = ok(toolbox.dispatch(session, "ppt_add_shape", arguments))
    again = ok(toolbox.dispatch(session, "ppt_add_shape", arguments))
    assert again.created == first.created and "already done" in again.summary
    assert len(session.entry("d1").document.slides[0].shapes) == 2


def test_render_takes_slides_and_caches_by_version(toolbox, session):
    first = ok(call(toolbox, session, "render", doc="d1", slides=[1, 2], width=640))
    assert [image.width for image in first.images] == [640, 640]
    again = ok(call(toolbox, session, "render", doc="d1", slides=[2], width=640))
    assert again.data["cached"] == [2]
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": "256.2", "text": "X"}]))
    fresh = ok(call(toolbox, session, "render", doc="d1", slides=[1], width=640))
    assert fresh.data["cached"] == []
    assert call(toolbox, session, "render", doc="d1").error.field == "slides"
    assert call(toolbox, session, "render", doc="d1", slides=[1, 2, 1, 2, 1]).error.code == "limit"


def test_render_and_check_take_slide_ids_refs_and_numbers(toolbox, session):
    # Production feedback: parallel loops adding slides shift positions between reading
    # them and the call; every other deck address is an id, so render and check take one.
    deck = session.entry("d1").document
    second = deck.slides[1].slide_id
    by_id = ok(call(toolbox, session, "render", doc="d1", slides=[f"s:{second}"], width=320))
    assert [image.label for image in by_id.images] == [f"slide 2 (s:{second})"]
    assert ok(call(toolbox, session, "render", doc="d1", slides=["2"], width=320)).data["cached"] == [2]
    added = ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Title Only", at=1,
                    ref="first"))
    (address,) = added.created
    # Positions moved; the id and the ref still name the same slides.
    moved = ok(call(toolbox, session, "render", doc="d1", slides=["$first", f"s:{second}"], width=320))
    assert [image.label for image in moved.images] == [f"slide 1 ({address})", f"slide 3 (s:{second})"]
    checked = ok(call(toolbox, session, "check", doc="d1", slides=[f"s:{second}", "$first"]))
    assert checked.data["slides"] == [f"s:{second}", address]
    for bad, code in (("s:9999", "not_found"), ("256.2", "invalid_arguments"), ("$nothing", "not_found")):
        error = call(toolbox, session, "render", doc="d1", slides=[bad]).error
        assert error.code == code and error.field == "slides[0]", (bad, error)
    assert call(toolbox, session, "check", doc="d1", slides=[True]).error.code == "invalid_arguments"


def test_render_reports_the_version_it_drew(toolbox, session, monkeypatch):
    """A render works on the deck's bytes with the lock let go: an edit made meanwhile does
    not wait, and the render reports (and caches under) the version it drew."""
    import threading

    started, finish = threading.Event(), threading.Event()
    real = toolbox.pool.run

    def slow(fn, *args, timeout=None, **kwargs):
        started.set()
        assert finish.wait(10)
        return real(fn, *args, timeout=timeout, **kwargs)

    monkeypatch.setattr(toolbox.pool, "run", slow)
    before = session.entry("d1").version
    out = {}
    thread = threading.Thread(target=lambda: out.update(
        result=call(toolbox, session, "render", doc="d1", slides=[1], width=320)))
    thread.start()
    assert started.wait(10)
    edit = ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[{"target": "256.2", "text": "Y"}]))
    finish.set()
    thread.join()
    rendered = ok(out["result"])
    assert edit.version == before + 1 and rendered.version == before
    assert session.entry("d1").render_cache.get((before + 1, 256, 320, ())) is None
    monkeypatch.setattr(toolbox.pool, "run", real)
    fresh = ok(call(toolbox, session, "render", doc="d1", slides=[1], width=320))
    assert fresh.version == before + 1 and fresh.data["cached"] == []


def test_render_says_which_text_its_image_leaves_out(toolbox, monkeypatch):
    """Production: without pptx2svg-fonts on a host with no CJK font, Japanese vanished
    from the render and nothing said so.  Simulated in this process: no bundle, neither
    Office's faces nor the host's font folders (the worker pool runs the render here, so
    the patches reach it)."""
    pytest.importorskip("pptx2svg.glyphs")
    monkeypatch.setenv("PPTX2SVG_OFFICE_FONTS", "0")
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    monkeypatch.setattr("pptx2svg.glyphs.system_font_dirs", lambda: [])
    monkeypatch.setattr(toolbox.pool, "run",
                        lambda fn, *args, timeout=None, **kwargs: fn(*args, **kwargs))
    session = toolbox.session(clock=CLOCK)
    session.open((FIXTURE_DIR / "real-financial-report.pptx").read_bytes(), "results.pptx")
    first = ok(call(toolbox, session, "render", doc="d1", slides=[1], width=320)).data
    cjk = [item for item in first["missing_glyphs"] if item["script"] == "CJK"]
    assert cjk and all(item["slide"] == "s:256" and item["sample"] for item in cjk)
    assert "PowerPoint draws it" in first["missing_glyphs_note"]
    again = ok(call(toolbox, session, "render", doc="d1", slides=[1], width=320)).data
    assert again["cached"] == [1] and again["missing_glyphs"] == first["missing_glyphs"]


def test_check_and_save_behind_the_gate(toolbox, session):
    ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": {"x": 100, "y": 200, "w": 94, "h": 29}, "text": "Kick-off"},
        {"preset": "textbox", "box": {"x": 191, "y": 200, "w": 94, "h": 29},
         "text": "SteerCo 1"}]))
    plain = ok(call(toolbox, session, "check", doc="d1")).data
    assert plain["collisions"] == [] and plain["validate"]["new"] == []
    boxes = ok(call(toolbox, session, "check", doc="d1", boxes=True)).data
    assert [c["detail"] for c in boxes["collisions"]] == ["box"]
    saved = ok(call(toolbox, session, "save_document", doc="d1", name="out.pptx",
                    format="pptx"))
    assert saved.data["validate"]["new"] == []
    (output,) = session.take_outputs()
    assert Document.open(output.data).slides[0].title == "Plan"
    outline = ok(call(toolbox, session, "save_document", doc="d1", name="out.md",
                      format="outline"))
    assert session.take_outputs()[0].data.startswith(b"<!--")


@pytest.mark.parametrize("fmt", ["xlsx", "xlsm", "xltx", "docx"])
def test_a_deck_refuses_to_save_as_another_kind(toolbox, session, fmt):
    # The shared format enum names every kind's formats (ooxml-edit 0.18: workbooks too).
    result = call(toolbox, session, "save_document", doc="d1", name=f"out.{fmt}", format=fmt)
    assert not result.ok and result.error.code == "invalid_arguments"
    assert result.error.field == "format" and session.take_outputs() == []


def test_save_lists_the_fit_and_collision_facts_left_as_facts_not_a_refusal(toolbox, session):
    # Trial 3: a run saved with two line-over-text collisions its checks had reported.
    # The deck's second slide has an empty body: filled, nothing is left.
    ok(call(toolbox, session, "ppt_set_text", doc="d1",
            items=[{"target": "257.3", "text": "The detail"}]))
    clean = ok(call(toolbox, session, "save_document", doc="d1", name="clean.pptx",
                    format="pptx"))
    assert "unresolved" not in clean.data and "still in the deck" not in clean.summary
    ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "textbox", "box": {"x": 100, "y": 200, "w": 200, "h": 30},
         "text": "Board decision"},
        {"preset": "textbox", "box": {"x": 110, "y": 205, "w": 200, "h": 30},
         "text": "Steering committee"},
        {"preset": "rect", "box": {"x": 400, "y": 300, "w": 60, "h": 20},
         "text": "Far too much text for a box this small to hold in one line"}]))
    found = ok(call(toolbox, session, "check", doc="d1")).data
    saved = ok(call(toolbox, session, "save_document", doc="d1", name="out.pptx",
                    format="pptx"))
    left = saved.data["unresolved"]
    assert left["collisions"] == found["collisions"] and left["collisions"]
    assert left["overflows"] == found["overflows"] and left["overflows"]
    assert "still in the deck" in saved.summary and "collisions" in saved.summary
    assert len(session.take_outputs()) == 2


def test_a_table_grown_past_the_slide_is_a_fact_that_says_where_to_split_it(toolbox, session):
    # Production feedback: a 12x3 table of wrapping text, cut off at the slide's bottom while
    # every check came back clean.
    text = "A longer cell text that wraps onto two or three lines in this column width"
    added = ok(call(toolbox, session, "ppt_add_table", doc="d1", slide="s:256",
                    box={"x": 36, "y": 108, "w": 888, "h": 346}, rows=12, columns=3,
                    data=[[f"R{r}C{c}: {text}" for c in (1, 2, 3)] for r in range(1, 13)]))
    (table,) = added.created
    (fact,) = [entry for entry in added.checks["off_slide"] if entry["shape"] == table]
    assert fact["rows_past"] == [7, 12] and fact["rows_fit"] == 6 and fact["past"] > 300
    assert "next slide" in added.checks["rows_note"]
    assert ok(call(toolbox, session, "check", doc="d1", slides=[1])).data["off_slide"] \
        == added.checks["off_slide"]
    saved = ok(call(toolbox, session, "save_document", doc="d1", name="out.pptx",
                    format="pptx"))
    assert saved.data["unresolved"]["off_slide"] == added.checks["off_slide"]


@pytest.mark.skipif(not SPIKE_P8_A1.exists(), reason="the spike's p8 A1 deck is not here")
def test_check_with_boxes_reports_the_spikes_p8_a1_label_overlap(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(SPIKE_P8_A1.read_bytes(), "p8-a1.pptx")
    plain = ok(call(toolbox, session, "check", doc="d1", slides=[3])).data
    assert plain["collisions"] == []
    boxes = ok(call(toolbox, session, "check", doc="d1", slides=[3], boxes=True)).data
    assert boxes["collisions"] == [{"detail": "box", "shape": "258.29", "other": "258.27",
                                    "area": boxes["collisions"][0]["area"]}]


def test_new_document_and_session_tools(toolbox):
    session = toolbox.session(clock=CLOCK)
    made = ok(call(toolbox, session, "new_document", kind="pptx", size="4:3", title="T"))
    assert made.data["size"] == {"w": 720.0, "h": 540.0}
    template = session.add_blob(Document.new().to_bytes(template=True), "brand.potx")
    branded = ok(call(toolbox, session, "new_document", kind="pptx", template_blob=template,
                      name="q4.pptx"))
    assert branded.data["doc"] == "d2"
    opened = ok(call(toolbox, session, "open_document", blob=template))
    assert any("TemplateOpenedWarning" in w for w in opened.warnings)
    assert list(session.documents) == ["d1", "d2", "d3"]
    assert "list_documents" not in toolbox.tools          # the application names them
    ok(call(toolbox, session, "set_properties", doc="d1", author="Ada", language="nl-NL"))
    assert session.entry("d1").document.author == "Ada"


def test_charts_and_smartart_on_a_deck(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open((FIXTURES / "real-financial-report.pptx").read_bytes(), "results.pptx")
    session.open((FIXTURES / "powerpoint-smartart.pptx").read_bytes(), "smartart.pptx")
    reading = ok(call(toolbox, session, "edit_chart", doc="d1", target="257.25",
                      action="read"))
    assert reading.version == 0 and not reading.changed     # read changes nothing
    read = reading.data
    assert read["categories"] == ["Q1", "Q2", "Q3"] and read["workbook"]["sheet"]
    edited = ok(call(toolbox, session, "edit_chart", doc="d1", target="257.25",
                     action="set_value", series="0", category="Q3", value=4310)).data
    assert edited["series"][0]["values"][2] == 4310
    wrong = call(toolbox, session, "edit_chart", doc="d1", target="257.25",
                 action="set_value", series="Profit", category="Q3", value=1)
    assert wrong.error.code == "label_not_found" and len(wrong.error.valid_options) == 2
    ok(call(toolbox, session, "edit_chart", doc="d1", target="257.25", action="set_legend",
            position="bottom"))
    nodes = ok(call(toolbox, session, "edit_smartart", doc="d2", target="256.3",
                    action="set_text", node=1, text="Build it")).data["nodes"]
    assert [n["text"] for n in nodes] == ["Plan", "Build it", "Ship"]


# -- review comments -----------------------------------------------------------------------------


def test_ppt_comments_add_reply_resolve_list_and_delete(toolbox, session):
    added = ok(call(toolbox, session, "ppt_comments", doc="d1", action="add", author="Dana",
                    items=[{"target": "s:256", "text": "Shorter title?", "ref": "t1"},
                           {"target": "s:257", "text": "Source?"}]))
    assert len(added.created) == 2 and added.refs["t1"] == added.created[0]
    ok(call(toolbox, session, "ppt_comments", doc="d1", action="reply", author="Sam",
            items=[{"comment": "$t1", "text": "Done.", "resolve": True}]))
    listed = ok(call(toolbox, session, "ppt_comments", doc="d1", action="list"))
    assert listed.total == 2 and listed.data[0]["done"] is True
    assert listed.data[0]["replies"][0]["text"] == "Done."
    assert listed.version == added.version + 1                  # list only reads
    open_only = ok(call(toolbox, session, "ppt_comments", doc="d1", action="list",
                        open_only=True))
    assert [t["text"] for t in open_only.data] == ["Source?"]
    described = ok(call(toolbox, session, "describe", doc="d1")).data
    assert described["comments"] == {"threads": 2, "open": 1, "read": "ppt_comments list"}
    ok(call(toolbox, session, "ppt_comments", doc="d1", action="delete",
            items=[{"comment": added.created[1]}]))
    assert ok(call(toolbox, session, "ppt_comments", doc="d1", action="list")).total == 1
    ok(call(toolbox, session, "undo", doc="d1"))
    assert ok(call(toolbox, session, "ppt_comments", doc="d1", action="list")).total == 2


def test_ppt_comments_errors_name_what_to_fix(toolbox, session):
    missing = call(toolbox, session, "ppt_comments", doc="d1", action="add",
                   items=[{"target": "s:256", "text": "x"}])
    assert missing.error.code == "invalid_arguments" and missing.error.field == "author"
    unknown = call(toolbox, session, "ppt_comments", doc="d1", action="resolve",
                   items=[{"comment": "c:12345678"}])
    assert unknown.error.code == "not_found"
    bad_target = call(toolbox, session, "ppt_comments", doc="d1", action="add", author="D",
                      items=[{"target": "256.99", "text": "x"}])
    assert bad_target.error.code == "not_found"


# -- aliases and refusals (round 5: T6, T7, T8) ---------------------------------------------------


def _box_shape(toolbox, session, **item):
    return ok(call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 40, "y": 160, "w": 200, "h": 60}, "ref": "box1",
         **item}]))


def test_autofit_shrink_is_read_as_normal(toolbox, session):
    made = _box_shape(toolbox, session, text="Plan", frame={"autofit": "shrink"})
    assert made.warnings[0] == "items[0].frame.autofit 'shrink' read as 'normal'"
    shape = session.entry("d1").document.shape(made.refs["box1"])
    assert shape.text_frame.autofit == "normal"
    result = ok(call(toolbox, session, "ppt_format_text", doc="d1",
                     items=[{"target": "$box1", "autofit": "Resize"}]))
    assert result.warnings[0] == "items[0].autofit 'Resize' read as 'shape'"
    assert shape.text_frame.autofit == "shape"
    # The aliases cost nothing: the definitions never carry them.
    assert "valueAliases" not in json.dumps(toolbox.definitions("anthropic"))


def test_an_autofit_close_to_a_mode_says_did_you_mean(toolbox, session):
    error = call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 0, "y": 0, "w": 9, "h": 9}, "text": "a",
         "frame": {"autofit": "Normal"}}]).error
    assert "did you mean 'normal'?" in error.message
    assert error.valid_options == ["none", "normal", "shape"]


def test_an_empty_paragraphs_list_beside_text_is_left_out(toolbox, session):
    made = _box_shape(toolbox, session, text="Plan", paragraphs=[])
    shape = session.entry("d1").document.shape(made.refs["box1"])
    assert shape.text_frame.text == "Plan"
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[
        {"target": "$box1", "text": "Next", "paragraphs": []}]))
    assert shape.text_frame.text == "Next"
    ok(call(toolbox, session, "ppt_set_text", doc="d1", items=[
        {"target": "$box1", "text": "", "paragraphs": [{"runs": [{"text": "Last"}]}]}]))
    assert shape.text_frame.text == "Last"
    both = call(toolbox, session, "ppt_set_text", doc="d1", items=[
        {"target": "$box1", "text": "a", "paragraphs": [{"runs": [{"text": "b"}]}]}])
    assert both.error.message == "give text or paragraphs, not both"


def test_a_key_on_set_text_says_setters_are_safe_to_repeat(toolbox, session):
    _box_shape(toolbox, session, text="Plan")
    error = call(toolbox, session, "ppt_set_text", doc="d1", key="k1",
                 items=[{"target": "$box1", "text": "x"}]).error
    assert error.field == "key" and "safe to repeat" in error.message


def test_office_theme_colour_names_are_read_as_scheme_colours(toolbox, session):
    made = _box_shape(toolbox, session, paragraphs=[
        {"runs": [{"text": "Plan", "color": "Text 2"}]}])
    deck = session.entry("d1").document
    shape = deck.shape(made.refs["box1"])
    assert shape.text_frame.paragraph(0).run(0).color == "tx2"
    ok(call(toolbox, session, "ppt_set_shape", doc="d1", items=[
        {"target": "$box1", "fill": "background1 lumMod=85%"}]))
    assert shape.fill.color == "bg1 lumMod=85%"
    error = call(toolbox, session, "ppt_set_shape", doc="d1", items=[
        {"target": "$box1", "fill": "text3"}]).error
    assert "tx1" in error.message and "text1" in error.message


def test_a_ref_name_with_capitals_is_refused_with_one_that_would_do(toolbox, session):
    shapes = len(session.entry("d1").document.slides[0].shapes)
    error = call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rect", "box": {"x": 0, "y": 0, "w": 90, "h": 90}, "text": "a",
         "ref": "catChart"}]).error
    assert error.code == "invalid_arguments" and error.field == "ref"
    assert "lowercase letter" in error.message and "'cat_chart'" in error.message
    assert len(session.entry("d1").document.slides[0].shapes) == shapes   # rolled back
    # On a valid edit_chart add too.
    error = call(toolbox, session, "edit_chart", doc="d1", action="add", target="s:256",
                 box={"x": 40, "y": 100, "w": 300, "h": 200}, chart_type="column",
                 categories=["A"], data=[{"name": "S", "values": [1]}], ref="catChart").error
    assert error.field == "ref" and "'cat_chart'" in error.message


def test_a_preset_called_rectangle_says_did_you_mean_rect(toolbox, session):
    error = call(toolbox, session, "ppt_add_shape", doc="d1", slide="s:256", items=[
        {"preset": "rectangle", "box": {"x": 0, "y": 0, "w": 9, "h": 9}}]).error
    assert error.field == "items[0].preset"
    assert "did you mean 'rect'?" in error.message


def test_edit_chart_takes_slide_for_target_when_adding(toolbox, session):
    result = ok(call(toolbox, session, "edit_chart", doc="d1", action="add", slide="s:256",
                     box={"x": 40, "y": 100, "w": 300, "h": 200}, chart_type="column", categories=["A", "B"],
                     data=[{"name": "S", "values": [1, 2]}]))
    assert "slide read as target" in result.warnings
