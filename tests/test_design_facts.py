"""Design facts (LP24) and the problem facts (LP15): measured, with addresses, no verdicts."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

pytest.importorskip("pptx2svg")

PT = 12700
CLOCK = lambda: dt.datetime(2026, 10, 7, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731
SPIKE = Path.home() / "Documents/agent-work/spike"


def _slide():
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    slide.title = "Facts"
    return deck, slide


def _rainbow(slide, legend: bool = False):
    """Four like boxes in one row, each a different accent: a colour-coded set."""
    boxes = []
    for n, accent in enumerate(("accent1", "accent2", "accent3", "accent4")):
        boxes.append(slide.add_shape("roundRect", (60 + 200 * n) * PT, 150 * PT, 180 * PT,
                                     60 * PT, text=f"Stream {n + 1}", fill=accent))
    if legend:
        for n, accent in enumerate(("accent1", "accent2", "accent3", "accent4")):
            slide.add_shape("rect", 60 * PT, (400 + 20 * n) * PT, 12 * PT, 12 * PT, fill=accent)
            slide.add_textbox(76 * PT, (398 + 20 * n) * PT, 120 * PT, 16 * PT, f"Kind {n}")
    return boxes


def test_palette_groups_theme_variants_and_lists_non_theme_colours():
    deck, slide = _slide()
    slide.add_shape("rect", 60 * PT, 150 * PT, 100 * PT, 50 * PT, fill="accent2")
    slide.add_shape("rect", 200 * PT, 150 * PT, 100 * PT, 50 * PT, fill="accent2 lumMod=60% lumOff=40%")
    odd = slide.add_shape("rect", 340 * PT, 150 * PT, 100 * PT, 50 * PT, fill="#C0392B")
    facts = slide.design_facts(include=["palette"])
    families = {f.family: f for f in facts.palette}
    assert [v.color for v in families["accent2"].variants] == ["accent2", "accent2 lumMod=60% lumOff=40%"]
    assert families["accent2"].accent and families["accent2"].theme
    assert [u.color for u in facts.non_theme] == ["#C0392B"]
    assert facts.non_theme[0].shapes == [odd.id]
    # LP15: the same colour is a problem fact, with its use.
    found = slide.facts()["non_theme_colors"]
    assert {"color": "#C0392B", "use": "fill"}.items() <= found[0].items()


def test_a_colour_coded_set_without_and_with_a_legend():
    deck, slide = _slide()
    boxes = _rainbow(slide)
    facts = slide.design_facts(include=["groups"])
    rainbow = next(g for g in facts.color_groups if g.preset == "roundRect")
    assert rainbow.shapes == [b.id for b in boxes]
    assert rainbow.accent_hues == ["accent1", "accent2", "accent3", "accent4"]
    assert rainbow.color_coded and rainbow.legend is None

    deck, slide = _slide()
    _rainbow(slide, legend=True)
    facts = slide.design_facts(include=["groups"])
    rainbow = next(g for g in facts.color_groups if g.preset == "roundRect")
    assert rainbow.legend is not None
    assert rainbow.legend.covers == ["accent1", "accent2", "accent3", "accent4"]
    assert rainbow.legend.misses == []
    assert [e["text"] for e in facts.legends[0].entries] == ["Kind 0", "Kind 1", "Kind 2", "Kind 3"]


def test_bars_of_one_height_are_like_shapes():
    deck, slide = _slide()
    for n, (x, w) in enumerate(((60, 100), (180, 300), (90, 50))):
        slide.add_shape("rect", x * PT, (150 + 40 * n) * PT, w * PT, 24 * PT, text=f"Bar {n}",
                        fill="accent1")
    groups = slide.design_facts(include=["groups"]).color_groups
    assert len(groups) == 1 and len(groups[0].shapes) == 3 and not groups[0].color_coded


def test_empty_regions_and_panels():
    deck, slide = _slide()
    left, top, width, height = slide.content_area
    # One box in the left half: the right half of the content area is empty.
    slide.add_shape("rect", left, top, width // 2, height, text="Left", fill="accent1")
    facts = slide.design_facts(include=["empty"])
    biggest = facts.empty[0]
    assert biggest.share == pytest.approx(0.5, abs=0.01)
    assert biggest.box[0] == pytest.approx(left + width // 2, abs=2)
    assert facts.panels == []
    # A panel covering it: a shaded background is reported as a panel, and the gap closes.
    slide.add_shape("rect", left + width // 2, top, width // 2, height, fill="bg2")
    facts = slide.design_facts(include=["empty"])
    assert facts.empty == []
    assert facts.panels[0]["fill"] == "bg2" and 0 <= facts.panels[0]["lightness"] <= 1


def test_alignment_lines_and_near_misses():
    deck, slide = _slide()
    a = slide.add_shape("rect", 100 * PT, 150 * PT, 80 * PT, 40 * PT)
    b = slide.add_shape("rect", 100 * PT, 220 * PT, 80 * PT, 40 * PT)
    c = slide.add_shape("rect", 101.5 * PT, 290 * PT, 80 * PT, 40 * PT)
    facts = slide.design_facts(include=["alignment"], within=2 * PT)
    left = next(line for line in facts.alignment["lines"] if line.edge == "left")
    assert left.shapes == [a.id, b.id]
    misses = [m for m in facts.alignment["near_misses"] if m.edge == "left"]
    assert [(m.shape, round(m.offset / PT, 2)) for m in misses] == [(c.id, 1.5)]
    # With a tighter tolerance, there is no near-miss.
    tight = slide.design_facts(include=["alignment"], within=1 * PT)
    assert not [m for m in tight.alignment["near_misses"] if m.edge == "left"]


def test_vocabulary_and_text_sizes():
    deck, slide = _slide()
    slide.add_shape("rect", 60 * PT, 150 * PT, 100 * PT, 50 * PT)
    slide.add_shape("roundRect", 200 * PT, 150 * PT, 100 * PT, 50 * PT,
                    adjustments={"adj": 20000})
    line = slide.add_connector("straight", (60 * PT, 300 * PT), (400 * PT, 300 * PT))
    line.line.dash = "dash"
    box = slide.add_textbox(60 * PT, 350 * PT, 200 * PT, 30 * PT, "Small print", autofit="none")
    box.text_frame.paragraphs[0].runs[0].size = 9
    facts = slide.design_facts(include=["vocabulary", "text_sizes"])
    v = facts.vocabulary
    assert v["presets"]["rect"] >= 1 and v["presets"]["roundRect"] == 1
    assert v["corners"]["rounded"] == [(10 * PT, 1)]   # 20% of the 50 pt side
    assert v["dashes"].get("dash") == 1
    assert sum(v["connectors"].values()) == 1
    nine = next(s for s in facts.text_sizes if s["size"] == 9.0)
    assert nine["shapes"] == [box.id]


def test_lines_over_text_with_their_z_order():
    deck, slide = _slide()
    bar = slide.add_shape("rect", 100 * PT, 200 * PT, 300 * PT, 30 * PT,
                          text="Supplier negotiations", fill="accent1")
    line = slide.add_connector("straight", (250 * PT, 150 * PT), (250 * PT, 300 * PT))
    facts = slide.design_facts(include=["z_order"])
    assert [(z.line, z.shape, z.in_front) for z in facts.z_order] == [(line.id, bar.id, True)]
    assert slide.collisions()                  # in front: a collision too
    line.send_to_back()
    facts = slide.design_facts(include=["z_order"])
    assert [(z.line, z.shape, z.in_front) for z in facts.z_order] == [(line.id, bar.id, False)]
    assert not slide.collisions()              # behind an opaque bar: hidden, still a fact


def test_region_limits_the_shapes():
    deck, slide = _slide()
    _rainbow(slide)
    region = (40 * PT, 140 * PT, 400 * PT, 100 * PT)        # the first two boxes
    facts = slide.design_facts(region=region, include=["groups"])
    assert len(facts.color_groups[0].shapes) == 2


def test_facts_report_title_wrap_margins():
    deck, slide = _slide()
    facts = slide.facts()
    titles = [w for w in facts["wrap_margins"] if w["title"]]
    assert titles and titles[0]["margin"] > 0


def test_unknown_kind_is_refused():
    deck, slide = _slide()
    with pytest.raises(ValueError):
        slide.design_facts(include=["taste"])


# -- an application's own rule, written against the facts (not shipped) --------------------------


def app_rule(facts) -> list[dict]:
    """An example rule an application might write: more than two accent hues on a set of
    like shapes, with no legend covering them."""
    findings = []
    for group in facts["groups"]["sets"]:
        if len(group["accent_hues"]) > 2 and group["legend"] is None:
            findings.append({"shapes": group["shapes"],
                             "finding": "more than 2 accent hues on like shapes and no "
                                        "legend: recolour to one accent and its tints, or add "
                                        "a legend"})
    return findings


def test_an_app_rule_over_the_tool_facts():
    deck, slide = _slide()
    _rainbow(slide)
    data = slide.design_facts().to_json()
    assert len(app_rule(data)) == 1
    deck, slide = _slide()
    _rainbow(slide, legend=True)
    assert app_rule(slide.design_facts().to_json()) == []


# -- the tool ------------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


@pytest.fixture
def session(toolbox):
    deck, slide = _slide()
    _rainbow(slide)
    session = toolbox.session(clock=CLOCK)
    session.open(deck.to_bytes(), "facts.pptx")
    return session


def test_the_tool_reports_facts_with_addresses(toolbox, session):
    result = toolbox.dispatch(session, "ppt_design_facts", {"doc": "d1", "slide": "s:256"})
    assert result.ok, result.to_text()
    data = result.data
    assert set(data) >= {"palette", "groups", "empty", "alignment", "vocabulary", "text_sizes",
                         "lines_over_text"}
    sets = data["groups"]["sets"]
    assert sets[0]["accent_hues"] == ["accent1", "accent2", "accent3", "accent4"]
    assert all(s.startswith("256.") for s in sets[0]["shapes"])
    assert len(result.to_text()) < 8000
    only = toolbox.dispatch(session, "ppt_design_facts",
                            {"doc": "d1", "slide": "s:256", "include": ["text_sizes"],
                             "region": {"x": 0, "y": 0, "w": 960, "h": 540}, "within": 1})
    assert only.ok and set(only.data) == {"slide", "area", "text_sizes"}
    bad = toolbox.dispatch(session, "ppt_design_facts", {"doc": "d1", "slide": "s:999"})
    assert not bad.ok and bad.error.code == "not_found"


def test_the_tool_in_a_batch_with_a_ref(toolbox, session):
    result = toolbox.dispatch(session, "batch", {"ops": [
        {"tool": "ppt_add_shape", "arguments": {"doc": "d1", "slide": "s:256", "items": [
            {"preset": "rect", "box": {"x": 60, "y": 300, "w": 100, "h": 40}, "fill": "#C0392B",
             "ref": "odd"}]}},
        {"tool": "ppt_design_facts", "arguments": {"doc": "d1", "slide": "s:256",
                                                    "include": ["palette"]}}]})
    assert result.ok, result.to_text()
    palette = result.data["ops"][1]["data"]["palette"]
    assert palette["non_theme"][0]["color"] == "#C0392B"
    assert palette["non_theme"][0]["shapes"] == [result.refs["odd"]]


def test_check_reports_facts_and_design(toolbox, session):
    result = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["facts"]})
    assert result.ok, result.to_text()
    entry = result.data["facts"][0]
    assert entry["slide"] == "s:256" and "non_theme_colors" in entry and "wrap_margins" in entry
    assert entry["design"]["sets"][0]["accent_hues"] == ["accent1", "accent2", "accent3", "accent4"]
    assert entry["design"]["sets"][0]["legend"] is None
    assert "notes" not in result.data
    design = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["design"]})
    entry = design.data["facts"][0]
    assert "non_theme_colors" not in entry and "design" in entry
    app = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["app"]})
    assert app.data["notes"]


# -- the spike's outputs: what the graders complained about, as facts ------------------------------


def _spike(task: str, run: str) -> Path | None:
    for root in (SPIKE / "mirror" / task / run / "work", SPIKE / "armC" / task / run / "work"):
        found = sorted(root.glob("halden-proposal-*.pptx")) if root.exists() else []
        if found:
            return found[0]
    return None


def _facts(task: str, run: str, number: int):
    path = _spike(task, run)
    if path is None:
        pytest.skip(f"the spike's {task} {run} output is not on this machine")
    return Document.open(path).slides[number - 1].design_facts()


def test_spike_o1_a1_rainbow_is_a_four_accent_set_without_legend():
    facts = _facts("o1-governance-chart", "A1", 2)
    four = [g for g in facts.color_groups if len(g.accent_hues) == 4]
    assert four and four[0].legend is None and len(four[0].shapes) == 4


def test_spike_p8_b1_colour_coded_bars_have_no_legend():
    facts = _facts("p8-engagement-plan", "B1", 3)
    bars = max(facts.color_groups, key=lambda g: len(g.shapes))
    assert len(bars.accent_hues) >= 4 and bars.legend is None


def test_spike_m1_b1_legend_is_found_and_the_plot_leaves_space():
    facts = _facts("m1-priority-matrix", "B1", 3)
    bubbles = next(g for g in facts.color_groups if g.preset == "ellipse")
    assert bubbles.legend is not None and bubbles.legend.misses == []
    assert facts.empty and facts.empty[0].share > 0.25


def test_spike_p7_b1_dead_band_is_an_empty_region_and_corners_are_mixed():
    facts = _facts("p7-phased-approach", "B1", 2)
    band = facts.empty[0]
    assert band.share > 0.2
    assert facts.vocabulary["corners"]["square"] and facts.vocabulary["corners"]["rounded"]


def test_spike_m1_a1_has_no_shading_panels():
    facts = _facts("m1-priority-matrix", "A1", 3)
    assert facts.panels == [] and facts.empty[0].share > 0.1


def test_spike_p8_a1_board_line_behind_then_in_front():
    path = _spike("p8-engagement-plan", "A1")
    if path is None:
        pytest.skip("the spike's p8 A1 output is not on this machine")
    slide = Document.open(path).slides[2]
    behind = slide.design_facts(include=["z_order"]).z_order
    assert behind and not any(z.in_front for z in behind)
    slide.document.shape(behind[0].line).bring_to_front()
    front = slide.design_facts(include=["z_order"]).z_order
    assert front and all(z.in_front for z in front)
