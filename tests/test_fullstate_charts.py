"""Charts and SmartArt in the full-state SVG (ROADMAP E4).

``data-ooxml-chart-data`` carries a chart's types, titles, legend, number format, categories
and series as compact JSON; ``data-ooxml-diagram-nodes`` a diagram's nodes (id, level,
text).  Edited in the SVG, they arrive through the E4 edits -- so the workbook moves with the
cache, and the cached SmartArt drawing with the data model -- in one undo step.
"""

from __future__ import annotations

import io
import math
import zipfile

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document
from pptx_agent.fullstate import FullStateError, Limits
from svgedit import FullStateSvg
from test_chart import check_deck_charts
from test_validity import assert_valid

pytest.importorskip("pptx2svg", reason="the full-state SVG is drawn by pptx2svg")

SMARTART = FIXTURE_DIR / "powerpoint-smartart.pptx"


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()
                if not info.is_dir()}


def _chart_slide(document: Document, shape: str) -> tuple[int, FullStateSvg]:
    index = next(i for i, s in enumerate(document.slides)
                 if any(sh.id == shape for sh in s.shapes))
    return index, FullStateSvg(document.slides[index].render_svg(full_state=True))


# ------------------------------------------------------------------------------------------
# Emitting
# ------------------------------------------------------------------------------------------


def test_a_chart_frame_carries_its_data(financial_report):
    document = Document.open(financial_report.read_bytes())
    _, svg = _chart_slide(document, "257.29")
    data = svg.json("257.29", "chart-data")
    assert data == document.shape("257.29").chart.data
    assert data["series"][1] == {"name": "売上総利益率（%）", "values": [42, 42.5, 43]}
    raw = svg.get("257.29", "chart-data")
    assert ", " not in raw and ": " not in raw  # compact


def test_a_diagram_frame_carries_its_nodes():
    document = Document.open(SMARTART.read_bytes())
    svg = FullStateSvg(document.slides[1].render_svg(full_state=True))
    model = svg.json("257.3", "diagram-nodes")
    assert [(n["lvl"], n["t"]) for n in model["nodes"]] == [
        (0, "Goals"), (1, "Faster edits"), (1, "Fewer prompts"), (0, "Risks"),
        (1, "Stale caches")]
    assert model["layout"].endswith("/layout/vList2")
    assert svg.get("257.3", "diagram") == "ppt/diagrams/data2.xml"


# ------------------------------------------------------------------------------------------
# Applying
# ------------------------------------------------------------------------------------------


def test_chart_edits_made_in_the_svg_arrive_with_the_workbook(financial_report):
    original = financial_report.read_bytes()
    document = Document.open(original)
    _, svg = _chart_slide(document, "257.29")
    data = svg.json("257.29", "chart-data")
    data["title"] = "Margins"
    data["axis_titles"]["value"] = "%"
    data["legend"] = False
    data["categories"] = ["Q1", "Q3", "Q4"]          # Q2 removed, Q4 added
    data["series"] = [
        {"name": "Operating", "values": [11.7, 11.9, 12.4]},   # renamed, values kept/added
        {"name": "純利益率（%）", "values": [7.4, 7.7, 8]},
        {"name": "New", "values": [1, 2, 3]},
    ]
    svg.set_json("257.29", "chart-data", data)

    report = document.apply_svg(str(svg))

    assert report.edited == {"257.29": ["chart-data"]}
    edited = document.to_bytes()
    assert Document.open(edited).shape("257.29").chart.data == {
        **data, "types": ["line"], "format": "General"}
    check_deck_charts(edited, original)
    assert_valid(edited, original)
    document.undo()
    assert _parts(document.to_bytes()) == _parts(original)


def test_chart_svg_round_trip_is_identity_after_an_edit(financial_report):
    document = Document.open(financial_report.read_bytes())
    document.shape("258.9").chart.add_series("Extra", [1, 2, 3, 4])
    saved = document.to_bytes()
    reopened = Document.open(saved)
    for slide in Document.open(saved).slides:
        assert not reopened.apply_svg(slide.render_svg(full_state=True))
    assert reopened.to_bytes() == saved


def test_diagram_edits_made_in_the_svg_arrive():
    original = SMARTART.read_bytes()
    document = Document.open(original)
    svg = FullStateSvg(document.slides[1].render_svg(full_state=True))
    model = svg.json("257.3", "diagram-nodes")
    model["nodes"][1]["t"] = "SVG edits"
    del model["nodes"][4]                                      # Stale caches
    model["nodes"].insert(4, {"lvl": 1, "t": "Long files"})   # a new child of Risks
    model["nodes"].append({"lvl": 0, "t": "Next"})            # a new top-level node
    svg.set_json("257.3", "diagram-nodes", model)

    document.apply_svg(str(svg))

    diagram = Document.open(document.to_bytes()).shape("257.3").diagram
    assert [(n.level, n.text) for n in diagram.nodes] == [
        (0, "Goals"), (1, "SVG edits"), (1, "Fewer prompts"), (0, "Risks"),
        (1, "Long files"), (0, "Next")]
    assert_valid(document.to_bytes(), original)
    document.undo()
    assert _parts(document.to_bytes()) == _parts(original)


def test_a_text_only_diagram_edit_keeps_the_drawing():
    document = Document.open(SMARTART.read_bytes())
    svg = FullStateSvg(document.slides[0].render_svg(full_state=True))
    model = svg.json("256.3", "diagram-nodes")
    model["nodes"][2]["t"] = "Go"  # short: the cached shape's 65 pt wraps anything long
    svg.set_json("256.3", "diagram-nodes", model)
    document.apply_svg(str(svg))
    diagram = document.shape("256.3").diagram
    assert diagram.drawing_part is not None
    assert ">Go</tspan>" in document.slides[0].render_svg()


# ------------------------------------------------------------------------------------------
# Refusals: validated before anything changes
# ------------------------------------------------------------------------------------------


@pytest.fixture
def chart_deck(financial_report):
    document = Document.open(financial_report.read_bytes())
    original = document.to_bytes()
    _, svg = _chart_slide(document, "257.29")
    yield document, svg
    assert document.to_bytes() == original, "a refused SVG changed the document"
    assert not document.history.can_undo()


def _mutate(svg: FullStateSvg, change) -> None:
    data = svg.json("257.29", "chart-data")
    change(data)
    svg.set_json("257.29", "chart-data", data)


@pytest.mark.parametrize("change", [
    lambda d: d["series"][0]["values"].__setitem__(0, "12"),
    lambda d: d["series"][0]["values"].__setitem__(0, True),
    lambda d: d["series"][0]["values"].append(1),             # one value too many
    lambda d: d["series"].append({"name": "x"}),               # no values
    lambda d: d.__setitem__("legend", "yes"),
    lambda d: d.__setitem__("types", ["pie3D", "teapot"]),
    lambda d: d["axis_titles"].__setitem__("depth", "z"),
    lambda d: d.pop("format"),
    lambda d: d.__setitem__("types", ["bar"]),                 # read-only
    lambda d: d.__setitem__("format", "0.0%"),                 # read-only
    lambda d: d.__setitem__("series", []),
], ids=["string", "bool", "length", "shape", "legend", "type", "axis", "keys", "types-ro",
        "format-ro", "no-series"])
def test_chart_data_is_validated(chart_deck, change):
    document, svg = chart_deck
    _mutate(svg, change)
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg))


def test_non_finite_numbers_are_refused(chart_deck):
    document, svg = chart_deck
    raw = svg.get("257.29", "chart-data").replace("11.7", "NaN", 1)
    svg.set("257.29", "chart-data", raw)
    with pytest.raises(FullStateError, match="finite"):
        document.apply_svg(str(svg))
    raw = raw.replace("NaN", "Infinity", 1)
    svg.set("257.29", "chart-data", raw)
    with pytest.raises(FullStateError, match="finite"):
        document.apply_svg(str(svg))
    assert not math.isnan(document.shape("257.29").chart.series[0].values[0])


def test_chart_size_is_bounded(chart_deck):
    document, svg = chart_deck
    _mutate(svg, lambda d: d["series"][0]["values"].__setitem__(0, 1))
    with pytest.raises(FullStateError, match="more than"):
        document.apply_svg(str(svg), limits=Limits(chart_points=4))


@pytest.fixture
def diagram_deck():
    document = Document.open(SMARTART.read_bytes())
    original = document.to_bytes()
    yield document, FullStateSvg(document.slides[1].render_svg(full_state=True))
    assert document.to_bytes() == original, "a refused SVG changed the document"
    assert not document.history.can_undo()


@pytest.mark.parametrize("change", [
    lambda m: m["nodes"][0].__setitem__("id", "{00000000-0000-0000-0000-000000000000}"),
    lambda m: m["nodes"].reverse(),                             # reordered
    lambda m: m["nodes"][1].__setitem__("lvl", 0),              # moved up a level
    lambda m: m["nodes"][0].__setitem__("lvl", 1),              # level 1 first
    lambda m: m.__setitem__("layout", "urn:other"),             # read-only
    lambda m: m["nodes"].__delitem__(0),                        # takes kept children with it
    lambda m: m.__setitem__("nodes", []),
    lambda m: m["nodes"][0].__setitem__("t", 5),
    lambda m: m["nodes"][0].__setitem__("id", "<script>"),
], ids=["unknown-id", "reorder", "level", "first-level", "layout", "orphans", "empty",
        "text-type", "bad-id"])
def test_diagram_nodes_are_validated(diagram_deck, change):
    document, svg = diagram_deck
    model = svg.json("257.3", "diagram-nodes")
    change(model)
    svg.set_json("257.3", "diagram-nodes", model)
    with pytest.raises(FullStateError):
        document.apply_svg(str(svg))


def test_diagram_size_is_bounded(diagram_deck):
    document, svg = diagram_deck
    model = svg.json("257.3", "diagram-nodes")
    model["nodes"][0]["t"] = "changed"
    svg.set_json("257.3", "diagram-nodes", model)
    with pytest.raises(FullStateError, match="more than"):
        document.apply_svg(str(svg), limits=Limits(diagram_nodes=3))


def test_the_diagram_part_is_read_only(diagram_deck):
    document, svg = diagram_deck
    svg.set("257.3", "diagram", "ppt/diagrams/data1.xml")
    with pytest.raises(FullStateError, match="read-only"):
        document.apply_svg(str(svg))
