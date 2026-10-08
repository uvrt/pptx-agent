"""E4's acceptance edits: every chart and SmartArt diagram in a deck, changed the ways an
agent would change them -- through the API and through the full-state SVG.

What each edit leaves for PowerPoint's PDF to show is returned, so the oracle test can look
for it: new labels and names in the page's text, a value large enough that the value axis
has to reach it, and node text in a diagram that PowerPoint lays out again.
"""

from __future__ import annotations

import warnings

from pptx_agent import Document
from svgedit import FullStateSvg

#: Large enough that no fixture's axis reaches it unedited, so a label at least this big on
#: the page means PowerPoint drew the new value.
BIG = 87654


def acceptance_edits(document: Document) -> dict:
    expected = {"pages": len(document.slides), "texts": {}, "absent": {}, "big": set()}

    def expect(page: int, text: str) -> None:
        expected["texts"].setdefault(page, []).append(text)

    number = 0
    for page, slide in enumerate(list(document.slides)):
        for shape in list(slide.shapes):
            if shape.has_chart:
                number += 1
                chart = shape.chart
                with warnings.catch_warnings():
                    warnings.simplefilter("error")  # every fixture chart has its workbook
                    _chart_edits(document, slide, shape.id, number)
                pie_like = chart.chart_type in {"pie", "pie3D", "doughnut", "ofPie"}
                expect(page, f"E4T{number}")
                expect(page, f"E4C{number}")
                expect(page, f"E4N{number}")
                if not pie_like:  # a pie's legend lists categories, not series
                    expect(page, f"E4S{number}")
                    expect(page, f"E4A{number}")
                    expected["big"].add(page)
            elif shape.has_diagram:
                diagram = shape.diagram
                if any(node.level for node in diagram.nodes):
                    _diagram_edits_through_svg(document, slide, shape.id)
                    expect(page, "E4-CHILD")
                    expect(page, "E4-NEWCHILD")
                    expected["absent"].setdefault(page, []).append("Stale caches")
                else:
                    diagram.set_text(1, "E4-BLOCK")
                    diagram.add_node("E4-ADDED")
                    expect(page, "E4-BLOCK")
                    expect(page, "E4-ADDED")
    return expected


def _chart_edits(document: Document, slide, shape: str, number: int) -> None:
    chart = document.shape(shape).chart
    last = chart.point_count - 1
    chart.series[-1].set_value(last, BIG)
    chart.series[0].name = f"E4S{number}"
    chart.set_category(0, f"E4C{number}")
    chart.add_series(f"E4A{number}", [1] * chart.point_count)
    chart.add_category(f"E4N{number}", [2] * len(chart.series))
    if chart.point_count > 3:
        chart.remove_category(1)
    # The title through the full-state SVG, so that path reaches PowerPoint too.
    svg = FullStateSvg(document.slide(slide.slide_id).render_svg(full_state=True))
    data = svg.json(shape, "chart-data")
    data["title"] = f"E4T{number}"
    svg.set_json(shape, "chart-data", data)
    document.apply_svg(str(svg))


def _diagram_edits_through_svg(document: Document, slide, shape: str) -> None:
    svg = FullStateSvg(document.slide(slide.slide_id).render_svg(full_state=True))
    model = svg.json(shape, "diagram-nodes")
    nodes = model["nodes"]
    nodes[1]["t"] = "E4-CHILD"
    model["nodes"] = [node for node in nodes if node["t"] != "Stale caches"]
    position = next(i for i, node in enumerate(model["nodes"]) if node["t"] == "Risks")
    model["nodes"].insert(position + 1, {"lvl": 1, "t": "E4-NEWCHILD"})
    svg.set_json(shape, "diagram-nodes", model)
    document.apply_svg(str(svg))
