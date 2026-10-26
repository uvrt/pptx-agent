"""S10 (review): ``edit_chart`` and ``edit_smartart`` behave the same in a deck and a Word
document.  ooxml-edit's table (``ooxml_edit.charts.cases``) is run here through this
package's tools, and through docx-agent's in its ``tests/test_chart_cases.py``: the same
calls, the same results, but for the chart's name in a result (here ``address``)."""

from __future__ import annotations

from pathlib import Path

import pytest

from ooxml_edit.charts.cases import CHART, CHART_CASES, SMARTART_CASES, problems
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

FIXTURES = Path(__file__).parent / "fixtures"
BOX = {"x": 40, "y": 120, "w": 640, "h": 360}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def _deck() -> bytes:
    deck = Document.new()
    deck.add_slide("Title Only").title = "Chart"
    return deck.to_bytes()


@pytest.mark.parametrize("case", CHART_CASES, ids=[case["name"] for case in CHART_CASES])
def test_a_chart_case(toolbox, case):
    session = toolbox.session()
    session.open(_deck(), "chart.pptx")

    def call(**arguments):
        return toolbox.dispatch(session, "edit_chart", {"doc": "d1", **arguments})

    made = call(target="s:256", action="add", box=BOX, ref="c", **CHART)
    assert made.ok, made.to_json()
    for arguments in case["before"]:
        assert call(target="$c", **arguments).ok
    assert problems(case, call(target="$c", **case["call"]), key="address") == []


@pytest.mark.parametrize("case", SMARTART_CASES, ids=[case["name"] for case in SMARTART_CASES])
def test_a_smartart_case(toolbox, case):
    session = toolbox.session()
    session.open((FIXTURES / "powerpoint-smartart.pptx").read_bytes(), "smartart.pptx")
    result = toolbox.dispatch(session, "edit_smartart",
                              {"doc": "d1", "target": "256.3", **case["call"]})
    assert problems(case, result) == []
