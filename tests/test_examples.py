"""``examples/restate_figures.py``: the end-to-end pilot's task, run on the pilot's deck.

Its figures are the answer key's.  With ``PPTX_AGENT_PILOT_OUTPUT`` naming the deck the
pilot's own script produced, every table cell and chart value is also compared with it,
through the public API (that deck is not part of the repository).
"""

from __future__ import annotations

import importlib.util
import os
from pathlib import Path

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document
from test_chart import check_deck_charts

REPORT = FIXTURE_DIR / "real-financial-report.pptx"
EXAMPLE = Path(__file__).resolve().parent.parent / "examples" / "restate_figures.py"


def _example():
    spec = importlib.util.spec_from_file_location("restate_figures", EXAMPLE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def restated(tmp_path_factory) -> Document:
    target = tmp_path_factory.mktemp("restated") / "restated.pptx"
    _example().main(str(REPORT), str(target))
    return Document.open(str(target))


def _text(deck: Document, address: str) -> str:
    return deck.resolve(address).text


def test_the_answer_keys_figures(restated):
    pl = restated.shape("257.3#5").table
    rows = {label: [pl.cell_by_label(label, c).text for c in ("当期実績", "増減額", "増減率")]
            for label in ("売上高", "営業利益", "当期純利益")}
    assert rows == {"売上高": ["4,310億円", "+418億", "+10.7%"],
                    "営業利益": ["520億円", "+82億", "+18.7%"],
                    "当期純利益": ["333億円", "+57億", "+20.7%"]}
    assert [_text(restated, a) for a in ("257.12", "257.16", "257.20")] == \
        ["77.0%", "79.4%", "78.5%"]
    assert [_text(restated, a) for a in ("256.13", "256.15", "256.17", "256.19", "256.21",
                                          "256.23", "256.29", "256.30")] == \
        ["4,310", "▲ 10.7%", "520", "▲ 18.7%", "333", "▲ 20.7%", "12.1%", "+0.8pt YoY"]
    segments = restated.shape("258.4#3").table
    assert [segments.cell_by_label("その他・調整", c).text
            for c in ("売上高", "構成比", "営業利益", "利益率", "YoY")] == \
        ["369億円", "8.6%", "35億円", "9.5%", "+3.1%"]  # its YoY cannot be derived: kept
    assert [segments.cell_by_label("合計", c).text
            for c in ("売上高", "構成比", "営業利益", "利益率", "YoY")] == \
        ["4,310億円", "100%", "520億円", "12.1%", "+18.7%"]
    shares = [segments.cell(row, 2).text for row in range(1, 5)]
    assert shares == ["42.7%", "29.8%", "18.9%", "8.6%"]
    charts = {sid: [s.values for s in restated.shape(sid).chart.series]
              for sid in ("257.25", "257.29", "258.9", "258.13")}
    assert charts == {"257.25": [[3980, 4120, 4310], [465, 488, 520]],
                      "257.29": [[11.7, 11.8, 12.1], [42, 42.5, 42.7], [7.4, 7.5, 7.7]],
                      "258.9": [[1599, 1185, 663, 334], [1842, 1285, 814, 369]],
                      "258.13": [[42.7, 29.8, 18.9, 8.6]]}


def test_only_what_the_task_names_changed(restated):
    before = {b.address: b.text for b in Document.open(str(REPORT)).outline_blocks()}
    after = {b.address: b.text for b in restated.outline_blocks()}
    assert before.keys() == after.keys()
    changed = {a for a in before if before[a] != after[a]}
    assert len(changed) == 31  # the 31 text figures of the pilot's self-check
    assert not any(a.startswith("259.") for a in changed)  # the forecast slide is untouched
    original = Document.open(str(REPORT))
    radar = [s.values for s in original.shape("259.11").chart.series]
    assert [s.values for s in restated.shape("259.11").chart.series] == radar


def test_the_charts_agree_with_their_workbooks(restated):
    check_deck_charts(restated.to_bytes(), REPORT.read_bytes())


@pytest.mark.skipif(not os.environ.get("PPTX_AGENT_PILOT_OUTPUT"),
                    reason="set PPTX_AGENT_PILOT_OUTPUT to the pilot's restated deck")
def test_the_same_cells_and_chart_values_as_the_pilots_output(restated):
    pilot = Document.open(os.environ["PPTX_AGENT_PILOT_OUTPUT"])
    assert [(b.address, b.text) for b in restated.outline_blocks()] == \
        [(b.address, b.text) for b in pilot.outline_blocks()]
    for slide in pilot.slides:
        for shape in slide.shapes:
            if shape.kind != "group" and shape.has_chart:
                ours = restated.shape(shape.id).chart
                assert ours.categories == shape.chart.categories
                assert [(s.name, s.values) for s in ours.series] == \
                    [(s.name, s.values) for s in shape.chart.series]
