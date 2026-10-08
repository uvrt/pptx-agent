"""Reading a chart's embedded workbook, so "Edit Data shows the same numbers" can be
checked rather than trusted (the trial's finding 13)."""

from __future__ import annotations

import warnings

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import ChartDataWarning, Document

REPORT = FIXTURE_DIR / "real-financial-report.pptx"


def _charts(deck):
    return [shape.chart for slide in deck.slides for shape in slide.shapes if shape.has_chart]


def test_the_workbook_holds_what_the_cache_holds():
    deck = Document.open(REPORT)
    charts = _charts(deck)
    assert len(charts) == 5
    for chart in charts:
        book = chart.workbook_values()
        assert book["sheet"] == "Sheet1"
        assert book["categories"] == chart.categories
        assert book["series"] == chart.data["series"]


def test_reading_changes_nothing():
    deck = Document.open(REPORT)
    before = deck.to_bytes()
    for chart in _charts(deck):
        chart.workbook_values()
    assert deck.to_bytes() == before


def test_edits_reach_the_workbook():
    deck = Document.open(REPORT)
    chart = deck.shape("257.25").chart
    chart.series["売上高（億円）"].set_value("Q3", 4310)
    chart.add_category("Q4", {"売上高（億円）": 4400})
    chart.add_series("純利益（億円）", [300, 315, 333, None])
    book = chart.workbook_values()
    assert book["categories"] == ["Q1", "Q2", "Q3", "Q4"]
    assert book["series"] == chart.data["series"]
    assert book["series"][0]["values"] == [3980, 4120, 4310, 4400]
    assert book["series"][2] == {"name": "純利益（億円）", "values": [300, 315, 333, None]}


def test_a_chart_without_a_workbook_has_none():
    deck = Document.open(REPORT)
    chart = deck.shape("257.25").chart
    host = chart._resolve()
    from ooxml_edit.charts.host import chart_part

    part = chart_part(host)
    for rel in list(deck.package.relationships(part).values()):
        if rel.type.endswith("/package"):
            deck.package.remove_relationship(part, rel.id)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", ChartDataWarning)
        assert chart.workbook_part is None
        assert chart.workbook_values() is None
