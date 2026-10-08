"""Drafted tables and pictures go below the title, and a table's columns follow its width
(the trial's finding 9: P1 drafted a table over the title of its template's TITLE_ONLY
layout, whose title sits lower than the master's body begins)."""

from __future__ import annotations

import warnings

from conftest import FIXTURE_DIR
from pptx_agent import Document

TEMPLATE = FIXTURE_DIR / "generated" / "trial" / "company-template.potx"
TABLE_SLIDE = """\
<!-- layout: TITLE_ONLY -->
# Costs and benefits of Option A (EUR thousand)

| Item | Year 1 | Year 2 | Year 3 |
| --- | --- | --- | --- |
| Licences | 380 | 380 | 380 |
| Implementation | 210 | 40 | 0 |
| Savings | -150 | -900 | -1,250 |
| Net | 440 | -480 | -870 |

Notes:

Option A pays back during year 2.
"""


def _template_deck() -> Document:
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        return Document.new(template=TEMPLATE)


def test_the_trial_table_is_drafted_below_the_title():
    deck = _template_deck()
    (slide,) = deck.insert_outline(TABLE_SLIDE)
    title, table = slide.shapes
    assert title.placeholder.type == "title" and table.has_table
    assert table.top >= title.top + title.height
    assert table.top + table.height <= deck.slide_size[1]
    assert slide.layout.name == "TITLE_ONLY"
    assert deck.overflows() == []
    assert slide.notes == "Option A pays back during year 2."


def test_free_content_keeps_the_master_area_when_the_title_is_above_it():
    deck = Document.new()
    (slide,) = deck.insert_outline("<!-- layout: Title Only -->\n# T\n\n| a | b |\n| - | - |\n"
                                   "| 1 | 2 |\n")
    title, table = slide.shapes
    body = deck.layout("Title and Content").placeholders[1]
    assert table.top == body.bounds[1]                  # unchanged: the title ends above it
    assert deck.overflows() == []


def test_a_table_width_rescales_its_columns():
    deck = _template_deck()
    (slide,) = deck.insert_outline(TABLE_SLIDE)
    frame = slide.shapes[1]
    table = frame.table
    table.set_column_width(0, 3000000)
    before = table.column_widths
    frame.width = 6000000
    after = table.column_widths
    assert sum(after) == 6000000 == frame.width
    assert [round(w / after[1], 2) for w in after] == [round(w / before[1], 2) for w in before]
    rows = table.row_heights
    frame.height = 2 * sum(rows)
    assert sum(table.row_heights) == frame.height
    assert deck.validate() == []
    deck.undo()
    deck.undo()
    assert slide.shapes[1].table.column_widths == before
