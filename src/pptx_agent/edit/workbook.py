"""The workbook behind a chart: now :mod:`ooxml_edit.charts.workbook`, re-exported here.

A chart's embedded ``.xlsx`` is the same SpreadsheetML in a deck and in a Word document, so
the cell-by-cell editing moved to ooxml-edit's charts subpackage, with its history.  This
module keeps the old import path working.
"""

from ooxml_edit.charts.workbook import *  # noqa: F401,F403  (re-exported)
from ooxml_edit.charts.workbook import (  # noqa: F401  (re-exported)
    MAX_COLUMN,
    MAX_ROW,
    SML_NS,
    Area,
    Cell,
    SheetRange,
    Table,
    Workbook,
    Worksheet,
    column_index,
    column_letters,
    number_text,
    parse_formula,
)
