"""Table tools: ``ppt_add_table`` and ``ppt_edit_table`` (cell fill and borders too; a
cell's text is formatted with ``ppt_format_text`` on its address).

Rows and columns are numbered from 0, as the library and cell addresses
(``256.7/cell1,2``) number them; a cell may also be named by its row and column labels.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, integer, number, obj, string, tool

from .common import (KIND, box_emu, emu, library_errors, shape as shape_at,
                     slide as slide_at)
from .shapes import OBJECTS, _box

_KEY = string("Retry key: a repeat with the same key makes nothing new.", optional=True)


def _table(deck, target: str):
    frame = shape_at(deck, target)
    table = frame.table if frame.kind == "graphic_frame" else None
    if table is None:
        raise ToolError("invalid_arguments", f"{target} is a {frame.kind}, not a table",
                        field="target")
    return frame, table


def table_json(frame, table) -> dict[str, Any]:
    """The table as the model reads it: size, labels, and its cells' text (up to 400)."""
    rows = []
    for row in range(min(table.rows, 40)):
        rows.append([table.cell(row, column).text for column in range(min(table.columns, 10))])
    return {"address": frame.id, "rows": table.rows, "columns": table.columns,
            "column_labels": table.column_labels(), "row_labels": table.row_labels(),
            "cells": rows}


@tool("ppt_add_table",
      "Add a table with its data, styled as PowerPoint inserts one: a header row, banded "
      "rows.",
      {"doc": string("Document id."),
       "slide": string("Slide, s:256 or $ref."),
       "box": _box("Position and size."),
       "rows": integer("1-200.", minimum=1, maximum=200),
       "columns": integer("1-50.", minimum=1, maximum=50),
       "data": array(array(string()), "Cell texts, row by row; may be shorter than the "
                     "table.", optional=True),
       "ref": string("Ref name for the table.", optional=True),
       "key": _KEY},
      kind=KIND, group=OBJECTS, mutates=True, refs=("slide",))
@library_errors
def ppt_add_table(call, doc, slide, box, rows, columns, data=None, ref=None, key=None):
    host = slide_at(call.document, slide)
    if rows * columns > 5000:
        raise ToolError("limit", "at most 5,000 cells per table", field="rows")
    x, y, w, h = box_emu(box)
    frame = host.add_table(rows, columns, x, y, w, h)
    table = frame.table
    for row, values in enumerate((data or [])[:rows]):
        for column, text in enumerate(values[:columns]):
            if text:
                table.cell(row, column).text = text
    if ref:
        call.define_ref(ref, frame.id)
    call.touch(host.slide_id)
    return Result(summary=f"Added a {rows}x{columns} table {frame.id}", created=[frame.id],
                  data=table_json(frame, call.document.shape(frame.id).table))


_CELL = obj({"row": integer("From 0.", minimum=0, maximum=10000, optional=True),
             "col": integer("From 0.", minimum=0, maximum=10000, optional=True),
             "row_label": string("Or the row's label (first column).", optional=True),
             "col_label": string("Or the column's label (header row).", optional=True),
             "text": string("Keeps the cell's formatting.")})


_SIDES = ["top", "bottom", "left", "right", "all"]


@tool("ppt_edit_table",
      "Change a table: cell text by position or label, insert or delete rows and columns, "
      "merge or split cells, column widths, row heights, cell fill and borders. Cell text "
      "formatting: ppt_format_text on cells. Returns the table after.",
      {"doc": string("Document id."),
       "target": string("The table, e.g. 256.7 or $ref."),
       "action": string("What to do.", enum=["set_cells", "insert_row", "delete_row",
                                             "insert_column", "delete_column", "merge",
                                             "split", "set_widths", "set_heights",
                                             "format_cells"]),
       "cells": array(_CELL, "set_cells: the cells.", optional=True),
       "index": integer("insert_*: new row/column's position, default last; delete_*: "
                        "which.", minimum=0, maximum=10000, optional=True),
       "like": integer("insert_*: copy this row/column's formatting.", minimum=0,
                       maximum=10000, optional=True),
       "span": obj({"row": integer(minimum=0, maximum=10000),
                    "col": integer(minimum=0, maximum=10000),
                    "to_row": integer("merge. Default row.", minimum=0, maximum=10000,
                                      optional=True),
                    "to_col": integer("merge. Default col.", minimum=0, maximum=10000,
                                      optional=True)},
                   "merge: the cells to join; split: a merged cell.", optional=True),
       "sizes": array(number(minimum=0, maximum=10000),
                      "set_widths/set_heights: one per column/row.", optional=True),
       "rows": array(integer(minimum=0, maximum=10000),
                     "format_cells: rows from 0. Default all.", optional=True),
       "columns": array(integer(minimum=0, maximum=10000),
                        "format_cells: columns from 0. Default all.", optional=True),
       "fill": string("format_cells: colour or none.", optional=True),
       "borders": array(obj({"side": string(enum=_SIDES),
                             "width": number("0 removes.", minimum=0, maximum=100),
                             "color": string(optional=True)}), "format_cells: borders.",
                        optional=True)},
      kind=KIND, group=OBJECTS, mutates=True, refs=("target",))
@library_errors
def ppt_edit_table(call, doc, target, action, cells=None, index=None, like=None, span=None,
                   sizes=None, rows=None, columns=None, fill=None, borders=None):
    frame, table = _table(call.document, target)

    def need(name, value):
        if value is None:
            raise ToolError("invalid_arguments", f"{action} needs {name}", field=name)
        return value

    changed = [frame.id]
    if action == "set_cells":
        for i, cell in enumerate(need("cells", cells)):
            row, col = cell.get("row"), cell.get("col")
            if "row_label" in cell or "col_label" in cell:
                if "row_label" not in cell or "col_label" not in cell:
                    raise ToolError("invalid_arguments", "give both row_label and col_label",
                                    field=f"cells[{i}].row_label")
                found = table.cell_by_label(cell["row_label"], cell["col_label"])
            elif row is None or col is None:
                raise ToolError("invalid_arguments", "give row and col, or the labels",
                                field=f"cells[{i}].row")
            else:
                if not (row < table.rows and col < table.columns):
                    raise ToolError("not_found", f"no cell {row},{col}; the table is "
                                    f"{table.rows}x{table.columns}", field=f"cells[{i}]")
                found = table.cell(row, col)
            found.text = cell["text"]
            changed.append(found.address)
    elif action == "insert_row":
        table.insert_row(index, like=like)
    elif action == "delete_row":
        table.delete_row(need("index", index))
    elif action == "insert_column":
        table.insert_column(index, like=like)
    elif action == "delete_column":
        table.delete_column(need("index", index))
    elif action == "merge":
        area = need("span", span)
        table.merge(area["row"], area["col"], area.get("to_row", area["row"]),
                    area.get("to_col", area["col"]))
    elif action == "split":
        area = need("span", span)
        table.split(area["row"], area["col"])
    elif action == "set_widths":
        values = need("sizes", sizes)
        if len(values) != table.columns:
            raise ToolError("invalid_arguments", f"{len(values)} widths for {table.columns} "
                            "columns", field="sizes")
        for column, value in enumerate(values):
            table.set_column_width(column, emu(value))
    elif action == "set_heights":
        values = need("sizes", sizes)
        if len(values) != table.rows:
            raise ToolError("invalid_arguments", f"{len(values)} heights for {table.rows} rows",
                            field="sizes")
        for row, value in enumerate(values):
            table.set_row_height(row, emu(value))
    elif action == "format_cells":
        if fill is None and not borders:
            raise ToolError("invalid_arguments", "format_cells needs fill or borders",
                            field="fill", valid_options=["fill", "borders"])
        _format_cells(table, rows, columns, fill, borders)
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id)
    return Result(summary=f"{action} on table {frame.id}", changed=changed,
                  data=table_json(fresh, fresh.table))


def _format_cells(table, rows, columns, fill, borders) -> int:
    """Fill and borders on the cells of ``rows`` x ``columns`` (every one by default)."""
    row_list = rows if rows is not None else list(range(table.rows))
    column_list = columns if columns is not None else list(range(table.columns))
    for value in row_list:
        if value >= table.rows:
            raise ToolError("not_found", f"no row {value}", field="rows")
    for value in column_list:
        if value >= table.columns:
            raise ToolError("not_found", f"no column {value}", field="columns")
    count = 0
    for row in row_list:
        for column in column_list:
            cell = table.cell(row, column)
            if fill is not None:
                cell.fill = fill
            for border in borders or ():
                sides = ["top", "bottom", "left", "right"] if border["side"] == "all" \
                    else [border["side"]]
                for side in sides:
                    cell.set_border(side, width=emu(border["width"]),
                                    color=border.get("color"))
            count += 1
    return count


TOOLS = [ppt_add_table, ppt_edit_table]
