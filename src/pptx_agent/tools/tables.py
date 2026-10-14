"""Table tools: ``ppt_add_table`` and ``ppt_edit_table`` (cell fill and borders too; a
cell's text is formatted with ``ppt_format_text`` on its address).

Rows and columns are numbered from 0, as the library and cell addresses
(``256.7/cell1,2``) number them; a cell may also be named by its row and column labels.

**Rows from a source.**  Both tools may take a table's text from a source instead of
literal strings (``from``: a dataset ``d:<id>``, or a table of an open document,
``t:<name>``; :mod:`ooxml_edit.tools.sources`), as ``edit_chart`` takes a chart's.
``from_columns`` picks the source's columns by name, in the order placed (default: every
column, in the source's order); every row is placed, in the source's order.  Above the
rows goes a header row: the column names as the source has them, unless ``header`` gives
its texts, or ``[]`` for none.  A value is written as the chart labels write a category
(:func:`ooxml_edit.tools.sources.label`): text as it is -- so a date as the source gives
it, ISO 8601 text from a workbook -- a number as the General format writes it with the
session locale's decimal separator, a logical value ``TRUE`` or ``FALSE``, a blank as an
empty cell.  Nothing is cut: a source too big for a table (200 rows with the header, 50
columns, 5,000 cells), or for the table it is written into, is refused, saying by how much.

The result's ``data.from`` says where each placed value came from, as ``edit_chart``'s
does: the source, its document and version, the table row of the header (``header_row``,
``None`` without one) and of the first source row (``first_row``; source row ``n``, from
1, is table row ``first_row + n - 1``), and per placed column its table column (``col``),
source column, source rows (from 1, as runs) and their sources, merged.

**The record in the deck.**  What ``from`` says is also kept in the deck, in the table's
graphic frame (:mod:`pptx_agent.edit.provenance`: a ``UVRT_PROVENANCE`` tag in the frame's
tags part), through :meth:`Call.keep_provenance <ooxml_edit.tools.dispatch.Call.keep_provenance>`,
in the same undo step; :func:`read_provenance` (``Session.provenance``) and
``Document.provenance()`` read it back.  The record is kept true or dropped, never left
naming a cell its value did not come to:

- ``ppt_add_table`` from a source writes it; ``set_cells`` from a source replaces it;
- ``set_cells`` without ``from``, ``ppt_set_text`` and ``replace_text`` on a cell of the
  block (its header row to its last source row), and ``merge`` over one, drop that column's
  entry; with no entry left, the record goes;
- ``insert_row`` above the block (or at its first row) moves it down, ``delete_row`` above
  it moves it up; either inside the block (the header too) drops the record; below it,
  nothing changes;
- ``insert_column`` moves the entries at or right of it right; ``delete_column`` drops its
  entry and moves those right of it left.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (Result, ToolError, array, boolean, integer, names_as, number, obj,
                              string, tool)
from ooxml_edit.tools import sources

from ..edit import provenance as _provenance
from ..edit.tablestyles import FLAGS

from .common import (KIND, box_emu, emu, library_errors, shape as shape_at,
                     host as slide_at)
from .shapes import OBJECTS, _box

_KEY = string("Retry key: a repeat with the same key makes nothing new.", optional=True)
_STYLE = string("Table style name or id; describe lists the deck's own.", optional=True)
#: ``columns`` given as names, as ``edit_chart`` takes them: ``from_columns`` here.
_COLUMN_NAMES = {"columns": names_as("columns", "from_columns")}
_FROM_COLUMNS = array(string(), "from's columns, in order. Default all.", optional=True)
_HEADER = array(string(), "from's header texts; [] none. Default column names.",
                optional=True)

#: The most a table takes: rows (the header included), columns and cells.
MAX_ROWS, MAX_COLUMNS, MAX_CELLS = 200, 50, 5000


def _style_id(deck, style):
    """``style``'s id, or ``None`` for none; ``not_found`` names the deck's own styles."""
    if style is None:
        return None
    from ..edit.tablestyles import resolve

    try:
        return resolve(deck, style)
    except ValueError as exc:
        raise ToolError("not_found", str(exc), field="style",
                        valid_options=[s["name"] for s in deck.table_styles
                                       if not s["builtin"]][:50] or None) from None


def _table(deck, target: str):
    frame = shape_at(deck, target)
    table = frame.table if frame.kind == "graphic_frame" else None
    if table is None:
        raise ToolError("invalid_arguments", f"{target} is a {frame.kind}, not a table",
                        field="target")
    return frame, table


def table_json(frame, table) -> dict[str, Any]:
    """The table as the model reads it: size, style and the flags that are on, labels, and
    its cells' text (up to 400)."""
    rows = []
    for row in range(min(table.rows, 40)):
        rows.append([table.cell(row, column).text for column in range(min(table.columns, 10))])
    style = table.style
    return {"address": frame.id, "rows": table.rows, "columns": table.columns,
            "style": (style["name"] or style["id"]) if style else None,
            "flags": [flag for flag, on in table.flags.items() if on],
            "column_labels": table.column_labels(), "row_labels": table.row_labels(),
            "cells": rows}


@tool("ppt_add_table",
      "Add a table with its data, styled as PowerPoint inserts one: a header row, banded "
      "rows, the default style or one named.",
      {"doc": string("Document id."),
       "slide": string("Slide, s:256 or $ref."),
       "box": _box("Position and size."),
       "rows": integer("1-200. Default fits from.", minimum=1, maximum=200,
                       optional=True),
       "columns": integer("1-50. Default fits from.", minimum=1, maximum=50, optional=True),
       "data": array(array(string()), "Cell texts, row by row; may be shorter than the "
                     "table.", optional=True),
       "from": string("Or rows of d:<id> or t:<table>.", optional=True),
       "from_columns": _FROM_COLUMNS,
       "header": _HEADER,
       "style": _STYLE,
       "ref": string("Ref name for the table.", optional=True),
       "key": _KEY},
      kind=KIND, group=OBJECTS, mutates=True, refs=("slide",), sources=("from",),
      aliases=_COLUMN_NAMES)
@library_errors
def ppt_add_table(call, doc, slide, box, rows=None, columns=None, data=None, style=None,
                  ref=None, key=None, header=None, from_columns=None, **given):
    source = given.get("from")
    placed = None
    if source is None:
        _no_source(from_columns, header)
        for name, value in (("rows", rows), ("columns", columns)):
            if value is None:
                raise ToolError("invalid_arguments", f"give {name}, or from", field=name)
    else:
        if data is not None:
            raise ToolError("invalid_arguments", "from replaces data: give one", field="data")
        placed = _placed(call, source, from_columns, header)
        need_rows, need_columns = placed.size
        _fits(placed, MAX_ROWS, MAX_COLUMNS)
        for name, value, needed in (("rows", rows, need_rows),
                                    ("columns", columns, need_columns)):
            if value is not None and value < needed:
                raise ToolError("invalid_arguments", f"{name} is {value}, but from needs "
                                f"{needed}: leave {name} out to fit it", field=name)
        rows = rows if rows is not None else need_rows
        columns = columns if columns is not None else need_columns
        data = placed.texts()
    if rows * columns > MAX_CELLS:
        raise ToolError("limit", f"{rows}x{columns} is {rows * columns:,} cells; a table takes "
                        f"at most {MAX_CELLS:,}" + (": filter the source or split it across "
                                                    "tables" if placed else ""),
                        field="from" if placed else "rows")
    host = slide_at(call.document, slide)
    x, y, w, h = box_emu(box)
    frame = host.add_table(rows, columns, x, y, w, h, style=_style_id(call.document, style))
    table = frame.table
    for row, values in enumerate((data or [])[:rows]):
        for column, text in enumerate(values[:columns]):
            if text:
                table.cell(row, column).text = text
    if ref:
        call.define_ref(ref, frame.id)
    call.touch(host.slide_id)
    result = table_json(frame, call.document.shape(frame.id).table)
    if placed is not None:
        result["from"] = placed.provenance(0)
        call.keep_provenance(frame.id, placed.rows, result["from"])
    return Result(summary=f"Added a {rows}x{columns} table {frame.id}"
                  + (f" from {placed.rows.address}" if placed else ""), created=[frame.id],
                  data=result)


class _Placed:
    """A source's rows as table text: the header (if any), then a row per source row."""

    def __init__(self, rows, at, header, locale):
        self.rows, self.at, self.header, self.locale = rows, at, header, locale
        #: The block's size: the header row and a row per source row; a column per pick.
        self.size = ((1 if header else 0) + len(rows.dataset.rows), len(at))

    def texts(self) -> list[list[str]]:
        """The block's text, row by row (the header first)."""
        return ([self.header] if self.header else []) + [
            [_text(self.rows, row, column, self.locale) for column in self.at]
            for row in range(len(self.rows.dataset.rows))]

    def provenance(self, top: int) -> dict[str, Any]:
        """``data.from`` for the block written with its first row at table row ``top``."""
        order = list(range(len(self.rows.dataset.rows)))
        first = top + (1 if self.header else 0)
        return {**self.rows.describe(), "header_row": top if self.header else None,
                "first_row": first,
                "columns": [{"col": col, **self.rows.provenance(column, order)}
                            for col, column in enumerate(self.at)]}


def _no_source(from_columns, header):
    for name, value in (("from_columns", from_columns), ("header", header)):
        if value is not None:
            raise ToolError("invalid_arguments", f"{name} is for from; give from too",
                            field=name)


def _placed(call, source, from_columns, header) -> _Placed:
    """The source's rows and the columns picked, as text; refusals name the cell."""
    rows = call.source(source, field="from")
    names = from_columns if from_columns is not None else list(rows.dataset.columns)
    if not names:
        raise ToolError("invalid_arguments", "from_columns names no column; leave it out "
                        "for every column", field="from_columns")
    at = [rows.column(name, field=f"from_columns[{i}]") for i, name in enumerate(names)]
    if len(set(at)) != len(at):
        raise ToolError("invalid_arguments", "a column is named twice", field="from_columns")
    if header is None:
        header = [rows.dataset.columns[column] for column in at]
    elif header and len(header) != len(at):
        raise ToolError("invalid_arguments", f"{len(header)} header texts for {len(at)} "
                        "columns; [] for no header row", field="header")
    return _Placed(rows, at, list(header), call.locale)


def _text(rows, row: int, column: int, locale) -> str:
    """A source value as cell text, as a chart's category label writes it; a blank is
    empty; anything else (not a finite number) is refused, naming the cell."""
    value = rows.dataset.rows[row][column]
    if value is None or value == "":
        return ""
    text = sources.label(value, locale)
    if text is None:
        raise ToolError("invalid_arguments", f"{rows.where(row, column)} is "
                        f"{sources.what(value)}: a cell takes text, a number, a logical value "
                        "or a blank", field="from")
    return text


def _fits(placed: _Placed, max_rows: int, max_columns: int, room: str | None = None) -> None:
    """Refuse a block bigger than ``max_rows`` x ``max_columns``, saying by how much: a
    table's limits, or (``room``) the room an existing table has."""
    address = placed.rows.address
    need_rows, need_columns = placed.size
    code = "refused" if room else "limit"
    if need_rows > max_rows:
        header = " with the header" if placed.header else ""
        fix = (f"{room}, or insert_row first" if room
               else f"a table takes at most {max_rows}: filter the source or split it "
                    "across tables")
        raise ToolError(code, f"{address} gives {need_rows} rows{header}; {fix} "
                        "(nothing is cut)", field="from")
    if need_columns > max_columns:
        fix = (f"the table has {max_columns}" if room
               else f"a table takes at most {max_columns}")
        raise ToolError(code, f"{address}: {need_columns} columns picked; {fix}: pick fewer "
                        "in from_columns", field="from_columns")


_CELL = obj({"row": integer("From 0.", minimum=0, maximum=10000, optional=True),
             "col": integer("From 0.", minimum=0, maximum=10000, optional=True),
             "row_label": string("Or the row's label (first column).", optional=True),
             "col_label": string("Or the column's label (header row).", optional=True),
             "text": string("Keeps the cell's formatting.")})


_SIDES = ["top", "bottom", "left", "right", "all"]


@tool("ppt_edit_table",
      "Change a table: cell text by position or label, insert or delete rows and columns, "
      "merge or split cells, column widths, row heights, cell fill and borders, its style "
      "and Table Design flags. Cell text formatting: ppt_format_text on cells. Returns the "
      "table after.",
      {"doc": string("Document id."),
       "target": string("The table, e.g. 256.7 or $ref."),
       "action": string("What to do.", enum=["set_cells", "insert_row", "delete_row",
                                             "insert_column", "delete_column", "merge",
                                             "split", "set_widths", "set_heights",
                                             "format_cells", "set_style", "set_flags"]),
       "cells": array(_CELL, "set_cells: the cells.", optional=True),
       "from": string("set_cells: rows of d:<id> or t:<table>, from row index.",
                      optional=True),
       "from_columns": _FROM_COLUMNS,
       "header": _HEADER,
       "index": integer("insert_*: new row/column's position, default last; delete_*: "
                        "which; set_cells from: first row.", minimum=0,
                        maximum=10000, optional=True),
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
                        optional=True),
       "style": _STYLE,
       "flags": obj({flag: boolean(optional=True) for flag in FLAGS},
                    "set_flags: on or off; others stay.", optional=True)},
      kind=KIND, group=OBJECTS, mutates=True, refs=("target",), sources=("from",),
      aliases=_COLUMN_NAMES)
@library_errors
def ppt_edit_table(call, doc, target, action, cells=None, index=None, like=None, span=None,
                   sizes=None, rows=None, columns=None, fill=None, borders=None, style=None,
                   flags=None, header=None, from_columns=None, **given):
    frame, table = _table(call.document, target)
    source = given.get("from")
    if source is not None:
        if action != "set_cells":
            raise ToolError("invalid_arguments", f"{action} takes no from; from is for "
                            "set_cells", field="from", valid_options=["set_cells"])
        if cells is not None:
            raise ToolError("invalid_arguments", "from replaces cells: give one",
                            field="cells")
        return _set_from(call, frame, table, source, from_columns, header, index or 0)
    _no_source(from_columns, header)

    def need(name, value):
        if value is None:
            raise ToolError("invalid_arguments", f"{action} needs {name}", field=name)
        return value

    changed = [frame.id]
    record = _provenance.read(frame)
    after = record
    if action == "set_cells":
        written = []
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
            written.append(_provenance.cell_of(found.address)[1:])
        after = _provenance.after_cells(record, written)
    elif action == "insert_row":
        at = table.rows if index is None else index
        table.insert_row(index, like=like)
        after = _provenance.after_row_insert(record, at)
    elif action == "delete_row":
        table.delete_row(need("index", index))
        after = _provenance.after_row_delete(record, index)
    elif action == "insert_column":
        at = table.columns if index is None else index
        table.insert_column(index, like=like)
        after = _provenance.after_column_insert(record, at)
    elif action == "delete_column":
        table.delete_column(need("index", index))
        after = _provenance.after_column_delete(record, index)
    elif action == "merge":
        area = need("span", span)
        to_row, to_col = area.get("to_row", area["row"]), area.get("to_col", area["col"])
        table.merge(area["row"], area["col"], to_row, to_col)
        after = _provenance.after_cells(record, [
            (row, col) for row in range(area["row"], to_row + 1)
            for col in range(area["col"], to_col + 1)])
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
    elif action == "set_style":
        table.style = _style_id(call.document, need("style", style))
    elif action == "set_flags":
        if not need("flags", flags):
            raise ToolError("invalid_arguments", "set_flags needs a flag", field="flags",
                            valid_options=list(FLAGS))
        table.set_flags(**flags)
    if after != record:
        call.store_provenance(frame.id, after)
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id)
    return Result(summary=f"{action} on table {frame.id}", changed=changed,
                  data=table_json(fresh, fresh.table))


def _set_from(call, frame, table, source, from_columns, header, top: int) -> Result:
    """``set_cells`` from a source: the block written from row ``top``, column 0."""
    placed = _placed(call, source, from_columns, header)
    if top >= table.rows:
        raise ToolError("not_found", f"no row {top}; the table has {table.rows}",
                        field="index")
    _fits(placed, table.rows - top, table.columns,
          room=f"the table has {table.rows - top} from row {top}")
    changed = [frame.id]
    for row, values in enumerate(placed.texts(), start=top):
        for column, text in enumerate(values):
            cell = table.cell(row, column)
            cell.text = text
            changed.append(cell.address)
    need_rows, need_columns = placed.size
    warnings = []
    if top + need_rows < table.rows or need_columns < table.columns:
        warnings.append(f"from filled rows {top}-{top + need_rows - 1}, columns "
                        f"0-{need_columns - 1} of a {table.rows}x{table.columns} table; the "
                        "other cells keep their text")
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id)
    data = table_json(fresh, fresh.table)
    data["from"] = placed.provenance(top)
    call.keep_provenance(frame.id, placed.rows, data["from"])
    return Result(summary=f"set_cells on table {frame.id} from {placed.rows.address}",
                  changed=changed, warnings=warnings, data=data)


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


def forget_cells(call, addresses) -> None:
    """Cells written without a source (``ppt_set_text``, ``replace_text``): the records of
    their tables drop the columns written in the block (:func:`~pptx_agent.edit.provenance.after_cells`)."""
    by_frame: dict[str, list[tuple[int, int]]] = {}
    for address in addresses:
        cell = _provenance.cell_of(address)
        if cell is not None:
            by_frame.setdefault(cell[0], []).append(cell[1:])
    for frame_id, cells in by_frame.items():
        try:
            frame = call.document.shape(frame_id)
        except (KeyError, IndexError, ValueError):
            continue
        record = _provenance.read(frame)
        after = _provenance.after_cells(record, cells)
        if after != record:
            call.store_provenance(frame.id, after)


def store_provenance(deck, address: str, record) -> None:
    """``DocumentFormat.store_provenance``: ``record`` into the frame at ``address`` (a
    table's or a chart's), ``None`` dropping it (:func:`pptx_agent.edit.provenance.store`)."""
    _provenance.store(deck.shape(str(address)), record)


def read_provenance(deck) -> dict[str, Any]:
    """``DocumentFormat.read_provenance``: ``{frame address: record}`` (:meth:`Document.provenance
    <pptx_agent.Document.provenance>`)."""
    return deck.provenance()


TOOLS = [ppt_add_table, ppt_edit_table]
